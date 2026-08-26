'use strict';
/**
 * Tests for the row identity guard.
 *
 * The scenario that motivated it: HR opens an editor on row N, someone sorts
 * the EMPLOYEES sheet, HR clicks Save, and the write lands on whoever now
 * occupies row N. 25 client-facing functions were exposed. The guard takes
 * the identity the client fetched alongside the row and either confirms the
 * row, re-finds the record, or refuses.
 */
const { load } = require('./harness');

/** Minimal in-memory Sheet — enough for guardRow_, hrVerifyBank and logChange_. */
function makeSheet(name, values) {
  return {
    _v: values,
    getName: () => name,
    getLastRow() { return this._v.length; },
    getLastColumn() { return this._v[0] ? this._v[0].length : 0; },
    getMaxRows() { return this._v.length; },
    getDataRange() {
      const self = this;
      return { getValues: () => self._v.map(r => r.slice()) };
    },
    getRange(r, c, nr, nc) {
      const self = this; nr = nr || 1; nc = nc || 1;
      return {
        getValue: () => (self._v[r - 1] || [])[c - 1],
        getValues: () => {
          const out = [];
          for (let i = 0; i < nr; i++) {
            const row = self._v[r - 1 + i] || [];
            out.push(row.slice(c - 1, c - 1 + nc));
          }
          return out;
        },
        setValue: v => {
          while (self._v.length < r) self._v.push([]);
          self._v[r - 1][c - 1] = v;
        },
        setValues: vals => {
          for (let i = 0; i < vals.length; i++) {
            while (self._v.length < r + i) self._v.push([]);
            for (let j = 0; j < vals[i].length; j++) self._v[r - 1 + i][c - 1 + j] = vals[i][j];
          }
        }
      };
    }
  };
}

function makeEnv(sheets, userEmail) {
  const SpreadsheetApp = {
    openById: () => ({ getSheetByName: n => sheets[n] || null })
  };
  const Session = {
    getScriptTimeZone: () => 'Africa/Cairo',
    getActiveUser: () => ({ getEmail: () => userEmail || 'test@konecta.com' })
  };
  return load(['Code.gs'], { SpreadsheetApp, Session });
}

const HDR = ['employee_id', 'national_id', 'full_name_en', 'bank_verified', 'updated_at', 'updated_by'];
function empRows() {
  return [
    HDR.slice(),
    ['EG0001', '29001010100001', 'Ahmed',  'Pending verification', '', ''],
    ['EG0002', '29001010100002', 'Basma',  'Pending verification', '', ''],
    ['EG0003', '29001010100003', 'Karim',  'Pending verification', '', ''],
  ];
}

describe('guardRow_ — the row identity guard', () => {
  test('a row still in place is confirmed unchanged', () => {
    const g = makeEnv({ EMPLOYEES: makeSheet('EMPLOYEES', empRows()) });
    const sh = g.sheet_('EMPLOYEES');
    eq(g.guardRow_(sh, HDR, 3, { employee_id: 'EG0002' }), 3);
  });

  test('THE scenario: the sheet is sorted under an open editor — the write follows the employee', () => {
    const rows = empRows();
    // sort descending by name: Karim, Basma, Ahmed
    const body = rows.slice(1).sort((a, b) => b[2].localeCompare(a[2]));
    const g = makeEnv({ EMPLOYEES: makeSheet('EMPLOYEES', [rows[0]].concat(body)) });
    const sh = g.sheet_('EMPLOYEES');
    // HR fetched Basma at row 3 before the sort; she now lives at row 3? No:
    // order is Karim(2), Basma(3), Ahmed(4) — Basma happens to stay, use Ahmed:
    // Ahmed was row 2, is now row 4.
    eq(g.guardRow_(sh, HDR, 2, { employee_id: 'EG0001' }), 4, 'Ahmed re-found at his new row');
  });

  test('a deleted record refuses instead of writing to whoever moved up', () => {
    const rows = empRows();
    rows.splice(2, 1);                    // Basma removed
    const g = makeEnv({ EMPLOYEES: makeSheet('EMPLOYEES', rows) });
    throws(() => g.guardRow_(g.sheet_('EMPLOYEES'), HDR, 3, { employee_id: 'EG0002' }),
           'must refuse when the record is gone');
  });

  test('a duplicate identity refuses — it will not guess between two rows', () => {
    const rows = empRows();
    rows.push(['EG0002', '29001010100009', 'Basma clone', 'Pending verification', '', '']);
    const g = makeEnv({ EMPLOYEES: makeSheet('EMPLOYEES', rows) });
    // stale row: the guard must go to the scan, then see two EG0002 rows
    throws(() => g.guardRow_(g.sheet_('EMPLOYEES'), HDR, 99, { employee_id: 'EG0002' }));
  });

  test('a blank key refuses — a stale pre-guard page must not fall through', () => {
    const g = makeEnv({ EMPLOYEES: makeSheet('EMPLOYEES', empRows()) });
    const sh = g.sheet_('EMPLOYEES');
    throws(() => g.guardRow_(sh, HDR, 2, { employee_id: '' }));
    throws(() => g.guardRow_(sh, HDR, 2, { employee_id: null }));
    throws(() => g.guardRow_(sh, HDR, 2, { employee_id: undefined }));
  });

  test('a renamed column is named in the error, not read as blank', () => {
    const g = makeEnv({ EMPLOYEES: makeSheet('EMPLOYEES', empRows()) });
    try {
      g.guardRow_(g.sheet_('EMPLOYEES'), HDR, 2, { no_such_column: 'x' });
      throw new Error('expected a throw');
    } catch (e) {
      if (!/no_such_column/.test(e.message)) throw new Error('error must name the column: ' + e.message);
    }
  });

  test('composite keys intersect — dependants are found by employee AND name', () => {
    const dhdr = ['employee_id', 'name', 'relation'];
    const dep = makeSheet('DEPENDANTS', [
      dhdr,
      ['EG0001', 'Salma', 'Daughter'],
      ['EG0001', 'Omar',  'Son'],
      ['EG0002', 'Omar',  'Son'],       // same name, different employee
    ]);
    const g = makeEnv({ DEPENDANTS: dep });
    eq(g.guardRow_(dep, dhdr, 99, { employee_id: 'EG0001', name: 'Omar' }), 3);
    eq(g.guardRow_(dep, dhdr, 99, { employee_id: 'EG0002', name: 'Omar' }), 4);
  });

  test('a Date cell matches the yyyy-MM-dd string the client was given', () => {
    const hdr2 = ['request_id', 'start_date'];
    const lv = makeSheet('LEAVE', [
      hdr2,
      ['LV-1', new Date(2026, 7, 3)],
      ['LV-2', new Date(2026, 7, 10)],
    ]);
    const g = makeEnv({ LEAVE: lv });
    // fmt_ made the client's copy '2026-08-10'; the cell holds a Date object
    eq(g.guardRow_(lv, hdr2, 99, { start_date: '2026-08-10' }), 3);
  });

  test('firstDataRow skips header junk — the INTAKE layout starts at sheet row 5', () => {
    const ihdr = ['national_id', 'full_name_en'];
    const intake = makeSheet('INTAKE', [
      ihdr,
      ['(instructions row)', ''],
      ['(instructions row)', ''],
      ['(column notes)', ''],
      ['29001010100005', 'Nour'],
      ['29001010100006', 'Hana'],
    ]);
    const g = makeEnv({ INTAKE: intake });
    eq(g.guardRow_(intake, ihdr, 99, { national_id: '29001010100006' }, 5), 6);
  });
});

describe('guardEmpRow_ — employee identity by whichever ID the record has', () => {
  test('an EG id is matched on employee_id', () => {
    const g = makeEnv({ EMPLOYEES: makeSheet('EMPLOYEES', empRows()) });
    eq(g.guardEmpRow_(g.sheet_('EMPLOYEES'), HDR, 99, 'EG0003'), 4);
  });
  test('a national id is matched on national_id — the pre-issue record', () => {
    const g = makeEnv({ EMPLOYEES: makeSheet('EMPLOYEES', empRows()) });
    eq(g.guardEmpRow_(g.sheet_('EMPLOYEES'), HDR, 99, '29001010100001'), 2);
  });
});

describe('end to end: hrVerifyBank on a sorted sheet', () => {
  test('the verification lands on the right employee, wherever they moved to', () => {
    // Ahmed(2) Basma(3) Karim(4). HR opens the bank queue, then the sheet is
    // sorted to Karim(2) Basma(3) Ahmed(4). HR clicks Verify on Ahmed's card,
    // which still says row 2 — where Karim now sits.
    const rows = empRows();
    const body = rows.slice(1).sort((a, b) => b[2].localeCompare(a[2]));
    const emp = makeSheet('EMPLOYEES', [rows[0]].concat(body));
    const log = makeSheet('CHANGE LOG', [['log_id']]);
    const g = makeEnv({ EMPLOYEES: emp, 'CHANGE LOG': log }, 'eghr@konecta.com');

    const res = g.hrVerifyBank(2, 'verify', 'EG0001');
    eq(res.ok, true);

    const bank = HDR.indexOf('bank_verified');
    const byId = {};
    emp._v.slice(1).forEach(r => { byId[r[0]] = r[bank]; });
    eq(byId['EG0001'], 'Verified', "Ahmed — the employee HR meant — is verified");
    eq(byId['EG0002'], 'Pending verification', 'Basma untouched');
    eq(byId['EG0003'], 'Pending verification', "Karim — who sits at Ahmed's old row — untouched");
  });

  test('before the guard, this exact call verified the wrong employee', () => {
    // Document the failure the guard closes: resolve row 2 blindly.
    const rows = empRows();
    const body = rows.slice(1).sort((a, b) => b[2].localeCompare(a[2]));
    const stale = [rows[0]].concat(body);
    eq(stale[1][0], 'EG0003', "row 2 now holds Karim — the old code would have verified HIM");
  });
});
