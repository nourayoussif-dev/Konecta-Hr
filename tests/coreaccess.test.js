'use strict';
/**
 * Tests for the spreadsheet access layer in Code.gs.
 *
 * These do not check business logic — they check COST. Apps Script charges a
 * round trip for every call into the Spreadsheet service, and the 6-minute
 * execution ceiling is spent in round trips, not in arithmetic. So the fake
 * spreadsheet below counts calls and cells, and the tests assert on those
 * numbers directly.
 */
const { load } = require('./harness');

const EMPLOYEE_ROWS = 2000;   // realistic for Konecta Egypt today
const EMPLOYEE_COLS = 80;

function makeFake() {
  const counts = {
    openById: 0, getSheetByName: 0, getLastRow: 0, getLastColumn: 0,
    getRange: 0, getDataRange: 0, cellsRead: 0, cellsWritten: 0
  };

  // header row + EMPLOYEE_ROWS employees
  const header = ['employee_id', 'full_name_en', 'job_title', 'konecta_email'];
  while (header.length < EMPLOYEE_COLS) header.push('col_' + header.length);
  const values = [header];
  for (let r = 1; r <= EMPLOYEE_ROWS; r++) {
    const row = new Array(EMPLOYEE_COLS).fill('x');
    row[0] = 'EG' + String(r).padStart(4, '0');
    row[1] = 'Employee ' + r;
    row[2] = 'Agent';
    row[3] = `e${r}@konecta.com`;
    values.push(row);
  }

  const sheet = {
    getName: () => 'EMPLOYEES',
    getLastRow: () => { counts.getLastRow++; return values.length; },
    getLastColumn: () => { counts.getLastColumn++; return values[0].length; },
    getDataRange: () => {
      counts.getDataRange++;
      return { getValues: () => { counts.cellsRead += values.length * values[0].length;
                                  return values.map(r => r.slice()); } };
    },
    getRange: (r, c, nr, nc) => {
      counts.getRange++;
      nr = nr || 1; nc = nc || 1;
      return {
        getValues: () => {
          counts.cellsRead += nr * nc;
          const out = [];
          for (let i = 0; i < nr; i++) out.push((values[r - 1 + i] || []).slice(c - 1, c - 1 + nc));
          return out;
        },
        getValue: () => { counts.cellsRead += 1; return (values[r - 1] || [])[c - 1]; },
        setValue: v => { counts.cellsWritten += 1; if (values[r - 1]) values[r - 1][c - 1] = v; },
        setValues: () => { counts.cellsWritten += nr * nc; }
      };
    }
  };

  const SpreadsheetApp = {
    openById: () => {
      counts.openById++;
      return { getSheetByName: n => { counts.getSheetByName++;
                                      return n === 'EMPLOYEES' ? sheet : null; } };
    }
  };

  const g = load(['Code.gs'], { SpreadsheetApp });
  const reset = () => {
    g._MEMO.ss = null; g._MEMO.sheets = {}; g._MEMO.headers = {};
    Object.keys(counts).forEach(k => counts[k] = 0);
  };
  return { g, counts, reset };
}

describe('Spreadsheet access — the file is opened once per execution', () => {
  const { g, counts, reset } = makeFake();

  test('ss_() opens the spreadsheet once however often it is called', () => {
    reset();
    for (let i = 0; i < 50; i++) g.ss_();
    eq(counts.openById, 1, 'openById calls');
  });

  test('sheet_() resolves a tab once and reuses it', () => {
    reset();
    for (let i = 0; i < 50; i++) g.sheet_('EMPLOYEES');
    eq(counts.openById, 1);
    eq(counts.getSheetByName, 1, 'getSheetByName calls');
  });

  test('a missing tab is NOT memoised, so a tab created later is still found', () => {
    // TAB_CONTRACTS and TAB_APPOINTMENTS are created on first use; caching the
    // miss would leave them permanently invisible.
    reset();
    eq(g.sheet_('DOES_NOT_EXIST'), null);
    eq(g.sheet_('DOES_NOT_EXIST'), null);
    eq(counts.getSheetByName, 2, 'a miss must be re-queried each time');
  });

  test('headers_() reads the header row once per tab', () => {
    reset();
    for (let i = 0; i < 50; i++) g.headers_('EMPLOYEES');
    eq(counts.getLastColumn, 1, 'getLastColumn calls');
    eq(counts.getRange, 1, 'getRange calls');
    eq(counts.cellsRead, EMPLOYEE_COLS, 'cells read');
  });

  test('headers_() no longer resolves the sheet twice in one call', () => {
    // It used to read: sheet_(n).getRange(1,1,1,sheet_(n).getLastColumn())
    reset();
    g.headers_('EMPLOYEES');
    eq(counts.getSheetByName, 1);
  });

  test('headers_() returns [] for a missing tab instead of throwing', () => {
    reset();
    eq(JSON.stringify(g.headers_('NOPE')), '[]');
  });

  test('the header memo can be cleared when a column is added', () => {
    reset();
    g.headers_('EMPLOYEES');
    g.clearHeaderMemo_('EMPLOYEES');
    g.headers_('EMPLOYEES');
    eq(counts.getRange, 2, 'header re-read after an explicit clear');
  });
});

describe('employeeFieldsOf_ — cost per lookup', () => {
  const { g, counts, reset } = makeFake();
  const FULL_TABLE = (EMPLOYEE_ROWS + 1) * EMPLOYEE_COLS;   // 160,800

  test('it returns the right values', () => {
    reset();
    const f = g.employeeFieldsOf_('EG0500', ['full_name_en', 'job_title']);
    eq(f.full_name_en, 'Employee 500');
    eq(f.job_title, 'Agent');
  });

  test('it never reads the whole table', () => {
    reset();
    g.employeeFieldsOf_('EG0500', ['full_name_en']);
    eq(counts.getDataRange, 0, 'getDataRange must not be used');
    if (counts.cellsRead >= FULL_TABLE) {
      throw new Error(`read ${counts.cellsRead} cells, the whole table is ${FULL_TABLE}`);
    }
  });

  test('one lookup costs roughly rows + columns, not rows x columns', () => {
    reset();
    g.employeeFieldsOf_('EG1234', ['full_name_en', 'job_title']);
    const budget = EMPLOYEE_ROWS + EMPLOYEE_COLS * 2;   // id column + header + one row
    if (counts.cellsRead > budget) {
      throw new Error(`read ${counts.cellsRead} cells, budget is ${budget}`);
    }
  });

  test('the loop call sites are now affordable — 100 lookups stay under one full table', () => {
    // hrInviteToSign, hrPendingDependants, hrMovementReport and hrMarkSigning
    // all call this inside a loop. 100 iterations used to cost 100 full-table
    // reads: over 16 million cells.
    reset();
    for (let i = 1; i <= 100; i++) {
      g.employeeFieldsOf_('EG' + String(i).padStart(4, '0'), ['full_name_en', 'konecta_email']);
    }
    const oldCost = 100 * FULL_TABLE;                    // 16,080,000
    if (counts.cellsRead >= oldCost / 50) {
      throw new Error(`100 lookups read ${counts.cellsRead} cells; want at least a 50x cut from ${oldCost}`);
    }
    console.log(`    \x1b[2m100 lookups: ${counts.cellsRead.toLocaleString()} cells ` +
                `vs ${oldCost.toLocaleString()} before — ` +
                `${Math.round(oldCost / counts.cellsRead)}x less\x1b[0m`);
  });

  test('an unknown employee id comes back empty rather than throwing', () => {
    reset();
    eq(JSON.stringify(g.employeeFieldsOf_('EG9999999', ['full_name_en'])), '{}');
  });

  test('a blank id short-circuits without touching the sheet', () => {
    reset();
    eq(JSON.stringify(g.employeeFieldsOf_('', ['full_name_en'])), '{}');
    eq(JSON.stringify(g.employeeFieldsOf_(null, ['full_name_en'])), '{}');
    eq(counts.openById, 0, 'must not open the spreadsheet at all');
  });

  test('a field that is not a real column comes back blank, not undefined', () => {
    reset();
    const f = g.employeeFieldsOf_('EG0500', ['full_name_en', 'no_such_column']);
    eq(f.no_such_column, '');
  });
});
