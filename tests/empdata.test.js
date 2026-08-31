'use strict';
// empData_ must read the EMPLOYEES sheet at most once per execution (a memo),
// and re-read after a write clears it. The previous CacheService approach
// never stored anything (the blob exceeded 100 KB), so every call re-read.
const { load } = require('./harness');

function makeFake(rowCount) {
  const counts = { getDataRange: 0, cellsRead: 0 };
  const header = ['employee_id', 'full_name_en', 'record_status', 'exit_date'];
  const values = [header];
  for (let r = 1; r <= rowCount; r++) {
    values.push(['EG' + String(r).padStart(4, '0'), 'Emp ' + r, 'Active', '']);
  }
  const sheet = {
    getName: () => 'EMPLOYEES',
    getLastRow: () => values.length,
    getLastColumn: () => header.length,
    getDataRange: () => { counts.getDataRange++;
      return { getValues: () => { counts.cellsRead += values.length * header.length;
                                  return values.map(r => r.slice()); } }; },
    getRange: (r, c, nr, nc) => ({
      getValues: () => [header.slice(c - 1, c - 1 + (nc || 1))],
      getValue: () => header[c - 1]
    })
  };
  const SpreadsheetApp = { openById: () => ({ getSheetByName: n => n === 'EMPLOYEES' ? sheet : null }) };
  const g = load(['Code.gs'], { SpreadsheetApp });
  return { g, counts };
}

describe('empData_ per-execution memo', () => {
  test('ten calls in one execution read the sheet once', () => {
    const { g, counts } = makeFake(500);
    for (let i = 0; i < 10; i++) g.empData_(false);
    eq(counts.getDataRange, 1, 'getDataRange calls');
  });

  test('active and full-history sets are memoised separately', () => {
    const { g, counts } = makeFake(300);
    g.empData_(false); g.empData_(false);
    g.empData_(true);  g.empData_(true);
    eq(counts.getDataRange, 2, 'one read for active, one for all');
  });

  test('clearEmpCache_ forces the next read to be fresh', () => {
    const { g, counts } = makeFake(300);
    g.empData_(false);
    g.clearEmpCache_();          // simulates a write
    g.empData_(false);
    eq(counts.getDataRange, 2);
  });

  test('the result carries the header and the visible rows', () => {
    const { g } = makeFake(3);
    const E = g.empData_(false);
    eq(E.hdr[0], 'employee_id');
    eq(E.rows.length, 3);
    eq(E.rows[0].values[E.hdr.indexOf('full_name_en')], 'Emp 1');
  });

  test('values keep their raw type — no JSON round-trip changes them', () => {
    const { g } = makeFake(1);
    const E = g.empData_(false);
    // strings stay strings (no ISO-date coercion of the kind the old cache would cause)
    eq(typeof E.rows[0].values[0], 'string');
  });
});

describe('getManagerIdentity_ per-execution memo', () => {
  function fakeWithEmail(matchEmail) {
    const counts = { getDataRange: 0 };
    const header = ['employee_id', 'full_name_en', 'record_status', 'exit_date', 'konecta_email', 'direct_manager'];
    const values = [header,
      ['EG0001', 'Boss',  'Active', '', matchEmail || 'boss@konecta.com', ''],
      ['EG0002', 'Report','Active', '', 'r@konecta.com', 'EG0001']];
    const sheet = {
      getName: () => 'EMPLOYEES', getLastRow: () => values.length, getLastColumn: () => header.length,
      getDataRange: () => { counts.getDataRange++; return { getValues: () => values.map(r => r.slice()) }; },
      getRange: (r,c,nr,nc) => ({ getValues: () => [header.slice(c-1, c-1+(nc||1))], getValue: () => header[c-1] })
    };
    const SpreadsheetApp = { openById: () => ({ getSheetByName: n => n === 'EMPLOYEES' ? sheet : null }) };
    const Session = { getScriptTimeZone: () => 'Africa/Cairo',
                      getActiveUser: () => ({ getEmail: () => 'boss@konecta.com' }) };
    const { load } = require('./harness');
    return { g: load(['Code.gs'], { SpreadsheetApp, Session }), counts };
  }

  test('repeated calls resolve identity but read the sheet once', () => {
    const { g, counts } = fakeWithEmail('boss@konecta.com');
    const a = g.getManagerIdentity_();
    const b = g.getManagerIdentity_();
    const c = g.getManagerIdentity_();
    eq(a && a.id, 'EG0001');
    eq(b && b.id, 'EG0001');
    eq(counts.getDataRange, 1, 'identity resolved once, then memoised');
  });

  test('a null result (not a manager) is memoised too, not recomputed', () => {
    const { g, counts } = fakeWithEmail('nobody@konecta.com');
    eq(g.getManagerIdentity_(), null);
    eq(g.getManagerIdentity_(), null);
    eq(counts.getDataRange, 1);
  });
});
