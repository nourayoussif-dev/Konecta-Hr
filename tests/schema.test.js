'use strict';
/**
 * Tests for the schema guard. A renamed header does not throw anywhere in
 * the app — indexOf returns -1 and reads come back blank — so the guard is
 * the only thing standing between a renamed column and silently wrong data.
 */
const { load } = require('./harness');

function sheetOf(name, hdr) {
  return {
    getName: () => name,
    getLastRow: () => 2,
    getLastColumn: () => hdr.length,
    getRange: (r, c, nr, nc) => ({
      getValues: () => [hdr.slice(c - 1, c - 1 + (nc || 1))],
      getValue: () => hdr[c - 1]
    })
  };
}

function envWith(tabs, extra) {
  const SpreadsheetApp = { openById: () => ({ getSheetByName: n => tabs[n] || null }) };
  const sent = [];
  const MailApp = { sendEmail: (to, subj, body) => sent.push({ to, subj, body }) };
  const g = load(['Code.gs'], Object.assign({ SpreadsheetApp, MailApp }, extra));
  return { g, sent };
}

/** Build a full, healthy tab set from the code's own requirements. */
function healthyTabs(g) {
  const want = g.requiredSchema_();
  const tabs = {};
  Object.keys(want).forEach(t => {
    // give layout-only tabs a dummy column so headers_ has something to read
    tabs[t] = sheetOf(t, want[t].length ? want[t].slice() : ['a']);
  });
  return tabs;
}

describe('Schema guard', () => {
  test('a healthy sheet reports no problems', () => {
    const probe = envWith({}).g;               // just to read requiredSchema_
    const { g } = envWith(healthyTabs(probe));
    eq(JSON.stringify(g.schemaProblems_()), '[]');
  });

  test('a renamed EMPLOYEES column is reported by name', () => {
    const probe = envWith({}).g;
    const tabs = healthyTabs(probe);
    const emp = probe.requiredSchema_()['EMPLOYEES'].slice();
    emp[emp.indexOf('basic_salary')] = 'salary';        // someone "tidied" the header
    tabs['EMPLOYEES'] = sheetOf('EMPLOYEES', emp);
    const { g } = envWith(tabs);
    const p = g.schemaProblems_();
    eq(p.length, 1);
    eq(p[0].tab, 'EMPLOYEES');
    eq(JSON.stringify(p[0].missing_columns), '["basic_salary"]');
  });

  test('a deleted tab is reported as missing', () => {
    const probe = envWith({}).g;
    const tabs = healthyTabs(probe);
    delete tabs['RESIGNATIONS'];
    const { g } = envWith(tabs);
    const p = g.schemaProblems_();
    eq(p.length, 1);
    eq(p[0].tab, 'RESIGNATIONS');
    eq(p[0].missing_tab, true);
  });

  test('the EMPLOYEES requirements track the gate constants automatically', () => {
    const { g } = envWith({});
    const req = g.requiredSchema_()['EMPLOYEES'];
    g.__eval('GATE1.concat(GATE2, GATE3)').forEach(f => {
      if (req.indexOf(f) === -1) throw new Error(`gate field ${f} not in schema guard`);
    });
  });

  test('the daily check emails HR while broken, and stays quiet when healthy', () => {
    const probe = envWith({}).g;
    const broken = healthyTabs(probe);
    delete broken['LEAVE'];
    const b = envWith(broken);
    eq(b.g.schemaDailyCheck_(), false);
    eq(b.sent.length, 1, 'one email to HR');
    if (!/LEAVE/.test(b.sent[0].body)) throw new Error('email must name the broken tab');

    const h = envWith(healthyTabs(probe));
    eq(h.g.schemaDailyCheck_(), true);
    eq(h.sent.length, 0, 'no email when healthy');
  });

  test('hrSchemaCheck refuses non-HR callers', () => {
    const probe = envWith({}).g;
    const { g } = envWith(healthyTabs(probe));   // Session stub = test@konecta.com
    throws(() => g.hrSchemaCheck());
  });
});
