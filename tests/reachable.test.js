'use strict';
/**
 * The reachable-set test.
 *
 * Apps Script exposes EVERY top-level function whose name does not end in `_`
 * to google.script.run — not just the ones the client happens to call. The
 * first security sweep only inspected client-referenced functions, so
 * buildPayslip and publishPayrollMonth (ungated globals sitting behind gated
 * wrappers) were never reviewed. This test enumerates the real reachable set
 * and fails on any newcomer that is neither gated nor explicitly accepted.
 *
 * When you add a top-level function:
 *   - internal helper  -> end its name with `_` (unreachable), or
 *   - real endpoint     -> gate it (isHR_/isIT_/actingFor_/self-scope), or
 *   - deliberately open  -> add it to ACCEPTED_OPEN below WITH a reason.
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const GS = ['Code.gs', 'payslip.gs', 'payslippdf.gs', 'payrollchecks.gs',
            'payrollarchive.gs', 'payrollengine.gs'];

// Gate tokens that, appearing near the top of a function, count as authorized.
const GATE = /\b(isHR_|isIT_|isFacilities_|getManagerIdentity_|actingFor_|getMyRecord|getMyLeaveInfo|getMyResignation|getMyContract|getMyDependants|currentUser_|guardRow_|guardEmpRow_|assertNotDirectCall_)\s*\(/;

// Deliberately open to any signed-in Konecta employee. Each MUST have a reason.
const ACCEPTED_OPEN = {
  doGet: 'the web-app entry point itself',
  onFormSubmit: 'installable Forms trigger; refuses browser calls via the e.namedValues guard',
  getBootstrap: 'self-scopes internally; role decided server-side',
  getRole: 'returns only the caller\'s own role',
  getLists: 'dropdown option lists, no personal data',
  getProjectMap: 'project -> cost-centre map, reference data',
  getManagerOptions: 'manager dropdown; names + ids only, reviewed',
  getHolidayList: 'the public holiday calendar',
  getLeaveTypes: 'leave-type reference list',
  countLeaveDays: 'pure calendar arithmetic on caller inputs',
  getDayPicker: 'pure calendar arithmetic on caller inputs',
  validateNationalId: 'validates a number the caller already holds',
  submitNewEmployee: 'joiner intake — the caller has no record yet',
  lookupEmployeeForNoShow: 'no-show reporting is open by design; name/title/status only',
  reportNoShow: 'no-show reporting is open by design; holds are HR-reviewed',
  listPayslipPeriods: 'period names only, no amounts',
  departmentList: 'department reference list',
  departmentNames_: 'helper',
  // payroll engine: pure computation on caller-supplied inputs, no stored PII
  firstBracketApplied: 'pure tax math', incomeTaxMonthly: 'pure tax math',
  insuranceWage: 'pure insurance math', overtimeAmount: 'pure overtime math',
  grossUp: 'pure gross-up math', daysThisMonth: 'pure day-count math',
  arrearsDays: 'pure day-count math', calculateEmployee: 'pure payroll math on caller input',
  priceTransportation: 'pure pricing', priceShiftAllowance: 'pure pricing',
  priceOnCall: 'pure pricing', priceIftar: 'pure pricing', priceKPI: 'pure pricing',
  priceLeaveEncashment: 'pure pricing', rulesFor: 'reads versioned constants only',
  useRulesFor: 'binds versioned constants only',
  wasPaidInPeriod: 'returns a boolean for one id+period; no PII',
};

function reachable() {
  const out = [];
  for (const f of GS) {
    const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
    const re = /^function ([a-zA-Z][A-Za-z0-9]*)\s*\(/gm;
    let m;
    while ((m = re.exec(src))) {
      const name = m[1];
      if (name.endsWith('_')) continue;             // private, unreachable
      const head = src.slice(m.index, m.index + 800);
      out.push({ name, file: f, gated: GATE.test(head) });
    }
  }
  return out;
}

describe('Reachable-set authorization', () => {
  const fns = reachable();

  test('every reachable global is gated or explicitly accepted-open', () => {
    const bad = fns.filter(x => !x.gated && !(x.name in ACCEPTED_OPEN));
    if (bad.length) {
      throw new Error('Ungated reachable globals (gate them, add a trailing _, ' +
        'or list in ACCEPTED_OPEN with a reason):\n  ' +
        bad.map(x => `${x.name}  (${x.file})`).join('\n  '));
    }
  });

  test('buildPayslip is no longer reachable — it is buildPayslip_', () => {
    eq(fns.some(x => x.name === 'buildPayslip'), false);
  });

  test('publishPayrollMonth is guarded, not open', () => {
    const p = fns.find(x => x.name === 'publishPayrollMonth');
    if (p) eq(p.gated, true, 'publishPayrollMonth must carry assertNotDirectCall_');
  });

  test('the accepted-open list has not rotted — every entry still exists', () => {
    const names = new Set(fns.map(x => x.name).concat(
      // helpers that may legitimately be renamed with _ later
      Object.keys(ACCEPTED_OPEN).filter(n => n.endsWith('_'))));
    Object.keys(ACCEPTED_OPEN).forEach(n => {
      if (!names.has(n)) throw new Error(`ACCEPTED_OPEN lists ${n}, which no longer exists — remove it`);
    });
  });
});
