'use strict';
/**
 * Regression tests for the Egyptian payroll engine.
 *
 * The engine was validated once, by hand, on 20/08/2026 against the July 2026
 * payroll from the outsourced provider (basic/overtime/martyrs 100%, income
 * tax 99.33%, social insurance 99.0%). Nothing has protected it since. These
 * tests lock in that behaviour so a future edit cannot quietly change what
 * people are paid.
 */
const { load } = require('./harness');
const g = load(['payrollengine.gs']);

const D = (y, m, d) => new Date(y, m - 1, d);
const JAN_START = D(2026, 1, 1);
const JAN_END   = D(2026, 1, 31);
const FEB_START = D(2026, 2, 1);
const FEB_END   = D(2026, 2, 28);

// ───────────────────────────────────────────────────────────── income tax ──
describe('Income tax — progressive brackets', () => {
  const si = s => g.insuranceWage('EG1', s, 30, D(2020, 1, 1), null, JAN_START, JAN_END) * g.CFG.SI_EMPLOYEE;
  const tax = s => g.incomeTaxMonthly(s, si(s), 0, 0);

  test('cumulative bracket rates are 0/10/15/20/22.5/25/27.5%', () => {
    let cum = 0;
    const got = g.RATE_DELTA.map(d => +( (cum += d) * 100 ).toFixed(1));
    eq(JSON.stringify(got), JSON.stringify([0, 10, 15, 20, 22.5, 25, 27.5]));
  });

  test('a salary under the exemption pays no tax', () => {
    eq(tax(5000), 0);
  });

  test('tax is monotonic — earning more never lowers the bill', () => {
    let prev = -1;
    for (let s = 1000; s <= 200000; s += 1000) {
      const t = tax(s);
      if (t < prev - 0.001) throw new Error(`tax fell at ${s}: ${t} < ${prev}`);
      prev = t;
    }
  });

  test('known reference points hold', () => {
    close(tax(10000),    592.50, 0.01, '10,000/month');
    close(tax(20000),   2445.10, 0.01, '20,000/month');
    close(tax(40000),   7019.92, 0.01, '40,000/month');
    close(tax(100000), 23290.75, 0.01, '100,000/month');
  });

  test('the personal exemption is applied monthly, not annually', () => {
    // 20,000/year => 1,666.67/month off the base
    const withExemption = g.incomeTaxMonthly(20000, 0, 0, 0);
    const manualBase = (20000 - g.CFG.PERSONAL_EXEMPTION_ANNUAL / 12) * 12;
    const fb = g.firstBracketApplied(manualBase);
    let expected = 0;
    for (let i = 0; i < 7; i++) {
      const lower = (i + 1 > fb) ? g.BRACKET_LOWER[i] : 0;
      if (manualBase > lower) expected += (manualBase - lower) * g.RATE_DELTA[i];
    }
    close(withExemption, expected / 12, 0.001);
  });

  test('non-taxable amounts inside total salary are excluded from the base', () => {
    const plain    = g.incomeTaxMonthly(30000, 1837, 0, 0);
    const withExempt = g.incomeTaxMonthly(30000, 1837, 0, 10000);
    if (withExempt >= plain) throw new Error('exempting 10,000 did not reduce the tax');
    close(withExempt, g.incomeTaxMonthly(20000, 1837, 0, 0), 0.01,
          'exempting 10,000 of 30,000 should tax like 20,000');
  });

  test('bracket cancellation kicks in at the documented thresholds', () => {
    eq(g.firstBracketApplied(0),         1);
    eq(g.firstBracketApplied(599999),    1);
    eq(g.firstBracketApplied(600000),    2);
    eq(g.firstBracketApplied(700000),    3);
    eq(g.firstBracketApplied(800000),    4);
    eq(g.firstBracketApplied(900000),    5);
    eq(g.firstBracketApplied(1200000),   6);
  });
});

// ─────────────────────────────────────────────────────── social insurance ──
describe('Social insurance — wage floor, ceiling and eligibility', () => {
  const iw = (basic, opts = {}) => g.insuranceWage(
    opts.id || 'EG1', basic, opts.workingDays ?? 30,
    opts.hire || D(2020, 1, 1), opts.exit || null, JAN_START, JAN_END);

  test('a mid-range salary is insured at its own value', () => eq(iw(10000), 10000));
  test('below the floor is lifted to the floor',  () => eq(iw(1000),  g.CFG.INS_WAGE_MIN));
  test('above the ceiling is capped at the ceiling', () => eq(iw(50000), g.CFG.INS_WAGE_MAX));
  test('an exempt employee is never insured', () => eq(iw(10000, { id: 'EG0009' }), 0));

  test('a joiner after day 1 is not insured this month (NOSI will not backdate)', () => {
    eq(iw(10000, { hire: D(2026, 1, 2) }), 0);
    eq(iw(10000, { hire: D(2026, 1, 20) }), 0);
  });

  test('a joiner on day 1 is insured for the full month', () => {
    eq(iw(10000, { hire: D(2026, 1, 1) }), 10000);
  });

  test('a leaver on day 3 or later stays insured for the whole month', () => {
    eq(iw(10000, { exit: D(2026, 1, 3),  workingDays: 0 }), 10000);
    eq(iw(10000, { exit: D(2026, 1, 25), workingDays: 0 }), 10000);
  });

  test('a leaver on day 1 or 2 with no working days is not insured', () => {
    eq(iw(10000, { exit: D(2026, 1, 2), workingDays: 0 }), 0);
  });

  test('employee and employer shares are 11% and 18.75%', () => {
    eq(g.CFG.SI_EMPLOYEE, 0.11);
    eq(g.CFG.SI_EMPLOYER, 0.1875);
  });
});

// ─────────────────────────────────────────────────────────────── gross-up ──
describe('Gross-up — net-to-gross for promised net allowances', () => {
  const taxableBefore = 20000, si = g.CFG.INS_WAGE_MAX * g.CFG.SI_EMPLOYEE;
  const netAt = x => {
    const t = taxableBefore + x;
    return t - si - g.incomeTaxMonthly(t, si, 0, 0) - t * g.CFG.MARTYRS_RATE;
  };

  test('delivers the promised net exactly, at several sizes', () => {
    [500, 1000, 3000, 5000, 12000].forEach(net => {
      const up = g.grossUp(net, taxableBefore, si, 0);
      close(netAt(up) - netAt(0), net, 0.01, `promised net ${net}`);
    });
  });

  test('the gross-up always exceeds the net it delivers', () => {
    [500, 3000, 12000].forEach(net => {
      const up = g.grossUp(net, taxableBefore, si, 0);
      if (up <= net) throw new Error(`gross ${up} should exceed net ${net}`);
    });
  });

  test('zero and falsy promised net cost nothing', () => {
    eq(g.grossUp(0, taxableBefore, si, 0), 0);
    eq(g.grossUp(null, taxableBefore, si, 0), 0);
    eq(g.grossUp(undefined, taxableBefore, si, 0), 0);
  });

  test('terminates rather than looping — bisection is bounded', () => {
    const t0 = Date.now();
    g.grossUp(50000, 0, 0, 0);
    if (Date.now() - t0 > 1000) throw new Error('gross-up took over a second');
  });
});

// ─────────────────────────────────────────────────────────────── overtime ──
describe('Overtime', () => {
  const basic = 9000; // 300/day, 37.50/hour
  test('day hours price at 1.35x the hourly rate', () => {
    close(g.overtimeAmount(basic, 10, 0, 0, 0), 37.5 * 10 * 1.35, 0.001);
  });
  test('night hours price at 1.70x the hourly rate', () => {
    close(g.overtimeAmount(basic, 0, 10, 0, 0), 37.5 * 10 * 1.70, 0.001);
  });
  test('a public holiday worked pays one extra day', () => {
    close(g.overtimeAmount(basic, 0, 0, 2, 0), 300 * 2, 0.001);
  });
  test('a weekend day worked pays one extra day', () => {
    close(g.overtimeAmount(basic, 0, 0, 0, 3), 300 * 3, 0.001);
  });
  test('missing inputs are treated as zero, not NaN', () => {
    eq(g.overtimeAmount(basic), 0);
    eq(g.overtimeAmount(basic, null, undefined, null, undefined), 0);
  });
  test('the hourly rate derives from a 30-day month and an 8-hour day', () => {
    eq(g.CFG.STANDARD_MONTH_DAYS, 30);
    eq(g.CFG.WORK_HOURS_PER_DAY, 8);
  });
});

// ──────────────────────────────────────────────────────────── day counts ──
describe('Day counts — joiners, leavers and arrears', () => {
  test('a full month is 30 days regardless of calendar length', () => {
    eq(g.daysThisMonth(D(2020, 1, 1), null, JAN_START, JAN_END), 30);
    eq(g.daysThisMonth(D(2020, 1, 1), null, FEB_START, FEB_END), 30);
  });

  test('a joiner is paid from their hire date', () => {
    eq(g.daysThisMonth(D(2026, 1, 1),  null, JAN_START, JAN_END), 30);
    eq(g.daysThisMonth(D(2026, 1, 15), null, JAN_START, JAN_END), 16);
  });

  test('a leaver is paid to their exit date, capped at 30', () => {
    eq(g.daysThisMonth(D(2020, 1, 1), D(2026, 1, 20), JAN_START, JAN_END), 20);
    eq(g.daysThisMonth(D(2020, 1, 1), D(2026, 1, 31), JAN_START, JAN_END), 30);
  });

  test('someone who left before this month is paid nothing', () => {
    eq(g.daysThisMonth(D(2020, 1, 1), D(2025, 12, 15), JAN_START, JAN_END), 0);
  });

  test('arrears only trigger for a previous-month hire who was not paid', () => {
    eq(g.arrearsDays(D(2025, 12, 12), D(2025, 12, 1), D(2025, 12, 31), false), 19);
    eq(g.arrearsDays(D(2025, 12, 12), D(2025, 12, 1), D(2025, 12, 31), true),   0,
       'already paid last month');
    eq(g.arrearsDays(D(2025, 11, 12), D(2025, 12, 1), D(2025, 12, 31), false),  0,
       'hired two months ago');
    eq(g.arrearsDays(null, D(2025, 12, 1), D(2025, 12, 31), false), 0);
  });

  test('the 30-day convention: a full month is 30 days and the 31st is nobody\'s payroll day', () => {
    // Not a bug, and deliberately left alone: nobody is paid for a 31st, so a
    // hire on the 31st is 0 days here and a full 30 the following month.
    // This is what produced the 100% Worth Basic match against the provider.
    eq(g.daysThisMonth(D(2020, 1, 1),  null, JAN_START, JAN_END), 30, 'full 31-day month');
    eq(g.daysThisMonth(D(2026, 1, 31), null, JAN_START, JAN_END),  0, 'hired on the 31st');
    eq(g.arrearsDays(D(2025, 12, 31), D(2025, 12, 1), D(2025, 12, 31), false), 0);
  });

  test('someone who joins AND leaves in the same month is paid only the days employed', () => {
    // Regression: the leaver branch used to win and count from the 1st,
    // paying 20 days for 11 worked. This is the no-show / drop-out shape.
    eq(g.daysThisMonth(D(2026, 1, 10), D(2026, 1, 20), JAN_START, JAN_END), 11,
       'hired 10th, exited 20th');
    eq(g.daysThisMonth(D(2026, 1, 1),  D(2026, 1, 30), JAN_START, JAN_END), 30,
       'hired 1st, exited 30th is still a full month');
    eq(g.daysThisMonth(D(2026, 1, 1),  D(2026, 1, 31), JAN_START, JAN_END), 30,
       'exit on the 31st is capped at 30');
    eq(g.daysThisMonth(D(2026, 1, 10), D(2026, 1, 31), JAN_START, JAN_END), 21,
       'hired 10th, exited 31st = 10th..30th');
    eq(g.daysThisMonth(D(2026, 1, 20), D(2026, 1, 20), JAN_START, JAN_END), 1,
       'hired and exited the same day still earns that day');
    eq(g.daysThisMonth(D(2026, 1, 25), D(2026, 1, 10), JAN_START, JAN_END), 0,
       'an exit before the hire date cannot go negative');
  });

  test('a hire dated after this month is paid nothing', () => {
    // Regression: a future hire fell through every branch to `return 30`,
    // paying a full month to someone who had not started.
    eq(g.daysThisMonth(D(2026, 2, 15), null, JAN_START, JAN_END), 0);
    eq(g.daysThisMonth(D(2026, 2, 1),  null, JAN_START, JAN_END), 0);
  });
});

// ─────────────────────────────────────────────────────── add-on pricing ──
describe('Allowance pricing', () => {
  test('transportation prorates on office days against working days', () => {
    close(g.priceTransportation(22, 22), 3000, 0.001);
    close(g.priceTransportation(11, 22), 1500, 0.001);
  });
  test('transportation falls back to a 22-day month when total is absent', () => {
    close(g.priceTransportation(22, 0), 3000, 0.001);
    close(g.priceTransportation(22, null), 3000, 0.001);
  });
  test('shift allowance prorates against a 22-day month', () => {
    close(g.priceShiftAllowance(22), 2000, 0.001);
    close(g.priceShiftAllowance(11), 1000, 0.001);
  });
  test('on-call prorates against a 30-day month', () => {
    close(g.priceOnCall(30), 2000, 0.001);
    close(g.priceOnCall(15), 1000, 0.001);
  });
  test('iftar is a flat rate per day', () => {
    close(g.priceIftar(10), 2500, 0.001);
  });
  test('KPI prorates on days worked when supplied', () => {
    close(g.priceKPI(5000, 1, 30), 5000, 0.001);
    close(g.priceKPI(5000, 1, 15), 2500, 0.001);
    close(g.priceKPI(5000, 0.8, 0), 4000, 0.001, 'no days worked -> full attainment');
  });
  test('leave encashment never goes negative when leave is overtaken', () => {
    eq(g.priceLeaveEncashment(9000, 365, 21, 100), 0);
  });
  test('leave encashment accrues pro rata on service days', () => {
    close(g.priceLeaveEncashment(9000, 365, 21, 0), 21 * 300, 0.001);
    close(g.priceLeaveEncashment(9000, 182.5, 21, 0), 10.5 * 300, 0.001);
  });
});

// ───────────────────────────────────────────────── whole-employee wiring ──
describe('calculateEmployee — the components must reconcile', () => {
  const base = {
    id: 'EG1234', basic: 12000, workingDays: 30, arrearsDays: 0,
    hireDate: D(2020, 1, 1), exitDate: null,
    grossAddOns: 0, netAddOns: 0, nonTaxableAmounts: 0, nonTaxableInTotal: 0,
    dedTaxable: 0, otherDeductions: 0,
    hours135: 0, hours170: 0, publicHolidayDays: 0, weekendDays: 0
  };
  const calc = o => g.calculateEmployee(Object.assign({}, base, o), JAN_START, JAN_END);

  test('net = total salary - deductions + non-taxable amounts', () => {
    const r = calc({ netAddOns: 2000, nonTaxableAmounts: 500, otherDeductions: 300 });
    close(r.netSalary, r.totalSalary - r.totalDeductions + 500, 0.001);
  });

  test('deductions are exactly SI + tax + martyrs + taxable and other deductions', () => {
    const r = calc({ dedTaxable: 400, otherDeductions: 250 });
    close(r.totalDeductions, r.employeeSI + r.incomeTax + r.martyrsFund + 400 + 250, 0.001);
  });

  test('worth basic prorates on working days plus arrears', () => {
    const r = calc({ workingDays: 15, arrearsDays: 5 });
    eq(r.workingDaysTotal, 20);
    close(r.worthBasic, 12000 * 20 / 30, 0.001);
  });

  test('employer cost = total salary + employer SI + emergency fund', () => {
    const r = calc({});
    close(r.totalCost, r.totalSalary + r.employerSI + r.emergencyFund, 0.001);
  });

  test('the emergency fund is charged only when the employee is insured', () => {
    eq(calc({}).emergencyFund, g.CFG.EMERGENCY_FUND_FLAT);
    eq(calc({ id: 'EG0009' }).emergencyFund, 0, 'SI-exempt employee');
  });

  test('the martyrs fund is 0.05% of total salary', () => {
    const r = calc({});
    close(r.martyrsFund, r.totalSalary * 0.0005, 0.001);
  });

  test('a promised net add-on arrives intact in net salary', () => {
    const without = calc({});
    const withNet = calc({ netAddOns: 2000 });
    close(withNet.netSalary - without.netSalary, 2000, 0.02);
  });

  test('no component comes back NaN on a sparse employee record', () => {
    const r = g.calculateEmployee(
      { id: 'EG9', basic: 8000, workingDays: 30, hireDate: D(2020, 1, 1) },
      JAN_START, JAN_END);
    Object.keys(r).forEach(k => {
      if (typeof r[k] === 'number' && isNaN(r[k])) throw new Error(`${k} is NaN`);
    });
  });

  test('a missing or unusable basic salary stops the run instead of paying NaN', () => {
    // Regression: NaN used to propagate through worthBasic, totalSalary,
    // incomeTax and netSalary without throwing, so the employee got a payslip
    // reading NaN and the run reported success.
    const bad = v => () => g.calculateEmployee(
      { id: 'EG777', basic: v, workingDays: 30, hireDate: D(2020, 1, 1) },
      JAN_START, JAN_END);
    throws(bad(undefined), 'undefined basic');
    throws(bad(null),      'null basic');
    throws(bad(''),        'blank cell');
    throws(bad('12000'),   'a number stored as text');
    throws(bad(NaN),       'NaN basic');
    throws(bad(-500),      'negative basic');
  });

  test('the refusal names the employee, so the record can be found and fixed', () => {
    try {
      g.calculateEmployee({ id: 'EG0421', workingDays: 30, hireDate: D(2020, 1, 1) },
                          JAN_START, JAN_END);
      throw new Error('expected a throw');
    } catch (e) {
      if (!/EG0421/.test(e.message)) throw new Error(`message lacks the id: ${e.message}`);
      if (!/basic_salary/.test(e.message)) throw new Error(`message lacks the field: ${e.message}`);
    }
  });

  test('a zero basic salary is allowed — it is a real case, unlike a blank one', () => {
    const r = g.calculateEmployee(
      { id: 'EG888', basic: 0, workingDays: 30, hireDate: D(2020, 1, 1) },
      JAN_START, JAN_END);
    eq(r.worthBasic, 0);
    eq(isNaN(r.netSalary), false);
  });
});
