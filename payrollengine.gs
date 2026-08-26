/**
 * KONECTA EGYPT — PAYROLL CALCULATION ENGINE
 * Google Apps Script. Paste into the HR tool as a library file.
 *
 * Validated 20/08/2026 against the July 2026 payroll produced by the
 * outsourced provider: Worth Basic, overtime and martyrs fund 100% match;
 * income tax 99.33%; social insurance 99.0%.
 *
 * Every constant below is taken from the CONFIG sheet of the payroll
 * workbook, not assumed. Update them here if the law or the scheme changes.
 */

// ============ CONFIGURATION =================================================
var CFG = {
  PERSONAL_EXEMPTION_ANNUAL : 20000,     // annual personal exemption
  MARTYRS_RATE              : 0.0005,    // 0.05% of total salary
  SI_EMPLOYEE               : 0.11,      // employee social insurance share
  SI_EMPLOYER               : 0.1875,    // employer social insurance share
  INS_WAGE_MAX              : 16700,     // insurance wage ceiling
  INS_WAGE_MIN              : 2300,      // insurance wage floor
  EMERGENCY_FUND_FLAT       : 23.70,     // flat emergency fund per insured employee
  STANDARD_MONTH_DAYS       : 30,
  WORK_HOURS_PER_DAY        : 8,
  OT_DAY                    : 1.35,      // overtime multiplier, day shift
  OT_NIGHT                  : 1.70,      // overtime multiplier, night shift
  OT_PUBLIC_HOLIDAY         : 1.00,      // public holiday worked, per day
  TRANSPORT_NET_PER_MONTH   : 3000,      // net, per 22 working days
  SHIFT_NET_PER_MONTH       : 2000,      // net, per 22 working days
  ONCALL_NET_PER_MONTH      : 2000,      // net, per 30-day month
  IFTAR_NET_PER_DAY         : 250,       // net
  WORKING_DAYS_PER_MONTH    : 22,        // for transport and shift allowance
  LEAVE_DAYS_PER_YEAR       : 15         // default; individual contracts may differ
};

// Income tax brackets. LOWER is the bracket floor, DELTA the marginal step.
var BRACKET_LOWER = [0, 40000, 55000, 70000, 200000, 400000, 1200000];
var RATE_DELTA    = [0, 0.10,  0.05,  0.05,  0.025,  0.025,  0.025];

// Bracket cancellation. Above each annual threshold the lower brackets are
// cancelled, so high earners lose the benefit of the lower rates.
var CANCEL = [[0,1],[600000,2],[700000,3],[800000,4],[900000,5],[1200000,6]];

// Employees exempt from social insurance, with the reason on record.
// EG0009 Walid Fairouz — army retirement, already covered.
var SI_EXEMPT = ['EG0009'];

// ============ CORE FUNCTIONS ================================================

/** Which bracket the tax starts from, given annual taxable income. */
function firstBracketApplied(annualTaxable) {
  var b = 1;
  for (var i = 0; i < CANCEL.length; i++) {
    if (annualTaxable >= CANCEL[i][0]) b = CANCEL[i][1];
  }
  return b;
}

/**
 * Monthly income tax.
 * Annualises the monthly base, applies bracket cancellation, taxes
 * marginally, then divides by 12.
 *
 * nonTaxableInTotal — amounts inside Total Salary that are NOT taxable:
 *   leave encashment (Remaining Vacation Amount) and WFR compensation
 *   (Termination Settlement Payments). Both confirmed exempt 20/08/2026.
 */
function incomeTaxMonthly(totalSalary, employeeSI, dedTaxable, nonTaxableInTotal) {
  dedTaxable = dedTaxable || 0;
  nonTaxableInTotal = nonTaxableInTotal || 0;
  var taxable = totalSalary - nonTaxableInTotal;
  var monthlyBase = Math.max(0, taxable - employeeSI
                     - CFG.PERSONAL_EXEMPTION_ANNUAL / 12 - dedTaxable);
  var annual = monthlyBase * 12;
  var fb = firstBracketApplied(annual);
  var tax = 0;
  for (var i = 0; i < 7; i++) {
    var lower = (i + 1 > fb) ? BRACKET_LOWER[i] : 0;
    if (annual > lower) tax += (annual - lower) * RATE_DELTA[i];
  }
  return tax / 12;
}

/**
 * Insurance wage for the month.
 * JOINER  hired on day 2 or later of the payroll month -> ZERO. NOSI will not
 *         accept a mid-month registration, so cover starts the next full month.
 * LEAVER  exit on day 3 or later of the payroll month -> still charged for the
 *         whole month, so the employee's cover stays continuous.
 * Zero working days and not a leaver this month -> ZERO.
 */
function insuranceWage(employeeId, basic, workingDays, hireDate, exitDate,
                       monthStart, monthEnd) {
  if (SI_EXEMPT.indexOf(employeeId) !== -1) return 0;
  if (hireDate && hireDate >= monthStart && hireDate <= monthEnd
      && hireDate.getDate() > 1) return 0;
  var leaverThisMonth = exitDate && exitDate >= monthStart
                        && exitDate <= monthEnd && exitDate.getDate() >= 3;
  if ((!workingDays || workingDays <= 0) && !leaverThisMonth) return 0;
  return Math.min(Math.max(basic, CFG.INS_WAGE_MIN), CFG.INS_WAGE_MAX);
}

/** Overtime and public holidays worked, priced from basic salary. */
function overtimeAmount(basic, hours135, hours170, publicHolidayDays, weekendDays) {
  var hourly = basic / CFG.STANDARD_MONTH_DAYS / CFG.WORK_HOURS_PER_DAY;
  var daily  = basic / CFG.STANDARD_MONTH_DAYS;
  return hourly * ((hours135 || 0) * CFG.OT_DAY + (hours170 || 0) * CFG.OT_NIGHT)
       + daily  * ((publicHolidayDays || 0) * CFG.OT_PUBLIC_HOLIDAY + (weekendDays || 0));
}

/**
 * Gross-up. Finds the gross addition that delivers a promised NET amount,
 * at the employee's own marginal rate.
 * Applies to: Transportation, Shift Allowance, On Call, Referral Bonus,
 * and Bonus & Other (the Iftar coupon).
 */
function grossUp(netPromised, taxableBefore, employeeSI, dedTaxable) {
  if (!netPromised) return 0;
  dedTaxable = dedTaxable || 0;
  function netAt(extra) {
    var total = taxableBefore + extra;
    var tax = incomeTaxMonthly(total, employeeSI, dedTaxable, 0);
    return total - employeeSI - tax - total * CFG.MARTYRS_RATE;
  }
  var base = netAt(0), lo = 0, hi = Math.max(netPromised * 3, 1000);
  for (var i = 0; i < 200; i++) {
    var mid = (lo + hi) / 2;
    if (netAt(mid) - base < netPromised) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

// ============ DAY COUNTS ====================================================

/**
 * Days of THIS month, on the Egyptian 30-day month convention.
 *
 *   full month  -> 30, even in a 31-day calendar month. The 31st is not a
 *                  payroll day for anyone, so a hire on the 31st is 0 days
 *                  here and a full 30 next month. That is deliberate.
 *   leaver      -> MIN(30, day of exit)
 *   joiner      -> 31 - day of hire   (hired on the 15th = 15th..30th = 16 days)
 *
 * FIXED 26/08/2026 — two cases the earlier version got wrong:
 *
 *   BOTH in the same month. The leaver branch was tested first and returned
 *   MIN(30, exit day), counting from the 1st — as though the person had been
 *   employed all along. Hired on the 10th and exited on the 20th paid 20 days
 *   for 11 worked. That is the exact shape of a no-show or a drop-out, which
 *   is why the app has a NO_SHOW module, and it was caught only when someone
 *   remembered to run holdPaymentFor_.
 *
 *   HIRE AFTER THIS MONTH. A hire date past monthEnd fell through every
 *   branch to `return 30`, paying a full month to somebody who had not
 *   started. Running payroll early, or a pushed-back start date, both reach
 *   this.
 */
function daysThisMonth(hireDate, exitDate, monthStart, monthEnd) {
  // Already gone, or not yet started: nothing to pay.
  if (exitDate && exitDate < monthStart) return 0;
  if (hireDate && hireDate > monthEnd)   return 0;

  var joins  = hireDate && hireDate >= monthStart && hireDate <= monthEnd;
  var leaves = exitDate && exitDate <= monthEnd;   // exit >= monthStart, checked above

  // Employed for only part of the month at both ends: pay the days between,
  // inclusive of the hire date itself.
  if (joins && leaves) {
    return Math.max(0, Math.min(30, exitDate.getDate()) - hireDate.getDate() + 1);
  }
  if (leaves) return Math.min(30, exitDate.getDate());
  if (joins)  return 31 - hireDate.getDate();
  return 30;
}

/**
 * ARREARS days — days worked LAST month that were never paid.
 *
 * LEARNED 21/08/2026, the hard way. August missed 14 people this way,
 * 190,659.48 EGP. My first version triggered only for hires after the 26th,
 * inferred from two examples. It was wrong: EG0836 was hired on the 12th and
 * still received nothing for July.
 *
 * So do NOT guess from a cut-off date. The only reliable test is whether the
 * person actually appeared in last month's payroll. A late contract, a missed
 * handover or a cut-off all produce the same result, and only the payroll
 * itself knows which happened.
 *
 * wasPaidLastMonth — pass TRUE only if the employee appears in the previous
 *   month's published payroll. Look it up; never assume.
 */
function arrearsDays(hireDate, prevMonthStart, prevMonthEnd, wasPaidLastMonth) {
  if (!hireDate) return 0;
  if (hireDate < prevMonthStart || hireDate > prevMonthEnd) return 0;
  if (wasPaidLastMonth) return 0;
  return 31 - hireDate.getDate();
}

/**
 * Was this employee in the previous month's published payroll?
 * Reads the archive, so it reflects what was actually paid rather than what
 * anyone believes was paid.
 */
function wasPaidInPeriod(employeeId, period) {
  var sh;
  try { sh = archiveTab_(period); } catch (e) { return false; }
  if (!sh) return false;
  var ids = sh.getRange(3, 1, Math.max(sh.getLastRow() - 2, 1), 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === employeeId) return true;
  }
  return false;
}

// ============ FULL CALCULATION ==============================================

/**
 * Calculate one employee.
 * emp must carry: id, basic, workingDays, arrearsDays, hireDate, exitDate,
 *   grossAddOns, netAddOns, nonTaxableAmounts, nonTaxableInTotal,
 *   dedTaxable, otherDeductions, hours135, hours170, publicHolidayDays, weekendDays
 */
function calculateEmployee(emp, monthStart, monthEnd) {
  // Refuse a record we cannot price, rather than paying out NaN.
  //
  // A blank or non-numeric basic_salary used to propagate NaN through
  // worthBasic, totalSalary, incomeTax and netSalary without throwing, so the
  // employee received a payslip reading NaN and the run reported success.
  // Stopping here forces the record to be fixed before anyone is paid.
  if (typeof emp.basic !== 'number' || !isFinite(emp.basic)) {
    throw new Error('Payroll stopped: ' + (emp.id || 'an employee') +
      ' has no usable basic salary (' + JSON.stringify(emp.basic) + '). ' +
      'Fix basic_salary on the EMPLOYEES tab and run again.');
  }
  if (emp.basic < 0) {
    throw new Error('Payroll stopped: ' + (emp.id || 'an employee') +
      ' has a negative basic salary (' + emp.basic + ').');
  }

  // arrearsDays must come from arrearsDays(), which checks the previous
  // month's payroll. Never default it to zero without looking.
  var totalDays = (emp.workingDays || 0) + (emp.arrearsDays || 0);
  var worthBasic = emp.basic * totalDays / CFG.STANDARD_MONTH_DAYS;
  var ot = overtimeAmount(emp.basic, emp.hours135, emp.hours170,
                          emp.publicHolidayDays, emp.weekendDays);
  var iw = insuranceWage(emp.id, emp.basic, emp.workingDays, emp.hireDate,
                         emp.exitDate, monthStart, monthEnd);
  var si = iw * CFG.SI_EMPLOYEE;
  var taxableBefore = worthBasic + ot + (emp.grossAddOns || 0);
  var up = grossUp(emp.netAddOns || 0, taxableBefore, si, emp.dedTaxable);
  var total = taxableBefore + up + (emp.nonTaxableInTotal || 0);
  var tax = incomeTaxMonthly(total, si, emp.dedTaxable, emp.nonTaxableInTotal);
  var mf = total * CFG.MARTYRS_RATE;
  var totalDeductions = si + tax + mf + (emp.dedTaxable || 0) + (emp.otherDeductions || 0);
  var net = total - totalDeductions + (emp.nonTaxableAmounts || 0);
  return {
    workingDaysTotal : totalDays,
    worthBasic       : worthBasic,
    overtimeAmount   : ot,
    insuranceWage    : iw,
    employeeSI       : si,
    grossUpAmount    : up,
    totalSalary      : total,
    incomeTax        : tax,
    martyrsFund      : mf,
    totalDeductions  : totalDeductions,
    netSalary        : net,
    employerSI       : iw * CFG.SI_EMPLOYER,
    emergencyFund    : iw > 0 ? CFG.EMERGENCY_FUND_FLAT : 0,
    totalCost        : total + iw * CFG.SI_EMPLOYER + (iw > 0 ? CFG.EMERGENCY_FUND_FLAT : 0)
  };
}

// ============ ADD-ON PRICING ================================================

function priceTransportation(officeDays, totalWorkingDays) {
  return CFG.TRANSPORT_NET_PER_MONTH * officeDays / (totalWorkingDays || CFG.WORKING_DAYS_PER_MONTH);
}
function priceShiftAllowance(days) {
  return CFG.SHIFT_NET_PER_MONTH * days / CFG.WORKING_DAYS_PER_MONTH;
}
function priceOnCall(days) {
  return CFG.ONCALL_NET_PER_MONTH * days / CFG.STANDARD_MONTH_DAYS;
}
function priceIftar(days) {
  return CFG.IFTAR_NET_PER_DAY * days;
}
function priceKPI(kpiTarget, attainmentPercent, daysWorked) {
  var full = kpiTarget * attainmentPercent;
  return daysWorked ? full * daysWorked / CFG.STANDARD_MONTH_DAYS : full;
}
/** Leave encashment. NOT taxable. entitlementDaysPerYear varies by contract. */
function priceLeaveEncashment(basic, serviceDays, entitlementDaysPerYear, daysTaken) {
  var accrued = (entitlementDaysPerYear || CFG.LEAVE_DAYS_PER_YEAR) * serviceDays / 365;
  return Math.max(0, accrued - (daysTaken || 0)) * basic / CFG.STANDARD_MONTH_DAYS;
}
