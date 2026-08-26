/**
 * KONECTA EGYPT — PAYSLIP
 * Renders one employee's payslip for a given month.
 *
 * VISIBILITY RULES (agreed 21/08/2026)
 *   Always shown  : working days, basic, total salary, total deductions, net,
 *                   and EVERY deduction regardless of value.
 *   Social insurance always shows. If zero it carries the reason, because a
 *                   new mid-month joiner needs to know registration starts
 *                   the following month rather than assume they were missed.
 *   KPI           : shown even at zero, but ONLY for employees on BSC Monthly.
 *                   Annual and semi-annual schemes have no monthly KPI, so a
 *                   zero line there would be misleading.
 *   Everything else: hidden when zero, so nobody reads forty empty rows.
 */

var PS = { PAYROLL_PREFIX: 'PAYROLL_', MASTER: 'EMPLOYEES' };

/** Earnings, in the order they appear. net:true means the figure is grossed up. */
var PS_EARNINGS = [
  { col:'Final Basic salary (EGP)',                 label:'Basic salary' },
  { col:'Retro / Back-pay',                         label:'Back-pay for a previous month' },
  { col:'KPI',                                      label:'KPI',            alwaysIfMonthly:true },
  { col:'Incentive',                                label:'Incentive' },
  { col:'Bonus (Annual)',                           label:'Annual bonus' },
  { col:'Transportation ( 3000 net)',               label:'Transportation', net:true },
  { col:'Shift Allowance ( 2000 net per month)',    label:'Shift allowance', net:true },
  { col:'On Call ( 2000 net per month)',            label:'On-call allowance', net:true },
  { col:'Bonus & Other',                            label:'Allowance',      net:true },
  { col:'Referral Bonus',                           label:'Referral bonus', net:true },
  // 'Amount  Overtime paid' is the provider's TOTAL of overtime hours plus
  // holiday and weekend days. 'Total Amount Days' is the days portion of that
  // same total. Showing both double-counts on the payslip even though the
  // employee is paid once. So: show the days portion as its own line, and show
  // overtime as the remainder.
  { col:'Amount  Overtime paid', label:'Overtime', minus:'Total Amount Days' },
  { col:'Total Amount Days',     label:'Public holidays and weekends worked' },
  { col:'Refunded Amounts',                         label:'Refunded amounts' },
  { col:'Remaining Vacation Amount',                label:'Unused annual leave', untaxed:true },
  { col:'End of Service Indemnity Amount',          label:'End of service',     untaxed:true },
  { col:'Termination Settlement Payments',          label:'Termination settlement', untaxed:true },
  { col:'Non taxable amounts',                      label:'Non-taxable payment',    untaxed:true },
  { col:'Paid Loan (Non-Taxable) – To Be Settled Accounting-Wise',
                                                    label:'Loan advanced to you',   untaxed:true }
];

/** Deductions. These ALWAYS appear, even at zero. */
var PS_DEDUCTIONS = [
  { col:'Social Security',                    label:'Social insurance', always:true },
  { col:'Final Monthly Tax',                  label:'Income tax',       always:true },
  { col:'Final Martyrs Fund',                 label:'Martyrs fund',     always:true },
  { col:'Amount Disiplinary Deduction',       label:'Disciplinary deduction' },
  { col:'Ded (salary in advance/Medical)',    label:'Loan repayment' },
  { col:'Medical Insurance',                  label:'Medical insurance — additional dependant' },
  { col:'Undeserved wages',                   label:'Recovery of an overpayment' },
  { col:'Ded (Taxable)',                      label:'Other deduction' }
];

function num_(v){ return (typeof v === 'number' && !isNaN(v)) ? v : 0; }
function money_(v){ return Utilities.formatString('%s', Utilities.formatString('%,.2f', v)); }

/**
 * Build the payslip for one employee.
 * period e.g. '2026_08'. Returns an object the UI renders.
 */
function buildPayslip(employeeId, period) {
  var sh = archiveTab_(period);            // published month, in the archive file
  if (!sh) throw new Error('No payroll has been published for that month yet.');

  var data = sh.getDataRange().getValues();
  var head = {}, hrow = data[1];              // row 2 holds the headings
  for (var i = 0; i < hrow.length; i++) head[String(hrow[i]).replace(/\n/g,' ').trim()] = i;

  var row = null;
  for (var r = 2; r < data.length; r++) {
    if (String(data[r][0]).trim() === employeeId) { row = data[r]; break; }
  }
  if (!row) throw new Error('No payslip found for this employee in that month.');

  function val(colName){ return head[colName] === undefined ? 0 : num_(row[head[colName]]); }
  function txt(colName){ return head[colName] === undefined ? '' : row[head[colName]]; }

  // is this person on a monthly KPI scheme?
  var master = ss_().getSheetByName(PS.MASTER).getDataRange().getValues();
  var mh = {}; for (var k = 0; k < master[0].length; k++) mh[String(master[0][k]).trim()] = k;
  var monthlyKPI = false, joinedThisMonth = false;
  for (var m = 1; m < master.length; m++) {
    if (String(master[m][mh['employee_id']]).trim() === employeeId) {
      monthlyKPI = String(master[m][mh['kpi_frequency']] || '')
                     .toLowerCase().indexOf('bsc monthly') === 0;
      var hire = master[m][mh['hire_date']];
      if (hire instanceof Date) {
        var y = Number(period.split('_')[0]), mo = Number(period.split('_')[1]) - 1;
        joinedThisMonth = (hire.getFullYear() === y && hire.getMonth() === mo && hire.getDate() > 1);
      }
      break;
    }
  }

  var earnings = [];
  PS_EARNINGS.forEach(function(e){
    var v = val(e.col);
    if (e.minus) v = v - val(e.minus);      // avoid double-counting, see note above
    var show = v !== 0 || (e.alwaysIfMonthly && monthlyKPI);
    if (!show) return;
    var note = '';
    if (e.net)     note = 'Paid so that the agreed net amount reaches you; tax on it is covered.';
    if (e.untaxed) note = 'Not subject to income tax.';
    earnings.push({ label:e.label, amount:v, note:note });
  });

  var deductions = [];
  PS_DEDUCTIONS.forEach(function(d){
    var v = val(d.col);
    if (v === 0 && !d.always) return;
    var note = '';
    if (d.col === 'Social Security' && v === 0) {
      note = joinedThisMonth
        ? 'You joined part-way through this month. Social insurance registration begins with your first full month.'
        : 'No social insurance was due this month.';
    }
    deductions.push({ label:d.label, amount:v, note:note });
  });

  return {
    period        : period.replace('_',' / '),
    employeeId    : employeeId,
    name          : txt('Employee Name'),
    jobTitle      : txt('Job Title'),
    workingDays   : val('Final No. of Days'),
    earnings      : earnings,
    totalSalary   : val('Total Income'),
    deductions    : deductions,
    totalDeducted : val('Total Deducted'),
    netSalary     : val('Final Net'),
    bankName      : txt('Bank Name'),
    accountMasked : maskAccount_(txt('ACCOUNT NUMBER'))
  };
}

/** Show only the last four digits. A payslip never needs the full number. */
function maskAccount_(acct) {
  var s = String(acct || '');
  if (s.length <= 4) return s;
  return '•••• ' + s.slice(-4);
}

/**
 * Called by the UI. An employee gets their own payslip and nobody else's;
 * HR can retrieve any. Identity comes from the signed-in account, never
 * from a field on the sheet.
 */
function getMyPayslip(period, employeeIdIfHR) {
  var me = currentUser_();
  if (isHR_() && employeeIdIfHR) return buildPayslip(employeeIdIfHR, period);

  var ss = ss_();
  var master = ss.getSheetByName(PS.MASTER).getDataRange().getValues();
  var mh = {}; for (var k = 0; k < master[0].length; k++) mh[String(master[0][k]).trim()] = k;
  for (var m = 1; m < master.length; m++) {
    var email = String(master[m][mh['konecta_email']] || '').toLowerCase().trim();
    if (email === me) return buildPayslip(String(master[m][mh['employee_id']]).trim(), period);
  }
  throw new Error('We could not find an employee record for your account. Please contact HR.');
}

// listPayslipPeriods() lives in PayrollArchive.gs — it scans the archive files.
