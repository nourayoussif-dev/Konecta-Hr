/**
 * KONECTA EGYPT — PAYROLL PRE-FLIGHT CHECKS
 * Run this BEFORE extracting the payroll file each month.
 *
 * Every check below corresponds to something that actually went wrong in
 * August 2026. The figure beside each is what it cost or nearly cost.
 */

var CHK = {
  MASTER_SHEET  : 'EMPLOYEES',
  ADDON_SHEET   : 'ADD-ONS',
  RESULTS_SHEET : 'PAYROLL CHECKS'
};

/** Column letters/headers on the master. Adjust if the sheet changes. */
function chkCols_(headers) {
  var m = {};
  for (var i = 0; i < headers.length; i++) m[String(headers[i]).trim()] = i;
  return m;
}

/**
 * MAIN. Run this, then read the PAYROLL CHECKS tab.
 * monthStart / monthEnd as Date objects, e.g. new Date(2026,7,1), new Date(2026,7,31)
 */
function runPayrollChecks_(monthStart, monthEnd) {
  var ss = ss_();
  var master = ss.getSheetByName(CHK.MASTER_SHEET).getDataRange().getValues();
  var head = chkCols_(master[0]);
  var addonSheet = ss.getSheetByName(CHK.ADDON_SHEET);
  var addons = addonSheet ? addonSheet.getDataRange().getValues() : [[]];
  var aHead = chkCols_(addons[0] || []);

  // who has an add-on this month
  var hasAddOn = {};
  for (var a = 1; a < addons.length; a++) {
    var id = addons[a][aHead['employee_id']];
    if (id) hasAddOn[String(id).trim()] = true;
  }

  var issues = [];
  function flag(sev, id, name, check, detail, action) {
    issues.push([sev, id, name, check, detail, action]);
  }

  for (var r = 1; r < master.length; r++) {
    var row  = master[r];
    var id   = String(row[head['employee_id']] || '').trim();
    if (!id) continue;
    var name = row[head['full_name_en']];
    var st   = String(row[head['record_status']] || '').trim();
    var exit = row[head['exit_date']];
    var hire = row[head['hire_date']];
    var basic= row[head['basic_salary']];
    var iban = row[head['iban']];
    var acct = row[head['account_number']];
    var ins  = String(row[head['insurance_number']] || '').trim().toLowerCase();
    var kfreq= String(row[head['kpi_frequency']] || '');
    var ktgt = row[head['kpi_target']];
    var closed = (st === 'Closed' || st === 'Cleared');
    var live = !closed;

    // 1. CLOSED WITH NO EXIT DATE
    //    August: 21 people stayed on the payroll file. 409,031.59 nearly paid.
    if (closed && !exit)
      flag('BLOCK', id, name, 'Closed with no exit date',
           'Status is ' + st + ' but exit_date is empty',
           'Payroll stops on the DATE, not the status. Add the exit date.');

    // 2. EXIT DATE STORED AS TEXT
    if (exit && !(exit instanceof Date))
      flag('BLOCK', id, name, 'Exit date is not a date',
           'exit_date reads "' + exit + '"',
           'Retype as a real date. Text in a date field behaves unpredictably.');

    // 3. ACTIVE BUT CARRYING AN EXIT DATE
    if (st === 'Active' && exit instanceof Date)
      flag('WARN', id, name, 'Active with an exit date',
           'Active but exit_date is ' + Utilities.formatDate(exit, 'GMT', 'dd/MM/yyyy'),
           'Contradiction. Either clear the date or close the record.');

    // 4. NO BANK DETAILS — the payment will simply fail
    //    August: 17 people, 340,390.37 that could not be transferred.
    if (live && (!iban || !acct))
      flag('BLOCK', id, name, 'No bank details',
           (!iban ? 'IBAN missing. ' : '') + (!acct ? 'Account number missing.' : ''),
           'Collect from the employee. They cannot be paid without it.');

    // 5. NO BASIC SALARY
    if (live && !basic)
      flag('BLOCK', id, name, 'No basic salary',
           'basic_salary is empty',
           'Every payroll calculation depends on this field.');

    // 6. MONTHLY KPI WITH NO TARGET — pays ZERO and nothing flags it
    //    August: 22 people would have been paid nothing.
    if (live && kfreq.toLowerCase().indexOf('bsc monthly') === 0 && !ktgt)
      flag('BLOCK', id, name, 'Monthly KPI with no target',
           'kpi_frequency is "' + kfreq + '" but kpi_target is empty',
           'Attainment multiplies to zero. Populate from the offer letter.');

    // 7. ADD-ON BUT THE PERSON IS EXCLUDED FROM PAYROLL
    if (hasAddOn[id] && closed && !exit)
      flag('WARN', id, name, 'Add-on with nowhere to post',
           'Has an add-on this month but is closed with no exit date',
           'Their money cannot be paid until the record is resolved.');

    // 8. MISSING INSURANCE NUMBER
    if (st === 'Active' && (ins === '' || ins === 'tbc' || ins === '0' || ins === 'n0'))
      flag('INFO', id, name, 'No social insurance number',
           'insurance_number reads "' + (row[head['insurance_number']] || 'blank') + '"',
           'Chase the employee. Required for NOSI registration.');
  }

  writeCheckResults_(ss, issues, monthStart);
  return issues.length;
}

function writeCheckResults_(ss, issues, monthStart) {
  var sh = ss.getSheetByName(CHK.RESULTS_SHEET);
  if (!sh) sh = ss.insertSheet(CHK.RESULTS_SHEET);
  sh.clear();

  var order = { BLOCK: 1, WARN: 2, INFO: 3 };
  issues.sort(function(a, b) {
    if (order[a[0]] !== order[b[0]]) return order[a[0]] - order[b[0]];
    return String(a[1]) > String(b[1]) ? 1 : -1;
  });

  var blocks = issues.filter(function(i){ return i[0] === 'BLOCK'; }).length;
  var warns  = issues.filter(function(i){ return i[0] === 'WARN';  }).length;

  sh.getRange('A1').setValue('PAYROLL PRE-FLIGHT CHECKS — '
      + Utilities.formatDate(monthStart, 'GMT', 'MMMM yyyy'))
    .setFontSize(14).setFontWeight('bold').setFontColor('#2800C8');
  sh.getRange('A2').setValue(
      blocks + ' BLOCKING, ' + warns + ' warnings, ' + issues.length + ' total.  '
      + 'Run at ' + Utilities.formatDate(new Date(), 'GMT', 'dd/MM/yyyy HH:mm')
      + '.  Clear every BLOCK before extracting the payroll file.')
    .setFontSize(9).setFontStyle('italic').setFontColor('#666666');

  var header = ['Severity','Employee ID','Name','Check','Detail','What to do','Fixed?'];
  sh.getRange(4,1,1,header.length).setValues([header])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2800C8')
    .setHorizontalAlignment('center');

  if (issues.length) {
    var rows = issues.map(function(i){ return i.concat(['']); });
    sh.getRange(5,1,rows.length,header.length).setValues(rows);
    for (var i = 0; i < issues.length; i++) {
      var c = issues[i][0] === 'BLOCK' ? '#FCE4E4'
            : issues[i][0] === 'WARN'  ? '#FFF2CC' : '#D9E1F2';
      sh.getRange(5+i,1).setBackground(c).setFontWeight('bold');
    }
  } else {
    sh.getRange(5,1).setValue('No issues found.').setFontWeight('bold');
  }

  sh.setColumnWidth(1,90);  sh.setColumnWidth(2,110); sh.setColumnWidth(3,240);
  sh.setColumnWidth(4,230); sh.setColumnWidth(5,320); sh.setColumnWidth(6,360);
  sh.setColumnWidth(7,70);
  sh.setFrozenRows(4);
    var existing = sh.getFilter();
  if (existing) existing.remove();
  sh.getRange(4,1,Math.max(issues.length,1)+1,header.length).createFilter();
}

/** Convenience: run for the current month. */
function runChecksThisMonth() {
  assertNotDirectCall_();   // monthly trigger, not a web endpoint
  var now = new Date();
  var s = new Date(now.getFullYear(), now.getMonth(), 1);
  var e = new Date(now.getFullYear(), now.getMonth()+1, 0);
  var n = runPayrollChecks_(s, e);
    Logger.log(n + ' issues found. See the PAYROLL CHECKS tab.');
}

function whichFile_() {
  var ss = ss_();
  if (!ss) { Logger.log('NOT BOUND — no active spreadsheet'); return; }
  Logger.log('Bound to: ' + ss.getName());
  Logger.log('ID: ' + ss.getId());
}
