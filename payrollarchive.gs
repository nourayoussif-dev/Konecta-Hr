/**
 * KONECTA EGYPT — PAYROLL ARCHIVE
 *
 * Payroll never lives in the employee master. Each month is written to a
 * separate archive spreadsheet as a frozen tab and is never touched again.
 *
 * Why separate:
 *   - the master stays the same size however many months accumulate
 *   - a published month cannot be edited by accident from the tab beside it
 *   - access to salary data can be restricted without restricting HR data
 *   - a payslip always shows what was actually paid, not what the master
 *     says today. Exit dates and salaries change after the event; the
 *     snapshot must not.
 */

// One archive per calendar year. Add the next line each January.
var PAYROLL_ARCHIVES = {
  '2026': '1I4JFlEPUbuh1OpxY_ya5bcjs_JRzRcEDhQrQw6CKXsE'
};

var ARCHIVE_PREFIX = 'PAYROLL_';

/** The archive file for a period such as '2026_08'. */
function archiveFor_(period) {
  var year = String(period).split('_')[0];
  var id = PAYROLL_ARCHIVES[year];
  if (!id) throw new Error('No payroll archive has been set up for ' + year
        + '. Create the file, then add its ID to PAYROLL_ARCHIVES.');
  return SpreadsheetApp.openById(id);
}

/** The tab for one month, or null if that month has not been published. */
function archiveTab_(period) {
  return archiveFor_(period).getSheetByName(ARCHIVE_PREFIX + period);
}

/**
 * Publish a month. values[0] is the heading row.
 * Refuses to overwrite: a published month is final. If it genuinely has to
 * be reissued, rename the existing tab first so the original survives.
 */
function publishPayrollMonth(period, values) {
  var ss = archiveFor_(period);
  var name = ARCHIVE_PREFIX + period;
  if (ss.getSheetByName(name))
    throw new Error(period + ' has already been published. Rename the existing '
      + 'tab before republishing, so the original is not lost.');

  var sh = ss.insertSheet(name);
  sh.getRange(1,1,1,1).setValue('KONECTA EGYPT — PAYROLL ' + period.replace('_',' / '))
    .setFontWeight('bold').setFontSize(13).setFontColor('#2800C8');
  sh.getRange(2,1,values.length,values[0].length).setValues(values);
  sh.getRange(2,1,1,values[0].length)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#2800C8')
    .setWrap(true).setHorizontalAlignment('center');
  sh.setFrozenRows(2);
  sh.setFrozenColumns(3);
  sh.protect().setDescription('Published payroll — do not edit')
    .setWarningOnly(true);
  return sh.getName();
}

/** Months available, newest first. */
function listPayslipPeriods() {
  var out = [];
  Object.keys(PAYROLL_ARCHIVES).forEach(function(year){
    try {
      SpreadsheetApp.openById(PAYROLL_ARCHIVES[year]).getSheets().forEach(function(s){
        var n = s.getName();
        if (n.indexOf(ARCHIVE_PREFIX) === 0) out.push(n.replace(ARCHIVE_PREFIX,''));
      });
    } catch(e) { /* an archive that cannot be opened must not break the list */ }
  });
  return out.sort().reverse();
}
