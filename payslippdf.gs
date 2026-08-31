/**
 * KONECTA EGYPT — PAYSLIP PDF
 *
 * Builds the PDF from the same data the screen shows, so the two can never
 * disagree. A payslip whose PDF differs from the screen invites the question
 * of which one is right.
 *
 * Paste as a new script file. Requires Payslip.gs and PayrollArchive.gs.
 */

/**
 * Called by the Download button. Returns a URL the browser can open.
 * The file is written to Drive, shared with the employee, and the link
 * returned. It is theirs to keep.
 */
function makePayslipPdf(period, employeeIdIfHR) {
  var d = (isHR_() && employeeIdIfHR)
        ? buildPayslip_(employeeIdIfHR, period)
        : getMyPayslip(period);

  var html = payslipHtml_(d, period);
  var blob = Utilities.newBlob(html, 'text/html', 'payslip.html')
                      .getAs('application/pdf')
                      .setName('Konecta payslip — ' + d.name + ' — ' + period + '.pdf');

  var folder = payslipFolder_();

  // If this payslip was already generated, hand back the existing file rather
  // than making another. Without this, a slow first click produces a duplicate
  // for every impatient second click — and each one emails the employee.
  var wanted = blob.getName();
  var existing = folder.getFilesByName(wanted);
  if (existing.hasNext()) return existing.next().getUrl();

  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

  // addViewer() emails the person. They are downloading their own payslip and
  // do not need telling that it exists, so add access without the notification.
  try {
    Drive.Permissions.insert(
      { role: 'reader', type: 'user', value: currentUser_() },
      file.getId(),
      { sendNotificationEmails: false }
    );
  } catch (e) { /* advanced Drive service not enabled; owner still has access */ }

  return file.getUrl();
}

/** One folder for generated payslips, created on first use. */
function payslipFolder_() {
  var name = 'Konecta payslips';
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function psMoneyPdf_(v) {
  if (v === null || v === undefined || isNaN(v)) return '—';
  return Number(v).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function prettyPeriodPdf_(p) {
  var M = ['January','February','March','April','May','June','July',
           'August','September','October','November','December'];
  var a = String(p).split('_');
  return (M[Number(a[1]) - 1] || a[1]) + ' ' + a[0];
}

function payslipHtml_(d, period) {
  var NAVY = '#2800C8';
  function rows(list) {
    return (list || []).map(function(x) {
      return '<tr><td class="lbl">' + x.label +
             (x.note ? '<div class="note">' + x.note + '</div>' : '') +
             '</td><td class="amt">' + psMoneyPdf_(x.amount) + '</td></tr>';
    }).join('');
  }

  return '' +
  '<html><head><meta charset="utf-8"><style>' +
  'body{font-family:Arial,Helvetica,sans-serif;font-size:11pt;color:#222;margin:34px 40px}' +
  'h1{color:' + NAVY + ';font-size:17pt;margin:0 0 2px}' +
  '.sub{color:#666;font-size:9.5pt;margin:0 0 18px}' +
  'table{width:100%;border-collapse:collapse;margin-bottom:16px}' +
  'th{background:' + NAVY + ';color:#fff;text-align:left;padding:7px 10px;font-size:10pt}' +
  'th.r{text-align:right}' +
  'td{padding:6px 10px;border-bottom:1px solid #e6e6ef;vertical-align:top}' +
  'td.amt{text-align:right;white-space:nowrap}' +
  '.note{color:#6b6b80;font-size:8.5pt;margin-top:2px}' +
  '.tot td{border-top:2px solid ' + NAVY + ';border-bottom:0;font-weight:bold;color:' + NAVY + '}' +
  '.net{background:#EEEDFE;border:1.5px solid ' + NAVY + ';padding:14px 16px;margin-top:6px}' +
  '.net .l{color:' + NAVY + ';font-weight:bold;font-size:12pt}' +
  '.net .v{color:' + NAVY + ';font-weight:bold;font-size:17pt;float:right}' +
  '.meta{width:100%;margin-bottom:18px;font-size:10pt}' +
  '.meta td{border-bottom:1px solid #e6e6ef;padding:5px 0}' +
  '.meta td:last-child{text-align:right;font-weight:bold}' +
  '.foot{margin-top:26px;color:#666;font-size:8.5pt;border-top:1px solid #e6e6ef;padding-top:10px}' +
  '</style></head><body>' +

  '<h1>Konecta Egypt</h1>' +
  '<p class="sub">Payslip for ' + prettyPeriodPdf_(period) + '</p>' +

  '<table class="meta">' +
  '<tr><td>Name</td><td>' + (d.name || '—') + '</td></tr>' +
  '<tr><td>Employee ID</td><td>' + (d.employeeId || '—') + '</td></tr>' +
  '<tr><td>Job title</td><td>' + (d.jobTitle || '—') + '</td></tr>' +
  '<tr><td>Working days</td><td>' + (d.workingDays || 0) + '</td></tr>' +
  '</table>' +

  '<table><tr><th>What you earned</th><th class="r">EGP</th></tr>' +
  rows(d.earnings) +
  '<tr class="tot"><td>Total</td><td class="amt">' + psMoneyPdf_(d.totalSalary) + '</td></tr>' +
  '</table>' +

  '<table><tr><th>What was deducted</th><th class="r">EGP</th></tr>' +
  rows(d.deductions) +
  '<tr class="tot"><td>Total</td><td class="amt">' + psMoneyPdf_(d.totalDeducted) + '</td></tr>' +
  '</table>' +

  '<div class="net"><span class="l">Paid to you</span>' +
  '<span class="v">' + psMoneyPdf_(d.netSalary) + ' EGP</span>' +
  '<div style="clear:both"></div>' +
  '<div style="color:#666;font-size:9pt;margin-top:6px">' +
  (d.bankName || '') + (d.accountMasked ? ' &middot; ' + d.accountMasked : '') + '</div></div>' +

  '<div class="foot">' +
  'Generated ' + Utilities.formatDate(new Date(), 'GMT+2', 'd MMMM yyyy') + '. ' +
  'If anything here does not look right, contact Egypt HR. ' +
  'This payslip reflects the payroll as published for ' + prettyPeriodPdf_(period) +
  ' and does not change if records are updated afterwards.' +
  '</div></body></html>';
}
