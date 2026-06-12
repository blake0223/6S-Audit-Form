/**
 * AuditId — generates a unique Audit ID for a submission, e.g. 6S-20260611-K7QD.
 * Used by RecordResults when the user submits the scores typed into the sheet.
 */
function newAuditId_() {
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var day = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // no easily-confused chars
  var suffix = '';
  for (var i = 0; i < 4; i++) suffix += chars.charAt(Math.floor(Math.random() * chars.length));
  return '6S-' + day + '-' + suffix;
}
