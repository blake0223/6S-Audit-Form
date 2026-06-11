/**
 * AuditId — generates a unique Audit ID and stamps it onto the form so each
 * printout is traceable. The same ID is typed back into "Record audit results".
 *
 * Placement: if a cell labeled "Audit ID" exists in the top rows, the ID goes in
 * the cell to its right. Otherwise the ID is written as "Audit ID: <id>" into the
 * first empty top cell (rows 1–3, above the print-hidden rows 4–5). Re-printing
 * reuses the same cell. Add an "Audit ID" label cell yourself to control placement.
 */
function newAuditId_() {
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var day = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // no easily-confused chars
  var suffix = '';
  for (var i = 0; i < 4; i++) suffix += chars.charAt(Math.floor(Math.random() * chars.length));
  return '6S-' + day + '-' + suffix;
}

/** Write the ID onto the sheet; returns the A1 notation of the cell used. */
function stampAuditId_(sheet, id) {
  var rows = Math.max(Math.min(sheet.getLastRow(), 12), 3);
  var cols = Math.max(Math.min(sheet.getLastColumn(), 12), 8);
  var vals = sheet.getRange(1, 1, rows, cols).getValues();

  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      var t = String(vals[r][c]).trim();
      if (/audit\s*id/i.test(t)) {
        if (t.toLowerCase().replace(/[:\s]/g, '') === 'auditid') {
          var tc = Math.min(c + 2, cols);                 // pure label → write to the right
          sheet.getRange(r + 1, tc).setValue(id);
          return sheet.getRange(r + 1, tc).getA1Notation();
        }
        sheet.getRange(r + 1, c + 1).setValue('Audit ID: ' + id); // combined cell → overwrite
        return sheet.getRange(r + 1, c + 1).getA1Notation();
      }
    }
  }
  for (var rr = 0; rr < Math.min(3, vals.length); rr++) {
    for (var cc = 0; cc < vals[rr].length; cc++) {
      if (String(vals[rr][cc]).trim() === '') {
        sheet.getRange(rr + 1, cc + 1).setValue('Audit ID: ' + id);
        return sheet.getRange(rr + 1, cc + 1).getA1Notation();
      }
    }
  }
  return '';
}
