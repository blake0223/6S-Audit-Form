/**
 * PrintForm — native pickers (form type, then facility), then open a print-ready
 * PDF in the browser.
 *
 * For the Monthly audit, the printout drops column A (No.), the Total Score column,
 * and rows 2–5 (the summary block). Those are non-rectangular, so we make a
 * throwaway "_print_tmp" copy of the tab, delete those rows/columns ON THE COPY,
 * and the user clicks a link to fetch the copy's PDF. This runs in the menu context
 * (no google.script.run) and uses no UrlFetch — so the only permission needed is
 * "edit this spreadsheet". The copy is visible (hidden sheets don't export by gid)
 * and is replaced on each print.
 *
 * Menu: 6S Audit Tools ▸ Print blank form
 */
/** Build the PDF export URL, trimming a throwaway copy for Monthly audits. */
function getPrintUrl_(sheetName, type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName(sheetName);
  if (!src) throw new Error('Tab not found: ' + sheetName);

  var gid;
  if (type === 'monthly') {
    var old = ss.getSheetByName('_print_tmp');
    if (old) ss.deleteSheet(old);
    var copy = src.copyTo(ss).setName('_print_tmp');

    var headerRow = findHeaderRow_(copy, copy.getLastRow());
    var totalCol = findColByHeader_(copy, headerRow, 'total score');
    if (totalCol) copy.deleteColumn(totalCol); // Total Score column
    copy.deleteColumn(1);                       // column A (No.)
    copy.deleteRows(2, 4);                       // rows 2–5 (summary block)

    // NOTE: do NOT hide this sheet — Google won't render a PDF export from a hidden
    // tab. It stays visible briefly and is removed by the cleanup after the PDF opens
    // (and by the start-of-print deletion above as a backstop).
    SpreadsheetApp.flush();
    gid = copy.getSheetId();
  } else {
    gid = src.getSheetId();
  }

  // Pin the export to the account that owns/uses the sheet. Without this, a user
  // signed into multiple Google accounts opens the link under their default
  // account and gets a "broken" / no-access page instead of the PDF.
  var authuser = '';
  try { authuser = Session.getActiveUser().getEmail() || ''; } catch (e) { /* not available */ }

  return 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?'
    + 'format=pdf&gid=' + gid
    + '&portrait=' + (type === 'monthly' ? 'true' : 'false')
    + '&fitw=true&size=letter&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=true'
    + '&top_margin=0.50&bottom_margin=0.50&left_margin=0.40&right_margin=0.40'
    + (authuser ? '&authuser=' + encodeURIComponent(authuser) : '');
}

/** Strip the form-type words from a tab name to get a clean facility label. */
function facilityLabel_(name) {
  return String(name)
    .replace(/6S/ig, '')
    .replace(/monthly audit sheet|monthly audit|daily checklist|weekly checklist|checklist|facility summary|sheet/ig, '')
    .replace(/\s+/g, ' ').trim() || name;
}
