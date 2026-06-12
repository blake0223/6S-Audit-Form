/**
 * PrintForm — native pickers (form type, then facility), then open a print-ready
 * PDF in the browser.
 *
 * For the Monthly audit, the printout drops column A (No.), the Total Score column,
 * and rows 2–5 (the summary block). Those are non-rectangular, so we make a hidden
 * throwaway copy of the tab, delete those rows/columns ON THE COPY, and let the
 * browser fetch the copy's PDF. This runs in the menu context (no google.script.run)
 * and uses no UrlFetch — so the only permission needed is "edit this spreadsheet".
 * The copy is replaced on each print.
 *
 * Menu: 6S Audit Tools ▸ Print blank form
 */
function printForm() {
  var ui = SpreadsheetApp.getUi();
  var typeLabel = pickOption_(['Monthly audit', 'Weekly checklist'], 'Print blank form — which form');
  if (!typeLabel) return;
  var type = (typeLabel === 'Weekly checklist') ? 'checklist' : 'monthly';
  var sheetName = pickFacility_(type, 'Print blank form');
  if (!sheetName) return;

  var url;
  try { url = buildPrintUrl_(sheetName, type); }
  catch (e) { ui.alert('Error: ' + e.message); return; }

  var html = HtmlService.createHtmlOutput(
      '<body style="font-family:Arial;margin:0;padding:16px;font-size:13px;color:#374151">'
    + 'Opening the print-ready PDF in a new tab…'
    + '<script>window.open(' + JSON.stringify(url) + ',"_blank");'
    + 'setTimeout(google.script.host.close,500);</script></body>')
    .setWidth(300).setHeight(90);
  ui.showModelessDialog(html, 'Print');
}

/** Build the export URL, trimming a hidden copy for Monthly audits. */
function buildPrintUrl_(sheetName, type) {
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

    copy.hideSheet();
    SpreadsheetApp.flush();
    gid = copy.getSheetId();
  } else {
    gid = src.getSheetId();
  }

  return 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?'
    + 'format=pdf&gid=' + gid
    + '&portrait=' + (type === 'monthly' ? 'true' : 'false')
    + '&fitw=true&size=letter&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=true'
    + '&top_margin=0.50&bottom_margin=0.50&left_margin=0.40&right_margin=0.40';
}

/** Strip the form-type words from a tab name to get a clean facility label. */
function facilityLabel_(name) {
  return String(name)
    .replace(/6S/ig, '')
    .replace(/monthly audit sheet|monthly audit|daily checklist|weekly checklist|checklist|facility summary|sheet/ig, '')
    .replace(/\s+/g, ' ').trim() || name;
}
