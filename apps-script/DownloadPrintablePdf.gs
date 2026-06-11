/**
 * DownloadPrintablePdf — exports the ACTIVE tab as a clean, blank, printable PDF
 * (portrait, fit-to-width, no gridlines), saves it to a "6S Audit PDFs" folder in
 * Drive, and shows a one-click download link.
 *
 * Menu: 6S Audit Tools ▸ Download printable PDF   (wired up in onOpen.gs)
 */
function downloadPrintablePdf() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var name = (sheet.getName() + ' - 6S Audit Form').replace(/[\\/:*?"<>|]/g, '-');

  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?'
    + 'format=pdf'
    + '&gid=' + sheet.getSheetId()
    + '&portrait=true'
    + '&fitw=true'
    + '&size=letter'
    + '&gridlines=false'
    + '&printtitle=false'
    + '&sheetnames=false'
    + '&pagenumbers=true'
    + '&top_margin=0.50&bottom_margin=0.50&left_margin=0.40&right_margin=0.40'
    + '&r1=0&c1=0&r2=' + lastRow + '&c2=' + lastCol;

  var blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName(name + '.pdf');

  var folder = getOrCreateFolder_('6S Audit PDFs');
  var file = folder.createFile(blob);

  var html = HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;padding:18px;font-size:14px">'
    + '<p style="margin:0 0 12px">Your blank <b>' + sheet.getName() + '</b> audit form is ready.</p>'
    + '<p style="margin:0 0 16px"><a href="' + file.getUrl() + '" target="_blank" '
    + 'style="background:#1F2A37;color:#fff;padding:9px 14px;border-radius:6px;text-decoration:none">'
    + 'Open / download PDF</a></p>'
    + '<p style="margin:0;color:#6B7280;font-size:12px">Also saved to Drive ▸ 6S Audit PDFs</p>'
    + '</div>')
    .setWidth(390).setHeight(165);
  SpreadsheetApp.getUi().showModalDialog(html, 'Printable PDF');
}

/** Return the named Drive folder, creating it if it doesn't exist. */
function getOrCreateFolder_(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
