/**
 * PrintForm — renders the ACTIVE tab as a clean, blank, print-ready PDF
 * (portrait, fit-to-width, no gridlines) and opens the browser's print dialog.
 * It does NOT auto-print — you still choose the printer and confirm. No
 * download, no Drive file.
 *
 * How it works: the PDF is fetched server-side, handed to a tiny dialog as a
 * same-origin blob URL, loaded in a hidden iframe, and printed on load.
 *
 * Menu: 6S Audit Tools ▸ Print blank form   (wired up in onOpen.gs)
 */
function printForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

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

  var pdf = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob();
  var b64 = Utilities.base64Encode(pdf.getBytes());

  var html = HtmlService.createHtmlOutput(
      '<body style="font-family:Arial;margin:0;padding:16px;font-size:14px;color:#374151">'
    + 'Opening the print dialog for <b>' + escapeHtml_(sheet.getName()) + '</b>…'
    + '<iframe id="pf" style="display:none"></iframe>'
    + '<script>'
    + 'var b64="' + b64 + '";'
    + 'var bytes=Uint8Array.from(atob(b64),function(c){return c.charCodeAt(0);});'
    + 'var blobUrl=URL.createObjectURL(new Blob([bytes],{type:"application/pdf"}));'
    + 'var f=document.getElementById("pf");'
    + 'f.onload=function(){setTimeout(function(){f.contentWindow.focus();f.contentWindow.print();},250);};'
    + 'f.src=blobUrl;'
    + '</script>'
    + '</body>')
    .setWidth(320).setHeight(110);
  SpreadsheetApp.getUi().showModalDialog(html, 'Print');
}

/** Minimal HTML escaping for the sheet name shown in the dialog. */
function escapeHtml_(s) {
  return String(s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}
