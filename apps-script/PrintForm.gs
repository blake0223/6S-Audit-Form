/**
 * PrintForm — renders the ACTIVE tab as a clean, blank, print-ready PDF
 * (portrait, fit-to-width, no gridlines), shows it in a quick preview, and opens
 * the browser's print dialog. A "Print this form" button is always available as a
 * fallback if the dialog doesn't auto-open. No download, no Drive file.
 *
 * The PDF is fetched server-side, handed to the dialog as a same-origin blob URL,
 * and loaded in a VISIBLE iframe (a hidden/display:none iframe can't be printed).
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
      '<style>'
    + 'html,body{margin:0;height:100%;font-family:Arial}'
    + '#bar{padding:8px;text-align:center;background:#1F2A37}'
    + '#bar button{font-size:14px;font-weight:bold;padding:8px 18px;border:0;'
    + 'border-radius:6px;background:#fff;color:#1F2A37;cursor:pointer}'
    + '#f{border:0;width:100%;height:calc(100% - 46px);display:block}'
    + '</style>'
    + '<div id="bar"><button onclick="doPrint()">🖨️ Print this form</button></div>'
    + '<iframe id="f"></iframe>'
    + '<script>'
    + 'var b64="' + b64 + '";'
    + 'var bytes=Uint8Array.from(atob(b64),function(c){return c.charCodeAt(0);});'
    + 'var url=URL.createObjectURL(new Blob([bytes],{type:"application/pdf"}));'
    + 'var f=document.getElementById("f");'
    + 'function doPrint(){try{f.contentWindow.focus();f.contentWindow.print();}catch(e){window.print();}}'
    + 'f.onload=function(){setTimeout(doPrint,500);};'
    + 'f.src=url;'
    + '</script>')
    .setWidth(840).setHeight(660);
  SpreadsheetApp.getUi().showModalDialog(html, 'Print — ' + sheet.getName());
}
