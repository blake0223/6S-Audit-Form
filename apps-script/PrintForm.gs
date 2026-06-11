/**
 * PrintForm — renders the ACTIVE tab as a clean, blank, print-ready PDF
 * (portrait, fit-to-width, no gridlines) and opens it in a NEW BROWSER TAB
 * (the browser's PDF viewer), where you print it with the viewer's print button
 * or Ctrl/Cmd-P. No download to Drive.
 *
 * Rows 4–5 are omitted from the printout: the PDF export takes only one contiguous
 * range, so those rows are hidden just for the export and restored immediately
 * after (Sheets leaves hidden rows out of the PDF).
 *
 * Why a new tab and not auto-print inside the dialog: Chrome renders a PDF in its
 * own (cross-origin) viewer, so printing it from inside an Apps Script dialog ends
 * up printing the Sheets page itself, not the form.
 *
 * Menu: 6S Audit Tools ▸ Print blank form   (wired up in onOpen.gs)
 */

var SKIP_ROW_START = 4; // first row to leave off the printout
var SKIP_ROW_COUNT = 2; // rows 4 and 5

function printForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var b64;

  // Hide rows 4–5 only while the PDF is generated, then put them back.
  try {
    sheet.hideRows(SKIP_ROW_START, SKIP_ROW_COUNT);
    SpreadsheetApp.flush();

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

    b64 = Utilities.base64Encode(
      UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
      }).getBlob().getBytes()
    );
  } finally {
    sheet.showRows(SKIP_ROW_START, SKIP_ROW_COUNT);
    SpreadsheetApp.flush();
  }

  var html = HtmlService.createHtmlOutput(
      '<body style="font-family:Arial;margin:0;padding:20px;text-align:center;color:#374151">'
    + '<p style="margin:0 0 14px">Your blank <b>' + escapeHtml_(sheet.getName()) + '</b> form is ready.</p>'
    + '<button id="b" style="font-size:15px;font-weight:bold;padding:10px 20px;border:0;'
    + 'border-radius:6px;background:#1F2A37;color:#fff;cursor:pointer">Open &amp; print form</button>'
    + '<p style="margin:14px 0 0;font-size:12px;color:#6B7280">Opens in a new tab — use its print button or Ctrl/⌘-P.</p>'
    + '<script>'
    + 'var bytes=Uint8Array.from(atob("' + b64 + '"),function(c){return c.charCodeAt(0);});'
    + 'var blobUrl=URL.createObjectURL(new Blob([bytes],{type:"application/pdf"}));'
    + 'document.getElementById("b").addEventListener("click",function(){'
    + 'window.open(blobUrl,"_blank");google.script.host.close();});'
    + '</script>'
    + '</body>')
    .setWidth(360).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, 'Print — ' + sheet.getName());
}

/** Minimal HTML escaping for the sheet name shown in the dialog. */
function escapeHtml_(s) {
  return String(s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}
