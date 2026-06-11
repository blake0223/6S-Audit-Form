/**
 * 6S Audit Tools — Google Sheets menu
 * Menu items:
 *   • Add room column      → inserts a scoring column for a new room, right of the questions
 *   • Download printable PDF → exports the active tab as a clean, blank, fillable PDF
 *
 * Install: Extensions ▸ Apps Script → paste this in Code.gs → Save → reload the sheet.
 * First run will ask you to authorize (Sheets, Drive, external request).
 */

var AUDIT_SHEET_NAME = '6S Audit Sheet'; // the tab these tools act on
var QUESTION_COL     = 3;                // column C = Description / Audit Question

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('6S Audit Tools')
    .addItem('➕  Add room column', 'addRoomColumn')
    .addSeparator()
    .addItem('🖨️  Download printable PDF', 'downloadPrintablePdf')
    .addToUi();
}

/** Insert a 0–3 scoring column for a new room, immediately to the right of the questions (col C). */
function addRoomColumn() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AUDIT_SHEET_NAME) || ss.getActiveSheet();

  var resp = ui.prompt('Add room', 'Name of the room / area to add:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var roomName = resp.getResponseText().trim();
  if (!roomName) { ui.alert('No room name entered — nothing added.'); return; }

  // New blank column right of the questions column.
  sheet.insertColumnsAfter(QUESTION_COL, 1);
  var newCol = QUESTION_COL + 1; // D

  var lastRow    = sheet.getLastRow();
  var headerRow  = findHeaderRow_(sheet, lastRow);
  var firstItem  = headerRow ? headerRow + 1 : 8;

  // Header = room name.
  if (headerRow) {
    sheet.getRange(headerRow, newCol)
      .setValue(roomName)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setWrap(true);
  }

  // 0–3 dropdown on every item row (rows whose column A holds an item number).
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([0, 1, 2, 3], true)
    .setAllowInvalid(false)
    .setHelpText('0 = Major Issue · 1 = Needs Work · 2 = Good · 3 = Excellent')
    .build();

  var colA = sheet.getRange(firstItem, 1, lastRow - firstItem + 1, 1).getValues();
  for (var i = 0; i < colA.length; i++) {
    if (typeof colA[i][0] === 'number' && colA[i][0] > 0) {
      var r = firstItem + i;
      sheet.getRange(r, newCol).setDataValidation(rule).setHorizontalAlignment('center');
    }
  }

  sheet.setColumnWidth(newCol, 90);
  ui.alert('Added room column "' + roomName + '".');
}

/** Export the ACTIVE tab as a clean, blank, printable PDF and give a download link. */
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

/* ---------- helpers ---------- */

function findHeaderRow_(sheet, lastRow) {
  var max = Math.min(lastRow, 30);
  var vals = sheet.getRange(1, 1, max, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    var a = String(vals[i][0]).trim().toLowerCase();
    var b = String(vals[i][1]).trim().toLowerCase();
    if (a === 'no.' || b === 'check item') return i + 1;
  }
  return 7; // default for this template
}

function getOrCreateFolder_(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
