/**
 * AddRoomColumn — inserts a 0–3 scoring column for a new room, immediately to
 * the right of the questions column (C). Prompts for the room name, labels the
 * header, and drops a 0–3 dropdown on every item row (standard + custom).
 *
 * Menu: 6S Audit Tools ▸ Add room column   (wired up in onOpen.gs)
 */

var AUDIT_SHEET_NAME = '6S Audit Sheet'; // the tab these tools act on
var QUESTION_COL     = 3;                // column C = Description / Audit Question

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

  var lastRow   = sheet.getLastRow();
  var headerRow = findHeaderRow_(sheet, lastRow);
  var firstItem = headerRow ? headerRow + 1 : 8;

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
      sheet.getRange(firstItem + i, newCol)
        .setDataValidation(rule)
        .setHorizontalAlignment('center');
    }
  }

  sheet.setColumnWidth(newCol, 90);
  ui.alert('Added room column "' + roomName + '".');
}

/** Find the header row (the one with "No." in col A or "Check Item" in col B). */
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
