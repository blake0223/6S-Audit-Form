/**
 * RoomTotals — fills the "Total Score" column so that, for every item row, it
 * sums that row's room-score columns (all columns between the questions column
 * and the Total Score column). Runs automatically when a room is added, and can
 * be run on demand from the menu.
 *
 * Menu: 6S Audit Tools ▸ Rebuild room totals   (wired up in onOpen.gs)
 */
function rebuildRoomTotals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AUDIT_SHEET_NAME) || ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var n = rebuildRoomTotals_(sheet);
  if (n > 0) {
    ui.alert('Total Score column updated — ' + n + ' rows now sum their room scores.');
  } else {
    ui.alert('No totals written. Make sure there are room columns between the '
      + 'questions and a "Total Score" column.');
  }
}

/**
 * Write per-row SUM formulas into the Total Score column. Returns the number of
 * rows updated (0 if the layout couldn't be resolved).
 */
function rebuildRoomTotals_(sheet) {
  var lastRow = sheet.getLastRow();
  var headerRow = findHeaderRow_(sheet, lastRow);
  var questionCol = findQuestionCol_(sheet, headerRow) || QUESTION_COL;
  var totalCol = findColByHeader_(sheet, headerRow, 'total score');

  // Need at least one room column between the questions and the Total Score column.
  if (!totalCol || totalCol <= questionCol + 1) return 0;

  var firstRoom = questionCol + 1;
  var lastRoom = totalCol - 1;
  var L1 = columnToLetter_(firstRoom);
  var L2 = columnToLetter_(lastRoom);

  var colA = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, 1).getValues();
  var count = 0;
  for (var i = 0; i < colA.length; i++) {
    if (typeof colA[i][0] === 'number' && colA[i][0] > 0) {
      var r = headerRow + 1 + i;
      sheet.getRange(r, totalCol).setFormula('=SUM(' + L1 + r + ':' + L2 + r + ')');
      count++;
    }
  }
  return count;
}

/* ---------- helpers ---------- */

/** First column (in the header row) whose text contains the given needle. */
function findColByHeader_(sheet, headerRow, needle) {
  var lastCol = sheet.getLastColumn();
  var hdr = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  needle = needle.toLowerCase();
  for (var c = 0; c < hdr.length; c++) {
    if (String(hdr[c]).toLowerCase().indexOf(needle) >= 0) return c + 1;
  }
  return 0;
}

/** 1 -> A, 4 -> D, 27 -> AA. */
function columnToLetter_(col) {
  var s = '';
  while (col > 0) { var m = (col - 1) % 26; s = String.fromCharCode(65 + m) + s; col = (col - m - 1) / 26; }
  return s;
}
