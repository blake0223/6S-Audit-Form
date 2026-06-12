/**
 * RoomTotals — keeps a facility audit tab's score formulas correct and dynamic as
 * rooms and rows are added. Re-stamps (matching the layout you built):
 *
 *   • Per-row Total Score (total column):
 *       =IF(<firstRoom>r="","-",SUM(<firstRoom>r:<lastRoom>r))      ← spans ALL rooms
 *   • Max Possible:
 *       =COUNTA(<checkItems>)*(COLUMN(<totalHdr>)-COLUMN(<firstRoomHdr>))*3
 *       (room count = columns between first room and the Total column, so it grows
 *        automatically when a room is inserted before the Total column)
 *   • Each room % and the Total %:
 *       =IFERROR(SUM(<col items>)/$<MaxPossible>$,0)   ← ABSOLUTE max ref, can't shift
 *
 * Total Score and RESULT reference the Total column directly, so they follow it on
 * insert and grow when rows are added inside — left as-is.
 *
 * Runs automatically on Add room / Add row, and on demand from the menu.
 * Menu: 6S Audit Tools ▸ Rebuild room totals   (wired up in onOpen.gs)
 */
/**
 * Re-stamp the per-row totals and summary formulas. Returns item-row count.
 * Called automatically by Add location / Add grading item (no menu item).
 */
function rebuildRoomTotals_(sheet) {
  var lastRow = sheet.getLastRow();
  var headerRow = findHeaderRow_(sheet, lastRow);
  var qCol = findQuestionCol_(sheet, headerRow) || QUESTION_COL;
  var totalCol = findColByHeader_(sheet, headerRow, 'total score');
  if (!totalCol || totalCol <= qCol + 1) return 0;

  var firstRoom = qCol + 1, lastRoom = totalCol - 1;
  var R1 = columnToLetter_(firstRoom), Rn = columnToLetter_(lastRoom);

  var colA = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, 1).getValues();
  var rows = [];
  for (var i = 0; i < colA.length; i++) if (isItemNumber_(colA[i][0])) rows.push(headerRow + 1 + i);
  if (!rows.length) return 0;
  var first = rows[0], last = rows[rows.length - 1];

  // Per-row Total Score — spans all room columns (matches your IF(<firstRoom>="") pattern).
  rows.forEach(function (r) {
    sheet.getRange(r, totalCol).setFormula('=IF(' + R1 + r + '="","-",SUM(' + R1 + r + ':' + Rn + r + '))');
  });

  refreshFacilitySummary_(sheet, headerRow, qCol, totalCol, firstRoom, lastRoom, first, last);
  return rows.length;
}

/** Re-stamp Max Possible (dynamic room count) and the room/total % cells (absolute max ref). */
function refreshFacilitySummary_(sheet, headerRow, qCol, totalCol, firstRoom, lastRoom, first, last) {
  var pctRow = headerRow - 1;
  if (pctRow < 1) return;

  var max = findLabelValueCell_(sheet, headerRow, 'max possible');
  if (!max) return;

  var checkL = columnToLetter_(Math.max(qCol - 1, 1));         // Check Item column (counts items)
  var totalHdr = columnToLetter_(totalCol) + headerRow;        // e.g. F6
  var firstHdr = columnToLetter_(firstRoom) + headerRow;       // e.g. D6
  sheet.getRange(max.row, max.col).setFormula(
    '=COUNTA(' + checkL + first + ':' + checkL + last + ')*(COLUMN(' + totalHdr + ')-COLUMN(' + firstHdr + '))*3');

  var maxRef = '$' + columnToLetter_(max.col) + '$' + max.row;
  for (var c = firstRoom; c <= lastRoom; c++) {
    var cl = columnToLetter_(c);
    sheet.getRange(pctRow, c).setFormula('=IFERROR(SUM(' + cl + first + ':' + cl + last + ')/' + maxRef + ',0)');
  }
  var tl = columnToLetter_(totalCol);
  sheet.getRange(pctRow, totalCol).setFormula('=IFERROR(SUM(' + tl + first + ':' + tl + last + ')/' + maxRef + ',0)');
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

/** Find a label above the header row and return its VALUE cell (the row below it). */
function findLabelValueCell_(sheet, headerRow, needle) {
  needle = needle.toLowerCase();
  var rows = Math.max(headerRow - 1, 1);
  var vals = sheet.getRange(1, 1, rows, sheet.getLastColumn()).getValues();
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      if (String(vals[r][c]).trim().toLowerCase().indexOf(needle) >= 0) {
        return { row: Math.min(r + 2, headerRow - 1), col: c + 1 };
      }
    }
  }
  return null;
}

/** 1 -> A, 4 -> D, 27 -> AA. */
function columnToLetter_(col) {
  var s = '';
  while (col > 0) { var m = (col - 1) % 26; s = String.fromCharCode(65 + m) + s; col = (col - m - 1) / 26; }
  return s;
}
