/**
 * SectionRows — add or remove item rows within a 6S section. The section acted on
 * is the one containing the active cell (the nearest "◆ SECTION" band above it).
 * After a change the section is renumbered (1, 2, 3 … restarting per section) and
 * the Total Score column is refreshed.
 *
 * Menu: 6S Audit Tools ▸ Add row to section / Remove selected row
 */

function addSectionRow() {
  var ui = SpreadsheetApp.getUi();
  var sheet = auditSheet_();
  var lastRow = sheet.getLastRow();
  var headerRow = findHeaderRow_(sheet, lastRow);
  var activeRow = sheet.getActiveRange().getRow();
  var b = sectionBounds_(sheet, activeRow, headerRow, lastRow);
  if (!b) { ui.alert('Select a cell inside a 6S section first.'); return; }

  // Insert after the active row when it's inside the section, else at section end.
  var afterRow = (activeRow >= b.first && activeRow <= b.last) ? activeRow : b.last;
  sheet.insertRowsAfter(afterRow, 1);
  var newRow = afterRow + 1;

  // Match the look of the row above it, then clear its contents.
  var lastCol = sheet.getLastColumn();
  sheet.getRange(afterRow, 1, 1, lastCol).copyTo(
    sheet.getRange(newRow, 1, 1, lastCol), { formatOnly: true });
  sheet.getRange(newRow, 1, 1, lastCol).clearContent();
  sheet.setRowHeight(newRow, sheet.getRowHeight(afterRow));

  renumberRange_(sheet, b.first, (b.last - b.first + 1) + 1); // +1 row added
  rebuildRoomTotals_(sheet);
}

function removeSectionRow() {
  var ui = SpreadsheetApp.getUi();
  var sheet = auditSheet_();
  var lastRow = sheet.getLastRow();
  var headerRow = findHeaderRow_(sheet, lastRow);
  var activeRow = sheet.getActiveRange().getRow();
  var b = sectionBounds_(sheet, activeRow, headerRow, lastRow);
  if (!b || activeRow < b.first || activeRow > b.last) {
    ui.alert('Select an item row inside a 6S section to remove.'); return;
  }
  if (b.last <= b.first) { ui.alert('A section must keep at least one row.'); return; }

  sheet.deleteRow(activeRow);
  renumberRange_(sheet, b.first, (b.last - b.first + 1) - 1); // -1 row removed
  rebuildRoomTotals_(sheet);
}

/* ---------- helpers ---------- */

function auditSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(AUDIT_SHEET_NAME) || ss.getActiveSheet();
}

/** True if column A of the row holds a "◆ …" section band. */
function isBand_(sheet, r) {
  return String(sheet.getRange(r, 1).getValue()).indexOf('◆') >= 0;
}

/** Bounds of the section containing `row`: { band, first, last }, or null. */
function sectionBounds_(sheet, row, headerRow, lastRow) {
  if (row <= headerRow) return null;
  var band = 0;
  for (var r = row; r > headerRow; r--) { if (isBand_(sheet, r)) { band = r; break; } }
  if (!band) return null;
  var end = lastRow;
  for (var r2 = band + 1; r2 <= lastRow; r2++) { if (isBand_(sheet, r2)) { end = r2 - 1; break; } }
  return { band: band, first: band + 1, last: end };
}

/** Write 1, 2, 3 … into column A for `count` rows starting at `first` (numeric). */
function renumberRange_(sheet, first, count) {
  if (count <= 0) return;
  var nums = [];
  for (var i = 1; i <= count; i++) nums.push([i]);
  sheet.getRange(first, 1, count, 1).setValues(nums);
}
