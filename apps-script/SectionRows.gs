/**
 * SectionRows — add a grading item (row) to a 6S category. Pick the facility + the
 * 6S category from a dialog; a new row is added at the end of that category and the
 * category is renumbered.
 *
 * Menu: 6S Audit Tools ▸ Monthly audit ▸ Add grading item
 */

/** Menu launcher — Add grading item. */
function addGradingItem() {
  showToolDialog_({
    title: 'Add grading item',
    type: 'monthly',
    button: 'Add item',
    fields: [
      { id: 'section', label: '6S category', kind: 'select',
        options: ['SORT', 'SET IN ORDER', 'SHINE', 'STANDARDIZE', 'SUSTAIN', 'SAFETY'] },
      { id: 'name', label: 'Check item (optional)', kind: 'text', placeholder: 'e.g. Refrigerant cylinders secured' }
    ],
    callback: 'addGradingItemFor'
  });
}

/** Core — add a grading item to the named tab's chosen category. */
function addGradingItemFor(sheetName, section, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Tab not found: ' + sheetName);
  var want = String(section || '').trim().toLowerCase();
  if (!want) throw new Error('Pick a 6S category.');

  var lastRow = sheet.getLastRow();
  var headerRow = findHeaderRow_(sheet, lastRow);

  var band = 0;
  for (var r = headerRow + 1; r <= lastRow; r++) {
    var a = String(sheet.getRange(r, 1).getValue());
    if (a.indexOf('◆') >= 0 && a.replace(/◆/g, '').trim().toLowerCase() === want) { band = r; break; }
  }
  if (!band) throw new Error('Category "' + section + '" not found on ' + sheetName + '.');

  var end = lastRow;
  for (var r2 = band + 1; r2 <= lastRow; r2++) {
    if (String(sheet.getRange(r2, 1).getValue()).indexOf('◆') >= 0) { end = r2 - 1; break; }
  }

  sheet.insertRowsAfter(end, 1);
  var newRow = end + 1, lastCol = sheet.getLastColumn();
  sheet.getRange(end, 1, 1, lastCol).copyTo(sheet.getRange(newRow, 1, 1, lastCol), { formatOnly: true });
  sheet.getRange(newRow, 1, 1, lastCol).clearContent();
  sheet.setRowHeight(newRow, sheet.getRowHeight(end));

  if (String(name || '').trim()) {
    var checkCol = Math.max((findQuestionCol_(sheet, headerRow) || QUESTION_COL) - 1, 1);
    sheet.getRange(newRow, checkCol).setValue(String(name).trim());
  }

  renumberRange_(sheet, band + 1, (end - band) + 1); // old item count + 1
  rebuildRoomTotals_(sheet);
  return 'Added a grading item to ' + want.toUpperCase() + ' on ' + facilityLabel_(sheetName) + '.';
}

/* ---------- helpers ---------- */

/** Write 1, 2, 3 … into column A for `count` rows starting at `first`. */
function renumberRange_(sheet, first, count) {
  if (count <= 0) return;
  var nums = [];
  for (var i = 1; i <= count; i++) nums.push([i]);
  sheet.getRange(first, 1, count, 1).setValues(nums);
}
