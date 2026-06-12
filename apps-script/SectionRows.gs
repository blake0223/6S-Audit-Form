/**
 * SectionRows — add a grading item (row) to a 6S category. Pick the facility + the
 * 6S category from a dialog; a new row is added at the end of that category and the
 * category is renumbered.
 *
 * Menu: 6S Audit Tools ▸ Monthly audit ▸ Add grading item
 */

/** Menu launcher — choose facility, choose category, then enter the item name. */
function addGradingItem() {
  var ui = SpreadsheetApp.getUi();
  var sheetName = pickFacility_('monthly', 'Add grading item');
  if (!sheetName) return;
  var section = pickOption_(['SORT', 'SET IN ORDER', 'SHINE', 'STANDARDIZE', 'SUSTAIN', 'SAFETY'],
    'Add grading item — 6S category');
  if (!section) return;
  var resp = ui.prompt('Add grading item',
    'Check item name (optional — leave blank to fill in later):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  try { ui.alert(addGradingItemFor(sheetName, section, resp.getResponseText().trim())); }
  catch (e) { ui.alert('Error: ' + e.message); }
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
