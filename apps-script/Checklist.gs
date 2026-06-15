/**
 * Checklist — add a facility-specific item (column) to a Weekly Checklist tab,
 * with a checkbox for each week row. Pick the facility from the dialog.
 *
 * Menu: 6S Audit Tools ▸ Weekly checklist ▸ Add checklist item
 */

/** Menu launcher — choose facility, then enter the checklist item. */
function addChecklistItem() {
  var ui = SpreadsheetApp.getUi();
  var sheetName = pickFacility_('checklist', 'Add weekly checklist item');
  if (!sheetName) return;
  var resp = ui.prompt('Add weekly checklist item',
    'Checklist item (e.g. Propane turned off):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var name = resp.getResponseText().trim();
  if (!name) { ui.alert('No item entered.'); return; }
  try { ui.alert(addChecklistItemFor(sheetName, name)); }
  catch (e) { ui.alert('Error: ' + e.message); }
}

/** Core — add a checklist item column to the named Weekly Checklist tab. */
function addChecklistItemFor(sheetName, name) {
  name = String(name || '').trim();
  if (!name) throw new Error('Enter a checklist item.');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Tab not found: ' + sheetName);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  var headerRow = 0;
  var scan = sheet.getRange(1, 1, Math.min(lastRow, 15), 1).getValues();
  for (var r = 0; r < scan.length; r++) {
    if (String(scan[r][0]).toLowerCase().indexOf('checklist') >= 0) { headerRow = r + 1; break; }
  }
  if (!headerRow) headerRow = 1; // fall back to the top row if no "Checklist" title cell

  var hdr = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var lastTaskCol = 1;
  for (var c = 0; c < hdr.length; c++) if (String(hdr[c]).trim() !== '') lastTaskCol = c + 1;
  var newCol = lastTaskCol + 1;
  if (sheet.getMaxColumns() < newCol) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), newCol - sheet.getMaxColumns());
  }

  sheet.getRange(headerRow, lastTaskCol).copyTo(sheet.getRange(headerRow, newCol), { formatOnly: true });
  sheet.getRange(headerRow, newCol).setValue(name);
  sheet.setColumnWidth(newCol, sheet.getColumnWidth(lastTaskCol));

  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  var weeks = 0;
  for (var rr = headerRow + 1; rr <= lastRow; rr++) {
    if (String(sheet.getRange(rr, 1).getValue()).trim() === '') continue;
    sheet.getRange(rr, newCol).setDataValidation(rule).setValue(false);
    weeks++;
  }

  return 'Added checklist item "' + name + '" to ' + facilityLabel_(sheetName)
    + ' with a checkbox on ' + weeks + ' week row(s).';
}
