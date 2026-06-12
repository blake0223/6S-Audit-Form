/**
 * Checklist — add a facility-specific item (column) to a Weekly / Daily 6S
 * Checklist tab, with a checkbox for each week row. Run it while viewing the
 * checklist tab you want to add to.
 *
 * Menu: 6S Audit Tools ▸ Weekly checklist ▸ Add checklist item
 */
function addChecklistItem() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  // Header row = the one whose first cell mentions "checklist".
  var headerRow = 0;
  var scan = sheet.getRange(1, 1, Math.min(lastRow, 15), 1).getValues();
  for (var r = 0; r < scan.length; r++) {
    if (String(scan[r][0]).toLowerCase().indexOf('checklist') >= 0) { headerRow = r + 1; break; }
  }
  if (!headerRow) { ui.alert('Open a facility Weekly Checklist tab first, then run this.'); return; }

  var resp = ui.prompt('Add checklist item',
    'New weekly checklist item (e.g. "Propane turned off"):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var name = resp.getResponseText().trim();
  if (!name) { ui.alert('No item entered — nothing added.'); return; }

  // Last filled task column in the header row → new column right after it.
  var hdr = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var lastTaskCol = 1;
  for (var c = 0; c < hdr.length; c++) if (String(hdr[c]).trim() !== '') lastTaskCol = c + 1;
  var newCol = lastTaskCol + 1;

  // Header — match the look of the previous task header.
  sheet.getRange(headerRow, lastTaskCol).copyTo(sheet.getRange(headerRow, newCol), { formatOnly: true });
  sheet.getRange(headerRow, newCol).setValue(name);
  sheet.setColumnWidth(newCol, sheet.getColumnWidth(lastTaskCol));

  // Checkbox (unchecked) for each week row — rows below the header with a label in col A.
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  for (var rr = headerRow + 1; rr <= lastRow; rr++) {
    if (String(sheet.getRange(rr, 1).getValue()).trim() === '') continue;
    sheet.getRange(rr, newCol).setDataValidation(rule).setValue(false);
  }

  ui.alert('Added checklist item "' + name + '".');
}
