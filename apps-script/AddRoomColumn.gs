/**
 * AddRoomColumn — "Add location": insert a scoring column for a new location/area
 * on a chosen facility's Monthly Audit tab. You pick the facility from a dropdown
 * (no need to be on the tab). The column is inserted just left of the Total Score
 * column so the "Audit Location Scores" band extends over it; formatting is copied
 * from the neighbouring location column cell-by-cell (never as a full-column range,
 * which would cross the title/legend/band merges and throw). rebuildRoomTotals_
 * then stamps the Score % and per-row totals.
 *
 * Menu: 6S Audit Tools ▸ Monthly audit ▸ Add location
 */

var QUESTION_COL = 3; // fallback if the question header isn't found

/** Menu launcher — choose facility, then enter the location name. */
function addLocation() {
  var ui = SpreadsheetApp.getUi();
  var sheetName = pickFacility_('monthly', 'Add location');
  if (!sheetName) return;
  var resp = ui.prompt('Add location', 'Location / area name (e.g. Parts Room):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var name = resp.getResponseText().trim();
  if (!name) { ui.alert('No location name entered.'); return; }
  try { ui.alert(addLocationFor(sheetName, name)); }
  catch (e) { ui.alert('Error: ' + e.message); }
}

/** Core — add a location column to the named Monthly Audit tab. */
function addLocationFor(sheetName, name) {
  name = String(name || '').trim();
  if (!name) throw new Error('Enter a location / area name.');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Tab not found: ' + sheetName);

  var lastRow = sheet.getLastRow();
  var headerRow = findHeaderRow_(sheet, lastRow);
  var qCol = findQuestionCol_(sheet, headerRow) || QUESTION_COL;
  var totalCol = findColByHeader_(sheet, headerRow, 'total score');
  if (!totalCol || totalCol <= qCol + 1) {
    throw new Error('Could not find the location block / "Total Score" column on ' + sheetName + '.');
  }

  sheet.insertColumnsBefore(totalCol, 1);
  var newCol = totalCol, srcCol = newCol - 1;
  sheet.setColumnWidth(newCol, sheet.getColumnWidth(srcCol));

  // Copy formatting from the neighbour column on non-merged cells only.
  var pctRow = headerRow - 1;
  if (pctRow >= 1) sheet.getRange(pctRow, srcCol).copyTo(sheet.getRange(pctRow, newCol), { formatOnly: true });
  sheet.getRange(headerRow, srcCol).copyTo(sheet.getRange(headerRow, newCol), { formatOnly: true });
  var colA = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, 1).getValues();
  for (var i = 0; i < colA.length; i++) {
    if (isItemNumber_(colA[i][0])) {
      var r = headerRow + 1 + i;
      sheet.getRange(r, srcCol).copyTo(sheet.getRange(r, newCol), { formatOnly: true });
    }
  }

  var color = nextRoomColor_();
  sheet.getRange(headerRow, newCol)
    .setValue(name).setFontWeight('bold')
    .setFontColor(color.font).setBackground(color.bg)
    .setHorizontalAlignment('center').setWrap(true);

  rebuildRoomTotals_(sheet);
  return 'Added location "' + name + '" to ' + facilityLabel_(sheetName) + '.';
}

/* ---------- shared helpers ---------- */

/** Header row = the one with "No." in col A or "Check Item" in col B. */
function findHeaderRow_(sheet, lastRow) {
  var max = Math.min(lastRow, 30);
  var vals = sheet.getRange(1, 1, max, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    var a = String(vals[i][0]).trim().toLowerCase();
    var b = String(vals[i][1]).trim().toLowerCase();
    if (a === 'no.' || b === 'check item') return i + 1;
  }
  return 7;
}

/** Find the question column by its header text ("description" / "audit question"). */
function findQuestionCol_(sheet, headerRow) {
  var lastCol = sheet.getLastColumn();
  var hdr = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  for (var c = 0; c < hdr.length; c++) {
    var t = String(hdr[c]).toLowerCase();
    if (t.indexOf('description') >= 0 || t.indexOf('audit question') >= 0) return c + 1;
  }
  return 0;
}

/** True if a column-A value is an item number — numeric (1) or "1.0"-style text. */
function isItemNumber_(v) {
  if (typeof v === 'number') return v > 0;
  var s = String(v).trim();
  return /^\d+(\.\d+)?$/.test(s) && parseFloat(s) > 0;
}

/** Next location header color, rotating through a palette; position remembered per doc. */
function nextRoomColor_() {
  var palette = [
    { bg: '#1F4E79', font: '#FFFFFF' }, { bg: '#2E7D32', font: '#FFFFFF' },
    { bg: '#8E44AD', font: '#FFFFFF' }, { bg: '#C0392B', font: '#FFFFFF' },
    { bg: '#D68910', font: '#FFFFFF' }, { bg: '#16A085', font: '#FFFFFF' },
    { bg: '#AD1457', font: '#FFFFFF' }, { bg: '#2C3E50', font: '#FFFFFF' }
  ];
  var props = PropertiesService.getDocumentProperties();
  var i = parseInt(props.getProperty('roomColorIndex') || '0', 10);
  props.setProperty('roomColorIndex', String(i + 1));
  return palette[i % palette.length];
}
