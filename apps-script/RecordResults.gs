/**
 * RecordResults — pick a facility, then submit the scores typed into that
 * facility's Monthly Audit room columns to the KPI Data tab's "Monthly Audit"
 * block. Generates one Audit ID per submission and writes a single row:
 * Facility, Audit ID, Date Completed, Result, Average Room Score (%),
 * Total Room Score (raw), and each section's score (%).
 *
 * Menu: 6S Audit Tools ▸ Record audit results
 */

/** Menu launcher — opens the facility picker. */
function recordAudit() {
  ensureAuthorized_();
  showToolDialog_({
    title: 'Record audit results',
    type: 'monthly',
    button: 'Record',
    fields: [],
    callback: 'recordResultsFor'
  });
}

/** Core — record the named facility's Monthly Audit into KPI Data. */
function recordResultsFor(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Tab not found: ' + sheetName);
  var layout = getAuditLayout_(sheet);
  if (!layout.rooms.length) throw new Error('No location columns found on ' + sheetName + '.');
  if (!layout.items.length) throw new Error('No audit items found on ' + sheetName + '.');

  var firstRow = layout.items[0].row;
  var lastItemRow = layout.items[layout.items.length - 1].row;
  var span = lastItemRow - firstRow + 1;
  var firstRoomCol = layout.rooms[0].col;
  var nRoomCols = layout.rooms[layout.rooms.length - 1].col - firstRoomCol + 1;
  var block = sheet.getRange(firstRow, firstRoomCol, span, nRoomCols).getValues();

  var total = 0, scoredCells = 0, zeros = 0;
  var roomAgg = {}, secAgg = {};
  layout.rooms.forEach(function (rm) { roomAgg[rm.col] = { sum: 0, scored: 0 }; });

  layout.items.forEach(function (it) {
    if (!secAgg[it.section]) secAgg[it.section] = { sum: 0, scored: 0 };
    layout.rooms.forEach(function (rm) {
      var v = block[it.row - firstRow][rm.col - firstRoomCol];
      if (v === '' || v === null) return;
      v = Number(v);
      if (isNaN(v)) return;
      total += v; scoredCells++; if (v === 0) zeros++;
      roomAgg[rm.col].sum += v; roomAgg[rm.col].scored++;
      secAgg[it.section].sum += v; secAgg[it.section].scored++;
    });
  });

  if (!scoredCells) throw new Error('No scores found on ' + sheetName + ' — type 0–3 into the location columns first.');

  var roomPcts = [];
  layout.rooms.forEach(function (rm) {
    var a = roomAgg[rm.col];
    if (a.scored > 0) roomPcts.push(a.sum / (a.scored * 3));
  });
  var avg = roomPcts.reduce(function (s, x) { return s + x; }, 0) / roomPcts.length;
  var overall = total / (scoredCells * 3);
  var result = (overall >= 0.67 && zeros === 0) ? 'PASS' : 'FAIL';

  var sections = {};
  Object.keys(secAgg).forEach(function (sec) {
    var a = secAgg[sec];
    sections[sec] = a.scored > 0 ? a.sum / (a.scored * 3) : '';
  });

  var id = newAuditId_();
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  appendMonthlyAuditRow_({
    facility: facilityLabel_(sheetName), id: id,
    date: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    result: result, avg: avg, total: total, sections: sections
  });

  return 'Recorded ' + facilityLabel_(sheetName) + ' — ' + result
    + ' (' + (Math.round(avg * 1000) / 10) + '%). Audit ID ' + id + '.';
}

/** Read the audit sheet structure: rooms (name, col) and items (row, section). */
function getAuditLayout_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headerRow = findHeaderRow_(sheet, lastRow);
  var qCol = findQuestionCol_(sheet, headerRow) || QUESTION_COL;
  var totalCol = findColByHeader_(sheet, headerRow, 'total score') || lastCol;

  var hdr = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var rooms = [];
  for (var idx = qCol; idx <= totalCol - 2; idx++) {
    var name = String(hdr[idx]).trim();
    if (name) rooms.push({ name: name, col: idx + 1 });
  }

  var data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, 3).getValues();
  var items = [], section = '';
  for (var i = 0; i < data.length; i++) {
    var a = data[i][0];
    if (String(a).indexOf('◆') >= 0) { section = String(a).replace(/◆/g, '').trim(); continue; }
    if (isItemNumber_(a)) items.push({ row: headerRow + 1 + i, section: section });
  }
  return { rooms: rooms, items: items };
}
