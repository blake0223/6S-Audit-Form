/**
 * RecordResults — submit the scores typed into the room columns to the KPI Data
 * tab's "Monthly Audit" block. Generates one Audit ID per submission and writes a
 * single row: Audit ID, Date Completed, Result, Average Room Score (%),
 * Total Room Score (raw points), and each section's score (%).
 *
 * Definitions:
 *   • Total Room Score   = sum of every room's scores (matches the row-6 total).
 *   • Average Room Score = average of the per-room score percentages.
 *   • <Section> Score    = section points ÷ section max, across all rooms.
 *   • Result             = PASS if overall ≥ 67% and no zeros, else FAIL.
 *
 * Workflow: type 0–3 into the room columns on the sheet, then run this.
 *
 * Menu: 6S Audit Tools ▸ Record audit results   (wired up in onOpen.gs)
 */
function recordResults() {
  var ui = SpreadsheetApp.getUi();
  var sheet = auditSheet_();
  var layout = getAuditLayout_();
  if (!layout.rooms.length) { ui.alert('No room columns found on the audit sheet.'); return; }
  if (!layout.items.length) { ui.alert('No audit items found.'); return; }

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

  if (!scoredCells) { ui.alert('No scores found. Type 0–3 into the room columns first, then record.'); return; }

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
  var date = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  appendMonthlyAuditRow_({ id: id, date: date, result: result, avg: avg, total: total, sections: sections });

  ui.alert('Recorded to KPI Data\n\n'
    + 'Audit ID:  ' + id + '\n'
    + 'Result:  ' + result + '\n'
    + 'Average Room Score:  ' + (Math.round(avg * 1000) / 10) + '%\n'
    + 'Total Room Score:  ' + total);
}

/** Read the audit sheet structure: rooms (name, col) and items (row, section). */
function getAuditLayout_() {
  var sheet = auditSheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headerRow = findHeaderRow_(sheet, lastRow);
  var qCol = findQuestionCol_(sheet, headerRow) || QUESTION_COL;
  var totalCol = findColByHeader_(sheet, headerRow, 'total score') || lastCol;

  var hdr = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var rooms = [];
  for (var idx = qCol; idx <= totalCol - 2; idx++) {     // columns between questions and total
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
