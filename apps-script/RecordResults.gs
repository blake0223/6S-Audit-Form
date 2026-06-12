/**
 * RecordResults — submit the scores currently typed into the room columns to the
 * KPI Data tab. Generates a single Audit ID for the submission and appends one
 * summary row per room that has scores: Audit ID, Date, Location, Room, Total,
 * Max, Score %, # of Zeros, Result.
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

  var id = newAuditId_();
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var date = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var submitted = [];
  for (var k = 0; k < layout.rooms.length; k++) {
    var room = layout.rooms[k];
    var colVals = sheet.getRange(firstRow, room.col, span, 1).getValues();
    var total = 0, zeros = 0, scored = 0;
    for (var i = 0; i < layout.items.length; i++) {
      var v = colVals[layout.items[i].row - firstRow][0];
      if (v === '' || v === null) continue;
      v = Number(v);
      if (isNaN(v)) continue;
      total += v; scored++;
      if (v === 0) zeros++;
    }
    if (!scored) continue; // skip rooms with nothing entered
    var max = scored * 3;
    var pct = Math.round(total / max * 1000) / 10;
    var result = (pct >= 67 && zeros === 0) ? 'PASS' : 'FAIL';
    appendKpiRow_({
      id: id, date: date, location: layout.location, room: room.name,
      total: total, max: max, pct: pct, zeros: zeros, result: result
    });
    submitted.push(room.name + ': ' + total + '/' + max + ' (' + pct + '%) ' + result);
  }

  if (!submitted.length) {
    ui.alert('No scores found. Type 0–3 into the room columns first, then record.');
    return;
  }
  ui.alert('Recorded to KPI Data\n\nAudit ID:  ' + id + '\n\n' + submitted.join('\n'));
}

/** Read the audit sheet structure: rooms, items (by section), max, location. */
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
    if (typeof a === 'number' && a > 0) items.push({ row: headerRow + 1 + i, section: section });
  }
  return { rooms: rooms, items: items, max: items.length * 3, location: readLocation_() };
}

/** Best-effort read of the Location value from the top of the audit sheet. */
function readLocation_() {
  var sheet = auditSheet_();
  var rows = Math.max(Math.min(sheet.getLastRow(), 12), 3);
  var cols = Math.max(Math.min(sheet.getLastColumn(), 12), 8);
  var vals = sheet.getRange(1, 1, rows, cols).getValues();
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      if (/^location/i.test(String(vals[r][c]).trim()) && c + 1 < cols) {
        var right = String(vals[r][c + 1]).trim();
        if (right) return right;
      }
    }
  }
  return '';
}
