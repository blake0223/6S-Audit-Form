/**
 * RecordResults — pick a facility, then submit the scores typed into that
 * facility's Monthly Audit room columns to the KPI Data tab's "Monthly Audit"
 * block. Generates one Audit ID per submission and writes a single row:
 * Facility, Audit ID, Date Completed, Result, Average Room Score (%),
 * Total Room Score (raw), and each section's score (%).
 *
 * Menu: 6S Audit Tools ▸ Record audit results
 */

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

  // Require a complete audit — every item must be scored in every location.
  // Partial sheets can't be submitted.
  var expected = layout.rooms.length * layout.items.length;
  if (scoredCells < expected) {
    var blanks = expected - scoredCells;
    throw new Error('This monthly audit isn’t complete, so it can’t be submitted yet. '
      + 'Every item needs a 0–3 score in every location first — '
      + blanks + ' of ' + expected + ' score cells are still blank. '
      + 'Fill them in, then run Record Audit Results again.');
  }

  var roomPcts = [];
  layout.rooms.forEach(function (rm) {
    var a = roomAgg[rm.col];
    if (a.scored > 0) roomPcts.push(a.sum / (a.scored * 3));
  });
  var avg = roomPcts.reduce(function (s, x) { return s + x; }, 0) / roomPcts.length;
  var overall = total / (scoredCells * 3);
  // Use the sheet's own RESULT (the PASS/FAIL cell shown above the header) so the
  // recorded result matches what the form displays; fall back to a computed result.
  var result = readSheetResult_(sheet, layout.headerRow)
    || ((overall >= 0.67 && zeros === 0) ? 'PASS' : 'FAIL');

  var sections = {};
  Object.keys(secAgg).forEach(function (sec) {
    var a = secAgg[sec];
    sections[sec] = a.scored > 0 ? a.sum / (a.scored * 3) : '';
  });

  // Capture each item's comment (Comments column is just right of Total Score) into
  // one string, labeled by check-item name, to store on the KPI Data row.
  var commentCol = layout.totalCol ? layout.totalCol + 1 : 0;
  var comments = '';
  if (commentCol && commentCol <= sheet.getLastColumn()) {
    var nameCol = Math.max((layout.qCol || QUESTION_COL) - 1, 1);
    var cVals = sheet.getRange(firstRow, commentCol, span, 1).getValues();
    var nVals = sheet.getRange(firstRow, nameCol, span, 1).getValues();
    var parts = [];
    layout.items.forEach(function (it) {
      var k = it.row - firstRow;
      var cmt = String(cVals[k][0] || '').trim();
      if (cmt) {
        var nm = String(nVals[k][0] || '').trim();
        parts.push((nm ? nm + ': ' : '') + cmt);
      }
    });
    comments = parts.join('  |  ');
  }

  var id = newAuditId_();
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  appendMonthlyAuditRow_({
    facility: facilityLabel_(sheetName), id: id,
    date: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    result: result, avg: avg, total: total, sections: sections, comments: comments
  });

  // Clear the typed room scores AND the comments so the form is fresh for the next
  // audit. Never touch the Total Score column — its per-row formula must survive.
  var clearCols = (layout.totalCol && layout.totalCol > firstRoomCol)
    ? (layout.totalCol - firstRoomCol)
    : nRoomCols;
  sheet.getRange(firstRow, firstRoomCol, span, clearCols).clearContent();
  if (commentCol && commentCol <= sheet.getLastColumn()) {
    sheet.getRange(firstRow, commentCol, span, 1).clearContent();
  }

  // Re-stamp the per-row Total Score formula so it is always retained after clearing.
  rebuildRoomTotals_(sheet);

  return 'Recorded ' + facilityLabel_(sheetName) + ' — ' + result
    + ' (' + (Math.round(avg * 1000) / 10) + '%). Audit ID ' + id
    + '. Score entries cleared for the next audit.';
}

/**
 * Read the sheet's own RESULT (PASS/FAIL) shown in the summary area above the header
 * row, so the recorded result matches what the audit form displays. Returns
 * 'PASS' / 'FAIL', or '' if no result cell is found.
 */
function readSheetResult_(sheet, headerRow) {
  if (!headerRow || headerRow < 2) return '';
  var vals = sheet.getRange(1, 1, headerRow - 1, sheet.getLastColumn()).getValues();
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      var s = String(vals[r][c]).toUpperCase();
      if (s.indexOf('PASS') >= 0) return 'PASS';
      if (s.indexOf('FAIL') >= 0) return 'FAIL';
    }
  }
  return '';
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
  return { rooms: rooms, items: items, headerRow: headerRow, totalCol: totalCol, qCol: qCol };
}
