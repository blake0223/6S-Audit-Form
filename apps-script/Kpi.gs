/**
 * Kpi — writes a completed audit into the KPI Data tab's "Monthly Audit" block.
 * Columns are matched by their header text (Audit ID, Date Completed, Result,
 * Average Room Score, Total Room Score, and "<SECTION> Score"), so the block can
 * live anywhere on the tab and sit beside the Daily Checklist block.
 *
 * Percentages are written as fractions with 0.00% format; Total Room Score is the
 * raw points total.
 */

var KPI_SHEET_NAME = 'KPI Data';

function appendMonthlyAuditRow_(o) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(KPI_SHEET_NAME);
  if (!sh) throw new Error('KPI Data tab not found.');

  // Locate the "Audit ID" header and map every header in that row to its column.
  var scan = sh.getRange(1, 1, Math.min(sh.getLastRow() || 1, 25), sh.getLastColumn() || 1).getValues();
  var headerRow = 0, map = {};
  for (var r = 0; r < scan.length && !headerRow; r++) {
    for (var c = 0; c < scan[r].length; c++) {
      if (String(scan[r][c]).trim().toLowerCase() === 'audit id') {
        headerRow = r + 1;
        for (var cc = 0; cc < scan[r].length; cc++) {
          var h = String(scan[r][cc]).trim();
          if (h) map[h.toLowerCase()] = cc + 1;
        }
        break;
      }
    }
  }
  if (!headerRow) throw new Error('Could not find the "Audit ID" header in KPI Data.');

  // First empty row in the Audit ID column, below the header.
  var idCol = map['audit id'];
  var below = sh.getRange(headerRow + 1, idCol, Math.max(sh.getLastRow() - headerRow, 1), 1).getValues();
  var target = headerRow + 1 + below.length;
  for (var i = 0; i < below.length; i++) {
    if (String(below[i][0]).trim() === '') { target = headerRow + 1 + i; break; }
  }

  function putVal(name, v) { var col = map[name]; if (col) sh.getRange(target, col).setValue(v); }
  function putPct(name, frac) {
    var col = map[name];
    if (!col) return;
    var cell = sh.getRange(target, col);
    if (frac === '' || frac === null) { cell.setValue(''); return; }
    cell.setValue(frac).setNumberFormat('0.00%');
  }

  putVal('facility', o.facility);
  putVal('audit id', o.id);
  putVal('date completed', o.date);
  putVal('result', o.result);
  putPct('average room score', o.avg);
  putVal('total room score', o.total);
  Object.keys(o.sections).forEach(function (sec) {
    putPct(sec.toLowerCase() + ' score', o.sections[sec]);
  });
  return target;
}
