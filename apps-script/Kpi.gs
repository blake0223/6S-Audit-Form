/**
 * Kpi — the KPI Data tab: results land here, one summary row per room per audit.
 * Columns are created on first use if the tab's header row is empty.
 */

var KPI_SHEET_NAME = 'KPI Data';
var KPI_HEADERS = ['Audit ID', 'Date', 'Location', 'Room',
                   'Total Score', 'Max Possible', 'Score %', '# of Zeros', 'Result'];

/** Return the KPI Data sheet, creating it (and its header row) if needed. */
function kpiSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(KPI_SHEET_NAME) || ss.insertSheet(KPI_SHEET_NAME);
  if (String(sh.getRange(1, 1).getValue()).trim() === '') {
    sh.getRange(1, 1, 1, KPI_HEADERS.length).setValues([KPI_HEADERS])
      .setFontWeight('bold').setBackground('#1F2A37').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Append one result row. */
function appendKpiRow_(o) {
  kpiSheet_().appendRow([o.id, o.date, o.location, o.room,
                         o.total, o.max, o.pct, o.zeros, o.result]);
}
