/**
 * ViewSubmissions — read-only viewer for recorded audits. Reads the KPI Data tab's
 * "Monthly Audit" block and shows every submission in a filterable table (newest
 * first). Runs in the menu context and embeds the data straight into the dialog
 * HTML, so there's no google.script.run call and no authorization prompt.
 *
 * Menu: 6S Audit Tools ▸ View past submissions
 */
function viewSubmissions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sh = ss.getSheetByName('KPI Data');
  if (!sh) { ui.alert('KPI Data tab not found.'); return; }

  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  var scan = sh.getRange(1, 1, Math.min(lastRow || 1, 25), lastCol || 1).getValues();
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
  if (!headerRow) { ui.alert('No recorded audits found (no "Audit ID" header in KPI Data).'); return; }

  var idCol = map['audit id'];
  var brandCol = (idCol > 1 && String(sh.getRange(headerRow, idCol - 1).getValue()).trim().toLowerCase() === 'brand')
    ? idCol - 1 : 0;
  var tz = ss.getSpreadsheetTimeZone();

  var defs = [
    { label: 'Facility', col: brandCol },
    { label: 'Date', col: map['date completed'], date: true },
    { label: 'Result', col: map['result'] },
    { label: 'Avg Score', col: map['average room score'], pct: true },
    { label: 'Total', col: map['total room score'] },
    { label: 'Audit ID', col: map['audit id'] },
    { label: 'Comments', col: map['comments'] }
  ];

  var rows = [];
  var n = lastRow - headerRow;
  if (n > 0) {
    var data = sh.getRange(headerRow + 1, 1, n, lastCol).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][idCol - 1] || '').trim() === '') continue;
      var rec = [];
      defs.forEach(function (d) {
        var v = d.col ? data[i][d.col - 1] : '';
        if (d.date && v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        else if (d.pct && typeof v === 'number') v = (Math.round(v * 1000) / 10) + '%';
        rec.push(v === null || v === undefined ? '' : String(v));
      });
      rows.push(rec);
    }
  }
  rows.reverse(); // newest first (records append to the bottom)

  var html = buildSubmissionsHtml_(defs.map(function (d) { return d.label; }), rows);
  ui.showModelessDialog(HtmlService.createHtmlOutput(html).setWidth(760).setHeight(540), 'Past audit submissions');
}

function buildSubmissionsHtml_(headers, rows) {
  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c];
    });
  }
  var brands = {};
  rows.forEach(function (r) { if (r[0]) brands[r[0]] = true; });
  var opts = '<option value="">All facilities</option>'
    + Object.keys(brands).sort().map(function (b) { return '<option value="' + esc(b) + '">' + esc(b) + '</option>'; }).join('');

  var thead = '<tr>' + headers.map(function (h) { return '<th>' + esc(h) + '</th>'; }).join('') + '</tr>';
  var tbody = rows.length
    ? rows.map(function (r) {
        return '<tr data-brand="' + esc(r[0]) + '">' + r.map(function (v, i) {
          return '<td' + (headers[i] === 'Comments' ? ' class="cmt"' : '') + '>' + esc(v) + '</td>';
        }).join('') + '</tr>';
      }).join('')
    : '<tr><td colspan="' + headers.length + '" style="text-align:center;color:#6b7280;padding:20px">No audits recorded yet.</td></tr>';

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
    + 'body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:16px;color:#111827;font-size:13px}'
    + 'h2{font-size:16px;margin:0 0 10px}'
    + '.bar{margin:0 0 10px}select{font-size:13px;padding:6px;border:1px solid #cbd5e1;border-radius:6px}'
    + '.count{color:#6b7280;margin-left:8px}'
    + 'table{border-collapse:collapse;width:100%}'
    + 'th,td{border:1px solid #e5e7eb;padding:6px 8px;text-align:left;vertical-align:top}'
    + 'th{background:#1F4E79;color:#fff;position:sticky;top:0}'
    + 'tr:nth-child(even) td{background:#f8fafc}'
    + 'td.cmt{max-width:260px;white-space:pre-wrap;color:#374151}'
    + '</style></head><body>'
    + '<h2>Past audit submissions</h2>'
    + '<div class="bar">Facility: <select id="f" onchange="filt()">' + opts + '</select><span class="count" id="cnt"></span></div>'
    + '<table><thead>' + thead + '</thead><tbody id="tb">' + tbody + '</tbody></table>'
    + '<script>function filt(){var v=document.getElementById("f").value;var tr=document.querySelectorAll("#tb tr");var n=0;'
    + 'for(var i=0;i<tr.length;i++){var b=tr[i].getAttribute("data-brand");var show=(!v||b===v);'
    + 'tr[i].style.display=show?"":"none";if(show&&b!==null)n++;}'
    + 'document.getElementById("cnt").textContent=n+" audit(s)";}filt();</script>'
    + '</body></html>';
}
