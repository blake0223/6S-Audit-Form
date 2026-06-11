/**
 * RecordResults — HTML form to enter a completed audit's scores and write them to
 * the KPI Data tab. The user types the Audit ID from the printed form, picks a
 * room, enters 0–3 for each item, and submits. One summary row per room is
 * appended to KPI Data; optionally the scores are written back into the room's
 * column on the audit sheet.
 *
 * Menu: 6S Audit Tools ▸ Record audit results   (wired up in onOpen.gs)
 */
function recordResults() {
  var layout = getAuditLayout_();
  if (!layout.rooms.length) {
    SpreadsheetApp.getUi().alert('No room columns found on the audit sheet — add a room first.');
    return;
  }
  var dlg = HtmlService.createHtmlOutput(buildRecordHtml_(layout)).setWidth(470).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(dlg, 'Record audit results');
}

/** Read the audit sheet structure: rooms, items (by section), max, location. */
function getAuditLayout_() {
  var sheet = auditSheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headerRow = findHeaderRow_(sheet, lastRow);
  var qCol = findQuestionCol_(sheet, headerRow) || QUESTION_COL;
  var totalCol = findColByHeader_(sheet, headerRow, 'total score') || (lastCol);

  var hdr = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var rooms = [];
  for (var idx = qCol; idx <= totalCol - 2; idx++) {           // columns between questions and total
    var name = String(hdr[idx]).trim();
    if (name) rooms.push({ name: name, col: idx + 1 });
  }

  var data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, 3).getValues();
  var items = [], section = '';
  for (var i = 0; i < data.length; i++) {
    var a = data[i][0];
    if (String(a).indexOf('◆') >= 0) { section = String(a).replace(/◆/g, '').trim(); continue; }
    if (typeof a === 'number' && a > 0) {
      var label = String(data[i][1]).trim() || String(data[i][2]).trim();
      items.push({ row: headerRow + 1 + i, section: section, no: a, label: label });
    }
  }
  return { rooms: rooms, items: items, max: items.length * 3, location: readLocation_() };
}

/** Compute a room's result, append it to KPI Data, optionally write back to the sheet. */
function submitRoomResults(id, roomName, scores, writeBack) {
  id = String(id).trim();
  if (!id) throw new Error('Audit ID is required.');
  var sheet = auditSheet_();
  var layout = getAuditLayout_();

  var roomCol = 0;
  for (var k = 0; k < layout.rooms.length; k++) {
    if (layout.rooms[k].name === roomName) { roomCol = layout.rooms[k].col; break; }
  }

  var total = 0, zeros = 0, scored = 0;
  for (var i = 0; i < layout.items.length; i++) {
    var v = scores[i];
    if (v === '' || v === null || v === undefined) continue;
    v = Number(v);
    if (isNaN(v)) continue;
    total += v; scored++;
    if (v === 0) zeros++;
    if (writeBack && roomCol) sheet.getRange(layout.items[i].row, roomCol).setValue(v);
  }

  var max = scored * 3;
  var pct = max ? Math.round(total / max * 1000) / 10 : 0;
  var result = (max && pct >= 67 && zeros === 0) ? 'PASS' : 'FAIL';
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  appendKpiRow_({
    id: id, date: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    location: layout.location, room: roomName,
    total: total, max: max, pct: pct, zeros: zeros, result: result
  });
  return { room: roomName, total: total, max: max, pct: pct, zeros: zeros, result: result };
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

/** Build the data-entry dialog HTML (layout embedded as JSON). */
function buildRecordHtml_(layout) {
  return ''
    + '<style>body{font-family:Arial;margin:0;padding:14px;font-size:13px;color:#111827}'
    + 'h3{margin:0 0 10px}label{font-weight:bold}select,input{font-size:13px}'
    + '.fld{margin-bottom:8px}.row{display:flex;justify-content:space-between;align-items:center;'
    + 'padding:3px 0;border-bottom:1px solid #eee}.row span{flex:1;padding-right:8px}'
    + '.sec{font-weight:bold;background:#E5E7EB;padding:4px 6px;margin-top:10px}'
    + '#go{margin-top:14px;background:#1F2A37;color:#fff;border:0;border-radius:6px;'
    + 'padding:10px 18px;font-weight:bold;cursor:pointer}#msg{margin-top:10px;font-weight:bold}</style>'
    + '<h3>Record audit results</h3>'
    + '<div class="fld"><label>Audit ID</label><br>'
    + '<input id="aid" style="width:100%" placeholder="6S-YYYYMMDD-XXXX"></div>'
    + '<div class="fld"><label>Room</label><br><select id="room" style="width:100%"></select></div>'
    + '<div class="fld"><label><input type="checkbox" id="wb" checked> '
    + 'Also fill this room’s column on the sheet</label></div>'
    + '<div id="items"></div>'
    + '<button id="go">Record results</button><div id="msg"></div>'
    + '<script>'
    + 'var L=' + JSON.stringify(layout) + ';'
    + 'var rs=document.getElementById("room");'
    + 'L.rooms.forEach(function(r,i){var o=document.createElement("option");o.value=i;o.text=r.name;rs.add(o);});'
    + 'var box=document.getElementById("items"),cur="";'
    + 'L.items.forEach(function(it,idx){if(it.section!==cur){cur=it.section;'
    + 'var s=document.createElement("div");s.className="sec";s.textContent=it.section;box.appendChild(s);}'
    + 'var d=document.createElement("div");d.className="row";var sp=document.createElement("span");'
    + 'sp.textContent=it.no+"  "+(it.label||"");var se=document.createElement("select");se.id="s"+idx;'
    + '["","0","1","2","3"].forEach(function(v){var o=document.createElement("option");o.value=v;'
    + 'o.text=(v===""?"–":v);se.add(o);});d.appendChild(sp);d.appendChild(se);box.appendChild(d);});'
    + 'document.getElementById("go").addEventListener("click",function(){'
    + 'var id=document.getElementById("aid").value.trim();'
    + 'if(!id){alert("Enter the Audit ID printed on the form.");return;}'
    + 'var room=L.rooms[rs.value];'
    + 'var scores=L.items.map(function(it,idx){return document.getElementById("s"+idx).value;});'
    + 'var wb=document.getElementById("wb").checked;var b=this;b.disabled=true;b.textContent="Saving…";'
    + 'google.script.run.withSuccessHandler(function(res){'
    + 'document.getElementById("msg").textContent="✔ Saved "+res.room+": "+res.total+"/"+res.max+" ("+res.pct+"%) — "+res.result;'
    + 'setTimeout(google.script.host.close,1600);})'
    + '.withFailureHandler(function(e){document.getElementById("msg").textContent="Error: "+e.message;'
    + 'b.disabled=false;b.textContent="Record results";})'
    + '.submitRoomResults(id,room.name,scores,wb);});'
    + '</script>';
}
