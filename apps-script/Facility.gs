/**
 * Facility — the per-facility picker shared by the menu tools. Each tool opens a
 * small dialog: choose the facility (and any tool-specific input), then it runs
 * against that facility's tab.
 *
 * Facility tabs are recognised by name: tabs containing "Monthly Audit" are the
 * monthly forms; tabs containing "Checklist" are the weekly forms. The dropdown
 * label is the tab name minus the form-type words (see facilityLabel_ in PrintForm.gs).
 */
function listFacilityTabs_(type) {
  var rx = (type === 'checklist') ? /checklist/i : /monthly audit/i;
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(function (s) { return rx.test(s.getName()); })
    .map(function (s) { return { name: s.getName(), label: facilityLabel_(s.getName()) }; });
}

/**
 * Open a facility-picker dialog.
 *   cfg = { title, type:'monthly'|'checklist', button, callback,
 *           fields:[{id,label,kind:'text'|'select',options,placeholder}] }
 * On submit it calls google.script.run[callback](facilityTabName, ...fieldValues),
 * shows the returned message, then closes.
 */
function showToolDialog_(cfg) {
  var ui = SpreadsheetApp.getUi();
  var tabs = listFacilityTabs_(cfg.type);
  if (!tabs.length) {
    ui.alert('No ' + (cfg.type === 'checklist' ? 'Weekly Checklist' : 'Monthly Audit') + ' tabs found.');
    return;
  }
  var html = HtmlService.createHtmlOutput(buildToolHtml_(cfg, tabs))
    .setWidth(400).setHeight(cfg.height || 250);
  ui.showModalDialog(html, cfg.title);
}

function buildToolHtml_(cfg, tabs) {
  var fields = cfg.fields || [];
  var fieldHtml = fields.map(function (f) {
    if (f.kind === 'select') {
      return '<div class="row"><label>' + escHtml_(f.label) + '</label><br><select id="' + f.id + '">'
        + f.options.map(function (o) { return '<option>' + escHtml_(o) + '</option>'; }).join('')
        + '</select></div>';
    }
    return '<div class="row"><label>' + escHtml_(f.label) + '</label><br><input id="' + f.id + '" type="text"'
      + (f.placeholder ? ' placeholder="' + escAttr_(f.placeholder) + '"' : '') + '></div>';
  }).join('');
  var ids = fields.map(function (f) { return f.id; });

  return ''
    + '<style>body{font-family:Arial;margin:0;padding:16px;font-size:13px;color:#111827}h3{margin:0 0 12px}'
    + '.row{margin-bottom:10px}label{font-weight:bold}select,input{width:100%;font-size:13px;padding:5px;box-sizing:border-box}'
    + '#go{background:#1F2A37;color:#fff;border:0;border-radius:6px;padding:10px 18px;font-weight:bold;cursor:pointer}'
    + '#msg{margin-top:10px;color:#6B7280}</style>'
    + '<h3>' + escHtml_(cfg.title) + '</h3>'
    + '<div class="row"><label>Facility</label><br><select id="_fac">'
    + tabs.map(function (t) { return '<option value="' + escAttr_(t.name) + '">' + escHtml_(t.label) + '</option>'; }).join('')
    + '</select></div>'
    + fieldHtml
    + '<button id="go">' + escHtml_(cfg.button || 'Apply') + '</button><div id="msg"></div>'
    + '<script>'
    + 'var IDS=' + JSON.stringify(ids) + ',CB=' + JSON.stringify(cfg.callback) + ';'
    + 'document.getElementById("go").addEventListener("click",function(){'
    + 'var args=[document.getElementById("_fac").value];'
    + 'IDS.forEach(function(id){args.push(document.getElementById(id).value);});'
    + 'var b=this;b.disabled=true;document.getElementById("msg").textContent="Working…";'
    + 'var r=google.script.run'
    + '.withSuccessHandler(function(m){document.getElementById("msg").textContent=m||"Done.";setTimeout(google.script.host.close,1300);})'
    + '.withFailureHandler(function(e){b.disabled=false;document.getElementById("msg").textContent="Error: "+e.message;});'
    + 'r[CB].apply(r,args);});'
    + '</script>';
}

function escHtml_(s) {
  return String(s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}
function escAttr_(s) { return escHtml_(s); }
