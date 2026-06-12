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
  ui.showModelessDialog(html, cfg.title);
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
    + 'var fac=document.getElementById("_fac").value;var f={};'
    + 'IDS.forEach(function(id){f[id]=document.getElementById(id).value;});'
    + 'var b=this;b.disabled=true;document.getElementById("msg").textContent="Working…";'
    + 'google.script.run'
    + '.withSuccessHandler(function(m){document.getElementById("msg").textContent=m||"Done.";setTimeout(google.script.host.close,1300);})'
    + '.withFailureHandler(function(e){b.disabled=false;var m=(e&&e.message)||"";'
    + 'document.getElementById("msg").textContent=/authoriz/i.test(m)'
    + '?"Permission needed: reload the sheet, run this tool, and approve the prompt that appears.":"Error: "+m;})'
    + '.runTool(CB,fac,f);});'
    + '</script>';
}

/**
 * Called at the start of every tool. The first time a user runs any tool, this
 * forces Google's one-time permission prompt to include WRITE access (a read-only
 * touch wouldn't), because it statically performs a reversible write (a temp sheet
 * created and deleted). Guarded so the write only happens once.
 */
function ensureAuthorized_() {
  // PER-USER flag (UserProperties), not script-wide — so every user's first click
  // actually runs the write below, which forces Google's one-time "edit this
  // spreadsheet" prompt in this (promptable) menu context. A script-wide flag
  // would skip the prompt for everyone after the first person.
  var up = PropertiesService.getUserProperties();
  if (up.getProperty('authedV3')) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stale = ss.getSheetByName('_6s_auth_check');
  if (stale) ss.deleteSheet(stale);
  var tmp = ss.insertSheet('_6s_auth_check');
  ss.deleteSheet(tmp);
  up.setProperty('authedV3', '1');
}

/**
 * Dispatcher called by every picker dialog (avoids google.script.run[name].apply,
 * which is unreliable in the dialog sandbox). Routes to the right core function.
 */
function runTool(cb, sheetName, f) {
  f = f || {};
  switch (cb) {
    case 'addLocationFor':      return addLocationFor(sheetName, f.name);
    case 'addGradingItemFor':   return addGradingItemFor(sheetName, f.section, f.name);
    case 'addChecklistItemFor': return addChecklistItemFor(sheetName, f.name);
    case 'recordResultsFor':    return recordResultsFor(sheetName);
    default: throw new Error('Unknown tool: ' + cb);
  }
}

function escHtml_(s) {
  return String(s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}
function escAttr_(s) { return escHtml_(s); }
