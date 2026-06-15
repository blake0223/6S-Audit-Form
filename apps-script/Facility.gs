/**
 * Facility — a small step-by-step dialog with CLICKABLE buttons for picking a
 * facility (and a 6S category, where needed), plus any text fields. The dialog
 * runs the chosen tool on the server via google.script.run.
 *
 * This is safe now that nothing in the project uses an external service: the only
 * permission required is "edit this spreadsheet", which the user has already
 * granted, so google.script.run runs without any consent prompt.
 *
 * Setup/master tabs (Safety Setup, Initiation, etc.) are excluded from the lists.
 */

var EXCLUDE_TABS = /initiation|safety setup|instruction|overview|kpi data/i;

/** Facility tabs of a type, as [{ name, label }]. */
function listFacilityTabs_(type) {
  var rx = (type === 'checklist') ? /checklist/i : /monthly audit/i;
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(function (s) { var n = s.getName(); return rx.test(n) && !EXCLUDE_TABS.test(n); })
    .map(function (s) { return { name: s.getName(), label: facilityLabel_(s.getName()) }; });
}

/**
 * Open the button-driven tool dialog. cfg:
 *   title, callback, argOrder:[keys], button, resultIsUrl,
 *   type            – fixed facility type (fills cfg.facilities)
 *   facilitiesByType – {monthly:[...],checklist:[...]} for type-dependent lists
 *   steps:[ {kind:'facility'|'pick'|'text', key, label, ...} ]
 */
function showToolDialog_(cfg) {
  if (cfg.type && !cfg.facilities) cfg.facilities = listFacilityTabs_(cfg.type);
  if (cfg.facilities && !cfg.facilities.length) {
    SpreadsheetApp.getUi().alert('No matching facility tabs were found.');
    return;
  }
  var json = JSON.stringify(cfg).replace(/<\//g, '<\\/');
  var html = HtmlService.createHtmlOutput(buildToolHtml_(json)).setWidth(360).setHeight(440);
  SpreadsheetApp.getUi().showModelessDialog(html, cfg.title);
}

function buildToolHtml_(cfgJson) {
  return '<!DOCTYPE html><html><head><base target="_top"><style>'
    + 'body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:18px;color:#222}'
    + 'h2{font-size:16px;margin:0 0 2px}'
    + '.sub{font-size:12px;color:#666;margin:0 0 14px}'
    + '.btn{display:block;width:100%;box-sizing:border-box;text-align:left;background:#fff;'
    + 'border:1px solid #cbd5e1;border-radius:8px;padding:12px 14px;margin:7px 0;font-size:14px;cursor:pointer}'
    + '.btn:hover{background:#f1f5f9;border-color:#1F4E79}'
    + '.primary{background:#1F4E79;color:#fff;border:none;text-align:center;font-weight:bold}'
    + '.primary:hover{background:#163a5b}'
    + 'input{width:100%;box-sizing:border-box;padding:11px;font-size:14px;border:1px solid #cbd5e1;border-radius:8px;margin:4px 0 12px}'
    + '.link{display:block;text-align:center;background:#1F4E79;color:#fff;text-decoration:none;'
    + 'border-radius:8px;padding:13px;font-size:15px;font-weight:bold;margin:10px 0}'
    + '.msg{font-size:14px;line-height:1.45}.err{color:#b91c1c}'
    + '#back{font-size:12px;color:#1F4E79;cursor:pointer;display:inline-block;margin-top:10px}'
    + '</style></head><body><div id="app"></div><script>'
    + 'var CFG=' + cfgJson + ';var state={};var stepIdx=0;'
    + 'var app=document.getElementById("app");'
    + 'function esc(s){return String(s).replace(/[&<>"]/g,function(c){return({"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;"})[c];});}'
    + 'function facs(st){return st.typeFrom?(CFG.facilitiesByType[state[st.typeFrom]]||[]):(CFG.facilities||[]);}'
    + 'function render(){var st=CFG.steps[stepIdx];if(!st){return submit();}'
    + 'var h="<h2>"+esc(CFG.title)+"</h2>";if(st.label)h+="<p class=\\"sub\\">"+esc(st.label)+"</p>";'
    + 'if(st.kind==="text"){h+="<input id=tf placeholder=\\""+esc(st.placeholder||"")+"\\">";'
    + 'h+="<button class=\\"btn primary\\" id=go>"+esc(CFG.button||"Continue")+"</button>";}'
    + 'else{var list=(st.kind==="facility")?facs(st):st.options;'
    + 'if(!list||!list.length){h+="<p class=\\"msg err\\">Nothing to choose here.</p>";}'
    + 'else{list.forEach(function(o){var lab=(st.kind==="facility")?o.label:o.label;var val=(st.kind==="facility")?o.name:o.value;'
    + 'h+="<button class=\\"btn pick\\" data-val=\\""+esc(val)+"\\">"+esc(lab)+"</button>";});}}'
    + 'if(stepIdx>0)h+="<span id=back>\\u2039 Back</span>";'
    + 'app.innerHTML=h;wire(st);}'
    + 'function wire(st){var b=document.getElementById("back");if(b)b.onclick=function(){stepIdx--;render();};'
    + 'if(st.kind==="text"){document.getElementById("go").onclick=function(){'
    + 'var v=document.getElementById("tf").value.trim();if(!v&&!st.optional){alert("Please enter a value.");return;}'
    + 'state[st.key]=v;stepIdx++;render();};return;}'
    + 'Array.prototype.forEach.call(app.querySelectorAll(".pick"),function(btn){btn.onclick=function(){'
    + 'state[st.key]=btn.getAttribute("data-val");stepIdx++;render();};});}'
    + 'function submit(){app.innerHTML="<p class=\\"msg\\">Working\\u2026</p>";'
    + 'var args=CFG.argOrder.map(function(k){return state[k];});'
    + 'var r=google.script.run.withSuccessHandler(done).withFailureHandler(fail);'
    + 'r[CFG.callback].apply(r,args);}'
    + 'function done(res){if(CFG.resultIsUrl){app.innerHTML="<p class=\\"msg\\">Your print-ready form is ready.</p>"'
    + '+"<a class=\\"link\\" href=\\""+esc(res)+"\\" target=_blank rel=noopener>Open PDF in new tab</a>"'
    + '+"<p class=\\"sub\\">Then use your browser\\u2019s Print or Download.</p>";}'
    + 'else{app.innerHTML="<p class=\\"msg\\">"+esc(res)+"</p><button class=\\"btn primary\\" onclick=\\"google.script.host.close()\\">Done</button>";}}'
    + 'function fail(e){app.innerHTML="<p class=\\"msg err\\">Error: "+esc(e.message)+"</p>'
    + '<button class=\\"btn primary\\" onclick=\\"google.script.host.close()\\">Close</button>";}'
    + 'render();</script></body></html>';
}
