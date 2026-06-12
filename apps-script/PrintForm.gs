/**
 * PrintForm — pick the facility, then Monthly audit / Weekly checklist, and open
 * that tab's print-ready PDF.
 *
 * IMPORTANT: the PDF is fetched by the USER'S BROWSER (they're already signed in),
 * not by the script. That means the script never needs the "connect to an external
 * service" permission — the only permission the whole tool set needs is "edit this
 * spreadsheet," which keeps the consent simple for non-technical users.
 *
 *   • Monthly audit  → portrait, fit-to-width.
 *   • Weekly checklist → landscape, fit-to-width (wide grid).
 *
 * Menu: 6S Audit Tools ▸ Print blank form
 */
function printForm() {
  ensureAuthorized_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var byKey = {};
  ss.getSheets().forEach(function (s) {
    var n = s.getName();
    var isMonthly = /monthly audit/i.test(n);
    var isChecklist = /checklist/i.test(n);
    if (!isMonthly && !isChecklist) return;
    var key = String(n).trim().split(/\s+/)[0].toUpperCase();
    if (!byKey[key]) byKey[key] = { label: facilityLabel_(n), monthly: null, checklist: null };
    if (isMonthly) { byKey[key].monthly = s.getSheetId(); byKey[key].label = facilityLabel_(n); }
    else { byKey[key].checklist = s.getSheetId(); }
  });
  var facilities = Object.keys(byKey).map(function (k) { return byKey[k]; });
  if (!facilities.length) {
    SpreadsheetApp.getUi().alert('No Monthly Audit or Weekly Checklist tabs found.');
    return;
  }

  var html = HtmlService.createHtmlOutput(buildPrintHtml_(ssId, facilities))
    .setWidth(390).setHeight(250);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Print blank form');
}

/** Strip the form-type words from a tab name to get a clean facility label. */
function facilityLabel_(name) {
  return String(name)
    .replace(/6S/ig, '')
    .replace(/monthly audit sheet|monthly audit|daily checklist|weekly checklist|checklist|facility summary|sheet/ig, '')
    .replace(/\s+/g, ' ').trim() || name;
}

function buildPrintHtml_(ssId, facilities) {
  return ''
    + '<style>body{font-family:Arial;margin:0;padding:16px;font-size:13px;color:#111827}'
    + 'h3{margin:0 0 12px}.row{margin-bottom:12px}label{font-weight:bold}'
    + 'select{width:100%;font-size:13px;padding:5px;box-sizing:border-box}'
    + '#go{background:#1F2A37;color:#fff;border:0;border-radius:6px;padding:10px 18px;font-weight:bold;cursor:pointer}'
    + '#msg{margin-top:10px;color:#6B7280}</style>'
    + '<h3>Print blank form</h3>'
    + '<div class="row"><label>Facility</label><br><select id="fac"></select></div>'
    + '<div class="row"><label>Form</label><br>'
    + '<label style="font-weight:normal"><input type="radio" name="t" value="monthly" checked> Monthly audit</label>&nbsp;&nbsp;'
    + '<label style="font-weight:normal"><input type="radio" name="t" value="checklist"> Weekly checklist</label></div>'
    + '<button id="go">Open PDF</button><div id="msg"></div>'
    + '<script>'
    + 'var SS=' + JSON.stringify(ssId) + ',F=' + JSON.stringify(facilities) + ';'
    + 'var sel=document.getElementById("fac");'
    + 'F.forEach(function(f,i){var o=document.createElement("option");o.value=i;o.text=f.label;sel.add(o);});'
    + 'function type(){return document.querySelector("input[name=t]:checked").value;}'
    + 'document.getElementById("go").addEventListener("click",function(){'
    + 'var f=F[sel.value];var t=type();var gid=(t==="monthly")?f.monthly:f.checklist;'
    + 'if(gid===null||gid===undefined){document.getElementById("msg").textContent="This facility has no "+(t==="monthly"?"Monthly Audit":"Weekly Checklist")+" tab.";return;}'
    + 'var url="https://docs.google.com/spreadsheets/d/"+SS+"/export?format=pdf&gid="+gid'
    + '+"&portrait="+(t==="monthly"?"true":"false")'
    + '+"&fitw=true&size=letter&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=true'
    + '&top_margin=0.50&bottom_margin=0.50&left_margin=0.40&right_margin=0.40";'
    + 'window.open(url,"_blank");google.script.host.close();});'
    + '</script>';
}
