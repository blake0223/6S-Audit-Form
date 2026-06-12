/**
 * PrintForm — pick the facility first, then Monthly audit / Weekly checklist, and
 * open that facility's tab as a clean, print-ready PDF in a new browser tab.
 *
 *   • Monthly audit  → portrait, fit-to-width, rows 4–5 hidden for the export.
 *   • Weekly checklist → landscape, fit-to-width (it's a wide grid).
 *
 * Facilities are grouped by the first word of the tab name, so a facility's
 * Monthly Audit and Checklist tabs map to the same dropdown entry.
 *
 * Menu: 6S Audit Tools ▸ Print blank form
 */

var SKIP_ROW_START = 4; // monthly audit: first row to leave off the printout
var SKIP_ROW_COUNT = 2; // rows 4 and 5

function printForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var byKey = {};
  ss.getSheets().forEach(function (s) {
    var n = s.getName();
    var isMonthly = /monthly audit/i.test(n);
    var isChecklist = /checklist/i.test(n);
    if (!isMonthly && !isChecklist) return;
    var key = String(n).trim().split(/\s+/)[0].toUpperCase();
    if (!byKey[key]) byKey[key] = { key: key, label: facilityLabel_(n), monthly: '', checklist: '' };
    if (isMonthly) { byKey[key].monthly = n; byKey[key].label = facilityLabel_(n); }
    else { byKey[key].checklist = n; }
  });
  var facilities = Object.keys(byKey).map(function (k) { return byKey[k]; });
  if (!facilities.length) {
    SpreadsheetApp.getUi().alert('No Monthly Audit or Weekly Checklist tabs found.');
    return;
  }

  var html = HtmlService.createHtmlOutput(buildPrintHtml_(facilities)).setWidth(390).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Print blank form');
}

/** Build the PDF for one tab and return it base64-encoded. Called from the dialog. */
function makeFormPdf(sheetName, type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Tab not found: ' + sheetName);

  var monthly = (type === 'monthly');
  var b64;
  try {
    if (monthly) { sheet.hideRows(SKIP_ROW_START, SKIP_ROW_COUNT); SpreadsheetApp.flush(); }
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?'
      + 'format=pdf'
      + '&gid=' + sheet.getSheetId()
      + '&portrait=' + (monthly ? 'true' : 'false')
      + '&fitw=true&size=letter&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=true'
      + '&top_margin=0.50&bottom_margin=0.50&left_margin=0.40&right_margin=0.40'
      + '&r1=0&c1=0&r2=' + lastRow + '&c2=' + lastCol;
    b64 = Utilities.base64Encode(
      UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
        .getBlob().getBytes());
  } finally {
    if (monthly) { sheet.showRows(SKIP_ROW_START, SKIP_ROW_COUNT); SpreadsheetApp.flush(); }
  }
  return b64;
}

/** Strip the form-type words from a tab name to get a clean facility label. */
function facilityLabel_(name) {
  return String(name)
    .replace(/6S/ig, '')
    .replace(/monthly audit sheet|monthly audit|daily checklist|weekly checklist|checklist|facility summary|sheet/ig, '')
    .replace(/\s+/g, ' ').trim() || name;
}

function buildPrintHtml_(facilities) {
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
    + '<button id="go">Open &amp; print</button><div id="msg"></div>'
    + '<script>'
    + 'var F=' + JSON.stringify(facilities) + ';'
    + 'var sel=document.getElementById("fac");'
    + 'F.forEach(function(f,i){var o=document.createElement("option");o.value=i;o.text=f.label;sel.add(o);});'
    + 'function type(){return document.querySelector("input[name=t]:checked").value;}'
    + 'document.getElementById("go").addEventListener("click",function(){'
    + 'var f=F[sel.value];var tab=(type()==="monthly")?f.monthly:f.checklist;'
    + 'if(!tab){document.getElementById("msg").textContent="This facility has no "+(type()==="monthly"?"Monthly Audit":"Weekly Checklist")+" tab.";return;}'
    + 'var b=this;b.disabled=true;document.getElementById("msg").textContent="Generating…";'
    + 'var w=window.open("about:blank","_blank");'
    + 'google.script.run.withSuccessHandler(function(b64){'
    + 'var bytes=Uint8Array.from(atob(b64),function(c){return c.charCodeAt(0);});'
    + 'var url=URL.createObjectURL(new Blob([bytes],{type:"application/pdf"}));'
    + 'if(w){w.location=url;}else{window.open(url,"_blank");}google.script.host.close();})'
    + '.withFailureHandler(function(e){if(w){w.close();}b.disabled=false;'
    + 'document.getElementById("msg").textContent="Error: "+e.message;})'
    + '.makeFormPdf(tab,type());});'
    + '</script>';
}
