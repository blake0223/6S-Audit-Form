/**
 * PrintForm — pick a form type (Monthly audit / Weekly checklist) and a facility
 * from a dropdown, then open that facility's tab as a clean, print-ready PDF in a
 * new browser tab to print. No download to Drive.
 *
 *   • Monthly audit  → portrait, fit-to-width, rows 4–5 hidden for the export.
 *   • Weekly checklist → landscape, fit-to-width (it's a wide grid).
 *
 * The facility dropdown is built from the tab names: tabs containing "Monthly
 * Audit" are the monthly forms; tabs containing "Checklist" are the weekly forms.
 *
 * Menu: 6S Audit Tools ▸ Print blank form   (wired up in onOpen.gs)
 */

var SKIP_ROW_START = 4; // monthly audit: first row to leave off the printout
var SKIP_ROW_COUNT = 2; // rows 4 and 5

function printForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var monthly = [], weekly = [];
  ss.getSheets().forEach(function (s) {
    var n = s.getName();
    if (/monthly audit/i.test(n)) monthly.push({ name: n, label: facilityLabel_(n) });
    else if (/checklist/i.test(n)) weekly.push({ name: n, label: facilityLabel_(n) });
  });
  if (!monthly.length && !weekly.length) {
    SpreadsheetApp.getUi().alert('No Monthly Audit or Weekly Checklist tabs found.');
    return;
  }

  var html = HtmlService.createHtmlOutput(buildPrintHtml_(monthly, weekly))
    .setWidth(380).setHeight(260);
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
      + '&fitw=true'
      + '&size=letter'
      + '&gridlines=false'
      + '&printtitle=false'
      + '&sheetnames=false'
      + '&pagenumbers=true'
      + '&top_margin=0.50&bottom_margin=0.50&left_margin=0.40&right_margin=0.40'
      + '&r1=0&c1=0&r2=' + lastRow + '&c2=' + lastCol;
    b64 = Utilities.base64Encode(
      UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
      }).getBlob().getBytes());
  } finally {
    if (monthly) { sheet.showRows(SKIP_ROW_START, SKIP_ROW_COUNT); SpreadsheetApp.flush(); }
  }
  return b64;
}

/* ---------- helpers ---------- */

/** Strip the form-type words from a tab name to get a clean facility label. */
function facilityLabel_(name) {
  return String(name)
    .replace(/6S/ig, '')
    .replace(/monthly audit sheet|monthly audit|daily checklist|weekly checklist|checklist|facility summary|sheet/ig, '')
    .replace(/\s+/g, ' ').trim() || name;
}

function buildPrintHtml_(monthly, weekly) {
  return ''
    + '<style>body{font-family:Arial;margin:0;padding:16px;font-size:13px;color:#111827}'
    + 'h3{margin:0 0 12px}.row{margin-bottom:12px}label{font-weight:bold}'
    + 'select{width:100%;font-size:13px;padding:4px}'
    + '#go{background:#1F2A37;color:#fff;border:0;border-radius:6px;padding:10px 18px;font-weight:bold;cursor:pointer}'
    + '#msg{margin-top:10px;color:#6B7280}</style>'
    + '<h3>Print blank form</h3>'
    + '<div class="row"><label>Form</label><br>'
    + '<label style="font-weight:normal"><input type="radio" name="t" value="monthly" checked> Monthly audit</label>&nbsp;&nbsp;'
    + '<label style="font-weight:normal"><input type="radio" name="t" value="checklist"> Weekly checklist</label></div>'
    + '<div class="row"><label>Facility</label><br><select id="fac"></select></div>'
    + '<button id="go">Open &amp; print</button><div id="msg"></div>'
    + '<script>'
    + 'var M=' + JSON.stringify(monthly) + ',W=' + JSON.stringify(weekly) + ';'
    + 'var fac=document.getElementById("fac");'
    + 'function type(){return document.querySelector("input[name=t]:checked").value;}'
    + 'function fill(){var L=(type()==="monthly")?M:W;fac.innerHTML="";'
    + 'if(!L.length){var o=document.createElement("option");o.text="(none found)";o.disabled=true;fac.add(o);return;}'
    + 'L.forEach(function(x){var o=document.createElement("option");o.value=x.name;o.text=x.label;fac.add(o);});}'
    + 'Array.prototype.forEach.call(document.getElementsByName("t"),function(r){r.addEventListener("change",fill);});fill();'
    + 'document.getElementById("go").addEventListener("click",function(){'
    + 'if(!fac.value){return;}var b=this;b.disabled=true;document.getElementById("msg").textContent="Generating…";'
    + 'var w=window.open("about:blank","_blank");'
    + 'google.script.run.withSuccessHandler(function(b64){'
    + 'var bytes=Uint8Array.from(atob(b64),function(c){return c.charCodeAt(0);});'
    + 'var url=URL.createObjectURL(new Blob([bytes],{type:"application/pdf"}));'
    + 'if(w){w.location=url;}else{window.open(url,"_blank");}google.script.host.close();})'
    + '.withFailureHandler(function(e){if(w){w.close();}b.disabled=false;'
    + 'document.getElementById("msg").textContent="Error: "+e.message;})'
    + '.makeFormPdf(fac.value,type());});'
    + '</script>';
}

/** Minimal HTML escaping (kept for reuse). */
function escapeHtml_(s) {
  return String(s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}
