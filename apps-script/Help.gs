/**
 * authorize — "Enable tools (first-time setup)" menu item. Each user runs this
 * once: clicking it touches the services that need permission, so Google shows its
 * normal in-sheet authorization prompt (no editor needed). After approving, every
 * tool works because the dialogs' google.script.run calls are then permitted.
 */
function authorize() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Reversible WRITE  → spreadsheets scope (the tools edit the sheet).
  var stale = ss.getSheetByName('_6s_auth_check');
  if (stale) ss.deleteSheet(stale);
  var tmp = ss.insertSheet('_6s_auth_check');
  ss.deleteSheet(tmp);

  // 2) Properties        → used by location-color + auth flag.
  PropertiesService.getDocumentProperties().getProperty('x');
  PropertiesService.getScriptProperties().getProperty('x');

  // NOTE: no UrlFetch here — printing is done by the browser, so the script never
  // needs the "connect to an external service" permission.
  SpreadsheetApp.getUi().alert('6S Audit Tools authorized — every tool is ready to use.');
}

/**
 * Help — the "How to use this sheet" dialog shown from the menu.
 * Menu: 6S Audit Tools ▸ How to use this sheet
 */
function showHelp() {
  var html = HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;padding:18px;font-size:13px;line-height:1.5;color:#111827">'
    + '<h2 style="margin:0 0 10px;font-size:18px">Welcome to the Hickory 6S Audit Forms</h2>'
    + '<p>Find <b>your facility&rsquo;s tabs</b> &mdash; they&rsquo;re color-coded. Each facility has a '
    + '<b>Monthly Audit</b>, a <b>Weekly Checklist</b>, and a <b>Facility Summary</b>.</p>'
    + '<p>Configure your forms from the <b>6S Audit Tools</b> menu above:</p>'
    + '<ul style="margin:0 0 10px 18px;padding:0">'
    + '<li><b>Monthly audit &rsaquo; Add room (column)</b> &mdash; add each location/area in your facility to the Monthly Audit.</li>'
    + '<li><b>Monthly audit &rsaquo; Add grading item (row)</b> &mdash; add a 6S grading line to a category (click into that category first).</li>'
    + '<li><b>Weekly checklist &rsaquo; Add checklist item</b> &mdash; add facility-specific daily checks (e.g. &ldquo;Propane turned off&rdquo;, &ldquo;Forklift parked with forks down&rdquo;).</li>'
    + '</ul>'
    + '<p><b>Before you start:</b> make sure every location in your facility has a column in the Monthly Audit, '
    + 'and add any location-specific checks to your Weekly Checklist.</p>'
    + '<p><b>Running an audit:</b> use <b>Print blank form</b> to fill it out by hand, then '
    + '<b>Record audit results</b> to log the scores into KPI Data.</p>'
    + '<p style="background:#FEF9C3;padding:8px;border-radius:5px"><b>First time only:</b> the first '
    + 'tool you run will ask for permission &mdash; click <b>Advanced &rsaquo; Go to &hellip; &rsaquo; Allow</b>, '
    + 'then run the tool again. Each person does this once.</p>'
    + '</div>')
    .setWidth(470).setHeight(440);
  SpreadsheetApp.getUi().showModelessDialog(html, 'How to use the Hickory 6S Audit Forms');
}
