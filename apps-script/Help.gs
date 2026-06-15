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
  var s = 'font-family:Arial;padding:20px;font-size:13px;line-height:1.5;color:#111827';
  var name = 'margin:12px 0 1px;font-size:13px';
  var desc = 'margin:0;color:#374151';
  var html = HtmlService.createHtmlOutput(
      '<div style="' + s + '">'
    + '<h2 style="margin:0 0 8px;font-size:18px">How to use the 6S Audit Forms</h2>'
    + '<p style="margin:0 0 10px">Each facility has three tabs: <b>Monthly Audit</b>, <b>Weekly Checklist</b>, '
    + 'and <b>Facility Summary</b>.</p>'
    + '<p style="margin:0 0 6px">Use the <b>6S Audit Tools</b> menu above. Every option opens a list of '
    + 'facilities &mdash; click yours. You don&rsquo;t need to be on a specific tab.</p>'

    + '<div style="border-top:1px solid #e5e7eb;margin:10px 0 0"></div>'

    + '<p style="' + name + '"><b>Add Monthly Audit Location</b></p>'
    + '<p style="' + desc + '">Pick facility, type an area name (e.g. Parts Room). Adds a scoring column to the Monthly Audit. Add one per area.</p>'

    + '<p style="' + name + '"><b>Add Weekly Checklist Item</b></p>'
    + '<p style="' + desc + '">Pick facility, type a daily check (e.g. Propane turned off). Adds a checkbox column to the Weekly Checklist.</p>'

    + '<p style="' + name + '"><b>Print Audit Form</b></p>'
    + '<p style="' + desc + '">Choose Monthly or Weekly, pick facility. A print-ready PDF opens in a new tab &mdash; print it from there.</p>'

    + '<p style="' + name + '"><b>Record Audit Results</b></p>'
    + '<p style="' + desc + '">Type the 0&ndash;3 scores into the Monthly Audit columns first, then pick facility. Scores are saved to the <b>KPI Data</b> tab and the cells are cleared.</p>'

    + '<div style="border-top:1px solid #e5e7eb;margin:14px 0 8px"></div>'

    + '<p style="margin:0 0 4px"><b>Typical order</b></p>'
    + '<ol style="margin:0 0 12px 18px;padding:0;color:#374151">'
    + '<li>Add your locations and any checklist items.</li>'
    + '<li>Print the form and score the facility 0&ndash;3 by hand.</li>'
    + '<li>Enter the scores in the Monthly Audit, then Record Audit Results.</li>'
    + '</ol>'

    + '<p style="background:#FEF9C3;padding:9px;border-radius:5px;margin:0">'
    + '<b>First time only:</b> the first option asks for permission &mdash; click '
    + '<b>Advanced &rsaquo; Go to &hellip; &rsaquo; Allow</b>, then run it again.</p>'
    + '</div>')
    .setWidth(470).setHeight(500);
  SpreadsheetApp.getUi().showModelessDialog(html, 'How to use the 6S Audit Forms');
}
