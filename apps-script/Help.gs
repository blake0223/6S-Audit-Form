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
  var s = 'font-family:Arial;padding:20px;font-size:13px;line-height:1.55;color:#111827';
  var step = 'margin:0 0 4px;color:#374151';
  var html = HtmlService.createHtmlOutput(
      '<div style="' + s + '">'
    + '<h2 style="margin:0 0 6px;font-size:18px">How to use the Hickory 6S Audit Forms</h2>'
    + '<p style="margin:0 0 14px">Find <b>your facility&rsquo;s tabs</b> at the bottom &mdash; they&rsquo;re color-coded. '
    + 'Each facility has a <b>Monthly Audit</b>, a <b>Weekly Checklist</b>, and a <b>Facility Summary</b>.</p>'

    + '<p style="margin:0 0 12px">Everything is run from the <b>6S Audit Tools</b> menu at the top. '
    + 'Each option opens a list of facilities &mdash; <b>just click your facility</b> (no typing required). '
    + 'You don&rsquo;t need to be on any particular tab first.</p>'

    + '<div style="border-top:1px solid #e5e7eb;margin:0 0 12px"></div>'

    + '<p style="margin:0 0 2px"><b>&#128205; Add Monthly Audit Location</b></p>'
    + '<p style="' + step + '">Click it &rarr; click your facility &rarr; type the location/area name '
    + '(e.g. &ldquo;Parts Room&rdquo;). A new scoring column is added to that facility&rsquo;s Monthly Audit. '
    + 'Add one for <b>every area</b> you audit.</p>'

    + '<p style="margin:10px 0 2px"><b>&#9745; Add checklist item</b></p>'
    + '<p style="' + step + '">Click it &rarr; click your facility &rarr; type the daily check '
    + '(e.g. &ldquo;Propane turned off&rdquo;). A new checkbox column is added to that facility&rsquo;s Weekly Checklist.</p>'

    + '<p style="margin:10px 0 2px"><b>&#128424; Print blank form</b></p>'
    + '<p style="' + step + '">Click it &rarr; choose <b>Monthly audit</b> or <b>Weekly checklist</b> &rarr; click your facility. '
    + 'A print-ready PDF opens in a new tab &mdash; use your browser&rsquo;s <b>Print</b> there. Fill it out by hand during the walk-through.</p>'

    + '<p style="margin:10px 0 2px"><b>&#128221; Record audit results</b></p>'
    + '<p style="' + step + '">First type the 0&ndash;3 scores straight into the location columns on the Monthly Audit. '
    + 'Then click this &rarr; click your facility. The scores are logged to the <b>KPI Data</b> tab and the entry '
    + 'cells are cleared for next time.</p>'

    + '<div style="border-top:1px solid #e5e7eb;margin:12px 0"></div>'

    + '<p style="margin:0 0 4px"><b>Typical order:</b></p>'
    + '<ol style="margin:0 0 12px 18px;padding:0;color:#374151">'
    + '<li>Add every location to your Monthly Audit, and any facility-specific checks to your Weekly Checklist.</li>'
    + '<li>Print the blank form and walk the facility, scoring 0&ndash;3.</li>'
    + '<li>Type the scores into the Monthly Audit columns, then Record audit results.</li>'
    + '</ol>'

    + '<p style="background:#FEF9C3;padding:9px;border-radius:5px;margin:0">'
    + '<b>First time only:</b> the first option you run asks for permission &mdash; click '
    + '<b>Advanced &rsaquo; Go to &hellip; &rsaquo; Allow</b>, then click the option again. Each person does this once.</p>'
    + '</div>')
    .setWidth(490).setHeight(560);
  SpreadsheetApp.getUi().showModelessDialog(html, 'How to use the Hickory 6S Audit Forms');
}
