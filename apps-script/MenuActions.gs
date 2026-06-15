/**
 * MenuActions — dispatchers the generated menu wrappers (MenuWrappers.gs) call.
 * Everything here runs in the menu context (clicked submenu item), so permission
 * prompts work and there is no google.script.run.
 *
 * Facility is chosen by the submenu the user clicked (index into the tab list).
 * The item name (a free-text value, not a list choice) is collected with a native
 * prompt, which is fine — the thing that had to become clickable was the facility
 * and category selection, and those are now menu items.
 */

var GRADING_CATEGORIES = ['SORT', 'SET IN ORDER', 'SHINE', 'STANDARDIZE', 'SUSTAIN', 'SAFETY'];

/** "SET IN ORDER" -> "Set in Order" for menu labels. */
function categoryLabel_(c) {
  return c.toLowerCase().replace(/(^|\s)\S/g, function (m) { return m.toUpperCase(); });
}

/** Facility-only tools: 'loc' | 'chk' | 'rec' | 'pm' | 'pw'. */
function runFacilityTool_(tool, type, idx) {
  var ui = SpreadsheetApp.getUi();
  var f = listFacilityTabs_(type)[idx];
  if (!f) { ui.alert('That facility is no longer available — reopen the sheet to refresh the menu.'); return; }
  try {
    if (tool === 'loc') {
      var r = ui.prompt('Add location — ' + f.label, 'Location / area name (e.g. Parts Room):', ui.ButtonSet.OK_CANCEL);
      if (r.getSelectedButton() !== ui.Button.OK) return;
      var v = r.getResponseText().trim();
      if (!v) { ui.alert('No location name entered.'); return; }
      ui.alert(addLocationFor(f.name, v));
    } else if (tool === 'chk') {
      var r2 = ui.prompt('Add checklist item — ' + f.label, 'Checklist item (e.g. Propane turned off):', ui.ButtonSet.OK_CANCEL);
      if (r2.getSelectedButton() !== ui.Button.OK) return;
      var v2 = r2.getResponseText().trim();
      if (!v2) { ui.alert('No item entered.'); return; }
      ui.alert(addChecklistItemFor(f.name, v2));
    } else if (tool === 'rec') {
      ui.alert(recordResultsFor(f.name));
    } else if (tool === 'pm') {
      openPrintDialog_(f.name, 'monthly');
    } else if (tool === 'pw') {
      openPrintDialog_(f.name, 'checklist');
    }
  } catch (e) { ui.alert('Error: ' + e.message); }
}

/** Add grading item: facility index + 6S category index, then prompt for the name. */
function runGradingTool_(fIdx, cIdx) {
  var ui = SpreadsheetApp.getUi();
  var f = listFacilityTabs_('monthly')[fIdx];
  var cat = GRADING_CATEGORIES[cIdx];
  if (!f || !cat) { ui.alert('That option is no longer available — reopen the sheet to refresh the menu.'); return; }
  try {
    var r = ui.prompt('Add grading item — ' + f.label + ' / ' + categoryLabel_(cat),
      'Check item name (optional — leave blank to fill in later):', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    ui.alert(addGradingItemFor(f.name, cat, r.getResponseText().trim()));
  } catch (e) { ui.alert('Error: ' + e.message); }
}

/** Build the print URL (trimmed copy for monthly) and show a clickable PDF link. */
function openPrintDialog_(sheetName, type) {
  var url = getPrintUrl_(sheetName, type);
  var html = HtmlService.createHtmlOutput(
      '<body style="font-family:Arial;margin:0;padding:22px 20px;text-align:center;color:#374151">'
    + '<p style="font-size:13px;margin:0 0 16px">Your print-ready form is ready.</p>'
    + '<a href="' + url + '" target="_blank" rel="noopener" '
    + 'style="display:inline-block;background:#1F4E79;color:#fff;text-decoration:none;'
    + 'padding:11px 20px;border-radius:6px;font-size:14px;font-weight:bold">Open PDF in new tab</a>'
    + '<p style="font-size:11px;color:#9ca3af;margin:16px 0 0">Then use your browser’s Print or Download.</p>'
    + '</body>')
    .setWidth(300).setHeight(160);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Print blank form');
}
