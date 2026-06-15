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

/** Build the print URL (trimmed copy for monthly) and open the PDF automatically. */
function openPrintDialog_(sheetName, type) {
  var url = getPrintUrl_(sheetName, type);
  var html = HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><base target="_blank"></head>'
    + '<body style="font-family:Arial;margin:0;padding:20px;text-align:center;color:#374151">'
    + '<p style="font-size:13px;margin:0 0 14px">Opening your print-ready PDF…</p>'
    + '<a href="' + url + '" target="_blank" rel="noopener" '
    + 'style="display:inline-block;background:#1F4E79;color:#fff;text-decoration:none;'
    + 'padding:11px 20px;border-radius:6px;font-size:14px;font-weight:bold">If it didn’t open, click here</a>'
    + '<script>window.open(' + JSON.stringify(url) + ',"_blank");</script>'
    + '</body></html>')
    .setWidth(300).setHeight(150);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Print Audit Form');
}
