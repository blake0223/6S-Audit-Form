/**
 * Facility — list the facility tabs of a given type for the menu.
 *
 * Selection happens entirely through native nested menus (built in onOpen.gs):
 * clicking "Add location ▸ Bell" runs in the menu context, where Apps Script can
 * prompt for permission if needed. We deliberately avoid HtmlService +
 * google.script.run for the actions, because that path cannot show the permission
 * prompt and fails with "Authorization is required."
 *
 * Setup/master tabs (Safety Setup, Initiation, etc.) are excluded.
 */

var EXCLUDE_TABS = /initiation|safety setup|instruction|overview|kpi data/i;

/** Facility tabs of a type, as [{ name, label }], in sheet order. */
function listFacilityTabs_(type) {
  var rx = (type === 'checklist') ? /checklist/i : /monthly audit/i;
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(function (s) { var n = s.getName(); return rx.test(n) && !EXCLUDE_TABS.test(n); })
    .map(function (s) { return { name: s.getName(), label: facilityLabel_(s.getName()) }; });
}
