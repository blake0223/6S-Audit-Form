/**
 * Facility — facility selection shared by the action tools, using NATIVE prompts.
 *
 * Why native: the action tools edit the sheet on the server. An HTML dropdown
 * dialog would have to call the server via google.script.run, which can't trigger
 * Google's permission prompt and fails with "Authorization is required." Native
 * prompts run in the menu context, which works. (Print keeps an HTML dropdown
 * because it never calls the server — the browser fetches the PDF.)
 *
 * Setup/master tabs (Safety Setup, Initiation, etc.) are excluded from the lists.
 */

var EXCLUDE_TABS = /initiation|safety setup|instruction|overview|kpi data/i;

function listFacilityTabs_(type) {
  var rx = (type === 'checklist') ? /checklist/i : /monthly audit/i;
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(function (s) { var n = s.getName(); return rx.test(n) && !EXCLUDE_TABS.test(n); })
    .map(function (s) { return { name: s.getName(), label: facilityLabel_(s.getName()) }; });
}

/** Native facility selection. Returns the chosen tab name, or null if cancelled. */
function pickFacility_(type, title) {
  var ui = SpreadsheetApp.getUi();
  var tabs = listFacilityTabs_(type);
  if (!tabs.length) {
    ui.alert('No ' + (type === 'checklist' ? 'Weekly Checklist' : 'Monthly Audit') + ' tabs found.');
    return null;
  }
  if (tabs.length === 1) return tabs[0].name;
  var list = tabs.map(function (t, i) { return (i + 1) + ')  ' + t.label; }).join('\n');
  var resp = ui.prompt(title + ' — choose facility',
    'Type the number of the facility:\n\n' + list, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return null;
  var n = parseInt(resp.getResponseText().trim(), 10);
  if (isNaN(n) || n < 1 || n > tabs.length) {
    ui.alert('Please enter a number between 1 and ' + tabs.length + '.');
    return null;
  }
  return tabs[n - 1].name;
}

/** Native selection over a plain list of strings. Returns the choice or null. */
function pickOption_(options, title) {
  var ui = SpreadsheetApp.getUi();
  var list = options.map(function (o, i) { return (i + 1) + ')  ' + o; }).join('\n');
  var resp = ui.prompt(title, 'Type the number:\n\n' + list, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return null;
  var n = parseInt(resp.getResponseText().trim(), 10);
  if (isNaN(n) || n < 1 || n > options.length) {
    ui.alert('Please enter a number between 1 and ' + options.length + '.');
    return null;
  }
  return options[n - 1];
}
