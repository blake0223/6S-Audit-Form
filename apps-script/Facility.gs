/**
 * Facility — the facility picker shared by the menu tools. Uses NATIVE prompts
 * (not an HTML dialog), so the tools run entirely in the menu context — the same
 * context that can show Google's permission prompt and that always worked. HTML
 * dialogs call the server via google.script.run, which can't trigger consent and
 * was the source of the "Authorization is required" loop.
 *
 * Facility tabs are recognised by name (Monthly Audit vs Checklist); the label is
 * the tab name minus the form-type words (facilityLabel_ lives in PrintForm.gs).
 */
function listFacilityTabs_(type) {
  var rx = (type === 'checklist') ? /checklist/i : /monthly audit/i;
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(function (s) { return rx.test(s.getName()); })
    .map(function (s) { return { name: s.getName(), label: facilityLabel_(s.getName()) }; });
}

/** Native numbered picker over the facility tabs of a type. Returns tab name or null. */
function pickFacility_(type, title) {
  var ui = SpreadsheetApp.getUi();
  var tabs = listFacilityTabs_(type);
  if (!tabs.length) {
    ui.alert('No ' + (type === 'checklist' ? 'Weekly Checklist' : 'Monthly Audit') + ' tabs found.');
    return null;
  }
  if (tabs.length === 1) return tabs[0].name;
  var list = tabs.map(function (t, i) { return (i + 1) + ')  ' + t.label; }).join('\n');
  var resp = ui.prompt(title, 'Type the number of the facility:\n\n' + list, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return null;
  var n = parseInt(resp.getResponseText().trim(), 10);
  if (isNaN(n) || n < 1 || n > tabs.length) {
    ui.alert('Please enter a number between 1 and ' + tabs.length + '.');
    return null;
  }
  return tabs[n - 1].name;
}

/** Native numbered picker over a plain list of strings. Returns the choice or null. */
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
