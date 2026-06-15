/**
 * onOpen — builds the "6S Audit Tools" menu with native nested submenus, one item
 * per facility (and per 6S category for grading). Clicking a submenu item runs in
 * the menu context, so it authorizes cleanly — no google.script.run, no auth wall.
 *
 * Submenu items call the generated wrappers in MenuWrappers.gs (t_loc_0, t_rec_1,
 * …), which forward to the dispatchers in MenuActions.gs. The wrapper index matches
 * the facility's position in listFacilityTabs_(type), which is stable (sheet order).
 *
 * onOpen is a simple trigger: it may read the bound spreadsheet's tab names (used
 * here) without extra authorization.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var monthly = listFacilityTabs_('monthly');
  var checklist = listFacilityTabs_('checklist');

  var locM = ui.createMenu('Add Monthly Audit Location');
  monthly.forEach(function (f, i) { locM.addItem(f.label, 't_loc_' + i); });

  var chkM = ui.createMenu('Add Weekly Checklist Item');
  checklist.forEach(function (f, i) { chkM.addItem(f.label, 't_chk_' + i); });

  var printM = ui.createMenu('Print Audit Form');
  var pmSub = ui.createMenu('Monthly audit');
  monthly.forEach(function (f, i) { pmSub.addItem(f.label, 't_pm_' + i); });
  var pwSub = ui.createMenu('Weekly checklist');
  checklist.forEach(function (f, i) { pwSub.addItem(f.label, 't_pw_' + i); });
  printM.addSubMenu(pmSub).addSubMenu(pwSub);

  var recM = ui.createMenu('Record Audit Results');
  monthly.forEach(function (f, i) { recM.addItem(f.label, 't_rec_' + i); });

  ui.createMenu('6S Audit Tools')
    .addItem('How to use this sheet', 'showHelp')
    .addSeparator()
    .addSubMenu(locM)
    .addSubMenu(chkM)
    .addSeparator()
    .addSubMenu(printM)
    .addSubMenu(recM)
    .addToUi();
}
