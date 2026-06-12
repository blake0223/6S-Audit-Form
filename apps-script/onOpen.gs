/**
 * onOpen — menu trigger.
 *
 * Builds the "6S Audit Tools" menu when the spreadsheet opens. Grouped so it's
 * clear what configures the Monthly Audit vs. the Weekly Checklist. Each item
 * calls a function in its own file:
 *   • showHelp()          → Help.gs
 *   • addRoomColumn()     → AddRoomColumn.gs
 *   • addSectionRow()     → SectionRows.gs
 *   • removeSectionRow()  → SectionRows.gs
 *   • rebuildRoomTotals() → RoomTotals.gs
 *   • addChecklistItem()  → Checklist.gs
 *   • printForm()         → PrintForm.gs
 *   • recordResults()     → RecordResults.gs
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('6S Audit Tools')
    .addItem('How to use this sheet', 'showHelp')
    .addSeparator()
    .addSubMenu(ui.createMenu('Monthly audit')
      .addItem('Add room (column)', 'addRoomColumn')
      .addItem('Add grading item (row)', 'addSectionRow')
      .addItem('Remove selected row', 'removeSectionRow')
      .addItem('Rebuild totals', 'rebuildRoomTotals'))
    .addSubMenu(ui.createMenu('Weekly checklist')
      .addItem('Add checklist item', 'addChecklistItem'))
    .addSeparator()
    .addItem('Print blank form', 'printForm')
    .addItem('Record audit results', 'recordResults')
    .addToUi();
}
