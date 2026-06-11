/**
 * onOpen — menu trigger.
 *
 * Runs automatically whenever the spreadsheet is opened and builds the
 * "6S Audit Tools" menu. Each item calls a function defined in its own file:
 *   • addRoomColumn()      → AddRoomColumn.gs
 *   • rebuildRoomTotals()  → RoomTotals.gs
 *   • addSectionRow()      → SectionRows.gs
 *   • removeSectionRow()   → SectionRows.gs
 *   • printForm()          → PrintForm.gs
 *   • recordResults()      → RecordResults.gs
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('6S Audit Tools')
    .addItem('Add room column', 'addRoomColumn')
    .addItem('Rebuild room totals', 'rebuildRoomTotals')
    .addSeparator()
    .addItem('Add row to section', 'addSectionRow')
    .addItem('Remove selected row', 'removeSectionRow')
    .addSeparator()
    .addItem('Print blank form', 'printForm')
    .addItem('Record audit results', 'recordResults')
    .addToUi();
}
