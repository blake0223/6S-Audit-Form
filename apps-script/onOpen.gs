/**
 * onOpen — menu trigger.
 *
 * Builds the "6S Audit Tools" menu when the spreadsheet opens. Each tool prompts
 * you to pick a facility first (except "Remove selected row", which acts on the
 * row your cursor is in). Functions live in their own files:
 *   • showHelp()         → Help.gs
 *   • addLocation()      → AddRoomColumn.gs   (facility picker)
 *   • addGradingItem()   → SectionRows.gs     (facility picker)
 *   • addChecklistItem() → Checklist.gs       (facility picker)
 *   • printForm()        → PrintForm.gs       (facility picker)
 *   • recordAudit()      → RecordResults.gs   (facility picker)
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('6S Audit Tools')
    .addItem('How to use this sheet', 'showHelp')
    .addSeparator()
    .addSubMenu(ui.createMenu('Monthly audit')
      .addItem('Add location', 'addLocation')
      .addItem('Add grading item', 'addGradingItem'))
    .addSubMenu(ui.createMenu('Weekly checklist')
      .addItem('Add checklist item', 'addChecklistItem'))
    .addSeparator()
    .addItem('Print blank form', 'printForm')
    .addItem('Record audit results', 'recordAudit')
    .addToUi();
}
