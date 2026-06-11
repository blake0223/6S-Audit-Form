/**
 * onOpen — menu trigger.
 *
 * Runs automatically whenever the spreadsheet is opened and builds the
 * "6S Audit Tools" menu. Each item calls a function defined in its own file:
 *   • addRoomColumn()        → AddRoomColumn.gs
 *   • downloadPrintablePdf() → DownloadPrintablePdf.gs
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('6S Audit Tools')
    .addItem('➕  Add room column', 'addRoomColumn')
    .addSeparator()
    .addItem('🖨️  Download printable PDF', 'downloadPrintablePdf')
    .addToUi();
}
