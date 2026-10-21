/**
 * Menu.js
 *
 * Adds a "Pacing Monitor" menu to the spreadsheet so all functions can be
 * run directly from the sheet without opening the script editor.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pacing Monitor')
    .addItem('Generate demo data', 'generateDemoData')
    .addToUi();
}
