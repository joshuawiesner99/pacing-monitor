/**
 * Menu.js
 *
 * Adds a "Pacing Monitor" menu to the spreadsheet so all functions can be
 * run directly from the sheet without opening the script editor.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pacing Monitor (deploy test)')
    .addItem('Generate demo data', 'generateDemoData')
    .addItem('Run monitor', 'runMonitor')
    .addSeparator()
    .addItem('Send report now', 'sendReportNow')
    .addSeparator()
    .addItem('Schedule daily run', 'installDailyTrigger')
    .addItem('Remove daily run', 'removeDailyTrigger')
    .addToUi();
}
