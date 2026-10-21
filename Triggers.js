/**
 * Triggers.js
 *
 * Installs or removes the daily time-driven trigger that runs dailyRun().
 * Existing triggers for dailyRun are removed first, so running the install
 * action twice never creates duplicate daily runs.
 */

const DAILY_TRIGGER_FUNCTION = 'dailyRun';
const DAILY_TRIGGER_HOUR = 7; // between 7 and 8 am in the project time zone

function installDailyTrigger() {
  removeDailyTrigger_();
  ScriptApp.newTrigger(DAILY_TRIGGER_FUNCTION)
    .timeBased()
    .everyDays(1)
    .atHour(DAILY_TRIGGER_HOUR)
    .create();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Daily run scheduled between ${DAILY_TRIGGER_HOUR}:00 and ${DAILY_TRIGGER_HOUR + 1}:00.`,
    'Pacing Monitor'
  );
}

function removeDailyTrigger() {
  const removed = removeDailyTrigger_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    removed > 0 ? 'Daily run removed.' : 'No daily run was scheduled.',
    'Pacing Monitor'
  );
}

function removeDailyTrigger_() {
  const triggers = ScriptApp.getProjectTriggers().filter(
    (t) => t.getHandlerFunction() === DAILY_TRIGGER_FUNCTION
  );
  triggers.forEach((t) => ScriptApp.deleteTrigger(t));
  return triggers.length;
}
