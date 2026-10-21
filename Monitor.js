/**
 * Monitor.js
 *
 * Core analysis of the Pacing Monitor:
 *  - Pacing: compares month-to-date spend with the monthly budget relative
 *    to how much of the month has passed.
 *  - Anomalies: compares the latest day's CPL with the average CPL of the
 *    days before it and flags deviations above the campaign's threshold.
 * Results are written to the Results sheet with colour-coded status cells.
 */

// Spend-to-time ratio outside 1 ± tolerance is flagged (0.15 = ±15%).
const PACING_TOLERANCE = 0.15;
// Number of days before the latest day used as the CPL baseline.
const BASELINE_DAYS = 7;

const PACING_STATUS = {
  OVER: 'Overpacing',
  UNDER: 'Underpacing',
  OK: 'On track',
};

const STATUS_COLORS = {
  alert: '#f4c7c3',
  warn: '#fce8b2',
  ok: '#b7e1cd',
};

const RESULTS_HEADER = [
  'Campaign',
  'Monthly Budget',
  'Spend MTD',
  'Budget Used (%)',
  'Month Elapsed (%)',
  'Pacing Status',
  'CPL Today',
  'CPL 7-Day Avg',
  'CPL Deviation (%)',
  'Anomaly',
];

/**
 * Entry point: analyses all active campaigns and writes the Results sheet.
 * Returns the results so later steps (e.g. notifications) can reuse them.
 */
function runMonitor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const campaigns = readConfig_(ss);
  const rawRows = readRawData_(ss);
  if (rawRows.length === 0) {
    throw new Error('RawData is empty. Run "Generate demo data" first.');
  }

  const latestDate = rawRows.reduce(
    (max, row) => (row.date > max ? row.date : max),
    rawRows[0].date
  );
  const results = campaigns.map((campaign) => analyzeCampaign_(campaign, rawRows, latestDate));

  writeResults_(ss, results);
  ss.toast(
    `Data up to ${dateKey_(latestDate)} analysed. ${countIssues_(results)} issue(s) found.`,
    'Pacing Monitor'
  );
  return { latestDate: latestDate, results: results };
}

/** Reads RawData into plain objects; rows without a valid date are skipped. */
function readRawData_(ss) {
  const values = getSheetOrThrow_(ss, SHEETS.RAW).getDataRange().getValues();
  values.shift();
  return values
    .filter((row) => row[0] instanceof Date && row[1] !== '')
    .map((row) => ({
      date: startOfDay_(row[0]),
      campaign: String(row[1]).trim(),
      cost: Number(row[2]),
      conversions: Number(row[5]),
    }));
}

/** Calculates pacing and CPL anomaly for one campaign. */
function analyzeCampaign_(campaign, rawRows, latestDate) {
  const rows = rawRows.filter((row) => row.campaign === campaign.name);

  // --- Pacing ---------------------------------------------------------------
  const monthRows = rows.filter(
    (row) =>
      row.date.getFullYear() === latestDate.getFullYear() &&
      row.date.getMonth() === latestDate.getMonth()
  );
  const spendMtd = sum_(monthRows, 'cost');
  const daysInMonth = new Date(latestDate.getFullYear(), latestDate.getMonth() + 1, 0).getDate();
  const monthElapsed = latestDate.getDate() / daysInMonth;
  const budgetUsed = campaign.monthlyBudget > 0 ? spendMtd / campaign.monthlyBudget : 0;
  const pacingRatio = budgetUsed / monthElapsed;

  let pacingStatus = PACING_STATUS.OK;
  if (pacingRatio > 1 + PACING_TOLERANCE) pacingStatus = PACING_STATUS.OVER;
  else if (pacingRatio < 1 - PACING_TOLERANCE) pacingStatus = PACING_STATUS.UNDER;

  // --- CPL anomaly ----------------------------------------------------------
  const latestKey = dateKey_(latestDate);
  const baselineStart = addDays_(latestDate, -BASELINE_DAYS);
  const todayRows = rows.filter((row) => dateKey_(row.date) === latestKey);
  const baselineRows = rows.filter((row) => row.date >= baselineStart && row.date < latestDate);

  const cplToday = cpl_(todayRows);
  const cplBaseline = cpl_(baselineRows);

  let cplDeviation = null;
  let anomaly = '';
  if (cplToday === null && sum_(todayRows, 'cost') > 0) {
    anomaly = 'No conversions';
  } else if (cplToday !== null && cplBaseline !== null) {
    cplDeviation = cplToday / cplBaseline - 1;
    // Only an increase in CPL is a problem; a cheaper CPL is not flagged.
    if (cplDeviation * 100 > campaign.cplThreshold) anomaly = 'CPL spike';
  }

  return {
    campaign: campaign.name,
    monthlyBudget: campaign.monthlyBudget,
    spendMtd: spendMtd,
    budgetUsed: budgetUsed,
    monthElapsed: monthElapsed,
    pacingStatus: pacingStatus,
    cplToday: cplToday,
    cplBaseline: cplBaseline,
    cplDeviation: cplDeviation,
    anomaly: anomaly,
  };
}

/** Rewrites the Results sheet: values in one batch, then formats and colours. */
function writeResults_(ss, results) {
  const sheet = getSheetOrThrow_(ss, SHEETS.RESULTS);
  sheet.clear();

  sheet.getRange(1, 1, 1, RESULTS_HEADER.length).setValues([RESULTS_HEADER]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (results.length === 0) return;

  const rows = results.map((r) => [
    r.campaign,
    r.monthlyBudget,
    round2_(r.spendMtd),
    r.budgetUsed,
    r.monthElapsed,
    r.pacingStatus,
    valueOrNa_(r.cplToday),
    valueOrNa_(r.cplBaseline),
    valueOrNa_(r.cplDeviation),
    r.anomaly || '–',
  ]);
  const n = rows.length;
  sheet.getRange(2, 1, n, RESULTS_HEADER.length).setValues(rows);

  // Number formats: currency-like columns and percentage columns.
  sheet.getRange(2, 2, n, 2).setNumberFormat('#,##0.00');
  sheet.getRange(2, 4, n, 2).setNumberFormat('0.0%');
  sheet.getRange(2, 7, n, 2).setNumberFormat('#,##0.00');
  sheet.getRange(2, 9, n, 1).setNumberFormat('+0.0%;-0.0%');

  // Colour-coded status cells.
  const pacingColors = results.map((r) => [
    r.pacingStatus === PACING_STATUS.OVER
      ? STATUS_COLORS.alert
      : r.pacingStatus === PACING_STATUS.UNDER
      ? STATUS_COLORS.warn
      : STATUS_COLORS.ok,
  ]);
  const anomalyColors = results.map((r) => [r.anomaly ? STATUS_COLORS.alert : STATUS_COLORS.ok]);
  sheet.getRange(2, 6, n, 1).setBackgrounds(pacingColors);
  sheet.getRange(2, 10, n, 1).setBackgrounds(anomalyColors);

  sheet.autoResizeColumns(1, RESULTS_HEADER.length);
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function sum_(rows, field) {
  return rows.reduce((total, row) => total + row[field], 0);
}

/** Cost per lead for a set of rows, or null if there are no conversions. */
function cpl_(rows) {
  const conversions = sum_(rows, 'conversions');
  return conversions > 0 ? sum_(rows, 'cost') / conversions : null;
}

function countIssues_(results) {
  return results.filter((r) => r.pacingStatus !== PACING_STATUS.OK || r.anomaly).length;
}

function valueOrNa_(value) {
  return value === null || value === undefined ? 'n/a' : value;
}

function dateKey_(date) {
  const pad = (x) => String(x).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}
