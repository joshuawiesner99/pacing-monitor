/**
 * DemoData.js
 *
 * Generates reproducible, synthetic Google Ads-like performance data for a
 * fictional client and writes it to the RawData sheet. A few anomalies are
 * injected on purpose so the monitor has something to detect.
 *
 * No real client data is used anywhere in this project.
 */

const SHEETS = {
  CONFIG: 'Config',
  RAW: 'RawData',
  RESULTS: 'Results',
};

// Fixed seed -> the same calendar date always produces the same data.
const DEMO_SEED = 42;

// Fictional performance profile per campaign type.
const CAMPAIGN_PROFILES = {
  Brand:       { cpc: 0.9, ctr: 0.12,  cvr: 0.09 },
  Generic:     { cpc: 2.6, ctr: 0.04,  cvr: 0.035 },
  Retargeting: { cpc: 1.1, ctr: 0.008, cvr: 0.05 },
};
const DEFAULT_PROFILE = { cpc: 1.5, ctr: 0.03, cvr: 0.04 };

// Deliberately injected anomalies the monitor should detect later.
const DEMO_ANOMALIES = [
  // CPL spike: conversion rate collapses on the most recent day.
  { campaign: 'Generic', type: 'cplSpike', conversionFactor: 0.3 },
  // Overpacing: spend runs ~45% above plan for the whole current month.
  { campaign: 'Retargeting', type: 'overpacing', spendFactor: 1.45 },
];

/**
 * Entry point: builds daily rows from the first day of the previous month
 * up to yesterday (the latest complete day) and writes them to RawData.
 */
function generateDemoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const campaigns = readConfig_(ss);
  if (campaigns.length === 0) {
    throw new Error('No active campaigns found in the Config sheet.');
  }

  const rand = createRandom_(DEMO_SEED);
  const endDate = addDays_(startOfDay_(new Date()), -1);
  const startDate = new Date(endDate.getFullYear(), endDate.getMonth() - 1, 1);

  const rows = [];
  for (let d = new Date(startDate); d <= endDate; d = addDays_(d, 1)) {
    const isCurrentMonth =
      d.getMonth() === endDate.getMonth() && d.getFullYear() === endDate.getFullYear();
    const isLastDay = d.getTime() === endDate.getTime();
    const daysInMonth = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
    const isWeekend = d.getDay() === 0 || d.getDay() === 6;
    const weekdayFactor = isWeekend ? 0.8 : 1.05;

    campaigns.forEach((campaign) => {
      const profile = CAMPAIGN_PROFILES[campaign.name] || DEFAULT_PROFILE;

      let cost = (campaign.monthlyBudget / daysInMonth) * weekdayFactor * randomBetween_(rand, 0.85, 1.15);
      let cvr = profile.cvr * randomBetween_(rand, 0.9, 1.1);

      DEMO_ANOMALIES.forEach((anomaly) => {
        if (anomaly.campaign !== campaign.name) return;
        if (anomaly.type === 'overpacing' && isCurrentMonth) cost *= anomaly.spendFactor;
        if (anomaly.type === 'cplSpike' && isLastDay) cvr *= anomaly.conversionFactor;
      });

      const cpc = profile.cpc * randomBetween_(rand, 0.9, 1.1);
      const clicks = Math.max(1, Math.round(cost / cpc));
      const ctr = profile.ctr * randomBetween_(rand, 0.85, 1.15);
      const impressions = Math.round(clicks / ctr);
      // Fractional conversions are realistic with data-driven attribution.
      const conversions = Math.round(clicks * cvr * 10) / 10;

      rows.push([new Date(d), campaign.name, round2_(cost), clicks, impressions, conversions]);
    });
  }

  writeRawData_(ss, rows);
  ss.toast(`${rows.length} rows of demo data generated.`, 'Pacing Monitor');
}

/**
 * Reads active campaigns from the Config sheet. Columns are looked up by
 * header name, so their order in the sheet does not matter.
 */
function readConfig_(ss) {
  const sheet = getSheetOrThrow_(ss, SHEETS.CONFIG);
  const values = sheet.getDataRange().getValues();
  const header = values.shift();

  const col = (name) => {
    const index = header.indexOf(name);
    if (index === -1) throw new Error(`Column "${name}" not found in sheet "${SHEETS.CONFIG}".`);
    return index;
  };
  const iName = col('Campaign');
  const iBudget = col('Monthly Budget');
  const iThreshold = col('CPL Threshold (%)');
  const iActive = col('Active');

  return values
    .filter((row) => row[iName] !== '' && isTruthy_(row[iActive]))
    .map((row) => ({
      name: String(row[iName]).trim(),
      monthlyBudget: Number(row[iBudget]),
      cplThreshold: Number(row[iThreshold]),
    }));
}

/** Replaces the content of RawData with a header row plus the given rows. */
function writeRawData_(ss, rows) {
  const sheet = getSheetOrThrow_(ss, SHEETS.RAW);
  const header = ['Date', 'Campaign', 'Cost', 'Clicks', 'Impressions', 'Conversions'];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);

  if (rows.length > 0) {
    // One batched write instead of cell-by-cell updates.
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat('#,##0.00');
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getSheetOrThrow_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet "${name}" not found.`);
  return sheet;
}

/** Accepts checkbox booleans as well as typed TRUE/WAHR values. */
function isTruthy_(value) {
  if (value === true) return true;
  const text = String(value).trim().toUpperCase();
  return text === 'TRUE' || text === 'WAHR';
}

/** Seeded pseudo-random generator (mulberry32) for reproducible demo data. */
function createRandom_(seed) {
  let state = seed >>> 0;
  return function () {
    state = (state + 0x6d2b79f5) | 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function randomBetween_(rand, min, max) {
  return min + rand() * (max - min);
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function addDays_(date, days) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);
}

function round2_(value) {
  return Math.round(value * 100) / 100;
}
