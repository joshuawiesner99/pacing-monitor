/**
 * Notify.js
 *
 * Sends the monitor results as an HTML email and, optionally, to Slack.
 *
 * Settings live in Script Properties (Project Settings -> Script Properties),
 * never in the code:
 *   RECIPIENT_EMAIL    optional, defaults to the script owner's address
 *   SLACK_WEBHOOK_URL  optional, Slack is skipped if not set
 */

/**
 * Daily job used by the time-driven trigger.
 * Demo mode: regenerates synthetic data first so the monitor always has
 * "fresh" data. In production, RawData would be filled by a Google Ads
 * export instead.
 */
function dailyRun() {
  generateDemoData();
  const report = runMonitor();
  sendReport_(report, false);
}

/** Menu action: analyse current data and always send a report. */
function sendReportNow() {
  const report = runMonitor();
  sendReport_(report, true);
}

/**
 * Sends the report. By default only when there are issues, so the
 * recipient is not flooded with "all good" messages.
 */
function sendReport_(report, force) {
  const issues = report.results.filter(
    (r) => r.pacingStatus !== PACING_STATUS.OK || r.anomaly
  );
  if (issues.length === 0 && !force) return;

  const props = PropertiesService.getScriptProperties();
  const recipient = props.getProperty('RECIPIENT_EMAIL') || Session.getEffectiveUser().getEmail();
  const webhook = props.getProperty('SLACK_WEBHOOK_URL');
  const dateLabel = dateKey_(report.latestDate);

  const subject =
    issues.length > 0
      ? `[Pacing Monitor] ${issues.length} issue(s) – data up to ${dateLabel}`
      : `[Pacing Monitor] All campaigns on track – data up to ${dateLabel}`;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: buildEmailHtml_(report, issues),
    body: buildPlainText_(report, issues),
  });

  if (webhook) {
    UrlFetchApp.fetch(webhook, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: buildSlackText_(report, issues) }),
      muteHttpExceptions: true,
    });
  }
}

// ---------------------------------------------------------------------------
// Message builders
// ---------------------------------------------------------------------------

function describeIssue_(r) {
  const parts = [];
  if (r.pacingStatus !== PACING_STATUS.OK) {
    parts.push(
      `${r.pacingStatus}: ${formatPercent_(r.budgetUsed)} of budget used after ` +
        `${formatPercent_(r.monthElapsed)} of the month`
    );
  }
  if (r.anomaly === 'CPL spike') {
    parts.push(
      `CPL spike: ${formatMoney_(r.cplToday)} vs. ${formatMoney_(r.cplBaseline)} ` +
        `7-day avg (${formatSignedPercent_(r.cplDeviation)})`
    );
  } else if (r.anomaly) {
    parts.push(r.anomaly);
  }
  return parts.join('; ');
}

function buildPlainText_(report, issues) {
  const lines = [`Pacing Monitor – data up to ${dateKey_(report.latestDate)}`, ''];
  if (issues.length === 0) {
    lines.push('All campaigns are on track. No anomalies detected.');
  } else {
    issues.forEach((r) => lines.push(`- ${r.campaign}: ${describeIssue_(r)}`));
  }
  lines.push('', 'Open the spreadsheet for the full overview:');
  lines.push(SpreadsheetApp.getActiveSpreadsheet().getUrl());
  return lines.join('\n');
}

function buildSlackText_(report, issues) {
  const header = `*Pacing Monitor* – data up to ${dateKey_(report.latestDate)}`;
  if (issues.length === 0) return `${header}\nAll campaigns are on track. :white_check_mark:`;
  const lines = issues.map((r) => `• *${r.campaign}*: ${describeIssue_(r)}`);
  return `${header}\n${lines.join('\n')}\n<${SpreadsheetApp.getActiveSpreadsheet().getUrl()}|Open spreadsheet>`;
}

function buildEmailHtml_(report, issues) {
  const cell = 'padding:6px 10px;border-bottom:1px solid #ddd;text-align:left;';
  const rows = report.results
    .map((r) => {
      const pacingColor =
        r.pacingStatus === PACING_STATUS.OVER
          ? STATUS_COLORS.alert
          : r.pacingStatus === PACING_STATUS.UNDER
          ? STATUS_COLORS.warn
          : STATUS_COLORS.ok;
      const anomalyColor = r.anomaly ? STATUS_COLORS.alert : STATUS_COLORS.ok;
      return (
        '<tr>' +
        `<td style="${cell}">${escapeHtml_(r.campaign)}</td>` +
        `<td style="${cell}">${formatPercent_(r.budgetUsed)} / ${formatPercent_(r.monthElapsed)}</td>` +
        `<td style="${cell}background:${pacingColor};">${r.pacingStatus}</td>` +
        `<td style="${cell}">${formatMoney_(r.cplToday)}</td>` +
        `<td style="${cell}">${formatMoney_(r.cplBaseline)}</td>` +
        `<td style="${cell}background:${anomalyColor};">${escapeHtml_(r.anomaly || '–')}</td>` +
        '</tr>'
      );
    })
    .join('');

  const summary =
    issues.length === 0
      ? '<p>All campaigns are on track. No anomalies detected.</p>'
      : '<ul>' +
        issues
          .map((r) => `<li><strong>${escapeHtml_(r.campaign)}</strong>: ${escapeHtml_(describeIssue_(r))}</li>`)
          .join('') +
        '</ul>';

  const th = `${cell}background:#f3f3f3;`;
  return (
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;">' +
    `<h2 style="margin:0 0 8px;">Pacing Monitor – data up to ${dateKey_(report.latestDate)}</h2>` +
    summary +
    '<table style="border-collapse:collapse;margin-top:12px;">' +
    '<tr>' +
    `<th style="${th}">Campaign</th><th style="${th}">Budget / Month</th>` +
    `<th style="${th}">Pacing</th><th style="${th}">CPL Today</th>` +
    `<th style="${th}">CPL 7-Day Avg</th><th style="${th}">Anomaly</th>` +
    '</tr>' +
    rows +
    '</table>' +
    `<p style="margin-top:16px;"><a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}">Open spreadsheet</a></p>` +
    '</div>'
  );
}

// ---------------------------------------------------------------------------
// Formatting helpers
// ---------------------------------------------------------------------------

function formatMoney_(value) {
  if (value === null || value === undefined) return 'n/a';
  const fixed = value.toFixed(2);
  const [whole, decimals] = fixed.split('.');
  return `€${whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',')}.${decimals}`;
}

function formatPercent_(value) {
  return `${(value * 100).toFixed(1)}%`;
}

function formatSignedPercent_(value) {
  if (value === null || value === undefined) return 'n/a';
  const sign = value >= 0 ? '+' : '';
  return `${sign}${(value * 100).toFixed(1)}%`;
}

function escapeHtml_(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
