// Data is appended to this sheet/tab.
var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1sxGFbLbdVFq3Ls8jjyEmQzq5jdH6hmZGHj0c5-HwsbA/edit?usp=sharing';
var SHEET_NAME = 'Sheet2';

// Use the GitHub RAW file URL containing one account ID per line.
var GITHUB_RAW_URL = 'https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main/account_ids.txt';

var TARGET_LABEL = '.DGO_Winsome';
var DISPLAY_TIMEZONE = 'GMT+8';
var REPORT_TIMEZONE = 'PST';
var DAY_OFFSET = 1; // 1 = yesterday, 2 = two days ago, etc.

function main() {
  var sheet = getOrCreateSheet();
  ensureHeaders(sheet);

  var accountIds = fetchAccountIdsFromGithub(GITHUB_RAW_URL);
  if (accountIds.length === 0) {
    Logger.log('No valid account IDs found. Stopping script.');
    return;
  }

  var dates = getReportDates(DAY_OFFSET);
  var rowsToAppend = collectRows(accountIds, dates.queryDate, dates.dayFormatted);

  if (rowsToAppend.length === 0) {
    Logger.log('No data to append for ' + dates.dayFormatted + '.');
    return;
  }

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  Logger.log('Successfully appended ' + rowsToAppend.length + ' rows to the spreadsheet.');
}

function getOrCreateSheet() {
  var spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    Logger.log('Created missing sheet: ' + SHEET_NAME);
  }
  return sheet;
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() !== 0) {
    return;
  }

  sheet.appendRow([
    'Day',
    'Account',
    'Customer ID',
    'Account Labels',
    'Clicks',
    'Impr.',
    'CTR',
    'Avg. CPC',
    'Cost',
    'Video Views'
  ]);
}

function fetchAccountIdsFromGithub(url) {
  Logger.log('Fetching account IDs from GitHub...');
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var statusCode = response.getResponseCode();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Failed to fetch account IDs. HTTP ' + statusCode + '. URL: ' + url);
  }

  var fileContent = response.getContentText() || '';
  var lines = fileContent.split(/\r?\n/);
  var ids = [];
  var seen = {};

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();

    // Allow comments and empty lines in account_ids.txt.
    if (!line || line.indexOf('#') === 0 || line.indexOf('//') === 0) {
      continue;
    }

    var normalized = normalizeCustomerId(line);
    if (!normalized) {
      Logger.log('Skipping invalid account ID line: ' + line);
      continue;
    }

    if (!seen[normalized]) {
      seen[normalized] = true;
      ids.push(normalized);
    }
  }

  Logger.log('Loaded ' + ids.length + ' valid account IDs from GitHub.');
  return ids;
}

function normalizeCustomerId(value) {
  var digits = String(value).replace(/\D/g, '');
  if (digits.length !== 10) {
    return null;
  }
  return digits;
}

function getReportDates(dayOffset) {
  var now = new Date();
  var targetDate = new Date(now.getTime() - (dayOffset * 24 * 60 * 60 * 1000));

  return {
    dayFormatted: Utilities.formatDate(targetDate, DISPLAY_TIMEZONE, 'yyyy-MM-dd'),
    queryDate: Utilities.formatDate(targetDate, REPORT_TIMEZONE, 'yyyyMMdd')
  };
}

function collectRows(accountIds, queryDate, dayFormatted) {
  var rowsToAppend = [];

  var accountIterator = AdsManagerApp.accounts()
    .withIds(accountIds)
    .withCondition("customer_client.status = 'ENABLED'")
    .get();

  while (accountIterator.hasNext()) {
    var account = accountIterator.next();
    AdsManagerApp.select(account);

    if (!accountHasTargetLabel(account, TARGET_LABEL)) {
      Logger.log('Account ' + account.getCustomerId() + ' skipped. Missing label: ' + TARGET_LABEL);
      continue;
    }

    var accountName = account.getName();
    var accountId = account.getCustomerId();
    var report = AdsApp.report(
      "SELECT Clicks, Impressions, Ctr, AverageCpc, Cost, VideoViews " +
      "FROM ACCOUNT_PERFORMANCE_REPORT " +
      "WHERE Date = '" + queryDate + "'"
    );

    var reportRows = report.rows();
    while (reportRows.hasNext()) {
      var row = reportRows.next();
      var cost = parseCost(row.Cost);

      if (cost < 0) {
        continue;
      }

      rowsToAppend.push([
        dayFormatted,
        accountName,
        accountId,
        TARGET_LABEL,
        row.Clicks,
        row.Impressions,
        row.Ctr,
        row.AverageCpc,
        row.Cost,
        row.VideoViews
      ]);
    }
  }

  return rowsToAppend;
}

function accountHasTargetLabel(account, labelName) {
  var mccLabelIterator = account.labels().get();
  while (mccLabelIterator.hasNext()) {
    if (mccLabelIterator.next().getName() === labelName) {
      return true;
    }
  }

  var accountLabelIterator = AdsApp.labels().get();
  while (accountLabelIterator.hasNext()) {
    if (accountLabelIterator.next().getName() === labelName) {
      return true;
    }
  }

  return false;
}

function parseCost(costValue) {
  var text = costValue ? String(costValue).replace(/,/g, '') : '0';
  var cost = parseFloat(text);
  return isNaN(cost) ? 0 : cost;
}