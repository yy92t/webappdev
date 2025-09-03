/**
 * @OnlyCurrentDoc
 * This script creates an auto-refreshing dashboard with four charts and a summary pivot table.
 * This version is optimized for performance and conciseness.
 */

// --- CONFIGURATION ---
const SOURCE_SHEET_NAME = "Weekly log_Thomas W";
const DASHBOARD_SHEET_NAME = "Full Dashboard";
const PIVOT_SHEET_NAME = "Summary Pivot Table";

// Column numbers (index-based for array access, e.g., Column A is 0)
const CLIENT_COL = 6, PLATFORM_COL = 7, AD_FORMAT_COL = 8, META_BUDGET_COL = 9;
const CAMPAIGN_TYPE_COL = 12, FREQUENCY_COL = 18, START_DATE_COL = 19, END_DATE_COL = 20;
// --- END CONFIGURATION ---

// --- PERFORMANCE / CONTROL CONFIG ---
const THROTTLE_MS = 30 * 1000;          // 最短重建間隔
const CACHE_TTL_MS = 60 * 1000;         // 聚合結果快取有效期
const EDIT_TRIGGER_COLUMNS = [          // 需要觸發的「一基底」欄號 (人類視角)
  CLIENT_COL + 1,
  PLATFORM_COL + 1,
  AD_FORMAT_COL + 1,
  META_BUDGET_COL + 1,
  CAMPAIGN_TYPE_COL + 1,
  FREQUENCY_COL + 1,
  START_DATE_COL + 1,
  END_DATE_COL + 1
];
// --- END PERFORMANCE / CONTROL CONFIG ---

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dashboard')
    .addItem('Refresh Full Dashboard', 'createFullDashboard')
    .addItem('Refresh Summary Pivot', 'createSummaryPivotTable')
    .addToUi();
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    if (e.range.getSheet().getName() !== SOURCE_SHEET_NAME) return;
    if (!shouldRefreshOnEdit(e.range)) return;
    createFullDashboard();
  } catch (_) {}
}

function shouldRefreshOnEdit(range) {
  const col = range.getColumn();
  return EDIT_TRIGGER_COLUMNS.indexOf(col) !== -1;
}

/**
 * Main function to generate or refresh all visualizations.
 */
function createFullDashboard() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    Logger.log('Skipped: lock not acquired.');
    return;
  }
  try {
    const props = PropertiesService.getDocumentProperties();
    const lastRun = Number(props.getProperty('LAST_DASHBOARD_RUN') || 0);
    const now = Date.now();
    if (now - lastRun < THROTTLE_MS) {
      Logger.log('Skipped: throttled.');
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) throw new Error(`Source sheet "${SOURCE_SHEET_NAME}" not found.`);

    let dashboardSheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
    if (dashboardSheet) {
      // 僅清除內容與圖表，不刪除工作表 (保留權限 / 保護)
      dashboardSheet.getCharts().forEach(c => dashboardSheet.removeChart(c));
      dashboardSheet.clear();
    } else {
      dashboardSheet = ss.insertSheet(DASHBOARD_SHEET_NAME);
    }

    const lastRow = sourceSheet.getLastRow();
    if (lastRow < 2) throw new Error('No data rows.');

    const lastCol = sourceSheet.getLastColumn();
    const data = sourceSheet.getRange(2, 1, lastRow - 1, lastCol).getValues(); // 跳過標題列

    const aggregates = getAggregatesWithCache(data);

    buildMonthlyBudgetChart(dashboardSheet, aggregates.monthlyBudgets);
    buildCampaignsByClientChart(dashboardSheet, aggregates.clientCounts);
    buildFrequencyByClientChart(dashboardSheet, aggregates.frequencyCounts);
    buildAugustBudgetChart(dashboardSheet, aggregates.augustBudgets);
    createSummaryPivotTable(ss, sourceSheet);

    dashboardSheet.activate();
    props.setProperty('LAST_DASHBOARD_RUN', String(now));
  } catch (e) {
    handleError('Dashboard creation failed', e);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// 快取 wrapper
function getAggregatesWithCache(data) {
  const props = PropertiesService.getDocumentProperties();
  const cacheStamp = Number(props.getProperty('AGG_CACHE_TS') || 0);
  const now = Date.now();
  if (now - cacheStamp < CACHE_TTL_MS) {
    const cached = props.getProperty('AGG_CACHE_JSON');
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (_) {}
    }
  }
  const agg = processDataForCharts(data);
  props.setProperty('AGG_CACHE_TS', String(now));
  props.setProperty('AGG_CACHE_JSON', JSON.stringify(agg));
  return agg;
}

// --- Chart Building Functions ---
function buildMonthlyBudgetChart(sheet, budgets) {
  const chartData = [['Campaign Type', 'Budget'], ...Object.entries(budgets)];
  if (chartData.length <= 1) return;
  const dataRange = sheet.getRange(1, 1, chartData.length, 2).setValues(chartData);
  const chart = sheet.newChart().setChartType(Charts.ChartType.BAR).addRange(dataRange)
    .setOption('title', 'Monthly Budget by Campaign Type').setOption('titleTextStyle', {fontSize: 16, bold: true})
    .setOption('hAxis', { title: 'Total Meta Budget', format: 'short' }).setOption('vAxis', { title: 'Campaign Type' })
    .setOption('series', { 0: { dataLabel: 'value' } }).setPosition(2, 3, 0, 0).build();
  sheet.insertChart(chart);
}

function buildCampaignsByClientChart(sheet, clientCounts) {
  const chartData = [['Client', 'Number of Campaigns'], ...Object.entries(clientCounts)];
  if (chartData.length <= 1) return;
  const dataRange = sheet.getRange(1, 10, chartData.length, 2).setValues(chartData);
  const chart = sheet.newChart().setChartType(Charts.ChartType.PIE).addRange(dataRange)
    .setOption('title', 'Campaigns by Client').setOption('titleTextStyle', {fontSize: 16, bold: true})
    .setOption('pieHole', 0.4).setPosition(2, 12, 0, 0).build();
  sheet.insertChart(chart);
}

function buildFrequencyByClientChart(sheet, frequencyData) {
  const { clients, allFrequencies } = frequencyData;
  const freqArray = Array.from(allFrequencies);
  const chartData = [['Client', ...freqArray]];
  for (const client in clients) {
    const row = [client];
    freqArray.forEach(freq => row.push(clients[client][freq] || 0));
    chartData.push(row);
  }
  if (chartData.length <= 1) return;
  const dataRange = sheet.getRange(25, 1, chartData.length, chartData[0].length).setValues(chartData);
  const chart = sheet.newChart().setChartType(Charts.ChartType.COLUMN).addRange(dataRange)
    .setOption('isStacked', 'true').setOption('title', 'Frequency by Client')
    .setOption('titleTextStyle', {fontSize: 16, bold: true}).setOption('hAxis', { title: 'Client' })
    .setOption('vAxis', { title: 'Count' }).setPosition(25, 3, 0, 0).build();
  sheet.insertChart(chart);
}

function buildAugustBudgetChart(sheet, budgets) {
  const chartData = [['Ad Format', 'Budget'], ...Object.entries(budgets)];
  if (chartData.length <= 1) return;
  const dataRange = sheet.getRange(25, 10, chartData.length, 2).setValues(chartData);
  const chart = sheet.newChart().setChartType(Charts.ChartType.BAR).addRange(dataRange)
    .setOption('title', 'Monthly Budget by Ad Format (August 2025)').setOption('titleTextStyle', {fontSize: 16, bold: true})
    .setOption('hAxis', { title: 'Total Meta Budget', format: 'short' }).setOption('vAxis', { title: 'Ad Format' })
    .setOption('series', { 0: { dataLabel: 'value' } }).setPosition(25, 12, 0, 0).build();
  sheet.insertChart(chart);
}

// --- Pivot Table Function ---
function createSummaryPivotTable(ss, sourceSheet) {
  try {
    if (!ss || !sourceSheet) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
      if (!sourceSheet) throw new Error(`Source sheet "${SOURCE_SHEET_NAME}" not found.`);
    }
    let pivotSheet = ss.getSheetByName(PIVOT_SHEET_NAME);
    if (pivotSheet) pivotSheet.clear();
    else pivotSheet = ss.insertSheet(PIVOT_SHEET_NAME);
    
    const sourceDataRange = sourceSheet.getRange("B1:U" + sourceSheet.getLastRow());
    const pivotTable = pivotSheet.getRange('A1').createPivotTable(sourceDataRange);
    pivotTable.addRowGroup(PLATFORM_COL + 1);
    pivotTable.addRowGroup(CLIENT_COL + 1);
    pivotTable.addColumnGroup(AD_FORMAT_COL + 1);
    pivotTable.addPivotValue(CAMPAIGN_TYPE_COL + 1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName("Campaign Count");
    pivotTable.addPivotValue(META_BUDGET_COL + 1, SpreadsheetApp.PivotTableSummarizeFunction.SUM).setDisplayName("Total Meta Budget");
    pivotSheet.autoResizeColumns(1, pivotSheet.getLastColumn());
  } catch (e) {
    handleError('Pivot table creation failed', e);
  }
}

// --- Helper Functions ---
const isValidDate = d => d instanceof Date && !isNaN(d);
function handleError(message, e) {
  const concise = `${message}: ${e && e.message ? e.message : e}`;
  Logger.log(concise + '\n' + (e && e.stack ? e.stack : ''));
  try {
    SpreadsheetApp.getUi().alert(concise);
  } catch (_) {}
}
