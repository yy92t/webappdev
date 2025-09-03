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

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dashboard')
    .addItem('Refresh Full Dashboard', 'createFullDashboard')
    .addItem('Refresh Summary Pivot', 'createSummaryPivotTable')
    .addToUi();
}

function onEdit(e) {
  if (e.range.getSheet().getName() === SOURCE_SHEET_NAME) {
    createFullDashboard();
  }
}

/**
 * Main function to generate or refresh all visualizations.
 */
function createFullDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) throw new Error(`Source sheet "${SOURCE_SHEET_NAME}" not found.`);

    let dashboardSheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
    if (dashboardSheet) {
      dashboardSheet.getCharts().forEach(chart => dashboardSheet.removeChart(chart));
      dashboardSheet.clear();
    } else {
      dashboardSheet = ss.insertSheet(DASHBOARD_SHEET_NAME);
    }

    const data = sourceSheet.getDataRange().getValues();
    data.shift(); // Remove header row

    // --- Process data for all charts in a single loop for performance ---
    const chartData = processDataForCharts(data);

    // --- Create all visualizations ---
    buildMonthlyBudgetChart(dashboardSheet, chartData.monthlyBudgets);
    buildCampaignsByClientChart(dashboardSheet, chartData.clientCounts);
    buildFrequencyByClientChart(dashboardSheet, chartData.frequencyCounts);
    buildAugustBudgetChart(dashboardSheet, chartData.augustBudgets);
    createSummaryPivotTable(ss, sourceSheet);

    dashboardSheet.activate();
  } catch (e) {
    handleError('Dashboard creation failed', e);
  }
}

/**
 * Processes the raw data and aggregates it for all charts in one pass.
 * @param {Array<Array<Object>>} data The raw data from the sheet.
 * @return {Object} An object containing the aggregated data for each chart.
 */
function processDataForCharts(data) {
  const today = new Date();
  const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const endOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  endOfMonth.setHours(23, 59, 59, 999);
  
  const startOfAugust = new Date(2025, 7, 1);
  const endOfAugust = new Date(2025, 8, 0);
  endOfAugust.setHours(23, 59, 59, 999);

  const aggregates = {
    monthlyBudgets: {},
    clientCounts: {},
    frequencyCounts: { clients: {}, allFrequencies: new Set() },
    augustBudgets: {}
  };

  data.forEach(row => {
    const client = row[CLIENT_COL];
    const campaignType = row[CAMPAIGN_TYPE_COL];
    const adFormat = row[AD_FORMAT_COL];
    const budget = row[META_BUDGET_COL];
    const frequency = row[FREQUENCY_COL];
    const startDate = new Date(row[START_DATE_COL]);
    const endDate = new Date(row[END_DATE_COL]);

    if (!isValidDate(startDate) || !isValidDate(endDate)) return;

    // Chart 1: Monthly Budget by Campaign Type
    if (startDate <= endOfMonth && endDate >= startOfMonth && campaignType && typeof campaignType === 'string' && campaignType.trim().toUpperCase() !== 'N/A' && typeof budget === 'number') {
      aggregates.monthlyBudgets[campaignType] = (aggregates.monthlyBudgets[campaignType] || 0) + budget;
    }

    // Chart 2: Campaigns by Client
    if (client && typeof client === 'string' && client.trim().toUpperCase() !== 'N/A') {
      aggregates.clientCounts[client] = (aggregates.clientCounts[client] || 0) + 1;
    }

    // Chart 3: Frequency by Client
    if (client && typeof client === 'string' && client.trim().toUpperCase() !== 'N/A' && frequency && typeof frequency === 'string') {
      if (!aggregates.frequencyCounts.clients[client]) aggregates.frequencyCounts.clients[client] = {};
      aggregates.frequencyCounts.clients[client][frequency] = (aggregates.frequencyCounts.clients[client][frequency] || 0) + 1;
      aggregates.frequencyCounts.allFrequencies.add(frequency);
    }

    // Chart 4: August Budget by Ad Format
    if (startDate <= endOfAugust && endDate >= startOfAugust && adFormat && typeof adFormat === 'string' && adFormat.trim().toUpperCase() !== 'N/A' && typeof budget === 'number') {
      aggregates.augustBudgets[adFormat] = (aggregates.augustBudgets[adFormat] || 0) + budget;
    }
  });
  return aggregates;
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
  const errorMessage = `${message}: ${e.message} (File: ${e.fileName}, Line: ${e.lineNumber})`;
  Logger.log(errorMessage + '\nStack: ' + e.stack);
  SpreadsheetApp.getUi().alert(`${message}. Please check script logs for details. Error: ${e.message}`);
}
