 * @OnlyCurrentDoc

 *

 * This script creates an auto-refreshing dashboard with four charts and a summary pivot table

 * summarizing campaign data from the "Weekly log_Thomas W" sheet.

 */



// --- CONFIGURATION ---

const SOURCE_SHEET_NAME = "Weekly log_Thomas W";

const DASHBOARD_SHEET_NAME = "Full Dashboard";

const PIVOT_SHEET_NAME = "Summary Pivot Table"; // New sheet for the pivot table



// Column numbers from your "Internal Hub" file

const CLIENT_COLUMN = 7;         // Column G

const PLATFORM_COLUMN = 8;       // Column H

const AD_FORMAT_COLUMN = 9;      // Column I

const META_BUDGET_COLUMN = 10;   // Column J

const CAMPAIGN_TYPE_COLUMN = 13; // Column M

const FREQUENCY_COLUMN = 19;     // Column S

const START_DATE_COLUMN = 20;    // Column T

const END_DATE_COLUMN = 21;      // Column U

// --- END CONFIGURATION ---



/**

 * Adds a custom menu to the spreadsheet UI.

 */

function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu('Dashboard')

    .addItem('Refresh Full Dashboard', 'createFullDashboard')

    .addItem('Refresh Summary Pivot', 'createSummaryPivotTable') // New menu item

    .addToUi();

}



/**

 * Automatically runs when a user changes a value in the source sheet.

 */

function onEdit(e) {

  if (e.range.getSheet().getName() === SOURCE_SHEET_NAME) {

    createFullDashboard();

  }

}



/**

 * Main function to generate or refresh all four charts and the pivot table.

 */

function createFullDashboard() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);

  if (!sourceSheet) {

    SpreadsheetApp.getUi().alert(`Source sheet "${SOURCE_SHEET_NAME}" not found.`);

    return;

  }



  // Get or create the dashboard sheet, then clear it.

  let dashboardSheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (dashboardSheet) {

    dashboardSheet.getCharts().forEach(chart => dashboardSheet.removeChart(chart));

    dashboardSheet.clear();

  } else {

    dashboardSheet = ss.insertSheet(DASHBOARD_SHEET_NAME);

  }



  const data = sourceSheet.getDataRange().getValues();

  data.shift(); // Remove header row



  // --- Create all visualizations ---

  createMonthlyBudgetByCampaignType(dashboardSheet, data);

  createCampaignsByClient(dashboardSheet, data);

  createFrequencyByClient(dashboardSheet, data);

  createAugustBudgetByAdFormat(dashboardSheet, data);

  createSummaryPivotTable(ss, sourceSheet); // Call the new pivot table function

 

  dashboardSheet.activate();

}



/**

 * --- NEW PIVOT TABLE FUNCTION ---

 * Creates a summary pivot table on a separate sheet.

 */

function createSummaryPivotTable(ss, sourceSheet) {

  let pivotSheet = ss.getSheetByName(PIVOT_SHEET_NAME);

  if (pivotSheet) {

    pivotSheet.clear();

  } else {

    pivotSheet = ss.insertSheet(PIVOT_SHEET_NAME);

  }



  const sourceDataRange = sourceSheet.getRange("B1:U" + sourceSheet.getLastRow());

  const pivotTable = pivotSheet.getRange('A1').createPivotTable(sourceDataRange);



  // Set up Rows

  pivotTable.addRowGroup(PLATFORM_COLUMN);

  pivotTable.addRowGroup(CLIENT_COLUMN);



  // Set up Columns

  pivotTable.addColumnGroup(AD_FORMAT_COLUMN);



  // Set up Values

  pivotTable.addPivotValue(CAMPAIGN_TYPE_COLUMN, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA)

            .setDisplayName("Campaign Count");

  pivotTable.addPivotValue(META_BUDGET_COLUMN, SpreadsheetApp.PivotTableSummarizeFunction.SUM)

            .setDisplayName("Total Meta Budget");

           

  pivotSheet.autoResizeColumns(1, pivotSheet.getLastColumn());

}





// --- CHART 1: Monthly Budget by Campaign Type (Bar Chart) ---

function createMonthlyBudgetByCampaignType(sheet, data) {

  const today = new Date();

  const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);

  const endOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);

  endOfMonth.setHours(23, 59, 59, 999);



  const budgets = {};

  data.forEach(row => {

    const startDate = new Date(row[START_DATE_COLUMN - 1]);

    const endDate = new Date(row[END_DATE_COLUMN - 1]);

    if (isValidDate(startDate) && isValidDate(endDate) && startDate <= endOfMonth && endDate >= startOfMonth) {

      const campaignType = row[CAMPAIGN_TYPE_COLUMN - 1];

      const budget = row[META_BUDGET_COLUMN - 1];

      if (campaignType && typeof campaignType === 'string' && campaignType.trim().toUpperCase() !== 'N/A' && typeof budget === 'number') {

        budgets[campaignType] = (budgets[campaignType] || 0) + budget;

      }

    }

  });



  const chartData = [['Campaign Type', 'Budget'], ...Object.entries(budgets)];

  if (chartData.length <= 1) return;

 

  const dataRange = sheet.getRange(1, 1, chartData.length, 2).setValues(chartData);



  const chart = sheet.newChart().setChartType(Charts.ChartType.BAR)

    .addRange(dataRange)

    .setOption('title', 'Monthly Budget by Campaign Type')

    .setOption('titleTextStyle', {fontSize: 16, bold: true})

    .setOption('hAxis', { title: 'Total Meta Budget', format: 'short' })

    .setOption('vAxis', { title: 'Campaign Type' })

    .setOption('series', { 0: { dataLabel: 'value' } })

    .setPosition(2, 3, 0, 0).build();

  sheet.insertChart(chart);

}



// --- CHART 2: Campaigns by Client (Pie Chart) ---

function createCampaignsByClient(sheet, data) {

  const clientCounts = {};

  data.forEach(row => {

    const client = row[CLIENT_COLUMN - 1];

    if (client && typeof client === 'string' && client.trim().toUpperCase() !== 'N/A') {

      clientCounts[client] = (clientCounts[client] || 0) + 1;

    }

  });



  const chartData = [['Client', 'Number of Campaigns'], ...Object.entries(clientCounts)];

  if (chartData.length <= 1) return;



  const dataRange = sheet.getRange(1, 10, chartData.length, 2).setValues(chartData);



  const chart = sheet.newChart().setChartType(Charts.ChartType.PIE)

    .addRange(dataRange)

    .setOption('title', 'Campaigns by Client')

    .setOption('titleTextStyle', {fontSize: 16, bold: true})

    .setOption('pieHole', 0.4)

    .setPosition(2, 12, 0, 0).build();

  sheet.insertChart(chart);

}



// --- CHART 3: Frequency by Client (Stacked Column Chart) ---

function createFrequencyByClient(sheet, data) {

  const clientFrequencies = {};

  const allFrequencies = new Set();

  data.forEach(row => {

    const client = row[CLIENT_COLUMN - 1];

    const frequency = row[FREQUENCY_COLUMN - 1];

    if (client && typeof client === 'string' && client.trim().toUpperCase() !== 'N/A' && frequency && typeof frequency === 'string') {

      if (!clientFrequencies[client]) clientFrequencies[client] = {};

      clientFrequencies[client][frequency] = (clientFrequencies[client][frequency] || 0) + 1;

      allFrequencies.add(frequency);

    }

  });



  const freqArray = Array.from(allFrequencies);

  const chartData = [['Client', ...freqArray]];

  for (const client in clientFrequencies) {

    const row = [client];

    freqArray.forEach(freq => {

      row.push(clientFrequencies[client][freq] || 0);

    });

    chartData.push(row);

  }



  if (chartData.length <= 1) return;



  const dataRange = sheet.getRange(25, 1, chartData.length, chartData[0].length).setValues(chartData);



  const chart = sheet.newChart().setChartType(Charts.ChartType.COLUMN)

    .addRange(dataRange)

    .setOption('isStacked', 'true')

    .setOption('title', 'Frequency by Client')

    .setOption('titleTextStyle', {fontSize: 16, bold: true})

    .setOption('hAxis', { title: 'Client' })

    .setOption('vAxis', { title: 'Count' })

    .setPosition(25, 3, 0, 0).build();

  sheet.insertChart(chart);

}



// --- CHART 4: Monthly Budget by Ad Format in August (Bar Chart) ---

function createAugustBudgetByAdFormat(sheet, data) {

  const startOfAugust = new Date(2025, 7, 1);

  const endOfAugust = new Date(2025, 8, 0);

  endOfAugust.setHours(23, 59, 59, 999);



  const budgets = {};

  data.forEach(row => {

    const startDate = new Date(row[START_DATE_COLUMN - 1]);

    const endDate = new Date(row[END_DATE_COLUMN - 1]);

    if (isValidDate(startDate) && isValidDate(endDate) && startDate <= endOfAugust && endDate >= startOfAugust) {

      const adFormat = row[AD_FORMAT_COLUMN - 1];

      const budget = row[META_BUDGET_COLUMN - 1];

      if (adFormat && typeof adFormat === 'string' && adFormat.trim().toUpperCase() !== 'N/A' && typeof budget === 'number') {

        budgets[adFormat] = (budgets[adFormat] || 0) + budget;

      }

    }

  });



  const chartData = [['Ad Format', 'Budget'], ...Object.entries(budgets)];

  if (chartData.length <= 1) return;



  const dataRange = sheet.getRange(25, 10, chartData.length, 2).setValues(chartData);



  const chart = sheet.newChart().setChartType(Charts.ChartType.BAR)

    .addRange(dataRange)

    .setOption('title', 'Monthly Budget by Ad Format (August 2025)')

    .setOption('titleTextStyle', {fontSize: 16, bold: true})

    .setOption('hAxis', { title: 'Total Meta Budget', format: 'short' })

    .setOption('vAxis', { title: 'Ad Format' })

    .setOption('series', { 0: { dataLabel: 'value' } })

    .setPosition(25, 12, 0, 0).build();

  sheet.insertChart(chart);

}



// --- HELPER FUNCTION ---

function isValidDate(d) {

  return d instanceof Date && !isNaN(d);

}