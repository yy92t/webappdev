/**
 * @OnlyCurrentDoc
 * This script powers a campaign dashboard web app and includes automation for data snapshots.
 */

// --- CONFIGURATION ---
const SOURCE_SHEET_NAME = "Weekly log_Thomas W";
const SNAPSHOT_SHEET_NAME = "Daily Data Snapshot";
// Column numbers are 0-indexed (Column A = 0)
const COLUMNS = {
  CAMPAIGN_ID: 22, CLIENT: 6, PLATFORM: 7, AD_FORMAT: 8, META_BUDGET: 9,
  CAMPAIGN_TYPE: 12, STATUS: 5, START_DATE: 19, END_DATE: 20
};
// --- END CONFIGURATION ---

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Progress Hub -- Ads Ops');
}

function _getSheetData() {
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) return { error: `Source sheet "${SOURCE_SHEET_NAME}" not found.` };
  const data = sourceSheet.getDataRange().getValues();
  data.shift(); // Remove header row
  return { data };
}

function getDashboardData() {
  const { data, error } = _getSheetData();
  if (error) return { error };
  return { charts: processDataForCharts(data) };
}

function searchCampaignByType(campaignType) {
  const { data, error } = _getSheetData();
  if (error) return { error };
  
  return data
    .filter(row => row[COLUMNS.CAMPAIGN_TYPE] && String(row[COLUMNS.CAMPAIGN_TYPE]).toUpperCase().includes(campaignType.toUpperCase()))
    .map(row => ({
      id: row[COLUMNS.CAMPAIGN_ID], client: row[COLUMNS.CLIENT], platform: row[COLUMNS.PLATFORM],
      campaignType: row[COLUMNS.CAMPAIGN_TYPE], status: row[COLUMNS.STATUS], budget: row[COLUMNS.META_BUDGET],
      startDate: Utilities.formatDate(new Date(row[COLUMNS.START_DATE]), "GMT", "yyyy-MM-dd"),
      endDate: Utilities.formatDate(new Date(row[COLUMNS.END_DATE]), "GMT", "yyyy-MM-dd")
    }))
    .sort((a, b) => new Date(b.startDate) - new Date(a.startDate));
}

function processDataForCharts(data) {
  const today = new Date();
  const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const endOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  endOfMonth.setHours(23, 59, 59, 999);

  const aggregators = {
    monthlyBudget: {}, platformClient: {}, allPlatforms: new Set(),
    frequencyByClient: {}, campaignDurations: {}
  };

  for (const row of data) {
    const client = row[COLUMNS.CLIENT];
    const platform = row[COLUMNS.PLATFORM];
    const budget = row[COLUMNS.META_BUDGET];
    const startDate = new Date(row[COLUMNS.START_DATE]);
    const endDate = new Date(row[COLUMNS.END_DATE]);

    if (!isValidDate(startDate) || !isValidDate(endDate)) continue;

    // Aggregate data for multiple charts in one go
    if (client && platform && typeof budget === 'number') {
      const clientStr = String(client).trim().toUpperCase();
      if (clientStr !== 'N/A') {
        // Platform by Client
        if (!aggregators.platformClient[client]) aggregators.platformClient[client] = {};
        aggregators.platformClient[client][platform] = (aggregators.platformClient[client][platform] || 0) + budget;
        aggregators.allPlatforms.add(platform);
        // Frequency by Client
        const key = `${client} - ${platform}`;
        aggregators.frequencyByClient[key] = (aggregators.frequencyByClient[key] || 0) + budget;
      }
    }

    // Meta budget in the current month
    const adFormat = row[COLUMNS.AD_FORMAT];
    if (startDate <= endOfMonth && endDate >= startOfMonth && adFormat && typeof budget === 'number') {
      aggregators.monthlyBudget[adFormat] = (aggregators.monthlyBudget[adFormat] || 0) + budget;
    }

    // Campaign Duration
    const campaignType = row[COLUMNS.CAMPAIGN_TYPE];
    if (campaignType) {
      const duration = (endDate - startDate) / 86400000; // ms to days
      aggregators.campaignDurations[campaignType] = (aggregators.campaignDurations[campaignType] || 0) + duration;
    }
  }

  // Format data for Google Charts
  const top5CampaignDurations = Object.entries(aggregators.campaignDurations).sort((a, b) => b[1] - a[1]).slice(0, 5);
  const sortedClients = Object.keys(aggregators.platformClient).sort();
  const sortedPlatforms = Array.from(aggregators.allPlatforms).sort();
  const platformClientData = [['Client', ...sortedPlatforms]];
  sortedClients.forEach(client => {
    const row = [client, ...sortedPlatforms.map(p => aggregators.platformClient[client][p] || 0)];
    platformClientData.push(row);
  });

  return {
    platformClientCounts: platformClientData,
    frequencyByClient: [['Client - Platform', 'Total Meta Budget'], ...Object.entries(aggregators.frequencyByClient)],
    currentMonthBudgets: [['Ad Format', 'Budget'], ...Object.entries(aggregators.monthlyBudget)],
    campaignDurations: [['Campaign Type', 'Total Duration (Days)'], ...top5CampaignDurations]
  };
}

function dailyDataRefresh() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) return Logger.log(`Error: Source sheet "${SOURCE_SHEET_NAME}" not found.`);
    
    let snapshotSheet = ss.getSheetByName(SNAPSHOT_SHEET_NAME);
    if (!snapshotSheet) snapshotSheet = ss.insertSheet(SNAPSHOT_SHEET_NAME);

    const sourceData = sourceSheet.getDataRange().getValues();
    snapshotSheet.clear();
    snapshotSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
    Logger.log(`Data refresh complete. Copied ${sourceData.length} rows.`);
  } catch (e) {
    Logger.log(`An error occurred: ${e.message}`);
  }
}

const isValidDate = d => d instanceof Date && !isNaN(d);