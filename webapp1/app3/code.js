/**
 * @OnlyCurrentDoc
 * Ad Ops Hub backend (app3) – optimized with caching and safe parsing.
 */

// --- CONFIGURATION ---
const SOURCE_SHEET_NAME = "Weekly log_Thomas W";
const PERMISSIONS_SHEET_NAME = "Permissions";
const SNAPSHOT_SHEET_NAME = "Daily Data Snapshot";
const MS_PER_DAY = 86400000;
const COLUMNS = {
  CAMPAIGN_ID: 22, CLIENT: 6, PLATFORM: 7, AD_FORMAT: 8, META_BUDGET: 9,
  CAMPAIGN_TYPE: 12, STATUS: 5, START_DATE: 19, END_DATE: 20,
  REMARKS: 42,
  GUIDE_LINK: 43
};
const ENTRY_COLUMN_INDEX = 43; // Column AQ is the 43rd column
// --- END CONFIGURATION ---

// --- CACHE CONFIG ---
const CACHE_EXP_SECONDS = 60; // base dashboard cache
const CACHE_KEY_BASE_DATA = 'APP3_DASHBOARD_BASE_V1';
const CACHE_KEY_SEARCH_PREFIX = 'APP3_SEARCH_CT_';

// --- UTILITIES ---
/** @param {any} v @returns {Date|null} */
function safeDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number') { const d = new Date(v); return isNaN(d) ? null : d; }
  if (typeof v === 'string' && v.trim()) { const d = new Date(v.trim()); return isNaN(d) ? null : d; }
  return null;
}
/** @param {any} v @returns {number|null} */
function toNumber(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (typeof v === 'string' && v.trim()) {
    const n = Number(v.replace(/[^0-9.+-]/g, ''));
    return isFinite(n) ? n : null;
  }
  return null;
}
/** @param {any} v */
function isNonEmpty(v) { return typeof v === 'string' ? v.trim() !== '' && v.trim().toUpperCase() !== 'N/A' : v != null; }

/**
 * Gets the user's role from the 'Permissions' sheet.
 */
function getUserRole() {
  const userEmail = Session.getActiveUser().getEmail();
  const permissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PERMISSIONS_SHEET_NAME);
  if (!permissionsSheet) return null;

  const roles = permissionsSheet.getDataRange().getValues();
  roles.shift(); // Remove header

  for (const row of roles) {
    if (row[0] && typeof row[0] === 'string') {
      if (row[0].trim().toLowerCase() === userEmail.toLowerCase()) {
        return row[1]; // Return the user's role
      }
    }
  }
  return null;
}

/**
 * Appends a code to the next empty cell in the entry column (AQ).
 * Includes a security check to ensure only admins can perform this action.
 */
function appendCodeToSheet(code) {
  const userRole = getUserRole();
  if (userRole !== 'admin') {
    throw new Error('Permission Denied: Only admins can submit new entries.');
  }

  if (!code || String(code).trim() === '') {
    throw new Error('Invalid input: Code cannot be empty.');
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SOURCE_SHEET_NAME}" not found.`);

    const columnValues = sheet.getRange(1, ENTRY_COLUMN_INDEX, sheet.getMaxRows(), 1).getValues();
    let firstEmptyRow = columnValues.findIndex(row => row[0] === '') + 1;
    if (firstEmptyRow === 0) { 
      firstEmptyRow = sheet.getLastRow() + 1;
    }

  sheet.getRange(firstEmptyRow, ENTRY_COLUMN_INDEX).setValue(code);
  invalidateDashboardCache();
    
    return { success: true, message: `Code "${code}" submitted to row ${firstEmptyRow}.` };
  } catch (e) {
    throw new Error(`Failed to append to sheet: ${e.message}`);
  }
}


function checkUserAccess() {
  return !!getUserRole();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Ad Ops Hub');
}


function _getSheetData() {
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) return { error: `Source sheet "${SOURCE_SHEET_NAME}" not found.` };

  const lastRow = sourceSheet.getLastRow();
  if (lastRow <= 1) return { data: [] };

  const neededCols = Object.values(COLUMNS);
  const minCol = Math.min.apply(null, neededCols);
  const maxCol = Math.max.apply(null, neededCols);
  const width = maxCol - minCol + 1;

  const rawData = sourceSheet.getRange(2, minCol, lastRow - 1, width).getValues();

  const processedData = [];
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const off = idx => row[idx - minCol];
    const startDate = safeDate(off(COLUMNS.START_DATE));
    const endDate = safeDate(off(COLUMNS.END_DATE));
    processedData.push({
      campaignId: off(COLUMNS.CAMPAIGN_ID),
      client: off(COLUMNS.CLIENT),
      platform: off(COLUMNS.PLATFORM),
      adFormat: off(COLUMNS.AD_FORMAT),
      budget: toNumber(off(COLUMNS.META_BUDGET)),
      campaignType: off(COLUMNS.CAMPAIGN_TYPE),
      status: off(COLUMNS.STATUS),
      startDate: startDate || null,
      endDate: endDate || null,
      remarks: off(COLUMNS.REMARKS),
      guideLink: off(COLUMNS.GUIDE_LINK)
    });
  }
  return { data: processedData };
}

function getDashboardData() {
  const cache = CacheService.getDocumentCache();
  let basePayload;
  try {
    const hit = cache.get(CACHE_KEY_BASE_DATA);
    if (hit) {
      basePayload = JSON.parse(hit);
    } else {
      basePayload = buildBaseDashboardData();
      if (!basePayload.error) cache.put(CACHE_KEY_BASE_DATA, JSON.stringify(basePayload), CACHE_EXP_SECONDS);
    }
  } catch (_) {
    basePayload = buildBaseDashboardData();
  }
  if (basePayload.error) return basePayload;

  const email = Session.getActiveUser().getEmail();
  const first = email ? email.split('@')[0].split('.')[0] : 'User';
  const userName = first ? first.charAt(0).toUpperCase() + first.slice(1) : 'User';

  return {
    charts: basePayload.charts,
    clients: basePayload.clients,
    userName,
    userRole: getUserRole()
  };
}

function buildBaseDashboardData() {
  const { data, error } = _getSheetData();
  if (error) return { error };
  const charts = processDataForCharts(data);
  const clientList = [...new Set(data.filter(r => isNonEmpty(r.client)).map(r => r.client))].sort();
  return { charts, clients: clientList };
}

function invalidateDashboardCache() {
  try { CacheService.getDocumentCache().remove(CACHE_KEY_BASE_DATA); } catch(_) {}
}

function searchCampaignByType(campaignType) {
  if (!campaignType || !campaignType.trim()) return [];
  const needle = campaignType.trim().toUpperCase();
  const { data, error } = _getSheetData();
  if (error) return { error };

  const cache = CacheService.getDocumentCache();
  const ck = CACHE_KEY_SEARCH_PREFIX + needle;
  try {
    const hit = cache.get(ck);
    if (hit) return JSON.parse(hit);
  } catch (_) {}

  const searchResults = data
    .filter(row => row.campaignType && String(row.campaignType).toUpperCase().includes(needle))
    .map(row => ({
      ...row,
      startDate: row.startDate ? Utilities.formatDate(new Date(row.startDate), 'GMT', 'yyyy-MM-dd') : '',
      endDate: row.endDate ? Utilities.formatDate(new Date(row.endDate), 'GMT', 'yyyy-MM-dd') : '',
      id: row.campaignId
    }));
  const sorted = searchResults.sort((a, b) => new Date(b.startDate) - new Date(a.startDate));
  try { cache.put(ck, JSON.stringify(sorted.slice(0, 300)), 30); } catch (_) {}
  return sorted;
}


function processDataForCharts(data) {
  if (!data || !data.length) {
    return {
      platformClientCounts: [['Client']],
      frequencyByClient: [['Client - Platform', 'Total Meta Budget']],
      currentMonthBudgets: [['Ad Format', 'Budget']],
      campaignDurations: [['Campaign Type', 'Total Duration (Days)']],
      campaignStatusCounts: [['Status', 'Count']],
      monthlyBudgetByClient: [['Client', 'Meta Budget']]
    };
  }

  const today = new Date();
  const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1).getTime();
  const endOfMonthDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  endOfMonthDate.setHours(23, 59, 59, 999);
  const endOfMonth = endOfMonthDate.getTime();

  const aggregators = {
    monthlyBudget: {}, platformClient: {}, allPlatforms: new Set(),
    frequencyByClient: {}, campaignDurations: {}, campaignStatus: {},
    monthlyBudgetByClient: {}
  };

  for (const row of data) {
    if (!isValidDate(row.startDate) || !isValidDate(row.endDate)) continue;
    const budget = toNumber(row.budget);
    const hasBudget = budget != null;
    const sd = row.startDate.getTime();
    const ed = row.endDate.getTime();

    if (row.status) aggregators.campaignStatus[row.status] = (aggregators.campaignStatus[row.status] || 0) + 1;

    if (isNonEmpty(row.client) && isNonEmpty(row.platform) && hasBudget) {
      if (!aggregators.platformClient[row.client]) aggregators.platformClient[row.client] = {};
      aggregators.platformClient[row.client][row.platform] = (aggregators.platformClient[row.client][row.platform] || 0) + budget;
      aggregators.allPlatforms.add(row.platform);
      const key = `${row.client} - ${row.platform}`;
      aggregators.frequencyByClient[key] = (aggregators.frequencyByClient[key] || 0) + budget;
    }

    if (sd <= endOfMonth && ed >= startOfMonth && isNonEmpty(row.adFormat) && hasBudget) {
      aggregators.monthlyBudget[row.adFormat] = (aggregators.monthlyBudget[row.adFormat] || 0) + budget;
    }

    if (sd <= endOfMonth && ed >= startOfMonth && isNonEmpty(row.client) && hasBudget) {
      aggregators.monthlyBudgetByClient[row.client] = (aggregators.monthlyBudgetByClient[row.client] || 0) + budget;
    }

    if (isNonEmpty(row.campaignType)) {
      const duration = (ed - sd) / MS_PER_DAY;
      aggregators.campaignDurations[row.campaignType] = (aggregators.campaignDurations[row.campaignType] || 0) + duration;
    }
  }

  const sortedPlatforms = Array.from(aggregators.allPlatforms).sort();
  
  const platformClientData = [
    ['Client', ...sortedPlatforms],
    ...Object.keys(aggregators.platformClient).sort().map(client => [client, ...sortedPlatforms.map(p => aggregators.platformClient[client][p] || 0)])
  ].filter((row, idx) => idx === 0 || row.slice(1).some(v => v !== 0));

  const top5CampaignDurations = Object.entries(aggregators.campaignDurations).sort((a, b) => b[1] - a[1]).slice(0, 5);
  
  const monthlyBudgetDataByClient = [['Client', 'Meta Budget'], ...Object.entries(aggregators.monthlyBudgetByClient)].filter((row, idx) => idx === 0 || row[1] !== 0);

  return {
    platformClientCounts: platformClientData,
    frequencyByClient: [['Client - Platform', 'Total Meta Budget'], ...Object.entries(aggregators.frequencyByClient)],
    currentMonthBudgets: [['Ad Format', 'Budget'], ...Object.entries(aggregators.monthlyBudget)],
    campaignDurations: [['Campaign Type', 'Total Duration (Days)'], ...top5CampaignDurations],
    campaignStatusCounts: [['Status', 'Count'], ...Object.entries(aggregators.campaignStatus)],
    monthlyBudgetByClient: monthlyBudgetDataByClient
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
    invalidateDashboardCache();
    Logger.log(`Data refresh complete. Copied ${sourceData.length} rows (cache invalidated).`);
  } catch (e) {
    Logger.log(`An error occurred: ${e.message}`);
  }
}

const isValidDate = d => d instanceof Date && !isNaN(d);
