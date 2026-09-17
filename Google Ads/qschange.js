// Copyright 2024 Google LLC
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     https://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/**
 * @name Quality Score Change Tracker
 * @version 1.0
 * @author Google Ads Scripts Team
 *
 * This script tracks the Quality Score of keywords over time using a Google
 * Sheet. It sends an email report detailing any keywords whose Quality Score
 * has changed since the last run.
 */

// --- Your Settings (MUST BE CHANGED) ---

// 1. The email address where you want to receive the report.
const RECIPIENT_EMAIL = 'thomas.wong@newimedia.com';

// 2. The URL of the Google Sheet you created to store the QS data.
//    Example: 'https://docs.google.com/spreadsheets/d/1234567890ABCD_EFG-HIJKLMNOP/edit'
const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1sAdzW4BoiLeoKoAXz24kTHMzu1tqGr0i4QfWapYuW9o/edit?gid=0#gid=0';

// 3. The name of the sheet within the spreadsheet where data will be stored.
const SHEET_NAME = 'Quality Score Log';

// 4. Optional: The specific Google Ads Account ID this script should run on.
//    This is a safety measure. Format: "XXX-XXX-XXXX"
const ACCOUNT_ID_TO_CHECK = '959-935-7325';

// --- End of Settings ---

const HEADERS = [
  'Keyword ID', 'Keyword Text', 'Ad Group', 'Campaign', 'Quality Score', 'Last Checked'
];

function main() {
  // Validate settings before running.
  if (SPREADSHEET_URL === 'PASTE_YOUR_GOOGLE_SHEET_URL_HERE') {
    throw new Error('Please paste the URL of your Google Sheet into the SPREADSHEET_URL setting.');
  }

  // Safety check: Halt execution if the script is running on the wrong account.
  const currentAccountId = AdsApp.currentAccount().getCustomerId();
  if (ACCOUNT_ID_TO_CHECK && currentAccountId !== ACCOUNT_ID_TO_CHECK) {
    const errorMessage = `Error: Script is running on the wrong account. Expected: ${ACCOUNT_ID_TO_CHECK}, Actual: ${currentAccountId}`;
    console.error(errorMessage);
    MailApp.sendEmail(RECIPIENT_EMAIL, `[Action Required] Ads Script Misconfiguration on Account ${currentAccountId}`, errorMessage);
    throw new Error(errorMessage);
  }

  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    console.log(`Sheet "${SHEET_NAME}" was not found and has been created.`);
  }

  const oldQualityScores = getOldQualityScores(sheet);
  const keywordIterator = AdsApp.keywords()
    .withCondition('campaign.status = "ENABLED"')
    .withCondition('ad_group.status = "ENABLED"')
    // Keyword status and Quality Score cannot be used in '.withCondition()'.
    // They will be checked inside the loop instead.
    .get();

  const improvedKeywords = [];
  const declinedKeywords = [];
  const newData = [HEADERS]; // Start with headers for the new sheet content
  const today = new Date().toLocaleDateString();

  console.log('Fetching current keyword Quality Scores...');

  while (keywordIterator.hasNext()) {
    const keyword = keywordIterator.next();
    const qualityScore = keyword.getQualityScore();
    
    // Check the conditions that were removed from the selector.
    if (keyword.isEnabled() && qualityScore > 0) {
      const keywordId = keyword.getId();
      
      const rowData = [
          keywordId,
          keyword.getText(),
          keyword.getAdGroup().getName(),
          keyword.getCampaign().getName(),
          qualityScore,
          today
      ];
      newData.push(rowData);

      if (oldQualityScores[keywordId]) {
        const oldQS = oldQualityScores[keywordId];
        if (qualityScore > oldQS) {
          improvedKeywords.push({ keyword, oldQS, newQS: qualityScore });
        } else if (qualityScore < oldQS) {
          declinedKeywords.push({ keyword, oldQS, newQS: qualityScore });
        }
      }
    }
  }
  
  console.log(`Found ${improvedKeywords.length} improved and ${declinedKeywords.length} declined keywords.`);

  updateSheet(sheet, newData);
  sendEmailReport(improvedKeywords, declinedKeywords);
  
  console.log('Script finished successfully.');
}

/**
 * Reads the existing Quality Score data from the spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to read from.
 * @return {!Object<string, number>} A map of Keyword ID to Quality Score.
 */
function getOldQualityScores(sheet) {
  const data = sheet.getDataRange().getValues();
  const qualityScoreMap = {};
  // Start from row 1 to skip the header
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const keywordId = row[0]; // Column A: Keyword ID
    const qualityScore = row[4]; // Column E: Quality Score
    if (keywordId && qualityScore) {
      qualityScoreMap[keywordId] = qualityScore;
    }
  }
  console.log(`Loaded ${Object.keys(qualityScoreMap).length} previous keyword scores from the sheet.`);
  return qualityScoreMap;
}

/**
 * Clears the sheet and writes the new data.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to update.
 * @param {Array<Array<string|number>>} data The new data to write.
 */
function updateSheet(sheet, data) {
    sheet.clearContents();
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    console.log(`Updated spreadsheet with ${data.length - 1} current keyword scores.`);
}

/**
 * Sends an email report with the changes in Quality Score.
 * @param {Array<Object>} improvedKeywords An array of keywords that improved.
 * @param {Array<Object>} declinedKeywords An array of keywords that declined.
 */
function sendEmailReport(improvedKeywords, declinedKeywords) {
  if (improvedKeywords.length === 0 && declinedKeywords.length === 0) {
    console.log('No Quality Score changes detected. No email will be sent.');
    return;
  }

  const accountName = AdsApp.currentAccount().getName();
  const today = new Date().toLocaleDateString();
  const subject = `Google Ads Quality Score Report for ${accountName} - ${today}`;

  let body = `<p>Hello,</p>
              <p>Here is the Quality Score change report for your Google Ads account: <strong>${accountName}</strong>.</p>`;

  body += '<h2>Improved Quality Scores</h2>';
  if (improvedKeywords.length > 0) {
    body += createHtmlTable(improvedKeywords);
  } else {
    body += '<p>No keywords with improved Quality Score since the last report.</p>';
  }

  body += '<h2>Declined Quality Scores</h2>';
  if (declinedKeywords.length > 0) {
    body += createHtmlTable(declinedKeywords);
  } else {
    body += '<p>No keywords with declined Quality Score since the last report.</p>';
  }
  
  body += `<p style="margin-top: 25px; color: #888; font-size: 12px;"><em>This is an automated notification from a Google Ads Script. The data has been updated in your <a href="${SPREADSHEET_URL}">Google Sheet</a>.</em></p>`;

  MailApp.sendEmail({
    to: RECIPIENT_EMAIL,
    subject: subject,
    htmlBody: body,
  });
  
  console.log('Change report email sent successfully.');
}

/**
 * Creates an HTML table from an array of keyword change objects.
 * @param {Array<Object>} keywordsData The array of keyword data.
 * @return {string} An HTML table as a string.
 */
function createHtmlTable(keywordsData) {
  let table = `<table style="width: 100%; border-collapse: collapse; text-align: left;">
    <thead>
      <tr>
        <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;">Keyword</th>
        <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;">Ad Group</th>
        <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;">Campaign</th>
        <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: center;">Old QS</th>
        <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: center;">New QS</th>
      </tr>
    </thead>
    <tbody>`;

  keywordsData.forEach(item => {
    table += `
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd;">${item.keyword.getText()}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${item.keyword.getAdGroup().getName()}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${item.keyword.getCampaign().getName()}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${item.oldQS}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${item.newQS}</td>
      </tr>`;
  });

  table += '</tbody></table>';
  return table;
}