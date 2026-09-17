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
 * @name Daily Campaign Spend Report
 * @version 1.0
 * @author Google Ads Scripts Team
 *
 * This script generates and emails a daily report summarizing the spend and
 * performance of all active campaigns in a specific Google Ads account.
 */

// --- Your Settings (MUST BE CHANGED) ---

// The email address where you want to receive the daily report.
const RECIPIENT_EMAIL = 'thomas.wong@newimedia.com';

// The specific Google Ads Account ID this script should run on.
// This is a safety measure to prevent it from running on the wrong account.
// Format: "XXX-XXX-XXXX"
const ACCOUNT_ID_TO_CHECK = '959-935-7325';

// --- End of Settings ---

function main() {
  const currentAccountId = AdsApp.currentAccount().getCustomerId();
  
  // Safety check: Halt execution if the script is running on the wrong account.
  if (currentAccountId !== ACCOUNT_ID_TO_CHECK) {
    const errorMessage = `Error: Script is running on the wrong account.
      Expected: ${ACCOUNT_ID_TO_CHECK}
      Actual: ${currentAccountId}`;
    console.error(errorMessage);
    // Optional: Send an email to notify of the misconfiguration.
    sendEmail(
        `[Action Required] Ads Script Misconfiguration on Account ${currentAccountId}`,
        errorMessage
    );
    throw new Error(errorMessage);
  }

  const accountName = AdsApp.currentAccount().getName();
  const today = new Date();
  const formattedDate = `${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;

  console.log(`Generating daily spend report for ${accountName} (${currentAccountId}) for ${formattedDate}...`);

  const campaignIterator = AdsApp.campaigns()
    .withCondition('campaign.status = "ENABLED"')
    // The 'metrics.cost' field cannot be used in a .withCondition() filter.
    // We will filter for campaigns with spend inside the loop instead.
    .forDateRange('TODAY')
    .get();

  let totalCost = 0;
  let totalClicks = 0;
  let totalImpressions = 0;
  let campaignRows = '';

  while (campaignIterator.hasNext()) {
    const campaign = campaignIterator.next();
    const stats = campaign.getStatsFor('TODAY');
    
    const cost = stats.getCost();

    // Only include campaigns in the report if they have actually spent money today.
    if (cost > 0) {
      const clicks = stats.getClicks();
      const impressions = stats.getImpressions();
      const cpc = clicks > 0 ? (cost / clicks).toFixed(2) : '0.00';

      totalCost += cost;
      totalClicks += clicks;
      totalImpressions += impressions;

      campaignRows += `
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;">${campaign.getName()}</td>
          <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">$${cost.toFixed(2)}</td>
          <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${clicks}</td>
          <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">$${cpc}</td>
          <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${impressions}</td>
        </tr>
      `;
    }
  }
  
  console.log(`Total account cost for today: $${totalCost.toFixed(2)}`);

  const subject = `Google Ads Daily Spend Report for ${accountName} - ${formattedDate}`;
  const totalCpc = totalClicks > 0 ? (totalCost / totalClicks).toFixed(2) : '0.00';

  const emailBody = `
    <p>Hello,</p>
    <p>Here is the daily spending report for your Google Ads account: <strong>${accountName} (${currentAccountId})</strong> for <strong>${formattedDate}</strong>.</p>
    
    <h3 style="margin-bottom: 5px;">Account Summary</h3>
    <table style="width: 500px; border-collapse: collapse; text-align: left;">
      <tr>
        <th style="padding: 8px; border-bottom: 1px solid #ddd; width: 40%;">Total Spend:</th>
        <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: right;">$${totalCost.toFixed(2)}</td>
      </tr>
      <tr>
        <th style="padding: 8px; border-bottom: 1px solid #ddd;">Total Clicks:</th>
        <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: right;">${totalClicks}</td>
      </tr>
       <tr>
        <th style="padding: 8px; border-bottom: 1px solid #ddd;">Avg. CPC:</th>
        <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: right;">$${totalCpc}</td>
      </tr>
      <tr>
        <th style="padding: 8px; border-bottom: 1px solid #ddd;">Total Impressions:</th>
        <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: right;">${totalImpressions}</td>
      </tr>
    </table>

    <h3 style="margin-top: 25px; margin-bottom: 5px;">Campaign Breakdown</h3>
    ${
      campaignRows ?
      `<table style="width: 100%; border-collapse: collapse; text-align: left;">
        <thead>
          <tr>
            <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;">Campaign Name</th>
            <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">Spend</th>
            <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">Clicks</th>
            <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">Avg. CPC</th>
            <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">Impressions</th>
          </tr>
        </thead>
        <tbody>
          ${campaignRows}
        </tbody>
      </table>` :
      '<p>There was no ad spend for any enabled campaigns today.</p>'
    }

    <p style="margin-top: 25px; color: #888; font-size: 12px;"><em>This is an automated notification from a Google Ads Script.</em></p>
  `;

  sendEmail(subject, emailBody);
  console.log('Daily report email sent successfully.');
}

/**
 * Sends an email with a specified subject and body.
 * @param {string} subject The subject line of the email.
 * @param {string} body The HTML body content of the email.
 */
function sendEmail(subject, body) {
  MailApp.sendEmail({
    to: RECIPIENT_EMAIL,
    subject: subject,
    htmlBody: body,
  });
}