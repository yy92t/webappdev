/**
 * Google Ads Script: Export to Sheets & Visual Email Dashboard
 * 
 * Logs daily data to a Google Sheet and sends an HTML-formatted
 * email containing inline visual data bars for top campaigns.
 */


const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1floMf5SpjcsAi8OQpBX4R--bzcWuV8q8S6dF01iBuzg/edit?gid=0#gid=0'; 
const EMAIL_ADDRESS = 'thomas.wyy9@outlook.com';

function main() {
  // 1. Prepare the Google Sheet
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = spreadsheet.getActiveSheet();
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Campaign Name', 'Impressions', 'Clicks', 'Cost', 'Conversions']);
  }

  const accountTimeZone = AdsApp.currentAccount().getTimeZone();
  const dateToday = Utilities.formatDate(new Date(), accountTimeZone, 'yyyy-MM-dd');

  // 2. Fetch Campaign Data and Log to Sheet
  const campaignIterator = AdsApp.campaigns().withCondition("Status = ENABLED").get();
  
  let campaignData = [];
  let totalCost = 0, totalClicks = 0, totalImpr = 0, totalConv = 0;

  while (campaignIterator.hasNext()) {
    const campaign = campaignIterator.next();
    const stats = campaign.getStatsFor("TODAY");
    
    const cost = stats.getCost();
    const clicks = stats.getClicks();
    const impr = stats.getImpressions();
    const conv = stats.getConversions();

    // Log to Sheet
    sheet.appendRow([dateToday, campaign.getName(), impr, clicks, cost, conv]);
    
    // Save to array for email visualization
    if (cost > 0 || clicks > 0) { // Only visualize active campaigns
      campaignData.push({ name: campaign.getName(), cost: cost, clicks: clicks, conv: conv });
    }
    
    totalCost += cost; totalClicks += clicks; totalImpr += impr; totalConv += conv;
  }
  
  // 3. Prepare Visual Data for Email
  // Sort campaigns by highest cost
  campaignData.sort((a, b) => b.cost - a.cost);
  
  // Find max cost to scale the visual bars correctly (avoid dividing by zero)
  const maxCost = campaignData.length > 0 ? campaignData[0].cost : 1; 

  // 4. Build the HTML Email Body
  let htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px;">
      <h2 style="color: #4285F4;">Google Ads Daily Dashboard</h2>
      <p>Here is your performance overview for <strong>${dateToday}</strong>.</p>
      
      <table style="width: 100%; text-align: left; border-collapse: collapse; margin-bottom: 20px;">
        <tr style="background-color: #f8f9fa;">
          <th style="padding: 10px; border: 1px solid #ddd;">Total Spend</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Clicks</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Conversions</th>
        </tr>
        <tr>
          <td style="padding: 10px; border: 1px solid #ddd; font-size: 18px; font-weight: bold;">$${totalCost.toFixed(2)}</td>
          <td style="padding: 10px; border: 1px solid #ddd; font-size: 18px; font-weight: bold;">${totalClicks}</td>
          <td style="padding: 10px; border: 1px solid #ddd; font-size: 18px; font-weight: bold;">${totalConv}</td>
        </tr>
      </table>

      <h3>Top Campaigns by Spend</h3>
      <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
        <tr style="border-bottom: 2px solid #ddd;">
          <th style="padding: 8px 0; text-align: left; width: 40%;">Campaign</th>
          <th style="padding: 8px 0; text-align: left; width: 40%;">Spend Bar</th>
          <th style="padding: 8px 0; text-align: right; width: 20%;">Cost</th>
        </tr>`;

  // Add a row and visual bar for each campaign
  campaignData.forEach(c => {
    let barWidth = Math.max((c.cost / maxCost) * 100, 1); // Minimum 1% width
    htmlBody += `
        <tr style="border-bottom: 1px solid #eee;">
          <td style="padding: 8px 0; truncate: nowrap;">${c.name.substring(0, 30)}</td>
          <td style="padding: 8px 0;">
            <div style="width: 100%; background-color: #f1f3f4; border-radius: 3px;">
              <div style="width: ${barWidth}%; height: 12px; background-color: #4285F4; border-radius: 3px;"></div>
            </div>
          </td>
          <td style="padding: 8px 0; text-align: right;">$${c.cost.toFixed(2)}</td>
        </tr>`;
  });

  htmlBody += `
      </table>
      <br>
      <a href="${SPREADSHEET_URL}" style="display: inline-block; padding: 10px 15px; background-color: #34A853; color: white; text-decoration: none; border-radius: 4px;">View Full Google Sheet</a>
    </div>`;

  // 5. Send the HTML Email
  MailApp.sendEmail({
    to: EMAIL_ADDRESS,
    subject: `📊 Visual Daily Report: ${AdsApp.currentAccount().getName()}`,
    htmlBody: htmlBody
  });
  
  Logger.log(`Visual HTML email sent to ${EMAIL_ADDRESS}.`);
}
