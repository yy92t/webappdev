/**
 * Google Ads Script: Monthly Summary with Campaign Types
 * 
 * Uses GAQL to fetch all campaign types (PMax, Search, Demand Gen, etc.)
 * and emails a styled HTML summary for the previous month.
 */

const EMAIL_ADDRESS = 'thomas.wyy9@outlook.com';
const TOP_CAMPAIGNS_LIMIT = 25; 

function main() {
  const account = AdsApp.currentAccount();
  const accountName = account.getName();
  
  // 1. Fetch Account Totals for Last Month
  const accountStats = account.getStatsFor("LAST_MONTH");
  const totalCost = accountStats.getCost();
  const totalClicks = accountStats.getClicks();
  const totalConv = accountStats.getConversions();

  // 2. Build the GAQL Query to get campaigns, spend, and channel type
  const query = `
    SELECT 
      campaign.name, 
      campaign.status, 
      campaign.advertising_channel_type,
      metrics.cost_micros, 
      metrics.clicks, 
      metrics.conversions 
    FROM campaign 
    WHERE campaign.status IN ('ENABLED', 'PAUSED') 
      AND segments.date DURING LAST_MONTH
    ORDER BY metrics.cost_micros DESC
    LIMIT ${TOP_CAMPAIGNS_LIMIT}
  `;
  
  const report = AdsApp.search(query);
  let campaignData = [];

  // 3. Process the Data
  while (report.hasNext()) {
    const row = report.next();
    
    // GAQL returns cost in "micros" (millionths of a dollar), so we divide by 1,000,000
    const cost = row.metrics.costMicros / 1000000;
    
    if (cost > 0) {
      campaignData.push({ 
        name: row.campaign.name, 
        status: row.campaign.status === 'ENABLED' ? 'Enabled' : 'Paused',
        type: formatCampaignType(row.campaign.advertisingChannelType),
        cost: cost, 
        clicks: row.metrics.clicks, 
        conv: row.metrics.conversions 
      });
    }
  }
  
  const formatNum = (num) => Number(num).toLocaleString('en-US');

  // 4. Build the HTML Email Body
  let htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 700px;">
      <h2 style="color: #4285F4;">Monthly Campaign Review</h2>
      <p>Here is the performance overview for <strong>${accountName}</strong> during the previous month.</p>
      
      <table style="width: 100%; text-align: left; border-collapse: collapse; margin-bottom: 25px;">
        <tr style="background-color: #f8f9fa;">
          <th style="padding: 12px; border: 1px solid #ddd;">Total Spend</th>
          <th style="padding: 12px; border: 1px solid #ddd;">Total Clicks</th>
          <th style="padding: 12px; border: 1px solid #ddd;">Total Conversions</th>
        </tr>
        <tr>
          <td style="padding: 12px; border: 1px solid #ddd; font-size: 20px; font-weight: bold;">$${formatNum(totalCost.toFixed(2))}</td>
          <td style="padding: 12px; border: 1px solid #ddd; font-size: 20px; font-weight: bold;">${formatNum(totalClicks)}</td>
          <td style="padding: 12px; border: 1px solid #ddd; font-size: 20px; font-weight: bold;">${formatNum(totalConv.toFixed(1))}</td>
        </tr>
      </table>

      <h3>Top Campaigns Last Month</h3>
      <table style="width: 100%; border-collapse: collapse; font-size: 13px; text-align: left;">
        <tr style="border-bottom: 2px solid #ddd; background-color: #f1f3f4;">
          <th style="padding: 10px;">Campaign Name</th>
          <th style="padding: 10px;">Type</th>
          <th style="padding: 10px;">Status</th>
          <th style="padding: 10px; text-align: right;">Spend</th>
          <th style="padding: 10px; text-align: right;">Conv.</th>
        </tr>`;

  // Add rows for top campaigns
  campaignData.forEach(c => {
    const statusColor = c.status === "Enabled" ? "#34A853" : "#9AA0A6";
    const typeBadge = `<span style="background:#e8f0fe; color:#1967d2; padding:3px 6px; border-radius:4px; font-size:11px;">${c.type}</span>`;
    
    htmlBody += `
        <tr style="border-bottom: 1px solid #eee;">
          <td style="padding: 10px;">${c.name}</td>
          <td style="padding: 10px;">${typeBadge}</td>
          <td style="padding: 10px;"><span style="color: ${statusColor}; font-weight: bold;">${c.status}</span></td>
          <td style="padding: 10px; text-align: right;">$${formatNum(c.cost.toFixed(2))}</td>
          <td style="padding: 10px; text-align: right;">${formatNum(c.conv.toFixed(1))}</td>
        </tr>`;
  });

  htmlBody += `
      </table>
    </div>`;

  // 5. Send Email
  MailApp.sendEmail({
    to: EMAIL_ADDRESS,
    subject: `📅 Last Month's Performance: ${accountName}`,
    htmlBody: htmlBody
  });
  
  Logger.log(`Monthly report sent to ${EMAIL_ADDRESS}.`);
}

/**
 * Helper function to turn GAQL enums into readable campaign types
 */
function formatCampaignType(typeEnum) {
  const types = {
    'SEARCH': 'Search',
    'DISPLAY': 'Display',
    'PERFORMANCE_MAX': 'Performance Max',
    'VIDEO': 'Video',
    'SHOPPING': 'Shopping',
    'DEMAND_GEN': 'Demand Gen',
    'DISCOVERY': 'Discovery',
    'SMART': 'Smart',
    'HOTEL': 'Hotel',
    'LOCAL': 'Local'
  };
  return types[typeEnum] || typeEnum; // Falls back to raw enum if unmapped
}
