function main() {
  // Replace with your actual email address
  const EMAIL_ADDRESS = "your.email@example.com"; 
  
  const today = new Date();
  const currentDay = today.getDate();
  const totalDaysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  
  const campaignIterator = AdsApp.campaigns()
      .withCondition("Status = ENABLED")
      .get();
      
  // Initialize HTML structure with new Conversions and CPA columns
  let htmlBody = `
    <h2 style="font-family: Arial, sans-serif; color: #333;">Google Ads: Performance & Pacing Dashboard</h2>
    <table style="font-family: Arial, sans-serif; border-collapse: collapse; width: 100%; font-size: 14px;">
      <tr style="background-color: #f2f2f2; text-align: left;">
        <th style="padding: 10px; border: 1px solid #ddd;">Campaign Name</th>
        <th style="padding: 10px; border: 1px solid #ddd;">Cost So Far</th>
        <th style="padding: 10px; border: 1px solid #ddd;">Conversions</th>
        <th style="padding: 10px; border: 1px solid #ddd;">CPA</th>
        <th style="padding: 10px; border: 1px solid #ddd;">Projected Spend</th>
        <th style="padding: 10px; border: 1px solid #ddd;">Monthly Target</th>
        <th style="padding: 10px; border: 1px solid #ddd;">Status</th>
      </tr>
  `;
  
  while (campaignIterator.hasNext()) {
    let campaign = campaignIterator.next();
    
    // Fetch stats for the current month
    let stats = campaign.getStatsFor("THIS_MONTH");
    let cost = stats.getCost();
    
    // Calculate Conversions and CPA
    let conversions = stats.getConversions();
    let cpa = conversions > 0 ? "$" + (cost / conversions).toFixed(2) : "N/A";
    
    // Calculate Pacing
    let dailyBudget = campaign.getBudget().getAmount();
    let monthlyBudgetLimit = dailyBudget * 30.4; 
    
    let runRate = cost / currentDay; 
    let projectedSpend = runRate * totalDaysInMonth;
    
    // Prevent division by zero if budget is missing
    let pacingPercentage = monthlyBudgetLimit > 0 ? (projectedSpend / monthlyBudgetLimit) * 100 : 0;
    
    // Determine status text and corresponding color badge
    let pacingStatus = "";
    let badgeColor = "";
    
    if (pacingPercentage > 105) {
      pacingStatus = "OVERSPENDING";
      badgeColor = "#d93025"; // Red
    } else if (pacingPercentage < 95) {
      pacingStatus = "UNDERSPENDING";
      badgeColor = "#f29900"; // Orange
    } else {
      pacingStatus = "ON TRACK";
      badgeColor = "#188038"; // Green
    }
    
    let statusBadge = <span style="background-color: ${badgeColor}; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 12px;">${pacingStatus} (${pacingPercentage.toFixed(1)}%)</span>;
    
    // Add row to HTML table, formatting conversions to 2 decimals for attribution models
    htmlBody += `
      <tr>
        <td style="padding: 10px; border: 1px solid #ddd;">${campaign.getName()}</td>
        <td style="padding: 10px; border: 1px solid #ddd;">$${cost.toFixed(2)}</td>
        <td style="padding: 10px; border: 1px solid #ddd;">${conversions.toFixed(2)}</td>
        <td style="padding: 10px; border: 1px solid #ddd;">${cpa}</td>
        <td style="padding: 10px; border: 1px solid #ddd;">$${projectedSpend.toFixed(2)}</td>
        <td style="padding: 10px; border: 1px solid #ddd;">$${monthlyBudgetLimit.toFixed(2)}</td>
        <td style="padding: 10px; border: 1px solid #ddd;">${statusBadge}</td>
      </tr>
    `;
  }
  
  htmlBody += "</table>";
  
  // Send the combined email
  MailApp.sendEmail({
    to: EMAIL_ADDRESS,
    subject: "Google Ads: Unified Performance & Pacing Dashboard",
    htmlBody: htmlBody
  });
  
  Logger.log("Combined visual dashboard email sent successfully.");
}