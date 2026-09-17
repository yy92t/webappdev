function main() {
  // ==========================================
  // CONFIGURATION
  // ==========================================
  const CONFIG = {
    recipientEmail: "thomas.wong@newimedia.com", // Enter your email here
    lookbackWindow: 30, // Days of history to analyze
    forecastDays: 7     // Days to predict into the future
  };
  
  // ==========================================
  // 1. GET DATA (Using GAQL)
  // ==========================================
  // We fetch cost_micros (Google stores money as millions, so $1 = 1,000,000 micros)
  const query = `
    SELECT 
      segments.date, 
      metrics.cost_micros 
    FROM 
      customer 
    WHERE 
      segments.date DURING LAST_${CONFIG.lookbackWindow}_DAYS 
    ORDER BY 
      segments.date ASC
  `;

  const rows = AdsApp.search(query);
  let xyPoints = [];
  let startDate = null;
  let count = 0;

  // Iterate through the report rows
  while (rows.hasNext()) {
    let row = rows.next();
    let dateString = row.segments.date; // Format YYYY-MM-DD
    let cost = row.metrics.costMicros / 1000000; // Convert to actual currency

    let currentDate = new Date(dateString);
    
    if (count === 0) startDate = currentDate;
    
    // Calculate "Day Number" (x-axis)
    let timeDiff = currentDate.getTime() - startDate.getTime();
    let dayIndex = Math.ceil(timeDiff / (1000 * 3600 * 24));

    xyPoints.push({ x: dayIndex, y: cost, date: dateString });
    count++;
  }

  if (xyPoints.length < 5) {
    Logger.log("Not enough data to predict trends. Account needs more history.");
    return;
  }

  // ==========================================
  // 2. CALCULATE TREND (Linear Regression)
  // ==========================================
  let n = xyPoints.length;
  let sumX = 0, sumY = 0, sumXY = 0, sumXX = 0;

  for (let p of xyPoints) {
    sumX += p.x;
    sumY += p.y;
    sumXY += (p.x * p.y);
    sumXX += (p.x * p.x);
  }

  let slope = (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX);
  let intercept = (sumY - slope * sumX) / n;

  // ==========================================
  // 3. GENERATE FORECAST DATA
  // ==========================================
  Logger.log("--- SPEND PREDICTION REPORT ---");
  
  let trendDirection = slope > 0 ? "INCREASING ðº" : "DECREASING ð»";
  let trendText = `Spend is ${trendDirection} by approx $${slope.toFixed(2)} per day.`;
  Logger.log(trendText);
  
  let forecastData = [];
  let totalForecast = 0;
  let lastDayIndex = xyPoints[xyPoints.length - 1].x;

  for (let i = 1; i <= CONFIG.forecastDays; i++) {
    let futureIndex = lastDayIndex + i;
    let predictedSpend = (slope * futureIndex) + intercept;
    
    // Safety: Spend cannot be negative
    if (predictedSpend < 0) predictedSpend = 0;
    
    // Create Date String for display
    let futureDate = new Date(startDate);
    futureDate.setDate(futureDate.getDate() + futureIndex);
    let dateStr = futureDate.toISOString().slice(0, 10);
    
    forecastData.push({
      date: dateStr,
      spend: predictedSpend
    });
    totalForecast += predictedSpend;
  }
  
  // ==========================================
  // 4. SEND EMAIL
  // ==========================================
  sendPredictionEmail(CONFIG.recipientEmail, trendText, forecastData, totalForecast, CONFIG);
}

// ------------------------------------------------
// HELPER FUNCTIONS
// ------------------------------------------------

function sendPredictionEmail(recipient, trendText, forecastData, totalForecast, config) {
  const accountName = AdsApp.currentAccount().getName();
  const cid = AdsApp.currentAccount().getCustomerId();
  
  // Build HTML Table
  let html = `
    <div style="font-family: Arial, sans-serif; color: #333;">
      <h2 style="color: #2980b9;">ð Google Ads Spend Prediction</h2>
      <p><strong>Account:</strong> ${accountName} (${cid})</p>
      <div style="background-color: #f8f9fa; padding: 15px; border-left: 5px solid #2980b9; margin-bottom: 20px;">
        <strong>Trend Analysis:</strong> ${trendText}<br>
        <em>Based on the last ${config.lookbackWindow} days of data.</em>
      </div>
      
      <h3>Forecast (Next ${config.forecastDays} Days)</h3>
      <table border="1" cellpadding="8" style="border-collapse: collapse; width: 100%; max-width: 600px; border: 1px solid #ddd;">
        <tr style="background-color: #2c3e50; color: white;">
          <th style="text-align: left;">Date</th>
          <th style="text-align: right;">Predicted Spend</th>
        </tr>
  `;
  
  for (let row of forecastData) {
    html += `
      <tr>
        <td>${row.date}</td>
        <td style="text-align: right;">$${row.spend.toFixed(2)}</td>
      </tr>
    `;
  }
  
  html += `
        <tr style="background-color: #ecf0f1; font-weight: bold;">
          <td>TOTAL</td>
          <td style="text-align: right;">$${totalForecast.toFixed(2)}</td>
        </tr>
      </table>
      <p style="font-size: 12px; color: #777; margin-top: 20px;">
        *Note: This is a linear projection based on past performance. Actual spend may vary due to seasonality or budget changes.
      </p>
    </div>
  `;
  
  MailApp.sendEmail({
    to: recipient,
    subject: `[Forecast] Spend Prediction for ${accountName}`,
    htmlBody: html
  });
  
  Logger.log(`ð§ Forecast email sent to ${recipient}`);
}