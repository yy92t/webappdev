function main() {
  // --- CONFIGURATION ---
  var TARGET_MONTHLY_BUDGET = 5000; 
  var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1cn-oKNd2qevSCYzpYyD9r3AK9JIuMJDub1v2khldWLg/edit?gid=1223119768#gid=1223119768'; // The Sheet where the chart lives
  var ALERT_EMAIL = 'thomas.wong@newimedia.com';
  var ALERT_THRESHOLD = 0.10; // 0.10 means it will email you if you are off pace by more than 10%
  // ---------------------

  var account = AdsApp.currentAccount();
  var timeZone = account.getTimeZone();
  var today = new Date();
  var currentDay = today.getDate();
  var daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  
  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  
  // Create or reset a tab for the current month
  var monthStr = Utilities.formatDate(today, timeZone, 'MMMM yyyy');
  var sheetName = "Pacing Visual - " + monthStr;
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
    var charts = sheet.getCharts();
    for (var c = 0; c < charts.length; c++) {
      sheet.removeChart(charts[c]);
    }
  }

  // --- 1. DATA EXTRACTION ---
  var query = "SELECT segments.date, metrics.cost_micros " +
              "FROM customer " +
              "WHERE segments.date DURING THIS_MONTH " +
              "ORDER BY segments.date ASC";
              
  var report = AdsApp.report(query);
  var rows = report.rows();
  var spendByDate = {};
  var totalSpendToDate = 0;
  
  while (rows.hasNext()) {
    var row = rows.next();
    var dailyCost = parseFloat(row['metrics.cost_micros']) / 1000000;
    spendByDate[row['segments.date']] = dailyCost; 
    totalSpendToDate += dailyCost; // Track total spend for the email alert
  }

  // --- 2. PACING CALCULATIONS ---
  var data = [];
  data.push(["Day of Month", "Actual Cumulative Spend", "Ideal Pacing Target"]);

  var cumulativeSpend = 0;
  var dailyTarget = TARGET_MONTHLY_BUDGET / daysInMonth;

  for (var i = 1; i <= daysInMonth; i++) {
    var idealSpend = dailyTarget * i;
    var loopDate = new Date(today.getFullYear(), today.getMonth(), i);
    var dateStr = Utilities.formatDate(loopDate, timeZone, 'yyyy-MM-dd');
    var actualToLog = "";
    
    if (spendByDate[dateStr] !== undefined) {
      cumulativeSpend += spendByDate[dateStr];
      actualToLog = cumulativeSpend;
    } else if (i <= currentDay) {
      actualToLog = cumulativeSpend;
    } else {
      actualToLog = ""; 
    }

    data.push(["Day " + i, actualToLog, idealSpend]);
  }

  // --- 3. GENERATE VISUALS (SPREADSHEET) ---
  var range = sheet.getRange(1, 1, data.length, data[0].length);
  range.setValues(data);
  
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#f3f3f3");
  sheet.getRange(2, 2, data.length - 1, 2).setNumberFormat("$#,##0.00");

  var chartBuilder = sheet.newChart();
  chartBuilder.addRange(sheet.getRange(1, 1, data.length, 1))
              .addRange(sheet.getRange(1, 2, data.length, 1))
              .addRange(sheet.getRange(1, 3, data.length, 1))
              .setChartType(Charts.ChartType.LINE)
              .setPosition(2, 5, 0, 0) 
              .setOption('title', 'Budget Pacing: ' + monthStr)
              .setOption('width', 800)
              .setOption('height', 500)
              .setOption('series', {
                 0: {color: '#4285F4', lineWidth: 4},          
                 1: {color: '#DB4437', lineDashStyle: [4, 4]}  
               })
              .setOption('vAxes', {
                 0: {title: 'Cumulative Spend ($)', viewWindow: {min: 0, max: TARGET_MONTHLY_BUDGET * 1.1}}
               })
              .setOption('legend', {position: 'bottom'});
               
  sheet.insertChart(chartBuilder.build());

  // --- 4. EMAIL LOGIC & SENDING ---
  
  // Calculate where we should be TODAY
  var expectedSpendToday = dailyTarget * currentDay;
  var spendDifference = totalSpendToDate - expectedSpendToday;
  var pacingPercentage = spendDifference / expectedSpendToday; 

  // Check if pacing exceeds the threshold (e.g., more than 10% off pace in either direction)
  if (Math.abs(pacingPercentage) > ALERT_THRESHOLD) {
    
    // Determine the status and color for the email
    var status = (spendDifference > 0) ? "OVERPACING" : "UNDERPACING";
    var color = (spendDifference > 0) ? "#DB4437" : "#F4B400"; // Red for over, Yellow for under
    
    var emailSubject = "Budget Alert: " + account.getName() + " is " + status;
    
    // Construct an HTML Email Body
    var htmlBody = "<h2>Budget Pacing Alert</h2>" +
                   "<p>Your Google Ads account <strong>" + account.getName() + "</strong> is currently off pace.</p>" +
                   "<table style='border-collapse: collapse; width: 300px;'>" +
                   "<tr><td style='padding: 8px; border: 1px solid #ddd;'>Target Monthly Budget</td><td style='padding: 8px; border: 1px solid #ddd;'>$" + TARGET_MONTHLY_BUDGET.toFixed(2) + "</td></tr>" +
                   "<tr><td style='padding: 8px; border: 1px solid #ddd;'>Expected Spend by Today</td><td style='padding: 8px; border: 1px solid #ddd;'>$" + expectedSpendToday.toFixed(2) + "</td></tr>" +
                   "<tr><td style='padding: 8px; border: 1px solid #ddd; font-weight: bold;'>Actual Spend to Date</td><td style='padding: 8px; border: 1px solid #ddd; font-weight: bold; color: " + color + ";'>$" + totalSpendToDate.toFixed(2) + "</td></tr>" +
                   "</table>" +
                   "<p>Status: <span style='color: " + color + "; font-weight: bold;'>" + status + " by " + Math.abs((pacingPercentage * 100)).toFixed(1) + "%</span></p>" +
                   "<br><hr>" +
                   "<p><strong><a href='" + SPREADSHEET_URL + "' target='_blank'>Click here to view your visual pacing dashboard</a></strong></p>";
                   
    // Send the email
    MailApp.sendEmail({
      to: ALERT_EMAIL,
      subject: emailSubject,
      htmlBody: htmlBody
    });
    
    Logger.log("Pacing alert triggered! Email sent to " + ALERT_EMAIL);
  } else {
    Logger.log("Account is pacing properly. No email sent.");
  }
}
