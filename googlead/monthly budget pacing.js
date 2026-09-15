function main() {
  // --- CONFIGURATION ---
  var TARGET_MONTHLY_BUDGET = 5000; 
  var ALERT_EMAIL = 'thomas.wong@newimedia.com';
  var THRESHOLD_PERCENTAGE = 0.15; // Alert if off pace by more than 15%
  // ---------------------

  var account = AdsApp.currentAccount();
  var costThisMonth = account.getStatsFor('THIS_MONTH').getCost();

  var today = new Date();
  var currentDay = today.getDate();
  var daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();

  // Calculate where spend should be by today
  var dailyTarget = TARGET_MONTHLY_BUDGET / daysInMonth;
  var expectedSpendToDate = dailyTarget * currentDay;

  Logger.log('Target Monthly Budget: $' + TARGET_MONTHLY_BUDGET);
  Logger.log('Current Spend: $' + costThisMonth);
  Logger.log('Expected Spend by Today: $' + expectedSpendToDate.toFixed(2));

  // Determine if pacing is off
  var difference = costThisMonth - expectedSpendToDate;
  var percentOff = Math.abs(difference / expectedSpendToDate);

  if (percentOff > THRESHOLD_PERCENTAGE) {
    var status = difference > 0 ? "OVERPACING" : "UNDERPACING";
    var body = "Your account is " + status + ".\n\n" +
               "Current Spend: $" + costThisMonth + "\n" +
               "Expected Spend: $" + expectedSpendToDate.toFixed(2);
               
    MailApp.sendEmail(ALERT_EMAIL, "Budget Alert: " + status, body);
    Logger.log('Alert sent: ' + status);
  } else {
    Logger.log('Budget is on track.');
  }
}
