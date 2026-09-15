// =======================================================================
// CONFIGURATION
// =======================================================================

// 1. Enter the email address where you want to receive the alerts
var EMAIL_ADDRESS = "thomas.wong@newimedia.com"; 

// 2. Set to true if you want the script to automatically PAUSE broken ads.
// Set to false if you just want to be emailed without pausing anything.
var PAUSE_BROKEN_ADS = false; 

// =======================================================================

function main() {
  Logger.log("Starting Link Checker...");
  
  // Grab all active ads in active ad groups and campaigns
  var adIterator = AdsApp.ads()
    .withCondition("CampaignStatus = ENABLED")
    .withCondition("AdGroupStatus = ENABLED")
    .withCondition("Status = ENABLED")
    .get();

  var brokenUrls = [];
  var checkedUrls = {};
  var adsChecked = 0;

  while (adIterator.hasNext()) {
    var ad = adIterator.next();
    var url = ad.urls().getFinalUrl();

    if (!url) continue; // Skip if the ad doesn't have a final URL
    adsChecked++;

    // Clean the URL: Remove Google Ads ValueTrack parameters (like {lpurl}) 
    // because they will cause the fetch function to fail.
    var cleanUrl = url.split('{')[0]; 

    // Only ping the URL if we haven't already checked it during this run
    if (checkedUrls[cleanUrl] === undefined) {
      try {
        // muteHttpExceptions prevents the script from crashing when it hits a 404
        var response = UrlFetchApp.fetch(cleanUrl, { muteHttpExceptions: true });
        var statusCode = response.getResponseCode();
        checkedUrls[cleanUrl] = statusCode;
      } catch (e) {
        // Catch DNS errors or invalid URL structures
        checkedUrls[cleanUrl] = "Error: " + e.message;
      }
    }

    var status = checkedUrls[cleanUrl];
    
    // Check if the response code is an error (400+) or a caught exception
    if ((typeof status === 'number' && status >= 400) || typeof status === 'string') {
      brokenUrls.push({
        url: cleanUrl,
        status: status,
        campaignName: ad.getCampaign().getName(),
        adGroupName: ad.getAdGroup().getName()
      });

      if (PAUSE_BROKEN_ADS) {
        ad.pause();
      }
    }
  }

  Logger.log("Total active ads scanned: " + adsChecked);
  Logger.log("Total unique URLs tested: " + Object.keys(checkedUrls).length);

  // Send an email if any broken links were found
  if (brokenUrls.length > 0) {
    Logger.log("Found " + brokenUrls.length + " broken ads. Sending email...");
    sendEmail(brokenUrls);
  } else {
    Logger.log("✅ Success! All active ad URLs are working perfectly.");
  }
}

function sendEmail(brokenUrls) {
  var subject = "🚨 Google Ads Alert: Broken Links Detected";
  var body = "The following active ads point to broken URLs:\n\n";

  for (var i = 0; i < brokenUrls.length; i++) {
    body += "URL: " + brokenUrls[i].url + "\n";
    body += "HTTP Error Code: " + brokenUrls[i].status + "\n";
    body += "Campaign: " + brokenUrls[i].campaignName + "\n";
    body += "Ad Group: " + brokenUrls[i].adGroupName + "\n";
    body += "----------------------------------------\n";
  }

  if (PAUSE_BROKEN_ADS) {
    body += "\nNote: PAUSE_BROKEN_ADS is set to true. These ads have been automatically paused in your account.";
  } else {
    body += "\nNote: PAUSE_BROKEN_ADS is set to false. These ads are STILL RUNNING and wasting money. Please fix them immediately.";
  }

  MailApp.sendEmail(EMAIL_ADDRESS, subject, body);
}
