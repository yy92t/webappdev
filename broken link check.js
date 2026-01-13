function main() {
  // ==============================================
  // CONFIGURATION
  // ==============================================
  const CONFIG = {
    recipientEmail: "thomas.wong@newimedia.com",
    checkAds: true,
    checkKeywords: true, 
    badCodes: [404, 500, 502, 503],
    limit: 10000, // Increased limit to ensure we catch items with URLs even if we fetch some empties
    batchSize: 50 // Number of URLs to check in parallel
  };

  Logger.log("🚀 Starting Optimized Link Checker...");

  // ==============================================
  // 1. GATHER URLS (Using GAQL for Speed)
  // ==============================================
  let urlsToCheck = {}; 

  if (CONFIG.checkAds) {
    Logger.log("Fetching Ads...");
    // REMOVED: "AND ad_group_ad.ad.final_urls IS NOT NULL" (Not supported for lists)
    const query = `
      SELECT 
        ad_group.name, 
        ad_group_ad.ad.id, 
        ad_group_ad.ad.final_urls 
      FROM ad_group_ad 
      WHERE 
        campaign.status = 'ENABLED' 
        AND ad_group.status = 'ENABLED' 
        AND ad_group_ad.status = 'ENABLED' 
      LIMIT ${CONFIG.limit}
    `;
    
    const rows = AdsApp.search(query);
    while (rows.hasNext()) {
      let row = rows.next();
      let urls = row.adGroupAd.ad.finalUrls;
      
      // We filter empty URLs here in JS instead of GAQL
      if (urls && urls.length > 0) {
        let url = urls[0]; 
        urlsToCheck[url] = { 
          type: "Ad", 
          id: row.adGroupAd.ad.id, 
          location: row.adGroup.name 
        };
      }
    }
  }

  if (CONFIG.checkKeywords) {
    Logger.log("Fetching Keywords...");
    // UPDATED: Use 'ad_group_criterion' resource instead of 'keyword'
    const query = `
      SELECT 
        ad_group.name, 
        ad_group_criterion.keyword.text, 
        ad_group_criterion.final_urls 
      FROM ad_group_criterion 
      WHERE 
        campaign.status = 'ENABLED' 
        AND ad_group.status = 'ENABLED' 
        AND ad_group_criterion.status = 'ENABLED' 
        AND ad_group_criterion.type = 'KEYWORD'
      LIMIT ${CONFIG.limit}
    `;

    const rows = AdsApp.search(query);
    while (rows.hasNext()) {
      let row = rows.next();
      // Access fields via adGroupCriterion
      let urls = row.adGroupCriterion.finalUrls;
      
      // Filter here
      if (urls && urls.length > 0) {
        let url = urls[0];
        urlsToCheck[url] = { 
          type: "Keyword", 
          id: row.adGroupCriterion.keyword.text, 
          location: row.adGroup.name 
        };
      }
    }
  }

  // ==============================================
  // 2. CHECK URLS (Parallel Execution)
  // ==============================================
  let urlList = Object.keys(urlsToCheck);
  
  if (urlList.length === 0) {
     Logger.log("⚠️ No URLs found to check. Check if your ads/keywords have Final URLs set.");
     return;
  }

  Logger.log(`Checking ${urlList.length} unique URLs in batches...`);

  let brokenUrls = [];

  for (let i = 0; i < urlList.length; i += CONFIG.batchSize) {
    let batch = urlList.slice(i, i + CONFIG.batchSize);
    
    let requests = batch.map(url => ({
      url: url,
      muteHttpExceptions: true,
      followRedirects: true
    }));

    try {
      let responses = UrlFetchApp.fetchAll(requests);

      responses.forEach((response, index) => {
        let url = batch[index];
        let code = response.getResponseCode();
        
        if (CONFIG.badCodes.indexOf(code) > -1) {
          Logger.log(`❌ BROKEN (${code}): ${url}`);
          brokenUrls.push({
            url: url,
            code: code,
            type: urlsToCheck[url].type,
            ref: urlsToCheck[url].id,
            group: urlsToCheck[url].location
          });
        }
      });
      
    } catch (e) {
      Logger.log(`⚠️ Batch failed (${e.message}). Falling back to single checks for this batch.`);
      checkBatchIndividually(batch, urlsToCheck, CONFIG, brokenUrls);
    }
  }

  // ==============================================
  // 3. REPORTING
  // ==============================================
  if (brokenUrls.length > 0) {
    sendAlertEmail(CONFIG.recipientEmail, brokenUrls);
  } else {
    Logger.log("✅ No broken links found. No email sent.");
  }
}

// ------------------------------------------------
// HELPER FUNCTIONS
// ------------------------------------------------

function checkBatchIndividually(urlList, dataMap, config, brokenList) {
  for (let url of urlList) {
    try {
      let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      let code = response.getResponseCode();
      if (config.badCodes.indexOf(code) > -1) {
        Logger.log(`❌ BROKEN (Fallback) (${code}): ${url}`);
        brokenList.push({
          url: url,
          code: code,
          type: dataMap[url].type,
          ref: dataMap[url].id,
          group: dataMap[url].location
        });
      }
    } catch (e) {
      Logger.log(`❌ INVALID URL: ${url} - ${e.message}`);
    }
  }
}

function sendAlertEmail(recipient, errors) {
  const accountName = AdsApp.currentAccount().getName();
  const cid = AdsApp.currentAccount().getCustomerId();
  
  let html = `
    <div style="font-family: Arial, sans-serif;">
      <h2 style="color: #c0392b;">🚨 Broken Links Detected</h2>
      <p>Account: <strong>${accountName}</strong> (${cid})</p>
      <p>The following URLs are returning error codes:</p>
      <table border="1" cellpadding="8" style="border-collapse: collapse; width: 100%; border: 1px solid #ddd;">
        <tr style="background-color: #f2f2f2; text-align: left;">
          <th>Code</th>
          <th>Type</th>
          <th>Ref (ID/Kw)</th>
          <th>AdGroup</th>
          <th>URL</th>
        </tr>
  `;
  
  for (let err of errors) {
    html += `
      <tr>
        <td style="color: red; font-weight:bold;">${err.code}</td>
        <td>${err.type}</td>
        <td>${err.ref}</td>
        <td>${err.group}</td>
        <td style="word-break: break-all;">
          <a href="${err.url}" style="text-decoration: none; color: #3498db;">${err.url}</a>
        </td>
      </tr>
    `;
  }
  
  html += `</table></div>`;
  
  MailApp.sendEmail({
    to: recipient,
    subject: `[URGENT] Broken Links Found in ${accountName}`,
    htmlBody: html
  });
  
  Logger.log(`📧 Alert email sent to ${recipient}`);
}