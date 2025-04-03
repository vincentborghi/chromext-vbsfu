// background.js - v50

console.log("Background service worker started (v50 - Was v14 Send Log Messages).");

// --- Helper Function to Parse Date String ---
function parseDateString(dateString) {
    // console.log(`Background: Parsing date: "${dateString}"`);
    if (!dateString) return null;
    let match = dateString.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/); // DD/MM/YYYY HH:MM
    if (match) {
        try {
            const day = parseInt(match[1]); const month = parseInt(match[2]) - 1; const year = parseInt(match[3]);
            const hour = parseInt(match[4]); const minute = parseInt(match[5]);
            if (year > 1970 && month >= 0 && month < 12 && day >= 1 && day <= 31 && hour >= 0 && hour < 24 && minute >= 0 && minute < 60) {
                 const dateObject = new Date(Date.UTC(year, month, day, hour, minute));
                 if (!isNaN(dateObject.getTime())) return dateObject;
            }
        } catch (e) { /* continue */ }
    }
    const parsedFallback = Date.parse(dateString);
    if (!isNaN(parsedFallback)) return new Date(parsedFallback);
    console.warn(`Background: Could not parse date format "${dateString}"`);
    return null;
}

// --- Function to process all items (Notes or Emails) via temporary tabs ---
// ** Added senderTabId parameter **
async function fetchAllDetailsViaTabs(itemsToFetch, itemType, senderTabId) {
    console.log(`Background: Starting tab automation for ${itemsToFetch.length} ${itemType}(s) for sender tab ${senderTabId}.`);
    const resultsMap = {};
    const scraperScript = itemType === 'Note' ? 'note_scraper.js' : 'email_scraper.js';
    let currentItemIndex = 0;

    // Process sequentially
    for (const itemInfo of itemsToFetch) {
        currentItemIndex++;
        const itemUrl = itemInfo.url;
        // ** Send log message back to original tab BEFORE processing **
        if (senderTabId) {
             console.log(`Background: Sending log message for ${itemUrl} to tab ${senderTabId}`);
             // Send message without waiting for response, catch potential error if tab closed
             chrome.tabs.sendMessage(senderTabId, {
                 action: "logUrlProcessing",
                 url: itemUrl,
                 itemType: itemType,
                 index: currentItemIndex,
                 total: itemsToFetch.length
             }).catch(err => console.warn(`Background: Failed to send log message to tab ${senderTabId}: ${err.message}`));
        } else {
            console.warn("Background: Cannot send log message, senderTabId is missing.")
        }

        console.log(`Background: Processing ${itemType} ${currentItemIndex}/${itemsToFetch.length}: ${itemUrl}`); // Keep existing log too
        let tempTab = null; let scrapeResult = null;

        try {
            console.log(`Background: Creating inactive tab for ${itemUrl}`);
            tempTab = await chrome.tabs.create({ url: itemUrl, active: false });
            const tempTabId = tempTab.id;
            if (!tempTabId) throw new Error("Failed to create temp tab.");
            console.log(`Background: Created tab ID: ${tempTabId}. Waiting for load...`);

            // Wait for tab load
            await new Promise((resolve, reject) => { /* ... same tab load wait logic ... */
                const timeout = setTimeout(() => { chrome.tabs.onUpdated.removeListener(tabUpdateListener); reject(new Error(`Timeout waiting for tab ${tempTabId} to load`)); }, 30000);
                const tabUpdateListener = (tabId, changeInfo, tab) => {
                    if (tabId === tempTabId) {
                         if (changeInfo.status === 'complete') { clearTimeout(timeout); console.log(`Tab ${tempTabId} loaded.`); chrome.tabs.onUpdated.removeListener(tabUpdateListener); resolve(); }
                         else if (changeInfo.status === 'error' || tab?.url?.includes('error') || tab?.url?.includes('login')) { clearTimeout(timeout); console.warn(`Tab ${tempTabId} failed load.`); chrome.tabs.onUpdated.removeListener(tabUpdateListener); reject(new Error(`Tab ${tempTabId} failed load.`)); }
                    }
                }; chrome.tabs.onUpdated.addListener(tabUpdateListener);
            });

             // Setup listener for scrape results
             const resultPromise = new Promise((resolve, reject) => { /* ... same result wait logic ... */
                 const timeout = setTimeout(() => { chrome.runtime.onMessage.removeListener(scrapeResultListener); reject(new Error(`Timeout waiting for scrape result from tab ${tempTabId}`)); }, 20000);
                 const expectedResponseType = itemType === 'Note' ? 'noteScrapeResult' : 'emailScrapeResult';
                 const scrapeResultListener = (message, sender, sendResponse) => {
                     if (message.type === expectedResponseType && sender.tab?.id === tempTabId) { clearTimeout(timeout); console.log(`Received '${expectedResponseType}' result.`); chrome.runtime.onMessage.removeListener(scrapeResultListener); resolve(message); return false; }
                 }; chrome.runtime.onMessage.addListener(scrapeResultListener);
             });

            // Inject scraping script
            console.log(`Background: Injecting ${scraperScript} into tab ${tempTabId}`);
            await chrome.scripting.executeScript({ target: { tabId: tempTabId }, files: [scraperScript] });
            console.log(`Background: Script injected. Waiting for result...`);

            // Wait for result
            scrapeResult = await resultPromise;

            // Process result
            if (scrapeResult) {
                 console.log(`Background: Processing result for ${itemUrl}.`);
                 const parsedDate = parseDateString(itemInfo.dateStr); // Use date from TABLE
                 if (itemType === 'Note') {
                     resultsMap[itemUrl] = { description: scrapeResult.description || '', dateObject: parsedDate };
                 } else { // Email
                      resultsMap[itemUrl] = { description: scrapeResult.bodyHTML || '', dateObject: parsedDate, subject: scrapeResult.subject || '', from: scrapeResult.from || '', to: scrapeResult.to || '' };
                 }
            } else { throw new Error("Scrape result was null/undefined."); }
        } catch (error) {
            console.error(`Background: Error processing ${itemType} ${itemUrl} in tab ${tempTab?.id}:`, error);
            resultsMap[itemUrl] = { description: `Error: ${error.message}`, dateObject: null };
        } finally {
            if (tempTab?.id) { // Close tab
                try { console.log(`Background: Closing temp tab ${tempTab.id}`); await chrome.tabs.remove(tempTab.id); }
                catch (closeError) { console.warn(`Background: Error closing temp tab ${tempTab.id}:`, closeError.message); }
            }
        }
    } // End for loop
    console.log(`Background: Finished processing all ${itemsToFetch.length} ${itemType}(s) via tabs.`);
    return resultsMap;
}


// --- Main Listener ---
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  // Handler for opening the results tab
  if (message.action === "openFullViewTab" && message.htmlContent) {
    console.log("Background: Received request to open results tab (using data: URL).");
    const url = 'data:text/html;charset=UTF-8,' + encodeURIComponent(message.htmlContent);
    console.log("Background: Created data URL.");
    chrome.tabs.create({ url: url }, (newTab) => { /* ... error handling ... */
        if (chrome.runtime.lastError) console.error("Background: Error creating tab:", chrome.runtime.lastError.message);
        else if (newTab) console.log("Background: Tab creation initiated.");
        else console.error("Background: Tab creation failed, newTab is null.");
    });
     return false; // Sync
  }

  // Handler for fetching Note OR Email details
  if (message.action === "fetchItemDetails" && message.items && message.items.length > 0) {
    const itemType = message.items[0].type;
    const itemsToFetch = message.items;
    // ** Get sender tab ID **
    const senderTabId = sender.tab?.id;

    if (!senderTabId) {
         console.error("Background: Could not get sender tab ID for fetchItemDetails request.");
         sendResponse({ status: "error", message: "Could not identify sender tab."});
         return false; // Sync error response
    }

    console.log(`Background: Received request 'fetchItemDetails' for ${itemsToFetch.length} item(s) of type '${itemType}' from tab ${senderTabId}.`);

    // ** Pass senderTabId to the processing function **
    fetchAllDetailsViaTabs(itemsToFetch, itemType, senderTabId)
      .then(resultsMap => {
        console.log(`Background: Finished ${itemType} batch. Sending results back.`);
        sendResponse({ status: "success", details: resultsMap });
      })
      .catch(error => {
        console.error(`Background: Error in fetchAllDetailsViaTabs chain for ${itemType}:`, error);
        sendResponse({ status: "error", message: error.message || `Unknown error during ${itemType} tab automation.` });
      });
    return true; // Async
  }

  // Default return false if no condition matched
  return false;

}); // End of addListener

console.log("Background: Listeners attached.");
