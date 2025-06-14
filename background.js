// background.js - v50 + Initiate Handler + CONCURRENT PROCESSING

console.log("Background service worker started (v50 + Concurrent Processing).");

// --- Helper Function to Parse Date String --- (No changes here)
function parseDateString(dateString) {
    // console.log(`Background: Parsing date: "${dateString}"`); // Optional debug log
    if (!dateString) return null; // Handle null or empty input

    // Regex specifically for DD/MM/YYYY HH:MM format
    let match = dateString.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);
    if (match) {
        try {
            // Extract components from regex match
            const day = parseInt(match[1]);
            const month = parseInt(match[2]) - 1; // JavaScript months are 0-indexed (0-11)
            const year = parseInt(match[3]);
            const hour = parseInt(match[4]);
            const minute = parseInt(match[5]);

            // Basic validation of the extracted date/time components
            if (year > 1970 && month >= 0 && month < 12 && day >= 1 && day <= 31 && hour >= 0 && hour < 24 && minute >= 0 && minute < 60) {
                 // Create a Date object using UTC to avoid timezone inconsistencies
                 const dateObject = new Date(Date.UTC(year, month, day, hour, minute));
                 // Final check: ensure the created Date object is valid (e.g., not Feb 30th)
                 if (!isNaN(dateObject.getTime())) {
                     return dateObject; // Return the valid Date object
                 }
            }
        } catch (e) {
            // Log error if parsing components failed
            console.error(`Background: Error parsing matched date parts "${dateString}":`, e);
            // Continue to fallback below
        }
    }

    // Fallback: Attempt parsing using Date.parse()
    const parsedFallback = Date.parse(dateString);
    if (!isNaN(parsedFallback)) {
        console.warn(`Background: Used Date.parse fallback for "${dateString}"`);
        return new Date(parsedFallback); // Return Date object from fallback
    }

    // If all parsing attempts fail
    console.warn(`Background: Could not parse date format "${dateString}"`);
    return null; // Return null if parsing failed
}

// --- MODIFIED: Function to process all items concurrently with a limit ---
async function fetchAllDetailsViaTabs(itemsToFetch, itemType, senderTabId) {
    console.log(`Background: Starting CONCURRENT tab automation for ${itemsToFetch.length} ${itemType}(s) from sender tab ${senderTabId}.`);
    const resultsMap = {};
    const scraperScript = itemType === 'Note' ? 'note_scraper.js' : 'email_scraper.js';

    // *** NEW: Concurrency configuration ***
    const CONCURRENCY_LIMIT = 4; // Process 4 tabs at a time. Adjust as needed.
    let activePromises = [];
    let itemIndex = 0;

    // Create a pool of promises that will be executed concurrently
    for (const itemInfo of itemsToFetch) {
        // Wrap the processing logic for a single item in a function that returns a promise
        const task = async () => {
            const currentIndex = ++itemIndex; // Use a dedicated counter for logging
            const itemUrl = itemInfo.url;

            // --- Send progress log message (No changes needed here) ---
            if (senderTabId) {
                 try {
                     await chrome.tabs.sendMessage(senderTabId, {
                         action: "logUrlProcessing",
                         url: itemUrl, itemType: itemType,
                         index: currentIndex, total: itemsToFetch.length
                     });
                 } catch (err) {
                     console.warn(`Background: Failed to send log message to tab ${senderTabId} (may be closed): ${err.message}`);
                 }
            }

            console.log(`Background: Processing ${itemType} ${currentIndex}/${itemsToFetch.length}: ${itemUrl}`);
            let tempTab = null;
            let scrapeResult = null;

            try {
                // The logic for a single tab remains largely the same
                tempTab = await chrome.tabs.create({ url: itemUrl, active: false });
                const tempTabId = tempTab.id;
                if (!tempTabId) throw new Error("Failed to create temp tab.");

                // Wait for tab to load (using the same robust promise-based listener)
                await new Promise((resolve, reject) => {
                    const timeout = setTimeout(() => {
                        chrome.tabs.onUpdated.removeListener(listener);
                        reject(new Error(`Timeout waiting for tab ${tempTabId} to load`));
                    }, 30000);
                    const listener = (tabId, changeInfo, tab) => {
                        if (tabId === tempTabId && changeInfo.status === 'complete') {
                            clearTimeout(timeout);
                            chrome.tabs.onUpdated.removeListener(listener);
                            resolve();
                        }
                    };
                    chrome.tabs.onUpdated.addListener(listener);
                });

                // Set up listener for scrape result
                const resultPromise = new Promise((resolve, reject) => {
                     const timeout = setTimeout(() => {
                        chrome.runtime.onMessage.removeListener(listener);
                        reject(new Error(`Timeout waiting for scrape result from tab ${tempTabId}`));
                    }, 20000);
                    const expectedResponseType = itemType === 'Note' ? 'noteScrapeResult' : 'emailScrapeResult';
                    const listener = (message, sender) => {
                        if (sender.tab?.id === tempTabId && message.type === expectedResponseType) {
                            clearTimeout(timeout);
                            chrome.runtime.onMessage.removeListener(listener);
                            resolve(message);
                            return false;
                        }
                        return false;
                    };
                    chrome.runtime.onMessage.addListener(listener);
                });

                await chrome.scripting.executeScript({ target: { tabId: tempTabId }, files: [scraperScript] });
                scrapeResult = await resultPromise;

                // Process result and store in the map (thread-safe as keys are unique URLs)
                const parsedDate = parseDateString(itemInfo.dateStr);
                if (itemType === 'Note') {
                    resultsMap[itemUrl] = {
                        description: scrapeResult.description || '', dateObject: parsedDate, isPublic: scrapeResult.isPublic
                    };
                } else { // Email
                    resultsMap[itemUrl] = {
                        description: scrapeResult.bodyHTML || '', dateObject: parsedDate, subject: scrapeResult.subject || '',
                        from: scrapeResult.from || '', to: scrapeResult.to || '', isPublic: null
                    };
                }
            } catch (error) {
                console.error(`Background: Error processing ${itemType} ${itemUrl} in tab ${tempTab?.id}:`, error);
                resultsMap[itemUrl] = {
                    description: `Error processing item: ${error.message}`,
                    dateObject: parseDateString(itemInfo.dateStr), isPublic: null, subject: '[Error]', from: '[Error]', to: '[Error]'
                };
            } finally {
                if (tempTab?.id) {
                    try { await chrome.tabs.remove(tempTab.id); } catch (e) {
                        console.warn(`Background: Error closing temp tab ${tempTab.id}:`, e.message);
                    }
                }
            }
        };

        // Add the task promise to the pool
        activePromises.push(task());

        // If the pool is full, wait for one promise to finish before adding the next
        if (activePromises.length >= CONCURRENCY_LIMIT) {
            await Promise.race(activePromises); // Wait for the FASTEST promise to complete
            activePromises = activePromises.filter(p => 'pending' in p ? p.pending : true); // Crude way to remove completed, better libraries exist but this is simple
        }
    }

    // Wait for all remaining promises in the pool to complete
    await Promise.all(activePromises);

    console.log(`Background: Finished processing all ${itemsToFetch.length} ${itemType}(s) via tabs.`);
    return resultsMap;
}


// --- Main Listener for Messages --- (No changes here)
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log(`Background: Received message action="${message.action}" from sender:`, sender.tab?.id, sender.url);

  if (message.action === "initiateGenerateFullCaseView") {
    if (sender.tab?.id) {
        const targetTabId = sender.tab.id;
        chrome.tabs.sendMessage(targetTabId, { action: "generateFullCaseView" }, (response) => {
            if (chrome.runtime.lastError) {
                console.error(`Background: Error sending 'generateFullCaseView' message to tab ${targetTabId}:`, chrome.runtime.lastError.message);
            } else {
                console.log(`Background: Response from content script (tab ${targetTabId}) after sending 'generateFullCaseView':`, response);
            }
        });
    } else {
        console.error("Background: Cannot initiate generation - sender tab ID is missing.");
    }
    return true;
  }

  if (message.action === "openFullViewTab" && message.htmlContent) {
    const dataUrl = 'data:text/html;charset=UTF-8,' + encodeURIComponent(message.htmlContent);
    chrome.tabs.create({ url: dataUrl });
    return false;
  }

  if (message.action === "fetchItemDetails" && message.items && message.items.length > 0) {
    const itemType = message.items[0].type;
    const itemsToFetch = message.items;
    const senderTabId = sender.tab?.id;

    if (!senderTabId) {
         console.error(`Background: Cannot process 'fetchItemDetails' - sender tab ID missing.`);
         sendResponse({ status: "error", message: "Could not identify sender tab."});
         return false;
    }

    console.log(`Background: Handling 'fetchItemDetails' for ${itemsToFetch.length} ${itemType}(s) from tab ${senderTabId}.`);

    fetchAllDetailsViaTabs(itemsToFetch, itemType, senderTabId)
      .then(resultsMap => {
        console.log(`Background: Finished ${itemType} batch for tab ${senderTabId}. Sending success response.`);
        sendResponse({ status: "success", details: resultsMap });
      })
      .catch(error => {
        console.error(`Background: Error during 'fetchItemDetails' for tab ${senderTabId}:`, error);
        sendResponse({ status: "error", message: error.message });
      });

    return true;
  }

  console.log(`Background: Received unhandled message action="${message.action}"`);
  return false;
});

console.log("Background: Service worker listeners attached and ready.");
