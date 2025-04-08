// background.js - v50 + Initiate Handler

console.log("Background service worker started (v50 + Initiate Handler).");

// --- Helper Function to Parse Date String ---
// Parses DD/MM/YYYY HH:MM format, with fallback for Date.parse
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
    // Note: This is less reliable for specific formats and can be timezone-dependent.
    const parsedFallback = Date.parse(dateString);
    if (!isNaN(parsedFallback)) {
        console.warn(`Background: Used Date.parse fallback for "${dateString}"`);
        return new Date(parsedFallback); // Return Date object from fallback
    }

    // If all parsing attempts fail
    console.warn(`Background: Could not parse date format "${dateString}"`);
    return null; // Return null if parsing failed
}

// --- Function to process all items (Notes or Emails) via temporary tabs ---
// Iterates through items, opens each URL in a hidden tab, injects a scraper script,
// waits for the result, and collects the data.
async function fetchAllDetailsViaTabs(itemsToFetch, itemType, senderTabId) {
    console.log(`Background: Starting tab automation for ${itemsToFetch.length} ${itemType}(s) from sender tab ${senderTabId}.`);
    const resultsMap = {}; // Object to store fetched details, keyed by item URL
    const scraperScript = itemType === 'Note' ? 'note_scraper.js' : 'email_scraper.js';
    let currentItemIndex = 0; // Counter for logging progress

    // Process items one by one sequentially to avoid overwhelming the browser or Salesforce
    for (const itemInfo of itemsToFetch) {
        currentItemIndex++;
        const itemUrl = itemInfo.url;

        // --- Send progress log message back to the original tab ---
        if (senderTabId) {
             console.log(`Background: Sending log message for ${itemUrl} to tab ${senderTabId}`);
             try {
                 // Send message without waiting for a response from the content script
                 await chrome.tabs.sendMessage(senderTabId, {
                     action: "logUrlProcessing",
                     url: itemUrl,
                     itemType: itemType,
                     index: currentItemIndex,
                     total: itemsToFetch.length
                 });
             } catch (err) {
                 // Log a warning if sending fails (e.g., original tab was closed)
                 console.warn(`Background: Failed to send log message to tab ${senderTabId} (may be closed): ${err.message}`);
                 // Optional: Decide whether to stop the whole process if the originating tab is gone
                 // Consider throwing an error here if continuing is pointless:
                 // throw new Error(`Originating tab ${senderTabId} closed or unresponsive.`);
             }
        } else {
            // Log if the sender tab ID wasn't available
            console.warn("Background: Cannot send log message, senderTabId is missing.")
        }

        console.log(`Background: Processing ${itemType} ${currentItemIndex}/${itemsToFetch.length}: ${itemUrl}`);
        let tempTab = null; // Variable to hold the temporary tab object
        let scrapeResult = null; // Variable to hold the result from the scraper

        try {
            // --- Create a new inactive tab for the item URL ---
            console.log(`Background: Creating inactive tab for ${itemUrl}`);
            tempTab = await chrome.tabs.create({ url: itemUrl, active: false });
            const tempTabId = tempTab.id;
            if (!tempTabId) {
                throw new Error("Failed to create temp tab (tab ID is null).");
            }
            console.log(`Background: Created tab ID: ${tempTabId}. Waiting for it to load...`);

            // --- Wait for the temporary tab to finish loading ---
            await new Promise((resolve, reject) => {
                const timeoutMillis = 30000; // 30 seconds timeout for tab loading
                // Set up a timeout timer
                const timeoutTimer = setTimeout(() => {
                    chrome.tabs.onUpdated.removeListener(tabUpdateListener); // Clean up the listener
                    console.error(`Background: Timeout waiting for tab ${tempTabId} to load URL: ${itemUrl}`);
                    reject(new Error(`Timeout (${timeoutMillis/1000}s) waiting for tab ${tempTabId} to load`));
                }, timeoutMillis);

                // Define the listener function for tab updates
                const tabUpdateListener = (tabId, changeInfo, tab) => {
                    // Check if the update is for our temporary tab
                    if (tabId === tempTabId) {
                         // Check if the tab is fully loaded
                         if (changeInfo.status === 'complete') {
                             clearTimeout(timeoutTimer); // Cancel the timeout
                             console.log(`Tab ${tempTabId} loaded successfully.`);
                             chrome.tabs.onUpdated.removeListener(tabUpdateListener); // Clean up the listener
                             resolve(); // Resolve the promise indicating success
                         }
                         // Check for potential errors or redirects to login pages
                         else if (changeInfo.status === 'error' || tab?.url?.includes('error') || tab?.url?.includes('login')) {
                             clearTimeout(timeoutTimer); // Cancel the timeout
                             console.warn(`Tab ${tempTabId} failed load or redirected. Status: ${changeInfo.status}, URL: ${tab?.url}`);
                             chrome.tabs.onUpdated.removeListener(tabUpdateListener); // Clean up the listener
                             reject(new Error(`Tab ${tempTabId} failed load or redirected.`)); // Reject the promise
                         }
                    }
                };
                // Attach the listener to tab update events
                chrome.tabs.onUpdated.addListener(tabUpdateListener);
            }); // End of tab load promise

             // --- Set up a listener specifically for the scrape result from this temp tab ---
             const resultPromise = new Promise((resolve, reject) => {
                 const timeoutMillis = 20000; // 20 seconds timeout for receiving scrape result
                 // Set up a timeout timer
                 const timeoutTimer = setTimeout(() => {
                     chrome.runtime.onMessage.removeListener(scrapeResultListener); // Clean up the listener
                     console.error(`Background: Timeout waiting for scrape result from tab ${tempTabId} for URL: ${itemUrl}`);
                     reject(new Error(`Timeout (${timeoutMillis/1000}s) waiting for scrape result from tab ${tempTabId}`));
                 }, timeoutMillis);

                 // Define the expected message type based on the item being scraped
                 const expectedResponseType = itemType === 'Note' ? 'noteScrapeResult' : 'emailScrapeResult';

                 // Define the listener function for incoming messages
                 const scrapeResultListener = (message, senderListener, sendResponseListener) => {
                     // Check if the message is from our temporary tab and has the correct type
                     if (senderListener.tab?.id === tempTabId && message.type === expectedResponseType) {
                         clearTimeout(timeoutTimer); // Cancel the timeout
                         console.log(`Received '${expectedResponseType}' result from tab ${tempTabId}.`);
                         chrome.runtime.onMessage.removeListener(scrapeResultListener); // Clean up the listener
                         resolve(message); // Resolve the promise with the received data
                         // Indicate synchronous handling for this specific message instance
                         // (prevents "message port closed" error for *this* listener)
                         return false;
                     }
                     // If the message is not the one we're waiting for, let other listeners handle it.
                     // Returning false indicates we didn't handle it asynchronously.
                     return false;
                 };
                 // Attach the listener for runtime messages
                 chrome.runtime.onMessage.addListener(scrapeResultListener);
             }); // End of result promise setup

            // --- Inject the appropriate scraping script into the loaded temporary tab ---
            console.log(`Background: Injecting ${scraperScript} into tab ${tempTabId}`);
            await chrome.scripting.executeScript({
                 target: { tabId: tempTabId }, // Target the specific temporary tab
                 files: [scraperScript]        // Inject the correct scraper file
            });
            console.log(`Background: Script ${scraperScript} injected. Waiting for result message...`);

            // --- Wait for the scraping result message ---
            scrapeResult = await resultPromise;

            // --- Process the received scrape result ---
            if (scrapeResult) {
                 console.log(`Background: Processing scrape result for ${itemUrl}.`);
                 // Parse the date string obtained from the *original* table data
                 const parsedDate = parseDateString(itemInfo.dateStr);

                 // Store the data based on item type
                 if (itemType === 'Note') {
                     resultsMap[itemUrl] = {
                         description: scrapeResult.description || '', // Scraped description (HTML)
                         dateObject: parsedDate, // Parsed Date object (or null)
                         isPublic: scrapeResult.isPublic // Scraped visibility status (boolean)
                     };
                 } else { // Email
                     resultsMap[itemUrl] = {
                         description: scrapeResult.bodyHTML || '', // Scraped email body (HTML)
                         dateObject: parsedDate, // Parsed Date object (or null)
                         subject: scrapeResult.subject || '', // Scraped subject
                         from: scrapeResult.from || '',       // Scraped 'from' address
                         to: scrapeResult.to || '',         // Scraped 'to' address(es)
                         isPublic: null // Visibility not applicable/fetched for emails
                     };
                 }
            } else {
                // This case should ideally not happen if the promise resolves correctly
                throw new Error("Scrape result promise resolved but result was null or undefined.");
            }
        } catch (error) {
            // Catch any errors that occurred during the process for this item
            console.error(`Background: Error processing ${itemType} ${itemUrl} in tab ${tempTab?.id}:`, error);
            // Store an error message in the results map for this item
            resultsMap[itemUrl] = {
                description: `Error processing item: ${error.message}`, // Store the error message
                dateObject: parseDateString(itemInfo.dateStr), // Still attempt to parse date from table data
                // Mark other fields appropriately, e.g., null or specific error indicators
                isPublic: null,
                subject: '[Error]',
                from: '[Error]',
                to: '[Error]'
            };
        } finally {
            // --- Ensure the temporary tab is closed, regardless of success or failure ---
            if (tempTab?.id) {
                try {
                    console.log(`Background: Attempting to close temp tab ${tempTab.id}`);
                    await chrome.tabs.remove(tempTab.id);
                    console.log(`Background: Temp tab ${tempTab.id} closed.`);
                } catch (closeError) {
                    // Log if closing the tab fails, but don't let it stop the overall process
                    console.warn(`Background: Error closing temp tab ${tempTab.id}:`, closeError.message);
                }
            }
        } // End try-catch-finally block for a single item
    } // End for loop iterating through all itemsToFetch

    console.log(`Background: Finished processing all ${itemsToFetch.length} ${itemType}(s) via tabs.`);
    return resultsMap; // Return the map containing fetched data or errors
}


// --- Main Listener for Messages from Content Scripts or Popups ---
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log(`Background: Received message action="${message.action}" from sender:`, sender.tab?.id, sender.url); // Log incoming messages

  // --- Handler to initiate the generation process (triggered by panel button in content script) ---
  if (message.action === "initiateGenerateFullCaseView") {
    console.log(`Background: Handling 'initiateGenerateFullCaseView' from tab ${sender.tab?.id}.`);
    // Check if the sender tab ID is available
    if (sender.tab?.id) {
        const targetTabId = sender.tab.id;
        console.log(`Background: Relaying 'generateFullCaseView' command to content script in tab ${targetTabId}`);
        // Send a message TO the specific content script, telling it to start generation
        chrome.tabs.sendMessage(targetTabId, { action: "generateFullCaseView" }, (response) => {
            // This callback runs after the content script's listener responds (or if sending fails)
            if (chrome.runtime.lastError) {
                // Log error if sending the message to the content script failed
                console.error(`Background: Error sending 'generateFullCaseView' message to tab ${targetTabId}:`, chrome.runtime.lastError.message);
                // Consider how to handle this - maybe log, maybe try again?
            } else {
                // Log the response received from the content script (if any)
                console.log(`Background: Response from content script (tab ${targetTabId}) after sending 'generateFullCaseView':`, response);
                // The content script should now be running its async HTML generation.
                // This background listener remains active (due to returning true)
                // to eventually receive the 'openFullViewTab' message.
            }
        });
    } else {
        // Log error if the sender tab ID was missing
        console.error("Background: Cannot initiate generation - sender tab ID is missing for 'initiateGenerateFullCaseView'.");
    }
    // ** IMPORTANT: Return true. **
    // This listener needs to stay active to handle the subsequent 'openFullViewTab' message,
    // which will be sent asynchronously by the content script after it finishes generating HTML.
    // Returning true keeps the message channel open for sendResponse.
    return true;
  }

  // --- Handler for opening the results tab (triggered by content script after HTML is ready) ---
  if (message.action === "openFullViewTab" && message.htmlContent) {
    console.log(`Background: Handling 'openFullViewTab' request from tab ${sender.tab?.id}.`);
    // Construct a 'data:' URL to display the generated HTML content directly
    const dataUrl = 'data:text/html;charset=UTF-8,' + encodeURIComponent(message.htmlContent);
    console.log(`Background: Created data URL (length approx: ${dataUrl.length})`);

    // Create a new tab to display the data URL
    chrome.tabs.create({ url: dataUrl }, (newTab) => {
        // This callback runs after Chrome attempts to create the tab
        if (chrome.runtime.lastError) {
            console.error(`Background: Error creating results tab: ${chrome.runtime.lastError.message}`);
            // Consider notifying the user if possible, though difficult from background script
        } else if (newTab) {
            console.log(`Background: Results tab created successfully (ID: ${newTab.id})`);
        } else {
            // This case should be rare if no error occurred, but check defensively
            console.error("Background: Results tab creation failed - newTab object is null/undefined.");
        }
    });
     // This listener function itself completes synchronously. Tab creation happens later via callback.
     // Therefore, we return false.
     return false;
  }

  // --- Handler for fetching Note OR Email details (triggered by content script) ---
  if (message.action === "fetchItemDetails" && message.items && message.items.length > 0) {
    const itemType = message.items[0].type; // Assume all items in the batch are the same type
    const itemsToFetch = message.items;
    const senderTabId = sender.tab?.id; // Get the ID of the tab that requested the details

    // Validate senderTabId
    if (!senderTabId) {
         console.error(`Background: Cannot process 'fetchItemDetails' - sender tab ID missing. Sender URL: ${sender.url}`);
         // Send an error response back to the content script
         sendResponse({ status: "error", message: "Could not identify sender tab."});
         // Return false as we handled this synchronously (by sending an error)
         return false;
    }

    console.log(`Background: Handling 'fetchItemDetails' for ${itemsToFetch.length} ${itemType}(s) from tab ${senderTabId}.`);

    // Call the asynchronous function that performs the tab automation and scraping
    fetchAllDetailsViaTabs(itemsToFetch, itemType, senderTabId)
      .then(resultsMap => {
        // Success: Send the map of fetched details back to the content script
        console.log(`Background: Finished ${itemType} batch for tab ${senderTabId}. Sending success response.`);
        sendResponse({ status: "success", details: resultsMap });
      })
      .catch(error => {
        // Error: Send an error message back to the content script
        console.error(`Background: Error during 'fetchItemDetails' for tab ${senderTabId}:`, error);
        sendResponse({ status: "error", message: error.message || `Unknown error during ${itemType} tab automation.` });
      });

    // ** IMPORTANT: Return true. **
    // Because fetchAllDetailsViaTabs is async and we use .then()/.catch() to eventually call
    // sendResponse, we must return true here to keep the message channel open for the response.
    return true;
  }

  // --- Default handling for any messages not matched above ---
  console.log(`Background: Received unhandled message action="${message.action}" from tab ${sender.tab?.id}, URL: ${sender.url}`);
  // Return false as we are handling this synchronously (by doing nothing).
  return false;

}); // End of chrome.runtime.onMessage.addListener

console.log("Background: Service worker listeners attached and ready.");
