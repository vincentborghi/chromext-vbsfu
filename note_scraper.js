// note_scraper.js - Injected into temporary Note tabs to scrape details

console.log("Note Scraper: Script Injected.");

/**
 * Waits for an element matching the selector to appear in the DOM.
 * Uses a shorter timeout as the page should be mostly loaded.
 * @param {string} selector - CSS selector
 * @param {Element} [baseElement=document] - Base element
 * @param {number} [timeout=10000] - Timeout in ms
 * @returns {Promise<Element|null>}
 */
function waitForElement(selector, baseElement = document, timeout = 10000) {
    // console.log(`Note Scraper: Waiting for "${selector}"...`); // Optional: verbose logging
    return new Promise((resolve) => {
        const startTime = Date.now();
        const interval = setInterval(() => {
            const element = baseElement.querySelector(selector);
            if (element) {
                // console.log(`Note Scraper: Found "${selector}"`);
                clearInterval(interval);
                resolve(element);
            } else if (Date.now() - startTime > timeout) {
                console.warn(`Note Scraper: Timeout waiting for "${selector}" in ${document.URL}`);
                clearInterval(interval);
                resolve(null); // Resolve with null on timeout
            }
        }, 250);
    });
}

// Main scraping logic wrapped in an async function
async function scrapeNoteDetails() {
    console.log("Note Scraper: Starting scrapeNoteDetails.");
    let description = null;
    let createdDateText = null;

    try {
        // --- Extract Description ---
        // Wait for the layout item container first for stability
        const descLayoutItemSelector = 'records-record-layout-item[field-label="Description"]';
        const descLayoutItem = await waitForElement(descLayoutItemSelector);
        console.log(`Note Scraper: Description layout item found?`, !!descLayoutItem);
        if (descLayoutItem) {
            // Try the rich text span selector within the item
            const descTextElement = await waitForElement('records-formatted-rich-text span[part="formatted-rich-text"]', descLayoutItem, 5000); // Shorter wait inside
            console.log(`Note Scraper: Description rich text span found?`, !!descTextElement);
            if (descTextElement) {
                description = descTextElement.innerHTML?.trim(); // Get innerHTML for rich text
                console.log(`Note Scraper: Extracted description length:`, description?.length);
            } else {
                 console.warn("Note Scraper: Found Description container, but inner rich text span not found.");
                 description = descLayoutItem.textContent?.trim(); // Fallback to all text in container
                 description = `[Fallback TextContent] ${description}`;
            }
        } else {
            console.warn("Note Scraper: Description layout item not found.");
            description = "N/A (Description Container Not Found)";
        }

        // --- Extract Created Date ---
        // Wait for the 'Created By' item container
        const createdByItemSelector = 'records-record-layout-item[field-label="Created By"]';
        const createdByItem = await waitForElement(createdByItemSelector);
         console.log(`Note Scraper: 'Created By' layout item found?`, !!createdByItem);

        if (createdByItem) {
            // Wait for the specific date element within the modstamp
            const dateElementSelector = 'records-modstamp lightning-formatted-text';
            const dateElement = await waitForElement(dateElementSelector, createdByItem, 5000); // Shorter wait inside
            console.log(`Note Scraper: Date element found?`, !!dateElement);
            if (dateElement) {
                createdDateText = dateElement.textContent?.trim();
                console.log("Note Scraper: Extracted createdDateText:", createdDateText);
            } else {
                // Fallback if lightning-formatted-text isn't used directly
                const modstampElement = createdByItem.querySelector('records-modstamp');
                if (modstampElement) {
                     let modstampText = modstampElement.textContent?.trim();
                     console.log("Note Scraper: Fallback - records-modstamp textContent:", modstampText);
                     const commaIndex = modstampText.lastIndexOf(','); // Simple split attempt
                     if (commaIndex !== -1) createdDateText = modstampText.substring(commaIndex + 1).trim();
                     else createdDateText = modstampText; // Assume it's just the date
                     console.log("Note Scraper: Fallback extracted date text:", createdDateText);
                } else {
                    console.warn(`Note Scraper: Failed to find date element or records-modstamp within Created By item.`);
                }
            }
        } else {
             console.warn(`Note Scraper: Failed to find '${createdByItemSelector}'.`);
        }

    } catch (error) {
        console.error("Note Scraper: Error during scraping:", error);
        description = description || `Error during scraping: ${error.message}`; // Append error if desc wasn't found
    }

    // Send results back to the background script
    console.log("Note Scraper: Sending results back:", { description: description ?? '', createdDateText: createdDateText });
    chrome.runtime.sendMessage({
        type: 'noteScrapeResult',
        description: description ?? '', // Ensure string even if null
        createdDateText: createdDateText // Send raw text for parsing in background
    });
}

// Execute the scraping
scrapeNoteDetails();
