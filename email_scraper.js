// email_scraper.js - Injected into temporary Email tabs to scrape details

console.log("Email Scraper: Script Injected.");

/**
 * Waits for an element matching the selector to appear in the DOM.
 * @param {string} selector - CSS selector
 * @param {Element} [baseElement=document] - Base element
 * @param {number} [timeout=10000] - Timeout in ms
 * @returns {Promise<Element|null>}
 */
function waitForElement(selector, baseElement = document, timeout = 10000) {
    // console.log(`Email Scraper: Waiting for "${selector}"...`);
    return new Promise((resolve) => {
        const startTime = Date.now();
        const interval = setInterval(() => {
            const element = baseElement.querySelector(selector);
            if (element) {
                // console.log(`Email Scraper: Found "${selector}"`);
                clearInterval(interval);
                resolve(element);
            } else if (Date.now() - startTime > timeout) {
                console.warn(`Email Scraper: Timeout waiting for "${selector}" in ${document.URL}`);
                clearInterval(interval);
                resolve(null); // Resolve with null on timeout
            }
        }, 250);
    });
}

// Main scraping logic
async function scrapeEmailDetails() {
    console.log("Email Scraper: Starting scrapeEmailDetails.");
    let subject = null;
    let from = null;
    let to = null;
    let dateText = null; // We get absolute date from background via table parse
    let bodyHTML = null;

    try {
        // Wait for a reliable element that indicates email view is loaded
        const emailArticle = await waitForElement('article.emailuiEmailMessage');
        if (!emailArticle) {
             throw new Error("Main email article element 'article.emailuiEmailMessage' not found.");
        }
        console.log("Email Scraper: Found email article.");

        // --- Extract Subject (From Header Highlights) ---
        const subjectElement = await waitForElement('header.forceHighlightsPanel h1.slds-page-header__title');
        if(subjectElement) {
            subject = subjectElement.textContent?.trim();
            console.log("Email Scraper: Found Subject:", subject);
        } else {
            console.warn("Email Scraper: Subject H1 not found in highlights.");
        }

        // --- Extract From ---
        const fromElement = emailArticle.querySelector('.fromDetail span.uiOutputText');
        if(fromElement) {
            from = fromElement.textContent?.trim();
            console.log("Email Scraper: Found From:", from);
        } else {
             console.warn("Email Scraper: From element not found.");
        }

        // --- Extract To ---
        const toElementList = emailArticle.querySelector('.toCcBccDetail ul.addressList');
        if(toElementList) {
             // Get combined text of all items, or iterate if more detail needed
            to = Array.from(toElementList.querySelectorAll('li'))
                      .map(li => li.textContent?.trim())
                      .filter(Boolean)
                      .join('; '); // Join multiple recipients with semicolon
            console.log("Email Scraper: Found To:", to);
        } else {
            console.warn("Email Scraper: To list element not found.");
        }

        // --- Extract Body from IFrame ---
        console.log("Email Scraper: Looking for iframe#emailuiFrame...");
        const iframe = await waitForElement('iframe#emailuiFrame');
        if (iframe) {
            console.log("Email Scraper: Found iframe. Attempting to access contentDocument...");
            // Add delay for iframe content to potentially load? Usually needed.
            await new Promise(resolve => setTimeout(resolve, 500));
            try {
                const iframeDoc = iframe.contentDocument || iframe.contentWindow?.document;
                if (iframeDoc && iframeDoc.body) {
                    console.log("Email Scraper: Accessed iframe body. Extracting innerHTML.");
                    bodyHTML = iframeDoc.body.innerHTML?.trim();
                    console.log("Email Scraper: Extracted bodyHTML length:", bodyHTML?.length);
                } else {
                    console.warn("Email Scraper: Could not access iframe contentDocument or body (possibly cross-origin or not loaded).");
                     // Check for srcdoc as fallback
                     if (iframe.srcdoc) {
                         console.log("Email Scraper: Found srcdoc attribute, using that.");
                         bodyHTML = iframe.srcdoc;
                     } else {
                         bodyHTML = "[Error accessing iframe content - no document or srcdoc]";
                     }
                }
            } catch (iframeError) {
                console.error("Email Scraper: Error accessing iframe content:", iframeError);
                bodyHTML = `[Error accessing iframe content: ${iframeError.message}]`;
            }
        } else {
            console.warn("Email Scraper: iframe#emailuiFrame not found.");
            bodyHTML = "[Email Body IFrame not found]";
        }

        // --- Extract Date (Relative - Not used, just for logging) ---
        const dateElement = emailArticle.querySelector('.dateContainer span.uiOutputDate');
        if(dateElement){
            dateText = dateElement.textContent?.trim(); // Get relative date text for logging
            console.log("Email Scraper: Found relative date text:", dateText);
        } else {
             console.warn("Email Scraper: Relative date element not found.");
        }


    } catch (error) {
        console.error("Email Scraper: Error during scraping:", error);
        bodyHTML = bodyHTML || `[Error during scraping: ${error.message}]`;
    }

    // Send results back to the background script
    // We don't send dateText because the absolute one parsed in background is needed
    console.log("Email Scraper: Sending results back.");
    chrome.runtime.sendMessage({
        type: 'emailScrapeResult',
        subject: subject ?? '',
        from: from ?? '',
        to: to ?? '',
        bodyHTML: bodyHTML ?? '' // Send extracted HTML body
    });
}

// Execute the scraping
scrapeEmailDetails();
