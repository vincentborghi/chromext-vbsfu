// content.js - Complete File (v50 + ... + Sliding Panel + Generate Fix v2 + Syntax Fix + View All Fix)

console.log("Salesforce Case Full View Content Script Loaded (v50 + ... + View All Fix).");

// --- Helper Functions ---
/**
 * Waits for an element matching the selector to appear in the DOM.
 * Returns a promise that resolves with the element or null if timed out.
 */
function waitForElement(selector, baseElement = document, timeout = 15000) {
  // console.log(`Waiting for "${selector}"...`); // Optional logging
  return new Promise((resolve) => {
    const startTime = Date.now();
    const interval = setInterval(() => {
      const element = baseElement.querySelector(selector);
      if (element) {
        // console.log(`waitForElement: Found "${selector}"`);
        clearInterval(interval);
        resolve(element); // Element found
      } else if (Date.now() - startTime > timeout) {
        // Timeout occurred
        console.warn(`waitForElement: Timeout waiting for "${selector}"`);
        clearInterval(interval);
        resolve(null); // Resolve with null on timeout
      }
    }, 300); // Check every 300ms
  });
}

/**
 * Wraps chrome.runtime.sendMessage in a Promise for async/await usage.
 * Useful when the content script needs a response from the background script.
 */
function sendMessagePromise(message) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(message, (response) => {
      // Check for errors during message sending/receiving
      if (chrome.runtime.lastError) {
        console.error("sendMessagePromise failed:", chrome.runtime.lastError.message);
        reject(chrome.runtime.lastError);
      } else {
        // Message sent and response received successfully
        resolve(response);
      }
    });
  });
}

/**
 * Escapes HTML special characters (&, <, >, ", ') in a string to prevent XSS.
 * Handles non-string inputs safely.
 * @param {string | null | undefined} unsafe - The string to escape.
 * @returns {string} - The escaped string, or an empty string for null/undefined input.
 */
function escapeHtml(unsafe) {
  // Handle null, undefined, or non-string types
  if (typeof unsafe !== 'string') {
      // Convert null/undefined to '', other types to string representation
      return unsafe === null || typeof unsafe === 'undefined' ? '' : String(unsafe);
  }
  // Perform HTML escaping
  return unsafe
    .replace(/&/g, "&amp;") // Must be first
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;"); // Used for single quotes
}

/**
 * Safely gets text content from a child element specified by a selector within a parent.
 * @param {Element | null} element - The parent element to search within.
 * @param {string} selector - The CSS selector for the desired child element.
 * @returns {string} - The trimmed text content of the child, or 'N/A' if not found.
 */
function getTextContentFromElement(element, selector) {
  if (!element) return 'N/A'; // Parent element doesn't exist
  const childElement = element.querySelector(selector);
  // Return trimmed text or 'N/A' if child or text content is missing
  return childElement ? childElement.textContent?.trim() : 'N/A';
}

/**
 * Parses a date string (DD/MM/YYYY HH:MM format expected) from the table into a Date object (UTC).
 * Includes basic validation and fallback parsing.
 * @param {string | null | undefined} dateString - The date string to parse.
 * @returns {Date | null} - The parsed Date object (in UTC) or null if parsing fails.
 */
function parseDateStringFromTable(dateString) {
  if (!dateString) return null;
  // Regex for DD/MM/YYYY HH:MM
  let match = dateString.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);
  if (match) {
    try {
      const day = parseInt(match[1]);
      const month = parseInt(match[2]) - 1; // Month is 0-indexed
      const year = parseInt(match[3]);
      const hour = parseInt(match[4]);
      const minute = parseInt(match[5]);

      // Basic validation of parsed date components
      if (year > 1970 && month >= 0 && month < 12 && day >= 1 && day <= 31 && hour >= 0 && hour < 24 && minute >= 0 && minute < 60) {
        // Create Date object in UTC
        const dateObject = new Date(Date.UTC(year, month, day, hour, minute));
        // Check if the resulting Date object is valid
        if (!isNaN(dateObject.getTime())) {
            return dateObject;
        } else {
             console.warn(`CS Date Invalid Comp: "${dateString}" resulted in invalid Date object`);
             return null; // Invalid date components combined
        }
      } else {
        console.warn(`CS Date Comp Range Invalid: "${dateString}"`);
        return null; // Components out of valid range
      }
    } catch (e) {
      console.error(`CS Error creating Date from matched parts: "${dateString}"`, e);
      return null; // Error during parsing
    }
  } else {
    // Fallback attempt using Date.parse if regex doesn't match
    const parsedFallback = Date.parse(dateString);
    if (!isNaN(parsedFallback)) {
        // Note: Date.parse interpretation can be inconsistent regarding timezones.
        console.warn(`CS Date parsed with fallback Date.parse: "${dateString}"`);
        return new Date(parsedFallback);
    } else {
        // If both regex and fallback fail
        console.warn(`CS Could not parse date format with regex or Date.parse: "${dateString}"`);
        return null;
    }
  }
}


// --- Functions to Extract Salesforce Record Details ---

// Finds Subject in the highlights panel
function findSubjectInContainer(container) {
    if (!container) return 'N/A';
    const element = container.querySelector('support-output-case-subject-field lightning-formatted-text');
    return element ? element.textContent?.trim() : 'N/A (Subject)';
}

// Older, potentially fragile way to find Case Number
function findCaseNumberInContainer(container) {
     if (!container) return 'N/A';
     const firstItem = container.querySelector('records-highlights-details-item:nth-of-type(1)');
     const element = firstItem?.querySelector('p.fieldComponent lightning-formatted-text');
     if (element &&
         !firstItem.querySelector('support-output-case-subject-field') &&
         !firstItem.querySelector('records-formula-output') &&
         !firstItem.querySelector('force-lookup') )
     {
         return element.textContent?.trim() || 'N/A (Case #)';
     }
     console.warn("Case Number selector (1st item, specific checks) failed.");
     return 'N/A (Case #)';
}

// More reliable way to find Case Number using title attribute
function findCaseNumberSpecific() {
    console.log("Content Script: Attempting to find Case/Record Number (Specific Selector)...");
    // Look for a highlights item that contains a paragraph with title="Case Number" (or adapt title if needed)
    // Adapt this title based on the actual Salesforce field label if it's not "Case Number"
    const itemSelector = 'records-highlights-details-item:has(p[title="Case Number"])';
    const textSelector = 'lightning-formatted-text'; // The element usually containing the number
    const detailsItem = document.querySelector(itemSelector);

    if (detailsItem) {
        const textElement = detailsItem.querySelector(textSelector);
        if (textElement) {
            const recordNum = textElement.textContent?.trim();
            // Basic validation: Check if it looks like a number (adjust regex if needed)
            if (recordNum && /^\d+$/.test(recordNum)) {
                 console.log("Content Script: Found Record Number:", recordNum);
                 return recordNum;
            } else {
                 console.warn("Content Script: Found element, but content doesn't look like a record number:", recordNum);
            }
        } else {
             console.warn("Content Script: Found details item, but inner text element not found with selector:", textSelector);
        }
    } else {
         console.warn("Content Script: Record Number details item not found with selector:", itemSelector);
         // Fallback attempt using the older function IF the new one fails
         console.log("Content Script: Falling back to older record number detection...");
         const container = document.querySelector('records-highlights2'); // Assume highlights container
         const fallbackRecordNum = findCaseNumberInContainer(container);
         // Check if fallback result is valid before returning
         if (fallbackRecordNum && !fallbackRecordNum.startsWith('N/A')) {
             return fallbackRecordNum;
         }
    }
    return null; // Return null if not found or invalid after all attempts
}

// Finds Status in the highlights panel
function findStatusInContainer(container) {
    if (!container) return 'N/A';
    const statusItem = container.querySelector('records-highlights-details-item:has(records-formula-output lightning-formatted-rich-text)');
    const element = statusItem?.querySelector('lightning-formatted-rich-text span[part="formatted-rich-text"]');
    return element ? element.textContent?.trim() : 'N/A (Status)';
}

// Finds Creator Name (async)
async function findCreatorName() {
    const createdByItem = await waitForElement('records-record-layout-item[field-label="Created By"]');
    if (!createdByItem) { console.warn("Creator layout item not found."); return 'N/A (Creator)'; }
    const nameElement = createdByItem.querySelector('force-lookup a');
    return nameElement ? nameElement.textContent?.trim() : 'N/A (Creator)';
}

// Finds Created Date string (async)
async function findCreatedDate() {
    const createdByItem = await waitForElement('records-record-layout-item[field-label="Created By"]');
    if (!createdByItem) { console.warn("Created Date layout item not found."); return 'N/A (Created Date)'; }
    const dateElement = createdByItem.querySelector('records-modstamp lightning-formatted-text');
    return dateElement ? dateElement.textContent?.trim() : 'N/A (Created Date)';
}

// Finds Owner Name in highlights panel
function findOwnerInContainer(container) {
     if (!container) return 'N/A';
     const ownerItem = container.querySelector('records-highlights-details-item:has(force-lookup)');
     const element = ownerItem?.querySelector('force-lookup a');
     return element ? element.textContent?.trim() : 'N/A (Owner)';
}

// Finds Account Name (async)
async function findAccountName() {
    console.log("Attempting to find Account Name...");
    const accountItem = await waitForElement('records-record-layout-item[field-label="Account Name"]');
    if (!accountItem) {
        console.warn("Could not find layout item with field-label='Account Name'.");
        return 'N/A (Account)';
    }
    const nameElement = accountItem.querySelector('force-lookup a');
    console.log("Account name element found:", !!nameElement);
    return nameElement ? nameElement.textContent?.trim() : 'N/A (Account)';
}

// Finds Description, handling 'View More' (async)
async function findCaseDescription() {
     // Adjust the selector based on the actual component used for Description
     const descriptionContainer = await waitForElement('article.cPSM_Case_Description');
     if (!descriptionContainer) {
         console.warn("Case Description container 'article.cPSM_Case_Description' not found.");
         return '';
     }
     // Find the text element within the container
     let textElement = descriptionContainer.querySelector('lightning-formatted-text.txtAreaReadOnly') || descriptionContainer.querySelector('lightning-formatted-text');
     if (!textElement) {
         console.warn("Case Description text element not found inside container.");
         return '';
     }
     // Look for a 'View More' or 'Show More' button within the description container
     const viewMoreButton = descriptionContainer.querySelector('button.slds-button:not([disabled])');
     let descriptionHTML = '';

     // If a visible 'View More' button exists
     if (viewMoreButton && (viewMoreButton.textContent.includes('View More') || viewMoreButton.textContent.includes('Show More'))) {
        console.log("Clicking 'View More' for description...");
        viewMoreButton.click();
        // Wait a short time for the content to expand/load
        await new Promise(resolve => setTimeout(resolve, 500));
        // Re-query for the text element, as it might have been replaced by the 'View More' action
        let updatedTextElement = descriptionContainer.querySelector('lightning-formatted-text.txtAreaReadOnly') || descriptionContainer.querySelector('lightning-formatted-text');
        descriptionHTML = updatedTextElement?.innerHTML?.trim() || ''; // Get HTML content
        console.log("Description length after 'View More':", descriptionHTML.length);
     } else {
        // If no 'View More' button, get the initial HTML content
        descriptionHTML = textElement?.innerHTML?.trim() || '';
        console.log("Description length (no 'View More'):", descriptionHTML.length);
     }
     return descriptionHTML;
}


// --- Function to Extract Note URLs from Table and Trigger Fetch ---
// Finds notes related list, clicks 'View All' if needed, extracts basic info,
// sends info to background script for detailed fetching.
async function extractAndFetchNotes() {
  const notesHeaderSelector = 'a.slds-card__header-link[href*="/related/PSM_Notes__r/view"]'; // Selector for Notes related list header link
  console.log("Extracting Notes: Looking for header link...");
  const headerLinkElement = await waitForElement(notesHeaderSelector);
  if (!headerLinkElement) { console.warn(`Notes header link not found ('${notesHeaderSelector}'). Skipping notes.`); return []; }
  console.log("Extracting Notes: Found header link.");

  const headerElement = headerLinkElement.closest('lst-list-view-manager-header');
  if (!headerElement) { console.warn("Notes parent header element ('lst-list-view-manager-header') not found. Skipping notes."); return []; }

  // Check count in header link span to potentially skip early if count is (0)
  const countSpan = headerLinkElement.querySelector('span[title*="("]');
  if (countSpan && countSpan.textContent?.trim() === '(0)') { console.log("Notes count in header is (0), skipping fetch."); return []; }

  // --- Find the container for the related list view ---
  // This might be lst-common-list-internal or lst-related-list-view-manager
  let initialContainer = headerElement.closest('lst-common-list-internal') || headerElement.closest('lst-related-list-view-manager');
  if (!initialContainer) { console.error("Cannot find initial Notes container (lst-common-list-internal or lst-related-list-view-manager). Skipping notes."); return []; }
  console.log("Extracting Notes: Found initial container.");

  // --- Find and click the 'View All' link if it's present and visible ---
  const listManager = initialContainer.closest('lst-related-list-view-manager') || initialContainer;
  const viewAllLinkSelector = 'a.slds-card__footer[href*="/related/PSM_Notes__r/view"]';
  const viewAllLink = listManager?.querySelector(viewAllLinkSelector);

  let currentContainer = initialContainer; // Start with the initial container

  // Check offsetParent !== null to ensure the link is actually visible on the page
  if (viewAllLink && viewAllLink.offsetParent !== null) {
    console.log("Extracting Notes: 'View All' link found and visible. Clicking...");
    viewAllLink.click();
    // ** INCREASED WAIT TIME ** after clicking 'View All'
    console.log("Extracting Notes: Waiting 3 seconds for table to update after 'View All' click...");
    await new Promise(resolve => setTimeout(resolve, 3000)); // Increased wait to 3 seconds
    console.log("Extracting Notes: Wait finished. Re-finding container and table elements...");

    // ** RE-FIND CONTAINER ** after click/wait, as the DOM might have significantly changed
    // We assume the header link is stable and find the container relative to it again.
    const updatedHeaderElement = await waitForElement(notesHeaderSelector); // Re-find header
    if (updatedHeaderElement) {
        const updatedContainerParent = updatedHeaderElement.closest('lst-list-view-manager-header');
        if (updatedContainerParent) {
            currentContainer = updatedContainerParent.closest('lst-common-list-internal') || updatedContainerParent.closest('lst-related-list-view-manager');
            if (!currentContainer) {
                 console.error("Extracting Notes: Failed to re-find container after 'View All'. Using initial container.");
                 currentContainer = initialContainer; // Fallback to initial if re-find fails
            } else {
                 console.log("Extracting Notes: Successfully re-found container after 'View All'.");
            }
        } else {
             console.warn("Extracting Notes: Failed to find parent header after 'View All'. Using initial container.");
             currentContainer = initialContainer;
        }
    } else {
        console.warn("Extracting Notes: Failed to re-find header link after 'View All'. Using initial container.");
        currentContainer = initialContainer;
    }

  } else {
      console.log("Extracting Notes: 'View All' link not found or not visible.");
  }

  // --- Wait for the datatable element within the *current* container ---
  console.log("Extracting Notes: Waiting for datatable within the container...");
  const dataTable = await waitForElement('lightning-datatable', currentContainer, 10000); // 10 sec timeout
  if (!dataTable) { console.warn("Notes datatable ('lightning-datatable') not found inside container. Skipping notes."); return []; }
  console.log("Extracting Notes: Found datatable.");

  // --- Wait for the table body to be present within the datatable ---
  console.log("Extracting Notes: Waiting for table body...");
  const tableBody = await waitForElement('tbody[data-rowgroup-body]', dataTable, 5000); // 5 sec timeout
  if (!tableBody) { console.warn("Notes table body ('tbody[data-rowgroup-body]') not found. Skipping notes."); return []; }
  console.log("Extracting Notes: Found table body.");

  // --- Get all data rows within the table body ---
  const rows = tableBody.querySelectorAll('tr[data-row-key-value]');
  // ** ADD LOGGING ** for number of rows found
  console.log(`Extracting Notes: Found ${rows.length} rows in the table body.`);

  const notesToFetch = []; // Array to hold info for notes that need detailed fetching
  const origin = window.location.origin; // Get base URL for resolving relative links

  // Iterate through each row found in the table
  rows.forEach((row, index) => {
    // Extract data from specific cells using data-label attributes for robustness
    const noteLink = row.querySelector('th[data-label="PSM Note Name"] a'); // Note Title link
    const authorLink = row.querySelector('td[data-label="Created By"] a'); // Author link
    const dateSpan = row.querySelector('td[data-label="Created Date"] lst-formatted-text span'); // Date span
    const snippetEl = row.querySelector('td[data-label="Description"] lightning-base-formatted-text'); // Description snippet

    const relativeUrl = noteLink?.getAttribute('href'); // Get the relative URL
    // Get date string - prefer title attribute, fallback to textContent
    const noteDateStr = dateSpan?.title || dateSpan?.textContent?.trim();
    const noteTitle = noteLink?.textContent?.trim() || 'N/A';
    const noteAuthor = authorLink?.textContent?.trim() || 'N/A';
    const noteSnippet = snippetEl?.textContent?.trim() || ''; // Get snippet text

    // Check if essential data (URL and Date) was found before adding to fetch list
    if (relativeUrl && noteDateStr) {
      notesToFetch.push({
        type: 'Note',
        url: new URL(relativeUrl, origin).href, // Construct absolute URL
        title: noteTitle,
        author: noteAuthor,
        dateStr: noteDateStr, // Store original date string from table
        descriptionSnippet: noteSnippet // Store snippet as potential fallback
      });
    } else {
      // Log if essential data is missing for a row
      console.warn(`Skipping Note row ${index+1}: Missing URL (${!!relativeUrl}), Date (${!!noteDateStr}), Title (${!!noteTitle}), or Author (${!!noteAuthor}).`);
    }
  });

  // If no valid notes found in table after processing rows, return empty array
  if (notesToFetch.length === 0) {
    console.log("Extracting Notes: No valid Note rows found after processing table.");
    return [];
  }

  // Send the list of notes to the background script for detailed fetching
  console.log(`Extracting Notes: Sending ${notesToFetch.length} notes to background for detail fetching...`);
  try {
    // Use sendMessagePromise to handle the asynchronous response from the background script
    const response = await sendMessagePromise({ action: "fetchItemDetails", items: notesToFetch });

    // Process the response received from the background script
    if (response?.status === "success" && response.details) {
      console.log(`Extracting Notes: Received details for ${Object.keys(response.details).length} notes from background.`);
      // Map the original note info (from table) with the fetched details (from background)
      return notesToFetch.map(noteInfo => {
        const fetched = response.details[noteInfo.url]; // Get details for this specific note URL
        // Use the parsed date object from background if available, otherwise parse the table string as fallback
        const finalDateObject = fetched?.dateObject ? new Date(fetched.dateObject) : parseDateStringFromTable(noteInfo.dateStr);
        let finalDescription = fetched?.description; // Fetched description (HTML or error message)

        // Handle cases where fetching the description failed or returned empty
        if (finalDescription?.startsWith('Error:')) {
          // If fetch error occurred, use snippet from table as fallback if available
          finalDescription = noteInfo.descriptionSnippet ? `[Fetch Error, Snippet: ${escapeHtml(noteInfo.descriptionSnippet)}]` : escapeHtml(finalDescription);
        } else if (!finalDescription && noteInfo.descriptionSnippet) {
            // If fetch succeeded but returned empty, still consider snippet
            finalDescription = `[Content Empty, Snippet: ${escapeHtml(noteInfo.descriptionSnippet)}]`;
        } else if (!finalDescription) {
            // If no description and no snippet
            finalDescription = '[Content Empty or Not Fetched]';
        }
        // If fetch succeeded and returned content, finalDescription is the raw HTML

        // Return the combined object containing all info for this note
        return {
          type: 'Note',
          url: noteInfo.url,
          title: noteInfo.title,
          author: noteInfo.author,
          dateStr: noteInfo.dateStr, // Keep original string for display if needed
          dateObject: finalDateObject, // Parsed date object for sorting
          content: finalDescription, // Fetched content (HTML or error/fallback message)
          isPublic: fetched?.isPublic ?? null, // Visibility status (boolean or null)
          attachments: 'N/A' // Placeholder for attachments (not currently implemented)
        };
      }).filter(note => note.dateObject !== null); // Filter out notes where date parsing failed completely
    } else {
      // Handle error response from background script (e.g., background reported failure)
      throw new Error(response?.message || "Invalid response or missing details for Notes from background");
    }
  } catch (error) {
    // Handle errors during the message sending/receiving process itself
    console.error("Error sending/receiving Note details to/from background:", error);
    return []; // Return empty array on error
  }
} // End of extractAndFetchNotes


// --- Function to Extract Email Info from Table and Trigger Fetch ---
// Finds emails related list, clicks 'View All' if needed, extracts basic info,
// sends info to background script for detailed fetching.
async function extractAndFetchEmails() {
  const emailsListContainerSelector = 'div.forceListViewManager[aria-label*="Emails"]'; // Selector for Emails related list container
  console.log("Extracting Emails: Looking for container...");
  const emailsListContainer = await waitForElement(emailsListContainerSelector);
  if (!emailsListContainer) { console.warn(`Emails container not found ('${emailsListContainerSelector}'). Skipping emails.`); return []; }
   console.log("Extracting Emails: Found container.");

  // Check count in header link span if available to potentially skip early
  const headerLinkElement = emailsListContainer.querySelector('lst-list-view-manager-header a.slds-card__header-link[title="Emails"]');
  if (headerLinkElement) {
      const countSpan = headerLinkElement.querySelector('span[title*="("]');
      if (countSpan && countSpan.textContent?.trim() === '(0)') {
          console.log("Emails count in header is (0), skipping fetch.");
          return [];
      }
  }

  // Find and click the 'View All' link if it's present and visible
  const parentCard = emailsListContainer.closest('.forceRelatedListSingleContainer'); // Find parent card
  const viewAllLinkSelector = 'a.slds-card__footer[href*="/related/EmailMessages/view"]';
  const viewAllLink = parentCard?.querySelector(viewAllLinkSelector);
  if (viewAllLink && viewAllLink.offsetParent !== null) { // Check visibility
    console.log("Extracting Emails: 'View All' link found and visible. Clicking...");
    viewAllLink.click();
    // ** INCREASED WAIT TIME ** after clicking 'View All'
    console.log("Extracting Emails: Waiting 3 seconds for table to update after 'View All' click...");
    await new Promise(resolve => setTimeout(resolve, 3000)); // Increased wait to 3 seconds
    console.log("Extracting Emails: Wait finished. Re-finding table elements...");
    // Note: Unlike notes, emails often use uiVirtualDataTable which might reload differently.
    // We'll rely on waitForElement finding the potentially new table within the container.
  } else {
      console.log("Extracting Emails: 'View All' link not found or not visible.");
  }

  // Wait for the email table (often uses 'uiVirtualDataTable' class) within the container
  console.log("Extracting Emails: Waiting for table...");
  const dataTable = await waitForElement('table.uiVirtualDataTable', emailsListContainer, 10000); // 10 sec timeout
  if (!dataTable) { console.warn("Emails table ('uiVirtualDataTable') not found. Skipping emails."); return []; }
   console.log("Extracting Emails: Found table.");

  // Wait for the table body
  console.log("Extracting Emails: Waiting for table body...");
  const tableBody = await waitForElement('tbody', dataTable, 5000); // 5 sec timeout
  if (!tableBody) { console.warn("Emails table body ('tbody') not found. Skipping emails."); return []; }
   console.log("Extracting Emails: Found table body.");

  // Get all table rows
  const rows = tableBody.querySelectorAll('tr');
  // ** ADD LOGGING ** for number of rows found
  console.log(`Extracting Emails: Found ${rows.length} rows in the table body.`);

  const emailsToFetch = []; // Array to hold info for emails to fetch
  const origin = window.location.origin; // Base URL for resolving relative links

  // Iterate through each row
  rows.forEach((row, index) => {
    const cells = row.querySelectorAll('th, td');
    // Basic check for expected number of cells to avoid errors on unusual rows
    if (cells.length < 5) {
      console.warn(`Skipping email row ${index+1} due to insufficient cells (${cells.length}).`);
      return;
    }
    // Selectors based on common cell indices/classes (adjust if Salesforce UI changes)
    const subjectLink = cells[1]?.querySelector('a.outputLookupLink'); // Cell 1: Subject link
    const fromEl = cells[2]?.querySelector('a.emailuiFormattedEmail'); // Cell 2: From link (often email address)
    const toEl = cells[3]?.querySelector('span.uiOutputText');         // Cell 3: To text (can be truncated)
    const dateEl = cells[4]?.querySelector('span.uiOutputDateTime');   // Cell 4: Date text

    const relativeUrl = subjectLink?.getAttribute('href');
    const emailSubject = subjectLink?.textContent?.trim();
    const emailDateStr = dateEl?.textContent?.trim();
    const emailFrom = fromEl?.textContent?.trim() || 'N/A';
    // Use the 'title' attribute as a fallback for 'To' field, as text content might be truncated
    const emailTo = toEl?.title || toEl?.textContent?.trim() || 'N/A';

    // Check if essential data (URL, Subject, Date) was found
    if (relativeUrl && emailSubject && emailDateStr) {
      emailsToFetch.push({
        type: 'Email',
        url: new URL(relativeUrl, origin).href, // Create absolute URL
        title: emailSubject, // Use Subject as title for consistency
        author: emailFrom, // Use 'author' for 'From' field for consistency
        to: emailTo,
        dateStr: emailDateStr // Store original date string
      });
    } else {
      // Log if essential data is missing
      console.warn(`Skipping Email row ${index+1}: Missing URL (${!!relativeUrl}), Subject (${!!emailSubject}), or Date (${!!emailDateStr}).`);
    }
  });

   // If no valid emails found after processing rows, return empty array
   if (emailsToFetch.length === 0) {
    console.log("Extracting Emails: No valid Email rows found after processing table.");
    return [];
  }

  // Send the list of emails to the background script for detailed fetching
  console.log(`Extracting Emails: Sending ${emailsToFetch.length} emails to background for detail fetching...`);
  try {
    const response = await sendMessagePromise({ action: "fetchItemDetails", items: emailsToFetch });

    // Process the response from the background script
    if (response?.status === "success" && response.details) {
      console.log(`Extracting Emails: Received details for ${Object.keys(response.details).length} emails from background.`);
      // Map the original email info with the fetched details
      return emailsToFetch.map(emailInfo => {
        const fetched = response.details[emailInfo.url]; // Get details for this URL
        // Use parsed date from background or parse table string as fallback
        const finalDateObject = fetched?.dateObject ? new Date(fetched.dateObject) : parseDateStringFromTable(emailInfo.dateStr);
        let finalContent = fetched?.description; // Fetched HTML body or error message

         // Handle fetch error or empty body cases
        if (finalContent?.startsWith('Error:')) {
            // If body fetch failed, display subject and the error message
            finalContent = `Subject: ${escapeHtml(fetched?.subject || emailInfo.title)}<br>[Body Fetch Error: ${escapeHtml(finalContent)}]`;
        } else if (!finalContent) {
             // If body wasn't fetched or is empty, display subject and a placeholder
             finalContent = `Subject: ${escapeHtml(fetched?.subject || emailInfo.title)}<br>[Body Not Fetched or Empty]`;
        }
        // If fetch succeeded, finalContent contains the raw HTML body from the email scraper

        // Return the combined object for this email
        return {
          type: 'Email',
          url: emailInfo.url,
          title: fetched?.subject || emailInfo.title, // Prefer fetched subject if available
          author: fetched?.from || emailInfo.author, // Prefer fetched 'from' if available
          to: fetched?.to || emailInfo.to, // Prefer fetched 'to' if available
          dateStr: emailInfo.dateStr, // Keep original date string
          dateObject: finalDateObject, // Parsed date object for sorting
          content: finalContent, // HTML body or error/placeholder message
          isPublic: fetched?.isPublic ?? null, // Always null for emails currently
          attachments: 'N/A' // Placeholder for attachments
        };
      }).filter(email => email.dateObject !== null); // Filter out items where date parsing failed
    } else {
      // Handle error response from background script
      throw new Error(response?.message || "Invalid response or missing details for Emails from background");
    }
  } catch (error) {
    // Handle errors during message sending/receiving
    console.error("Error sending/receiving Email details to/from background:", error);
    return []; // Return empty array on error
  }
} // End of extractAndFetchEmails


// --- Main Function to Orchestrate Extraction and Generate HTML ---
// Gathers all data by calling helper functions, fetches related item details via background script,
// combines and sorts timeline items, and builds the final HTML report string.
async function generateCaseViewHtml(generatedTime) {
    console.log("Starting async HTML generation...");
    // Determine the type of Salesforce record (Case, WorkOrder, etc.) based on the URL
    let objectType = 'Record'; // Default type
    const currentHref = window.location.href;
    if (currentHref.includes('/Case/')) {
        objectType = 'Case';
    } else if (currentHref.includes('/WorkOrder/')) {
        objectType = 'WorkOrder';
    } // Add more 'else if' blocks here to support other object types if needed
    console.log(`Detected Object Type: ${objectType}`);

    // --- Find the main highlights container element ---
    const HIGHLIGHTS_CONTAINER_SELECTOR = 'records-highlights2'; // Selector for the highlights panel
    const highlightsContainerElement = await waitForElement(HIGHLIGHTS_CONTAINER_SELECTOR);
    // If the main container isn't found, we can't proceed. Return error HTML.
    if (!highlightsContainerElement) {
      console.error(`FATAL: Highlights container ("${HIGHLIGHTS_CONTAINER_SELECTOR}") not found. Cannot extract details.`);
      // Generate a minimal HTML page indicating the error
      return `<html><head><title>Error</title></head><body><h1>Extraction Error</h1><p>Could not find the main details section (<code>${escapeHtml(HIGHLIGHTS_CONTAINER_SELECTOR)}</code>) on the page. Cannot generate view.</p></body></html>`;
    } else {
      console.log("Main highlights container found.");
    }

    // --- Start extracting header details and related list data concurrently ---
    console.log("Extracting header details and related lists concurrently...");
    // Create an array of promises for all data extraction tasks
    const extractionPromises = [
        // Wrap synchronous calls in Promise.resolve for consistency
        Promise.resolve(findSubjectInContainer(highlightsContainerElement)),
        Promise.resolve(findCaseNumberSpecific()), // Use the more reliable function
        Promise.resolve(findStatusInContainer(highlightsContainerElement)),
        Promise.resolve(findOwnerInContainer(highlightsContainerElement)),
        findCreatorName(), // Async function
        findAccountName(), // Async function
        findCreatedDate(), // Async function
        findCaseDescription(), // Async function (handles 'View More')
        extractAndFetchNotes(), // Async function (includes background communication)
        extractAndFetchEmails() // Async function (includes background communication)
    ];

    // --- Wait for all extraction promises to resolve ---
    console.log("Waiting for all data extraction promises to resolve...");
    // Use Promise.all to wait for all tasks, destructure results
    const [
        subject, recordNumber, status, owner, creatorName, accountName, createdDateStr, description,
        notesData, emailsData
    ] = await Promise.all(extractionPromises);

    // Log a summary of the extracted data counts/presence
    console.log(`Extraction results: Subject=${!!subject}, ${objectType}#=${recordNumber}, Status=${!!status}, Owner=${!!owner}, Creator=${!!creatorName}, AccountName=${!!accountName}, CreatedDate=${!!createdDateStr}, Desc=${description?.length>0}, Notes=${notesData?.length || 0}, Emails=${emailsData?.length || 0}`);

    // --- Combine and Sort Notes & Emails into a single timeline ---
    let allTimelineItems = [];
    // Ensure notesData and emailsData are arrays before concatenating
    if (Array.isArray(notesData)) { allTimelineItems = allTimelineItems.concat(notesData); }
    if (Array.isArray(emailsData)) { allTimelineItems = allTimelineItems.concat(emailsData); }

    console.log(`Sorting ${allTimelineItems.length} combined timeline items by date...`);
    // Sort the combined array by the parsed dateObject (oldest first)
    allTimelineItems.sort((a, b) => {
        // Use getTime() for reliable numerical comparison of Date objects
        const timeA = a.dateObject?.getTime() || 0; // Use 0 for invalid/missing dates
        const timeB = b.dateObject?.getTime() || 0;
        // Handle cases where one or both dates are invalid/missing
        if (!timeA && !timeB) return 0; // Keep original relative order if both invalid
        if (!timeA) return 1;  // Put items without valid dates after items with valid dates
        if (!timeB) return -1; // Put items without valid dates after items with valid dates
        return timeA - timeB; // Sort numerically by timestamp (ascending order)
    });

    // --- Construct the Final HTML Output String ---
    console.log("Constructing final HTML output string...");
    // Escape data that will be placed directly into HTML attributes or text nodes
    const safeRecordNumber = escapeHtml(recordNumber || 'N/A');
    const safeSubject = escapeHtml(subject || 'N/A');
    const safeObjectType = escapeHtml(objectType);

    // Start building the HTML string using template literals
    let htmlOutput = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${safeObjectType} ${safeRecordNumber}: ${safeSubject}</title>
    <style>
        /* CSS styles for the generated HTML report page */
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            line-height: 1.4;
            padding: 15px 25px; /* Add padding around the content */
            margin: 0;
            color: #333; /* Default text color */
            background-color: #f9f9f9; /* Light background */
        }
        h1, h2 {
            border-bottom: 1px solid #ccc; /* Underline for headers */
            padding-bottom: 6px;
            margin-top: 25px;
            margin-bottom: 15px;
            color: #1a5f90; /* Dark blue header color */
            font-weight: 600; /* Slightly bolder headers */
        }
        h1 { font-size: 1.7em; text-align: left; }
        h2 { font-size: 1.4em; }
        .generation-info { /* Timestamp for generation */
            font-size: 0.8em;
            color: #777; /* Grey color */
            margin-bottom: 20px;
            text-align: right; /* Align to the right */
        }
        .record-details { /* Container for main record details */
            background-color: #fff; /* White background */
            border: 1px solid #e1e5eb; /* Light border */
            padding: 15px 20px;
            border-radius: 5px; /* Rounded corners */
            margin-bottom: 25px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05); /* Subtle shadow */
        }
        .record-details h2 { margin-top: 0; margin-bottom: 12px; } /* Adjust spacing for h2 inside details box */

        /* Grid layout for the main details section */
        .details-grid {
           display: grid;
           grid-template-columns: auto 1fr; /* Label column auto-width, Value column takes remaining space */
           gap: 4px 10px; /* Row gap, Column gap */
           margin-bottom: 15px;
           align-items: start; /* Align items at the top of their grid cell */
       }
       .details-grid dt { /* Definition Term (Label) styling */
           grid-column: 1; /* Place in the first column */
           font-weight: 600; /* Bold label */
           color: #005fb2; /* Salesforce blue */
           text-align: right; /* Align label text to the right */
           padding-right: 8px; /* Space between label and value */
           white-space: nowrap; /* Prevent labels from wrapping */
       }
       .details-grid dd { /* Definition Description (Value) styling */
           grid-column: 2; /* Place in the second column */
           margin-left: 0; /* Reset browser default margin */
           word-wrap: break-word; /* Allow long values to wrap */
           text-align: left; /* Align value text to the left */
       }

        /* Styling for the Description section */
        .description-label {
            font-weight: 600;
            color: #005fb2;
            margin-bottom: 5px;
            display: block; /* Make it a block element */
        }
        .record-details .description-content {
            white-space: pre-wrap; /* Preserve whitespace formatting and wrap lines */
            word-wrap: break-word; /* Break long words if necessary */
            margin-top: 0px;
            padding: 10px 12px;
            background-color: #f1f1f1; /* Light grey background for distinction */
            border-radius: 4px;
            line-height: 1.45;
            font-size: 0.95em;
            max-height: 400px; /* Limit height and enable vertical scrolling */
            overflow-y: auto;
            border: 1px solid #e0e0e0; /* Light border around description */
        }

        /* Styling for individual timeline items (Notes/Emails) */
        .timeline-item {
            border: 1px solid #e1e5eb;
            padding: 12px 18px;
            margin-bottom: 10px;
            border-radius: 5px;
            background-color: #fff;
            box-shadow: 0 1px 2px rgba(0,0,0,0.04);
            position: relative; /* For potential future absolute positioning inside */
        }
        .timeline-item.type-note { border-left: 5px solid #6b92dc; } /* Blue left border for Notes */
        .timeline-item.type-email { border-left: 5px solid #5cb85c; } /* Green left border for Emails */

        /* Header within each timeline item */
        .item-header {
            font-size: 0.95em;
            color: #444;
            margin-bottom: 8px;
            border-bottom: 1px dashed #eee; /* Dashed separator */
            padding-bottom: 6px;
            line-height: 1.4;
        }
        .item-timestamp { /* Timestamp styling */
            color: #555;
            font-family: monospace; /* Monospace font for dates */
            margin-right: 10px;
            font-size: 0.9em;
            background-color:#f0f0f0; /* Light background for timestamp */
            padding: 1px 4px;
            border-radius: 3px;
        }
        .item-type-label { /* "NOTE" or "EMAIL" label */
            font-weight: bold;
            text-transform: uppercase;
            font-size: 0.85em;
            margin-right: 5px;
        }
        .item-type-label.type-note { color: #6b92dc; } /* Match border color */
        .item-type-label.type-email { color: #5cb85c; } /* Match border color */

        .item-subject-title { /* Subject/Title of the item */
            font-weight: 600;
            color: #222; /* Darker color for title */
            margin-left: 4px;
            font-size: 1.05em;
        }
        .item-meta { /* Container for meta info like From/To/By */
            display: block; /* Put meta info on a new line */
            font-size: 0.85em;
            color: #666;
            margin-top: 3px;
        }
        .item-meta-label { /* "From:", "To:", "By:" labels */
            color: #005fb2;
            font-weight: 600;
        }
        .item-meta-info { /* Actual From/To/By values */
            color: #555;
            margin-left: 3px;
        }

        /* Styling for the main content body of a timeline item */
        .item-content {
            white-space: normal; /* Allow wrapping */
            word-wrap: break-word;
            overflow-wrap: break-word; /* Ensure long words break */
            color: #333;
            margin-top: 10px;
            font-size: 0.95em;
            line-height: 1.45;
        }
        /* Basic styling for common HTML elements within the content */
        .item-content p { margin-top: 0; margin-bottom: 0.5em; }
        .item-content strong, .item-content b { font-weight: bold; }
        .item-content em, .item-content i { font-style: italic; }
        .item-content ul { list-style: disc; margin-left: 1.8em; padding-left: 0; margin-top: 0.5em; margin-bottom: 0.5em; }
        .item-content ol { list-style: decimal; margin-left: 1.8em; padding-left: 0; margin-top: 0.5em; margin-bottom: 0.5em; }
        .item-content a { color: #007bff; text-decoration: underline; }
        .item-content blockquote { border-left: 3px solid #ccc; padding-left: 10px; margin-left: 5px; color: #666; font-style: italic; }
        .item-content pre { background-color: #eee; padding: 5px; border-radius: 3px; white-space: pre-wrap; word-wrap: break-word; font-family: monospace; }
        .item-content code { font-family: monospace; background-color: #eee; padding: 1px 3px; border-radius: 2px; }

        /* Styling for attachments placeholder and error messages */
        .item-attachments {
            font-style: italic;
            color: #888;
            font-size: 0.85em;
            margin-top: 10px;
        }
        .error-message { /* Styling for error messages within content */
            color: red;
            font-weight: bold;
            background-color: #ffebeb; /* Light red background */
            border: 1px solid red;
            padding: 5px 8px;
            border-radius: 3px;
            display: inline-block; /* Make it wrap nicely */
            margin-top: 5px;
        }

        /* Styling for Note visibility labels (Public/Internal) */
        .item-visibility {
            margin-left: 8px;
            font-size: 0.9em;
            font-weight: bold;
            text-transform: lowercase;
            padding: 1px 5px;
            border-radius: 3px;
            border: 1px solid transparent; /* Base border */
        }
        .item-visibility.public { /* Reddish style for Public notes */
            color: #8e1b03;
            background-color: #fdd;
            border-color: #fbb;
         }
        .item-visibility.internal { /* Grey style for Internal notes */
            color: #333;
            background-color: #eee;
            border-color: #ddd;
        }
    </style>
</head>
<body>
    <h1>${safeObjectType} ${safeRecordNumber}: ${safeSubject}</h1>
    <div class="generation-info">Generated: ${escapeHtml(generatedTime)}</div>

    <div class="record-details">
        <h2>Details</h2>
        <dl class="details-grid">
            <dt>${safeObjectType}:</dt><dd>${safeRecordNumber}</dd>
            <dt>Customer Account:</dt><dd>${escapeHtml(accountName || 'N/A')}</dd>
            <dt>Subject:</dt><dd>${safeSubject}</dd>
            <dt>Date Created:</dt><dd>${escapeHtml(createdDateStr || 'N/A')}</dd>
            <dt>Created By:</dt><dd>${escapeHtml(creatorName || 'N/A')}</dd>
            <dt>Status:</dt><dd>${escapeHtml(status || 'N/A')}</dd>
            <dt>Owner:</dt><dd>${escapeHtml(owner || 'N/A')}</dd>
        </dl>
        <div class="description-label">Description:</div>
        <div class="description-content">
            ${description || '<p><i>No description found or extracted.</i></p>'}
        </div>
    </div>

    <h2>Timeline (${allTimelineItems.length} items)</h2>
`;

    // Check if there are any items to display in the timeline
    if (allTimelineItems.length === 0) {
        htmlOutput += "<p>No Notes or Emails found or extracted successfully.</p>";
    } else {
        // Iterate through the sorted timeline items and generate HTML for each
        allTimelineItems.forEach(item => {
            // --- Prepare Content HTML ---
            let contentHtml = '';
            let isErrorContent = false; // Flag if content represents an error
            if (item.content && typeof item.content === 'string') {
                // Check for specific error/placeholder strings generated during fetching
                if (item.content.startsWith('Error:') || item.content.startsWith('[Fetch Error') || item.content.startsWith('[Body Fetch Error') || item.content.startsWith('[Content Not Fetched]') || item.content.startsWith('[Content Empty]')) {
                   // Display error messages with specific styling
                   contentHtml = `<span class="error-message">${escapeHtml(item.content)}</span>`;
                   isErrorContent = true;
                } else {
                   // If not an error, assume it's valid HTML content (from Note or Email body)
                   // Pass it through directly without escaping. Scrapers should ensure basic safety if needed.
                   contentHtml = item.content;
                }
            } else {
                 // Handle case where content is missing entirely
                 contentHtml = '<i>[Content Missing]</i>';
                 isErrorContent = true;
            }

            // --- Prepare Visibility Label (for Notes only) ---
            let visibilityLabel = '';
            let visibilityClass = '';
            if (item.type === 'Note') {
                if (item.isPublic === true) { // Explicitly public
                    visibilityClass = 'public';
                    visibilityLabel = `<span class="item-visibility ${visibilityClass}">(public)</span>`;
                } else if (item.isPublic === false) { // Explicitly internal
                    visibilityClass = 'internal';
                    visibilityLabel = `<span class="item-visibility ${visibilityClass}">(internal)</span>`;
                }
                // If item.isPublic is null (e.g., fetch failed, or for Emails), no label is shown.
            }

            // --- Format Timestamp ---
            let formattedTimestamp = 'N/A';
            try {
                // Check if dateObject is a valid Date
                if (item.dateObject && !isNaN(item.dateObject.getTime())) {
                    // Format using locale string (e.g., "4/8/2025, 10:30") with 24-hour time
                    formattedTimestamp = item.dateObject.toLocaleString(undefined, { // Use browser's default locale
                        year: 'numeric', month: 'numeric', day: 'numeric',
                        hour: '2-digit', minute: '2-digit',
                        hour12: false // Force 24-hour format
                    });
                } else if (item.dateStr) {
                    // Fallback to the raw date string from the table if parsing failed
                    formattedTimestamp = escapeHtml(item.dateStr);
                }
            } catch (e) {
                // Catch errors during date formatting
                console.warn("Error formatting date for item:", item.url, e);
                formattedTimestamp = escapeHtml(item.dateStr || 'Date Error'); // Show original string or error
            }

            // --- Prepare other common item details (escaped for safety) ---
            const itemTypeClass = `type-${escapeHtml(item.type?.toLowerCase() || 'unknown')}`;
            const itemTypeLabel = escapeHtml(item.type || 'Item');
            const itemTitle = escapeHtml(item.title || 'N/A');
            const itemAuthor = escapeHtml(item.author || 'N/A'); // 'From' for email, 'Created By' for note
            const itemTo = escapeHtml(item.to || 'N/A'); // Relevant for email

            // --- Construct Header Meta Details based on item type ---
            let headerMetaDetails = '';
            if (item.type === 'Email') {
                headerMetaDetails = `<span class="item-meta"><span class="item-meta-label">From:</span> <span class="item-meta-info">${itemAuthor}</span> | <span class="item-meta-label">To:</span> <span class="item-meta-info">${itemTo}</span></span>`;
            } else { // Assume Note or other types
                headerMetaDetails = `<span class="item-meta"><span class="item-meta-label">By:</span> <span class="item-meta-info"><strong>${itemAuthor}</strong></span></span>`;
            }

            // --- Append the HTML block for this timeline item ---
            htmlOutput += `
            <div class="timeline-item ${itemTypeClass}">
                <div class="item-header">
                    <strong class="item-type-label ${itemTypeClass}">${itemTypeLabel}</strong>
                    ${visibilityLabel} <span class="item-timestamp">[${formattedTimestamp}]</span> -
                    <span class="item-subject-title">${itemTitle}</span>
                    ${headerMetaDetails} </div>
                <div class="item-content">
                    ${contentHtml} </div>
                <div class="item-attachments">
                    Attachments: ${escapeHtml(item.attachments || 'N/A')} </div>
            </div>`;
        }); // End loop through allTimelineItems
    }

    // Close the HTML structure
    htmlOutput += `
</body>
</html>`;

  console.log("HTML generation complete.");
  return htmlOutput; // Return the complete HTML string
} // End of generateCaseViewHtml


// --- Message Listener (Handles messages from background or other parts of the extension) ---
// IMPORTANT: This listener must be at the top level of the script (not inside another function).
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log(`Content Script: Received message action="${message.action}" from sender:`, sender.id, sender.url); // Log incoming messages

  // --- Handler for the generation process (now triggered by background script) ---
  if (message.action === "generateFullCaseView") {
    console.log("Content Script: Handling 'generateFullCaseView' command received from background script.");
    const now = new Date();
    // Format the generation time for display in the report
    const generatedTime = now.toLocaleString('en-US', { dateStyle: 'short', timeStyle: 'medium' });

    // Update panel status
    const statusDiv = document.getElementById('vbsfu-status');
    if (statusDiv) {
        statusDiv.textContent = 'Extracting data...';
        statusDiv.style.color = 'orange';
    }

    // Start the asynchronous HTML generation process
    generateCaseViewHtml(generatedTime)
      .then(fullHtml => {
        // HTML generation successful
        console.log("Content Script: HTML generation complete. Sending HTML to background script to open in new tab.");
        if (statusDiv) {
            statusDiv.textContent = 'Opening results tab...';
            statusDiv.style.color = 'green';
        }
        // Send the generated HTML content to the background script via 'openFullViewTab' message
        chrome.runtime.sendMessage({ action: "openFullViewTab", htmlContent: fullHtml });
        // We don't strictly need to send a response back to the background script here,
        // but could do so if the background needed confirmation.
        // sendResponse({ status: "html_sent_to_background" });
      })
      .catch(error => {
        // HTML generation failed
        console.error("Content Script: Error during generateCaseViewHtml execution:", error);
        if (statusDiv) {
            statusDiv.textContent = 'Error generating view!';
            statusDiv.style.color = 'red';
        }
        // Send an error HTML page to the background script instead
        chrome.runtime.sendMessage({
            action: "openFullViewTab",
            htmlContent: `<html><head><title>Generation Error</title></head><body><h1>Error Generating View</h1><p>An error occurred while generating the report:</p><pre>${escapeHtml(error.message)}</pre></body></html>`
        });
        // Send an error response back to the background script if it was waiting
        // sendResponse({ status: "error", message: error.message });
      });

    // ** IMPORTANT: Return true immediately. **
    // This indicates to the Chrome messaging system that this listener function
    // will respond asynchronously (because generateCaseViewHtml is async).
    // This keeps the message channel open and prevents the "message port closed" error
    // for the message sent *from* the background script *to* this listener.
    return true;
  }

  // --- Handler for getting Case Number and URL (Used by Copy Button - likely redundant now) ---
  // The copy button now calls findCaseNumberSpecific directly. This handler might be removed
  // unless other parts of the extension rely on it.
  if (message.action === "getCaseNumberAndUrl") {
    console.log("Content Script: Handling 'getCaseNumberAndUrl' message (likely redundant).");
    const recordNumber = findCaseNumberSpecific(); // Use the specific function
    const currentUrl = window.location.href;
    if (recordNumber) {
      // Send success response back synchronously
      sendResponse({ status: "success", caseNumber: recordNumber, url: currentUrl });
    } else {
      // Send error response back synchronously
      sendResponse({ status: "error", message: "Could not find Record Number on page." });
    }
    // Return false because sendResponse is called synchronously within this block
    return false;
  }

  // --- Handler for background logging messages ---
  // Listens for progress updates sent from the background script during note/email fetching.
  if (message.action === "logUrlProcessing") {
    console.log(`[BACKGROUND] Progress: Fetching ${message.itemType} ${message.index}/${message.total}: ${message.url}`);
    // Update the status display in the sliding panel
    const statusDiv = document.getElementById('vbsfu-status');
    if (statusDiv) {
        statusDiv.textContent = `Fetching ${message.itemType} ${message.index}/${message.total}...`;
        statusDiv.style.color = 'orange'; // Indicate ongoing work
    }
    // No response needed back to the background script for log messages
    return false;
  }

  // --- Default handling for any messages not matched above ---
  console.log(`Content Script: Received unhandled message action="${message.action}" from sender:`, sender.id, sender.url);
  // Return false as we are handling this synchronously (by doing nothing).
  return false;

}); // End of chrome.runtime.onMessage.addListener

console.log('Content script initial load complete. Top-level message listener attached.');


// --- Sliding Panel Code ---

// --- Function to Create and Inject the Sliding Panel UI ---
function injectSlidingPanel() {
    console.log('Injecting Sliding Panel UI...');

    // Prevent duplicate injection if script runs multiple times somehow
    if (document.getElementById('vbsfu-panel')) {
        console.warn('Sliding panel already exists. Aborting injection.');
        return;
    }

    // --- Create Panel Elements ---
    // Main panel container
    const panel = document.createElement('div');
    panel.id = 'vbsfu-panel';
    // Styling for the panel
    panel.style.position = 'fixed';
    panel.style.top = '100px'; // Position from top edge of viewport
    panel.style.right = '-180px'; // Start hidden off-screen (adjust based on width + padding)
    panel.style.width = '160px'; // Width of the panel
    panel.style.zIndex = '2147483647'; // Max z-index to stay on top
    panel.style.backgroundColor = '#f0f0f0'; // Light grey background
    panel.style.border = '1px solid #ccc'; // Border
    panel.style.borderRight = 'none'; // No border on the edge side
    panel.style.borderRadius = '5px 0 0 5px'; // Rounded corners on the visible side
    panel.style.boxShadow = '-2px 2px 5px rgba(0,0,0,0.2)'; // Shadow for depth
    panel.style.transition = 'right 0.3s ease-in-out'; // Smooth slide animation
    panel.style.padding = '10px'; // Internal padding
    panel.style.fontFamily = 'sans-serif'; // Standard font
    panel.style.fontSize = '14px';
    panel.style.display = 'flex'; // Use flexbox for layout
    panel.style.flexDirection = 'column'; // Stack items vertically
    panel.style.alignItems = 'stretch'; // Make buttons fill width

    // Button to toggle the panel's visibility
    const toggleButton = document.createElement('button');
    toggleButton.id = 'vbsfu-toggle';
    toggleButton.textContent = '🛠️'; // Tool icon
    // Styling for the toggle button
    toggleButton.style.position = 'fixed'; // Fixed position relative to viewport
    toggleButton.style.top = '100px';    // Align with panel's top edge
    toggleButton.style.right = '10px';   // Position near the right edge of the viewport
    toggleButton.style.zIndex = '2147483647'; // Max z-index
    toggleButton.style.padding = '8px';
    toggleButton.style.cursor = 'pointer'; // Indicate clickable
    toggleButton.style.border = '1px solid #fd0';
    toggleButton.style.backgroundColor = '#bde'; // Slightly darker background than panel
    toggleButton.style.color = '#33F'; // Text/icon color
    toggleButton.style.borderRadius = '5px'; // Rounded corners
    toggleButton.style.fontSize = '16px'; // Icon size
    toggleButton.setAttribute('aria-label', 'Toggle Salesforce Utilities Panel'); // Accessibility

    // Panel Title
    const title = document.createElement('h4');
    title.textContent = 'VB SF Utils';
    title.style.textAlign = 'center';
    title.style.marginTop = '0';
    title.style.marginBottom = '10px';
    title.style.color = '#000099';

    // Generate Full View Button
    const generateButton = document.createElement('button');
    generateButton.id = 'vbsfu-generate';
    generateButton.textContent = 'Generate Full View';
    generateButton.style.padding = '8px';
    generateButton.style.marginBottom = '5px'; // Space below button
    generateButton.style.cursor = 'pointer';

    // Copy Case/Record Link Button
    const copyButton = document.createElement('button');
    copyButton.id = 'vbsfu-copy';
    copyButton.textContent = 'Copy Record Link'; // Generic name
    copyButton.style.padding = '8px';
    copyButton.style.marginBottom = '5px';
    copyButton.style.cursor = 'pointer';

    // Status Message Area within the panel
    const statusDiv = document.createElement('div');
    statusDiv.id = 'vbsfu-status';
    statusDiv.style.fontSize = '12px'; // Smaller font for status
    statusDiv.style.marginTop = '8px'; // Space above status
    statusDiv.style.textAlign = 'center';
    statusDiv.style.minHeight = '1em'; // Prevent layout jump when empty
    statusDiv.style.color = '#00802b'; // Default success color

    // --- Assemble Panel and Add to Page ---
    // Add elements to the panel itself
    panel.appendChild(title);
    panel.appendChild(generateButton);
    panel.appendChild(copyButton);
    panel.appendChild(statusDiv);
    // Add the panel and the separate toggle button to the document body
    document.body.appendChild(panel);
    document.body.appendChild(toggleButton);


    // --- Event Listeners for Buttons ---

    // Toggle Panel Visibility onClick
    toggleButton.onclick = () => {
        // Check current state based on 'right' style
        if (panel.style.right === '0px') { // If panel is currently visible
            panel.style.right = '-180px'; // Slide it out (hide)
        } else { // If panel is hidden
            panel.style.right = '0px'; // Slide it in (show)
            statusDiv.textContent = ''; // Clear status message when opening
        }
    };

    // Generate Full View Button onClick
    generateButton.onclick = () => {
        console.log('Panel Generate button clicked');
        statusDiv.textContent = 'Initiating...'; // Initial feedback
        statusDiv.style.color = 'orange';

        // ** Send "initiate" message TO BACKGROUND script **
        console.log("Sending 'initiateGenerateFullCaseView' message to background script.");
        chrome.runtime.sendMessage({ action: "initiateGenerateFullCaseView" }, (response) => {
            // This callback runs IF the background script sends a response OR if an error occurs sending the message
            if (chrome.runtime.lastError) {
                // Handle error sending message TO background script
                console.error("Error sending 'initiateGenerateFullCaseView' message:", chrome.runtime.lastError.message);
                // Display the specific error message if available
                statusDiv.textContent = `Error: ${chrome.runtime.lastError.message}`;
                statusDiv.style.color = 'red';
            } else {
                // Background script acknowledged the initiation request (response might be undefined if background didn't send one)
                console.log("Background script acknowledged initiation. Response:", response);
                // Update status. Further updates will come via the 'logUrlProcessing' message handler.
                statusDiv.textContent = "Processing initiated...";
                statusDiv.style.color = 'orange';
            }
        });
    }; // End generateButton.onclick

    // Copy Case/Record Link Button onClick
    copyButton.onclick = () => {
        console.log('Panel Copy button clicked');
        statusDiv.textContent = ''; // Clear previous status
        statusDiv.style.color = 'green'; // Reset color

        // Get record number and URL using functions available in this script
        const recordNumber = findCaseNumberSpecific();
        const currentUrl = window.location.href;

        // Determine object type for the link text
        let objectType = 'Record'; // Default
         if (currentUrl.includes('/Case/')) objectType = 'Case';
         else if (currentUrl.includes('/WorkOrder/')) objectType = 'WorkOrder';
         // Add more types if needed

        // Proceed only if both record number and URL were found
        if (recordNumber && currentUrl) {
            const linkText = `${objectType} ${recordNumber}`; // Dynamic link text
            const richTextHtml = `<a href="${currentUrl}">${linkText}</a>`; // HTML for rich text clipboard
            console.log(`Attempting to copy: ${richTextHtml}`);

            try {
                // Create Blob objects for HTML and plain text formats
                const blobHtml = new Blob([richTextHtml], { type: 'text/html' });
                const blobText = new Blob([linkText], { type: 'text/plain' });

                // Create a ClipboardItem containing both formats
                const clipboardItem = new ClipboardItem({
                    'text/html': blobHtml,
                    'text/plain': blobText // Plain text fallback
                });

                // Use the Clipboard API to write the item
                navigator.clipboard.write([clipboardItem]).then(() => {
                    // Success callback
                    console.log('Rich text link copied to clipboard!');
                    statusDiv.textContent = `Copied: ${linkText}`; // Display success message
                }).catch(err => {
                    // Error callback for clipboard write
                    console.error('Failed to copy rich text: ', err);
                    // Attempt fallback to copying plain text only
                    navigator.clipboard.writeText(linkText).then(() => {
                        console.log('Fallback: Copied plain text link label!');
                        statusDiv.textContent = `Copied text: ${linkText}`;
                    }).catch(fallbackErr => {
                        // Error callback for plain text fallback attempt
                        console.error('Failed to copy plain text fallback: ', fallbackErr);
                        statusDiv.textContent = 'Error: Copy failed.';
                        statusDiv.style.color = 'red';
                    });
                });
            } catch (error) {
                // Catch errors related to the Clipboard API itself (e.g., browser support, permissions)
                console.error('Clipboard API error:', error);
                statusDiv.textContent = 'Error: Clipboard API failed.';
                statusDiv.style.color = 'red';
            }
        } else {
            // Handle case where record number couldn't be found
            console.error("Failed to get record number/URL for copy.");
            statusDiv.textContent = 'Error: Record number not found.';
            statusDiv.style.color = 'red';
        }
    }; // End copyButton.onclick

    console.log('Sliding Panel UI Injected and event listeners attached.');
} // End of injectSlidingPanel


// --- Initial Check and Injection Trigger ---
// Define the URL pattern to target specific Salesforce record pages (Case, WorkOrder)
// ** Ensure this is declared only ONCE at the top level **
const urlPattern = /\.lightning\.force\.com\/lightning\/r\/(Case|WorkOrder)\//;
// Check if the current page URL matches the pattern
if (urlPattern.test(window.location.href)) {
    // Use setTimeout to delay injection slightly. This helps ensure that the
    // Salesforce page has finished its initial rendering, making DOM manipulation more reliable.
    // 1500ms (1.5 seconds) is a starting point; adjust if needed based on page load performance.
    setTimeout(injectSlidingPanel, 1500);
} else {
    // Log if the script is running on a page that doesn't match the pattern
    console.log('Not a target Salesforce record page (Case/WorkOrder), panel not injected.');
}

console.log("Content Script: End of script execution. Waiting for messages or UI interaction.");
