// content.js - Complete File (v50 + Copy Feature + Detail Layout Fix)

console.log("Salesforce Case Full View Content Script Loaded (v50 + Copy Feature + Detail Layout Fix).");

// --- Helper Functions ---
/**
 * Waits for an element matching the selector to appear in the DOM.
 */
function waitForElement(selector, baseElement = document, timeout = 15000) {
  // console.log(`Waiting for "${selector}"...`);
  return new Promise((resolve) => {
    const startTime = Date.now();
    const interval = setInterval(() => {
      const element = baseElement.querySelector(selector);
      if (element) {
        // console.log(`waitForElement: Found "${selector}"`);
        clearInterval(interval);
        resolve(element);
      } else if (Date.now() - startTime > timeout) {
        console.warn(`waitForElement: Timeout waiting for "${selector}"`);
        clearInterval(interval);
        resolve(null);
      }
    }, 300);
  });
}

/**
 * Wraps chrome.runtime.sendMessage in a Promise for async/await usage.
 */
function sendMessagePromise(message) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(message, (response) => {
      if (chrome.runtime.lastError) {
        console.error("sendMessagePromise failed:", chrome.runtime.lastError);
        reject(chrome.runtime.lastError);
      } else {
        resolve(response);
      }
    });
  });
}

function escapeHtml(unsafe) {
  if (typeof unsafe !== 'string') {
      return unsafe === null || typeof unsafe === 'undefined' ? '' : String(unsafe);
  }
  return unsafe
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function getTextContentFromElement(element, selector) {
  if (!element) return 'N/A';
  const childElement = element.querySelector(selector);
  return childElement ? childElement.textContent?.trim() : 'N/A';
}

/**
 * Parses a date string (DD/MM/YYYY HH:MM) into a Date object.
 */
function parseDateStringFromTable(dateString) {
  if (!dateString) return null;
  let match = dateString.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/); // DD/MM/YYYY HH:MM
  if (match) {
    try {
      const day = parseInt(match[1]);
      const month = parseInt(match[2]) - 1; // Month is 0-indexed
      const year = parseInt(match[3]);
      const hour = parseInt(match[4]);
      const minute = parseInt(match[5]);

      // Basic validation
      if (year > 1970 && month >= 0 && month < 12 && day >= 1 && day <= 31 && hour >= 0 && hour < 24 && minute >= 0 && minute < 60) {
        const dateObject = new Date(Date.UTC(year, month, day, hour, minute));
        if (!isNaN(dateObject.getTime())) {
            return dateObject;
        } else {
             console.warn(`CS Date Invalid Comp: "${dateString}" resulted in invalid Date object`);
             return null;
        }
      } else {
        console.warn(`CS Date Comp Range Invalid: "${dateString}"`);
        return null;
      }
    } catch (e) {
      console.error(`CS Error creating Date from matched parts: "${dateString}"`, e);
      return null;
    }
  } else {
    // Fallback attempt if regex doesn't match
    const parsedFallback = Date.parse(dateString);
    if (!isNaN(parsedFallback)) {
        return new Date(parsedFallback); // Note: Might not be UTC, depends on browser interpretation
    } else {
        console.warn(`CS Could not parse date format with regex or Date.parse: "${dateString}"`);
        return null;
    }
  }
}


// --- Functions to Extract Case Details (REVISED Selectors from v50) ---

function findSubjectInContainer(container) {
    if (!container) return 'N/A';
    // Strategy: Find the unique LWC for subject, then the text inside.
    const element = container.querySelector('support-output-case-subject-field lightning-formatted-text');
    // console.log("Finding Subject:", element ? 'Found' : 'Not Found');
    return element ? element.textContent?.trim() : 'N/A (Subject)';
}

// Older function, kept for reference but likely less reliable than findCaseNumberSpecific
function findCaseNumberInContainer(container) {
     if (!container) return 'N/A';
     // Strategy: Assume it's the first item AND directly contains lightning-formatted-text
     // This is positional and potentially fragile.
     // console.log("Finding Case Number (assuming 1st item)...");
     const firstItem = container.querySelector('records-highlights-details-item:nth-of-type(1)');
     const element = firstItem?.querySelector('p.fieldComponent lightning-formatted-text');
     // Check it doesn't contain other known unique components found in other fields
     if (element &&
         !firstItem.querySelector('support-output-case-subject-field') &&
         !firstItem.querySelector('records-formula-output') &&
         !firstItem.querySelector('force-lookup') )
     {
         // console.log("Case Number element found:", !!element);
         return element.textContent?.trim() || 'N/A (Case #)';
     }
     console.warn("Case Number selector (1st item, specific checks) failed.");
     return 'N/A (Case #)';
}

// --- Helper function specifically for the Case Number based on user's markup ---
// This targets the specific structure provided more reliably than the older findCaseNumberInContainer
function findCaseNumberSpecific() {
    console.log("Content Script: Attempting to find Case Number (Specific Selector)...");
    // Find the item containing the title "Case Number", then get the formatted text within it.
    const itemSelector = 'records-highlights-details-item:has(p[title="Case Number"])';
    const textSelector = 'lightning-formatted-text';

    const detailsItem = document.querySelector(itemSelector);
    if (detailsItem) {
        const textElement = detailsItem.querySelector(textSelector);
        if (textElement) {
            const caseNum = textElement.textContent?.trim();
            // Optional: Add a simple validation like checking if it's digits
            if (caseNum && /^\d+$/.test(caseNum)) {
                 console.log("Content Script: Found Case Number:", caseNum);
                 return caseNum;
            } else {
                 console.warn("Content Script: Found element, but content doesn't look like a case number:", caseNum);
            }
        } else {
             console.warn("Content Script: Found details item, but inner text element not found with selector:", textSelector);
        }
    } else {
         console.warn("Content Script: Case Number details item not found with selector:", itemSelector);
         // Fallback attempt using the older function IF the new one fails
         console.log("Content Script: Falling back to older case number detection...");
         const container = document.querySelector('records-highlights2'); // Assuming this is still the main container
         return findCaseNumberInContainer(container);
    }
    return null; // Return null if not found or invalid after all attempts
}


function findStatusInContainer(container) {
    if (!container) return 'N/A';
    // console.log("Attempting to find Status (item with specific rich text/img)...");
    // Strategy: Find the item containing the rich text formula output, specifically looking for the span part
    const statusItem = container.querySelector('records-highlights-details-item:has(records-formula-output lightning-formatted-rich-text)');
    const element = statusItem?.querySelector('lightning-formatted-rich-text span[part="formatted-rich-text"]');
    // console.log("Status element found:", !!element);
    return element ? element.textContent?.trim() : 'N/A (Status)';
}

async function findCreatorName() { // Searches document using field-label
    // console.log("Finding Creator Name (System Info)...");
    const createdByItem = await waitForElement('records-record-layout-item[field-label="Created By"]');
    if (!createdByItem) { console.warn("Creator layout item not found."); return 'N/A (Creator)'; }
    const nameElement = createdByItem.querySelector('force-lookup a');
    // console.log("Creator name element found:", !!nameElement);
    return nameElement ? nameElement.textContent?.trim() : 'N/A (Creator)';
}

async function findCreatedDate() { // Searches document using field-label
    // console.log("Finding Created Date (System Info)...");
    const createdByItem = await waitForElement('records-record-layout-item[field-label="Created By"]');
    if (!createdByItem) { console.warn("Created Date layout item not found."); return 'N/A (Created Date)'; }
    const dateElement = createdByItem.querySelector('records-modstamp lightning-formatted-text');
     // console.log("Created Date element found:", !!dateElement);
    return dateElement ? dateElement.textContent?.trim() : 'N/A (Created Date)';
}

function findOwnerInContainer(container) {
     if (!container) return 'N/A';
     // console.log("Finding Owner (force-lookup a)...");
     // Strategy: Find the specific item containing force-lookup, then the link.
     const ownerItem = container.querySelector('records-highlights-details-item:has(force-lookup)');
     const element = ownerItem?.querySelector('force-lookup a');
     // console.log("Owner element found:", !!element);
     return element ? element.textContent?.trim() : 'N/A (Owner)';
}

// ** NEW: Function to find Account Name from main details section **
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

async function findCaseDescription() {
    // console.log("Attempting to find Case Description...");
     const descriptionContainer = await waitForElement('article.cPSM_Case_Description');
     if (!descriptionContainer) return '';
     let textElement = descriptionContainer.querySelector('lightning-formatted-text.txtAreaReadOnly') || descriptionContainer.querySelector('lightning-formatted-text');
     if (!textElement) return '';
     // Check for "View More" button specifically within the description container
     const viewMoreButton = descriptionContainer.querySelector('button.slds-button:not([disabled])');
     let descriptionHTML = '';

     if (viewMoreButton && (viewMoreButton.textContent.includes('View More') || viewMoreButton.textContent.includes('Show More'))) {
        // console.log("Clicking 'View More' for description...");
        viewMoreButton.click();
        // Wait briefly for content to potentially expand
        await new Promise(resolve => setTimeout(resolve, 500));
        // Re-query for the text element after clicking
        let updatedTextElement = descriptionContainer.querySelector('lightning-formatted-text.txtAreaReadOnly') || descriptionContainer.querySelector('lightning-formatted-text');
        descriptionHTML = updatedTextElement?.innerHTML?.trim() || '';
     } else {
        descriptionHTML = textElement?.innerHTML?.trim() || '';
     }
     return descriptionHTML;
}


// --- Function to Extract Note URLs from Table (to send to background) ---
async function extractAndFetchNotes() {
  const notesHeaderSelector = 'a.slds-card__header-link[href*="/related/PSM_Notes__r/view"]';
  // console.log(`Finding Notes header link ("${notesHeaderSelector}")...`);
  const headerLinkElement = await waitForElement(notesHeaderSelector);
  if (!headerLinkElement) { console.warn(`Notes header link not found.`); return []; }

  const headerElement = headerLinkElement.closest('lst-list-view-manager-header');
  if (!headerElement) { console.warn("Notes parent header element not found."); return []; }

  // Check count directly in the header link's span
  const countSpan = headerLinkElement.querySelector('span[title*="("]');
  if (countSpan && countSpan.textContent?.trim() === '(0)') { console.log("Notes count is (0), skipping."); return []; }

  // Find the common container for the list (might vary slightly)
  let commonContainer = headerElement.closest('lst-common-list-internal') || headerElement.closest('lst-related-list-view-manager');
  if (!commonContainer) { console.error("Cannot find Notes container (lst-common-list-internal or lst-related-list-view-manager)."); return []; }

  // Check for and click 'View All' if present and visible
  const listManager = commonContainer.closest('lst-related-list-view-manager') || commonContainer;
  const viewAllLink = listManager?.querySelector('a.slds-card__footer[href*="/related/PSM_Notes__r/view"]');
  if (viewAllLink && viewAllLink.offsetParent !== null) { // Check if visible
    console.log("Clicking 'View All' Notes...");
    viewAllLink.click();
    await new Promise(resolve => setTimeout(resolve, 1500)); // Wait for potential load/update
  }

  // Wait for the datatable within the container
  const dataTable = await waitForElement('lightning-datatable', commonContainer, 10000);
  if (!dataTable) { console.warn("Notes datatable not found inside container."); return []; }

  const tableBody = await waitForElement('tbody[data-rowgroup-body]', dataTable, 5000);
  if (!tableBody) { console.warn("Notes table body not found."); return []; }

  const rows = tableBody.querySelectorAll('tr[data-row-key-value]');
  const notesToFetch = [];
  const origin = window.location.origin;

  rows.forEach((row, index) => {
    const noteLink = row.querySelector('th[data-label="PSM Note Name"] a');
    const authorLink = row.querySelector('td[data-label="Created By"] a');
    const dateSpan = row.querySelector('td[data-label="Created Date"] lst-formatted-text span');
    const snippetEl = row.querySelector('td[data-label="Description"] lightning-base-formatted-text');

    const relativeUrl = noteLink?.getAttribute('href');
    const noteDateStr = dateSpan?.title || dateSpan?.textContent?.trim(); // Use title first if available

    if (relativeUrl && noteDateStr) {
      notesToFetch.push({
        type: 'Note',
        url: new URL(relativeUrl, origin).href,
        title: noteLink?.textContent?.trim() || 'N/A',
        author: authorLink?.textContent?.trim() || 'N/A',
        dateStr: noteDateStr,
        descriptionSnippet: snippetEl?.textContent?.trim() || ''
      });
    } else {
      console.warn(`Skipping Note row ${index+1}: Missing URL (${!!relativeUrl}) or Date (${!!noteDateStr}).`);
    }
  });

  if (notesToFetch.length === 0) {
    console.log("No valid Note rows found in table.");
    return [];
  }

  console.log(`Sending ${notesToFetch.length} Notes to background for detail fetching...`);
  try {
    const response = await sendMessagePromise({ action: "fetchItemDetails", items: notesToFetch });
    if (response?.status === "success" && response.details) {
      console.log(`Received details for ${Object.keys(response.details).length} Notes from background.`);
      return notesToFetch.map(noteInfo => {
        const fetched = response.details[noteInfo.url];
        // Prioritize date from background (already parsed), fallback to table string
        const finalDateObject = fetched?.dateObject ? new Date(fetched.dateObject) : parseDateStringFromTable(noteInfo.dateStr);
        let finalDescription = fetched?.description;
        // Use snippet as fallback ONLY if fetch resulted in an error string
        if (finalDescription?.startsWith('Error:')) {
          finalDescription = noteInfo.descriptionSnippet ? `[Fetch Error, Snippet: ${noteInfo.descriptionSnippet}]` : finalDescription;
        }
        return {
          type: 'Note',
          url: noteInfo.url,
          title: noteInfo.title,
          author: noteInfo.author,
          dateStr: noteInfo.dateStr, // Keep original string for display if needed
          dateObject: finalDateObject, // Parsed date object for sorting
          content: finalDescription || '[Content Not Fetched]', // Use fetched content
          attachments: 'N/A' // Placeholder
        };
      }).filter(note => note.dateObject !== null); // Filter out notes where date parsing failed entirely
    } else {
      throw new Error(response?.message || "Invalid response or missing details for Notes from background");
    }
  } catch (error) {
    console.error("Error sending/receiving Note details to/from background:", error);
    return []; // Return empty array on error
  }
}


// --- Function to Extract Email Info from Table and Trigger Fetch ---
async function extractAndFetchEmails() {
  const emailsListContainerSelector = 'div.forceListViewManager[aria-label*="Emails"]';
  // console.log(`Finding Emails container ("${emailsListContainerSelector}")...`);
  const emailsListContainer = await waitForElement(emailsListContainerSelector);
  if (!emailsListContainer) { console.warn(`Emails container not found.`); return []; }

  // Check count in header if possible
  const headerLinkElement = emailsListContainer.querySelector('lst-list-view-manager-header a.slds-card__header-link[title="Emails"]');
  if (headerLinkElement) {
      const countSpan = headerLinkElement.querySelector('span[title*="("]');
      if (countSpan && countSpan.textContent?.trim() === '(0)') {
          console.log("Emails count is (0), skipping.");
          return [];
      }
  }

  // Click 'View All' if present and visible
  const parentCard = emailsListContainer.closest('.forceRelatedListSingleContainer');
  const viewAllLink = parentCard?.querySelector('a.slds-card__footer[href*="/related/EmailMessages/view"]');
  if (viewAllLink && viewAllLink.offsetParent !== null) {
    console.log("Clicking 'View All' Emails...");
    viewAllLink.click();
    await new Promise(resolve => setTimeout(resolve, 1500));
  }

  // Wait for table (might be different selector)
  const dataTable = await waitForElement('table.uiVirtualDataTable', emailsListContainer, 10000);
  if (!dataTable) { console.warn("Emails table (uiVirtualDataTable) not found."); return []; }

  const tableBody = await waitForElement('tbody', dataTable, 5000);
  if (!tableBody) { console.warn("Emails table body not found."); return []; }

  const rows = tableBody.querySelectorAll('tr');
  const emailsToFetch = [];
  const origin = window.location.origin;

  rows.forEach((row, index) => {
    const cells = row.querySelectorAll('th, td');
    if (cells.length < 5) { // Basic check for enough cells
      console.warn(`Skipping email row ${index+1} due to insufficient cells (${cells.length}).`);
      return;
    }
    // Selectors based on common structure (might need adjustment)
    const subjectLink = cells[1]?.querySelector('a.outputLookupLink');
    const fromEl = cells[2]?.querySelector('a.emailuiFormattedEmail'); // Often a link
    const toEl = cells[3]?.querySelector('span.uiOutputText'); // Often just text
    const dateEl = cells[4]?.querySelector('span.uiOutputDateTime');

    const relativeUrl = subjectLink?.getAttribute('href');
    const emailSubject = subjectLink?.textContent?.trim();
    const emailDateStr = dateEl?.textContent?.trim();

    if (relativeUrl && emailSubject && emailDateStr) {
      emailsToFetch.push({
        type: 'Email',
        url: new URL(relativeUrl, origin).href,
        title: emailSubject,
        author: fromEl?.textContent?.trim() || 'N/A',
        to: toEl?.title || toEl?.textContent?.trim() || 'N/A', // Use title attr if available
        dateStr: emailDateStr
      });
    } else {
      console.warn(`Skipping Email row ${index+1}: Missing URL (${!!relativeUrl}), Subject (${!!emailSubject}), or Date (${!!emailDateStr}).`);
    }
  });

   if (emailsToFetch.length === 0) {
    console.log("No valid Email rows found in table.");
    return [];
  }

  console.log(`Sending ${emailsToFetch.length} Emails to background for detail fetching...`);
  try {
    const response = await sendMessagePromise({ action: "fetchItemDetails", items: emailsToFetch });
    if (response?.status === "success" && response.details) {
      console.log(`Received details for ${Object.keys(response.details).length} Emails from background.`);
      return emailsToFetch.map(emailInfo => {
        const fetched = response.details[emailInfo.url];
        // Prioritize date from background, fallback to table string
        const finalDateObject = fetched?.dateObject ? new Date(fetched.dateObject) : parseDateStringFromTable(emailInfo.dateStr);
        let finalContent = fetched?.description;
         // Handle fetch error case
        if (finalContent?.startsWith('Error:')) {
            finalContent = `Subject: ${escapeHtml(fetched?.subject || emailInfo.title)}<br>[Body Fetch Error: ${escapeHtml(finalContent)}]`;
        } else if (!finalContent) {
             finalContent = `Subject: ${escapeHtml(fetched?.subject || emailInfo.title)}<br>[Body Not Fetched or Empty]`;
        }
        return {
          type: 'Email',
          url: emailInfo.url,
          title: fetched?.subject || emailInfo.title, // Use fetched subject if available
          author: fetched?.from || emailInfo.author, // Use fetched from if available
          to: fetched?.to || emailInfo.to, // Use fetched to if available
          dateStr: emailInfo.dateStr,
          dateObject: finalDateObject,
          content: finalContent, // Use fetched content (HTML body)
          attachments: 'N/A' // Placeholder
        };
      }).filter(email => email.dateObject !== null); // Filter out emails where date parsing failed
    } else {
      throw new Error(response?.message || "Invalid response or missing details for Emails from background");
    }
  } catch (error) {
    console.error("Error sending/receiving Email details to/from background:", error);
    return [];
  }
}


// --- Main Function to Orchestrate Extraction and Generate HTML ---
async function generateCaseViewHtml(generatedTime) {
  console.log("Starting async HTML generation (v50 + Copy Feature + Detail Layout Fix)...");

  // Optional scroll to ensure dynamic content might load, but less critical if using waitForElement robustly
  // console.log("Scrolling to bottom (quick)...");
  window.scrollTo({ top: document.body.scrollHeight, behavior: 'auto' });
  await new Promise(resolve => setTimeout(resolve, 500)); // Shorter wait after scroll
  // console.log("Scroll attempt complete.");

  // --- Find Main Case Details Container ---
  const CASE_CONTAINER_SELECTOR = 'records-highlights2'; // Selector for the main highlights area
  const caseContainerElement = await waitForElement(CASE_CONTAINER_SELECTOR);
  if (!caseContainerElement) {
      console.error(`FATAL: Case highlights container ("${CASE_CONTAINER_SELECTOR}") not found. Cannot extract details.`);
      // Maybe return a minimal error HTML?
      return `<html><body><h1>Error</h1><p>Could not find the main case details section on the page.</p></body></html>`;
  } else {
      console.log("Main Case highlights container found.");
  }

  // --- Extract Header Details Concurrently ---
  console.log("Extracting Case header details...");
  const subjectPromise = Promise.resolve(findSubjectInContainer(caseContainerElement)); // Sync call wrapped
  const caseNumberPromise = Promise.resolve(findCaseNumberSpecific()); // Use the NEW specific function (sync)
  const statusPromise = Promise.resolve(findStatusInContainer(caseContainerElement)); // Sync
  const ownerPromise = Promise.resolve(findOwnerInContainer(caseContainerElement)); // Sync
  const creatorNamePromise = findCreatorName(); // Async
  const createdDateCaseStrPromise = findCreatedDate(); // Async
  const caseDescriptionPromise = findCaseDescription(); // Async
  const accountNamePromise = findAccountName(); // VB
    // console.log(`accountName is ${accountNamePromise}`); // VB
    
  // --- Extract Related Lists Concurrently ---
  console.log("--- Starting Notes & Emails Extraction Concurrently ---");
  const notesPromise = extractAndFetchNotes(); // Async
  const emailsPromise = extractAndFetchEmails(); // Async

  // --- Wait for all extractions to complete ---
  const [
      subject, caseNumber, status, owner, creatorName, accountName, createdDateCaseStr, caseDescription,
      notesData, emailsData
  ] = await Promise.all([
      subjectPromise, caseNumberPromise, statusPromise, ownerPromise, creatorNamePromise, accountNamePromise, createdDateCaseStrPromise, caseDescriptionPromise,
      notesPromise, emailsPromise
  ]);

  console.log(`Extraction results: Subject=${!!subject}, Case#=${caseNumber}, Status=${!!status}, Owner=${!!owner}, Creator=${!!creatorName}, AccountNamer=${!!accountName},  CreatedDate=${!!createdDateCaseStr}, Desc=${caseDescription?.length>0}, Notes=${notesData.length}, Emails=${emailsData.length}`);


  // --- Combine and Sort Notes & Emails ---
  let allItems = [];
  if (notesData) { allItems = allItems.concat(notesData); }
  if (emailsData) { allItems = allItems.concat(emailsData); }

  console.log(`Sorting ${allItems.length} combined items by date...`);
  allItems.sort((a, b) => {
      const dateA = a.dateObject?.getTime() || 0; // Use getTime() for comparison
      const dateB = b.dateObject?.getTime() || 0;
      if (!dateA && !dateB) return 0; // Both invalid/missing
      if (!dateA) return 1;  // Put items without dates last
      if (!dateB) return -1; // Put items without dates last
      return dateA - dateB; // Sort oldest first
  });

  console.log("Constructing final HTML output...");
  // --- Construct the Final HTML Output ---
  // Use textContent for safety in title, escape for display
  const safeCaseNumber = escapeHtml(caseNumber || 'N/A');
  const safeSubject = escapeHtml(subject || 'N/A');

  let html = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Case ${safeCaseNumber}: ${safeSubject}</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; line-height: 1.4; padding: 15px 25px; margin: 0; color: #333; background-color: #f9f9f9; }
        h1, h2 { border-bottom: 1px solid #ccc; padding-bottom: 6px; margin-top: 25px; margin-bottom: 15px; color: #1a5f90; font-weight: 600; }
        h1 { font-size: 1.7em; text-align: left; }
        h2 { font-size: 1.4em; }
        .generation-info { font-size: 0.8em; color: #777; margin-bottom: 20px; text-align: right; }
        .case-details { background-color: #fff; border: 1px solid #e1e5eb; padding: 15px 20px; border-radius: 5px; margin-bottom: 25px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .case-details h2 { margin-top: 0; margin-bottom: 12px; }

        /* --- UPDATED CSS FOR CASE DETAILS GRID --- */
        .case-details-grid {
           display: grid;
           grid-template-columns: auto 1fr; /* Always 2 columns: Label | Value */
           gap: 4px 10px; /* Adjust gap */
           margin-bottom: 15px;
           align-items: start; /* Align items to the top of their cell */
       }
       /* REMOVED @media query changing columns */
       .case-details-grid dt {
           grid-column: 1; /* Label in column 1 */
           font-weight: 600;
           color: #005fb2;
           text-align: right;
           padding-right: 8px;
           white-space: nowrap; /* Prevent label from wrapping */
       }
       .case-details-grid dd {
           grid-column: 2; /* Value in column 2 */
           margin-left: 0;
           word-wrap: break-word; /* Allow long values to wrap */
           text-align: left; /* Ensure values align left */
       }
       /* --- END OF UPDATED CSS --- */

        .case-description-label { font-weight: 600; color: #005fb2; margin-bottom: 5px; display: block; }
        .case-details .description-content { white-space: pre-wrap; word-wrap: break-word; margin-top: 0px; padding: 10px 12px; background-color: #f1f1f1; border-radius: 4px; line-height: 1.45; font-size: 0.95em; max-height: 400px; overflow-y: auto; }
        .timeline-item { border: 1px solid #e1e5eb; padding: 12px 18px; margin-bottom: 10px; border-radius: 5px; background-color: #fff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); position: relative; }
        .timeline-item.type-note { border-left: 5px solid #6b92dc; }
        .timeline-item.type-email { border-left: 5px solid #5cb85c; }
        .item-header { font-size: 0.95em; color: #444; margin-bottom: 8px; border-bottom: 1px dashed #eee; padding-bottom: 6px; line-height: 1.4; }
        .item-timestamp { color: #555; font-family: monospace; margin-right: 10px; font-size: 0.9em; background-color:#f0f0f0; padding: 1px 4px; border-radius: 3px;}
        .item-type-label { font-weight: bold; text-transform: uppercase; font-size: 0.85em; margin-right: 5px; }
        .item-type-label.type-note { color: #6b92dc; }
        .item-type-label.type-email { color: #5cb85c; }
        .item-subject-title { font-weight: 600; color: #222; margin-left: 4px; font-size: 1.05em;}
        .item-meta { display: block; font-size: 0.85em; color: #666; margin-top: 3px; }
        .item-meta-label { color: #005fb2; font-weight: 600; }
        .item-meta-info { color: #555; margin-left: 3px; }
        .item-content { white-space: normal; word-wrap: break-word; color: #333; margin-top: 10px; font-size: 0.95em; line-height: 1.45; }
        .item-content p { margin-top: 0; margin-bottom: 0.5em; }
        .item-content strong { font-weight: bold; }
        .item-content em { font-style: italic; }
        .item-content ul { list-style: disc; margin-left: 1.8em; padding-left: 0; margin-top: 0.5em; margin-bottom: 0.5em; }
        .item-content ol { list-style: decimal; margin-left: 1.8em; padding-left: 0; margin-top: 0.5em; margin-bottom: 0.5em; }
        .item-content a { color: #007bff; text-decoration: underline; }
        .item-content blockquote { border-left: 3px solid #ccc; padding-left: 10px; margin-left: 5px; color: #666; font-style: italic; }
        .item-attachments { font-style: italic; color: #888; font-size: 0.85em; margin-top: 10px; }
        .error-message { color: red; font-weight: bold; background-color: #ffebeb; border: 1px solid red; padding: 5px; border-radius: 3px;}
    </style>
</head>
<body>
    <h1>Case ${safeCaseNumber}: ${safeSubject}</h1>
    <div class="generation-info">Generated: ${escapeHtml(generatedTime)}</div>

    <div class="case-details">
        <h2>Case Details</h2>
        <dl class="case-details-grid">
            <dt>Case Number:</dt><dd>${safeCaseNumber}</dd> 
            <dt>Customer Account:</dt><dd>${escapeHtml(accountName || 'N/A')}</dd>
            <dt>Subject:</dt><dd>${safeSubject}</dd>
            <dt>Date Created:</dt><dd>${escapeHtml(createdDateCaseStr || 'N/A')}</dd>
            <dt>Created By:</dt><dd>${escapeHtml(creatorName || 'N/A')}</dd>
            <dt>Status:</dt><dd>${escapeHtml(status || 'N/A')}</dd>
            <dt>Owner:</dt><dd>${escapeHtml(owner || 'N/A')}</dd>
        </dl>
        <div class="case-description-label">Description:</div>
        <div class="description-content">${caseDescription || '<p><i>No description found or extracted.</i></p>'}</div>
    </div>

    <h2>Notes and Emails (${allItems.length} items)</h2>
`;

    if (allItems.length === 0) {
        html += "<p>No Notes or Emails found or extracted successfully.</p>";
    } else {
        allItems.forEach(item => {
            // Safely get content, handle potential errors stored in content
            let contentHtml = '';
            let isErrorContent = false;
            if (item.content && typeof item.content === 'string') {
                if (item.content.startsWith('Error:') || item.content.startsWith('[Fetch Error') || item.content.startsWith('[Body Fetch Error') || item.content.startsWith('[Content Not Fetched]')) {
                   contentHtml = `<span class="error-message">${escapeHtml(item.content)}</span>`;
                   isErrorContent = true;
                } else {
                   // Assume it's HTML content if not an error string
                   contentHtml = item.content; // Don't escape HTML content from emails/notes
                }
            } else {
                 contentHtml = '<i>[Content Missing]</i>';
                 isErrorContent = true;
            }


            let headerDetails = '';
            let formattedTimestamp = 'N/A';
            try {
                if (item.dateObject && !isNaN(item.dateObject.getTime())) {
                    // Use locale string for better readability
                    formattedTimestamp = item.dateObject.toLocaleString(undefined, {
                        year: 'numeric', month: 'numeric', day: 'numeric',
                        hour: '2-digit', minute: '2-digit', // Removed seconds
                        hour12: false // Use 24-hour format
                    });
                } else {
                    formattedTimestamp = escapeHtml(item.dateStr || 'N/A'); // Fallback to raw string
                }
            } catch (e) {
                console.warn("Error formatting date for item:", item.url, e);
                formattedTimestamp = escapeHtml(item.dateStr || 'Date Error');
            }

            const itemTypeClass = `type-${escapeHtml(item.type?.toLowerCase() || 'unknown')}`;
            const itemTypeLabel = escapeHtml(item.type || 'Item');
            const itemTitle = escapeHtml(item.title || 'N/A');
            const itemAuthor = escapeHtml(item.author || 'N/A');
            const itemTo = escapeHtml(item.to || 'N/A');

            if (item.type === 'Email') {
                headerDetails = `<span class="item-meta"><span class="item-meta-label">From:</span> <span class="item-meta-info">${itemAuthor}</span> | <span class="item-meta-label">To:</span> <span class="item-meta-info">${itemTo}</span></span>`;
            } else { // Note
                headerDetails = `<span class="item-meta"><span class="item-meta-label">By:</span> <span class="item-meta-info"><strong>${itemAuthor}</strong></span></span>`;
            }

            html += `
            <div class="timeline-item ${itemTypeClass}">
                <div class="item-header">
                    <span class="item-timestamp">[${formattedTimestamp}]</span>
                    <strong class="item-type-label ${itemTypeClass}">${itemTypeLabel}</strong>:
                    <span class="item-subject-title">${itemTitle}</span>
                    ${headerDetails}
                </div>
                <div class="item-content">
                    ${contentHtml} 
                </div>
                <div class="item-attachments">
                    Attachments: ${escapeHtml(item.attachments || 'N/A')}
                </div>
            </div>`;
        });
    }

    html += `
</body>
</html>`;

  console.log("HTML generation complete.");
  return html;
} // End of generateCaseViewHtml


// --- Message Listener (Handles both actions) ---
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  // --- Handler for generating the full case view ---
  if (message.action === "generateFullCaseView") {
    console.log("Content Script: Received request 'generateFullCaseView'.");
    const now = new Date();
    const generatedTime = now.toLocaleString('en-US', { dateStyle: 'short', timeStyle: 'medium' });
    generateCaseViewHtml(generatedTime)
      .then(fullHtml => {
        console.log("Content Script: Async generation complete, sending HTML to background.");
        if (fullHtml.length > 50 * 1024 * 1024) { // Example size check
           console.error("Generated HTML is too large to send directly via message.");
           // Consider sending an error message or handling large data differently
           chrome.runtime.sendMessage({ action: "openFullViewTab", htmlContent: "<html><body><h1>Error</h1><p>Generated content is too large to display.</p></body></html>" });
        } else {
           chrome.runtime.sendMessage({ action: "openFullViewTab", htmlContent: fullHtml });
        }
        // No need to call sendResponse here if background handles the action entirely
      })
      .catch(error => {
        console.error("Content Script: Error generating HTML:", error);
        // Optionally send an error status back to the popup if needed
        // sendResponse({ status: "error", message: "HTML generation failed." });
        // Send minimal error page to background?
         chrome.runtime.sendMessage({ action: "openFullViewTab", htmlContent: `<html><body><h1>Error</h1><p>Failed to generate case view: ${escapeHtml(error.message)}</p></body></html>` });
      });
    // Return true to indicate that sendResponse might be called asynchronously
    // (although in this case, the response pathway is primarily background->new tab)
    return true;
  }

  // --- Handler for getting Case Number and URL ---
  if (message.action === "getCaseNumberAndUrl") {
    console.log("Content Script: Received request 'getCaseNumberAndUrl'.");
    const caseNumber = findCaseNumberSpecific(); // Use the specific function
    const currentUrl = window.location.href;

    if (caseNumber) {
      console.log("Content Script: Sending back Case Number and URL:", caseNumber, currentUrl);
      sendResponse({ status: "success", caseNumber: caseNumber, url: currentUrl });
    } else {
      console.log("Content Script: Case Number not found.");
      sendResponse({ status: "error", message: "Could not find Case Number on page." });
    }
    // Return false because sendResponse is called synchronously within this block
    return false;
  }

  // --- Handler for background logging messages ---
  if (message.action === "logUrlProcessing") {
    // Log messages coming *from* the background script about its progress
    console.log(`[BACKGROUND] Processing ${message.itemType} ${message.index}/${message.total}: ${message.url}`);
    return false; // No response needed
  }

  // If the message action doesn't match any known handlers
  console.log("Content Script: Received unknown message action:", message.action);
  return false; // Indicate no response will be sent for unknown actions
});

console.log("Content Script: Message listener attached (or updated).");
