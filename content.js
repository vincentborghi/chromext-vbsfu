// content.js - Complete File (v54 - Context-Aware Scraping)

console.log("Salesforce Full View: Core Logic Loaded.");

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
 * Scrolls an element into view to trigger lazy-loading and waits briefly.
 * @param {string} selector - The CSS selector for the element to scroll to.
 * @param {Element} statusDiv - The UI element for status updates.
 * @returns {Promise<boolean>} - True if the element was found and scrolled to.
 */
async function scrollIntoViewAndWait(selector, statusDiv, baseElement = document) {
    console.log(`Ensuring visibility for: ${selector}`);
    if (statusDiv) {
         statusDiv.textContent = `Loading section...`;
    }
    const element = await waitForElement(selector, baseElement);
    if (element) {
        console.log(`Found element for ${selector}. Scrolling into view.`);
        element.scrollIntoView({ behavior: 'auto', block: 'center' });
        await new Promise(resolve => setTimeout(resolve, 750));
        console.log(`Finished ensuring visibility for: ${selector}`);
        return true;
    } else {
        console.warn(`Could not find element to scroll to for selector: ${selector}.`);
        return false;
    }
}


/**
 * Wraps chrome.runtime.sendMessage in a Promise for async/await usage.
 */
function sendMessagePromise(message) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(message, (response) => {
      if (chrome.runtime.lastError) {
        console.error("sendMessagePromise failed:", chrome.runtime.lastError.message);
        reject(chrome.runtime.lastError);
      } else {
        resolve(response);
      }
    });
  });
}

/**
 * Escapes HTML special characters to prevent XSS.
 */
function escapeHtml(unsafe) {
  if (typeof unsafe !== 'string') {
      return unsafe === null || typeof unsafe === 'undefined' ? '' : String(unsafe);
  }
  return unsafe
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

/**
 * Parses a date string (DD/MM/YYYY HH:MM) into a Date object (UTC).
 */
function parseDateStringFromTable(dateString) {
  if (!dateString) return null;
  let match = dateString.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);
  if (match) {
    try {
      const day = parseInt(match[1]);
      const month = parseInt(match[2]) - 1;
      const year = parseInt(match[3]);
      const hour = parseInt(match[4]);
      const minute = parseInt(match[5]);
      if (year > 1970 && month >= 0 && month < 12 && day >= 1 && day <= 31 && hour >= 0 && hour < 24 && minute >= 0 && minute < 60) {
        const dateObject = new Date(Date.UTC(year, month, day, hour, minute));
        if (!isNaN(dateObject.getTime())) {
            return dateObject;
        }
      }
    } catch (e) {
      console.error(`CS Error creating Date from matched parts: "${dateString}"`, e);
      return null;
    }
  }
  console.warn(`CS Could not parse date format: "${dateString}"`);
  return null;
}

/**
 * Uses MutationObserver to wait for a table's content to update after "View All".
 */
function waitForTableUpdate(targetNode, timeout = 10000) {
    console.log("Setting up MutationObserver to wait for table update...");
    return new Promise((resolve) => {
        if (!targetNode) {
            console.warn("waitForTableUpdate: Target node is null.");
            return resolve();
        }
        const timeoutTimer = setTimeout(() => {
            observer.disconnect();
            console.warn("MutationObserver timed out waiting for table update.");
            resolve();
        }, timeout);

        const observer = new MutationObserver((mutationsList, obs) => {
            for (const mutation of mutationsList) {
                if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
                    console.log("MutationObserver detected added nodes. Table has updated.");
                    clearTimeout(timeoutTimer);
                    obs.disconnect();
                    resolve();
                    return;
                }
            }
        });
        observer.observe(targetNode, { childList: true, subtree: true });
    });
}


// --- Functions to Extract Salesforce Record Details (NOW CONTEXT-AWARE) ---
function findSubjectInContainer(container) {
    if (!container) return 'N/A';
    const element = container.querySelector('support-output-case-subject-field lightning-formatted-text');
    return element ? element.textContent?.trim() : 'N/A (Subject)';
}

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

function findCaseNumberSpecific(baseElement) {
    console.log("Content Script: Attempting to find Case/Record Number (Specific Selector)...");
    const itemSelector = 'records-highlights-details-item:has(p[title="Case Number"])';
    const textSelector = 'lightning-formatted-text';
    const detailsItem = baseElement.querySelector(itemSelector);

    if (detailsItem) {
        const textElement = detailsItem.querySelector(textSelector);
        if (textElement) {
            const recordNum = textElement.textContent?.trim();
            if (recordNum && /^\d+$/.test(recordNum)) {
                 console.log("Content Script: Found Record Number:", recordNum);
                 return recordNum;
            }
        }
    }
    console.warn("Content Script: Record Number details item not found. Falling back...");
    const container = baseElement.querySelector('records-highlights2');
    const fallbackRecordNum = findCaseNumberInContainer(container);
    if (fallbackRecordNum && !fallbackRecordNum.startsWith('N/A')) {
        return fallbackRecordNum;
    }
    return null;
}

function findStatusInContainer(container) {
    if (!container) return 'N/A';
    const statusItem = container.querySelector('records-highlights-details-item:has(records-formula-output lightning-formatted-rich-text)');
    const element = statusItem?.querySelector('lightning-formatted-rich-text span[part="formatted-rich-text"]');
    return element ? element.textContent?.trim() : 'N/A (Status)';
}

async function findCreatorName(baseElement) {
    const createdByItem = await waitForElement('records-record-layout-item[field-label="Created By"]', baseElement);
    if (!createdByItem) { return 'N/A (Creator)'; }
    const nameElement = createdByItem.querySelector('force-lookup a');
    return nameElement ? nameElement.textContent?.trim() : 'N/A (Creator)';
}

async function findCreatedDate(baseElement) {
    const createdByItem = await waitForElement('records-record-layout-item[field-label="Created By"]', baseElement);
    if (!createdByItem) { return 'N/A (Created Date)'; }
    const dateElement = createdByItem.querySelector('records-modstamp lightning-formatted-text');
    return dateElement ? dateElement.textContent?.trim() : 'N/A (Created Date)';
}

function findOwnerInContainer(container) {
     if (!container) return 'N/A';
     const ownerItem = container.querySelector('records-highlights-details-item:has(force-lookup)');
     const element = ownerItem?.querySelector('force-lookup a');
     return element ? element.textContent?.trim() : 'N/A (Owner)';
}

async function findAccountName(baseElement) {
    const accountItem = await waitForElement('records-record-layout-item[field-label="Account Name"]', baseElement);
    if (!accountItem) {
        return 'N/A (Account)';
    }
    const nameElement = accountItem.querySelector('force-lookup a');
    return nameElement ? nameElement.textContent?.trim() : 'N/A (Account)';
}

async function findCaseDescription(baseElement) {
     const descriptionContainer = await waitForElement('article.cPSM_Case_Description', baseElement);
     if (!descriptionContainer) { return ''; }
     let textElement = descriptionContainer.querySelector('lightning-formatted-text.txtAreaReadOnly') || descriptionContainer.querySelector('lightning-formatted-text');
     if (!textElement) { return ''; }
     const viewMoreButton = descriptionContainer.querySelector('button.slds-button:not([disabled])');
     let descriptionHTML = '';
     if (viewMoreButton && (viewMoreButton.textContent.includes('View More') || viewMoreButton.textContent.includes('Show More'))) {
        viewMoreButton.click();
        await new Promise(resolve => setTimeout(resolve, 500));
        let updatedTextElement = descriptionContainer.querySelector('lightning-formatted-text.txtAreaReadOnly') || descriptionContainer.querySelector('lightning-formatted-text');
        descriptionHTML = updatedTextElement?.innerHTML?.trim() || '';
     } else {
        descriptionHTML = textElement?.innerHTML?.trim() || '';
     }
     return descriptionHTML;
}

// --- Function to Extract Note URLs from Table and Trigger Fetch (CONTEXT-AWARE) ---
async function extractAndFetchNotes(baseElement) {
    const notesHeaderSelector = 'a.slds-card__header-link[href*="/related/PSM_Notes__r/view"]';
    const headerLinkElement = await waitForElement(notesHeaderSelector, baseElement);
    if (!headerLinkElement) { console.warn("Notes header not found. Skipping notes."); return []; }

    const listManager = headerLinkElement.closest('lst-related-list-view-manager');
    if (!listManager) { console.warn("Notes list manager not found. Skipping notes."); return []; }

    const countSpan = headerLinkElement.querySelector('span[title*="("]');
    if (countSpan && countSpan.textContent?.trim() === '(0)') { console.log("Notes count is 0, skipping."); return []; }

    const viewAllLink = listManager.querySelector('a.slds-card__footer[href*="/related/PSM_Notes__r/view"]');
    if (viewAllLink && viewAllLink.offsetParent !== null) {
        console.log("Notes: 'View All' link found. Clicking...");
        viewAllLink.click();
        try {
            const tableBodyToWatch = await waitForElement('tbody[data-rowgroup-body]', listManager, 5000);
            await waitForTableUpdate(tableBodyToWatch, 15000);
        } catch (error) {
            console.error("Error waiting for Notes table to update:", error);
        }
    }

    const dataTable = await waitForElement('lightning-datatable', listManager, 10000);
    if (!dataTable) { console.warn("Notes datatable not found. Skipping notes."); return []; }

    const tableBody = await waitForElement('tbody[data-rowgroup-body]', dataTable, 5000);
    if (!tableBody) { console.warn("Notes table body not found. Skipping notes."); return []; }

    const rows = tableBody.querySelectorAll('tr[data-row-key-value]');
    if (rows.length === 0) return [];

    const notesToFetch = Array.from(rows).map(row => {
        const noteLink = row.querySelector('th[data-label="PSM Note Name"] a');
        const dateSpan = row.querySelector('td[data-label="Created Date"] lst-formatted-text span');
        const relativeUrl = noteLink?.getAttribute('href');
        const noteDateStr = dateSpan?.title || dateSpan?.textContent?.trim();
        if (!relativeUrl || !noteDateStr) return null;
        return {
            type: 'Note',
            url: new URL(relativeUrl, window.location.origin).href,
            title: noteLink?.textContent?.trim() || 'N/A',
            author: row.querySelector('td[data-label="Created By"] a')?.textContent?.trim() || 'N/A',
            dateStr: noteDateStr,
            descriptionSnippet: row.querySelector('td[data-label="Description"] lightning-base-formatted-text')?.textContent?.trim() || ''
        };
    }).filter(Boolean);

    if (notesToFetch.length === 0) return [];

    console.log(`Sending ${notesToFetch.length} notes to background for fetching...`);
    const response = await sendMessagePromise({ action: "fetchItemDetails", items: notesToFetch });

    if (response?.status === "success" && response.details) {
        return notesToFetch.map(noteInfo => {
            const fetched = response.details[noteInfo.url] || {};
            const finalDateObject = fetched.dateObject ? new Date(fetched.dateObject) : parseDateStringFromTable(noteInfo.dateStr);
            let finalDescription = fetched.description;
            if (finalDescription && finalDescription.startsWith('Error:')) {
                finalDescription = `[Fetch Error]`;
            } else if (!finalDescription) {
                finalDescription = '[Description Empty or Not Fetched]';
            }
            return {
                ...noteInfo, dateObject: finalDateObject, content: finalDescription,
                isPublic: fetched.isPublic ?? null, attachments: 'N/A'
            };
        }).filter(note => note.dateObject);
    }
    console.error("Failed to fetch note details from background:", response?.message);
    return [];
}


// --- Function to Extract Email Info from Table and Trigger Fetch (CONTEXT-AWARE) ---
async function extractAndFetchEmails(baseElement) {
    const emailsListContainerSelector = 'div.forceListViewManager[aria-label*="Emails"]';
    const emailsListContainer = await waitForElement(emailsListContainerSelector, baseElement);
    if (!emailsListContainer) { console.warn("Emails container not found. Skipping emails."); return []; }

    const countSpan = emailsListContainer.querySelector('a.slds-card__header-link span[title*="("]');
    if (countSpan && countSpan.textContent?.trim() === '(0)') { console.log("Emails count is 0, skipping."); return []; }

    const parentCard = emailsListContainer.closest('.forceRelatedListSingleContainer');
    const viewAllLink = parentCard?.querySelector('a.slds-card__footer[href*="/related/EmailMessages/view"]');
    if (viewAllLink && viewAllLink.offsetParent !== null) {
        console.log("Emails: 'View All' link found. Clicking...");
        viewAllLink.click();
        try {
            const tableBodyToWatch = await waitForElement('table.uiVirtualDataTable tbody', emailsListContainer, 5000);
            await waitForTableUpdate(tableBodyToWatch, 15000);
        } catch (error) {
            console.error("Error waiting for Emails table to update:", error);
        }
    }

    const dataTable = await waitForElement('table.uiVirtualDataTable', emailsListContainer, 10000);
    if (!dataTable) { console.warn("Emails table not found. Skipping emails."); return []; }

    const tableBody = await waitForElement('tbody', dataTable, 5000);
    if (!tableBody) { console.warn("Emails table body not found. Skipping emails."); return []; }

    const rows = tableBody.querySelectorAll('tr');
    if (rows.length === 0) return [];

    const emailsToFetch = Array.from(rows).map(row => {
        const cells = row.querySelectorAll('th, td');
        if (cells.length < 5) return null;
        const subjectLink = cells[1]?.querySelector('a.outputLookupLink');
        const dateEl = cells[4]?.querySelector('span.uiOutputDateTime');
        const relativeUrl = subjectLink?.getAttribute('href');
        const emailDateStr = dateEl?.textContent?.trim();
        if (!relativeUrl || !emailDateStr) return null;
        return {
            type: 'Email',
            url: new URL(relativeUrl, window.location.origin).href,
            title: subjectLink?.textContent?.trim(),
            author: cells[2]?.querySelector('a.emailuiFormattedEmail')?.textContent?.trim() || 'N/A',
            to: cells[3]?.querySelector('span.uiOutputText')?.title || cells[3]?.querySelector('span.uiOutputText')?.textContent?.trim() || 'N/A',
            dateStr: emailDateStr
        };
    }).filter(Boolean);

    if (emailsToFetch.length === 0) return [];

    console.log(`Sending ${emailsToFetch.length} emails to background for fetching...`);
    const response = await sendMessagePromise({ action: "fetchItemDetails", items: emailsToFetch });

    if (response?.status === "success" && response.details) {
        return emailsToFetch.map(emailInfo => {
            const fetched = response.details[emailInfo.url] || {};
            const finalDateObject = fetched.dateObject ? new Date(fetched.dateObject) : parseDateStringFromTable(emailInfo.dateStr);
            let finalContent = fetched.description;
            if (finalContent && finalContent.startsWith('Error:')) {
                finalContent = `Subject: ${escapeHtml(fetched.subject || emailInfo.title)}<br>[Body Fetch Error]`;
            } else if (!finalContent) {
                 finalContent = `Subject: ${escapeHtml(fetched.subject || emailInfo.title)}<br>[Body Not Fetched or Empty]`;
            }
            return {
                ...emailInfo, title: fetched.subject || emailInfo.title,
                author: fetched.from || emailInfo.author, to: fetched.to || emailInfo.to,
                dateObject: finalDateObject, content: finalContent,
                isPublic: null, attachments: 'N/A'
            };
        }).filter(email => email.dateObject);
    }
    console.error("Failed to fetch email details from background:", response?.message);
    return [];
}


// --- Main Function to Orchestrate Extraction and Generate HTML (CONTEXT-AWARE) ---
async function generateCaseViewHtml(generatedTime) {
    // 1. Find the currently active tab button to get the context.
    const activeTabButton = await waitForElement('li.slds-is-active[role="presentation"] a[role="tab"]');
    if (!activeTabButton) {
        console.error("FATAL: Could not find the active Salesforce tab button.");
        return `<html><head><title>Error</title></head><body><h1>Extraction Error</h1><p>Could not find the active Salesforce tab to generate the view from. Is a Case tab currently selected?</p></body></html>`;
    }
    
    // 2. Get the ID of the content panel from the button.
    const panelId = activeTabButton.getAttribute('aria-controls');
    if (!panelId) {
        console.error("FATAL: Active tab button has no 'aria-controls' ID.");
        return `<html><head><title>Error</title></head><body><h1>Extraction Error</h1><p>Could not identify the content panel for the active tab.</p></body></html>`;
    }

    // 3. Get the content panel element itself. This is our base element for all searches.
    const activeTabPanel = document.getElementById(panelId);
    if (!activeTabPanel) {
        console.error(`FATAL: Could not find tab panel with ID: ${panelId}`);
        return `<html><head><title>Error</title></head><body><h1>Extraction Error</h1><p>Could not find the content for the active tab (ID: ${panelId}).</p></body></html>`;
    }

    console.log("Starting async HTML generation for the active tab...");
    let objectType = 'Record';
    const currentHref = window.location.href;
    if (currentHref.includes('/Case/')) objectType = 'Case';
    else if (currentHref.includes('/WorkOrder/')) objectType = 'WorkOrder';
    console.log(`Detected Object Type: ${objectType}`);

    const highlightsContainerElement = await waitForElement('records-highlights2', activeTabPanel);
    if (!highlightsContainerElement) {
      console.error(`FATAL: Highlights container not found in the active tab.`);
      return `<html><head><title>Error</title></head><body><h1>Extraction Error</h1><p>Could not find the main details section on the page.</p></body></html>`;
    }

    console.log("Extracting header details and related lists from the active tab...");
    const extractionPromises = [
        Promise.resolve(findSubjectInContainer(highlightsContainerElement)),
        Promise.resolve(findCaseNumberSpecific(activeTabPanel)),
        Promise.resolve(findStatusInContainer(highlightsContainerElement)),
        Promise.resolve(findOwnerInContainer(highlightsContainerElement)),
        findCreatorName(activeTabPanel),
        findAccountName(activeTabPanel),
        findCreatedDate(activeTabPanel),
        findCaseDescription(activeTabPanel),
        extractAndFetchNotes(activeTabPanel),
        extractAndFetchEmails(activeTabPanel)
    ];

    const [
        subject, recordNumber, status, owner, creatorName, accountName, createdDateStr, description,
        notesData, emailsData
    ] = await Promise.all(extractionPromises);

    console.log(`Extraction results: Subject=${!!subject}, ${objectType}#=${recordNumber}, Notes=${notesData?.length || 0}, Emails=${emailsData?.length || 0}`);

    let allTimelineItems = [];
    if (Array.isArray(notesData)) { allTimelineItems = allTimelineItems.concat(notesData); }
    if (Array.isArray(emailsData)) { allTimelineItems = allTimelineItems.concat(emailsData); }

    console.log(`Sorting ${allTimelineItems.length} combined timeline items...`);
    allTimelineItems.sort((a, b) => {
        const timeA = a.dateObject?.getTime() || 0;
        const timeB = b.dateObject?.getTime() || 0;
        if (!timeA && !timeB) return 0;
        if (!timeA) return 1;
        if (!timeB) return -1;
        return timeA - timeB;
    });

    console.log("Constructing final HTML output...");
    const safeRecordNumber = escapeHtml(recordNumber || 'N/A');
    const safeSubject = escapeHtml(subject || 'N/A');
    const safeObjectType = escapeHtml(objectType);
    const safeAccountName = escapeHtml(accountName || 'N/A');

    let htmlOutput = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${safeObjectType} ${safeRecordNumber}: ${safeSubject}</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; line-height: 1.4; padding: 15px 25px; margin: 0; color: #333; background-color: #f9f9f9; }
        h1, h2 { border-bottom: 1px solid #ccc; padding-bottom: 6px; margin-top: 25px; margin-bottom: 15px; color: #1a5f90; font-weight: 600; }
        h1 { font-size: 1.7em; } h2 { font-size: 1.4em; }
        .meta-info-bar { display: flex; justify-content: space-between; align-items: center; padding: 8px 12px; margin-bottom: 25px; border-radius: 5px; background-color: #eef3f8; border: 1px solid #d1e0ee; }
        .customer-account-info { font-size: 1.1em; font-weight: 600; color: #005a9e; }
        .generation-info { font-size: 0.85em; color: #555; }
        .record-details { background-color: #fff; border: 1px solid #e1e5eb; padding: 15px 20px; border-radius: 5px; margin-bottom: 25px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .record-details h2 { margin-top: 0; margin-bottom: 12px; }
        .details-grid { display: grid; grid-template-columns: auto 1fr; gap: 4px 10px; margin-bottom: 15px; align-items: start; }
        .details-grid dt { grid-column: 1; font-weight: 600; color: #005fb2; text-align: right; padding-right: 8px; white-space: nowrap; }
        .details-grid dd { grid-column: 2; margin-left: 0; word-wrap: break-word; text-align: left; }
        .description-label { font-weight: 600; color: #005fb2; margin-bottom: 5px; display: block; }
        .record-details .description-content { white-space: pre-wrap; word-wrap: break-word; margin-top: 0px; padding: 10px 12px; background-color: #f1f1f1; border-radius: 4px; font-size: 0.95em; max-height: 400px; overflow-y: auto; border: 1px solid #e0e0e0; }
        .timeline-item { border: 1px solid #e1e5eb; padding: 12px 18px; margin-bottom: 10px; border-radius: 5px; background-color: #fff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); position: relative; }
        .timeline-item.type-note { border-left: 5px solid #6b92dc; }
        .timeline-item.type-email { border-left: 5px solid #5cb85c; }
        .item-header { font-size: 0.95em; color: #444; margin-bottom: 8px; border-bottom: 1px dashed #eee; padding-bottom: 6px; line-height: 1.4; }
        .item-timestamp { color: #555; font-family: monospace; margin-right: 10px; font-weight: bold; font-size: 1.2em; background-color:#f0f0f0; padding: 1px 4px; border-radius: 3px; }
        .item-type-label { font-weight: bold; text-transform: uppercase; font-size: 0.85em; margin-right: 5px; }
        .item-type-label.type-note { color: #6b92dc; }
        .item-type-label.type-email { color: #5cb85c; }
        .item-subject-title { font-weight: 600; color: #222; margin-left: 4px; font-size: 1.05em; }
        .item-meta { display: block; font-size: 0.85em; color: #666; margin-top: 3px; }
        .item-meta-label { color: #005fb2; font-weight: 600; }
        .item-meta-info { color: #555; margin-left: 3px; }
        .item-content { white-space: normal; word-wrap: break-word; overflow-wrap: break-word; color: #333; margin-top: 10px; font-size: 0.95em; line-height: 1.45; }
        .item-content p, .item-content ul, .item-content ol { margin-top: 0; margin-bottom: 0.5em; }
        .item-content a { color: #007bff; text-decoration: underline; }
        .item-content blockquote { border-left: 3px solid #ccc; padding-left: 10px; margin-left: 5px; color: #666; font-style: italic; }
        .item-content pre, .item-content code { font-family: monospace; background-color: #eee; padding: 1px 3px; border-radius: 2px; white-space: pre-wrap; }
        .item-attachments { font-style: italic; color: #888; font-size: 0.85em; margin-top: 10px; }
        .error-message { color: red; font-weight: bold; background-color: #ffebeb; border: 1px solid red; padding: 5px 8px; border-radius: 3px; display: inline-block; margin-top: 5px; }
        .item-visibility { margin-left: 8px; font-size: 0.9em; font-weight: bold; text-transform: lowercase; padding: 1px 5px; border-radius: 3px; border: 1px solid transparent; }
        .item-visibility.public { color: #8e1b03; background-color: #fdd; border-color: #fbb; }
        .item-visibility.internal { color: #333; background-color: #eee; border-color: #ddd; }
    </style>
</head>
<body>
    <h1>${safeObjectType} ${safeRecordNumber}: ${safeSubject}</h1>
    
    <div class="meta-info-bar">
        <div class="customer-account-info"><strong>Customer Account:</strong> ${safeAccountName}</div>
        <div class="generation-info">Generated: ${escapeHtml(generatedTime)}</div>
    </div>

    <div class="record-details">
        <h2>Details</h2>
        <dl class="details-grid">
            <dt>Date Created:</dt><dd>${escapeHtml(createdDateStr || 'N/A')}</dd>
            <dt>Created By:</dt><dd>${escapeHtml(creatorName || 'N/A')}</dd>
            <dt>Status:</dt><dd>${escapeHtml(status || 'N/A')}</dd>
            <dt>Owner:</dt><dd>${escapeHtml(owner || 'N/A')}</dd>
        </dl>
        <div class="description-label">Description:</div>
        <div class="description-content">${description || '<p><i>Description empty or not found.</i></p>'}</div>
    </div>
    <h2>Timeline (${allTimelineItems.length} items)</h2>
`;

    if (allTimelineItems.length === 0) {
        htmlOutput += "<p>No Notes or Emails found or extracted successfully.</p>";
    } else {
        allTimelineItems.forEach(item => {
            let contentHtml = '';
            if (item.content && (item.content.startsWith('Error:') || item.content.startsWith('[Fetch Error') || item.content.startsWith('[Body Fetch Error') || item.content.startsWith('[Content'))) {
               contentHtml = `<span class="error-message">${escapeHtml(item.content)}</span>`;
            } else {
               contentHtml = item.content || '<i>[Content Missing]</i>';
            }

            let visibilityLabel = '';
            if (item.type === 'Note') {
                if (item.isPublic === true) visibilityLabel = `<span class="item-visibility public">(public)</span>`;
                else if (item.isPublic === false) visibilityLabel = `<span class="item-visibility internal">(internal)</span>`;
            }

            let formattedTimestamp = 'N/A';
            if (item.dateObject && !isNaN(item.dateObject.getTime())) {
                formattedTimestamp = item.dateObject.toLocaleString(undefined, {
                    year: 'numeric', month: 'numeric', day: 'numeric',
                    hour: '2-digit', minute: '2-digit', hour12: false
                });
            } else {
                formattedTimestamp = escapeHtml(item.dateStr || 'Date Error');
            }

            const itemTypeClass = `type-${escapeHtml(item.type?.toLowerCase() || 'unknown')}`;
            const itemTypeLabel = escapeHtml(item.type || 'Item');
            const itemTitle = escapeHtml(item.title || 'N/A');
            const itemAuthor = escapeHtml(item.author || 'N/A');
            const itemTo = escapeHtml(item.to || 'N/A');

            let headerMetaDetails = (item.type === 'Email')
                ? `<span class="item-meta"><span class="item-meta-label">From:</span> <span class="item-meta-info">${itemAuthor}</span> | <span class="item-meta-label">To:</span> <span class="item-meta-info">${itemTo}</span></span>`
                : `<span class="item-meta"><span class="item-meta-label">By:</span> <span class="item-meta-info"><strong>${itemAuthor}</strong></span></span>`;

            htmlOutput += `
            <div class="timeline-item ${itemTypeClass}">
                <div class="item-header">
                    <strong class="item-type-label ${itemTypeClass}">${itemTypeLabel}</strong>
                    ${visibilityLabel} <span class="item-timestamp">[${formattedTimestamp}]</span> -
                    <span class="item-subject-title">${itemTitle}</span>
                    ${headerMetaDetails}
                </div>
                <div class="item-content">${contentHtml}</div>
                <div class="item-attachments">Attachments: ${escapeHtml(item.attachments || 'N/A')}</div>
            </div>`;
        });
    }

    htmlOutput += `</body></html>`;
    console.log("HTML generation complete.");
    return htmlOutput;
}


// --- Message Listener ---
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    console.log(`Content Script: Received message action="${message.action}"`);

    if (message.action === "generateFullCaseView") {
        console.log("Content Script: Handling 'generateFullCaseView' command.");
        const now = new Date();
        const generatedTime = now.toLocaleString('en-US', { dateStyle: 'short', timeStyle: 'medium' });
        const statusDiv = document.getElementById('vbsfu-status');
        if (statusDiv) {
            statusDiv.textContent = 'Extracting data...';
            statusDiv.style.color = 'orange';
        }

        generateCaseViewHtml(generatedTime)
          .then(fullHtml => {
            console.log("Content Script: HTML generation complete. Sending to background.");
            if (statusDiv) {
                statusDiv.textContent = 'Opening results...';
                statusDiv.style.color = 'green';
            }
            chrome.runtime.sendMessage({ action: "openFullViewTab", htmlContent: fullHtml });
          })
          .catch(error => {
            console.error("Content Script: Error during HTML generation:", error);
            if (statusDiv) {
                statusDiv.textContent = 'Error generating view!';
                statusDiv.style.color = 'red';
            }
            chrome.runtime.sendMessage({
                action: "openFullViewTab",
                htmlContent: `<html><body><h1>Error Generating View</h1><pre>${escapeHtml(error.message)}</pre></body></html>`
            });
          });

        return true; // Indicate async response
    }

    if (message.action === "getCaseNumberAndUrl") {
        // Use the active tab as the context for finding the case number
        const activeTabButton = document.querySelector('li.slds-is-active[role="presentation"] a[role="tab"]');
        if (!activeTabButton) {
             sendResponse({ status: "error", message: "Could not find active tab button." });
             return false;
        }
        const panelId = activeTabButton.getAttribute('aria-controls');
        const activeTabPanel = panelId ? document.getElementById(panelId) : null;
        
        const recordNumber = activeTabPanel ? findCaseNumberSpecific(activeTabPanel) : null;
        const currentUrl = window.location.href;
        if (recordNumber) {
          sendResponse({ status: "success", caseNumber: recordNumber, url: currentUrl });
        } else {
          sendResponse({ status: "error", message: "Could not find Record Number in active tab." });
        }
        return false;
    }

    if (message.action === "logUrlProcessing") {
        console.log(`[BACKGROUND] Progress: Fetching ${message.itemType} ${message.index}/${message.total}`);
        const statusDiv = document.getElementById('vbsfu-status');
        if (statusDiv) {
            statusDiv.textContent = `Fetching ${message.itemType} ${message.index}/${message.total}...`;
            statusDiv.style.color = 'orange';
        }
        return false;
    }

    return false;
});
