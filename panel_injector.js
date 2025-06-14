// panel_injector.js - Injects the UI, makes it draggable, and handles events.

console.log("Salesforce Full View: UI Panel Injector Loaded.");

// --- Helper: Wait for an element to appear in the DOM ---
/**
 * Waits for an element matching the selector to appear in the DOM.
 * @param {string} selector - CSS selector
 * @param {Element} [baseElement=document] - Base element
 * @param {number} [timeout=15000] - Timeout in ms
 * @returns {Promise<Element|null>}
 */
function waitForElement(selector, baseElement = document, timeout = 15000) {
    return new Promise((resolve) => {
        const startTime = Date.now();
        const interval = setInterval(() => {
            const element = baseElement.querySelector(selector);
            if (element) {
                clearInterval(interval);
                resolve(element);
            } else if (Date.now() - startTime > timeout) {
                console.warn(`VBSFU waitForElement: Timeout for "${selector}"`);
                clearInterval(interval);
                resolve(null);
            }
        }, 300);
    });
}

// --- Feature: Inject Custom Information Below the Header Row ---
/**
 * Finds key information on the page and injects it into a custom, visible container.
 * This is now triggered by a button click.
 */
async function injectCustomHeaderInfo() {
    console.log('VBSFU: "Show Key Info" button clicked. Attempting to inject info...');
    const statusDiv = document.getElementById('vbsfu-status');
    
    // Find the currently active Salesforce tab button to reliably get the panel ID.
    const activeTabButton = await waitForElement('li.slds-is-active[role="presentation"] a[role="tab"]');
    if (!activeTabButton) {
        console.error("VBSFU: Could not find the active Salesforce tab button.");
        if (statusDiv) {
            statusDiv.textContent = 'Error: Active tab not found.';
            statusDiv.style.color = 'var(--vbsfu-status-error)';
        }
        return;
    }

    const panelId = activeTabButton.getAttribute('aria-controls');
    if (!panelId) {
        console.error("VBSFU: Active tab button has no 'aria-controls' ID.");
        if (statusDiv) {
            statusDiv.textContent = 'Error: Could not identify tab content.';
            statusDiv.style.color = 'var(--vbsfu-status-error)';
        }
        return;
    }
    
    const activeTabPanel = document.getElementById(panelId);
    if (!activeTabPanel) {
        console.error(`VBSFU: Could not find tab panel with ID: ${panelId}`);
        if (statusDiv) {
            statusDiv.textContent = 'Error: Could not find tab content.';
            statusDiv.style.color = 'var(--vbsfu-status-error)';
        }
        return;
    }

    // 1. Find the anchor element to insert our container after.
    const anchorElement = await waitForElement('div.slds-grid.primaryFieldRow', activeTabPanel);
    if (!anchorElement) {
        console.log('VBSFU: Anchor element "div.slds-grid.primaryFieldRow" not found.');
        return;
    }

    // 2. Find or create our main container for custom info within the active tab.
    let infoContainer = activeTabPanel.querySelector('.vbsfu-custom-info-container');
    if (!infoContainer) {
        infoContainer = document.createElement('div');
        infoContainer.className = 'vbsfu-custom-info-container';
        infoContainer.style.cssText = `
            padding: 10px;
            margin: 10px 0;
            border-top: 1px solid #dddbda;
            border-bottom: 1px solid #dddbda;
        `;
        // Insert the container right after the anchor element.
        anchorElement.insertAdjacentElement('afterend', infoContainer);
    }
    
    // --- Inject Account Name ---
    // Clear previous content to ensure it's fresh, but only if we are re-injecting.
    infoContainer.innerHTML = ''; 
    
    const accountItemContainer = await waitForElement('records-record-layout-item[field-label="Account Name"]', activeTabPanel);
    if (accountItemContainer) {
        const accountNameElement = accountItemContainer.querySelector('force-lookup a');
        const accountName = accountNameElement?.textContent?.trim();

        if (accountName) {
            const accountDisplay = document.createElement('div');
            accountDisplay.className = 'vbsfu-account-header-display';
            accountDisplay.style.cssText = `
                font-size: 1.1em;
                font-weight: 600;
                color: #664d03;
                background-color: #fffbe6;
                padding: 8px 14px;
                border-radius: 6px;
                border-left: 5px solid #ffc107;
            `;
            accountDisplay.textContent = `Account: ${accountName}`;
            infoContainer.appendChild(accountDisplay); // Append to our container
            console.log(`VBSFU: Successfully displayed Account Name "${accountName}"`);
            if (statusDiv) statusDiv.textContent = 'Key info shown.';
        } else {
             console.log('VBSFU: Could not extract Account Name text.');
             if (statusDiv) statusDiv.textContent = 'Info not found.';
        }
    } else {
         console.log('VBSFU: Account Name field container not found on page.');
         if (statusDiv) statusDiv.textContent = 'Not a Case page?';
    }
}


// --- Draggable Panel Logic ---
function makePanelDraggable(panel, header) {
    let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
    let toggleButton = document.getElementById('vbsfu-toggle');

    header.onmousedown = dragMouseDown;

    function dragMouseDown(e) {
        e.preventDefault();
        pos3 = e.clientX;
        pos4 = e.clientY;
        document.onmouseup = closeDragElement;
        document.onmousemove = elementDrag;
        header.style.cursor = 'grabbing';
    }

    function elementDrag(e) {
        e.preventDefault();
        pos1 = pos3 - e.clientX;
        pos2 = pos4 - e.clientY;
        pos3 = e.clientX;
        pos4 = e.clientY;
        let newTop = panel.offsetTop - pos2;
        let newLeft = panel.offsetLeft - pos1;
        panel.style.top = newTop + "px";
        panel.style.left = newLeft + "px";
        if (toggleButton) {
           toggleButton.style.top = newTop + "px";
        }
    }

    function closeDragElement() {
        document.onmouseup = null;
        document.onmousemove = null;
        header.style.cursor = 'grab';
    }
}


// --- Main UI Injection ---
function injectUI() {
    if (document.getElementById('vbsfu-panel')) return; 
    console.log('VBSFU: Injecting main UI panel...');

    const panel = document.createElement('div');
    panel.id = 'vbsfu-panel';
    const header = document.createElement('div');
    header.id = 'vbsfu-header';
    const title = document.createElement('h4');
    title.textContent = 'VB SF Utils';
    header.appendChild(title);

    const content = document.createElement('div');
    content.id = 'vbsfu-content';

    const showInfoButton = document.createElement('button');
    showInfoButton.id = 'vbsfu-show-info';
    showInfoButton.textContent = 'Show Key Info';
    showInfoButton.className = 'vbsfu-button';
    
    const generateButton = document.createElement('button');
    generateButton.id = 'vbsfu-generate';
    generateButton.textContent = 'Generate Full View';
    generateButton.className = 'vbsfu-button';

    const copyButton = document.createElement('button');
    copyButton.id = 'vbsfu-copy';
    copyButton.textContent = 'Copy Record Link';
    copyButton.className = 'vbsfu-button';
    const aboutButton = document.createElement('button');
    aboutButton.id = 'vbsfu-about';
    aboutButton.textContent = 'About';
    aboutButton.className = 'vbsfu-button';
    const statusDiv = document.createElement('div');
    statusDiv.id = 'vbsfu-status';
    statusDiv.textContent = 'Ready.';
    
    // New button order
    content.appendChild(showInfoButton);
    content.appendChild(generateButton);
    content.appendChild(copyButton);
    content.appendChild(aboutButton);

    panel.appendChild(header);
    panel.appendChild(content);
    panel.appendChild(statusDiv);

    const toggleButton = document.createElement('button');
    toggleButton.id = 'vbsfu-toggle';
    toggleButton.innerHTML = '&#x1F6E0;&#xFE0F;';
    toggleButton.setAttribute('aria-label', 'Toggle Salesforce Utilities Panel');

    const modalOverlay = document.createElement('div');
    modalOverlay.id = 'vbsfu-modal-overlay';
    const modalContent = document.createElement('div');
    modalContent.id = 'vbsfu-modal-content';
    const modalClose = document.createElement('button');
    modalClose.id = 'vbsfu-modal-close';
    modalClose.innerHTML = '&times;';
    const modalTitle = document.createElement('h5');
    modalTitle.textContent = 'About VB Salesforce Utils';
    const modalBody = document.createElement('div');
    modalBody.id = 'vbsfu-modal-body';
    const extensionVersion = chrome.runtime.getManifest().version;
    modalBody.innerHTML = `<p><strong>Version:</strong> ${extensionVersion}</p><p>For information on this Chrome extension, contact Vincent Borghi.</p>`;
    modalContent.appendChild(modalClose);
    modalContent.appendChild(modalTitle);
    modalContent.appendChild(modalBody);
    modalOverlay.appendChild(modalContent);

    document.body.appendChild(panel);
    document.body.appendChild(toggleButton);
    document.body.appendChild(modalOverlay);

    makePanelDraggable(panel, header);

    // --- Event Listeners ---
    toggleButton.onclick = () => {
        panel.classList.toggle('vbsfu-visible');
        if (panel.classList.contains('vbsfu-visible')) {
            statusDiv.textContent = 'Ready.';
            statusDiv.style.color = 'var(--vbsfu-button-text)';
        }
    };
    aboutButton.onclick = () => modalOverlay.classList.add('vbsfu-visible');
    modalClose.onclick = () => modalOverlay.classList.remove('vbsfu-visible');
    modalOverlay.onclick = (e) => {
        if (e.target === modalOverlay) modalOverlay.classList.remove('vbsfu-visible');
    };
    
    // Add click listener for the new "Show Key Info" button
    showInfoButton.onclick = injectCustomHeaderInfo;

    generateButton.onclick = async () => {
        generateButton.disabled = true;
        copyButton.disabled = true;
        try {
            statusDiv.textContent = 'Preparing page...';
            statusDiv.style.color = 'var(--vbsfu-status-warn)';
            await waitForElement('a.slds-card__header-link[href*="/related/PSM_Notes__r/view"]', document, 10000).then(el => el.scrollIntoView({ block: 'center' }));
            await new Promise(r => setTimeout(r, 500));
            await waitForElement('div.forceListViewManager[aria-label*="Emails"]', document, 10000).then(el => el.scrollIntoView({ block: 'center' }));
            await new Promise(r => setTimeout(r, 500));
            window.scrollTo({ top: 0, behavior: 'auto' });
            statusDiv.textContent = 'Initiating...';
            chrome.runtime.sendMessage({ action: "initiateGenerateFullCaseView" }, response => {
                if (chrome.runtime.lastError) {
                    statusDiv.textContent = `Error: ${chrome.runtime.lastError.message}`;
                    statusDiv.style.color = 'var(--vbsfu-status-error)';
                } else {
                    statusDiv.textContent = "Processing initiated...";
                }
            });
        } catch (error) {
             statusDiv.textContent = 'Error preparing page!';
             statusDiv.style.color = 'var(--vbsfu-status-error)';
        } finally {
             generateButton.disabled = false;
             copyButton.disabled = false;
        }
    };

    copyButton.onclick = () => {
        statusDiv.textContent = '';
        statusDiv.style.color = 'var(--vbsfu-status-success)';
        const recordNumber = findCaseNumberSpecific(); // Note: This function is in content.js
        const currentUrl = window.location.href;
        let objectType = currentUrl.includes('/Case/') ? 'Case' : 'WorkOrder';

        if (recordNumber && currentUrl) {
            const linkText = `${objectType} ${recordNumber}`;
            const richTextHtml = `<a href="${currentUrl}">${linkText}</a>`;
            const blobHtml = new Blob([richTextHtml], { type: 'text/html' });
            const blobText = new Blob([linkText], { type: 'text/plain' });
            navigator.clipboard.write([new ClipboardItem({ 'text/html': blobHtml, 'text/plain': blobText })]).then(() => {
                statusDiv.textContent = `Copied: ${linkText}`;
            }).catch(err => {
                statusDiv.textContent = 'Error: Copy failed.';
                statusDiv.style.color = 'var(--vbsfu-status-error)';
            });
        } else {
            statusDiv.textContent = 'Error: Record # not found.';
            statusDiv.style.color = 'var(--vbsfu-status-error)';
        }
    };
}

// --- Initial Check and Injection Trigger ---
function init() {
    // We only need to inject the UI panel once. The logic is now entirely user-driven.
    // A slight delay is still good practice to avoid interrupting Salesforce's initial load.
    if (document.body) {
         setTimeout(injectUI, 1500);
    } else {
         document.addEventListener('DOMContentLoaded', () => {
            setTimeout(injectUI, 1500);
         });
    }
}

init();
