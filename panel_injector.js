// panel_injector.js - Injects the UI, makes it draggable, and handles events.

console.log("Salesforce Full View: UI Panel Injector Loaded.");

// --- Draggable Panel Logic ---
function makePanelDraggable(panel, header) {
    let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
    let toggleButton = document.getElementById('vbsfu-toggle');

    header.onmousedown = dragMouseDown;

    function dragMouseDown(e) {
        e.preventDefault();
        // Get the initial mouse cursor position
        pos3 = e.clientX;
        pos4 = e.clientY;
        document.onmouseup = closeDragElement;
        // Call a function whenever the cursor moves
        document.onmousemove = elementDrag;
        header.style.cursor = 'grabbing';
    }

    function elementDrag(e) {
        e.preventDefault();
        // Calculate the new cursor position
        pos1 = pos3 - e.clientX;
        pos2 = pos4 - e.clientY;
        pos3 = e.clientX;
        pos4 = e.clientY;
        // Set the element's new position
        let newTop = panel.offsetTop - pos2;
        let newLeft = panel.offsetLeft - pos1;

        panel.style.top = newTop + "px";
        panel.style.left = newLeft + "px";
        
        // Also move the toggle button if it exists
        if (toggleButton) {
           toggleButton.style.top = newTop + "px";
           // The toggle button's right position is relative to the viewport, not the panel's left.
           // No change to its 'right' style is needed during drag.
        }
    }

    function closeDragElement() {
        // Stop moving when mouse button is released
        document.onmouseup = null;
        document.onmousemove = null;
        header.style.cursor = 'grab';
    }
}


// --- Main UI Injection ---
function injectUI() {
    console.log('Injecting UI...');
    if (document.getElementById('vbsfu-panel')) return;

    // --- Create Panel and Header ---
    const panel = document.createElement('div');
    panel.id = 'vbsfu-panel';

    const header = document.createElement('div');
    header.id = 'vbsfu-header';
    const title = document.createElement('h4');
    title.textContent = 'SalesForce Helper for PSM';
    header.appendChild(title);

    // --- Create Buttons and Content Area ---
    const content = document.createElement('div');
    content.id = 'vbsfu-content';

    const generateButton = document.createElement('button');
    generateButton.id = 'vbsfu-generate';
    generateButton.textContent = 'Generate Full View';
    generateButton.className = 'vbsfu-button';

    const copyButton = document.createElement('button');
    copyButton.id = 'vbsfu-copy';
    copyButton.textContent = 'Copy Link';
    copyButton.className = 'vbsfu-button';
    
    const aboutButton = document.createElement('button');
    aboutButton.id = 'vbsfu-about';
    aboutButton.textContent = 'About';
    aboutButton.className = 'vbsfu-button';

    const statusDiv = document.createElement('div');
    statusDiv.id = 'vbsfu-status';
    statusDiv.textContent = 'Ready.';

    content.appendChild(generateButton);
    content.appendChild(copyButton);
    content.appendChild(aboutButton);

    panel.appendChild(header);
    panel.appendChild(content);
    panel.appendChild(statusDiv);

    // --- Create Toggle Button ---
    const toggleButton = document.createElement('button');
    toggleButton.id = 'vbsfu-toggle';
    toggleButton.innerHTML = '&#x1F6E0;&#xFE0F;'; // Hammer and wrench emoji
    toggleButton.setAttribute('aria-label', 'Toggle Helper Panel');

    // --- Create About Modal ---
    const modalOverlay = document.createElement('div');
    modalOverlay.id = 'vbsfu-modal-overlay';
    const modalContent = document.createElement('div');
    modalContent.id = 'vbsfu-modal-content';
    const modalClose = document.createElement('button');
    modalClose.id = 'vbsfu-modal-close';
    modalClose.innerHTML = '&times;';
    
    const modalTitle = document.createElement('h5');
    modalTitle.textContent = 'About Salesforce for PSM Helper';
    
    const modalBody = document.createElement('div');
    modalBody.id = 'vbsfu-modal-body';
    
    // -- Customieable HTML content for the modal --
    const extensionVersion = chrome.runtime.getManifest().version;
    modalBody.innerHTML = `
        <p><strong>Version:</strong> ${extensionVersion} (June 2025)</p>
        <p>This Chrome extension, "Salesforce for PSM Helper", is an experimental tool 
        to help you using PSM SalesForce.</p>
        <p>For feedback or information, contact Vincent Borghi (by email preferably).</p>
    `;
    
    modalContent.appendChild(modalClose);
    modalContent.appendChild(modalTitle);
    modalContent.appendChild(modalBody);
    modalOverlay.appendChild(modalContent);

    // --- Append everything to the body ---
    document.body.appendChild(panel);
    document.body.appendChild(toggleButton);
    document.body.appendChild(modalOverlay);

    // --- Make the panel draggable ---
    makePanelDraggable(panel, header);

    // --- Event Listeners ---
    toggleButton.onclick = () => {
        panel.classList.toggle('vbsfu-visible');
        if (panel.classList.contains('vbsfu-visible')) {
            statusDiv.textContent = 'Ready.';
            statusDiv.style.color = 'var(--vbsfu-button-text)';
        }
    };
    
    aboutButton.onclick = () => {
        modalOverlay.classList.add('vbsfu-visible');
    };

    modalClose.onclick = () => {
        modalOverlay.classList.remove('vbsfu-visible');
    };
    
    modalOverlay.onclick = (e) => {
        if (e.target === modalOverlay) { // Only close if clicking the overlay itself
            modalOverlay.classList.remove('vbsfu-visible');
        }
    };


    // This handler relies on functions from content.js (scrollIntoViewAndWait)
    generateButton.onclick = async () => {
        console.log('Panel Generate button clicked');
        generateButton.disabled = true;
        copyButton.disabled = true;

        try {
            statusDiv.textContent = 'Preparing page...';
            statusDiv.style.color = 'var(--vbsfu-status-warn)';

            // Scroll to wake up lazy-loaded related lists
            await scrollIntoViewAndWait('a.slds-card__header-link[href*="/related/PSM_Notes__r/view"]', statusDiv);
            await scrollIntoViewAndWait('div.forceListViewManager[aria-label*="Emails"]', statusDiv);

            window.scrollTo({ top: 0, behavior: 'auto' });
            console.log('Finished preparing page. Initiating generation.');

            statusDiv.textContent = 'Initiating...';
            statusDiv.style.color = 'var(--vbsfu-status-warn)';
            chrome.runtime.sendMessage({ action: "initiateGenerateFullCaseView" }, (response) => {
                if (chrome.runtime.lastError) {
                    console.error("Error sending initiate message:", chrome.runtime.lastError.message);
                    statusDiv.textContent = `Error: ${chrome.runtime.lastError.message}`;
                    statusDiv.style.color = 'var(--vbsfu-status-error)';
                } else {
                    console.log("Background acknowledged initiation.");
                    statusDiv.textContent = "Processing initiated...";
                    statusDiv.style.color = 'var(--vbsfu-status-success)';
                }
            });
        } catch (error) {
             console.error("Error during pre-generation scroll:", error);
             statusDiv.textContent = 'Error preparing page!';
             statusDiv.style.color = 'var(--vbsfu-status-error)';
        } finally {
             generateButton.disabled = false;
             copyButton.disabled = false;
        }
    };

    // This handler relies on functions from content.js (findCaseNumberSpecific)
    copyButton.onclick = () => {
        console.log('Panel Copy button clicked');
        statusDiv.textContent = '';
        statusDiv.style.color = 'var(--vbsfu-status-success)';

        const recordNumber = findCaseNumberSpecific();
        const currentUrl = window.location.href;
        let objectType = 'Record';
        if (currentUrl.includes('/Case/')) objectType = 'Case';
        else if (currentUrl.includes('/WorkOrder/')) objectType = 'WorkOrder';

        if (recordNumber && currentUrl) {
            const linkText = `${objectType} ${recordNumber}`;
            const richTextHtml = `<a href="${currentUrl}">${linkText}</a>`;
            try {
                const blobHtml = new Blob([richTextHtml], { type: 'text/html' });
                const blobText = new Blob([linkText], { type: 'text/plain' });
                const clipboardItem = new ClipboardItem({ 'text/html': blobHtml, 'text/plain': blobText });

                navigator.clipboard.write([clipboardItem]).then(() => {
                    console.log('Rich text link copied!');
                    statusDiv.textContent = `Copied: ${linkText}`;
                }).catch(err => {
                    console.error('Failed to copy rich text: ', err);
                    statusDiv.textContent = 'Error: Copy failed.';
                    statusDiv.style.color = 'var(--vbsfu-status-error)';
                });
            } catch (error) {
                console.error('Clipboard API error:', error);
                statusDiv.textContent = 'Error: Clipboard API failed.';
                statusDiv.style.color = 'var(--vbsfu-status-error)';
            }
        } else {
            console.error("Failed to get record number for copy.");
            statusDiv.textContent = 'Error: Record # not found.';
            statusDiv.style.color = 'var(--vbsfu-status-error)';
        }
    };
    console.log('Sliding Panel UI Injected.');
}

// --- Initial Check and Injection Trigger ---
function init() {
    const urlPattern = /\.lightning\.force\.com\/lightning\/r\/(Case|WorkOrder)\//;
    if (urlPattern.test(window.location.href)) {
        // Delay injection to ensure Salesforce page has finished its initial render.
        // Also check if the body is already loaded.
        if (document.body) {
             setTimeout(injectUI, 1500);
        } else {
             document.addEventListener('DOMContentLoaded', () => setTimeout(injectUI, 1500));
        }
    } else {
        console.log('Not a target Salesforce record page, panel not injected.');
    }
}

init();
