// panel_injector.js - Injects the UI and triggers the extension.

console.log("Salesforce Full View: UI Panel Injector Loaded.");

// --- Sliding Panel Code ---
function injectSlidingPanel() {
    console.log('Injecting Sliding Panel UI...');
    if (document.getElementById('vbsfu-panel')) return;

    const panel = document.createElement('div');
    panel.id = 'vbsfu-panel';
    panel.style.cssText = 'position:fixed; top:100px; right:-180px; width:160px; z-index:2147483647; background-color:#f0f0f0; border:1px solid #ccc; border-right:none; border-radius:5px 0 0 5px; box-shadow:-2px 2px 5px rgba(0,0,0,0.2); transition:right 0.3s ease-in-out; padding:10px; font-family:sans-serif; font-size:14px; display:flex; flex-direction:column; align-items:stretch;';

    const toggleButton = document.createElement('button');
    toggleButton.id = 'vbsfu-toggle';
    toggleButton.textContent = '🛠️';
    toggleButton.style.cssText = 'position:fixed; top:100px; right:10px; z-index:2147483647; padding:8px; cursor:pointer; border:1px solid #fd0; background-color:#bde; color:#33F; border-radius:5px; font-size:16px;';
    toggleButton.setAttribute('aria-label', 'Toggle Salesforce Utilities Panel');

    const title = document.createElement('h4');
    title.textContent = 'VB SF Utils';
    title.style.cssText = 'text-align:center; margin:0 0 10px 0; color:#000099;';

    const generateButton = document.createElement('button');
    generateButton.id = 'vbsfu-generate';
    generateButton.textContent = 'Generate Full View';
    generateButton.style.cssText = 'padding:8px; margin-bottom:5px; cursor:pointer;';

    const copyButton = document.createElement('button');
    copyButton.id = 'vbsfu-copy';
    copyButton.textContent = 'Copy Record Link';
    copyButton.style.cssText = 'padding:8px; margin-bottom:5px; cursor:pointer;';

    const statusDiv = document.createElement('div');
    statusDiv.id = 'vbsfu-status';
    statusDiv.style.cssText = 'font-size:12px; margin-top:8px; text-align:center; min-height:1em; color:#00802b;';

    panel.appendChild(title);
    panel.appendChild(generateButton);
    panel.appendChild(copyButton);
    panel.appendChild(statusDiv);
    document.body.appendChild(panel);
    document.body.appendChild(toggleButton);

    toggleButton.onclick = () => {
        panel.style.right = (panel.style.right === '0px') ? '-180px' : '0px';
        if (panel.style.right === '0px') statusDiv.textContent = '';
    };

    // This handler relies on functions from content.js (scrollIntoViewAndWait)
    generateButton.onclick = async () => {
        console.log('Panel Generate button clicked');
        generateButton.disabled = true;
        copyButton.disabled = true;

        try {
            statusDiv.textContent = 'Preparing page...';
            statusDiv.style.color = 'orange';

            // Scroll to wake up lazy-loaded related lists
            await scrollIntoViewAndWait('a.slds-card__header-link[href*="/related/PSM_Notes__r/view"]', statusDiv);
            await scrollIntoViewAndWait('div.forceListViewManager[aria-label*="Emails"]', statusDiv);

            window.scrollTo({ top: 0, behavior: 'auto' });
            console.log('Finished preparing page. Initiating generation.');

            statusDiv.textContent = 'Initiating...';
            chrome.runtime.sendMessage({ action: "initiateGenerateFullCaseView" }, (response) => {
                if (chrome.runtime.lastError) {
                    console.error("Error sending initiate message:", chrome.runtime.lastError.message);
                    statusDiv.textContent = `Error: ${chrome.runtime.lastError.message}`;
                    statusDiv.style.color = 'red';
                } else {
                    console.log("Background acknowledged initiation.");
                    statusDiv.textContent = "Processing initiated...";
                }
            });
        } catch (error) {
             console.error("Error during pre-generation scroll:", error);
             statusDiv.textContent = 'Error preparing page!';
             statusDiv.style.color = 'red';
        } finally {
             generateButton.disabled = false;
             copyButton.disabled = false;
        }
    };

    // This handler relies on functions from content.js (findCaseNumberSpecific)
    copyButton.onclick = () => {
        console.log('Panel Copy button clicked');
        statusDiv.textContent = '';
        statusDiv.style.color = 'green';

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
                    statusDiv.style.color = 'red';
                });
            } catch (error) {
                console.error('Clipboard API error:', error);
                statusDiv.textContent = 'Error: Clipboard API failed.';
                statusDiv.style.color = 'red';
            }
        } else {
            console.error("Failed to get record number for copy.");
            statusDiv.textContent = 'Error: Record # not found.';
            statusDiv.style.color = 'red';
        }
    };
    console.log('Sliding Panel UI Injected.');
}

// --- Initial Check and Injection Trigger ---
const urlPattern = /\.lightning\.force\.com\/lightning\/r\/(Case|WorkOrder)\//;
if (urlPattern.test(window.location.href)) {
    // Delay injection to ensure Salesforce page has finished its initial render.
    setTimeout(injectSlidingPanel, 1500);
} else {
    console.log('Not a target Salesforce record page, panel not injected.');
}
