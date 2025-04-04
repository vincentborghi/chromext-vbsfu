// Existing listener for generateViewBtn
document.getElementById('generateViewBtn').addEventListener('click', () => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    const urlPattern = /\.lightning\.force\.com\/lightning\/r\/(Case|WorkOrder)\//;
    if (tabs[0]?.id && tabs[0]?.url && urlPattern.test(tabs[0].url)) {
    // OLD if (tabs[0] && tabs[0].id && tabs[0].url && tabs[0].url.includes('.lightning.force.com/lightning/r/Case/')) {
      // Clear any previous status messages
      const statusDiv = document.getElementById('statusMessage');
      statusDiv.textContent = '';
      statusDiv.style.color = 'green'; // Reset color

      chrome.tabs.sendMessage(tabs[0].id, { action: "generateFullCaseView" }, (response) => {
        if (chrome.runtime.lastError) {
          console.error("Error sending generateFullCaseView message:", chrome.runtime.lastError.message);
          statusDiv.textContent = "Error: Could not connect. Refresh SF page.";
          statusDiv.style.color = 'red';
        } else if (response && response.status === "processing") {
           console.log("Content script is processing full view.");
           statusDiv.textContent = "Processing... New tab opening.";
           // Optionally close popup
           // setTimeout(() => window.close(), 1500);
        } else if (response && response.status === "error") {
           console.error("Error from content script (generateFullCaseView):", response.message);
           statusDiv.textContent = `Error: ${response.message}`;
           statusDiv.style.color = 'red';
        }
      });
    } else {
       document.getElementById('statusMessage').textContent = "Error: Not on SF Lightning Case page.";
       document.getElementById('statusMessage').style.color = 'red';
    }
  });
});

// --- NEW Listener for copyCaseNumberBtn ---
document.getElementById('copyCaseNumberBtn').addEventListener('click', () => {
  const statusDiv = document.getElementById('statusMessage');
  statusDiv.textContent = ''; // Clear previous messages
  statusDiv.style.color = 'green'; // Reset color

  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    if (tabs[0] && tabs[0].id && tabs[0].url && tabs[0].url.includes('.lightning.force.com/lightning/r/Case/')) {
      const tabId = tabs[0].id;
      console.log("Sending getCaseNumberAndUrl message to tab:", tabId);

      // Send message to content script to get case number and URL
      chrome.tabs.sendMessage(tabId, { action: "getCaseNumberAndUrl" }, (response) => {
        if (chrome.runtime.lastError) {
          console.error("Error sending getCaseNumberAndUrl message:", chrome.runtime.lastError.message);
          statusDiv.textContent = 'Error communicating with page.';
          statusDiv.style.color = 'red';
          return;
        }

        if (response && response.status === "success" && response.caseNumber && response.url) {
          const caseNumber = response.caseNumber;
          const caseUrl = response.url;
          const linkText = `Case ${caseNumber}`;
          const richTextHtml = `<a href="${caseUrl}">${linkText}</a>`;

          console.log(`Attempting to copy: ${richTextHtml}`);

          // Use the Clipboard API to write rich text
          try {
            // Create a Blob with HTML content
            const blobHtml = new Blob([richTextHtml], { type: 'text/html' });
            // Create a Blob with plain text fallback
            const blobText = new Blob([linkText], { type: 'text/plain' });

            // Create ClipboardItem
            const clipboardItem = new ClipboardItem({
              'text/html': blobHtml,
              'text/plain': blobText // Provide plain text fallback
            });

            // Write to clipboard
            navigator.clipboard.write([clipboardItem]).then(() => {
              console.log('Rich text link copied to clipboard!');
              statusDiv.textContent = `Copied: ${linkText}`;
              // Optionally close popup after success
              // setTimeout(() => window.close(), 1500);
            }).catch(err => {
              console.error('Failed to copy rich text: ', err);
              // Fallback to plain text if rich text fails (less common now)
              navigator.clipboard.writeText(linkText).then(() => {
                 console.log('Fallback: Copied plain text link label to clipboard!');
                 statusDiv.textContent = `Copied text: ${linkText}`;
              }).catch(fallbackErr => {
                 console.error('Failed to copy plain text fallback: ', fallbackErr);
                 statusDiv.textContent = 'Error: Failed to copy.';
                 statusDiv.style.color = 'red';
              });
            });
          } catch (error) {
            console.error('Clipboard API error:', error);
            statusDiv.textContent = 'Error: Clipboard API failed.';
            statusDiv.style.color = 'red';
          }

        } else {
          console.error("Failed to get case number/URL from content script. Response:", response);
          statusDiv.textContent = response?.message || 'Error: Case number not found.';
          statusDiv.style.color = 'red';
        }
      });
    } else {
      console.log("Button clicked, but not on a valid Salesforce Case page.");
      statusDiv.textContent = "Error: Not on SF Lightning Case page.";
      statusDiv.style.color = 'red';
    }
  });
});
