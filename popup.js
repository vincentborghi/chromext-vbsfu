// popup.js - Remove the event listeners

// Keep this console log to confirm the script runs if needed
console.log("Popup script loaded. Functionality moved to content script panel.");

// The status message can be set directly in popup.html or dynamically here if needed
// For example, check if on the right page:
/*
chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    const statusDiv = document.getElementById('statusMessage');
    const urlPattern = /\.lightning\.force\.com\/lightning\/r\/(Case|WorkOrder)\//;
    if (tabs[0]?.url && urlPattern.test(tabs[0].url)) {
        statusDiv.textContent = 'Tools active on this page.';
        statusDiv.style.color = 'green';
    } else {
        statusDiv.textContent = 'Tools only work on SF Case/WorkOrder pages.';
        statusDiv.style.color = 'grey';
    }
});
*/

// No button listeners needed here anymore.

