/**
 * NoHands Background Script
 * Manages context menu and side panel
 */

// Open side panel when clicking the extension icon
chrome.action.onClicked.addListener(async (tab) => {
  await chrome.sidePanel.open({ windowId: tab.windowId });
});

// Create context menu on installation
chrome.runtime.onInstalled.addListener(() => {
  // Allow the side panel to be opened by clicking the icon
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });

  chrome.contextMenus.create({
    id: 'copyInputName',
    title: 'Copier le nom de l\'input',
    contexts: ['editable'],
    documentUrlPatterns: ['http://*/*', 'https://*/*']
  });

  console.log('NoHands: Context menu created');
});

// Handle context menu click
chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  if (info.menuItemId === 'copyInputName') {
    try {
      const response = await chrome.tabs.sendMessage(tab.id, {
        action: 'copyInputName'
      });

      if (response && response.success) {
        console.log('NoHands: Input name copied:', response.inputName);
      } else {
        console.warn('NoHands: Failed to copy input name:', response?.error || 'Unknown error');
      }
    } catch (error) {
      console.debug('NoHands: Content script not available:', error.message);
    }
  }
});

// Auto-inject content script into new tabs/popups as a fallback
// (declarative content_scripts should handle most cases, this covers edge cases)
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.status !== 'complete') return;
  const url = tab.url || '';
  if (!url.startsWith('http://') && !url.startsWith('https://')) return;

  chrome.scripting.executeScript({
    target: { tabId: tabId, allFrames: true },
    files: ['content.js']
  }).catch(() => {
    // Silently ignore — content script may already be injected via manifest
  });
});

console.log('NoHands: Background script loaded');
