// background.js — NoHands OSA
// Ouvre le side panel au clic sur l'icône, gère le menu contextuel
// "Copier le nom de l'input" et l'injection de secours du content script.

chrome.runtime.onInstalled.addListener(() => {
  // Ouvrir le panneau latéral en cliquant sur l'icône de l'extension
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });

  // Menu contextuel : clic droit sur un champ -> copie son attribut name
  chrome.contextMenus.create({
    id: "copyInputName",
    title: "Copier le nom de l'input",
    contexts: ["editable"],
    documentUrlPatterns: ["http://*/*", "https://*/*"]
  });
});

chrome.action.onClicked.addListener(async (tab) => {
  await chrome.sidePanel.open({ windowId: tab.windowId });
});

chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  if (info.menuItemId !== "copyInputName") return;
  try {
    const response = await chrome.tabs.sendMessage(tab.id, { action: "copyInputName" });
    if (!response || !response.success) {
      console.warn("NoHands OSA: copie du nom échouée:", response?.error || "inconnu");
    }
  } catch (error) {
    console.debug("NoHands OSA: content script indisponible:", error.message);
  }
});

// Injection de secours du content script dans les nouveaux onglets / popups
// (les content_scripts déclaratifs couvrent la plupart des cas, ceci couvre
// les pages ouvertes avant l'installation ou certains popups window.open).
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.status !== "complete") return;
  const url = tab.url || "";
  if (!url.startsWith("http://") && !url.startsWith("https://")) return;

  chrome.scripting.executeScript({
    target: { tabId: tabId, allFrames: true },
    files: ["content.js"]
  }).catch(() => {
    // Ignoré : déjà injecté via le manifest, ou page protégée.
  });
});
