// background.js
// Fait en sorte qu'un clic sur l'icône de l'extension ouvre le panneau latéral
// (le panneau reste ouvert même quand on clique sur la page, contrairement à un popup,
// ce qui est indispensable pour la fonction "choisir un élément sur la page").
chrome.sidePanel
  .setPanelBehavior({ openPanelOnActionClick: true })
  .catch((err) => console.error("sidePanel setPanelBehavior error:", err));
