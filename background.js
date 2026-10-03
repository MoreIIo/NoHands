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

/* ====================================================================
 * SIGEO — champ « Mandat MG » (SelectorControl + AutoCompletion)
 * --------------------------------------------------------------------
 * Le champ n'est pas un simple input :
 *   - #body_x_selMan_x_txt_x__ctl0 : texte visible (libellé « MG0396166 ») ;
 *   - #body_x_selMan_x            : input caché = valeur SOUMISE (« 396166 ») ;
 *   - window.__ivCtrl["body_x_selMan_x"] : objet JS du contrôle (AutoPostBack).
 * ctrl.add(id, libellé) remplit caché + texte et déclenche un __doPostBack
 * qui recharge la liste Journal et les updatePanels. __ivCtrl n'existe que
 * dans le monde de la page : injection en world 'MAIN' obligatoire.
 * Appelé par le content script (marqueur « mg: ») via runtime.sendMessage.
 * ==================================================================== */

const MANDAT_MG = {
  CTRL_ID: "body_x_selMan_x",
  TIMEOUT_MS: 8000,        // attente max de la fin du postback
  // Mandat inexistant : si SIGEO accepte l'id mais ne charge aucun journal,
  // on considère que le mandat n'existe pas. Passer à false si un mandat
  // valide peut légitimement n'avoir aucun journal.
  REQUIRE_JOURNAL: true
};

// « MG0396166 », « mg396166 », « 396166 », « MG 0396166 » → { id, label }
function normalizeMandatMG(raw) {
  const s = String(raw ?? "").replace(/[\s\u00a0\u202f\u2007]/g, "").toUpperCase();
  const m = s.match(/^(?:MG)?(\d+)$/);
  if (!m) {
    return { error: `Mandat MG invalide : « ${String(raw ?? "").trim()} » (attendu MG0396166, mg396166 ou 396166)` };
  }
  const id = String(parseInt(m[1], 10));
  if (id === "0") return { error: `Mandat MG invalide : « ${String(raw).trim()} » (numéro nul)` };
  return { id, label: "MG" + id.padStart(7, "0") };
}

// --- Code injecté dans la page (world MAIN) : doit rester autonome ---------
// Pose la valeur, attend la fin du postback, puis vérifie le résultat.
// Les nœuds peuvent être remplacés par le postback : on ré-interroge le DOM
// par ID à chaque lecture, sans garder de référence.
async function mgPageSetAndWait(ctrlId, id, label, timeoutMs, requireJournal) {
  const JOURNAL_ID = "body_x_ddlJournal_ddlJournal";
  const txtId = ctrlId + "_txt_x__ctl0";
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  // Espaces insécables du DOM SIGEO : normalisés avant toute comparaison.
  const norm = (s) => String(s ?? "").replace(/[\u00a0\u202f\u2007]/g, " ").replace(/\s+/g, " ").trim().toUpperCase();
  const labelOk = (t) => {
    const n = norm(t);
    return n === label || (n.startsWith(label) && !/\d/.test(n.charAt(label.length)));
  };
  const read = () => ({
    hidden: String(document.getElementById(ctrlId)?.value ?? "").trim(),
    text: document.getElementById(txtId)?.value ?? ""
  });
  const journalSig = () => {
    const j = document.getElementById(JOURNAL_ID);
    return j ? Array.from(j.options).map((o) => o.value).join("|") : null;
  };
  const isBusy = () => {
    try {
      if (window.Sys?.WebForms?.PageRequestManager?.getInstance?.()?.get_isInAsyncPostBack?.()) return true;
    } catch (_) { /* ignoré */ }
    const p = document.getElementById("progessUpdatePanel") || document.querySelector('[id$="progessUpdatePanel"]');
    if (!p) return false;
    const cs = getComputedStyle(p);
    return cs.display !== "none" && cs.visibility !== "hidden" && p.getClientRects().length > 0;
  };

  const ctrl = window.__ivCtrl?.[ctrlId];
  if (!ctrl || typeof ctrl.add !== "function") {
    return { ok: false, code: "NOT_FOUND", error: "champ Mandat MG introuvable sur cette page" };
  }

  // Idempotence : déjà le bon mandat → aucun postback (évite les boucles
  // quand le re-remplissage automatique repasse sur le champ).
  const before = read();
  if (before.hidden === id && labelOk(before.text)) {
    return { ok: true, already: true, ...before };
  }

  const journalNodeBefore = document.getElementById(JOURNAL_ID);
  const sigBefore = journalSig();
  let journalTouched = false;
  let endRequestSeen = false;
  let serverError = null;

  const touches = (n) => n.nodeType === 1 && (n.id === JOURNAL_ID || !!n.querySelector?.("#" + JOURNAL_ID));
  const mo = new MutationObserver((muts) => {
    for (const m of muts) {
      if (m.target?.id === JOURNAL_ID ||
          Array.from(m.addedNodes).some(touches) ||
          Array.from(m.removedNodes).some(touches)) { journalTouched = true; return; }
    }
  });
  mo.observe(document.documentElement, { childList: true, subtree: true });

  let prm = null;
  const onEnd = (_sender, args) => {
    endRequestSeen = true;
    try { const err = args?.get_error?.(); if (err) serverError = err.message || String(err); } catch (_) { /* ignoré */ }
  };
  try {
    prm = window.Sys?.WebForms?.PageRequestManager?.getInstance?.() || null;
    prm?.add_endRequest(onEnd);
  } catch (_) { prm = null; }

  const cleanup = () => {
    mo.disconnect();
    try { prm?.remove_endRequest(onEnd); } catch (_) { /* ignoré */ }
  };

  try {
    if (typeof ctrl.removeAllValues === "function") ctrl.removeAllValues();
    ctrl.add(id, label); // remplit caché + texte, puis OnChange() → __doPostBack
  } catch (e) {
    cleanup();
    return { ok: false, code: "CTRL_ERROR", error: "erreur du contrôle SIGEO : " + (e?.message || e) };
  }

  // Attente de la fin du postback : signal PageRequestManager, ou liste
  // Journal remplacée / modifiée, ET plus d'indicateur de chargement.
  const t0 = Date.now();
  let finished = false;
  while (Date.now() - t0 < timeoutMs) {
    await sleep(100);
    if (isBusy()) continue;
    const changed = endRequestSeen || journalTouched ||
      journalSig() !== sigBefore || document.getElementById(JOURNAL_ID) !== journalNodeBefore;
    if (changed) {
      await sleep(250); // laisse les scripts de fin de postback s'exécuter
      if (!isBusy()) { finished = true; break; }
    }
  }
  cleanup();

  const after = read();
  const j = document.getElementById(JOURNAL_ID);
  const journalOptions = j ? Array.from(j.options).filter((o) => o.value && !/^(0|-1)$/.test(o.value)).length : null;
  const base = { ...after, journalOptions, elapsedMs: Date.now() - t0 };

  if (serverError) return { ok: false, code: "SERVER", error: "erreur serveur pendant le postback : " + serverError, ...base };
  if (!finished) return { ok: false, code: "TIMEOUT", error: `postback non terminé après ${timeoutMs} ms (liste Journal non rechargée)`, ...base };
  if (after.hidden !== id) {
    return { ok: false, code: "MISMATCH", error: `valeur soumise « ${after.hidden || "(vide)"} » au lieu de « ${id} » — mandat inexistant ?`, ...base };
  }
  if (!labelOk(after.text)) {
    return { ok: false, code: "MISMATCH", error: `libellé affiché « ${norm(after.text) || "(vide)"} » au lieu de « ${label} » — mandat inexistant ?`, ...base };
  }
  if (requireJournal && journalOptions === 0) {
    return { ok: false, code: "NO_JOURNAL", error: `aucun journal chargé pour ${label} — mandat inexistant ?`, ...base };
  }
  return { ok: true, ...base };
}

// Vérification seule (si le postback s'est transformé en rechargement
// complet de la page, l'injection précédente a été détruite).
function mgPageVerify(ctrlId, id, label) {
  const norm = (s) => String(s ?? "").replace(/[\u00a0\u202f\u2007]/g, " ").replace(/\s+/g, " ").trim().toUpperCase();
  const hidden = String(document.getElementById(ctrlId)?.value ?? "").trim();
  const text = document.getElementById(ctrlId + "_txt_x__ctl0")?.value ?? "";
  const n = norm(text);
  const ok = hidden === id && (n === label || (n.startsWith(label) && !/\d/.test(n.charAt(label.length))));
  return ok
    ? { ok: true, hidden, text, reloaded: true }
    : { ok: false, code: "MISMATCH", hidden, text, error: `après rechargement : valeur « ${hidden || "(vide)"} », libellé « ${norm(text) || "(vide)"} » (attendu ${id} / ${label})` };
}

function waitTabComplete(tabId, timeoutMs) {
  return new Promise((resolve) => {
    const t0 = Date.now();
    const tick = async () => {
      try {
        const tab = await chrome.tabs.get(tabId);
        if (tab.status === "complete") return resolve(true);
      } catch (_) { return resolve(false); }
      if (Date.now() - t0 > timeoutMs) return resolve(false);
      setTimeout(tick, 200);
    };
    setTimeout(tick, 300);
  });
}

/**
 * Sélectionne un mandat dans le champ « Mandat MG » de SIGEO et attend la
 * fin du postback. Ne clique JAMAIS sur un bouton de validation.
 * @param {number} tabId
 * @param {string|number} mg - MG0396166 | mg396166 | 396166
 * @param {{frameId?: number, ctrlId?: string, timeoutMs?: number}} [opts]
 * @returns {Promise<{ok: boolean, id?: string, label?: string, error?: string, code?: string, detail?: object}>}
 */
async function setMandatMG(tabId, mg, opts = {}) {
  const ctrlId = opts.ctrlId || MANDAT_MG.CTRL_ID;
  const timeoutMs = opts.timeoutMs || MANDAT_MG.TIMEOUT_MS;
  const target = { tabId, frameIds: [opts.frameId ?? 0] };

  const n = normalizeMandatMG(mg);
  if (n.error) {
    console.warn("NoHands OSA [Mandat MG]:", n.error);
    return { ok: false, code: "INVALID", error: n.error };
  }
  const { id, label } = n;
  console.log(`NoHands OSA [Mandat MG]: ${label} (id ${id}) → onglet ${tabId}, frame ${target.frameIds[0]}`);

  let res;
  try {
    const [inj] = await chrome.scripting.executeScript({
      target, world: "MAIN",
      func: mgPageSetAndWait,
      args: [ctrlId, id, label, timeoutMs, MANDAT_MG.REQUIRE_JOURNAL]
    });
    res = inj?.result;
  } catch (e) {
    // Rechargement complet de la page pendant l'attente : on vérifie après.
    if (!/frame|navigat|removed|context|closed/i.test(e.message || "")) {
      console.error("NoHands OSA [Mandat MG]: injection impossible:", e);
      return { ok: false, id, label, code: "INJECT", error: "injection impossible : " + e.message };
    }
    console.log("NoHands OSA [Mandat MG]: page rechargée pendant le postback, vérification…");
    await waitTabComplete(tabId, timeoutMs);
    try {
      const [inj] = await chrome.scripting.executeScript({
        target: { tabId, frameIds: [0] }, world: "MAIN",
        func: mgPageVerify, args: [ctrlId, id, label]
      });
      res = inj?.result;
    } catch (e2) {
      return { ok: false, id, label, code: "INJECT", error: "vérification impossible après rechargement : " + e2.message };
    }
  }

  if (!res) return { ok: false, id, label, code: "INJECT", error: "aucun résultat de la page" };
  const { ok, error, code, ...detail } = res;
  if (ok) {
    console.log(`NoHands OSA [Mandat MG]: ✓ ${label}${detail.already ? " (déjà en place)" : ` (${detail.elapsedMs ?? "?"} ms, ${detail.journalOptions ?? "?"} journal(aux))`}`);
    return { ok: true, id, label, detail };
  }
  console.warn(`NoHands OSA [Mandat MG]: ✗ ${label} — ${error}`, detail);
  return { ok: false, id, label, code, error: `Mandat MG ${label} : ${error}`, detail };
}

/* ====================================================================
 * SIGEO — moteur générique des sélecteurs (SelectorControl + AutoCompletion)
 * --------------------------------------------------------------------
 * Utilisé par l'étape de scénario « Sélecteur SIGEO » (sigeoSelector) :
 *   - mode « direct »    : id connu → ctrl.add(id, libellé) (cas Mandat MG) ;
 *   - mode « recherche » : frappe simulée dans l'input visible, attente des
 *     suggestions AJAX, clic sur la bonne, attente de la sélection (cas Compte).
 * Validation finale : SelectedValues.length === 1 ET texte visible qui
 * commence par la valeur attendue. Aucun clic sur un bouton de validation.
 * ==================================================================== */

const SIGEO_SELECTOR = {
  COMPTE_TIMEOUT_MS: 5000,
  DEFAULT_TIMEOUT_MS: 5000,
  // Mots-clés servant à repérer les contrôles « Compte » des lignes de
  // contrepartie quand aucune clé explicite n'est fournie.
  COMPTE_KEY_RE: "(cpt|compte|account)",
  // Clés connues (format id, {n} = n° de ligne) — essayées avant la détection
  // auto. Relevé sur « Saisie d'opérations diverses » (ventilation).
  COMPTE_KEY_TEMPLATES: ["body_x_proxyTabSaisie_x__compte_ajax_selector_{n}_x_selCompte_x"]
};

// --- Injecté en world MAIN : autonome -----------------------------------
// cfg : { ctrlKey?, compteAuto?: { ligne, template, keyRe }, mode, value,
//         id?, label?, expectText, timeoutMs }
async function sselPageRun(cfg) {
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const norm = (s) => String(s ?? "").replace(/[\u00a0\u202f\u2007]/g, " ").replace(/\s+/g, " ").trim().toUpperCase();
  const t0 = Date.now();
  const deadline = t0 + (cfg.timeoutMs || 5000);
  const reg = () => window.__ivCtrl || {};
  const allKeys = () => Object.keys(reg());
  const hiddenEl = (k) => document.getElementById(k);
  const textEl = (k) => document.getElementById(k + "_txt_x__ctl0") ||
    document.querySelector(`input[id^="${CSS.escape(k)}"][type="text"], input[id^="${CSS.escape(k)}"]:not([type])`);
  const isVisible = (el) => {
    if (!el) return false;
    const cs = getComputedStyle(el);
    return cs.display !== "none" && cs.visibility !== "hidden" && el.getClientRects().length > 0;
  };
  const isBusy = () => {
    try { if (window.Sys?.WebForms?.PageRequestManager?.getInstance?.()?.get_isInAsyncPostBack?.()) return true; } catch (_) { /* ignoré */ }
    return isVisible(document.getElementById("progessUpdatePanel") || document.querySelector('[id$="progessUpdatePanel"]'));
  };
  const selCount = (ctrl, k) => {
    try {
      let sv = ctrl.SelectedValues;
      if (typeof sv === "function") sv = sv.call(ctrl);
      if (sv && typeof sv.length === "number") return sv.length;
    } catch (_) { /* ignoré */ }
    // Repli : valeur cachée soumise (une valeur, pas de séparateur)
    const h = String(hiddenEl(k)?.value ?? "").trim();
    return h ? h.split(/[,;|]/).filter(Boolean).length : 0;
  };
  const textOk = (k) => {
    const want = norm(cfg.expectText);
    return !!want && norm(textEl(k)?.value).startsWith(want);
  };
  const snapshot = (k) => ({
    ctrlKey: k,
    id: String(hiddenEl(k)?.value ?? "").trim(),
    label: String(textEl(k)?.value ?? "").replace(/[\u00a0\u202f\u2007]/g, " ").trim()
  });

  // 1. Résolution de la clé __ivCtrl (attendue jusqu'au timeout : le
  //    contrôle peut n'apparaître qu'après « Saisir les contreparties »).
  // Clé normalisée : format name (body:x:…) → id (body_x_…), et clé de
  // l'input texte (…_txt_x__ctl0, aussi présente dans __ivCtrl) → clé du
  // contrôle lui-même (sinon le texte tapé passerait pour une sélection).
  const canon = (raw) => {
    const k = String(raw || "").trim().replace(/[:$]/g, "_");
    const base = k.replace(/_txt_x__ctl0$/, "");
    if (base !== k && reg()[base]) return base;
    return reg()[k] ? k : null;
  };
  const resolveKey = () => {
    if (cfg.ctrlKey) return canon(cfg.ctrlKey);
    const ca = cfg.compteAuto;
    if (!ca) return null;
    const n = Math.max(1, parseInt(ca.ligne, 10) || 1);
    // Modèle : accepte aussi le format name ASP.NET (body:x:… ou body$x$…),
    // converti au format id utilisé par __ivCtrl (body_x_…).
    const fromTpl = (tpl) => canon(String(tpl)
      .replace(/\{n\}/g, n).replace(/\{n0\}/g, n - 1).replace(/\{ctl\}/g, String(n + 1).padStart(2, "0")));
    if (ca.template) return fromTpl(ca.template);
    for (const tpl of ca.knownTemplates || []) {
      const k = fromTpl(tpl);
      if (k) return k;
    }
    const re = new RegExp(ca.keyRe, "i");
    let keys = allKeys().filter((k) => re.test(k) && !/selMan/i.test(k));
    // « …_x_txt_x__ctl0 » est l'input texte d'un contrôle déjà listé : doublon.
    keys = keys.filter((k) => !(k.endsWith("_txt_x__ctl0") && keys.includes(k.slice(0, -"_txt_x__ctl0".length))));
    // N° de ligne intégré à la clé (…_selector_1_x_…) : on le préfère à l'ordre de la page.
    const byNum = keys.filter((k) => { const m = k.match(/_(\d+)_x/); return m && parseInt(m[1], 10) === n; });
    if (byNum.length === 1) return byNum[0];
    const found = keys
      .map((k) => ({ k, el: hiddenEl(k) || textEl(k) }))
      .filter((x) => x.el);
    found.sort((a, b) => (a.el.compareDocumentPosition(b.el) & Node.DOCUMENT_POSITION_FOLLOWING) ? -1 : 1);
    return found[n - 1]?.k || null;
  };
  let key = resolveKey();
  while (!key && Date.now() < deadline) { await sleep(200); key = resolveKey(); }
  if (!key) {
    const keys = allKeys();
    return {
      ok: false, code: "NOT_FOUND",
      error: (cfg.ctrlKey ? `contrôle « ${cfg.ctrlKey} » introuvable` : `contrôle Compte de la ligne ${cfg.compteAuto?.ligne} introuvable`) +
        (keys.length ? ` — clés __ivCtrl présentes : ${keys.slice(0, 12).join(", ")}${keys.length > 12 ? "…" : ""}` : " — aucun __ivCtrl sur cette page"),
      availableKeys: keys
    };
  }
  const ctrl = reg()[key];

  // 2. Idempotence : déjà la bonne sélection → rien à faire.
  if (selCount(ctrl, key) === 1 && textOk(key) && (!cfg.id || snapshot(key).id === String(cfg.id))) {
    return { ok: true, already: true, ...snapshot(key) };
  }

  // Suivi des postbacks déclenchés par la sélection
  let prm = null, endSeen = false, serverError = null;
  const onEnd = (_s, args) => {
    endSeen = true;
    try { const e = args?.get_error?.(); if (e) serverError = e.message || String(e); } catch (_) { /* ignoré */ }
  };
  try { prm = window.Sys?.WebForms?.PageRequestManager?.getInstance?.() || null; prm?.add_endRequest(onEnd); } catch (_) { prm = null; }
  const cleanup = () => { try { prm?.remove_endRequest(onEnd); } catch (_) { /* ignoré */ } };
  const settle = async () => {
    // Laisse partir un éventuel postback, puis attend qu'il soit fini.
    await sleep(150);
    while (Date.now() < deadline && isBusy()) await sleep(100);
    await sleep(200);
  };

  try {
    if (typeof ctrl.removeAllValues === "function") ctrl.removeAllValues();

    if (cfg.mode === "direct") {
      if (typeof ctrl.add !== "function") return { ok: false, code: "CTRL_ERROR", error: `le contrôle « ${key} » n'a pas de méthode add() — essaie le mode « recherche »` };
      ctrl.add(String(cfg.id), String(cfg.label));
      await settle();
    } else {
      // --- mode recherche : frappe simulée -----------------------------
      const txt = textEl(key);
      if (!txt) return { ok: false, code: "NOT_FOUND", error: `input texte du contrôle « ${key} » introuvable` };
      const typed = String(cfg.value);
      const suggestSelect = () => {
        const id = txt.id || "";
        let sel = id ? document.querySelector(`select[name="searchResultSelect_${CSS.escape(id)}"]`) : null;
        if (!sel && id) sel = document.getElementById("search:" + id)?.querySelector("select") || null;
        return sel;
      };
      const prevSig = (() => { const s = suggestSelect(); return s ? Array.from(s.options).map((o) => o.text).join("|") : ""; })();
      // Repli si la page n'utilise pas searchResultSelect_ : on collecte les
      // éléments ajoutés au DOM pendant l'attente (liste AJAX quelconque).
      const added = new Set();
      const addObs = new MutationObserver((muts) => {
        for (const m of muts) for (const nd of m.addedNodes) if (nd.nodeType === 1) added.add(nd);
      });
      addObs.observe(document.body || document.documentElement, { childList: true, subtree: true });
      const genericSuggestions = () => {
        const want = norm(typed);
        const out = [];
        for (const root of added) {
          if (!root.isConnected) continue;
          const nodes = [root, ...root.querySelectorAll("li, option, tr, a, div, span, td")];
          for (const el of nodes) {
            if (el === txt || el.contains(txt) || !isVisible(el)) continue;
            const t = norm(el.textContent);
            if (!t || t.length > 200 || !t.startsWith(want)) continue;
            // Élément le plus « feuille » : pas d'enfant qui porte déjà le même texte
            if (Array.from(el.children).some((c) => norm(c.textContent) === t)) continue;
            if (!out.some((o) => o.el === el)) out.push({ el, text: el.textContent.replace(/\s+/g, " ").trim() });
          }
        }
        // Un conteneur qui englobe d'autres suggestions n'en est pas une
        const leaves = out.filter((o) => !out.some((x) => x !== o && o.el.contains(x.el)));
        // Dédoublonne par texte (ex. <tr> et <td> identiques)
        const seen = new Set();
        return leaves.filter((o) => (seen.has(norm(o.text)) ? false : (seen.add(norm(o.text)), true)));
      };
      txt.focus();
      txt.value = "";
      for (const ch of typed) {
        txt.value += ch;
        const kc = ch.toUpperCase().charCodeAt(0);
        txt.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: ch, keyCode: kc, which: kc }));
        txt.dispatchEvent(new KeyboardEvent("keypress", { bubbles: true, cancelable: true, key: ch, keyCode: ch.charCodeAt(0), which: ch.charCodeAt(0) }));
        txt.dispatchEvent(new Event("input", { bubbles: true }));
        txt.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, cancelable: true, key: ch, keyCode: kc, which: kc }));
        await sleep(30);
      }

      // Attente : sélection automatique (match exact) OU liste de suggestions.
      let picked = null;
      while (Date.now() < deadline) {
        await sleep(200);
        if (selCount(ctrl, key) === 1 && textOk(key)) break; // auto-sélectionné
        const sel = suggestSelect();
        if (!sel) {
          const gen = genericSuggestions();
          if (!gen.length) continue;
          const want = norm(typed);
          const code = (t) => norm(t).split(/\s+-\s+|\s+/)[0];
          const exact = gen.filter((o) => code(o.text) === want);
          const choiceG = exact.length === 1 ? exact[0] : (exact.length === 0 && gen.length === 1 ? gen[0] : null);
          if (!choiceG) {
            addObs.disconnect();
            const pool = exact.length ? exact : gen;
            return {
              ok: false, code: "AMBIGUOUS",
              error: `plusieurs suggestions pour « ${typed} » : ${pool.slice(0, 5).map((o) => o.text).join(" / ")}${pool.length > 5 ? "…" : ""}`,
              ...snapshot(key)
            };
          }
          picked = choiceG.text;
          for (const type of ["mouseover", "mousedown", "mouseup", "click"]) {
            choiceG.el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
          }
          break;
        }
        const opts = Array.from(sel.options).filter((o) => norm(o.text));
        if (!opts.length) continue;
        const sig = opts.map((o) => o.text).join("|");
        if (sig === prevSig && !isVisible(sel)) continue; // ancienne liste cachée
        const want = norm(typed);
        const code = (t) => norm(t).split(/\s+-\s+|\s+/)[0];
        const exact = opts.filter((o) => code(o.text) === want);
        const starts = opts.filter((o) => norm(o.text).startsWith(want));
        const choice = exact.length === 1 ? exact[0] : (exact.length === 0 && starts.length === 1 ? starts[0] : null);
        if (!choice) {
          addObs.disconnect();
          const pool = exact.length ? exact : (starts.length ? starts : opts);
          return {
            ok: false, code: exact.length || starts.length ? "AMBIGUOUS" : "NO_MATCH",
            error: (exact.length || starts.length
              ? `plusieurs suggestions pour « ${typed} »`
              : `aucune suggestion ne commence par « ${typed} »`) +
              ` : ${pool.slice(0, 5).map((o) => o.text.trim()).join(" / ")}${pool.length > 5 ? "…" : ""}`,
            ...snapshot(key)
          };
        }
        picked = choice.text.trim();
        choice.selected = true;
        sel.value = choice.value;
        for (const type of ["mousedown", "mouseup", "click"]) {
          choice.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
        }
        sel.dispatchEvent(new Event("change", { bubbles: true }));
        break;
      }
      addObs.disconnect();
      await settle();
      if (picked) cfg._picked = picked;
    }
  } catch (e) {
    cleanup();
    return { ok: false, code: "CTRL_ERROR", error: "erreur du contrôle SIGEO : " + (e?.message || e), ...snapshot(key) };
  }

  // 3. Validation (attend la sélection jusqu'au timeout)
  while (Date.now() < deadline && !(selCount(ctrl, key) === 1 && textOk(key))) await sleep(150);
  while (Date.now() < deadline && isBusy()) await sleep(100);
  cleanup();

  const snap = snapshot(key);
  const n = selCount(ctrl, key);
  const base = { ...snap, selectedCount: n, picked: cfg._picked || null, elapsedMs: Date.now() - t0, postback: endSeen };
  if (serverError) return { ok: false, code: "SERVER", error: "erreur serveur : " + serverError, ...base };
  if (n !== 1) return { ok: false, code: n ? "MULTIPLE" : "NOT_SELECTED", error: n ? `${n} valeurs sélectionnées au lieu d'une` : `aucune valeur sélectionnée pour « ${cfg.value} » (délai ${cfg.timeoutMs} ms dépassé ?)`, ...base };
  if (!textOk(key)) return { ok: false, code: "MISMATCH", error: `texte affiché « ${snap.label || "(vide)"} » ne commence pas par « ${cfg.expectText} »`, ...base };
  if (cfg.id && snap.id !== String(cfg.id)) return { ok: false, code: "MISMATCH", error: `valeur soumise « ${snap.id || "(vide)"} » au lieu de « ${cfg.id} »`, ...base };
  return { ok: true, ...base };
}

/**
 * Sélecteur SIGEO générique.
 * @param {number} tabId
 * @param {{ctrlKey?: string, compteAuto?: object, mode: 'direct'|'recherche', value: string,
 *          timeoutMs?: number, frameId?: number}} opts
 * @returns {Promise<{ok: boolean, id?: string, label?: string, error?: string, code?: string, detail?: object}>}
 */
async function setSigeoSelector(tabId, opts) {
  const mode = opts.mode === "direct" ? "direct" : "recherche";
  const value = String(opts.value ?? "").replace(/[\u00a0\u202f\u2007]/g, " ").trim();
  if (!value) return { ok: false, code: "INVALID", error: "valeur vide" };
  if (!opts.ctrlKey && !opts.compteAuto) return { ok: false, code: "INVALID", error: "clé __ivCtrl manquante" };

  // Mode direct générique : « id|libellé », sinon id = chiffres de la valeur.
  let id = null, label = null, expectText = value;
  if (mode === "direct") {
    const pipe = value.split("|");
    if (pipe.length >= 2) { id = pipe[0].trim(); label = pipe.slice(1).join("|").trim(); }
    else {
      const digits = value.replace(/\D/g, "");
      id = digits ? String(parseInt(digits, 10)) : value;
      label = value;
    }
    if (!id || !label) return { ok: false, code: "INVALID", error: `valeur « ${value} » : attendu « id|libellé » ou un code numérique` };
    expectText = label;
  }

  const cfg = {
    ctrlKey: opts.ctrlKey || null, compteAuto: opts.compteAuto || null,
    mode, value, id, label, expectText,
    timeoutMs: opts.timeoutMs || SIGEO_SELECTOR.DEFAULT_TIMEOUT_MS
  };
  const what = cfg.ctrlKey || `compte ligne ${cfg.compteAuto?.ligne}`;
  console.log(`NoHands OSA [Sélecteur SIGEO]: ${what} ← « ${value} » (${mode})`);
  try {
    const [inj] = await chrome.scripting.executeScript({
      target: { tabId, frameIds: [opts.frameId ?? 0] }, world: "MAIN",
      func: sselPageRun, args: [cfg]
    });
    const res = inj?.result;
    if (!res) return { ok: false, code: "INJECT", error: "aucun résultat de la page" };
    const { ok, error, code, ...detail } = res;
    if (ok) {
      console.log(`NoHands OSA [Sélecteur SIGEO]: ✓ ${what} → id ${detail.id}, « ${detail.label} »${detail.already ? " (déjà en place)" : ""}`);
      return { ok: true, id: detail.id, label: detail.label, detail };
    }
    console.warn(`NoHands OSA [Sélecteur SIGEO]: ✗ ${what} — ${error}`, detail);
    return { ok: false, id: detail.id, label: detail.label, code, error, detail };
  } catch (e) {
    console.error("NoHands OSA [Sélecteur SIGEO]: injection impossible:", e);
    return { ok: false, code: "INJECT", error: "injection impossible : " + e.message };
  }
}

/**
 * Compte d'une ligne de contrepartie (OD) — mode recherche.
 * La clé __ivCtrl est trouvée automatiquement (N-ième contrôle dont la clé
 * contient cpt/compte/account, dans l'ordre de la page), ou construite depuis
 * un modèle : {n} = ligne, {n0} = ligne-1, {ctl} = ligne+1 sur 2 chiffres.
 */
async function setCompte(tabId, ligne, valeur, opts = {}) {
  const n = Math.max(1, parseInt(ligne, 10) || 1);
  return setSigeoSelector(tabId, {
    compteAuto: {
      ligne: n, template: (opts.keyTemplate || "").trim() || null,
      knownTemplates: SIGEO_SELECTOR.COMPTE_KEY_TEMPLATES, keyRe: SIGEO_SELECTOR.COMPTE_KEY_RE
    },
    mode: opts.mode || "recherche",
    value: valeur,
    timeoutMs: opts.timeoutMs || SIGEO_SELECTOR.COMPTE_TIMEOUT_MS,
    frameId: opts.frameId
  });
}

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request?.action !== "setCompte" && request?.action !== "setSigeoSelector") return false;
  const tabId = request.tabId ?? sender.tab?.id;
  if (tabId == null) { sendResponse({ ok: false, code: "NO_TAB", error: "onglet inconnu" }); return false; }
  const p = request.action === "setCompte"
    ? setCompte(tabId, request.ligne, request.valeur, request)
    : setSigeoSelector(tabId, request);
  p.then(sendResponse).catch((e) => sendResponse({ ok: false, code: "INTERNAL", error: e.message }));
  return true;
});

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request?.action !== "setMandatMG") return false;
  const tabId = request.tabId ?? sender.tab?.id;
  if (tabId == null) {
    sendResponse({ ok: false, code: "NO_TAB", error: "onglet inconnu" });
    return false;
  }
  setMandatMG(tabId, request.mg, {
    frameId: request.frameId ?? sender.frameId ?? 0,
    ctrlId: request.ctrlId,
    timeoutMs: request.timeoutMs
  })
    .then(sendResponse)
    .catch((e) => sendResponse({ ok: false, code: "INTERNAL", error: e.message }));
  return true; // réponse asynchrone
});
