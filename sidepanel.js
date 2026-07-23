// ============================================================
// NoHands OSA — sidepanel.js
// Fusion des projets OSA (extraction Excel <-> web en boucle) et
// NoHands (remplissage de formulaires depuis des données Excel).
//
// Structure :
//   1. État global + utilitaires
//   2. Colonnes : détection d'en-têtes, modèles configurables
//   3. Chargement des données (fichier / collage / JSON)
//   4. Saisie : sélection de ligne, champs, mapping, remplissage, scénario
//   5. Extraction : conditions, recherche, boucle, résultats
//   6. Export, profils, onglets, initialisation
// ============================================================

/* ================== 1. ÉTAT GLOBAL ================== */

const state = {
  workbook: null,          // workbook XLSX si chargé depuis un fichier
  sheetName: null,
  rows: [],                // tableau 2D brut (0-indexé)
  originalFileName: "resultat.xlsx",
  headerMode: "auto",      // "auto" | "yes" | "no"
  modelName: "",           // modèle appliqué quand pas d'en-têtes ("" = lettres)
  selectedRowIdx: null,    // index absolu dans rows (ligne active en Saisie)
  mapping: {},             // nomColonne -> [nomsInput]
  customFields: [],        // [{name, value}]
  valueRules: [],          // [{from:[synonymes], to:"valeur cible"}] — conversion avant saisie
  targetTabId: null,       // onglet du formulaire mémorisé (dernier 🎯) — indép. de l'onglet affiché
  originalArrayBuffer: null, // octets bruts du .xlsx d'origine (pour export fidèle via ExcelJS)
  hasOriginalFile: false    // true si les données viennent d'un fichier .xlsx (mise en forme dispo)
};

/* ---------- Stockage du fichier d'origine (IndexedDB) ----------
   Les octets du .xlsx sont conservés hors de chrome.storage (trop petit / non binaire)
   afin de pouvoir réécrire le fichier en gardant toute la mise en forme, même après
   réouverture du panneau ou pour de gros fichiers. */
const IDB_NAME = "osa-files", IDB_STORE = "original", IDB_KEY = "current";
function idbOpen() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => { req.result.createObjectStore(IDB_STORE); };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function idbSaveOriginal(buf) {
  const db = await idbOpen();
  try {
    await new Promise((res, rej) => {
      const tx = db.transaction(IDB_STORE, "readwrite");
      tx.objectStore(IDB_STORE).put(buf, IDB_KEY);
      tx.oncomplete = res; tx.onerror = () => rej(tx.error);
    });
  } finally { db.close(); }
}
async function idbLoadOriginal() {
  const db = await idbOpen();
  try {
    return await new Promise((res, rej) => {
      const tx = db.transaction(IDB_STORE, "readonly");
      const r = tx.objectStore(IDB_STORE).get(IDB_KEY);
      r.onsuccess = () => res(r.result || null);
      r.onerror = () => rej(r.error);
    });
  } finally { db.close(); }
}
async function idbClearOriginal() {
  try {
    const db = await idbOpen();
    await new Promise((res) => {
      const tx = db.transaction(IDB_STORE, "readwrite");
      tx.objectStore(IDB_STORE).delete(IDB_KEY);
      tx.oncomplete = res; tx.onerror = res;
    });
    db.close();
  } catch (e) { /* ignoré */ }
}

let models = [];           // [{name, columns: []}]
let allMappings = {};      // clé de colonnes -> mapping
let profiles = {};         // nom -> config
let stopRequested = false;
let isRunning = false;
let hasStartedRun = false;
let hasDownloaded = false;
let runLog = [];
let lastRunOutputs = [];
let runDurationMs = 0;     // durée réelle de la dernière extraction (pour le récap)
let workingReady = false;  // true une fois l'init terminée : autorise la sauvegarde auto des options

// Modèles par défaut (hérités de NoHands) — modifiables/supprimables par l'utilisateur.
const DEFAULT_MODELS = [
  {
    name: "PROP",
    columns: [
      "N° PROP (TW)", "CIVILITE PROP", "NOM PROP", "PRENOM PROP",
      "ADRESSE LIGNE 1 PROP", "ADRESSE LIGNE 2 PROP", "CP PROP", "VILLE PROP",
      "TELEPHONE DOMICILE PROP", "TELEPHONE BUREAU PROP", "TELEPHONE PORTABLE PROP",
      "EMAIL PROP", "IBAN PROP", "FREQUENCE REGLT ACOMPTE PROP",
      "FREQUENCE REEDITION PROP", "MODE REGLT AU PROP", "TAUX HONOS PROP",
      "ASSURANCE GL (O/N)", "TAUX ASSURANCE GLI", "TAUX HONOS/ASSURANCE BASE 1",
      "DECLARATION REVENUS FONCIERS ADRF (O/N)", "TYPE GARANTIE",
      "DATE DEBUT MANDAT PROP", "NOM GESTIONNAIRE", "PRENOM GESTIONNAIRE",
      "Opérateur saisie"
    ]
  },
  {
    name: "LOTS",
    columns: [
      "N° PROPRIETAIRE (Tw)", "NOM PROPRIETAIRE", "ADRESSE LOT", "CATEGORIE",
      "N° LOT", "ETAT DU LOT", "ETAGE DU LOT", "TYPE DE LOT",
      "LIBELLE TYPE DE LOT", "N°APPARTEMENT", "SURFACE DU LOT",
      "NOM LOCATAIRE", "REGIME FISCAL"
    ]
  },
  {
    name: "BAIL",
    columns: [
      "N° PROP", "NOM PROPRIETAIRE", "NOM IMMEUBLE", "CIVILITE LOCATAIRE",
      "NOM LOCATAIRE", "PRENOM LOCATAIRE", "DATE DE NAISSANCE", "LIEU DE NAISSANCE",
      "ADRESSE LIGNE 1 LOCATAIRE", "ADRESSE LIGNE 2 LOCATAIRE", "CP LOCATAIRE",
      "VILLE LOCATAIRE", "TELEPHONE DOMICILE LOCATAIRE", "TELEPHONE N°2",
      "TELEPHONE PORTABLE LOCATAIRE", "EMAIL LOCATAIRE", "IBAN MANDAT SEPA LOCATAIRE",
      "BIC MANDAT SEPA LOCATAIRE", "DATE ENTREE LOCATAIRE", "CODE TYPE BAIL",
      "LIBELLE BAIL", "N° INDICE (5 IRL/1 ICC INSEE/11 ILC COMMERCIAUX )",
      "DATE PROCHAINE REVISION LOYER", "DATE DERNIERE REVISION LOYER",
      "FREQUENCE REVISION LOYER", "ANNEE DERNIERE REVISION LOYER",
      "TRIMESTRE REFERENCE DERNIERE REVISION LOYER", "ANNEE PROCHAINE REVISION LOYER",
      "TRIMESTRE PROCHAINE REVISION LOYER", "MODE REGLT LOCATAIRE",
      "TERME AVANCE/ECHU LOYER", "FREQUENCE APPEL LOYER",
      "DEPOT DE GARANTIE CONSERVE EN AGENCE", "DEPOT DE GARANTIE REVERSE AU PROPRIETAIRE",
      "DATE DEBUT ASSURANCE MULTIRISQUES", "DATE FIN ASSURANCE MULTIRISQUES",
      "SURFACE DU LOT", "TYPE DE LOT", "N° PORTE", "NOMBRE DE GARANTS",
      "CIVILITE GARANT", "NOM GARANT", "PRENOM GARANT", "ADRESSE LIGNE 1 GARANT",
      "ADRESSE LIGNE 2 GARANT", "CP GARANT", "VILLE GARANT", "RUM MANDAT SEPA LOCATAIRE"
    ]
  }
];

/* ---------- Utilitaires ---------- */

function $(id) { return document.getElementById(id); }

function escapeHtml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeAttr(str) {
  return String(str ?? "").replace(/&/g, "&amp;").replace(/"/g, "&quot;");
}

function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

// Formate une durée en millisecondes → "12 s", "3 min 20 s", "1 h 05 min".
function formatDuration(ms) {
  ms = Math.max(0, Math.round(ms));
  const totalS = Math.round(ms / 1000);
  if (totalS < 60) return totalS + " s";
  const m = Math.floor(totalS / 60), rs = totalS % 60;
  if (m < 60) return rs ? `${m} min ${String(rs).padStart(2, "0")} s` : `${m} min`;
  const h = Math.floor(m / 60), rm = m % 60;
  return rm ? `${h} h ${String(rm).padStart(2, "0")} min` : `${h} h`;
}

function hashString(str) {
  let h = 5381;
  for (let i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) >>> 0;
  return h.toString(36);
}

function indexToLetter(idx) {
  let n = idx + 1, s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function letterToIndex(letter) {
  letter = (letter || "").trim().toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(letter)) return -1;
  let n = 0;
  for (let i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
  return n - 1;
}

function formatIban(str) {
  if (!str || typeof str !== "string") return str;
  const s = str.replace(/\s/g, "").trim();
  if (s.length < 4) return str;
  const parts = [s.slice(0, 4)];
  for (let i = 4; i < s.length; i += 4) parts.push(s.slice(i, i + 4));
  return parts.join(" ");
}

function formatTrimester(value) {
  if (value === undefined || value === null) return value;
  const s = String(value).trim();
  if (!s) return value;
  if (/^T\d+$/i.test(s)) return s.toUpperCase();
  const digits = s.replace(/[^\d]/g, "");
  if (!digits) return value;
  return `T${digits}`;
}

// Formatage automatique selon le nom de colonne (hérité de NoHands)
function formatValueForColumn(colName, value) {
  const upper = (colName || "").toUpperCase();
  if (upper.includes("IBAN")) return formatIban(String(value));
  if (upper.includes("TRIMESTRE")) return formatTrimester(value);
  return value;
}

/* ---------- Formatage d'une valeur extraite via regex ----------
   Appliqué (mode Extraction) à chaque résultat avant écriture dans le tableau.
   - mode "extract" : renvoie le 1er groupe capturant ( ) s'il existe, sinon
     toute la correspondance ; si rien ne correspond, renvoie la valeur brute.
   - mode "replace" : remplace toutes les correspondances (drapeau global) par
     le texte de remplacement ($1, $2… acceptés).
   Regex invalide ou motif vide → valeur brute renvoyée (jamais d'exception). */
function applyOutputRegex(value, o) {
  if (!o || !o.regexEnabled) return value;
  const pattern = (o.regexPattern || "").trim();
  if (!pattern) return value;
  const str = value == null ? "" : String(value);
  const flags = o.regexFlagI ? "i" : "";
  try {
    if (o.regexMode === "replace") {
      return str.replace(new RegExp(pattern, "g" + flags), o.regexReplace || "");
    }
    const m = str.match(new RegExp(pattern, flags));
    if (!m) return str;
    return m[1] != null ? m[1] : m[0];
  } catch (e) {
    return str;
  }
}

/* ---------- Règles de valeurs (conversion d'une valeur avant la saisie) ----------
   Ex. : l'Excel contient "female" → on saisit "Mme" (option d'un <select>).
   Générique : marche pour civilité, Oui/Non, codes pays, etc. La valeur convertie
   passe ensuite par la logique de correspondance normale (select/radio/texte). */
const DEFAULT_VALUE_RULES = [
  { from: ["female", "femme", "f", "madame", "mme", "woman", "mrs", "ms", "féminin", "feminin"], to: "Mme" },
  { from: ["male", "homme", "m", "monsieur", "mr", "man", "masculin"], to: "M." }
];

// Normalise pour comparaison : minuscules, sans accents, sans espaces superflus.
function normalizeForMatch(s) {
  return String(s ?? "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().trim();
}

// Renvoie la valeur convertie si une règle correspond, sinon la valeur d'origine.
function applyValueRules(value) {
  if (value === undefined || value === null || value === "") return value;
  const norm = normalizeForMatch(value);
  if (!norm) return value;
  for (const rule of state.valueRules) {
    if (!rule || !rule.to) continue;
    const froms = Array.isArray(rule.from) ? rule.from : String(rule.from || "").split(",");
    if (froms.some((f) => normalizeForMatch(f) === norm)) return rule.to;
  }
  return value;
}

async function copyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch (err) {
    return false;
  }
}

/* ---------- Toast ---------- */
let statusTimer = null;
function showStatus(message, type = "info") {
  const el = $("statusMessage");
  el.textContent = message;
  el.className = `status-message show ${type}`;
  clearTimeout(statusTimer);
  statusTimer = setTimeout(() => el.classList.remove("show"), 2500);
}

/* ================== 2. COLONNES ================== */

// Largeur effective du tableau (nombre max de cellules sur les 200 premières lignes)
function tableWidth() {
  let w = 0;
  const scan = Math.min(state.rows.length, 200);
  for (let i = 0; i < scan; i++) w = Math.max(w, (state.rows[i] || []).length);
  return Math.min(w, 200);
}

// Heuristique : la première ligne ressemble-t-elle à des en-têtes ?
function detectHeaders() {
  if (state.rows.length < 2) return false;
  const first = state.rows[0] || [];
  const nonEmpty = first.filter((c) => String(c ?? "").trim() !== "");
  if (nonEmpty.length < 2) return false;
  const textCells = nonEmpty.filter((c) => {
    const s = String(c).trim();
    return s !== "" && isNaN(Number(s.replace(",", ".")));
  });
  const unique = new Set(nonEmpty.map((c) => String(c).trim().toLowerCase()));
  return textCells.length / nonEmpty.length >= 0.7 && unique.size === nonEmpty.length;
}

function hasHeaders() {
  if (state.headerMode === "yes") return true;
  if (state.headerMode === "no") return false;
  return detectHeaders();
}

function dataStartIdx() { return hasHeaders() ? 1 : 0; }

// Colonnes effectives : [{name, index, letter}]
function getColumns() {
  const width = tableWidth();
  const cols = [];
  if (!state.rows.length) return cols;

  if (hasHeaders()) {
    const seen = {};
    const header = state.rows[0] || [];
    for (let i = 0; i < width; i++) {
      let name = String(header[i] ?? "").trim() || indexToLetter(i);
      if (seen[name]) { seen[name]++; name = `${name} (${seen[name]})`; }
      else seen[name] = 1;
      cols.push({ name, index: i, letter: indexToLetter(i) });
    }
    return cols;
  }

  const model = models.find((m) => m.name === state.modelName);
  if (model) {
    const n = Math.max(width, model.columns.length);
    const seen = {};
    for (let i = 0; i < n; i++) {
      let name = (model.columns[i] || "").trim() || indexToLetter(i);
      if (seen[name]) { seen[name]++; name = `${name} (${seen[name]})`; }
      else seen[name] = 1;
      cols.push({ name, index: i, letter: indexToLetter(i) });
    }
    return cols;
  }

  for (let i = 0; i < width; i++) {
    cols.push({ name: indexToLetter(i), index: i, letter: indexToLetter(i) });
  }
  return cols;
}

function colIndexByName(name) {
  const col = getColumns().find((c) => c.name === name);
  if (col) return col.index;
  const li = letterToIndex(name);
  return li; // -1 si introuvable
}

function getCellByIndex(row, idx) {
  if (idx < 0 || !row) return "";
  const v = row[idx];
  return v === undefined || v === null ? "" : String(v);
}

function setCellByIndex(row, idx, value) {
  if (idx < 0) return;
  while (row.length <= idx) row.push("");
  row[idx] = value;
}

// Clé de persistance du mapping pour le jeu de colonnes courant
function mappingKey() {
  if (!hasHeaders() && state.modelName) return `model:${state.modelName}`;
  const names = getColumns().map((c) => c.name).join("");
  return `cols:${hashString(names)}`;
}

/* ---------- Rendu : chips + aperçu + selects de colonnes ---------- */

function renderColumns() {
  const cols = getColumns();
  const chips = $("columnsChips");
  chips.innerHTML = "";
  cols.slice(0, 60).forEach((c) => {
    const span = document.createElement("span");
    span.className = "chip";
    span.innerHTML = `<span class="chip-letter">${escapeHtml(c.letter)}</span>${escapeHtml(c.name)}`;
    span.title = `${c.name} (colonne ${c.letter})`;
    chips.appendChild(span);
  });
  if (cols.length > 60) {
    const span = document.createElement("span");
    span.className = "chip";
    span.textContent = `… +${cols.length - 60}`;
    chips.appendChild(span);
  }

  const info = $("columnsInfo");
  if (!state.rows.length) {
    info.textContent = "";
  } else if (hasHeaders()) {
    info.textContent = `${cols.length} colonnes — noms lus sur la 1re ligne${state.headerMode === "auto" ? " (détection auto)" : ""}.`;
  } else if (state.modelName) {
    const model = models.find((m) => m.name === state.modelName);
    const diff = model ? tableWidth() - model.columns.length : 0;
    info.textContent = `${cols.length} colonnes — modèle « ${state.modelName} »` +
      (diff > 0 ? ` (⚠ ${diff} colonne(s) de plus que le modèle)` : diff < 0 && tableWidth() > 0 ? ` (données plus courtes que le modèle)` : "") + ".";
  } else {
    info.textContent = `${cols.length} colonnes — sans nom (lettres A, B, C…). Choisis un modèle ou active les en-têtes.`;
  }

  $("modelRow").style.display = hasHeaders() || !state.rows.length ? "none" : "flex";
  renderPreview();
  refreshAllColumnSelects();
  renderRowList();
  renderSelectedRowFields();
  loadMappingForCurrentColumns();
  updateDoneMarkers();
}

function renderPreview() {
  const wrap = $("previewWrap");
  if (!state.rows.length) {
    wrap.innerHTML = '<p class="hint">Charge des données pour voir l\'aperçu.</p>';
    return;
  }
  const cols = getColumns().slice(0, 15);
  const start = dataStartIdx();
  const preview = state.rows.slice(start, start + 5);
  let html = '<table class="preview-table"><thead><tr><th>#</th>';
  html += cols.map((c) => `<th title="${escapeAttr(c.name)}">${escapeHtml(c.name)}</th>`).join("");
  html += "</tr></thead><tbody>";
  preview.forEach((row, i) => {
    html += `<tr><td>${start + i + 1}</td>`;
    html += cols.map((c) => `<td>${escapeHtml(getCellByIndex(row, c.index))}</td>`).join("");
    html += "</tr>";
  });
  html += "</tbody></table>";
  wrap.innerHTML = html;
}

// Remplit un <select> avec les colonnes courantes ; garde la valeur si possible.
// extra = [{value, label}] options supplémentaires en fin de liste.
function fillColumnSelect(sel, current, extra = []) {
  const cols = getColumns();
  sel.innerHTML = "";
  cols.forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c.name;
    opt.textContent = c.name === c.letter ? c.letter : `${c.name} (${c.letter})`;
    sel.appendChild(opt);
  });
  extra.forEach((e) => {
    const opt = document.createElement("option");
    opt.value = e.value;
    opt.textContent = e.label;
    sel.appendChild(opt);
  });
  if (current) {
    const known = cols.some((c) => c.name === current) || extra.some((e) => e.value === current);
    if (!known) {
      const opt = document.createElement("option");
      opt.value = current;
      opt.textContent = `⚠ ${current}`;
      sel.appendChild(opt);
    }
    sel.value = current;
  }
}

function refreshAllColumnSelects() {
  document.querySelectorAll("select[data-colselect]").forEach((sel) => {
    const extra = sel.dataset.colselect === "target" ? [{ value: "__other__", label: "➕ Autre (lettre ou nouvelle colonne)…" }] : [];
    fillColumnSelect(sel, sel.value, extra);
    if (sel.dataset.colselect === "target") {
      const inp = sel.parentElement.querySelector(".new-col-input");
      if (inp) inp.style.display = sel.value === "__other__" ? "block" : "none";
    }
  });
}

/* ---------- Persistance de session ---------- */

let sessionSizeWarned = false;
function persistSession() {
  try {
    const session = {
      rows: state.rows,
      sheetName: state.sheetName,
      headerMode: state.headerMode,
      modelName: state.modelName,
      selectedRowIdx: state.selectedRowIdx,
      originalFileName: state.originalFileName,
      hasOriginalFile: state.hasOriginalFile
    };
    const size = JSON.stringify(session).length;
    if (size > 4_000_000) {
      if (!sessionSizeWarned) {
        sessionSizeWarned = true;
        showStatus("Fichier volumineux : les données ne seront pas conservées à la fermeture du panneau.", "info");
      }
      chrome.storage.local.remove("session");
      return;
    }
    chrome.storage.local.set({ session });
  } catch (e) { /* ignoré */ }
}

/* ================== 3. CHARGEMENT DES DONNÉES ================== */

function setLoadedInfo(label) {
  const n = state.rows.length;
  $("fileInfo").textContent = `${label} : ${n} ligne${n > 1 ? "s" : ""}${hasHeaders() ? " (en-tête incluse)" : ""}.`;
  $("endRow").placeholder = "auto (" + n + ")";
  $("scnEndRow").placeholder = "auto (" + n + ")";
}

function afterDataLoaded(label) {
  state.selectedRowIdx = state.rows.length > dataStartIdx() ? dataStartIdx() : null;
  setLoadedInfo(label);
  $("outputNameInput").value = state.originalFileName;
  autoSuggestModel();
  renderColumns();
  persistSession();
}

// Si pas d'en-têtes et qu'un seul modèle correspond exactement au nombre
// de colonnes, on le propose automatiquement.
function autoSuggestModel() {
  if (hasHeaders() || state.modelName || !state.rows.length) return;
  const width = tableWidth();
  const matching = models.filter((m) => m.columns.length === width);
  if (matching.length === 1) {
    state.modelName = matching[0].name;
    $("modelSelect").value = state.modelName;
    showStatus(`Modèle « ${matching[0].name} » détecté automatiquement (${width} colonnes).`, "success");
  }
}

/* ---------- Fichier Excel / CSV ---------- */

$("fileInput").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  state.originalFileName = file.name.replace(/\.(xlsx|xls|csv)$/i, "") + "_maj.xlsx";

  try {
    if (/\.csv$/i.test(file.name)) {
      const text = await file.text();
      state.workbook = XLSX.read(text, { type: "string" });
      state.originalArrayBuffer = null;   // CSV : pas de mise en forme d'origine
      state.hasOriginalFile = false;
      idbClearOriginal();
    } else {
      const buf = await file.arrayBuffer();
      // On conserve les octets bruts du fichier pour l'export fidèle via ExcelJS
      // (SheetJS ne relit pas les styles). buf est copié car il sera relu plusieurs fois.
      state.originalArrayBuffer = buf.slice(0);
      state.hasOriginalFile = true;
      // Persistance hors session pour survivre à la réouverture du panneau.
      idbSaveOriginal(state.originalArrayBuffer).catch(() => {});
      state.workbook = XLSX.read(buf, { type: "array", cellStyles: true, cellNF: true, sheetStubs: true });
    }
    const sheetSelect = $("sheetSelect");
    sheetSelect.innerHTML = "";
    state.workbook.SheetNames.forEach((name) => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      sheetSelect.appendChild(opt);
    });
    $("sheetRow").style.display = state.workbook.SheetNames.length > 1 ? "flex" : "none";
    loadSheet(state.workbook.SheetNames[0]);
  } catch (err) {
    $("fileInfo").textContent = "Erreur de lecture du fichier : " + err.message;
  }
});

$("sheetSelect").addEventListener("change", () => loadSheet($("sheetSelect").value));

function loadSheet(name) {
  state.sheetName = name;
  const ws = state.workbook.Sheets[name];
  state.rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
  afterDataLoaded(`Feuille « ${name} » chargée`);
}

/* ---------- Coller un tableau ---------- */

document.querySelectorAll('input[name="sourceMode"]').forEach((r) => {
  r.addEventListener("change", () => {
    const mode = document.querySelector('input[name="sourceMode"]:checked').value;
    $("sourceFile").style.display = mode === "file" ? "block" : "none";
    $("sourcePaste").style.display = mode === "paste" ? "block" : "none";
    $("sourceJson").style.display = mode === "json" ? "block" : "none";
  });
});

function parseDelimitedText(text) {
  const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").filter((l) => l.length > 0);
  const delim = text.includes("\t") ? "\t" : ",";
  return lines.map((line) => line.split(delim));
}

$("usePasteBtn").addEventListener("click", () => {
  const text = $("pasteArea").value;
  if (!text.trim()) { showStatus("Colle d'abord un tableau.", "error"); return; }
  state.workbook = null;
  state.originalArrayBuffer = null;
  state.hasOriginalFile = false;
  idbClearOriginal();
  state.sheetName = "Tableau collé";
  state.originalFileName = "resultat.xlsx";
  state.rows = parseDelimitedText(text);
  $("sheetRow").style.display = "none";
  afterDataLoaded("Tableau collé chargé");
});

/* ---------- JSON ---------- */

function jsonToRows(data) {
  if (!Array.isArray(data)) throw new Error("Le JSON doit être un tableau.");
  if (data.length === 0) return [[]];
  if (Array.isArray(data[0])) {
    return data.map((r) => r.map((v) => (v === null || v === undefined ? "" : String(v))));
  }
  const headers = [];
  data.forEach((obj) => {
    if (obj && typeof obj === "object") {
      Object.keys(obj).forEach((k) => { if (!headers.includes(k)) headers.push(k); });
    }
  });
  const out = [headers];
  data.forEach((obj) => {
    out.push(headers.map((h) => (obj && obj[h] !== undefined && obj[h] !== null ? String(obj[h]) : "")));
  });
  return out;
}

function applyJsonText(text) {
  if (!text.trim()) { showStatus("Colle ou choisis d'abord un JSON.", "error"); return; }
  try {
    const data = JSON.parse(text);
    state.rows = jsonToRows(data);
    state.workbook = null;
    state.originalArrayBuffer = null;
    state.hasOriginalFile = false;
    idbClearOriginal();
    state.sheetName = "JSON";
    state.originalFileName = "resultat.xlsx";
    $("sheetRow").style.display = "none";
    afterDataLoaded("JSON chargé");
  } catch (err) {
    showStatus("Erreur JSON : " + err.message, "error");
  }
}

$("useJsonBtn").addEventListener("click", () => applyJsonText($("jsonTextArea").value));

$("jsonFileInput").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const text = await file.text();
  $("jsonTextArea").value = text;
  applyJsonText(text);
});

/* ---------- Mode en-têtes + modèle ---------- */

$("headerModeSelect").addEventListener("change", () => {
  state.headerMode = $("headerModeSelect").value;
  if (state.rows.length) {
    state.selectedRowIdx = state.rows.length > dataStartIdx() ? dataStartIdx() : null;
    setLoadedInfo(`Feuille « ${state.sheetName || "?"} »`);
  }
  renderColumns();
  persistSession();
});

$("modelSelect").addEventListener("change", () => {
  state.modelName = $("modelSelect").value;
  renderColumns();
  persistSession();
});

/* ---------- Modèles de colonnes (CRUD) ---------- */

function persistModels() { chrome.storage.local.set({ models }); }

function renderModelSelect() {
  const sel = $("modelSelect");
  sel.innerHTML = '<option value="">— lettres (A, B, C…) —</option>';
  models.forEach((m) => {
    const opt = document.createElement("option");
    opt.value = m.name;
    opt.textContent = `${m.name} (${m.columns.length} col.)`;
    sel.appendChild(opt);
  });
  sel.value = state.modelName || "";
  if (sel.value !== (state.modelName || "")) { state.modelName = ""; sel.value = ""; }
}

function renderModelsList() {
  const list = $("modelsList");
  list.innerHTML = "";
  if (!models.length) {
    list.innerHTML = '<p class="hint">Aucun modèle.</p>';
    return;
  }
  models.forEach((m) => {
    const div = document.createElement("div");
    div.className = "model-item";
    div.innerHTML = `
      <span class="model-name">${escapeHtml(m.name)}</span>
      <span class="model-count">${m.columns.length} col.</span>
      <button class="btn icon-only" data-edit title="Modifier" type="button"><svg class="icon icon-sm"><use href="#icon-pen"/></svg></button>
      <button class="remove-btn" data-del title="Supprimer" type="button"><svg class="icon icon-sm"><use href="#icon-trash"/></svg></button>
    `;
    div.querySelector("[data-edit]").addEventListener("click", () => {
      $("newModelName").value = m.name;
      $("newModelColumns").value = m.columns.join("\n");
      $("newModelColumns").focus();
    });
    div.querySelector("[data-del]").addEventListener("click", () => {
      models = models.filter((x) => x.name !== m.name);
      if (state.modelName === m.name) state.modelName = "";
      persistModels();
      renderModelsList();
      renderModelSelect();
      renderColumns();
    });
    list.appendChild(div);
  });
}

$("manageModelsBtn").addEventListener("click", () => {
  renderModelsList();
  $("modelsModal").hidden = false;
});
$("closeModelsBtn").addEventListener("click", () => { $("modelsModal").hidden = true; });
$("modelsModal").addEventListener("click", (e) => {
  if (e.target === $("modelsModal")) $("modelsModal").hidden = true;
});

$("addModelBtn").addEventListener("click", () => {
  const name = $("newModelName").value.trim();
  const columns = $("newModelColumns").value.split("\n").map((c) => c.trim()).filter(Boolean);
  if (!name) { showStatus("Donne un nom au modèle.", "error"); return; }
  if (!columns.length) { showStatus("Liste au moins une colonne (une par ligne).", "error"); return; }
  const existing = models.find((m) => m.name === name);
  if (existing) existing.columns = columns;
  else models.push({ name, columns });
  persistModels();
  renderModelsList();
  renderModelSelect();
  state.modelName = name;
  $("modelSelect").value = name;
  renderColumns();
  persistSession();
  showStatus(`Modèle « ${name} » ${existing ? "mis à jour" : "créé"} (${columns.length} colonnes).`, "success");
  $("newModelName").value = "";
  $("newModelColumns").value = "";
});

$("modelFromColumnsBtn").addEventListener("click", () => {
  if (!state.rows.length) { showStatus("Charge d'abord des données.", "error"); return; }
  $("newModelColumns").value = getColumns().map((c) => c.name).join("\n");
  if (!$("newModelName").value) $("newModelName").value = state.sheetName || "";
  $("newModelName").focus();
});

/* ================== 4. SAISIE (remplissage de formulaires) ================== */

/* ---------- Sélection de la ligne active ---------- */

function dataRowIndices() {
  const start = dataStartIdx();
  const out = [];
  for (let i = start; i < state.rows.length; i++) out.push(i);
  return out;
}

function rowSummary(row) {
  const parts = [];
  for (let i = 0; i < row.length && parts.length < 3; i++) {
    const v = String(row[i] ?? "").trim();
    if (v) parts.push(v.length > 28 ? v.slice(0, 28) + "…" : v);
  }
  return parts.join(" — ") || "(ligne vide)";
}

const ROW_PAGE_SIZE = 300;
let rowListPage = 0;

/* Indices des lignes de données correspondant au filtre courant. */
function getFilteredRowIndices() {
  const filter = $("rowFilterInput").value.trim().toLowerCase();
  const indices = dataRowIndices();
  if (!filter) return indices;
  return indices.filter((idx) =>
    (state.rows[idx] || []).join(" ").toLowerCase().includes(filter)
  );
}

/* Page (0-index) sur laquelle se trouve une ligne donnée. */
function pageForRowIndex(idx) {
  const pos = getFilteredRowIndices().indexOf(idx);
  return pos === -1 ? rowListPage : Math.floor(pos / ROW_PAGE_SIZE);
}

function renderRowList() {
  const list = $("rowList");
  list.innerHTML = "";
  if (!state.rows.length) {
    $("selectedRowLabel").textContent = "—";
    renderRowPager(1, 0);
    return;
  }
  const filtered = getFilteredRowIndices();
  const pageCount = Math.max(1, Math.ceil(filtered.length / ROW_PAGE_SIZE));
  rowListPage = Math.min(Math.max(rowListPage, 0), pageCount - 1);
  const startPos = rowListPage * ROW_PAGE_SIZE;
  const pageIndices = filtered.slice(startPos, startPos + ROW_PAGE_SIZE);

  for (const idx of pageIndices) {
    const row = state.rows[idx] || [];
    const div = document.createElement("div");
    div.className = "row-item" + (idx === state.selectedRowIdx ? " selected" : "");
    div.innerHTML = `<span class="row-num">L${idx + 1}</span><span class="row-summary">${escapeHtml(rowSummary(row))}</span>`;
    div.addEventListener("click", () => selectRow(idx));
    list.appendChild(div);
  }
  if (!filtered.length) {
    const p = document.createElement("div");
    p.className = "row-item";
    p.innerHTML = '<span class="row-summary" style="color:var(--text-dim)">Aucune ligne ne correspond au filtre.</span>';
    list.appendChild(p);
  }
  renderRowPager(pageCount, filtered.length);
  $("selectedRowLabel").textContent = state.selectedRowIdx !== null ? `L${state.selectedRowIdx + 1}` : "—";
}

function renderRowPager(pageCount, total) {
  const pager = $("rowPager");
  if (!pager) return;
  pager.innerHTML = "";
  if (pageCount <= 1) return;

  const startNum = rowListPage * ROW_PAGE_SIZE + 1;
  const endNum = Math.min((rowListPage + 1) * ROW_PAGE_SIZE, total);

  const prev = document.createElement("button");
  prev.type = "button";
  prev.className = "btn icon-only row-pager-btn";
  prev.textContent = "‹";
  prev.title = "Page précédente";
  prev.disabled = rowListPage === 0;
  prev.addEventListener("click", () => { rowListPage--; renderRowList(); });

  const info = document.createElement("span");
  info.className = "row-pager-info";
  info.textContent = `${startNum}–${endNum} sur ${total} · page ${rowListPage + 1}/${pageCount}`;

  const next = document.createElement("button");
  next.type = "button";
  next.className = "btn icon-only row-pager-btn";
  next.textContent = "›";
  next.title = "Page suivante";
  next.disabled = rowListPage >= pageCount - 1;
  next.addEventListener("click", () => { rowListPage++; renderRowList(); });

  pager.appendChild(prev);
  pager.appendChild(info);
  pager.appendChild(next);
}

function selectRow(idx) {
  state.selectedRowIdx = idx;
  rowListPage = pageForRowIndex(idx);
  renderRowList();
  renderSelectedRowFields();
  updateFillButtonState();
  updateDoneMarkers();
  persistSession();
}

$("rowFilterInput").addEventListener("input", () => { rowListPage = 0; renderRowList(); });

$("prevRowBtn").addEventListener("click", () => {
  const indices = dataRowIndices();
  if (!indices.length) return;
  const pos = indices.indexOf(state.selectedRowIdx);
  const next = pos <= 0 ? indices[0] : indices[pos - 1];
  selectRow(next);
});

$("nextRowBtn").addEventListener("click", () => {
  const indices = dataRowIndices();
  if (!indices.length) return;
  const pos = indices.indexOf(state.selectedRowIdx);
  const next = pos === -1 ? indices[0] : indices[Math.min(indices.length - 1, pos + 1)];
  selectRow(next);
});

/* ---------- Champs de la ligne (avec copie) ---------- */

function renderSelectedRowFields() {
  const container = $("fieldsList");
  container.innerHTML = "";
  if (state.selectedRowIdx === null || !state.rows.length) {
    container.innerHTML = '<p class="hint">Sélectionne une ligne ci-dessus.</p>';
    return;
  }
  const row = state.rows[state.selectedRowIdx] || [];
  getColumns().forEach((c) => {
    const raw = getCellByIndex(row, c.index);
    const display = String(formatValueForColumn(c.name, raw) ?? "");
    const item = document.createElement("div");
    item.className = "field-item";

    const info = document.createElement("div");
    info.className = "field-info-container";
    const label = document.createElement("div");
    label.className = "field-label";
    label.textContent = c.name;
    const value = document.createElement("div");
    value.className = "field-value" + (display ? "" : " empty");
    value.textContent = display || "(vide)";
    info.appendChild(label);
    info.appendChild(value);

    const btn = document.createElement("button");
    btn.className = "field-copy-btn";
    btn.title = "Copier cette valeur";
    btn.innerHTML = '<svg class="icon icon-sm"><use href="#icon-copy"/></svg>';
    btn.addEventListener("click", async () => {
      const ok = await copyToClipboard(display);
      if (ok) {
        btn.classList.add("copied");
        showStatus(`✓ « ${c.name} » copié !`, "success");
        setTimeout(() => btn.classList.remove("copied"), 1000);
      } else {
        showStatus("Erreur lors de la copie", "error");
      }
    });

    item.appendChild(info);
    item.appendChild(btn);
    container.appendChild(item);
  });
}

/* ---------- Mapping colonnes -> inputs ---------- */

function updateMappingInfo() {
  const n = Object.keys(state.mapping).length;
  $("mappingInfo").textContent = `${n} colonne${n > 1 ? "s" : ""} mappée${n > 1 ? "s" : ""}`;
  updateFillButtonState();
}

function loadMappingForCurrentColumns() {
  if (!state.rows.length) { state.mapping = {}; updateMappingInfo(); return; }
  state.mapping = allMappings[mappingKey()] || {};
  updateMappingInfo();
}

function openMappingModal() {
  if (!state.rows.length) { showStatus("Charge d'abord des données.", "error"); return; }
  const container = $("mappingFields");
  container.innerHTML = "";

  getColumns().forEach((c) => {
    const fieldContainer = document.createElement("div");
    fieldContainer.className = "mapping-field-container";

    const label = document.createElement("label");
    label.className = "mapping-label";
    label.textContent = c.name;
    fieldContainer.appendChild(label);

    let existing = state.mapping[c.name] || [];
    if (typeof existing === "string") existing = existing ? [existing] : [];
    if (!Array.isArray(existing)) existing = [];

    const inputsContainer = document.createElement("div");
    inputsContainer.className = "mapping-inputs-container";
    inputsContainer.dataset.column = c.name;

    if (existing.length > 0) existing.forEach((v) => addMappingInput(inputsContainer, v));
    else addMappingInput(inputsContainer, "");

    fieldContainer.appendChild(inputsContainer);

    const addButton = document.createElement("button");
    addButton.className = "add-mapping-btn";
    addButton.type = "button";
    addButton.innerHTML = '<svg class="icon"><use href="#icon-plus"/></svg> Ajouter un champ';
    addButton.addEventListener("click", () => addMappingInput(inputsContainer, ""));
    fieldContainer.appendChild(addButton);

    container.appendChild(fieldContainer);
  });

  $("mappingModal").hidden = false;
}

function addMappingInput(container, value) {
  const group = document.createElement("div");
  group.className = "mapping-input-group";

  const input = document.createElement("input");
  input.type = "text";
  input.placeholder = "name ou id (ex: body:x:tabc:x:txtNom)";
  input.value = value;

  const pickBtn = document.createElement("button");
  pickBtn.className = "btn pick icon-only";
  pickBtn.type = "button";
  pickBtn.title = "Cliquer sur le champ du site pour récupérer son name";
  pickBtn.innerHTML = '<svg class="icon"><use href="#icon-target"/></svg>';
  pickBtn.addEventListener("click", async () => {
    const picked = await pickTargetOnActiveTab();
    if (!picked) return;
    const identifier = picked.name || picked.id;
    if (identifier) {
      input.value = identifier;
      const kind = picked.name ? "name" : "id";
      showStatus(`✓ ${kind} récupéré : ${identifier}`, "success");
    } else {
      showStatus("Cet élément n'a ni name ni id — impossible de le cibler.", "error");
    }
  });

  const removeBtn = document.createElement("button");
  removeBtn.className = "remove-btn";
  removeBtn.type = "button";
  removeBtn.title = "Supprimer";
  removeBtn.innerHTML = '<svg class="icon icon-sm"><use href="#icon-close"/></svg>';
  removeBtn.addEventListener("click", () => {
    if (container.children.length > 1) group.remove();
    else input.value = "";
  });

  group.appendChild(input);
  group.appendChild(pickBtn);
  group.appendChild(removeBtn);
  container.appendChild(group);
}

function closeMappingModal() { $("mappingModal").hidden = true; }

function saveMappingFromModal() {
  const newMapping = {};
  $("mappingFields").querySelectorAll(".mapping-inputs-container").forEach((container) => {
    const column = container.dataset.column;
    const values = [];
    container.querySelectorAll("input").forEach((input) => {
      const v = input.value.trim();
      if (v) values.push(v);
    });
    if (values.length > 0) newMapping[column] = values;
  });
  state.mapping = newMapping;
  allMappings[mappingKey()] = newMapping;
  chrome.storage.local.set({ allMappings });
  updateMappingInfo();
  updateDoneMarkers();
  closeMappingModal();
  showStatus("Mapping sauvegardé !", "success");
}

$("openMappingBtn").addEventListener("click", openMappingModal);
$("closeMappingBtn").addEventListener("click", closeMappingModal);
$("cancelMappingBtn").addEventListener("click", closeMappingModal);
$("saveMappingBtn").addEventListener("click", saveMappingFromModal);
$("mappingModal").addEventListener("click", (e) => {
  if (e.target === $("mappingModal")) closeMappingModal();
});

/* ---------- Règles de valeurs (UI) ---------- */

function persistValueRules() {
  try { chrome.storage.local.set({ valueRules: state.valueRules }); } catch (e) { /* ignoré */ }
}

function updateValueRulesInfo() {
  const n = (state.valueRules || []).filter(
    (r) => r && r.to && (Array.isArray(r.from) ? r.from.length : String(r.from || "").trim())
  ).length;
  const el = $("valueRulesInfo");
  if (el) el.textContent = `${n} règle${n > 1 ? "s" : ""}`;
}

function addValueRuleRow(rule = { from: [], to: "" }) {
  const fromStr = Array.isArray(rule.from) ? rule.from.join(", ") : String(rule.from || "");
  const div = document.createElement("div");
  div.className = "value-rule-row field-row";
  div.innerHTML = `
    <input type="text" class="vr-from" style="flex:2;min-width:120px" placeholder="valeurs sources (ex: female, femme, f)" value="${escapeAttr(fromStr)}" />
    <span class="vr-arrow" aria-hidden="true">→</span>
    <input type="text" class="vr-to" style="flex:1;min-width:80px" placeholder="valeur cible (ex: Mme)" value="${escapeAttr(rule.to || "")}" />
    <button class="remove-btn vr-remove" type="button" title="Supprimer la règle"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
  `;
  div.querySelector(".vr-remove").addEventListener("click", () => div.remove());
  $("valueRulesList").appendChild(div);
}

function renderValueRulesModal() {
  $("valueRulesList").innerHTML = "";
  const rules = (state.valueRules && state.valueRules.length) ? state.valueRules : [{ from: [], to: "" }];
  rules.forEach((r) => addValueRuleRow(r));
}

function getValueRulesFromDOM() {
  return Array.from($("valueRulesList").querySelectorAll(".value-rule-row"))
    .map((row) => ({
      from: row.querySelector(".vr-from").value.split(",").map((s) => s.trim()).filter(Boolean),
      to: row.querySelector(".vr-to").value.trim()
    }))
    .filter((r) => r.to && r.from.length);
}

function openValueRulesModal() {
  renderValueRulesModal();
  $("valueRulesModal").hidden = false;
}
function closeValueRulesModal() { $("valueRulesModal").hidden = true; }

function saveValueRulesFromModal() {
  state.valueRules = getValueRulesFromDOM();
  persistValueRules();
  updateValueRulesInfo();
  persistWorkingConfig();
  closeValueRulesModal();
  showStatus("Règles de valeurs sauvegardées.", "success");
}

$("openValueRulesBtn").addEventListener("click", openValueRulesModal);
$("closeValueRulesBtn").addEventListener("click", closeValueRulesModal);
$("cancelValueRulesBtn").addEventListener("click", closeValueRulesModal);
$("saveValueRulesBtn").addEventListener("click", saveValueRulesFromModal);
$("addValueRuleBtn").addEventListener("click", () => addValueRuleRow());
$("valueRulesModal").addEventListener("click", (e) => {
  if (e.target === $("valueRulesModal")) closeValueRulesModal();
});

/* ---------- Champs personnalisés ---------- */

function getCustomFieldsFromDOM() {
  const out = [];
  $("customFieldsList").querySelectorAll(".custom-field-row").forEach((row) => {
    out.push({
      name: row.querySelector(".custom-field-name").value.trim(),
      value: row.querySelector(".custom-field-value").value.trim()
    });
  });
  return out;
}

function saveCustomFields() {
  state.customFields = getCustomFieldsFromDOM();
  chrome.storage.local.set({ customFields: state.customFields });
  updateFillButtonState();
}

function renderCustomFields() {
  const list = $("customFieldsList");
  list.innerHTML = "";
  const items = state.customFields.length ? state.customFields : [{ name: "", value: "" }];
  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "custom-field-row";
    row.innerHTML = `
      <input type="text" class="custom-field-name" placeholder="name, id ou classe" value="${escapeAttr(item.name)}" />
      <button class="btn pick icon-only custom-field-pick" type="button" title="Cliquer sur le champ du site pour récupérer son name/id"><svg class="icon"><use href="#icon-target"/></svg></button>
      <input type="text" class="custom-field-value" placeholder="valeur (ex : MG{N° MG})" title="Valeur dynamique : {Nom de colonne} ou {A} est remplacé par la valeur de la ligne active. Ex : MG{N° MG}" value="${escapeAttr(item.value)}" />
      <button class="remove-btn" type="button" title="Supprimer"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
    `;
    row.querySelectorAll("input").forEach((inp) => inp.addEventListener("blur", saveCustomFields));
    row.querySelector(".custom-field-pick").addEventListener("click", async () => {
      const picked = await pickTargetOnActiveTab();
      if (!picked) return;
      const identifier = picked.name || picked.id;
      if (identifier) {
        row.querySelector(".custom-field-name").value = identifier;
        saveCustomFields();
        const kind = picked.name ? "name" : "id";
        showStatus(`✓ ${kind} récupéré : ${identifier}`, "success");
      } else {
        showStatus("Cet élément n'a ni name ni id — saisis une classe manuellement.", "error");
      }
    });
    row.querySelector(".remove-btn").addEventListener("click", () => {
      row.remove();
      saveCustomFields();
      if (!$("customFieldsList").children.length) {
        state.customFields = [];
        renderCustomFields();
      }
    });
    list.appendChild(row);
  });
}

$("addCustomFieldBtn").addEventListener("click", () => {
  state.customFields = getCustomFieldsFromDOM();
  state.customFields.push({ name: "", value: "" });
  renderCustomFields();
});

$("exportCustomFieldsBtn").addEventListener("click", () => {
  const data = getCustomFieldsFromDOM().filter((f) => f.name);
  if (!data.length) { showStatus("Aucun champ personnalisé à exporter.", "error"); return; }
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "nohands-osa-champs-perso.json";
  a.click();
  URL.revokeObjectURL(url);
  showStatus(`${data.length} champ(s) exporté(s) !`, "success");
});

$("importCustomFieldsBtn").addEventListener("click", () => $("importCustomFieldsFile").click());
$("importCustomFieldsFile").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const data = JSON.parse(ev.target.result);
      if (!Array.isArray(data)) { showStatus("Format invalide : tableau JSON attendu.", "error"); return; }
      const valid = data
        .filter((x) => x && typeof x === "object" && typeof x.name === "string")
        .map((x) => ({ name: x.name.trim(), value: String(x.value || "").trim() }));
      if (!valid.length) { showStatus("Aucun champ valide dans le fichier.", "error"); return; }
      state.customFields = valid;
      renderCustomFields();
      saveCustomFields();
      showStatus(`${valid.length} champ(s) importé(s) !`, "success");
    } catch (err) {
      showStatus("Erreur JSON : " + err.message, "error");
    }
  };
  reader.readAsText(file);
  e.target.value = "";
});

/* ---------- Remplissage ---------- */

function updateFillButtonState() {
  const hasRow = state.selectedRowIdx !== null && state.rows.length > 0;
  const hasMapping = Object.keys(state.mapping).length > 0;
  const hasCustom = getCustomFieldsFromDOM().some((f) => f.name);
  $("fillBtn").disabled = !((hasRow && hasMapping) || hasCustom);
}

// Envoie le message de remplissage à une liste d'onglets déjà résolue.
async function sendFillToTabList(targetTabs, message) {
  if (!targetTabs.length) throw new Error("Aucun onglet cible trouvé");

  let totalFilled = 0;
  const totalErrors = [];
  let tabsReached = 0;

  await Promise.all(targetTabs.map(async (tab) => {
    try {
      const response = await chrome.tabs.sendMessage(tab.id, message);
      if (response) {
        tabsReached++;
        totalFilled += response.filledCount || 0;
        if (response.errors) totalErrors.push(...response.errors);
      }
    } catch (msgErr) {
      // Secours : injecter le content script puis réessayer
      try {
        await chrome.scripting.executeScript({
          target: { tabId: tab.id, allFrames: true },
          files: ["content.js"]
        });
        const response = await chrome.tabs.sendMessage(tab.id, message);
        if (response) {
          tabsReached++;
          totalFilled += response.filledCount || 0;
          if (response.errors) totalErrors.push(...response.errors);
        }
      } catch (retryErr) {
        console.warn(`NoHands OSA: onglet ${tab.id} injoignable:`, retryErr.message);
      }
    }
  }));

  return { totalFilled, totalErrors, tabsReached, tabCount: targetTabs.length };
}

/**
 * Remplissage cloisonné sur UN onglet (mode scénario multi-onglets).
 * Chaque onglet traite sa propre ligne : diffuser à tous les onglets du
 * site écraserait les lignes des autres. On inclut malgré tout les popups
 * ouvertes PAR cet onglet (window.open), car certains formulaires — SIGEO
 * notamment — déportent une partie de la saisie dans une popup.
 */
async function sendFillToOneTab(tabId, message) {
  let tab;
  try {
    tab = await chrome.tabs.get(tabId);
  } catch (_) {
    throw new Error("Onglet de travail fermé pendant l'exécution.");
  }

  const url = tab.url || "";
  if (!/^https?:/.test(url)) throw new Error("Page protégée : ouvre d'abord le site cible");

  const origin = new URL(url).origin;
  const sameOrigin = await chrome.tabs.query({ url: origin + "/*" });
  const popups = sameOrigin.filter(
    (t) => t.openerTabId === tabId && /^https?:/.test(t.url || "")
  );

  return sendFillToTabList([tab, ...popups], message);
}

// Envoie le message de remplissage à tous les onglets de la même origine
// que l'onglet actif (popups window.open inclus). Hérité de NoHands.
async function sendFillToAllTabs(message) {
  // Onglet du formulaire (mémorisé via 🎯), et non l'onglet affiché : le
  // remplissage fonctionne donc même quand tu es sur une autre page.
  const targetTabId = await resolveTargetTabId();
  let targetTab;
  try {
    targetTab = await chrome.tabs.get(targetTabId);
  } catch (_) {
    state.targetTabId = null;
    throw new Error("Onglet cible fermé : rouvre le formulaire puis clique 🎯.");
  }

  const targetUrl = targetTab.url || "";
  if (targetUrl.startsWith("chrome://") || targetUrl.startsWith("about:") || targetUrl.startsWith("chrome-extension://")) {
    throw new Error("Page protégée : ouvre d'abord le site cible");
  }

  const origin = new URL(targetUrl).origin;
  const allTabs = await chrome.tabs.query({ url: origin + "/*" });
  const targetTabs = allTabs.filter((t) => {
    const u = t.url || "";
    return u.startsWith("http://") || u.startsWith("https://");
  });
  // L'onglet mémorisé en tête (popups window.open de même origine inclus).
  if (!targetTabs.some((t) => t.id === targetTabId)) targetTabs.unshift(targetTab);

  return sendFillToTabList(targetTabs, message);
}

// Construit data/mapping/customFields pour une ligne donnée.
// Utilisé par le bouton « Remplir le formulaire » et par le scénario de saisie.
function buildFillPayload(rowIdx) {
  const data = {};
  const mapping = {};
  const activeRow = (rowIdx !== null && rowIdx !== undefined && state.rows.length) ? (state.rows[rowIdx] || []) : null;
  if (activeRow) {
    const row = activeRow;
    const colByName = {};
    getColumns().forEach((c) => { colByName[c.name] = c; });
    for (const [colName, inputNames] of Object.entries(state.mapping)) {
      const col = colByName[colName];
      if (!col) continue;
      data[colName] = applyValueRules(formatValueForColumn(colName, getCellByIndex(row, col.index)));
      mapping[colName] = inputNames;
    }
  }
  const customFields = {};
  state.customFields.forEach(({ name, value }) => { if (name) customFields[name] = resolveRowTemplate(value, activeRow); });
  // Contexte de ligne complet (toutes les colonnes) : sert au content script
  // à départager les suggestions des champs à autocomplétion asynchrone
  // (ex. CP tapé → choisit « 70600 - ARGILLIERES » si la ligne contient la ville).
  const rowContext = {};
  if (activeRow) {
    getColumns().forEach((c) => {
      const v = applyValueRules(formatValueForColumn(c.name, getCellByIndex(activeRow, c.index)));
      if (v !== undefined && v !== null && String(v).trim() !== "") rowContext[c.name] = String(v);
    });
  }
  return { data, mapping, customFields, rowContext };
}

$("fillBtn").addEventListener("click", async () => {
  saveCustomFields();

  const { data, mapping, customFields, rowContext } = buildFillPayload(state.selectedRowIdx);

  if (!Object.keys(mapping).length && !Object.keys(customFields).length) {
    showStatus("Aucune donnée exploitable : configure le mapping ou des champs personnalisés.", "error");
    return;
  }

  try {
    const message = {
      action: "fillForm",
      data,
      mapping,
      customFields: Object.keys(customFields).length ? customFields : undefined,
      rowContext
    };
    const { totalFilled, totalErrors, tabsReached, tabCount } = await sendFillToAllTabs(message);

    if (totalFilled > 0) {
      const tabInfo = tabCount > 1 ? ` (${tabsReached} onglet${tabsReached > 1 ? "s" : ""})` : "";
      showStatus(`✓ ${totalFilled} champ${totalFilled > 1 ? "s" : ""} rempli${totalFilled > 1 ? "s" : ""}${tabInfo} !`, "success");
    } else if (totalErrors.length) {
      showStatus(`Erreur : ${totalErrors.slice(0, 3).join(", ")}`, "error");
    } else {
      showStatus("Aucun champ rempli. Vérifie le mapping.", "error");
    }
  } catch (error) {
    showStatus("Erreur : " + error.message, "error");
  }
});

/* ---------- Saisie multi-onglets (une ligne différente par onglet) ---------- */

const mtState = {
  active: false,
  origin: null,
  tabIds: [],              // onglets gérés par le mode
  assignments: new Map(),  // tabId -> index de ligne en cours
  queue: [],               // indices des lignes restantes
  done: 0,
  total: 0,
  lastAdvance: new Map()   // tabId -> timestamp (anti double déclenchement)
};

// Envoie un message à UN onglet, avec injection de secours du content script.
async function sendMessageToTab(tabId, message) {
  try {
    return await chrome.tabs.sendMessage(tabId, message);
  } catch (_) {
    await chrome.scripting.executeScript({
      target: { tabId, allFrames: true },
      files: ["content.js"]
    });
    return await chrome.tabs.sendMessage(tabId, message);
  }
}

function mtRowLabel(rowIdx) { return "Ligne " + (rowIdx + 1); }

function renderMtTabs() {
  const list = $("mtTabsList");
  list.innerHTML = "";
  if (!mtState.active) { $("mtProgress").textContent = ""; return; }
  $("mtProgress").textContent =
    `${mtState.done} / ${mtState.total} traitée(s) — ${mtState.queue.length} en attente`;
  mtState.tabIds.forEach((tabId, i) => {
    const div = document.createElement("div");
    div.className = "mt-tab-item";
    const rowIdx = mtState.assignments.get(tabId);
    const label = document.createElement("span");
    label.className = "mt-tab-label";
    label.textContent = `Onglet ${i + 1} : ` + (rowIdx !== undefined ? mtRowLabel(rowIdx) : "terminé ✓");
    const goBtn = document.createElement("button");
    goBtn.type = "button";
    goBtn.className = "btn sm";
    goBtn.textContent = "Voir";
    goBtn.addEventListener("click", async () => {
      try {
        const t = await chrome.tabs.get(tabId);
        await chrome.tabs.update(tabId, { active: true });
        if (t.windowId != null) await chrome.windows.update(t.windowId, { focused: true });
      } catch (_) { showStatus("Onglet fermé.", "error"); }
    });
    const nextBtn = document.createElement("button");
    nextBtn.type = "button";
    nextBtn.className = "btn sm";
    nextBtn.textContent = "Ligne suivante";
    nextBtn.title = "Marquer la ligne comme validée et remplir la suivante dans cet onglet";
    nextBtn.addEventListener("click", () => { mtAdvanceTab(tabId, true).catch(console.warn); });
    div.append(label, goBtn, nextBtn);
    list.appendChild(div);
  });
}

// Remplit l'onglet avec sa ligne assignée et affiche le badge.
async function mtFillTab(tabId) {
  const rowIdx = mtState.assignments.get(tabId);
  if (rowIdx === undefined) return null;
  const { data, mapping, customFields, rowContext } = buildFillPayload(rowIdx);
  const message = {
    action: "fillForm",
    data,
    mapping,
    customFields: Object.keys(customFields).length ? customFields : undefined,
    rowContext
  };
  try {
    const response = await sendMessageToTab(tabId, message);
    await sendMessageToTab(tabId, { action: "showRowBadge", label: mtRowLabel(rowIdx), state: "active" });
    return response;
  } catch (err) {
    console.warn("NoHands OSA multi-onglets: onglet " + tabId + " injoignable:", err.message);
    return null;
  }
}

// Assigne la prochaine ligne de la file à l'onglet (ou le marque terminé).
async function mtAssignNext(tabId) {
  if (!mtState.active) return;
  if (!mtState.queue.length) {
    mtState.assignments.delete(tabId);
    try { await sendMessageToTab(tabId, { action: "showRowBadge", label: "✓ Terminé", state: "done" }); } catch (_) {}
    renderMtTabs();
    if (!mtState.assignments.size) {
      showStatus(`Saisie multi-onglets terminée : ${mtState.done} ligne(s) traitée(s).`, "success");
      mtStop(false, false);
    }
    return;
  }
  const rowIdx = mtState.queue.shift();
  mtState.assignments.set(tabId, rowIdx);
  await mtFillTab(tabId);
  renderMtTabs();
}

// L'onglet vient d'être validé (rechargement de page, ou clic manuel) :
// la ligne en cours est comptée traitée et l'onglet reçoit la suivante.
async function mtAdvanceTab(tabId, manual = false) {
  if (!mtState.active) return;
  const now = Date.now();
  if (!manual && now - (mtState.lastAdvance.get(tabId) || 0) < 1500) return;
  mtState.lastAdvance.set(tabId, now);
  if (mtState.assignments.has(tabId)) mtState.done++;
  await mtAssignNext(tabId);
}

chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
  if (!mtState.active || !mtState.tabIds.includes(tabId)) return;
  if (changeInfo.status !== "complete") return;
  if (!$("mtAutoNext").checked) {
    // Pas d'avance auto : on ré-affiche juste le badge de la ligne en cours.
    const rowIdx = mtState.assignments.get(tabId);
    if (rowIdx !== undefined) {
      sendMessageToTab(tabId, { action: "showRowBadge", label: mtRowLabel(rowIdx), state: "active" }).catch(() => {});
    }
    return;
  }
  const delay = Math.max(0, parseInt($("mtDelayMs").value, 10) || 0);
  setTimeout(() => { mtAdvanceTab(tabId).catch(console.warn); }, delay);
});

chrome.tabs.onRemoved.addListener((tabId) => {
  if (!mtState.active || !mtState.tabIds.includes(tabId)) return;
  // Onglet fermé : sa ligne non validée retourne en tête de file.
  const rowIdx = mtState.assignments.get(tabId);
  if (rowIdx !== undefined) mtState.queue.unshift(rowIdx);
  mtState.assignments.delete(tabId);
  mtState.tabIds = mtState.tabIds.filter((id) => id !== tabId);
  if (!mtState.tabIds.length) {
    showStatus("Tous les onglets multi-saisie sont fermés — mode arrêté.", "error");
    mtStop(false);
  } else {
    renderMtTabs();
  }
});

async function mtStart() {
  saveCustomFields();
  if (!state.rows.length) { showStatus("Charge d'abord des données.", "error"); return; }
  const hasMapping = Object.keys(state.mapping).length > 0;
  const hasCustom = state.customFields.some((f) => f.name);
  if (!hasMapping && !hasCustom) {
    showStatus("Configure d'abord le mapping (ou des champs personnalisés).", "error");
    return;
  }

  // File de lignes : lignes filtrées, à partir de la ligne active.
  const indices = getFilteredRowIndices();
  if (!indices.length) { showStatus("Aucune ligne à saisir (filtre trop restrictif ?).", "error"); return; }
  const startPos = state.selectedRowIdx !== null ? Math.max(0, indices.indexOf(state.selectedRowIdx)) : 0;
  const queue = indices.slice(startPos);

  // Onglet/site cible (dernier 🎯, sinon onglet actif).
  let targetTab;
  try {
    targetTab = await chrome.tabs.get(await resolveTargetTabId());
  } catch (_) {
    showStatus("Ouvre d'abord le site cible (ou clique 🎯 sur le formulaire).", "error");
    return;
  }
  const targetUrl = targetTab.url || "";
  if (!/^https?:/.test(targetUrl)) {
    showStatus("Page protégée : ouvre d'abord le site cible.", "error");
    return;
  }
  const origin = new URL(targetUrl).origin;
  const wanted = Math.min(Math.max(parseInt($("mtTabCount").value, 10) || 2, 2), 10);

  // Onglets existants du même site (l'onglet cible en tête).
  const existing = (await chrome.tabs.query({ url: origin + "/*" }))
    .filter((t) => /^https?:/.test(t.url || ""));
  if (!existing.some((t) => t.id === targetTab.id)) existing.unshift(targetTab);
  const initialTabIds = existing.slice(0, wanted).map((t) => t.id);

  mtState.active = true;
  mtState.origin = origin;
  mtState.tabIds = initialTabIds.slice();
  mtState.assignments = new Map();
  mtState.queue = queue;
  mtState.done = 0;
  mtState.total = queue.length;
  mtState.lastAdvance = new Map();
  $("mtStartBtn").disabled = true;
  $("mtStopBtn").disabled = false;

  // S'il manque des onglets, on duplique l'onglet du formulaire.
  // Ils seront remplis à la fin de leur chargement (onUpdated).
  const toCreate = Math.min(wanted, queue.length) - mtState.tabIds.length;
  for (let i = 0; i < toCreate; i++) {
    try {
      const dup = await chrome.tabs.duplicate(targetTab.id);
      mtState.tabIds.push(dup.id);
    } catch (err) {
      showStatus("Duplication d'onglet impossible : " + err.message, "error");
      break;
    }
  }

  // Les onglets déjà chargés reçoivent leur ligne immédiatement.
  for (const tabId of mtState.tabIds) {
    let tab;
    try { tab = await chrome.tabs.get(tabId); } catch (_) { continue; }
    if (tab.status === "complete" && !mtState.assignments.has(tabId)) {
      mtState.lastAdvance.set(tabId, Date.now());
      await mtAssignNext(tabId);
    }
  }
  renderMtTabs();
  showStatus(`Saisie multi-onglets : ${queue.length} ligne(s) réparties sur ${mtState.tabIds.length} onglet(s).`, "success");
}

function mtStop(notify = true, hideBadges = true) {
  const tabIds = mtState.tabIds.slice();
  mtState.active = false;
  mtState.tabIds = [];
  mtState.assignments = new Map();
  mtState.queue = [];
  $("mtStartBtn").disabled = false;
  $("mtStopBtn").disabled = true;
  renderMtTabs();
  if (hideBadges) {
    tabIds.forEach((tabId) => {
      sendMessageToTab(tabId, { action: "hideRowBadge" }).catch(() => {});
    });
  }
  if (notify) showStatus("Saisie multi-onglets arrêtée.", "info");
}

$("mtStartBtn").addEventListener("click", () => {
  mtStart().catch((e) => {
    mtStop(false);
    showStatus("Erreur : " + e.message, "error");
  });
});
$("mtStopBtn").addEventListener("click", () => mtStop());

/* ---------- Sélection d'un élément sur la page (🎯) ---------- */

async function getActiveTabId() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab) throw new Error("Aucun onglet actif trouvé.");
  return tab.id;
}

// Onglet du formulaire à remplir. On privilégie l'onglet mémorisé (dernier 🎯),
// afin que le remplissage/l'extraction visent TOUJOURS le formulaire, même si tu
// as basculé sur une autre page. Repli sur l'onglet actif si aucun cible valide.
async function resolveTargetTabId() {
  if (state.targetTabId != null) {
    try {
      const t = await chrome.tabs.get(state.targetTabId);
      const u = t.url || "";
      if (u.startsWith("http://") || u.startsWith("https://")) return state.targetTabId;
    } catch (_) {
      // Onglet fermé entre-temps : on oublie et on retombe sur l'onglet actif.
      state.targetTabId = null;
    }
  }
  const id = await getActiveTabId();
  state.targetTabId = id;
  return id;
}

// Retourne { selector, name } ou null (Échap).
async function pickTargetOnActiveTab() {
  try {
    const tabId = await getActiveTabId();
    // L'utilisateur pointe le formulaire : on mémorise cet onglet comme cible.
    state.targetTabId = tabId;
    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId },
      func: pickElementOnPageInjected
    });
    return result;
  } catch (err) {
    showStatus("Erreur sélection : " + err.message, "error");
    return null;
  }
}

// Fonction injectée : l'utilisateur clique un élément, on renvoie
// son sélecteur CSS et son attribut name. Autonome (aucune référence externe).
function pickElementOnPageInjected() {
  return new Promise((resolve) => {
    const prevCursor = document.documentElement.style.cursor;
    document.documentElement.style.cursor = "crosshair";
    let highlighted = null;

    function computeSelector(el) {
      if (el.id) return "#" + CSS.escape(el.id);
      const parts = [];
      let node = el;
      let depth = 0;
      while (node && node.nodeType === 1 && depth < 6) {
        if (node.id) { parts.unshift("#" + CSS.escape(node.id)); break; }
        let part = node.tagName.toLowerCase();
        if (node.classList && node.classList.length) {
          part += "." + Array.from(node.classList).slice(0, 2).map((c) => CSS.escape(c)).join(".");
        }
        const parent = node.parentElement;
        if (parent) {
          const siblings = Array.from(parent.children).filter((c) => c.tagName === node.tagName);
          if (siblings.length > 1) part += ":nth-of-type(" + (siblings.indexOf(node) + 1) + ")";
        }
        parts.unshift(part);
        node = parent;
        depth++;
      }
      return parts.join(" > ");
    }

    function clearHighlight() {
      if (highlighted) {
        highlighted.style.outline = highlighted.__prevOutline || "";
        highlighted.style.outlineOffset = highlighted.__prevOutlineOffset || "";
      }
    }

    function onMouseOver(e) {
      clearHighlight();
      highlighted = e.target;
      highlighted.__prevOutline = highlighted.style.outline;
      highlighted.__prevOutlineOffset = highlighted.style.outlineOffset;
      highlighted.style.outline = "2px solid #6d5ef0";
      highlighted.style.outlineOffset = "1px";
    }

    function cleanup() {
      document.documentElement.style.cursor = prevCursor;
      clearHighlight();
      document.removeEventListener("mouseover", onMouseOver, true);
      document.removeEventListener("click", onClick, true);
      document.removeEventListener("keydown", onKeyDown, true);
    }

    function onClick(e) {
      e.preventDefault();
      e.stopPropagation();
      const target = e.target;
      const result = {
        selector: computeSelector(target),
        name: target.getAttribute ? (target.getAttribute("name") || null) : null,
        id: target.id || null
      };
      cleanup();
      resolve(result);
    }

    function onKeyDown(e) {
      if (e.key === "Escape") { cleanup(); resolve(null); }
    }

    document.addEventListener("mouseover", onMouseOver, true);
    document.addEventListener("click", onClick, true);
    document.addEventListener("keydown", onKeyDown, true);
  });
}

/* ================== 5. EXTRACTION (boucle recherche -> résultats) ================== */

const OPERATORS = [
  { v: "equals", t: "= égal à" },
  { v: "not_equals", t: "≠ différent de" },
  { v: "contains", t: "contient" },
  { v: "not_contains", t: "ne contient pas" },
  { v: "empty", t: "est vide" },
  { v: "not_empty", t: "n'est pas vide" }
];

/* ---------- Conditions ---------- */

function addConditionRow(col = "", op = "equals", val = "") {
  const div = document.createElement("div");
  div.className = "cond-item";
  const opOptions = OPERATORS.map((o) => `<option value="${o.v}" ${o.v === op ? "selected" : ""}>${o.t}</option>`).join("");
  div.innerHTML = `
    <select class="col-select" data-colselect="plain"></select>
    <select class="op-select">${opOptions}</select>
    <input type="text" class="val-input" placeholder="valeur (ex : MG{N° MG})" title="Valeur dynamique : {Nom de colonne} ou {A} est remplacé par la valeur de la ligne. Ex : MG{N° MG}" value="${escapeAttr(val)}" />
    <button class="remove-btn" title="Supprimer" type="button"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
  `;
  fillColumnSelect(div.querySelector(".col-select"), col);
  const valInput = div.querySelector(".val-input");
  const opSelect = div.querySelector(".op-select");
  const toggleVal = () => {
    valInput.style.display = ["empty", "not_empty"].includes(opSelect.value) ? "none" : "block";
  };
  opSelect.addEventListener("change", toggleVal);
  toggleVal();
  div.querySelector(".remove-btn").addEventListener("click", () => { div.remove(); persistWorkingConfig(); });
  $("conditionsList").appendChild(div);
}

$("addConditionBtn").addEventListener("click", () => addConditionRow());

function getConditions() {
  return Array.from($("conditionsList").querySelectorAll(".cond-item")).map((el) => ({
    col: el.querySelector(".col-select").value,
    op: el.querySelector(".op-select").value,
    val: el.querySelector(".val-input").value
  })).filter((c) => c.col);
}

// Teste une cellule contre un opérateur (partagé : extraction + scénario de saisie).
function cellMatches(cell, op, val) {
  const c = String(cell ?? "");
  const v = String(val ?? "");
  switch (op) {
    case "equals": return c.trim().toLowerCase() === v.trim().toLowerCase();
    case "not_equals": return c.trim().toLowerCase() !== v.trim().toLowerCase();
    case "contains": return c.toLowerCase().includes(v.toLowerCase());
    case "not_contains": return !c.toLowerCase().includes(v.toLowerCase());
    case "empty": return c.trim() === "";
    case "not_empty": return c.trim() !== "";
  }
  return false;
}

function rowMatchesSkipCondition(row, conditions) {
  for (const c of conditions) {
    const idx = colIndexByName(c.col);
    const cell = getCellByIndex(row, idx);
    if (cellMatches(cell, c.op, resolveRowTemplate(c.val, row))) return true; // une condition qui matche => ligne ignorée
  }
  return false;
}

/* ---------- Champs de recherche ---------- */

function addSearchFieldRow(selector = "", col = "") {
  const div = document.createElement("div");
  div.className = "search-field-item";
  div.innerHTML = `
    <input type="text" class="selector-input" placeholder="sélecteur CSS du champ" value="${escapeAttr(selector)}" />
    <button class="btn pick icon-only" data-pick-inline="1" title="Choisir sur la page" type="button"><svg class="icon"><use href="#icon-target"/></svg></button>
    <select class="col-select" data-colselect="plain"></select>
    <button class="remove-btn" title="Supprimer" type="button"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
  `;
  fillColumnSelect(div.querySelector(".col-select"), col);
  div.querySelector(".remove-btn").addEventListener("click", () => { div.remove(); persistWorkingConfig(); });
  div.querySelector("[data-pick-inline]").addEventListener("click", async () => {
    const picked = await pickTargetOnActiveTab();
    if (picked && picked.selector) { div.querySelector(".selector-input").value = picked.selector; persistWorkingConfig(); }
  });
  $("searchFieldsList").appendChild(div);
}

$("addSearchFieldBtn").addEventListener("click", () => addSearchFieldRow());

function getSearchFields() {
  return Array.from($("searchFieldsList").querySelectorAll(".search-field-item")).map((el) => ({
    selector: el.querySelector(".selector-input").value.trim(),
    col: el.querySelector(".col-select").value
  })).filter((f) => f.selector && f.col);
}

/* ---------- Mode de validation ---------- */

document.querySelectorAll('input[name="submitMode"]').forEach((r) => {
  r.addEventListener("change", () => {
    $("submitSelectorRow").style.display =
      document.querySelector('input[name="submitMode"]:checked').value === "click" ? "flex" : "none";
  });
});

// Boutons 🎯 "top-level" (data-target)
document.querySelectorAll(".btn.pick[data-target]").forEach((btn) => {
  btn.addEventListener("click", async () => {
    const picked = await pickTargetOnActiveTab();
    if (picked && picked.selector) $(btn.getAttribute("data-target")).value = picked.selector;
  });
});

/* ---------- Résultats à récupérer (outputs) ---------- */

// o = { mode: "css"|"tableMatch"|"tableCheck", col, selector, rowSelector, matchSourceCol,
//       matchType, matchTdIndex, extractTdIndex, newCol,
//       tcSelector, tcCheck, tcText, tcRegex }
function addOutputRow(o = {}) {
  const mode = o.mode || "css";
  const div = document.createElement("div");
  div.className = "out-item";
  div.innerHTML = `
    <div class="out-item-row1">
      <select class="mode-select">
        <option value="css">Info sur la page (sélecteur CSS)</option>
        <option value="tableMatch">Ligne de tableau (par valeur)</option>
        <option value="tableCheck">Tableau : rempli / contient un texte</option>
      </select>
      <input type="text" class="selector-input" placeholder="sélecteur CSS du résultat" value="${escapeAttr(o.selector || "")}" />
      <button class="btn pick icon-only" data-pick-inline="1" title="Choisir sur la page" type="button"><svg class="icon"><use href="#icon-target"/></svg></button>
      <button class="remove-btn" title="Supprimer" type="button"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
    </div>
    <p class="hint out-item-csshint">Récupère le texte de n'importe quel élément de la page (un titre, un prix, un message…), pas seulement dans un tableau. Utilise le bouton 🎯 pour le choisir directement sur la page.</p>
    <div class="out-item-row2">
      <label>→ écrire dans :</label>
      <select class="col-target-select" data-colselect="target"></select>
      <input type="text" class="new-col-input" placeholder="lettre (C) ou nom de nouvelle colonne" style="display:none" value="${escapeAttr(o.newCol || "")}" />
    </div>
    <div class="out-item-tablematch" ${mode === "tableMatch" ? "" : "hidden"}>
      <p class="hint">Pour lire une valeur dans un tableau de résultats. On cherche la ligne dont une cellule correspond à ta donnée, puis on récupère une autre cellule de <b>cette même ligne</b>.</p>
      <label class="tm-label">1. Lignes du tableau (sélecteur CSS)</label>
      <div class="tm-row">
        <input type="text" class="row-selector-input" placeholder="ex : #dataTable tbody tr" value="${escapeAttr(o.rowSelector || "")}" style="flex:1;min-width:110px" />
        <button class="btn pick icon-only" data-pick-row="1" title="Choisir la ligne du tableau sur la page" type="button"><svg class="icon"><use href="#icon-target"/></svg></button>
      </div>
      <label class="tm-label">2. Comparer avec ma colonne…</label>
      <div class="tm-row">
        <select class="match-col-select" data-colselect="plain"></select>
        <select class="match-type-select">
          <option value="contains">contient</option>
          <option value="exact">= exact</option>
        </select>
      </div>
      <label class="tm-label">3. N° des cellules dans le tableau (1 = 1<sup>re</sup> colonne)</label>
      <div class="tm-row">
        <label>cellule à comparer :</label>
        <input type="number" class="match-td-input" min="1" value="${escapeAttr(o.matchTdIndex || 1)}" />
        <label>cellule à extraire :</label>
        <input type="number" class="extract-td-input" min="1" value="${escapeAttr(o.extractTdIndex || 2)}" />
      </div>
    </div>
    <div class="out-item-tablecheck" ${mode === "tableCheck" ? "" : "hidden"}>
      <p class="hint">Regarde un tableau de la page : d'abord s'il est <b>rempli</b> (au moins une ligne de données), puis, si tu veux, s'il <b>contient un certain texte</b> (ex : DPE, même écrit test_dpe ou DpeDiga). Utilise les messages « trouvé » / « rien trouvé » plus bas pour choisir ce qui est écrit (ex : OUI / NON).</p>
      <label class="tm-label">1. Tableau à examiner (sélecteur CSS)</label>
      <div class="tm-row">
        <input type="text" class="tc-selector-input" placeholder="ex : .PowerGridClass  ou  #dataTable tbody tr" value="${escapeAttr(o.tcSelector || "")}" style="flex:1;min-width:110px" />
        <button class="btn pick icon-only" data-pick-tc="1" title="Choisir le tableau sur la page" type="button"><svg class="icon"><use href="#icon-target"/></svg></button>
      </div>
      <label class="tm-label">2. Que vérifier ?</label>
      <div class="tm-row">
        <select class="tc-check-select">
          <option value="filled">Seulement : est-il rempli ?</option>
          <option value="text">Est-il rempli ET contient un texte ?</option>
        </select>
      </div>
      <div class="tc-text-block" ${o.tcCheck === "text" ? "" : "hidden"}>
        <label class="tm-label">3. Texte(s) à chercher dans le tableau</label>
        <div class="tm-row">
          <input type="text" class="tc-text-input" placeholder="ex : dpe   (plusieurs termes séparés par une virgule = OU)" value="${escapeAttr(o.tcText || "")}" style="flex:1;min-width:110px" />
        </div>
        <label class="checkbox-row"><input type="checkbox" class="tc-regex-check" ${o.tcRegex ? "checked" : ""} /> Utiliser une expression régulière (regex, insensible à la casse)</label>
        <p class="hint">Sans regex : recherche « contient », insensible à la casse. « dpe » trouve DPE, test_dpe, DpeDiga… Plusieurs termes séparés par une virgule = au moins un doit être présent. Tu peux aussi utiliser <code>{Nom de colonne}</code> pour chercher une valeur de la ligne.</p>
      </div>
    </div>
    <div class="out-item-notfound">
      <label class="checkbox-row">
        <input type="checkbox" class="notfound-check" ${o.notFoundEnabled ? "checked" : ""} />
        Si rien n'est trouvé, écrire un message
      </label>
      <input type="text" class="notfound-msg" placeholder="ex : KO" value="${escapeAttr(o.notFoundMsg || "")}" ${o.notFoundEnabled ? "" : "style=display:none"} />
    </div>
    <div class="out-item-found">
      <label class="checkbox-row">
        <input type="checkbox" class="found-check" ${o.foundEnabled ? "checked" : ""} />
        Si un résultat est trouvé, écrire un message à la place
      </label>
      <input type="text" class="found-msg" placeholder="ex : OK" value="${escapeAttr(o.foundMsg || "")}" ${o.foundEnabled ? "" : "style=display:none"} />
    </div>
    <div class="out-item-regex">
      <label class="checkbox-row">
        <input type="checkbox" class="regex-check" ${o.regexEnabled ? "checked" : ""} />
        Formater le résultat (regex)
      </label>
      <div class="regex-body" ${o.regexEnabled ? "" : "hidden"}>
        <div class="regex-row">
          <select class="regex-mode">
            <option value="extract">Extraire (garder ce qui correspond)</option>
            <option value="replace">Remplacer</option>
          </select>
          <label class="checkbox-row regex-flag"><input type="checkbox" class="regex-i" ${o.regexFlagI ? "checked" : ""} /> ignorer la casse</label>
        </div>
        <div class="regex-row">
          <label class="regex-lbl">Motif :</label>
          <input type="text" class="regex-pattern" spellcheck="false" placeholder="ex : M\\.?\\s+([A-ZÀ-Ÿ'\\-]+)" value="${escapeAttr(o.regexPattern || "")}" style="flex:1;min-width:120px" />
        </div>
        <div class="regex-row regex-replace-row" ${o.regexMode === "replace" ? "" : "hidden"}>
          <label class="regex-lbl">Remplacer par :</label>
          <input type="text" class="regex-replace" spellcheck="false" placeholder="ex : $1 (vide = supprimer)" value="${escapeAttr(o.regexReplace || "")}" style="flex:1;min-width:120px" />
        </div>
        <p class="hint">Extraire : le 1<sup>er</sup> groupe entre parenthèses <code>( )</code> est gardé, sinon toute la correspondance. Remplacer : utilise <code>$1</code>, <code>$2</code>… pour réinjecter les groupes.</p>
        <div class="regex-row regex-test-row">
          <label class="regex-lbl">Exemple :</label>
          <input type="text" class="regex-test" spellcheck="false" placeholder="colle une valeur pour tester" style="flex:1;min-width:120px" />
          <span class="regex-arrow">→</span>
          <output class="regex-result"></output>
        </div>
      </div>
    </div>
  `;

  div.querySelector(".mode-select").value = mode;
  div.querySelector(".match-type-select").value = o.matchType || "contains";
  div.querySelector(".regex-mode").value = o.regexMode || "extract";

  const targetSel = div.querySelector(".col-target-select");
  fillColumnSelect(targetSel, o.col || "", [{ value: "__other__", label: "➕ Autre (lettre ou nouvelle colonne)…" }]);
  const newColInput = div.querySelector(".new-col-input");
  if (o.col === "__other__" || (o.newCol && !o.col)) {
    targetSel.value = "__other__";
    newColInput.style.display = "block";
  }
  targetSel.addEventListener("change", () => {
    newColInput.style.display = targetSel.value === "__other__" ? "block" : "none";
  });

  fillColumnSelect(div.querySelector(".match-col-select"), o.matchSourceCol || "");

  const tablematchDiv = div.querySelector(".out-item-tablematch");
  const tablecheckDiv = div.querySelector(".out-item-tablecheck");
  const cssHint = div.querySelector(".out-item-csshint");
  const selectorInput = div.querySelector(".selector-input");
  const pickCssBtn = div.querySelector("[data-pick-inline]");
  const modeSelect = div.querySelector(".mode-select");
  const syncMode = () => {
    const isMatch = modeSelect.value === "tableMatch";
    const isCheck = modeSelect.value === "tableCheck";
    const isCss = !isMatch && !isCheck;
    tablematchDiv.hidden = !isMatch;
    tablecheckDiv.hidden = !isCheck;
    cssHint.style.display = isCss ? "block" : "none";
    selectorInput.style.display = isCss ? "block" : "none";
    pickCssBtn.style.display = isCss ? "" : "none";
  };
  modeSelect.addEventListener("change", () => { syncMode(); persistWorkingConfig(); });
  syncMode();

  // Mode "tableCheck" : sélecteur du tableau, sous-mode rempli/texte, terme, regex
  div.querySelector(".tc-check-select").value = o.tcCheck || "filled";
  const tcCheckSelect = div.querySelector(".tc-check-select");
  const tcTextBlock = div.querySelector(".tc-text-block");
  const tcSelInput = div.querySelector(".tc-selector-input");
  const syncTcMode = () => { tcTextBlock.hidden = tcCheckSelect.value !== "text"; };
  tcCheckSelect.addEventListener("change", () => { syncTcMode(); persistWorkingConfig(); });
  syncTcMode();
  div.querySelector("[data-pick-tc]").addEventListener("click", async () => {
    const picked = await pickTargetOnActiveTab();
    if (picked && picked.selector) { tcSelInput.value = picked.selector; persistWorkingConfig(); }
  });
  div.querySelectorAll(".tc-selector-input, .tc-text-input, .tc-regex-check").forEach((el) => {
    el.addEventListener("change", persistWorkingConfig);
  });

  div.querySelector(".remove-btn").addEventListener("click", () => { div.remove(); persistWorkingConfig(); });
  pickCssBtn.addEventListener("click", async () => {
    const picked = await pickTargetOnActiveTab();
    if (picked && picked.selector) { selectorInput.value = picked.selector; persistWorkingConfig(); }
  });

  // 🎯 Choisir la ligne du tableau directement sur la page
  const rowSelInput = div.querySelector(".row-selector-input");
  div.querySelector("[data-pick-row]").addEventListener("click", async () => {
    const picked = await pickTargetOnActiveTab();
    if (picked && picked.selector) { rowSelInput.value = picked.selector; persistWorkingConfig(); }
  });

  // Case à cocher "message si rien n'est trouvé"
  const notFoundCheck = div.querySelector(".notfound-check");
  const notFoundMsg = div.querySelector(".notfound-msg");
  notFoundCheck.addEventListener("change", () => {
    notFoundMsg.style.display = notFoundCheck.checked ? "block" : "none";
    if (notFoundCheck.checked) notFoundMsg.focus();
  });

  // Case à cocher "message si un résultat est trouvé"
  const foundCheck = div.querySelector(".found-check");
  const foundMsg = div.querySelector(".found-msg");
  foundCheck.addEventListener("change", () => {
    foundMsg.style.display = foundCheck.checked ? "block" : "none";
    if (foundCheck.checked) foundMsg.focus();
  });

  // Bloc "Formater le résultat (regex)"
  const regexCheck = div.querySelector(".regex-check");
  const regexBody = div.querySelector(".regex-body");
  const regexModeSel = div.querySelector(".regex-mode");
  const regexReplaceRow = div.querySelector(".regex-replace-row");
  const regexTestInput = div.querySelector(".regex-test");
  const regexResult = div.querySelector(".regex-result");
  const readRegexConfig = () => ({
    regexEnabled: regexCheck.checked,
    regexMode: regexModeSel.value,
    regexPattern: div.querySelector(".regex-pattern").value,
    regexReplace: div.querySelector(".regex-replace").value,
    regexFlagI: div.querySelector(".regex-i").checked
  });
  const refreshRegexTest = () => {
    const cfg = readRegexConfig();
    const sample = regexTestInput.value;
    if (!cfg.regexEnabled || !cfg.regexPattern.trim() || !sample) {
      regexResult.textContent = "";
      regexResult.classList.remove("regex-err");
      return;
    }
    try {
      new RegExp(cfg.regexPattern); // validation
      regexResult.textContent = applyOutputRegex(sample, cfg);
      regexResult.classList.remove("regex-err");
    } catch (e) {
      regexResult.textContent = "regex invalide";
      regexResult.classList.add("regex-err");
    }
  };
  regexCheck.addEventListener("change", () => {
    regexBody.hidden = !regexCheck.checked;
    refreshRegexTest();
    persistWorkingConfig();
  });
  regexModeSel.addEventListener("change", () => {
    regexReplaceRow.hidden = regexModeSel.value !== "replace";
    refreshRegexTest();
    persistWorkingConfig();
  });
  div.querySelectorAll(".regex-pattern, .regex-replace, .regex-i").forEach((el) => {
    el.addEventListener("input", refreshRegexTest);
    el.addEventListener("change", persistWorkingConfig);
  });
  regexTestInput.addEventListener("input", refreshRegexTest);

  $("outputsList").appendChild(div);
}

$("addOutputBtn").addEventListener("click", () => addOutputRow());

function getOutputs() {
  return Array.from($("outputsList").querySelectorAll(".out-item")).map((el) => {
    const mode = el.querySelector(".mode-select").value;
    let col = el.querySelector(".col-target-select").value;
    let newCol = "";
    if (col === "__other__") {
      newCol = el.querySelector(".new-col-input").value.trim();
      col = newCol;
    }
    const notFoundEnabled = el.querySelector(".notfound-check").checked;
    const notFoundMsg = el.querySelector(".notfound-msg").value;
    const foundEnabled = el.querySelector(".found-check").checked;
    const foundMsg = el.querySelector(".found-msg").value;
    const regexEnabled = el.querySelector(".regex-check").checked;
    const regexMode = el.querySelector(".regex-mode").value;
    const regexPattern = el.querySelector(".regex-pattern").value;
    const regexReplace = el.querySelector(".regex-replace").value;
    const regexFlagI = el.querySelector(".regex-i").checked;
    const base = { mode, col, newCol, notFoundEnabled, notFoundMsg, foundEnabled, foundMsg, regexEnabled, regexMode, regexPattern, regexReplace, regexFlagI };
    if (mode === "tableMatch") {
      return {
        ...base,
        rowSelector: el.querySelector(".row-selector-input").value.trim(),
        matchSourceCol: el.querySelector(".match-col-select").value,
        matchType: el.querySelector(".match-type-select").value,
        matchTdIndex: parseInt(el.querySelector(".match-td-input").value, 10) || 1,
        extractTdIndex: parseInt(el.querySelector(".extract-td-input").value, 10) || 1
      };
    }
    if (mode === "tableCheck") {
      return {
        ...base,
        tcSelector: el.querySelector(".tc-selector-input").value.trim(),
        tcCheck: el.querySelector(".tc-check-select").value,
        tcText: el.querySelector(".tc-text-input").value.trim(),
        tcRegex: el.querySelector(".tc-regex-check").checked
      };
    }
    return { ...base, selector: el.querySelector(".selector-input").value.trim() };
  }).filter((o) => {
    if (!o.col) return false;
    if (o.mode === "tableMatch") return Boolean(o.rowSelector && o.matchSourceCol);
    if (o.mode === "tableCheck") return Boolean(o.tcSelector && (o.tcCheck !== "text" || o.tcText));
    return Boolean(o.selector);
  });
}

/* ---------- Fonction injectée : action sur une ligne ---------- */
// (Portée telle quelle depuis OSA — autonome, aucune référence externe.)
function performRowActionInjected(config) {
  return new Promise((resolve) => {
    try {
      const fields = config.searchFields.map((f) => ({
        selector: f.selector,
        value: f.value,
        el: document.querySelector(f.selector)
      }));
      const missingFields = fields.filter((f) => !f.el).map((f) => f.selector);
      if (missingFields.length) {
        return resolve({ ok: false, error: "Champ(s) de recherche introuvable(s) : " + missingFields.join(", ") });
      }

      // Aucun champ de recherche (mode navigation seule) : on lit directement
      // les résultats après le délai d'attente, sans remplir ni valider.
      if (!fields.length) {
        setTimeout(readResults, config.waitMs || 0);
        return;
      }

      function setElementValue(el, val) {
        if ("value" in el) {
          el.focus();
          el.value = val;
          el.dispatchEvent(new Event("input", { bubbles: true }));
          el.dispatchEvent(new Event("change", { bubbles: true }));
        } else {
          el.innerText = val;
        }
      }

      function textOf(el) {
        if (!el) return "";
        if ("value" in el && el.tagName !== "DIV") return el.value;
        return (el.innerText || el.textContent || "").trim();
      }

      function readTableMatch(out) {
        const trs = document.querySelectorAll(out.rowSelector);
        if (!trs.length) return { found: false, value: "", reason: "noRows" };
        const needle = (out.matchValue || "").trim().toLowerCase();
        for (const tr of trs) {
          const cells = tr.querySelectorAll("td");
          const matchCell = cells[out.matchTdIndex - 1];
          if (!matchCell) continue;
          const cellText = matchCell.textContent.trim().toLowerCase();
          const isMatch = out.matchType === "exact" ? cellText === needle : cellText.includes(needle);
          if (isMatch) {
            const extractCell = cells[out.extractTdIndex - 1];
            return { found: true, value: extractCell ? extractCell.textContent.trim() : "" };
          }
        }
        return { found: false, value: "", reason: "noMatch" };
      }

      // Vérifie si un tableau est rempli, et éventuellement s'il contient un texte.
      // Robuste : ignore les lignes d'en-tête, gère un sélecteur qui vise le
      // tableau, le tbody ou directement les lignes.
      function readTableCheck(out) {
        let nodes;
        try { nodes = document.querySelectorAll(out.tcSelector); }
        catch (e) { return { found: false, filled: false, reason: "badSelector" }; }
        if (!nodes.length) return { found: false, filled: false, reason: "noNode" };

        // Récupère les lignes candidates (tr) ou, à défaut, les nœuds eux-mêmes.
        const candidates = [];
        nodes.forEach((n) => {
          const tag = n.tagName ? n.tagName.toLowerCase() : "";
          if (tag === "tr") {
            candidates.push(n);
          } else if (n.querySelectorAll) {
            const trs = n.querySelectorAll("tr");
            if (trs.length) trs.forEach((tr) => candidates.push(tr));
            else candidates.push(n);
          } else {
            candidates.push(n);
          }
        });

        // Ne garde que les lignes de données non vides (on écarte les en-têtes).
        const meaningful = candidates.filter((n) => {
          const tag = n.tagName ? n.tagName.toLowerCase() : "";
          if (tag === "tr") {
            const cls = (n.className || "").toLowerCase();
            if (cls.includes("header")) return false;          // ex : PowerGridHeaderClass
            if (!n.querySelector("td")) return false;           // ligne d'en-têtes (th) seule
          }
          return (n.innerText || n.textContent || "").trim() !== "";
        });

        const filled = meaningful.length > 0;
        if (out.tcCheck !== "text") return { found: filled, filled };

        // Recherche de texte : le tableau doit d'abord être rempli.
        if (!filled) return { found: false, filled: false, reason: "empty" };
        const haystack = meaningful.map((n) => (n.innerText || n.textContent || "")).join("\n");
        const raw = (out.tcText || "").trim();
        if (!raw) return { found: false, filled, reason: "noTerm" };

        let matched = false;
        if (out.tcRegex) {
          try { matched = new RegExp(raw, "i").test(haystack); }
          catch (e) { return { found: false, filled, reason: "badRegex" }; }
        } else {
          const hay = haystack.toLowerCase();
          matched = raw.split(",").map((t) => t.trim().toLowerCase()).filter(Boolean)
            .some((t) => hay.includes(t));
        }
        return { found: matched, filled };
      }

      function readResults() {
        const values = [];
        const notFound = [];
        for (const out of config.outputs) {
          const fallback = out.notFoundEnabled ? (out.notFoundMsg || "") : "";
          if (out.mode === "tableMatch") {
            const { found, value, reason } = readTableMatch(out);
            if (found) {
              // Si un message perso "trouvé" est défini, il remplace la valeur extraite.
              values.push(out.foundEnabled ? (out.foundMsg || "") : value);
            } else {
              values.push(fallback);
              // Si un message perso est défini, on ne signale plus d'erreur "introuvable".
              if (!out.notFoundEnabled) {
                notFound.push(reason === "noRows"
                  ? 'aucune ligne trouvée pour le sélecteur "' + out.rowSelector + '" (vérifie ce sélecteur ou augmente le délai d\'attente)'
                  : 'aucune ligne où la cellule n°' + out.matchTdIndex + ' correspond à "' + out.matchValue + '"');
              }
            }
          } else if (out.mode === "tableCheck") {
            const r = readTableCheck(out);
            // Erreurs de configuration : on le signale comme "introuvable".
            if (r.reason === "badSelector" || r.reason === "badRegex") {
              values.push(out.notFoundEnabled ? (out.notFoundMsg || "") : "");
              notFound.push(r.reason === "badRegex"
                ? 'expression régulière invalide : "' + out.tcText + '"'
                : 'sélecteur de tableau invalide : "' + out.tcSelector + '"');
            } else if (r.found) {
              // Trouvé (rempli, ou rempli + texte présent) : message perso ou "OUI".
              values.push(out.foundEnabled ? (out.foundMsg || "") : "OUI");
            } else {
              // Pas trouvé (vide, ou texte absent) : c'est un résultat normal, pas une erreur.
              values.push(out.notFoundEnabled ? (out.notFoundMsg || "") : "NON");
            }
          } else {
            const el = document.querySelector(out.selector);
            if (!el) {
              values.push(fallback);
              if (!out.notFoundEnabled) notFound.push(out.selector);
              continue;
            }
            values.push(out.foundEnabled ? (out.foundMsg || "") : textOf(el));
          }
        }
        resolve({ ok: true, values, notFound });
      }

      fields.forEach((f) => setElementValue(f.el, f.value));

      setTimeout(() => {
        try {
          if (config.submitMode === "click") {
            const btn = document.querySelector(config.submitSelector);
            if (!btn) return resolve({ ok: false, error: "Bouton de validation introuvable (" + config.submitSelector + ")" });
            btn.click();
          } else {
            const lastEl = fields[fields.length - 1].el;
            lastEl.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", code: "Enter", keyCode: 13, which: 13, bubbles: true }));
            lastEl.dispatchEvent(new KeyboardEvent("keyup", { key: "Enter", code: "Enter", keyCode: 13, which: 13, bubbles: true }));
            if (lastEl.form) {
              try { lastEl.form.requestSubmit ? lastEl.form.requestSubmit() : lastEl.form.submit(); } catch (e) {}
            }
          }
          setTimeout(readResults, config.waitMs || 0);
        } catch (e) {
          resolve({ ok: false, error: String(e) });
        }
      }, 50);
    } catch (e) {
      resolve({ ok: false, error: String(e) });
    }
  });
}

/* ---------- Navigation : ouvrir une URL construite depuis la ligne ---------- */

$("navEnabled").addEventListener("change", () => {
  $("navOptions").style.display = $("navEnabled").checked ? "block" : "none";
});

// Remplace les {Nom de colonne} (ou {A}) du modèle d'URL par les valeurs de la ligne.
// Retourne { url, missing: [colonnes introuvables], values: [valeurs insérées] }.
function buildNavUrl(template, row) {
  const missing = [];
  const values = [];
  const url = String(template || "").replace(/\{([^{}]+)\}/g, (_, name) => {
    const idx = colIndexByName(name.trim());
    if (idx < 0) { missing.push(name.trim()); return ""; }
    const v = getCellByIndex(row, idx).trim();
    values.push(v);
    return encodeURIComponent(v);
  });
  return { url, missing, values };
}

// Résout les variables {Nom de colonne} (ou {A}) d'une chaîne avec les valeurs
// de la ligne active. Sert à rendre dynamiques les valeurs de condition et les
// champs personnalisés (ex : « MG{N° MG} » → « MG » + valeur de la colonne).
// Contrairement à buildNavUrl, la valeur n'est PAS encodée pour l'URL, et une
// variable introuvable est laissée telle quelle pour rester visible.
function resolveRowTemplate(str, row) {
  if (str === null || str === undefined) return str;
  if (!row) return String(str);
  return String(str).replace(/\{([^{}]+)\}/g, (whole, name) => {
    const idx = colIndexByName(name.trim());
    if (idx < 0) return whole;
    return getCellByIndex(row, idx);
  });
}

// Navigue l'onglet vers l'URL. Si waitForLoad, attend la fin du chargement
// (status "complete") avec timeout ; en cas de timeout on continue quand même
// (timedOut: true) car la page est souvent déjà exploitable.
function navigateTabAndWait(tabId, url, waitForLoad, timeoutMs = 15000) {
  return new Promise((resolve) => {
    let settled = false;
    const finish = (res) => {
      if (settled) return;
      settled = true;
      try { chrome.tabs.onUpdated.removeListener(onUpdated); } catch (_) {}
      resolve(res);
    };
    function onUpdated(id, info) {
      if (id === tabId && info.status === "complete") finish({ ok: true });
    }
    if (waitForLoad) {
      chrome.tabs.onUpdated.addListener(onUpdated);
      setTimeout(() => finish({ ok: true, timedOut: true }), Math.max(500, timeoutMs || 0));
    }
    chrome.tabs.update(tabId, { url }, () => {
      if (chrome.runtime.lastError) return finish({ ok: false, error: chrome.runtime.lastError.message });
      if (!waitForLoad) finish({ ok: true });
    });
  });
}

/* ---------- Log & progression ---------- */

function logLine(text, cls = "") {
  const line = document.createElement("div");
  if (cls) line.className = cls;
  line.textContent = text;
  $("log").appendChild(line);
  $("log").scrollTop = $("log").scrollHeight;
}

function setProgress(current, total) {
  $("progressFill").style.width = total ? Math.round((current / total) * 100) + "%" : "0%";
  $("progressText").textContent = `${current} / ${total}`;
}

// Affiche le temps écoulé et une estimation du temps restant pendant l'exécution.
function updateRunTiming(done, total, runStart) {
  const el = $("progressTiming");
  if (!el) return;
  const elapsed = Date.now() - runStart;
  let txt = `Écoulé ${formatDuration(elapsed)}`;
  if (done > 0 && done < total) {
    const eta = (elapsed / done) * (total - done);
    txt += ` · reste ~${formatDuration(eta)}`;
  }
  el.textContent = txt;
}

// Estime le temps d'une extraction à partir de la configuration courante.
// Renvoie { total, perRowMs, totalMs } ou null si aucune donnée.
function estimateExtractionMs() {
  if (!state.rows.length) return null;
  const waitMs = parseInt($("waitMs").value, 10) || 0;
  const rowDelayMs = parseInt($("rowDelayMs").value, 10) || 0;
  const navEnabled = $("navEnabled").checked;
  const navExtraWaitMs = parseInt($("navExtraWaitMs").value, 10) || 0;
  const startRowInput = parseInt($("startRow").value, 10) || (dataStartIdx() + 1);
  const endRowInput = $("endRow").value.trim();
  const startIdx = Math.max(0, startRowInput - 1);
  const endIdx = Math.min(state.rows.length - 1, endRowInput ? parseInt(endRowInput, 10) - 1 : state.rows.length - 1);
  const total = Math.max(0, endIdx - startIdx + 1);
  const SCRIPT_OVERHEAD = 500;         // exécution du script de recherche/lecture
  const NAV_LOAD = navEnabled ? 2000 : 0; // chargement moyen estimé de la page
  const perRowMs = waitMs + rowDelayMs + navExtraWaitMs + SCRIPT_OVERHEAD + NAV_LOAD;
  return { total, perRowMs, totalMs: total * perRowMs };
}

// Met à jour la ligne d'estimation affichée sous les boutons d'exécution.
function updateExtractEstimate() {
  const el = $("extractEstimate");
  if (!el || isRunning) return;
  const est = estimateExtractionMs();
  if (!est || !est.total) { el.textContent = ""; return; }
  el.textContent = `Estimation : ~${formatDuration(est.totalMs)} pour ${est.total} ligne${est.total > 1 ? "s" : ""} (~${formatDuration(est.perRowMs)}/ligne).`;
}

/* ---------- Boucle principale ---------- */

$("startBtn").addEventListener("click", runAutomation);
$("stopBtn").addEventListener("click", () => { stopRequested = true; });

// Résout l'index de colonne cible d'un output ; crée la colonne si besoin.
function resolveOutputTarget(ref, createdCols) {
  let idx = colIndexByName(ref);
  if (idx >= 0) return idx;
  if (createdCols[ref] !== undefined) return createdCols[ref];
  // Nouvelle colonne : ajoutée à droite du tableau
  idx = tableWidth();
  Object.values(createdCols).forEach((v) => { idx = Math.max(idx, v + 1); });
  createdCols[ref] = idx;
  if (hasHeaders()) {
    setCellByIndex(state.rows[0], idx, ref);
  }
  return idx;
}

async function runAutomation() {
  if (!state.rows.length) { logLine("Charge d'abord des données (onglet Données).", "err"); return; }

  const conditions = getConditions();
  const outputs = getOutputs();
  const searchFields = getSearchFields();
  const submitMode = document.querySelector('input[name="submitMode"]:checked').value;
  const submitSelector = $("submitSelector").value.trim();
  const waitMs = parseInt($("waitMs").value, 10) || 0;
  const rowDelayMs = parseInt($("rowDelayMs").value, 10) || 0;
  const navEnabled = $("navEnabled").checked;
  const navUrlTemplate = $("navUrlTemplate").value.trim();
  const navWaitLoad = $("navWaitLoad").checked;
  const navWaitTimeout = parseInt($("navWaitTimeout").value, 10) || 15000;
  const navExtraWaitMs = parseInt($("navExtraWaitMs").value, 10) || 0;

  if (navEnabled && !navUrlTemplate) { logLine("Navigation activée : indique l'URL à ouvrir.", "err"); return; }
  if (!searchFields.length && !navEnabled) { logLine("Ajoute au moins un champ de recherche (ou active la navigation).", "err"); return; }
  if (!outputs.length) { logLine("Ajoute au moins un résultat à récupérer.", "err"); return; }

  // Vérifie en amont que les colonnes citées dans l'URL existent.
  if (navEnabled) {
    const { missing } = buildNavUrl(navUrlTemplate, state.rows[dataStartIdx()] || []);
    if (missing.length) {
      logLine("Colonne(s) introuvable(s) dans l'URL de navigation : " + missing.join(", "), "err");
      return;
    }
  }

  // Résolution des indices de colonnes en amont
  const searchResolved = searchFields.map((f) => ({ ...f, colIdx: colIndexByName(f.col) }));
  const badSearch = searchResolved.filter((f) => f.colIdx < 0);
  if (badSearch.length) {
    logLine("Colonne(s) introuvable(s) : " + badSearch.map((f) => f.col).join(", "), "err");
    return;
  }
  const createdCols = {};
  const outputsResolved = outputs.map((o) => ({
    ...o,
    targetIdx: resolveOutputTarget(o.col, createdCols),
    matchSourceIdx: o.mode === "tableMatch" ? colIndexByName(o.matchSourceCol) : -1
  }));
  const badMatch = outputsResolved.filter((o) => o.mode === "tableMatch" && o.matchSourceIdx < 0);
  if (badMatch.length) {
    logLine("Colonne de comparaison introuvable : " + badMatch.map((o) => o.matchSourceCol).join(", "), "err");
    return;
  }

  const startRowInput = parseInt($("startRow").value, 10) || (dataStartIdx() + 1);
  const endRowInput = $("endRow").value.trim();
  const startIdx = Math.max(0, startRowInput - 1);
  const endIdx = endRowInput ? parseInt(endRowInput, 10) - 1 : state.rows.length - 1;

  let tabId;
  try { tabId = await resolveTargetTabId(); }
  catch (err) { logLine(err.message, "err"); return; }

  isRunning = true;
  hasStartedRun = true;
  stopRequested = false;
  $("startBtn").disabled = true;
  $("stopBtn").disabled = false;
  $("log").innerHTML = "";
  runLog = [];
  lastRunOutputs = outputsResolved;
  const total = Math.max(0, endIdx - startIdx + 1);
  let done = 0;
  const runStart = Date.now();
  runDurationMs = 0;
  setProgress(0, total);
  $("extractEstimate").textContent = "";
  updateRunTiming(0, total, runStart);
  updateDoneMarkers();

  for (let idx = startIdx; idx <= endIdx; idx++) {
    if (stopRequested) { logLine("Arrêté par l'utilisateur.", "skip"); break; }
    const row = state.rows[idx] || [];
    const rowNum = idx + 1;

    if (rowMatchesSkipCondition(row, conditions)) {
      logLine(`Ligne ${rowNum} : ignorée (condition).`, "skip");
      runLog.push({ row: rowNum, search: "", values: [], status: "skip", note: "Condition" });
      done++; setProgress(done, total); updateRunTiming(done, total, runStart);
      continue;
    }

    const searchFieldValues = searchResolved.map((f) => ({ selector: f.selector, value: getCellByIndex(row, f.colIdx) }));
    let searchLabel = searchFieldValues.map((f) => f.value).filter((v) => v.trim()).join(" / ");
    if (searchResolved.length && !searchFieldValues.some((f) => f.value.trim())) {
      logLine(`Ligne ${rowNum} : ignorée (valeur(s) de recherche vide(s)).`, "skip");
      runLog.push({ row: rowNum, search: "", values: [], status: "skip", note: "Valeur vide" });
      done++; setProgress(done, total); updateRunTiming(done, total, runStart);
      continue;
    }

    // Navigation : ouvrir l'URL construite pour cette ligne avant la recherche.
    if (navEnabled) {
      const { url, values } = buildNavUrl(navUrlTemplate, row);
      if (values.length && values.every((v) => !v)) {
        logLine(`Ligne ${rowNum} : ignorée (valeur de navigation vide).`, "skip");
        runLog.push({ row: rowNum, search: searchLabel, values: [], status: "skip", note: "Valeur de navigation vide" });
        done++; setProgress(done, total); updateRunTiming(done, total, runStart);
        continue;
      }
      if (!searchLabel) searchLabel = values.filter(Boolean).join(" / ");
      const nav = await navigateTabAndWait(tabId, url, navWaitLoad, navWaitTimeout);
      if (!nav.ok) {
        logLine(`Ligne ${rowNum} : erreur navigation — ${nav.error}`, "err");
        runLog.push({ row: rowNum, search: searchLabel, values: [], status: "err", note: "Navigation : " + nav.error });
        done++; setProgress(done, total); updateRunTiming(done, total, runStart);
        continue;
      }
      if (nav.timedOut) logLine(`Ligne ${rowNum} : page toujours en chargement après ${navWaitTimeout} ms — on continue.`, "skip");
      if (navExtraWaitMs > 0) await sleep(navExtraWaitMs);
      if (stopRequested) { logLine("Arrêté par l'utilisateur.", "skip"); break; }
    }

    try {
      const [{ result }] = await chrome.scripting.executeScript({
        target: { tabId },
        func: performRowActionInjected,
        args: [{
          searchFields: searchFieldValues,
          submitMode,
          submitSelector,
          waitMs,
          outputs: outputsResolved.map((o) => o.mode === "tableMatch" ? {
            mode: "tableMatch",
            rowSelector: o.rowSelector,
            matchType: o.matchType,
            matchTdIndex: o.matchTdIndex,
            extractTdIndex: o.extractTdIndex,
            matchValue: getCellByIndex(row, o.matchSourceIdx),
            notFoundEnabled: o.notFoundEnabled,
            notFoundMsg: o.notFoundMsg,
            foundEnabled: o.foundEnabled,
            foundMsg: o.foundMsg
          } : o.mode === "tableCheck" ? {
            mode: "tableCheck",
            tcSelector: o.tcSelector,
            tcCheck: o.tcCheck,
            tcText: resolveRowTemplate(o.tcText, row),
            tcRegex: o.tcRegex,
            notFoundEnabled: o.notFoundEnabled,
            notFoundMsg: o.notFoundMsg,
            foundEnabled: o.foundEnabled,
            foundMsg: o.foundMsg
          } : {
            mode: "css",
            selector: o.selector,
            notFoundEnabled: o.notFoundEnabled,
            notFoundMsg: o.notFoundMsg,
            foundEnabled: o.foundEnabled,
            foundMsg: o.foundMsg
          })
        }]
      });

      if (!result || !result.ok) {
        const msg = result ? result.error : "pas de réponse";
        logLine(`Ligne ${rowNum} : erreur — ${msg}`, "err");
        runLog.push({ row: rowNum, search: searchLabel, values: [], status: "err", note: msg });
      } else {
        outputsResolved.forEach((o, i) => {
          let val = result.values[i];
          // Ne pas reformater le message "non trouvé" éventuel.
          const isNotFoundMsg = o.notFoundEnabled && val === (o.notFoundMsg || "");
          const isFoundMsg = o.foundEnabled && val === (o.foundMsg || "");
          if (!isNotFoundMsg && !isFoundMsg) val = applyOutputRegex(val, o);
          setCellByIndex(row, o.targetIdx, val);
        });
        state.rows[idx] = row;
        const missingSelectors = result.notFound || [];
        const allValuesEmpty = result.values.every((v) => !String(v || "").trim());
        if (missingSelectors.length) {
          logLine(`Ligne ${rowNum} (${searchLabel}) : sélecteur introuvable (${missingSelectors.join(", ")}).`, "err");
          runLog.push({ row: rowNum, search: searchLabel, values: result.values, status: "err", note: "Sélecteur introuvable" });
        } else if (allValuesEmpty) {
          logLine(`Ligne ${rowNum} (${searchLabel}) : aucun résultat trouvé.`, "skip");
          runLog.push({ row: rowNum, search: searchLabel, values: result.values, status: "skip", note: "Aucun résultat" });
        } else {
          logLine(`Ligne ${rowNum} (${searchLabel}) : OK`, "ok");
          runLog.push({ row: rowNum, search: searchLabel, values: result.values, status: "ok" });
        }
      }
    } catch (err) {
      logLine(`Ligne ${rowNum} : erreur script — ${err.message}`, "err");
      runLog.push({ row: rowNum, search: searchLabel, values: [], status: "err", note: err.message });
    }

    done++; setProgress(done, total); updateRunTiming(done, total, runStart);
    if (rowDelayMs > 0 && idx < endIdx && !stopRequested) await sleep(rowDelayMs);
  }

  runDurationMs = Date.now() - runStart;
  isRunning = false;
  $("startBtn").disabled = false;
  $("stopBtn").disabled = true;
  $("progressTiming").textContent = `Terminé en ${formatDuration(runDurationMs)}`;
  logLine(`Terminé en ${formatDuration(runDurationMs)}.`, "ok");
  renderResultSummary();
  updateExtractEstimate();
  renderPreview();
  persistSession();
  updateDoneMarkers();
}

/* ---------- Récapitulatif ---------- */

function renderResultSummary() {
  const container = $("resultSummary");
  if (!runLog.length) {
    container.innerHTML = '<p class="hint">Aucune exécution récente.</p>';
    return;
  }
  const headers = lastRunOutputs.map((o) => o.col);
  const statusLabels = { ok: "OK", err: "Erreur", skip: "Ignoré" };
  const counts = runLog.reduce((a, e) => { a[e.status] = (a[e.status] || 0) + 1; return a; }, {});
  const n = runLog.length;
  let html = '<div class="run-recap">';
  if (runDurationMs > 0) {
    html += `<p class="recap-time"><strong>Temps passé :</strong> ${formatDuration(runDurationMs)} pour ${n} ligne${n > 1 ? "s" : ""} · ~${formatDuration(runDurationMs / Math.max(1, n))}/ligne</p>`;
  }
  html += `<p class="recap-counts"><span class="status-ok">${counts.ok || 0} OK</span> · <span class="status-err">${counts.err || 0} erreur(s)</span> · <span class="status-skip">${counts.skip || 0} ignorée(s)</span></p>`;
  html += '</div>';
  html += '<div class="summary-table-wrap"><table class="summary-table"><thead><tr><th>Ligne</th><th>Recherche</th>';
  html += headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("");
  html += "<th>Statut</th></tr></thead><tbody>";
  runLog.forEach((entry) => {
    const statusLabel = statusLabels[entry.status] || entry.status;
    const note = entry.note ? ` — ${escapeHtml(entry.note)}` : "";
    html += `<tr class="status-${entry.status}"><td>${entry.row}</td><td>${escapeHtml(entry.search)}</td>`;
    html += lastRunOutputs.map((_, i) => `<td>${escapeHtml(entry.values[i] || "")}</td>`).join("");
    html += `<td class="status-cell">${statusLabel}${note}</td></tr>`;
  });
  html += "</tbody></table></div>";
  container.innerHTML = html;
}

/* ================== 5bis. SCÉNARIO DE SAISIE ================== */
// Enchaînement d'étapes exécutées sur l'onglet cible (mémorisé via 🎯) :
//   remplir / cliquer / attendre / condition.
// Pensé pour les saisies réparties sur plusieurs panels d'une même page
// (onglets internes, sections dépliables, popups ASP.NET…).

let scnRunning = false;
let scnStopRequested = false;

const SCN_PAGE_OPERATORS = [
  { v: "exists", t: "existe / est visible" },
  { v: "not_exists", t: "n'existe pas" },
  { v: "contains", t: "contient le texte" },
  { v: "not_contains", t: "ne contient pas" },
  { v: "equals", t: "= égal à" }
];

function scnLog(text, cls = "") {
  const line = document.createElement("div");
  if (cls) line.className = cls;
  line.textContent = text;
  $("scnLog").appendChild(line);
  $("scnLog").scrollTop = $("scnLog").scrollHeight;
}

function scnSetProgress(current, total) {
  $("scnProgressFill").style.width = total ? Math.round((current / total) * 100) + "%" : "0%";
  $("scnProgressText").textContent = `${current} / ${total}`;
}

// Affiche le temps écoulé et le temps restant estimé pendant la boucle de saisie.
function scnUpdateTiming(done, total, runStart) {
  const el = $("scnProgressTiming");
  if (!el) return;
  const elapsed = Date.now() - runStart;
  let txt = `Écoulé ${formatDuration(elapsed)}`;
  if (done > 0 && done < total) {
    const eta = (elapsed / done) * (total - done);
    txt += ` · reste ~${formatDuration(eta)}`;
  }
  el.textContent = txt;
}

// Estimation indicative de la durée d'une étape de scénario.
function estimateScenarioStepMs(s) {
  switch (s.type) {
    case "fill": return 700;
    case "goto": return s.gotoWait !== false ? Math.min(2500, parseInt(s.gotoTimeout, 10) || 15000) : 400;
    case "click": return 700;
    case "wait":
      if (s.waitMode === "delay") return parseInt(s.waitMs, 10) || 0;
      return 1200; // attente d'un sélecteur : moyenne estimée
    case "cond": return 250;
    case "pdfcheck": case "pdfwrite": return 50; // local, quasi instantané
    case "sigeo": return (s.sigeoNav === false ? 300 : 2500) + 4500; // nav + remplissage + résolution ville + postback
    // Le nombre de lots n'est connu qu'à l'exécution : on table sur 3.
    case "batchedit": {
      const dl = s.batchWaitMode === "start" || s.batchWaitMode === "complete";
      const parLot = dl
        ? 2500 + (parseInt(s.batchDlSettleMs, 10) || 300)   // génération + détection
        : (parseInt(s.batchWaitMs, 10) || 3000);
      return (parLot + 600) * 3;
    }
  }
  return 300;
}

// Met à jour la ligne d'estimation de la boucle de saisie.
function updateScenarioEstimate() {
  const el = $("scnEstimate");
  if (!el || scnRunning) return;
  const steps = getScenarioSteps();
  if (!steps.length || !state.rows.length) { el.textContent = ""; return; }
  const startRowInput = parseInt($("scnStartRow").value, 10) || (dataStartIdx() + 1);
  const endRowInput = $("scnEndRow").value.trim();
  const startIdx = Math.max(0, startRowInput - 1);
  const endIdx = Math.min(state.rows.length - 1, endRowInput ? parseInt(endRowInput, 10) - 1 : state.rows.length - 1);
  const total = Math.max(0, endIdx - startIdx + 1);
  if (!total) { el.textContent = ""; return; }
  const rowDelayMs = parseInt($("scnRowDelayMs").value, 10) || 0;
  const perRowMs = steps.reduce((sum, s) => sum + estimateScenarioStepMs(s), 0) + rowDelayMs;
  el.textContent = `Estimation boucle : ~${formatDuration(perRowMs * total)} pour ${total} ligne${total > 1 ? "s" : ""} (~${formatDuration(perRowMs)}/ligne).`;
}

// Pause interruptible par le bouton Arrêter.
async function scnSleep(ms) {
  const end = Date.now() + Math.max(0, ms | 0);
  while (Date.now() < end && !scnStopRequested) {
    await sleep(Math.min(100, end - Date.now()));
  }
}

/* ---------- Construction des étapes (UI) ---------- */

function addScenarioStep(step = {}) {
  const div = document.createElement("div");
  div.className = "scn-step";
  div.dataset.type = step.type || "fill";

  const excelOpOptions = OPERATORS.map((o) => `<option value="${o.v}">${o.t}</option>`).join("");
  const pageOpOptions = SCN_PAGE_OPERATORS.map((o) => `<option value="${o.v}">${o.t}</option>`).join("");

  div.innerHTML = `
    <div class="scn-head">
      <button class="btn icon-only scn-drag" title="Glisser pour réordonner" type="button" aria-label="Déplacer l'étape"><svg class="icon icon-sm"><use href="#icon-grip"/></svg></button>
      <span class="scn-num"></span>
      <select class="scn-type" title="Type d'étape">
        <option value="fill">Remplir les champs</option>
        <option value="goto">Ouvrir une URL</option>
        <option value="click">Cliquer sur un élément</option>
        <option value="wait">Attendre</option>
        <option value="cond">Condition (si… alors…)</option>
        <option value="pdfcheck">PDF β : vérifier un champ</option>
        <option value="pdfwrite">PDF β : écrire un champ</option>
        <option value="sigeo">SIGEO : saisir une adresse</option>
        <option value="batchedit">Éditer par lots (cocher N → cliquer)</option>
      </select>
      <button class="btn icon-only scn-move-up" title="Monter" type="button"><svg class="icon icon-sm"><use href="#icon-arrow-up"/></svg></button>
      <button class="btn icon-only scn-move-down" title="Descendre" type="button"><svg class="icon icon-sm"><use href="#icon-arrow-down"/></svg></button>
      <button class="remove-btn" title="Supprimer l'étape" type="button"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
    </div>
    <div class="scn-body">
      <p class="hint scn-only-fill">Remplit tous les champs mappés trouvés sur la page (+ champs personnalisés). Les champs absents — par ex. sur un autre panel — sont simplement ignorés.</p>

      <div class="scn-row scn-only-goto">
        <input type="text" class="scn-goto-url" placeholder="ex : http://sigeo.evoriel.net/…/quick_access?var={N° MG}&metier=ger;;2,27" value="${escapeAttr(step.gotoUrl || "")}" />
      </div>
      <p class="hint scn-only-goto">Navigue l'onglet cible vers cette URL. <code>{Nom de colonne}</code> (ou <code>{A}</code>) est remplacé par la valeur de la ligne active.</p>
      <div class="scn-row scn-only-goto">
        <label class="checkbox-row"><input type="checkbox" class="scn-goto-wait" ${step.gotoWait === false ? "" : "checked"} /> attendre le chargement</label>
        <label>timeout (ms) :</label>
        <input type="number" class="scn-goto-timeout" min="500" step="500" value="${escapeAttr(step.gotoTimeout ?? 15000)}" />
      </div>

      <div class="scn-row scn-only-click">
        <input type="text" class="scn-click-selector" placeholder="sélecteur CSS de l'élément / bouton" value="${escapeAttr(step.clickSelector || "")}" />
        <button class="btn pick icon-only scn-pick-click" title="Choisir sur la page" type="button"><svg class="icon"><use href="#icon-target"/></svg></button>
      </div>
      <div class="scn-row scn-only-click">
        <label>attendre l'élément max (ms) :</label>
        <input type="number" class="scn-click-timeout" min="0" step="100" value="${escapeAttr(step.clickTimeout ?? 5000)}" />
      </div>

      <div class="scn-row scn-only-wait">
        <select class="scn-wait-mode">
          <option value="delay">un délai fixe</option>
          <option value="appear">qu'un élément apparaisse</option>
          <option value="gone">qu'un élément disparaisse</option>
        </select>
        <span class="scn-inline scn-wait-delay-wrap">
          <input type="number" class="scn-wait-ms" min="0" step="100" value="${escapeAttr(step.waitMs ?? 1000)}" />
          <label>ms</label>
        </span>
      </div>
      <div class="scn-row scn-only-wait scn-wait-elt-wrap">
        <input type="text" class="scn-wait-selector" placeholder="sélecteur CSS de l'élément" value="${escapeAttr(step.waitSelector || "")}" />
        <button class="btn pick icon-only scn-pick-wait" title="Choisir sur la page" type="button"><svg class="icon"><use href="#icon-target"/></svg></button>
        <label>timeout (ms) :</label>
        <input type="number" class="scn-wait-timeout" min="100" step="100" value="${escapeAttr(step.waitTimeout ?? 10000)}" />
      </div>

      <div class="scn-row scn-only-cond">
        <label>Si</label>
        <select class="scn-cond-source">
          <option value="excel">une colonne Excel</option>
          <option value="page">un élément de la page</option>
        </select>
      </div>
      <div class="scn-row scn-only-cond scn-cond-excel-wrap">
        <select class="scn-cond-col" data-colselect="plain"></select>
        <select class="scn-cond-op">${excelOpOptions}</select>
        <input type="text" class="scn-cond-val" placeholder="valeur (ex : MG{N° MG})" title="Valeur dynamique : {Nom de colonne} ou {A} est remplacé par la valeur de la ligne active. Ex : MG{N° MG}" value="${escapeAttr(step.condVal || "")}" />
      </div>
      <div class="scn-row scn-only-cond scn-cond-page-wrap">
        <input type="text" class="scn-cond-selector" placeholder="sélecteur CSS" value="${escapeAttr(step.condSelector || "")}" />
        <button class="btn pick icon-only scn-pick-cond" title="Choisir sur la page" type="button"><svg class="icon"><use href="#icon-target"/></svg></button>
        <select class="scn-cond-pageop">${pageOpOptions}</select>
        <input type="text" class="scn-cond-pageval" placeholder="texte (ex : MG{N° MG})" title="Valeur dynamique : {Nom de colonne} ou {A} est remplacé par la valeur de la ligne active." value="${escapeAttr(step.condPageVal || "")}" />
      </div>
      <div class="scn-row scn-only-cond">
        <label>alors</label>
        <select class="scn-cond-action">
          <option value="skip">sauter les étapes suivantes</option>
          <option value="stop">arrêter le scénario</option>
        </select>
        <span class="scn-inline scn-cond-skip-wrap">
          <input type="number" class="scn-cond-skip" min="1" value="${escapeAttr(step.condSkip ?? 1)}" />
          <label>étape(s)</label>
        </span>
      </div>
      <p class="hint scn-only-cond">Valeur dynamique : <code>{Nom de colonne}</code> (ou <code>{A}</code>) est remplacé par la valeur de la ligne active. Ex : <code>MG{N° MG}</code>.</p>

      <div class="scn-row scn-only-pdf">
        <label>Document</label>
        <select class="scn-pdf-docmode">
          <option value="active">PDF sélectionné dans la Toolbox</option>
          <option value="match">retrouver par nom de fichier</option>
        </select>
        <input type="text" class="scn-pdf-docmatch" placeholder="le nom contient… (ex : {N° DPE})" title="Valeur dynamique : {Nom de colonne} ou {A}. Le PDF dont le nom de fichier contient cette valeur est utilisé." value="${escapeAttr(step.pdfDocMatch || "")}" />
      </div>
      <div class="scn-row scn-only-pdf">
        <label>Champ</label>
        <select class="scn-pdf-field"></select>
      </div>
      <div class="scn-row scn-only-pdfcheck">
        <select class="scn-pdf-op">${excelOpOptions}</select>
        <input type="text" class="scn-pdf-val" placeholder="valeur attendue (ex : {SIRET})" title="Valeur dynamique : {Nom de colonne} ou {A} est remplacé par la valeur de la ligne active." value="${escapeAttr(step.pdfVal || "")}" />
      </div>
      <div class="scn-row scn-only-pdfcheck">
        <label>si écart</label>
        <select class="scn-pdf-missaction">
          <option value="error">erreur (interrompt la ligne)</option>
          <option value="warn">avertir et continuer</option>
          <option value="skip">sauter les étapes suivantes</option>
          <option value="stop">arrêter le scénario</option>
        </select>
        <span class="scn-inline scn-pdf-skip-wrap">
          <input type="number" class="scn-pdf-skip" min="1" value="${escapeAttr(step.pdfMissSkip ?? 1)}" />
          <label>étape(s)</label>
        </span>
      </div>
      <p class="hint scn-only-pdfcheck">« égal à / contient » : au moins une valeur du champ PDF correspond. Charge les PDF dans l'onglet <strong>Toolbox</strong>.</p>
      <div class="scn-row scn-only-pdfwrite">
        <label>vers la colonne</label>
        <select class="scn-pdf-targetcol" data-colselect="target"></select>
        <input type="text" class="new-col-input scn-pdf-newcol" placeholder="lettre (C) ou nom de nouvelle colonne" style="display:none" value="${escapeAttr(step.pdfNewCol || "")}" />
      </div>
      <p class="hint scn-only-pdfwrite">Écrit la valeur du champ PDF (valeurs multiples séparées par « | ») dans la colonne, sur la ligne active.</p>

      <p class="hint scn-only-sigeo">Saisit une adresse dans SIGEO (formulaire <code>address_manage</code>) : pays, CP, ville (code commune résolu via le sélecteur de la page), voie, n° de voie, puis « Enreg. et fermer ». <code>{Nom de colonne}</code> (ou <code>{A}</code>) est remplacé par la valeur de la ligne active. Nécessite une session SIGEO déjà connectée dans l'onglet cible.</p>
      <div class="scn-row scn-only-sigeo">
        <label>ID adresse :</label>
        <input type="text" class="scn-sigeo-addressid" placeholder="ex : {ID adresse}" value="${escapeAttr(step.sigeoAddressId ?? "")}" />
        <label>table :</label>
        <input type="text" class="scn-sigeo-table" title="Paramètre d'URL « table » de la fiche adresse" value="${escapeAttr(step.sigeoTable ?? "body_x_tabc_x_desc_tab_x_desc_x_address_addressHolder")}" />
      </div>
      <div class="scn-row scn-only-sigeo">
        <label class="checkbox-row" title="Construit l'URL popup.aspx/…/address_manage/{ID} et navigue avant de remplir. Décocher si la fiche adresse est déjà ouverte dans l'onglet cible."><input type="checkbox" class="scn-sigeo-nav" ${step.sigeoNav === false ? "" : "checked"} /> ouvrir la fiche adresse</label>
        <input type="text" class="scn-sigeo-baseurl" title="Base du site SIGEO" value="${escapeAttr(step.sigeoBaseUrl ?? "http://sigeo.veille.evoriel.net")}" />
      </div>
      <div class="scn-row scn-only-sigeo">
        <label>pays :</label>
        <input type="text" class="scn-sigeo-pays" title="Code ISO2 minuscule (ex : fr)" style="max-width:70px" value="${escapeAttr(step.sigeoPays ?? "fr")}" />
        <label>CP :</label>
        <input type="text" class="scn-sigeo-cp" placeholder="ex : {CP}" value="${escapeAttr(step.sigeoCp ?? "")}" />
        <label>ville :</label>
        <input type="text" class="scn-sigeo-ville" placeholder="ex : {Ville}" value="${escapeAttr(step.sigeoVille ?? "")}" />
      </div>
      <div class="scn-row scn-only-sigeo">
        <label>voie :</label>
        <input type="text" class="scn-sigeo-voie" placeholder="ex : {Voie}" value="${escapeAttr(step.sigeoVoie ?? "")}" />
        <label>n° voie :</label>
        <input type="text" class="scn-sigeo-numvoie" placeholder="libellé, ex : {N°}" title="Libellé affiché dans la liste « numéro de voie » (match strict)" value="${escapeAttr(step.sigeoNumVoie ?? "")}" />
      </div>
      <div class="scn-row scn-only-sigeo">
        <label>BP :</label>
        <input type="text" class="scn-sigeo-bp" placeholder="(optionnel)" value="${escapeAttr(step.sigeoBp ?? "")}" />
        <label>complément :</label>
        <input type="text" class="scn-sigeo-complt" placeholder="(optionnel)" value="${escapeAttr(step.sigeoComplt ?? "")}" />
        <label>compl. nom :</label>
        <input type="text" class="scn-sigeo-compnom" placeholder="(optionnel)" value="${escapeAttr(step.sigeoCompNom ?? "")}" />
      </div>
      <div class="scn-row scn-only-sigeo">
        <label class="checkbox-row" title="Remplit et contrôle tout, mais ne clique PAS sur « Enreg. et fermer ». Idéal pour valider un lot avant écriture réelle."><input type="checkbox" class="scn-sigeo-dryrun" ${step.sigeoDryRun === false ? "" : "checked"} /> simulation (dryRun) — ne pas enregistrer</label>
      </div>
      <div class="scn-row scn-only-sigeo">
        <label>timeout (ms) :</label>
        <input type="number" class="scn-sigeo-timeout" min="2000" step="500" value="${escapeAttr(step.sigeoTimeout ?? 15000)}" />
        <label>résultat → colonne :</label>
        <input type="text" class="new-col-input scn-sigeo-resultcol" placeholder="ex : Résultat SIGEO (vide = ne pas écrire)" title="Écrit OK / SIMULATION OK / ERREUR + détail dans cette colonne (créée si besoin), sur la ligne active" value="${escapeAttr(step.sigeoResultCol ?? "")}" />
      </div>
      <p class="hint scn-only-batchedit">Coche les cases par paquets de N, clique un bouton entre chaque paquet, et recommence jusqu'à épuisement. Prévu pour les pages qui limitent le nombre d'éléments traitables en une fois : édition de feuille de présence (22 clés, lots de 10 → 3 fichiers téléchargés).</p>
      <div class="scn-row scn-only-batchedit">
        <label>Tableau :</label>
        <input type="text" class="scn-batch-scope" placeholder="id du tableau ou sélecteur CSS (vide = toute la page)" value="${escapeAttr(step.batchScope || "")}" />
        <button class="btn pick icon-only scn-pick-batch-scope" title="Choisir le tableau sur la page" type="button"><svg class="icon"><use href="#icon-target"/></svg></button>
      </div>
      <div class="scn-row scn-only-batchedit">
        <label>Bouton à cliquer :</label>
        <input type="text" class="scn-batch-button" placeholder="sélecteur CSS du bouton (ex : Editer)" value="${escapeAttr(step.batchButton || "")}" />
        <button class="btn pick icon-only scn-pick-batch-button" title="Choisir le bouton sur la page" type="button"><svg class="icon"><use href="#icon-target"/></svg></button>
      </div>
      <div class="scn-row scn-only-batchedit">
        <label>Taille de lot :</label>
        <input type="number" class="scn-batch-size" min="1" step="1" value="${escapeAttr(step.batchSize ?? 10)}" />
        <label>après le clic, attendre :</label>
        <select class="scn-batch-waitmode">
          <option value="delay">un délai fixe</option>
          <option value="start">le début du téléchargement</option>
          <option value="complete">la fin du téléchargement</option>
        </select>
      </div>
      <div class="scn-row scn-only-batchedit scn-batch-delay-wrap">
        <label>délai (ms) :</label>
        <input type="number" class="scn-batch-wait" min="0" step="100" value="${escapeAttr(step.batchWaitMs ?? 3000)}" />
      </div>
      <div class="scn-row scn-only-batchedit scn-batch-dl-wrap">
        <label>abandon après (ms) :</label>
        <input type="number" class="scn-batch-dltimeout" min="1000" step="500" title="Durée maximale d'attente du téléchargement avant de continuer malgré tout" value="${escapeAttr(step.batchDlTimeoutMs ?? 30000)}" />
        <label>puis pause (ms) :</label>
        <input type="number" class="scn-batch-dlsettle" min="0" step="100" title="Petit répit une fois le téléchargement détecté, avant le lot suivant" value="${escapeAttr(step.batchDlSettleMs ?? 300)}" />
      </div>
      <div class="scn-row scn-only-batchedit">
        <label>Filtre :</label>
        <input type="text" class="scn-batch-filter" placeholder="optionnel : texte du libellé, ou /regex/i" value="${escapeAttr(step.batchFilter || "")}" />
        <label>lots max :</label>
        <input type="number" class="scn-batch-max" min="1" step="1" title="Garde-fou : nombre maximum de lots avant arrêt" value="${escapeAttr(step.batchMaxRounds ?? 50)}" />
      </div>
      <p class="hint scn-only-batchedit">« Début du téléchargement » enchaîne dès que le navigateur commence à recevoir le fichier ; « fin » attend qu'il soit complet — plus sûr si le serveur est lent. Ces deux modes demandent l'accès aux téléchargements (Chrome l'invite à la sélection). En délai fixe, prévois large : la génération + le téléchargement. Le filtre permet d'exclure certaines lignes (ex. une clé dont le total est nul).</p>

      <p class="hint scn-only-sigeo">La simulation est activée par défaut : décoche-la pour enregistrer réellement. Le ViewState est géré par le navigateur (aucune requête forgée).</p>
    </div>
  `;


  // Valeurs initiales des selects
  div.querySelector(".scn-type").value = step.type || "fill";
  div.querySelector(".scn-wait-mode").value = step.waitMode || "delay";
  div.querySelector(".scn-cond-source").value = step.condSource || "excel";
  div.querySelector(".scn-cond-op").value = step.condOp || "equals";
  div.querySelector(".scn-cond-pageop").value = step.condPageOp || "exists";
  div.querySelector(".scn-cond-action").value = step.condAction || "skip";
  fillColumnSelect(div.querySelector(".scn-cond-col"), step.condCol || "");
  div.querySelector(".scn-pdf-docmode").value = step.pdfDocMode || "active";
  tbFillFieldSelect(div.querySelector(".scn-pdf-field"), step.pdfField || "");
  div.querySelector(".scn-pdf-op").value = step.pdfOp || "equals";
  div.querySelector(".scn-pdf-missaction").value = step.pdfMissAction || "error";
  fillColumnSelect(div.querySelector(".scn-pdf-targetcol"), step.pdfTargetCol || "",
    [{ value: "__other__", label: "➕ Autre (lettre ou nouvelle colonne)…" }]);

  // Affichage conditionnel interne à l'étape
  const typeSelect = div.querySelector(".scn-type");
  const waitMode = div.querySelector(".scn-wait-mode");
  const waitDelayWrap = div.querySelector(".scn-wait-delay-wrap");
  const waitEltWrap = div.querySelector(".scn-wait-elt-wrap");
  const condSource = div.querySelector(".scn-cond-source");
  const condExcelWrap = div.querySelector(".scn-cond-excel-wrap");
  const condPageWrap = div.querySelector(".scn-cond-page-wrap");
  const condOp = div.querySelector(".scn-cond-op");
  const condVal = div.querySelector(".scn-cond-val");
  const condPageOp = div.querySelector(".scn-cond-pageop");
  const condPageVal = div.querySelector(".scn-cond-pageval");
  const condAction = div.querySelector(".scn-cond-action");
  const condSkipWrap = div.querySelector(".scn-cond-skip-wrap");
  const pdfDocMode = div.querySelector(".scn-pdf-docmode");
  const pdfDocMatch = div.querySelector(".scn-pdf-docmatch");
  const pdfOp = div.querySelector(".scn-pdf-op");
  const pdfVal = div.querySelector(".scn-pdf-val");
  const pdfMissAction = div.querySelector(".scn-pdf-missaction");
  const pdfSkipWrap = div.querySelector(".scn-pdf-skip-wrap");
  const pdfTargetCol = div.querySelector(".scn-pdf-targetcol");
  const pdfNewCol = div.querySelector(".scn-pdf-newcol");

  function syncStepUI() {
    div.dataset.type = typeSelect.value;
    const delayMode = waitMode.value === "delay";
    waitDelayWrap.hidden = !delayMode;
    waitEltWrap.hidden = delayMode;
    condExcelWrap.hidden = condSource.value !== "excel";
    condPageWrap.hidden = condSource.value !== "page";
    condVal.hidden = ["empty", "not_empty"].includes(condOp.value);
    condPageVal.hidden = ["exists", "not_exists"].includes(condPageOp.value);
    condSkipWrap.hidden = condAction.value !== "skip";
    pdfDocMatch.hidden = pdfDocMode.value !== "match";
    pdfVal.hidden = ["empty", "not_empty"].includes(pdfOp.value);
    pdfSkipWrap.hidden = pdfMissAction.value !== "skip";
    pdfNewCol.style.display = pdfTargetCol.value === "__other__" ? "block" : "none";
  }
  [typeSelect, waitMode, condSource, condOp, condPageOp, condAction,
    pdfDocMode, pdfOp, pdfMissAction, pdfTargetCol]
    .forEach((sel) => sel.addEventListener("change", syncStepUI));
  syncStepUI();

  // Réordonner / supprimer
  div.querySelector(".scn-move-up").addEventListener("click", () => {
    const prev = div.previousElementSibling;
    if (prev) { div.parentElement.insertBefore(div, prev); persistWorkingConfig(); }
  });
  div.querySelector(".scn-move-down").addEventListener("click", () => {
    const next = div.nextElementSibling;
    if (next) { div.parentElement.insertBefore(next, div); persistWorkingConfig(); }
  });
  div.querySelector(".remove-btn").addEventListener("click", () => {
    div.remove();
    persistWorkingConfig();
  });

  // Réordonner par glisser-déposer (drag & drop) via la poignée
  const dragHandle = div.querySelector(".scn-drag");
  dragHandle.addEventListener("mousedown", () => { div.draggable = true; });
  dragHandle.addEventListener("mouseup", () => { div.draggable = false; });
  div.addEventListener("dragstart", (e) => {
    div.classList.add("scn-dragging");
    e.dataTransfer.effectAllowed = "move";
    try { e.dataTransfer.setData("text/plain", ""); } catch (_) {}
  });
  div.addEventListener("dragend", () => {
    div.classList.remove("scn-dragging");
    div.draggable = false;
    div.parentElement.querySelectorAll(".scn-drop-before, .scn-drop-after")
      .forEach((el) => el.classList.remove("scn-drop-before", "scn-drop-after"));
    persistWorkingConfig();
  });

  // Boutons 🎯
  const wirePick = (btnSel, inputSel) => {
    div.querySelector(btnSel).addEventListener("click", async () => {
      const picked = await pickTargetOnActiveTab();
      if (picked && picked.selector) {
        div.querySelector(inputSel).value = picked.selector;
        persistWorkingConfig();
      }
    });
  };
  wirePick(".scn-pick-click", ".scn-click-selector");
  wirePick(".scn-pick-wait", ".scn-wait-selector");
  wirePick(".scn-pick-cond", ".scn-cond-selector");
  wirePick(".scn-pick-batch-scope", ".scn-batch-scope");
  wirePick(".scn-pick-batch-button", ".scn-batch-button");

  // Mode d'attente après le clic (étape « Éditer par lots »).
  const bWaitMode = div.querySelector(".scn-batch-waitmode");
  const bDelayWrap = div.querySelector(".scn-batch-delay-wrap");
  const bDlWrap = div.querySelector(".scn-batch-dl-wrap");
  bWaitMode.value = step.batchWaitMode || "delay";
  const syncBatchWaitUI = () => {
    const dl = bWaitMode.value !== "delay";
    bDelayWrap.hidden = dl;
    bDlWrap.hidden = !dl;
  };
  syncBatchWaitUI();
  bWaitMode.addEventListener("change", async () => {
    // chrome.permissions.request exige un geste utilisateur : on le fait
    // ici, dans le handler du change, avant tout autre await.
    if (bWaitMode.value !== "delay") {
      const granted = await requestDownloadsPermission();
      if (!granted) {
        bWaitMode.value = "delay";
        showStatus("Accès aux téléchargements refusé : l'étape restera en délai fixe.", "error");
      }
    }
    syncBatchWaitUI();
    persistWorkingConfig();
  });

  $("scenarioSteps").appendChild(div);
}

function getScenarioSteps() {
  return Array.from($("scenarioSteps").querySelectorAll(".scn-step")).map((el) => ({
    type: el.dataset.type,
    gotoUrl: el.querySelector(".scn-goto-url").value.trim(),
    gotoWait: el.querySelector(".scn-goto-wait").checked,
    gotoTimeout: parseInt(el.querySelector(".scn-goto-timeout").value, 10) || 15000,
    clickSelector: el.querySelector(".scn-click-selector").value.trim(),
    clickTimeout: parseInt(el.querySelector(".scn-click-timeout").value, 10) || 0,
    waitMode: el.querySelector(".scn-wait-mode").value,
    waitMs: parseInt(el.querySelector(".scn-wait-ms").value, 10) || 0,
    waitSelector: el.querySelector(".scn-wait-selector").value.trim(),
    waitTimeout: parseInt(el.querySelector(".scn-wait-timeout").value, 10) || 10000,
    condSource: el.querySelector(".scn-cond-source").value,
    condCol: el.querySelector(".scn-cond-col").value,
    condOp: el.querySelector(".scn-cond-op").value,
    condVal: el.querySelector(".scn-cond-val").value,
    condSelector: el.querySelector(".scn-cond-selector").value.trim(),
    condPageOp: el.querySelector(".scn-cond-pageop").value,
    condPageVal: el.querySelector(".scn-cond-pageval").value,
    condAction: el.querySelector(".scn-cond-action").value,
    condSkip: parseInt(el.querySelector(".scn-cond-skip").value, 10) || 1,
    pdfDocMode: el.querySelector(".scn-pdf-docmode").value,
    pdfDocMatch: el.querySelector(".scn-pdf-docmatch").value.trim(),
    pdfField: el.querySelector(".scn-pdf-field").value,
    pdfOp: el.querySelector(".scn-pdf-op").value,
    pdfVal: el.querySelector(".scn-pdf-val").value,
    pdfMissAction: el.querySelector(".scn-pdf-missaction").value,
    pdfMissSkip: parseInt(el.querySelector(".scn-pdf-skip").value, 10) || 1,
    pdfTargetCol: el.querySelector(".scn-pdf-targetcol").value,
    pdfNewCol: el.querySelector(".scn-pdf-newcol").value.trim(),
    sigeoAddressId: el.querySelector(".scn-sigeo-addressid").value.trim(),
    sigeoTable: el.querySelector(".scn-sigeo-table").value.trim(),
    sigeoBaseUrl: el.querySelector(".scn-sigeo-baseurl").value.trim(),
    sigeoNav: el.querySelector(".scn-sigeo-nav").checked,
    sigeoPays: el.querySelector(".scn-sigeo-pays").value.trim(),
    sigeoCp: el.querySelector(".scn-sigeo-cp").value.trim(),
    sigeoVille: el.querySelector(".scn-sigeo-ville").value.trim(),
    sigeoVoie: el.querySelector(".scn-sigeo-voie").value.trim(),
    sigeoNumVoie: el.querySelector(".scn-sigeo-numvoie").value.trim(),
    sigeoBp: el.querySelector(".scn-sigeo-bp").value.trim(),
    sigeoComplt: el.querySelector(".scn-sigeo-complt").value.trim(),
    sigeoCompNom: el.querySelector(".scn-sigeo-compnom").value.trim(),
    sigeoDryRun: el.querySelector(".scn-sigeo-dryrun").checked,
    sigeoTimeout: parseInt(el.querySelector(".scn-sigeo-timeout").value, 10) || 15000,
    sigeoResultCol: el.querySelector(".scn-sigeo-resultcol").value.trim(),
    batchScope: el.querySelector(".scn-batch-scope").value.trim(),
    batchButton: el.querySelector(".scn-batch-button").value.trim(),
    batchSize: parseInt(el.querySelector(".scn-batch-size").value, 10) || 10,
    batchWaitMode: el.querySelector(".scn-batch-waitmode").value,
    batchWaitMs: parseInt(el.querySelector(".scn-batch-wait").value, 10) || 0,
    batchDlTimeoutMs: parseInt(el.querySelector(".scn-batch-dltimeout").value, 10) || 30000,
    batchDlSettleMs: parseInt(el.querySelector(".scn-batch-dlsettle").value, 10) || 0,
    batchFilter: el.querySelector(".scn-batch-filter").value.trim(),
    batchMaxRounds: parseInt(el.querySelector(".scn-batch-max").value, 10) || 50
  }));
}

// Gestion du dépôt (drag & drop) au niveau du conteneur d'étapes
(function initScenarioDnD() {
  const container = $("scenarioSteps");
  if (!container) return;

  function stepAfterCursor(y) {
    const steps = [...container.querySelectorAll(".scn-step:not(.scn-dragging)")];
    return steps.reduce((closest, child) => {
      const box = child.getBoundingClientRect();
      const offset = y - box.top - box.height / 2;
      if (offset < 0 && offset > closest.offset) return { offset, element: child };
      return closest;
    }, { offset: Number.NEGATIVE_INFINITY, element: null }).element;
  }

  container.addEventListener("dragover", (e) => {
    const dragging = container.querySelector(".scn-dragging");
    if (!dragging) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    const after = stepAfterCursor(e.clientY);
    if (after == null) container.appendChild(dragging);
    else container.insertBefore(dragging, after);
  });
})();

$("addStepBtn").addEventListener("click", () => {
  addScenarioStep();
  persistWorkingConfig();
  updateScenarioEstimate();
});

$("clearStepsBtn").addEventListener("click", () => {
  if (!$("scenarioSteps").children.length) return;
  $("scenarioSteps").innerHTML = "";
  persistWorkingConfig();
  updateScenarioEstimate();
  showStatus("Étapes du scénario supprimées.", "info");
});

// Recalcule l'estimation quand une étape du scénario est modifiée.
$("scenarioSteps").addEventListener("input", updateScenarioEstimate);
$("scenarioSteps").addEventListener("change", updateScenarioEstimate);

/* ---------- Fonctions injectées (autonomes) ---------- */

// Attend qu'un élément apparaisse (visible) ou disparaisse, avec timeout.
function scnWaitInjected(cfg) {
  return new Promise((resolve) => {
    const deadline = Date.now() + Math.max(0, cfg.timeoutMs || 0);
    function isVisible(el) {
      if (!el) return false;
      const cs = window.getComputedStyle(el);
      if (cs.display === "none" || cs.visibility === "hidden") return false;
      return el.getClientRects().length > 0;
    }
    (function poll() {
      let el = null;
      try { el = document.querySelector(cfg.selector); }
      catch (e) { return resolve({ ok: false, error: "sélecteur invalide : " + cfg.selector }); }
      const visible = isVisible(el);
      if (cfg.mode === "gone" ? !visible : visible) return resolve({ ok: true });
      if (Date.now() >= deadline) {
        return resolve({
          ok: false,
          error: (cfg.mode === "gone" ? "l'élément est toujours visible après attente : " : "élément introuvable/invisible après attente : ") + cfg.selector
        });
      }
      setTimeout(poll, 120);
    })();
  });
}

// Attend l'élément (jusqu'au timeout) puis clique dessus.
function scnClickInjected(cfg) {
  return new Promise((resolve) => {
    const deadline = Date.now() + Math.max(0, cfg.timeoutMs || 0);
    function isVisible(el) {
      if (!el) return false;
      const cs = window.getComputedStyle(el);
      if (cs.display === "none" || cs.visibility === "hidden") return false;
      return el.getClientRects().length > 0;
    }
    function doClick(el, note) {
      try {
        if (el.scrollIntoView) el.scrollIntoView({ block: "center", inline: "center" });
        const opts = { bubbles: true, cancelable: true, view: window };
        el.dispatchEvent(new MouseEvent("mousedown", opts));
        el.dispatchEvent(new MouseEvent("mouseup", opts));
        el.click();
        resolve({ ok: true, info: note });
      } catch (e) {
        resolve({ ok: false, error: "échec du clic : " + (e.message || e) });
      }
    }
    (function poll() {
      let el = null;
      try { el = document.querySelector(cfg.selector); }
      catch (e) { return resolve({ ok: false, error: "sélecteur invalide : " + cfg.selector }); }
      if (el && isVisible(el)) return doClick(el, "");
      if (Date.now() >= deadline) {
        // Présent mais masqué : on tente quand même (menus/onglets techniques).
        if (el) return doClick(el, "cliqué (élément non visible)");
        return resolve({ ok: false, error: "élément à cliquer introuvable : " + cfg.selector });
      }
      setTimeout(poll, 120);
    })();
  });
}

// Teste l'état d'un élément (condition « page »).
function scnCheckInjected(cfg) {
  let el = null;
  try { el = document.querySelector(cfg.selector); }
  catch (e) { return { ok: false, error: "sélecteur invalide : " + cfg.selector }; }
  function isVisible(node) {
    if (!node) return false;
    const cs = window.getComputedStyle(node);
    if (cs.display === "none" || cs.visibility === "hidden") return false;
    return node.getClientRects().length > 0;
  }
  let text = "";
  if (el) {
    text = ("value" in el && el.tagName !== "DIV") ? String(el.value) : (el.innerText || el.textContent || "");
    text = text.trim();
  }
  const present = isVisible(el);
  const needle = String(cfg.value || "").trim().toLowerCase();
  const hay = text.toLowerCase();
  let match = false;
  switch (cfg.op) {
    case "exists": match = present; break;
    case "not_exists": match = !present; break;
    case "contains": match = !!el && hay.includes(needle); break;
    case "not_contains": match = !el || !hay.includes(needle); break;
    case "equals": match = !!el && hay === needle; break;
  }
  return { ok: true, match };
}

/* ---------- Étape SIGEO : saisie d'adresse (address_manage) ---------- */
// Automatise le formulaire d'adresse SIGEO (ASP.NET WebForms). Principe :
// pilotage du DOM dans la page (remplissage + clic « Enreg. et fermer »), jamais
// de POST forgé — le navigateur gère __VIEWSTATE/__EVENTVALIDATION lui-même.
// Le point délicat est la résolution des ids internes (code commune
// « Migrated_code », value du n° de voie) : on privilégie la cascade CP→ville de
// la page puis son sélecteur de commune (iframe ou fenêtre séparée).

// Noms (attributs name) des champs du formulaire — stables entre postbacks.
const SIGEO_FIELDS = {
  pays: "body:x:ddlPays:ddlPays",
  cp: "body:x:city:x:Migrated_txtCP:x:CP",
  ville: "body:x:city:x:Migrated_lVille:x:lVille",
  migratedCode: "body:x:city:x:Migrated_code",
  cityCountry: "body:x:city:x:country",
  voie: "body:x:txtVoie:x:Voie",
  idVoie: "body:x:idVoie",
  numVoie: "body:x:ddlNum_voie:ddlNum_voie",
  bp: "body:x:txtBoitePostale:x:Num",
  complt: "body:x:txtComplt:x:Complt",
  compNom: "body:x:txtCompNom:x:CompNom",
  saveBtn: "proxyActionBar:x:_cmdEnd:x:_btn"
};

// Sonde une frame : formulaire adresse présent ? page de connexion ?
function sigeoProbeInjected(cfg) {
  const has = (n) => document.getElementsByName(n).length > 0;
  const form = has(cfg.paysName) && has(cfg.btnName);
  const login = !!document.querySelector('input[type="password"]') ||
    /login|logon|connexion/i.test(location.pathname);
  return { form, login };
}

// Cherche (avec attente) la frame qui contient le formulaire adresse.
// Renvoie { frameId } | { auth: true } | { notFound: true }.
async function sigeoProbe(tabId, timeoutMs) {
  const deadline = Date.now() + Math.max(3000, timeoutMs || 15000);
  let sawLogin = false;
  while (Date.now() < deadline && !scnStopRequested) {
    let results = [];
    try {
      results = await chrome.scripting.executeScript({
        target: { tabId, allFrames: true },
        func: sigeoProbeInjected,
        args: [{ paysName: SIGEO_FIELDS.pays, btnName: SIGEO_FIELDS.saveBtn }]
      });
    } catch (_) { /* page en cours de navigation : on réessaie */ }
    const hit = (results || []).find((r) => r && r.result && r.result.form);
    if (hit) return { frameId: hit.frameId };
    if ((results || []).some((r) => r && r.result && r.result.login)) sawLogin = true;
    await scnSleep(400);
  }
  return sawLogin ? { auth: true } : { notFound: true };
}

// Remplit le formulaire (idempotent : ne touche que ce qui diffère) puis, hors
// simulation, clique « Enreg. et fermer ». Autonome (sérialisée dans la page).
async function sigeoFillInjected(cfg) {
  const F = cfg.fields, V = cfg.values;
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const byName = (n, doc = document) => (doc.getElementsByName(n)[0] || null);
  const log = [];

  // Tous les documents accessibles (frame courante + iframes same-origin),
  // pour retrouver un éventuel sélecteur de commune ouvert en surimpression.
  function allDocs() {
    const docs = [];
    (function walk(win) {
      try { if (win.document) docs.push(win.document); } catch (_) { return; }
      try { for (let i = 0; i < win.frames.length; i++) walk(win.frames[i]); } catch (_) {}
    })(window);
    return docs;
  }
  function isVisible(el) {
    if (!el) return false;
    try {
      const cs = el.ownerDocument.defaultView.getComputedStyle(el);
      if (cs.display === "none" || cs.visibility === "hidden") return false;
      return el.getClientRects().length > 0;
    } catch (_) { return false; }
  }
  const fire = (el, type) => el.dispatchEvent(new Event(type, { bubbles: true }));
  function setInput(el, val) {
    try { el.focus(); } catch (_) {}
    el.value = val;
    fire(el, "input");
    fire(el, "change");
    try { el.blur(); } catch (_) {}
  }
  const norm = (s) => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase().replace(/\s+/g, " ");
  const optLabel = (sel) => (sel && sel.selectedIndex >= 0 ? sel.options[sel.selectedIndex].text.trim() : "");

  // — Références des champs. Un postback partiel (UpdatePanel) peut remplacer
  // les nœuds : on re-requête avant chaque lecture/écriture sensible.
  const el = {};
  const refreshEls = () => { for (const k of Object.keys(F)) { const n = byName(F[k]); if (n) el[k] = n; } };
  for (const k of Object.keys(F)) el[k] = byName(F[k]);
  if (!el.pays || !el.saveBtn) return { found: false };

  const codeOk = () => { refreshEls(); return el.migratedCode && String(el.migratedCode.value || "").trim() !== ""; };
  const villeOk = () => { refreshEls(); return el.ville && norm(el.ville.value) === norm(V.ville); };

  // — Idempotence : si tout est déjà à la valeur cible, ne pas resoumettre.
  const optionalsMatch = () =>
    (!V.bp || norm(el.bp && el.bp.value) === norm(V.bp)) &&
    (!V.complt || norm(el.complt && el.complt.value) === norm(V.complt)) &&
    (!V.compNom || norm(el.compNom && el.compNom.value) === norm(V.compNom));
  const allMatch = () =>
    norm(el.pays.value) === norm(V.pays) &&
    norm(el.cp && el.cp.value) === norm(V.cp) &&
    villeOk() && codeOk() &&
    norm(el.voie && el.voie.value) === norm(V.voie) &&
    (el.numVoie ? norm(optLabel(el.numVoie)) === norm(V.numVoie) : false) &&
    optionalsMatch();
  if (allMatch()) {
    return {
      found: true, status: "already",
      resolvedIds: { migratedCode: el.migratedCode.value, idVoie: el.idVoie ? el.idVoie.value : "", numVoieValue: el.numVoie.value },
      log: ["valeurs déjà identiques — pas de resoumission"]
    };
  }

  // — 1. Pays (select par value, ISO2 minuscule)
  if (norm(el.pays.value) !== norm(V.pays)) {
    const opt = Array.from(el.pays.options).find((o) => norm(o.value) === norm(V.pays));
    if (!opt) return { found: true, status: "validation_error", errors: [`pays « ${V.pays} » absent de la liste`], log };
    el.pays.value = opt.value;
    fire(el.pays, "change");
    log.push(`pays = ${opt.value}`);
    await sleep(300); // laisse le JS de la page réagir
    refreshEls();
  }

  // — 2/3. Code postal + ville (le champ caché Migrated_code doit être cohérent)
  let cityVia = "déjà résolu";
  if (!(villeOk() && codeOk())) {
    if (!el.cp || !el.ville) return { found: true, status: "validation_error", errors: ["champs CP/ville introuvables"], log };
    if (norm(el.cp.value) !== norm(V.cp)) {
      // Frappe façon autocomplétion « searchResult » : valeur + keyup, SANS blur
      // (un blur risquerait un __doPostBack complet qui tuerait l'injection).
      try { el.cp.focus(); } catch (_) {}
      el.cp.value = V.cp;
      fire(el.cp, "input");
      el.cp.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, cancelable: true, key: String(V.cp).slice(-1) || "Unidentified" }));
      log.push(`CP = ${V.cp}`);
    }
    // 2a. suggestions d'autocomplétion (mécanisme « searchResult » du site) :
    // elles arrivent dans select[name="searchResultSelect_{id}"] (ou
    // #search:{id}) ; cliquer une <option> appelle setDataFieldValue, qui
    // remplit la ville ET le code commune (Migrated_code) de façon cohérente.
    const findSuggSelect = () => {
      refreshEls();
      const id = (el.cp && el.cp.id) || "";
      if (!id) return null;
      let sel = null;
      try { sel = document.querySelector(`select[name="searchResultSelect_${CSS.escape(id)}"]`); } catch (_) {}
      if (!sel) {
        const cont = document.getElementById("search:" + id);
        if (cont) sel = cont.querySelector("select");
      }
      return sel;
    };
    let until = Date.now() + 8000;
    let suggPicked = false;
    while (Date.now() < until && !(villeOk() && codeOk())) {
      const sel = findSuggSelect();
      if (sel && !suggPicked) {
        const opts = Array.from(sel.options).filter((o) => (o.text || "").trim() !== "");
        if (opts.length) {
          const wantV = norm(V.ville), wantCp = String(V.cp).trim();
          let best = null, bestScore = 0;
          for (const o of opts) {
            const t = norm(o.text);
            let sc = 0;
            if (t === wantV || t.endsWith("- " + wantV)) sc = 3;
            else if (t.includes(wantV)) sc = 1;
            if (wantCp && t.includes(wantCp)) sc += 1;
            if (sc > bestScore) { bestScore = sc; best = o; }
          }
          if (best) {
            best.selected = true;
            for (const t of ["mousedown", "mouseup", "click"]) {
              best.dispatchEvent(new MouseEvent(t, { bubbles: true, cancelable: true, view: window }));
            }
            suggPicked = true;
            log.push("suggestion choisie : " + best.text.trim());
            until = Date.now() + 4000; // laisse setDataFieldValue propager ville + code
          }
        }
      }
      await sleep(200);
    }
    if (villeOk() && codeOk()) cityVia = suggPicked ? "autocomplétion CP" : "cascade CP";
    else if (cfg.forceDirect) {
      if (!villeOk()) setInput(el.ville, V.ville);
      cityVia = "saisie directe";
    } else {
      // 2b. sélection assistée : bouton du sélecteur de commune (« … »)
      const looksCity = (node) => {
        const s = ((node.getAttribute("onclick") || "") + " " + (node.getAttribute("href") || "") + " " +
          (node.id || "") + " " + (node.title || "") + " " + (node.getAttribute("src") || "")).toLowerCase();
        return s.includes("city_selector") || s.includes("city") || s.includes("commune");
      };
      let btn = Array.from(document.querySelectorAll("[onclick],a[href]"))
        .find((n) => isVisible(n) && (((n.getAttribute("onclick") || "") + (n.getAttribute("href") || "")).toLowerCase().includes("city_selector")));
      if (!btn) {
        for (const ref of [el.ville, el.cp]) {
          let node = ref;
          for (let up = 0; up < 4 && node && !btn; up++) {
            node = node.parentElement;
            if (!node) break;
            btn = Array.from(node.querySelectorAll("a,button,input[type=button],input[type=image],img"))
              .find((n) => isVisible(n) && looksCity(n)) || null;
          }
          if (btn) break;
        }
      }
      if (btn) {
        // instantané avant clic : seules les lignes NOUVELLES seront candidates
        const pre = new WeakSet();
        allDocs().forEach((d) => d.querySelectorAll("tr,a").forEach((n) => pre.add(n)));
        btn.click();
        log.push("sélecteur de commune ouvert");
        const wantVille = norm(V.ville), wantCp = String(V.cp).trim();
        let clicked = false;
        until = Date.now() + 3500;
        while (Date.now() < until && !(villeOk() && codeOk()) && !clicked) {
          for (const d of allDocs()) {
            let best = null, bestScore = -1;
            for (const tr of d.querySelectorAll("tr")) {
              if (pre.has(tr) || !isVisible(tr)) continue;
              const cells = Array.from(tr.cells || []).map((td) => norm(td.innerText));
              if (!cells.some((c) => c === wantVille)) continue;
              const score = 1 + (wantCp && cells.some((c) => c.includes(wantCp)) ? 1 : 0);
              if (score > bestScore) { best = tr; bestScore = score; }
            }
            if (!best) {
              const link = Array.from(d.querySelectorAll("a"))
                .find((a) => !pre.has(a) && isVisible(a) && norm(a.innerText) === wantVille);
              if (link) { link.click(); clicked = true; log.push("commune cliquée (lien)"); break; }
            } else {
              const link = best.querySelector("a");
              (link || best).click();
              if (!link) { try { best.dispatchEvent(new MouseEvent("dblclick", { bubbles: true })); } catch (_) {} }
              clicked = true;
              log.push("commune cliquée (ligne" + (bestScore > 1 ? " CP+ville" : "") + ")");
              break;
            }
          }
          if (!clicked) await sleep(200);
        }
        if (clicked) {
          until = Date.now() + 4000;
          while (Date.now() < until && !codeOk()) await sleep(150);
          cityVia = "sélecteur (popup interne)";
        } else if (!(villeOk() && codeOk())) {
          // rien d'accessible dans la page : le sélecteur s'est sans doute ouvert
          // dans une fenêtre séparée — la barre latérale prend le relais.
          return { found: true, status: "city_popup", log };
        }
      } else {
        if (!villeOk()) setInput(el.ville, V.ville);
        cityVia = "saisie directe (aucun sélecteur trouvé)";
      }
    }
    if (!codeOk()) {
      return {
        found: true, status: "validation_error",
        errors: [`UNRESOLVED_CITY — code commune (Migrated_code) non résolu pour « ${V.ville} » (${V.cp}) via ${cityVia} : vérifier le couple CP/ville`],
        log
      };
    }
    if (!villeOk()) setInput(el.ville, V.ville);
    log.push(`ville = ${el.ville.value} (code ${el.migratedCode.value}, via ${cityVia})`);
  }

  // — 4. Voie (idVoie laissé tel quel : géré par la page si référentiel)
  refreshEls();
  if (el.voie && norm(el.voie.value) !== norm(V.voie)) {
    setInput(el.voie, V.voie);
    log.push(`voie = ${V.voie}`);
    await sleep(300);
    refreshEls();
  }

  // — 5. Numéro de voie (select : match strict sur le libellé, value = id interne)
  if (!el.numVoie) return { found: true, status: "validation_error", errors: ["liste « numéro de voie » introuvable"], log };
  let numRes = null;
  const untilNum = Date.now() + 4000; // la liste peut se recharger après la voie
  while (Date.now() < untilNum && !numRes) {
    refreshEls();
    const want = norm(V.numVoie);
    const opt = Array.from(el.numVoie.options).find((o) => norm(o.text) === want);
    if (opt) {
      if (el.numVoie.value !== opt.value) { el.numVoie.value = opt.value; fire(el.numVoie, "change"); }
      numRes = { value: opt.value, label: opt.text.trim() };
    } else await sleep(200);
  }
  if (!numRes) {
    const avail = Array.from(el.numVoie.options).map((o) => o.text.trim()).filter(Boolean).slice(0, 12);
    return {
      found: true, status: "validation_error",
      errors: [`NUM_VOIE_INTROUVABLE — libellé « ${V.numVoie} » absent de la liste${avail.length ? " (dispo : " + avail.join(", ") + "…)" : " (liste vide)"}`],
      log
    };
  }
  log.push(`n° voie = ${numRes.label} (value ${numRes.value})`);

  // — 6. Champs optionnels (renseignés seulement s'ils sont fournis)
  const optFields = [["bp", V.bp], ["complt", V.complt], ["compNom", V.compNom]];
  for (const [k, v] of optFields) {
    if (v && el[k] && norm(el[k].value) !== norm(v)) { setInput(el[k], v); log.push(`${k} = ${v}`); }
  }

  // — 7. Contrôles avant soumission
  refreshEls();
  const errors = [];
  if (!String(el.pays.value || "").trim()) errors.push("pays vide");
  if (!String(el.cp.value || "").trim()) errors.push("code postal vide");
  if (norm(V.pays) === "FR" && !/^\d{5}$/.test(String(el.cp.value).trim())) errors.push(`CP incohérent pour la France : « ${el.cp.value} »`);
  if (!String(el.ville.value || "").trim()) errors.push("ville vide");
  if (!String(el.voie.value || "").trim()) errors.push("voie vide");
  if (!String(el.numVoie.value || "").trim()) errors.push("numéro de voie non sélectionné");
  // messages de validation déjà affichés ?
  document.querySelectorAll('[id*="Validator"], [id*="ValidationSummary"], .validation-summary-errors, .field-validation-error').forEach((n) => {
    if (!isVisible(n)) return;
    const t = (n.innerText || "").trim();
    if (t && t.length < 300 && !errors.includes(t)) errors.push(t);
  });
  const report = {
    submittedValues: {
      pays: el.pays.value, codePostal: el.cp.value, ville: el.ville.value,
      voie: el.voie.value, numeroVoie: numRes.label,
      boitePostale: el.bp ? el.bp.value : "", complement: el.complt ? el.complt.value : "",
      complementNom: el.compNom ? el.compNom.value : ""
    },
    resolvedIds: {
      migratedCode: el.migratedCode ? el.migratedCode.value : "",
      idVoie: el.idVoie ? el.idVoie.value : "",
      numVoieValue: numRes.value
    },
    cityVia, log
  };
  if (errors.length) return { found: true, status: "validation_error", errors, ...report };

  // — 8. Simulation : tout est rempli et contrôlé, on n'enregistre pas.
  if (cfg.dryRun) return { found: true, status: "dry_run_ok", ...report };

  // — 9. Soumission : clic différé pour laisser le rapport repartir avant le postback.
  const saveBtn = el.saveBtn;
  setTimeout(() => { try { saveBtn.click(); } catch (_) {} }, 60);
  return { found: true, status: "submitted", ...report };
}

// Dans la fenêtre séparée du sélecteur de commune : clique la bonne ligne.
function sigeoCityPickInjected(cfg) {
  return new Promise((resolve) => {
    const norm = (s) => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase().replace(/\s+/g, " ");
    const want = norm(cfg.ville), cp = String(cfg.cp || "").trim();
    const deadline = Date.now() + Math.max(2000, cfg.timeoutMs || 8000);
    (function poll() {
      let best = null, bestScore = -1;
      document.querySelectorAll("tr").forEach((tr) => {
        const cells = Array.from(tr.cells || []).map((td) => norm(td.innerText));
        if (!cells.some((c) => c === want)) return;
        const score = 1 + (cp && cells.some((c) => c.includes(cp)) ? 1 : 0);
        if (score > bestScore) { best = tr; bestScore = score; }
      });
      if (best) {
        const link = best.querySelector("a");
        (link || best).click();
        if (!link) { try { best.dispatchEvent(new MouseEvent("dblclick", { bubbles: true })); } catch (_) {} }
        return resolve({ ok: true });
      }
      const link = Array.from(document.querySelectorAll("a")).find((a) => norm(a.innerText) === want);
      if (link) { link.click(); return resolve({ ok: true }); }
      if (Date.now() >= deadline) return resolve({ ok: false, error: "commune introuvable dans le sélecteur : " + cfg.ville });
      setTimeout(poll, 200);
    })();
  });
}

// Après postback : détecte succès / erreurs de validation dans chaque frame.
function sigeoResultInjected(cfg) {
  const F = cfg.fields;
  const byName = (n) => (document.getElementsByName(n)[0] || null);
  const login = !!document.querySelector('input[type="password"]') || /login|logon|connexion/i.test(location.pathname);
  const form = !!byName(F.pays) && !!byName(F.saveBtn);
  function isVisible(el) {
    if (!el) return false;
    const cs = getComputedStyle(el);
    if (cs.display === "none" || cs.visibility === "hidden") return false;
    return el.getClientRects().length > 0;
  }
  const errors = [];
  document.querySelectorAll('[id*="Validator"], [id*="ValidationSummary"], .validation-summary-errors, .field-validation-error, .error, span, div').forEach((el) => {
    if (errors.length >= 5 || !isVisible(el) || el.children.length > 2) return;
    const t = (el.innerText || "").trim();
    if (!t || t.length > 300) return;
    const idCls = (el.id + " " + el.className).toLowerCase();
    let flagged = /valid|error|erreur/.test(idCls);
    if (!flagged) {
      const m = getComputedStyle(el).color.match(/rgb\((\d+),\s*(\d+),\s*(\d+)/);
      flagged = !!m && +m[1] > 150 && +m[2] < 100 && +m[3] < 100; // texte « rouge »
    }
    if (flagged && !errors.includes(t)) errors.push(t);
  });
  return {
    login, form, errors,
    values: form ? {
      cp: byName(F.cp) ? byName(F.cp).value : "",
      code: byName(F.migratedCode) ? byName(F.migratedCode).value : ""
    } : null
  };
}

// Attend la fin du postback ; détecte la fermeture de l'onglet (popup fermée).
async function sigeoWaitPostback(tabId, timeoutMs) {
  const deadline = Date.now() + Math.max(2000, timeoutMs || 15000);
  await scnSleep(700); // laisse partir le POST
  while (Date.now() < deadline && !scnStopRequested) {
    let t;
    try { t = await chrome.tabs.get(tabId); }
    catch (_) { return { closed: true }; }
    if (t.status === "complete") { await scnSleep(400); return { closed: false }; }
    await scnSleep(250);
  }
  return { closed: false, timedOut: true };
}

// Le sélecteur de commune s'est ouvert dans une fenêtre séparée : on la
// retrouve, on clique la commune, on attend la fermeture (propagation opener).
async function sigeoHandleCityPopup(tabId, vals, timeoutMs) {
  const deadline = Date.now() + 3000;
  let popup = null;
  while (Date.now() < deadline && !popup && !scnStopRequested) {
    const tabs = await chrome.tabs.query({});
    popup = tabs.find((t) => t.id !== tabId && /city|commune/i.test(t.url || "")) ||
      tabs.find((t) => t.id !== tabId && /popup\.aspx/i.test(t.url || "") && !/address_manage/i.test(t.url || ""));
    if (!popup) await scnSleep(300);
  }
  if (!popup) return false;
  let res = null;
  try {
    const r = await chrome.scripting.executeScript({
      target: { tabId: popup.id },
      func: sigeoCityPickInjected,
      args: [{ ville: vals.ville, cp: vals.cp, timeoutMs: Math.min(8000, timeoutMs || 8000) }]
    });
    res = r && r[0] && r[0].result;
  } catch (_) { return false; }
  if (!res || !res.ok) return false;
  const closeDeadline = Date.now() + 6000;
  while (Date.now() < closeDeadline && !scnStopRequested) {
    try { await chrome.tabs.get(popup.id); } catch (_) { return true; } // fermée
    await scnSleep(250);
  }
  return true; // sélection cliquée ; la reprise vérifiera la propagation
}

// Écrit le résultat de l'étape dans la colonne configurée (créée si besoin).
function sigeoWriteResult(step, rowIdx, text) {
  const ref = (step.sigeoResultCol || "").trim();
  if (!ref || rowIdx === null || rowIdx === undefined || !state.rows.length) return;
  const row = state.rows[rowIdx] = state.rows[rowIdx] || [];
  const existed = colIndexByName(ref) >= 0;
  const idx = resolveOutputTarget(ref, {});
  if (idx < 0) return;
  setCellByIndex(row, idx, text);
  persistSession();
  if (existed) { renderPreview(); renderSelectedRowFields(); }
  else renderColumns();
}

// Orchestrateur de l'étape « SIGEO : saisir une adresse ».
async function execSigeoStep(step, rowIdx, tabId) {
  const t0 = Date.now();
  const row = (rowIdx !== null && rowIdx !== undefined) ? (state.rows[rowIdx] || []) : null;
  const rt = (v) => String(resolveRowTemplate(v ?? "", row) ?? "").trim();
  const timeoutMs = step.sigeoTimeout || 15000;
  const vals = {
    addressId: rt(step.sigeoAddressId),
    table: rt(step.sigeoTable),
    pays: (rt(step.sigeoPays) || "fr").toLowerCase(),
    cp: rt(step.sigeoCp),
    ville: rt(step.sigeoVille),
    voie: rt(step.sigeoVoie),
    numVoie: rt(step.sigeoNumVoie),
    bp: rt(step.sigeoBp),
    complt: rt(step.sigeoComplt),
    compNom: rt(step.sigeoCompNom)
  };
  const dur = () => ` — ${((Date.now() - t0) / 1000).toFixed(1)} s`;
  const fail = (msg) => { sigeoWriteResult(step, rowIdx, "ERREUR — " + msg); return { ok: false, error: msg + dur() }; };

  // — Contrôles d'entrée
  const missing = [];
  if (step.sigeoNav !== false && !vals.addressId) missing.push("ID adresse");
  if (step.sigeoNav !== false && !vals.table) missing.push("table");
  if (!vals.cp) missing.push("code postal");
  if (!vals.ville) missing.push("ville");
  if (!vals.voie) missing.push("voie");
  if (!vals.numVoie) missing.push("n° de voie");
  if (missing.length) return fail("champ(s) requis vide(s) : " + missing.join(", "));
  if (vals.pays === "fr" && !/^\d{5}$/.test(vals.cp)) return fail(`CP invalide pour la France : « ${vals.cp} »`);

  // — Navigation vers la fiche adresse (chargement frais → ViewState neuf)
  if (step.sigeoNav !== false) {
    const base = (rt(step.sigeoBaseUrl) || "http://sigeo.veille.evoriel.net").replace(/\/+$/, "");
    const url = `${base}/popup.aspx/fr/sig/address_manage/${encodeURIComponent(vals.addressId)}` +
      `?table=${encodeURIComponent(vals.table)}&adrTypeId=atype&deletable=true&atypeCode=PRI`;
    const nav = await navigateTabAndWait(tabId, url, true, timeoutMs);
    if (!nav.ok) return fail("navigation : " + nav.error);
  }

  // — Remplissage (avec reprise : popup de commune, frame rechargée…)
  let rep = null;
  let forceDirect = false;
  let popupTried = false;
  for (let attempt = 1; attempt <= 4; attempt++) {
    if (scnStopRequested) return fail("arrêté par l'utilisateur");
    const probe = await sigeoProbe(tabId, timeoutMs);
    if (probe.auth) return fail("AUTH_REQUIRED — page de connexion détectée (session SIGEO expirée ?)");
    if (probe.notFound) return fail("formulaire adresse introuvable (page inattendue ?)");
    let results = null;
    try {
      results = await chrome.scripting.executeScript({
        target: { tabId, frameIds: [probe.frameId] },
        func: sigeoFillInjected,
        args: [{ fields: SIGEO_FIELDS, values: vals, dryRun: step.sigeoDryRun !== false, forceDirect, timeoutMs }]
      });
    } catch (e) {
      if (attempt === 4) return fail("injection impossible : " + (e.message || e));
      await scnSleep(800);
      continue; // la frame a probablement rechargé (autopostback) : on reprend
    }
    rep = results && results[0] && results[0].result;
    if (!rep || !rep.found) { await scnSleep(800); continue; }
    if (rep.status === "city_popup") {
      if (popupTried) forceDirect = true; // deuxième ouverture : on n'insiste pas
      else {
        popupTried = true;
        const handled = await sigeoHandleCityPopup(tabId, vals, timeoutMs);
        if (handled) await scnSleep(400); // propagation des hidden fields via opener
        else forceDirect = true;          // pas de popup accessible → repli saisie directe
      }
      rep = null;
      continue; // reprise idempotente du remplissage
    }
    break;
  }
  if (!rep || !rep.found) return fail("le remplissage n'a pas abouti (page instable ?)");

  const ids = rep.resolvedIds || {};
  const idsTxt = `commune ${ids.migratedCode || "?"}, n° voie value=${ids.numVoieValue || "?"}`;

  // — Issues sans soumission
  if (rep.status === "validation_error") {
    return fail("contrôle avant soumission : " + (rep.errors || ["erreur inconnue"]).join(" | "));
  }
  if (rep.status === "already") {
    sigeoWriteResult(step, rowIdx, "OK — déjà à jour (" + idsTxt + ")");
    return { ok: true, info: "déjà à jour, non resoumis (" + idsTxt + ")" + dur() };
  }
  if (rep.status === "dry_run_ok") {
    sigeoWriteResult(step, rowIdx, "SIMULATION OK — " + idsTxt + " (via " + (rep.cityVia || "?") + ")");
    return { ok: true, info: `simulation OK (non enregistré) — ${idsTxt}, ville via ${rep.cityVia || "?"}` + dur() };
  }
  if (rep.status !== "submitted") return fail("état inattendu : " + rep.status);

  // — Postback : on laisse le navigateur soumettre, puis on lit le résultat.
  const pb = await sigeoWaitPostback(tabId, timeoutMs);
  if (pb.closed) {
    sigeoWriteResult(step, rowIdx, "OK — enregistré (popup fermée) — " + idsTxt);
    return { ok: true, info: "enregistré — la popup s'est fermée (" + idsTxt + ")" + dur() };
  }
  let results = null;
  try {
    results = await chrome.scripting.executeScript({
      target: { tabId, allFrames: true },
      func: sigeoResultInjected,
      args: [{ fields: SIGEO_FIELDS }]
    });
  } catch (e) {
    sigeoWriteResult(step, rowIdx, "INCERTAIN — soumis, résultat illisible");
    return { ok: true, warn: true, info: "soumis, mais résultat illisible (" + (e.message || e) + ")" + dur() };
  }
  const frames = (results || []).map((r) => r && r.result).filter(Boolean);
  if (frames.some((f) => f.login)) return fail("AUTH_REQUIRED — session perdue pendant l'enregistrement");
  const errs = [...new Set(frames.flatMap((f) => f.errors || []))];
  if (errs.length) return fail("VALIDATION — " + errs.join(" | "));
  const formFrame = frames.find((f) => f.form);
  sigeoWriteResult(step, rowIdx, "OK — enregistré — " + idsTxt);
  if (formFrame) {
    return { ok: true, warn: true, info: "enregistré (aucune erreur détectée, formulaire toujours affiché) — " + idsTxt + dur() };
  }
  return { ok: true, info: "enregistré — " + idsTxt + dur() };
}

/* ---------- Moteur d'exécution ---------- */

function scnTrunc(s, n = 45) {
  s = String(s || "");
  return s.length > n ? s.slice(0, n) + "…" : s;
}

function scnStepLabel(s) {
  switch (s.type) {
    case "fill": return "Remplir";
    case "goto": return "Ouvrir " + scnTrunc(s.gotoUrl || "?");
    case "click": return "Cliquer " + scnTrunc(s.clickSelector || "?");
    case "wait":
      if (s.waitMode === "delay") return `Attendre ${s.waitMs} ms`;
      return (s.waitMode === "gone" ? "Attendre disparition de " : "Attendre ") + scnTrunc(s.waitSelector || "?");
    case "cond":
      return s.condSource === "page" ? `Si ${scnTrunc(s.condSelector || "?", 30)}…` : `Si « ${s.condCol} »…`;
    case "pdfcheck":
      return `PDF : vérifier « ${tbFieldLabel(s.pdfField)} »`;
    case "pdfwrite":
      return `PDF : « ${tbFieldLabel(s.pdfField)} » → ${s.pdfTargetCol === "__other__" ? (s.pdfNewCol || "?") : (s.pdfTargetCol || "?")}`;
    case "sigeo":
      return `SIGEO : adresse ${scnTrunc(s.sigeoAddressId || "?", 20)}${s.sigeoDryRun !== false ? " (simulation)" : ""}`;
    case "batchedit": return `Éditer par lots de ${s.batchSize || 10}`;
  }
  return s.type;
}

/* ---------- Attente d'un téléchargement (étape « Éditer par lots ») ----------
   La permission « downloads » est optionnelle (optional_permissions du
   manifest) : elle n'est demandée que quand l'utilisateur choisit un mode
   d'attente basé sur le téléchargement — ainsi la mise à jour de
   l'extension ne réclame aucune nouvelle permission en bloc. Tant qu'elle
   n'est pas accordée, chrome.downloads est indéfini. */

async function hasDownloadsPermission() {
  try {
    if (!chrome.permissions) return false;
    return await chrome.permissions.contains({ permissions: ["downloads"] });
  } catch (_) { return false; }
}

async function requestDownloadsPermission() {
  try {
    if (!chrome.permissions) return false;
    // Renvoie true sans invite si déjà accordée.
    return await chrome.permissions.request({ permissions: ["downloads"] });
  } catch (_) { return false; }
}

/**
 * Guette le PROCHAIN téléchargement créé par le navigateur.
 * À armer AVANT le clic qui le déclenche : un fichier peut partir
 * immédiatement, un écouteur posé après manquerait l'événement.
 *
 * mode 'start'    → résout dès la création du téléchargement ;
 * mode 'complete' → attend en plus que son état passe à « complete »
 *                   (ou « interrupted », signalé comme échec).
 *
 * Renvoie une promesse de { started, completed?, interrupted?, filename }
 * qui expose .cancel() pour désarmer le guet.
 */
function watchNextDownload(mode, timeoutMs) {
  let cancel = () => {};
  const promise = new Promise((resolve) => {
    let done = false;
    let timer = null;
    let watchedId = null;

    const cleanup = () => {
      clearTimeout(timer);
      try { chrome.downloads.onCreated.removeListener(onCreated); } catch (_) { /* ignoré */ }
      try { chrome.downloads.onChanged.removeListener(onChanged); } catch (_) { /* ignoré */ }
    };
    const finish = (res) => {
      if (done) return;
      done = true;
      cleanup();
      resolve(res);
    };

    const baseName = (item) =>
      String(item.filename || "").split(/[\\/]/).pop() || item.url || "";

    const onChanged = (delta) => {
      if (delta.id !== watchedId || !delta.state) return;
      if (delta.state.current === "complete") finish({ started: true, completed: true, filename: lastName });
      else if (delta.state.current === "interrupted") finish({ started: true, interrupted: true, filename: lastName });
    };

    let lastName = "";
    const onCreated = (item) => {
      lastName = baseName(item);
      if (mode !== "complete") { finish({ started: true, filename: lastName }); return; }
      watchedId = item.id;
      if (item.state === "complete") { finish({ started: true, completed: true, filename: lastName }); return; }
      try {
        chrome.downloads.onChanged.addListener(onChanged);
      } catch (_) {
        finish({ started: true, filename: lastName }); // repli : au moins démarré
      }
    };

    timer = setTimeout(() => finish({ started: false, timedOut: true }), timeoutMs);
    cancel = () => finish({ started: false, cancelled: true });
    try {
      chrome.downloads.onCreated.addListener(onCreated);
    } catch (_) {
      finish({ started: false, unavailable: true });
    }
  });
  promise.cancel = () => cancel();
  return promise;
}

// Exécute une étape. Retourne { ok, error?, info?, warn?, skip?, stop? }.
// opts.soloTab : le remplissage ne vise que `tabId` (mode multi-onglets),
// au lieu de tous les onglets du site.
async function execScenarioStep(step, rowIdx, tabId, opts = {}) {
  try {
    switch (step.type) {
      case "fill": {
        const { data, mapping, customFields, rowContext } = buildFillPayload(rowIdx);
        if (!Object.keys(mapping).length && !Object.keys(customFields).length) {
          return { ok: false, error: "aucun mapping ni champ personnalisé configuré" };
        }
        const fillMsg = {
          action: "fillForm",
          data,
          mapping,
          customFields: Object.keys(customFields).length ? customFields : undefined,
          rowContext
        };
        const { totalFilled } = opts.soloTab
          ? await sendFillToOneTab(tabId, fillMsg)
          : await sendFillToAllTabs(fillMsg);
        return { ok: true, info: `${totalFilled} champ(s) rempli(s)`, warn: totalFilled === 0 };
      }
      case "goto": {
        if (!step.gotoUrl) return { ok: false, error: "URL manquante" };
        let url = step.gotoUrl;
        if (/\{[^{}]+\}/.test(url)) {
          if (rowIdx === null || rowIdx === undefined) return { ok: false, error: "aucune ligne active pour construire l'URL" };
          const { url: built, missing, values } = buildNavUrl(url, state.rows[rowIdx] || []);
          if (missing.length) return { ok: false, error: "colonne(s) introuvable(s) : " + missing.join(", ") };
          if (values.length && values.every((v) => !v)) return { ok: false, error: "valeur(s) de la ligne vide(s) pour l'URL" };
          url = built;
        }
        const res = await navigateTabAndWait(tabId, url, step.gotoWait !== false, step.gotoTimeout);
        if (!res.ok) return { ok: false, error: res.error };
        return {
          ok: true,
          info: res.timedOut ? "page toujours en chargement après " + step.gotoTimeout + " ms — on continue" : scnTrunc(url, 60),
          warn: !!res.timedOut
        };
      }
      case "click": {
        if (!step.clickSelector) return { ok: false, error: "sélecteur manquant" };
        const [{ result }] = await chrome.scripting.executeScript({
          target: { tabId },
          func: scnClickInjected,
          args: [{ selector: step.clickSelector, timeoutMs: step.clickTimeout }]
        });
        return result || { ok: false, error: "pas de réponse de la page" };
      }
      case "wait": {
        if (step.waitMode === "delay") {
          await scnSleep(step.waitMs);
          return { ok: true, info: `${step.waitMs} ms` };
        }
        if (!step.waitSelector) return { ok: false, error: "sélecteur manquant" };
        const [{ result }] = await chrome.scripting.executeScript({
          target: { tabId },
          func: scnWaitInjected,
          args: [{ selector: step.waitSelector, timeoutMs: step.waitTimeout, mode: step.waitMode }]
        });
        return result || { ok: false, error: "pas de réponse de la page" };
      }
      case "cond": {
        let match;
        const condRow = (rowIdx !== null && rowIdx !== undefined) ? (state.rows[rowIdx] || []) : null;
        if (step.condSource === "page") {
          if (!step.condSelector) return { ok: false, error: "sélecteur manquant" };
          const pageVal = resolveRowTemplate(step.condPageVal, condRow);
          const [{ result }] = await chrome.scripting.executeScript({
            target: { tabId },
            func: scnCheckInjected,
            args: [{ selector: step.condSelector, op: step.condPageOp, value: pageVal }]
          });
          if (!result || !result.ok) return { ok: false, error: result ? result.error : "pas de réponse de la page" };
          match = result.match;
        } else {
          if (condRow === null) return { ok: false, error: "aucune ligne active pour la condition Excel" };
          const idx = colIndexByName(step.condCol);
          if (idx < 0) return { ok: false, error: `colonne introuvable (${step.condCol})` };
          match = cellMatches(getCellByIndex(condRow, idx), step.condOp, resolveRowTemplate(step.condVal, condRow));
        }
        if (!match) return { ok: true, info: "condition non remplie → on continue" };
        if (step.condAction === "stop") return { ok: true, stop: true, info: "condition remplie → arrêt du scénario" };
        const n = Math.max(1, step.condSkip || 1);
        return { ok: true, skip: n, info: `condition remplie → saute ${n} étape(s)` };
      }
      case "pdfcheck": {
        const row = (rowIdx !== null && rowIdx !== undefined) ? (state.rows[rowIdx] || []) : null;
        const { doc, error, warn } = tbGetDocForStep(step, row);
        if (error) return { ok: false, error };
        const values = tbFieldValues(doc, step.pdfField);
        const expected = resolveRowTemplate(step.pdfVal ?? "", row);
        const match = tbFieldMatches(values, step.pdfOp, expected);
        const shown = values.length ? scnTrunc(values.join(" | "), 60) : "(vide)";
        if (match) {
          return { ok: true, info: `OK — ${doc.name} : ${shown}${warn ? " ⚠ " + warn : ""}` };
        }
        const opLabel = (OPERATORS.find((o) => o.v === step.pdfOp) || {}).t || step.pdfOp;
        const msg = `écart PDF (${doc.name}) — « ${tbFieldLabel(step.pdfField)} » = ${shown}, attendu : ${opLabel} « ${scnTrunc(expected, 40)} »`;
        switch (step.pdfMissAction) {
          case "warn": return { ok: true, warn: true, info: msg + " → on continue" };
          case "stop": return { ok: true, stop: true, info: msg + " → arrêt du scénario" };
          case "skip": {
            const n = Math.max(1, step.pdfMissSkip || 1);
            return { ok: true, skip: n, info: msg + ` → saute ${n} étape(s)` };
          }
          default: return { ok: false, error: msg };
        }
      }
      case "pdfwrite": {
        if (rowIdx === null || rowIdx === undefined || !state.rows.length) {
          return { ok: false, error: "aucune ligne active pour écrire la valeur" };
        }
        const row = state.rows[rowIdx] = state.rows[rowIdx] || [];
        const { doc, error, warn } = tbGetDocForStep(step, row);
        if (error) return { ok: false, error };
        const values = tbFieldValues(doc, step.pdfField);
        const val = values.join(" | ");
        const ref = step.pdfTargetCol === "__other__" ? (step.pdfNewCol || "").trim() : step.pdfTargetCol;
        if (!ref) return { ok: false, error: "colonne cible manquante" };
        const existed = colIndexByName(ref) >= 0;
        const idx = resolveOutputTarget(ref, {});
        if (idx < 0) return { ok: false, error: `colonne cible introuvable (${ref})` };
        setCellByIndex(row, idx, val);
        persistSession();
        if (existed) { renderPreview(); renderSelectedRowFields(); }
        else renderColumns(); // nouvelle colonne : rafraîchit chips + selects
        return {
          ok: true,
          info: `« ${ref} » ← ${val ? scnTrunc(val, 50) : "(vide)"}${warn ? " ⚠ " + warn : ""}`,
          warn: !val
        };
      }
      case "sigeo": return await execSigeoStep(step, rowIdx, tabId);

      case "batchedit": {
        if (!step.batchButton) return { ok: false, error: "sélecteur du bouton manquant" };

        const size = Math.max(1, parseInt(step.batchSize, 10) || 10);
        const waitMs = Math.max(0, parseInt(step.batchWaitMs, 10) || 0);
        const maxRounds = Math.max(1, parseInt(step.batchMaxRounds, 10) || 50);

        let offset = 0;
        let rounds = 0;
        let total = null;
        let lotsSansDl = 0;

        // Mode d'attente basé sur le téléchargement : la permission peut
        // avoir été révoquée depuis la configuration de l'étape.
        let dlMode = (step.batchWaitMode === "start" || step.batchWaitMode === "complete")
          ? step.batchWaitMode : null;
        if (dlMode && !(await hasDownloadsPermission())) {
          dlMode = null;
          scnLog("   accès aux téléchargements non accordé — repli sur le délai fixe", "skip");
        }
        if (dlMode && opts.soloTab) {
          // chrome.downloads ne dit pas quel onglet a déclenché un fichier :
          // deux onglets qui guettent en même temps peuvent se voler
          // l'événement. On prévient, sans bloquer.
          scnLog("   ⚠ mode multi-onglets : l'attente du téléchargement peut confondre les onglets", "skip");
        }
        const dlTimeout = Math.max(1000, parseInt(step.batchDlTimeoutMs, 10) || 30000);
        const dlSettle = Math.max(0, parseInt(step.batchDlSettleMs, 10) || 0);

        while (rounds < maxRounds) {
          if (scnStopRequested) return { ok: true, stop: true, info: `arrêté après ${rounds} lot(s)` };

          // Coche la tranche suivante et décoche tout le reste.
          const sel = await sendMessageToTab(tabId, {
            action: "batchSelectSlice",
            config: {
              scopeSelector: step.batchScope,
              matchFilter: step.batchFilter,
              offset,
              count: size
            }
          });

          if (!sel) return { ok: false, error: "pas de réponse de la page" };
          if (!sel.success) return { ok: false, error: sel.error || "sélection impossible" };
          if (sel.scopeFound === false) return { ok: false, error: "tableau des cases introuvable" };

          if (total === null) {
            total = sel.total;
            if (!total) return { ok: false, error: "aucune case à cocher trouvée dans le périmètre" };
          }
          if (!sel.selected) break; // plus rien à traiter

          // Guet armé AVANT le clic : un téléchargement instantané serait
          // manqué par un écouteur posé après.
          const dlWatch = dlMode ? watchNextDownload(dlMode, dlTimeout) : null;

          const [{ result: clicked }] = await chrome.scripting.executeScript({
            target: { tabId },
            func: scnClickInjected,
            args: [{ selector: step.batchButton, timeoutMs: 5000 }]
          });
          if (!clicked || !clicked.ok) {
            if (dlWatch) dlWatch.cancel();
            return {
              ok: false,
              error: `lot ${rounds + 1} (éléments ${sel.from}-${sel.to}) : ${(clicked && clicked.error) || "clic impossible"}`
            };
          }

          rounds++;
          offset = sel.to;
          scnLog(`   lot ${rounds} : éléments ${sel.from}-${sel.to} / ${total} → clic`, "");

          if (dlWatch) {
            const dl = await dlWatch;
            const nom = dl.filename ? " — " + dl.filename : "";
            if (dl.interrupted) {
              lotsSansDl++;
              scnLog(`   lot ${rounds} : téléchargement interrompu${nom}`, "err");
            } else if (dl.started) {
              scnLog(`   lot ${rounds} : téléchargement ${dl.completed ? "terminé" : "démarré"}${nom}`, "ok");
            } else {
              lotsSansDl++;
              scnLog(`   lot ${rounds} : aucun téléchargement après ${dlTimeout} ms — on continue`, "err");
            }
          }

          if (!sel.remaining) break;      // dernier lot : pas d'attente inutile
          if (dlMode) { if (dlSettle) await scnSleep(dlSettle); }
          else if (waitMs) await scnSleep(waitMs);
        }

        if (total === null) return { ok: false, error: "aucune case à cocher trouvée" };

        const reste = Math.max(0, total - offset);
        const manque = lotsSansDl ? `, ${lotsSansDl} lot(s) sans téléchargement` : "";
        if (reste) {
          return {
            ok: true,
            warn: true,
            info: `${rounds} lot(s) — arrêt sur le garde-fou « lots max », ${reste} élément(s) non traité(s)${manque}`
          };
        }
        return {
          ok: true,
          warn: lotsSansDl > 0,
          info: `${rounds} lot(s) de ${size} sur ${total} élément(s)${manque}`
        };
      }
    }
    return { ok: false, error: "type d'étape inconnu" };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

// Exécute toutes les étapes pour une ligne. prefix = préfixe de log (mode boucle).
async function runScenarioForRow(rowIdx, tabId, steps, prefix = "", opts = {}) {
  for (let i = 0; i < steps.length; i++) {
    if (scnStopRequested) return { stopped: true };
    const step = steps[i];
    const label = `${prefix}Étape ${i + 1} — ${scnStepLabel(step)}`;
    const res = await execScenarioStep(step, rowIdx, tabId, opts);
    if (!res.ok) {
      scnLog(`✗ ${label} : ${res.error}`, "err");
      return { ok: false };
    }
    scnLog(
      `${res.stop || res.skip ? "→" : "✓"} ${label}${res.info ? " : " + res.info : ""}`,
      res.stop || res.skip || res.warn ? "skip" : "ok"
    );
    if (!prefix) scnSetProgress(i + 1, steps.length);
    if (res.stop) return { ok: true };
    if (res.skip) i += res.skip;
  }
  return { ok: true };
}

function scenarioNeedsRow(steps) {
  // Une étape "fill" ne nécessite une ligne que si elle s'appuie sur un
  // mapping colonnes → inputs ; avec uniquement des champs personnalisés
  // (valeurs fixes), elle fonctionne sans donnée Excel chargée.
  return steps.some((s) =>
    (s.type === "fill" && Object.keys(state.mapping).length > 0) ||
    (s.type === "cond" && s.condSource === "excel") ||
    (s.type === "goto" && /\{[^{}]+\}/.test(s.gotoUrl || "")) ||
    (s.type === "pdfwrite") ||
    (s.type === "pdfcheck" && (/\{[^{}]+\}/.test(s.pdfVal || "") ||
      (s.pdfDocMode === "match" && /\{[^{}]+\}/.test(s.pdfDocMatch || "")))) ||
    (s.type === "sigeo" && ([s.sigeoAddressId, s.sigeoTable, s.sigeoPays, s.sigeoCp, s.sigeoVille,
      s.sigeoVoie, s.sigeoNumVoie, s.sigeoBp, s.sigeoComplt, s.sigeoCompNom]
      .some((v) => /\{[^{}]+\}/.test(v || "")) || !!s.sigeoResultCol))
  );
}

function scnBegin() {
  scnRunning = true;
  scnStopRequested = false;
  $("runScenarioBtn").disabled = true;
  $("runScenarioLoopBtn").disabled = true;
  $("runScenarioMultiBtn").disabled = true;
  $("mtStartBtn").disabled = true;
  $("addStepBtn").disabled = true;
  $("stopScenarioBtn").disabled = false;
  $("scnLog").innerHTML = "";
}

function scnFinish() {
  scnRunning = false;
  $("runScenarioBtn").disabled = false;
  $("runScenarioLoopBtn").disabled = false;
  $("runScenarioMultiBtn").disabled = false;
  if (!mtState.active) $("mtStartBtn").disabled = false;
  $("addStepBtn").disabled = false;
  $("stopScenarioBtn").disabled = true;
  persistSession();
}

$("stopScenarioBtn").addEventListener("click", () => { scnStopRequested = true; });

// Exécution sur la ligne active
$("runScenarioBtn").addEventListener("click", async () => {
  if (scnRunning) return;
  const steps = getScenarioSteps();
  if (!steps.length) { showStatus("Ajoute au moins une étape au scénario.", "error"); return; }
  if (scenarioNeedsRow(steps) && (state.selectedRowIdx === null || !state.rows.length)) {
    showStatus("Sélectionne d'abord une ligne à saisir.", "error");
    return;
  }
  saveCustomFields();

  let tabId;
  try { tabId = await resolveTargetTabId(); }
  catch (err) { showStatus(err.message, "error"); return; }

  scnBegin();
  scnSetProgress(0, steps.length);
  scnLog(state.selectedRowIdx !== null ? `— Scénario : ligne L${state.selectedRowIdx + 1} —` : "— Scénario —");
  const res = await runScenarioForRow(state.selectedRowIdx, tabId, steps);
  if (res.stopped) scnLog("Arrêté par l'utilisateur.", "skip");
  else scnLog(res.ok ? "Scénario terminé." : "Scénario interrompu (erreur).", res.ok ? "ok" : "err");
  scnFinish();
});

// Exécution en boucle sur une plage de lignes
$("runScenarioLoopBtn").addEventListener("click", async () => {
  if (scnRunning) return;
  const steps = getScenarioSteps();
  if (!steps.length) { showStatus("Ajoute au moins une étape au scénario.", "error"); return; }
  if (!state.rows.length) { showStatus("Charge d'abord des données (onglet Données).", "error"); return; }
  saveCustomFields();

  const startRowInput = parseInt($("scnStartRow").value, 10) || (dataStartIdx() + 1);
  const endRowInput = $("scnEndRow").value.trim();
  const startIdx = Math.max(0, startRowInput - 1);
  const endIdx = Math.min(state.rows.length - 1, endRowInput ? parseInt(endRowInput, 10) - 1 : state.rows.length - 1);
  if (endIdx < startIdx) { showStatus("Plage de lignes invalide.", "error"); return; }
  const rowDelayMs = parseInt($("scnRowDelayMs").value, 10) || 0;

  let tabId;
  try { tabId = await resolveTargetTabId(); }
  catch (err) { showStatus(err.message, "error"); return; }

  scnBegin();
  const total = endIdx - startIdx + 1;
  let done = 0, okCount = 0, errCount = 0, skippedCount = 0;
  const runStart = Date.now();
  scnSetProgress(0, total);
  $("scnEstimate").textContent = "";
  scnUpdateTiming(0, total, runStart);

  for (let idx = startIdx; idx <= endIdx; idx++) {
    if (scnStopRequested) { scnLog("Arrêté par l'utilisateur.", "skip"); break; }
    const row = state.rows[idx] || [];
    if (!row.some((v) => String(v ?? "").trim() !== "")) {
      scnLog(`L${idx + 1} : ligne vide, ignorée.`, "skip");
      skippedCount++; done++; scnSetProgress(done, total); scnUpdateTiming(done, total, runStart);
      continue;
    }
    selectRow(idx); // la ligne active suit la boucle (visible dans l'UI)
    scnLog(`— Ligne L${idx + 1} —`);
    const res = await runScenarioForRow(idx, tabId, steps, `L${idx + 1} · `);
    if (res.stopped) { scnLog("Arrêté par l'utilisateur.", "skip"); break; }
    if (res.ok) okCount++; else errCount++;
    done++; scnSetProgress(done, total); scnUpdateTiming(done, total, runStart);
    if (rowDelayMs > 0 && idx < endIdx && !scnStopRequested) await scnSleep(rowDelayMs);
  }

  const scnDurationMs = Date.now() - runStart;
  $("scnProgressTiming").textContent = `Terminé en ${formatDuration(scnDurationMs)}`;
  scnLog(`Boucle terminée en ${formatDuration(scnDurationMs)} : ${okCount} OK, ${errCount} erreur(s), ${skippedCount} ligne(s) vide(s).`, errCount ? "err" : "ok");
  scnFinish();
  updateScenarioEstimate();
});

/* ---------- Exécution du scénario en parallèle sur plusieurs onglets ----------
   Chaque onglet prend la ligne suivante dans une file commune et déroule
   tout le scénario dessus, indépendamment des autres. Deux précautions
   propres à ce mode :
     - le remplissage est cloisonné (soloTab) : diffusé à tous les onglets
       du site comme en mode simple, un onglet écraserait la ligne des
       autres ;
     - la ligne active de l'interface n'est PAS déplacée : plusieurs
       onglets avancent en même temps, il n'y a plus de « ligne en cours »
       unique à afficher. */

// Attend qu'un onglet ait fini de charger (onglets fraîchement dupliqués).
function waitForTabComplete(tabId, timeoutMs = 15000) {
  return new Promise((resolve) => {
    let done = false;
    const finish = () => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      try { chrome.tabs.onUpdated.removeListener(onUpd); } catch (_) { /* ignoré */ }
      resolve();
    };
    const onUpd = (id, info) => { if (id === tabId && info.status === "complete") finish(); };
    const timer = setTimeout(finish, timeoutMs);
    chrome.tabs.onUpdated.addListener(onUpd);
    chrome.tabs.get(tabId)
      .then((t) => { if (t && t.status === "complete") finish(); })
      .catch(finish);
  });
}

// Réunit `wanted` onglets de travail sur le site cible : ceux déjà ouverts
// d'abord, complétés au besoin par duplication de l'onglet du formulaire.
async function acquireWorkerTabs(wanted) {
  const targetTab = await chrome.tabs.get(await resolveTargetTabId());
  const url = targetTab.url || "";
  if (!/^https?:/.test(url)) {
    throw new Error("Page protégée : ouvre d'abord le site cible (ou clique 🎯 sur le formulaire).");
  }
  const origin = new URL(url).origin;

  const existing = (await chrome.tabs.query({ url: origin + "/*" }))
    .filter((t) => /^https?:/.test(t.url || ""));
  if (!existing.some((t) => t.id === targetTab.id)) existing.unshift(targetTab);

  const tabIds = existing.slice(0, wanted).map((t) => t.id);
  const created = [];
  while (tabIds.length < wanted) {
    const dup = await chrome.tabs.duplicate(targetTab.id);
    tabIds.push(dup.id);
    created.push(dup.id);
  }

  // Un onglet encore en chargement ferait échouer la première étape.
  await Promise.all(created.map((id) => waitForTabComplete(id)));
  return { tabIds, origin, created };
}

$("runScenarioMultiBtn").addEventListener("click", async () => {
  if (scnRunning) return;
  if (mtState.active) {
    showStatus("Arrête d'abord la saisie multi-onglets (remplissage assisté).", "error");
    return;
  }

  const steps = getScenarioSteps();
  if (!steps.length) { showStatus("Ajoute au moins une étape au scénario.", "error"); return; }
  if (!state.rows.length) { showStatus("Charge d'abord des données (onglet Données).", "error"); return; }
  saveCustomFields();

  const startRowInput = parseInt($("scnStartRow").value, 10) || (dataStartIdx() + 1);
  const endRowInput = $("scnEndRow").value.trim();
  const startIdx = Math.max(0, startRowInput - 1);
  const endIdx = Math.min(state.rows.length - 1, endRowInput ? parseInt(endRowInput, 10) - 1 : state.rows.length - 1);
  if (endIdx < startIdx) { showStatus("Plage de lignes invalide.", "error"); return; }
  const rowDelayMs = parseInt($("scnRowDelayMs").value, 10) || 0;

  const queue = [];
  for (let i = startIdx; i <= endIdx; i++) queue.push(i);
  const total = queue.length;

  // Inutile d'ouvrir plus d'onglets que de lignes à traiter.
  const wanted = Math.min(
    Math.min(Math.max(parseInt($("mtTabCount").value, 10) || 2, 2), 10),
    total
  );
  if (wanted < 2) {
    showStatus("Une seule ligne à traiter : utilise « Demarrer le scénario ».", "error");
    return;
  }

  let tabIds;
  try {
    ({ tabIds } = await acquireWorkerTabs(wanted));
  } catch (err) {
    showStatus(err.message, "error");
    return;
  }

  scnBegin();
  let done = 0, okCount = 0, errCount = 0, skippedCount = 0;
  const runStart = Date.now();
  scnSetProgress(0, total);
  $("scnEstimate").textContent = "";
  scnUpdateTiming(0, total, runStart);
  scnLog(`— Scénario en parallèle : ${total} ligne(s) réparties sur ${tabIds.length} onglet(s) —`);

  const tick = () => {
    done++;
    scnSetProgress(done, total);
    scnUpdateTiming(done, total, runStart);
  };

  // Un « worker » par onglet : il pioche dans la file commune jusqu'à
  // épuisement. La file étant un simple tableau et JS mono-thread, le
  // shift() ne peut pas servir deux fois la même ligne.
  const worker = async (tabId, n) => {
    const tag = `O${n}`;
    while (!scnStopRequested) {
      const idx = queue.shift();
      if (idx === undefined) return;

      const row = state.rows[idx] || [];
      if (!row.some((v) => String(v ?? "").trim() !== "")) {
        scnLog(`${tag} · L${idx + 1} : ligne vide, ignorée.`, "skip");
        skippedCount++;
        tick();
        continue;
      }

      const res = await runScenarioForRow(idx, tabId, steps, `${tag} · L${idx + 1} · `, { soloTab: true });
      if (res.stopped) return;
      if (res.ok) okCount++; else errCount++;
      tick();

      if (rowDelayMs > 0 && !scnStopRequested) await scnSleep(rowDelayMs);
    }
  };

  await Promise.all(tabIds.map((id, i) => worker(id, i + 1).catch((err) => {
    scnLog(`O${i + 1} : onglet interrompu — ${err.message}`, "err");
    errCount++;
  })));

  const durMs = Date.now() - runStart;
  $("scnProgressTiming").textContent = `Terminé en ${formatDuration(durMs)}`;
  if (scnStopRequested) scnLog("Arrêté par l'utilisateur.", "skip");
  const reste = queue.length;
  scnLog(
    `Parallèle terminé en ${formatDuration(durMs)} : ${okCount} OK, ${errCount} erreur(s), ` +
    `${skippedCount} ligne(s) vide(s)${reste ? `, ${reste} ligne(s) non traitée(s)` : ""}.`,
    errCount ? "err" : "ok"
  );
  scnFinish();
  updateScenarioEstimate();
});

/* ================== 5ter. TOOLBOX PDF (BÊTA) ================== */
// Analyse locale de PDF (pdf.js) : extraction du texte page par page avec
// reconstruction des espaces (certains PDF — polices mal encodées — perdent
// les espaces), détection de champs par motifs intégrés (SIRET, n° ADEME,
// dates, emails…) + libellés « Libellé : valeur » + motifs personnalisés.
// Les documents ne sont PAS persistés (session en cours uniquement).

const tbState = {
  docs: [],          // [{id, name, size, pages, text, fields:[{key,label,values}], error, loading}]
  activeDocId: null,
  sendKeys: new Set(), // champs cochés pour « Vers la ligne active »
  patterns: []       // motifs perso [{name, kind:"label"|"regex", value}]
};
let tbDocSeq = 1;

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = chrome.runtime.getURL("lib/pdf.worker.min.js");
}

// Normalisation : minuscules + accents retirés (longueur conservée) — permet
// de faire correspondre libellés/valeurs sans se soucier des accents.
function tbNorm(s) {
  return String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[’‘]/g, "'").toLowerCase();
}

function tbEscapeRegex(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

// Motifs intégrés (regex, valeurs multiples possibles).
const TB_BUILTIN_PATTERNS = [
  { key: "ademe", label: "N° ADEME / DPE", re: "\\b(\\d{4}[A-Z]\\d{7}[A-Z])\\b" },
  { key: "siret", label: "SIRET", re: "\\b(\\d{3}[ .]?\\d{3}[ .]?\\d{3}[ .]?\\d{5})\\b" },
  { key: "siren", label: "SIREN", re: "\\b(\\d{3}[ .]?\\d{3}[ .]?\\d{3})\\b(?![ .]?\\d)" },
  { key: "tva", label: "TVA intracom.", re: "\\b(FR ?\\d{2} ?\\d{9})\\b" },
  { key: "iban", label: "IBAN", re: "\\b(FR\\d{2}(?: ?[A-Z0-9]{4}){5} ?[A-Z0-9]{0,3})\\b" },
  { key: "date", label: "Dates", re: "\\b(\\d{2}/\\d{2}/\\d{4})\\b" },
  { key: "email", label: "Emails", re: "\\b([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,})\\b" },
  { key: "tel", label: "Téléphones", re: "\\b(0[1-9](?:[ .]?\\d{2}){4})\\b" },
  { key: "cpville", label: "CP + Ville", re: "\\b(\\d{5} ?[A-ZÀ-Ü][A-ZÀ-Üa-zà-ü' -]{2,})" },
  { key: "surface", label: "Surfaces (m²)", re: "\\b(\\d+(?:[.,]\\d+)?) ?m²" },
  // Étiquette énergie DPE — motifs tolérants aux erreurs d'OCR (² → ?/*/2,
  // kWh → kWiv/kWhimt…, valeurs « 325 |10* » au-dessus des unités).
  { key: "dpe_conso", label: "Conso énergie (kWh/m²/an)", re: [
    "\\b(\\d{1,4})\\s*[kK][WwVv]\\S{0,6}/an",
    { re: "\\b(\\d{2,4})\\s*[|Il!1]\\s*\\d{1,3}\\s*\\*", group: 1 },
    { re: "\\b(\\d{1,4})\\s*[|Il!]\\s*\\d{1,4}\\s*\\*?\\s*\\n\\s*[kK][WwVv]", group: 1 }
  ] },
  { key: "dpe_emission", label: "Émissions (kgCO₂/m²/an)", re: [
    "\\b(\\d{1,4})\\s*[kK][gq]\\s?[CG]O\\S{0,2}\\s*/\\s*m\\S{0,3}\\s*/?\\s*an",
    { re: "\\b\\d{2,4}\\s*[|Il!1]\\s*(\\d{1,3})\\s*\\*", group: 1 },
    { re: "\\b\\d{1,4}\\s*[|Il!]\\s*(\\d{1,4})\\s*\\*?\\s*\\n\\s*[kK][WwVv]", group: 1 }
  ] },
  { key: "dpe_co2_total", label: "CO₂ émis (kg/an)", re: ["[ée]met\\s*(\\d[\\d ]{0,6})\\s*kg\\s*de\\s*CO"] }
];

// Libellés intégrés : cherche « Libellé : valeur » ligne par ligne.
const TB_BUILTIN_LABELS = [
  { key: "proprietaire", label: "Propriétaire" },
  { key: "adresse", label: "Adresse" },
  { key: "type_bien", label: "Type de bien" },
  { key: "surface_habitable", label: "Surface habitable" },
  { key: "annee_construction", label: "Année de construction" },
  { key: "etabli_le", label: "Établi le" },
  { key: "valable_jusquau", label: "Valable jusqu'au" },
  { key: "diagnostiqueur", label: "Diagnostiqueur" },
  { key: "certification", label: "N° de certification" }
];

// Regex de libellé : accents/espaces tolérés, deux-points requis.
function tbLabelRegex(label, valueRequired) {
  const esc = tbEscapeRegex(tbNorm(label)).replace(/ +/g, "[ ]?").replace(/'/g, "'?");
  return new RegExp("(^|[^a-z0-9])" + esc + "[ ]*:[ ]*" + (valueRequired ? "(.+)$" : "$"), "i");
}

/* ---------- Extraction du texte (pdf.js) ---------- */

// Reconstruit le texte d'un PDF : regroupe les items par ligne (y proche),
// trie par x et réinsère les espaces d'après les écarts horizontaux.
async function tbExtractTextFromPdf(arrayBuffer) {
  if (!window.pdfjsLib) throw new Error("pdf.js non chargé (lib/pdf.min.js manquant)");
  const doc = await pdfjsLib.getDocument({ data: arrayBuffer, useSystemFonts: true }).promise;
  const pageTexts = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const tc = await page.getTextContent();
    const rows = new Map();
    for (const it of tc.items) {
      if (!it.str) continue;
      const key = Math.round(it.transform[5] / 2) * 2; // tolérance verticale ~2pt
      if (!rows.has(key)) rows.set(key, []);
      rows.get(key).push({ x: it.transform[4], str: it.str, w: it.width || 0, h: it.height || 0 });
    }
    const lines = [...rows.entries()]
      .sort((a, b) => b[0] - a[0]) // haut → bas
      .map(([, items]) => {
        items.sort((a, b) => a.x - b.x);
        let line = "", endX = null, h = 0;
        for (const it of items) {
          if (endX !== null) {
            const ref = Math.max(h, it.h, 6);
            if (it.x - endX > ref * 0.12) line += " "; // écart → espace
          }
          line += it.str;
          endX = it.x + it.w;
          h = Math.max(h, it.h);
        }
        return line.trim();
      })
      .filter(Boolean);
    pageTexts.push(lines.join("\n"));
  }
  try { doc.destroy(); } catch (_) {}
  return {
    text: pageTexts.join("\n\n"),
    pages: pageTexts.length,
    pageTextLens: pageTexts.map((t) => t.length) // sert au mode OCR « auto »
  };
}

// Texte complet d'un document : couche texte + résultat d'OCR éventuel.
function tbDocFullText(d) {
  if (!d) return "";
  if (!d.ocrText) return d.text || "";
  return (d.text ? d.text + "\n\n" : "") + d.ocrText;
}

/* ---------- Détection des champs ---------- */

function tbAllFieldDefs() {
  const defs = [];
  TB_BUILTIN_LABELS.forEach((l) => defs.push({ key: l.key, label: l.label, kind: "label", value: l.label }));
  TB_BUILTIN_PATTERNS.forEach((p) => defs.push({ key: p.key, label: p.label, kind: "regex", value: p.re }));
  tbState.patterns.forEach((p, i) => {
    if ((p.name || "").trim() && (p.value || "").trim()) {
      defs.push({ key: "custom:" + p.name.trim(), label: p.name.trim(), kind: p.kind === "regex" ? "regex" : "label", value: p.value.trim(), custom: true });
    }
  });
  return defs;
}

function tbExtractFields(text) {
  const lines = text.split("\n");
  const normLines = lines.map(tbNorm);
  const fields = [];

  for (const def of tbAllFieldDefs()) {
    let values = [];
    if (def.kind === "label") {
      const reVal = tbLabelRegex(def.value, true);
      const reEnd = tbLabelRegex(def.value, false);
      for (let i = 0; i < lines.length; i++) {
        let m = normLines[i].match(reVal);
        if (m) {
          // normalisation à longueur constante → les index correspondent à la ligne d'origine
          const start = m.index + m[0].length - m[2].length;
          const v = lines[i].slice(start).trim();
          if (v && !/^[-—.]+$/.test(v)) values.push(v);
        } else if (reEnd.test(normLines[i])) {
          // libellé en fin de ligne → la valeur est sur la ligne suivante
          const next = (lines[i + 1] || "").trim();
          if (next) values.push(next);
        }
      }
    } else {
      // def.value : regex unique, ou liste de variantes ("…" ou {re, group}).
      const variants = (Array.isArray(def.value) ? def.value : [def.value])
        .map((v) => (typeof v === "string" ? { re: v, group: 1 } : v));
      for (const va of variants) {
        let re = null;
        try { re = new RegExp(va.re, "g"); } catch (_) {}
        if (!re) { values.push("(regex invalide)"); continue; }
        let m, guard = 0;
        while ((m = re.exec(text)) && guard++ < 500) {
          const g = va.group || 1;
          values.push(((m[g] !== undefined ? m[g] : m[0]) || "").trim());
          if (m.index === re.lastIndex) re.lastIndex++;
        }
      }
    }
    values = [...new Set(values)].slice(0, 50);
    fields.push({ key: def.key, label: def.label, values });
  }
  return fields;
}

/* ---------- Chargement des documents ---------- */

$("tbFileInput").addEventListener("change", async (e) => {
  const files = Array.from(e.target.files || []);
  e.target.value = ""; // permet de recharger le même fichier
  if (!files.length) return;
  for (const file of files) {
    const entry = { id: "d" + tbDocSeq++, name: file.name, size: file.size, pages: 0, text: "", fields: [], error: "", loading: true, bytes: null, pageTextLens: null, ocrText: "", ocrPages: null };
    tbState.docs.push(entry);
    tbState.activeDocId = entry.id;
    tbRenderDocs();
    try {
      const buf = await file.arrayBuffer();
      entry.bytes = new Uint8Array(buf); // conservé pour l'OCR (rendu des pages)
      const { text, pages, pageTextLens } = await tbExtractTextFromPdf(entry.bytes.slice());
      entry.text = text;
      entry.pages = pages;
      entry.pageTextLens = pageTextLens;
      entry.fields = tbExtractFields(text);
      if (!text.trim()) entry.error = "aucun texte trouvé (PDF scanné/image) — lance l'OCR ci-dessous";
    } catch (err) {
      entry.error = err && err.message ? err.message : String(err);
    }
    entry.loading = false;
    tbRenderDocs();
    tbRenderActiveDoc();
  }
  updateDoneMarkers();
  const okCount = tbState.docs.filter((d) => !d.error && !d.loading).length;
  showStatus(`${okCount} PDF analysé(s).`, "success");
});

function tbActiveDoc() {
  return tbState.docs.find((d) => d.id === tbState.activeDocId) || null;
}

function tbRenderDocs() {
  const wrap = $("tbDocsList");
  wrap.innerHTML = "";
  if (!tbState.docs.length) {
    wrap.innerHTML = '<p class="hint">Aucun document chargé.</p>';
    return;
  }
  tbState.docs.forEach((d) => {
    const div = document.createElement("div");
    div.className = "tb-doc" + (d.id === tbState.activeDocId ? " active" : "");
    div.innerHTML = `
      <svg class="icon icon-sm"><use href="#icon-file"/></svg>
      <span class="tb-doc-name" title="${escapeAttr(d.name)}">${escapeHtml(d.name)}</span>
      ${d.loading ? '<span class="tb-doc-meta">analyse…</span>'
        : d.error ? `<span class="tb-doc-err" title="${escapeAttr(d.error)}">⚠ erreur</span>`
        : `<span class="tb-doc-meta">${d.pages} p. · ${Math.round(d.size / 1024)} Ko${d.ocrText ? " · OCR ✓" : ""}</span>`}
      <button class="remove-btn" title="Retirer" type="button"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
    `;
    div.addEventListener("click", () => {
      tbState.activeDocId = d.id;
      tbRenderDocs();
      tbRenderActiveDoc();
    });
    div.querySelector(".remove-btn").addEventListener("click", (ev) => {
      ev.stopPropagation();
      tbState.docs = tbState.docs.filter((x) => x.id !== d.id);
      if (tbState.activeDocId === d.id) tbState.activeDocId = tbState.docs.length ? tbState.docs[tbState.docs.length - 1].id : null;
      tbRenderDocs();
      tbRenderActiveDoc();
      updateDoneMarkers();
    });
    wrap.appendChild(div);
  });
}

/* ---------- Rendu des champs + texte ---------- */

function tbRenderActiveDoc() {
  tbRenderFields();
  tbRenderText();
}

function tbRenderFields() {
  const wrap = $("tbFieldsWrap");
  const doc = tbActiveDoc();
  if (!doc || doc.loading) {
    wrap.innerHTML = '<p class="hint">Charge un PDF pour voir les champs détectés.</p>';
    return;
  }
  if (doc.error && !tbDocFullText(doc).trim()) {
    wrap.innerHTML = `<p class="hint">⚠ ${escapeHtml(doc.error)} — essaie le bouton « Lancer l'OCR ».</p>`;
    return;
  }
  const filter = tbNorm($("tbFieldFilter").value.trim());
  const shown = doc.fields.filter((f) =>
    f.values.length &&
    (!filter || tbNorm(f.label).includes(filter) || f.values.some((v) => tbNorm(v).includes(filter)))
  );
  if (!shown.length) {
    wrap.innerHTML = '<p class="hint">Aucun champ détecté' + (filter ? " pour ce filtre" : "") + ".</p>";
    return;
  }
  wrap.innerHTML = "";
  shown.forEach((f) => {
    const div = document.createElement("div");
    div.className = "tb-field";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "tb-field-check";
    cb.title = "Inclure dans « Vers la ligne active »";
    cb.checked = tbState.sendKeys.has(f.key);
    cb.addEventListener("change", () => {
      if (cb.checked) tbState.sendKeys.add(f.key); else tbState.sendKeys.delete(f.key);
    });
    const lab = document.createElement("span");
    lab.className = "tb-field-label";
    lab.textContent = f.label;
    lab.title = f.label;
    const vals = document.createElement("span");
    vals.className = "tb-field-vals";
    f.values.slice(0, 20).forEach((v) => {
      const chip = document.createElement("span");
      chip.className = "tb-val";
      chip.textContent = v.length > 80 ? v.slice(0, 80) + "…" : v;
      chip.title = "Cliquer pour copier : " + v;
      chip.addEventListener("click", async () => {
        try { await navigator.clipboard.writeText(v); showStatus("Valeur copiée.", "success"); }
        catch (_) { showStatus("Copie impossible.", "error"); }
      });
      vals.appendChild(chip);
    });
    if (f.values.length > 20) {
      const more = document.createElement("span");
      more.className = "hint";
      more.textContent = `+${f.values.length - 20}`;
      vals.appendChild(more);
    }
    div.append(cb, lab, vals);
    wrap.appendChild(div);
  });
}

$("tbFieldFilter").addEventListener("input", tbRenderFields);

// Écrit les champs cochés dans la ligne active (colonne du même nom, créée si absente).
$("tbSendRowBtn").addEventListener("click", () => {
  const doc = tbActiveDoc();
  if (!doc || !doc.fields.length) { showStatus("Charge d'abord un PDF.", "error"); return; }
  if (!state.rows.length || state.selectedRowIdx === null) {
    showStatus("Sélectionne d'abord une ligne active (onglet Saisie).", "error");
    return;
  }
  const picked = doc.fields.filter((f) => tbState.sendKeys.has(f.key) && f.values.length);
  if (!picked.length) { showStatus("Coche au moins un champ (case à gauche).", "error"); return; }
  const createdCols = {};
  const row = state.rows[state.selectedRowIdx] = state.rows[state.selectedRowIdx] || [];
  picked.forEach((f) => {
    const idx = resolveOutputTarget(f.label, createdCols);
    setCellByIndex(row, idx, f.values.join(" | "));
  });
  persistSession();
  renderColumns();
  showStatus(`${picked.length} champ(s) écrit(s) dans la ligne L${state.selectedRowIdx + 1}.`, "success");
});

function tbRenderText() {
  const doc = tbActiveDoc();
  const view = $("tbTextView");
  const full = tbDocFullText(doc);
  if (!doc || doc.loading || !full) {
    view.innerHTML = '<span class="hint">Charge un PDF pour voir son texte.</span>';
    $("tbTextSearchInfo").textContent = "";
    return;
  }
  const q = $("tbTextSearch").value;
  if (!q.trim()) {
    view.textContent = full;
    $("tbTextSearchInfo").textContent = "";
    return;
  }
  const normText = tbNorm(full);
  const needle = tbNorm(q);
  let html = "", pos = 0, count = 0;
  let i = normText.indexOf(needle);
  while (i !== -1 && count < 500) {
    html += escapeHtml(full.slice(pos, i)) + "<mark>" + escapeHtml(full.slice(i, i + needle.length)) + "</mark>";
    pos = i + needle.length;
    count++;
    i = normText.indexOf(needle, pos);
  }
  html += escapeHtml(full.slice(pos));
  view.innerHTML = html;
  $("tbTextSearchInfo").textContent = count ? `${count} occurrence(s)` : "aucune occurrence";
  const first = view.querySelector("mark");
  if (first) first.scrollIntoView({ block: "nearest" });
}

$("tbTextSearch").addEventListener("input", tbRenderText);

$("tbCopyTextBtn").addEventListener("click", async () => {
  const doc = tbActiveDoc();
  const full = tbDocFullText(doc);
  if (!full) { showStatus("Aucun texte à copier.", "error"); return; }
  try { await navigator.clipboard.writeText(full); showStatus("Texte du PDF copié.", "success"); }
  catch (_) { showStatus("Copie impossible.", "error"); }
});

/* ---------- OCR (bêta) — pages scannées / images ---------- */
// Tesseract.js embarqué (lib/ocr/, 100 % local). Les pages sont rendues en
// image par pdf.js puis reconnues ; le texte OCR s'ajoute à la couche texte
// et les champs sont re-détectés (utile pour l'étiquette énergie des DPE).

let tbOcrWorker = null;
let tbOcrRunning = false;
let tbOcrStopRequested = false;
const TB_OCR_HINT_DEFAULT = $("tbOcrStatus") ? $("tbOcrStatus").textContent : "";

function tbSetOcrStatus(msg) {
  $("tbOcrStatus").textContent = msg || TB_OCR_HINT_DEFAULT;
}

async function tbGetOcrWorker() {
  if (tbOcrWorker) return tbOcrWorker;
  if (!window.Tesseract) throw new Error("Tesseract non chargé (fichiers lib/ocr/ manquants)");
  tbOcrWorker = await Tesseract.createWorker("fra", 1, {
    workerPath: chrome.runtime.getURL("lib/ocr/worker.min.js"),
    corePath: chrome.runtime.getURL("lib/ocr/core"),
    langPath: chrome.runtime.getURL("lib/ocr/lang"),
    workerBlobURL: false,   // requis en extension MV3 (pas de worker blob:)
    cacheMethod: "none"
  });
  return tbOcrWorker;
}

async function tbRunOcrOnActiveDoc() {
  if (tbOcrRunning) { tbOcrStopRequested = true; tbSetOcrStatus("Arrêt demandé…"); return; }
  const doc = tbActiveDoc();
  if (!doc || doc.loading) { showStatus("Charge et sélectionne d'abord un PDF.", "error"); return; }
  if (!doc.bytes) { showStatus("Données du PDF indisponibles — recharge le fichier.", "error"); return; }

  tbOcrRunning = true;
  tbOcrStopRequested = false;
  const btn = $("tbOcrBtn");
  btn.innerHTML = '<svg class="icon icon-sm"><use href="#icon-stop"/></svg> Arrêter l\'OCR';
  const mode = $("tbOcrMode").value;
  let pdf = null;
  let done = 0;
  try {
    tbSetOcrStatus("Initialisation de l'OCR (première fois : quelques secondes)…");
    const worker = await tbGetOcrWorker();
    pdf = await pdfjsLib.getDocument({ data: doc.bytes.slice(), useSystemFonts: true }).promise;

    // Pages cibles : celles (quasi) sans texte, ou toutes si demandé.
    const targets = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const len = doc.pageTextLens ? (doc.pageTextLens[p - 1] || 0) : 0;
      if (mode === "all" || len < 60) targets.push(p);
    }
    if (!targets.length) {
      tbSetOcrStatus("Toutes les pages contiennent déjà du texte — choisis « toutes les pages » pour forcer l'OCR (ex. étiquette énergie en image).");
      return;
    }

    const parts = doc.ocrPages || {};
    for (const p of targets) {
      if (tbOcrStopRequested) break;
      tbSetOcrStatus(`OCR : page ${p} — ${done + 1}/${targets.length} (~5-20 s/page)…`);
      const page = await pdf.getPage(p);
      let vp = page.getViewport({ scale: 1 });
      const scale = Math.min(3, Math.max(1.5, 1700 / vp.width));
      vp = page.getViewport({ scale });
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(vp.width);
      canvas.height = Math.ceil(vp.height);
      await page.render({ canvasContext: canvas.getContext("2d", { willReadFrequently: true }), viewport: vp }).promise;
      const { data } = await worker.recognize(canvas);
      parts[p] = (data.text || "").trim();
      canvas.width = canvas.height = 0; // libère la mémoire
      done++;
    }

    doc.ocrPages = parts;
    doc.ocrText = Object.keys(parts).map(Number).sort((a, b) => a - b)
      .map((p) => `— OCR page ${p} —\n${parts[p]}`).join("\n\n");
    if (doc.error && doc.ocrText.trim()) doc.error = ""; // le PDF scanné a maintenant du texte
    doc.fields = tbExtractFields(tbDocFullText(doc));
    tbRenderDocs();
    tbRenderActiveDoc();
    updateDoneMarkers();
    tbSetOcrStatus(tbOcrStopRequested
      ? `OCR interrompu : ${done}/${targets.length} page(s) traitée(s).`
      : `OCR terminé : ${done} page(s). Champs re-détectés.`);
  } catch (e) {
    tbSetOcrStatus("Erreur OCR : " + (e && e.message ? e.message : e));
    try { if (tbOcrWorker) tbOcrWorker.terminate(); } catch (_) {}
    tbOcrWorker = null; // retentera une initialisation propre la prochaine fois
  } finally {
    try { if (pdf) pdf.destroy(); } catch (_) {}
    tbOcrRunning = false;
    tbOcrStopRequested = false;
    btn.innerHTML = '<svg class="icon icon-sm"><use href="#icon-search"/></svg> Lancer l\'OCR';
  }
}

$("tbOcrBtn").addEventListener("click", tbRunOcrOnActiveDoc);

/* ---------- Motifs personnalisés ---------- */

function tbPersistPatterns() {
  chrome.storage.local.set({ tbPatterns: tbState.patterns });
}

function tbReextractAll() {
  tbState.docs.forEach((d) => {
    if (!d.loading && tbDocFullText(d)) d.fields = tbExtractFields(tbDocFullText(d));
  });
  tbRenderFields();
  tbRefreshFieldSelects();
}

function tbRenderPatterns() {
  const wrap = $("tbPatternsList");
  wrap.innerHTML = "";
  tbState.patterns.forEach((p, i) => {
    const div = document.createElement("div");
    div.className = "tb-pattern-item";
    div.innerHTML = `
      <input type="text" class="tb-pat-name" placeholder="nom du champ" value="${escapeAttr(p.name || "")}" />
      <select class="tb-pat-kind">
        <option value="label">Libellé :</option>
        <option value="regex">Regex</option>
      </select>
      <input type="text" class="tb-pat-value" placeholder="ex : Référence cadastrale — ou — \\bMG\\d{7}\\b" value="${escapeAttr(p.value || "")}" />
      <button class="remove-btn" title="Supprimer" type="button"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
    `;
    div.querySelector(".tb-pat-kind").value = p.kind === "regex" ? "regex" : "label";
    const save = () => {
      tbState.patterns[i] = {
        name: div.querySelector(".tb-pat-name").value,
        kind: div.querySelector(".tb-pat-kind").value,
        value: div.querySelector(".tb-pat-value").value
      };
      tbPersistPatterns();
      tbReextractAll();
    };
    div.querySelectorAll("input, select").forEach((el) => el.addEventListener("change", save));
    div.querySelector(".remove-btn").addEventListener("click", () => {
      tbState.patterns.splice(i, 1);
      tbPersistPatterns();
      tbRenderPatterns();
      tbReextractAll();
    });
    wrap.appendChild(div);
  });
}

$("tbAddPatternBtn").addEventListener("click", () => {
  tbState.patterns.push({ name: "", kind: "label", value: "" });
  tbRenderPatterns();
  $("tbPatternsList").querySelector(".tb-pattern-item:last-child .tb-pat-name").focus();
});

/* ---------- Intégration au scénario ---------- */

// Remplit un <select> avec les champs PDF connus (+ texte complet).
function tbFillFieldSelect(sel, current) {
  sel.innerHTML = "";
  const defs = tbAllFieldDefs();
  defs.forEach((d) => {
    const opt = document.createElement("option");
    opt.value = d.key;
    opt.textContent = d.label + (d.custom ? " (perso)" : "");
    sel.appendChild(opt);
  });
  const optText = document.createElement("option");
  optText.value = "__text";
  optText.textContent = "Texte complet du PDF";
  sel.appendChild(optText);
  if (current) {
    if (![...sel.options].some((o) => o.value === current)) {
      const opt = document.createElement("option");
      opt.value = current;
      opt.textContent = "⚠ " + current;
      sel.appendChild(opt);
    }
    sel.value = current;
  }
}

function tbRefreshFieldSelects() {
  document.querySelectorAll(".scn-pdf-field").forEach((sel) => tbFillFieldSelect(sel, sel.value));
}

function tbFieldLabel(key) {
  if (key === "__text") return "texte complet";
  const def = tbAllFieldDefs().find((d) => d.key === key);
  return def ? def.label : (key || "?");
}

// Valeurs d'un champ pour un document.
function tbFieldValues(doc, key) {
  if (key === "__text") {
    const full = tbDocFullText(doc);
    return full ? [full] : [];
  }
  const f = (doc.fields || []).find((x) => x.key === key);
  return f ? f.values : [];
}

// Test d'un champ multi-valeurs : equals/contains = au moins une valeur
// correspond ; not_* = aucune ; empty / not_empty sur l'ensemble.
function tbFieldMatches(values, op, expected) {
  const nonEmpty = values.filter((v) => String(v).trim() !== "");
  switch (op) {
    case "empty": return nonEmpty.length === 0;
    case "not_empty": return nonEmpty.length > 0;
    case "equals": return nonEmpty.some((v) => cellMatches(v, "equals", expected));
    case "contains": return nonEmpty.some((v) => cellMatches(v, "contains", expected));
    case "not_equals": return !nonEmpty.some((v) => cellMatches(v, "equals", expected));
    case "not_contains": return !nonEmpty.some((v) => cellMatches(v, "contains", expected));
  }
  return false;
}

// Sélectionne le document visé par une étape : actif, ou retrouvé par nom.
function tbGetDocForStep(step, row) {
  const ready = tbState.docs.filter((d) => !d.error || tbDocFullText(d).trim()).filter((d) => !d.loading);
  if (!ready.length) return { error: "aucun PDF chargé dans la Toolbox" };
  if (step.pdfDocMode === "match") {
    const needle = tbNorm(resolveRowTemplate(step.pdfDocMatch || "", row)).trim();
    if (!needle) return { error: "valeur vide pour retrouver le PDF par son nom" };
    const hits = ready.filter((d) => tbNorm(d.name).includes(needle));
    if (!hits.length) return { error: `aucun PDF dont le nom contient « ${needle} »` };
    return { doc: hits[0], warn: hits.length > 1 ? `${hits.length} PDF correspondent — « ${hits[0].name} » utilisé` : "" };
  }
  const active = ready.find((d) => d.id === tbState.activeDocId) || ready[0];
  return { doc: active };
}

/* ================== 6. EXPORT / PROFILS / ONGLETS / INIT ================== */

/* ---------- Export ---------- */

function getOrBuildWorkbook() {
  // Cas 1 : fichier d'origine chargé → on met à jour UNIQUEMENT les valeurs
  // dans la feuille existante, en conservant toute la mise en forme
  // (styles .s, formats de nombre .z, largeurs de colonnes !cols,
  //  hauteurs de lignes !rows, fusions !merges).
  if (state.workbook && state.sheetName && state.workbook.Sheets[state.sheetName]) {
    patchSheetValues(state.workbook.Sheets[state.sheetName], state.rows);
    return state.workbook;
  }
  // Cas 2 : données collées / JSON → aucune mise en forme d'origine à préserver.
  const ws = XLSX.utils.aoa_to_sheet(state.rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, (state.sheetName || "Feuil1").slice(0, 31));
  return wb;
}

// Écrit les valeurs de `rows` (tableau 2D) dans la feuille `ws` sans détruire
// le style existant de chaque cellule. Une cellule inchangée n'est pas touchée
// du tout (valeur, type, format et style d'origine préservés à l'identique).
function patchSheetValues(ws, rows) {
  const nRows = rows.length;
  let nCols = 0;
  for (const r of rows) if (r && r.length > nCols) nCols = r.length;

  for (let R = 0; R < nRows; R++) {
    const row = rows[R] || [];
    for (let C = 0; C < nCols; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      let v = row[C];
      if (v === undefined || v === null) v = "";
      v = String(v);
      const cell = ws[addr];

      if (cell) {
        // Valeur affichée actuelle de la cellule d'origine.
        const cur = cell.w != null ? cell.w : (cell.v != null ? String(cell.v) : "");
        if (v === cur) continue; // inchangée → on ne touche à rien

        const style = cell.s;                 // on garde le style visuel
        const wasNumeric = cell.t === "n";
        const numFmt = cell.z;                 // format de nombre éventuel
        delete cell.w;                          // le texte formaté n'est plus valide

        // Si la cellule était numérique et que la nouvelle valeur est un nombre,
        // on réécrit un nombre en conservant son format ; sinon on écrit du texte.
        const num = wasNumeric ? parseLocaleNumber(v) : null;
        if (num !== null && isFinite(num)) {
          cell.t = "n";
          cell.v = num;
          if (numFmt != null) cell.z = numFmt;
        } else {
          cell.t = "s";
          cell.v = v;
          delete cell.z;
        }
        if (style !== undefined) cell.s = style;
      } else if (v !== "") {
        // Nouvelle cellule (ligne/colonne ajoutée) : pas de style d'origine.
        ws[addr] = { t: "s", v: v };
      }
    }
  }

  // Étendre la plage utilisée si des lignes/colonnes ont été ajoutées.
  if (nRows > 0 && nCols > 0) {
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
    range.s.r = 0; range.s.c = 0;
    if (nRows - 1 > range.e.r) range.e.r = nRows - 1;
    if (nCols - 1 > range.e.c) range.e.c = nCols - 1;
    ws["!ref"] = XLSX.utils.encode_range(range);
  }
}

// Convertit une valeur texte éventuellement localisée ("1 234,56") en nombre,
// ou renvoie null si ce n'est pas un nombre.
function parseLocaleNumber(v) {
  const s = String(v).trim();
  if (s === "") return null;
  // Retire les espaces (normaux et insécables) ; virgule décimale → point.
  const cleaned = s.replace(/[   \t]/g, "").replace(",", ".");
  if (!/^-?\d*\.?\d+(e-?\d+)?$/i.test(cleaned)) return null;
  const n = Number(cleaned);
  return isNaN(n) ? null : n;
}

function rowsToJson(allRows) {
  if (!allRows.length) return [];
  const cols = getColumns();
  const start = dataStartIdx();
  return allRows.slice(start).map((r) => {
    const obj = {};
    cols.forEach((c) => { obj[c.name] = r[c.index] !== undefined ? r[c.index] : ""; });
    return obj;
  });
}

// Nom du fichier de sortie : modifiable par l'utilisateur dans l'onglet Export.
function sanitizeFileName(name) {
  return (name || "").trim().replace(/[\\/:*?"<>|]/g, "_");
}

function currentOutputName() {
  let name = sanitizeFileName($("outputNameInput").value) || state.originalFileName || "resultat.xlsx";
  if (!/\.xlsx$/i.test(name)) name += ".xlsx";
  return name;
}

$("outputNameInput").addEventListener("change", () => {
  state.originalFileName = currentOutputName();
  $("outputNameInput").value = state.originalFileName;
  persistSession();
});

$("downloadBtn").addEventListener("click", async () => {
  if (!state.rows.length) { showStatus("Aucune donnée chargée.", "error"); return; }
  try {
    // Si les octets d'origine ne sont plus en mémoire (réouverture du panneau),
    // tenter de les recharger depuis IndexedDB.
    if (!state.originalArrayBuffer && state.hasOriginalFile) {
      try { state.originalArrayBuffer = await idbLoadOriginal(); } catch (e) { /* ignoré */ }
    }
    // Fichier .xlsx d'origine disponible → export fidèle via ExcelJS (conserve
    // couleurs, polices, tailles, bordures, formats, largeurs, fusions).
    if (state.originalArrayBuffer && typeof ExcelJS !== "undefined") {
      await exportWithFormatting();
    } else {
      // Prévenir clairement si on s'apprête à perdre la mise en forme d'un fichier.
      if (state.hasOriginalFile) {
        showStatus("Impossible de retrouver le fichier d'origine : recharge-le dans l'onglet Données pour conserver la mise en forme. Export sans styles pour l'instant.", "info");
      }
      XLSX.writeFile(getOrBuildWorkbook(), currentOutputName());
    }
    hasDownloaded = true;
    updateDoneMarkers();
  } catch (err) {
    showStatus("Erreur lors de l'export : " + err.message, "error");
  }
});

// Recharge le fichier d'origine avec ExcelJS, met à jour UNIQUEMENT les valeurs
// modifiées (le style de chaque cellule est conservé automatiquement), puis
// télécharge le résultat.
async function exportWithFormatting() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(state.originalArrayBuffer.slice(0));
  let ws = state.sheetName ? wb.getWorksheet(state.sheetName) : null;
  if (!ws) ws = wb.worksheets[0];
  if (!ws) throw new Error("Feuille introuvable dans le fichier.");

  const rows = state.rows;
  let nCols = 0;
  for (const r of rows) if (r && r.length > nCols) nCols = r.length;

  for (let R = 0; R < rows.length; R++) {
    const row = rows[R] || [];
    for (let C = 0; C < nCols; C++) {
      const cell = ws.getCell(R + 1, C + 1);
      let v = row[C];
      if (v === undefined || v === null) v = "";
      v = String(v);

      // Valeur affichée actuelle de la cellule d'origine.
      const curText = cell.text != null ? String(cell.text) : "";
      if (v === curText) continue; // inchangée → on ne touche à rien (style + valeur préservés)

      // La valeur a changé : ExcelJS conserve le style (cell.style) quand on
      // réaffecte cell.value. On garde le type numérique si c'en était un.
      const wasNumber = typeof cell.value === "number";
      const num = wasNumber ? parseLocaleNumber(v) : null;
      cell.value = (num !== null && isFinite(num)) ? num : v;
    }
  }

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = currentOutputName();
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  showStatus("Fichier Excel exporté (mise en forme conservée).", "success");
}

$("downloadJsonBtn").addEventListener("click", () => {
  if (!state.rows.length) { showStatus("Aucune donnée chargée.", "error"); return; }
  const data = rowsToJson(state.rows);
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = currentOutputName().replace(/\.xlsx$/i, "") + ".json";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  hasDownloaded = true;
  updateDoneMarkers();
  showStatus("Export JSON téléchargé.", "success");
});

/* ---------- Profils (configurations nommées) ---------- */

function buildProfileConfig() {
  return {
    headerMode: state.headerMode,
    modelName: state.modelName,
    mapping: state.mapping,
    customFields: getCustomFieldsFromDOM().filter((f) => f.name || f.value),
    valueRules: state.valueRules,
    extraction: {
      conditions: getConditions(),
      searchFields: Array.from($("searchFieldsList").querySelectorAll(".search-field-item")).map((el) => ({
        selector: el.querySelector(".selector-input").value.trim(),
        col: el.querySelector(".col-select").value
      })),
      navEnabled: $("navEnabled").checked,
      navUrlTemplate: $("navUrlTemplate").value,
      navWaitLoad: $("navWaitLoad").checked,
      navWaitTimeout: $("navWaitTimeout").value,
      navExtraWaitMs: $("navExtraWaitMs").value,
      submitMode: document.querySelector('input[name="submitMode"]:checked').value,
      submitSelector: $("submitSelector").value,
      waitMs: $("waitMs").value,
      rowDelayMs: $("rowDelayMs").value,
      startRow: $("startRow").value,
      endRow: $("endRow").value,
      outputs: getOutputs()
    },
    scenario: {
      steps: getScenarioSteps(),
      startRow: $("scnStartRow").value,
      endRow: $("scnEndRow").value,
      rowDelayMs: $("scnRowDelayMs").value
    }
  };
}

function applyProfileConfig(cfg) {
  if (!cfg) return;
  state.headerMode = cfg.headerMode || "auto";
  $("headerModeSelect").value = state.headerMode;
  state.modelName = cfg.modelName || "";
  renderModelSelect();

  if (cfg.mapping && Object.keys(cfg.mapping).length) {
    state.mapping = cfg.mapping;
    if (state.rows.length) {
      allMappings[mappingKey()] = cfg.mapping;
      chrome.storage.local.set({ allMappings });
    }
  }

  state.customFields = Array.isArray(cfg.customFields) ? cfg.customFields : [];
  renderCustomFields();

  if (Array.isArray(cfg.valueRules)) {
    state.valueRules = cfg.valueRules;
    persistValueRules();
    updateValueRulesInfo();
  }

  const ex = cfg.extraction || {};
  $("conditionsList").innerHTML = "";
  (ex.conditions || []).forEach((c) => addConditionRow(c.col, c.op, c.val));
  $("searchFieldsList").innerHTML = "";
  (ex.searchFields || []).forEach((f) => addSearchFieldRow(f.selector, f.col));
  $("outputsList").innerHTML = "";
  (ex.outputs || []).forEach((o) => addOutputRow(o.newCol ? { ...o, col: "__other__", newCol: o.newCol || o.col } : o));
  $("navEnabled").checked = !!ex.navEnabled;
  $("navUrlTemplate").value = ex.navUrlTemplate || "";
  $("navWaitLoad").checked = ex.navWaitLoad !== false;
  $("navWaitTimeout").value = ex.navWaitTimeout || 15000;
  $("navExtraWaitMs").value = ex.navExtraWaitMs || 0;
  $("navOptions").style.display = ex.navEnabled ? "block" : "none";
  const submitMode = ex.submitMode || "enter";
  document.querySelector(`input[name="submitMode"][value="${submitMode}"]`).checked = true;
  $("submitSelectorRow").style.display = submitMode === "click" ? "flex" : "none";
  $("submitSelector").value = ex.submitSelector || "";
  $("waitMs").value = ex.waitMs || 1200;
  $("rowDelayMs").value = ex.rowDelayMs || 300;
  $("startRow").value = ex.startRow || (dataStartIdx() + 1);
  $("endRow").value = ex.endRow || "";

  const sc = cfg.scenario || {};
  $("scenarioSteps").innerHTML = "";
  (sc.steps || []).forEach((s) => addScenarioStep(s));
  $("scnStartRow").value = sc.startRow || 2;
  $("scnEndRow").value = sc.endRow || "";
  $("scnRowDelayMs").value = sc.rowDelayMs || 500;

  renderColumns();
  updateMappingInfo();
  updateFillButtonState();
}

function persistProfiles() { chrome.storage.local.set({ profiles }); }

/* ---------- Sauvegarde automatique des options de travail ----------
   Mémorise en continu tous les réglages (Saisie + Extraction) pour les
   restaurer à la réouverture du panneau, sans avoir à enregistrer un profil. */
let workingConfigTimer = null;
function persistWorkingConfig() {
  if (!workingReady) return;            // on n'écrit pas pendant l'initialisation
  clearTimeout(workingConfigTimer);
  workingConfigTimer = setTimeout(() => {
    try {
      chrome.storage.local.set({ workingConfig: buildProfileConfig() });
    } catch (e) { /* ignoré */ }
  }, 400);
}

// Toute modification d'un champ de réglage déclenche la sauvegarde auto.
["input", "change"].forEach((evt) => {
  document.addEventListener(evt, (e) => {
    const t = e.target;
    if (!t) return;
    if (t.type === "file" || t.id === "profileSelect") return; // gérés à part
    persistWorkingConfig();
  });
});

function renderProfileSelect(selected) {
  const sel = $("profileSelect");
  sel.innerHTML = '<option value="">— aucun profil —</option>';
  Object.keys(profiles).sort().forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    sel.appendChild(opt);
  });
  sel.value = selected || "";
}

$("profileSelect").addEventListener("change", () => {
  const name = $("profileSelect").value;
  chrome.storage.local.set({ activeProfileName: name });
  if (name && profiles[name]) {
    applyProfileConfig(profiles[name]);
    showStatus(`Profil « ${name} » chargé.`, "success");
  }
});

$("saveProfileBtn").addEventListener("click", () => {
  const name = $("profileSelect").value;
  if (!name) { showStatus("Choisis un profil, ou crée-en un avec +.", "error"); return; }
  profiles[name] = buildProfileConfig();
  persistProfiles();
  showStatus(`Profil « ${name} » sauvegardé.`, "success");
});

$("newProfileBtn").addEventListener("click", () => {
  $("newProfileRow").hidden = false;
  $("newProfileName").focus();
});
$("cancelNewProfileBtn").addEventListener("click", () => {
  $("newProfileRow").hidden = true;
  $("newProfileName").value = "";
});
$("confirmNewProfileBtn").addEventListener("click", () => {
  const name = $("newProfileName").value.trim();
  if (!name) { showStatus("Donne un nom au profil.", "error"); return; }
  profiles[name] = buildProfileConfig();
  persistProfiles();
  renderProfileSelect(name);
  chrome.storage.local.set({ activeProfileName: name });
  $("newProfileRow").hidden = true;
  $("newProfileName").value = "";
  showStatus(`Profil « ${name} » créé.`, "success");
});
$("newProfileName").addEventListener("keydown", (e) => {
  if (e.key === "Enter") $("confirmNewProfileBtn").click();
});

$("deleteProfileBtn").addEventListener("click", () => {
  const name = $("profileSelect").value;
  if (!name) { showStatus("Aucun profil sélectionné.", "error"); return; }
  delete profiles[name];
  persistProfiles();
  renderProfileSelect("");
  chrome.storage.local.set({ activeProfileName: "" });
  showStatus(`Profil « ${name} » supprimé.`, "success");
});

/* ---------- Onglets principaux ---------- */

function showTab(tabName) {
  document.querySelectorAll(".panel[data-tab]").forEach((p) => {
    p.hidden = p.getAttribute("data-tab") !== tabName;
  });
  document.querySelectorAll(".tab-btn[data-tab]").forEach((b) => {
    b.classList.toggle("active", b.getAttribute("data-tab") === tabName);
  });
  updateDoneMarkers();
  if (tabName === "extraction") updateExtractEstimate();
  if (tabName === "saisie") updateScenarioEstimate();
  if (workingReady) chrome.storage.local.set({ lastTab: tabName });
}

document.querySelectorAll(".tab-btn[data-tab]").forEach((btn) => {
  btn.addEventListener("click", () => showTab(btn.getAttribute("data-tab")));
});

// Recalcule l'estimation d'extraction dès qu'un paramètre de temps / de plage change.
["waitMs", "rowDelayMs", "navExtraWaitMs", "navEnabled", "startRow", "endRow"].forEach((id) => {
  const el = $(id);
  if (el) el.addEventListener("input", updateExtractEstimate);
});
// Idem pour le scénario de saisie (plage de la boucle et pause).
["scnStartRow", "scnEndRow", "scnRowDelayMs"].forEach((id) => {
  const el = $(id);
  if (el) el.addEventListener("input", updateScenarioEstimate);
});

function updateDoneMarkers() {
  const done = {
    donnees: state.rows.length > 0,
    saisie: state.selectedRowIdx !== null && Object.keys(state.mapping).length > 0,
    extraction: hasStartedRun && !isRunning,
    export: hasDownloaded,
    toolbox: tbState.docs.some((d) => !d.loading && !d.error)
  };
  document.querySelectorAll(".tab-btn[data-tab]").forEach((btn) => {
    btn.classList.toggle("done", Boolean(done[btn.getAttribute("data-tab")]));
  });
}

/* ---------- Migration de l'ancienne config OSA ---------- */

function migrateLegacyOsaConfig(savedConfig) {
  // Convertit l'ancienne config OSA (colonnes en lettres) en profil.
  const legacyModelCols = [];
  (savedConfig.mappings || []).forEach((m) => {
    const idx = letterToIndex(m.col);
    if (idx >= 0) {
      while (legacyModelCols.length <= idx) legacyModelCols.push("");
      legacyModelCols[idx] = m.label;
    }
  });
  const searchFields = savedConfig.searchFields
    || (savedConfig.searchSelector ? [{ selector: savedConfig.searchSelector, col: savedConfig.searchCol }] : []);
  return {
    headerMode: "auto",
    modelName: "",
    mapping: {},
    customFields: [],
    extraction: {
      conditions: (savedConfig.conditions || []).map((c) => ({ col: (c.col || "").toUpperCase(), op: c.op, val: c.val })),
      searchFields: searchFields.map((f) => ({ selector: f.selector, col: (f.col || "").toUpperCase() })),
      submitMode: savedConfig.submitMode || "enter",
      submitSelector: savedConfig.submitSelector || "",
      waitMs: savedConfig.waitMs || 1200,
      rowDelayMs: savedConfig.rowDelayMs || 300,
      startRow: savedConfig.startRow || 2,
      endRow: savedConfig.endRow || "",
      outputs: (savedConfig.outputs || []).map((o) => ({
        mode: o.mode || "css",
        col: (o.col || "").toUpperCase(),
        selector: o.selector || "",
        rowSelector: o.rowSelector || "",
        matchSourceCol: (o.matchSourceCol || "").toUpperCase(),
        matchType: o.matchType || "contains",
        matchTdIndex: o.matchTdIndex || 1,
        extractTdIndex: o.extractTdIndex || 2
      }))
    }
  };
}

/* ---------- Initialisation ---------- */

async function init() {
  const stored = await chrome.storage.local.get([
    "models", "allMappings", "profiles", "activeProfileName",
    "customFields", "session", "savedConfig", "migratedOsaLegacy",
    "workingConfig", "lastTab", "valueRules", "tbPatterns"
  ]);

  // Règles de valeurs (seed avec la civilité au premier lancement).
  if (Array.isArray(stored.valueRules)) {
    state.valueRules = stored.valueRules;
  } else {
    state.valueRules = JSON.parse(JSON.stringify(DEFAULT_VALUE_RULES));
    persistValueRules();
  }

  // Modèles (seed au premier lancement)
  if (Array.isArray(stored.models) && stored.models.length) {
    models = stored.models;
  } else {
    models = JSON.parse(JSON.stringify(DEFAULT_MODELS));
    persistModels();
  }

  allMappings = stored.allMappings || {};
  profiles = stored.profiles || {};

  // Migration de l'ancienne config OSA -> profil "OSA (importé)"
  if (stored.savedConfig && !stored.migratedOsaLegacy) {
    profiles["OSA (importé)"] = migrateLegacyOsaConfig(stored.savedConfig);
    persistProfiles();
    chrome.storage.local.set({ migratedOsaLegacy: true });
  }

  // Champs personnalisés
  if (Array.isArray(stored.customFields)) state.customFields = stored.customFields;
  renderCustomFields();

  // Motifs personnalisés de la Toolbox PDF
  tbState.patterns = Array.isArray(stored.tbPatterns) ? stored.tbPatterns : [];
  tbRenderPatterns();

  // Session précédente (données + réglages)
  if (stored.session && Array.isArray(stored.session.rows) && stored.session.rows.length) {
    state.rows = stored.session.rows;
    state.sheetName = stored.session.sheetName || null;
    state.headerMode = stored.session.headerMode || "auto";
    state.modelName = stored.session.modelName || "";
    state.selectedRowIdx = stored.session.selectedRowIdx ?? null;
    state.originalFileName = stored.session.originalFileName || "resultat.xlsx";
    state.hasOriginalFile = !!stored.session.hasOriginalFile;
    if (state.hasOriginalFile) {
      try { state.originalArrayBuffer = await idbLoadOriginal(); } catch (e) { state.originalArrayBuffer = null; }
    }
    $("headerModeSelect").value = state.headerMode;
    $("outputNameInput").value = state.originalFileName;
    setLoadedInfo(`Session restaurée (« ${state.sheetName || "données"} »)`);
  }

  renderModelSelect();

  // Profil actif
  renderProfileSelect(stored.activeProfileName || "");
  if (stored.activeProfileName && profiles[stored.activeProfileName]) {
    applyProfileConfig(profiles[stored.activeProfileName]);
  } else {
    renderColumns();
    updateMappingInfo();
  }

  // Dernier état de travail auto-sauvegardé : reflète les options exactes
  // laissées à la fermeture précédente (prioritaire sur les valeurs du profil).
  if (stored.workingConfig) {
    applyProfileConfig(stored.workingConfig);
  }

  // Lignes par défaut dans l'onglet Extraction si vide
  if (!$("searchFieldsList").children.length) addSearchFieldRow();
  if (!$("outputsList").children.length) addOutputRow();

  // Onglet actif restauré
  if (stored.lastTab) showTab(stored.lastTab);

  updateFillButtonState();
  updateDoneMarkers();
  updateValueRulesInfo();

  // À partir d'ici, toute modification est sauvegardée automatiquement.
  workingReady = true;
}

init();

// (fin du fichier)