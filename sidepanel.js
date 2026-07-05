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
  targetTabId: null        // onglet du formulaire mémorisé (dernier 🎯) — indép. de l'onglet affiché
};

let models = [];           // [{name, columns: []}]
let allMappings = {};      // clé de colonnes -> mapping
let profiles = {};         // nom -> config
let stopRequested = false;
let isRunning = false;
let hasStartedRun = false;
let hasDownloaded = false;
let runLog = [];
let lastRunOutputs = [];
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
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
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
      originalFileName: state.originalFileName
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
    } else {
      const buf = await file.arrayBuffer();
      state.workbook = XLSX.read(buf, { type: "array" });
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

function renderRowList() {
  const list = $("rowList");
  list.innerHTML = "";
  if (!state.rows.length) {
    $("selectedRowLabel").textContent = "—";
    return;
  }
  const filter = $("rowFilterInput").value.trim().toLowerCase();
  const indices = dataRowIndices();
  let shown = 0;
  for (const idx of indices) {
    if (shown >= 300) break;
    const row = state.rows[idx] || [];
    if (filter && !row.join(" ").toLowerCase().includes(filter)) continue;
    shown++;
    const div = document.createElement("div");
    div.className = "row-item" + (idx === state.selectedRowIdx ? " selected" : "");
    div.innerHTML = `<span class="row-num">L${idx + 1}</span><span class="row-summary">${escapeHtml(rowSummary(row))}</span>`;
    div.addEventListener("click", () => selectRow(idx));
    list.appendChild(div);
  }
  if (!shown) {
    const p = document.createElement("div");
    p.className = "row-item";
    p.innerHTML = '<span class="row-summary" style="color:var(--text-dim)">Aucune ligne ne correspond au filtre.</span>';
    list.appendChild(p);
  }
  $("selectedRowLabel").textContent = state.selectedRowIdx !== null ? `L${state.selectedRowIdx + 1}` : "—";
}

function selectRow(idx) {
  state.selectedRowIdx = idx;
  renderRowList();
  renderSelectedRowFields();
  updateFillButtonState();
  updateDoneMarkers();
  persistSession();
}

$("rowFilterInput").addEventListener("input", renderRowList);

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
  input.placeholder = "attribut name (ex: body:x:tabc:x:txtNom)";
  input.value = value;

  const pickBtn = document.createElement("button");
  pickBtn.className = "btn pick icon-only";
  pickBtn.type = "button";
  pickBtn.title = "Cliquer sur le champ du site pour récupérer son name";
  pickBtn.innerHTML = '<svg class="icon"><use href="#icon-target"/></svg>';
  pickBtn.addEventListener("click", async () => {
    const picked = await pickTargetOnActiveTab();
    if (!picked) return;
    if (picked.name) {
      input.value = picked.name;
      showStatus(`✓ name récupéré : ${picked.name}`, "success");
    } else {
      showStatus("Cet élément n'a pas d'attribut name — mapping impossible sur ce champ.", "error");
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
      <input type="text" class="custom-field-name" placeholder="nom de l'input" value="${escapeAttr(item.name)}" />
      <input type="text" class="custom-field-value" placeholder="valeur" value="${escapeAttr(item.value)}" />
      <button class="remove-btn" type="button" title="Supprimer"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
    `;
    row.querySelectorAll("input").forEach((inp) => inp.addEventListener("blur", saveCustomFields));
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

// Construit data/mapping/customFields pour une ligne donnée.
// Utilisé par le bouton « Remplir le formulaire » et par le scénario de saisie.
function buildFillPayload(rowIdx) {
  const data = {};
  const mapping = {};
  if (rowIdx !== null && rowIdx !== undefined && state.rows.length) {
    const row = state.rows[rowIdx] || [];
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
  state.customFields.forEach(({ name, value }) => { if (name) customFields[name] = value; });
  return { data, mapping, customFields };
}

$("fillBtn").addEventListener("click", async () => {
  saveCustomFields();

  const { data, mapping, customFields } = buildFillPayload(state.selectedRowIdx);

  if (!Object.keys(mapping).length && !Object.keys(customFields).length) {
    showStatus("Aucune donnée exploitable : configure le mapping ou des champs personnalisés.", "error");
    return;
  }

  try {
    const message = {
      action: "fillForm",
      data,
      mapping,
      customFields: Object.keys(customFields).length ? customFields : undefined
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
        name: target.getAttribute ? (target.getAttribute("name") || null) : null
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
    <input type="text" class="val-input" placeholder="valeur" value="${escapeAttr(val)}" />
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
    if (cellMatches(cell, c.op, c.val)) return true; // une condition qui matche => ligne ignorée
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

// o = { mode: "css"|"tableMatch", col, selector, rowSelector, matchSourceCol,
//       matchType, matchTdIndex, extractTdIndex, newCol }
function addOutputRow(o = {}) {
  const mode = o.mode || "css";
  const div = document.createElement("div");
  div.className = "out-item";
  div.innerHTML = `
    <div class="out-item-row1">
      <select class="mode-select">
        <option value="css">Info sur la page (sélecteur CSS)</option>
        <option value="tableMatch">Ligne de tableau (par valeur)</option>
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
    <div class="out-item-notfound">
      <label class="checkbox-row">
        <input type="checkbox" class="notfound-check" ${o.notFoundEnabled ? "checked" : ""} />
        Si rien n'est trouvé, écrire un message
      </label>
      <input type="text" class="notfound-msg" placeholder="ex : non trouvé" value="${escapeAttr(o.notFoundMsg || "")}" ${o.notFoundEnabled ? "" : "style=display:none"} />
    </div>
  `;

  div.querySelector(".mode-select").value = mode;
  div.querySelector(".match-type-select").value = o.matchType || "contains";

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
  const cssHint = div.querySelector(".out-item-csshint");
  const selectorInput = div.querySelector(".selector-input");
  const pickCssBtn = div.querySelector("[data-pick-inline]");
  const modeSelect = div.querySelector(".mode-select");
  const syncMode = () => {
    const isTable = modeSelect.value === "tableMatch";
    tablematchDiv.hidden = !isTable;
    cssHint.style.display = isTable ? "none" : "block";
    selectorInput.style.display = isTable ? "none" : "block";
    pickCssBtn.style.display = isTable ? "none" : "";
  };
  modeSelect.addEventListener("change", syncMode);
  syncMode();

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
    const base = { mode, col, newCol, notFoundEnabled, notFoundMsg };
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
    return { ...base, selector: el.querySelector(".selector-input").value.trim() };
  }).filter((o) => {
    if (!o.col) return false;
    if (o.mode === "tableMatch") return Boolean(o.rowSelector && o.matchSourceCol);
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

      function readResults() {
        const values = [];
        const notFound = [];
        for (const out of config.outputs) {
          const fallback = out.notFoundEnabled ? (out.notFoundMsg || "") : "";
          if (out.mode === "tableMatch") {
            const { found, value, reason } = readTableMatch(out);
            if (found) {
              values.push(value);
            } else {
              values.push(fallback);
              // Si un message perso est défini, on ne signale plus d'erreur "introuvable".
              if (!out.notFoundEnabled) {
                notFound.push(reason === "noRows"
                  ? 'aucune ligne trouvée pour le sélecteur "' + out.rowSelector + '" (vérifie ce sélecteur ou augmente le délai d\'attente)'
                  : 'aucune ligne où la cellule n°' + out.matchTdIndex + ' correspond à "' + out.matchValue + '"');
              }
            }
          } else {
            const el = document.querySelector(out.selector);
            if (!el) {
              values.push(fallback);
              if (!out.notFoundEnabled) notFound.push(out.selector);
              continue;
            }
            values.push(textOf(el));
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

  if (!searchFields.length) { logLine("Ajoute au moins un champ de recherche.", "err"); return; }
  if (!outputs.length) { logLine("Ajoute au moins un résultat à récupérer.", "err"); return; }

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
  setProgress(0, total);
  updateDoneMarkers();

  for (let idx = startIdx; idx <= endIdx; idx++) {
    if (stopRequested) { logLine("Arrêté par l'utilisateur.", "skip"); break; }
    const row = state.rows[idx] || [];
    const rowNum = idx + 1;

    if (rowMatchesSkipCondition(row, conditions)) {
      logLine(`Ligne ${rowNum} : ignorée (condition).`, "skip");
      runLog.push({ row: rowNum, search: "", values: [], status: "skip", note: "Condition" });
      done++; setProgress(done, total);
      continue;
    }

    const searchFieldValues = searchResolved.map((f) => ({ selector: f.selector, value: getCellByIndex(row, f.colIdx) }));
    const searchLabel = searchFieldValues.map((f) => f.value).filter((v) => v.trim()).join(" / ");
    if (!searchFieldValues.some((f) => f.value.trim())) {
      logLine(`Ligne ${rowNum} : ignorée (valeur(s) de recherche vide(s)).`, "skip");
      runLog.push({ row: rowNum, search: "", values: [], status: "skip", note: "Valeur vide" });
      done++; setProgress(done, total);
      continue;
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
            notFoundMsg: o.notFoundMsg
          } : {
            mode: "css",
            selector: o.selector,
            notFoundEnabled: o.notFoundEnabled,
            notFoundMsg: o.notFoundMsg
          })
        }]
      });

      if (!result || !result.ok) {
        const msg = result ? result.error : "pas de réponse";
        logLine(`Ligne ${rowNum} : erreur — ${msg}`, "err");
        runLog.push({ row: rowNum, search: searchLabel, values: [], status: "err", note: msg });
      } else {
        outputsResolved.forEach((o, i) => setCellByIndex(row, o.targetIdx, result.values[i]));
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

    done++; setProgress(done, total);
    if (rowDelayMs > 0 && idx < endIdx && !stopRequested) await sleep(rowDelayMs);
  }

  isRunning = false;
  $("startBtn").disabled = false;
  $("stopBtn").disabled = true;
  logLine("Terminé.", "ok");
  renderResultSummary();
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
  let html = '<div class="summary-table-wrap"><table class="summary-table"><thead><tr><th>Ligne</th><th>Recherche</th>';
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
        <option value="click">Cliquer sur un élément</option>
        <option value="wait">Attendre</option>
        <option value="cond">Condition (si… alors…)</option>
      </select>
      <button class="btn icon-only scn-move-up" title="Monter" type="button"><svg class="icon icon-sm"><use href="#icon-arrow-up"/></svg></button>
      <button class="btn icon-only scn-move-down" title="Descendre" type="button"><svg class="icon icon-sm"><use href="#icon-arrow-down"/></svg></button>
      <button class="remove-btn" title="Supprimer l'étape" type="but