// sidepanel.js
// Toute la logique du panneau : chargement Excel, mapping colonnes, conditions,
// sélection d'éléments sur la page cible, boucle d'automatisation, export.

/* ---------- État global ---------- */
let workbook = null;
let sheetName = null;
let rows = [];          // tableau 2D brut de la feuille active (0-indexé)
let originalFileName = "resultat.xlsx";
let stopRequested = false;
let isRunning = false;

/* ---------- Utilitaires colonnes ---------- */
function colLetterToIndex(letter) {
  letter = (letter || "").trim().toUpperCase();
  if (!letter) return -1;
  let n = 0;
  for (let i = 0; i < letter.length; i++) {
    const c = letter.charCodeAt(i) - 64;
    if (c < 1 || c > 26) return -1;
    n = n * 26 + c;
  }
  return n - 1;
}

function getCell(row, colLetter) {
  const idx = colLetterToIndex(colLetter);
  if (idx < 0 || !row) return "";
  const v = row[idx];
  return v === undefined || v === null ? "" : String(v);
}

function setCell(row, colLetter, value) {
  const idx = colLetterToIndex(colLetter);
  if (idx < 0) return;
  while (row.length <= idx) row.push("");
  row[idx] = value;
}

/* ---------- Chargement du fichier ---------- */
const fileInput = document.getElementById("fileInput");
const sheetRow = document.getElementById("sheetRow");
const sheetSelect = document.getElementById("sheetSelect");
const fileInfo = document.getElementById("fileInfo");

fileInput.addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  originalFileName = file.name.replace(/\.(xlsx|xls|csv)$/i, "") + "_maj.xlsx";

  try {
    if (/\.csv$/i.test(file.name)) {
      const text = await file.text();
      workbook = XLSX.read(text, { type: "string" });
    } else {
      const buf = await file.arrayBuffer();
      workbook = XLSX.read(buf, { type: "array" });
    }
    sheetSelect.innerHTML = "";
    workbook.SheetNames.forEach((name) => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      sheetSelect.appendChild(opt);
    });
    sheetRow.style.display = workbook.SheetNames.length > 1 ? "flex" : "none";
    loadSheet(workbook.SheetNames[0]);
  } catch (err) {
    fileInfo.textContent = "Erreur de lecture du fichier : " + err.message;
  }
});

sheetSelect.addEventListener("change", () => loadSheet(sheetSelect.value));

function setLoadedInfo(label, count) {
  fileInfo.textContent = `${label} : ${count} lignes (en-tête incluse).`;
  document.getElementById("endRow").placeholder = "auto (" + count + ")";
}

function loadSheet(name) {
  sheetName = name;
  const ws = workbook.Sheets[name];
  rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
  setLoadedInfo(`Feuille "${name}" chargée`, rows.length);
}

/* ---------- Sources alternatives : coller un tableau / JSON ---------- */
const sourceFileDiv = document.getElementById("sourceFile");
const sourcePasteDiv = document.getElementById("sourcePaste");
const sourceJsonDiv = document.getElementById("sourceJson");

document.querySelectorAll('input[name="sourceMode"]').forEach((r) => {
  r.addEventListener("change", () => {
    const mode = document.querySelector('input[name="sourceMode"]:checked').value;
    sourceFileDiv.style.display = mode === "file" ? "block" : "none";
    sourcePasteDiv.style.display = mode === "paste" ? "block" : "none";
    sourceJsonDiv.style.display = mode === "json" ? "block" : "none";
  });
});

function parseDelimitedText(text) {
  const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").filter((l) => l.length > 0);
  const delim = text.includes("\t") ? "\t" : ",";
  return lines.map((line) => line.split(delim));
}

document.getElementById("usePasteBtn").addEventListener("click", () => {
  const text = document.getElementById("pasteArea").value;
  if (!text.trim()) { fileInfo.textContent = "Colle d'abord un tableau."; return; }
  workbook = null;
  sheetName = "Tableau collé";
  originalFileName = "resultat.xlsx";
  rows = parseDelimitedText(text);
  sheetRow.style.display = "none";
  setLoadedInfo("Tableau collé chargé", rows.length);
});

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
  if (!text.trim()) { fileInfo.textContent = "Colle ou choisis d'abord un JSON."; return; }
  try {
    const data = JSON.parse(text);
    rows = jsonToRows(data);
    workbook = null;
    sheetName = "JSON";
    originalFileName = "resultat.xlsx";
    sheetRow.style.display = "none";
    setLoadedInfo("JSON chargé", rows.length);
  } catch (err) {
    fileInfo.textContent = "Erreur JSON : " + err.message;
  }
}

document.getElementById("useJsonBtn").addEventListener("click", () => {
  applyJsonText(document.getElementById("jsonTextArea").value);
});

document.getElementById("jsonFileInput").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const text = await file.text();
  document.getElementById("jsonTextArea").value = text;
  applyJsonText(text);
});

/* ---------- Mapping des colonnes ---------- */
const mappingList = document.getElementById("mappingList");

function addMappingRow(label = "", col = "") {
  const div = document.createElement("div");
  div.className = "map-item";
  div.innerHTML = `
    <input type="text" class="label-input" placeholder="nom du champ (ex: nom)" value="${escapeAttr(label)}" />
    <input type="text" class="col-input" placeholder="col" value="${escapeAttr(col)}" />
    <button class="remove-btn" title="Supprimer" type="button"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
  `;
  div.querySelector(".remove-btn").addEventListener("click", () => {
    div.remove();
    refreshAllColumnSelects();
  });
  mappingList.appendChild(div);
  refreshAllColumnSelects();
}

document.getElementById("addMappingBtn").addEventListener("click", () => addMappingRow());
mappingList.addEventListener("input", refreshAllColumnSelects);

function getMappings() {
  return Array.from(mappingList.querySelectorAll(".map-item")).map((el) => ({
    label: el.querySelector(".label-input").value.trim(),
    col: el.querySelector(".col-input").value.trim(),
  })).filter((m) => m.label && m.col);
}

/* ---------- Champs de recherche (un ou plusieurs) ---------- */
const searchFieldsList = document.getElementById("searchFieldsList");

function refreshAllColumnSelects() {
  const mappings = getMappings();
  document.querySelectorAll("#searchFieldsList .col-select, #outputsList .match-col-select").forEach((sel) => {
    const current = sel.value;
    sel.innerHTML = "";
    mappings.forEach((m) => {
      const opt = document.createElement("option");
      opt.value = m.col;
      opt.textContent = `${m.label} (col. ${m.col.toUpperCase()})`;
      sel.appendChild(opt);
    });
    if (mappings.some((m) => m.col === current)) sel.value = current;
  });
}

function addSearchFieldRow(selector = "", col = "") {
  const div = document.createElement("div");
  div.className = "search-field-item";
  div.innerHTML = `
    <input type="text" class="selector-input" placeholder="sélecteur CSS du champ" value="${escapeAttr(selector)}" />
    <button class="btn pick icon-only" data-pick-inline="1" title="Choisir sur la page" type="button"><svg class="icon"><use href="#icon-target"/></svg></button>
    <select class="col-select"></select>
    <button class="remove-btn" title="Supprimer" type="button"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
  `;
  div.querySelector(".remove-btn").addEventListener("click", () => div.remove());
  div.querySelector("[data-pick-inline]").addEventListener("click", async () => {
    const selInput = div.querySelector(".selector-input");
    const picked = await pickSelectorOnActiveTab();
    if (picked) selInput.value = picked;
  });
  searchFieldsList.appendChild(div);
  refreshAllColumnSelects();
  if (col) div.querySelector(".col-select").value = col;
}

document.getElementById("addSearchFieldBtn").addEventListener("click", () => addSearchFieldRow());

function getSearchFields() {
  return Array.from(searchFieldsList.querySelectorAll(".search-field-item")).map((el) => ({
    selector: el.querySelector(".selector-input").value.trim(),
    col: el.querySelector(".col-select").value,
  })).filter((f) => f.selector && f.col);
}

/* ---------- Conditions ---------- */
const conditionsList = document.getElementById("conditionsList");
const OPERATORS = [
  { v: "equals", t: "= égal à" },
  { v: "not_equals", t: "≠ différent de" },
  { v: "contains", t: "contient" },
  { v: "not_contains", t: "ne contient pas" },
  { v: "empty", t: "est vide" },
  { v: "not_empty", t: "n'est pas vide" },
];

function addConditionRow(col = "", op = "equals", val = "") {
  const div = document.createElement("div");
  div.className = "cond-item";
  const opOptions = OPERATORS.map((o) => `<option value="${o.v}" ${o.v === op ? "selected" : ""}>${o.t}</option>`).join("");
  div.innerHTML = `
    <input type="text" class="col-input" placeholder="col" value="${escapeAttr(col)}" />
    <select class="op-select">${opOptions}</select>
    <input type="text" class="val-input" placeholder="valeur" value="${escapeAttr(val)}" />
    <button class="remove-btn" title="Supprimer" type="button"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
  `;
  const valInput = div.querySelector(".val-input");
  const opSelect = div.querySelector(".op-select");
  const toggleVal = () => {
    const needsVal = !["empty", "not_empty"].includes(opSelect.value);
    valInput.style.display = needsVal ? "block" : "none";
  };
  opSelect.addEventListener("change", toggleVal);
  toggleVal();
  div.querySelector(".remove-btn").addEventListener("click", () => div.remove());
  conditionsList.appendChild(div);
}

document.getElementById("addConditionBtn").addEventListener("click", () => addConditionRow());

function getConditions() {
  return Array.from(conditionsList.querySelectorAll(".cond-item")).map((el) => ({
    col: el.querySelector(".col-input").value.trim(),
    op: el.querySelector(".op-select").value,
    val: el.querySelector(".val-input").value,
  })).filter((c) => c.col);
}

function rowMatchesSkipCondition(row, conditions) {
  for (const c of conditions) {
    const cell = getCell(row, c.col);
    let match = false;
    switch (c.op) {
      case "equals": match = cell.trim().toLowerCase() === c.val.trim().toLowerCase(); break;
      case "not_equals": match = cell.trim().toLowerCase() !== c.val.trim().toLowerCase(); break;
      case "contains": match = cell.toLowerCase().includes(c.val.toLowerCase()); break;
      case "not_contains": match = !cell.toLowerCase().includes(c.val.toLowerCase()); break;
      case "empty": match = cell.trim() === ""; break;
      case "not_empty": match = cell.trim() !== ""; break;
    }
    if (match) return true; // une condition qui matche => on ignore la ligne
  }
  return false;
}

/* ---------- Résultats à récupérer (outputs) ---------- */
const outputsList = document.getElementById("outputsList");

// o = { mode: "css"|"tableMatch", selector, col, rowSelector, matchSourceCol, matchType, matchTdIndex, extractTdIndex }
function addOutputRow(o = {}) {
  const mode = o.mode || "css";
  const div = document.createElement("div");
  div.className = "out-item";
  div.innerHTML = `
    <div class="out-item-row1">
      <select class="mode-select">
        <option value="css">Sélecteur CSS</option>
        <option value="tableMatch">Ligne de tableau (par valeur)</option>
      </select>
      <input type="text" class="selector-input" placeholder="sélecteur CSS du résultat" value="${escapeAttr(o.selector || "")}" />
      <button class="btn pick icon-only" data-pick-inline="1" title="Choisir sur la page" type="button"><svg class="icon"><use href="#icon-target"/></svg></button>
      <input type="text" class="col-input" placeholder="col." value="${escapeAttr(o.col || "")}" />
      <button class="remove-btn" title="Supprimer" type="button"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
    </div>
    <div class="out-item-tablematch" ${mode === "tableMatch" ? "" : "hidden"}>
      <p class="hint">Compare la valeur choisie à la colonne n° indiquée de chaque ligne trouvée par le sélecteur, puis récupère la colonne n° indiquée sur cette même ligne.</p>
      <input type="text" class="row-selector-input" placeholder="sélecteur des lignes (ex: #dataTable tbody tr)" value="${escapeAttr(o.rowSelector || "")}" />
      <select class="match-col-select"></select>
      <select class="match-type-select">
        <option value="contains">contient</option>
        <option value="exact">= exact</option>
      </select>
      <input type="number" class="match-td-input" placeholder="td n° à comparer" min="1" value="${escapeAttr(o.matchTdIndex || 1)}" />
      <input type="number" class="extract-td-input" placeholder="td n° à extraire" min="1" value="${escapeAttr(o.extractTdIndex || 2)}" />
    </div>
  `;
  div.querySelector(".mode-select").value = mode;
  div.querySelector(".match-type-select").value = o.matchType || "contains";
  const tablematchDiv = div.querySelector(".out-item-tablematch");
  div.querySelector(".mode-select").addEventListener("change", (e) => {
    tablematchDiv.hidden = e.target.value !== "tableMatch";
  });
  div.querySelector(".remove-btn").addEventListener("click", () => div.remove());
  div.querySelector("[data-pick-inline]").addEventListener("click", async () => {
    const selInput = div.querySelector(".selector-input");
    const picked = await pickSelectorOnActiveTab();
    if (picked) selInput.value = picked;
  });
  outputsList.appendChild(div);
  refreshAllColumnSelects();
  if (o.matchSourceCol) div.querySelector(".match-col-select").value = o.matchSourceCol;
}

document.getElementById("addOutputBtn").addEventListener("click", () => addOutputRow());

function getOutputs() {
  return Array.from(outputsList.querySelectorAll(".out-item")).map((el) => {
    const mode = el.querySelector(".mode-select").value;
    const col = el.querySelector(".col-input").value.trim();
    if (mode === "tableMatch") {
      return {
        mode,
        col,
        rowSelector: el.querySelector(".row-selector-input").value.trim(),
        matchSourceCol: el.querySelector(".match-col-select").value,
        matchType: el.querySelector(".match-type-select").value,
        matchTdIndex: parseInt(el.querySelector(".match-td-input").value, 10) || 1,
        extractTdIndex: parseInt(el.querySelector(".extract-td-input").value, 10) || 1,
      };
    }
    return {
      mode: "css",
      col,
      selector: el.querySelector(".selector-input").value.trim(),
    };
  }).filter((o) => {
    if (!o.col) return false;
    if (o.mode === "tableMatch") return Boolean(o.rowSelector && o.matchSourceCol);
    return Boolean(o.selector);
  });
}

/* ---------- Sélecteur de soumission (radio) ---------- */
const submitSelectorRow = document.getElementById("submitSelectorRow");
document.querySelectorAll('input[name="submitMode"]').forEach((r) => {
  r.addEventListener("change", () => {
    submitSelectorRow.style.display = document.querySelector('input[name="submitMode"]:checked').value === "click" ? "flex" : "none";
  });
});

/* ---------- Boutons "Choisir sur la page" (top-level fields) ---------- */
document.querySelectorAll(".btn.pick[data-target]").forEach((btn) => {
  btn.addEventListener("click", async () => {
    const targetId = btn.getAttribute("data-target");
    const picked = await pickSelectorOnActiveTab();
    if (picked) document.getElementById(targetId).value = picked;
  });
});

async function getActiveTabId() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab) throw new Error("Aucun onglet actif trouvé.");
  return tab.id;
}

async function pickSelectorOnActiveTab() {
  try {
    const tabId = await getActiveTabId();
    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId },
      func: pickElementOnPageInjected,
    });
    return result;
  } catch (err) {
    logLine("Erreur sélection : " + err.message, "err");
    return null;
  }
}

// Fonction injectée sur la page cible pour choisir un élément au clic.
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
      highlighted.style.outline = "2px solid #5a4bda";
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
      const sel = computeSelector(e.target);
      cleanup();
      resolve(sel);
    }

    function onKeyDown(e) {
      if (e.key === "Escape") { cleanup(); resolve(null); }
    }

    document.addEventListener("mouseover", onMouseOver, true);
    document.addEventListener("click", onClick, true);
    document.addEventListener("keydown", onKeyDown, true);
  });
}

/* ---------- Fonction injectée : action sur une ligne (recherche + lecture résultats) ---------- */
function performRowActionInjected(config) {
  return new Promise((resolve) => {
    try {
      const fields = config.searchFields.map((f) => ({
        selector: f.selector,
        value: f.value,
        el: document.querySelector(f.selector),
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

      // Cherche, parmi toutes les lignes correspondant à rowSelector, celle dont la
      // cellule n° matchTdIndex correspond à matchValue, puis renvoie le texte de la
      // cellule n° extractTdIndex de cette même ligne. Utile quand les résultats sont
      // affichés dans un tableau sans id/name exploitable (plusieurs <td> identiques).
      function readTableMatch(out) {
        const trs = document.querySelectorAll(out.rowSelector);
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
        return { found: false, value: "" };
      }

      function readResults() {
        const values = [];
        const notFound = [];
        for (const out of config.outputs) {
          if (out.mode === "tableMatch") {
            const { found, value } = readTableMatch(out);
            values.push(value);
            if (!found) notFound.push(out.rowSelector + " (aucune ligne correspondant à \"" + out.matchValue + "\")");
          } else {
            const el = document.querySelector(out.selector);
            if (!el) { values.push(""); notFound.push(out.selector); continue; }
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
const logEl = document.getElementById("log");
const progressFill = document.getElementById("progressFill");
const progressText = document.getElementById("progressText");

function logLine(text, cls = "") {
  const line = document.createElement("div");
  if (cls) line.className = cls;
  line.textContent = text;
  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
}

function setProgress(current, total) {
  progressFill.style.width = total ? Math.round((current / total) * 100) + "%" : "0%";
  progressText.textContent = `${current} / ${total}`;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/* ---------- Exécution principale ---------- */
const startBtn = document.getElementById("startBtn");
const stopBtn = document.getElementById("stopBtn");

startBtn.addEventListener("click", runAutomation);
stopBtn.addEventListener("click", () => { stopRequested = true; });

let runLog = [];
let lastOutputs = [];

async function runAutomation() {
  if (!rows.length) { logLine("Chargez d'abord un fichier Excel.", "err"); return; }

  const mappings = getMappings();
  const conditions = getConditions();
  const outputs = getOutputs();
  const searchFields = getSearchFields();
  const submitMode = document.querySelector('input[name="submitMode"]:checked').value;
  const submitSelector = document.getElementById("submitSelector").value.trim();
  const waitMs = parseInt(document.getElementById("waitMs").value, 10) || 0;
  const rowDelayMs = parseInt(document.getElementById("rowDelayMs").value, 10) || 0;

  if (!searchFields.length) { logLine("Ajoutez au moins un champ de recherche.", "err"); return; }
  if (!outputs.length) { logLine("Ajoutez au moins un résultat à récupérer.", "err"); return; }

  const startRowInput = parseInt(document.getElementById("startRow").value, 10) || 2;
  const endRowInput = document.getElementById("endRow").value.trim();
  const startIdx = startRowInput - 1; // 0-indexé
  const endIdx = endRowInput ? parseInt(endRowInput, 10) - 1 : rows.length - 1;

  let tabId;
  try { tabId = await getActiveTabId(); }
  catch (err) { logLine(err.message, "err"); return; }

  isRunning = true;
  stopRequested = false;
  startBtn.disabled = true;
  stopBtn.disabled = false;
  logEl.innerHTML = "";
  runLog = [];
  lastOutputs = outputs;
  const total = Math.max(0, endIdx - startIdx + 1);
  let done = 0;
  setProgress(0, total);

  for (let idx = startIdx; idx <= endIdx; idx++) {
    if (stopRequested) { logLine("Arrêté par l'utilisateur.", "skip"); break; }
    const row = rows[idx] || [];
    const excelRowNum = idx + 1;

    if (rowMatchesSkipCondition(row, conditions)) {
      logLine(`Ligne ${excelRowNum} : ignorée (condition).`, "skip");
      runLog.push({ row: excelRowNum, search: "", values: [], status: "skip", note: "Condition" });
      done++; setProgress(done, total);
      continue;
    }

    const searchFieldValues = searchFields.map((f) => ({ selector: f.selector, value: getCell(row, f.col) }));
    const searchLabel = searchFieldValues.map((f) => f.value).filter((v) => v.trim()).join(" / ");
    const hasAnyValue = searchFieldValues.some((f) => f.value.trim());
    if (!hasAnyValue) {
      logLine(`Ligne ${excelRowNum} : ignorée (valeur(s) de recherche vide(s)).`, "skip");
      runLog.push({ row: excelRowNum, search: "", values: [], status: "skip", note: "Valeur vide" });
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
          outputs: outputs.map((o) => o.mode === "tableMatch" ? {
            mode: "tableMatch",
            rowSelector: o.rowSelector,
            matchType: o.matchType,
            matchTdIndex: o.matchTdIndex,
            extractTdIndex: o.extractTdIndex,
            matchValue: getCell(row, o.matchSourceCol),
          } : { mode: "css", selector: o.selector }),
        }],
      });

      if (!result || !result.ok) {
        const msg = result ? result.error : "pas de réponse";
        logLine(`Ligne ${excelRowNum} : erreur — ${msg}`, "err");
        runLog.push({ row: excelRowNum, search: searchLabel, values: [], status: "err", note: msg });
      } else {
        outputs.forEach((o, i) => setCell(row, o.col, result.values[i]));
        rows[idx] = row;
        const missingSelectors = result.notFound || [];
        const allValuesEmpty = result.values.every((v) => !String(v || "").trim());
        if (missingSelectors.length) {
          logLine(`Ligne ${excelRowNum} (${searchLabel}) : sélecteur introuvable sur la page (${missingSelectors.join(", ")}).`, "err");
          runLog.push({ row: excelRowNum, search: searchLabel, values: result.values, status: "err", note: "Sélecteur introuvable" });
        } else if (allValuesEmpty) {
          logLine(`Ligne ${excelRowNum} (${searchLabel}) : aucun résultat trouvé (case vide sur la page).`, "skip");
          runLog.push({ row: excelRowNum, search: searchLabel, values: result.values, status: "skip", note: "Aucun résultat" });
        } else {
          logLine(`Ligne ${excelRowNum} (${searchLabel}) : OK`, "ok");
          runLog.push({ row: excelRowNum, search: searchLabel, values: result.values, status: "ok" });
        }
      }
    } catch (err) {
      logLine(`Ligne ${excelRowNum} : erreur script — ${err.message}`, "err");
      runLog.push({ row: excelRowNum, search: searchLabel, values: [], status: "err", note: err.message });
    }

    done++; setProgress(done, total);

    if (rowDelayMs > 0 && idx < endIdx && !stopRequested) {
      await sleep(rowDelayMs);
    }
  }

  isRunning = false;
  startBtn.disabled = false;
  stopBtn.disabled = true;
  logLine("Terminé.", "ok");
  renderResultSummary();
}

/* ---------- Téléchargement du résultat ---------- */
function getOrBuildWorkbook() {
  const ws = XLSX.utils.aoa_to_sheet(rows);
  if (workbook && sheetName) {
    workbook.Sheets[sheetName] = ws;
    return workbook;
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, (sheetName || "Feuil1").slice(0, 31));
  return wb;
}

function rowsToJson(allRows) {
  if (!allRows.length) return [];
  const headers = allRows[0];
  return allRows.slice(1).map((r) => {
    const obj = {};
    headers.forEach((h, i) => { obj[h || `col${i + 1}`] = r[i] !== undefined ? r[i] : ""; });
    return obj;
  });
}

document.getElementById("downloadBtn").addEventListener("click", () => {
  if (!rows.length) { logLine("Aucune donnée chargée.", "err"); return; }
  const wb = getOrBuildWorkbook();
  XLSX.writeFile(wb, originalFileName);
});

document.getElementById("downloadJsonBtn").addEventListener("click", () => {
  if (!rows.length) { logLine("Aucune donnée chargée.", "err"); return; }
  const data = rowsToJson(rows);
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = originalFileName.replace(/\.xlsx$/i, "") + ".json";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  logLine("Export JSON téléchargé.", "ok");
});

/* ---------- Sauvegarde / chargement de la configuration ---------- */
function buildConfig() {
  return {
    mappings: getMappings(),
    conditions: getConditions(),
    outputs: getOutputs(),
    searchFields: getSearchFields(),
    submitMode: document.querySelector('input[name="submitMode"]:checked').value,
    submitSelector: document.getElementById("submitSelector").value,
    waitMs: document.getElementById("waitMs").value,
    rowDelayMs: document.getElementById("rowDelayMs").value,
    startRow: document.getElementById("startRow").value,
    endRow: document.getElementById("endRow").value,
  };
}

function applyConfig(cfg) {
  mappingList.innerHTML = "";
  (cfg.mappings || []).forEach((m) => addMappingRow(m.label, m.col));
  conditionsList.innerHTML = "";
  (cfg.conditions || []).forEach((c) => addConditionRow(c.col, c.op, c.val));
  outputsList.innerHTML = "";
  (cfg.outputs || []).forEach((o) => addOutputRow(o));
  searchFieldsList.innerHTML = "";
  // Compatibilité avec les anciennes configs sauvegardées (un seul champ de recherche).
  const legacySearchFields = cfg.searchFields
    || (cfg.searchSelector ? [{ selector: cfg.searchSelector, col: cfg.searchCol }] : []);
  legacySearchFields.forEach((f) => addSearchFieldRow(f.selector, f.col));
  document.querySelector(`input[name="submitMode"][value="${cfg.submitMode || "enter"}"]`).checked = true;
  submitSelectorRow.style.display = cfg.submitMode === "click" ? "flex" : "none";
  document.getElementById("submitSelector").value = cfg.submitSelector || "";
  document.getElementById("waitMs").value = cfg.waitMs || 1200;
  document.getElementById("rowDelayMs").value = cfg.rowDelayMs || 300;
  document.getElementById("startRow").value = cfg.startRow || 2;
  document.getElementById("endRow").value = cfg.endRow || "";
}

document.getElementById("saveConfigBtn").addEventListener("click", () => {
  chrome.storage.local.set({ savedConfig: buildConfig() }, () => logLine("Configuration sauvegardée.", "ok"));
});

document.getElementById("loadConfigBtn").addEventListener("click", () => {
  chrome.storage.local.get("savedConfig", (data) => {
    if (data.savedConfig) { applyConfig(data.savedConfig); logLine("Configuration chargée.", "ok"); }
    else logLine("Aucune configuration sauvegardée.", "err");
  });
});

/* ---------- Récapitulatif de la dernière exécution ---------- */
function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function renderResultSummary() {
  const container = document.getElementById("resultSummary");
  if (!container) return;
  if (!runLog.length) {
    container.innerHTML = '<p class="hint">Aucune exécution récente.</p>';
    return;
  }
  const headers = lastOutputs.map((o) => `Col. ${o.col.toUpperCase()}`);
  const statusLabels = { ok: "OK", err: "Erreur", skip: "Ignoré" };
  let html = '<div class="summary-table-wrap"><table class="summary-table"><thead><tr><th>Ligne</th><th>Recherche</th>';
  html += headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("");
  html += "<th>Statut</th></tr></thead><tbody>";
  runLog.forEach((entry) => {
    const statusLabel = statusLabels[entry.status] || entry.status;
    const note = entry.note ? ` — ${escapeHtml(entry.note)}` : "";
    html += `<tr class="status-${entry.status}"><td>${entry.row}</td><td>${escapeHtml(entry.search)}</td>`;
    html += lastOutputs.map((_, i) => `<td>${escapeHtml(entry.values[i] || "")}</td>`).join("");
    html += `<td class="status-cell">${statusLabel}${note}</td></tr>`;
  });
  html += "</tbody></table></div>";
  container.innerHTML = html;
}

/* ---------- Init ---------- */
function escapeAttr(str) {
  return String(str || "").replace(/"/g, "&quot;");
}

// Au démarrage : essaie de recharger la dernière configuration, sinon met des exemples.
chrome.storage.local.get("savedConfig", (data) => {
  if (data.savedConfig) {
    applyConfig(data.savedConfig);
  } else {
    addMappingRow("nom", "J");
    addMappingRow("prenom", "A");
    addMappingRow("valeur_recherche", "B");
    addConditionRow("E", "equals", "NON");
    addSearchFieldRow("", "B");
    addOutputRow({ col: "C" });
  }
});

/* ---------- UX : dashboard (barre d'onglets + panneau unique) ---------- */
(function initDashboardUX() {
  const tabButtons = Array.from(document.querySelectorAll(".tab-btn[data-step]"));
  const panels = Array.from(document.querySelectorAll(".panel[data-step]"));
  let hasDownloaded = false;
  let hasStartedRun = false;

  function stepIsDone(step) {
    switch (step) {
      case "1": return rows.length > 0;
      case "2": return getMappings().length > 0;
      case "3": return getConditions().length > 0;
      case "4": return getSearchFields().length > 0 && getOutputs().length > 0;
      case "5": return hasStartedRun && !isRunning;
      case "6": return hasDownloaded;
      default: return false;
    }
  }

  function updateDoneMarkers() {
    tabButtons.forEach((btn) => {
      const step = btn.getAttribute("data-step");
      btn.classList.toggle("done", stepIsDone(step));
    });
  }

  function showStep(step) {
    panels.forEach((p) => { p.hidden = p.getAttribute("data-step") !== step; });
    tabButtons.forEach((btn) => btn.classList.toggle("active", btn.getAttribute("data-step") === step));
    updateDoneMarkers();
  }

  tabButtons.forEach((btn) => btn.addEventListener("click", () => showStep(btn.getAttribute("data-step"))));
  document.querySelectorAll(".next-btn[data-next]").forEach((btn) => {
    btn.addEventListener("click", () => showStep(btn.getAttribute("data-next")));
  });
  document.querySelectorAll(".prev-btn[data-prev]").forEach((btn) => {
    btn.addEventListener("click", () => showStep(btn.getAttribute("data-prev")));
  });

  // Recalcule les indicateurs "terminé" sur les onglets à chaque interaction utile,
  // sans jamais changer d'onglet automatiquement (navigation restée sous contrôle de l'utilisateur).
  fileInput.addEventListener("change", () => setTimeout(updateDoneMarkers, 0));
  document.getElementById("usePasteBtn").addEventListener("click", () => setTimeout(updateDoneMarkers, 0));
  document.getElementById("useJsonBtn").addEventListener("click", () => setTimeout(updateDoneMarkers, 0));
  document.getElementById("jsonFileInput").addEventListener("change", () => setTimeout(updateDoneMarkers, 0));
  mappingList.addEventListener("input", updateDoneMarkers);
  conditionsList.addEventListener("input", updateDoneMarkers);
  outputsList.addEventListener("input", updateDoneMarkers);
  searchFieldsList.addEventListener("input", updateDoneMarkers);
  document.getElementById("addSearchFieldBtn").addEventListener("click", () => setTimeout(updateDoneMarkers, 0));

  startBtn.addEventListener("click", () => { hasStartedRun = true; updateDoneMarkers(); });
  document.getElementById("downloadBtn").addEventListener("click", () => { hasDownloaded = true; updateDoneMarkers(); });
  document.getElementById("downloadJsonBtn").addEventListener("click", () => { hasDownloaded = true; updateDoneMarkers(); });
  document.querySelector('.tab-btn[data-step="6"]').addEventListener("click", renderResultSummary);

  updateDoneMarkers();
})();
