// ============================================================
// SHEET TYPE DEFINITIONS
// ============================================================

const SHEET_TYPES = {
  PROP: {
    id: 'PROP',
    label: 'Propriétaire',
    columnCount: 26,
    multiRow: false,
    columns: [
      'N° PROP (TW)',
      'CIVILITE PROP',
      'NOM PROP',
      'PRENOM PROP',
      'ADRESSE LIGNE 1 PROP',
      'ADRESSE LIGNE 2 PROP',
      'CP PROP',
      'VILLE PROP',
      'TELEPHONE DOMICILE PROP',
      'TELEPHONE BUREAU PROP',
      'TELEPHONE PORTABLE PROP',
      'EMAIL PROP',
      'IBAN PROP',
      'FREQUENCE REGLT ACOMPTE PROP',
      'FREQUENCE REEDITION PROP',
      'MODE REGLT AU PROP',
      'TAUX HONOS PROP',
      'ASSURANCE GL (O/N)',
      'TAUX ASSURANCE GLI',
      'TAUX HONOS/ASSURANCE BASE 1',
      'DECLARATION REVENUS FONCIERS ADRF (O/N)',
      'TYPE GARANTIE',
      'DATE DEBUT MANDAT PROP',
      'NOM GESTIONNAIRE',
      'PRENOM GESTIONNAIRE',
      'Opérateur saisie'
    ],
    summaryColumns: ['N° PROP (TW)', 'NOM PROP', 'PRENOM PROP'],
    storageKey: 'propData',
    mappingKey: 'fieldMapping_PROP',
    pasteHint: '1 ligne, 26 colonnes (N° PROP, CIVILITE, NOM, PRENOM, etc.)'
  },
  LOTS: {
    id: 'LOTS',
    label: 'Lots',
    columnCount: 13,
    multiRow: true,
    columns: [
      'N° PROPRIETAIRE (Tw)',
      'NOM PROPRIETAIRE',
      'ADRESSE LOT',
      'CATEGORIE',
      'N° LOT',
      'ETAT DU LOT',
      'ETAGE DU LOT',
      'TYPE DE LOT',
      'LIBELLE TYPE DE LOT',
      'N°APPARTEMENT',
      'SURFACE DU LOT',
      'NOM LOCATAIRE',
      'REGIME FISCAL'
    ],
    summaryColumns: ['N° LOT', 'ADRESSE LOT', 'NOM LOCATAIRE'],
    storageKey: 'lotsData',
    mappingKey: 'fieldMapping_LOTS',
    pasteHint: '1 ou plusieurs lignes, 13 colonnes par ligne'
  },
  BAIL: {
    id: 'BAIL',
    label: 'Bails',
    columnCount: 48,
    multiRow: true,
    columns: [
      'N° PROP',
      'NOM PROPRIETAIRE',
      'NOM IMMEUBLE',
      'CIVILITE LOCATAIRE',
      'NOM LOCATAIRE',
      'PRENOM LOCATAIRE',
      'DATE DE NAISSANCE',
      'LIEU DE NAISSANCE',
      'ADRESSE LIGNE 1 LOCATAIRE',
      'ADRESSE LIGNE 2 LOCATAIRE',
      'CP LOCATAIRE',
      'VILLE LOCATAIRE',
      'TELEPHONE DOMICILE LOCATAIRE',
      'TELEPHONE N°2',
      'TELEPHONE PORTABLE LOCATAIRE',
      'EMAIL LOCATAIRE',
      'IBAN MANDAT SEPA LOCATAIRE',
      'BIC MANDAT SEPA LOCATAIRE',
      'DATE ENTREE LOCATAIRE',
      'CODE TYPE BAIL',
      'LIBELLE BAIL',
      'N° INDICE (5 IRL/1 ICC INSEE/11 ILC COMMERCIAUX )',
      'DATE PROCHAINE REVISION LOYER',
      'DATE DERNIERE REVISION LOYER',
      'FREQUENCE REVISION LOYER',
      'ANNEE DERNIERE REVISION LOYER',
      'TRIMESTRE REFERENCE DERNIERE REVISION LOYER',
      'ANNEE PROCHAINE REVISION LOYER',
      'TRIMESTRE PROCHAINE REVISION LOYER',
      'MODE REGLT LOCATAIRE',
      'TERME AVANCE/ECHU LOYER',
      'FREQUENCE APPEL LOYER',
      'DEPOT DE GARANTIE CONSERVE EN AGENCE',
      'DEPOT DE GARANTIE REVERSE AU PROPRIETAIRE',
      'DATE DEBUT ASSURANCE MULTIRISQUES',
      'DATE FIN ASSURANCE MULTIRISQUES',
      'SURFACE DU LOT',
      'TYPE DE LOT',
      'N° PORTE',
      'NOMBRE DE GARANTS',
      'CIVILITE GARANT',
      'NOM GARANT',
      'PRENOM GARANT',
      'ADRESSE LIGNE 1 GARANT',
      'ADRESSE LIGNE 2 GARANT',
      'CP GARANT',
      'VILLE GARANT',
      'RUM MANDAT SEPA LOCATAIRE'
    ],
    summaryColumns: ['NOM LOCATAIRE', 'PRENOM LOCATAIRE', 'NOM IMMEUBLE'],
    storageKey: 'bailData',
    mappingKey: 'fieldMapping_BAIL',
    pasteHint: '1 ou plusieurs lignes, 48 colonnes par ligne'
  }
};

// ============================================================
// STATE
// ============================================================

let activeSheetTab = 'PROP';
let activeConfigSheet = 'PROP';

const sheetState = {
  PROP: { data: null, selectedIndex: null },
  LOTS: { data: null, selectedIndex: null },
  BAIL: { data: null, selectedIndex: null }
};

const sheetMappings = {
  PROP: {},
  LOTS: {},
  BAIL: {}
};

/** @type {{ name: string, value: string }[]} */
let customFieldsArray = [];

// ============================================================
// DOM REFERENCES (resolved after DOMContentLoaded)
// ============================================================

let statusMessage;
let fillButton;
let configButton;
let configPanel;
let closeConfigButton;
let cancelConfigButton;
let saveConfigButton;
let mappingFields;
let customFieldsList;
let addCustomFieldButton;

// ============================================================
// UTILITY FUNCTIONS
// ============================================================

/**
 * Format IBAN with spaces every 4 characters
 */
function formatIban(str) {
  if (!str || typeof str !== 'string') return str;
  const s = str.replace(/\s/g, '').trim();
  if (s.length < 4) return str;
  const parts = [s.slice(0, 4)];
  for (let i = 4; i < s.length; i += 4) {
    parts.push(s.slice(i, i + 4));
  }
  return parts.join(' ');
}

/**
 * Format trimestre value as "T<number>"
 * Ex: "1" -> "T1", "T2" stays "T2"
 */
function formatTrimester(value) {
  if (value === undefined || value === null) return value;
  const s = String(value).trim();
  if (!s) return value;

  // Already in the right format
  if (/^T\d+$/i.test(s)) {
    return s.toUpperCase();
  }

  // Extract digits and prefix with T
  const digits = s.replace(/[^\d]/g, '');
  if (!digits) return value;
  return `T${digits}`;
}

/**
 * Copy text to clipboard
 */
async function copyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch (err) {
    console.error('Clipboard copy failed:', err);
    return false;
  }
}

/**
 * Show status message
 */
function showStatus(message, type = 'info') {
  statusMessage.textContent = message;
  statusMessage.className = `status-message show ${type}`;
  setTimeout(() => {
    statusMessage.classList.remove('show');
  }, 2000);
}

function escapeHtml(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

// ============================================================
// DATA PARSING
// ============================================================

/**
 * Parse tab-separated Excel data for a specific sheet type.
 * Returns a single object for PROP, or an array for LOTS/BAIL.
 */
function parseExcelData(rawData, sheetType) {
  const sheet = SHEET_TYPES[sheetType];
  if (!rawData || rawData.trim() === '' || !sheet) return null;

  const lines = rawData.trim().split(/\r?\n/).filter(line => line.trim() !== '');

  if (sheet.multiRow) {
    const rows = [];
    for (let i = 0; i < lines.length; i++) {
      const values = lines[i].split('\t');
      // Excel omits trailing tabs for empty cells — pad with empty strings
      if (values.length > sheet.columnCount) {
        throw new Error(
          `Ligne ${i + 1} invalide (${sheetType}) : ${values.length} colonnes trouvées, ${sheet.columnCount} attendues`
        );
      }
      while (values.length < sheet.columnCount) {
        values.push('');
      }
      const row = {};
      sheet.columns.forEach((col, idx) => { row[col] = values[idx].trim(); });
      rows.push(row);
    }
    return rows.length > 0 ? rows : null;
  } else {
    if (lines.length > 1) {
      throw new Error('PROP : une seule ligne attendue');
    }
    const values = lines[0].split('\t');
    // Excel omits trailing tabs for empty cells — pad with empty strings
    if (values.length > sheet.columnCount) {
      throw new Error(
        `Format invalide (${sheetType}) : ${values.length} colonnes trouvées, ${sheet.columnCount} attendues`
      );
    }
    while (values.length < sheet.columnCount) {
      values.push('');
    }
    const data = {};
    sheet.columns.forEach((col, idx) => { data[col] = values[idx].trim(); });
    return data;
  }
}

// ============================================================
// UI: FIELD DISPLAY
// ============================================================

/**
 * Create a field item with copy button
 */
function createFieldItem(label, value) {
  // Format IBAN fields (any label containing "IBAN")
  const isIban = label.toUpperCase().includes('IBAN');
  const displayValue = isIban ? formatIban(value) : value;
  const copyValue = displayValue;

  const item = document.createElement('div');
  item.className = 'field-item';

  const fieldInfo = document.createElement('div');
  fieldInfo.className = 'field-info-container';

  const labelEl = document.createElement('div');
  labelEl.className = 'field-label';
  labelEl.textContent = label;

  const valueEl = document.createElement('div');
  valueEl.className = 'field-value';
  valueEl.textContent = displayValue || '(vide)';
  if (!value) valueEl.classList.add('empty');

  fieldInfo.appendChild(labelEl);
  fieldInfo.appendChild(valueEl);

  const copyBtn = document.createElement('button');
  copyBtn.className = 'field-copy-btn';
  copyBtn.title = 'Copier cette valeur';
  copyBtn.innerHTML = `
    <svg class="copy-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
      <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"></path>
      <rect x="8" y="2" width="8" height="4" rx="1" ry="1"></rect>
    </svg>
  `;

  copyBtn.addEventListener('click', async () => {
    const success = await copyToClipboard(copyValue);
    if (success) {
      copyBtn.classList.add('copied');
      showStatus(`\u2713 "${label}" copié !`, 'success');
      setTimeout(() => copyBtn.classList.remove('copied'), 1000);
    } else {
      showStatus('Erreur lors de la copie', 'error');
    }
  });

  item.appendChild(fieldInfo);
  item.appendChild(copyBtn);
  return item;
}

/**
 * Display parsed data fields in a target element
 */
function displayFields(data, columns, targetElement) {
  targetElement.innerHTML = '';
  columns.forEach(column => {
    const fieldItem = createFieldItem(column, data[column]);
    targetElement.appendChild(fieldItem);
  });
}

// ============================================================
// UI: ROW SELECTOR (LOTS / BAIL)
// ============================================================

/**
 * Display a radio-button row selector for multi-row sheets, with delete buttons
 */
function displayRowSelector(rows, sheetType, container) {
  const sheet = SHEET_TYPES[sheetType];
  container.innerHTML = '';

  if (!rows || rows.length === 0) return;

  const selectorDiv = document.createElement('div');
  selectorDiv.className = 'row-selector';

  const label = document.createElement('div');
  label.className = 'row-selector-label';
  label.textContent = `${rows.length} ${sheet.label.toLowerCase()} ajouté(s) :`;
  selectorDiv.appendChild(label);

  rows.forEach((row, index) => {
    const item = document.createElement('div');
    item.className = 'row-selector-item';

    const radioLabel = document.createElement('label');
    radioLabel.className = 'row-selector-label-wrap';

    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = `${sheetType}-row-select`;
    radio.value = index;
    radio.className = 'row-radio';
    if (sheetState[sheetType].selectedIndex === index) {
      radio.checked = true;
    }

    radio.addEventListener('change', () => {
      sheetState[sheetType].selectedIndex = index;
      chrome.storage.local.set({ [`selected${sheetType}Index`]: index });
      const fieldsList = document.getElementById(`${sheetType}-fieldsList`);
      if (fieldsList) {
        displayFields(row, sheet.columns, fieldsList);
      }
    });

    const summary = document.createElement('span');
    summary.className = 'row-summary';
    const summaryParts = sheet.summaryColumns
      .map(col => row[col] || '')
      .filter(Boolean);
    summary.textContent = summaryParts.join(' \u2014 ') || `Ligne ${index + 1}`;

    radioLabel.appendChild(radio);
    radioLabel.appendChild(summary);

    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'row-delete-btn';
    deleteBtn.type = 'button';
    deleteBtn.title = 'Supprimer';
    deleteBtn.innerHTML = `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <line x1="18" y1="6" x2="6" y2="18"></line>
        <line x1="6" y1="6" x2="18" y2="18"></line>
      </svg>
    `;
    deleteBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      handleRemoveRow(sheetType, index);
    });

    item.appendChild(radioLabel);
    item.appendChild(deleteBtn);
    selectorDiv.appendChild(item);
  });

  container.appendChild(selectorDiv);
}

/**
 * Remove a row from a multi-row sheet
 */
async function handleRemoveRow(sheetType, index) {
  const sheet = SHEET_TYPES[sheetType];
  const rows = sheetState[sheetType].data;
  if (!rows) return;

  rows.splice(index, 1);

  if (rows.length === 0) {
    sheetState[sheetType].data = null;
    sheetState[sheetType].selectedIndex = null;
    await chrome.storage.local.remove([sheet.storageKey, `selected${sheetType}Index`]);
  } else {
    // Adjust selectedIndex
    if (sheetState[sheetType].selectedIndex >= rows.length) {
      sheetState[sheetType].selectedIndex = rows.length - 1;
    }
    if (sheetState[sheetType].selectedIndex === null) {
      sheetState[sheetType].selectedIndex = 0;
    }
    await chrome.storage.local.set({
      [sheet.storageKey]: rows,
      [`selected${sheetType}Index`]: sheetState[sheetType].selectedIndex
    });
  }

  // Refresh UI
  const selectorContainer = document.getElementById(`${sheetType}-rowSelector`);
  const fieldsList = document.getElementById(`${sheetType}-fieldsList`);

  if (rows.length > 0) {
    if (selectorContainer) displayRowSelector(rows, sheetType, selectorContainer);
    const selectedRow = rows[sheetState[sheetType].selectedIndex];
    if (fieldsList && selectedRow) displayFields(selectedRow, sheet.columns, fieldsList);
  } else {
    if (selectorContainer) selectorContainer.innerHTML = '';
    if (fieldsList) fieldsList.innerHTML = '';
  }

  updateFillButtonState();
  updateTabIndicators();
  updateClearAllButtons();
}

// ============================================================
// UI: SHEET TAB NAVIGATION
// ============================================================

function switchSheetTab(sheetType) {
  activeSheetTab = sheetType;

  ['PROP', 'LOTS', 'BAIL', 'CUSTOM'].forEach(type => {
    const tab = document.getElementById(`sheetTab-${type}`);
    const panel = document.getElementById(`sheetPanel-${type}`);
    if (tab) {
      tab.classList.toggle('active', type === sheetType);
      tab.setAttribute('aria-selected', type === sheetType);
    }
    if (panel) {
      panel.classList.toggle('active', type === sheetType);
      panel.hidden = (type !== sheetType);
    }
  });
}

/**
 * Switch between paste view and data view for a sheet
 */
function switchSheetView(sheetType, view) {
  const pasteSection = document.getElementById(`${sheetType}-pasteSection`);
  const fieldsSection = document.getElementById(`${sheetType}-fieldsSection`);
  if (pasteSection) pasteSection.style.display = view === 'paste' ? 'block' : 'none';
  if (fieldsSection) fieldsSection.style.display = view === 'data' ? 'block' : 'none';
}

/**
 * Update the green dot indicator on sheet tabs
 */
function updateTabIndicators() {
  ['PROP', 'LOTS', 'BAIL'].forEach(type => {
    const tab = document.getElementById(`sheetTab-${type}`);
    if (tab) {
      tab.classList.toggle('has-data', sheetState[type].data !== null);
    }
  });
}

/**
 * Show/hide the "Tout effacer" buttons for LOTS and BAIL
 */
function updateClearAllButtons() {
  ['LOTS', 'BAIL'].forEach(type => {
    const btn = document.getElementById(`${type}-clearAllButton`);
    if (btn) {
      const hasData = sheetState[type].data && sheetState[type].data.length > 0;
      btn.style.display = hasData ? 'flex' : 'none';
    }
  });
}

// ============================================================
// PARSE / RESET HANDLERS
// ============================================================

/**
 * Handle parse for PROP (single row, replaces data)
 */
async function handleParse() {
  const textarea = document.getElementById(`${activeSheetTab}-textarea`);
  if (!textarea) return;

  const rawData = textarea.value;
  if (!rawData || rawData.trim() === '') {
    showStatus('Veuillez coller des données dans la zone de texte', 'error');
    return;
  }

  try {
    const sheet = SHEET_TYPES[activeSheetTab];
    const result = parseExcelData(rawData, activeSheetTab);

    if (!result) {
      showStatus('Données invalides ou vides', 'error');
      return;
    }

    sheetState[activeSheetTab].data = result;
    await chrome.storage.local.set({ [sheet.storageKey]: result });

    switchSheetView(activeSheetTab, 'data');

    const fieldsList = document.getElementById(`${activeSheetTab}-fieldsList`);
    if (fieldsList) {
      displayFields(result, sheet.columns, fieldsList);
    }
    showStatus('Données analysées avec succès !', 'success');

    updateFillButtonState();
    updateTabIndicators();
    updateClearAllButtons();
  } catch (error) {
    showStatus(error.message, 'error');
  }
}

/**
 * Handle adding rows for multi-row sheets (LOTS/BAIL).
 * Parses pasted data and appends to existing rows.
 */
async function handleAddRows(sheetType) {
  const textarea = document.getElementById(`${sheetType}-textarea`);
  if (!textarea) return;

  const rawData = textarea.value;
  if (!rawData || rawData.trim() === '') {
    showStatus('Veuillez coller des données dans la zone de texte', 'error');
    return;
  }

  try {
    const sheet = SHEET_TYPES[sheetType];
    const newRows = parseExcelData(rawData, sheetType);

    if (!newRows || newRows.length === 0) {
      showStatus('Données invalides ou vides', 'error');
      return;
    }

    // Append to existing rows
    const existing = sheetState[sheetType].data || [];
    const merged = [...existing, ...newRows];
    sheetState[sheetType].data = merged;

    // Auto-select first if none selected
    if (sheetState[sheetType].selectedIndex === null) {
      sheetState[sheetType].selectedIndex = 0;
    }

    await chrome.storage.local.set({
      [sheet.storageKey]: merged,
      [`selected${sheetType}Index`]: sheetState[sheetType].selectedIndex
    });

    // Clear textarea
    textarea.value = '';
    const addBtn = document.getElementById(`${sheetType}-addButton`);
    if (addBtn) addBtn.disabled = true;

    // Refresh UI
    const selectorContainer = document.getElementById(`${sheetType}-rowSelector`);
    if (selectorContainer) displayRowSelector(merged, sheetType, selectorContainer);

    const fieldsList = document.getElementById(`${sheetType}-fieldsList`);
    const selectedRow = merged[sheetState[sheetType].selectedIndex];
    if (fieldsList && selectedRow) {
      displayFields(selectedRow, sheet.columns, fieldsList);
    }

    const added = newRows.length;
    showStatus(`${added} ${sheet.label.toLowerCase()} ajouté${added > 1 ? 's' : ''} !`, 'success');
    updateFillButtonState();
    updateTabIndicators();
    updateClearAllButtons();
  } catch (error) {
    showStatus(error.message, 'error');
  }
}

async function handleReset() {
  const sheet = SHEET_TYPES[activeSheetTab];
  if (!sheet) return;

  sheetState[activeSheetTab].data = null;
  sheetState[activeSheetTab].selectedIndex = null;

  const keysToRemove = [sheet.storageKey];
  if (sheet.multiRow) keysToRemove.push(`selected${activeSheetTab}Index`);
  await chrome.storage.local.remove(keysToRemove);

  const textarea = document.getElementById(`${activeSheetTab}-textarea`);
  if (textarea) textarea.value = '';

  if (!sheet.multiRow) {
    switchSheetView(activeSheetTab, 'paste');
  } else {
    // Clear the row list and fields
    const selectorContainer = document.getElementById(`${activeSheetTab}-rowSelector`);
    const fieldsList = document.getElementById(`${activeSheetTab}-fieldsList`);
    if (selectorContainer) selectorContainer.innerHTML = '';
    if (fieldsList) fieldsList.innerHTML = '';
  }

  updateFillButtonState();
  updateTabIndicators();
  updateClearAllButtons();
}

// ============================================================
// FILL BUTTON STATE
// ============================================================

function updateFillButtonState() {
  if (!fillButton) return;

  const hasPropData = sheetState.PROP.data !== null;
  const hasPropMapping = Object.keys(sheetMappings.PROP).length > 0;
  const shouldEnable = hasPropData && hasPropMapping;

  fillButton.disabled = !shouldEnable;

  if (shouldEnable) {
    fillButton.title = 'Remplir le formulaire avec les données';
  } else if (!hasPropData) {
    fillButton.title = 'Analysez d\'abord des données PROP';
  } else {
    fillButton.title = 'Configurez d\'abord le mapping PROP';
  }

  // Per-sheet fill buttons
  const propFillBtn = document.getElementById('PROP-fillSheetButton');
  if (propFillBtn) {
    propFillBtn.disabled = !(hasPropData && hasPropMapping);
  }

  const bailFillBtn = document.getElementById('BAIL-fillSheetButton');
  if (bailFillBtn) {
    const hasBailData = sheetState.BAIL.data && sheetState.BAIL.selectedIndex !== null;
    const hasBailMapping = Object.keys(sheetMappings.BAIL).length > 0;
    const show = hasBailData && hasBailMapping;
    bailFillBtn.style.display = show ? 'flex' : 'none';
  }
}

// ============================================================
// FILL FORM
// ============================================================

async function handleFillForm() {
  if (!sheetState.PROP.data) {
    showStatus('Veuillez d\'abord analyser des données PROP', 'error');
    return;
  }

  if (Object.keys(sheetMappings.PROP).length === 0) {
    showStatus('Veuillez d\'abord configurer le mapping PROP', 'error');
    return;
  }

  try {
    // Find the active tab to determine the target origin
    const [activeTab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!activeTab) {
      showStatus('Aucun onglet actif trouvé', 'error');
      return;
    }

    const activeUrl = activeTab.url || '';
    if (activeUrl.startsWith('chrome://') || activeUrl.startsWith('about:') || activeUrl.startsWith('chrome-extension://')) {
      showStatus('Impossible d\'injecter sur cette page (page protégée)', 'error');
      return;
    }

    // Extract origin to find all same-origin tabs (including popups)
    let origin;
    try {
      origin = new URL(activeUrl).origin;
    } catch {
      showStatus('URL invalide', 'error');
      return;
    }

    // Find ALL tabs matching the same origin (covers window.open popups)
    const allTabs = await chrome.tabs.query({ url: origin + '/*' });
    const targetTabs = allTabs.filter(t => {
      const u = t.url || '';
      return u.startsWith('http://') || u.startsWith('https://');
    });

    if (targetTabs.length === 0) {
      showStatus('Aucun onglet cible trouvé', 'error');
      return;
    }

    // Build merged data and mapping
    const mergedData = {};
    const mergedMapping = {};

    // PROP (always present)
    const propData = { ...sheetState.PROP.data };
    if (propData['IBAN PROP']) propData['IBAN PROP'] = formatIban(propData['IBAN PROP']);
    Object.assign(mergedData, propData);
    Object.assign(mergedMapping, sheetMappings.PROP);

    // LOTS (optional)
    if (sheetState.LOTS.data && sheetState.LOTS.selectedIndex !== null) {
      const lotRow = sheetState.LOTS.data[sheetState.LOTS.selectedIndex];
      if (lotRow) {
        Object.assign(mergedData, lotRow);
        Object.assign(mergedMapping, sheetMappings.LOTS);
      }
    }

    // BAIL (optional)
    if (sheetState.BAIL.data && sheetState.BAIL.selectedIndex !== null) {
      const bailRow = { ...sheetState.BAIL.data[sheetState.BAIL.selectedIndex] };
      if (bailRow) {
        if (bailRow['IBAN MANDAT SEPA LOCATAIRE']) {
          bailRow['IBAN MANDAT SEPA LOCATAIRE'] = formatIban(bailRow['IBAN MANDAT SEPA LOCATAIRE']);
        }
        // Format trimesters as "T<number>"
        if (bailRow['TRIMESTRE REFERENCE DERNIERE REVISION LOYER']) {
          bailRow['TRIMESTRE REFERENCE DERNIERE REVISION LOYER'] =
            formatTrimester(bailRow['TRIMESTRE REFERENCE DERNIERE REVISION LOYER']);
        }
        if (bailRow['TRIMESTRE PROCHAINE REVISION LOYER']) {
          bailRow['TRIMESTRE PROCHAINE REVISION LOYER'] =
            formatTrimester(bailRow['TRIMESTRE PROCHAINE REVISION LOYER']);
        }
        Object.assign(mergedData, bailRow);
        Object.assign(mergedMapping, sheetMappings.BAIL);
      }
    }

    // Custom fields
    syncCustomFieldsFromDOM();
    saveCustomFields();
    const customFields = getCustomFieldsObject();

    const message = {
      action: 'fillForm',
      data: mergedData,
      mapping: mergedMapping,
      customFields: Object.keys(customFields).length > 0 ? customFields : undefined
    };

    // Inject and fill all same-origin tabs (including popups and iframes)
    let totalFilled = 0;
    let totalErrors = [];
    let tabsReached = 0;

    const fillPromises = targetTabs.map(async (tab) => {
      try {
        // Inject content script with allFrames for iframe support
        await chrome.scripting.executeScript({
          target: { tabId: tab.id, allFrames: true },
          files: ['content.js']
        });
        await new Promise(resolve => setTimeout(resolve, 100));
      } catch (injErr) {
        console.warn(`NoHands: Injection skipped for tab ${tab.id}:`, injErr.message);
        return;
      }

      try {
        const response = await chrome.tabs.sendMessage(tab.id, message);
        if (response) {
          tabsReached++;
          totalFilled += response.filledCount || 0;
          if (response.errors) totalErrors.push(...response.errors);
        }
      } catch (msgErr) {
        console.warn(`NoHands: Message skipped for tab ${tab.id}:`, msgErr.message);
      }
    });

    await Promise.all(fillPromises);

    if (totalFilled > 0) {
      const tabInfo = targetTabs.length > 1 ? ` (${tabsReached} onglet${tabsReached > 1 ? 's' : ''})` : '';
      showStatus(`\u2713 ${totalFilled} champ${totalFilled > 1 ? 's' : ''} rempli${totalFilled > 1 ? 's' : ''}${tabInfo} !`, 'success');
    } else if (totalErrors.length > 0) {
      showStatus(`Erreur: ${totalErrors.slice(0, 3).join(', ')}`, 'error');
    } else {
      showStatus('Aucun champ rempli. Vérifiez le mapping.', 'error');
    }
  } catch (error) {
    console.error('Error filling form:', error);
    showStatus('Erreur: ' + error.message, 'error');
  }
}

/**
 * Send a fill message to all same-origin tabs.
 * Returns { totalFilled, totalErrors, tabsReached }.
 */
async function sendFillToAllTabs(message) {
  const [activeTab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!activeTab) throw new Error('Aucun onglet actif trouvé');

  const activeUrl = activeTab.url || '';
  if (activeUrl.startsWith('chrome://') || activeUrl.startsWith('about:') || activeUrl.startsWith('chrome-extension://')) {
    throw new Error('Page protégée');
  }

  const origin = new URL(activeUrl).origin;
  const allTabs = await chrome.tabs.query({ url: origin + '/*' });
  const targetTabs = allTabs.filter(t => {
    const u = t.url || '';
    return u.startsWith('http://') || u.startsWith('https://');
  });

  if (targetTabs.length === 0) throw new Error('Aucun onglet cible trouvé');

  let totalFilled = 0;
  let totalErrors = [];
  let tabsReached = 0;

  const fillPromises = targetTabs.map(async (tab) => {
    try {
      await chrome.scripting.executeScript({
        target: { tabId: tab.id, allFrames: true },
        files: ['content.js']
      });
      await new Promise(resolve => setTimeout(resolve, 100));
    } catch (injErr) {
      console.warn(`NoHands: Injection skipped for tab ${tab.id}:`, injErr.message);
      return;
    }
    try {
      const response = await chrome.tabs.sendMessage(tab.id, message);
      if (response) {
        tabsReached++;
        totalFilled += response.filledCount || 0;
        if (response.errors) totalErrors.push(...response.errors);
      }
    } catch (msgErr) {
      console.warn(`NoHands: Message skipped for tab ${tab.id}:`, msgErr.message);
    }
  });

  await Promise.all(fillPromises);
  return { totalFilled, totalErrors, tabsReached, tabCount: targetTabs.length };
}

/**
 * Fill only one sheet's data (PROP or BAIL)
 */
async function handleFillSheet(sheetType) {
  const sheet = SHEET_TYPES[sheetType];
  const mapping = sheetMappings[sheetType];

  if (!mapping || Object.keys(mapping).length === 0) {
    showStatus(`Configurez d'abord le mapping ${sheet.label}`, 'error');
    return;
  }

  let data;
  if (sheet.multiRow) {
    if (!sheetState[sheetType].data || sheetState[sheetType].selectedIndex === null) {
      showStatus(`Aucune donnée ${sheet.label} sélectionnée`, 'error');
      return;
    }
    data = { ...sheetState[sheetType].data[sheetState[sheetType].selectedIndex] };
  } else {
    if (!sheetState[sheetType].data) {
      showStatus(`Aucune donnée ${sheet.label}`, 'error');
      return;
    }
    data = { ...sheetState[sheetType].data };
  }

  // Apply formatting
  if (sheetType === 'PROP') {
    if (data['IBAN PROP']) data['IBAN PROP'] = formatIban(data['IBAN PROP']);
  }
  if (sheetType === 'BAIL') {
    if (data['IBAN MANDAT SEPA LOCATAIRE']) {
      data['IBAN MANDAT SEPA LOCATAIRE'] = formatIban(data['IBAN MANDAT SEPA LOCATAIRE']);
    }
    if (data['TRIMESTRE REFERENCE DERNIERE REVISION LOYER']) {
      data['TRIMESTRE REFERENCE DERNIERE REVISION LOYER'] =
        formatTrimester(data['TRIMESTRE REFERENCE DERNIERE REVISION LOYER']);
    }
    if (data['TRIMESTRE PROCHAINE REVISION LOYER']) {
      data['TRIMESTRE PROCHAINE REVISION LOYER'] =
        formatTrimester(data['TRIMESTRE PROCHAINE REVISION LOYER']);
    }
  }

  try {
    const message = {
      action: 'fillForm',
      data: data,
      mapping: mapping
    };

    const { totalFilled, totalErrors, tabsReached, tabCount } = await sendFillToAllTabs(message);

    if (totalFilled > 0) {
      const tabInfo = tabCount > 1 ? ` (${tabsReached} onglet${tabsReached > 1 ? 's' : ''})` : '';
      showStatus(`\u2713 ${totalFilled} champ${totalFilled > 1 ? 's' : ''} ${sheet.label} rempli${totalFilled > 1 ? 's' : ''}${tabInfo} !`, 'success');
    } else if (totalErrors.length > 0) {
      showStatus(`Erreur: ${totalErrors.slice(0, 3).join(', ')}`, 'error');
    } else {
      showStatus(`Aucun champ ${sheet.label} rempli. Vérifiez le mapping.`, 'error');
    }
  } catch (error) {
    console.error('Error filling sheet:', error);
    showStatus('Erreur: ' + error.message, 'error');
  }
}

// ============================================================
// CONFIG PANEL (MAPPING)
// ============================================================

function openConfigPanel() {
  // Build sheet tabs inside modal
  const configSheetTabs = document.getElementById('configSheetTabs');
  if (configSheetTabs) {
    configSheetTabs.innerHTML = '';
    ['PROP', 'LOTS', 'BAIL'].forEach(type => {
      const tab = document.createElement('button');
      tab.type = 'button';
      tab.className = `config-sheet-tab ${type === activeConfigSheet ? 'active' : ''}`;
      tab.textContent = SHEET_TYPES[type].label;
      tab.addEventListener('click', () => renderMappingForSheet(type));
      configSheetTabs.appendChild(tab);
    });
  }

  renderMappingForSheet(activeConfigSheet);
  configPanel.style.display = 'flex';
}

function renderMappingForSheet(sheetType) {
  activeConfigSheet = sheetType;

  // Update tab active state
  const tabs = document.querySelectorAll('.config-sheet-tab');
  tabs.forEach(t => {
    t.classList.toggle('active', t.textContent === SHEET_TYPES[sheetType].label);
  });

  mappingFields.innerHTML = '';
  const sheet = SHEET_TYPES[sheetType];
  const currentMapping = sheetMappings[sheetType] || {};

  sheet.columns.forEach(column => {
    const fieldContainer = document.createElement('div');
    fieldContainer.className = 'mapping-field-container';

    const label = document.createElement('label');
    label.className = 'mapping-label';
    label.textContent = column;
    fieldContainer.appendChild(label);

    let existingMappings = currentMapping[column] || [];
    if (typeof existingMappings === 'string') {
      existingMappings = existingMappings ? [existingMappings] : [];
    }
    if (!Array.isArray(existingMappings)) existingMappings = [];

    const inputsContainer = document.createElement('div');
    inputsContainer.className = 'mapping-inputs-container';
    inputsContainer.dataset.column = column;

    if (existingMappings.length > 0) {
      existingMappings.forEach(value => addMappingInput(inputsContainer, column, value));
    } else {
      addMappingInput(inputsContainer, column, '');
    }

    fieldContainer.appendChild(inputsContainer);

    const addButton = document.createElement('button');
    addButton.className = 'add-mapping-btn';
    addButton.type = 'button';
    addButton.innerHTML = `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <circle cx="12" cy="12" r="10"></circle>
        <line x1="12" y1="8" x2="12" y2="16"></line>
        <line x1="8" y1="12" x2="16" y2="12"></line>
      </svg>
      Ajouter un champ
    `;
    addButton.addEventListener('click', () => addMappingInput(inputsContainer, column, ''));

    fieldContainer.appendChild(addButton);
    mappingFields.appendChild(fieldContainer);
  });
}

function addMappingInput(container, column, value) {
  const inputGroup = document.createElement('div');
  inputGroup.className = 'mapping-input-group';

  const input = document.createElement('input');
  input.type = 'text';
  input.placeholder = 'Ex: body:x:tabc:x:infoBail:x:txtNom';
  input.value = value;
  input.dataset.column = column;

  const removeButton = document.createElement('button');
  removeButton.className = 'remove-mapping-btn';
  removeButton.type = 'button';
  removeButton.innerHTML = `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <circle cx="12" cy="12" r="10"></circle>
      <line x1="15" y1="9" x2="9" y2="15"></line>
      <line x1="9" y1="9" x2="15" y2="15"></line>
    </svg>
  `;
  removeButton.addEventListener('click', () => {
    if (container.children.length > 1) {
      inputGroup.remove();
    } else {
      input.value = '';
    }
  });

  inputGroup.appendChild(input);
  inputGroup.appendChild(removeButton);
  container.appendChild(inputGroup);
}

function closeConfigPanel() {
  configPanel.style.display = 'none';
}

async function handleConfigSave() {
  const newMapping = {};

  const inputsContainers = mappingFields.querySelectorAll('.mapping-inputs-container');
  inputsContainers.forEach(container => {
    const column = container.dataset.column;
    const inputs = container.querySelectorAll('input');
    const values = [];
    inputs.forEach(input => {
      const value = input.value.trim();
      if (value) values.push(value);
    });
    if (values.length > 0) {
      newMapping[column] = values.length === 1 ? values[0] : values;
    }
  });

  const sheet = SHEET_TYPES[activeConfigSheet];
  sheetMappings[activeConfigSheet] = newMapping;

  try {
    await chrome.storage.local.set({ [sheet.mappingKey]: newMapping });
    showStatus(`Mapping ${sheet.label} sauvegardé !`, 'success');
    updateFillButtonState();
    console.log(`Mapping ${activeConfigSheet} saved:`, newMapping);
  } catch (error) {
    console.error('Error saving mapping:', error);
    showStatus('Erreur lors de la sauvegarde', 'error');
  }

  closeConfigPanel();
}

// ============================================================
// CUSTOM FIELDS
// ============================================================

function getCustomFieldsFromDOM() {
  if (!customFieldsList) return [];
  const rows = customFieldsList.querySelectorAll('.custom-field-row');
  const result = [];
  rows.forEach(row => {
    const nameInput = row.querySelector('.custom-field-name');
    const valueInput = row.querySelector('.custom-field-value');
    result.push({
      name: nameInput ? nameInput.value.trim() : '',
      value: valueInput ? valueInput.value.trim() : ''
    });
  });
  return result;
}

function getCustomFieldsObject() {
  const arr = getCustomFieldsFromDOM();
  const obj = {};
  arr.forEach(({ name, value }) => {
    if (name) obj[name] = value;
  });
  return obj;
}

async function saveCustomFields() {
  const data = getCustomFieldsFromDOM();
  try {
    await chrome.storage.local.set({ customFields: data });
    customFieldsArray = data;
  } catch (error) {
    console.error('Error saving custom fields:', error);
  }
}

function renderCustomFields() {
  if (!customFieldsList) return;
  customFieldsList.innerHTML = '';
  const list = customFieldsArray.length > 0 ? customFieldsArray : [{ name: '', value: '' }];
  list.forEach((item, index) => {
    const row = document.createElement('div');
    row.className = 'custom-field-row';
    row.innerHTML = `
      <input type="text" class="custom-field-name" placeholder="Nom de l'input" value="${escapeHtml(item.name)}" data-index="${index}">
      <input type="text" class="custom-field-value" placeholder="Valeur" value="${escapeHtml(item.value)}" data-index="${index}">
      <button type="button" class="remove-custom-field-btn" data-index="${index}" title="Supprimer">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
      </button>
    `;
    const nameInput = row.querySelector('.custom-field-name');
    const valueInput = row.querySelector('.custom-field-value');
    const removeBtn = row.querySelector('.remove-custom-field-btn');
    nameInput.addEventListener('blur', () => { syncCustomFieldsFromDOM(); saveCustomFields(); });
    valueInput.addEventListener('blur', () => { syncCustomFieldsFromDOM(); saveCustomFields(); });
    removeBtn.addEventListener('click', () => {
      const r = removeBtn.closest('.custom-field-row');
      const idx = r ? Array.from(customFieldsList.children).indexOf(r) : index;
      removeCustomField(idx);
    });
    customFieldsList.appendChild(row);
  });
}

function syncCustomFieldsFromDOM() {
  customFieldsArray = getCustomFieldsFromDOM();
  if (customFieldsArray.length === 0) customFieldsArray = [{ name: '', value: '' }];
}

function addCustomField() {
  syncCustomFieldsFromDOM();
  customFieldsArray.push({ name: '', value: '' });
  renderCustomFields();
  saveCustomFields();
}

function removeCustomField(index) {
  syncCustomFieldsFromDOM();
  customFieldsArray.splice(index, 1);
  if (customFieldsArray.length === 0) customFieldsArray = [{ name: '', value: '' }];
  renderCustomFields();
  saveCustomFields();
}

// ============================================================
// STORAGE: LOAD + MIGRATION
// ============================================================

async function loadMapping() {
  try {
    const keys = [
      // New keys
      'propData', 'lotsData', 'bailData',
      'fieldMapping_PROP', 'fieldMapping_LOTS', 'fieldMapping_BAIL',
      'selectedLOTSIndex', 'selectedBAILIndex',
      'customFields',
      // Legacy keys
      'parsedData', 'fieldMapping'
    ];

    const result = await chrome.storage.local.get(keys);

    // Migration: old keys → new keys
    if (result.parsedData && !result.propData) {
      result.propData = result.parsedData;
      await chrome.storage.local.set({ propData: result.propData });
      await chrome.storage.local.remove('parsedData');
    }
    if (result.fieldMapping && !result.fieldMapping_PROP) {
      result.fieldMapping_PROP = result.fieldMapping;
      await chrome.storage.local.set({ fieldMapping_PROP: result.fieldMapping_PROP });
      await chrome.storage.local.remove('fieldMapping');
    }

    // Load per-sheet data
    sheetState.PROP.data = result.propData || null;
    sheetState.LOTS.data = result.lotsData || null;
    sheetState.BAIL.data = result.bailData || null;
    sheetState.LOTS.selectedIndex = result.selectedLOTSIndex ?? null;
    sheetState.BAIL.selectedIndex = result.selectedBAILIndex ?? null;

    // Load per-sheet mappings
    sheetMappings.PROP = result.fieldMapping_PROP || {};
    sheetMappings.LOTS = result.fieldMapping_LOTS || {};
    sheetMappings.BAIL = result.fieldMapping_BAIL || {};

    // Restore UI for each sheet
    restoreSheetUI('PROP');
    restoreSheetUI('LOTS');
    restoreSheetUI('BAIL');

    // Custom fields
    if (Array.isArray(result.customFields)) {
      customFieldsArray = result.customFields.length > 0 ? result.customFields : [{ name: '', value: '' }];
    } else if (result.customFields && typeof result.customFields === 'object' && !Array.isArray(result.customFields)) {
      customFieldsArray = Object.entries(result.customFields).map(([name, value]) => ({ name, value: String(value) }));
    } else {
      customFieldsArray = [];
    }
    renderCustomFields();

    updateFillButtonState();
    updateTabIndicators();
    updateClearAllButtons();
    console.log('Data loaded. PROP:', !!sheetState.PROP.data, 'LOTS:', !!sheetState.LOTS.data, 'BAIL:', !!sheetState.BAIL.data);
  } catch (error) {
    console.error('Error loading data:', error);
  }
}

/**
 * Restore the UI state for a sheet from loaded data
 */
function restoreSheetUI(sheetType) {
  const sheet = SHEET_TYPES[sheetType];
  const data = sheetState[sheetType].data;

  if (sheet.multiRow) {
    // LOTS/BAIL: paste area is always visible; just repopulate row list + fields
    const selectorContainer = document.getElementById(`${sheetType}-rowSelector`);
    const fieldsList = document.getElementById(`${sheetType}-fieldsList`);

    if (data && data.length > 0) {
      if (sheetState[sheetType].selectedIndex === null) {
        sheetState[sheetType].selectedIndex = 0;
      }
      if (selectorContainer) displayRowSelector(data, sheetType, selectorContainer);
      const selectedRow = data[sheetState[sheetType].selectedIndex];
      if (fieldsList && selectedRow) displayFields(selectedRow, sheet.columns, fieldsList);
    } else {
      if (selectorContainer) selectorContainer.innerHTML = '';
      if (fieldsList) fieldsList.innerHTML = '';
    }
  } else {
    // PROP: toggle between paste view and data view
    if (!data) {
      switchSheetView(sheetType, 'paste');
      return;
    }
    switchSheetView(sheetType, 'data');
    const fieldsList = document.getElementById(`${sheetType}-fieldsList`);
    if (fieldsList) displayFields(data, sheet.columns, fieldsList);
  }
}

// ============================================================
// INITIALIZATION
// ============================================================

document.addEventListener('DOMContentLoaded', async () => {
  // Resolve DOM references
  statusMessage = document.getElementById('statusMessage');
  fillButton = document.getElementById('fillButton');
  configButton = document.getElementById('configButton');
  configPanel = document.getElementById('configPanel');
  closeConfigButton = document.getElementById('closeConfigButton');
  cancelConfigButton = document.getElementById('cancelConfigButton');
  saveConfigButton = document.getElementById('saveConfigButton');
  mappingFields = document.getElementById('mappingFields');
  customFieldsList = document.getElementById('customFieldsList');
  addCustomFieldButton = document.getElementById('addCustomFieldButton');

  // Sheet tab navigation
  ['PROP', 'LOTS', 'BAIL', 'CUSTOM'].forEach(type => {
    const tab = document.getElementById(`sheetTab-${type}`);
    if (tab) {
      tab.addEventListener('click', () => switchSheetTab(type));
    }
  });

  // PROP: parse + reset buttons
  const propParseBtn = document.getElementById('PROP-parseButton');
  const propResetBtn = document.getElementById('PROP-resetButton');
  const propTextarea = document.getElementById('PROP-textarea');

  if (propParseBtn) {
    propParseBtn.disabled = true;
    propParseBtn.addEventListener('click', () => {
      activeSheetTab = 'PROP';
      handleParse();
    });
  }
  if (propResetBtn) {
    propResetBtn.addEventListener('click', () => {
      activeSheetTab = 'PROP';
      handleReset();
    });
  }
  if (propTextarea) {
    propTextarea.addEventListener('input', () => {
      if (propParseBtn) propParseBtn.disabled = propTextarea.value.trim() === '';
    });
    propTextarea.addEventListener('keydown', (e) => {
      if (e.ctrlKey && e.key === 'Enter') {
        activeSheetTab = 'PROP';
        handleParse();
      }
    });
  }

  // PROP: per-sheet fill button
  const propFillSheetBtn = document.getElementById('PROP-fillSheetButton');
  if (propFillSheetBtn) {
    propFillSheetBtn.addEventListener('click', () => handleFillSheet('PROP'));
  }

  // BAIL: per-sheet fill button
  const bailFillSheetBtn = document.getElementById('BAIL-fillSheetButton');
  if (bailFillSheetBtn) {
    bailFillSheetBtn.addEventListener('click', () => handleFillSheet('BAIL'));
  }

  // LOTS & BAIL: add button + textarea + clear all
  ['LOTS', 'BAIL'].forEach(type => {
    const addBtn = document.getElementById(`${type}-addButton`);
    const textarea = document.getElementById(`${type}-textarea`);
    const clearAllBtn = document.getElementById(`${type}-clearAllButton`);

    if (addBtn) {
      addBtn.disabled = true;
      addBtn.addEventListener('click', () => handleAddRows(type));
    }

    if (clearAllBtn) {
      clearAllBtn.addEventListener('click', () => {
        activeSheetTab = type;
        handleReset();
      });
    }

    if (textarea) {
      textarea.addEventListener('input', () => {
        if (addBtn) addBtn.disabled = textarea.value.trim() === '';
      });
      textarea.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.key === 'Enter') handleAddRows(type);
      });
    }
  });

  // Config and fill buttons
  if (configButton) configButton.addEventListener('click', openConfigPanel);
  if (fillButton) fillButton.addEventListener('click', handleFillForm);
  if (addCustomFieldButton) addCustomFieldButton.addEventListener('click', addCustomField);

  // Config panel buttons
  if (closeConfigButton) closeConfigButton.addEventListener('click', closeConfigPanel);
  if (cancelConfigButton) cancelConfigButton.addEventListener('click', closeConfigPanel);
  if (saveConfigButton) saveConfigButton.addEventListener('click', handleConfigSave);

  // Close config panel on overlay click
  if (configPanel) {
    configPanel.addEventListener('click', (e) => {
      if (e.target === configPanel || e.target.classList.contains('config-overlay')) {
        closeConfigPanel();
      }
    });
  }

  // Load saved data and restore UI
  await loadMapping();
  updateFillButtonState();
});
