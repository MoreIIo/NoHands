/**
 * NoHands OSA — Content Script
 * Injecté dans les pages web pour remplir les formulaires automatiquement
 * (mode Saisie). Provient du projet NoHands.
 */

// Garde contre la double injection (manifest + chrome.scripting)
if (window.__nohandsOsaInjected) {
  // Déjà injecté, on ne fait rien.
} else {
  window.__nohandsOsaInjected = true;

// Dernier élément cliqué-droit (pour le menu contextuel)
let lastContextMenuTarget = null;

// Dernière requête de remplissage (pour re-remplissage via MutationObserver)
let lastFillData = null;
let lastFillMapping = null;
let lastFillCustomFields = null;
let lastFillRowContext = null;
let fillObserver = null;

document.addEventListener('contextmenu', (event) => {
  lastContextMenuTarget = event.target;
}, true);

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'fillForm') {
    console.log('NoHands OSA: fillForm reçu', request);

    lastFillData = request.data;
    lastFillMapping = request.mapping;
    lastFillCustomFields = request.customFields || null;
    lastFillRowContext = request.rowContext || null;

    // Stoppe l'observer d'un remplissage précédent pour éviter les
    // re-remplissages concurrents pendant ce remplissage-ci.
    if (fillObserver) {
      fillObserver.disconnect();
      fillObserver = null;
    }

    // Remplissage asynchrone : les champs à autocomplétion attendent les
    // suggestions AJAX avant que la réponse ne parte.
    performFill(request.data, request.mapping, request.customFields, lastFillRowContext)
      .then((result) => {
        // Observe le contenu chargé dynamiquement (UpdatePanels ASP.NET, etc.)
        startFillObserver();
        sendResponse(result);
      })
      .catch((err) => {
        startFillObserver();
        sendResponse({ success: false, filledCount: 0, filled: [], errors: [err.message], error: err.message });
      });
  } else if (request.action === 'copyInputName') {
    if (lastContextMenuTarget) {
      // On privilégie le name ; à défaut on récupère l'id (beaucoup de
      // formulaires n'ont pas d'attribut name sur leurs champs).
      const inputName = lastContextMenuTarget.getAttribute('name') || lastContextMenuTarget.id;
      if (inputName) {
        navigator.clipboard.writeText(inputName).then(() => {
          showCopyNotification(inputName);
          sendResponse({ success: true, inputName: inputName });
        }).catch(err => {
          sendResponse({ success: false, error: err.message });
        });
      } else {
        sendResponse({ success: false, error: 'No name attribute' });
      }
    } else {
      sendResponse({ success: false, error: 'No target element' });
    }
    return true;
  } else if (request.action === 'showRowBadge') {
    showRowBadge(request.label || '', request.state || 'active');
    sendResponse({ success: true });
  } else if (request.action === 'hideRowBadge') {
    const badge = document.getElementById('nohands-osa-row-badge');
    if (badge) badge.remove();
    sendResponse({ success: true });
  }
  return true;
});

/**
 * Badge persistant (mode multi-onglets) : indique quelle ligne du fichier
 * est saisie dans CET onglet. state: 'active' (bleu) ou 'done' (vert).
 * Uniquement dans le cadre principal, pas dans les iframes.
 */
function showRowBadge(label, state) {
  if (window.top !== window) return;
  let badge = document.getElementById('nohands-osa-row-badge');
  if (!badge) {
    badge = document.createElement('div');
    badge.id = 'nohands-osa-row-badge';
    document.documentElement.appendChild(badge);
  }
  badge.textContent = label;
  badge.style.cssText = `
    position: fixed;
    top: 12px;
    right: 12px;
    background: ${state === 'done' ? 'rgba(20, 83, 45, 0.95)' : 'rgba(30, 58, 138, 0.95)'};
    color: ${state === 'done' ? '#4ade80' : '#93c5fd'};
    padding: 8px 16px;
    border-radius: 10px;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    font-size: 15px;
    font-weight: 700;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.35), 0 0 0 1px rgba(255, 255, 255, 0.08);
    z-index: 2147483647;
    pointer-events: none;
  `;
}

/**
 * Remplit avec data+mapping, plus les champs personnalisés éventuels.
 * Asynchrone : les champs à autocomplétion sont attendus séquentiellement.
 */
async function performFill(data, mapping, customFields, rowContext) {
  const result = await fillFormFields(data, mapping, rowContext);
  if (customFields && typeof customFields === 'object') {
    const customResult = await fillCustomFields(customFields, rowContext);
    result.filledCount += customResult.filledCount;
    if (customResult.filled.length) result.filled.push(...customResult.filled);
    if (customResult.errors && customResult.errors.length) {
      result.errors = (result.errors || []).concat(customResult.errors);
      result.error = result.errors.slice(0, 3).join(', ');
    }
    result.success = result.filledCount > 0;
  }
  return result;
}

/**
 * MutationObserver pour le contenu ASP.NET chargé dynamiquement :
 * re-remplit quand de nouveaux champs apparaissent.
 */
function startFillObserver() {
  if (fillObserver) {
    fillObserver.disconnect();
    fillObserver = null;
  }
  if (!lastFillData || !lastFillMapping) return;

  let retryCount = 0;
  const maxRetries = 10;
  let debounceTimer = null;
  let refillRunning = false;

  // Nœuds créés par le mécanisme de suggestions (autocomplétion) : à
  // ignorer, sinon chaque liste de suggestions relancerait un remplissage.
  const isSuggestionNode = (node) =>
    (typeof node.id === 'string' && node.id.startsWith('search:')) ||
    (node.matches && node.matches('select[name^="searchResultSelect_"]'));

  fillObserver = new MutationObserver((mutations) => {
    let hasNewInputs = false;
    for (const mutation of mutations) {
      for (const node of mutation.addedNodes) {
        if (node.nodeType === Node.ELEMENT_NODE) {
          if (isSuggestionNode(node)) continue;
          if ((node.matches && node.matches('input, select, textarea, form')) ||
              (node.querySelector && node.querySelector('input, select, textarea'))) {
            hasNewInputs = true;
            break;
          }
        }
      }
      if (hasNewInputs) break;
    }

    if (hasNewInputs && retryCount < maxRetries) {
      retryCount++;
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(async () => {
        if (refillRunning) return;
        refillRunning = true;
        console.log(`NoHands OSA: nouveaux champs détectés (essai ${retryCount}/${maxRetries}), re-remplissage...`);
        try {
          await performFill(lastFillData, lastFillMapping, lastFillCustomFields, lastFillRowContext);
        } finally {
          refillRunning = false;
        }
      }, 300);
    }
  });

  fillObserver.observe(document.body || document.documentElement, {
    childList: true,
    subtree: true
  });

  // Déconnexion automatique après 30 secondes
  setTimeout(() => {
    if (fillObserver) {
      fillObserver.disconnect();
      fillObserver = null;
    }
  }, 30000);
}

/**
 * Résout un champ de formulaire à partir d'un identifiant fourni par
 * l'utilisateur. Supporte name, id ET classe. Cherche successivement
 * (le premier trouvé gagne) :
 *   1. attribut name exact
 *   2. attribut id exact (getElementById puis [id="…"])
 *   3. sélecteur CSS brut si l'utilisateur en tape un (ex: "#monId",
 *      ".maClasse", "input[data-x=y]")
 *   4. nom de classe brut (ex: "form-control")
 *   5. repli : name OU id se terminant par l'identifiant (utile pour les
 *      ID/name dynamiques type ASP.NET « ctl00$...$txtNom »).
 * Quand plusieurs éléments correspondent (classe), on privilégie le
 * premier champ réellement remplissable (input / select / textarea).
 * @param {string} identifier
 * @returns {Element|null}
 */
function findFormInput(identifier) {
  if (!identifier) return null;
  const key = String(identifier).trim();
  if (!key) return null;

  const FILLABLE = 'input, select, textarea';
  // Parmi une liste de correspondances, renvoie d'abord un champ remplissable.
  const pickBest = (list) => {
    if (!list || !list.length) return null;
    for (const el of list) {
      if (el.matches && el.matches(FILLABLE)) return el;
      // ou un conteneur qui enveloppe un champ remplissable
      const inner = el.querySelector && el.querySelector(FILLABLE);
      if (inner) return inner;
    }
    return list[0];
  };

  const esc = CSS.escape(key);

  // 1. name exact
  let el = document.querySelector(`[name="${esc}"]`);
  if (el) return el;

  // 2. id exact
  el = document.getElementById(key) || document.querySelector(`[id="${esc}"]`);
  if (el) return el;

  // 3. sélecteur CSS brut (l'utilisateur a tapé #id, .classe, [attr]…)
  if (/[#.\[\]>\s,]/.test(key)) {
    try {
      const found = pickBest(document.querySelectorAll(key));
      if (found) return found;
    } catch (_) { /* sélecteur invalide : on ignore */ }
  }

  // 4. nom de classe brut (sans le point)
  try {
    const found = pickBest(document.querySelectorAll(`.${esc}`));
    if (found) return found;
  } catch (_) { /* classe invalide : on ignore */ }

  // 5. repli : name/id se terminant par l'identifiant
  el = document.querySelector(`[name$="${esc}"], [id$="${esc}"]`);
  if (el) return el;

  return null;
}

/* ====================================================================
 * Champs à autocomplétion asynchrone (ASP.NET « searchResult »)
 * --------------------------------------------------------------------
 * Certains formulaires internes (adresses, comptes bancaires…) ont des
 * champs dont la frappe déclenche une recherche AJAX debouncée
 * (searchResult → SearchStart → processTimerAdresse, ~300 ms à 1 s).
 * Les suggestions arrivent dans <div id="search:ID_DU_CHAMP"> contenant
 * un <select name="searchResultSelect_ID_DU_CHAMP"> ; cliquer une
 * <option> appelle setDataFieldValue(...) qui remplit les champs liés
 * (ex. la ville à partir du code postal).
 * Stratégie : taper la valeur, attendre les suggestions par polling
 * (jamais de délai fixe), choisir la meilleure option en s'aidant des
 * autres colonnes de la ligne, puis la cliquer.
 * Détection automatique + marqueur manuel « ac: » dans le mapping.
 * ==================================================================== */

const OSA_AC = {
  POLL_MS: 200,      // intervalle de vérification des suggestions
  TIMEOUT_MS: 8000,  // attente max (debounce + AJAX serveur)
  SETTLE_MS: 400     // délai après clic (setDataFieldValue remplit les champs liés)
};

// Sépare un identifiant de mapping de son éventuel marqueur « ac: »
// (force le traitement autocomplétion si la détection auto échoue).
function parseInputIdentifier(raw) {
  const trimmed = String(raw).trim();
  if (/^ac:/i.test(trimmed)) {
    return { name: trimmed.slice(3).trim(), forceAutocomplete: true };
  }
  return { name: trimmed, forceAutocomplete: false };
}

// Détection automatique d'un champ à autocomplétion :
// handler inline searchResult/SearchStart, ou conteneur « search:<id> ».
function isAutocompleteInput(input) {
  if (!input || input.tagName.toLowerCase() !== 'input') return false;
  const type = (input.type || 'text').toLowerCase();
  if (!['text', 'search'].includes(type)) return false;
  for (const attr of ['onkeyup', 'onkeydown', 'onkeypress']) {
    const code = input.getAttribute(attr) || '';
    if (code.includes('searchResult') || code.includes('SearchStart')) return true;
  }
  if (input.id && document.getElementById('search:' + input.id)) return true;
  return false;
}

function osaSleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// Normalise pour comparaison : majuscules, sans accents, espaces réduits
function osaNormalize(s) {
  return String(s ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

// Le <select> de suggestions associé à un champ
function findSuggestionSelect(input) {
  const id = input.id || '';
  if (!id) return null;
  let sel = document.querySelector(`select[name="searchResultSelect_${CSS.escape(id)}"]`);
  if (!sel) {
    const container = document.getElementById('search:' + id);
    if (container) sel = container.querySelector('select');
  }
  return sel;
}

function nonEmptyOptions(sel) {
  return Array.from(sel.options).filter((o) => (o.text || '').trim() !== '');
}

function isElementVisible(el) {
  return !!el && el.getClientRects().length > 0;
}

// Meilleure option : privilégie celles contenant la valeur tapée, départage
// avec les autres valeurs de la ligne (ex. colonne Ville pour un CP :
// « 70600 - ARGILLIERES » gagne si la ligne contient ARGILLIERES).
function pickBestSuggestion(options, typedValue, rowContext) {
  const typed = osaNormalize(typedValue);
  const contextValues = [];
  if (rowContext && typeof rowContext === 'object') {
    for (const v of Object.values(rowContext)) {
      const n = osaNormalize(v);
      if (n && n.length >= 2 && n !== typed && !contextValues.includes(n)) contextValues.push(n);
    }
  }
  let best = null;
  let bestScore = -1;
  for (const opt of options) {
    const text = osaNormalize(opt.text);
    if (!text) continue;
    let score = 0;
    if (text === typed) score += 4;
    if (typed && text.includes(typed)) score += 2;
    for (const cv of contextValues) {
      if (text.includes(cv)) score += 3;
    }
    if (score > bestScore) { bestScore = score; best = opt; }
  }
  return best;
}

// Clique une option comme le ferait l'utilisateur : sélection puis vrais
// événements souris pour déclencher son handler inline (setDataFieldValue).
function clickSuggestionOption(option) {
  option.selected = true;
  for (const type of ['mousedown', 'mouseup', 'click']) {
    option.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
  }
}

/**
 * Remplit un champ à autocomplétion : tape la valeur, attend les suggestions
 * AJAX (polling toutes les 200 ms, timeout 8 s), sélectionne la meilleure.
 * @returns {Promise<{success: boolean, detail?: string, warning?: string, error?: string}>}
 */
async function fillAutocompleteField(input, value, rowContext) {
  const typedValue = String(value).trim();
  const label = input.id || input.name || '(champ)';

  // Déjà traité avec cette valeur (re-remplissage MutationObserver) :
  // ne pas rouvrir les suggestions.
  if (input.dataset.osaAcDone === typedValue && input.value !== '') {
    return { success: true, detail: 'déjà sélectionné' };
  }

  // Signature des suggestions déjà affichées, pour détecter les nouvelles
  const prevSelect = findSuggestionSelect(input);
  const prevSignature = prevSelect ? nonEmptyOptions(prevSelect).map((o) => o.text).join('|') : '';

  // Frappe simulée : valeur + keyup (déclenche searchResult côté page).
  // Pas de blur volontairement (risque de __doPostBack, cf. triggerInputEvents).
  input.value = typedValue;
  input.dispatchEvent(new Event('input', { bubbles: true }));
  input.dispatchEvent(new KeyboardEvent('keyup', {
    bubbles: true, cancelable: true, key: typedValue.slice(-1) || 'Unidentified'
  }));

  // Attente des suggestions (comportement asynchrone : on ne se fie pas à
  // un délai fixe, on attend que des <option> non vides apparaissent).
  const deadline = Date.now() + OSA_AC.TIMEOUT_MS;
  let options = null;
  while (Date.now() < deadline) {
    await osaSleep(OSA_AC.POLL_MS);
    const sel = findSuggestionSelect(input);
    if (!sel) continue;
    const opts = nonEmptyOptions(sel);
    if (!opts.length) continue;
    const signature = opts.map((o) => o.text).join('|');
    // Nouvelles suggestions, ou liste visible (une ancienne liste cachée ne compte pas)
    if (signature !== prevSignature || isElementVisible(sel)) {
      options = opts;
      break;
    }
  }

  if (!options) {
    // Pas de suggestion : la valeur tapée reste, on signale seulement.
    input.dispatchEvent(new Event('change', { bubbles: true }));
    return {
      success: true,
      warning: `${label} : aucune suggestion pour « ${typedValue} » (délai dépassé) — champ lié non rempli`
    };
  }

  const option = pickBestSuggestion(options, typedValue, rowContext);
  if (!option) {
    return { success: false, error: `${label} : suggestions illisibles pour « ${typedValue} »` };
  }

  clickSuggestionOption(option);
  input.dataset.osaAcDone = typedValue;

  // Laisse setDataFieldValue remplir les champs liés (ex. la ville)
  await osaSleep(OSA_AC.SETTLE_MS);

  return { success: true, detail: option.text.trim() };
}

/**
 * Remplit les champs du formulaire à partir des données et du mapping
 * @param {Object} data - Données de la ligne (nomColonne -> valeur)
 * @param {Object} mapping - nomColonne -> nom(s) d'input
 * @param {Object|null} rowContext - toutes les colonnes de la ligne (désambiguïsation)
 */
async function fillFormFields(data, mapping, rowContext) {
  let filledCount = 0;
  const errors = [];
  const filled = [];

  for (const [columnName, inputNames] of Object.entries(mapping)) {
    const value = data[columnName];
    if (value === undefined || value === null) continue;

    let inputNamesArray = inputNames;
    if (typeof inputNames === 'string') inputNamesArray = [inputNames];
    if (!Array.isArray(inputNamesArray)) continue;

    for (const rawName of inputNamesArray) {
      if (!rawName || rawName.trim() === '') continue;

      try {
        const { name: inputName, forceAutocomplete } = parseInputIdentifier(rawName);
        const input = findFormInput(inputName);
        if (!input) {
          errors.push(`Input non trouvé (name/id/classe): ${inputName}`);
          continue;
        }

        // Champ à autocomplétion asynchrone (détection auto ou marqueur « ac: »)
        if (forceAutocomplete || isAutocompleteInput(input)) {
          const acResult = await fillAutocompleteField(input, value, rowContext);
          if (acResult.success) {
            filledCount++;
            filled.push(`${columnName} → ${inputName}${acResult.detail ? ` (suggestion : ${acResult.detail})` : ''}`);
            if (acResult.warning) errors.push(acResult.warning);
          } else {
            errors.push(acResult.error || `Échec autocomplétion pour ${columnName} → ${inputName}`);
          }
          continue;
        }

        const success = fillInputByType(input, value);
        if (success) {
          filledCount++;
          filled.push(`${columnName} → ${inputName}`);
        } else {
          errors.push(`Échec pour ${columnName} → ${inputName}`);
        }
      } catch (error) {
        errors.push(`Erreur pour ${columnName} → ${rawName}: ${error.message}`);
      }
    }
  }

  return {
    success: filledCount > 0,
    filledCount: filledCount,
    filled: filled,
    errors: errors.length > 0 ? errors : null,
    error: errors.length > 0 ? errors.slice(0, 3).join(', ') : null
  };
}

/**
 * Remplit des champs personnalisés (nom d'input -> valeur fixe).
 * Gère aussi les champs à autocomplétion (auto ou marqueur « ac: »).
 */
async function fillCustomFields(customFields, rowContext) {
  let filledCount = 0;
  const errors = [];
  const filled = [];

  for (const [rawName, value] of Object.entries(customFields)) {
    if (!rawName || rawName.trim() === '') continue;
    try {
      const { name: inputName, forceAutocomplete } = parseInputIdentifier(rawName);
      const input = findFormInput(inputName);
      if (!input) {
        errors.push(`Input non trouvé (name/id/classe): ${inputName}`);
        continue;
      }

      if (forceAutocomplete || isAutocompleteInput(input)) {
        const acResult = await fillAutocompleteField(input, value, rowContext);
        if (acResult.success) {
          filledCount++;
          filled.push(`custom:${inputName}${acResult.detail ? ` (suggestion : ${acResult.detail})` : ''}`);
          if (acResult.warning) errors.push(acResult.warning);
        } else {
          errors.push(acResult.error || `Échec autocomplétion pour custom → ${inputName}`);
        }
        continue;
      }

      const success = fillInputByType(input, value);
      if (success) {
        filledCount++;
        filled.push(`custom:${inputName}`);
      } else {
        errors.push(`Échec pour custom → ${inputName}`);
      }
    } catch (error) {
      errors.push(`Erreur custom ${rawName}: ${error.message}`);
    }
  }

  return { filledCount, filled, errors };
}

/**
 * Remplit un champ selon son type (text, select, checkbox, radio, date, hidden…)
 */
function fillInputByType(input, value) {
  const tagName = input.tagName.toLowerCase();
  const type = input.type ? input.type.toLowerCase() : 'text';

  // Champs texte
  if (tagName === 'input' && ['text', 'email', 'tel', 'number', 'url', 'search', 'password'].includes(type)) {
    // Détection des champs IBAN / RIB découpés : si la valeur dépasse maxLength
    // et que le nom finit par un numéro, on répartit sur les champs numérotés.
    const stripped = String(value).replace(/\s/g, '');
    if (input.maxLength > 0 && stripped.length > input.maxLength) {
      const splitCount = trySplitAcrossNumberedInputs(input, stripped);
      if (splitCount > 0) return true;
    }

    input.value = value;
    triggerInputEvents(input);
    return true;
  }

  // Textarea
  if (tagName === 'textarea') {
    input.value = value;
    triggerInputEvents(input);
    return true;
  }

  // Select : plusieurs stratégies de correspondance
  if (tagName === 'select') {
    const valueLower = value.toString().toLowerCase().trim();
    let option = null;

    option = Array.from(input.options).find(opt => opt.value === value);
    if (!option) option = Array.from(input.options).find(opt => opt.text === value);
    if (!option) option = Array.from(input.options).find(opt => opt.value.toLowerCase().trim() === valueLower);
    if (!option) option = Array.from(input.options).find(opt => opt.text.toLowerCase().trim() === valueLower);
    if (!option) {
      option = Array.from(input.options).find(opt =>
        opt.text.toLowerCase().includes(valueLower) || valueLower.includes(opt.text.toLowerCase())
      );
    }
    if (!option) {
      option = Array.from(input.options).find(opt =>
        opt.value.toLowerCase().includes(valueLower) || valueLower.includes(opt.value.toLowerCase())
      );
    }

    if (option) {
      input.value = option.value;
      triggerChangeEvent(input);
      return true;
    }
    console.warn(`NoHands OSA: aucune option correspondant à "${value}" dans le select`, input.name);
    return false;
  }

  // Checkbox : O/OUI/YES/TRUE/1/ON => coché
  if (tagName === 'input' && type === 'checkbox') {
    const shouldCheck = ['o', 'oui', 'yes', 'true', '1', 'on', 'checked'].includes(
      value.toString().toLowerCase()
    );
    input.checked = shouldCheck;
    triggerChangeEvent(input);
    return true;
  }

  // Radio
  if (tagName === 'input' && type === 'radio') {
    const radios = document.querySelectorAll(`[name="${CSS.escape(input.name)}"]`);
    const matchingRadio = Array.from(radios).find(r =>
      r.value === value || r.value.toLowerCase() === String(value).toLowerCase()
    );
    if (matchingRadio) {
      matchingRadio.checked = true;
      triggerChangeEvent(matchingRadio);
      return true;
    }
    return false;
  }

  // Date : conversion DD/MM/YYYY -> YYYY-MM-DD
  if (tagName === 'input' && type === 'date') {
    const convertedDate = convertDateFormat(String(value));
    if (convertedDate) {
      input.value = convertedDate;
      triggerChangeEvent(input);
      return true;
    }
    return false;
  }

  // Hidden
  if (tagName === 'input' && type === 'hidden') {
    input.value = value;
    triggerChangeEvent(input);
    return true;
  }

  console.warn(`NoHands OSA: type de champ non géré: ${tagName} (${type})`);
  return false;
}

/**
 * Répartit une valeur longue (ex: IBAN) sur des champs numérotés successifs.
 * Gère les noms ASP.NET où le numéro apparaît plusieurs fois
 * (txtIBAN1:x:txtIBAN1 -> txtIBAN2:x:txtIBAN2).
 * @returns {number} nombre de champs remplis (0 si motif non détecté)
 */
function trySplitAcrossNumberedInputs(firstInput, stripped) {
  const name = firstInput.getAttribute('name') || '';
  const endMatch = name.match(/(\d+)$/);
  if (!endMatch) return 0;

  const startNum = parseInt(endMatch[1], 10);
  const numStr = endMatch[1];

  const prefixLetters = name.slice(0, endMatch.index).match(/([A-Za-z_]+)$/);
  let makeName;

  if (prefixLetters) {
    const token = prefixLetters[1] + numStr;
    makeName = (n) => {
      const newToken = prefixLetters[1] + n;
      return name.split(token).join(newToken);
    };
  } else {
    const beforeNum = name.slice(0, endMatch.index);
    makeName = (n) => beforeNum + n;
  }

  const inputs = [];
  for (let n = startNum; ; n++) {
    const candidateName = makeName(n);
    const el = document.querySelector(`[name="${CSS.escape(candidateName)}"]`);
    if (!el) break;
    inputs.push(el);
  }

  if (inputs.length < 2) return 0;

  let offset = 0;
  let filledCount = 0;
  for (const inp of inputs) {
    if (offset >= stripped.length) break;
    const len = inp.maxLength > 0 ? inp.maxLength : (stripped.length - offset);
    const chunk = stripped.slice(offset, offset + len);
    inp.value = chunk;
    triggerInputEvents(inp);
    offset += len;
    filledCount++;
  }

  return filledCount;
}

/**
 * Déclenche input+change pour que les frameworks détectent la modification.
 * NOTE : pas de 'blur' volontairement — sur les pages ASP.NET WebForms, le blur
 * de certains champs déclenche __doPostBack et ouvre des popups indésirables.
 */
function triggerInputEvents(element) {
  element.dispatchEvent(new Event('input', { bubbles: true }));
  element.dispatchEvent(new Event('change', { bubbles: true }));
}

function triggerChangeEvent(element) {
  element.dispatchEvent(new Event('change', { bubbles: true }));
}

/**
 * Convertit une date DD/MM/YYYY (ou D/M/YYYY) en YYYY-MM-DD
 */
function convertDateFormat(dateStr) {
  const match1 = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (match1) {
    const [, day, month, year] = match1;
    return `${year}-${month}-${day}`;
  }
  const match2 = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match2) {
    const [, day, month, year] = match2;
    return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
  }
  if (/^(\d{4})-(\d{2})-(\d{2})$/.test(dateStr)) return dateStr;
  return dateStr;
}

/**
 * Notification visuelle temporaire (copie du nom d'input)
 */
function showCopyNotification(inputName) {
  const notification = document.createElement('div');
  notification.textContent = `✓ Copié: ${inputName}`;
  notification.style.cssText = `
    position: fixed;
    top: 20px;
    right: 20px;
    background: rgba(24, 24, 27, 0.92);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    color: #4ade80;
    padding: 12px 20px;
    border-radius: 12px;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    font-size: 14px;
    font-weight: 600;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3), 0 0 0 1px rgba(255, 255, 255, 0.06);
    z-index: 999999;
  `;
  document.body.appendChild(notification);
  setTimeout(() => notification.remove(), 2000);
}

console.log('NoHands OSA: content script chargé');

} // Fin de la garde d'injection
