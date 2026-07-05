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

    const result = performFill(request.data, request.mapping, request.customFields);

    // Observe le contenu chargé dynamiquement (UpdatePanels ASP.NET, etc.)
    startFillObserver();

    sendResponse(result);
  } else if (request.action === 'copyInputName') {
    if (lastContextMenuTarget) {
      const inputName = lastContextMenuTarget.getAttribute('name');
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
  }
  return true;
});

/**
 * Remplit avec data+mapping, plus les champs personnalisés éventuels
 */
function performFill(data, mapping, customFields) {
  const result = fillFormFields(data, mapping);
  if (customFields && typeof customFields === 'object') {
    const customResult = fillCustomFields(customFields);
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

  fillObserver = new MutationObserver((mutations) => {
    let hasNewInputs = false;
    for (const mutation of mutations) {
      for (const node of mutation.addedNodes) {
        if (node.nodeType === Node.ELEMENT_NODE) {
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
      debounceTimer = setTimeout(() => {
        console.log(`NoHands OSA: nouveaux champs détectés (essai ${retryCount}/${maxRetries}), re-remplissage...`);
        performFill(lastFillData, lastFillMapping, lastFillCustomFields);
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
 * Remplit les champs du formulaire à partir des données et du mapping
 * @param {Object} data - Données de la ligne (nomColonne -> valeur)
 * @param {Object} mapping - nomColonne -> nom(s) d'input
 */
function fillFormFields(data, mapping) {
  let filledCount = 0;
  const errors = [];
  const filled = [];

  for (const [columnName, inputNames] of Object.entries(mapping)) {
    const value = data[columnName];
    if (value === undefined || value === null) continue;

    let inputNamesArray = inputNames;
    if (typeof inputNames === 'string') inputNamesArray = [inputNames];
    if (!Array.isArray(inputNamesArray)) continue;

    inputNamesArray.forEach(inputName => {
      if (!inputName || inputName.trim() === '') return;

      try {
        const input = document.querySelector(`[name="${CSS.escape(inputName)}"]`);
        if (!input) {
          errors.push(`Input non trouvé: ${inputName}`);
          return;
        }
        const success = fillInputByType(input, value);
        if (success) {
          filledCount++;
          filled.push(`${columnName} → ${inputName}`);
        } else {
          errors.push(`Échec pour ${columnName} → ${inputName}`);
        }
      } catch (error) {
        errors.push(`Erreur pour ${columnName} → ${inputName}: ${error.message}`);
      }
    });
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
 * Remplit des champs personnalisés (nom d'input -> valeur fixe)
 */
function fillCustomFields(customFields) {
  let filledCount = 0;
  const errors = [];
  const filled = [];

  for (const [inputName, value] of Object.entries(customFields)) {
    if (!inputName || inputName.trim() === '') continue;
    try {
      const input = document.querySelector(`[name="${CSS.escape(inputName)}"]`);
      if (!input) {
        errors.push(`Input non trouvé: ${inputName}`);
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
      errors.push(`Erreur custom ${inputName}: ${error.message}`);
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
