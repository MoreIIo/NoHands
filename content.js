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

// Sépare un identifiant de mapping de ses éventuels marqueurs (cumulables) :
//   « ac:nom » force le traitement autocomplétion (si la détection auto échoue) ;
//   « pb:nom » force un postback après remplissage : blur ciblé + attente du
//   rechargement partiel ASP.NET (champs « en cascade », ex. code postal, voie).
function parseInputIdentifier(raw) {
  let name = String(raw).trim();
  let forceAutocomplete = false;
  let forcePostback = false;
  let m;
  while ((m = name.match(/^(ac|pb):/i)) !== null) {
    if (m[1].toLowerCase() === 'ac') forceAutocomplete = true;
    else forcePostback = true;
    name = name.slice(m[0].length).trim();
  }
  return { name, forceAutocomplete, forcePostback };
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

  return { success: true, detail: 'suggestion : ' + option.text.trim() };
}

/* ====================================================================
 * Formulaires « en cascade » (ASP.NET WebForms / __doPostBack)
 * --------------------------------------------------------------------
 * Sur ces pages, certains champs reconstruisent la suite du formulaire
 * côté serveur : un select Pays déclenche __doPostBack au changement, un
 * code postal ou une voie déclenchent un appel serveur au blur
 * (initCacheresultCall…), et des listes comme « Numéro » restent vides
 * tant que le serveur ne les a pas peuplées.
 * Réponses apportées ici :
 *   1. idempotence : une valeur déjà en place n'est jamais re-remplie
 *      (sinon chaque re-remplissage relancerait les postbacks en boucle) ;
 *   2. selects auto-postback (__doPostBack dans onchange) : après le
 *      changement, on attend la fin du rechargement partiel (quiescence
 *      DOM) avant le champ suivant ;
 *   3. selects en cascade vides : on attend que les options apparaissent ;
 *   4. marqueur « pb: » : après remplissage, déclenche blur+focusout puis
 *      attend le rechargement (blur ciblé uniquement — un blur global
 *      ouvre des popups sur certaines pages WebForms).
 * L'ordre de remplissage suit l'ordre des colonnes du mapping : placer
 * Pays avant Code postal, Code postal avant Voie, etc.
 * ==================================================================== */

const OSA_PB = {
  QUIET_MS: 800,           // durée sans mutation DOM = page « posée »
  TIMEOUT_MS: 10000,       // attente max d'un rechargement partiel
  SELECT_RETRY_MS: 300,    // intervalle de re-vérification d'un select vide
  SELECT_TIMEOUT_MS: 8000  // attente max des options d'un select en cascade
};

// Attend que le DOM se stabilise (aucune mutation de structure pendant
// QUIET_MS). Sert à laisser un __doPostBack / UpdatePanel se terminer.
// Ne regarde que childList (les animations d'attributs ne comptent pas).
function waitForDomSettle(quietMs = OSA_PB.QUIET_MS, timeoutMs = OSA_PB.TIMEOUT_MS) {
  return new Promise((resolve) => {
    let quietTimer = null;
    let hardTimer = null;
    let done = false;
    const obs = new MutationObserver(() => arm());
    const finish = () => {
      if (done) return;
      done = true;
      obs.disconnect();
      clearTimeout(quietTimer);
      clearTimeout(hardTimer);
      resolve();
    };
    const arm = () => {
      clearTimeout(quietTimer);
      quietTimer = setTimeout(finish, quietMs);
    };
    obs.observe(document.documentElement, { childList: true, subtree: true });
    hardTimer = setTimeout(finish, timeoutMs);
    arm(); // si aucun postback ne part, on ressort après quietMs
  });
}

// Le changement de ce select déclenche-t-il un postback ASP.NET ?
function selectTriggersPostback(el) {
  if (!el || !el.getAttribute) return false;
  const code = el.getAttribute('onchange') || '';
  return code.includes('__doPostBack') || code.includes('WebForm_DoPostBack');
}

// Sortie de champ ciblée : déclenche les handlers inline onblur
// (initCacheresultCall…) sans toucher aux autres champs.
function triggerBlurEvents(element) {
  element.dispatchEvent(new FocusEvent('blur'));
  element.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
}

// Cherche l'option d'un select correspondant à une valeur (stratégies
// successives : valeur exacte, texte exact, insensible à la casse, puis
// inclusion — en ignorant les options vides, qui matchaient tout avant).
function findSelectOption(select, value) {
  const strVal = String(value);
  const valueLower = strVal.toLowerCase().trim();
  const opts = Array.from(select.options);
  let option = opts.find(opt => opt.value === strVal);
  if (!option) option = opts.find(opt => opt.text === strVal);
  if (!option) option = opts.find(opt => opt.value.toLowerCase().trim() === valueLower);
  if (!option) option = opts.find(opt => opt.text.toLowerCase().trim() === valueLower);
  if (!option) {
    option = opts.find(opt => {
      const t = opt.text.toLowerCase().trim();
      return t !== '' && (t.includes(valueLower) || valueLower.includes(t));
    });
  }
  if (!option) {
    option = opts.find(opt => {
      const v = opt.value.toLowerCase().trim();
      return v !== '' && (v.includes(valueLower) || valueLower.includes(v));
    });
  }
  return option || null;
}

// La valeur est-elle déjà en place ? (idempotence : évite de re-déclencher
// change/postback à chaque re-remplissage de l'observer)
function isAlreadyFilled(input, value) {
  const tagName = input.tagName.toLowerCase();
  const type = input.type ? input.type.toLowerCase() : 'text';
  const strVal = String(value);
  if (tagName === 'select') {
    const option = findSelectOption(input, value);
    return !!option && option.value !== '' && input.value === option.value;
  }
  if (tagName === 'input' && type === 'checkbox') {
    const shouldCheck = ['o', 'oui', 'yes', 'true', '1', 'on', 'checked'].includes(strVal.toLowerCase());
    return input.checked === shouldCheck;
  }
  if (tagName === 'input' && type === 'radio') {
    const radios = document.querySelectorAll(`[name="${CSS.escape(input.name)}"]`);
    const target = Array.from(radios).find(r =>
      r.value === strVal || r.value.toLowerCase() === strVal.toLowerCase()
    );
    return !!target && target.checked;
  }
  if (tagName === 'input' && type === 'date') {
    return strVal !== '' && input.value === convertDateFormat(strVal);
  }
  if (tagName === 'textarea' || tagName === 'input') {
    return strVal !== '' && input.value === strVal;
  }
  return false;
}

// Remplit un select en attendant au besoin que ses options apparaissent
// (listes en cascade peuplées par le postback d'un champ précédent).
// Re-résout l'élément à chaque essai : le postback a pu le remplacer.
async function fillSelectWaiting(identifier, value) {
  const deadline = Date.now() + OSA_PB.SELECT_TIMEOUT_MS;
  let waited = false;
  while (true) {
    const el = findFormInput(identifier);
    if (el && el.tagName.toLowerCase() === 'select') {
      const option = findSelectOption(el, value);
      if (option) {
        if (el.value === option.value) {
          return { success: true, element: el, changed: false, detail: 'déjà en place' };
        }
        el.value = option.value;
        triggerChangeEvent(el);
        return {
          success: true, element: el, changed: true,
          detail: waited ? 'option apparue après rechargement' : null
        };
      }
      // Des options réelles existent mais aucune ne correspond : échec
      // immédiat (comportement historique). On n'attend que si la liste
      // est vide (select en cascade pas encore peuplé).
      const realOptions = Array.from(el.options).filter(o => (o.value || o.text.trim()) !== '');
      if (realOptions.length > 0) {
        return { success: false, error: `aucune option correspondant à « ${value} » dans ${identifier}` };
      }
    }
    if (Date.now() >= deadline) {
      return {
        success: false,
        error: `select ${identifier} : liste restée vide, option « ${value} » jamais apparue — cascade non déclenchée ? (vérifie l'ordre des colonnes, ou ajoute pb: au champ précédent)`
      };
    }
    waited = true;
    await osaSleep(OSA_PB.SELECT_RETRY_MS);
  }
}

/**
 * Remplit UN champ (logique commune mapping + champs personnalisés) :
 * marqueurs ac:/pb:, autocomplétion, idempotence, selects en cascade,
 * attente des postbacks.
 * @returns {Promise<{success: boolean, identifier?: string, detail?: string, warning?: string, error?: string}>}
 */
async function fillOneField(rawIdentifier, value, rowContext) {
  const { name: identifier, forceAutocomplete, forcePostback } = parseInputIdentifier(rawIdentifier);
  const input = findFormInput(identifier);
  if (!input) {
    return { success: false, identifier, error: `Input non trouvé (name/id/classe): ${identifier}` };
  }

  // 1. Champ à autocomplétion asynchrone (suggestions AJAX)
  if (forceAutocomplete || isAutocompleteInput(input)) {
    const acResult = await fillAutocompleteField(input, value, rowContext);
    acResult.identifier = identifier;
    if (acResult.success && forcePostback) {
      triggerBlurEvents(findFormInput(identifier) || input);
      await waitForDomSettle();
      acResult.detail = (acResult.detail ? acResult.detail + ' — ' : '') + 'rechargement attendu';
    }
    return acResult;
  }

  // 2. Valeur déjà en place : ne rien re-déclencher
  if (isAlreadyFilled(input, value)) {
    return { success: true, identifier, detail: 'déjà en place' };
  }

  // 3. Select : attente des options (cascade) + postback éventuel
  if (input.tagName.toLowerCase() === 'select') {
    const res = await fillSelectWaiting(identifier, value);
    res.identifier = identifier;
    if (!res.success) return res;
    const el = res.element || findFormInput(identifier);
    if (res.changed && (forcePostback || selectTriggersPostback(el))) {
      await waitForDomSettle();
      res.detail = (res.detail ? res.detail + ' — ' : '') + 'rechargement attendu';
    }
    return res;
  }

  // 4. Champs classiques
  const success = fillInputByType(input, value);
  if (!success) {
    return { success: false, identifier, error: `Échec pour ${identifier}` };
  }

  // 5. Marqueur pb: : sortie de champ + attente du rechargement
  //    (ex. code postal initCacheresultCall, voie onblur)
  if (forcePostback) {
    triggerBlurEvents(input);
    await waitForDomSettle();
    return { success: true, identifier, detail: 'blur + rechargement attendu' };
  }
  return { success: true, identifier };
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
        const res = await fillOneField(rawName, value, rowContext);
        const shownName = res.identifier || rawName;
        if (res.success) {
          filledCount++;
          filled.push(`${columnName} → ${shownName}${res.detail ? ` (${res.detail})` : ''}`);
          if (res.warning) errors.push(res.warning);
        } else {
          errors.push(res.error || `Échec pour ${columnName} → ${shownName}`);
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
      const res = await fillOneField(rawName, value, rowContext);
      const shownName = res.identifier || rawName;
      if (res.success) {
        filledCount++;
        filled.push(`custom:${shownName}${res.detail ? ` (${res.detail})` : ''}`);
        if (res.warning) errors.push(res.warning);
      } else {
        errors.push(res.error || `Échec pour custom → ${shownName}`);
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

  // Select : correspondance via findSelectOption (stratégies successives)
  if (tagName === 'select') {
    const option = findSelectOption(input, value);

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
