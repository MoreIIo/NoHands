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
  } else if (request.action === 'batchSelectSlice') {
    // Sélection d'une tranche de cases (étape de scénario « éditer par lots »).
    runSelectSlice(request.config || {})
      .then((report) => sendResponse({ success: true, ...report }))
      .catch((err) => sendResponse({ success: false, error: err.message }));
    return true; // réponse asynchrone
  } else if (request.action === 'scanDataTable') {
    // Détection + extraction d'une table de données SIGEO (liste de cases
    // à cocher) : attend la table (rechargement asynchrone possible), puis
    // renvoie les lignes { checked, label, hiddenValue } + le diagnostic.
    scanSigeoDataTable(request.config || {})
      .then((report) => sendResponse({ success: report.ok, ...report }))
      .catch((err) => sendResponse({ success: false, error: err.message }));
    return true; // réponse asynchrone
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

/* ====================================================================
 * Sélection par lots dans un tableau de cases à cocher
 * --------------------------------------------------------------------
 * Certaines pages internes affichent une table de cases à cocher où
 * chaque case id="X" possède un champ hidden miroir id="hdnX" ; le
 * onclick natif de la case fait hdnX.value = checked ? '1' : '0'.
 * C'est ce hidden qui est réellement posté au serveur : régler
 * input.checked = true ne suffit donc PAS, il faut passer par
 * input.click() (qui exécute le onclick) ou synchroniser le miroir
 * à la main.
 * Sert à l'étape de scénario « éditer par lots » : la page d'édition
 * de feuille de présence n'accepte qu'un nombre limité de clés à la
 * fois, on coche donc N clés, on clique le bouton d'édition (qui
 * télécharge un .doc), puis on recommence avec les N suivantes.
 * Les éléments sont re-résolus par id à chaque opération : un postback
 * partiel remplace les noeuds du DOM, et les références gardées
 * deviendraient des orphelins silencieux.
 * Cible de référence : syn_man_edition_feuille_presence / tabCleRepart.
 * ==================================================================== */

// Résout le périmètre de recherche : id de table, sélecteur CSS, name
// d'un champ, ou rien du tout (= toute la page).
function resolveBatchScope(rawScope) {
  const key = String(rawScope ?? '').trim();
  if (!key) return document;

  const byId = document.getElementById(key);
  if (byId) return byId;

  if (/[#.\[\]>\s,:]/.test(key)) {
    try {
      const found = document.querySelector(key);
      if (found) return found;
    } catch (_) { /* sélecteur invalide : on continue */ }
  }

  // Repli : périmètre désigné par le name d'un champ qu'il contient.
  try {
    const input = findFormInput(key);
    if (input) return input.closest('table') || input.parentElement || input;
  } catch (_) { /* recherche infructueuse : on continue */ }

  try {
    const esc = CSS.escape(key);
    const tail = document.querySelector(`[id$="${esc}"], [name$="${esc}"]`);
    if (tail) return tail;
  } catch (_) { /* ignoré */ }

  return null;
}

// Détecte l'écran de blocage « Chargement en cours… » (UpdateProgress
// ASP.NET) affiché pendant un postback. Renvoie le texte de l'overlay
// visible, ou '' s'il n'y en a pas.
//   deep = false : ne teste que les conteneurs classiques (rapide, appelé
//                  à chaque tranche) ;
//   deep = true  : balaie aussi div/span/p/td (plus lent, réservé au
//                  diagnostic quand la sélection a échoué).
function detectLoadingOverlay(deep) {
  const re = /chargement\s+en\s+cours/i;
  const known = document.querySelectorAll(
    '[id*="Progress" i],[id*="progress"],.modalPopup,.blockUI,.ajax__updateprogress,[class*="chargement" i],[class*="loading" i]'
  );
  for (const el of known) {
    if (isElementVisible(el) && re.test(el.textContent || '')) {
      return String(el.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 80);
    }
  }
  if (!deep) return '';
  // Repli : le message peut vivre dans un conteneur quelconque. On ne garde
  // que les noeuds proches du texte (peu d'enfants) pour éviter de remonter
  // sur <body>.
  const tags = document.querySelectorAll('div,span,p,td,th');
  for (const el of tags) {
    if (el.children.length > 3) continue;
    if (re.test(el.textContent || '') && isElementVisible(el)) {
      return String(el.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 80);
    }
  }
  return '';
}

// Diagnostic du périmètre quand la tranche n'a pas pu être sélectionnée :
// dit si l'élément de périmètre existe, s'il est visible, combien de cases
// il contient, et combien il y en a sur la page entière.
function batchScopeDiag(rawScope) {
  const key = String(rawScope ?? '').trim();
  const d = {
    scope: key || '(toute la page)',
    scopeInDom: false,
    scopeVisible: false,
    checkboxesInScope: 0,
    checkboxesOnPage: document.querySelectorAll('input[type="checkbox"]').length,
    url: location.href,
    title: document.title
  };
  if (!key) {
    d.scopeInDom = true;
    d.scopeVisible = true;
    d.checkboxesInScope = d.checkboxesOnPage;
    return d;
  }
  let el = document.getElementById(key);
  if (!el && /[#.\[\]>\s,:]/.test(key)) {
    try { el = document.querySelector(key); } catch (_) { /* sélecteur invalide */ }
  }
  if (!el) {
    try {
      const esc = CSS.escape(key);
      el = document.querySelector(`[id$="${esc}"],[name$="${esc}"]`);
    } catch (_) { /* ignoré */ }
  }
  if (el) {
    d.scopeInDom = true;
    d.scopeVisible = isElementVisible(el);
    d.checkboxesInScope = el.querySelectorAll
      ? el.querySelectorAll('input[type="checkbox"]').length : 0;
  }
  return d;
}

// Le hidden miroir d'une case (hdn + id) ne doit jamais être traité
// comme une cible à part entière : il est piloté par sa case.
function isBatchMirrorHidden(el) {
  if (!el || !el.id || !/^hdn/i.test(el.id)) return false;
  return !!document.getElementById(el.id.replace(/^hdn/i, ''));
}

// Libellé associé à un élément, pour le filtre texte : <label for>,
// sinon le texte de la cellule, sinon celui de la ligne du tableau.
function batchTargetLabel(el) {
  let txt = '';
  if (el.id) {
    try {
      const lab = document.querySelector(`label[for="${CSS.escape(el.id)}"]`);
      if (lab) txt = lab.textContent;
    } catch (_) { /* ignoré */ }
  }
  if (!txt.trim()) {
    const td = el.closest('td');
    if (td) txt = td.textContent;
  }
  if (!txt.trim()) {
    const tr = el.closest('tr');
    if (tr) txt = tr.textContent;
  }
  return String(txt || '').replace(/\s+/g, ' ').trim();
}

// Filtre optionnel : « /motif/i » = expression régulière, sinon
// recherche de texte insensible à la casse et aux accents.
function makeBatchMatcher(filter) {
  const raw = String(filter ?? '').trim();
  if (!raw) return null;
  const m = raw.match(/^\/(.*)\/([gimsuy]*)$/);
  if (m) {
    try {
      const re = new RegExp(m[1], m[2].replace(/g/g, ''));
      return (txt) => re.test(txt);
    } catch (_) { /* regex invalide : repli sur la recherche texte */ }
  }
  const needle = osaNormalize(raw);
  return (txt) => osaNormalize(txt).includes(needle);
}

/**
 * Liste ordonnée des éléments à traiter dans le périmètre demandé.
 * Renvoie { scopeFound, targets }.
 */
function collectBatchTargets(config) {
  const scope = resolveBatchScope(config.scopeSelector);
  if (!scope) return { scopeFound: false, targets: [] };

  const type = String(config.elementType || 'auto').toLowerCase();
  const q = (sel) => Array.from(scope.querySelectorAll(sel));

  const checkboxes = () => {
    const inCells = q('td input[type="checkbox"]');
    return inCells.length ? inCells : q('input[type="checkbox"]');
  };
  const textInputs = () =>
    q('td input[type="text"], td input[type="hidden"], td textarea')
      .concat(q('input[type="text"], input[type="hidden"], textarea'))
      .filter((el, i, arr) => arr.indexOf(el) === i);

  let nodes;
  if (type === 'checkbox') nodes = checkboxes();
  else if (type === 'input') nodes = textInputs();
  else {
    nodes = checkboxes();
    if (!nodes.length) nodes = textInputs();
  }

  const matcher = makeBatchMatcher(config.matchFilter);

  return {
    scopeFound: true,
    targets: nodes.filter((el) => {
      if (el.disabled || el.readOnly) return false;
      if (isBatchMirrorHidden(el)) return false;
      // Les champs hidden n'ont pas de boîte : on ne teste la visibilité
      // que sur les éléments censés être affichés.
      const t = (el.type || '').toLowerCase();
      if (t !== 'hidden' && !isElementVisible(el)) return false;
      if (matcher && !matcher(batchTargetLabel(el))) return false;
      return true;
    })
  };
}

// Interprète l'état voulu pour une case à cocher (le panneau peut
// envoyer un booléen, ou une chaîne quand le mode « Valeur… » est actif).
function toCheckedState(desired) {
  if (typeof desired === 'boolean') return desired;
  const s = osaNormalize(desired);
  return !['', '0', 'FALSE', 'NON', 'NO', 'N', 'DECOCHER'].includes(s);
}

// Force le hidden miroir hdn+id, au cas où le onclick natif serait
// absent ou n'aurait pas fait son travail. Sans effet s'il est déjà bon.
function syncMirrorHidden(checkbox, checked) {
  if (!checkbox.id) return;
  const mirror = document.getElementById('hdn' + checkbox.id);
  if (!mirror) return;
  const want = checked ? '1' : '0';
  if (mirror.value !== want) {
    mirror.value = want;
    triggerChangeEvent(mirror);
  }
}

/**
 * Applique l'état voulu à UN élément.
 * Renvoie 'already' (déjà dans l'état voulu, rien fait) ou 'done'.
 * Idempotent : on ne clique jamais une case déjà dans le bon état,
 * sinon on l'inverserait.
 */
function applyToTarget(el, desiredState) {
  const type = (el.type || '').toLowerCase();

  if (type === 'checkbox' || type === 'radio') {
    const want = toCheckedState(desiredState);
    if (el.checked === want) {
      syncMirrorHidden(el, want); // le hidden peut être désynchronisé
      return 'already';
    }
    el.click();                   // exécute le onclick natif → met à jour hdnX
    if (el.checked !== want) {    // repli si un handler a annulé le clic
      el.checked = want;
      triggerInputEvents(el);
    }
    syncMirrorHidden(el, want);   // filet de sécurité (cases sans onclick)
    return 'done';
  }

  const value = desiredState === true ? '1'
    : desiredState === false ? '0'
    : String(desiredState ?? '');
  if (el.value === value) return 'already';
  el.value = value;
  triggerInputEvents(el);
  return 'done';
}

// Un postback partiel remplace les noeuds : on retient une clé stable
// plutôt que la référence DOM, et on re-résout à chaque lot.
function batchKeyOf(el) {
  if (el.id) return { id: el.id };
  const name = el.getAttribute && el.getAttribute('name');
  if (name) return { name };
  return { el };
}

function resolveBatchKey(key) {
  if (key.el) return document.contains(key.el) ? key.el : null;
  if (key.id) return document.getElementById(key.id);
  if (key.name) {
    try { return document.querySelector(`[name="${CSS.escape(key.name)}"]`); }
    catch (_) { return null; }
  }
  return null;
}

/**
 * Sélectionne UNE tranche de cases et désélectionne tout le reste.
 * config = { scopeSelector, matchFilter, offset, count, uncheckOthers }
 *
 * L'ordre du document sert de repère : la tranche est
 * targets[offset … offset+count[. Le décompte est donc stable d'un
 * appel à l'autre, même si la page se recharge partiellement entre
 * deux lots (les cases reviennent dans le même ordre).
 *
 * Renvoie { scopeFound, total, offset, selected, unchecked, from, to,
 *           remaining, labels }.
 */
async function runSelectSlice(config) {
  // La zone de données peut être rechargée en asynchrone (bouton
  // « Réactualiser la liste des clés de répartition ») : on ATTEND que la
  // table cible existe et soit peuplée avant de compter quoi que ce soit,
  // sinon la tranche serait calculée sur un DOM vide et l'étape
  // « passerait » silencieusement. Le scopeSelector sert d'indice de cible
  // s'il ressemble à un id simple (pas un sélecteur CSS composé).
  const waited = await waitForSigeoTable({
    tableId: /^[\w-]+$/.test(String(config.scopeSelector || '').trim())
      ? String(config.scopeSelector).trim() : '',
    minRows: config.minRows,
    timeoutMs: config.waitTimeoutMs,
    pollMs: config.pollMs
  });

  let { scopeFound, targets } = collectBatchTargets({
    scopeSelector: config.scopeSelector,
    matchFilter: config.matchFilter,
    elementType: 'checkbox'
  });

  // Périmètre explicite introuvable (ou sans case) mais table de données
  // auto-détectée par son id body_x_… : on se replie dessus. Couvre le cas
  // où le nom de page change (PfeuillePresenceChoixCle → autre écran) alors
  // que le scénario contient encore l'ancien id complet.
  let usedScope = String(config.scopeSelector || '').trim() || '(toute la page)';
  if ((!scopeFound || !targets.length) && waited.table) {
    const fallback = collectBatchTargets({
      scopeSelector: waited.table.id,
      matchFilter: config.matchFilter,
      elementType: 'checkbox'
    });
    if (fallback.scopeFound && fallback.targets.length) {
      scopeFound = fallback.scopeFound;
      targets = fallback.targets;
      usedScope = '#' + waited.table.id + ' (auto-détectée)';
      console.log(OSA_SCAN_TAG + ' périmètre « ' + (config.scopeSelector || '(vide)') +
        ' » sans case exploitable — repli sur #' + waited.table.id);
    }
  }

  const total = targets.length;
  const offset = Math.max(0, parseInt(config.offset, 10) || 0);
  const count = Math.max(1, parseInt(config.count, 10) || 10);

  const report = {
    scopeFound,
    total,
    offset,
    selected: 0,
    unchecked: 0,
    from: 0,
    to: 0,
    remaining: Math.max(0, total - offset),
    labels: []
  };

  // Détection de l'overlay de chargement : recherche rapide à chaque appel,
  // recherche approfondie seulement si la sélection a échoué (diagnostic).
  report.overlay = detectLoadingOverlay(!scopeFound || !total) || waited.overlay || '';
  if (!scopeFound) report.diag = batchScopeDiag(config.scopeSelector);

  // Diagnostic de détection : combien de tables candidates, laquelle est
  // retenue, sur quel périmètre on a compté, et combien de temps on a
  // attendu le rechargement. Visible dans la console de la page.
  report.scan = {
    candidates: waited.diag ? waited.diag.candidates : 0,
    pickedId: waited.diag ? waited.diag.pickedId : null,
    rejected: waited.diag ? waited.diag.rejected : [],
    scope: usedScope,
    waitedMs: waited.waitedMs,
    timedOut: waited.timedOut
  };
  console.log(OSA_SCAN_TAG + ' lot : ' + report.scan.candidates + ' table(s) candidate(s), retenue : ' +
    (report.scan.pickedId ? '#' + report.scan.pickedId : 'aucune') + ', périmètre : ' + usedScope +
    ', cases : ' + total + ', attente : ' + waited.waitedMs + ' ms' +
    (waited.timedOut ? ' (TIMEOUT)' : ''));

  if (!scopeFound || !total || offset >= total) return report;

  // On travaille sur des clés stables plutôt que sur les références :
  // si un handler de la page reconstruit le tableau pendant qu'on
  // coche, les noeuds d'origine seraient détachés silencieusement.
  const keys = targets.map(batchKeyOf);
  const wanted = new Set();
  for (let i = offset; i < Math.min(offset + count, total); i++) wanted.add(i);

  for (let i = 0; i < keys.length; i++) {
    const want = wanted.has(i);
    if (!want && config.uncheckOthers === false) continue;

    const el = resolveBatchKey(keys[i]);
    if (!el) continue;

    const changed = applyToTarget(el, want) === 'done';
    if (want) {
      report.selected++;
      report.labels.push(batchTargetLabel(el).slice(0, 70));
    } else if (changed) {
      report.unchecked++;
    }
  }

  report.from = offset + 1;
  report.to = offset + report.selected;
  report.remaining = Math.max(0, total - report.to);
  return report;
}

/* ====================================================================
 * Détection fiable des tables de données SIGEO / Evoriel
 * --------------------------------------------------------------------
 * Ces pages (ASP.NET WebForms « maison », sans jQuery/React) contiennent
 * ~66 <table> dont la quasi-totalité sont des tables de MISE EN PAGE :
 * menus masqués (content_ivmenu00_menu_*_table, largeur/hauteur nulles)
 * et conteneurs à cellule unique. Aucune n'a de <thead>, de <th>, ni de
 * role="grid" : une détection générique « table avec en-têtes » ne
 * trouve rien — ou retient une table de layout — et l'étape « passe ».
 *
 * La seule signature fiable est l'id serveur, stable et normé :
 *   body_x_<NomPage>_x_tab<Nom>  (ex : body_x_PfeuillePresenceChoixCle_x_tabCleRepart)
 *   body_x_<NomPage>_x_f<Nom>
 * <NomPage> change à chaque écran ; le préfixe body_x_ et le motif
 * _x_tab / _x_f restent. On détecte donc PAR ID, puis on filtre par
 * pertinence (visibilité, nombre de cellules, contenu exploitable).
 *
 * La zone de données peut être rechargée en asynchrone (bouton
 * « Réactualiser la liste des clés de répartition ») : au moment du
 * scan, la table peut être absente ou vide. waitForSigeoTable() attend
 * donc — MutationObserver + re-test périodique — qu'une table
 * pertinente d'au moins minRows lignes soit là avant toute lecture.
 * ==================================================================== */

const OSA_SCAN_TAG = 'NoHands OSA [scan]';

// Motif exact d'un id de table de données : body_x_…_x_tab<Nom> ou _x_f<Nom>.
// Les sélecteurs CSS servent de filet large ; cette regex resserre (elle
// écarte p. ex. un id qui contiendrait « tab » par accident).
const SIGEO_TABLE_ID_RE = /^body_x_.+_x_(?:tab|f)[A-Za-z0-9_]/;

/**
 * Toutes les tables candidates de la page, ciblées PAR ID — jamais par
 * <thead>/<th>/role, inexistants sur ces écrans.
 */
function findSigeoTableCandidates() {
  const nodes = document.querySelectorAll(
    'table[id^="body_x_"][id*="tab"], table[id^="body_x_"][id*="_f"]'
  );
  return Array.from(nodes).filter((t) => SIGEO_TABLE_ID_RE.test(t.id));
}

/**
 * Juge la pertinence d'une table candidate.
 * Renvoie { ok, reason, rows, cells, inputs } ; reason explique le rejet
 * (repris tel quel dans les logs de diagnostic).
 */
function assessSigeoTable(table, minRows) {
  // Menus masqués (content_ivmenu*) et zones repliées : aucune boîte de rendu.
  if (table.offsetWidth === 0 || table.offsetHeight === 0) {
    return { ok: false, reason: 'invisible (largeur ou hauteur nulle)' };
  }
  const rows = table.rows ? table.rows.length : 0;
  const cells = table.querySelectorAll('td, th').length;
  // Pur layout : une seule cellule dans toute la table.
  if (cells <= 1) {
    return { ok: false, reason: 'une seule cellule (mise en page)', rows, cells };
  }
  // Rien d'exploitable : ni texte ni champ de saisie.
  const inputs = table.querySelectorAll('input, select, textarea').length;
  const text = (table.innerText || table.textContent || '').trim();
  if (!text && !inputs) {
    return { ok: false, reason: 'ni texte ni champ exploitable', rows, cells };
  }
  // Trop peu de lignes : zone probablement en cours de rechargement.
  if (rows < minRows) {
    return { ok: false, reason: rows + ' ligne(s) < ' + minRows + ' attendue(s)', rows, cells };
  }
  return { ok: true, rows, cells, inputs };
}

/**
 * Une passe de détection : renvoie { table, diag }.
 * config = { tableId?: id exact ou fragment (ex « tabCleRepart »), minRows? }
 * diag = { candidates, candidateIds, rejected[], kept[], pickedId, pickedWhy }
 */
function pickSigeoTable(config) {
  const mr = parseInt(config.minRows, 10);
  const minRows = Number.isFinite(mr) && mr > 0 ? mr : 1;
  const wanted = String(config.tableId || '').trim();

  const candidates = findSigeoTableCandidates();
  // Un id complet a pu être fourni sans suivre le motif : on l'ajoute au filet.
  if (wanted) {
    const el = document.getElementById(wanted);
    if (el && el.tagName === 'TABLE' && candidates.indexOf(el) === -1) candidates.unshift(el);
  }

  const diag = {
    candidates: candidates.length,
    candidateIds: candidates.map((t) => t.id),
    rejected: [],
    kept: [],
    pickedId: null,
    pickedWhy: ''
  };

  const relevant = [];
  for (const t of candidates) {
    const a = assessSigeoTable(t, minRows);
    if (a.ok) {
      relevant.push({ table: t, rows: a.rows, inputs: a.inputs });
      diag.kept.push(t.id + ' (' + a.rows + ' lignes, ' + a.inputs + ' champs)');
    } else {
      diag.rejected.push(t.id + ' : ' + a.reason);
    }
  }
  if (!relevant.length) return { table: null, diag };

  let pick = null;
  if (wanted) {
    pick = relevant.find((r) => r.table.id === wanted)
        || relevant.find((r) => r.table.id.indexOf(wanted) !== -1);
    if (pick) diag.pickedWhy = 'correspond à « ' + wanted + ' »';
  }
  if (!pick) {
    // Sans cible nommée (ou introuvable) : la table la plus « dense ».
    pick = relevant.reduce((best, r) =>
      (r.rows + r.inputs > best.rows + best.inputs ? r : best));
    diag.pickedWhy = wanted
      ? '« ' + wanted + ' » introuvable — repli sur la plus dense'
      : 'table la plus dense (lignes + champs)';
  }
  diag.pickedId = pick.table.id;
  return { table: pick.table, diag };
}

/**
 * Attend qu'une table de données pertinente soit disponible.
 * MutationObserver pour réagir dès que le DOM bouge, PLUS un re-test
 * périodique de sécurité : les postbacks partiels ASP.NET remplacent des
 * sous-arbres entiers et certains changements (display, valeurs) passent
 * entre les mailles de l'observer. Refuse de conclure tant que l'overlay
 * « Chargement en cours… » est affiché (la table présente serait
 * l'ancienne version, sur le point d'être remplacée).
 * config = { tableId?, minRows?, timeoutMs? (défaut 10000), pollMs? (défaut 300) }
 * Résout TOUJOURS (jamais de rejet) :
 *   { table|null, diag, overlay, timedOut, waitedMs }
 */
function waitForSigeoTable(config) {
  const to = parseInt(config.timeoutMs, 10);
  const timeoutMs = Number.isFinite(to) && to >= 0 ? to : 10000;
  const pollMs = Math.max(100, parseInt(config.pollMs, 10) || 300);
  const started = Date.now();

  return new Promise((resolve) => {
    let observer = null;
    let pollTimer = null;
    let killTimer = null;
    let lastTry = 0;
    let done = false;

    const finish = (table, diag, overlay, timedOut) => {
      if (done) return;
      done = true;
      if (observer) observer.disconnect();
      clearInterval(pollTimer);
      clearTimeout(killTimer);
      resolve({
        table,
        diag,
        overlay: overlay || '',
        timedOut: !!timedOut,
        waitedMs: Date.now() - started
      });
    };

    const attempt = (isLast) => {
      if (done) return true;
      const overlay = detectLoadingOverlay(false);
      const { table, diag } = pickSigeoTable(config);
      if (table && !overlay) { finish(table, diag, '', false); return true; }
      if (isLast) { finish(table, diag, overlay, true); return true; }
      return false;
    };

    if (attempt(timeoutMs === 0)) return; // déjà prêt : aucun délai ajouté

    observer = new MutationObserver(() => {
      // Anti-rafale : un postback déclenche des salves de mutations.
      const now = Date.now();
      if (now - lastTry < 100) return;
      lastTry = now;
      attempt(false);
    });
    observer.observe(document.documentElement, { childList: true, subtree: true });
    pollTimer = setInterval(() => attempt(false), pollMs);
    killTimer = setTimeout(() => attempt(true), timeoutMs);
  });
}

/**
 * Extraction adaptée au format « liste de cases à cocher » (cf.
 * tabCleRepart : chaque ligne = une cellule unique contenant
 * INPUT[checkbox] + LABEL + INPUT[hidden]). Pour chaque ligne : l'état de
 * la case, le libellé du <label>, et la valeur du hidden associé — c'est
 * LUI qui est réellement posté au serveur (cf. syncMirrorHidden).
 * Les lignes sans case (séparateurs éventuels) sont comptées à part.
 */
function extractSigeoChecklist(table) {
  const rows = [];
  let skipped = 0;
  Array.from(table.rows || []).forEach((tr, index) => {
    const cb = tr.querySelector('input[type="checkbox"]');
    if (!cb) { skipped++; return; }

    // Hidden associé : d'abord le miroir hdn<id> (posté au serveur),
    // sinon le premier hidden de la même cellule/ligne.
    let hidden = cb.id ? document.getElementById('hdn' + cb.id) : null;
    if (!hidden) {
      const cell = cb.closest('td') || tr;
      hidden = cell.querySelector('input[type="hidden"]');
    }

    rows.push({
      index,                                  // n° de ligne dans la table
      id: cb.id || cb.name || '',             // clé stable de re-résolution
      checked: !!cb.checked,
      label: batchTargetLabel(cb),            // <label for> sinon texte de cellule/ligne
      hiddenId: hidden ? (hidden.id || hidden.name || '') : '',
      hiddenValue: hidden ? hidden.value : null
    });
  });
  return { rows, skipped };
}

/**
 * Point d'entrée de l'action « scanDataTable » : attend la table de
 * données, extrait la liste de cases, journalise le diagnostic complet
 * dans la console de la page.
 * config = { tableId?, minRows?, timeoutMs?, pollMs? }
 * Renvoie { ok, tableId, rowCount, rows, skippedRows, waitedMs, overlay, diag }.
 */
async function scanSigeoDataTable(config) {
  const waited = await waitForSigeoTable(config || {});
  const diag = waited.diag ||
    { candidates: 0, candidateIds: [], rejected: [], kept: [], pickedId: null, pickedWhy: '' };

  // ---- Logs de diagnostic (console de la page, filtre « OSA ») ----
  console.log(OSA_SCAN_TAG + ' tables candidates : ' + diag.candidates, diag.candidateIds);
  diag.rejected.forEach((r) => console.log(OSA_SCAN_TAG + ' rejetée — ' + r));
  diag.kept.forEach((k) => console.log(OSA_SCAN_TAG + ' pertinente — ' + k));
  if (waited.table) {
    console.log(OSA_SCAN_TAG + ' retenue : #' + diag.pickedId + ' (' + diag.pickedWhy +
      ') après ' + waited.waitedMs + ' ms');
  } else {
    console.warn(OSA_SCAN_TAG + ' AUCUNE table retenue après ' + waited.waitedMs + ' ms' +
      (waited.overlay ? ' — overlay encore affiché : « ' + waited.overlay + ' »' : ''));
  }

  if (!waited.table) {
    return {
      ok: false,
      error: 'aucune table de données pertinente' +
        (waited.timedOut ? ' après ' + waited.waitedMs + ' ms d\'attente' : ''),
      tableId: null,
      rowCount: 0,
      rows: [],
      skippedRows: 0,
      waitedMs: waited.waitedMs,
      overlay: waited.overlay,
      diag
    };
  }

  const extracted = extractSigeoChecklist(waited.table);
  console.log(OSA_SCAN_TAG + ' extraction : ' + extracted.rows.length +
    ' ligne(s) à case, ' + extracted.skipped + ' sans case');

  return {
    ok: true,
    tableId: waited.table.id,
    rowCount: extracted.rows.length,
    rows: extracted.rows,
    skippedRows: extracted.skipped,
    waitedMs: waited.waitedMs,
    overlay: waited.overlay,
    diag
  };
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
