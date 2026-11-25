/* ====================================================================
 * NoHands OSA — Recherche d'immeuble (IM) dans SIGEO à partir d'une adresse
 * --------------------------------------------------------------------
 * Onglet Extraction, type « Recherche IM SIGEO » (adresse → N° E.I.).
 *
 * Constat (test manuel, recherche Nom = « LAMARCK », Site vide → 66
 * résultats sur 5 pages de 15) :
 *   - le champ « Nom » de la recherche est un « contient » sur un libellé
 *     libre, saisi de façon incohérente (LAMARCK "112", 14 LAMARCK,
 *     *98 LAMARCK***30/03/06****) → on ne compare JAMAIS sur « Nom E.I. » ;
 *   - la colonne « Adresse » est normalisée :
 *     « <num>[ BIS|TER] <TYPE VOIE> <NOM VOIE> <CP> <VILLE> » → c'est elle
 *     qui sert à la correspondance.
 *
 * Stratégie :
 *   1. parser l'adresse source (numéros multiples « 158/160 », « 128 128BIS »,
 *      suffixes BIS/TER, voie, CP/ville) ;
 *   2. lancer la recherche avec SEULEMENT le mot le plus distinctif de la
 *      voie (champ Nom, Site forcé à vide) — ou, au choix, par adresse
 *      (Adresse = ce mot, Code postal si connu), ou les deux ;
 *   3. attendre la fin du postback, puis PARCOURIR TOUTES LES PAGES
 *      (« Page suivante >> » jusqu'à désactivation) et vérifier que le
 *      total lu = « N Résultat(s) » ;
 *   4. filtrer ici sur la colonne Adresse (numéro + suffixe exacts, voie,
 *      CP/ville si fournis) ;
 *   5. statut : 1 IM → OK ; 0 → Introuvable (+ 3 adresses les plus
 *      proches) ; > 1 → Multiple (liste des IM, jamais de choix arbitraire) ;
 *   6. cache par recherche (voie → lignes), vidé au début de chaque exécution.
 *
 * Ce fichier est chargé par sidepanel.html AVANT sidepanel.js. Il est aussi
 * importable sous Node (tests) : module.exports en bas de fichier.
 * La fonction imPageOp est injectée telle quelle dans la page (world MAIN) :
 * elle doit rester autonome (aucune référence à ce module).
 * ==================================================================== */

/* ---------------------------------------------------------------- *
 * 1. Normalisation et parsing d'adresses (logique pure)            *
 * ---------------------------------------------------------------- */

// Types de voie → forme canonique. Les abréviations courtes (R, CH, PL…)
// ne sont reconnues qu'en tête de voie (juste après le numéro).
const IM_TYPES_VOIE = (() => {
  const groups = {
    "RUE": ["RUE", "R", "RUES"],
    "BOULEVARD": ["BOULEVARD", "BD", "BLD", "BLVD", "BOUL", "BVD"],
    "AVENUE": ["AVENUE", "AV", "AVE", "AVN"],
    "PLACE": ["PLACE", "PL"],
    "IMPASSE": ["IMPASSE", "IMP"],
    "ALLEE": ["ALLEE", "ALLEES", "ALL", "AL"],
    "CHEMIN": ["CHEMIN", "CHE", "CHEM", "CH"],
    "ROUTE": ["ROUTE", "RTE"],
    "QUAI": ["QUAI", "QU", "QAI"],
    "COURS": ["COURS", "CRS"],
    "SQUARE": ["SQUARE", "SQ"],
    "PASSAGE": ["PASSAGE", "PASS", "PAS", "PSG"],
    "FAUBOURG": ["FAUBOURG", "FBG", "FG"],
    "CITE": ["CITE", "CTE"],
    "RESIDENCE": ["RESIDENCE", "RES", "RESID"],
    "LOTISSEMENT": ["LOTISSEMENT", "LOT"],
    "HAMEAU": ["HAMEAU", "HAM"],
    "PROMENADE": ["PROMENADE", "PROM"],
    "ESPLANADE": ["ESPLANADE", "ESP", "ESPL"],
    "CARREFOUR": ["CARREFOUR", "CAR", "CARR"],
    "TRAVERSE": ["TRAVERSE", "TRA", "TRAV"],
    "VILLA": ["VILLA", "VLA"],
    "SENTIER": ["SENTIER", "SEN", "SENT"],
    "RUELLE": ["RUELLE", "RLE"],
    "MONTEE": ["MONTEE", "MTE"],
    "CHAUSSEE": ["CHAUSSEE", "CHS"],
    "QUARTIER": ["QUARTIER", "QUA", "QRT"],
    "PARVIS": ["PARVIS", "PRV"],
    "MAIL": ["MAIL"],
    "VOIE": ["VOIE"],
    "DOMAINE": ["DOMAINE", "DOM"],
    "LIEU DIT": ["LIEU-DIT", "LIEUDIT", "LD"],
    "ROND POINT": ["ROND-POINT", "RPT", "RDPT"],
    "ZONE": ["ZA", "ZI", "ZAC", "ZONE"]
  };
  const map = Object.create(null);
  for (const [canon, list] of Object.entries(groups)) list.forEach((t) => { map[t] = canon; });
  return map;
})();

// Mots vides ignorés dans la comparaison des noms de voie / villes.
const IM_ARTICLES = new Set(["DE", "DU", "DES", "D", "LA", "LE", "LES", "L", "ET", "AU", "AUX", "A", "EN"]);
// Mots trop fréquents pour servir de terme de recherche (gardés pour la comparaison).
const IM_GENERIQUES = new Set([
  "SAINT", "SAINTE", "GRAND", "GRANDE", "PETIT", "PETITE", "VIEUX", "VIEILLE", "NOTRE", "DAME",
  "GENERAL", "MARECHAL", "DOCTEUR", "PRESIDENT", "PROFESSEUR", "ABBE", "COMMANDANT", "CAPITAINE",
  "COLONEL", "LIEUTENANT", "ROUTE", "RUE", "AVENUE", "BOULEVARD", "PLACE", "CHEMIN", "SUR", "SOUS",
  "NORD", "SUD", "EST", "OUEST", "HAUT", "BAS", "NOUVELLE", "NOUVEAU"
]);
// Bruit fréquent dans les adresses saisies (bâtiment, escalier…).
const IM_BRUIT = new Set(["BAT", "BATIMENT", "ESC", "ESCALIER", "APPT", "APT", "APPARTEMENT", "ETAGE", "ETG", "PORTE", "CEDEX", "BP"]);
const IM_SUFFIXES = { "BIS": "BIS", "B": null, "TER": "TER", "T": null, "QUATER": "QUATER", "QUINQUIES": "QUINQUIES" };

function imFold(s) {
  return String(s ?? "")
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[   ]/g, " ")
    .toUpperCase();
}

// Texte → jetons comparables. Les « 158-160 » / « 158 / 160 » deviennent
// « 158/160 » ; toute autre ponctuation devient espace.
function imTokens(s) {
  let t = imFold(s);
  t = t.replace(/(\d)\s*[\/\-]\s*(?=\d)/g, "$1/");
  t = t.replace(/[^A-Z0-9\/]+/g, " ").replace(/\//g, (m, i, str) =>
    (/\d/.test(str[i - 1] || "") && /\d/.test(str[i + 1] || "")) ? "/" : " ");
  return t.split(/\s+/).filter(Boolean)
    .map((w) => (w === "ST" ? "SAINT" : w === "STE" ? "SAINTE" : w));
}

// Mots significatifs d'un nom de voie / ville (sans articles ni bruit).
function imSignificant(words) {
  return words.filter((w) => !IM_ARTICLES.has(w) && !IM_BRUIT.has(w) && !/^\d+$/.test(w));
}

// Lit les numéros en tête d'une liste de jetons à partir de i.
// Renvoie { nums: ["158","160","128BIS"], next: index suivant }.
function imReadNumbers(tok, i) {
  const nums = [];
  while (i < tok.length) {
    const w = tok[i];
    let m;
    if ((m = w.match(/^(\d{1,4})(BIS|TER|QUATER)?$/))) {
      let n = m[1].replace(/^0+(?=\d)/, "") + (m[2] || "");
      i++;
      const nxt = tok[i];
      if (!m[2] && nxt && Object.prototype.hasOwnProperty.call(IM_SUFFIXES, nxt) && IM_SUFFIXES[nxt]) {
        n += IM_SUFFIXES[nxt];
        i++;
      }
      nums.push(n);
      continue;
    }
    if (/^\d{1,4}(?:BIS|TER)?(?:\/\d{1,4}(?:BIS|TER)?)+$/.test(w)) {
      w.split("/").forEach((p) => nums.push(p.replace(/^0+(?=\d)/, "")));
      i++;
      continue;
    }
    break;
  }
  return { nums, next: i };
}

/**
 * Adresse source (ligne du fichier) → structure.
 *   imParseSource("158/160 RUE LAMARCK")       → nums [158,160], type RUE, voie [LAMARCK]
 *   imParseSource("CHARONNE 128 128BIS")       → nums [128,128BIS], voie [CHARONNE]
 *   imParseSource("SAINT SULPICE 28")          → nums [28], voie [SAINT,SULPICE]
 *   imParseSource("8 BD LAMARCK", {ville:"PARIS"}) → ville [PARIS]
 * extra.cp / extra.ville : valeurs de colonnes séparées (prioritaires si
 * la chaîne n'en contient pas).
 */
function imParseSource(str, extra = {}) {
  const tok = imTokens(str);
  const nums = [];
  const voie = [];
  let cp = "";
  let villeTok = [];
  let i = 0;
  while (i < tok.length) {
    const w = tok[i];
    if (!cp && /^\d{5}$/.test(w)) {
      cp = w;
      villeTok = tok.slice(i + 1).filter((x) => !/^\d+$/.test(x) && x !== "CEDEX");
      break;
    }
    if (/^\d/.test(w)) {
      const r = imReadNumbers(tok, i);
      if (r.nums.length) { nums.push(...r.nums); i = r.next; continue; }
      i++; // nombre non reconnu (ex. 6 chiffres) : ignoré
      continue;
    }
    // Suffixe isolé collé au numéro précédent : « 128 BIS » déjà géré ;
    // « BIS » ailleurs (ex. « CHARONNE BIS ») → ignoré.
    if ((w === "BIS" || w === "TER") && nums.length) { i++; continue; }
    voie.push(w);
    i++;
  }
  let type = "";
  if (voie.length > 1 && IM_TYPES_VOIE[voie[0]]) type = IM_TYPES_VOIE[voie.shift()];
  // Types en deux mots (ROND POINT, LIEU DIT)
  if (!type && voie.length > 2 && ((voie[0] === "ROND" && voie[1] === "POINT") || (voie[0] === "LIEU" && voie[1] === "DIT"))) {
    type = voie[0] + " " + voie[1];
    voie.splice(0, 2);
  }
  const xCp = String(extra.cp ?? "").replace(/\D/g, "");
  if (!cp && /^\d{5}$/.test(xCp)) cp = xCp;
  const xVille = imSignificant(imTokens(extra.ville ?? "")).filter((w) => w !== "CEDEX");
  if (!villeTok.length && xVille.length) villeTok = xVille;
  return {
    raw: String(str ?? ""),
    nums: [...new Set(nums)],
    type,
    voie,
    voieSig: imSignificant(voie),
    cp,
    ville: imSignificant(villeTok)
  };
}

// Adresse normalisée SIGEO (colonne « Adresse ») → structure.
function imParseTarget(str) {
  let tok = imTokens(str);
  // La cellule peut commencer par le nom de l'immeuble :
  //   « grande Armee 62 62 AVENUE DE LA GRANDE ARMEE 75017 PARIS »
  // → l'adresse commence au dernier « <num> <TYPE VOIE> » avant le CP.
  let cpAt = tok.length;
  for (let k = tok.length - 1; k >= 0; k--) { if (/^\d{5}$/.test(tok[k])) { cpAt = k; break; } }
  let start = -1;
  for (let k = 0; k < cpAt; k++) {
    if (!/^\d/.test(tok[k])) continue;
    const rr = imReadNumbers(tok, k);
    if (rr.nums.length && rr.next < cpAt && IM_TYPES_VOIE[tok[rr.next]] && rr.next + 1 < cpAt) start = k;
  }
  // Que des numéros avant (« 158 160 RUE … ») : pas de préfixe, on garde tout.
  // Sinon on coupe le nom en tête (son éventuel numéro, « … 62 », est ignoré).
  if (start > 0 && !tok.slice(0, start).every((w) => /^\d{1,4}(BIS|TER)?$|^(BIS|TER)$/.test(w))) {
    tok = tok.slice(start);
  }
  const r = imReadNumbers(tok, 0);
  let i = r.next;
  const rest = tok.slice(i);
  let cpIdx = -1;
  for (let k = rest.length - 1; k >= 0; k--) { if (/^\d{5}$/.test(rest[k])) { cpIdx = k; break; } }
  let voie = cpIdx >= 0 ? rest.slice(0, cpIdx) : rest.slice();
  const cp = cpIdx >= 0 ? rest[cpIdx] : "";
  const ville = cpIdx >= 0 ? rest.slice(cpIdx + 1).filter((x) => x !== "CEDEX") : [];
  let type = "";
  if (voie.length > 1 && IM_TYPES_VOIE[voie[0]]) type = IM_TYPES_VOIE[voie.shift()];
  if (!type && voie.length > 2 && ((voie[0] === "ROND" && voie[1] === "POINT") || (voie[0] === "LIEU" && voie[1] === "DIT"))) {
    type = voie[0] + " " + voie[1];
    voie = voie.slice(2);
  }
  return { raw: String(str ?? ""), nums: r.nums, type, voie, voieSig: imSignificant(voie), cp, ville: imSignificant(ville) };
}

// Mot le plus distinctif de la voie : le plus long hors mots génériques
// (à égalité, le dernier — « JEAN JAURES » → JAURES).
function imPickSearchTerm(src) {
  const words = (src.voieSig || []).filter((w) => w.length > 1);
  const pick = (list) => list.reduce((best, w) => (!best || w.length >= best.length ? w : best), "");
  return pick(words.filter((w) => !IM_GENERIQUES.has(w))) || pick(words) || "";
}

const imSubset = (a, b) => a.every((w) => b.includes(w));

/**
 * Compare une adresse source à une adresse SIGEO.
 * Renvoie { ok, nameScore (2 = voie identique, 1 = incluse), numHits, why }.
 */
function imMatchAddress(src, tgt) {
  let numHits = 0;
  if (src.nums.length) {
    numHits = src.nums.filter((n) => tgt.nums.includes(n)).length;
    if (!numHits) return { ok: false, why: "numéro" };
  }
  if (src.type && tgt.type && src.type !== tgt.type) return { ok: false, why: "type de voie" };
  const S = src.voieSig, T = tgt.voieSig;
  let nameScore = 0;
  if (S.length && imSubset(S, T) && imSubset(T, S)) nameScore = 2;
  else if (S.length && T.length && (imSubset(S, T) || imSubset(T, S))) nameScore = 1;
  if (!nameScore) return { ok: false, why: "voie" };
  if (src.cp && tgt.cp !== src.cp) return { ok: false, why: "code postal" };
  if (src.ville.length && !(tgt.ville.length && imSubset(src.ville, tgt.ville))) return { ok: false, why: "ville" };
  return { ok: true, nameScore, numHits };
}

// Proximité (pour proposer les adresses les plus proches quand rien ne colle).
function imSimilarity(src, tgt) {
  const S = new Set(src.voieSig), T = new Set(tgt.voieSig);
  const inter = [...S].filter((w) => T.has(w)).length;
  const union = new Set([...S, ...T]).size || 1;
  let score = 3 * inter / union;
  if (src.type && src.type === tgt.type) score += 0.5;
  if (src.nums.length && tgt.nums.length) {
    const n = (x) => parseInt(x, 10);
    const d = Math.min(...src.nums.flatMap((a) => tgt.nums.map((b) => Math.abs(n(a) - n(b)))));
    if (src.nums.some((a) => tgt.nums.includes(a))) score += 2;
    else if (d === 0) score += 1.5;   // même numéro, autre suffixe
    else if (d <= 4) score += 1;
    else if (d <= 20) score += 0.3;
  }
  if (src.cp && src.cp === tgt.cp) score += 1;
  if (src.ville.length && imSubset(src.ville, tgt.ville)) score += 1;
  return score;
}

/**
 * Filtre les lignes de résultats {im, nom, adresse} pour une adresse source.
 * Renvoie { status: "OK"|"Multiple"|"Introuvable", ims, matches, closest }.
 */
// Clé de comparaison d'un nom : jetons normalisés (casse, accents, ST/SAINT…).
const imNameKey = (s) => imTokens(s).join(" ");

function imResolve(src, rows, opts = {}) {
  const scored = [];
  for (const row of rows) {
    const tgt = imParseTarget(row.adresse);
    const m = imMatchAddress(src, tgt);
    if (m.ok) scored.push({ row, tgt, ...m });
  }
  let cands = scored;
  if (cands.length) {
    // Voie identique prioritaire sur voie « incluse ».
    const bestName = Math.max(...cands.map((c) => c.nameScore));
    cands = cands.filter((c) => c.nameScore === bestName);
    // Plage « 158/160 » : préférer l'adresse qui contient tous les numéros.
    if (src.nums.length > 1) {
      const bestHits = Math.max(...cands.map((c) => c.numHits));
      cands = cands.filter((c) => c.numHits === bestHits);
    }
  }
  // Nom E.I. strictement identique à la valeur cherchée (ex. « GRANDE ARMEE 62 »).
  const key = opts.name ? imNameKey(opts.name) : "";
  const byName = key ? rows.filter((row) => imNameKey(row.nom) === key) : [];
  const matchRows = [...byName, ...cands.map((c) => c.row).filter((r) => !byName.includes(r))];
  const ims = [...new Set(matchRows.map((r) => r.im))];
  const status = ims.length === 1 ? "OK" : ims.length ? "Multiple" : "Introuvable";
  const via = byName.length && cands.length ? "nom + adresse" : byName.length ? "nom" : cands.length ? "adresse" : "";
  let closest = [];
  if (!ims.length) {
    const seen = new Set();
    closest = rows
      .map((row) => ({ row, s: imSimilarity(src, imParseTarget(row.adresse)) }))
      .sort((a, b) => b.s - a.s)
      .filter(({ row }) => { const k = row.im + "|" + row.adresse; if (seen.has(k)) return false; seen.add(k); return true; })
      .slice(0, 3)
      .map(({ row }) => row);
  }
  return { status, ims, matches: matchRows, closest, via };
}

// Texte écrit dans la colonne de résultat.
function imStatusText(res) {
  if (res.status === "OK") return "OK — " + res.ims[0];
  if (res.status === "Multiple") return "Multiple — " + res.ims.join(", ");
  if (res.status === "Introuvable") {
    const near = (res.closest || []).map((r) => `${r.adresse} (${r.im})`).join(" ; ");
    return "Introuvable" + (near ? " — proches : " + near : "");
  }
  return "ERREUR — " + (res.error || "inconnue");
}

/* ---------------------------------------------------------------- *
 * 2. Opérations dans la page (injectées, world MAIN, autonomes)    *
 * ---------------------------------------------------------------- */

/**
 * cfg.op :
 *   "probe"   → diagnostic : champs trouvés, bouton, tableau, total…
 *   "prepare" → Site à vide, remplit Nom (ou Adresse + CP), marque les
 *               résultats actuels comme périmés, clique « Rechercher ».
 *   "poll"    → état du postback (prêt ? occupé ? nouveaux résultats ?).
 *   "read"    → lignes {im, nom, adresse} de la page + total + pagination.
 *   "next"    → clique « Page suivante » (ou signale la dernière page).
 * cfg.sel    : sélecteurs forcés { nom, site, adresse, cp, submit, toggle, table }.
 * cfg.ids    : id exacts essayés AVANT la détection par libellés (IM_GER_IDS :
 *              nom, siret, adresse, cp, ville, site, siteLabel, filtre, type,
 *              code, codeCopro, codeProg, submit, table) ; absents → détection.
 * cfg.clearAll     : vide aussi Siret, N° E.I., codes copro / programme,
 *                    Filtre et Type remis sur « Tous » (jamais « Réinit. »).
 * cfg.changeEvents : « input » + « change » sur chaque champ modifié (sauf
 *                    s'il déclencherait un postback).
 * cfg.token  : marqueur des éléments « périmés » (avant postback).
 * Aucune alert()/confirm() ne peut bloquer : le clic sur Rechercher n'est pas
 * fait si Adresse est remplie sans CP ni ville (alerte SIGEO), et les
 * alert/confirm éventuels pendant le clic sont interceptés (statut "alert").
 */
function imPageOp(cfg) {
  const ATTR = "data-osa-im-stale";
  const sel = cfg.sel || {};
  const ids = cfg.ids || {};
  const byId = (k) => {
    if (!ids[k]) return null;
    const el = document.getElementById(ids[k]);
    return el && el.isConnected ? el : null;
  };
  const norm = (s) => String(s ?? "").replace(/[   ]/g, " ").replace(/\s+/g, " ").trim();
  const fold = (s) => norm(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase();
  const isVisible = (el) => {
    if (!el || !el.isConnected) return false;
    try {
      const cs = getComputedStyle(el);
      if (cs.display === "none" || cs.visibility === "hidden") return false;
      return el.getClientRects().length > 0;
    } catch (_) { return false; }
  };
  const resolve = (raw) => {
    const t = String(raw || "").trim();
    if (!t) return null;
    let el = null;
    try { el = document.querySelector(t); } catch (_) { /* pas un sélecteur CSS */ }
    if (el) return el;
    const byName = Array.from(document.getElementsByName(t));
    return byName.find(isVisible) || byName[0] || document.getElementById(t) || null;
  };
  const describe = (el) => {
    if (!el) return "";
    const lbl = norm(el.value && el.tagName !== "SELECT" ? el.value : el.textContent).slice(0, 40);
    return (el.id ? "#" + el.id : el.name ? `[name=${el.name}]` : el.tagName.toLowerCase()) + (lbl ? ` « ${lbl} »` : "");
  };
  const login = !!document.querySelector('input[type="password"]');
  const busy = () => {
    try { if (window.Sys?.WebForms?.PageRequestManager?.getInstance?.()?.get_isInAsyncPostBack?.()) return true; } catch (_) { /* ignoré */ }
    const p = document.getElementById("progessUpdatePanel") || document.querySelector('[id$="progessUpdatePanel"],[id$="UpdateProgress"],.ajax__updateprogress');
    if (p && isVisible(p)) return true;
    for (const el of document.querySelectorAll('[id*="rogress"],.modalPopup,.blockUI,[class*="chargement" i],[class*="loading" i]')) {
      if (isVisible(el) && /chargement\s+en\s+cours/i.test(el.textContent || "")) return true;
    }
    return false;
  };

  // ---- Libellés d'un contrôle (label, cellule voisine, texte précédent)
  function labelTexts(el) {
    const out = [];
    const push = (t) => { t = norm(t).replace(/[:*\s]+$/, ""); if (t && t.length < 60) out.push(t); };
    try {
      if (el.id) document.querySelectorAll(`label[for="${CSS.escape(el.id)}"]`).forEach((l) => push(l.textContent));
    } catch (_) { /* ignoré */ }
    const lab = el.closest("label");
    if (lab) push(lab.textContent);
    push(el.getAttribute("aria-label"));
    push(el.getAttribute("title"));
    push(el.getAttribute("placeholder"));
    let n = el.previousSibling;
    for (let k = 0; n && k < 4; k++, n = n.previousSibling) {
      if (n.nodeType === 3) { if (norm(n.textContent)) { push(n.textContent); break; } continue; }
      if (n.nodeType === 1) {
        if (n.matches("input,select,textarea") || n.querySelector("input,select,textarea")) break;
        if (norm(n.textContent)) { push(n.textContent); break; }
      }
    }
    const cell = el.closest("td,th");
    if (cell) {
      let prev = cell.previousElementSibling;
      while (prev && !norm(prev.textContent)) prev = prev.previousElementSibling;
      if (prev && !prev.querySelector("input,select,textarea")) push(prev.textContent);
    }
    const parent = el.parentElement;
    if (parent && parent !== cell) {
      const ps = parent.previousElementSibling;
      if (ps && !ps.querySelector("input,select,textarea")) push(ps.textContent);
    }
    return out;
  }

  const FIELD_RULES = {
    nom: { tags: "text", label: /^NOM\b(?!.*PRENOM)/, id: /(^|[_:$])(TXT)?(NOM|NAME|LIBELLE|LABEL)([_:$]|$)/ }, // SIGEO : #body_x_txtLabel_x_Label
    site: { tags: "select", label: /^SITE\b/, id: /(^|[_:$])(SEL|DDL|LST|CBO|CMB)?SITES?([_:$]|$)/ },
    adresse: { tags: "text", label: /^(ADRESSE|VOIE)\b/, id: /(^|[_:$])(TXT)?(ADRESSE|ADR|VOIE|RUE)([_:$]|$)/ },
    cp: { tags: "text", label: /^(CODE\s*POSTAL|C\.?\s*P\.?)$|^CODE\s*POSTAL\b/, id: /(^|[_:$])(TXT)?(CP|CODEPOSTAL|CODE_POSTAL|ZIP)([_:$]|$)/ },
    ville: { tags: "text", label: /^(VILLE|COMMUNE)\b/, id: /(^|[_:$])(TXT)?(VILLE|CITY|COMMUNE)([_:$]|$)/ }
  };
  function textInputs() {
    return Array.from(document.querySelectorAll("input")).filter((i) =>
      /^(text|search|)$/i.test(i.getAttribute("type") || "") && !i.disabled && !i.readOnly);
  }
  function submitCandidates() {
    const out = [];
    for (const el of document.querySelectorAll('input[type="submit"],input[type="button"],input[type="image"],button,a')) {
      if (!isVisible(el)) continue;
      const t = fold(el.tagName === "INPUT" ? (el.value || el.alt || el.title) : (el.textContent || el.title));
      const t2 = fold(el.title || el.alt || "");
      if (/^(RECHERCHER|RECHERCHE|CHERCHER|LANCER LA RECHERCHE|FILTRER)$/.test(t) || /^(RECHERCHER|RECHERCHE)$/.test(t2)) out.push(el);
    }
    return out;
  }
  const depthToCommon = (a, b) => {
    if (!a || !b) return 99;
    let d = 0;
    for (let p = a; p; p = p.parentElement, d++) if (p.contains(b)) return d;
    return 99;
  };
  function findField(kind, anchor) {
    if (sel[kind]) return resolve(sel[kind]);
    const exact = byId(kind);
    if (exact) return exact;
    const rule = FIELD_RULES[kind];
    if (!rule) return null;
    const pool = rule.tags === "select" ? Array.from(document.querySelectorAll("select")) : textInputs();
    let best = null, bestScore = 0;
    for (const el of pool) {
      if (!isVisible(el)) continue;
      let score = 0;
      if (labelTexts(el).some((t) => rule.label.test(fold(t)))) score = 3;
      else if (rule.id.test(fold(el.id || "")) || rule.id.test(fold(el.name || ""))) score = 2;
      if (!score) continue;
      if (anchor) score += Math.max(0, 1 - depthToCommon(el, anchor) / 12);
      if (score > bestScore) { best = el; bestScore = score; }
    }
    return best;
  }
  function findSubmit(anchor) {
    if (sel.submit) return resolve(sel.submit);
    const exact = byId("submit");
    if (exact) return exact;
    const c = submitCandidates();
    if (!anchor) return c[0] || null;
    return c.sort((a, b) => depthToCommon(a, anchor) - depthToCommon(b, anchor))[0] || null;
  }
  function findToggle(text) {
    if (sel.toggle && cfg.mode === "adresse") return resolve(sel.toggle);
    // 1. Boutons radio dont le libellé est « Par nom » / « Par adresse »
    for (const el of document.querySelectorAll('input[type="radio"]')) {
      if (isVisible(el) && labelTexts(el).some((t) => fold(t) === text)) return el;
    }
    // 2. Onglets / liens / libellés portant exactement ce texte
    for (const el of document.querySelectorAll("a,label,span,li,div,button,td")) {
      if (el.children.length > 2 || fold(el.textContent) !== text || !isVisible(el)) continue;
      if (el.tagName === "LABEL" && el.control) return el.control;
      const radio = el.querySelector('input[type="radio"],input[type="checkbox"]');
      return radio || el;
    }
    return null;
  }

  // ---- Tableau de résultats (en-têtes « N° E.I. », « Nom E.I. », « Adresse »)
  function cellAt(row, logical) {
    let pos = 0;
    for (const c of row.cells) {
      const span = c.colSpan || 1;
      if (logical >= pos && logical < pos + span) return c;
      pos += span;
    }
    return null;
  }
  function headerInfo(table) {
    const rows = Array.from(table.rows);
    for (let r = 0; r < Math.min(rows.length, 6); r++) {
      const idx = { num: -1, nom: -1, adr: -1 };
      let pos = 0;
      for (const c of rows[r].cells) {
        const t = fold(c.textContent);
        if (idx.num < 0 && /^N\W*\s*E\W*\s*I\W*$/.test(t)) idx.num = pos;
        else if (idx.nom < 0 && /^NOM\b/.test(t)) idx.nom = pos;
        else if (idx.adr < 0 && /^ADRESSE\b/.test(t)) idx.adr = pos;
        pos += c.colSpan || 1;
      }
      if (idx.adr >= 0 && (idx.num >= 0 || idx.nom >= 0)) return { headerRow: r, ...idx };
    }
    return null;
  }
  function findTable() {
    const grid = !sel.table && byId("table");
    if (grid && grid.tagName === "TABLE" && isVisible(grid)) {
      const h = headerInfo(grid);
      // Grille SIGEO #body_x_grid_x : [0] icônes, [1] N° E.I., [2] Nom E.I., [3] Adresse.
      return h ? { table: grid, ...h } : { table: grid, headerRow: -1, num: 1, nom: 2, adr: 3 };
    }
    let pool;
    if (sel.table) {
      const el = resolve(sel.table);
      pool = el ? (el.tagName === "TABLE" ? [el, ...el.querySelectorAll("table")] : Array.from(el.querySelectorAll("table"))) : [];
    } else pool = Array.from(document.querySelectorAll("table"));
    for (const t of pool) {
      if (!isVisible(t)) continue;
      const h = headerInfo(t);
      if (h) return { table: t, ...h };
    }
    return null;
  }
  function readRows(info) {
    const out = [];
    const rows = Array.from(info.table.rows);
    for (let r = info.headerRow + 1; r < rows.length; r++) {
      const row = rows[r];
      const txt = (i) => (i >= 0 && cellAt(row, i) ? norm(cellAt(row, i).textContent) : "");
      let im = (txt(info.num).match(/IM\d{7}/i) || [])[0];
      if (!im) {
        const first = Array.from(row.cells).find((c) => norm(c.textContent));
        im = first && (norm(first.textContent).match(/IM\d{7}/i) || [])[0];
      }
      if (!im) continue; // pagination, ligne vide, séparateur…
      // Id interne de l'immeuble : fin du lien de la fiche (immeuble_manage/ger/399982).
      const a = row.querySelector('a.edit_link[href], a[href*="immeuble_manage"]');
      const id = a ? ((a.getAttribute("href") || "").match(/\/(\d+)\/?(?:[?#].*)?$/) || [])[1] || "" : "";
      out.push({ im: im.toUpperCase(), nom: txt(info.nom), adresse: txt(info.adr), id });
    }
    return out;
  }
  // Éléments « N Résultat(s) » / « Aucun résultat » (les plus petits).
  function countEls() {
    const out = [];
    for (const el of document.querySelectorAll("span,div,td,p,b,strong,label,font,th,li,h1,h2,h3,h4,em")) {
      const t = norm(el.textContent);
      if (!t || t.length > 160) continue;
      if (!/\d+\s*r[ée]sultat|aucun\s+r[ée]sultat/i.test(t)) continue;
      if (Array.from(el.children).some((ch) => /\d+\s*r[ée]sultat|aucun\s+r[ée]sultat/i.test(ch.textContent || ""))) continue;
      out.push(el);
    }
    return out;
  }
  function readCount(els) {
    for (const el of els) {
      if (el.hasAttribute(ATTR)) continue;
      const t = norm(el.textContent);
      const m = t.match(/(\d[\d\s.]*)\s*r[ée]sultat/i);
      if (m) return { count: parseInt(m[1].replace(/\D/g, ""), 10), zero: false };
      if (/aucun\s+r[ée]sultat/i.test(t)) return { count: 0, zero: true };
    }
    return { count: null, zero: false };
  }
  function findNext(info) {
    const scope = info ? (info.table.closest("div,form,body") || document) : document;
    const look = (root) => {
      for (const el of root.querySelectorAll('a,input[type="submit"],input[type="button"],input[type="image"],button,span')) {
        const lbl = fold(el.tagName === "INPUT" ? (el.value || el.alt || el.title) : (el.textContent || el.title || ""));
        if (!/PAGE\s+SUIVANTE|^SUIVANT(E)?\s*>*$|^>>$/.test(lbl)) continue;
        if (!isVisible(el)) continue;
        if (el.tagName === "SPAN" && el.querySelector("a,input,button")) continue;
        return el;
      }
      return null;
    };
    const el = (info && look(info.table)) || look(scope) || look(document);
    if (!el) return { el: null, enabled: false };
    const href = el.getAttribute("href");
    const disabled = el.disabled || el.getAttribute("aria-disabled") === "true" ||
      /disabled/i.test(el.className || "") || el.hasAttribute("disabled") ||
      el.tagName === "SPAN" ||
      (el.tagName === "A" && !href && !el.getAttribute("onclick"));
    return { el, enabled: !disabled };
  }
  function markStale(info) {
    const tok = cfg.token;
    document.documentElement.setAttribute("data-osa-im-doc", tok);
    if (info) info.table.setAttribute(ATTR, tok);
    countEls().forEach((el) => el.setAttribute(ATTR, tok));
  }
  function setVal(el, v) {
    if (el.value === v) return false;
    const proto = el.tagName === "TEXTAREA" ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, "value")?.set;
    if (setter) setter.call(el, v); else el.value = v;
    // Pas de blur, et « change » seulement si demandé (cfg.changeEvents) ET
    // sans postback : un TextBox AutoPostBack déclencherait un __doPostBack
    // avant le clic sur Rechercher.
    el.dispatchEvent(new Event("input", { bubbles: true }));
    if (cfg.changeEvents && !triggersPostback(el)) el.dispatchEvent(new Event("change", { bubbles: true }));
    return true;
  }
  const triggersPostback = (el) => /__doPostBack|WebForm_DoPostBack|AutoPostBack/i.test(
    (el.getAttribute("onchange") || "") + (el.getAttribute("onclick") || "") + (el.getAttribute("href") || ""));

  const op = cfg.op;

  if (op === "probe") {
    const nom = findField("nom");
    const main = cfg.mode === "adresse" ? findField("adresse") : nom;
    const toggle = findToggle(cfg.mode === "adresse" ? "PAR ADRESSE" : "PAR NOM");
    const anchor = main || nom || findField("adresse") || toggle;
    const submit = findSubmit(anchor);
    const info = findTable();
    const ce = countEls();
    const nx = findNext(info);
    const fields = {};
    for (const k of ["nom", "site", "adresse", "cp"]) {
      const f = k === "nom" ? nom : findField(k, anchor);
      fields[k] = f ? describe(f) : "";
    }
    return {
      login, url: location.href, title: document.title, fields,
      submit: submit ? describe(submit) : "",
      table: info ? (info.table.id ? "#" + info.table.id : "table") : "",
      rows: info ? readRows(info).length : 0,
      count: readCount(ce).count,
      next: nx.el ? (nx.enabled ? "actif" : "désactivé") : "",
      toggle: !!toggle,
      siteInfo: (() => {
        const st = findField("site", anchor);
        if (!st || st.tagName !== "SELECT") return null;
        const o = st.options[st.selectedIndex];
        return { value: st.value, text: o ? norm(o.textContent) : "" };
      })(),
      // Options de la liste « Site » (réglage « Site SIGEO » du panneau).
      siteOptions: (() => {
        const st = findField("site", anchor);
        if (!st || st.tagName !== "SELECT") return [];
        return Array.from(st.options).slice(0, 400).map((o) => ({ value: o.value, text: norm(o.textContent) }));
      })(),
      busy: busy(),
      // Champ principal visible, ou masqué derrière la bascule « Par nom /
      // Par adresse » (l'opération prepare cliquera dessus).
      ok: !!(main || toggle) && !!submit
    };
  }

  if (op === "prepare") {
    if (login) return { status: "login" };
    // 1. Bascule « Par nom » / « Par adresse » si le formulaire en a une.
    const mainKind = cfg.mode === "adresse" ? "adresse" : "nom";
    let main = findField(mainKind);
    if (!cfg.noToggle) {
      const tg = findToggle(cfg.mode === "adresse" ? "PAR ADRESSE" : "PAR NOM");
      const needs = tg && (!main || !isVisible(main) || (tg.type === "radio" && !tg.checked));
      if (needs) {
        markStale(null);
        tg.click();
        return { status: triggersPostback(tg) ? "toggle_postback" : "toggled", detail: describe(tg) };
      }
    }
    if (!main) return { status: "no_field", detail: mainKind };
    // 2. Site : valeur imposée (cfg.siteValue), ou forcé à vide (option « »).
    const wantSite = typeof cfg.siteValue === "string";
    if (wantSite || cfg.siteEmpty) {
      const site = findField("site", main);
      if (wantSite && !(site && site.tagName === "SELECT")) return { status: "no_site_field" };
      if (site && site.tagName === "SELECT") {
        const opt = wantSite ? Array.from(site.options).find((o) => o.value === cfg.siteValue)
          : Array.from(site.options).find((o) => o.value === "") ||
          Array.from(site.options).find((o) => /^(|TOUS|\(TOUS\)|TOUS LES SITES|-+|\*)$/.test(fold(o.textContent)));
        if (!opt) return { status: wantSite ? "no_site_value" : "no_site_empty", detail: describe(site) };
        if (site.value !== opt.value) {
          site.value = opt.value;
          if (triggersPostback(site)) {
            markStale(null);
            site.setAttribute(ATTR, cfg.token);
            site.dispatchEvent(new Event("change", { bubbles: true }));
            return { status: "site_postback", detail: describe(site) };
          }
          // SIGEO : le onchange copie seulement le libellé dans le champ caché.
          site.dispatchEvent(new Event("change", { bubbles: true }));
        }
        // Libellé caché du site (body_x_sites_x_hdSiteLabel) : écrit aussi ici,
        // pour que le serveur ne reçoive jamais un site incohérent.
        const lab = byId("siteLabel");
        if (lab && lab.value !== norm(opt.textContent)) lab.value = norm(opt.textContent);
      }
    }
    // 3. Critères : le champ principal reçoit le terme, l'autre mode est vidé.
    main = findField(mainKind);
    const filled = [];
    setVal(main, cfg.term);
    filled.push(`${mainKind}=${cfg.term}`);
    // Les critères d'une autre recherche restés dans le formulaire (Nom tapé
    // pour la ligne précédente, Ville…) se combineraient avec celui-ci.
    if (cfg.mode === "adresse") {
      const cp = findField("cp", main);
      if (cp) { setVal(cp, cfg.cp || ""); filled.push(`cp=${cfg.cp || ""}`); }
      for (const k of ["nom", "ville"]) {
        const f = findField(k, main);
        if (f && f !== main && f !== cp && f.value) { setVal(f, ""); filled.push(`${k} vidé`); }
      }
    } else {
      for (const k of ["adresse", "cp", "ville"]) {
        const f = findField(k, main);
        if (f && f !== main && f.value) { setVal(f, ""); filled.push(`${k} vidé`); }
      }
    }
    if (cfg.clearAll) {
      // Tous les autres critères (ils se combinent en ET) — jamais « Réinit. ».
      for (const k of ["siret", "code", "codeCopro", "codeProg"]) {
        const f = byId(k);
        if (f && f !== main && f.value) { setVal(f, ""); filled.push(`${k} vidé`); }
      }
      for (const k of ["filtre", "type"]) {
        const f = byId(k);
        if (f && f.tagName === "SELECT" && f.value !== "" && Array.from(f.options).some((o) => o.value === "") && !triggersPostback(f)) {
          f.value = "";
          f.dispatchEvent(new Event("change", { bubbles: true }));
          filled.push(`${k}=Tous`);
        }
      }
    }
    // 4. Alerte SIGEO « Merci de préciser un code postal ou une ville » :
    // Adresse remplie sans CP ni ville → on ne clique pas (DPE extraction ;
    // ailleurs, l'alerte éventuelle est interceptée ci-dessous).
    if (cfg.clearAll) {
      const fa = findField("adresse", main), fc = findField("cp", main), fv = findField("ville", main);
      if (fa && norm(fa.value) && !(fc && norm(fc.value)) && !(fv && norm(fv.value))) return { status: "would_alert", filled };
    }
    // 5. Résultats actuels marqués périmés, puis clic sur Rechercher. Une
    // alert()/confirm() pendant le clic est interceptée (jamais bloquante).
    const submit = findSubmit(main);
    if (!submit) return { status: "no_submit" };
    markStale(findTable());
    const alerts = [];
    const oldAlert = window.alert, oldConfirm = window.confirm;
    window.alert = (m) => { alerts.push(String(m ?? "")); };
    window.confirm = (m) => { alerts.push(String(m ?? "")); return false; };
    try { submit.click(); } finally { window.alert = oldAlert; window.confirm = oldConfirm; }
    if (alerts.length) {
      document.documentElement.removeAttribute("data-osa-im-doc");
      return { status: "alert", detail: alerts.join(" / "), filled };
    }
    return { status: "submitted", detail: describe(submit), filled };
  }

  if (op === "poll") {
    const info = findTable();
    const stale = !!document.querySelector(`[${ATTR}="${cfg.token}"]`);
    const c = readCount(countEls());
    return {
      login,
      ready: document.readyState === "complete",
      busy: busy(),
      docReplaced: document.documentElement.getAttribute("data-osa-im-doc") !== cfg.token,
      stale,
      table: !!info && !info.table.hasAttribute(ATTR),
      count: c.count,
      zero: c.zero
    };
  }

  if (op === "read") {
    const info = findTable();
    const c = readCount(countEls().filter((el) => !el.hasAttribute(ATTR) || el.getAttribute(ATTR) !== cfg.token));
    const nx = findNext(info);
    const rows = info ? readRows(info) : [];
    return {
      found: !!info,
      rows,
      count: c.count,
      zero: c.zero,
      next: !!nx.el && nx.enabled,
      sig: rows.length ? rows[0].im + "|" + rows[0].adresse + "|" + rows[rows.length - 1].im + "|" + rows.length : ""
    };
  }

  if (op === "next") {
    const info = findTable();
    const nx = findNext(info);
    if (!nx.el || !nx.enabled) return { status: "last" };
    markStale(info);
    nx.el.click();
    return { status: "clicked", detail: describe(nx.el) };
  }

  return { status: "unknown_op" };
}

/* ---------------------------------------------------------------- *
 * 3. Orchestration (recherche + pagination + cache)                *
 * ---------------------------------------------------------------- */

/*
 * L'orchestrateur ne dépend pas de chrome.* : il reçoit un adaptateur
 *   ad.execAll(cfg)          → [{ frameId, result }] (imPageOp dans toutes les frames)
 *   ad.exec(frameId, cfg)    → résultat de imPageOp (lève une erreur si la
 *                              frame navigue / n'est pas joignable)
 *   ad.navigate(url)         → { ok, error? }
 *   ad.sleep(ms), ad.stopped(), ad.log(text, cls)
 * Ainsi le même code tourne dans l'extension et dans les tests (Playwright).
 */

let imTokenSeq = 0;
const imNewToken = () => "osa" + Date.now().toString(36) + (++imTokenSeq);

// Cherche la frame qui porte le formulaire de recherche (nav si besoin).
async function imLocate(ad, p) {
  const tryOnce = async () => {
    let res = [];
    try { res = await ad.execAll({ op: "probe", mode: p.mode, sel: p.sel, ids: p.ids }); } catch (_) { /* navigation */ }
    const hits = (res || []).filter((r) => r && r.result);
    if (hits.some((r) => r.result.login)) return { auth: true };
    const ok = hits.filter((r) => r.result.ok)
      .sort((a, b) => (b.result.table ? 1 : 0) - (a.result.table ? 1 : 0));
    return ok.length ? { frameId: ok[0].frameId, probe: ok[0].result } : { probes: hits.map((h) => h.result) };
  };
  let r = await tryOnce();
  if (r.frameId !== undefined || r.auth) return r;
  if (p.url) {
    const nav = await ad.navigate(p.url);
    if (!nav.ok) return { error: "navigation vers la page de recherche : " + nav.error };
    const deadline = Date.now() + (p.timeoutMs || 20000);
    while (Date.now() < deadline && !ad.stopped()) {
      r = await tryOnce();
      if (r.frameId !== undefined || r.auth) return r;
      await ad.sleep(400);
    }
  }
  const pr = (r.probes || [])[0];
  let why = "formulaire de recherche introuvable — ouvre la recherche des immeubles dans l'onglet cible (ou renseigne l'URL de la page)";
  if (pr) {
    const k = p.mode === "adresse" ? "adresse" : "nom";
    if (!pr.fields[k]) why += ` ; champ « ${k === "nom" ? "Nom" : "Adresse"} » non trouvé (désigne-le avec 🎯)`;
    else if (!pr.submit) why += " ; bouton « Rechercher » non trouvé (désigne-le avec 🎯)";
  }
  return { error: why };
}

/**
 * Attend la fin du postback déclenché après marquage `token`.
 * kind = "results" : de nouveaux résultats (tableau, « N Résultat(s) » ou
 *        « Aucun résultat ») doivent être visibles ;
 * kind = "settle"  : la page doit simplement être revenue au repos.
 * Accepte aussi bien un postback partiel (UpdatePanel) qu'un rechargement
 * complet de la page.
 */
async function imWaitForSearchDone(ad, frameId, token, timeoutMs, kind = "results", ids) {
  const t0 = Date.now();
  const deadline = t0 + Math.max(2000, timeoutMs || 20000);
  let stable = 0, sawBusy = false, idleSince = null, last = null;
  while (Date.now() < deadline) {
    if (ad.stopped()) return { ok: false, error: "arrêté par l'utilisateur" };
    await ad.sleep(200);
    let r;
    try { r = await ad.exec(frameId, { op: "poll", token, ids }); }
    catch (_) { sawBusy = true; stable = 0; idleSince = null; continue; } // rechargement en cours
    if (!r) { stable = 0; continue; }
    last = r;
    if (r.login) return { ok: false, auth: true, error: "AUTH_REQUIRED — page de connexion détectée (session SIGEO expirée ?)" };
    if (r.busy || !r.ready) { sawBusy = true; stable = 0; idleSince = null; continue; }
    if (idleSince === null) idleSince = Date.now();
    let done;
    if (kind === "settle") {
      done = r.docReplaced || !r.stale || (sawBusy && Date.now() - idleSince > 600) || Date.now() - t0 > 2500;
    } else {
      const fresh = !r.stale && (r.table || r.count !== null || r.zero);
      // Postback vu puis page au repos, mais le tableau n'a pas été recréé
      // (mise à jour « sur place ») : on accepte après un court délai.
      const inPlace = sawBusy && r.stale && Date.now() - idleSince > 1500;
      done = fresh || inPlace;
    }
    if (done) { if (++stable >= 2) return { ok: true, poll: r }; }
    else stable = 0;
  }
  return {
    ok: false,
    error: `pas de réponse de la recherche (délai ${Math.round((timeoutMs || 20000) / 1000)} s dépassé` +
      (last ? (last.busy ? ", page toujours en chargement" : last.stale ? ", résultats non rafraîchis" : "") : ", page injoignable") + ")"
  };
}

// Lecture d'une page avec quelques reprises (DOM en cours de remplacement).
async function imReadPage(ad, frameId, token, ids) {
  for (let k = 0; k < 6; k++) {
    try {
      const r = await ad.exec(frameId, { op: "read", token, ids });
      if (r && (r.found || r.zero || r.count === 0)) return r;
    } catch (_) { /* frame en transition */ }
    await ad.sleep(300);
  }
  return null;
}

/**
 * Lance UNE recherche et parcourt toutes les pages.
 * p = { mode: "nom"|"adresse", term, cp, siteEmpty, siteValue, sel, ids, clearAll,
 *       changeEvents, url, timeoutMs, maxPages }
 * Renvoie { ok, rows, total, pages, warnings, error?, siteMissing? }.
 */
async function imSearchAllPages(ad, p) {
  const warnings = [];
  const loc = await imLocate(ad, p);
  if (loc.auth) return { ok: false, error: "AUTH_REQUIRED — page de connexion détectée (session SIGEO expirée ?)" };
  if (loc.error) return { ok: false, error: loc.error };
  const frameId = loc.frameId;

  // — Préparation + clic sur Rechercher (reprises après postback Site / bascule)
  let token = null, noToggle = false;
  for (let attempt = 1; ; attempt++) {
    if (ad.stopped()) return { ok: false, error: "arrêté par l'utilisateur" };
    if (attempt > 5) return { ok: false, error: "le formulaire de recherche ne se stabilise pas (postbacks répétés)" };
    token = imNewToken();
    let r;
    try {
      r = await ad.exec(frameId, { op: "prepare", token, mode: p.mode, term: p.term, cp: p.cp || "", siteEmpty: p.siteEmpty !== false, siteValue: p.siteValue,
        sel: p.sel, ids: p.ids, clearAll: !!p.clearAll, changeEvents: !!p.changeEvents, noToggle });
    } catch (e) {
      await ad.sleep(500);
      continue;
    }
    if (!r) return { ok: false, error: "préparation de la recherche : aucune réponse de la page" };
    if (r.status === "submitted") break;
    if (r.status === "login") return { ok: false, error: "AUTH_REQUIRED — page de connexion détectée (session SIGEO expirée ?)" };
    if (r.status === "no_field") return { ok: false, error: `champ « ${r.detail === "nom" ? "Nom" : "Adresse"} » introuvable — désigne-le avec 🎯 (sélecteurs avancés)` };
    if (r.status === "no_submit") return { ok: false, error: "bouton « Rechercher » introuvable — désigne-le avec 🎯 (sélecteurs avancés)" };
    if (r.status === "no_site_empty") return { ok: false, error: `liste « Site » ${r.detail} : aucune option vide` };
    if (r.status === "no_site_value") return { ok: false, siteMissing: true, error: `site « ${p.siteValue} » introuvable dans la liste « Site » (${r.detail})` };
    if (r.status === "no_site_field") return { ok: false, siteMissing: true, error: `liste « Site » introuvable : impossible de choisir le site « ${p.siteValue} »` };
    if (r.status === "would_alert") return { ok: false, error: "Adresse remplie sans code postal ni ville : SIGEO afficherait « Merci de préciser un code postal ou une ville » — recherche non lancée" };
    if (r.status === "alert") return { ok: false, error: `SIGEO a affiché un message (intercepté, recherche annulée) : « ${r.detail} »` };
    if (r.status === "toggled") { noToggle = true; await ad.sleep(300); continue; } // bascule côté client
    if (r.status === "site_postback" || r.status === "toggle_postback") {
      if (r.status === "toggle_postback") noToggle = true;
      const w = await imWaitForSearchDone(ad, frameId, token, p.timeoutMs, "settle", p.ids);
      if (!w.ok && w.auth) return { ok: false, error: w.error };
      continue;
    }
    return { ok: false, error: "préparation de la recherche : " + r.status };
  }

  const w = await imWaitForSearchDone(ad, frameId, token, p.timeoutMs, "results", p.ids);
  if (!w.ok) return { ok: false, error: w.error };

  // — Lecture de la page 1 puis des suivantes
  const rows = [];
  const keys = new Set();
  const add = (list) => list.forEach((r) => {
    const k = r.im + "|" + r.nom + "|" + r.adresse;
    if (!keys.has(k)) { keys.add(k); rows.push(r); }
  });
  let page = await imReadPage(ad, frameId, token, p.ids);
  if (!page) return { ok: false, error: "tableau de résultats illisible (colonnes « N° E.I. » / « Adresse » introuvables)" };
  const total = page.count;
  add(page.rows);
  let pages = 1;
  const sigs = new Set([page.sig]);
  const maxPages = Math.max(1, p.maxPages || 30);
  while (page.next && pages < maxPages) {
    if (total !== null && rows.length >= total) break;
    if (ad.stopped()) return { ok: false, error: "arrêté par l'utilisateur" };
    token = imNewToken();
    let r;
    try { r = await ad.exec(frameId, { op: "next", token, ids: p.ids }); }
    catch (e) { return { ok: false, error: "clic « Page suivante » impossible : " + (e.message || e) }; }
    if (!r || r.status === "last") break;
    const wn = await imWaitForSearchDone(ad, frameId, token, p.timeoutMs, "results", p.ids);
    if (!wn.ok) return { ok: false, error: `page ${pages + 1} : ${wn.error}` };
    page = await imReadPage(ad, frameId, token, p.ids);
    if (!page) return { ok: false, error: `page ${pages + 1} illisible` };
    if (sigs.has(page.sig)) { warnings.push(`page ${pages + 1} identique à une page déjà lue — arrêt de la pagination`); break; }
    sigs.add(page.sig);
    add(page.rows);
    pages++;
  }
  if (page.next && pages >= maxPages) warnings.push(`limite de ${maxPages} pages atteinte`);
  if (total !== null && rows.length !== total) warnings.push(`${rows.length} ligne(s) lue(s) sur ${total} annoncée(s)`);

  // — Cohérence : une recherche par Nom « contient » le terme sur chaque ligne.
  if (rows.length) {
    const col = p.mode === "adresse" ? "adresse" : "nom";
    const T = imNameKey(p.term);
    if (T && !rows.some((r) => imNameKey(r[col]).includes(T))) {
      return { ok: false, error: `résultats incohérents : aucune ligne ne contient « ${p.term} » (la recherche n'a pas été prise en compte ?)` };
    }
  }
  return { ok: true, rows, total, pages, warnings };
}

// Cache des recherches : clé → Promise du résultat. Vidé à chaque exécution.
const imSearchCache = new Map();
function imClearCache() { imSearchCache.clear(); }

async function imCachedSearch(ad, p) {
  const key = [p.mode, imFold(p.term), p.mode === "adresse" ? (p.cp || "") : "",
    typeof p.siteValue === "string" ? "site=" + p.siteValue : p.siteEmpty !== false ? "site=" : "site*"].join("|");
  if (imSearchCache.has(key)) {
    const res = await imSearchCache.get(key);
    return { ...res, cached: true };
  }
  const prom = imSearchAllPages(ad, p);
  imSearchCache.set(key, prom);
  const res = await prom;
  if (!res.ok) imSearchCache.delete(key); // une erreur n'est pas mise en cache
  return res;
}

/**
 * Recherche complète pour une ligne du fichier.
 * cfg = { adresse, cp, ville, term?, mode, siteEmpty, sel, url, timeoutMs, maxPages }
 * mode :
 *   "auto"    → tape la valeur COMPLÈTE dans Nom (« GRANDE ARMEE 62 »), valide,
 *               lit le tableau ; si rien ne correspond, relance avec le seul
 *               nom de voie (« ARMEE ») ;
 *   "nom"     → nom de voie seul dans Nom ;
 *   "adresse" → nom de voie dans « Par adresse » (+ CP) ;
 *   "both"    → nom de voie, par Nom ET par adresse.
 * Une ligne du tableau correspond si son Nom E.I. est identique à la valeur
 * cherchée, ou si son Adresse correspond (numéro, voie, CP/ville).
 * Renvoie { status: "OK"|"Multiple"|"Introuvable"|"ERREUR", ims, closest,
 *           term, searches: [{mode, term, total, pages, rows, cached, warnings, matches, ims}], error? }
 */
async function imLookupAddress(ad, cfg) {
  const raw = String(cfg.adresse ?? "").replace(/[\u00a0\u202f\u2007]/g, " ").replace(/\s+/g, " ").trim();
  const src = imParseSource(raw, { cp: cfg.cp, ville: cfg.ville });
  const forced = String(cfg.term || "").replace(/\s+/g, " ").trim();
  const voieTerm = imPickSearchTerm(src);
  const mode = ["auto", "nom", "adresse", "both"].includes(cfg.mode) ? cfg.mode : "auto";

  // Étapes successives ; on s'arrête dès qu'une étape donne au moins une correspondance.
  let stages;
  if (forced) {
    stages = [mode === "both" ? [["nom", forced], ["adresse", forced]] : [[mode === "adresse" ? "adresse" : "nom", forced]]];
  } else if (mode === "auto") {
    stages = [[["nom", raw]]];
    if (voieTerm && imNameKey(voieTerm) !== imNameKey(raw)) stages.push([["nom", voieTerm]]);
  } else if (mode === "both") {
    stages = [[["nom", voieTerm], ["adresse", voieTerm]]];
  } else {
    stages = [[[mode, voieTerm]]];
  }
  if (!raw || stages.some((st) => st.some(([, t]) => !t))) {
    return { status: "ERREUR", ims: [], closest: [], error: `valeur illisible : « ${raw} » (aucun nom de voie)`, src };
  }

  const searches = [];
  const all = [];
  const keys = new Set();
  let res = null;
  for (const stage of stages) {
    for (const [m, term] of stage) {
      const r = await imCachedSearch(ad, {
        mode: m, term, cp: src.cp, siteEmpty: cfg.siteEmpty, siteValue: cfg.siteValue, sel: cfg.sel || {}, url: cfg.url,
        ids: cfg.ids, clearAll: cfg.clearAll, changeEvents: cfg.changeEvents,
        timeoutMs: cfg.timeoutMs, maxPages: cfg.maxPages
      });
      if (!r.ok) {
        return { status: "ERREUR", ims: [], closest: [], error: `recherche ${m === "nom" ? "Nom" : "par adresse"} « ${term} » : ${r.error}`, siteMissing: !!r.siteMissing, src, term, searches };
      }
      const own = imResolve(src, r.rows, { name: raw });
      searches.push({
        mode: m, term, total: r.total, pages: r.pages, rows: r.rows.length, cached: !!r.cached,
        warnings: r.warnings || [], matches: own.matches.length, ims: own.ims
      });
      for (const row of r.rows) {
        const k = row.im + "|" + row.nom + "|" + row.adresse;
        if (!keys.has(k)) { keys.add(k); all.push(row); }
      }
    }
    res = imResolve(src, all, { name: raw });
    if (res.ims.length) break;
  }
  return { ...res, src, term: searches.length ? searches[searches.length - 1].term : voieTerm, searches, rowsRead: all.length };
}

/* ---------------------------------------------------------------- *
 * 4. Recherche « par désignation + site » (DPE extraction, v2.22.0)  *
 * ---------------------------------------------------------------- */

// Id exacts de la recherche d'immeubles SIGEO
// (http://sigeo.evoriel.net/page.aspx/fr/sig_imm/immeuble_browse/ger).
// Absents de la page → détection par libellés (imPageOp).
const IM_GER_IDS = {
  nom: "body_x_txtLabel_x_Label",
  siret: "body_x_txtSiret_x_txtSiret",
  adresse: "body_x_txtAdresse_x_txtAdresse",
  cp: "body_x_txtCP_x_CP",
  ville: "body_x_txtVille_x_Ville",
  site: "body_x_sites_x_ddlSite_x",
  siteLabel: "body_x_sites_x_hdSiteLabel",
  filtre: "body_x_ddlMetier_x",
  type: "body_x_ddlTypeMan_x",
  code: "body_x_txtCode_x_Code",
  codeCopro: "body_x_txtCodeCopro_x_txtCodeCopro",
  codeProg: "body_x_txtCodeProgrammeCommercial_x_txtCodeProgrammeCommercial",
  submit: "body_x_searchRaz_x_cmdSearchBtn_x__btn",
  table: "body_x_grid_x"
};
const IM_DEFAULT_SITE = "391"; // 391 - ORALIA SOGIMAT

// Nom E.I. comparable : majuscules, sans accents, espaces réduits.
const imStrictKey = (s) => imFold(s).replace(/\s+/g, " ").trim();
// Ville d'une adresse SIGEO (mots après le dernier CP) / d'un texte libre.
function imVilleKey(adr, isTarget) {
  const tok = imTokens(adr);
  if (!isTarget) return tok.join(" ");
  let k = tok.length - 1;
  while (k >= 0 && !/^\d{5}$/.test(tok[k])) k--;
  return k >= 0 ? tok.slice(k + 1).filter((x) => x !== "CEDEX").join(" ") : "";
}

/**
 * Choix d'un immeuble dans les résultats d'une recherche par désignation
 * (le champ Nom est un « contient » : « LE BOURG » ramène aussi « LE BOURG
 * NORD », et plusieurs immeubles peuvent s'appeler « LE BOURG »).
 *   - exactement 1 Nom E.I. identique → retenu (CP différent → à vérifier) ;
 *   - sinon (0 ou plusieurs) : départage par le CP, puis la ville, comparés
 *     à la colonne Adresse — parmi les noms identiques s'il y en a, sinon
 *     parmi tous les résultats (alors toujours « à vérifier ») ;
 *   - sinon rien : JAMAIS la première ligne.
 * opts = { cp, ville }. Renvoie { row|null, exact, by, check, candidates }.
 */
function imPickByName(rows, nom, opts = {}) {
  const K = imStrictKey(nom);
  const seen = new Set();
  const uniq = (rows || []).filter((r) => (seen.has(r.im) ? false : (seen.add(r.im), true)));
  const exact = uniq.filter((r) => imStrictKey(r.nom) === K);
  const cp = String(opts.cp || "").trim();
  const ville = imVilleKey(opts.ville || "", false);
  const cpOf = (r) => imParseTarget(r.adresse).cp;
  if (exact.length === 1) {
    const r = exact[0];
    const check = cp && cpOf(r) && cpOf(r) !== cp ? `CP de l'immeuble SIGEO ${cpOf(r)} ≠ CP Intralia ${cp}` : "";
    return { row: r, exact: 1, by: "", check, candidates: exact };
  }
  const pool = exact.length ? exact : uniq;
  let left = pool;
  const by = [];
  if (cp && pool.length) { left = left.filter((r) => cpOf(r) === cp); by.push("CP " + cp); }
  if (ville && pool.length && (!cp || left.length > 1)) { left = left.filter((r) => imVilleKey(r.adresse, true) === ville); by.push("ville " + opts.ville); }
  if (by.length && left.length === 1) {
    const r = left[0];
    return {
      row: r, exact: exact.length, by: by.join(" + "), candidates: left,
      check: exact.length ? "" : `Nom E.I. « ${r.nom} » ≠ « ${nom} » : immeuble retenu par ${by.join(" + ")} seulement`
    };
  }
  return { row: null, exact: exact.length, by: by.join(" + "), check: "", candidates: by.length ? left : pool };
}

/**
 * Recherche « par désignation + site » d'UN immeuble : Nom = nom de
 * l'immeuble, Site = cfg.site (vide = tous les sites), tous les autres
 * critères vidés ; puis imPickByName.
 * cfg = { nom, site, cp, ville, ids, sel, url, timeoutMs, maxPages }
 * Renvoie { status: OK | AMBIGU | INTROUVABLE | ERREUR, im, id, row, check,
 *           nom, site, total, exact, by, candidates, error?, siteMissing?, auth? }.
 */
async function imFindByDesignation(ad, cfg) {
  const nom = String(cfg.nom ?? "").replace(/[   ]/g, " ").replace(/\s+/g, " ").trim();
  const site = String(cfg.site ?? "").trim();
  if (!nom) return { status: "ERREUR", nom, site, error: "nom d'immeuble vide" };
  const r = await imCachedSearch(ad, {
    mode: "nom", term: nom, cp: "", siteValue: site || undefined, siteEmpty: !site,
    ids: cfg.ids || IM_GER_IDS, clearAll: true, changeEvents: true,
    sel: cfg.sel || {}, url: cfg.url, timeoutMs: cfg.timeoutMs, maxPages: cfg.maxPages
  });
  if (!r.ok) return { status: "ERREUR", nom, site, error: r.error, siteMissing: !!r.siteMissing, auth: /AUTH_REQUIRED/.test(r.error || "") };
  const pick = imPickByName(r.rows, nom, { cp: cfg.cp, ville: cfg.ville });
  const base = {
    nom, site, total: r.total ?? r.rows.length, exact: pick.exact, by: pick.by, candidates: pick.candidates,
    cached: !!r.cached, warnings: r.warnings || []
  };
  if (pick.row) return { ...base, status: "OK", im: pick.row.im, id: pick.row.id || "", row: pick.row, check: pick.check };
  return { ...base, status: pick.candidates.length > 1 ? "AMBIGU" : "INTROUVABLE" };
}

// Résumé lisible d'un résultat imFindByDesignation (journal, colonne Méthode).
function imDesignationText(r) {
  if (r.status === "ERREUR") return `« ${r.nom} » : ${r.error}`;
  const n = (k, one, many) => `${k} ${k > 1 ? many : one}`;
  let t = `« ${r.nom} » : ${n(r.total || 0, "résultat", "résultats")}, ${n(r.exact || 0, "nom identique", "noms identiques")}`;
  if (r.by) t += `, ${r.status === "OK" ? "retenu par" : "puis"} ${r.by}`;
  if (r.status === "AMBIGU") {
    t += ` → ${r.candidates.length} candidats : ` + r.candidates.slice(0, 5).map((c) => `${c.im} ${c.nom} (${c.adresse})`).join(" ; ") +
      (r.candidates.length > 5 ? " ; …" : "");
  }
  return t;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    IM_GER_IDS, IM_DEFAULT_SITE, imStrictKey, imPickByName, imFindByDesignation, imDesignationText,
    imFold, imTokens, imParseSource, imParseTarget, imPickSearchTerm, imMatchAddress,
    imResolve, imNameKey, imStatusText, imPageOp, imLocate, imWaitForSearchDone, imSearchAllPages,
    imCachedSearch, imClearCache, imLookupAddress, imSearchCache
  };
}
