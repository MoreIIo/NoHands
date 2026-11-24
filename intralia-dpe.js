/* ====================================================================
 * NoHands OSA — DPE le plus récent par lot (Intralia → FlowerDocs)
 * --------------------------------------------------------------------
 * Onglet Extraction, type « DPE Intralia ».
 *
 * Pour chaque ligne (nom d'immeuble + VRAI n° de lot) :
 *  1. Index des immeubles (une fois par traitement, mis en cache) :
 *       GET /gerance/diag.php?action=R&ID_SOCIETE={id}&ID_GEST=0&TYPEB=0&DPE_TOUS=checked&GES=0
 *     HTML rendu côté serveur (DataTables côté client, pas d'AJAX), encodé
 *     en ISO-8859-1 → décodé en windows-1252 (jamais res.text()). Parsing
 *     par regex des lignes de #example3 (91 Mo pour « Tous les cabinets ») :
 *     td 0 Cabinet, 1 N° Immeuble, 2 Nom, 3 Adresse, 4 CP, 5 Ville, 6 N° Lot,
 *     7 type ; ids lus dans le lien diaglot.php?id_societe&id_immeuble&id_lot.
 *     Clé « NOM_NORMALISÉ|LOT » ; plusieurs correspondances = AMBIGU.
 *  2. Page du lot : GET /gerance/diaglot.php?id_societe&id_immeuble&id_lot
 *     → DPE déjà saisi (form1, lu, JAMAIS soumis) + documents GED
 *     (input.IDGED, jamais cochés) : nom, date « crée le », UUID.
 *  3. DPE le plus récent (filtre regex sur le nom, tri par date « crée le »
 *     décroissante, à égalité le premier de la page, doublons nom+date
 *     écartés) ; téléchargement FlowerDocs :
 *       GET /flower-docs-gui/download?documentId={UUID}&index=0
 *     contrôlé (statut HTTP + « %PDF ») avant tout enregistrement.
 *
 * LECTURE SEULE sur Intralia : uniquement des GET.
 *
 * Chargé par sidepanel.html AVANT sidepanel.js ; importable sous Node
 * (tests). intraPageOp est injectée telle quelle dans un onglet Intralia
 * (même origine → cookies de session) : elle doit rester autonome.
 * ==================================================================== */

const INTRA_DEFAULT_BASE = "https://intralia.oralia.local";
const INTRA_DEFAULT_FLOWER = "https://www.oralia.flowerdocs.cloud";
const INTRA_DEFAULT_SOC = "19";      // Sicov
const INTRA_ALL_SOC = "100";         // « Tous les cabinets »
const INTRA_DEFAULT_CABINETS = Object.freeze([
  { id: "19", label: "Sicov" },
  { id: "100", label: "Tous les cabinets" }
]);
// Appliquée au nom du fichier GED normalisé (sans accents, minuscules).
const INTRA_DEFAULT_REGEX = "dpe|diag.*perf|perf.*energ|(^|[_\\s.-])dp([_\\s.-]|$)";
const INTRA_DEFAULT_NAME = "{Immeuble}_lot{lots}_{DPE_Date_GED}.pdf";
const INTRA_DEFAULT_FOLDER = "DPE";
const INTRA_DEFAULT_DELAY = 400;     // pause entre lots : 300–500 ms (± 25 %)
const INTRA_DEFAULT_MAX_AGE_DAYS = 7;

const INTRA_STATUS = Object.freeze({
  OK: "OK",
  AUCUN_DPE_GED: "AUCUN_DPE_GED",
  LOT_INTROUVABLE: "LOT_INTROUVABLE",
  IMMEUBLE_INTROUVABLE: "IMMEUBLE_INTROUVABLE",
  AMBIGU: "AMBIGU",
  INTRALIA_NON_CONNECTE: "INTRALIA_NON_CONNECTE",
  FLOWERDOCS_NON_CONNECTE: "FLOWERDOCS_NON_CONNECTE",
  ERREUR: "ERREUR"
});

// Option « analyser le PDF » : colonnes de l'Analyse DPE (DS_COLS de
// dpe-scan.js : N° FU, N° ADEME, Établi le…, lues par l'étape « SIGEO :
// saisir un DPE ») + le nom du fournisseur.
const INTRA_SUPPLIER_COL = "Fournisseur DPE";

// Fichier final « prêt pour la saisie SIGEO » (v2.19.0) : colonnes lues telles
// quelles par l'étape « SIGEO : saisir un DPE » (préréglage DIAG_PRESETS.intralia
// de sigeo-diag.js, choisi tout seul par « Modèle DPE »).
const INTRA_SAISIE_COLS = Object.freeze([
  "DPE analyse", "URL fiche DPE", "Type de local", "PDF analysé", "N° FU", INTRA_SUPPLIER_COL,
  "Établi le", "Valable jusqu'au", "N° ADEME", "Énergie primaire (kWh/m²/an)", "Classe énergie",
  "GES (kg CO2/m²/an)", "Classe GES", "Coût min (€/an)", "Coût max (€/an)", "Date d'indexation des prix",
  "Remarques DPE"
]);
const INTRA_ID_COL = "ID interne SIGEO";
const INTRA_IM_COL = "IM SIGEO";
// Popup « Diagnostic » SIGEO d'un lot : {IM} = n° d'immeuble, {ID} = id interne du lot.
const INTRA_FICHE_DEFAULT = "http://sigeo.evoriel.net/popup.aspx/fr/sig_imm/diag_manage/{IM}/{ID}/0";

function intraFicheUrl(tpl, im, id) {
  return String(tpl || INTRA_FICHE_DEFAULT).trim().replace(/\{IM\}/gi, String(im)).replace(/\{ID\}/gi, String(id));
}

// Adresse de l'index Intralia (« voie CP ville », voie parfois vide) →
// { street, cp, ville } ; le CP est le dernier nombre à 5 chiffres.
function intraSplitAdr(adr) {
  const s = String(adr ?? "").replace(/\s+/g, " ").trim();
  const m = s.match(/^(.*)\b(\d{5})\b(.*)$/);
  if (!m) return { street: s, cp: "", ville: "" };
  return { street: m[1].trim(), cp: m[2], ville: m[3].trim() };
}

// Type de lot Intralia (« Appartement », « Commerce »…) → type de local SIGEO.
function intraTypeLocal(t) {
  const f = intraFold(t);
  if (!f) return { value: "", known: false };
  if (/appart|maison|studio|logement|habitation|duplex|triplex|pavillon|chambre|villa|loft|(^|\W)[tf]\d/.test(f)) return { value: "habitation", known: true };
  if (/commerc|bureau|local|activit|entrepot|boutique|professionnel|tertiaire|atelier|hotel|restaurant|magasin/.test(f)) return { value: "hors habitation", known: true };
  return { value: "", known: false };
}

// Colonnes toujours écrites (créées si absentes).
const INTRA_COLS = Object.freeze([
  "Intralia_Cabinet", "Intralia_ID_Immeuble", "Intralia_URL_Lot",
  "DPE_Statut", "DPE_Detail", "DPE_Fichier", "DPE_Date_GED", "DPE_UUID",
  "DPE_Lien_FlowerDocs", "DPE_Fichier_Local"
]);

// DPE déjà saisi dans Intralia (form1 de diaglot.php) → colonnes à cocher.
const INTRA_FIELDS = Object.freeze([
  { k: "libdpe1", col: "Intralia_DPE_Classe", label: "lettre DPE", def: true },
  { k: "DPE1", col: "Intralia_DPE_kWh_m2", label: "kWh/m²", def: true },
  { k: "libges1", col: "Intralia_GES_Classe", label: "lettre GES", def: true },
  { k: "GES1", col: "Intralia_GES_kgCO2_m2", label: "kgCO2/m²", def: true },
  { k: "typediag", col: "Intralia_Type_DPE", label: "type (avant / après le 01/07/2021)", def: false },
  { k: "DATEDIAG", col: "Intralia_Date_Diag", label: "date du diagnostic (DATEDIAG)", def: true },
  { k: "datefinval", col: "Intralia_Fin_Validite", label: "fin de validité", def: false },
  { k: "MTDIAGMIN", col: "Intralia_Cout_Min", label: "coût min", def: false },
  { k: "MTDIAGMAX", col: "Intralia_Cout_Max", label: "coût max", def: false },
  { k: "MTDIAG", col: "Intralia_Cout", label: "coût estimé (MTDIAG)", def: false },
  { k: "ANNEREFDIAG", col: "Intralia_Annee_Ref", label: "année de référence des prix", def: false },
  { k: "REFERENCE", col: "Intralia_Reference", label: "référence (ex. 1032/5004/0732)", def: false }
]);

/* ---------------------------------------------------------------- *
 * 1. Logique pure                                                  *
 * ---------------------------------------------------------------- */

// Nom d'immeuble : sans accents, majuscules, tout non-alphanumérique →
// espace, espaces réduits. « Île-Verte — Le Concerto 1 » → « ILE VERTE LE CONCERTO 1 ».
function intraNormName(s) {
  return String(s ?? "")
    .replace(/[Œœ]/g, "OE").replace(/[Ææ]/g, "AE")
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ").trim();
}

// N° de lot : trim + zéros de tête retirés (FlowerDocs « 0732 », Intralia « 732 »).
function intraNormLot(s) {
  return String(s ?? "").replace(/[\s ]+/g, "").toUpperCase().replace(/^0+(?=.)/, "");
}

// Nom de fichier GED pour le filtre : sans accents, minuscules.
function intraFold(s) {
  return String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[  ]/g, " ").toLowerCase().trim();
}

const intraTrimSlash = (u) => String(u ?? "").trim().replace(/\/+$/, "");

function intraIndexUrl(base, soc) {
  return `${intraTrimSlash(base)}/gerance/diag.php?action=R&ID_SOCIETE=${encodeURIComponent(soc)}&ID_GEST=0&TYPEB=0&DPE_TOUS=checked&GES=0`;
}
function intraFormUrl(base) {
  return `${intraTrimSlash(base)}/gerance/diag.php`;
}
function intraLotUrl(base, soc, imm, lot) {
  return `${intraTrimSlash(base)}/gerance/diaglot.php?id_societe=${encodeURIComponent(soc)}&id_immeuble=${encodeURIComponent(imm)}&id_lot=${encodeURIComponent(lot)}`;
}
function intraFlowerViewUrl(flower, uuid) {
  return `${intraTrimSlash(flower)}/flower-docs-gui/#/documents/edit:${uuid}`;
}
function intraFlowerDownloadUrl(flower, uuid) {
  return `${intraTrimSlash(flower)}/flower-docs-gui/download?documentId=${encodeURIComponent(uuid)}&index=0`;
}

// Lignes compactes de l'index (renvoyées par intraPageOp « index ») :
//   [soc, imm, lot, nom, cabinet, type, idLot, adresse]
// → { map: "NOM|LOT" → [entrée], names: NOM → { label, builds }, count }
function intraBuildIndex(rows) {
  const map = new Map();
  const names = new Map();
  let count = 0;
  for (const r of rows || []) {
    if (!r) continue;
    const [soc, imm, lot, nom, cabinet, type, idLot, adr] = r;
    const nn = intraNormName(nom), nl = intraNormLot(lot);
    if (!nn || !nl) continue;
    const e = {
      soc: String(soc ?? ""), imm: String(imm ?? ""), lot: String(lot ?? ""), idLot: String(idLot || lot || ""),
      nom: String(nom ?? ""), cabinet: String(cabinet ?? ""), type: String(type ?? ""), adr: String(adr ?? "")
    };
    const key = nn + "|" + nl;
    let list = map.get(key);
    if (!list) map.set(key, list = []);
    if (list.some((x) => x.soc === e.soc && x.imm === e.imm)) continue; // même lot listé deux fois
    list.push(e);
    count++;
    let nm = names.get(nn);
    if (!nm) names.set(nn, nm = { label: e.nom, builds: new Map() });
    const bk = e.soc + "|" + e.imm;
    let b = nm.builds.get(bk);
    if (!b) nm.builds.set(bk, b = { soc: e.soc, imm: e.imm, cabinet: e.cabinet, nom: e.nom, adr: e.adr, lots: 0 });
    b.lots++;
  }
  return { map, names, count };
}

// Similarité de deux noms normalisés (bigrammes, Dice) + bonus d'inclusion.
function intraSimilarity(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const grams = (s) => {
    const t = " " + s + " ", m = new Map();
    for (let i = 0; i < t.length - 1; i++) { const g = t.slice(i, i + 2); m.set(g, (m.get(g) || 0) + 1); }
    return m;
  };
  const A = grams(a), B = grams(b);
  let inter = 0, na = 0, nb = 0;
  A.forEach((v) => { na += v; });
  B.forEach((v, g) => { nb += v; inter += Math.min(v, A.get(g) || 0); });
  let s = (2 * inter) / (na + nb);
  if ((" " + a + " ").includes(" " + b + " ") || (" " + b + " ").includes(" " + a + " ")) s += 0.25;
  return s;
}

// Noms d'immeubles les plus proches (au plus n, score ≥ 0,35).
function intraCloseNames(idx, name, n = 3) {
  const nn = intraNormName(name);
  if (!nn || !idx || !idx.names) return [];
  const scored = [];
  idx.names.forEach((v, k) => {
    const s = intraSimilarity(nn, k);
    if (s >= 0.35) scored.push({ s, label: v.label });
  });
  return scored.sort((x, y) => y.s - x.s || x.label.localeCompare(y.label)).slice(0, n).map((x) => x.label);
}

// → { status: FOUND | AMBIGU | LOT_INTROUVABLE | IMMEUBLE_INTROUVABLE, matches, buildings?, close? }
function intraLookup(idx, nom, lot) {
  const nn = intraNormName(nom), nl = intraNormLot(lot);
  if (!nn) return { status: INTRA_STATUS.IMMEUBLE_INTROUVABLE, matches: [], close: [] };
  const matches = (idx && idx.map.get(nn + "|" + nl)) || [];
  if (matches.length === 1) return { status: "FOUND", matches };
  if (matches.length > 1) return { status: INTRA_STATUS.AMBIGU, matches };
  const nm = idx && idx.names.get(nn);
  if (nm) return { status: INTRA_STATUS.LOT_INTROUVABLE, matches: [], buildings: Array.from(nm.builds.values()) };
  return { status: INTRA_STATUS.IMMEUBLE_INTROUVABLE, matches: [], close: intraCloseNames(idx, nn, 3) };
}

const intraDescribe = (m) => `${m.cabinet || "cabinet " + m.soc} — imm ${m.imm} ${m.nom}${m.adr ? " (" + m.adr + ")" : ""}`;

// Regex de détection (saisie utilisateur) ; jamais d'exception.
function intraCompileRegex(src) {
  const s = String(src ?? "").trim() || INTRA_DEFAULT_REGEX;
  try { return { re: new RegExp(s, "i"), source: s, error: "" }; }
  catch (e) { return { re: null, source: s, error: `regex invalide « ${s} » (${e.message})` }; }
}

function intraIsDpeName(name, re) {
  return !!re && re.test(intraFold(name));
}

// docs : [{ uuid, name, date (AAAA-MM-JJ), order }] (ordre de la page)
// → { total, candidates (dédoublonnés, ordre de la page), sorted (plus récent d'abord), chosen }
function intraSelectDpe(docs, re) {
  const seen = new Set();
  const candidates = [];
  (docs || []).forEach((d, i) => {
    if (!intraIsDpeName(d.name, re)) return;
    const k = intraFold(d.name) + "|" + (d.date || "");
    if (seen.has(k)) return; // même nom + même date déposés deux fois
    seen.add(k);
    candidates.push(Object.assign({}, d, { order: d.order ?? i }));
  });
  const sorted = candidates.slice().sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")) || (a.order - b.order));
  return { total: (docs || []).length, candidates, sorted, chosen: sorted[0] || null };
}

// Nom de fichier Windows : caractères interdits → « _ », espaces réduits,
// pas de point / espace final, noms réservés (CON, NUL, COM1…) préfixés.
function intraSanitize(s) {
  let t = String(s ?? "")
    .replace(/[\\/:*?"<>|\u0000-\u001F]/g, "_")
    .replace(/[\s ]+/g, " ")
    .trim()
    .replace(/[. ]+$/g, "");
  if (/^(con|prn|aux|nul|com\d|lpt\d)(\.|$)/i.test(t)) t = "_" + t;
  return t;
}

function intraFolder(f) {
  return String(f ?? "").split(/[\\/]+/).map(intraSanitize).filter((p) => p && p !== "." && p !== "..").join("/");
}

const INTRA_KNOWN_EXT = /\.(pdf|docx?|odt|rtf|txt|jpe?g|png|tiff?|gif|bmp|msg|eml|xlsx?|zip)$/i;

function intraExtOf(name) {
  const m = String(name ?? "").trim().match(INTRA_KNOWN_EXT);
  return m ? "." + m[1].toLowerCase() : "";
}

// tpl : « {Immeuble}_lot{lots}_{DPE_Date_GED}.pdf ». Variables de sortie
// (vars, insensibles à la casse) d'abord, puis {Colonne} de la ligne
// (resolve = resolveRowTemplate), variables inconnues retirées. L'extension
// est celle du fichier réellement reçu (ext), sinon celle du modèle, sinon .pdf.
function intraFileName(tpl, vars, resolve, ext) {
  const lower = {};
  Object.keys(vars || {}).forEach((k) => { lower[k.toLowerCase()] = vars[k]; });
  let s = String(tpl || INTRA_DEFAULT_NAME).replace(/\{([^{}]+)\}/g, (w, k) => {
    const v = lower[k.trim().toLowerCase()];
    return v === undefined || v === null ? w : String(v);
  });
  if (resolve) s = resolve(s);
  s = s.replace(/\{[^{}]*\}/g, "");
  const tplExt = intraExtOf(s);
  if (tplExt) s = s.slice(0, -tplExt.length);
  const base = intraSanitize(s) || "DPE";
  return base + (ext || tplExt || ".pdf");
}

// Réponse du téléchargement FlowerDocs → { kind, ext?, mime?, error? }
//   PDF      : commence par %PDF → enregistrable
//   LOGIN    : 401/403, page HTML (connexion) → FLOWERDOCS_NON_CONNECTE
//   HTTP     : autre statut / fichier vide → ERREUR (la boucle continue)
//   NOT_PDF  : fichier GED « .pdf » dont le contenu n'est pas un PDF → ERREUR
//   OTHER    : fichier GED d'un autre type (.doc, .jpg…) → enregistré tel quel
function intraClassifyFile(r, gedName) {
  if (!r) return { kind: "HTTP", error: "aucune réponse" };
  const bytes = r.bytes || new Uint8Array(0);
  const head = String.fromCharCode.apply(null, Array.from(bytes.subarray(0, 512)));
  if (head.startsWith("%PDF")) return { kind: "PDF", ext: ".pdf", mime: "application/pdf" };
  const status = Number(r.status) || 0;
  if (status === 401 || status === 403) return { kind: "LOGIN", error: `FlowerDocs a refusé l'accès (HTTP ${status})` };
  const html = /text\/html/i.test(r.contentType || "") || /^\s*</.test(head);
  if (status >= 200 && status < 300 && html) return { kind: "LOGIN", error: "FlowerDocs a renvoyé une page HTML (page de connexion ?)" };
  if (status < 200 || status >= 300) return { kind: "HTTP", error: `FlowerDocs : HTTP ${status || "?"}` };
  if (!bytes.length) return { kind: "HTTP", error: "FlowerDocs : fichier vide" };
  const dispo = /filename\*?=(?:utf-8'')?["']?([^;"']+)/i.exec(r.disposition || "");
  const ext = intraExtOf(gedName) || intraExtOf(dispo ? decodeURIComponent(dispo[1]) : "");
  if (!ext || ext === ".pdf") return { kind: "NOT_PDF", error: "le fichier reçu n'est pas un PDF (" + (r.contentType || "type inconnu") + ")" };
  return { kind: "OTHER", ext, mime: r.contentType || "application/octet-stream" };
}

// Configuration complète (valeurs par défaut, migrations).
// v2 (v2.20.0, « DPE extraction » tout en un) : une config d'avant passe en
// téléchargement + analyse + fichier prêt pour la saisie.
function intraNormalizeConfig(c = {}) {
  const old = !(Number(c.v) >= 2);
  const num = (v, d) => { const n = parseInt(v, 10); return Number.isFinite(n) && n >= 0 ? n : d; };
  const fields = Array.isArray(c.fields)
    ? c.fields.filter((k) => INTRA_FIELDS.some((f) => f.k === k))
    : INTRA_FIELDS.filter((f) => f.def).map((f) => f.k);
  return {
    v: 2,
    nomCol: String(c.nomCol ?? ""),
    lotCol: String(c.lotCol ?? ""),
    soc: String(c.soc ?? "").trim() || INTRA_DEFAULT_SOC,
    regex: String(c.regex ?? "").trim() || INTRA_DEFAULT_REGEX,
    mode: old || c.mode !== "find" ? "download" : "find",
    folder: c.folder === undefined || c.folder === null ? INTRA_DEFAULT_FOLDER : String(c.folder),
    name: String(c.name ?? "").trim() || INTRA_DEFAULT_NAME,
    verify: !!c.verify,
    toolbox: !!c.toolbox,
    analyze: old || c.analyze !== false,
    sigeo: c.sigeo !== false,
    imCol: String(c.imCol ?? ""),
    idCol: String(c.idCol ?? ""),
    fiche: String(c.fiche ?? "").trim() || INTRA_FICHE_DEFAULT,
    out: !old && c.out === "all" ? "all" : "saisie",
    ocr: ["auto", "always", "never"].includes(c.ocr) ? c.ocr : "auto",
    fields,
    base: intraTrimSlash(c.base) || INTRA_DEFAULT_BASE,
    flower: intraTrimSlash(c.flower) || INTRA_DEFAULT_FLOWER,
    delay: num(c.delay, INTRA_DEFAULT_DELAY),
    maxAgeDays: num(c.maxAgeDays, INTRA_DEFAULT_MAX_AGE_DAYS)
  };
}

// Valeurs à écrire dans la ligne pour un résultat (colonnes → texte).
function intraRowValues(res, fieldKeys) {
  const c = res.chosen || null;
  const v = {
    Intralia_Cabinet: res.cabinet || "",
    Intralia_ID_Immeuble: res.imm || "",
    Intralia_URL_Lot: res.lotUrl || "",
    DPE_Statut: res.status || "",
    DPE_Detail: res.detail || "",
    DPE_Fichier: c ? c.name : "",
    DPE_Date_GED: c ? c.date : "",
    DPE_UUID: c ? c.uuid : "",
    DPE_Lien_FlowerDocs: c ? c.view || "" : "",
    DPE_Fichier_Local: res.file || ""
  };
  for (const k of fieldKeys || []) {
    const f = INTRA_FIELDS.find((x) => x.k === k);
    if (f) v[f.col] = res.fields ? String(res.fields[k] ?? "") : "";
  }
  return v;
}

// Colonnes « saisie SIGEO » (mêmes noms que l'Analyse DPE, lues par l'étape
// « SIGEO : saisir un DPE ») pour un résultat :
//  - données du DPE saisies dans INTRALIA (form1) : établi le (DATEDIAG),
//    énergie primaire, classe énergie, GES, classe GES, coûts (DPE ≥ 01/07/2021) ;
//    « Valable jusqu'au » calculé comme SIGEO (ds.validity) — la datefinval
//    d'Intralia n'est pas fiable ; valeur du PDF seulement si Intralia est vide ;
//  - lus dans le PDF (texte / OCR) : fournisseur → N° FU, N° ADEME.
// ds = { DS_COLS, dsRecordValues, validity(date) }. Aucun DPE → « Pas de PDF ».
function intraSaisieValues(res, ds) {
  const v = {};
  INTRA_SAISIE_COLS.forEach((c) => { v[c] = ""; });
  const S = INTRA_STATUS;
  const why = `${res.status || ""}${res.detail ? " — " + res.detail : ""}`;
  if ([S.AUCUN_DPE_GED, S.LOT_INTROUVABLE, S.IMMEUBLE_INTROUVABLE, S.AMBIGU].includes(res.status)) {
    v["DPE analyse"] = "Pas de PDF";
    v["Remarques DPE"] = why;
    return v;
  }
  if (res.status !== S.OK || !res.chosen) {
    if (res.status) { v["DPE analyse"] = "ÉCHEC"; v["Remarques DPE"] = why; }
    return v;
  }
  const f = res.fields || {};
  const a = res.analysis || null;
  const pdf = {};
  if (a && !a.error) {
    const vals = ds.dsRecordValues(a, "");
    ds.DS_COLS.forEach((c, k) => { pdf[c] = vals[k] ?? ""; });
  }
  const rem = [];
  // Vide, « 0 », « 0,00 »… = pas d'info dans Intralia → valeur du scan.
  const num = (x) => {
    const t = String(x ?? "").trim().replace(/\s+/g, "").replace(".", ",");
    return /^0*(,0*)?$/.test(t) ? "" : t;
  };
  const date = (x) => {
    const t = String(x ?? "").trim();
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(t);
    return m ? `${m[3]}/${m[2]}/${m[1]}` : /^0+\/0+\/0+$/.test(t) ? "" : t;
  };
  const pick = (col, intraVal, label) => {
    if (intraVal) return intraVal;
    if (pdf[col]) { rem.push(`${label} vide dans Intralia : valeur du PDF`); return pdf[col]; }
    rem.push(`${label} non renseigné`);
    return "";
  };
  v["Établi le"] = pick("Établi le", date(f.DATEDIAG), "date du DPE");
  v["Valable jusqu'au"] = (ds.validity && ds.validity(v["Établi le"])) || pdf["Valable jusqu'au"] || "";
  v["Énergie primaire (kWh/m²/an)"] = pick("Énergie primaire (kWh/m²/an)", num(f.DPE1), "énergie primaire");
  v["Classe énergie"] = pick("Classe énergie", f.libdpe1, "classe énergie");
  v["GES (kg CO2/m²/an)"] = pick("GES (kg CO2/m²/an)", num(f.GES1), "GES");
  v["Classe GES"] = pick("Classe GES", f.libges1, "classe GES");
  const dm = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(v["Établi le"]);
  const isNew = !!dm && (+dm[3] > 2021 || (+dm[3] === 2021 && +dm[2] >= 7));
  if (isNew) {
    v["Coût min (€/an)"] = pick("Coût min (€/an)", num(f.MTDIAGMIN), "coût min");
    v["Coût max (€/an)"] = pick("Coût max (€/an)", num(f.MTDIAGMAX), "coût max");
    // Date d'indexation des prix : année de référence d'Intralia, sinon le PDF.
    const yr = String(f.ANNEREFDIAG ?? "").trim();
    const pr = a && !a.error && a.prixRef;
    if (/^(19|20)\d{2}$/.test(yr)) v["Date d'indexation des prix"] = `01/01/${yr}`;
    else if (pr) {
      v["Date d'indexation des prix"] = `${String(pr.d).padStart(2, "0")}/${String(pr.m).padStart(2, "0")}/${pr.y}`;
      rem.push("année de référence vide dans Intralia : date d'indexation du PDF");
    } else rem.push("date d'indexation des prix non renseignée");
  }
  // Type de local (type du lot dans l'index Intralia).
  const tl = intraTypeLocal(res.lotType);
  v["Type de local"] = tl.value;
  if (!tl.known) rem.push(res.lotType ? `type de lot « ${res.lotType} » : type de local à vérifier (habitation par défaut)` : "type de lot inconnu : type de local à vérifier");
  // Fiche SIGEO (id interne du lot) : res.sigeo = { status, url, text }.
  const sg = res.sigeo;
  if (sg) {
    if (sg.url) v["URL fiche DPE"] = sg.url;
    else rem.push(`fiche SIGEO introuvable : ${sg.text || sg.status}`);
  }
  if (!a) rem.push("PDF non analysé : fournisseur et n° ADEME à saisir");
  else if (a.error) rem.push(`analyse du PDF impossible (${a.error}) : fournisseur et n° ADEME à saisir`);
  else {
    v["N° FU"] = pdf["N° FU"] || "";
    v[INTRA_SUPPLIER_COL] = a.supplier ? a.supplier.entry.name : (a.societe || "");
    v["N° ADEME"] = pdf["N° ADEME"] || "";
    // Remarques du PDF utiles ici : fournisseur, et n° ADEME pour un DPE ≥ 01/07/2021
    // (avant, SIGEO n'a pas de champ N° ADEME en création).
    (a.remarques || []).filter((r) => /fournisseur|n° FU/i.test(r) || (isNew && /ADEME/i.test(r))).forEach((r) => rem.push(r));
    if (pdf["Établi le"] && v["Établi le"] && pdf["Établi le"] !== v["Établi le"]) {
      rem.push(`PDF établi le ${pdf["Établi le"]} ≠ Intralia ${v["Établi le"]} : vérifier que le PDF est bien ce DPE`);
    }
  }
  v["PDF analysé"] = res.file || res.chosen.name;
  v["DPE analyse"] = rem.length ? "À VÉRIFIER" : "OK";
  v["Remarques DPE"] = rem.join(" ; ");
  return v;
}

// Colonnes écrites en mode « utiles à la saisie ».
function intraSaisieCols() {
  return INTRA_SAISIE_COLS.slice();
}

// Variables du modèle de nom de fichier pour le DPE retenu c.
function intraTemplateVars(c, res) {
  const v = intraRowValues(Object.assign({}, res || {}, { chosen: c, file: "" }), INTRA_FIELDS.map((f) => f.k));
  v.DPE_Fichier = String((c && c.name) || "").replace(INTRA_KNOWN_EXT, "");
  v.DPE_Date_GED = (c && c.date) || "sans-date";
  delete v.DPE_Fichier_Local; delete v.DPE_Statut; delete v.DPE_Detail;
  return v;
}

/* ---------------------------------------------------------------- *
 * 2. Fonction injectée dans un onglet Intralia / FlowerDocs        *
 * ---------------------------------------------------------------- *
 * cfg.op :
 *   "index"    { url }       → GET diag.php → { ok, login, rows, cabinets, size }
 *   "cabinets" { url }       → GET diag.php (formulaire) → { ok, cabinets }
 *   "lot"      { url }       → GET diaglot.php → { ok, login, fields, docs }
 *   "file"     { url }       → GET (FlowerDocs) → { status, contentType, disposition, b64 }
 *   "parseIndex" / "parseCabinets" / "parseLot" { html, url } → idem sur un HTML (tests)
 * Uniquement des GET : rien n'est soumis, rien n'est coché. */
async function intraPageOp(cfg) {
  const ENT = {
    amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " ", laquo: "«", raquo: "»", deg: "°", euro: "€",
    rsquo: "’", lsquo: "‘", ldquo: "“", rdquo: "”", ndash: "–", mdash: "—", hellip: "…", sup2: "²", middot: "·",
    eacute: "é", egrave: "è", ecirc: "ê", euml: "ë", agrave: "à", aacute: "á", acirc: "â", auml: "ä", ccedil: "ç",
    icirc: "î", iuml: "ï", ocirc: "ô", ouml: "ö", ugrave: "ù", ucirc: "û", uuml: "ü", oelig: "œ", aelig: "æ",
    Eacute: "É", Egrave: "È", Ecirc: "Ê", Euml: "Ë", Agrave: "À", Acirc: "Â", Ccedil: "Ç", Icirc: "Î", Iuml: "Ï",
    Ocirc: "Ô", Ugrave: "Ù", Ucirc: "Û", OElig: "Œ", AElig: "Æ"
  };
  const decodeEnt = (s) => String(s ?? "").replace(/&(#[xX][0-9a-fA-F]+|#\d+|[a-zA-Z][a-zA-Z0-9]*);/g, (w, e) => {
    if (e[0] === "#") {
      const n = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      try { return String.fromCodePoint(n); } catch (_) { return w; }
    }
    return ENT[e] !== undefined ? ENT[e] : w;
  });
  const norm = (s) => String(s ?? "").replace(/[\s  ]+/g, " ").trim();
  const htmlText = (h) => norm(decodeEnt(String(h ?? "").replace(/<br\s*\/?>/gi, " ").replace(/<[^>]*>/g, " ")));
  // ISO-8859-1 / windows-1252 : jamais res.text() (supposerait UTF-8).
  const decode = (buf) => new TextDecoder("windows-1252").decode(buf);

  async function get(url) {
    const res = await fetch(url, { credentials: "include", cache: "no-store" });
    const buf = await res.arrayBuffer();
    return { status: res.status, url: res.url || url, redirected: !!res.redirected, html: decode(buf), size: buf.byteLength };
  }
  const looksLogin = (html, url) =>
    /<input\b[^>]*type\s*=\s*["']?password/i.test(html) || /login|connexion|authentif|identif/i.test(String(url || ""));

  function parseCabinets(html) {
    const sm = /<select\b[^>]*\bname\s*=\s*["']?ID_SOCIETE\b[^>]*>([\s\S]*?)<\/select\s*>/i.exec(html);
    if (!sm) return [];
    const out = [];
    const re = /<option\b([^>]*)>([\s\S]*?)(?=<option\b|$)/gi;
    let m;
    while ((m = re.exec(sm[1]))) {
      const vm = /\bvalue\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/i.exec(m[1]);
      const label = htmlText(m[2].replace(/<\/option\s*>[\s\S]*$/i, ""));
      const id = norm(decodeEnt(vm ? (vm[1] ?? vm[2] ?? vm[3]) : label));
      if (!id || out.some((o) => o.id === id)) continue;
      out.push({ id, label: label || id, selected: /\bselected\b/i.test(m[1]) });
    }
    return out;
  }

  // Lignes de #example3 par regex (pas de DOMParser : jusqu'à 91 Mo).
  function parseIndex(html, url) {
    const out = { ok: false, login: false, url, rows: [], cabinets: parseCabinets(html), size: html.length };
    const tm = /<table\b[^>]*\bid\s*=\s*["']?example3\b[^>]*>/i.exec(html);
    if (!tm) { out.login = looksLogin(html, url); return out; }
    out.ok = true;
    let body = html.slice(tm.index + tm[0].length);
    const end = body.search(/<\/table\s*>/i);
    if (end >= 0) body = body.slice(0, end);
    const tb = body.search(/<tbody\b/i);
    if (tb >= 0) body = body.slice(tb);
    const trs = body.split(/<tr\b[^>]*>/i);
    for (let i = 1; i < trs.length; i++) {
      const tr = trs[i];
      const cells = tr.split(/<td\b[^>]*>/i).slice(1).map((c) => c.replace(/<\/td\s*>[\s\S]*$/i, ""));
      if (cells.length < 7) continue; // en-tête (th) ou ligne vide
      const c = cells.map(htmlText);
      const hm = /href\s*=\s*["']?([^"'\s>]*diaglot\.php\?[^"'\s>]*)/i.exec(tr);
      let soc = "", imm = "", idLot = "";
      if (hm) {
        const q = new URLSearchParams(decodeEnt(hm[1]).split("?")[1] || "");
        soc = q.get("id_societe") || ""; imm = q.get("id_immeuble") || ""; idLot = q.get("id_lot") || "";
      }
      if (!soc && cfg.soc && String(cfg.soc) !== "100") soc = String(cfg.soc);
      const lot = c[6] || idLot;
      const adr = [c[3], c[4], c[5]].filter(Boolean).join(" ");
      out.rows.push([soc, imm || c[1], lot, c[2], c[0], c[7] || "", idLot || lot, adr]);
    }
    return out;
  }

  function parseLot(html, url) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const form = doc.querySelector('form[name="form1"]');
    const out = { ok: !!form, login: false, url, fields: {}, docs: [] };
    if (!form) { out.login = looksLogin(html, url); return out; }

    // Champ de form1 (y compris hors de la balise si le HTML est mal imbriqué).
    const el = (n) => {
      let e = null;
      try { e = form.elements && form.elements.namedItem(n); } catch (_) { e = null; }
      if (e && !e.tagName && e.length) e = e[0];
      if (e) return e;
      e = doc.querySelector(`[name="${n}"]`);
      if (e) return e;
      const low = n.toLowerCase(); // nom écrit avec une autre casse
      return Array.from(doc.querySelectorAll("[name]")).find((x) => String(x.getAttribute("name")).toLowerCase() === low) || null;
    };
    // Valeur posée en JavaScript dans la page (document.form1.X.value = '…').
    const jsValue = (n) => {
      const m = new RegExp("\\b" + n + "\\b[^;\\n]{0,60}?(?:\\.value\\s*=|\\.val\\(|\\.selectedValue\\s*=)\\s*['\"]([^'\"]*)['\"]").exec(html);
      return m ? m[1] : null;
    };
    // Comme le navigateur : dernière option « selected », sinon la première.
    const pick = (sel) => {
      const opts = Array.from(sel.querySelectorAll("option"));
      const o = opts.filter((x) => x.hasAttribute("selected")).pop() || opts[0];
      return o ? { value: norm(o.getAttribute("value") ?? o.textContent), text: norm(o.textContent) } : { value: "", text: "" };
    };
    const read = (n) => {
      const e = el(n);
      const js = jsValue(n);
      if (!e) return js === null ? null : { value: norm(js), text: "" };
      if (e.tagName === "SELECT") {
        if (js !== null) {
          const o = Array.from(e.querySelectorAll("option")).find((x) => norm(x.getAttribute("value") ?? x.textContent) === norm(js));
          if (o) return { value: norm(js), text: norm(o.textContent) };
        }
        return pick(e);
      }
      if (e.tagName === "TEXTAREA") return { value: norm(e.textContent), text: "" };
      return { value: norm(js !== null && !e.getAttribute("value") ? js : (e.getAttribute("value") ?? "")), text: "" };
    };
    const letter = (v) => {
      if (!v) return "";
      if (/^[A-G]$/i.test(v.value)) return v.value.toUpperCase();
      const m = /^([A-G])\b/i.exec(v.text || "");
      if (m) return m[1].toUpperCase();
      const t = v.text || v.value;
      return /^(-+|\.+|0|choisir.*|s[ée]lection.*|aucun.*)$/i.test(t) ? "" : t;
    };
    const f = out.fields;
    f.libdpe1 = letter(read("libdpe1"));
    f.libges1 = letter(read("libges1"));
    const td = read("typediag");
    f.typediag = !td || td.value === "" ? "" : td.value === "0" ? "avant le 01/07/2021" : "après le 01/07/2021";
    ["DPE1", "GES1", "DATEDIAG", "datefinval", "MTDIAGMIN", "MTDIAGMAX", "MTDIAG", "ANNEREFDIAG", "REFERENCE"]
      .forEach((n) => { const v = read(n); f[n] = v ? v.value : ""; });
    // « Montant théorique des dépenses énergétiques : Entre [..] et [..] » :
    // cases lues par leur libellé (ce que l'on voit à l'écran), prioritaires
    // sur MTDIAGMIN / MTDIAGMAX.
    const lab = Array.from(doc.querySelectorAll("td, th, label, b, strong, span, div, p"))
      .find((x) => x.textContent.length < 120 && /montant\s+th[ée]orique\s+des\s+d[ée]penses/i.test(x.textContent) && !x.querySelector("input"));
    if (lab) {
      const scope = lab.closest("tr") || lab.parentElement;
      const ins = scope ? Array.from(scope.querySelectorAll("input")).filter((i) => !/^(hidden|checkbox|radio|submit|button|image)$/i.test(i.getAttribute("type") || "")) : [];
      const vals = ins.map((i) => norm(i.getAttribute("value") ?? ""));
      if (ins.length >= 2 && (vals[0] || vals[1])) { f.MTDIAGMIN = vals[0]; f.MTDIAGMAX = vals[1]; }
    }

    // Documents GED : chaque input.IDGED (jamais coché) → sa ligne.
    const seen = new Set();
    Array.from(doc.querySelectorAll('input.IDGED, input[name="IDGED"]')).forEach((inp) => {
      if (seen.has(inp)) return;
      seen.add(inp);
      const uuid = norm(inp.getAttribute("value") || "");
      const scope = inp.closest("tr") || inp.parentElement || doc.body;
      const labels = Array.from(scope.querySelectorAll("span.label")).map((s) => norm(s.textContent));
      const b = scope.querySelector("b");
      let name = b ? norm(b.textContent) : "";
      if (!name) {
        const tds = Array.from(scope.querySelectorAll("td"));
        const cell = tds.find((t) => !t.querySelector("input,a")) || tds[tds.length - 1];
        name = cell ? norm(cell.textContent.replace(new RegExp(labels.map((l) => l.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|") || "$^", "g"), "")) : "";
      }
      const lbl = labels.join(" | ");
      let date = "";
      const dm = /cr\S{0,4}\s+le\s+(\d{4}-\d{2}-\d{2})/i.exec(lbl) || /(\d{4}-\d{2}-\d{2})/.exec(lbl);
      if (dm) date = dm[1];
      else {
        const fm = /cr\S{0,4}\s+le\s+(\d{1,2})\/(\d{1,2})\/(\d{4})/i.exec(lbl);
        if (fm) date = `${fm[3]}-${fm[2].padStart(2, "0")}-${fm[1].padStart(2, "0")}`;
      }
      const a = scope.querySelector('a[href*="documents/edit"], a[target="GED"]');
      out.docs.push({
        uuid, name, date,
        kind: labels.filter((l) => !/cr\S{0,4}\s+le/i.test(l)).join(", "),
        link: a ? a.getAttribute("href") || "" : "",
        order: out.docs.length
      });
    });
    return out;
  }

  try {
    switch (cfg.op) {
      case "parseIndex": return parseIndex(cfg.html, cfg.url || "");
      case "parseCabinets": return { ok: true, cabinets: parseCabinets(cfg.html) };
      case "parseLot": return parseLot(cfg.html, cfg.url || "");
      case "index": {
        const r = await get(cfg.url);
        const out = parseIndex(r.html, r.url);
        return Object.assign(out, { status: r.status, size: r.size, redirected: r.redirected });
      }
      case "cabinets": {
        const r = await get(cfg.url);
        const cabinets = parseCabinets(r.html);
        return { ok: cabinets.length > 0, login: !cabinets.length && looksLogin(r.html, r.url), cabinets, status: r.status, url: r.url };
      }
      case "lot": {
        const r = await get(cfg.url);
        return Object.assign(parseLot(r.html, r.url), { status: r.status, redirected: r.redirected });
      }
      case "file": {
        const res = await fetch(cfg.url, { credentials: "include", cache: "no-store" });
        const u8 = new Uint8Array(await res.arrayBuffer());
        let s = "";
        for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
        return {
          status: res.status, url: res.url, redirected: !!res.redirected,
          contentType: res.headers.get("content-type") || "",
          disposition: res.headers.get("content-disposition") || "",
          size: u8.length, b64: btoa(s)
        };
      }
      default: return { ok: false, error: "opération inconnue : " + cfg.op };
    }
  } catch (e) {
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

/* ---------------------------------------------------------------- *
 * 3. Orchestration d'une ligne (côté panneau, via un adaptateur)   *
 * ---------------------------------------------------------------- *
 * ad = {
 *   intra(cfg):          résultat de intraPageOp dans l'onglet Intralia
 *   fetchFile(url):      { status, contentType, disposition, bytes, via } | { error }
 *   save(bytes, path, mime): { ok, filename, error? }  (chrome.downloads)
 *   verify(bytes):       { ok, how, why }  (« performance énergétique » p. 1)
 *   analyze(bytes, name): enregistrement de dsAnalyzeDoc (dpe-scan.js)
 *   stopped(), log(text, cls)
 * }
 * p = { idx, nom, lot, base, flower, re, download, verify, analyze, socLabel,
 *       name(chosen, ext, res) → chemin relatif sous Téléchargements }
 * Résultat : { status (INTRA_STATUS), detail, cabinet, soc, imm, lotUrl,
 *   fields, chosen, candidates, nDocs, file, analysis?, bytes?, info[], error? }
 * L'analyse (n° FU, n° ADEME…) ne change jamais DPE_Statut : un échec
 * d'analyse est rapporté dans res.analysis.error. */

async function intraGetFile(ad, p, c) {
  const url = intraFlowerDownloadUrl(p.flower, c.uuid);
  let r;
  try { r = await ad.fetchFile(url); } catch (e) { r = { error: (e && e.message) || String(e) }; }
  if (!r || r.error) return { kind: "LOGIN", error: "FlowerDocs injoignable (" + ((r && r.error) || "aucune réponse") + ")" };
  return Object.assign(intraClassifyFile(r, c.name), { bytes: r.bytes, via: r.via || "" });
}

async function intraProcessRow(ad, p) {
  const S = INTRA_STATUS;
  const info = [];
  const res = { status: S.ERREUR, detail: "", cabinet: "", soc: "", imm: "", lotUrl: "", fields: null,
    chosen: null, candidates: [], nDocs: 0, file: "", info };
  const done = (status, detail, extra) => Object.assign(res, { status, detail: detail || "" }, extra || {});

  const look = intraLookup(p.idx, p.nom, p.lot);
  if (look.status === S.AMBIGU) {
    return done(S.AMBIGU, `${look.matches.length} correspondances : ` + look.matches.map(intraDescribe).join(" ; "), { matches: look.matches });
  }
  if (look.status === S.LOT_INTROUVABLE) {
    const bs = look.buildings || [];
    const one = bs.length === 1 ? bs[0] : null;
    return done(S.LOT_INTROUVABLE,
      `immeuble trouvé (${bs.map((b) => `${b.cabinet || "cabinet " + b.soc} imm ${b.imm}, ${b.lots} lot(s)`).join(" ; ")}) mais pas le lot ${intraNormLot(p.lot)}`,
      one ? { cabinet: one.cabinet, soc: one.soc, imm: one.imm } : {});
  }
  if (look.status === S.IMMEUBLE_INTROUVABLE) {
    const close = look.close || [];
    return done(S.IMMEUBLE_INTROUVABLE, `nom absent de l'index ${p.socLabel || ""}`.trim() + (close.length ? ` — proches : ${close.join(" ; ")}` : ""), { close });
  }

  const m = look.matches[0];
  Object.assign(res, { cabinet: m.cabinet, soc: m.soc, imm: m.imm, lotType: m.type || "", adr: m.adr || "", nomIntralia: m.nom || "", lotUrl: intraLotUrl(p.base, m.soc, m.imm, m.idLot || m.lot) });
  if (ad.stopped()) return done(S.ERREUR, "arrêté", { error: "arrêté" });

  let page;
  try { page = await ad.intra({ op: "lot", url: res.lotUrl }); }
  catch (e) { return done(S.ERREUR, "page du lot : " + ((e && e.message) || e), { error: (e && e.message) || String(e) }); }
  if (!page || !page.ok) {
    if (page && page.login) return done(S.INTRALIA_NON_CONNECTE, "page de connexion Intralia (session expirée ?)", { error: "Intralia : non connecté" });
    const why = page && page.error ? page.error : `page du lot inattendue (HTTP ${page && page.status ? page.status : "?"}, pas de formulaire form1)`;
    return done(S.ERREUR, why, { error: why });
  }
  res.fields = page.fields || {};
  res.nDocs = (page.docs || []).length;
  const sel = intraSelectDpe(page.docs, p.re);
  const view = (c) => Object.assign({}, c, { view: c.link && c.link.includes(c.uuid) ? c.link : intraFlowerViewUrl(p.flower, c.uuid) });
  res.candidates = sel.sorted.map(view);
  if (!sel.sorted.length) {
    const names = (page.docs || []).map((d) => d.name).filter(Boolean);
    return done(S.AUCUN_DPE_GED, `${res.nDocs} document(s) GED, aucun DPE` +
      (names.length ? ` (${names.slice(0, 4).join(", ")}${names.length > 4 ? "…" : ""})` : ""));
  }

  let chosen = null, file = null;
  if (p.verify) {
    let rejected = 0;
    for (const c of res.candidates) {
      if (ad.stopped()) return done(S.ERREUR, "arrêté", { error: "arrêté" });
      const f = await intraGetFile(ad, p, c);
      if (f.kind === "LOGIN") return done(S.FLOWERDOCS_NON_CONNECTE, f.error, { error: f.error });
      if (f.kind !== "PDF") { rejected++; info.push(`« ${c.name} » (${c.date || "sans date"}) non vérifiable : ${f.error || "pas un PDF"}`); continue; }
      let v;
      try { v = await ad.verify(f.bytes); } catch (e) { v = { ok: false, why: "lecture impossible : " + ((e && e.message) || e) }; }
      if (v && v.ok) { chosen = c; file = f; info.push(`« ${c.name} » vérifié (${v.how || "texte"})`); break; }
      rejected++;
      info.push(`« ${c.name} » (${c.date || "sans date"}) écarté : ${(v && v.why) || "contenu non reconnu"}`);
    }
    if (!chosen) {
      return done(S.AUCUN_DPE_GED, `${res.nDocs} document(s) GED, ${rejected} candidat(s) écarté(s) par la vérification du contenu`);
    }
  } else {
    chosen = res.candidates[0];
  }
  res.chosen = chosen;
  const others = res.candidates.length - 1;
  const okDetail = `${res.candidates.length} DPE sur ${res.nDocs} document(s) GED` + (others > 0 ? ` (autres : ${res.candidates.filter((x) => x !== chosen).map((x) => `${x.name} ${x.date}`).join(" ; ")})` : "");
  // Le PDF n'est chargé que pour l'enregistrer ou l'analyser.
  if (!p.download && !p.analyze) return done(S.OK, okDetail);

  if (!file) {
    file = await intraGetFile(ad, p, chosen);
    if (file.kind === "LOGIN") return done(S.FLOWERDOCS_NON_CONNECTE, file.error, { error: file.error });
    if (file.kind !== "PDF" && file.kind !== "OTHER") {
      if (p.download) return done(S.ERREUR, file.error, { error: file.error });
      res.analysis = { error: file.error }; // repérage seul : DPE repéré, seule l'analyse échoue
      return done(S.OK, okDetail);
    }
  }
  if (file.via) res.via = file.via;
  if (p.download) {
    const path = p.name(chosen, file.ext, res);
    let sv;
    try { sv = await ad.save(file.bytes, path, file.mime); } catch (e) { sv = { ok: false, error: (e && e.message) || String(e) }; }
    if (!sv || !sv.ok) {
      const why = "enregistrement impossible : " + ((sv && sv.error) || "?");
      return done(S.ERREUR, why, { error: why });
    }
    res.file = sv.filename || path.split("/").pop();
    res.path = path;
    res.bytes = file.bytes;
  }
  if (p.analyze) {
    if (file.kind !== "PDF") {
      res.analysis = { error: `fichier ${file.ext || "non PDF"} : analyse impossible` };
    } else {
      try { res.analysis = await ad.analyze(file.bytes, res.file || chosen.name); }
      catch (e) { res.analysis = { error: (e && e.message) || String(e) }; }
      if (!res.analysis) res.analysis = { error: "analyse sans résultat" };
    }
  }
  return done(S.OK, okDetail);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    INTRA_DEFAULT_BASE, INTRA_DEFAULT_FLOWER, INTRA_DEFAULT_SOC, INTRA_ALL_SOC, INTRA_DEFAULT_CABINETS,
    INTRA_DEFAULT_REGEX, INTRA_DEFAULT_NAME, INTRA_DEFAULT_FOLDER, INTRA_DEFAULT_DELAY, INTRA_DEFAULT_MAX_AGE_DAYS,
    INTRA_STATUS, INTRA_COLS, INTRA_FIELDS, INTRA_SUPPLIER_COL, INTRA_SAISIE_COLS, INTRA_ID_COL, INTRA_IM_COL, INTRA_FICHE_DEFAULT,
    intraFicheUrl, intraTypeLocal, intraSplitAdr,
    intraNormName, intraNormLot, intraFold, intraIndexUrl, intraFormUrl, intraLotUrl, intraFlowerViewUrl, intraFlowerDownloadUrl,
    intraBuildIndex, intraSimilarity, intraCloseNames, intraLookup, intraCompileRegex, intraIsDpeName, intraSelectDpe,
    intraSanitize, intraFolder, intraExtOf, intraFileName, intraClassifyFile, intraNormalizeConfig, intraRowValues, intraTemplateVars, intraSaisieValues, intraSaisieCols,
    intraPageOp, intraGetFile, intraProcessRow
  };
}
