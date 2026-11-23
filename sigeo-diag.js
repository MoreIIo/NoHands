/* ====================================================================
 * NoHands OSA — Étape de scénario « SIGEO : saisir un DPE »
 * --------------------------------------------------------------------
 * Remplit le formulaire « Diagnostic » de SIGEO (popup ouverte en onglet :
 *   …/popup.aspx/fr/sig_imm/diag_manage/{id}/10/{idDiag}, idDiag = 0 en
 * création), clique « Enregistrer », puis joint le PDF du DPE
 * (description « DPE »).
 *
 * Principe : le side panel pilote ; chaque action est une fonction
 * autonome injectée dans la page (diagPageOp). Un __doPostBack recharge
 * toute la page, ce qui ne coupe pas le déroulé : le panneau attend le
 * nouveau document (identifiant posé sur <html>, readyState complete),
 * le relit et reprend. La saisie est une boucle « lire la page → faire
 * la PREMIÈRE action manquante (ordre du formulaire) → attendre » : elle
 * est donc idempotente et reprend d'elle-même après n'importe quel
 * rechargement.
 *
 * Garde-fous : jamais de touche Entrée (Fournisseur et Description
 * déclenchent Enregistrer), jamais « Enreg. et fermer » ; enregistrement
 * et upload protégés contre les doublons (état de la page + journal dans
 * chrome.storage.session) ; le PDF est retenu par correspondance stricte
 * (nom exact, sinon nom contenant le N° ADEME, sinon erreur) et son N°
 * ADEME est contrôlé contre celui affiché par SIGEO après enregistrement.
 *
 * Chargé par sidepanel.html AVANT sidepanel.js ; importable sous Node
 * (tests). diagPageOp est injectée telle quelle : elle doit rester
 * autonome (aucune référence au reste du module).
 * ==================================================================== */

const DIAG_IDS = Object.freeze({
  type: "body_x_ddlDiagnostic_ddlDiagnostic",
  fou: "body_x_proxyFou_x_txt_x_txt",
  fouBtn: "body_x_proxyFou_x_btn",
  dateReal: "body_x_txtDateReal_x_txtDateReal",
  dateFin: "body_x_txtDateFinValid_x_txtDateFinValid",
  dateFinLabel: "fake_txtDateFinValid_x_txtDateFinValid",
  ademeTitle: "body_x_titreIdDpe",
  ademeLabel: "body_x_labelIdDpe",
  divConso: "div_conso",
  conso: "body_x_txtConsoE_x_txtConsoE",
  consoLettre: "body_x_txtConsoLegende_x__ctl0",
  ges: "body_x_txtEmission_x_txtEmission",
  gesLettre: "body_x_txtEmisLegende_x__ctl0",
  forceEp: "body_x_txtForceDiag_x__ctl0",
  forceGes: "body_x_txtForceGes_x__ctl0",
  altitude: "body_x_chkAltitude_x__ctl0",
  nouvelles: "body_x_chkSurface_x__ctl0",
  divCouts: "div_couts_annuels",
  coutMin: "body_x_txtEstimationBasse_x_txtEstimationBasse",
  coutMax: "body_x_txtEstimationHaute_x_txtEstimationHaute",
  datePrix: "body_x_txtDateReferencePrix_x_txtDateReferencePrix",
  resultat: "body_x_txtResultat_x_txtResultat",
  comment: "body_x_txtComment_x_txtComment",
  save: "proxyActionBar_x__cmdSave_x__btn",
  saveClose: "proxyActionBar_x__cmdEnd_x__btn",
  close: "proxyActionBar_x_BackAction_x__btn",
  msgLine: "msgline",
  msgBody: "msgbody",
  trDesc: "body_x_prxFileManager_x_trDesc",
  trFichier: "body_x_prxFileManager_x_trFichier",
  docDesc: "body_x_prxFileManager_x_txtDocDescription_x__ctl0",
  docFile: "body_x_prxFileManager_x_fileUpload",
  linkAdd: "body_x_prxFileManager_x_linkAjoutDocument",
  grid: "body_x_grid_x"
});

// Colonnes du scénario (modèles {Colonne} modifiables dans l'étape).
const DIAG_FIELDS = Object.freeze([
  { k: "type_local", label: "Type de local", hint: "habitation (défaut) / hors habitation" },
  { k: "code_fournisseur", label: "Code fournisseur", req: true },
  { k: "date_etabli", label: "Établi le", req: true },
  { k: "date_validite", label: "Valable jusqu'au (opt.)" },
  { k: "numero_ademe", label: "N° ADEME" },
  { k: "conso_ep", label: "Conso EP (kWh/m²/an)", req: true },
  { k: "emission_ges", label: "Émission GES (kgCO2/m²/an)", req: true },
  { k: "lettre_ep_forcee", label: "Lettre EP forcée (opt.)" },
  { k: "lettre_ges_forcee", label: "Lettre GES forcée (opt.)" },
  { k: "altitude_800m", label: "Altitude > 800 m (oui/non)" },
  { k: "nouvelles_valeurs", label: "Nouvelles valeurs (oui/non)" },
  { k: "cout_min", label: "Coût min (€)" },
  { k: "cout_max", label: "Coût max (€)" },
  { k: "date_prix", label: "Date d'indexation des prix" },
  { k: "resultat", label: "Résultat (opt.)" },
  { k: "commentaire", label: "Commentaire (opt.)" },
  { k: "fichier_dpe", label: "Fichier DPE (nom)" },
  { k: "lot", label: "Lot (affichage)" }
]);

const DIAG_PRESETS = Object.freeze({
  // Noms de colonnes du cahier des charges.
  default: Object.freeze(Object.fromEntries(DIAG_FIELDS.map((f) => [f.k, "{" + f.k + "}"]))),
  // Colonnes écrites par l'extraction « Analyse DPE » (v2.14.0).
  analyse: Object.freeze({
    type_local: "{type_local}", code_fournisseur: "{N° FU}", date_etabli: "{Établi le}",
    date_validite: "", numero_ademe: "{N° ADEME}", conso_ep: "{Énergie primaire (kWh/m²/an)}",
    emission_ges: "{GES (kg CO2/m²/an)}", lettre_ep_forcee: "", lettre_ges_forcee: "",
    altitude_800m: "", nouvelles_valeurs: "", cout_min: "{Coût min (€/an)}", cout_max: "{Coût max (€/an)}",
    date_prix: "{date_prix}", resultat: "", commentaire: "", fichier_dpe: "{DPE fichier}", lot: "{lot}"
  })
});

const DIAG_DEFAULTS = Object.freeze({
  desc: "DPE", resultCol: "Saisie DPE", timeoutMs: 20000, uploadTimeoutMs: 60000, maxPdfBytes: 40 * 1024 * 1024
});

/* ---------------------------------------------------------------- *
 * 1. Normalisation des données de la ligne                         *
 * ---------------------------------------------------------------- */

// NBSP → espace, ’ → ', espaces réduits, trim.
function diagNorm(s) {
  return String(s ?? "")
    .replace(/[   ]/g, " ")
    .replace(/[‘’ʼ]/g, "'")
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/ *\n */g, "\n")
    .trim();
}
function diagFold(s) {
  return diagNorm(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
}

const diagPad = (n) => String(n).padStart(2, "0");
const diagDaysIn = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
const diagFmtDate = (d, m, y) => `${diagPad(d)}/${diagPad(m)}/${y}`;

// N° de série Excel (1 = 01/01/1900, 25569 = 01/01/1970) → jj/mm/aaaa.
function diagSerialToDate(n) {
  const dt = new Date(Math.round((Math.floor(Number(n)) - 25569) * 86400000));
  return diagFmtDate(dt.getUTCDate(), dt.getUTCMonth() + 1, dt.getUTCFullYear());
}

// Date de la ligne → { ok, value: "jj/mm/aaaa" | "" } | { ok:false, error }.
// cell : cellule SheetJS brute { t, v } si disponible (prioritaire : la
// valeur affichée par SheetJS pour un format date par défaut est m/d/yy).
function diagDate(input, cell) {
  if (cell && cell.t === "n" && Number.isFinite(cell.v) && cell.v >= 1 && cell.v < 2958466) {
    return { ok: true, value: diagSerialToDate(cell.v), from: "série Excel" };
  }
  if (cell && cell.t === "d" && cell.v instanceof Date && !isNaN(cell.v)) {
    const d = cell.v;
    return { ok: true, value: diagFmtDate(d.getDate(), d.getMonth() + 1, d.getFullYear()) };
  }
  const s = diagNorm(input);
  if (!s) return { ok: true, value: "" };
  let d, mo, y, m;
  if ((m = s.match(/^(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(\d{4}|\d{2})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/))) {
    d = +m[1]; mo = +m[2]; y = +m[3];
    if (m[3].length === 2) y += y < 70 ? 2000 : 1900;
  } else if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ][\d:.]+Z?)?$/))) {
    y = +m[1]; mo = +m[2]; d = +m[3];
  } else if (/^\d{4,5}(?:[.,]\d+)?$/.test(s)) {
    const n = parseFloat(s.replace(",", "."));
    if (n >= 10000 && n < 80000) return { ok: true, value: diagSerialToDate(n), from: "série Excel" };
    return { ok: false, error: `date illisible « ${s} »` };
  } else {
    return { ok: false, error: `date illisible « ${s} » (attendu jj/mm/aaaa)` };
  }
  if (mo < 1 || mo > 12 || d < 1 || y < 1900 || y > 2100 || d > diagDaysIn(y, mo)) {
    return { ok: false, error: `date invalide « ${s} »` };
  }
  return { ok: true, value: diagFmtDate(d, mo, y) };
}

// « jj/mm/aaaa » → timestamp UTC (NaN si invalide).
function diagDateTs(s) {
  const m = String(s || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  return m ? Date.UTC(+m[3], +m[2] - 1, +m[1]) : NaN;
}

// Égalité de dates (valeur de la page éventuellement non complétée de 0).
function diagDateEq(pageVal, want) {
  const r = diagDate(pageVal);
  return r.ok && !!r.value && r.value === want;
}

// Règles de validité (calcul de SIGEO) : < 2013 → +10 ans −1 j ;
// 2013–2017 → 31/12/2022 ; 01/2018–06/2021 → 31/12/2024 ; ≥ 01/07/2021 → +10 ans −1 j.
function diagExpectedValidity(date) {
  const m = String(date || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return "";
  const d = +m[1], mo = +m[2], y = +m[3];
  const plus10 = () => {
    const t = new Date(Date.UTC(y + 10, mo - 1, d) - 86400000);
    return diagFmtDate(t.getUTCDate(), t.getUTCMonth() + 1, t.getUTCFullYear());
  };
  if (y < 2013) return plus10();
  if (y <= 2017) return "31/12/2022";
  if (y < 2021 || (y === 2021 && mo < 7)) return "31/12/2024";
  return plus10();
}

// Le champ N° ADEME n'apparaît (hypothèse) que pour un DPE ≥ 01/07/2021.
function diagIsNewDpe(date) {
  const ts = diagDateTs(date);
  return Number.isFinite(ts) && ts >= Date.UTC(2021, 6, 1);
}

// Texte → nombre (null si illisible). « 1 234,5 kWh/m²/an » → 1234.5 ;
// « 320,00 » → 320 ; « 1.234,5 » → 1234.5.
function diagParseNum(s) {
  const t0 = diagNorm(s).replace(/(\d)\s+(?=\d)/g, "$1");
  const m = t0.match(/-?\d[\d.,]*/);
  if (!m) return null;
  let t = m[0].replace(/[.,]+$/, "");
  const lc = t.lastIndexOf(","), ld = t.lastIndexOf(".");
  if (lc >= 0 && ld >= 0) t = lc > ld ? t.replace(/\./g, "").replace(",", ".") : t.replace(/,/g, "");
  else if (lc >= 0) t = (t.match(/,/g).length > 1) ? t.replace(/,/g, "") : t.replace(",", ".");
  else if ((t.match(/\./g) || []).length > 1) t = t.replace(/\./g, "");
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}

// Nombre de la ligne → texte à saisir : entier, sinon virgule et au plus
// 2 décimales, dans la limite de maxLen caractères (maxlength du champ).
function diagNum(input, cell, maxLen) {
  let n;
  if (cell && cell.t === "n" && Number.isFinite(cell.v)) n = cell.v;
  else {
    const s = diagNorm(input);
    if (!s) return { ok: true, value: "", num: null };
    n = diagParseNum(s);
    if (n === null) return { ok: false, error: `nombre illisible « ${s} »` };
  }
  if (n < 0) return { ok: false, error: `nombre négatif : ${n}` };
  for (let dec = 2; dec >= 0; dec--) {
    const r = Math.round(n * 10 ** dec) / 10 ** dec;
    const txt = (dec ? r.toFixed(dec).replace(/0+$/, "").replace(/\.$/, "") : String(Math.round(n))).replace(".", ",");
    if (!maxLen || txt.length <= maxLen) {
      const num = Number(txt.replace(",", "."));
      return { ok: true, value: txt, num, rounded: Math.abs(num - n) > 1e-9 };
    }
  }
  return { ok: false, error: `nombre trop long (${maxLen} caractères max) : ${n}` };
}

// Égalité numérique entre la valeur relue (« 320,00 ») et la valeur saisie.
function diagNumEq(pageVal, want) {
  const a = diagParseNum(pageVal), b = diagParseNum(want);
  return a !== null && b !== null && Math.abs(a - b) < 0.005;
}

function diagBool(input) {
  const f = diagFold(input);
  if (!f) return { ok: true, value: null };
  if (/^(oui|o|yes|y|1|x|vrai|true|coche|cochee)$/.test(f)) return { ok: true, value: true };
  if (/^(non|n|no|0|faux|false|-)$/.test(f)) return { ok: true, value: false };
  return { ok: false, error: `valeur oui/non illisible « ${input} »` };
}

function diagLetter(input) {
  const s = diagNorm(input).toUpperCase();
  if (!s) return { ok: true, value: "" };
  return /^[A-G]$/.test(s) ? { ok: true, value: s } : { ok: false, error: `lettre invalide « ${input} » (A à G)` };
}

const DIAG_ADEME_FMT = /^\d{4}[A-Z]\d{7}[A-Z]$/;
function diagAdeme(input) {
  const s = diagNorm(input).toUpperCase().replace(/\s+/g, "");
  return { ok: true, value: s, plausible: !s || DIAG_ADEME_FMT.test(s) };
}

function diagType(input) {
  const f = diagFold(input);
  if (!f || f === "11" || /^(hab|habitation|logement)$/.test(f) || (f.includes("habitation") && !f.includes("hors"))) return { ok: true, value: "11" };
  if (f === "15" || f.includes("hors") || f.includes("tertiaire")) return { ok: true, value: "15" };
  return { ok: false, error: `type_local inconnu « ${input} » (habitation / hors habitation)` };
}

// raw[k] = { text, cell?, missing?: [colonnes introuvables] }
// → { v, errors[], warnings[] }
function diagNormalizeData(raw, now = Date.now()) {
  const errors = [], warnings = [];
  const g = (k) => raw[k] || { text: "" };
  const txt = (k) => diagNorm(g(k).text);
  const label = (k) => (DIAG_FIELDS.find((f) => f.k === k) || {}).label || k;
  const emptyMsg = (k) => {
    const r = g(k);
    return r.missing && r.missing.length
      ? `${label(k)} : colonne ${r.missing.map((m) => "« " + m + " »").join(", ")} introuvable`
      : `${label(k)} vide`;
  };
  const v = {};

  const t = diagType(txt("type_local"));
  if (t.ok) v.type = t.value; else errors.push(t.error);

  v.fou = txt("code_fournisseur").replace(/\s+/g, "");
  if (!v.fou) errors.push(emptyMsg("code_fournisseur"));
  else if (v.fou.length > 9) errors.push(`code fournisseur trop long (9 caractères max) : ${v.fou}`);

  const date = (k, key, required) => {
    const r = diagDate(g(k).text, g(k).cell);
    if (!r.ok) { errors.push(`${label(k)} : ${r.error}`); v[key] = ""; return; }
    v[key] = r.value;
    if (!r.value && required) errors.push(emptyMsg(k));
  };
  date("date_etabli", "dateEtabli", true);
  date("date_validite", "dateValidite", false);
  date("date_prix", "datePrix", false);
  if (v.dateEtabli && diagDateTs(v.dateEtabli) > now) errors.push(`date d'établissement dans le futur : ${v.dateEtabli}`);

  const a = diagAdeme(txt("numero_ademe"));
  v.ademe = a.value;
  if (v.ademe && !a.plausible) warnings.push(`N° ADEME au format inhabituel : ${v.ademe} (attendu 4 chiffres, 1 lettre, 7 chiffres, 1 lettre)`);
  if (!v.ademe && v.dateEtabli && diagIsNewDpe(v.dateEtabli)) warnings.push(`N° ADEME vide pour un DPE établi le ${v.dateEtabli}`);

  const num = (k, key, maxLen, required) => {
    const r = diagNum(g(k).text, g(k).cell, maxLen);
    if (!r.ok) { errors.push(`${label(k)} : ${r.error}`); v[key] = ""; return; }
    v[key] = r.value;
    if (!r.value && required) errors.push(emptyMsg(k));
    if (r.rounded) warnings.push(`${label(k)} arrondi à ${r.value}`);
  };
  num("conso_ep", "conso", 6, true);
  num("emission_ges", "ges", 6, true);
  num("cout_min", "coutMin", 5, false);
  num("cout_max", "coutMax", 5, false);
  const cMin = diagParseNum(v.coutMin), cMax = diagParseNum(v.coutMax);
  if (cMin !== null && cMax !== null && cMin > cMax) warnings.push(`coût min (${v.coutMin}) > coût max (${v.coutMax})`);
  if ((v.coutMin || v.coutMax || v.datePrix) && !(v.coutMin && v.coutMax && v.datePrix)) {
    warnings.push("coûts annuels incomplets (min, max et date d'indexation attendus)");
  }

  for (const [k, key] of [["lettre_ep_forcee", "forceEp"], ["lettre_ges_forcee", "forceGes"]]) {
    const r = diagLetter(txt(k));
    if (r.ok) v[key] = r.value; else { errors.push(`${label(k)} : ${r.error}`); v[key] = ""; }
  }
  for (const [k, key] of [["altitude_800m", "altitude"], ["nouvelles_valeurs", "nouvelles"]]) {
    const r = diagBool(txt(k));
    if (r.ok) v[key] = r.value; else { errors.push(`${label(k)} : ${r.error}`); v[key] = null; }
  }
  v.resultat = diagNorm(g("resultat").text);
  if (v.resultat.length > 1000) errors.push(`résultat trop long (${v.resultat.length} / 1000 caractères)`);
  v.comment = diagNorm(g("commentaire").text);
  v.fichier = diagNorm(g("fichier_dpe").text);
  v.lot = diagNorm(g("lot").text);
  return { v, errors, warnings };
}

/* ---------------------------------------------------------------- *
 * 2. Fichier : correspondance stricte, contrôles                    *
 * ---------------------------------------------------------------- */

const diagBaseName = (s) => String(s ?? "").split(/[\\/]/).pop().trim();

// files : [{ name, path? }]. Règles : 1) nom exact (fichier_dpe, « .pdf »
// ajouté s'il n'a pas d'extension) ; 2) sinon nom contenant le N° ADEME ;
// 3) sinon, ou plusieurs candidats → erreur. Jamais de « meilleure
// supposition ».
function diagMatchFile(files, { name, ademe } = {}) {
  const list = files || [];
  const show = (hits) => hits.slice(0, 4).map((f) => f.path || f.name).join(", ") + (hits.length > 4 ? "…" : "");
  const tried = [];
  const base = diagBaseName(name);
  if (base) {
    const want = diagFold(base);
    const wantPdf = /\.[a-z0-9]{2,5}$/i.test(base) ? "" : want + ".pdf";
    const hits = list.filter((f) => { const n = diagFold(f.name); return n === want || (wantPdf && n === wantPdf); });
    if (hits.length === 1) return { file: hits[0], via: "nom exact" };
    if (hits.length > 1) return { error: `ambigu : ${hits.length} fichiers nommés « ${base} » (${show(hits)})`, ambiguous: true };
    tried.push(`aucun fichier nommé « ${base} »`);
  }
  const a = String(ademe || "").toUpperCase().replace(/\s+/g, "");
  if (a) {
    const hits = list.filter((f) => String(f.name).toUpperCase().includes(a));
    if (hits.length === 1) return { file: hits[0], via: "nom contenant le N° ADEME" };
    if (hits.length > 1) return { error: `ambigu : ${hits.length} fichiers contiennent le N° ADEME ${a} (${show(hits)})`, ambiguous: true };
    tried.push(`aucun fichier dont le nom contient ${a}`);
  }
  if (!base && !a) return { error: "introuvable : ni nom de fichier (fichier_dpe) ni N° ADEME pour le retrouver" };
  return { error: `introuvable : ${tried.join(" ; ")} (${list.length} fichier(s) dans le dossier)` };
}

const DIAG_ADEME_RE = /(?<![A-Z0-9])(\d{4}[A-Z]\d{7}[A-Z])(?![A-Z0-9])/g;
const diagAdemeTokens = (s) => [...new Set([...String(s || "").toUpperCase().matchAll(DIAG_ADEME_RE)].map((m) => m[1]))];

// Le fichier correspond-il au N° ADEME ? status : "name" (dans le nom),
// "text" (dans le texte du PDF), "other" (un AUTRE n° ADEME trouvé),
// "unknown" (rien de lisible), "needText" (appeler à nouveau avec le
// texte), "none" (pas de N° ADEME de référence).
function diagCheckFileAdeme(fileName, text, ademe) {
  const A = String(ademe || "").toUpperCase().replace(/\s+/g, "");
  if (!A) return { status: "none" };
  const up = String(fileName || "").toUpperCase();
  if (up.includes(A)) return { status: "name" };
  const otherName = diagAdemeTokens(up).filter((x) => x !== A);
  if (otherName.length) return { status: "other", where: "nom du fichier", others: otherName };
  if (text === undefined) return { status: "needText" };
  const T = String(text || "").toUpperCase();
  if (T.replace(/\s+/g, "").includes(A)) return { status: "text" };
  const others = diagAdemeTokens(T).filter((x) => x !== A);
  if (others.length) return { status: "other", where: "texte du PDF", others };
  return { status: "unknown" };
}

function diagIsPdf(bytes) {
  if (!bytes || !bytes.length) return false;
  let i = 0;
  while (i < bytes.length && i < 16 && [0x20, 0x09, 0x0a, 0x0d, 0xef, 0xbb, 0xbf].includes(bytes[i])) i++;
  return bytes[i] === 0x25 && bytes[i + 1] === 0x50 && bytes[i + 2] === 0x44 && bytes[i + 3] === 0x46; // %PDF
}

function diagToBase64(bytes) {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

function diagFmtSize(n) {
  return n >= 1048576 ? (n / 1048576).toFixed(1).replace(".", ",") + " Mo" : Math.max(1, Math.round(n / 1024)) + " Ko";
}

/* ---------------------------------------------------------------- *
 * 3. URL, journal, messages, grille                                *
 * ---------------------------------------------------------------- */

function diagUrlParts(url) {
  const m = String(url || "").match(/\/diag_manage\/([^/?#]+)\/([^/?#]+)\/([^/?#]+)/i);
  return m ? { lotId: decodeURIComponent(m[1]), kind: m[2], idDiag: m[3] } : null;
}
function diagUrlWithId(url, id) {
  return String(url).replace(/(\/diag_manage\/[^/?#]+\/[^/?#]+\/)([^/?#]+)/i, "$1" + encodeURIComponent(id));
}

// Clé du journal : fiche (id de l'URL) + N° ADEME (à défaut date|conso|ges).
function diagJournalKey(lotId, v) {
  return `${lotId || "?"}#${v.ademe || [v.dateEtabli, v.conso, v.ges].join("|")}`;
}

// #msgbody : <li class="infomsg">Enregistrement réussi !</li> = succès ;
// infomsg autre = information ; toute autre li = erreur.
function diagClassifyMessages(m) {
  const out = { success: false, infos: [], errors: [] };
  for (const it of (m && m.items) || []) {
    const info = /infomsg/i.test(it.cls || "");
    if (info && /enregistrement reussi/.test(diagFold(it.text))) out.success = true;
    else if (info) out.infos.push(it.text);
    else out.errors.push(it.text);
  }
  return out;
}

function diagFindDocRow(grid, desc) {
  const d = diagFold(desc || DIAG_DEFAULTS.desc);
  return ((grid && grid.rows) || []).find((cells) => cells.some((c) => diagFold(c) === d)) || null;
}
function diagRowHasFile(row, name) {
  const n = diagFold(diagBaseName(name)), base = n.replace(/\.[a-z0-9]{2,5}$/, "");
  return !!row && row.some((c) => { const f = diagFold(c); return f.includes(n) || (base && f.includes(base)); });
}

/* ---------------------------------------------------------------- *
 * 4. Fonction injectée dans la page SIGEO (autonome)               *
 * ---------------------------------------------------------------- *
 * cfg.ids = DIAG_IDS ; cfg.op :
 *   "read"       → instantané de la page (champs, messages, grille…)
 *   "dump"       → instantané + liste de tous les champs (mode debug)
 *   "set"        { actions:[{ key, kind:value|select|check, id|locate, value|checked, trigger? }] }
 *                  l'action « trigger » (postback attendu) est déclenchée
 *                  en différé, pour que le résultat reparte avant le rechargement
 *   "save"       → clic différé sur « Enregistrer » (jamais « Enreg. et fermer »)
 *   "showAttach" → clic sur « Joindre un document » si le champ Fichier est masqué
 *   "attach"     { desc, b64, name, size } → Description puis fichier
 *                  (DataTransfer) ; le change part en différé (upload).
 * Aucune touche clavier n'est jamais simulée.
 */
function diagPageOp(cfg) {
  const ID = cfg.ids;
  const doc = document;
  const byId = (id) => doc.getElementById(id);
  const norm = (s) => String(s ?? "").replace(/[   ]/g, " ").replace(/[‘’]/g, "'").replace(/\s+/g, " ").trim();
  const fold = (s) => norm(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const fire = (el, t) => el.dispatchEvent(new Event(t, { bubbles: true }));

  // Identifiant du document : change à chaque rechargement (postback).
  const root = doc.documentElement;
  if (root && !root.getAttribute("data-nh-diag-doc")) {
    root.setAttribute("data-nh-diag-doc", Date.now().toString(36) + Math.random().toString(36).slice(2, 8));
  }
  const docId = root ? root.getAttribute("data-nh-diag-doc") : "";

  function hidden(el) {
    for (let e = el; e && e.nodeType === 1; e = e.parentElement) {
      if (e.hidden) return true;
      if (e.tagName === "INPUT" && String(e.getAttribute("type") || "").toLowerCase() === "hidden") return true;
      if (e.style && (e.style.display === "none" || e.style.visibility === "hidden")) return true;
      let cs = null;
      try { cs = getComputedStyle(e); } catch (_) { /* jsdom / frame */ }
      if (cs && (cs.display === "none" || cs.visibility === "hidden")) return true;
    }
    return false;
  }
  const visible = (el) => !!el && el.isConnected && !hidden(el);
  const valOf = (el) => {
    if (!el) return "";
    if (/^(INPUT|SELECT|TEXTAREA)$/.test(el.tagName)) return String(el.value ?? "");
    return norm(el.textContent);
  };
  const info = (el) => el ? {
    present: true, visible: visible(el), id: el.id, value: valOf(el),
    checked: el.type === "checkbox" ? !!el.checked : undefined,
    maxLength: el.maxLength > 0 ? el.maxLength : 0, disabled: !!el.disabled, readOnly: !!el.readOnly
  } : { present: false, visible: false, value: "" };

  const knownIds = new Set(Object.values(ID));
  const textInput = (i) => i && i.tagName === "INPUT" && /^(text|)$/i.test(i.getAttribute("type") || "") && !knownIds.has(i.id);

  // N° ADEME : en création, input sans id connu (rendu serveur) dans la
  // DERNIÈRE cellule de la ligne qui contient « Établi le » ; une fois
  // enregistré, texte dans span#body_x_labelIdDpe (plus d'input).
  function ademeEl() {
    const dr = byId(ID.dateReal);
    if (dr) {
      for (let tr = dr.closest("tr"); tr; tr = tr.parentElement && tr.parentElement.closest("tr")) {
        const tds = Array.from(tr.children).filter((c) => c.tagName === "TD");
        const last = tds[tds.length - 1];
        if (!last || last.contains(dr)) continue;
        const inp = Array.from(last.querySelectorAll("input")).find((i) => textInput(i) && visible(i));
        if (inp) return { el: inp, via: "dernière cellule de la ligne « Établi le »" };
      }
    }
    const byName = Array.from(doc.querySelectorAll("input"))
      .find((i) => textInput(i) && visible(i) && /ademe|iddpe|numdpe/i.test((i.id || "") + " " + (i.name || "")));
    if (byName) return { el: byName, via: "id/name" };
    for (const el of doc.querySelectorAll("label,span,td,th")) {
      if (el.id === ID.ademeLabel || !/ademe/.test(fold(el.textContent)) || norm(el.textContent).length > 40) continue;
      const f = el.getAttribute("for");
      let cand = f ? byId(f) : null;
      if (!(textInput(cand) && visible(cand))) {
        const cell = el.closest("td,th");
        cand = cell ? Array.from(cell.querySelectorAll("input")).find((i) => textInput(i) && visible(i)) : null;
      }
      if (cand) return { el: cand, via: "libellé « ADEME »" };
    }
    return { el: null, via: "" };
  }
  function ademe() {
    const lab = byId(ID.ademeLabel);
    const labelText = lab ? norm(lab.textContent) : "";
    const a = ademeEl();
    return {
      mode: a.el ? "input" : (labelText ? "label" : "absent"),
      input: a.el ? info(a.el) : null, via: a.via, labelText, title: !!byId(ID.ademeTitle)
    };
  }

  // Nom du fournisseur affiché à côté du code (après enregistrement).
  function supplier() {
    const inp = byId(ID.fou), btn = byId(ID.fouBtn);
    if (!inp) return { name: "", dedicated: false, src: "" };
    const own = Array.from(doc.querySelectorAll('[id^="body_x_proxyFou"]'))
      .filter((e) => e !== inp && e !== btn && !e.contains(inp) && !/^(SCRIPT|STYLE)$/.test(e.tagName) &&
        !(e.tagName === "INPUT" && !e.readOnly) && !(e.tagName === "INPUT" && /hidden|button|submit|image/i.test(e.type)));
    for (const e of own) {
      const t = valOf(e);
      if (t && t !== norm(inp.value)) return { name: t, dedicated: true, src: "#" + e.id };
    }
    const td = inp.closest("td");
    const strip = (t) => norm(String(t).replace(/s[ée]lectionner/gi, "").replace(norm(inp.value), ""));
    let name = td ? strip(td.textContent) : "";
    let src = td ? "cellule du code" : "";
    if (!name && td) {
      const n = td.nextElementSibling;
      if (n && !n.querySelector("input,select,textarea")) { name = strip(n.textContent); src = "cellule suivante"; }
    }
    return { name, dedicated: own.length > 0, src };
  }

  function messages() {
    const line = byId(ID.msgLine), body = byId(ID.msgBody);
    const shown = !!body && (line ? visible(line) : visible(body));
    const items = shown ? Array.from(body.querySelectorAll("li")).map((li) => ({ cls: li.className || "", text: norm(li.textContent) })).filter((x) => x.text) : [];
    if (shown && !items.length && norm(body.textContent)) items.push({ cls: "", text: norm(body.textContent) });
    return { shown, items };
  }

  function grid() {
    const t = byId(ID.grid);
    if (!t) return { present: false, count: 0, countText: false, rows: [] };
    // Compteur « N Résultat(s) » : lu dans SA cellule (le texte global colle
    // la date de la dernière ligne au compteur : « …/20261 Résultat(s) »).
    let m = null;
    for (const c of t.querySelectorAll("td,th,span,div")) {
      m = norm(c.textContent).match(/^(\d+)\s*r[ée]sultat\(?s?\)?$/i);
      if (m) break;
    }
    const rows = [];
    for (const tr of t.querySelectorAll("tr")) {
      if (/header|pager|footer|empty/i.test(tr.className || "") || tr.querySelector("th") || tr.querySelector("tr")) continue;
      const cells = Array.from(tr.children).filter((c) => /^T[DH]$/.test(c.tagName)).map((c) => norm(c.textContent));
      if (!cells.some(Boolean) || /^\d+\s*r[ée]sultat/i.test(cells.join(" ").trim())) continue;
      rows.push(cells);
    }
    const empty = !!t.querySelector("tr.PowerGridEmptyPagerClass");
    return { present: true, count: m ? +m[1] : (empty ? 0 : rows.length), countText: !!m, rows, empty };
  }

  function urlParts(u) {
    const m = String(u || "").match(/\/diag_manage\/([^/?#]+)\/([^/?#]+)\/([^/?#]+)/i);
    return m ? { lotId: decodeURIComponent(m[1]), idDiag: m[3] } : null;
  }

  function snapshot() {
    const up = urlParts(location.href);
    const form = byId("mainForm") || doc.forms[0];
    let ap = null;
    try { ap = form && form.getAttribute("action") ? urlParts(new URL(form.getAttribute("action"), location.href).href) : null; } catch (_) { /* ignoré */ }
    const idDiag = (up && up.idDiag !== "0" && up.idDiag) || (ap && ap.idDiag !== "0" && ap.idDiag) || (up ? up.idDiag : "");
    const g = (k) => info(byId(ID[k]));
    const f = {
      type: g("type"), fou: g("fou"), dateReal: g("dateReal"), dateFin: g("dateFin"),
      conso: g("conso"), consoLettre: g("consoLettre"), ges: g("ges"), gesLettre: g("gesLettre"),
      forceEp: g("forceEp"), forceGes: g("forceGes"), altitude: g("altitude"), nouvelles: g("nouvelles"),
      coutMin: g("coutMin"), coutMax: g("coutMax"), datePrix: g("datePrix"),
      resultat: g("resultat"), comment: g("comment"), ademe: ademe(), fouName: supplier()
    };
    const sel = byId(ID.type);
    f.type.text = sel && sel.selectedIndex >= 0 && sel.options[sel.selectedIndex] ? norm(sel.options[sel.selectedIndex].text) : "";
    const lab = byId(ID.dateFinLabel);
    f.dateFin.label = lab ? norm(lab.textContent) : "";
    const divC = byId(ID.divConso), divK = byId(ID.divCouts);
    const descEl = byId(ID.docDesc), fileEl = byId(ID.docFile), link = byId(ID.linkAdd);
    const attach = {
      present: !!(descEl || fileEl || link || byId(ID.trDesc)),
      descVisible: visible(descEl), fileVisible: visible(fileEl), linkPresent: !!link,
      descValue: descEl ? descEl.value : ""
    };
    const needSave = /veuillez enregistrer le diagnostic/.test(fold(doc.body ? doc.body.textContent : ""));
    const save = byId(ID.save);
    return {
      docId, url: location.href, readyState: doc.readyState, title: doc.title,
      login: !!doc.querySelector('input[type="password"]'),
      isDiag: /\/diag_manage\//i.test(location.href) && !!byId(ID.type),
      lotId: up ? up.lotId : "", idDiag,
      saved: !!f.ademe.labelText || (!!idDiag && idDiag !== "0") || (attach.present && !needSave),
      needSave, f,
      divConso: divC ? visible(divC) : visible(byId(ID.conso)),
      couts: { present: !!divK, visible: divK ? visible(divK) : visible(byId(ID.coutMin)) },
      attach, saveBtn: { present: !!save, visible: visible(save) },
      messages: messages(), grid: grid()
    };
  }

  function labelFor(el) {
    if (el.id) {
      try {
        const l = doc.querySelector(`label[for="${CSS.escape(el.id)}"]`);
        if (l && norm(l.textContent)) return norm(l.textContent).slice(0, 60);
      } catch (_) { /* ignoré */ }
    }
    const td = el.closest("td");
    if (!td) return "";
    for (let p = td.previousElementSibling; p; p = p.previousElementSibling) {
      if (norm(p.textContent)) return norm(p.textContent).slice(0, 60);
    }
    const tr = td.parentElement, prev = tr && tr.previousElementSibling;
    if (prev) {
      const h = prev.children[Array.from(tr.children).indexOf(td)];
      if (h && norm(h.textContent)) return "[colonne] " + norm(h.textContent).slice(0, 60);
    }
    return "";
  }

  function dump() {
    const out = [];
    const sel = "input,select,textarea,button,a[id],span[id],label[id],div[id],table[id],tr[id],li";
    for (const el of doc.querySelectorAll(sel)) {
      if (out.length >= 400) break;
      const tag = el.tagName.toLowerCase();
      const type = (el.getAttribute("type") || "").toLowerCase();
      const name = el.getAttribute("name") || "";
      if (name.startsWith("__")) continue; // __VIEWSTATE, __EVENTTARGET…
      const field = /^(input|select|textarea)$/.test(tag);
      if (tag === "li" && !el.closest("#" + ID.msgBody)) continue;
      if (!field && tag !== "button" && tag !== "li" && !/^(body_x|proxyActionBar|msg|div_|fake_)/.test(el.id || "")) continue;
      const item = { tag, type, id: el.id || "", name, visible: visible(el) };
      if (field) item.label = labelFor(el);
      if (tag === "select") item.value = `${el.value} (${el.selectedIndex >= 0 && el.options[el.selectedIndex] ? norm(el.options[el.selectedIndex].text) : ""})`;
      else if (type === "checkbox" || type === "radio") item.checked = el.checked;
      else if (type !== "file") item.value = (field ? String(el.value ?? "") : norm(el.textContent)).slice(0, 80);
      if (field && el.maxLength > 0) item.maxLength = el.maxLength;
      if (el.readOnly) item.readOnly = true;
      if (tag === "li") item.cls = el.className;
      for (const h of ["onchange", "onclick", "onblur", "onkeydown", "onkeyup", "href"]) {
        const v = el.getAttribute(h);
        if (v && (h !== "href" || /^javascript:/i.test(v))) item[h] = v.slice(0, 120);
      }
      out.push(item);
    }
    return out;
  }

  // Élément visé par une action.
  const target = (a) => a.locate === "ademe" ? ademeEl().el : byId(a.id);
  function check(a, el) {
    if (!el) return `champ introuvable (${a.label || a.key})`;
    if (a.kind === "select" && !Array.from(el.options || []).some((o) => o.value === a.value)) {
      return `option « ${a.value} » absente de la liste ${a.label || a.key}`;
    }
    if (a.kind === "check" && el.type !== "checkbox") return `${a.label || a.key} n'est pas une case à cocher`;
    return "";
  }
  function apply(a, el) {
    if (a.kind === "check") {
      if (!!el.checked !== !!a.checked) el.click(); // clic natif : onclick + change de la page
      return;
    }
    try { el.focus(); } catch (_) { /* ignoré */ }
    el.value = a.value;
    fire(el, "input");
    fire(el, "change");
  }

  if (cfg.op === "read") return snapshot();
  if (cfg.op === "dump") return Object.assign(snapshot(), { dump: dump() });

  if (cfg.op === "set") {
    const acts = cfg.actions || [];
    const results = [];
    for (const a of acts.filter((x) => !x.trigger)) {
      const el = target(a);
      const err = check(a, el);
      if (err) return { ok: false, error: err, results };
      apply(a, el);
      results.push({ key: a.key, after: a.kind === "check" ? el.checked : el.value });
    }
    const trig = acts.find((x) => x.trigger);
    if (trig) {
      const el = target(trig);
      const err = check(trig, el);
      if (err) return { ok: false, error: err, results };
      // Différé : le résultat doit repartir avant le rechargement de la page.
      setTimeout(() => apply(trig, el), 60);
      results.push({ key: trig.key, deferred: true });
    }
    return { ok: true, results, docId };
  }

  if (cfg.op === "save") {
    const btn = byId(ID.save);
    if (!btn) return { ok: false, error: "bouton « Enregistrer » introuvable" };
    setTimeout(() => btn.click(), 60);
    return { ok: true, docId };
  }

  if (cfg.op === "showAttach") {
    const fileEl = byId(ID.docFile);
    if (visible(fileEl)) return { ok: true, already: true, docId };
    const link = byId(ID.linkAdd);
    if (!link) return { ok: false, error: "lien « Joindre un document » introuvable" };
    link.click(); // showAddDocument() bascule l'affichage : un seul clic
    return { ok: true, clicked: true, visibleNow: visible(byId(ID.docFile)), docId };
  }

  if (cfg.op === "attach") {
    const desc = byId(ID.docDesc), inp = byId(ID.docFile);
    if (!desc || !inp) return { ok: false, error: "champs Description / Fichier introuvables" };
    if (!visible(inp) || !visible(desc)) return { ok: false, error: "champs Description / Fichier masqués" };
    if (norm(desc.value) !== cfg.desc) {
      try { desc.focus(); } catch (_) { /* ignoré */ }
      desc.value = cfg.desc;
      fire(desc, "input");
      fire(desc, "change");
    }
    if (norm(desc.value) !== cfg.desc) return { ok: false, error: `Description non acceptée (« ${desc.value} »)` };
    const bin = atob(cfg.b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    if (bytes.length !== cfg.size) return { ok: false, error: `taille reçue ${bytes.length} ≠ ${cfg.size} octets` };
    const file = new File([bytes], cfg.name, { type: "application/pdf" });
    const dt = new DataTransfer();
    dt.items.add(file);
    inp.files = dt.files;
    const got = inp.files && inp.files[0];
    if (!got || inp.files.length !== 1 || got.name !== cfg.name || got.size !== cfg.size) {
      return { ok: false, error: "le fichier n'a pas pu être posé dans le champ Fichier" };
    }
    // onchange="__doPostBack('')" : l'upload part dès le change (différé).
    setTimeout(() => fire(inp, "change"), 60);
    return { ok: true, size: got.size, docId };
  }

  return { ok: false, error: "opération inconnue : " + cfg.op };
}

/* ---------------------------------------------------------------- *
 * 5. Prochaine action (logique pure)                                *
 * ---------------------------------------------------------------- *
 * Parcourt les champs dans l'ordre du formulaire et renvoie les actions
 * « douces » (sans rechargement attendu) jusqu'à la PREMIÈRE action qui
 * déclenche un postback (hard). Les valeurs saisies avant le
 * déclenchement partent avec le POST et survivent ; tout ce qui suivrait
 * serait perdu, d'où l'arrêt au premier postback.
 * st = { attempts: {clé: n}, warned: {clé: true} } (conservé entre passes).
 * → { soft[], hard|null, done, fatal, warnings[] } */
const DIAG_MAX_ATTEMPTS = 3;

function diagNextActions(s, v, st) {
  const I = DIAG_IDS;
  const F = s.f;
  const out = { soft: [], hard: null, done: false, fatal: "", warnings: [] };
  const warnOnce = (k, msg) => { if (!st.warned[k]) { st.warned[k] = true; out.warnings.push(msg); } };
  const can = (a) => {
    const n = st.attempts[a.key] || 0;
    if (n < DIAG_MAX_ATTEMPTS) return true;
    const want = a.kind === "check" ? (a.checked ? "coché" : "décoché") : `« ${a.value} »`;
    out.fatal = `la page n'accepte pas ${a.label} : saisi ${n} fois ${want}, relu « ${a.got ?? ""} »`;
    return false;
  };
  const soft = (a) => { if (!can(a)) return false; out.soft.push(a); return true; };
  const hard = (a) => { if (can(a)) out.hard = a; return out; };

  // 1. Type (postback)
  if (!F.type.present) { out.fatal = "liste « Type » introuvable"; return out; }
  if (F.type.value !== v.type) return hard({ key: "type", kind: "select", id: I.type, value: v.type, label: "Type", got: F.type.value });
  if (!s.divConso) { out.fatal = "bloc des consommations (div#div_conso) non affiché alors que le type est sélectionné"; return out; }

  // 2. Fournisseur (jamais d'Entrée : elle déclencherait Enregistrer)
  if (!F.fou.present) { out.fatal = "champ Fournisseur introuvable"; return out; }
  if (diagNorm(F.fou.value) !== v.fou && !soft({ key: "fou", kind: "value", id: I.fou, value: v.fou, label: "Fournisseur", got: F.fou.value })) return out;

  // 3. Établi le (postback ; calcule « Valable jusqu'au », fait apparaître le N° ADEME)
  if (!F.dateReal.present) { out.fatal = "champ « Établi le » introuvable"; return out; }
  const finVal = F.dateFin.visible ? F.dateFin.value : (F.dateFin.label || F.dateFin.value);
  if (!diagDateEq(F.dateReal.value, v.dateEtabli)) {
    return hard({ key: "dateReal", kind: "value", id: I.dateReal, value: v.dateEtabli, label: "Établi le", got: F.dateReal.value });
  }
  if (!diagNorm(finVal)) {
    // Date déjà à la bonne valeur sans calcul (saisie perdue avant postback) : on relance le change une fois.
    if ((st.attempts.dateReal || 0) >= 1) { out.fatal = "« Valable jusqu'au » non calculé après la saisie de la date"; return out; }
    return hard({ key: "dateReal", kind: "value", id: I.dateReal, value: v.dateEtabli, label: "Établi le", got: F.dateReal.value });
  }
  if (v.dateValidite) {
    if (F.dateFin.present && F.dateFin.visible) {
      if (!diagDateEq(F.dateFin.value, v.dateValidite) &&
        !soft({ key: "dateFin", kind: "value", id: I.dateFin, value: v.dateValidite, label: "Valable jusqu'au", got: F.dateFin.value })) return out;
    } else if (!diagDateEq(finVal, v.dateValidite)) {
      warnOnce("dateFinRO", `« Valable jusqu'au » affiché ${finVal} ≠ ${v.dateValidite} (non modifiable sur cette fiche)`);
    }
  } else {
    const exp = diagExpectedValidity(v.dateEtabli);
    if (exp && !diagDateEq(finVal, exp)) warnOnce("dateFinCalc", `« Valable jusqu'au » = ${finVal}, la règle donnerait ${exp}`);
  }

  // 4. N° ADEME (présent seulement pour certaines dates ; non modifiable une fois enregistré)
  const A = F.ademe;
  if (A.mode === "input") {
    if (v.ademe) {
      const cur = diagNorm(A.input.value).toUpperCase().replace(/\s+/g, "");
      if (cur !== v.ademe && !soft({ key: "ademe", kind: "value", locate: "ademe", value: v.ademe, label: "N° ADEME", got: A.input.value })) return out;
    } else warnOnce("ademeEmpty", "champ N° ADEME affiché mais N° ADEME vide dans les données");
  } else if (A.mode === "absent" && v.ademe) {
    warnOnce("ademeAbsent", `champ ADEME non affiché pour cette date (${v.dateEtabli}) : N° ${v.ademe} non saisi`);
  }

  // 5. Coûts annuels (bloc affiché par le serveur ; revérifié après chaque postback)
  if (s.couts.visible) {
    for (const [key, val, label, num] of [["coutMin", v.coutMin, "Coût min", true], ["coutMax", v.coutMax, "Coût max", true], ["datePrix", v.datePrix, "Date d'indexation des prix", false]]) {
      if (!val) { warnOnce("vide_" + key, `bloc « coûts annuels » affiché mais ${label} vide dans les données`); continue; }
      const el = F[key];
      if (!el.present) { warnOnce("abs_" + key, `champ ${label} introuvable dans le bloc coûts`); continue; }
      const eq = num ? diagNumEq(el.value, val) : diagDateEq(el.value, val);
      if (!eq && !soft({ key, kind: "value", id: I[key], value: val, label, got: el.value })) return out;
    }
  }

  // 6–7. Consommation puis émissions (postbacks ; lettres calculées par le serveur)
  if (!F.conso.present) { out.fatal = "champ « Conso énergie primaire » introuvable"; return out; }
  if (!diagNumEq(F.conso.value, v.conso)) return hard({ key: "conso", kind: "value", id: I.conso, value: v.conso, label: "Conso EP", got: F.conso.value });
  if (!F.ges.present) { out.fatal = "champ « Émission GES » introuvable"; return out; }
  if (!diagNumEq(F.ges.value, v.ges)) return hard({ key: "ges", kind: "value", id: I.ges, value: v.ges, label: "Émission GES", got: F.ges.value });

  // 8. Cases optionnelles (bloc masqué par le serveur selon le DPE)
  for (const [key, val, label] of [["altitude", v.altitude, "Altitude > 800 m"], ["nouvelles", v.nouvelles, "Nouvelles valeurs du DPE"]]) {
    if (val === null || val === undefined) continue;
    const c = F[key];
    if (!c.present || !c.visible) continue; // signalé au contrôle final s'il fallait la cocher
    if (!!c.checked !== val && !soft({ key, kind: "check", id: I[key], checked: val, label, got: c.checked ? "coché" : "décoché" })) return out;
  }

  // 9. Lettres forcées, résultat, commentaire
  for (const [key, val, label] of [["forceEp", v.forceEp, "Lettre EP forcée"], ["forceGes", v.forceGes, "Lettre GES forcée"]]) {
    if (!val) continue;
    const el = F[key];
    if (!el.present || !el.visible) { warnOnce("abs_" + key, `champ ${label} non affiché : lettre ${val} non saisie`); continue; }
    if (diagNorm(el.value).toUpperCase() !== val && !soft({ key, kind: "value", id: I[key], value: val, label, got: el.value })) return out;
  }
  for (const [key, val, label] of [["resultat", v.resultat, "Résultat"], ["comment", v.comment, "Commentaires"]]) {
    if (!val) continue;
    const el = F[key];
    if (!el.present) { warnOnce("abs_" + key, `champ ${label} introuvable`); continue; }
    if (diagNorm(el.value) !== val && !soft({ key, kind: "value", id: I[key], value: val, label, got: el.value })) return out;
  }

  out.done = !out.soft.length && !out.hard && !out.fatal;
  return out;
}

// Relit tout le formulaire et compare aux données (contrôle final).
function diagCompare(s, v) {
  const F = s.f, out = [];
  const add = (label, want, got, ok) => out.push({ label, want, got: got ?? "", ok: !!ok });
  add("Type", v.type, F.type.value, F.type.value === v.type);
  add("Fournisseur", v.fou, F.fou.value, diagNorm(F.fou.value) === v.fou);
  add("Établi le", v.dateEtabli, F.dateReal.value, diagDateEq(F.dateReal.value, v.dateEtabli));
  if (v.dateValidite && F.dateFin.visible) add("Valable jusqu'au", v.dateValidite, F.dateFin.value, diagDateEq(F.dateFin.value, v.dateValidite));
  if (v.ademe && F.ademe.mode !== "absent") {
    const cur = F.ademe.mode === "input" ? F.ademe.input.value : F.ademe.labelText;
    add("N° ADEME", v.ademe, cur, diagNorm(cur).toUpperCase().replace(/\s+/g, "") === v.ademe);
  }
  add("Conso EP", v.conso, F.conso.value, diagNumEq(F.conso.value, v.conso));
  add("Émission GES", v.ges, F.ges.value, diagNumEq(F.ges.value, v.ges));
  if (s.couts.visible) {
    if (v.coutMin) add("Coût min", v.coutMin, F.coutMin.value, diagNumEq(F.coutMin.value, v.coutMin));
    if (v.coutMax) add("Coût max", v.coutMax, F.coutMax.value, diagNumEq(F.coutMax.value, v.coutMax));
    if (v.datePrix) add("Date d'indexation", v.datePrix, F.datePrix.value, diagDateEq(F.datePrix.value, v.datePrix));
  }
  for (const [key, val, label] of [["altitude", v.altitude, "Altitude > 800 m"], ["nouvelles", v.nouvelles, "Nouvelles valeurs"]]) {
    if (val === null || val === undefined || !F[key].present || !F[key].visible) continue;
    add(label, val ? "coché" : "décoché", F[key].checked ? "coché" : "décoché", !!F[key].checked === val);
  }
  for (const [key, val, label] of [["forceEp", v.forceEp, "Lettre EP forcée"], ["forceGes", v.forceGes, "Lettre GES forcée"]]) {
    if (val && F[key].present && F[key].visible) add(label, val, F[key].value, diagNorm(F[key].value).toUpperCase() === val);
  }
  if (v.resultat && F.resultat.present) add("Résultat", v.resultat, F.resultat.value, diagNorm(F.resultat.value) === v.resultat);
  if (v.comment && F.comment.present) add("Commentaires", v.comment, F.comment.value, diagNorm(F.comment.value) === v.comment);
  return out;
}

/* ---------------------------------------------------------------- *
 * 6. Orchestration (côté panneau, via un adaptateur)               *
 * ---------------------------------------------------------------- *
 * ad = {
 *   read():            instantané (op "read") ou null si la page est en
 *                      cours de chargement
 *   exec(cfg):         résultat d'une op de diagPageOp dans la frame du formulaire
 *   navigate(url):     { ok, error? }
 *   resolveFile({ name, ademe }): { file:{ name, path, size }, via, bytes() } | { error }
 *   pdfText(bytes):    texte du PDF ("" si illisible)
 *   confirm(info):     Promise<boolean>
 *   journalGet(key) / journalSet(key, entry|null)  (chrome.storage.session)
 *   debug(snapshot, label)  (mode debug : relevé des champs)
 *   log(text, cls), sleep(ms), now(), stopped()
 * }
 * p = { values, warnings, url, attach, desc, simulate, batch, confirmUpload,
 *       debug, timeoutMs, uploadTimeoutMs, rowLabel }
 * → { ok, status, error?, warnings[], idDiag, file, info } */
async function diagProcess(ad, p) {
  const v = p.values;
  const warnings = [];
  const res = { ok: false, status: "ERREUR", warnings, idDiag: "", file: "", info: "" };
  const T = Math.max(3000, p.timeoutMs || DIAG_DEFAULTS.timeoutMs);
  const TU = Math.max(T, p.uploadTimeoutMs || DIAG_DEFAULTS.uploadTimeoutMs);
  const desc = diagNorm(p.desc) || DIAG_DEFAULTS.desc;
  const log = (t, c) => { if (ad.log) ad.log(t, c); };
  const warn = (w) => { if (!warnings.includes(w)) { warnings.push(w); log("⚠ " + w, "skip"); } };
  const fail = (msg, status = "ERREUR") => Object.assign(res, { ok: false, status, error: msg });
  (p.warnings || []).forEach(warn);

  let lastDoc = "";
  const read = async () => {
    let s = null;
    try { s = await ad.read(); } catch (_) { s = null; }
    if (s && s.docId !== lastDoc && s.readyState === "complete") {
      lastDoc = s.docId;
      if (p.debug && ad.debug) {
        try { ad.debug(await ad.exec({ op: "dump" }), s); } catch (_) { /* relevé facultatif */ }
      }
    }
    return s;
  };
  // Attend un NOUVEAU document (postback terminé) — erreur au-delà du délai.
  const waitReload = async (docId, timeoutMs) => {
    const t0 = ad.now();
    await ad.sleep(300);
    for (;;) {
      if (ad.stopped()) return { stopped: true };
      const s = await read();
      if (s && s.docId !== docId && s.readyState === "complete") return { snap: s };
      if (ad.now() - t0 > timeoutMs) return { timeout: true };
      await ad.sleep(250);
    }
  };
  const pageError = (s) => {
    if (!s) return "page illisible";
    if (s.login && !s.isDiag) return "session SIGEO expirée (page de connexion)";
    if (!s.isDiag) return `page inattendue (${s.url}) : formulaire diag_manage introuvable`;
    return "";
  };
  const msgErrors = (s) => {
    const m = diagClassifyMessages(s.messages);
    m.infos.forEach((t) => log("SIGEO : " + t));
    return m.errors;
  };

  // — 1. Ouverture de la fiche (facultative) ; un DPE déjà enregistré dans
  //   cette session est rouvert au lieu d'en créer un second.
  if (p.url) {
    let url = p.url;
    const parts = diagUrlParts(url);
    if (parts && parts.idDiag === "0") {
      const j0 = await ad.journalGet(diagJournalKey(parts.lotId, v));
      if (j0 && j0.idDiag) {
        url = diagUrlWithId(url, j0.idDiag);
        log(`DPE déjà enregistré pendant cette session (id ${j0.idDiag}) : réouverture de la fiche au lieu d'une création`);
      }
    }
    const nav = await ad.navigate(url);
    if (!nav || !nav.ok) return fail("ouverture de la fiche impossible : " + ((nav && nav.error) || url), "PAGE");
  }

  // — 2. Vérification de la page
  let s = null;
  for (const t0 = ad.now(); ;) {
    if (ad.stopped()) return fail("arrêté par l'utilisateur", "ARRÊT");
    s = await read();
    if (s && s.readyState === "complete" && (s.isDiag || s.login)) break;
    if (ad.now() - t0 > T) break;
    await ad.sleep(250);
  }
  const pe = pageError(s);
  if (pe) return fail(pe, s && s.login ? "AUTH" : "PAGE");
  if (!s.f.type.present) return fail("liste « Type » introuvable sur la fiche", "PAGE");

  const key = diagJournalKey(s.lotId, v);
  const J = await ad.journalGet(key);
  if (s.f.ademe.mode === "label" && v.ademe && diagNorm(s.f.ademe.labelText).toUpperCase().replace(/\s+/g, "") !== v.ademe) {
    return fail(`mauvaise fiche : N° ADEME affiché ${s.f.ademe.labelText} ≠ ${v.ademe} — rien n'a été modifié`, "PAGE");
  }
  if (!s.saved && J) {
    if (J.idDiag || ["saved", "uploading", "done"].includes(J.phase)) {
      return fail(`ce DPE a déjà été enregistré pendant cette session (fiche ${J.idDiag || "id inconnu"}) : ouvre la fiche existante${J.idDiag ? " (…/" + J.idDiag + ")" : ""} ou vide l'historique DPE — aucun nouvel enregistrement`, "DOUBLON");
    }
    if (J.phase === "saving") {
      return fail("un enregistrement de ce DPE a été lancé sans confirmation (rechargement pendant l'enregistrement ?) : vérifie dans SIGEO ; s'il n'existe pas, vide l'historique DPE puis relance", "DOUBLON");
    }
  }
  log(`fiche ${s.lotId || "?"} · ${s.saved ? "diagnostic déjà enregistré (id " + (s.idDiag || "?") + ")" : "création"} · type ${s.f.type.text || s.f.type.value || "(vide)"}`);

  // — Fichier : retenu et contrôlé AVANT toute saisie (rien n'est enregistré si le PDF manque).
  let fileRef = null, bytes = null, fileChk = null;
  const already = () => diagFindDocRow(s.grid, desc);
  // N° ADEME du fichier : dans le nom, sinon dans le texte du PDF (pdf.js).
  const checkFile = async (ref) => {
    let c = diagCheckFileAdeme(fileRef.file.name, undefined, ref);
    if (c.status === "needText") {
      let text = "";
      try { text = await ad.pdfText(bytes); } catch (_) { text = ""; }
      c = diagCheckFileAdeme(fileRef.file.name, text, ref);
    }
    return c;
  };
  const resolveFile = async () => {
    const r = await ad.resolveFile({ name: v.fichier, ademe: v.ademe });
    if (!r || r.error) return (r && r.error) || "aucune source de fichier";
    bytes = await r.bytes();
    if (!diagIsPdf(bytes)) return `« ${r.file.name} » n'est pas un PDF (en-tête %PDF absent)`;
    if (bytes.length > DIAG_DEFAULTS.maxPdfBytes) return `« ${r.file.name} » trop volumineux (${diagFmtSize(bytes.length)})`;
    fileRef = r;
    fileChk = await checkFile(v.ademe);
    if (fileChk.status === "other") {
      fileRef = null;
      return `« ${r.file.name} » : le ${fileChk.where} contient un autre N° ADEME (${fileChk.others.join(", ")}) que la ligne (${v.ademe})`;
    }
    if (fileChk.status === "unknown" && p.batch && !p.confirmUpload) {
      fileRef = null;
      return `« ${r.file.name} » : N° ADEME ${v.ademe} introuvable dans le nom et dans le texte du PDF — refusé en boucle sans confirmation`;
    }
    log(`fichier retenu : ${r.file.path || r.file.name} (${r.via}, ${diagFmtSize(bytes.length)}) · ${{ name: "ADEME dans le nom", text: "ADEME dans le texte du PDF", unknown: "ADEME non vérifiable (scan ?)", none: "pas de N° ADEME" }[fileChk.status]}`);
    return "";
  };
  if (p.attach && !p.simulate && !already()) {
    const e = await resolveFile();
    if (e) return fail(`fichier DPE ${e}${s.saved ? "" : " — rien n'a été saisi"}`, "FICHIER");
  }

  // — 3. Saisie : boucle « lire → première action manquante → attendre »
  if (!s.saved && !J) await ad.journalSet(key, { phase: "filling", row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
  const st = { attempts: {}, warned: {} };
  let changed = false, ademeSeen = s.f.ademe.mode === "input";
  for (let iter = 0; ; iter++) {
    if (ad.stopped()) return fail("arrêté par l'utilisateur", "ARRÊT");
    if (iter > 40) return fail("saisie non stabilisée après 40 passes");
    const plan = diagNextActions(s, v, st);
    plan.warnings.forEach(warn);
    if (plan.fatal) return fail(plan.fatal);
    if (s.f.ademe.mode === "input") ademeSeen = true;
    if (plan.done) break;
    const acts = plan.soft.concat(plan.hard ? [Object.assign({}, plan.hard, { trigger: true })] : []);
    acts.forEach((a) => { st.attempts[a.key] = (st.attempts[a.key] || 0) + 1; });
    log("saisie : " + acts.map((a) => `${a.label} = ${a.kind === "check" ? (a.checked ? "coché" : "décoché") : "« " + a.value + " »"}`).join(", ") +
      (plan.hard ? " → rechargement" : ""));
    if (plan.hard && !s.saved) await ad.journalSet(key, { phase: "filling", step: plan.hard.key, row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
    const before = s.docId;
    let r;
    try { r = await ad.exec({ op: "set", actions: acts }); } catch (e) { r = { ok: false, error: e.message || String(e) }; }
    if (!r || !r.ok) return fail("saisie impossible : " + ((r && r.error) || "page indisponible"));
    changed = true;
    let reloaded = !!plan.hard;
    if (!plan.hard) {
      await ad.sleep(700); // un champ « doux » a-t-il quand même déclenché un postback ?
      const s2 = await read();
      if (s2 && s2.docId === before) s = s2;
      else { reloaded = true; log("rechargement inattendu après la saisie — reprise"); }
    }
    if (reloaded) {
      const w = await waitReload(before, T);
      if (w.stopped) return fail("arrêté par l'utilisateur", "ARRÊT");
      if (w.timeout) return fail(`pas de rechargement de la page après « ${plan.hard ? plan.hard.label : "saisie"} » (${Math.round(T / 1000)} s)`);
      s = w.snap;
      const e = pageError(s);
      if (e) return fail(e, s.login ? "AUTH" : "PAGE");
      const errs = msgErrors(s);
      if (errs.length) return fail("SIGEO : " + errs.join(" | "), "REFUS");
      if (plan.hard && plan.hard.key === "type" && !s.divConso) return fail("bloc des consommations non affiché après le choix du type");
      if (plan.hard && plan.hard.key === "dateReal") {
        const fin = s.f.dateFin.visible ? s.f.dateFin.value : (s.f.dateFin.label || s.f.dateFin.value);
        log(`valable jusqu'au : ${fin || "(vide)"} · champ ADEME : ${s.f.ademe.mode === "input" ? "affiché (" + s.f.ademe.via + ")" : s.f.ademe.mode === "label" ? "enregistré" : "absent"} · coûts annuels : ${s.couts.visible ? "affichés" : "masqués"}`);
      }
      if (plan.hard && plan.hard.key === "conso") log(`lettre EP : ${s.f.consoLettre.value || "(vide)"}`);
      if (plan.hard && plan.hard.key === "ges") log(`lettre GES : ${s.f.gesLettre.value || "(vide)"}`);
    }
  }

  // — Contrôle final : toutes les valeurs relues, lettres calculées
  if (!diagNorm(s.f.consoLettre.value) && !v.forceEp) return fail(`lettre EP non calculée par SIGEO (conso ${v.conso})`);
  if (!diagNorm(s.f.gesLettre.value) && !v.forceGes) return fail(`lettre GES non calculée par SIGEO (émission ${v.ges})`);
  if (!s.couts.visible && (v.coutMin || v.coutMax || v.datePrix)) warn("bloc « coûts annuels » non affiché : coûts non saisis");
  for (const [key, val, label] of [["altitude", v.altitude, "Altitude > 800 m"], ["nouvelles", v.nouvelles, "Nouvelles valeurs du DPE"]]) {
    if (val && !(s.f[key].present && s.f[key].visible)) warn(`case « ${label} » non affichée : non cochée`);
  }
  const cmp = diagCompare(s, v);
  const bad = cmp.filter((c) => !c.ok);
  if (bad.length) return fail("contrôle avant enregistrement : " + bad.map((c) => `${c.label} attendu « ${c.want} », relu « ${c.got} »`).join(" ; "));
  log(`contrôle : ${cmp.length} valeur(s) conforme(s) · lettres ${s.f.consoLettre.value || v.forceEp}/${s.f.gesLettre.value || v.forceGes}`, "ok");

  if (p.simulate) {
    return Object.assign(res, { ok: true, status: "SIMULATION", idDiag: s.idDiag !== "0" ? s.idDiag : "", info: `simulation : formulaire rempli et contrôlé (${cmp.length} valeurs), non enregistré` });
  }

  // — 4. Enregistrer (une seule fois ; jamais « Enreg. et fermer »)
  if (!s.saved || changed) {
    if (!s.saveBtn.present) return fail("bouton « Enregistrer » introuvable");
    const isNew = !s.saved;
    if (isNew) await ad.journalSet(key, { phase: "saving", row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
    const before = s.docId;
    const r = await ad.exec({ op: "save" });
    if (!r || !r.ok) return fail("Enregistrer : " + ((r && r.error) || "page indisponible"));
    log(isNew ? "Enregistrer → création du diagnostic" : "Enregistrer → mise à jour du diagnostic");
    const w = await waitReload(before, T);
    if (w.stopped) return fail("arrêté pendant l'enregistrement — vérifie dans SIGEO", "ARRÊT");
    if (w.timeout) return fail(`pas de rechargement après « Enregistrer » (${Math.round(T / 1000)} s) — vérifie dans SIGEO si le diagnostic a été créé`);
    s = w.snap;
    const e = pageError(s);
    if (e) return fail(e, s.login ? "AUTH" : "PAGE");
    const m = diagClassifyMessages(s.messages);
    m.infos.forEach((t) => log("SIGEO : " + t));
    if (m.errors.length) {
      if (isNew) await ad.journalSet(key, { phase: "filling", row: p.rowLabel || "", lotId: s.lotId, t: ad.now(), refus: m.errors.join(" | ") });
      return fail("SIGEO a refusé l'enregistrement : " + m.errors.join(" | "), "REFUS");
    }
    if (!m.success) return fail("pas de message « Enregistrement réussi ! » après Enregistrer — vérifie dans SIGEO");
    if (!s.attach.present || s.needSave) return fail("enregistré, mais le bloc « Joindre un document » n'apparaît pas");
    res.idDiag = s.idDiag && s.idDiag !== "0" ? s.idDiag : "";
    await ad.journalSet(key, { phase: "saved", idDiag: res.idDiag, url: s.url, row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
    log(`enregistrement réussi${res.idDiag ? " — diagnostic " + res.idDiag : ""}`, "ok");
  } else {
    res.idDiag = s.idDiag && s.idDiag !== "0" ? s.idDiag : "";
    if (!J) await ad.journalSet(key, { phase: "saved", idDiag: res.idDiag, url: s.url, row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
    log("diagnostic déjà enregistré et conforme : pas de nouvel enregistrement");
  }

  // — 5. Contrôles après enregistrement (avant tout upload)
  const serverAdeme = s.f.ademe.mode === "label" ? diagNorm(s.f.ademe.labelText).toUpperCase().replace(/\s+/g, "") : "";
  if (v.ademe) {
    if (serverAdeme && serverAdeme !== v.ademe) return fail(`N° ADEME enregistré par SIGEO ${serverAdeme} ≠ ${v.ademe} — upload annulé`);
    if (!serverAdeme && ademeSeen) return fail("N° ADEME non affiché par SIGEO après l'enregistrement — upload annulé");
  }
  if (serverAdeme) log(`N° ADEME confirmé par SIGEO : ${serverAdeme}`);
  const fn = s.f.fouName || {};
  if (fn.name) log(`fournisseur : ${v.fou} — ${fn.name}`);
  else if (fn.dedicated) return fail(`fournisseur ${v.fou} non reconnu par SIGEO (nom vide à côté du code) — upload annulé`);
  else warn("nom du fournisseur non repéré sur la page (active le mode debug pour le localiser)");
  const after = diagCompare(s, v).filter((c) => !c.ok);
  if (after.length) return fail("valeurs différentes après l'enregistrement : " + after.map((c) => `${c.label} « ${c.got} » au lieu de « ${c.want} »`).join(" ; ") + " — upload annulé");

  const okInfo = () => [res.idDiag ? "diagnostic " + res.idDiag : "diagnostic enregistré", serverAdeme ? "ADEME " + serverAdeme : ""].filter(Boolean).join(", ");
  if (!p.attach) return Object.assign(res, { ok: true, status: "OK", info: okInfo() + " — PDF non joint (option désactivée)" });

  // — 6. Pièce jointe (idempotente : jamais deux documents « DPE »)
  const existing = diagFindDocRow(s.grid, desc);
  if (existing) {
    if (fileRef && !diagRowHasFile(existing, fileRef.file.name)) warn(`un document « ${desc} » est déjà joint (${existing.join(" · ")}) mais pas sous le nom ${fileRef.file.name}`);
    return Object.assign(res, { ok: true, status: "DÉJÀ JOINT", info: `${okInfo()} — document « ${desc} » déjà joint, pas de nouvel upload` });
  }
  if (!fileRef) {
    const e = await resolveFile();
    if (e) return fail(`fichier DPE ${e} — diagnostic enregistré, PDF non joint`, "FICHIER");
  }
  const name = fileRef.file.name;
  const ref = serverAdeme || v.ademe;
  // Contrôle contre le N° ADEME confirmé par SIGEO (identique à celui de la ligne, sinon arrêt plus haut).
  const chk = ref === v.ademe && fileChk ? fileChk : await checkFile(ref);
  if (chk.status === "other") return fail(`le ${chk.where} de ${name} contient un autre N° ADEME (${chk.others.join(", ")}) que la fiche (${ref}) — upload annulé`, "FICHIER");
  const chkTxt = {
    name: "N° ADEME présent dans le nom du fichier",
    text: "N° ADEME trouvé dans le texte du PDF",
    unknown: "N° ADEME introuvable dans le nom et dans le texte du PDF (scan ?)",
    none: "pas de N° ADEME (DPE sans champ ADEME) : fichier retenu par son nom"
  }[chk.status];
  if (chk.status === "unknown" && p.batch && !p.confirmUpload) return fail(`${chkTxt} — upload refusé en batch sans confirmation`, "FICHIER");
  if (!p.batch || p.confirmUpload || chk.status === "unknown") {
    const ok = await ad.confirm({
      row: p.rowLabel || "", lot: v.lot || s.lotId, fiche: res.idDiag || s.idDiag || "", ademe: ref,
      ademeSrc: serverAdeme ? "confirmé par SIGEO" : (ref ? "données de la ligne" : ""),
      file: name, path: fileRef.file.path || name, size: diagFmtSize(bytes.length), via: fileRef.via,
      check: chkTxt, risky: chk.status === "unknown"
    });
    if (!ok) return fail("upload annulé (non confirmé) — diagnostic enregistré, PDF non joint", "ANNULÉ");
    // La page a pu être rechargée pendant la confirmation : on relit tout.
    s = await read();
    const e2 = pageError(s);
    if (e2) return fail(e2 + " — PDF non joint", s && s.login ? "AUTH" : "PAGE");
    if (diagFindDocRow(s.grid, desc)) return Object.assign(res, { ok: true, status: "DÉJÀ JOINT", info: `${okInfo()} — document « ${desc} » déjà joint` });
  }

  if (!s.attach.fileVisible) {
    const before = s.docId;
    const r = await ad.exec({ op: "showAttach" });
    if (!r || !r.ok) return fail("« Joindre un document » : " + ((r && r.error) || "page indisponible"));
    for (const t0 = ad.now(); ;) {
      await ad.sleep(200);
      const s2 = await read();
      if (s2 && s2.readyState === "complete" && s2.attach.fileVisible) { s = s2; break; }
      if (ad.now() - t0 > Math.min(T, 8000)) return fail("champs Description / Fichier non affichés après « Joindre un document »");
      if (s2 && s2.docId !== before && diagFindDocRow(s2.grid, desc)) return fail("un document « DPE » est apparu entre-temps — vérifie la fiche");
    }
  }
  const count0 = s.grid.count;
  await ad.journalSet(key, { phase: "uploading", idDiag: res.idDiag, count0, file: name, row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
  log(`upload : ${name} (${diagFmtSize(bytes.length)}), description « ${desc} »`);
  const before = s.docId;
  let r;
  try { r = await ad.exec({ op: "attach", desc, b64: diagToBase64(bytes), name, size: bytes.length }); }
  catch (e) { r = { ok: false, error: e.message || String(e) }; }
  if (!r || !r.ok) {
    await ad.journalSet(key, { phase: "saved", idDiag: res.idDiag, row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
    return fail("upload : " + ((r && r.error) || "page indisponible"));
  }
  const w = await waitReload(before, TU);
  if (w.stopped) return fail("arrêté pendant l'upload — vérifie la liste des documents", "ARRÊT");
  if (w.timeout) return fail(`pas de rechargement après l'envoi du PDF (${Math.round(TU / 1000)} s) — vérifie la liste des documents`);
  s = w.snap;
  const e3 = pageError(s);
  if (e3) return fail(e3, s.login ? "AUTH" : "PAGE");
  const errs = msgErrors(s);
  const row = diagFindDocRow(s.grid, desc);
  if (errs.length || s.grid.count !== count0 + 1 || !row) {
    await ad.journalSet(key, { phase: "saved", idDiag: res.idDiag, row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
    return fail(`upload non vérifié : ${count0} → ${s.grid.count} document(s), ligne « ${desc} » ${row ? "présente" : "absente"}${errs.length ? " — SIGEO : " + errs.join(" | ") : ""}`);
  }
  if (!diagRowHasFile(row, name)) warn(`nom ${name} non retrouvé dans la ligne « ${desc} » (${row.join(" · ")})`);
  await ad.journalSet(key, { phase: "done", idDiag: res.idDiag, file: name, row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
  log(`PDF joint : ${s.grid.count} document(s) — ${row.join(" · ")}`, "ok");
  return Object.assign(res, { ok: true, status: "OK", file: name, info: `${okInfo()} — PDF joint (${name})` });
}

/* ---------------------------------------------------------------- *
 * 7. Corps HTML de l'étape (panneau)                               *
 * ---------------------------------------------------------------- */
function diagStepHtml(step, esc) {
  const cols = Object.assign({}, DIAG_PRESETS.default, step.diagCols || {});
  const grid = DIAG_FIELDS.map((f) => `
        <label title="${esc(f.hint || f.k)}">${f.label}${f.req ? " *" : ""}</label>
        <input type="text" class="scn-diag-col" data-k="${f.k}" value="${esc(cols[f.k] ?? "")}" />`).join("");
  return `
      <p class="hint scn-only-diag">Remplit le formulaire « Diagnostic » SIGEO (popup <code>diag_manage</code>) : type, fournisseur, dates, N° ADEME, consommations, coûts…, contrôle toutes les valeurs, clique « Enregistrer », puis joint le PDF. Chaque postback recharge la page : l'étape attend le rechargement et reprend là où elle en est. Jamais de touche Entrée ni « Enreg. et fermer » ; la popup reste ouverte.</p>
      <div class="scn-row scn-only-diag">
        <label>URL de la fiche :</label>
        <input type="text" class="scn-diag-url" placeholder="vide = fiche déjà ouverte · ex : …/popup.aspx/fr/sig_imm/diag_manage/{ID interne}/10/0" value="${esc(step.diagUrl || "")}" />
      </div>
      <details class="scn-only-diag scn-diag-cols">
        <summary>Colonnes du fichier <span class="h2-note">(* obligatoire)</span></summary>
        <div class="scn-diag-grid">${grid}
        </div>
        <div class="scn-row">
          <button class="btn sm scn-diag-preset" data-preset="default" type="button">Noms par défaut</button>
          <button class="btn sm scn-diag-preset" data-preset="analyse" type="button" title="Colonnes écrites par l'extraction « Analyse DPE »">Colonnes de l'Analyse DPE</button>
        </div>
      </details>
      <div class="scn-row scn-only-diag">
        <label class="checkbox-row"><input type="checkbox" class="scn-diag-attach" ${step.diagAttach === false ? "" : "checked"} /> joindre le PDF</label>
        <label>description :</label>
        <input type="text" class="scn-diag-desc" style="max-width:90px" value="${esc(step.diagDesc ?? DIAG_DEFAULTS.desc)}" />
      </div>
      <div class="scn-row scn-only-diag">
        <div class="scn-diag-drop" tabindex="0" role="button" title="Fichier utilisé par « Démarrer sur la ligne active » (jamais en boucle)">PDF (ligne active) : <strong class="scn-diag-unitname">aucun</strong> — clique ou glisse le fichier ici</div>
        <input type="file" class="scn-diag-unitfile" accept="application/pdf,.pdf" hidden />
        <button class="btn icon-only scn-diag-unitclear" type="button" title="Oublier ce fichier"><svg class="icon icon-sm"><use href="#icon-close"/></svg></button>
      </div>
      <div class="scn-row scn-only-diag">
        <button class="btn sm scn-diag-folder" type="button" title="Dossier où chercher les PDF (boucle et ligne active)"><svg class="icon icon-sm"><use href="#icon-file"/></svg> Dossier des DPE</button>
        <span class="hint scn-diag-foldername">aucun dossier</span>
        <input type="file" class="scn-diag-folderinput" webkitdirectory multiple hidden />
      </div>
      <div class="scn-row scn-only-diag">
        <label class="checkbox-row" title="En boucle, chaque upload attend ta confirmation (toujours le cas sur la ligne active)"><input type="checkbox" class="scn-diag-confirm" ${step.diagConfirm === false ? "" : "checked"} /> boucle : confirmer chaque upload</label>
        <label class="checkbox-row" title="Remplit et contrôle tout, sans cliquer Enregistrer ni joindre le PDF"><input type="checkbox" class="scn-diag-simulate" ${step.diagSimulate ? "checked" : ""} /> simulation</label>
      </div>
      <div class="scn-row scn-only-diag">
        <label>résultat → colonne :</label>
        <input type="text" class="new-col-input scn-diag-resultcol" value="${esc(step.diagResultCol ?? DIAG_DEFAULTS.resultCol)}" />
        <label class="checkbox-row" title="Une ligne dont la colonne résultat commence par « OK » est sautée"><input type="checkbox" class="scn-diag-skipdone" ${step.diagSkipDone === false ? "" : "checked"} /> sauter les lignes « OK »</label>
      </div>
      <div class="scn-row scn-only-diag">
        <label>rechargement max (ms) :</label>
        <input type="number" class="scn-diag-timeout" min="3000" step="1000" value="${esc(step.diagTimeout ?? DIAG_DEFAULTS.timeoutMs)}" />
        <label>upload max (ms) :</label>
        <input type="number" class="scn-diag-uptimeout" min="5000" step="5000" value="${esc(step.diagUploadTimeout ?? DIAG_DEFAULTS.uploadTimeoutMs)}" />
      </div>
      <div class="scn-row scn-only-diag">
        <label class="checkbox-row" title="À chaque chargement de la page : liste des champs (id, libellé, visible/caché, valeur, handlers) dans la console et le relevé à copier"><input type="checkbox" class="scn-diag-debug" ${step.diagDebug ? "checked" : ""} /> debug : relever les champs à chaque chargement</label>
      </div>
      <div class="scn-row scn-only-diag">
        <button class="btn sm scn-diag-dump" type="button" title="Relève les champs de la page SIGEO ouverte, sans rien modifier">Lister les champs de la page</button>
        <button class="btn sm scn-diag-copy" type="button" title="Copie les relevés (JSON) pour les envoyer">Copier le relevé</button>
        <button class="btn sm scn-diag-forget" type="button" title="Oublie les DPE enregistrés pendant cette session (protection anti-doublon)">Vider l'historique DPE</button>
        <span class="hint scn-diag-out"></span>
      </div>`;
}

// Réglages de l'étape lus dans le DOM (null-safe).
function diagStepRead(el) {
  const q = (c) => el.querySelector(c);
  const val = (c, d = "") => (q(c) ? q(c).value : d);
  const chk = (c, d) => (q(c) ? q(c).checked : d);
  const cols = {};
  el.querySelectorAll(".scn-diag-col").forEach((i) => { cols[i.dataset.k] = i.value.trim(); });
  return {
    diagUrl: val(".scn-diag-url").trim(),
    diagCols: cols,
    diagAttach: chk(".scn-diag-attach", true),
    diagDesc: val(".scn-diag-desc", DIAG_DEFAULTS.desc).trim() || DIAG_DEFAULTS.desc,
    diagConfirm: chk(".scn-diag-confirm", true),
    diagSimulate: chk(".scn-diag-simulate", false),
    diagResultCol: val(".scn-diag-resultcol", DIAG_DEFAULTS.resultCol).trim(),
    diagSkipDone: chk(".scn-diag-skipdone", true),
    diagTimeout: parseInt(val(".scn-diag-timeout"), 10) || DIAG_DEFAULTS.timeoutMs,
    diagUploadTimeout: parseInt(val(".scn-diag-uptimeout"), 10) || DIAG_DEFAULTS.uploadTimeoutMs,
    diagDebug: chk(".scn-diag-debug", false)
  };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    DIAG_IDS, DIAG_FIELDS, DIAG_PRESETS, DIAG_DEFAULTS,
    diagNorm, diagFold, diagDate, diagDateEq, diagSerialToDate, diagExpectedValidity, diagIsNewDpe,
    diagParseNum, diagNum, diagNumEq, diagBool, diagLetter, diagAdeme, diagType, diagNormalizeData,
    diagMatchFile, diagCheckFileAdeme, diagIsPdf, diagToBase64, diagFmtSize,
    diagUrlParts, diagUrlWithId, diagJournalKey, diagClassifyMessages, diagFindDocRow, diagRowHasFile,
    diagPageOp, diagNextActions, diagCompare, diagProcess, diagStepHtml, diagStepRead
  };
}
