/* ====================================================================
 * NoHands OSA — Étape de scénario « SIGEO : saisir un DPE »
 * --------------------------------------------------------------------
 * Remplit le formulaire « Diagnostic » de SIGEO (popup ouverte en onglet :
 *   …/popup.aspx/fr/sig_imm/diag_manage/{id}/10/{idDiag}, idDiag = 0 en
 * création), clique « Enregistrer », puis joint le PDF du DPE
 * (description « DPE »).
 *
 * Principe : le side panel pilote ; chaque action est une fonction
 * autonome injectée dans la page (diagPageOp). Certains champs déclenchent
 * un __doPostBack COMPLET : le formulaire est sérialisé au submit() puis
 * toute la page est remplacée — ce qu'on écrit ensuite dans l'ancien
 * document est perdu. D'où (v2.16.0) :
 *   - un docId par document (content.js et diagPageOp, attribut data-nh-doc
 *     sur <html>) et un drapeau « navigation en cours » posé par les
 *     événements formdata / beforeunload / pagehide ;
 *   - fireAndSettle : un champ à postback à la fois, puis attente d'un
 *     nouveau document (docId ≠, readyState complete) ou de ~1,5 s de calme
 *     sans navigation en attente ; 15 s max → POSTBACK_TIMEOUT. Si un champ
 *     à postback n'a rien déclenché, son handler onchange est appelé
 *     directement dans le monde de la page (cas d'« Établi le ») ;
 *   - saisie en phases, toujours basée sur l'écart avec le DOM relu :
 *     A champs à postback (Type → Fournisseur → Établi le → Conso → GES),
 *     B champs sans postback (ADEME, dates, coûts, lettres forcées…),
 *     C contrôle avant enregistrement (2 nouvelles saisies max),
 *     D Enregistrer + contrôle après, E upload du PDF.
 *     Tout rechargement inattendu fait repartir de la phase A.
 * v2.17.0 (faits vérifiés sur SIGEO le 23/11/2026) :
 *   - « Établi le » : un change synthétique n'atteint pas
 *     performanceEnergetique ; après l'écriture, performanceEnergetique(el)
 *     est appelé directement dans le monde de la page (repli __doPostBack).
 *   - En création, l'input N° ADEME n'existe que si Type = DPE ET « Établi
 *     le » ≥ 01/07/2021 (ademeSaisissableAvantEnreg). Sinon il apparaît après
 *     le 1er enregistrement : enregistrement en deux temps (Enregistrer n°1,
 *     N° ADEME, Enregistrer n°2). Champ attendu mais absent → ADEME_CHAMP_ABSENT,
 *     rien n'est enregistré. Une fois enregistré avec une valeur, le N° est un
 *     texte dans span#body_x_labelIdDpe.
 *   - Échec après l'Enregistrer n°1 : « fiche DPE {id} créée mais incomplète »
 *     (+ lien) ; une reprise rouvre cette fiche (journal, colonne résultat ou
 *     URL avec id) en modification — jamais de second DPE.
 * v2.17.1 : SIGEO n'affiche pas toujours le PDF juste après l'upload ; si la
 *   liste ne le montre pas, « Enregistrer » est recliqué (3 fois max, 10 s
 *   d'attente active chacun) et la liste relue — JAMAIS de nouvel upload.
 *   Une relance après échec revérifie de même au lieu de ré-uploader.
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
  // Libellé du fournisseur résolu par SIGEO (id avec deux-points).
  fouLabel: "body:x:proxyFou:x:lbl:x:lbl",
  dateReal: "body_x_txtDateReal_x_txtDateReal",
  dateFin: "body_x_txtDateFinValid_x_txtDateFinValid",
  dateFinLabel: "fake_txtDateFinValid_x_txtDateFinValid",
  dateDeb: "body_x_txtDateDebValid_x_txtDateDebValid",
  ademe: "body_x_textIdDpe_x_textIdDpe",
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
  { k: "date_validite", label: "Valable jusqu'au (opt.)", hint: "jamais saisi : calculé par SIGEO ; avertissement si différent du calcul" },
  { k: "date_debut_validite", label: "Date début validité (opt.)" },
  { k: "numero_ademe", label: "N° ADEME" },
  { k: "conso_ep", label: "Conso EP (kWh/m²/an)", req: true },
  { k: "emission_ges", label: "Émission GES (kgCO2/m²/an)", req: true },
  { k: "lettre_ep_forcee", label: "Lettre EP du DPE (opt.)", hint: "mise dans « Forcer la lettre » seulement si SIGEO calcule une autre lettre (option)" },
  { k: "lettre_ges_forcee", label: "Lettre GES du DPE (opt.)", hint: "mise dans « Forcer la lettre » seulement si SIGEO calcule une autre lettre (option)" },
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
    date_validite: "{Valable jusqu'au}", date_debut_validite: "", numero_ademe: "{N° ADEME}", conso_ep: "{Énergie primaire (kWh/m²/an)}",
    emission_ges: "{GES (kg CO2/m²/an)}", lettre_ep_forcee: "{Classe énergie}", lettre_ges_forcee: "{Classe GES}",
    altitude_800m: "", nouvelles_valeurs: "", cout_min: "{Coût min (€/an)}", cout_max: "{Coût max (€/an)}",
    date_prix: "", resultat: "", commentaire: "", fichier_dpe: "{PDF analysé}", lot: "{lot}"
  }),
  // Fichier final de l'extraction « DPE Intralia » (v2.19.0) ; URL de la fiche :
  // colonne « URL fiche DPE » (DIAG_INTRALIA_URL).
  intralia: Object.freeze({
    type_local: "{Type de local}", code_fournisseur: "{N° FU}", date_etabli: "{Établi le}",
    date_validite: "{Valable jusqu'au}", date_debut_validite: "", numero_ademe: "{N° ADEME}", conso_ep: "{Énergie primaire (kWh/m²/an)}",
    emission_ges: "{GES (kg CO2/m²/an)}", lettre_ep_forcee: "{Classe énergie}", lettre_ges_forcee: "{Classe GES}",
    altitude_800m: "", nouvelles_valeurs: "", cout_min: "{Coût min (€/an)}", cout_max: "{Coût max (€/an)}",
    date_prix: "{Date d'indexation des prix}", resultat: "", commentaire: "", fichier_dpe: "{PDF analysé}", lot: ""
  })
});
const DIAG_INTRALIA_URL = "{URL fiche DPE}";

const DIAG_DEFAULTS = Object.freeze({
  desc: "DPE", resultCol: "Saisie DPE", timeoutMs: 15000, uploadTimeoutMs: 60000, maxPdfBytes: 40 * 1024 * 1024,
  settleMs: 1500 // calme sans navigation au-delà duquel un change n'a pas déclenché de postback
});

// Codes d'erreur (en tête du message, donc de la colonne résultat).
const DIAG_CODES = Object.freeze({
  timeout: "POSTBACK_TIMEOUT", before: "ECART_AVANT_ENREG", after: "ECART_APRES_ENREG",
  tooLong: "VALEUR_TROP_LONGUE", ademe: "ADEME_INVALIDE", ademeAbsent: "ADEME_CHAMP_ABSENT"
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

// DPE « nouveau format » : établi le 01/07/2021 ou après.
function diagIsNewDpe(date) {
  const ts = diagDateTs(diagDate(date).value || "");
  return Number.isFinite(ts) && ts >= Date.UTC(2021, 6, 1);
}
// Règle vérifiée sur SIGEO (23/11/2026) : en CRÉATION, l'input N° ADEME
// (#body_x_textIdDpe_x_textIdDpe) n'est généré au rechargement que si
// Type = DPE ET « Établi le » ≥ 01/07/2021. Sinon le N° ADEME ne peut être
// saisi qu'après le 1er enregistrement de la fiche.
function ademeSaisissableAvantEnreg(dateEtabli) {
  return diagIsNewDpe(dateEtabli);
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
// 2 décimales, sans espace ni unité. S'il dépasse maxLen caractères
// (maxlength du champ) : VALEUR_TROP_LONGUE — jamais d'arrondi pour faire tenir.
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
  const r = Math.round(n * 100) / 100;
  const txt = (Number.isInteger(r) ? String(r) : r.toFixed(2).replace(/0+$/, "").replace(/\.$/, "")).replace(".", ",");
  if (maxLen && txt.length > maxLen) {
    return { ok: false, code: DIAG_CODES.tooLong, error: `« ${txt} » : ${txt.length} caractères, ${maxLen} au maximum` };
  }
  return { ok: true, value: txt, num: r, rounded: Math.abs(r - n) > 1e-9 };
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
const diagAdemeKey = (s) => diagNorm(s).toUpperCase().replace(/\s+/g, "");
// N° ADEME « non communiqué » : NC, non communiqué, non trouvé, [NON EMIS
// ADEME], aucun, néant, -, 99999… → « NC » est saisi dans le champ N° ADEME
// (v2.17.7), sans erreur de format ; aucun N° ne sert à retrouver le PDF.
const DIAG_ADEME_NC = /^(n\.?\s*c\.?|non\s*communique|non\s*trouve|non\s*emis(\s*(a|l')?\s*ademe)?|non\s*enregistre|aucun|neant|sans|inconnu|-+|9{4,}|0{4,})$/;
function diagAdeme(input) {
  const raw = diagNorm(input);
  const f = diagFold(raw).replace(/^[\[(«"\s]+|[\])»"\s.]+$/g, "");
  if (raw && DIAG_ADEME_NC.test(f)) return { ok: true, value: "NC", plausible: true, nc: true, raw };
  const s = diagAdemeKey(input);
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

  // « FU » est un préfixe affiché hors du champ : retiré, comme les espaces.
  v.fou = txt("code_fournisseur").replace(/\s+/g, "").replace(/^fu/i, "");
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
  date("date_debut_validite", "dateDeb", false);
  date("date_prix", "datePrix", false);
  if (v.dateEtabli && diagDateTs(v.dateEtabli) > now) errors.push(`date d'établissement dans le futur : ${v.dateEtabli}`);
  // « Valable jusqu'au » n'est jamais saisi (SIGEO le calcule) : la valeur
  // du PDF n'est comparée au calcul que pour avertir.
  const expFin = diagExpectedValidity(v.dateEtabli);
  if (v.dateValidite && expFin && v.dateValidite !== expFin) {
    warnings.push(`« Valable jusqu'au » du DPE (${v.dateValidite}) ≠ calcul SIGEO (${expFin}) — non saisi, SIGEO garde ${expFin}`);
  }

  const a = diagAdeme(txt("numero_ademe"));
  v.ademe = a.value;
  v.ademeNc = !!a.nc;
  if (a.nc) warnings.push(`N° ADEME non communiqué (« ${a.raw} ») : « NC » saisi`);
  if (v.ademe && !a.plausible) errors.push(`${DIAG_CODES.ademe} — « ${v.ademe} » : format attendu 4 chiffres, 1 lettre, 7 chiffres, 1 lettre (ex. 2021L2000833E)`);
  if (!v.ademe && !a.nc && v.dateEtabli && diagIsNewDpe(v.dateEtabli)) warnings.push(`N° ADEME vide pour un DPE établi le ${v.dateEtabli}`);

  const num = (k, key, maxLen, required) => {
    const r = diagNum(g(k).text, g(k).cell, maxLen);
    if (!r.ok) { errors.push(r.code ? `${r.code} — ${label(k).replace(/ \(.*$/, "")} ${r.error}` : `${label(k)} : ${r.error}`); v[key] = ""; return; }
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

  // Lettres du DPE (PDF) : forcées dans SIGEO seulement si elles diffèrent
  // de la lettre calculée par SIGEO (voir diagLetterSpecs).
  for (const [k, key] of [["lettre_ep_forcee", "lettreEp"], ["lettre_ges_forcee", "lettreGes"]]) {
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
 * 1bis. Comparateur normalisé                                       *
 * ---------------------------------------------------------------- *
 * kind : num    → parse FR (espaces, virgule, « € », « kWh… » ignorés),
 *                 égalité si |a − b| < 0,005 (« 224 » = « 224,00 »)
 *        date   → jj/mm/aaaa, zéros de tête normalisés (« 1/2/2024 » = « 01/02/2024 »)
 *        ademe  → sans espaces, en majuscules
 *        letter → une lettre, majuscules ; code → sans espaces ;
 *        select → valeur exacte ; check → booléen ; text → espaces normalisés */
function diagSame(kind, got, want) {
  switch (kind) {
    case "num": return diagNumEq(got, want);
    case "date": {
      const a = diagDate(got), b = diagDate(want);
      return a.ok && b.ok && !!b.value && a.value === b.value;
    }
    case "ademe": return diagAdemeKey(got) === diagAdemeKey(want);
    case "letter": return diagNorm(got).toUpperCase() === diagNorm(want).toUpperCase();
    case "code": return String(got ?? "").replace(/\s+/g, "") === String(want ?? "").replace(/\s+/g, "");
    case "fou": return diagFouKey(got) === diagFouKey(want);
    case "check": return !!got === !!want;
    case "select": return String(got ?? "") === String(want ?? "");
    default: return diagNorm(got) === diagNorm(want);
  }
}
// Code fournisseur : SIGEO le stocke comme un nombre (zéros de tête
// supprimés : 0448957 → 448957) ; « FU » est un préfixe d'affichage.
function diagFouKey(v) {
  return String(v ?? "").trim().toUpperCase().replace(/\s+/g, "").replace(/^FU/, "").replace(/^0+/, "") || "0";
}

// Valeur lisible dans le journal et les messages d'écart.
function diagShow(kind, x) {
  if (kind === "check") return x ? "coché" : "décoché";
  return diagNorm(x);
}
// Entre guillemets (« » pour une valeur vide).
const diagQ = (x) => (x === "" || x === null || x === undefined) ? "« »" : `« ${x} »`;

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
  const a = v.ademeNc ? "" : v.ademe; // « NC » n'identifie pas un DPE
  return `${lotId || "?"}#${a || [v.dateEtabli, v.conso, v.ges].join("|")}`;
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
 *   "fire"       { action } → champ à postback : valeur + input/change
 *                  DIFFÉRÉS (la réponse repart avant le rechargement) ;
 *                  l'élément est re-résolu par id au moment d'écrire
 *   "write"      { action } → champ sans postback : valeur + input/change
 *                  immédiats, relecture, drapeau de navigation
 *   "save"       → clic différé sur « Enregistrer » (jamais « Enreg. et fermer »)
 *   "showAttach" → clic sur « Joindre un document » si le champ Fichier est masqué
 *   "attach"     { desc, b64, name, size } → Description puis fichier
 *                  (DataTransfer) ; le change part en différé (upload).
 * action = { key, label, kind: value|select|check, id | locate:"ademe", value | checked }
 * Aucune touche clavier n'est jamais simulée ; aucun élément n'est conservé
 * d'un appel à l'autre (chaque appel refait getElementById).
 */
function diagPageOp(cfg) {
  const ID = cfg.ids;
  const doc = document;
  const byId = (id) => doc.getElementById(id);
  const norm = (s) => String(s ?? "").replace(/[   ]/g, " ").replace(/[‘’]/g, "'").replace(/\s+/g, " ").trim();
  const fold = (s) => norm(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const fire = (el, t) => el.dispatchEvent(new Event(t, { bubbles: true }));

  // Identifiant du document : posé une seule fois par document (par
  // content.js au chargement, ou ici s'il n'y est pas encore), il change
  // donc à chaque rechargement. crypto.randomUUID n'existe pas sur une page
  // http (contexte non sécurisé) : repli sur getRandomValues.
  const root = doc.documentElement;
  const uuid = () => {
    try { if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") return crypto.randomUUID(); } catch (_) { /* contexte non sécurisé */ }
    const b = new Uint8Array(16);
    try { crypto.getRandomValues(b); } catch (_) { for (let i = 0; i < 16; i++) b[i] = Math.floor(Math.random() * 256); }
    b[6] = (b[6] & 0x0f) | 0x40;
    b[8] = (b[8] & 0x3f) | 0x80;
    const h = Array.from(b, (x) => x.toString(16).padStart(2, "0")).join("");
    return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}`;
  };
  if (root && !root.getAttribute("data-nh-doc")) root.setAttribute("data-nh-doc", uuid());
  const docId = root ? root.getAttribute("data-nh-doc") || "" : "";

  // Navigation en cours. formdata part de façon SYNCHRONE au submit() (le
  // formulaire vient d'être sérialisé : toute écriture ultérieure est
  // perdue) ; beforeunload / pagehide quand la page s'en va. Un écouteur
  // par monde JS (drapeau global, pas un attribut partagé).
  if (root && !window.__nhDiagLeaveHook) {
    window.__nhDiagLeaveHook = true;
    const mark = (attr) => () => { try { root.setAttribute(attr, String(Date.now())); } catch (_) { /* ignoré */ } };
    doc.addEventListener("formdata", mark("data-nh-submit"), true);
    window.addEventListener("beforeunload", mark("data-nh-leaving"), true);
    window.addEventListener("pagehide", mark("data-nh-leaving"), true);
  }
  // Un FormData construit par un script (sans navigation) ne bloque que 3 s.
  const leaving = () => {
    if (!root) return false;
    if (root.hasAttribute("data-nh-leaving")) return true;
    const t = +(root.getAttribute("data-nh-submit") || 0);
    return t > 0 && Date.now() - t < 3000;
  };

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
  // Champ à postback (détection automatique) : onchange qui appelle
  // __doPostBack / changeType / performanceEnergetique, case dont le onclick
  // le fait, ou bouton submit.
  const PB_RE = /__doPostBack|changeType|performanceEnergetique/;
  const pbOf = (el) => !!el && (PB_RE.test(el.getAttribute("onchange") || "") ||
    (/^(checkbox|radio)$/i.test(el.type || "") && PB_RE.test(el.getAttribute("onclick") || "")) ||
    (el.tagName === "INPUT" && /^(submit|image)$/i.test(el.type || "")));
  const info = (el) => el ? {
    present: true, visible: visible(el), id: el.id, value: valOf(el),
    checked: el.type === "checkbox" ? !!el.checked : undefined,
    maxLength: el.maxLength > 0 ? el.maxLength : 0, disabled: !!el.disabled, readOnly: !!el.readOnly,
    pb: pbOf(el)
  } : { present: false, visible: false, value: "", pb: false };

  const knownIds = new Set(Object.values(ID));
  const textInput = (i) => i && i.tagName === "INPUT" && /^(text|)$/i.test(i.getAttribute("type") || "") && !knownIds.has(i.id);

  // N° ADEME : id réel body_x_textIdDpe_x_textIdDpe ; à défaut (ancienne
  // page) input de la dernière cellule de la ligne « Établi le », id/name,
  // ou libellé « ADEME » ; texte seul dans span#body_x_labelIdDpe.
  function ademeEl() {
    const direct = byId(ID.ademe);
    if (direct && direct.tagName === "INPUT") return { el: direct, via: "id " + ID.ademe };
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
    const inp = a.el ? info(a.el) : null;
    return {
      // N° enregistré (texte non modifiable) prioritaire sur l'input.
      mode: labelText ? "label" : (inp && inp.visible ? "input" : (inp ? "hidden" : "absent")),
      input: inp, via: a.via, labelText, title: !!byId(ID.ademeTitle)
    };
  }

  // Nom du fournisseur affiché à côté du code.
  function supplier() {
    const inp = byId(ID.fou), btn = byId(ID.fouBtn);
    if (!inp) return { name: "", dedicated: false, src: "" };
    // Libellé officiel #body:x:proxyFou:x:lbl:x:lbl : non vide = fournisseur résolu.
    const lab = ID.fouLabel ? byId(ID.fouLabel) : null;
    if (lab) return { name: norm(lab.textContent), dedicated: true, labelFound: true, src: "#" + ID.fouLabel };
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
    let ap = null, formAction = "";
    try {
      formAction = form && form.getAttribute("action") ? new URL(form.getAttribute("action"), location.href).href : "";
      ap = formAction ? urlParts(formAction) : null;
    } catch (_) { /* ignoré */ }
    const idDiag = (up && up.idDiag !== "0" && up.idDiag) || (ap && ap.idDiag !== "0" && ap.idDiag) || (up ? up.idDiag : "");
    const g = (k) => info(byId(ID[k]));
    const f = {
      type: g("type"), fou: g("fou"), dateReal: g("dateReal"), dateFin: g("dateFin"), dateDeb: g("dateDeb"),
      conso: g("conso"), consoLettre: g("consoLettre"), ges: g("ges"), gesLettre: g("gesLettre"),
      forceEp: g("forceEp"), forceGes: g("forceGes"), altitude: g("altitude"), nouvelles: g("nouvelles"),
      coutMin: g("coutMin"), coutMax: g("coutMax"), datePrix: g("datePrix"),
      resultat: g("resultat"), comment: g("comment"), ademe: ademe(), fouName: supplier()
    };
    const sel = byId(ID.type);
    f.type.text = sel && sel.selectedIndex >= 0 && sel.options[sel.selectedIndex] ? norm(sel.options[sel.selectedIndex].text) : "";
    // « Valable jusqu'au » : input caché recalculé par performanceEnergetique, doublé d'un libellé.
    const lab = byId(ID.dateFinLabel);
    f.dateFin.label = lab ? norm(lab.textContent) : "";
    f.dateFin.shown = norm(f.dateFin.value) || f.dateFin.label;
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
      docId, leaving: leaving(), url: location.href, formAction, readyState: doc.readyState, title: doc.title,
      login: !!doc.querySelector('input[type="password"]'),
      isDiag: /\/diag_manage\//i.test(location.href) && !!byId(ID.type),
      lotId: up ? up.lotId : "", idDiag,
      saved: !!f.ademe.labelText || (!!idDiag && idDiag !== "0") || (attach.present && !needSave),
      needSave, f,
      divConso: divC ? visible(divC) : visible(byId(ID.conso)),
      couts: { present: !!divK, visible: divK ? visible(divK) : visible(byId(ID.coutMin)) },
      attach, saveBtn: { present: !!save, visible: visible(save), disabled: !!save && (!!save.disabled || save.getAttribute("aria-disabled") === "true") },
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
      if (field || tag === "button") item.postback = pbOf(el);
      if (tag === "li") item.cls = el.className;
      for (const h of ["onchange", "onclick", "onblur", "onkeydown", "onkeyup", "href"]) {
        const v = el.getAttribute(h);
        if (v && (h !== "href" || /^javascript:/i.test(v))) item[h] = v.slice(0, 120);
      }
      out.push(item);
    }
    return out;
  }

  // Élément visé par une action (toujours re-résolu).
  const target = (a) => a.locate === "ademe" ? (byId(a.id) || ademeEl().el) : byId(a.id);
  function check(a, el) {
    const lbl = a.label || a.key;
    if (!el) return { msg: `champ introuvable (${lbl})` };
    if (a.kind === "select" && !Array.from(el.options || []).some((o) => o.value === a.value)) {
      return { msg: `option « ${a.value} » absente de la liste ${lbl}` };
    }
    if (a.kind === "check" && el.type !== "checkbox") return { msg: `${lbl} n'est pas une case à cocher` };
    if (a.kind !== "check" && el.maxLength > 0 && String(a.value ?? "").length > el.maxLength) {
      return { code: "VALEUR_TROP_LONGUE", msg: `${lbl} « ${a.value} » : ${String(a.value).length} caractères, la page en accepte ${el.maxLength}` };
    }
    return null;
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
    // Fournisseur : validé par la perte du focus (jamais la touche Entrée,
    // dont le onkeydown déclenche l'enregistrement).
    if (a.blur) { try { el.blur(); } catch (_) { /* ignoré */ } fire(el, "blur"); }
  }

  if (cfg.op === "read") return snapshot();
  if (cfg.op === "dump") return Object.assign(snapshot(), { dump: dump() });

  if (cfg.op === "fire" || cfg.op === "write") {
    const a = cfg.action || {};
    const el = target(a);
    const bad = check(a, el);
    if (bad) return { ok: false, error: bad.msg, code: bad.code || "", docId };
    const cur = (e) => (a.kind === "check" ? !!e.checked : valOf(e));
    const before = cur(el);
    const pb = pbOf(el);
    if (cfg.op === "fire") {
      // Différé : la réponse repart avant un éventuel rechargement ; l'élément
      // est re-résolu par id au moment d'écrire.
      setTimeout(() => { const e2 = target(a); if (e2) apply(a, e2); }, 30);
      return { ok: true, before, pb, docId };
    }
    apply(a, el);
    const e2 = target(a);
    return { ok: true, before, after: e2 ? cur(e2) : "", pb, docId, leaving: leaving() };
  }

  if (cfg.op === "save") {
    const btn = byId(ID.save);
    if (!btn) return { ok: false, error: "bouton « Enregistrer » introuvable" };
    setTimeout(() => { const b = byId(ID.save); if (b) b.click(); }, 60);
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
    setTimeout(() => { const i2 = byId(ID.docFile) || inp; fire(i2, "change"); }, 60);
    return { ok: true, size: got.size, docId };
  }

  return { ok: false, error: "opération inconnue : " + cfg.op };
}

/* ---------------------------------------------------------------- *
 * 4bis. Repli : appel direct du handler onchange (monde de la page) *
 * ---------------------------------------------------------------- *
 * Injectée avec world: "MAIN" quand un change synthétique n'a déclenché
 * aucun postback sur un champ qui en attend un (« Établi le » : vérifié
 * sur la vraie page). Exécute une à une les instructions de l'attribut
 * onchange (ex. « ValidatorOnChange(event); performanceEnergetique(this); »),
 * chacune dans son try/catch : this → le champ, event → un Event change.
 * Seuls des appels de fonctions de la page avec des arguments simples sont
 * exécutés (pas d'eval). Autonome. */
function diagMainHandler(cfg) {
  const el = document.getElementById(cfg.id);
  if (!el) return { ok: false, error: `champ introuvable (${cfg.id})` };
  // « Établi le » (mode "date") : performanceEnergetique(el) calcule
  // « Valable jusqu'au » puis fait __doPostBack(date) ; repli : __doPostBack
  // direct (UniqueID = name, à défaut id).
  if (cfg.mode === "date") {
    const tries = [];
    const msg = (e) => String((e && e.message) || e);
    if (typeof window.performanceEnergetique === "function") {
      try { window.performanceEnergetique(el); return { ok: true, via: "performanceEnergetique" }; }
      catch (e) { tries.push("performanceEnergetique : " + msg(e)); }
    } else tries.push("performanceEnergetique absente de la page");
    if (typeof window.__doPostBack === "function") {
      try { window.__doPostBack(el.name || el.id, ""); return { ok: true, via: "__doPostBack", note: tries.join(" ; ") }; }
      catch (e) { tries.push("__doPostBack : " + msg(e)); }
    } else tries.push("__doPostBack absente de la page");
    return { ok: false, error: tries.join(" ; ") };
  }
  const code = String(el.getAttribute(cfg.attr || "onchange") || "").replace(/^\s*javascript:/i, "");
  if (!code.trim()) return { ok: false, error: "aucun handler " + (cfg.attr || "onchange") };
  const ev = new Event("change", { bubbles: true });
  // Découpe hors chaînes et hors parenthèses.
  const split = (src, sep) => {
    const out = [];
    let cur = "", q = "", depth = 0;
    for (let i = 0; i < src.length; i++) {
      const c = src[i];
      if (q) {
        cur += c;
        if (c === "\\" && i + 1 < src.length) { cur += src[++i]; continue; }
        if (c === q) q = "";
        continue;
      }
      if (c === "'" || c === '"') { q = c; cur += c; continue; }
      if (c === "(") depth++;
      if (c === ")") depth--;
      if (c === sep && depth === 0) { out.push(cur); cur = ""; continue; }
      cur += c;
    }
    out.push(cur);
    return out;
  };
  const parseArg = (a) => {
    a = a.trim();
    if (a === "this") return { v: el };
    if (a === "event") return { v: ev };
    if (/^(['"])[\s\S]*\1$/.test(a)) return { v: a.slice(1, -1).replace(/\\(.)/g, "$1") };
    if (/^-?\d+(\.\d+)?$/.test(a)) return { v: Number(a) };
    if (a === "true" || a === "false") return { v: a === "true" };
    if (a === "null") return { v: null };
    return null;
  };
  const calls = [];
  for (const raw of split(code, ";")) {
    const st = raw.trim().replace(/^return\s+/, "");
    if (!st) continue;
    const m = st.match(/^([A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)*)\s*\(([\s\S]*)\)$/);
    if (!m) { calls.push({ st, skipped: "instruction non reconnue" }); continue; }
    const path = m[1].split(".");
    let owner = window, fn = window;
    for (const k of path) { owner = fn; fn = fn == null ? undefined : fn[k]; }
    if (typeof fn !== "function") { calls.push({ st, skipped: "fonction absente" }); continue; }
    const args = [];
    let bad = false;
    for (const a of (m[2].trim() ? split(m[2], ",") : [])) {
      const p = parseArg(a);
      if (!p) { bad = true; break; }
      args.push(p.v);
    }
    if (bad) { calls.push({ st, skipped: "arguments non reconnus" }); continue; }
    try { fn.apply(path.length > 1 ? owner : window, args); calls.push({ st, ok: true }); }
    catch (e) { calls.push({ st, error: String((e && e.message) || e) }); }
  }
  return { ok: calls.some((c) => c.ok), calls };
}

/* ---------------------------------------------------------------- *
 * 5. Plan de saisie et comparaison (logique pure)                  *
 * ---------------------------------------------------------------- *
 * diagPlan(s, v, opts) → { specs[], warnings[] }, s = instantané (op read).
 * spec = { key, label, phase, kind (comparateur), act (null = non saisi),
 *          want, got, ok, present, visible, pb, note? }
 *   phase "A"    : champs à postback, ordre Type → Fournisseur → Établi le
 *                  → Conso → GES (un à la fois, fireAndSettle) ;
 *   phase "calc" : « Valable jusqu'au », calculé par SIGEO à partir
 *                  d'« Établi le » : vérifié, jamais saisi ;
 *   phase "B"    : champs sans postback (ADEME, date début validité, coûts,
 *                  date des prix, lettres forcées, cases, résultat, commentaire).
 * Les lettres EP/GES ne sont jamais comparées au PDF : la lettre du DPE va
 * dans « Forcer la lettre » seulement si SIGEO en calcule une autre
 * (opts.forceLetters, défaut oui) — c'est ce champ-là qui est vérifié. */
const DIAG_PHASE_A = Object.freeze(["type", "fou", "dateReal", "conso", "ges"]);

function diagPlan(s, v, opts = {}) {
  const I = DIAG_IDS, F = s.f || {};
  const specs = [], warnings = [];
  const fld = (k) => F[k] || { present: false, visible: false, value: "" };
  const add = (o) => {
    const f = o.field || fld(o.key);
    const sp = Object.assign({ present: !!f.present, visible: !!f.visible, pb: !!f.pb, note: "" }, o);
    delete sp.field;
    sp.ok = diagSame(sp.kind, sp.got, sp.want);
    specs.push(sp);
    return sp;
  };
  const val = (k) => fld(k).value;

  // — Phase A (postback) + « Valable jusqu'au » calculé
  add({ key: "type", label: "Type", phase: "A", kind: "select", act: { kind: "select", id: I.type }, want: v.type, got: val("type") });
  add({ key: "fou", label: "Fournisseur", phase: "A", kind: "fou", act: { kind: "value", id: I.fou, blur: true }, want: v.fou, got: val("fou") });
  add({ key: "dateReal", label: "Établi le", phase: "A", kind: "date", act: { kind: "value", id: I.dateReal }, want: v.dateEtabli, got: val("dateReal") });
  const exp = diagExpectedValidity(v.dateEtabli);
  const fin = fld("dateFin");
  if (exp && (fin.present || fin.label)) {
    add({ key: "dateFin", label: "Valable jusqu'au", phase: "calc", kind: "date", act: null, want: exp,
      got: fin.shown ?? fin.value, field: Object.assign({}, fin, { present: true, visible: true, pb: false }) });
  } else if (exp) warnings.push("« Valable jusqu'au » introuvable sur la page : non vérifié");
  add({ key: "conso", label: "Conso EP", phase: "A", kind: "num", act: { kind: "value", id: I.conso }, want: v.conso, got: val("conso") });
  add({ key: "ges", label: "Émission GES", phase: "A", kind: "num", act: { kind: "value", id: I.ges }, want: v.ges, got: val("ges") });

  // — Phase B (sans postback)
  // N° ADEME : lu dans span#body_x_labelIdDpe s'il existe (enregistré), sinon
  // dans l'input. opts.ademeDeferred (DPE < 01/07/2021 en création) : saisi
  // après le 1er enregistrement, absent du plan d'ici là. Champ absent : écart
  // non saisissable (ADEME_CHAMP_ABSENT est levé avant, par l'orchestrateur).
  const A = F.ademe || { mode: "absent" };
  if (v.ademe && !opts.ademeDeferred) {
    if (A.mode === "label") {
      add({ key: "ademe", label: "N° ADEME", phase: "B", kind: "ademe", act: null, want: v.ademe, got: A.labelText,
        field: { present: true, visible: true, pb: false } });
    } else if (A.mode === "input") {
      add({ key: "ademe", label: "N° ADEME", phase: "B", kind: "ademe", act: { kind: "value", locate: "ademe", id: A.input.id },
        want: v.ademe, got: A.input.value, field: A.input });
    } else {
      add({ key: "ademe", label: "N° ADEME", phase: "B", kind: "ademe", act: null, want: v.ademe, got: "", missing: true,
        field: { present: false, visible: false, pb: false } });
    }
  } else if (!v.ademe && !v.ademeNc && A.mode === "input" && diagIsNewDpe(v.dateEtabli)) warnings.push("champ N° ADEME affiché mais N° ADEME vide dans les données");

  if (v.dateDeb) {
    const f = fld("dateDeb");
    if (f.present && f.visible) add({ key: "dateDeb", label: "Date début validité", phase: "B", kind: "date", act: { kind: "value", id: I.dateDeb }, want: v.dateDeb, got: f.value });
    else warnings.push(`champ « Date début validité » ${f.present ? "masqué" : "absent"} : ${v.dateDeb} non saisie`);
  }

  const couts = [["coutMin", v.coutMin, "Coût min", "num"], ["coutMax", v.coutMax, "Coût max", "num"], ["datePrix", v.datePrix, "Date d'indexation des prix", "date"]];
  if (s.couts && s.couts.visible) {
    for (const [key, want, label, kind] of couts) {
      // Date d'indexation des prix : facultative (non écrite par l'Analyse DPE).
      if (!want) { if (key !== "datePrix") warnings.push(`bloc « coûts annuels » affiché mais ${label} vide dans les données`); continue; }
      const f = fld(key);
      if (!f.present) { warnings.push(`champ ${label} introuvable dans le bloc coûts`); continue; }
      add({ key, label, phase: "B", kind, act: { kind: "value", id: I[key] }, want, got: f.value });
    }
  } else if (couts.some((c) => c[1])) warnings.push("bloc « coûts annuels » non affiché : coûts non saisis");

  for (const [key, want, legend, L] of [["forceEp", v.lettreEp, "consoLettre", "EP"], ["forceGes", v.lettreGes, "gesLettre", "GES"]]) {
    if (!want) continue;
    const f = fld(key);
    const sigeo = diagNorm(val(legend)).toUpperCase();
    const forced = diagNorm(f.value).toUpperCase();
    let target = null, note = "";
    if (forced === want) target = want;                    // déjà forcée à la lettre du DPE
    else if (sigeo === want) { if (forced) target = ""; }  // SIGEO calcule la bonne lettre : on retire un forçage contraire
    else if (opts.forceLetters !== false) {
      target = want;
      note = `lettre ${L} du DPE ${want} ≠ lettre calculée par SIGEO ${sigeo || "(vide)"} → forcée`;
    } else warnings.push(`lettre ${L} du DPE ${want} ≠ lettre calculée par SIGEO ${sigeo || "(vide)"} — non forcée (option désactivée)`);
    if (target === null) continue;
    if (!f.present || !f.visible) { warnings.push(`champ « Forcer la lettre ${L} » non affiché : lettre ${want} non forcée`); continue; }
    add({ key, label: `Lettre ${L} forcée`, phase: "B", kind: "letter", act: { kind: "value", id: I[key] }, want: target, got: f.value, note });
  }

  for (const [key, want, label] of [["altitude", v.altitude, "Altitude > 800 m"], ["nouvelles", v.nouvelles, "Nouvelles valeurs du DPE"]]) {
    if (want === null || want === undefined) continue;
    const f = fld(key);
    if (!f.present || !f.visible) { if (want) warnings.push(`case « ${label} » non affichée : non cochée`); continue; }
    add({ key, label, phase: "B", kind: "check", act: { kind: "check", id: I[key] }, want, got: !!f.checked });
  }
  for (const [key, want, label] of [["resultat", v.resultat, "Résultat"], ["comment", v.comment, "Commentaires"]]) {
    if (!want) continue;
    const f = fld(key);
    if (!f.present) { warnings.push(`champ ${label} introuvable`); continue; }
    add({ key, label, phase: "B", kind: "text", act: { kind: "value", id: I[key] }, want, got: f.value });
  }
  return { specs, warnings };
}

// Comparaison complète (contrôles avant / après enregistrement).
function diagCompare(s, v, opts = {}) {
  return diagPlan(s, v, opts).specs.map((x) => ({ key: x.key, label: x.label, kind: x.kind, want: x.want, got: x.got ?? "", ok: x.ok }));
}

// « 2 écart(s) : N° ADEME attendu « … », relu « … » ; … » (TOUS les écarts).
function diagGapText(gaps) {
  return `${gaps.length} écart(s) : ` + gaps.map((g) => `${g.label} attendu ${diagQ(diagShow(g.kind, g.want))}, relu ${diagQ(diagShow(g.kind, g.got))}${g.missing ? " (champ absent)" : ""}`).join(" ; ");
}

/* ---------------------------------------------------------------- *
 * 6. Attente après un déclenchement (fireAndSettle)                 *
 * ---------------------------------------------------------------- *
 * Interroge la frame jusqu'à :
 *   a) un NOUVEAU document (docId ≠ docBefore) en readyState complete
 *      → { reloaded: true, snap } ;
 *   b) ou quietMs (~1,5 s) d'affilée sans navigation en attente — même
 *      document, chargé, pas de drapeau formdata/beforeunload, onglet pas
 *      en chargement — → { reloaded: false, snap } (pas de postback) ;
 *      jamais si o.mustReload (Enregistrer, upload) ;
 *   c) au-delà de timeoutMs → { timeout: true } ;
 *   o.doneWhen(snap) : fin acceptée aussi sur le MÊME document (mise à jour
 *      partielle d'un update panel) → { reloaded: false, partial: true }.
 * read() renvoie null pendant que la page est remplacée ; une lecture qui
 * ne répond pas (frame en pleine navigation) est abandonnée après
 * o.readCapMs et comptée comme « navigation en cours », pour que le délai
 * global s'applique toujours. o.ignoreLoading : l'onglet était déjà
 * « loading » avant le déclenchement (non significatif). */
function diagCapped(promise, ms) {
  return new Promise((resolve) => {
    let done = false;
    const end = (x) => { if (!done) { done = true; clearTimeout(t); resolve(x); } };
    const t = setTimeout(() => end(null), ms);
    Promise.resolve(promise).then(end, () => end(null));
  });
}
async function diagSettle(ad, read, docBefore, o = {}) {
  const quietMs = o.quietMs ?? DIAG_DEFAULTS.settleMs;
  const timeoutMs = o.timeoutMs ?? DIAG_DEFAULTS.timeoutMs;
  const capMs = o.readCapMs ?? 2500;
  const t0 = ad.now();
  let quietSince = null, last = null, sawLeaving = false;
  for (;;) {
    if (ad.stopped()) return { stopped: true };
    let s = null;
    try { s = await diagCapped(read(), capMs); } catch (_) { s = null; }
    if (s) last = s;
    if (s && s.docId && s.docId !== docBefore && s.readyState === "complete") return { reloaded: true, snap: s, sawLeaving };
    if (s && s.leaving) sawLeaving = true;
    if (o.doneWhen && s && s.docId === docBefore && s.readyState === "complete" && !s.leaving && o.doneWhen(s)) {
      return { reloaded: false, partial: true, snap: s, sawLeaving };
    }
    let loading = false;
    if (!o.ignoreLoading && ad.tabLoading) { try { loading = !!(await ad.tabLoading()); } catch (_) { loading = false; } }
    const pending = !s || s.readyState !== "complete" || s.docId !== docBefore || !!s.leaving || loading;
    if (!pending && !o.mustReload) {
      if (quietSince === null) quietSince = ad.now();
      if (ad.now() - quietSince >= quietMs) return { reloaded: false, snap: s, sawLeaving };
    } else quietSince = null;
    if (ad.now() - t0 > timeoutMs) return { timeout: true, snap: last, sawLeaving };
    await ad.sleep(o.pollMs ?? 150);
  }
}

/* ---------------------------------------------------------------- *
 * 7. Orchestration (côté panneau, via un adaptateur)               *
 * ---------------------------------------------------------------- *
 * ad = {
 *   read():            instantané (op "read") ou null si la page est en
 *                      cours de chargement
 *   exec(cfg):         résultat d'une op de diagPageOp dans la frame du formulaire
 *   execMain(cfg):     diagMainHandler dans le monde de la page (facultatif)
 *   tabLoading():      l'onglet est-il en chargement ? (facultatif)
 *   navigate(url):     { ok, error? }
 *   resolveFile({ name, ademe }): { file:{ name, path, size }, via, bytes() } | { error }
 *   pdfText(bytes):    texte du PDF ("" si illisible)
 *   confirm(info):     Promise<boolean>
 *   journalGet(key) / journalSet(key, entry|null)  (chrome.storage.session)
 *   debug(snapshot, label)  (mode debug : relevé des champs)
 *   log(text, cls), sleep(ms), now(), stopped()
 * }
 * p = { values, warnings, url, attach, desc, simulate, batch, confirmUpload,
 *       forceLetters, debug, timeoutMs, uploadTimeoutMs, settleMs, rowLabel,
 *       knownIdDiag (fiche déjà créée d'après la colonne résultat) }
 * → { ok, status, code, error?, warnings[], idDiag, file, info } */
async function diagProcess(ad, p) {
  const v = p.values;
  const opts = { forceLetters: p.forceLetters !== false, ademeDeferred: false };
  const warnings = [];
  const res = { ok: false, status: "ERREUR", code: "", warnings, idDiag: "", file: "", info: "" };
  const T = Math.max(3000, p.timeoutMs || DIAG_DEFAULTS.timeoutMs);
  const TU = Math.max(T, p.uploadTimeoutMs || DIAG_DEFAULTS.uploadTimeoutMs);
  const QUIET = p.settleMs || DIAG_DEFAULTS.settleMs;
  const desc = diagNorm(p.desc) || DIAG_DEFAULTS.desc;
  const sec = (ms) => Math.round(ms / 1000);
  const log = (t, c) => { if (ad.log) ad.log(t, c); };
  const warn = (w) => { if (!warnings.includes(w)) { warnings.push(w); log("⚠ " + w, "skip"); } };
  // Fiche enregistrée pendant ce passage (Enregistrer n°1 réussi) : tout
  // échec ultérieur le signale, avec le lien, pour une reprise sans doublon.
  let saved1 = null, ademeConfirmed = false, pdfDone = false, pdfSent = false;
  const fail = (msg, status = "ERREUR", code = "") => {
    const c = code || (String(msg).match(/^([A-Z_]{8,}) — /) || [])[1] || "";
    if (saved1) {
      const miss = [v.ademe && !ademeConfirmed ? "ADEME" : "", p.attach && !pdfDone && !pdfSent ? "PDF" : ""].filter(Boolean);
      let what = miss.length ? `${miss.join("/")} manquant${miss.length > 1 ? "s" : ""}` : "";
      if (p.attach && !pdfDone && pdfSent) what = (what ? what + ", " : "") + "PDF envoyé mais non vérifié";
      const note = `fiche DPE ${saved1.id || "(id introuvable)"} ${saved1.isNew ? "créée" : "mise à jour"} mais incomplète` +
        `${what ? ` (${what})` : ""} : ${saved1.url}`;
      log(note, "skip");
      msg = `${msg} — ${note}`;
    }
    return Object.assign(res, { ok: false, status, code: c, error: msg });
  };
  const E = (error, status = "ERREUR", code = "") => ({ error, status, code });
  (p.warnings || []).forEach(warn);

  let lastDoc = "";
  // Lecture bornée : pendant une navigation la frame peut ne pas répondre.
  const readCap = Math.min(T, 2500);
  const read = async () => {
    let s = null;
    try { s = await diagCapped(ad.read(), readCap); } catch (_) { s = null; }
    if (s && s.docId !== lastDoc && s.readyState === "complete") {
      lastDoc = s.docId;
      if (p.debug && ad.debug) {
        try { ad.debug(await ad.exec({ op: "dump" }), s); } catch (_) { /* relevé facultatif */ }
      }
    }
    return s;
  };
  const tabLoading = async () => {
    if (!ad.tabLoading) return false;
    try { return !!(await ad.tabLoading()); } catch (_) { return false; }
  };
  const settle = (docBefore, o = {}) => diagSettle(ad, read, docBefore, Object.assign({ quietMs: QUIET, timeoutMs: T }, o));
  // Document chargé, sans navigation en attente.
  const stable = async () => {
    const t0 = ad.now();
    for (;;) {
      if (ad.stopped()) return { stopped: true };
      const s = await read();
      if (s && s.readyState === "complete" && !s.leaving) return { snap: s };
      if (ad.now() - t0 > T) return { timeout: true };
      await ad.sleep(150);
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

  // — 1. Ouverture de la fiche (facultative). Un DPE déjà créé (journal de
  //   session, ou colonne résultat « fiche DPE {id} créée mais incomplète »)
  //   est rouvert en modification : jamais de second DPE. Une URL qui porte
  //   déjà un idDiag est ouverte telle quelle (modification).
  const known = String(p.knownIdDiag || "").replace(/\D/g, "");
  if (p.url) {
    let url = p.url;
    const parts = diagUrlParts(url);
    if (parts && parts.idDiag === "0") {
      const j0 = await ad.journalGet(diagJournalKey(parts.lotId, v));
      if (j0 && j0.idDiag) {
        url = diagUrlWithId(url, j0.idDiag);
        log(`DPE déjà enregistré pendant cette session (id ${j0.idDiag}) : réouverture de la fiche au lieu d'une création`);
      } else if (known) {
        url = diagUrlWithId(url, known);
        log(`fiche DPE ${known} déjà créée (colonne résultat) : réouverture en modification au lieu d'une création`);
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
  const shownAdeme = s.f.ademe.mode === "label" ? s.f.ademe.labelText : (s.saved && s.f.ademe.input ? s.f.ademe.input.value : "");
  if (v.ademe && diagAdemeKey(shownAdeme) && diagAdemeKey(shownAdeme) !== v.ademe) {
    return fail(`mauvaise fiche : N° ADEME affiché ${diagAdemeKey(shownAdeme)} ≠ ${v.ademe} — rien n'a été modifié`, "PAGE");
  }
  if (!s.saved && known) {
    return fail(`la colonne résultat indique la fiche DPE ${known} déjà créée : ouvre-la (…/${known}) au lieu d'en créer une autre — aucun nouvel enregistrement`, "DOUBLON");
  }
  if (!s.saved && J) {
    if (J.idDiag || ["saved", "uploading", "uploaded", "done"].includes(J.phase)) {
      return fail(`ce DPE a déjà été enregistré pendant cette session (fiche ${J.idDiag || "id inconnu"}) : ouvre la fiche existante${J.idDiag ? " (…/" + J.idDiag + ")" : ""} ou vide l'historique DPE — aucun nouvel enregistrement`, "DOUBLON");
    }
    if (J.phase === "saving") {
      return fail("un enregistrement de ce DPE a été lancé sans confirmation (rechargement pendant l'enregistrement ?) : vérifie dans SIGEO ; s'il n'existe pas, vide l'historique DPE puis relance", "DOUBLON");
    }
  }
  log(`fiche ${s.lotId || "?"} · ${s.saved ? "diagnostic déjà enregistré (id " + (s.idDiag || "?") + ") : modification" : "création"} · type ${s.f.type.text || s.f.type.value || "(vide)"}`);

  // N° ADEME : saisissable avant l'enregistrement seulement pour un DPE
  // ≥ 01/07/2021 ; sinon (création) après le 1er enregistrement.
  const ademeAvant = ademeSaisissableAvantEnreg(v.dateEtabli);
  opts.ademeDeferred = !!v.ademe && !s.saved && !ademeAvant;
  if (opts.ademeDeferred) {
    log(p.simulate
      ? "simulation : DPE avant 01/07/2021 — N° ADEME non tenté (il n'est saisissable qu'après le 1er enregistrement)"
      : "ADEME saisi après 1er enregistrement (DPE avant 01/07/2021)");
  }

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
  // N° servant à retrouver / contrôler le PDF : aucun pour « NC ».
  const fileAdeme = v.ademeNc ? "" : v.ademe;
  const resolveFile = async () => {
    const r = await ad.resolveFile({ name: v.fichier, ademe: fileAdeme });
    if (!r || r.error) return (r && r.error) || "aucune source de fichier";
    bytes = await r.bytes();
    if (!diagIsPdf(bytes)) return `« ${r.file.name} » n'est pas un PDF (en-tête %PDF absent)`;
    if (bytes.length > DIAG_DEFAULTS.maxPdfBytes) return `« ${r.file.name} » trop volumineux (${diagFmtSize(bytes.length)})`;
    fileRef = r;
    fileChk = await checkFile(fileAdeme);
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

  // — 3. Saisie en phases (toujours d'après l'écart avec le DOM relu)
  if (!s.saved && !J) await ad.journalSet(key, { phase: "filling", row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
  const st = { attempts: {}, restarts: 0, cRounds: 0 };
  let changed = false;
  const actOf = (sp) => Object.assign({ key: sp.key, label: sp.label, value: sp.want, checked: !!sp.want }, sp.act);
  const gotAfter = (snap, sp) => {
    const x = snap ? diagPlan(snap, v, opts).specs.find((y) => y.key === sp.key) : null;
    return x ? diagShow(x.kind, x.got) : "(champ absent)";
  };

  // fireAndSettle : UN champ à postback, puis rechargement complet ou calme.
  const fireAndSettle = async (sp, why) => {
    const w0 = await stable();
    if (w0.stopped) return E("arrêté par l'utilisateur", "ARRÊT");
    if (w0.timeout) return E(`${DIAG_CODES.timeout} — page toujours en rechargement avant « ${sp.label} » (${sec(T)} s)`, "ERREUR", DIAG_CODES.timeout);
    const s0 = w0.snap;
    const pe0 = pageError(s0);
    if (pe0) return E(pe0, s0.login ? "AUTH" : "PAGE");
    const loading0 = await tabLoading();
    if (sp.note) log(sp.note);
    // « Établi le » : un change synthétique n'atteint pas performanceEnergetique
    // (vérifié) → écriture synchrone puis appel direct dans le monde de la page.
    const isDate = sp.key === "dateReal" && !!ad.execMain;
    let r;
    try { r = await ad.exec({ op: isDate ? "write" : "fire", action: actOf(sp) }); } catch (e) { r = { ok: false, error: e.message || String(e) }; }
    if (!r || !r.ok) {
      if (r && r.code) return E(`${r.code} — ${r.error}`, "ERREUR", r.code);
      return E(`saisie impossible (${sp.label}) : ${(r && r.error) || "page indisponible"}`);
    }
    changed = true;
    let how = "";
    if (isDate && !r.leaving) {
      let m;
      try { m = await ad.execMain({ mode: "date", id: sp.act.id }); } catch (e) { m = { ok: false, error: e.message || String(e) }; }
      if (m && m.ok) {
        how = ` (${m.via} appelé)`;
        log(`${sp.label} : ${m.via}() appelé dans la page${m.via === "__doPostBack" ? " — repli" : ""}${m.note ? " (" + m.note + ")" : ""}`);
      } else log(`${sp.label} : déclenchement dans la page impossible — ${(m && m.error) || "résultat vide"}`, "skip");
    }
    let w = await settle(s0.docId, { ignoreLoading: loading0 });
    if (!isDate && !w.stopped && !w.timeout && !w.reloaded && r.pb && ad.execMain) {
      // Champ à postback resté sans effet : le change synthétique n'a pas
      // atteint le postback (« Établi le ») → appel direct du handler.
      log(`${sp.label} : pas de postback après le change synthétique → appel direct du handler onchange`, "skip");
      let m;
      try { m = await ad.execMain({ id: (sp.act && sp.act.id) || "" }); } catch (e) { m = { ok: false, error: e.message || String(e) }; }
      if (m && m.calls) log(`handler : ${m.calls.map((c) => `${c.st.slice(0, 60)} ${c.ok ? "✓" : c.error ? "✗ " + c.error : "— " + c.skipped}`).join(" · ")}`, "skip");
      else if (m && m.error) log(`handler : ${m.error}`, "skip");
      w = await settle(s0.docId, { ignoreLoading: loading0 });
      how = " (handler appelé)";
    }
    if (w.stopped) return E("arrêté par l'utilisateur", "ARRÊT");
    if (w.timeout) return E(`${DIAG_CODES.timeout} — « ${sp.label} » : rechargement non terminé après ${sec(T)} s`, "ERREUR", DIAG_CODES.timeout);
    const s1 = w.snap;
    if (w.reloaded) log(`↻ postback (${sp.label})`);
    log(`${sp.label} : attendu ${diagQ(diagShow(sp.kind, sp.want))} · avant ${diagQ(diagShow(sp.kind, r.before))} · après ${diagQ(gotAfter(s1, sp))}` +
      `${w.reloaded ? "" : " — sans postback" + how}${why ? " — " + why : ""}`);
    if (sp.key === "dateReal" && s1) {
      const fin = s1.f.dateFin, got = fin.shown ?? fin.value, exp = diagExpectedValidity(v.dateEtabli);
      log(`Valable jusqu'au : relu ${diagQ(diagNorm(got))} · calcul attendu ${exp || "?"} ${exp && diagSame("date", got, exp) ? "✓" : "✗"}`);
    }
    if (w.reloaded) {
      const pe1 = pageError(s1);
      if (pe1) return E(pe1, s1.login ? "AUTH" : "PAGE");
      const errs = msgErrors(s1);
      if (errs.length) return E("SIGEO : " + errs.join(" | "), "REFUS");
    }
    return { snap: s1, reloaded: !!w.reloaded };
  };

  // Phase A : champs à postback, un par un, dans l'ordre ; on repart
  // toujours du DOM relu (un postback peut en défaire un autre).
  const phaseA = async () => {
    for (let guard = 0; guard < 40; guard++) {
      const w = await stable();
      if (w.stopped) return E("arrêté par l'utilisateur", "ARRÊT");
      if (w.timeout) return E(`${DIAG_CODES.timeout} — page toujours en rechargement (${sec(T)} s)`, "ERREUR", DIAG_CODES.timeout);
      s = w.snap;
      const pe1 = pageError(s);
      if (pe1) return E(pe1, s.login ? "AUTH" : "PAGE");
      const plan = diagPlan(s, v, opts);
      const gap = plan.specs.find((x) => !x.ok && (x.phase === "A" || (x.phase === "calc" && (st.attempts.dateFin || 0) < 2)));
      if (!gap) return { snap: s };
      if (gap.key !== "type" && gap.key !== "fou" && !s.divConso) return E("bloc des consommations (div#div_conso) non affiché alors que le type est sélectionné");
      let sp = gap, why = "";
      if (gap.phase === "calc") {
        // « Valable jusqu'au » ≠ calcul : « Établi le » n'a pas été pris en compte → nouveau déclenchement.
        st.attempts.dateFin = (st.attempts.dateFin || 0) + 1;
        sp = plan.specs.find((x) => x.key === "dateReal");
        why = `« Valable jusqu'au » = ${diagQ(diagShow("date", gap.got))} au lieu de ${gap.want} : nouveau déclenchement`;
      } else if (!gap.present) return E(`champ « ${gap.label} » introuvable sur la fiche`, "PAGE");
      const n = st.attempts[sp.key] = (st.attempts[sp.key] || 0) + 1;
      if (n > 3) {
        return E(`${DIAG_CODES.before} — la page ne retient pas ${sp.label} : saisi ${n - 1} fois ${diagQ(diagShow(sp.kind, sp.want))}, relu ${diagQ(diagShow(sp.kind, sp.got))}`, "ERREUR", DIAG_CODES.before);
      }
      const r = await fireAndSettle(sp, why);
      if (r.error) return r;
      if (sp.key === "type" && !r.snap.divConso) return E("bloc des consommations non affiché après le choix du type");
      if (sp.key === "conso" && r.reloaded) log(`lettre EP calculée par SIGEO : ${diagNorm(r.snap.f.consoLettre.value) || "(vide)"}`);
      if (sp.key === "ges" && r.reloaded) log(`lettre GES calculée par SIGEO : ${diagNorm(r.snap.f.gesLettre.value) || "(vide)"}`);
    }
    return E("saisie non stabilisée (phase A : 40 passes)");
  };

  // Phase B : champs sans postback, SEULEMENT après la phase A. Valeur +
  // input/change ; tout rechargement observé → reprise en phase A.
  const phaseB = async () => {
    const w = await stable();
    if (w.stopped) return E("arrêté par l'utilisateur", "ARRÊT");
    if (w.timeout) return E(`${DIAG_CODES.timeout} — page toujours en rechargement (${sec(T)} s)`, "ERREUR", DIAG_CODES.timeout);
    s = w.snap;
    let docRef = s.docId, wrote = false;
    for (const sp of diagPlan(s, v, opts).specs.filter((x) => x.phase === "B" && !x.ok && x.act)) {
      const cur = await read();
      if (!cur || cur.docId !== docRef || cur.leaving || cur.readyState !== "complete") return { restart: `avant « ${sp.label} »` };
      if (sp.pb) {
        // Détecté comme champ à postback (onchange) : traité comme en phase A.
        const n = st.attempts[sp.key] = (st.attempts[sp.key] || 0) + 1;
        if (n > 3) return E(`${DIAG_CODES.before} — la page ne retient pas ${sp.label} : saisi ${n - 1} fois ${diagQ(diagShow(sp.kind, sp.want))}, relu ${diagQ(diagShow(sp.kind, sp.got))}`, "ERREUR", DIAG_CODES.before);
        const r = await fireAndSettle(sp, "champ à postback détecté");
        if (r.error) return r;
        if (r.reloaded) return { restart: `après « ${sp.label} »`, expected: true };
        docRef = r.snap.docId;
        continue;
      }
      if (sp.note) log(sp.note);
      let r;
      try { r = await ad.exec({ op: "write", action: actOf(sp) }); } catch (e) { r = { ok: false, error: e.message || String(e) }; }
      if (!r || !r.ok) {
        if (r && r.code) return E(`${r.code} — ${r.error}`, "ERREUR", r.code);
        const again = await read();
        if (!again || again.docId !== docRef) return { restart: `pendant « ${sp.label} »` };
        return E(`saisie impossible (${sp.label}) : ${(r && r.error) || "page indisponible"}`);
      }
      changed = true;
      wrote = true;
      log(`${sp.label} : attendu ${diagQ(diagShow(sp.kind, sp.want))} · avant ${diagQ(diagShow(sp.kind, r.before))} · après ${diagQ(diagShow(sp.kind, r.after))}`);
      if (r.leaving || (r.docId && r.docId !== docRef)) return { restart: `après « ${sp.label} »` };
    }
    if (!wrote) return { snap: s };
    // Aucun champ « sans postback » ne doit avoir rechargé la page.
    const w2 = await settle(docRef);
    if (w2.stopped) return E("arrêté par l'utilisateur", "ARRÊT");
    if (w2.timeout) return E(`${DIAG_CODES.timeout} — rechargement non terminé après la phase B (${sec(T)} s)`, "ERREUR", DIAG_CODES.timeout);
    if (w2.reloaded) return { restart: "après la phase B" };
    s = w2.snap;
    return { snap: s };
  };

  let finalPlan = null;
  for (;;) {
    if (ad.stopped()) return fail("arrêté par l'utilisateur", "ARRÊT");
    if (st.restarts > 4) return fail("saisie non stabilisée : trop de rechargements inattendus");
    const a = await phaseA();
    if (a.error) return fail(a.error, a.status, a.code);
    // Après le rechargement de la date, l'input N° ADEME DOIT exister pour un
    // DPE ≥ 01/07/2021 (et pour une fiche déjà enregistrée sans N°) : sinon
    // ADEME_CHAMP_ABSENT et aucun enregistrement.
    const am = a.snap.f.ademe.mode;
    if (!opts.ademeDeferred && am !== "input" && am !== "label" && (v.ademe || (!a.snap.saved && ademeAvant))) {
      return fail(`${DIAG_CODES.ademeAbsent} — champ N° ADEME (#${DIAG_IDS.ademe}) ${am === "hidden" ? "masqué" : "absent"} après le rechargement de la date` +
        `${!a.snap.saved && ademeAvant ? ` alors que le DPE est établi le ${v.dateEtabli} (≥ 01/07/2021) — type DPE et date bien pris en compte ?` : ""} — rien n'a été enregistré`,
        "ERREUR", DIAG_CODES.ademeAbsent);
    }
    const b = await phaseB();
    if (b.error) return fail(b.error, b.status, b.code);
    if (b.restart !== undefined) {
      if (b.expected) log(`↻ reprise en phase A (postback ${b.restart})`);
      else { st.restarts++; log(`↻ rechargement inattendu ${b.restart} — reprise en phase A`, "skip"); }
      continue;
    }
    // — Phase C : contrôle AVANT enregistrement (relecture complète)
    s = b.snap;
    const L = (k) => diagNorm(s.f[k].value);
    if (!L("consoLettre") && !(v.lettreEp && opts.forceLetters)) return fail(`lettre EP non calculée par SIGEO (conso ${v.conso}) : le postback de la consommation n'a pas eu lieu`);
    if (!L("gesLettre") && !(v.lettreGes && opts.forceLetters)) return fail(`lettre GES non calculée par SIGEO (émission ${v.ges}) : le postback des émissions n'a pas eu lieu`);
    const plan = diagPlan(s, v, opts);
    const gaps = plan.specs.filter((x) => !x.ok);
    if (!gaps.length) { plan.warnings.forEach(warn); finalPlan = plan; break; }
    const fixable = gaps.some((g) => g.act || (g.phase === "calc" && (st.attempts.dateFin || 0) < 2));
    st.cRounds++;
    if (!fixable || st.cRounds > 2) return fail(`${DIAG_CODES.before} — ${diagGapText(gaps)}`, "ERREUR", DIAG_CODES.before);
    log(`contrôle avant enregistrement : ${gaps.length} écart(s) → nouvelle saisie (essai ${st.cRounds}/2) : ${gaps.map((g) => g.label).join(", ")}`, "skip");
  }
  const showOf = (plan, k) => { const x = plan.specs.find((y) => y.key === k); return x ? `${x.label} ${diagShow(x.kind, x.got)}` : ""; };
  const summary = (plan) => ["conso", "ges", "dateFin", "ademe"].map((k) => showOf(plan, k)).filter(Boolean).join(" · ");
  const letters = () => `${diagNorm(s.f.consoLettre.value) || "?"}/${diagNorm(s.f.gesLettre.value) || "?"}`;
  log(`contrôle avant enregistrement : ${finalPlan.specs.length} valeur(s) conforme(s) · lettres SIGEO ${letters()}`, "ok");

  if (p.simulate) {
    return Object.assign(res, { ok: true, status: "SIMULATION", idDiag: s.idDiag !== "0" ? s.idDiag : "",
      info: `simulation : formulaire rempli et contrôlé (${finalPlan.specs.length} valeurs : ${summary(finalPlan)})` +
        `${opts.ademeDeferred ? `, N° ADEME ${v.ademe} non tenté (DPE avant 01/07/2021 : saisi après le 1er enregistrement)` : ""}, non enregistré` });
  }

  // — Phase D : Enregistrer (jamais « Enreg. et fermer ») — en deux temps si
  //   le N° ADEME n'est saisissable qu'après le 1er enregistrement.
  const ficheUrl = (snap, id) => {
    if (id && new RegExp("/" + id + "(?:[/?#]|$)").test(snap.url)) return snap.url;
    if (id && snap.formAction && new RegExp("/" + id + "(?:[/?#]|$)").test(snap.formAction)) return snap.formAction;
    return id ? diagUrlWithId(snap.url, id) : snap.url;
  };
  const saveOnce = async (n, isNew) => {
    if (!s.saveBtn.present) return E("bouton « Enregistrer » introuvable");
    if (isNew) await ad.journalSet(key, { phase: "saving", row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
    const before = s.docId;
    const r = await ad.exec({ op: "save" });
    if (!r || !r.ok) return E(`Enregistrer n°${n} : ${(r && r.error) || "page indisponible"}`);
    log(`Enregistrer n°${n} → ${isNew ? "création du diagnostic" : "mise à jour du diagnostic"}`);
    const w = await settle(before, { mustReload: true });
    if (w.stopped) return E(`arrêté pendant l'enregistrement n°${n} — vérifie dans SIGEO`, "ARRÊT");
    if (w.timeout) return E(`${DIAG_CODES.timeout} — pas de rechargement terminé ${sec(T)} s après « Enregistrer » n°${n} — vérifie dans SIGEO si le diagnostic a été ${isNew ? "créé" : "mis à jour"}`, "ERREUR", DIAG_CODES.timeout);
    s = w.snap;
    log(`↻ postback (Enregistrer n°${n})`);
    const e = pageError(s);
    if (e) return E(e, s.login ? "AUTH" : "PAGE");
    const m = diagClassifyMessages(s.messages);
    m.infos.forEach((t) => log("SIGEO : " + t));
    if (m.errors.length) {
      if (isNew) await ad.journalSet(key, { phase: "filling", row: p.rowLabel || "", lotId: s.lotId, t: ad.now(), refus: m.errors.join(" | ") });
      return E(`SIGEO a refusé l'enregistrement${n > 1 ? " n°" + n : ""} : ${m.errors.join(" | ")}`, "REFUS");
    }
    // Page rechargée mais toujours en création (id 0, « Veuillez enregistrer… ») :
    // le diagnostic n'a PAS été créé → l'historique ne bloque pas la relance.
    if (isNew && !s.saved) {
      await ad.journalSet(key, { phase: "filling", row: p.rowLabel || "", lotId: s.lotId, t: ad.now(), refus: "non créé" });
      return E(`SIGEO n'a pas créé le diagnostic (page rechargée toujours en création${m.infos.length ? " — " + m.infos.join(" | ") : ""}) — vérifie les champs obligatoires`, "REFUS");
    }
    if (!m.success) return E(`pas de message « Enregistrement réussi ! » après Enregistrer n°${n} — vérifie dans SIGEO`);
    if (!s.attach.present || s.needSave) return E(`enregistré (n°${n}), mais le bloc « Joindre un document » n'apparaît pas`);
    const id = s.idDiag && s.idDiag !== "0" ? s.idDiag : "";
    if (id) res.idDiag = id;
    await ad.journalSet(key, { phase: "saved", idDiag: res.idDiag, url: ficheUrl(s, res.idDiag), row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
    return { ok: true };
  };

  if (!s.saved || changed) {
    const isNew = !s.saved;
    const r1 = await saveOnce(1, isNew);
    if (r1.error) return fail(r1.error, r1.status, r1.code);
    saved1 = { id: res.idDiag, url: ficheUrl(s, res.idDiag), isNew };
    if (isNew && !res.idDiag) warn("id du diagnostic introuvable dans l'URL et dans form.action");
    log(`enregistrement n°1 réussi — fiche DPE ${res.idDiag || "(id introuvable)"} ${isNew ? "créée" : "mise à jour"} : ${saved1.url}`, "ok");
  } else {
    res.idDiag = s.idDiag && s.idDiag !== "0" ? s.idDiag : "";
    if (!J) await ad.journalSet(key, { phase: "saved", idDiag: res.idDiag, url: ficheUrl(s, res.idDiag), row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
    log("diagnostic déjà enregistré et conforme : pas de nouvel enregistrement");
  }

  // — N° ADEME pas encore enregistré (DPE < 01/07/2021 : le champ n'apparaît
  //   qu'après le 1er enregistrement) : saisi, relu, puis Enregistrer n°2.
  //   Enregistré = texte dans span#body_x_labelIdDpe (comparé au contrôle).
  opts.ademeDeferred = false;
  if (v.ademe) {
    const A = s.f.ademe;
    if (A.mode === "input") {
      const sp = { key: "ademe", label: "N° ADEME", kind: "ademe", want: v.ademe, act: { kind: "value", locate: "ademe", id: A.input.id } };
      const w0 = await stable();
      if (w0.stopped) return fail("arrêté par l'utilisateur", "ARRÊT");
      if (w0.timeout) return fail(`${DIAG_CODES.timeout} — page toujours en rechargement avant la saisie du N° ADEME`, "ERREUR", DIAG_CODES.timeout);
      s = w0.snap;
      let r;
      try { r = await ad.exec({ op: "write", action: actOf(sp) }); } catch (e) { r = { ok: false, error: e.message || String(e) }; }
      if (!r || !r.ok) return fail(r && r.code ? `${r.code} — ${r.error}` : `saisie du N° ADEME impossible : ${(r && r.error) || "page indisponible"}`, "ERREUR", (r && r.code) || "");
      log(`N° ADEME : attendu ${diagQ(v.ademe)} · avant ${diagQ(diagNorm(r.before))} · après ${diagQ(diagNorm(r.after))} (après le 1er enregistrement)`);
      if (r.leaving || (r.docId && r.docId !== s.docId)) return fail("la page s'est rechargée pendant la saisie du N° ADEME — relance l'étape (la fiche sera reprise en modification)");
      const s2 = await read();
      const relu = s2 && s2.f.ademe.input ? s2.f.ademe.input.value : "";
      if (!s2 || s2.docId !== s.docId || diagAdemeKey(relu) !== v.ademe) {
        return fail(`${DIAG_CODES.before} — N° ADEME non retenu par la page : relu ${diagQ(diagNorm(relu))} au lieu de ${diagQ(v.ademe)}`, "ERREUR", DIAG_CODES.before);
      }
      s = s2;
      const r2 = await saveOnce(2, false);
      if (r2.error) return fail(r2.error, r2.status, r2.code);
      log(`enregistrement n°2 réussi — N° ADEME ${v.ademe} sur la fiche DPE ${res.idDiag || "?"}`, "ok");
    } else if (A.mode !== "label") {
      return fail(`${DIAG_CODES.ademeAbsent} — champ N° ADEME (#${DIAG_IDS.ademe}) ${A.mode === "hidden" ? "masqué" : "absent"} après l'enregistrement — N° ${v.ademe} non saisi, upload annulé`, "ERREUR", DIAG_CODES.ademeAbsent);
    }
  }

  // — Contrôle APRÈS enregistrement (même comparateur), avant tout upload.
  //   N° ADEME lu dans span#body_x_labelIdDpe s'il existe, sinon dans l'input.
  const pa = diagPlan(s, v, opts);
  pa.warnings.forEach(warn);
  const bad = pa.specs.filter((x) => !x.ok);
  // Fournisseur : code comparé sans zéros de tête (SIGEO stocke un nombre)
  // ET libellé #body:x:proxyFou:x:lbl:x:lbl non vide, sinon écart réel.
  const fn = s.f.fouName || {};
  const fouSpec = pa.specs.find((x) => x.key === "fou");
  if (fouSpec && fouSpec.ok && (fn.labelFound || fn.dedicated) && !fn.name) {
    bad.push({ key: "fouLabel", label: "Fournisseur non reconnu par SIGEO", kind: "text", want: `libellé du code ${v.fou}`, got: "" });
  }
  const ademeSpec = pa.specs.find((x) => x.key === "ademe");
  ademeConfirmed = !v.ademe || !!(ademeSpec && ademeSpec.ok);
  if (bad.length) return fail(`${DIAG_CODES.after} — ${diagGapText(bad)} — upload annulé`, "ERREUR", DIAG_CODES.after);
  const ademeOk = ademeSpec && ademeSpec.ok ? diagAdemeKey(ademeSpec.got) : "";
  if (ademeOk) log(`N° ADEME relu après enregistrement : ${ademeOk} (${s.f.ademe.mode === "label" ? "texte enregistré" : "champ de saisie"})`);
  log(`contrôle après enregistrement : ${pa.specs.length} valeur(s) conforme(s) · ${summary(pa)}`, "ok");
  const fouRelu = fouSpec ? diagNorm(fouSpec.got) : "";
  if (fn.name && fouRelu && fouRelu !== v.fou) log(`Fournisseur ${v.fou} relu ${fouRelu} (zéros de tête supprimés par SIGEO) — OK, libellé : ${fn.name}`);
  else if (fn.name) log(`fournisseur : ${v.fou} — ${fn.name}`);
  else warn("nom du fournisseur non repéré sur la page (active le mode debug pour le localiser)");

  const okInfo = () => [res.idDiag ? "diagnostic " + res.idDiag : "diagnostic enregistré", ademeOk ? "ADEME " + ademeOk : ""].filter(Boolean).join(", ");
  if (!p.attach) return Object.assign(res, { ok: true, status: "OK", info: okInfo() + " — PDF non joint (option désactivée)" });

  // — Phase E : pièce jointe (idempotente : jamais deux documents « DPE »)
  // Vérification relue dans le DOM : plus de documents qu'avant l'upload ET
  // une ligne « DPE ». SIGEO n'affiche pas toujours le document tout de
  // suite : on reclique alors « Enregistrer » (même fiche, même popup) et on
  // revérifie — 3 tentatives, 10 s d'attente active chacune, JAMAIS de
  // nouvel upload (il créerait un doublon).
  const RESAVE_MAX = 3, RESAVE_MS = 10000;
  const ord = (k) => (k === 1 ? "1er" : k + "e");
  const uploadSeen = (snap, count0) => {
    const row = snap && snap.grid ? diagFindDocRow(snap.grid, desc) : null;
    return { ok: !!row && snap.grid.count > count0, row, count: snap && snap.grid ? snap.grid.count : 0 };
  };
  const recheckBySaving = async (count0) => {
    let last = "";
    for (let k = 1; k <= RESAVE_MAX; k++) {
      const w0 = await stable();
      if (w0.stopped) return { stopped: true };
      if (w0.timeout) return { error: `${DIAG_CODES.timeout} — page toujours en rechargement avant le ${ord(k)} nouvel enregistrement`, code: DIAG_CODES.timeout };
      s = w0.snap;
      const pe0 = pageError(s);
      if (pe0) return { error: pe0, status: s.login ? "AUTH" : "PAGE" };
      if (res.idDiag && s.idDiag !== res.idDiag) return { error: `la popup n'affiche plus la fiche DPE ${res.idDiag} (${s.url}) — rien n'a été recliqué`, status: "PAGE" };
      if (!s.saveBtn.present) return { error: "bouton « Enregistrer » introuvable pour revérifier l'upload" };
      const before = s.docId, t0 = ad.now();
      const r = await ad.exec({ op: "save" });
      if (!r || !r.ok) return { error: `Enregistrer (revérification ${k}/${RESAVE_MAX}) : ${(r && r.error) || "page indisponible"}` };
      log(`document « ${desc} » absent de la liste : Enregistrer recliqué (${k}/${RESAVE_MAX}), sans nouvel upload`, "skip");
      // Rechargement terminé (nouveau docId, readyState complete)…
      const w = await settle(before, { mustReload: true, timeoutMs: RESAVE_MS });
      if (w.stopped) return { stopped: true };
      if (w.timeout) { last = `pas de rechargement terminé ${sec(RESAVE_MS)} s après Enregistrer`; log(`${last} (${k}/${RESAVE_MAX})`, "skip"); continue; }
      log(`↻ postback (Enregistrer — revérification ${k}/${RESAVE_MAX})`);
      // … puis bouton de nouveau cliquable et liste rendue ; la liste est
      // relue dans le DOM jusqu'au document ou jusqu'aux 10 s.
      let snap = w.snap, seen = null;
      for (;;) {
        if (snap && snap.readyState === "complete" && !snap.leaving && snap.saveBtn.present && snap.saveBtn.visible &&
            !snap.saveBtn.disabled && snap.grid.present) {
          seen = uploadSeen(snap, count0);
          if (seen.ok) break;
        }
        if (ad.stopped() || ad.now() - t0 > RESAVE_MS) break;
        await ad.sleep(250);
        const n = await read();
        if (n) snap = n;
      }
      s = snap;
      const errs2 = msgErrors(s);
      if (errs2.length) log(`SIGEO : ${errs2.join(" | ")}`, "skip");
      if (seen && seen.ok) return { ok: true, attempt: k, row: seen.row };
      last = `${count0} → ${s.grid.count} document(s), ligne « ${desc} » ${diagFindDocRow(s.grid, desc) ? "présente" : "absente"}${errs2.length ? " — SIGEO : " + errs2.join(" | ") : ""}`;
    }
    return { failed: true, last };
  };

  const existing = diagFindDocRow(s.grid, desc);
  if (existing) {
    if (fileRef && !diagRowHasFile(existing, fileRef.file.name)) warn(`un document « ${desc} » est déjà joint (${existing.join(" · ")}) mais pas sous le nom ${fileRef.file.name}`);
    if (J && J.phase === "uploaded") await ad.journalSet(key, { phase: "done", idDiag: res.idDiag, file: J.file || "", row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
    return Object.assign(res, { ok: true, status: "DÉJÀ JOINT", info: `${okInfo()} — document « ${desc} » déjà joint, pas de nouvel upload` });
  }
  // PDF déjà envoyé sur cette fiche par un passage précédent, jamais apparu
  // dans la liste : Enregistrer + revérification, pas de second upload.
  const Jup = await ad.journalGet(key);
  if (Jup && Jup.phase === "uploaded" && (!Jup.idDiag || !res.idDiag || Jup.idDiag === res.idDiag)) {
    log(`un PDF (${Jup.file || "?"}) a déjà été envoyé sur cette fiche lors d'un passage précédent sans apparaître dans la liste : Enregistrer + revérification, sans nouvel upload`, "skip");
    pdfSent = true;
    const c0 = Number.isFinite(Jup.count0) ? Jup.count0 : s.grid.count;
    const rv = await recheckBySaving(c0);
    if (rv.stopped) return fail("arrêté pendant la revérification de l'upload — vérifie la liste des documents", "ARRÊT");
    if (rv.error) return fail(`${rv.error} — PDF déjà envoyé, non ré-uploadé`, rv.status || "ERREUR", rv.code || "");
    if (!rv.ok) {
      return fail(`upload non vérifié après ${RESAVE_MAX} tentatives d'enregistrement (PDF envoyé lors d'un passage précédent) : ${rv.last} — pas de nouvel upload pour éviter un doublon : vérifie dans SIGEO ; s'il n'y est vraiment pas, « Vider l'historique DPE » puis relance`);
    }
    warn(`document « ${desc} » apparu après le ${ord(rv.attempt)} nouvel enregistrement (${rv.attempt}/${RESAVE_MAX})`);
    await ad.journalSet(key, { phase: "done", idDiag: res.idDiag, file: Jup.file || "", row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
    pdfDone = true;
    log(`PDF joint : ${s.grid.count} document(s) — ${rv.row.join(" · ")}`, "ok");
    return Object.assign(res, { ok: true, status: "OK", file: Jup.file || "", info: `${okInfo()} — PDF joint (${Jup.file || "envoyé précédemment"}), visible après ${rv.attempt} nouvel(s) enregistrement(s)` });
  }
  if (!fileRef) {
    const e = await resolveFile();
    if (e) return fail(`fichier DPE ${e} — diagnostic enregistré, PDF non joint`, "FICHIER");
  }
  const name = fileRef.file.name;
  const ref = v.ademeNc ? "" : (ademeOk || v.ademe);
  const chk = ref === fileAdeme && fileChk ? fileChk : await checkFile(ref);
  if (chk.status === "other") return fail(`le ${chk.where} de ${name} contient un autre N° ADEME (${chk.others.join(", ")}) que la fiche (${ref}) — upload annulé`, "FICHIER");
  const chkTxt = {
    name: "N° ADEME présent dans le nom du fichier",
    text: "N° ADEME trouvé dans le texte du PDF",
    unknown: "N° ADEME introuvable dans le nom et dans le texte du PDF (scan ?)",
    none: "pas de N° ADEME (DPE sans champ ADEME) : fichier retenu par son nom"
  }[chk.status];
  if (chk.status === "unknown" && p.batch && !p.confirmUpload) return fail(`${chkTxt} — upload refusé en batch sans confirmation`, "FICHIER");
  // Confirmation avant upload selon l'option (ligne active comme boucle) ;
  // toujours demandée si le N° ADEME du PDF n'a pas pu être vérifié.
  if (p.confirmUpload || chk.status === "unknown") {
    const ok = await ad.confirm({
      row: p.rowLabel || "", lot: v.lot || s.lotId, fiche: res.idDiag || s.idDiag || "", ademe: ref || (v.ademeNc ? "NC" : ""),
      ademeSrc: ademeOk ? "confirmé par SIGEO" : (ref ? "données de la ligne" : ""),
      file: name, path: fileRef.file.path || name, size: diagFmtSize(bytes.length), via: fileRef.via,
      check: chkTxt, risky: chk.status === "unknown"
    });
    if (!ok) return fail("upload annulé (non confirmé) — diagnostic enregistré, PDF non joint", "ANNULÉ");
    // La page a pu être rechargée pendant la confirmation : on relit tout.
    const w = await stable();
    if (w.stopped) return fail("arrêté par l'utilisateur", "ARRÊT");
    s = w.snap || await read();
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
  // Postback complet attendu ; une mise à jour partielle de la grille
  // (update panel body_x_upcDiag) est acceptée aussi.
  const w = await settle(before, { mustReload: true, timeoutMs: TU, doneWhen: (x) => !!x.grid && x.grid.count > count0 });
  if (w.stopped) return fail("arrêté pendant l'upload — vérifie la liste des documents", "ARRÊT");
  if (w.timeout) return fail(`${DIAG_CODES.timeout} — pas de rechargement terminé ${sec(TU)} s après l'envoi du PDF — vérifie la liste des documents`, "ERREUR", DIAG_CODES.timeout);
  s = w.snap;
  pdfSent = true;
  log(w.partial ? "↻ postback partiel (upload du PDF, grille mise à jour)" : "↻ postback (upload du PDF)");
  const e3 = pageError(s);
  if (e3) return fail(e3, s.login ? "AUTH" : "PAGE");
  const errs = msgErrors(s);
  let seen = uploadSeen(s, count0);
  let row = seen.row;
  if (!seen.ok) {
    // Pas (encore) dans la liste : PDF envoyé, à revérifier en recliquant
    // « Enregistrer » — noté au journal pour qu'une relance ne ré-uploade pas.
    log(`upload non vérifié tout de suite : ${count0} → ${s.grid.count} document(s), ligne « ${desc} » ${row ? "présente" : "absente"}` +
      `${errs.length ? " — SIGEO : " + errs.join(" | ") : ""} → nouvel enregistrement de la fiche`, "skip");
    await ad.journalSet(key, { phase: "uploaded", idDiag: res.idDiag, count0, file: name, row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
    const rv = await recheckBySaving(count0);
    if (rv.stopped) return fail("arrêté pendant la revérification de l'upload — vérifie la liste des documents", "ARRÊT");
    if (rv.error) return fail(`${rv.error} — PDF envoyé, non vérifié`, rv.status || "ERREUR", rv.code || "");
    if (!rv.ok) {
      return fail(`upload non vérifié après ${RESAVE_MAX} tentatives d'enregistrement : ${rv.last}${errs.length ? " — SIGEO (upload) : " + errs.join(" | ") : ""} — PDF envoyé une seule fois, non ré-uploadé : vérifie la liste des documents dans SIGEO`);
    }
    row = rv.row;
    warn(`document « ${desc} » apparu après le ${ord(rv.attempt)} nouvel enregistrement (${rv.attempt}/${RESAVE_MAX})`);
  } else if (errs.length) warn(`SIGEO après l'upload : ${errs.join(" | ")}`);
  if (s.grid.count > count0 + 1) warn(`${s.grid.count - count0} documents ajoutés au lieu d'un (${count0} → ${s.grid.count}) — vérifie la liste`);
  if (!diagRowHasFile(row, name)) warn(`nom ${name} non retrouvé dans la ligne « ${desc} » (${row.join(" · ")})`);
  await ad.journalSet(key, { phase: "done", idDiag: res.idDiag, file: name, row: p.rowLabel || "", lotId: s.lotId, t: ad.now() });
  pdfDone = true;
  log(`PDF joint : ${s.grid.count} document(s) — ${row.join(" · ")}`, "ok");
  return Object.assign(res, { ok: true, status: "OK", file: name, info: `${okInfo()} — PDF joint (${name})` });
}

/* ---------------------------------------------------------------- *
 * 8. Corps HTML de l'étape (panneau)                               *
 * ---------------------------------------------------------------- */
// (v2.16.0 : 20000, ancien défaut du délai de rechargement, devient 15000.)
function diagStepHtml(step, esc) {
  const cols = Object.assign({}, DIAG_PRESETS.default, step.diagCols || {});
  const grid = DIAG_FIELDS.map((f) => `
        <label title="${esc(f.hint || f.k)}">${f.label}${f.req ? " *" : ""}</label>
        <input type="text" class="scn-diag-col" data-k="${f.k}" value="${esc(cols[f.k] ?? "")}" />`).join("");
  return `
      <p class="hint scn-only-diag">Remplit le formulaire « Diagnostic » SIGEO (popup <code>diag_manage</code>) en phases : A) champs à postback un par un (type → fournisseur → établi le → conso → GES, rechargement attendu à chaque fois), B) champs sans postback (N° ADEME, dates, coûts, lettres forcées…), C) contrôle complet avant « Enregistrer », D) enregistrement et contrôle après, E) PDF joint. « Valable jusqu'au » n'est jamais saisi (calculé par SIGEO, vérifié). Tout rechargement inattendu fait repartir de la phase A. Jamais de touche Entrée ni « Enreg. et fermer » ; la popup reste ouverte.</p>
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
          <button class="btn sm scn-diag-preset" data-preset="intralia" type="button" title="Fichier final de l'extraction « DPE extraction » (colonnes + URL de la fiche)">Fichier DPE extraction</button>
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
        <label class="checkbox-row" title="Coché : chaque upload attend un clic sur « Joindre ». Décoché : le PDF est joint automatiquement à la fin (ligne active et boucle) ; une confirmation reste demandée si le N° ADEME du PDF n'a pas pu être vérifié (scan)."><input type="checkbox" class="scn-diag-confirm" ${step.diagConfirm === false ? "" : "checked"} /> confirmer avant de joindre le PDF</label>
        <label class="checkbox-row" title="Remplit et contrôle tout, sans cliquer Enregistrer ni joindre le PDF"><input type="checkbox" class="scn-diag-simulate" ${step.diagSimulate ? "checked" : ""} /> simulation</label>
      </div>
      <div class="scn-row scn-only-diag">
        <label class="checkbox-row" title="Si la lettre EP/GES du DPE (colonnes « Lettre … du DPE ») diffère de celle calculée par SIGEO, elle est saisie dans « Forcer la lettre » et contrôlée ; sinon simple avertissement"><input type="checkbox" class="scn-diag-forceletters" ${step.diagForceLetters === false ? "" : "checked"} /> forcer la lettre du DPE si SIGEO en calcule une autre</label>
      </div>
      <div class="scn-row scn-only-diag">
        <label>résultat → colonne :</label>
        <input type="text" class="new-col-input scn-diag-resultcol" value="${esc(step.diagResultCol ?? DIAG_DEFAULTS.resultCol)}" />
        <label class="checkbox-row" title="Une ligne dont la colonne résultat commence par « OK » est sautée"><input type="checkbox" class="scn-diag-skipdone" ${step.diagSkipDone === false ? "" : "checked"} /> sauter les lignes « OK »</label>
      </div>
      <div class="scn-row scn-only-diag">
        <label title="Délai maximal d'un rechargement (postback) ; au-delà : POSTBACK_TIMEOUT">rechargement max (ms) :</label>
        <input type="number" class="scn-diag-timeout" min="3000" step="1000" value="${esc(step.diagTimeout && step.diagTimeout !== 20000 ? step.diagTimeout : DIAG_DEFAULTS.timeoutMs)}" />
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
    diagForceLetters: chk(".scn-diag-forceletters", true),
    diagResultCol: val(".scn-diag-resultcol", DIAG_DEFAULTS.resultCol).trim(),
    diagSkipDone: chk(".scn-diag-skipdone", true),
    diagTimeout: parseInt(val(".scn-diag-timeout"), 10) || DIAG_DEFAULTS.timeoutMs,
    diagUploadTimeout: parseInt(val(".scn-diag-uptimeout"), 10) || DIAG_DEFAULTS.uploadTimeoutMs,
    diagDebug: chk(".scn-diag-debug", false)
  };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    DIAG_IDS, DIAG_FIELDS, DIAG_PRESETS, DIAG_DEFAULTS, DIAG_INTRALIA_URL,
    diagNorm, diagFold, diagDate, diagDateEq, diagSerialToDate, diagExpectedValidity, diagIsNewDpe, ademeSaisissableAvantEnreg,
    diagParseNum, diagNum, diagNumEq, diagBool, diagLetter, diagAdeme, diagAdemeKey, diagType, diagNormalizeData,
    diagSame, diagShow, diagQ, diagFouKey, DIAG_CODES, DIAG_PHASE_A,
    diagMatchFile, diagCheckFileAdeme, diagIsPdf, diagToBase64, diagFmtSize,
    diagUrlParts, diagUrlWithId, diagJournalKey, diagClassifyMessages, diagFindDocRow, diagRowHasFile,
    diagPageOp, diagMainHandler, diagPlan, diagCompare, diagGapText, diagCapped, diagSettle, diagProcess, diagStepHtml, diagStepRead
  };
}
