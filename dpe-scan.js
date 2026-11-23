/* ====================================================================
 * NoHands OSA — Analyse des DPE téléchargés (dossier → Excel)
 * --------------------------------------------------------------------
 * Onglet Extraction, type « Analyse DPE (dossier → Excel) ».
 *
 * Lit les PDF de DPE d'un dossier (ex. Téléchargements/DPE_SIGEO, rempli
 * par le type « DPE SIGEO »), en extrait les informations utiles à la
 * saisie SIGEO (établi le, valable jusqu'au, énergie primaire, GES, coûts,
 * n° ADEME, surface, diagnostiqueur…) et retrouve le n° fournisseur (FU)
 * grâce à la liste de correspondance « NOM : n° FU ».
 *
 * Port JavaScript de l'outil DPE Scan (Rust) : mêmes expressions, même
 * repérage de la page DPE et de la page des consommations (p.3), mêmes
 * recoupements (étiquette = plancher(total ÷ surface)) et mêmes barèmes de
 * lettres (2021 double seuil, petites surfaces < 40 m² depuis le
 * 01/07/2024, ancien DPE).
 *
 * Ce fichier ne contient que de la logique pure (aucun accès au DOM ni à
 * chrome.*) : il est chargé par sidepanel.html AVANT sidepanel.js et
 * importable sous Node (tests). Le rendu des pages et l'OCR sont fournis
 * par un adaptateur (`io`) : Tesseract.js dans le panneau, tesseract en
 * ligne de commande dans les tests.
 * ==================================================================== */

const DS_TYPE_DIAG = "Performance Energetique";

// Colonnes écrites dans le tableau (créées si absentes), dans cet ordre.
// Uniquement ce qui sert à la saisie du DPE dans SIGEO (formulaire
// diag_manage) + statut, PDF à joindre et remarques de contrôle (v2.17.2).
// Les totaux (kWh/an, kg CO2/an), la surface, l'adresse, le diagnostiqueur…
// restent lus pour les recoupements mais ne sont plus écrits. Coûts et date
// d'indexation : seulement pour un DPE établi depuis le 01/07/2021 (avant,
// SIGEO n'a pas de champ d'estimation).
const DS_COLS = [
  "DPE analyse", "PDF analysé", "N° FU", "Établi le", "Valable jusqu'au", "N° ADEME",
  "Énergie primaire (kWh/m²/an)", "Classe énergie", "GES (kg CO2/m²/an)", "Classe GES",
  "Coût min (€/an)", "Coût max (€/an)", "Prix indexés au", "Remarques DPE"
];

// Liste fournisseurs fournie par Maëlle (modifiable dans le panneau).
const DS_DEFAULT_SUPPLIERS = [
  "EXPERT IMMOBILIER : 0528825",
  "DIAG A FAIRE : 0410962",
  "ACADE : 0448957",
  "ACDI : 0044021",
  "ALPES CONTROLE : 0178989",
  "ROUSSEL EXPERTISE : 0176995",
  "AV ENVIRONNEMENT : 0425954",
  "DIAGMTER : 0373934",
  "ROUSSEL Frédérique : 0418483",
  "ACT'IM : 0429872",
  "SOCOTECH : 0048730",
  "PROTECH : 0426828",
  "2DI : 0520226",
  "BOMTEMS : 0594816",
  "AGENDA EXPERTISE : 0515321",
  "ARTHEMIS : 0596024",
  "OPERA GROUP : 0603258",
  "DIAG3 : 0468380",
  "SPEED DIAG : 0570646",
  "DIAG 75 : 0394076x",
  "ANAPLOMB : 0552043",
  "CAP EXPERTISE : 0374697",
  "POLYEXPERT : 0417277",
  "EXIM : 0480110",
  "BATISCES : 0375853",
  "ACDE : 0376064",
  "ARTHEMIS : 0596024",
  "APAVE INFRASTRUCTURES ET CONSTRUCTION FRANCE : 0559809"
].join("\n");

/* ---------------------------------------------------------------- *
 * 1. Normalisation                                                 *
 * ---------------------------------------------------------------- */

// Caractères « exotiques » → équivalent simple (casse conservée).
function dsClean(s) {
  return String(s ?? "")
    .replace(/[      \t]/g, " ")
    .replace(/\r/g, "")
    .replace(/[’‘`´ʼ]/g, "'")
    .replace(/[“”«»]/g, "\"")
    .replace(/[₂²]/g, "2").replace(/³/g, "3")
    .replace(/[–—‑]/g, "-")
    .replace(/[￾﻿\u0002\u0000]/g, "");
}

// Texte de recherche : nettoyé, sans accents, minuscules (retours à la ligne conservés).
function dsNorm(s) {
  return dsClean(s).replace(/œ/g, "oe").replace(/Œ/g, "OE").replace(/æ/g, "ae").replace(/Æ/g, "AE")
    .normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
}

// Clé compacte : MAJUSCULES sans accents, lettres/chiffres, mots séparés par un espace.
function dsKey(s) {
  return dsNorm(s).toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

// Nombre écrit à la française (« 14 954 », « 318,99 », « 1.029 ») → nombre.
function dsParseNum(s) {
  let t = dsClean(s).replace(/\s+/g, "").replace(/^[^0-9]+|[^0-9]+$/g, "");
  if (!t) return null;
  if (t.includes(",")) t = t.replace(/\./g, "").replace(",", ".");
  else {
    const pos = t.lastIndexOf(".");
    if (pos >= 0 && t.length - pos - 1 === 3 && pos <= 3) t = t.replace(/\./g, "");
  }
  const v = Number(t);
  return Number.isFinite(v) ? v : null;
}

// Confusions OCR classiques dans une chaîne censée ne contenir que des chiffres.
function dsOcrDigits(s) {
  const map = { O: "0", o: "0", D: "0", Q: "0", I: "1", l: "1", "|": "1", i: "1", "!": "1", "]": "1", "[": "1",
    Z: "2", z: "2", S: "5", s: "5", B: "8", G: "6", T: "7" };
  let out = "";
  for (const c of String(s)) {
    if (c >= "0" && c <= "9") out += c;
    else if (c === " ") continue;
    else out += map[c] || "?";
  }
  return out;
}

// Nombre pour le tableau : entier tel quel, sinon 2 décimales max, virgule.
function dsFmtNum(v) {
  if (v === null || v === undefined || !Number.isFinite(v)) return "";
  if (Math.abs(v - Math.round(v)) < 1e-9) return String(Math.round(v));
  return (Math.round(v * 100) / 100).toFixed(2).replace(/0+$/, "").replace(/\.$/, "").replace(".", ",");
}

/* ---------------------------------------------------------------- *
 * 2. Dates                                                         *
 * ---------------------------------------------------------------- */

const DS_MONTHS = { janvier: 1, janv: 1, jan: 1, fevrier: 2, fevr: 2, fev: 2, mars: 3, avril: 4, avr: 4, mai: 5, juin: 6,
  juillet: 7, juil: 7, aout: 8, septembre: 9, sept: 9, sep: 9, octobre: 10, oct: 10, novembre: 11, nov: 11, decembre: 12, dec: 12 };

function dsDaysInMonth(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }

function dsDate(d, m, y) {
  d = +d; m = +m; y = +y;
  if (y < 100) y += 2000;
  if (!(y >= 1990 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= dsDaysInMonth(y, m))) return null;
  return { y, m, d };
}
function dsDateCmp(a, b) { return a.y - b.y || a.m - b.m || a.d - b.d; }
function dsDateFr(x) { return x ? `${String(x.d).padStart(2, "0")}/${String(x.m).padStart(2, "0")}/${x.y}` : ""; }
function dsPlusYearsMinusOneDay(x, years) {
  const y = x.y + years;
  let d = Math.min(x.d, dsDaysInMonth(y, x.m)), m = x.m, yy = y;
  if (d > 1) d -= 1;
  else if (m > 1) { m -= 1; d = dsDaysInMonth(yy, m); }
  else { yy -= 1; m = 12; d = 31; }
  return { y: yy, m, d };
}

/* ---------------------------------------------------------------- *
 * 3. Expressions (sur texte normalisé : minuscules, sans accents)  *
 * ---------------------------------------------------------------- */

const DS_DATE = "(\\d{1,2})\\s?[/.\\-]\\s?(\\d{1,2})\\s?[/.\\-]\\s?(\\d{4}|\\d{2}\\b)";
const DS_DATE_LONG = "(\\d{1,2})(?:er)?\\s+(janvier|fevrier|mars|avril|mai|juin|juillet|aout|septembre|octobre|novembre|decembre)\\s+(\\d{4})";
const DS_NUM = "(\\d{1,3}(?:[ .]\\d{3})*(?:,\\d+)?|\\d+(?:,\\d+)?)";
const dsRe = (s, f = "") => new RegExp(s, f);

const DS_RE = {
  etabli: dsRe(`etabli[e]?\\s*l[e€]\\s*[:;.,]?\\s*${DS_DATE}`, "g"),
  etabliLong: dsRe(`etabli[e]?\\s*l[e€]\\s*[:;.,]?\\s*${DS_DATE_LONG}`, "g"),
  valable: dsRe(`valable\\s*jusqu.{0,3}au\\s*[:;.,]?\\s*${DS_DATE}`, "g"),
  valableLong: dsRe(`valable\\s*jusqu.{0,3}au\\s*[:;.,]?\\s*${DS_DATE_LONG}`, "g"),
  dateEtabOld: dsRe(`(?:date\\s*d.?\\s*etablissement(?:\\s*du\\s*(?:rapport|dpe|diagnostic))?|date\\s*du\\s*rapport|date\\s*de\\s*realisation(?:\\s*du\\s*dpe)?)\\s*[:;.]?\\s*${DS_DATE}`, "g"),
  dateVisite: dsRe(`date\\s*de\\s*(?:la\\s*)?visite(?:\\s*du\\s*bien)?\\s*[:;.]?\\s*${DS_DATE}`, "g"),
  ademeStrict: /(?:^|[^0-9a-z])(\d{2}(?:\d{2}|2[ab])[a-z]\d{7}[a-z])(?=[^0-9a-z]|$)/g,
  ademeLabel: /(?:n\s*[°o*º]\s*(?:ademe|dpe)?|ademe|numero\s*(?:du\s*)?dpe|reference\s*ademe|n°\s*d.enregistrement)\s*[:;.]?\s*([0-9a-z][0-9a-z ]{11,17})/g,
  cout: /entre\s*(\d[\d ]{0,7}?)\s*(?:[€¢]|eur|e\b)\s*,?\s*(?:et|a)\s*(\d[\d ]{0,7}?)\s*(?:[€¢]|eur|e\b)\s*(?:\/|par)?\s*an/,
  coutLoose: /entre\s*(\d[\d ]{0,7}?)\s*(?:[€¢ce]|eur)?\s*,?\s*(?:et|a)\s*(\d[\d ]{0,7}?)\s*(?:[€¢ce]|eur)?\s*(?:\/|par)\s*an/,
  // Ordre du flux PDF parfois mélangé : « par an entre et 1430 € 1980 € »
  coutSwap: /entre\s*et\s*(\d[\d ]{0,7}?)\s*(?:€|eur)\s*(\d[\d ]{0,7}?)\s*(?:€|eur)/,
  // Montants sur une autre ligne que « entre … et … par an » (polices de tailles différentes)
  // « Prix moyens des énergies indexés au 1er janvier 2021 » (date de référence des prix)
  prixRefLong: dsRe(`index[e]?s?\\s*(?:au|a\\s*la\\s*date\\s*du)\\s*[:;.]?\\s*${DS_DATE_LONG}`, "g"),
  prixRef: dsRe(`index[e]?s?\\s*(?:au|a\\s*la\\s*date\\s*du)\\s*[:;.]?\\s*${DS_DATE}`, "g"),
  coutNear: /entre[^0-9€\n]{0,40}?\n?[^0-9€]{0,40}?(\d[\d ]{0,7}?)\s*(?:€|eur)[^0-9€]{0,25}?(\d[\d ]{0,7}?)\s*(?:€|eur)/,
  twoAmounts: /(?:^|[^0-9])(\d{1,3}(?: \d{3})?|\d{4,5})\s*[€¢][^\d€]{0,25}?(\d{1,3}(?: \d{3})?|\d{4,5})\s*[€¢]/,
  coutTotalRow: /energie\s*totale\s*pour\s*les[\s\S]{0,90}?entre\s*(\d[\d ]{0,7}?)\s*(?:[€¢]|e\b)\s*et\s*(\d[\d ]{0,7}?)\s*(?:[€¢]|e\b)/,
  surface: /surface\s*(?:habitable|de\s*reference|utile|de\s*ref\.?)\s*(?:\(m2\))?\s*[:;.]?\s*(\d{1,4}(?:[,.]\d{1,2})?)\s*m/,
  gesTotal: /ce\s*logement\s*emet\s*(\d[\d ]{0,7}?)\s*kg/,
  total2021: /usages\s*recenses\s*[:;.]?[^0-9(]{0,60}?(\d[\d ]{0,8}?)\s*kwh(?:[^0-9]{0,12}?(\d[\d ]{0,8}?)\s*kwh\s*[ée]?\s*\.?\s*f)?/,
  total2021B: /energie\s*totale\s*pour\s*les\s*(?:usages\s*recenses\s*[:;.]?\s*)?(\d[\d ]{0,8}?)\s*kwh/,
  efTotal: /\(\s*(\d[\d ]{0,8}?)\s*kwh\s*e\s*\.?\s*f/,
  totalOld: /consommations?\s*d.?\s*energie\s*pour\s*les[\s\S]{0,70}?(\d[\d ]*,\d+)\s+(\d[\d ]*,\d+)\s+(\d[\d ]*,\d+)\s*(?:€|e\b|eur)/,
  consoConv: dsRe(`consommation\\s*conventionnelle\\s*[:;.]?\\s*${DS_NUM}\\s*k?k?wh`),
  emissionOld: dsRe(`estimation\\s*des\\s*emissions\\s*[:;.]?\\s*${DS_NUM}\\s*kg`),
  labelPair: /(?:^|[^0-9,./])(\d{1,4})\s*[|!/\]\[:;i]?\s*[|]?\s*(\d{1,3})\s*[*°'"x»+#^]\s*(?:\|?\s*([a-g])\b)?/g,
  loneNum: /^\W{0,3}(\d{1,4})\W{0,3}$/,
  gesScale: /(?:^|[^0-9])(\d(?: ?\d){0,2})\s*kg\s*\.?\s*(?:eq)?\s*co\s*[2z,.]?\s*\/\s*m/g,
  efM2: /(\d{1,4})\s*kwh\s*\/\s*m2?\s*\/?\s*an\s*d.?\s*energie\s*finale/
};

function dsDateFromMatch(m) {
  const d = m[1], ms = m[2], y = m[3];
  const mo = /^\d+$/.test(ms) ? +ms : DS_MONTHS[ms.replace(/\.$/, "")];
  return mo ? dsDate(d, mo, y) : null;
}

function dsFirstDate(texts, res) {
  for (const t of texts) {
    if (!t) continue;
    for (const re of res) {
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(t))) {
        const d = dsDateFromMatch(m);
        if (d) return d;
        if (m[0] === "") re.lastIndex++;
      }
    }
  }
  return null;
}

/* ---------------------------------------------------------------- *
 * 4. N° ADEME                                                      *
 * ---------------------------------------------------------------- */

function dsAdemePlausible(s) {
  if (!s || s.length !== 13) return false;
  const yy = parseInt(s.slice(0, 2), 10);
  return yy >= 6 && yy <= 40;
}

function dsFindAdemeStrict(t) {
  const re = DS_RE.ademeStrict;
  re.lastIndex = 0;
  let m;
  while ((m = re.exec(t || ""))) {
    const s = m[1].toUpperCase();
    if (dsAdemePlausible(s)) return s;
  }
  return null;
}

function dsToDigit(c) { const d = dsOcrDigits(c); return /^\d$/.test(d) ? d : null; }
function dsToLetter(c) {
  if (/^[A-Z]$/.test(c)) return c;
  return { 0: "O", 1: "I", 2: "Z", 5: "S", 6: "G", 8: "B" }[c] || null;
}

// N° ADEME depuis une lecture OCR approximative (« 2121É0449628F », « 2121 E O449628 F »).
function dsRepairAdeme(raw) {
  const chars = [...dsNorm(raw).replace(/\s+/g, "").toUpperCase()];
  for (let start = 0; start + 13 <= chars.length; start++) {
    let out = "", ok = true;
    for (let i = 0; i < 13 && ok; i++) {
      const c = chars[start + i];
      let f = null;
      if (i <= 2 || (i >= 5 && i <= 11)) f = dsToDigit(c);
      else if (i === 3) f = (out.endsWith("2") && (c === "A" || c === "B")) ? c : dsToDigit(c);
      else f = dsToLetter(c);
      if (f === null) ok = false; else out += f;
    }
    if (ok && dsAdemePlausible(out)) return out;
  }
  return null;
}

/* ---------------------------------------------------------------- *
 * 5. Barèmes des lettres                                           *
 * ---------------------------------------------------------------- */

const DS_LETTERS = "ABCDEFG";
const DS_EP_2021 = [70, 110, 180, 250, 330, 420];
const DS_GES_2021 = [6, 11, 30, 50, 70, 100];
const DS_EP_OLD = [50, 90, 150, 230, 330, 450];
const DS_GES_OLD = [5, 10, 20, 35, 55, 80];
// Petites surfaces (arrêté du 25 mars 2024), de 8 à 40 m² : [surface, seuils EP A..F, seuils GES A..F]
const DS_SMALL = [
  [8, [146, 186, 386, 505, 622, 739], [11, 16, 44, 68, 90, 122]],
  [9, [134, 174, 355, 464, 574, 685], [11, 16, 42, 65, 87, 118]],
  [10, [124, 164, 329, 428, 533, 640], [10, 15, 40, 62, 84, 115]],
  [11, [117, 157, 311, 403, 503, 607], [10, 15, 39, 61, 82, 113]],
  [12, [112, 152, 296, 381, 478, 578], [9, 14, 38, 59, 81, 111]],
  [13, [108, 148, 283, 363, 456, 554], [9, 14, 37, 58, 79, 110]],
  [14, [104, 144, 273, 348, 437, 532], [9, 14, 37, 58, 78, 108]],
  [15, [100, 140, 263, 333, 421, 514], [8, 13, 36, 56, 76, 107]],
  [16, [97, 137, 255, 325, 412, 504], [8, 13, 35, 55, 76, 106]],
  [17, [95, 135, 248, 318, 404, 496], [8, 13, 35, 55, 75, 105]],
  [18, [92, 132, 241, 311, 397, 489], [8, 13, 34, 54, 75, 105]],
  [19, [90, 130, 235, 305, 390, 482], [8, 13, 34, 54, 74, 105]],
  [20, [88, 128, 230, 300, 385, 476], [8, 13, 34, 54, 74, 104]],
  [21, [87, 127, 225, 295, 380, 471], [8, 13, 33, 53, 74, 104]],
  [22, [85, 125, 221, 291, 375, 466], [7, 12, 33, 53, 73, 103]],
  [23, [84, 124, 217, 287, 371, 462], [7, 12, 33, 53, 73, 103]],
  [24, [82, 122, 214, 284, 367, 458], [7, 12, 33, 53, 73, 103]],
  [25, [81, 121, 210, 280, 363, 454], [7, 12, 32, 52, 73, 103]],
  [26, [80, 120, 207, 277, 360, 451], [7, 12, 32, 52, 72, 103]],
  [27, [79, 119, 204, 274, 357, 447], [7, 12, 32, 52, 72, 102]],
  [28, [78, 118, 202, 272, 354, 444], [7, 12, 32, 52, 72, 102]],
  [29, [77, 117, 199, 269, 351, 442], [7, 12, 32, 52, 72, 102]],
  [30, [76, 116, 197, 267, 349, 439], [7, 12, 32, 52, 72, 102]],
  [31, [76, 116, 195, 265, 346, 437], [7, 12, 31, 51, 72, 102]],
  [32, [75, 115, 193, 263, 344, 434], [7, 12, 31, 51, 71, 101]],
  [33, [74, 114, 191, 261, 342, 432], [7, 12, 31, 51, 71, 101]],
  [34, [74, 114, 189, 259, 340, 430], [7, 12, 31, 51, 71, 101]],
  [35, [73, 113, 188, 258, 338, 428], [7, 12, 31, 51, 71, 101]],
  [36, [72, 112, 186, 256, 337, 427], [7, 12, 31, 51, 71, 101]],
  [37, [72, 112, 185, 255, 335, 425], [7, 12, 31, 51, 71, 101]],
  [38, [71, 111, 183, 253, 333, 423], [7, 12, 31, 51, 71, 101]],
  [39, [71, 111, 182, 252, 332, 422], [7, 12, 31, 51, 71, 101]],
  [40, [70, 110, 180, 250, 330, 420], [6, 11, 30, 50, 70, 100]]
];

function dsClassStrict(v, t) { const i = t.findIndex((s) => v < s); return i < 0 ? 6 : i; }
function dsClassInclusive(v, t) { const i = t.findIndex((s) => v <= s); return i < 0 ? 6 : i; }

function dsSmallThresholds(surface) {
  if (surface <= 8) return [DS_SMALL[0][1], DS_SMALL[0][2]];
  for (let k = 0; k + 1 < DS_SMALL.length; k++) {
    const [s0, e0, g0] = DS_SMALL[k], [s1, e1, g1] = DS_SMALL[k + 1];
    if (surface >= s0 && surface <= s1) {
      const f = (surface - s0) / (s1 - s0);
      return [e0.map((v, i) => Math.floor(v + (e1[i] - v) * f)), g0.map((v, i) => Math.floor(v + (g1[i] - v) * f))];
    }
  }
  return [DS_EP_2021, DS_GES_2021];
}

// → { ep: "D", ges: "C", note }
function dsLetters(gen, ep, ges, surface, etabli) {
  if (gen === "ancien") return { ep: DS_LETTERS[dsClassInclusive(ep, DS_EP_OLD)], ges: DS_LETTERS[dsClassInclusive(ges, DS_GES_OLD)], note: "" };
  const small = surface != null && surface < 40 && (!etabli || dsDateCmp(etabli, { y: 2024, m: 7, d: 1 }) >= 0);
  const [te, tg] = small ? dsSmallThresholds(surface) : [DS_EP_2021, DS_GES_2021];
  const ce = dsClassStrict(ep, te), cg = dsClassStrict(ges, tg);
  return { ep: DS_LETTERS[Math.max(ce, cg)], ges: DS_LETTERS[cg], note: small ? "barème petites surfaces (< 40 m²)" : "" };
}

/* ---------------------------------------------------------------- *
 * 6. OCR : mots, lignes reconstruites, étiquette                   *
 * ---------------------------------------------------------------- */

// TSV Tesseract (même format en ligne de commande et dans Tesseract.js).
function dsParseTsv(tsv) {
  const words = [];
  String(tsv || "").split(/\r?\n/).forEach((l) => {
    const c = l.split("\t");
    if (c.length < 12 || c[0] !== "5") return;
    const text = c.slice(11).join("\t").trim();
    if (!text) return;
    words.push({ text, x: +c[6], y: +c[7], w: +c[8], h: +c[9], conf: +c[10],
      line: `${c[2]}.${c[3]}.${c[4]}` });
  });
  return words;
}

// Texte « ligne par ligne » à partir des lignes Tesseract.
function dsWordsText(words) {
  const lines = [];
  let cur = null;
  for (const w of words) {
    if (!cur || cur.key !== w.line) { cur = { key: w.line, parts: [] }; lines.push(cur); }
    cur.parts.push(w.text);
  }
  return lines.map((l) => l.parts.join(" ")).join("\n");
}

// Lignes visuelles sur toute la largeur de la page (colonnes recollées) :
// « entre 1 230 € » + « et 1 700 € par an » écrits dans deux blocs.
function dsRowText(words) {
  if (!words || !words.length) return "";
  const hs = words.map((w) => w.h).sort((a, b) => a - b);
  const medH = hs[Math.floor(hs.length / 2)] || 10;
  const sorted = words.slice().sort((a, b) => (a.y + a.h / 2) - (b.y + b.h / 2));
  const rows = [];
  for (const w of sorted) {
    const cy = w.y + w.h / 2;
    const r = rows[rows.length - 1];
    if (r && Math.abs(cy - r.cy) <= medH * 0.6) { r.words.push(w); r.cy = (r.cy * (r.words.length - 1) + cy) / r.words.length; }
    else rows.push({ cy, words: [w] });
  }
  return rows.map((r) => r.words.sort((a, b) => a.x - b.x).map((w) => w.text).join(" ")).join("\n");
}

function dsLabelPairsFromText(t) {
  const out = [];
  const lines = String(t || "").split("\n");
  // Nombre seul sur sa ligne, juste au-dessus de l'unité kWh/m²/an
  for (let i = 0; i + 1 < lines.length; i++) {
    const unit = lines[i + 1];
    if (unit.includes("kwh") && (unit.includes("/m") || unit.includes("m2") || unit.includes("m?")) && !unit.includes("finale")) {
      const m = DS_RE.loneNum.exec(lines[i].trim());
      if (m) { const ep = +m[1]; if (ep >= 1 && ep <= 1500) out.push({ ep, ges: null, letter: null }); }
    }
  }
  for (const line of lines) {
    // Les températures, dates et unités ne sont pas l'étiquette
    if (line.includes("°c") || line.includes("/20") || line.includes("kwh") || line.length > 40) continue;
    // Téléphones, e-mails, montants : pas l'étiquette (« tel : 01 30 12 80 00 ’ »)
    if (/\btel\b|\bfax\b|@|€|%|\b0\d(?:[ .]?\d{2}){3,4}\b/.test(line)) continue;
    const re = DS_RE.labelPair;
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(line))) {
      const ep = +m[1], ges = +m[2];
      if (ep >= 1 && ep <= 1500 && ges <= 400) out.push({ ep, ges, letter: m[3] ? m[3].toUpperCase() : null });
      if (m[0] === "") re.lastIndex++;
    }
  }
  return out;
}

function dsWordNum(w) {
  const t = String(w.text).replace(/^[*°'"|\[\](){}.,:;!]+|[*°'"|\[\](){}.,:;!]+$/g, "");
  return /^\d{1,4}$/.test(t) ? +t : null;
}

// Lecture « spatiale » : nombre juste au-dessus des unités kWh/m²/an et kgCO2/m²/an.
function dsLabelPairsSpatial(words) {
  if (!words || !words.length) return [];
  const out = [];
  const units = words.filter((w) => { const n = dsNorm(w.text); return n.includes("kwh") && n.includes("/m") && !n.includes("ef"); });
  const above = (u) => {
    let best = null;
    for (const w of words) {
      const bottom = w.y + w.h, cx = w.x + w.w / 2;
      if (bottom <= u.y + u.h / 2 && (u.y - bottom) < u.h * 5 && cx >= u.x - u.w * 0.4 && cx <= u.x + u.w * 1.2) {
        const v = dsWordNum(w);
        if (v !== null && (!best || (u.y - bottom) < best.d)) best = { v, d: u.y - bottom };
      }
    }
    return best ? best.v : null;
  };
  for (const k of units) {
    const ep = above(k);
    if (ep === null) continue;
    const gu = words.find((w) => {
      const n = dsNorm(w.text);
      return (n.includes("kgco") || n.includes("co2/m") || n.includes("co,/m")) && w.x > k.x &&
        Math.abs((w.y + w.h / 2) - (k.y + k.h / 2)) < k.h * 1.5 && (w.x - k.x) < k.w * 4;
    });
    out.push({ ep, ges: gu ? above(gu) : null, letter: null });
  }
  return out;
}

function dsGesScaleValues(t) {
  const v = [];
  const re = DS_RE.gesScale;
  re.lastIndex = 0;
  let m;
  while ((m = re.exec(t || ""))) {
    const x = +m[1].replace(/\D/g, "");
    if (Number.isFinite(x) && x <= 400) v.push(x);
  }
  return v;
}

const dsClose = (a, b, rel, abs) => Math.abs(a - b) <= Math.max(abs, rel * Math.abs(b));

/* ---------------------------------------------------------------- *
 * 7. Repérage des pages                                            *
 * ---------------------------------------------------------------- */

// Score « page de synthèse DPE » d'un texte normalisé.
function dsPageScore(t) {
  let s = 0;
  const has = (x) => t.includes(x);
  const test = (re) => { re.lastIndex = 0; return re.test(t); };
  if (has("diagnostic de performance")) s += 2;
  if (test(DS_RE.etabli)) s += 3;
  if (test(DS_RE.valable)) s += 3;
  if (dsFindAdemeStrict(t)) s += 2;
  if (has("estimation des couts annuels")) s += 2;
  if (has("ce logement emet")) s += 2;
  if (has("performance energetique")) s += 1;
  if (DS_RE_INFO_DIAG.test(t)) s += 1;
  if (has("consommation conventionnelle") || has("estimation des emissions")) s += 2;
  if (has("consommations annuelles par energie")) s += 2;
  if (has("fiche technique")) s -= 3;
  if (has("montants et consommations annuels")) s -= 2;
  return s;
}
const DS_SCORE_FOUND = 9;
const DS_SCORE_MIN = 5;

function dsIsTotalsPage(t) { return t.includes("usages recenses") && t.includes("kwh"); }

function dsCountAlnum(s) { return (String(s).match(/[\p{L}\p{N}]/gu) || []).length; }

/* ---------------------------------------------------------------- *
 * 8. Extraction des champs                                         *
 * ---------------------------------------------------------------- */

// « Informations diagnostiqueur », y compris lu par OCR (« dlagnostiqueur »)
const DS_RE_INFO_DIAG = /informations?\s*d[il1|]a[gq]nost[il1|][qg]ueur/;
const DS_RE_DIAG_NAME = /(?<!informations?[ \t]{1,3})d[il]agnostiqueur[ \t]*[:;][ \t]*\n?[ \t]*([^\n]{3,60})/i;
const DS_RE_CERTIF = /n\s*[°o]\s*de\s*certification\s*[:;]?\s*([A-Za-z0-9][A-Za-z0-9 /._\-]{1,30})/i;
const DS_RE_CAB = /^\s*(?:cabinet|nom de la soci[ée]t[ée]|soci[ée]t[ée]|raison sociale|[ée]tablissement du rapport)\s*:\s*([^\n]{2,60})/im;
const DS_RE_ADDR = /^\s*adresse(?:\s*du\s*bien)?\s*:\s*(.+)$/im;

// Coupe un texte à la première « colonne » (3 espaces ou plus) et nettoie.
function dsFirstCol(s) { return String(s || "").split(/ {3,}/)[0].trim(); }

// Retire les « mots » d'un seul caractère en début/fin (bruit d'OCR : « SOCOTEC Diagnostic À »).
function dsTrimNoise(s, keepDigits = false) {
  const w = String(s || "").trim().split(/\s+/);
  const junk = (t) => (t.length === 1 && !(keepDigits && /\d/.test(t))) || /^[^\p{L}\p{N}]+$/u.test(t);
  while (w.length > 1 && junk(w[w.length - 1])) w.pop();
  while (w.length > 1 && junk(w[0])) w.shift();
  return w.join(" ").trim().replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N})]+$/gu, "");
}

function dsExtractPeople(src, rec) {
  const raws = [src.dpeRaw, src.dpeOcrRaw].filter(Boolean);
  // Société : première ligne exploitable après « Informations diagnostiqueur »
  outer: for (const raw of raws) {
    const lines = raw.split("\n").map((l) => l.trim());
    for (let i = 0; i < lines.length; i++) {
      if (!DS_RE_INFO_DIAG.test(dsNorm(lines[i]))) continue;
      for (const next of lines.slice(i + 1, i + 4)) {
        let s = dsFirstCol(next);
        const n = dsNorm(s);
        for (const cut of ["diagnostiqueur", "tel :", "tel:", "email", "le present rapport"]) {
          const pos = n.indexOf(cut);
          if (pos > 1) s = s.slice(0, pos); else if (pos >= 0) s = "";
        }
        s = s.trim().replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, "").trim();
        s = dsTrimNoise(s);
        if ((s.match(/\p{L}/gu) || []).length >= 2) { rec.societe = s; break outer; }
      }
    }
  }
  if (!rec.societe) {
    for (const raw of [...raws, src.allRaw]) {
      const m = DS_RE_CAB.exec(raw || "");
      if (m) { rec.societe = dsFirstCol(m[1].split("  ")[0]); break; }
    }
  }
  for (const raw of raws) {
    const m = DS_RE_DIAG_NAME.exec(raw);
    if (!m) continue;
    let s = dsFirstCol(m[1]);
    const n = dsNorm(s);
    for (const cut of ["tel", "email", "n°", "organisme"]) {
      const pos = n.indexOf(cut);
      if (pos > 2) s = s.slice(0, pos);
    }
    s = dsTrimNoise(s.replace(/[.,;:]+$/, ""));
    if ((s.match(/\p{L}/gu) || []).length >= 3) { rec.diagnostiqueur = s; break; }
  }
  for (const raw of raws) {
    const m = DS_RE_CERTIF.exec(raw);
    if (!m) continue;
    const s = (m[1].split("  ")[0].trim().split(/\s+/)[0] || "");
    if (s.length >= 2) { rec.certification = s; break; }
  }
  // Adresse du bien (1re « adresse : » de la page DPE ; CP sur la ligne suivante accepté)
  for (const raw of [...raws, src.allRaw]) {
    const m = DS_RE_ADDR.exec(raw || "");
    if (!m) continue;
    let a = dsFirstCol(m[1]);
    if (!/\b\d{5}\b/.test(a)) {
      const after = (raw || "").slice(m.index + m[0].length).split("\n").map((l) => l.trim()).filter(Boolean)[0] || "";
      if (/^\d{5}\s+\S/.test(after)) a += " " + dsFirstCol(after);
    }
    // Colonne de droite recollée par l'OCR : « … 75011 PARIS “ °n° : 2575E… »
    const cut = /\s+[^\p{L}\p{N}]*\s*(?:n\s*°|n°|[ée]tabli\s*le|valable|type\s*de\s*bien|ann[ée]e\s*de\s*construction|surface\s*habitable|propri[ée]taire)/iu.exec(a);
    if (cut && cut.index > 4) a = a.slice(0, cut.index);
    a = dsTrimNoise(a.replace(/[,;]\s*$/, ""), true);
    if (a.length >= 5) { rec.adresse = a; break; }
  }
  // Bloc diagnostiqueur (pour retrouver le fournisseur)
  const zone = [];
  for (const raw of raws) {
    const lines = raw.split("\n");
    const i = lines.findIndex((l) => DS_RE_INFO_DIAG.test(dsNorm(l)));
    if (i >= 0) zone.push(...lines.slice(i + 1, i + 11));
    lines.forEach((l) => { if (/e-?mail|@|www\.|cabinet|soci[ée]t[ée]|raison sociale|diagnostiqueur\s*:/i.test(l)) zone.push(l); });
  }
  rec.zone = zone.join("\n");
}

function dsNewRecord() {
  return {
    generation: null, nonOfficiel: false, societe: "", diagnostiqueur: "", certification: "", adresse: "", zone: "",
    etabli: null, valable: null, epM2: null, epTotal: null, lettreEp: "", gesM2: null, gesTotal: null, lettreGes: "",
    coutMin: null, coutMax: null, prixRef: null, ademe: "", surface: null, efM2: null, efTotal: null,
    remarques: [], notes: [], erreur: ""
  };
}

// Fourchette de coûts plausible (le max d'un DPE vaut ~1,3 à 1,6 × le min) :
// écarte les montants abîmés par l'OCR (« entre z 080 € et 1 470 € »).
function dsCostOk(a, b) { return a !== null && b !== null && a >= 10 && a <= b && b <= a * 2.2 && b < 100000; }

// src : { dpeText, dpeRaw, dpeOcrText, dpeOcrRaw, dpeOcrWords, dpeOcred, totalsText, allText, allRaw }
function dsExtract(src, rec = dsNewRecord()) {
  const dpeTexts = [src.dpeText, src.dpeOcrText].filter(Boolean);
  const all = src.allText || "";
  const joined = dpeTexts.join("\n");

  // --- Génération (DPE 2021 ou ancien) ---
  const looksOld = /consommation conventionnelle|estimation des emissions|consommations annuelles par energie|logement econome/.test(joined);
  const looks2021 = /estimation des couts annuels|ce logement emet|dont emissions de gaz/.test(joined) || DS_RE_INFO_DIAG.test(joined);
  if (/evaluation thermique|non valable pour une transaction|non emis/.test(joined)) {
    rec.nonOfficiel = true;
    rec.remarques.push("évaluation thermique / DPE non enregistré à l'ADEME");
  }

  // --- Dates ---
  rec.etabli = dsFirstDate(dpeTexts, [DS_RE.etabli, DS_RE.etabliLong]) ||
    dsFirstDate(dpeTexts, [DS_RE.dateEtabOld]) || dsFirstDate([all], [DS_RE.etabli, DS_RE.etabliLong]);
  rec.valable = dsFirstDate(dpeTexts, [DS_RE.valable, DS_RE.valableLong]) || dsFirstDate([all], [DS_RE.valable, DS_RE.valableLong]);
  if (!rec.etabli) {
    const d = dsFirstDate(dpeTexts, [DS_RE.dateVisite]);
    if (d) { rec.etabli = d; rec.remarques.push("« établi le » absent : date de visite utilisée"); }
  }
  const gen = looks2021 && !looksOld ? "2021" : looksOld && !looks2021 ? "ancien"
    : rec.etabli && dsDateCmp(rec.etabli, { y: 2021, m: 7, d: 1 }) < 0 ? "ancien" : "2021";
  rec.generation = gen;

  if (rec.etabli && rec.valable) {
    if (dsDateCmp(rec.valable, rec.etabli) <= 0) rec.remarques.push("date de validité antérieure à la date d'établissement");
    else if (gen === "2021") {
      const exp = dsPlusYearsMinusOneDay(rec.etabli, 10);
      if (dsDateCmp(rec.valable, exp) !== 0) rec.remarques.push(`validité ≠ établi + 10 ans (attendu ${dsDateFr(exp)})`);
    }
  }
  if (!rec.valable && !rec.nonOfficiel && rec.etabli) {
    rec.valable = dsPlusYearsMinusOneDay(rec.etabli, 10);
    rec.remarques.push(`« valable jusqu'au » non lu : calculé (${dsDateFr(rec.valable)})`);
  }
  if (gen === "ancien" && rec.etabli) {
    // DPE d'avant la réforme : validité légale raccourcie (décret 2020-1610).
    const lim = dsDateCmp(rec.etabli, { y: 2018, m: 1, d: 1 }) < 0 ? "31/12/2022" : "31/12/2024";
    rec.notes.push(`ancien DPE : validité légale jusqu'au ${lim}`);
  }

  // --- N° ADEME ---
  const ademeText = dsFindAdemeStrict(src.dpeText);
  let ademeOcr = dsFindAdemeStrict(src.dpeOcrText);
  if (!ademeOcr && src.dpeOcrText) {
    const re = DS_RE.ademeLabel;
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(src.dpeOcrText))) { const a = dsRepairAdeme(m[1]); if (a) { ademeOcr = a; break; } }
  }
  rec.ademe = ademeText || ademeOcr || dsFindAdemeStrict(all) || "";
  if (rec.ademe && !ademeText && ademeOcr === rec.ademe) rec.remarques.push("n° ADEME lu par OCR : à contrôler");

  // --- Surface ---
  for (const t of [...dpeTexts, all]) {
    const m = DS_RE.surface.exec(t || "");
    const v = m ? dsParseNum(m[1]) : null;
    if (v !== null && v > 3 && v < 5000) { rec.surface = v; break; }
  }

  // --- Totaux & coûts ---
  const totalsTexts = [src.totalsText, src.dpeText, src.dpeOcrText].filter(Boolean);
  if (gen === "2021") {
    for (const t of totalsTexts) {
      const m = DS_RE.total2021.exec(t) || DS_RE.total2021B.exec(t);
      if (!m) continue;
      rec.epTotal = dsParseNum(m[1]);
      rec.efTotal = m[2] ? dsParseNum(m[2]) : null;
      if (rec.efTotal === null) {
        const tail = t.slice(m.index + m[0].length, m.index + m[0].length + 80);
        const e = DS_RE.efTotal.exec(tail);
        if (e) rec.efTotal = dsParseNum(e[1]);
      }
      break;
    }
    for (const t of dpeTexts) {
      const m = DS_RE.gesTotal.exec(t);
      if (m) { rec.gesTotal = dsParseNum(m[1]); break; }
    }
    const coutRes = [DS_RE.cout, DS_RE.coutLoose, DS_RE.coutSwap];
    outer: for (const t of [...dpeTexts, src.totalsText].filter(Boolean)) {
      for (const re of coutRes) {
        const m = re.exec(t);
        if (!m) continue;
        const a = dsParseNum(m[1]), b = dsParseNum(m[2]);
        if (dsCostOk(a, b)) { rec.coutMin = a; rec.coutMax = b; break outer; }
      }
    }
    if (rec.coutMin === null && src.totalsText) {
      const m = DS_RE.coutTotalRow.exec(src.totalsText);
      if (m && dsCostOk(dsParseNum(m[1]), dsParseNum(m[2]))) { rec.coutMin = dsParseNum(m[1]); rec.coutMax = dsParseNum(m[2]); }
    }
    // Page DPE : deux montants juste après « entre », ou juste après le titre
    // « Estimation des coûts annuels » (ordre du flux PDF mélangé).
    if (rec.coutMin === null) {
      outer2: for (const t of dpeTexts) {
        const k = t.indexOf("estimation des couts annuels");
        for (const [re, txt] of [[DS_RE.coutNear, t], [DS_RE.twoAmounts, k >= 0 ? t.slice(k, k + 600) : ""]]) {
          const m = txt && re.exec(txt);
          if (!m) continue;
          const a = dsParseNum(m[1]), b = dsParseNum(m[2]);
          if (dsCostOk(a, b)) { rec.coutMin = a; rec.coutMax = b; break outer2; }
        }
      }
    }
  } else {
    for (const t of totalsTexts.concat([all])) {
      const m = DS_RE.totalOld.exec(t);
      if (!m) continue;
      rec.efTotal = dsParseNum(m[1]);
      rec.epTotal = dsParseNum(m[2]);
      const cost = dsParseNum(m[3]);
      rec.coutMin = cost; rec.coutMax = cost;
      break;
    }
  }

  // --- Énergie primaire & GES par m² ---
  const ocrText = src.dpeOcrText || "";
  const pairs = dsLabelPairsFromText(ocrText).concat(dsLabelPairsFromText(src.dpeText || ""), dsLabelPairsSpatial(src.dpeOcrWords));
  const scaleGes = dsGesScaleValues(ocrText).concat(dsGesScaleValues(src.dpeText || ""));

  if (gen === "2021") {
    const calcEp = rec.epTotal !== null && rec.surface ? Math.floor(rec.epTotal / rec.surface + 1e-6) : null;
    const calcGes = rec.gesTotal !== null && rec.surface ? Math.floor(rec.gesTotal / rec.surface + 1e-6) : null;
    dsChooseValues(rec, pairs, scaleGes, calcEp, calcGes, !!src.dpeOcred);
    if (rec.efTotal !== null && rec.surface) rec.efM2 = Math.floor(rec.efTotal / rec.surface + 1e-6);
    if (rec.efM2 === null) { const m = DS_RE.efM2.exec(ocrText); if (m) rec.efM2 = dsParseNum(m[1]); }
  } else {
    for (const t of dpeTexts) { const m = DS_RE.consoConv.exec(t); if (m) { rec.epM2 = dsParseNum(m[1]); break; } }
    for (const t of dpeTexts) { const m = DS_RE.emissionOld.exec(t); if (m) { rec.gesM2 = dsParseNum(m[1]); break; } }
    if (!src.dpeText && (rec.epM2 !== null || rec.gesM2 !== null)) rec.remarques.push("valeurs lues par OCR : à contrôler");
    if (rec.epTotal !== null && rec.surface && rec.epM2 !== null && !dsClose(rec.epTotal / rec.surface, rec.epM2, 0.06, 2)) {
      rec.remarques.push(`énergie primaire ${dsFmtNum(rec.epM2)} ≠ total/surface (${Math.round(rec.epTotal / rec.surface)})`);
    }
    if (rec.gesM2 !== null && rec.surface) {
      rec.gesTotal = Math.round(rec.gesM2 * rec.surface);
    }
    if (rec.efTotal !== null && rec.surface) rec.efM2 = Math.round(rec.efTotal / rec.surface * 100) / 100;
  }

  // --- Lettres ---
  if (rec.epM2 !== null && rec.gesM2 !== null) {
    const L = dsLetters(gen, rec.epM2, rec.gesM2, rec.surface, rec.etabli);
    rec.lettreEp = L.ep; rec.lettreGes = L.ges;
    if (L.note) rec.notes.push(L.note);
    const printed = pairs.find((p) => p.letter && dsClose(p.ep, rec.epM2, 0, 0.5) && p.ges !== null && dsClose(p.ges, rec.gesM2, 0, 0.5));
    if (printed && printed.letter !== L.ep) {
      rec.remarques.push(`lettre lue sur l'étiquette « ${printed.letter} » ≠ lettre calculée « ${L.ep} »${L.note ? ` (${L.note})` : ""}`);
    }
  }

  dsExtractPeople(src, rec);

  rec.prixRef = dsFirstDate([...dpeTexts, src.totalsText, all].filter(Boolean), [DS_RE.prixRefLong, DS_RE.prixRef]);

  if (!rec.ademe && !rec.nonOfficiel) rec.remarques.push("n° ADEME non trouvé");
  if (dsCoutsUtiles(rec)) {
    if (rec.coutMin === null) rec.remarques.push("estimation des coûts non trouvée");
    else if (!rec.prixRef) rec.remarques.push("date d'indexation des prix non trouvée");
  }
  return rec;
}

// Choix EP/GES (DPE 2021) : étiquette (texte ou OCR) recoupée avec total ÷ surface
// (les valeurs de l'étiquette sont tronquées à l'entier).
function dsChooseValues(rec, pairs, scaleGes, calcEp, calcGes, ocred) {
  const eps = pairs.map((p) => p.ep);
  if (calcEp !== null) {
    const hit = eps.find((v) => dsClose(v, calcEp, 0.015, 1));
    if (hit !== undefined) rec.epM2 = hit;
    else {
      rec.epM2 = calcEp;
      if (eps.length) rec.remarques.push(`énergie primaire : étiquette lue ${dsFmtNum(eps[0])} mais total ÷ surface = ${dsFmtNum(calcEp)} (valeur calculée retenue)`);
      else if (ocred) rec.notes.push("énergie primaire calculée (total ÷ surface), étiquette illisible");
    }
  } else if (eps.length) {
    rec.epM2 = eps[0];
    rec.remarques.push("énergie primaire lue sur l'étiquette uniquement (non recoupée) : à contrôler");
  }

  const ges = [];
  if (rec.epM2 !== null) pairs.filter((p) => dsClose(p.ep, rec.epM2, 0.015, 1) && p.ges !== null).forEach((p) => ges.push(p.ges));
  pairs.filter((p) => p.ges !== null).forEach((p) => ges.push(p.ges));
  ges.push(...scaleGes);
  if (calcGes !== null) {
    const hit = ges.find((v) => dsClose(v, calcGes, 0.03, 1));
    if (hit !== undefined) rec.gesM2 = hit;
    else {
      rec.gesM2 = calcGes;
      if (ges.length) rec.remarques.push(`GES : étiquette lue ${dsFmtNum(ges[0])} mais total ÷ surface = ${dsFmtNum(calcGes)} (valeur calculée retenue)`);
      else if (ocred) rec.notes.push("GES calculé (total ÷ surface), étiquette illisible");
    }
  } else {
    const label = (pairs.find((p) => p.ges !== null && (rec.epM2 === null || dsClose(p.ep, rec.epM2, 0, 0.5))) || {}).ges;
    const sc = scaleGes[0];
    if (label != null && sc != null && dsClose(label, sc, 0, 0.5)) rec.gesM2 = label;
    else if (label != null || sc != null) {
      rec.gesM2 = label != null ? label : sc;
      rec.remarques.push("GES lu sur l'étiquette uniquement (non recoupé) : à contrôler");
    }
  }
}

/* ---------------------------------------------------------------- *
 * 9. Fournisseurs (liste « NOM : n° FU »)                          *
 * ---------------------------------------------------------------- */

const DS_SUPPLIER_SEPS = ["->", "=>", "→", ";", "\t", "=", "|", ",", ":"];

// Texte de la liste → { entries:[{name, code, key}], warnings:[], empty:[noms sans n°] }
function dsParseSuppliers(text) {
  const entries = [], warnings = [], empty = [];
  const seen = new Map();
  String(text || "").split(/\r?\n/).forEach((raw, i) => {
    const line = raw.replace(/^﻿/, "").trim();
    if (!line || line.startsWith("#") || line.startsWith("//")) return;
    let k = null, v = null;
    for (const sep of DS_SUPPLIER_SEPS) {
      let pos = line.indexOf(sep);
      if (pos < 0) continue;
      if (sep === ":" || sep === ",") pos = line.lastIndexOf(sep); // les noms peuvent en contenir
      k = line.slice(0, pos); v = line.slice(pos + sep.length);
      break;
    }
    if (k === null) { warnings.push(`ligne ${i + 1} ignorée (pas de séparateur) : « ${line} »`); return; }
    k = k.trim().replace(/^"|"$/g, "").trim();
    v = v.trim().replace(/^"|"$/g, "").trim();
    const key = dsKey(k);
    if (!entries.length && /^(CLE|KEY|NOM|SOCIETE|FOURNISSEUR)$/.test(key) && !/\d/.test(v)) return; // en-tête
    if (key.length < 2) return;
    if (!v) { empty.push(k); return; }
    const prev = seen.get(key);
    if (prev) {
      if (prev.code !== v) warnings.push(`« ${k} » a deux numéros (${prev.code} et ${v}) : le premier est utilisé`);
      return;
    }
    if (!/^(FU)?\d{5,8}$/i.test(v)) warnings.push(`« ${k} » : numéro inhabituel « ${v} »`);
    const e = { name: k, code: v, key };
    seen.set(key, e);
    entries.push(e);
  });
  return { entries, warnings, empty };
}

function dsLevenshtein(a, b, max = 3) {
  if (Math.abs(a.length - b.length) > max) return max + 1;
  let prev = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    let rowMin = i;
    for (let j = 1; j <= b.length; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      rowMin = Math.min(rowMin, cur[j]);
    }
    if (rowMin > max) return max + 1;
    prev = cur;
  }
  return prev[b.length];
}

// Cherche le fournisseur : société lue > bloc diagnostiqueur > page DPE > document.
// ctx : { societe, zone, dpeText, allText } (textes bruts). → { entry, how, fuzzy, alternatives } | null
function dsFindSupplier(entries, ctx) {
  if (!entries || !entries.length) return null;
  const levels = [["société", ctx.societe, true], ["bloc diagnostiqueur", ctx.zone, true], ["page DPE", ctx.dpeText, false], ["document", ctx.allText, false]];
  for (const [how, text, tokenSet] of levels) {
    if (!text) continue;
    const hay = " " + dsKey(text) + " ";
    const words = new Set(hay.trim().split(" "));
    const hits = entries.filter((e) => hay.includes(" " + e.key + " ") ||
      (tokenSet && e.key.includes(" ") && e.key.split(" ").every((w) => w.length > 1 && words.has(w))));
    if (!hits.length) continue;
    hits.sort((a, b) => b.key.length - a.key.length);
    const alternatives = hits.filter((h) => h.code !== hits[0].code && !hits[0].key.includes(h.key));
    return { entry: hits[0], how, fuzzy: false, alternatives };
  }
  // Rapprochement approximatif (fautes de frappe : SOCOTECH ≈ SOCOTEC, DIAGMTER ≈ DIAGAMTER)
  const cands = [];
  for (const text of [ctx.societe, (ctx.zone || "").split("\n").slice(0, 3).join(" ")]) {
    const ws = dsKey(text).split(" ").filter(Boolean);
    for (let n = 1; n <= 4; n++) for (let i = 0; i + n <= ws.length; i++) cands.push({ c: ws.slice(i, i + n).join(""), s: ws.slice(i, i + n).join(" ") });
  }
  // Pas de rapprochement pour les noms courts ou contenant des chiffres
  // (DIAG3, DIAG 75, 2DI : une lettre de différence change de société).
  let best = null;
  for (const e of entries) {
    const k = e.key.replace(/ /g, "");
    if (k.length < 6 || /\d/.test(k)) continue;
    const max = k.length >= 10 ? 2 : 1;
    for (const { c, s } of cands) {
      if (c.length < 5 || /\d/.test(c)) continue;
      const d = dsLevenshtein(k, c, max);
      if (d <= max && (!best || d < best.d || (d === best.d && k.length > best.k.length))) best = { e, d, k, s };
    }
  }
  return best ? { entry: best.e, how: "rapprochement", fuzzy: true, matched: best.s, alternatives: [] } : null;
}

/* ---------------------------------------------------------------- *
 * 10. Rattachement PDF ↔ lignes                                    *
 * ---------------------------------------------------------------- */

const DS_FILE_RE = /\.(pdf|png|jpe?g)$/i;

// « …/IM0393664_10_DPE_2022-10-19 (1).pdf » → « im0393664_10_dpe_2022-10-19 (1) »
function dsBaseName(p) {
  return String(p || "").split(/[\\/]/).pop().replace(/\.[A-Za-z0-9]{1,5}$/, "").trim().toLowerCase();
}

function dsNormIm(v) {
  const d = String(v ?? "").replace(/^\s*IM/i, "").replace(/\D/g, "").replace(/^0+/, "");
  return d;
}
function dsNormLot(v) {
  return String(v ?? "").trim().toUpperCase().replace(/^LT\s*/, "").replace(/\s+/g, "").replace(/^0+(?=.)/, "");
}

// Nom produit par « DPE SIGEO » ({IM}_{lot}_DPE_{date}) → { im, lot, date }
function dsParseFileName(name) {
  const base = String(name || "").split(/[\\/]/).pop().replace(/\.[A-Za-z0-9]{1,5}$/, "").replace(/\s*\(\d+\)$/, "");
  const m = /^(?:IM)?0*(\d{3,7})[_ -]+(?:LT)?0*([A-Z0-9]+?)[_ -]+DPE(?:[_ -]+(\d{4}-\d{2}-\d{2}))?/i.exec(base);
  return m ? { im: m[1], lot: dsNormLot(m[2]), date: m[3] || "" } : null;
}

// rows : tableau complet ; idxs : lignes à traiter ; cols : { file, im, lot } (index, -1 si absent)
// files : [{ name }] → { byRow: Map(rowIdx → fileIdx), via: Map(rowIdx → "fichier"|"IM + lot"),
//                       dupes: [fileIdx], unmatched: [fileIdx] }
function dsAssignFiles(rows, idxs, cols, files) {
  const byRow = new Map(), via = new Map(), used = new Set(), dupes = [];
  const cell = (r, i) => (i >= 0 && r ? String(r[i] ?? "").trim() : "");
  const byBase = new Map();
  files.forEach((f, i) => { const b = dsBaseName(f.name); if (!byBase.has(b)) byBase.set(b, i); });
  // 1) colonne « DPE fichier » (nom exact)
  for (const ri of idxs) {
    const v = cell(rows[ri], cols.file);
    if (!v) continue;
    const fi = byBase.get(dsBaseName(v));
    if (fi !== undefined) { byRow.set(ri, fi); via.set(ri, "fichier"); used.add(fi); }
  }
  // 2) IM + lot lus dans le nom du fichier
  const byKey = new Map();
  files.forEach((f, i) => {
    const p = dsParseFileName(f.name);
    if (!p) return;
    const k = p.im + "|" + p.lot;
    if (!byKey.has(k)) byKey.set(k, []);
    byKey.get(k).push({ i, date: p.date });
  });
  for (const list of byKey.values()) list.sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  if (cols.im >= 0 && cols.lot >= 0) {
    for (const ri of idxs) {
      if (byRow.has(ri)) continue;
      const k = dsNormIm(cell(rows[ri], cols.im)) + "|" + dsNormLot(cell(rows[ri], cols.lot));
      const list = byKey.get(k);
      if (!list || !list.length) continue;
      const pick = list.find((x) => !used.has(x.i)) || list[0];
      byRow.set(ri, pick.i); via.set(ri, "IM + lot"); used.add(pick.i);
    }
  }
  // Fichiers en double pour un même lot (le plus récent est retenu)
  for (const list of byKey.values()) {
    if (list.length < 2 || !list.some((x) => used.has(x.i))) continue;
    list.forEach((x) => { if (!used.has(x.i)) { dupes.push(x.i); used.add(x.i); } });
  }
  const unmatched = files.map((_, i) => i).filter((i) => !used.has(i));
  return { byRow, via, dupes, unmatched };
}

/* ---------------------------------------------------------------- *
 * 11. Analyse d'un document (orchestration)                        *
 * ---------------------------------------------------------------- */

// Texte d'une page pdf.js : lignes reconstituées (y), espaces d'après les
// écarts x, et 3 espaces entre deux colonnes éloignées.
async function dsPdfPageText(page) {
  const tc = await page.getTextContent();
  const rows = new Map();
  for (const it of tc.items) {
    if (!it.str) continue;
    const key = Math.round(it.transform[5] / 2) * 2;
    if (!rows.has(key)) rows.set(key, []);
    rows.get(key).push({ x: it.transform[4], str: it.str, w: it.width || 0, h: it.height || Math.abs(it.transform[3]) || 0 });
  }
  return [...rows.entries()].sort((a, b) => b[0] - a[0]).map(([, items]) => {
    items.sort((a, b) => a.x - b.x);
    let line = "", endX = null, h = 0;
    for (const it of items) {
      if (endX !== null) {
        const ref = Math.max(h, it.h, 6), gap = it.x - endX;
        if (gap > ref * 3) line = line.replace(/ +$/, "") + "   ";
        else if (gap > ref * 0.12 && !/ $/.test(line) && !/^ /.test(it.str)) line += " ";
      }
      line += it.str;
      endX = it.x + it.w;
      h = Math.max(h, it.h);
    }
    return line.replace(/\s+$/, "");
  }).filter((l) => l.trim()).join("\n");
}

// Ouvre un PDF avec pdf.js et lit la couche texte de chaque page.
async function dsReadPdf(pdfjs, bytes) {
  const doc = await pdfjs.getDocument({ data: bytes, useSystemFonts: true }).promise;
  const pages = [];
  for (let p = 1; p <= doc.numPages; p++) pages.push(await dsPdfPageText(await doc.getPage(p)));
  return { doc, pages };
}

// pages : textes bruts de la couche texte ([] pour une image).
// io : { ocr(pageIdx, zone) → Promise<{ text, words, blank }|null>, stopped(), log(msg) }
// opts : { ocr: "auto" | "always" | "never", suppliers: [entries] }
async function dsAnalyzeDoc(pages, io, opts = {}) {
  const mode = opts.ocr || "auto";
  const n = Math.max(1, pages.length);
  const st = Array.from({ length: n }, (_, i) => {
    const raw = pages[i] || "";
    return { raw, norm: dsNorm(raw), hasText: dsCountAlnum(raw) >= 150, ocr: null, ocrNorm: "" };
  });
  const ocrCount = { n: 0 };
  const combined = (s) => s.norm + (s.ocr ? "\n" + s.ocrNorm : "");
  const doOcr = async (i, zone = false) => {
    if (mode === "never" || !io || !io.ocr) return false;
    if (io.stopped && io.stopped()) throw new Error("arrêté");
    const r = await io.ocr(i, zone);
    if (!r || r.blank) return false;
    ocrCount.n++;
    const s = st[i];
    const rowTxt = dsRowText(r.words || []);
    if (zone) {
      s.zoneText = (s.zoneText ? s.zoneText + "\n" : "") + (r.text || "") + "\n" + rowTxt;
    } else {
      s.ocr = { text: r.text || "", rowText: rowTxt, words: r.words || [] };
    }
    s.ocrNorm = dsNorm([s.ocr ? s.ocr.text : "", s.ocr ? s.ocr.rowText : "", s.zoneText || ""].join("\n"));
    return true;
  };

  // 1) Meilleure page parmi celles qui ont du texte
  let best = null;
  st.forEach((s, i) => { if (s.hasText) { const sc = dsPageScore(s.norm); if (!best || sc > best.sc) best = { i, sc }; } });
  // 2) Sinon : OCR des pages sans texte, dans l'ordre, jusqu'à trouver la page DPE
  const warnings = [];
  if (!best || best.sc < DS_SCORE_FOUND) {
    if (mode === "never" || !io || !io.ocr) {
      if (st.some((s) => !s.hasText)) warnings.push("pages scannées non lues (OCR désactivé)");
    } else {
      for (let i = 0; i < n; i++) {
        if (st[i].hasText) continue;
        if (!(await doOcr(i))) continue;
        const sc = dsPageScore(combined(st[i]));
        if (!best || sc > best.sc) best = { i, sc };
        if (sc >= DS_SCORE_FOUND) break;
      }
    }
  }
  if (!best) return { error: "aucun texte exploitable (document vide ou scanné sans OCR)", ocrPages: ocrCount.n };
  const di = best.i;
  if (best.sc < DS_SCORE_MIN) warnings.push(`page DPE non identifiée avec certitude (score ${best.sc})`);

  // 3) OCR de la page DPE elle-même (étiquette en image), si demandé
  if (mode === "always" && !st[di].ocr) await doOcr(di);

  // 4) Page des consommations (p.3) : page DPE ou 5 suivantes
  const findTotals = () => {
    if (dsIsTotalsPage(combined(st[di]))) return di;
    for (let i = di + 1; i < Math.min(n, di + 6); i++) if ((st[i].hasText || st[i].ocr) && dsIsTotalsPage(combined(st[i]))) return i;
    return -1;
  };
  let ti = findTotals();
  const looksOld = /consommation conventionnelle|estimation des emissions/.test(combined(st[di]));
  if (ti < 0 && !looksOld && (mode === "always" || !st[di].hasText)) {
    for (let i = di + 1; i < Math.min(n, di + 6) && ti < 0; i++) {
      if (st[i].hasText || st[i].ocr) continue;
      if ((await doOcr(i)) && dsIsTotalsPage(combined(st[i]))) ti = i;
    }
  }

  const build = () => {
    const d = st[di];
    const others = st.filter((_, i) => i !== di);
    return {
      dpeText: d.hasText ? d.norm : "",
      dpeRaw: d.hasText ? d.raw : "",
      dpeOcrText: d.ocrNorm || "",
      dpeOcrRaw: d.ocr ? d.ocr.text + "\n" + (d.zoneText || "") : (d.zoneText || ""),
      dpeOcrWords: d.ocr ? d.ocr.words : [],
      dpeOcred: !!d.ocr,
      totalsText: ti >= 0 ? combined(st[ti]) : "",
      allText: [combined(d), ...others.map(combined)].join("\n"),
      allRaw: [d.raw, d.ocr ? d.ocr.text : "", ...others.map((s) => s.raw + (s.ocr ? "\n" + s.ocr.text : ""))].join("\n")
    };
  };
  let src = build();
  let rec = dsExtract(src);
  const missing = (r) => !r.etabli || !r.valable || r.epM2 === null || r.gesM2 === null || (!r.ademe && !r.nonOfficiel);

  // 5) Mode auto : OCR de la page DPE si des valeurs manquent (étiquette, en-tête)
  if (mode === "auto" && missing(rec) && !st[di].ocr) {
    if (await doOcr(di)) {
      if (ti < 0) ti = findTotals();
      src = build(); rec = dsExtract(src);
    }
  }
  // 6) Zone contrastée (en-tête n°/dates en gris clair, étiquette) si toujours incomplet
  if (mode !== "never" && st[di].ocr && missing(rec)) {
    if (await doOcr(di, true)) { src = build(); rec = dsExtract(src); }
  }

  rec.page = di + 1;
  rec.pages = n;
  rec.ocrPages = ocrCount.n;
  rec.methode = !st.some((s) => s.hasText) ? `OCR (${ocrCount.n} page(s))`
    : st[di].hasText ? (st[di].ocr ? "texte + OCR étiquette" : "texte") : `OCR (${ocrCount.n} page(s))`;
  if (rec.generation === "2021" && rec.epM2 !== null && !st[di].ocr && rec.epTotal !== null) rec.methode += " (total ÷ surface)";
  rec.remarques.push(...warnings);

  // Fournisseur
  const sup = dsFindSupplier(opts.suppliers || [], { societe: rec.societe, zone: rec.zone, dpeText: src.dpeRaw || src.dpeOcrRaw, allText: src.allRaw });
  rec.supplier = sup;
  if (sup) {
    if (sup.fuzzy) rec.remarques.push(`fournisseur rapproché : « ${sup.matched} » ≈ « ${sup.entry.name} » de la liste`);
    else if (sup.how === "page DPE" || sup.how === "document") rec.remarques.push(`fournisseur « ${sup.entry.name} » trouvé dans le texte (${sup.how}), pas dans le bloc diagnostiqueur`);
    if (sup.alternatives.length) rec.remarques.push(`autres fournisseurs possibles : ${sup.alternatives.map((a) => a.name).join(", ")}`);
    if (!/^(FU)?\d{5,8}$/i.test(sup.entry.code)) rec.remarques.push(`n° FU inhabituel dans la liste : « ${sup.entry.code} »`);
  } else {
    rec.remarques.push(rec.societe ? `fournisseur absent de la liste : « ${rec.societe} »` : "fournisseur non trouvé");
  }
  return rec;
}

function dsMissing(rec) {
  const m = [];
  if (!rec.etabli) m.push("établi le");
  if (!rec.valable) m.push("valable jusqu'au");
  if (rec.epM2 === null) m.push("énergie primaire");
  if (rec.gesM2 === null) m.push("émission GES");
  return m;
}

function dsStatus(rec) {
  if (rec.error) return "ÉCHEC";
  return dsMissing(rec).length || rec.remarques.length ? "À VÉRIFIER" : "OK";
}

// Estimations (coûts + date d'indexation) saisissables dans SIGEO seulement
// pour un DPE établi depuis le 01/07/2021.
function dsCoutsUtiles(rec) {
  return !!rec.etabli && dsDateCmp(rec.etabli, { y: 2021, m: 7, d: 1 }) >= 0;
}

// Valeurs des colonnes DS_COLS pour un résultat.
function dsRecordValues(rec, fileName) {
  if (rec.error) {
    const v = DS_COLS.map(() => "");
    v[0] = "ÉCHEC"; v[1] = fileName || ""; v[DS_COLS.length - 1] = rec.error;
    return v;
  }
  const miss = dsMissing(rec);
  const rem = [...(miss.length ? [`non trouvé : ${miss.join(", ")}`] : []), ...rec.remarques, ...rec.notes];
  const sup = rec.supplier;
  const couts = dsCoutsUtiles(rec);
  return [
    dsStatus(rec), fileName || "", sup ? sup.entry.code : "",
    dsDateFr(rec.etabli), dsDateFr(rec.valable), rec.ademe || "",
    dsFmtNum(rec.epM2), rec.lettreEp || "", dsFmtNum(rec.gesM2), rec.lettreGes || "",
    couts ? dsFmtNum(rec.coutMin) : "", couts ? dsFmtNum(rec.coutMax) : "", couts ? dsDateFr(rec.prixRef) : "",
    rem.join(" ; ")
  ];
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    DS_COLS, DS_TYPE_DIAG, DS_DEFAULT_SUPPLIERS, dsCoutsUtiles, DS_FILE_RE, dsClean, dsNorm, dsKey, dsParseNum, dsFmtNum, dsDate, dsDateFr,
    dsPlusYearsMinusOneDay, dsFindAdemeStrict, dsRepairAdeme, dsLetters, dsParseTsv, dsWordsText, dsRowText,
    dsLabelPairsFromText, dsLabelPairsSpatial, dsPageScore, dsExtract, dsNewRecord, dsParseSuppliers, dsFindSupplier,
    dsLevenshtein, dsBaseName, dsNormIm, dsNormLot, dsParseFileName, dsAssignFiles, dsPdfPageText, dsReadPdf,
    dsAnalyzeDoc, dsMissing, dsStatus, dsRecordValues
  };
}
