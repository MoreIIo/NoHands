/* ====================================================================
 * NoHands OSA — DPE le plus récent par lot (SIGEO, page Documents)
 * --------------------------------------------------------------------
 * Onglet Extraction, type « DPE SIGEO (vérifier + télécharger) ».
 *
 * Pour chaque ligne : ouvre la page Documents du lot
 *   …/popup.aspx/fr/fil/ged_multi_file_manager/sig_imm/lot_manage?…
 * (ou la fiche lot, d'où le lien vers cette page est suivi automatiquement),
 * attend la fin du « Chargement en cours... », lit les sections dont le
 * titre contient « EPC » (« Diagnostics >> EPC », « Diagnostics Multys >>
 * EPC »), filtre les fichiers DPE par leur nom, garde le plus récent et le
 * télécharge via l'URL ged_file_download du onclick popup(…) du lien,
 * (même origine → cookies de session), sans cliquer sur le lien.
 *
 * Repérage des sections : par TEXTE de titre, jamais par position ni par
 * id. Chaque tableau de fichiers (lignes avec un lien __doPostBack) est
 * rattaché au dernier texte contenant « >> » situé entre le tableau
 * précédent et lui (à défaut, au dernier texte court).
 *
 * Chargé par sidepanel.html AVANT sidepanel.js ; importable sous Node
 * (tests). dpePageOp est injectée telle quelle dans la page : elle doit
 * rester autonome (aucune référence au reste du module).
 * ==================================================================== */

// Règles par défaut (expressions régulières, appliquées au texte
// normalisé : sans accents, minuscules, « ? » et « � » supprimés — ainsi
// « chaudie?re » devient « chaudiere » et « ?nerg?tique » « nergtique »).
const DPE_DEFAULT_RULES = Object.freeze({
  section: "epc",
  include: [
    "(?<![a-z])dpe(?![a-z])",
    "diagnostics? de performance e?nerge?tique",
    "performance e?nerge?tique"
  ],
  exclude: [
    "gaz", "elec", "chaudie?re", "entretien", "(?<![a-z])ent(?![a-z])", "amiante",
    "plomb", "termite", "carrez", "(?<![a-z])erp(?![a-z])", "(?<![a-z])ernmt(?![a-z])"
  ],
  // L'exclusion ne s'applique pas si ce motif est présent (DPE explicite).
  keep: "(?<![a-z])dpe(?![a-z])"
});

const DPE_DEFAULT_NAME = "{IM}_{lot}_DPE_{date}";
const DPE_DEFAULT_FOLDER = "DPE_SIGEO";
const DPE_DEFAULT_URL = "{URL lot}";
// Repli quand rien n'est trouvé en gérance (metier=ger) : la fiche lot en
// syndic, …/page.aspx/fr/sig_imm/lot_manage/syn/{immeuble}/{lot}
// (ex. …/syn/393151/90) ; la page Documents y est suivie automatiquement.
const DPE_DEFAULT_URL_SYN = "http://sigeo.evoriel.net/page.aspx/fr/sig_imm/lot_manage/syn/{IM_format}/{lot}";

/* ---------------------------------------------------------------- *
 * 1. Logique pure                                                  *
 * ---------------------------------------------------------------- */

// Même normalisation que foldText (sidepanel/content), plus la suppression
// des caractères d'encodage abîmé.
function dpeFold(s) {
  return String(s ?? "")
    .replace(/[   -​  　﻿]/g, " ")
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[?�]/g, "")
    .replace(/\s+/g, " ").trim().toLowerCase();
}

// Règles saisies (une regex par ligne, ou tableau) → regex compilées.
// Une regex invalide est signalée et ignorée (jamais d'exception).
function dpeCompileRules(r = {}) {
  const lines = (v, def) => (v === undefined || v === null ? def : Array.isArray(v) ? v : String(v).split(/\r?\n/))
    .map((x) => String(x).trim()).filter(Boolean);
  const errors = [];
  const rx = (src) => { try { return new RegExp(src, "i"); } catch (e) { errors.push(`regex invalide « ${src} »`); return null; } };
  const one = (v, def) => { const s = String(v ?? def ?? "").trim(); return s ? rx(s) : null; };
  return {
    section: one(r.section, DPE_DEFAULT_RULES.section),
    include: lines(r.include, DPE_DEFAULT_RULES.include).map(rx).filter(Boolean),
    exclude: lines(r.exclude, DPE_DEFAULT_RULES.exclude).map(rx).filter(Boolean),
    keep: one(r.keep, DPE_DEFAULT_RULES.keep),
    errors
  };
}

function dpeIsSection(title, rules) {
  return !!rules.section && rules.section.test(dpeFold(title));
}

function dpeIsDpe(name, rules) {
  const f = dpeFold(name);
  if (!rules.include.some((r) => r.test(f))) return false;
  if (rules.keep && rules.keep.test(f)) return true;
  return !rules.exclude.some((r) => r.test(f));
}

// « 19/10/2022 » → { iso: "2022-10-19", ts } ; null si illisible.
function dpeParseDate(s) {
  const m = String(s ?? "").match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
  if (!m) return null;
  const d = +m[1], mo = +m[2], y = +m[3];
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  const iso = `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  return { iso, ts: Date.UTC(y, mo - 1, d) };
}

const dpeNameYear = (name) => {
  const all = String(name ?? "").match(/20\d\d/g);
  return all ? Math.max(...all.map(Number)) : 0;
};

// sections : [{ title, files:[{ name, desc, date, target, order }] }]
// → { sections (EPC), candidates, chosen }
// Plus récent : date de la colonne Date, puis année dans le nom, puis
// ordre dans le tableau (le premier gagne).
function dpeSelect(sections, rulesIn) {
  const rules = rulesIn && rulesIn.include ? rulesIn : dpeCompileRules(rulesIn);
  const epc = (sections || []).filter((s) => dpeIsSection(s.title, rules));
  const candidates = [];
  let order = 0;
  for (const s of epc) {
    for (const f of s.files || []) {
      const o = f.order ?? order;
      order = o + 1;
      if (!dpeIsDpe(f.name, rules)) continue;
      const d = dpeParseDate(f.date);
      candidates.push(Object.assign({}, f, { section: s.title, iso: d ? d.iso : "", ts: d ? d.ts : -Infinity, year: dpeNameYear(f.name), order: o }));
    }
  }
  const sorted = candidates.slice().sort((a, b) => (b.ts - a.ts) || (b.year - a.year) || (a.order - b.order));
  return { sections: epc, candidates, chosen: sorted[0] || null };
}

// Nom de fichier Windows propre. tpl : « {IM}_{lot}_DPE_{date} » ; les
// variables de colonnes sont résolues par `resolve` (resolveRowTemplate
// dans le panneau), {date} = date ISO du fichier SIGEO.
function dpeSanitize(s) {
  return String(s ?? "")
    .replace(/[\\/:*?"<>|\u0000-\u001F]/g, "")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^[_.]+|[_. ]+$/g, "");
}

function dpeExtOf(filename) {
  const m = String(filename ?? "").match(/\.([A-Za-z0-9]{1,5})\s*$/);
  return m ? "." + m[1].toLowerCase() : "";
}

function dpeFileName(tpl, vars, resolve) {
  const nom = String(vars.nom ?? "").replace(/\.[A-Za-z0-9]{1,5}\s*$/, "");
  let s = String(tpl || DPE_DEFAULT_NAME).replace(/\{date\}/gi, vars.date || "sans-date").replace(/\{nom\}/gi, () => nom);
  if (resolve) s = resolve(s);
  s = s.replace(/\{([^{}]+)\}/g, (w, k) => (vars[k] ?? vars[k.toLowerCase()] ?? ""));
  const base = dpeSanitize(s) || "DPE";
  const ext = dpeExtOf(vars.serverName) || ".pdf";
  return base + ext;
}

// Dossier relatif accepté par chrome.downloads (pas de « .. », pas de lecteur).
function dpeFolder(f) {
  return String(f ?? "").split(/[\\/]+/).map(dpeSanitize).filter((p) => p && p !== "." && p !== "..").join("/");
}

/* ---------------------------------------------------------------- *
 * 2. Fonction injectée dans la page SIGEO (autonome)               *
 * ---------------------------------------------------------------- *
 * cfg.op :
 *   "scan"     {}            → état de la page Documents (sections, fichiers)
 *   "parse"    { html, baseUrl } → idem sur un HTML (tests)
 * Chaque fichier porte `download` : URL absolue lue dans le onclick
 *   popup('/bare.aspx/fr/fil/ged_file_download/<ID>', …) du lien (le
 *   href __doPostBack n'est PAS utilisé : il renvoie la page Documents).
 */
function dpePageOp(cfg) {
  const norm = (s) => String(s ?? "").replace(/[   ]/g, " ").replace(/\s+/g, " ").trim();
  const fold = (s) => norm(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const DATE = /\b\d{1,2}[\/.-]\d{1,2}[\/.-]\d{4}\b/;
  const HEAD = /^(nom|description|date|taille|type|auteur)$/;
  const LOADING = /chargement en cours/i;

  // onclick="popup('/bare.aspx/fr/fil/ged_file_download/123', name, …)"
  const popupOf = (a) => {
    const m = (a.getAttribute("onclick") || "").match(/popup\(\s*['"]([^'"]+)['"]/);
    return m ? m[1] : "";
  };
  const targetOf = (a) => {
    const code = (a.getAttribute("href") || "") + " " + (a.getAttribute("onclick") || "");
    const m = code.match(/__doPostBack\(\s*['"]([^'"]+)['"]\s*,\s*['"]([^'"]*)['"]/) ||
      code.match(/WebForm_PostBackOptions\(\s*["']([^"']+)["']\s*,\s*["']([^"']*)["']/);
    return m ? { target: m[1], arg: m[2] } : null;
  };

  function scan(doc, url) {
    const body = doc.body;
    if (!body) return { ready: false, loading: true, titles: [], sections: [], gedLinks: [] };
    const login = !!doc.querySelector('input[type="password"]');
    // Texte « Chargement en cours... » visible = page pas prête.
    let loading = false;
    const walkL = doc.createTreeWalker(body, 4 /* SHOW_TEXT */);
    for (let n = walkL.nextNode(); n; n = walkL.nextNode()) {
      if (!LOADING.test(n.nodeValue)) continue;
      const el = n.parentElement;
      const hidden = el && (el.closest('[style*="display:none"],[style*="display: none"],[hidden]') ||
        (doc.defaultView && doc.defaultView.getComputedStyle && el.isConnected && (() => {
          for (let e = el; e && e !== body; e = e.parentElement) {
            const cs = doc.defaultView.getComputedStyle(e);
            if (cs && (cs.display === "none" || cs.visibility === "hidden")) return true;
          }
          return false;
        })()));
      if (!hidden) { loading = true; break; }
    }

    // Lignes de fichiers : <tr> contenant un lien __doPostBack nommé.
    const grids = new Map(); // table → [rows]
    let order = 0;
    for (const tr of Array.from(doc.querySelectorAll("tr"))) {
      if (tr.closest("thead")) continue;
      const links = Array.from(tr.querySelectorAll("a")).filter((a) => (targetOf(a) || popupOf(a)) && norm(a.textContent));
      if (!links.length) continue;
      // Ligne d'un tableau imbriqué : on ne prend que la ligne la plus profonde.
      if (Array.from(tr.querySelectorAll("tr")).length) continue;
      const a = links[0];
      const name = norm(a.textContent);
      if (HEAD.test(fold(name))) continue; // lien de tri d'en-tête
      const cells = Array.from(tr.children).filter((c) => /^(TD|TH)$/.test(c.tagName));
      const nameCell = cells.find((c) => c.contains(a));
      const dateCell = cells.find((c) => c !== nameCell && DATE.test(norm(c.textContent)));
      const descCell = cells.find((c) => c !== nameCell && c !== dateCell && norm(c.textContent));
      const pb = targetOf(a) || {};
      let download = "";
      const pop = popupOf(a);
      if (pop) { try { download = new URL(pop, url || undefined).href; } catch (_) { download = pop; } }
      const table = tr.closest("table");
      if (!grids.has(table)) grids.set(table, []);
      grids.get(table).push({
        name,
        desc: descCell ? norm(descCell.textContent) : "",
        date: dateCell ? (norm(dateCell.textContent).match(DATE) || [""])[0] : "",
        target: pb.target || "", download, order: order++
      });
    }

    // Titres. 1) <fieldset><legend>Diagnostics >> EPC</legend> (structure
    // SIGEO) : légende du fieldset le plus proche du tableau. 2) À défaut,
    // texte le plus proche avant le tableau.
    const legendOf = (el) => {
      for (let e = el; e && e !== body; e = e.parentElement) {
        if (e.tagName !== "FIELDSET") continue;
        const lg = Array.from(e.children).find((c) => c.tagName === "LEGEND");
        if (lg && norm(lg.textContent)) return norm(lg.textContent);
      }
      return "";
    };
    const isTitle = (t) => />>|»/.test(t);
    const gridTables = Array.from(grids.keys());
    const inGrid = (node) => gridTables.some((t) => t.contains(node));
    const sections = [];
    const titles = [];
    let gap = []; // textes vus depuis le tableau précédent
    const walk = doc.createTreeWalker(body, 1 | 4 /* ELEMENT | TEXT */);
    const seen = new Set();
    for (let n = walk.nextNode(); n; n = walk.nextNode()) {
      if (n.nodeType === 1) {
        if (grids.has(n) && !seen.has(n)) {
          seen.add(n);
          // Titre = texte le plus proche avant le tableau (hors liens,
          // boutons, compteurs « (3) »).
          const title = legendOf(n) || gap.filter((t) => t.length <= 120 && !/^[(\[]?\d+[)\]]?$/.test(t)).pop() || "";
          sections.push({ title, files: grids.get(n) });
          gap = [];
        }
        continue;
      }
      const p = n.parentElement;
      if (!p || /^(SCRIPT|STYLE|NOSCRIPT|OPTION|BUTTON|TEXTAREA|SELECT)$/.test(p.tagName) || p.closest("a,button")) continue;
      const t = norm(n.nodeValue);
      if (!t || LOADING.test(t) || HEAD.test(fold(t)) || inGrid(n)) continue;
      gap.push(t);
      if (isTitle(t) || p.closest("legend")) titles.push(t);
    }
    // Sections sans tableau (vides) : titres « >> » / légendes non rattachés.
    const attached = new Set(sections.map((s) => s.title));
    titles.filter((t) => !attached.has(t)).forEach((t) => sections.push({ title: t, files: [], empty: true }));

    // Liens vers la page Documents (depuis la fiche lot).
    const gedLinks = [];
    const html = body.innerHTML || "";
    const re = /["'(]([^"'()\s]*ged_multi_file_manager[^"'()\s]*)["')]/g;
    let m;
    while ((m = re.exec(html)) && gedLinks.length < 5) {
      let u = m[1].replace(/&amp;/g, "&");
      try { u = new URL(u, url).href; } catch (_) { continue; }
      if (!gedLinks.includes(u)) gedLinks.push(u);
    }
    const isDocPage = /ged_multi_file_manager/i.test(url || "") || titles.length > 0 || gridTables.length > 0;
    return {
      ready: !loading && doc.readyState !== "loading",
      loading, login, url, isDocPage,
      hasForm: !!doc.forms && doc.forms.length > 0,
      titles, sections, gedLinks
    };
  }

  if (cfg.op === "parse") {
    const d = new DOMParser().parseFromString(cfg.html, "text/html");
    return scan(d, cfg.baseUrl || "");
  }

  if (cfg.op === "scan") return scan(document, location.href);

  return { ready: false, error: "opération inconnue : " + cfg.op };
}

/* ---------------------------------------------------------------- *
 * 3. Orchestration (côté panneau, via un adaptateur)               *
 * ---------------------------------------------------------------- *
 * ad = {
 *   navigate(url):            { ok, error? } (attend le chargement)
 *   execAll(cfg):             [{ frameId, result }] toutes frames
 *   exec(frameId, cfg):       résultat de dpePageOp dans une frame
 *   download(url, filename):  { ok, filename?, error? } (service worker →
 *                             chrome.downloads : cookies + redirections)
 *   sleep(ms), stopped(), log(text, cls)
 * }
 * p = { url, rules, timeoutMs, download, name(chosen) → chemin relatif }
 * Résultat : { status: OK|ABSENT|PAGE|AUTH|ERREUR, present, count, chosen,
 *              candidates, file, statut, error?, info[] } */

// Attend la page Documents (polling 250 ms) ; suit le lien
// ged_multi_file_manager si l'URL ouverte est la fiche lot — y compris
// quand la fiche lot (ex. syndic) contient déjà des tableaux mais aucune
// section recherchée (wants(page) faux).
async function dpeOpenDocs(ad, url, timeoutMs, info, wants) {
  const nav = await ad.navigate(url);
  if (!nav.ok) return { error: "ouverture impossible : " + (nav.error || url) };
  let followed = false;
  const t0 = Date.now();
  let last = null;
  for (;;) {
    if (ad.stopped()) return { error: "arrêté" };
    const all = await ad.execAll({ op: "scan" }).catch(() => []);
    const frames = all.filter((f) => f && f.result);
    const login = frames.some((f) => f.result.login && !f.result.isDocPage);
    const doc = frames.filter((f) => f.result.isDocPage)
      .sort((a, b) => (b.result.titles.length - a.result.titles.length))[0];
    last = doc || frames[0] || null;
    if (doc && doc.result.ready && (doc.result.titles.length || doc.result.sections.length)) {
      const link = frames.map((f) => f.result.gedLinks || []).flat()[0];
      if (!followed && wants && !wants(doc.result) && link && !/ged_multi_file_manager/i.test(doc.result.url || "")) {
        info.push("page Documents ouverte depuis la fiche lot : " + link);
        followed = true;
        const n2 = await ad.navigate(link);
        if (!n2.ok) return { error: "ouverture de la page Documents impossible : " + (n2.error || link) };
        continue;
      }
      return { frameId: doc.frameId, page: doc.result };
    }
    if (login) return { auth: true };
    if (!doc && !followed && Date.now() - t0 > 300) {
      const link = frames.map((f) => f.result.gedLinks || []).flat()[0];
      if (link && !frames.some((f) => f.result.loading)) {
        info.push("page Documents ouverte depuis la fiche lot : " + link);
        followed = true;
        const n2 = await ad.navigate(link);
        if (!n2.ok) return { error: "ouverture de la page Documents impossible : " + (n2.error || link) };
        continue;
      }
    }
    if (Date.now() - t0 > timeoutMs) {
      if (doc && doc.result.ready) return { frameId: doc.frameId, page: doc.result }; // page prête mais sans titre « >> »
      return { error: doc ? "« Chargement en cours... » toujours affiché (délai dépassé)" : "page Documents introuvable (" + ((last && last.result.url) || url) + ")", page: last && last.result };
    }
    await ad.sleep(250);
  }
}

// Gérance d'abord (p.url) ; si aucun DPE n'y est trouvé (ou la page est
// introuvable) et qu'une URL syndic est fournie (p.urlSyn), même recherche
// en syndic. res.source : « gérance », « syndic » ou « gérance + syndic ».
async function dpeProcess(ad, p) {
  const g = await dpeProcessOne(ad, p);
  g.source = "gérance";
  if (!p.urlSyn || ad.stopped() || !(g.status === "ABSENT" || g.status === "PAGE")) return g;
  const info = g.info.slice();
  info.push(`${g.status === "ABSENT" ? "aucun DPE en gérance" + (g.sections && g.sections.length ? " (" + g.sections.join(", ") + ")" : "") : "gérance : " + (g.error || g.statut)} → recherche en syndic`);
  const s = await dpeProcessOne(ad, Object.assign({}, p, { url: p.urlSyn }));
  s.source = "syndic";
  s.info = info.concat(s.info.map((x) => "syndic : " + x));
  if (!(s.status === "ABSENT" || s.status === "PAGE")) return s;
  // Ni en gérance ni en syndic.
  const both = (g.sections || []).map((t) => "gérance : " + t).concat((s.sections || []).map((t) => "syndic : " + t));
  if (g.status === "ABSENT" || s.status === "ABSENT") {
    const err = [g.status === "PAGE" ? "gérance : " + g.error : "", s.status === "PAGE" ? "syndic : " + s.error : ""].filter(Boolean).join(" ; ");
    return Object.assign(s.status === "ABSENT" ? s : g, { status: "ABSENT", present: "NON", statut: "Absent", count: 0, chosen: null, candidates: [],
      sections: both, source: "gérance + syndic", info: s.info, error: err || undefined });
  }
  return Object.assign(s, { source: "gérance + syndic", error: `gérance : ${g.error} ; syndic : ${s.error}` });
}

async function dpeProcessOne(ad, p) {
  const info = [];
  const rules = p.rules && p.rules.include ? p.rules : dpeCompileRules(p.rules);
  rules.errors.forEach((e) => info.push(e));
  const base = { present: "ERREUR", count: 0, chosen: null, candidates: [], file: "", info };
  if (!p.url) return Object.assign(base, { status: "PAGE", statut: "Page introuvable", error: "URL de la page Documents vide" });
  const timeout = Math.max(2000, p.timeoutMs || 20000);
  const open = await dpeOpenDocs(ad, p.url, timeout, info, (pg) => (pg.sections || []).some((s) => dpeIsSection(s.title, rules)));
  if (open.auth) return Object.assign(base, { status: "AUTH", statut: "Page introuvable", error: "session SIGEO expirée (page de connexion)" });
  if (open.error) return Object.assign(base, { status: "PAGE", statut: "Page introuvable", error: open.error });

  const sel = dpeSelect(open.page.sections, rules);
  if (!sel.sections.length) {
    const nf = open.page.sections.reduce((k, s) => k + s.files.length, 0);
    info.push("aucune section « EPC » trouvée (titres : " + (open.page.sections.map((s) => s.title || "?").join(" | ") || "aucun") + ` ; ${nf} fichier(s) lus)`);
  }
  const res = Object.assign(base, { candidates: sel.candidates, count: sel.candidates.length, chosen: sel.chosen, sections: sel.sections.map((s) => s.title) });
  if (!sel.chosen) return Object.assign(res, { status: "ABSENT", present: "NON", statut: "Absent" });
  res.present = "OUI";
  if (!p.download) return Object.assign(res, { status: "OK", statut: "Trouvé (non téléchargé)" });

  const url = sel.chosen.download || "";
  if (!/ged_file_download/i.test(url)) {
    return Object.assign(res, { status: "ERREUR", statut: "KO téléchargement : lien de téléchargement introuvable", error: "lien de téléchargement introuvable" });
  }
  const filename = p.name(sel.chosen);
  let dl;
  try { dl = await ad.download(url, filename); }
  catch (e) { dl = { ok: false, error: e.message || String(e) }; }
  if (!dl || !dl.ok) {
    const why = (dl && dl.error) || "aucune réponse du service worker";
    return Object.assign(res, { status: "ERREUR", statut: "KO téléchargement : " + why, error: why });
  }
  return Object.assign(res, { status: "OK", statut: "Téléchargé", file: dl.filename || filename.split("/").pop(), size: dl.size, url });
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    DPE_DEFAULT_RULES, DPE_DEFAULT_NAME, DPE_DEFAULT_FOLDER, DPE_DEFAULT_URL, DPE_DEFAULT_URL_SYN, dpeProcessOne,
    dpeFold, dpeCompileRules, dpeIsSection, dpeIsDpe, dpeParseDate, dpeSelect,
    dpeSanitize, dpeFileName, dpeFolder, dpePageOp, dpeOpenDocs, dpeProcess
  };
}
