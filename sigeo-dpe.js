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
 * télécharge par un POST __doPostBack rejoué en fetch depuis la page
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
  let s = String(tpl || DPE_DEFAULT_NAME).replace(/\{date\}/gi, vars.date || "sans-date");
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
 *   "download" { target, timeoutMs } → POST __doPostBack en fetch →
 *              { ok, b64, contentType, filename, size } | { ok:false, error }
 */
function dpePageOp(cfg) {
  const norm = (s) => String(s ?? "").replace(/[   ]/g, " ").replace(/\s+/g, " ").trim();
  const fold = (s) => norm(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const DATE = /\b\d{1,2}[\/.-]\d{1,2}[\/.-]\d{4}\b/;
  const HEAD = /^(nom|description|date|taille|type|auteur)$/;
  const LOADING = /chargement en cours/i;

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
      const links = Array.from(tr.querySelectorAll("a")).filter((a) => targetOf(a) && norm(a.textContent));
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
      const pb = targetOf(a);
      const table = tr.closest("table");
      if (!grids.has(table)) grids.set(table, []);
      grids.get(table).push({
        name,
        desc: descCell ? norm(descCell.textContent) : "",
        date: dateCell ? (norm(dateCell.textContent).match(DATE) || [""])[0] : "",
        target: pb.target, arg: pb.arg, order: order++
      });
    }

    // Titres : textes hors tableaux de fichiers, rattachés au tableau suivant.
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
          const title = gap.filter((t) => t.length <= 120 && !/^[(\[]?\d+[)\]]?$/.test(t)).pop() || "";
          sections.push({ title, files: grids.get(n) });
          gap = [];
        }
        continue;
      }
      const p = n.parentElement;
      if (!p || /^(SCRIPT|STYLE|NOSCRIPT|OPTION|BUTTON|TEXTAREA|SELECT)$/.test(p.tagName) || p.closest("a,button,[onclick]")) continue;
      const t = norm(n.nodeValue);
      if (!t || LOADING.test(t) || HEAD.test(fold(t)) || inGrid(n)) continue;
      gap.push(t);
      if (t.includes(">>")) titles.push(t);
    }
    // Sections sans tableau (vides) : titres « >> » non rattachés.
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

  if (cfg.op === "download") {
    return (async () => {
      const form = document.forms[0];
      if (!form) return { ok: false, error: "formulaire ASP.NET absent de la page" };
      const fd = new FormData(form);
      fd.set("__EVENTTARGET", cfg.target);
      fd.set("__EVENTARGUMENT", cfg.arg || "");
      const body = new URLSearchParams();
      for (const [k, v] of fd.entries()) if (typeof v === "string") body.append(k, v);
      let action = location.href;
      try { action = new URL(form.getAttribute("action") || location.href, location.href).href; } catch (_) {}
      const timeout = Math.max(3000, cfg.timeoutMs || 60000);
      const ctl = new AbortController();
      const t = setTimeout(() => ctl.abort(), timeout);
      let r, buf;
      try {
        r = await fetch(action, {
          method: "POST", credentials: "include", cache: "no-store", redirect: "follow", signal: ctl.signal,
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: body.toString()
        });
        buf = await r.arrayBuffer();
      } catch (e) {
        return { ok: false, error: e && e.name === "AbortError" ? `délai de ${Math.round(timeout / 1000)} s dépassé` : "fetch : " + ((e && e.message) || e) };
      } finally { clearTimeout(t); }
      const type = (r.headers.get("Content-Type") || "").toLowerCase();
      const disp = r.headers.get("Content-Disposition") || "";
      let filename = "";
      const fm = disp.match(/filename\*\s*=\s*(?:UTF-8|ISO-8859-1)?''([^;]+)/i) || disp.match(/filename\s*=\s*"([^"]+)"/i) || disp.match(/filename\s*=\s*([^;]+)/i);
      if (fm) { try { filename = decodeURIComponent(fm[1].trim()); } catch (_) { filename = fm[1].trim(); } }
      const bytes = new Uint8Array(buf);
      const head = String.fromCharCode.apply(null, Array.from(bytes.subarray(0, 5)));
      if (!r.ok) return { ok: false, error: `HTTP ${r.status}`, status: r.status };
      const declared = /attachment|filename/i.test(disp);
      const isHtml = /text\/html|xhtml/.test(type) && !declared && head !== "%PDF-";
      const isFile = !isHtml && (/pdf|octet-stream|force-download|^application\//.test(type) || declared || head === "%PDF-");
      if (!isFile) {
        const txt = new TextDecoder().decode(bytes.subarray(0, 4000));
        const why = /invalid postback|eventvalidation|viewstate/i.test(txt) ? " (validation ASP.NET refusée)"
          : /type="password"/i.test(txt) ? " (page de connexion : session expirée ?)" : "";
        const title = (txt.match(/<title[^>]*>([^<]*)</i) || [, ""])[1].trim();
        return { ok: false, html: true, error: `réponse HTML au lieu d'un fichier${why}${title ? ` — « ${title} »` : ""}`, contentType: type };
      }
      let bin = "";
      for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + 0x8000)));
      return { ok: true, b64: btoa(bin), size: bytes.length, contentType: type || "application/octet-stream", filename, pdf: head === "%PDF-" };
    })();
  }

  return { ready: false, error: "opération inconnue : " + cfg.op };
}

/* ---------------------------------------------------------------- *
 * 3. Orchestration (côté panneau, via un adaptateur)               *
 * ---------------------------------------------------------------- *
 * ad = {
 *   navigate(url):            { ok, error? } (attend le chargement)
 *   execAll(cfg):             [{ frameId, result }] toutes frames
 *   exec(frameId, cfg):       résultat de dpePageOp dans une frame
 *   save(b64, filename, type): { ok, filename?, error? } (chrome.downloads)
 *   sleep(ms), stopped(), log(text, cls)
 * }
 * p = { url, rules, timeoutMs, download, name(chosen, serverName) → nom }
 * Résultat : { status: OK|ABSENT|PAGE|AUTH|ERREUR, present, count, chosen,
 *              candidates, file, statut, error?, info[] } */

// Attend la page Documents (polling 250 ms) ; suit le lien
// ged_multi_file_manager si l'URL ouverte est la fiche lot.
async function dpeOpenDocs(ad, url, timeoutMs, info) {
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
    if (doc && doc.result.ready && (doc.result.titles.length || doc.result.sections.length)) return { frameId: doc.frameId, page: doc.result };
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

async function dpeProcess(ad, p) {
  const info = [];
  const rules = p.rules && p.rules.include ? p.rules : dpeCompileRules(p.rules);
  rules.errors.forEach((e) => info.push(e));
  const base = { present: "ERREUR", count: 0, chosen: null, candidates: [], file: "", info };
  if (!p.url) return Object.assign(base, { status: "PAGE", statut: "Page introuvable", error: "URL de la page Documents vide" });
  const timeout = Math.max(2000, p.timeoutMs || 20000);
  const open = await dpeOpenDocs(ad, p.url, timeout, info);
  if (open.auth) return Object.assign(base, { status: "AUTH", statut: "Page introuvable", error: "session SIGEO expirée (page de connexion)" });
  if (open.error) return Object.assign(base, { status: "PAGE", statut: "Page introuvable", error: open.error });

  const sel = dpeSelect(open.page.sections, rules);
  if (!sel.sections.length) info.push("aucune section « EPC » trouvée (titres : " + (open.page.titles.join(" | ") || "aucun") + ")");
  const res = Object.assign(base, { candidates: sel.candidates, count: sel.candidates.length, chosen: sel.chosen, sections: sel.sections.map((s) => s.title) });
  if (!sel.chosen) return Object.assign(res, { status: "ABSENT", present: "NON", statut: "Absent" });
  res.present = "OUI";
  if (!p.download) return Object.assign(res, { status: "OK", statut: "Trouvé (non téléchargé)" });

  let dl;
  try { dl = await ad.exec(open.frameId, { op: "download", target: sel.chosen.target, arg: sel.chosen.arg, timeoutMs: Math.max(timeout, 60000) }); }
  catch (e) { dl = { ok: false, error: "injection impossible : " + (e.message || e) }; }
  if (!dl || !dl.ok) {
    const why = (dl && dl.error) || "aucune réponse";
    if (dl && dl.html) ad.log("Téléchargement par POST __doPostBack refusé : " + why + ". (Aucune autre méthode tentée.)", "err");
    return Object.assign(res, { status: "ERREUR", statut: "KO téléchargement : " + why, error: why });
  }
  const filename = p.name(sel.chosen, dl.filename);
  const saved = await ad.save(dl.b64, filename, dl.contentType);
  if (!saved || !saved.ok) {
    const why = (saved && saved.error) || "enregistrement impossible";
    return Object.assign(res, { status: "ERREUR", statut: "KO téléchargement : " + why, error: why });
  }
  if (!dl.pdf) info.push("le fichier reçu n'est pas un PDF (" + (dl.contentType || "type inconnu") + ")");
  return Object.assign(res, { status: "OK", statut: "Téléchargé", file: saved.filename || filename, size: dl.size, b64: p.keepBytes ? dl.b64 : undefined });
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    DPE_DEFAULT_RULES, DPE_DEFAULT_NAME, DPE_DEFAULT_FOLDER, DPE_DEFAULT_URL,
    dpeFold, dpeCompileRules, dpeIsSection, dpeIsDpe, dpeParseDate, dpeSelect,
    dpeSanitize, dpeFileName, dpeFolder, dpePageOp, dpeOpenDocs, dpeProcess
  };
}
