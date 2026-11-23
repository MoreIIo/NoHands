/* ====================================================================
 * NoHands OSA — Recherche d'id interne de lot SIGEO (IM + LT → id)
 * --------------------------------------------------------------------
 * Onglet Extraction, type « Recherche d'id interne de lot ».
 *
 * Les fiches lot SIGEO ont l'URL
 *   …/page.aspx/fr/sig_imm/lot_manage/syn/{IM}/{ID_INTERNE}
 * où l'id interne n'a aucun rapport avec le numéro de lot (LT 204 → id 28).
 * On le retrouve via la recherche « Par code » de
 *   …/page.aspx/fr/sig_imm/lot_browse/syn
 * (formulaire ASP.NET WebForms, recherche par postback POST uniquement).
 *
 * Deux méthodes :
 *   1. « fetch » (préférée) : exécutée DANS l'onglet SIGEO (même origine →
 *      cookies de session envoyés). GET de lot_browse → champs cachés
 *      (__VIEWSTATE, __EVENTVALIDATION…) → POST du formulaire avec IM, LT,
 *      Niv. vide, Statut et le bouton Rechercher (name=value ou
 *      __EVENTTARGET selon le bouton) → HTML parsé avec DOMParser.
 *   2. « onglet » (repli) : remplit les champs de la page lot_browse, clique
 *      sur Rechercher, attend la fin du postback (rechargement complet ou
 *      UpdatePanel) et lit les liens.
 * En mode auto, si le POST direct n'est pas pris en compte (validation
 * d'événement, HTTP 500, champ LT non renvoyé…), on bascule sur l'onglet
 * pour le reste de l'exécution.
 *
 * Résultats : liens a[href*="lot_manage"] dont le texte ressemble à
 * « N° LT000204 Appartement Type 4 » ; on garde ceux dont l'IM (href) et le
 * numéro LT (texte) correspondent numériquement. 1 → OK, 0 → Introuvable,
 * plusieurs → Ambigu (jamais de choix arbitraire). Page de connexion →
 * AUTH (la boucle s'arrête).
 *
 * Chargé par sidepanel.html AVANT sidepanel.js ; importable sous Node
 * (tests). lotPageOp est injectée telle quelle dans la page : elle doit
 * rester autonome (aucune référence au reste du module).
 * ==================================================================== */

const LOT_DEFAULT_URL = "http://sigeo.evoriel.net/page.aspx/fr/sig_imm/lot_browse/syn";

/* ---------------------------------------------------------------- *
 * 1. Normalisation des entrées (logique pure)                      *
 * ---------------------------------------------------------------- */

const lotClean = (v) => String(v ?? "").replace(/[\s   ]+/g, "").toUpperCase();

// « IM0393528 », « 0393528 », « 393528 », « 393528.0 » → { num:"393528", pad:"0393528" }
function lotNormIm(v) {
  const s = lotClean(v).replace(/^IM/, "").replace(/[.,]0+$/, "");
  if (!/^\d+$/.test(s)) return null;
  const num = String(parseInt(s, 10));
  return { num, pad: num.padStart(7, "0") };
}

// « LT000204 », « LT 204 », « 000204 », « 204 », « 204.0 » → "204"
function lotNormLt(v) {
  const s = lotClean(v).replace(/^LT/, "").replace(/[.,]0+$/, "");
  if (!/^\d+$/.test(s)) return null;
  return String(parseInt(s, 10));
}

// Texte de statut écrit dans la colonne « Résultat ».
function lotStatusText(res) {
  if (!res) return "";
  if (res.status === "OK") return `OK — id ${res.matches[0].id}`;
  if (res.status === "INTROUVABLE") return "Introuvable";
  if (res.status === "AMBIGU") {
    return "Ambigu — " + res.matches.map((m) => `id ${m.id}${m.label ? ` (${m.label})` : ""}`).join(", ");
  }
  if (res.status === "AUTH") return "ERREUR — session SIGEO expirée (page de connexion)";
  return "ERREUR — " + (res.error || "inconnue");
}

/* ---------------------------------------------------------------- *
 * 2. Fonction injectée dans la page SIGEO (autonome)               *
 * ---------------------------------------------------------------- *
 * cfg.op :
 *   "fetch"  { url, im, imPad, lt, statut, sel, timeoutMs } → résultat (async)
 *   "probe"  { sel }                → champs / bouton trouvés sur la page
 *   "submit" { token, imPad, lt, statut, sel } → remplit + clique Rechercher
 *   "read"   { token, im, lt, sel } → { ready, …résultat }
 *   "parse"  { html, baseUrl, im, lt } → résultat (tests)
 * Résultat : { status: OK|INTROUVABLE|AMBIGU|AUTH|ERREUR, matches:[{im,id,lt,label,url}],
 *              total, others, info[], error?, fallback? }
 */
function lotPageOp(cfg) {
  const SEL = Object.assign({
    lt: "#body_x_txtLotId_x_txtLotId",
    niv: "#body_x_txtBatNiv_x_txtBatNiv",
    im: "#body_x_txtImmId_x_txtImmId",
    statut: "#body_x_ddlLotStatut_ddlLotStatut",
    submit: ""
  }, Object.fromEntries(Object.entries(cfg.sel || {}).filter(([, v]) => v)));
  const SUFFIX = { lt: 'input[id$="txtLotId"]', niv: 'input[id$="txtBatNiv"]', im: 'input[id$="txtImmId"]', statut: 'select[id$="ddlLotStatut"]' };

  const norm = (s) => String(s ?? "").replace(/[   ]/g, " ").replace(/\s+/g, " ").trim();
  const num = (s) => { const m = String(s ?? "").match(/\d+/); return m ? String(parseInt(m[0], 10)) : ""; };
  const q = (root, sel) => { if (!sel) return null; try { return root.querySelector(sel); } catch (_) { return null; } };
  const field = (root, k) => q(root, SEL[k]) || (SUFFIX[k] ? root.querySelector(SUFFIX[k]) : null);
  const isLogin = (doc, url) => !field(doc, "lt") &&
    (!!doc.querySelector('input[type="password"]') || /login|logon|connexion|authentif/i.test(url || ""));
  const describe = (el) => {
    if (!el) return "";
    const t = el.tagName.toLowerCase();
    return t + (el.id ? "#" + el.id : el.getAttribute("name") ? `[name="${el.getAttribute("name")}"]` : "") +
      ` « ${norm(t === "input" ? el.getAttribute("value") || el.getAttribute("alt") : el.textContent)} »`;
  };

  // Bouton « Rechercher » : sélecteur forcé, sinon libellé exact, sinon
  // « recherch » dans libellé/id/name ; si plusieurs, le plus proche du champ LT.
  function findSubmit(root, ltEl) {
    const forced = q(root, SEL.submit);
    if (forced) return forced;
    const all = Array.from(root.querySelectorAll('input[type="submit"],input[type="button"],input[type="image"],button,a'));
    const label = (el) => norm(el.tagName === "INPUT" ? (el.getAttribute("value") || "") : el.textContent);
    const extra = (el) => [el.getAttribute("alt"), el.getAttribute("title"), el.id, el.getAttribute("name")].join(" ");
    let cands = all.filter((el) => /^rechercher$/i.test(label(el)) || /^rechercher$/i.test(norm(el.getAttribute("alt") || el.getAttribute("title"))));
    if (!cands.length) cands = all.filter((el) => /recherch|search/i.test(label(el) + " " + extra(el)));
    if (cands.length <= 1 || !ltEl) return cands[0] || null;
    let best = null, bestD = Infinity;
    for (const c of cands) {
      let a = ltEl, d = 0;
      while (a && !a.contains(c)) { a = a.parentElement; d++; }
      if (a && d < bestD) { best = c; bestD = d; }
    }
    return best || cands[0];
  }

  // Ce que le clic sur le bouton envoie au serveur.
  function postbackOf(btn) {
    if (!btn) return null;
    const code = (btn.getAttribute("onclick") || "") + " " + (btn.getAttribute("href") || "");
    const tag = btn.tagName, type = (btn.getAttribute("type") || (tag === "BUTTON" ? "submit" : "")).toLowerCase();
    const name = btn.getAttribute("name");
    if ((tag === "INPUT" && (type === "submit" || type === "image")) || (tag === "BUTTON" && type === "submit")) {
      if (name) return { submitName: name, submitValue: btn.getAttribute("value") || "", image: type === "image" };
      return { plain: true };
    }
    let m = code.match(/__doPostBack\(\s*['"]([^'"]*)['"]\s*,\s*['"]([^'"]*)['"]/);
    if (!m) m = code.match(/WebForm_PostBackOptions\(\s*["']([^"']*)["']\s*,\s*["']([^"']*)["']/);
    if (m) return { target: m[1], arg: m[2] };
    if (name) return { target: name, arg: "" };
    return null;
  }

  // Valeur de la liste Statut : par valeur puis par libellé ; "" = défaut de la page.
  function setStatut(sel, wanted, info) {
    if (!sel || !wanted) return;
    const w = norm(wanted).toLowerCase();
    const opt = Array.from(sel.options).find((o) => norm(o.value).toLowerCase() === w) ||
      Array.from(sel.options).find((o) => norm(o.textContent).toLowerCase() === w);
    if (opt) { sel.value = opt.value; return; }
    info.push(`statut « ${wanted} » absent de la liste (valeur de la page conservée)`);
  }

  function fillFields(root, c, info, events) {
    const set = (el, v) => {
      if (!el) return;
      el.value = v;
      if (events) { el.dispatchEvent(new Event("input", { bubbles: true })); el.dispatchEvent(new Event("change", { bubbles: true })); }
    };
    set(field(root, "lt"), c.lt);
    set(field(root, "niv"), "");
    set(field(root, "im"), c.imPad);
    const st = field(root, "statut");
    if (st && c.statut) {
      const before = st.value;
      setStatut(st, c.statut, info);
      if (events && st.value !== before) st.dispatchEvent(new Event("change", { bubbles: true }));
    }
  }

  // Sérialisation façon navigateur (champs réussis du formulaire).
  function serialize(form, pb) {
    const out = [];
    const els = form.elements ? Array.from(form.elements) : Array.from(form.querySelectorAll("input,select,textarea"));
    for (const el of els) {
      const name = el.getAttribute("name");
      if (!name || el.disabled) continue;
      const tag = el.tagName, type = (el.getAttribute("type") || "text").toLowerCase();
      if (tag === "INPUT") {
        if (["submit", "button", "image", "reset", "file"].includes(type)) continue;
        if (type === "checkbox" || type === "radio") { if (el.checked) out.push([name, el.getAttribute("value") ?? "on"]); continue; }
        out.push([name, el.value]);
      } else if (tag === "SELECT") {
        Array.from(el.options).filter((o) => o.selected).forEach((o) => out.push([name, o.value]));
      } else if (tag === "TEXTAREA") {
        out.push([name, el.value]);
      }
    }
    const rest = out.filter(([n]) => n !== "__EVENTTARGET" && n !== "__EVENTARGUMENT");
    rest.unshift(["__EVENTTARGET", pb.target || ""], ["__EVENTARGUMENT", pb.arg || ""]);
    if (pb.submitName) {
      if (pb.image) rest.push([pb.submitName + ".x", "1"], [pb.submitName + ".y", "1"]);
      else rest.push([pb.submitName, pb.submitValue]);
    }
    return rest;
  }

  // Liens de résultats → lots ; filtre IM + LT exacts (comparaison numérique).
  function parseResults(doc, baseUrl, im, lt) {
    const byKey = new Map();
    for (const a of Array.from(doc.querySelectorAll('a[href*="lot_manage"]'))) {
      const href = a.getAttribute("href") || "";
      const m = href.match(/lot_manage\/syn\/(\d+)\/(\d+)/);
      if (!m) continue;
      const label = norm(a.textContent) || norm(a.getAttribute("title"));
      const lm = label.match(/(?:^|[^A-Z])LT\s*0*(\d+)/i);
      let url = href;
      try { url = new URL(href, baseUrl).href; } catch (_) {}
      const key = m[1] + "/" + m[2];
      const prev = byKey.get(key);
      const item = { im: String(parseInt(m[1], 10)), id: m[2], lt: lm ? String(parseInt(lm[1], 10)) : "", label, url };
      if (!prev) byKey.set(key, item);
      else if (!prev.lt && item.lt) byKey.set(key, Object.assign(item, { label: item.label || prev.label }));
    }
    const all = Array.from(byKey.values());
    const imN = String(parseInt(im, 10)), ltN = String(parseInt(lt, 10));
    const info = [];
    let matches = all.filter((x) => x.im === imN && x.lt === ltN);
    if (!matches.length && all.length && all.every((x) => !x.lt)) {
      // Aucun numéro LT lisible dans les libellés : on ne peut que filtrer sur l'IM.
      matches = all.filter((x) => x.im === imN);
      if (matches.length) info.push("numéro LT illisible dans les résultats : correspondance sur l'IM seul");
    }
    const others = all.filter((x) => !matches.includes(x)).slice(0, 5);
    return {
      status: matches.length === 1 ? "OK" : matches.length > 1 ? "AMBIGU" : "INTROUVABLE",
      matches, total: all.length, others, info
    };
  }

  if (cfg.op === "parse") {
    const d = new DOMParser().parseFromString(cfg.html, "text/html");
    if (isLogin(d, cfg.baseUrl)) return { status: "AUTH", matches: [], info: [] };
    return parseResults(d, cfg.baseUrl, cfg.im, cfg.lt);
  }

  if (cfg.op === "probe") {
    const lt = field(document, "lt");
    const btn = findSubmit(document, lt);
    return {
      url: location.href,
      login: isLogin(document, location.href),
      fields: { lt: !!lt, im: !!field(document, "im"), niv: !!field(document, "niv"), statut: !!field(document, "statut") },
      submit: describe(btn),
      postback: postbackOf(btn),
      links: document.querySelectorAll('a[href*="lot_manage"]').length
    };
  }

  if (cfg.op === "submit") {
    const lt = field(document, "lt");
    if (!lt) return { ok: false, login: isLogin(document, location.href), error: "champ LT absent de la page" };
    const btn = findSubmit(document, lt);
    if (!btn) return { ok: false, error: "bouton Rechercher introuvable" };
    const info = [];
    fillFields(document, cfg, info, true);
    window.__nhLotPending = cfg.token;
    window.__nhLotDone = "";
    try {
      const prm = window.Sys && Sys.WebForms && Sys.WebForms.PageRequestManager.getInstance();
      if (prm && !window.__nhLotHooked) {
        prm.add_endRequest(() => { window.__nhLotDone = window.__nhLotPending; });
        window.__nhLotHooked = true;
      }
    } catch (_) {}
    setTimeout(() => btn.click(), 0); // après le retour de executeScript
    return { ok: true, button: describe(btn), info };
  }

  if (cfg.op === "read") {
    if (document.readyState !== "complete") return { ready: false };
    try { if (Sys.WebForms.PageRequestManager.getInstance().get_isInAsyncPostBack()) return { ready: false }; } catch (_) {}
    const reloaded = window.__nhLotPending !== cfg.token; // nouveau document après postback complet
    if (!reloaded && window.__nhLotDone !== cfg.token) return { ready: false };
    const lt = field(document, "lt");
    if (!lt) {
      return isLogin(document, location.href)
        ? { ready: true, status: "AUTH", matches: [] }
        : { ready: true, status: "ERREUR", matches: [], error: "page inattendue après la recherche (" + location.href + ")" };
    }
    if (num(lt.value) !== String(parseInt(cfg.lt, 10))) {
      return { ready: true, status: "ERREUR", matches: [], error: "la page affichée ne correspond pas à la recherche (LT " + (lt.value || "vide") + ")" };
    }
    return Object.assign({ ready: true }, parseResults(document, location.href, cfg.im, cfg.lt));
  }

  if (cfg.op === "fetch") {
    return (async () => {
      const timeout = Math.max(2000, cfg.timeoutMs || 20000);
      const go = async (url, opts) => {
        const ctl = new AbortController();
        const t = setTimeout(() => ctl.abort(), timeout);
        try {
          const r = await fetch(url, Object.assign({ credentials: "include", cache: "no-store", redirect: "follow", signal: ctl.signal }, opts));
          return { r, html: await r.text() };
        } finally { clearTimeout(t); }
      };
      const errOf = (e) => (e && e.name === "AbortError") ? `délai de ${Math.round(timeout / 1000)} s dépassé` : (e && e.message) || String(e);
      const info = [];

      let g;
      try { g = await go(cfg.url, {}); }
      catch (e) { return { status: "ERREUR", matches: [], error: "GET lot_browse : " + errOf(e), fallback: true }; }
      const d1 = new DOMParser().parseFromString(g.html, "text/html");
      if (isLogin(d1, g.r.url)) return { status: "AUTH", matches: [] };
      const lt1 = field(d1, "lt");
      if (!lt1) return { status: "ERREUR", matches: [], error: `formulaire de recherche absent de la page (HTTP ${g.r.status})`, fallback: true };
      const form = lt1.form || lt1.closest("form") || d1.querySelector("form");
      if (!form) return { status: "ERREUR", matches: [], error: "balise <form> introuvable", fallback: true };
      const btn = findSubmit(d1, lt1);
      const pb = postbackOf(btn);
      if (!pb) return { status: "ERREUR", matches: [], error: "bouton Rechercher introuvable", fallback: true };
      fillFields(d1, cfg, info, false);
      const pairs = serialize(form, pb);
      let action = g.r.url || cfg.url;
      try { action = new URL(form.getAttribute("action") || action, g.r.url || cfg.url).href; } catch (_) {}

      let p;
      try {
        p = await go(action, {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams(pairs).toString()
        });
      } catch (e) { return { status: "ERREUR", matches: [], error: "POST : " + errOf(e), fallback: true }; }
      const d2 = new DOMParser().parseFromString(p.html, "text/html");
      if (isLogin(d2, p.r.url)) return { status: "AUTH", matches: [] };
      if (!p.r.ok) {
        const why = /invalid postback|eventvalidation|viewstate/i.test(p.html) ? " (validation ASP.NET refusée)" : "";
        return { status: "ERREUR", matches: [], error: `HTTP ${p.r.status} au postback${why}`, fallback: true };
      }
      const res = parseResults(d2, p.r.url || action, cfg.im, cfg.lt);
      res.info = info.concat(res.info);
      res.button = describe(btn);
      const lt2 = field(d2, "lt");
      const echoed = !!lt2 && num(lt2.getAttribute("value")) === String(parseInt(cfg.lt, 10));
      if (!res.matches.length && !echoed) {
        return Object.assign(res, { status: "ERREUR", error: "le postback n'a pas été pris en compte (champ LT non renvoyé par le serveur)", fallback: true });
      }
      return res;
    })();
  }

  return { status: "ERREUR", matches: [], error: "opération inconnue : " + cfg.op };
}

/* ---------------------------------------------------------------- *
 * 3. Orchestration (côté panneau, via un adaptateur)               *
 * ---------------------------------------------------------------- *
 * ad = {
 *   tabUrl():                 URL actuelle de l'onglet cible
 *   navigate(url):            { ok, error? } (attend le chargement)
 *   exec(frameId, cfg, world) résultat de lotPageOp dans une frame
 *   execAll(cfg, world):      [{ frameId, result }] toutes frames
 *   sleep(ms), stopped(), log(text, cls)
 * }
 * run = { fetchBroken } : état partagé sur une exécution (repli définitif). */

const lotCache = new Map();
function lotClearCache() { lotCache.clear(); }

let lotTokenSeq = 0;
const lotNewToken = () => "nhlot" + Date.now().toString(36) + (++lotTokenSeq);

const lotOrigin = (u) => { try { return new URL(u).origin; } catch (_) { return ""; } };

// Cherche une ligne : p = { im, lt (bruts), url, statut, sel, timeoutMs, method: auto|fetch|ui }
async function lotLookup(ad, p, run = {}) {
  const im = lotNormIm(p.im), lt = lotNormLt(p.lt);
  if (!im) return { status: "ERREUR", matches: [], error: `IM invalide : « ${p.im} »` };
  if (!lt) return { status: "ERREUR", matches: [], error: `LT invalide : « ${p.lt} »` };
  const url = p.url || LOT_DEFAULT_URL;
  const key = [im.num, lt, p.statut || ""].join("|");
  if (lotCache.has(key)) return Object.assign({}, lotCache.get(key), { cached: true });
  const cfg = { im: im.num, imPad: im.pad, lt, statut: p.statut || "", sel: p.sel || {}, url, timeoutMs: p.timeoutMs || 20000 };

  let res = null;
  if (p.method !== "ui" && !run.fetchBroken) {
    // fetch depuis la page SIGEO elle-même : même origine → cookies de session.
    if (lotOrigin(await ad.tabUrl()) !== lotOrigin(url)) {
      const nav = await ad.navigate(url);
      if (!nav.ok) return { status: "ERREUR", matches: [], error: "ouverture de SIGEO impossible : " + (nav.error || "") };
    }
    try { res = await ad.exec(0, Object.assign({ op: "fetch" }, cfg), "ISOLATED"); }
    catch (e) { res = { status: "ERREUR", matches: [], error: "injection impossible : " + (e.message || e), fallback: true }; }
    res = res || { status: "ERREUR", matches: [], error: "aucune réponse de la page", fallback: true };
    res.via = "fetch";
    if (res.status === "ERREUR" && res.fallback && p.method !== "fetch") {
      ad.log(`POST direct impossible (${res.error}) → repli : recherche dans l'onglet pour la suite.`, "skip");
      run.fetchBroken = true;
      res = null;
    }
  }
  if (!res) res = await lotLookupUi(ad, cfg);
  if (res.status !== "ERREUR" && res.status !== "AUTH") lotCache.set(key, res);
  return res;
}

// Repli : pilote le formulaire de l'onglet.
async function lotLookupUi(ad, cfg) {
  const timeout = Math.max(2000, cfg.timeoutMs);
  const locate = async () => {
    const all = await ad.execAll({ op: "probe", sel: cfg.sel }, "MAIN").catch(() => []);
    const hit = all.find((f) => f.result && f.result.fields && f.result.fields.lt);
    const login = all.some((f) => f.result && f.result.login);
    return { frameId: hit ? hit.frameId : null, probe: hit && hit.result, login };
  };
  let loc = await locate();
  if (loc.frameId === null) { // pas sur lot_browse (ou page de connexion) : on l'ouvre, puis on juge
    const nav = await ad.navigate(cfg.url);
    if (!nav.ok) return { status: "ERREUR", matches: [], error: "ouverture de lot_browse impossible : " + (nav.error || ""), via: "onglet" };
    loc = await locate();
  }
  if (loc.frameId === null) {
    return loc.login
      ? { status: "AUTH", matches: [], via: "onglet" }
      : { status: "ERREUR", matches: [], error: "formulaire de recherche de lots introuvable dans l'onglet", via: "onglet" };
  }
  const token = lotNewToken();
  const sub = await ad.exec(loc.frameId, Object.assign({ op: "submit", token }, cfg), "MAIN");
  if (!sub || !sub.ok) {
    return sub && sub.login ? { status: "AUTH", matches: [], via: "onglet" }
      : { status: "ERREUR", matches: [], error: (sub && sub.error) || "remplissage impossible", via: "onglet" };
  }
  const t0 = Date.now();
  await ad.sleep(300);
  while (Date.now() - t0 < timeout) {
    if (ad.stopped()) return { status: "ERREUR", matches: [], error: "arrêté", via: "onglet" };
    let r = null;
    try { r = await ad.exec(loc.frameId, { op: "read", token, im: cfg.im, lt: cfg.lt, sel: cfg.sel }, "MAIN"); }
    catch (_) { /* page en cours de rechargement */ }
    if (r && r.ready) {
      delete r.ready;
      r.info = (sub.info || []).concat(r.info || []);
      r.via = "onglet";
      return r;
    }
    await ad.sleep(250);
  }
  return { status: "ERREUR", matches: [], error: `pas de réponse de la recherche (délai ${Math.round(timeout / 1000)} s dépassé)`, via: "onglet" };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    LOT_DEFAULT_URL, lotNormIm, lotNormLt, lotStatusText, lotPageOp,
    lotLookup, lotLookupUi, lotClearCache
  };
}
