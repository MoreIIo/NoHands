// Test d'intégration du panneau : onglet Extraction, type « DPE Intralia ».
// Le vrai sidepanel.html est servi en HTTP local (worker pdf.js pour
// « vérifier le contenu ») ; chrome.* est simulé : l'« onglet Intralia » est
// un onglet Chrome ouvert sur le faux Intralia (tests/fixtures/intralia-mock.js,
// pages windows-1252), chrome.downloads enregistre en mémoire, FlowerDocs est
// un second faux serveur (autre origine, CORS avec cookies simulé).
//   NODE_PATH=$(npm root -g) node tests/intralia-dpe.panel.e2e.js
const fs = require("fs");
const http = require("http");
const path = require("path");
const assert = require("assert");
const { chromium } = require("playwright");
const mock = require("./fixtures/intralia-mock.js");
const lotMock = require("./fixtures/sigeo-lot-server.js");

const ROOT = path.resolve(__dirname, "..");
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".svg": "image/svg+xml", ".png": "image/png", ".gz": "application/gzip", ".wasm": "application/wasm" };
function serve() {
  return new Promise((resolve) => {
    const srv = http.createServer((req, res) => {
      const p = path.join(ROOT, decodeURIComponent(req.url.split("?")[0]));
      if (!p.startsWith(ROOT) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); res.end(); return; }
      res.writeHead(200, { "content-type": TYPES[path.extname(p)] || "application/octet-stream" });
      fs.createReadStream(p).pipe(res);
    });
    srv.listen(0, "127.0.0.1", () => resolve({ srv, origin: `http://127.0.0.1:${srv.address().port}` }));
  });
}

const CHROME_STUB = `
(() => {
  const store = {};
  window.__store = store;
  const updated = new Set();
  const dl = new Map(); const dlL = new Set(); let dlSeq = 0;
  window.chrome = {
    storage: { local: {
      get: (k, cb) => {
        const keys = k == null ? Object.keys(store) : Array.isArray(k) ? k : typeof k === "string" ? [k] : Object.keys(k);
        const r = {}; keys.forEach((x) => { if (x in store) r[x] = JSON.parse(JSON.stringify(store[x])); });
        if (typeof cb === "function") cb(r); return Promise.resolve(r);
      },
      set: (o, cb) => { Object.assign(store, JSON.parse(JSON.stringify(o))); if (typeof cb === "function") cb(); return Promise.resolve(); },
      remove: (k, cb) => { [].concat(k).forEach((x) => delete store[x]); if (typeof cb === "function") cb(); return Promise.resolve(); }
    } },
    runtime: { getURL: (p) => new URL(p, location.href).href, getManifest: () => ({ version: "0.0.0" }),
               sendMessage: async () => ({}), lastError: null, onMessage: { addListener() {} } },
    // Onglet 1 : Intralia ; onglet 2 : SIGEO (recherche d'id interne de lot).
    tabs: {
      get: async (id) => { if (id !== 1 && id !== 2) throw new Error("No tab with id: " + id); return { id, url: await window.__tabUrl(id), status: "complete" }; },
      query: async () => [{ id: 1, url: await window.__tabUrl(1), active: true }, { id: 2, url: await window.__tabUrl(2), active: false }],
      create: async () => { throw new Error("création d'onglet non prévue dans ce test"); },
      sendMessage: async () => ({}), duplicate: async () => ({}),
      update: (id, o, cb) => { window.__navigate(id, o.url).then(() => { if (cb) cb({ id }); updated.forEach((l) => l(id, { status: "complete" })); }); },
      onUpdated: { addListener(l) { updated.add(l); }, removeListener(l) { updated.delete(l); } }, onRemoved: { addListener() {} }
    },
    scripting: { executeScript: async ({ target, func, args }) => {
      if (target.tabId !== 1 && target.tabId !== 2) throw new Error("onglet inconnu");
      return [{ frameId: 0, result: await window.__runInTab(target.tabId, func.toString(), args || []) }];
    } },
    downloads: {
      download: async ({ url, filename, conflictAction, saveAs }) => {
        const u8 = new Uint8Array(await (await fetch(url)).arrayBuffer());
        let s = ""; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
        const full = await window.__saveDl({ url: url.slice(0, 5), filename, conflictAction, saveAs, b64: btoa(s) });
        const id = ++dlSeq;
        dl.set(id, { id, state: "in_progress", filename: full, mime: "application/pdf" });
        setTimeout(() => { dl.get(id).state = "complete"; dlL.forEach((l) => l({ id, state: { current: "complete" } })); }, 20);
        return id;
      },
      search: async ({ id }) => (dl.has(id) ? [Object.assign({}, dl.get(id))] : []),
      onChanged: { addListener: (l) => dlL.add(l), removeListener: (l) => dlL.delete(l) },
      onCreated: { addListener() {}, removeListener() {} }
    },
    permissions: { contains: async () => true, request: async () => true },
    windows: { update: async () => ({}) }
  };
})();`;

const HEAD = ["Immeuble", "IM", "lots", "IM_format"];
const ROWS = [
  HEAD,
  ["PARC FLEURI LE CONCERTO 1", "IM0393001", "0732", "393001"],   // 2  OK
  ["Cour Saint-Hubert", "IM0393002", "130", "393002"],            // 3  OK (…_DP.pdf)
  ["RESIDENCE BEAUSEJOUR", "IM0393003", "543", "393003"],         // 4  OK (accents côté Intralia)
  ["RÉSIDENCE CHATEAUBRIAND", "IM0393004", "278", "393004"],      // 5  OK (doublon dans la GED)
  ["OPERA", "IM0393005", "9", "393005"],                          // 6  AUCUN_DPE_GED (0 doc)
  ["PARC DES LILAS", "IM0393006", "211", "393006"],               // 7  AUCUN_DPE_GED (2 docs)
  ["LES ACACIAS", "IM0393007", "3", "393007"],                    // 8  AMBIGU
  ["PARC FLEURI LE CONCERTO 1", "IM0393001", "999", "393001"],    // 9  LOT_INTROUVABLE
  ["PARC FLEURI CONCERTO", "IM0393008", "732", "393008"],         // 10 IMMEUBLE_INTROUVABLE
  ["LE ZOLA", "IM0393009", "7", "393009"],                        // 11 IMMEUBLE_INTROUVABLE (autre cabinet)
  ["Parc Fleuri - le Concerto 1", "IM0393001", "732", "393001"],  // 12 OK (déjà traité, ligne 2)
  ["LES TILLEULS", "IM0393010", "12", "393010"],                  // 13 OK (le plus récent n'est pas un DPE)
  ["", "IM0393011", "5", "393011"]                                // 14 ignorée (nom vide)
];

(async () => {
  const { srv: web, origin } = await serve();
  const srv = await mock.start();
  // Faux SIGEO « lot_browse » : lots des lignes du test (IM_format → LT → id interne).
  Object.assign(lotMock.DB, {
    393001: [{ lt: 732, id: 38, type: "Appartement" }],
    393002: [{ lt: 130, id: 12, type: "Appartement" }],
    393003: [{ lt: 543, id: 5, type: "Appartement" }],
    393010: [{ lt: 12, id: 3, type: "Appartement" }, { lt: 120, id: 4, type: "Cave" }]
  }); // 393004 (ligne 5) absent → introuvable
  const lsrv = await lotMock.start();
  const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/usr/bin/google-chrome", headless: true });
  const ctx = await browser.newContext();
  const tab = await ctx.newPage();
  await tab.goto(srv.intraOrigin + "/");
  const stab = await ctx.newPage();
  await stab.goto(lsrv.origin + "/__login");
  await stab.goto(lsrv.url);
  const tabs = { 1: tab, 2: stab };
  const saved = [];
  const panel = await ctx.newPage();
  const errors = [];
  panel.on("pageerror", (e) => { errors.push(e.message); console.log("[panel] " + e.message); });
  await panel.exposeFunction("__tabUrl", (id) => tabs[id].url());
  await panel.exposeFunction("__navigate", async (id, u) => { await tabs[id].goto(u); });
  await panel.exposeFunction("__runInTab", async (id, src, args) => {
    for (let i = 0; ; i++) {
      try { return await tabs[id].evaluate(([s, a]) => (new Function("return (" + s + ")"))()(...a), [src, args]); }
      catch (e) { if (i >= 5 || !/context was destroyed|navigat/i.test(e.message)) throw e; await new Promise((r) => setTimeout(r, 200)); }
    }
  });
  await panel.exposeFunction("__saveDl", (o) => {
    const buf = Buffer.from(o.b64, "base64");
    let name = o.filename, k = 0;
    while (saved.some((x) => x.name === name)) name = o.filename.replace(/(\.[^.]+)$/, ` (${++k})$1`);
    saved.push({ name, requested: o.filename, head: buf.subarray(0, 5).toString("latin1"), size: buf.length, url: o.url, conflictAction: o.conflictAction, saveAs: o.saveAs });
    return "C:\\Users\\m\\Downloads\\" + name.replace(/\//g, "\\");
  });
  await panel.addInitScript(CHROME_STUB);
  await panel.goto(origin + "/sidepanel.html");
  await panel.waitForFunction(() => typeof runIntraExtraction === "function" && typeof intraPageOp === "function");

  await panel.evaluate((rows) => { state.headerMode = "yes"; state.rows = rows; renderColumns(); showTab("extraction"); }, ROWS);
  await panel.check('input[name="extractKind"][value="intraliaDpe"]');
  const ui = await panel.evaluate(() => ({
    box: !$("intraExtractBox").hidden, std: !$("stdExtractBox").hidden, dpe: !$("dpeExtractBox").hidden,
    cfg: getIntraConfig(), socs: Array.from($("intraSoc").options).map((o) => o.textContent),
    warn: $("intraSocWarn").hidden, info: $("intraIdxInfo").textContent, nFields: document.querySelectorAll(".intra-field").length
  }));
  console.log("UI :", JSON.stringify(ui));
  assert.ok(ui.box && !ui.std && !ui.dpe, "seul le bloc Intralia est visible");
  assert.deepStrictEqual([ui.cfg.nomCol, ui.cfg.lotCol], ["Immeuble", "lots"], "colonnes devinées (jamais IM_format)");
  assert.deepStrictEqual([ui.cfg.soc, ui.cfg.mode, ui.cfg.folder, ui.cfg.name, ui.cfg.delay, ui.cfg.verify, ui.cfg.analyze, ui.cfg.ocr],
    ["19", "download", "DPE", "{Immeuble}_lot{lots}_{DPE_Date_GED}.pdf", 400, false, true, "auto"]);
  assert.strictEqual(await panel.evaluate(() => document.querySelector('input[name="extractKind"][value="intraliaDpe"]').parentElement.textContent.trim()),
    "DPE extraction (immeuble + lot → fichier prêt pour la saisie SIGEO)");
  await panel.click("#intraOpts summary");
  assert.ok(/^\d+ fournisseurs$/.test(await panel.textContent("#intraSupCount")), "liste fournisseurs de l'Analyse DPE");
  assert.deepStrictEqual(ui.cfg.fields, ["libdpe1", "DPE1", "libges1", "GES1", "DATEDIAG"]);
  assert.deepStrictEqual(ui.socs, ["Sicov (19)", "Tous les cabinets (100)"]);
  assert.ok(ui.warn, "avertissement « tous les cabinets » masqué");
  assert.strictEqual(ui.nFields, 12);
  assert.ok(/pas encore chargé/.test(ui.info), ui.info);
  await panel.selectOption("#intraSoc", "100");
  assert.strictEqual(await panel.evaluate(() => $("intraSocWarn").hidden), false, "avertissement affiché pour « Tous les cabinets »");
  await panel.selectOption("#intraSoc", "19");
  await panel.click("#intraAdv summary");
  await panel.fill("#intraBase", srv.intraOrigin);
  await panel.fill("#intraFlower", srv.flowerOrigin);
  await panel.fill("#intraDelay", "0");
  await panel.evaluate((u) => { $("lotUrl").value = u; }, lsrv.url); // réglage de « Recherche d'id interne de lot »
  assert.deepStrictEqual(await panel.evaluate(() => [getIntraConfig().imCol, getIntraConfig().idCol]), ["IM_format", ""], "IM SIGEO deviné (IM_format)");
  await panel.evaluate(() => updateExtractEstimate());
  const est = await panel.textContent("#extractEstimate");
  assert.ok(/12 lots \(11 différents\)/.test(est), est);

  // Tester sur la ligne de départ : rien n'est écrit ni téléchargé.
  await panel.fill("#startRow", "2");
  await panel.click("#intraTestBtn");
  await panel.waitForFunction(() => !$("intraTestBtn").disabled && /Ligne 2 :/.test($("intraTestResult").textContent), null, { timeout: 30000 });
  const testTxt = await panel.textContent("#intraTestResult");
  console.log("test :", testTxt);
  assert.ok(testTxt.startsWith("✓ Ligne 2 : SICOV imm 5004 → « dpe.pdf » du 2013-07-17 (2 DPE / 5 document(s) GED) → DPE/PARC FLEURI LE CONCERTO 1_lot0732_2013-07-17.pdf"), testTxt);
  assert.ok(testTxt.includes("DPE saisi dans Intralia : C 105 / GES D 25, du 05/06/2012"), testTxt);
  assert.ok(testTxt.includes("saisie OK : EXPERT IMMOBILIER (FU 0528825), ADEME 1259V1000732X, établi le 05/06/2012 → 04/06/2022, 105 kWh/m² C, 25 kg CO2/m² D"), testTxt);
  assert.strictEqual(await panel.evaluate(() => state.rows[0].length), 4, "le test n'écrit rien");
  assert.strictEqual(saved.length, 0, "le test n'enregistre rien");
  srv.stats.flower.length = 0;
  assert.strictEqual(srv.stats.index.length, 1);
  const socs2 = await panel.evaluate(() => Array.from($("intraSoc").options).map((o) => o.textContent));
  assert.deepStrictEqual(socs2, ["SICOV (19)", "CABINET NORD (31)", "CABINET SUD (42)", "Tous les cabinets (100)"], "liste lue dans diag.php");
  assert.ok(/Index SICOV : 14 lots, lu le/.test(await panel.textContent("#intraIdxInfo")));

  // 1. Repérage seul, sans analyse du PDF : aucun accès à FlowerDocs.
  await panel.uncheck("#intraAnalyze");
  await panel.selectOption("#intraMode", "find");
  assert.strictEqual(await panel.inputValue("#intraOut"), "saisie", "par défaut : colonnes utiles à la saisie");
  await panel.selectOption("#intraOut", "all");
  const run = async () => {
    await panel.click("#startBtn");
    try { await panel.waitForFunction(() => !isRunning && /Terminé|Arrêté/.test($("log").textContent), null, { timeout: 90000 }); }
    catch (e) { console.log("TIMEOUT log:\n" + await panel.evaluate(() => [isRunning, $("log").innerText].join("\n"))); throw e; }
    return panel.evaluate(() => ({ rows: state.rows, log: $("log").innerText, recap: $("resultSummary").innerText }));
  };
  const lots0 = srv.stats.lots.length;
  let out = await run();
  console.log(out.rows.map((r) => r.join(" | ")).join("\n"));
  console.log("--- journal ---\n" + out.log + "\n--- récap ---\n" + out.recap);
  const H = out.rows[0];
  const col = (n) => { const i = H.indexOf(n); assert.ok(i >= 4, "colonne " + n); return i; };
  const C = Object.fromEntries(["Intralia_Cabinet", "Intralia_ID_Immeuble", "Intralia_URL_Lot", "DPE_Statut", "DPE_Detail", "DPE_Fichier", "DPE_Date_GED", "DPE_UUID",
    "DPE_Lien_FlowerDocs", "DPE_Fichier_Local", "Intralia_DPE_Classe", "Intralia_DPE_kWh_m2", "Intralia_GES_Classe", "Intralia_GES_kgCO2_m2", "Intralia_Date_Diag"].map((n) => [n, col(n)]));
  assert.strictEqual(H.length, 4 + 15 + 1 + 17, "« toutes » : 10 colonnes fixes + 5 champs cochés + ID interne SIGEO + 17 colonnes du fichier prêt");
  const gs = (r, n) => out.rows[r - 1][H.indexOf(n)];
  assert.deepStrictEqual(["DPE analyse", "Établi le", "Valable jusqu'au", "Énergie primaire (kWh/m²/an)", "Classe énergie", "N° FU"].map((n) => gs(2, n)),
    ["À VÉRIFIER", "05/06/2012", "04/06/2022", "105", "C", ""], "sans analyse du PDF : données Intralia, fournisseur à saisir");
  const g = (r, n) => out.rows[r - 1][C[n]];
  assert.deepStrictEqual(["Intralia_Cabinet", "Intralia_ID_Immeuble", "Intralia_URL_Lot", "DPE_Statut", "DPE_Fichier", "DPE_Date_GED", "DPE_UUID", "DPE_Lien_FlowerDocs", "DPE_Fichier_Local",
    "Intralia_DPE_Classe", "Intralia_DPE_kWh_m2", "Intralia_GES_Classe", "Intralia_GES_kgCO2_m2", "Intralia_Date_Diag"].map((n) => g(2, n)),
  ["SICOV", "5004", srv.intraOrigin + "/gerance/diaglot.php?id_societe=19&id_immeuble=5004&id_lot=732", "OK", "dpe.pdf", "2013-07-17", mock.U(2),
    "https://www.oralia.flowerdocs.cloud/flower-docs-gui/#/documents/edit:" + mock.U(2), "", "C", "105", "D", "25", "05/06/2012"]);
  const st = (r) => [g(r, "DPE_Statut"), g(r, "Intralia_ID_Immeuble"), g(r, "DPE_Fichier"), g(r, "DPE_Date_GED")];
  assert.deepStrictEqual(st(3), ["OK", "5305", "DURAND_928_24.08.11_DP.pdf", "2011-08-24"]);
  assert.deepStrictEqual(st(4), ["OK", "49", "dpedupontmarie.pdf", "2007-09-25"]);
  assert.deepStrictEqual(st(5), ["OK", "5762", "MARTIN_DPE.pdf", "2015-09-01"]);
  assert.strictEqual(g(5, "DPE_UUID"), mock.U(10), "doublon : le premier de la page");
  assert.deepStrictEqual(st(6), ["AUCUN_DPE_GED", "175", "", ""]);
  assert.strictEqual(g(6, "DPE_Detail"), "0 document(s) GED, aucun DPE");
  assert.deepStrictEqual(st(7), ["AUCUN_DPE_GED", "242", "", ""]);
  assert.strictEqual(g(7, "DPE_Detail"), "2 document(s) GED, aucun DPE (entretien chaudiere 2019.pdf, additif bail.msg)");
  assert.deepStrictEqual([g(8, "DPE_Statut"), g(8, "Intralia_ID_Immeuble"), g(8, "Intralia_URL_Lot")], ["AMBIGU", "", ""]);
  assert.ok(/2 correspondances : SICOV — imm 30[01] LES ACACIAS/.test(g(8, "DPE_Detail")), g(8, "DPE_Detail"));
  assert.deepStrictEqual([g(9, "DPE_Statut"), g(9, "Intralia_ID_Immeuble"), g(9, "Intralia_URL_Lot")], ["LOT_INTROUVABLE", "5004", ""]);
  assert.strictEqual(g(10, "DPE_Statut"), "IMMEUBLE_INTROUVABLE");
  assert.ok(/proches : PARC FLEURI LE CONCERTO [12] ; PARC FLEURI LE CONCERTO [12]/.test(g(10, "DPE_Detail")), g(10, "DPE_Detail"));
  assert.strictEqual(g(11, "DPE_Statut"), "IMMEUBLE_INTROUVABLE");
  assert.deepStrictEqual(st(12), st(2), "même nom + lot que la ligne 2");
  assert.deepStrictEqual(st(13), ["OK", "600", "DPE 2021.pdf", "2021-03-01"]);
  assert.strictEqual(out.rows[13].length, 4, "ligne 14 (nom vide) non écrite");
  assert.ok(/Ligne 14 : ignorée \(nom d'immeuble vide\)/.test(out.log), out.log);
  assert.ok(/Ligne 12 .*\[déjà traité plus haut\]/.test(out.log), out.log);
  assert.ok(/Index SICOV \(19\) : 14 lots \(cache du/.test(out.log), "index chargé une seule fois (test) puis réutilisé");
  assert.ok(/6 OK · 0 erreur\(s\) ·/.test(out.recap), out.recap);
  assert.strictEqual(srv.stats.index.length, 1, "diag.php lu une seule fois");
  assert.strictEqual(srv.stats.lots.length - lots0, 7, "une page diaglot.php par lot trouvé (ligne 12 = cache)");
  assert.strictEqual(saved.length, 0, "repérage seul : rien téléchargé");
  assert.strictEqual(srv.stats.flower.length, 0, "repérage seul : aucun GET FlowerDocs");

  // 2. Repérage + téléchargement + analyse du PDF (+ Toolbox). OCR coupé :
  // les DPE simulés ont une couche texte (le « DPE 2021 » n'en a pas assez).
  await panel.selectOption("#intraMode", "download");
  await panel.check("#intraToolbox");
  await panel.check("#intraAnalyze");
  await panel.selectOption("#intraOcr", "never");
  await panel.selectOption("#intraOut", "saisie");
  await panel.evaluate((rows) => { state.rows = rows; renderColumns(); }, ROWS); // tableau vierge
  out = await run();
  console.log("--- journal (téléchargement) ---\n" + out.log);
  const H2 = out.rows[0];
  const ac = ["DPE analyse", "PDF analysé", "N° FU", "Fournisseur DPE", "N° ADEME", "Établi le", "Valable jusqu'au", "Énergie primaire (kWh/m²/an)", "Classe énergie", "GES (kg CO2/m²/an)", "Classe GES", "Remarques DPE"];
  const ga = (r) => ac.map((n) => { const i = H2.indexOf(n); assert.ok(i >= 0, "colonne " + n); return out.rows[r - 1][i]; });
  assert.deepStrictEqual(H2.slice(4), ["ID interne SIGEO", "DPE analyse", "URL fiche DPE", "Type de local", "PDF analysé", "N° FU", "Fournisseur DPE", "Établi le", "Valable jusqu'au",
    "N° ADEME", "Énergie primaire (kWh/m²/an)", "Classe énergie", "GES (kg CO2/m²/an)", "Classe GES", "Coût min (€/an)", "Coût max (€/an)",
    "Date d'indexation des prix", "Remarques DPE"], "uniquement le fichier prêt pour la saisie");
  const gx = (r, n) => out.rows[r - 1][H2.indexOf(n)];
  assert.deepStrictEqual([gx(2, "ID interne SIGEO"), gx(2, "URL fiche DPE"), gx(2, "Type de local")],
    ["38", "http://sigeo.evoriel.net/popup.aspx/fr/sig_imm/diag_manage/393001/38/0", "habitation"]);
  assert.deepStrictEqual([gx(3, "ID interne SIGEO"), gx(4, "ID interne SIGEO"), gx(12, "ID interne SIGEO"), gx(6, "ID interne SIGEO")], ["12", "5", "38", ""]);
  assert.deepStrictEqual([gx(5, "DPE analyse"), gx(5, "URL fiche DPE")], ["À VÉRIFIER", ""]);
  assert.strictEqual(gx(5, "Remarques DPE"), "fiche SIGEO introuvable : lot 278 introuvable dans SIGEO — IM0393004 : Introuvable");
  assert.ok(/Ligne 2 : SIGEO IM0393001 lot id 38 .*IM de la colonne → http:\/\/sigeo\.evoriel\.net\/popup\.aspx\/fr\/sig_imm\/diag_manage\/393001\/38\/0/.test(out.log), out.log);
  assert.ok(/Ligne 12 : SIGEO IM0393001 lot id 38 .*\[cache\]/.test(out.log), out.log);
  assert.strictEqual(H2.indexOf("Fournisseur DPE"), H2.indexOf("N° FU") + 1);
  assert.deepStrictEqual(ga(2).slice(0, 11), ["OK", "PARC FLEURI LE CONCERTO 1_lot0732_2013-07-17.pdf", "0528825", "EXPERT IMMOBILIER", "1259V1000732X", "05/06/2012", "04/06/2022", "105", "C", "25", "D"]);
  assert.deepStrictEqual(ga(3).slice(0, 5), ["À VÉRIFIER", "Cour Saint-Hubert_lot130_2011-08-24.pdf", "0394076x", "DIAG 75", "1159V1000130B"]);
  assert.deepStrictEqual(ga(4).slice(0, 5), ["À VÉRIFIER", "RESIDENCE BEAUSEJOUR_lot543_2007-09-25.pdf", "", "DIAGNOSTICS DU NORD", "non trouvé"]);
  assert.strictEqual(ga(4)[11], "fournisseur absent de la liste : « DIAGNOSTICS DU NORD »", "pas de remarque ADEME pour un DPE ancien");
  assert.strictEqual(ga(5)[7], "180,5", "énergie primaire d'Intralia (le PDF dit 180)");
  assert.strictEqual(ga(6)[11], "AUCUN_DPE_GED — 0 document(s) GED, aucun DPE");
  assert.deepStrictEqual(ga(5).slice(2, 5), ["0448957", "ACADE", "1559V1000278C"]);
  // Onglet Saisie → « Modèle DPE » : étape réglée toute seule sur le fichier prêt.
  const step = await panel.evaluate(() => {
    document.getElementById("scenarioSteps").innerHTML = "";
    $("addDpeTemplateBtn").click();
    const st = getScenarioSteps().find((x) => x.type === "sigeoDpe");
    const row = 1;
    const norm = diagNormalizeData(diagRowData(st, row));
    return { url: st.diagUrl, cols: st.diagCols, nav: buildNavUrl(st.diagUrl, state.rows[row]).url, errors: norm.errors, v: norm.v };
  });
  assert.strictEqual(step.url, "{URL fiche DPE}");
  assert.strictEqual(step.nav, "http://sigeo.evoriel.net/popup.aspx/fr/sig_imm/diag_manage/393001/38/0", "URL reprise telle quelle (pas encodée)");
  assert.deepStrictEqual(Object.assign({}, step.cols, { lot: "" }), Object.assign({}, await panel.evaluate(() => DIAG_PRESETS.intralia), { lot: "" }));
  assert.strictEqual(step.cols.lot, "{lots}");
  assert.deepStrictEqual(step.errors, [], "ligne 2 prête pour la saisie");
  assert.deepStrictEqual([step.v.type, step.v.fou, step.v.dateEtabli, step.v.conso, step.v.ges, step.v.fichier], ["11", "0528825", "05/06/2012", "105", "25", "PARC FLEURI LE CONCERTO 1_lot0732_2013-07-17.pdf"]);
  assert.deepStrictEqual([ga(6)[0], ga(8)[0], ga(9)[0]], ["Pas de PDF", "Pas de PDF", "Pas de PDF"]);
  assert.deepStrictEqual(ga(12), ga(2), "ligne déjà traitée : même analyse");
  assert.deepStrictEqual([ga(13)[0], ga(13)[5], ga(13)[11]], ["À VÉRIFIER", "02/02/2016", "analyse du PDF impossible (aucun texte exploitable (document vide ou scanné sans OCR)) : fournisseur et n° ADEME à saisir"]);
  assert.ok(/Ligne 2 : saisie OK : EXPERT IMMOBILIER \(FU 0528825\), ADEME 1259V1000732X/.test(out.log), out.log);
  assert.ok(/Ligne 4 : saisie À VÉRIFIER : DIAGNOSTICS DU NORD \(FU \?\), ADEME non trouvé.*fournisseur absent de la liste/.test(out.log), out.log);
  assert.ok(/prêts pour la saisie : 1 OK, 4 à vérifier ; 5 URL de fiche SIGEO\./.test(out.log), out.log);
  assert.ok(/Fournisseurs absents de la liste \(n° FU vide\) : DIAGNOSTICS DU NORD/.test(out.log), out.log);
  assert.strictEqual(await panel.evaluate(() => $("dsSupAddUnknown").hidden), false);
  const pre = await panel.evaluate(() => DIAG_PRESETS.analyse);
  ["N° FU", "N° ADEME", "Établi le", "PDF analysé"].forEach((c) => assert.ok(Object.values(pre).some((v) => v.includes("{" + c + "}")), "colonne lue par « SIGEO : saisir un DPE » : " + c));
  assert.deepStrictEqual(saved.map((s) => [s.name, s.head, s.url, s.conflictAction, s.saveAs]), [
    ["DPE/PARC FLEURI LE CONCERTO 1_lot0732_2013-07-17.pdf", "%PDF-", "blob:", "uniquify", false],
    ["DPE/Cour Saint-Hubert_lot130_2011-08-24.pdf", "%PDF-", "blob:", "uniquify", false],
    ["DPE/RESIDENCE BEAUSEJOUR_lot543_2007-09-25.pdf", "%PDF-", "blob:", "uniquify", false],
    ["DPE/RÉSIDENCE CHATEAUBRIAND_lot278_2015-09-01.pdf", "%PDF-", "blob:", "uniquify", false],
    ["DPE/LES TILLEULS_lot12_2021-03-01.pdf", "%PDF-", "blob:", "uniquify", false]
  ]);
  const gl = (r) => out.rows[r - 1][H2.indexOf("PDF analysé")];
  assert.deepStrictEqual([gl(2), gl(5), gl(6), gl(12)], ["PARC FLEURI LE CONCERTO 1_lot0732_2013-07-17.pdf", "RÉSIDENCE CHATEAUBRIAND_lot278_2015-09-01.pdf", "", "PARC FLEURI LE CONCERTO 1_lot0732_2013-07-17.pdf"]);
  assert.ok(/6 DPE repéré\(s\), 5 téléchargé\(s\) dans Téléchargements\/DPE/.test(out.log), out.log);
  assert.ok(/FlowerDocs : fichiers lus depuis le panneau\./.test(out.log), out.log);
  const tb = await panel.evaluate(() => tbState.docs.map((d) => [d.name, d.pages, /performance énergétique/i.test(d.text)]));
  assert.deepStrictEqual(tb.map((d) => d[0]), saved.map((s) => s.name.split("/").pop()), "PDF ajoutés à la Toolbox");
  assert.ok(tb[0][1] === 1 && tb[0][2], "texte lu par pdf.js : " + JSON.stringify(tb[0]));
  const tbDoc = await panel.evaluate(() => { const r = tbGetDocForStep({ pdfDocMode: "match", pdfDocMatch: "{PDF analysé}" }, state.rows[1]); return r.doc ? r.doc.name : r.error; });
  assert.strictEqual(tbDoc, "PARC FLEURI LE CONCERTO 1_lot0732_2013-07-17.pdf", "étape PDF β : PDF retrouvé par {PDF analysé}");
  assert.strictEqual(srv.stats.index.length, 1, "index toujours en cache");

  // 3. Vérifier le contenu (pdf.js) : le DPE 2021 est un rapport de chaudière.
  await panel.check("#intraVerify");
  await panel.uncheck("#intraToolbox");
  await panel.fill("#startRow", "13");
  await panel.fill("#endRow", "13");
  out = await run();
  console.log("--- journal (vérification) ---\n" + out.log);
  assert.deepStrictEqual(ga(13).slice(0, 5), ["OK", "LES TILLEULS_lot12_2016-02-02.pdf", "0048730", "SOCOTECH", "1659V1000012D"], "analyse du DPE vérifié");
  assert.ok(/« DPE 2021\.pdf » \(2021-03-01\) écarté : « performance énergétique » absent de la page 1 \(texte du PDF\)/.test(out.log), out.log);
  assert.ok(/Ligne 13 : ✓ « dpe\.pdf » vérifié \(texte du PDF\)/.test(out.log), out.log);
  await panel.uncheck("#intraVerify");

  // 4. FlowerDocs déconnecté : arrêt, aucun fichier enregistré.
  srv.setFlowerLogged(false);
  const n0 = saved.length;
  await panel.fill("#startRow", "2");
  await panel.fill("#endRow", "5");
  out = await run();
  console.log("--- journal (FlowerDocs déconnecté) ---\n" + out.log);
  assert.deepStrictEqual([ga(2)[0], ga(2)[1]], ["ÉCHEC", ""]);
  assert.ok(ga(2)[11].startsWith("FLOWERDOCS_NON_CONNECTE"), ga(2)[11]);
  assert.strictEqual(saved.length, n0, "jamais de page de connexion enregistrée");
  assert.ok(/relance à partir de la ligne 2/.test(out.log), out.log);
  assert.ok(!/Ligne 3 /.test(out.log), "boucle arrêtée");
  srv.setFlowerLogged(true);

  // 5. Intralia déconnecté : arrêt.
  srv.setIntraLogged(false);
  await panel.selectOption("#intraMode", "find");
  out = await run();
  assert.ok(ga(2)[11].startsWith("INTRALIA_NON_CONNECTE"), ga(2)[11]);
  assert.ok(/Intralia : session expirée .*relance à partir de la ligne 2/.test(out.log), out.log);
  assert.ok(!/Ligne 3 /.test(out.log), "boucle arrêtée");
  // Bouton « Rafraîchir » sans session : message clair, ancien index gardé.
  await panel.click("#intraIdxRefresh");
  await panel.waitForFunction(() => !$("intraIdxRefresh").disabled && /INTRALIA_NON_CONNECTE/.test($("intraIdxInfo").textContent));
  srv.setIntraLogged(true);
  await panel.click("#intraIdxRefresh");
  await panel.waitForFunction(() => !$("intraIdxRefresh").disabled && /14 lots/.test($("intraIdxInfo").textContent));
  assert.strictEqual(srv.stats.index.length, 2, "rafraîchissement = nouvelle lecture de diag.php");
  // Cache chrome.storage.local : réutilisé sans relire diag.php.
  const cached = await panel.evaluate(async () => { intraIdxMem = null; const r = await intraEnsureIndex(getIntraConfig()); return [!!r.cached, r.index.count, Object.keys(window.__store).filter((k) => k.startsWith("intralia"))]; });
  assert.deepStrictEqual(cached, [true, 14, ["intraliaCabinets", "intraliaIdx:19"]]);
  assert.strictEqual(srv.stats.index.length, 2);

  // id interne déjà connu (ex. colonne id_server) : réutilisé, pas de recherche SIGEO.
  const posts0 = lsrv.stats.post;
  const known = await panel.evaluate(async () => intraSigeoLookup({ imIdx: 0, knownIdx: 1, fiche: "", lc: getLotConfig(), run: {}, ad: null }, ["IM0393999", "77"], "5"));
  assert.deepStrictEqual(known, { status: "OK", im: "393999", id: "77", known: true, url: "http://sigeo.evoriel.net/popup.aspx/fr/sig_imm/diag_manage/393999/77/0" });
  assert.strictEqual(lsrv.stats.post, posts0);
  // Session SIGEO expirée : arrêt avec la ligne de reprise.
  await ctx.clearCookies();
  await panel.uncheck("#intraAnalyze");
  await panel.selectOption("#intraMode", "find");
  await panel.fill("#startRow", "2");
  await panel.fill("#endRow", "4");
  out = await run();
  assert.ok(/SIGEO : session expirée\. Reconnecte-toi dans l'onglet SIGEO puis relance à partir de la ligne 2\./.test(out.log), out.log);
  assert.ok(!/Ligne 3 /.test(out.log), "boucle arrêtée");
  await stab.goto(lsrv.origin + "/__login");

  // 6. Profil : configuration sauvegardée / restaurée.
  const prof = await panel.evaluate(() => {
    $("intraMode").value = "download"; $("intraFolder").value = "DPE/Intralia"; $("intraSoc").value = "31";
    $("intraAnalyze").checked = false; $("intraOcr").value = "always";
    document.querySelector('.intra-field[data-k="REFERENCE"]').checked = true;
    const c = buildProfileConfig(); applyIntraConfig({}); const def = getIntraConfig(); applyProfileConfig(c);
    const after = getIntraConfig();
    return { kind: c.extraction.kind, after: extractKind(), def: [def.mode, def.folder, def.soc, def.analyze], cfg: [after.mode, after.folder, after.soc, after.fields.join(","), after.base === c.extraction.intra.base, after.analyze, after.ocr] };
  });
  assert.deepStrictEqual(prof, { kind: "intraliaDpe", after: "intraliaDpe", def: ["download", "DPE", "19", true],
    cfg: ["download", "DPE/Intralia", "31", "libdpe1,DPE1,libges1,GES1,DATEDIAG,REFERENCE", true, false, "always"] });

  // 7. Le cas réel : seulement « immeuble » + « lots ». IM retrouvé dans la
  // recherche d'immeubles SIGEO par l'adresse Intralia, départagé par le lot.
  // Faux SIGEO sur UNE origine : recherche d'immeubles (sigeo-im-mock.html) +
  // lot_browse (sigeo-lot-server.js, relayé).
  const imHtml = fs.readFileSync(path.join(__dirname, "fixtures/sigeo-im-mock.html"), "utf8").replace("// Autres voies", [
    'add("IM0393001", "CONCERTO 1", "12 RUE DES MUSICIENS 59000 LILLE");',
    'add("IM0393005", "CONCERTO 2", "14 RUE DES MUSICIENS 59000 LILLE");',
    'add("IM0393002", "ST HUBERT GERANCE", "5 COUR SAINT HUBERT 59000 LILLE");',
    'add("IM0393022", "ST HUBERT SYNDIC", "5 COUR SAINT HUBERT 59000 LILLE");',
    // Cas Dijon : adresse Intralia vide (le nom est l'adresse), BD abrégé, nom seul.
    'add("IM0393030", "BRIFAUT 9", "9 RUE CHARLES BRIFAUT 21000 DIJON");',
    'add("IM0393031", "AMBROISE PARE", "18 BOULEVARD JEANNE D ARC 21000 DIJON");',
    'add("IM0393032", "LES PINS", "3 RUE DES ROSES 21000 DIJON");',
    'add("IM0393033", "LES PINS", "5 RUE DES PINS 75011 PARIS");', "// Autres voies"].join("\n"));
  Object.assign(lotMock.DB, { 393022: [{ lt: 7, id: 90, type: "Cave" }], // même adresse, pas le lot 130
    393030: [{ lt: 1, id: 3, type: "Appartement" }], 393031: [{ lt: 37, id: 21, type: "Appartement" }],
    393032: [{ lt: 2, id: 8, type: "Appartement" }], 393033: [{ lt: 2, id: 9, type: "Appartement" }] });
  const sigeo = await new Promise((resolve) => {
    const px = http.createServer((req, res) => {
      if (req.url.startsWith("/imsearch")) { res.writeHead(200, { "content-type": "text/html; charset=utf-8" }); res.end(imHtml); return; }
      const up = http.request(lsrv.origin + req.url, { method: req.method, headers: Object.assign({}, req.headers, { host: new URL(lsrv.origin).host }) }, (r) => {
        const h = Object.assign({}, r.headers);
        if (h.location) h.location = h.location.replace(lsrv.origin, "");
        res.writeHead(r.statusCode, h); r.pipe(res);
      });
      req.pipe(up);
    });
    px.listen(0, "127.0.0.1", () => resolve({ px, origin: `http://127.0.0.1:${px.address().port}` }));
  });
  await stab.goto(sigeo.origin + "/__login");
  await stab.goto(sigeo.origin + "/imsearch");
  await panel.evaluate(([u, rows]) => {
    $("lotUrl").value = u;
    state.rows = rows; renderColumns();
    applyIntraConfig(Object.assign(getIntraConfig(), { soc: "19", nomCol: "", lotCol: "", imCol: "", idCol: "", mode: "find", analyze: false, out: "saisie" }));
  }, [sigeo.origin + "/page.aspx/fr/sig_imm/lot_browse/syn", [["immeuble", "lots"], ["PARC FLEURI LE CONCERTO 1", "732"], ["COUR SAINT-HUBERT", "130"], ["RESIDENCE CHATEAUBRIAND", "278"]]]);
  assert.deepStrictEqual(await panel.evaluate(() => { const c = getIntraConfig(); return [c.nomCol, c.lotCol, c.imCol, c.idCol, c.sigeo]; }), ["immeuble", "lots", "", "", true]);
  await panel.fill("#startRow", "2");
  await panel.fill("#endRow", "");
  out = await run();
  console.log("--- journal (immeuble + lots seulement) ---\n" + out.log);
  const H7 = out.rows[0];
  assert.deepStrictEqual(H7.slice(2, 6), ["IM SIGEO", "ID interne SIGEO", "DPE analyse", "URL fiche DPE"]);
  const g7 = (r, n) => out.rows[r - 1][H7.indexOf(n)];
  assert.deepStrictEqual([g7(2, "IM SIGEO"), g7(2, "ID interne SIGEO"), g7(2, "URL fiche DPE")], ["IM0393001", "38", "http://sigeo.evoriel.net/popup.aspx/fr/sig_imm/diag_manage/393001/38/0"]);
  assert.deepStrictEqual([g7(3, "IM SIGEO"), g7(3, "ID interne SIGEO"), g7(3, "URL fiche DPE")], ["IM0393002", "12", "http://sigeo.evoriel.net/popup.aspx/fr/sig_imm/diag_manage/393002/12/0"]);
  assert.ok(/Ligne 3 : SIGEO IM0393002 lot id 12 .*IM par l'adresse « 5 COUR SAINT-HUBERT » \(2 candidats, départagés par le lot\)/.test(out.log), out.log);
  assert.ok(/SIGEO : recherche d'immeubles — onglet http:\/\/127\.0\.0\.1:\d+\/imsearch/.test(out.log), out.log);
  // Nom du PDF : {Immeuble} / {lots} = colonnes choisies (ici « immeuble » / « lots », en minuscules).
  assert.strictEqual(await panel.evaluate(() => intraParams(getIntraConfig(), state.rows[1], 0, null).name({ name: "dpe.pdf", date: "2013-07-17" }, ".pdf", {})),
    "DPE/Intralia/PARC FLEURI LE CONCERTO 1_lot732_2013-07-17.pdf");
  assert.deepStrictEqual([g7(4, "IM SIGEO"), g7(4, "URL fiche DPE")], ["", ""]);
  assert.ok(/^fiche SIGEO introuvable : IM introuvable \(adresse Intralia « 9 BOULEVARD DE LA LIBERTÉ 59800 LILLE »\) — recherches SIGEO : Nom « 9 BOULEVARD DE LA LIBERTÉ » : 0 résultat\(s\) ; Nom « LIBERTE » : 0 résultat\(s\) ; par adresse « LIBERTE » \+ CP 59800 : 0 résultat\(s\) ; Nom « RESIDENCE CHATEAUBRIAND » : 0 résultat\(s\)\. Ajoute une colonne « IM » au fichier pour le donner à la main/.test(g7(4, "Remarques DPE")), g7(4, "Remarques DPE"));
  // Recherches directes (cas réels de Dijon) : adresse Intralia sans voie → nom = adresse ;
  // « BD » abrégé ; nom seul (Nom E.I. identique, même CP) ; introuvable avec le détail.
  const dij = await panel.evaluate(async () => {
    const ctx = intraSigeoCtx(getIntraConfig());
    ctx.stopped = () => false;
    const go = (nom, lot, res) => intraSigeoLookup(ctx, [nom, lot], lot, res);
    return [
      await go("9 RUE CHARLES BRIFAUT", "1", { adr: "21000 DIJON", nomIntralia: "9 RUE CHARLES BRIFAUT" }),
      await go("AMBROISE PARE", "37", { adr: "18 BD JEANNE D ARC 21000 DIJON", nomIntralia: "AMBROISE PARE" }),
      await go("LES PINS", "2", { adr: "21000 DIJON", nomIntralia: "LES PINS" }),
      await go("95 RUE DE MAZY", "1", { adr: "95 RUE DE MAZY 21160 MARSANNAY LA COTE", nomIntralia: "95 RUE DE MAZY" }),
      // Colonne IM ajoutée à la main : vide → recherche auto ; remplie → utilisée telle quelle.
      await intraSigeoLookup(Object.assign({}, ctx, { imIdx: 2 }), ["9 RUE CHARLES BRIFAUT", "1", ""], "1", { adr: "21000 DIJON", nomIntralia: "" }),
      await intraSigeoLookup(Object.assign({}, ctx, { imIdx: 2 }), ["95 RUE DE MAZY", "37", "393031"], "37", { adr: "95 RUE DE MAZY 21160 MARSANNAY LA COTE" })
    ];
  });
  console.log(JSON.stringify(dij, null, 1));
  assert.deepStrictEqual(dij.slice(0, 3).map((r) => [r.status, r.im, r.id, r.via]), [
    ["OK", "393030", "3", "IM par l'adresse « 9 RUE CHARLES BRIFAUT »"],
    ["OK", "393031", "21", "IM par l'adresse « 18 BD JEANNE D ARC »"],
    ["OK", "393032", "8", "IM par le nom « LES PINS »"]]);
  assert.deepStrictEqual(dij.slice(4).map((r) => [r.status, r.im, r.id, r.via]), [
    ["OK", "393030", "3", "IM par l'adresse « 9 RUE CHARLES BRIFAUT »"], ["OK", "393031", "21", "IM de la colonne"]]);
  assert.strictEqual(dij[3].status, "INTROUVABLE");
  assert.strictEqual(dij[3].text, "IM introuvable (adresse Intralia « 95 RUE DE MAZY 21160 MARSANNAY LA COTE ») — recherches SIGEO : " +
    "Nom « 95 RUE DE MAZY » : 0 résultat(s) ; Nom « MAZY » : 0 résultat(s) ; par adresse « MAZY » + CP 21160 : 0 résultat(s). " +
    "Ajoute une colonne « IM » au fichier pour le donner à la main");
  // Recherche d'immeubles absente de l'onglet (et pas d'URL réglée) : arrêt avec un message clair.
  await stab.goto(sigeo.origin + "/page.aspx/fr/sig_imm/lot_browse/syn");
  out = await run();
  assert.ok(/SIGEO : ouvre la recherche d'immeubles SIGEO .* Relance à partir de la ligne 2\./.test(out.log), out.log);
  assert.ok(!/Ligne 3 /.test(out.log), "boucle arrêtée");
  sigeo.px.close();

  assert.deepStrictEqual(srv.stats.posts, [], "lecture seule : aucun POST vers Intralia");
  assert.deepStrictEqual(errors, [], "erreurs JS du panneau");
  console.log("\n✓ onglet Extraction (DPE extraction) OK");
  await browser.close();
  await srv.close();
  web.close();
  lsrv.server.close();
  process.exit(0);
})().catch((e) => { console.error(e); process.exit(1); });
