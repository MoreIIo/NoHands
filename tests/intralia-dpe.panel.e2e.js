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
    tabs: {
      get: async (id) => { if (id !== 1) throw new Error("No tab with id: " + id); return { id: 1, url: await window.__tabUrl(), status: "complete" }; },
      query: async () => [{ id: 1, url: await window.__tabUrl(), active: true }],
      create: async () => { throw new Error("création d'onglet non prévue dans ce test"); },
      sendMessage: async () => ({}), duplicate: async () => ({}),
      update: (id, o, cb) => { window.__navigate(o.url).then(() => { if (cb) cb({ id }); updated.forEach((l) => l(id, { status: "complete" })); }); },
      onUpdated: { addListener(l) { updated.add(l); }, removeListener(l) { updated.delete(l); } }, onRemoved: { addListener() {} }
    },
    scripting: { executeScript: async ({ target, func, args }) => {
      if (target.tabId !== 1) throw new Error("onglet inconnu");
      return [{ frameId: 0, result: await window.__runInTab(func.toString(), args || []) }];
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
  const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/usr/bin/google-chrome", headless: true });
  const ctx = await browser.newContext();
  const tab = await ctx.newPage();
  await tab.goto(srv.intraOrigin + "/");
  const saved = [];
  const panel = await ctx.newPage();
  const errors = [];
  panel.on("pageerror", (e) => { errors.push(e.message); console.log("[panel] " + e.message); });
  await panel.exposeFunction("__tabUrl", () => tab.url());
  await panel.exposeFunction("__navigate", async (u) => { await tab.goto(u); });
  await panel.exposeFunction("__runInTab", async (src, args) => tab.evaluate(([s, a]) => (new Function("return (" + s + ")"))()(...a), [src, args]));
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
  assert.deepStrictEqual([ui.cfg.soc, ui.cfg.mode, ui.cfg.folder, ui.cfg.name, ui.cfg.delay, ui.cfg.verify],
    ["19", "find", "DPE", "{Immeuble}_lot{lots}_{DPE_Date_GED}.pdf", 400, false]);
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
  assert.strictEqual(await panel.evaluate(() => state.rows[0].length), 4, "le test n'écrit rien");
  assert.strictEqual(saved.length, 0, "le test ne télécharge rien");
  assert.strictEqual(srv.stats.index.length, 1);
  const socs2 = await panel.evaluate(() => Array.from($("intraSoc").options).map((o) => o.textContent));
  assert.deepStrictEqual(socs2, ["SICOV (19)", "CABINET NORD (31)", "CABINET SUD (42)", "Tous les cabinets (100)"], "liste lue dans diag.php");
  assert.ok(/Index SICOV : 14 lots, lu le/.test(await panel.textContent("#intraIdxInfo")));

  // 1. Repérage seul
  const run = async () => {
    await panel.click("#startBtn");
    await panel.waitForFunction(() => !isRunning && /Terminé|Arrêté/.test($("log").textContent), null, { timeout: 90000 });
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
  assert.strictEqual(H.length, 4 + 15, "10 colonnes fixes + 5 champs cochés");
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
  assert.ok(/6 OK · 0 erreur\(s\) · 4 à vérifier · 3 ignorée\(s\)/.test(out.recap), out.recap);
  assert.strictEqual(srv.stats.index.length, 1, "diag.php lu une seule fois");
  assert.strictEqual(srv.stats.lots.length - lots0, 7, "une page diaglot.php par lot trouvé (ligne 12 = cache)");
  assert.strictEqual(saved.length, 0, "repérage seul : rien téléchargé");
  assert.strictEqual(srv.stats.flower.length, 0, "repérage seul : aucun GET FlowerDocs");

  // 2. Repérage + téléchargement (+ Toolbox)
  await panel.selectOption("#intraMode", "download");
  await panel.check("#intraToolbox");
  out = await run();
  console.log("--- journal (téléchargement) ---\n" + out.log);
  assert.deepStrictEqual(saved.map((s) => [s.name, s.head, s.url, s.conflictAction, s.saveAs]), [
    ["DPE/PARC FLEURI LE CONCERTO 1_lot0732_2013-07-17.pdf", "%PDF-", "blob:", "uniquify", false],
    ["DPE/Cour Saint-Hubert_lot130_2011-08-24.pdf", "%PDF-", "blob:", "uniquify", false],
    ["DPE/RESIDENCE BEAUSEJOUR_lot543_2007-09-25.pdf", "%PDF-", "blob:", "uniquify", false],
    ["DPE/RÉSIDENCE CHATEAUBRIAND_lot278_2015-09-01.pdf", "%PDF-", "blob:", "uniquify", false],
    ["DPE/LES TILLEULS_lot12_2021-03-01.pdf", "%PDF-", "blob:", "uniquify", false]
  ]);
  const gl = (r) => out.rows[r - 1][C.DPE_Fichier_Local];
  assert.deepStrictEqual([gl(2), gl(5), gl(6), gl(12)], ["PARC FLEURI LE CONCERTO 1_lot0732_2013-07-17.pdf", "RÉSIDENCE CHATEAUBRIAND_lot278_2015-09-01.pdf", "", "PARC FLEURI LE CONCERTO 1_lot0732_2013-07-17.pdf"]);
  assert.ok(/6 DPE repéré\(s\), 5 téléchargé\(s\) dans Téléchargements\/DPE/.test(out.log), out.log);
  assert.ok(/FlowerDocs : fichiers lus depuis le panneau\./.test(out.log), out.log);
  const tb = await panel.evaluate(() => tbState.docs.map((d) => [d.name, d.pages, /PERFORMANCE ENERGETIQUE/.test(d.text)]));
  assert.deepStrictEqual(tb.map((d) => d[0]), saved.map((s) => s.name.split("/").pop()), "PDF ajoutés à la Toolbox");
  assert.ok(tb[0][1] === 1 && tb[0][2], "texte lu par pdf.js : " + JSON.stringify(tb[0]));
  const tbDoc = await panel.evaluate(() => { const r = tbGetDocForStep({ pdfDocMode: "match", pdfDocMatch: "{DPE_Fichier_Local}" }, state.rows[1]); return r.doc ? r.doc.name : r.error; });
  assert.strictEqual(tbDoc, "PARC FLEURI LE CONCERTO 1_lot0732_2013-07-17.pdf", "étape PDF β : PDF retrouvé par {DPE_Fichier_Local}");
  assert.strictEqual(srv.stats.index.length, 1, "index toujours en cache");

  // 3. Vérifier le contenu (pdf.js) : le DPE 2021 est un rapport de chaudière.
  await panel.check("#intraVerify");
  await panel.uncheck("#intraToolbox");
  await panel.fill("#startRow", "13");
  await panel.fill("#endRow", "13");
  out = await run();
  console.log("--- journal (vérification) ---\n" + out.log);
  assert.deepStrictEqual([out.rows[12][C.DPE_Statut], out.rows[12][C.DPE_Fichier], out.rows[12][C.DPE_Date_GED], out.rows[12][C.DPE_Fichier_Local]],
    ["OK", "dpe.pdf", "2016-02-02", "LES TILLEULS_lot12_2016-02-02.pdf"]);
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
  assert.strictEqual(out.rows[1][C.DPE_Statut], "FLOWERDOCS_NON_CONNECTE");
  assert.strictEqual(out.rows[1][C.DPE_Fichier_Local], "");
  assert.strictEqual(saved.length, n0, "jamais de page de connexion enregistrée");
  assert.ok(/relance à partir de la ligne 2/.test(out.log), out.log);
  assert.ok(!/Ligne 3 /.test(out.log), "boucle arrêtée");
  srv.setFlowerLogged(true);

  // 5. Intralia déconnecté : arrêt.
  srv.setIntraLogged(false);
  await panel.selectOption("#intraMode", "find");
  out = await run();
  assert.strictEqual(out.rows[1][C.DPE_Statut], "INTRALIA_NON_CONNECTE");
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

  // 6. Profil : configuration sauvegardée / restaurée.
  const prof = await panel.evaluate(() => {
    $("intraMode").value = "download"; $("intraFolder").value = "DPE/Intralia"; $("intraSoc").value = "31";
    document.querySelector('.intra-field[data-k="REFERENCE"]').checked = true;
    const c = buildProfileConfig(); applyIntraConfig({}); const def = getIntraConfig(); applyProfileConfig(c);
    const after = getIntraConfig();
    return { kind: c.extraction.kind, after: extractKind(), def: [def.mode, def.folder, def.soc], cfg: [after.mode, after.folder, after.soc, after.fields.join(","), after.base === c.extraction.intra.base] };
  });
  assert.deepStrictEqual(prof, { kind: "intraliaDpe", after: "intraliaDpe", def: ["find", "DPE", "19"],
    cfg: ["download", "DPE/Intralia", "31", "libdpe1,DPE1,libges1,GES1,DATEDIAG,REFERENCE", true] });

  assert.deepStrictEqual(srv.stats.posts, [], "lecture seule : aucun POST vers Intralia");
  assert.deepStrictEqual(errors, [], "erreurs JS du panneau");
  console.log("\n✓ onglet Extraction (DPE Intralia) OK");
  await browser.close();
  await srv.close();
  web.close();
})().catch((e) => { console.error(e); process.exit(1); });
