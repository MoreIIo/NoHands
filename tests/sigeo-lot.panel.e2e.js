// Test d'intégration du panneau : onglet Extraction, type « Recherche d'id
// interne de lot ». Le vrai sidepanel.html est chargé ; chrome.scripting /
// chrome.tabs sont redirigés vers un onglet ouvert sur le faux serveur SIGEO
// (tests/fixtures/sigeo-lot-server.js).
//   NODE_PATH=$(npm root -g) node tests/sigeo-lot.panel.e2e.js
const path = require("path");
const assert = require("assert");
const { chromium } = require("playwright");
const mock = require("./fixtures/sigeo-lot-server.js");

const PANEL = "file://" + path.join(path.resolve(__dirname, ".."), "sidepanel.html");

const CHROME_STUB = `
(() => {
  const updated = new Set();
  window.chrome = {
    storage: { local: {
      get: (k, cb) => { const r = {}; if (typeof cb === "function") cb(r); return Promise.resolve(r); },
      set: (o, cb) => { if (typeof cb === "function") cb(); return Promise.resolve(); },
      remove: (k, cb) => { if (typeof cb === "function") cb(); return Promise.resolve(); }
    } },
    runtime: { getURL: (p) => p, getManifest: () => ({ version: "0.0.0" }), sendMessage: async () => ({}),
               lastError: null, onMessage: { addListener() {} } },
    tabs: { get: async (id) => ({ id, url: await window.__tabUrl(), status: "complete" }),
            query: async () => [{ id: 1, url: await window.__tabUrl(), active: true }],
            sendMessage: async () => ({}), duplicate: async () => ({}),
            update: (id, o, cb) => { window.__navigate(o.url).then(() => { if (cb) cb({ id }); updated.forEach((l) => l(id, { status: "complete" })); }); },
            onUpdated: { addListener(l) { updated.add(l); }, removeListener(l) { updated.delete(l); } }, onRemoved: { addListener() {} } },
    scripting: { executeScript: async ({ func, args }) => {
      const r = await window.__runInTab(func.toString(), args || []);
      return [{ frameId: 0, result: r }];
    } },
    permissions: { contains: async () => false, request: async () => false },
    windows: { update: async () => ({}) }
  };
})();`;

(async () => {
  const srv = await mock.start();
  const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/usr/bin/google-chrome", headless: true });
  const ctx = await browser.newContext();
  const tab = await ctx.newPage();
  await tab.goto(srv.origin + "/__login");
  await tab.goto(srv.url);
  const panel = await ctx.newPage();
  const errors = [];
  panel.on("pageerror", (e) => errors.push(e.message));
  await panel.exposeFunction("__tabUrl", () => tab.url());
  await panel.exposeFunction("__navigate", async (u) => { await tab.goto(u); });
  await panel.exposeFunction("__runInTab", async (src, args) => {
    for (let i = 0; ; i++) {
      try { return await tab.evaluate(([s, a]) => (new Function("return (" + s + ")"))()(...a), [src, args]); }
      catch (e) { if (i >= 5 || !/context was destroyed|navigat/i.test(e.message)) throw e; await new Promise((r) => setTimeout(r, 200)); }
    }
  });
  await panel.addInitScript(CHROME_STUB);
  await panel.goto(PANEL);
  await panel.waitForTimeout(500);

  // Données comme dans Excel : Immeuble | IM | lot | IM_format
  await panel.evaluate(() => {
    state.headerMode = "yes";
    state.rows = [
      ["Immeuble", "IM", "lot", "IM_format"],
      ["LA CROISEE", "IM0393528", "204", "393528"],
      ["RESIDENCE V", "IM0393271", "305", "393271"],
      ["PARC AVENUE", "IM0393151", "260", "393151"],
      ["LE REMBRANDT", "IM0393528", "999", "393528"],
      ["LES JARDINS", "IM0393528", "LT000204", "393528"],
      ["", "", "", ""]
    ];
    renderColumns();
    showTab("extraction");
  });
  await panel.check('input[name="extractKind"][value="sigeoLot"]');
  const ui = await panel.evaluate(() => ({ lot: !$("lotExtractBox").hidden, std: !$("stdExtractBox").hidden, im: !$("imExtractBox").hidden, cfg: getLotConfig(), est: $("extractEstimate").textContent }));
  console.log("UI :", JSON.stringify(ui));
  assert.ok(ui.lot && !ui.std && !ui.im, "seul le bloc lot est visible");
  assert.strictEqual(ui.cfg.imCol, "IM", "colonne IM devinée");
  assert.strictEqual(ui.cfg.ltCol, "lot", "colonne LT devinée");
  assert.strictEqual(ui.cfg.idCol, "ID interne");
  assert.strictEqual(ui.cfg.statusCol, "Résultat lot");
  assert.strictEqual(ui.cfg.method, "auto");
  assert.ok(/5 couples IM\/LT \(4 différents\)/.test(ui.est), "estimation : " + ui.est);

  assert.strictEqual(ui.cfg.url, "http://sigeo.evoriel.net/page.aspx/fr/sig_imm/lot_browse/syn", "URL SIGEO par défaut");
  await panel.fill("#lotUrl", srv.url);

  // Tester (n'écrit rien)
  await panel.fill("#startRow", "2");
  await panel.click("#lotTestBtn");
  await panel.waitForFunction(() => !$("lotTestBtn").disabled && /→/.test($("lotTestResult").textContent), null, { timeout: 30000 });
  const testTxt = await panel.textContent("#lotTestResult");
  console.log("test :", testTxt);
  assert.ok(testTxt.includes("OK — id 28") && testTxt.includes("POST direct"), testTxt);
  assert.strictEqual(await panel.evaluate(() => state.rows[0].length), 4, "le test n'écrit rien");

  // Exécution
  await panel.click("#startBtn");
  await panel.waitForFunction(() => !isRunning && /Terminé/.test($("log").textContent), null, { timeout: 60000 });
  let out = await panel.evaluate(() => ({ rows: state.rows, log: $("log").innerText, recap: $("resultSummary").innerText }));
  console.log(out.rows.map((r) => r.join(" | ")).join("\n"));
  console.log("--- journal ---\n" + out.log + "\n--- récap ---\n" + out.recap);
  const H = out.rows[0];
  const iId = H.indexOf("ID interne"), iUrl = H.indexOf("URL lot"), iSt = H.indexOf("Résultat lot");
  assert.ok(iId >= 4 && iUrl >= 4 && iSt >= 4, "colonnes créées à droite");
  assert.strictEqual(out.rows[1][iId], "28");
  assert.strictEqual(out.rows[1][iUrl], srv.origin + "/page.aspx/fr/sig_imm/lot_manage/syn/393528/28");
  assert.strictEqual(out.rows[1][iSt], "OK — id 28");
  assert.strictEqual(out.rows[2][iId], "7");
  assert.ok(out.rows[3][iSt].startsWith("Ambigu — id 40"), out.rows[3][iSt]);
  assert.strictEqual(out.rows[3][iId], "");
  assert.strictEqual(out.rows[4][iSt], "Introuvable");
  assert.strictEqual(out.rows[5][iId], "28", "LT000204 = 204 (cache)");
  assert.ok(/3 OK · 0 erreur\(s\) · 2 à vérifier · 1 ignorée/.test(out.recap), "récap : " + out.recap);

  // Profil : configuration sauvegardée / restaurée
  const prof = await panel.evaluate(() => { $("lotDelay").value = "50"; const c = buildProfileConfig(); applyProfileConfig(c); return { kind: c.extraction.kind, lot: c.extraction.lot, after: extractKind(), delay: $("lotDelay").value }; });
  assert.strictEqual(prof.kind, "sigeoLot");
  assert.strictEqual(prof.after, "sigeoLot");
  assert.strictEqual(prof.lot.imCol, "IM");
  assert.strictEqual(prof.delay, "50");

  // Session expirée : arrêt dès la première ligne
  await ctx.clearCookies();
  await panel.click("#startBtn");
  await panel.waitForFunction(() => !isRunning && /Terminé/.test($("log").textContent), null, { timeout: 60000 });
  out = await panel.evaluate(() => ({ rows: state.rows, log: $("log").innerText }));
  console.log("--- journal (session expirée) ---\n" + out.log);
  assert.ok(/session SIGEO expirée/.test(out.rows[1][iSt]), out.rows[1][iSt]);
  assert.strictEqual(out.rows[2][iSt], "OK — id 7", "lignes suivantes non touchées");
  assert.ok(/relance à partir de la ligne 2/.test(out.log));

  assert.deepStrictEqual(errors, [], "erreurs JS du panneau");
  console.log("\n✓ onglet Extraction (recherche de lot) OK");
  await browser.close();
  srv.server.close();
})().catch((e) => { console.error(e); process.exit(1); });
