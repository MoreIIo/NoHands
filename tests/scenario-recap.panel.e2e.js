// Test du panneau : récap de fin de boucle du scénario + reprise après arrêt.
// Le déroulé d'une ligne (runScenarioForRow) est simulé : on vérifie la
// mécanique de suivi (statuts, raisons, reprise, relance des erreurs,
// restauration après fermeture du panneau), pas les étapes elles-mêmes.
//   NODE_PATH=$(npm root -g) node tests/scenario-recap.panel.e2e.js
const path = require("path");
const assert = require("assert");
const { chromium } = require("playwright");

const PANEL = "file://" + path.join(path.resolve(__dirname, ".."), "sidepanel.html");

const CHROME_STUB = `
(() => {
  // Stockage tenu côté Node (survit au rechargement du panneau).
  const mem = {};
  const ready = window.__memLoad().then((m) => Object.assign(mem, m));
  const save = () => window.__memSave(JSON.parse(JSON.stringify(mem)));
  const pick = (k) => { const r = {}; (k == null ? Object.keys(mem) : [].concat(k)).forEach((x) => { if (x in mem) r[x] = mem[x]; }); return r; };
  window.chrome = {
    storage: { local: {
      get: (k, cb) => ready.then(() => { const r = pick(k); if (typeof cb === "function") cb(r); return r; }),
      set: (o, cb) => { Object.assign(mem, JSON.parse(JSON.stringify(o))); save(); if (typeof cb === "function") cb(); return Promise.resolve(); },
      remove: (k, cb) => { [].concat(k).forEach((x) => delete mem[x]); save(); if (typeof cb === "function") cb(); return Promise.resolve(); }
    } },
    runtime: { getURL: (p) => p, getManifest: () => ({ version: "0.0.0" }), sendMessage: async () => ({}),
               lastError: null, onMessage: { addListener() {} } },
    tabs: { get: async (id) => ({ id, url: "https://x.test/", status: "complete" }),
            query: async () => [{ id: 1, url: "https://x.test/", active: true }],
            sendMessage: async () => ({}), duplicate: async () => ({ id: 99 }),
            onUpdated: { addListener() {}, removeListener() {} }, onRemoved: { addListener() {} } },
    scripting: { executeScript: async () => [{ result: null }] },
    permissions: { contains: async () => false, request: async () => false },
    windows: { update: async () => ({}) }
  };
})();`;

const SETUP = () => {
  state.headerMode = "yes";
  state.originalFileName = "lots.xlsx";
  state.rows = [
    ["Immeuble", "N° lot", "Ville"],
    ["LA CROISEE", "204", "Rennes"],
    ["RESIDENCE V", "305", "Brest"],
    ["", "", ""],
    ["PARC AVENUE", "260", "Vannes"],
    ["LE REMBRANDT", "999", "Nantes"],
    ["LES JARDINS", "410", "Lorient"]
  ];
  renderColumns();
  getScenarioSteps = () => [{ type: "click" }, { type: "sigeoSelector" }];
  diagPrepareRun = async () => {};
  resolveTargetTabId = async () => 1;
  acquireWorkerTabs = async (n) => ({ tabIds: Array.from({ length: n }, (_, i) => 10 + i), created: [] });
  $("scnRowDelayMs").value = "0";
  $("scnStartRow").value = "2";
  $("scnEndRow").value = "";
  // L3 (RESIDENCE V) : erreur simple ; L6 (LE REMBRANDT) : erreur « stopper ».
  window.__fail = { 2: { abort: false }, 5: { abort: true } };
  window.__calls = [];
  runScenarioForRow = async (idx, tabId, steps, prefix) => {
    window.__calls.push(idx);
    await new Promise((r) => setTimeout(r, 20));
    const f = window.__fail[idx];
    if (!f) return { ok: true };
    if (f.abort) scnStopRequested = true;
    return { ok: false, stepIdx: 1, stepLabel: "Sélecteur SIGEO — Mandat MG", error: f.abort ? "mandat introuvable" : "valeur absente de la liste", aborted: f.abort };
  };
};

(async () => {
  const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/usr/bin/google-chrome", headless: true });
  const ctx = await browser.newContext();
  const panel = await ctx.newPage();
  const errors = [];
  panel.on("pageerror", (e) => errors.push(e.message));
  panel.on("dialog", (d) => d.accept());
  let store = {};
  await panel.exposeFunction("__memLoad", () => store);
  await panel.exposeFunction("__memSave", (m) => { store = m; });
  await panel.addInitScript(CHROME_STUB);
  await panel.goto(PANEL);
  await panel.waitForTimeout(400);
  await panel.evaluate(SETUP);

  // 1. Boucle simple : s'arrête sur L6 (étape « stopper »).
  await panel.evaluate(() => $("runScenarioLoopBtn").click());
  await panel.waitForFunction(() => !scnRunning && scnRun && scnRun.status !== "running");
  let r = await panel.evaluate(() => ({ run: scnRun, recap: $("scnRecap").innerText, hidden: $("scnRecap").hidden, calls: window.__calls }));
  console.log("--- récap 1 ---\n" + r.recap);
  assert.strictEqual(r.hidden, false);
  assert.strictEqual(r.run.status, "aborted");
  assert.deepStrictEqual(r.calls, [1, 2, 4, 5]);
  assert.strictEqual(r.run.res[3].s, "empty");
  assert.strictEqual(r.run.res[2].label, "RESIDENCE V · 305");
  assert.ok(/2 réussie\(s\).*2 en erreur.*1 vide\(s\).*1 non traitée/s.test(r.recap), r.recap);
  assert.ok(r.recap.includes("valeur absente de la liste") && r.recap.includes("mandat introuvable (boucle stoppée)"));
  assert.ok(r.recap.includes("Étape 2 — Sélecteur SIGEO — Mandat MG : 2"));
  assert.ok(r.recap.includes("Reprendre à la ligne 6 (2 lignes)"), r.recap);
  assert.ok(r.recap.includes("Relancer les 2 erreurs"));
  if (process.env.SHOT) { await panel.evaluate(() => { showTab("saisie"); $("scnRecap").scrollIntoView(); }); await panel.setViewportSize({ width: 560, height: 900 }); await panel.screenshot({ path: process.env.SHOT }); }

  // 2. Panneau rechargé : le récap et la reprise sont restaurés.
  await panel.reload();
  await panel.waitForTimeout(400);
  r = await panel.evaluate(() => ({ status: scnRun && scnRun.status, recap: $("scnRecap").innerText }));
  assert.strictEqual(r.status, "aborted");
  assert.ok(r.recap.includes("Reprendre à la ligne 6"));

  // 3. Reprise (L6 corrigée) : seules L6 et L7 sont rejouées.
  await panel.evaluate(SETUP);
  await panel.evaluate(() => { window.__fail = { 2: { abort: false } }; });
  await panel.evaluate(() => $("scnResumeBtn").click());
  await panel.waitForFunction(() => !scnRunning && scnRun.status !== "running");
  r = await panel.evaluate(() => ({ run: scnRun, recap: $("scnRecap").innerText, calls: window.__calls }));
  console.log("--- récap 2 ---\n" + r.recap);
  assert.deepStrictEqual(r.calls, [5, 6]);
  assert.strictEqual(r.run.status, "done");
  assert.ok(/4 réussie\(s\).*1 en erreur.*1 vide/s.test(r.recap), r.recap);
  assert.ok(!r.recap.includes("Reprendre à la ligne"));

  // 4. Relance des erreurs sur plusieurs onglets → 1 seule ligne, repli en boucle simple.
  await panel.evaluate(() => { window.__fail = {}; window.__calls = []; scnRun.mode = "multi"; renderScnRecap(); $("scnRetryErrBtn").click(); });
  await panel.waitForFunction(() => !scnRunning && scnRun.status !== "running");
  r = await panel.evaluate(() => ({ run: scnRun, recap: $("scnRecap").innerText, calls: window.__calls }));
  assert.deepStrictEqual(r.calls, [2]);
  assert.ok(/5 réussie\(s\).*0 en erreur/s.test(r.recap), r.recap);

  // 5. Multi-onglets : panneau fermé en cours → reprenable au rechargement.
  await panel.evaluate(SETUP);
  await panel.evaluate(() => { window.__fail = {}; $("mtTabCount").value = "2"; $("runScenarioMultiBtn").click(); });
  await panel.waitForFunction(() => scnRun && Object.keys(scnRun.res).length >= 2 && scnRun.status === "running");
  await panel.reload();
  await panel.waitForTimeout(400);
  r = await panel.evaluate(() => ({ run: scnRun, recap: $("scnRecap").innerText }));
  console.log("--- récap 3 ---\n" + r.recap);
  assert.strictEqual(r.run.status, "stopped");
  assert.strictEqual(r.run.mode, "multi");
  assert.ok(r.recap.includes("panneau fermé"));
  assert.ok(/Reprendre à la ligne \d/.test(r.recap));

  // 6. Données différentes : reprise refusée.
  await panel.evaluate(() => { state.originalFileName = "autre.xlsx"; renderScnRecap(); });
  r = await panel.evaluate(() => $("scnRecap").innerText);
  assert.ok(r.includes("ne correspondent plus"));

  assert.deepStrictEqual(errors, [], "erreurs JS : " + errors.join(" | "));
  console.log("OK — récap + reprise");
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
