// Test d'intégration du panneau : onglet Extraction, type « Recherche IM
// SIGEO ». Le vrai sidepanel.html est chargé, chrome.scripting étant
// redirigé vers la page SIGEO simulée (tests/fixtures/sigeo-im-mock.html).
//   NODE_PATH=… node tests/sigeo-im.panel.e2e.js
const path = require("path");
const assert = require("assert");
const { chromium } = require("playwright");

const ROOT = path.resolve(__dirname, "..");
const PANEL = "file://" + path.join(ROOT, "sidepanel.html");
const MOCK = "file://" + path.join(ROOT, "tests/fixtures/sigeo-im-mock.html");

const CHROME_STUB = `
window.chrome = {
  storage: { local: {
    get: (k, cb) => { const r = {}; if (typeof cb === "function") cb(r); return Promise.resolve(r); },
    set: (o, cb) => { if (typeof cb === "function") cb(); return Promise.resolve(); },
    remove: (k, cb) => { if (typeof cb === "function") cb(); return Promise.resolve(); }
  } },
  runtime: { getURL: (p) => p, getManifest: () => ({ version: "0.0.0" }), sendMessage: async () => ({}),
             lastError: null, onMessage: { addListener() {} } },
  tabs: { get: async (id) => ({ id, url: "http://sigeo.test/", status: "complete" }),
          query: async () => [{ id: 1, url: "http://sigeo.test/", active: true }],
          sendMessage: async () => ({}), update: async () => ({}), duplicate: async () => ({}),
          onUpdated: { addListener() {}, removeListener() {} }, onRemoved: { addListener() {} } },
  scripting: { executeScript: async ({ func, args }) => {
    window.__calls = (window.__calls || 0) + 1;
    const r = await window.__runInMock(func.toString(), args || []);
    return [{ frameId: 0, result: r }];
  } },
  permissions: { contains: async () => false, request: async () => false },
  windows: { update: async () => ({}) }
};`;

(async () => {
  const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/usr/bin/google-chrome", headless: true });
  const ctx = await browser.newContext();
  const mock = await ctx.newPage();
  await mock.goto(MOCK);
  const panel = await ctx.newPage();
  const errors = [];
  panel.on("pageerror", (e) => errors.push(e.message));
  await panel.exposeFunction("__runInMock", (src, args) =>
    mock.evaluate(([s, a]) => (new Function("return (" + s + ")"))()(...a), [src, args]));
  await panel.addInitScript(CHROME_STUB);
  await panel.goto(PANEL);
  await panel.waitForTimeout(500);

  // L'étape n'existe plus dans le scénario de Saisie
  const inScenario = await panel.evaluate(() => { addScenarioStep({}); const r = !!document.querySelector('.scn-type option[value="sigeoIm"]'); $("scenarioSteps").innerHTML = ""; return r; });
  assert.strictEqual(inScenario, false, "plus d'étape IM dans la Saisie");

  // Données (avant le choix du type : les listes se remplissent avec les colonnes)
  await panel.evaluate(() => {
    state.headerMode = "yes";
    state.rows = [
      ["Adresse", "Ville"],
      ["129 RUE LAMARCK", ""],
      ["137 RUE LAMARCK", ""],
      ["8 BOULEVARD LAMARCK", "PARIS"],
      ["SAINT SULPICE 28", ""],
      ["158/160 RUE LAMARCK", ""],
      ["", ""]
    ];
    renderColumns();
    showTab("extraction");
  });
  // Valeurs par défaut : IM et statut vers de NOUVELLES colonnes (jamais la colonne A)
  const defaults = await panel.evaluate(() => getImxConfig());
  console.log("défauts :", JSON.stringify(defaults));
  assert.strictEqual(defaults.imCol, "N° E.I.");
  assert.strictEqual(defaults.statusCol, "Résultat IM");
  assert.strictEqual(defaults.adrCol, "Adresse");

  // Choix du type « Recherche IM SIGEO » via l'interface
  await panel.check('input[name="extractKind"][value="sigeoIm"]');
  await panel.selectOption("#imxVilleCol", "Ville");
  const vis = await panel.evaluate(() => ({ im: !$("imExtractBox").hidden, std: !$("stdExtractBox").hidden, est: $("extractEstimate").textContent }));
  console.log("UI :", JSON.stringify(vis));
  assert.ok(vis.im && !vis.std, "bloc IM visible, bloc standard masqué");
  assert.ok(/2 voies différentes/.test(vis.est), "estimation par voie");

  // Bouton « Tester sur la ligne de départ » (n'écrit rien)
  await panel.fill("#startRow", "2");
  await panel.click("#imxTestBtn");
  await panel.waitForFunction(() => !$("imxTestBtn").disabled && /→/.test($("imxTestResult").textContent), null, { timeout: 60000 });
  const testTxt = await panel.textContent("#imxTestResult");
  console.log("test :", testTxt);
  assert.ok(testTxt.includes("OK — IM0398636"));
  assert.strictEqual(await panel.evaluate(() => state.rows[0].length), 2, "le test n'écrit rien");

  // Démarrer
  await panel.click("#startBtn");
  await panel.waitForFunction(() => !isRunning && !$("startBtn").disabled && /Terminé/.test($("log").textContent), null, { timeout: 120000 });
  const out = await panel.evaluate(() => ({ rows: state.rows, log: $("log").innerText, recap: $("resultSummary").innerText }));
  console.log(out.rows.map((r) => r.join(" | ")).join("\n"));
  console.log("--- journal ---\n" + out.log + "\n--- récap ---\n" + out.recap);
  const H = out.rows[0];
  const ri = H.indexOf("Résultat IM"), ii = H.indexOf("N° E.I.");
  assert.ok(ri > 1 && ii > 1, "colonnes créées à droite");
  assert.ok(out.rows.every((r) => r.length <= 4), "pas de colonne parasite");
  assert.strictEqual(out.rows[1][ii], "IM0398636");
  assert.strictEqual(out.rows[1][ri], "OK — IM0398636");
  assert.strictEqual(out.rows[2][ri], "Multiple — IM0038854, IM0069522");
  assert.strictEqual(out.rows[2][ii], "");
  assert.ok(out.rows[3][ri].startsWith("Introuvable — proches : 8 BOULEVARD LAMARCK 18000 BOURGES"));
  assert.strictEqual(out.rows[4][ii], "IM0280028");
  assert.strictEqual(out.rows[5][ii], "IM0158160");
  assert.ok(/3 OK · 0 erreur\(s\) · 2 à vérifier · 1 ignorée/.test(out.recap), "récap : " + out.recap);
  const searches = await mock.evaluate(() => window.__searches);
  // test (LAMARCK) + exécution (LAMARCK, SULPICE) : le cache est vidé au démarrage
  assert.strictEqual(searches, 3);

  // Colonne IM = colonne adresse : un statut non OK n'efface pas l'adresse
  await panel.evaluate(() => {
    state.rows = [["Adresse"], ["129 RUE LAMARCK"], ["137 RUE LAMARCK"]];
    renderColumns();
    imxSetTarget("imxImCol", "imxImColNew", "Adresse", "target");
    imxSetTarget("imxStatusCol", "imxStatusColNew", "", "targetopt");
    fillColumnSelect($("imxVilleCol"), "", colSelectExtras("optional"));
    $("log").innerHTML = "";
  });
  await panel.click("#startBtn");
  await panel.waitForFunction(() => !isRunning && /Terminé/.test($("log").textContent), null, { timeout: 120000 });
  const same = await panel.evaluate(() => state.rows);
  console.log("même colonne :", JSON.stringify(same));
  assert.deepStrictEqual(same, [["Adresse"], ["IM0398636"], ["137 RUE LAMARCK"]]);

  // Profil : la configuration IM est sauvegardée / restaurée
  const prof = await panel.evaluate(() => { const c = buildProfileConfig(); applyProfileConfig(c); return { kind: c.extraction.kind, im: c.extraction.im, after: extractKind() }; });
  assert.strictEqual(prof.kind, "sigeoIm");
  assert.strictEqual(prof.after, "sigeoIm");

  assert.deepStrictEqual(errors, [], "erreurs JS du panneau");
  console.log("\n✓ onglet Extraction OK");
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
