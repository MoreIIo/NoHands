// Test d'intégration du panneau : l'étape « SIGEO : rechercher un IM »
// est créée dans la vraie interface (sidepanel.html), exécutée par le vrai
// moteur de scénario, chrome.scripting étant redirigé vers la page SIGEO
// simulée (tests/fixtures/sigeo-im-mock.html).
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

  // Données + étape
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
    addScenarioStep({ type: "sigeoIm", imAdresse: "{Adresse}", imVille: "{Ville}", imTimeout: 15000 });
  });
  // Visibilité : seuls les champs de l'étape IM sont affichés
  const vis = await panel.evaluate(() => {
    const st = document.querySelector(".scn-step");
    // Visible à l'intérieur de l'étape (le panneau « Saisie » peut être masqué).
    const v = (sel) => {
      for (let e = st.querySelector(sel); e && e !== st; e = e.parentElement) {
        if (e.hidden || getComputedStyle(e).display === "none") return false;
      }
      return !!st.querySelector(sel);
    };
    return { im: v(".scn-im-adresse"), test: v(".scn-im-test"), sigeo: v(".scn-sigeo-cp"), goto: v(".scn-goto-url"),
             icon: st.querySelector(".scn-icon use").getAttribute("href"), label: scnStepLabel(getScenarioSteps()[0]) };
  });
  console.log("UI :", JSON.stringify(vis));
  assert.ok(vis.im && vis.test && !vis.sigeo && !vis.goto, "visibilité des champs");
  assert.strictEqual(vis.icon, "#icon-step-sigeoIm");

  const step = await panel.evaluate(() => getScenarioSteps()[0]);
  assert.strictEqual(step.type, "sigeoIm");
  assert.strictEqual(step.imResultCol, "Résultat IM");
  assert.strictEqual(step.imImCol, "N° E.I.");
  assert.strictEqual(step.imMode, "nom");

  // Exécution en boucle via le moteur réel
  await panel.evaluate(async () => {
    const steps = getScenarioSteps();
    scnBegin();
    for (let i = 1; i < state.rows.length; i++) await runScenarioForRow(i, 1, steps, `L${i + 1} · `);
    scnFinish();
  });
  const out = await panel.evaluate(() => ({ rows: state.rows, log: $("scnLog").innerText, calls: window.__calls }));
  console.log(out.rows.map((r) => r.join(" | ")).join("\n"));
  assert.ok(out.rows.every((r) => r.length <= 4), "pas de colonne parasite");
  console.log("--- journal ---\n" + out.log);
  const H = out.rows[0];
  const ri = H.indexOf("Résultat IM"), ii = H.indexOf("N° E.I.");
  assert.ok(ri > 0 && ii > 0, "colonnes créées");
  assert.strictEqual(out.rows[1][ii], "IM0398636");
  assert.strictEqual(out.rows[1][ri], "OK — IM0398636");
  assert.strictEqual(out.rows[2][ri], "Multiple — IM0038854, IM0069522");
  assert.strictEqual(out.rows[2][ii], "");
  assert.ok(out.rows[3][ri].startsWith("Introuvable — proches : 8 BOULEVARD LAMARCK 18000 BOURGES"));
  assert.strictEqual(out.rows[4][ii], "IM0280028");
  assert.strictEqual(out.rows[5][ii], "IM0158160");
  assert.strictEqual(out.rows[6][ri], "ERREUR — adresse source vide");
  const searches = await mock.evaluate(() => window.__searches);
  assert.strictEqual(searches, 2, "LAMARCK + SULPICE seulement (cache)");
  assert.deepStrictEqual(errors, [], "erreurs JS du panneau");
  console.log(`\n✓ panneau OK (${searches} recherches lancées, ${out.calls} injections)`);
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
