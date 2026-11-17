// Tests de l'étape « SIGEO : rechercher un IM ».
//   node tests/sigeo-im.e2e.js
// Nécessite playwright (NODE_PATH) et un Chrome/Chromium (CHROME_PATH).
// Joue la vraie fonction injectée (imPageOp) et le vrai orchestrateur
// (imLookupAddress) contre une page SIGEO simulée, en postback partiel
// (UpdatePanel) puis complet (rechargement de page).
const path = require("path");
const assert = require("assert");
const { chromium } = require("playwright");
const im = require("../sigeo-im.js");

const FIXTURE = "file://" + path.resolve(__dirname, "fixtures/sigeo-im-mock.html");

function adapter(page) {
  const run = (cfg) => page.evaluate(([src, c]) => (new Function("return (" + src + ")"))()(c), [im.imPageOp.toString(), cfg]);
  return {
    execAll: async (cfg) => [{ frameId: 0, result: await run(cfg) }],
    exec: async (_fid, cfg) => run(cfg),
    navigate: async (url) => { try { await page.goto(url); return { ok: true }; } catch (e) { return { ok: false, error: e.message }; } },
    sleep: (ms) => new Promise((r) => setTimeout(r, ms)),
    stopped: () => false,
    log: () => {}
  };
}

const CASES = [
  { adresse: "129 RUE LAMARCK", status: "OK", ims: ["IM0398636"] },
  { adresse: "137 RUE LAMARCK", status: "Multiple", ims: ["IM0038854", "IM0069522"] },
  { adresse: "8 BOULEVARD LAMARCK", status: "Multiple", ims: ["IM0101010", "IM0101011"] },
  { adresse: "8 BOULEVARD LAMARCK", ville: "PARIS", status: "Introuvable", ims: [] },
  { adresse: "158/160 RUE LAMARCK", status: "OK", ims: ["IM0158160"] },
  { adresse: "SAINT SULPICE 28", status: "OK", ims: ["IM0280028"] },
  { adresse: "SEBASTOPOL 73", status: "OK", ims: ["IM0730073"] },
  { adresse: "CHARONNE 128 128BIS…", status: "OK", ims: ["IM0128128"] },
  { adresse: "SEBASTOPOL 99", status: "Introuvable", ims: [], closest: true },
  { adresse: "131 RUE LAMARCK", status: "Introuvable", ims: [] },                 // Nom ≠ voie
  { adresse: "131 RUE LAMARCK", mode: "adresse", status: "OK", ims: ["IM0555555"] },
  { adresse: "131 RUE LAMARCK 75018 PARIS", mode: "both", status: "OK", ims: ["IM0555555"] },
  { adresse: "12 RUE ZZTOP", status: "Introuvable", ims: [] },                    // « Aucun résultat »
  // Mode « auto » : valeur complète dans Nom, puis nom de voie si rien
  { adresse: "GRANDE ARMEE 62", mode: "auto", status: "OK", ims: ["IM0398898"], searches: 1 },
  { adresse: "GRANDE ARMEE 62", mode: "nom", status: "OK", ims: ["IM0398898"] },    // Adresse préfixée du nom
  { adresse: "RESIDENCE DU PARC", mode: "auto", status: "OK", ims: ["IM0777777"] },  // Nom E.I. exact
  { adresse: "129 RUE LAMARCK", mode: "auto", status: "OK", ims: ["IM0398636"], searches: 2 },
  { adresse: "8 RUE DE LA 1ERE ARMEE", mode: "auto", status: "OK", ims: ["IM0352337"] },
  { adresse: "129 RUE LAMARCK", sel: { submit: "#body_x_tabCriteres" }, timeoutMs: 3000, status: "ERREUR", ims: [], nocache: true }
];

async function runSuite(browser, label, url) {
  const page = await browser.newPage();
  await page.goto(url);
  const ad = adapter(page);
  im.imClearCache();
  let fails = 0;
  for (const c of CASES) {
    const t0 = Date.now();
    if (c.nocache) im.imClearCache();
    const r = await im.imLookupAddress(ad, {
      adresse: c.adresse, cp: "", ville: c.ville || "", mode: c.mode || "nom",
      siteEmpty: true, sel: c.sel || {}, url: url, timeoutMs: c.timeoutMs || 15000, maxPages: 30
    });
    const ok = r.status === c.status && JSON.stringify([...r.ims].sort()) === JSON.stringify(c.ims) &&
      (!c.closest || r.closest.length === Math.min(3, r.rowsRead)) &&
      (!c.searches || (r.searches || []).length === c.searches);
    if (!ok) fails++;
    const s = (r.searches || []).map((x) => `${x.mode}«${x.term}»:${x.rows}/${x.total} en ${x.pages}p${x.cached ? " (cache)" : ""}${x.warnings.length ? " ⚠ " + x.warnings.join(";") : ""}`).join(" + ");
    console.log(`${ok ? "✓" : "✗"} [${label}] ${c.adresse}${c.ville ? " / " + c.ville : ""}${c.mode ? " (" + c.mode + ")" : ""} → ${im.imStatusText(r)}  [${s}] ${Date.now() - t0} ms${r.error ? " ERR " + r.error : ""}`);
  }
  const siteVal = await page.$eval("#body_x_selSite_x_Site", (e) => e.value);
  assert.strictEqual(siteVal, "", "Site doit être forcé à vide");
  if (label === "partiel") {
    const n = await page.evaluate(() => window.__searches);
    // LAMARCK, SULPICE, SEBASTOPOL, CHARONNE, LAMARCK par adresse ×2 (CP), ZZTOP,
    // « GRANDE ARMEE 62 », ARMEE, « RESIDENCE DU PARC », « 129 RUE LAMARCK »
    // (LAMARCK en cache), « 8 RUE DE LA 1ERE ARMEE »
    console.log(`  recherches lancées : ${n} (attendu 12)`);
    if (n !== 12) fails++;
  }
  await page.close();
  return fails;
}

(async () => {
  // Tests unitaires de parsing / filtrage
  const S = im.imParseSource;
  assert.deepStrictEqual(S("158/160 RUE LAMARCK").nums, ["158", "160"]);
  assert.deepStrictEqual(S("CHARONNE 128 128BIS").nums, ["128", "128BIS"]);
  assert.deepStrictEqual(S("128 bis rue de Charonne").nums, ["128BIS"]);
  assert.strictEqual(im.imPickSearchTerm(S("SAINT SULPICE 28")), "SULPICE");
  assert.strictEqual(im.imPickSearchTerm(S("12 av. du Général de Gaulle")), "GAULLE");
  assert.strictEqual(im.imResolve(S("128 RUE LAMARCK"), [{ im: "IM1", adresse: "128 BIS RUE LAMARCK 75018 PARIS" }]).status, "Introuvable");
  assert.strictEqual(im.imResolve(S("8 RUE LAMARCK"), [{ im: "IM1", adresse: "8 BOULEVARD LAMARCK 18000 BOURGES" }]).status, "Introuvable");
  assert.strictEqual(im.imResolve(S("8 LAMARCK"), [{ im: "IM1", adresse: "8 BOULEVARD LAMARCK 18000 BOURGES" }]).status, "OK");
  assert.strictEqual(im.imResolve(S("129 RUE LAMARCK", { cp: "75018" }), [{ im: "IM1", adresse: "129 RUE LAMARCK 18000 BOURGES" }]).status, "Introuvable");
  console.log("✓ tests unitaires");

  const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/usr/bin/google-chrome", headless: true });
  let fails = 0;
  fails += await runSuite(browser, "partiel", FIXTURE);
  fails += await runSuite(browser, "complet", FIXTURE + "?full=1");
  await browser.close();
  console.log(fails ? `\n${fails} échec(s)` : "\nTout est OK");
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
