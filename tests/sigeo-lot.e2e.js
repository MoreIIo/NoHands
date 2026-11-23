// Tests de la « Recherche d'id interne de lot » (IM + LT → id SIGEO).
//   NODE_PATH=$(npm root -g) node tests/sigeo-lot.e2e.js
// Nécessite playwright et un Chrome/Chromium (CHROME_PATH).
// Joue la vraie fonction injectée (lotPageOp) et le vrai orchestrateur
// (lotLookup) contre un faux serveur ASP.NET (tests/fixtures/sigeo-lot-server.js) :
// POST direct (bouton submit nommé, puis lien __doPostBack), repli onglet,
// ambigu / introuvable / statut, session expirée, normalisation IM / LT.
const assert = require("assert");
const { chromium } = require("playwright");
const lot = require("../sigeo-lot.js");
const mock = require("./fixtures/sigeo-lot-server.js");

function adapter(page, logs) {
  const run = (cfg) => page.evaluate(([src, c]) => (new Function("return (" + src + ")"))()(c), [lot.lotPageOp.toString(), cfg]);
  return {
    tabUrl: async () => page.url(),
    navigate: async (url) => { try { await page.goto(url); return { ok: true }; } catch (e) { return { ok: false, error: e.message }; } },
    exec: async (_fid, cfg) => run(cfg),
    execAll: async (cfg) => [{ frameId: 0, result: await run(cfg) }],
    sleep: (ms) => new Promise((r) => setTimeout(r, ms)),
    stopped: () => false,
    log: (t) => logs.push(t)
  };
}

let failures = 0;
async function check(name, fn) {
  try { await fn(); console.log("✓ " + name); }
  catch (e) { failures++; console.log("✗ " + name + "\n   " + (e.stack || e).toString().split("\n").slice(0, 3).join("\n   ")); }
}

(async () => {
  // Normalisation (logique pure)
  await check("normalisation IM / LT", async () => {
    assert.deepStrictEqual(lot.lotNormIm("IM0393528"), { num: "393528", pad: "0393528" });
    assert.deepStrictEqual(lot.lotNormIm(" 393528 "), { num: "393528", pad: "0393528" });
    assert.deepStrictEqual(lot.lotNormIm("393528.0"), { num: "393528", pad: "0393528" });
    assert.strictEqual(lot.lotNormIm("LES JARDINS"), null);
    assert.strictEqual(lot.lotNormLt("LT000204"), "204");
    assert.strictEqual(lot.lotNormLt("lt 204"), "204");
    assert.strictEqual(lot.lotNormLt("000204"), "204");
    assert.strictEqual(lot.lotNormLt("2019"), "2019");
    assert.strictEqual(lot.lotNormLt(""), null);
  });

  const srv = await mock.start();
  const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/usr/bin/google-chrome", headless: true });
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  page.on("pageerror", (e) => console.log("   [page] " + e.message));
  await page.goto(srv.origin + "/__login");      // cookie de session
  await page.goto(srv.origin + "/login.aspx");   // onglet sur SIGEO, pas sur lot_browse
  const logs = [];
  const ad = adapter(page, logs);
  const P = (im, lt, extra = {}) => Object.assign({ im, lt, url: srv.url, statut: "actif", timeoutMs: 8000, method: "auto" }, extra);

  // Analyse des résultats (texte avec espaces insécables, lien image en double)
  await check("parse : N° LT000204 avec \\u00A0, doublon de lien, LT 2040 exclu", async () => {
    const html = `<a href="/page.aspx/fr/sig_imm/lot_manage/syn/393528/29">N° LT002040 Cave</a>
      <a href="/page.aspx/fr/sig_imm/lot_manage/syn/393528/28"><img></a>
      <a href="/page.aspx/fr/sig_imm/lot_manage/syn/393528/28">N° LT000204 Appartement Type 4</a>`;
    const r = await page.evaluate(([src, c]) => (new Function("return (" + src + ")"))()(c),
      [lot.lotPageOp.toString(), { op: "parse", html, baseUrl: "http://sigeo.evoriel.net/page.aspx/fr/sig_imm/lot_browse/syn", im: "393528", lt: "204" }]);
    assert.strictEqual(r.status, "OK");
    assert.strictEqual(r.matches[0].id, "28");
    assert.strictEqual(r.matches[0].label, "N° LT000204 Appartement Type 4");
    assert.strictEqual(r.matches[0].url, "http://sigeo.evoriel.net/page.aspx/fr/sig_imm/lot_manage/syn/393528/28");
    assert.strictEqual(r.total, 2);
  });

  // ---- POST direct, bouton submit nommé ----
  srv.setMode("input");
  lot.lotClearCache();
  await check("fetch : 0393528 / 204 → id 28 (2040 écarté, bon bouton parmi deux « Rechercher »)", async () => {
    const before = srv.stats.post;
    const r = await lot.lotLookup(ad, P("0393528", "204"), {});
    assert.strictEqual(r.status, "OK", JSON.stringify(r));
    assert.strictEqual(r.via, "fetch");
    assert.strictEqual(r.matches[0].id, "28");
    assert.strictEqual(r.matches[0].url, srv.origin + "/page.aspx/fr/sig_imm/lot_manage/syn/393528/28");
    assert.strictEqual(srv.stats.post - before, 1);
    assert.ok(page.url().startsWith(srv.origin), "l'onglet reste sur SIGEO");
    assert.ok(!page.url().includes("lot_browse"), "l'onglet n'a pas été piloté");
  });
  await check("fetch : IM0393271 / LT000305 → id 7", async () => {
    const r = await lot.lotLookup(ad, P("IM0393271", "LT000305"), {});
    assert.strictEqual(r.status, "OK", JSON.stringify(r));
    assert.strictEqual(r.matches[0].id, "7");
    assert.strictEqual(lot.lotStatusText(r), "OK — id 7");
  });
  await check("fetch : deux lots LT 260 → Ambigu avec la liste", async () => {
    const r = await lot.lotLookup(ad, P("393151", "260"), {});
    assert.strictEqual(r.status, "AMBIGU");
    assert.deepStrictEqual(r.matches.map((m) => m.id), ["40", "41"]);
    assert.ok(/^Ambigu — id 40 \(N° LT000260/.test(lot.lotStatusText(r)), lot.lotStatusText(r));
  });
  await check("fetch : LT 999 → Introuvable", async () => {
    const r = await lot.lotLookup(ad, P("393528", "999"), {});
    assert.strictEqual(r.status, "INTROUVABLE", JSON.stringify(r));
  });
  await check("fetch : statut actif → lot inactif introuvable ; statut inactif → id 30", async () => {
    const a = await lot.lotLookup(ad, P("393528", "205"), {});
    assert.strictEqual(a.status, "INTROUVABLE");
    const b = await lot.lotLookup(ad, P("393528", "205", { statut: "Inactif" }), {}); // libellé accepté
    assert.strictEqual(b.status, "OK", JSON.stringify(b));
    assert.strictEqual(b.matches[0].id, "30");
  });
  await check("cache : même couple non recherché deux fois", async () => {
    const before = srv.stats.post;
    const r = await lot.lotLookup(ad, P("IM0393528", "LT000204"), {});
    assert.strictEqual(r.cached, true);
    assert.strictEqual(srv.stats.post, before);
  });
  await check("IM / LT illisibles → ERREUR sans requête", async () => {
    const r = await lot.lotLookup(ad, P("LES JARDINS", "130"), {});
    assert.strictEqual(r.status, "ERREUR");
    assert.ok(/IM invalide/.test(r.error));
  });

  // ---- POST direct, bouton lien __doPostBack ----
  srv.setMode("link");
  lot.lotClearCache();
  await check("fetch (lien __doPostBack) : 393710 / 2019 → id 3", async () => {
    const r = await lot.lotLookup(ad, P("393710", "2019"), {});
    assert.strictEqual(r.status, "OK", JSON.stringify(r));
    assert.strictEqual(r.via, "fetch");
    assert.strictEqual(r.matches[0].id, "3");
  });

  // ---- POST direct refusé → repli onglet ----
  srv.setMode("nofetch");
  lot.lotClearCache();
  await check("auto : POST refusé (500) → repli onglet → id 28, puis onglet directement", async () => {
    logs.length = 0;
    const run = {};
    const rej = srv.stats.rejected;
    const r = await lot.lotLookup(ad, P("393528", "204"), run);
    assert.strictEqual(r.status, "OK", JSON.stringify(r));
    assert.strictEqual(r.via, "onglet");
    assert.strictEqual(r.matches[0].id, "28");
    assert.strictEqual(run.fetchBroken, true);
    assert.strictEqual(srv.stats.rejected - rej, 1);
    assert.ok(logs.some((l) => /repli/.test(l)), logs.join("\n"));
    const r2 = await lot.lotLookup(ad, P("393271", "301"), run);
    assert.strictEqual(r2.via, "onglet");
    assert.strictEqual(r2.matches[0].id, "8");
    assert.strictEqual(srv.stats.rejected - rej, 1, "plus de POST direct après le repli");
    const r3 = await lot.lotLookup(ad, P("393151", "260"), run);
    assert.strictEqual(r3.status, "AMBIGU");
  });
  await check("méthode fetch seule : POST refusé → ERREUR (pas de repli)", async () => {
    lot.lotClearCache();
    const r = await lot.lotLookup(ad, P("393528", "204", { method: "fetch" }), {});
    assert.strictEqual(r.status, "ERREUR");
    assert.ok(/HTTP 500/.test(r.error) && /validation/.test(r.error), r.error);
  });
  srv.setMode("input");
  await check("méthode onglet : 393528 / 999 → Introuvable", async () => {
    lot.lotClearCache();
    const r = await lot.lotLookup(ad, P("393528", "999", { method: "ui" }), {});
    assert.strictEqual(r.status, "INTROUVABLE", JSON.stringify(r));
    assert.strictEqual(r.via, "onglet");
  });

  // ---- Session expirée ----
  await ctx.clearCookies();
  await check("session expirée (fetch) → AUTH", async () => {
    lot.lotClearCache();
    const r = await lot.lotLookup(ad, P("393528", "204"), {});
    assert.strictEqual(r.status, "AUTH", JSON.stringify(r));
    assert.ok(/session SIGEO expirée/.test(lot.lotStatusText(r)));
  });
  await check("session expirée (onglet) → AUTH", async () => {
    lot.lotClearCache();
    const r = await lot.lotLookup(ad, P("393528", "204", { method: "ui" }), {});
    assert.strictEqual(r.status, "AUTH", JSON.stringify(r));
  });

  await browser.close();
  srv.server.close();
  console.log(failures ? `\n${failures} échec(s)` : "\nTout est OK");
  process.exit(failures ? 1 : 0);
})();
