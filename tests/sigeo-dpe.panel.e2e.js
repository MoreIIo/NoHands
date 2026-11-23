// Test d'intégration du panneau : onglet Extraction, type « DPE SIGEO ».
// Le vrai sidepanel.html est chargé ; chrome.scripting / chrome.tabs sont
// redirigés vers un onglet Chrome ouvert sur le faux SIGEO
// (tests/fixtures/sigeo-dpe-mock.js) ; le message « dpeDownload » au service
// worker est simulé (GET avec cookie + redirection, refus du HTML).
//   NODE_PATH=$(npm root -g) node tests/sigeo-dpe.panel.e2e.js
const path = require("path");
const assert = require("assert");
const { chromium } = require("playwright");
const mock = require("./fixtures/sigeo-dpe-mock.js");

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
    runtime: { getURL: (p) => p, getManifest: () => ({ version: "0.0.0" }),
               sendMessage: async (m) => m && m.action === "dpeDownload" ? window.__bgDownload(m.url, m.filename) : ({}),
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
    permissions: { contains: async () => true, request: async () => true },
    windows: { update: async () => ({}) }
  };
})();`;

(async () => {
  const srv = await mock.start();
  const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/usr/bin/google-chrome", headless: true });
  const ctx = await browser.newContext();
  const tab = await ctx.newPage();
  await tab.goto(srv.origin + "/__login");
  const bg = [];
  let cookie = "sigeo=ok";
  const panel = await ctx.newPage();
  const errors = [];
  panel.on("pageerror", (e) => { errors.push(e.message); console.log("[panel] " + e.message); });
  await panel.exposeFunction("__tabUrl", () => tab.url());
  await panel.exposeFunction("__bgDownload", async (url, filename) => {
    const r = await fetch(url, { headers: { cookie }, redirect: "follow" });
    const buf = Buffer.from(await r.arrayBuffer());
    if (/text\/html/.test(r.headers.get("content-type") || "")) return { ok: false, html: true, error: "SIGEO a renvoyé une page HTML (session expirée ?)" };
    bg.push({ url, filename, head: buf.subarray(0, 5).toString() });
    return { ok: true, filename: filename.split("/").pop(), size: buf.length };
  });
  await panel.exposeFunction("__navigate", async (u) => { await tab.goto(u); });
  await panel.exposeFunction("__runInTab", async (src, args) => {
    for (let i = 0; ; i++) {
      try { return await tab.evaluate(([s, a]) => (new Function("return (" + s + ")"))()(...a), [src, args]); }
      catch (e) { if (i >= 5 || !/context was destroyed|navigat/i.test(e.message)) throw e; await new Promise((r) => setTimeout(r, 200)); }
    }
  });
  await panel.addInitScript(CHROME_STUB);
  await panel.goto(PANEL);

  // Données comme dans Excel (sortie de la recherche de lot).
  await panel.evaluate((u) => {
    state.headerMode = "yes";
    state.rows = [
      ["Immeuble", "IM", "lot", "IM_format", "ID interne", "URL lot", "Résultat lot"],
      ["COUR DES M", "IM0393664", "10", "393664", "1", u.A, "OK — id 1"],
      ["RESIDENCE B", "IM0393664", "11", "393664", "2", u.B, "OK — id 2"],
      ["RESIDENCE C", "IM0393665", "3", "393665", "3", u.C, "OK — id 3"],
      ["SANS URL", "IM0393666", "4", "393666", "", "", "Introuvable"],
      ["RESIDENCE H", "IM0393667", "5", "393667", "5", u.H, "OK — id 5"]
    ];
    renderColumns();
    showTab("extraction");
  }, { A: srv.ficheUrl("A"), B: srv.docsUrl("B"), C: srv.docsUrl("C"), H: srv.docsUrl("H") });
  await panel.check('input[name="extractKind"][value="sigeoDpe"]');
  const ui = await panel.evaluate(() => ({ dpe: !$("dpeExtractBox").hidden, std: !$("stdExtractBox").hidden, lot: !$("lotExtractBox").hidden, cfg: getDpeConfig(), est: $("extractEstimate").textContent }));
  console.log("UI :", JSON.stringify(ui));
  assert.ok(ui.dpe && !ui.std && !ui.lot, "seul le bloc DPE est visible");
  assert.strictEqual(ui.cfg.url, "{URL lot}");
  assert.strictEqual(ui.cfg.folder, "DPE_SIGEO");
  assert.strictEqual(ui.cfg.name, "{IM}_{lot}_DPE_{date}");
  assert.ok(/4 lots/.test(ui.est), "estimation : " + ui.est);
  await panel.fill("#dpeDelay", "0");

  // Tester (fiche lot → page Documents, ne télécharge ni n'écrit rien)
  await panel.fill("#startRow", "2");
  await panel.click("#dpeTestBtn");
  await panel.waitForFunction(() => !$("dpeTestBtn").disabled && /Ligne 2/.test($("dpeTestResult").textContent), null, { timeout: 30000 });
  const testTxt = await panel.textContent("#dpeTestResult");
  console.log("test :", testTxt);
  assert.ok(testTxt.includes("« DPE 2022 » du 19/10/2022") && testTxt.includes("DPE_SIGEO/IM0393664_10_DPE_2022-10-19.pdf"), testTxt);
  assert.strictEqual(await panel.evaluate(() => state.rows[0].length), 7, "le test n'écrit rien");
  assert.strictEqual(bg.length, 0, "le test ne télécharge rien");

  // Exécution
  await panel.click("#startBtn");
  await panel.waitForFunction(() => !isRunning && /Terminé/.test($("log").textContent), null, { timeout: 90000 });
  const out = await panel.evaluate(() => ({ rows: state.rows, log: $("log").innerText, recap: $("resultSummary").innerText }));
  console.log(out.rows.map((r) => r.join(" | ")).join("\n"));
  console.log("--- journal ---\n" + out.log + "\n--- récap ---\n" + out.recap);
  const H = out.rows[0];
  const col = (n) => { const i = H.indexOf(n); assert.ok(i >= 7, "colonne " + n); return i; };
  const [P, N, NOM, D, S, F, ST] = ["DPE présent", "Nb DPE", "DPE nom SIGEO", "DPE date", "DPE section", "DPE fichier", "DPE statut"].map(col);
  const r1 = out.rows[1];
  assert.deepStrictEqual([r1[P], r1[N], r1[NOM], r1[D], r1[S], r1[F], r1[ST]],
    ["OUI", "1", "DPE 2022", "19/10/2022", "Diagnostics >> EPC", "IM0393664_10_DPE_2022-10-19.pdf", "OK & téléchargé"]);
  const r2 = out.rows[2];
  assert.deepStrictEqual([r2[P], r2[N], r2[NOM], r2[S], r2[F], r2[ST]],
    ["OUI", "2", "Diagnostic de performance énergétique", "Diagnostics Multys >> EPC", "IM0393664_11_DPE_2023-03-15.pdf", "OK & téléchargé"]);
  assert.deepStrictEqual([out.rows[3][P], out.rows[3][N], out.rows[3][ST], out.rows[3][F]], ["NON", "0", "KO pas trouvé", ""]);
  assert.deepStrictEqual([out.rows[4][P], out.rows[4][ST]], ["ERREUR", "Page introuvable"]);
  assert.strictEqual(out.rows[5][P], "OUI");
  assert.strictEqual(out.rows[5][ST], "KO téléchargement : SIGEO a renvoyé une page HTML (session expirée ?)");
  assert.strictEqual(out.rows[1][5], srv.ficheUrl("A"), "colonnes source intactes");
  assert.deepStrictEqual(bg.map((d) => [d.url.replace(srv.origin, ""), d.filename, d.head]), [
    ["/bare.aspx/fr/fil/ged_file_download/A03", "DPE_SIGEO/IM0393664_10_DPE_2022-10-19.pdf", "%PDF-"],
    ["/bare.aspx/fr/fil/ged_file_download/B10", "DPE_SIGEO/IM0393664_11_DPE_2023-03-15.pdf", "%PDF-"]
  ]);
  assert.strictEqual(srv.stats.posts.length, 0, "aucun POST __doPostBack");
  assert.ok(/2 DPE téléchargé/.test(out.log), out.log);

  // Textes personnalisés de « DPE statut » (variables + {Colonne})
  await panel.click("#dpeLabels summary");
  await panel.fill("#dpeLblOk", "OK {nom} du {date} ({Immeuble})");
  await panel.fill("#dpeLblAbsent", "rien");
  await panel.fill("#dpeLblPage", "pas de page");
  await panel.fill("#endRow", "5");
  await panel.click("#startBtn");
  await panel.waitForFunction(() => !isRunning && /Terminé/.test($("log").textContent), null, { timeout: 90000 });
  const out2 = await panel.evaluate(() => state.rows.map((r) => r[r.length - 1]));
  assert.deepStrictEqual(out2.slice(1, 5), ["OK DPE 2022 du 19/10/2022 (COUR DES M)", "OK Diagnostic de performance énergétique du 15/03/2023 (RESIDENCE B)", "rien", "pas de page"]);
  const lbl = await panel.evaluate(() => { const c = buildProfileConfig(); applyDpeConfig({}); const def = $("dpeLblAbsent").value; applyProfileConfig(c); return [def, $("dpeLblAbsent").value]; });
  assert.deepStrictEqual(lbl, ["KO pas trouvé", "rien"]);

  // Profil : configuration sauvegardée / restaurée
  const prof = await panel.evaluate(() => {
    $("dpeFolder").value = "DPE/2026"; $("dpeRuleExclude").value = "gaz";
    const c = buildProfileConfig(); applyDpeConfig({}); applyProfileConfig(c);
    return { kind: c.extraction.kind, after: extractKind(), folder: $("dpeFolder").value, excl: $("dpeRuleExclude").value };
  });
  assert.deepStrictEqual(prof, { kind: "sigeoDpe", after: "sigeoDpe", folder: "DPE/2026", excl: "gaz" });

  assert.deepStrictEqual(errors, [], "erreurs JS du panneau");
  console.log("\n✓ onglet Extraction (DPE SIGEO) OK");
  await browser.close();
  await srv.close();
})().catch((e) => { console.error(e); process.exit(1); });
