// Tests « DPE SIGEO » (vérification + téléchargement du DPE le plus récent).
//   NODE_PATH=$(npm root -g) node tests/sigeo-dpe.jsdom.js
// Nécessite jsdom (npm i -g jsdom). La vraie fonction injectée (dpePageOp)
// tourne dans une fenêtre jsdom chargée avec la page Documents simulée
// (tests/fixtures/sigeo-dpe-mock.js) ; le POST __doPostBack part vers le
// faux serveur ASP.NET (PDF ou HTML selon le cas).
const assert = require("assert");
const { JSDOM } = require("jsdom");
const dpe = require("../sigeo-dpe.js");
const mock = require("./fixtures/sigeo-dpe-mock.js");

let failures = 0;
async function check(name, fn) {
  try { await fn(); console.log("✓ " + name); }
  catch (e) { failures++; console.log("✗ " + name + "\n   " + (e.stack || e).toString().split("\n").slice(0, 4).join("\n   ")); }
}

// Fenêtre jsdom sur une URL donnée, avec fetch / TextDecoder de Node.
function makeDom(html, url) {
  const dom = new JSDOM(html, { url, runScripts: "dangerously", pretendToBeVisual: true });
  const w = dom.window;
  w.fetch = (u, o) => fetch(new URL(u, url).href, o);
  w.TextDecoder = TextDecoder;
  return dom;
}
const runOp = (dom, cfg) => dom.window.eval("(" + dpe.dpePageOp.toString() + ")")(cfg);
const parse = (html) => runOp(makeDom("<html><body></body></html>", "http://sigeo.test/"), { op: "parse", html, baseUrl: "http://sigeo.test/popup.aspx/fr/fil/ged_multi_file_manager/sig_imm/lot_manage" });

// Adaptateur de dpeProcess : l'« onglet » est une fenêtre jsdom.
function adapter(srv, saved, logs) {
  let dom = makeDom("<html><body>accueil</body></html>", srv.origin + "/");
  return {
    navigate: async (url) => {
      const r = await fetch(url);
      dom = makeDom(await r.text(), url);
      return { ok: true };
    },
    execAll: async (cfg) => [{ frameId: 0, result: await runOp(dom, cfg) }],
    exec: async (_f, cfg) => runOp(dom, cfg),
    save: async (b64, filename, type) => { saved.push({ b64, filename, type }); return { ok: true, filename }; },
    sleep: (ms) => new Promise((r) => setTimeout(r, ms)),
    stopped: () => false,
    log: (t) => logs.push(t)
  };
}

(async () => {
  const rules = dpe.dpeCompileRules({});

  await check("normalisation : accents, encodage abîmé « ? », espaces insécables", async () => {
    assert.strictEqual(dpe.dpeFold("Entretien chaudie?re. 09.2023"), "entretien chaudiere. 09.2023");
    assert.strictEqual(dpe.dpeFold("Diagnostic de performance énergétique"), "diagnostic de performance energetique");
    assert.strictEqual(dpe.dpeFold("Diagnostic de performance ?nerg?tique"), "diagnostic de performance nergtique");
  });

  await check("filtre : DPE retenus, gaz / élec / chaudière / amiante… écartés", async () => {
    const yes = ["DPE 2022", "dpe.pdf", "DPE_2019", "Diagnostic de performance énergétique", "Diagnostics de performance ?nerg?tique",
      "Rapport performance energetique", "DPE ERP 2020", "DPE + DIAG GAZ"];
    const no = ["DIAG GAZ 2022", "DIAG ELEC 2022", "Entretien chaudie?re. 09.2023", "ENT CHAUDIERE 10.2020.pdf", "Amiante DTA",
      "Rapport performance énergétique chaudière", "ERNMT 2021", "Plan", "DPEXX", "Carrez"];
    yes.forEach((n) => assert.ok(dpe.dpeIsDpe(n, rules), "devrait être un DPE : " + n));
    no.forEach((n) => assert.ok(!dpe.dpeIsDpe(n, rules), "ne devrait pas être un DPE : " + n));
  });

  await check("règles modifiables : une regex par ligne, regex invalide signalée", async () => {
    const r = dpe.dpeCompileRules({ include: "\\bdpe\\b\nbilan thermique", exclude: "gaz\n(", section: "epc|thermique" });
    assert.strictEqual(r.errors.length, 1);
    assert.ok(dpe.dpeIsDpe("Bilan thermique 2020", r));
    assert.ok(dpe.dpeIsSection("Diagnostics >> Thermique", r));
  });

  await check("dates : dd/mm/yyyy → ISO", async () => {
    assert.strictEqual(dpe.dpeParseDate("19/10/2022").iso, "2022-10-19");
    assert.strictEqual(dpe.dpeParseDate(" 1/2/2018 ").iso, "2018-02-01");
    assert.strictEqual(dpe.dpeParseDate("2022"), null);
  });

  await check("nom de fichier : {IM}_{lot}_DPE_{date}.pdf, caractères interdits, extension serveur", async () => {
    const resolve = (s) => s.replace("{IM}", "IM0393664").replace("{lot}", "10");
    assert.strictEqual(dpe.dpeFileName(dpe.DPE_DEFAULT_NAME, { date: "2022-10-19" }, resolve), "IM0393664_10_DPE_2022-10-19.pdf");
    assert.strictEqual(dpe.dpeFileName("{IM} / {lot}:DPE*{date}?", { date: "2022-10-19" }, (s) => s.replace("{IM}", "IM 1").replace("{lot}", "A|B")),
      "IM_1_ABDPE2022-10-19.pdf");
    assert.strictEqual(dpe.dpeFileName("x_{date}", { date: "", serverName: "Rapport.PDF" }), "x_sans-date.pdf");
    assert.strictEqual(dpe.dpeFileName("x", { date: "2020-01-01", serverName: "scan.tif" }), "x.tif");
    assert.strictEqual(dpe.dpeFolder("..\\DPE_SIGEO//2026: lot"), "DPE_SIGEO/2026_lot");
  });

  // ---- Lecture de la page Documents ----
  await check("page lot A : sections par titre, en-têtes exclus, seul « DPE 2022 » retenu (pas celui de Divers)", async () => {
    const page = parse(mock.docsPage("A"));
    assert.ok(page.ready && page.isDocPage);
    const titles = page.sections.map((s) => s.title);
    assert.ok(titles.includes("Diagnostics >> EPC"), titles.join(" | "));
    assert.ok(titles.includes("Divers"), titles.join(" | "));
    assert.ok(titles.includes("Diagnostics Multys >> EPC"), "section vide signalée : " + titles.join(" | "));
    const epc = page.sections.find((s) => s.title === "Diagnostics >> EPC");
    assert.strictEqual(epc.files.length, 6);
    assert.ok(!epc.files.some((f) => /^nom$/i.test(f.name)));
    const f3 = epc.files[3];
    assert.deepStrictEqual([f3.name, f3.desc, f3.date], ["DPE 2022", "RapportDiagnosticLot", "19/10/2022"]);
    assert.ok(epc.files[3].target.endsWith("$grid$x$_ctl6$_ctl3"), epc.files[3].target);
    const sel = dpe.dpeSelect(page.sections, rules);
    assert.strictEqual(sel.sections.length, 2);
    assert.strictEqual(sel.candidates.length, 1);
    assert.strictEqual(sel.chosen.name, "DPE 2022");
    assert.strictEqual(sel.chosen.iso, "2022-10-19");
    assert.strictEqual(sel.chosen.section, "Diagnostics >> EPC");
  });

  await check("lot B : deux DPE → le plus récent (section Multys)", async () => {
    const sel = dpe.dpeSelect(parse(mock.docsPage("B")).sections, rules);
    assert.strictEqual(sel.candidates.length, 2);
    assert.strictEqual(sel.chosen.name, "Diagnostic de performance énergétique");
    assert.strictEqual(sel.chosen.section, "Diagnostics Multys >> EPC");
  });

  await check("lot C : aucun DPE", async () => {
    const sel = dpe.dpeSelect(parse(mock.docsPage("C")).sections, rules);
    assert.strictEqual(sel.candidates.length, 0);
    assert.strictEqual(sel.chosen, null);
  });

  await check("lot D : date égale → année du nom, puis ordre du tableau", async () => {
    const sel = dpe.dpeSelect(parse(mock.docsPage("D")).sections, rules);
    assert.strictEqual(sel.candidates.length, 4);
    assert.strictEqual(sel.chosen.name, "DPE 2023 (refait)");
    // Sans année dans les noms : le premier du tableau gagne.
    const tie = dpe.dpeSelect([{ title: "Diagnostics >> EPC", files: [
      { name: "DPE b", date: "10/05/2023", order: 0 }, { name: "DPE a", date: "10/05/2023", order: 1 }] }], rules);
    assert.strictEqual(tie.chosen.name, "DPE b");
  });

  await check("lot F (structure réelle : fieldset/legend, form onclick) : « 17312 (DIAG DPE2021).pdf » retenu", async () => {
    const page = parse(mock.docsPage("F"));
    const titles = page.sections.map((s) => s.title);
    ["Diagnostics >> EPC", "Diagnostics Multys >> EPC", "Divers", "Plans"].forEach((t) => assert.ok(titles.includes(t), titles.join(" | ")));
    const sel = dpe.dpeSelect(page.sections, rules);
    assert.strictEqual(sel.candidates.length, 1, JSON.stringify(sel.candidates.map((c) => c.name)));
    assert.strictEqual(sel.chosen.name, "17312 (DIAG DPE2021).pdf");
    assert.strictEqual(sel.chosen.iso, "2025-01-09");
    assert.strictEqual(sel.chosen.section, "Diagnostics >> EPC");
  });

  await check("« Chargement en cours... » visible → page pas prête", async () => {
    const page = parse(mock.docsPage("A", { loaded: false }));
    assert.strictEqual(page.ready, false);
    assert.strictEqual(page.loading, true);
  });

  await check("fiche lot : lien window.open vers la page Documents repéré", async () => {
    const page = runOp(makeDom("<html><body></body></html>", "http://sigeo.test/"),
      { op: "parse", html: mock.fichePage("A", "http://sigeo.test"), baseUrl: "http://sigeo.test/page.aspx/fr/sig_imm/lot_manage/syn/393664/1" });
    assert.strictEqual(page.isDocPage, false);
    assert.strictEqual(page.gedLinks[0], "http://sigeo.test/popup.aspx/fr/fil/ged_multi_file_manager/sig_imm/lot_manage?lot=A&x=1");
  });

  // ---- Parcours complet + téléchargement (POST __doPostBack) ----
  const srv = await mock.start();
  const P = (url, extra = {}) => Object.assign({
    url, rules: {}, timeoutMs: 5000, download: true,
    name: (c, serverName) => dpe.dpeFileName(dpe.DPE_DEFAULT_NAME, { date: c.iso, serverName }, (s) => s.replace("{IM}", "IM0393664").replace("{lot}", "10"))
  }, extra);

  await check("lot A : attente du chargement, POST avec __VIEWSTATE à jour, PDF enregistré", async () => {
    const saved = [], logs = [];
    const r = await dpe.dpeProcess(adapter(srv, saved, logs), P(srv.docsUrl("A")));
    assert.strictEqual(r.status, "OK", JSON.stringify(r));
    assert.strictEqual(r.present, "OUI");
    assert.strictEqual(r.count, 1);
    assert.strictEqual(r.statut, "Téléchargé");
    assert.strictEqual(r.file, "IM0393664_10_DPE_2022-10-19.pdf");
    assert.strictEqual(saved.length, 1);
    assert.ok(Buffer.from(saved[0].b64, "base64").toString().startsWith("%PDF-"));
    const post = srv.stats.posts[srv.stats.posts.length - 1];
    assert.strictEqual(post.viewstate, mock.VIEWSTATE);
    assert.strictEqual(post.ev, "ev-A");
    assert.ok(post.target.endsWith("$grid$x$_ctl6$_ctl3"), post.target);
  });

  await check("fiche lot → page Documents suivie automatiquement", async () => {
    const saved = [], logs = [];
    const r = await dpe.dpeProcess(adapter(srv, saved, logs), P(srv.ficheUrl("B")));
    assert.strictEqual(r.status, "OK", JSON.stringify(r));
    assert.strictEqual(r.chosen.name, "Diagnostic de performance énergétique");
    assert.strictEqual(r.file, "IM0393664_10_DPE_2023-03-15.pdf");
    assert.ok(r.info.some((i) => /fiche lot/.test(i)));
  });

  await check("lot C : Absent, rien de téléchargé", async () => {
    const saved = [];
    const r = await dpe.dpeProcess(adapter(srv, saved, []), P(srv.docsUrl("C")));
    assert.strictEqual(r.status, "ABSENT");
    assert.strictEqual(r.present, "NON");
    assert.strictEqual(r.statut, "Absent");
    assert.strictEqual(saved.length, 0);
  });

  await check("lot H : réponse HTML au lieu du PDF → KO téléchargement, sans exception", async () => {
    const saved = [], logs = [];
    const r = await dpe.dpeProcess(adapter(srv, saved, logs), P(srv.docsUrl("H")));
    assert.strictEqual(r.status, "ERREUR");
    assert.strictEqual(r.present, "OUI");
    assert.ok(/^KO téléchargement : réponse HTML au lieu d'un fichier \(validation ASP\.NET refusée\)/.test(r.statut), r.statut);
    assert.strictEqual(saved.length, 0);
    assert.ok(logs.some((l) => /POST __doPostBack refusé/.test(l)));
  });

  await check("lot F : parcours complet + téléchargement", async () => {
    const saved = [];
    const r = await dpe.dpeProcess(adapter(srv, saved, []), P(srv.docsUrl("F")));
    assert.strictEqual(r.statut, "Téléchargé", JSON.stringify(r));
    assert.strictEqual(r.file, "IM0393664_10_DPE_2025-01-09.pdf");
  });

  await check("vérification seule : trouvé, pas de POST", async () => {
    const before = srv.stats.posts.length;
    const r = await dpe.dpeProcess(adapter(srv, [], []), P(srv.docsUrl("D"), { download: false }));
    assert.strictEqual(r.statut, "Trouvé (non téléchargé)");
    assert.strictEqual(r.count, 4);
    assert.strictEqual(srv.stats.posts.length, before);
  });

  await check("URL vide / page sans documents → Page introuvable", async () => {
    let r = await dpe.dpeProcess(adapter(srv, [], []), P(""));
    assert.strictEqual(r.statut, "Page introuvable");
    r = await dpe.dpeProcess(adapter(srv, [], []), P(srv.origin + "/autre", { timeoutMs: 2000 }));
    assert.strictEqual(r.statut, "Page introuvable");
    assert.strictEqual(r.present, "ERREUR");
  });

  await srv.close();
  console.log(failures ? `\n${failures} échec(s)` : "\nTous les tests passent.");
  process.exit(failures ? 1 : 0);
})();
