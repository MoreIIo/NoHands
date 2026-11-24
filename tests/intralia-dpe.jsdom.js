// Tests « DPE Intralia » (index diag.php → page diaglot.php → DPE FlowerDocs).
//   NODE_PATH=$(npm root -g) node tests/intralia-dpe.jsdom.js
// Nécessite jsdom (npm i -g jsdom). La vraie fonction injectée (intraPageOp)
// tourne dans une fenêtre jsdom « ouverte sur Intralia » (fetch de Node vers
// le faux serveur tests/fixtures/intralia-mock.js, pages en windows-1252) ;
// le téléchargement FlowerDocs passe par le même faux serveur.
const assert = require("assert");
const path = require("path");
const { JSDOM } = require("jsdom");
const intra = require("../intralia-dpe.js");
const ds = require("../dpe-scan.js");
const diag = require("../sigeo-diag.js");
const mock = require("./fixtures/intralia-mock.js");
const pdfjs = require(path.join(__dirname, "..", "lib/pdf.min.js"));
pdfjs.GlobalWorkerOptions.workerSrc = path.join(__dirname, "..", "lib/pdf.worker.min.js");
const SUPPLIERS = ds.dsParseSuppliers(ds.DS_DEFAULT_SUPPLIERS).entries;
// Analyse du PDF par le vrai moteur « Analyse DPE » (texte pdf.js, sans OCR).
async function analyzeReal(bytes) {
  const r = await ds.dsReadPdf(pdfjs, bytes.slice());
  try { return await ds.dsAnalyzeDoc(r.pages, null, { ocr: "never", suppliers: SUPPLIERS }); }
  finally { r.doc.destroy(); }
}

let failures = 0;
async function check(name, fn) {
  try { await fn(); console.log("✓ " + name); }
  catch (e) { failures++; console.log("✗ " + name + "\n   " + (e.stack || e).toString().split("\n").slice(0, 5).join("\n   ")); }
}

function makeDom(url) {
  const dom = new JSDOM("<html><body></body></html>", { url, runScripts: "outside-only" });
  const w = dom.window;
  w.fetch = (u, o) => fetch(new URL(u, url).href, o);
  w.TextDecoder = TextDecoder;
  return dom;
}
// Résultat sérialisé comme par chrome.scripting.executeScript.
const runOp = async (dom, cfg) => JSON.parse(JSON.stringify(await dom.window.eval("(" + intra.intraPageOp.toString() + ")")(cfg)));

(async () => {
  const srv = await mock.start();
  const tab = makeDom(srv.intraOrigin + "/");
  const S = intra.INTRA_STATUS;
  const re = intra.intraCompileRegex("").re;

  await check("normalisation du nom (accents, ponctuation, Œ) et du lot (zéros de tête)", async () => {
    assert.strictEqual(intra.intraNormName("Île-Verte — le Concerto 1"), "ILE VERTE LE CONCERTO 1");
    assert.strictEqual(intra.intraNormName("  Résidence   Beauséjour "), "RESIDENCE BEAUSEJOUR");
    assert.strictEqual(intra.intraNormName("Cœur de ville"), "COEUR DE VILLE");
    assert.strictEqual(intra.intraNormName("COUR ST-HUBERT (bât. A)"), "COUR ST HUBERT BAT A");
    assert.strictEqual(intra.intraNormLot(" 0732 "), "732");
    assert.strictEqual(intra.intraNormLot("000"), "0");
    assert.strictEqual(intra.intraNormLot("012a"), "12A");
    assert.strictEqual(intra.intraNormLot(732), "732");
  });

  let rows19, idx19;
  await check("diag.php : décodage windows-1252 (pas UTF-8), lignes #example3 par regex", async () => {
    const r = await runOp(tab, { op: "index", url: intra.intraIndexUrl(srv.intraOrigin, "19"), soc: "19" });
    assert.ok(r.ok && !r.login, JSON.stringify(r).slice(0, 200));
    rows19 = r.rows;
    const n = mock.IMMS.filter((i) => i[0] === 19).reduce((k, i) => k + i[6].length, 0);
    assert.strictEqual(r.rows.length, n, "toutes les lignes, y compris celles sans </td>");
    const names = r.rows.map((x) => x[3]);
    assert.ok(names.includes("RÉSIDENCE BEAUSÉJOUR"), names.join(" | "));
    assert.ok(names.includes("CŒUR DE VILLE"), "Œ (0x8C) = windows-1252, pas ISO-8859-1 strict");
    assert.ok(!names.some((x) => /�|Ã/.test(x)), "aucun caractère abîmé");
    // Contre-épreuve : un res.text() (UTF-8) aurait abîmé les accents.
    const raw = await (await fetch(intra.intraIndexUrl(srv.intraOrigin, "19"))).text();
    assert.ok(raw.includes("�"), "le décodage UTF-8 abîmerait bien les accents");
    const concerto = r.rows.find((x) => x[1] === "5004" && x[2] === "732");
    assert.deepStrictEqual(concerto.slice(0, 7), ["19", "5004", "732", "PARC FLEURI LE CONCERTO 1", "SICOV", "Appartement", "732"]);
    assert.strictEqual(concerto[7], "12 RUE DES MUSICIENS 59000 LILLE");
    assert.deepStrictEqual(r.cabinets.map((c) => [c.id, c.label]), [["19", "SICOV"], ["31", "CABINET NORD"], ["42", "CABINET SUD"], ["100", "Tous les cabinets"]]);
    idx19 = intra.intraBuildIndex(rows19);
    assert.strictEqual(idx19.count, n);
  });

  await check("index : recherche nom + lot, CONCERTO 1 ≠ CONCERTO 2, zéros de tête", async () => {
    const a = intra.intraLookup(idx19, "Parc Fleuri - le Concerto 1", "0732");
    assert.strictEqual(a.status, "FOUND");
    assert.strictEqual(a.matches[0].imm, "5004");
    const b = intra.intraLookup(idx19, "PARC FLEURI LE CONCERTO 2", "732");
    assert.strictEqual(b.matches[0].imm, "5005");
    const c = intra.intraLookup(idx19, "coeur de ville", "000");
    assert.deepStrictEqual([c.status, c.matches[0].imm, c.matches[0].lot], ["FOUND", "610", "0"]);
  });

  await check("index : AMBIGU (même cabinet), LOT_INTROUVABLE, IMMEUBLE_INTROUVABLE + 3 noms proches max", async () => {
    const a = intra.intraLookup(idx19, "LES ACACIAS", "3");
    assert.strictEqual(a.status, S.AMBIGU);
    assert.deepStrictEqual(a.matches.map((m) => m.imm).sort(), ["300", "301"]);
    const l = intra.intraLookup(idx19, "PARC FLEURI LE CONCERTO 1", "999");
    assert.strictEqual(l.status, S.LOT_INTROUVABLE);
    assert.deepStrictEqual(l.buildings.map((b) => [b.imm, b.lots]), [["5004", 2]]);
    const i = intra.intraLookup(idx19, "PARC FLEURI CONCERTO", "732");
    assert.strictEqual(i.status, S.IMMEUBLE_INTROUVABLE);
    assert.ok(i.close.length >= 2 && i.close.length <= 3, i.close.join(" | "));
    assert.ok(i.close.slice(0, 2).every((x) => /CONCERTO/.test(x)), i.close.join(" | "));
    const z = intra.intraLookup(idx19, "LE ZOLA", "7");
    assert.strictEqual(z.status, S.IMMEUBLE_INTROUVABLE, "LE ZOLA n'est pas chez Sicov");
  });

  await check("« Tous les cabinets » (100) : collisions entre cabinets → AMBIGU", async () => {
    const r = await runOp(tab, { op: "index", url: intra.intraIndexUrl(srv.intraOrigin, "100"), soc: "100" });
    const idx = intra.intraBuildIndex(r.rows);
    const z = intra.intraLookup(idx, "LE ZOLA", "7");
    assert.strictEqual(z.status, S.AMBIGU);
    assert.deepStrictEqual(z.matches.map((m) => `${m.cabinet}/${m.imm}`).sort(), ["CABINET NORD/880", "CABINET SUD/1201"]);
    const c = intra.intraLookup(idx, "PARC FLEURI LE CONCERTO 1", "732");
    assert.strictEqual(c.status, S.AMBIGU, "même nom + lot chez Sicov et Cabinet Nord");
    const res = await intra.intraProcessRow({ stopped: () => false }, { idx, nom: "LE ZOLA", lot: "7", re });
    assert.strictEqual(res.status, S.AMBIGU);
    assert.ok(/CABINET NORD — imm 880 LE ZOLA \(7 RUE ZOLA 62000 ARRAS\)/.test(res.detail), res.detail);
  });

  await check("diaglot.php : DPE saisi (form1, valeur posée en JS comprise) + documents GED", async () => {
    const p = await runOp(tab, { op: "lot", url: intra.intraLotUrl(srv.intraOrigin, "19", "5004", "732") });
    assert.ok(p.ok && !p.login);
    assert.deepStrictEqual(p.fields, {
      libdpe1: "C", libges1: "D", typediag: "avant le 01/07/2021", DPE1: "105", GES1: "25", DATEDIAG: "05/06/2012",
      datefinval: "31/12/2012", MTDIAGMIN: "", MTDIAGMAX: "", MTDIAG: "574", ANNEREFDIAG: "2010", REFERENCE: "1032/5004/0732"
    });
    assert.deepStrictEqual(p.docs.map((d) => [d.name, d.date]), mock.LOTS["19/5004/732"].docs.map((d) => [d[0], d[1]]));
    assert.strictEqual(p.docs[1].uuid, mock.U(2));
    assert.strictEqual(p.docs[1].kind, "RapportDiagnostic");
    assert.ok(p.docs[1].link.endsWith("documents/edit:" + mock.U(2)));
  });

  await check("filtre DPE par nom (regex par défaut) : cas réels et documents écartés", async () => {
    const yes = ["dpe.pdf", "dpemoniotfrancoi.pdf", "dpedupontmarie.pdf", "TISSERAND_DPE.pdf", "MARTIN_DPE.pdf", "GHESQUIER_928_24.08.11_DP.pdf",
      "DURAND_928_24.08.11_DP.pdf", "Diagnostic de performance énergétique.pdf", "DPE ancien.pdf", "perf énergétique 2019.pdf", "DP.pdf"];
    const no = ["entretien chaudière 2020.pdf", "EDL entrée.pdf", "additif bail.doc", "courrier.msg", "photo façade.png", "diagnostic gaz.pdf",
      "état des lieux.pdf", "DPPE.pdf", "Etat_daté.pdf", "independance.pdf"];
    yes.forEach((n) => assert.ok(intra.intraIsDpeName(n, re), "devrait être un DPE : " + n));
    no.forEach((n) => assert.ok(!intra.intraIsDpeName(n, re), "ne devrait pas être un DPE : " + n));
    assert.ok(intra.intraCompileRegex("dpe(").error, "regex invalide signalée");
    assert.ok(!intra.intraCompileRegex("^dpe").error);
  });

  await check("plus récent par date « crée le », doublons nom + date, égalité = premier de la page", async () => {
    const docs = [
      { uuid: "a", name: "dpe.pdf", date: "2013-07-17" },
      { uuid: "b", name: "DPE ancien.pdf", date: "2009-05-02" },
      { uuid: "c", name: "MARTIN_DPE.pdf", date: "2015-09-01" },
      { uuid: "d", name: "MARTIN_DPE.pdf", date: "2015-09-01" },
      { uuid: "e", name: "autre dpe.pdf", date: "2015-09-01" },
      { uuid: "f", name: "dpe sans date.pdf", date: "" },
      { uuid: "g", name: "entretien.pdf", date: "2020-01-01" }
    ];
    const s = intra.intraSelectDpe(docs, re);
    assert.strictEqual(s.total, 7);
    assert.deepStrictEqual(s.sorted.map((d) => d.uuid), ["c", "e", "a", "b", "f"]);
    assert.strictEqual(s.chosen.uuid, "c");
  });

  await check("page de connexion Intralia détectée (pas de #example3 / pas de form1)", async () => {
    srv.setIntraLogged(false);
    try {
      const i = await runOp(tab, { op: "index", url: intra.intraIndexUrl(srv.intraOrigin, "19") });
      assert.deepStrictEqual([i.ok, i.login], [false, true]);
      const l = await runOp(tab, { op: "lot", url: intra.intraLotUrl(srv.intraOrigin, "19", "5004", "732") });
      assert.deepStrictEqual([l.ok, l.login], [false, true]);
    } finally { srv.setIntraLogged(true); }
    const u = await runOp(tab, { op: "lot", url: intra.intraLotUrl(srv.intraOrigin, "19", "5004", "1") });
    assert.deepStrictEqual([u.ok, u.login], [false, false], "lot inconnu ≠ page de connexion");
    const c = await runOp(tab, { op: "parseCabinets", html: mock.diagFormPage() });
    assert.strictEqual(c.cabinets.length, 4);
  });

  await check("contrôle du fichier FlowerDocs : %PDF, page HTML, 401, 500, autre type", async () => {
    const b = (s) => new Uint8Array(Buffer.from(s, "latin1"));
    const k = (r, n) => intra.intraClassifyFile(r, n).kind;
    assert.strictEqual(k({ status: 200, contentType: "application/octet-stream", bytes: b("%PDF-1.4 …") }, "dpe.pdf"), "PDF");
    assert.strictEqual(k({ status: 200, contentType: "text/html", bytes: b("<!DOCTYPE html><form>…") }, "dpe.pdf"), "LOGIN");
    assert.strictEqual(k({ status: 200, contentType: "application/octet-stream", bytes: b("  <html>login") }, "dpe.pdf"), "LOGIN");
    assert.strictEqual(k({ status: 401, bytes: b("") }, "dpe.pdf"), "LOGIN");
    assert.strictEqual(k({ status: 500, contentType: "text/html", bytes: b("<html>erreur") }, "dpe.pdf"), "HTTP");
    assert.strictEqual(k({ status: 200, bytes: b("") }, "dpe.pdf"), "HTTP");
    assert.strictEqual(k({ status: 200, contentType: "application/octet-stream", bytes: b("PK\u0003\u0004") }, "dpe.pdf"), "NOT_PDF");
    const o = intra.intraClassifyFile({ status: 200, contentType: "application/msword", bytes: b("ÐÏ\u0011à") }, "dpe.doc");
    assert.deepStrictEqual([o.kind, o.ext], ["OTHER", ".doc"]);
  });

  await check("nom de fichier : modèle par défaut, {Colonne}, caractères interdits → _, extension réelle", async () => {
    const row = { Immeuble: "PARC FLEURI LE CONCERTO 1", lots: "0732" };
    const resolve = (s) => s.replace(/\{([^{}]+)\}/g, (w, k) => (k in row ? row[k] : w));
    const c = { name: "dpe.pdf", date: "2013-07-17", uuid: mock.U(2) };
    const v = intra.intraTemplateVars(c, { cabinet: "SICOV", imm: "5004", fields: { DATEDIAG: "05/06/2012" } });
    assert.strictEqual(intra.intraFileName(intra.INTRA_DEFAULT_NAME, v, resolve, ".pdf"), "PARC FLEURI LE CONCERTO 1_lot0732_2013-07-17.pdf");
    const row2 = { Immeuble: 'A/B: "C"* <D>|E?', lots: "7" };
    const r2 = (s) => s.replace(/\{([^{}]+)\}/g, (w, k) => (k in row2 ? row2[k] : w));
    assert.strictEqual(intra.intraFileName("{Immeuble}_lot{lots}_{DPE_Date_GED}.pdf", v, r2, ".pdf"), "A_B_ _C__ _D__E__lot7_2013-07-17.pdf");
    assert.strictEqual(intra.intraFileName("{Intralia_ID_Immeuble}_{DPE_Fichier}_{Intralia_Date_Diag}", v, resolve, ".pdf"), "5004_dpe_05_06_2012.pdf");
    assert.strictEqual(intra.intraFileName("{Immeuble}.pdf", v, resolve, ".doc"), "PARC FLEURI LE CONCERTO 1.doc", "extension du fichier reçu");
    assert.strictEqual(intra.intraFileName("{Inconnue}", {}, resolve, ".pdf"), "DPE.pdf");
    assert.strictEqual(intra.intraFileName("CON", {}, null, ".pdf"), "_CON.pdf");
    assert.strictEqual(intra.intraFolder("DPE\\2026/../x:y"), "DPE/2026/x_y");
  });

  await check("configuration par défaut (profils / stockage)", async () => {
    const c = intra.intraNormalizeConfig({});
    assert.deepStrictEqual([c.soc, c.mode, c.folder, c.name, c.regex, c.verify, c.delay, c.maxAgeDays, c.base, c.flower, c.analyze, c.ocr],
      ["19", "find", "DPE", "{Immeuble}_lot{lots}_{DPE_Date_GED}.pdf", intra.INTRA_DEFAULT_REGEX, false, 400, 7, "https://intralia.oralia.local", "https://www.oralia.flowerdocs.cloud", true, "auto"]);
    assert.deepStrictEqual([intra.intraNormalizeConfig({ analyze: false, ocr: "never" }).analyze, intra.intraNormalizeConfig({ ocr: "x" }).ocr], [false, "auto"]);
    assert.deepStrictEqual(c.fields, ["libdpe1", "DPE1", "libges1", "GES1", "DATEDIAG"]);
    const d = intra.intraNormalizeConfig({ mode: "download", folder: "", fields: ["REFERENCE", "inconnu"], soc: "100", base: "https://x.local/", delay: "0" });
    assert.deepStrictEqual([d.mode, d.folder, d.fields, d.soc, d.base, d.delay], ["download", "", ["REFERENCE"], "100", "https://x.local", 0]);
  });

  // Adaptateur de intraProcessRow : onglet Intralia = jsdom ; FlowerDocs = fetch de Node.
  const saved = [];
  const pdfText = (u8) => Buffer.from(u8).toString("latin1");
  function adapter(over = {}) {
    return Object.assign({
      intra: (cfg) => runOp(tab, cfg),
      fetchFile: async (url) => {
        const r = await fetch(url, { redirect: "follow" });
        return { status: r.status, contentType: r.headers.get("content-type") || "", disposition: r.headers.get("content-disposition") || "", bytes: new Uint8Array(await r.arrayBuffer()) };
      },
      save: async (bytes, path) => { saved.push({ path, head: Buffer.from(bytes.subarray(0, 5)).toString("latin1") }); return { ok: true, filename: path.split("/").pop() }; },
      verify: async (bytes) => ({ ok: /performance [ée]nerg/i.test(pdfText(bytes)), how: "texte", why: "absent" }),
      analyze: (bytes) => analyzeReal(bytes),
      stopped: () => false, log: () => {}
    }, over);
  }
  const params = (nom, lot, over = {}) => Object.assign({
    idx: idx19, nom, lot, re, base: srv.intraOrigin, flower: srv.flowerOrigin, download: true, verify: false, socLabel: "SICOV",
    name: (c, ext, res) => "DPE/" + intra.intraFileName(intra.INTRA_DEFAULT_NAME, intra.intraTemplateVars(c, res), (s) => s.replace("{Immeuble}", nom).replace("{lots}", lot), ext)
  }, over);

  await check("6 cas réels (anonymisés) : repérage + téléchargement", async () => {
    saved.length = 0;
    for (const k of mock.REAL_CASES) {
      const res = await intra.intraProcessRow(adapter(), params(k.nom, k.lot));
      assert.strictEqual(res.status, k.status, `${k.nom} / ${k.lot} : ${res.status} — ${res.detail}`);
      assert.strictEqual(res.imm, k.imm);
      assert.strictEqual(res.cabinet, "SICOV");
      assert.strictEqual(res.lotUrl, intra.intraLotUrl(srv.intraOrigin, "19", k.imm, k.lot));
      if (k.status === "OK") {
        assert.deepStrictEqual([res.chosen.name, res.chosen.date, res.chosen.uuid], [k.file, k.date, k.uuid]);
        assert.strictEqual(res.file, `${k.nom}_lot${k.lot}_${k.date}.pdf`);
        assert.ok(res.chosen.view.endsWith("documents/edit:" + k.uuid));
      } else {
        assert.strictEqual(res.nDocs, k.nDocs);
        assert.ok(res.detail.startsWith(`${k.nDocs} document(s) GED, aucun DPE`), res.detail);
        assert.strictEqual(res.file, "");
      }
    }
    assert.deepStrictEqual(saved.map((s) => [s.path, s.head]), [
      ["DPE/PARC FLEURI LE CONCERTO 1_lot732_2013-07-17.pdf", "%PDF-"],
      ["DPE/COUR SAINT-HUBERT_lot130_2011-08-24.pdf", "%PDF-"],
      ["DPE/RESIDENCE BEAUSEJOUR_lot543_2007-09-25.pdf", "%PDF-"],
      ["DPE/RESIDENCE CHATEAUBRIAND_lot278_2015-09-01.pdf", "%PDF-"]
    ]);
    const v = intra.intraRowValues(await intra.intraProcessRow(adapter(), params("PARC FLEURI LE CONCERTO 1", "732", { download: false })), ["libdpe1", "DATEDIAG"]);
    assert.deepStrictEqual(v, {
      Intralia_Cabinet: "SICOV", Intralia_ID_Immeuble: "5004", Intralia_URL_Lot: intra.intraLotUrl(srv.intraOrigin, "19", "5004", "732"),
      DPE_Statut: "OK", DPE_Detail: "2 DPE sur 5 document(s) GED (autres : DPE ancien.pdf 2009-05-02)",
      DPE_Fichier: "dpe.pdf", DPE_Date_GED: "2013-07-17", DPE_UUID: mock.U(2),
      DPE_Lien_FlowerDocs: "https://www.oralia.flowerdocs.cloud/flower-docs-gui/#/documents/edit:" + mock.U(2),
      DPE_Fichier_Local: "", Intralia_DPE_Classe: "C", Intralia_Date_Diag: "05/06/2012"
    });
  });

  await check("repérage seul : rien n'est téléchargé", async () => {
    saved.length = 0;
    const n0 = srv.stats.flower.length;
    const res = await intra.intraProcessRow(adapter(), params("COUR SAINT-HUBERT", "130", { download: false }));
    assert.strictEqual(res.status, "OK");
    assert.strictEqual(saved.length, 0);
    assert.strictEqual(srv.stats.flower.length, n0, "aucun GET FlowerDocs");
  });

  await check("vérifier le contenu : le plus récent qui contient « performance énergétique »", async () => {
    const plain = await intra.intraProcessRow(adapter(), params("LES TILLEULS", "12", { download: false }));
    assert.strictEqual(plain.chosen.name, "DPE 2021.pdf", "sans vérification : le plus récent par date");
    const res = await intra.intraProcessRow(adapter(), params("LES TILLEULS", "12", { verify: true }));
    assert.strictEqual(res.status, "OK");
    assert.deepStrictEqual([res.chosen.name, res.chosen.date], ["dpe.pdf", "2016-02-02"]);
    assert.ok(res.info.some((x) => /« DPE 2021\.pdf » \(2021-03-01\) écarté/.test(x)), res.info.join(" | "));
    const none = await intra.intraProcessRow(adapter({ verify: async () => ({ ok: false, why: "non" }) }), params("LES TILLEULS", "12", { verify: true }));
    assert.strictEqual(none.status, S.AUCUN_DPE_GED);
    assert.ok(/2 candidat\(s\) écarté\(s\)/.test(none.detail), none.detail);
  });

  await check("FlowerDocs déconnecté → FLOWERDOCS_NON_CONNECTE, rien d'enregistré", async () => {
    saved.length = 0;
    srv.setFlowerLogged(false);
    try {
      const res = await intra.intraProcessRow(adapter(), params("PARC FLEURI LE CONCERTO 1", "732"));
      assert.strictEqual(res.status, S.FLOWERDOCS_NON_CONNECTE);
      assert.ok(/page HTML/.test(res.detail), res.detail);
      assert.strictEqual(saved.length, 0, "jamais de page de connexion enregistrée en .pdf");
      const net = await intra.intraProcessRow(adapter({ fetchFile: async () => ({ error: "Failed to fetch" }) }), params("PARC FLEURI LE CONCERTO 1", "732"));
      assert.strictEqual(net.status, S.FLOWERDOCS_NON_CONNECTE);
    } finally { srv.setFlowerLogged(true); }
  });

  await check("Intralia déconnecté pendant le traitement → INTRALIA_NON_CONNECTE", async () => {
    srv.setIntraLogged(false);
    try {
      const res = await intra.intraProcessRow(adapter(), params("PARC FLEURI LE CONCERTO 1", "732"));
      assert.strictEqual(res.status, S.INTRALIA_NON_CONNECTE);
      assert.strictEqual(res.imm, "5004");
    } finally { srv.setIntraLogged(true); }
  });

  await check("autres statuts : LOT_INTROUVABLE / IMMEUBLE_INTROUVABLE / fichier GED non PDF", async () => {
    const l = await intra.intraProcessRow(adapter(), params("PARC FLEURI LE CONCERTO 1", "999"));
    assert.deepStrictEqual([l.status, l.cabinet, l.imm, l.lotUrl], [S.LOT_INTROUVABLE, "SICOV", "5004", ""]);
    assert.ok(/immeuble trouvé \(SICOV imm 5004, 2 lot\(s\)\) mais pas le lot 999/.test(l.detail), l.detail);
    const i = await intra.intraProcessRow(adapter(), params("PARC FLEURI CONCERTO", "732"));
    assert.strictEqual(i.status, S.IMMEUBLE_INTROUVABLE);
    assert.ok(/proches : PARC FLEURI LE CONCERTO/.test(i.detail), i.detail);
    const custom = intra.intraCompileRegex("additif").re; // .doc → enregistré tel quel
    const o = await intra.intraProcessRow(adapter(), params("COUR SAINT-HUBERT", "130", { re: custom }));
    assert.deepStrictEqual([o.status, o.file], ["OK", "COUR SAINT-HUBERT_lot130_2015-01-01.doc"]);
  });

  const anv = (res) => intra.intraSaisieValues(res, { DS_COLS: ds.DS_COLS, dsRecordValues: ds.dsRecordValues, validity: diag.diagExpectedValidity });
  await check("analyse du PDF (moteur Analyse DPE) : fournisseur → n° FU, n° ADEME, dates, conso, GES", async () => {
    saved.length = 0;
    const res = await intra.intraProcessRow(adapter(), params("PARC FLEURI LE CONCERTO 1", "732", { analyze: true }));
    assert.strictEqual(res.status, "OK");
    const v = anv(res);
    assert.deepStrictEqual(Object.keys(v), [...ds.DS_COLS, intra.INTRA_SUPPLIER_COL]);
    assert.deepStrictEqual(
      ["DPE analyse", "PDF analysé", "N° FU", "Fournisseur DPE", "N° ADEME", "Établi le", "Valable jusqu'au", "Énergie primaire (kWh/m²/an)", "Classe énergie", "GES (kg CO2/m²/an)", "Classe GES"].map((c) => v[c]),
      ["OK", "PARC FLEURI LE CONCERTO 1_lot732_2013-07-17.pdf", "0528825", "EXPERT IMMOBILIER", "1259V1000732X", "05/06/2012", "04/06/2022", "105", "C", "25", "D"]);
    assert.strictEqual(saved.length, 1, "enregistré puis analysé");
    // DPE de 2007 : pas de n° ADEME, cabinet absent de la liste → à vérifier, nom du cabinet conservé.
    const b = anv(await intra.intraProcessRow(adapter(), params("RESIDENCE BEAUSEJOUR", "543", { analyze: true })));
    assert.deepStrictEqual([b["DPE analyse"], b["N° FU"], b["Fournisseur DPE"], b["N° ADEME"]], ["À VÉRIFIER", "", "DIAGNOSTICS DU NORD", "non trouvé"]);
    assert.ok(/fournisseur absent de la liste : « DIAGNOSTICS DU NORD »/.test(b["Remarques DPE"]), b["Remarques DPE"]);
    const d = anv(await intra.intraProcessRow(adapter(), params("COUR SAINT-HUBERT", "130", { analyze: true })));
    assert.deepStrictEqual([d["N° FU"], d["Fournisseur DPE"], d["N° ADEME"]], ["0394076x", "DIAG 75", "1159V1000130B"]);
  });

  await check("repérage seul + analyse : PDF lu en mémoire, jamais enregistré", async () => {
    saved.length = 0;
    const n0 = srv.stats.flower.length;
    const res = await intra.intraProcessRow(adapter(), params("RESIDENCE CHATEAUBRIAND", "278", { download: false, analyze: true }));
    const v = anv(res);
    assert.deepStrictEqual([res.status, res.file, v["PDF analysé"], v["N° FU"], v["N° ADEME"]], ["OK", "", "MARTIN_DPE.pdf", "0448957", "1559V1000278C"]);
    assert.strictEqual(saved.length, 0);
    assert.strictEqual(srv.stats.flower.length, n0 + 1, "un seul GET FlowerDocs");
  });

  await check("saisie : données Intralia + fournisseur / ADEME du PDF ; un échec d'analyse ne change pas DPE_Statut", async () => {
    const boom = await intra.intraProcessRow(adapter({ analyze: async () => { throw new Error("PDF illisible"); } }), params("PARC FLEURI LE CONCERTO 1", "732", { analyze: true }));
    const v = anv(boom);
    assert.deepStrictEqual([boom.status, v["DPE analyse"], v["Établi le"], v["Énergie primaire (kWh/m²/an)"], v["Classe énergie"], v["N° FU"]], ["OK", "À VÉRIFIER", "05/06/2012", "105", "C", ""]);
    assert.ok(/analyse du PDF impossible \(PDF illisible\)/.test(v["Remarques DPE"]), v["Remarques DPE"]);
    // Intralia prioritaire : « 180.5 » saisi dans Intralia → 180,5 (le PDF dit 180).
    const m = anv(await intra.intraProcessRow(adapter(), params("RESIDENCE CHATEAUBRIAND", "278", { analyze: true })));
    assert.deepStrictEqual([m["DPE analyse"], m["Énergie primaire (kWh/m²/an)"], m["N° FU"], m["Valable jusqu'au"], m["Remarques DPE"]], ["OK", "180,5", "0448957", "31/12/2022", ""],
      "validité calculée comme SIGEO (DPE 2013-2017 → 31/12/2022)");
    // Champ vide dans Intralia → valeur du PDF, signalée.
    const e = anv(await intra.intraProcessRow(adapter(), params("COEUR DE VILLE", "5", { analyze: true })));
    assert.deepStrictEqual([e["DPE analyse"], e["Établi le"], e["Énergie primaire (kWh/m²/an)"], e["N° FU"]], ["À VÉRIFIER", "04/04/2019", "95", "0426828"]);
    assert.ok(/date du DPE vide dans Intralia : valeur du PDF/.test(e["Remarques DPE"]), e["Remarques DPE"]);
    // Date du PDF ≠ Intralia → à vérifier ; DPE récent : coûts d'Intralia, remarque ADEME conservée.
    const fake = (rec) => adapter({ analyze: async () => Object.assign(ds.dsNewRecord(), rec) });
    const d = anv(await intra.intraProcessRow(fake({ etabli: { y: 2013, m: 1, d: 2 }, remarques: ["n° ADEME non trouvé", "valeurs lues par OCR : à contrôler"], supplier: { entry: { name: "ACADE", code: "0448957" } } }),
      params("PARC FLEURI LE CONCERTO 1", "732", { analyze: true })));
    assert.ok(/PDF établi le 02\/01\/2013 ≠ Intralia 05\/06\/2012/.test(d["Remarques DPE"]), d["Remarques DPE"]);
    assert.ok(!/ADEME|OCR/.test(d["Remarques DPE"]), "DPE ancien : pas de remarque ADEME / OCR des valeurs (prises dans Intralia)");
    const n = intra.intraSaisieValues({ status: "OK", chosen: { name: "x.pdf" }, file: "x.pdf", fields: { DATEDIAG: "2023-04-17", DPE1: "221", libdpe1: "D", GES1: "6", libges1: "B", MTDIAGMIN: "479", MTDIAGMAX: "649" },
      analysis: Object.assign(ds.dsNewRecord(), { ademe: "2321E1261645S", remarques: ["n° ADEME lu par OCR : à contrôler"], supplier: { entry: { name: "ACADE", code: "448957" } } }) },
    { DS_COLS: ds.DS_COLS, dsRecordValues: ds.dsRecordValues, validity: diag.diagExpectedValidity });
    assert.deepStrictEqual([n["Établi le"], n["Valable jusqu'au"], n["Coût min (€/an)"], n["Coût max (€/an)"], n["N° ADEME"], n["DPE analyse"], n["Remarques DPE"]],
      ["17/04/2023", "16/04/2033", "479", "649", "2321E1261645S", "À VÉRIFIER", "n° ADEME lu par OCR : à contrôler"]);
    // Sans DPE / erreurs.
    const none = anv(await intra.intraProcessRow(adapter(), params("OPERA", "9", { analyze: true })));
    assert.deepStrictEqual([none["DPE analyse"], none["Remarques DPE"]], ["Pas de PDF", "AUCUN_DPE_GED — 0 document(s) GED, aucun DPE"]);
    assert.strictEqual(anv(await intra.intraProcessRow(adapter(), params("LES ACACIAS", "3", { analyze: true })))["DPE analyse"], "Pas de PDF");
    srv.setFlowerLogged(false);
    try {
      const fl = await intra.intraProcessRow(adapter(), params("PARC FLEURI LE CONCERTO 1", "732", { download: false, analyze: true }));
      assert.deepStrictEqual([fl.status, anv(fl)["DPE analyse"]], [S.FLOWERDOCS_NON_CONNECTE, "ÉCHEC"], "repérage + analyse : FlowerDocs déconnecté = arrêt");
    } finally { srv.setFlowerLogged(true); }
    assert.deepStrictEqual(intra.intraSaisieCols(ds.DS_COLS).slice(0, 4), ["DPE analyse", "PDF analysé", "N° FU", "Fournisseur DPE"]);
    assert.strictEqual(intra.intraNormalizeConfig({}).out, "saisie");
  });

  await check("performance : 36 000 lignes parsées par regex", async () => {
    let extra = "";
    for (let i = 0; i < 36000; i++) {
      const h = `../gerance/diaglot.php?id_societe=${20 + (i % 9)}&amp;id_immeuble=${10000 + Math.floor(i / 12)}&amp;id_lot=${i % 12 + 1}`;
      extra += `<tr>${["CAB " + (i % 9), 10000 + Math.floor(i / 12), "IMMEUBLE N° " + Math.floor(i / 12), "1 RUE X", "75000", "PARIS", i % 12 + 1, "Appartement", "NC"]
        .map((v) => `<td><a href="${h}">${v}</a></td>`).join("")}</tr>\n`;
    }
    const html = mock.diagHtml("19", extra);
    const t0 = Date.now();
    const r = await runOp(tab, { op: "parseIndex", html });
    const idx = intra.intraBuildIndex(r.rows);
    const ms = Date.now() - t0;
    console.log(`   ${(html.length / 1e6).toFixed(1)} Mo, ${r.rows.length} lignes, index ${idx.count} lots en ${ms} ms`);
    assert.ok(r.rows.length > 36000);
    assert.ok(ms < 15000, ms + " ms");
  });

  await check("lecture seule : aucun POST envoyé à Intralia", async () => {
    assert.deepStrictEqual(srv.stats.posts, []);
  });

  await srv.close();
  console.log(failures ? `\n${failures} test(s) en échec.` : "\nTous les tests passent.");
  process.exit(failures ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
