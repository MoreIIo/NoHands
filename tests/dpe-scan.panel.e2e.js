// Test d'intégration du panneau : onglet Extraction, type « Analyse DPE ».
// Le vrai sidepanel.html est servi en HTTP local (les workers pdf.js et
// Tesseract.js ne démarrent pas en file://) ; chrome.* est simulé. Les DPE
// fictifs sont générés par tests/fixtures/make_dpe.py (un PDF scanné compris :
// l'OCR passe par le vrai Tesseract.js embarqué, compter 1 à 3 minutes).
//   NODE_PATH=$(npm root -g) node tests/dpe-scan.panel.e2e.js
const fs = require("fs");
const os = require("os");
const http = require("http");
const path = require("path");
const assert = require("assert");
const { execFileSync } = require("child_process");
const { chromium } = require("playwright");

const ROOT = path.resolve(__dirname, "..");
const FIX = path.join(os.tmpdir(), "nohands-dpe-fixtures");
if (!fs.existsSync(path.join(FIX, "IM0393664_11_DPE_2025-03-15.pdf"))) {
  execFileSync("python3", [path.join(__dirname, "fixtures", "make_dpe.py"), FIX], { stdio: "inherit" });
}

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
    tabs: { get: async (id) => ({ id, url: "about:blank" }), query: async () => [], sendMessage: async () => ({}),
            onUpdated: { addListener() {}, removeListener() {} }, onRemoved: { addListener() {} } },
    scripting: { executeScript: async () => [] },
    permissions: { contains: async () => true, request: async () => true },
    windows: { update: async () => ({}) }
  };
})();`;

const HEAD = ["Immeuble", "IM", "lot", "IM_format", "ID interne", "URL lot", "Résultat lot", "DPE présent", "Nb DPE", "DPE nom SIGEO", "DPE date", "DPE section", "DPE fichier", "DPE statut"];
const ROWS = [
  HEAD,
  ["COUR DES M", "IM0393664", "10", "393664", "1", "", "OK — id 1", "OUI", "1", "DPE 2022", "19/10/2022", "Diagnostics >> EPC", "IM0393664_10_DPE_2022-10-19.pdf", "OK & téléchargé"],
  ["RESIDENCE B", "IM0393664", "11", "393664", "2", "", "OK — id 2", "OUI", "1", "DPE", "15/03/2025", "Diagnostics >> EPC", "", "OK trouvé (non téléchargé)"],
  ["RESIDENCE C", "IM0393665", "3", "393665", "3", "", "OK — id 3", "OUI", "1", "DPE 2014", "12/05/2014", "Diagnostics >> EPC", "IM0393665_3_DPE_2014-05-12.pdf", "OK & téléchargé"],
  ["SANS DPE", "IM0393666", "4", "393666", "4", "", "OK — id 4", "NON", "0", "", "", "", "", "KO pas trouvé"],
  ["RESIDENCE H", "IM0393667", "5", "393667", "5", "", "OK — id 5", "OUI", "1", "DPE", "02/02/2023", "Diagnostics >> EPC", "IM0393667_5_DPE_2023-02-02.pdf", "OK & téléchargé"],
  ["MANQUANT", "IM0393668", "6", "393668", "6", "", "OK — id 6", "OUI", "1", "DPE", "01/01/2020", "Diagnostics >> EPC", "IM0393668_6_DPE_2020-01-01.pdf", "OK & téléchargé"]
];

(async () => {
  const { srv, origin } = await serve();
  const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/usr/bin/google-chrome", headless: true });
  const ctx = await browser.newContext({ acceptDownloads: true });
  const panel = await ctx.newPage();
  const errors = [];
  panel.on("pageerror", (e) => { errors.push(e.message); console.log("[panel] " + e.message); });
  panel.on("console", (m) => { if (m.type() === "error") console.log("[console] " + m.text()); });
  await panel.addInitScript(CHROME_STUB);
  await panel.goto(origin + "/sidepanel.html");
  await panel.waitForFunction(() => typeof runDsExtraction === "function" && $("dsSupCount").textContent !== "…");

  // Liste fournisseurs par défaut (celle de Maëlle)
  const sup0 = await panel.evaluate(() => ({ n: $("dsSupCount").textContent, warn: $("dsSupWarn").textContent, first: $("dsSuppliers").value.split("\n")[0] }));
  console.log("liste :", JSON.stringify(sup0));
  assert.strictEqual(sup0.n, "27 fournisseurs");
  assert.strictEqual(sup0.first, "EXPERT IMMOBILIER : 0528825");
  assert.ok(/0394076x/.test(sup0.warn), sup0.warn);

  // Données : sortie de « DPE SIGEO »
  await panel.evaluate((rows) => { state.headerMode = "yes"; state.rows = rows; renderColumns(); showTab("extraction"); }, ROWS);
  await panel.check('input[name="extractKind"][value="dpeScan"]');
  const ui = await panel.evaluate(() => ({ ds: !$("dsExtractBox").hidden, std: !$("stdExtractBox").hidden, dpe: !$("dpeExtractBox").hidden, cfg: getDsConfig(), est: $("extractEstimate").textContent }));
  console.log("UI :", JSON.stringify(ui));
  assert.ok(ui.ds && !ui.std && !ui.dpe, "seul le bloc Analyse DPE est visible");
  assert.deepStrictEqual([ui.cfg.fileCol, ui.cfg.imCol, ui.cfg.lotCol, ui.cfg.ocr, ui.cfg.append], ["DPE fichier", "IM", "lot", "auto", true]);
  assert.ok(/choisis d'abord le dossier/.test(ui.est), ui.est);

  // Choix du dossier (input webkitdirectory)
  await panel.setInputFiles("#dsFolderInput", FIX);
  const info = await panel.evaluate(() => ({ info: $("dsFilesInfo").textContent, est: $("extractEstimate").textContent }));
  console.log("dossier :", JSON.stringify(info));
  assert.ok(/^5 DPE dans « nohands-dpe-fixtures » \(5 PDF\) — 1 autre\(s\) fichier\(s\) ignoré\(s\)\.$/.test(info.info), info.info);
  assert.ok(/^5 DPE choisis/.test(info.est), info.est);

  // Tester (ligne 2, rien n'est écrit)
  await panel.fill("#startRow", "2");
  await panel.click("#dsTestBtn");
  await panel.waitForFunction(() => !$("dsTestBtn").disabled && /^[✓⚠✗]/.test($("dsTestResult").textContent), null, { timeout: 60000 });
  const testTxt = await panel.textContent("#dsTestResult");
  console.log("test :", testTxt);
  assert.ok(testTxt.startsWith("✓ IM0393664_10_DPE_2022-10-19.pdf (ligne 2)") && testTxt.includes("N° FU : 0448957") && testTxt.includes("Énergie primaire (kWh/m²/an) : 208"), testTxt);
  assert.strictEqual(await panel.evaluate(() => state.rows[0].length), HEAD.length, "le test n'écrit rien");

  // Exécution complète
  await panel.click("#startBtn");
  await panel.waitForFunction(() => !isRunning && /Terminé/.test($("log").textContent), null, { timeout: 400000, polling: 1000 });
  const out = await panel.evaluate(() => ({ rows: state.rows, log: $("log").innerText, recap: $("resultSummary").innerText, unk: !$("dsSupAddUnknown").hidden }));
  console.log("--- journal ---\n" + out.log + "\n--- récap ---\n" + out.recap);
  const H = out.rows[0];
  const col = (n) => { const i = H.indexOf(n); assert.ok(i >= HEAD.length, "colonne " + n); return i; };
  const get = (r, n) => out.rows[r][col(n)];
  const pick = (r, names) => names.map((n) => get(r, n));
  const MAIN = ["DPE analyse", "N° FU", "Établi le", "Valable jusqu'au", "Énergie primaire (kWh/m²/an)", "Classe énergie", "GES (kg CO2/m²/an)", "Classe GES", "N° ADEME"];
  assert.deepStrictEqual(pick(1, MAIN), ["OK", "0448957", "19/10/2022", "18/10/2032", "208", "D", "19", "C", "2275E1234567X"]);
  assert.deepStrictEqual(pick(1, ["Fournisseur DPE", "Fournisseur (liste)", "Type diagnostic", "Coût min (€/an)", "Coût max (€/an)", "Surface (m²)", "Adresse DPE", "Diagnostiqueur", "PDF analysé"]),
    ["ACADE", "ACADE", "Performance Energetique", "1230", "1700", "45,2", "12 rue Lamarck 75018 PARIS", "Jean DUPONT", "IM0393664_10_DPE_2022-10-19.pdf"]);
  // Scan (rattaché par IM + lot), OCR Tesseract.js, SOCOTEC ≈ SOCOTECH, petite surface
  assert.deepStrictEqual(pick(2, ["DPE analyse", "N° FU", "Établi le", "Valable jusqu'au", "Énergie primaire (kWh/m²/an)", "Classe énergie", "GES (kg CO2/m²/an)", "Classe GES", "Surface (m²)"]),
    ["À VÉRIFIER", "0048730", "15/03/2025", "14/03/2035", "340", "E", "21", "C", "28,5"]);
  assert.strictEqual(get(2, "PDF analysé"), "IM0393664_11_DPE_2025-03-15.pdf");
  assert.ok(/^OCR/.test(get(2, "Lecture")), get(2, "Lecture"));
  assert.ok(/rapproché/.test(get(2, "Remarques DPE")), get(2, "Remarques DPE"));
  // Ancien DPE
  assert.deepStrictEqual(pick(3, ["Établi le", "Énergie primaire (kWh/m²/an)", "Classe énergie", "GES (kg CO2/m²/an)", "Classe GES", "N° FU"]),
    ["12/05/2014", "318,99", "E", "74,64", "F", "0394076x"]);
  assert.deepStrictEqual(pick(4, ["DPE analyse", "N° FU", "Remarques DPE"]), ["Pas de PDF", "", ""]);
  assert.deepStrictEqual(pick(5, ["DPE analyse", "Fournisseur DPE", "N° FU"]), ["À VÉRIFIER", "DIAG EXPRESS SUD-OUEST", ""]);
  assert.deepStrictEqual(pick(6, ["DPE analyse", "Remarques DPE"]), ["PDF introuvable", "« IM0393668_6_DPE_2020-01-01.pdf » absent du dossier"]);
  // PDF sans ligne → ajouté en bas
  assert.strictEqual(out.rows.length, 8);
  assert.deepStrictEqual([out.rows[7][1], out.rows[7][2], out.rows[7][12], get(7, "N° FU"), get(7, "Fournisseur (liste)")],
    ["IM0400000", "7", "IM0400000_7_DPE_2024-01-01.pdf", "0425954", "AV ENVIRONNEMENT"]);
  assert.strictEqual(out.rows[1][0], "COUR DES M", "colonnes source intactes");
  assert.ok(/1 OK, 4 à vérifier, 0 échec\(s\), 2 sans PDF/.test(out.log), out.log);
  assert.ok(out.unk, "bouton « Ajouter les fournisseurs inconnus » visible");

  // Fournisseurs inconnus → ajoutés à la liste (numéro à compléter), liste enregistrée
  await panel.click("#dsSupAddUnknown");
  await panel.waitForTimeout(500);
  const sup1 = await panel.evaluate(() => ({ last: $("dsSuppliers").value.split("\n").pop(), saved: window.__store.dsSuppliers, warn: $("dsSupWarn").textContent, open: $("dsSupBox").open }));
  assert.strictEqual(sup1.last, "DIAG EXPRESS SUD-OUEST : ");
  assert.ok(sup1.saved.endsWith("DIAG EXPRESS SUD-OUEST : "), "liste enregistrée");
  assert.ok(/1 nom\(s\) sans numéro : DIAG EXPRESS SUD-OUEST/.test(sup1.warn), sup1.warn);
  // Numéro ajouté à la main → relance sur la ligne 6 (Excel ligne 6 = RESIDENCE H), sans OCR
  await panel.evaluate(() => {
    const t = $("dsSuppliers"); t.value = t.value.replace(/DIAG EXPRESS SUD-OUEST : $/, "DIAG EXPRESS SUD-OUEST : 0612345");
    t.dispatchEvent(new Event("input", { bubbles: true }));
  });
  await panel.fill("#startRow", "6"); await panel.fill("#endRow", "6");
  await panel.selectOption("#dsOcr", "never");
  await panel.click("#startBtn");
  await panel.waitForFunction(() => !isRunning && /Terminé/.test($("log").textContent), null, { timeout: 60000 });
  const r6 = await panel.evaluate(() => { const H = state.rows[0]; return ["DPE analyse", "N° FU", "Fournisseur (liste)"].map((c) => state.rows[5][H.indexOf(c)]); });
  assert.deepStrictEqual(r6, ["OK", "0612345", "DIAG EXPRESS SUD-OUEST"]);
  assert.strictEqual(await panel.evaluate(() => state.rows.length), 8, "pas de nouvel ajout hors plage");

  // Export Excel (bouton du bloc)
  const [dl] = await Promise.all([panel.waitForEvent("download"), panel.click("#dsExportBtn")]);
  const xlsx = path.join(os.tmpdir(), "dpe-export-test.xlsx");
  await dl.saveAs(xlsx);
  const xl = JSON.parse(execFileSync("python3", ["-c", `
import openpyxl, json
ws = openpyxl.load_workbook("${xlsx}").active
print(json.dumps([[c.value for c in r] for r in ws.iter_rows()]))`]).toString());
  const XH = xl[0];
  assert.strictEqual(xl.length, 8);
  assert.strictEqual(xl[1][XH.indexOf("N° FU")], "0448957", "n° FU exporté en texte (zéro conservé)");
  assert.strictEqual(xl[7][XH.indexOf("Fournisseur (liste)")], "AV ENVIRONNEMENT");
  console.log("export :", dl.suggestedFilename(), XH.length, "colonnes");

  // Profil : réglages sauvegardés / restaurés
  const prof = await panel.evaluate(() => {
    $("dsAppend").checked = false;
    const c = buildProfileConfig(); applyDsConfig({}); const def = getDsConfig(); applyProfileConfig(c);
    return { kind: c.extraction.kind, after: extractKind(), saved: c.extraction.ds, def: def.ocr, restored: getDsConfig() };
  });
  assert.deepStrictEqual([prof.kind, prof.after, prof.def], ["dpeScan", "dpeScan", "auto"]);
  assert.deepStrictEqual(prof.restored, { fileCol: "DPE fichier", imCol: "IM", lotCol: "lot", append: false, ocr: "never" });

  // Autre sélection (1 seul PDF) : les lignes déjà analysées sans PDF sont conservées
  await panel.setInputFiles("#dsFilesInput", path.join(FIX, "IM0393664_10_DPE_2022-10-19.pdf"));
  await panel.fill("#startRow", "2"); await panel.fill("#endRow", "");
  await panel.click("#startBtn");
  await panel.waitForFunction(() => !isRunning && /Terminé/.test($("log").textContent), null, { timeout: 60000 });
  const keep = await panel.evaluate(() => { const H = state.rows[0], c = H.indexOf("DPE analyse"), f = H.indexOf("N° FU");
    return { st: state.rows.slice(1).map((r) => r[c]), fu3: state.rows[2][f], n: state.rows.length, log: $("log").innerText }; });
  assert.deepStrictEqual(keep.st, ["OK", "À VÉRIFIER", "À VÉRIFIER", "Pas de PDF", "OK", "PDF introuvable", "À VÉRIFIER"], JSON.stringify(keep.st));
  assert.strictEqual(keep.fu3, "0048730", "résultat du scan conservé");
  assert.strictEqual(keep.n, 8);
  assert.ok(/analyse précédente conservée \(À VÉRIFIER\)/.test(keep.log), keep.log);

  // Sans tableau chargé : une ligne par PDF
  await panel.setInputFiles("#dsFolderInput", FIX);
  await panel.evaluate(() => { state.rows = []; renderColumns(); $("startRow").value = ""; $("endRow").value = ""; });
  await panel.click("#startBtn");
  await panel.waitForFunction(() => !isRunning && /Terminé/.test($("log").textContent), null, { timeout: 60000 });
  const fresh = await panel.evaluate(() => ({ rows: state.rows, name: state.originalFileName, cfg: getDsConfig() }));
  assert.deepStrictEqual(fresh.rows.slice(0, 3).map((r) => r.slice(0, 4)), [
    ["DPE fichier", "IM", "lot", "DPE analyse"],
    ["IM0393664_10_DPE_2022-10-19.pdf", "IM0393664", "10", "OK"],
    ["IM0393664_11_DPE_2025-03-15.pdf", "IM0393664", "11", "ÉCHEC"]
  ]);
  assert.ok(/aucun texte exploitable/.test(fresh.rows[2][fresh.rows[0].indexOf("Remarques DPE")]), "scan sans OCR → ÉCHEC expliqué");
  assert.strictEqual(fresh.rows.length, 6);
  assert.ok(/^DPE_analyse_\d{4}-\d{2}-\d{2}\.xlsx$/.test(fresh.name), fresh.name);

  assert.deepStrictEqual(errors, [], "erreurs JS du panneau");
  console.log("\n✓ onglet Extraction (Analyse DPE) OK");
  await browser.close();
  srv.close();
})().catch((e) => { console.error(e); process.exit(1); });
