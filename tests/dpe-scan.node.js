// Tests du moteur « Analyse DPE » (dpe-scan.js) sous Node, sur des DPE fictifs.
// Couche texte : lib/pdf.min.js (pdf.js du panneau). OCR : tesseract en ligne
// de commande (même format TSV que Tesseract.js), pages rendues par pdftoppm.
//   node tests/dpe-scan.node.js [dossier des PDF]
// Sans dossier : les DPE fictifs sont générés par tests/fixtures/make_dpe.py
// dans un dossier temporaire et les valeurs attendues sont vérifiées.
const fs = require("fs");
const path = require("path");
const os = require("os");
const assert = require("assert");
const { execFileSync } = require("child_process");
const ds = require("../dpe-scan.js");

const ROOT = path.resolve(__dirname, "..");
const DIR = process.argv[2] || (() => {
  const d = path.join(os.tmpdir(), "nohands-dpe-fixtures");
  if (!fs.existsSync(path.join(d, "IM0393664_11_DPE_2025-03-15.pdf"))) {
    execFileSync("python3", [path.join(__dirname, "fixtures", "make_dpe.py"), d], { stdio: "inherit" });
  }
  return d;
})();
const pdfjs = require(path.join(ROOT, "lib/pdf.min.js"));
pdfjs.GlobalWorkerOptions.workerSrc = path.join(ROOT, "lib/pdf.worker.min.js");

function ocrIo(file, isPdf) {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "dpeocr-"));
  return {
    stopped: () => false,
    async ocr(i, zone) {
      let img;
      if (isPdf) {
        const base = path.join(tmp, `p${i}${zone ? "z" : ""}`);
        execFileSync("pdftoppm", ["-r", zone ? "300" : "200", "-gray", "-png", "-f", String(i + 1), "-l", String(i + 1), "-singlefile", file, base]);
        img = base + ".png";
      } else img = file;
      if (zone) {
        execFileSync("python3", ["-c", `
from PIL import Image, ImageOps
im = Image.open("${img}").convert("L"); w, h = im.size
im = im.crop((0, 0, w, int(h * 0.55))); im = ImageOps.autocontrast(im, cutoff=1)
im = im.point(lambda v: 0 if v < 200 else 255); im.save("${img}")`]);
      }
      const tsv = execFileSync("tesseract", [img, "-", "-l", "fra", "--psm", zone ? "6" : "3", "tsv"], { stdio: ["ignore", "pipe", "ignore"] }).toString();
      const words = ds.dsParseTsv(tsv);
      return { text: ds.dsWordsText(words), words, blank: !words.length };
    }
  };
}

async function analyze(file, ocr, suppliers) {
  const isPdf = /\.pdf$/i.test(file);
  let pages = [""];
  if (isPdf) {
    const r = await ds.dsReadPdf(pdfjs, new Uint8Array(fs.readFileSync(file)));
    pages = r.pages;
    r.doc.destroy();
  }
  return ds.dsAnalyzeDoc(pages, ocrIo(file, isPdf), { ocr, suppliers });
}

function unitTests() {
  // Lettres (mêmes cas que l'outil Rust)
  const d21 = ds.dsDate(5, 10, 2021);
  assert.deepStrictEqual([ds.dsLetters("2021", 208, 19, 71.8, d21).ep, ds.dsLetters("2021", 208, 19, 71.8, d21).ges], ["D", "C"]);
  assert.strictEqual(ds.dsLetters("2021", 86, 17, 41.62, d21).ep, "C");
  assert.strictEqual(ds.dsLetters("2021", 91, 5, 42.73, d21).ep, "B");
  const r = ds.dsLetters("2021", 690, 27, 13.36, ds.dsDate(1, 9, 2026));
  assert.deepStrictEqual([r.ep, r.ges], ["G", "C"]);
  const o = ds.dsLetters("ancien", 318.99, 74.64, null, null);
  assert.deepStrictEqual([o.ep, o.ges], ["E", "F"]);
  // N° ADEME
  assert.strictEqual(ds.dsFindAdemeStrict("n° : 2121e0449628f\n"), "2121E0449628F");
  assert.strictEqual(ds.dsRepairAdeme("2121É0449628F"), "2121E0449628F");
  assert.strictEqual(ds.dsRepairAdeme("2121 E O449628 F"), "2121E0449628F");
  // Nombres, coûts dans le désordre du flux PDF
  assert.strictEqual(ds.dsParseNum("14 954"), 14954);
  assert.strictEqual(ds.dsParseNum("318,99"), 318.99);
  assert.strictEqual(ds.dsParseNum("1.029"), 1029);
  const rec = ds.dsExtract({ dpeText: ds.dsNorm("ce logement emet 860 kg\npar anentre et1430 € 1980 €\nsurface habitable : 50 m²"), allText: "" });
  assert.deepStrictEqual([rec.coutMin, rec.coutMax], [1430, 1980]);
  // Étiquette lue par OCR
  const p = ds.dsLabelPairsFromText("208 |19*\nkkwh/m/an |kgco2/m°/an");
  assert.deepStrictEqual([p[0].ep, p[0].ges], [208, 19]);
  // Liste fournisseurs
  const L = ds.dsParseSuppliers("Clé;Code\nACADE -> 0448957\nARTHEMIS CÔTE D'OR ; FU002\nSOCOTEC\t0048730\nARTHEMIS => 0596024\nVIDE : \nACADE : 9999999");
  assert.deepStrictEqual(L.entries.map((e) => [e.key, e.code]), [["ACADE", "0448957"], ["ARTHEMIS COTE D OR", "FU002"], ["SOCOTEC", "0048730"], ["ARTHEMIS", "0596024"]]);
  assert.deepStrictEqual(L.empty, ["VIDE"]);
  assert.ok(L.warnings.some((w) => /deux numéros/.test(w)), L.warnings.join(" | "));
  const f = (ctx) => { const x = ds.dsFindSupplier(L.entries, ctx); return x ? [x.entry.code, x.how, x.fuzzy] : null; };
  assert.deepStrictEqual(f({ societe: "ACADE", zone: "" }), ["0448957", "société", false]);
  assert.deepStrictEqual(f({ societe: "Arthemis Côte d'Or", zone: "" }), ["FU002", "société", false]);
  assert.strictEqual(f({ societe: "ACADEMIE", zone: "" }), null);
  assert.deepStrictEqual(f({ societe: "X", zone: "email : jean@acade.fr" }), ["0448957", "bloc diagnostiqueur", false]);
  const D = ds.dsParseSuppliers(ds.DS_DEFAULT_SUPPLIERS).entries;
  const g = (ctx) => { const x = ds.dsFindSupplier(D, ctx); return x ? [x.entry.name, x.fuzzy] : null; };
  assert.deepStrictEqual(g({ societe: "Diagamter Paris 15", zone: "" }), ["DIAGMTER", true]);
  assert.deepStrictEqual(g({ societe: "Cabinet", zone: "diagnostiqueur : Frédérique ROUSSEL" }), ["ROUSSEL Frédérique", false]);
  assert.strictEqual(g({ societe: "DIAG EXPRESS", zone: "DIAG EXPRESS" }), null, "pas de DIAG3 approximatif");
  // Rattachement PDF ↔ lignes
  assert.deepStrictEqual(ds.dsParseFileName("IM0393664_10_DPE_2022-10-19 (1).pdf"), { im: "393664", lot: "10", date: "2022-10-19" });
  assert.deepStrictEqual(ds.dsParseFileName("0393664_LT000204_DPE_sans-date.pdf"), { im: "393664", lot: "204", date: "" });
  const rows = [["Immeuble", "IM", "lot", "DPE fichier"],
    ["A", "IM0393664", "10", "IM0393664_10_DPE_2022-10-19 (1).pdf"],
    ["B", "0393664", "LT000011", ""],
    ["C", "IM0393665", "3", ""],
    ["D", "IM0393666", "4", "absent.pdf"]];
  const files = [{ name: "IM0393664_10_DPE_2022-10-19.pdf" }, { name: "IM0393664_10_DPE_2022-10-19 (1).pdf" },
    { name: "IM0393664_11_DPE_2021-01-01.pdf" }, { name: "IM0393664_11_DPE_2025-03-15.pdf" }, { name: "IM0400000_7_DPE_2024-01-01.pdf" }];
  const a = ds.dsAssignFiles(rows, [1, 2, 3, 4], { file: 3, im: 1, lot: 2 }, files);
  assert.deepStrictEqual([...a.byRow.entries()], [[1, 1], [2, 3]]);
  assert.deepStrictEqual([a.via.get(1), a.via.get(2)], ["fichier", "IM + lot"]);
  assert.deepStrictEqual(a.dupes.sort(), [0, 2]);
  assert.deepStrictEqual(a.unmatched, [4]);
  console.log("✓ tests unitaires");
}

(async () => {
  unitTests();
  const sup = ds.dsParseSuppliers(ds.DS_DEFAULT_SUPPLIERS);
  console.log("liste :", sup.entries.length, "fournisseurs ;", sup.warnings.join(" | "));
  const files = fs.readdirSync(DIR).filter((f) => ds.DS_FILE_RE.test(f)).sort();
  const out = {};
  for (const f of files) {
    const t0 = Date.now();
    const rec = await analyze(path.join(DIR, f), process.env.OCR || "auto", sup.entries);
    const v = ds.dsRecordValues(rec, f);
    out[f] = Object.fromEntries(ds.DS_COLS.map((c, i) => [c, v[i]]));
    console.log(`\n=== ${f} (${Date.now() - t0} ms)`);
    ds.DS_COLS.forEach((c, i) => { if (v[i]) console.log(`  ${c.padEnd(34)} ${v[i]}`); });
  }
  if (process.argv[2]) return;
  const g = (f, c) => out[f][c];
  // DPE 2021 texte (ACADE) : valeurs calculées total ÷ surface, sans OCR en mode auto
  const A = "IM0393664_10_DPE_2022-10-19.pdf";
  // Colonnes : uniquement ce qui sert à la saisie SIGEO
  assert.deepStrictEqual(ds.DS_COLS, ["DPE analyse", "PDF analysé", "N° FU", "Établi le", "Valable jusqu'au", "N° ADEME",
    "Énergie primaire (kWh/m²/an)", "Classe énergie", "GES (kg CO2/m²/an)", "Classe GES", "Coût min (€/an)", "Coût max (€/an)",
    "Prix indexés au", "Remarques DPE"]);
  assert.deepStrictEqual(ds.DS_COLS.map((c) => out[A][c]),
    ["OK", A, "0448957", "19/10/2022", "18/10/2032", "2275E1234567X", "208", "D", "19", "C", "1230", "1700", "01/01/2021", ""]);
  // DPE 2021 scanné (SOCOTEC ≈ SOCOTECH) : OCR, petite surface (E au lieu de F)
  const B = "IM0393664_11_DPE_2025-03-15.pdf";
  assert.strictEqual(g(B, "N° FU"), "0048730");
  assert.deepStrictEqual([g(B, "Établi le"), g(B, "Valable jusqu'au"), g(B, "Énergie primaire (kWh/m²/an)"), g(B, "GES (kg CO2/m²/an)"),
    g(B, "Classe énergie"), g(B, "Classe GES"), g(B, "Coût min (€/an)"), g(B, "Coût max (€/an)")],
  ["15/03/2025", "14/03/2035", "340", "21", "E", "C", "1080", "1470"]);
  assert.strictEqual(g(B, "N° ADEME"), "2575E0456789B");
  assert.strictEqual(g(B, "DPE analyse"), "À VÉRIFIER");
  assert.ok(/rapproché : « .+ » ≈ « SOCOTECH »/.test(g(B, "Remarques DPE")), g(B, "Remarques DPE"));
  // Ancien DPE 2014
  const C = "IM0393665_3_DPE_2014-05-12.pdf";
  assert.deepStrictEqual([g(C, "Établi le"), g(C, "Valable jusqu'au"), g(C, "Énergie primaire (kWh/m²/an)"), g(C, "GES (kg CO2/m²/an)"),
    g(C, "Classe énergie"), g(C, "Classe GES"), g(C, "N° FU"), g(C, "N° ADEME")],
  ["12/05/2014", "11/05/2024", "318,99", "74,64", "E", "F", "0394076x", "1475V1001234X"]);
  // DPE d'avant 07/2021 : pas d'estimation dans SIGEO → coûts et date des prix non écrits
  assert.deepStrictEqual([g(C, "Coût min (€/an)"), g(C, "Coût max (€/an)"), g(C, "Prix indexés au")], ["", "", ""]);
  assert.ok(!/coûts|montant annuel|indexation/.test(g(C, "Remarques DPE")), g(C, "Remarques DPE"));
  assert.ok(/validité légale jusqu'au 31\/12\/2022/.test(g(C, "Remarques DPE")));
  // Fournisseur inconnu
  const E = "IM0393667_5_DPE_2023-02-02.pdf";
  assert.strictEqual(g(E, "N° FU"), "");
  assert.strictEqual(g(E, "DPE analyse"), "À VÉRIFIER");
  assert.ok(/absent de la liste : « DIAG EXPRESS SUD-OUEST »/.test(g(E, "Remarques DPE")), g(E, "Remarques DPE"));
  assert.strictEqual(g(E, "Classe énergie"), "C");
  // N° ADEME absent : « NC » (évaluation thermique / non émis ADEME) ou « non trouvé »
  const ix = ds.DS_COLS.indexOf("N° ADEME");
  const r0 = Object.assign(ds.dsNewRecord(), { etabli: { y: 2025, m: 4, d: 30 } });
  assert.strictEqual(ds.dsRecordValues(r0, "x.pdf")[ix], "non trouvé");
  assert.strictEqual(ds.dsRecordValues(Object.assign(r0, { nonOfficiel: true }), "x.pdf")[ix], "NC");
  console.log("\n✓ moteur Analyse DPE OK");
})().catch((e) => { console.error(e); process.exit(1); });
