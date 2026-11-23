// Tests unitaires (logique pure) de l'étape « SIGEO : saisir un DPE ».
//   node tests/sigeo-diag.node.js
const assert = require("assert");
const D = require("../sigeo-diag.js");

let failures = 0;
function check(name, fn) {
  try { fn(); console.log("✓ " + name); }
  catch (e) { failures++; console.log("✗ " + name + "\n   " + (e.stack || e).toString().split("\n").slice(0, 3).join("\n   ")); }
}

check("dates : texte, série Excel, ISO, invalides", () => {
  assert.deepStrictEqual(D.diagDate("15/1/21"), { ok: true, value: "15/01/2021" });
  assert.strictEqual(D.diagDate("  15.03.2024 ").value, "15/03/2024");
  assert.strictEqual(D.diagDate("2024-03-15").value, "15/03/2024");
  assert.strictEqual(D.diagDate("2024-03-15T00:00:00Z").value, "15/03/2024");
  assert.strictEqual(D.diagDate("", { t: "n", v: 44211 }).value, "15/01/2021", "cellule numérique prioritaire");
  assert.strictEqual(D.diagDate("1/15/21", { t: "n", v: 44211 }).value, "15/01/2021", "m/d/yy de SheetJS ignoré");
  assert.strictEqual(D.diagDate("44211").value, "15/01/2021", "série tapée en texte");
  assert.strictEqual(D.diagDate("").value, "");
  assert.ok(!D.diagDate("31/02/2024").ok);
  assert.ok(!D.diagDate("13/13/2024").ok);
  assert.ok(!D.diagDate("hier").ok);
  assert.ok(D.diagDateEq("1/2/2024", "01/02/2024"));
});

check("validité calculée (règles SIGEO)", () => {
  assert.strictEqual(D.diagExpectedValidity("15/03/2012"), "14/03/2022");
  assert.strictEqual(D.diagExpectedValidity("01/06/2015"), "31/12/2022");
  assert.strictEqual(D.diagExpectedValidity("15/01/2018"), "31/12/2024");
  assert.strictEqual(D.diagExpectedValidity("30/06/2021"), "31/12/2024");
  assert.strictEqual(D.diagExpectedValidity("01/07/2021"), "30/06/2031");
  assert.strictEqual(D.diagExpectedValidity("29/02/2024"), "28/02/2034");
  assert.ok(D.diagIsNewDpe("01/07/2021") && !D.diagIsNewDpe("30/06/2021"));
});

check("nombres : unités, séparateurs, maxlength", () => {
  assert.strictEqual(D.diagNum("1 234,5 kWh/m2/an", null, 6).value, "1234,5");
  assert.strictEqual(D.diagNum("320", null, 6).value, "320");
  assert.strictEqual(D.diagNum("320.456", null, 6).value, "320,46");
  assert.strictEqual(D.diagNum("1.234,5", null, 6).value, "1234,5");
  assert.strictEqual(D.diagNum("", { t: "n", v: 9 }, 6).value, "9");
  assert.strictEqual(D.diagNum("1,240", { t: "n", v: 1240 }, 5).value, "1240", "cellule numérique prioritaire");
  const r = D.diagNum("1240,75", null, 5);
  assert.strictEqual(r.value, "1241");
  assert.ok(r.rounded);
  assert.ok(!D.diagNum("1234567", null, 6).ok);
  assert.ok(!D.diagNum("abc", null, 6).ok);
  assert.ok(D.diagNumEq("320,00", "320") && D.diagNumEq("9,00", "9") && D.diagNumEq("1 240,00", "1240") && !D.diagNumEq("320,00", "321"));
});

check("type, oui/non, lettres, ADEME", () => {
  assert.strictEqual(D.diagType("").value, "11");
  assert.strictEqual(D.diagType("Habitation").value, "11");
  assert.strictEqual(D.diagType("local d'habitation").value, "11");
  assert.strictEqual(D.diagType("hors habitation").value, "15");
  assert.strictEqual(D.diagType("15").value, "15");
  assert.ok(!D.diagType("bureau").ok);
  assert.strictEqual(D.diagBool("Oui").value, true);
  assert.strictEqual(D.diagBool("non").value, false);
  assert.strictEqual(D.diagBool("").value, null);
  assert.ok(!D.diagBool("peut-être").ok);
  assert.strictEqual(D.diagLetter("c").value, "C");
  assert.ok(!D.diagLetter("H").ok);
  assert.deepStrictEqual(D.diagAdeme(" 2475 e1234567x "), { ok: true, value: "2475E1234567X", plausible: true });
  assert.strictEqual(D.diagAdeme("ABC").plausible, false);
});

check("données de la ligne : obligatoires, colonnes absentes, avertissements", () => {
  const raw = (o) => Object.fromEntries(Object.entries(o).map(([k, v]) => [k, typeof v === "string" ? { text: v } : v]));
  let r = D.diagNormalizeData(raw({ code_fournisseur: "448957", date_etabli: "15/03/2024", conso_ep: "320", emission_ges: "9" }), Date.UTC(2026, 10, 23));
  assert.deepStrictEqual(r.errors, []);
  assert.ok(r.warnings.some((w) => /N° ADEME vide/.test(w)));
  assert.strictEqual(r.v.type, "11");
  r = D.diagNormalizeData(raw({ code_fournisseur: { text: "", missing: ["N° FU"] }, date_etabli: "15/03/2030", conso_ep: "", emission_ges: "x" }), Date.UTC(2026, 10, 23));
  assert.ok(r.errors.some((e) => /colonne « N° FU » introuvable/.test(e)), r.errors.join(" | "));
  assert.ok(r.errors.some((e) => /futur/.test(e)));
  assert.ok(r.errors.some((e) => /Conso EP.*vide/.test(e)));
  assert.ok(r.errors.some((e) => /illisible/.test(e)));
  r = D.diagNormalizeData(raw({ code_fournisseur: "448957", date_etabli: "15/03/2024", conso_ep: "320", emission_ges: "9", numero_ademe: "2475E1234567X", cout_min: "2000", cout_max: "1000" }), Date.UTC(2026, 10, 23));
  assert.ok(r.warnings.some((w) => /coût min/.test(w)) && r.warnings.some((w) => /incomplets/.test(w)));
});

check("fichier : correspondance stricte", () => {
  const files = [
    { name: "IM0393664_10_DPE_2022-10-19.pdf", path: "a/IM0393664_10_DPE_2022-10-19.pdf" },
    { name: "DPE_2475E1234567X.pdf", path: "a/DPE_2475E1234567X.pdf" },
    { name: "x_2475E7654321Y.pdf", path: "a/x" }, { name: "y_2475E7654321Y.pdf", path: "a/y" },
    { name: "double.pdf", path: "a/double.pdf" }, { name: "double.pdf", path: "b/double.pdf" }
  ];
  assert.strictEqual(D.diagMatchFile(files, { name: "IM0393664_10_DPE_2022-10-19.pdf" }).via, "nom exact");
  assert.strictEqual(D.diagMatchFile(files, { name: "im0393664_10_dpe_2022-10-19" }).file.name, "IM0393664_10_DPE_2022-10-19.pdf", "sans extension, casse ignorée");
  assert.strictEqual(D.diagMatchFile(files, { name: "C:\\DPE\\DPE_2475E1234567X.pdf" }).file.name, "DPE_2475E1234567X.pdf", "chemin complet");
  assert.ok(D.diagMatchFile(files, { name: "double.pdf" }).ambiguous);
  assert.strictEqual(D.diagMatchFile(files, { name: "absent.pdf", ademe: "2475E1234567X" }).via, "nom contenant le N° ADEME");
  assert.ok(D.diagMatchFile(files, { ademe: "2475E7654321Y" }).ambiguous);
  assert.ok(/introuvable/.test(D.diagMatchFile(files, { ademe: "2475E0000000A" }).error));
  assert.ok(/ni nom de fichier/.test(D.diagMatchFile(files, {}).error));
  assert.ok(!D.diagMatchFile(files, { name: "IM0393664_10_DPE" }).file, "pas de correspondance partielle");
});

check("fichier : N° ADEME dans le nom / le texte", () => {
  assert.strictEqual(D.diagCheckFileAdeme("DPE_2475E1234567X.pdf", undefined, "2475E1234567X").status, "name");
  const o = D.diagCheckFileAdeme("DPE_2475E2222222W.pdf", undefined, "2475E1234567X");
  assert.deepStrictEqual([o.status, o.others], ["other", ["2475E2222222W"]]);
  assert.strictEqual(D.diagCheckFileAdeme("IM0393664_10_DPE_2022-10-19.pdf", undefined, "2475E1234567X").status, "needText");
  assert.strictEqual(D.diagCheckFileAdeme("a.pdf", "N° ADEME : 2475 E 1234567 X", "2475E1234567X").status, "text");
  assert.strictEqual(D.diagCheckFileAdeme("a.pdf", "N° ADEME : 2475E9999999Z", "2475E1234567X").status, "other");
  assert.strictEqual(D.diagCheckFileAdeme("a.pdf", "", "2475E1234567X").status, "unknown");
  assert.strictEqual(D.diagCheckFileAdeme("a.pdf", "", "").status, "none");
  assert.ok(D.diagIsPdf(Buffer.from("%PDF-1.7")) && D.diagIsPdf(Buffer.from("\n %PDF-1.4")) && !D.diagIsPdf(Buffer.from("<html>")));
  assert.strictEqual(Buffer.from(D.diagToBase64(new Uint8Array([37, 80, 68, 70])), "base64").toString(), "%PDF");
});

check("URL, journal, messages, grille", () => {
  const u = "http://sigeo.evoriel.net/popup.aspx/fr/sig_imm/diag_manage/393664/10/0";
  assert.deepStrictEqual(D.diagUrlParts(u), { lotId: "393664", kind: "10", idDiag: "0" });
  assert.strictEqual(D.diagUrlWithId(u, "1147729"), u.replace(/0$/, "1147729"));
  assert.strictEqual(D.diagJournalKey("393664", { ademe: "2475E1234567X" }), "393664#2475E1234567X");
  assert.strictEqual(D.diagJournalKey("393664", { ademe: "", dateEtabli: "12/05/2019", conso: "245,5", ges: "12" }), "393664#12/05/2019|245,5|12");
  const m = D.diagClassifyMessages({ items: [{ cls: "infomsg", text: "Enregistrement réussi !" }, { cls: "infomsg", text: "Info" }, { cls: "errormsg", text: "Erreur" }] });
  assert.deepStrictEqual(m, { success: true, infos: ["Info"], errors: ["Erreur"] });
  const g = { rows: [["Autre", "DPE_vieux.pdf"], ["DPE", "DPE_2475E1234567X.pdf", "23/11/2026"]] };
  assert.strictEqual(D.diagFindDocRow(g, "DPE"), g.rows[1], "description exactement « DPE »");
  assert.ok(D.diagRowHasFile(g.rows[1], "DPE_2475E1234567X.pdf"));
});

// Instantané de page minimal pour diagNextActions.
function snap(o = {}) {
  const f = (v = "", extra = {}) => Object.assign({ present: true, visible: true, value: v }, extra);
  const s = {
    divConso: true, couts: { present: true, visible: false },
    f: {
      type: f(""), fou: f(""), dateReal: f(""), dateFin: f(""), conso: f(""), ges: f(""),
      consoLettre: f(""), gesLettre: f(""), forceEp: f(""), forceGes: f(""),
      altitude: f(undefined, { checked: false, visible: false }), nouvelles: f(undefined, { checked: false, visible: false }),
      coutMin: f(""), coutMax: f(""), datePrix: f(""), resultat: f(""), comment: f(""),
      ademe: { mode: "absent", input: null, labelText: "" }, fouName: { name: "" }
    }
  };
  for (const [k, v] of Object.entries(o.f || {})) Object.assign(s.f[k], v);
  if (o.couts) s.couts = o.couts;
  if (o.divConso !== undefined) s.divConso = o.divConso;
  return s;
}
const V = {
  type: "11", fou: "448957", dateEtabli: "15/03/2024", dateValidite: "", ademe: "2475E1234567X", conso: "320", ges: "9",
  forceEp: "", forceGes: "", altitude: false, nouvelles: true, coutMin: "1240", coutMax: "1680", datePrix: "01/01/2021",
  resultat: "", comment: "ok", fichier: "", lot: ""
};

check("prochaine action : ordre du formulaire, arrêt au premier postback", () => {
  const st = { attempts: {}, warned: {} };
  let p = D.diagNextActions(snap(), V, st);
  assert.deepStrictEqual([p.soft.length, p.hard.key], [0, "type"]);
  p = D.diagNextActions(snap({ f: { type: { value: "11" } } }), V, st);
  assert.deepStrictEqual([p.soft.map((a) => a.key), p.hard.key], [["fou"], "dateReal"], "fournisseur saisi avant le postback de la date");
  p = D.diagNextActions(snap({ divConso: false, f: { type: { value: "11" } } }), V, st);
  assert.ok(/div_conso/.test(p.fatal));
  // date faite, ADEME affiché, coûts encore masqués → ADEME puis conso
  const after = { type: { value: "11" }, fou: { value: "448957" }, dateReal: { value: "15/03/2024" }, dateFin: { value: "14/03/2034" },
    ademe: { mode: "input", input: { value: "", id: "x" } } };
  p = D.diagNextActions(snap({ f: after }), V, st);
  assert.deepStrictEqual([p.soft.map((a) => a.key), p.hard.key], [["ademe"], "conso"]);
  // coûts apparus après la conso → saisis avant le postback des GES
  const after2 = Object.assign({}, after, { ademe: { mode: "input", input: { value: "2475E1234567X" } }, conso: { value: "320,00" } });
  p = D.diagNextActions(snap({ f: after2, couts: { present: true, visible: true } }), V, st);
  assert.deepStrictEqual([p.soft.map((a) => a.key), p.hard.key], [["coutMin", "coutMax", "datePrix"], "ges"]);
  // tout est saisi sauf cases / commentaire
  const after3 = Object.assign({}, after2, { ges: { value: "9,00" }, coutMin: { value: "1240" }, coutMax: { value: "1680" }, datePrix: { value: "01/01/2021" },
    nouvelles: { visible: true, checked: false }, altitude: { visible: true, checked: false } });
  p = D.diagNextActions(snap({ f: after3, couts: { present: true, visible: true } }), V, st);
  assert.deepStrictEqual([p.soft.map((a) => a.key + ":" + (a.kind === "check" ? a.checked : a.value)), p.hard], [["nouvelles:true", "comment:ok"], null]);
  const after4 = Object.assign({}, after3, { nouvelles: { visible: true, checked: true }, comment: { value: "ok" } });
  p = D.diagNextActions(snap({ f: after4, couts: { present: true, visible: true } }), V, st);
  assert.ok(p.done && !p.fatal);
  assert.deepStrictEqual(D.diagCompare(snap({ f: after4, couts: { present: true, visible: true } }), V).filter((c) => !c.ok), []);
});

check("prochaine action : avertissements ADEME, champ refusé", () => {
  const st = { attempts: {}, warned: {} };
  const base = { type: { value: "11" }, fou: { value: "448957" }, dateReal: { value: "12/05/2019" }, dateFin: { value: "31/12/2024" } };
  let p = D.diagNextActions(snap({ f: base }), Object.assign({}, V, { dateEtabli: "12/05/2019" }), st);
  assert.ok(p.warnings.some((w) => /champ ADEME non affiché/.test(w)));
  p = D.diagNextActions(snap({ f: base }), Object.assign({}, V, { dateEtabli: "12/05/2019" }), st);
  assert.ok(!p.warnings.some((w) => /champ ADEME non affiché/.test(w)), "avertissement émis une seule fois");
  // la page refuse obstinément la conso → erreur après 3 tentatives
  st.attempts.conso = 3;
  p = D.diagNextActions(snap({ f: base }), Object.assign({}, V, { dateEtabli: "12/05/2019" }), st);
  assert.ok(/n'accepte pas Conso EP/.test(p.fatal), p.fatal);
  // date à la bonne valeur mais « Valable jusqu'au » vide → un seul nouveau change
  const st2 = { attempts: {}, warned: {} };
  const V19 = Object.assign({}, V, { dateEtabli: "12/05/2019" });
  p = D.diagNextActions(snap({ f: Object.assign({}, base, { dateFin: { value: "" } }) }), V19, st2);
  assert.strictEqual(p.hard && p.hard.key, "dateReal");
  assert.strictEqual(p.hard.value, "12/05/2019");
  st2.attempts.dateReal = 1;
  p = D.diagNextActions(snap({ f: Object.assign({}, base, { dateFin: { value: "" } }) }), V19, st2);
  assert.ok(/Valable jusqu'au/.test(p.fatal));
});

check("HTML de l'étape et lecture des réglages", () => {
  const html = D.diagStepHtml({ diagCols: { code_fournisseur: "{N° FU}" }, diagConfirm: false }, (s) => String(s ?? "").replace(/"/g, "&quot;"));
  assert.ok(html.includes('data-k="code_fournisseur" value="{N° FU}"'));
  assert.ok(/scn-diag-confirm"\s*\/>/.test(html), "case confirmer décochée");
  assert.strictEqual(Object.keys(D.DIAG_PRESETS.analyse).length, D.DIAG_FIELDS.length);
});

console.log(failures ? `\n${failures} échec(s)` : "\nTous les tests passent.");
process.exit(failures ? 1 : 0);
