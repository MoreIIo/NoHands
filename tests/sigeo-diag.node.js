// Tests unitaires (logique pure) de l'étape « SIGEO : saisir un DPE ».
//   node tests/sigeo-diag.node.js
const assert = require("assert");
const D = require("../sigeo-diag.js");

let failures = 0;
const tests = [];
function check(name, fn) { tests.push([name, fn]); } // exécutés dans l'ordre, async compris

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

check("validité calculée (règles SIGEO) — cas de la demande", () => {
  assert.strictEqual(D.diagExpectedValidity("08/10/2020"), "31/12/2024");
  assert.strictEqual(D.diagExpectedValidity("15/03/2016"), "31/12/2022");
  assert.strictEqual(D.diagExpectedValidity("12/09/2022"), "11/09/2032");
});

check("validité calculée (règles SIGEO) — bornes", () => {
  assert.strictEqual(D.diagExpectedValidity("31/12/2012"), "30/12/2022", "avant 2013 : +10 ans −1 jour");
  assert.strictEqual(D.diagExpectedValidity("01/01/2013"), "31/12/2022");
  assert.strictEqual(D.diagExpectedValidity("31/12/2017"), "31/12/2022");
  assert.strictEqual(D.diagExpectedValidity("01/01/2018"), "31/12/2024");
  assert.strictEqual(D.diagExpectedValidity("15/03/2012"), "14/03/2022");
  assert.strictEqual(D.diagExpectedValidity("01/06/2015"), "31/12/2022");
  assert.strictEqual(D.diagExpectedValidity("15/01/2018"), "31/12/2024");
  assert.strictEqual(D.diagExpectedValidity("30/06/2021"), "31/12/2024");
  assert.strictEqual(D.diagExpectedValidity("01/07/2021"), "30/06/2031");
  assert.strictEqual(D.diagExpectedValidity("29/02/2024"), "28/02/2034");
  assert.ok(D.diagIsNewDpe("01/07/2021") && !D.diagIsNewDpe("30/06/2021"));
});

check("nombres à saisir : 2 décimales max, virgule, VALEUR_TROP_LONGUE", () => {
  assert.strictEqual(D.diagNum("1 234,5 kWh/m2/an", null, 6).value, "1234,5");
  assert.strictEqual(D.diagNum("224", null, 6).value, "224");
  assert.strictEqual(D.diagNum("52,6", null, 6).value, "52,6");
  assert.strictEqual(D.diagNum("52.60", null, 6).value, "52,6");
  const r = D.diagNum("320.456", null, 6);
  assert.deepStrictEqual([r.value, r.rounded], ["320,46", true]);
  assert.strictEqual(D.diagNum("1.234,5", null, 6).value, "1234,5");
  assert.strictEqual(D.diagNum("", { t: "n", v: 9 }, 6).value, "9");
  assert.strictEqual(D.diagNum("1,240", { t: "n", v: 1240 }, 5).value, "1240", "cellule numérique prioritaire");
  assert.strictEqual(D.diagNum("1 240 €", null, 5).value, "1240", "estimation : sans espace ni €");
  const long = D.diagNum("1240,75", null, 5);
  assert.deepStrictEqual([long.ok, long.code], [false, "VALEUR_TROP_LONGUE"], "plus d'arrondi pour faire tenir");
  assert.ok(/7 caractères, 5 au maximum/.test(long.error), long.error);
  assert.strictEqual(D.diagNum("1234,56", null, 6).code, "VALEUR_TROP_LONGUE");
  assert.strictEqual(D.diagNum("1234567", null, 6).code, "VALEUR_TROP_LONGUE");
  assert.ok(!D.diagNum("abc", null, 6).ok);
});

check("comparateur normalisé", () => {
  const S = D.diagSame;
  // nombres : parse FR, |a−b| < 0,005
  assert.ok(S("num", "224,00", "224") && S("num", "224", "224,00"));
  assert.ok(S("num", "52,60", "52,6") && S("num", "52,6", "52,60"));
  assert.ok(S("num", "1 240,00 €", "1240") && S("num", "224,00 kWh/m²/an", "224") && S("num", "52,6 kgCO2/m²/an", "52,60"));
  assert.ok(S("num", "52,604", "52,6") && !S("num", "52,61", "52,6") && !S("num", "", "52,6") && !S("num", "224,00", "225"));
  // dates : jj/mm/aaaa, zéros de tête
  assert.ok(S("date", "8/10/2020", "08/10/2020") && S("date", "08/10/2020", "8/10/2020") && S("date", " 31/12/2024 ", "31/12/2024"));
  assert.ok(!S("date", "", "31/12/2024") && !S("date", "31/12/2025", "31/12/2024"));
  // ADEME : trim + majuscules
  assert.ok(S("ademe", " 2021l2000833e ", "2021L2000833E") && !S("ademe", "", "2021L2000833E"));
  // divers
  assert.ok(S("code", " 448957", "448957") && S("letter", "d", "D") && S("check", true, true) && !S("check", false, true));
  assert.ok(S("select", "11", "11") && !S("select", "15", "11") && S("text", "a  b", "a b"));
  assert.strictEqual(D.diagShow("check", true), "coché");
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
  assert.deepStrictEqual(D.diagAdeme(" 2021l2000833e"), { ok: true, value: "2021L2000833E", plausible: true });
  assert.strictEqual(D.diagAdeme("ABC").plausible, false);
  assert.strictEqual(D.diagAdeme("2021L200083").plausible, false);
  // N° non communiqué : pas d'erreur, rien à saisir
  for (const x of ["NC", "nc", "N.C.", "non communiqué", "Non trouvé", "[NON EMIS ADEME]", "non émis", "99999", "-", "aucun"]) {
    const r = D.diagAdeme(x);
    assert.ok(r.nc && r.value === "" && r.plausible, "NC : " + x);
  }
  assert.ok(!D.diagAdeme("2021N1001671X").nc);
});

check("fournisseur : préfixe FU, zéros de tête supprimés par SIGEO", () => {
  const K = D.diagFouKey;
  assert.deepStrictEqual(["0448957", "448957", "FU0448957", "fu 0448957", " 448957 ", "000", ""].map(K), ["448957", "448957", "448957", "448957", "448957", "0", "0"]);
  assert.ok(D.diagSame("fou", "448957", "0448957") && D.diagSame("fou", "0448957", "FU448957") && !D.diagSame("fou", "448957", "448958"));
  const raw = (o) => Object.fromEntries(Object.entries(o).map(([k, v]) => [k, { text: v }]));
  const r = D.diagNormalizeData(raw({ code_fournisseur: "FU 0448957", date_etabli: "12/09/2022", numero_ademe: "2222E0000001A", conso_ep: "224", emission_ges: "52,6" }), Date.UTC(2026, 10, 24));
  assert.strictEqual(r.v.fou, "0448957", "« FU » et espaces retirés avant la saisie, zéros conservés");
  // plan : relu « 448957 » après enregistrement = conforme ; saisie validée par blur
  const p = D.diagPlan(snap({ f: { fou: { value: "448957" } } }), Object.assign({}, REALV, { fou: "0448957" }));
  const f = p.specs.find((x) => x.key === "fou");
  assert.deepStrictEqual([f.kind, f.ok, f.act.blur], ["fou", true, true]);
  assert.strictEqual(D.DIAG_IDS.fouLabel, "body:x:proxyFou:x:lbl:x:lbl");
});

check("données de la ligne : N° ADEME « NC »", () => {
  const raw = (o) => Object.fromEntries(Object.entries(o).map(([k, v]) => [k, { text: v }]));
  const r = D.diagNormalizeData(raw({ code_fournisseur: "448957", date_etabli: "30/04/2025", numero_ademe: "NC", conso_ep: "107", emission_ges: "5" }), Date.UTC(2026, 10, 23));
  assert.deepStrictEqual(r.errors, [], "plus d'ADEME_INVALIDE pour NC");
  assert.deepStrictEqual([r.v.ademe, r.v.ademeNc], ["", true]);
  assert.ok(r.warnings.includes("N° ADEME non communiqué (« NC ») : non saisi"), r.warnings.join(" | "));
  assert.ok(!r.warnings.some((w) => /N° ADEME vide/.test(w)));
  const p = D.diagPlan(snap({ f: { ademe: { mode: "input", input: { present: true, visible: true, value: "", id: "x" }, labelText: "" } } }),
    Object.assign({}, REALV, { ademe: "", ademeNc: true, dateEtabli: "30/04/2025" }));
  assert.ok(!p.specs.some((x) => x.key === "ademe") && !p.warnings.some((w) => /ADEME/.test(w)));
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

check("données de la ligne : cas réel, ADEME_INVALIDE, VALEUR_TROP_LONGUE, validité du PDF", () => {
  const raw = (o) => Object.fromEntries(Object.entries(o).map(([k, v]) => [k, { text: v }]));
  const real = { code_fournisseur: "448957", date_etabli: "08/10/2020", numero_ademe: "2021L2000833E", conso_ep: "224", emission_ges: "52,6",
    lettre_ep_forcee: "d", date_debut_validite: "8/10/2020" };
  let r = D.diagNormalizeData(raw(real), Date.UTC(2026, 10, 23));
  assert.deepStrictEqual(r.errors, []);
  assert.deepStrictEqual([r.v.fou, r.v.dateEtabli, r.v.ademe, r.v.conso, r.v.ges, r.v.lettreEp, r.v.dateDeb], ["448957", "08/10/2020", "2021L2000833E", "224", "52,6", "D", "08/10/2020"]);
  r = D.diagNormalizeData(raw(Object.assign({}, real, { numero_ademe: "2021L200083" })), Date.UTC(2026, 10, 23));
  assert.ok(r.errors.some((e) => /^ADEME_INVALIDE — « 2021L200083 »/.test(e)), r.errors.join(" | "));
  r = D.diagNormalizeData(raw(Object.assign({}, real, { conso_ep: "1234,56", cout_min: "1240,75" })), Date.UTC(2026, 10, 23));
  assert.ok(r.errors.some((e) => /^VALEUR_TROP_LONGUE — Conso EP « 1234,56 » : 7 caractères, 6 au maximum/.test(e)), r.errors.join(" | "));
  assert.ok(r.errors.some((e) => /^VALEUR_TROP_LONGUE — Coût min « 1240,75 »/.test(e)), r.errors.join(" | "));
  // « Valable jusqu'au » du PDF : avertissement seulement
  r = D.diagNormalizeData(raw(Object.assign({}, real, { date_validite: "07/10/2030" })), Date.UTC(2026, 10, 23));
  assert.deepStrictEqual(r.errors, []);
  assert.ok(r.warnings.some((w) => /Valable jusqu'au » du DPE \(07\/10\/2030\) ≠ calcul SIGEO \(31\/12\/2024\)/.test(w)), r.warnings.join(" | "));
  r = D.diagNormalizeData(raw(Object.assign({}, real, { date_validite: "31/12/2024" })), Date.UTC(2026, 10, 23));
  assert.ok(!r.warnings.some((w) => /Valable jusqu'au/.test(w)));
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

// Instantané de page minimal (op "read").
function snap(o = {}) {
  const f = (v = "", extra = {}) => Object.assign({ present: true, visible: true, value: v, pb: false }, extra);
  const s = {
    docId: "d1", readyState: "complete", leaving: false,
    divConso: true, couts: { present: true, visible: false },
    f: {
      type: f("", { pb: true }), fou: f(""), dateReal: f("", { pb: true }), dateFin: f("", { visible: false, label: "", shown: "" }), dateDeb: f(""),
      conso: f("", { pb: true }), ges: f("", { pb: true }),
      consoLettre: f(""), gesLettre: f(""), forceEp: f(""), forceGes: f(""),
      altitude: f(undefined, { checked: false, visible: false }), nouvelles: f(undefined, { checked: false, visible: false }),
      coutMin: f(""), coutMax: f(""), datePrix: f(""), resultat: f(""), comment: f(""),
      ademe: { mode: "input", input: { present: true, visible: true, value: "", id: "body_x_textIdDpe_x_textIdDpe", pb: false }, labelText: "" },
      fouName: { name: "" }
    }
  };
  for (const [k, v] of Object.entries(o.f || {})) s.f[k] = Object.assign({}, s.f[k], v);
  if (o.couts) s.couts = o.couts;
  if (o.divConso !== undefined) s.divConso = o.divConso;
  return s;
}
const REALV = {
  type: "11", fou: "448957", dateEtabli: "08/10/2020", dateValidite: "", dateDeb: "", ademe: "2021L2000833E", conso: "224", ges: "52,6",
  lettreEp: "", lettreGes: "", altitude: null, nouvelles: null, coutMin: "", coutMax: "", datePrix: "", resultat: "", comment: "", fichier: "", lot: ""
};
const gaps = (s, v, o) => D.diagPlan(s, v, o).specs.filter((x) => !x.ok);

check("plan : phase A (postback) dans l'ordre, puis phase B", () => {
  assert.deepStrictEqual(D.DIAG_PHASE_A, ["type", "fou", "dateReal", "conso", "ges"]);
  const p = D.diagPlan(snap(), REALV);
  assert.deepStrictEqual(p.specs.map((x) => x.key + ":" + x.phase), ["type:A", "fou:A", "dateReal:A", "dateFin:calc", "conso:A", "ges:A", "ademe:B"]);
  assert.strictEqual(p.specs.find((x) => x.key === "ademe").act.id, "body_x_textIdDpe_x_textIdDpe", "ADEME par son id réel");
  assert.strictEqual(p.specs.find((x) => x.key === "dateFin").act, null, "« Valable jusqu'au » jamais saisi");
  assert.strictEqual(p.specs.find((x) => x.key === "dateFin").want, "31/12/2024");
  // page remplie comme après enregistrement (valeurs réaffichées par SIGEO)
  const done = snap({ f: { type: { value: "11" }, fou: { value: "448957" }, dateReal: { value: "08/10/2020" },
    dateFin: { value: "31/12/2024", label: "31/12/2024", shown: "31/12/2024" }, conso: { value: "224,00" }, ges: { value: "52,60" },
    consoLettre: { value: "D" }, gesLettre: { value: "E" },
    ademe: { mode: "input", input: { present: true, visible: true, value: "2021L2000833E", id: "body_x_textIdDpe_x_textIdDpe" }, labelText: "" } } });
  assert.deepStrictEqual(gaps(done, REALV), []);
  // ADEME vide après enregistrement → un écart, lu dans le champ de saisie
  const lost = snap({ f: Object.assign({}, done.f, { ademe: { mode: "input", input: { present: true, visible: true, value: "", id: "x" }, labelText: "" } }) });
  const g = gaps(lost, REALV);
  assert.deepStrictEqual(g.map((x) => [x.label, x.want, x.got]), [["N° ADEME", "2021L2000833E", ""]]);
  assert.strictEqual(D.diagGapText(g), "1 écart(s) : N° ADEME attendu « 2021L2000833E », relu « »");
  // ADEME en libellé (fiche enregistrée, ancienne page)
  const lab = snap({ f: Object.assign({}, done.f, { ademe: { mode: "label", input: null, labelText: "2021L2000833E" } }) });
  assert.deepStrictEqual(gaps(lab, REALV), []);
  assert.strictEqual(D.diagPlan(lab, REALV).specs.find((x) => x.key === "ademe").act, null);
  // « Valable jusqu'au » ≠ calcul → écart (signe que « Établi le » n'a pas été pris en compte)
  const badFin = snap({ f: Object.assign({}, done.f, { dateFin: { value: "", label: "", shown: "" } }) });
  assert.deepStrictEqual(gaps(badFin, REALV).map((x) => x.key), ["dateFin"]);
  // tous les écarts listés
  const two = snap({ f: Object.assign({}, lost.f, { ges: { value: "53,60" } }) });
  assert.strictEqual(D.diagGapText(gaps(two, REALV)), "2 écart(s) : Émission GES attendu « 52,6 », relu « 53,60 » ; N° ADEME attendu « 2021L2000833E », relu « »");
});

check("plan : phase B complète, avertissements", () => {
  const v = Object.assign({}, REALV, { dateEtabli: "15/03/2024", dateDeb: "15/03/2024", coutMin: "1240", coutMax: "1680", datePrix: "01/01/2021",
    nouvelles: true, altitude: false, comment: "ok" });
  let p = D.diagPlan(snap({ couts: { present: true, visible: true }, f: { nouvelles: { visible: true, checked: false }, altitude: { visible: true, checked: false } } }), v);
  assert.deepStrictEqual(p.specs.filter((x) => x.phase === "B").map((x) => x.key), ["ademe", "dateDeb", "coutMin", "coutMax", "datePrix", "altitude", "nouvelles", "comment"]);
  p = D.diagPlan(snap(), v);
  assert.ok(p.warnings.includes("bloc « coûts annuels » non affiché : coûts non saisis"));
  assert.ok(p.warnings.some((w) => /case « Nouvelles valeurs du DPE » non affichée/.test(w)));
});

check("règle de présence de l'ADEME : ademeSaisissableAvantEnreg", () => {
  const f = D.ademeSaisissableAvantEnreg;
  assert.strictEqual(f("30/06/2021"), false);
  assert.strictEqual(f("01/07/2021"), true);
  assert.strictEqual(f("24/11/2020"), false);
  assert.strictEqual(f("12/09/2022"), true);
  assert.strictEqual(f("1/7/2021"), true, "zéros de tête normalisés");
  assert.strictEqual(f(""), false);
  assert.strictEqual(f("n'importe quoi"), false);
});

check("plan : ADEME absent (plus d'avertissement), différé, libellé prioritaire", () => {
  const OLD = Object.assign({}, REALV, { dateEtabli: "24/11/2020", ademe: "2021N1001671X" });
  // champ absent → écart non saisissable, plus de simple avertissement
  let p = D.diagPlan(snap({ f: { ademe: { mode: "absent", input: null, labelText: "" } } }), OLD);
  const a = p.specs.find((x) => x.key === "ademe");
  assert.deepStrictEqual([a.ok, a.missing, a.act], [false, true, null]);
  assert.ok(!p.warnings.some((w) => /ADEME/.test(w)), p.warnings.join(" | "));
  assert.strictEqual(D.diagGapText([a]), "1 écart(s) : N° ADEME attendu « 2021N1001671X », relu « » (champ absent)");
  // DPE < 01/07/2021 en création : ADEME différé, absent du plan
  p = D.diagPlan(snap({ f: { ademe: { mode: "absent", input: null, labelText: "" } } }), OLD, { ademeDeferred: true });
  assert.ok(!p.specs.some((x) => x.key === "ademe") && !p.warnings.some((w) => /ADEME/.test(w)));
  // enregistré : lu dans le libellé, non saisissable
  p = D.diagPlan(snap({ f: { ademe: { mode: "label", input: null, labelText: " 2021n1001671x " } } }), OLD);
  const l = p.specs.find((x) => x.key === "ademe");
  assert.deepStrictEqual([l.ok, l.act], [true, null]);
  // input vide après le 1er enregistrement → saisissable
  p = D.diagPlan(snap({ f: { ademe: { mode: "input", input: { present: true, visible: true, value: "", id: "body_x_textIdDpe_x_textIdDpe" }, labelText: "" } } }), OLD);
  assert.strictEqual(p.specs.find((x) => x.key === "ademe").act.id, "body_x_textIdDpe_x_textIdDpe");
});

check("plan : lettres du DPE forcées seulement si SIGEO calcule autre chose", () => {
  const v = Object.assign({}, REALV, { lettreEp: "D", lettreGes: "B" });
  const page = (o) => snap({ f: Object.assign({ consoLettre: { value: "E" }, gesLettre: { value: "B" } }, o) });
  let p = D.diagPlan(page(), v);
  const ep = p.specs.find((x) => x.key === "forceEp");
  assert.deepStrictEqual([ep.want, ep.ok, ep.phase], ["D", false, "B"]);
  assert.ok(/lettre EP du DPE D ≠ lettre calculée par SIGEO E → forcée/.test(ep.note));
  assert.ok(!p.specs.some((x) => x.key === "forceGes"), "GES identique : pas de forçage");
  // forcée → conforme, même si la légende affiche désormais D
  assert.ok(D.diagPlan(page({ forceEp: { value: "D" }, consoLettre: { value: "D" } }), v).specs.find((x) => x.key === "forceEp").ok);
  // forçage contraire alors que SIGEO calcule la bonne lettre → à vider
  const clr = D.diagPlan(page({ gesLettre: { value: "B" }, forceGes: { value: "C" } }), v).specs.find((x) => x.key === "forceGes");
  assert.deepStrictEqual([clr.want, clr.ok], ["", false]);
  // option désactivée → avertissement, rien à saisir
  p = D.diagPlan(page(), v, { forceLetters: false });
  assert.ok(!p.specs.some((x) => x.key === "forceEp"));
  assert.ok(p.warnings.some((w) => /non forcée \(option désactivée\)/.test(w)));
});

// Faux adaptateur à horloge virtuelle pour diagSettle.
function fakeAd(script, o = {}) {
  let t = 0;
  return {
    now: () => t, sleep: async (ms) => { t += ms; }, stopped: () => false,
    tabLoading: o.tabLoading ? async () => o.tabLoading(t) : undefined,
    read: async () => script(t)
  };
}
const page = (docId, extra = {}) => Object.assign({ docId, readyState: "complete", leaving: false }, extra);

check("fireAndSettle : rechargement, calme, navigation en attente, délai", async () => {
  // a) postback : nouveau docId après 400 ms (null pendant le remplacement)
  let ad = fakeAd((t) => t < 200 ? page("A", { leaving: t > 30 }) : t < 400 ? null : page("B"));
  let r = await D.diagSettle(ad, ad.read, "A", { quietMs: 1500, timeoutMs: 15000 });
  assert.deepStrictEqual([r.reloaded, r.snap.docId], [true, "B"]);
  // b) pas de postback : 1,5 s de calme
  ad = fakeAd(() => page("A"));
  r = await D.diagSettle(ad, ad.read, "A", { quietMs: 1500, timeoutMs: 15000 });
  assert.strictEqual(r.reloaded, false);
  assert.ok(ad.now() >= 1500 && ad.now() < 2000, "≈ 1,5 s : " + ad.now());
  // c) serveur lent : ancien document encore là 2,5 s, mais formulaire envoyé → on attend
  ad = fakeAd((t) => t < 2500 ? page("A", { leaving: true }) : page("B"));
  r = await D.diagSettle(ad, ad.read, "A", { quietMs: 1500, timeoutMs: 15000 });
  assert.deepStrictEqual([r.reloaded, r.snap.docId], [true, "B"], "jamais « pas de postback » tant que la navigation est en attente");
  // c bis) même cas signalé seulement par l'onglet en chargement
  ad = fakeAd((t) => t < 2500 ? page("A") : page("B"), { tabLoading: (t) => t < 2500 });
  r = await D.diagSettle(ad, ad.read, "A", { quietMs: 1500, timeoutMs: 15000 });
  assert.strictEqual(r.reloaded, true);
  // d) nouveau document pas encore complet → on attend complete
  ad = fakeAd((t) => t < 300 ? page("A") : t < 900 ? page("B", { readyState: "interactive" }) : page("B"));
  r = await D.diagSettle(ad, ad.read, "A", { quietMs: 1500, timeoutMs: 15000 });
  assert.ok(r.reloaded && ad.now() >= 900);
  // e) navigation qui ne finit jamais → timeout 15 s
  ad = fakeAd(() => page("A", { leaving: true }));
  r = await D.diagSettle(ad, ad.read, "A", { quietMs: 1500, timeoutMs: 15000 });
  assert.ok(r.timeout && ad.now() > 15000 && ad.now() < 15500);
  // f) mustReload (Enregistrer) : le calme ne suffit pas
  ad = fakeAd(() => page("A"));
  r = await D.diagSettle(ad, ad.read, "A", { quietMs: 1500, timeoutMs: 4000, mustReload: true });
  assert.ok(r.timeout);
  // f bis) mustReload + doneWhen : mise à jour partielle sur le même document
  ad = fakeAd((t) => page("A", { grid: { count: t < 800 ? 0 : 1 } }));
  r = await D.diagSettle(ad, ad.read, "A", { quietMs: 1500, timeoutMs: 4000, mustReload: true, doneWhen: (x) => x.grid.count > 0 });
  assert.deepStrictEqual([r.partial, r.reloaded], [true, false]);
  // g) lecture qui ne répond plus (frame en navigation) : le délai s'applique quand même
  ad = fakeAd(() => new Promise(() => {}));
  r = await D.diagSettle(ad, ad.read, "A", { quietMs: 1500, timeoutMs: 3000, readCapMs: 5 });
  assert.ok(r.timeout, "POSTBACK_TIMEOUT même si la page ne répond pas");
  assert.strictEqual(await D.diagCapped(new Promise(() => {}), 5), null);
  assert.strictEqual(await D.diagCapped(Promise.resolve(7), 50), 7);
});

check("date : performanceEnergetique(el) dans la page, repli __doPostBack", () => {
  const calls = [];
  const el = { id: "body_x_txtDateReal_x_txtDateReal", name: "body$x$txtDateReal$x$txtDateReal", getAttribute: () => null };
  const saved = { window: global.window, document: global.document };
  global.document = { getElementById: (id) => id === el.id ? el : null };
  try {
    global.window = { performanceEnergetique: (e) => calls.push(["P", e === el]), __doPostBack: (t, a) => calls.push(["D", t, a]) };
    assert.deepStrictEqual(D.diagMainHandler({ mode: "date", id: el.id }), { ok: true, via: "performanceEnergetique" });
    assert.deepStrictEqual(calls, [["P", true]], "pas de double postback");
    calls.length = 0;
    global.window = { performanceEnergetique: () => { throw new Error("boum"); }, __doPostBack: (t, a) => calls.push(["D", t, a]) };
    let r = D.diagMainHandler({ mode: "date", id: el.id });
    assert.deepStrictEqual([r.ok, r.via, calls], [true, "__doPostBack", [["D", "body$x$txtDateReal$x$txtDateReal", ""]]], "repli avec l'UniqueID");
    assert.ok(/boum/.test(r.note));
    global.window = {};
    r = D.diagMainHandler({ mode: "date", id: el.id });
    assert.ok(!r.ok && /performanceEnergetique absente/.test(r.error) && /__doPostBack absente/.test(r.error), r.error);
    assert.strictEqual(D.diagMainHandler({ mode: "date", id: "x" }).ok, false);
  } finally { global.window = saved.window; global.document = saved.document; }
});

check("repli : handler onchange appelé instruction par instruction (monde MAIN)", () => {
  const calls = [];
  const el = { id: "d", value: "08/10/2020", getAttribute: (a) => a === "onchange" ? "javascript:ValidatorOnChange(event); performanceEnergetique(this); return x.y('a\\'b', 2);" : null };
  const saved = { window: global.window, document: global.document };
  global.document = { getElementById: (id) => id === "d" ? el : null };
  global.window = {
    ValidatorOnChange: (ev) => { calls.push(["V", ev.type]); throw new Error("vals is undefined"); },
    performanceEnergetique(e) { calls.push(["P", e === el, this === global.window]); },
    x: { y(a, n) { calls.push(["Y", a, n, this === global.window.x]); } }
  };
  try {
    const r = D.diagMainHandler({ id: "d" });
    assert.ok(r.ok);
    assert.deepStrictEqual(calls, [["V", "change"], ["P", true, true], ["Y", "a'b", 2, true]], "une erreur n'empêche pas la suite");
    assert.ok(/vals is undefined/.test(r.calls[0].error));
    assert.deepStrictEqual(D.diagMainHandler({ id: "zz" }).ok, false);
    el.getAttribute = () => "foo(bar)";
    assert.strictEqual(D.diagMainHandler({ id: "d" }).calls[0].skipped, "fonction absente");
  } finally { global.window = saved.window; global.document = saved.document; }
});

check("HTML de l'étape et lecture des réglages", () => {
  const html = D.diagStepHtml({ diagCols: { code_fournisseur: "{N° FU}" }, diagConfirm: false }, (s) => String(s ?? "").replace(/"/g, "&quot;"));
  assert.ok(html.includes('data-k="code_fournisseur" value="{N° FU}"'));
  assert.ok(/scn-diag-confirm"\s*\/>/.test(html), "case confirmer décochée");
  assert.ok(/scn-diag-forceletters" checked/.test(html), "forcer la lettre : coché par défaut");
  assert.ok(!/scn-diag-forceletters" checked/.test(D.diagStepHtml({ diagForceLetters: false }, String)));
  assert.strictEqual(Object.keys(D.DIAG_PRESETS.analyse).length, D.DIAG_FIELDS.length);
  assert.strictEqual(D.DIAG_PRESETS.analyse.lettre_ep_forcee, "{Classe énergie}");
  assert.strictEqual(D.DIAG_IDS.ademe, "body_x_textIdDpe_x_textIdDpe");
  assert.strictEqual(D.DIAG_CODES.ademeAbsent, "ADEME_CHAMP_ABSENT");
  assert.strictEqual(D.DIAG_DEFAULTS.timeoutMs, 15000);
  assert.ok(/scn-diag-timeout"[^>]*value="15000"/.test(D.diagStepHtml({ diagTimeout: 20000 }, String)), "ancien défaut 20 s → 15 s");
  assert.ok(/scn-diag-timeout"[^>]*value="30000"/.test(D.diagStepHtml({ diagTimeout: 30000 }, String)), "valeur personnalisée conservée");
});

(async () => {
  for (const [name, fn] of tests) {
    try { await fn(); console.log("✓ " + name); }
    catch (e) { failures++; console.log("✗ " + name + "\n   " + (e.stack || e).toString().split("\n").slice(0, 8).join("\n   ")); }
  }
  console.log(failures ? `\n${failures} échec(s)` : "\nTous les tests passent.");
  process.exit(failures ? 1 : 0);
})();
