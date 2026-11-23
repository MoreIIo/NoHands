// Test d'intégration : cas réels et phases A→E de l'étape « SIGEO : saisir
// un DPE », sur un faux SIGEO conforme aux faits vérifiés le 23/11/2026 :
//   - « Établi le » : un change synthétique n'atteint pas performanceEnergetique ;
//   - en création, l'input N° ADEME n'existe que si la date ≥ 01/07/2021 ;
//     après enregistrement il est présent (vide) tant qu'aucun N° n'est
//     enregistré, puis le N° devient un texte (span#body_x_labelIdDpe).
//   NODE_PATH=$(npm root -g) node tests/sigeo-diag.realcase.e2e.js
// Critères :
//   1. simulation (dryRun) d'un DPE d'avant 07/2021 : ADEME non tenté, rien d'enregistré ;
//   2. « ancien DPE » 24/11/2020 + ADEME 2021N1001671X (cas de la fiche 1147787) :
//      2 enregistrements, ADEME relu = 2021N1001671X, PDF joint ;
//   3. « nouveau DPE » (≥ 01/07/2021) : 1 seul enregistrement ;
//   4. serveur lent (2,5 s > fenêtre de calme) sur un ancien DPE ;
//   5. ADEME_INVALIDE et VALEUR_TROP_LONGUE avant toute saisie ;
//   6. ECART_AVANT_ENREG (tous les écarts) ;
//   7. ECART_APRES_ENREG + « fiche DPE {id} créée mais incomplète » ;
//   8. ADEME_CHAMP_ABSENT (DPE ≥ 07/2021 sans champ) : aucun enregistrement ;
//   9. échec après l'Enregistrer n°1 puis reprise (journal) : aucune nouvelle fiche ;
//  10. idem, reprise par la colonne résultat après « Vider l'historique DPE » ;
//  11. POSTBACK_TIMEOUT ;
//  12. lettre du DPE ≠ lettre SIGEO → « Forcer la lettre » (et seulement là) ;
//  13. PDF absent de la liste juste après l'upload : « Enregistrer » recliqué,
//      document vu au 2e → OK avec avertissement, un seul upload ;
//  14. PDF jamais visible : erreur après 3 tentatives, puis relances sans
//      nouvel upload jusqu'à ce qu'il apparaisse ;
//  15. N° ADEME « NC » : rien de saisi, pas d'erreur ; PDF joint
//      automatiquement sur la ligne active (confirmation décochée).
const path = require("path");
const fs = require("fs");
const os = require("os");
const assert = require("assert");
const { chromium } = require("playwright");
const mock = require("./fixtures/sigeo-diag-server.js");
const { makePdf, COLS, row, openPanel, runActive } = require("./fixtures/diag-panel-harness.js");

const OLD = { code_fournisseur: "590971", date_etabli: "24/11/2020", numero_ademe: "2021N1001671X", conso_ep: "65", emission_ges: "8,1" };
const NEW = { code_fournisseur: "448957", date_etabli: "12/09/2022", numero_ademe: "2222E0000001A", conso_ep: "224", emission_ges: "52,6",
  cout_min: "900", cout_max: "1300", date_prix: "01/01/2021" };
const OLD2 = { code_fournisseur: "448957", date_etabli: "08/10/2020", numero_ademe: "2021L2000833E", conso_ep: "224", emission_ges: "52,6" };

(async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "nh-dpe-real-"));
  for (const a of ["2021N1001671X", "2222E0000001A", "2021L2000833E", "2475E1234567X"]) {
    fs.writeFileSync(path.join(dir, `DPE_${a}.pdf`), makePdf(["Diagnostic de performance energetique", "N ADEME : " + a]));
  }

  const srv = await mock.start();
  const cfg = (q) => fetch(srv.origin + "/__config?" + q);
  const { browser, tab, panel, errors } = await openPanel(chromium, srv.url("590001", 0, 2));
  const cell = (r, col) => panel.evaluate(([r, c]) => getCellByIndex(state.rows[r], colIndexByName(c)), [r, col]);
  const lotDiags = (lot) => srv.diags().filter((d) => d.lot === lot);
  const stats = () => Object.assign({}, srv.stats);
  const delta = (a, b) => ({ creates: b.creates - a.creates, updates: b.updates - a.updates, uploads: b.uploads - a.uploads, saves: b.saves - a.saves });
  const pageVals = () => tab.evaluate(() => {
    const v = (id) => { const e = document.getElementById(id); return e ? e.value : null; };
    const t = (id) => { const e = document.getElementById(id); return e ? e.textContent.trim() : null; };
    return {
      conso: v("body_x_txtConsoE_x_txtConsoE"), ges: v("body_x_txtEmission_x_txtEmission"),
      fin: v("body_x_txtDateFinValid_x_txtDateFinValid"), ademeInput: v("body_x_textIdDpe_x_textIdDpe"),
      ademeLabel: t("body_x_labelIdDpe"), date: v("body_x_txtDateReal_x_txtDateReal")
    };
  });

  await panel.evaluate(([cols, rows]) => {
    state.headerMode = "yes";
    state.rows = [cols].concat(rows);
    renderColumns();
  }, [COLS, [
    /* L2  */ row(Object.assign({ lot: "590001" }, OLD)),
    /* L3  */ row(Object.assign({ lot: "590002" }, NEW)),
    /* L4  */ row(Object.assign({ lot: "590003" }, OLD2)),
    /* L5  */ row(Object.assign({}, OLD, { lot: "590004", numero_ademe: "2021N100167" })),
    /* L6  */ row(Object.assign({}, OLD, { lot: "590005", conso_ep: "1234,56", cout_min: "1240,75" })),
    /* L7  */ row(Object.assign({}, NEW, { lot: "590006", date_debut_validite: "12/09/2022", commentaire: "test écart" })),
    /* L8  */ row(Object.assign({}, NEW, { lot: "590007" })),
    /* L9  */ row(Object.assign({}, NEW, { lot: "590008" })),
    /* L10 */ row(Object.assign({ lot: "590009" }, OLD)),
    /* L11 */ row(Object.assign({ lot: "590010" }, OLD)),
    /* L12 */ row(Object.assign({ lot: "590011" }, OLD)),
    /* L13 */ row({ code_fournisseur: "448957", date_etabli: "15/03/2024", numero_ademe: "2475E1234567X", conso_ep: "320", emission_ges: "9",
      lettre_ep_forcee: "D", lettre_ges_forcee: "B", cout_min: "1240", cout_max: "1680", date_prix: "01/01/2021", lot: "590012" }),
    /* L14 */ row(Object.assign({}, NEW, { lot: "590013" })),
    /* L15 */ row(Object.assign({}, NEW, { lot: "590014" })),
    /* L16 */ row(Object.assign({}, OLD, { lot: "590015", numero_ademe: "NC", fichier_dpe: "DPE_2021N1001671X.pdf" }))
  ]]);

  await panel.evaluate((u) => {
    document.getElementById("scenarioSteps").innerHTML = "";
    $("addDpeTemplateBtn").click();
    document.querySelector(".scn-step .scn-diag-url").value = u;
  }, srv.origin + "/popup.aspx/fr/sig_imm/diag_manage/{lot}/2/0");
  await panel.setInputFiles(".scn-diag-folderinput", dir);
  const setOpt = (cls, checked) => panel.evaluate(([c, v]) => { document.querySelector(c).checked = v; }, [cls, checked]);
  const yes = async () => { await panel.click("#scnConfirmYes"); };

  // ===== 1. Simulation d'un ancien DPE : ADEME non tenté =====
  await setOpt(".scn-diag-simulate", true);
  let s0 = stats();
  let run = await runActive(panel, 1);
  console.log("--- 1. simulation (DPE du 24/11/2020) ---\n" + run.log);
  await setOpt(".scn-diag-simulate", false);
  let c = await cell(1, "Saisie DPE");
  assert.ok(/^SIMULATION OK — .*N° ADEME 2021N1001671X non tenté \(DPE avant 01\/07\/2021/.test(c), c);
  assert.ok(/simulation : DPE avant 01\/07\/2021 — N° ADEME non tenté/.test(run.log), run.log);
  assert.strictEqual(delta(s0, stats()).saves, 0, "rien d'enregistré en simulation");
  let pv = await pageVals();
  assert.deepStrictEqual([pv.conso, pv.ges, pv.fin, pv.ademeInput, pv.date], ["65,00", "8,10", "31/12/2024", null, "24/11/2020"], "pas d'input ADEME en création avant 07/2021 : " + JSON.stringify(pv));
  assert.ok(/Établi le : performanceEnergetique\(\) appelé dans la page/.test(run.log), "date déclenchée directement");
  assert.ok(!/pas de postback après le change synthétique/.test(run.log), "plus d'attente inutile avant le handler");
  assert.ok(/Valable jusqu'au : relu « 31\/12\/2024 » · calcul attendu 31\/12\/2024 ✓/.test(run.log), run.log);

  // ===== 2. Ancien DPE (fiche 1147787) : 2 enregistrements =====
  s0 = stats();
  run = await runActive(panel, 1, async (body) => { assert.ok(body.includes("2021N1001671X (confirmé par SIGEO)"), body); await yes(); });
  console.log("--- 2. ancien DPE 24/11/2020 ---\n" + run.log);
  assert.deepStrictEqual(delta(s0, stats()), { creates: 1, updates: 1, uploads: 1, saves: 2 }, "2 enregistrements, 1 seule fiche");
  let d = lotDiags("590001");
  assert.strictEqual(d.length, 1);
  assert.deepStrictEqual([d[0].fou, d[0].dateReal, d[0].dateFin, d[0].conso, d[0].ges, d[0].ademe, d[0].docs.length],
    ["590971", "24/11/2020", "31/12/2024", "65,00", "8,10", "2021N1001671X", 1], JSON.stringify(d[0]));
  assert.ok(/ADEME saisi après 1er enregistrement \(DPE avant 01\/07\/2021\)/.test(run.log));
  assert.ok(/enregistrement n°1 réussi — fiche DPE \d+ créée : http:\/\/127\.0\.0\.1:\d+\/popup\.aspx\/fr\/sig_imm\/diag_manage\/590001\/2\/\d+/.test(run.log), "id et lien logués");
  assert.ok(/N° ADEME : attendu « 2021N1001671X » · avant « » · après « 2021N1001671X » \(après le 1er enregistrement\)/.test(run.log));
  assert.ok(/Enregistrer n°2 → mise à jour du diagnostic/.test(run.log) && /enregistrement n°2 réussi/.test(run.log));
  assert.ok(/N° ADEME relu après enregistrement : 2021N1001671X \(texte enregistré\)/.test(run.log), run.log);
  c = await cell(1, "Saisie DPE");
  assert.ok(/^OK — diagnostic \d+, ADEME 2021N1001671X — PDF joint \(DPE_2021N1001671X\.pdf\)/.test(c), c);
  pv = await pageVals();
  assert.deepStrictEqual([pv.ademeLabel, pv.ademeInput], ["2021N1001671X", null], "ADEME devenu texte non modifiable");

  // ===== 3. Nouveau DPE : 1 seul enregistrement =====
  s0 = stats();
  run = await runActive(panel, 2, yes);
  console.log("--- 3. nouveau DPE 12/09/2022 ---\n" + run.log);
  assert.deepStrictEqual(delta(s0, stats()), { creates: 1, updates: 0, uploads: 1, saves: 1 }, "1 seul enregistrement");
  d = lotDiags("590002")[0];
  assert.deepStrictEqual([d.ademe, d.dateFin, d.coutMin, d.docs.length], ["2222E0000001A", "11/09/2032", "900", 1], JSON.stringify(d));
  assert.ok(!/Enregistrer n°2/.test(run.log) && !/après 1er enregistrement/.test(run.log));
  assert.ok(/N° ADEME : attendu « 2222E0000001A » · avant « » · après « 2222E0000001A »/.test(run.log));
  assert.ok(/^OK — diagnostic \d+, ADEME 2222E0000001A — PDF joint/.test(await cell(2, "Saisie DPE")));

  // ===== 4. Serveur lent (2,5 s par réponse), ancien DPE =====
  await cfg("delayMs=2500");
  s0 = stats();
  run = await runActive(panel, 3, yes);
  console.log("--- 4. serveur lent ---\n" + run.log);
  await cfg("delayMs=0");
  d = lotDiags("590003");
  assert.strictEqual(d.length, 1, "une seule fiche (serveur lent)");
  assert.deepStrictEqual([d[0].conso, d[0].ges, d[0].dateFin, d[0].ademe, d[0].docs.length], ["224,00", "52,60", "31/12/2024", "2021L2000833E", 1], JSON.stringify(d[0]));
  assert.deepStrictEqual(delta(s0, stats()), { creates: 1, updates: 1, uploads: 1, saves: 2 });

  // ===== 5. Données invalides : rien n'est saisi =====
  const p5 = srv.stats.posts;
  await runActive(panel, 4);
  assert.ok(/^ERREUR — données de la ligne : ADEME_INVALIDE — « 2021N100167 »/.test(await cell(4, "Saisie DPE")), await cell(4, "Saisie DPE"));
  await runActive(panel, 5);
  c = await cell(5, "Saisie DPE");
  assert.ok(/VALEUR_TROP_LONGUE — Conso EP/.test(c) && /VALEUR_TROP_LONGUE — Coût min/.test(c), c);
  assert.strictEqual(srv.stats.posts, p5, "aucun POST pour des données invalides");

  // ===== 6. Écarts avant enregistrement =====
  await cfg("clearOnChange=dateDeb,comment");
  s0 = stats();
  run = await runActive(panel, 6);
  await cfg("clearOnChange=");
  c = await cell(6, "Saisie DPE");
  assert.ok(/^ERREUR — ECART_AVANT_ENREG — 2 écart\(s\) : Date début validité attendu « 12\/09\/2022 », relu « » ; Commentaires attendu « test écart », relu « »/.test(c), c);
  assert.ok(/nouvelle saisie \(essai 2\/2\)/.test(run.log));
  assert.strictEqual(delta(s0, stats()).saves, 0);

  // ===== 7. Écarts après enregistrement + fiche incomplète =====
  await cfg("corruptOnSave=ges");
  s0 = stats();
  run = await runActive(panel, 7, async () => { throw new Error("aucune confirmation attendue"); });
  console.log("--- 7. écarts après enregistrement ---\n" + run.log);
  await cfg("corruptOnSave=");
  c = await cell(7, "Saisie DPE");
  assert.ok(/^ERREUR — ECART_APRES_ENREG — 1 écart\(s\) : Émission GES attendu « 52,6 », relu « 53,60 » — upload annulé — fiche DPE (\d+) créée mais incomplète \(PDF manquant\) : http:\/\/\S+\/590007\/2\/\1$/.test(c), c);
  assert.deepStrictEqual(delta(s0, stats()), { creates: 1, updates: 0, uploads: 0, saves: 1 });

  // ===== 8. ADEME_CHAMP_ABSENT : DPE ≥ 07/2021 sans champ ADEME =====
  await cfg("ademeNever=1");
  s0 = stats();
  run = await runActive(panel, 8);
  console.log("--- 8. champ ADEME absent ---\n" + run.log);
  await cfg("ademeNever=0");
  c = await cell(8, "Saisie DPE");
  assert.ok(/^ERREUR — ADEME_CHAMP_ABSENT — champ N° ADEME \(#body_x_textIdDpe_x_textIdDpe\) absent après le rechargement de la date alors que le DPE est établi le 12\/09\/2022 \(≥ 01\/07\/2021\)/.test(c), c);
  assert.ok(/rien n'a été enregistré/.test(c));
  assert.strictEqual(delta(s0, stats()).saves, 0, "pas d'enregistrement");

  // ===== 9. Échec après l'Enregistrer n°1, puis reprise par le journal =====
  await cfg("refuseUpdate=1");
  s0 = stats();
  run = await runActive(panel, 9);
  console.log("--- 9a. Enregistrer n°2 refusé ---\n" + run.log);
  await cfg("refuseUpdate=0");
  c = await cell(9, "Saisie DPE");
  const m9 = c.match(/^ERREUR — SIGEO a refusé l'enregistrement n°2 : Mise à jour refusée \(test\)\. — fiche DPE (\d+) créée mais incomplète \(ADEME\/PDF manquants\) : (http:\/\/\S+\/590009\/2\/(\d+))$/);
  assert.ok(m9 && m9[1] === m9[3], c);
  assert.ok(run.log.includes(`fiche DPE ${m9[1]} créée mais incomplète (ADEME/PDF manquants) : ${m9[2]}`), "dans le journal aussi");
  assert.deepStrictEqual(delta(s0, stats()), { creates: 1, updates: 0, uploads: 0, saves: 2 });
  s0 = stats();
  run = await runActive(panel, 9, yes);
  console.log("--- 9b. reprise (journal) ---\n" + run.log);
  assert.ok(new RegExp(`DPE déjà enregistré pendant cette session \\(id ${m9[1]}\\) : réouverture`).test(run.log), run.log);
  assert.ok(/diagnostic déjà enregistré \(id \d+\) : modification/.test(run.log));
  assert.deepStrictEqual(delta(s0, stats()), { creates: 0, updates: 1, uploads: 1, saves: 1 }, "modification, jamais de second DPE");
  assert.strictEqual(lotDiags("590009").length, 1);
  assert.strictEqual(lotDiags("590009")[0].ademe, "2021N1001671X");
  assert.ok(/^OK — diagnostic \d+, ADEME 2021N1001671X — PDF joint/.test(await cell(9, "Saisie DPE")));

  // ===== 10. Reprise par la colonne résultat (journal vidé) =====
  await cfg("refuseUpdate=1");
  await runActive(panel, 10);
  await cfg("refuseUpdate=0");
  c = await cell(10, "Saisie DPE");
  const id10 = (c.match(/fiche DPE (\d+) créée mais incomplète/) || [])[1];
  assert.ok(id10, c);
  await panel.evaluate(() => document.querySelector(".scn-diag-forget").click());
  await panel.waitForFunction(() => /historique DPE vidé/.test(document.querySelector(".scn-diag-out").textContent));
  s0 = stats();
  run = await runActive(panel, 10, yes);
  console.log("--- 10. reprise (colonne résultat) ---\n" + run.log);
  assert.ok(run.log.includes(`fiche DPE ${id10} déjà créée (colonne résultat) : réouverture en modification`), run.log);
  assert.deepStrictEqual(delta(s0, stats()), { creates: 0, updates: 1, uploads: 1, saves: 1 });
  assert.strictEqual(lotDiags("590010").length, 1, "toujours une seule fiche");

  // ===== 11. POSTBACK_TIMEOUT =====
  await panel.evaluate(() => { document.querySelector(".scn-diag-timeout").value = "3000"; });
  await cfg("delayMs=6000");
  run = await runActive(panel, 11);
  await cfg("delayMs=0");
  await panel.evaluate(() => { document.querySelector(".scn-diag-timeout").value = "15000"; });
  assert.ok(/^ERREUR — POSTBACK_TIMEOUT — « Type » : rechargement non terminé après 3 s/.test(await cell(11, "Saisie DPE")), await cell(11, "Saisie DPE"));
  await tab.waitForTimeout(6500); // laisse arriver la réponse lente

  // ===== 12. Lettres : forcée seulement si SIGEO calcule autre chose =====
  run = await runActive(panel, 12, yes);
  console.log("--- 12. lettres forcées ---\n" + run.log);
  d = lotDiags("590012")[0];
  assert.ok(d, "diagnostic L13");
  assert.strictEqual(d.forceEp, "D", "EP : PDF D ≠ SIGEO E → forcée");
  assert.strictEqual(d.forceGes, "", "GES : PDF B = SIGEO B → non forcée");
  assert.ok(/lettre EP du DPE D ≠ lettre calculée par SIGEO E → forcée/.test(run.log), run.log);
  assert.ok(/^OK —/.test(await cell(12, "Saisie DPE")), await cell(12, "Saisie DPE"));

  // ===== 13. PDF visible seulement après un nouvel « Enregistrer » =====
  await cfg("uploadHidden=2"); // caché à l'affichage de l'upload et au 1er réenregistrement
  s0 = stats();
  run = await runActive(panel, 13, yes);
  console.log("--- 13. PDF affiché au 2e nouvel enregistrement ---\n" + run.log);
  await cfg("uploadHidden=0");
  assert.deepStrictEqual(delta(s0, stats()), { creates: 1, updates: 2, uploads: 1, saves: 3 }, "un seul upload, 2 Enregistrer de revérification");
  c = await cell(13, "Saisie DPE");
  assert.ok(/^OK — diagnostic \d+, ADEME 2222E0000001A — PDF joint \(DPE_2222E0000001A\.pdf\) ⚠ document « DPE » apparu après le 2e nouvel enregistrement \(2\/3\)$/.test(c), c);
  assert.ok(/upload non vérifié tout de suite : 0 → 0 document\(s\), ligne « DPE » absente → nouvel enregistrement de la fiche/.test(run.log), run.log);
  assert.ok(/Enregistrer recliqué \(1\/3\), sans nouvel upload/.test(run.log) && /Enregistrer recliqué \(2\/3\)/.test(run.log) && !/\(3\/3\)/.test(run.log));
  assert.ok(/⚠ document « DPE » apparu après le 2e nouvel enregistrement/.test(run.log), "avertissement journalisé");
  assert.ok(/✓ Étape 1 — SIGEO : saisir un DPE \+ PDF : .*1 avertissement\(s\) — \d+\.\d s/.test(run.log), "succès avec durée de l'étape");
  d = lotDiags("590013")[0];
  assert.strictEqual(d.docs.length, 1, "un seul document");

  // ===== 14. PDF jamais visible : erreur après 3 tentatives, relances sans doublon =====
  await cfg("uploadHidden=9");
  s0 = stats();
  run = await runActive(panel, 14, yes);
  console.log("--- 14a. PDF jamais visible ---\n" + run.log);
  await cfg("uploadHidden=0");
  assert.deepStrictEqual(delta(s0, stats()), { creates: 1, updates: 3, uploads: 1, saves: 4 }, "3 tentatives, un seul upload");
  c = await cell(14, "Saisie DPE");
  assert.ok(/^ERREUR — upload non vérifié après 3 tentatives d'enregistrement : 0 → 0 document\(s\), ligne « DPE » absente — PDF envoyé une seule fois, non ré-uploadé : vérifie la liste des documents dans SIGEO — fiche DPE (\d+) créée mais incomplète \(PDF envoyé mais non vérifié\) : http:\/\/\S+\/590014\/2\/\1 — \d+\.\d s$/.test(c) ||
    /^ERREUR — upload non vérifié après 3 tentatives d'enregistrement : 0 → 0 document\(s\), ligne « DPE » absente — PDF envoyé une seule fois, non ré-uploadé : vérifie la liste des documents dans SIGEO — fiche DPE (\d+) créée mais incomplète \(PDF envoyé mais non vérifié\) : http:\/\/\S+\/590014\/2\/\1$/.test(c), c);
  assert.ok(/Enregistrer recliqué \(3\/3\)/.test(run.log));
  s0 = stats();
  run = await runActive(panel, 14);
  console.log("--- 14b. relance : revérification sans nouvel upload ---\n" + run.log);
  assert.ok(/a déjà été envoyé sur cette fiche lors d'un passage précédent/.test(run.log), run.log);
  assert.ok(/pas de nouvel upload pour éviter un doublon/.test(await cell(14, "Saisie DPE")), await cell(14, "Saisie DPE"));
  assert.deepStrictEqual(delta(s0, stats()), { creates: 0, updates: 3, uploads: 0, saves: 3 }, "aucun upload, aucune fiche");
  s0 = stats();
  run = await runActive(panel, 14);
  console.log("--- 14c. relance : document enfin visible ---\n" + run.log);
  assert.deepStrictEqual(delta(s0, stats()), { creates: 0, updates: 1, uploads: 0, saves: 1 });
  c = await cell(14, "Saisie DPE");
  assert.ok(/^OK — diagnostic \d+, ADEME 2222E0000001A — PDF joint \(DPE_2222E0000001A\.pdf\), visible après 1 nouvel\(s\) enregistrement\(s\) ⚠ document « DPE » apparu après le 1er nouvel enregistrement \(1\/3\)/.test(c), c);
  assert.strictEqual(lotDiags("590014").length, 1, "une seule fiche");
  assert.strictEqual(lotDiags("590014")[0].docs.length, 1, "un seul document : jamais ré-uploadé");

  // ===== 15. N° ADEME « NC » + PDF joint automatiquement (ligne active, sans « Joindre ») =====
  await setOpt(".scn-diag-confirm", false);
  s0 = stats();
  run = await runActive(panel, 15, async () => { throw new Error("aucune confirmation attendue"); });
  console.log("--- 15. ADEME NC, jonction automatique ---\n" + run.log);
  await setOpt(".scn-diag-confirm", true);
  assert.deepStrictEqual(delta(s0, stats()), { creates: 1, updates: 0, uploads: 1, saves: 1 }, "1 enregistrement, pas d'Enregistrer n°2, PDF joint");
  d = lotDiags("590015")[0];
  assert.deepStrictEqual([d.ademe, d.docs.length, d.docs[0].name], ["", 1, "DPE_2021N1001671X.pdf"]);
  c = await cell(15, "Saisie DPE");
  assert.ok(/^OK — diagnostic \d+ — PDF joint \(DPE_2021N1001671X\.pdf\) ⚠ N° ADEME non communiqué \(« NC »\) : non saisi/.test(c), c);
  assert.ok(!/ADEME_INVALIDE/.test(c) && !/après 1er enregistrement/.test(run.log));

  assert.strictEqual(srv.stats.enterSaves, 0, "jamais de touche Entrée");
  assert.strictEqual(srv.stats.saveCloses, 0, "jamais « Enreg. et fermer »");
  assert.deepStrictEqual(errors, [], "erreurs JS du panneau");
  console.log("\nstats serveur :", JSON.stringify(srv.stats));
  console.log("\n✓ cas réels, enregistrement en deux temps et phases A→E OK");
  await browser.close();
  srv.server.close();
})().catch((e) => { console.error(e); process.exit(1); });
