// Test d'intégration : cas réel du 23/11/2026 et phases A→E de l'étape
// « SIGEO : saisir un DPE » sur un faux SIGEO conforme à la vraie page
// (fournisseur à postback, change synthétique sans effet sur « Établi le »,
// ADEME body_x_textIdDpe_x_textIdDpe toujours en saisie après enregistrement).
//   NODE_PATH=$(npm root -g) node tests/sigeo-diag.realcase.e2e.js
// Critères :
//   1. simulation (dryRun) du cas réel : rien d'enregistré, la page affiche
//      224,00 / 52,60 / 31/12/2024 / 2021L2000833E ;
//   2. enregistrement du cas réel : mêmes valeurs relues après enregistrement,
//      PDF joint ; journal « attendu / avant / après » et « ↻ postback (…) » ;
//   3. serveur lent (2,5 s > fenêtre de calme de 1,5 s) : la navigation en
//      attente est détectée, rien n'est écrit dans la page condamnée ;
//   4. ADEME_INVALIDE et VALEUR_TROP_LONGUE avant toute saisie ;
//   5. ECART_AVANT_ENREG (2 essais puis abandon, TOUS les écarts listés) ;
//   6. ECART_APRES_ENREG (tous les écarts, upload annulé) ;
//   7. POSTBACK_TIMEOUT ;
//   8. lettre du DPE ≠ lettre SIGEO → « Forcer la lettre » (et seulement là).
const path = require("path");
const fs = require("fs");
const os = require("os");
const assert = require("assert");
const { chromium } = require("playwright");
const mock = require("./fixtures/sigeo-diag-server.js");
const { makePdf, COLS, row, openPanel, runActive } = require("./fixtures/diag-panel-harness.js");

const REAL = { code_fournisseur: "448957", date_etabli: "08/10/2020", numero_ademe: "2021L2000833E", conso_ep: "224", emission_ges: "52,6" };

(async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "nh-dpe-real-"));
  fs.writeFileSync(path.join(dir, "DPE_2021L2000833E.pdf"), makePdf(["Diagnostic de performance energetique", "N ADEME : 2021L2000833E", "FU ACADE Diagnostic"]));
  fs.writeFileSync(path.join(dir, "DPE_2475E1234567X.pdf"), makePdf(["N ADEME : 2475E1234567X"]));

  const srv = await mock.start();
  const cfg = (q) => fetch(srv.origin + "/__config?" + q);
  const { browser, tab, panel, errors } = await openPanel(chromium, srv.url("448001", 0, 3));
  const cell = (r, col) => panel.evaluate(([r, c]) => getCellByIndex(state.rows[r], colIndexByName(c)), [r, col]);
  const lotDiags = (lot) => srv.diags().filter((d) => d.lot === lot);
  const pageVals = () => tab.evaluate(() => {
    const v = (id) => { const e = document.getElementById(id); return e ? e.value : null; };
    return {
      conso: v("body_x_txtConsoE_x_txtConsoE"), ges: v("body_x_txtEmission_x_txtEmission"),
      fin: v("body_x_txtDateFinValid_x_txtDateFinValid"), ademe: v("body_x_textIdDpe_x_textIdDpe"),
      fou: v("body_x_proxyFou_x_txt_x_txt"), date: v("body_x_txtDateReal_x_txtDateReal")
    };
  });

  await panel.evaluate(([cols, rows]) => {
    state.headerMode = "yes";
    state.rows = [cols].concat(rows);
    renderColumns();
  }, [COLS, [
    /* L2 */ row(Object.assign({ lot: "448001" }, REAL)),
    /* L3 */ row(Object.assign({ lot: "448002" }, REAL)),
    /* L4 */ row(Object.assign({}, REAL, { lot: "448003", numero_ademe: "2021L200083" })),
    /* L5 */ row(Object.assign({}, REAL, { lot: "448004", conso_ep: "1234,56", cout_min: "1240,75" })),
    /* L6 */ row(Object.assign({}, REAL, { lot: "448005", date_debut_validite: "08/10/2020", commentaire: "test écart" })),
    /* L7 */ row(Object.assign({}, REAL, { lot: "448006" })),
    /* L8 */ row(Object.assign({}, REAL, { lot: "448007" })),
    /* L9 */ row({ code_fournisseur: "448957", date_etabli: "15/03/2024", numero_ademe: "2475E1234567X", conso_ep: "320", emission_ges: "9",
      lettre_ep_forcee: "D", lettre_ges_forcee: "B", cout_min: "1240", cout_max: "1680", date_prix: "01/01/2021", lot: "448008" })
  ]]);

  await panel.evaluate((u) => {
    document.getElementById("scenarioSteps").innerHTML = "";
    $("addDpeTemplateBtn").click();
    const st = document.querySelector(".scn-step");
    st.querySelector(".scn-diag-url").value = u;
  }, srv.origin + "/popup.aspx/fr/sig_imm/diag_manage/{lot}/3/0");
  await panel.setInputFiles(".scn-diag-folderinput", dir);
  const setOpt = (cls, checked) => panel.evaluate(([c, v]) => { document.querySelector(c).checked = v; }, [cls, checked]);
  assert.strictEqual(await panel.evaluate(() => getScenarioSteps()[0].diagForceLetters), true, "option « forcer la lettre » cochée par défaut");

  // ===== 1. Simulation (dryRun) du cas réel =====
  await setOpt(".scn-diag-simulate", true);
  let run = await runActive(panel, 1);
  console.log("--- 1. simulation du cas réel ---\n" + run.log);
  assert.ok(/^SIMULATION OK — /.test(await cell(1, "Saisie DPE")), await cell(1, "Saisie DPE"));
  assert.strictEqual(srv.stats.saves, 0, "rien d'enregistré en simulation");
  let pv = await pageVals();
  assert.deepStrictEqual([pv.conso, pv.ges, pv.fin, pv.ademe, pv.fou, pv.date], ["224,00", "52,60", "31/12/2024", "2021L2000833E", "448957", "08/10/2020"], JSON.stringify(pv));
  assert.ok(/Conso EP : attendu « 224 » · avant « » · après « 224,00 »/.test(run.log), "journal attendu/avant/après");
  assert.ok(/↻ postback \(Conso EP\)/.test(run.log) && /↻ postback \(Émission GES\)/.test(run.log) && /↻ postback \(Fournisseur\)/.test(run.log), "postbacks journalisés");
  assert.ok(/Établi le : pas de postback après le change synthétique → appel direct du handler/.test(run.log), "repli du handler « Établi le »");
  assert.ok(/↻ postback \(Établi le\)/.test(run.log));
  assert.ok(/Valable jusqu'au 31\/12\/2024/.test(await cell(1, "Saisie DPE")), await cell(1, "Saisie DPE"));
  const order = ["Type :", "Fournisseur :", "Établi le :", "Conso EP :", "Émission GES :", "N° ADEME :"].map((k) => run.log.indexOf(k));
  assert.ok(order.every((x, i) => x >= 0 && (i === 0 || x > order[i - 1])), "ordre A puis B : " + order);

  // ===== 2. Enregistrement du cas réel =====
  await setOpt(".scn-diag-simulate", false);
  run = await runActive(panel, 1, async (body) => {
    assert.ok(body.includes("2021L2000833E"), body);
    await panel.click("#scnConfirmYes");
  });
  console.log("--- 2. enregistrement du cas réel ---\n" + run.log);
  let d = lotDiags("448001");
  assert.strictEqual(d.length, 1, "un diagnostic");
  d = d[0];
  assert.deepStrictEqual([d.fou, d.dateReal, d.conso, d.ges, d.dateFin, d.ademe], ["448957", "08/10/2020", "224,00", "52,60", "31/12/2024", "2021L2000833E"], JSON.stringify(d));
  assert.deepStrictEqual(d.docs.map((x) => [x.desc, x.name]), [["DPE", "DPE_2021L2000833E.pdf"]]);
  assert.ok(/contrôle après enregistrement : \d+ valeur\(s\) conforme\(s\)/.test(run.log), run.log);
  assert.ok(/^OK — diagnostic \d+, ADEME 2021L2000833E — PDF joint \(DPE_2021L2000833E\.pdf\)/.test(await cell(1, "Saisie DPE")), await cell(1, "Saisie DPE"));
  pv = await pageVals();
  assert.deepStrictEqual([pv.conso, pv.ges, pv.fin, pv.ademe], ["224,00", "52,60", "31/12/2024", "2021L2000833E"], "page après enregistrement : " + JSON.stringify(pv));

  // ===== 3. Serveur lent : 2,5 s par réponse =====
  await cfg("delayMs=2500");
  const posts0 = srv.stats.posts;
  run = await runActive(panel, 2, async () => { await panel.click("#scnConfirmYes"); });
  console.log("--- 3. serveur lent ---\n" + run.log);
  await cfg("delayMs=0");
  d = lotDiags("448002");
  assert.strictEqual(d.length, 1, "un diagnostic (serveur lent)");
  assert.deepStrictEqual([d[0].conso, d[0].ges, d[0].dateFin, d[0].ademe, d[0].docs.length], ["224,00", "52,60", "31/12/2024", "2021L2000833E", 1], JSON.stringify(d[0]));
  assert.ok(srv.stats.posts - posts0 <= 8, "pas de postbacks en rafale : " + (srv.stats.posts - posts0));

  // ===== 4. Données invalides : rien n'est saisi =====
  const p4 = srv.stats.posts;
  run = await runActive(panel, 3);
  assert.ok(/^ERREUR — données de la ligne : ADEME_INVALIDE — « 2021L200083 »/.test(await cell(3, "Saisie DPE")), await cell(3, "Saisie DPE"));
  run = await runActive(panel, 4);
  const c5 = await cell(4, "Saisie DPE");
  assert.ok(/VALEUR_TROP_LONGUE — Conso EP/.test(c5) && /VALEUR_TROP_LONGUE — Coût min/.test(c5), c5);
  assert.strictEqual(srv.stats.posts, p4, "aucun POST pour des données invalides");

  // ===== 5. Écarts avant enregistrement (valeurs refusées par la page) =====
  await cfg("clearOnChange=dateDeb,comment");
  const s5 = srv.stats.saves;
  run = await runActive(panel, 5);
  console.log("--- 5. écarts avant enregistrement ---\n" + run.log);
  await cfg("clearOnChange=");
  const c6 = await cell(5, "Saisie DPE");
  assert.ok(/^ERREUR — ECART_AVANT_ENREG — 2 écart\(s\) : Date début validité attendu « 08\/10\/2020 », relu « » ; Commentaires attendu « test écart », relu « »/.test(c6), c6);
  assert.ok(/contrôle avant enregistrement : 2 écart\(s\) → nouvelle saisie \(essai 2\/2\)/.test(run.log), run.log);
  assert.strictEqual(srv.stats.saves, s5, "pas d'enregistrement avec des écarts");

  // ===== 6. Écarts après enregistrement =====
  await cfg("corruptOnSave=ademe,ges");
  const u6 = srv.stats.uploads;
  run = await runActive(panel, 6, async () => { throw new Error("aucune confirmation attendue"); });
  console.log("--- 6. écarts après enregistrement ---\n" + run.log);
  await cfg("corruptOnSave=");
  const c7 = await cell(6, "Saisie DPE");
  assert.ok(/^ERREUR — ECART_APRES_ENREG — 2 écart\(s\) : Émission GES attendu « 52,6 », relu « 53,60 » ; N° ADEME attendu « 2021L2000833E », relu « » — upload annulé/.test(c7), c7);
  assert.strictEqual(srv.stats.uploads, u6, "pas d'upload");

  // ===== 7. POSTBACK_TIMEOUT =====
  await panel.evaluate(() => { document.querySelector(".scn-diag-timeout").value = "3000"; });
  await cfg("delayMs=6000");
  run = await runActive(panel, 7);
  console.log("--- 7. délai dépassé ---\n" + run.log);
  await cfg("delayMs=0");
  await panel.evaluate(() => { document.querySelector(".scn-diag-timeout").value = "15000"; });
  assert.ok(/^ERREUR — POSTBACK_TIMEOUT — « Type » : rechargement non terminé après 3 s/.test(await cell(7, "Saisie DPE")), await cell(7, "Saisie DPE"));
  await tab.waitForTimeout(6500); // laisse arriver la réponse lente

  // ===== 8. Lettres : forcée seulement si SIGEO calcule autre chose =====
  run = await runActive(panel, 8, async () => { await panel.click("#scnConfirmYes"); });
  console.log("--- 8. lettres forcées ---\n" + run.log);
  d = lotDiags("448008")[0];
  assert.ok(d, "diagnostic L9");
  assert.strictEqual(d.forceEp, "D", "EP : PDF D ≠ SIGEO E → forcée");
  assert.strictEqual(d.forceGes, "", "GES : PDF B = SIGEO B → non forcée");
  assert.ok(/lettre EP du DPE D ≠ lettre calculée par SIGEO E → forcée/.test(run.log), run.log);
  assert.ok(/^OK —/.test(await cell(8, "Saisie DPE")), await cell(8, "Saisie DPE"));

  assert.strictEqual(srv.stats.enterSaves, 0, "jamais de touche Entrée");
  assert.strictEqual(srv.stats.saveCloses, 0, "jamais « Enreg. et fermer »");
  assert.deepStrictEqual(errors, [], "erreurs JS du panneau");
  console.log("\nstats serveur :", JSON.stringify(srv.stats));
  console.log("\n✓ cas réel et phases A→E OK");
  await browser.close();
  srv.server.close();
})().catch((e) => { console.error(e); process.exit(1); });
