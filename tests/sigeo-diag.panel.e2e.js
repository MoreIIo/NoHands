// Test d'intégration : étape de scénario « SIGEO : saisir un DPE ».
// Le vrai sidepanel.html est chargé ; chrome.scripting / chrome.tabs sont
// redirigés vers un onglet ouvert sur le faux SIGEO diag_manage
// (tests/fixtures/sigeo-diag-server.js), chrome.storage.session est simulé.
//   NODE_PATH=$(npm root -g) node tests/sigeo-diag.panel.e2e.js
// Critères vérifiés :
//   1. DPE ≥ 07/2021 (ADEME + coûts) saisi de bout en bout, PDF joint « DPE », compteur à 1 ;
//   2. DPE d'avant 07/2021 : N° ADEME saisi aussi (champ présent quelle que soit la date) ;
//   3. fichier absent / ambigu / mauvais ADEME / pas un PDF → arrêt AVANT toute saisie ;
//   4. refus serveur après Enregistrer → arrêt, texte de #msgbody dans le journal ;
//   5. rechargements (pendant la confirmation, pendant la saisie, relance, page /0)
//      → ni double enregistrement ni double upload.
const path = require("path");
const fs = require("fs");
const os = require("os");
const assert = require("assert");
const { chromium } = require("playwright");
const mock = require("./fixtures/sigeo-diag-server.js");
const H = require("./fixtures/diag-panel-harness.js");
const { makePdf, COLS, row } = H;

(async () => {
  // — Dossier des DPE
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "nh-dpe-"));
  const files = {
    "DPE_2475E1234567X.pdf": makePdf(["Diagnostic de performance energetique", "N ADEME : 2475E1234567X"]),
    "IM0393600_12_DPE_2019-05-12.pdf": makePdf(["DPE 2019", "Numero d'enregistrement ADEME : 1975V1000123A"]),
    "A_2475E7654321Y.pdf": makePdf(["N ADEME : 2475E7654321Y"]),
    "B_2475E7654321Y.pdf": makePdf(["N ADEME : 2475E7654321Y"]),
    "DPE_2475E5555555C.pdf": makePdf(["N ADEME : 2475E5555555C"]),
    "WRONG_2475E2222222W.pdf": makePdf(["N ADEME : 2475E2222222W"]),
    "IM0393700_14_DPE_2023-01-10.pdf": makePdf(["Diagnostic de performance energetique", "N ADEME : 2375 E 3333333 A"]),
    "SCAN_lot15.pdf": makePdf([]),
    "notpdf.pdf": Buffer.from("<html>session expiree</html>"),
    "DPE_2475E6666666D.pdf": makePdf(["N ADEME : 2475E6666666D"]),
    "DPE_2475E7777777E.pdf": makePdf(["N ADEME : 2475E7777777E"])
  };
  for (const [n, b] of Object.entries(files)) fs.writeFileSync(path.join(dir, n), b);

  const srv = await mock.start();
  const { browser, tab, panel, errors } = await H.openPanel(chromium, srv.url("393664", 0));

  await panel.evaluate(([cols, rows]) => {
    state.headerMode = "yes";
    state.rows = [cols].concat(rows);
    renderColumns();
  }, [COLS, [
    // L2 : DPE 2024 complet (ligne active, PDF choisi)
    row({ type_local: "habitation", code_fournisseur: "448957", date_etabli: "15/03/2024", numero_ademe: "2475E1234567X", conso_ep: "320", emission_ges: "9",
      altitude_800m: "non", nouvelles_valeurs: "oui", cout_min: "1 240", cout_max: "1680", date_prix: "01/01/2021", commentaire: "saisie auto", lot: "393664" }),
    // L3 : DPE 2019, champ ADEME absent → avertissement ; fichier par nom exact
    row({ code_fournisseur: "528825", date_etabli: "12/05/2019", numero_ademe: "1975V1000123A", conso_ep: "245,5", emission_ges: "12",
      fichier_dpe: "IM0393600_12_DPE_2019-05-12.pdf", lot: "393600" }),
    // L4 : deux fichiers contiennent l'ADEME → ambigu
    row({ code_fournisseur: "448957", date_etabli: "02/02/2024", numero_ademe: "2475E7654321Y", conso_ep: "150", emission_ges: "20", cout_min: "800", cout_max: "1100", date_prix: "01/01/2021", lot: "393604" }),
    // L5 : fournisseur inconnu → refus serveur à l'enregistrement
    row({ code_fournisseur: "999999", date_etabli: "05/05/2024", numero_ademe: "2475E5555555C", conso_ep: "200", emission_ges: "30", cout_min: "900", cout_max: "1200", date_prix: "01/01/2021", lot: "393605" }),
    // L6 : le fichier nommé porte un autre N° ADEME
    row({ code_fournisseur: "448957", date_etabli: "06/06/2024", numero_ademe: "2475E1111111Z", conso_ep: "210", emission_ges: "31", fichier_dpe: "WRONG_2475E2222222W.pdf", lot: "393606" }),
    // L7 : ADEME absent du nom mais présent dans le texte du PDF
    row({ type_local: "hors habitation", code_fournisseur: "410962", date_etabli: "10/01/2023", numero_ademe: "2375E3333333A", conso_ep: "99", emission_ges: "4",
      cout_min: "500", cout_max: "700", date_prix: "01/01/2021", fichier_dpe: "IM0393700_14_DPE_2023-01-10", lot: "393607" }),
    // L8 : PDF scanné (aucun texte) → refusé en boucle sans confirmation
    row({ code_fournisseur: "448957", date_etabli: "08/08/2024", numero_ademe: "2475E4444444B", conso_ep: "180", emission_ges: "25", fichier_dpe: "SCAN_lot15.pdf", lot: "393608" }),
    // L9 : « PDF » qui n'en est pas un
    row({ code_fournisseur: "448957", date_etabli: "09/09/2024", numero_ademe: "2475E9999999F", conso_ep: "180", emission_ges: "25", fichier_dpe: "notpdf.pdf", lot: "393609" }),
    // L10 : rechargement pendant la confirmation
    row({ code_fournisseur: "448957", date_etabli: "10/10/2024", numero_ademe: "2475E6666666D", conso_ep: "260", emission_ges: "40", cout_min: "1000", cout_max: "1500", date_prix: "01/01/2024", lot: "393610" }),
    // L11 : rechargement de la page pendant la saisie
    row({ code_fournisseur: "448957", date_etabli: "11/11/2024", numero_ademe: "2475E7777777E", conso_ep: "270", emission_ges: "41", cout_min: "1100", cout_max: "1600", date_prix: "01/01/2024",
      resultat: "conforme", lot: "393611" })
  ]]);

  // — Étape ajoutée par « Modèle DPE »
  await panel.evaluate(() => { $("clearStepsBtn") && (document.getElementById("scenarioSteps").innerHTML = ""); $("addDpeTemplateBtn").click(); });
  let steps = await panel.evaluate(() => getScenarioSteps());
  assert.strictEqual(steps.length, 1);
  assert.strictEqual(steps[0].type, "sigeoDpe");
  assert.strictEqual(steps[0].diagCols.code_fournisseur, "{code_fournisseur}");
  assert.strictEqual(steps[0].diagAttach, true);
  assert.strictEqual(steps[0].diagConfirm, true);
  assert.strictEqual(steps[0].diagResultCol, "Saisie DPE");
  const label = await panel.evaluate(() => scnStepLabel(getScenarioSteps()[0]));
  assert.strictEqual(label, "SIGEO : saisir un DPE + PDF");
  const warn0 = await panel.evaluate(() => $("scnWarnings").innerText);
  assert.ok(/aucun PDF ni dossier/.test(warn0), "avertissement sans fichier : " + warn0);

  const runActive = (idx, onConfirm) => H.runActive(panel, idx, onConfirm);
  const cell = (r, col) => panel.evaluate(([r, c]) => getCellByIndex(state.rows[r], colIndexByName(c)), [r, col]);
  const lotDiags = (lot) => srv.diags().filter((d) => d.lot === lot);

  // ===== 1. Ligne active, DPE 2024 : de bout en bout, PDF choisi par glisser/sélection =====
  await panel.setInputFiles(".scn-diag-unitfile", path.join(dir, "DPE_2475E1234567X.pdf"));
  assert.ok(/DPE_2475E1234567X\.pdf/.test(await panel.textContent(".scn-diag-unitname")));
  let run = await runActive(1, async (body) => {
    assert.ok(body.includes("2475E1234567X (confirmé par SIGEO)"), "confirmation : " + body);
    assert.ok(body.includes("DPE_2475E1234567X.pdf"), body);
    assert.ok(/393664/.test(body), body);
    await panel.click("#scnConfirmYes");
  });
  console.log("--- 1. ligne active ---\n" + run.log);
  assert.ok(run.confirmed, "confirmation demandée en unitaire");
  let d = lotDiags("393664");
  assert.strictEqual(d.length, 1, "un seul diagnostic créé");
  d = d[0];
  assert.strictEqual(d.type, "11");
  assert.strictEqual(d.fou, "448957");
  assert.strictEqual(d.dateReal, "15/03/2024");
  assert.strictEqual(d.dateFin, "14/03/2034", "validité calculée par la page");
  assert.strictEqual(d.ademe, "2475E1234567X");
  assert.strictEqual(d.conso, "320,00");
  assert.strictEqual(d.ges, "9,00");
  assert.strictEqual(d.coutMin, "1240", "coûts saisis (bloc apparu après la conso)");
  assert.strictEqual(d.coutMax, "1680");
  assert.strictEqual(d.datePrix, "01/01/2021");
  assert.strictEqual(d.nouvelles, true, "case « nouvelles valeurs » cochée");
  assert.strictEqual(d.altitude, false);
  assert.strictEqual(d.comment, "saisie auto");
  assert.deepStrictEqual(d.docs.map((x) => [x.desc, x.name, x.pdf]), [["DPE", "DPE_2475E1234567X.pdf", true]], "PDF joint");
  assert.strictEqual(srv.stats.enterSaves, 0, "jamais de touche Entrée");
  assert.strictEqual(srv.stats.saveCloses, 0, "jamais « Enreg. et fermer »");
  assert.strictEqual(srv.stats.saves, 1);
  assert.ok(/^OK — diagnostic 1147729, ADEME 2475E1234567X — PDF joint \(DPE_2475E1234567X\.pdf\)/.test(await cell(1, "Saisie DPE")), await cell(1, "Saisie DPE"));
  assert.ok(/1 document\(s\)/.test(run.log) && /contrôle avant enregistrement : \d+ valeur\(s\) conforme/.test(run.log), "journal");
  assert.ok(/contrôle après enregistrement : \d+ valeur\(s\) conforme/.test(run.log), "contrôle après enregistrement");
  assert.ok(/Nouvelles valeurs du DPE : attendu « coché » · avant « décoché » · après « coché »/.test(run.log), "case journalisée");
  assert.ok(/\/diag_manage\/393664\/10\/1147729$/.test(tab.url()), "popup laissée ouverte sur la fiche (form action → id) : " + tab.url());

  // ===== 5a. Relance sur la même fiche : ni enregistrement ni upload =====
  await panel.evaluate(() => { setCellByIndex(state.rows[1], colIndexByName("Saisie DPE"), ""); });
  run = await runActive(1, async () => { throw new Error("aucune confirmation attendue"); });
  console.log("--- 5a. relance ---\n" + run.log);
  assert.strictEqual(srv.stats.saves, 1, "pas de nouvel enregistrement");
  assert.strictEqual(srv.stats.uploads, 1, "pas de nouvel upload");
  assert.ok(/déjà joint/.test(await cell(1, "Saisie DPE")), await cell(1, "Saisie DPE"));

  // ===== « sauter les lignes OK » =====
  run = await runActive(1);
  assert.ok(/ligne déjà saisie/.test(run.log), run.log);

  // ===== 5b. Page de création /0 rouverte pour un DPE déjà enregistré =====
  await tab.goto(srv.url("393664", 0));
  await panel.evaluate(() => { setCellByIndex(state.rows[1], colIndexByName("Saisie DPE"), ""); });
  run = await runActive(1);
  console.log("--- 5b. page /0 ---\n" + run.log);
  assert.ok(/déjà été enregistré pendant cette session \(fiche 1147729\)/.test(run.log), run.log);
  assert.strictEqual(lotDiags("393664").length, 1, "pas de doublon");
  assert.ok(/^ERREUR — ce DPE a déjà été enregistré/.test(await cell(1, "Saisie DPE")));

  // ===== Dossier des DPE (repli <input webkitdirectory>) =====
  await panel.evaluate(() => diagSetUnitFile(null));
  await panel.setInputFiles(".scn-diag-folderinput", dir);
  assert.ok(/11 fichier/.test(await panel.textContent(".scn-diag-foldername")), await panel.textContent(".scn-diag-foldername"));

  // ===== 5c. Rechargement de la page pendant la confirmation (après Enregistrer) =====
  await tab.goto(srv.url("393610", 0));
  await fetch(srv.origin + "/__config?redirect=1");
  run = await runActive(9, async () => {
    await tab.reload();
    await tab.waitForLoadState("load");
    await panel.click("#scnConfirmYes");
  });
  console.log("--- 5c. rechargement pendant la confirmation ---\n" + run.log);
  await fetch(srv.origin + "/__config?redirect=0");
  d = lotDiags("393610");
  assert.strictEqual(d.length, 1, "un seul diagnostic");
  assert.strictEqual(d[0].docs.length, 1, "un seul upload");
  assert.ok(/\/diag_manage\/393610\/10\/1147730$/.test(tab.url()), "redirection après enregistrement : " + tab.url());
  assert.ok(/^OK —/.test(await cell(9, "Saisie DPE")), await cell(9, "Saisie DPE"));

  // ===== Boucle : 2, 3 et 4 =====
  await panel.evaluate((u) => {
    const st = document.querySelector(".scn-step");
    st.querySelector(".scn-diag-url").value = u;
    st.querySelector(".scn-diag-confirm").checked = false;
    $("scnStartRow").value = "3";
    $("scnEndRow").value = "9";
    $("scnRowDelayMs").value = "0";
    $("scnLog").innerHTML = "";
    $("runScenarioLoopBtn").click();
  }, srv.origin + "/popup.aspx/fr/sig_imm/diag_manage/{lot}/10/0");
  await panel.waitForFunction(() => !scnRunning && /Boucle terminée/.test($("scnLog").textContent), null, { timeout: 180000 });
  let log = await panel.evaluate(() => $("scnLog").innerText);
  console.log("--- boucle L3-L9 ---\n" + log);
  const res = {};
  for (let r = 2; r <= 8; r++) res[r + 1] = await cell(r, "Saisie DPE");
  console.log(res);
  // 2. ancien DPE (2019) : N° ADEME saisi aussi, validité 2018–06/2021 → 31/12/2024
  assert.ok(/^OK — diagnostic \d+, ADEME 1975V1000123A — PDF joint \(IM0393600_12_DPE_2019-05-12\.pdf\)/.test(res[3]), res[3]);
  d = lotDiags("393600")[0];
  assert.ok(d && d.ademe === "1975V1000123A" && d.dateFin === "31/12/2024" && d.docs.length === 1, JSON.stringify(d));
  // 3. fichiers : ambigu / autre ADEME / scan / pas un PDF → arrêt avant toute saisie
  assert.ok(/^ERREUR — fichier DPE ambigu : 2 fichiers contiennent le N° ADEME 2475E7654321Y .* rien n'a été saisi/.test(res[4]), res[4]);
  assert.ok(/^ERREUR — fichier DPE « WRONG_2475E2222222W\.pdf » : le nom du fichier contient un autre N° ADEME \(2475E2222222W\)/.test(res[6]), res[6]);
  assert.ok(/^ERREUR — fichier DPE « SCAN_lot15\.pdf » : N° ADEME 2475E4444444B introuvable .* refusé en boucle sans confirmation/.test(res[8]), res[8]);
  assert.ok(/^ERREUR — fichier DPE « notpdf\.pdf » n'est pas un PDF/.test(res[9]), res[9]);
  for (const lot of ["393604", "393606", "393608", "393609"]) assert.strictEqual(lotDiags(lot).length, 0, "rien enregistré pour " + lot);
  // 4. refus serveur : texte de #msgbody dans le journal et la colonne
  assert.ok(/^ERREUR — SIGEO a refusé l'enregistrement : Le fournisseur 999999 n'existe pas\./.test(res[5]), res[5]);
  assert.ok(/SIGEO a refusé l'enregistrement : Le fournisseur 999999 n'existe pas/.test(log));
  assert.strictEqual(lotDiags("393605").length, 0);
  // ADEME lu dans le texte du PDF (pas dans le nom), type hors habitation
  assert.ok(/^OK — diagnostic \d+, ADEME 2375E3333333A — PDF joint \(IM0393700_14_DPE_2023-01-10\.pdf\)/.test(res[7]), res[7]);
  assert.ok(/ADEME dans le texte du PDF/.test(log));
  assert.strictEqual(lotDiags("393607")[0].type, "15");
  assert.ok(/3 OK|2 OK/.test(log));

  // ===== 5d. Rechargement complet pendant la saisie (après la date) =====
  let reloaded = false;
  const watcher = setInterval(async () => {
    if (reloaded) return;
    try {
      const fin = /393611\/10\/0$/.test(tab.url()) ? await tab.$eval("#fake_txtDateFinValid_x_txtDateFinValid", (e) => e.textContent.trim()) : "";
      if (fin) {
        reloaded = true;
        await tab.goto(srv.url("393611", 0)); // fiche vierge : tout est à refaire
      }
    } catch (_) { /* navigation en cours */ }
  }, 50);
  await panel.evaluate(() => {
    $("scnStartRow").value = "11"; $("scnEndRow").value = "11"; $("scnLog").innerHTML = "";
    $("runScenarioLoopBtn").click();
  });
  await panel.waitForFunction(() => !scnRunning && /Boucle terminée/.test($("scnLog").textContent), null, { timeout: 120000 });
  clearInterval(watcher);
  log = await panel.evaluate(() => $("scnLog").innerText);
  console.log("--- 5d. rechargement pendant la saisie ---\n" + log);
  assert.ok(reloaded, "la page a bien été rechargée pendant la saisie");
  d = lotDiags("393611");
  assert.strictEqual(d.length, 1, "un seul diagnostic malgré le rechargement");
  assert.strictEqual(d[0].docs.length, 1);
  assert.strictEqual(d[0].resultat, "conforme");
  assert.strictEqual(d[0].ademe, "2475E7777777E");
  assert.ok(/^OK —/.test(await cell(10, "Saisie DPE")), await cell(10, "Saisie DPE"));

  // ===== 5e. Relance en boucle avec l'URL /0 : le journal rouvre les fiches existantes =====
  const before = { creates: srv.stats.creates, uploads: srv.stats.uploads };
  await panel.evaluate(() => {
    document.querySelector(".scn-diag-skipdone").checked = false;
    $("scnStartRow").value = "2"; $("scnEndRow").value = "3"; $("scnLog").innerHTML = "";
    $("runScenarioLoopBtn").click();
  });
  await panel.waitForFunction(() => !scnRunning && /Boucle terminée/.test($("scnLog").textContent), null, { timeout: 120000 });
  log = await panel.evaluate(() => $("scnLog").innerText);
  console.log("--- 5e. relance /0 ---\n" + log);
  assert.ok(/réouverture de la fiche au lieu d'une création/.test(log), "journal utilisé");
  assert.strictEqual(srv.stats.creates, before.creates, "aucune création");
  assert.strictEqual(srv.stats.uploads, before.uploads, "aucun upload");
  assert.ok(/déjà joint/.test(await cell(1, "Saisie DPE")) && /déjà joint/.test(await cell(2, "Saisie DPE")));

  // ===== Debug : relevé des champs =====
  await panel.evaluate(() => document.querySelector(".scn-diag-dump").click());
  await panel.waitForFunction(() => /éléments relevés/.test(document.querySelector(".scn-diag-out").textContent), null, { timeout: 10000 });
  const dump = await panel.evaluate(() => diagDumps[diagDumps.length - 1]);
  assert.ok(dump.fields.some((f) => f.id === "body_x_proxyFou_x_txt_x_txt" && /press\(cmdSave\)/.test(f.onkeydown || "")), "handlers relevés");
  assert.ok(dump.fields.some((f) => f.id === "body_x_txtConsoE_x_txtConsoE" && f.label === "Conso énergie primaire"), "libellés relevés");

  // Profil : réglages de l'étape sauvegardés / restaurés
  const prof = await panel.evaluate(() => { const c = buildProfileConfig(); applyProfileConfig(c); return getScenarioSteps()[0]; });
  assert.strictEqual(prof.type, "sigeoDpe");
  assert.strictEqual(prof.diagConfirm, false);
  assert.ok(/\{lot\}\/10\/0$/.test(prof.diagUrl));

  assert.strictEqual(srv.stats.enterSaves, 0, "jamais de touche Entrée");
  assert.strictEqual(srv.stats.saveCloses, 0, "jamais « Enreg. et fermer »");
  assert.deepStrictEqual(errors, [], "erreurs JS du panneau");
  console.log("\nstats serveur :", JSON.stringify(srv.stats));
  console.log("\n✓ étape « SIGEO : saisir un DPE » OK");
  await browser.close();
  srv.server.close();
})().catch((e) => { console.error(e); process.exit(1); });
