// Faux SIGEO « diag_manage » (popup Diagnostic) pour les tests de l'étape
// « SIGEO : saisir un DPE ». Reproduit ce qui compte pour l'automatisation
// (ids réels relevés sur sigeo.evoriel.net le 23/11/2026) :
//   - chaque __doPostBack = POST multipart complet → page entièrement rechargée ;
//     le formulaire est sérialisé au submit() : toute écriture postérieure dans
//     l'ancien document est perdue (cfg.delayMs simule un serveur lent) ;
//   - Type (changeType) → affiche le bloc consommations ;
//   - Fournisseur (relevé sur SIGEO le 25/11/2026) : input code SANS handler
//     de recherche (cfg.fouPostback = true pour l'ancien onchange → postback),
//     libellé <label id="body:x:proxyFou:x:lbl:x:lbl">, champs cachés
//     hidden_body:x:proxyFou:x:{txt|lbl}:x:{txt|lbl} toujours rendus VIDES ;
//     au postback, le libellé n'est réaffiché que s'il a été posté dans
//     hidden_…lbl (sinon vide, même avec un code valide) ; bouton
//     « Sélectionner » → popup('/popup.aspx/fr/sig_tec/fou_browse/ger?…') ;
//   - popup fou_browse : GET (Nom pré-rempli par le dernier filtre
//     « Positherm »), POST « Rechercher » (urlencoded) → grille
//     #body_x_gFournisseurs_x dont chaque lien fait, sur l'opener,
//     lbl.innerHTML = '<libellé>'; txt.value = '<code avec zéros>' ;
//     recherche par N° FU avec ou sans zéros de tête (cfg.fouNoZeroSearch :
//     avec zéro de tête → aucun résultat) ;
//   - fiche lot …/page.aspx/fr/sig_imm/lot_manage/ger/{imm}/{lot} : onglet
//     Diagnostics (#body_x_tabc_x_diagnostic_tab_x_diagnostic_x_grid_x) listant
//     les fiches enregistrées du lot (cfg.lotPageMissing : page sans grille) ;
//   - Établi le : onchange = ValidatorOnChange('body_x_txtDateReal_x_txtDateReal');
//     performanceEnergetique(this) → calcule « Valable jusqu'au » (input caché +
//     libellé) puis __doPostBack(date) ; comme sur la vraie page, un change
//     SYNTHÉTIQUE n'atteint pas performanceEnergetique (cfg.dateSynthetic =
//     false par défaut) — l'appeler directement dans la page, si ;
//   - N° ADEME (faits vérifiés le 23/11/2026) : en CRÉATION, l'input
//     body_x_textIdDpe_x_textIdDpe (avec span#body_x_titreIdDpe et les
//     estimations) n'est généré au rechargement que si Type = DPE ET « Établi
//     le » ≥ 01/07/2021 (quel que soit le champ qui a provoqué le rechargement) ;
//     une fois la fiche enregistrée, l'input est présent quelle que soit la date
//     tant qu'aucun N° n'est enregistré, puis le N° devient un texte non
//     modifiable dans span#body_x_labelIdDpe (cfg.ademeNever : jamais d'input
//     en création, pour simuler un champ manquant) ;
//   - cfg.refuseUpdate : tout Enregistrer sur une fiche existante est refusé ;
//   - Enregistrer : code fournisseur EXACT (avec ses zéros, comme la popup le
//     renvoie) → fournisseur retenu ; même code sans ses zéros (552043 pour
//     0552043) → fiche créée SANS fournisseur (bug d'origine) ; code inconnu →
//     refus « Le fournisseur X n'existe pas. » ; code stocké sans zéros de tête
//     (0448957 → 448957), libellé du fournisseur réaffiché ensuite
//     (cfg.fouLabelBlank : libellé vide après enregistrement) ;
//   - cfg.uploadHidden = N : un PDF uploadé est bien enregistré mais n'apparaît
//     dans la liste qu'au (N+1)e affichage de la fiche (comme SIGEO, qui ne le
//     montre parfois qu'après un nouvel « Enregistrer ») ;
//   - Date début validité : body_x_txtDateDebValid_x_txtDateDebValid (sans postback) ;
//   - cfg.clearOnChange = "dateDeb,comment" : la page vide ces champs au change
//     (valeur refusée côté client) ; cfg.corruptOnSave = "ademe,ges" : le serveur
//     enregistre une autre valeur (ADEME vidé, GES + 1) → écarts après enregistrement ;
//   - les cases et les coûts annuels n'apparaissent qu'avec une date ≥ 07/2021
//     (revérifiés après chaque postback) ;
//   - conso / GES → postback, lettres calculées, valeurs réaffichées « 320,00 » ;
//   - Fournisseur et Description : onkeydown = press(cmdSave) → Entrée = Enregistrer
//     (compté dans stats.enterSaves : doit rester à 0) ;
//   - Enregistrer : infomsg « Enregistrement réussi ! » ou li d'erreur ;
//     création d'un id (form action → …/id, ou redirection si cfg.redirect) ;
//   - après enregistrement : N° ADEME en span#body_x_labelIdDpe, « Valable
//     jusqu'au » en libellé, bloc « Joindre un document » (lien bascule) ;
//   - upload : onchange="__doPostBack('')", grille PowerGrid « N Résultat(s) ».
//   const srv = await require("./sigeo-diag-server.js").start();
const http = require("http");

// Code fournisseur stocké sans zéros de tête par SIGEO.
const fouKey = (c) => String(c || "").trim().replace(/^0+(?=\d)/, "");
// Fournisseurs (code exact tel que renvoyé par fou_browse → libellé, statut).
const SUPPLIER_LIST = [
  { code: "0448957", label: "ACADE Diagnostic", status: "Actif" },
  { code: "0528825", label: "EXPERT IMMOBILIER", status: "Actif" },
  { code: "0410962", label: "DIAG A FAIRE", status: "Actif" },
  { code: "0590971", label: "DIAG IMMO 590971", status: "Actif" },
  { code: "0552043", label: "ANAPLOMB", status: "Interdit" },
  { code: "0612345", label: "L'EXPERT & FILS", status: "Actif" }
];
const SUPPLIERS = Object.fromEntries(SUPPLIER_LIST.map((x) => [fouKey(x.code), x.label]));
const supplierExact = (code) => SUPPLIER_LIST.find((x) => x.code === String(code || "").trim()) || null;
const supplierByKey = (code) => SUPPLIER_LIST.find((x) => fouKey(x.code) === fouKey(code)) || null;
const FOU_BROWSE = "/popup.aspx/fr/sig_tec/fou_browse/ger";
const jsq = (s) => String(s).replace(/\\/g, "\\\\").replace(/'/g, "\\'");

function esc(s) { return String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;"); }
function parseDate(s) {
  const m = String(s || "").trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  return m ? { d: +m[1], m: +m[2], y: +m[3], ts: Date.UTC(+m[3], +m[2] - 1, +m[1]) } : null;
}
const isNew2021 = (s) => { const d = parseDate(s); return !!d && d.ts >= Date.UTC(2021, 6, 1); };
function num(s) {
  const t = String(s ?? "").replace(/\s/g, "").replace(",", ".");
  if (!/^\d+(\.\d+)?$/.test(t)) return null;
  return Number(t);
}
const fmt2 = (n) => n === null ? "" : n.toFixed(2).replace(".", ",");
function letter(v, t) { if (v === null) return ""; const i = t.findIndex((x) => v < x); return "ABCDEFG"[i < 0 ? 6 : i]; }
const EP = [70, 110, 180, 250, 330, 420], GES = [6, 11, 30, 50, 70, 100];

// multipart/form-data → { fields, files }
function parseMultipart(buf, boundary) {
  const fields = {}, files = {};
  const sep = Buffer.from("--" + boundary);
  let pos = buf.indexOf(sep);
  while (pos >= 0) {
    const start = pos + sep.length;
    if (buf.slice(start, start + 2).toString() === "--") break;
    const next = buf.indexOf(sep, start);
    if (next < 0) break;
    const part = buf.slice(start + 2, next - 2); // \r\n … \r\n
    const hEnd = part.indexOf("\r\n\r\n");
    const head = part.slice(0, hEnd).toString("utf8");
    const data = part.slice(hEnd + 4);
    const name = (head.match(/name="([^"]*)"/) || [])[1];
    const filename = (head.match(/filename="([^"]*)"/) || [])[1];
    if (name !== undefined) {
      if (filename !== undefined) files[name] = { filename, data };
      else fields[name] = data.toString("utf8");
    }
    pos = next;
  }
  return { fields, files };
}

const N = {
  type: "body$x$ddlDiagnostic$ddlDiagnostic", fou: "body$x$proxyFou$x$txt$x$txt",
  dateReal: "body$x$txtDateReal$x$txtDateReal", dateFin: "body$x$txtDateFinValid$x$txtDateFinValid",
  ademe: "body$x$textIdDpe$x$textIdDpe", dateDeb: "body$x$txtDateDebValid$x$txtDateDebValid", conso: "body$x$txtConsoE$x$txtConsoE", consoL: "body$x$txtConsoLegende$x$_ctl0",
  ges: "body$x$txtEmission$x$txtEmission", gesL: "body$x$txtEmisLegende$x$_ctl0",
  forceEp: "body$x$txtForceDiag$x$_ctl0", forceGes: "body$x$txtForceGes$x$_ctl0",
  altitude: "body$x$chkAltitude$x$_ctl0", nouvelles: "body$x$chkSurface$x$_ctl0",
  coutMin: "body$x$txtEstimationBasse$x$txtEstimationBasse", coutMax: "body$x$txtEstimationHaute$x$txtEstimationHaute",
  datePrix: "body$x$txtDateReferencePrix$x$txtDateReferencePrix",
  resultat: "body$x$txtResultat$x$txtResultat", comment: "body$x$txtComment$x$txtComment",
  save: "proxyActionBar$x$_cmdSave$x$_btn", end: "proxyActionBar$x$_cmdEnd$x$_btn",
  fouHidTxt: "hidden_body:x:proxyFou:x:txt:x:txt", fouHidLbl: "hidden_body:x:proxyFou:x:lbl:x:lbl",
  desc: "body$x$prxFileManager$x$txtDocDescription$x$_ctl0", file: "body$x$prxFileManager$x$fileUpload"
};
const idOf = (name) => name.replace(/\$/g, "_");

function blank() {
  return { type: "", fou: "", dateReal: "", dateFin: "", dateDeb: "", ademe: "", conso: "", ges: "", forceEp: "", forceGes: "",
    altitude: false, nouvelles: false, coutMin: "", coutMax: "", datePrix: "", resultat: "", comment: "", docs: [] };
}

function render(url, d, opts) {
  const cfg = opts.cfg || {};
  const saved = !!opts.saved;
  const cons = num(d.conso), ges = num(d.ges);
  const showConso = !!d.type;
  const dateDone = !!parseDate(d.dateReal) && opts.dateKnown;
  const ademeLabel = saved && !!d.ademe;
  const showAdeme = showConso && !ademeLabel && (saved || (isNew2021(d.dateReal) && !cfg.ademeNever));
  const showCases = dateDone && isNew2021(d.dateReal);
  const showCouts = showConso && isNew2021(d.dateReal);
  const msgs = opts.msgs || [];
  const clear = String(cfg.clearOnChange || "").split(",").map((x) => x.trim()).filter(Boolean);
  const clr = (k) => clear.includes(k) ? ` onchange="this.value=''"` : "";
  const inp = (k, extra = "") => `<input type="text" id="${idOf(N[k])}" name="${N[k]}" value="${esc(d[k])}" ${extra}${clr(k)}>`;
  const chk = (k) => `<input type="checkbox" id="${idOf(N[k])}" name="${N[k]}" ${d[k] ? "checked" : ""}>`;
  // Documents enregistrés mais pas encore affichés (cfg.uploadHidden).
  const docsAll = d.docs || [];
  const docs = docsAll.filter((x) => !(x.hiddenRenders > 0));
  docsAll.forEach((x) => { if (x.hiddenRenders > 0) x.hiddenRenders--; });
  const grid = docs.length
    ? docs.map((x) => `<tr class="PowerGridRowClass"><td>${esc(x.desc)}</td><td><a href="#">${esc(x.name)}</a></td><td>23/11/2026</td></tr>`).join("") +
      `<tr class="PowerGridPagerClass"><td colspan="3">${docs.length} Résultat(s)</td></tr>`
    : `<tr class="PowerGridEmptyPagerClass"><td colspan="3">0 Résultat(s)</td></tr>`;
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Diagnostic</title></head><body>
<form id="mainForm" method="post" enctype="multipart/form-data" action="${esc(opts.action || url)}">
<input type="hidden" name="__EVENTTARGET" id="__EVENTTARGET" value="">
<input type="hidden" name="__EVENTARGUMENT" id="__EVENTARGUMENT" value="">
<input type="hidden" name="__VIEWSTATE" value="dDwtMTIz">
<input type="hidden" name="__nhEnter" id="__nhEnter" value="">
<input type="hidden" name="__dateKnown" value="${dateDone ? 1 : ""}">
<table id="msgline" style="${msgs.length ? "" : "display:none"}"><tr><td><div id="msgbody"><ul>${msgs.map((m) => `<li class="${m.cls}">${esc(m.text)}</li>`).join("")}</ul></div></td></tr></table>
<div class="actionbar">
  <input type="button" id="proxyActionBar_x__cmdSave_x__btn" name="${N.save}" value="Enregistrer" onclick="__doPostBack('${N.save}','')">
  <input type="button" id="proxyActionBar_x__cmdEnd_x__btn" name="${N.end}" value="Enreg. et fermer" onclick="__doPostBack('${N.end}','')">
  <input type="button" id="proxyActionBar_x_BackAction_x__btn" value="Fermer" onclick="window.close()">
</div>
<table class="form">
  <tr><td>Type</td><td><select id="${idOf(N.type)}" name="${N.type}" onchange="changeType()">
    <option value=""></option><option value="3"${d.type === "3" ? " selected" : ""}>Amiante</option>
    <option value="11"${d.type === "11" ? " selected" : ""}>Performance Energetique</option>
    <option value="15"${d.type === "15" ? " selected" : ""}>Performance Energetique</option></select></td></tr>
  <tr><td>Fournisseur</td><td><span id="body_x_proxyFou"><input type="text" id="${idOf(N.fou)}" name="${N.fou}" maxlength="9" value="${esc(d.fou)}" onkeydown="return press(cmdSave)"${cfg.fouPostback ? ` onchange="__doPostBack('${N.fou}','')"` : ""}>
    <input type="submit" id="body_x_proxyFou_x_btn" name="body:x:proxyFou:x:btn" value="Sélectionner" onclick="document.getElementById('${idOf(N.fou)}').focus();popup('${FOU_BROWSE}?opener=diag_manage&amp;txt=${idOf(N.fou)}&amp;lbl=body:x:proxyFou:x:lbl:x:lbl');return false;;return abort(this);"><label id="body:x:proxyFou:x:lbl:x:lbl">${esc(opts.fouLabel || "")}</label>
    <input type="hidden" id="${N.fouHidLbl}" name="${N.fouHidLbl}" value=""><input type="hidden" id="${N.fouHidTxt}" name="${N.fouHidTxt}" value=""></span></td></tr>
</table>
<div id="div_conso" style="${showConso ? "" : "display:none"}">
<table class="dpe">
  <tr><td>Établi le</td><td>Valable jusqu'au</td><td>${showAdeme || ademeLabel ? `<span id="body_x_titreIdDpe">N° ADEME</span>` : ""}</td></tr>
  <tr><td>${inp("dateReal", `maxlength="10" onchange="ValidatorOnChange('body_x_txtDateReal_x_txtDateReal'); performanceEnergetique(this);"`)}</td>
      <td><label id="fake_txtDateFinValid_x_txtDateFinValid">${esc(d.dateFin)}</label><input type="text" style="display:none" id="${idOf(N.dateFin)}" name="${N.dateFin}" value="${esc(d.dateFin)}"></td>
      <td>${ademeLabel ? `<span id="body_x_labelIdDpe">${esc(d.ademe)}</span>` : showAdeme ? inp("ademe", 'maxlength="13"') : ""}</td></tr>
  <tr><td>Date début validité</td><td>${inp("dateDeb", 'maxlength="10"')}</td><td></td></tr>
</table>
<div id="body_x_blocCases" style="${showCases ? "" : "display:none"}">
  ${chk("altitude")} <label for="${idOf(N.altitude)}">Logement à plus de 800m d'altitude en H1B, H1C ou H2D</label>
  ${chk("nouvelles")} <label for="${idOf(N.nouvelles)}">Nouvelles valeurs du DPE</label>
</div>
<table>
  <tr><td>Conso énergie primaire</td><td><input type="text" id="${idOf(N.conso)}" name="${N.conso}" maxlength="6" value="${esc(cons === null ? d.conso : fmt2(cons))}" onchange="__doPostBack('${N.conso}','')"></td>
      <td><input type="text" readonly id="${idOf(N.consoL)}" name="${N.consoL}" value="${letter(cons, EP)}"></td></tr>
  <tr><td>Émission GES</td><td><input type="text" id="${idOf(N.ges)}" name="${N.ges}" maxlength="6" value="${esc(ges === null ? d.ges : fmt2(ges))}" onchange="__doPostBack('${N.ges}','')"></td>
      <td><input type="text" readonly id="${idOf(N.gesL)}" name="${N.gesL}" value="${letter(ges, GES)}"></td></tr>
  <tr><td>Forcer lettre EP</td><td>${inp("forceEp", 'maxlength="1"')}</td><td>Forcer lettre GES ${inp("forceGes", 'maxlength="1"')}</td></tr>
</table>
<div id="div_couts_annuels" style="${showCouts ? "" : "display:none"}">
  Estimation ${inp("coutMin", 'maxlength="5"')} € à ${inp("coutMax", 'maxlength="5"')} € — prix indexés au ${inp("datePrix", 'maxlength="10"')}
</div>
</div>
<table>
  <tr><td>Résultat</td><td><textarea id="${idOf(N.resultat)}" name="${N.resultat}" maxlength="1000">${esc(d.resultat)}</textarea></td></tr>
  <tr><td>Commentaires</td><td><textarea id="${idOf(N.comment)}" name="${N.comment}"${clr("comment")}>${esc(d.comment)}</textarea></td></tr>
</table>
${saved ? `<fieldset><legend>Documents</legend>
  <a id="body_x_prxFileManager_x_linkAjoutDocument" href="javascript:showAddDocument()">Joindre un document</a>
  <table>
    <tr id="body_x_prxFileManager_x_trDesc" style="display:none"><td>Description</td><td><input type="text" id="${idOf(N.desc)}" name="${N.desc}" onkeydown="return press(cmdSave)"></td></tr>
    <tr id="body_x_prxFileManager_x_trFichier" style="display:none"><td>Fichier</td><td><input type="file" id="${idOf(N.file)}" name="${N.file}" onchange="__doPostBack('')"></td></tr>
  </table>
  <table id="body_x_grid_x" class="PowerGridClass"><tr class="PowerGridHeaderClass"><td>Description</td><td>Fichier</td><td>Date</td></tr>${grid}</table>
</fieldset>` : `<p>Veuillez enregistrer le diagnostic avant d'y ajouter un document.</p>`}
</form>
<script>
var cmdSave = "proxyActionBar_x__cmdSave_x__btn";
function __doPostBack(t, a) {
  var f = document.getElementById("mainForm");
  document.getElementById("__EVENTTARGET").value = t;
  document.getElementById("__EVENTARGUMENT").value = a || "";
  f.submit();
}
function press(btn) {
  var e = window.event;
  if (e && (e.keyCode === 13 || e.key === "Enter")) { document.getElementById("__nhEnter").value = "1"; document.getElementById(btn).click(); return false; }
  return true;
}
function changeType() { __doPostBack("${N.type}", ""); }
function ValidatorOnChange(x) {
  var t = typeof x === "string" ? document.getElementById(x) : ((x || window.event).srcElement || (x || window.event).target);
  return !!t;
}
var NH_DATE_SYNTHETIC = ${cfg.dateSynthetic ? "true" : "false"};
function performanceEnergetique(el) {
  // Vraie page : un change synthétique (dispatchEvent) n'atteint pas le postback.
  var e = window.event;
  if (e && e.isTrusted === false && !NH_DATE_SYNTHETIC) return;
  var m = /^(\\d{2})\\/(\\d{2})\\/(\\d{4})$/.exec(el.value);
  var fin = document.getElementById("${idOf(N.dateFin)}");
  if (m && fin) {
    var d = +m[1], mo = +m[2], y = +m[3], t = Date.UTC(y, mo - 1, d), r;
    var p10 = function () { var x = new Date(Date.UTC(y + 10, mo - 1, d) - 86400000); return ("0" + x.getUTCDate()).slice(-2) + "/" + ("0" + (x.getUTCMonth() + 1)).slice(-2) + "/" + x.getUTCFullYear(); };
    if (y < 2013) r = p10(); else if (y <= 2017) r = "31/12/2022"; else if (t < Date.UTC(2021, 6, 1)) r = "31/12/2024"; else r = p10();
    fin.value = r;
    var lab = document.getElementById("fake_txtDateFinValid_x_txtDateFinValid");
    if (lab) lab.textContent = r;
  }
  __doPostBack("${N.dateReal}", "");
}
function showAddDocument() {
  ["body_x_prxFileManager_x_trDesc", "body_x_prxFileManager_x_trFichier"].forEach(function (id) {
    var e = document.getElementById(id); e.style.display = e.style.display === "none" ? "" : "none";
  });
}
</script>
</body></html>`;
}

// Popup fou_browse : GET = formulaire (Nom pré-rempli par le dernier filtre de
// l'utilisateur), POST « Rechercher » = grille des fournisseurs.
function fouBrowse(req, res, u, stats, config = {}) {
  const txt = u.searchParams.get("txt") || "", lbl = u.searchParams.get("lbl") || "";
  const page = (crit, rows) => `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Fournisseurs</title></head><body>
<form id="mainForm" method="post" action="${esc(u.pathname + u.search)}">
<input type="hidden" name="__VIEWSTATE" value="fou">
<input type="hidden" id="REQUEST_METHOD" name="REQUEST_METHOD" value="GET">
Nom <input type="text" id="body_x_tbNom_x_nTbNom" name="body:x:tbNom:x:nTbNom" value="${esc(crit.nom)}">
<select id="body_x_maRecherche_x_ddlSearch_ddlSearch" name="body:x:maRecherche:x:ddlSearch:ddlSearch"><option value="1">Commence par</option><option value="2" selected>Contient</option></select>
SIREN <input type="text" id="body_x_txtSiren_x_txtSiren" name="body:x:txtSiren:x:txtSiren" value="">
N° FU <input type="text" id="body_x_tbNum_x_nTbNum" name="body:x:tbNum:x:nTbNum" value="${esc(crit.num)}">
<input type="submit" id="body_x_searchRaz_x_cmdSearchBtn_x__btn" name="body:x:searchRaz:x:cmdSearchBtn:x:_btn" value="Rechercher">
<input type="submit" id="body_x_searchRaz_x_cmdRazBtn_x__btn" name="body:x:searchRaz:x:cmdRazBtn:x:_btn" value="RAZ">
${rows === null ? "<p>Veuillez renseigner des critères et lancez votre recherche</p>" : `<table id="body_x_gFournisseurs_x" class="PowerGridClass">
<tr class="PowerGridHeaderClass"><td></td><td>Nom</td><td>SIREN</td><td>Statut</td><td>Mode de paiement</td></tr>
${rows.map((x) => `<tr class="PowerGridItemClass"><td><a class="edit_link" href="javascript:void(0)" onclick="${esc(`window.opener.document.getElementById('${jsq(lbl)}').innerHTML = '${jsq(esc(x.label))}';window.opener.document.getElementById('${jsq(txt)}').value = '${jsq(x.code)}';window.close();return false;`)}"><img src="/image/ico_selector.gif">&nbsp;FU${esc(x.code)}</a></td><td>${esc(x.label)}</td><td>408194629</td><td>${esc(x.status)}</td><td>Chèque</td></tr>`).join("\n")}
<tr><td colspan="5"><b>${rows.length} Résultat(s)</b></td></tr></table>`}
</form></body></html>`;
  res.setHeader("Content-Type", "text/html; charset=utf-8");
  if (req.method === "GET") { res.end(page({ nom: "Positherm", num: "" }, null)); return; }
  const chunks = [];
  req.on("data", (c) => chunks.push(c));
  req.on("end", () => {
    stats.fouSearches++;
    const f = new URLSearchParams(Buffer.concat(chunks).toString("utf8"));
    const nom = (f.get("body:x:tbNom:x:nTbNom") || "").trim().toLowerCase();
    const num = (f.get("body:x:tbNum:x:nTbNum") || "").trim();
    let rows = SUPPLIER_LIST.slice();
    if (nom) rows = rows.filter((x) => x.label.toLowerCase().includes(nom));
    // cfg.fouNoZeroSearch : une recherche avec zéro de tête ne trouve rien
    // (pour vérifier le repli sur la variante sans zéro).
    if (num) rows = config.fouNoZeroSearch && /^0/.test(num) ? [] : rows.filter((x) => x.code.includes(num) || fouKey(x.code) === fouKey(num));
    res.end(page({ nom: f.get("body:x:tbNom:x:nTbNom") || "", num }, rows));
  });
}

// Fiche lot, onglet Diagnostics : fiches enregistrées du lot.
function lotPage(imm, lot, rows) {
  const lab = (t) => (t === "11" || t === "15") ? "Performance Energetique" : t === "3" ? "Amiante" : "Diagnostic " + t;
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Lot N°LT${lot}</title></head><body>
<div id="diagnostic_tab" style="display:none"><table class="PowerGridClass" id="body_x_tabc_x_diagnostic_tab_x_diagnostic_x_grid_x">
${rows.length ? `<tr class="PowerGridHeaderClass"><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><a href="#">Diagnostic</a></td><td>Commentaires</td><td><a href="#">Date de début de validité</a></td><td><a href="#">Date de fin de validité</a></td><td>Résultat</td></tr>` : ""}
${rows.map((d) => `<tr class="PowerGridItemClass"><td><a href="javascript:void(0)" onclick="if (confirm('Confirmer la supression de ce diagnostic ?')){document.getElementById('x_diag_id').value = '${d.id}';}return false;"><img src="/image/ico_del.gif"></a></td><td><a class="edit_link" href="javascript:void(0);" onclick="popup('/popup.aspx/fr/sig_imm/diag_manage/${imm}/${lot}/${d.id}?upcid=body_x_tabc_x_diagnostic_tab_x_diagnostic_x_upcDiag');return false;"><img src="/image/ico_mod.gif"></a></td><td></td><td>${esc(lab(d.type))}</td><td>${esc(d.comment)}</td><td>${esc(d.dateReal)}</td><td>${esc(d.dateFin)}</td><td>${esc(d.resultat)}</td></tr>`).join("\n")}
<tr><td colspan="8"><b>${rows.length} Résultat(s)</b></td></tr></table></div></body></html>`;
}

async function start(cfg = {}) {
  const store = new Map(); // id → données enregistrées
  const flash = new Map();
  const stats = { saves: 0, creates: 0, updates: 0, uploads: 0, enterSaves: 0, saveCloses: 0, posts: 0, refused: 0,
    fouSearches: 0, lotPages: 0, hiddenPosted: 0, hiddenLabelPosted: 0, savedWithoutSupplier: 0 };
  const config = Object.assign({ redirect: false, delayMs: 0, fouPostback: false, lotPageMissing: false, dateSynthetic: false, ademeNever: false, refuseUpdate: false, uploadHidden: 0, fouLabelBlank: false, silentRefuse: false }, cfg);
  let nextId = 1147729;
  const savedLabel = (d) => (config.fouLabelBlank ? "" : (d && d.fouLabel) || "");

  const server = http.createServer((req, res) => {
    const u = new URL(req.url, "http://x");
    if (u.pathname === "/__config") {
      for (const [k, v] of u.searchParams) config[k] = v === "1" ? true : v === "0" ? false : /^\d+$/.test(v) ? Number(v) : v;
      res.end("ok");
      return;
    }
    if (u.pathname === FOU_BROWSE) { fouBrowse(req, res, u, stats, config); return; }
    const ml = u.pathname.match(/^\/page\.aspx\/fr\/sig_imm\/lot_manage\/(ger|syn)\/([^/]+)\/(\d+)$/);
    if (ml) {
      stats.lotPages++;
      res.setHeader("Content-Type", "text/html; charset=utf-8");
      if (config.lotPageMissing) { res.end("<!DOCTYPE html><html><head><title>Erreur</title></head><body>Erreur</body></html>"); return; }
      const rows = [...store.entries()].filter(([, d]) => d.lot === ml[2]).map(([id, d]) => Object.assign({ id }, d));
      res.end(lotPage(ml[2], ml[3], rows));
      return;
    }
    const m = u.pathname.match(/^\/popup\.aspx\/fr\/sig_imm\/diag_manage\/([^/]+)\/\d+\/(\d+)$/);
    if (!m) { res.statusCode = 404; res.end("not found"); return; }
    const lot = m[1], id = m[2];
    const url = `http://${req.headers.host}${u.pathname}`;
    const send = (html) => {
      setTimeout(() => { res.setHeader("Content-Type", "text/html; charset=utf-8"); res.end(html); }, config.delayMs || 0);
    };
    if (req.method === "GET") {
      if (id !== "0" && !store.has(id)) { res.statusCode = 404; res.end("diagnostic inconnu"); return; }
      const d = id === "0" ? blank() : JSON.parse(JSON.stringify(store.get(id)));
      if (id !== "0") d.docs = store.get(id).docs;
      const msgs = flash.get(id) || [];
      flash.delete(id);
      send(render(url, d, { saved: id !== "0", dateKnown: id !== "0", msgs, cfg: config, fouLabel: id !== "0" ? savedLabel(d) : "" }));
      return;
    }
    const chunks = [];
    req.on("data", (c) => chunks.push(c));
    req.on("end", () => {
      stats.posts++;
      const boundary = ((req.headers["content-type"] || "").match(/boundary=(.+)$/) || [])[1];
      const { fields, files } = parseMultipart(Buffer.concat(chunks), boundary);
      const target = fields.__EVENTTARGET || "";
      const saved = id !== "0" && store.has(id);
      const d = saved ? JSON.parse(JSON.stringify(store.get(id))) : blank();
      // Valeurs postées (le N° ADEME n'est modifiable qu'avant enregistrement)
      for (const k of ["type", "fou", "dateReal", "dateDeb", "conso", "ges", "forceEp", "forceGes", "coutMin", "coutMax", "datePrix", "resultat", "comment"]) {
        if (fields[N[k]] !== undefined) d[k] = fields[N[k]];
      }
      if (fields[N.dateFin] !== undefined) d.dateFin = fields[N.dateFin];
      // N° ADEME modifiable tant qu'il n'est pas enregistré (ensuite : texte seul).
      if (fields[N.ademe] !== undefined && !(saved && store.get(id).ademe)) d.ademe = fields[N.ademe];
      d.altitude = fields[N.altitude] !== undefined;
      d.nouvelles = fields[N.nouvelles] !== undefined;
      const dateKnown = !!fields.__dateKnown || target === N.dateReal;
      // Libellé du fournisseur réaffiché seulement s'il a été posté dans le champ caché.
      // Fiche enregistrée et même code : libellé enregistré (base) à défaut.
      const postedLabel = String(fields[N.fouHidLbl] || "") ||
        (saved && fields[N.fou] !== undefined && fouKey(fields[N.fou]) === fouKey(store.get(id).fou) ? savedLabel(store.get(id)) : "");
      if (fields[N.fouHidLbl] !== undefined || fields[N.fouHidTxt] !== undefined) stats.hiddenPosted++;
      if (postedLabel) stats.hiddenLabelPosted++;
      let msgs = [];
      let action = url;

      if (target === N.end) stats.saveCloses++;
      if (target === N.save) {
        stats.saves++;
        if (fields.__nhEnter) stats.enterSaves++;
        const errs = [];
        if (!d.type) errs.push("Le type de diagnostic est obligatoire.");
        if (!supplierByKey(d.fou)) errs.push(`Le fournisseur ${d.fou || "(vide)"} n'existe pas.`);
        if (!parseDate(d.dateReal)) errs.push("La date de réalisation est invalide.");
        if (num(d.conso) === null || num(d.ges) === null) errs.push("Consommation / émission invalides.");
        if (!saved && isNew2021(d.dateReal) && !/^(\d{4}[A-Z]\d{7}[A-Z]|NC)$/.test(d.ademe)) errs.push("Le N° ADEME est obligatoire pour un DPE établi depuis le 01/07/2021.");
        if (saved && config.refuseUpdate) errs.push("Mise à jour refusée (test).");
        // cfg.silentRefuse : création refusée sans message (validateur côté page) :
        // la page revient en création, rien n'est créé.
        if (!saved && config.silentRefuse) {
          stats.refused++;
          send(render(url, d, { saved: false, dateKnown, msgs: [], cfg: config, fouLabel: postedLabel }));
          return;
        }
        if (errs.length) {
          stats.refused++;
          msgs = errs.map((t) => ({ cls: "errormsg", text: t }));
          send(render(url, d, { saved, dateKnown, msgs, cfg: config, fouLabel: postedLabel }));
          return;
        }
        const n = (k) => { const x = num(d[k]); if (x !== null) d[k] = fmt2(x); };
        n("conso"); n("ges");
        // Code exact (avec zéros) → fournisseur retenu ; même code sans ses
        // zéros → fiche enregistrée SANS fournisseur (bug d'origine).
        const ex = supplierExact(d.fou);
        d.fouLabel = ex ? ex.label : "";
        if (!ex) stats.savedWithoutSupplier++;
        d.fou = fouKey(d.fou); // 0448957 → 448957
        const bad = String(config.corruptOnSave || "").split(",").map((x) => x.trim());
        if (bad.includes("ademe")) d.ademe = "";
        if (bad.includes("ges") && num(d.ges) !== null) d.ges = fmt2(num(d.ges) + 1);
        let newId = id;
        if (!saved) { newId = String(nextId++); stats.creates++; } else stats.updates++;
        store.set(newId, Object.assign({ lot }, d));
        msgs = [{ cls: "infomsg", text: "Enregistrement réussi !" }];
        const newUrl = url.replace(/\/\d+$/, "/" + newId);
        if (config.redirect && newId !== id) {
          flash.set(newId, msgs);
          setTimeout(() => { res.statusCode = 302; res.setHeader("Location", newUrl); res.end(); }, config.delayMs || 0);
          return;
        }
        send(render(url, store.get(newId), { saved: true, dateKnown: true, msgs, action: newUrl, cfg: config, fouLabel: savedLabel(store.get(newId)) }));
        return;
      }
      if (target === "" && files[N.file] && files[N.file].filename) {
        if (!saved) { msgs = [{ cls: "errormsg", text: "Veuillez enregistrer le diagnostic avant d'y ajouter un document." }]; }
        else if (!String(fields[N.desc] || "").trim()) { msgs = [{ cls: "errormsg", text: "La description du document est obligatoire." }]; }
        else {
          const f = files[N.file];
          store.get(id).docs.push({ desc: fields[N.desc].trim(), name: f.filename, size: f.data.length, pdf: f.data.slice(0, 4).toString() === "%PDF",
            hiddenRenders: Number(config.uploadHidden) || 0 });
          stats.uploads++;
        }
      }
      if (saved) d.docs = store.get(id).docs;
      send(render(url, d, { saved, dateKnown, msgs, action, cfg: config, fouLabel: postedLabel }));
    });
  });
  await new Promise((r) => server.listen(0, "127.0.0.1", r));
  const origin = `http://127.0.0.1:${server.address().port}`;
  return {
    server, origin, store, stats, config,
    url: (lot, id = 0, kind = 10) => `${origin}/popup.aspx/fr/sig_imm/diag_manage/${lot}/${kind}/${id}`,
    diags: () => [...store.entries()].map(([id, d]) => Object.assign({ id }, d))
  };
}

module.exports = { start, SUPPLIERS, SUPPLIER_LIST };
