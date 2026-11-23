// Faux SIGEO pour les tests DPE. Comme le vrai : chaque lien de fichier a
// href="javascript:__doPostBack(…)" (qui renvoie la page Documents, PAS le
// fichier) et onclick="popup('/bare.aspx/fr/fil/ged_file_download/<ID>', …)" ;
// ce GET redirige vers le fichier (PDF), ou vers une page HTML pour le lot H.
// Lot N : liens sans popup(…) → « lien de téléchargement introuvable ».
//
// Faux SIGEO pour les tests DPE : page Documents d'un lot (popup.aspx
// ged_multi_file_manager) reproduisant la structure observée — un seul
// <form> ASP.NET, sections titrées « X >> Y » dans un tableau de mise en
// page, chaque section avec un tableau Nom | Description | Date dont les
// noms sont des liens __doPostBack, en-têtes cliquables (tri).
//   pages(lot)  → HTML (utilisable sous jsdom)
//   start()     → serveur HTTP (Playwright) : GET page / fiche lot, POST →
//                 PDF (Content-Disposition) ou HTML selon la cible.
const http = require("http");

const VIEWSTATE = "dDwtMTIzNDU2Nzg5O3Q8O2w8aTwxPjs+O2w8dDw7bDxpPDE+Oz47bDx0PDtsPGk8MT47";

// Données des lots. Chaque fichier : [nom, description, date].
const LOTS = {
  // Cas réel : un seul DPE au milieu des entretiens chaudière / diag gaz-élec.
  A: {
    "Diagnostics >> EPC": [
      ["Entretien chaudie?re. 09.2023", "RapportDiagnosticLot", "23/10/2023"],
      ["Entretien chaudiere 06.2022.pdf", "RapportDiagnosticLot", "29/12/2022"],
      ["DIAG GAZ 2022", "RapportDiagnosticLot", "19/10/2022"],
      ["DPE 2022", "RapportDiagnosticLot", "19/10/2022"],
      ["DIAG ELEC 2022", "RapportDiagnosticLot", "19/10/2022"],
      ["ENT CHAUDIERE 10.2020.pdf", "RapportDiagnosticLot", "23/11/2020"]
    ],
    "Diagnostics Multys >> EPC": [],
    "Divers": [["DPE 2025 brouillon", "Divers", "01/01/2025"]], // hors EPC : ignoré
    "Photos": [["Salon.jpg", "", "02/02/2021"]]
  },
  // Deux DPE de dates différentes, dans les deux sections EPC.
  B: {
    "Diagnostics >> EPC": [
      ["DPE 2018.pdf", "RapportDiagnosticLot", "01/02/2018"],
      ["Diagnostic gaz", "RapportDiagnosticLot", "01/02/2018"]
    ],
    "Diagnostics Multys >> EPC": [
      ["Diagnostic de performance énergétique", "Multys", "15/03/2023"]
    ]
  },
  // Aucun DPE (dont un « performance énergétique » de chaudière : exclu).
  C: {
    "Diagnostics >> EPC": [
      ["DIAG ELEC 2021", "RapportDiagnosticLot", "05/05/2021"],
      ["Rapport performance énergétique chaudière", "RapportDiagnosticLot", "05/05/2021"],
      ["Amiante DTA", "RapportDiagnosticLot", "05/05/2021"]
    ],
    "Diagnostics Multys >> EPC": []
  },
  // Dates égales : l'année du nom départage, puis l'ordre du tableau.
  D: {
    "Diagnostics >> EPC": [
      ["DPE 2021", "RapportDiagnosticLot", "10/05/2023"],
      ["DPE 2023 (refait)", "RapportDiagnosticLot", "10/05/2023"],
      ["DPE ancien", "RapportDiagnosticLot", "01/01/2015"]
    ],
    "Diagnostics Multys >> EPC": [
      ["DPE", "Multys", "10/05/2023"]
    ]
  },
  // Structure réelle observée (capture du 23/11) : sections en <fieldset>
  // + <legend>, formulaire avec onclick, « Joindre un document ».
  F: {
    "Diagnostics >> EPC": [
      ["17312 (DIAG DPE2021).pdf", "RapportDiagnosticLot", "09/01/2025"],
      ["diag_elec_2017.pdf", "RapportDiagnosticLot", "17/08/2017"],
      ["additif_edle_1_emetteur_en_plus.msg", "RapportDiagnosticLot", "03/07/2017"],
      ["EDLE_WROBEL_ORALIA.pdf", "RapportDiagnosticLot", "31/05/2017"],
      ["Votre_entree_dans_le_logement_Proprietaire_BUNLON.msg", "RapportDiagnosticLot", "26/04/2017"],
      ["BUNLON_392_05.10.10_DP.pdf", "RapportDiagnosticLot", "26/04/2017"]
    ],
    "Diagnostics Multys >> EPC": [],
    "Divers": [],
    "Etats des lieux": [],
    "Photos": [],
    "Plans": [["plan.jpg", "PlanLot", "26/04/2017"]]
  },
  // Liens sans popup(…) : pas d'URL de téléchargement.
  N: { "Diagnostics >> EPC": [["DPE 2019", "RapportDiagnosticLot", "04/04/2019"]] },
  // Le DPE renvoie du HTML au lieu d'un PDF (ViewState rejeté côté serveur).
  H: {
    "Diagnostics >> EPC": [["DPE 2020", "RapportDiagnosticLot", "03/03/2020"]]
  },
  // Lot S : pas de DPE en gérance, mais un en syndic (metier=syn).
  S: { "Diagnostics >> EPC": [["DIAG GAZ 2021", "RapportDiagnosticLot", "05/05/2021"]] },
  S_syn: { "Diagnostics >> EPC": [["DPE 2023", "RapportDiagnosticLot", "10/02/2023"], ["DIAG ELEC 2023", "RapportDiagnosticLot", "10/02/2023"]] },
  // Lot T : rien ni en gérance ni en syndic.
  T: { "Diagnostics >> EPC": [] },
  // Page Documents d'un lot sans données : sections vides.
  EMPTY: { "Diagnostics >> EPC": [], "Divers >> Autres": [] }
};

const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
const fileId = (lot, si, fi) => `${lot}${si}${fi}`;
const onclickOf = (lot, si, fi, name) => lot === "N" ? "" :
  ` onclick="popup('/bare.aspx/fr/fil/ged_file_download/${fileId(lot, si, fi)}', '${esc(name).replace(/'/g, "")}', 100, 100, 800, 600); return false;"`;
const target = (lot, si, fi) => `body$x$reManagers$_ctl${si}$prxFileManager$x$grid$x$_ctl${fi + 3}$_ctl3` + (lot === "H" ? "$html" : "");

function sectionsHtml(lot) {
  const secs = LOTS[lot] || {};
  return Object.entries(secs).map(([title, files], si) => {
    const grid = files.length ? `
        <table class="grid" id="body_x_reManagers__ctl${si}_prxFileManager_x_grid">
          <tr class="header">
            <td><a href="javascript:__doPostBack('body$x$reManagers$_ctl${si}$prxFileManager$x$grid','Sort$Nom')">Nom</a></td>
            <td>Description</td><td>Date</td>
          </tr>
          ${files.map(([n, d, dt], fi) => `<tr class="row">
            <td><img src="pdf.gif"> <a href="javascript:__doPostBack('${target(lot, si, fi)}','')"${onclickOf(lot, si, fi, n)}>${esc(n)}</a></td>
            <td>${esc(d)}</td><td>${esc(dt)}</td>
            <td><a href="javascript:__doPostBack('${target(lot, si, fi).replace("_ctl3", "_ctl9")}','')"><img src="del.gif"></a></td>
          </tr>`).join("")}
        </table>` : `<span class="empty">Aucun document</span>`;
    return `
    <tr><td class="title"><span>${esc(title)}</span> <input type="button" value="Ajouter"></td></tr>
    <tr><td>${grid}</td></tr>`;
  }).join("");
}

// Variante fieldset / legend (lot F) : structure de la vraie page.
function fieldsetHtml(lot) {
  return Object.entries(LOTS[lot] || {}).map(([title, files], si) => `
  <fieldset class="ged"><legend>${esc(title)}</legend>
    ${title === "Divers" || title === "Photos" || title === "Plans" ? `<div><img src="add.gif"><a href="javascript:__doPostBack('body$x$reManagers$_ctl${si}$lnkAdd','')">Joindre un document</a></div>` : ""}
    ${files.length ? `<table cellspacing="0" rules="all" border="1" style="border-collapse:collapse;">
      <tr style="color:White;background-color:#333;font-weight:bold;"><td>Nom</td><td>Description</td><td>Date</td></tr>
      ${files.map(([n, d, dt], fi) => `<tr><td><a href="javascript:__doPostBack('${target(lot, si, fi)}','')"${onclickOf(lot, si, fi, n)}>${esc(n)}</a></td><td>${esc(d)}</td><td>${esc(dt)}</td></tr>`).join("")}
    </table>` : ""}
  </fieldset>`).join("");
}

// loaded=false : page telle qu'à l'ouverture (« Chargement en cours... »),
// sections injectées par un script après `delay` ms (si delay fourni).
function docsPage(lot, { loaded = true, delay = 0 } = {}) {
  const inner = lot === "F" ? fieldsetHtml(lot) : `<table class="layout">${sectionsHtml(lot)}</table>`;
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Gestion des documents</title></head><body>
<form name="aspnetForm" method="post" action="./lot_manage?lot=${lot}" id="aspnetForm"${lot === "F" ? ' onclick="if(window.hideMenus)hideMenus();"' : ""}>
<input type="hidden" name="__EVENTTARGET" id="__EVENTTARGET" value="">
<input type="hidden" name="__EVENTARGUMENT" id="__EVENTARGUMENT" value="">
<input type="hidden" name="__VIEWSTATE" id="__VIEWSTATE" value="${loaded ? VIEWSTATE : "initial"}">
<input type="hidden" name="__EVENTVALIDATION" id="__EVENTVALIDATION" value="ev-${lot}">
<script>function popup(u){window.open(u);}function __doPostBack(t,a){var f=document.forms[0];f.__EVENTTARGET.value=t;f.__EVENTARGUMENT.value=a;f.submit();}</script>
<div id="loading" style="${loaded ? "display:none" : ""}">Chargement en cours...</div>
<div id="content">${loaded ? inner : ""}</div>
${!loaded && delay ? `<script>setTimeout(function(){
  document.getElementById("content").innerHTML = ${JSON.stringify(inner)};
  document.getElementById("__VIEWSTATE").value = ${JSON.stringify(VIEWSTATE)};
  document.getElementById("loading").style.display = "none";
}, ${delay});</script>` : ""}
</form></body></html>`;
}

// Fiche lot : la page Documents s'ouvre par un window.open. En syndic
// (metier « syn »), la fiche contient aussi un tableau sans rapport (liste
// des copropriétaires) : la page Documents doit quand même être suivie.
function fichePage(lot, origin, metier = "") {
  const extra = metier === "syn" ? `<table><tr><td><a href="javascript:__doPostBack('body$x$grdCopro','Select$0')">DUPONT Jean</a></td><td>12/01/2019</td></tr></table>` : "";
  return `<!DOCTYPE html><html><body><form><h1>Lot ${lot}${metier ? " (" + metier + ")" : ""}</h1>${extra}
<a href="#" onclick="window.open('${origin}/popup.aspx/fr/fil/ged_multi_file_manager/sig_imm/lot_manage?${metier ? "metier=" + metier + "&amp;" : ""}lot=${lot}&amp;x=1','docs');return false;">Documents</a>
</form></body></html>`;
}

const PDF = Buffer.from("%PDF-1.4\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF\n");

function start() {
  const stats = { posts: [], gets: [] };
  const server = http.createServer((req, res) => {
    const u = new URL(req.url, "http://x");
    const lot = u.searchParams.get("lot") || "A";
    if (req.method === "POST") {
      // Comme SIGEO : le postback du lien renvoie la page Documents.
      stats.posts.push({ lot });
      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      return res.end(docsPage(lot));
    }
    let m = u.pathname.match(/\/bare\.aspx\/fr\/fil\/ged_file_download\/(\w+)$/);
    if (m) {
      stats.gets.push(m[1]);
      if (!/sigeo=ok/.test(req.headers.cookie || "")) { res.writeHead(302, { Location: "/login.aspx" }); return res.end(); }
      res.writeHead(302, { Location: m[1].startsWith("H") ? "/files/erreur.aspx" : `/files/${m[1]}.pdf` });
      return res.end();
    }
    m = u.pathname.match(/^\/files\/(\w+)\.pdf$/);
    if (m) {
      res.writeHead(200, { "Content-Type": "application/pdf", "Content-Disposition": `attachment; filename="RapportDiagnostic_${m[1]}.pdf"` });
      return res.end(PDF);
    }
    if (u.pathname === "/__login") {
      res.writeHead(200, { "Set-Cookie": "sigeo=ok; Path=/", "Content-Type": "text/html" });
      return res.end("<html><body>connecté</body></html>");
    }
    if (u.pathname === "/login.aspx") {
      res.writeHead(200, { "Content-Type": "text/html" });
      return res.end('<html><body><form><input type="password" name="pw"></form></body></html>');
    }
    res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
    if (u.pathname.includes("ged_multi_file_manager")) {
      const key = u.searchParams.get("metier") === "syn" ? lot + "_syn" : lot;
      return res.end(docsPage(LOTS[key] ? key : "EMPTY", { loaded: false, delay: 400 }));
    }
    if (u.pathname.includes("lot_manage/syn")) {
      // …/syn/{immeuble}/{lot} ; ?lot=… (anciens tests) = fiche gérance.
      const pm = u.pathname.match(/lot_manage\/syn\/\d+\/(\w+)$/);
      const syn = !u.searchParams.has("lot");
      return res.end(fichePage(syn && pm ? pm[1] : lot, `http://127.0.0.1:${server.address().port}`, syn ? "syn" : ""));
    }
    res.end("<html><body>ok</body></html>");
  });
  return new Promise((resolve) => server.listen(0, "127.0.0.1", () => {
    const origin = `http://127.0.0.1:${server.address().port}`;
    resolve({
      origin, stats, close: () => new Promise((r) => server.close(r)),
      docsUrl: (lot) => `${origin}/popup.aspx/fr/fil/ged_multi_file_manager/sig_imm/lot_manage?lot=${lot}`,
      ficheUrl: (lot) => `${origin}/page.aspx/fr/sig_imm/lot_manage/syn/393664/1?lot=${lot}`
    });
  }));
}

module.exports = { LOTS, VIEWSTATE, PDF, docsPage, fichePage, start };
