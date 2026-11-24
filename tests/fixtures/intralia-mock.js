// Faux Intralia + faux FlowerDocs pour les tests « DPE Intralia ».
// Données ANONYMISÉES (noms d'immeubles, adresses, noms de fichiers et UUID
// fictifs) reprenant la forme des cas réels : ids d'immeuble, n° de lot,
// dates de dépôt, doublons, documents non-DPE, lots sans DPE.
//
// Intralia (origine 1) — pages en ISO-8859-1 / windows-1252, sans charset
// dans l'en-tête (comme le vrai serveur) :
//   GET /gerance/diag.php?action=R&ID_SOCIETE=…   index (#example3, DataTables)
//   GET /gerance/diag.php                         formulaire (liste ID_SOCIETE)
//   GET /gerance/diaglot.php?id_societe&id_immeuble&id_lot   page du lot
//   déconnecté → page de connexion (champ mot de passe) ; tout POST est
//   compté (les tests vérifient qu'il n'y en a aucun).
// FlowerDocs (origine 2) :
//   GET /flower-docs-gui/download?documentId=…&index=0 → PDF (octet-stream)
//   déconnecté → 302 vers une page de connexion HTML.
const http = require("http");

const CAB = { 19: "SICOV", 31: "CABINET NORD", 42: "CABINET SUD" };
const U = (n) => `ef1e3745-0000-4000-8000-${String(n).padStart(12, "0")}`;

// [soc, imm, nom, adresse, cp, ville, [[lot, type]…]]
const IMMS = [
  [19, 5004, "PARC FLEURI LE CONCERTO 1", "12 RUE DES MUSICIENS", "59000", "LILLE", [["732", "Appartement"], ["733", "Garage"]]],
  [19, 5005, "PARC FLEURI LE CONCERTO 2", "14 RUE DES MUSICIENS", "59000", "LILLE", [["732", "Appartement"], ["740", "Cave"]]],
  [19, 5305, "COUR SAINT-HUBERT", "5 COUR SAINT-HUBERT", "59000", "LILLE", [["130", "Appartement"]]],
  [19, 49, "RÉSIDENCE BEAUSÉJOUR", "3 PLACE DU MARCHÉ", "59800", "LILLE", [["543", "Appartement"]]],
  [19, 5762, "RÉSIDENCE CHÂTEAUBRIAND", "9 BOULEVARD DE LA LIBERTÉ", "59800", "LILLE", [["278", "Appartement"]]],
  [19, 175, "OPÉRA", "1 RUE DE LA SCÈNE", "59000", "LILLE", [["9", "Commerce"]]],
  [19, 242, "PARC DES LILAS", "20 AVENUE DES LILAS", "59700", "MARCQ-EN-BARŒUL", [["211", "Appartement"]]],
  [19, 300, "LES ACACIAS", "1 ALLÉE DES ACACIAS", "59100", "ROUBAIX", [["3", "Appartement"]]],
  [19, 301, "LES ACACIAS", "8 RUE VERTE", "59200", "TOURCOING", [["3", "Garage"]]],
  [19, 600, "LES TILLEULS", "4 RUE DES TILLEULS", "59000", "LILLE", [["12", "Appartement"]]],
  [19, 610, "CŒUR DE VILLE", "2 PLACE D'ARMES", "59000", "LILLE", [["0", "Parking"], ["5", "Appartement"]]],
  [31, 880, "LE ZOLA", "7 RUE ZOLA", "62000", "ARRAS", [["7", "Appartement"]]],
  [42, 1201, "LE ZOLA", "70 AVENUE ÉMILE ZOLA", "80000", "AMIENS", [["7", "Appartement"]]],
  [31, 881, "PARC FLEURI LE CONCERTO 1", "1 RUE HAUTE", "62000", "ARRAS", [["732", "Appartement"]]]
];

// Contenu des fichiers GED (texte de la page 1 du PDF généré).
const DPE_TEXT = ["DIAGNOSTIC DE PERFORMANCE ENERGETIQUE (logement)", "Etabli le : 05/06/2012", "Adresse : 12 rue des Musiciens 59000 LILLE", "Consommation conventionnelle : 105 kWh/m2/an"];
const CHAUD_TEXT = ["RAPPORT D'ENTRETIEN DE CHAUDIERE GAZ", "Intervention du technicien le 01/03/2021", "Controle de combustion et nettoyage du bruleur", "Attestation d'entretien annuel"];

// Pages des lots : formulaire DPE (form1) + documents GED [nom, date « crée le », uuid, contenu].
const LOTS = {
  "19/5004/732": {
    form: { libdpe1: "C", DPE1: "105", libges1: "JS:D", GES1: "25", typediag: "0", DATEDIAG: "05/06/2012", datefinval: "31/12/2012", MTDIAGMIN: "", MTDIAGMAX: "", MTDIAG: "574", ANNEREFDIAG: "2010", REFERENCE: "1032/5004/0732" },
    docs: [
      ["entretien chaudière 2020.pdf", "2020-01-10", U(1), CHAUD_TEXT],
      ["dpe.pdf", "2013-07-17", U(2), DPE_TEXT],
      ["EDL entrée.pdf", "2012-09-01", U(3), CHAUD_TEXT],
      ["DPE ancien.pdf", "2009-05-02", U(4), DPE_TEXT],
      ["photo façade.png", "2014-02-02", U(5), null]
    ]
  },
  "19/5305/130": { docs: [["DURAND_928_24.08.11_DP.pdf", "2011-08-24", U(6), DPE_TEXT], ["additif bail.doc", "2015-01-01", U(7), null]] },
  "19/49/543": { docs: [["dpedupontmarie.pdf", "2007-09-25", U(8), DPE_TEXT], ["courrier.msg", "2010-01-01", U(9), null]] },
  "19/5762/278": {
    docs: [["MARTIN_DPE.pdf", "2015-09-01", U(10), DPE_TEXT], ["MARTIN_DPE.pdf", "2015-09-01", U(11), DPE_TEXT], ["diagnostic gaz.pdf", "2016-01-01", U(12), CHAUD_TEXT]]
  },
  "19/175/9": { docs: [] },
  "19/242/211": { docs: [["entretien chaudiere 2019.pdf", "2019-10-10", U(13), CHAUD_TEXT], ["additif bail.msg", "2018-03-03", U(14), null]] },
  "19/600/12": { docs: [["DPE 2021.pdf", "2021-03-01", U(15), CHAUD_TEXT], ["dpe.pdf", "2016-02-02", U(16), DPE_TEXT]] },
  "19/610/5": { docs: [["Diagnostic de performance énergétique.pdf", "2019-04-04", U(17), DPE_TEXT]] }
};

// Les 6 cas réels attendus (anonymisés), cabinet Sicov (19).
const REAL_CASES = [
  { nom: "PARC FLEURI LE CONCERTO 1", lot: "732", imm: "5004", status: "OK", file: "dpe.pdf", date: "2013-07-17", uuid: U(2) },
  { nom: "COUR SAINT-HUBERT", lot: "130", imm: "5305", status: "OK", file: "DURAND_928_24.08.11_DP.pdf", date: "2011-08-24", uuid: U(6) },
  { nom: "RESIDENCE BEAUSEJOUR", lot: "543", imm: "49", status: "OK", file: "dpedupontmarie.pdf", date: "2007-09-25", uuid: U(8) },
  { nom: "RESIDENCE CHATEAUBRIAND", lot: "278", imm: "5762", status: "OK", file: "MARTIN_DPE.pdf", date: "2015-09-01", uuid: U(10) },
  { nom: "OPERA", lot: "9", imm: "175", status: "AUCUN_DPE_GED", nDocs: 0 },
  { nom: "PARC DES LILAS", lot: "211", imm: "242", status: "AUCUN_DPE_GED", nDocs: 2 }
];

// windows-1252 (Latin-1 + Œ œ ’ € … – —).
const CP1252 = { "€": 0x80, "…": 0x85, "’": 0x92, "–": 0x96, "—": 0x97, "Œ": 0x8c, "œ": 0x9c };
function enc1252(s) {
  const out = Buffer.alloc(s.length);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    out[i] = c < 0x80 || (c >= 0xa0 && c < 0x100) ? c : CP1252[s[i]] ?? 0x3f;
  }
  return out;
}

const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function cabSelect(selected) {
  const opts = Object.entries(CAB).map(([id, l]) => `<option value="${id}"${String(id) === String(selected) ? " selected" : ""}>${l}</option>`);
  opts.push(`<option value="100"${String(selected) === "100" ? " selected" : ""}>Tous les cabinets</option>`);
  return `<select name="ID_SOCIETE" class="form-control">${opts.join("")}</select>`;
}

function diagForm(soc) {
  return `<form name="recherche" method="get" action="diag.php"><input type="hidden" name="action" value="R">
Cabinet : ${cabSelect(soc)} Gestionnaire : <select name="ID_GEST"><option value="0">Tous</option></select>
Type de bail : <select name="TYPEB"><option value="0">Tous</option><option value="1">Habitation</option></select>
<input type="checkbox" name="DPE_TOUS" value="checked" checked> Tous les DPE <input type="submit" value="Rechercher"></form>`;
}

function diagHtml(soc, extraRows = "") {
  const rows = [];
  let k = 0;
  for (const [s, imm, nom, adr, cp, ville, lots] of IMMS) {
    if (String(soc) !== "100" && String(s) !== String(soc)) continue;
    for (const [lot, type] of lots) {
      k++;
      const amp = k % 2 ? "&amp;" : "&"; // les deux formes existent
      const href = `../gerance/diaglot.php?id_societe=${s}${amp}id_immeuble=${imm}${amp}id_lot=${lot}`;
      const cells = [CAB[s], imm, nom, adr, cp, ville, lot, type, "NC"].map((v) => `<a href="${href}">${esc(v)}</a>`);
      // Une ligne sur trois sans balises </td> (HTML PHP mal fermé).
      rows.push(k % 3 ? `<tr class="${k % 2 ? "odd" : "even"}">${cells.map((c) => `<td>${c}</td>`).join("")}</tr>` : `<tr>${cells.map((c) => `<td>${c}`).join("")}\n`);
    }
  }
  return `<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><title>Diagnostics</title>
<script src="js/jquery.dataTables.min.js"></script></head><body>
<h3>Diagnostics de performance énergétique</h3>${diagForm(soc)}
<input type="search" placeholder="Recherche rapide">
<table id="example3" class="display table table-striped"><thead><tr><th>Cabinet</th><th>N° Immeuble</th><th>Nom Immeuble</th><th>Adresse</th><th>CP</th><th>Ville</th><th>N° Lot</th><th>Type de lot</th><th>DPE</th></tr></thead>
<tbody>
${rows.join("\n")}${extraRows}
</tbody></table>
<script>$('#example3').DataTable({ pageLength: 10 });</script></body></html>`;
}

function diagFormPage() {
  return `<!DOCTYPE html><html><head><title>Diagnostics</title></head><body><h3>Diagnostics</h3>${diagForm(19)}</body></html>`;
}

function loginPage() {
  return `<!DOCTYPE html><html><head><title>Intralia - Identification</title></head><body>
<form name="login" method="post" action="/index.php"><input type="text" name="login"><input type="password" name="pass"><input type="submit" value="Connexion"></form></body></html>`;
}

function letterSelect(name, v) {
  const opts = ["", "A", "B", "C", "D", "E", "F", "G"].map((l) => `<option value="${l}"${l === v ? " selected" : ""}>${l}</option>`).join("");
  return `<select name="${name}">${opts}</select>`;
}

function lotHtml(soc, imm, lot) {
  const L = LOTS[`${soc}/${imm}/${lot}`] || { docs: [] };
  const f = Object.assign({ libdpe1: "", DPE1: "", libges1: "", GES1: "", typediag: "0", DATEDIAG: "", datefinval: "", MTDIAGMIN: "", MTDIAGMAX: "", MTDIAG: "", ANNEREFDIAG: "", REFERENCE: `1032/${imm}/${String(lot).padStart(4, "0")}` }, L.form || {});
  const js = [];
  const letter = (n) => { if (String(f[n]).startsWith("JS:")) { js.push(`document.form1.${n}.value = '${f[n].slice(3)}';`); return letterSelect(n, ""); } return letterSelect(n, f[n]); };
  const inp = (n) => `<input type="text" name="${n}" value="${esc(f[n])}" size="8">`;
  const docs = L.docs.map(([name, date, uuid]) => `<tr>
  <td><a href="https://www.oralia.flowerdocs.cloud/flower-docs-gui/#/documents/edit:${uuid}" target="GED"><i class="fa fa-eye"></i></a></td>
  <td><input type="checkbox" name="IDGED" class="IDGED" value="${uuid}" onclick="majchk('${uuid}')"></td>
  <td><b>${esc(name)}</b><br><span class="label label-danger">RapportDiagnostic</span>
      <span class="label label-danger">crée le ${date}</span></td>
</tr>`).join("\n");
  return `<!DOCTYPE html><html><head><title>Diagnostic du lot ${lot}</title></head><body>
<div class="container"><table class="table"><tr><td>
<form name="form1" method="post" action="diaglot.php?id_societe=${soc}&amp;id_immeuble=${imm}&amp;id_lot=${lot}">
<input type="hidden" name="REFERENCE" value="${esc(f.REFERENCE)}">
<table>
<tr><td>Étiquette énergie</td><td>${letter("libdpe1")} ${inp("DPE1")} kWh/m²/an</td></tr>
<tr><td>Étiquette climat</td><td>${letter("libges1")} ${inp("GES1")} kgCO2/m²/an</td></tr>
<tr><td>Type de DPE</td><td><select name="typediag"><option value="0"${f.typediag === "0" ? " selected" : ""}>DPE réalisé avant le 1er juillet 2021</option><option value="1"${f.typediag === "1" ? " selected" : ""}>DPE réalisé à partir du 1er juillet 2021</option></select></td></tr>
<tr><td>Date du diagnostic</td><td>${inp("DATEDIAG")} fin de validité ${inp("datefinval")}</td></tr>
<tr><td>Montants</td><td>min ${inp("MTDIAGMIN")} max ${inp("MTDIAGMAX")} estimé ${inp("MTDIAG")} année ${inp("ANNEREFDIAG")}</td></tr>
</table>
<input type="submit" value="Enregistrer le DPE">
<button type="button" class="btn btn-danger" onclick="commandeDPE()">Passer la commande DPE</button>
</td></tr></table>
</form>
<h4>Documents de la GED</h4>
<table class="table table-condensed">
${docs}
</table>
<button type="button" onclick="rattacher()">Rattacher les documents cochés</button>
</div>
<script>${js.join("\n")}</script>
</body></html>`;
}

function unknownLotPage() {
  return `<!DOCTYPE html><html><head><title>Lot</title></head><body><p>Lot inconnu.</p></body></html>`;
}

// PDF minimal (une page, Helvetica, texte ASCII).
function makePdf(lines) {
  const content = "BT /F1 12 Tf 50 750 Td " + lines.map((l, i) => (i ? "0 -16 Td " : "") + "(" + l.replace(/[\\()]/g, "\\$&") + ") Tj").join(" ") + " ET";
  const objs = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
    `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
  ];
  let out = "%PDF-1.4\n";
  const offs = [];
  objs.forEach((o, i) => { offs.push(out.length); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const xref = out.length;
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n` + offs.map((o) => String(o).padStart(10, "0") + " 00000 n \n").join("");
  out += `trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, "latin1");
}

function docByUuid(uuid) {
  for (const L of Object.values(LOTS)) for (const d of L.docs) if (d[2] === uuid) return d;
  return null;
}

function listen(handler) {
  return new Promise((resolve) => {
    const s = http.createServer(handler);
    s.listen(0, "127.0.0.1", () => resolve({ s, origin: `http://127.0.0.1:${s.address().port}` }));
  });
}

async function start() {
  const stats = { index: [], forms: 0, lots: [], posts: [], flower: [] };
  const st = { intraLogged: true, flowerLogged: true };
  const html = (res, body, status = 200) => { res.writeHead(status, { "content-type": "text/html" }); res.end(enc1252(body)); };

  const intra = await listen((req, res) => {
    const u = new URL(req.url, "http://x");
    if (req.method !== "GET") { stats.posts.push(req.method + " " + req.url); res.writeHead(405); res.end(); return; }
    if (u.pathname === "/" || u.pathname === "/index.php") return html(res, st.intraLogged ? "<html><body><h1>Intralia</h1></body></html>" : loginPage());
    if (!st.intraLogged) return html(res, loginPage());
    if (u.pathname === "/gerance/diag.php") {
      if (u.searchParams.get("action") !== "R") { stats.forms++; return html(res, diagFormPage()); }
      stats.index.push(u.search);
      return html(res, diagHtml(u.searchParams.get("ID_SOCIETE") || "19"));
    }
    if (u.pathname === "/gerance/diaglot.php") {
      const k = ["id_societe", "id_immeuble", "id_lot"].map((x) => u.searchParams.get(x));
      stats.lots.push(k.join("/"));
      const exists = IMMS.some(([s, imm, , , , , lots]) => String(s) === k[0] && String(imm) === k[1] && lots.some(([l]) => l === k[2]));
      return html(res, exists ? lotHtml(...k) : unknownLotPage());
    }
    res.writeHead(404); res.end();
  });

  const flower = await listen((req, res) => {
    const u = new URL(req.url, "http://x");
    const cors = req.headers.origin ? {
      "access-control-allow-origin": req.headers.origin, "access-control-allow-credentials": "true",
      "access-control-expose-headers": "Content-Disposition, Content-Type"
    } : {};
    if (req.method === "OPTIONS") { res.writeHead(204, cors); res.end(); return; }
    if (u.pathname === "/flower-docs-gui/login") {
      res.writeHead(200, Object.assign({ "content-type": "text/html;charset=UTF-8" }, cors));
      res.end("<!DOCTYPE html><html><head><title>FlowerDocs - Connexion</title></head><body><form><input name=user><input type=password name=pwd></form></body></html>");
      return;
    }
    if (u.pathname === "/flower-docs-gui/download") {
      const id = u.searchParams.get("documentId");
      if (!st.flowerLogged) { res.writeHead(302, Object.assign({ location: "/flower-docs-gui/login" }, cors)); res.end(); return; }
      const d = docByUuid(id);
      if (!d) { res.writeHead(404, cors); res.end("not found"); return; }
      stats.flower.push(id);
      const body = d[3] ? makePdf(d[3]) : Buffer.from("PK\u0003\u0004 fichier non PDF", "latin1");
      res.writeHead(200, Object.assign({
        "content-type": "application/octet-stream",
        "content-disposition": `attachment;filename=${encodeURIComponent(d[0])};filename*=utf-8''${encodeURIComponent(d[0])}`
      }, cors));
      res.end(body);
      return;
    }
    res.writeHead(404, cors); res.end();
  });

  return {
    intraOrigin: intra.origin, flowerOrigin: flower.origin, stats,
    setIntraLogged: (v) => { st.intraLogged = !!v; },
    setFlowerLogged: (v) => { st.flowerLogged = !!v; },
    close: () => Promise.all([new Promise((r) => intra.s.close(r)), new Promise((r) => flower.s.close(r))])
  };
}

module.exports = { start, CAB, IMMS, LOTS, U, REAL_CASES, DPE_TEXT, CHAUD_TEXT, enc1252, diagHtml, diagFormPage, lotHtml, loginPage, unknownLotPage, makePdf };
