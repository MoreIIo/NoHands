// Faux serveur SIGEO pour les tests de « Recherche d'id interne de lot ».
// Reproduit ce qui compte côté ASP.NET WebForms :
//   - GET  lot_browse/syn : formulaire avec __VIEWSTATE / __EVENTVALIDATION,
//     deux boutons « Rechercher » (bloc Nom = piège, bloc Par code = le bon) ;
//   - POST lot_browse/syn : refusé (HTTP 500 « Invalid postback ») si le
//     viewstate / l'event validation / le bouton ne correspondent pas ;
//     recherche LT en « contient » (204 renvoie aussi 2040) ; champs renvoyés ;
//   - session : cookie sid=ok, sinon redirection vers login.aspx.
// Modes (GET /__mode?m=…) : input (bouton submit nommé), link (__doPostBack),
// nofetch (POST refusé hors navigation de formulaire → force le repli onglet).
const http = require("http");

const DB = {
  393528: [{ lt: 204, id: 28, type: "Appartement Type 4" }, { lt: 2040, id: 29, type: "Cave" }, { lt: 205, id: 30, type: "Parking", statut: "inactif" }],
  393271: [{ lt: 305, id: 7, type: "Appartement Type 2" }, { lt: 301, id: 8, type: "Studio" }],
  393151: [{ lt: 260, id: 40, type: "Appartement Type 3" }, { lt: 260, id: 41, type: "Appartement Type 3" }],
  393710: [{ lt: 2019, id: 3, type: "Local commercial" }]
};
const BASE = "/page.aspx/fr/sig_imm/";

function start(port = 0) {
  let mode = "input";
  let seq = 0;
  const issued = new Set();
  const stats = { get: 0, post: 0, rejected: 0 };

  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
  function page({ lt = "", im = "", statut = "actif", results = null }) {
    const n = ++seq;
    issued.add("VS-" + n);
    const btn = mode === "link"
      ? `<a id="body_x_lnkLot" href="javascript:__doPostBack('ctl00$body$x$lnkLot','')">Rechercher</a>`
      : `<input type="submit" name="ctl00$body$x$btnLot" value="Rechercher" onclick="javascript:WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions(&quot;ctl00$body$x$btnLot&quot;, &quot;&quot;, true, &quot;&quot;, &quot;&quot;, false, false))" />`;
    const opt = (v, l) => `<option value="${v}"${statut === v ? " selected" : ""}>${l}</option>`;
    let rows = "";
    if (results) {
      rows = results.length ? results.map(({ imNum, l }) => {
        const href = `${BASE}lot_manage/syn/${imNum}/${l.id}`;
        return `<tr><td><a href="${href}">N°&nbsp;LT${String(l.lt).padStart(6, "0")}&nbsp;${esc(l.type)}</a></td>` +
          `<td><a href="${href}"><img alt="voir" src="data:,"></a></td></tr>`;
      }).join("") : `<tr><td>Aucun lot trouvé</td></tr>`;
    }
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Lots</title>
<script>
function __doPostBack(t, a) { var f = document.getElementById("aspnetForm"); f.__EVENTTARGET.value = t; f.__EVENTARGUMENT.value = a; f.submit(); }
function WebForm_PostBackOptions() {} function WebForm_DoPostBackWithOptions() {}
</script></head><body>
<form method="post" action="./syn" id="aspnetForm">
<input type="hidden" name="__EVENTTARGET" id="__EVENTTARGET" value="" />
<input type="hidden" name="__EVENTARGUMENT" id="__EVENTARGUMENT" value="" />
<input type="hidden" name="__VIEWSTATE" id="__VIEWSTATE" value="VS-${n}" />
<input type="hidden" name="__VIEWSTATEGENERATOR" value="A1B2C3D4" />
<input type="hidden" name="__EVENTVALIDATION" value="EV-${n}" />
<div id="menu"><a href="${BASE}imm_browse/syn">Immeubles</a></div>
<div class="bloc"><h3>Par nom</h3>
  <input name="ctl00$body$x$txtNom" id="body_x_txtNom" value="" />
  <input type="submit" name="ctl00$body$x$btnNom" value="Rechercher" />
</div>
<div class="bloc"><h3>Par code</h3><table><tr>
  <td>LT <input name="ctl00$body$x$txtLotId$x$txtLotId" id="body_x_txtLotId_x_txtLotId" value="${esc(lt)}" /></td>
  <td>Niv. <input name="ctl00$body$x$txtBatNiv$x$txtBatNiv" id="body_x_txtBatNiv_x_txtBatNiv" value="" /></td>
  <td>IM <input name="ctl00$body$x$txtImmId$x$txtImmId" id="body_x_txtImmId_x_txtImmId" value="${esc(im)}" /></td>
  <td><select name="ctl00$body$x$ddlLotStatut$ddlLotStatut" id="body_x_ddlLotStatut_ddlLotStatut">${opt("", "Tous")}${opt("actif", "Actif")}${opt("inactif", "Inactif")}</select></td>
  <td>${btn}</td></tr></table>
</div>
<table id="results">${rows}</table>
</form></body></html>`;
  }

  const loginPage = `<!DOCTYPE html><html><body><form method="post" action="/login.aspx">
<input name="user" /><input type="password" name="pwd" /><input type="submit" value="Connexion" /></form></body></html>`;

  const server = http.createServer((req, res) => {
    const url = new URL(req.url, "http://x");
    const send = (code, html, headers = {}) => { res.writeHead(code, Object.assign({ "Content-Type": "text/html; charset=utf-8" }, headers)); res.end(html); };
    if (url.pathname === "/__mode") { mode = url.searchParams.get("m") || "input"; return send(200, "mode " + mode); }
    if (url.pathname === "/__login") return send(200, "ok", { "Set-Cookie": "sid=ok; Path=/" });
    if (url.pathname === "/login.aspx") return send(200, loginPage);
    const authed = /(?:^|;\s*)sid=ok/.test(req.headers.cookie || "");
    if (url.pathname.startsWith(BASE) && !authed) return send(302, "", { Location: "/login.aspx?ReturnUrl=" + encodeURIComponent(url.pathname) });
    if (url.pathname.startsWith(BASE + "lot_manage/syn/")) return send(200, "<html><body><h1>Fiche lot</h1></body></html>");
    if (url.pathname !== BASE + "lot_browse/syn") return send(404, "not found");
    if (req.method === "GET") { stats.get++; return send(200, page({})); }

    let body = "";
    req.on("data", (c) => { body += c; });
    req.on("end", () => {
      stats.post++;
      const f = new URLSearchParams(body);
      const vs = f.get("__VIEWSTATE") || "";
      const bad = (why) => { stats.rejected++; send(500, `<html><body><h1>Server Error</h1><p>Invalid postback or callback argument. ${why}</p></body></html>`); };
      if (!issued.has(vs) || f.get("__EVENTVALIDATION") !== vs.replace("VS", "EV") || f.get("__VIEWSTATEGENERATOR") !== "A1B2C3D4") return bad("viewstate");
      if (mode === "nofetch" && req.headers["sec-fetch-mode"] !== "navigate") return bad("fetch refusé");
      const lt = f.get("ctl00$body$x$txtLotId$x$txtLotId") || "";
      const im = f.get("ctl00$body$x$txtImmId$x$txtImmId") || "";
      const statut = f.get("ctl00$body$x$ddlLotStatut$ddlLotStatut") ?? "actif";
      const lotBtn = mode === "link" ? f.get("__EVENTTARGET") === "ctl00$body$x$lnkLot" : f.get("ctl00$body$x$btnLot") === "Rechercher";
      if (!lotBtn) {
        if (f.get("ctl00$body$x$btnNom") === "Rechercher") return send(200, page({ lt, im, statut, results: [] })); // recherche par nom : pas de lots
        return bad("bouton");
      }
      const imNum = parseInt(im, 10);
      const ltDigits = String(lt).replace(/\D/g, "").replace(/^0+/, "");
      const results = (DB[imNum] || [])
        .filter((l) => String(l.lt).includes(ltDigits))
        .filter((l) => !statut || (l.statut || "actif") === statut)
        .map((l) => ({ imNum, l }));
      send(200, page({ lt, im, statut, results }));
    });
  });
  return new Promise((resolve) => server.listen(port, "127.0.0.1", () => {
    const origin = `http://127.0.0.1:${server.address().port}`;
    resolve({ server, origin, url: origin + BASE + "lot_browse/syn", stats, setMode: (m) => { mode = m; } });
  }));
}

module.exports = { start, DB };
