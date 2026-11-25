// Faux SIGEO : recherche d'immeubles « immeuble_browse/ger » telle que décrite
// le 25/11/2026 (page ASP.NET WebForms, formulaire dans le document principal) :
//   - « Rechercher » envoie TOUT le formulaire (POST) : la page se recharge
//     entièrement, les champs reviennent remplis comme côté serveur ;
//   - id exacts : body_x_txtLabel_x_Label (Nom), body_x_txtSiret_x_txtSiret,
//     body_x_txtAdresse_x_txtAdresse, body_x_txtCP_x_CP, body_x_txtVille_x_Ville,
//     body_x_sites_x_ddlSite_x (Site) + body_x_sites_x_hdSiteLabel (caché),
//     body_x_ddlMetier_x (Filtre), body_x_ddlTypeMan_x (Type),
//     body_x_txtCode_x_Code (N° E.I.), codes copro / programme commercial,
//     boutons body_x_searchRaz_x_cmdSearchBtn_x__btn / …cmdRazBtn_x__btn ;
//   - Site : options m0 « Mes mandats », m1 « Mes sites », puis « <code> - <NOM> » ;
//     son onchange copie SEULEMENT le libellé dans le champ caché (pas de postback) ;
//   - onclick de Rechercher : Adresse remplie ET CP + Ville vides → alert(), annulé ;
//   - grille #body_x_grid_x : [0] icônes, [1] N° E.I. (a.edit_link →
//     immeuble_manage/ger/<id>), [2] Nom E.I., [3] Adresse ; lignes
//     tr.PowerGridItemClass / PowerGridAlternatingItemClass ; pied
//     tr.PowerGridPagerClass « N Résultat(s) ». ?nohead=1 : grille sans en-têtes.
// Le serveur filtre : Nom « contient » (sans accents), Site exact, Adresse
// « contient », CP, Ville, N° E.I. ; chaque POST est noté dans log.
const qs = require("querystring");

const SITES = [["m0", "Mes mandats"], ["m1", "Mes sites"], ["390", "390 - ORALIA SICOV - OPTIM"], ["391", "391 - ORALIA SOGIMAT"], ["392", "392 - ORALIA AUTRE"]];
const F = {
  nom: "body_x_txtLabel_x_Label", siret: "body_x_txtSiret_x_txtSiret", adresse: "body_x_txtAdresse_x_txtAdresse",
  cp: "body_x_txtCP_x_CP", ville: "body_x_txtVille_x_Ville", site: "body_x_sites_x_ddlSite_x", siteLabel: "body_x_sites_x_hdSiteLabel",
  filtre: "body_x_ddlMetier_x", type: "body_x_ddlTypeMan_x", code: "body_x_txtCode_x_Code",
  codeCopro: "body_x_txtCodeCopro_x_txtCodeCopro", codeProg: "body_x_txtCodeProgrammeCommercial_x_txtCodeProgrammeCommercial",
  search: "body_x_searchRaz_x_cmdSearchBtn_x__btn", raz: "body_x_searchRaz_x_cmdRazBtn_x__btn"
};
const PATH = "/page.aspx/fr/sig_imm/immeuble_browse/ger";

function create(data) {
  const DATA = data.map(([site, im, nom, adresse]) => ({ site, im, nom, adresse, id: String(parseInt(im.replace(/^IM/, ""), 10)) }));
  const log = [];
  const fold = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/\s+/g, " ").trim();
  const esc = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

  function filter(c) {
    if (!c.nom && !c.adresse && !c.cp && !c.ville && !c.code && !c.siret) return [];
    return DATA.filter((d) => {
      const a = " " + d.adresse + " ";
      if (/^\d+$/.test(c.site) && d.site !== c.site) return false;
      if (c.nom && !fold(d.nom).includes(fold(c.nom))) return false;
      if (c.adresse && !fold(d.adresse).includes(fold(c.adresse))) return false;
      if (c.cp && !a.includes(" " + c.cp + " ")) return false;
      if (c.ville && !fold(d.adresse).includes(fold(c.ville))) return false;
      if (c.code && fold(d.im) !== fold(c.code)) return false;
      if (c.siret) return false;
      return true;
    });
  }

  function page(c, rows, opts) {
    const v = (k) => esc(c[k] || "");
    const site = c.site || "390"; // à l'ouverture : un autre site que 391 est affiché
    const siteOpts = SITES.map(([val, t]) => `<option value="${val}"${val === site ? " selected" : ""}>${esc(t)}</option>`).join("");
    const label = (SITES.find(([val]) => val === site) || ["", ""])[1];
    const sel = (id, cur, list) => `<select id="${id}" name="${id}">` + list.map(([val, t]) => `<option value="${val}"${val === cur ? " selected" : ""}>${t}</option>`).join("") + "</select>";
    const head = opts.nohead ? "" : `<tr class="PowerGridHeaderClass"><th></th><th><a href="#">N° E.I.</a></th><th><a href="#">Nom E.I.</a></th><th><a href="#">Adresse</a></th></tr>`;
    const body = rows.map((d, i) => `<tr class="${i % 2 ? "PowerGridAlternatingItemClass" : "PowerGridItemClass"}">` +
      `<td><a href="#" title="Supprimer">&#x2715;</a> <a href="#" title="Editer">&#x270E;</a></td>` +
      `<td><a class="edit_link" href="/page.aspx/fr/sig_imm/immeuble_manage/ger/${d.id}"><img alt=""> ${d.im}</a></td>` +
      `<td>${esc(d.nom)}</td><td>${esc(d.adresse)}</td></tr>`).join("");
    return `<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><title>SIGEO - Immeubles</title></head><body>
<form id="aspnetForm" name="aspnetForm" method="post" action="${PATH}${opts.nohead ? "?nohead=1" : ""}">
<input type="hidden" name="__VIEWSTATE" value="x">
<fieldset><legend>Par désignation</legend>
  Nom <input type="text" id="${F.nom}" name="${F.nom}" value="${v("nom")}"><br>
  N° Siret <input type="text" id="${F.siret}" name="${F.siret}" value="${v("siret")}">
</fieldset>
<fieldset><legend>Par adresse</legend>
  Adresse <input type="text" id="${F.adresse}" name="${F.adresse}" value="${v("adresse")}"><br>
  Code postal <input type="text" id="${F.cp}" name="${F.cp}" size="5" value="${v("cp")}"> Ville <input type="text" id="${F.ville}" name="${F.ville}" value="${v("ville")}">
</fieldset>
<fieldset><legend>Par mandat</legend>
  Site <select id="${F.site}" name="${F.site}" onchange="changeSitebody_x_sites_x_ddlSite_x()">${siteOpts}</select>
  <input type="hidden" id="${F.siteLabel}" name="${F.siteLabel}" value="${esc(label)}"><br>
  Filtre ${sel(F.filtre, c.filtre || "", [["", "Tous"], ["G", "Gérance"], ["S", "Syndic"]])}<br>
  Type ${sel(F.type, c.type || "", [["", "Tous"], ["M", "Mandat"]])}
</fieldset>
<fieldset><legend>Par code</legend>N°E.I. IM<input type="text" id="${F.code}" name="${F.code}" value="${v("code")}"></fieldset>
<fieldset><legend>Par références pôle logement</legend>
  Code copropriété <input type="text" id="${F.codeCopro}" name="${F.codeCopro}" value="${v("codeCopro")}"><br>
  Code programme commercial <input type="text" id="${F.codeProg}" name="${F.codeProg}" value="${v("codeProg")}"><br>
  Provient du SAS CLEO <input type="checkbox" id="body_x_chkCleo" name="body_x_chkCleo">
</fieldset>
<div>
  <input type="submit" id="${F.search}" name="${F.search}" value="Rechercher"
    onclick="if (document.getElementById('${F.adresse}').value != '' &amp;&amp; document.getElementById('${F.cp}').value == '' &amp;&amp; document.getElementById('${F.ville}').value == '') { alert('Merci de préciser un code postal ou une ville'); return false; }">
  <input type="submit" id="${F.raz}" name="${F.raz}" value="Réinit.">
</div>
<table id="body_x_grid_x" class="PowerGrid" border="1">${head}${body}
  <tr class="PowerGridPagerClass"><td colspan="4"><b>${rows.length} Résultat(s)</b></td></tr>
</table>
</form>
<script>
function changeSitebody_x_sites_x_ddlSite_x() {
  var s = document.getElementById("${F.site}");
  document.getElementById("${F.siteLabel}").value = s.options[s.selectedIndex].text;
}
</script>
</body></html>`;
  }

  // Gestionnaire HTTP : true si la requête est pour cette page.
  function handle(req, res) {
    const u = new URL(req.url, "http://x");
    if (u.pathname !== PATH) return false;
    const opts = { nohead: u.searchParams.get("nohead") === "1" };
    const send = (html) => { res.writeHead(200, { "content-type": "text/html; charset=utf-8" }); res.end(html); };
    if (req.method !== "POST") { send(page({}, [], opts)); return true; }
    let raw = "";
    req.on("data", (ch) => { raw += ch; });
    req.on("end", () => {
      const b = qs.parse(raw);
      const c = {};
      for (const [k, id] of Object.entries(F)) c[k] = String(b[id] ?? "");
      c.raz = F.raz in b;
      c.searchBtn = F.search in b;
      log.push(c);
      if (c.raz) { send(page({}, [], opts)); return; }
      send(page(c, filter(c), opts));
    });
    return true;
  }
  return { handle, log, PATH, DATA };
}

module.exports = { create, PATH, F };
