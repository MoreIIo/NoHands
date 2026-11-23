// Harnais commun des tests d'intégration de l'étape « SIGEO : saisir un DPE ».
// Le vrai sidepanel.html est chargé dans un onglet Playwright ; chrome.scripting
// et chrome.tabs sont redirigés vers un second onglet ouvert sur le faux SIGEO
// (tests/fixtures/sigeo-diag-server.js) ; chrome.storage.session est simulé.
// world: "MAIN" est ignoré : tout s'exécute dans le monde de la page.
const path = require("path");

const PANEL = "file://" + path.join(path.resolve(__dirname, "..", ".."), "sidepanel.html");

const CHROME_STUB = `
(() => {
  const updated = new Set();
  const mem = {};
  window.__sessionMem = mem;
  window.chrome = {
    storage: {
      local: {
        get: (k, cb) => { const r = {}; if (typeof cb === "function") cb(r); return Promise.resolve(r); },
        set: (o, cb) => { if (typeof cb === "function") cb(); return Promise.resolve(); },
        remove: (k, cb) => { if (typeof cb === "function") cb(); return Promise.resolve(); }
      },
      session: {
        get: async (k) => {
          if (k === null || k === undefined) return Object.assign({}, mem);
          const r = {};
          (Array.isArray(k) ? k : [k]).forEach((x) => { if (x in mem) r[x] = JSON.parse(JSON.stringify(mem[x])); });
          return r;
        },
        set: async (o) => { Object.assign(mem, JSON.parse(JSON.stringify(o))); },
        remove: async (k) => { (Array.isArray(k) ? k : [k]).forEach((x) => delete mem[x]); }
      }
    },
    runtime: { getURL: (p) => p, getManifest: () => ({ version: "0.0.0" }), sendMessage: async () => ({}),
               lastError: null, onMessage: { addListener() {} } },
    tabs: { get: async (id) => ({ id, url: await window.__tabUrl(), status: "complete" }),
            query: async () => [{ id: 1, url: await window.__tabUrl(), active: true }],
            sendMessage: async () => ({}), duplicate: async () => ({}),
            update: (id, o, cb) => { window.__navigate(o.url).then(() => { if (cb) cb({ id }); updated.forEach((l) => l(id, { status: "complete" })); }); },
            onUpdated: { addListener(l) { updated.add(l); }, removeListener(l) { updated.delete(l); } }, onRemoved: { addListener() {} } },
    scripting: { executeScript: async ({ func, args }) => {
      const r = await window.__runInTab(func.toString(), args || []);
      return [{ frameId: 0, result: r }];
    } },
    permissions: { contains: async () => false, request: async () => false },
    windows: { update: async () => ({}) }
  };
})();`;

// PDF minimal avec une couche texte (lignes ASCII).
function makePdf(lines) {
  const content = lines.length
    ? "BT /F1 12 Tf 50 780 Td " + lines.map((l, i) => (i ? "0 -16 Td " : "") + `(${l.replace(/[()\\]/g, "\\$&")}) Tj `).join("") + "ET"
    : "q Q";
  const objs = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
    `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
  ];
  let out = "%PDF-1.4\n";
  const offs = [];
  objs.forEach((o, i) => { offs.push(out.length); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const xref = out.length;
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n` + offs.map((o) => String(o).padStart(10, "0") + " 00000 n \n").join("");
  out += `trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, "latin1");
}

const COLS = ["type_local", "code_fournisseur", "date_etabli", "date_validite", "date_debut_validite", "numero_ademe", "conso_ep", "emission_ges",
  "lettre_ep_forcee", "lettre_ges_forcee", "altitude_800m", "nouvelles_valeurs", "cout_min", "cout_max", "date_prix",
  "resultat", "commentaire", "fichier_dpe", "lot"];
const row = (o) => COLS.map((c) => o[c] ?? "");

// Ouvre le panneau et l'onglet SIGEO. → { browser, ctx, tab, panel, errors }
async function openPanel(chromium, startUrl) {
  const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/usr/bin/google-chrome", headless: true });
  const ctx = await browser.newContext();
  const tab = await ctx.newPage();
  await tab.goto(startUrl);
  const panel = await ctx.newPage();
  const errors = [];
  panel.on("pageerror", (e) => errors.push(e.message));
  tab.on("dialog", (d) => d.accept());
  await panel.exposeFunction("__tabUrl", () => tab.url());
  await panel.exposeFunction("__navigate", async (u) => { await tab.goto(u); });
  await panel.exposeFunction("__runInTab", async (src, args) => {
    for (let i = 0; ; i++) {
      try { return await tab.evaluate(([s, a]) => (new Function("return (" + s + ")"))()(...a), [src, args]); }
      catch (e) {
        // Page en cours de remplacement : comme executeScript, on échoue vite
        // (le panneau relit ensuite) au lieu de bloquer.
        if (i >= 8 || !/context was destroyed|navigat|Execution context/i.test(e.message)) throw e;
        await new Promise((r) => setTimeout(r, 100));
      }
    }
  });
  await panel.addInitScript(CHROME_STUB);
  await panel.goto(PANEL);
  await panel.waitForTimeout(500);
  return { browser, ctx, tab, panel, errors };
}

// Lance l'étape sur la ligne active ; onConfirm(texte) répond à la confirmation d'upload.
async function runActive(panel, idx, onConfirm) {
  await panel.evaluate((i) => { selectRow(i); $("scnLog").innerHTML = ""; $("runScenarioBtn").click(); }, idx);
  let confirmed = null;
  for (let t = 0; t < 1200; t++) {
    const st = await panel.evaluate(() => ({ running: scnRunning, confirm: !$("scnConfirm").hidden, body: $("scnConfirmBody").innerText }));
    if (st.confirm && onConfirm && confirmed === null) { confirmed = st.body; await onConfirm(st.body); }
    if (!st.running && t > 2) break;
    await new Promise((r) => setTimeout(r, 100));
  }
  return { confirmed, log: await panel.evaluate(() => $("scnLog").innerText) };
}

module.exports = { PANEL, CHROME_STUB, makePdf, COLS, row, openPanel, runActive };
