// Test du téléchargement DPE côté service worker (background.js,
// dpeDownload) avec un faux chrome.downloads : OK, page HTML (fichier
// supprimé), interruption, nom réel après « uniquify ».
//   node tests/sigeo-dpe.background.js
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const src = fs.readFileSync(path.join(__dirname, "..", "background.js"), "utf8");
const start = src.indexOf("function dpeWaitDownload");
const end = src.indexOf("chrome.runtime.onMessage.addListener", start);
assert.ok(start > 0 && end > start, "fonctions dpe introuvables dans background.js");

function fakeChrome(scenario) {
  const listeners = new Set();
  const items = new Map();
  const calls = { download: [], removeFile: [], erase: [] };
  let seq = 0;
  return {
    calls,
    downloads: {
      download: async (o) => {
        calls.download.push(o);
        const id = ++seq;
        items.set(id, { id, state: "in_progress", filename: "C:\\Users\\m\\Downloads\\" + o.filename.replace(/\//g, "\\"), mime: "" });
        setTimeout(() => {
          Object.assign(items.get(id), scenario(o));
          listeners.forEach((l) => l({ id, state: { current: items.get(id).state } }));
        }, 30);
        return id;
      },
      search: async ({ id }) => (items.has(id) ? [Object.assign({}, items.get(id))] : []),
      removeFile: async (id) => { calls.removeFile.push(id); },
      erase: async ({ id }) => { calls.erase.push(id); },
      onChanged: { addListener: (l) => listeners.add(l), removeListener: (l) => listeners.delete(l) }
    }
  };
}

function load(chrome) {
  const ctx = vm.createContext({ chrome, setTimeout, clearTimeout, Promise, String, Math, Object });
  vm.runInContext(src.slice(start, end), ctx);
  return ctx.dpeDownload;
}

let failures = 0;
async function check(name, fn) {
  try { await fn(); console.log("✓ " + name); }
  catch (e) { failures++; console.log("✗ " + name + "\n   " + (e.stack || e).toString().split("\n").slice(0, 3).join("\n   ")); }
}

(async () => {
  const req = { url: "http://sigeo.evoriel.net/bare.aspx/fr/fil/ged_file_download/123", filename: "DPE_SIGEO/IM0393664_10_DPE_2022-10-19.pdf", timeoutMs: 5000 };

  await check("PDF : download({url, filename, uniquify, saveAs:false}) → OK + nom réel", async () => {
    const c = fakeChrome(() => ({ state: "complete", mime: "application/pdf", fileSize: 1234,
      filename: "C:\\Users\\m\\Downloads\\DPE_SIGEO\\IM0393664_10_DPE_2022-10-19 (1).pdf" }));
    const r = await load(c)(req);
    assert.strictEqual(JSON.stringify(c.calls.download), JSON.stringify([{ url: req.url, filename: req.filename, conflictAction: "uniquify", saveAs: false }]));
    assert.strictEqual(r.ok, true);
    assert.strictEqual(r.filename, "IM0393664_10_DPE_2022-10-19 (1).pdf");
    assert.strictEqual(r.size, 1234);
  });

  await check("page HTML : fichier supprimé (removeFile + erase), KO", async () => {
    const c = fakeChrome(() => ({ state: "complete", mime: "text/html" }));
    const r = await load(c)(req);
    assert.strictEqual(r.ok, false);
    assert.strictEqual(r.error, "SIGEO a renvoyé une page HTML (session expirée ?)");
    assert.deepStrictEqual(c.calls.removeFile, [1]);
    assert.deepStrictEqual(c.calls.erase, [1]);
  });

  await check("interrompu : KO avec l'erreur Chrome", async () => {
    const c = fakeChrome(() => ({ state: "interrupted", error: "NETWORK_FAILED" }));
    const r = await load(c)(req);
    assert.strictEqual(r.ok, false);
    assert.ok(/interrompu \(NETWORK_FAILED\)/.test(r.error), r.error);
    assert.deepStrictEqual(c.calls.removeFile, []);
  });

  console.log(failures ? `\n${failures} échec(s)` : "\nTous les tests passent.");
  process.exit(failures ? 1 : 0);
})();
