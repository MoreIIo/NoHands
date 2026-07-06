// update-check.js — NoHands OSA
// -----------------------------------------------------------------------------
// Système de vérification de mise à jour par rapport au dépôt GitHub.
//
// Fonctionnement :
//   • Compare la version installée (manifest.json local) à celle publiée sur
//     GitHub (manifest.json de la branche `main`).
//   • Récupère les notes de version (dernière release ou, à défaut, les derniers
//     messages de commit).
//   • Affiche une bannière si une version plus récente existe.
//   • Le dépôt étant PRIVÉ, l'appel à l'API GitHub nécessite un token de lecture.
//     Ce token est saisi par l'utilisateur dans le modal « Mises à jour » et
//     stocké UNIQUEMENT dans chrome.storage.local — jamais dans le code, jamais
//     poussé sur GitHub. Créez un « fine-grained token » limité au dépôt
//     MoreIIo/NoHands avec la permission « Contents: Read-only ».
//
// Aucune modification du manifest n'est nécessaire : host_permissions inclut
// déjà <all_urls>, ce qui autorise les appels à api.github.com.
// -----------------------------------------------------------------------------

(function () {
  "use strict";

  // ---- Configuration du dépôt --------------------------------------------
  const REPO_OWNER = "MoreIIo";
  const REPO_NAME  = "NoHands";
  const REPO_BRANCH = "main";
  const REPO_URL = `https://github.com/${REPO_OWNER}/${REPO_NAME}`;
  const API_BASE = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}`;

  // Intervalle mini entre deux vérifications automatiques (3 h).
  const AUTO_CHECK_INTERVAL_MS = 3 * 60 * 60 * 1000;

  // Clés de stockage local.
  const K_TOKEN      = "osaUpdate.token";
  const K_AUTOCHECK  = "osaUpdate.autoCheck";
  const K_LASTCHECK  = "osaUpdate.lastCheck";
  const K_DISMISSED  = "osaUpdate.dismissedVersion";

  // ---- Utilitaires storage ------------------------------------------------
  const store = {
    get(keys) {
      return new Promise((res) => chrome.storage.local.get(keys, res));
    },
    set(obj) {
      return new Promise((res) => chrome.storage.local.set(obj, res));
    },
  };

  // ---- Version -----------------------------------------------------------
  function localVersion() {
    try {
      return chrome.runtime.getManifest().version || "0.0.0";
    } catch (_) {
      return "0.0.0";
    }
  }

  // Compare deux versions semver simples ("2.1.0" > "2.0.3"). Renvoie true si a > b.
  function versionGreater(a, b) {
    const pa = String(a).split(".").map((n) => parseInt(n, 10) || 0);
    const pb = String(b).split(".").map((n) => parseInt(n, 10) || 0);
    const len = Math.max(pa.length, pb.length);
    for (let i = 0; i < len; i++) {
      const x = pa[i] || 0;
      const y = pb[i] || 0;
      if (x > y) return true;
      if (x < y) return false;
    }
    return false;
  }

  // ---- Appels API GitHub -------------------------------------------------
  async function ghFetch(url, token) {
    const headers = {
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
    };
    if (token) headers.Authorization = `Bearer ${token}`;
    const resp = await fetch(url, { headers, cache: "no-store" });
    if (!resp.ok) {
      const err = new Error(`GitHub API ${resp.status}`);
      err.status = resp.status;
      throw err;
    }
    return resp.json();
  }

  // Décode un contenu base64 (avec sauts de ligne) en UTF-8.
  function decodeBase64Utf8(b64) {
    const clean = String(b64).replace(/\s/g, "");
    const bin = atob(clean);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return new TextDecoder("utf-8").decode(bytes);
  }

  // Récupère la version publiée depuis manifest.json de la branche.
  async function fetchRemoteVersion(token) {
    const data = await ghFetch(
      `${API_BASE}/contents/manifest.json?ref=${REPO_BRANCH}`,
      token
    );
    const raw = decodeBase64Utf8(data.content || "");
    const manifest = JSON.parse(raw);
    return manifest.version || "0.0.0";
  }

  // Récupère les notes de version : dernière release, sinon derniers commits.
  async function fetchReleaseNotes(token) {
    // 1) Tenter la dernière release publiée.
    try {
      const rel = await ghFetch(`${API_BASE}/releases/latest`, token);
      return {
        kind: "release",
        title: rel.name || rel.tag_name || "Dernière version",
        body: (rel.body || "").trim(),
        url: rel.html_url || REPO_URL,
      };
    } catch (e) {
      if (e.status && e.status !== 404) throw e; // 404 = pas de release, on continue
    }
    // 2) À défaut : derniers messages de commit de la branche.
    const commits = await ghFetch(
      `${API_BASE}/commits?sha=${REPO_BRANCH}&per_page=15`,
      token
    );
    const items = (commits || []).map((c) => {
      const msg = (c.commit && c.commit.message) || "";
      return msg.split("\n")[0].trim(); // sujet uniquement
    });
    return {
      kind: "commits",
      title: "Derniers changements",
      body: items.map((s) => `• ${s}`).join("\n"),
      url: `${REPO_URL}/commits/${REPO_BRANCH}`,
    };
  }

  // ---- Références DOM (résolues au chargement) ---------------------------
  let el = {};
  function cacheDom() {
    el = {
      banner:      document.getElementById("updateBanner"),
      bannerVer:   document.getElementById("updateBannerVersion"),
      bannerLink:  document.getElementById("updateBannerLink"),
      bannerClose: document.getElementById("updateBannerClose"),
      notesToggle: document.getElementById("updateNotesToggle"),
      notes:       document.getElementById("updateNotes"),
      // Modal
      openBtn:     document.getElementById("openUpdateBtn"),
      modal:       document.getElementById("updateModal"),
      closeBtn:    document.getElementById("closeUpdateBtn"),
      localVer:    document.getElementById("updLocalVer"),
      remoteVer:   document.getElementById("updRemoteVer"),
      tokenInput:  document.getElementById("updTokenInput"),
      autoCheck:   document.getElementById("updAutoCheck"),
      saveTokenBtn:document.getElementById("updSaveTokenBtn"),
      checkNowBtn: document.getElementById("updCheckNowBtn"),
      statusBox:   document.getElementById("updStatus"),
      repoLink:    document.getElementById("updRepoLink"),
    };
  }

  // ---- Rendu ------------------------------------------------------------
  function showBanner(remoteVer, notes) {
    if (!el.banner) return;
    el.bannerVer.textContent = `v${localVersion()} → v${remoteVer}`;
    el.bannerLink.href = notes && notes.url ? notes.url : REPO_URL;
    if (notes && notes.body) {
      el.notes.textContent = notes.body;
    } else {
      el.notes.textContent = "";
    }
    el.banner.hidden = false;
  }

  function setStatus(msg, type) {
    if (!el.statusBox) return;
    el.statusBox.textContent = msg || "";
    el.statusBox.className = "hint upd-status" + (type ? " " + type : "");
  }

  function humanError(e) {
    if (e && e.status === 401) return "Token invalide ou expiré (401).";
    if (e && e.status === 403) return "Accès refusé ou limite d'API atteinte (403).";
    if (e && e.status === 404)
      return "Dépôt ou fichier introuvable (404) — vérifiez le token et ses droits sur le dépôt.";
    if (e && /Failed to fetch|NetworkError/i.test(e.message || ""))
      return "Pas de connexion à GitHub.";
    return "Erreur : " + (e && e.message ? e.message : "inconnue");
  }

  // ---- Logique de vérification ------------------------------------------
  // silent=true : appel automatique (pas de message si à jour / pas de token).
  async function runCheck(silent) {
    const cfg = await store.get([K_TOKEN, K_AUTOCHECK]);
    const token = cfg[K_TOKEN];

    if (!token) {
      if (!silent) setStatus("Ajoutez d'abord un token GitHub, puis enregistrez.", "err");
      return;
    }

    if (!silent) setStatus("Vérification en cours…", "");
    try {
      const remoteVer = await fetchRemoteVersion(token);
      await store.set({ [K_LASTCHECK]: Date.now() });
      if (el.remoteVer) el.remoteVer.textContent = "v" + remoteVer;

      const local = localVersion();
      if (versionGreater(remoteVer, local)) {
        let notes = null;
        try { notes = await fetchReleaseNotes(token); } catch (_) { /* notes best-effort */ }

        // Ne pas ré-afficher une version explicitement masquée (auto seulement).
        const dism = (await store.get([K_DISMISSED]))[K_DISMISSED];
        if (!(silent && dism === remoteVer)) {
          showBanner(remoteVer, notes);
        }
        if (!silent) {
          setStatus(`Nouvelle version disponible : v${remoteVer}.`, "ok");
          if (notes && notes.body) {
            el.notes.textContent = notes.body;
            el.notes.hidden = false;
          }
        }
      } else {
        if (el.banner) el.banner.hidden = true;
        if (!silent) setStatus(`À jour (v${local}).`, "ok");
      }
    } catch (e) {
      if (!silent) setStatus(humanError(e), "err");
      else console.debug("NoHands OSA — auto update-check:", e.message);
    }
  }

  // ---- Modal -------------------------------------------------------------
  async function openModal() {
    if (!el.modal) return;
    const cfg = await store.get([K_TOKEN, K_AUTOCHECK]);
    if (el.localVer) el.localVer.textContent = "v" + localVersion();
    if (el.tokenInput) el.tokenInput.value = cfg[K_TOKEN] || "";
    if (el.autoCheck) el.autoCheck.checked = cfg[K_AUTOCHECK] !== false; // défaut activé
    if (el.repoLink) el.repoLink.href = REPO_URL;
    setStatus("", "");
    if (el.notes) el.notes.hidden = true;
    el.modal.hidden = false;
  }
  function closeModal() {
    if (el.modal) el.modal.hidden = true;
  }

  async function saveToken() {
    const token = (el.tokenInput && el.tokenInput.value || "").trim();
    const auto = el.autoCheck ? el.autoCheck.checked : true;
    await store.set({ [K_TOKEN]: token, [K_AUTOCHECK]: auto });
    setStatus(token ? "Token enregistré (stocké localement)." : "Token effacé.", "ok");
  }

  async function dismissBanner() {
    if (el.banner) el.banner.hidden = true;
    // Mémorise la version publiée pour ne plus la re-signaler automatiquement.
    const cfg = await store.get([K_TOKEN]);
    if (cfg[K_TOKEN]) {
      try {
        const remoteVer = await fetchRemoteVersion(cfg[K_TOKEN]);
        await store.set({ [K_DISMISSED]: remoteVer });
      } catch (_) { /* ignore */ }
    }
  }

  // ---- Câblage des événements -------------------------------------------
  function wire() {
    if (el.openBtn)  el.openBtn.addEventListener("click", openModal);
    if (el.closeBtn) el.closeBtn.addEventListener("click", closeModal);
    if (el.modal) {
      el.modal.addEventListener("click", (ev) => {
        if (ev.target === el.modal) closeModal(); // clic hors carte
      });
    }
    if (el.saveTokenBtn) el.saveTokenBtn.addEventListener("click", saveToken);
    if (el.checkNowBtn)  el.checkNowBtn.addEventListener("click", () => runCheck(false));
    if (el.bannerClose)  el.bannerClose.addEventListener("click", dismissBanner);
    if (el.bannerLink)   el.bannerLink.addEventListener("click", () => { /* lien natif */ });
    if (el.notesToggle && el.notes) {
      el.notesToggle.addEventListener("click", () => {
        el.notes.hidden = !el.notes.hidden;
      });
    }
  }

  // ---- Démarrage ---------------------------------------------------------
  async function init() {
    cacheDom();
    wire();

    // Vérification automatique throttlée à l'ouverture du panneau.
    const cfg = await store.get([K_TOKEN, K_AUTOCHECK, K_LASTCHECK]);
    const autoOn = cfg[K_AUTOCHECK] !== false;
    const hasToken = !!cfg[K_TOKEN];
    const last = cfg[K_LASTCHECK] || 0;
    if (autoOn && hasToken && Date.now() - last > AUTO_CHECK_INTERVAL_MS) {
      runCheck(true);
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
