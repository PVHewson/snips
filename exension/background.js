// background.js (MV3 service worker)
// - Fetches manifest/snippet files from GitHub (raw + API fallback for private/rate-limited)
// - Executes snippets in the PAGE "MAIN" world so global Xrm is accessible.
//
// The injected runner supports BOTH:
//   - snippets referencing global `Xrm`
//   - snippets referencing `ctx.Xrm`
// by calling the wrapper as (ctx, Xrm) and passing window.Xrm.

const DEFAULT_CONFIG = {
  owner: "PVHewson",
  repo: "snips",
  ref: "main",
  manifestPath: "snippets/manifest.json",
  requireConfirmForRisky: true,
  riskyLevels: ["medium", "high"]
};

async function getConfig() {
  const { config } = await chrome.storage.local.get(["config"]);
  return { ...DEFAULT_CONFIG, ...(config ?? {}) };
}

async function getToken() {
  const { githubToken } = await chrome.storage.local.get(["githubToken"]);
  return (githubToken ?? "").trim();
}

function rawUrl({ owner, repo, ref }, path) {
  const p = String(path).replace(/^\/+/, "");
  return `https://raw.githubusercontent.com/${owner}/${repo}/${ref}/${p}`;
}

function getGithubHeaders(token) {
  const headers = { Accept: "application/vnd.github+json" };
  if (token) headers["Authorization"] = `Bearer ${token}`;
  return headers;
}

async function safeRead(res) {
  try { return await res.text(); } catch { return ""; }
}

async function fetchViaGithubApi(config, rawUrlThatFailed, token) {
  const marker = `https://raw.githubusercontent.com/${config.owner}/${config.repo}/${config.ref}/`;
  const path = rawUrlThatFailed.startsWith(marker)
    ? rawUrlThatFailed.slice(marker.length)
    : null;

  if (!path) throw new Error("Could not map raw URL to repo path for API fallback.");

  const apiUrl = `https://api.github.com/repos/${config.owner}/${config.repo}/contents/${path}?ref=${encodeURIComponent(config.ref)}`;

  const res = await fetch(apiUrl, {
    method: "GET",
    headers: getGithubHeaders(token),
    cache: "no-store"
  });

  if (!res.ok) {
    const body = await safeRead(res);
    throw new Error(`GitHub API fetch failed (${res.status}) for ${apiUrl}\n${body ? "Details: " + body : ""}`);
  }

  const json = await res.json();
  if (!json || json.type !== "file" || !json.content) {
    throw new Error(`GitHub API returned non-file content for ${path}`);
  }

  const b64 = String(json.content).replace(/\n/g, "");
  const txt = atob(b64);

  // Convert to utf-8 (best effort)
  try {
    const bytes = Uint8Array.from(txt, c => c.charCodeAt(0));
    return new TextDecoder("utf-8").decode(bytes);
  } catch {
    return txt;
  }
}

async function fetchTextWithFallback(config, url) {
  // 1) Attempt raw fetch (public-friendly)
  const res = await fetch(url, { method: "GET", cache: "no-store", credentials: "omit" });
  if (res.ok) return res.text();

  // 2) If forbidden/404 and token exists, retry via GitHub API (private-friendly)
  if ([401, 403, 404].includes(res.status)) {
    const token = await getToken();
    if (token) return fetchViaGithubApi(config, url, token);
  }

  const body = await safeRead(res);
  throw new Error(`Fetch failed (${res.status}) for ${url}\n${body ? "Details: " + body : ""}`);
}

async function getActiveTabId() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) throw new Error("No active tab.");
  return tab.id;
}

async function runInMainWorld(tabId, userCode) {
  const [{ result }] = await chrome.scripting.executeScript({
    target: { tabId },
    world: "MAIN",
    func: async (code) => {
      // Runs inside the page MAIN world, so can see window.Xrm.
      if (!window.Xrm) {
        return { ok: false, error: "Xrm not found on this page (are you on a D365/Model-driven app tab?)." };
      }

      const ctx = {
        Xrm: window.Xrm,
        window,
        document,
        location,
        console,
        sleep: (ms) => new Promise(r => setTimeout(r, ms)),
        notify: (msg) => console.info("[SnippetRunner]", msg)
      };

      try {
        // Support snippets that reference global `Xrm` by passing it as a parameter.
        const wrapped = `(async (ctx, Xrm) => {\n"use strict";\n${code}\n})`;
        const fn = (0, eval)(wrapped);
        const out = await fn(ctx, window.Xrm);
        return { ok: true, out };
      } catch (e) {
        return { ok: false, error: String(e?.message ?? e), stack: String(e?.stack ?? "") };
      }
    },
    args: [String(userCode ?? "")]
  });

  return result;
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  (async () => {
    try {
      if (msg?.type === "GET_CONFIG") {
        const config = await getConfig();
        const token = await getToken();
        sendResponse({ ok: true, config, hasToken: !!token });
        return;
      }

      if (msg?.type === "SET_TOKEN") {
        const token = String(msg?.token ?? "").trim();
        await chrome.storage.local.set({ githubToken: token });
        sendResponse({ ok: true, hasToken: !!token });
        return;
      }

      if (msg?.type === "LOAD_MANIFEST") {
        const cfg = await getConfig();
        const url = rawUrl(cfg, cfg.manifestPath);
        const text = await fetchTextWithFallback(cfg, url);
        const manifest = JSON.parse(text);
        if (!manifest || !Array.isArray(manifest.snippets)) {
          throw new Error("manifest.json must contain { snippets: [...] }");
        }
        const snippets = manifest.snippets
          .slice()
          .sort((a, b) => (a.name || "").localeCompare(b.name || ""));
        sendResponse({ ok: true, snippets, manifestUrl: url });
        return;
      }

      if (msg?.type === "LOAD_SNIPPET_FILES") {
        const cfg = await getConfig();
        const s = msg?.snippet;
        if (!s?.id) throw new Error("Missing snippet.");

        const readmePath = s.readmePath || `snippets/${s.id}/README.md`;
        const entryPath = s.entryPath || `snippets/${s.id}/run.js`;

        const readmeUrl = rawUrl(cfg, readmePath);
        const entryUrl = rawUrl(cfg, entryPath);

        const [readme, code] = await Promise.all([
          fetchTextWithFallback(cfg, readmeUrl),
          fetchTextWithFallback(cfg, entryUrl)
        ]);

        sendResponse({ ok: true, readme, code, entryPath, entryUrl, readmeUrl });
        return;
      }

      if (msg?.type === "RUN_SNIPPET") {
        const tabId = await getActiveTabId();
        const result = await runInMainWorld(tabId, msg?.code);
        sendResponse({ ok: true, result });
        return;
      }

      sendResponse({ ok: false, error: "Unknown message type." });
    } catch (e) {
      sendResponse({ ok: false, error: String(e?.message ?? e) });
    }
  })();

  return true;
});
