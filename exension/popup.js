// popup.js
// Mimics the DevTools snippet-runner UI + behavior, but uses background.js for fetching + running
// (because popup cannot inject MAIN-world code directly).

const CONFIG = {
  uiTitle: "Team Snippets (GitHub)",
  requireConfirmForRisky: true,
  riskyLevels: new Set(["medium", "high"]),
};

// ------------------------------
// Utilities (ported from your snippet runner)
// ------------------------------
const esc = (s) =>
  String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");

function renderMarkdown(mdText) {
  const t = String(mdText ?? "");
  const lines = t.split("\n");
  let html = "";
  let inCode = false;
  let codeLang = "";

  for (const rawLine of lines) {
    const line = rawLine ?? "";

    const codeFence = line.match(/^```(\w+)?\s*$/);
    if (codeFence) {
      inCode = !inCode;
      codeLang = codeFence[1] || "";
      html += inCode
        ? `<pre class="sr-code"><code data-lang="${esc(codeLang)}">`
        : `</code></pre>`;
      continue;
    }

    if (inCode) {
      html += esc(line) + "\n";
      continue;
    }

    const h = line.match(/^(#{1,4})\s+(.*)$/);
    if (h) {
      const lvl = h[1].length;
      html += `<h${lvl} class="sr-h">${esc(h[2])}</h${lvl}>`;
      continue;
    }

    const bullet = line.match(/^\s*[-*]\s+(.*)$/);
    if (bullet) {
      html += `<li class="sr-li">${inlineMd(esc(bullet[1]))}</li>`;
      continue;
    }

    if (!line.trim()) {
      html += `<div class="sr-sp"></div>`;
      continue;
    }

    html += `<p class="sr-p">${inlineMd(esc(line))}</p>`;
  }

  html = html.replace(
    /(?:<li class="sr-li">[\s\S]*?<\/li>)+/g,
    (m) => `<ul class="sr-ul">${m}</ul>`
  );
  return html;

  function inlineMd(safeText) {
    safeText = safeText.replace(
      /`([^`]+)`/g,
      (_, c) => `<code class="sr-inlinecode">${esc(c)}</code>`
    );
    safeText = safeText.replace(/\*\*([^*]+)\*\*/g, `<strong>$1</strong>`);
    safeText = safeText.replace(/\*([^*]+)\*/g, `<em>$1</em>`);
    safeText = safeText.replace(
      /\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g,
      (_, txt, url) =>
        `<a class="sr-a" href="${esc(url)}" target="_blank" rel="noreferrer noopener">${esc(
          txt
        )}</a>`
    );
    return safeText;
  }
}

// ------------------------------
// DOM
// ------------------------------
const $ = (id) => document.getElementById(id);

const titleEl = $("srTitle");
const reloadBtn = $("srReload");
const tokenBtn = $("srToken");

const search = $("srSearch");
const list = $("srList");
const meta = $("srMeta");
const preview = $("srPreview");

const runBtn = $("srRun");
const copyBtn = $("srCopy");
const openBtn = $("srOpen");

const tokenDialog = $("srTokenDialog");
const tokenInput = $("srTokenInput");
const tokenSaveBtn = $("srTokenSave");

titleEl.textContent = CONFIG.uiTitle;

// ------------------------------
// Background messaging
// ------------------------------
function bg(message) {
  return chrome.runtime.sendMessage(message);
}

// ------------------------------
// State
// ------------------------------
let snippets = [];
let selected = null;
let selectedCode = null;
let selectedReadme = null;
let selectedEntryPath = null;
let selectedEntryUrl = null;
let selectedReadmeUrl = null;
let manifestUrl = null;

// ------------------------------
// UI helpers (ported feel)
// ------------------------------
function setMeta(text) {
  meta.textContent = text;
}

function removeSelectionHighlight() {
  for (const el of list.querySelectorAll(".sr-item")) el.classList.remove("sr-selected");
}

function renderList(items) {
  list.innerHTML = "";
  if (!items.length) {
    const empty = document.createElement("div");
    empty.style.padding = "12px";
    empty.style.opacity = "0.8";
    empty.textContent = "No snippets found.";
    list.appendChild(empty);
    return;
  }

  for (const s of items) {
    const badges = document.createElement("div");
    badges.className = "sr-badges";

    if (s.version) {
      const b = document.createElement("span");
      b.className = "sr-badge";
      b.textContent = `v${s.version}`;
      badges.appendChild(b);
    }
    if (s.risk) {
      const b = document.createElement("span");
      b.className = "sr-badge";
      b.textContent = `risk: ${s.risk}`;
      badges.appendChild(b);
    }
    if (Array.isArray(s.tags) && s.tags.length) {
      const b = document.createElement("span");
      b.className = "sr-badge";
      b.textContent = s.tags.join(", ");
      badges.appendChild(b);
    }

    const item = document.createElement("div");
    item.className = "sr-item";
    if (selected?.id === s.id) item.classList.add("sr-selected");

    const t = document.createElement("div");
    t.className = "sr-item-title";
    t.textContent = s.name || s.id;

    const sub = document.createElement("div");
    sub.className = "sr-item-sub";
    sub.textContent = s.description || "";

    item.append(t, sub, badges);
    item.addEventListener("click", () => selectSnippet(s, item));
    list.appendChild(item);
  }
}

// ------------------------------
// Load manifest + wire behavior (extension version)
// ------------------------------
async function loadManifest() {
  runBtn.disabled = true;
  copyBtn.disabled = true;
  openBtn.disabled = true;
  preview.innerHTML = "";
  selected = null;
  selectedCode = null;
  selectedReadme = null;
  selectedEntryPath = null;
  selectedEntryUrl = null;
  selectedReadmeUrl = null;

  setMeta("Loading manifest…");
  const resp = await bg({ type: "LOAD_MANIFEST" });
  if (!resp?.ok) throw new Error(resp?.error ?? "Failed to load manifest.");

  manifestUrl = resp.manifestUrl || null;
  snippets = resp.snippets || [];
  renderList(snippets);

  setMeta(`Loaded ${snippets.length} snippets from manifest.${manifestUrl ? `\n${manifestUrl}` : ""}`);
}

search.addEventListener("input", () => {
  const q = search.value.trim().toLowerCase();
  const filtered = !q
    ? snippets
    : snippets.filter((s) => {
        const hay = `${s.id} ${s.name} ${s.description} ${s.tags?.join(" ")}`.toLowerCase();
        return hay.includes(q);
      });
  renderList(filtered);
});

async function selectSnippet(s, itemEl) {
  removeSelectionHighlight();
  itemEl.classList.add("sr-selected");

  selected = s;
  selectedCode = null;
  selectedReadme = null;
  selectedEntryPath = null;
  selectedEntryUrl = null;
  selectedReadmeUrl = null;

  runBtn.disabled = true;
  copyBtn.disabled = true;
  openBtn.disabled = false;

  setMeta(`Selected: ${s.name || s.id} — loading README and code…`);

  const resp = await bg({ type: "LOAD_SNIPPET_FILES", snippet: s });
  if (!resp?.ok) {
    preview.innerHTML = `<p class="sr-p">Could not load README/code. Check paths, permissions, and ref.</p>`;
    throw new Error(resp?.error ?? "Failed to load snippet files.");
  }

  selectedReadme = resp.readme;
  selectedCode = resp.code;
  selectedEntryPath = resp.entryPath;
  selectedEntryUrl = resp.entryUrl;
  selectedReadmeUrl = resp.readmeUrl;

  preview.innerHTML = renderMarkdown(selectedReadme);

  setMeta(
    `${s.name || s.id} — ${s.description || ""} ` +
      `${s.version ? `| v${s.version} ` : ""}` +
      `${s.risk ? `| risk: ${s.risk} ` : ""}` +
      `${s.owner ? `| owner: ${s.owner}` : ""}` +
      `\n${selectedEntryUrl}`
  );

  runBtn.disabled = false;
  copyBtn.disabled = false;
  openBtn.disabled = false;
}

async function runSelected() {
  if (!selected || !selectedCode) return;

  const risk = String(selected.risk || "").toLowerCase();
  if (CONFIG.requireConfirmForRisky && CONFIG.riskyLevels.has(risk)) {
    const ok = confirm(
      `This snippet is marked risk: ${risk}.\n\n` +
        `Name: ${selected.name || selected.id}\n` +
        `Version: ${selected.version || "?"}\n\n` +
        `Proceed to run it on this page?`
    );
    if (!ok) return;
  }

  setMeta(`▶ Running: ${selected.name || selected.id} …`);

  const resp = await bg({ type: "RUN_SNIPPET", code: selectedCode });
  if (!resp?.ok) {
    setMeta(`❌ ${resp?.error ?? "Run failed."}`);
    return;
  }

  const r = resp.result;
  if (r?.ok) {
    setMeta(
      `✅ Completed: ${selected.name || selected.id}` +
        (r.out !== undefined ? ` (returned: ${String(r.out)})` : "")
    );
    // Snippet itself will log to the page console if it wants
    console.info("[Popup] snippet returned:", r.out);
  } else {
    setMeta(`❌ Error running ${selected.name || selected.id}: ${r?.error ?? "Unknown error"}`);
    console.error("[Popup] execution error:", r);
    alert(`Snippet error: ${r?.error ?? "Unknown error"}\n\n${r?.stack ?? ""}`);
  }
}

async function copySelected() {
  if (!selectedCode) return;
  try {
    await navigator.clipboard.writeText(selectedCode);
    setMeta(`✅ Copied code for ${selected.name || selected.id} to clipboard.`);
  } catch (e) {
    console.warn("Clipboard failed, falling back to prompt()", e);
    prompt("Copy the code below:", selectedCode);
  }
}

async function openInGitHub() {
  if (!selectedEntryPath) return;
  const cfg = await bg({ type: "GET_CONFIG" });
  if (!cfg?.ok) return;
  const c = cfg.config;
  const ghUrl = `https://github.com/${c.owner}/${c.repo}/blob/${c.ref}/${selectedEntryPath}`;
  chrome.tabs.create({ url: ghUrl });
}

// Token dialog
tokenBtn.addEventListener("click", async () => {
  const cfg = await bg({ type: "GET_CONFIG" });
  if (cfg?.ok) {
    // don't pre-fill, but show whether token exists
    setMeta(cfg.hasToken ? "Token is set (will use GitHub API fallback when needed)." : "No token set.");
  }
  tokenInput.value = "";
  tokenDialog.showModal();
});

tokenSaveBtn.addEventListener("click", async (e) => {
  e.preventDefault();
  const token = tokenInput.value.trim();
  const resp = await bg({ type: "SET_TOKEN", token });
  if (resp?.ok) {
    setMeta(resp.hasToken ? "Token saved." : "Token cleared.");
    tokenDialog.close();
  } else {
    alert(resp?.error ?? "Failed to save token.");
  }
});

reloadBtn.addEventListener("click", async () => {
  try {
    await loadManifest();
  } catch (e) {
    setMeta(`❌ Failed to load manifest. ${String(e?.message ?? e)}`);
  }
});

runBtn.addEventListener("click", runSelected);
copyBtn.addEventListener("click", copySelected);
openBtn.addEventListener("click", openInGitHub);

// Init
(async function init() {
  try {
    await loadManifest();
  } catch (e) {
    setMeta(`❌ Failed to load manifest. ${String(e?.message ?? e)}`);
  }
})();
