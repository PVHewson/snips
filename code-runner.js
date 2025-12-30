/**
* D365 DevTools Snippet Runner (GitHub-backed)
* - Loads a manifest from a GitHub repo (raw content)
* - Shows a UI: list of snippets + README preview
* - Runs selected snippet on demand
*
* ✅ Paste this into a Chrome DevTools "Snippet" and run it.
*
* Repo expectation (you can change paths in CONFIG):
* /snippets/manifest.json
* /snippets/<snippet-id>/run.js
* /snippets/<snippet-id>/README.md
*
* Manifest example is included at the bottom of this file.
*/
(async function SnippetRunner() {
const CONFIG = {
owner: "PVHewson",
repo: "snips",
ref: "main", // branch or tag (recommended: tag like "v1.2.3")
manifestPath: "snippets/manifest.json",
// For private repos: set to "prompt" to ask user each time, or paste a token (not recommended).
githubToken: "prompt", // "prompt" | "" | "<token>"
uiTitle: "Team Snippets (GitHub)",
// Safety defaults:
requireConfirmForRisky: true,
riskyLevels: new Set(["medium", "high"]),
};

// ------------------------------
// Utilities
// ------------------------------

const esc = (s) =>
String(s ?? "")
.replaceAll("&", "&amp;")
.replaceAll("<", "&lt;")
.replaceAll(">", "&gt;")
.replaceAll('"', "&quot;")
.replaceAll("'", "&#039;");

// Very small markdown-ish renderer (safe-ish, no HTML passthrough):
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

// Headings
const h = line.match(/^(#{1,4})\s+(.*)$/);
if (h) {
const lvl = h[1].length;
html += `<h${lvl} class="sr-h">${esc(h[2])}</h${lvl}>`;
continue;
}

// Bullets
const bullet = line.match(/^\s*[-*]\s+(.*)$/);
if (bullet) {
// naive list handling: wrap each bullet in <li>, and we’ll auto-wrap with <ul> blocks later
html += `<li class="sr-li">${inlineMd(esc(bullet[1]))}</li>`;
continue;
}

// Blank line
if (!line.trim()) {
html += `<div class="sr-sp"></div>`;
continue;
}

// Paragraph
html += `<p class="sr-p">${inlineMd(esc(line))}</p>`;
}

// Wrap consecutive <li> into <ul>
html = html.replace(/(?:<li class="sr-li">[\s\S]*?<\/li>)+/g, (m) => `<ul class="sr-ul">${m}</ul>`);
return html;

function inlineMd(safeText) {
// inline code: `code`
safeText = safeText.replace(/`([^`]+)`/g, (_, c) => `<code class="sr-inlinecode">${esc(c)}</code>`);
// bold: **text**
safeText = safeText.replace(/\*\*([^*]+)\*\*/g, `<strong>$1</strong>`);
// italic: *text*
safeText = safeText.replace(/\*([^*]+)\*/g, `<em>$1</em>`);
// links: [text](url) - url must be http(s)
safeText = safeText.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, (_, txt, url) => {
return `<a class="sr-a" href="${esc(url)}" target="_blank" rel="noreferrer noopener">${esc(txt)}</a>`;
});
return safeText;
}
}

function getXrmSafely() {
// D365: usually window.Xrm exists; sometimes inside frames.
// Keep it simple: prefer current window.
return window.Xrm || null;
}

function createEl(tag, attrs = {}, children = []) {
const el = document.createElement(tag);
for (const [k, v] of Object.entries(attrs)) {
if (k === "class") el.className = v;
else if (k === "style") Object.assign(el.style, v);
else if (k.startsWith("on") && typeof v === "function") el.addEventListener(k.slice(2), v);
else el.setAttribute(k, String(v));
}
for (const c of [].concat(children)) {
if (c == null) continue;
el.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
}
return el;
}

function getGithubHeaders() {
const headers = { "Accept": "application/vnd.github+json" };
const token = (CONFIG.githubToken || "").trim();
if (token) headers["Authorization"] = `Bearer ${token}`;
return headers;
}

async function ensureTokenIfNeeded(url) {
// Raw GitHub URLs can be fetched without token for public repos.
// For private repos, a token is needed and raw fetch may still fail depending on org settings.
if (CONFIG.githubToken === "prompt") {
// only prompt once, lazily, when a fetch fails
return;
}
if (typeof CONFIG.githubToken !== "string") CONFIG.githubToken = "";
}

function rawUrl(path) {
const p = String(path).replace(/^\/+/, "");
return `https://raw.githubusercontent.com/${CONFIG.owner}/${CONFIG.repo}/${CONFIG.ref}/${p}`;
}

async function fetchText(url, { allowRetryWithTokenPrompt = true } = {}) {
await ensureTokenIfNeeded(url);

// 1) Attempt raw fetch (public-friendly)
let res = await fetch(url, {
method: "GET",
credentials: "omit",
cache: "no-store",
});

// 2) If forbidden/404 and token prompting is enabled, prompt and retry via GitHub API (private-friendly)
if (!res.ok && allowRetryWithTokenPrompt && CONFIG.githubToken === "prompt") {
const status = res.status;

// Only prompt if it looks like access is the issue
if ([401, 403, 404].includes(status)) {
const token = prompt(
"This repo/path may be private or blocked. Paste a GitHub Personal Access Token (classic or fine-grained) with read access.\n\n(Leave blank to cancel.)"
);
if (token && token.trim()) {
CONFIG.githubToken = token.trim();
return fetchViaGithubApi(url);
}
}
}

if (!res.ok) {
const body = await safeRead(res);
throw new Error(`Fetch failed (${res.status}) for ${url}\n${body ? "Details: " + body : ""}`);
}
return res.text();
}

async function fetchViaGithubApi(rawUrlThatFailed) {
// Convert raw URL into GitHub "contents" API call:
// https://api.github.com/repos/:owner/:repo/contents/:path?ref=:ref
const marker = `https://raw.githubusercontent.com/${CONFIG.owner}/${CONFIG.repo}/${CONFIG.ref}/`;
const path = rawUrlThatFailed.startsWith(marker) ? rawUrlThatFailed.slice(marker.length) : null;
if (!path) throw new Error("Could not map raw URL to repo path for API fallback.");

const apiUrl = `https://api.github.com/repos/${CONFIG.owner}/${CONFIG.repo}/contents/${path}?ref=${encodeURIComponent(
CONFIG.ref
)}`;

const res = await fetch(apiUrl, {
method: "GET",
headers: getGithubHeaders(),
cache: "no-store",
});

if (!res.ok) {
const body = await safeRead(res);
throw new Error(`GitHub API fetch failed (${res.status}) for ${apiUrl}\n${body ? "Details: " + body : ""}`);
}

const json = await res.json();
if (!json || json.type !== "file" || !json.content) {
throw new Error(`GitHub API returned non-file content for ${path}`);
}

// content is base64 (may contain newlines)
const b64 = String(json.content).replace(/\n/g, "");
const txt = atob(b64);
// Handle UTF-8 properly:
// (atob returns binary string; decode to utf-8)
try {
const bytes = Uint8Array.from(txt, (c) => c.charCodeAt(0));
return new TextDecoder("utf-8").decode(bytes);
} catch {
return txt;
}
}

async function safeRead(res) {
try {
return await res.text();
} catch {
return "";
}
}

function removeExistingUi() {
const old = document.getElementById("sr-root");
if (old) old.remove();
}

// ------------------------------
// UI
// ------------------------------

removeExistingUi();

const root = createEl("div", { id: "sr-root", class: "sr-root" });
const backdrop = createEl("div", { class: "sr-backdrop" });
const panel = createEl("div", { class: "sr-panel", role: "dialog", "aria-modal": "true" });

const header = createEl("div", { class: "sr-header" }, [
createEl("div", { class: "sr-title" }, [CONFIG.uiTitle]),
createEl("div", { class: "sr-actions" }, [
createEl("button", { class: "sr-btn sr-btn-ghost", onclick: () => cleanup() }, ["Close"]),
]),
]);

const body = createEl("div", { class: "sr-body" });

const left = createEl("div", { class: "sr-left" });
const right = createEl("div", { class: "sr-right" });

const search = createEl("input", {
class: "sr-search",
placeholder: "Search snippets…",
autocomplete: "off",
});

const list = createEl("div", { class: "sr-list" });

const meta = createEl("div", { class: "sr-meta" }, ["Select a snippet to preview its README."]);
const preview = createEl("div", { class: "sr-preview" });
const footer = createEl("div", { class: "sr-footer" });

const runBtn = createEl("button", { class: "sr-btn sr-btn-primary", disabled: "true" }, ["Run"]);
const copyBtn = createEl("button", { class: "sr-btn", disabled: "true" }, ["Copy code"]);
const openBtn = createEl("button", { class: "sr-btn", disabled: "true" }, ["Open in GitHub"]);

footer.append(runBtn, copyBtn, openBtn);
left.append(search, list);
right.append(meta, preview, footer);

body.append(left, right);
panel.append(header, body);
root.append(backdrop, panel);
document.body.appendChild(root);

// Styles (scoped-ish)
const style = createEl("style", {}, [`
.sr-root { position: fixed; inset: 0; z-index: 2147483647; font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; }
.sr-backdrop { position: absolute; inset: 0; background: rgba(0,0,0,0.45); }
.sr-panel { position: absolute; top: 5vh; left: 50%; transform: translateX(-50%); width: min(1100px, 92vw); height: 90vh; background: #111; color: #eee; border: 1px solid rgba(255,255,255,0.12); border-radius: 12px; overflow: hidden; box-shadow: 0 20px 80px rgba(0,0,0,0.55); display:flex; flex-direction:column; }
.sr-header { display:flex; justify-content:space-between; align-items:center; padding: 12px 14px; background: #161616; border-bottom: 1px solid rgba(255,255,255,0.08); }
.sr-title { font-weight: 700; letter-spacing: 0.2px; }
.sr-actions { display:flex; gap: 8px; }
.sr-body { display:flex; flex:1; min-height: 0; }
.sr-left { width: 360px; max-width: 45%; border-right: 1px solid rgba(255,255,255,0.08); display:flex; flex-direction:column; min-height:0; }
.sr-right { flex:1; display:flex; flex-direction:column; min-height:0; }
.sr-search { margin: 12px; padding: 10px 10px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.14); background:#0d0d0d; color:#eee; outline:none; }
.sr-list { padding: 6px; overflow:auto; min-height:0; }
.sr-item { padding: 10px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.08); background:#141414; margin: 6px; cursor:pointer; }
.sr-item:hover { border-color: rgba(255,255,255,0.20); }
.sr-item.sr-selected { border-color: rgba(120,170,255,0.65); background: #101825; }
.sr-item-title { font-weight: 700; margin-bottom: 4px; }
.sr-item-sub { opacity: 0.85; font-size: 12px; line-height: 1.35; }
.sr-badges { display:flex; gap:6px; margin-top:6px; flex-wrap:wrap; }
.sr-badge { font-size: 11px; padding: 2px 8px; border-radius: 999px; border: 1px solid rgba(255,255,255,0.14); opacity: 0.9; }
.sr-meta { padding: 12px 14px; border-bottom: 1px solid rgba(255,255,255,0.08); background:#101010; font-size: 12px; opacity: 0.9; }
.sr-preview { padding: 14px; overflow:auto; min-height:0; }
.sr-footer { padding: 12px 14px; display:flex; gap: 10px; border-top: 1px solid rgba(255,255,255,0.08); background:#161616; }
.sr-btn { padding: 9px 12px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.18); background: #1e1e1e; color:#eee; cursor:pointer; }
.sr-btn:disabled { opacity: 0.45; cursor:not-allowed; }
.sr-btn-primary { background:#2a63ff; border-color: rgba(42,99,255,0.75); }
.sr-btn-ghost { background: transparent; }
/* Markdown-ish */
.sr-h { margin: 10px 0 8px; }
.sr-p { margin: 8px 0; line-height: 1.45; }
.sr-ul { margin: 8px 0 8px 18px; }
.sr-li { margin: 4px 0; }
.sr-code { background:#0b0b0b; border: 1px solid rgba(255,255,255,0.10); padding: 10px; border-radius: 10px; overflow:auto; }
.sr-inlinecode { background:#0b0b0b; border: 1px solid rgba(255,255,255,0.10); padding: 1px 6px; border-radius: 8px; }
.sr-a { color: #8db2ff; }
`]);
document.head.appendChild(style);

function cleanup() {
root.remove();
style.remove();
window.removeEventListener("keydown", onKey);
}

function onKey(e) {
if (e.key === "Escape") cleanup();
}
window.addEventListener("keydown", onKey);

// ------------------------------
// Load manifest + wire up behavior
// ------------------------------

let manifest;
let snippets = [];
let selected = null;
let selectedCode = null;
let selectedReadme = null;

const manifestUrl = rawUrl(CONFIG.manifestPath);

setMeta(`Loading manifest… (${manifestUrl})`);

try {
const manifestText = await fetchText(manifestUrl);
manifest = JSON.parse(manifestText);
if (!manifest || !Array.isArray(manifest.snippets)) {
throw new Error("manifest.json must contain { snippets: [...] }");
}
snippets = manifest.snippets.slice().sort((a, b) => (a.name || "").localeCompare(b.name || ""));
} catch (err) {
setMeta(`❌ Failed to load manifest. ${String(err?.message || err)}`);
console.error(err);
return;
}

renderList(snippets);

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

function setMeta(text) {
meta.textContent = text;
}

function renderList(items) {
list.innerHTML = "";
if (!items.length) {
list.appendChild(createEl("div", { style: { padding: "12px", opacity: 0.8 } }, ["No snippets found."]));
return;
}
for (const s of items) {
const badges = createEl("div", { class: "sr-badges" }, [
s.version ? createEl("span", { class: "sr-badge" }, [`v${s.version}`]) : null,
s.risk ? createEl("span", { class: "sr-badge" }, [`risk: ${s.risk}`]) : null,
s.tags?.length ? createEl("span", { class: "sr-badge" }, [s.tags.join(", ")]) : null,
].filter(Boolean));

const item = createEl("div", {
class: "sr-item",
onclick: () => selectSnippet(s, item),
}, [
createEl("div", { class: "sr-item-title" }, [s.name || s.id]),
createEl("div", { class: "sr-item-sub" }, [s.description || ""]),
badges,
]);

// keep highlight if currently selected
if (selected?.id === s.id) item.classList.add("sr-selected");
list.appendChild(item);
}
}

async function selectSnippet(s, itemEl) {
// update selection UI
for (const el of list.querySelectorAll(".sr-item")) el.classList.remove("sr-selected");
itemEl.classList.add("sr-selected");

selected = s;
selectedCode = null;
selectedReadme = null;

runBtn.disabled = true;
copyBtn.disabled = true;
openBtn.disabled = false;

const readmePath = s.readmePath || `snippets/${s.id}/README.md`;
const entryPath = s.entryPath || `snippets/${s.id}/run.js`;

const readmeUrl = rawUrl(readmePath);
const entryUrl = rawUrl(entryPath);

setMeta(`Selected: ${s.name || s.id} — loading README and code…`);

try {
const [md, code] = await Promise.all([
fetchText(readmeUrl, { allowRetryWithTokenPrompt: true }),
fetchText(entryUrl, { allowRetryWithTokenPrompt: true }),
]);

selectedReadme = md;
selectedCode = code;

// render README
preview.innerHTML = renderMarkdown(md);

// update meta
setMeta(
`${s.name || s.id} — ${s.description || ""} ` +
`${s.version ? `| v${s.version} ` : ""}` +
`${s.risk ? `| risk: ${s.risk} ` : ""}` +
`${s.owner ? `| owner: ${s.owner}` : ""}`
);

runBtn.disabled = false;
copyBtn.disabled = false;

openBtn.onclick = () => {
const ghUrl = `https://github.com/${CONFIG.owner}/${CONFIG.repo}/blob/${CONFIG.ref}/${entryPath}`;
window.open(ghUrl, "_blank", "noreferrer");
};

copyBtn.onclick = async () => {
try {
await navigator.clipboard.writeText(code);
setMeta(`✅ Copied code for ${s.name || s.id} to clipboard.`);
} catch (e) {
console.warn("Clipboard failed, falling back to prompt()", e);
prompt("Copy the code below:", code);
}
};

runBtn.onclick = () => runSnippet(s, code);

} catch (err) {
setMeta(`❌ Failed to load selected snippet files. ${String(err?.message || err)}`);
preview.innerHTML = `<p class="sr-p">Could not load README/code. Check paths, permissions, and ref.</p>`;
console.error(err);
}
}

async function runSnippet(s, code) {
try {
const risk = String(s.risk || "").toLowerCase();
if (CONFIG.requireConfirmForRisky && CONFIG.riskyLevels.has(risk)) {
const ok = confirm(
`This snippet is marked risk: ${risk}.\n\n` +
`Name: ${s.name || s.id}\n` +
`Version: ${s.version || "?"}\n\n` +
`Proceed to run it on this page?`
);
if (!ok) return;
}

// Provide a controlled context object. Snippet can use `ctx.Xrm`, `ctx.window`, etc.
const ctx = {
Xrm: getXrmSafely(),
window,
document,
location,
console,
// helpers for snippet authors
sleep: (ms) => new Promise((r) => setTimeout(r, ms)),
notify: (msg) => console.info("[SnippetRunner]", msg),
};

// Wrap the snippet in an async function so `await` works inside.
// IMPORTANT: this is essentially eval. Only run code you trust (repo + reviews + tags).
const wrapped = `(async (ctx) => {\n"use strict";\n${code}\n})`;
const fn = (0, eval)(wrapped); // avoid local scope capture

setMeta(`▶ Running: ${s.name || s.id} …`);
const result = await fn(ctx);

setMeta(`✅ Completed: ${s.name || s.id}${result !== undefined ? ` (returned: ${String(result)})` : ""}`);
console.info("[SnippetRunner] result:", result);
} catch (err) {
setMeta(`❌ Error running ${s.name || s.id}: ${String(err?.message || err)}`);
console.error("[SnippetRunner] execution error:", err);
alert(`Snippet error: ${String(err?.message || err)}`);
}
}

// initial meta
setMeta(`Loaded ${snippets.length} snippets from manifest. Select one to preview.`);

// ------------------------------
// Manifest example (for your repo)
// ------------------------------
console.info("[SnippetRunner] Manifest example:\n", {
snippets: [
{
id: "fix.case.timeline.refresh",
name: "Fix: Case form refresh timeline",
description: "Refreshes the timeline control when it fails to update automatically.",
version: "1.2.0",
risk: "low",
owner: "team-d365",
tags: ["case", "form", "ui"],
readmePath: "snippets/fix.case.timeline.refresh/README.md",
entryPath: "snippets/fix.case.timeline.refresh/run.js",
},
],
});
})();
