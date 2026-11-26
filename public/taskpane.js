/* Multi-Cursor for Word (Office.js)
 *
 * Word's API has no native multi-selection, so this emulates it:
 * "cursors" are tracked Range objects (collapsed = caret, non-empty = a selected match).
 * Word keeps tracked ranges in sync as the document is edited.
 */

const HIGHLIGHT = "#FFE699";
let ctx = null;          // one shared RequestContext so tracked ranges stay valid
let cursors = [];        // { range: Word.Range, orig: string|null, isMatch: boolean }

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus("This add-in only runs in Word.");
    return;
  }
  ctx = new Word.RequestContext();
  const on = (id, fn) => document.getElementById(id).addEventListener("click", () => run(fn));
  on("btnAdd", addCursorAtSelection);
  on("btnClear", clearAll);
  on("btnLineStart", () => addLineEdges("start"));
  on("btnLineEnd", () => addLineEdges("end"));
  on("btnMatches", selectMatches);
  on("btnType", () => typeText("replace"));
  on("btnPrefix", () => typeText("before"));
  on("btnSuffix", () => typeText("after"));
  on("btnCase", applyCase);
  on("btnBold", () => applyFont("bold"));
  on("btnItalic", () => applyFont("italic"));
  on("btnUnderline", () => applyFont("underline"));
  on("btnDelete", deleteText);
  document.getElementById("find").addEventListener("keydown", (e) => { if (e.key === "Enter") run(selectMatches); });
  document.getElementById("text").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) run(() => typeText("replace"));
  });
  render();
  setStatus("Ready. Put the caret somewhere and press “Add cursor at selection”, or search for text to select all matches.");
});

/* ---------- helpers ---------- */

async function run(fn) {
  try { await fn(); }
  catch (e) { console.error(e); setStatus("Error: " + (e.message || e)); }
}

function setStatus(msg) { document.getElementById("status").textContent = msg; }

function track(range) { ctx.trackedObjects.add(range); return range; }
function untrack(range) { try { ctx.trackedObjects.remove(range); } catch (_) {} }

async function render() {
  const list = document.getElementById("list");
  list.innerHTML = "";
  if (!cursors.length) { list.innerHTML = '<div style="color:var(--muted)">No cursors yet.</div>'; return; }
  const shown = cursors.slice(0, 60);
  const paras = shown.map((c) => { const p = c.range.paragraphs.getFirst(); p.load("text"); return p; });
  try { await ctx.sync(); } catch (_) {}
  shown.forEach((c, i) => {
    const d = document.createElement("div");
    const txt = (paras[i].isNullObject === false && paras[i].text) ? paras[i].text.trim() : "";
    const b = document.createElement("b");
    b.textContent = String(i + 1) + (c.isMatch ? "▮" : "|");
    d.appendChild(b);
    d.appendChild(document.createTextNode(txt.slice(0, 70) || "(empty line)"));
    list.appendChild(d);
  });
  if (cursors.length > shown.length) {
    const d = document.createElement("div");
    d.textContent = `… and ${cursors.length - shown.length} more`;
    list.appendChild(d);
  }
}

async function loadAll() {
  cursors.forEach((c) => c.range.load("text,isEmpty"));
  await ctx.sync();
}

function requireCursors() {
  if (!cursors.length) { setStatus("Add some cursors first."); return false; }
  return true;
}

/* ---------- placing cursors ---------- */

async function addCursorAtSelection() {
  const sel = ctx.document.getSelection();
  const r = track(sel.getRange("Whole"));
  r.load("isEmpty");
  await ctx.sync();
  cursors.push({ range: r, orig: null, isMatch: !r.isEmpty });
  await render();
  setStatus(`${cursors.length} cursor${cursors.length === 1 ? "" : "s"}. Move the caret and add another.`);
}

async function addLineEdges(which) {
  const sel = ctx.document.getSelection();
  const paras = sel.paragraphs;
  paras.load("items");
  await ctx.sync();
  const added = paras.items.map((p) => {
    const content = p.getRange("Content");
    return track(content.getRange(which === "start" ? "Start" : "End"));
  });
  await ctx.sync();
  added.forEach((r) => cursors.push({ range: r, orig: null, isMatch: false }));
  await render();
  setStatus(`Added ${added.length} cursor${added.length === 1 ? "" : "s"} at line ${which}s.`);
}

async function selectMatches() {
  const q = document.getElementById("find").value;
  if (!q) { setStatus("Type the text to find first."); return; }
  const opts = {
    matchCase: document.getElementById("matchCase").checked,
    matchWholeWord: document.getElementById("wholeWord").checked,
    matchWildcards: document.getElementById("wildcards").checked,
  };
  const scope = document.getElementById("inSel").checked ? ctx.document.getSelection() : ctx.document.body;
  const found = scope.search(q, opts);
  found.load("items/font/highlightColor");
  await ctx.sync();

  const items = found.items;
  items.forEach((r) => {
    track(r);
    cursors.push({ range: r, orig: r.font.highlightColor, isMatch: true });
    r.font.highlightColor = HIGHLIGHT;
  });
  await ctx.sync();
  await render();
  setStatus(items.length ? `Selected ${items.length} match${items.length === 1 ? "" : "es"} (highlighted).` : "No matches found.");
}

async function clearAll(silent) {
  // restore original highlight on matches
  cursors.forEach((c) => { if (c.isMatch) c.range.font.highlightColor = c.orig; });
  try { await ctx.sync(); } catch (_) {}
  cursors.forEach((c) => untrack(c.range));
  try { await ctx.sync(); } catch (_) {}
  cursors = [];
  await render();
  if (silent !== true) setStatus("Cleared all cursors.");
}

/* ---------- editing ---------- */

function expand(template, index, start) {
  return template.replace(/\{n\}/g, String(start + index));
}

async function typeText(mode) {
  if (!requireCursors()) return;
  const template = document.getElementById("text").value;
  if (!template) { setStatus("Enter the text to insert."); return; }
  const start = parseInt(document.getElementById("nStart").value, 10) || 0;

  const next = [];
  cursors.forEach((c, i) => {
    const txt = expand(template, i, start);
    let inserted;
    if (mode === "replace") {
      inserted = c.range.insertText(txt, "Replace");
    } else if (mode === "before") {
      inserted = c.range.insertText(txt, "Start");
    } else {
      inserted = c.range.insertText(txt, "End");
    }
    if (c.isMatch) inserted.font.highlightColor = c.orig;   // drop our highlight on new text
    // new caret sits after the inserted text so repeated typing keeps appending
    const caret = inserted.getRange("End");
    next.push({ old: c, caret });
  });
  await ctx.sync();

  next.forEach(({ old, caret }) => { track(caret); });
  await ctx.sync();
  next.forEach(({ old }) => untrack(old.range));
  cursors = next.map(({ caret }) => ({ range: caret, orig: null, isMatch: false }));
  await ctx.sync();
  await render();
  setStatus(`Inserted at ${cursors.length} location${cursors.length === 1 ? "" : "s"}. Cursors now sit after the new text.`);
}

async function applyCase() {
  if (!requireCursors()) return;
  await loadAll();
  const mode = document.getElementById("caseMode").value;
  let n = 0;
  cursors.forEach((c) => {
    if (c.range.isEmpty) return;
    const t = c.range.text;
    let out = t;
    if (mode === "upper") out = t.toUpperCase();
    else if (mode === "lower") out = t.toLowerCase();
    else out = t.toLowerCase().replace(/(^|[\s\-"'(])(\p{L})/gu, (_, a, b) => a + b.toUpperCase());
    if (out !== t) { c.range.insertText(out, "Replace"); n++; }
  });
  await ctx.sync();
  await render();
  setStatus(`Changed case in ${n} selection${n === 1 ? "" : "s"}. (Mixed formatting inside a match is flattened.)`);
}

async function applyFont(prop) {
  if (!requireCursors()) return;
  await loadAll();
  const targets = cursors.filter((c) => !c.range.isEmpty);
  if (!targets.length) { setStatus("These are empty cursors. Formatting needs selected text (use Select all matches)."); return; }
  targets.forEach((c) => {
    if (prop === "underline") c.range.font.underline = "Single";
    else c.range.font[prop] = true;
  });
  await ctx.sync();
  setStatus(`Applied ${prop} to ${targets.length} selection${targets.length === 1 ? "" : "s"}.`);
}

async function deleteText() {
  if (!requireCursors()) return;
  await loadAll();
  const targets = cursors.filter((c) => !c.range.isEmpty);
  if (!targets.length) { setStatus("Nothing to delete: all cursors are empty."); return; }
  targets.forEach((c) => c.range.insertText("", "Replace"));
  await ctx.sync();
  // the matches are gone, so keep them as plain cursors
  targets.forEach((c) => { c.isMatch = false; c.orig = null; });
  await render();
  setStatus(`Deleted text at ${targets.length} location${targets.length === 1 ? "" : "s"}.`);
}
