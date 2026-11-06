"use strict";

const $ = (id) => document.getElementById(id);

let currentState = null;
let allSections = [];
let currentMessages = [];

function show(el) { el.classList.remove("hidden"); }
function hide(el) { el.classList.add("hidden"); }

function showError(text) {
  hide($("result"));
  const s = $("status");
  s.textContent = text;
  show(s);
}

function showResult(html) {
  hide($("status"));
  const r = $("result");
  r.innerHTML = html;
  show(r);
}

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

/* ------------------------------------------------- Message detection */
/*
 * Multi-stage, because tabs.query({currentWindow}) does not reliably return
 * the window containing the opened message in a Thunderbird/Betterbird popup:
 *   1. Active tab in the current window: opened message OR
 *      selection in a mail tab (preview / multi-selection)
 *   2. Fallback: search all windows for opened messages
 */
async function findMessages() {
  const seen = new Set();
  const result = [];
  const add = (m) => {
    if (m && m.id != null && !seen.has(m.id)) {
      seen.add(m.id);
      result.push(m);
    }
  };

  async function probeTab(tab, includeSelection) {
    if (!tab) return;
    try {
      const m = await browser.messageDisplay.getDisplayedMessage(tab.id);
      if (m) { add(m); return; }
    } catch { /* not a message tab */ }
    if (includeSelection) {
      try {
        const sel = await browser.mailTabs.getSelectedMessages(tab.id);
        if (sel?.messages?.length) sel.messages.forEach(add);
      } catch { /* not a mail tab */ }
    }
  }

  try {
    const win = await browser.windows.getCurrent({ populate: true });
    for (const t of win.tabs || []) {
      if (t.active) await probeTab(t, true);
    }
  } catch { /* continue with fallback */ }

  if (!result.length) {
    try {
      const wins = await browser.windows.getAll({ populate: true });
      outer: for (const w of wins) {
        for (const t of w.tabs || []) {
          await probeTab(t, false);
          if (result.length) break outer;
        }
      }
    } catch { /* nothing found */ }
  }
  return result;
}

/* --------------------------------------------------------- Section list */

function fillSections(filter = "") {
  const sel = $("section");
  const f = filter.trim().toLowerCase();
  const filtered = f
    ? allSections.filter((s) => `${s.name} ${s.notebook}`.toLowerCase().includes(f))
    : allSections;

  sel.innerHTML = "";
  $("section-count").textContent = allSections.length
    ? `${filtered.length} of ${allSections.length} sections`
    : "";

  if (!filtered.length) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = allSections.length
      ? "(no matches – adjust your search)"
      : "(no sections – please refresh)";
    sel.appendChild(opt);
    return;
  }

  const byNotebook = new Map();
  for (const s of filtered) {
    if (!byNotebook.has(s.notebook)) byNotebook.set(s.notebook, []);
    byNotebook.get(s.notebook).push(s);
  }
  for (const [notebook, list] of [...byNotebook.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
    const group = document.createElement("optgroup");
    group.label = notebook;
    for (const s of list.sort((a, b) => a.name.localeCompare(b.name))) {
      const opt = document.createElement("option");
      opt.value = s.id;
      opt.textContent = s.name;
      group.appendChild(opt);
    }
    sel.appendChild(group);
  }

  if (currentState?.lastSectionId && [...sel.options].some((o) => o.value === currentState.lastSectionId)) {
    sel.value = currentState.lastSectionId;
  } else if (filtered.length === 1) {
    sel.value = filtered[0].id;
  }
}

/* ----------------------------------------------------------------- Views */

function renderMessageInfo() {
  const info = $("message-info");
  const btn = $("btn-clip");
  if (!currentMessages.length) {
    info.textContent = "⚠️ No message found. Open an email (or select it in the list), then open this popup again.";
    info.className = "small warn";
    btn.disabled = true;
  } else if (currentMessages.length === 1) {
    info.innerHTML = `Message: <b>${escapeHtml(currentMessages[0].subject || "(no subject)")}</b>`;
    info.className = "small msg";
    btn.disabled = false;
    btn.textContent = "Save to OneNote";
  } else {
    info.innerHTML = `<b>${currentMessages.length} selected messages</b> – will be saved one after another in the chosen section.`;
    info.className = "small msg";
    btn.disabled = false;
    btn.textContent = `Save ${currentMessages.length} emails to OneNote`;
  }
}

function render() {
  hide($("view-setup")); hide($("view-signin")); hide($("view-clip"));
  hide($("status")); hide($("result"));

  if (!currentState.clientIdSet) { show($("view-setup")); return; }
  if (!currentState.signedIn) { show($("view-signin")); return; }

  $("account").textContent = `Signed in as ${currentState.accountName || "Microsoft account"}`;
  allSections = currentState.sections || [];
  fillSections($("section-search").value);
  renderMessageInfo();
  show($("view-clip"));
}

async function init() {
  try {
    currentState = await browser.runtime.sendMessage({ type: "getState" });
  } catch (e) {
    showError(`Add-on background not reachable:\n${e.message}`);
    return;
  }
  currentMessages = await findMessages();
  render();
}

/* --------------------------------------------------------------- Actions */

$("btn-open-options").addEventListener("click", () => {
  browser.runtime.openOptionsPage();
  window.close();
});

$("btn-signin").addEventListener("click", async (ev) => {
  ev.target.disabled = true;
  try {
    currentState = await browser.runtime.sendMessage({ type: "signIn" });
    if (currentState?.error) throw new Error(currentState.error);
    render();
  } catch (e) {
    showError(`Sign-in failed:\n${e.message}`);
  } finally {
    ev.target.disabled = false;
  }
});

$("section-search").addEventListener("input", (ev) => fillSections(ev.target.value));

$("btn-clip").addEventListener("click", async (ev) => {
  const sectionId = $("section").value;
  if (!sectionId) {
    showError("Please select a section first (clear the search or refresh sections).");
    return;
  }
  if (!currentMessages.length) {
    showError("No message found. Please open/select an email and open the popup again.");
    return;
  }

  const btn = ev.target;
  const includeAttachments = $("include-attachments").checked;
  const errors = [];
  const warningsAll = [];
  let okCount = 0;
  let firstUrl = null;
  btn.disabled = true;

  try {
    for (let i = 0; i < currentMessages.length; i++) {
      btn.textContent = currentMessages.length > 1 ? `Saving ${i + 1}/${currentMessages.length} …` : "Saving …";
      try {
        const res = await browser.runtime.sendMessage({
          type: "clip",
          messageId: currentMessages[i].id,
          sectionId,
          includeAttachments,
        });
        if (res?.error) throw new Error(res.error);
        okCount++;
        firstUrl = firstUrl || res.pageUrl;
        if (res.warnings?.length) warningsAll.push(...res.warnings);
      } catch (e) {
        errors.push(`"${currentMessages[i].subject || "(no subject)"}": ${e.message}`);
      }
    }
  } finally {
    btn.disabled = false;
    btn.textContent = "Save to OneNote";
  }

  if (errors.length) {
    showError(`${okCount} of ${currentMessages.length} messages saved.\nErrors:\n${errors.join("\n")}`);
  } else {
    let html = currentMessages.length > 1
      ? `✅ ${okCount} emails saved to OneNote.`
      : "✅ Saved to OneNote.";
    if (firstUrl) html += ` <a href="${escapeHtml(firstUrl)}" target="_blank" rel="noopener">Open page</a>`;
    if (warningsAll.length) html += `\n⚠️ ${escapeHtml(warningsAll.join("\n"))}`;
    showResult(html);
  }
});

$("link-refresh").addEventListener("click", async (ev) => {
  ev.preventDefault();
  ev.target.textContent = "refreshing …";
  try {
    currentState = await browser.runtime.sendMessage({ type: "refreshSections" });
    if (currentState?.error) throw new Error(currentState.error);
    allSections = currentState.sections || [];
    fillSections($("section-search").value);
  } catch (e) {
    showError(`Refresh failed:\n${e.message}`);
  } finally {
    ev.target.textContent = "Refresh sections";
  }
});

$("link-signout").addEventListener("click", async (ev) => {
  ev.preventDefault();
  currentState = await browser.runtime.sendMessage({ type: "signOut" });
  render();
});

$("link-manual-auth").addEventListener("click", (ev) => {
  ev.preventDefault();
  browser.tabs.create({ url: browser.runtime.getURL("auth.html") });
  window.close();
});

$("link-options").addEventListener("click", (ev) => {
  ev.preventDefault();
  browser.runtime.openOptionsPage();
});

init();
