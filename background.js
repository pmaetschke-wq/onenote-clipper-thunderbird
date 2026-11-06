"use strict";

/*
 * OneNote Clipper (local) – background script
 *
 * All network traffic goes exclusively to:
 *   - https://login.microsoftonline.com (OAuth2, directly with Microsoft)
 *   - https://graph.microsoft.com       (Microsoft Graph API)
 * There is no third-party server. Tokens are stored in browser.storage.local
 * inside the user's own Thunderbird/Betterbird profile.
 */

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const LOGIN_BASE = "https://login.microsoftonline.com";
const SCOPES = "Notes.ReadWrite User.Read offline_access";
const MAX_ATTACHMENT_TOTAL = 20 * 1024 * 1024; // 20 MB total budget for binary parts

/*
 * AADSTS90023 workaround:
 * fetch() from an extension page automatically sends an Origin header
 * (moz-extension://…). Microsoft allows "cross-origin token redemption" for
 * native clients only with an allowlist entry. Without an Origin header the
 * request is not considered cross-origin and token redemption succeeds.
 * This only ever touches requests to login.microsoftonline.com.
 */
try {
  browser.webRequest.onBeforeSendHeaders.addListener(
    (details) => ({
      requestHeaders: details.requestHeaders.filter(
        (h) => h.name.toLowerCase() !== "origin"
      ),
    }),
    { urls: ["https://login.microsoftonline.com/*"] },
    ["blocking", "requestHeaders"]
  );
} catch (e) {
  console.warn("OneNote Clipper: webRequest unavailable, trying declarativeNetRequest", e);
}
if (browser.declarativeNetRequest?.updateSessionRules) {
  browser.declarativeNetRequest
    .updateSessionRules({
      removeRuleIds: [1],
      addRules: [
        {
          id: 1,
          priority: 1,
          action: {
            type: "modifyHeaders",
            requestHeaders: [{ header: "Origin", operation: "remove" }],
          },
          condition: {
            urlFilter: "|https://login.microsoftonline.com/",
            resourceTypes: ["xmlhttprequest", "other"],
          },
        },
      ],
    })
    .catch(() => {});
}

/* ---------------------------------------------------------------- Helpers */

function randomBytesUrlSafe(bytes) {
  const a = new Uint8Array(bytes);
  crypto.getRandomValues(a);
  let s = "";
  for (const b of a) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

async function pkceChallenge(verifier) {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier));
  const bytes = new Uint8Array(digest);
  let s = "";
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function escapeXml(s) {
  return String(s ?? "").replace(/[<>&'"]/g, (c) => ({
    "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&apos;", '"': "&quot;",
  }[c]));
}

function getState() {
  return browser.storage.local.get({
    clientId: "",
    tenant: "common",
    refreshToken: null,
    accessToken: null,
    accessTokenExp: 0,
    accountName: null,
    sections: null,
    sectionsFetched: 0,
    lastSectionId: null,
  });
}

/* ------------------------------------------------------------------ OAuth */

async function tokenRequest(params) {
  const { tenant } = await getState();
  const res = await fetch(`${LOGIN_BASE}/${encodeURIComponent(tenant)}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(params),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.error_description || data.error || `Token request failed (HTTP ${res.status})`);
  }
  return data;
}

async function saveTokens(tok) {
  const upd = {
    accessToken: tok.access_token,
    accessTokenExp: Date.now() + (tok.expires_in || 3600) * 1000,
  };
  if (tok.refresh_token) upd.refreshToken = tok.refresh_token;
  await browser.storage.local.set(upd);
}

async function createAuthRequest() {
  const { clientId, tenant } = await getState();
  if (!clientId) {
    throw new Error("No client ID configured. Please enter it in the add-on settings first (see setup guide).");
  }

  const redirectUri = browser.identity.getRedirectURL();
  const verifier = randomBytesUrlSafe(32);
  const challenge = await pkceChallenge(verifier);
  const state = randomBytesUrlSafe(16);

  const url = new URL(`${LOGIN_BASE}/${encodeURIComponent(tenant)}/oauth2/v2.0/authorize`);
  url.searchParams.set("client_id", clientId);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", redirectUri);
  url.searchParams.set("scope", SCOPES);
  url.searchParams.set("response_mode", "query");
  url.searchParams.set("state", state);
  url.searchParams.set("code_challenge", challenge);
  url.searchParams.set("code_challenge_method", "S256");
  // Always show the account picker: prevents the window from disappearing
  // instantly when a Microsoft session already exists.
  url.searchParams.set("prompt", "select_account");

  return { url: url.toString(), redirectUri, verifier, state };
}

async function handleAuthRedirect(redirect, req) {
  const ru = new URL(redirect);

  if (ru.searchParams.get("error")) {
    throw new Error(ru.searchParams.get("error_description") || ru.searchParams.get("error"));
  }
  if (ru.searchParams.get("state") !== req.state) {
    throw new Error("State parameter mismatch (CSRF protection). Aborting.");
  }
  const code = ru.searchParams.get("code");
  if (!code) throw new Error("No authorization code received from Microsoft.");

  const tok = await tokenRequest({
    grant_type: "authorization_code",
    code,
    redirect_uri: req.redirectUri,
    code_verifier: req.verifier,
    client_id: (await getState()).clientId,
  });
  await saveTokens(tok);
  await fetchAccountName().catch(() => {});
  await loadSections().catch(() => {});
  return buildPublicState();
}

async function signIn() {
  const req = await createAuthRequest();
  const redirect = await browser.identity.launchWebAuthFlow({ url: req.url, interactive: true });
  return handleAuthRedirect(redirect, req);
}

/* Alternative sign-in in a regular tab (fallback in case
   launchWebAuthFlow does not work reliably in a popup window). */
async function beginManualAuth() {
  const req = await createAuthRequest();
  await browser.storage.local.set({
    pendingAuth: { verifier: req.verifier, state: req.state, created: Date.now() },
  });
  return { url: req.url, redirectUri: req.redirectUri };
}

async function completeManualAuth(pastedUrl) {
  const { pendingAuth } = await browser.storage.local.get({ pendingAuth: null });
  if (!pendingAuth) {
    throw new Error("No pending sign-in found. Please run step 1 again.");
  }
  await browser.storage.local.remove("pendingAuth");
  if (Date.now() - pendingAuth.created > 15 * 60 * 1000) {
    throw new Error("The sign-in attempt has expired. Please run step 1 again.");
  }
  return handleAuthRedirect(pastedUrl, {
    verifier: pendingAuth.verifier,
    state: pendingAuth.state,
    redirectUri: browser.identity.getRedirectURL(),
  });
}

async function getAccessToken() {
  const s = await getState();
  if (s.accessToken && Date.now() < s.accessTokenExp - 60 * 1000) return s.accessToken;
  if (!s.refreshToken) throw new Error("Not signed in. Please sign in with Microsoft in the popup.");
  let tok;
  try {
    tok = await tokenRequest({
      grant_type: "refresh_token",
      refresh_token: s.refreshToken,
      client_id: s.clientId,
      scope: SCOPES,
    });
  } catch (e) {
    if (/invalid_grant|interaction_required/i.test(e.message)) {
      // Refresh token expired/revoked -> sign out cleanly instead of looping
      await browser.storage.local.set({ refreshToken: null, accessToken: null, accessTokenExp: 0 });
      throw new Error("Your sign-in has expired or was revoked. Please sign in with Microsoft again in the popup.");
    }
    throw e;
  }
  await saveTokens(tok);
  return tok.access_token;
}

async function signOut() {
  await browser.storage.local.set({
    refreshToken: null,
    accessToken: null,
    accessTokenExp: 0,
    accountName: null,
    sections: null,
    sectionsFetched: 0,
  });
  return buildPublicState();
}

/* ----------------------------------------------------------------- Graph */

async function graphFetch(path, options = {}, retry = true) {
  const token = await getAccessToken();
  const url = path.startsWith("http") ? path : GRAPH_BASE + path;
  const res = await fetch(url, {
    ...options,
    headers: { ...(options.headers || {}), Authorization: `Bearer ${token}` },
  });
  if (res.status === 401 && retry) {
    // Token may have been invalidated server-side -> refresh once and retry
    await browser.storage.local.set({ accessToken: null, accessTokenExp: 0 });
    return graphFetch(path, options, false);
  }
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Microsoft Graph: HTTP ${res.status} – ${text.slice(0, 400)}`);
  }
  return res;
}

async function fetchAccountName() {
  const res = await graphFetch("/me?$select=displayName,userPrincipalName,mail");
  const me = await res.json();
  const name = me.displayName || me.userPrincipalName || me.mail || "Microsoft account";
  await browser.storage.local.set({ accountName: name });
}

async function loadSections() {
  // Deliberately simple query: $orderby and nested $select inside $expand
  // are not consistently supported by the OneNote endpoints.
  let url = "/me/onenote/sections?$select=id,displayName&$expand=parentNotebook&$top=100";
  const all = [];
  let guard = 0;
  while (url && guard++ < 20) {
    const res = await graphFetch(url);
    const data = await res.json();
    for (const s of data.value || []) {
      all.push({
        id: s.id,
        name: s.displayName || "(unnamed)",
        notebook: s.parentNotebook?.displayName || "OneNote",
      });
    }
    url = data["@odata.nextLink"] || null;
  }
  all.sort((a, b) => a.notebook.localeCompare(b.notebook) || a.name.localeCompare(b.name));
  await browser.storage.local.set({ sections: all, sectionsFetched: Date.now() });
  return all;
}

/* ------------------------------------------------- Email -> OneNote page */

function sanitizeToXhtml(html) {
  try {
    const doc = new DOMParser().parseFromString(html, "text/html");
    doc.querySelectorAll(
      "script,style,iframe,frame,frameset,object,embed,form,input,button,select,textarea,link,meta,base,audio,video,canvas,svg,applet"
    ).forEach((el) => el.remove());
    for (const el of doc.querySelectorAll("*")) {
      for (const attr of [...el.attributes]) {
        const n = attr.name.toLowerCase();
        if (n.startsWith("on") || n === "srcset" || n === "formaction" || n === "background" || n === "class" || n === "id") {
          el.removeAttribute(attr.name);
        } else if ((n === "href" || n === "src") && /^\s*(javascript|vbscript):/i.test(attr.value)) {
          el.removeAttribute(attr.name);
        }
      }
    }
    const xs = new XMLSerializer();
    return [...doc.body.childNodes]
      .map((n) => xs.serializeToString(n))
      .join("")
      .replaceAll(' xmlns="http://www.w3.org/1999/xhtml"', "");
  } catch (e) {
    return `<p>${escapeXml(html)}</p>`;
  }
}

function plainToHtml(text) {
  return escapeXml(text)
    .split(/\r?\n/)
    .map((line) => `<div>${line || "<br />"}</div>`)
    .join("");
}

function buildPage(message, parsed, includeAttachments) {
  const binParts = []; // {name, blob}
  const warnings = [];
  let bytesTotal = 0;

  // Body (HTML preferred, plain text otherwise)
  let bodyHtml;
  if (parsed.htmlBody) {
    bodyHtml = sanitizeToXhtml(parsed.htmlBody);
    // Embed inline images (cid:) as binary parts
    for (const img of parsed.inlineImages) {
      if (!bodyHtml.includes(`cid:${img.contentId}`)) continue;
      if (bytesTotal + img.bytes.length > MAX_ATTACHMENT_TOTAL) {
        warnings.push(`Image "${img.filename || img.contentId}" skipped (size limit).`);
        continue;
      }
      const name = `img${binParts.length}`;
      binParts.push({ name, blob: new Blob([img.bytes], { type: img.mimeType }) });
      bytesTotal += img.bytes.length;
      bodyHtml = bodyHtml.split(`cid:${img.contentId}`).join(`name:${name}`);
    }
    // Remove unresolvable cid: images entirely (OneNote does not understand
    // cid: and may reject <img> without a valid src)
    bodyHtml = bodyHtml.replace(/<img\b[^>]*?src="cid:[^"]*"[^>]*?>/gi, "");
  } else if (parsed.textBody) {
    bodyHtml = plainToHtml(parsed.textBody);
  } else {
    bodyHtml = "<p><i>(No message body available – possibly a very large or encrypted message.)</i></p>";
  }

  // Attachments
  let attachmentHtml = "";
  const attachedNames = [];
  if (includeAttachments && parsed.attachments.length) {
    const objects = [];
    for (const att of parsed.attachments) {
      const label = att.filename || "attachment";
      if (bytesTotal + att.bytes.length > MAX_ATTACHMENT_TOTAL) {
        warnings.push(`Attachment "${label}" skipped (20 MB total size limit).`);
        continue;
      }
      const name = `att${binParts.length}`;
      binParts.push({ name, blob: new Blob([att.bytes], { type: att.mimeType || "application/octet-stream" }) });
      bytesTotal += att.bytes.length;
      objects.push(
        `<object data="name:${name}" data-attachment="${escapeXml(label)}" type="${escapeXml(att.mimeType || "application/octet-stream")}" />`
      );
      attachedNames.push(label);
    }
    if (objects.length) {
      attachmentHtml = `<hr /><p><b>Attachments:</b></p><p>${objects.join("<br />")}</p>`;
    }
  }

  // Header table
  const dateStr = message.date ? new Date(message.date).toLocaleString() : "";
  const rows = [
    ["From", message.author || ""],
    ["To", (message.recipients || []).join(", ")],
    ["Date", dateStr],
    ["Subject", message.subject || "(no subject)"],
  ]
    .map(([k, v]) => `<tr><td><b>${escapeXml(k)}</b></td><td>${escapeXml(v)}</td></tr>`)
    .join("");

  const title = (message.subject || "(no subject)").slice(0, 200);

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<title>${escapeXml(title)}</title>
<meta name="created" content="${new Date().toISOString()}" />
</head>
<body>
<table>${rows}</table>
<hr />
${bodyHtml}
${attachmentHtml}
</body>
</html>`;

  return { html, binParts, warnings, attachedNames };
}

async function normalizeRawMessage(raw) {
  // getRaw() returns a string, a Blob/File or binary data depending on version.
  // The MIME parser expects a string in which each character maps to one
  // original byte (latin1 decoding = 1:1 byte mapping).
  if (typeof raw === "string") return raw;
  if (raw == null) throw new Error("Could not read the message (empty raw data).");
  if (typeof Blob !== "undefined" && raw instanceof Blob) {
    const buf = await raw.arrayBuffer();
    return new TextDecoder("iso-8859-1").decode(buf);
  }
  if (raw instanceof ArrayBuffer) {
    return new TextDecoder("iso-8859-1").decode(raw);
  }
  if (ArrayBuffer.isView(raw)) {
    return new TextDecoder("iso-8859-1").decode(raw);
  }
  if (typeof raw.text === "function") {
    return raw.text();
  }
  throw new Error(`Unknown raw data format: ${Object.prototype.toString.call(raw)}`);
}

async function clipMessage(messageId, sectionId, includeAttachments) {
  let stage = "Reading message metadata";
  try {
    const message = await browser.messages.get(messageId);

    stage = "Reading raw message";
    const raw = await browser.messages.getRaw(messageId);

    stage = "Decoding message";
    const normalized = await normalizeRawMessage(raw);
    const parsed = parseMimeMessage(normalized);

    stage = "Building page";
    const { html, binParts, warnings, attachedNames } = buildPage(message, parsed, includeAttachments);

  /*
   * Build the multipart body manually instead of using FormData: this gives
   * the Presentation part exactly the shape documented by OneNote (no
   * filename attribute, explicit content type with charset), and binary
   * parts keep their MIME types.
   */
  const boundary = "----OneNoteClipper" + randomBytesUrlSafe(12);
  const chunks = [
    `--${boundary}\r\nContent-Disposition: form-data; name="Presentation"\r\nContent-Type: text/html; charset=utf-8\r\n\r\n${html}\r\n`,
  ];
  for (const p of binParts) {
    chunks.push(
      `--${boundary}\r\nContent-Disposition: form-data; name="${p.name}"\r\nContent-Type: ${p.blob.type || "application/octet-stream"}\r\n\r\n`
    );
    chunks.push(p.blob);
    chunks.push("\r\n");
  }
  chunks.push(`--${boundary}--\r\n`);
  const body = new Blob(chunks);

  stage = "Sending to OneNote";
  const res = await graphFetch(`/me/onenote/sections/${encodeURIComponent(sectionId)}/pages`, {
    method: "POST",
    headers: { "Content-Type": `multipart/form-data; boundary=${boundary}` },
    body,
  });
  const page = await res.json();
  await browser.storage.local.set({ lastSectionId: sectionId });
  return {
    ok: true,
    pageUrl: page?.links?.oneNoteWebUrl?.href || null,
    warnings,
    attachedNames,
  };
  } catch (e) {
    // Include the processing stage so errors can be pinpointed immediately
    throw new Error(`[${stage}] ${e.message}`);
  }
}

/* ------------------------------------------------------------ Message API */

async function buildPublicState() {
  const s = await getState();
  return {
    clientIdSet: !!s.clientId,
    tenant: s.tenant,
    signedIn: !!s.refreshToken,
    accountName: s.accountName,
    sections: s.sections || [],
    sectionsFetched: s.sectionsFetched || 0,
    lastSectionId: s.lastSectionId,
  };
}

browser.runtime.onMessage.addListener((msg) => {
  switch (msg?.type) {
    case "getState":
      return buildPublicState();
    case "signIn":
      return signIn();
    case "beginManualAuth":
      return beginManualAuth();
    case "completeManualAuth":
      return completeManualAuth(msg.url);
    case "signOut":
      return signOut();
    case "refreshSections":
      return loadSections().then(() => buildPublicState());
    case "clip":
      return clipMessage(msg.messageId, msg.sectionId, msg.includeAttachments);
    default:
      return Promise.resolve({ error: "Unknown message" });
  }
});
