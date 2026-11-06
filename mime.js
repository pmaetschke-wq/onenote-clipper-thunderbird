"use strict";

/*
 * Minimal, robust MIME parser for browser.messages.getRaw().
 *
 * Background: messages.getFull() does not reliably return text bodies
 * depending on the version. getRaw(), on the other hand, always returns the
 * complete original message – parsed here ourselves, independent of the
 * API version.
 *
 * Returns: {
 *   htmlBody:     string|null,
 *   textBody:     string|null,
 *   inlineImages: [{ contentId, mimeType, filename, bytes: Uint8Array }],
 *   attachments:  [{ filename, mimeType, bytes: Uint8Array }]
 * }
 */

function parseMimeMessage(raw) {
  const result = { htmlBody: null, textBody: null, inlineImages: [], attachments: [] };

  function splitHeaders(block) {
    const m = block.match(/\r?\n\r?\n/);
    let headerText, body;
    if (m) {
      headerText = block.slice(0, m.index);
      body = block.slice(m.index + m[0].length);
    } else {
      headerText = block;
      body = "";
    }
    // Unfold folded headers
    const lines = headerText.replace(/\r?\n[ \t]+/g, " ").split(/\r?\n/);
    const headers = {};
    for (const line of lines) {
      const idx = line.indexOf(":");
      if (idx <= 0) continue;
      const key = line.slice(0, idx).trim().toLowerCase();
      if (headers[key] === undefined) headers[key] = line.slice(idx + 1).trim();
    }
    return { headers, body };
  }

  function baseType(ct) {
    return (ct || "").split(";")[0].trim().toLowerCase();
  }

  function getParam(value, param) {
    const re = new RegExp(`${param}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^;\\s]+))`, "i");
    const m = (value || "").match(re);
    return m ? (m[1] ?? m[2] ?? m[3] ?? "").trim() : "";
  }

  function getRfc2231Param(value, param) {
    // filename*=utf-8''My%20Attachment.pdf
    const m = (value || "").match(new RegExp(`${param}\\*\\s*=\\s*([^;]+)`, "i"));
    if (!m) return "";
    const v = m[1].trim().replace(/^["']|["']$/g, "");
    const parts = v.split("'");
    const encoded = parts.length === 3 ? parts[2] : v;
    try {
      const bytes = Uint8Array.from(encoded.replace(/%([0-9A-Fa-f]{2})/g, (_, h) =>
        String.fromCharCode(parseInt(h, 16))
      ), (c) => c.charCodeAt(0) & 0xff);
      const charset = parts.length === 3 && parts[0] ? parts[0] : "utf-8";
      return new TextDecoder(charset).decode(bytes);
    } catch {
      return encoded;
    }
  }

  function decodeBytes(str, cte) {
    cte = (cte || "").toLowerCase().trim();
    if (cte === "base64") {
      const clean = str.replace(/[^A-Za-z0-9+/=]/g, "");
      try {
        const bin = atob(clean);
        const out = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
        return out;
      } catch {
        return new Uint8Array(0);
      }
    }
    if (cte === "quoted-printable") {
      const bytes = [];
      let i = 0;
      while (i < str.length) {
        if (str[i] === "=") {
          const soft = str.slice(i).match(/^=\r?\n/);
          if (soft) { i += soft[0].length; continue; }
          const hex = str.substr(i + 1, 2);
          if (/^[0-9A-Fa-f]{2}$/.test(hex)) {
            bytes.push(parseInt(hex, 16));
            i += 3;
            continue;
          }
        }
        bytes.push(str.charCodeAt(i) & 0xff);
        i++;
      }
      return new Uint8Array(bytes);
    }
    // 7bit / 8bit / binary
    const out = new Uint8Array(str.length);
    for (let i = 0; i < str.length; i++) out[i] = str.charCodeAt(i) & 0xff;
    return out;
  }

  function decodeText(bytes, charset) {
    const cs = (charset || "utf-8").trim().replace(/^["']|["']$/g, "");
    try {
      return new TextDecoder(cs).decode(bytes);
    } catch {
      try {
        return new TextDecoder("utf-8").decode(bytes);
      } catch {
        return "";
      }
    }
  }

  function decodeBodyText(str, cte, charset) {
    const c = (cte || "").toLowerCase().trim();
    const cs = (charset || "utf-8").trim().replace(/^["']|["']$/g, "");
    if (c === "base64" || c === "quoted-printable") {
      return decodeText(decodeBytes(str, c), cs);
    }
    // 7bit/8bit/binary: depending on the API version, the raw string may
    // already be decoded text OR still hold raw bytes (1 char = 1 byte).
    // Try strict byte decoding first; if it fails, the string was apparently
    // already correctly decoded.
    if (/^[\x00-\x7F]*$/.test(str)) return str; // pure ASCII
    try {
      return new TextDecoder(cs, { fatal: true }).decode(
        Uint8Array.from(str, (ch) => ch.charCodeAt(0) & 0xff)
      );
    } catch {
      return str;
    }
  }

  function parseEntity(block, depth) {
    if (depth > 12 || !block) return;
    const { headers, body } = splitHeaders(block);
    const ct = headers["content-type"] || "text/plain";
    const type = baseType(ct);
    const cte = headers["content-transfer-encoding"] || "";
    const disp = headers["content-disposition"] || "";
    const dispType = disp.split(";")[0].trim().toLowerCase();

    if (type.startsWith("multipart/")) {
      const boundary = getParam(ct, "boundary");
      if (!boundary) return;
      const dash = "--" + boundary;
      const entities = [];
      let current = null;
      for (const line of body.split(/\r?\n/)) {
        if (line.startsWith(dash)) {
          if (current !== null) entities.push(current);
          current = null;
          if (line.slice(dash.length).trim().startsWith("--")) break; // end
          current = "";
        } else if (current !== null) {
          current += line + "\r\n";
        }
      }
      if (current !== null) entities.push(current);
      for (const ent of entities) parseEntity(ent, depth + 1);
      return;
    }

    const filename =
      getRfc2231Param(disp, "filename") ||
      getParam(disp, "filename") ||
      getRfc2231Param(ct, "name") ||
      getParam(ct, "name");
    const contentId = (headers["content-id"] || "").replace(/[<>]/g, "").trim();

    if (contentId && type.startsWith("image/")) {
      result.inlineImages.push({
        contentId,
        mimeType: type,
        filename,
        bytes: decodeBytes(body, cte),
      });
      return;
    }
    if (dispType === "attachment" || filename) {
      result.attachments.push({
        filename: filename || (type === "message/rfc822" ? "attachment.eml" : "attachment"),
        mimeType: type,
        bytes: decodeBytes(body, cte),
      });
      return;
    }
    if (type === "text/html" && result.htmlBody === null) {
      result.htmlBody = decodeBodyText(body, cte, getParam(ct, "charset"));
      return;
    }
    if (type === "text/plain" && result.textBody === null) {
      result.textBody = decodeBodyText(body, cte, getParam(ct, "charset"));
    }
    // Ignore everything else (e.g. application/pgp-signature without a name)
  }

  parseEntity(raw, 0);
  return result;
}

// Export for local tests (not used inside the add-on)
if (typeof module !== "undefined" && module.exports) {
  module.exports = { parseMimeMessage };
}
