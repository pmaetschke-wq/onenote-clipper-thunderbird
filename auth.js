"use strict";

const $ = (id) => document.getElementById(id);

function showMsg(text, kind) {
  const el = $("msg");
  el.className = kind;
  el.textContent = text;
}

// Show the redirect prefix so users know what to look for in the address bar
$("redirect-prefix").textContent = browser.identity.getRedirectURL();

$("open-auth").addEventListener("click", async () => {
  $("msg").className = "hidden";
  try {
    const res = await browser.runtime.sendMessage({ type: "beginManualAuth" });
    if (res?.error) throw new Error(res.error);
    $("redirect-prefix").textContent = res.redirectUri;
    await browser.tabs.create({ url: res.url });
  } catch (e) {
    showErrorSafe(e);
  }
});

$("complete").addEventListener("click", async () => {
  const url = $("pasted").value.trim();
  if (!url) {
    showMsg("Please paste the address from the sign-in tab's address bar first.", "error");
    return;
  }
  if (!/^https:\/\/\S+[?&](code|error)=/.test(url)) {
    showMsg("This does not look like the final address. Please copy the COMPLETE address – it must contain \"code=\" (or \"error=\").", "error");
    return;
  }
  try {
    const st = await browser.runtime.sendMessage({ type: "completeManualAuth", url });
    if (st?.error) throw new Error(st.error);
    showMsg(`✅ Signed in as ${st.accountName || "Microsoft account"}.\nYou can close this tab and start using the clipper in the message view.`, "ok");
  } catch (e) {
    showErrorSafe(e);
  }
});

function showErrorSafe(e) {
  showMsg(`Sign-in failed:\n${e.message || e}`, "error");
}
