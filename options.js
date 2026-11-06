"use strict";

const $ = (id) => document.getElementById(id);

async function init() {
  // The redirect URL depends on the add-on ID and is therefore fixed –
  // it must be registered exactly like this in the Azure app registration.
  $("redirect-url").textContent = browser.identity.getRedirectURL();

  const s = await browser.storage.local.get({ clientId: "", tenant: "common" });
  $("client-id").value = s.clientId;
  $("tenant").value = s.tenant;
}

$("save").addEventListener("click", async () => {
  const clientId = $("client-id").value.trim();
  const tenant = $("tenant").value;

  if (clientId && !/^[0-9a-fA-F-]{36}$/.test(clientId)) {
    $("saved").textContent = "⚠️ This does not look like a valid client ID (GUID).";
    return;
  }

  // If the client ID changed, discard old tokens – they belong to the old app.
  const old = await browser.storage.local.get({ clientId: "" });
  const upd = { clientId, tenant };
  if (old.clientId && old.clientId !== clientId) {
    Object.assign(upd, {
      refreshToken: null, accessToken: null, accessTokenExp: 0,
      accountName: null, sections: null, sectionsFetched: 0,
    });
  }
  await browser.storage.local.set(upd);
  $("saved").textContent = "✅ Saved. Now sign in via the add-on popup.";
  setTimeout(() => { $("saved").textContent = ""; }, 6000);
});

$("signout").addEventListener("click", async () => {
  await browser.runtime.sendMessage({ type: "signOut" });
  $("saved").textContent = "✅ Signed out, tokens deleted.";
  setTimeout(() => { $("saved").textContent = ""; }, 6000);
});

init();
