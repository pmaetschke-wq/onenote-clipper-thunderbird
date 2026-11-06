# OneNote Clipper (local) for Thunderbird / Betterbird

Save emails to OneNote with one click – like Outlook's "Send to OneNote", but for
Thunderbird and Betterbird, and **without any third-party server**. The add-on
talks directly to Microsoft (`login.microsoftonline.com`, `graph.microsoft.com`)
using **your own free Azure app registration**. Tokens stay in your local profile.

Built as a privacy-respecting alternative to server-based clippers: you own the
app registration, you can revoke access at any time, and there is no telemetry,
tracking or update backend.

## Features

- One-click clipping from the message view: headers (From/To/Date/Subject),
  formatted body, inline images and optionally attachments
- Target **section picker** grouped by notebook, with **live search filter**
  and memory of the last used section
- **Multi-select**: clip several emails in one go
- Alternative tab-based sign-in if the OAuth window misbehaves
- Own MIME parser (base64, quoted-printable, charsets, nested multiparts) –
  tested with unit and integration tests against realistic sample emails

## Requirements

- Thunderbird **128+** or Betterbird **128+** (Manifest V3)
- A Microsoft account with OneNote (personal or work/school)
- A Microsoft Entra directory for the app registration. **Personal accounts**
  (@outlook.com etc.) get one by creating a
  [free Azure account](https://azure.microsoft.com/free) (card only for identity
  verification, no charges; no paid Azure resources are used). Work/school
  accounts can use their existing directory.

## Setup (one time, ~5 minutes)

1. **Register the app** at [entra.microsoft.com](https://entra.microsoft.com)
   → *Entra ID* → *App registrations* → *New registration*:
   - Account types: *any organizational directory and personal Microsoft accounts*
     (or personal only)
   - Redirect URI: platform **"Mobile and desktop applications"** (**NOT "Web"!**),
     custom URI = the URL shown in the add-on settings
   - *API permissions* → *Microsoft Graph* → *Delegated* → **`Notes.ReadWrite`**
2. **Install the add-on**: *Tools* → *Add-ons and Themes* → gear icon →
   *Install Add-on From File…* → select the ZIP.
   (For temporary testing via `about:debugging`: the add-on ID stays stable,
   so the redirect URL does not change.)
3. **Enter the client ID** from the app overview in the add-on settings → Save.
4. **Sign in**: open an email → OneNote Clipper button → *Sign in with Microsoft*
   → consent to `Notes.ReadWrite`.

## Usage

1. Open an email, or select one or more emails in the list.
2. Click the **OneNote Clipper** button in the message view. The popup shows
   the detected message's subject for confirmation.
3. Pick a **section** (filter with the search box if the list is long).
4. Optionally toggle *Include attachments*.
5. *Save to OneNote* – the success message links to the created page.

New sections in OneNote appear via *Refresh sections* in the popup.

## Troubleshooting

| Error | Fix |
|---|---|
| "Account not found in tenant 'Microsoft Services'" | Personal account without a directory → create a free Azure account (see Requirements) |
| `AADSTS70002` / `AADSTS7000218` (client_secret required) | Redirect URI registered as "Web" → re-register it under "Mobile and desktop applications" |
| `AADSTS90023` (cross-origin token redemption) | Update the add-on (Origin workaround built in since v1.2); alternative: register the URI as "Single-page application" |
| `AADSTS50011` (redirect_uri mismatch) | Registered URI differs – copy it exactly from the settings, including trailing `/` |
| Sign-in window closes instantly | Usually fine (existing session redirects immediately); check the popup. If needed: "Alternative tab-based sign-in" |
| "No message found" | Click/open the email again, then reopen the popup |

More documentation (technical architecture, permissions explained) is available
at the bottom of the add-on's settings page.

## Limits

- ~20 MB total per page for inline images + attachments (larger items are skipped
  and reported)
- Page title = email subject (max. 200 characters)
- Encrypted emails (S/MIME/OpenPGP) may not have a readable body
- HTML is reduced to the subset OneNote supports (scripts, forms etc. are removed)

## Privacy

- Network traffic goes **only** to Microsoft (`login.microsoftonline.com`,
  `graph.microsoft.com`). No third-party servers, no analytics.
- Tokens and settings are stored only in your local profile (`storage.local`).
- *Sign out* deletes all local tokens. Deleting the app registration in Entra
  revokes all access immediately.

## Project status: shared as-is, no maintenance

This tool was built with AI assistance and is shared in the hope that it is useful
to others. **It works for the author – and that is where the story ends.** Please
respect the following:

- **No issues, bug reports, feature requests or ideas** – they will not be read
  or answered.
- **No pull requests** – they will not be reviewed or merged.
- **No support, no updates, no warranty** of any kind. Use at your own risk.

**Want to change or extend it? Fork it and make it your own.** The MIT license
explicitly encourages this. If you build something on top of it, you are the
maintainer of your fork – not this project.

For the technically curious: the code is plain JavaScript with no build step
(`manifest.json`, `background.js` = OAuth, Graph, page creation; `mime.js` =
MIME parser; `popup.*`, `options.*`, `auth.*` = UI). Comments are in English,
and the architecture is documented in the add-on's settings page.

## License & warranty

[MIT](LICENSE). Provided **as is**, without warranty of any kind. Works for the
author; use at your own risk. See "Project status" above: no maintenance, no
support, no contributions accepted.
