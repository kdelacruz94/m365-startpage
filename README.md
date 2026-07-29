# M365 Calendar Start Page

A single static page that shows what's next on your Microsoft 365 calendar,
styled to match the Everforest Dark Hard YASB bar. Two surfaces share one Entra
app registration:

| Surface | Where it runs | Auth flow |
|---|---|---|
| This start page | any browser, any machine | SPA + PKCE (MSAL.js) |
| YASB `calendar` widget | the desktop only | device code (MSAL Python) |

No secrets live in this repo. An SPA using PKCE has no client secret, and the
client ID is public by design.

**Privacy note.** Meeting subjects can carry patient names. This page holds
events in memory only — nothing about a meeting is written to disk, logged, or
sent anywhere except Microsoft. The **Blur** button hides subjects until you
hover, which is worth leaving on if your screen is visible to others.

---

## 1. Register the Entra app (one time)

You need permission to register apps in the Griffin Health tenant. Many
healthcare tenants disable this — if step 1 is greyed out, skip to
"If registration is blocked" below.

1. Go to <https://entra.microsoft.com> → **Applications** → **App registrations**
   → **New registration**.
   - **Name:** `Calendar Start Page`
   - **Supported account types:** *Accounts in this organizational directory only
     (Single tenant)*
   - **Redirect URI:** platform **Single-page application (SPA)**, value
     `https://kdelacruz94.github.io/m365-startpage/`
     (include the trailing slash — it must match exactly)
   - **Register**

2. On the **Overview** page, copy the **Application (client) ID**.

3. **Authentication** → **Add a platform** → **Single-page application** → add
   `http://localhost:8731/` as a second redirect URI, so you can test locally.

4. Still on **Authentication**, scroll to **Advanced settings** and set
   **Allow public client flows** = **Yes**. The YASB device-code script needs
   this; the browser page does not.

5. **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Delegated permissions** → tick **Calendars.Read** → **Add permissions**.
   - `Calendars.Read` does not normally require admin consent, so the first
     sign-in will just prompt you. If your tenant restricts user consent, click
     **Grant admin consent** (needs an admin) or ask IT.

### If registration is blocked

Ask IT for a **single-tenant SPA app registration with delegated
`Calendars.Read`** and the two redirect URIs above. It is read-only, has no
client secret, and cannot see anyone's calendar but your own — usually an easy
approval. Send them this README.

---

## 2. Configure

Paste the client ID into **two** places:

- `config.js` → `clientId`
- `%USERPROFILE%\.config\yasb\scripts\calendar_config.json` → `client_id`

The authority stays `.../organizations`, which keeps the tenant GUID out of a
public repo while still only admitting your work account.

---

## 3. Run it

**Locally:**

```sh
py -3.14 -m http.server 8731
# then open http://localhost:8731/
```

**On GitHub Pages:** push this repo, then Settings → Pages → deploy from
`main` / root. It lands at `https://kdelacruz94.github.io/m365-startpage/`.

**As your start page:** set that URL as the homepage / startup page. Note that
*New Tab* cannot be pointed at a custom URL from browser settings — that needs
an extension (Zen/Firefox: "New Tab Override"; Chromium: an extension using
`chrome_url_overrides`) or the `NewTabPageLocation` enterprise policy. On the
work laptop, both are likely your employer's call, so homepage or a pinned tab
is the realistic option there.

---

## 4. YASB widget

```sh
py -3.14 %USERPROFILE%\.config\yasb\scripts\calendar_login.py
```

Sign in once with the device code it prints. The refresh token is cached at
`%LOCALAPPDATA%\YASB\calendar_token.json`, and the widget renews silently from
then on. It polls every 2 minutes and shows e.g. `󰃰 in 25m  Ops Standup`, with
the next few meetings in the tooltip.

If it ever shows `󰀦 cal login`, the cached token is gone — rerun the line above.

---

## Known wrinkles

- **Zen/Firefox silent renewal.** Total Cookie Protection tends to break MSAL's
  hidden-iframe token renewal, so the page uses a full redirect for interactive
  sign-in and may bounce you through Microsoft occasionally. Chromium is
  smoother here.
- **Tokens in `localStorage`.** Standard for SPAs and scoped to read-only
  calendar access, but it does mean a token sits in browser storage on that
  machine. Use **Sign out** on shared machines.
