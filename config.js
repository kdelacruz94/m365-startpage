// ---------------------------------------------------------------------------
// Fill in clientId after registering the app (see README.md).
// Nothing here is a secret: an SPA using PKCE has no client secret, and the
// client ID is public by design. The tenant GUID is deliberately NOT hardcoded
// -- "organizations" keeps your employer's tenant id out of a public repo while
// still only admitting accounts from tenants the app is registered in.
// ---------------------------------------------------------------------------
window.CAL_CONFIG = {
  clientId: "e3fe964c-6413-4d71-aefb-8de528e548f3",

  // Use "organizations" for work/school accounts. Swap for your tenant GUID if
  // you'd rather pin it explicitly (and keep the repo private).
  authority: "https://login.microsoftonline.com/organizations",

  // Read-only. offline_access lets MSAL refresh silently in the background.
  scopes: ["Calendars.Read", "offline_access"],

  // How far ahead the agenda looks.
  lookaheadDays: 3,

  // How often to re-query Graph (ms). The countdown re-renders far more often.
  refreshIntervalMs: 5 * 60 * 1000,

  // Hide meeting subjects until you hover/focus the page. Useful when the
  // screen is visible to others and subjects may name patients.
  privacyMode: false,
};
