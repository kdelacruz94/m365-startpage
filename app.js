/* M365 calendar start page.
 *
 * Auth: MSAL.js public client (SPA + PKCE, no secret). Tokens live in
 * localStorage so a fresh tab can renew silently instead of bouncing you
 * through a sign-in every time.
 *
 * Data: Microsoft Graph /me/calendarView, read-only. Events are held in memory
 * only -- nothing about a meeting is ever written to disk or sent anywhere but
 * Microsoft. That matters here: meeting subjects can carry patient names.
 */
(() => {
  "use strict";

  const CFG = window.CAL_CONFIG;
  const GRAPH = "https://graph.microsoft.com/v1.0";

  const $ = (id) => document.getElementById(id);
  const el = {
    clock: $("clock"),
    today: $("today"),
    status: $("status"),
    statusMsg: $("status-msg"),
    statusAction: $("status-action"),
    next: $("next"),
    countdown: $("next-countdown"),
    nextTime: $("next-time"),
    subject: $("next-subject"),
    location: $("next-location"),
    join: $("next-join"),
    agendaWrap: $("agenda-wrap"),
    agenda: $("agenda"),
    updated: $("updated"),
    authBtn: $("auth-btn"),
    signoutBtn: $("signout-btn"),
    refreshBtn: $("refresh-btn"),
    privacyBtn: $("privacy-btn"),
  };

  let pca = null;   // PublicClientApplication (window.msal is the library)
  let account = null;
  let events = [];          // in-memory only
  let lastFetch = null;

  // ---------------------------------------------------------------- helpers
  const pad = (n) => String(n).padStart(2, "0");

  function fmtTime(d) {
    return d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  }

  function fmtRelative(ms) {
    if (ms <= 0) return "now";
    const mins = Math.round(ms / 60000);
    if (mins < 60) return `in ${mins}m`;
    const h = Math.floor(mins / 60);
    const m = mins % 60;
    if (h < 24) return m ? `in ${h}h ${m}m` : `in ${h}h`;
    const d = Math.floor(h / 24);
    return `in ${d}d`;
  }

  function dayLabel(d) {
    const today = new Date();
    const t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const d0 = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const diff = Math.round((d0 - t0) / 86400000);
    if (diff === 0) return "Today";
    if (diff === 1) return "Tomorrow";
    return d.toLocaleDateString([], { weekday: "long", month: "short", day: "numeric" });
  }

  // Graph returns naive local-time strings plus a timeZone field. We ask for
  // UTC via the Prefer header, so append Z and let the browser localise.
  function parseGraphDate(dt) {
    if (!dt || !dt.dateTime) return null;
    const raw = dt.dateTime;
    return new Date(raw.endsWith("Z") ? raw : raw + "Z");
  }

  function showStatus(msg, { action = null, isError = false } = {}) {
    el.status.hidden = false;
    el.status.classList.toggle("error", isError);
    el.statusMsg.textContent = msg;
    if (action) {
      el.statusAction.hidden = false;
      el.statusAction.textContent = action.label;
      el.statusAction.onclick = action.fn;
    } else {
      el.statusAction.hidden = true;
      el.statusAction.onclick = null;
    }
  }

  function hideStatus() {
    el.status.hidden = true;
    el.status.classList.remove("error");
  }

  // ------------------------------------------------------------------ clock
  function tickClock() {
    const now = new Date();
    el.clock.textContent = `${now.getHours()}:${pad(now.getMinutes())}`;
    el.today.textContent = now.toLocaleDateString([], {
      weekday: "long", month: "long", day: "numeric",
    });
  }

  // ------------------------------------------------------------------- auth
  async function initAuth() {
    if (!CFG.clientId || CFG.clientId.startsWith("PASTE-")) {
      showStatus(
        "Not configured yet. Register the Entra app, then paste the Application (client) ID into config.js. See README.md.",
        { isError: true }
      );
      return false;
    }

    pca = new msal.PublicClientApplication({
      auth: {
        clientId: CFG.clientId,
        authority: CFG.authority,
        redirectUri: window.location.origin + window.location.pathname,
        navigateToLoginRequestUrl: true,
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false,
      },
    });

    await pca.initialize();

    // Completes a redirect sign-in if we just came back from one.
    const result = await pca.handleRedirectPromise();
    if (result && result.account) account = result.account;

    if (!account) {
      const all = pca.getAllAccounts();
      if (all.length) account = all[0];
    }
    return true;
  }

  async function getToken({ interactive = false } = {}) {
    const request = { scopes: CFG.scopes.filter((s) => s !== "offline_access") };
    if (account) request.account = account;

    if (!interactive && account) {
      try {
        const r = await pca.acquireTokenSilent(request);
        return r.accessToken;
      } catch (e) {
        // Fall through to interactive below.
        if (!(e instanceof msal.InteractionRequiredAuthError)) {
          console.warn("silent token failed:", e);
        }
      }
    }
    if (interactive) {
      // Redirect rather than popup: popups get blocked, and Firefox-family
      // storage partitioning makes iframe-based silent renewal unreliable.
      await pca.acquireTokenRedirect(request);
    }
    return null;
  }

  function signIn() {
    getToken({ interactive: true }).catch((e) => {
      showStatus(`Sign-in failed: ${e.message}`, { isError: true });
    });
  }

  function signOut() {
    events = [];
    if (pca && account) {
      pca.logoutRedirect({ account }).catch(() => {});
    }
  }

  // ------------------------------------------------------------------ graph
  async function fetchEvents() {
    const token = await getToken();
    if (!token) {
      el.next.hidden = true;
      el.agendaWrap.hidden = true;
      el.authBtn.hidden = false;
      el.signoutBtn.hidden = true;
      showStatus("Sign in with your work account to see your calendar.", {
        action: { label: "Sign in with work account", fn: signIn },
      });
      return;
    }

    const start = new Date();
    const end = new Date(start.getTime() + CFG.lookaheadDays * 86400000);
    const params = new URLSearchParams({
      startDateTime: start.toISOString(),
      endDateTime: end.toISOString(),
      $orderby: "start/dateTime",
      $top: "60",
      $select: "subject,start,end,location,isAllDay,isCancelled,onlineMeeting,onlineMeetingUrl,showAs",
    });

    const res = await fetch(`${GRAPH}/me/calendarView?${params}`, {
      headers: {
        Authorization: `Bearer ${token}`,
        // Ask Graph to hand back UTC so we can localise consistently.
        Prefer: 'outlook.timezone="UTC"',
      },
    });

    if (!res.ok) {
      const body = await res.text().catch(() => "");
      if (res.status === 401 || res.status === 403) {
        showStatus(
          `Graph refused the request (${res.status}). The app may be missing Calendars.Read consent.`,
          { action: { label: "Sign in again", fn: signIn }, isError: true }
        );
        return;
      }
      throw new Error(`Graph ${res.status}: ${body.slice(0, 200)}`);
    }

    const data = await res.json();
    events = (data.value || [])
      .map((e) => ({
        subject: e.subject || "(no subject)",
        start: parseGraphDate(e.start),
        end: parseGraphDate(e.end),
        allDay: !!e.isAllDay,
        cancelled: !!e.isCancelled,
        location: (e.location && e.location.displayName) || "",
        join: (e.onlineMeeting && e.onlineMeeting.joinUrl) || e.onlineMeetingUrl || "",
      }))
      .filter((e) => e.start && !e.cancelled)
      .sort((a, b) => a.start - b.start);

    lastFetch = new Date();
    hideStatus();
    el.authBtn.hidden = true;
    el.signoutBtn.hidden = false;
    render();
  }

  // ----------------------------------------------------------------- render
  function nextEvent() {
    const now = Date.now();
    // "Next" = the next thing that hasn't ended yet, ignoring all-day blocks
    // so a full-day marker doesn't mask a real meeting.
    return events.find((e) => !e.allDay && e.end && e.end.getTime() > now) || null;
  }

  function render() {
    const ev = nextEvent();

    if (!ev) {
      el.next.hidden = false;
      el.next.classList.add("none");
      el.countdown.textContent = "clear";
      el.countdown.className = "countdown";
      el.nextTime.textContent = "";
      el.subject.textContent = `Nothing scheduled in the next ${CFG.lookaheadDays} days`;
      el.location.textContent = "";
      el.join.hidden = true;
    } else {
      el.next.hidden = false;
      el.next.classList.remove("none");
      const ms = ev.start.getTime() - Date.now();
      const started = ms <= 0;

      el.countdown.textContent = started ? "in progress" : fmtRelative(ms);
      el.countdown.className =
        "countdown" + (started ? " now" : ms <= 10 * 60000 ? " soon" : "");
      el.nextTime.textContent = `${fmtTime(ev.start)} – ${ev.end ? fmtTime(ev.end) : ""}`;
      el.subject.textContent = ev.subject;
      el.location.textContent = ev.location;
      el.join.hidden = !ev.join;
      if (ev.join) el.join.href = ev.join;
    }

    // Agenda: everything after the headline event.
    const rest = events.filter((e) => e !== ev && e.end && e.end.getTime() > Date.now());
    el.agenda.textContent = "";
    let lastDay = null;

    for (const e of rest) {
      const label = dayLabel(e.start);
      if (label !== lastDay) {
        const li = document.createElement("li");
        li.className = "daybreak";
        const span = document.createElement("span");
        span.className = "daylabel";
        span.textContent = label;
        li.appendChild(span);
        el.agenda.appendChild(li);
        lastDay = label;
      }

      const li = document.createElement("li");
      const when = document.createElement("span");
      when.className = "when";
      when.textContent = e.allDay ? "all day" : `${fmtTime(e.start)} – ${fmtTime(e.end)}`;
      const what = document.createElement("span");
      what.className = "what";
      what.textContent = e.subject;
      const where = document.createElement("span");
      where.className = "where";
      where.textContent = e.location;
      li.append(when, what, where);
      el.agenda.appendChild(li);
    }

    el.agendaWrap.hidden = rest.length === 0;
    if (lastFetch) el.updated.textContent = `updated ${fmtTime(lastFetch)}`;
  }

  // ------------------------------------------------------------------- boot
  function setPrivacy(on) {
    document.body.classList.toggle("privacy", on);
    el.privacyBtn.setAttribute("aria-pressed", String(on));
    try { localStorage.setItem("cal.privacy", on ? "1" : "0"); } catch {}
  }

  async function refresh() {
    try {
      await fetchEvents();
    } catch (e) {
      console.error(e);
      showStatus(`Couldn't load the calendar: ${e.message}`, {
        action: { label: "Try again", fn: refresh },
        isError: true,
      });
    }
  }

  async function main() {
    tickClock();
    setInterval(tickClock, 1000);

    // Countdown drifts if we only redraw on fetch.
    setInterval(() => { if (events.length) render(); }, 30000);

    const savedPrivacy = (() => {
      try { return localStorage.getItem("cal.privacy") === "1"; } catch { return false; }
    })();
    setPrivacy(CFG.privacyMode || savedPrivacy);

    el.privacyBtn.onclick = () => setPrivacy(!document.body.classList.contains("privacy"));
    el.refreshBtn.onclick = refresh;
    el.authBtn.onclick = signIn;
    el.signoutBtn.onclick = signOut;

    if (!(await initAuth())) return;

    await refresh();
    setInterval(refresh, CFG.refreshIntervalMs);

    // A start page often sits in a background tab for hours; catch up on focus.
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") {
        render();
        if (!lastFetch || Date.now() - lastFetch > CFG.refreshIntervalMs) refresh();
      }
    });
  }

  main().catch((e) => {
    console.error(e);
    showStatus(`Startup failed: ${e.message}`, { isError: true });
  });
})();
