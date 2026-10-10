const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { test } = require("node:test");

const source = fs.readFileSync(path.join(__dirname, "../app.js"), "utf8");
const nextLink = "https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=synthetic-page-2";
const finalLink = "https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=synthetic-page-3";

function element() {
  const classes = new Set();
  return {
    hidden: true,
    textContent: "",
    children: [],
    classList: {
      add: (name) => classes.add(name),
      remove: (name) => classes.delete(name),
      contains: (name) => classes.has(name),
      toggle(name, on) { on ? classes.add(name) : classes.delete(name); },
    },
    appendChild(child) { this.children.push(child); },
    append(...children) { this.children.push(...children); },
    setAttribute() {},
  };
}

function event(subject, minutes, extra = {}) {
  return {
    subject,
    start: { dateTime: new Date(Date.now() + minutes * 60000).toISOString() },
    end: { dateTime: new Date(Date.now() + (minutes + 30) * 60000).toISOString() },
    ...extra,
  };
}

function page(value, link) {
  return { ok: true, json: async () => ({ value, "@odata.nextLink": link }) };
}

async function boot(responses) {
  const elements = new Map();
  const calls = [];
  const intervals = [];
  vm.runInNewContext(source, {
    window: {
      CAL_CONFIG: {
        clientId: "synthetic-client", scopes: ["Calendars.Read"],
        lookaheadDays: 7, refreshIntervalMs: 300000,
      },
      location: { origin: "https://example.invalid", pathname: "/" },
    },
    document: {
      getElementById(id) {
        if (!elements.has(id)) elements.set(id, element());
        return elements.get(id);
      },
      body: element(), createElement: element, addEventListener() {},
    },
    msal: {
      PublicClientApplication: class {
        async initialize() {}
        async handleRedirectPromise() { return null; }
        getAllAccounts() { return [{}]; }
        async acquireTokenSilent() { return { accessToken: "synthetic-placeholder" }; }
      },
    },
    localStorage: { getItem() {}, setItem() {} },
    fetch: async (url, options) => {
      calls.push({ url, options });
      assert.ok(responses.length, "unexpected fetch");
      const response = responses.shift();
      return typeof response === "function" ? response() : response;
    },
    setInterval: (fn, ms) => intervals.push({ fn, ms }),
    Date, URLSearchParams, console: { error() {}, warn() {} },
  });
  // Let the app's real asynchronous startup and fetch chain settle.
  await new Promise(setImmediate);
  return { elements, calls, intervals };
}

test("a qualifying event on page two replaces a cancelled first page", async () => {
  const { elements, calls } = await boot([
    page(Array.from({ length: 60 }, () => event("Synthetic cancelled", 10, { isCancelled: true })), nextLink),
    page([event("Synthetic page two", 20)]),
  ]);
  assert.equal(elements.get("next-subject").textContent, "Synthetic page two");
  assert.equal(calls.length, 2);
  assert.equal(calls[1].url, nextLink);
  assert.deepEqual(calls[1].options.headers, calls[0].options.headers);
});

test("all pages are assembled, filtered and sorted before publication", async () => {
  let finish;
  const { elements, calls } = await boot([
    page([event("Synthetic later", 90)], nextLink),
    page([
      event("Synthetic cancelled", 5, { isCancelled: true }),
      event("Synthetic middle", 60),
      event("Synthetic missing start", 10, { start: null }),
    ], finalLink),
    () => new Promise((resolve) => { finish = resolve; }),
  ]);
  assert.deepEqual(calls.slice(1).map((call) => call.url), [nextLink, finalLink]);
  assert.equal(elements.get("next").hidden, true, "no partial headline while a page is pending");
  assert.equal(elements.get("updated").textContent, "");
  finish(page([event("Synthetic earliest", 20)]));
  await new Promise(setImmediate);
  assert.equal(elements.get("next-subject").textContent, "Synthetic earliest");
  const subjects = elements.get("agenda").children
    .filter((row) => row.children.length === 3)
    .map((row) => row.children[1].textContent);
  assert.deepEqual(subjects, ["Synthetic middle", "Synthetic later"]);
  assert.equal(elements.get("status").hidden, true);
});

for (const status of [500, 403]) {
  test(`a later-page ${status} does not publish an empty or partial calendar`, async () => {
    const responses = [page([event("Synthetic previous", 20)])];
    const { elements, intervals } = await boot(responses);
    const updated = elements.get("updated").textContent;
    responses.push(
      page(Array.from({ length: 60 }, () => event("Synthetic cancelled", 10, { isCancelled: true })), nextLink),
      { ok: false, status, text: async () => "synthetic failure" },
    );
    await elements.get("refresh-btn").onclick();
    intervals.find((interval) => interval.ms === 30000).fn();
    assert.equal(elements.get("next-subject").textContent, "Synthetic previous");
    assert.equal(elements.get("updated").textContent, updated);
    assert.equal(elements.get("status").hidden, false);
    assert.match(elements.get("status-msg").textContent, /Graph/);
  });
}

test("a failed initial paginated load keeps the calendar unpublished", async () => {
  const { elements } = await boot([
    page([], nextLink),
    { ok: false, status: 500, text: async () => "synthetic failure" },
  ]);
  assert.equal(elements.get("next").hidden, true);
  assert.equal(elements.get("updated").textContent, "");
  assert.equal(elements.get("status").hidden, false);
  assert.match(elements.get("status-msg").textContent, /Couldn't load the calendar/);
});
