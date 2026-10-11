const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { test } = require('node:test');
const { join } = require('node:path');
const vm = require('node:vm');

// Run the actual app offline with synthetic auth, Graph responses and DOM nodes.
const source = readFileSync(join(__dirname, '../app.js'), 'utf8');

function element() {
  const classes = new Set();
  return {
    hidden: true,
    _text: '',
    get textContent() { return this._text; },
    set textContent(value) { this._text = value; this.children = []; },
    children: [],
    classList: {
      add: (name) => classes.add(name),
      remove: (name) => classes.delete(name),
      contains: (name) => classes.has(name),
      toggle(name, on) { on ? classes.add(name) : classes.delete(name); },
    },
    setAttribute() {},
    removeAttribute(name) { delete this[name]; },
    append(...children) { this.children.push(...children); },
    appendChild(child) { this.children.push(child); },
  };
}

async function boot({ signedIn = true } = {}) {
  const elements = new Map();
  const timers = new Map();
  const listeners = new Map();
  const state = { tokenAvailable: signedIn, graphStatus: 200, now: new Date(2026, 9, 11, 12).getTime(), created: 0 };
  class Clock extends Date {
    constructor(...args) { super(...(args.length ? args : [state.now])); }
    static now() { return state.now; }
  }
  class InteractionRequiredAuthError extends Error {}
  class PublicClientApplication {
    async initialize() {}
    async handleRedirectPromise() { return null; }
    getAllAccounts() { return signedIn ? [{}] : []; }
    async acquireTokenSilent() {
      if (!state.tokenAvailable) throw new InteractionRequiredAuthError();
      return { accessToken: 'synthetic' };
    }
    async logoutRedirect() {}
  }
  const document = {
    body: element(),
    visibilityState: 'visible',
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, element());
      return elements.get(id);
    },
    createElement() { state.created++; return element(); },
    addEventListener: (name, callback) => listeners.set(name, callback),
  };
  const context = vm.createContext({
    window: {
      CAL_CONFIG: {
        clientId: 'synthetic-client', scopes: ['Calendars.Read'],
        lookaheadDays: 7, refreshIntervalMs: 300000,
      },
      location: { origin: 'http://localhost', pathname: '/' },
    },
    document,
    msal: { PublicClientApplication, InteractionRequiredAuthError },
    localStorage: { getItem() { return null; }, setItem() {} },
    setInterval: (callback, ms) => timers.set(ms, callback),
    Date: Clock,
    URLSearchParams,
    console,
    fetch: async () => state.fetch ? state.fetch() : ({
      ok: state.graphStatus === 200,
      status: state.graphStatus,
      text: async () => '',
      json: async () => {
        if (state.beforeJson) await state.beforeJson;
        return { value: [1, 2].map((offset) => ({
          subject: `Synthetic event ${offset}`,
          start: { dateTime: new Date(state.now + offset * 3600000).toISOString() },
          end: { dateTime: new Date(state.now + (offset + 1) * 3600000).toISOString() },
          location: { displayName: 'Synthetic location' },
          onlineMeetingUrl: 'https://example.invalid/synthetic',
        })) };
      },
    }),
  });
  vm.runInContext(source, context);
  // main() has no exported promise; drain its async boot through the event loop.
  await new Promise(setImmediate);
  return {
    state,
    el: (id) => elements.get(id),
    tick: () => timers.get(30000)(),
    periodicRefresh: () => timers.get(300000)(),
    async visible() {
      listeners.get('visibilitychange')();
      await new Promise(setImmediate);
    },
    refresh: () => elements.get('refresh-btn').onclick(),
    signOut: () => elements.get('signout-btn').onclick(),
  };
}

function assertHidden(app) {
  assert.equal(app.el('next').hidden, true, 'headline stays hidden');
  assert.equal(app.el('agenda-wrap').hidden, true, 'agenda stays hidden');
}

function assertRestored(app) {
  assert.equal(app.el('next').hidden, false);
  assert.equal(app.el('next-subject').textContent, 'Synthetic event 1');
  assert.equal(app.el('agenda-wrap').hidden, false);
  assert.equal(app.el('auth-btn').hidden, true);
  assert.equal(app.el('signout-btn').hidden, false);
}

test('failed token renewal stays hidden through timer and visibility, then recovers', async () => {
  const app = await boot();
  assertRestored(app);
  app.state.tokenAvailable = false;
  await app.refresh();
  assertHidden(app);
  assert.equal(app.el('next-subject').textContent, '');
  assert.equal(app.el('agenda').children.length, 0);
  assert.equal(app.el('next-join').href, undefined);
  assert.equal(app.el('updated').textContent, '');
  const signInMessage = app.el('status-msg').textContent;
  app.tick();
  assertHidden(app);
  await app.visible();
  assertHidden(app);
  assert.equal(app.el('status-msg').textContent, signInMessage);
  app.state.tokenAvailable = true;
  await app.refresh();
  assertRestored(app);
});

test('signed-out visibility retains the sign-in state', async () => {
  const app = await boot({ signedIn: false });
  const signInMessage = app.el('status-msg').textContent;
  assert.match(signInMessage, /Sign in with your work account/);
  await app.visible();
  assertHidden(app);
  assert.equal(app.el('status-msg').textContent, signInMessage);
  assert.doesNotMatch(app.el('next-subject').textContent, /Nothing scheduled/);
});

for (const status of [401, 403]) {
  test(`Graph ${status} invalidates the calendar`, async () => {
    const app = await boot();
    app.state.graphStatus = status;
    await app.refresh();
    app.tick();
    await app.visible();
    assertHidden(app);
  });
}

test('sign-out immediately hides calendar content before redirect', async () => {
  const app = await boot();
  app.signOut();
  app.tick();
  await app.visible();
  assertHidden(app);
});

test('an in-flight response cannot restore content after sign-out', async () => {
  const app = await boot();
  let release;
  app.state.beforeJson = new Promise((resolve) => { release = resolve; });
  const pending = app.refresh();
  await new Promise(setImmediate);
  app.signOut();
  release();
  await pending;
  assertHidden(app);
  assert.equal(app.el('next-subject').textContent, '');
});

function deferred() {
  let resolve, reject;
  const promise = new Promise((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

function response(subject, now = new Date(2026, 9, 11, 12).getTime()) {
  return { ok: true, json: async () => ({ value: [{
    subject,
    start: { dateTime: new Date(now + 3600000).toISOString() },
    end: { dateTime: new Date(now + 7200000).toISOString() },
  }] }) };
}

for (const trigger of ['refresh', 'periodicRefresh', 'visible']) {
 test(`newer ${trigger} wins when an older response arrives last`, async () => {
  const app = await boot();
  const old = deferred();
  app.state.fetch = () => old.promise;
  const pending = app.refresh();
  await new Promise(setImmediate);
  app.state.fetch = () => response('Newest synthetic event');
  app.state.now += 300001;
  await app[trigger]();
  old.resolve(response('Old synthetic event'));
  await pending;
  assert.equal(app.el('next-subject').textContent, 'Newest synthetic event');
 });
}

for (const failure of ['network', 'access']) {
  test(`superseded ${failure} failure cannot change newer calendar state`, async () => {
    const app = await boot();
    const old = deferred();
    app.state.fetch = () => old.promise;
    const pending = app.refresh();
    await new Promise(setImmediate);
    app.state.fetch = () => response('Newest synthetic event');
    await app.refresh();
    if (failure === 'network') old.reject(new Error('synthetic network failure'));
    else old.resolve({ ok: false, status: 403, text: async () => '' });
    await pending;
    assert.equal(app.el('next-subject').textContent, 'Newest synthetic event');
    assert.equal(app.el('status').hidden, true);
  });
}

test('countdown ticks preserve agenda nodes and advance countdown styling', async () => {
  const app = await boot();
  const nodes = [...app.el('agenda').children];
  const created = app.state.created;
  app.state.now += 51 * 60000;
  app.tick();
  assert.equal(app.el('next-countdown').textContent, 'in 9m');
  assert.equal(app.el('next-countdown').className, 'countdown soon');
  assert.deepEqual(app.el('agenda').children, nodes);
  assert.equal(app.state.created, created);
  app.state.now += 10 * 60000;
  app.tick();
  assert.equal(app.el('next-countdown').textContent, 'in progress');
  assert.equal(app.el('next-countdown').className, 'countdown now');
  assert.equal(app.state.created, created);
});

test('expiry, day rollover and fresh responses still rebuild calendar content', async () => {
  const app = await boot();
  app.state.now += 121 * 60000;
  app.tick();
  assert.equal(app.el('next-subject').textContent, 'Synthetic event 2');
  const midnight = new Date(app.state.now);
  midnight.setHours(24, 0, 0, 0);
  app.state.now = midnight.getTime();
  app.tick();
  assert.equal(app.el('agenda-wrap').hidden, true);
  app.state.now = Date.now();
  await app.refresh();
  const nodes = [...app.el('agenda').children];
  await app.refresh();
  assert.notEqual(app.el('agenda').children[0], nodes[0]);
});

test('midnight updates day labels without waiting for a refresh', async () => {
  const app = await boot();
  const midnight = new Date(app.state.now);
  midnight.setHours(24, 0, 0, 0);
  app.state.fetch = () => ({ ok: true, json: async () => ({ value: [1, 2].map((offset) => ({
    subject: `Synthetic tomorrow ${offset}`,
    start: { dateTime: new Date(midnight.getTime() + offset * 3600000).toISOString() },
    end: { dateTime: new Date(midnight.getTime() + (offset + 1) * 3600000).toISOString() },
  })) }) });
  await app.refresh();
  assert.equal(app.el('agenda').children[0].children[0].textContent, 'Tomorrow');
  app.state.now = midnight.getTime();
  app.tick();
  assert.equal(app.el('agenda').children[0].children[0].textContent, 'Today');
});
