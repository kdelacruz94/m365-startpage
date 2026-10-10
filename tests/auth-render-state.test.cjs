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
  const state = { tokenAvailable: signedIn, graphStatus: 200 };
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
    createElement: element,
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
    URLSearchParams,
    console,
    fetch: async () => ({
      ok: state.graphStatus === 200,
      status: state.graphStatus,
      text: async () => '',
      json: async () => {
        if (state.beforeJson) await state.beforeJson;
        return { value: [1, 2].map((offset) => ({
          subject: `Synthetic event ${offset}`,
          start: { dateTime: new Date(Date.now() + offset * 3600000).toISOString() },
          end: { dateTime: new Date(Date.now() + (offset + 1) * 3600000).toISOString() },
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
