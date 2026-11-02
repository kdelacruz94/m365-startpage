const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { test } = require('node:test');
const vm = require('node:vm');

test('CSP permits the MSAL iframe authorization and redirect URI', () => {
  const html = readFileSync(join(__dirname, '../index.html'), 'utf8');
  const app = readFileSync(join(__dirname, '../app.js'), 'utf8');
  const meta = html.match(/<meta\b[^>]*http-equiv="Content-Security-Policy"[^>]*>/i);
  assert.ok(meta, 'index.html must declare its CSP');
  const content = meta[0].match(/\bcontent="([^"]*)"/i);
  assert.ok(content, 'CSP meta must have content');
  const directives = content[1].split(';').map(value => value.trim().split(/\s+/));
  const frameSrc = directives.find(([name]) => name === 'frame-src');
  assert.ok(frameSrc, 'CSP must declare frame-src');
  const sources = frameSrc.slice(1);
  assert.ok(sources.includes('https://login.microsoftonline.com'),
    'MSAL authorization must be allowed in the iframe');

  const redirect = app.match(/\bredirectUri\s*:\s*([^,\n]+),/);
  assert.ok(redirect, 'app.js must configure redirectUri');
  const location = new URL('https://calendar.example.test/startpage/index.html');
  const redirectUri = vm.runInNewContext(`(${redirect[1]})`, {
    window: { location: { origin: location.origin, pathname: location.pathname } },
  });
  const redirectOrigin = new URL(redirectUri, location).origin;
  assert.ok(
    sources.includes(redirectOrigin) ||
      (sources.includes("'self'") && redirectOrigin === location.origin),
    `frame-src must allow the MSAL redirect origin ${redirectOrigin}`,
  );
});
