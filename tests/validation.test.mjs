import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const backend = readFileSync(resolve(root, 'backend/Code.gs'), 'utf8');
const siteBase = 'https://example.com/juki-search-test/';
const headers = ['timestamp', 'sku', 'machine_model', 'part_number', 'email', 'preference', 'notes', 'page_url', 'lead_status', 'owner_minutes', 'quoted', 'paid', 'estimated_contribution', 'actual_contribution'];
const skus = ['110-38650', '400-90753', '110-40359'];

function request(overrides = {}) {
  const fields = {
    machine_model: 'Unknown', part_number: 'Unknown', preference: 'either',
    email: 'buyer@example.com', notes: '', sku: '400-90753',
    page_url: siteBase + 'juki/400-90753.html', ...overrides
  };
  const parameters = {};
  const body = new URLSearchParams();
  for (const [key, value] of Object.entries(fields)) {
    if (value === undefined) continue;
    parameters[key] = Array.isArray(value) ? value : [value];
    for (const item of parameters[key]) body.append(key, item);
  }
  return { parameters, postData: { type: 'application/x-www-form-urlencoded', contents: body.toString() } };
}

function harness(options = {}) {
  const rows = options.rows ?? [headers.slice()];
  const mail = [], logs = [], calls = [];
  const properties = { SITE_BASE_URL: siteBase, NOTIFICATION_EMAIL: 'operator@example.com', ...options.properties };
  const sheet = {
    getLastRow: () => rows.length,
    getLastColumn: () => Math.max(0, ...rows.map(row => row.length)),
    appendRow(row) {
      calls.push('append');
      if (options.writeFailure) throw new Error('PRIVATE spreadsheet failure');
      rows.push(Array.from(row));
    },
    getRange(row, column, rowCount = 1, columnCount = 1) {
      return {
        getDisplayValues: () => rows.slice(row - 1, row - 1 + rowCount).map(values => values.slice(column - 1, column - 1 + columnCount).map(String)),
        setValue(value) { rows[row - 1][column - 1] = value; }
      };
    }
  };
  const context = vm.createContext({
    Date,
    console: { error: message => logs.push(message) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: key => properties[key] }) },
    LockService: { getScriptLock: () => ({
      tryLock: () => { calls.push('lock'); return !options.busy; },
      releaseLock: () => calls.push('unlock')
    }) },
    SpreadsheetApp: {
      openById(id) {
        assert.equal(id, '16IcDQZ9omcNJaIzrkCXfK53N--qwcxhR5QQuS5P1-Ts');
        return { getSheetByName(name) { assert.equal(name, 'Leads'); return options.missingSheet ? null : sheet; } };
      },
      flush() { calls.push('flush'); }
    },
    GmailApp: { sendEmail(...args) {
      calls.push('mail');
      if (options.mailFailure) throw new Error('PRIVATE Gmail failure buyer@example.com');
      mail.push(args);
    } },
    HtmlService: { createHtmlOutput(html) { return { html, setTitle(title) { this.title = title; return this; } }; } }
  });
  vm.runInContext(backend, context, { filename: 'Code.gs' });
  return { rows, mail, logs, calls, post: event => context.doPost(event), get: () => context.doGet() };
}

for (const sku of skus) {
  test(`stores ${sku} with all 14 columns, sends notification after flush, then redirects`, () => {
    const h = harness();
    const response = h.post(request({ sku, page_url: siteBase + `juki/${sku}.html` }));
    assert.equal(response.title, 'Request recorded');
    assert.equal(h.rows.length, 2);
    assert.ok(h.rows[1][0] instanceof Date);
    assert.deepEqual(h.rows[1].slice(1), [sku, 'Unknown', 'Unknown', 'buyer@example.com', 'either', '', siteBase + `juki/${sku}.html`, 'new', '', '', '', '', '']);
    assert.deepEqual(h.calls, ['lock', 'append', 'flush', 'unlock', 'mail']);
    assert.equal(h.mail[0][0], 'operator@example.com');
    assert.equal(h.mail[0][3].replyTo, 'buyer@example.com');
    assert.match(h.mail[0][2], /Machine model: Unknown/);
    assert.match(response.html, /window\.location\.replace\("https:\/\/example\.com\/juki-search-test\/thank-you\.html"\)/);
    assert.match(response.html, /target="_top"/);
    assert.match(response.html, /noindex,nofollow/);
    assert.doesNotMatch(response.html, /buyer@example\.com|16IcDQZ9/);
  });
}

test('sanitizes control characters, strips tracking URL data, and escapes formula text', () => {
  const h = harness();
  const response = h.post(request({ machine_model: '\t=HYPERLINK("bad")\u0000', part_number: '+110-96500', notes: '  @SUM(1)\r\nline\u0007\t2 ', page_url: siteBase + 'juki/400-90753.html?email=private@example.com#tracking' }));
  assert.equal(response.title, 'Request recorded');
  assert.equal(h.rows[1][2], '\'=HYPERLINK("bad")');
  assert.equal(h.rows[1][3], "'+110-96500");
  assert.equal(h.rows[1][6], "'@SUM(1)\nline\t2");
  assert.equal(h.rows[1][7], siteBase + 'juki/400-90753.html');
  assert.doesNotMatch(response.html, /HYPERLINK|SUM|private@example/);
});

test('optional notes can be omitted and an empty ledger receives the exact header', () => {
  const h = harness({ rows: [] });
  assert.equal(h.post(request({ notes: undefined })).title, 'Request recorded');
  assert.deepEqual(h.rows[0], headers);
  assert.equal(h.rows[1][6], '');
});

test('HTML in notes is stored as text, never reflected in public HTML or logs', () => {
  const h = harness();
  const payload = '<script>alert("private")</script>';
  const response = h.post(request({ notes: payload }));
  assert.equal(h.rows[1][6], payload);
  assert.ok(!response.html.includes(payload));
  assert.equal(h.logs.length, 0);
});

const rejected = [
  ['historical reference as SKU', { sku: '110-96500' }],
  ['unlisted SKU', { sku: '123-45678' }],
  ['duplicate email', { email: ['first@example.com', 'second@example.com'] }],
  ['unknown field / redirect', { redirect_url: 'https://evil.example/' }],
  ['client ledger field', { paid: 'true' }],
  ['missing machine model', { machine_model: undefined }],
  ['blank part number', { part_number: '   ' }],
  ['bad email', { email: 'buyer@@example.com' }],
  ['bad domain', { email: 'buyer@example..com' }],
  ['email header injection', { email: 'buyer@example.com\r\nBcc:other@example.com' }],
  ['invalid preference', { preference: 'best' }],
  ['external page URL', { page_url: 'https://evil.example/juki/400-90753.html' }],
  ['wrong SKU page URL', { page_url: siteBase + 'juki/110-38650.html' }],
  ['overlong notes', { notes: 'x'.repeat(2001) }]
];
for (const [name, fields] of rejected) {
  test(`rejects ${name} without ledger write, Gmail or redirect`, () => {
    const h = harness();
    const response = h.post(request(fields));
    assert.equal(response.title, 'Request not confirmed');
    assert.equal(h.rows.length, 1);
    assert.equal(h.mail.length, 0);
    assert.doesNotMatch(response.html, /window\.location|buyer@example|PRIVATE|evil\.example/);
    assert.ok(h.logs.every(log => !log.includes('buyer@example')));
  });
}

test('rejects malformed or oversized bodies and unsupported content types', () => {
  for (const event of [undefined, {}, { ...request(), postData: { type: 'application/json', contents: '{}' } }, { ...request(), postData: { type: 'application/x-www-form-urlencoded', contents: 'x'.repeat(32769) } }]) {
    const h = harness();
    assert.equal(h.post(event).title, 'Request not confirmed');
    assert.equal(h.rows.length, 1);
    assert.equal(h.mail.length, 0);
  }
});

test('preserves existing ledger rows and refuses a different schema', () => {
  const rows = [headers.slice(), ['existing record']];
  rows[0][4] = 'wrong_email_column';
  const snapshot = JSON.stringify(rows);
  const h = harness({ rows });
  assert.equal(h.post(request()).title, 'Request not confirmed');
  assert.equal(JSON.stringify(rows), snapshot);
  assert.equal(h.mail.length, 0);
  assert.ok(h.calls.includes('unlock'));
});

test('missing tab, busy lock and storage errors never become success', () => {
  for (const options of [{ missingSheet: true }, { busy: true }, { writeFailure: true }]) {
    const h = harness(options);
    const response = h.post(request());
    assert.equal(response.title, 'Request not confirmed');
    assert.equal(h.rows.length, 1);
    assert.equal(h.mail.length, 0);
    assert.doesNotMatch(response.html, /PRIVATE|window\.location/);
    if (!options.busy) assert.ok(h.calls.includes('unlock'));
  }
});

test('Gmail failure preserves stored success and flags the row for operator follow-up', () => {
  const h = harness({ mailFailure: true });
  const response = h.post(request());
  assert.equal(response.title, 'Request recorded');
  assert.equal(h.rows.length, 2);
  assert.equal(h.rows[1][8], 'new_notification_failed');
  assert.match(response.html, /thank-you\.html/);
  assert.ok(h.logs.every(log => !log.includes('buyer@example') && !log.includes('PRIVATE')));
});

test('requires valid operator configuration before storing any lead', () => {
  for (const properties of [
    { SITE_BASE_URL: '' }, { SITE_BASE_URL: 'http://example.com/' },
    { SITE_BASE_URL: 'https://example.com/path/../' },
    { SITE_BASE_URL: 'https://example.com/?redirect=bad' },
    { NOTIFICATION_EMAIL: '' }, { NOTIFICATION_EMAIL: 'one@example.com,two@example.com' }
  ]) {
    const h = harness({ properties });
    assert.equal(h.post(request()).title, 'Request not confirmed');
    assert.equal(h.rows.length, 1);
    assert.equal(h.mail.length, 0);
  }
});

test('GET never records a lead or sends mail', () => {
  const h = harness();
  assert.equal(h.get().title, 'Fit request endpoint');
  assert.equal(h.rows.length, 1);
  assert.equal(h.mail.length, 0);
});

test('catalog has exactly the requested SKU roles and reference-only historical number', () => {
  const data = JSON.parse(readFileSync(resolve(root, 'data/parts.json'), 'utf8'));
  assert.equal(data.system_version, '0.2');
  assert.ok(data.lead_endpoint === null || /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(data.lead_endpoint));
  assert.deepEqual(data.parts.map(part => [part.sku, part.validation_role]), [['110-38650', 'control'], ['400-90753', 'active_test'], ['110-40359', 'watch_candidate']]);
  assert.equal(data.parts[1].references[0].part_number, '110-96500');
  assert.equal(data.parts[1].references[0].relationship, null);
  assert.equal(data.parts[1].references[0].supersession_verified, false);
  for (const part of data.parts) {
    for (const key of ['machine_models', 'compatibility', 'stock', 'price', 'lead_time', 'oem_availability', 'compatible_availability', 'specifications']) assert.equal(part[key], null);
  }
  assert.deepEqual(readdirSync(resolve(root, 'juki')).sort(), skus.map(sku => sku + '.html').sort());
});

test('every page retains noindex,nofollow and local assets/links resolve', () => {
  const pages = ['index.html', 'privacy.html', 'thank-you.html', ...skus.map(sku => `juki/${sku}.html`)];
  for (const page of pages) {
    const path = resolve(root, page);
    const html = readFileSync(path, 'utf8');
    assert.match(html, /<meta name="robots" content="noindex,nofollow">/);
    assert.match(html, /<html lang="en">/);
    for (const [, target] of html.matchAll(/(?:href|src)="([^"]+)"/g)) {
      if (!/^(?:https?:|#)/.test(target)) assert.ok(existsSync(resolve(dirname(path), target)), `${page}: ${target}`);
    }
  }
});
