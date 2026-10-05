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
const liveBase = 'https://cho-leung.github.io/juki-search-test/';
const leadEndpoint = 'https://script.google.com/macros/s/AKfycbybM7TeEMbAdSPI6XdRSeYk_YtdnA1OqTZvBwwP1mi1qc-nS49oXhBdmaU_2dBl-hEDjg/exec';
const catalog = JSON.parse(readFileSync(resolve(root, 'data/parts.json'), 'utf8'));
const manufacturerSources = {
  '110-38650': 'https://www.juki.co.jp/industrial_j/download_j/manual_j/ddl8000a/ddl8000a/menu/8000a/partslist.pdf',
  '400-90753': 'https://www.juki.co.jp/industrial_j/download_j/manual_j/dln5410n/menu/dln5410n/pdf/partslist_dln5410n-7.pdf',
  '110-96500': 'https://www.juki.co.jp/industrial_j/download_j/manual_j/ddl8700/menu/ddl8700-7/pdf/partslist.pdf'
};
const indexablePages = new Map([
  ['index.html', liveBase],
  ['juki/400-90753.html', liveBase + 'juki/400-90753.html'],
  ['juki/110-38650.html', liveBase + 'juki/110-38650.html']
]);
const noindexPages = ['juki/110-40359.html', 'privacy.html', 'thank-you.html'];

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
  ['reference-only number as SKU', { sku: '110-96500' }],
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

test('catalog retains exactly three SKU roles and the exact existing lead endpoint', () => {
  const data = catalog;
  assert.equal(data.system_version, '0.2');
  assert.equal(data.lead_endpoint, leadEndpoint);
  assert.deepEqual(data.parts.map(part => [part.sku, part.validation_role]), [['110-38650', 'control'], ['400-90753', 'active_test'], ['110-40359', 'watch_candidate']]);
  for (const part of data.parts) {
    for (const key of ['compatibility', 'stock', 'price', 'lead_time', 'oem_availability', 'compatible_availability', 'specifications']) assert.equal(part[key], null);
  }
  assert.deepEqual(readdirSync(resolve(root, 'juki')).sort(), skus.map(sku => sku + '.html').sort());
});

for (const [sku, machine, description] of [
  ['400-90753', 'DLN-5410N-7', 'THREAD TRIMMER SOLENOID'],
  ['110-38650', 'DDL-8000A', 'HOOK ASM.']
]) {
  test(`${sku} has only the authorized manufacturer documentation reference`, () => {
    const part = catalog.parts.find(item => item.sku === sku);
    assert.deepEqual(part.machine_models, [machine]);
    assert.equal(part.part_type, description);
    assert.equal(part.manufacturer_documentation.machine_reference, machine);
    assert.equal(part.manufacturer_documentation.part_description, description);
    assert.equal(part.manufacturer_documentation.url, manufacturerSources[sku]);
    assert.match(part.identity_source, /documentation reference; exact-variant fit still requires verification/);
    assert.equal(part.compatibility, null);
    if (sku === '110-38650') assert.equal(part.name, 'Juki Hook Assembly');
  });
}

test('110-96500 is separately documented, reference-only, with all relationships unverified', () => {
  const references = catalog.parts.flatMap(part => part.references);
  assert.equal(references.length, 1);
  const reference = references[0];
  assert.equal(reference.part_number, '110-96500');
  assert.equal(reference.role, 'separately_documented_reference_only');
  assert.equal(reference.part_type, 'THREAD TRIMMER SOLENOID');
  assert.equal(reference.manufacturer_documentation.machine_reference, 'DDL-8700-7');
  assert.equal(reference.manufacturer_documentation.url, manufacturerSources['110-96500']);
  assert.equal(reference.relationship, null);
  assert.equal(reference.supersession_verified, false);
  assert.equal(reference.replacement_verified, false);
  assert.equal(reference.interchangeability_verified, false);
  for (const phrase of ['separately identifies 110-96500', 'no verified supersession', 'no verified replacement relationship', 'no verified interchangeability']) assert.ok(reference.note.includes(phrase));
  assert.doesNotMatch(reference.note, /historical|superseded/i);
  assert.ok(!catalog.parts.some(part => part.sku === '110-96500'));
  assert.ok(!existsSync(resolve(root, 'juki/110-96500.html')));
});

test('110-40359 remains unverified with no added manufacturer evidence', () => {
  const watch = catalog.parts.find(part => part.sku === '110-40359');
  assert.equal(watch.machine_models, null);
  assert.equal(watch.compatibility, null);
  assert.equal(watch.manufacturer_documentation, undefined);
  assert.deepEqual(watch.references, []);
  assert.match(watch.identity_source, /manufacturer documentation has not been verified/);
});

for (const [page, canonical] of indexablePages) {
  test(`${page} is indexable with its exact self-referencing canonical`, () => {
    const html = readFileSync(resolve(root, page), 'utf8');
    const robots = [...html.matchAll(/<meta\b[^>]*name="robots"[^>]*>/gi)];
    assert.ok(robots.every(([tag]) => !/noindex|nofollow/i.test(tag)));
    const canonicals = [...html.matchAll(/<link\b[^>]*rel="canonical"[^>]*>/gi)];
    assert.equal(canonicals.length, 1);
    assert.ok(canonicals[0][0].includes(`href="${canonical}"`));
  });
}

for (const page of noindexPages) {
  test(`${page} keeps noindex,nofollow and has no canonical`, () => {
    const html = readFileSync(resolve(root, page), 'utf8');
    assert.match(html, /<meta name="robots" content="noindex,nofollow">/);
    assert.doesNotMatch(html, /rel="canonical"/i);
  });
}

test('the sitemap contains exactly the three authorized URLs', () => {
  const sitemap = readFileSync(resolve(root, 'sitemap.xml'), 'utf8');
  assert.match(sitemap, /<urlset xmlns="http:\/\/www\.sitemaps\.org\/schemas\/sitemap\/0\.9">/);
  const urls = [...sitemap.matchAll(/<loc>([^<]+)<\/loc>/g)].map(([, url]) => url);
  assert.equal((sitemap.match(/<url>/g) || []).length, 3);
  assert.deepEqual(urls, [...indexablePages.values()]);
  assert.doesNotMatch(sitemap, /110-40359|privacy\.html|thank-you\.html|110-96500/);
});

test('robots.txt allows crawling and points to the authorized sitemap', () => {
  const robots = readFileSync(resolve(root, 'robots.txt'), 'utf8');
  assert.equal(robots, `User-agent: *\nAllow: /\n\nSitemap: ${liveBase}sitemap.xml\n`);
  assert.doesNotMatch(robots, /Disallow:/i);
});

test('all six pages have valid local assets and links, with no extra HTML pages', () => {
  const pages = [...indexablePages.keys(), ...noindexPages];
  assert.deepEqual(readdirSync(root).filter(name => name.endsWith('.html')).sort(), ['index.html', 'privacy.html', 'thank-you.html']);
  for (const page of pages) {
    const path = resolve(root, page);
    const html = readFileSync(path, 'utf8');
    assert.match(html, /<html lang="en">/);
    for (const [, target] of html.matchAll(/(?:href|src)="([^"]+)"/g)) {
      if (!/^(?:https?:|#)/.test(target)) assert.ok(existsSync(resolve(dirname(path), target)), `${page}: ${target}`);
    }
  }
});
