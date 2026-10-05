/**
 * Validation System v0.2 — deploy as a standalone Apps Script web app.
 * Script properties required: SITE_BASE_URL (HTTPS directory URL), NOTIFICATION_EMAIL.
 * Use the deployment owner's identity; never expose credentials in the static site.
 */
const SPREADSHEET_ID = '16IcDQZ9omcNJaIzrkCXfK53N--qwcxhR5QQuS5P1-Ts';
const SHEET_NAME = 'Leads';
const LEAD_HEADERS = [
  'timestamp', 'sku', 'machine_model', 'part_number', 'email', 'preference',
  'notes', 'page_url', 'lead_status', 'owner_minutes', 'quoted', 'paid',
  'estimated_contribution', 'actual_contribution'
];
const ALLOWED_SKUS = ['110-38650', '400-90753', '110-40359'];
const FIELD_LIMITS = {
  sku: 20, machine_model: 120, part_number: 80, email: 254,
  preference: 20, notes: 2000, page_url: 1000
};

function doGet() {
  return htmlResponse_('Fit request endpoint',
    '<p>Submit a fit request from an ExactPart SKU page. Visiting this endpoint does not record a lead.</p>');
}

function doPost(e) {
  let stage = 'configuration';
  try {
    const config = getConfig_();
    stage = 'validation';
    const lead = validateLead_(e, config.siteBase);
    stage = 'ledger';
    const saved = appendLead_(lead);
    stage = 'notification';
    try {
      GmailApp.sendEmail(config.notificationEmail,
        '[ExactPart] Fit request: ' + lead.sku,
        notificationBody_(lead), { replyTo: lead.email, name: 'ExactPart' });
    } catch (_) {
      // The lead is already durable. Returning failure would invite duplicate submissions.
      // No visitor data or exception text is logged or reflected in the response.
      console.error('Lead saved; Gmail notification failed. Check the Leads ledger.');
      try {
        saved.sheet.getRange(saved.row, 9).setValue('new_notification_failed');
        SpreadsheetApp.flush();
      } catch (_) {
        console.error('Could not mark notification failure. Check the Leads ledger.');
      }
    }
    return successResponse_(config.siteBase + 'thank-you.html');
  } catch (_) {
    console.error('Fit request could not be confirmed at stage: ' + stage + '.');
    return htmlResponse_('Request not confirmed',
      '<p>We could not confirm receipt of your request. Please use your browser’s Back button to review the required fields, then try again later.</p>' +
      '<p>No fit, stock or order has been confirmed. If you already received an operator reply, check it before resubmitting.</p>');
  }
}

function getConfig_() {
  const properties = PropertiesService.getScriptProperties();
  const site = (properties.getProperty('SITE_BASE_URL') || '').trim();
  // Restrict the server-owned redirect to an HTTPS host and simple directory path.
  // No credentials, query, fragment, port or encoded path is accepted here.
  const match = /^https:\/\/([a-z0-9]+(?:[.-][a-z0-9]+)*)(\/[a-z0-9_~./-]*)?$/i.exec(site);
  if (!match) throw new Error('Invalid site configuration');
  const path = match[2] || '/';
  if (path.split('/').some(function (segment) { return segment === '.' || segment === '..'; })) {
    throw new Error('Invalid site configuration');
  }
  const siteBase = 'https://' + match[1].toLowerCase() + path.replace(/\/*$/, '/');
  const notificationEmail = (properties.getProperty('NOTIFICATION_EMAIL') || '').trim();
  if (!validEmail_(notificationEmail)) throw new Error('Invalid notification configuration');
  return { siteBase: siteBase, notificationEmail: notificationEmail };
}

function validateLead_(e, siteBase) {
  if (!e || !e.postData || typeof e.postData.contents !== 'string' ||
      e.postData.contents.length > 32768 ||
      !/^application\/x-www-form-urlencoded(?:;|$)/i.test(e.postData.type || '')) {
    throw new Error('Invalid request');
  }
  const parameters = e.parameters;
  if (!parameters || typeof parameters !== 'object' || Array.isArray(parameters)) throw new Error('Invalid fields');
  const keys = Object.keys(parameters);
  if (keys.some(function (key) { return !Object.prototype.hasOwnProperty.call(FIELD_LIMITS, key); })) {
    throw new Error('Unexpected fields');
  }
  const lead = {};
  Object.keys(FIELD_LIMITS).forEach(function (key) {
    const values = parameters[key];
    if (values === undefined && key === 'notes') { lead.notes = ''; return; }
    if (!Array.isArray(values) || values.length !== 1 || typeof values[0] !== 'string' || values[0].length > FIELD_LIMITS[key]) {
      throw new Error('Invalid field');
    }
    const raw = values[0];
    if (key === 'email' && /[\r\n]/.test(raw)) throw new Error('Invalid email');
    let value = raw.replace(/\r\n?/g, '\n').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, '');
    if (key !== 'notes') value = value.replace(/[\n\t]/g, ' ');
    value = value.trim();
    if (key !== 'notes' && !value) throw new Error('Required field');
    lead[key] = value;
  });
  if (ALLOWED_SKUS.indexOf(lead.sku) === -1 ||
      ['oem', 'compatible', 'either'].indexOf(lead.preference) === -1 ||
      !validEmail_(lead.email)) throw new Error('Invalid lead');
  // Client page_url is attribution metadata, not proof of the request's origin.
  // Strip query/fragment data; never accept a client-supplied redirect destination.
  const canonicalPage = siteBase + 'juki/' + lead.sku + '.html';
  if (lead.page_url.split(/[?#]/)[0] !== canonicalPage) throw new Error('Invalid page');
  lead.page_url = canonicalPage;
  return lead;
}

function validEmail_(value) {
  if (typeof value !== 'string' || value.length > 254 ||
      !/^[^\s@<>(),;:"\\]+@[a-z0-9.-]+$/i.test(value)) return false;
  const pieces = value.split('@');
  const local = pieces[0];
  const labels = pieces[1].split('.');
  return local.length <= 64 && !/^\.|\.$|\.\./.test(local) && labels.length >= 2 &&
    /^[a-z]{2,63}$/i.test(labels[labels.length - 1]) &&
    labels.every(function (label) { return /^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$/i.test(label); });
}

function sheetText_(value) {
  // Prevent user text from becoming a spreadsheet formula, including on CSV export.
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function appendLead_(lead) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('Ledger busy');
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Missing ledger tab');
    if (sheet.getLastRow() === 0) sheet.appendRow(LEAD_HEADERS.slice());
    const headers = sheet.getRange(1, 1, 1, LEAD_HEADERS.length).getDisplayValues()[0];
    if (sheet.getLastColumn() !== LEAD_HEADERS.length ||
        headers.some(function (value, index) { return value !== LEAD_HEADERS[index]; })) {
      throw new Error('Ledger schema mismatch');
    }
    const row = [new Date(), lead.sku, lead.machine_model, lead.part_number,
      lead.email, lead.preference, lead.notes, lead.page_url, 'new',
      '', '', '', '', ''];
    sheet.appendRow(row.map(function (value) { return typeof value === 'string' ? sheetText_(value) : value; }));
    SpreadsheetApp.flush();
    return { sheet: sheet, row: sheet.getLastRow() };
  } finally {
    lock.releaseLock();
  }
}

function notificationBody_(lead) {
  return [
    'New fit / availability inquiry. This is not an order or a confirmed fit.',
    '', 'SKU: ' + lead.sku, 'Machine model: ' + lead.machine_model,
    'Existing part number: ' + lead.part_number, 'Email: ' + lead.email,
    'OEM / compatible preference: ' + lead.preference,
    'Notes: ' + (lead.notes || '(none)'), 'Page: ' + lead.page_url,
    '', 'Review the Leads ledger. Owner minutes, quoted, paid and contribution fields require operator entry.'
  ].join('\n');
}

function escapeHtml_(value) {
  return value.replace(/[&<>"']/g, function (character) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character];
  });
}

function successResponse_(thankYouUrl) {
  // HtmlService is sandboxed. Redirect the content frame; a user-activated _top
  // link is the supported fallback for full-page navigation. No CORS assumptions.
  const href = escapeHtml_(thankYouUrl);
  const scriptUrl = JSON.stringify(thankYouUrl).replace(/</g, '\\u003c');
  return htmlResponse_('Request recorded',
    '<p>Your request was recorded for review. Fit and availability remain unconfirmed.</p>' +
    '<p><a href="' + href + '" target="_top">Continue to the thank-you page</a></p>' +
    '<script>window.location.replace(' + scriptUrl + ');</script>');
}

function htmlResponse_(title, body) {
  return HtmlService.createHtmlOutput(
    '<!doctype html><html lang="en"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<meta name="robots" content="noindex,nofollow"><meta name="referrer" content="no-referrer">' +
    '<title>' + escapeHtml_(title) + '</title></head><body><main><h1>' +
    escapeHtml_(title) + '</h1>' + body + '</main></body></html>').setTitle(title);
}
