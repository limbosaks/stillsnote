/**
 * Stillsnote backend (Google Apps Script)
 *
 * Handles: membership (sign-up requests, approvals, fixed access codes, sign-in),
 * every AI call (photo reads, captions, ideas, direction), saved profiles and drafts,
 * and a daily trends refresh. The website never holds the Claude API key.
 *
 * One-time setup: see the Stillsnote setup guide.
 *   1. Paste this file, set the time zone to Seoul, run setup().
 *   2. Add the Script Property ANTHROPIC_API_KEY, run checkSetup().
 *   3. Deploy as a web app (Execute as: Me, Who has access: Anyone).
 */

const CFG = {
  SITE_NAME: 'Stillsnote',
  SITE_URL: 'https://stillsnote.com',
  FROM_EMAIL: 'admin@stillsnote.com',
  MODEL: 'claude-sonnet-5-5',
  TOKEN_DAYS: 30,
  OPTIN_WORDING: 'Yes, send me occasional Stillsnote news and tips. I can unsubscribe anytime.',
  MAX_THUMB_CHARS: 45000,
  MAX_IMAGE_CHARS: 6000000
};

const TABS = {
  SETTINGS: 'Settings',
  MEMBERS: 'Members',
  PROFILES: 'Profiles',
  DRAFTS: 'Drafts',
  TRENDS: 'Trends'
};

const MEMBER_HEADERS = ['Requested', 'Name', 'Email', 'Instagram', 'Status', 'Access code', 'Plan',
  'Marketing opt-in', 'Opt-in wording', 'AI uses this month', 'Usage month', 'Last sign-in', 'Note'];
const PROFILE_HEADERS = ['Email', 'Updated', 'Profile JSON'];
const DRAFT_HEADERS = ['ID', 'Email', 'Updated', 'Planned date', 'Status', 'Pillar', 'Caption', 'Hashtags', 'Story', 'Thumbnail', 'Extra JSON'];
const TREND_HEADERS = ['Updated', 'Region', 'Trends JSON'];

const DEFAULT_SETTINGS = [
  ['Sign-up mode', 'Approve', 'Approve = you approve each request. Automatic = codes are emailed instantly.'],
  ['Admin email', 'admin@stillsnote.com', 'Where new sign-up requests are sent, and where member replies go.'],
  ['Free monthly limit', 30, 'AI uses per month for Free members. 0 = unlimited.'],
  ['Paid monthly limit', 300, 'AI uses per month for Paid members. 0 = unlimited.'],
  ['Trend regions', 'Korea, Global', 'Regions refreshed each morning, comma-separated.'],
  ['Niches', 'travel, food, cafe, fitness, fashion, beauty, lifestyle, photography, pets, outdoors', 'Niche filters on the Trends page.']
];

/* ------------------------------------------------------------------ */
/* Setup                                                              */
/* ------------------------------------------------------------------ */

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const existing = props.getProperty('SHEET_ID');
  if (existing) {
    ss = SpreadsheetApp.openById(existing);
  } else {
    ss = SpreadsheetApp.create('Stillsnote members');
    props.setProperty('SHEET_ID', ss.getId());
  }
  if (!props.getProperty('SECRET')) {
    props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  }

  ensureTab_(ss, TABS.SETTINGS, ['Setting', 'Value', 'What it does']);
  const settings = ss.getSheetByName(TABS.SETTINGS);
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
  }
  ensureTab_(ss, TABS.MEMBERS, MEMBER_HEADERS);
  ensureTab_(ss, TABS.PROFILES, PROFILE_HEADERS);
  ensureTab_(ss, TABS.DRAFTS, DRAFT_HEADERS);
  ensureTab_(ss, TABS.TRENDS, TREND_HEADERS);
  const blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);

  const members = ss.getSheetByName(TABS.MEMBERS);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Declined', 'Removed'], true).build();
  members.getRange(2, 5, 999, 1).setDataValidation(statusRule);
  const planRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Free', 'Paid'], true).build();
  members.getRange(2, 7, 999, 1).setDataValidation(planRule);
  const modeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Approve', 'Automatic'], true).build();
  settings.getRange(2, 2).setDataValidation(modeRule);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'onSheetEdit' || fn === 'refreshTrends') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('refreshTrends').timeBased().everyDays(1).atHour(6).create();

  Logger.log('Ready. Your Stillsnote Sheet: ' + ss.getUrl());
  Logger.log('Next: add the Script Property ANTHROPIC_API_KEY, then run checkSetup.');
}

function checkSetup() {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) { Logger.log('Missing: add ANTHROPIC_API_KEY under Project Settings > Script Properties.'); return; }
  const out = callClaude_({
    model: CFG.MODEL,
    max_tokens: 20,
    messages: [{ role: 'user', content: 'Reply with the single word: ready' }]
  });
  Logger.log('Claude says: ' + textOf_(out));
  Logger.log('Sheet: ' + sheet_().getUrl());
  Logger.log('Time zone: ' + Session.getScriptTimeZone() + ' (should be Asia/Seoul)');
}

function ensureTab_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    headers.forEach(function (h) {
      if (current.indexOf(h) === -1) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(h).setFontWeight('bold');
      }
    });
  }
  return sh;
}

/* ------------------------------------------------------------------ */
/* Web entry points                                                    */
/* ------------------------------------------------------------------ */

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const handler = ACTIONS[body.action];
    if (!handler) throw userError_('Unknown action.');
    const result = handler(body) || {};
    result.ok = true;
    return json_(result);
  } catch (err) {
    const out = { ok: false, error: err.userMessage || 'Something went wrong. Please try again.' };
    if (err.code) out.code = err.code;
    if (!err.userMessage) console.error(body.action, err && err.stack || err);
    return json_(out);
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'review') return reviewPage_(p.e, p.s);
  if (p.action === 'unsub') return unsubscribe_(p.e, p.s);
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Stillsnote service is running.</p>');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function userError_(message, code) {
  const err = new Error(message);
  err.userMessage = message;
  if (code) err.code = code;
  return err;
}

const ACTIONS = {
  signup: actionSignup_,
  login: actionLogin_,
  me: function (b) { const m = requireMember_(b.token); return { member: publicMember_(m) }; },
  generate: actionGenerate_,
  getProfile: function (b) { const m = requireMember_(b.token); return { profile: getProfile_(m.email) }; },
  saveProfile: actionSaveProfile_,
  listDrafts: actionListDrafts_,
  saveDraft: actionSaveDraft_,
  deleteDraft: actionDeleteDraft_,
  trends: actionTrends_,
  ideas: actionIdeas_,
  direction: actionDirection_
};

/* ------------------------------------------------------------------ */
/* Sheet helpers                                                       */
/* ------------------------------------------------------------------ */

function sheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Run setup() first.');
  return SpreadsheetApp.openById(id);
}

function tab_(name) { return sheet_().getSheetByName(name); }

function settings_() {
  const rows = tab_(TABS.SETTINGS).getDataRange().getValues();
  const s = {};
  rows.slice(1).forEach(function (r) { s[String(r[0]).trim()] = r[1]; });
  return {
    mode: String(s['Sign-up mode'] || 'Approve').trim(),
    admin: String(s['Admin email'] || '').trim() || Session.getEffectiveUser().getEmail(),
    freeLimit: Number(s['Free monthly limit'] || 0),
    paidLimit: Number(s['Paid monthly limit'] || 0),
    regions: String(s['Trend regions'] || 'Korea, Global').split(',').map(trim_).filter(Boolean),
    niches: String(s['Niches'] || '').split(',').map(trim_).filter(Boolean)
  };
}

function trim_(x) { return String(x).trim(); }
function normEmail_(x) { return String(x || '').trim().toLowerCase(); }
function now_() { return new Date(); }
function monthKey_() { return Utilities.formatDate(now_(), Session.getScriptTimeZone(), 'yyyy-MM'); }
function stamp_() { return Utilities.formatDate(now_(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }

function colMap_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function (h, i) { map[h] = i; });
  return map;
}

function findMember_(email) {
  const sh = tab_(TABS.MEMBERS);
  const values = sh.getDataRange().getValues();
  const c = colMap_(sh);
  const target = normEmail_(email);
  for (let i = 1; i < values.length; i++) {
    if (normEmail_(values[i][c['Email']]) === target) {
      return rowToMember_(values[i], c, i + 1);
    }
  }
  return null;
}

function rowToMember_(r, c, rowNum) {
  return {
    row: rowNum,
    name: String(r[c['Name']] || ''),
    email: normEmail_(r[c['Email']]),
    instagram: String(r[c['Instagram']] || ''),
    status: String(r[c['Status']] || '').trim(),
    code: String(r[c['Access code']] || '').trim(),
    plan: String(r[c['Plan']] || 'Free').trim() || 'Free',
    uses: Number(r[c['AI uses this month']] || 0),
    usageMonth: String(r[c['Usage month']] || '')
  };
}

function setMemberCell_(rowNum, header, value) {
  const sh = tab_(TABS.MEMBERS);
  const c = colMap_(sh);
  sh.getRange(rowNum, c[header] + 1).setValue(value);
}

function publicMember_(m) {
  const st = settings_();
  const limit = m.plan === 'Paid' ? st.paidLimit : st.freeLimit;
  const uses = m.usageMonth === monthKey_() ? m.uses : 0;
  return { name: m.name, email: m.email, instagram: m.instagram, plan: m.plan, uses: uses, limit: limit };
}

/* ------------------------------------------------------------------ */
/* Security: signatures, tokens, codes                                 */
/* ------------------------------------------------------------------ */

function sign_(text) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  const bytes = Utilities.computeHmacSha256Signature(text, secret);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function makeToken_(email) {
  const exp = Date.now() + CFG.TOKEN_DAYS * 86400000;
  const payload = email + '|' + exp;
  return Utilities.base64EncodeWebSafe(payload).replace(/=+$/, '') + '.' + sign_(payload);
}

function readToken_(token) {
  if (!token || String(token).indexOf('.') < 0) return null;
  const parts = String(token).split('.');
  let payload;
  try {
    payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(padB64_(parts[0]))).getDataAsString();
  } catch (e) { return null; }
  if (sign_(payload) !== parts[1]) return null;
  const bits = payload.split('|');
  if (Number(bits[1]) < Date.now()) return null;
  return bits[0];
}

function padB64_(s) { while (s.length % 4) s += '='; return s; }

function requireMember_(token) {
  const email = readToken_(token);
  if (!email) throw userError_('Please sign in again.', 'signin');
  const m = findMember_(email);
  if (!m || m.status !== 'Approved') throw userError_('Your membership is not active. Contact us if this is a mistake.', 'signin');
  return m;
}

function newCode_() {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 8; i++) s += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  return 'SN-' + s.slice(0, 4) + '-' + s.slice(4);
}

function cleanCode_(x) { return String(x || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }

/* ------------------------------------------------------------------ */
/* Membership                                                          */
/* ------------------------------------------------------------------ */

function actionSignup_(b) {
  if (b.website) return { status: 'pending' }; // honeypot: bots fill hidden field
  const email = normEmail_(b.email);
  const name = String(b.name || '').trim().slice(0, 80);
  const instagram = String(b.instagram || '').trim().replace(/^@/, '').slice(0, 60);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw userError_('Please enter a valid email address.');
  if (!name) throw userError_('Please enter your name.');

  const cache = CacheService.getScriptCache();
  const throttleKey = 'su:' + email;
  if (cache.get(throttleKey)) throw userError_('We just got a request from this email. Check your inbox, or try again in 10 minutes.');
  cache.put(throttleKey, '1', 600);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const existing = findMember_(email);
    if (existing) {
      if (existing.status === 'Approved') {
        if (!existing.code) { existing.code = newCode_(); setMemberCell_(existing.row, 'Access code', existing.code); }
        sendCodeEmail_(existing, true);
        return { status: 'resent' };
      }
      return { status: 'pending' };
    }

    const st = settings_();
    const auto = st.mode.toLowerCase() === 'automatic';
    const optin = b.optin === true;
    const code = auto ? newCode_() : '';
    const sh = tab_(TABS.MEMBERS);
    const c = colMap_(sh);
    const row = new Array(sh.getLastColumn()).fill('');
    row[c['Requested']] = stamp_();
    row[c['Name']] = name;
    row[c['Email']] = email;
    row[c['Instagram']] = instagram ? '@' + instagram : '';
    row[c['Status']] = auto ? 'Approved' : 'Pending';
    row[c['Access code']] = code;
    row[c['Plan']] = 'Free';
    row[c['Marketing opt-in']] = optin ? 'Yes (' + stamp_() + ')' : 'No';
    row[c['Opt-in wording']] = optin ? CFG.OPTIN_WORDING : '';
    row[c['AI uses this month']] = 0;
    row[c['Usage month']] = monthKey_();
    sh.appendRow(row);

    const member = { name: name, email: email, instagram: instagram, code: code };
    if (auto) {
      sendCodeEmail_(member, false);
      notifyAdmin_(member, st, true);
      return { status: 'approved' };
    }
    sendReceivedEmail_(member);
    notifyAdmin_(member, st, false);
    return { status: 'pending' };
  } finally {
    lock.releaseLock();
  }
}

function actionLogin_(b) {
  const email = normEmail_(b.email);
  const cache = CacheService.getScriptCache();
  const key = 'li:' + email;
  const tries = Number(cache.get(key) || 0);
  if (tries >= 8) throw userError_('Too many attempts. Please wait 15 minutes and try again.');

  const m = findMember_(email);
  const ok = m && m.status === 'Approved' && m.code && cleanCode_(m.code) === cleanCode_(b.code);
  if (!ok) {
    cache.put(key, String(tries + 1), 900);
    if (m && m.status === 'Pending') throw userError_('Your request is still waiting for approval. You will get your code by email.');
    throw userError_('That email and code do not match. Check the code in your Stillsnote email.');
  }
  cache.remove(key);
  setMemberCell_(m.row, 'Last sign-in', stamp_());
  return { token: makeToken_(m.email), member: publicMember_(m) };
}

/** Installable trigger: approving someone in the Sheet emails their code. */
function onSheetEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== TABS.MEMBERS || e.range.getRow() < 2) return;
    const c = colMap_(sh);
    const statusCol = c['Status'] + 1;
    const first = e.range.getColumn();
    const last = first + e.range.getNumColumns() - 1;
    if (statusCol < first || statusCol > last) return;
    for (let r = e.range.getRow(); r < e.range.getRow() + e.range.getNumRows(); r++) {
      const values = sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0];
      const m = rowToMember_(values, c, r);
      if (m.status !== 'Approved' || !m.email) continue;
      if (e.oldValue === 'Approved') continue;
      if (!m.code) { m.code = newCode_(); sh.getRange(r, c['Access code'] + 1).setValue(m.code); }
      if (!values[c['Plan']]) sh.getRange(r, c['Plan'] + 1).setValue('Free');
      if (!values[c['Requested']]) sh.getRange(r, c['Requested'] + 1).setValue(stamp_());
      sendCodeEmail_(m, false);
    }
  } catch (err) {
    console.error('onSheetEdit', err);
  }
}

function reviewPage_(email, sig) {
  email = normEmail_(email);
  if (!email || sign_('review|' + email) !== sig) {
    return HtmlService.createHtmlOutput(pageShell_('<p>This link is not valid.</p>'));
  }
  const m = findMember_(email);
  if (!m) return HtmlService.createHtmlOutput(pageShell_('<p>This request no longer exists.</p>'));
  const body =
    '<h1>Sign-up request</h1>' +
    '<p><b>' + esc_(m.name) + '</b><br>' + esc_(m.email) + (m.instagram ? '<br>' + esc_(m.instagram) : '') + '</p>' +
    '<p>Current status: <b>' + esc_(m.status || 'Pending') + '</b></p>' +
    '<p><button onclick="decide(\'Approved\')">Approve and send the code</button> ' +
    '<button class="ghost" onclick="decide(\'Declined\')">Decline</button></p>' +
    '<p id="msg"></p>' +
    '<script>function decide(d){document.getElementById("msg").textContent="Working…";' +
    'google.script.run.withSuccessHandler(function(t){document.getElementById("msg").textContent=t;})' +
    '.withFailureHandler(function(){document.getElementById("msg").textContent="That did not work. Try again.";})' +
    '.reviewDecision(' + JSON.stringify(email) + ',' + JSON.stringify(sig) + ',d);}</script>';
  return HtmlService.createHtmlOutput(pageShell_(body)).setTitle('Stillsnote sign-up');
}

/** Called from the review page after the confirm click (link scanners cannot trigger it). */
function reviewDecision(email, sig, decision) {
  email = normEmail_(email);
  if (sign_('review|' + email) !== sig) return 'This link is not valid.';
  if (decision !== 'Approved' && decision !== 'Declined') return 'Unknown choice.';
  const m = findMember_(email);
  if (!m) return 'This request no longer exists.';
  if (decision === 'Declined') {
    setMemberCell_(m.row, 'Status', 'Declined');
    return 'Declined. No email was sent.';
  }
  if (m.status === 'Approved') return 'Already approved.';
  if (!m.code) { m.code = newCode_(); setMemberCell_(m.row, 'Access code', m.code); }
  setMemberCell_(m.row, 'Status', 'Approved');
  sendCodeEmail_(m, false);
  return 'Approved. ' + m.name + ' has been emailed their code.';
}

function unsubscribe_(email, sig) {
  email = normEmail_(email);
  if (!email || sign_('unsub|' + email) !== sig) {
    return HtmlService.createHtmlOutput(pageShell_('<p>This link is not valid.</p>'));
  }
  const m = findMember_(email);
  if (m) setMemberCell_(m.row, 'Marketing opt-in', 'No (unsubscribed ' + stamp_() + ')');
  return HtmlService.createHtmlOutput(pageShell_('<h1>Unsubscribed</h1><p>You will not get Stillsnote news emails. Account emails, like your access code, still arrive.</p>'));
}

function pageShell_(inner) {
  return '<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:Helvetica,Arial,sans-serif;background:#35302C;color:#EFE8DE;padding:40px 24px;max-width:520px;margin:0 auto;line-height:1.5}' +
    'h1{font-family:Georgia,serif;font-weight:400;font-size:32px}button{font:inherit;padding:12px 20px;border:0;background:#E8A066;color:#2A2420;cursor:pointer;margin:4px 4px 4px 0}' +
    'button.ghost{background:transparent;color:#EFE8DE;border:1px solid #6E645B}</style></head><body>' +
    '<p style="letter-spacing:6px;font-size:12px">STILLSNOTE</p>' + inner + '</body></html>';
}

function esc_(s) {
  return String(s || '').replace(/[&<>"']/g, function (ch) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch];
  });
}

/* ------------------------------------------------------------------ */
/* Email                                                               */
/* ------------------------------------------------------------------ */

function serviceUrl_() { return ScriptApp.getService().getUrl(); }

function emailShell_(inner, email) {
  const unsub = serviceUrl_() + '?action=unsub&e=' + encodeURIComponent(email) + '&s=' + sign_('unsub|' + email);
  return '<div style="font-family:Helvetica,Arial,sans-serif;max-width:520px;color:#2A2420;line-height:1.55">' +
    '<p style="letter-spacing:6px;font-size:12px;color:#8A7F74">STILLSNOTE</p>' + inner +
    '<p style="font-size:12px;color:#8A7F74;margin-top:32px">You are getting this because you asked for a Stillsnote account. ' +
    '<a href="' + unsub + '" style="color:#8A7F74">Unsubscribe from news emails</a></p></div>';
}

/** Sends from admin@stillsnote.com when the script's Google account can send as it; otherwise from the account itself with replies going to admin@. */
function send_(to, subject, html) {
  const opts = { htmlBody: html, name: CFG.SITE_NAME, replyTo: CFG.FROM_EMAIL };
  if (canSendAsFrom_()) {
    opts.from = CFG.FROM_EMAIL;
    GmailApp.sendEmail(to, subject, 'Open this email in an app that shows HTML.', opts);
  } else {
    MailApp.sendEmail({ to: to, subject: subject, htmlBody: html, name: CFG.SITE_NAME, replyTo: CFG.FROM_EMAIL });
  }
}

function canSendAsFrom_() {
  const me = Session.getEffectiveUser().getEmail().toLowerCase();
  if (me === CFG.FROM_EMAIL) return false; // already sending as admin@, no alias needed
  return GmailApp.getAliases().map(function (a) { return a.toLowerCase(); }).indexOf(CFG.FROM_EMAIL) >= 0;
}

/** Run this once to check which address emails will come from. */
function checkSender() {
  const me = Session.getEffectiveUser().getEmail().toLowerCase();
  if (me === CFG.FROM_EMAIL || canSendAsFrom_()) {
    Logger.log('Good: emails will come from ' + CFG.FROM_EMAIL + '.');
  } else {
    Logger.log('Emails will come from ' + me + ', with replies going to ' + CFG.FROM_EMAIL + '. ' +
      'To send as ' + CFG.FROM_EMAIL + ', add it in Gmail: Settings > See all settings > Accounts > Send mail as, then run checkSender again.');
  }
}

function sendCodeEmail_(m, resend) {
  const html = emailShell_(
    '<p>Hi ' + esc_(m.name || 'there') + ',</p>' +
    '<p>' + (resend ? 'Here is your Stillsnote access code again.' : 'You are in. Here is your Stillsnote access code.') + '</p>' +
    '<p style="font-size:26px;letter-spacing:3px;font-family:Menlo,monospace;background:#F1ECE4;padding:14px 18px;display:inline-block">' + esc_(m.code) + '</p>' +
    '<p>Sign in at <a href="' + CFG.SITE_URL + '">' + CFG.SITE_URL.replace('https://', '') + '</a> with this email address and the code. ' +
    'The code does not change, so keep this email. Lost it? Request access again with the same email and we will resend it.</p>', m.email);
  send_(m.email, resend ? 'Your Stillsnote code' : 'Welcome to Stillsnote: your access code', html);
}

function sendReceivedEmail_(m) {
  const html = emailShell_('<p>Hi ' + esc_(m.name) + ',</p><p>We got your request for Stillsnote. ' +
    'You will get your access code by email as soon as it is approved.</p>', m.email);
  send_(m.email, 'We got your Stillsnote request', html);
}

function notifyAdmin_(m, st, auto) {
  const review = serviceUrl_() + '?action=review&e=' + encodeURIComponent(m.email) + '&s=' + sign_('review|' + m.email);
  const html = '<div style="font-family:Helvetica,Arial,sans-serif;line-height:1.55">' +
    '<p><b>' + esc_(m.name) + '</b> &lt;' + esc_(m.email) + '&gt;' + (m.instagram ? ' · @' + esc_(m.instagram) : '') + '</p>' +
    (auto ? '<p>Automatic mode is on, so their code has already been sent.</p>'
      : '<p><a href="' + review + '" style="display:inline-block;padding:12px 18px;background:#2A2420;color:#fff;text-decoration:none">Review: approve or decline</a></p>' +
      '<p style="color:#777;font-size:12px">You can also set their Status to Approved in the Members tab.</p>') +
    '</div>';
  send_(st.admin, (auto ? 'New Stillsnote member: ' : 'New Stillsnote request: ') + m.name, html);
}

/* ------------------------------------------------------------------ */
/* Usage limits                                                        */
/* ------------------------------------------------------------------ */

function useCredit_(m) {
  const st = settings_();
  const limit = m.plan === 'Paid' ? st.paidLimit : st.freeLimit;
  const month = monthKey_();
  const uses = m.usageMonth === month ? m.uses : 0;
  if (limit > 0 && uses >= limit) {
    throw userError_('You have used all ' + limit + ' AI requests for this month. They reset on the 1st.', 'limit');
  }
  setMemberCell_(m.row, 'AI uses this month', uses + 1);
  if (m.usageMonth !== month) setMemberCell_(m.row, 'Usage month', month);
  m.uses = uses + 1;
  m.usageMonth = month;
}

/** Charge one AI use, run fn, and give the use back if the AI call fails. */
function withCredit_(m, fn) {
  useCredit_(m);
  try {
    return fn();
  } catch (err) {
    try { setMemberCell_(m.row, 'AI uses this month', Math.max(0, m.uses - 1)); m.uses = Math.max(0, m.uses - 1); } catch (e) { /* ignore */ }
    throw err;
  }
}

/* ------------------------------------------------------------------ */
/* Claude                                                              */
/* ------------------------------------------------------------------ */

function callClaude_(payload) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('ANTHROPIC_API_KEY is not set.');
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code !== 200) {
    console.error('Claude API ' + code + ': ' + text.slice(0, 800));
    if (code === 429 || code === 529) throw userError_('The AI is busy right now. Please try again in a minute.');
    throw new Error('Claude API error ' + code);
  }
  return JSON.parse(text);
}

function textOf_(resp) {
  return (resp.content || []).filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; }).join('\n').trim();
}

/** Ask Claude for structured output by forcing a single tool call. */
function claudeStructured_(system, content, tool, maxTokens) {
  const resp = callClaude_({
    model: CFG.MODEL,
    max_tokens: maxTokens || 2000,
    system: system,
    tools: [tool],
    tool_choice: { type: 'tool', name: tool.name },
    messages: [{ role: 'user', content: content }]
  });
  const block = (resp.content || []).filter(function (b) { return b.type === 'tool_use'; })[0];
  if (!block) throw new Error('No structured output returned.');
  return block.input;
}

const VOICE_RULES = [
  'Write like a real person posting to their own account, not like a brand or a copywriter.',
  'Avoid filler and cliches: no "unforgettable", "embrace", "journey", "vibes", "elevate", "core memory", "living my best life", "not gonna lie", "a moment".',
  'No em dashes. No hashtags inside the caption text. Emoji only if the style profile uses them, and at most one or two.',
  'The first line is the hook: it must make sense and pull people in before Instagram cuts the caption off (about 125 characters).',
  'Use concrete details from the photo and the story. Never invent facts, names or places that are not in the photo, the story or the profile.',
  'Keep every caption under 2,200 characters.'
].join('\n');

function profileText_(p) {
  if (!p || !Object.keys(p).length) return 'No style profile saved yet. Use a natural, warm, casual voice.';
  const lines = [];
  if (p.niche) lines.push('Account niche: ' + p.niche);
  if (p.audience) lines.push('Audience: ' + p.audience);
  if (p.tone) lines.push('Tone words: ' + p.tone);
  if (p.language) lines.push('Caption language: ' + p.language);
  if (p.location) lines.push('Based in: ' + p.location);
  if (p.pillars && p.pillars.length) lines.push('Content pillars: ' + [].concat(p.pillars).join(', '));
  if (p.avoid) lines.push('Avoid: ' + p.avoid);
  if (p.samples) lines.push('Captions they wrote and liked (match this voice, do not copy):\n' + String(p.samples).slice(0, 3000));
  return lines.join('\n');
}

function trendHint_() {
  const t = latestTrends_();
  const tags = [];
  Object.keys(t.regions || {}).forEach(function (r) {
    (t.regions[r].hashtags || []).slice(0, 12).forEach(function (h) { tags.push(h.tag); });
  });
  return tags.length ? 'Hashtags currently rising (use only if they truly fit this photo): ' + tags.join(' ') : '';
}

/* ------------------------------------------------------------------ */
/* Studio: photo read, captions, hashtags                              */
/* ------------------------------------------------------------------ */

const POST_TOOL = {
  name: 'deliver_post',
  description: 'Return the photo read, three caption options and exactly five hashtags.',
  input_schema: {
    type: 'object',
    properties: {
      read: {
        type: 'object',
        properties: {
          tags: { type: 'array', items: { type: 'string' }, description: '5 short lowercase descriptors of what is in the photo, e.g. "golden light".' },
          palette: { type: 'array', items: { type: 'string' }, description: '5 dominant colors as #RRGGBB, most dominant first.' },
          summary: { type: 'string', description: 'One plain sentence on what the photo shows and its mood.' }
        },
        required: ['tags', 'palette', 'summary']
      },
      captions: {
        type: 'object',
        properties: {
          story: { type: 'string', description: 'Storytelling caption, 2 to 4 short paragraphs, ends with a light call to comment or save.' },
          short: { type: 'string', description: 'One or two lines.' },
          question: { type: 'string', description: 'Short caption built around a question that invites comments.' }
        },
        required: ['story', 'short', 'question']
      },
      hashtags: {
        type: 'array',
        description: 'Exactly 5 hashtags (Instagram allows at most 5 per post): 1 big, 2 medium, 2 local or niche.',
        items: {
          type: 'object',
          properties: {
            tag: { type: 'string', description: 'Starts with #, no spaces.' },
            tier: { type: 'string', enum: ['big', 'medium', 'local'] },
            why: { type: 'string', description: 'Under 12 words.' }
          },
          required: ['tag', 'tier', 'why']
        }
      },
      altText: { type: 'string', description: 'Alt text for accessibility, one sentence.' }
    },
    required: ['read', 'captions', 'hashtags', 'altText']
  }
};

function actionGenerate_(b) {
  const m = requireMember_(b.token);
  const image = String(b.image || '');
  if (!image) throw userError_('Add a photo first.');
  if (image.length > CFG.MAX_IMAGE_CHARS) throw userError_('That photo is too large. Try a smaller one.');
  const mediaType = /^image\/(jpeg|png|webp|gif)$/.test(b.mediaType) ? b.mediaType : 'image/jpeg';
  const profile = getProfile_(m.email);
  const system = [
    'You help an Instagram creator turn one photo into a ready-to-post caption and hashtags.',
    VOICE_RULES,
    'Hashtags: Instagram now allows at most 5 hashtags per post, so return exactly 5: one big broad tag, two medium tags, and two local or niche tags (a specific place, community or sub-niche). Lowercase. No banned, spammy or engagement-bait tags (#like4like, #follow, #instagood). Korean hashtags are good when the place or audience is Korean.',
    'The three captions must be clearly different from each other, not rewordings.',
    '',
    'STYLE PROFILE',
    profileText_(profile),
    '',
    trendHint_()
  ].join('\n');

  const story = String(b.story || '').slice(0, 1500);
  const vibe = String(b.vibe || 'Chill').slice(0, 30);
  const content = [
    { type: 'image', source: { type: 'base64', media_type: mediaType, data: image } },
    { type: 'text', text: 'What I was doing: ' + (story || '(nothing added, go by the photo)') + '\nVibe: ' + vibe }
  ];
  const out = withCredit_(m, function () { return claudeStructured_(system, content, POST_TOOL, 2000); });
  out.hashtags = (out.hashtags || []).slice(0, 5).map(function (h) {
    h.tag = '#' + String(h.tag || '').replace(/^#+/, '').replace(/\s+/g, '').toLowerCase();
    return h;
  });
  return { post: out, member: publicMember_(m) };
}

/* ------------------------------------------------------------------ */
/* Profiles                                                            */
/* ------------------------------------------------------------------ */

function getProfile_(email) {
  const sh = tab_(TABS.PROFILES);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (normEmail_(values[i][0]) === email) {
      try { return JSON.parse(values[i][2] || '{}'); } catch (e) { return {}; }
    }
  }
  return {};
}

function actionSaveProfile_(b) {
  const m = requireMember_(b.token);
  const p = b.profile || {};
  const clean = {
    niche: str_(p.niche, 200), audience: str_(p.audience, 300), tone: str_(p.tone, 200),
    language: str_(p.language, 60), location: str_(p.location, 120), avoid: str_(p.avoid, 400),
    samples: str_(p.samples, 4000),
    pillars: [].concat(p.pillars || []).map(function (x) { return str_(x, 40); }).filter(Boolean).slice(0, 6)
  };
  const sh = tab_(TABS.PROFILES);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (normEmail_(values[i][0]) === m.email) {
      sh.getRange(i + 1, 2, 1, 2).setValues([[stamp_(), JSON.stringify(clean)]]);
      return { profile: clean };
    }
  }
  sh.appendRow([m.email, stamp_(), JSON.stringify(clean)]);
  return { profile: clean };
}

function str_(x, max) { return String(x == null ? '' : x).trim().slice(0, max); }

/* ------------------------------------------------------------------ */
/* Drafts (Planner)                                                    */
/* ------------------------------------------------------------------ */

function draftsFor_(email) {
  const sh = tab_(TABS.DRAFTS);
  const values = sh.getDataRange().getValues();
  const c = colMap_(sh);
  const list = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (normEmail_(r[c['Email']]) !== email) continue;
    let extra = {};
    try { extra = JSON.parse(r[c['Extra JSON']] || '{}'); } catch (e) { extra = {}; }
    const planned = r[c['Planned date']];
    list.push({
      id: String(r[c['ID']]),
      row: i + 1,
      updated: String(r[c['Updated']]),
      date: planned instanceof Date ? Utilities.formatDate(planned, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(planned || ''),
      status: String(r[c['Status']] || 'Draft'),
      pillar: String(r[c['Pillar']] || ''),
      caption: String(r[c['Caption']] || ''),
      hashtags: String(r[c['Hashtags']] || '').split(/\s+/).filter(Boolean),
      story: String(r[c['Story']] || ''),
      thumb: String(r[c['Thumbnail']] || ''),
      readTags: extra.readTags || [],
      palette: extra.palette || []
    });
  }
  list.sort(function (a, b) { return (b.date || '9999').localeCompare(a.date || '9999'); });
  return list;
}

function actionListDrafts_(b) {
  const m = requireMember_(b.token);
  return { drafts: draftsFor_(m.email).map(stripRow_) };
}

function stripRow_(d) { const o = Object.assign({}, d); delete o.row; return o; }

function actionSaveDraft_(b) {
  const m = requireMember_(b.token);
  const d = b.draft || {};
  let thumb = String(d.thumb || '');
  if (thumb.length > CFG.MAX_THUMB_CHARS) thumb = '';
  const date = /^\d{4}-\d{2}-\d{2}$/.test(d.date || '') ? d.date : '';
  const status = ['Draft', 'Ready', 'Posted'].indexOf(d.status) >= 0 ? d.status : 'Draft';
  const tags = [].concat(d.hashtags || []).slice(0, 5).join(' ');
  const extra = JSON.stringify({
    readTags: [].concat(d.readTags || []).slice(0, 6),
    palette: [].concat(d.palette || []).slice(0, 6)
  });

  const sh = tab_(TABS.DRAFTS);
  const c = colMap_(sh);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let rowNum = 0;
    if (d.id) {
      const found = draftsFor_(m.email).filter(function (x) { return x.id === String(d.id); })[0];
      if (found) rowNum = found.row;
    }
    const id = rowNum ? String(d.id) : Utilities.getUuid().slice(0, 12);
    const row = new Array(sh.getLastColumn()).fill('');
    row[c['ID']] = id;
    row[c['Email']] = m.email;
    row[c['Updated']] = stamp_();
    row[c['Planned date']] = date ? "'" + date : '';
    row[c['Status']] = status;
    row[c['Pillar']] = str_(d.pillar, 40);
    row[c['Caption']] = str_(d.caption, 2200);
    row[c['Hashtags']] = tags;
    row[c['Story']] = str_(d.story, 1500);
    row[c['Extra JSON']] = extra;
    if (rowNum) {
      if (!thumb) row[c['Thumbnail']] = sh.getRange(rowNum, c['Thumbnail'] + 1).getValue();
      else row[c['Thumbnail']] = thumb;
      sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
    } else {
      row[c['Thumbnail']] = thumb;
      sh.appendRow(row);
    }
    return { id: id };
  } finally {
    lock.releaseLock();
  }
}

function actionDeleteDraft_(b) {
  const m = requireMember_(b.token);
  const found = draftsFor_(m.email).filter(function (x) { return x.id === String(b.id); })[0];
  if (found) tab_(TABS.DRAFTS).deleteRow(found.row);
  return {};
}

/* ------------------------------------------------------------------ */
/* Trends                                                              */
/* ------------------------------------------------------------------ */

function latestTrends_() {
  const sh = tab_(TABS.TRENDS);
  const values = sh.getDataRange().getValues();
  const out = { updated: '', regions: {} };
  for (let i = values.length - 1; i >= 1; i--) {
    const region = String(values[i][1]);
    if (out.regions[region]) continue;
    try {
      out.regions[region] = JSON.parse(values[i][2]);
      if (!out.updated || String(values[i][0]) > out.updated) out.updated = String(values[i][0]);
    } catch (e) { /* skip bad rows */ }
  }
  return out;
}

function actionTrends_(b) {
  requireMember_(b.token);
  const t = latestTrends_();
  t.niches = settings_().niches;
  return { trends: t };
}

/** Runs every morning (trigger made by setup). You can also run it by hand. */
function refreshTrends() {
  const st = settings_();
  const today = Utilities.formatDate(now_(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  st.regions.forEach(function (region) {
    try {
      const data = fetchTrendsFor_(region, st.niches, today);
      tab_(TABS.TRENDS).appendRow([stamp_(), region, JSON.stringify(data)]);
    } catch (err) {
      console.error('refreshTrends ' + region, err && err.stack || err);
    }
  });
  pruneTrends_();
}

function fetchTrendsFor_(region, niches, today) {
  const korea = /korea/i.test(region);
  const tool = { type: 'web_search_20250305', name: 'web_search', max_uses: 6 };
  if (korea) tool.user_location = { type: 'approximate', country: 'KR', timezone: 'Asia/Seoul' };

  const prompt = [
    'Today is ' + today + '. Research what is working on Instagram right now for creators' + (korea ? ' in South Korea (search in Korean and English)' : ' globally') + '.',
    'Focus on the last 30 days. Niches to cover: ' + niches.join(', ') + '.',
    'Find: which post formats are getting reach (Reels, carousels, single photos, specific styles), topics and seasonal moments coming up in the next 6 weeks, and hashtags that are rising in these niches.',
    'Remember Instagram allows only 5 hashtags per post, so favor specific, relevant tags over huge generic ones.',
    'Only include things your searches support. If something is uncertain, leave it out.',
    '',
    'Reply with ONLY a JSON object, no other text, in this shape:',
    '{"summary":"2 sentences","formats":[{"name":"","signal":"what is happening","tip":"how to use it"}],',
    '"topics":[{"title":"","why":"","timing":"e.g. mid-November","niches":["travel"]}],',
    '"hashtags":[{"tag":"#...","niche":"travel","trend":"rising|steady","note":"under 12 words"}],',
    '"sources":[{"title":"","url":""}]}',
    'Give 3 to 5 formats, 4 to 8 topics and 8 to 15 hashtags. Use niche names from the list given.'
  ].join('\n');

  let messages = [{ role: 'user', content: prompt }];
  let resp;
  for (let turn = 0; turn < 4; turn++) {
    resp = callClaude_({ model: CFG.MODEL, max_tokens: 4000, tools: [tool], messages: messages });
    if (resp.stop_reason !== 'pause_turn') break;
    messages = messages.concat([{ role: 'assistant', content: resp.content }]);
  }

  const text = textOf_(resp);
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start < 0 || end < start) throw new Error('No JSON in trends reply.');
  const data = JSON.parse(text.slice(start, end + 1));

  const seen = {};
  const sources = [];
  [].concat(data.sources || []).forEach(function (s) {
    if (s && s.url && !seen[s.url]) { seen[s.url] = 1; sources.push({ title: s.title || s.url, url: s.url }); }
  });
  (resp.content || []).forEach(function (blk) {
    (blk.citations || []).forEach(function (ct) {
      if (ct.url && !seen[ct.url]) { seen[ct.url] = 1; sources.push({ title: ct.title || ct.url, url: ct.url }); }
    });
  });
  data.sources = sources.slice(0, 12);
  data.region = region;
  data.date = today;
  return data;
}

function pruneTrends_() {
  const sh = tab_(TABS.TRENDS);
  const extra = sh.getLastRow() - 1 - 60;
  if (extra > 0) sh.deleteRows(2, extra);
}

/* ------------------------------------------------------------------ */
/* Creative studio                                                     */
/* ------------------------------------------------------------------ */

const IDEAS_TOOL = {
  name: 'deliver_ideas',
  description: 'Return creative ideas for the next post.',
  input_schema: {
    type: 'object',
    properties: {
      concepts: {
        type: 'array', description: '3 to 4 distinct post concepts.',
        items: {
          type: 'object',
          properties: {
            title: { type: 'string' },
            shot: { type: 'string', description: 'What to shoot and how: framing, light, timing. 1-2 sentences.' },
            why: { type: 'string', description: 'Why it should land with this audience. Under 20 words.' }
          },
          required: ['title', 'shot', 'why']
        }
      },
      sequence: {
        type: 'array', description: 'For a carousel: the slide order (5 to 8). For a reel: the beats (4 to 6). For a single photo: 2 to 3 alternate frames to try.',
        items: { type: 'object', properties: { label: { type: 'string' }, purpose: { type: 'string' } }, required: ['label', 'purpose'] }
      },
      hooks: { type: 'array', items: { type: 'string' }, description: '4 opening lines or on-screen hooks, each under 12 words.' },
      coverText: { type: 'array', items: { type: 'string' }, description: '3 short cover or overlay texts, under 6 words each.' },
      gradeTip: { type: 'string', description: 'One sentence on color and editing so it matches their feed.' }
    },
    required: ['concepts', 'sequence', 'hooks', 'coverText', 'gradeTip']
  }
};

function actionIdeas_(b) {
  const m = requireMember_(b.token);
  const kind = ['single', 'carousel', 'reel'].indexOf(b.kind) >= 0 ? b.kind : 'single';
  const brief = str_(b.brief, 1500);
  const system = [
    'You are a creative director helping one Instagram creator plan their next post.',
    VOICE_RULES,
    'Ideas must be practical to shoot on a phone by one person.',
    '',
    'STYLE PROFILE',
    profileText_(getProfile_(m.email)),
    trendHint_()
  ].join('\n');
  const recent = draftsFor_(m.email).slice(0, 6).map(function (d) { return '- ' + (d.caption || '').slice(0, 140); }).join('\n');
  const content = 'Post type: ' + kind + '\nWhat I am planning: ' + (brief || '(open, suggest something that fits my account)') +
    (recent ? '\nMy recent drafts, for context (avoid repeating them):\n' + recent : '');
  const ideas = withCredit_(m, function () { return claudeStructured_(system, content, IDEAS_TOOL, 2200); });
  return { ideas: ideas, kind: kind, member: publicMember_(m) };
}

/* ------------------------------------------------------------------ */
/* Planner direction                                                   */
/* ------------------------------------------------------------------ */

const DIRECTION_TOOL = {
  name: 'deliver_direction',
  description: 'Return a short read on where the account should go next.',
  input_schema: {
    type: 'object',
    properties: {
      summary: { type: 'string', description: '2 to 3 sentences, plain and direct.' },
      pillars: {
        type: 'array',
        items: {
          type: 'object',
          properties: { pillar: { type: 'string' }, count: { type: 'integer' }, advice: { type: 'string', description: 'Under 15 words.' } },
          required: ['pillar', 'count', 'advice']
        }
      },
      nextPosts: {
        type: 'array', description: '3 to 5 specific next posts.',
        items: {
          type: 'object',
          properties: {
            idea: { type: 'string' }, pillar: { type: 'string' },
            when: { type: 'string', description: 'A suggested day or timing.' },
            why: { type: 'string', description: 'Under 20 words.' }
          },
          required: ['idea', 'pillar', 'when', 'why']
        }
      }
    },
    required: ['summary', 'pillars', 'nextPosts']
  }
};

function actionDirection_(b) {
  const m = requireMember_(b.token);
  const profile = getProfile_(m.email);
  const drafts = draftsFor_(m.email).slice(0, 15).map(function (d) {
    return '- [' + (d.date || 'no date') + ', ' + d.status + ', pillar: ' + (d.pillar || 'none') + '] ' + (d.caption || '').slice(0, 160);
  }).join('\n');
  const t = latestTrends_();
  const trendLines = [];
  Object.keys(t.regions || {}).forEach(function (r) {
    const d = t.regions[r];
    if (d.summary) trendLines.push(r + ': ' + d.summary);
    (d.topics || []).slice(0, 4).forEach(function (x) { trendLines.push('- ' + x.title + ' (' + (x.timing || '') + ')'); });
  });
  const today = Utilities.formatDate(now_(), Session.getScriptTimeZone(), 'EEEE, yyyy-MM-dd');
  const system = [
    'You are a calm, honest content strategist for one Instagram creator. Give direction they can act on this week.',
    'Count posts per content pillar from the drafts given. If there are no drafts, say so and suggest a starting mix.',
    'No hype words. Be specific.',
    '',
    'STYLE PROFILE',
    profileText_(profile)
  ].join('\n');
  const content = 'Today is ' + today + '.\nMy drafts and planned posts:\n' + (drafts || '(none yet)') +
    '\n\nCurrent trends:\n' + (trendLines.join('\n') || '(no trend data yet)');
  const direction = withCredit_(m, function () { return claudeStructured_(system, content, DIRECTION_TOOL, 1800); });
  return { direction: direction, member: publicMember_(m) };
}
