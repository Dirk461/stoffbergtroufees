/**
 * Stoffberg Troufees – guest list, visit log and directory.
 *
 * Paste this into the Google Sheet: Extensions > Apps Script, replace everything, save.
 * Then reload the sheet and use the "Troufees" menu:
 *   1. "Set up tabs"            – creates Guests, Log and Directory (run once)
 *   2. "Make personal links"    – gives every guest without a code its own link
 * Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone),
 * and paste the web-app URL into config.js on GitHub.
 *
 * Guests tab, one row per invite (or per person – your choice):
 *   Code | Invite name | Names (separate with ;) | Language (en/af/nl) | Personal link
 *   "Invite name" is how the page greets them ("Jan & Marie").
 *   "Names" are the full names allowed to open that link ("Jan van der Merwe; Marie van der Merwe").
 */

const SITE_URL = 'https://stoffbergtroufees.com';
const TABS = {
  Guests:    ['Code', 'Invite name', 'Names (separate with ;)', 'Language (en/af/nl)', 'Personal link'],
  Log:       ['Time', 'Event', 'Code', 'Invite', 'Name typed', 'Matched', 'Language', 'Device'],
  Directory: ['Invite', 'Code', 'Personal link', 'Name entered', 'First visit', 'Last visit', 'Visits', 'Added to calendar', 'Calendar clicked at'],
  Scores:    ['Code', 'Player', 'Best score', 'Caught Ané', 'Games played', 'Last played'],
};

// ---------- Menu ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Troufees')
    .addItem('Set up tabs', 'setupTabs')
    .addItem('Make personal links', 'makeLinks')
    .addToUi();
}

function setupTabs() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(TABS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  const first = ss.getSheets()[0];
  if (first.getName() === 'Sheet1' && first.getLastRow() === 0) ss.deleteSheet(first);
  SpreadsheetApp.getUi().alert('Tabs ready. Fill in Guests, then run "Make personal links".');
}

function makeLinks() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Guests');
  const rows = sh.getLastRow() - 1;
  if (rows < 1) return;
  const data = sh.getRange(2, 1, rows, 5).getValues();
  const used = new Set(data.map(r => String(r[0]).trim()).filter(Boolean));
  data.forEach(r => {
    if (!String(r[1]).trim() && !String(r[2]).trim()) return;
    if (!String(r[0]).trim()) {
      let c; do { c = newCode(); } while (used.has(c));
      used.add(c); r[0] = c;
    }
    r[4] = SITE_URL + '/g/' + String(r[0]).trim();
  });
  sh.getRange(2, 1, rows, 5).setValues(data);
  SpreadsheetApp.getUi().alert('Personal links are in column E. Existing codes were kept.');
}

function newCode() {
  const abc = 'abcdefghjkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 8; i++) s += abc[Math.floor(Math.random() * abc.length)];
  return s;
}

// ---------- Web app ----------
function doGet() {
  return ContentService.createTextOutput('Troufees is running.');
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json({ ok: false, error: 'bad_request' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (req.action === 'verify') return json(verify(req));
    if (req.action === 'open') { record('Opened', findByCode(req.code), req, 'yes'); return json({ ok: true }); }
    if (req.action === 'calendar') { record('Added to calendar', findByCode(req.code), req, 'yes'); return json({ ok: true }); }
    if (req.action === 'score') return json(saveScore(req));
    if (req.action === 'leaderboard') return json({ ok: true, rows: leaderboard(req.code) });
    return json({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}

function verify(req) {
  const typed = clip(req.name, 120);
  const byCode = findByCode(req.code);
  let guest = byCode && (nameMatches(typed, byCode.names) || firstNameMatches(typed, byCode.names)) ? byCode : null;
  if (!guest) guest = guests().find(g => nameMatches(typed, g.names)) || null;   // right name, other (or no) link
  if (!guest) {
    record('Name not matched', byCode, req, 'no');
    return { ok: false, reason: 'no_match' };
  }
  const canonical = guest.names.find(n => nameMatches(typed, [n])) || titleCase(typed);   // first-name-only entries keep the full name they typed
  record('Name entered', guest, req, 'yes', canonical);
  return { ok: true, code: guest.code, greeting: guest.invite, names: guest.names, lang: guest.lang, name: canonical };
}

// ---------- Game leaderboard ----------
// One row per player (link + name). Only guests with a real link can post a score.
function scoresSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Scores');
  if (!sh) {
    sh = ss.insertSheet('Scores');
    sh.getRange(1, 1, 1, TABS.Scores.length).setValues([TABS.Scores]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function saveScore(req) {
  const guest = findByCode(req.code);
  const score = Math.floor(Number(req.score));
  if (!guest || !(score >= 0 && score <= 5000)) return { ok: false };
  const player = clip(req.name, 120) || guest.invite;
  const sh = scoresSheet(), now = new Date();
  const last = sh.getLastRow();
  const data = last > 1 ? sh.getRange(2, 1, last - 1, 6).getValues() : [];
  const i = data.findIndex(r => String(r[0]) === guest.code && fold(r[1]) === fold(player));
  if (i < 0) {
    sh.appendRow([guest.code, player, score, req.won ? 'Yes' : 'No', 1, now]);
  } else {
    const r = data[i];
    r[2] = Math.max(Number(r[2]) || 0, score);
    if (req.won) r[3] = 'Yes';
    r[4] = (Number(r[4]) || 0) + 1;
    r[5] = now;
    sh.getRange(i + 2, 1, 1, 6).setValues([r]);
  }
  return { ok: true, rows: leaderboard(guest.code, player) };
}
// Top 10, shown as first name + surname initial ("Marie v.").
function leaderboard(code, player) {
  const sh = scoresSheet();
  const last = sh.getLastRow();
  if (last < 2) return [];
  code = String(code || '').toLowerCase();
  return sh.getRange(2, 1, last - 1, 4).getValues()
    .filter(r => String(r[0]) && Number(r[2]) > 0)
    .sort((a, b) => Number(b[2]) - Number(a[2]))
    .slice(0, 10)
    .map(r => ({ name: shortName(r[1]), score: Number(r[2]), caught: r[3] === 'Yes', me: String(r[0]) === code && (!player || fold(r[1]) === fold(player)) }));
}
function shortName(n) {
  const parts = String(n).trim().split(/\s+/);
  return parts.length > 1 ? parts[0] + ' ' + parts[parts.length - 1].charAt(0).toUpperCase() + '.' : parts[0];
}

// ---------- Sheet helpers ----------
function guests() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Guests');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({
      code: String(r[0]).trim().toLowerCase(),
      invite: String(r[1]).trim() || String(r[2]).split(';')[0].trim(),
      names: String(r[2]).split(';').map(s => s.trim()).filter(Boolean),
      lang: ['en', 'af', 'nl'].indexOf(String(r[3]).trim().toLowerCase()) >= 0 ? String(r[3]).trim().toLowerCase() : '',   // empty: page keeps the visitor's own language
      link: String(r[4]).trim() || SITE_URL + '/g/' + String(r[0]).trim(),
    }));
}
function findByCode(code) {
  code = String(code || '').trim().toLowerCase();
  return code ? guests().find(g => g.code === code) || null : null;
}

// "Jan van der Merwe" matches "jan VAN DER merwe", "Jan v.d. Merwe" does not.
// A middle name may be left out: "Anna Maria Botha" also opens with "Anna Botha".
function fold(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/[^a-z0-9\s'-]/g, ' ').replace(/\s+/g, ' ').trim();
}
function nameMatches(typed, names) {
  const t = fold(typed);
  if (!t || t.indexOf(' ') < 0) return false;
  return names.some(n => {
    const f = fold(n);
    if (f === t) return true;
    const a = f.split(' '), b = t.split(' ');
    return b.length === 2 && a.length > 2 && a[0] === b[0] && a[a.length - 1] === b[1];   // first + last name only
  });
}

// While the guest list only has a first name (or "Oom Rene", "Tannie Yvonne"), the personal link plus
// the right first name is enough. The full name they type is saved in the Directory. Once you put full
// names in the Guests tab, the stricter full-name check above takes over for that guest.
function firstNameMatches(typed, names) {
  const first = fold(typed).split(' ')[0];
  return names.some(n => {
    const f = fold(n).replace(/^(oom|tannie|ouma|oupa|tant) /, '');
    return f.indexOf(' ') < 0 && f === first;
  });
}
function titleCase(s) {
  return String(s).trim().replace(/\s+/g, ' ').split(' ').map(w => /^(van|der|den|de|du|le|la|von)$/i.test(w) ? w.toLowerCase() : w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
}

function record(event, guest, req, matched, canonical) {
  const ss = SpreadsheetApp.getActive();
  const now = new Date();
  const code = guest ? guest.code : String(req.code || '').toLowerCase();
  ss.getSheetByName('Log').appendRow([now, event, code, guest ? guest.invite : '', clip(req.name, 120), matched,
    clip(req.lang, 5), device(req.ua)]);
  if (!guest) return;

  const dir = ss.getSheetByName('Directory');
  const last = dir.getLastRow();
  const codes = last > 1 ? dir.getRange(2, 2, last - 1, 1).getValues().map(r => String(r[0])) : [];
  let row = codes.indexOf(guest.code);
  let vals;
  if (row < 0) {
    vals = [guest.invite, guest.code, guest.link, '', now, now, 0, 'No', ''];
    row = last + 1;                      // new row at the bottom
  } else {
    row = row + 2;
    vals = dir.getRange(row, 1, 1, 9).getValues()[0];
  }
  if (req.name && event !== 'Name not matched') {
    const seen = String(vals[3]).split(', ').filter(Boolean);
    const nm = clip(canonical || req.name, 120);
    if (seen.map(fold).indexOf(fold(nm)) < 0) seen.push(nm);
    vals[3] = seen.join(', ');
  }
  vals[5] = now;
  if (event === 'Name entered' || event === 'Opened') vals[6] = Number(vals[6] || 0) + 1;
  if (event === 'Added to calendar') { vals[7] = 'Yes'; vals[8] = now; }
  dir.getRange(row, 1, 1, 9).setValues([vals]);
}

function device(ua) {
  ua = String(ua || '');
  if (/iPhone|iPad/.test(ua)) return 'iPhone/iPad';
  if (/Android/.test(ua)) return 'Android';
  if (/Macintosh/.test(ua)) return 'Mac';
  if (/Windows/.test(ua)) return 'Windows';
  return ua ? 'Other' : '';
}
function clip(s, n) { return String(s == null ? '' : s).slice(0, n); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
