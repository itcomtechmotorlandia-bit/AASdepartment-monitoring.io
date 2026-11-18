/* Google Apps Script - two-way link between the dashboard's Mechanic Performance and this Google Sheet.
   The dashboard can SEND a month to the sheet and GET it back after you edit it here.
   Only the 3 allowed dashboard users can read or write. */

const SHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';   // the long text in the sheet's address between /d/ and /edit
const FIREBASE_API_KEY = 'AIzaSyBL2yVpMgaFn0ex2RvpLiNEsjDBKphYV5E';
const SITE = 'https://itcomtechmotorlandia-bit.github.io/';
const ALLOWED = ['motoassap@gmail.com', 'aas.dept01@gmail.com', 'itcomtechmotorlandia@gmail.com'];
const FIRST = 5;   // first data row (rows 1-4 are the title, note, blank and header)

function doGet() { return out({ ok: true, msg: 'Sheet link is running.' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    const email = whoIs(d.token);
    if (!email || ALLOWED.indexOf(email) < 0) return out({ ok: false, error: 'This account is not allowed.' });
    const name = String(d.tab || '').replace(/[\[\]\*\?\:\/\\]/g, ' ').trim().slice(0, 90);
    if (!name) return out({ ok: false, error: 'Missing tab name.' });
    const ss = SpreadsheetApp.openById(SHEET_ID);
    if (d.action === 'get') return getTab(ss, name);
    if (d.action === 'put') return putTab(ss, name, d, email);
    return out({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  } finally { try { lock.releaseLock(); } catch (x) {} }
}

function dataValues(sh, w) {
  const last = sh.getLastRow();
  if (last < FIRST) return [];
  return sh.getRange(FIRST, 1, last - FIRST + 1, w).getValues();
}
function hashOf(vals) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(vals));
  return b.map(x => ('0' + (x & 0xff).toString(16)).slice(-2)).join('');
}
function tabUrl(ss, sh) { return ss.getUrl() + '#gid=' + sh.getSheetId(); }

function getTab(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) return out({ ok: true, found: false });
  const w = Math.max(1, sh.getLastColumn());
  const vals = dataValues(sh, w);
  const head = w ? sh.getRange(FIRST - 1, 1, 1, w).getValues()[0] : [];
  const rows = vals.filter(r => String(r[0]).trim() && String(r[0]).trim().toUpperCase() !== 'TOTAL');
  return out({ ok: true, found: true, head: head, rows: rows, hash: hashOf(vals), url: tabUrl(ss, sh) });
}

function putTab(ss, name, d, email) {
  if (!d.head || !d.rows || d.head.length > 60 || d.rows.length > 500) return out({ ok: false, error: 'Unexpected data size.' });
  const w = d.head.length;
  let sh = ss.getSheetByName(name);
  if (sh && !d.force) {
    const cur = hashOf(dataValues(sh, w));
    const has = sh.getLastRow() >= FIRST;
    if (has && cur !== String(d.baseHash || '')) return out({ ok: false, conflict: true, error: 'The sheet was edited since the last sync.' });
  }
  const clean = v => {                       // stop text from running as a formula
    if (typeof v === 'number') return v;
    const s = String(v == null ? '' : v);
    return /^[=+\-@]/.test(s) ? "'" + s : s;
  };
  const fix = r => { const a = r.map(clean); while (a.length < w) a.push(''); return a.slice(0, w); };
  const when = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd HH:mm');
  const n = d.rows.length;
  const vals = [
    fix([String(d.title || name)]),
    fix(['Last sent from the dashboard ' + when + ' by ' + String(d.by || email) + '. You can edit any cell here, then press "Get from Sheet" on the website.']),
    fix(['']),
    fix(d.head)
  ].concat(d.rows.map(fix));
  const total = new Array(w).fill('');
  total[0] = 'TOTAL';
  vals.push(total);

  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, vals.length, w).setValues(vals);
  const L = c => { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
  const tr = FIRST + n;
  for (let c = 3; c <= w; c++) if (c !== 4) sh.getRange(tr, c).setFormula('=SUM(' + L(c) + FIRST + ':' + L(c) + (tr - 1) + ')');
  sh.getRange(tr, 1, 1, w).setFontWeight('bold').setBackground('#e8eefb');
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  sh.getRange(2, 1).setFontStyle('italic').setFontColor('#555555');
  sh.getRange(FIRST - 1, 1, 1, w).setFontWeight('bold').setBackground('#1b2d50').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
  if (n) {
    sh.getRange(FIRST, 3, n, 1).setNumberFormat('#,##0');
    sh.getRange(FIRST, 5, n, w - 4).setNumberFormat('#,##0');
    sh.getRange(FIRST, 4, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['1ST GRADE', '2ND GRADE', '3RD GRADE'], true).setAllowInvalid(false).build());
    sh.getRange(FIRST, 1, n, w).setBackground('#fffdf0');
  }
  sh.setFrozenRows(FIRST - 1); sh.setFrozenColumns(1);
  sh.autoResizeColumns(1, w);
  SpreadsheetApp.flush();
  return out({ ok: true, tab: name, rows: n, hash: hashOf(dataValues(sh, w)), url: tabUrl(ss, sh) });
}

// Checks the dashboard login token with Firebase and returns the user's email.
function whoIs(token) {
  if (!token) return '';
  const r = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Referer: SITE },
    payload: JSON.stringify({ idToken: token })
  });
  if (r.getResponseCode() !== 200) throw new Error('Login check failed (' + r.getResponseCode() + '): ' + r.getContentText().slice(0, 200));
  const u = (JSON.parse(r.getContentText()).users || [])[0];
  return u && u.email ? String(u.email).toLowerCase() : '';
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
