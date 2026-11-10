/* Runs first on index.html:
   1) checks the Firebase login (not signed in -> login.html)
   2) loads the shared data from Firestore into the browser storage
   3) starts app.js
   4) keeps every change saved to Firestore and warns when someone else changes the data */
(function () {
  const MAX_AGE = SFD_SESSION_HOURS * 3600 * 1000, CH = 250000, SAVE_DELAY = 800;
  const KEYS = ['bd2', 'afterSalesReport', 'sparePartsWeekly', 'sparePartsTargets', 'sparePartsTargetsMonthly',
    'sparePartsClassifications', 'uniformMonitoring', 'memoPreview', 'travelExpense',
    'contractPreview', 'contractPreviewByBranch', 'facilityConcerns', 'facilityConcernsByMonth'];
  const SID = Math.random().toString(36).slice(2) + Date.now().toString(36);
  const origSet = Storage.prototype.setItem;
  let db, col, me, ready = false, firstSnap = true;
  const lastPushed = {}, known = {}, timers = {};

  function show() { const s = document.getElementById('sfdHide'); if (s) s.remove(); }
  function fatal(msg, retry) {
    show();
    const d = document.createElement('div');
    d.style.cssText = 'position:fixed;inset:0;z-index:200000;display:grid;place-items:center;background:#0a1324;color:#e9eef8;font:15px system-ui,sans-serif;padding:20px;visibility:visible';
    d.innerHTML = '<div style="max-width:420px;text-align:center"><h2 style="margin:0 0 10px">Cannot open the dashboard</h2><p style="line-height:1.5"></p><p><button id="sfdRe" style="padding:10px 18px;border-radius:8px;border:0;background:#3b82f6;color:#fff;font:inherit;cursor:pointer">Try again</button> <button id="sfdOut" style="padding:10px 18px;border-radius:8px;border:1px solid #556;background:none;color:inherit;font:inherit;cursor:pointer">Back to login</button></p></div>';
    d.querySelector('p').textContent = msg;
    document.body ? document.body.appendChild(d) : document.documentElement.appendChild(d);
    d.querySelector('#sfdRe').onclick = () => location.reload();
    d.querySelector('#sfdOut').onclick = () => { try { sfdAuth.signOut(); } catch (e) {} location.replace('login.html'); };
  }
  function toLogin() { location.replace('login.html'); }
  function onDom(fn) { document.readyState === 'loading' ? document.addEventListener('DOMContentLoaded', fn) : fn(); }

  /* ---------- Firestore sync ---------- */
  async function loadRemote() {
    const snap = await col.get(), metas = {}, chunks = {};
    snap.forEach(d => {
      const id = d.id, x = d.data();
      if (id.indexOf('~') > -1) { const p = id.split('~'); (chunks[p[0]] = chunks[p[0]] || [])[+p[1]] = x.s; }
      else metas[id] = x;
    });
    const out = {};
    Object.keys(metas).forEach(k => {
      const n = metas[k].n || 0, c = chunks[k] || []; let ok = true;
      for (let i = 0; i < n; i++) if (typeof c[i] !== 'string') ok = false;
      known[k] = n;
      if (ok) out[k] = c.slice(0, n).join('');
    });
    return out;
  }
  async function push(k) {
    const v = localStorage.getItem(k);
    if (v == null || v === lastPushed[k]) return;
    const parts = []; for (let i = 0; i < v.length; i += CH) parts.push(v.slice(i, i + CH));
    if (!parts.length) parts.push('');
    const b = db.batch();
    parts.forEach((p, i) => b.set(col.doc(k + '~' + i), { s: p }));
    b.set(col.doc(k), { n: parts.length, t: firebase.firestore.FieldValue.serverTimestamp(), by: me.name, sid: SID });
    for (let i = parts.length; i < (known[k] || 0); i++) b.delete(col.doc(k + '~' + i));
    await b.commit();
    known[k] = parts.length; lastPushed[k] = v;
  }
  function status(t, bad) {
    const w = document.getElementById('whoami'); if (!w) return;
    w.textContent = '👤 ' + me.name + (t ? ' · ' + t : ''); w.style.color = bad ? '#ec6b73' : '';
  }
  function schedule(k) {
    clearTimeout(timers[k]); status('saving…');
    timers[k] = setTimeout(async () => {
      delete timers[k];
      try { await push(k); if (!Object.keys(timers).length) status('saved ✓'); }
      catch (e) { console.error('Save failed', k, e); status('NOT saved - check internet', true); }
    }, SAVE_DELAY);
  }
  async function flushAll() {
    const pend = Object.keys(timers); pend.forEach(k => { clearTimeout(timers[k]); delete timers[k]; });
    for (const k of pend) await push(k);
  }
  function hookStorage() {
    Storage.prototype.setItem = function (k, v) {
      origSet.call(this, k, v);
      if (this === window.localStorage && ready && KEYS.indexOf(k) > -1) schedule(k);
    };
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') Object.keys(timers).forEach(k => { clearTimeout(timers[k]); delete timers[k]; push(k).catch(() => {}); }); });
    window.addEventListener('beforeunload', e => { if (Object.keys(timers).length) { e.preventDefault(); e.returnValue = ''; } });
  }
  function banner(by) {
    if (document.getElementById('sfdBn')) return;
    const d = document.createElement('div'); d.id = 'sfdBn';
    d.style.cssText = 'position:fixed;left:50%;bottom:18px;transform:translateX(-50%);z-index:99999;background:#1b2d50;color:#fff;padding:12px 16px;border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,.4);font:14px system-ui,sans-serif;display:flex;gap:12px;align-items:center';
    d.innerHTML = '<span></span><button style="border:0;border-radius:8px;padding:8px 12px;background:#3b82f6;color:#fff;font:inherit;cursor:pointer">Refresh now</button>';
    d.querySelector('span').textContent = '🔄 ' + by + ' updated the shared data.';
    d.querySelector('button').onclick = () => { Object.keys(timers).forEach(k => { clearTimeout(timers[k]); delete timers[k]; }); location.reload(); };
    document.body.appendChild(d);
  }
  function listen() {
    col.onSnapshot(snap => {
      if (firstSnap) { firstSnap = false; return; }
      snap.docChanges().forEach(c => {
        const x = c.doc.data();
        if (c.doc.id.indexOf('~') < 0 && x && x.sid && x.sid !== SID && !c.doc.metadata.hasPendingWrites) banner(x.by || 'Someone');
      });
    }, e => console.error('Live updates stopped', e));
  }

  /* ---------- start ---------- */
  async function start(user) {
    me = { email: user.email, name: SFD_NAMES[(user.email || '').toLowerCase()] || user.email };
    window.SFD_USER = me;
    db = firebase.firestore(); col = db.collection('appdata');
    let remote;
    try {
      remote = await Promise.race([loadRemote(), new Promise((_, rej) => setTimeout(() => rej({ code: 'timeout' }), 20000))]);
    } catch (e) {
      console.error(e);
      const c = (e && e.code) || (e && e.message) || 'unknown';
      return fatal(c === 'permission-denied'
        ? 'This account is not allowed to open the data. In Firebase > Firestore > Rules, publish the rules and make sure your email is listed. (permission-denied)'
        : c === 'timeout' || c === 'unavailable' || c === 'failed-precondition' || c === 'not-found'
        ? 'Cannot reach the Firestore database. In the Firebase console open Build > Firestore Database and click "Create database" if you have not yet, then try again. (' + c + ')'
        : 'Could not load the shared data. Error: ' + c);
    }
    const toUpload = [];
    KEYS.forEach(k => {
      if (remote[k] != null) { origSet.call(localStorage, k, remote[k]); lastPushed[k] = remote[k]; }
      else if (localStorage.getItem(k) != null) toUpload.push(k);   /* first time: upload what is already in this browser */
    });
    hookStorage();
    for (const k of toUpload) { try { await push(k); } catch (e) { console.error('Upload failed', k, e); } }
    ready = true;
    onDom(() => {
      const w = document.getElementById('whoami'); if (w) w.textContent = '👤 ' + me.name;
      const b = document.getElementById('logout');
      if (b) b.addEventListener('click', async () => {
        b.disabled = true;
        try { await flushAll(); }
        catch (e) { if (!confirm('Your latest changes could not be saved (no connection?). Log out anyway and lose them?')) { b.disabled = false; return; } }
        KEYS.forEach(k => localStorage.removeItem(k));   /* do not leave company data on this computer */
        try { await sfdAuth.signOut(); } catch (e) {}
        sessionStorage.removeItem('sfd_login_t'); toLogin();
      });
      const s = document.createElement('script'); s.src = 'app.js';
      s.onload = () => { show(); listen(); };
      s.onerror = () => fatal('app.js could not be loaded. Make sure app.js is uploaded next to index.html.');
      document.body.appendChild(s);
    });
  }

  if (typeof firebase === 'undefined' || !sfdAuth) {
    onDom(() => fatal('Firebase could not be loaded. Check your internet connection and try again.')); return;
  }
  const t = +sessionStorage.getItem('sfd_login_t') || 0;
  sfdAuth.onAuthStateChanged(user => {
    if (!user) return toLogin();
    if (t && Date.now() - t > MAX_AGE) { sfdAuth.signOut().then(toLogin, toLogin); return; }
    if (!t) sessionStorage.setItem('sfd_login_t', String(Date.now()));
    if (ready || window.__sfdStarted) return; window.__sfdStarted = true;
    start(user).catch(e => { console.error(e); fatal('Unexpected error: ' + ((e && e.message) || e)); });
  }, e => { console.error(e); fatal('Login check failed: ' + ((e && e.message) || e)); });
})();
