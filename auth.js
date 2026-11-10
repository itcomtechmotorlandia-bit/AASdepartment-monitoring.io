/* Firebase setup + login helpers (used by login.html and index.html).
   The firebaseConfig below is NOT a secret. Your data is protected by the Firestore rules
   and by the users you added in Firebase Authentication. */
const firebaseConfig = {
  apiKey: "AIzaSyBL2yVpMgaFn0ex2RvpLiNEsjDBKphYV5E",
  authDomain: "aasdepartmentmonitoring.firebaseapp.com",
  projectId: "aasdepartmentmonitoring",
  storageBucket: "aasdepartmentmonitoring.firebasestorage.app",
  messagingSenderId: "672532075383",
  appId: "1:672532075383:web:16df130cedaaddef0f0772",
  measurementId: "G-EMDLKJ501F"
};
const SFD_MAX_FAILS = 5, SFD_LOCK_SECONDS = 60, SFD_SESSION_HOURS = 8;
const SFD_NAMES = {
  'motoassap@gmail.com': 'Moto ASSAP',
  'aas.dept01@gmail.com': 'AAS Department',
  'itcomtechmotorlandia@gmail.com': 'IT Comtech'
};
let sfdAuth = null;
try {
  if (!firebase.apps.length) firebase.initializeApp(firebaseConfig);
  sfdAuth = firebase.auth();
} catch (e) { console.error('Firebase failed to start', e); }

/* Sign in. Returns {ok:true} or {ok:false, msg:'...'} */
async function sfdLogin(email, pw) {
  if (!sfdAuth) return { ok: false, msg: 'Could not connect to Firebase. Check your internet and reload.' };
  try {
    /* SESSION = signed out when the browser tab/window is closed (same as before) */
    await sfdAuth.setPersistence(firebase.auth.Auth.Persistence.SESSION);
    await sfdAuth.signInWithEmailAndPassword(email.trim(), pw);
    sessionStorage.setItem('sfd_login_t', String(Date.now()));
    return { ok: true };
  } catch (e) {
    const c = e && e.code || '';
    if (c === 'auth/network-request-failed') return { ok: false, msg: 'No internet connection.' };
    if (c === 'auth/too-many-requests') return { ok: false, msg: 'Too many attempts. Please wait a few minutes or reset your password.' };
    if (c === 'auth/user-disabled') return { ok: false, msg: 'This account is disabled.' };
    if (c === 'auth/unauthorized-domain') return { ok: false, msg: 'This website address is not authorized in Firebase (Authentication > Settings > Authorized domains).' };
    return { ok: false, msg: 'Wrong email or password.', bad: true };
  }
}

/* Send a password reset link to the email. Returns '' on success or an error message. */
async function sfdReset(email) {
  if (!sfdAuth) return 'Could not connect to Firebase. Check your internet and reload.';
  try { await sfdAuth.sendPasswordResetEmail(email.trim()); return ''; }
  catch (e) {
    const c = e && e.code || '';
    if (c === 'auth/invalid-email') return 'Please type a valid email.';
    if (c === 'auth/network-request-failed') return 'No internet connection.';
    if (c === 'auth/too-many-requests') return 'Too many requests. Please try again later.';
    return 'Could not send the reset email. Try again.';
  }
}
