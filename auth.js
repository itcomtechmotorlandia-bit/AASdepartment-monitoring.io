/* Login accounts and helpers (passwords are stored only as salted PBKDF2 hashes).
   NOTE: this is a front-end only login. It is NOT secure against someone who knows web development. */
const SFD_SESSION_KEY='sfd_session',SFD_SESSION_HOURS=8,SFD_MAX_FAILS=5,SFD_LOCK_SECONDS=60,SFD_ITER=150000;
const USERS=[
  {u:'motoassap@gmail.com',name:'Moto ASSAP',s:'005898a6d627bbbee90744ab7bd02831',h:'f4bc53bcba854c69b55e03fd9327309cbbe8ad0edb9fb2a8d4094e31cfda463a'},
  {u:'aas.dept01@gmail.com',name:'AAS Department',s:'8d4857d54a19aaa5c01e343ef489e536',h:'8daf4138d6ba059d8ab267f74ce4260131bddff1d6c74b98f453f5cdd49da5f0'},
  {u:'itcomtechmotorlandia@gmail.com',name:'IT Comtech',s:'8fac45258671ea1dd64e984f233383d7',h:'135fad0512a9388865a68fb15b5d0d6c0562367f6ccd2cd5295cc405040e47b4'}
];
const RECOVERY={s:'6277dc24266e12e6b23edf622c637ffc',h:'86ca0efd0e3bea70b3e41c5e3cb8bc8793a232b05004cdd5abfdb726dd54c8b8'};
const hex=b=>[...new Uint8Array(b)].map(x=>x.toString(16).padStart(2,'0')).join(''),unhex=h=>new Uint8Array(h.match(/../g).map(x=>parseInt(x,16)));
async function sfdPb(pw,salt){const k=await crypto.subtle.importKey('raw',new TextEncoder().encode(pw),'PBKDF2',false,['deriveBits']);return hex(await crypto.subtle.deriveBits({name:'PBKDF2',hash:'SHA-256',salt:unhex(salt),iterations:SFD_ITER},k,256))}
function sfdOverrides(){try{return JSON.parse(localStorage.getItem('sfd_pw_override'))||{}}catch(e){return{}}}
async function sfdLogin(email,pw){email=email.trim().toLowerCase();const acc=USERS.find(a=>a.u===email);if(!acc)return null;const rec=sfdOverrides()[email]||acc;if(await sfdPb(pw,rec.s)!==rec.h)return null;
  sessionStorage.setItem(SFD_SESSION_KEY,JSON.stringify({u:acc.u,name:acc.name,t:Date.now()}));return acc}
async function sfdReset(email,code,newPw){email=email.trim().toLowerCase();if(!USERS.find(a=>a.u===email))return'This email is not authorized.';
  if(await sfdPb(code.trim().toUpperCase(),RECOVERY.s)!==RECOVERY.h)return'Wrong recovery code.';
  if(newPw.length<8)return'The new password must have at least 8 characters.';
  const s=hex(crypto.getRandomValues(new Uint8Array(16))),o=sfdOverrides();o[email]={s,h:await sfdPb(newPw,s)};try{localStorage.setItem('sfd_pw_override',JSON.stringify(o))}catch(e){return'Could not save the new password in this browser.'}return''}
