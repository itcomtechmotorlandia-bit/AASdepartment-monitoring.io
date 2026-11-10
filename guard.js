/* Runs first on index.html: sends visitors to the login page when they are not signed in. */
(function(){
  const K='sfd_session',MAX=8*3600*1000;let s=null;
  try{s=JSON.parse(sessionStorage.getItem(K))}catch(e){}
  if(!s||!s.u||Date.now()-s.t>MAX){sessionStorage.removeItem(K);location.replace('login.html');return}
  window.SFD_USER=s;
  document.addEventListener('DOMContentLoaded',function(){
    var w=document.getElementById('whoami');if(w)w.textContent='👤 '+s.name;
    var b=document.getElementById('logout');
    if(b)b.addEventListener('click',function(){sessionStorage.removeItem(K);location.replace('login.html')});
  });
})();
