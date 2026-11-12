/* Standalone download helper: lets the Excel/PDF export buttons work outside Claude. */
if(!window.claude){window.claude={use:async()=>({save:async({filename,data})=>{const a=document.createElement('a');a.href=URL.createObjectURL(data);a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1500)}})}}
const $=s=>document.querySelector(s),peso=n=>'₱'+Math.round(n).toLocaleString('en-PH');
const D=26;
const N=['ALFONSO','ANGELES','ATIMONAN','BACOOR 1','BACOOR 2','BATAAN BALANGA','BATANGAS','BAY','BAYAMBANG','BINALONAN','BINANGONAN','CALAMBA','3S CALAMBA','CALAPAN','CALAUAG','CALIT-CALIT','CARMEN','CATANAUAN','DAET','DAGUPAN','GENERAL TRIAS','GUMACA','MS GUMACA','INFANTA','MS INFANTA','LAIYA SAN JUAN','LIMAY BATAAN','LIPA','LOPEZ','LOS BAÑOS','LUCENA','3S LUCENA','MALOLOS','MARILAO','NAGCARLAN','NAIC','ORANI','PAGBILAO','PANIQUI','PILILLA','PINAMALAYAN','PLARIDEL','PUERTO GALERA','SAN FRANCISCO','SAN PABLO','3S SAN PABLO','SILANG','SINILOAN','STA. CRUZ','STA. ROSA','3S STA. ROSA','STO TOMAS','TANAUAN','TANAY','TANZA','TARLAC'];
const AREAS={'BATAAN AREA':['BATAAN BALANGA','ORANI','LIMAY BATAAN'],'BATANGAS AREA':['BATANGAS','CALIT-CALIT','LAIYA SAN JUAN','LIPA','STO TOMAS','TANAUAN'],'BICOL AREA':['DAET'],'BULACAN AREA':['MALOLOS','MARILAO','PLARIDEL'],'CAVITE AREA':['ALFONSO','BACOOR 1','BACOOR 2','GENERAL TRIAS','NAIC','SILANG','TANZA'],'LAGUNA AREA':['BAY','CALAMBA','3S CALAMBA','LOS BAÑOS','NAGCARLAN','SAN PABLO','3S SAN PABLO','SINILOAN','STA. CRUZ','STA. ROSA','3S STA. ROSA'],'QUEZON AREA':['ATIMONAN','CATANAUAN','GUMACA','MS GUMACA','INFANTA','MS INFANTA','LOPEZ','LUCENA','3S LUCENA','CALAUAG','PAGBILAO','SAN FRANCISCO'],'MINDORO AREA':['CALAPAN','PINAMALAYAN','PUERTO GALERA'],'PAMPANGA AREA':['ANGELES'],'PANGASINAN AREA':['BAYAMBANG','BINALONAN','CARMEN','DAGUPAN'],'RIZAL AREA':['BINANGONAN','PILILLA','TANAY','SAMPALOC'],'TARLAC AREA':['TARLAC','PANIQUI']};
const B=[];
const SPB=['MS GUMACA','MS INFANTA','3S CALAMBA','3S LUCENA','3S SAN PABLO','3S STA ROSA'];let SPT={},SPC={};
const P=[];
let F=[];
const tip=document.createElement('div');tip.id='tip';document.body.appendChild(tip);
function showTip(e){const c=e.target.closest('.tt[data-j="3"]');if(!c){tip.style.opacity=0;return}const x=B[c.dataset.i],s=x.wi.reduce((a,b)=>a+b,0);tip.innerHTML=`<b>${x.b}</b><br>Monthly Target: ${peso(TG(x))}<br>Total Week Income: ${peso(s)}<hr>Difference: <b>${peso(TG(x)-s)}</b>`;tip.style.opacity=1;tip.style.left=Math.max(8,Math.min(e.clientX+14,innerWidth-250))+'px';tip.style.top=Math.max(8,Math.min(e.clientY+14,innerHeight-130))+'px'}
['mousemove','click'].forEach(t=>$('#tb').addEventListener(t,showTip));$('#tb').addEventListener('mouseleave',()=>tip.style.opacity=0);
const num=v=>+String(v).replace(/[^0-9.\-]/g,'')||0;
function grade(v){const s=String(v).toUpperCase();return /1ST|^1|\bA\b/.test(s)?'A':/2ND|^2|\bB\b/.test(s)?'B':/3RD|^3|\bC\b/.test(s)?'C':''}
function imsg(t){$('#im').textContent=t}
function imp(f){if(typeof XLSX=='undefined'){imsg('Excel library failed to load.');return}if(!PER.m){imsg('Pick a month first to import weekly incomes.');return}const r=new FileReader();r.onload=ev=>{try{
const wb=XLSX.read(ev.target.result,{type:'array'});
const K=s=>String(s==null?'':s).toUpperCase().replace(/[^A-Z0-9]/g,''),H=s=>String(s==null?'':s).toLowerCase().replace(/[^a-z0-9]/g,''),sp1=s=>String(s==null?'':s).replace(/\s+/g,' ').trim().toUpperCase();
const isBr=h=>/^(branch|branchname|branches|branchlocation|store|outlet)$/.test(h);
const areaOf=v=>{const k=K(v);if(!k)return null;return AO.find(a=>K(a)==k||K(a)==k+'AREA')||null};
let add=0,upd=0,skip=0,sheets=0;const noArea=[],weeks=new Set();
wb.SheetNames.forEach(sn=>{
const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:true,blankrows:false});
const hi=rows.findIndex((row,i)=>i<40&&row.some(c=>isBr(H(c))));if(hi<0)return;sheets++;
const col={wk:{},sr:{}};
rows[hi].map(H).forEach((h,c)=>{if(!h)return;
if(col.b==null&&isBr(h))col.b=c;
else if(col.ar==null&&/^area/.test(h))col.ar=c;
else if(col.m==null&&/mechanic/.test(h)&&!/perf|income/.test(h))col.m=c;
else if(col.c==null&&/^(classification|grade|class|branchclass)/.test(h))col.c=c;
else if(col.t==null&&/target/.test(h)&&!/ideal|daily|spare|need/.test(h))col.t=c;
else if(/repair/.test(h)){const mm=h.match(/(?:week|wk|w)(\d)/)||h.match(/(\d)(?:st|nd|rd|th)?(?:week|wk)/);if(mm){const k=+mm[1]-1;if(k>=0&&k<5&&col.sr[k]==null)col.sr[k]=c}}
else if(!/daily|perf|percent|target|achieve|ideal/.test(h)){const mm=h.match(/(?:week|wk|w)(\d)/)||h.match(/(\d)(?:st|nd|rd|th)?(?:week|wk)/);if(mm){const k=+mm[1]-1;if(k>=0&&k<5&&col.wk[k]==null)col.wk[k]=c}}});
let lastArea=null;
rows.slice(hi+1).forEach(row=>{
const nm=sp1(row[col.b]);if(!nm){skip++;return}
if(/^(TOTAL|GRAND TOTAL|OVERALL|SUBTOTAL)/.test(nm)){skip++;return}
const cell=c=>c==null||row[c]===''||row[c]==null?undefined:row[c];
const dataCols=[col.m,col.c,col.t,...Object.values(col.wk)];
if(/AREA$/.test(nm)&&areaOf(nm)&&!dataCols.some(c=>cell(c)!==undefined)){lastArea=areaOf(nm);skip++;return}
const arv=cell(col.ar);let a=null;
if(arv!==undefined){a=areaOf(arv);if(a)lastArea=a}else if(col.ar!=null&&lastArea)a=lastArea;else if(col.ar==null&&lastArea)a=lastArea;
const canon=N.find(n=>K(n)==K(nm)),b=canon||nm;
let x=B.find(y=>K(y.b)==K(b));const isNew=!x;
if(isNew){x={b,m:0,t:0,a:0,ar:AO[0],c:'',w:[0,0,0,0,0],wp:{}};sx(x);if(!a){const hit=AO.find(k=>AREAS[k].some(n=>K(n)==K(b)));if(hit)x.ar=hit;else noArea.push(b)}}
if(a)x.ar=a;
const mc=cell(col.m),t=cell(col.t),cl=cell(col.c);
if(mc!==undefined)x.m=num(mc);if(t!==undefined)x.t=num(t);
if(cl!==undefined){const g=grade(cl);if(g)x.c=g}
for(let k=0;k<5;k++){if(col.wk[k]==null)continue;const v=cell(col.wk[k]);if(v!==undefined){x.wi[k]=Math.max(0,num(v));weeks.add(k+1)}}
for(let k=0;k<5;k++){if(col.sr[k]==null)continue;const v=cell(col.sr[k]);if(v!==undefined)x.sw[k]=Math.max(0,Math.round(num(v)))}
if(isNew){B.push(x);add++}else upd++;rc(x)})});
if(!sheets){imsg('Could not find a "Branch" column header in the file.');return}
refreshAll();
const miss=[1,2,3,4,5].filter(k=>!weeks.has(k));
imsg(`Imported from ${sheets} sheet${sheets>1?'s':''}: ${upd} updated, ${add} added, ${skip} rows skipped (blank/total/area-title).`+(miss.length?` No data found for Week ${miss.join(', ')}.`:'')+(noArea.length?` Area not set for: ${noArea.slice(0,5).join(', ')}${noArea.length>5?'…':''}.`:''))
}catch(e){imsg('Could not read that file: '+(e&&e.message||e))}};r.readAsArrayBuffer(f)}
$('#xf').onchange=e=>{const f=e.target.files[0];if(f)imp(f);e.target.value=''};
function bpData(){const H=['Branch','Area','No. of Mechanics','Classification','Monthly Target','Ideal Daily Service Income'];
for(let k=1;k<=5;k++)H.push('Week '+k+' Income','Week '+k+' Daily Service Income','Week '+k+' Service Repaired','Week '+k+' Performance (%)');
H.push('Total Week Income','Total Week Daily Service Income','Total Week Service Repaired','Total Performance (%)','Need to Achieve');
const R=x=>{const s=x.wi.reduce((a,b)=>a+b,0),r=[x.b,x.ar,x.m,GR[x.c]||'',TG(x),Math.round(x.d)];x.wi.forEach((w,k)=>r.push(w,Math.round(w/(WD[k]*pm())),(x.sw&&x.sw[k])||0,x.t?+(w/TG(x)*100).toFixed(1):0));r.push(s,Math.round(s/(26*pm())),(x.sw||[]).reduce((a,b)=>a+(+b||0),0),x.t?+(s/TG(x)*100).toFixed(1):0,Math.max(0,TG(x)-s));return r};
const T={b:'TOTAL',ar:'',m:B.reduce((s,x)=>s+x.m,0),c:'',t:B.reduce((s,x)=>s+x.t,0),ty:B.reduce((s,x)=>s+x.ty,0),d:B.reduce((s,x)=>s+x.d,0),wi:[0,1,2,3,4].map(k=>B.reduce((s,x)=>s+x.wi[k],0)),sw:[0,1,2,3,4].map(k=>B.reduce((s,x)=>s+((x.sw&&x.sw[k])||0),0))};
return{H,rows:[...B.map(R),R(T)]}}
async function expXlsx(){const dl=window.claude?await claude.use('downloads'):null;if(!dl){imsg('Download is not available in this view.');return}
if(typeof XLSX=='undefined'){imsg('Excel library failed to load.');return}
const{H,rows}=bpData();const ws=XLSX.utils.aoa_to_sheet([H,...rows]);ws['!cols']=H.map((h,i)=>({wch:i?Math.max(14,h.length+2):22}));
const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Branch Performance');
const out=XLSX.write(wb,{bookType:'xlsx',type:'array'});
try{await dl.save({filename:'branch-performance-'+(PER.m?pk(PER.m):PER.y)+'.xlsx',data:new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})});imsg('Excel file ready.')}catch(e){imsg(e&&e.code=='declined'?'Download cancelled.':'Could not download the file.')}}
async function expPdf(){const dl=window.claude?await claude.use('downloads'):null;if(!dl){imsg('Download is not available in this view.');return}
const J=window.jspdf&&window.jspdf.jsPDF;if(!J){imsg('PDF library failed to load.');return}
try{const{H,rows}=bpData(),doc=new J({orientation:'landscape',unit:'mm',format:'a3'});if(typeof doc.autoTable!='function'){imsg('PDF table library failed to load.');return}
const lab=PER.m?MN[PER.m-1]+' '+PER.y:'Year '+PER.y,fm=(v,h)=>/Performance/.test(h)?v+'%':/Income|Target|Need to Achieve/.test(h)?(/Week/.test(h)&&!v?'-':'PHP '+Math.round(v).toLocaleString('en-PH')):/Service Repaired/.test(h)&&!v?'-':String(v);
doc.setFontSize(14);doc.text('Branch Performance - '+lab,6,10);
doc.autoTable({head:[H],body:rows.map(r=>r.map((v,j)=>fm(v,H[j]))),startY:14,theme:'grid',styles:{fontSize:5.2,cellPadding:.9,halign:'center',valign:'middle',overflow:'linebreak'},headStyles:{fillColor:[20,40,75],textColor:255,fontSize:5,halign:'center'},columnStyles:{0:{halign:'left',fontStyle:'bold'}},margin:{left:6,right:6,top:10,bottom:8},didParseCell:d=>{if(d.section=='body'&&d.row.index==rows.length-1){d.cell.styles.fontStyle='bold';d.cell.styles.fillColor=[230,235,244]}}});
const out=doc.output('arraybuffer');
await dl.save({filename:'branch-performance-'+(PER.m?pk(PER.m):PER.y)+'.pdf',data:new Blob([out],{type:'application/pdf'})});imsg('PDF file ready.')}catch(e){imsg(e&&e.code=='declined'?'Download cancelled.':'Could not create the PDF file.')}}
let dlT='bp';const dlOpen=t=>{dlT=t;$('#dlSub').textContent='Choose the file format for the '+(t=='sp'?'Spare Parts Report':'Branch Performance')+' table.';$('#dlDlg').showModal()};
$('#xe').onclick=()=>dlOpen('bp');
$('#dlC').onclick=()=>$('#dlDlg').close();
$('#dlX').onclick=()=>{$('#dlDlg').close();dlT=='sp'?spExp():expXlsx()};
$('#dlP').onclick=()=>{$('#dlDlg').close();dlT=='sp'?spPdf():expPdf()};
function openPrintPreview(sp){const src=document.querySelector(sp?'#spCard table':'#bpCard table');if(!src)return;const t=src.cloneNode(true);t.removeAttribute('style');t.querySelectorAll('input').forEach(i=>i.replaceWith(document.createTextNode(i.value)));
const nm=sp?'Spare Parts Report':'Branch Performance';$('#ppBody').classList.toggle('sp',!!sp);$('#ppName').textContent=nm;$('#ppTitle').textContent=nm+' · '+(PER.m?MN[PER.m-1]+' '+PER.y:'Year '+PER.y);$('#ppTable').innerHTML='';$('#ppTable').appendChild(t);$('#ppDlg').showModal()}
$('#xp').onclick=()=>openPrintPreview(0);$('#spxp').onclick=()=>openPrintPreview(1);
$('#ppClose').onclick=()=>$('#ppDlg').close();
$('#ppPrint').onclick=()=>{$('#ppDlg').close();try{window.print()}catch(e){imsg('Printing is not available in this view.')}};


let SPW={};try{const p=JSON.parse(localStorage.getItem('sparePartsWeekly'));if(p&&typeof p=='object'&&!Array.isArray(p))SPW=p}catch(e){}
const spSave=()=>{try{localStorage.setItem('sparePartsWeekly',JSON.stringify(SPW))}catch(e){}};
let SPTM={};try{const p=JSON.parse(localStorage.getItem('sparePartsTargetsMonthly'));if(p&&typeof p=='object'&&!Array.isArray(p))SPTM=p}catch(e){}
const spM=(b,y,m)=>{const o=SPTM[y+'-'+String(m).padStart(2,'0')];return o&&o[b]!=null?(+o[b]||0):(+SPT[b]||0)};
const spMo=b=>PER.m?spM(b,PER.y,PER.m):(+SPT[b]||0);
const spTarget=b=>PER.m?spM(b,PER.y,PER.m):[1,2,3,4,5,6,7,8,9,10,11,12].reduce((n,m)=>n+spM(b,PER.y,m),0);
const spWeeks=b=>{if(PER.m){const a=SPW[pk(PER.m)]&&SPW[pk(PER.m)][b];return[0,1,2,3,4].map(i=>+(a&&a[i])||0)}return[0,1,2,3,4].map(i=>{let s=0;for(let m=1;m<=12;m++){const a=SPW[pk(m)]&&SPW[pk(m)][b];if(a)s+=+a[i]||0}return s})};
const spTot=(t,s)=>[s?peso(s):'—',s?peso(s/(26*pm())):'—',pc(t?s/t*100:0,s),t?(t>s?peso(t-s):'✔ Achieved'):'—'];
function fillSpPer(){const y0=new Date().getFullYear();$('#sys').innerHTML=[y0-2,y0-1,y0,y0+1,y0+2].map(y=>`<option${y==PER.y?' selected':''}>${y}</option>`).join('');$('#sms').innerHTML='<option value="0">Whole year</option>'+MN.map((n,i)=>`<option value="${i+1}"${i+1==PER.m?' selected':''}>${n}</option>`).join('');$('#sys').onchange=$('#sms').onchange=()=>{PER.y=+$('#sys').value;PER.m=+$('#sms').value;refreshAll(1)}}
function spFoot(){const T=spL().map(b=>({t:spTarget(b),w:spWeeks(b),i:spMo(b)/D})),sum=f=>T.reduce((a,x)=>a+f(x),0),tt=sum(x=>x.t),ws=[0,1,2,3,4].map(k=>sum(x=>x.w[k])),ts=ws.reduce((a,c)=>a+c,0);
let h=`<tr><td>TOTAL</td><td></td><td class="num">${tt?peso(tt):'—'}</td><td class="num">${peso(sum(x=>x.i))}</td>`;ws.forEach((w,k)=>{h+=`<td class="num">${w?peso(w):'—'}</td><td class="num">${w?peso(w/(WD[k]*pm())):'—'}</td><td>${pc(tt?w/tt*100:0,w)}</td>`});
$('#spinvf').innerHTML=h+spTot(tt,ts).map(v=>`<td class="num">${v}</td>`).join('')+'</tr>'}
const spCls=b=>{const br=B.find(x=>x.b==b);return SPC[b]||(br?(br.dc??br.c):'')};
function spL(){const sel=$('#spb1').value,q=$('#spq').value.toLowerCase(),c=$('#spc1').value;return SPB.filter(b=>(!sel||b==sel)&&b.toLowerCase().includes(q)&&(!c||spCls(b)==c))}
function fillSpBr(){const e=$('#spb1'),v=e.value;e.innerHTML='<option value="">All branches</option>'+SPB.map(b=>`<option value="${esc(b)}">${esc(b)}</option>`).join('');e.value=SPB.includes(v)?v:''}
const spChW=()=>{const w=$('#spch').clientWidth;return w>0?Math.max(520,Math.min(1400,w)):720};
function drawSpCh(){
 if(!SPB.length){$('#spch').innerHTML='<p class="sub">No spare parts branches yet.</p>';return}
 const L=spL(),isM=!!PER.m,prevYear=PER.y-1,key=(y,m)=>y+'-'+String(m).padStart(2,'0'),sel=$('#spb1').value,active=sel||$('#spq').value||$('#spc1').value,
 label=sel?sel:active?'Filtered Branches ('+L.length+')':'All Branches';
 $('#spchT').textContent=(isM?'Weekly Income · '+MN[PER.m-1]+' '+PER.y:'Total Week Income by Month · '+PER.y)+' · Spare Parts · '+label+' (₱)';
 if(!L.length){$('#spch').innerHTML='<p class="sub">No branches match the current filters.</p>';return}
 const sum=a=>a.reduce((n,v)=>n+v,0),wk=(b,y,m)=>(SPW[key(y,m)]&&SPW[key(y,m)][b])||[],mt=sum(L.map(b=>spMo(b)));
 let labels,full,current,previous,scaleExtra,curName,prevName,cards;
 if(isM){
  const ws=y=>[0,1,2,3,4].map(i=>L.reduce((t,b)=>t+(+wk(b,y,PER.m)[i]||0),0));
  labels=[1,2,3,4,5].map(i=>'Week '+i);full=labels;current=ws(PER.y);previous=ws(prevYear);scaleExtra=[];
  curName=MN[PER.m-1]+' '+PER.y;prevName=MN[PER.m-1]+' '+prevYear;
  const cs=sum(current),ps=sum(previous),pf=mt?cs/mt*100:0;
  cards=[['Total Income · '+curName,peso(cs)],['Total Income · '+prevName,peso(ps)],['Monthly Target',mt?peso(mt):'—'],['Performance',cs&&mt?pf.toFixed(1)+'%':'—']];
 }else{
  const months=Array.from({length:12},(_,i)=>i+1),tot=y=>months.map(m=>L.reduce((t,b)=>t+sum(wk(b,y,m).map(n=>+n||0)),0));
  labels=MN.map(m=>m.slice(0,3));full=MN;current=tot(PER.y);previous=tot(prevYear);scaleExtra=[mt];
  curName=String(PER.y);prevName=String(prevYear);
  const bi=current.indexOf(Math.max(...current));
  cards=[[PER.y+' Total Week Income',peso(sum(current))],[prevYear+' Total Week Income',peso(sum(previous))],['Top month in '+PER.y,current[bi]>0?MN[bi]:'—']];
 }
 if(!current.concat(previous).some(v=>v>0)){$('#spch').innerHTML='<p class="sub">No income recorded for '+esc(label)+' in '+(isM?curName+' or '+prevName:'these years')+'.</p>';return}
 const n=labels.length,short=v=>v>=1000000?'₱'+(v/1000000).toFixed(1)+'m':v>=1000?'₱'+(v/1000).toFixed(0)+'k':peso(v),
 W=spChW(),H=Math.max(300,Math.min(480,Math.round(W*.38))),left=70,right=44,top=30,bottom=50,pw=W-left-right,ph=H-top-bottom,mx=Math.max(1,...scaleExtra,...current,...previous),px=i=>left+i*pw/(n-1),py=v=>top+ph-v/mx*ph,base=top+ph;
 let s=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;max-width:100%;height:auto;display:block" role="img" aria-label="Spare parts income for ${esc(label)}: ${curName} and ${prevName}">`;
 for(let i=0;i<=4;i++){const v=mx*(4-i)/4,y=top+i*ph/4;s+=`<line x1="${left}" y1="${y}" x2="${W-right}" y2="${y}" stroke="var(--bd)" stroke-dasharray="4 6"/><text x="${left-12}" y="${y+4}" text-anchor="end" font-size="11" fill="var(--mu)">${short(v)}</text>`}
 const pathFor=a=>a.map((v,i)=>`${i?'L':'M'} ${px(i)} ${py(v)}`).join(' ');
 s+=`<path d="${pathFor(previous)}" fill="none" stroke="var(--ac2)" stroke-width="2.5" stroke-dasharray="8 6" stroke-linecap="round" stroke-linejoin="round"/><path d="${pathFor(current)}" fill="none" stroke="var(--ac)" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>`;
 previous.forEach((v,i)=>{s+=`<g><title>${full[i]} ${prevName}: ${peso(v)} · ${esc(label)}</title><circle cx="${px(i)}" cy="${py(v)}" r="4.5" fill="var(--card)" stroke="var(--ac2)" stroke-width="2.5"/></g>`});
 current.forEach((v,i)=>{const x=px(i),y=py(v),ly=Math.max(top+13,y-12);s+=`<g><title>${full[i]} ${curName}: ${peso(v)} · ${esc(label)}</title><circle cx="${x}" cy="${y}" r="5" fill="var(--ac)" stroke="var(--card)" stroke-width="2"/><text x="${Math.min(W-24,Math.max(left,x))}" y="${ly}" text-anchor="middle" font-size="10" font-weight="700" fill="var(--tx)">${short(v)}</text></g>`});
 labels.forEach((m,i)=>{const x=px(i);s+=`<line x1="${x}" y1="${base}" x2="${x}" y2="${base+5}" stroke="var(--bd)"/><text x="${x}" y="${base+22}" text-anchor="middle" font-size="11" fill="var(--mu)">${m}</text>`});s+='</svg>';
 $('#spch').innerHTML='<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px;margin:2px 0 12px">'+cards.map(c=>'<div class="card" style="padding:10px 12px"><div class="sub">'+c[0]+'</div><b>'+c[1]+'</b></div>').join('')+'</div><div style="width:100%;overflow:hidden">'+s+'</div><div style="display:flex;gap:18px;flex-wrap:wrap;margin-top:10px"><span><i style="display:inline-block;width:18px;border-top:3px solid var(--ac);vertical-align:middle;margin-right:6px"></i>'+curName+'</span><span><i style="display:inline-block;width:18px;border-top:3px dashed var(--ac2);vertical-align:middle;margin-right:6px"></i>'+prevName+'</span></div><div class="sub" style="margin-top:8px">Showing '+esc(label)+'. Use the Year, Month, Branch, Search and Class filters in Spare Parts Report to change this graph. Pick a month to see its weekly income; choose Whole year to see all 12 months. Hover a point for its exact amount.</div>';
}
let spLastW=0;let spRaf=0;new ResizeObserver(()=>{cancelAnimationFrame(spRaf);spRaf=requestAnimationFrame(()=>{const w=$('#spch').clientWidth;if(w&&Math.abs(w-spLastW)>8){spLastW=w;drawSpCh()}})}).observe($('#spch'));
function drawPartsReport(){fillSpBr();$('#sys').value=PER.y;$('#sms').value=PER.m;$('#ys').value=PER.y;$('#ms').value=PER.m;
$('#spinvb').innerHTML=SPB.map((b,i)=>[b,i]).filter(([b])=>spL().includes(b)).map(([b,i])=>{const br=B.find(x=>x.b==b),c=SPC[b]||(br?(br.dc??br.c):''),t=spTarget(b),w=spWeeks(b),s=w.reduce((a,v)=>a+v,0);
let h=`<tr><td><b>${esc(b)}</b></td><td>${c?`<span class="tag ${c}">CLASS ${c}</span>`:'<span class="sub">—</span>'}</td><td class="num">${t?peso(t):'<span class="sub">—</span>'}</td><td class="num">${peso(spMo(b)/D)}</td>`;
w.forEach((v,k)=>{h+=`<td><input class="wv sw" type="text" inputmode="decimal" data-i="${i}" data-k="${k}" value="${v?peso(v):''}" ${PER.m?'':'readonly'} placeholder="₱0"></td><td class="num sdy" data-i="${i}" data-k="${k}">${v?peso(v/(WD[k]*pm())):'—'}</td><td class="spf" data-i="${i}" data-k="${k}">${pc(t?v/t*100:0,v)}</td>`});
return h+spTot(t,s).map((v,j)=>`<td class="num stt" data-i="${i}" data-j="${j}">${v}</td>`).join('')+'</tr>'}).join('')||'<tr><td colspan="23" class="sub">No branches match the current filters.</td></tr>';
spFoot();drawSpCh();
document.querySelectorAll('#spinvb .sw').forEach(e=>{e.onbeforeinput=ev=>{if(ev.data&&/[^0-9,.₱]/.test(ev.data)){ev.preventDefault();alert('Input Numbers Only')}};
e.oninput=()=>{if(/[^0-9,.₱]/.test(e.value)){e.value=e.value.replace(/[^0-9,.₱]/g,'');alert('Input Numbers Only')}
const i=+e.dataset.i,k=+e.dataset.k,b=SPB[i];if(!PER.m)return;const key=pk(PER.m),o=SPW[key]||(SPW[key]={}),a=o[b]||(o[b]=[0,0,0,0,0]);a[k]=Math.max(0,num(e.value));spSave();
const t=spTarget(b),w=spWeeks(b),sel=`[data-i="${i}"][data-k="${k}"]`;document.querySelector('#spinvb .sdy'+sel).textContent=w[k]?peso(w[k]/(WD[k]*pm())):'—';document.querySelector('#spinvb .spf'+sel).innerHTML=pc(t?w[k]/t*100:0,w[k]);
const tot=spTot(t,w.reduce((x,y)=>x+y,0));document.querySelectorAll(`#spinvb .stt[data-i="${i}"]`).forEach(c=>c.innerHTML=tot[c.dataset.j]);spFoot();drawSpCh()};
e.onblur=()=>{const i=+e.dataset.i,k=+e.dataset.k;const v=spWeeks(SPB[i])[k];e.value=v?peso(v):''}})}
const spmsg=t=>{$('#spim').textContent=t};
function spImp(f){if(typeof XLSX=='undefined'){spmsg('Excel library failed to load.');return}const r=new FileReader();r.onload=ev=>{try{
const wb=XLSX.read(ev.target.result,{type:'array'}),K=s=>String(s==null?'':s).toUpperCase().replace(/[^A-Z0-9]/g,''),H=s=>String(s==null?'':s).toLowerCase().replace(/[^a-z0-9]/g,'');
let upd=0,skip=0,sheets=0,hasWk=false;const unk=[];
wb.SheetNames.forEach(sn=>{const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:true,blankrows:false});
const hi=rows.findIndex((row,i)=>i<40&&row.some(c=>/^(branch|branchname|branches|store|outlet)$/.test(H(c))));if(hi<0)return;sheets++;
const col={wk:{}};rows[hi].map(H).forEach((h,c)=>{if(!h)return;if(col.b==null&&/^(branch|branchname|branches|store|outlet)$/.test(h))col.b=c;else if(col.c==null&&/^(classification|grade|class)/.test(h))col.c=c;else if(col.t==null&&/target/.test(h)&&!/ideal|daily/.test(h))col.t=c;else if(!/daily|perf|percent|target|achieve|ideal/.test(h)){const mm=h.match(/(?:week|wk|w)(\d)/)||h.match(/(\d)(?:st|nd|rd|th)?(?:week|wk)/);if(mm){const k=+mm[1]-1;if(k>=0&&k<5&&col.wk[k]==null){col.wk[k]=c;hasWk=true}}}});
rows.slice(hi+1).forEach(row=>{const nm=String(row[col.b]==null?'':row[col.b]).replace(/\s+/g,' ').trim().toUpperCase();if(!nm||/^(TOTAL|GRAND TOTAL)/.test(nm)){skip++;return}
const b=SPB.find(n=>K(n)==K(nm));if(!b){unk.push(nm);return}
const cl=col.c!=null&&row[col.c]!==''?row[col.c]:undefined,t=col.t!=null&&row[col.t]!==''?row[col.t]:undefined;
if(cl!==undefined){const g=grade(cl);if(g)SPC[b]=g}if(t!==undefined)SPT[b]=Math.max(0,num(t));if(PER.m)for(let k=0;k<5;k++){if(col.wk[k]==null)continue;const v=row[col.wk[k]];if(v!==''&&v!=null){const key=pk(PER.m),o=SPW[key]||(SPW[key]={}),a=o[b]||(o[b]=[0,0,0,0,0]);a[k]=Math.max(0,num(v))}}upd++})});
if(!sheets){spmsg('Could not find a "Branch" column header in the file.');return}
try{localStorage.setItem('sparePartsTargets',JSON.stringify(SPT));localStorage.setItem('sparePartsClassifications',JSON.stringify(SPC))}catch(e){}
spSave();drawPartsTargets();drawPartsKpis();anim();
spmsg(`Imported: ${upd} branches updated, ${skip} rows skipped.`+(hasWk&&!PER.m?' Weekly incomes were not imported: pick a month first.':'')+(unk.length?` Not in the Spare Parts branch list: ${unk.slice(0,5).join(', ')}${unk.length>5?'…':''}.`:''))
}catch(e){spmsg('Could not read that file: '+(e&&e.message||e))}};r.readAsArrayBuffer(f)}
$('#spxf').onchange=e=>{const f=e.target.files[0];if(f)spImp(f);e.target.value=''};
function spData(){const H=['Branch','Classification','Monthly Spare Parts Target','Ideal Daily Spare Parts Income'];for(let k=1;k<=5;k++)H.push('Week '+k+' Income','Week '+k+' Daily Spare Parts Income','Week '+k+' Performance (%)');H.push('Total Week Income','Total Week Daily Spare Parts Income','Total Performance (%)','Need to Achieve');
const row=(n,c,t,id,w)=>{const s=w.reduce((a,v)=>a+v,0),r=[n,c,t,Math.round(id)];w.forEach((v,k)=>r.push(v,Math.round(v/(WD[k]*pm())),t?+(v/t*100).toFixed(1):0));r.push(s,Math.round(s/(26*pm())),t?+(s/t*100).toFixed(1):0,t?Math.max(0,t-s):0);return r};
const R=SPB.map(b=>{const br=B.find(x=>x.b==b),c=SPC[b]||(br?(br.dc??br.c):'');return row(b,c?'CLASS '+c:'',spTarget(b),spMo(b)/D,spWeeks(b))});
const tt=SPB.reduce((n,b)=>n+spTarget(b),0),ti=SPB.reduce((n,b)=>n+spMo(b)/D,0),tw=[0,1,2,3,4].map(k=>SPB.reduce((n,b)=>n+spWeeks(b)[k],0));
return{H,rows:[...R,row('TOTAL','',tt,ti,tw)]}}
async function spExp(){const dl=window.claude?await claude.use('downloads'):null;if(!dl){spmsg('Download is not available in this view.');return}
if(typeof XLSX=='undefined'){spmsg('Excel library failed to load.');return}
const{H,rows}=spData();const ws=XLSX.utils.aoa_to_sheet([H,...rows]);ws['!cols']=H.map((h,i)=>({wch:i?Math.max(14,h.length+2):22}));
const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Spare Parts Report');const out=XLSX.write(wb,{bookType:'xlsx',type:'array'});
try{await dl.save({filename:'spare-parts-report-'+(PER.m?pk(PER.m):PER.y)+'.xlsx',data:new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})});spmsg('Excel file ready.')}catch(e){spmsg(e&&e.code=='declined'?'Download cancelled.':'Could not download the file.')}}
async function spPdf(){const dl=window.claude?await claude.use('downloads'):null;if(!dl){spmsg('Download is not available in this view.');return}
const J=window.jspdf&&window.jspdf.jsPDF;if(!J){spmsg('PDF library failed to load.');return}
try{const{H,rows}=spData(),doc=new J({orientation:'landscape',unit:'mm',format:'a3'});if(typeof doc.autoTable!='function'){spmsg('PDF table library failed to load.');return}
const lab=PER.m?MN[PER.m-1]+' '+PER.y:'Year '+PER.y,fm=(v,h)=>/Performance/.test(h)?v+'%':/Income|Target|Need to Achieve/.test(h)?(/Week/.test(h)&&!v?'-':'PHP '+Math.round(v).toLocaleString('en-PH')):String(v);
doc.setFontSize(14);doc.text('Spare Parts Report - '+lab,6,10);
doc.autoTable({head:[H],body:rows.map(r=>r.map((v,j)=>fm(v,H[j]))),startY:14,theme:'grid',styles:{fontSize:5.2,cellPadding:.9,halign:'center',valign:'middle',overflow:'linebreak'},headStyles:{fillColor:[20,40,75],textColor:255,fontSize:5,halign:'center'},columnStyles:{0:{halign:'left',fontStyle:'bold'}},margin:{left:6,right:6,top:10,bottom:8},didParseCell:d=>{if(d.section=='body'&&d.row.index==rows.length-1){d.cell.styles.fontStyle='bold';d.cell.styles.fillColor=[230,235,244]}}});
const out=doc.output('arraybuffer');
await dl.save({filename:'spare-parts-report-'+(PER.m?pk(PER.m):PER.y)+'.pdf',data:new Blob([out],{type:'application/pdf'})});spmsg('PDF file ready.')}catch(e){spmsg(e&&e.code=='declined'?'Download cancelled.':'Could not create the PDF file.')}}
$('#spxe').onclick=()=>dlOpen('sp');
// theme/nav
document.querySelectorAll('.tw').forEach(w=>{let d=0,x0=0,s0=0;w.addEventListener('pointerdown',e=>{if(e.pointerType!='mouse'||e.target.closest('input,select,button'))return;d=1;x0=e.clientX;s0=w.scrollLeft;w.classList.add('drag')});addEventListener('pointermove',e=>{if(d)w.scrollLeft=s0-(e.clientX-x0)});addEventListener('pointerup',()=>{d=0;w.classList.remove('drag')})});
$('#hz').onclick=()=>$('.app').classList.toggle('open');
document.querySelector('main').addEventListener('click',()=>$('.app').classList.remove('open'));
$('#dt').textContent=new Date().toLocaleDateString('en-PH',{weekday:'long',year:'numeric',month:'long',day:'numeric'});
$('#navToggle').onclick=e=>{const hidden=$('.app').classList.toggle('nav-hidden'),b=e.currentTarget;b.setAttribute('aria-label',hidden?'Show sidebar':'Hide sidebar');b.title=hidden?'Show sidebar':'Hide sidebar'};$('#th').onclick=()=>{const r=document.documentElement,d=r.dataset.theme!=='light';r.dataset.theme=d?'light':'dark'};
const T={svc:'Service Mechanic Performance',asr:'After Sales Report',parts:'Spare Parts Monitoring',fac:'Admin · Branch Facility Concerns',mgr:'Manage Branches',spbr:'Spare Parts Branches','ad-memo':'Memo Preview','ad-unif':'Uniform Monitoring','ad-reno':'Branch Renovation Report','ad-trav':'Travel Expense','ad-cont':'Contract Preview'};
document.querySelectorAll('.nav button').forEach(b=>b.onclick=()=>{document.querySelectorAll('.nav button,.sec').forEach(e=>e.classList.remove('on'));b.classList.add('on');$('#'+b.dataset.s).classList.add('on');$('#ttl').textContent=T[b.dataset.s];if(b.dataset.s=='asr')drawAsr();anim()});
// kpi
function kpi(el,a){el.innerHTML=a.map(([l,v,d,f,c])=>`<div class="card kpi"><div class="l">${l}</div><div class="v" data-v="${v}" data-f="${f||''}">0</div><div class="d ${c||''}">${d}</div></div>`).join('')}
function anim(){document.querySelectorAll('.sec.on .v').forEach(e=>{const v=+e.dataset.v,f=e.dataset.f,t0=performance.now();(function s(t){const p=Math.min(1,(t-t0)/900),x=v*(1-Math.pow(1-p,3));e.textContent=f=='s'?e.dataset.v:f=='p'?peso(x):f=='%'?x.toFixed(1)+'%':Math.round(x);if(p<1)requestAnimationFrame(s)})(t0)});
 setTimeout(()=>document.querySelectorAll('.sec.on .bar i').forEach(i=>i.style.width=i.dataset.w+'%'),50)}
function drawK1(){const tT=B.reduce((s,x)=>s+TG(x),0),byArea=new Map();B.forEach(x=>{const daily=x.wi.reduce((s,w)=>s+w,0)/(26*pm());byArea.set(x.ar,(byArea.get(x.ar)||0)+daily)});const top=[...byArea].filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1])[0],area=top?top[0]:'—',daily=top?top[1]:0;kpi($('#k1'),[['Total Mechanics',B.reduce((s,x)=>s+x.m,0),B.length+' branches'],['Monthly Target',tT,'Combined all branches','p'],['Top Performing Area',area,daily?'Total week daily service income · '+peso(daily):'No weekly income yet','s'],['Branches',B.length,'Total branches']]);drawNotices()}
function drawPartsKpis(){const total=SPB.reduce((s,b)=>s+spMo(b),0),set=SPB.filter(b=>spMo(b)>0).length;kpi($('#k2'),[['Spare Parts Branches',SPB.length,'TOTAL 3S &amp; MS SHOPS'],['Targets Set',set,'Branches with a target'],['Monthly Spare Parts Target',total,'Combined spare-parts target','p']])}
function fk(){}
// chart
function noticeRows(list){return list.map(({x,target,income})=>'<div style="display:flex;justify-content:space-between;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid var(--bd)"><div><b>'+esc(x.b)+'</b><div class="sub">'+esc(x.ar)+'</div></div><span class="sub">Needs '+peso(target-income)+' · '+peso(income)+' / '+peso(target)+'</span></div>').join('')}
function drawNotices(){if(!B.length){if($('#noticeDlg').open)$('#noticeDlg').close();$('#mix').innerHTML='<p class="sub">No branch data yet. Add a branch to begin.</p>';return}const L=B.map(x=>({x,target:TG(x),income:x.wi.reduce((s,w)=>s+w,0)})).filter(o=>o.target>o.income).sort((a,b)=>(b.target-b.income)-(a.target-a.income));if(!L.length){if($('#noticeDlg').open)$('#noticeDlg').close();$('#mix').innerHTML='<p class="sub">All branches have achieved the selected period target.</p>';return}const shown=L.slice(0,5),extra=L.slice(5);$('#mix').innerHTML='<p class="sub">Branches below the selected period target'+(L.length>5?' · Showing 5 of '+L.length:'')+':</p>'+noticeRows(shown)+(L.length>5?'<p><button type="button" class="tbtn" id="moreNotices">Show More</button></p>':'');if(L.length>5)$('#moreNotices').onclick=()=>{const dlg=$('#noticeDlg');$('#noticeDlgSub').textContent='Showing '+extra.length+' additional branches below the selected period target.';$('#noticeItems').innerHTML=noticeRows(extra);dlg.showModal()};if($('#noticeDlg').open&&L.length>5){$('#noticeDlgSub').textContent='Showing '+extra.length+' additional branches below the selected period target.';$('#noticeItems').innerHTML=noticeRows(extra)}}
$('#closeNotices').onclick=()=>$('#noticeDlg').close();
function drawCh(){
 if(!B.length){$('#ch').innerHTML='<p class="sub">No branch data yet. Add a branch to begin.</p>';$('#mix').innerHTML='<p class="sub">No branch data yet.</p>';return}
 const L=fB(),isM=!!PER.m,prevYear=PER.y-1,key=(y,m)=>y+'-'+String(m).padStart(2,'0'),wk=(x,y,m)=>(x.wp&&x.wp[key(y,m)])||[],sel=$('#b1').value,active=sel||$('#q1').value||$('#a1').value||$('#c1').value,
 label=sel?sel:active?'Filtered Branches ('+L.length+')':'All Branches';
 $('#chT').textContent=(isM?'Weekly Income · '+MN[PER.m-1]+' '+PER.y:'Monthly Income · '+PER.y)+' · '+label+' (₱)';
 if(!L.length){$('#ch').innerHTML='<p class="sub">No branches match the current filters.</p>';return}
 const tgtOf=(y,m)=>L.reduce((sum,x)=>{const q=x.mo&&x.mo[key(y,m)];return sum+(q&&q.t!=null?+q.t:+(x.dt??x.t??0))},0),sum=a=>a.reduce((n,v)=>n+v,0);
 let labels,full,current,previous,scaleExtra,curName,prevName,cards,tline=[],cum=null;
 if(isM){
  const ws=y=>[0,1,2,3,4].map(i=>L.reduce((t,x)=>t+(+wk(x,y,PER.m)[i]||0),0));
  labels=[1,2,3,4,5].map(i=>'Week '+i);full=labels;current=ws(PER.y);previous=ws(prevYear);cum=current.map((v,i)=>current.slice(0,i+1).reduce((a,b)=>a+b,0));tline=labels.map(()=>tgtOf(PER.y,PER.m));scaleExtra=tline.concat(cum);
  curName=MN[PER.m-1]+' '+PER.y;prevName=MN[PER.m-1]+' '+prevYear;
  const cs=sum(current),ps=sum(previous),t=tgtOf(PER.y,PER.m),pf=t?cs/t*100:0;
  cards=[['Total Income · '+curName,peso(cs)],['Total Income · '+prevName,peso(ps)],['Monthly Target',peso(t)],['Performance',cs?pf.toFixed(1)+'%':'—']];
 }else{
  const months=Array.from({length:12},(_,i)=>i+1),tot=y=>months.map(m=>L.reduce((t,x)=>t+sum(wk(x,y,m).map(n=>+n||0)),0));
  labels=MN.map(m=>m.slice(0,3));full=MN;current=tot(PER.y);previous=tot(prevYear);scaleExtra=months.map(m=>tgtOf(PER.y,m)).concat(months.map(m=>tgtOf(prevYear,m)));tline=months.map(m=>tgtOf(PER.y,m));
  curName=String(PER.y);prevName=String(prevYear);
  const bi=current.indexOf(Math.max(...current));
  cards=[['Total Income of Year '+PER.y,peso(sum(current))],['Total Income of Year '+prevYear,peso(sum(previous))],['Top month in '+PER.y,current[bi]>0?MN[bi]:'—'],['Months that reached the target',current.filter((v,i)=>tline[i]>0&&v>=tline[i]).length+' of '+tline.filter(v=>v>0).length]];
 }
 if(!current.concat(previous).some(v=>v>0)){$('#ch').innerHTML='<p class="sub">No income recorded for '+esc(label)+' in '+(isM?curName+' or '+prevName:'these years')+'.</p>';return}
 const n=labels.length,short=v=>v>=1000000?'₱'+(v/1000000).toFixed(1)+'m':v>=1000?'₱'+(v/1000).toFixed(0)+'k':peso(v),
 W=720,H=360,left=64,right=84,top=28,bottom=48,pw=W-left-right,ph=H-top-bottom,mx=Math.max(1,...scaleExtra,...current,...previous),px=i=>left+i*pw/(n-1),py=v=>top+ph-v/mx*ph,base=top+ph;
 let s=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;max-width:100%;height:auto;display:block" role="img" aria-label="Income comparison for ${esc(label)}: ${curName} and ${prevName}">`;
 for(let i=0;i<=4;i++){const v=mx*(4-i)/4,y=top+i*ph/4;s+=`<line x1="${left}" y1="${y}" x2="${W-right}" y2="${y}" stroke="var(--bd)" stroke-dasharray="4 6"/><text x="${left-12}" y="${y+4}" text-anchor="end" font-size="11" fill="var(--mu)">${short(v)}</text>`}
 const pathFor=a=>a.map((v,i)=>`${i?'L':'M'} ${px(i)} ${py(v)}`).join(' ');
 if(tline.some(v=>v>0)){s+=`<path d="${pathFor(tline)}" fill="none" stroke="#16a34a" stroke-width="2.5" stroke-dasharray="2 6" stroke-linecap="round"/>`;
  tline.forEach((v,i)=>{s+=`<g><title>${full[i]} Target: ${peso(v)} · ${esc(label)}</title><circle cx="${px(i)}" cy="${py(v)}" r="3.5" fill="#16a34a"/></g>`});
  const same=tline.every(v=>v===tline[0]),lt=tline[tline.length-1];
  s+=`<text x="${W-right+8}" y="${py(lt)-3}" font-size="10" font-weight="700" fill="#16a34a">Target</text><text x="${W-right+8}" y="${py(lt)+10}" font-size="10" font-weight="700" fill="#16a34a">${same?short(tline[0]):'varies'}</text>`}
 if(cum)s+=`<path d="${pathFor(cum)}" fill="none" stroke="var(--ac)" stroke-width="2" stroke-opacity=".55" stroke-dasharray="1 5" stroke-linecap="round"/>`+cum.map((v,i)=>{const ok=tline[i]>0&&v>=tline[i];return`<g><title>${full[i]} cumulative: ${peso(v)}${tline[i]?' · '+(v/tline[i]*100).toFixed(1)+'% of target':''}</title><circle cx="${px(i)}" cy="${py(v)}" r="4" fill="${ok?'#16a34a':'var(--card)'}" stroke="${ok?'#16a34a':'var(--ac)'}" stroke-width="2"/></g>`}).join('');
 s+=`<path d="${pathFor(previous)}" fill="none" stroke="var(--ac2)" stroke-width="2.5" stroke-dasharray="8 6" stroke-linecap="round" stroke-linejoin="round"/><path d="${pathFor(current)}" fill="none" stroke="var(--ac)" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>`;
 previous.forEach((v,i)=>{s+=`<g><title>${full[i]} ${prevName}: ${peso(v)} · ${esc(label)}</title><circle cx="${px(i)}" cy="${py(v)}" r="4.5" fill="var(--card)" stroke="var(--ac2)" stroke-width="2.5"/></g>`});
 current.forEach((v,i)=>{const x=px(i),y=py(v),ly=Math.max(top+13,y-12);s+=`<g><title>${full[i]} ${curName}: ${peso(v)} · ${esc(label)}${!isM&&tline[i]>0?(v>=tline[i]?' · target reached ✓':' · below target ('+(v/tline[i]*100).toFixed(1)+'%)'):''}</title><circle cx="${x}" cy="${y}" r="5" fill="${!isM&&tline[i]>0&&v>=tline[i]?'#16a34a':'var(--ac)'}" stroke="var(--card)" stroke-width="2"/><text x="${Math.min(W-24,Math.max(left,x))}" y="${ly}" text-anchor="middle" font-size="10" font-weight="700" fill="var(--tx)">${short(v)}</text></g>`});
 labels.forEach((m,i)=>{const x=px(i);s+=`<line x1="${x}" y1="${base}" x2="${x}" y2="${base+5}" stroke="var(--bd)"/><text x="${x}" y="${base+22}" text-anchor="middle" font-size="11" fill="var(--mu)">${m}</text>`});s+='</svg>';
 $('#ch').innerHTML='<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px;margin:2px 0 12px">'+cards.map(c=>'<div class="card" style="padding:10px 12px"><div class="sub">'+c[0]+'</div><b>'+c[1]+'</b></div>').join('')+'</div><div style="width:100%;overflow:hidden">'+s+'</div><div style="display:flex;gap:18px;flex-wrap:wrap;margin-top:10px"><span><i style="display:inline-block;width:18px;border-top:3px solid var(--ac);vertical-align:middle;margin-right:6px"></i>'+curName+'</span><span><i style="display:inline-block;width:18px;border-top:3px dashed var(--ac2);vertical-align:middle;margin-right:6px"></i>'+prevName+'</span>'+(tline.some(v=>v>0)?'<span><i style="display:inline-block;width:18px;border-top:3px dotted #16a34a;vertical-align:middle;margin-right:6px"></i>Monthly Target'+(isM?'':' (green dot = target reached)')+'</span>':'')+(cum?'<span><i style="display:inline-block;width:18px;border-top:3px dotted var(--ac);vertical-align:middle;margin-right:6px"></i>Running total of the weeks (green = target reached)</span>':'')+'</div><div class="sub" style="margin-top:8px">Showing '+esc(label)+'. Use the Year, Month, Branch, Area and Grade filters in Branch Performance to change this graph. Pick a month to see its weekly income; choose Whole year to see all 12 months. Hover a point for its exact amount.</div>';
}
// table
function fB(){const q=$('#q1').value.toLowerCase(),c=$('#c1').value,a=$('#a1').value,b=$('#b1').value;return B.filter(x=>(!b||x.b==b)&&x.b.toLowerCase().includes(q)&&(!c||x.c==c)&&(!a||x.ar==a))}
function fillBr(){const e=$('#b1'),v=e.value;e.innerHTML='<option value="">All branches</option>'+B.map(x=>`<option value="${esc(x.b)}">${esc(x.b)}</option>`).join('');e.value=B.some(x=>x.b==v)?v:''}
let sk='b',sd=1;
function rt(){const q=$('#q1').value.toLowerCase(),c=$('#c1').value,a=$('#a1').value;
$('#tb').innerHTML=fB().sort((a,b)=>(a[sk]>b[sk]?1:-1)*sd).map(x=>`<tr><td><b>${x.b}</b></td><td>${x.ar}</td><td class="num">${x.m}</td><td><span class="tag ${x.c}">${GR[x.c]}</span></td><td class="num">${peso(TG(x))}</td><td class="num">${peso(x.d)}</td>${wk(x)}</tr>`).join('');document.querySelectorAll('#tb .wi').forEach(e=>e.oninput=()=>{const x=B[e.dataset.i],k=+e.dataset.k;x.w[k]=Math.max(0,+e.value||0);document.querySelector(`#tb .pf[data-i="${e.dataset.i}"][data-k="${k}"]`).innerHTML=pc(perf(x,k),x.w[k]);sv()});document.querySelectorAll('#tb .wv').forEach(e=>{
 e.onbeforeinput=ev=>{if(ev.data&&/[^0-9,.₱]/.test(ev.data)){ev.preventDefault();alert('Input Numbers Only')}};
 e.oninput=()=>{const invalid=/[^0-9,.₱]/.test(e.value);if(invalid){e.value=e.value.replace(/[^0-9,.₱]/g,'');alert('Input Numbers Only')}const x=B[e.dataset.i],k=+e.dataset.k,s=`[data-i="${e.dataset.i}"][data-k="${k}"]`;x.wi[k]=Math.max(0,num(e.value));document.querySelector('#tb .dy'+s).textContent=dv(x,k);document.querySelector('#tb .pf'+s).innerHTML=pc(perf(x,k),x.wi[k]);document.querySelectorAll('#tb .tt[data-i="'+e.dataset.i+'"]').forEach(c=>c.innerHTML=tc(x)[c.dataset.j]);ft();drawK1();drawCh();anim();sv()};
 e.onblur=()=>{const x=B[e.dataset.i],k=+e.dataset.k;e.value=x.wi[k]?peso(x.wi[k]):''};
});document.querySelectorAll('#tb .sv').forEach(e=>{e.oninput=()=>{const c=e.value.replace(/[^0-9]/g,'');if(c!==e.value)e.value=c;const x=B[e.dataset.i],k=+e.dataset.k;if(!PER.m)return;x.sw[k]=parseInt(c,10)||0;document.querySelectorAll('#tb .tt[data-i="'+e.dataset.i+'"]').forEach(c=>c.innerHTML=tc(x)[c.dataset.j]);ft();sv()}});ft();anim()}
document.querySelectorAll('th[data-k]').forEach(h=>h.onclick=()=>{sd=sk==h.dataset.k?-sd:1;sk=h.dataset.k;rt()});
$('#q1').oninput=$('#c1').onchange=$('#a1').onchange=$('#b1').onchange=()=>{rt();drawCh()};

function drawPartsTargets(){const q=$('#qpt').value.toLowerCase(),visible=SPB.filter(b=>b.toLowerCase().includes(q)),render=(id,list,ro)=>{$('#'+id).innerHTML=list.map(b=>{const branch=B.find(x=>x.b==b),fallback=branch?(branch.dc??branch.c):'',c=SPC[b]||fallback,target=spMo(b);if(ro)return`<tr><td><b>${esc(b)}</b></td><td>${c?`<span class="tag ${c}">CLASS ${c}</span>`:'<span class="sub">—</span>'}</td><td class="num">${target?peso(target):'<span class="sub">—</span>'}</td><td class="num">${peso(target/D)}</td></tr>`;return`<tr><td><b>${esc(b)}</b></td><td><select class="pt-class tbtn ${c}" data-b="${esc(b)}" aria-label="Classification for ${esc(b)}"><option value="">Select classification</option>${['A','B','C'].map(k=>`<option value="${k}"${k==c?' selected':''}>CLASS ${k}</option>`).join('')}</select></td><td><input class="pt-target" type="number" min="0" step="1000" data-b="${esc(b)}" value="${target||''}" placeholder="₱" aria-label="Monthly spare-parts target for ${esc(b)}"></td><td class="num">${peso(target/D)}</td></tr>`}).join('')||'<tr><td colspan="4" class="sub">No branches found.</td></tr>'};render('sptb',visible);drawPartsReport();document.querySelectorAll('.pt-class').forEach(e=>e.onchange=()=>{SPC[e.dataset.b]=e.value;try{localStorage.setItem('sparePartsClassifications',JSON.stringify(SPC))}catch(e){}drawPartsTargets();anim()});document.querySelectorAll('.pt-target').forEach(e=>e.onchange=()=>{const v=Math.max(0,+e.value||0);if(PER.m){const k=PER.y+'-'+String(PER.m).padStart(2,'0');(SPTM[k]||(SPTM[k]={}))[e.dataset.b]=v;try{localStorage.setItem('sparePartsTargetsMonthly',JSON.stringify(SPTM))}catch(e){}}else{SPT[e.dataset.b]=v;try{localStorage.setItem('sparePartsTargets',JSON.stringify(SPT))}catch(e){}}drawPartsTargets();drawPartsKpis();anim()})}
$('#qpt').oninput=drawPartsTargets;
$('#spq').oninput=$('#spb1').onchange=$('#spc1').onchange=drawPartsReport;
// facility
function fo(){}
function rf(){const s='',p='';
$('#fl').innerHTML=F.filter(f=>(!s||f.s==s)&&(!p||f.p==p)).map(f=>`<div class="card ${f.p}"><div style="display:flex;justify-content:space-between;gap:6px"><b>${f.b}</b><span class="tag ${f.p}">${f.p}</span></div><div class="sub">#${f.id} · ${f.k} · ${f.dt}</div><p>${f.d}</p><select data-i="${f.id}" class="tbtn">${['Open','Ongoing','Resolved'].map(x=>`<option${x==f.s?' selected':''}>${x}</option>`).join('')}</select></div>`).join('')||'<p class="sub">No concerns match.</p>';
document.querySelectorAll('#fl select').forEach(e=>e.onchange=()=>{F.find(f=>f.id==e.dataset.i).s=e.value;fk();rf();anim()})}


// branch manager

const esc=s=>String(s).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');
const WD=[5.2,5.2,5.2,5.2,5.2];
function ft(){const q=$('#q1').value.toLowerCase(),c=$('#c1').value,a=$('#a1').value,L=fB(),S=f=>L.reduce((s,x)=>s+f(x),0),T=S(x=>TG(x));
let h=`<td>TOTAL (${L.length})</td><td></td><td class="num">${S(x=>x.m)}</td><td></td><td class="num">${peso(T)}</td><td class="num">${peso(S(x=>x.d))}</td>`;
for(let k=0;k<5;k++){const w=S(x=>x.wi[k]),sr=S(x=>(x.sw&&x.sw[k])||0);h+=`<td class="num">${w?peso(w):'—'}</td><td class="num">${w?peso(w/(WD[k]*pm())):'—'}</td><td class="num src">${sr||'—'}</td><td>${pc(T?w/T*100:0,w)}</td>`}
const s=S(x=>x.wi.reduce((p,n)=>p+n,0)),tsr=S(x=>(x.sw||[]).reduce((p,n)=>p+(+n||0),0));
h+=`<td class="num">${s?peso(s):'—'}</td><td class="num">${s?peso(s/(26*pm())):'—'}</td><td class="num">${tsr||'—'}</td><td>${pc(T?s/T*100:0,s)}</td><td class="num">${T>s?peso(T-s):'✔ Achieved'}</td>`;
$('#tf').innerHTML='<tr>'+h+'</tr>'}
const MN=['January','February','March','April','May','June','July','August','September','October','November','December'],PER={y:new Date().getFullYear(),m:new Date().getMonth()+1},pk=m=>PER.y+'-'+String(m).padStart(2,'0'),pm=()=>PER.m?1:12,TG=x=>PER.m?x.t:x.ty;
function sx(x){if(!x.sr)x.sr={};if(!x.wp)x.wp={};if(!x.mo)x.mo={};if(x.dm==null){x.dm=x.m;x.dt=x.t;x.dc=x.c}
const o=PER.m&&x.mo[pk(PER.m)]||{m:x.dm,t:x.dt,c:x.dc};x.m=o.m;x.t=o.t;x.c=o.c;x.d=x.t/D;
x.ty=0;for(let m=1;m<=12;m++){const q=x.mo[pk(m)];x.ty+=q?q.t:x.dt}
if(PER.m){const k=pk(PER.m);x.wi=x.wp[k]||(x.wp[k]=[0,0,0,0,0])}else x.wi=[0,1,2,3,4].map(i=>{let s=0;for(let m=1;m<=12;m++){const a=x.wp[pk(m)];if(a)s+=a[i]}return s});if(PER.m){const k=pk(PER.m);x.sw=x.sr[k]||(x.sr[k]=[0,0,0,0,0])}else x.sw=[0,1,2,3,4].map(i=>{let s=0;for(let m=1;m<=12;m++){const a=x.sr[pk(m)];if(a)s+=+a[i]||0}return s})}
function sp(){B.forEach(sx);const L=PER.m?MN[PER.m-1]+' '+PER.y+' · monthly view':'Year '+PER.y+' · all 12 months combined (targets added up)';$('#pl').textContent=L;$('#spl').textContent=L}
function fillPer(){const y0=new Date().getFullYear();$('#ys').innerHTML=[y0-2,y0-1,y0,y0+1,y0+2].map(y=>`<option${y==PER.y?' selected':''}>${y}</option>`).join('');$('#ms').innerHTML='<option value="0">Whole year</option>'+MN.map((n,i)=>`<option value="${i+1}"${i+1==PER.m?' selected':''}>${n}</option>`).join('');$('#ys').onchange=$('#ms').onchange=()=>{PER.y=+$('#ys').value;PER.m=+$('#ms').value;refreshAll(1)}}
function perf(x,k){return x.t?x.wi[k]/TG(x)*100:0}
function tc(x){const s=x.wi.reduce((a,b)=>a+b,0),t=TG(x),r=(x.sw||[]).reduce((a,b)=>a+(+b||0),0);return[s?peso(s):'—',s?peso(s/(26*pm())):'—',pc(t?s/t*100:0,s),t>s?peso(t-s):'✔ Achieved',r||'—']}
function dv(x,k){return x.wi[k]?peso(x.wi[k]/(WD[k]*pm())):'—'}
function pc(p,w){return w?`<span class="tag ${p>=100?'ok':p>=80?'wn':'er'}">${p.toFixed(1)}%</span>`:'<span class="sub">—</span>'}
function wk(x){const i=B.indexOf(x),ic=(v,k)=>`<td><input class="wv" type="text" inputmode="decimal" data-i="${i}" data-k="${k}" value="${v?peso(v):''}" ${PER.m?'':'readonly'} placeholder="₱0"></td>`;return x.wi.map((w,k)=>ic(w,k)+`<td class="num dy" data-i="${i}" data-k="${k}">${dv(x,k)}</td><td class="src"><input class="sv" type="text" inputmode="numeric" data-i="${i}" data-k="${k}" value="${(x.sw&&x.sw[k])||''}" ${PER.m?'':'readonly'} placeholder="0"></td><td class="pf" data-i="${i}" data-k="${k}">${pc(perf(x,k),w)}</td>`).join('')+[0,1,4,2,3].map(j=>`<td class="num tt" data-i="${i}" data-j="${j}">${tc(x)[j]}</td>`).join('')}
function rc(x){if(!x.w)x.w=[0,0,0,0,0];if(!x.wp)x.wp={};if(!x.c)x.c=x.m>=8?'A':x.m>=5?'B':'C';x.d=x.t/D;x.p=x.t?x.a/x.t*100:0}
function load(a){B.length=0;a.forEach(([b,m,t,v,ar,c,w,wp,mo,dm,dt,dc,spt,sr])=>{const x={b,m,t,a:v,ar,c,w,mo,dm,dt,dc,spt:+spt||0,sr:sr&&typeof sr=='object'&&!Array.isArray(sr)?sr:{},wp:Array.isArray(wp)?{[pk(PER.m||new Date().getMonth()+1)]:wp}:wp};rc(x);B.push(x)})}
function sv(){try{localStorage.setItem('bd2',JSON.stringify(B.map(x=>[x.b,x.m,x.t,x.a,x.ar,x.c,x.w,x.wp,x.mo,x.dm,x.dt,x.dc,x.spt,x.sr])))}catch(e){}}
const DEF=JSON.stringify([]);
try{const s=JSON.parse(localStorage.getItem('bd2'));if(Array.isArray(s)&&s.length)load(s)}catch(e){}
try{const p=JSON.parse(localStorage.getItem('sparePartsTargets'));if(p&&typeof p=='object'&&!Array.isArray(p))SPT=p}catch(e){}
try{const p=JSON.parse(localStorage.getItem('sparePartsClassifications'));if(p&&typeof p=='object'&&!Array.isArray(p))SPC=p}catch(e){}
B.forEach(x=>{if(SPB.includes(x.b)&&SPT[x.b]==null&&x.spt!=null)SPT[x.b]=x.spt});
function wb(){B.forEach(x=>{if(!x.mo)x.mo={};if(PER.m)x.mo[pk(PER.m)]={m:x.m,t:x.t,c:x.c};else{x.dm=x.m;x.dt=x.t;x.dc=x.c}})}
const ASP=['3S SAN PABLO','3S STA ROSA','3S CALAMBA','3S LUCENA','MS GUMACA','MS INFANTA'],ASPL=ASP.map(b=>({b})),asList=k=>k=='parts'?ASPL:B;
const ASS=[['parts','3S Parts Sales'],['svc','Service Revenue'],['oil','Multi-Brand Oil Sales']];
let AS={parts:{},svc:{},oil:{}};try{const p=JSON.parse(localStorage.getItem('afterSalesReport'));if(p&&typeof p=='object'&&!Array.isArray(p))AS=Object.assign(AS,p)}catch(e){}
const asSave=()=>{try{localStorage.setItem('afterSalesReport',JSON.stringify(AS))}catch(e){}};
const asGet=(k,b,y=PER.y,mo=PER.m)=>{const g=m=>{const ky=y+'-'+String(m).padStart(2,'0'),o=AS[k]&&AS[k][ky]&&AS[k][ky][b];let t=+(o&&o.t)||0;if(k=='parts'){t=spM(b,y,m)}else if(k=='svc'&&!(t>0)){const x=B.find(z=>z.b==b),q=x&&x.mo&&x.mo[ky];t=x?(q&&q.t!=null?+q.t:+(x.dt??x.t??0))||0:0}const sf=+(o&&(o.sf!=null?o.sf:o.a))||0,nc=+(o&&o.nc)||0,fcs=k=='svc'?nc*30:+(o&&o.fcs)||0;return{t,a:k=='svc'?sf+fcs:+(o&&o.a)||0,sf,fcs,nc,pm:+(o&&o.pm)||0,py:+(o&&o.py)||0}};if(mo)return g(mo);let t=0,a=0,sf=0,fcs=0,nc=0,pm=0,py=0;for(let m=1;m<=12;m++){const o=g(m);t+=o.t;a+=o.a;sf+=o.sf;fcs+=o.fcs;nc+=o.nc;pm+=o.pm;py+=o.py}return{t,a,sf,fcs,nc,pm,py}};
const asUG=(y=PER.y,mo=PER.m)=>{const g=m=>{const ky=y+'-'+String(m).padStart(2,'0'),o=AS.units&&AS.units[ky],t=B.reduce((s,x)=>s+((x.sr&&x.sr[ky])||[]).reduce((q,n)=>q+(+n||0),0),0);return{t,a:t}};if(mo)return g(mo);let t=0,a=0;for(let m=1;m<=12;m++){const o=g(m);t+=o.t;a+=o.a}return{t,a}};
const nf=n=>Math.round(n).toLocaleString('en-PH');
function asUCalc(){const o=asUG(),pv=asUG(PER.y-1,PER.m).a,v=o.t?(o.a/o.t-1)*100:0,yo=pv?(pv-o.a)/pv*100:null;
 $('#asUv').innerHTML=o.t?`<span class="tag ${v>=0?'ok':'er'}">${v>=0?'+':'−'}${Math.abs(v).toFixed(1)}%</span>`:'—';
 $('#asUp').innerHTML=o.t?pc(o.a/o.t*100,1):'—';
 $('#asUy').innerHTML=yo==null?'—':`<span class="tag ${yo>=0?'ok':'er'}">${yo>=0?'▲ +':'▼ −'}${Math.abs(yo).toFixed(1)}%</span>`}
function asUnits(){const o=asUG(),pY=PER.m>1?PER.y:PER.y-1,pM=PER.m?(PER.m>1?PER.m-1:12):0,v=n=>n?nf(n):'—';
 $('#asUn').innerHTML=`<tr><td>Total Unit Served</td><td class="num">${v(o.t)}</td><td class="num">${v(asUG(PER.y-1,PER.m).a)}</td><td class="num">${v(o.a)}</td><td class="num" id="asUv"></td><td id="asUp"></td><td class="num" id="asUy"></td></tr>`;
 asUCalc()}
const asTot=(k,list,y=PER.y,mo=PER.m)=>list.reduce((r,x)=>{const o=asGet(k,x.b,y,mo);r.t+=o.t;r.a+=o.a;return r},{t:0,a:0});
function fillAsPer(){const y0=new Date().getFullYear();$('#asy').innerHTML=[y0-2,y0-1,y0,y0+1,y0+2].map(y=>`<option${y==PER.y?' selected':''}>${y}</option>`).join('');$('#asm').innerHTML='<option value="0">Whole year</option>'+MN.map((n,i)=>`<option value="${i+1}"${i+1==PER.m?' selected':''}>${n}</option>`).join('');$('#asy').onchange=$('#asm').onchange=()=>{PER.y=+$('#asy').value;PER.m=+$('#asm').value;refreshAll(1)}}
function asCalc(){
 const cells=(t,a)=>[a?pc(t?a/t*100:0,a):'<span class="sub">—</span>',t?(t>a?peso(t-a):'✔ Achieved'):'—'];
 document.querySelectorAll('#asr .asn').forEach(e=>{const x=asList(e.dataset.k)[+e.dataset.i],o=asGet(e.dataset.k,x.b),c=cells(o.t,o.a);e.querySelector('.ap').innerHTML=c[0];e.querySelector('.an').textContent=c[1]});
 let T=0,A=0,h='';const L=B;
 ASS.forEach(([k,n])=>{const o=asTot(k,asList(k)),vis=asTot(k,asVis(k));T+=o.t;A+=o.a;const c=cells(vis.t,vis.a);if(k=='parts'){$('#asf-parts').innerHTML=asPartsFoot(asVis('parts'));return}if(k=='oil'){$('#asf-oil').innerHTML=asOilFoot(asVis('oil'));asOilLive();return}if(k=='svc'){$('#asf-svc').innerHTML=asSvcFoot(asVis('svc'));asSvcLive();return}
  $('#asf-'+k).innerHTML=`<tr><td>TOTAL</td><td class="num">${vis.t?peso(vis.t):'—'}</td><td class="num">${vis.a?peso(vis.a):'—'}</td><td>${c[0]}</td><td class="num">${c[1]}</td></tr>`});
 const pct=(a,t)=>t?(a/t*100):0,f1=v=>v.toFixed(1)+'%';
  $('#asPh').textContent='Previous Year';$('#asMh').textContent='Current Year';
 const pY=PER.m>1?PER.y:PER.y-1,pM=PER.m?(PER.m>1?PER.m-1:12):0;
 const row=(n,t,pp,a,pv)=>{const v=t?(a/t-1)*100:0,yo=pp?(pp-a)/pp*100:null,b=x=>x;
  return`<tr><td>${b(esc(n))}</td><td class="num">${t?peso(t):'—'}</td><td class="num">${pp?peso(pp):'—'}</td><td class="num">${a?peso(a):'—'}</td><td class="num">${t?`<span class="tag ${v>=0?'ok':'er'}">${v>=0?'+':'−'}${Math.abs(v).toFixed(1)}%</span>`:'—'}</td><td>${t?pc(a/t*100,1):'—'}</td><td class="num">${yo==null?'—':`<span class="tag ${yo>=0?'ok':'er'}">${yo>=0?'▲ +':'▼ −'}${Math.abs(yo).toFixed(1)}%</span>`}</td></tr>`};
 let rt_=0,ra=0,rp=0;
 $('#asTb').innerHTML=ASS.map(([k,n])=>{const o=asTot(k,asList(k)),pv=asTot(k,asList(k),PER.y-1).a,pp=asTot(k,asList(k),PER.y-1,PER.m)[k=='parts'?'t':'a'];rt_+=o.t;ra+=o.a;rp+=pv;return row({parts:'Parts Sales Revenue (3S Shops)',svc:'Service Revenue',oil:'MB Oil Sales'}[k]||n,o.t,pp,k=='parts'?o.t:o.a,pv)}).join('');
  anim()}
const asVis=k=>{const q=$('#asq').value.toLowerCase();return asList(k).filter(x=>x.b.toLowerCase().includes(q))};
const asPV=x=>{const pm=PER.m?(PER.m>1?PER.m-1:12):0,pY=PER.m>1?PER.y:PER.y-1,t=asGet('parts',x.b).t;return{t,py:asGet('parts',x.b,PER.y-1,PER.m).t,cy:t,pm:pm?asGet('parts',x.b,pY,pm).t:0}};
const asPf=(t,c)=>t?pc(c/t*100,1):'<span class="sub">—</span>',asF=v=>v?peso(v):'—',asGr=(p,c)=>{if(!p)return'—';const g=(p-c)/p*100;return`<span class="tag ${g>=0?'ok':'er'}">${g>=0?'▲ +':'▼ −'}${Math.abs(g).toFixed(1)}%</span>`};
const asPartsRows=list=>list.map(x=>{const o=asPV(x);return`<tr><td><b>${esc(x.b)}</b></td><td class="num">${asF(o.t)}</td><td class="num">${asF(o.py)}</td><td class="num">${asF(o.cy)}</td><td class="num">${asF(o.pm)}</td><td>${asGr(o.pm,o.cy)}</td><td>${asGr(o.py,o.cy)}</td><td>${asPf(o.t,o.cy)}</td></tr>`}).join('')||'<tr><td colspan="8" class="sub">No branches match the search.</td></tr>';
const asPartsFoot=list=>{const s=list.map(asPV).reduce((r,o)=>({t:r.t+o.t,py:r.py+o.py,cy:r.cy+o.cy,pm:r.pm+o.pm}),{t:0,py:0,cy:0,pm:0});return`<tr><td>TOTAL</td><td class="num">${asF(s.t)}</td><td class="num">${asF(s.py)}</td><td class="num">${asF(s.cy)}</td><td class="num">${asF(s.pm)}</td><td>${asGr(s.pm,s.cy)}</td><td>${asGr(s.py,s.cy)}</td><td>${asPf(s.t,s.cy)}</td></tr>`};
const asSvcRows=list=>{const ro=PER.m?'':'readonly title="Select a month to enter data"';return list.map(x=>{const i=asList('svc').indexOf(x),o=asGet('svc',x.b),inp=(f,v,pf)=>`<td><input class="wv as" data-f="${f}" data-i="${i}" data-k="svc" type="text" inputmode="decimal" value="${v?pf(v):''}" ${ro} placeholder="${f=='nc'?'0':'₱0'}"></td>`;return`<tr data-i="${i}"><td><b>${esc(x.b)}</b></td>${inp('t',o.t,peso)}${inp('sf',o.sf,peso)}<td class="num fc">${asF(o.fcs)}</td>${inp('nc',o.nc,nf)}<td class="num st">${asF(o.a)}</td><td class="sp">${asPf(o.t,o.a)}</td></tr>`}).join('')||'<tr><td colspan="7" class="sub">No branches match the search.</td></tr>'};
const asSvcFoot=list=>{const s=list.map(x=>asGet('svc',x.b)).reduce((r,o)=>({t:r.t+o.t,a:r.a+o.a,sf:r.sf+o.sf,fcs:r.fcs+o.fcs,nc:r.nc+o.nc}),{t:0,a:0,sf:0,fcs:0,nc:0});return`<tr><td>TOTAL</td><td class="num">${asF(s.t)}</td><td class="num">${asF(s.sf)}</td><td class="num">${asF(s.fcs)}</td><td class="num">${s.nc?nf(s.nc):'—'}</td><td class="num">${asF(s.a)}</td><td>${asPf(s.t,s.a)}</td></tr>`};
const asSvcLive=()=>document.querySelectorAll('#asb-svc tr[data-i]').forEach(tr=>{const o=asGet('svc',asList('svc')[+tr.dataset.i].b);tr.querySelector('.fc').textContent=asF(o.fcs);tr.querySelector('.st').textContent=asF(o.a);tr.querySelector('.sp').innerHTML=asPf(o.t,o.a)});
const asOilV=b=>{const o=asGet('oil',b);return{c:o.a,pm:o.pm,py:o.py}};
const asGo=(p,c)=>{if(!c)return'—';const g=(c-p)/c*100;return`<span class="tag ${g>=0?'ok':'er'}">${g>=0?'▲ +':'▼ −'}${Math.abs(g).toFixed(1)}%</span>`};
const asOilRows=list=>{const ro=PER.m?'':'readonly title="Select a month to enter data"';return list.map(x=>{const i=asList('oil').indexOf(x),o=asOilV(x.b);return`<tr data-i="${i}"><td><b>${esc(x.b)}</b></td><td><input class="wv as" data-f="a" data-i="${i}" data-k="oil" type="text" inputmode="decimal" value="${o.c?peso(o.c):''}" ${ro} placeholder="₱0"></td><td><input class="wv as" data-f="pm" data-i="${i}" data-k="oil" type="text" inputmode="decimal" value="${o.pm?peso(o.pm):''}" ${ro} placeholder="₱0"></td><td><input class="wv as" data-f="py" data-i="${i}" data-k="oil" type="text" inputmode="decimal" value="${o.py?peso(o.py):''}" ${ro} placeholder="₱0"></td><td class="mm">${asGo(o.pm,o.c)}</td><td class="yy">${asGo(o.py,o.c)}</td></tr>`}).join('')||'<tr><td colspan="6" class="sub">No branches match the search.</td></tr>'};
const asOilFoot=list=>{const s=list.map(x=>asOilV(x.b)).reduce((r,o)=>({c:r.c+o.c,pm:r.pm+o.pm,py:r.py+o.py}),{c:0,pm:0,py:0});return`<tr><td>TOTAL</td><td class="num">${asF(s.c)}</td><td class="num">${asF(s.pm)}</td><td class="num">${asF(s.py)}</td><td>${asGo(s.pm,s.c)}</td><td>${asGo(s.py,s.c)}</td></tr>`};
const asOilLive=()=>document.querySelectorAll('#asb-oil tr[data-i]').forEach(tr=>{const o=asOilV(asList('oil')[+tr.dataset.i].b);tr.querySelector('.mm').innerHTML=asGo(o.pm,o.c);tr.querySelector('.yy').innerHTML=asGo(o.py,o.c)});
function drawAsr(){
 $('#asy').value=PER.y;$('#asm').value=PER.m;$('#asl').textContent=PER.m?MN[PER.m-1]+' '+PER.y+' · monthly view':'Year '+PER.y+' · all 12 months combined · select a month to enter Multi-Brand Oil data';
 const ro=PER.m?'':'readonly title="Select a month to enter data"';
 ASS.forEach(([k])=>{const vis=asVis(k);if(k=='parts'){$('#asb-parts').innerHTML=asPartsRows(vis);return}if(k=='oil'){$('#asb-oil').innerHTML=asOilRows(vis);return}if(k=='svc'){$('#asb-svc').innerHTML=asSvcRows(vis);return}$('#asb-'+k).innerHTML=vis.map(x=>{const i=asList(k).indexOf(x),o=asGet(k,x.b);return`<tr class="asn" data-i="${i}" data-k="${k}"><td><b>${esc(x.b)}</b></td><td><input class="wv as" data-f="t" data-i="${i}" data-k="${k}" type="text" inputmode="decimal" value="${o.t?peso(o.t):''}" ${k=='parts'?'readonly title="Taken from the Monthly Spare Parts Target on the Spare Parts page"':ro} placeholder="₱0"></td><td><input class="wv as" data-f="a" data-i="${i}" data-k="${k}" type="text" inputmode="decimal" value="${o.a?peso(o.a):''}" ${ro} placeholder="₱0"></td><td class="ap"></td><td class="num an"></td></tr>`}).join('')||'<tr><td colspan="5" class="sub">No branches match the search.</td></tr>'});
 document.querySelectorAll('#asr .as').forEach(e=>{
  e.onbeforeinput=ev=>{if(ev.data&&/[^0-9,.₱]/.test(ev.data)){ev.preventDefault();alert('Input Numbers Only')}};
  e.oninput=()=>{if(/[^0-9,.₱]/.test(e.value)){e.value=e.value.replace(/[^0-9,.₱]/g,'');alert('Input Numbers Only')}
   if(!PER.m)return;const k=e.dataset.k,x=asList(k)[+e.dataset.i],key=pk(PER.m),o=AS[k][key]||(AS[k][key]={}),r=o[x.b]||(o[x.b]={t:0,a:0});r[e.dataset.f]=Math.max(0,num(e.value));asSave();asCalc()};
  e.onblur=()=>{const o=asGet(e.dataset.k,asList(e.dataset.k)[+e.dataset.i].b)[e.dataset.f];e.value=o?(e.dataset.f=='nc'?nf(o):peso(o)):''}});
 asCalc();asUnits()}
$('#asq').oninput=drawAsr;
function refreshAll(per){if(!per)wb();sp();fillBr();sv();drawK1();drawCh();rt();drawAsr();drawPartsKpis();drawPartsTargets();fo();mg();anim()}
function mg(){
 const q=$('#qm').value.toLowerCase(),ag=$('#am').value;
 $('#mb').innerHTML=B.map((x,i)=>({x,i})).filter(o=>o.x.b.toLowerCase().includes(q)&&(!ag||o.x.ar==ag)).map(({x,i})=>{
  const m=x.dm??x.m,t=x.dt??x.t,c=x.dc??x.c;
  return `<tr><td class="num">${i+1}</td><td><input data-i="${i}" data-f="b" value="${esc(x.b)}"></td><td><div class="area-edit-wrap"><select class="tbtn" data-i="${i}" data-f="ar" disabled>${AO.map(a=>`<option${a==x.ar?' selected':''}>${a}</option>`).join('')}</select><button type="button" class="tbtn edit-area" data-i="${i}" aria-label="Edit area" title="Edit area">✎</button></div></td><td><input type="number" min="0" data-i="${i}" data-f="m" value="${m}" style="width:90px"></td><td><select class="tbtn ${c}" data-i="${i}" data-f="c">${['A','B','C'].map(k=>`<option value="${k}"${k==c?' selected':''}>${GR[k]}</option>`).join('')}</select></td><td><input type="number" min="0" step="1000" data-i="${i}" data-f="t" value="${t}" style="width:130px"></td><td class="num">${peso(t/D)}</td><td><button class="tbtn del" data-i="${i}">🗑 Delete</button></td></tr>`;
 }).join('')||'<tr><td colspan="8" class="sub">No branches found.</td></tr>';
 document.querySelectorAll('#mb input,#mb select').forEach(e=>e.onchange=()=>{const x=B[e.dataset.i],f=e.dataset.f;if(f=='b'){const v=e.value.trim().toUpperCase();if(!v){e.value=x.b;return}x.b=v}else if(f=='ar')x.ar=e.value;else if(f=='c')x.dc=e.value;else if(f=='m')x.dm=Math.max(0,+e.value||0);else if(f=='t')x.dt=Math.max(0,+e.value||0);rc(x);refreshAll()});
 document.querySelectorAll('#mb .edit-area').forEach(e=>e.onclick=()=>{const a=e.parentElement.querySelector('select[data-f="ar"]');a.disabled=false;a.focus();e.textContent='✓';e.title='Save area';e.setAttribute('aria-label','Save area');e.onclick=()=>a.dispatchEvent(new Event('change',{bubbles:true}))});
 document.querySelectorAll('#mb .del').forEach(e=>e.onclick=()=>{const x=B[e.dataset.i];if(confirm('Delete '+x.b+'?')){B.splice(e.dataset.i,1);refreshAll()}})
}
$('#qm').oninput=$('#am').onchange=mg;
$('#ab').onclick=()=>{const b=$('#nb').value.trim().toUpperCase();if(!b)return $('#nb').focus();if(B.some(x=>x.b==b)){alert('That branch already exists.');return}const x={b,m:+$('#nm').value||0,t:+$('#nt').value||0,a:0,ar:$('#nar').value,c:$('#ncl').value,spt:0};rc(x);B.unshift(x);['#nb','#nm','#nt'].forEach(s=>$(s).value='');refreshAll()};
$('#rs').onclick=()=>{if(confirm('Clear all branches?')){load(JSON.parse(DEF));refreshAll()}};
const AO=Object.keys(AREAS),GR={A:'1ST GRADE',B:'2ND GRADE',C:'3RD GRADE'};
function fillAreas(){const o=AO.map(a=>`<option>${a}</option>`).join('');$('#a1').innerHTML='<option value="">All areas</option>'+o;$('#am').innerHTML='<option value="">All areas</option>'+o;$('#nar').innerHTML=o}
fillPer();fillSpPer();sp();fillAreas();fillBr();fillAsPer();drawK1();drawCh();rt();drawAsr();drawPartsKpis();drawPartsTargets();fk();rf();fo();mg();anim();
(function(){const g=document.querySelector('.grp'),t=document.getElementById('grt');const tog=()=>{const shut=g.classList.toggle('shut');t.setAttribute('aria-expanded',String(!shut))};t.addEventListener('click',tog);t.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();tog()}})})();
(function(){const b=document.querySelector('.nav button[data-s="parts"]'),g=b.closest('.grp2');b.addEventListener('click',()=>{const shut=g.classList.toggle('shut');b.setAttribute('aria-expanded',String(!shut))})})();
(function(){const t=document.getElementById('grt2'),g=t.closest('.grp');const tog=()=>{const shut=g.classList.toggle('shut');t.setAttribute('aria-expanded',String(!shut))};t.addEventListener('click',tog);t.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();tog()}})})();
(function(){const el=document.getElementById('ch'),N='http://www.w3.org/2000/svg';
function enh(){const s=el.querySelector('svg');if(!s||s.dataset.e)return;s.dataset.e=1;
const cur=[...s.querySelectorAll('path')].find(p=>p.getAttribute('stroke-width')==='3.5');if(!cur)return;
const d=cur.getAttribute('d'),n=d.match(/-?\d+\.?\d*/g).map(Number);if(n.length<4)return;
const ys=[...s.querySelectorAll('line')].map(l=>+l.getAttribute('y1')||0),base=ys.length?Math.max(...ys):cur.getBBox().y+cur.getBBox().height;
let defs=s.querySelector('defs');if(!defs){defs=document.createElementNS(N,'defs');s.prepend(defs)}
defs.insertAdjacentHTML('beforeend','<linearGradient id="gArea" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#e8590c" stop-opacity=".38"/><stop offset="1" stop-color="#e8590c" stop-opacity="0"/></linearGradient>');
const a=document.createElementNS(N,'path');a.setAttribute('d',d+' L'+n[n.length-2]+' '+base+' L'+n[0]+' '+base+' Z');a.setAttribute('fill','url(#gArea)');a.setAttribute('class','ch-area');a.style.pointerEvents='none';cur.before(a);
const L=cur.getTotalLength();cur.style.strokeDasharray=L;cur.style.strokeDashoffset=L;cur.classList.add('ch-draw')}
new MutationObserver(enh).observe(el,{childList:true});enh()})();
(function(){let U=[];try{const p=JSON.parse(localStorage.getItem('uniformMonitoring'));if(Array.isArray(p))U=p}catch(e){}
const sv=()=>{try{localStorage.setItem('uniformMonitoring',JSON.stringify(U))}catch(e){$('#umst').textContent='Storage is full — the last record may not be saved.'}},today=()=>new Date().toISOString().slice(0,10);
const AREAS={'BATAAN AREA':['BATAAN BALANGA','ORANI','LIMAY'],'BATANGAS AREA':['BATANGAS','CALIT-CALIT','LAIYA','LIPA','MALVAR','STO.TOMAS'],'BICOL AREA':['DAET'],'BULACAN AREA':['MALOLOS','MARILAO','PLARIDEL'],'CAVITE AREA':['ALFONSO','BACOOR 1','BACOOR 2','GENERAL TRIAS','NAIC','SILANG','TANZA'],'LAGUNA AREA':['BAY','CALAMBA','3S CALAMBA','LOS BAÑOS','NAGCARLAN','SAN PABLO','3S SAN PABLO','SINILOAN','STA CRUZ','STA ROSA','3S STA ROSA'],'QUEZON AREA':['ATIMONAN','CATANUAN','GUMACA','INFANTA','MS INFANTA','LOPEZ','LUCENA','3S LUCENA','CALAUAG','PAGBILAO','SAN FRANCISCO'],'MINDORO AREA':['CALAPAN','PINAMALAYAN','PUERTO GALERA'],'PAMPANGA AREA':['ANGELES'],'PANGASINAN AREA':['BAYAMBANGG','BINALONAN','CARMEN','DAGUPAN'],'RIZAL AREA':['BINANGONAN','PILILIA','TANAY','SAMPALOC'],'TARLAC AREA':['TARLAC','PANIQUI']};
window.__AREAS=AREAS;const aOpts=Object.keys(AREAS).map(a=>'<option>'+a+'</option>').join('');
function fillB(){const a=$('#uma').value,v=$('#umb').value,l=AREAS[a]||[];$('#umb').innerHTML=l.map(b=>'<option>'+esc(b)+'</option>').join('');if(v&&l.includes(v))$('#umb').value=v}
function draw(){const f=$('#umf'),fv=f.value;if(!$('#uma').options.length){$('#uma').innerHTML=aOpts;fillB()}f.innerHTML='<option value="">All areas</option>'+aOpts;f.value=fv;
const mEl=$('#umfm'),mPrev=mEl.value,ms=[...new Set(U.map(r=>(r.d||'').slice(0,7)).filter(Boolean))].sort().reverse();mEl.innerHTML='<option value="">All months</option>'+ms.map(k=>'<option value="'+k+'">'+MN[+k.slice(5,7)-1]+' '+k.slice(0,4)+'</option>').join('');mEl.value=ms.includes(mPrev)?mPrev:'';const fmv=mEl.value;
const q=$('#umq').value.toLowerCase(),list=U.map((r,i)=>({r,i})).filter(o=>(!fv||o.r.a==fv)&&(!fmv||(o.r.d||'').slice(0,7)==fmv)&&(!q||Object.values(o.r).join(' ').toLowerCase().includes(q)));
$('#umt').innerHTML=list.length?list.map(o=>'<tr><td><b>'+esc(o.r.b)+'</b></td><td>'+esc(o.r.a)+'</td><td>'+esc(o.r.p)+'</td><td>'+esc(o.r.n)+'</td><td>'+esc(o.r.v)+'</td><td class="num">'+esc(o.r.d)+'</td><td>'+esc(o.r.m)+'</td><td>'+(o.r.ph?'<img class="umimg" src="'+o.r.ph+'" alt="Photo" style="width:48px;height:48px;object-fit:cover;border-radius:8px;cursor:zoom-in">':'—')+'</td><td><button class="tbtn umview" style="white-space:nowrap" data-i="'+o.i+'">👁 View</button></td><td><button class="tbtn umdel" style="white-space:nowrap" data-i="'+o.i+'">🗑 Delete</button></td></tr>').join(''):'<tr><td colspan="10" class="sub">No records yet.</td></tr>'}
$('#umd').value=today();
$('#umadd').onclick=()=>{const n=$('#umn').value.trim();if(!n){$('#umn').focus();return}
U.unshift({b:$('#umb').value,a:$('#uma').value.trim(),p:$('#ump').value.trim(),n,v:$('#umv').value.trim(),d:$('#umd').value||today(),m:$('#umm').value.trim(),ph:curPh});
curPh='';$('#umph').value='';$('#umst').textContent='';sv();['ump','umn','umv','umm'].forEach(id=>$('#'+id).value='');draw()};
$('#umt').addEventListener('click',e=>{const b=e.target.closest('.umdel');if(!b)return;const i=+b.dataset.i,o=document.createElement('div');o.style.cssText='position:fixed;inset:0;z-index:99;background:rgba(0,0,0,.5);display:flex;align-items:center;justify-content:center;padding:20px';o.innerHTML='<div class="card" style="max-width:360px;text-align:center"><h2>Delete this record?</h2><p class="sub">Are you sure you want to delete this data? This cannot be undone.</p><p><button class="btn umyes">Yes, delete</button> <button class="tbtn umno">Cancel</button></p></div>';o.querySelector('.umyes').onclick=()=>{U.splice(i,1);sv();draw();o.remove()};o.querySelector('.umno').onclick=()=>o.remove();o.onclick=ev=>{if(ev.target===o)o.remove()};document.body.appendChild(o)});
/* View: percentage of each violation for one employee (same name + branch) */
$('#umt').addEventListener('click',e=>{const b=e.target.closest('.umview');if(!b)return;const r0=U[+b.dataset.i];if(!r0)return;
const nk=s=>String(s||'').replace(/\s+/g,' ').trim().toLowerCase(),recs=U.filter(r=>nk(r.n)==nk(r0.n)&&nk(r.b)==nk(r0.b)),cnt={};
recs.forEach(r=>{const v=(r.v||'').trim()||'Not specified';cnt[v]=(cnt[v]||0)+1});
const tot=recs.length,rows=Object.entries(cnt).sort((x,y)=>y[1]-x[1]||x[0].localeCompare(y[0])),top=rows[0],dates=recs.map(r=>r.d).filter(Boolean).sort();
const o=document.createElement('div');o.style.cssText='position:fixed;inset:0;z-index:99;background:rgba(0,0,0,.55);display:flex;align-items:center;justify-content:center;padding:20px';
o.innerHTML='<div class="card" style="width:min(520px,100%);max-height:90vh;overflow:auto"><div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start"><div><h2 style="margin:0">Violation Summary</h2><p class="sub" style="margin:4px 0 0"><b>'+esc(r0.n)+'</b> · '+esc(r0.b)+(r0.p?' · '+esc(r0.p):'')+'</p></div><button class="tbtn umcl" aria-label="Close">✕</button></div>'
+'<p class="sub" style="margin:12px 0">'+tot+' violation record'+(tot>1?'s':'')+(dates.length?' · '+esc(dates[0])+(dates[dates.length-1]!=dates[0]?' to '+esc(dates[dates.length-1]):''):'')+'</p>'
+'<div style="padding:10px 12px;border:1px solid var(--bd);border-radius:10px;margin-bottom:14px"><span class="sub">Most frequent violation</span><div style="font-weight:700;font-size:16px;margin-top:2px">'+esc(top[0])+' <span class="tag er">'+(top[1]/tot*100).toFixed(1)+'%</span></div></div>'
+rows.map(([k,n])=>{const p=n/tot*100;return'<div style="margin-bottom:12px"><div style="display:flex;justify-content:space-between;gap:10px;font-weight:600"><span>'+esc(k)+'</span><span class="num">'+p.toFixed(1)+'% <span class="sub" style="font-weight:400">('+n+' of '+tot+')</span></span></div><div class="bar" style="margin-top:6px"><i style="width:'+p+'%"></i></div></div>'}).join('')
+'<p class="sub" style="margin:6px 0 0;font-size:12px">Percentage = number of times for each violation ÷ all violation records of this employee (same name and branch).</p></div>';
o.querySelector('.umcl').onclick=()=>o.remove();o.onclick=ev=>{if(ev.target===o)o.remove()};document.body.appendChild(o)});
$('#umq').oninput=draw;$('#umf').onchange=draw;$('#uma').onchange=fillB;$('#umfm').onchange=draw;
let curPh='';
$('#umph').onchange=e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=()=>{const im=new Image();im.onload=()=>{const k=Math.min(1,900/Math.max(im.width,im.height)),c=document.createElement('canvas');c.width=Math.round(im.width*k);c.height=Math.round(im.height*k);c.getContext('2d').drawImage(im,0,0,c.width,c.height);curPh=c.toDataURL('image/jpeg',.72);$('#umst').textContent='Photo attached ✓'};im.src=r.result};r.readAsDataURL(f)};
$('#umt').addEventListener('click',e=>{const m=e.target.closest('.umimg');if(!m)return;const o=document.createElement('div');o.style.cssText='position:fixed;inset:0;z-index:99;background:rgba(0,0,0,.8);display:flex;align-items:center;justify-content:center;cursor:zoom-out;padding:20px';o.innerHTML='<img src="'+m.src+'" style="max-width:100%;max-height:100%;border-radius:12px">';o.onclick=()=>o.remove();document.body.appendChild(o)});
document.querySelector('.nav button[data-s="ad-unif"]').addEventListener('click',draw);draw()})();
(function(){let M=[];try{const p=JSON.parse(localStorage.getItem('memoPreview'));if(Array.isArray(p))M=p}catch(e){}
const sv=()=>{try{localStorage.setItem('memoPreview',JSON.stringify(M))}catch(e){$('#mpst').textContent='Storage is full — the last memo may not be saved.'}},today=()=>new Date().toISOString().slice(0,10);
function draw(){const q=$('#mpq').value.toLowerCase(),list=M.map((r,i)=>({r,i})).filter(o=>!q||Object.values(o.r).join(' ').toLowerCase().includes(q));
$('#mpt').innerHTML=list.length?list.map(o=>'<tr><td><b>'+esc(o.r.r)+'</b></td><td class="num">'+esc(o.r.d)+'</td><td class="num">'+esc(o.r.i)+'</td><td>'+esc(o.r.m)+'</td><td>'+esc(o.r.b)+'</td><td>'+esc(o.r.s)+'</td><td>'+esc(o.r.k)+'</td><td>'+(o.r.ph?'<img class="mpimg" src="'+o.r.ph+'" alt="Photo" style="width:48px;height:48px;object-fit:cover;border-radius:8px;cursor:zoom-in">':'—')+'</td><td><button class="tbtn mpdel" data-i="'+o.i+'">🗑 Delete</button></td></tr>').join(''):'<tr><td colspan="9" class="sub">No memos yet.</td></tr>'}
$('#mpd').value=today();
$('#mpadd').onclick=()=>{const m=$('#mpm').value.trim();if(!m){$('#mpm').focus();return}
M.unshift({r:$('#mpr').value.trim(),d:$('#mpd').value||today(),i:$('#mpi').value,m,b:$('#mpb').value.trim(),s:$('#mps').value.trim(),k:$('#mpk').value.trim(),ph:mPh});
mPh='';$('#mph').value='';$('#mpst').textContent='';sv();['mpr','mpm','mpb','mps','mpk','mpi'].forEach(id=>$('#'+id).value='');$('#mpd').value=today();draw()};
$('#mpt').addEventListener('click',e=>{const b=e.target.closest('.mpdel');if(!b)return;const i=+b.dataset.i,o=document.createElement('div');o.style.cssText='position:fixed;inset:0;z-index:99;background:rgba(0,0,0,.5);display:flex;align-items:center;justify-content:center;padding:20px';o.innerHTML='<div class="card" style="max-width:360px;text-align:center"><h2>Delete this memo?</h2><p class="sub">Are you sure you want to delete this data? This cannot be undone.</p><p><button class="btn mpyes">Yes, delete</button> <button class="tbtn mpno">Cancel</button></p></div>';o.querySelector('.mpyes').onclick=()=>{M.splice(i,1);sv();draw();o.remove()};o.querySelector('.mpno').onclick=()=>o.remove();o.onclick=ev=>{if(ev.target===o)o.remove()};document.body.appendChild(o)});
let mPh='';
$('#mph').onchange=e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=()=>{const im=new Image();im.onload=()=>{const k=Math.min(1,1400/Math.max(im.width,im.height)),c=document.createElement('canvas');c.width=Math.round(im.width*k);c.height=Math.round(im.height*k);c.getContext('2d').drawImage(im,0,0,c.width,c.height);mPh=c.toDataURL('image/jpeg',.75);$('#mpst').textContent='Photo attached ✓'};im.src=r.result};r.readAsDataURL(f)};
$('#mpt').addEventListener('click',e=>{const m=e.target.closest('.mpimg');if(!m)return;const o=document.createElement('div');o.style.cssText='position:fixed;inset:0;z-index:99;background:rgba(0,0,0,.8);display:flex;align-items:center;justify-content:center;cursor:zoom-out;padding:20px';o.innerHTML='<img src="'+m.src+'" style="max-width:100%;max-height:100%;border-radius:12px">';o.onclick=()=>o.remove();document.body.appendChild(o)});
$('#mpq').oninput=draw;document.querySelector('.nav button[data-s="ad-memo"]').addEventListener('click',draw);draw()})();
(function(){let X=[];try{const p=JSON.parse(localStorage.getItem('travelExpense'));if(Array.isArray(p))X=p}catch(e){}
const sv=()=>{try{localStorage.setItem('travelExpense',JSON.stringify(X))}catch(e){$('#test').textContent='Storage is full — the last record may not be saved.'}},today=()=>new Date().toISOString().slice(0,10),F=v=>v?peso(v):'—',n=id=>+$('#'+id).value||0;
$('#tea').innerHTML='<option value="">Branch</option>'+Object.values(window.__AREAS||{}).flat().map(b=>'<option>'+esc(b)+'</option>').join('');
function draw(){const mEl=$('#tefm'),mPrev=mEl.value,ms=[...new Set(X.map(r=>(r.d||'').slice(0,7)).filter(Boolean))].sort().reverse();mEl.innerHTML='<option value="">All months</option>'+ms.map(k=>'<option value="'+k+'">'+MN[+k.slice(5,7)-1]+' '+k.slice(0,4)+'</option>').join('');mEl.value=ms.includes(mPrev)?mPrev:'';const fm=mEl.value,q=$('#teq').value.toLowerCase(),list=X.map((r,i)=>({r,i})).filter(o=>(!fm||(o.r.d||'').slice(0,7)==fm)&&(!q||Object.values(o.r).join(' ').toLowerCase().includes(q)));
$('#tet2').innerHTML=list.length?list.map(o=>'<tr><td class="num">'+esc(o.r.d)+'</td><td>'+esc(o.r.a)+'</td><td>'+esc(o.r.p)+'</td><td>'+esc(o.r.k)+'</td><td>'+esc(o.r.dp)+'</td><td>'+esc(o.r.pe)+'</td><td class="num">'+F(o.r.g)+'</td><td class="num">'+F(o.r.m)+'</td><td class="num">'+F(o.r.o)+'</td><td class="num"><b>'+F(o.r.t)+'</b></td><td><button class="tbtn tedel" data-i="'+o.i+'">🗑 Delete</button></td></tr>').join(''):'<tr><td colspan="11" class="sub">No travel expenses yet.</td></tr>';
const s=k=>list.reduce((a,o)=>a+(+o.r[k]||0),0);$('#tef').innerHTML=list.length?'<tr><td>TOTAL</td><td></td><td></td><td></td><td></td><td></td><td class="num">'+F(s('g'))+'</td><td class="num">'+F(s('m'))+'</td><td class="num">'+F(s('o'))+'</td><td class="num">'+F(s('t'))+'</td><td></td></tr>':''}
const live=()=>$('#tet').textContent='Total Expense: '+peso(n('teg')+n('tem')+n('teo'));
['teg','tem','teo'].forEach(id=>$('#'+id).oninput=live);
$('#ted').value=today();
$('#teadd').onclick=()=>{const p=$('#tep').value.trim();if(!p){$('#tep').focus();return}
const g=n('teg'),m=n('tem'),o=n('teo');X.unshift({d:$('#ted').value||today(),a:$('#tea').value,p,k:$('#tek').value.trim(),dp:$('#tedp').value.trim(),pe:$('#tepe').value.trim(),g,m,o,t:g+m+o});
$('#test').textContent='';sv();['tep','tek','tedp','tepe','teg','tem','teo'].forEach(id=>$('#'+id).value='');$('#ted').value=today();live();draw()};
$('#tet2').addEventListener('click',e=>{const b=e.target.closest('.tedel');if(!b)return;const i=+b.dataset.i,o=document.createElement('div');o.style.cssText='position:fixed;inset:0;z-index:99;background:rgba(0,0,0,.5);display:flex;align-items:center;justify-content:center;padding:20px';o.innerHTML='<div class="card" style="max-width:360px;text-align:center"><h2>Delete this record?</h2><p class="sub">Are you sure you want to delete this data? This cannot be undone.</p><p><button class="btn teyes">Yes, delete</button> <button class="tbtn teno">Cancel</button></p></div>';o.querySelector('.teyes').onclick=()=>{X.splice(i,1);sv();draw();o.remove()};o.querySelector('.teno').onclick=()=>o.remove();o.onclick=ev=>{if(ev.target===o)o.remove()};document.body.appendChild(o)});
$('#teq').oninput=draw;$('#tefm').onchange=draw;document.querySelector('.nav button[data-s="ad-trav"]').addEventListener('click',draw);draw()})();
(function(){const KEY='contractPreviewByBranch';let CT={};try{const p=JSON.parse(localStorage.getItem(KEY));if(p&&typeof p=='object'&&!Array.isArray(p))CT=p;else{const o=JSON.parse(localStorage.getItem('contractPreview'));if(Array.isArray(o))o.slice().reverse().forEach(r=>{CT[r.b]={l:r.l,e:r.e,s:r.s,x:r.x,n:r.n,y:r.y,m:r.m,w:r.w}})}}catch(e){}
const BR=()=>Object.values(window.__AREAS||{}).flat().sort((a,b)=>a.localeCompare(b,undefined,{numeric:true})),F=v=>v?peso(v):'—',g=(b,k)=>(CT[b]&&CT[b][k])||'';
const sv=()=>{try{localStorage.setItem(KEY,JSON.stringify(CT));$('#ctst').textContent=''}catch(e){$('#ctst').textContent='Storage is full.'}};
let PH={};try{const p=JSON.parse(localStorage.getItem('contractPhotos'));if(p&&typeof p=='object'&&!Array.isArray(p))PH=p}catch(e){}
const svp=()=>{try{localStorage.setItem('contractPhotos',JSON.stringify(PH));$('#ctst').textContent=''}catch(e){$('#ctst').textContent='Storage is full - the photo was not saved.'}};
const vis=()=>{const q=$('#ctq').value.toLowerCase();return BR().filter(b=>!q||(b+' '+g(b,'l')+' '+g(b,'n')).toLowerCase().includes(q))};
function cls(b){const x=g(b,'x');if(!x)return'ct-violet';const t=new Date();t.setHours(0,0,0,0);const e=new Date(x+'T00:00:00'),m6=new Date(t),y1=new Date(t);m6.setMonth(m6.getMonth()+6);y1.setFullYear(y1.getFullYear()+1);return e<=t?'ct-gray':e<=m6?'ct-red':e<=y1?'ct-yellow':'ct-green'}
function foot(){const L=vis(),s=k=>L.reduce((a,b)=>a+(+g(b,k)||0),0);$('#ctf').innerHTML='<tr><td>TOTAL</td><td></td><td></td><td></td><td></td><td></td><td></td><td class="num">'+F(s('m'))+'</td><td class="num">'+F(s('w'))+'</td><td></td></tr>'}
function phCell(b,i){const u=PH[b];return'<td style="white-space:nowrap">'+(u?'<img class="ctimg" data-i="'+i+'" src="'+u+'" alt="Photo" style="width:44px;height:44px;object-fit:cover;border-radius:8px;vertical-align:middle;margin-right:6px;cursor:zoom-in"> ':'')+'<label class="tbtn" style="cursor:pointer" title="'+(u?'Replace photo':'Upload photo')+'">'+(u?'🔄':'📷 Upload')+'<input type="file" accept="image/*" class="ctfile" data-i="'+i+'" hidden></label>'+(u?' <button type="button" class="tbtn ctrm" data-i="'+i+'" title="Remove photo">✕</button>':'')+'</td>'}
function draw(){const all=BR(),inp=(b,f,t,ph,w)=>{const v=g(b,f);return'<td><input class="wv" data-i="'+all.indexOf(b)+'" data-f="'+f+'" type="'+t+'"'+(t=='number'?' min="0" step="any"':'')+' value="'+esc(v||'')+'" placeholder="'+ph+'" style="min-width:'+w+'px"></td>'};
$('#ctr').innerHTML=vis().map(b=>'<tr class="'+cls(b)+'"><td><b>'+esc(b)+'</b></td>'+inp(b,'l','text','Lessor',130)+inp(b,'e','text','e.g. 5% yearly',110)+inp(b,'s','date','',140)+inp(b,'x','date','',140)+inp(b,'n','text','Notary',120)+inp(b,'y','number','Years',70)+inp(b,'m','number','₱0',110)+inp(b,'w','number','₱0',120)+phCell(b,all.indexOf(b))+'</tr>').join('')||'<tr><td colspan="10" class="sub">No branches match the search.</td></tr>';foot()}
$('#ctr').addEventListener('input',e=>{const t=e.target,f=t.dataset.f;if(!f)return;const b=BR()[+t.dataset.i],o=CT[b]||(CT[b]={});o[f]=t.type=='number'?(+t.value||0):t.value;
if(f=='s'||f=='x'){const tr=t.closest('tr'),a=o.s,z=o.x;if(a&&z&&z>a){o.y=Math.round((new Date(z)-new Date(a))/(365.25*864e5)*10)/10;tr.querySelector('[data-f="y"]').value=o.y}}
t.closest('tr').className=cls(b);sv();foot()});
$('#ctr').addEventListener('change',e=>{const t=e.target;if(!t.classList.contains('ctfile'))return;const f=t.files[0];if(!f)return;const b=BR()[+t.dataset.i],r=new FileReader();r.onload=()=>{const im=new Image();im.onload=()=>{const k=Math.min(1,700/Math.max(im.width,im.height)),c=document.createElement('canvas');c.width=Math.round(im.width*k);c.height=Math.round(im.height*k);c.getContext('2d').drawImage(im,0,0,c.width,c.height);PH[b]=c.toDataURL('image/jpeg',.6);svp();draw()};im.src=r.result};r.readAsDataURL(f)});
$('#ctr').addEventListener('click',e=>{const m=e.target.closest('.ctimg'),x=e.target.closest('.ctrm');
if(m){const o=document.createElement('div');o.style.cssText='position:fixed;inset:0;z-index:99;background:rgba(0,0,0,.8);display:flex;flex-direction:column;gap:10px;align-items:center;justify-content:center;cursor:zoom-out;padding:20px';o.innerHTML='<div style="color:#fff;font-weight:700">'+esc(BR()[+m.dataset.i])+'</div><img src="'+m.src+'" style="max-width:100%;max-height:85%;border-radius:12px">';o.onclick=()=>o.remove();document.body.appendChild(o);return}
if(x){const b=BR()[+x.dataset.i],o=document.createElement('div');o.style.cssText='position:fixed;inset:0;z-index:99;background:rgba(0,0,0,.5);display:flex;align-items:center;justify-content:center;padding:20px';o.innerHTML='<div class="card" style="max-width:360px;text-align:center"><h2>Remove this photo?</h2><p class="sub">Are you sure you want to delete the photo of '+esc(b)+'? This cannot be undone.</p><p><button class="btn ctyes">Yes, delete</button> <button class="tbtn ctno">Cancel</button></p></div>';o.querySelector('.ctyes').onclick=()=>{delete PH[b];svp();draw();o.remove()};o.querySelector('.ctno').onclick=()=>o.remove();o.onclick=ev=>{if(ev.target===o)o.remove()};document.body.appendChild(o)}});
$('#ctq').oninput=draw;document.querySelector('.nav button[data-s="ad-cont"]').addEventListener('click',draw);draw()})();
(function(){const KEY='facilityConcernsByMonth',td=new Date(),today=()=>td.toISOString().slice(0,10),pad=n=>String(n).padStart(2,'0');let D={};
try{const p=JSON.parse(localStorage.getItem(KEY));if(p&&typeof p=='object'&&!Array.isArray(p))D=p;else{const o=JSON.parse(localStorage.getItem('facilityConcerns'));if(Array.isArray(o))o.forEach(r=>{const k=(r.d||today()).slice(0,7);(D[k]||(D[k]=[])).push(r)})}}catch(e){}
let ym=td.getFullYear()+'-'+pad(td.getMonth()+1);const L=()=>D[ym]||(D[ym]=[]);
const sv=()=>{try{localStorage.setItem(KEY,JSON.stringify(D));$('#fcst').textContent=''}catch(e){$('#fcst').textContent='Storage is full.'}};
const BR=()=>Object.values(window.__AREAS||{}).flat(),ST=['Pending','Assessed','Done','Not Assessed','Not Done'],S2=['Pending Approval','Approved','Ongoing','Done'];
const O=(a,v)=>a.map(x=>'<option'+(x==v?' selected':'')+'>'+esc(x)+'</option>').join('');
const T=(i,f,v,t,w,d)=>'<td><input class="wv" data-i="'+i+'" data-f="'+f+'" type="'+t+'" value="'+esc(v||'')+'" style="min-width:'+w+'px"'+(d?' disabled':'')+'></td>';
const Sel=(i,f,a,v,d,ph)=>'<td><select class="wv" data-i="'+i+'" data-f="'+f+'"'+(d?' disabled':'')+'>'+(ph?'<option value="">'+ph+'</option>':'')+O(a,v)+'</select></td>';
const Y=()=>+ym.slice(0,4),M=()=>+ym.slice(5,7),go=(y,m)=>{while(m<1){m+=12;y--}while(m>12){m-=12;y++}ym=y+'-'+pad(m);draw()};
$('#fcm').innerHTML=MN.map((n,i)=>'<option value="'+(i+1)+'">'+n+'</option>').join('');
function draw(){const q=$('#fcq').value.toLowerCase(),rows=L(),R=L().map((r,i)=>({r,i})).filter(o=>!q||Object.values(o.r).join(' ').toLowerCase().includes(q));
const y0=td.getFullYear(),ys=new Set();for(let y=y0-4;y<=y0+1;y++)ys.add(y);Object.keys(D).forEach(k=>ys.add(+k.slice(0,4)));ys.add(Y());$('#fcy').innerHTML=[...ys].sort().map(y=>'<option>'+y+'</option>').join('');$('#fcy').value=Y();$('#fcm').value=M();
$('#fcl').textContent=MN[M()-1]+' '+Y()+' · '+rows.length+' concern'+(rows.length==1?'':'s');
$('#fct').innerHTML=R.length?R.map(({r,i})=>{const fu=r.s=='Not Assessed'||r.s=='Not Done';return'<tr>'+T(i,'r',r.r,'text',110)+Sel(i,'b',BR(),r.b,0,'Branch')+T(i,'d',r.d,'date',140)+T(i,'c',r.c,'text',200)+T(i,'da',r.da,'date',140)+T(i,'ab',r.ab,'text',120)+T(i,'at',r.at,'text',200)+T(i,'ds',r.ds,'date',140)+T(i,'dr',r.dr,'date',140)+T(i,'rb',r.rb,'text',120)+Sel(i,'s',ST,r.s)+T(i,'rm',r.rm,'text',180)+T(i,'rd',r.rd,'date',140,!fu)+T(i,'sb',r.sb,'text',120,!fu)+T(i,'dp',r.dp,'date',140,!fu)+Sel(i,'s2',S2,r.s2,!fu)+T(i,'td',r.td,'date',140,!fu)+'<td><button class="tbtn fcdel" data-i="'+i+'">🗑 Delete</button></td></tr>'}).join(''):'<tr><td colspan="18" class="sub">No concerns for this month. Click “Add concern”.</td></tr>'}
$('#fcpy').onclick=()=>go(Y()-1,M());$('#fcpm').onclick=()=>go(Y(),M()-1);$('#fcnm').onclick=()=>go(Y(),M()+1);$('#fcm').onchange=()=>go(Y(),+$('#fcm').value);$('#fcy').onchange=()=>go(+$('#fcy').value,M());
$('#fcadd').onclick=()=>{const cur=ym==td.getFullYear()+'-'+pad(td.getMonth()+1);L().unshift({r:'',b:'',d:cur?today():ym+'-01',s:'Pending',s2:'Pending Approval'});sv();draw()};
$('#fct').addEventListener('input',e=>{const t=e.target,f=t.dataset.f;if(!f)return;L()[+t.dataset.i][f]=t.value;sv();if(f=='s')draw()});
$('#fct').addEventListener('click',e=>{const b=e.target.closest('.fcdel');if(!b)return;const i=+b.dataset.i,o=document.createElement('div');o.style.cssText='position:fixed;inset:0;z-index:99;background:rgba(0,0,0,.5);display:flex;align-items:center;justify-content:center;padding:20px';o.innerHTML='<div class="card" style="max-width:360px;text-align:center"><h2>Delete this concern?</h2><p class="sub">Are you sure you want to delete this data? This cannot be undone.</p><p><button class="btn fcyes">Yes, delete</button> <button class="tbtn fcno">Cancel</button></p></div>';o.querySelector('.fcyes').onclick=()=>{L().splice(i,1);sv();draw();o.remove()};o.querySelector('.fcno').onclick=()=>o.remove();o.onclick=ev=>{if(ev.target===o)o.remove()};document.body.appendChild(o)});
$('#fcq').oninput=draw;document.querySelector('.nav button[data-s="fac"]').addEventListener('click',draw);draw()})();

/* ===== After Sales Report: Import Excel + Download (Excel / PDF) for Service Revenue and Multi-Brand Oil Sales ===== */
(function(){
const KK=s=>String(s==null?'':s).toUpperCase().replace(/[^A-Z0-9]/g,''),HH=s=>String(s==null?'':s).toLowerCase().replace(/[^a-z0-9]/g,'');
const isBr=h=>/^(branch|branchname|branches|branchlocation|store|outlet)$/.test(h);
const rnd1=v=>v===''?'':Math.round(v*10)/10;
const gr=(p,c)=>c?rnd1((c-p)/c*100):'';
const CF={
 svc:{label:'Service Revenue',file:'service-revenue',m:'#asSvcM',inp:'#asSvcF',btn:'#asSvcD',
  head:['Branch','Target','Service Fee','FCS','Number of Coupon','Total','Performance (%)'],money:[1,2,3,5],
  data(){const R=B.map(x=>{const o=asGet('svc',x.b);return[x.b,o.t,o.sf,o.fcs,o.nc,o.a,o.t?rnd1(o.a/o.t*100):'']});
   const s=R.reduce((r,v)=>[r[0]+v[1],r[1]+v[2],r[2]+v[3],r[3]+v[4],r[4]+v[5]],[0,0,0,0,0]);
   return R.concat([['TOTAL',s[0],s[1],s[2],s[3],s[4],s[0]?rnd1(s[4]/s[0]*100):'']])},
  col(h){if(/growth|perf|percent|total|fcs/.test(h))return null;if(/target/.test(h))return't';if(/servicefee|^fee/.test(h))return'sf';if(/coupon/.test(h))return'nc';return null}},
 oil:{label:'Multi-Brand Oil Sales',file:'multi-brand-oil-sales',m:'#asOilM',inp:'#asOilF',btn:'#asOilD',
  head:['Branch','Current Year','Previous Month','Previous Year','Month Over Month Growth (%)','Year Over Year Growth (%)'],money:[1,2,3],pctCols:[4,5],
  data(){const R=B.map(x=>{const o=asOilV(x.b);return[x.b,o.c,o.pm,o.py,gr(o.pm,o.c),gr(o.py,o.c)]});
   const s=R.reduce((r,v)=>[r[0]+v[1],r[1]+v[2],r[2]+v[3]],[0,0,0]);
   return R.concat([['TOTAL',s[0],s[1],s[2],gr(s[1],s[0]),gr(s[2],s[0])]])},
  col(h){if(/growth|mom|yoy|perf|percent/.test(h))return null;if(/current/.test(h))return'a';if(/previousmonth|prevmonth|lastmonth/.test(h))return'pm';if(/previousyear|prevyear|lastyear/.test(h))return'py';return null}}
};
const say=(k,t)=>{$(CF[k].m).textContent=t};
const period=()=>PER.m?MN[PER.m-1]+' '+PER.y:'Year '+PER.y;
const stamp=()=>PER.m?pk(PER.m):String(PER.y);
function saveFile(name,blob){const a=document.createElement('a'),u=URL.createObjectURL(blob);a.href=u;a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{a.remove();URL.revokeObjectURL(u)},1500)}

function doImport(k,f){
 const C=CF[k];
 if(typeof XLSX=='undefined'){say(k,'Excel library failed to load.');return}
 if(!PER.m){say(k,'Pick a month first, then import.');return}
 const r=new FileReader();
 r.onerror=()=>say(k,'Could not read the file.');
 r.onload=ev=>{try{
  const wb=XLSX.read(ev.target.result,{type:'array'}),key=pk(PER.m);
  const byName={};B.forEach(x=>{byName[KK(x.b)]=x});
  let upd=0,sheets=0;const miss=[],seen=new Set();
  wb.SheetNames.forEach(sn=>{
   const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:true,blankrows:false});
   const hi=rows.findIndex((row,i)=>i<40&&row.some(c=>isBr(HH(c))));if(hi<0)return;
   const col={};rows[hi].map(HH).forEach((h,c)=>{if(!h)return;if(col.b==null&&isBr(h)){col.b=c;return}const f2=C.col(h);if(f2&&col[f2]==null)col[f2]=c});
   const fields=['t','sf','nc','a','pm','py'].filter(z=>col[z]!=null);if(!fields.length)return;
   sheets++;
   rows.slice(hi+1).forEach(row=>{
    const nm=String(row[col.b]==null?'':row[col.b]).replace(/\s+/g,' ').trim();if(!nm)return;
    if(/^(TOTAL|GRAND TOTAL|OVERALL|SUBTOTAL)/i.test(nm))return;
    const x=byName[KK(nm)];if(!x){if(miss.indexOf(nm)<0)miss.push(nm);return}
    const vals={};fields.forEach(z=>{const v=row[col[z]];if(v===''||v==null)return;vals[z]=Math.max(0,num(v))});
    if(!Object.keys(vals).length)return;
    AS[k]=AS[k]||{};const o=AS[k][key]||(AS[k][key]={}),rec=o[x.b]||(o[x.b]={t:0,a:0});
    Object.assign(rec,vals);if(!seen.has(x.b)){seen.add(x.b);upd++}
   })});
  if(!sheets){say(k,'No usable columns found. The file needs a "Branch" column and the same column names as the table.');return}
  if(upd)asSave();drawAsr();asCalc();
  say(k,(upd?'Imported '+upd+' branch'+(upd>1?'es':'')+' for '+MN[PER.m-1]+' '+PER.y+'.':'No matching branches were found.')+(miss.length?' Not found in your branch list: '+miss.slice(0,6).join(', ')+(miss.length>6?' and '+(miss.length-6)+' more':'')+'.':''));
 }catch(e){console.error(e);say(k,'Could not read this file. Please use an .xlsx, .xls or .csv file.')}};
 r.readAsArrayBuffer(f);
}

let dlK='svc';
function xlsxOut(k){
 if(typeof XLSX=='undefined'){say(k,'Excel library failed to load.');return}
 const C=CF[k],rows=C.data(),ws=XLSX.utils.aoa_to_sheet([C.head,...rows]);
 ws['!cols']=C.head.map((h,i)=>({wch:i?Math.max(16,h.length+2):26}));
 const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,C.label.slice(0,31));
 const out=XLSX.write(wb,{bookType:'xlsx',type:'array'});
 saveFile(C.file+'-'+stamp()+'.xlsx',new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));say(k,'Excel file downloaded.');
}
function pdfOut(k){
 const J=window.jspdf&&window.jspdf.jsPDF;if(!J){say(k,'PDF library failed to load.');return}
 const C=CF[k];
 try{const doc=new J({orientation:'landscape',unit:'mm',format:'a4'});
  if(typeof doc.autoTable!='function'){say(k,'PDF table library failed to load.');return}
  const rows=C.data(),php=v=>v===''||v==null||v===0?'-':'PHP '+Math.round(v).toLocaleString('en-PH');
  const fmt=(v,j)=>{if(j==0)return String(v);if(k=='svc'){if(j==4)return v?Math.round(v).toLocaleString('en-PH'):'-';if(j==6)return v===''?'-':v+'%';return php(v)}
   if(j>=4)return v===''?'-':(v>=0?'+':'')+v+'%';return php(v)};
  doc.setFontSize(14);doc.text(C.label+' - '+period(),10,12);
  doc.autoTable({head:[C.head],body:rows.map(r=>r.map(fmt)),startY:17,theme:'grid',styles:{fontSize:8,cellPadding:1.6,halign:'center',valign:'middle',overflow:'linebreak'},headStyles:{fillColor:[20,40,75],textColor:255,halign:'center'},columnStyles:{0:{halign:'left',fontStyle:'bold'}},margin:{left:10,right:10,top:12,bottom:10},
   didParseCell:d=>{if(d.section=='body'&&d.row.index==rows.length-1){d.cell.styles.fontStyle='bold';d.cell.styles.fillColor=[230,235,244]}}});
  saveFile(C.file+'-'+stamp()+'.pdf',new Blob([doc.output('arraybuffer')],{type:'application/pdf'}));say(k,'PDF file downloaded.');
 }catch(e){console.error(e);say(k,'Could not create the PDF file.')}
}
['svc','oil'].forEach(k=>{
 const C=CF[k];
 $(C.inp).onchange=e=>{const f=e.target.files[0];if(f)doImport(k,f);e.target.value=''};
 $(C.btn).onclick=()=>{dlK=k;$('#asDlSub').textContent='Choose the file format for the '+C.label+' table ('+period()+').';$('#asDlg').showModal()};
});
$('#asDlC').onclick=()=>$('#asDlg').close();
$('#asDlX').onclick=()=>{$('#asDlg').close();xlsxOut(dlK)};
$('#asDlP').onclick=()=>{$('#asDlg').close();pdfOut(dlK)};
})();
