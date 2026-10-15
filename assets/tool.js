// Shared in-browser tool. Mode from <div id="tool" data-mode data-target-kb data-format>. Nothing is uploaded.
(function(){
const el=document.getElementById('tool'); if(!el) return;
const mode=el.dataset.mode, targetKB=+el.dataset.targetKb||0, fixedFormat=el.dataset.format||'';
const $=s=>el.querySelector(s), out=$('.out'), status=$('.status');
let files=[];
const say=(t,err)=>{status.textContent=t;status.className='status'+(err?' err':'');};
const kb=n=>n<1048576?(n/1024).toFixed(0)+' KB':(n/1048576).toFixed(2)+' MB';
function dl(bytes,name,mime){const b=new Blob([bytes],{type:mime||'application/pdf'});const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=name;a.textContent='Download '+name+' ('+kb(b.size)+')';a.className='btn dl';out.appendChild(a);return b.size;}
const read=f=>f.arrayBuffer();
const input=$('input[type=file]'), drop=$('.drop');
const multiModes=new Set(['merge','images','imageskb','mixed','heic']);
function setFiles(list){files=[...list];$('.list').innerHTML=files.map((f,i)=>`<li>${i+1}. ${f.name.replace(/</g,'&lt;')} <small>${kb(f.size)}</small>${multiModes.has(mode)?` <button data-up="${i}" aria-label="Move up">↑</button>`:''}</li>`).join('');say(files.length?files.length+' file(s) ready':'');}
input.addEventListener('change',()=>setFiles(input.files));
$('.list').addEventListener('click',e=>{const i=+e.target.dataset.up;if(e.target.dataset.up&&i>0){[files[i-1],files[i]]=[files[i],files[i-1]];setFiles(files);}});
['dragover','dragenter'].forEach(t=>drop.addEventListener(t,e=>{e.preventDefault();drop.classList.add('hot');}));
['dragleave','drop'].forEach(t=>drop.addEventListener(t,e=>{e.preventDefault();drop.classList.remove('hot');}));
drop.addEventListener('drop',e=>setFiles(e.dataTransfer.files));

let pdfjs, heicReady;
async function getPdfjs(){if(!pdfjs){pdfjs=await import(window.ASSET+'pdf.min.mjs');pdfjs.GlobalWorkerOptions.workerSrc=window.ASSET+'pdf.worker.min.mjs';}return pdfjs;}
async function getQpdfFactory(){
  if(window.__qpdfFactory) return window.__qpdfFactory;
  await new Promise((res,rej)=>{const s=document.createElement('script');s.src=window.ASSET+'qpdf.js';s.onload=res;s.onerror=()=>rej(new Error('Could not load qpdf'));document.head.appendChild(s);});
  if(typeof Module!=='function') throw new Error('qpdf Module factory missing');
  window.__qpdfFactory=Module;
  return Module;
}
// Fresh qpdf instance each run — callMain exits the program
async function runQpdf(args, inputBytes){
  const factory=await getQpdfFactory();
  const qpdf=await factory({locateFile:()=>window.ASSET+'qpdf.wasm', noInitialRun:true});
  qpdf.FS.writeFile('/input.pdf', new Uint8Array(inputBytes));
  let code=0;
  try{code=qpdf.callMain(args);}catch(e){if(e&&(e.name==='ExitStatus'||e.status!=null)) code=e.status|0; else throw e;}
  if(code!==0) throw new Error(code===2?'Wrong password or the PDF could not be decrypted.':'qpdf failed (exit '+code+'). Check the password and try again.');
  return qpdf.FS.readFile('/output.pdf');
}
async function loadHeic(){
  if(heicReady) return;
  await new Promise((res,rej)=>{const s=document.createElement('script');s.src=window.ASSET+'heic2any.min.js';s.onload=res;s.onerror=()=>rej(new Error('Could not load HEIC converter'));document.head.appendChild(s);});
  heicReady=true;
}

async function rasterise(buf,scale,q,gray){
  const lib=await getPdfjs();const src=await lib.getDocument({data:new Uint8Array(buf.slice(0))}).promise;
  const doc=await PDFLib.PDFDocument.create();
  for(let i=1;i<=src.numPages;i++){const p=await src.getPage(i);const vp1=p.getViewport({scale:1});const vp=p.getViewport({scale});
    const c=document.createElement('canvas');c.width=Math.ceil(vp.width);c.height=Math.ceil(vp.height);const ctx=c.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,c.width,c.height);
    await p.render({canvasContext:ctx,viewport:vp}).promise;
    if(gray){const d=ctx.getImageData(0,0,c.width,c.height),a=d.data;for(let k=0;k<a.length;k+=4){const g=a[k]*.3+a[k+1]*.59+a[k+2]*.11;a[k]=a[k+1]=a[k+2]=g;}ctx.putImageData(d,0,0);}
    const blob=await new Promise(r=>c.toBlob(r,'image/jpeg',q));const img=await doc.embedJpg(new Uint8Array(await blob.arrayBuffer()));
    doc.addPage([vp1.width,vp1.height]).drawImage(img,{x:0,y:0,width:vp1.width,height:vp1.height});}
  return doc.save();
}
async function compress(buf){
  const tKB=+($('#target')||{}).value||targetKB||0, gray=($('#gray')||{}).checked;
  const re=await (await PPTCore.reorder(buf,[...Array(await PPTCore.pageCount(buf)).keys()]));
  let best=re.byteLength<buf.byteLength?re:new Uint8Array(buf);
  if(tKB&&best.byteLength<=tKB*1024) return {bytes:best,note:'Already under '+tKB+' KB with no quality loss.'};
  const steps=[[1.6,.8],[1.3,.7],[1.1,.6],[.9,.5],[.75,.42],[.6,.35],[.5,.3],[.4,.25]];
  for(const [s,q] of steps){say('Compressing… trying quality '+Math.round(q*100)+'%');const r=await rasterise(buf,s,q,gray);if(r.byteLength<best.byteLength)best=r;if(!tKB&&s<=1.1)break;if(tKB&&r.byteLength<=tKB*1024)return{bytes:r,note:'Hit target: '+kb(r.byteLength)+' (limit '+tKB+' KB).'};}
  return {bytes:best,note:tKB?'Smallest we could reach is '+kb(best.byteLength)+'. Try greyscale, or split the file into parts.':'Reduced from '+kb(buf.byteLength)+' to '+kb(best.byteLength)+'.'};
}

async function pdfToImages(buf, format, hq, tKB){
  const lib=await getPdfjs();
  const src=await lib.getDocument({data:new Uint8Array(buf.slice(0))}).promise;
  const mime=format==='png'?'image/png':'image/jpeg';
  const ext=format==='png'?'png':'jpg';
  const results=[];
  const scales=hq?[2.5,2,1.6]:[1.5,1.2,1,0.8,0.6];
  const quals=hq?[0.95,0.9]:[0.85,0.7,0.55,0.4,0.3];
  for(let i=1;i<=src.numPages;i++){
    say('Rendering page '+i+' of '+src.numPages+'…');
    let best=null;
    const tryScales=tKB?scales: [hq?2:1.5];
    const tryQuals=tKB?quals: [hq?0.92:0.85];
    outer: for(const sc of tryScales){
      for(const q of tryQuals){
        const p=await src.getPage(i);const vp=p.getViewport({scale:sc});
        const c=document.createElement('canvas');c.width=Math.ceil(vp.width);c.height=Math.ceil(vp.height);
        const ctx=c.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,c.width,c.height);
        await p.render({canvasContext:ctx,viewport:vp}).promise;
        const blob=await new Promise(r=>c.toBlob(r,mime, format==='png'?undefined:q));
        const bytes=new Uint8Array(await blob.arrayBuffer());
        if(!best||bytes.byteLength<best.byteLength) best=bytes;
        if(!tKB) break outer;
        if(bytes.byteLength<=tKB*1024){best=bytes;break outer;}
      }
    }
    results.push({name:'page-'+i+'.'+ext, bytes:best});
  }
  return results;
}

function sigCanvasSetup(){
  const canvas=$('#sig'); if(!canvas) return null;
  const ctx=canvas.getContext('2d');
  let drawing=false, last=null;
  const pos=e=>{const r=canvas.getBoundingClientRect();const t=e.touches?e.touches[0]:e;return {x:(t.clientX-r.left)*(canvas.width/r.width), y:(t.clientY-r.top)*(canvas.height/r.height)};};
  const start=e=>{e.preventDefault();drawing=true;last=pos(e);};
  const move=e=>{if(!drawing)return;e.preventDefault();const p=pos(e);ctx.strokeStyle='#111';ctx.lineWidth=2.2;ctx.lineCap='round';ctx.beginPath();ctx.moveTo(last.x,last.y);ctx.lineTo(p.x,p.y);ctx.stroke();last=p;};
  const end=()=>{drawing=false;};
  canvas.addEventListener('mousedown',start);canvas.addEventListener('mousemove',move);window.addEventListener('mouseup',end);
  canvas.addEventListener('touchstart',start,{passive:false});canvas.addEventListener('touchmove',move,{passive:false});canvas.addEventListener('touchend',end);
  const clearBtn=$('#sigclear'); if(clearBtn) clearBtn.onclick=()=>{ctx.clearRect(0,0,canvas.width,canvas.height);};
  return {canvas, ctx, isEmpty(){const d=ctx.getImageData(0,0,canvas.width,canvas.height).data; for(let i=3;i<d.length;i+=4) if(d[i]>10) return false; return true;}};
}
const sig=sigCanvasSetup();

// Fill-form: when a PDF is chosen, list fields
async function refreshFields(){
  const box=$('#fields'); if(!box||mode!=='fill') return;
  box.innerHTML='';
  if(!files.length) return;
  try{
    const fields=await PPTCore.listFields(await read(files[0]));
    if(!fields.length){box.innerHTML='<p class="note">This PDF has no fillable form fields (AcroForm). You can still sign it on the Sign PDF page.</p>';return;}
    box.innerHTML=fields.map((f,i)=>{
      if(f.type==='CheckBox') return `<label><input type="checkbox" data-f="${f.name}" ${f.value?'checked':''}> ${f.name.replace(/</g,'&lt;')}</label>`;
      if(f.type==='Dropdown'||f.type==='OptionList'||f.type==='RadioGroup'){
        const opts=(f.options||[]).map(o=>`<option value="${String(o).replace(/"/g,'&quot;')}" ${o===f.value?'selected':''}>${String(o).replace(/</g,'&lt;')}</option>`).join('');
        return `<label>${f.name.replace(/</g,'&lt;')} <select data-f="${f.name}"><option value="">—</option>${opts}</select></label>`;
      }
      if(f.type==='Signature'||f.type==='Button') return '';
      return `<label>${f.name.replace(/</g,'&lt;')} <input type="text" data-f="${f.name}" value="${String(f.value||'').replace(/"/g,'&quot;')}"></label>`;
    }).join('');
  }catch(e){box.innerHTML='<p class="err">'+String(e.message||e).replace(/</g,'&lt;')+'</p>';}
}
if(mode==='fill'){input.addEventListener('change',()=>setTimeout(refreshFields,50));}

$('.go').addEventListener('click',async()=>{
  out.innerHTML='';if(!files.length)return say('Choose a file first.',1);
  try{say('Working… (all on your device)');
    if(mode==='merge'){if(files.length<2)return say('Choose at least two PDFs.',1);dl(await PPTCore.merge(await Promise.all(files.map(read))),'merged.pdf');}
    else if(mode==='images'||mode==='imageskb'){
      let pdfBytes=await PPTCore.imagesToPdf(await Promise.all(files.map(read)),($('#size')||{}).value||'a4');
      const t=+($('#target')||{}).value||targetKB||0;
      if(mode==='imageskb'||t){const r=await compress(pdfBytes);pdfBytes=r.bytes;dl(pdfBytes,'images.pdf');return say(r.note);}
      dl(pdfBytes,'images.pdf');
    }
    else if(mode==='mixed'){
      if(files.length<2)return say('Choose at least two files (PDF and/or JPG/PNG).',1);
      const items=[];
      for(const f of files){
        const bytes=await read(f);
        if(/\.(jpg|jpeg|png)$/i.test(f.name)||/^image\//.test(f.type)) items.push({type:'image',bytes});
        else items.push({type:'pdf',bytes});
      }
      dl(await PPTCore.mixedMerge(items,($('#size')||{}).value||'a4'),'merged.pdf');
    }
    else if(mode==='heic'){
      await loadHeic();
      const jpgs=[];
      for(const f of files){
        say('Converting '+f.name+'…');
        let blob=f;
        if(!/\.heic$/i.test(f.name)&&f.type!=='image/heic'&&f.type!=='image/heif'){
          // allow jpg/png passthrough
          if(/^image\/(jpeg|png)/.test(f.type)||/\.(jpg|jpeg|png)$/i.test(f.name)){jpgs.push(await read(f));continue;}
        }
        const converted=await heic2any({blob:f, toType:'image/jpeg', quality:0.92});
        const one=Array.isArray(converted)?converted[0]:converted;
        jpgs.push(await one.arrayBuffer());
      }
      dl(await PPTCore.imagesToPdf(jpgs,($('#size')||{}).value||'a4'),'heic.pdf');
    }
    else if(mode==='pdf2img'){
      const fmt=fixedFormat||($('#fmt')||{}).value||'jpg';
      const hq=!!($('#hq')||{}).checked || el.dataset.hq==='1';
      const t=+($('#target')||{}).value||targetKB||0;
      const imgs=await pdfToImages(await read(files[0]), fmt, hq, t);
      imgs.forEach(im=>dl(im.bytes, im.name, fmt==='png'?'image/png':'image/jpeg'));
      const note=t?imgs.every(im=>im.bytes.byteLength<=t*1024)?'Each image is under '+t+' KB.':'Some pages stayed above '+t+' KB (image content limit).':'Done. One image per page.';
      return say(note);
    }
    else if(mode==='sign'){
      const buf=await read(files[0]);
      const n=await PPTCore.pageCount(buf);
      const page=(+($('#sigpage')||{}).value||1)-1;
      if(page<0||page>=n) return say('Page must be between 1 and '+n+'.',1);
      const place=($('#sigplace')||{}).value||'br';
      const sizes=await PPTCore.pageSizes(buf);
      const {width:W,height:H}=sizes[page];
      const typed=(($('#sigtext')||{}).value||'').trim();
      let pngBytes=null;
      if(sig && !sig.isEmpty()){
        pngBytes=await new Promise(r=>sig.canvas.toBlob(b=>b.arrayBuffer().then(r),'image/png'));
      } else if(typed){
        const c=document.createElement('canvas');c.width=500;c.height=140;
        const ctx=c.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,c.width,c.height);
        ctx.fillStyle='#111';ctx.font='italic 56px "Segoe Script", "Comic Sans MS", cursive, serif';
        ctx.fillText(typed, 20, 90);
        pngBytes=await new Promise(r=>c.toBlob(b=>b.arrayBuffer().then(r),'image/png'));
      } else return say('Draw a signature or type your name first.',1);
      const w=Math.min(180, W*0.35);
      const margin=36;
      let cx,cy;
      if(place==='br'){cx=W-margin-w/2;cy=margin+40;}
      else if(place==='bl'){cx=margin+w/2;cy=margin+40;}
      else if(place==='tr'){cx=W-margin-w/2;cy=H-margin-20;}
      else if(place==='tl'){cx=margin+w/2;cy=H-margin-20;}
      else {cx=W/2;cy=H/2;}
      dl(await PPTCore.addImageStamp(buf, pngBytes, {page, cx, cy, width:w}), files[0].name.replace(/\.pdf$/i,'')+'-signed.pdf');
    }
    else if(mode==='fill'){
      const vals={};
      el.querySelectorAll('[data-f]').forEach(inp=>{
        if(inp.type==='checkbox') vals[inp.dataset.f]=inp.checked;
        else if(inp.value!=='') vals[inp.dataset.f]=inp.value;
      });
      const flatten=!!($('#flatten')||{}).checked;
      dl(await PPTCore.fillForm(await read(files[0]), vals, flatten), files[0].name.replace(/\.pdf$/i,'')+'-filled.pdf');
    }
    else if(mode==='pagenum'){
      dl(await PPTCore.addPageNumbers(await read(files[0]), {
        start:+($('#start')||{}).value||1,
        format:($('#fmt')||{}).value||'n',
        position:($('#pos')||{}).value||'bc',
        size:+($('#size')||{}).value||10
      }), files[0].name.replace(/\.pdf$/i,'')+'-numbered.pdf');
    }
    else if(mode==='watermark'){
      const text=(($('#wmtext')||{}).value||'').trim();
      if(!text) return say('Enter watermark text.',1);
      dl(await PPTCore.addWatermark(await read(files[0]), {
        text, opacity:+($('#opacity')||{}).value||0.25, size:+($('#wmsize')||{}).value||48, angle:+($('#angle')||{}).value||45
      }), files[0].name.replace(/\.pdf$/i,'')+'-watermarked.pdf');
    }
    else if(mode==='protect'){
      const pass=(($('#pass')||{}).value||'');
      const pass2=(($('#pass2')||{}).value||'');
      if(pass.length<1) return say('Enter a password.',1);
      if(pass!==pass2) return say('Passwords do not match.',1);
      say('Encrypting with AES-256…');
      const outBytes=await runQpdf(['--encrypt', pass, pass, '256', '--', '/input.pdf', '/output.pdf'], await read(files[0]));
      dl(outBytes, files[0].name.replace(/\.pdf$/i,'')+'-protected.pdf');
    }
    else if(mode==='unlock'){
      const pass=(($('#pass')||{}).value||'');
      if(!pass) return say('Enter the current password. This tool cannot crack unknown passwords.',1);
      say('Decrypting…');
      const outBytes=await runQpdf(['--password='+pass, '--decrypt', '/input.pdf', '/output.pdf'], await read(files[0]));
      dl(outBytes, files[0].name.replace(/\.pdf$/i,'')+'-unlocked.pdf');
    }
    else{const buf=await read(files[0]),base=files[0].name.replace(/\.pdf$/i,'');
      if(mode==='compress'){const r=await compress(buf);dl(r.bytes,base+'-compressed.pdf');return say(r.note);}
      if(mode==='extract'){dl(await PPTCore.extract(buf,$('#pages').value),base+'-pages.pdf');}
      if(mode==='split'){const spec=$('#pages').value.trim()||'each';const parts=await PPTCore.split(buf,spec);parts.forEach((p,i)=>dl(p,base+'-part'+(i+1)+'.pdf'));}
      if(mode==='rotate'){const n=await PPTCore.pageCount(buf);const deg=+$('#deg').value;const sel=$('#pages').value.trim()?PPTCore.parseRanges($('#pages').value,n):[...Array(n).keys()];dl(await PPTCore.reorder(buf,[...Array(n).keys()].map(i=>({index:i,rotate:sel.includes(i)?deg:0}))),base+'-rotated.pdf');}
    }
    say('Done. Your file never left this device.');
  }catch(e){say(e.message||String(e),1);console.error(e);}
});
})();
