// In-browser .xlsx reader + Wave Tracker parser. Exposes window.WaveParser.
(function(){
  async function unzip(blob){
    const buf = new Uint8Array(await blob.arrayBuffer());
    const dv = new DataView(buf.buffer);
    let eocd=-1;
    for(let i=buf.length-22;i>=0;i--){ if(dv.getUint32(i,true)===0x06054b50){eocd=i;break;} }
    if(eocd<0) throw new Error('ملف غير صالح');
    const cdCount=dv.getUint16(eocd+10,true);
    let cdOff=dv.getUint32(eocd+16,true);
    const files={}; let p=cdOff;
    for(let n=0;n<cdCount;n++){
      if(dv.getUint32(p,true)!==0x02014b50) break;
      const method=dv.getUint16(p+10,true);
      const compSize=dv.getUint32(p+20,true);
      const nameLen=dv.getUint16(p+28,true);
      const extraLen=dv.getUint16(p+30,true);
      const commLen=dv.getUint16(p+32,true);
      const lho=dv.getUint32(p+42,true);
      const name=new TextDecoder().decode(buf.subarray(p+46,p+46+nameLen));
      const lnameLen=dv.getUint16(lho+26,true);
      const lextraLen=dv.getUint16(lho+28,true);
      const dataStart=lho+30+lnameLen+lextraLen;
      const comp=buf.subarray(dataStart,dataStart+compSize);
      let data;
      if(method===0){ data=comp; }
      else { const ds=new DecompressionStream('deflate-raw');
        const stream=new Blob([comp]).stream().pipeThrough(ds);
        data=new Uint8Array(await new Response(stream).arrayBuffer()); }
      files[name]=new TextDecoder().decode(data);
      p+=46+nameLen+extraLen+commLen;
    }
    return files;
  }
  function parseShared(xml){ if(!xml) return [];
    const arr=[]; const siRe=/<si>([\s\S]*?)<\/si>/g; let m;
    while((m=siRe.exec(xml))){ arr.push(m[1].replace(/<[^>]+>/g,'').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"')); }
    return arr; }
  function parseSheet(xml, shared){
    const grid={}; const rowRe=/<row[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g; let rm;
    while((rm=rowRe.exec(xml))){
      const cellRe=/<c r="([A-Z]+)(\d+)"(?:[^>]*?\st="([^"]+)")?[^>]*>(?:<v>([\s\S]*?)<\/v>|<is>([\s\S]*?)<\/is>)?<\/c>/g; let cm;
      while((cm=cellRe.exec(rm[2]))){
        const col=cm[1], r=cm[2], t=cm[3], v=cm[4], is=cm[5]; let val=null;
        if(t==='s' && v!=null){ val=shared[+v]; }
        else if(t==='inlineStr' && is!=null){ val=is.replace(/<[^>]+>/g,''); }
        else if(v!=null){ val=v; }
        if(val!=null){ val=String(val).replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&#10;/g,'\n').trim(); }
        grid[col+r]={col,r:+r,val};
      }
    }
    return grid;
  }
  function wbSheets(files){
    const wb=files['xl/workbook.xml']; const rels=files['xl/_rels/workbook.xml.rels'];
    const relMap={}; const relRe=/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g; let r;
    while((r=relRe.exec(rels))){ relMap[r[1]]=r[2]; }
    const sheets=[]; const shRe=/<sheet name="([^"]+)"[^>]*r:id="([^"]+)"/g; let s;
    while((s=shRe.exec(wb))){ sheets.push({name:s[1], path:'xl/'+(relMap[s[2]]||'').replace(/^\//,'')}); }
    return sheets;
  }
  function xserial(v){ const n=+v; if(!n||isNaN(n)) return null;
    const d=new Date(Date.UTC(1899,11,30)+n*86400000);
    const dd=String(d.getUTCDate()).padStart(2,'0'); const mm=String(d.getUTCMonth()+1).padStart(2,'0'); const yy=d.getUTCFullYear();
    return dd+'/'+mm+'/'+yy; }

  const waveMap={ PP:{1:[1,2,3,4],2:[5,6,7,8]}, SMRO:{1:[1,2,3,4],2:[5,6,7,8]}, SS:{1:[1,2,3],2:[4,5,6,7]} };
  function waveOf(dept,num){ const m=waveMap[dept]; for(const w in m){ if(m[w].includes(num)) return +w; } return null; }

  async function parseTrackerBlob(dept, blob){
    const files=await unzip(blob);
    const shared=parseShared(files['xl/sharedStrings.xml']);
    const sheets=wbSheets(files);
    const projects=[]; const issues=[];
    for(const sh of sheets){
      const g=parseSheet(files[sh.path]||'',shared);
      if(new RegExp('^'+dept+'-\\d+$').test(sh.name)){
        const num=+sh.name.split('-')[1];
        const name=((g['B4']&&g['B4'].val)||'').replace(/^اسم المشروع\s*:?\s*/,'').trim();
        const spec=(g['B5']&&g['B5'].val)||'';
        const log=[];
        for(let r=13;r<=72;r++){
          const stage=g['C'+r]&&g['C'+r].val;
          const upd=g['E'+r]&&g['E'+r].val;
          const note=g['F'+r]&&g['F'+r].val;
          const stt=g['G'+r]&&g['G'+r].val;
          const deliv=g['D'+r]&&g['D'+r].val;
          const date=g['B'+r]&&g['B'+r].val;
          if(stage||upd||note||stt||deliv){
            log.push({seq:(g['A'+r]&&g['A'+r].val)||'', date:xserial(date)||'', stage:stage||'', deliverable:deliv||'', update:upd||'', note:note||'', status:stt||''});
          }
        }
        const last=log[log.length-1]||{};
        projects.push({id:sh.name, num, name, specialist:spec, dept, wave:waveOf(dept,num),
          currentStage:last.stage||'', currentStatus:last.status||'', lastDate:last.date||'',
          lastUpdate:last.update||'', logCount:log.length, log});
      }
      if(sh.name==='ISSUES'){
        for(let r=4;r<=53;r++){
          const desc=g['B'+r]&&g['B'+r].val;
          if(desc){ issues.push({seq:(g['A'+r]&&g['A'+r].val)||'', desc, stage:(g['C'+r]&&g['C'+r].val)||'', impact:(g['D'+r]&&g['D'+r].val)||'', owner:(g['E'+r]&&g['E'+r].val)||'', status:(g['F'+r]&&g['F'+r].val)||'', start:xserial(g['G'+r]&&g['G'+r].val)||(g['G'+r]&&g['G'+r].val)||'', resolved:xserial(g['H'+r]&&g['H'+r].val)||''}); }
        }
      }
    }
    projects.sort((a,b)=>a.num-b.num);
    return {projects, issues};
  }

  async function parseOverviewBlob(blob){
    const files=await unzip(blob);
    const shared=parseShared(files['xl/sharedStrings.xml']);
    const sheets=wbSheets(files);
    const mt=sheets.find(s=>s.name==='Management Tracker');
    const out=[];
    if(mt){
      const g=parseSheet(files[mt.path],shared);
      for(let r=2;r<=20;r++){
        const topic=g['B'+r]&&g['B'+r].val;
        if(topic){ out.push({seq:(g['A'+r]&&g['A'+r].val)||'', topic:topic.trim(), date:xserial(g['C'+r]&&g['C'+r].val)||'', action:(g['D'+r]&&g['D'+r].val)||'', doneDate:xserial(g['E'+r]&&g['E'+r].val)||'', status:(g['F'+r]&&g['F'+r].val)||''}); }
      }
    }
    return out;
  }

  // Map a filename to a role
  function roleOf(name){
    const n=name.toLowerCase();
    if(/(^|[^a-z])pp([^a-z]|$)/.test(n)||n.includes('- pp')) return {kind:'tracker',dept:'PP'};
    if(n.includes('smro')) return {kind:'tracker',dept:'SMRO'};
    if(/(^|[^a-z])ss([^a-z]|$)/.test(n)||n.includes('- ss')) return {kind:'tracker',dept:'SS'};
    if(n.includes('overview')||n.includes('نظرة')) return {kind:'overview'};
    return null;
  }

  // Detect a role from the workbook's contents (sheet names) when the filename
  // doesn't make it obvious — so uploads work regardless of how they're named.
  async function detectRole(blob){
    try{
      const files=await unzip(blob);
      const names=wbSheets(files).map(s=>s.name);
      for(const dept of ['PP','SMRO','SS']){
        if(names.some(n=>new RegExp('^'+dept+'-\\d+$').test(n))) return {kind:'tracker',dept};
      }
      if(names.includes('Management Tracker')) return {kind:'overview'};
    }catch(e){}
    return null;
  }

  const DEPT_FILES = { PP:'uploads/Wave Tracker V2 - PP.xlsx', SMRO:'uploads/Wave Tracker V2 - SMRO.xlsx', SS:'uploads/Wave Tracker V2 - SS.xlsx' };
  const OVERVIEW_FILE = 'uploads/Waves Overview.xlsx';
  const RES_ID = { PP:'ppFile', SMRO:'smroFile', SS:'ssFile', overview:'overviewFile' };
  // When bundled as a standalone file, resources are inlined as blob URLs on window.__resources.
  function srcFor(key){ const R=window.__resources||{}; return R[RES_ID[key]] || (key==='overview'?OVERVIEW_FILE:DEPT_FILES[key]); }
  const DEPT_NAMES = { PP:'قسم PP', SMRO:'قسم SMRO', SS:'قسم SS' };

  // Build a full data object from a set of {role -> blob}
  async function buildFromBlobs(map){
    const result={departments:{}, managementTracker:[], deptNames:DEPT_NAMES};
    for(const dept of ['PP','SMRO','SS']){
      if(map['tracker:'+dept]) result.departments[dept]=await parseTrackerBlob(dept, map['tracker:'+dept]);
      else result.departments[dept]={projects:[],issues:[]};
    }
    if(map['overview']) result.managementTracker=await parseOverviewBlob(map['overview']);
    return result;
  }

  // Load directly from the bundled uploads/ files via fetch
  async function loadFromProject(){
    const map={};
    for(const dept of ['PP','SMRO','SS']){
      const res=await fetch(srcFor(dept)); if(res.ok) map['tracker:'+dept]=await res.blob();
    }
    const ov=await fetch(srcFor('overview')); if(ov.ok) map['overview']=await ov.blob();
    return buildFromBlobs(map);
  }

  // Load from a FileList the user uploaded; fall back to project files for any role not provided
  async function loadFromFiles(fileList){
    const map={}; const recognized=[];
    for(const f of fileList){
      const role=roleOf(f.name) || await detectRole(f);
      if(!role) continue;
      if(role.kind==='tracker'){ map['tracker:'+role.dept]=f; recognized.push(role.dept); }
      else if(role.kind==='overview'){ map['overview']=f; if(recognized.indexOf('Overview')<0) recognized.push('Overview'); }
    }
    // fill gaps from project
    for(const dept of ['PP','SMRO','SS']){
      if(!map['tracker:'+dept]){ try{ const r=await fetch(srcFor(dept)); if(r.ok) map['tracker:'+dept]=await r.blob(); }catch(e){} }
    }
    if(!map['overview']){ try{ const r=await fetch(srcFor('overview')); if(r.ok) map['overview']=await r.blob(); }catch(e){} }
    const result=await buildFromBlobs(map);
    result.recognized=recognized;
    return result;
  }

  window.WaveParser = { loadFromProject, loadFromFiles, roleOf };
})();
