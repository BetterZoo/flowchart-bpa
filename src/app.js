
"use strict";
/* ================= Libraries (inlined in this standalone build) ================= */
const LIBS = {
  xlsx: 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  jspdf: 'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js',
  autotable: 'https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js',
  docx: 'https://cdn.jsdelivr.net/npm/docx@9.15.3/dist/index.iife.js'
};
const _libs = {};
function loadScript(url){ return Promise.resolve(); }
function _unusedLoadScript(url){
  if(!_libs[url]) _libs[url] = new Promise((res,rej)=>{
    const s=document.createElement('script'); s.src=url; s.onload=()=>res(); s.onerror=()=>{delete _libs[url]; rej(new Error('Gagal memuat pustaka: '+url));};
    document.head.appendChild(s);
  });
  return _libs[url];
}

/* ================= Helpers ================= */
const $ = id => document.getElementById(id);
const esc = s => String(s==null?'':s).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const norm = s => String(s==null?'':s).trim().toLowerCase().replace(/\s+/g,' ');
const uniq = a => [...new Set(a)];
const cmpCode = (a,b) => { const sa=String(a), sb=String(b), x=sa.replace(/-G(\d+)$/,'.0.$1').split('.').map(Number), y=sb.replace(/-G(\d+)$/,'.0.$1').split('.').map(Number); for(let i=0;i<Math.max(x.length,y.length);i++){ const d=(x[i]||0)-(y[i]||0); if(d) return d; } return 0; };
function toast(msg, ms=2600){ const t=$('toast'); t.textContent=msg; t.classList.add('show'); clearTimeout(toast._t); toast._t=setTimeout(()=>t.classList.remove('show'), ms); }
let storageFailed=false;
const store = {
  get(k){ try{ return localStorage.getItem(k); }catch(e){ return null; } },
  set(k,v){ try{ localStorage.setItem(k,v); return true; }catch(e){ storageFailed=true; toast('Penyimpanan browser gagal. Export Excel untuk backup.',9000); return false; } }
};
const CODE_RE = /\b\d+(?:\.\d+){1,}(?:-G\d+)?\b/g; // v1.8.0: gateway codes, e.g. 12.01.04-G1
const oneEnd = () => { try{ return localStorage.getItem('flowbpa:oneEnd')!=='0'; }catch(e){ return true; } }; // v1.9.2
const isGateRow = r => !!r && /-G\d+$/.test(String(r.l5no||'')); // v1.8.0: a gateway that follows another gateway
const gateQ = r => String((r&&r.decisionQuestion)||'').replace(/^\s*\[[^\]]*\]\s*/,'') || String((r&&r.l5)||'').replace(/^Gateway:\s*/i,'');

/* ================= Parsing workbook ================= */
const COLS = {
  l1no:['no l1'], l1:['level 1'], l2no:['no l2'], l2:['level 2'], l3no:['no l3'], l3:['level 3'],
  l4no:['no l4'], l4:['level 4'], l5no:['no l5'], l5:['level 5'],
  desc:['activity description','deskripsi aktivitas','deskripsi'], role:['responsible role','peran penanggung jawab','pic','role'],
  dec:['decision ref','referensi keputusan'], inParty:['pihak pemberi input'], outParty:['pihak penerima output'],
  prev:['proses sebelum'], next:['proses sesudah'], noteRole:['catatan peran'], iface:['dasar antarmuka'], note:['catatan'],
  slaX:['sla'], outX:['dokumen/output','dokumen / output'], posX:['jabatan (khusus aktivitas ini)','jabatan khusus','jabatan aktivitas','jabatan'], unitX:['unit kerja','unit organisasi','divisi','unit'], decisionQuestion:['pertanyaan keputusan'], decisionRole:['role keputusan'], inputX:['input'], outputX:['output'], docX:['dokumen'], systemX:['sistem/aplikasi'], startLabel:['bpa label mulai'], endLabel:['bpa label selesai'], symbolDocs:['bpa label dokumen'], dataFrom:['dokumen dari l4 lain dipakai di'] // v1.7.1
};
const BPMN_META_SHEET='BPMN Metadata';
const BPMN_META_HEAD=['No L4','Kode L5','Versi','Jenis Gateway'];
function parseBpmnMetaSheet(wb){
 const ws=wb.Sheets[BPMN_META_SHEET];if(!ws)return null;
 const data=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:''});
 if(!data.length||BPMN_META_HEAD.some((h,i)=>String(data[0][i]||'').trim()!==h))throw Error('Header BPMN Metadata tidak sesuai format Master Workbook.');
 const result={},seen=new Set();
 for(let i=1;i<data.length;i++){
  const [l4,code,mode,type]=data[i].slice(0,4).map(v=>String(v||'').trim());
  if(![l4,code,mode,type].some(Boolean))continue;
  if(!/^\d+(?:\.\d+)+$/.test(l4)||!code.startsWith(l4+'.')||!['asis','tobe'].includes(mode)||!['XOR','OR','AND'].includes(type))throw Error('BPMN Metadata baris '+(i+1)+' tidak valid.');
  const id=l4+'|'+mode+'|'+code;if(seen.has(id))throw Error('Duplikasi BPMN Metadata: '+id);seen.add(id);
  (result[l4+'|'+mode] ||= {})[code]=type;
 }
 return result;
}
function appendBpmnMetaSheet(wb){
 const rows=[BPMN_META_HEAD],data=bpmnSemanticData();
 for(const [scope,items] of Object.entries(data)){
  const i=scope.lastIndexOf('|');if(i<0)continue;
  const l4=scope.slice(0,i),mode=scope.slice(i+1);
  if(!['asis','tobe'].includes(mode))continue;
  for(const [code,type] of Object.entries(items))if(['XOR','OR','AND'].includes(type))rows.push([l4,code,mode,type]);
 }
 const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=[{wch:14},{wch:18},{wch:12},{wch:20}];
 XLSX.utils.book_append_sheet(wb,ws,BPMN_META_SHEET);
}
function validateBpmnMetaAgainstRows(data,rows){
 if(!data)return;
 const by=new Map(rows.map(a=>[a.l5no,a]));
 for(const [scope,items] of Object.entries(data)){
  const [l4,mode]=scope.split('|');if(!['asis','tobe'].includes(mode))throw Error('Versi BPMN tidak dikenal: '+scope);
  for(const code of Object.keys(items)){
   const a=by.get(code);if(!a||a.l4no!==l4)throw Error('BPMN Metadata merujuk aktivitas tidak dikenal: '+code);
  }
 }
}
function chooseSopImportSheet(names){
  return new Promise(resolve=>{
    const overlay=document.createElement('div');
    overlay.style.cssText='position:fixed;inset:0;z-index:20000;background:rgba(15,23,42,.6);display:flex;align-items:center;justify-content:center;padding:20px';
    const dialog=document.createElement('div');dialog.setAttribute('role','dialog');dialog.setAttribute('aria-modal','true');dialog.setAttribute('aria-label','Pilih sheet SOP');
    dialog.style.cssText='background:#fff;color:#172033;border-radius:12px;padding:24px;width:min(480px,96vw);box-shadow:0 20px 60px #0005;font:14px Arial,sans-serif';
    const h=document.createElement('h2');h.textContent='Pilih sheet SOP untuk diimport';h.style.cssText='font-size:19px;margin:0 0 12px';
    const note=document.createElement('p');note.textContent='Workbook memiliki beberapa sheet proses yang valid. Pilih satu sebagai sumber aktivitas; sheet pendukung tetap dibaca sesuai kebutuhan.';note.style.cssText='line-height:1.5;margin:0 0 16px';
    const sel=document.createElement('select');sel.setAttribute('aria-label','Sheet SOP');sel.style.cssText='width:100%;padding:10px;border:1px solid #b6bfca;border-radius:7px;font-size:14px';
    for(const name of names){const opt=document.createElement('option');opt.value=name;opt.textContent=name;sel.appendChild(opt);}
    const official=names.find(n=>n.trim().toLowerCase()==='sop');if(official)sel.value=official;
    const buttons=document.createElement('div');buttons.style.cssText='display:flex;justify-content:flex-end;gap:10px;margin-top:20px';
    const cancel=document.createElement('button');cancel.type='button';cancel.textContent='Batal';cancel.style.cssText='padding:9px 16px;border:1px solid #cbd5e1;border-radius:6px;background:#fff;cursor:pointer';
    const ok=document.createElement('button');ok.type='button';ok.textContent='Import sheet';ok.style.cssText='padding:9px 16px;border:0;border-radius:6px;background:#1e3a5f;color:#fff;cursor:pointer';
    function finish(v){overlay.remove();resolve(v);}
    cancel.onclick=()=>finish(null);ok.onclick=()=>finish(sel.value);buttons.append(cancel,ok);
    dialog.append(h,note,sel,buttons);overlay.append(dialog);document.body.append(overlay);sel.focus();
  });
}
async function readWorkbook(buf){
  await loadScript(LIBS.xlsx);
  const wb = XLSX.read(buf, {type:'array'});
  const candidates=[];
  for(const sn of wb.SheetNames){
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, raw:false, defval:''});
    for(let i=0;i<Math.min(15,aoa.length);i++){
      const hdr=(aoa[i]||[]).map(norm);
      if(hdr.includes('no l5')&&hdr.includes('proses sesudah')){
        candidates.push(sn);break;
      }
    }
  }
  if(!candidates.length) throw new Error('Tidak ada sheet dengan kolom "No L5" dan "Proses Sesudah".');
  const chosen = candidates.length===1 ? candidates[0] : await chooseSopImportSheet(candidates);
  if(!chosen) throw new Error('Import dibatalkan: belum memilih sheet SOP.');
  for(const sn of [chosen]){
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, raw:false, defval:''});
    for(let i=0;i<Math.min(15,aoa.length);i++){
      const hdr = aoa[i].map(norm);
      if(hdr.includes('no l5') && hdr.includes('proses sesudah')){
        const col = {};
        for(const [k,names] of Object.entries(COLS)){ const j = hdr.findIndex(h=>names.includes(h)); if(j>=0) col[k]=j; }
        const need = ['l4no','l4','l5no','l5','next'];
        const miss = need.filter(k=>col[k]==null);
        if(miss.length) throw new Error('Kolom wajib tidak ditemukan: '+miss.join(', '));
        // v1.6.6: repair Master Workbooks saved by v1.3.x–v1.5.0 (Level 3 column missing, every later column shifted left).
        if(col.l3!=null && col.l4no!=null){
          const isCode=v=>/^\d+(\.\d+)+$/.test(String(v||'').trim());
          const sample=aoa.slice(i+1,i+21).filter(row=>String(row[col.l3]||'').trim());
          if(sample.length && sample.every(row=>isCode(row[col.l3]) && !isCode(row[col.l4no]))){
            for(let r=i+1;r<aoa.length;r++){ const row=aoa[r]; aoa[r]=[...row.slice(0,col.l3),'',...row.slice(col.l3)]; }
            state._repairedShift=true;
          }
        }
        const rows = [];
        for(let r=i+1;r<aoa.length;r++){
          const row = aoa[r]; const o = {};
          for(const k of Object.keys(COLS)) o[k] = col[k]==null ? '' : String(row[col[k]]==null?'':row[col[k]]).trim();
          if(!o.l5no) continue;
          if(!o.outX) o.outX=[o.outputX,o.docX].map(v=>String(v||'').trim()).filter(v=>v&&v.toUpperCase()!=='X').join(' / '); /* v1.6.1: "X" is a marker, not a name */ rows.push(o);
        }
        if(!rows.length) throw new Error('Sheet "'+sn+'" tidak berisi aktivitas.');
        if(new Set(rows.map(r=>r.l5no)).size!==rows.length) throw Error('Kode L5 harus unik dalam workbook.'); if(new Set(rows.map(r=>r.l4no+'|'+r.l1no+'|'+r.l2no+'|'+r.l3no)).size!==new Set(rows.map(r=>r.l4no)).size) throw Error('Kode L4 harus unik lintas hierarki.'); const bpmn=parseBpmnMetaSheet(wb);validateBpmnMetaAgainstRows(bpmn,rows);return {sheet:sn, rows, bpmn, used:readUsedCodes(wb), elab: parseElabSheet(wb), pos: parsePosSheets(wb), extra: parseExtraSheet(wb), roles: parseRoleSheet(wb), analysis:readAnalysisSheet(wb), register:readRegisterSheet(wb), master:wb.SheetNames.includes('Riwayat Kode')&&wb.SheetNames.includes('SOP')};
      }
    }
  }
  throw new Error('Tidak ada sheet dengan kolom "No L5" dan "Proses Sesudah".');
}

function parseFlow(s){
  s = String(s||'').trim(); if(!s) return [];
  const out = []; let lastLabel = null, anyLabel = false;
  for(const p of s.split(/;|\n/).map(x=>x.trim()).filter(Boolean)){
    let label = null, rest = p;
    const m = p.match(/^([^:]+):\s*(.+)$/);
    if(m && !/^\s*\d/.test(m[1])){ label = m[1].trim(); rest = m[2]; anyLabel = true; lastLabel = label; }
    else if(anyLabel) label = lastLabel;
    // v1.5.0: "A + B (catatan)" — each part keeps its own note, so a parallel hand-off is not misread as a decision.
    for(const part of rest.split(/\s*\+\s*/).filter(Boolean)){
      const note = ((part.match(/\(([^)]*)\)/)||[])[1]||'').trim();
      const bare = part.replace(/\([^)]*\)/g,' ');
      const ts = [];
      if(/\b(SELESAI|END)\b/i.test(bare)) ts.push({kind:'end'});
      if(/\b(MULAI|START)\b/i.test(bare)) ts.push({kind:'start'});
      (bare.match(CODE_RE)||[]).forEach(c=>ts.push({kind:'code', code:c}));
      if(!ts.length) out.push({label, kind:'unknown', raw:part, note});
      for(const t of ts) out.push({label, note, ...t});
    }
  }
  return out;
}
const splitParty = s => String(s||'').split(/;|\n/).map(x=>x.trim()).filter(Boolean);
function roleInfo(role){ if(String(role).startsWith('unit::')) return {code:'',name:String(role).slice(6),pos:true}; if(String(role).startsWith('pos::')){ return {code:'', name:String(role).slice(5), pos:true}; } const m = String(role).match(/^(PF-\d+)\s*(.*)$/i); const r = m ? {code:m[1].toUpperCase(), name:(m[2]||m[1]).trim()} : {code:'', name:role};
  const al = (typeof state!=='undefined' && state.roleEdit && state.roleEdit.alias[role]); if(al){ r.orig = r.name; r.name = al; } return r; }
const STANDARD_ROLES = ['Pemrakarsa','Pelaksana','Pemeriksa','Pemberi Persetujuan','Penerima'];
const STANDARD_EN = ['Initiator','Executor','Reviewer','Approver','Recipient'];
function standardRank(r){ const n=String(r||'').trim().toLowerCase(); return STANDARD_ROLES.findIndex((v,i)=>[v,STANDARD_EN[i],v+' ('+STANDARD_EN[i]+')','R0'+(i+1)].some(x=>x.toLowerCase()===n)); }
const roleNum = r => { const i=standardRank(r); if(i>=0)return i+1; const m = String(r).match(/PF-(\d+)/i); return m ? +m[1] : 999; };
const sortRoles = a => a.slice().sort((x,y)=>roleNum(x)-roleNum(y) || x.localeCompare(y));

/* ================= State ================= */
const state = { fileName:'', rows:[], byCode:new Map(), l4s:new Map(), warnings:[], current:null, zoom:1, diagram:null,
  elab:{}, extra:{}, roleEdit:{alias:{}, added:{}, hidden:{}, order:{}}, posmap:{}, decmap:{}, unitmap:{}, laneModes:{}, posOn:true, colorView:false, showProposed:true, sideMode:null, editing:null, newItem:null, aiCtl:null, aiMsg:null, sampleAvail:undefined };

/* ================= Data Validation v1.2 (baseline workbook, read-only) ================= */
function validateWorkbookRows(rows){
  const issues = [], by = new Map(), counts = new Map(), edges = new Map(), incoming = new Map();
  const add = (r, severity, msg) => issues.push({code:r.l5no, l4:r.l4no, severity, msg});
  for(const r of rows){
    if(['docX','systemX'].some(k=>String(r[k]||'').trim().toUpperCase()==='X') && !String(r.outputX||'').trim() && !String(r.inputX||'').trim()) add(r,'warning','Kolom '+(String(r.docX||'').trim()?'Dokumen (U)':'Sistem/Aplikasi (W)')+' berisi "X" tetapi Input (S) dan Output (T) kosong, jadi nama dokumen/data tidak diketahui dan simbolnya tidak digambar. Isi Output (T) untuk dokumen yang dihasilkan, Input (S) untuk dokumen yang diterima, atau tulis namanya langsung di kolom tersebut.');
    if(String(r.docX||'').trim()&&String(r.systemX||'').trim())add(r,'warning','Kolom Dokumen (U) dan Sistem/Aplikasi (W) sama-sama terisi. Canvas menampilkan Data Object terlebih dahulu; periksa apakah keduanya memang diperlukan.');
    counts.set(r.l5no,(counts.get(r.l5no)||0)+1); by.set(r.l5no,r); }
  for(const r of rows){
    if(counts.get(r.l5no)>1) add(r,'error','Kode L5 duplikat: '+r.l5no+'.');
    if(!r.l4no || !r.l4 || !r.l5) add(r,'error','Kode/nama L4 atau nama aktivitas kosong.');
    if(!r.role) add(r,'error','Responsible Role kosong.');
    const nx=parseFlow(r.next), pv=parseFlow(r.prev);
    if(!nx.length) add(r,'error','Proses Sesudah kosong.');
    if(!pv.length) add(r,'warning','Proses Sebelum kosong.');
    for(const [field,flow] of [['Proses Sesudah',nx],['Proses Sebelum',pv]]) for(const f of flow){
      if(f.kind==='unknown') add(r,'warning',field+' berisi kondisi/keterangan tanpa kode tujuan: "'+f.raw+'". Periksa apakah perlu jalur atau hanya catatan.');
      if(f.kind==='code' && !by.has(f.code)){ const isL4=rows.some(x=>x.l4no===f.code); add(r,isL4?'warning':'error',field+' menyebut '+f.code+(isL4?' sebagai referensi L4 (bukan aktivitas L5). Periksa hubungan lintas-proses.':', yang tidak ada di workbook.')); }
      if((field==='Proses Sesudah' && f.kind==='start') || (field==='Proses Sebelum' && f.kind==='end')) add(r,'warning',field+' memiliki terminator pada arah yang tidak lazim.');
    }
    { // v1.4.1: same branch label, several destinations (other than "SELESAI + kode L4 lain")
      const byLab=new Map();
      for(const f of nx){ if(!f.label||!['code','end'].includes(f.kind))continue; if(!byLab.has(f.label))byLab.set(f.label,[]); byLab.get(f.label).push(f); }
      // A label written in one segment with "+" (e.g. "Terpenuhi: 07.05.11 + 04.04.01") is an intentional parallel split.
      const segCount=new Map(); for(const seg of String(r.next||'').split(/;|\n/)){ const m=seg.match(/^\s*([^:]+):/); if(m&&!/^\s*\d/.test(m[1])){ const k=m[1].trim(); segCount.set(k,(segCount.get(k)||0)+1); } }
      for(const [lab,fs] of byLab){
        if((segCount.get(lab)||0)<2) continue;
        const local=fs.filter(f=>f.kind==='code' && by.has(f.code) && by.get(f.code).l4no===r.l4no);
        const cross=fs.filter(f=>f.kind==='code' && !(by.has(f.code) && by.get(f.code).l4no===r.l4no));
        const ends=fs.filter(f=>f.kind==='end');
        if(local.length>1 || (local.length && (ends.length || cross.length)))
          add(r,'warning','Cabang "'+lab+'" menuju beberapa tujuan ('+fs.map(f=>f.kind==='end'?'SELESAI':f.code).join(', ')+'). Pastikan memang berjalan bersamaan; jika tidak, beri label cabang yang berbeda.');
      }
    }
    const targets=nx.filter(f=>f.kind==='code' && by.has(f.code)).map(f=>f.code);
    edges.set(r.l5no,targets);
    for(const c of targets){ if(!incoming.has(c)) incoming.set(c,new Set()); incoming.get(c).add(r.l5no); }
    for(const c of new Set(targets)) if(!parseFlow(by.get(c).prev).some(f=>f.kind==='code' && f.code===r.l5no)) add(r,'warning','Proses Sesudah mengarah ke '+c+', tetapi Proses Sebelum tujuan tidak mencantumkan '+r.l5no+'.');
    for(const c of new Set(pv.filter(f=>f.kind==='code' && by.has(f.code)).map(f=>f.code))) if(!parseFlow(by.get(c).next).some(f=>f.kind==='code' && f.code===r.l5no)) add(r,'warning','Proses Sebelum menyebut '+c+', tetapi Proses Sesudah '+c+' tidak mengarah ke '+r.l5no+'.');
  }
  const reachable=new Set(), queue=rows.filter(r=>parseFlow(r.prev).some(f=>f.kind==='start')).map(r=>r.l5no);
  for(let i=0;i<queue.length;i++){ const c=queue[i]; if(reachable.has(c)) continue; reachable.add(c); for(const t of edges.get(c)||[]) if(!reachable.has(t)) queue.push(t); }
  for(const r of rows){
    const start=parseFlow(r.prev).some(f=>f.kind==='start'), end=parseFlow(r.next).some(f=>f.kind==='end');
    if(!start && !end && !(edges.get(r.l5no)||[]).length && !incoming.has(r.l5no)) add(r,'warning','Aktivitas terisolasi: tidak memiliki hubungan valid atau terminator.');
    else if(!reachable.has(r.l5no)) add(r,'warning','Aktivitas tidak dapat dicapai dari MULAI melalui Proses Sesudah. Periksa jalur masuk.');
  }
  const exits=new Set(rows.filter(r=>parseFlow(r.next).some(f=>f.kind==='end')).map(r=>r.l5no)), reverse=[...exits];
  for(let i=0;i<reverse.length;i++)for(const c of incoming.get(reverse[i])||[])if(!exits.has(c)){exits.add(c);reverse.push(c);}
  for(const r of rows)if(!exits.has(r.l5no))add(r,'warning','Tidak memiliki jalur menuju SELESAI.');
  return issues;
}

function ingest(rows, fileName){
  state.rows = rows; state.fileName = fileName; reserveCodes(rows); rememberCodes();
  state.byCode = new Map(rows.map(r=>[r.l5no, r]));
  state.l4s = new Map();
  for(const r of rows){
    let g = state.l4s.get(r.l4no);
    if(!g){ g = {l4no:r.l4no, l4:r.l4, l3no:r.l3no, l3:r.l3, l2no:r.l2no, l2:r.l2, l1no:r.l1no, l1:r.l1, acts:[], warn:0}; state.l4s.set(r.l4no, g); }
    g.acts.push(r);
  }
  // data checks (read-only; does not alter diagram generation)
  state.warnings = validateWorkbookRows(rows);
  for(const w of state.warnings){ const g=state.l4s.get(w.l4); if(g) g.warn++; }
  let seeded = 0;
  for(const r of rows){ if(!(r.slaX || r.outX)) continue; const x = state.extra[r.l5no] || (state.extra[r.l5no] = {}); if(r.slaX && !x.sla){ x.sla = r.slaX; seeded++; } if(r.outX && !x.out){ x.out = r.outX; seeded++; } }
  if(seeded) saveExtra();
  updatePosToggle();
  store.set('flowbpa:data', JSON.stringify({fileName, rows}));
  $('fileInfo').textContent = fileName+' · '+state.l4s.size+' subproses · '+rows.filter(r=>!isGateRow(r)).length+' aktivitas';
  $('search').disabled = false; $('btnExport').disabled = false; $('btnNew').disabled = false;
  $('empty').hidden = true; $('canvas').hidden = false; $('toolbar').hidden = false;
  renderTree(); renderWarnings();
  const fromHash = hashToL4();
  select(fromHash && state.l4s.has(fromHash) ? fromHash : [...state.l4s.keys()].sort(cmpCode)[0], true);
}

/* ================= Sidebar ================= */
function renderTree(){
  const q = norm($('search').value);
  const groups = [...state.l4s.values()].sort((a,b)=>cmpCode(a.l4no,b.l4no));
  const match = g => !q || norm(g.l4no+' '+g.l4+' '+g.l3+' '+g.l2).includes(q) || g.acts.some(a=>norm(a.l5no+' '+a.l5+' '+a.desc).includes(q));
  let html = '', lastL2 = null, lastL3 = null, open = false;
  const shown = groups.filter(match);
  for(const g of shown){
    if(g.l2no!==lastL2){ if(open){ html+='</details>'; open=false; } html += '<div class="l2h">'+esc(g.l2no+' · '+g.l2)+'</div>'; lastL2=g.l2no; lastL3=null; }
    if(g.l3no!==lastL3){
      if(open) html+='</details>';
      const n = shown.filter(x=>x.l3no===g.l3no && x.l2no===g.l2no).length;
      const isOpen = q || (state.current && state.l4s.get(state.current) && state.l4s.get(state.current).l3no===g.l3no);
      html += '<details class="l3"'+(isOpen?' open':'')+'><summary>'+esc(g.l3no+' '+g.l3)+'<span class="cnt">'+n+'</span></summary>';
      open = true; lastL3 = g.l3no;
    }
    html += '<a class="l4'+(g.l4no===state.current?' active':'')+'" data-l4="'+esc(g.l4no)+'"><span class="code">'+esc(g.l4no)+'</span><span>'+esc(g.l4)+'</span>'+propBadge(g.l4no)+(g.warn?'<span class="warn" title="Ada catatan data">⚠ '+g.warn+'</span>':'')+'</a>';
  }
  if(open) html+='</details>';
  if(!shown.length) html = '<div class="l2h">Tidak ada yang cocok.</div>';
  $('tree').innerHTML = html;
}
function propBadge(l4no){ const p = (state.elab[l4no]||[]).filter(i=>i.status==='proposed').length; return p ? '<span class="prop" title="Usulan elaborasi belum direviu">✦ '+p+'</span>' : ''; }
function renderWarnings(){
  const mode=$('processEditMode').value;
  const W = mode==='tobe'?validateWorkbookRows([...state.l4s.values()].flatMap(g=>{try{return structuralRows(g,'tobe');}catch(e){return g.acts;}})):state.warnings;
  if(mode==='tobe')for(const g of state.l4s.values())try{const rr=structuralRows(g,'tobe');if(!Array.isArray(rr)||!rr.length||rr.some(r=>r.l4no!==g.l4no))throw Error('invalid');}catch(e){W.push({code:g.acts[0].l5no,l4:g.l4no,severity:'error',msg:'Snapshot TO-BE tidak valid.'});}
  $('warnBox').hidden = !state.rows.length;
  const errors=W.filter(w=>w.severity==='error').length, warnings=W.length-errors;
  $('warnSum').textContent = 'Validasi data · '+errors+' error · '+warnings+' warning';
  $('warnList').innerHTML = W.length ? W.map(w=>'<li><b>'+ (w.severity==='error'?'ERROR':'WARNING') +'</b> · <a data-l4="'+esc(w.l4)+'" data-code="'+esc(w.code)+'">'+esc(w.code)+'</a> '+esc(w.msg)+'</li>').join('') : '<li>Tidak ada temuan pada pemeriksaan data baseline.</li>';
}

/* ================= Text measuring ================= */
const FONT = 'Arial, Helvetica, sans-serif';
const _mctx = document.createElement('canvas').getContext('2d');
function tw(text, size, bold){ _mctx.font = (bold?'bold ':'')+size+'px '+FONT; return _mctx.measureText(text).width; }
function wrap(text, size, bold, maxW){
  const words = String(text||'').split(/\s+/).filter(Boolean); const lines = []; let cur = '';
  for(const w of words){ const t = cur ? cur+' '+w : w; if(tw(t,size,bold) <= maxW || !cur) cur = t; else { lines.push(cur); cur = w; } }
  if(cur) lines.push(cur);
  return lines.length ? lines : [''];
}

/* ================= Diagram model + layout ================= */
const G = { m:16, headerH:58, poolH:28, laneHeadH:48, laneW:440, gut:28, panelW:430, panelGap:24, band:30 };
/* v1.6.0: job-level rank for default lane order (lower = left). */
const JABATAN_RANK=[[/direktur utama|president director|\bceo\b/,7],[/direktur|director|\bsevp\b|\bevp\b/,6],[/vice president|\bvp\b|general manager|senior manager|kepala divisi/,5],[/manager|manajer|kepala bagian|kepala departemen/,4],[/asisten manajer|assistant manager|supervisor|kepala seksi|\bkasi\b/,3],[/officer|analis|analyst|spesialis|specialist|koordinator|engineer/,2],[/staff|\bstaf\b|pelaksana|admin|clerk|operator|juru/,1]];
function jabatanRank(name){ const first=String(name||'').split(' / ')[0].split(':').pop().toLowerCase(); for(const [re,r] of JABATAN_RANK) if(re.test(first)) return r; return 2.5; }
function actLaneKey(l4no, a){ const role=a.role||'(Tanpa peran)', mode=laneMode(l4no); const unit=mode==='unit'?unitFor(l4no,role,a):''; const pos=mode==='position'?posFor(l4no,role,a.dec,a):''; return unit?'unit::'+unit:pos?'pos::'+pos:role; }
function addedLaneKey(l4no, r){ const mode=laneMode(l4no); const u=(state.unitmap[l4no]||{})[r]||'', p=(state.posmap[l4no]||{})[r]||''; return mode==='unit'&&u?'unit::'+u:mode==='position'&&p?'pos::'+p:r; }
const RAIL_STEP=88;
function railEntryY(n,i){ return Math.min(n.cy+n.h/2-6, n.cy-6+i*9); } // v1.6.3: association entry points stay inside the task edge
function slaOff(n){ return n&&n.sla?30:0; } // v1.7.0: SLA annotation sits at the top of the rail, icons move down
function railExtent(n){ const k=n.crossInputs?.length||0, xo=n.extraOuts?.length||0; return Math.max(n.h/2, k*RAIL_STEP + (n.tag?(xo?92:78):0) + xo*RAIL_STEP - (xo?10:0) + slaOff(n) + (n.sla?14:0) + 4); } // v1.6.2: artifact rail starts at the task centre and grows downward
function legendLabel(kind){ // v1.6.4: reader-friendly legend names, editable in Pengaturan
  const id=kind==='store'?'legendData':'legendDoc', def=kind==='store'?'Data di sistem (Data Store)':'Dokumen (Data Object)';
  try{ const el=document.getElementById(id); const v=(el&&el.value||'').trim(); return v||def; }catch(e){ return def; } }
function cleanOut(v){ return String(v||'').replace(/(^|\s*\/\s*)X\s*$/i,'').trim().split(/\s*\n+\s*/).filter(Boolean).join(' · '); } // v1.6.1: drop a stray "X" marker from older saved output labels
function showMergeGw(){ try{ const el=document.getElementById('mergeGw'); return !!(el&&el.checked); }catch(e){ return false; } }
function clampWrap(text,size,bold,w,max){ const L=wrap(String(text||''),size,bold,w); if(L.length<=max) return L; const out=L.slice(0,max); let last=out[max-1]; while(last.length>1 && tw(last+'…',size,bold)>w) last=last.replace(/\s*\S*$/,'')||last.slice(0,-1); out[max-1]=last+'…'; return out; }
function poolName(){ try{ const el=document.getElementById('poolLabel'); const v=(el&&el.value||'').trim(); return v||'Organisasi'; }catch(e){ return 'Organisasi'; } }
const C = {
  navy:'#8a1e45', laneHead:'#f3ece9', laneA:'#ffffff', laneB:'#faf7f5', grid:'#bdb3ae',
  box:'#fdf6ef', boxS:'#58595b', teal:'#d9680a', decF:'#fff3d6', decS:'#e0a010',
  termF:'#fbecee', termS:'#a3212f', docF:'#fff8e8', docS:'#c08a1e', offF:'#f3ece9', offS:'#58595b',
  edge:'#2b2b2b', back:'#a3212f', text:'#2b2b2d', muted:'#6b6b70', amber:'#f2a71b',
  prop:'#7c3aed', propF:'#f5f0ff', newF:'#fff7ed', newS:'#c2410c', chgS:'#d97706', ink:'#58595b', termT:'#7a1a2a', docT:'#6b4e00', holdT:'#8a1c1c', dataT:'#3b2f86', propT:'#4c1d95', holdF:'#fdecec', holdS:'#c0262d', dataF:'#efeafd', dataS:'#6d5bd0'
};

// Grayscale palette shared by canvas and exported diagrams.
const MONO = { navy:'#292929', amber:'#ffffff', laneHead:'#ececec', laneA:'#ffffff', laneB:'#fafafa', grid:'#8c8c8c', box:'#ffffff', boxS:'#1a1a1a', teal:'#1a1a1a',
  decF:'#ffffff', decS:'#1a1a1a', termF:'#ffffff', termS:'#1a1a1a', docF:'#ffffff', docS:'#1a1a1a', offF:'#ffffff', offS:'#1a1a1a',
  edge:'#1a1a1a', back:'#1a1a1a', text:'#111111', muted:'#555555', prop:'#777777', propF:'#ffffff', newF:'#ffffff', newS:'#000000', chgS:'#000000',
  holdF:'#ffffff', holdS:'#1a1a1a', dataF:'#ffffff', dataS:'#1a1a1a', ink:'#111111', termT:'#111111', docT:'#111111', holdT:'#111111', dataT:'#111111', propT:'#111111' };
let CROSS_VER='asis'; // v1.6.5: which version of OTHER L4s feeds cross-L4 inputs (asis / tobe / approved)
function crossSourceRow(code){
  const r=state.byCode && state.byCode.get(code); if(!r || CROSS_VER==='asis') return r;
  try{ const g=elabFor(r.l4no, CROSS_VER==='tobe').filter(i=>i.type==='graph').slice(-1)[0]; if(!g) return r; const rows=JSON.parse(g.text); return rows.find(x=>x.l5no===code) || null; }catch(e){ return r; }
}
function effectiveData(a){ // v1.6.5: same rule as the output icon of the source task
  if(!a) return null; const doc=String(a.docX||'').trim(), sys=String(a.systemX||'').trim();
  if(doc==='-') return null;
  const kind=doc?'object':sys?'store':(cleanOut((state.extra[a.l5no]||{}).out)?'object':null); if(!kind) return null;
  const raw=kind==='store'?sys:doc, named=raw && raw.toUpperCase()!=='X' ? raw : '';
  const label=named || cleanOut((state.extra[a.l5no]||{}).out) || String(a.outputX||a.outX||'').trim().split(/\s*\n+\s*/).filter(Boolean).join(' · ');
  if(!label) return null; // v1.6.7: an "X" that marks only the activity's INPUT produces nothing downstream
  return {kind, label};
}
function buildDiagram(l4, statusLabel, opts){
  CROSS_VER = (opts && opts.ver) || ((opts && opts.items && opts.items.length) ? 'tobe' : 'asis');
  const graph=(opts?.items||[]).filter(i=>i.type==='graph').slice(-1)[0];
  if(graph){ try{ l4={...l4,acts:JSON.parse(graph.text)}; opts={...opts,items:opts.items.filter(i=>i.type!=='graph'&&!consumedIds(l4.acts).has(i.id))}; }catch(e){console.error(e);} }

  if(!(opts && opts.mono)) return _buildDiagram(l4, statusLabel, opts);
  const saved = {...C}; Object.assign(C, MONO);
  try{ return _buildDiagram(l4, statusLabel, opts); } finally { Object.assign(C, saved); }
}
function decQuestion(flow){
  const labels = uniq(flow.filter(f=>f.label).map(f=>f.label));
  if(labels.some(l=>/^disetujui$/i.test(l))) return 'Disetujui?';
  const pos = labels.find(l=>!/^(tidak|belum|gagal)\b/i.test(l)) || labels[0] || 'Keputusan';
  return pos.charAt(0).toUpperCase()+pos.slice(1)+'?';
}

function hasFormalDecision(a){
  const flow = parseFlow(a.next);
  const labels = uniq(flow.filter(f=>f.label).map(f=>f.label));
  if(labels.length >= 2) return true;
  const tg = flow.filter(f=>f.kind==='code' || f.kind==='end');
  return !labels.length && tg.length>=2 && tg.length===flow.length && tg.every(f=>f.note);
}
function flowOrder(rows){
 const sorted=rows.slice().sort((a,b)=>cmpCode(a.l5no,b.l5no)), by=new Map(sorted.map(r=>[r.l5no,r])), seen=new Set(), out=[];
 function visit(r){if(!r||seen.has(r.l5no))return;seen.add(r.l5no);out.push(r);for(const f of parseFlow(r.next))if(f.kind==='code')visit(by.get(f.code));}
 sorted.filter(r=>parseFlow(r.prev).some(f=>f.kind==='start'||(f.kind==='code'&&!by.has(f.code)))).forEach(visit);sorted.forEach(visit);
 // Remove only DFS back edges from the layout graph; never from process data.
 const mark=new Map(), adj=new Map(out.map(r=>[r.l5no,[]])), deg=new Map(out.map(r=>[r.l5no,0]));
 function scan(c){mark.set(c,1);for(const f of parseFlow(by.get(c).next)){if(f.kind!=='code'||!by.has(f.code)||mark.get(f.code)===1)continue; if(!adj.get(c).includes(f.code)){adj.get(c).push(f.code);deg.set(f.code,deg.get(f.code)+1);}if(!mark.has(f.code))scan(f.code);}mark.set(c,2);}
 out.forEach(r=>{if(!mark.has(r.l5no))scan(r.l5no);});
 const pending=new Set(out.map(r=>r.l5no)), result=[];while(pending.size){const r=out.find(r=>pending.has(r.l5no)&&deg.get(r.l5no)===0);if(!r)return out;pending.delete(r.l5no);result.push(r);for(const c of adj.get(r.l5no))deg.set(c,deg.get(c)-1);}return result;
}
function _buildDiagram(l4, statusLabel, opts){
  opts = opts || {};
  const acts = flowOrder(l4.acts);
  const inL4 = new Set(acts.map(a=>a.l5no));
  const byType = {gate:new Map(), hold:new Map(), doc:new Map(), datain:new Map(), declane:new Map()};
  for(const it of (opts.items||[])){ const m = byType[it.type]; if(!m || !inL4.has(it.after)) continue; if(!m.has(it.after)) m.set(it.after, []); m.get(it.after).push(it); }
  const drawn = [];
  // ---- structural changes: proses baru, pindah urutan, ubah PIC ----
  const flowOf = new Map(acts.map(a=>[a.l5no, parseFlow(a.next)]));
  const startOf = new Set(acts.filter(a=>parseFlow(a.prev).some(r=>r.kind==='start')).map(a=>a.l5no));
  const startNote = new Map(acts.map(a=>[a.l5no, (parseFlow(a.prev).find(r=>r.kind==='start')||{}).note||'']));
  const roleOv = new Map(), chg = new Map(), byCodeL = new Map(acts.map(a=>[a.l5no,a]));
  let order = acts.map(a=>a.l5no);
  const linear = c => { const f = flowOf.get(c)||[]; return f.length===1 && !f[0].label && (f[0].kind==='code' || f[0].kind==='end'); };
  const addChg = (c, it) => { if(!chg.has(c)) chg.set(c, []); chg.get(c).push(it); };
  for(const it of (opts.items||[])) if(it.type==='rename' && byCodeL.has(it.after) && it.text){
    const a = byCodeL.get(it.after); byCodeL.set(it.after, {...a, l5:it.text, origL5:a.origL5||a.l5}); addChg(it.after, it); drawn.push(it); }
  for(const it of (opts.items||[])){
    if(it.type==='role' && byCodeL.has(it.after) && it.role){ roleOv.set(it.after, it.role); addChg(it.after, it); drawn.push(it); }
    else if(it.type==='reorder' && byCodeL.has(it.subject) && inL4.has(it.after) && it.subject!==it.after && linear(it.subject)){
      const X = it.subject, Y = it.after, S = flowOf.get(X)[0];
      for(const [c,f] of flowOf) if(c!==X) f.forEach((t,i)=>{ if(t.kind==='code' && t.code===X) f[i] = {...S, label:t.label, note:t.note||S.note||''}; });
      if(startOf.has(X)){ startOf.delete(X); if(S.kind==='code' && inL4.has(S.code)) startOf.add(S.code); }
      flowOf.set(X, flowOf.get(Y)); flowOf.set(Y, [{kind:'code', code:X, label:null, note:''}]);
      order = order.filter(c=>c!==X); order.splice(order.indexOf(Y)+1, 0, X);
      addChg(X, it); drawn.push(it);
    }
    else if(it.type==='newact' && inL4.has(it.after) && it.text){
      const Y = it.after, base = byCodeL.get(Y);
      const pa = {l5no:it.id, l5:it.text, role: it.role || (roleOv.get(Y) || (base && base.role) || ''), desc:'', prev:'', next:'', inParty:'', outParty:'', isNew:true, elab:it,
        l4no:l4.l4no, l4:l4.l4, l3:l4.l3, l2:l4.l2, l1:l4.l1};
      byCodeL.set(it.id, pa); inL4.add(it.id);
      flowOf.set(it.id, flowOf.get(Y)); flowOf.set(Y, [{kind:'code', code:it.id, label:null, note:''}]);
      order.splice(order.indexOf(Y)+1, 0, it.id);
      drawn.push(it);
    }
  }
  const baseRole = a => roleOv.get(a.l5no) || a.role || '(Tanpa peran)';
  const laneRoles = new Map();   // lane key -> roles that sit in it (for the lane subtitle)
  const keyFor = (role, dec, act) => { const mode = laneMode(l4.l4no); const unit = mode==='unit' ? unitFor(l4.l4no,role,act) : ''; const pos = mode==='position' ? posFor(l4.l4no, role, dec, act) : ''; const k = unit ? 'unit::'+unit : pos ? 'pos::'+pos : role; if(!laneRoles.has(k)) laneRoles.set(k, new Set()); laneRoles.get(k).add(role); return k; };
  const R = a => keyFor(baseRole(a), a.dec, a);
  const seq = order.map(c=>byCodeL.get(c));
  const noteNo = new Map(); { let ni = 0; for(const a of seq){ const x = state.extra[a.l5no]; if(x && x.note && !a.isNew) noteNo.set(a.l5no, ++ni); } }
  // Order lanes by first actual involvement for organization views; keep standard role order in role view.
  const firstLane = new Map();
  // v1.6.0: default = job level (Staff → … → Direktur), ties by first involvement; a manual order per L4 wins.
  const laneRank = k => { if(String(k).startsWith('unit::')){ const rs=seq.filter(a=>R(a)===k).map(a=>jabatanRank(a.posX)); return rs.length?Math.min(...rs):2.5; } const ri=roleInfo(k); return ri.pos ? jabatanRank(ri.name) : 2.5; };
  const manualOrder = (state.roleEdit.order||{})[l4.l4no] || null;
  const sortKeys = a => { const auto=a.slice().sort((x,y)=>{
      const fx=firstLane.get(x)??Infinity, fy=firstLane.get(y)??Infinity;
      return (laneRank(x)-laneRank(y)) || (fx===fy?0:fx-fy) || x.localeCompare(y);
    });
    if(!manualOrder) return auto;
    return [...manualOrder.filter(k=>auto.includes(k)), ...auto.filter(k=>!manualOrder.includes(k))];
  };
  // decision / gateway diamonds may sit in another column than their activity
  const diamondTarget = (a, kind) => {
    if(kind==='gate'){ const g = (byType.gate.get(a.l5no)||[])[0]; if(g && g.role) return g; }
    const d = (byType.declane.get(a.l5no)||[])[0]; return d && d.role ? d : (kind==='dec' && hasFormalDecision(a) && a.decisionRole ? {role:a.decisionRole} : null);
  };
  // lane key for a moved diamond: the exact column when one was picked (pos), else the role's usual column
  const diamondKey = (a, kind) => { const t = diamondTarget(a, kind); if(!t) return null;
    if(t.pos && laneMode(l4.l4no)==='position'){ const k = 'pos::'+t.pos; if(!laneRoles.has(k)) laneRoles.set(k, new Set()); laneRoles.get(k).add(t.role); return k; }
    const candidates=seq.filter(x=>baseRole(x)===t.role); const match=candidates.length===1?candidates[0]:null; return keyFor(t.role, '',t.role===baseRole(a)?a:match); };
  const diamondRole = (a, kind) => { const t = diamondTarget(a, kind); return t ? t.role : null; };
  const extraKeys = [];
  const touchLane=k=>{if(k&&!firstLane.has(k))firstLane.set(k,firstLane.size);};
  if(laneMode(l4.l4no)==='role'){ for(const a of flowOrder(state.l4s.get(l4.l4no)?.acts||[])){touchLane(keyFor(a.role||'(Tanpa peran)','',a));if(a.decisionRole)touchLane(keyFor(a.decisionRole,''));} }
  for(const a of seq){ touchLane(R(a)); for(const k of ['dec','gate']){ const r = diamondKey(a, k); if(r){extraKeys.push(r);touchLane(r);} } }
  for(const r of (state.roleEdit.added[l4.l4no]||[])) extraKeys.push(keyFor(r, ''));
  // Retain baseline role columns in TO-BE so reassignment moves the activity, not the header.
  // Only roles actually used in this L4 baseline or active version are retained.
  if(laneMode(l4.l4no)==='role'){
    const baseline=state.l4s.get(l4.l4no);
    for(const a of (baseline?.acts||[])){
      const k=keyFor(a.role||'(Tanpa peran)','',a); extraKeys.push(k);
      if(a.decisionRole)extraKeys.push(keyFor(a.decisionRole,''));
    }
  }
  const actRoles = sortKeys(uniq([...seq.map(R), ...extraKeys]));
  const hiddenP = new Set(state.roleEdit.hidden[l4.l4no]||[]);
  const partyKey = p => keyFor(p, '');
  const inP = [], outP = [];
  for(const a of seq){
    for(const p0 of splitParty(a.inParty)){ if(hiddenP.has(p0)) continue; const p = partyKey(p0); if(!actRoles.includes(p) && !inP.includes(p)) inP.push(p); }
    for(const p0 of splitParty(a.outParty)){ if(hiddenP.has(p0)) continue; const p = partyKey(p0); if(!actRoles.includes(p) && !outP.includes(p) && !inP.includes(p)) outP.push(p); }
  }
  const lanes = [...sortKeys(inP).map(r=>({role:r,kind:'in'})), ...actRoles.map(r=>({role:r,kind:'act'})), ...sortKeys(outP).map(r=>({role:r,kind:'out'}))];
  const laneIdx = new Map(lanes.map((l,i)=>[l.role,i]));
  const nodes = [], edges = [];
  const mk = (type, role, p) => { const n = {id:nodes.length, type, lane:laneIdx.get(role), ...p}; nodes.push(n); return n; };
  const pack = (arr, lane) => { let last = arr[arr.length-1]; if(!last || last.has(lane)){ last = new Set(); arr.push(last); } last.add(lane); return arr.length-1; };
  const preRows = [], postRows = [], preEdges = [];
  const offIn = new Map(), offOut = new Map(), nodeOf = new Map(), crossInputs = new Map(); let startNode = null, endNode = null; const endSrc = [];

  for(const a of seq){
    const lane = laneIdx.get(R(a));
    const refs = parseFlow(a.prev);
    if(startOf.has(a.l5no)){ startNode = mk('start', R(a), {code:a.l5no,text:a.startLabel||('Mulai'+(startNote.get(a.l5no)?' — '+startNote.get(a.l5no):'')), nearStart:a.l5no}); preEdges.push([startNode, a, '']); }
    const crossRefs = refs.filter(r=>r.kind==='code' && !inL4.has(r.code));
    if(crossRefs.length > 1){ // v1.7.0: two or more triggers share one start event, in the task's own lane
      // v1.4.1: many triggers from other L4s -> one Message Start Event listing all sources.
      for(const r of crossRefs){ const d=bpaCrossL4Data(r.code,l4.l4no); if(d){ const list=crossInputs.get(a.l5no)||[]; if(!list.some(x=>x.source===d.source))list.push(d); crossInputs.set(a.l5no,list);} }
      const useLane = lane;
      const g = mk('offin', lanes[useLane].role, {target:a.l5no, code:'multi:'+a.l5no, group:crossRefs.map(r=>({code:r.code, note:r.note||'', ext:state.byCode.get(r.code)})), prerow:pack(preRows,useLane)});
      for(const r of crossRefs) offIn.set(r.code+'|'+a.l5no, g);
      preEdges.push([g, a]);
    }
    else for(const r of refs) if(r.kind==='code' && !inL4.has(r.code)){
      // A documented source artifact is data provenance, not a separate process node
      // above the Start Event. Attach it to the receiving task instead.
      const incomingData=bpaCrossL4Data(r.code,l4.l4no);
      if(incomingData){
        // v1.4.0: the documented data stays a side-car Data Object; the trigger itself
        // is still drawn as a Message Start Event below.
        const list=crossInputs.get(a.l5no)||[];
        if(!list.some(x=>x.source===incomingData.source))list.push(incomingData);
        crossInputs.set(a.l5no,list);
      }
      const okey = r.code+'|'+a.l5no; // v1.6.3: one trigger per receiving task (different branches of the source stay distinct)
      let n = offIn.get(okey);
      if(!n){
        // prefer an empty neighbouring activity lane in the current top row, so Mulai and this box sit side by side
        const useLane = lane; // v1.7.0: a trigger sits in the lane of the activity it starts
        n = mk('offin', lanes[useLane].role, {target:a.l5no, code:r.code, note:r.note||'', ext:state.byCode.get(r.code), prerow:pack(preRows,useLane)}); offIn.set(okey,n);
      }
      preEdges.push([n, a]);
    }
  }
  if(!nodes.length && seq.length) preEdges.push([mk('start', R(seq[0]), {text:'Mulai', prerow:pack(preRows, laneIdx.get(R(seq[0])))}), seq[0]]);

  // v1.7.1: a document received from another L4 stays with the step that uses it (dataFrom),
  // even when a new step is inserted in front of it and takes over the trigger.
  for(const a of seq) for(const code of String(a.dataFrom||'').split(/[;,\s]+/).filter(Boolean)){
    if(!seq.some(x=>parseFlow(x.prev).some(f=>f.kind==='code'&&f.code===code))) continue;
    const d=bpaCrossL4Data(code,l4.l4no); if(!d) continue;
    for(const list of crossInputs.values()){ const i=list.findIndex(x=>x.source===d.source); if(i>=0) list.splice(i,1); }
    const L=crossInputs.get(a.l5no)||[]; if(!L.some(x=>x.source===d.source)) L.push(d); crossInputs.set(a.l5no,L);
  }
  let row = preRows.length;
  for(const a of seq){
    if(isGateRow(a)){ // v1.8.0: a gateway directly after a gateway: only a diamond, no task box
      const fl=flowOf.get(a.l5no)||[]; const gn=mk('dec', R(a), {row:row++, text:gateQ(a)||decQuestion(fl), act:a, code:a.l5no, isGate:true});
      gn.flow=fl; gn.crossInputs=[]; nodeOf.set(a.l5no, gn); continue;
    }
    const n = mk('act', R(a), {act:a, code:a.l5no, row:row++, isNew:!!a.isNew, elab:a.isNew ? a.elab : (chg.get(a.l5no)||[]).find(i=>i.status==='proposed'), chg:chg.get(a.l5no)});
    nodeOf.set(a.l5no, n);
    n.crossInputs=crossInputs.get(a.l5no)||[];
    const flow = flowOf.get(a.l5no) || [];
    const labels = uniq(flow.filter(f=>f.label).map(f=>f.label));
    let isDec = labels.length >= 2;
    if(!isDec && !labels.length){
      const tg = flow.filter(f=>f.kind==='code' || f.kind==='end');
      if(tg.length>=2 && tg.length===flow.length && tg.every(f=>f.note)){ isDec = true; tg.forEach(f=>{ f.label = f.note; f.note=''; }); }
    }
    n.flow = flow;
    if(isDec){ const dr = diamondKey(a,'dec'); n.dec = mk('dec', dr || R(a), {row:row++, text:String(a.decisionQuestion||'').replace(/^\s*\[[^\]]*\]\s*/,'') || decQuestion(flow), act:a, moved: dr ? (byType.declane.get(a.l5no)||[])[0] : null}); if(n.dec.moved) drawn.push(n.dec.moved); }
    else if(byType.gate.has(a.l5no)){
      const g = byType.gate.get(a.l5no).find(x=>inL4.has(x.noTarget));
      if(g){ let q = String(g.text||'').trim() || 'Lolos pemeriksaan?'; if(!/\?$/.test(q)) q += '?';
        const gr = diamondKey(a,'gate'); n.gate = mk('dec', gr || R(a), {row:row++, text:q, act:a, elab:g}); drawn.push(g); }
    }
    const docs = (byType.doc.get(a.l5no)||[]).slice();
    if(docs.length) drawn.push(...docs);
    const xo = state.extra[a.l5no];
    const dataKind=bpmnDataKind(a);
    // v1.6.1: "X" only marks that a document/data is produced; its name comes from Output (T).
    const rawMark=String((dataKind==='store'?a.systemX:a.docX)||'').trim();
    const markedText=rawMark && rawMark.toUpperCase()!=='X' ? rawMark : String(a.outputX||'').trim().split(/\n+/).filter(Boolean).join(' · ');
    // v1.6.7: "X" with an empty Output but a filled Input marks the INPUT type (common on the first activity of an L4).
    const inputOnly = dataKind && rawMark.toUpperCase()==='X' && !String(a.outputX||'').trim() && !cleanOut(xo&&xo.out) && String(a.inputX||'').trim();
    if(inputOnly){ n.crossInputs=[...(n.crossInputs||[]), {kind:dataKind, label:String(a.inputX).trim().split(/\s*\n+\s*/).filter(Boolean).join(' · '), source:''}]; }
    else if(dataKind && (markedText || cleanOut(xo&&xo.out))){
      // Existing Dokumen / Sistem-Aplikasi fields determine the BPMN symbol.
      // Keep the existing output note as the label when it was entered in the editor.
      const label=(rawMark&&rawMark.toUpperCase()!=='X'?rawMark:'')||cleanOut(xo&&xo.out)||markedText||(dataKind==='store'?'Penyimpanan data':'Data');
      n.tag={items:[{id:'mark:'+a.l5no,status:'accepted',text:label,isOut:true}],text:label,kind:dataKind};
    } else if(docs.length){
      // Legacy elaboration remains readable as annotation, not a BPMN data object.
      n.tag={items:docs,text:docs.map(d=>d.text).join(' · '),kind:'annotation'};
    }
    { const xd=docsOf(a.l5no); // v1.7.1: extra inputs join the input rail, extra outputs go below the main output
      const ins=xd.filter(d=>d.d!=='out').map(d=>({kind:d.k==='store'?'store':'object',label:d.n,source:'',extra:true}));
      if(ins.length) n.crossInputs=[...(n.crossInputs||[]),...ins];
      n.extraOuts=xd.filter(d=>d.d==='out').map(d=>({kind:d.k==='store'?'store':'object',label:d.n})); }
    if(xo && cleanOut(xo.out) && !dataKind && String(a.docX||'').trim()!=='-'){ // v1.6.4: Output text shows as a document unless marked otherwise
      // Manual editor must continue to create an explicit artifact.
      // The manual 'Dokumen / output' field represents a Data Object, not a process task.
      n.tag={items:[{id:'manual:'+a.l5no,status:'accepted',text:cleanOut(xo.out),isOut:true}],text:cleanOut(xo.out),kind:'object'};
    }
    if(n.tag && n.tag.kind!=='annotation') n.tag.crossL4Targets=bpaCrossL4OutputTargets(a,l4.l4no);
    if(xo && xo.sla) n.sla = xo.sla;
    if(noteNo.has(a.l5no)){ n.noteNo = noteNo.get(a.l5no); n.note = xo.note; }
  }
  const mainEnd = row;

  { // v1.4.0: BPMN start entries. Several start events into one task meet in an explicit XOR merge.
    const entries=new Map();
    for(const [s,a,lab] of preEdges){ if(!entries.has(a.l5no))entries.set(a.l5no,{a,list:[]}); entries.get(a.l5no).list.push([s,lab]); }
    for(const {a,list} of entries.values()){
      const t=nodeOf.get(a.l5no);
      if(list.length>1 && showMergeGw()){
        const m=mk('merge', R(a), {code:a.l5no, nearStart:a.l5no});
        for(const [s,lab] of list) edges.push({s, t:m, label:lab||''});
        edges.push({s:m, t, label:''});
      } else for(const [s,lab] of list) edges.push({s, t, label:lab||''});
    }
  }
  for(const a of seq){
    const n = nodeOf.get(a.l5no); let src = n;
    if(n.dec){ edges.push({s:n, t:n.dec, label:''}); src = n.dec; }
    else if(n.gate){ edges.push({s:n, t:n.gate, label:''}); src = n.gate; }
    let firstYes = !!n.gate;
    // v1.4.1: a branch that both ends here and continues to another L4 is one Message End Event.
    const crossLab = new Set(n.flow.filter(f=>f.kind==='code' && !inL4.has(f.code)).map(f=>f.label||''));
    let endHere = null;
    for(const f of n.flow){
      let t = null;
      if(f.kind==='end' && crossLab.has(f.label||'')) continue;
      if(f.kind==='end' && oneEnd()){ // v1.9.2: one shared Selesai at the bottom, in the lane of the last path that ends
        if(!endNode) endNode = mk('end', lanes[n.lane].role, {code:'', shared:true, text:(seq.find(x=>x.endLabel&&parseFlow(x.next).some(y=>y.kind==='end'))||{}).endLabel||'Selesai'});
        if(!endSrc.includes(src)) endSrc.push(src); t = endNode; }
      else if(f.kind==='end'){ if(!endHere) endHere = mk('end', lanes[n.lane].role, {code:a.l5no,text:a.endLabel||'Selesai', nearEnd:src}); /* v1.7.0: lane where the flow ends */ t = endHere; }
      else if(f.kind==='code'){
        if(inL4.has(f.code)) t = nodeOf.get(f.code);
        else { const ok2=f.code+'|'+src.id; t = offOut.get(ok2); if(!t){ t = mk('offout', lanes[n.lane].role, {src:a.l5no, code:f.code, ext:state.byCode.get(f.code), nearEnd:src}); offOut.set(ok2, t); } } /* v1.7.0: right after, and in the lane of, the step where the process ends */
      } else continue;
      let label = f.label || '';
      if(f.note) label = label ? label+' ('+f.note+')' : f.note;
      if(firstYes){ label = label ? 'Ya · '+label : 'Ya'; firstYes = false; }
      edges.push({s:src, t, label});
    }
    if(n.gate){ const g = n.gate.elab; edges.push({s:n.gate, t:nodeOf.get(g.noTarget), label:g.noLabel||'Tidak', elab:g}); }
  }
  if(endNode){ endNode.postrow = postRows.length; postRows.push(new Set()); const last = endSrc.reduce((m,x)=>x.row>m.row?x:m, endSrc[0]); endNode.lane = last.lane; }
  for(const a of seq){
    const n = nodeOf.get(a.l5no);
    const seen = new Set();
    for(const p of splitParty(a.inParty).filter(x=>!hiddenP.has(x)).map(partyKey)) if(laneIdx.has(p) && lanes[laneIdx.get(p)].kind!=='act' && !seen.has(p)){
      seen.add(p); const d = mk('docin', p, {row:n.row, text:symbolDocLabel(a,'docin',p)||'Input dari '+laneShort(roleInfo(p).name), code:a.l5no, party:p}); edges.push({s:d, t:n, label:'', side:true});
    }
    for(const p of splitParty(a.outParty).filter(x=>!hiddenP.has(x)).map(partyKey)) if(laneIdx.has(p) && lanes[laneIdx.get(p)].kind!=='act' && !seen.has(p)){
      seen.add(p); const d = mk('docout', p, {row:n.row, text:symbolDocLabel(a,'docout',p)||'Output ke '+laneShort(roleInfo(p).name), code:a.l5no, party:p}); edges.push({s:n, t:d, label:'', side:true});
    }
  }
  // elaboration side nodes (HOLD, input data) in the same row, nearest free lane
  const occupied = (r, l) => nodes.some(x=>x.row===r && x.lane===l);
  const freeLane = (r, l) => { for(let d=1; d<lanes.length; d++) for(const c of [l+d, l-d]) if(c>=0 && c<lanes.length && !occupied(r,c)) return c; return null; };
  for(const a of seq){
    const n = nodeOf.get(a.l5no);
    for(const it of (byType.hold.get(a.l5no)||[])){
      const src = n.gate || n; const l = freeLane(src.row, src.lane); if(l==null) continue;
      const h = mk('hold', lanes[l].role, {row:src.row, text:it.text||'HOLD', elab:it});
      edges.push({s:src, t:h, label:it.label||'Hold', side:true, elab:it}); drawn.push(it);
    }
    for(const it of (byType.datain.get(a.l5no)||[])){
      const l = freeLane(n.row, n.lane); if(l==null) continue;
      const d = mk('datain', lanes[l].role, {row:n.row, text:it.text||'Input data', elab:it});
      edges.push({s:d, t:n, label:'', side:true, elab:it}); drawn.push(it);
    }
  }
  for(const n of nodes){ if(n.prerow!=null) n.row = n.prerow; if(n.postrow!=null) n.row = mainEnd + n.postrow; }
  // Insert local terminator rows without moving documents away from their activities.
  const before=new Map(), after=new Map();
  for(const n of nodes){if(n.nearStart){const r=nodeOf.get(n.nearStart).row;before.set(r,[...(before.get(r)||[]),n]);}if(n.nearEnd){const r=n.nearEnd.row;after.set(r,[...(after.get(r)||[]),n]);}}
  let offset=0;for(let r=0;r<mainEnd+postRows.length;r++){for(const n of before.get(r)||[])n.row=r+offset++;const same=nodes.filter(n=>!n.nearStart&&!n.nearEnd&&n.row===r&&!n._layoutPlaced);for(const n of same){n.row=r+offset;n._layoutPlaced=true;}for(const n of after.get(r)||[])n.row=r+offset+1,offset++;}
  let nRows = mainEnd + postRows.length + offset;
  { // v1.4.0: start entries of one XOR merge share a row (Mulai beside the Message Start Event).
    const entryTargets=[...new Set(edges.filter(e=>e.s.type==='start'||e.s.type==='offin').map(e=>e.t))];
    for(const m of entryTargets){
      const srcs=edges.filter(e=>e.t===m).map(e=>e.s), st=srcs.find(n=>n.type==='start');
      if(!st) continue;
      let sideUsed = false;
      for(const o of srcs.filter(n=>n.type==='offin')){
        if(o.lane===st.lane){ if(sideUsed) continue; o._dx=140; o._paired=true; st._paired=true; sideUsed=true; } // v1.7.0: always beside Mulai
        else if(nodes.some(n=>n!==o && n.row===st.row && n.lane===o.lane)) continue;
        o.row=st.row;
      }
    }
    const used=[...new Set(nodes.map(n=>n.row))].sort((a,b)=>a-b), rmap=new Map(used.map((r,i)=>[r,i]));
    for(const n of nodes) n.row=rmap.get(n.row);
    nRows=used.length;
  }

  /* ---- sizes ---- */
  const bw = 250; // fixed BPMN task width; artifact rail is not part of task width
  for(const n of nodes){
    if(n.type==='act'){ n.lines = wrap(n.act.l5, 12.5, true, bw-20); n.w = bw; n.h = Math.max(62, n.lines.length*16+30); }
    else if(n.type==='dec'){ n.w = 46; n.h = 46; n.lines = clampWrap(n.text, 10.5, true, 116, 6); n._labelExtra = Math.max(0, n.lines.length*13+10-n.h)+6; }
    else if(n.type==='merge'){ n.w = 40; n.h = 40; n.lines = []; }
    else if(n.type==='start' || n.type==='end'){ n.w = 32; n.h = 32; n.lines = clampWrap(n.text,11,true,n._paired?100:160,3); n._labelExtra = Math.max(0, n.lines.length*13+6-n.h); }
    else if(n.type==='offin' || n.type==='offout'){
      n.w = bw;
      const head = (n.type==='offin'?'Dari ':'Lanjut ke ')+n.code;
      const l4ref = !n.ext && !n.group && state.l4s && state.l4s.get(n.code);
      n.crossData = n.type==='offin'?bpaCrossL4Data(n.code,l4.l4no):null;
      const body = n.ext ? wrap(n.ext.l5, 9.5, false, bw-26).slice(0,2) : l4ref ? wrap(l4ref.l4||'Referensi prosedur L4',9.5,false,bw-26).slice(0,2) : ['Referensi proses '+n.code];
      const sub = []; // Keep cross-L4 references compact; details remain in activity data.
      if(n.crossData)sub.push(...wrap((n.crossData.kind==='store'?'Data Store: ':'Data Object: ')+n.crossData.label,8.5,false,bw-22).slice(0,2));
      n.parts = {head, body, sub};
      // Cross-L4 predecessor with explicit source data is a data artifact, not a process-reference box.
      // Keep the process reference as metadata; do not infer a data artifact for an unmarked source.
      // v1.4.0: cross-L4 references are Message Start / Message End Events (BPMN 2.0).
      // Documented source data is drawn separately as a side-car Data Object.
      n.crossData = null;
      n.w = 32; n.h = 32;
      n.evLines = [{t:(n.type==='offin'?'Dari ':'Ke ')+n.code, b:true}, ...(n.note?[{t:'('+n.note+')', b:false}]:[]), ...clampWrap(n.ext ? n.ext.l5 : body.join(' '), 10, false, n._paired?118:165, n.note?2:3).map(t=>({t,b:false}))];
      if(n.group) n.evLines = [{t:'Dipicu dari L4 lain', b:true}, ...n.group.flatMap(g=>clampWrap(g.code+(g.note?' · '+g.note:(g.ext?' · '+g.ext.l5:'')),10,false,n._paired?118:165,2).map(t=>({t,b:false})))];
      n._labelExtra = Math.max(0, n.evLines.length*13+6-n.h);
    }
    else if(n.type==='hold'){ n.w = bw-14; n.lines = wrap(n.text, 10, true, n.w-30); n.h = n.lines.length*13+20; }
    else if(n.type==='datain'){ n.w = bw-20; n.lines = wrap(n.text, 9.5, false, n.w-34); n.h = n.lines.length*12+22; }
    else { n.w = 34; n.h = 40; n.lines = clampWrap(n.text, 10, false, 160, 3); n._labelExtra = n.lines.length*12+8; } // docin/docout = Data Object
    if(n.tag){ n.tag.w = Math.min(G.laneW-30,Math.max(95,Math.round(n.w/1.45))); n.tag.lines = wrap(n.tag.text, 8.5, false, n.tag.w-14); n.tag.h = Math.max(56,n.tag.lines.length*10.5+43); n.tag.extra = Math.max(0, n.tag.h-10); if(n.tag.crossL4Targets?.length)n.tag.extra+=18; }
  }

  /* ---- rows & coordinates ---- */
  const rowH = new Array(nRows).fill(0);
  // v1.7.0: rows are measured above and below the centre line separately, because the artifact rail only grows downward.
  const rowUp=new Array(nRows).fill(17), rowDown=new Array(nRows).fill(17);
  for(const n of nodes){ const half=(n.h+(n._labelExtra||0))/2; rowUp[n.row]=Math.max(rowUp[n.row],half); rowDown[n.row]=Math.max(rowDown[n.row],half,(n.tag||n.sla||n.crossInputs?.length||n.extraOuts?.length)?railExtent(n)+(n._labelExtra||0)/2:0); }
  for(let i=0;i<nRows;i++) rowH[i] = Math.max(rowH[i], rowUp[i]+rowDown[i]) + 2*G.band;
  const lanesX = G.m;
  // lane headers: position names can wrap to several lines
  const laneHead = lanes.map(ln=>{
    const info = roleInfo(ln.role);
    let lines = wrap(info.pos ? laneShort(info.name) : info.name, info.pos?11:12, true, G.laneW-18);
    if(lines.length > 4){ lines = lines.slice(0,4); lines[3] = lines[3].replace(/\s*\S*$/,'')+' …'; }
    const roles = [...(laneRoles.get(ln.role)||[])].map(r=>roleInfo(r));
    // v1.5.0: lane = jabatan (or unit); subtitle = organisasi (or jabatan). Process roles are not shown on the canvas.
    const unitLane = String(ln.role).startsWith('unit::');
    const vals = uniq(seq.filter(a=>R(a)===ln.role).map(a=>unitLane ? sepPos(a.posX||'').join(' / ') : String(a.unitX||'').trim()).filter(Boolean));
    if(!vals.length){ for(const r of (state.roleEdit.added[l4.l4no]||[])) if(addedLaneKey(l4.l4no,r)===ln.role){ const u=unitLane?((state.posmap[l4.l4no]||{})[r]||''):((state.unitmap[l4.l4no]||{})[r]||''); if(u) vals.push(u); } }
    let orgSub = vals.join(' · '); if(orgSub.length>95) orgSub = orgSub.slice(0,92)+'…';
    const sub = ln.kind==='in' ? '(Pihak pemberi input)' : ln.kind==='out' ? '(Pihak penerima output)' : info.pos ? orgSub : '';
    return {lines, sub, size: info.pos?11:12};
  });
  const LH = Math.max(G.laneHeadH, Math.max(0, ...laneHead.map(h=>h.lines.length))*14 + 30);
  const y0 = G.m + G.headerH + G.poolH + LH + 8;
  const rowTop = []; let yy = y0;
  for(let i=0;i<nRows;i++){ rowTop.push(yy); yy += rowH[i]; }
  const rowsEnd = yy;
  const laneX = i => lanesX + i*G.laneW;
  for(const n of nodes){ n.cx = laneX(n.lane) + 150 + (n._dx||0); n.cy = rowTop[n.row] + G.band + rowUp[n.row]; }
  const boundary = b => b < nRows ? rowTop[b] : rowsEnd;

  /* ---- routing ---- */
  // Edges heading to the same node share one horizontal "bus" and one gutter line, so they merge instead of running side by side.
  const gutUse = new Map(), gutMemo = new Map(), gapUse = new Map(), gapMemo = new Map();
  const gutterX = (lane, side, key) => {
    const mk = lane+side+'|'+key;
    if(key && gutMemo.has(mk)) return gutMemo.get(mk);
    const k = lane+side; const c = gutUse.get(k)||0; gutUse.set(k, c+1);
    const x = side==='left' ? laneX(lane)+12+(c%4)*10 : laneX(lane)+G.laneW-12-(c%4)*10;
    if(key) gutMemo.set(mk, x); return x;
  };
  const gapY = (b, key) => {
    const mk = b+'|'+key; if(key && gapMemo.has(mk)) return gapMemo.get(mk);
    const c = gapUse.get(b)||0; gapUse.set(b, c+1); const offs = [0,9,-9,18,-18,26,-26];
    const y = boundary(b) + offs[c%offs.length]; if(key) gapMemo.set(mk, y); return y;
  };
  const clear = (lane, r1, r2) => !nodes.some(n=>n.lane===lane && n.row>Math.min(r1,r2) && n.row<Math.max(r1,r2));
  const rowFree = (row, l1, l2) => !nodes.some(n=>n.row===row && n.lane>Math.min(l1,l2) && n.lane<=Math.max(l1,l2) && n.lane!==l1);
  const top = n => [n.cx, n.cy - n.h/2];

  // exit sides: straight-down edge takes the bottom; the others leave on the side facing their target
  const railN = n => !!(n && n.type==='act' && (n.tag || n.sla || n.crossInputs?.length||n.extraOuts?.length));
  const outBy = new Map(), usedSide = new Map();
  for(const e of edges){
    if(e.side){
      for(const [a,b] of [[e.s,e.t],[e.t,e.s]]) if(a.type==='dec'){ if(!usedSide.has(a.id)) usedSide.set(a.id,new Set()); usedSide.get(a.id).add(b.cx>a.cx?'right':'left'); }
      continue;
    }
    if(!outBy.has(e.s.id)) outBy.set(e.s.id, []); outBy.get(e.s.id).push(e);
  }
  for(const [sid,list] of outBy){
    const s0 = list[0].s;
    const fw = list.filter(e=>e.t.row > e.s.row).sort((a,b)=>a.t.row-b.t.row);
    const bk = list.filter(e=>e.t.row <= e.s.row);
    const used = new Set(usedSide.get(sid) || []);
    const take = pref => { for(const p of pref) if(!used.has(p)){ used.add(p); return p; } return pref[0]; };
    const prim = fw.find(e=>e.t.lane===s0.lane && clear(s0.lane, s0.row, e.t.row)) || fw.find(e=>e.t.lane===s0.lane) || fw[0];
    if(prim){ prim.exit = 'bottom'; prim.back = false; }
    // in its own column a vertical run should avoid the side where input/output arrows leave the boxes it passes
    const sideBusy = (side, r1, r2) => edges.some(x=>{ if(!x.side) return false; const n = x.s.lane===s0.lane ? x.s : x.t.lane===s0.lane ? x.t : null; if(!n) return false;
      const o = n===x.s ? x.t : x.s; return n.row>Math.min(r1,r2) && n.row<Math.max(r1,r2) && ((o.cx>n.cx) === (side==='right')); });
    const prefFor = e => { if(['start','offin'].includes(s0.type)) return ['left','right']; // v1.6.3: an event's label sits on its right
      if(railN(s0) && e.t.lane<=s0.lane) return ['left','right']; // v1.4.0: keep the artifact rail clear
      if(e.t.lane!==s0.lane) return e.t.lane < s0.lane ? ['left','right'] : ['right','left'];
      return sideBusy('right', s0.row, e.t.row) && !sideBusy('left', s0.row, e.t.row) ? ['left','right'] : ['right','left']; };
    for(const e of fw) if(e!==prim){ e.back = false; e.exit = take(prefFor(e)); }
    for(const e of bk){ e.back = true; e.exit = take(prefFor(e)); }
  }
  // Independent return tracks and entry ports. Allocate once per edge, not per target.
  const returnGroups=new Map(), returnPorts=new Map();
  for(const e of edges.filter(e=>e.back&&!e.side)){
    const side=e.s.lane!==e.t.lane && e.exit===(e.t.lane<e.s.lane?'left':'right') && rowFree(e.s.row,e.s.lane,e.t.lane) && !nodes.some(n=>n.row===e.s.row&&n.lane===e.t.lane) ? (e.t.lane<e.s.lane?'right':'left') : e.exit;
    e.returnSide=side; const k=e.t.lane+'|'+side, p=e.t.id+'|'+side;
    if(!returnGroups.has(k))returnGroups.set(k,[]);returnGroups.get(k).push(e);
    if(!returnPorts.has(p))returnPorts.set(p,[]);returnPorts.get(p).push(e);
  }
  for(const group of returnGroups.values()){
    group.sort((a,b)=>b.s.row-a.s.row||a.t.row-b.t.row||a.s.id-b.s.id);
    group.forEach((e,i)=>{const d=6+i*Math.min(10,40/Math.max(1,group.length-1));e.returnX=laneX(e.t.lane)+(e.returnSide==='left'?d:G.laneW-d);});
  }
  for(const group of returnPorts.values()){
    group.sort((a,b)=>Math.abs(b.returnX-b.t.cx)-Math.abs(a.returnX-a.t.cx));
    group.forEach((e,i)=>{e.returnY=e.t.cy+(group.length===1?(e.t.type==='act'?-8:0):(-1+2*i/(group.length-1))*Math.min(18,e.t.h/2-10));});
    // v1.6.2: returns into a task with an artifact rail come in along the task's upper edge, above the icons.
    group.forEach((e,i)=>{ if(railN(e.t) && e.returnSide==='right') e.returnY=Math.min(e.t.cy-8, e.t.cy-e.t.h/2+9+i*7); });
  }
  for(const e of edges){
    const s = e.s, t = e.t; let pts;
    if(e.side){
      const y = s.type==='dec' ? s.cy : t.type==='dec' ? t.cy : s.type==='act' ? s.cy+8 : t.cy+8;
      pts = s.cx < t.cx ? [[s.cx+s.w/2, y],[t.cx-t.w/2, y]] : [[s.cx-s.w/2, y],[t.cx+t.w/2, y]];
      e.exit = s.cx < t.cx ? 'right' : 'left';
      // v1.4.0: a task's right-hand artifact rail (data object, SLA) must not be crossed by a data association.
      const rail = n => n.type==='act' && (n.tag || n.sla || n.crossInputs?.length||n.extraOuts?.length);
      const railSide = (s.cx < t.cx && rail(s)) ? s : (s.cx > t.cx && rail(t)) ? t : null;
      if(railSide){
        const gy = rowTop[railSide.row] + rowH[railSide.row] - 8;
        const [a,b] = pts, dir = Math.sign(b[0]-a[0]);
        pts = [a,[a[0]+dir*10,a[1]],[a[0]+dir*10,gy],[b[0]-dir*10,gy],[b[0]-dir*10,b[1]],b];
      }
    } else if(!e.back){
      const p1 = top(t);
      if(e.exit==='bottom'){
        const p0 = [s.cx, s.cy+s.h/2];
        if(s.lane===t.lane && clear(s.lane, s.row, t.row)) pts = Math.abs(s.cx-t.cx)<1 ? [p0, p1] : (()=>{ const g = gapY(t.row, 'in'+t.id); return [p0,[s.cx,g],[t.cx,g],p1]; })();
        else if(s.lane!==t.lane && clear(t.lane, s.row, t.row)){ const g = gapY(s.row+1, (s.type==='offin'||s.type==='start') ? 'in'+t.id : 'out'+s.id); pts = [p0,[s.cx,g],[t.cx,g],p1]; }
        else if(s.lane!==t.lane && clear(s.lane, s.row, t.row)){ const g = gapY(t.row, 'in'+t.id); pts = [p0,[s.cx,g],[t.cx,g],p1]; }
        else {
          const g1 = gapY(s.row+1, 'out'+s.id), g2 = gapY(t.row, 'in'+t.id);
          const xg = gutterX(t.lane, t.lane<s.lane ? 'right' : 'left', 'to'+t.id);
          pts = [p0,[s.cx,g1],[xg,g1],[xg,g2],[t.cx,g2],p1];
        }
      } else {
        const dir = e.exit==='right' ? 1 : -1, sx = s.cx + dir*s.w/2;
        const straight = t.lane!==s.lane && (t.lane-s.lane)*dir > 0 && rowFree(s.row, s.lane, t.lane) && clear(t.lane, s.row, t.row);
        if(straight) pts = [[sx,s.cy],[t.cx,s.cy],p1];
        else { const xg = gutterX(s.lane, e.exit, 'from'+s.id+e.exit), g2 = gapY(t.row, 'in'+t.id); pts = [[sx,s.cy],[xg,s.cy],[xg,g2],[t.cx,g2],p1]; }
      }
    } else {
      // back edge to another column: leave toward it, climb in the target column's near gutter, enter the near side
      if(s.lane!==t.lane){
        const toLeft = t.lane < s.lane, near = toLeft ? 'right' : 'left';
        if(e.exit===(toLeft?'left':'right') && e.returnSide===near && rowFree(s.row, s.lane, t.lane) && !nodes.some(n=>n.row===s.row && n.lane===t.lane)){
          const sx = s.cx + (toLeft?-1:1)*s.w/2, sy = s.type==='act' ? s.cy+6 : s.cy;
          const xgT = e.returnX, ty = e.returnY, tx = t.cx + (toLeft?1:-1)*t.w/2;
          e.pts = simplifyPts([[sx,sy],[xgT,sy],[xgT,ty],[tx,ty]]); continue;
        }
      }
      const side = e.exit==='left' ? 'left' : 'right', sgn = side==='right'?1:-1;
      const sy = s.type==='act' ? s.cy+6 : s.cy;
      const ts = e.returnSide==='left' ? -1 : 1;
      const sx = s.cx + (s.lane===t.lane ? ts : sgn)*s.w/2;
      const ty = e.returnY, tx = t.cx + ts*t.w/2;
      if(s.lane===t.lane){ const xg = e.returnX; pts = [[sx,sy],[xg,sy],[xg,ty],[tx,ty]]; }
      else { const xg = gutterX(s.lane, side, 'from'+s.id+side), g = gapY(t.row, 'bkin'+s.id+'-'+t.id), xgT = e.returnX; pts = [[sx,sy],[xg,sy],[xg,g],[xgT,g],[xgT,ty],[tx,ty]]; }
    }
    e.pts = simplifyPts(pts);
  }
  // v1.7.0: a right-hand return into a task with an SLA annotation drops onto the task's top edge,
  // so it never runs through the annotation that sits at the top of the artifact rail.
  { const cnt=new Map();
    for(const e of edges){ const t=e.t; if(!e.back||e.returnSide!=='right'||!t||t.type!=='act'||!t.sla||!e.pts||e.pts.length<2) continue;
      const p=e.pts, a=p[p.length-2], b=p[p.length-1]; if(Math.abs(a[1]-b[1])>1) continue;
      const i=cnt.get(t.id)||0; cnt.set(t.id,i+1);
      const yTop=t.cy-t.h/2-20-i*7, tx2=t.cx+30+i*10;
      e.pts=simplifyPts([...p.slice(0,-2),[a[0],yTop],[tx2,yTop],[tx2,t.cy-t.h/2]]);
    } }


  /* v1.3.8.1: conservative obstacle/overlap repair, before SVG rendering. */
  const routeSegments = p => p.slice(1).map((b,i)=>[p[i],b]);
  const railBox = n => (n.type==='act' && (n.tag||n.sla||n.crossInputs?.length||n.extraOuts?.length)) ? {l:laneX(n.lane)+G.laneW-84-50, r:laneX(n.lane)+G.laneW-84+50, a:n.sla?n.cy-n.h/2:n.cy-6, b:n.cy+railExtent(n)} : null;
  const routeHits = (e,p) => routeSegments(p).reduce((sum,[[x,y],[u,v]])=>sum+nodes.filter(n=>{
    if(n===e.s||n===e.t)return false;
    const rb=railBox(n);
    if(rb && (x===u ? x>rb.l&&x<rb.r&&Math.max(y,v)>rb.a&&Math.min(y,v)<rb.b : y===v && y>rb.a&&y<rb.b&&Math.max(x,u)>rb.l&&Math.min(x,u)<rb.r)) return true;
    const l=n.cx-n.w/2-2,r=n.cx+n.w/2+2,a=n.cy-n.h/2-2,b=n.cy+n.h/2+2;
    return x===u ? x>l&&x<r&&Math.max(y,v)>a&&Math.min(y,v)<b :
      y===v && y>a&&y<b&&Math.max(x,u)>l&&Math.min(x,u)<r;
  }).length,0);
  const routeOverlap = (e,p) => {
    let total=0;
    for(const f of edges){
      if(f===e||f.s===e.s||f.t===e.t||!f.pts)continue;
      for(const [[x,y],[u,v]] of routeSegments(p))
      for(const [[a,b],[c,d]] of routeSegments(f.pts)){
        // v1.5.0: near-parallel runs (< 9 px apart) count as overlap too, so the repair pass spreads them.
        if(x===u&&a===c&&Math.abs(x-a)<9)total+=Math.max(0,Math.min(Math.max(y,v),Math.max(b,d))-Math.max(Math.min(y,v),Math.min(b,d))-(x===a?0:20));
        if(y===v&&b===d&&Math.abs(y-b)<9)total+=Math.max(0,Math.min(Math.max(x,u),Math.max(a,c))-Math.max(Math.min(x,u),Math.min(a,c))-(y===b?0:20));
      }
    }
    return total;
  };
  const routeCost=(e,p)=>routeHits(e,p)*1000000+routeOverlap(e,p)*100+
    routeSegments(p).reduce((a,[[x,y],[u,v]])=>a+Math.abs(x-u)+Math.abs(y-v),0)+p.length*8;
  for(let pass=0;pass<2;pass++)for(const e of edges){
    const orig=e.pts;
    if(!routeHits(e,orig)&&!routeOverlap(e,orig))continue;
    let best=orig, score=routeCost(e,orig);
    const offer=p=>{
      p=simplifyPts(p);
      if(p.length<2||p.some(([x,y])=>x<lanesX+2||x>laneX(lanes.length)-2||y<y0-4||y>rowsEnd+4))return;
      if(routeSegments(p).some(([[x,y],[u,v]])=>x!==u&&y!==v))return;
      // Preserve original endpoint direction; never reverse into a symbol.
      const dir=(a,b)=>[Math.sign(b[0]-a[0]),Math.sign(b[1]-a[1])].join(',');
      if(dir(p[0],p[1])!==dir(orig[0],orig[1])||dir(p[p.length-2],p[p.length-1])!==dir(orig[orig.length-2],orig[orig.length-1]))return;
      const c=routeCost(e,p);if(c<score){best=p;score=c;}
    };
    // Shift interior tracks while retaining ports and endpoint directions.
    for(let i=1;i<orig.length-2;i++){
      const a=orig[i],b=orig[i+1],axis=a[0]===b[0]?0:1;
      for(const delta of [6,-6,12,-12,18,-18,24,-24,30,-30]){
        const p=orig.map(q=>q.slice());p[i][axis]+=delta;p[i+1][axis]+=delta;offer(p);
      }
    }
    // Side documents: detour through the row gap with short outward port stubs.
    if(e.side&&orig.length===2&&orig[0][1]===orig[1][1]){
      const [a,b]=orig,dir=Math.sign(b[0]-a[0]);
      for(const stub of [8,12,16,20])for(const gy of [rowTop[e.s.row]+5,rowTop[e.s.row]+rowH[e.s.row]-5]){
        offer([a,[a[0]+dir*stub,a[1]],[a[0]+dir*stub,gy],[b[0]-dir*stub,gy],[b[0]-dir*stub,b[1]],b]);
      }
    }
    e.pts=best;
  }

  /* ---- keterangan panel ---- */
  const panel = buildPanel(l4, seq.filter(a=>!isGateRow(a)).map(a=>({...a, _role:R(a), _base:baseRole(a), _note:noteNo.get(a.l5no)})), nodes, edges, offIn, offOut, drawn, lanes.filter(ln=>roleInfo(ln.role).pos).map(ln=>({pos:roleInfo(ln.role).name, roles:[...(laneRoles.get(ln.role)||[])]})));
  const hasProp = drawn.some(i=>i.status==='proposed'), hasAcc = drawn.some(i=>i.status!=='proposed' && ['gate','hold','doc','datain'].includes(i.type));
  const px = laneX(lanes.length) + G.panelGap, pw = G.panelW;
  const pLines = layoutPanel(panel, pw-28);
  const panelEnd = y0 + 10 + pLines.h;
  const bodyEnd = Math.max(rowsEnd, panelEnd) + 14;
  const W = px + pw + G.m;

  /* ---- legend ---- */
  const legendItems = [['bstart','Start Event'],['bend','End Event'],['act','Task (aktivitas L5)']];
  const gatewayCodes = [...new Set(nodes.filter(n=>n.type==='dec').map(n=>bpmnGatewayType(n.act.l5no)))];
  for(const kind of gatewayCodes) legendItems.push([kind==='AND'?'gand':kind==='OR'?'gor':'gxor',kind==='AND'?'Parallel Gateway (AND)':kind==='OR'?'Inclusive Gateway (OR)':'Exclusive Gateway (XOR)']);
  legendItems.push(['arrow','Sequence Flow']);
  if(nodes.some(n=>(n.tag&&n.tag.kind==='object')||n.extraOuts?.some(x=>x.kind==='object')||n.crossInputs?.some(x=>x.kind==='object'&&!x.source))) legendItems.push(['tag',legendLabel('object')]);
  if(nodes.some(n=>(n.tag&&n.tag.kind==='store')||n.extraOuts?.some(x=>x.kind==='store')||n.crossInputs?.some(x=>x.kind==='store'&&!x.source))) legendItems.push(['store',legendLabel('store')]);
  if(nodes.some(n=>n.crossInputs?.some(x=>x.kind==='object'&&x.source)||n.type==='offin'&&n.crossData?.kind==='object'))legendItems.push(['tag',legendLabel('object')+' — input dari L4 lain']);
  if(nodes.some(n=>n.crossInputs?.some(x=>x.kind==='store'&&x.source)||n.type==='offin'&&n.crossData?.kind==='store'))legendItems.push(['store',legendLabel('store')+' — input dari L4 lain']);
  if(nodes.some(n=>n.type==='merge') && !gatewayCodes.includes('XOR')) legendItems.splice(3,0,['gxor','Exclusive Gateway (XOR)']);
  if(nodes.some(n=>n.type==='offin')) legendItems[0]=['bstart', nodes.some(n=>n.type==='start')?'Start Event (Mulai / dari L4 lain)':'Start Event (dari L4 lain)'];
  if(nodes.some(n=>n.type==='offout')) legendItems[1]=['bend', nodes.some(n=>n.type==='end')?'End Event (Selesai / ke L4 lain)':'End Event (ke L4 lain)'];
  if(nodes.some(n=>n.type==='docin'||n.type==='docout') && !legendItems.some(i=>i[0]==='tag')) legendItems.push(['tag',legendLabel('object')]);
  if(legendItems.some(i=>i[0]==='tag'||i[0]==='store')) legendItems.push(['assoc','Data Association']);

  const has = t => drawn.some(i=>i.type===t);
  if(has('hold')) legendItems.push(['hold','HOLD (anotasi non-BPMN)']);
  const anyOut = seq.some(a=>state.extra[a.l5no] && state.extra[a.l5no].out), anySla = seq.some(a=>state.extra[a.l5no] && state.extra[a.l5no].sla);
  if(has('doc') || anyOut) legendItems.push(['note','Catatan / output tambahan (non-BPMN)']);
  if(seq.some(a=>noteNo.has(a.l5no))) legendItems.push(['note','Penanda catatan (non-BPMN)']);
  if(anySla) legendItems.push(['sla','Text Annotation (SLA)']);
  if(has('datain')) legendItems.push(['datain','Input data eksternal (anotasi non-BPMN)']);
  if(has('newact')) legendItems.push(['new','Aktivitas baru (penanda perubahan)']);
  if(has('reorder') || has('role') || has('rename')) legendItems.push(['chg','Perubahan aktivitas (penanda perubahan)']);
  if(hasProp) legendItems.push(['prop','Usulan (penanda perubahan)']);
  let lx = G.m, ly = bodyEnd + 34; const legPos = [];
  for(const [k,txt] of legendItems){ const w = 44 + tw(txt, 10, false) + 26; if(lx + w > W - G.m){ lx = G.m; ly += 26; } legPos.push({k,txt,x:lx,y:ly}); lx += w; }
  const footY = ly + 28;
  let footer = 'Simbol data mengikuti kolom U: Dokumen / W: Sistem-Aplikasi pada sheet SOP. · Dibuat otomatis dari '+state.fileName+' · swimlane = Jabatan / Unit Kerja sesuai pilihan. Diagram mengikuti notasi BPMN 2.0: pool & lane, start/end event (termasuk keterkaitan ke/dari L4 lain), task, gateway, sequence flow, data object, dan data association. Panel Keterangan berada di luar pool. SLA ditulis sebagai Text Annotation. HOLD, penanda catatan, dan penanda perubahan adalah anotasi tambahan non-BPMN.';
  if(hasAcc) footer += ' Gateway pemeriksaan, HOLD, dokumen/bukti, dan input data adalah elaborasi dari Activity Description yang telah direviu.';
  if(drawn.some(i=>['newact','reorder','role','rename'].includes(i.type) && i.status!=='proposed')) footer += ' Kotak berlabel BARU, DIPINDAH, PIC DIUBAH, atau NAMA DIUBAH adalah perubahan terhadap workbook (lihat bagian E).';
  if(hasProp) footer += ' Elemen ungu bergaris putus-putus adalah usulan yang belum direviu.';
  const footLines = wrap(footer, 9, false, W - 2*G.m);
  const H = footY + footLines.length*12 + G.m;

  /* ================= SVG ================= */
  const o = [];
  o.push('<svg xmlns="http://www.w3.org/2000/svg" width="'+W+'" height="'+H+'" viewBox="0 0 '+W+' '+H+'" font-family="'+FONT+'">');
  o.push('<defs><marker id="ar" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="8" markerHeight="8" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="'+C.edge+'"/></marker>'+
         '<marker id="arR" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="8" markerHeight="8" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="'+C.back+'"/></marker>'+
         '<marker id="arP" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="8" markerHeight="8" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="'+C.prop+'"/></marker>'+
         '<marker id="arA" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="8" markerHeight="8" orient="auto-start-reverse"><path d="M1,1 L9,5 L1,9" fill="none" stroke="'+C.docS+'" stroke-width="1.4"/></marker></defs>');
  o.push('<rect width="'+W+'" height="'+H+'" fill="#fff"/>');
  // header
  o.push('<rect x="'+G.m+'" y="'+G.m+'" width="'+(W-2*G.m)+'" height="'+G.headerH+'" fill="'+C.navy+'"/>');
  o.push(txt(G.m+14, G.m+25, l4.l4no+' — '+l4.l4, 17, true, '#fff'));
  const sub = 'SOP '+l4.l3no+' '+l4.l3+' · '+(l4.l2no?l4.l2no+' ':'')+l4.l2+' · '+l4.l1+' · Sumber: '+state.fileName;
  o.push(txt(G.m+14, G.m+45, sub, 10, false, opts.mono?'#eeeeee':'#f1d6de'));
  if(statusLabel){ o.push(txt(W-G.m-14, G.m+25, statusLabel, 12, true, C.amber, 'end')); }
  { const nNew = seq.filter(a=>a.isNew).length; o.push(txt(W-G.m-14, G.m+45, acts.filter(a=>!isGateRow(a)).length+' aktivitas'+(nNew?' + '+nNew+' proses baru':'')+' · dibuat otomatis dari workbook BPA', 9.5, false, opts.mono?'#eeeeee':'#f1d6de', 'end')); }
  // lanes
  const poolY = G.m + G.headerH, hy = poolY + G.poolH, poolW = laneX(lanes.length) - lanesX;
  o.push('<rect x="'+lanesX+'" y="'+poolY+'" width="'+poolW+'" height="'+G.poolH+'" fill="'+C.laneHead+'" stroke="'+C.grid+'"/>');
  o.push(txt(lanesX+12, poolY+G.poolH/2+4, 'Pool: '+poolName(), 11, true, C.ink));
  o.push(txt(lanesX+poolW-12, poolY+G.poolH/2+4, l4.l4no+' '+l4.l4, 9.5, false, C.muted, 'end'));
  lanes.forEach((ln,i)=>{
    const x = laneX(i), lh = laneHead[i];
    o.push('<rect x="'+x+'" y="'+hy+'" width="'+G.laneW+'" height="'+LH+'" fill="'+C.laneHead+'" stroke="'+C.grid+'"/>');
    const blockH = lh.lines.length*14 + 14, top0 = hy + (LH - blockH)/2 + 11;
    lh.lines.forEach((l,j)=>o.push(txt(x+G.laneW/2, top0 + j*14, l, lh.size, true, C.ink, 'middle')));
    o.push(txt(x+G.laneW/2, top0 + lh.lines.length*14 + 2, lh.sub, 9, false, C.muted, 'middle', ln.kind!=='act'));
    o.push('<rect x="'+x+'" y="'+(hy+LH)+'" width="'+G.laneW+'" height="'+(bodyEnd-hy-LH)+'" fill="'+(i%2?C.laneB:C.laneA)+'" stroke="'+C.grid+'"/>');
  });
  o.push('<rect x="'+lanesX+'" y="'+poolY+'" width="'+poolW+'" height="'+(bodyEnd-poolY)+'" fill="none" stroke="'+C.ink+'" stroke-width="1.6"/>');
  o.push('<rect x="'+px+'" y="'+poolY+'" width="'+pw+'" height="'+(hy+LH-poolY)+'" fill="'+C.laneHead+'" stroke="'+C.grid+'"/>');
  o.push(txt(px+pw/2, poolY+(hy+LH-poolY)/2, 'Keterangan', 12, true, C.ink, 'middle'));
  o.push(txt(px+pw/2, poolY+(hy+LH-poolY)/2+15, '(catatan, di luar diagram BPMN)', 9, false, C.muted, 'middle', true));
  o.push('<rect x="'+px+'" y="'+(hy+LH)+'" width="'+pw+'" height="'+(bodyEnd-hy-LH)+'" fill="#fff" stroke="'+C.grid+'"/>');
  // Semantic loop styling is independent of the geometric routing direction.
  const flowAdj = new Map();
  for(const e of edges) if(!e.side&&!e.reference){ if(!flowAdj.has(e.s))flowAdj.set(e.s,[]); flowAdj.get(e.s).push(e.t); }
  const reaches = (from,to) => { const seen=new Set(), todo=[from]; while(todo.length){const n=todo.pop();if(n===to)return true;if(seen.has(n))continue;seen.add(n);todo.push(...(flowAdj.get(n)||[]));}return false; };
  for(const e of edges){
    if(e.reference){ e.dataReference=!!e.s.crossData; e.loopStyle=false; }
    else e.loopStyle=!!(!e.side && e.back && reaches(e.t,e.s));
  }
  // v1.4.0: gateway question goes beside the diamond, on a side no flow uses.
  for(const n of nodes) if(n.type==='dec'){
    const used=new Set();
    for(const e of edges){ if(!e.pts)continue; const p=e.s===n?e.pts[0]:e.t===n?e.pts[e.pts.length-1]:null; if(!p)continue;
      if(p[0]<n.cx-1)used.add('left'); else if(p[0]>n.cx+1)used.add('right'); }
    const lw=Math.max(0,...n.lines.map(t=>tw(t,10.5,true)))+8, lh=n.lines.length*13+6;
    const cand={left:{x:n.cx-n.w/2-6-lw,y:n.cy-lh/2,w:lw,h:lh}, right:{x:n.cx+n.w/2+6,y:n.cy-lh/2,w:lw,h:lh}, top:{x:n.cx-8-lw,y:n.cy-n.h/2-4-lh,w:lw,h:lh}, bottom:{x:n.cx-8-lw,y:n.cy+n.h/2+4,w:lw,h:lh}};
    // v1.5.0: pick the side no flow touches and the fewest passing flows cross.
    const crossCount=B=>edges.reduce((c,e)=>c+(e.pts?routeSegments(e.pts).filter(([[x,y],[u,v]])=> x===u ? x>B.x&&x<B.x+B.w&&Math.max(y,v)>B.y&&Math.min(y,v)<B.y+B.h : y===v&&y>B.y&&y<B.y+B.h&&Math.max(x,u)>B.x&&Math.min(x,u)<B.x+B.w).length:0),0);
    const outLane=B=>B.x<laneX(n.lane)+2||B.x+B.w>laneX(n.lane)+G.laneW-2;
    const score=k=>(used.has(k)||(k==='top'&&used.has('left'))||(k==='bottom'&&used.has('left'))?1000:0)+crossCount(cand[k])*100+(outLane(cand[k])?50:0)+({left:0,right:1,top:2,bottom:3})[k];
    n._qSide=['left','right','top','bottom'].sort((a,b)=>score(a)-score(b))[0];
    n._qBox=cand[n._qSide];
  }
  // edges
  for(const e of edges){
    const d = roundPath(e.pts, 6);
    if(e.side && (e.s.type==='docin'||e.t.type==='docout')){
      // BPMN Data Association: dotted line, open arrowhead.
      o.push('<path d="'+d+'" fill="none" stroke="'+C.docS+'" stroke-width="1.2" stroke-dasharray="2.5 3.5" marker-end="url(#arA)"/>');
      continue;
    }
    const prop = isProp(e) || isProp(e.s) || isProp(e.t);
    const col = prop ? C.prop : e.loopStyle ? C.back : C.edge, mk2 = prop ? 'arP' : e.loopStyle ? 'arR' : 'ar';
    if(e.reference){
      // Cross-L4 source is informational provenance, not BPMN sequence flow.
      // Do not draw an incoming sequence arrow from a separate L4 reference.
      o.push('<path d="'+d+'" fill="none" stroke="'+C.muted+'" stroke-width="1.2" stroke-dasharray="2 5"/>');
    }else o.push('<path d="'+d+'" fill="none" stroke="'+col+'" stroke-width="1.4"'+(prop?' stroke-dasharray="6 4"':'')+' marker-end="url(#'+mk2+')"/>');
  }
  // nodes
  for(const n of nodes) o.push(drawNode(n));
  // Detached artifacts share the SVG coordinate system with PDF/Word exports.
  for(const n of nodes) if(n.type==='act'&&(n.tag||n.sla||n.crossInputs?.length||n.extraOuts?.length)) o.push(bpaDetachedArtifact(n));
  // Edge labels: measure full text, wrap and choose a low-conflict position near its own route.
  const labelRects=[];
  const intersects=(a,b)=>a.x<b.x+b.w&&a.x+a.w>b.x&&a.y<b.y+b.h&&a.y+a.h>b.y;
  const obstacles=[...nodes.map(n=>({x:n.cx-n.w/2-3,y:n.cy-n.h/2-3,w:n.w+6,h:n.h+6})), ...nodes.filter(n=>n._qBox).map(n=>n._qBox), ...nodes.filter(n=>n._evBox).map(n=>n._evBox)];
  const allSeg=edges.flatMap(e=>routeSegments(e.pts).map(p=>({e,p})));
  for(const e of edges){
    if(!e.label)continue;
    const lines=[];let line='';
    for(const word of String(e.label).split(/\s+/)){
      if(line&&tw(line+' '+word,10.5,true)>180){lines.push(line);line='';}
      let part=word;while(tw(part,10.5,true)>180&&part.length>1){let k=part.length-1;while(k>1&&tw(part.slice(0,k),10.5,true)>180)k--;if(line){lines.push(line);line='';}lines.push(part.slice(0,k));part=part.slice(k);}
      line+=(line?' ':'')+part;
    }if(line)lines.push(line);if(!lines.length)continue;
    const w=Math.max(...lines.map(l=>tw(l,10.5,true)))+6,h=lines.length*13+3;
    let best=null,score=Infinity;
    const offer=(x,y,dist)=>{
      const r={x,y,w,h};if(x<lanesX+2||x+w>laneX(lanes.length)-2||y<y0||y+h>bodyEnd) return;
      let c=dist+obstacles.filter(n=>intersects(r,n)).length*100000+labelRects.filter(n=>intersects(r,n)).length*100000;
      for(const {e:f,p:[[a,b],[u,v]]} of allSeg){if(f===e)continue; if(a===u&&a>x&&a<x+w&&Math.max(b,v)>y&&Math.min(b,v)<y+h||b===v&&b>y&&b<y+h&&Math.max(a,u)>x&&Math.min(a,u)<x+w)c+=8000;} // v1.6.3: a label on another flow is worse than a longer distance
      if(c<score){score=c;best=r;}
    };
    let dist=0;
    for(const [[a,b],[u,v]] of routeSegments(e.pts)){
      const len=Math.abs(u-a)+Math.abs(v-b);
      for(const t of [.1,.25,.4,.55,.7,.85]){const x=a+(u-a)*t,y=b+(v-b)*t,d=dist+len*t;
        if(a===u){offer(x+5,y-h/2,d);offer(x-w-5,y-h/2,d+1);}
        else{offer(x-w/2,y-h-5,d);offer(x-w/2,y+5,d+1);}
      }dist+=len;
    }
    if(!best){const [x,y]=e.pts[0];best={x:Math.max(lanesX+2,Math.min(x+6,laneX(lanes.length)-w-2)),y:Math.max(y0,Math.min(y+6,bodyEnd-h)),w,h};}
    labelRects.push(best);e.labelBox=best;
    o.push('<rect x="'+best.x.toFixed(1)+'" y="'+best.y.toFixed(1)+'" width="'+w.toFixed(1)+'" height="'+h+'" fill="#fff" opacity=".92"/>');
    o.push('<g><title>'+esc(e.label)+'</title>'+lines.map((l,i)=>txt(best.x+3,best.y+12+i*13,l,10.5,true,(isProp(e)||isProp(e.t))?C.prop:e.loopStyle?C.back:C.text,'start')).join('')+'</g>');
  }
  // panel
  o.push(drawPanel(pLines, px+14, y0+10));
  // legend
  o.push(txt(G.m, bodyEnd+20, 'Legenda notasi BPMN dan anotasi', 11, true, C.ink));
  for(const it of legPos) o.push(drawLegend(it));
  footLines.forEach((l,i)=>o.push(txt(G.m, footY+i*12, l, 9, false, C.muted)));
  o.push('</svg>');
  return {svg:o.join(''), w:W, h:H, acts:seq.map(a=>({...a,role:baseRole(a),next:flowString(flowOf.get(a.l5no)||[]),prev:flowString([...(startOf.has(a.l5no)?[{kind:'start',note:startNote.get(a.l5no)}]:[]),...Array.from(flowOf).filter(([c,fs])=>fs.some(f=>f.kind==='code'&&f.code===a.l5no)).map(([c])=>({kind:'code',code:c})),...parseFlow(a.prev).filter(f=>f.kind==='code'&&!inL4.has(f.code))])})), nodes, edges, drawn, lanes: lanes.map((l,i)=>({key:l.role, kind:l.kind, roles:[...(laneRoles.get(l.role)||[l.role])], x:laneX(i)})), laneW:G.laneW, top:G.m+G.headerH+G.poolH, bodyTop:y0};
}
const isProp = x => !!(x && x.elab && x.elab.status==='proposed');
function slaShort(s){
  return String(s||'').replace(/hari\s+kerja/ig,'HK').replace(/hari\s+kalender/ig,'HKal').replace(/\bhari\b/ig,'H').replace(/minggu/ig,'Mg').replace(/bulan/ig,'Bln')
    .replace(/tahun/ig,'Thn').replace(/\bjam\b/ig,'Jam').replace(/menit/ig,'Mnt').replace(/^maks(imal|\.)?\s*/i,'≤').replace(/sesuai\s+jadwal/ig,'Jadwal').replace(/\s+/g,' ').trim();
}
function simplifyPts(pts){
  const P = [];
  for(const p of pts){ const q = P[P.length-1]; if(q && Math.abs(q[0]-p[0])<0.5 && Math.abs(q[1]-p[1])<0.5) continue; P.push(p); }
  for(let i=1;i<P.length-1;){ const a = P[i-1], b = P[i], c = P[i+1];
    if((Math.abs(a[0]-b[0])<0.5 && Math.abs(b[0]-c[0])<0.5) || (Math.abs(a[1]-b[1])<0.5 && Math.abs(b[1]-c[1])<0.5)) P.splice(i,1); else i++; }
  return P;
}
function roundPath(P, r){
  const f = p => p[0].toFixed(1)+','+p[1].toFixed(1);
  let d = 'M'+f(P[0]);
  for(let i=1;i<P.length;i++){
    const b = P[i];
    if(i < P.length-1){
      const a = P[i-1], c = P[i+1];
      const l1 = Math.hypot(b[0]-a[0], b[1]-a[1]), l2 = Math.hypot(c[0]-b[0], c[1]-b[1]); const rr = Math.min(r, l1/2, l2/2);
      const p1 = [b[0]-(b[0]-a[0])/l1*rr, b[1]-(b[1]-a[1])/l1*rr], p2 = [b[0]+(c[0]-b[0])/l2*rr, b[1]+(c[1]-b[1])/l2*rr];
      d += ' L'+f(p1)+' Q'+f(b)+' '+f(p2);
    } else d += ' L'+f(b);
  }
  return d;
}
function changeText(i){
  const nm = c => { const r = state.byCode.get(c); return r ? c+' '+r.l5 : c; };
  if(i.type==='newact') return 'Proses baru setelah '+nm(i.after)+': "'+i.text+'"'+(i.role?' (PIC: '+roleInfo(i.role).name+')':'');
  if(i.type==='reorder') return nm(i.subject)+' dipindah ke setelah '+nm(i.after);
  if(i.type==='rename') return nm(i.after)+' → nama baru: "'+i.text+'"';
  if(i.type==='role'){ const r = state.byCode.get(i.after); return 'PIC '+nm(i.after)+': '+roleInfo((r&&r.role)||'-').name+' → '+roleInfo(i.role).name; }
  return i.text;
}

function txt(x, y, s, size, bold, fill, anchor, italic){
  return '<text x="'+(+x).toFixed(1)+'" y="'+(+y).toFixed(1)+'" font-size="'+size+'"'+(bold?' font-weight="bold"':'')+(italic?' font-style="italic"':'')+' fill="'+fill+'"'+(anchor&&anchor!=='start'?' text-anchor="'+anchor+'"':'')+'>'+esc(s)+'</text>';
}
function textBlock(lines, cx, cy, size, bold, fill, lh){
  const h = lines.length*lh; let y = cy - h/2 + lh*0.78;
  return lines.map((l,i)=>txt(cx, y+i*lh, l, size, bold, fill, 'middle')).join('');
}
let bpmnPreviewRendering=true;
function bpmnMark(v){return /^x$/i.test(String(v||'').trim());}
function bpaCrossL4Data(sourceCode, receivingCode){
  // Explicit predecessor only. Never infer data equivalence from similar wording.
  const source=crossSourceRow(sourceCode); // v1.6.5: follows the source L4's TO-BE draft when a TO-BE view is shown
  if(!source || source.l4no===receivingCode) return null;
  const eff=effectiveData(source); if(eff) return {kind:eff.kind, label:eff.label||('Data dari '+sourceCode), source:sourceCode};
  const kind=bpmnDataKind(source);
  if(!kind)return null;
  const raw=String(kind==='store'?source.systemX:source.docX||'').trim();
  const fallback=String(source.outputX||source.outX||'').trim().split(/\s*\n+\s*/).filter(Boolean).join(' · '); // v1.6.3: keep item boundaries
  return {kind,label:raw.toUpperCase()==='X'?(fallback||'Data dari '+sourceCode):raw,source:sourceCode};
}
function bpaCrossL4OutputTargets(a, currentL4){
 const refs=[];
 for(const f of parseFlow(a.next||'')){
  if(f.kind!=='code'||!f.code)continue;
  const dest=state.byCode&&state.byCode.get(f.code);
  const l4=dest?.l4no || (/^\d{2}\.\d{2}$/.test(f.code)?f.code:null);
  if(l4&&l4!==currentL4)refs.push({l4,code:f.code});
 }
 return [...new Map(refs.map(x=>[x.code,x])).values()];
}
function bpaDetachedArtifact(n){
  // Side-car layout: artifacts must NEVER occupy the central sequence-flow axis.
  // Use a single compact column beside the task, within the active swimlane.
  const out=[];
  const d=state.diagram;
  const lane=d&&d.lanes&&d.lanes[n.lane];
  // Reserve a right-hand artifact rail inside each 400px lane.
  // Task center is intentionally offset left by 75px in _buildDiagram.
  // No clamp against the task: the rail is a separate 110px-wide zone.
  const side=1;
  const taskEdge=n.cx+n.w/2;
  const artifactCenter=lane ? lane.x+d.laneW-84 : taskEdge+84; // v1.6.2: keeps a clear corridor for return flows at the lane edge
  const y=n.cy+10+slaOff(n)+(n.crossInputs?.length||0)*RAIL_STEP; // v1.6.2: icons sit below the task's upper edge so return flows can pass above them
  if(n.tag){
    const t=n.tag, kind=t.kind;
    const p=t.items&&t.items.some(i=>i.status==='proposed');
    const fill=p?C.propF:C.docF, stroke=p?C.prop:C.docS;
    const gx=artifactCenter-12,gy=y-7;
    // Short horizontal BPMN data association (not a sequence flow).
    const oy=gy+15, k=n.crossInputs?.length||0, sy0=railEntryY(n,k), xo=taskEdge+14+k*10; // v1.6.3: output leaves the task edge below the input entries
    out.push('<path d="M'+taskEdge+','+sy0+(Math.abs(oy-sy0)>1?' H'+xo+' V'+oy:'')+' H'+(artifactCenter-side*15)+'" fill="none" stroke="'+stroke+'" stroke-width="1.1" stroke-dasharray="2.5 3.5"'+(kind==='annotation'?'':' marker-end="url(#arA)"')+'/>');
    if(kind==='annotation'){
      out.push('<path d="M'+(gx+7)+','+gy+' H'+gx+' V'+(gy+30)+' H'+(gx+7)+'" fill="none" stroke="'+stroke+'" stroke-width="1.2"/>');
    }else out.push(bpmnDataGlyph(gx,gy,24,30,kind,stroke,fill));
    const label=String(t.text||'').trim();
    const lines=clampWrap(label,9,false,98,3); // v1.6.1: up to 3 full lines before "…"
    const targets=t.crossL4Targets||[];
    const targetLabel=''; // v1.4.0: destination is shown by the Message End Event, not inferred on the data object
    // A white label backing avoids connector bleed-through.
    out.push('<rect x="'+(artifactCenter-52)+'" y="'+(gy+33)+'" width="104" height="'+(lines.length*12+(targetLabel?18:5))+'" rx="2" fill="#fff" fill-opacity=".96"/>');
    out.push(textBlock(lines,artifactCenter,gy+37+(lines.length*12)/2,9,false,C.docT,12));
    if(targetLabel){
      const shortTarget=targetLabel.length>23?targetLabel.slice(0,20)+'…':targetLabel;
      out.push('<g><title>'+esc('Referensi tujuan eksplisit: '+targets.map(x=>x.code).join(', ')+'. Kesesuaian data belum diverifikasi.')+'</title>'+txt(artifactCenter,gy+45+lines.length*11,shortTarget,8,true,C.docT,'middle')+'</g>');
    }
  }
  if(n.crossInputs?.length){
    const baseY=n.cy-4+slaOff(n);
    n.crossInputs.forEach((info,i)=>{
      // The icon is side-car to the receiver task; no connector goes into Start.
      const gy=baseY+i*RAIL_STEP, gx=artifactCenter-12;
      const kind=info.kind==='store'?'store':'object';
      out.push(bpmnDataGlyph(gx,gy,24,30,kind,C.docS,C.docF));
      const label=String(info.label||'Data').trim();
      const lines=clampWrap(label,9,false,98,3);
      out.push('<rect x="'+(artifactCenter-52)+'" y="'+(gy+32)+'" width="104" height="'+(lines.length*11+16)+'" rx="2" fill="#fff" fill-opacity=".97"/>');
      out.push(txt(artifactCenter,gy+41,info.source?'Dari '+info.source:'Input',8.5,true,C.text,'middle'));
      lines.forEach((line,j)=>out.push(txt(artifactCenter,gy+53+j*11,line,9,false,C.docT,'middle')));
      // v1.6.3: every Data Association ends ON the task edge; lower icons bend up into the task.
      const ey=railEntryY(n,i), sy=gy+16, xv=taskEdge+14+i*10;
      out.push('<path d="M'+(artifactCenter-16)+','+sy+(Math.abs(sy-ey)>1?' H'+xv+' V'+ey:'')+' H'+(taskEdge+3)+'" fill="none" stroke="'+C.docS+'" stroke-width="1.1" stroke-dasharray="2.5 3.5" marker-end="url(#arA)"/>');
    });
  }
  if(n.extraOuts?.length){
    // v1.7.1: extra output documents, below the inputs and the main output; association points away from the task.
    const k=n.crossInputs?.length||0, base=n.cy-4+slaOff(n)+k*RAIL_STEP+(n.tag?92:0);
    n.extraOuts.forEach((info,j)=>{
      const gy=base+j*RAIL_STEP, gx=artifactCenter-12, kind=info.kind==='store'?'store':'object';
      out.push(bpmnDataGlyph(gx,gy,24,30,kind,C.docS,C.docF));
      const lines=clampWrap(String(info.label||'Data').trim(),9,false,98,3);
      out.push('<rect x="'+(artifactCenter-52)+'" y="'+(gy+32)+'" width="104" height="'+(lines.length*11+16)+'" rx="2" fill="#fff" fill-opacity=".97"/>');
      out.push(txt(artifactCenter,gy+41,'Output',8.5,true,C.text,'middle'));
      lines.forEach((line,i)=>out.push(txt(artifactCenter,gy+53+i*11,line,9,false,C.docT,'middle')));
      const idx=k+(n.tag?1:0)+j, sy0=railEntryY(n,idx), xv=taskEdge+14+idx*10, ey=gy+16;
      out.push('<path d="M'+taskEdge+','+sy0+(Math.abs(ey-sy0)>1?' H'+xv+' V'+ey:'')+' H'+(artifactCenter-16)+'" fill="none" stroke="'+C.docS+'" stroke-width="1.1" stroke-dasharray="2.5 3.5" marker-end="url(#arA)"/>');
    });
  }
  if(n.sla){
    // SLA is a BPMN text annotation, placed at the side, never below the task.
    // v1.7.0: first item of the rail, joined by a short straight association from the upper part of the task.
    const sx=artifactCenter-42, sy=n.cy-n.h/2+2, ay=n.cy-n.h/2+10;
    const lines=wrap('SLA: '+slaShort(n.sla),8.5,false,85).slice(0,3);
    const hh=Math.max(20,lines.length*11+7);
    out.push('<path d="M'+taskEdge+','+ay+' H'+sx+'" fill="none" stroke="'+C.docS+'" stroke-width="1" stroke-dasharray="2.5 3.5"/>');
    out.push('<path d="M'+(sx+8)+','+sy+' H'+sx+' V'+(sy+hh)+' H'+(sx+8)+'" fill="none" stroke="'+C.docS+'" stroke-width="1.2"/>');
    out.push(textBlock(lines,sx+49,sy+hh/2,8.5,false,C.text,11));
  }
  return out.join('');
}
function ghForData(t){return 32;}
function bpmnDataKind(act){
  if(!act)return null;
  // Existing workbook fields only: U = Dokumen, W = Sistem/Aplikasi.
  // A value or X indicates that a corresponding artifact was recorded.
  // Data Object takes precedence if both fields are populated; report a warning.
  if(String(act.docX||'').trim()==='-')return null; // v1.6.4: "-" in Dokumen = deliberately no symbol
  if(String(act.docX||'').trim())return 'object';
  if(String(act.systemX||'').trim())return 'store';
  return null;
}
function bpmnDataGlyph(x,y,w,h,kind,stroke,fill){
  if(kind==='store'){
    const ry=Math.min(8,h/5), rx=w/2;
    return '<path d="M'+x+','+(y+ry)+' A'+rx+','+ry+' 0 0 1 '+(x+w)+','+(y+ry)+' V'+(y+h-ry)+' A'+rx+','+ry+' 0 0 1 '+x+','+(y+h-ry)+' Z" fill="'+fill+'" stroke="'+stroke+'" stroke-width="1.2"/>'+
      '<ellipse cx="'+(x+rx)+'" cy="'+(y+ry)+'" rx="'+rx+'" ry="'+ry+'" fill="'+fill+'" stroke="'+stroke+'" stroke-width="1.2"/>';
  }
  return '<path d="M'+x+','+y+' H'+(x+w-11)+' L'+(x+w)+','+(y+11)+' V'+(y+h)+' H'+x+' Z M'+(x+w-11)+','+y+' V'+(y+11)+' H'+(x+w)+'" fill="'+fill+'" stroke="'+stroke+'" stroke-width="1.2"/>';
}
function bpmnEventLabel(n, lines, color){
  // Event label sits to the right of the circle, vertically centred; never on the outgoing flow.
  const x0=n.cx+21, h=lines.length*13, top=n.cy-h/2+10;
  const w=Math.max(0,...lines.map(l=>tw(l.t,l.b?11:10,l.b)))+6;
  n._evBox={x:x0-3,y:top-12,w:w,h:h+4};
  return '<rect x="'+(x0-3)+'" y="'+(top-12)+'" width="'+w+'" height="'+(h+4)+'" rx="2" fill="#fff" fill-opacity=".9"/>'+
    lines.map((l,i)=>txt(x0, top+i*13, l.t, l.b?11:10, l.b, l.b?color:C.text, 'start')).join('');
}
function drawNode(n){
  const x = n.cx-n.w/2, y = n.cy-n.h/2, o = [];
  if(n.type==='act'){
    const p = isProp(n), badge = n.isNew ? 'BARU' : n.code;
    const fill = n.isNew ? (p?C.propF:C.newF) : (p?C.propF:C.box);
    const stroke = p ? C.prop : n.isNew ? C.newS : (n.chg ? C.chgS : C.boxS);
    o.push(n.isNew ? '<g data-elab="'+esc(n.elab.id)+'">' : '<g data-code="'+esc(n.code)+'">');
    o.push('<rect class="box" x="'+x+'" y="'+y+'" width="'+n.w+'" height="'+n.h+'"'+(bpmnPreviewRendering?' rx="12"':'')+' fill="'+fill+'" stroke="'+stroke+'" stroke-width="'+(n.chg||n.isNew?1.9:1.4)+'"'+(p?' stroke-dasharray="5 3"':'')+'/>');
    const bwid = tw(badge, 9.5, true)+16;
    o.push('<rect x="'+(x+n.w-bwid-6)+'" y="'+(y-9)+'" width="'+bwid+'" height="18" rx="9" fill="'+(n.isNew?(p?C.prop:C.newS):C.teal)+'"/>');
    o.push(txt(x+n.w-6-bwid/2, y+4, badge, 9.5, true, '#fff', 'middle'));
    if(n.chg && !n.isNew){
      const lab = uniq(n.chg.map(i=>i.type==='reorder'?'DIPINDAH':i.type==='rename'?'NAMA DIUBAH':'PIC DIUBAH')).join(' · ');
      o.push(txt(x+(n.noteNo?14:2), y-4, (p?'USULAN: ':'')+lab, 8, true, p?C.prop:C.chgS));
    } else if(n.isNew && p) o.push(txt(x+2, y-4, 'USULAN', 8, true, C.prop));
    o.push(textBlock(n.lines, n.cx, n.cy+3, 12.5, true, C.text, 16));
    o.push('</g>');
    if(n.noteNo){ const nt = String(n.noteNo), r = nt.length>1 ? 9.5 : 8.5;
      o.push('<g data-x="'+esc(n.code)+'"><title>Catatan '+nt+': '+esc(n.note)+'</title><circle cx="'+x+'" cy="'+y+'" r="'+r+'" fill="'+C.text+'" stroke="#fff" stroke-width="1.2"/>'+txt(x, y+3.3, nt, 9, true, '#fff', 'middle')+'</g>'); }
    // Data artifacts and SLA are rendered as detached BPMN elements below all task boxes.
  } else if(n.type==='dec'){
    const p = isProp(n) || !!(n.moved && n.moved.status==='proposed');
    o.push('<g class="dmd" data-dec="'+esc(n.act.l5no)+'" data-dk="'+(n.elab?'gate':'formal')+'"><path d="M'+n.cx+','+y+' L'+(x+n.w)+','+n.cy+' L'+n.cx+','+(y+n.h)+' L'+x+','+n.cy+' Z" fill="'+(p?C.propF:C.decF)+'" stroke="'+(p?C.prop:C.decS)+'" stroke-width="1.4"'+(p?' stroke-dasharray="5 3"':'')+'/>');
    if(bpmnPreviewRendering){
      const gType=bpmnGatewayType(n.act.l5no);
      if(gType==='OR')o.push('<circle cx="'+n.cx+'" cy="'+n.cy+'" r="10" stroke="#333" stroke-width="2.5" fill="none"/>');
      else if(gType==='AND')o.push('<path d="M'+n.cx+','+(n.cy-12)+' V'+(n.cy+12)+' M'+(n.cx-12)+','+n.cy+' H'+(n.cx+12)+'" stroke="#333" stroke-width="3.5"/>');
      // v1.7.0: Exclusive (XOR) = empty diamond, which BPMN 2.0 allows; OR and AND keep their markers.
    }
    if(bpmnPreviewRendering && n._qBox){
      // BPMN: the question is a label beside the gateway, never inside or on a flow.
      const b=n._qBox, anchor=n._qSide==='right'?'start':'end', tx=n._qSide==='right'?b.x+4:b.x+b.w-4;
      o.push('<rect x="'+b.x+'" y="'+b.y+'" width="'+b.w+'" height="'+b.h+'" rx="2" fill="#fff" fill-opacity="0.95"/>');
      n.lines.forEach((t,i)=>o.push(txt(tx, b.y+13+i*13, t, 10.5, true, p?C.propT:C.text, anchor)));
    } else o.push(textBlock(n.lines,n.cx,n.cy,10.5,true,p?C.propT:C.text,13));
    o.push('</g>');
    if(p) o.push(txt(x+n.w-14, y+6, 'USULAN', 8, true, C.prop));
  } else if(n.type==='hold'){
    const p = isProp(n);
    o.push('<g data-elab="'+esc(n.elab.id)+'"><rect x="'+x+'" y="'+y+'" width="'+n.w+'" height="'+n.h+'" rx="'+Math.min(16,n.h/2)+'" fill="'+(p?C.propF:C.holdF)+'" stroke="'+(p?C.prop:C.holdS)+'" stroke-width="1.4"'+(p?' stroke-dasharray="5 3"':'')+'/>');
    o.push(textBlock(n.lines, n.cx, n.cy, 10, true, p?C.propT:C.holdT, 13)+'</g>');
    if(p) o.push(txt(x+4, y-4, 'USULAN', 8, true, C.prop));
  } else if(n.type==='datain'){
    const p = isProp(n), sk = 10;
    o.push('<g data-elab="'+esc(n.elab.id)+'"><path d="M'+(x+sk)+','+y+' H'+(x+n.w)+' L'+(x+n.w-sk)+','+(y+n.h)+' H'+x+' Z" fill="'+(p?C.propF:C.dataF)+'" stroke="'+(p?C.prop:C.dataS)+'" stroke-width="1.3"'+(p?' stroke-dasharray="5 3"':'')+'/>');
    o.push(textBlock(n.lines, n.cx, n.cy, 9.5, false, p?C.propT:C.dataT, 12)+'</g>');
    if(p) o.push(txt(x+4, y-4, 'USULAN', 8, true, C.prop));
  } else if(n.type==='start' || n.type==='end'){
    // BPMN: thin circle = Start Event, thick circle = End Event; label beside the circle, clear of the flow.
    o.push('<g data-symbol="'+n.type+'" data-symbol-code="'+esc(n.code||'')+'"><circle cx="'+n.cx+'" cy="'+n.cy+'" r="15" fill="'+C.termF+'" stroke="'+C.termS+'" stroke-width="'+(n.type==='end'?3.3:1.4)+'"/>');
    o.push(bpmnEventLabel(n, n.lines.map(t=>({t,b:true})), C.termT)+'</g>');
  } else if(n.type==='merge'){
    o.push('<g data-merge="xor"><path d="M'+n.cx+','+y+' L'+(x+n.w)+','+n.cy+' L'+n.cx+','+(y+n.h)+' L'+x+','+n.cy+' Z" fill="'+C.decF+'" stroke="'+C.decS+'" stroke-width="1.4"/>');
    o.push('</g>'); // v1.7.0: merge gateway is an empty diamond (XOR)
  } else if(n.type==='offin' || n.type==='offout'){
    // BPMN Message Start Event (from another L4) / Message End Event (to another L4).
    const end=n.type==='offout', code=n.ext ? n.code : '';
    o.push('<g'+(code?' data-code="'+esc(code)+'"':'')+' data-message-event="'+(end?'end':'start')+'" data-ev="'+(end?'hand':'trig')+'" data-ev-act="'+esc(end?(n.src||''):(n.target||''))+'" data-ev-src="'+esc(n.group?n.group.map(g=>g.code).join(','):n.code)+'" style="cursor:pointer"><title>'+esc((end?'Lanjut ke ':'Dipicu dari ')+(n.group?n.group.map(g=>g.code).join(', '):n.code))+'</title>');
    o.push('<circle cx="'+n.cx+'" cy="'+n.cy+'" r="15" fill="'+C.termF+'" stroke="'+C.termS+'" stroke-width="'+(end?3.3:1.4)+'"/>');
    // v1.7.0: plain Start / End Event; the label says which process it comes from or goes to (no envelope).
    o.push(bpmnEventLabel(n, n.evLines, C.termT)+'</g>');
  } else if(n.type==='offin' && n.crossData){
    // The source's marked data type controls the incoming symbol (Data Object / Data Store).
    // The reference is still a predecessor reference; do not fabricate a BPMN message flow.
    const kind=n.crossData.kind==='store'?'store':'object';
    const glyphW=34,glyphH=36,glyphX=n.cx-glyphW/2,glyphY=y+4;
    o.push('<g data-cross-l4="'+esc(n.code)+'" data-cross-kind="'+kind+'">');
    o.push(bpmnDataGlyph(glyphX,glyphY,glyphW,glyphH,kind,C.docS,C.docF));
    o.push(txt(n.cx,glyphY+glyphH+13,'Dari '+n.code,9,true,C.text,'middle'));
    n.crossDataLines.forEach((line,i)=>o.push(txt(n.cx,glyphY+glyphH+26+i*12,line,9,false,C.text,'middle')));
    o.push('</g>');
  } else if(n.type==='offin' || n.type==='offout'){
    const code = n.ext ? n.code : '';
    o.push('<g'+(code?' data-code="'+esc(code)+'"':'')+'><rect x="'+x+'" y="'+y+'" width="'+n.w+'" height="'+n.h+'" fill="'+C.offF+'" stroke="'+C.offS+'" stroke-width="1.4"/>');
    o.push('<line x1="'+(x+7)+'" y1="'+y+'" x2="'+(x+7)+'" y2="'+(y+n.h)+'" stroke="'+C.offS+'" stroke-width="1.2"/><line x1="'+(x+n.w-7)+'" y1="'+y+'" x2="'+(x+n.w-7)+'" y2="'+(y+n.h)+'" stroke="'+C.offS+'" stroke-width="1.2"/>');
    let yy = y+16;
    o.push(txt(n.cx, yy, n.parts.head, 10, true, C.ink, 'middle')); yy += 13;
    n.parts.body.forEach(l=>{ o.push(txt(n.cx, yy, l, 9.5, false, C.text, 'middle')); yy += 12; });
    n.parts.sub.forEach(l=>{ o.push(txt(n.cx, yy, l, 8.5, false, C.muted, 'middle', true)); yy += 11; });
    o.push('</g>');
  } else {
    const b = y+n.h;
    o.push('<g data-symbol="'+esc(n.type)+'" data-symbol-code="'+esc(n.code||'')+'" data-symbol-party="'+esc(n.party||'')+'">'+bpmnDataGlyph(x,y,n.w,n.h,'object',C.docS,C.docF));
    o.push('<rect x="'+(n.cx-84)+'" y="'+(b+3)+'" width="168" height="'+(n.lines.length*12+4)+'" rx="2" fill="#fff" fill-opacity=".95"/>');
    n.lines.forEach((t,i)=>o.push(txt(n.cx, b+14+i*12, t, 10, false, C.docT, 'middle')));
    o.push('</g>');
  }
  return o.join('');
}
function drawLegend(it){
  const x = it.x, y = it.y, o = [];
  if(it.k==='bstart') o.push('<circle cx="'+(x+18)+'" cy="'+y+'" r="8" fill="'+C.termF+'" stroke="'+C.termS+'" stroke-width="1.3"/>');
  if(it.k==='mstart'||it.k==='mend'){ const e=it.k==='mend'; o.push('<circle cx="'+(x+18)+'" cy="'+y+'" r="9" fill="'+C.termF+'" stroke="'+C.termS+'" stroke-width="'+(e?3:1.3)+'"/>'+(e?'<rect x="'+(x+13)+'" y="'+(y-3.5)+'" width="10" height="7" fill="'+C.termS+'"/><path d="M'+(x+13)+','+(y-3.5)+' L'+(x+18)+','+y+' L'+(x+23)+','+(y-3.5)+'" fill="none" stroke="#fff" stroke-width=".9"/>':'<rect x="'+(x+13)+'" y="'+(y-3.5)+'" width="10" height="7" fill="#fff" stroke="'+C.termS+'" stroke-width=".9"/><path d="M'+(x+13)+','+(y-3.5)+' L'+(x+18)+','+y+' L'+(x+23)+','+(y-3.5)+'" fill="none" stroke="'+C.termS+'" stroke-width=".9"/>')); }
  if(it.k==='assoc') o.push('<path d="M'+x+','+y+' L'+(x+34)+','+y+'" stroke="'+C.docS+'" stroke-width="1.2" stroke-dasharray="2.5 3.5" marker-end="url(#arA)"/>');
  if(it.k==='bend') o.push('<circle cx="'+(x+18)+'" cy="'+y+'" r="8" fill="'+C.termF+'" stroke="'+C.termS+'" stroke-width="3"/>');
  if(['gxor','gor','gand'].includes(it.k)){
    o.push('<path d="M'+(x+18)+','+(y-11)+' L'+(x+36)+','+y+' L'+(x+18)+','+(y+11)+' L'+x+','+y+' Z" fill="'+C.decF+'" stroke="'+C.decS+'"/>');
    if(it.k==='gor')o.push('<circle cx="'+(x+18)+'" cy="'+y+'" r="5" stroke="'+C.decS+'" stroke-width="2" fill="none"/>');
    else if(it.k==='gand')o.push('<path d="M'+(x+18)+','+(y-6)+' V'+(y+6)+' M'+(x+11)+','+y+' H'+(x+25)+'" stroke="'+C.decS+'" stroke-width="2"/>');
    /* v1.7.0: XOR legend glyph is an empty diamond */
  }
  if(it.k==='term') o.push('<rect x="'+x+'" y="'+(y-8)+'" width="36" height="16" rx="8" fill="'+C.termF+'" stroke="'+C.termS+'"/>');
  if(it.k==='act') o.push('<rect x="'+x+'" y="'+(y-8)+'" width="36" height="16" rx="4" fill="'+C.box+'" stroke="'+C.boxS+'"/><rect x="'+(x+20)+'" y="'+(y-12)+'" width="14" height="7" rx="3.5" fill="'+C.teal+'"/>');
  if(it.k==='dec') o.push('<path d="M'+(x+18)+','+(y-9)+' L'+(x+36)+','+y+' L'+(x+18)+','+(y+9)+' L'+x+','+y+' Z" fill="'+C.decF+'" stroke="'+C.decS+'"/>');
  if(it.k==='doc') o.push('<path d="M'+(x+8)+','+(y-11)+' H'+(x+26)+' L'+(x+34)+','+(y-3)+' V'+(y+11)+' H'+(x+8)+' Z M'+(x+26)+','+(y-11)+' V'+(y-3)+' H'+(x+34)+'" fill="'+C.docF+'" stroke="'+C.docS+'"/>');
  if(it.k==='off') o.push('<rect x="'+x+'" y="'+(y-8)+'" width="36" height="16" fill="'+C.offF+'" stroke="'+C.offS+'"/><line x1="'+(x+5)+'" y1="'+(y-8)+'" x2="'+(x+5)+'" y2="'+(y+8)+'" stroke="'+C.offS+'"/><line x1="'+(x+31)+'" y1="'+(y-8)+'" x2="'+(x+31)+'" y2="'+(y+8)+'" stroke="'+C.offS+'"/>');
  if(it.k==='arrow') o.push('<path d="M'+x+','+y+' L'+(x+34)+','+y+'" stroke="'+C.edge+'" stroke-width="1.4" marker-end="url(#ar)"/>');
  if(it.k==='back') o.push('<path d="M'+x+','+y+' L'+(x+34)+','+y+'" stroke="'+C.back+'" stroke-width="1.4" stroke-dasharray="6 4" marker-end="url(#arR)"/>');
  if(it.k==='note') o.push('<circle cx="'+(x+18)+'" cy="'+y+'" r="8.5" fill="'+C.text+'"/>'+txt(x+18, y+3.3, '1', 9, true, '#fff', 'middle'));
  if(it.k==='sla') o.push('<path d="M'+(x+14)+','+(y-10)+' H'+(x+6)+' V'+(y+10)+' H'+(x+14)+'" fill="none" stroke="'+C.docS+'" stroke-width="1.2"/>'+txt(x+16, y+3, 'SLA', 7.5, false, C.text, 'start')); // v1.7.0: SLA = BPMN Text Annotation
  if(it.k==='hold') o.push('<rect x="'+x+'" y="'+(y-8)+'" width="36" height="16" rx="8" fill="'+C.holdF+'" stroke="'+C.holdS+'"/>');
  if(it.k==='tag') o.push(bpmnDataGlyph(x+8,y-11,26,22,'object',C.docS,C.docF));
  if(it.k==='store') o.push(bpmnDataGlyph(x+8,y-11,26,22,'store',C.docS,C.docF));
  if(it.k==='datain') o.push('<path d="M'+(x+6)+','+(y-8)+' H'+(x+36)+' L'+(x+30)+','+(y+8)+' H'+x+' Z" fill="'+C.dataF+'" stroke="'+C.dataS+'"/>');
  if(it.k==='new') o.push('<rect x="'+x+'" y="'+(y-8)+'" width="36" height="16" fill="'+C.newF+'" stroke="'+C.newS+'" stroke-width="1.6"/><rect x="'+(x+18)+'" y="'+(y-12)+'" width="16" height="7" rx="3.5" fill="'+C.newS+'"/>');
  if(it.k==='chg') o.push('<rect x="'+x+'" y="'+(y-8)+'" width="36" height="16" fill="'+C.box+'" stroke="'+C.chgS+'" stroke-width="1.6"/>');
  if(it.k==='prop') o.push('<rect x="'+x+'" y="'+(y-8)+'" width="36" height="16" fill="'+C.propF+'" stroke="'+C.prop+'" stroke-dasharray="4 2"/>');
  o.push(txt(x+44, y+4, it.txt, 10, false, C.text));
  return o.join('');
}

/* ---- Keterangan content ---- */
function buildPanel(l4, acts, nodes, edges, offIn, offOut, drawn, posLanes){
  const S = []; drawn = drawn || [];
  const tagE = i => i.status==='proposed' ? ' [usulan AI]' : ' (elaborasi)';
  const ofT = t => drawn.filter(i=>i.type===t).sort((a,b)=>cmpCode(a.after,b.after));
  const L = n => !n ? '' : (n.act && n.act.isNew) ? '"'+n.act.l5+'" (baru)' : (n.code || (n.act && n.act.l5no) || (n.type==='end'?'SELESAI':''));
  const nm = r => { const pos = posFor(l4.l4no, r, ''); return pos ? laneShort(pos) : roleInfo(r).name; };
  // A. INPUT
  const A = [];
  const byParty = new Map();
  for(const a of acts) for(const p of splitParty(a.inParty)){ if(!byParty.has(p)) byParty.set(p, []); byParty.get(p).push(a.l5no); }
  for(const [p,codes] of byParty) A.push('Dari '+nm(p)+' ('+roleInfo(p).code+') ke aktivitas '+codes.join(', ')+'.');
  for(const [okey,n] of offIn){
    const code = String(okey).split('|')[0];
    const targets = edges.filter(e=>e.s===n).map(e=>L(e.t)).join(', ');
    { const src=n.group?state.byCode.get(code):n.ext; A.push('Dari prosedur lain: '+code+(src?' '+src.l5:'')+' → '+targets+'.'); }
  }
  const starts = edges.filter(e=>e.s.type==='start').map(e=>L(e.t));
  if(starts.length) A.unshift('Titik mulai (MULAI): aktivitas '+uniq(starts).join(', ')+'.');
  for(const i of ofT('datain')) A.push('Input data dari luar prosedur ke '+i.after+': '+i.text+tagE(i)+'.');
  for(const a of acts){ const ins=docsOf(a.l5no).filter(d=>d.d!=='out'); if(ins.length) A.push('Input '+(a.isNew?'"'+a.l5+'"':a.l5no)+': '+ins.map(d=>d.n+(d.k==='store'?' (data)':'')).join(', ')+'.'); } // v1.7.1
  S.push({h:'A. INPUT', items:A.length?A:['Tidak ada input pihak lain yang tercatat di workbook.']});
  // B. PROSES
  const roleOf = a => { if(a.posX) return laneShort(sepPos(a.posX).join(' / ')); const ri = roleInfo(a._role||a.role||''); return ri.pos ? laneShort(ri.name) : ri.name; }; // v1.5.0: jabatan, not process role
  S.push({h:'B. PROSES', codes:acts.map(a=>{ const x = state.extra[a.l5no]; return [a.isNew?'BARU':a.l5no, a.l5, roleOf(a)+(x && x.sla ? '; SLA: '+x.sla : '')]; })});
  // C. OUTPUT
  const Cc = [];
  const outParty = new Map();
  for(const a of acts) for(const p of splitParty(a.outParty)){ if(!outParty.has(p)) outParty.set(p, []); outParty.get(p).push(a.l5no); }
  for(const [p,codes] of outParty) Cc.push('Ke '+nm(p)+' ('+roleInfo(p).code+') dari aktivitas '+codes.join(', ')+'.');
  { const byCode=new Map(); for(const n of offOut.values()){ if(!byCode.has(n.code)) byCode.set(n.code,[]); byCode.get(n.code).push(n); } /* v1.7.0: one end event per ending lane, one line per target */
  for(const [code,ns] of byCode){
    const n=ns[0]; const srcs = uniq(edges.filter(e=>ns.includes(e.t)).map(e=>L(e.s))).join(', ');
    Cc.push('Lanjut ke '+code+(n.ext?' '+n.ext.l5+' (prosedur '+n.ext.l4no+')':' (tidak ditemukan di workbook)')+', dari '+srcs+'.');
  } }
  const ends = uniq(edges.filter(e=>e.t.type==='end').map(e=>L(e.s)));
  if(ends.length) Cc.push('Proses selesai (SELESAI) setelah '+ends.join(', ')+'.');
  for(const a of acts){ const outs=docsOf(a.l5no).filter(d=>d.d==='out'); if(outs.length) Cc.push('Output '+(a.isNew?'"'+a.l5+'"':a.l5no)+': '+outs.map(d=>d.n+(d.k==='store'?' (data)':'')).join(', ')+'.'); } // v1.7.1
  for(const a of acts){ const x = state.extra[a.l5no]; if(x && cleanOut(x.out)) Cc.push('Dokumen/output '+(a.isNew?'"'+a.l5+'"':a.l5no)+': '+cleanOut(x.out)+'.'); }
  for(const i of ofT('doc')) Cc.push('Dokumen/bukti dari '+i.after+': '+i.text+tagE(i)+'.');
  S.push({h:'C. OUTPUT', items:Cc.length?Cc:['Tidak ada output pihak lain yang tercatat di workbook.']});
  // D. KENDALI
  const D = [];
  const decs = new Map();
  for(const a of acts) if(a.dec){ if(!decs.has(a.dec)) decs.set(a.dec, []); decs.get(a.dec).push(a.l5no); }
  for(const [d,codes] of decs) D.push(d+': aktivitas '+codes.join(', ')+'.');
  for(const n of nodes) if(n.type==='dec' && !n.elab){
    const br = edges.filter(e=>e.s===n).map(e=>(e.label||'?')+' → '+L(e.t));
    D.push('Percabangan setelah '+n.act.l5no+': '+br.join('; ')+'.');
  }
  for(const i of ofT('gate')) D.push('Gateway pemeriksaan setelah '+i.after+(i.role?' (di kolom '+roleInfo(i.role).name+')':'')+': "'+i.text+'" Bila tidak lolos → '+i.noTarget+(i.noLabel?' ('+i.noLabel+')':'')+tagE(i)+'.');
  for(const i of ofT('hold')) D.push('HOLD dari '+i.after+': '+i.text+tagE(i)+'.');
  for(const i of ofT('declane')) D.push('Keputusan setelah '+i.after+' digambar di kolom '+laneShort(roleInfo(i.role).name)+tagE(i)+'.');
  const group = (field) => { const m = new Map(); for(const a of acts){ const v=a[field]; if(!v) continue; if(!m.has(v)) m.set(v,[]); m.get(v).push(a.l5no); } return m; };
  const changed = new Set(drawn.filter(i=>['reorder','role','newact'].includes(i.type)).flatMap(i=>[i.after, i.subject].filter(Boolean)));
  drawn.filter(i=>i.type==='role' || i.type==='reorder').forEach(i=>changed.add(i.type==='role' ? i.after : i.subject));
  const flag = (v, codes) => { const refs = new Set([...codes, ...(String(v).match(CODE_RE)||[])]); return [...refs].some(c=>drawn.some(i=>(i.type==='role' && i.after===c) || (i.type==='reorder' && i.subject===c))) ? ' [cek: terkait perubahan di E]' : ''; };
  for(const [v,codes] of group('noteRole')) D.push(codes.join(', ')+' — '+v+flag(v, codes));
  for(const [v,codes] of group('note')) D.push(codes.join(', ')+' — '+v+flag(v, codes));
  S.push({h:'D. KENDALI DAN ANTARMUKA', items:D.length?D:['Tidak ada catatan kendali di workbook.']});
  const E = [];
  for(const i of drawn.filter(i=>['newact','reorder','role','rename'].includes(i.type))) E.push(changeText(i)+(i.status==='proposed'?' [usulan]':'')+'.');
  const RE = state.roleEdit, inThis = new Set(rolesInL4(l4));
  for(const [r,al] of Object.entries(RE.alias)) if(inThis.has(r) || (RE.added[l4.l4no]||[]).includes(r)) E.push('Nama peran "'+(roleInfo(r).orig||r)+'" ditampilkan sebagai "'+al+'".');
  for(const r of (RE.added[l4.l4no]||[])) E.push('Peran baru ditambahkan: '+roleInfo(r).name+'.');
  for(const r of (RE.hidden[l4.l4no]||[])) E.push('Lane pihak '+roleInfo(r).name+' disembunyikan dari flowchart.');
  const NT = acts.filter(a=>a._note).map(a=>'('+a._note+') '+a.l5no+' '+a.l5+': '+state.extra[a.l5no].note);
  if(NT.length) S.push({h:'CATATAN DISKUSI (nomor sesuai penanda bulat hitam di flowchart)', items:NT});
  if(E.length) S.push({h:'E. PERUBAHAN TERHADAP WORKBOOK', items:E});
  if(false && posLanes && posLanes.length) S.push({h:'F. JABATAN PER LANE', items: posLanes.map(l=>l.roles.join(', ')+': '+sepPos(l.pos).join('; ')+'.').concat(['Sumber: sheet Peran Jabatan; jabatan khusus per aktivitas dari kolom "Jabatan (khusus aktivitas ini)" di SOP.'])});
  return S;
}
function layoutPanel(S, maxW){
  const L = []; let y = 0;
  for(const s of S){
    if(L.length) y += 10;
    L.push({t:'h', s:s.h, y:y+12}); y += 20;
    if(s.items) for(const it of s.items){
      const ls = wrap(it, 9.5, false, maxW-12);
      ls.forEach((l,i)=>{ L.push({t:'li', s:l, y:y+11, bullet:i===0}); y += 12.5; });
      y += 2;
    }
    if(s.codes) for(const [c,name,role] of s.codes){
      const ls = wrap(name+(role?' ('+role+')':''), 9.5, false, maxW-62);
      ls.forEach((l,i)=>{ L.push({t:'code', c:i===0?c:'', s:l, y:y+11}); y += 12.5; });
      y += 2;
    }
  }
  L.h = y; return L;
}
function drawPanel(L, x, y){
  return L.map(l=>{
    if(l.t==='h') return txt(x, y+l.y, l.s, 11, true, C.ink);
    if(l.t==='li') return (l.bullet?txt(x, y+l.y, '•', 9.5, false, C.text):'')+txt(x+11, y+l.y, l.s, 9.5, false, C.text);
    return (l.c?txt(x, y+l.y, l.c, 9.5, true, C.teal):'')+txt(x+62, y+l.y, l.s, 9.5, false, C.text);
  }).join('');
}

/* ================= Rendering in page ================= */
function hashToL4(){ const h = location.hash.replace('#',''); const m = h.match(/^p-([\w-]+)$/); return m ? m[1].replace(/-/g,'.') : null; }
function select(l4no, initial){
  if(!l4no || !state.l4s.has(l4no)) return;
  const changed = state.current !== l4no;
  state.current = l4no; updatePosToggle();
  const g = state.l4s.get(l4no);
  $('crumb').textContent = g.l1+' › '+g.l2+' › '+g.l3+' › '+g.l4no+' '+g.l4;
  try{ history.replaceState(null,'','#p-'+l4no.replace(/\./g,'-')); }catch(e){}
  redraw();
  renderTree();
  const a = document.querySelector('.l4.active'); if(a && !initial) a.scrollIntoView({block:'nearest'});
  if(initial) fitZoom();
  if(changed){ state.editing = null; state.freeText = ''; state.removing = null; if(state.sideMode==='elab') renderElab(); else if(state.sideMode==='extra') renderExtra(); else if(state.sideMode==='roles') renderRoles(); else $('detail').classList.add('hidden'); }
}
function redraw(){
  syncWorkspace();
  const g = state.l4s.get(state.current); if(!g) return;
  bpmnPreviewRendering=true;
  try{state.diagram = buildDiagram(g, $('statusLabel').value.trim(), {ver: $('processEditMode').value==='asis' ? 'asis' : ($('processEditMode').value==='tobe' || state.showProposed) ? 'tobe' : 'approved', items: $('processEditMode').value==='asis' ? [] : elabFor(g.l4no, $('processEditMode').value==='tobe' || state.showProposed), mono: !state.colorView});}
  finally{bpmnPreviewRendering=true;}
  closeLanePop(); $('canvas').innerHTML = state.diagram.svg;
  applyZoom(); renderWarnings();
}
function applyZoom(){
  const svg = $('canvas').querySelector('svg'); if(!svg || !state.diagram) return;
  svg.style.width = (state.diagram.w*state.zoom)+'px'; svg.style.height = (state.diagram.h*state.zoom)+'px';
  $('zoomTxt').textContent = Math.round(state.zoom*100)+'%';
}
function fitZoom(){ if(!state.diagram) return; const w = $('canvas').clientWidth - 34; state.zoom = Math.max(0.2, Math.min(1.5, w/state.diagram.w)); applyZoom(); }
function renameOf(code){ const r = state.byCode.get(code); return r ? (state.elab[r.l4no]||[]).find(i=>i.type==='rename' && i.after===code && i.status==='accepted') : null; }
function actName(a){ const it = renameOf(a.l5no); return it && it.text ? it.text : a.l5; }
function setRename(code, text){
  const r = state.byCode.get(code); if(!r) return; text = String(text||'').replace(/\s+/g,' ').trim();
  const list = elabOf(r.l4no); let it = list.find(i=>i.type==='rename' && i.after===code);
  if(!text || text===r.l5){ if(it) list.splice(list.indexOf(it), 1); toast('Nama aktivitas '+code+' kembali ke nama di workbook.'); }
  else { if(!it){ it = {id:newId(), l4:r.l4no, type:'rename', after:code, source:'manual', reason:'Diganti langsung di flowchart'}; list.push(it); }
    it.text = text; it.status = 'accepted'; toast('Nama aktivitas '+code+' diganti.'); }
  elabChanged(); showDetail(code);
}
function showDetail(code, editName){
  const r = state.byCode.get(code); if(!r) return;
  if(r.l4no !== state.current){ select(r.l4no); }
  const fromElab = state.sideMode === 'elab';
  state.sideMode = 'detail'; ['btnElab','btnX','btnR'].forEach(x=>$(x).classList.remove('on'));
  const f = (k,v) => v ? '<dt>'+esc(k)+'</dt><dd>'+esc(v)+'</dd>' : '';
  $('detail').innerHTML = '<button class="tb x" data-a="close">✕</button>'+(fromElab?'<button class="sbtn" data-a="openElab" style="margin-bottom:8px">← Usulan &amp; Revisi</button>':'')+
    '<div class="dcode">'+esc(r.l5no)+'</div>'+(()=>{ const rn = renameOf(code), cur = rn ? rn.text : r.l5;
      if(editName) return '<div class="rnbox"><label style="font-size:12px;color:var(--muted)">Nama aktivitas</label><textarea id="rnInput" rows="2" style="width:100%;font:inherit;font-size:14px;padding:6px 8px;border:1px solid var(--line);border-radius:6px;resize:vertical">'+esc(cur)+'</textarea>'+
        '<div class="cbtn" style="margin-top:6px"><button class="sbtn ok" data-a="rnSave">Simpan</button><button class="sbtn" data-a="rnCancel">Batal</button>'+(rn?'<button class="sbtn no" data-a="rnReset">Kembalikan nama asli</button>':'')+'</div>'+
        '<div style="font-size:11.5px;color:var(--muted);margin-top:6px">Nama di workbook: '+esc(r.l5)+'</div></div>';
      return '<h3>'+esc(cur)+' <button class="sbtn" data-a="rnEdit" title="Ganti nama aktivitas" style="font-size:12px;padding:2px 8px;vertical-align:middle">✎ Ganti nama</button></h3>'+(rn?'<div style="font-size:11.5px;color:var(--muted);margin:-4px 0 8px">Nama di workbook: '+esc(r.l5)+'</div>':''); })()+'<dl>'+
    f('Deskripsi aktivitas', r.desc)+f('Unit',r.unitX)+f('Jabatan aktivitas',r.posX)+f('Role keputusan',r.decisionRole)+f('Input',r.inputX)+f('Output',r.outputX)+f('Dokumen',r.docX)+f('Sistem/Aplikasi',r.systemX)+f('Responsible Role', r.role)+f('Jabatan', posFor(r.l4no, r.role, r.dec, r))+f('SLA', (state.extra[r.l5no]||{}).sla)+f('Dokumen/output', (state.extra[r.l5no]||{}).out)+f('Catatan diskusi', (state.extra[r.l5no]||{}).note)+f('Decision Ref', r.dec)+f('Pihak pemberi input', r.inParty)+f('Pihak penerima output', r.outParty)+
    f('Proses sebelum', r.prev)+f('Proses sesudah', r.next)+f('Catatan peran', r.noteRole)+f('Dasar antarmuka', r.iface)+f('Catatan', r.note)+
    f('Hierarki', r.l1+' › '+r.l2+' › '+r.l3+' › '+r.l4no+' '+r.l4)+'</dl>';
  $('detail').classList.remove('hidden');
  $('detail').dataset.code = code;
  $('detail').insertAdjacentHTML('beforeend',annotationSection(state.l4s.get(r.l4no),code,'view'));
  if(editName){ const t = $('rnInput'); t.focus(); t.select();
    t.onkeydown = e => { if(e.key==='Enter' && !e.shiftKey){ e.preventDefault(); setRename(code, t.value); } else if(e.key==='Escape') showDetail(code); }; }
}

/* ================= Elaboration (tahap 2) ================= */
const TYPE_NAME = {gate:'Gateway', hold:'HOLD', doc:'Dokumen', datain:'Input data', declane:'Kolom keputusan', newact:'Proses baru', reorder:'Pindah urutan', role:'Ubah PIC', rename:'Ganti nama', graph:'Struktur proses'};
const STRUCT = ['newact','reorder','role','rename'];
const TYPE_FROM = {'gateway':'gate','hold':'hold','dokumen':'doc','dokumen/bukti':'doc','input data':'datain','input':'datain','proses baru':'newact','pindah urutan':'reorder','ubah pic':'role','kolom keputusan':'declane','ganti nama':'rename','ganti nama aktivitas':'rename','struktur proses':'graph'};
const STATUS_NAME = {accepted:'Diterima', proposed:'Usulan', rejected:'Ditolak'};
const STATUS_FROM = {'diterima':'accepted','usulan':'proposed','ditolak':'rejected'};
const newId = () => 'e'+Date.now().toString(36)+Math.random().toString(36).slice(2,6);
function elabOf(l4no){ return state.elab[l4no] || (state.elab[l4no] = []); }
function elabFor(l4no, withProposed){ return (state.elab[l4no]||[]).filter(i=>i.status==='accepted' || (withProposed && i.status==='proposed')); }
function saveElab(){ const clean = {}; for(const [k,v] of Object.entries(state.elab)) if(v && v.length) clean[k] = v; store.set('flowbpa:elab', JSON.stringify(clean)); }
function elabChanged(){ saveElab(); redraw(); renderTree(); if(state.sideMode==='elab') renderElab(); }
function actsOf(g){ return g.acts.filter(a=>!isGateRow(a)).sort((a,b)=>cmpCode(a.l5no,b.l5no)); }

function openElab(focusId){
  state.legacyAiPanel=false;
  if($('processEditMode').value!=='tobe'){$('processEditMode').value='tobe';redraw();}
  state.sideMode = 'elab'; $('btnElab').classList.add('on');
  $('detail').classList.remove('hidden');
  renderElab(focusId);
}
function cardHtml(it, g){
  if(state.editing === it.id && it.type!=='graph') return editHtml(it, g);
  const meta = it.type==='gate' ? 'Tidak lolos → '+esc(it.noTarget||'?')+(it.noLabel?' · label "'+esc(it.noLabel)+'"':'')+(it.role?' · kolom: '+esc(roleInfo(it.role).name):'') :
    it.type==='declane' ? 'Belah ketupat dipindah ke kolom '+esc(roleInfo(it.role||'').name) : it.type==='hold' ? 'Label cabang: "'+esc(it.label||'Hold')+'"' :
    it.type==='newact' ? 'PIC: '+esc(roleInfo(it.role||'(sama dengan '+it.after+')').name) : STRUCT.includes(it.type) ? esc(changeText(it)) : '';
  const anchor = (it.type==='role'||it.type==='rename') ? 'aktivitas '+esc(it.after) : it.type==='declane' ? 'keputusan setelah '+esc(it.after) : it.type==='reorder' ? esc(it.subject||'?')+' → setelah '+esc(it.after) : 'setelah '+esc(it.after);
  const headText = it.type==='graph' ? 'Perubahan struktur proses (snapshot L4)' : it.type==='rename' ? 'Nama baru: '+it.text : it.type==='reorder' ? 'Pindah '+(it.subject||'') : it.type==='role' ? 'PIC → '+roleInfo(it.role||'').name : it.type==='declane' ? 'Keputusan → kolom '+roleInfo(it.role||'').name : it.text;
  const warn = it.type==='gate' && g && (()=>{ const a = g.acts.find(x=>x.l5no===it.after); return a && hasFormalDecision(a); })();
  let btn = '';
  if(it.status==='proposed') btn = '<button class="sbtn ok" data-a="accept">✓ Terima</button><button class="sbtn no" data-a="reject">✕ Tolak</button><button class="sbtn" data-a="edit">Ubah</button>';
  else if(it.status==='accepted') btn = '<button class="sbtn" data-a="edit">Ubah</button><button class="sbtn" data-a="unaccept">Jadikan usulan</button><button class="sbtn no" data-a="del">Hapus</button>';
  else btn = '<button class="sbtn" data-a="restore">Pulihkan</button><button class="sbtn no" data-a="del">Hapus</button>';
  if(ANNOT_TYPES.includes(it.type)) btn=btn.replace(/<button[^>]*data-a="edit"[^>]*>.*?<\/button>/g,'<button class="sbtn" data-a="annotationDetail">Buka aktivitas</button>');
  return '<div class="card'+(it.status==='proposed'?' prop':'')+'" data-id="'+esc(it.id)+'"><span class="src">'+(it.source==='ai'?'AI':it.source==='free'?'Teks bebas':'Manual')+'</span>'+
    '<span class="chip t-'+it.type+'">'+TYPE_NAME[it.type]+'</span><span class="after">'+anchor+'</span>'+
    '<div class="ctext">'+esc(headText)+'</div>'+(meta?'<div class="cmeta">'+meta+'</div>':'')+
    (warn?'<div class="why" style="color:var(--red)">Aktivitas ini sudah punya keputusan resmi; gateway tidak digambar.</div>':'')+
    (it.type==='reorder' && !canMove(g, it.subject) ? '<div class="why" style="color:var(--red)">Aktivitas ini bercabang, jadi tidak bisa dipindah otomatis.</div>':'')+
    (it.type==='declane' && g && !(()=>{ const a = g.acts.find(x=>x.l5no===it.after); return a && hasFormalDecision(a); })() ? '<div class="why" style="color:var(--red)">Aktivitas ini tidak punya keputusan resmi. Untuk gateway elaborasi, atur kolomnya lewat Ubah pada kartu gateway.</div>':'')+
    (it.reason?'<div class="why">'+esc(it.reason)+'</div>':'')+'<div class="cbtn">'+(it.type==='graph'?btn.replace(/<button[^>]*data-a="edit"[^>]*>.*?<\/button>/g,''):btn)+'</div></div>';
}
function editHtml(it, g){
  const acts = actsOf(g);
  const opt = (v, sel, label) => '<option value="'+esc(v)+'"'+(v===sel?' selected':'')+'>'+esc(label)+'</option>';
  const actOpts = sel => acts.map(a=>opt(a.l5no, sel, a.l5no+' '+a.l5)).join('');
  return '<div class="card ef" data-id="'+esc(it.id)+'">'+
    '<label>Jenis</label><select data-f="type">'+Object.keys(TYPE_NAME).filter(t=>t!=='graph'&&(!STRUCT.includes(t)||t===it.type)).map(t=>opt(t, it.type, TYPE_NAME[t])).join('')+'</select>'+
    (it.type==='reorder' ? '<label>Aktivitas yang dipindah</label><select data-f="subject">'+actOpts(it.subject)+'</select>' : '')+
    '<label>'+(it.type==='role'||it.type==='rename'?'Aktivitas':it.type==='reorder'?'Pindahkan ke setelah':it.type==='declane'?'Keputusan setelah aktivitas':'Setelah aktivitas')+'</label><select data-f="after">'+actOpts(it.after)+'</select>'+
    (['reorder','role','declane'].includes(it.type) ? '' : '<label>'+(it.type==='gate'?'Pertanyaan pemeriksaan':it.type==='newact'?'Nama proses baru':it.type==='rename'?'Nama baru aktivitas':'Teks')+'</label><input data-f="text" value="'+esc(it.text||'')+'" placeholder="'+(it.type==='gate'?'mis. Data lengkap?':'')+'">')+
    (['newact','role','gate','declane'].includes(it.type) ? '<label>'+(it.type==='role'?'PIC baru':it.type==='gate'?'Kolom (kosongkan = sama dengan aktivitas)':it.type==='declane'?'Pindah ke kolom':'PIC (kosongkan = sama dengan aktivitas sebelumnya)')+'</label><input data-f="role" list="roleList" value="'+esc(it.role||'')+'" placeholder="mis. PF-03 Pelaksana HC"><datalist id="roleList">'+allRoles().map(r=>'<option value="'+esc(r)+'">').join('')+'</datalist>' : '')+
    (it.type==='gate' ? '<label>Bila tidak lolos, kembali ke</label><select data-f="noTarget">'+actOpts(it.noTarget)+'</select><label>Label cabang tidak lolos</label><input data-f="noLabel" value="'+esc(it.noLabel||'Tidak')+'">' : '')+
    (it.type==='hold' ? '<label>Label cabang</label><input data-f="label" value="'+esc(it.label||'Hold')+'">' : '')+
    '<div class="cbtn"><button class="sbtn ok" data-a="save">Simpan</button><button class="sbtn" data-a="cancel">Batal</button></div></div>';
}
/* v1.3.7 stage 1: read-only semantic comparison, independent of visual settings. */
function changeProjection(g,items){
 const graph=items.filter(i=>i.type==='graph').slice(-1)[0];
 const rows=JSON.parse(JSON.stringify(graph?JSON.parse(graph.text):g.acts));
 if(!Array.isArray(rows)||rows.some(r=>!r||!r.l5no||r.l4no!==g.l4no)||new Set(rows.map(r=>r.l5no)).size!==rows.length)throw Error('Snapshot TO-BE tidak valid.');
 const consumed=consumedIds(rows); items=items.filter(i=>i.type!=='graph'&&!consumed.has(i.id));
 const by=new Map(rows.map(r=>[r.l5no,r])), flows=new Map(rows.map(r=>[r.l5no,parseFlow(r.next)])), starts=new Map(rows.map(r=>[r.l5no,parseFlow(r.prev).filter(f=>f.kind==='start')]));
 for(const i of items)if(i.type==='rename'&&by.has(i.after)&&i.text)by.get(i.after).l5=i.text;
 for(const i of items){
  if(i.type==='role'&&by.has(i.after)&&i.role)by.get(i.after).role=i.role;
  if(i.type==='reorder'&&by.has(i.subject)&&by.has(i.after)&&i.subject!==i.after){const x=i.subject,y=i.after,f=flows.get(x);if(f.length!==1||f[0].label||!['code','end'].includes(f[0].kind))continue;const t=f[0];for(const [c,fs] of flows)if(c!==x)flows.set(c,fs.map(v=>v.kind==='code'&&v.code===x?{...t,label:v.label,note:v.note||t.note}:v));if(starts.get(x)?.length){if(by.has(t.code))starts.set(t.code,starts.get(x));starts.set(x,[]);}flows.set(x,flows.get(y));flows.set(y,[{kind:'code',code:x}]);}
  if(i.type==='newact'&&by.has(i.after)&&i.text){if(by.has(i.id))throw Error('Identitas usulan aktivitas ganda.');const a=by.get(i.after);by.set(i.id,{l5no:i.id,l4no:g.l4no,l5:i.text,role:i.role||a.role,prev:'',next:''});flows.set(i.id,flows.get(i.after));flows.set(i.after,[{kind:'code',code:i.id}]);starts.set(i.id,[]);}
 }
 const clean=v=>String(v??'').trim().replace(/\s+/g,' ');
 const fmt=f=>[f.label?clean(f.label)+': ':'',f.kind==='code'?f.code:f.kind==='start'?'MULAI':f.kind==='end'?'SELESAI':clean(f.raw),f.note?' ('+clean(f.note)+')':''].join('');
 const list=fs=>[...new Set(fs.map(fmt))].sort().join('; ');
 const result=new Map();
 for(const [code,a] of by){
  const local=items.filter(i=>i.after===code), role=a.role||'', formal=hasFormalDecision({...a,next:flowString(flows.get(code))});
  const gate=!formal?local.find(i=>i.type==='gate'):null, decl=local.find(i=>i.type==='declane');
  const incoming=[...(starts.get(code)||[]),...parseFlow(a.prev).filter(f=>f.kind!=='start'&&(f.kind!=='code'||!by.has(f.code))),...[...flows].flatMap(([c,fs])=>fs.some(f=>f.kind==='code'&&f.code===code)?[{kind:'code',code:c}]:[])];
  const annotations=local.filter(i=>['hold','doc','datain'].includes(i.type)).map(i=>[i.type,clean(i.text),clean(i.label)].filter(Boolean).join(': ')).sort().join('; ');
  const x=state.extra[code]||{};
  const fields={'Nama aktivitas':a.l5,'Deskripsi':a.desc,'Role':role,'Jabatan':a.posX||(state.posmap[g.l4no]||{})[role],'Unit':a.unitX||(state.unitmap[g.l4no]||{})[role],'Proses sebelum':list(incoming),'Proses sesudah':list(flows.get(code)),'Pertanyaan keputusan':formal?(a.decisionQuestion||decQuestion(flows.get(code))):gate?.text,'Role keputusan':formal||gate?(gate?.role||decl?.role||a.decisionRole||role):'','Cabang gateway':gate?[gate.noLabel||'Tidak',gate.noTarget].join(': '):'','Input':a.inputX,'Output':a.outputX,'Dokumen':a.docX,'SLA':x.sla||a.slaX,'Sistem/Aplikasi':a.systemX,'Output tambahan':x.out||a.outX,'Pihak input':a.inParty,'Pihak output':a.outParty,'Label mulai':a.startLabel,'Label selesai':a.endLabel,'Label dokumen':a.symbolDocs,'Anotasi':annotations};
  result.set(code,{name:clean(a.l5),fields:Object.fromEntries(Object.entries(fields).map(([k,v])=>[k,clean(v)]))});
 }
 return result;
}
function compareProcess(g,latest=true){
 const before=changeProjection(g,[]),after=changeProjection(g,elabFor(g.l4no,latest)),changes=[];
 for(const code of new Set([...before.keys(),...after.keys()])){const a=before.get(code),b=after.get(code);const type=!a?'added':!b?'removed':'modified';const fields=Object.keys((a||b).fields).filter(k=>(a?.fields[k]||'')!==(b?.fields[k]||'')).map(k=>({label:k,before:a?.fields[k]||'',after:b?.fields[k]||''}));if(type!=='modified'||fields.length)changes.push({code,name:(b||a).name,type,fields});}
 return {before:before.size,after:after.size,changes};
}
const REGISTER_KEY='flowbpa:changeRegister';
let changeRegister=[];
function registerId(l4,code){return 'CHG-'+encodeURIComponent(l4)+'/'+encodeURIComponent(code);}
function registerSignature(c){return JSON.stringify([c.type,c.fields.map(f=>[f.label,f.before,f.after])]);}
function registerEntry(l4,code){return changeRegister.find(r=>r.id===registerId(l4,code));}
function registerStatus(r,c){return !r?'Belum dilengkapi':r.signature!==registerSignature(c)?'Perlu ditinjau ulang':'Tersimpan';}
function normalizeRegister(r){if(!r||!r.l4||!r.code)return null;return {id:registerId(String(r.l4),String(r.code)),l4:String(r.l4),code:String(r.code),finding:String(r.finding||''),reason:String(r.reason||''),benefit:String(r.benefit||''),signature:String(r.signature||'')};}
function saveRegister(next){if(!store.set(REGISTER_KEY,JSON.stringify(next)))return false;changeRegister=next;return true;}
function mergeRegister(rows){if(!rows)return;const m=new Map(changeRegister.map(r=>[r.id,r]));for(const raw of rows){const r=normalizeRegister(raw);if(r)m.set(r.id,r);}saveRegister([...m.values()]);}
function registerHtml(g,c){const r=registerEntry(g.l4no,c.code), findings=analysisRecords.filter(x=>x.l4===g.l4no), selected=r?.finding||'';return '<div class="register-fields" style="border-top:1px solid #ddd;margin-top:12px;padding-top:10px" data-register-code="'+esc(c.code)+'"><b>'+esc(registerId(g.l4no,c.code))+'</b><p class="muted">'+esc(registerStatus(r,c))+'</p><label>Referensi temuan</label><select data-reg-finding><option value="">Tanpa referensi</option>'+findings.map(x=>'<option value="'+esc(x.id)+'"'+(selected===x.id?' selected':'')+'>'+esc(x.id+' — '+x.finding)+'</option>').join('')+(selected&&!findings.some(x=>x.id===selected)?'<option selected value="'+esc(selected)+'">'+esc(selected+' (temuan tidak tersedia)')+'</option>':'')+'</select><label>Alasan perubahan</label><textarea data-reg-reason rows="3">'+esc(r?.reason||'')+'</textarea><label>Expected benefit</label><textarea data-reg-benefit rows="3">'+esc(r?.benefit||'')+'</textarea><button class="sbtn ok" data-a="saveRegister">Simpan catatan perubahan</button><p class="muted" data-reg-status></p></div>';}
function handleRegisterSave(b,g){const el=b.closest('[data-register-code]');if(!el||!g)return;const c=compareProcess(g,state.changeCompareVersion!=='accepted').changes.find(x=>x.code===el.dataset.registerCode);if(!c)return toast('Perubahan tidak lagi aktif.');const reason=el.querySelector('[data-reg-reason]').value.trim();if(!reason)return toast('Isi alasan perubahan terlebih dahulu.');const r=normalizeRegister({l4:g.l4no,code:c.code,finding:el.querySelector('[data-reg-finding]').value,reason,benefit:el.querySelector('[data-reg-benefit]').value.trim(),signature:registerSignature(c)});if(saveRegister([...changeRegister.filter(x=>x.id!==r.id),r])){el.querySelector('[data-reg-status]').textContent='Catatan tersimpan.';el.querySelector('p.muted').textContent='Tersimpan';toast('Catatan perubahan tersimpan.');}}
const REGISTER_HEAD=['ID Perubahan','No L4','No L5','ID Temuan','Alasan Perubahan','Expected Benefit','Signature'];
function appendRegisterSheet(wb){const ws=XLSX.utils.aoa_to_sheet([REGISTER_HEAD,...changeRegister.map(r=>[r.id,r.l4,r.code,r.finding,r.reason,r.benefit,r.signature])]);ws['!cols']=[30,15,20,30,60,60,30].map(w=>({wch:w}));XLSX.utils.book_append_sheet(wb,ws,'Daftar Perubahan');}
function readRegisterSheet(wb){const ws=wb.Sheets['Daftar Perubahan'];if(!ws)return null;const a=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:''});if(!a.length||REGISTER_HEAD.some((h,i)=>a[0][i]!==h))throw Error('Format Daftar Perubahan tidak sesuai.');return a.slice(1).filter(r=>r[1]&&r[2]).map(r=>normalizeRegister({l4:r[1],code:r[2],finding:r[3],reason:r[4],benefit:r[5],signature:r[6]}));}
function changeSummaryHtml(g){
 try{const latest=state.changeCompareVersion!=='accepted',d=compareProcess(g,latest),filter=state.changeCompareFilter||'all', labels={added:'Ditambahkan',modified:'Diubah',removed:'Dihapus'};
 return '<div class="io"><div class="sec">Perbandingan AS-IS vs TO-BE</div><label>Versi TO-BE</label><select data-change-version><option value="latest"'+(latest?' selected':'')+'>Terbaru (termasuk usulan)</option><option value="accepted"'+(!latest?' selected':'')+'>Diterima</option></select><p class="muted">AS-IS: '+d.before+' aktivitas · TO-BE: '+d.after+' aktivitas</p><p>'+Object.keys(labels).map(k=>labels[k]+': <b>'+d.changes.filter(c=>c.type===k).length+'</b>').join(' · ')+'</p><label>Filter perubahan</label><select data-change-filter>'+[['all','Semua'],...Object.entries(labels)].map(([k,v])=>'<option value="'+k+'"'+(filter===k?' selected':'')+'>'+v+'</option>').join('')+'</select>'+ (d.changes.filter(c=>filter==='all'||c.type===filter).map(c=>'<details class="card"><summary style="cursor:pointer"><b>'+esc(c.code+' · '+c.name)+'</b><div class="muted">'+labels[c.type]+' · '+c.fields.length+' aspek</div></summary>'+c.fields.map(f=>'<div style="margin-top:8px"><b>'+esc(f.label)+'</b><div class="muted">AS-IS: '+esc(f.before||'—')+'</div><div>TO-BE: '+esc(f.after||'—')+'</div></div>').join('')+registerHtml(g,c)+'</details>').join('')||'<p class="muted">Tidak ada perubahan pada filter ini.</p>')+'<button class="sbtn" data-a="copyNetChanges">Salin daftar perubahan L4</button><div id="netChangeOut"></div><p class="muted">Perbandingan substansi; pengaturan tampilan tidak dihitung. Catatan dan pemetaan bersama tidak memiliki versi terpisah.</p></div>';
 }catch(e){return '<div class="io"><div class="sec">Perbandingan AS-IS vs TO-BE</div><p class="muted">Perbandingan tidak tersedia: '+esc(e.message)+'</p></div>';}
}

function renderElab(focusId){
  if(state.legacyAiPanel)return renderLegacyAi();
  const g = state.l4s.get(state.current); if(!g) return;
  const graph=draftGraph(g), absorbed=graph?consumedIds(JSON.parse(graph.text)):new Set();const archive=elabOf(g.l4no).filter(i=>absorbed.has(i.id));const items = elabOf(g.l4no).filter(i=>!absorbed.has(i.id)).slice().sort((a,b)=>cmpCode(a.after,b.after) || a.type.localeCompare(b.type));
  const P = items.filter(i=>i.status==='proposed'), A = items.filter(i=>i.status==='accepted'), X = items.filter(i=>i.status==='rejected');
  const busy = !!state.aiCtl;
  let h = '<button class="tb x" data-a="close">✕</button><div class="dcode" style="color:#5b21b6">RIWAYAT TO-BE</div><h3>'+esc(g.l4no+' '+g.l4)+'</h3>'+
    '<p class="muted">Edit nama, role dan alur langsung dari aktivitas. Panel ini untuk meninjau rancangan, dan persetujuan. Anotasi dikelola dari Detail Aktivitas. Export dapat memilih rancangan terbaru atau yang diterima.</p>';
  h += changeSummaryHtml(g);
  h += '<label class="chk"><input type="checkbox" data-a="toggleProp"'+(state.showProposed?' checked':'')+'> Tampilkan usulan di flowchart</label>';
  h += '<div class="sec">Usulan ('+P.length+')<span class="sp"></span>'+(P.length?'<button class="sbtn ok" data-a="acceptAll">Terima semua</button><button class="sbtn no" data-a="rejectAll">Tolak semua</button>':'')+'</div>';
  h += P.length ? P.map(i=>cardHtml(i,g)).join('') : '<p class="muted">Belum ada usulan.</p>';
  h += '<div class="sec">Diterima ('+A.length+')</div>';
  h += A.length ? A.map(i=>cardHtml(i,g)).join('') : '<p class="muted">Belum ada elemen yang diterima.</p>';
  if(X.length) h += '<details'+(X.some(i=>i.id===state.editing||i.id===focusId)?' open':'')+'><summary class="sec" style="cursor:pointer">Ditolak ('+X.length+')</summary>'+X.map(i=>cardHtml(i,g)).join('')+'</details>';
  const aiOk = state.sampleAvail !== false;
  if(archive.length)h+='<details><summary class="sec">Riwayat usulan yang digabung ('+archive.length+')</summary>'+archive.map(i=>'<p class="muted">'+esc(TYPE_NAME[i.type]+' · '+i.after+' · '+(i.text||i.role||'')+' · '+STATUS_NAME[i.status])+'</p>').join('')+'</details>';
  h += '<div class="io"><div class="sec">Ringkasan perubahan</div><button class="sbtn" data-a="copyChanges">Salin ringkasan perubahan</button><div id="chgOut"></div></div>';
  $('detail').innerHTML = h;
  if(focusId){ const c = $('detail').querySelector('.card[data-id="'+focusId+'"]'); if(c){ c.scrollIntoView({block:'center'}); c.classList.add('flash'); setTimeout(()=>c.classList.remove('flash'), 1400); } }
}
function setAiMsg(text){ state.aiMsg = {l4:state.current, text}; const s = $('aiStatus'); if(s) s.textContent = text; }

function onSideClick(e){
  const b = e.target.closest('[data-a]'); if(!b) return;
  const a = b.dataset.a; const card = b.closest('.card'); const id = card && card.dataset.id;
  const g = state.l4s.get(state.current); const list = g ? elabOf(g.l4no) : [];
  const it = id ? list.find(x=>x.id===id) : null;
  if(a==='close'){ $('detail').classList.add('hidden'); state.sideMode = null; ['btnElab','btnX','btnR'].forEach(x=>$(x).classList.remove('on')); return; }
  if(a==='annotationDetail'){if(it){$('processEditMode').value='tobe';redraw();directEditor(it.after);}return;}
  if(a.startsWith('ann')){handleAnnotation(a,b,g);return;}
  if(onRoleClick(a, b)) return;
  if(a==='rnEdit' || a==='rnSave' || a==='rnCancel' || a==='rnReset'){ const code = $('detail').dataset.code;
    if(a==='rnEdit') return showDetail(code, true); if(a==='rnCancel') return showDetail(code);
    return setRename(code, a==='rnReset' ? '' : $('rnInput').value); }
  if(a==='saveRegister') return handleRegisterSave(b,g);
  if(a==='openElab') return openElab();
  if(a==='toggleProp'){ state.showProposed = b.checked; store.set('flowbpa:showprop', b.checked?'1':'0'); redraw(); return; }
  if(a==='aiSettings') return openAiSettings();
  if(a==='ai') return proposeAI();
  if(a==='free') return freeAI();
  if(a==='stopAi'){ if(state.aiCtl) state.aiCtl.abort(); return; }
  if(a==='acceptAll'){ const graph=draftGraph(g), ids=graph?consumedIds(JSON.parse(graph.text)):new Set();list.forEach(x=>{ if(x.status==='proposed'&&!ids.has(x.id)) x.status='accepted'; }); return elabChanged(); }
  if(a==='rejectAll'){ const graph=draftGraph(g), ids=graph?consumedIds(JSON.parse(graph.text)):new Set();list.forEach(x=>{ if(x.status==='proposed'&&!ids.has(x.id)) x.status='rejected'; }); return elabChanged(); }
  if(a==='add'){
    const acts = actsOf(g); const n = {id:newId(), l4:g.l4no, type:'gate', after:acts[0].l5no, text:'', noTarget:acts[0].l5no, noLabel:'Tidak', label:'Hold', status:'accepted', source:'manual', reason:''};
    list.push(n); state.editing = n.id; state.newItem = n.id; saveElab(); renderElab(n.id); return;
  }
  if(a==='exportElab') return exportElabSheet();
  if(a==='copyNetChanges'){try{const d=compareProcess(g,state.changeCompareVersion!=='accepted');const text=g.l4no+' '+g.l4+' — AS-IS vs TO-BE '+(state.changeCompareVersion==='accepted'?'diterima':'terbaru')+'\n'+d.changes.map(c=>c.code+' '+c.type+' '+c.name+'\n'+c.fields.map(f=>'  '+f.label+': '+(f.before||'—')+' → '+(f.after||'—')).join('\n')).join('\n');$('netChangeOut').innerHTML='<textarea readonly style="width:100%;min-height:100px">'+esc(text)+'</textarea>';if(navigator.clipboard)navigator.clipboard.writeText(text).then(()=>toast('Daftar perubahan disalin.'),()=>toast('Salin manual dari kotak.'));}catch(e){toast(e.message);}return;}
  if(a==='copyChanges'){
    const lines = [];
    for(const k of Object.keys(state.elab).sort(cmpCode)){
      const ch = state.elab[k].filter(i=>i.status==='accepted' && STRUCT.includes(i.type)); if(!ch.length) continue;
      const gg = state.l4s.get(k); lines.push(k+' '+(gg?gg.l4:''));
      ch.forEach(i=>lines.push('  - '+changeText(i)+(i.reason?' (dasar: '+i.reason+')':'')));
    }
    const txt2 = lines.length ? 'Perubahan terhadap workbook BPA ('+state.fileName+')\n'+lines.join('\n') : '';
    if(!txt2){ toast('Belum ada perubahan yang diterima.'); return; }
    const show = () => { $('chgOut').innerHTML = '<textarea readonly style="width:100%;min-height:120px;margin-top:6px;font-size:12px;border:1px solid var(--line);border-radius:6px;padding:6px">'+esc(txt2)+'</textarea>'; };
    try{ navigator.clipboard.writeText(txt2).then(()=>{ toast('Ringkasan perubahan disalin.'); show(); }, ()=>{ show(); toast('Salin manual dari kotak di bawah.'); }); }catch(err){ show(); }
    return;
  }
  if(a==='importElab') return $('elabIn').click();
  if(!it) return;
  if(a==='accept'){ it.status='accepted'; return elabChanged(); }
  if(a==='reject'){ it.status='rejected'; return elabChanged(); }
  if(a==='unaccept' || a==='restore'){ it.status='proposed'; return elabChanged(); }
  if(a==='del'){ list.splice(list.indexOf(it),1); return elabChanged(); }
  if(a==='edit'){ state.editing = it.id; return renderElab(it.id); }
  if(a==='cancel'){
    if(state.newItem===it.id) list.splice(list.indexOf(it),1);
    state.editing = null; state.newItem = null; saveElab(); return renderElab();
  }
  if(a==='save'){
    const v = f => { const el = card.querySelector('[data-f="'+f+'"]'); return el ? el.value.trim() : undefined; };
    it.type = v('type'); it.after = v('after'); it.text = v('text') || '';
    if(v('subject')!==undefined) it.subject = v('subject');
    if(v('role')!==undefined) it.role = matchRole(v('role'));
    if(it.type==='reorder'){ if(it.subject===it.after){ toast('Aktivitas tujuan harus berbeda.'); return; } it.text = 'Pindah '+it.subject+' ke setelah '+it.after; }
    if(it.type==='role'){ if(!it.role){ toast('PIC baru belum diisi.'); return; } it.text = 'PIC '+it.after+' → '+roleInfo(it.role).name; }
    if(it.type==='declane'){ if(!it.role){ toast('Kolom tujuan belum diisi.'); return; } it.text = 'Keputusan setelah '+it.after+' → kolom '+roleInfo(it.role).name; }
    if(v('noTarget')!==undefined) it.noTarget = v('noTarget');
    if(v('noLabel')!==undefined) it.noLabel = v('noLabel');
    if(v('label')!==undefined) it.label = v('label');
    if(!it.text){ toast('Teks belum diisi.'); return; }
    if(it.type==='gate' && !/\?$/.test(it.text)) it.text += '?';
    if(it.type!=='gate'){ delete it.noTarget; delete it.noLabel; }
    if(it.type!=='hold') delete it.label;
    if(it.type!=='reorder') delete it.subject;
    if(!['newact','role','gate','declane'].includes(it.type) || !it.role) delete it.role;
    if(it.source==='ai' && state.newItem!==it.id) it.edited = true;
    state.editing = null; state.newItem = null; return elabChanged();
  }
}
function onSideChange(e){
  if(e.target.matches("[data-change-version]")){state.changeCompareVersion=e.target.value;return renderElab();}
  if(e.target.matches("[data-change-filter]")){state.changeCompareFilter=e.target.value;return renderElab();}
  // switching type in the edit form re-renders the form with the right fields
  const sel = e.target.closest('[data-f="type"]'); if(!sel) return;
  const card = sel.closest('.card'); const g = state.l4s.get(state.current);
  const it = elabOf(g.l4no).find(x=>x.id===card.dataset.id); if(!it) return;
  const keep = {}; card.querySelectorAll('[data-f]').forEach(el=>keep[el.dataset.f]=el.value);
  Object.assign(it, {after:keep.after, text:keep.text||'', type:sel.value});
  if(it.type==='reorder' && !it.subject) it.subject = it.after;
  if(it.type==='gate' && !it.noTarget) it.noTarget = it.after;
  renderElab(it.id);
}

/* ---- AI proposals ---- */
let _sample;
const AI_MODELS = [['claude-sonnet-5-5','Claude Sonnet 5.5 (disarankan)'],['claude-haiku-5-5','Claude Haiku 5.5 (lebih hemat)'],['claude-opus-5-5','Claude Opus 5.5 (paling teliti)']];
function apiKey(){ return store.get('flowbpa:apikey') || ''; }
function apiModel(){ return store.get('flowbpa:apimodel') || AI_MODELS[0][0]; }
function anthropicSample(){
  const call = async (input, opts) => {
    opts = opts || {};
    let res;
    try{
      res = await fetch('https://api.anthropic.com/v1/messages', { method:'POST', signal: opts.signal,
        headers:{ 'content-type':'application/json', 'x-api-key': apiKey(), 'anthropic-version':'2023-06-01', 'anthropic-dangerous-direct-browser-access':'true' },
        body: JSON.stringify({ model: apiModel(), max_tokens: 4096, messages:[{role:'user', content: typeof input==='string' ? input : String(input)}] }) });
    }catch(e){ throw {code: e && e.name==='AbortError' ? 'cancelled' : 'upstream_error', message:String(e)}; }
    if(!res.ok){
      let msg = ''; try{ const j = await res.json(); msg = j && j.error && j.error.message || ''; }catch(e){}
      throw {code: res.status===401 || res.status===403 ? 'not_granted' : res.status===429 || res.status===529 ? 'rate_limited' : res.status===400 ? 'invalid_request' : 'upstream_error', message:msg, status:res.status};
    }
    const j = await res.json();
    const text = (j.content||[]).filter(c=>c.type==='text').map(c=>c.text).join('').trim();
    if(!text) throw {code:'empty_completion', message:'Jawaban kosong'};
    return {text, truncated: j.stop_reason==='max_tokens'};
  };
  const fn = (input, opts) => call(input, opts);
  fn.json = async (input, opts) => {
    const {text} = await call(input, opts);
    const tryParse = s => { try{ return JSON.parse(s); }catch(e){ return undefined; } };
    let v = tryParse(text);
    if(v===undefined){ const m = text.match(/```(?:json)?\s*([\s\S]*?)```/); if(m) v = tryParse(m[1]); }
    if(v===undefined){ const a = text.search(/[\[{]/), b = Math.max(text.lastIndexOf('}'), text.lastIndexOf(']')); if(a>=0 && b>a) v = tryParse(text.slice(a, b+1)); }
    if(v===undefined) throw {code:'invalid_json', message:'Jawaban bukan JSON', text};
    return v;
  };
  return fn;
}
async function getSample(){
  if(_sample !== undefined) return _sample;
  _sample = apiKey() ? anthropicSample() : null;
  state.sampleAvail = !!_sample;
  return _sample;
}
function buildPrompt(g, items){
  const acts = actsOf(g);
  const line = a => '- '+a.l5no+(hasFormalDecision(a)?' [KEPUTUSAN RESMI]':'')+' | '+a.l5+' | PIC: '+(a.role||'-')+' | Sebelum: '+(a.prev||'-')+' | Sesudah: '+(a.next||'-')+
    ' | Input dari: '+(a.inParty||'-')+' | Output ke: '+(a.outParty||'-')+(a.dec?' | Decision Ref: '+a.dec:'')+'\n  Deskripsi: '+(a.desc||'(kosong)').slice(0,1200);
  const show = i => '- '+i.type+' setelah '+i.after+': '+i.text+(i.type==='gate'?' (tidak lolos → '+i.noTarget+')':'');
  const live = items.filter(i=>i.status!=='rejected'), rej = items.filter(i=>i.status==='rejected');
  return [
    'Anda membantu konsultan HR mengelaborasi flowchart SOP yang dibuat dari Business Process Architecture (BPA).',
    'Subproses: '+g.l4no+' '+g.l4+' (SOP '+g.l3no+' '+g.l3+'; '+g.l2+').',
    'Alur resmi dari kolom "Proses Sesudah" SUDAH digambar dan tidak boleh diubah. Tugas Anda: usulkan elemen elaborasi yang BENAR-BENAR tersirat di Activity Description agar flowchart lebih jelas.',
    '',
    'Jenis elemen:',
    '1. "gate": gateway pemeriksaan (belah ketupat) tepat setelah aktivitas yang memeriksa/memverifikasi/memvalidasi/mereviu, bila deskripsinya menyiratkan hasil yang tidak lolos harus dikembalikan, diklarifikasi, atau diperbaiki. Isi "text" = pertanyaan singkat diakhiri "?" (maks 5 kata, mis. "Data lengkap?"), "no_target" = kode aktivitas di subproses ini yang dituju bila tidak lolos, "no_label" = label singkat cabang tidak lolos (mis. "Tidak: klarifikasi"). Cabang "Ya" otomatis mengikuti alur resmi. JANGAN usulkan gate pada aktivitas bertanda [KEPUTUSAN RESMI].',
    '2. "hold": status HOLD/tertunda bila deskripsi menyebut kondisi belum dapat dipenuhi, ditunda, atau menunggu. Isi "text" (maks 8 kata) dan "label" (label cabang, mis. "Hold").',
    '3. "doc": dokumen/bukti keluaran yang disebut eksplisit di deskripsi (mis. "Catatan permintaan rekrutmen", "SK penetapan"). Isi "text" maks 5 kata. Maks 1 per aktivitas.',
    '4. "datain": input data dari luar prosedur yang disebut eksplisit (mis. formasi, anggaran, kebijakan, hasil prosedur lain). Isi "text" maks 8 kata.',
    '',
    'Aturan:',
    '- Hanya berdasarkan teks deskripsi; jangan mengarang langkah, pihak, atau dokumen yang tidak disebut. Lebih baik sedikit tapi tepat (umumnya 0–8 elemen untuk satu subproses).',
    '- "after" wajib salah satu kode aktivitas di daftar.',
    '- Sertakan "reason": satu kalimat singkat yang mengutip atau merujuk kata di deskripsi.',
    '- Jangan mengulang elemen yang sudah ada atau yang sudah ditolak.',
    '- Bahasa Indonesia baku, singkat.',
    '',
    'Balas HANYA dengan JSON berbentuk:',
    '{"items":[{"type":"gate","after":"03.01.03","text":"Data lengkap?","no_target":"03.01.05","no_label":"Tidak: klarifikasi","reason":"..."},{"type":"doc","after":"03.01.02","text":"Catatan permintaan rekrutmen","reason":"..."}]}',
    'Bila tidak ada yang layak, balas {"items":[]}.',
    '',
    'Aktivitas:',
    acts.map(line).join('\n'),
    '',
    'Elemen yang sudah ada:', live.length ? live.map(show).join('\n') : '(belum ada)',
    'Elemen yang sudah ditolak:', rej.length ? rej.map(show).join('\n') : '(belum ada)'
  ].join('\n');
}
function allRoles(){ const r = new Set(); for(const a of state.rows){ if(a.role) r.add(a.role); if(a.decisionRole)r.add(a.decisionRole); splitParty(a.inParty).forEach(x=>r.add(x)); splitParty(a.outParty).forEach(x=>r.add(x)); } Object.values(state.roleEdit.added).forEach(l=>l.forEach(x=>r.add(x))); return sortRoles([...r]); }
function matchRole(s){
  s = String(s||'').trim(); if(!s) return '';
  const all = allRoles(), n = norm(s);
  return all.find(r=>norm(r)===n) || all.find(r=>roleInfo(r).code.toLowerCase()===n) || all.find(r=>norm(roleInfo(r).name)===n) ||
         all.find(r=>norm(r).includes(n) || n.includes(norm(roleInfo(r).name))) || s;
}
function canMove(g, code){ const a = g.acts.find(x=>x.l5no===code); if(!a) return false; const f = parseFlow(a.next); return f.length===1 && !f[0].label && (f[0].kind==='code'||f[0].kind==='end'); }
function validateProposals(g, raw, source){
  const acts = actsOf(g); const codes = new Set(acts.map(a=>a.l5no));
  const existing = elabOf(g.l4no);
  const arr = Array.isArray(raw) ? raw : (raw && Array.isArray(raw.items) ? raw.items : []);
  const out = []; let dropped = 0;
  const clip = (s, n) => String(s==null?'':s).replace(/\s+/g,' ').trim().slice(0, n);
  for(const r of arr){
    const type = String(r && r.type || '').toLowerCase();
    const after = clip(r && r.after, 20);
    if(!TYPE_NAME[type] || !codes.has(after) || (STRUCT.includes(type) && source!=='free')){ dropped++; continue; }
    const it = {id:newId(), l4:g.l4no, type, after, text:clip(r.text, 90), status:'proposed', source:source||'ai', reason:clip(r.reason, 240)};
    if(type==='rename' && !it.text){ dropped++; continue; }
    if(type==='reorder'){
      it.subject = clip(r.subject, 20);
      if(!codes.has(it.subject) || it.subject===after || !canMove(g, it.subject)){ dropped++; continue; }
      it.text = 'Pindah '+it.subject+' ke setelah '+after;
    }
    if(type==='role'){
      it.role = matchRole(r.role); const cur = acts.find(a=>a.l5no===after).role;
      if(!it.role || it.role===cur){ dropped++; continue; }
      it.text = 'PIC '+after+' → '+roleInfo(it.role).name;
    }
    if(type==='newact') it.role = matchRole(r.role);
    if(type==='gate' && r.role && source==='free') it.role = matchRole(r.role);
    if(type==='declane'){
      if(source!=='free'){ dropped++; continue; }
      it.role = matchRole(r.role); const a0 = acts.find(a=>a.l5no===after);
      if(!it.role || !hasFormalDecision(a0)){
        // a gateway proposed/accepted for that activity: move the gateway instead
        const gw = [...existing, ...out].find(x=>x.type==='gate' && x.after===after && x.status!=='rejected');
        if(gw && it.role){ gw.role = it.role; continue; }
        dropped++; continue;
      }
      it.text = 'Keputusan setelah '+after+' → kolom '+roleInfo(it.role).name;
    }
    if(!it.text){ dropped++; continue; }
    if(type==='gate'){
      it.noTarget = clip(r.no_target, 20); it.noLabel = clip(r.no_label, 40) || 'Tidak';
      if(!codes.has(it.noTarget) || hasFormalDecision(acts.find(a=>a.l5no===after))){ dropped++; continue; }
      if(!/\?$/.test(it.text)) it.text += '?';
    }
    if(type==='hold') it.label = clip(r.label, 30) || 'Hold';
    const dup = [...existing, ...out].some(x=> x.status!=='rejected' && x.type===type && (type==='reorder' ? x.subject===it.subject : x.after===after && (['gate','doc','role','declane'].includes(type) || norm(x.text)===norm(it.text))));
    if(dup){ dropped++; continue; }
    out.push(it);
  }
  return {items:out, dropped};
}
async function proposeAI(){
  const g = state.l4s.get(state.current); if(!g || state.aiCtl) return;
  const epoch=projectEpoch; const sample = await getSample();
  if(epoch!==projectEpoch)return;
  if(!sample){ renderElab(); return; }
  const ctl = new AbortController(); state.aiCtl = ctl; const l4no = g.l4no;
  renderElab(); setAiMsg('Claude sedang membaca '+g.acts.length+' deskripsi aktivitas… biasanya 20–60 detik.');
  try{
    const raw = await sample.json(buildPrompt(g, elabOf(l4no)), {modelTier:'default', signal:ctl.signal, cache:false});
    if(epoch!==projectEpoch)return;
    const {items, dropped} = validateProposals(g, raw);
    elabOf(l4no).push(...items);
    state.aiMsg = {l4:l4no, text: items.length ? items.length+' usulan baru. Periksa satu per satu, lalu terima atau tolak.'+(dropped?' ('+dropped+' usulan dibuang karena tidak sesuai aturan.)':'') : 'Tidak ada usulan baru untuk subproses ini.'+(dropped?' ('+dropped+' usulan dibuang karena tidak sesuai aturan.)':'')};
    state.showProposed = true;
  }catch(e){
    if(epoch!==projectEpoch)return;
    const c = e && e.code;
    const msg = c==='cancelled' ? 'Dihentikan.' :
      (c==='not_granted') ? 'API key ditolak (tidak valid atau tidak punya akses). Periksa di ⚙ Pengaturan AI.' :
      (c==='invalid_request') ? 'Permintaan ditolak API (periksa model di ⚙ Pengaturan AI).' :
      c==='rate_limited' ? 'Terlalu banyak permintaan atau batas pemakaian tercapai. Coba lagi beberapa saat lagi.' :
      c==='session_expired' ? 'Sesi berakhir. Silakan masuk lagi ke Claude.' :
      c==='invalid_json' || c==='empty_completion' ? 'Jawaban Claude tidak bisa dibaca. Coba sekali lagi.' :
      c==='refused' ? 'Claude tidak bisa memproses subproses ini.' :
      'Gagal terhubung ke Claude. Coba lagi.';
    state.aiMsg = {l4:l4no, text:msg};
  }finally{
    if(epoch!==projectEpoch)return;
    state.aiCtl = null; saveElab();
    if(state.current===l4no){ redraw(); renderTree(); if(state.sideMode==='elab') renderElab(); } else renderTree();
  }
}

function buildFreePrompt(g, text){
  const acts = actsOf(g);
  return [
    'Ubah instruksi konsultan HR di bawah menjadi elemen elaborasi flowchart SOP. Jangan menambah elemen yang tidak diminta.',
    'Subproses: '+g.l4no+' '+g.l4+'.',
    '',
    'Jenis elemen dan isiannya:',
    '- "gate": gateway pemeriksaan setelah satu aktivitas. "text" = pertanyaan singkat diakhiri "?", "no_target" = kode aktivitas tujuan bila tidak lolos, "no_label" = label cabang tidak lolos (default "Tidak"). Tidak boleh pada aktivitas bertanda [KEPUTUSAN RESMI].',
    '- "hold": status HOLD/tertunda dari satu aktivitas. "text" singkat, "label" cabang (default "Hold").',
    '- "doc": dokumen/bukti keluaran satu aktivitas. "text" = nama dokumen singkat.',
    '- "datain": input data dari luar prosedur ke satu aktivitas. "text" singkat.',
    '- "newact": proses/aktivitas BARU yang disisipkan tepat setelah aktivitas "after" (alur lama "after" diteruskan dari proses baru). "text" = nama proses diawali kata kerja (gaya nama aktivitas di daftar), "role" = PIC dari daftar peran bila disebut (kosongkan bila tidak).',
    '- "reorder": memindah aktivitas "subject" ke posisi tepat setelah aktivitas "after".',
    '- "role": mengganti PIC aktivitas "after" menjadi "role" (pilih dari daftar peran).',
    '- "rename": mengganti NAMA aktivitas "after" menjadi "text" (nama baru diawali kata kerja, gaya nama aktivitas di daftar).',
    '- "declane": memindah belah ketupat KEPUTUSAN RESMI setelah aktivitas "after" ke kolom peran "role" (mis. ke kolom Pemberi Persetujuan). Untuk gateway elaborasi, isi saja "role" pada elemen "gate" sebagai kolomnya.',
    '"after" (dan "subject") wajib kode aktivitas dari daftar. Bila konsultan menyebut nama aktivitas, cocokkan ke kodenya. Rapikan teks seperlunya, pertahankan maksudnya.',
    'Isi "reason" dengan potongan instruksi yang menjadi dasar elemen itu.',
    'Bila ada bagian instruksi yang tidak bisa dipetakan, jelaskan singkat di "unclear".',
    '',
    'Balas HANYA JSON: {"items":[{"type":"gate","after":"03.01.03","text":"Data lengkap?","no_target":"03.01.05","no_label":"Tidak: klarifikasi","reason":"..."},{"type":"newact","after":"03.01.08","text":"Mengarsipkan Dokumen Permintaan","role":"PF-03 Pelaksana HC","reason":"..."},{"type":"reorder","subject":"03.01.05","after":"03.01.03","reason":"..."},{"type":"role","after":"03.01.04","role":"PF-03 Pelaksana HC","reason":"..."},{"type":"declane","after":"03.01.06","role":"PF-08 Pemberi Persetujuan","reason":"..."}],"unclear":""}',
    '',
    'Aktivitas:',
    acts.map(a=>'- '+a.l5no+(hasFormalDecision(a)?' [KEPUTUSAN RESMI]':'')+' '+a.l5+' (PIC: '+(a.role||'-')+')').join('\n'),
    '',
    'Daftar peran:', allRoles().join('; '),
    '',
    'Instruksi konsultan:',
    text.slice(0, 4000)
  ].join('\n');
}
async function freeAI(){
  const g = state.l4s.get(state.current); if(!g || state.aiCtl) return;
  const text = (state.freeText||'').trim();
  if(!text){ toast('Tulis dulu apa yang mau ditambahkan.'); return; }
  const epoch=projectEpoch; const sample = await getSample();
  if(epoch!==projectEpoch)return;
  if(!sample){ renderElab(); return; }
  const ctl = new AbortController(); state.aiCtl = ctl; const l4no = g.l4no;
  renderElab(); setAiMsg('Claude sedang menerjemahkan teks lw…');
  try{
    const raw = await sample.json(buildFreePrompt(g, text), {modelTier:'default', signal:ctl.signal, cache:false});
    if(epoch!==projectEpoch)return;
    const {items, dropped} = validateProposals(g, raw, 'free');
    elabOf(l4no).push(...items);
    const unclear = raw && typeof raw.unclear==='string' && raw.unclear.trim() ? ' Belum terpetakan: '+raw.unclear.trim() : '';
    if(items.length && !unclear) state.freeText = '';
    const bad = (raw && Array.isArray(raw.items) ? raw.items : []).filter(r=>String(r&&r.type).toLowerCase()==='reorder' && !canMove(g, String(r.subject||'')));
    state.aiMsg = {l4:l4no, text: (items.length ? items.length+' elemen ditambahkan ke Usulan. Cek lalu terima.' : 'Tidak ada elemen yang bisa dibuat dari teks itu.')+(dropped?' ('+dropped+' dibuang: kode tidak ditemukan, sudah ada, aktivitas sudah punya keputusan resmi, atau aktivitas bercabang tidak bisa dipindah.)':'')+unclear};
    state.showProposed = true;
  }catch(e){
    if(epoch!==projectEpoch)return;
    const c = e && e.code;
    state.aiMsg = {l4:l4no, text: c==='cancelled' ? 'Dihentikan.' :
      (c==='not_granted') ? 'API key ditolak. Periksa di ⚙ Pengaturan AI.' :
      c==='rate_limited' ? 'Batas pemakaian tercapai. Coba lagi beberapa saat lagi.' :
      c==='invalid_json' || c==='empty_completion' ? 'Jawaban Claude tidak bisa dibaca. Coba sekali lagi.' : 'Gagal terhubung ke Claude. Coba lagi.'};
  }finally{
    if(epoch!==projectEpoch)return;
    state.aiCtl = null; saveElab();
    if(state.current===l4no){ redraw(); renderTree(); if(state.sideMode==='elab') renderElab(); } else renderTree();
  }
}

/* ================= Jabatan (dari sheet "Pemetaan Jabatan" / "Pemetaan DEC") ================= */
const sepPos = s => String(s||'').split(/\s*\/\s*|\n/).map(x=>x.trim()).filter(Boolean);
function rolesInL4(g){
  const s = new Set();
  g.acts.forEach(a=>{ if(a.role) s.add(a.role); if(a.decisionRole)s.add(a.decisionRole); splitParty(a.inParty).forEach(p=>s.add(p)); splitParty(a.outParty).forEach(p=>s.add(p)); });
  return sortRoles([...s]);
}
function laneMode(l4no){ const m=state.laneModes[l4no]; return m==='unit'?'unit':'position'; }
function unitFor(l4no,role,act){ return (act && act.role===role && act.unitX) || (state.unitmap[l4no]||{})[role] || ''; }
function saveLaneSettings(){ store.set('flowbpa:unitmap',JSON.stringify(state.unitmap)); store.set('flowbpa:lanemodes',JSON.stringify(state.laneModes)); }
function posFor(l4no, role, dec, act){
  if(laneMode(l4no)!=='position') return '';
  if(act && act.posX && role===act.role) return sepPos(act.posX).join(' / ');
  if(roleInfo(role).code==='PF-08' && dec && state.decmap[dec]) return state.decmap[dec];
  const m = state.posmap[l4no]; return m ? (m[role]||'') : '';
}
function laneShort(pos){ const p = sepPos(pos); return p.length > 3 ? p.slice(0,3).join(' / ')+' +'+(p.length-3)+' lainnya' : p.join(' / '); }
function savePos(){ store.set('flowbpa:posmap', JSON.stringify(state.posmap)); store.set('flowbpa:decmap', JSON.stringify(state.decmap)); store.set('flowbpa:poson', state.posOn?'1':'0'); }
const hasLaneSettings = () => Object.keys(state.unitmap).length || Object.keys(state.laneModes).length;
const hasPos = () => hasLaneSettings() || Object.keys(state.posmap).length > 0 || Object.keys(state.decmap).length > 0 || state.rows.some(r=>r.posX);
function parsePosSheets(wb){
  const find = re => wb.SheetNames.find(n=>re.test(n));
  const sm = find(/peran jabatan/i) || find(/pemetaan jabatan/i), sd = find(/pemetaan dec/i); if(!sm && !sd && !find(/^pengaturan swimlane$/i)) return null;
  const rowsOf = sn => { const a = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, raw:false, defval:''}); const hi = a.findIndex(r=>r.map(norm).some(h=>h==='jabatan')); if(hi<0) return []; const h = a[hi].map(norm);
    return a.slice(hi+1).map(r=>{ const o = {}; h.forEach((k,i)=>o[k]=String(r[i]==null?'':r[i]).trim()); return o; }); };
  const out = {posmap:{}, decmap:{}, unitmap:{}, laneModes:{}};
  if(sm) for(const r of rowsOf(sm)){ const l4 = r['no l4'], role = r['peran'], pos = sepPos(r['jabatan']).join(' / '); if(!l4 || !role) continue; (out.posmap[l4] = out.posmap[l4] || {}); if(pos) out.posmap[l4][role] = pos; const unit=r['unit kerja']||r['unit organisasi']||r['divisi']; if(unit) (out.unitmap[l4] ||= {})[role]=unit; }
  if(sd) for(const r of rowsOf(sd)){ const d = r['kode dec'], pos = sepPos(r['jabatan']).join(' / '); if(d && pos) out.decmap[d] = pos; }
  const modes=find(/^pengaturan swimlane$/i); if(modes){ const a=XLSX.utils.sheet_to_json(wb.Sheets[modes],{header:1,raw:false,defval:''}); for(const r of a.slice(1)) if(r[0] && ['role','unit','position'].includes(r[1])) out.laneModes[r[0]]=r[1]; }
  return out;
}
function applyPos(p){
  if(!p) return 0;
  state.posmap = p.posmap; state.decmap = p.decmap; state.unitmap=p.unitmap||{}; state.laneModes=p.laneModes||{}; saveLaneSettings(); state.posOn = true; savePos(); updatePosToggle();
  return Object.values(p.posmap).reduce((n,m)=>n+Object.keys(m).length,0) + Object.keys(p.decmap).length;
}
function updatePosToggle(){ $('posTg').style.display='none'; $('posOn').checked = !!state.posOn; $('laneMode').value=laneMode(state.current); }

/* ---- Sheet "Elaborasi Flowchart" ---- */
const ELAB_HEAD = ['ID','No L4','Subproses','Jenis','Setelah Aktivitas','Teks','Label Cabang','Tujuan Tidak Lolos','Status','Sumber','Alasan','Aktivitas Dipindah','PIC'];
function elabRowsAll(){
  const rows = [];
  for(const k of Object.keys(state.elab).sort(cmpCode)) for(const i of state.elab[k].slice().sort((a,b)=>cmpCode(a.after,b.after))){
    const g = state.l4s.get(k);
    rows.push([i.id, k, g?g.l4:'', TYPE_NAME[i.type], i.after, i.text, i.type==='gate'?(i.noLabel||''):i.type==='hold'?(i.label||''):'', i.type==='gate'?(i.noTarget||''):'', STATUS_NAME[i.status]||i.status, i.source==='ai'?'AI':i.source==='free'?'Teks bebas':'Manual', i.reason||'', i.subject||'', i.role||'']);
  }
  return rows;
}
async function exportElabSheet(){
  const rows = elabRowsAll();
  if(!state.rows.length && !rows.length && !hasPos() && !hasExtra() && !hasRoleEdits() && !Object.keys(state.laneModes).length && !state.rows.some(r=>r._asisEdited)){ toast('Belum ada elaborasi, SLA/output, atau pemetaan untuk disimpan.'); return 'empty'; }
  await loadScript(LIBS.xlsx);
  const wb = XLSX.utils.book_new();
  if(state.rows.length){ const keys=['l1no','l1','l2no','l2','l3no','l3','l4no','l4','l5no','l5','desc','prev','next','role','posX','unitX','decisionQuestion','decisionRole','inputX','outputX','docX','slaX','systemX']; const heads=['No L1','Level 1','No L2','Level 2','No L3','Level 3','No L4','Level 4','No L5','Level 5','Activity Description','Proses Sebelum','Proses Sesudah','Role','Jabatan','Unit','Pertanyaan Keputusan','Role Keputusan','Input','Output','Dokumen','SLA','Sistem/Aplikasi']; const legacy=Object.keys(COLS).filter(k=>!keys.includes(k));const all=[...keys,...legacy]; const ws=XLSX.utils.aoa_to_sheet([[...heads,...legacy.map(k=>COLS[k][0])],...state.rows.map(r=>all.map(k=>r[k]||''))]);ws['!cols']=all.map((k,i)=>({wch:i%2===0?18:35,hidden:i>=keys.length}));ws['!autofilter']={ref:'A1:W'+(state.rows.length+1)}; XLSX.utils.book_append_sheet(wb,ws,'SOP'); }
  if(rows.length){
    const ws = XLSX.utils.aoa_to_sheet([ELAB_HEAD, ...rows]);
    ws['!cols'] = [8,8,34,13,12,40,22,14,10,10,60,14,26].map(w=>({wch:w}));
    ws['!autofilter'] = {ref: ws['!ref']};
    XLSX.utils.book_append_sheet(wb, ws, 'Elaborasi Flowchart');
  }
  if(hasExtra()){
    const X = [['No L4','No L5','Aktivitas','SLA','Dokumen/Output','Catatan','Input/Output Tambahan']];
    for(const c of Object.keys(state.extra).sort(cmpCode)){ const r = state.byCode.get(c), x = state.extra[c]; X.push([r?r.l4no:c.split('.').slice(0,2).join('.'), c, r?r.l5:'', x.sla||'', x.out||'', x.note||'', encodeDocs(x.docs)]); }
    const wx = XLSX.utils.aoa_to_sheet(X); wx['!cols'] = [8,11,46,20,44,50,60].map(w=>({wch:w})); XLSX.utils.book_append_sheet(wb, wx, 'SLA & Output');
  }
  if(hasRoleEdits()){ const wr = XLSX.utils.aoa_to_sheet(roleSheet()); wr['!cols'] = [14,8,30,40].map(w=>({wch:w})); XLSX.utils.book_append_sheet(wb, wr, 'Pengaturan Lane'); }
  if(hasPos()){
    const M = [['No L4','Subproses','Peran','Jabatan','Unit Kerja']];
    for(const k of uniq([...Object.keys(state.posmap),...Object.keys(state.unitmap)]).sort(cmpCode)){ const g=state.l4s.get(k), p=state.posmap[k]||{}, u=state.unitmap[k]||{}; for(const role of uniq([...Object.keys(p),...Object.keys(u)])) M.push([k,g?g.l4:'',role,p[role]||'',u[role]||'']); }
    const D = [['Kode DEC','Jabatan']]; for(const [d,pos] of Object.entries(state.decmap).sort()) D.push([d, pos]);
    const wm = XLSX.utils.aoa_to_sheet(M); wm['!cols'] = [8,38,26,60].map(w=>({wch:w}));
    const wd = XLSX.utils.aoa_to_sheet(D); wd['!cols'] = [16,60].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, wm, 'Peran Jabatan'); if(Object.keys(state.decmap).length) XLSX.utils.book_append_sheet(wb, wd, 'Pemetaan DEC');
  }
  if(Object.keys(state.laneModes).length) XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['No L4','Mode'],...Object.entries(state.laneModes)]),'Pengaturan Swimlane');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Kode L5'],...[...usedCodes].map(c=>[c])]),'Riwayat Kode');
  appendAnalysisSheet(wb); appendRegisterSheet(wb); appendBpmnMetaSheet(wb); wb.Workbook=wb.Workbook||{};wb.Workbook.Sheets=wb.SheetNames.map(name=>({name,Hidden:['SOP','Analisis Proses','Daftar Perubahan','BPMN Metadata'].includes(name)?0:1}));
  const buf = XLSX.write(wb, {type:'array', bookType:'xlsx'});
  const base = state.fileName.replace(/\.[^.]+$/,'').replace(/^(?:BPA Master - )+/,'');
  const r = await saveFile('BPA Master - '+base+'.xlsx', new Blob([buf]));
  toast(r==='saved' ? 'Sheet elaborasi & pemetaan tersimpan.' : r==='declined' ? 'Unduhan dibatalkan.' : 'Gunakan tautan unduh yang muncul.');
  if(r==='link'){ const fb = $('fallback').innerHTML; const s1 = $('aiStatus'); if(s1) s1.innerHTML = fb; const s2 = $('newFb'); if(s2) s2.innerHTML = fb; }
  return r;
}
function readUsedCodes(wb){return wb.Sheets['Riwayat Kode']?XLSX.utils.sheet_to_json(wb.Sheets['Riwayat Kode'],{header:1}).slice(1).filter(r=>r[0]).map(r=>String(r[0])):[];}
function parseElabSheet(wb){
  const sn = wb.SheetNames.find(n=>/elaborasi/i.test(n)); if(!sn) return null;
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, raw:false, defval:''});
  const hi = aoa.findIndex(r=>r.map(norm).includes('jenis') && r.map(norm).includes('setelah aktivitas')); if(hi<0) return null;
  const hdr = aoa[hi].map(norm); if(hdr.includes('versi pembanding')&&hdr.includes('as-is')&&hdr.includes('to-be'))throw Error('File laporan Elaborasi Flowchart tidak dapat diimpor. Gunakan Master Workbook.'); const c = n => hdr.indexOf(norm(n));
  const out = {};
  for(const r of aoa.slice(hi+1)){
    const get = n => { const j = c(n); return j<0 ? '' : String(r[j]==null?'':r[j]).trim(); };
    const type = TYPE_FROM[norm(get('Jenis'))]; const l4 = get('No L4'); const after = get('Setelah Aktivitas');
    if(!type || !l4 || !after) continue;
    const it = {id:get('ID')||newId(), l4, type, after, text:get('Teks'), status:STATUS_FROM[norm(get('Status'))]||'proposed', source:norm(get('Sumber'))==='ai'?'ai':norm(get('Sumber'))==='teks bebas'?'free':'manual', reason:get('Alasan')};
    if(type==='gate'){ it.noLabel = get('Label Cabang')||'Tidak'; it.noTarget = get('Tujuan Tidak Lolos'); }
    if(type==='hold') it.label = get('Label Cabang')||'Hold';
    if(type==='reorder') it.subject = get('Aktivitas Dipindah');
    if(['newact','role','gate','declane'].includes(type) && get('PIC')) it.role = get('PIC');
    (out[l4] = out[l4] || []).push(it);
  }
  return out;
}
function mergeElab(imp){ let n = 0; for(const items of Object.values(imp))for(const it of items)if(it.type==='graph')try{reserveCodes(JSON.parse(it.text));}catch(e){} rememberCodes(); for(const [k,v] of Object.entries(imp)){ state.elab[k] = v; n += v.length; } saveElab(); return n; }
async function importElabFile(file){
  if(!file) return;
  try{
    await loadScript(LIBS.xlsx);
    const wb = XLSX.read(await file.arrayBuffer(), {type:'array'});
    const imp = parseElabSheet(wb), pos = parsePosSheets(wb), ex = parseExtraSheet(wb), rl = parseRoleSheet(wb);
    if(rl) mergeRoles(rl);
    if(!imp && !pos && !ex && !rl){ toast('Sheet "Elaborasi Flowchart", "SLA & Output", atau "Pemetaan Jabatan" tidak ditemukan di file itu.'); return; }
    const n = imp ? mergeElab(imp) : 0, np = pos ? applyPos(pos) : 0, nx = ex ? mergeExtra(ex) : 0; elabChanged();
    toast([n?n+' elemen elaborasi':'', nx?nx+' SLA/output':'', np?np+' pemetaan jabatan':''].filter(Boolean).join(', ')+' diimpor.');
  }catch(e){ console.error(e); toast('Gagal membaca file elaborasi.'); }
}

/* ================= Export ================= */
function svgToCanvas(svg, w, h, scale){
  return new Promise((res, rej)=>{
    const img = new Image();
    img.onload = () => { const c = document.createElement('canvas'); c.width = Math.round(w*scale); c.height = Math.round(h*scale);
      const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0,0,c.width,c.height); x.drawImage(img, 0, 0, c.width, c.height); res(c); };
    img.onerror = () => rej(new Error('Gagal mengubah flowchart menjadi gambar.'));
    img.src = 'data:image/svg+xml;charset=utf-8,'+encodeURIComponent(svg);
  });
}
const pdfSafe = s => String(s==null?'':s).replace(/[—–]/g,'-').replace(/›/g,'>').replace(/→/g,'->').replace(/[“”]/g,'"').replace(/[‘’]/g,"'").replace(/…/g,'...').replace(/•/g,'-').replace(/[^\x00-\xFF]/g,'');
function scopeL4s(scope){
  const all = [...state.l4s.values()].sort((a,b)=>cmpCode(a.l4no,b.l4no));
  const cur = state.l4s.get(state.current);
  if(scope==='one') return [cur];
  if(scope==='l3') return all.filter(g=>g.l3no===cur.l3no && g.l2no===cur.l2no);
  return all;
}
const TABLE_HEAD = ['Kode','Aktivitas','Deskripsi','PIC','Input dari','Output ke','SLA','Dokumen/Output','Catatan','Sebelum','Sesudah'];
const tableRows = g => g.acts.filter(a=>!isGateRow(a)).map(a=>{ const pos = posFor(g.l4no, a.role, a.dec, a); const x = state.extra[a.l5no] || {}; return [a.l5no, a.l5, a.desc, pos ? pos+' ('+a.role+')' : a.role, a.inParty, a.outParty, x.sla||'', x.out||'', x.note||'', a.prev, a.next]; });
const blobToU8 = async b => new Uint8Array(await b.arrayBuffer());
const canvasBlob = (c, type, q) => new Promise(r=>c.toBlob(r, type, q));

function exportItems(k,m){return m==='asis'?[]:elabFor(k,m==='tobe');}
const exportVersionName=m=>({asis:'AS-IS',tobe:'TO-BE',approved:'TO-BE Diterima',both:'AS-IS & TO-BE'})[m];
function exportPlan(list,version){return list.flatMap(g=>(version==='both'?['asis','tobe']:[version]).map(m=>({...g,exportMode:m})));}
function exportHeading(g){return g.l4no+' — '+g.l4+' ['+exportVersionName(g.exportMode)+']';}
async function makeExcelReport(list,prog){
  await loadScript(LIBS.xlsx); const wb=XLSX.utils.book_new();
  for(const m of [...new Set(list.map(g=>g.exportMode))]){
    const rows=[['No L2','L2','No L3','L3','No L4','L4','Versi',...TABLE_HEAD]];
    for(const g of list.filter(x=>x.exportMode===m)){
      prog('Memproses '+exportHeading(g));
      const d=buildDiagram(g,'',{ver:m==='asis'?'asis':m==='tobe'?'tobe':'approved',items:exportItems(g.l4no,m)});
      for(const r of tableRows({...g,acts:d.acts}))rows.push([g.l2no,g.l2,g.l3no,g.l3,g.l4no,g.l4,exportVersionName(m),...r]);
    }
    const ws=XLSX.utils.aoa_to_sheet(rows); ws['!cols']=[10,30,10,30,10,35,20,14,45,55,30,25,25,18,35,35,25,25].map(w=>({wch:w})); ws['!autofilter']={ref:ws['!ref']};
    XLSX.utils.book_append_sheet(wb,ws,'Laporan '+exportVersionName(m));
  }
  return new Blob([XLSX.write(wb,{type:'array',bookType:'xlsx'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function isChangeReport(){return document.querySelector('input[name=fmt]:checked').value==='docx'&&$('wordPurpose').value==='change';}
async function makeChangeDocx(g,version,paper,prog,mono,includeAnalysis,includeApproval){
 await loadScript(LIBS.docx);
 const D=window.docx, latest=version==='tobe', diff=compareProcess(g,latest), findings=analysisRecords.filter(r=>r.l4===g.l4no), sections=[];
 const accent=mono?'222222':'8A1E45', mg=850, dm=567;
 const p=(t,size=22,bold=false,color='222222')=>new D.Paragraph({spacing:{after:100},children:[new D.TextRun({text:String(t??''),font:'Arial',size,bold,color})]});
 const h=t=>new D.Paragraph({heading:D.HeadingLevel.HEADING_1,spacing:{before:240,after:120},children:[new D.TextRun({text:t,font:'Arial',size:28,bold:true,color:accent})]});
 const sub=t=>new D.Paragraph({keepNext:true,spacing:{before:180,after:100},children:[new D.TextRun({text:t,font:'Arial',size:23,bold:true,color:accent})]});
 const table=(headers,rows,widths)=>new D.Table({width:{size:100,type:D.WidthType.PERCENTAGE},rows:[headers,...rows].map((row,i)=>new D.TableRow({tableHeader:i===0,children:row.map((v,j)=>new D.TableCell({width:{size:widths[j],type:D.WidthType.PERCENTAGE},shading:i===0?{fill:accent,type:D.ShadingType.CLEAR,color:'auto'}:undefined,children:[p(v||'—',19,i===0,i===0?'FFFFFF':'222222')]}))}))});
 const portrait=children=>({properties:{page:{size:{width:11906,height:16838,orientation:D.PageOrientation.PORTRAIT},margin:{top:mg,bottom:mg,left:mg,right:mg}}},children});
 const labels={added:'Ditambahkan',modified:'Diubah',removed:'Dihapus'};
 const overview=[p('PROCESS CHANGE REPORT',40,true,accent),p(g.l4no+' — '+g.l4,30,true),p('Versi pembanding: '+(latest?'TO-BE Terbaru (termasuk usulan)':'TO-BE Diterima')),p('Dokumen kerja — bukan bukti persetujuan process owner.',20,false,'666666'),h('1. Process Overview'),table(['Identitas','Keterangan'],[['L2',g.l2no+' — '+g.l2],['L3',g.l3no+' — '+g.l3],['L4',g.l4no+' — '+g.l4],['Sumber data',state.fileName],['Tanggal export',new Date().toLocaleDateString('id-ID')]], [25,75])];
 if(includeAnalysis){overview.push(h('2. Process Analysis'));if(!findings.length)overview.push(p('Belum ada temuan Analisis Proses pada subproses ini.'));for(const f of findings){overview.push(sub(f.id+' — '+f.focus),table(['Aspek','Keterangan'],[['Aktivitas',f.activity],['Temuan',f.finding],['Bukti / Konteks',f.evidence],['Risiko / Dampak',f.risk],['Rekomendasi',f.recommendation],['Status',f.status],['Catatan',f.note]],[28,72]));}}else overview.push(p('Bagian 2. Process Analysis tidak disertakan sesuai pilihan export.',19,false,'666666'));
 sections.push(portrait(overview));
 for(const [m,title] of [['asis','3. AS-IS Process'],[version,'4. TO-BE Process']]){
  prog('Membuat '+title);const d=buildDiagram(g,'',{ver:m==='asis'?'asis':m==='tobe'?'tobe':'approved',items:exportItems(g.l4no,m),mono});
  const canvas=await svgToCanvas(d.svg,d.w,d.h,Math.min(2,8192/Math.max(d.w,d.h)));const blob=await canvasBlob(canvas,'image/png');if(!blob)throw Error('Gagal membuat gambar flowchart.');
  const dims=paper==='a3'?{w:16838,h:23811}:{w:11906,h:16838};const sc=Math.min((dims.h-2*dm)/15/d.w,((dims.w-2*dm)/15-85)/d.h);
  sections.push({properties:{page:{size:{width:dims.w,height:dims.h,orientation:D.PageOrientation.LANDSCAPE},margin:{top:dm,bottom:dm,left:dm,right:dm}}},children:[h(title),p(g.l4no+' — '+g.l4+' | '+exportVersionName(m),20),new D.Paragraph({alignment:D.AlignmentType.CENTER,children:[new D.ImageRun({type:'png',data:await blobToU8(blob),transformation:{width:Math.max(1,Math.floor(d.w*sc)),height:Math.max(1,Math.floor(d.h*sc))}})]})]});
 }
 const counts=['added','modified','removed'].map(t=>[labels[t],String(diff.changes.filter(c=>c.type===t).length)]);
 const body=[h('5. Change Summary'),table(['Indikator','Jumlah'],[['Aktivitas AS-IS',String(diff.before)],['Aktivitas TO-BE',String(diff.after)],...counts],[75,25]),p('Perhitungan berdasarkan perubahan bersih per aktivitas, bukan jumlah editing.'),h('6. Change Register')];
 if(!diff.changes.length)body.push(p('Tidak ada perubahan substansi proses pada versi pembanding ini.'));
 for(const c of diff.changes){const r=registerEntry(g.l4no,c.code),f=findings.find(x=>x.id===r?.finding);body.push(sub(registerId(g.l4no,c.code)+' — '+labels[c.type]),p(c.code+' | '+c.name,22,true),table(['Aspek','AS-IS','TO-BE'],c.fields.map(x=>[x.label,x.before,x.after]),[28,36,36]),table(['Dokumentasi','Keterangan'],[['Status',registerStatus(r,c)],['Referensi temuan',r?.finding?(f?f.id+' — '+f.finding:r.finding+' (temuan tidak tersedia)'):'Tanpa referensi'],['Alasan perubahan',r?.reason||'Belum diisi'],['Expected benefit',r?.benefit||'Belum diisi']],[30,70]));if(r&&registerStatus(r,c)==='Perlu ditinjau ulang')body.push(p('Catatan berasal dari versi perubahan sebelumnya; perlu dikonfirmasi kembali.',19,true,accent));}
 if(includeApproval)body.push(h('7. Validation & Approval'),p('Diisi setelah validasi oleh pihak berwenang. Status TO-BE Diterima pada aplikasi bukan persetujuan formal.'),table(['Peran','Nama / Tanggal / Paraf'],[['Disusun oleh',''],['Diperiksa oleh',''],['Disetujui oleh / Process Owner','']],[40,60]));
 sections.push(portrait(body));prog('Menyusun dokumen Word…');return D.Packer.toBlob(new D.Document({creator:'Flowchart BPA',title:'Process Change Report — '+g.l4no,sections}));
}

function isElabReport(){return document.querySelector('input[name=fmt]:checked').value==='xlsx'&&document.querySelector('input[name=excelPurpose]:checked').value==='elab';}
async function makeElabReport(groups,version,prog){
 await loadScript(LIBS.xlsx);
 const latest=version==='tobe', rows=[], labels={added:'Proses baru',removed:'Hapus aktivitas',modified:'Ubah aktivitas'};
 const heads=[...ELAB_HEAD,'Atribut','AS-IS','TO-BE','No L5','Expected Benefit','ID Temuan','Status Catatan','Versi Pembanding'];
 for(const g of groups){
  prog('Membandingkan '+g.l4no+' '+g.l4);
  const accepted=new Map(compareProcess(g,false).changes.map(c=>[c.code,c]));
  for(const c of compareProcess(g,latest).changes){
   const r=registerEntry(g.l4no,c.code), a=accepted.get(c.code), status=a&&registerSignature(a)===registerSignature(c)?'Diterima':'Usulan';
   const fields=c.fields.length?c.fields:[{label:'Aktivitas',before:'',after:''}];
   for(const f of fields){
    const kind=c.type==='modified'?({'Nama aktivitas':'Ganti nama','Role':'Ganti PIC','Proses sebelum':'Ubah koneksi','Proses sesudah':'Ubah koneksi'}[f.label]||labels.modified):labels[c.type];
    rows.push([registerId(g.l4no,c.code),g.l4no,g.l4,kind,c.type==='added'?'':c.code,c.type==='removed'?f.before:f.after,'','',status,'Perbandingan AS-IS/TO-BE',r?.reason||'','',f.label==='Role'?f.after:'',f.label,f.before,f.after,c.code,r?.benefit||'',r?.finding||'',registerStatus(r,c),exportVersionName(version)]);
   }
  }
 }
 const wb=XLSX.utils.book_new(), ws=XLSX.utils.aoa_to_sheet([heads,...rows]);
 ws['!cols']=[32,12,42,23,20,50,18,20,14,28,50,20,25,26,55,55,20,45,24,25,25].map(wch=>({wch}));ws['!autofilter']={ref:ws['!ref']};
 XLSX.utils.book_append_sheet(wb,ws,'Elaborasi Flowchart');
 const info=[['Keterangan','Nilai'],['Jenis file','Laporan perubahan bersih AS-IS vs TO-BE; bukan file import atau backup.'],['Versi',exportVersionName(version)],['Cakupan',groups.map(g=>g.l4no).join(', ')],['Jumlah baris detail',rows.length],['Catatan','Satu baris per atribut berubah; ID dapat berulang. Kolom legacy yang tidak dapat diturunkan secara aman dibiarkan kosong. Status Diterima berarti sama dengan perubahan TO-BE diterima, bukan persetujuan formal. Sumber adalah hasil perbandingan, bukan asal tindakan.'],['Pemulihan','Gunakan Simpan Master Workbook untuk menyimpan dan mengimpor kembali pekerjaan.'],['Catatan perubahan','Status Perlu ditinjau ulang berarti alasan/manfaat berasal dari versi perubahan sebelumnya.'],['Batasan','Perbandingan menggunakan engine existing; perubahan AS-IS langsung dan atribut bersama tidak dihitung sebagai perubahan TO-BE.']];
 const wi=XLSX.utils.aoa_to_sheet(info);wi['!cols']=[{wch:25},{wch:110}];XLSX.utils.book_append_sheet(wb,wi,'Keterangan');
 return new Blob([XLSX.write(wb,{type:'array',bookType:'xlsx'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function updateExportOptions(){
 const fmt=document.querySelector('input[name=fmt]:checked').value;
 const backup=fmt==='xlsx'&&document.querySelector('input[name=excelPurpose]:checked').value==='backup';
 const change=isChangeReport(), elab=isElabReport(); $('elabExportOptions').hidden=!elab; $('wordPurposeBox').hidden=fmt!=='docx'; $('changeReportOptions').hidden=!change; $('scopeSeg').parentElement.hidden=change; $('exportVersion').parentElement.hidden=change||elab; $('exportTableOption').hidden=change||fmt==='xlsx';
 $('exportPurpose').hidden=fmt!=='xlsx'; $('exportReportOptions').hidden=backup;
 $('exportDocOptions').hidden=fmt==='xlsx'; $('exportTableOption').hidden=change||fmt==='xlsx';
 $('exportHelp').textContent=elab?'Daftar perubahan aktual AS-IS vs TO-BE per atribut. Bukan file untuk diimpor kembali; gunakan Master Workbook untuk backup.':change?'Process Change Report untuk L4 aktif. Narasi A4 portrait; diagram pada halaman landscape sesuai ukuran kertas.':backup?'Backup mencakup seluruh proyek, AS-IS, rancangan TO-BE, catatan dan mapping. Gunakan Buka File untuk melanjutkan pekerjaan.':fmt==='xlsx'?'Laporan tabel aktivitas sesuai cakupan dan versi. Bukan file backup untuk diimpor kembali.':'Dokumen flowchart dengan opsi tabel uraian aktivitas.';
 $('exGo').textContent=elab?'Unduh Elaborasi Flowchart':change?'Buat Process Change Report':backup?'Unduh backup':'Buat file';
}
async function makePdf(list, paper, incTable, prog, mono){
  await loadScript(LIBS.jspdf); await loadScript(LIBS.autotable);
  const { jsPDF } = window.jspdf; let doc = null; const label = $('statusLabel').value.trim();
  for(let i=0;i<list.length;i++){
    const g = list[i]; prog('Memproses '+(i+1)+' dari '+list.length+': '+g.l4no+' '+g.l4);
    const d = buildDiagram(g, label, {ver:g.exportMode==='asis'?'asis':g.exportMode==='tobe'?'tobe':'approved', items: exportItems(g.l4no,g.exportMode), mono});
    const canvas = await svgToCanvas(d.svg, d.w, d.h, 2);
    const orient = d.w >= d.h*1.05 ? 'landscape' : 'portrait';
    if(!doc) doc = new jsPDF({orientation:orient, unit:'pt', format:paper, compress:true}); else doc.addPage(paper, orient);
    const pw = doc.internal.pageSize.getWidth(), ph = doc.internal.pageSize.getHeight(), mg = 22;
    const sc = Math.min((pw-2*mg)/d.w, (ph-2*mg-12)/d.h);
    doc.addImage(canvas.toDataURL('image/jpeg', 0.9), 'JPEG', (pw-d.w*sc)/2, mg, d.w*sc, d.h*sc);
    doc.setFontSize(7.5); doc.setTextColor(110); doc.text(pdfSafe(exportHeading(g)+' - flowchart'), mg, ph-12);
    if(incTable){
      doc.addPage(paper, 'landscape');
      const tw2 = doc.internal.pageSize.getWidth();
      doc.setFontSize(13); if(mono) doc.setTextColor(0); else doc.setTextColor(138,30,69); doc.setFont(undefined,'bold'); doc.text(pdfSafe(exportHeading(g)+': Uraian Aktivitas'), mg, mg+14);
      doc.setFont(undefined,'normal'); doc.setFontSize(8.5); doc.setTextColor(90); doc.text(pdfSafe('SOP '+g.l3no+' '+g.l3+' · '+g.l2+' · Sumber: '+state.fileName), mg, mg+28);
      const avail = tw2 - 2*mg; const pct = [5,11,22,10,7,7,6,9,9,7,7];
      const colStyles = {}; pct.forEach((p,j)=>colStyles[j]={cellWidth:avail*p/100}); colStyles[0].fontStyle='bold';
      doc.autoTable({ startY:mg+38, head:[TABLE_HEAD], body:tableRows({...g,acts:d.acts}).map(r=>r.map(pdfSafe)), margin:{left:mg,right:mg},
        styles:{fontSize:paper==='a3'?8.5:7, cellPadding:3, overflow:'linebreak', valign:'top', lineColor:mono?[205,205,205]:[200,208,220], lineWidth:0.4},
        headStyles:{fillColor:mono?[41,41,41]:[138,30,69], textColor:255}, alternateRowStyles:{fillColor: mono ? [245,245,245] : [250,246,244]}, columnStyles:colStyles });
    }
    await new Promise(r=>setTimeout(r,0));
  }
  return doc.output('blob');
}

async function makeDocx(list, paper, incTable, prog, mono){
  await loadScript(LIBS.docx);
  const D = window.docx; const label = $('statusLabel').value.trim();
  const dims = paper==='a3' ? {w:16838, h:23811} : {w:11906, h:16838};
  const mg = 567; const pxPerTwip = 96/1440;
  const sections = [];
  const cellPara = (t, bold, color) => new D.Paragraph({ children:[ new D.TextRun({text:String(t||''), bold:!!bold, size:paper==='a3'?17:15, color:color||(mono?'222222':'1F2A37'), font:'Arial'}) ] });
  for(let i=0;i<list.length;i++){
    const g = list[i]; prog('Memproses '+(i+1)+' dari '+list.length+': '+g.l4no+' '+g.l4);
    const d = buildDiagram(g, label, {ver:g.exportMode==='asis'?'asis':g.exportMode==='tobe'?'tobe':'approved', items: exportItems(g.l4no,g.exportMode), mono});
    const canvas = await svgToCanvas(d.svg, d.w, d.h, 2);
    const blob = await canvasBlob(canvas, 'image/jpeg', 0.9);
    const land = d.w >= d.h*1.05;
    const pgW = land ? dims.h : dims.w, pgH = land ? dims.w : dims.h;
    const maxW = (pgW-2*mg)*pxPerTwip, maxH = (pgH-2*mg)*pxPerTwip - 40;
    const sc = Math.min(maxW/d.w, maxH/d.h);
    sections.push({ properties:{ page:{ size:{ width:dims.w, height:dims.h, orientation: land ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT }, margin:{top:mg,bottom:mg,left:mg,right:mg} } },
      children:[
        new D.Paragraph({ children:[ new D.TextRun({text:exportHeading(g), bold:true, size:24, color: mono ? '000000' : '8A1E45', font:'Arial'}) ] }),
        new D.Paragraph({ alignment:D.AlignmentType.CENTER, children:[ new D.ImageRun({ type:'jpg', data:await blobToU8(blob), transformation:{ width:Math.round(d.w*sc), height:Math.round(d.h*sc) } }) ] })
      ]});
    if(incTable){
      const widths = [5,11,22,10,7,7,6,9,9,7,7];
      const head = new D.TableRow({ tableHeader:true, children:TABLE_HEAD.map((t,j)=>new D.TableCell({ width:{size:widths[j], type:D.WidthType.PERCENTAGE}, shading:{fill:mono?'292929':'8A1E45', type:D.ShadingType.CLEAR, color:'auto'}, children:[cellPara(t, true, 'FFFFFF')] })) });
      const body = tableRows({...g,acts:d.acts}).map((r,ri)=>new D.TableRow({ children:r.map((t,j)=>new D.TableCell({ width:{size:widths[j], type:D.WidthType.PERCENTAGE}, shading: ri%2 ? {fill: mono ? 'F5F5F5' : 'FAF6F4', type:D.ShadingType.CLEAR, color:'auto'} : undefined, children:[cellPara(t, j===0)] })) }));
      sections.push({ properties:{ page:{ size:{ width:dims.w, height:dims.h, orientation:D.PageOrientation.LANDSCAPE }, margin:{top:mg,bottom:mg,left:mg,right:mg} } },
        children:[
          new D.Paragraph({ children:[ new D.TextRun({text:exportHeading(g)+': Uraian Aktivitas', bold:true, size:26, color: mono ? '000000' : '8A1E45', font:'Arial'}) ] }),
          new D.Paragraph({ spacing:{after:160}, children:[ new D.TextRun({text:'SOP '+g.l3no+' '+g.l3+' · '+g.l2+' · Sumber: '+state.fileName, size:17, color:mono?'555555':'5B6B7F', font:'Arial'}) ] }),
          new D.Table({ width:{size:100, type:D.WidthType.PERCENTAGE}, rows:[head, ...body] })
        ]});
    }
    await new Promise(r=>setTimeout(r,0));
  }
  const doc = new D.Document({ creator:'Flowchart BPA', title:'Flowchart BPA', sections });
  return D.Packer.toBlob(doc);
}

let _dl;
async function getDownloads(){
  if(_dl !== undefined) return _dl;
  try{ _dl = (window.claude && window.claude.use) ? await window.claude.use('downloads') : null; }catch(e){ _dl = null; }
  return _dl;
}
async function saveFile(name, blob){
  if(!(window.claude && window.claude.use)){
    const url = URL.createObjectURL(blob); const a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url), 60000); return 'saved';
  }
  const dl = await getDownloads();
  if(dl){
    try{ await dl.save({filename:name, data:blob}); return 'saved'; }
    catch(e){ if(e && e.code==='declined') return 'declined'; if(e && e.code==='rate_limited') throw new Error('Masih ada konfirmasi unduhan yang terbuka. Coba lagi sebentar.'); }
  }
  // fallback: link (works when the page is opened outside a sandboxed frame)
  const url = URL.createObjectURL(blob);
  $('fallback').innerHTML = 'Unduhan otomatis tidak tersedia di tampilan ini. <a href="'+url+'" download="'+esc(name)+'">Klik di sini untuk mengunduh '+esc(name)+'</a>.';
  return 'link';
}
function openExport(){
  const cur = state.l4s.get(state.current);
  $('scopeSeg').innerHTML =
    '<label><input type="radio" name="scope" value="one" checked>Subproses ini ('+esc(cur.l4no)+')</label>'+
    '<label><input type="radio" name="scope" value="l3">SOP '+esc(cur.l3no)+' ('+scopeL4s('l3').length+' subproses)</label>'+
    '<label><input type="radio" name="scope" value="all">Semua ('+state.l4s.size+' subproses)</label>';
  $('prog').textContent = ''; $('fallback').innerHTML = ''; $('exportVersion').value=$('processEditMode').value==='asis'?'asis':$('processEditMode').value==='tobe'?'tobe':'approved';
  updateExportOptions(); $('exportDlg').classList.remove('hidden');
}
async function runExport(){
  const fmt = document.querySelector('input[name=fmt]:checked').value;
  const scope = document.querySelector('input[name=scope]:checked').value;
  const paper = document.querySelector('input[name=paper]:checked').value;
  const inc = $('incTable').checked;
  const mono = document.querySelector('input[name=color]:checked').value === 'mono';
  const version=$('exportVersion').value; const list = exportPlan(scopeL4s(scope),version); const cur = state.l4s.get(state.current);
  const base = scope==='one' ? cur.l4no+' '+cur.l4 : scope==='l3' ? 'SOP '+cur.l3no+' '+cur.l3 : 'Semua subproses';
  const change=isChangeReport();
  const name = (isElabReport()?('Elaborasi Flowchart - '+base+' - '+exportVersionName($('elabExportVersion').value)):change?('Process Change Report - '+cur.l4no+' '+cur.l4+' - '+exportVersionName($('changeReportVersion').value)):((fmt==='xlsx'?'Laporan ':'Flowchart ')+base+' - '+exportVersionName(version))).replace(/[\\/:*?"<>|]+/g,' ').replace(/\s+/g,' ').trim().slice(0,150)+'.'+fmt;
  const controls=[...$('exportDlg').querySelectorAll('input,select,button')]; const priorDisabled=controls.map(c=>c.disabled); controls.forEach(c=>c.disabled=true); $('fallback').innerHTML = '';
  const prog = t => $('prog').textContent = t;
  try{
    if(fmt==='xlsx'&&document.querySelector('input[name=excelPurpose]:checked').value==='backup'){const r=await exportElabSheet();prog(r==='saved'?'Backup tersimpan.':r==='declined'?'Unduhan dibatalkan.':r==='empty'?'Belum ada data.':'Backup siap diunduh.');return;}
    const blob = isElabReport() ? await makeElabReport(scopeL4s(scope),$('elabExportVersion').value,prog) : change ? await makeChangeDocx(cur,$('changeReportVersion').value,paper,prog,mono,$('changeIncludeAnalysis').checked,$('changeIncludeApproval').checked) : fmt==='pdf' ? await makePdf(list, paper, inc, prog, mono) : fmt==='docx' ? await makeDocx(list, paper, inc, prog, mono) : await makeExcelReport(list,prog);
    prog('File siap ('+(blob.size/1024/1024).toFixed(1)+' MB). Menyimpan…');
    const r = await saveFile(name, blob);
    prog(r==='saved' ? 'Tersimpan: '+name : r==='declined' ? 'Unduhan dibatalkan.' : 'File siap.');
  }catch(e){ prog('Gagal: '+(e.message||e)); console.error(e); }
  finally{ controls.forEach((c,i)=>c.disabled=priorDisabled[i]); }
}

/* ================= SLA & Output ================= */
const hasExtra = () => Object.values(state.extra).some(x=>x && (x.sla || x.out || x.note || (x.docs&&x.docs.length)));
/* v1.7.1: extra input / output documents per activity: extra[code].docs = [{d:'in'|'out', k:'object'|'store', n:name}] */
function docsOf(code){ const x=state.extra[code]; return (x&&Array.isArray(x.docs))?x.docs.filter(d=>d&&String(d.n||'').trim()):[]; }
function encodeDocs(ds){ return (ds||[]).filter(d=>String(d.n||'').trim()).map(d=>(d.d==='out'?'Output':'Input')+' · '+(d.k==='store'?'Data':'Dokumen')+': '+String(d.n).replace(/;/g,',').trim()).join('; '); }
function parseDocs(t){ return String(t||'').split(/;|\n/).map(x=>x.trim()).filter(Boolean).map(x=>{ const m=x.match(/^(input|output)\s*[·\-–|]?\s*(dokumen|data)?\s*:\s*(.+)$/i); return m?{d:/^o/i.test(m[1])?'out':'in',k:/^data/i.test(m[2]||'')?'store':'object',n:m[3].trim()}:{d:'in',k:'object',n:x}; }); }
function saveExtra(){ const c = {}, mem = {}; for(const [k,v] of Object.entries(state.extra)){ if(!v) continue; const docs=Array.isArray(v.docs)?v.docs.filter(d=>d&&String(d.n||'').trim()):[]; const keep=v.sla||v.out||v.note||docs.length; if(keep||(Array.isArray(v.docs)&&v.docs.length)) mem[k]=v; if(keep){ const o={...v}; if(docs.length) o.docs=docs.map(d=>({d:d.d,k:d.k,n:String(d.n).trim()})); else delete o.docs; c[k]=o; } } state.extra = mem; store.set('flowbpa:extra', JSON.stringify(c)); }
function openExtra(focusCode){ state.sideMode = 'extra'; $('btnX').classList.add('on'); $('btnElab').classList.remove('on'); $('btnR').classList.remove('on'); $('detail').classList.remove('hidden'); renderExtra(focusCode); }
function renderExtra(focusCode){
  const g = state.l4s.get(state.current); if(!g) return;
  const acts = actsOf(g); const filled = acts.filter(a=>state.extra[a.l5no] && (state.extra[a.l5no].sla || state.extra[a.l5no].out || state.extra[a.l5no].note)).length;
  let h = '<button class="tb x" data-a="close">✕</button><div class="dcode" style="color:#334155">✎ CATATAN</div><h3>'+esc(g.l4no+' '+g.l4)+'</h3>'+
    '<p class="muted">Isi SLA, dokumen/output, dan catatan per aktivitas. Ketik langsung saat diskusi dengan klien; tersimpan otomatis dan langsung tampil di flowchart. Pindah kolom dengan Tab. '+filled+' dari '+acts.length+' aktivitas sudah terisi.</p>';
  h += '<datalist id="slaList">'+['1 hari kerja','2 hari kerja','3 hari kerja','5 hari kerja','7 hari kerja','14 hari kerja','1 bulan','Sesuai jadwal','Maks. 1 hari'].map(v=>'<option value="'+v+'">').join('')+'</datalist>';
  for(const a of acts){
    const x = state.extra[a.l5no] || {};
    h += '<div class="xrow" data-row="'+esc(a.l5no)+'"><div class="xh"><b>'+esc(a.l5no)+'</b> '+esc(a.l5)+'</div>'+
      '<div class="xin"><input data-xf="sla" data-code="'+esc(a.l5no)+'" list="slaList" placeholder="SLA (mis. 2 HK)" value="'+esc(x.sla||'')+'"><input data-xf="out" data-code="'+esc(a.l5no)+'" placeholder="Tambah dokumen / output (manual)" value="'+esc(x.out||'')+'"><input data-xf="note" data-code="'+esc(a.l5no)+'" class="xnote" placeholder="Catatan (opsional, muncul sebagai nomor di flowchart)" value="'+esc(x.note||'')+'"></div></div>';
  }
  h += '<div class="io"><p class="muted" style="margin-top:0">Ikut tersimpan saat lw klik <b>Unduh sheet</b> (sheet "SLA &amp; Output"), bersama elaborasi dan pemetaan jabatan.</p><div class="cbtn"><button class="sbtn" data-a="exportElab">Unduh sheet (.xlsx)</button></div></div>';
  $('detail').innerHTML = h;
  if(focusCode){ const r = $('detail').querySelector('.xrow[data-row="'+focusCode+'"]'); if(r){ r.scrollIntoView({block:'center'}); r.classList.add('flash'); setTimeout(()=>r.classList.remove('flash'), 1400); const i = r.querySelector('input'); if(i) i.focus(); } }
}
function onExtraInput(e){
  const t = e.target; if(!t.dataset || !t.dataset.xf) return false;
  const c = t.dataset.code; const x = state.extra[c] || (state.extra[c] = {}); x[t.dataset.xf] = t.value.replace(/\s+/g,' ').trim();
  saveExtra(); clearTimeout(window._xt); window._xt = setTimeout(()=>{ redraw(); renderTree(); }, 350);
  return true;
}
function parseExtraSheet(wb){
  const sn = wb.SheetNames.find(n=>/sla/i.test(n)); if(!sn) return null;
  const a = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, raw:false, defval:''}); const hi = a.findIndex(r=>r.map(norm).includes('no l5')); if(hi<0) return null;
  const h = a[hi].map(norm), out = {};
  for(const r of a.slice(hi+1)){ const g = k => { const j = h.indexOf(k); return j<0 ? '' : String(r[j]==null?'':r[j]).replace(/\s+/g,' ').trim(); };
    const c = g('no l5'); const sla = g('sla'), o = g('dokumen/output'), nt = g('catatan'), dk = g('input/output tambahan'); if(c && (sla || o || nt || dk)){ out[c] = {}; if(sla) out[c].sla = sla; if(o) out[c].out = o; if(nt) out[c].note = nt; if(dk) out[c].docs = parseDocs(dk); } }
  return out;
}
function mergeExtra(x){ if(!x) return 0; Object.assign(state.extra, x); saveExtra(); return Object.keys(x).length; }

/* ================= Pindah kolom belah ketupat (klik / seret) ================= */
function decInfo(actCode, kind){
  const g = state.l4s.get(state.current), list = elabOf(g.l4no), act = g.acts.find(a=>a.l5no===actCode);
  const ov = list.find(i=>i.type==='role' && i.after===actCode && i.status==='accepted');
  const actRole = ov ? ov.role : (act ? act.role : '');
  const gate = list.find(i=>i.type==='gate' && i.after===actCode && i.status!=='rejected');
  const dl = list.find(i=>i.type==='declane' && i.after===actCode && i.status!=='rejected');
  const cur = kind==='gate' ? (gate && gate.role) || (dl && dl.role) || actRole : (dl && dl.role) || actRole;
  return {g, list, actRole, gate, dl, cur};
}
const laneRole = ln => ln.roles[0];
function applyDecLane(actCode, kind, ln){
  if(laneMode(state.current)==='unit'){ toast('Pindahkan decision pada mode Role atau Jabatan agar PIC tetap jelas.'); closeLanePop(); return; }
  const I = decInfo(actCode, kind), d0 = state.diagram;
  const actNode = d0.nodes.find(n=>n.type==='act' && n.act.l5no===actCode);
  const back = actNode && d0.lanes[actNode.lane].key === ln.key;
  const role = laneRole(ln), pos = ln.key.startsWith('pos::') ? ln.key.slice(5) : '';
  const setT = it => { it.role = role; if(pos) it.pos = pos; else delete it.pos; };
  if(kind==='gate' && I.gate){ if(back){ delete I.gate.role; delete I.gate.pos; } else setT(I.gate); }
  else {
    if(back){ if(I.dl) I.list.splice(I.list.indexOf(I.dl), 1); }
    else { let d = I.dl; if(!d){ d = {id:newId(), l4:I.g.l4no, type:'declane', after:actCode, status:'accepted', source:'manual', reason:'Dipindah langsung di flowchart'}; I.list.push(d); }
      setT(d); d.text = 'Keputusan setelah '+actCode+' → kolom '+laneShort(roleInfo(ln.key).name); }
  }
  closeLanePop(); elabChanged();
  toast(back ? 'Belah ketupat dikembalikan ke kolom aktivitasnya.' : 'Belah ketupat dipindah ke kolom '+laneShort(roleInfo(ln.key).name)+'.');
}
function closeLanePop(){ const p = document.getElementById('lanePop'); if(p) p.remove(); }

/* v1.3.8.9: TO-BE decision removal, preserving L1-L5 and AS-IS.
   Special case requested for 01.02.07: close the L4 and remove the
   reciprocal cross-L4 predecessor on 01.03.01 in the same undo unit. */
function removeDecisionToEnd(actCode){
  if($('processEditMode').value!=='tobe')return toast('Hapus gateway hanya tersedia dalam Edit TO-BE.');
  const g=state.l4s.get(state.current);
  if(!g||actCode!=='01.02.07'||g.l4no!=='01.02')
    return toast('Penghapusan ini hanya untuk keputusan 01.02.07 yang sudah disepakati.');
  const other=state.l4s.get('01.03');
  if(!other)return toast('Subproses 01.03 tidak ditemukan. Koneksi lintas L4 belum bisa diperbarui.');
  let aRows,bRows;
  try{
    aRows=structuralRows(g,'tobe');bRows=structuralRows(other,'tobe');
    const a=aRows.find(x=>x.l5no===actCode),b=bRows.find(x=>x.l5no==='01.03.01');
    if(!a||!b)throw Error('Aktivitas terkait tidak ditemukan.');
    const next=parseFlow(a.next);
    if(!next.some(f=>f.kind==='code'&&f.code==='01.03.01'))
      throw Error('Koneksi 01.02.07 → 01.03.01 tidak ditemukan. Tidak ada perubahan dilakukan.');
    if(next.some(f=>f.kind==='code'&&f.code!=='01.03.01'))
      throw Error('Ada cabang aktivitas lain. Periksa mapping sebelum menghapus gateway.');
    const prev=parseFlow(b.prev);
    if(!prev.some(f=>f.kind==='code'&&f.code===actCode))
      throw Error('Koneksi balik 01.03.01 belum konsisten. Periksa mapping terlebih dahulu.');
    const conflicts=[g,other].flatMap(group=>(state.elab[group.l4no]||[])
      .filter(i=>i.status!=='rejected'&&i.type!=='graph'&&['gate','declane','newact','reorder'].includes(i.type)
       &&(i.after===actCode||i.after==='01.03.01'||i.subject===actCode||i.subject==='01.03.01')));
    if(conflicts.length)throw Error('Ada elaborasi lama terkait gateway/urutan. Tinjau dahulu agar tidak tertimpa.');
    a.next='SELESAI';
    a.decisionQuestion='';a.decisionRole='';a.dec='';
    b.prev=flowString(prev.filter(f=>!(f.kind==='code'&&f.code===actCode)));
    if(!confirm('Hapus gateway TO-BE setelah 01.02.07?\n\n'
      +'01.02.07 → SELESAI\n'
      +'Hapus koneksi lintas L4 01.02.07 → 01.03.01\n'
      +'01.03.01 tetap memiliki pemicu lainnya (jika ada).\n\n'
      +'AS-IS, aktivitas, Role/Jabatan/Unit, dan Process Architecture tetap.'))return;
    const prior=snapshot();storageFailed=false;
    saveUnifiedDraft(g,aRows,materializeDraft(g).ids,'Hapus gateway 01.02.07; lanjut SELESAI');
    saveUnifiedDraft(other,bRows,materializeDraft(other).ids,'Sinkronisasi koneksi dari 01.02.07');
    closeLanePop();
    undoStack.push({before:prior,after:snapshot()});
    toast('Gateway 01.02.07 dihapus dari TO-BE. Jalur langsung ke SELESAI.');
  }catch(e){toast(e.message,9000);}
}
function openLanePop(el, actCode, kind){
  closeLanePop();
  const d = state.diagram, I = decInfo(actCode, kind);
  const dn = d.nodes.find(n=>n.type==='dec' && n.act.l5no===actCode && (kind==='gate') === !!n.elab), an = d.nodes.find(n=>n.type==='act' && n.act.l5no===actCode);
  const curKey = dn ? d.lanes[dn.lane] : null, actKey = an ? d.lanes[an.lane].key : ''; const lanes = d.lanes;
  const pop = document.createElement('div'); pop.className = 'lanepop'; pop.id = 'lanePop';
  pop.innerHTML = '<div class="t">'+(kind==='gate'?'Gateway':'Keputusan')+' setelah '+esc(actCode)+'</div><div style="color:var(--muted);margin-bottom:4px">Pindahkan ke kolom:</div>'+
    lanes.map((l,i)=>'<button class="ln'+(curKey && curKey.key===l.key?' cur':'')+'" data-li="'+i+'">'+(curKey && curKey.key===l.key?'✓ ':'')+esc(laneShort(roleInfo(l.key).name))+(l.key===actKey?' <span style="color:var(--muted);font-weight:400">(kolom aktivitas)</span>':'')+(l.kind!=='act'?' <span style="color:var(--muted);font-weight:400">(pihak luar)</span>':'')+'</button>').join('')+
    '<div class="hint">Bisa juga: seret belah ketupat langsung ke kolom lain.</div><div style="margin-top:8px">'+
    (kind==='gate' && I.gate ? '<button class="lnk2" data-pa="elab">Edit gateway</button>' : '<button class="lnk2" data-pa="detail">Lihat aktivitas</button>')+(actCode==='01.02.07' && $('processEditMode').value==='tobe' ? '<button class="lnk2" style="color:#b4233a;font-weight:700" data-pa="remove-end">Hapus gateway → SELESAI</button>' : '')+'<button class="lnk2" data-pa="close">Tutup</button></div>';
  const cv = $('canvas'), r = el.getBoundingClientRect(), c = cv.getBoundingClientRect();
  pop.style.left = Math.max(8, Math.min(cv.scrollLeft + r.right - c.left + 8, cv.scrollLeft + cv.clientWidth - 272))+'px';
  cv.appendChild(pop);
  const want = cv.scrollTop + r.top - c.top, maxTop = cv.scrollTop + cv.clientHeight - pop.offsetHeight - 8;
  pop.style.top = Math.max(cv.scrollTop + 8, Math.min(want, maxTop))+'px';
  pop.onclick = e => {
    e.stopPropagation();
    const b = e.target.closest('[data-li]'); if(b){ applyDecLane(actCode, kind, lanes[+b.dataset.li]); return; }
    const a = e.target.closest('[data-pa]'); if(!a) return;
    if(a.dataset.pa==='remove-end'){ removeDecisionToEnd(actCode); return; }
    if(a.dataset.pa==='close') closeLanePop();
    if(a.dataset.pa==='elab'){ closeLanePop(); openElab(I.gate.id); }
    if(a.dataset.pa==='detail'){ closeLanePop(); showDetail(actCode); }
  };
}
// drag a diamond sideways onto another column
let _drag = null;
function svgX(ev){ const svg = $('canvas').querySelector('svg'); const r = svg.getBoundingClientRect(); return (ev.clientX - r.left) / state.zoom; }
function laneAt(x){ return state.diagram.lanes.find(l=>x >= l.x && x < l.x + state.diagram.laneW); }
$('canvas').addEventListener('pointerdown', e => {
  const el = e.target.closest('.dmd'); if(!el || e.button!==0) return;
  _drag = {el, code:el.dataset.dec, kind:el.dataset.dk, x0:e.clientX, y0:e.clientY, moved:false, startLane: laneAt(svgX(e))};
  try{ el.setPointerCapture(e.pointerId); }catch(err){}
});
$('canvas').addEventListener('pointermove', e => {
  if(!_drag) return;
  const dx = e.clientX - _drag.x0, dy = e.clientY - _drag.y0;
  if(!_drag.moved && Math.hypot(dx, dy) < 6) return;
  if(!_drag.moved){ _drag.moved = true; closeLanePop(); $('canvas').classList.add('dragging'); }
  _drag.el.setAttribute('transform', 'translate('+(dx/state.zoom)+','+(dy/state.zoom*0.15)+')'); _drag.el.style.opacity = .75;
  const ln = laneAt(svgX(e)); const svg = $('canvas').querySelector('svg'); let hi = svg.querySelector('#laneHi');
  if(!hi){ hi = document.createElementNS('http://www.w3.org/2000/svg','rect'); hi.id = 'laneHi'; hi.setAttribute('fill','rgba(224,122,16,.10)'); hi.setAttribute('stroke','#e07a10'); hi.setAttribute('stroke-width','2'); hi.setAttribute('stroke-dasharray','8 5'); hi.style.pointerEvents = 'none'; svg.appendChild(hi); }
  if(ln){ hi.setAttribute('x', ln.x); hi.setAttribute('y', state.diagram.top); hi.setAttribute('width', state.diagram.laneW); hi.setAttribute('height', state.diagram.h - state.diagram.top - 90); hi.style.display = ''; }
  else hi.style.display = 'none';
});
$('canvas').addEventListener('pointerup', e => {
  if(!_drag) return; const D = _drag; _drag = null; $('canvas').classList.remove('dragging');
  if(!D.moved){ openLanePop(D.el, D.code, D.kind); return; }
  const ln = laneAt(svgX(e));
  if(ln && (!D.startLane || ln.key !== D.startLane.key)) applyDecLane(D.code, D.kind, ln);
  else { redraw(); if(!ln) toast('Lepaskan di dalam salah satu kolom (lane) untuk memindahkan.'); }
});
document.addEventListener('pointerdown', e => { if(!e.target.closest('#lanePop') && !e.target.closest('.dmd')) closeLanePop(); });

/* ================= Peran (lane) ================= */
function saveRoles(){ store.set('flowbpa:roles', JSON.stringify(state.roleEdit)); }
const hasRoleEdits = () => { const R = state.roleEdit; return Object.keys(R.alias).length || Object.values(R.added).some(l=>l.length) || Object.values(R.hidden).some(l=>l.length) || Object.values(R.order||{}).some(l=>l&&l.length); };
function organizationRows(g,mode){
  if(mode==='tobe') return materializeDraft(g).rows;
  if(mode==='asis') return g.acts;
  const graph=elabFor(g.l4no,state.showProposed).filter(i=>i.type==='graph').slice(-1)[0];
  const rows=graph?JSON.parse(graph.text):g.acts, ids=consumedIds(rows);
  const items=elabFor(g.l4no,state.showProposed).filter(i=>!ids.has(i.id));
  const ov=new Map(items.filter(i=>i.type==='role').map(i=>[i.after,i.role]));
  return rows.map(r=>({...r,role:ov.get(r.l5no)||r.role})).concat(items.filter(i=>i.type==='newact').map(i=>({role:i.role,l5no:i.id})));
}
function lanesOfL4(g){
  const mode=$('processEditMode').value; let rows, draft=[];
  try{rows=organizationRows(g,mode);}catch(e){rows=g.acts;}
  try{draft=materializeDraft(g).rows;}catch(e){draft=[];}
  const added=state.roleEdit.added[g.l4no]||[], parties=new Set();
  [...g.acts,...rows,...draft].forEach(a=>{splitParty(a.inParty).forEach(p=>parties.add(p));splitParty(a.outParty).forEach(p=>parties.add(p));});
  const roles=new Set([...g.acts,...rows,...draft].map(a=>a.role||'(Tanpa peran)'));
  added.forEach(r=>roles.add(r));
  const out=sortRoles([...roles]).map(role=>({role,kind:added.includes(role)?'added':'act',count:rows.filter(a=>(a.role||'(Tanpa peran)')===role).length,used:[...g.acts,...draft,...rows].some(a=>(a.role||'(Tanpa peran)')===role)}));
  sortRoles([...parties].filter(p=>!roles.has(p))).forEach(role=>out.push({role,kind:'party',count:0,hidden:(state.roleEdit.hidden[g.l4no]||[]).includes(role)}));
  return out;
}
function openRoles(){ state.sideMode = 'roles'; ['btnElab','btnX'].forEach(b=>$(b).classList.remove('on')); $('btnR').classList.add('on'); $('detail').classList.remove('hidden'); renderRoles(); }
function renderRoles(){
  // v1.6.0: swimlane settings by jabatan / unit (process roles are not shown on the canvas).
  const g = state.l4s.get(state.current); if(!g || !state.diagram) return;
  const d = state.diagram, mode = laneMode(g.l4no), byUnit = mode==='unit';
  const counts = d.lanes.map((ln,i)=>d.nodes.filter(n=>n.type==='act'&&n.lane===i).length);
  const actL = d.lanes.map((ln,i)=>({...ln,i,count:counts[i]})).filter(ln=>ln.kind==='act');
  const party = d.lanes.map((ln,i)=>({...ln,i})).filter(ln=>ln.kind!=='act');
  const R = state.roleEdit, manual = !!(R.order||{})[g.l4no];
  const nm = k => { const ri=roleInfo(k); return ri.name; };
  const addedOf = k => (R.added[g.l4no]||[]).find(r=>addedLaneKey(g.l4no,r)===k);
  let h = '<button class="tb x" data-a="close">✕</button><div class="ws-panel-title">Pengaturan Swimlane</div><div class="dcode">Proses aktif</div><h3>'+esc(g.l4no+' '+g.l4)+'</h3>'+
    '<p class="muted">Lane = '+(byUnit?'unit kerja':'jabatan')+'. Pengaturan ini berlaku untuk subproses ini, sama untuk AS-IS dan TO-BE. Mode aktif: '+esc($('processEditMode').value.toUpperCase())+'.</p>';
  h += '<section class="ws-role-add"><div class="sec">Tambah lane</div><div class="rgrid">'+
    '<label for="newLanePos">Jabatan'+(byUnit?' (opsional)':'')+'</label><input id="newLanePos" list="posListR" placeholder="mis. Supervisor Pengelolaan Payroll">'+
    '<label for="newLaneUnit">Unit kerja'+(byUnit?'':' (opsional)')+'</label><input id="newLaneUnit" list="unitListR" placeholder="mis. Layanan SDM">'+
    '<label for="newLaneAt">Posisi</label><select id="newLaneAt"><option value="0">Paling kiri</option>'+actL.map((ln,j)=>'<option value="'+(j+1)+'"'+(j===actL.length-1?' selected':'')+'>Sesudah: '+esc(laneShort(nm(ln.key)))+'</option>').join('')+'</select></div>'+
    '<div class="cbtn"><button class="sbtn ok" data-a="laneAdd">+ Tambah lane</button></div>'+
    '<p class="ws-field-hint" style="margin-top:8px">Lane baru tampil kosong. Untuk mengisinya, klik aktivitas di mode TO-BE (atau AS-IS) lalu ubah Jabatan / Unit.</p></section>';
  h += '<div class="sec">Urutan lane '+(manual?'(diatur manual)':'(otomatis: tingkat jabatan)')+'</div><div class="cbtn" style="flex-wrap:wrap;gap:6px">'+
    (manual?'<button class="sbtn" data-a="laneAuto">Kembalikan urutan otomatis</button>':'')+
    '<button class="sbtn" data-a="laneApplySop">Terapkan urutan ke semua L4 di SOP '+esc(g.l3no)+'</button></div>';
  actL.forEach((ln,j)=>{
    const sub = d.lanes[ln.i] && (d.nodes.filter(n=>n.type==='act'&&n.lane===ln.i).map(n=>byUnit?String(n.act.posX||''):String(n.act.unitX||'')).filter(Boolean)[0]||'');
    h += '<div class="rcard" data-lane="'+esc(ln.key)+'"><div class="rhead"><span class="k">'+(j+1)+' · '+ln.count+' aktivitas'+(addedOf(ln.key)?' · lane tambahan':'')+'</span><span class="ws-role-name">'+esc(laneShort(nm(ln.key)))+'</span></div>'+
      (sub?'<div class="muted" style="font-size:12px;margin:2px 0 4px">'+esc(sub)+'</div>':'');
    if(state.renamingUnit===ln.key){
      const cu=laneUnitOf(ln.i);
      h += '<div class="rgrid" style="margin-top:6px"><label for="laneNewUnit">Nama unit baru</label><input id="laneNewUnit" list="unitListR" value="'+esc(cu)+'"></div>'+
        '<label style="display:block;font-size:13px;margin:4px 0"><input type="checkbox" id="laneUnitAll"> Terapkan ke semua L4 yang memakai unit "'+esc(cu)+'"</label>'+
        '<p class="ws-field-hint" style="margin:4px 0">Mengubah Unit aktivitas di lane ini pada mode '+esc($('processEditMode').value.toUpperCase())+'.'+($('processEditMode').value==='tobe'?' Untuk semua L4: hanya L4 yang sudah punya rancangan TO-BE.':'')+'</p>'+
        '<div class="cbtn"><button class="sbtn ok" data-a="laneUnitGo">Simpan unit</button><button class="sbtn" data-a="laneUnitCancel">Batal</button></div>';
    } else if(state.renaming===ln.key){
      h += '<div class="rgrid" style="margin-top:6px"><label for="laneNewName">Nama '+(byUnit?'unit kerja':'jabatan')+' baru</label><input id="laneNewName" list="'+(byUnit?'unitListR':'posListR')+'" value="'+esc(nm(ln.key))+'"></div>'+
        '<p class="ws-field-hint" style="margin:4px 0">Mengubah '+(byUnit?'Unit':'Jabatan')+' semua aktivitas di lane ini pada mode '+esc($('processEditMode').value.toUpperCase())+'.</p>'+
        '<div class="cbtn"><button class="sbtn ok" data-a="laneRenameGo">Simpan nama</button><button class="sbtn" data-a="laneRenameCancel">Batal</button></div>';
    } else if(state.removing===ln.key){
      const opts = actL.filter(x=>x.key!==ln.key).map(x=>'<option value="'+esc(x.key)+'">'+esc(laneShort(nm(x.key)))+'</option>').join('');
      h += '<div class="rgrid" style="margin-top:6px"><label>Pindahkan '+ln.count+' aktivitas ke</label><select data-rf="target">'+opts+'</select></div>'+
        '<div class="cbtn"><button class="sbtn no" data-a="laneRemoveGo">Pindahkan &amp; hapus lane</button><button class="sbtn" data-a="laneRemoveCancel">Batal</button></div>';
    } else {
      h += '<div class="cbtn"><button class="sbtn" data-a="laneRename" title="Ganti nama lane">✎ Ganti nama</button>'+(byUnit?'':'<button class="sbtn" data-a="laneUnit" title="Ganti nama unit kerja di lane ini">✎ Ganti unit</button>')+'<button class="sbtn" data-a="laneUp"'+(j===0?' disabled':'')+' title="Geser ke kiri">⇤ Kiri</button><button class="sbtn" data-a="laneDown"'+(j===actL.length-1?' disabled':'')+' title="Geser ke kanan">Kanan ⇥</button>'+
        '<button class="sbtn no" data-a="laneRemove">'+(ln.count?'Pindahkan aktivitas…':'Hapus lane')+'</button></div>';
    }
    h += '</div>';
  });
  if(party.length){
    h += '<div class="sec">Pihak pemberi input / penerima output</div>';
    for(const ln of party){ const p0=(ln.roles&&ln.roles[0])||ln.key, hid=(R.hidden[g.l4no]||[]).includes(p0);
      h += '<div class="rcard'+(hid?' hid':'')+'" data-party="'+esc(p0)+'"><div class="rhead"><span class="k">Pihak luar</span><span class="ws-role-name">'+esc(nm(ln.key))+'</span></div><div class="cbtn"><button class="sbtn" data-a="partyHide">'+(hid?'Tampilkan lagi':'Sembunyikan lane')+'</button></div></div>'; }
    for(const p0 of (R.hidden[g.l4no]||[]).filter(p=>!party.some(ln=>((ln.roles&&ln.roles[0])||ln.key)===p)))
      h += '<div class="rcard hid" data-party="'+esc(p0)+'"><div class="rhead"><span class="k">Pihak luar (disembunyikan)</span><span class="ws-role-name">'+esc(roleInfo(p0).name)+'</span></div><div class="cbtn"><button class="sbtn" data-a="partyHide">Tampilkan lagi</button></div></div>';
  }
  const plist = new Set(); state.rows.forEach(r=>sepPos(r.posX||'').forEach(x=>plist.add(x))); Object.values(state.posmap).forEach(m=>Object.values(m).forEach(v=>sepPos(v).forEach(x=>plist.add(x))));
  h += '<datalist id="posListR">'+[...plist].sort().map(v=>'<option value="'+esc(v)+'">').join('')+'</datalist>';
  const units=new Set(); Object.values(state.unitmap).forEach(m=>Object.values(m).forEach(v=>units.add(v))); state.rows.forEach(r=>{if(r.unitX)units.add(r.unitX);});
  h += '<datalist id="unitListR">'+[...units].sort().map(v=>'<option value="'+esc(v)+'">').join('')+'</datalist>';
  $('detail').innerHTML = h;
  const nu=$('laneNewUnit'); if(nu) nu.onkeydown=e=>{ if(e.key==='Enter'){ e.preventDefault(); renameUnit(g, state.renamingUnit, nu.value, !!($('laneUnitAll')&&$('laneUnitAll').checked)); } if(e.key==='Escape'){ state.renamingUnit=null; renderRoles(); } };
  const ni=$('laneNewName'); if(ni) ni.onkeydown=e=>{ if(e.key==='Enter'){ e.preventDefault(); renameLane(g, state.renaming, ni.value); } if(e.key==='Escape'){ state.renaming=null; renderRoles(); } };
}
function roleChanged(){ saveRoles(); savePos(); saveLaneSettings(); updatePosToggle(); redraw(); renderTree(); if(state.sideMode==='roles') renderRoles(); }
function onRoleClick(a, b){
  const g = state.l4s.get(state.current); if(!g) return false;
  if(a && a.startsWith('lane') || a==='partyHide'){ try{ return onLaneClick(a,b,g); }catch(e){ toast(e.message||String(e)); return true; } }
  const card = b.closest('.rcard'), role = card && card.dataset.role;
  if(a==='roleAdd'){
    const name = ($('newRole').value||'').replace(/\s+/g,' ').trim(); if(!name){ toast('Isi nama peran dulu.'); return true; }
    if(lanesOfL4(g).some(x=>x.role.toLowerCase()===name.toLowerCase())){toast('Role sudah tersedia.');return true;}
    const unit=$('newRoleUnit').value.trim(); if(unit)(state.unitmap[g.l4no] ||= {})[name]=unit;
    const R = state.roleEdit; (R.added[g.l4no] = R.added[g.l4no] || []); if(!R.added[g.l4no].includes(name)) R.added[g.l4no].push(name);
    const pos = sepPos($('newRolePos').value).join(' / '); if(pos){ (state.posmap[g.l4no] = state.posmap[g.l4no] || {})[name] = pos; } 
    state.newRoleText = ''; roleChanged(); toast('Lane "'+name+'" ditambahkan.'); return true;
  }
  if(!role) return false;
  if(a==='roleRemove'){ const x = lanesOfL4(g).find(z=>z.role===role);
    if(x && x.kind==='added' && !x.used){ const R = state.roleEdit; R.added[g.l4no] = (R.added[g.l4no]||[]).filter(r=>r!==role); if(state.posmap[g.l4no]) delete state.posmap[g.l4no][role]; if(state.unitmap[g.l4no]) delete state.unitmap[g.l4no][role]; saveLaneSettings(); roleChanged(); return true; }
    state.removing = role; renderRoles(); return true; }
  if(a==='roleRemoveCancel'){ state.removing = null; renderRoles(); return true; }
  if(a==='roleRemoveGo'){
    const target=card.querySelector('[data-rf="target"]').value, mode=$('processEditMode').value;
    if(!target||target===role){toast('Pilih role tujuan yang berbeda.');return true;}
    if(mode==='view'){toast('Pilih mode AS-IS atau TO-BE untuk memindahkan aktivitas.');return true;}
    try{
      if(mode==='asis'&&draftGraph(g))throw Error('Selesaikan rancangan TO-BE sebelum koreksi AS-IS.');
      const draft=mode==='tobe'?materializeDraft(g):null, rows=draft?draft.rows:g.acts;
      const affected=rows.filter(r=>r.role===role);if(!affected.length)throw Error('Tidak ada aktivitas role ini pada mode aktif.');
      if(!confirm('Pindahkan '+affected.length+' aktivitas ke '+roleInfo(target).name+' pada '+mode.toUpperCase()+'?'))return true;
      const prior=snapshot();storageFailed=false;
      if(draft){affected.forEach(r=>r.role=target);saveUnifiedDraft(g,rows,draft.ids,'Pindahkan role '+role+' → '+target);undoStack.push({before:prior,after:snapshot()});}
      else{affected.forEach(r=>{r.role=target;r._asisEdited=true;});ingest(state.rows,state.fileName);select(g.l4no);undoStack.push({before:prior,after:snapshot()});}
      state.removing=null;openRoles();if(!storageFailed)toast(affected.length+' aktivitas dipindahkan. Role tetap tersedia untuk versi proses lain.');
    }catch(e){toast(e.message);}return true;
  }
  if(a==='roleHide'){ const R = state.roleEdit; const l = R.hidden[g.l4no] = R.hidden[g.l4no] || []; const i = l.indexOf(role); if(i>=0) l.splice(i,1); else l.push(role); roleChanged(); return true; }
  return false;
}
function laneOrderNow(g){ return (state.diagram?.lanes||[]).filter(l=>l.kind==='act').map(l=>l.key); }
function onLaneClick(a, b, g){
  const R = state.roleEdit; R.order = R.order || {};
  const card = b.closest('.rcard'), key = card && card.dataset.lane;
  const setOrder = list => { R.order[g.l4no] = uniq(list); };
  if(a==='laneAdd'){
    const pos = sepPos($('newLanePos').value).join(' / '), unit = ($('newLaneUnit').value||'').replace(/\s+/g,' ').trim(), byUnit = laneMode(g.l4no)==='unit';
    if(byUnit ? !unit : !pos){ toast(byUnit?'Isi unit kerja dulu.':'Isi jabatan dulu.'); return true; }
    const newKey = byUnit ? 'unit::'+unit : 'pos::'+pos;
    const order = laneOrderNow(g); if(order.includes(newKey)){ toast('Lane itu sudah ada.'); return true; }
    const name = byUnit ? unit : pos; // internal lane id
    (R.added[g.l4no] = R.added[g.l4no] || []); if(!R.added[g.l4no].includes(name)) R.added[g.l4no].push(name);
    if(pos) (state.posmap[g.l4no] = state.posmap[g.l4no] || {})[name] = pos;
    if(unit) (state.unitmap[g.l4no] = state.unitmap[g.l4no] || {})[name] = unit;
    const at = Math.max(0, Math.min(order.length, parseInt($('newLaneAt').value)||0)); order.splice(at, 0, newKey); setOrder(order);
    roleChanged(); toast('Lane "'+name+'" ditambahkan.'); return true;
  }
  if(a==='laneAuto'){ delete R.order[g.l4no]; roleChanged(); toast('Urutan lane kembali otomatis (tingkat jabatan).'); return true; }
  if(a==='laneApplySop'){
    const order = laneOrderNow(g), same = [...state.l4s.values()].filter(x=>x.l3no===g.l3no);
    if(!confirm('Terapkan urutan lane ini ke '+same.length+' subproses di SOP '+g.l3no+'? Lane yang tidak ada di subproses lain diabaikan.')) return true;
    for(const x of same) R.order[x.l4no] = order.slice(); roleChanged(); toast('Urutan diterapkan ke '+same.length+' subproses.'); return true;
  }
  if(a==='partyHide'){ const p0 = card && card.dataset.party; if(!p0) return true; const l = R.hidden[g.l4no] = R.hidden[g.l4no] || []; const i = l.indexOf(p0); if(i>=0) l.splice(i,1); else l.push(p0); roleChanged(); return true; }
  if(!key) return false;
  if(a==='laneUp' || a==='laneDown'){
    const order = laneOrderNow(g), i = order.indexOf(key), j = a==='laneUp' ? i-1 : i+1;
    if(i<0 || j<0 || j>=order.length) return true; [order[i], order[j]] = [order[j], order[i]]; setOrder(order); roleChanged(); return true;
  }
  const dropAdded = (k, keepOrder) => { const r = (R.added[g.l4no]||[]).find(x=>addedLaneKey(g.l4no,x)===k); if(r){ R.added[g.l4no] = R.added[g.l4no].filter(x=>x!==r); if(state.posmap[g.l4no]) delete state.posmap[g.l4no][r]; if(state.unitmap[g.l4no]) delete state.unitmap[g.l4no][r]; } if(!keepOrder && R.order[g.l4no]) R.order[g.l4no] = R.order[g.l4no].filter(x=>x!==k); };
  if(a==='laneRemove'){
    const idx = state.diagram.lanes.findIndex(l=>l.key===key), cnt = state.diagram.nodes.filter(n=>n.type==='act'&&n.lane===idx).length;
    if(!cnt){ if(!(R.added[g.l4no]||[]).some(x=>addedLaneKey(g.l4no,x)===key)){ toast('Lane ini berasal dari data workbook pada versi lain; pindahkan aktivitasnya di versi tersebut.'); return true; } dropAdded(key); saveLaneSettings(); roleChanged(); toast('Lane dihapus.'); return true; }
    state.removing = key; renderRoles(); return true;
  }
  if(a==='laneRemoveCancel'){ state.removing = null; renderRoles(); return true; }
  if(a==='laneRename'){ state.renaming = key; state.removing = null; renderRoles(); setTimeout(()=>{ const i=$('laneNewName'); if(i){ i.focus(); i.select(); } },0); return true; }
  if(a==='laneRenameCancel'){ state.renaming = null; renderRoles(); return true; }
  if(a==='laneUnit'){ state.renamingUnit = key; state.renaming = null; state.removing = null; renderRoles(); setTimeout(()=>{ const i=$('laneNewUnit'); if(i){ i.focus(); i.select(); } },0); return true; }
  if(a==='laneUnitCancel'){ state.renamingUnit = null; renderRoles(); return true; }
  if(a==='laneUnitGo'){ renameUnit(g, key, ($('laneNewUnit').value||''), !!($('laneUnitAll')&&$('laneUnitAll').checked)); return true; }
  if(a==='laneRenameGo'){ renameLane(g, key, ($('laneNewName').value||'')); return true; }
  if(a==='laneRemoveGo'){
    const target = card.querySelector('[data-rf="target"]').value, mode = $('processEditMode').value, byUnit = laneMode(g.l4no)==='unit';
    if(!target || target===key){ toast('Pilih lane tujuan yang berbeda.'); return true; }
    if(mode==='view'){ toast('Pilih mode AS-IS atau TO-BE untuk memindahkan aktivitas.'); return true; }
    if(mode==='asis' && draftGraph(g)) throw Error('Selesaikan rancangan TO-BE sebelum koreksi AS-IS.');
    const draft = mode==='tobe' ? materializeDraft(g) : null, rows = draft ? draft.rows : g.acts;
    const affected = rows.filter(r=>actLaneKey(g.l4no,r)===key); if(!affected.length) throw Error('Tidak ada aktivitas lane ini pada mode aktif.');
    const tVal = target.replace(/^(pos|unit)::/,''), tUnit = !byUnit ? (state.diagram.nodes.find(n=>n.type==='act'&&state.diagram.lanes[n.lane]?.key===target)?.act.unitX || ((state.unitmap[g.l4no]||{})[(R.added[g.l4no]||[]).find(x=>addedLaneKey(g.l4no,x)===target)]||'')) : '';
    if(!confirm('Pindahkan '+affected.length+' aktivitas ke '+laneShort(tVal)+' pada '+mode.toUpperCase()+'?')) return true;
    const prior = snapshot(); storageFailed = false;
    for(const r of affected){ if(byUnit) r.unitX = tVal; else { r.posX = tVal; if(tUnit) r.unitX = tUnit; } if(!draft) r._asisEdited = true; }
    if(draft){ saveUnifiedDraft(g, rows, draft.ids, 'Pindahkan lane '+key.replace(/^(pos|unit)::/,'')+' → '+tVal); }
    else { ingest(state.rows, state.fileName); select(g.l4no); }
    undoStack.push({before:prior, after:snapshot()});
    dropAdded(key, true); saveLaneSettings(); state.removing = null; roleChanged(); if(!storageFailed) toast(affected.length+' aktivitas dipindahkan ke '+laneShort(tVal)+'.'); return true;
  }
  return false;
}
/* v1.7.0: rename a lane = change Jabatan (or Unit) of every activity in it, in the active mode. */
function renameLane(g, key, raw){
  const R = state.roleEdit; R.order = R.order || {};
  const byUnit = laneMode(g.l4no)==='unit', mode = $('processEditMode').value;
  const newVal = byUnit ? String(raw).replace(/\s+/g,' ').trim() : sepPos(raw).join(' / ');
  const oldVal = key.replace(/^(pos|unit)::/,'');
  if(!newVal){ toast('Nama lane tidak boleh kosong.'); return; }
  if(newVal===oldVal){ state.renaming=null; renderRoles(); return; }
  const newKey = (byUnit?'unit::':'pos::')+newVal;
  if(laneOrderNow(g).includes(newKey)){ toast('Lane "'+laneShort(newVal)+'" sudah ada. Gunakan "Pindahkan aktivitas…" untuk menggabungkan.'); return; }
  const addedR = (R.added[g.l4no]||[]).find(x=>addedLaneKey(g.l4no,x)===key);
  const idx = state.diagram.lanes.findIndex(l=>l.key===key), cnt = state.diagram.nodes.filter(n=>n.type==='act'&&n.lane===idx).length;
  if(cnt && mode==='view'){ toast('Pilih mode AS-IS atau TO-BE dulu untuk mengganti nama lane yang berisi aktivitas.'); return; }
  if(cnt && mode==='asis' && draftGraph(g)){ toast('Selesaikan rancangan TO-BE sebelum koreksi AS-IS.'); return; }
  const order0 = (R.order[g.l4no]||laneOrderNow(g)).slice();
  const prior = snapshot(); storageFailed = false; let n = 0;
  if(cnt){
    const draft = mode==='tobe' ? materializeDraft(g) : null, rows = draft ? draft.rows : g.acts;
    const affected = rows.filter(r=>actLaneKey(g.l4no,r)===key);
    for(const r of affected){ if(byUnit) r.unitX = newVal; else r.posX = newVal; if(!draft) r._asisEdited = true; }
    n = affected.length;
    if(draft) saveUnifiedDraft(g, rows, draft.ids, 'Ganti nama lane '+oldVal+' → '+newVal);
    else { ingest(state.rows, state.fileName); select(g.l4no); }
  }
  if(addedR){ const m = byUnit ? (state.unitmap[g.l4no] ||= {}) : (state.posmap[g.l4no] ||= {}); m[addedR] = newVal; }
  // role-level fallbacks (Pemetaan Jabatan / Unit) that produced the old name
  const map = byUnit ? state.unitmap[g.l4no] : state.posmap[g.l4no];
  if(map && cnt) for(const [role,v] of Object.entries(map)) if(v===oldVal && role!==addedR && mode==='asis') map[role] = newVal;
  // keep the old key next to the new one, so the other mode (AS-IS/TO-BE) keeps its lane position too
  R.order[g.l4no] = uniq(order0.flatMap(k=>k===key?[newKey,key]:[k]));
  undoStack.push({before:prior, after:snapshot()});
  state.renaming = null; saveLaneSettings(); roleChanged();
  if(!storageFailed) toast('Lane diganti menjadi "'+laneShort(newVal)+'"'+(n?' ('+n+' aktivitas, '+mode.toUpperCase()+').':'.'));
}
function onRoleChange(e){
  const t = e.target; if(!t.dataset || !t.dataset.rf) return false; if(t.dataset.rf==='target') return true;
  const g = state.l4s.get(state.current); const role = t.closest('.rcard').dataset.role; const v = t.value.replace(/\s+/g,' ').trim();
  if(t.dataset.rf==='alias'){ if(v && v!==(roleInfo(role).orig||roleInfo(role).name)) state.roleEdit.alias[role] = v; else delete state.roleEdit.alias[role]; }
  if(t.dataset.rf==='unit'){ const m=state.unitmap[g.l4no] ||= {}; if(v) m[role]=v; else delete m[role]; saveLaneSettings(); }
  if(t.dataset.rf==='pos'){ const m = state.posmap[g.l4no] = state.posmap[g.l4no] || {}; if(v){ m[role] = sepPos(v).join(' / '); } else delete m[role]; }
  roleChanged(); return true;
}
function roleSheet(){
  const R = state.roleEdit, rows = [['Jenis','No L4','Peran','Nilai']];
  for(const [r,al] of Object.entries(R.alias)) rows.push(['Ganti nama','',r,al]);
  for(const [k,l] of Object.entries(R.added)) l.forEach(r=>rows.push(['Tambah',k,r,'']));
  for(const [k,l] of Object.entries(R.hidden)) l.forEach(r=>rows.push(['Sembunyikan',k,r,'']));
  for(const [k,l] of Object.entries(R.order||{})) (l||[]).forEach((key,i)=>rows.push(['Urutan lane',k,key,String(i+1)]));
  return rows;
}
function parseRoleSheet(wb){
  const sn = wb.SheetNames.find(n=>/^peran edit$|ubah peran|^pengaturan lane$/i.test(n)); if(!sn) return null;
  const a = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, raw:false, defval:''}); const h = (a[0]||[]).map(norm);
  const out = {alias:{}, added:{}, hidden:{}, order:{}};
  for(const r of a.slice(1)){ const g = k => String(r[h.indexOf(k)]==null?'':r[h.indexOf(k)]).trim(); const j = norm(g('jenis')), k = g('no l4'), role = g('peran'); if(!role) continue;
    if(j==='ganti nama' && g('nilai')) out.alias[role] = g('nilai'); else if(j==='tambah' && k) (out.added[k] = out.added[k]||[]).push(role); else if(j==='sembunyikan' && k) (out.hidden[k] = out.hidden[k]||[]).push(role); else if(j==='urutan lane' && k){ (out.order[k] = out.order[k]||[])[Math.max(0,(parseInt(g('nilai'))||1)-1)] = role; } }
  for(const k of Object.keys(out.order)) out.order[k]=out.order[k].filter(Boolean);
  return out;
}
function mergeRoles(x){ if(!x) return 0; const R = state.roleEdit; Object.assign(R.alias, x.alias); Object.assign(R.added, x.added); Object.assign(R.hidden, x.hidden); R.order=R.order||{}; Object.assign(R.order, x.order||{}); saveRoles(); return Object.keys(x.alias).length + Object.values(x.added).flat().length + Object.values(x.hidden).flat().length; }

/* ================= Proyek baru ================= */
const hasWork = () => state.rows.some(r=>r._asisEdited) || Object.values(state.elab).some(v=>v && v.length) || hasPos() || hasExtra() || hasRoleEdits() || analysisRecords.length>0 || changeRegister.length>0;
// Project-only storage: never clear application preferences or unrelated sites.
const PROJECT_KEYS=['data','elab','extra','roles','posmap','decmap','unitmap','lanemodes','poson','usedCodes','analysis','changeRegister'].map(k=>'flowbpa:'+k);
let projectEpoch=0;
function resetProject(){
  // Clear persisted project before changing memory; rollback on storage failure.
  const old=new Map();
  try{
    for(const k of PROJECT_KEYS) old.set(k,localStorage.getItem(k));
    for(const k of PROJECT_KEYS) localStorage.removeItem(k);
  }catch(e){
    for(const [k,v] of old) try{if(v===null)localStorage.removeItem(k);else localStorage.setItem(k,v);}catch(_){}
    toast('Reset gagal: penyimpanan browser tidak tersedia. Backup project terlebih dahulu.',9000);return false;
  }
  for(const k of Object.keys(localStorage))if(k.startsWith('bpa_bpmn_semantics_v1391_'))localStorage.removeItem(k);
  bpmnSemanticCacheKey='';bpmnSemanticCache={};
  projectEpoch++; if(state.aiCtl)state.aiCtl.abort(); state.aiCtl=null;
  analysisRecords=[]; changeRegister=[]; analysisEditing=null;
  $('analysisDlg').classList.add('hidden'); $('analysisForm').hidden=true;
  usedCodes.clear(); undoStack=[];
  Object.assign(state,{rows:[],byCode:new Map(),l4s:new Map(),warnings:[],current:null,diagram:null,fileName:'',elab:{},extra:{},roleEdit:{alias:{},added:{},hidden:{},order:{}},posmap:{},decmap:{},unitmap:{},laneModes:{},posOn:true,editing:null,newItem:null,aiMsg:null,freeText:'',removing:null,newRoleText:'',sideMode:null,zoom:1});
  $('tree').innerHTML=''; $('canvas').innerHTML=''; $('canvas').hidden=true; $('toolbar').hidden=true; $('empty').hidden=false;
  $('warnBox').hidden=true; $('detail').classList.add('hidden'); $('detail').innerHTML=''; delete $('detail').dataset.code; $('fileInfo').textContent='';
  $('search').value=''; $('search').disabled=true; $('processEditMode').value='view'; $('btnExport').disabled=true; $('btnNew').disabled=true; $('loadErr').textContent='';
  ['btnElab','btnX','btnR'].forEach(x=>$(x).classList.remove('on')); syncWorkspace(); updatePosToggle();
  try{history.replaceState(null,'',location.pathname+location.search);}catch(e){}
  return true;
}
let _newResolve = null;
function askNewProject(fileName,master=false){
  // resolves 'new' | 'keep' | 'cancel'
  $('newTitle').textContent = fileName ? (master?'Pulihkan Master Workbook?':'Import workbook ke project aktif?') : 'Mulai proyek baru?';
  $('newText').innerHTML = fileName
    ? (master?'Master Workbook <b>'+esc(fileName)+'</b> terdeteksi. Pilih Kosongkan / mulai baru untuk memulihkan isi backup tanpa mencampur catatan lama. Pilih pertahankan hanya jika memang ingin menggabungkan pekerjaan.':'Workbook <b>'+esc(fileName)+'</b> siap diimpor. Kosongkan &amp; mulai baru untuk menggunakan data Excel tanpa history project lama, termasuk jika nama file sama. Pertahankan hanya jika ingin menggabungkan data lama.')
    : 'Aplikasi akan dikosongkan: workbook, elaborasi, analisis proses, Change Register, SLA/output, dan pemetaan jabatan yang tersimpan di browser ini dihapus. Label status tetap.';
  $('newKeep').hidden = !fileName; $('newFb').innerHTML = '';
  $('newSave').disabled = !state.rows.length && !hasWork();
  $('newDlg').classList.remove('hidden');
  return new Promise(r=>{ _newResolve = r; });
}
function closeNew(v){ $('newDlg').classList.add('hidden'); if(_newResolve){ const r = _newResolve; _newResolve = null; r(v); } }

/* ================= Events ================= */
let importingWorkbook=false;
async function handleFile(file){ state._repairedShift=false;
  if(!file || importingWorkbook)return; importingWorkbook=true;
  $('loadErr').textContent = '';
  toast('Membaca '+file.name+'…');
  try{
    const buf = await file.arrayBuffer(); let res;
    try{ res = await readWorkbook(buf); }
    catch(err){ // not a BPA workbook: maybe a standalone mapping / elaboration file
      if(!state.rows.length) throw err;
      const wb = XLSX.read(buf, {type:'array'}); if(wb.SheetNames.some(n=>XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:false}).slice(0,15).some(r=>r.map(norm).includes('no l5')&&r.map(norm).includes('proses sesudah'))))throw err; const pos = parsePosSheets(wb), el = parseElabSheet(wb), ex = parseExtraSheet(wb), rl = parseRoleSheet(wb);
      if(!pos && !el && !ex && !rl) throw err;
      const np = pos ? applyPos(pos) : 0, ne = el ? mergeElab(el) : 0, nx = ex ? mergeExtra(ex) : 0; if(rl) mergeRoles(rl); saveElab(); redraw(); renderTree();
      toast([np?np+' pemetaan jabatan':'', ne?ne+' elemen elaborasi':'', nx?nx+' SLA/output':''].filter(Boolean).join(', ')+' dimuat dari '+file.name+'.', 4000); return;
    }
    if(state.rows.length || hasWork()){
      const ans = await askNewProject(file.name,res.master);
      if(ans==='cancel'){ toast('Upload dibatalkan.'); return; }
      if(ans==='new'&&!resetProject()) return;
    }
    undoStack=[]; (res.used||[]).forEach(c=>usedCodes.add(c)); const {rows, elab, pos, extra, roles, analysis, register} = res; mergeRegister(register); if(analysis) {for(const r of analysis)r.activityName=rows.find(a=>a.l5no===r.activity&&a.l4no===r.l4)?.l5||r.activityName; mergeAnalysis(analysis);} let ne = 0, np = 0; if(elab) ne = mergeElab(elab); if(pos) np = applyPos(pos); if(extra) mergeExtra(extra); if(roles) mergeRoles(roles);
    if(!pos && rows.some(r=>r.posX||r.unitX)){state.posOn=false;savePos();} ingest(rows, file.name); if(res.bpmn!==null)bpmnSemanticImport(res.bpmn); toast(state.l4s.size+' subproses berhasil dibaca.'+(state._repairedShift?' Kolom Master Workbook versi lama (v1.5.0 ke bawah) diperbaiki otomatis.':'')+(ne?' '+ne+' elemen elaborasi ikut terbaca.':'')+(np?' Pemetaan jabatan ikut terbaca.':''), 4000); }
  catch(e){ console.error(e); $('loadErr').textContent = 'Gagal membaca file: '+(e.message||e); toast('Gagal membaca file.'); }
  finally{importingWorkbook=false;}
}
$('fileIn').onchange = e => { handleFile(e.target.files[0]); e.target.value=''; };
$('fileIn2').onchange = e => { handleFile(e.target.files[0]); e.target.value=''; };
['dragover','dragenter'].forEach(ev=>document.addEventListener(ev, e=>{ e.preventDefault(); $('drop').classList.add('over'); }));
['dragleave','drop'].forEach(ev=>document.addEventListener(ev, e=>{ e.preventDefault(); $('drop').classList.remove('over'); }));
document.addEventListener('drop', e=>{ const f = e.dataTransfer && e.dataTransfer.files[0]; if(f) handleFile(f); });
$('tree').onclick = e => { const a = e.target.closest('[data-l4]'); if(a) select(a.dataset.l4); };
$('warnList').onclick = e => { const a = e.target.closest('a[data-code]'); if(a){ select(a.dataset.l4); showDetail(a.dataset.code); } };
$('search').oninput = () => renderTree();
/* Direct process editing v1.3.3: baseline corrections vs proposals. */
const usedCodes=new Set(); let undoStack=[];
function reserveCodes(rows){for(const r of rows)usedCodes.add(r.l5no);}
function rememberCodes(){return store.set('flowbpa:usedCodes',JSON.stringify([...usedCodes]));}
function snapshot(){return JSON.stringify({rows:state.rows,elab:state.elab,extra:state.extra,bpmn:bpmnSemanticData()});}
function restoreSnapshot(raw){const x=JSON.parse(raw), l4=state.current;state.elab=x.elab;state.extra=x.extra;saveElab();saveExtra();ingest(x.rows,state.fileName);bpmnSemanticImport(x.bpmn||{});select(l4);}
function undoStructural(){if(!undoStack.length)return toast('Tidak ada perubahan untuk Undo.');if(!confirm('Batalkan perubahan terakhir?'))return;storageFailed=false;const last=undoStack[undoStack.length-1];if(snapshot()!==last.after)return toast('Data telah berubah setelah edit struktur. Undo dibatasi agar perubahan lain tidak hilang.');undoStack.pop();restoreSnapshot(last.before);if(undoStack.length)undoStack[undoStack.length-1].after=snapshot();if(!storageFailed)toast('Perubahan dibatalkan.');}
// Unified draft: old proposals remain intact as audit records; snapshots identify absorbed records.
function consumedIds(rows){return new Set(rows.flatMap(r=>Array.isArray(r._consumedElab)?r._consumedElab:[]));}
function stampConsumed(rows,ids){rows.forEach(r=>delete r._consumedElab);if(rows.length)rows[0]._consumedElab=[...ids];return rows;}
function draftGraph(g){return (state.elab[g.l4no]||[]).filter(i=>i.type==='graph'&&i.status!=='rejected').slice(-1)[0];}
function materializeDraft(g){
 const graph=draftGraph(g), rows=JSON.parse(JSON.stringify(graph?JSON.parse(graph.text):g.acts));
 if(!Array.isArray(rows)||!rows.length||rows.some(r=>!r||r.l4no!==g.l4no||!r.l5no))throw Error('Snapshot TO-BE tidak valid.');
 const ids=consumedIds(rows), pending=(state.elab[g.l4no]||[]).filter(i=>i.status!=='rejected'&&STRUCT.includes(i.type)&&!ids.has(i.id));
 const outside=state.rows.filter(r=>r.l4no!==g.l4no), used=new Set([...state.rows,...rows].map(r=>r.l5no).concat([...usedCodes]));
 const by=c=>rows.find(r=>r.l5no===c), local=f=>f.kind!=='code'||!!by(f.code), map={};
 const fail=i=>{throw Error('Usulan lama '+TYPE_NAME[i.type]+' pada '+i.after+' belum dapat digabung aman. Tinjau usulan tersebut di Riwayat TO-BE.');};
 const sync=()=>{for(const r of rows){const incoming=rows.filter(a=>parseFlow(a.next).some(f=>f.kind==='code'&&f.code===r.l5no));const p=parseFlow(r.prev).filter(f=>f.kind!=='code'||!by(f.code)||incoming.some(a=>a.l5no===f.code));for(const a of incoming)if(!p.some(f=>f.kind==='code'&&f.code===a.l5no))p.push({kind:'code',code:a.l5no});r.prev=flowString(p);}};
 const before=new Set(validateWorkbookRows([...outside,...rows]).map(w=>w.code+'|'+w.msg));
 for(const i of pending.filter(i=>i.type==='rename'||i.type==='role')){const a=by(i.after);if(!a||!(i.type==='rename'?i.text:i.role))fail(i);if(i.type==='rename')a.l5=i.text;else a.role=i.role;ids.add(i.id);}
 for(const i of pending.filter(i=>i.type==='newact'||i.type==='reorder')){
  const a=by(i.after);if(!a)fail(i);
  if(i.type==='newact'){
   if(!i.text||!parseFlow(a.next).every(local)||a.decisionQuestion||hasFormalDecision(a))fail(i);
   let n=1,c;do{c=g.l4no+'.'+String(n++).padStart(2,'0');}while(used.has(c));used.add(c);map[i.id]=c;
   rows.push({...a,l5no:c,l5:i.text,role:i.role||a.role,desc:'',prev:a.l5no,next:a.next,inParty:'',outParty:'',startLabel:'',endLabel:'',symbolDocs:'',dec:'',decisionQuestion:'',decisionRole:'',inputX:'',outputX:'',docX:'',systemX:'',slaX:'',outX:'',posX:'',unitX:'',note:'',noteRole:'',iface:''});a.next=c;
  }else{
   const x=by(i.subject), f=x&&parseFlow(x.next);if(!x||x===a||f.length!==1||f[0].label||!['code','end'].includes(f[0].kind)||!f.every(local)||!parseFlow(x.prev).every(local)||!parseFlow(a.next).every(local)||outside.some(r=>[...parseFlow(r.prev),...parseFlow(r.next)].some(t=>t.kind==='code'&&t.code===x.l5no)))fail(i);
   for(const r of rows)if(r!==x)r.next=flowString(parseFlow(r.next).map(t=>t.kind==='code'&&t.code===x.l5no?{...f[0],label:t.label,note:t.note||f[0].note}:t));
   const starts=parseFlow(x.prev).filter(t=>t.kind==='start');if(starts.length){const target=by(f[0].code);if(!target)fail(i);target.prev=flowString([...parseFlow(target.prev),...starts]);x.prev=flowString(parseFlow(x.prev).filter(t=>t.kind!=='start'));}
   x.next=a.next;a.next=x.l5no;
  }ids.add(i.id);sync();
 }
 if(pending.length){const issue=validateWorkbookRows([...outside,...rows]).find(w=>!before.has(w.code+'|'+w.msg));if(issue)throw Error('Penggabungan usulan lama ditunda: '+issue.msg);}
 return {rows:stampConsumed(rows,ids),ids,map,pending};
}
function structuralRows(g,mode){return mode==='tobe'?materializeDraft(g).rows:JSON.parse(JSON.stringify(g.acts));}
function saveUnifiedDraft(g,rows,ids,reason){
 stampConsumed(rows,ids);const text=JSON.stringify(rows);if(text.length>32000)throw Error('Snapshot terlalu besar untuk satu sel Excel.');
 const list=elabOf(g.l4no),old=draftGraph(g);const data={text,reason};
 if(old&&old.status==='proposed')Object.assign(old,data);else list.push({id:newId(),l4:g.l4no,type:'graph',after:g.acts[0].l5no,status:'proposed',source:'manual',...data});
 reserveCodes(rows);rememberCodes();elabChanged();
}
function flowString(fs){ const t=f=>(f.kind==='code'?f.code:f.kind==='start'?'MULAI':f.kind==='end'?'SELESAI':f.raw)+(f.note?' ('+f.note+')':''); const out=[]; let last=null;
  for(const f of fs){ if(f.label&&last&&last.label===f.label&&f.kind!=='unknown'&&last.kind!=='unknown'){ out[out.length-1]+=' + '+t(f); } else out.push((f.label?f.label+': ':'')+t(f)); last=f; } // v1.8.0: parallel targets of one branch as "Label: A + B"
  return out.join('; '); }
function moveLinearActivity(rows,code,target,position){
 const fail=()=>{throw Error('Pindah urutan hanya tersedia untuk L4 linear tanpa percabangan, label koneksi, atau hubungan lintas L4. Gunakan Ubah aktivitas berikutnya untuk alur kompleks.');};
 const map=new Map(rows.map(r=>[r.l5no,r]));
 if(!map.has(code)||!map.has(target)||code===target||!['before','after'].includes(position))throw Error('Pilih aktivitas tujuan yang berbeda.');
 const plain=f=>!f.label&&!f.note;
 for(const r of rows){const p=parseFlow(r.prev),n=parseFlow(r.next);if(p.length!==1||n.length!==1||!p.every(plain)||!n.every(plain)||r.decisionQuestion||hasFormalDecision(r)||!(p[0].kind==='start'||p[0].kind==='code'&&map.has(p[0].code))||!(n[0].kind==='end'||n[0].kind==='code'&&map.has(n[0].code)))fail();}
 const starts=rows.filter(r=>parseFlow(r.prev)[0].kind==='start');if(starts.length!==1)fail();
 const order=[],seen=new Set();let r=starts[0];while(r){if(seen.has(r.l5no))fail();seen.add(r.l5no);order.push(r.l5no);const n=parseFlow(r.next)[0];if(n.kind==='end')break;const t=map.get(n.code);if(!t||parseFlow(t.prev)[0].code!==r.l5no)fail();r=t;}
 if(order.length!==rows.length)fail();
 const moved=order.filter(c=>c!==code);moved.splice(moved.indexOf(target)+(position==='after'?1:0),0,code);
 if(moved.every((c,i)=>c===order[i]))throw Error('Urutan aktivitas sudah sesuai.');
 return rows.map(r=>{const i=moved.indexOf(r.l5no);return {...r,prev:i?moved[i-1]:'MULAI',next:i===moved.length-1?'SELESAI':moved[i+1]};});
}
function moveFlexibleActivity(rows,code,target,position){
 const fail=m=>{throw Error(m);};
 const map=new Map(rows.map(r=>[r.l5no,{...r}]));
 if(!map.has(code)||!map.has(target)||code===target||!['before','after'].includes(position))fail('Pilih aktivitas tujuan yang berbeda.');
 const x=map.get(code), t=map.get(target);
 const plain=f=>!f.label&&!f.note;
 const outgoing=r=>{const fs=parseFlow(r.next);return fs.length===1&&plain(fs[0])&&!r.decisionQuestion&&!hasFormalDecision(r)&& (fs[0].kind==='end'||fs[0].kind==='code'&&map.has(fs[0].code))?fs[0]:null;};
 const incoming=r=>{const fs=parseFlow(r.prev);const candidates=fs.filter(f=>plain(f)&&(f.kind==='start'||f.kind==='code'&&map.has(f.code)&&outgoing(map.get(f.code))?.code===r.l5no));return candidates.length===1?candidates[0]:null;};
 const same=(a,b)=>a.kind===b.kind&&(a.kind!=='code'||a.code===b.code);
 const ref=c=>({kind:'code',code:c});
 const replacePrev=(r,old,nw)=>{const fs=parseFlow(r.prev),hits=fs.filter(f=>plain(f)&&same(f,old));if(hits.length!==1||fs.some(f=>same(f,nw)))fail('Koneksi masuk tidak unik pada '+r.l5no+'. Sesuaikan koneksi secara manual.');const result=fs.map(f=>plain(f)&&same(f,old)?nw:f);const text=flowString(result);if(JSON.stringify(parseFlow(text).map(f=>[f.kind,f.code||'',f.label||'',f.note||'']))!==JSON.stringify(result.map(f=>[f.kind,f.code||'',f.label||'',f.note||''])))fail('Format label koneksi tidak dapat dipertahankan pada '+r.l5no+'.');r.prev=text;};
 const pred=incoming(x),succ=outgoing(x);
 if(!pred||!succ||succ.code===code||pred.code===code)fail('Aktivitas asal harus memiliki satu jalur utama masuk dan keluar yang jelas. Aktivitas keputusan atau koneksi lintas L4 perlu diedit manual.');
 if(position==='before'&&succ.code===target||position==='after'&&pred.code===target)fail('Urutan aktivitas sudah sesuai.');
 // Detach only the main-path edges; other incoming decision/return edges stay attached to X.
 if(pred.kind==='code')map.get(pred.code).next=flowString([succ]);
 if(succ.kind==='code')replacePrev(map.get(succ.code),ref(code),pred);
 // Remove the old main incoming edge while preserving labelled/other incoming edges.
 const retained=parseFlow(x.prev).filter(f=>!(plain(f)&&same(f,pred)));
 x.prev=flowString(retained);x.next='';
 let left,right;
 if(position==='before'){left=incoming(t);right=ref(target);if(!left)fail('Tujuan memiliki beberapa jalur masuk atau koneksi lintas L4. Pilih posisi sesudah aktivitas pendahulu yang spesifik.');}
 else{left=ref(target);right=outgoing(t);if(!right)fail('Tujuan memiliki percabangan keputusan atau koneksi lintas L4. Pilih posisi sebelum aktivitas pada cabang yang dimaksud.');}
 if(left.code===code||right.code===code||left.kind==='start'&&right.kind==='end')fail('Posisi tujuan tidak valid.');
 if(left.kind==='code')map.get(left.code).next=flowString([ref(code)]);
 if(right.kind==='code')replacePrev(map.get(right.code),left,ref(code));
 if(retained.some(f=>same(f,left)))fail('Pemindahan akan membuat koneksi masuk ganda.');
 // Put the unlabelled main predecessor first to avoid inherited branch labels in parseFlow.
 x.prev=flowString([left,...retained]);x.next=flowString([right]);
 const result=rows.map(r=>map.get(r.l5no));
 // Ensure all non-main edges (labels, notes, external links) survive unchanged.
 const protectedEdges=rs=>rs.flatMap(r=>['prev','next'].flatMap(k=>parseFlow(r[k]).filter(f=>f.label||f.note||f.kind==='code'&&!map.has(f.code)).map(f=>JSON.stringify([r.l5no,k,f.kind,f.code||'',f.label||'',f.note||''])))).sort();
 if(JSON.stringify(protectedEdges(rows))!==JSON.stringify(protectedEdges(result)))fail('Pemindahan akan mengubah cabang berlabel atau koneksi lintas L4. Gunakan editor koneksi.');
 return result;
}
function structuralChange(rows,code,op,arg){
  const a=rows.find(r=>r.l5no===code); if(!a)throw Error('Aktivitas tidak ditemukan.');
  const local=new Set(rows.map(r=>r.l5no));
  const next=parseFlow(a.next);
  const incoming=rows.filter(r=>parseFlow(r.next).some(f=>f.kind==='code'&&f.code===code));
  const outside=state.rows.filter(r=>r.l4no!==a.l4no);
  const external=outside.some(r=>[...parseFlow(r.next),...parseFlow(r.prev)].some(f=>f.kind==='code'&&f.code===code));
  const cross=fs=>fs.some(f=>f.kind==='code'&&!local.has(f.code));
  const patches={out:[],warn:[],newCode:''};
  const outsideRefs=c=>outside.filter(r=>[...parseFlow(r.next),...parseFlow(r.prev)].some(f=>f.kind==='code'&&f.code===c));
  const sync=(r,old,newf)=>{for(const t of rows){const had=old.some(f=>f.kind==='code'&&f.code===t.l5no),has=newf.some(f=>f.kind==='code'&&f.code===t.l5no);if(had===has)continue;let p=parseFlow(t.prev).filter(f=>!(f.kind==='code'&&f.code===r.l5no));if(has)p.push({kind:'code',code:r.l5no});t.prev=flowString(p);}};
  if(op==='move'){
    rows=moveFlexibleActivity(rows,code,arg.target,arg.position);
  }else if(op==='add'){
    // v1.7.1: add before / after / inside a gateway branch; cross-L4 links move with the step they belong to.
    if(!arg.name||!arg.role)throw Error('Nama dan role wajib diisi.');
    const where=arg.where||'after';
    let n=1,c;const used=new Set([...state.rows,...rows].map(r=>r.l5no).concat([...usedCodes]));do{c=a.l4no+'.'+String(n++).padStart(2,'0');}while(used.has(c));
    const b={...a,l5no:c,l5:arg.name,role:arg.role,desc:'',prev:'',next:'',dataFrom:'',inParty:'',outParty:'',startLabel:'',endLabel:'',symbolDocs:'',dec:'',decisionQuestion:'',decisionRole:'',inputX:'',outputX:'',docX:'',systemX:'',slaX:'',outX:'',posX:arg.posX||'',unitX:arg.unitX||'',noteRole:'',iface:'',note:'',_asisEdited:true};
    rows.push(b); patches.newCode=c;
    if(where==='after'){
      // the new step takes over everything that followed, including the gateway (question and branches)
      b.prev=code; b.next=a.next; b.decisionQuestion=a.decisionQuestion||''; b.decisionRole=a.decisionRole||''; b.startLabel=''; b.endLabel=a.endLabel||'';
      sync(a,next,[]); a.next=c; a.decisionQuestion=''; a.decisionRole=''; a.endLabel=''; sync(b,[],next);
      for(const f of next) if(f.kind==='code'&&!local.has(f.code)) patches.out.push({code:f.code,field:'prev',from:code,to:c});
    }else if(where==='before'){
      // predecessors now lead to the new step (sync moves them into its Proses Sebelum)
      for(const r of incoming){ if(r===a) continue; const old=parseFlow(r.next), nf=old.map(f=>f.kind==='code'&&f.code===code?{...f,code:c}:f); r.next=flowString(nf); sync(r,old,nf); }
      // start event and triggers from other processes also move to the new step
      const pv=parseFlow(a.prev), moved=pv.filter(f=>f.kind==='start'||(f.kind==='code'&&!local.has(f.code)));
      b.prev=flowString([...moved,...parseFlow(b.prev)]); b.startLabel=a.startLabel||''; a.startLabel='';
      for(const f of moved) if(f.kind==='code') patches.out.push({code:f.code,field:'next',from:code,to:c});
      // the trigger moves to the new step, but documents from that process stay with this step
      const mc=moved.filter(f=>f.kind==='code').map(f=>f.code); if(mc.length) a.dataFrom=uniq([...String(a.dataFrom||'').split(/[;,\s]+/).filter(Boolean),...mc]).join('; ');
      b.dataFrom='';
      b.next=code; a.prev=flowString([...parseFlow(a.prev).filter(f=>!(f.kind==='start'||(f.kind==='code'&&!local.has(f.code)))),{kind:'code',code:c}]);
    }else if(/^branch:\d+$/.test(where)){
      // all flows of one branch label (e.g. "Ya: SELESAI + 12.03.01") go through the new step
      const i=+where.split(':')[1], f=next[i]; if(!f)throw Error('Cabang tidak ditemukan.');
      const grp=next.filter((x,j)=>j===i||(f.label&&x.label===f.label));
      b.prev=code; b.next=flowString(grp.map(x=>({...x,label:''})));
      const nf=[]; next.forEach(x=>{ if(!grp.includes(x)) nf.push(x); else if(x===grp[0]) nf.push({kind:'code',code:c,label:f.label,note:''}); });
      sync(a,next,nf); a.next=flowString(nf);
      for(const x of grp){
        if(x.kind==='code'){ const t=rows.find(r=>r.l5no===x.code); if(t){ const tp=parseFlow(t.prev).filter(y=>!(y.kind==='code'&&y.code===code)); if(nf.some(y=>y.kind==='code'&&y.code===x.code)) tp.push({kind:'code',code}); if(!tp.some(y=>y.kind==='code'&&y.code===c)) tp.push({kind:'code',code:c}); t.prev=flowString(tp); } else patches.out.push({code:x.code,field:'prev',from:code,to:c}); }
        if(x.kind==='end') b.endLabel=a.endLabel||'';
      }
    }else throw Error('Posisi tambah tidak dikenal.');
  }else if(op==='delete'){
    // v1.7.1: delete any step. Predecessors are joined to what followed it; a gateway on the step
    // either moves to the predecessor or one branch is kept.
    if(rows.length===1)throw Error('Aktivitas terakhir di subproses ini tidak dapat dihapus.');
    const own=next.filter(f=>!(f.kind==='code'&&f.code===code));
    const preds=incoming.filter(r=>r!==a);
    const gate=own.filter(f=>f.label).length>1;
    let succ=own;
    if(gate){
      const how=arg.gate||'';
      if(how==='pred'){
        const p=preds.length===1?preds[0]:null, pf=p?parseFlow(p.next):[];
        if(!p||pf.length!==1)throw Error('Gateway hanya bisa dipindah ke aktivitas sebelumnya bila aktivitas ini punya tepat satu pendahulu dengan satu jalur.');
        p.decisionQuestion=a.decisionQuestion||p.decisionQuestion||''; p.decisionRole=a.decisionRole||p.decisionRole||'';
        succ=own; // keep labels: predecessor becomes the gateway
        const old=pf, nf=own.map(f=>({...f})); p.next=flowString(nf); sync(p,old,nf);
        for(const f of own) if(f.kind==='code'&&!local.has(f.code)) patches.out.push({code:f.code,field:'prev',from:code,to:p.l5no});
        preds.length=0; // already reconnected
      }else if(/^keep:\d+$/.test(how)){ const f=own[+how.split(':')[1]]; if(!f)throw Error('Cabang tidak ditemukan.'); succ=[{...f,label:''}]; }
      else throw Error('Pilih apa yang terjadi dengan gateway pada aktivitas ini.');
    }
    if(!succ.length)throw Error('Aktivitas ini tidak punya kelanjutan; ubah dulu Proses Sesudah-nya.');
    for(const r of preds){
      const old=parseFlow(r.next), nf=[];
      for(const f of old){ if(f.kind==='code'&&f.code===code){ succ.forEach((s0,k)=>nf.push({...s0,label:f.label||(succ.length>1&&!gate?s0.label:''),note:k?s0.note:(f.note||s0.note)})); } else nf.push(f); }
      const seen=new Set(), ded=nf.filter(f=>{const k=f.label+'|'+f.kind+'|'+(f.code||''); if(seen.has(k))return false; seen.add(k); return true;});
      r.next=flowString(ded); sync(r,old,ded);
      for(const f of succ) if(f.kind==='code'&&!local.has(f.code)) patches.out.push({code:f.code,field:'prev',from:code,to:r.l5no});
    }
    sync(a,next,[]);
    // start event and triggers from other processes move to the first local successor
    const firstLocal=succ.map(f=>f.kind==='code'&&rows.find(r=>r.l5no===f.code&&r!==a)).find(Boolean);
    const pvMove=parseFlow(a.prev).filter(f=>f.kind==='start'||(f.kind==='code'&&!local.has(f.code)));
    if(pvMove.length){
      if(!firstLocal)throw Error('Aktivitas ini titik awal proses; tambahkan dulu aktivitas penggantinya.');
      firstLocal.prev=flowString([...parseFlow(firstLocal.prev),...pvMove]); if(a.startLabel&&!firstLocal.startLabel) firstLocal.startLabel=a.startLabel;
      for(const f of pvMove) if(f.kind==='code') patches.out.push({code:f.code,field:'next',from:code,to:firstLocal.l5no});
    }
    for(const r of outsideRefs(code)) if(!patches.out.some(x=>x.code===r.l5no)) patches.warn.push(r.l5no);
    rows=rows.filter(r=>r!==a);
  }else{
    let nf=arg.branches||parseFlow(arg.next);
    // v1.8.0: a branch can lead to a NEW gateway (question + its own branches), stored as row <code>-G<n>
    if(nf.some(f=>f.kind==='newgate')){
      const base=code.replace(/-G\d+$/,''), used=new Set([...state.rows,...rows].map(r=>r.l5no).concat([...usedCodes])); let gi=1;
      nf=nf.map(f=>{
        if(f.kind!=='newgate') return f;
        const q=String(f.question||'').trim(); if(!q) throw Error('Isi pertanyaan untuk gateway baru di cabang "'+(f.label||'')+'".');
        const sub=(f.branches||[]).filter(x=>x&&(x.kind==='end'||x.code));
        if(sub.length<2) throw Error('Gateway baru di cabang "'+(f.label||'')+'" perlu minimal 2 kondisi.');
        if(sub.some(x=>!String(x.label||'').trim())||new Set(sub.map(x=>x.label.trim().toLowerCase())).size!==sub.length) throw Error('Setiap kondisi gateway baru wajib berlabel berbeda.');
        if(sub.some(x=>/[:;\n]/.test(x.label))) throw Error('Label tidak boleh mengandung titik dua, titik koma, atau baris baru.');
        while(used.has(base+'-G'+gi)) gi++; const gc=base+'-G'+gi; used.add(gc); patches.newGates=(patches.newGates||[]).concat(gc);
        const G={...a,l5no:gc,l5:'Gateway: '+q,desc:'',prev:'',next:'',inParty:'',outParty:'',startLabel:'',endLabel:'',symbolDocs:'',dec:'',decisionQuestion:q,decisionRole:a.decisionRole||'',inputX:'',outputX:'',docX:'',systemX:'',slaX:'',outX:'',dataFrom:'',noteRole:'',iface:'',note:'',_asisEdited:true};
        const gf=sub.map(x=>({label:x.label.trim(),kind:x.kind==='end'?'end':'code',code:x.kind==='end'?undefined:x.code,note:''}));
        const oldCross=c=>next.some(x=>x.kind==='code'&&x.code===c);
        if(gf.some(x=>x.kind==='code'&&!local.has(x.code)&&!oldCross(x.code))) throw Error('Tujuan gateway baru harus aktivitas di L4 ini, Selesai, atau proses lain yang sudah terhubung.');
        for(const x of gf) if(x.kind==='code'&&!local.has(x.code)) patches.out.push({code:x.code,field:'prev',from:code,to:gc});
        rows.push(G); G.prev=code; G.next=flowString(gf); sync(G,[],gf); local.add(gc);
        return {label:f.label,kind:'code',code:gc,note:''};
      });
    }
    if(!nf.length||nf.some(f=>!['code','end'].includes(f.kind)||f.code===code)||nf.some(f=>f.kind==='code'&&!local.has(f.code)&&!next.some(x=>x.kind==='code'&&x.code===f.code)))throw Error('Gunakan kode L5 dalam L4 ini atau SELESAI; self-loop dan lintas L4 tidak didukung.');
    if(nf.some(f=>/[:;\n]/.test(f.label||'')))throw Error('Label tidak boleh mengandung titik dua, titik koma, atau baris baru.');
    if(nf.length>1&&(nf.some(f=>!f.label)||new Set(nf.map(f=>f.label.trim().toLowerCase())).size!==nf.length))throw Error('Setiap cabang wajib memiliki label, mis. Ya: 03.01.02; Tidak: 03.01.03.');
    nf=nf.flatMap(f=>[f,...(f.also||[]).map(x=>({...x,label:f.label}))]).map(({also,...f})=>f); // v1.8.0: parallel targets of one branch stay together
    a.next=flowString(nf);sync(a,next,nf); if(arg.question!==undefined){a.decisionQuestion=arg.question;if(!arg.question)a.decisionRole='';} if(isGateRow(a)){ if(nf.filter(f=>f.label).length<2) throw Error('Gateway memerlukan minimal dua kondisi. Untuk menghapusnya, pakai "Hapus gateway ini".'); if(arg.question) a.l5='Gateway: '+arg.question; }
  }
  const combined=[...outside,...rows];const before=validateWorkbookRows([...outside,...structuralRows(state.l4s.get(a.l4no),arg.mode)]);const old=new Set(before.map(w=>w.code+'|'+w.msg));const issues=validateWorkbookRows(combined).filter(w=>!old.has(w.code+'|'+w.msg));
  // v1.7.1: links to other processes are fixed in AS-IS (patches) or stay as recorded in the other L4 (TO-BE)
  const crossMsg=w=>w.l4!==a.l4no||(()=>{const m=String(w.msg).match(/(?:mengarah ke|menyebut|mencantumkan) (\d+\.\d+(?:\.\d+)?)/g)||[];return m.some(x=>!local.has(x.split(' ').pop())&&!rows.some(r=>r.l5no===x.split(' ').pop()));})();
  const hard=issues.filter(w=>!crossMsg(w));
  if(hard.length){ const h0=hard[0]; if(op==='delete'&&/tidak dapat dicapai|Proses Sebelum kosong/.test(h0.msg)) throw Error('Perubahan ditolak: aktivitas '+h0.code+' jadi tidak tersambung dari Mulai kalau '+code+' dihapus. Hapus atau sambungkan dulu aktivitas '+h0.code+', atau pilih cabang lain.'); throw Error('Perubahan ditolak: '+h0.msg); }
  rows.patches=patches;
  return rows;
}
/* v1.7.1: list of extra input / output documents in Detail Aktivitas; saved immediately. */
function mountDocsSection(code){
  const host=document.querySelector('#detail .ws-section'); if(!host) return;
  const box=document.createElement('div'); box.className='ws-section'; box.id='docsSec'; host.insertAdjacentElement('beforebegin',box);
  const render=()=>{
    const ds=(state.extra[code]&&Array.isArray(state.extra[code].docs))?state.extra[code].docs:[];
    const row=(d,i)=>'<div class="docrow" data-i="'+i+'" style="display:grid;grid-template-columns:82px 96px 1fr 28px;gap:4px;margin:4px 0;align-items:center">'+
      '<select data-df="d" aria-label="Arah"><option value="in"'+(d.d!=='out'?' selected':'')+'>Input</option><option value="out"'+(d.d==='out'?' selected':'')+'>Output</option></select>'+
      '<select data-df="k" aria-label="Jenis"><option value="object"'+(d.k!=='store'?' selected':'')+'>Dokumen</option><option value="store"'+(d.k==='store'?' selected':'')+'>Data</option></select>'+
      '<input data-df="n" aria-label="Nama dokumen" value="'+esc(d.n||'')+'" placeholder="Nama dokumen / data">'+
      '<button class="sbtn no" data-df="del" title="Hapus" style="padding:2px 6px">✕</button></div>';
    box.innerHTML='<strong class="ws-section-title">Input / output tambahan</strong><p class="ws-field-hint">Untuk dokumen selain output utama di atas. Tersimpan langsung dan tampil di samping aktivitas.</p>'+
      ds.map(row).join('')+'<div class="cbtn"><button class="sbtn" data-df="addin">＋ Input</button><button class="sbtn" data-df="addout">＋ Output</button></div>';
  };
  const list=()=>{ const x=state.extra[code]||(state.extra[code]={}); if(!Array.isArray(x.docs)) x.docs=[]; return x.docs; };
  const commit=()=>{ saveExtra(); clearTimeout(window._dt); window._dt=setTimeout(()=>{ redraw(); },250); };
  box.addEventListener('click',e=>{ const b=e.target.closest('[data-df]'); if(!b||b.tagName!=='BUTTON') return; e.preventDefault();
    const t=b.dataset.df;
    if(t==='addin'||t==='addout'){ list().push({d:t==='addin'?'in':'out',k:'object',n:''}); render(); const ins=box.querySelectorAll('input[data-df=n]'); ins[ins.length-1]?.focus(); return; }
    if(t==='del'){ const i=+b.closest('.docrow').dataset.i; list().splice(i,1); render(); commit(); }
  });
  box.addEventListener('input',e=>{ const el=e.target; if(el.dataset.df!=='n') return; const i=+el.closest('.docrow').dataset.i; const L=list(); if(L[i]){ L[i].n=el.value; commit(); } });
  box.addEventListener('change',e=>{ const el=e.target; if(!['d','k'].includes(el.dataset.df)) return; const i=+el.closest('.docrow').dataset.i; const L=list(); if(L[i]){ L[i][el.dataset.df]=el.value; commit(); } });
  render();
}
/* v1.7.1: forms for adding (before / after / in a branch) and deleting (incl. steps with a gateway) */
function flowTargetName(rows,f){ if(f.kind==='end') return 'Selesai'; if(f.kind!=='code') return f.raw||''; const r=rows.find(x=>x.l5no===f.code)||state.byCode.get(f.code); if(isGateRow(r)) return '◇ '+gateQ(r); return f.code+(r?' — '+r.l5:''); }
function laneChoices(a){ const d=state.diagram; if(!d) return []; return d.lanes.map((l,i)=>({l,i})).filter(x=>x.l.kind==='act').map(x=>({key:x.l.key,name:laneShort(roleInfo(x.l.key).name)})); }
function laneFill(key,a){
  const d=state.diagram, i=d?d.lanes.findIndex(l=>l.key===key):-1, sample=i>=0?d.nodes.find(n=>n.type==='act'&&n.lane===i&&n.act):null;
  const s=sample&&sample.act, out={role:s?s.role:a.role, posX:'', unitX:''};
  if(key.startsWith('pos::')){ out.posX=key.slice(5); out.unitX=s?String(s.unitX||''):''; }
  else if(key.startsWith('unit::')){ out.unitX=key.slice(6); out.posX=s?String(s.posX||''):''; }
  else { out.role=key; }
  if(!s){ const g=state.l4s.get(state.current), R=state.roleEdit, add=(R.added[g.l4no]||[]).find(x=>addedLaneKey(g.l4no,x)===key); if(add){ out.role=a.role; out.posX=(state.posmap[g.l4no]||{})[add]||out.posX; out.unitX=(state.unitmap[g.l4no]||{})[add]||out.unitX; } }
  return out;
}
function addFormHtml(rows,a){
  const nx=parseFlow(a.next), gate=nx.filter(f=>f.label).length>1;
  const opts=isGateRow(a)?[['before','Sebelum gateway ini']]:[['after',gate?'Sesudah aktivitas ini, sebelum gateway':'Sesudah aktivitas ini'+(nx.every(f=>f.kind==='end'||(f.kind==='code'&&!rows.some(r=>r.l5no===f.code)))?' (jadi aktivitas terakhir)':'')],['before','Sebelum aktivitas ini'+(parseFlow(a.prev).some(f=>f.kind==='start')?' (jadi aktivitas pertama)':'')]];
  if(gate||isGateRow(a)){ const seen=new Set(); nx.forEach((f,i)=>{ if(!(f.kind==='code'||f.kind==='end')||seen.has(f.label)) return; seen.add(f.label); const grp=nx.filter(x=>x.label===f.label); opts.push(['branch:'+i,'Di cabang "'+(f.label||'?')+'" (sebelum '+grp.map(x=>flowTargetName(rows,x)).join(' + ')+')']); }); }
  const lanes=laneChoices(a), curKey=state.diagram?.nodes.find(n=>n.type==='act'&&n.act&&n.act.l5no===a.l5no), cur=curKey?state.diagram.lanes[curKey.lane]?.key:'';
  return '<label for="structWhere">Posisi</label><select id="structWhere" style="width:100%">'+opts.map(([v,t])=>'<option value="'+esc(v)+'">'+esc(t)+'</option>').join('')+'</select>'+
    '<label for="structName">Nama aktivitas baru</label><textarea id="structName" style="width:100%" placeholder="mis. Memeriksa kelengkapan dokumen"></textarea>'+
    (lanes.length?'<label for="structLane">Lane (jabatan / unit)</label><select id="structLane" style="width:100%">'+lanes.map(l=>'<option value="'+esc(l.key)+'"'+(l.key===cur?' selected':'')+'>'+esc(l.name)+'</option>').join('')+'</select>':'')+
    '<input id="structRole" type="hidden" value="'+esc(a.role)+'">'+
    '<p class="muted">Kode baru memakai nomor berikutnya yang belum dipakai. Alur sebelum/sesudah disambung otomatis.</p>';
}
function deleteFormHtml(rows,a){
  const nx=parseFlow(a.next).filter(f=>!(f.kind==='code'&&f.code===a.l5no)), gate=nx.filter(f=>f.label).length>1;
  const preds=rows.filter(r=>r!==a&&parseFlow(r.next).some(f=>f.kind==='code'&&f.code===a.l5no));
  const canPred=preds.length===1&&parseFlow(preds[0].next).length===1;
  if(isGateRow(a)){ let g='<p><b>Hapus gateway "'+esc(gateQ(a))+'".</b> Jalur sebelum gateway akan lanjut ke cabang yang dipilih:</p>'; nx.forEach((f,i)=>{ g+='<label style="display:block;margin:4px 0"><input type="radio" name="delGate" value="keep:'+i+'"'+(i===0?' checked':'')+'> Cabang "'+esc(f.label||'')+'" → '+esc(flowTargetName(rows,f))+'</label>'; }); return g+'<p class="muted">Cabang lain yang jadi tidak tersambung harus dihapus atau disambung dulu.</p>'; }
  let h='<p>Aktivitas dihapus; aktivitas sebelumnya langsung disambung ke aktivitas sesudahnya.'+(parseFlow(a.prev).some(f=>f.kind==='start')?' Titik Mulai pindah ke aktivitas berikutnya.':'')+'</p>';
  if(gate){
    h+='<p><b>Aktivitas ini punya gateway'+(a.decisionQuestion?' ("'+esc(a.decisionQuestion)+'")':'')+'.</b> Pilih:</p>';
    if(canPred) h+='<label style="display:block;margin:4px 0"><input type="radio" name="delGate" value="pred" checked> Pindahkan gateway ke aktivitas sebelumnya ('+esc(preds[0].l5no)+')</label>';
    nx.forEach((f,i)=>{ h+='<label style="display:block;margin:4px 0"><input type="radio" name="delGate" value="keep:'+i+'"'+(!canPred&&i===0?' checked':'')+'> Hapus gateway, lanjut ke cabang "'+esc(f.label||'')+'" → '+esc(flowTargetName(rows,f))+'</label>'; });
    h+='<p class="muted">Cabang lain yang jadi tidak tersambung harus dihapus atau disambung dulu.</p>';
  }
  return h;
}
function applyCrossPatches(px){
  for(const x of (px&&px.out)||[]){ const r=state.rows.find(z=>z.l5no===x.code); if(!r) continue;
    const fs=parseFlow(r[x.field]); const has=fs.some(f=>f.kind==='code'&&f.code===x.to);
    const nf=fs.flatMap(f=>f.kind==='code'&&f.code===x.from?((has||!x.to)?[]:[{...f,code:x.to}]):[f]); r[x.field]=flowString(nf); r._asisEdited=true; }
}
function structuralConflict(g,code,op,mode){
  const absorbed=mode==='tobe'?materializeDraft(g).ids:new Set();const active=(state.elab[g.l4no]||[]).filter(i=>i.status!=='rejected'&&i.type!=='graph'&&!absorbed.has(i.id));
  // Legacy structural overlays are not yet materialized into snapshots. Never silently discard them.
  const legacy=active.find(i=>['newact','reorder'].includes(i.type));
  if(legacy)return 'Konflik: '+(legacy.type==='newact'?'aktivitas tambahan':'perubahan urutan')+' pada '+legacy.after+' masih berupa elaborasi lama. Tinjau usulan ini sebelum mengubah struktur L4.';
  const hit=active.find(i=>op==='move'&&['gate','declane'].includes(i.type)||op==='delete'&&(i.after===code||i.subject===code||i.noTarget===code)||op==='connector'&&i.after===code&&['gate','declane'].includes(i.type)||op==='add'&&i.after===code&&['gate','declane'].includes(i.type));
  return hit?'Konflik pada '+code+': '+(TYPE_NAME[hit.type]||hit.type)+' masih terkait aktivitas ini. Tinjau usulan terkait; usulan lain tetap dapat dipertahankan.':'';
}
function structuralEditor(code,op,style){
  const mode=$('processEditMode').value,g=state.l4s.get(state.current);let rows;try{rows=structuralRows(g,mode);}catch(e){return toast(e.message);}const a=rows.find(r=>r.l5no===code);if(!a)return;const absorbed=consumedIds(rows);
  const existing=parseFlow(a.next), branching=existing.length>1;
  const simple=op==='connector'&&style==='next', decision=op==='connector'&&style==='decision';
  if(simple&&branching&&!confirm('Aktivitas ini memiliki beberapa jalur. Ubah menjadi satu jalur?'))return;
  $('detail').innerHTML='<button class="tb x" data-a="close">✕</button><h3>'+({move:'Pindahkan urutan aktivitas',add:'Tambah aktivitas',delete:'Hapus aktivitas',connector:simple?'Ubah aktivitas berikutnya':(isGateRow(a)||branching?'Edit gateway':'Tambahkan gateway / keputusan')}[op])+'</h3><p>'+esc(code+' · '+a.l5)+'</p>'+(op==='move'?'<p class="muted">Koneksi sebelum dan sesudah akan disesuaikan otomatis. Kode L5 tetap. Dapat memindahkan aktivitas melewati beberapa aktivitas dan antar-swimlane. Jalur utama asal dan tujuan harus jelas; cabang keputusan dan koneksi lintas L4 tetap dipertahankan.</p><label>Posisi baru</label><select id="movePosition"><option value="before">Sebelum</option><option value="after">Sesudah</option></select><label>Aktivitas tujuan</label><select id="moveTarget"><option value="">Pilih aktivitas</option>'+rows.filter(r=>r!==a).map(r=>'<option value="'+esc(r.l5no)+'">'+esc(r.l5no+' — '+r.l5)+'</option>').join('')+'</select>':op==='add'?addFormHtml(rows,a):op==='connector'?(simple?'<label>Setelah aktivitas ini selesai, lanjut ke mana?</label>':'<label>Pertanyaan keputusan</label><input id="decisionQuestion" style="width:100%" placeholder="Contoh: Apakah dokumen lengkap?" value="'+esc(a.decisionQuestion||(branching?decQuestion(existing):''))+'"><p class="muted">Tentukan tujuan untuk setiap kondisi.</p>')+'<div id="branchRows"></div>'+(!simple?'<button class="sbtn" id="branchAdd">+ Tambah kondisi</button>':''):deleteFormHtml(rows,a))+'<div class="cbtn"><button class="sbtn ok" id="structSave">Simpan '+(mode==='asis'?'koreksi':'usulan')+'</button><button class="sbtn" id="structCancel">Batal</button></div>';
  if(op==='connector'){
    const host=$('branchRows');
    // v1.8.0: target list = activities, gateways (◇), Selesai, existing links to other L4; decision mode can create a new gateway in a branch
    const crossT=uniq(existing.filter(f=>f.kind==='code'&&!rows.some(r=>r.l5no===f.code)).map(f=>f.code));
    const targetOpts=(sel,allowNew)=>'<option value="" disabled'+(!sel?' selected':'')+'>Pilih tujuan</option>'+[...rows.filter(r=>r!==a).map(r=>[r.l5no,isGateRow(r)?'◇ Gateway '+r.l5no.split('-').pop()+': '+gateQ(r):r.l5no+' — '+r.l5]),...crossT.map(c=>[c,'→ '+c+(state.byCode.get(c)?' — '+state.byCode.get(c).l5:'')+' (proses lain)']),['SELESAI','Selesai']].map(([c,n])=>'<option value="'+esc(c)+'"'+(c===sel?' selected':'')+'>'+esc(n)+'</option>').join('')+(allowNew?'<option value="__NEWGATE__">◇ Gateway baru setelah cabang ini…</option>':'');
    const subRow=(host2,f)=>{const d=document.createElement('div');d.className='subBranch';d.style.cssText='display:grid;grid-template-columns:1fr 1.6fr 28px;gap:4px;margin:4px 0';const t=f.kind==='end'?'SELESAI':f.code||'';d.innerHTML='<input class="subLabel" placeholder="Kondisi" value="'+esc(f.label||'')+'"><select class="subTarget">'+targetOpts(t,false)+'</select><button class="sbtn no" style="padding:2px 6px" title="Hapus kondisi">✕</button>';d.querySelector('button').onclick=e=>{e.preventDefault();if(host2.children.length<=2)return toast('Gateway memerlukan minimal dua kondisi.');d.remove();};host2.appendChild(d);};
    const addBranch=f=>{const div=document.createElement('div');div.className='branchRow';div.dataset.note=f.note||'';div.dataset.also=JSON.stringify(f.also||[]);div.style.margin='10px 0';const target=f.kind==='end'?'SELESAI':f.code||'';
      const alsoTxt=(f.also||[]).map(x=>x.kind==='end'?'Selesai':x.code).join(' + ');
      div.innerHTML=(simple?'<input class="branchLabel" type="hidden" value="">':'<label>Jika kondisi:</label><input class="branchLabel" style="width:100%" placeholder="Ya / Tidak" value="'+esc(f.label||'')+'">')+'<label>Lanjut ke:</label><select class="branchTarget" style="width:100%">'+targetOpts(target,!simple)+'</select>'+(alsoTxt?'<p class="ws-field-hint" style="margin:2px 0">Juga bersamaan ke: '+esc(alsoTxt)+'</p>':'')+
        '<div class="newGate" style="display:none;border-left:3px solid #c9a227;padding:6px 0 6px 10px;margin:6px 0"><label>Pertanyaan gateway baru</label><input class="newGateQ" style="width:100%" placeholder="mis. Perlu persetujuan Direksi?"><div class="subRows"></div><button class="sbtn subAdd" style="margin-top:2px">+ kondisi</button></div>'+
        (!simple?'<button class="sbtn no delBranch" style="margin-top:4px">Hapus kondisi</button>':'');
      const sel=div.querySelector('.branchTarget'), ng=div.querySelector('.newGate'), subs=div.querySelector('.subRows');
      sel.onchange=()=>{ const on=sel.value==='__NEWGATE__'; ng.style.display=on?'block':'none'; if(on&&!subs.children.length){ subRow(subs,{label:'Ya'}); subRow(subs,{label:'Tidak'}); div.querySelector('.newGateQ').focus(); } };
      div.querySelector('.subAdd').onclick=e=>{e.preventDefault();subRow(subs,{});};
      if(!simple)div.querySelector('.delBranch').onclick=()=>{if(host.children.length<=2)return toast('Keputusan memerlukan minimal dua kondisi.');div.remove();};host.appendChild(div);};
    const grouped=[]; for(const f of existing){ const g=f.label&&grouped.find(x=>x.label===f.label); if(g) (g.also=g.also||[]).push(f); else grouped.push({...f}); }
    if(simple)addBranch(existing.length===1?existing[0]:{});
    else{if(branching)grouped.forEach(addBranch);else{addBranch({...existing[0],label:'Ya'});addBranch({label:'Tidak'});}$('branchAdd').onclick=()=>addBranch({});}
  }
 $('detail').classList.remove('hidden');$('structCancel').onclick=()=>{ if(isGateRow(a)){ $('detail').classList.add('hidden'); state.sideMode=null; } else directEditor(code); };
  if(isGateRow(a)&&op==='connector'){ // v1.8.0: gateway panel actions
    const bar=document.createElement('div'); bar.className='cbtn'; bar.style.marginTop='12px';
    bar.innerHTML='<button class="sbtn" id="gateAddAct">＋ Tambah aktivitas di cabang / sebelum gateway</button><button class="sbtn no" id="gateDel">Hapus gateway ini</button>';
    $('structSave').parentElement.insertAdjacentElement('afterend',bar);
    $('gateAddAct').onclick=()=>structuralEditor(code,'add'); $('gateDel').onclick=()=>structuralEditor(code,'delete');
  }
  $('structSave').onclick=()=>{const oldError=$('structError');if(oldError)oldError.remove();try{
    const conflict=structuralConflict(g,code,op,mode);
    if(conflict)throw Error(conflict);
    if(mode==='asis'&&(state.elab[g.l4no]||[]).some(i=>i.type==='graph'&&i.status!=='rejected'))throw Error('Selesaikan usulan struktur TO-BE sebelum mengubah AS-IS.');
    if(op==='connector'&&[...$('branchRows').children].some(d=>!d.querySelector('.branchTarget').value||[...d.querySelectorAll('.subTarget')].some(x=>d.querySelector('.branchTarget').value==='__NEWGATE__'&&!x.value)))throw Error('Pilih tujuan untuk setiap jalur.');
    if(decision&&!$('decisionQuestion').value.trim())throw Error('Isi pertanyaan keputusan.');
    const laneSel=$('structLane')?.value||'', laneInfo=laneSel?laneFill(laneSel,a):{};
    const arg={mode,where:$('structWhere')?.value,gate:(document.querySelector('input[name=delGate]:checked')||{}).value,posX:laneInfo.posX,unitX:laneInfo.unitX,target:$('moveTarget')?.value,position:$('movePosition')?.value,question:op==='connector'?(simple?'':$('decisionQuestion').value.trim()):undefined,name:$('structName')?.value.trim(),role:(laneInfo.role||$('structRole')?.value.trim()),branches:op==='connector'?[...$('branchRows').children].map(d=>{const c=d.querySelector('.branchTarget').value, lab=d.querySelector('.branchLabel').value.trim();
      if(c==='__NEWGATE__') return {label:lab,kind:'newgate',question:d.querySelector('.newGateQ').value.trim(),branches:[...d.querySelectorAll('.subBranch')].map(x=>{const t=x.querySelector('.subTarget').value;return {label:x.querySelector('.subLabel').value.trim(),kind:t==='SELESAI'?'end':'code',code:t};})};
      let also=[]; try{also=JSON.parse(d.dataset.also||'[]');}catch(e){} return {label:lab,kind:c==='SELESAI'?'end':'code',code:c,note:d.dataset.note||'',also};}):undefined};const result=structuralChange(rows,code,op,arg);
    const preview=op==='move'?result.filter(r=>{const old=rows.find(x=>x.l5no===r.l5no);return old&&(old.prev!==r.prev||old.next!==r.next);}).map(r=>{const old=rows.find(x=>x.l5no===r.l5no);return r.l5no+'\n  Sebelum: '+old.prev+' → '+old.next+'\n  Sesudah: '+r.prev+' → '+r.next;}).join('\n\n'):'';
    const px=result.patches||{out:[],warn:[]}, cross=[...new Set([...px.out.map(x=>x.code),...px.warn])];
    const crossNote=cross.length?(mode==='asis'?'\n\nHubungan dengan proses lain ('+cross.join(', ')+') ikut disesuaikan.':'\n\nCatatan: proses lain ('+cross.join(', ')+') di AS-IS masih mencatat hubungan ke '+code+'. Sesuaikan saat TO-BE diterima.'):'';
    if(!confirm('Simpan perubahan struktur '+mode.toUpperCase()+'?'+crossNote+(preview?'\n\nTinjau koneksi yang berubah:\n'+preview+'\n\nPastikan logika bisnis dan jalur keputusan tetap sesuai.':'')))return;
    const prior=snapshot();storageFailed=false;reserveCodes(result);rememberCodes();
    if(mode==='asis'){result.forEach(r=>r._asisEdited=true);if(op==='delete'){delete state.extra[code];saveExtra();}applyCrossPatches(px);ingest([...state.rows.filter(r=>r.l4no!==g.l4no),...result],state.fileName);select(g.l4no);}
    else{saveUnifiedDraft(g,result,absorbed,'Perubahan struktur: '+op+' '+code);}
    $('detail').classList.add('hidden');state.sideMode=null;undoStack.push({before:prior,after:snapshot()});if(!storageFailed)toast(op==='add'&&px.newCode?'Aktivitas '+px.newCode+' ditambahkan.':op==='delete'?(isGateRow(a)?'Gateway '+code+' dihapus.':'Aktivitas '+code+' dihapus, alur disambung.'):(px.newGates&&px.newGates.length)?'Gateway baru '+px.newGates.join(', ')+' ditambahkan. Klik dua kali belah ketupatnya untuk mengedit.':'Perubahan struktur tersimpan.');else toast('Perubahan hanya di memori. Export Excel sekarang.',9000);
  }catch(e){const msg=document.createElement('p');msg.id='structError';msg.setAttribute('role','alert');msg.style.cssText='color:#a21d32;font-weight:600;margin:10px 0;';msg.textContent=e.message;$('structSave').parentElement.insertAdjacentElement('beforebegin',msg);toast(e.message,9000);}};
}

/* v1.3.5: contextual annotations, using the existing elaboration schema. */
const ANNOT_TYPES=['hold','doc','datain','declane','gate'];
function annotationSection(g,code,mode){
  if(!g)return '';
  const items=(state.elab[g.l4no]||[]).filter(i=>ANNOT_TYPES.includes(i.type)&&i.after===code&&i.status!=='rejected');
  const editable=mode==='tobe';
  return '<div class="ws-section"><strong class="ws-section-title">Anotasi flowchart</strong><p class="ws-field-hint">'+(editable?'Tambahkan simbol pada rancangan TO-BE. Keputusan baru dibuat melalui Ubah alur proses.':'Anotasi adalah bagian dari TO-BE. Pilih mode TO-BE untuk mengubahnya.')+'</p>'+items.map(i=>'<div class="card" data-id="'+esc(i.id)+'"><span class="chip t-'+i.type+'">'+esc(TYPE_NAME[i.type])+'</span> <span class="after">'+esc(STATUS_NAME[i.status])+'</span><div class="ctext">'+esc(i.text||'')+'</div>'+(editable?'<div class="cbtn"><button class="sbtn" data-a="annEdit">Ubah</button><button class="sbtn no" data-a="annDelete">Hapus</button></div>':'')+'</div>').join('')+(editable?'<div class="cbtn"><button class="sbtn" data-a="annAdd">＋ Tambah anotasi</button></div><div id="annForm"></div>':'')+'</div>';
}
function annotationForm(g,code,it){
  const rows=structuralRows(g,'tobe'), r=rows.find(x=>x.l5no===code);if(!r)return;
  const types=['hold','doc','datain',...(hasFormalDecision(r)?['declane']:[]),...(it&&!['hold','doc','datain'].includes(it.type)?[it.type]:[])];
  const type=it?it.type:types[0];const opt=(v,t,sel)=>'<option value="'+esc(v)+'"'+(v===sel?' selected':'')+'>'+esc(t)+'</option>';
  const roles=uniq([...rows.map(x=>x.role),...(state.roleEdit.added[g.l4no]||[])]);
  $('annForm').innerHTML='<div class="card ef" data-ann-id="'+esc(it?.id||'')+'"><label>Jenis anotasi</label><select id="annType">'+uniq(types).map(t=>opt(t,TYPE_NAME[t],type)).join('')+'</select><div id="annFields"></div><div class="cbtn"><button class="sbtn ws-primary" data-a="annSave">Simpan anotasi</button><button class="sbtn" data-a="annCancel">Batal</button></div></div>';
  const fields=()=>{const t=$('annType').value;const v=it&&it.type===t?it:{};$('annFields').innerHTML=(t==='declane'?'':'<label>Keterangan</label><input id="annText" value="'+esc(v.text||'')+'">')+(t==='hold'?'<label>Label cabang</label><input id="annLabel" value="'+esc(v.label||'Hold')+'">':'')+(['gate','declane'].includes(t)?'<label>Kolom keputusan</label><select id="annRole">'+(t==='gate'?opt('','Sama dengan aktivitas',v.role||''):'')+roles.map(x=>opt(x,roleInfo(x).name,v.role)).join('')+'</select>':'')+(t==='gate'?'<label>Bila tidak lolos, kembali ke</label><select id="annTarget">'+rows.map(x=>opt(x.l5no,x.l5no+' '+x.l5,v.noTarget)).join('')+'</select><label>Label cabang tidak lolos</label><input id="annNoLabel" value="'+esc(v.noLabel||'Tidak')+'">':'');};
  $('annType').onchange=fields;fields();$('annForm').scrollIntoView({block:'nearest'});
}
function handleAnnotation(a,b,g){
  if(!g||$('processEditMode').value!=='tobe'||state.sideMode!=='direct')return;
  const code=$('detail').dataset.code,list=elabOf(g.l4no),id=b.closest('[data-id]')?.dataset.id,it=list.find(x=>x.id===id);
  if(a==='annAdd'||a==='annEdit')return annotationForm(g,code,a==='annEdit'?it:null);
  if(a==='annCancel'){ $('annForm').innerHTML='';return; }
  const prior=snapshot();
  if(a==='annDelete'){if(!it||!confirm('Hapus anotasi ini dari TO-BE?'))return;list.splice(list.indexOf(it),1);}
  else if(a==='annSave'){
    const form=b.closest('[data-ann-id]'),existing=list.find(x=>x.id===form.dataset.annId),type=$('annType').value;
    const text=($('annText')?.value||'').trim(),role=$('annRole')?.value||'';
    if(type!=='declane'&&!text)return toast('Keterangan wajib diisi.');
    if(type==='declane'&&!role)return toast('Kolom keputusan wajib dipilih.');
    const n={id:existing?.id||newId(),l4:g.l4no,after:code,type,text:type==='declane'?'Keputusan setelah '+code+' → kolom '+roleInfo(role).name:type==='gate'&&!text.endsWith('?')?text+'?':text,status:existing?.status||'accepted',source:existing?.source||'manual',reason:existing?.reason||''};
    if(type==='hold')n.label=$('annLabel').value.trim()||'Hold';
    if(role)n.role=role;
    if(type==='gate'){n.noTarget=$('annTarget').value;n.noLabel=$('annNoLabel').value.trim()||'Tidak';}
    if(existing){if(existing.source==='ai')n.edited=true;list[list.indexOf(existing)]=n;}else list.push(n);
  }else return;
  storageFailed=false;elabChanged();undoStack.push({before:prior,after:snapshot()});directEditor(code);if(!storageFailed)toast('Anotasi TO-BE tersimpan.');
}
function renderLegacyAi(){
 const g=state.l4s.get(state.current);if(!g)return;
 const busy=!!state.aiCtl;
 let h='<button class="tb x" data-a="close">✕</button><div class="ws-panel-title">Usulan AI (opsional)</div><p class="muted">Fitur existing. Usulan ditinjau melalui Riwayat TO-BE.</p>';
 if(!apiKey())h+='<p class="muted">API key Anthropic diperlukan.</p><button class="sbtn" data-a="aiSettings">Pengaturan AI</button>';
 else h+=(busy?'<button class="sbtn" data-a="stopAi">Hentikan</button>':'<button class="pbtn" data-a="ai">Usulkan dengan AI</button>')+'<textarea id="freeTxt" placeholder="Instruksi usulan">'+esc(state.freeText||'')+'</textarea><button class="sbtn" data-a="free"'+(busy?' disabled':'')+'>Tambahkan usulan</button>';
 h+='<div id="aiStatus">'+esc(state.aiMsg?.text||'')+'</div><button class="sbtn" data-a="openElab">Buka Riwayat TO-BE</button>';$('detail').innerHTML=h;
}
function directEditor(code){
  const mode=$('processEditMode').value, base=state.byCode.get(code), group=state.l4s.get(state.current);let draft;try{if(mode==='tobe'&&group)draft=materializeDraft(group);}catch(e){return toast(e.message);}if(draft&&draft.map[code])code=draft.map[code];const r=draft?draft.rows.find(a=>a.l5no===code):base; if(!r || mode==='view') return showDetail(code);
  const g=state.l4s.get(r.l4no), list=state.elab[r.l4no]||[];
  const latest=type=>list.filter(i=>i.type===type && i.after===code && i.status!=='rejected').slice(-1)[0];
  const rn=null, ro=null;
  const roles=sortRoles(uniq([...g.acts.map(a=>a.role),...(mode==='tobe'?structuralRows(g,mode):g.acts).map(a=>a.role),...(state.roleEdit.added[g.l4no]||[]),r.role,ro&&ro.role].filter(Boolean)));
  const perspective=laneMode(g.l4no), primary=perspective==='position'?'directPosition':perspective==='unit'?'directUnit':'directRole';
  const draftRows=mode==='tobe'?draft.rows:g.acts;
  const optionsFor=(field,mapping)=>uniq([...state.rows.map(a=>a[field]),...draftRows.map(a=>a[field]),...Object.values(mapping[g.l4no]||{}),r[field]].filter(Boolean)).sort((a,b)=>a.localeCompare(b));
  const laneField=(id,label,field,mapping)=>{const effective=r[field]||(mapping[g.l4no]||{})[r.role]||'';return '<label for="'+id+'">'+label+'</label><input id="'+id+'" list="'+id+'List" value="'+esc(effective)+'" placeholder="Kosong = mapping role"><datalist id="'+id+'List">'+optionsFor(field,mapping).map(x=>'<option value="'+esc(x)+'"></option>').join('')+'</datalist>';};
  const roleField='<label for="directRole">Role</label><select id="directRole">'+roles.map(x=>'<option value="'+esc(x)+'"'+(x===r.role?' selected':'')+'>'+esc(roleInfo(x).name)+'</option>').join('')+'</select>';
  const fields={directRole:roleField,directPosition:laneField('directPosition','Jabatan','posX',state.posmap),directUnit:laneField('directUnit','Unit','unitX',state.unitmap)};
  // v1.6.4: output type (Dokumen / Data) + name, written to columns U / W.
  const noSym=String(r.docX||'').trim()==='-', autoOut=!noSym && !bpmnDataKind(r) && !!cleanOut((state.extra[code]||{}).out);
  const curKind=noSym?null:(bpmnDataKind(r)||(autoOut?'object':null)), curRaw=noSym?'':String((curKind==='store'?r.systemX:r.docX)||'').trim(), curName=curRaw.toUpperCase()==='X'?'':curRaw;
  const outFields='<label for="directOutKind">Jenis dokumen / data</label><select id="directOutKind"><option value="">Tidak ada simbol</option><option value="object"'+(curKind==='object'?' selected':'')+'>Dokumen (ikon kertas)</option><option value="store"'+(curKind==='store'?' selected':'')+'>Data di sistem (ikon silinder)</option></select>'+
    '<label for="directOutName">Nama dokumen / data</label><input id="directOutName" value="'+esc(curName)+'" placeholder="'+esc(String(r.outputX||'').split(/\n/)[0]||'mis. SK Pengangkatan / Data Kandidat')+'">'+
    '<p class="ws-field-hint">'+(autoOut?'Saat ini tampil otomatis sebagai Dokumen dari kolom Output. ':'')+'Kosongkan nama untuk memakai isi kolom Output. Disimpan ke kolom Dokumen (U) / Sistem/Aplikasi (W); "Tidak ada simbol" ditulis sebagai "-" di kolom Dokumen.</p>';
  const assignmentFields=fields[primary]+'<p class="ws-field-hint">Atribut utama mengikuti perspektif swimlane aktif. Perubahan atribut lain tidak otomatis mengubah lane ini.</p><details><summary>Atribut organisasi lainnya</summary>'+Object.keys(fields).filter(k=>k!==primary).map(k=>fields[k]).join('')+'</details>';
  state.sideMode='direct'; ['btnElab','btnX','btnR'].forEach(x=>$(x).classList.remove('on'));
  $('detail').dataset.code=code;
  const df=(k,v)=>v?'<dt>'+esc(k)+'</dt><dd>'+esc(v)+'</dd>':'';
  const extra=state.extra[code]||{};
  $('detail').innerHTML='<button class="tb x" data-a="close" aria-label="Tutup detail">✕</button><div class="ws-panel-title">Detail Aktivitas</div><div class="dcode">'+esc(code)+' · '+(mode==='asis'?'AS-IS':'TO-BE')+'</div><h3>'+esc(rn?rn.text:r.l5)+'</h3><p class="muted">'+(mode==='asis'?'Koreksi proses existing.':'Rancang perbaikan tanpa mengubah AS-IS.')+'</p><div class="ws-detail-fields"><label for="directName">Nama aktivitas</label><textarea id="directName" rows="3">'+esc(rn?rn.text:r.l5)+'</textarea>'+assignmentFields+outFields+'</div><div class="cbtn"><button class="sbtn ws-primary" id="directSave">Simpan '+(mode==='asis'?'koreksi':'usulan')+'</button><button class="sbtn" id="directCancel">Batal</button></div><div class="ws-section"><details><summary>Informasi aktivitas</summary><dl>'+df('Uraian',r.desc)+df('Jabatan',r.posX)+df('Unit',r.unitX)+df('Role keputusan',r.decisionRole)+df('Input',r.inputX)+df('Output',r.outputX)+df('Dokumen',r.docX)+df('Sistem/Aplikasi',r.systemX)+df('Proses sebelum',r.prev)+df('Proses sesudah',r.next)+df('SLA',extra.sla)+df('Dokumen / output',extra.out)+df('Catatan',extra.note||r.note)+df('Pihak pemberi input',r.inParty)+df('Pihak penerima output',r.outParty)+df('Decision Ref',r.dec)+df('Catatan peran',r.noteRole)+df('Dasar antarmuka',r.iface)+'</dl><button class="sbtn" id="workspaceNotes">Tambah / edit dokumen & SLA</button></details></div><div class="ws-section"><strong class="ws-section-title">Ubah alur proses</strong><div class="ws-detail-actions"><button class="sbtn" id="structAdd">＋ Tambah aktivitas (sebelum / sesudah / di cabang)</button><button class="sbtn" id="structMove">↕ Pindahkan urutan aktivitas</button><button class="sbtn" id="structConn">→ Ubah aktivitas berikutnya</button><button class="sbtn" id="structDecision">◇ Tambahkan / edit keputusan</button><button class="sbtn no" id="structDel">Hapus aktivitas</button></div></div>';
  $('detail').insertAdjacentHTML('beforeend',annotationSection(g,code,mode));
  $('workspaceNotes').onclick=()=>openExtra(code);
  mountDocsSection(code);
  $('structMove').onclick=()=>structuralEditor(code,'move');$('structAdd').onclick=()=>structuralEditor(code,'add');$('structConn').onclick=()=>structuralEditor(code,'connector','next');$('structDecision').onclick=()=>structuralEditor(code,'connector','decision');$('structDel').onclick=()=>structuralEditor(code,'delete');
  $('detail').classList.remove('hidden');
  $('directCancel').onclick=()=>{$('detail').classList.add('hidden');state.sideMode=null;};
  $('directSave').onclick=()=>{
    const graph=list.filter(i=>i.type==='graph'&&i.status!=='rejected').slice(-1)[0];
    const name=$('directName').value.trim(), role=$('directRole').value; const posValue=$('directPosition').value.trim(), unitValue=$('directUnit').value.trim(); const posX=posValue===(r.posX||(state.posmap[g.l4no]||{})[r.role]||'')?r.posX:posValue, unitX=unitValue===(r.unitX||(state.unitmap[g.l4no]||{})[r.role]||'')?r.unitX:unitValue; const outKind=$('directOutKind').value, outName=$('directOutName').value.replace(/\s+/g,' ').trim();
    if(outKind && !outName && !String(r.outputX||'').trim() && !cleanOut((state.extra[code]||{}).out)){toast('Isi nama dokumen/data (kolom Output aktivitas ini kosong).');return;}
    const docX=outKind==='object'?(outName||'X'):outKind===''?'-':'', systemX=outKind==='store'?(outName||'X'):'';
    const sameOut=String(r.docX||'')===docX && String(r.systemX||'')===systemX;
    const unchanged=name===r.l5&&role===r.role&&posX===r.posX&&unitX===r.unitX&&sameOut; if(!name||!role){toast('Nama dan role wajib diisi.');return;}
    if(mode==='asis'){
      if(graph){toast('Selesaikan usulan struktur TO-BE sebelum koreksi AS-IS.');return;}
      if(unchanged){toast('Tidak ada perubahan.');return;}
      if(!confirm('Simpan koreksi ke data AS-IS? Workbook Excel asli tidak ditimpa.'))return;
      storageFailed=false;r.l5=name;r.role=role;r.posX=posX;r.unitX=unitX;r.docX=docX;r.systemX=systemX;r._asisEdited=true;
      ingest(state.rows,state.fileName); select(r.l4no); if(!storageFailed)toast('Koreksi AS-IS tersimpan.');
    }else{
      if(unchanged&&!draft.pending.length)return toast('Tidak ada perubahan.');
      const prior=snapshot();try{storageFailed=false;r.l5=name;r.role=role;r.posX=posX;r.unitX=unitX;r.docX=docX;r.systemX=systemX;saveUnifiedDraft(g,draft.rows,draft.ids,'Edit aktivitas '+code);undoStack.push({before:prior,after:snapshot()});if(!storageFailed)toast('Rancangan TO-BE tersimpan.');}catch(e){return toast(e.message);}

    }
    directEditor(code);
  };
}
function syncWorkspace(){
  const m=$('processEditMode').value;
  $('workspaceHint').textContent=m==='asis'?'Koreksi data existing':m==='tobe'?'Rancangan tanpa mengubah AS-IS':'Lihat tanpa mengubah data';
  $('workspaceOrg').disabled=!state.current; $('btnAnalysis').disabled=!state.current;
  $('btnUndoStructural').hidden=m==='view';
}
$('workspaceOrg').onclick=()=>{$('btnR').click();};
$('workspaceNav').onclick=()=>{const side=document.querySelector('.side');const hidden=side.classList.toggle('ws-collapsed');$('workspaceNav').setAttribute('aria-expanded',String(!hidden));};
document.addEventListener('click',e=>{document.querySelectorAll('.ws-menu[open]').forEach(d=>{if(!d.contains(e.target)||e.target.closest('button'))d.open=false;});});
$('processEditMode').onchange=()=>{ $('detail').classList.add('hidden');state.sideMode=null; ['btnElab','btnX','btnR'].forEach(x=>$(x).classList.remove('on'));redraw(); };
$('btnUndoStructural').onclick=undoStructural;
$('canvas').ondblclick = e => { const g = e.target.closest('[data-code]');
  { const dm=e.target.closest('.dmd'); if(dm && dm.dataset.dk==='formal'){ e.preventDefault(); closeLanePop&&closeLanePop(); const code=dm.dataset.dec; // v1.8.0: double-click a gateway = gateway panel
      if($('processEditMode').value==='view'){ toast('Pilih mode AS-IS atau TO-BE untuk mengedit gateway.'); return; }
      structuralEditor(code,'connector','decision'); return; } }
  if(!g && state.diagram){ const svg=$('canvas').querySelector('svg'); const r=svg&&svg.getBoundingClientRect(); const y=r?(e.clientY-r.top)/state.zoom:-1, ln=laneAt(svgX(e));
    if(ln && ln.kind==='act' && y>=state.diagram.top && y<=state.diagram.bodyTop){ e.preventDefault(); openRoles();
      const unit=laneUnitOf(state.diagram.lanes.indexOf(ln)), tx=e.target.closest('text'), onUnit=laneMode(state.current)!=='unit' && unit && tx && unit.startsWith(String(tx.textContent||'').replace(/…$/,'').trim()) && String(tx.textContent||'').trim().length>1;
      if(onUnit){ state.renamingUnit=ln.key; state.renaming=null; renderRoles(); setTimeout(()=>{ const i=$('laneNewUnit'); if(i){ i.focus(); i.select(); i.closest('.rcard')?.scrollIntoView({block:'center'}); } },0); return; }
      state.renaming=ln.key; renderRoles(); setTimeout(()=>{ const i=$('laneNewName'); if(i){ i.focus(); i.select(); i.closest('.rcard')?.scrollIntoView({block:'center'}); } },0); return; } } if(g && !e.target.closest('.dmd') && state.byCode.has(g.dataset.code)){ e.preventDefault(); if($('processEditMode').value==='view') showDetail(g.dataset.code, true); else directEditor(g.dataset.code); } };

function symbolMap(a){try{const v=JSON.parse(a.symbolDocs||'{}');return v&&typeof v==='object'&&!Array.isArray(v)?v:{};}catch(e){return {};}}
function symbolDocLabel(a,t,p){return symbolMap(a)[t+'|'+p]||'';}
function symbolSetDoc(a,t,p,label){const m=symbolMap(a),k=t+'|'+p;if(label)m[k]=label;else delete m[k];a.symbolDocs=Object.keys(m).length?JSON.stringify(m):'';}
function symbolEditor(el){
 const g=state.l4s.get(state.current),mode=$('processEditMode').value;if(!g)return;
 const t=el.dataset.symbol,term=t==='start'||t==='end',field=t==='start'?'prev':'next',lf=t==='start'?'startLabel':'endLabel',pf=t==='docin'?'inParty':'outParty',party=el.dataset.symbolParty||'',code=el.dataset.symbolCode||'';
 const rows=mode==='view'?(state.diagram?.acts||g.acts):structuralRows(g,mode),linked=term?rows.filter(a=>(!code||a.l5no===code)&&parseFlow(a[field]).some(f=>f.kind===t)):rows.filter(a=>a.l5no===code),a=linked[0];
 const label=term?(a?.[lf]||(t==='start'?'Mulai':'Selesai')):(a&&symbolDocLabel(a,t,party)|| (t==='docin'?'Input dari ':'Output ke ')+laneShort(roleInfo(party).name));
 const editable=mode!=='view'&&!!a,detail=$('detail');state.sideMode='symbol';delete detail.dataset.code;['btnElab','btnX','btnR'].forEach(x=>$(x).classList.remove('on'));
 const options=(selected)=>rows.map(r=>'<option value="'+esc(r.l5no)+'"'+(r.l5no===selected?' selected':'')+'>'+esc(r.l5no+' — '+r.l5)+'</option>').join('');
 detail.innerHTML='<button class="tb x" data-a="close">✕</button><div class="ws-panel-title">Detail Simbol</div><h3>'+esc(term?(t==='start'?'Mulai':'Selesai'):(t==='docin'?'Dokumen Input':'Dokumen Output'))+'</h3><label>Label simbol</label><input id="symLabel" maxlength="120" value="'+esc(label)+'"'+(!editable?' readonly':'')+'><p class="muted">'+esc(term?'Terminator mengikuti struktur proses.':'Pihak terkait: '+roleInfo(party).name)+'</p><label>Aktivitas terhubung</label>'+linked.map(r=>'<p>'+esc(r.l5no+' — '+r.l5)+'</p>').join('')+(!a?'<p class="muted">Mulai otomatis dari engine; belum ada terminator eksplisit pada data.</p>':'')+(editable&&!term?'<label>Aktivitas tujuan</label><select id="symTarget">'+options(a.l5no)+'</select>':'')+(editable&&!term?'<label>Jenis dokumen</label><select id="symType"><option value="docin"'+(t==='docin'?' selected':'')+'>Input</option><option value="docout"'+(t==='docout'?' selected':'')+'>Output</option></select>':'')+(editable?'<div class="cbtn"><button class="sbtn ws-primary" id="symSave">Simpan '+(mode==='asis'?'koreksi':'usulan')+'</button>'+(!term?'<button class="sbtn no" id="symDelete">Hapus dokumen</button>':'')+'</div>':'')+'<div id="symError" role="alert" style="color:#a21d32"></div>';
 detail.classList.remove('hidden');
 if(!editable)return;
 const commit=del=>{try{
  if(state.current!==g.l4no||$('processEditMode').value!==mode)throw Error('Buka kembali simbol pada mode aktif.');
  if(mode==='asis'&&(state.elab[g.l4no]||[]).some(i=>i.type==='graph'&&i.status!=='rejected'))throw Error('Selesaikan usulan struktur TO-BE sebelum koreksi AS-IS.');
  const result=JSON.parse(JSON.stringify(rows)),src=result.find(r=>r.l5no===a.l5no),target=result.find(r=>r.l5no===($('symTarget')?.value||a.l5no)),name=$('symLabel').value.trim();
  if(!del&&!name)throw Error('Label wajib diisi.');
  if(term){
   if(target!==src){const conflict=structuralConflict(g,src.l5no,'connector',mode)||structuralConflict(g,target.l5no,'connector',mode);if(conflict)throw Error(conflict);
    const fs=parseFlow(src[field]),ts=parseFlow(target[field]);if(fs.length!==1||fs[0].kind!==t||fs[0].label||ts.some(f=>f.kind===t)||hasFormalDecision(src)||hasFormalDecision(target))throw Error('Perpindahan terminator bercabang tidak didukung. Gunakan editing alur aktivitas.');
    // A terminal cannot be detached leaving an empty mandatory field; move activity instead.
    throw Error('Gunakan Pindahkan urutan aktivitas untuk memindahkan titik '+(t==='start'?'Mulai':'Selesai')+' tanpa memutus alur.');
   }
   result.filter(r=>(!code||r.l5no===code)&&parseFlow(r[field]).some(f=>f.kind===t)).forEach(r=>r[lf]=name===(t==='start'?'Mulai':'Selesai')?'':name);
  }else{
   const rawParty=splitParty(src[pf]).find(p=>p===party||('unit::'+unitFor(g.l4no,p)===party)||('pos::'+posFor(g.l4no,p,'')===party));if(!rawParty)throw Error('Pihak dokumen tidak ditemukan.');
   const nt=$('symType').value,np=nt==='docin'?'inParty':'outParty';
   if(!del&&(target!==src||np!==pf)&&splitParty(target[np]).includes(rawParty))throw Error('Dokumen pihak ini sudah terhubung ke aktivitas tujuan.');
   src[pf]=splitParty(src[pf]).filter(p=>p!==rawParty).join('; ');symbolSetDoc(src,t,party,'');
   if(!del){target[np]=[...splitParty(target[np]),rawParty].join('; ');symbolSetDoc(target,nt,party,name);}
  }
  const outside=state.rows.filter(r=>r.l4no!==g.l4no),before=new Set(validateWorkbookRows([...outside,...rows]).map(w=>w.code+'|'+w.msg)),issues=validateWorkbookRows([...outside,...result]).filter(w=>!before.has(w.code+'|'+w.msg));if(issues.length)throw Error(issues[0].msg);
  if(!confirm(del?'Hapus hubungan dokumen ini?':'Simpan perubahan simbol '+mode.toUpperCase()+'?'))return;
  const prior=snapshot();storageFailed=false;
  if(mode==='asis'){result.forEach(r=>r._asisEdited=true);ingest([...outside,...result],state.fileName);select(g.l4no);}else saveUnifiedDraft(g,result,consumedIds(rows),'Edit simbol '+t+' '+a.l5no);
  undoStack.push({before:prior,after:snapshot()});detail.classList.add('hidden');state.sideMode=null;toast(storageFailed?'Perubahan hanya di memori. Export Master Workbook.':'Perubahan simbol tersimpan.');
 }catch(e){$('symError').textContent=e.message;}};
 $('symSave').onclick=()=>commit(false);if($('symDelete'))$('symDelete').onclick=()=>commit(true);
 if(term&&a){const b=document.createElement('button');b.className='sbtn';b.textContent='Edit alur aktivitas';b.onclick=()=>directEditor(a.l5no);detail.append(b);if(linked.length===1){const m=document.createElement('button');m.className='sbtn';m.textContent='Pindahkan aktivitas terhubung';m.onclick=()=>structuralEditor(a.l5no,'move');detail.append(m);}}
}

$('canvas').onclick = e => { const ev=e.target.closest('[data-ev]'); if(ev) return eventEditor(ev.dataset.ev, ev.dataset.evAct, ev.dataset.evSrc); const sym=e.target.closest('[data-symbol]');if(sym){ if(sym.dataset.symbol==='start'||sym.dataset.symbol==='end') return eventEditor(sym.dataset.symbol, sym.dataset.symbolCode||''); return symbolEditor(sym); } if(e.target.closest('.dmd') || e.target.closest('#lanePop')) return; const xx = e.target.closest('[data-x]'); if(xx) return openExtra(xx.dataset.x); const el = e.target.closest('[data-elab]'); if(el){const g=state.l4s.get(state.current),i=g&&(state.elab[g.l4no]||[]).find(x=>x.id===el.dataset.elab);if($('processEditMode').value==='tobe'&&i?.type==='newact')return directEditor(i.id);if(i&&ANNOT_TYPES.includes(i.type)){if($('processEditMode').value==='tobe')return directEditor(i.after);if($('processEditMode').value==='asis')return directEditor(i.after);return showDetail(i.after);}return openElab(el.dataset.elab);} const g = e.target.closest('[data-code]'); if(g){ if($('processEditMode').value!=='view' ) return directEditor(g.dataset.code); if(state.sideMode==='extra' && !$('detail').classList.contains('hidden')) return openExtra(g.dataset.code); showDetail(g.dataset.code); } };
$('detail').addEventListener('click', onSideClick);
$('detail').addEventListener('change', e => { if(onRoleChange(e)) return; onSideChange(e); });
$('detail').addEventListener('input', e => { if(e.target.id==='freeTxt') state.freeText = e.target.value; else if(e.target.id==='newRole') state.newRoleText = e.target.value; else onExtraInput(e); });
$('btnR').onclick = () => { if(state.sideMode==='roles' && !$('detail').classList.contains('hidden')){ $('detail').classList.add('hidden'); state.sideMode=null; $('btnR').classList.remove('on'); } else openRoles(); };
$('btnX').onclick = () => { $('btnR').classList.remove('on'); if(state.sideMode==='extra' && !$('detail').classList.contains('hidden')){ $('detail').classList.add('hidden'); state.sideMode=null; $('btnX').classList.remove('on'); } else openExtra(); };
$('legacyImport').onclick=()=>$('elabIn').click();
$('legacyExport').onclick=()=>exportElabSheet();
$('legacyAi').onclick=()=>{if(!state.current)return toast('Buka proses terlebih dahulu.');openElab();state.legacyAiPanel=true;renderLegacyAi();};
$('btnElab').onclick = () => { $('btnX').classList.remove('on'); $('btnR').classList.remove('on'); if(state.sideMode==='elab' && !state.legacyAiPanel && !$('detail').classList.contains('hidden')){ $('detail').classList.add('hidden'); state.sideMode=null; $('btnElab').classList.remove('on'); } else { openElab(); } };
$('elabIn').onchange = e => { importElabFile(e.target.files[0]); e.target.value=''; };
$('laneMode').onchange = e => { if(!state.current) return; if(!['position','unit'].includes(e.target.value)) return; state.laneModes[state.current]=e.target.value; saveLaneSettings(); redraw(); if(state.sideMode==='roles') renderRoles(); else if(state.sideMode==='direct'&&$('detail').dataset.code) directEditor($('detail').dataset.code); }; $('posOn').onchange = e => { state.posOn = e.target.checked; savePos(); redraw(); };
$('colorOn').onchange = e => { state.colorView = e.target.checked; store.set('flowbpa:color', state.colorView?'1':'0'); redraw(); };
$('zIn').onclick = () => { state.zoom = Math.min(3, state.zoom*1.2); applyZoom(); };
$('zOut').onclick = () => { state.zoom = Math.max(0.15, state.zoom/1.2); applyZoom(); };
$('zFit').onclick = fitZoom;
$('z100').onclick = () => { state.zoom = 1; applyZoom(); };
$('canvas').addEventListener('wheel', e => { if(!(e.ctrlKey||e.metaKey)) return; e.preventDefault(); state.zoom = Math.max(0.15, Math.min(3, state.zoom*(e.deltaY<0?1.1:1/1.1))); applyZoom(); }, {passive:false});
$('statusLabel').oninput = () => { store.set('flowbpa:status', $('statusLabel').value); clearTimeout(window._st); window._st = setTimeout(()=>{ if(state.current) redraw(); }, 300); };
$('btnExport').onclick = openExport;
function openAiSettings(){
  $('aiModel').innerHTML = AI_MODELS.map(([v,l])=>'<option value="'+v+'"'+(v===apiModel()?' selected':'')+'>'+l+'</option>').join('');
  $('aiKey').value = apiKey(); $('aiTest').textContent = ''; $('aiDlg').classList.remove('hidden');
}
function aiApply(){ _sample = undefined; state.sampleAvail = undefined; getSample().then(()=>{ if(state.sideMode==='elab') renderElab(); }); }
$('btnAi').onclick = openAiSettings;
$('aiCancel').onclick = () => $('aiDlg').classList.add('hidden');
$('aiSave').onclick = () => { store.set('flowbpa:apikey', $('aiKey').value.trim()); store.set('flowbpa:apimodel', $('aiModel').value); aiApply(); $('aiDlg').classList.add('hidden'); toast($('aiKey').value.trim() ? 'API key tersimpan di browser ini.' : 'API key dikosongkan.'); };
$('aiClear').onclick = () => { try{ localStorage.removeItem('flowbpa:apikey'); }catch(e){} $('aiKey').value = ''; aiApply(); $('aiTest').textContent = 'API key dihapus dari browser ini.'; };
$('aiCheck').onclick = async () => {
  const keep = [apiKey(), apiModel()]; store.set('flowbpa:apikey', $('aiKey').value.trim()); store.set('flowbpa:apimodel', $('aiModel').value);
  $('aiTest').textContent = 'Menghubungi API…';
  try{ const r = await anthropicSample()('Balas dengan satu kata: siap'); $('aiTest').textContent = '✓ Terhubung. Jawaban: '+r.text.slice(0,40); }
  catch(e){ $('aiTest').textContent = '✕ Gagal: '+(e.code==='not_granted'?'API key tidak valid.':e.message||e.code); }
  finally{ store.set('flowbpa:apikey', keep[0]); store.set('flowbpa:apimodel', keep[1]); }
};
$('btnNew').onclick = async () => { if(importingWorkbook||_newResolve)return; const a = await askNewProject(''); if(a==='new'&&resetProject()){ toast('Siap untuk proyek baru. Upload workbook BPA.'); } };
$('newCancel').onclick = () => closeNew('cancel');
$('newKeep').onclick = () => closeNew('keep');
$('newGo').onclick = () => closeNew('new');
$('newSave').onclick = () => exportElabSheet();
$('exCancel').onclick = () => $('exportDlg').classList.add('hidden');
$('exGo').onclick = runExport;
$('exportDlg').addEventListener('change',e=>{if(['fmt','excelPurpose','wordPurpose'].includes(e.target.name))updateExportOptions();});
window.addEventListener('hashchange', () => { const h = hashToL4(); if(h && h!==state.current) select(h); });

/* v1.3.5 Tahap 2 — manual analysis, independent of the process engine. */
const ANALYSIS_KEY='flowbpa:analysis';
const ANALYSIS_FOCUS=['Efisiensi','Efektivitas','Governance & Control','Organisasi','Digitalisasi'];
const ANALYSIS_STATUS=['Terbuka','Dalam review','Disepakati','Ditutup'];
let analysisRecords=[], analysisEditing=null;
function normalizeAnalysis(r){
  if(!r||!r.id||!r.l4||!String(r.finding||'').trim())return null;
  return {id:String(r.id),l4:String(r.l4),activity:String(r.activity||''),activityName:String(r.activityName||''),focus:ANALYSIS_FOCUS.includes(r.focus)?r.focus:ANALYSIS_FOCUS[0],finding:String(r.finding),evidence:String(r.evidence||''),risk:String(r.risk||''),recommendation:String(r.recommendation||''),status:ANALYSIS_STATUS.includes(r.status)?r.status:ANALYSIS_STATUS[0],note:String(r.note||'')};
}
function saveAnalysis(next){analysisRecords=next;return store.set(ANALYSIS_KEY,JSON.stringify(next));}
function mergeAnalysis(records){if(!records)return;const map=new Map(analysisRecords.map(r=>[r.id,r]));for(const raw of records){const r=normalizeAnalysis(raw);if(r)map.set(r.id,r);}saveAnalysis([...map.values()]);}
const ANALYSIS_HEAD=['ID Temuan','No L4','No L5','Fokus Analisis','Temuan','Bukti / Konteks','Risiko / Dampak','Rekomendasi','Status','Catatan'];
function appendAnalysisSheet(wb){const ws=XLSX.utils.aoa_to_sheet([ANALYSIS_HEAD,...analysisRecords.map(r=>[r.id,r.l4,r.activity,r.focus,r.finding,r.evidence,r.risk,r.recommendation,r.status,r.note])]);ws['!cols']=[26,12,15,25,55,55,55,55,18,45].map(w=>({wch:w}));ws['!autofilter']={ref:ws['!ref']};XLSX.utils.book_append_sheet(wb,ws,'Analisis Proses');}
function readAnalysisSheet(wb){const ws=wb.Sheets['Analisis Proses'];if(!ws)return null;const a=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:''});const hi=a.findIndex(r=>r.map(norm).includes('temuan')&&r.map(norm).includes('no l4'));if(hi<0)throw Error('Format sheet Analisis Proses tidak sesuai.');const h=a[hi].map(norm);return a.slice(hi+1).map((r,i)=>{const get=(...names)=>{const k=h.findIndex(x=>names.includes(x));return k<0?'':r[k]||'';};const finding=get('temuan'),l4=get('no l4');if(!finding&&!l4)return null;if(!finding||!l4)throw Error('Analisis Proses baris '+(hi+i+2)+': No L4 dan Temuan wajib diisi.');return normalizeAnalysis({id:get('id temuan','id')||'import-'+l4+'-'+(hi+i+2),l4,activity:get('no l5'),activityName:get('aktivitas'),focus:get('fokus analisis','fokus'),finding,evidence:get('bukti / konteks'),risk:get('risiko / dampak'),recommendation:get('rekomendasi'),status:get('status'),note:get('catatan')});}).filter(Boolean);}
function analysisOptions(values,selected){return values.map(v=>'<option value="'+esc(v)+'"'+(v===selected?' selected':'')+'>'+esc(v)+'</option>').join('');}
function renderAnalysis(){const g=state.l4s.get(state.current);if(!g)return;const list=analysisRecords.filter(r=>r.l4===g.l4no);$('analysisTitle').textContent=g.l4no+' — '+g.l4;$('analysisSummary').textContent=g.acts.length+' aktivitas AS-IS · '+new Set(g.acts.map(a=>a.role).filter(Boolean)).size+' role · '+list.length+' temuan';
 $('analysisList').innerHTML=list.length?list.map(r=>'<article class="analysis-card"><div class="analysis-meta">'+esc(r.focus)+' · '+esc(r.status)+'</div><strong>'+esc(r.finding)+'</strong><p class="muted">'+esc(r.activity?(r.activity+' — '+r.activityName):'Seluruh subproses')+'</p>'+[['Bukti / konteks',r.evidence],['Risiko / dampak',r.risk],['Rekomendasi',r.recommendation],['Catatan',r.note]].filter(x=>x[1]).map(x=>'<p><b>'+x[0]+':</b> '+esc(x[1])+'</p>').join('')+'<div class="cbtn"><button class="sbtn" data-analysis-edit="'+esc(r.id)+'">Ubah</button><button class="sbtn no" data-analysis-delete="'+esc(r.id)+'">Hapus</button></div></article>').join(''):'<p class="muted">Belum ada temuan. Tambahkan hasil observasi atau diskusi dengan klien.</p>';
}
function editAnalysis(r){const g=state.l4s.get(state.current);if(!g)return;analysisEditing=r?r.id:null;const f=$('analysisForm');f.hidden=false;const acts=g.acts.slice();$('analysisActivity').innerHTML='<option value="">Seluruh subproses</option>'+acts.map(a=>'<option value="'+esc(a.l5no)+'">'+esc(a.l5no+' — '+a.l5)+'</option>').join('');if(r?.activity&&!acts.some(a=>a.l5no===r.activity))$('analysisActivity').insertAdjacentHTML('beforeend','<option value="'+esc(r.activity)+'">'+esc(r.activity+' — '+r.activityName+' (referensi lama)')+'</option>');$('analysisActivity').value=r?.activity||'';$('analysisFocus').innerHTML=analysisOptions(ANALYSIS_FOCUS,r?.focus||ANALYSIS_FOCUS[0]);$('analysisStatus').innerHTML=analysisOptions(ANALYSIS_STATUS,r?.status||ANALYSIS_STATUS[0]);for(const [id,key] of [['analysisFinding','finding'],['analysisEvidence','evidence'],['analysisRisk','risk'],['analysisRecommendation','recommendation']])$(id).value=r?.[key]||'';$('analysisFinding').focus();}
function closeAnalysis(){if(!$('analysisForm').hidden&&!confirm('Tutup tanpa menyimpan perubahan temuan?'))return;$('analysisForm').hidden=true;$('analysisDlg').classList.add('hidden');analysisEditing=null;}
$('btnAnalysis').onclick=()=>{if(!state.current)return;$('analysisForm').hidden=true;analysisEditing=null;renderAnalysis();$('analysisDlg').classList.remove('hidden');};
$('analysisClose').onclick=closeAnalysis;
$('analysisAdd').onclick=()=>{if(!$('analysisForm').hidden&&!confirm('Buang perubahan formulir yang belum disimpan?'))return;editAnalysis(null);};
$('analysisCancel').onclick=()=>{$('analysisForm').hidden=true;analysisEditing=null;};
$('analysisList').onclick=e=>{const edit=e.target.closest('[data-analysis-edit]'),del=e.target.closest('[data-analysis-delete]');if(edit){if(!$('analysisForm').hidden&&!confirm('Buang perubahan formulir yang belum disimpan?'))return;editAnalysis(analysisRecords.find(r=>r.id===edit.dataset.analysisEdit));}if(del&&confirm('Hapus temuan ini?')){const id=del.dataset.analysisDelete;const ok=saveAnalysis(analysisRecords.filter(r=>r.id!==id));if(analysisEditing===id){$('analysisForm').hidden=true;analysisEditing=null;}renderAnalysis();if(ok)toast('Temuan dihapus.');}};
$('analysisForm').onsubmit=e=>{e.preventDefault();const g=state.l4s.get(state.current);if(!g)return;const finding=$('analysisFinding').value.trim();if(!finding){$('analysisFinding').focus();return;}const activity=$('analysisActivity').value,old=analysisRecords.find(r=>r.id===analysisEditing);const r=normalizeAnalysis({id:analysisEditing||('an-'+Date.now().toString(36)+'-'+Math.random().toString(36).slice(2,10)),l4:g.l4no,activity,activityName:g.acts.find(a=>a.l5no===activity)?.l5||(activity===old?.activity?old.activityName:''),focus:$('analysisFocus').value,finding,evidence:$('analysisEvidence').value.trim(),risk:$('analysisRisk').value.trim(),recommendation:$('analysisRecommendation').value.trim(),status:$('analysisStatus').value,note:old?.note||''});const ok=saveAnalysis([...analysisRecords.filter(x=>x.id!==r.id),r]);$('analysisForm').hidden=true;analysisEditing=null;renderAnalysis();if(ok)toast('Temuan tersimpan.');};

/* ================= Restore last session ================= */
(function init(){
  syncWorkspace();
  try{const a=JSON.parse(store.get(REGISTER_KEY)||'[]');if(Array.isArray(a))changeRegister=a.map(normalizeRegister).filter(Boolean);}catch(e){}
  try{const a=JSON.parse(store.get(ANALYSIS_KEY)||'[]');if(Array.isArray(a))analysisRecords=a.map(normalizeAnalysis).filter(Boolean);}catch(e){}
  try{JSON.parse(store.get('flowbpa:usedCodes')||'[]').forEach(c=>usedCodes.add(c));}catch(e){}
  try{ state.unitmap=JSON.parse(store.get('flowbpa:unitmap')||'{}')||{}; state.laneModes=JSON.parse(store.get('flowbpa:lanemodes')||'{}')||{}; }catch(e){}
  const st = store.get('flowbpa:status'); if(st!=null) $('statusLabel').value = st;
  state.showProposed = store.get('flowbpa:showprop') !== '0';
  try{ const e = JSON.parse(store.get('flowbpa:elab')||'{}'); if(e && typeof e==='object') state.elab = e; }catch(err){}
  try{ state.posmap = JSON.parse(store.get('flowbpa:posmap')||'{}') || {}; state.decmap = JSON.parse(store.get('flowbpa:decmap')||'{}') || {}; state.posOn = store.get('flowbpa:poson') !== '0'; }catch(err){}
  try{ state.extra = JSON.parse(store.get('flowbpa:extra')||'{}') || {}; }catch(err){}
  try{ const r = JSON.parse(store.get('flowbpa:roles')||'null'); if(r && r.alias) state.roleEdit = {alias:r.alias||{}, added:r.added||{}, hidden:r.hidden||{}, order:r.order||{}}; }catch(err){}
  updatePosToggle();
  state.colorView = store.get('flowbpa:color') === '1'; $('colorOn').checked = state.colorView;
  const raw = store.get('flowbpa:data');
  if(raw){ try{ const d = JSON.parse(raw); if(d && d.rows && d.rows.length) ingest(d.rows, d.fileName||'workbook.xlsx'); }catch(e){} }
})();

/* v1.4.0: BPMN pool name (Pengaturan). */
(()=>{const el=$('poolLabel');if(!el)return;const v=store.get('flowbpa:pool');if(v!=null)el.value=v;
el.oninput=()=>{store.set('flowbpa:pool',el.value);clearTimeout(window._pl);window._pl=setTimeout(()=>{if(state.current)redraw();},300);};
if(v!=null&&state.current)redraw();})();

/* v1.5.0: optional XOR merge gateway at process start (Pengaturan). */
(()=>{const el=$('mergeGw');if(!el)return;const v=store.get('flowbpa:mergegw');if(v!=null)el.checked=v==='1';
el.onchange=()=>{store.set('flowbpa:mergegw',el.checked?'1':'0');if(state.current)redraw();};
if(v!=null&&state.current)redraw();})();

/* v1.6.4: legend names for Dokumen / Data (Pengaturan). */
(()=>{for(const [id,key] of [['legendDoc','flowbpa:legendDoc'],['legendData','flowbpa:legendData']]){const el=$(id);if(!el)continue;const v=store.get(key);if(v!=null)el.value=v;
el.oninput=()=>{store.set(key,el.value);clearTimeout(window._lg);window._lg=setTimeout(()=>{if(state.current)redraw();},300);};}})();

/* ================= v1.9.1: edit Mulai / Selesai / pemicu & serah terima lintas L4 ================= */
function eventChange(rows, s){
  rows=JSON.parse(JSON.stringify(rows));
  const by=c=>rows.find(r=>r.l5no===c), a=by(s.code); if(!a) throw Error('Aktivitas '+s.code+' tidak ditemukan.');
  const l4=a.l4no, local=new Set(rows.map(r=>r.l5no)), patches={out:[],warn:[]};
  const addPrev=(t,f)=>{ const p=parseFlow(t.prev); if(!p.some(x=>x.kind===f.kind&&(x.kind!=='code'||x.code===f.code))) (f.kind==='start'?p.unshift(f):p.push(f)); t.prev=flowString(p); };
  const linkFrom=(fromCode)=>{ if(!fromCode) return; const p=by(fromCode); if(!p) throw Error('Aktivitas '+fromCode+' tidak ditemukan.'); if(p===a) throw Error('Pilih aktivitas lain sebagai pendahulu.');
    const nx=parseFlow(p.next); if(nx.some(f=>f.label)) throw Error(fromCode+' punya gateway. Sambungkan lewat panel gateway (klik dua kali belah ketupatnya).');
    if(!nx.some(f=>f.kind==='code'&&f.code===a.l5no)) nx.push({kind:'code',code:a.l5no}); p.next=flowString(nx); addPrev(a,{kind:'code',code:fromCode}); };
  const linkTo=(isRem, nextTo)=>{ // replace removed end/handoff flows of a by nextTo, in place (keeps branch labels and order)
    const fl=parseFlow(a.next), removed=fl.filter(isRem), keep=fl.filter(f=>!isRem(f));
    if(nextTo){ const t=by(nextTo); if(!t) throw Error('Aktivitas '+nextTo+' tidak ditemukan.'); if(t===a) throw Error('Pilih aktivitas lain sebagai kelanjutan.'); addPrev(t,{kind:'code',code:a.l5no}); }
    else for(const f of removed) if(f.label && !keep.some(k=>k.label===f.label)) throw Error('Cabang "'+f.label+'" jadi tidak punya tujuan. Pilih "lanjut ke" untuk cabang itu.');
    const out=[], done=new Set();
    for(const f of fl){ if(!isRem(f)){ out.push(f); continue; } const lab=f.label||''; if(nextTo && !done.has(lab)){ done.add(lab); if(!fl.some(x=>!isRem(x)&&(x.label||'')===lab&&x.kind==='code'&&x.code===nextTo)) out.push({kind:'code',code:nextTo,label:lab,note:''}); } }
    if(!out.length) throw Error('Aktivitas '+a.l5no+' jadi tidak punya kelanjutan. Pilih "lanjut ke".');
    a.next=flowString(out); };
  const addEndTo=(code)=>{ if(!code) return; const t=by(code); if(!t) throw Error('Aktivitas '+code+' tidak ditemukan.'); const nx=parseFlow(t.next); if(nx.some(f=>f.label)) throw Error(code+' punya gateway. Atur Selesai lewat panel gateway.'); if(!nx.some(f=>f.kind==='end')) nx.push({kind:'end'}); t.next=flowString(nx); if(a.endLabel&&!t.endLabel) t.endLabel=a.endLabel; };
  if(s.kind==='start'){
    if(s.label!==undefined) a.startLabel=s.label==='Mulai'?'':s.label;
    if(s.remove){ const pv=parseFlow(a.prev); if(!pv.some(f=>f.kind==='start')) throw Error('Aktivitas ini tidak punya titik Mulai.');
      a.prev=flowString(pv.filter(f=>f.kind!=='start'));
      if(s.addTo){ const t=by(s.addTo); if(!t) throw Error('Aktivitas tujuan tidak ditemukan.'); addPrev(t,{kind:'start'}); if(a.startLabel&&!t.startLabel) t.startLabel=a.startLabel; }
      a.startLabel=''; linkFrom(s.connectFrom); }
  } else if(s.kind==='trig'){
    const pv=parseFlow(a.prev), f0=pv.find(f=>f.kind==='code'&&f.code===s.src); if(!f0) throw Error('Pemicu dari '+s.src+' tidak ditemukan di '+a.l5no+'.');
    a.prev=flowString(pv.filter(f=>!(f.kind==='code'&&f.code===s.src)));
    a.dataFrom=String(a.dataFrom||'').split(/[;,\s]+/).filter(c=>c&&c!==s.src).join('; ');
    if(s.addTo){ const t=by(s.addTo); if(!t) throw Error('Aktivitas tujuan tidak ditemukan.'); addPrev(t,{kind:'code',code:s.src,note:f0.note||''}); }
    patches.out.push({code:s.src, field:'next', from:a.l5no, to:s.addTo||''});
    linkFrom(s.connectFrom);
  } else if(s.kind==='end'){
    if(s.label!==undefined) a.endLabel=s.label==='Selesai'?'':s.label;
    if(s.remove){ const fl=parseFlow(a.next), rem=fl.filter(f=>f.kind==='end'); if(!rem.length) throw Error('Aktivitas ini tidak punya titik Selesai.');
      const carry=a.endLabel; linkTo(f=>f.kind==='end', s.nextTo); a.endLabel=''; if(s.addTo){ a.endLabel=carry; addEndTo(s.addTo); a.endLabel=''; } }
  } else if(s.kind==='hand'){
    const fl=parseFlow(a.next), rem=fl.filter(f=>f.kind==='code'&&f.code===s.src); if(!rem.length) throw Error('Serah terima ke '+s.src+' tidak ditemukan.');
    linkTo(f=>f.kind==='code'&&f.code===s.src, s.nextTo); patches.out.push({code:s.src, field:'prev', from:a.l5no, to:''});
  } else throw Error('Jenis simbol tidak dikenal.');
  const outside=state.rows.filter(r=>r.l4no!==l4);
  const before=validateWorkbookRows([...outside,...structuralRows(state.l4s.get(l4),s.mode)]), old=new Set(before.map(w=>w.code+'|'+w.msg));
  const issues=validateWorkbookRows([...outside,...rows]).filter(w=>!old.has(w.code+'|'+w.msg));
  const crossMsg=w=>w.l4!==l4||(String(w.msg).match(/(?:mengarah ke|menyebut|mencantumkan) (\d+\.\d+(?:\.\d+)?)/g)||[]).some(x=>!local.has(x.split(' ').pop()));
  const hard=issues.filter(w=>!crossMsg(w));
  if(hard.length){ const h=hard[0]; if(/tidak dapat dicapai|Proses Sebelum kosong/.test(h.msg)) throw Error('Aktivitas '+h.code+' jadi tidak tersambung dari Mulai. Isi "sambungkan dari" atau "pindahkan ke" supaya alurnya tidak putus.'); if(/SELESAI/.test(h.msg)) throw Error('Proses jadi tidak punya jalur ke Selesai. Isi "tambahkan Selesai setelah" atau "lanjut ke".'); throw Error('Perubahan ditolak: '+h.msg); }
  rows.patches=patches; return rows;
}
function eventEditor(kind, actCode, src){
  const g=state.l4s.get(state.current), mode=$('processEditMode').value; if(!g) return;
  const detail=$('detail'); state.sideMode='symbol'; delete detail.dataset.code; ['btnElab','btnX','btnR'].forEach(x=>$(x).classList.remove('on'));
  const view=mode==='view'; let rows=[]; try{ rows=view?(state.diagram?.acts||g.acts):structuralRows(g,mode); }catch(e){ toast(e.message); return; }
  const endActs=kind==='end'?rows.filter(r=>parseFlow(r.next).some(f=>f.kind==='end')):[];
  if(kind==='end'&&!actCode&&endActs.length) actCode=endActs[0].l5no;
  const acts=rows.filter(r=>!isGateRow(r)), a=rows.find(r=>r.l5no===actCode);
  const opt=(sel,ph)=>'<option value="">'+esc(ph)+'</option>'+acts.filter(r=>r.l5no!==actCode).map(r=>'<option value="'+esc(r.l5no)+'"'+(r.l5no===sel?' selected':'')+'>'+esc(r.l5no+' — '+r.l5)+'</option>').join('');
  const srcs=String(src||'').split(',').filter(Boolean), ext=c=>{ const r=state.byCode.get(c); return c+(r?' — '+r.l5:''); };
  const T={start:'Titik Mulai',end:'Titik Selesai',trig:'Dipicu dari L4 lain',hand:'Lanjut ke L4 lain'}[kind];
  let h='<button class="tb x" data-a="close">✕</button><div class="ws-panel-title">Detail Simbol</div><h3>'+esc(T)+'</h3>';
  if(!a){ detail.innerHTML=h+'<p class="muted">Simbol ini tidak terhubung ke aktivitas yang bisa diedit.</p>'; detail.classList.remove('hidden'); return; }
  if(kind==='end'&&endActs.length>1&&oneEnd()) h+='<label for="evWhich">Selesai setelah</label><select id="evWhich">'+endActs.map(r=>'<option value="'+esc(r.l5no)+'"'+(r.l5no===actCode?' selected':'')+'>'+esc((isGateRow(r)?'◇ '+gateQ(r):r.l5no+' — '+r.l5))+'</option>').join('')+'</select><p class="ws-field-hint">Titik Selesai ini dipakai bersama oleh '+endActs.length+' jalur. Pilih jalur yang mau diubah; label berlaku untuk semuanya.</p>';
  else h+='<p class="muted">'+esc({start:'Mulai sebelum ',end:'Selesai setelah ',trig:'Memicu ',hand:'Dikirim dari '}[kind]+a.l5no+' — '+a.l5)+'</p>';
  if(view){ detail.innerHTML=h+'<p class="ws-field-hint">Pilih mode AS-IS atau TO-BE untuk mengubah simbol ini.</p>'; detail.classList.remove('hidden'); return; }
  const lab=kind==='start'?(a.startLabel||'Mulai'):kind==='end'?(a.endLabel||'Selesai'):'';
  if(kind==='start'||kind==='end') h+='<label for="evLabel">Label</label><input id="evLabel" maxlength="120" value="'+esc(lab)+'"><div class="cbtn"><button class="sbtn" id="evSaveLabel">Simpan label</button></div>';
  if(kind==='trig'&&srcs.length>1) h+='<label for="evSrc">Pemicu</label><select id="evSrc">'+srcs.map(c=>'<option value="'+esc(c)+'">'+esc(ext(c))+'</option>').join('')+'</select>';
  if(kind==='hand') h+='<p>Serah terima ke <b>'+esc(ext(srcs[0]))+'</b>.</p>';
  h+='<div class="ws-section"><strong class="ws-section-title">'+esc({start:'Hapus / pindahkan Mulai',end:'Hapus / pindahkan Selesai',trig:'Hapus / pindahkan pemicu',hand:'Hapus serah terima'}[kind])+'</strong>';
  if(kind==='start'||kind==='trig'){
    h+='<label for="evAddTo">Pindahkan ke aktivitas (opsional)</label><select id="evAddTo">'+opt('','— tidak dipindah, hapus saja —')+'</select>'+
       '<label for="evFrom">Sambungkan '+esc(a.l5no)+' dari (opsional)</label><select id="evFrom">'+opt('','— tidak perlu —')+'</select>'+
       '<p class="ws-field-hint">Kalau setelah dihapus '+esc(a.l5no)+' tidak punya aktivitas sebelumnya, pilih dari aktivitas mana alurnya datang.'+(kind==='trig'?' Di AS-IS, Proses Sesudah di L4 asal ikut disesuaikan; di TO-BE L4 asal tetap seperti semula.':'')+'</p>';
  } else {
    h+='<label for="evNext">'+esc(a.l5no)+' lanjut ke (opsional)</label><select id="evNext">'+opt('','— tidak perlu —')+'</select>'+
      (kind==='end'?'<label for="evAddTo">Tambahkan Selesai setelah (opsional)</label><select id="evAddTo">'+opt('','— tidak perlu —')+'</select>':'')+
      '<p class="ws-field-hint">Kalau setelah dihapus '+esc(a.l5no)+' tidak punya kelanjutan, pilih "lanjut ke".'+(kind==='hand'?' Di AS-IS, Proses Sebelum di L4 tujuan ikut disesuaikan.':'')+'</p>';
  }
  h+='<div class="cbtn"><button class="sbtn no" id="evRemove">'+esc({start:'Hapus Mulai',end:'Hapus Selesai',trig:'Hapus pemicu',hand:'Hapus serah terima'}[kind])+'</button></div></div>';
  detail.innerHTML=h; detail.classList.remove('hidden');
  const run=(spec,msg)=>{ const old=$('evError'); if(old) old.remove();
    try{
      if(state.current!==g.l4no||$('processEditMode').value!==mode) throw Error('Buka kembali simbol pada mode aktif.');
      if(mode==='asis'&&draftGraph(g)) throw Error('Selesaikan rancangan TO-BE sebelum koreksi AS-IS.');
      const fresh=structuralRows(g,mode), absorbed=consumedIds(fresh), result=eventChange(fresh,{...spec,code:actCode,mode}), px=result.patches;
      if(spec.labelOnly&&kind==='end'&&oneEnd()) for(const r of result) if(parseFlow(r.next).some(f=>f.kind==='end')) r.endLabel=spec.label==='Selesai'?'':spec.label;
      if(!spec.labelOnly){ const cross=uniq(px.out.map(x=>x.code)); if(!confirm(msg+(cross.length?(mode==='asis'?'\n\nHubungan di '+cross.join(', ')+' ikut disesuaikan.':'\n\nCatatan: '+cross.join(', ')+' di AS-IS masih mencatat hubungan lama.'):''))) return; }
      const prior=snapshot(); storageFailed=false;
      if(mode==='asis'){ result.forEach(r=>r._asisEdited=true); applyCrossPatches(px); ingest([...state.rows.filter(r=>r.l4no!==g.l4no),...result],state.fileName); select(g.l4no); }
      else saveUnifiedDraft(g,result,absorbed,'Ubah simbol '+kind+' '+actCode);
      undoStack.push({before:prior,after:snapshot()}); detail.classList.add('hidden'); state.sideMode=null;
      if(!storageFailed) toast(spec.labelOnly?'Label disimpan.':'Perubahan disimpan.');
    }catch(e){ const p=document.createElement('p'); p.id='evError'; p.setAttribute('role','alert'); p.style.cssText='color:#a21d32;font-weight:600;margin:10px 0'; p.textContent=e.message; $('evRemove').parentElement.insertAdjacentElement('beforebegin',p); toast(e.message,9000); } };
  const v=id=>($(id)&&$(id).value)||'';
  if($('evWhich')) $('evWhich').onchange=()=>eventEditor('end',$('evWhich').value,'');
  if($('evSaveLabel')) $('evSaveLabel').onclick=()=>{ const t=v('evLabel').trim(); if(!t) return toast('Label wajib diisi.'); run({kind,label:t,labelOnly:true},''); };
  $('evRemove').onclick=()=>{
    const spec={kind,remove:true,src:kind==='trig'?(v('evSrc')||srcs[0]):srcs[0],addTo:v('evAddTo'),connectFrom:v('evFrom'),nextTo:v('evNext')};
    const what={start:'titik Mulai sebelum '+actCode,end:'titik Selesai setelah '+actCode,trig:'pemicu dari '+spec.src+' ke '+actCode,hand:'serah terima '+actCode+' ke '+spec.src}[kind];
    run(spec,(spec.addTo?'Pindahkan ':'Hapus ')+what+(spec.addTo?' ke '+spec.addTo:'')+' pada '+(mode==='asis'?'AS-IS':'TO-BE')+'?');
  };
}

/* ================= v1.9.1: ganti nama unit (subjudul lane) ================= */
function laneUnitOf(laneIdx){ const d=state.diagram; if(!d) return ''; const cnt=new Map(); for(const n of d.nodes) if(n.type==='act'&&n.lane===laneIdx&&n.act){ const u=String(n.act.unitX||'').trim()||unitFor(state.current,n.act.role,n.act); if(u) cnt.set(u,(cnt.get(u)||0)+1); } return [...cnt].sort((x,y)=>y[1]-x[1])[0]?.[0]||''; }
function renameUnit(g, laneKey, raw, allL4){
  const mode=$('processEditMode').value, newU=String(raw||'').replace(/\s+/g,' ').trim(); if(!newU){ toast('Nama unit tidak boleh kosong.'); return; }
  const idx=state.diagram.lanes.findIndex(l=>l.key===laneKey), oldU=laneUnitOf(idx);
  if(mode==='view'){ toast('Pilih mode AS-IS atau TO-BE dulu untuk mengganti nama unit.'); return; }
  if(oldU===newU){ state.renamingUnit=null; renderRoles(); return; }
  const prior=snapshot(); storageFailed=false; let n=0, l4n=0, skipped=[];
  const touch=(rows, pick)=>{ let k=0; for(const r of rows) if(!isGateRow(r)&&pick(r)){ r.unitX=newU; if(mode==='asis') r._asisEdited=true; k++; } return k; };
  const isOld=(l4no,r)=>(String(r.unitX||'').trim()||unitFor(l4no,r.role,r))===oldU;
  const targets=allL4?[...state.l4s.values()]:[g];
  if(mode==='tobe'){
    for(const x of targets){ if(x.l4no!==g.l4no && !draftGraph(x)) continue; const d=materializeDraft(x); const k=touch(d.rows, r=>x.l4no===g.l4no&&!allL4?actLaneKey(x.l4no,r)===laneKey:isOld(x.l4no,r)); if(k){ saveUnifiedDraft(x,d.rows,d.ids,'Ganti nama unit '+oldU+' → '+newU); n+=k; l4n++; } }
  } else {
    let changed=false;
    for(const x of targets){ if(x.l4no!==g.l4no && draftGraph(x)){ if(state.rows.some(r=>r.l4no===x.l4no&&isOld(x.l4no,r))) skipped.push(x.l4no); continue; }
      if(x.l4no===g.l4no && draftGraph(x)){ toast('Selesaikan rancangan TO-BE sebelum koreksi AS-IS.'); return; }
      const k=touch(state.rows.filter(r=>r.l4no===x.l4no), r=>x.l4no===g.l4no&&!allL4?actLaneKey(x.l4no,r)===laneKey:isOld(x.l4no,r)); if(k){ n+=k; l4n++; changed=true; } }
    // role-level fallbacks (Pemetaan Unit) that produced the old name
    for(const [l4no,m] of Object.entries(state.unitmap)) if(allL4||l4no===g.l4no) for(const [role,v] of Object.entries(m||{})) if(v===oldU) m[role]=newU;
    if(changed){ ingest(state.rows,state.fileName); select(g.l4no); }
  }
  undoStack.push({before:prior,after:snapshot()}); state.renamingUnit=null; saveLaneSettings(); roleChanged();
  if(!storageFailed) toast('Unit "'+oldU+'" → "'+newU+'": '+n+' aktivitas di '+l4n+' L4 ('+(mode==='asis'?'AS-IS':'TO-BE')+').'+(skipped.length?' Dilewati (ada rancangan TO-BE): '+skipped.join(', ')+'.':''),6000);
}

/* v1.9.2: Pengaturan — one shared Selesai (default) or one per path */
(()=>{const el=$('oneEndCk');if(!el)return;el.checked=oneEnd();el.onchange=()=>{store.set('flowbpa:oneEnd',el.checked?'1':'0');if(state.current)redraw();};})();
