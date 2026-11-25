
/* ================= v1.9.0: Simpan ke cloud (Supabase) =================
   One project = one row in public.projects holding a snapshot of the browser-side project keys.
   Team access is enforced in the database (members table + row level security).
   Only the PUBLISHABLE key belongs here; never put a secret key in this file. */
const CLOUD = { url:'https://mwntajmqujwlygnutwjy.supabase.co', key:'sb_publishable_gp5m8aejRNsyHSN0PYocGg_p10NvmvA' };
const CLOUD_PROJ_KEY = 'flowbpa:cloudProject', CLOUD_AUTO_KEY = 'flowbpa:cloudAuto';
const CLOUD_HISTORY_EVERY = 15*60*1000; // automatic saves keep a history version at most every 15 minutes
const CLOUD_USER_DOMAIN = 'flowbpa.local'; // v1.9.7: username "dea" signs in as dea@flowbpa.local (no email is ever sent)
let sbc = null, cloudUser = null, cloudMember = false, cloudMemberErr = '', cloudProj = null, cloudHash = '', cloudBusy = false, cloudLastHist = 0, cloudState = '', cloudRows = 0;
function cloudLib(){ return window.__sbMock || window.supabase || null; }
function cloudLoginEmail(u){ u = String(u||'').trim().toLowerCase(); return u.includes('@') ? u : u + '@' + CLOUD_USER_DOMAIN; }
function cloudUserName(email){ email = String(email||''); return email.endsWith('@'+CLOUD_USER_DOMAIN) ? email.slice(0, -CLOUD_USER_DOMAIN.length-1) : email; }
const CLOUD_PART_NAMES = {data:'Data AS-IS', elab:'Rancangan TO-BE', extra:'SLA/Output', roles:'Peran', posmap:'Pemetaan jabatan', decmap:'Keputusan', unitmap:'Unit kerja', lanemodes:'Mode swimlane', poson:'Tampilan jabatan', usedCodes:'Nomor kode', analysis:'Analisis', changeRegister:'Register perubahan'};
function cloudPartNames(arr){ arr = arr || []; return arr.length ? arr.map(k => CLOUD_PART_NAMES[k] || (/^bpa_bpmn_semantics/.test(k) ? 'Jenis simbol BPMN' : k)).filter((v,i,a)=>a.indexOf(v)===i).join(', ') : '(tidak ada isi berubah)'; }
// v1.9.7: number of activity rows in a snapshot; used to refuse saving an empty/shrunken browser state over the cloud copy
function cloudRowCount(snap){ try{ const d = JSON.parse((snap||{})['flowbpa:data']||'null'); return d && Array.isArray(d.rows) ? d.rows.length : 0; }catch(e){ return 0; } }
function cloudClient(){
  if(sbc) return sbc; const L = cloudLib(); if(!L || !L.createClient) return null;
  try{ sbc = L.createClient(CLOUD.url, CLOUD.key, {auth:{persistSession:true, autoRefreshToken:true, detectSessionInUrl:true, flowType:'pkce', storageKey:'flowbpa-auth'}}); }catch(e){ sbc = null; }
  return sbc;
}
function cloudKeys(){ const ks = [...PROJECT_KEYS]; try{ for(const k of Object.keys(localStorage)) if(k.startsWith('bpa_bpmn_semantics_v1391_')) ks.push(k); }catch(e){} return ks; }
function cloudSnapshot(){ const o = {}; for(const k of cloudKeys()){ const v = store.get(k); if(v != null) o[k] = v; } return o; }
function cloudHashOf(o){ const s = JSON.stringify(o); let h = 0; for(let i=0;i<s.length;i++) h = (h*31 + s.charCodeAt(i))|0; return h+':'+s.length; }
function cloudAuto(){ return store.get(CLOUD_AUTO_KEY) !== '0'; }
function cloudTime(t){ try{ return new Date(t).toLocaleString('id-ID',{day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit'}); }catch(e){ return String(t||''); } }
function cloudSetState(txt, kind){
  cloudState = txt; const b = $('btnCloud'); if(!b) return;
  b.textContent = '☁ ' + (cloudProj ? cloudProj.name : 'Cloud') + (txt ? ' · '+txt : '');
  b.title = cloudProj ? 'Proyek cloud: '+cloudProj.name+(txt?' — '+txt:'') : 'Simpan proyek ke cloud supaya bisa dibuka dari mana saja';
  b.style.borderColor = kind==='err' ? '#a21d32' : kind==='warn' ? '#c9a227' : '';
  b.style.color = kind==='err' ? '#a21d32' : '';
}
async function cloudRefreshUser(){
  const c = cloudClient(); if(!c){ cloudUser = null; return; }
  try{ const { data } = await c.auth.getSession(); cloudUser = data && data.session ? data.session.user : null; }catch(e){ cloudUser = null; }
  cloudMember = false; cloudMemberErr = '';
  if(cloudUser){ try{ const { data, error } = await c.from('members').select('email').limit(1); cloudMember = !error && !!(data && data.length); cloudMemberErr = error ? (error.message||String(error.code||'')) : (data && data.length ? '' : 'tidak ada baris yang cocok'); }catch(e){ cloudMemberErr = e.message||String(e); } }
}
function cloudWriteLocal(snap){
  for(const k of cloudKeys()) try{ localStorage.removeItem(k); }catch(e){}
  for(const [k,v] of Object.entries(snap||{})) if(k.startsWith('flowbpa:') || k.startsWith('bpa_bpmn_semantics_v1391_')) store.set(k, v);
}
async function cloudOpen(id, versionRow){
  const c = cloudClient(); if(!c) return;
  if(cloudProj && cloudProj.id !== id && cloudDirty() && !confirm('Perubahan di proyek "'+cloudProj.name+'" belum tersimpan ke cloud. Tetap buka proyek lain?')) return;
  const { data, error } = await c.from('projects').select('id,name,version,data').eq('id', id).single();
  if(error || !data){ toast('Gagal membuka proyek: '+((error&&error.message)||'tidak ditemukan'), 7000); return; }
  cloudWriteLocal(versionRow ? versionRow.data : data.data);
  store.set(CLOUD_PROJ_KEY, JSON.stringify({id:data.id, name:data.name, version:data.version, restored:!!versionRow, rows:cloudRowCount(versionRow ? versionRow.data : data.data)}));
  window.__cloudLeaving = true; location.reload();
}
function cloudDirty(){ return !!cloudProj && cloudHashOf(cloudSnapshot()) !== cloudHash; }
async function cloudSave(manual, note){
  const c = cloudClient(); if(!c || !cloudProj || !cloudUser || cloudBusy) return false;
  const snap = cloudSnapshot(), h = cloudHashOf(snap);
  if(!manual && h === cloudHash) return true;
  // v1.9.7: never overwrite the cloud copy with an empty or much smaller project by accident
  const n = cloudRowCount(snap);
  if(n === 0 || (cloudRows > 0 && n < cloudRows * 0.5)){
    const msg = n === 0 ? 'Data di browser ini kosong' : 'Data di browser ini jauh lebih sedikit ('+n+' baris, di cloud '+cloudRows+' baris)';
    if(!manual){ cloudSetState('tidak disimpan — cek data', 'warn'); return false; }
    if(!confirm(msg+'. Tetap simpan dan TIMPA proyek cloud "'+cloudProj.name+'"?\n\nPilih Batal lalu "Lepas dari cloud" jika lw sedang memulai proyek lain.')) return false;
  }
  cloudBusy = true; cloudSetState('menyimpan…');
  try{
    const now = new Date().toISOString();
    const { data, error } = await c.from('projects').update({data:snap, version:cloudProj.version+1, updated_by:cloudUser.email, updated_at:now, file_name:state.fileName||null}).eq('id', cloudProj.id).eq('version', cloudProj.version).select('version');
    if(error){ cloudSetState('gagal simpan', 'err'); toast('Gagal simpan ke cloud: '+error.message, 8000); return false; }
    if(!data || !data.length){ cloudSetState('ada versi lebih baru', 'warn'); cloudConflict(); return false; }
    cloudProj.version = data[0].version; cloudHash = h; cloudRows = n;
    store.set(CLOUD_PROJ_KEY, JSON.stringify({id:cloudProj.id, name:cloudProj.name, version:cloudProj.version, rows:cloudRows}));
    if(manual || Date.now() - cloudLastHist > CLOUD_HISTORY_EVERY){
      cloudLastHist = Date.now();
      await c.from('project_versions').insert({project_id:cloudProj.id, version:cloudProj.version, data:snap, note:note||(manual?'Simpan manual':'Simpan otomatis'), saved_by:cloudUser.email});
    }
    cloudSetState('tersimpan '+cloudTime(now).split(' ').pop());
    return true;
  }catch(e){ cloudSetState('gagal simpan', 'err'); toast('Gagal simpan ke cloud: '+(e.message||e), 8000); return false; }
  finally{ cloudBusy = false; }
}
async function cloudConflict(){
  const c = cloudClient(); const { data } = await c.from('projects').select('version,updated_by,updated_at').eq('id', cloudProj.id).single();
  cloudShow('<h3>Ada versi lebih baru di cloud</h3><p style="font-size:13px;line-height:1.5">Proyek <b>'+esc(cloudProj.name)+'</b> sudah disimpan oleh <b>'+esc((data&&data.updated_by)||'orang lain')+'</b> pada '+esc(cloudTime(data&&data.updated_at))+' setelah lw membukanya. Simpan otomatis dihentikan supaya tidak saling menimpa.</p>'+
    '<p style="font-size:13px;line-height:1.5">Pilih <b>Muat versi terbaru</b> (perubahan lw sejak itu hilang, kecuali di-backup dulu lewat Export → Master Workbook), atau <b>Timpa dengan versi saya</b> (versi mereka tetap tersimpan di Riwayat).</p>'+
    '<div class="row"><button class="btn" data-c="close">Nanti</button><button class="btn" data-c="reload">Muat versi terbaru</button><button class="btn primary" data-c="overwrite">Timpa dengan versi saya</button></div>');
}
async function cloudOverwrite(){
  const c = cloudClient(); const { data } = await c.from('projects').select('version').eq('id', cloudProj.id).single();
  if(data){ cloudProj.version = data.version; cloudHash = ''; await cloudSave(true, 'Menimpa versi lain'); }
}
async function cloudNew(name){
  const c = cloudClient(); name = String(name||'').trim(); if(!name){ toast('Isi nama proyek.'); return; }
  const snap = cloudSnapshot();
  const { data, error } = await c.from('projects').insert({name, data:snap, version:1, file_name:state.fileName||null, created_by:cloudUser.email, updated_by:cloudUser.email}).select('id,name,version').single();
  if(error){ toast('Gagal membuat proyek cloud: '+error.message, 8000); return; }
  cloudProj = {id:data.id, name:data.name, version:data.version}; cloudHash = cloudHashOf(snap); cloudRows = cloudRowCount(snap); cloudLastHist = Date.now();
  store.set(CLOUD_PROJ_KEY, JSON.stringify(cloudProj));
  await c.from('project_versions').insert({project_id:cloudProj.id, version:1, data:snap, note:'Proyek dibuat', saved_by:cloudUser.email});
  cloudSetState('tersimpan'); toast('Proyek "'+name+'" tersimpan di cloud. Perubahan berikutnya tersimpan otomatis.', 5000); cloudPanel();
}
function cloudDetach(){ if(cloudDirty() && !confirm('Ada perubahan yang belum tersimpan ke cloud. Tetap lepaskan?')) return; cloudProj = null; cloudHash = ''; cloudRows = 0; try{ localStorage.removeItem(CLOUD_PROJ_KEY); }catch(e){} cloudSetState(''); cloudPanel(); toast('Proyek dilepas dari cloud. Data tetap ada di browser ini.'); }
// v1.10.0: delete a team project (and, by cascade, its history and save log). Typed-name confirmation.
async function cloudDelete(id, name){
  const c = cloudClient(); if(!c) return;
  const active = cloudProj && cloudProj.id === id;
  const typed = window.prompt ? prompt('HAPUS PERMANEN proyek "'+name+'" dari cloud?\n\nSeluruh riwayat versi dan log simpannya ikut terhapus dan tidak bisa dikembalikan.'+(active?' Data di browser ini tetap ada, tetapi tidak lagi terhubung ke cloud.':'')+'\n\nKetik nama proyek persis untuk mengonfirmasi:', '') : null;
  if(typed === null) return;
  if(typed.trim() !== String(name).trim()){ toast('Nama tidak cocok. Proyek tidak dihapus.', 5000); return; }
  const { data, error } = await c.from('projects').delete().eq('id', id).select('id');
  if(error){ toast('Gagal menghapus: '+error.message, 8000); return; }
  if(!data || !data.length){ toast('Proyek tidak terhapus (tidak ditemukan atau tidak ada akses).', 7000); return; }
  if(active){ cloudProj = null; cloudHash = ''; cloudRows = 0; try{ localStorage.removeItem(CLOUD_PROJ_KEY); }catch(e){} cloudSetState(''); }
  toast('Proyek "'+name+'" dihapus dari cloud.', 5000); cloudPanel();
}
function cloudShow(html){
  let m = $('cloudDlg');
  if(!m){ m = document.createElement('div'); m.className = 'modal hidden'; m.id = 'cloudDlg'; m.setAttribute('role','dialog'); m.innerHTML = '<div class="dlg" style="max-width:640px;width:min(640px,94vw);max-height:86vh;overflow:auto"></div>'; document.body.appendChild(m);
    m.addEventListener('click', cloudClick); m.addEventListener('keydown', e=>{ if(e.key==='Escape') m.classList.add('hidden'); if(e.key==='Enter' && (e.target.id==='cloudUser' || e.target.id==='cloudPass')){ e.preventDefault(); cloudClick({target:m.querySelector('[data-c=login]')}); } if(e.key==='Enter' && e.target.id==='cloudNewName'){ e.preventDefault(); cloudClick({target:m.querySelector('[data-c=new]')}); } }); }
  m.querySelector('.dlg').innerHTML = html; m.classList.remove('hidden');
}
async function cloudPanel(){
  const c = cloudClient(), P = s => '<p style="font-size:13px;line-height:1.5;color:#334155">'+s+'</p>';
  if(!c){ cloudShow('<h3>Simpan ke cloud</h3>'+P('Library cloud tidak termuat. Periksa koneksi internet lalu muat ulang halaman.')+'<div class="row"><button class="btn" data-c="close">Tutup</button></div>'); return; }
  if(location.protocol === 'file:'){ cloudShow('<h3>Simpan ke cloud</h3>'+P('Login cloud hanya bisa dari versi online aplikasi: <b>https://betterzoo.github.io/flowchart-bpa/</b>. File HTML yang dibuka langsung dari komputer tetap bisa dipakai seperti biasa (tersimpan di browser ini).')+'<div class="row"><button class="btn" data-c="close">Tutup</button></div>'); return; }
  cloudShow('<h3>Simpan ke cloud</h3>'+P('Memuat…'));
  await cloudRefreshUser();
  if(!cloudUser){
    const inp = 'style="width:100%;font-size:16px;padding:8px;margin:4px 0 10px;box-sizing:border-box"';
    cloudShow('<h3>Masuk ke cloud tim</h3>'+P('Masuk dengan username dan password dari admin.')+
      '<label for="cloudUser" style="font-size:13px">Username</label><input id="cloudUser" autocomplete="username" autocapitalize="none" spellcheck="false" '+inp+' placeholder="mis. dea">'+
      '<label for="cloudPass" style="font-size:13px">Password</label><input id="cloudPass" type="password" autocomplete="current-password" '+inp+'>'+
      '<div id="cloudMsg" style="font-size:13px"></div><div class="row"><button class="btn" data-c="close">Tutup</button><button class="btn primary" data-c="login">Masuk</button></div>');
    setTimeout(()=>$('cloudUser')&&$('cloudUser').focus(), 0); return;
  }
  if(!cloudMember){
    cloudShow('<h3>Belum terdaftar</h3>'+P('Lw masuk sebagai <b>'+esc(cloudUserName(cloudUser.email))+'</b>, tetapi akun ini belum terdaftar sebagai anggota tim. Minta admin menambahkan <code>'+esc(cloudUser.email)+'</code> ke tabel <code>members</code> di Supabase, lalu buka panel ini lagi.')+
      '<p style="font-size:12px;color:#64748b">Detail pengecekan: '+esc(cloudMemberErr||'-')+'</p><div class="row" style="justify-content:flex-start"><button class="btn" data-c="recheck">Cek lagi</button></div>'+
      '<div class="row"><button class="btn" data-c="logout">Keluar</button><button class="btn primary" data-c="close">Tutup</button></div>'); return;
  }
  const { data:list, error } = await c.from('projects').select('id,name,version,updated_at,updated_by,file_name').order('updated_at', {ascending:false});
  let h = '<h3>Simpan ke cloud</h3>'+P('Masuk sebagai <b>'+esc(cloudUserName(cloudUser.email))+'</b> · <a href="#" data-c="passwd">Ganti password</a> · <a href="#" data-c="logout">Keluar</a>');
  if(cloudProj){
    h += '<div style="border:1px solid var(--line,#d6dbe3);border-radius:8px;padding:10px 12px;margin:8px 0"><div style="font-size:12px;color:#64748b">Proyek aktif</div><div style="font-size:16px;font-weight:600">'+esc(cloudProj.name)+'</div>'+
      '<div style="font-size:12.5px;color:#334155;margin:4px 0 8px">'+(cloudDirty()?'Ada perubahan yang belum tersimpan.':'Semua perubahan sudah tersimpan.')+' Versi '+cloudProj.version+'.</div>'+
      '<label style="font-size:13px;display:block;margin-bottom:8px"><input type="checkbox" id="cloudAutoCk"'+(cloudAuto()?' checked':'')+'> Simpan otomatis setiap ada perubahan</label>'+
      '<div class="row" style="justify-content:flex-start"><button class="btn primary" data-c="save">Simpan sekarang</button><button class="btn" data-c="history">Riwayat versi</button><button class="btn" data-c="detach">Lepas dari cloud</button></div></div>';
  } else {
    h += '<div style="border:1px solid var(--line,#d6dbe3);border-radius:8px;padding:10px 12px;margin:8px 0"><div style="font-size:13px;font-weight:600;margin-bottom:4px">Simpan pekerjaan di browser ini sebagai proyek cloud baru</div>'+
      (state.rows.length?'':'<div style="font-size:12.5px;color:#a21d32;margin-bottom:6px">Belum ada workbook yang dibuka.</div>')+
      '<input id="cloudNewName" style="width:100%;font-size:16px;padding:8px;margin:4px 0 8px" placeholder="mis. PELNI — SOP HC 2026" value="'+esc(String(state.fileName||'').replace(/\.xlsx?$/i,''))+'">'+
      '<div class="row" style="justify-content:flex-start"><button class="btn primary" data-c="new"'+(state.rows.length?'':' disabled')+'>Simpan sebagai proyek baru</button></div></div>';
  }
  h += '<div style="font-size:13px;font-weight:600;margin:14px 0 6px">Proyek tim</div>';
  if(error) h += P('Gagal memuat daftar: '+esc(error.message));
  else if(!list.length) h += P('Belum ada proyek di cloud.');
  else h += '<div>'+list.map(p=>'<div style="display:flex;gap:8px;align-items:center;justify-content:space-between;border-top:1px solid var(--line,#e5e7eb);padding:8px 0"><div style="min-width:0"><div style="font-weight:600">'+esc(p.name)+(cloudProj&&cloudProj.id===p.id?' <span style="font-size:11px;color:#0a7d32">(aktif)</span>':'')+'</div><div style="font-size:12px;color:#64748b">'+esc(cloudTime(p.updated_at))+' · '+esc(p.updated_by||'')+(p.file_name?' · '+esc(p.file_name):'')+'</div></div>'+
    '<div style="display:flex;gap:6px;flex:none">'+(cloudProj&&cloudProj.id===p.id?'':'<button class="btn" data-c="open" data-id="'+esc(p.id)+'">Buka</button>')+'<button class="btn" data-c="rename" data-id="'+esc(p.id)+'" data-name="'+esc(p.name)+'">Ganti nama</button><button class="btn" data-c="delete" data-id="'+esc(p.id)+'" data-name="'+esc(p.name)+'" style="color:#a21d32" title="Hapus proyek ini dari cloud beserta seluruh riwayatnya">Hapus</button></div></div>').join('')+'</div>';
  if(!error && list.length >= 2) h += '<div class="row" style="justify-content:flex-start;margin-top:8px"><button class="btn" data-c="merge" title="Gabungkan beberapa proyek menjadi satu proyek baru">⇄ Gabungkan proyek…</button></div>';
  h += '<div class="row"><button class="btn" data-c="close">Tutup</button></div>';
  cloudShow(h);
  const ck = $('cloudAutoCk'); if(ck) ck.onchange = () => { store.set(CLOUD_AUTO_KEY, ck.checked?'1':'0'); toast(ck.checked?'Simpan otomatis aktif.':'Simpan otomatis dimatikan. Pakai "Simpan sekarang".'); };
}
async function cloudHistory(){
  const c = cloudClient(); if(!cloudProj) return;
  const { data, error } = await c.from('project_versions').select('id,version,note,saved_by,saved_at').eq('project_id', cloudProj.id).order('id', {ascending:false}).limit(30);
  let h = '<h3>Riwayat versi · '+esc(cloudProj.name)+'</h3><p style="font-size:13px;color:#334155">Disimpan saat Simpan manual, saat proyek dibuat, dan otomatis paling sering tiap 15 menit. 30 versi terakhir dipertahankan.</p>';
  // v1.9.9: the project's current state, which often has no snapshot of its own (saved <15 min after the last one)
  try{
    const cur = await c.from('projects').select('version,updated_by,updated_at').eq('id', cloudProj.id).single();
    if(cur.data && !(data||[]).some(v=>v.version===cur.data.version))
      h += '<div style="display:flex;justify-content:space-between;align-items:center;gap:8px;border-top:1px solid var(--line,#e5e7eb);padding:8px 0;background:#f0fdf4"><div style="padding-left:6px"><div style="font-weight:600">Versi '+cur.data.version+' · saat ini · '+esc(cloudTime(cur.data.updated_at))+'</div><div style="font-size:12px;color:#64748b">'+esc(cloudUserName(cur.data.updated_by))+' · Isi proyek sekarang (belum ada snapshot terpisah)</div></div></div>';
  }catch(e){}
  if(error) h += '<p>Gagal memuat: '+esc(error.message)+'</p>';
  else h += (data||[]).map(v=>'<div style="display:flex;justify-content:space-between;align-items:center;gap:8px;border-top:1px solid var(--line,#e5e7eb);padding:8px 0"><div><div style="font-weight:600">Versi '+v.version+' · '+esc(cloudTime(v.saved_at))+'</div><div style="font-size:12px;color:#64748b">'+esc(cloudUserName(v.saved_by))+' · '+esc(v.note||'')+'</div></div><button class="btn" data-c="restore" data-id="'+v.id+'">Pulihkan</button></div>').join('') || '<p>Belum ada riwayat.</p>';
  // v1.9.8: every save (who, when, which parts) from public.project_saves, filled by a database trigger.
  // Older databases without that table simply skip this section.
  try{
    const r = await c.from('project_saves').select('version,saved_by,saved_at,changed').eq('project_id', cloudProj.id).order('id', {ascending:false}).limit(60);
    if(!r.error && r.data && r.data.length){
      h += '<div style="font-size:13px;font-weight:600;margin:16px 0 4px">Log simpan (semua simpan, tanpa tombol pulihkan)</div>'+
        '<table style="width:100%;border-collapse:collapse;font-size:12.5px"><tr style="text-align:left;color:#64748b"><th style="padding:4px 6px 4px 0">Versi</th><th style="padding:4px 6px">Waktu</th><th style="padding:4px 6px">Oleh</th><th style="padding:4px 0 4px 6px">Bagian berubah</th></tr>'+
        r.data.map(s=>'<tr style="border-top:1px solid var(--line,#e5e7eb)"><td style="padding:4px 6px 4px 0">'+s.version+'</td><td style="padding:4px 6px;white-space:nowrap">'+esc(cloudTime(s.saved_at))+'</td><td style="padding:4px 6px">'+esc(cloudUserName(s.saved_by))+'</td><td style="padding:4px 0 4px 6px">'+esc(cloudPartNames(s.changed))+'</td></tr>').join('')+'</table>';
    }
  }catch(e){}
  h += '<div class="row"><button class="btn" data-c="back">Kembali</button></div>';
  cloudShow(h);
}
async function cloudRestore(vid){
  const c = cloudClient(); const { data, error } = await c.from('project_versions').select('version,data').eq('id', +vid).single();
  if(error || !data){ toast('Gagal memuat versi: '+((error&&error.message)||''), 7000); return; }
  if(!confirm('Pulihkan versi '+data.version+'? Pekerjaan saat ini diganti; versi sekarang tetap ada di Riwayat setelah tersimpan.')) return;
  if(cloudDirty()) await cloudSave(true, 'Sebelum memulihkan versi '+data.version);
  cloudOpen(cloudProj.id, data);
}
async function cloudClick(e){
  const b = e.target && e.target.closest ? e.target.closest('[data-c]') : null; if(!b) return; if(e.preventDefault) e.preventDefault();
  const a = b.dataset.c, c = cloudClient(), m = $('cloudDlg');
  if(a==='close'){ m.classList.add('hidden'); return; }
  if(a==='login'){
    const u = ($('cloudUser').value||'').trim(), pw = $('cloudPass').value||'';
    if(!u || /\s/.test(u)){ $('cloudMsg').textContent = 'Isi username (tanpa spasi).'; return; }
    if(!pw){ $('cloudMsg').textContent = 'Isi password.'; return; }
    b.disabled = true; $('cloudMsg').textContent = 'Masuk…';
    const { data, error } = await c.auth.signInWithPassword({email: cloudLoginEmail(u), password: pw});
    b.disabled = false;
    if(error){ $('cloudMsg').innerHTML = '<span style="color:#a21d32">'+esc(/invalid/i.test(error.message||'') ? 'Username atau password salah.' : 'Gagal masuk: '+error.message)+'</span>'; return; }
    cloudUser = data && data.user ? data.user : cloudUser; cloudPanel();
    return;
  }
  if(a==='passwd'){
    cloudShow('<h3>Ganti password</h3><p style="font-size:13px;color:#334155">Akun: <b>'+esc(cloudUserName(cloudUser&&cloudUser.email))+'</b>. Minimal 8 karakter.</p>'+
      '<label for="cloudNewPass" style="font-size:13px">Password baru</label><input id="cloudNewPass" type="password" autocomplete="new-password" style="width:100%;font-size:16px;padding:8px;margin:4px 0 10px;box-sizing:border-box">'+
      '<label for="cloudNewPass2" style="font-size:13px">Ulangi password baru</label><input id="cloudNewPass2" type="password" autocomplete="new-password" style="width:100%;font-size:16px;padding:8px;margin:4px 0 10px;box-sizing:border-box">'+
      '<div id="cloudMsg" style="font-size:13px"></div><div class="row"><button class="btn" data-c="back">Kembali</button><button class="btn primary" data-c="setpass">Simpan password</button></div>');
    setTimeout(()=>$('cloudNewPass')&&$('cloudNewPass').focus(), 0); return;
  }
  if(a==='setpass'){
    const p1 = $('cloudNewPass').value||'', p2 = $('cloudNewPass2').value||'';
    if(p1.length < 8){ $('cloudMsg').textContent = 'Password minimal 8 karakter.'; return; }
    if(p1 !== p2){ $('cloudMsg').textContent = 'Kedua password tidak sama.'; return; }
    b.disabled = true; const { error } = await c.auth.updateUser({password: p1}); b.disabled = false;
    if(error){ $('cloudMsg').innerHTML = '<span style="color:#a21d32">Gagal: '+esc(error.message)+'</span>'; return; }
    toast('Password diganti.'); cloudPanel(); return;
  }
  if(a==='logout'){ if(cloudDirty() && !confirm('Ada perubahan yang belum tersimpan ke cloud. Tetap keluar?')) return; await c.auth.signOut(); cloudUser = null; cloudSetState(''); cloudPanel(); return; }
  if(a==='new'){ await cloudNew($('cloudNewName').value); return; }
  if(a==='save'){ const ok = await cloudSave(true); if(ok) toast('Tersimpan ke cloud.'); cloudPanel(); return; }
  if(a==='open'){ cloudOpen(b.dataset.id); return; }
  if(a==='rename'){ const nm = (window.prompt ? prompt('Nama baru proyek:', b.dataset.name) : null); if(!nm || !nm.trim()) return; const { error } = await c.from('projects').update({name:nm.trim()}).eq('id', b.dataset.id); if(error) toast('Gagal ganti nama: '+error.message); else { if(cloudProj && cloudProj.id===b.dataset.id){ cloudProj.name = nm.trim(); store.set(CLOUD_PROJ_KEY, JSON.stringify(cloudProj)); cloudSetState(cloudState); } cloudPanel(); } return; }
  if(a==='delete'){ await cloudDelete(b.dataset.id, b.dataset.name); return; }
  if(a==='merge'){ cloudMergeStart(); return; }
  if(a==='mapStart'){ b.disabled = true; try{ await cloudMapStart(); } finally { if(b.isConnected) b.disabled = false; } return; }
  if(a==='mapAll'){ const i = +b.dataset.i; for(const l3 of cloudMapCtx.l3s) if(l3.has[i]){ cloudMapCtx.map[l3.l3no] = i; l3.touched = true; } cloudMapRender(); return; }
  if(a==='mapCheck'){ cloudMapCheck(); return; }
  if(a==='mapBack'){ cloudMapRender(); return; }
  if(a==='mapCreate'){ b.disabled = true; try{ await cloudMapCreate(); } finally { if(b.isConnected) b.disabled = false; } return; }
  if(a==='mergeDiff'){ b.disabled = true; try{ await cloudMergeDiff(); } finally { b.disabled = false; } return; }
  if(a==='mergeCreate'){ b.disabled = true; try{ await cloudMergeCreate(); } finally { if(b.isConnected) b.disabled = false; } return; }
  if(a==='detach'){ cloudDetach(); return; }
  if(a==='history'){ cloudHistory(); return; }
  if(a==='back'||a==='recheck'){ cloudPanel(); return; }
  if(a==='restore'){ cloudRestore(b.dataset.id); return; }
  if(a==='reload'){ m.classList.add('hidden'); cloudOpen(cloudProj.id); return; }
  if(a==='overwrite'){ m.classList.add('hidden'); cloudOverwrite(); return; }
}
(async function cloudInit(){
  const anchor = $('btnExport'); if(!anchor) return;
  const btn = document.createElement('button'); btn.className = 'btn'; btn.id = 'btnCloud'; btn.type = 'button'; btn.textContent = '☁ Cloud'; btn.onclick = cloudPanel;
  anchor.parentNode.insertBefore(btn, anchor);
  try{ const p = JSON.parse(store.get(CLOUD_PROJ_KEY)||'null'); if(p && p.id){ cloudProj = {id:p.id, name:p.name, version:p.version}; const s0 = cloudSnapshot(); cloudHash = cloudHashOf(s0); cloudRows = p.rows || cloudRowCount(s0); if(p.restored) cloudHash = ''; } }catch(e){}
  cloudSetState(cloudProj ? '' : '');
  const c = cloudClient(); if(!c) return;
  try{ if(/[?&]code=/.test(location.search)) { await c.auth.getSession(); history.replaceState(null, '', location.pathname + location.hash); } }catch(e){}
  await cloudRefreshUser();
  if(c.auth.onAuthStateChange) c.auth.onAuthStateChange((ev, session) => { cloudUser = session ? session.user : null; if(ev==='SIGNED_IN') cloudRefreshUser().then(()=>{ if(!cloudProj) toast('Masuk cloud sebagai '+cloudUser.email+'. Klik ☁ Cloud untuk menyimpan atau membuka proyek.', 6000); }); });
  if(cloudProj && cloudUser && cloudMember){
    // someone else may have saved while this browser was closed
    try{ const { data } = await c.from('projects').select('version,updated_by,updated_at').eq('id', cloudProj.id).single();
      if(data && data.version > cloudProj.version){ if(cloudHash && cloudHashOf(cloudSnapshot())===cloudHash){ cloudOpen(cloudProj.id); return; } cloudSetState('ada versi lebih baru','warn'); cloudConflict(); }
      else if(!data){ cloudSetState('proyek tidak ditemukan','err'); }
      else cloudSetState('');
    }catch(e){}
  } else if(cloudProj && !cloudUser){ cloudSetState('belum masuk','warn'); }
  setInterval(() => { if(cloudProj && cloudUser && cloudMember && cloudAuto() && !cloudBusy && cloudState!=='ada versi lebih baru' && document.visibilityState!=='hidden') cloudSave(false); }, 8000);
  window.addEventListener('beforeunload', e => { if(!window.__cloudLeaving && cloudProj && cloudUser && cloudDirty()){ e.preventDefault(); e.returnValue = ''; } });
})();

/* ================= v1.11.0: Gabungkan (merge) beberapa proyek cloud menjadi proyek baru =================
   Each project snapshot is split into small units (one L4's AS-IS rows, one L4's TO-BE draft, one activity's
   SLA/notes, one lane mapping, …). Per unit:
   - all projects equal → taken as is;
   - a project "changed" the unit when its value differs from the COMMON ANCESTOR: the oldest first snapshot
     among the selected projects (v1.11.1; before, any project's first snapshot counted, so a merged project —
     whose first snapshot already holds everything — made other projects' "nothing here" look like a deletion).
     Absent counts as a value. Only changed values compete; if none changed, all compete;
   - the winner is the project whose LOG (project_saves) shows the latest save touching that part
     (fallback: the project's last save time). Every choice can be overridden in the preview.
   The source projects are never modified. */
let cloudMergeCtx = null;
const CLOUD_MERGE_PART = {data:'Data AS-IS', elab:'Rancangan TO-BE', extra:'SLA/Output', roles:'Peran', posmap:'Pemetaan jabatan', decmap:'Keputusan', unitmap:'Unit kerja', lanemodes:'Mode swimlane', poson:'Tampilan jabatan', analysis:'Analisis', changeRegister:'Register perubahan', sem:'Jenis simbol BPMN', other:'Lainnya'};
function cloudMergeParse(v){ try{ return JSON.parse(v); }catch(e){ return null; } }
// snapshot → Map(unitKey → JSON string or null) plus helpers for rebuilding
function cloudMergeSplit(snap){
  const U = new Map(), J = v => JSON.stringify(v); snap = snap || {};
  for(const [k, raw] of Object.entries(snap)){
    if(k.startsWith('bpa_bpmn_semantics_')){ const o = cloudMergeParse(raw) || {}; for(const [ik, iv] of Object.entries(o)) U.set('sem|'+ik, J(iv)); continue; }
    if(!k.startsWith('flowbpa:')){ U.set('other|'+k, J(raw)); continue; }
    const part = k.slice(8), v = cloudMergeParse(raw);
    if(part === 'usedCodes') continue; // union, handled separately
    if(part === 'data' && v && Array.isArray(v.rows)){
      const by = new Map(); for(const r of v.rows){ const l4 = String(r.l4no||''); if(!by.has(l4)) by.set(l4, []); by.get(l4).push(r); }
      for(const [l4, rows] of by) U.set('data|'+l4, J(rows));
      for(const [dk, dv] of Object.entries(v)) if(dk !== 'rows') U.set('datameta|'+dk, J(dv));
    } else if(part === 'roles' && v && typeof v === 'object'){
      for(const [sk, sv] of Object.entries(v)) for(const [ik, iv] of Object.entries(sv||{})) U.set('roles|'+sk+'|'+ik, J(iv));
    } else if((part === 'analysis' || part === 'changeRegister') && Array.isArray(v)){
      v.forEach((r, i) => U.set(part+'|'+(r && r.id ? r.id : '#'+i), J(r)));
    } else if(v && typeof v === 'object' && !Array.isArray(v)){
      for(const [ik, iv] of Object.entries(v)) U.set(part+'|'+ik, J(iv));
    } else U.set(part+'|', J(v === null ? raw : v));
  }
  return U;
}
function cloudMergeJoin(units, used, semKey, order){
  const P = s => cloudMergeParse(s), out = {}, obj = {}, data = {meta:{}, by:new Map()}, roles = {alias:{}, added:{}, hidden:{}, order:{}}, lists = {analysis:[], changeRegister:[]}, sem = {};
  for(const [key, val] of units){ if(val == null) continue;
    const [part, ...rest] = key.split('|'), k1 = rest.join('|');
    if(part === 'data') data.by.set(k1, P(val));
    else if(part === 'datameta') data.meta[k1] = P(val);
    else if(part === 'roles'){ const [sk, ...ik] = rest; (roles[sk] = roles[sk] || {})[ik.join('|')] = P(val); }
    else if(part === 'analysis' || part === 'changeRegister') lists[part].push(P(val));
    else if(part === 'sem') sem[k1] = P(val);
    else if(part === 'other') out[k1] = P(val);
    else if(k1 === '' ) out['flowbpa:'+part] = typeof P(val) === 'string' ? P(val) : JSON.stringify(P(val));
    else (obj[part] = obj[part] || {})[k1] = P(val);
  }
  const l4s = [...data.by.keys()].sort((a,b)=>{ const ia = order.indexOf(a), ib = order.indexOf(b); return (ia<0?1e9:ia)-(ib<0?1e9:ib) || (typeof cmpCode==='function' ? cmpCode(a,b) : a.localeCompare(b)); });
  out['flowbpa:data'] = JSON.stringify({...data.meta, rows: l4s.flatMap(l4 => data.by.get(l4))});
  for(const [part, o] of Object.entries(obj)) out['flowbpa:'+part] = JSON.stringify(o);
  out['flowbpa:roles'] = JSON.stringify(roles);
  out['flowbpa:analysis'] = JSON.stringify(lists.analysis); out['flowbpa:changeRegister'] = JSON.stringify(lists.changeRegister);
  out['flowbpa:usedCodes'] = JSON.stringify([...used].sort());
  if(semKey) out[semKey] = JSON.stringify(sem);
  return out;
}
function cloudMergeLabel(key, names){
  const [part, ...rest] = key.split('|'), k = rest.join('|'), P = CLOUD_MERGE_PART[part] || part;
  if(part === 'data' || part === 'elab' || part === 'posmap' || part === 'unitmap' || part === 'decmap' || part === 'lanemodes'){ const n = names.get(k); return P+' '+k+(n?' — '+n:''); }
  if(part === 'extra') return P+' '+k+(names.get(k)?' — '+names.get(k):'');
  if(part === 'roles') return P+' ('+rest[0]+') '+rest.slice(1).join('|')+(names.get(rest.slice(1).join('|'))?' — '+names.get(rest.slice(1).join('|')):'');
  return P+(k?' '+k:'');
}
function cloudMergeLogPart(key){ const p = key.split('|')[0]; return p === 'datameta' ? 'data' : p === 'sem' ? 'bpa_bpmn_semantics' : p; }
async function cloudMergeLoad(ids){
  const c = cloudClient(), docs = [];
  for(const id of ids){
    const { data:p, error } = await c.from('projects').select('id,name,version,data,updated_at,updated_by').eq('id', id).single();
    if(error || !p) throw Error('Gagal memuat proyek: '+((error&&error.message)||id));
    let base = null; try{ const r = await c.from('project_versions').select('version,data,saved_at').eq('project_id', id).order('id', {ascending:true}).limit(1); base = r.data && r.data[0] ? r.data[0] : null; }catch(e){}
    let log = []; try{ const r = await c.from('project_saves').select('saved_at,changed').eq('project_id', id).order('id', {ascending:false}).limit(500); log = (!r.error && r.data) || []; }catch(e){}
    docs.push({id:p.id, name:p.name, version:p.version, updated_at:p.updated_at, updated_by:p.updated_by, snap:p.data, units:cloudMergeSplit(p.data), base:base ? cloudMergeSplit(base.data) : null, baseVer:base ? base.version : null, baseAt:base ? Date.parse(base.saved_at) || 0 : 0, log});
  }
  return docs;
}
// last time (ms) this project's log shows a save touching the part; falls back to the project's last save
function cloudMergeTime(doc, part){
  let t = 0; for(const s of doc.log) if((s.changed||[]).some(x => x === part || (part === 'bpa_bpmn_semantics' && /^bpa_bpmn_semantics/.test(x)))){ const v = Date.parse(s.saved_at); if(v > t) t = v; }
  return t || Date.parse(doc.updated_at) || 0;
}
function cloudMergePlan(docs){
  const keys = new Set(); docs.forEach(d => d.units.forEach((_, k) => keys.add(k)));
  const anc = docs.filter(d => d.base).sort((a, b) => a.baseAt - b.baseAt)[0] || null;
  const units = [];
  for(const key of keys){
    const vals = docs.map(d => d.units.has(key) ? d.units.get(key) : null);
    if(vals.every(v => v === vals[0])){ units.push({key, vals, pick:0, same:true}); continue; }
    const av = anc ? (anc.base.has(key) ? anc.base.get(key) : null) : undefined;
    const changed = vals.map(v => anc ? v !== av : false);
    const part = cloudMergeLogPart(key), times = docs.map(d => cloudMergeTime(d, part));
    let cand = docs.map((_, i) => i).filter(i => changed[i]); if(!cand.length) cand = docs.map((_, i) => i);
    const pick = cand.reduce((a, b) => times[b] > times[a] ? b : a);
    const distinctChanged = new Set(cand.filter(i => changed[i]).map(i => vals[i])).size;
    units.push({key, vals, pick, changed, times, conflict: distinctChanged > 1 || (!changed.some(Boolean))});
  }
  return units;
}
// TO-BE-only activity codes: SLA/notes follow the project whose TO-BE draft of that L4 was taken
function cloudMergeAlign(plan){
  const byKey = new Map(plan.map(u => [u.key, u]));
  for(const u of plan){ if(u.same || !u.key.startsWith('extra|') || u.touched) continue;
    const code = u.key.slice(6), l4 = code.split('.').slice(0, -1).join('.'), asis = byKey.get('data|'+l4), el = byKey.get('elab|'+l4);
    const inAsis = asis && asis.vals[asis.pick] && asis.vals[asis.pick].includes('"l5no":"'+code+'"');
    if(!inAsis && el && !el.same && u.vals[el.pick] != null) u.pick = el.pick;
  }
}
function cloudMergeNames(docs){
  const names = new Map();
  for(const d of docs){ try{ const rows = JSON.parse(d.snap['flowbpa:data']).rows; for(const r of rows){ if(r.l4no && !names.has(r.l4no)) names.set(r.l4no, r.l4); if(r.l5no && !names.has(r.l5no)) names.set(r.l5no, r.l5); } }catch(e){} }
  return names;
}
async function cloudMergeStart(){
  const c = cloudClient(); const { data:list, error } = await c.from('projects').select('id,name,version,updated_at,updated_by').order('updated_at', {ascending:false});
  if(error){ toast('Gagal memuat daftar: '+error.message, 7000); return; }
  cloudShow('<h3>Gabungkan proyek</h3><p style="font-size:13px;line-height:1.5;color:#334155">Pilih 2 proyek atau lebih. Hasilnya proyek <b>baru</b>; proyek asal tidak diubah.</p><ul style="font-size:12.5px;line-height:1.5;color:#334155;padding-left:18px;margin:4px 0 8px"><li><b>Bandingkan otomatis</b>: bagian yang berbeda diambil dari perubahan terbaru menurut log simpan; bisa diubah sebelum dibuat.</li><li><b>Mapping per L3</b>: lw pilih dokumen sumber untuk tiap L3; tanpa membandingkan isi.</li></ul>'+
    list.map(p => '<label style="display:flex;gap:8px;align-items:flex-start;border-top:1px solid var(--line,#e5e7eb);padding:7px 0;font-size:13.5px"><input type="checkbox" class="cmSel" value="'+esc(p.id)+'" style="margin-top:3px"><span><b>'+esc(p.name)+'</b><br><span style="font-size:12px;color:#64748b">v'+p.version+' · '+esc(cloudTime(p.updated_at))+' · '+esc(cloudUserName(p.updated_by))+'</span></span></label>').join('')+
    '<div id="cloudMsg" style="font-size:13px;margin-top:8px"></div><div class="row"><button class="btn" data-c="back">Kembali</button><button class="btn" data-c="mapStart">Mapping per L3</button><button class="btn primary" data-c="mergeDiff">Bandingkan otomatis</button></div>');
}
async function cloudMergeDiff(){
  const ids = [...document.querySelectorAll('.cmSel:checked')].map(x => x.value);
  if(ids.length < 2){ $('cloudMsg').textContent = 'Pilih minimal 2 proyek.'; return; }
  $('cloudMsg').textContent = 'Memuat dan membandingkan '+ids.length+' proyek…';
  let docs; try{ docs = await cloudMergeLoad(ids); }catch(e){ $('cloudMsg').textContent = e.message; return; }
  const plan = cloudMergePlan(docs); cloudMergeAlign(plan);
  cloudMergeCtx = {docs, plan, names:cloudMergeNames(docs)};
  cloudMergeRender();
}
function cloudMergeRender(){
  const {docs, plan, names} = cloudMergeCtx, D = docs;
  const diff = plan.filter(u => !u.same), conf = diff.filter(u => u.conflict), auto = diff.filter(u => !u.conflict);
  const opt = (u, i) => {
    const v = u.vals[i], tag = v == null ? 'tidak ada' : u.changed[i] ? 'diubah, log '+cloudTime(u.times[i]) : 'sama dengan awal';
    return '<option value="'+i+'"'+(u.pick === i ? ' selected' : '')+'>'+esc(D[i].name)+' ('+esc(tag)+')</option>';
  };
  const row = u => { const ui = plan.indexOf(u); return '<div style="border-top:1px solid var(--line,#e5e7eb);padding:6px 0;font-size:12.5px"><div style="font-weight:600">'+esc(cloudMergeLabel(u.key, names))+'</div><select data-mu="'+ui+'" style="width:100%;font-size:13px;margin-top:3px">'+D.map((_, i) => opt(u, i)).join('')+'</select></div>'; };
  const sum = {}; for(const u of auto){ const k = D[u.pick].name; sum[k] = (sum[k]||0) + 1; }
  const defName = 'Gabungan — '+D.map(d => d.name).join(' + ');
  cloudShow('<h3>Hasil perbandingan</h3>'+
    '<p style="font-size:13px;line-height:1.5;color:#334155">'+D.length+' proyek, '+plan.length+' bagian dicek: <b>'+(plan.length-diff.length)+'</b> sama, <b>'+auto.length+'</b> berbeda dan hanya diubah di satu proyek (diambil otomatis), <b>'+conf.length+'</b> perlu perhatian (diubah di lebih dari satu proyek; default: log terbaru).</p>'+
    (D.some(d => !d.base) ? '<p style="font-size:12px;color:#a16207">Sebagian proyek tidak punya snapshot awal; perbandingan memakai waktu log saja.</p>' : '')+
    '<details open><summary style="font-weight:600;font-size:13px;margin:8px 0">Perlu perhatian ('+conf.length+')</summary>'+(conf.length ? conf.map(row).join('') : '<p style="font-size:12.5px;color:#64748b">Tidak ada.</p>')+'</details>'+
    '<details><summary style="font-weight:600;font-size:13px;margin:8px 0">Diambil otomatis ('+auto.length+')'+(auto.length?' — '+Object.entries(sum).map(([n, k]) => esc(n)+': '+k).join(', '):'')+'</summary>'+auto.map(row).join('')+'</details>'+
    '<label for="cmName" style="font-size:13px;display:block;margin-top:12px">Nama proyek baru</label><input id="cmName" style="width:100%;font-size:15px;padding:8px;margin:4px 0 8px;box-sizing:border-box" value="'+esc(defName.slice(0, 120))+'">'+
    '<div id="cloudMsg" style="font-size:13px"></div><div class="row"><button class="btn" data-c="merge">Kembali</button><button class="btn primary" data-c="mergeCreate">Buat proyek gabungan</button></div>');
  document.querySelectorAll('[data-mu]').forEach(s => s.onchange = () => { const u = plan[+s.dataset.mu]; u.pick = +s.value; u.touched = true; });
}
function cloudMergeBuild(){
  const {docs, plan} = cloudMergeCtx, units = new Map(), used = new Set();
  for(const u of plan) units.set(u.key, u.vals[u.pick]);
  for(const d of docs){ try{ JSON.parse(d.snap['flowbpa:usedCodes']||'[]').forEach(x => used.add(x)); }catch(e){} }
  // every activity code present in the result stays reserved
  try{ for(const r of JSON.parse(cloudMergeJoin(units, new Set(), '', []) ['flowbpa:data']).rows) used.add(r.l5no); }catch(e){}
  const newest = docs.slice().sort((a, b) => Date.parse(b.updated_at) - Date.parse(a.updated_at))[0];
  const semKey = Object.keys(newest.snap).find(k => k.startsWith('bpa_bpmn_semantics_')) || docs.map(d => Object.keys(d.snap).find(k => k.startsWith('bpa_bpmn_semantics_'))).find(Boolean) || '';
  let order = []; try{ order = [...new Set(JSON.parse(newest.snap['flowbpa:data']).rows.map(r => r.l4no))]; }catch(e){}
  return cloudMergeJoin(units, used, semKey, order);
}
async function cloudMergeCreate(){
  const c = cloudClient(), name = ($('cmName').value||'').trim(); if(!name){ $('cloudMsg').textContent = 'Isi nama proyek baru.'; return; }
  let snap; try{ snap = cloudMergeBuild(); }catch(e){ $('cloudMsg').textContent = 'Gagal menggabungkan: '+e.message; return; }
  const n = cloudRowCount(snap); if(!n){ $('cloudMsg').textContent = 'Hasil gabungan kosong; tidak dibuat.'; return; }
  $('cloudMsg').textContent = 'Menyimpan…';
  const src = cloudMergeCtx.docs.map(d => d.name+' v'+d.version).join(', ');
  const fileName = (()=>{ try{ return JSON.parse(snap['flowbpa:data']).fileName || null; }catch(e){ return null; } })();
  const { data, error } = await c.from('projects').insert({name, data:snap, version:1, file_name:fileName, created_by:cloudUser.email, updated_by:cloudUser.email}).select('id,name,version').single();
  if(error){ $('cloudMsg').textContent = 'Gagal membuat proyek: '+error.message; return; }
  await c.from('project_versions').insert({project_id:data.id, version:1, data:snap, note:('Hasil gabungan: '+src).slice(0, 500), saved_by:cloudUser.email});
  cloudMergeCtx = null;
  cloudShow('<h3>Proyek gabungan dibuat</h3><p style="font-size:13px;line-height:1.5;color:#334155"><b>'+esc(name)+'</b> berisi '+n+' baris aktivitas, digabung dari: '+esc(src)+'. Proyek asal tidak diubah.</p><div class="row"><button class="btn" data-c="back">Kembali ke daftar</button><button class="btn primary" data-c="open" data-id="'+esc(data.id)+'">Buka proyek gabungan</button></div>');
}

/* ================= v1.12.0: Gabungkan dengan mapping per L3 (tanpa diff) =================
   The user picks, per L3, which project supplies it. Everything attached to an L4 under that L3 comes from
   that project only: AS-IS rows, TO-BE draft, SLA/output/notes of its activities, position/unit maps, lane
   mode, lane roles (added/hidden/order), analysis and change register entries. Data not tied to a process
   (role aliases, decision map, position display, BPMN symbol settings, file name) comes from the main project.
   Used activity codes are the union of all selected projects. */
let cloudMapCtx = null;
function cloudMapL4s(docs){
  // l4no → {l3no, l3, l4}
  const m = new Map();
  for(const d of docs){ try{ for(const r of JSON.parse(d.snap['flowbpa:data']).rows){ const l4 = String(r.l4no||''); if(l4 && !m.has(l4)) m.set(l4, {l3no:String(r.l3no||l4), l3:r.l3||'', l4:r.l4||''}); } }catch(e){} }
  return m;
}
function cloudMapL4OfKey(key, l4s){
  const [part, ...rest] = key.split('|'), k = rest.join('|');
  if(part === 'data' || part === 'elab' || part === 'posmap' || part === 'unitmap' || part === 'lanemodes') return k;
  if(part === 'roles') return rest[0] === 'alias' ? null : rest.slice(1).join('|');
  if(part === 'extra'){ let best = null; for(const l4 of l4s.keys()) if(k.startsWith(l4+'.') && (!best || l4.length > best.length)) best = l4; return best || k.replace(/-G\d+$/,'').split('.').slice(0,-1).join('.'); }
  return undefined; // decided by the record itself (analysis/register) or global
}
function cloudMapOwnerL4(key, val, l4s){
  const [part] = key.split('|'); let l4 = cloudMapL4OfKey(key, l4s);
  if(l4 === undefined && (part === 'analysis' || part === 'changeRegister')){ const r = cloudMergeParse(val); l4 = r && r.l4 ? String(r.l4) : null; }
  return l4 === undefined ? null : l4;
}
function cloudMapRender(){
  const {docs, l3s, map, main} = cloudMapCtx, D = docs;
  const opt = (l3, i) => { const has = l3.has[i]; return '<option value="'+i+'"'+(map[l3.l3no] === i ? ' selected' : '')+(has ? '' : ' disabled')+'>'+esc(D[i].name)+(has ? ' ('+(l3.chg[i] ? 'diubah' : 'sama dengan awal')+')' : ' (tidak ada L3 ini)')+'</option>'; };
  cloudShow('<h3>Mapping per L3</h3>'+
    '<p style="font-size:13px;line-height:1.5;color:#334155">Pilih dokumen sumber untuk tiap L3. Seluruh isi L4 di bawahnya (AS-IS, rancangan TO-BE, SLA/output/catatan, pemetaan jabatan/unit, pengaturan lane, analisis) diambil utuh dari dokumen itu, tanpa membandingkan isi. Keterangan "diubah" hanya informasi: isinya berbeda dari snapshot awal.</p>'+
    '<label style="font-size:13px;display:block">Dokumen utama (default L3 dan data umum: alias peran, pemetaan keputusan, pengaturan tampilan)</label><select id="cmMain" style="width:100%;font-size:13.5px;margin:4px 0 8px">'+D.map((d, i) => '<option value="'+i+'"'+(i === main ? ' selected' : '')+'>'+esc(d.name)+' · '+esc(cloudTime(d.updated_at))+'</option>').join('')+'</select>'+
    '<div style="font-size:12.5px;margin:4px 0 8px">Semua L3 dari: '+D.map((d, i) => '<a href="#" data-c="mapAll" data-i="'+i+'">'+esc(d.name)+'</a>').join(' · ')+'</div>'+
    l3s.map(l3 => '<div style="border-top:1px solid var(--line,#e5e7eb);padding:6px 0;font-size:12.5px"><div><b>'+esc(l3.l3no)+'</b> '+esc(l3.l3)+' <span style="color:#64748b">· '+l3.l4s.length+' L4</span></div><div style="font-size:11.5px;color:#64748b">'+esc(l3.l4s.map(x => x.l4no+' '+x.l4).join('; '))+'</div><select data-ml3="'+esc(l3.l3no)+'" style="width:100%;font-size:13px;margin-top:3px">'+D.map((_, i) => opt(l3, i)).join('')+'</select></div>').join('')+
    '<div id="cloudMsg" style="font-size:13px;margin-top:8px"></div><div class="row"><button class="btn" data-c="merge">Kembali</button><button class="btn primary" data-c="mapCheck">Periksa mapping</button></div>');
  $('cmMain').onchange = () => { const nm = +$('cmMain').value; for(const l3 of l3s) if(!l3.touched){ map[l3.l3no] = l3.has[nm] ? nm : map[l3.l3no]; } cloudMapCtx.main = nm; cloudMapRender(); };
  document.querySelectorAll('[data-ml3]').forEach(s => s.onchange = () => { const l3 = l3s.find(x => x.l3no === s.dataset.ml3); map[l3.l3no] = +s.value; l3.touched = true; });
}
async function cloudMapStart(){
  const ids = [...document.querySelectorAll('.cmSel:checked')].map(x => x.value);
  if(ids.length < 2){ $('cloudMsg').textContent = 'Pilih minimal 2 proyek.'; return; }
  $('cloudMsg').textContent = 'Memuat '+ids.length+' proyek…';
  let docs; try{ docs = await cloudMergeLoad(ids); }catch(e){ $('cloudMsg').textContent = e.message; return; }
  const l4s = cloudMapL4s(docs), anc = docs.filter(d => d.base).sort((a, b) => a.baseAt - b.baseAt)[0] || null;
  const byL3 = new Map();
  for(const [l4no, x] of l4s){ if(!byL3.has(x.l3no)) byL3.set(x.l3no, {l3no:x.l3no, l3:x.l3, l4s:[]}); byL3.get(x.l3no).l4s.push({l4no, l4:x.l4}); }
  const cmp = (a, b) => typeof cmpCode === 'function' ? cmpCode(a, b) : String(a).localeCompare(String(b));
  const l3s = [...byL3.values()].sort((a, b) => cmp(a.l3no, b.l3no));
  for(const l3 of l3s){
    const set = new Set(l3.l4s.map(x => x.l4no));
    l3.has = docs.map(d => [...set].some(l4 => d.units.has('data|'+l4)));
    l3.chg = docs.map(d => { if(!anc) return false; const keys = new Set(); for(const k of d.units.keys()) if(set.has(cloudMapOwnerL4(k, d.units.get(k), l4s))) keys.add(k); for(const k of anc.base.keys()) if(set.has(cloudMapOwnerL4(k, anc.base.get(k), l4s))) keys.add(k); for(const k of keys) if((d.units.get(k) ?? null) !== (anc.base.get(k) ?? null)) return true; return false; });
  }
  const main = docs.map((d, i) => i).sort((a, b) => Date.parse(docs[b].updated_at) - Date.parse(docs[a].updated_at))[0];
  const map = {}; for(const l3 of l3s){ map[l3.l3no] = l3.has[main] ? main : l3.has.indexOf(true); }
  cloudMapCtx = {docs, l4s, l3s, map, main};
  cloudMapRender();
}
function cloudMapBuild(){
  const {docs, l4s, map, main} = cloudMapCtx, units = new Map(), used = new Set();
  const ownerOf = l4 => { const x = l4s.get(l4); return x && map[x.l3no] !== undefined ? map[x.l3no] : main; };
  docs.forEach((d, i) => { for(const [k, v] of d.units){ const l4 = cloudMapOwnerL4(k, v, l4s); const owner = l4 ? ownerOf(l4) : main; if(owner === i) units.set(k, v); } });
  for(const d of docs){ try{ JSON.parse(d.snap['flowbpa:usedCodes']||'[]').forEach(x => used.add(x)); }catch(e){} }
  const semKey = Object.keys(docs[main].snap).find(k => k.startsWith('bpa_bpmn_semantics_')) || docs.map(d => Object.keys(d.snap).find(k => k.startsWith('bpa_bpmn_semantics_'))).find(Boolean) || '';
  let order = []; try{ order = [...new Set(JSON.parse(docs[main].snap['flowbpa:data']).rows.map(r => r.l4no))]; }catch(e){}
  const snap = cloudMergeJoin(units, used, semKey, order);
  try{ for(const r of JSON.parse(snap['flowbpa:data']).rows) used.add(r.l5no); }catch(e){}
  try{ for(const items of Object.values(JSON.parse(snap['flowbpa:elab']||'{}'))) for(const i of items||[]) if(i.type === 'graph') try{ JSON.parse(i.text).forEach(r => !/-G\d+$/.test(r.l5no) && used.add(r.l5no)); }catch(e){} }catch(e){}
  snap['flowbpa:usedCodes'] = JSON.stringify([...used].sort());
  return snap;
}
// references (Proses Sebelum/Sesudah, Data dari) that point to activities missing from the result
function cloudMapDangling(snap){
  const RE = /\b\d+(?:\.\d+){2,}(?:-G\d+)?\b/g, asis = new Set(), tobe = new Set(), out = [];
  let rows = [], elab = {}; try{ rows = JSON.parse(snap['flowbpa:data']).rows; }catch(e){} try{ elab = JSON.parse(snap['flowbpa:elab']||'{}'); }catch(e){}
  rows.forEach(r => asis.add(r.l5no));
  const drafts = [];
  for(const [l4, items] of Object.entries(elab)){ const g = (items||[]).filter(i => i.type === 'graph' && i.status !== 'rejected').slice(-1)[0]; if(!g) continue; try{ const dr = JSON.parse(g.text); dr.forEach(r => tobe.add(r.l5no)); drafts.push([l4, dr]); }catch(e){} }
  const check = (r, where, known) => { for(const f of ['prev','next','dataFrom']) for(const c of String(r[f]||'').match(RE)||[]) if(!known.has(c)) out.push({where, from:r.l5no, ref:c}); };
  rows.forEach(r => check(r, 'AS-IS', asis));
  const all = new Set([...asis, ...tobe]); drafts.forEach(([l4, dr]) => dr.forEach(r => check(r, 'TO-BE', all)));
  const seen = new Set(); return out.filter(x => { const k = x.where+x.from+x.ref; if(seen.has(k)) return false; seen.add(k); return true; });
}
function cloudMapCheck(){
  const {docs, l3s, map, main} = cloudMapCtx;
  let snap; try{ snap = cloudMapBuild(); }catch(e){ $('cloudMsg').textContent = 'Gagal menyusun: '+e.message; return; }
  const dang = cloudMapDangling(snap), n = cloudRowCount(snap);
  cloudMapCtx.snap = snap;
  const per = {}; for(const l3 of l3s){ const nm = docs[map[l3.l3no]] ? docs[map[l3.l3no]].name : '(tidak ada)'; (per[nm] = per[nm] || []).push(l3.l3no); }
  cloudShow('<h3>Ringkasan mapping</h3>'+
    '<p style="font-size:13px;line-height:1.5;color:#334155">Hasil: '+n+' baris aktivitas dari '+l3s.length+' L3. Data umum dari <b>'+esc(docs[main].name)+'</b>.</p>'+
    Object.entries(per).map(([nm, l]) => '<div style="font-size:12.5px;border-top:1px solid var(--line,#e5e7eb);padding:5px 0"><b>'+esc(nm)+'</b>: '+esc(l.join(', '))+'</div>').join('')+
    (dang.length ? '<div style="margin-top:10px;font-size:12.5px;color:#a16207"><b>'+dang.length+' rujukan ke aktivitas yang tidak ada di hasil</b> (biasanya karena proses terkait diambil dari dokumen lain). Bisa dibuat tetap, lalu dirapikan di aplikasi:<div style="max-height:160px;overflow:auto;margin-top:4px">'+dang.slice(0, 200).map(x => '<div>'+esc(x.where)+' '+esc(x.from)+' → '+esc(x.ref)+'</div>').join('')+'</div></div>' : '<p style="font-size:12.5px;color:#0a7d32;margin-top:10px">Semua rujukan Proses Sebelum/Sesudah tersambung.</p>')+
    '<label for="cmName" style="font-size:13px;display:block;margin-top:12px">Nama proyek baru</label><input id="cmName" style="width:100%;font-size:15px;padding:8px;margin:4px 0 8px;box-sizing:border-box" value="'+esc(('Mapping — '+docs.map(d => d.name).join(' + ')).slice(0, 120))+'">'+
    '<div id="cloudMsg" style="font-size:13px"></div><div class="row"><button class="btn" data-c="mapBack">Ubah mapping</button><button class="btn primary" data-c="mapCreate">Buat proyek baru</button></div>');
}
async function cloudMapCreate(){
  const c = cloudClient(), name = ($('cmName').value||'').trim(); if(!name){ $('cloudMsg').textContent = 'Isi nama proyek baru.'; return; }
  const {docs, l3s, map, main, snap} = cloudMapCtx, n = cloudRowCount(snap);
  if(!n){ $('cloudMsg').textContent = 'Hasil kosong; tidak dibuat.'; return; }
  $('cloudMsg').textContent = 'Menyimpan…';
  const groups = {}; for(const l3 of l3s){ const i = map[l3.l3no]; if(i !== main) (groups[i] = groups[i] || []).push(l3.l3no); }
  const note = ('Mapping L3: '+Object.entries(groups).map(([i, l]) => l.join(', ')+' ← '+docs[i].name+' v'+docs[i].version).join('; ')+(Object.keys(groups).length ? '; ' : '')+'lainnya ← '+docs[main].name+' v'+docs[main].version).slice(0, 1000);
  const fileName = (()=>{ try{ return JSON.parse(snap['flowbpa:data']).fileName || null; }catch(e){ return null; } })();
  const { data, error } = await c.from('projects').insert({name, data:snap, version:1, file_name:fileName, created_by:cloudUser.email, updated_by:cloudUser.email}).select('id,name,version').single();
  if(error){ $('cloudMsg').textContent = 'Gagal membuat proyek: '+error.message; return; }
  await c.from('project_versions').insert({project_id:data.id, version:1, data:snap, note, saved_by:cloudUser.email});
  cloudMapCtx = null;
  cloudShow('<h3>Proyek baru dibuat</h3><p style="font-size:13px;line-height:1.5;color:#334155"><b>'+esc(name)+'</b> berisi '+n+' baris aktivitas. '+esc(note)+'. Dokumen asal tidak diubah.</p><div class="row"><button class="btn" data-c="back">Kembali ke daftar</button><button class="btn primary" data-c="open" data-id="'+esc(data.id)+'">Buka proyek baru</button></div>');
}
