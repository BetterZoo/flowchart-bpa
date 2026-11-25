
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
    '<div style="display:flex;gap:6px;flex:none">'+(cloudProj&&cloudProj.id===p.id?'':'<button class="btn" data-c="open" data-id="'+esc(p.id)+'">Buka</button>')+'<button class="btn" data-c="rename" data-id="'+esc(p.id)+'" data-name="'+esc(p.name)+'">Ganti nama</button></div></div>').join('')+'</div>';
  h += '<div class="row"><button class="btn" data-c="close">Tutup</button></div>';
  cloudShow(h);
  const ck = $('cloudAutoCk'); if(ck) ck.onchange = () => { store.set(CLOUD_AUTO_KEY, ck.checked?'1':'0'); toast(ck.checked?'Simpan otomatis aktif.':'Simpan otomatis dimatikan. Pakai "Simpan sekarang".'); };
}
async function cloudHistory(){
  const c = cloudClient(); if(!cloudProj) return;
  const { data, error } = await c.from('project_versions').select('id,version,note,saved_by,saved_at').eq('project_id', cloudProj.id).order('id', {ascending:false}).limit(30);
  let h = '<h3>Riwayat versi · '+esc(cloudProj.name)+'</h3><p style="font-size:13px;color:#334155">Disimpan saat Simpan manual, saat proyek dibuat, dan otomatis paling sering tiap 15 menit. 30 versi terakhir dipertahankan.</p>';
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
