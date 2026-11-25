// In-memory Supabase stand-in for tests (persists in localStorage '__mockdb' so page reloads keep data).
(function(){
  const load=()=>{try{return JSON.parse(localStorage.getItem('__mockdb'))||{members:[{email:'dea@better.co'},{email:'dea@flowbpa.local'}],projects:[],project_versions:[],seq:1};}catch(e){return {members:[],projects:[],project_versions:[],seq:1};}};
  const save=db=>localStorage.setItem('__mockdb',JSON.stringify(db));
  const user=()=>{const e=localStorage.getItem('__mockuser');return e?{email:e,id:'u-'+e}:null;};
  function Q(table){ this.t=table; this.f=[]; this.op='select'; this.cols='*'; this.ret=false; this.one=false; this.ord=null; this.lim=null; }
  Q.prototype.select=function(c){ if(this.op==='select') this.cols=c||'*'; else this.ret=true; return this; };
  Q.prototype.eq=function(k,v){ this.f.push([k,v]); return this; };
  Q.prototype.order=function(k,o){ this.ord=[k,o&&o.ascending===false?-1:1]; return this; };
  Q.prototype.limit=function(n){ this.lim=n; return this; };
  Q.prototype.single=function(){ this.one=true; return this; };
  Q.prototype.insert=function(o){ this.op='insert'; this.val=o; return this; };
  Q.prototype.update=function(o){ this.op='update'; this.val=o; return this; };
  Q.prototype.run=function(){
    const db=load(), u=user(); if(!u) return {data:null,error:{message:'not authenticated'}};
    const member=db.members.some(m=>m.email===u.email); let rows=db[this.t]||[];
    if(this.t!=='members' && !member) return {data:this.op==='select'?[]:null,error:this.op==='select'?null:{message:'RLS'}};
    if(this.t==='members') rows=rows.filter(m=>m.email===u.email);
    const match=r=>this.f.every(([k,v])=>r[k]===v);
    let out;
    if(this.op==='insert'){ const r=Object.assign({},this.val); if(this.t==='projects'){ r.id='p'+(db.seq++); r.updated_at=new Date().toISOString(); } else r.id=db.seq++; if(!r.saved_at) r.saved_at=new Date().toISOString(); db[this.t].push(r); save(db); out=[r]; window.__mockWrites=(window.__mockWrites||0)+1; }
    else if(this.op==='update'){ out=[]; for(const r of db[this.t]) if(match(r)){ const old=JSON.parse(JSON.stringify(r)); Object.assign(r,this.val); out.push(r); if(this.t==='projects' && this.val.data){ /* v1.9.8: mimic trg_log_project_save */ db.project_saves=db.project_saves||[]; const ks=[...new Set([...Object.keys(old.data||{}),...Object.keys(r.data||{})])].filter(k=>JSON.stringify((old.data||{})[k])!==JSON.stringify(r.data[k])).map(k=>k.replace('flowbpa:','')).sort(); db.project_saves.push({id:db.seq++,project_id:r.id,version:r.version,saved_by:r.updated_by,saved_at:r.updated_at,changed:ks}); if(old.updated_by!==r.updated_by && !db.project_versions.some(v=>v.project_id===r.id&&v.version===old.version)) db.project_versions.push({id:db.seq++,project_id:r.id,version:old.version,data:old.data,note:'Versi terakhir sebelum diedit '+String(r.updated_by).split('@')[0],saved_by:old.updated_by,saved_at:old.updated_at}); } } save(db); window.__mockWrites=(window.__mockWrites||0)+1; }
    else { out=rows.filter(match); if(this.ord) out=out.slice().sort((a,b)=>(a[this.ord[0]]>b[this.ord[0]]?1:-1)*this.ord[1]); if(this.lim) out=out.slice(0,this.lim); }
    out=JSON.parse(JSON.stringify(out));
    if(this.one) return out.length?{data:out[0],error:null}:{data:null,error:{message:'no rows'}};
    return {data:out,error:null};
  };
  Q.prototype.then=function(res,rej){ return Promise.resolve(this.run()).then(res,rej); };
  window.__sbMock={createClient(){ return {
    auth:{ async getSession(){ const u=user(); return {data:{session:u?{user:u}:null}}; },
      async signInWithOtp({email}){ window.__otp=email; return {error:null}; },
      // v1.9.7: password login; test accounts live in localStorage '__mockpw' as {email: password}, default dea@flowbpa.local / rahasia123
      async signInWithPassword({email,password}){ const pw=JSON.parse(localStorage.getItem('__mockpw')||'{"dea@flowbpa.local":"rahasia123"}'); if(pw[email]!==password) return {data:{user:null,session:null},error:{message:'Invalid login credentials'}}; localStorage.setItem('__mockuser',email); const u={email,id:'u-'+email}; return {data:{user:u,session:{user:u}},error:null}; },
      async updateUser({password}){ const u=user(); if(!u) return {data:null,error:{message:'not authenticated'}}; const pw=JSON.parse(localStorage.getItem('__mockpw')||'{"dea@flowbpa.local":"rahasia123"}'); pw[u.email]=password; localStorage.setItem('__mockpw',JSON.stringify(pw)); return {data:{user:u},error:null}; },
      async signOut(){ localStorage.removeItem('__mockuser'); return {error:null}; },
      onAuthStateChange(){ return {data:{subscription:{unsubscribe(){}}}}; } },
    from(t){ return new Q(t); } }; }};
})();
