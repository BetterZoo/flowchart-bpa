import sys, random
from playwright.sync_api import sync_playwright
CHECK=open('/home/claude/check2.py').read().split('JS = r"""')[1].split('"""')[0]
URL='file:///home/claude/out/Flowchart_BPA_v1.8.0_BPMN.html'
FZ=r"""
([code,seed])=>{
 let s=seed; const rnd=n=>{s=(s*9301+49297)%233280; return Math.floor(s/233280*n);};
 select(code); $('processEditMode').value='tobe'; $('processEditMode').dispatchEvent(new Event('change'));
 const g=state.l4s.get(code); const log=[];
 for(let k=0;k<10;k++){
   let rows; try{ rows=structuralRows(g,'tobe'); }catch(e){ log.push('rows:'+e.message); break; }
   const acts=rows; if(!acts.length) break;
   const a=acts[rnd(acts.length)], b=acts[rnd(acts.length)];
   const ops=['add','add','delete','delete','move','dec','sla','next','ng','ng'];
   const op=ops[rnd(ops.length)];
   try{
     let res, absorbed=consumedIds(rows);
     if(op==='add'){ const nx=parseFlow(a.next); const ws=['after','before',...nx.map((f,i)=>'branch:'+i).filter((w,i)=>nx[i].label)]; res=structuralChange(rows,a.l5no,'add',{mode:'tobe',name:'Aktivitas uji '+k,role:a.role,where:ws[rnd(ws.length)]}); }
     else if(op==='move'){ if(a===b) continue; res=structuralChange(rows,a.l5no,'move',{mode:'tobe',target:b.l5no,position:rnd(2)?'before':'after'}); }
     else if(op==='delete'){ if(acts.length<3) continue; const nx=parseFlow(a.next).filter(f=>f.label); res=structuralChange(rows,a.l5no,'delete',{mode:'tobe',gate:nx.length>1?(rnd(2)?'pred':'keep:'+rnd(nx.length)):''}); }
     else if(op==='dec'){ if(a===b) continue; const ex=parseFlow(a.next).find(f=>f.kind==='code'&&rows.some(r=>r.l5no===f.code)); if(!ex) continue; res=structuralChange(rows,a.l5no,'connector',{mode:'tobe',question:'Uji?',branches:[{label:'Ya',kind:'code',code:ex.code},{label:'Tidak',kind:'code',code:b.l5no}]}); }
     else if(op==='next'){ if(a===b) continue; if(parseFlow(a.next).some(f=>f.kind==='code'&&!rows.some(r=>r.l5no===f.code))) continue; res=structuralChange(rows,a.l5no,'connector',{mode:'tobe',question:'',branches:[{label:'',kind:'code',code:b.l5no}]}); }
     else if(op==='ng'){ const nx=parseFlow(a.next); const labs=nx.map(f=>f.label); if(nx.length<2||labs.some(l=>!l)||new Set(labs).size!==labs.length||a===b) continue; const i=rnd(nx.length); const br=nx.map((f,j)=>j===i?{label:f.label,kind:'newgate',question:'Uji bertingkat '+k+'?',branches:[{label:'Ya',kind:f.kind,code:f.code},{label:'Tidak',kind:'code',code:b.l5no}]}:f); res=structuralChange(rows,a.l5no,'connector',{mode:'tobe',question:a.decisionQuestion||'Q?',branches:br}); }
     else if(op==='sla'){ state.extra[a.l5no]=Object.assign(state.extra[a.l5no]||{},{sla:'1 HK'}); saveExtra(); redraw(); log.push('sla '+a.l5no); continue; }
     const c=structuralConflict(g,a.l5no,['dec','next','ng'].includes(op)?'connector':op,'tobe'); if(c){ log.push('conflict '+op); continue; }
     saveUnifiedDraft(g,res,absorbed,'fuzz '+op); log.push(op+' '+a.l5no+(op==='move'||op==='dec'||op==='next'?'→'+b.l5no:''));
   }catch(e){ log.push('ERR '+op+' '+a.l5no+': '+e.message); }
   if(!document.querySelector('#canvas svg')) { log.push('NO SVG'); break; }
 }
 return log;
}"""
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1600,'height':1100}); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('dialog',lambda d: d.accept())
    pg.goto(URL); pg.wait_for_timeout(1200); pg.set_input_files('#fileIn',sys.argv[1] if len(sys.argv)>1 else '/home/claude/wb/pelni.xlsx'); pg.wait_for_timeout(4500)
    codes=pg.evaluate("[...state.l4s.keys()]"); tot={'dangling':0,'tov':0,'tin':0,'ex':0}; bad=[]
    for i,c in enumerate(codes):
        log=pg.evaluate(FZ,[c,i*7+3])
        r=pg.evaluate(CHECK,c)
        for k in tot: tot[k]+=r.get(k,0)
        if any(r.get(k) for k in tot) or any(x.startswith(('ERR','NO')) for x in log): bad.append((c,log,{k:r.get(k) for k in tot}))
    print('TOTAL',tot); 
    for x in bad:
        if any(x[2].values()): print(x)
    print('pageerrors',errs[:5]); b.close()
