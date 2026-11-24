from playwright.sync_api import sync_playwright
URL='file:///home/claude/out/Flowchart_BPA_v1.8.0_BPMN.html'
CHECK=open('/home/claude/check2.py').read().split('JS = r"""')[1].split('"""')[0]
H=r"""
([l4,code,op,arg,mode])=>{ select(l4); $('processEditMode').value=mode; $('processEditMode').dispatchEvent(new Event('change'));
 const g=state.l4s.get(l4); const rows=structuralRows(g,mode); const absorbed=consumedIds(rows);
 try{ const res=structuralChange(rows,code,op,Object.assign({mode},arg)); const px=res.patches;
   if(mode==='asis'){ res.forEach(r=>r._asisEdited=true); applyCrossPatches(px); ingest([...state.rows.filter(r=>r.l4no!==g.l4no),...res],state.fileName); select(l4);}
   else saveUnifiedDraft(g,res,absorbed,'t');
   const now=structuralRows(state.l4s.get(l4),mode);
   return {ok:true,new:px.newCode,out:px.out,warn:px.warn,rows:now.map(r=>r.l5no+' | '+r.prev+' -> '+r.next+(r.decisionQuestion?' ['+r.decisionQuestion+']':''))};
 }catch(e){ return {ok:false,err:e.message}; } }"""
def show(t,r):
    print('==',t, 'OK' if r.get('ok') else 'ERR '+r.get('err',''), r.get('new',''), r.get('out',''), r.get('warn',''))
    for x in r.get('rows',[]): print('   ',x)
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1600,'height':1100}); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('dialog',lambda d: d.accept())
    pg.goto(URL); pg.wait_for_timeout(1200); pg.set_input_files('#fileIn','/home/claude/wb/pelni.xlsx'); pg.wait_for_timeout(4500)
    T=lambda *a: pg.evaluate(H,list(a))
    show('12.01 sebelum aktivitas pertama (ada trigger L4 lain)', T('12.01','12.01.01','add',{'where':'before','name':'Menerima permintaan','role':'x'},'tobe'))
    show('12.01 di cabang gateway 12.01.04 [0]', T('12.01','12.01.04','add',{'where':'branch:0','name':'Melengkapi dokumen','role':'x'},'tobe'))
    show('12.01 sesudah 12.01.04 sebelum gateway', T('12.01','12.01.04','add',{'where':'after','name':'Mencatat hasil verifikasi','role':'x'},'tobe'))
    print('check', pg.evaluate(CHECK,'12.01'))
    show('01.03 sesudah aktivitas terakhir 01.03.06 (gateway ke L4 lain)', T('01.03','01.03.06','add',{'where':'after','name':'Mengarsipkan','role':'x'},'tobe'))
    show('01.03 di cabang ke Selesai', T('01.03','01.03.07' if False else '01.03.06','add',{'where':'branch:0','name':'Menyusun rencana','role':'x'},'tobe'))
    print('check', pg.evaluate(CHECK,'01.03'))
    # deletes on 12.03 TO-BE fresh
    show('12.02 hapus aktivitas pertama', T('12.02','12.02.01','delete',{},'tobe'))
    print('check', pg.evaluate(CHECK,'12.02'))
    r=pg.evaluate("(()=>{const g=state.l4s.get('12.01');return structuralRows(g,'asis').filter(r=>parseFlow(r.next).filter(f=>f.label).length>1).map(r=>r.l5no)})()"); print('gateway acts 12.01', r)
    show('12.01 AS-IS? (draft ada) hapus gateway pindah ke pendahulu', T('12.01','12.01.04','delete',{'gate':'pred'},'tobe'))
    print('check', pg.evaluate(CHECK,'12.01'))
    show('03.04 hapus gateway keep:0', T('03.04','03.04.09','delete',{'gate':'keep:1'},'tobe'))
    show('01.01 AS-IS hapus 01.01.01 (dipicu 01.03.06)', T('01.01','01.01.01','delete',{},'asis'))
    print('01.03.06 next sekarang:', pg.evaluate("state.rows.find(r=>r.l5no==='01.03.06').next"))
    show('01.01 AS-IS tambah sebelum 01.01.02 (jadi pertama)', T('01.01','01.01.02','add',{'where':'before','name':'Menerima hasil evaluasi','role':'x'},'asis'))
    print('01.03.06 next sekarang:', pg.evaluate("state.rows.find(r=>r.l5no==='01.03.06').next"))
    print('check', pg.evaluate(CHECK,'01.01'))
    print('errors',errs[:5]); b.close()
