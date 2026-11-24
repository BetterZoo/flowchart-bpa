from playwright.sync_api import sync_playwright
URL='file:///home/claude/out/Flowchart_BPA_v1.8.0_BPMN.html'
CHECK=open('/home/claude/check2.py').read().split('JS = r"""')[1].split('"""')[0]
ROWS="(l4)=>structuralRows(state.l4s.get(l4),$('processEditMode').value).map(r=>r.l5no+' | '+r.prev+' -> '+r.next+(r.decisionQuestion?' ['+r.decisionQuestion+']':''))"
with sync_playwright() as p:
    b=p.chromium.launch(); ctx=b.new_context(viewport={'width':1600,'height':1100},accept_downloads=True); pg=ctx.new_page(); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('dialog',lambda d: d.accept())
    pg.add_init_script("delete window.showSaveFilePicker; delete window.showDirectoryPicker;")
    pg.goto(URL); pg.wait_for_timeout(1200); pg.set_input_files('#fileIn','/home/claude/wb/pelni.xlsx'); pg.wait_for_timeout(4500)
    pg.evaluate("select('01.01')"); pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(600)
    # open gateway panel by double-clicking diamond of 01.01.08
    box=pg.evaluate("(()=>{const el=document.querySelector('.dmd[data-dec=\"01.01.08\"]');const r=el.getBoundingClientRect();return [r.x+r.width/2,r.y+r.height/2]})()")
    pg.mouse.dblclick(box[0],box[1]); pg.wait_for_timeout(500)
    print('panel:', pg.evaluate("document.querySelector('#detail h3')?.textContent"))
    rows=pg.query_selector_all('#branchRows .branchRow'); print('cabang:', [r.query_selector('.branchLabel').input_value() for r in rows])
    # branch "Disetujui" -> new gateway
    r=[r for r in rows if r.query_selector('.branchLabel').input_value()=='Disetujui'][0]
    r.query_selector('.branchTarget').select_option('__NEWGATE__'); pg.wait_for_timeout(200)
    r.query_selector('.newGateQ').fill('Perlu persetujuan Dewan Komisaris?')
    subs=r.query_selector_all('.subBranch'); 
    subs[0].query_selector('.subLabel').fill('Perlu'); subs[0].query_selector('.subTarget').select_option('01.01.06')
    subs[1].query_selector('.subLabel').fill('Tidak perlu'); subs[1].query_selector('.subTarget').select_option('01.02.01')
    pg.click('#structSave'); pg.wait_for_timeout(1000)
    print('err:', pg.evaluate("document.querySelector('#structError')?.textContent||''"), '| toast:', pg.evaluate("document.querySelector('#toast').textContent"))
    for x in pg.evaluate(ROWS,'01.01'): print('  ',x)
    print('dec nodes:', pg.evaluate("state.diagram.nodes.filter(n=>n.type==='dec').map(n=>n.code||n.act.l5no)"))
    print('check', pg.evaluate(CHECK,'01.01'))
    # nested again: double-click G1 diamond, add a gateway inside its branch "Perlu"
    box=pg.evaluate("(()=>{const el=document.querySelector('.dmd[data-dec=\"01.01.08-G1\"]');const r=el.getBoundingClientRect();return [r.x+r.width/2,r.y+r.height/2]})()")
    pg.mouse.dblclick(box[0],box[1]); pg.wait_for_timeout(500)
    print('panel2:', pg.evaluate("document.querySelector('#detail h3')?.textContent"), pg.evaluate("!!document.querySelector('#gateDel')"))
    rows=pg.query_selector_all('#branchRows .branchRow'); r=rows[0]
    r.query_selector('.branchTarget').select_option('__NEWGATE__'); r.query_selector('.newGateQ').fill('Anggaran tersedia?')
    subs=r.query_selector_all('.subBranch'); subs[0].query_selector('.subTarget').select_option('01.01.06'); subs[1].query_selector('.subTarget').select_option('SELESAI')
    pg.click('#structSave'); pg.wait_for_timeout(1000)
    print('err:', pg.evaluate("document.querySelector('#structError')?.textContent||''"))
    for x in pg.evaluate(ROWS,'01.01'): print('  ',x)
    print('check', pg.evaluate(CHECK,'01.01'))
    svg=pg.evaluate("document.querySelector('#canvas svg').outerHTML")
    p2=ctx.new_page(); p2.set_content('<body style="margin:0">'+svg.replace('<svg','<svg style="width:1600px;height:auto"',1)+'</body>'); p2.screenshot(path='/home/claude/gw.png',full_page=True); p2.close()
    # add activity in G2 branch
    pg.evaluate("structuralEditor('01.01.08-G2','add')"); pg.wait_for_timeout(300)
    print('add opts:', pg.evaluate("[...document.querySelectorAll('#structWhere option')].map(o=>o.textContent)"))
    pg.select_option('#structWhere', index=1); pg.fill('#structName','Mengajukan tambahan anggaran'); pg.click('#structSave'); pg.wait_for_timeout(900)
    print('err:', pg.evaluate("document.querySelector('#structError')?.textContent||''"))
    # delete G2 keep branch 0
    pg.evaluate("structuralEditor('01.01.08-G2','delete')"); pg.wait_for_timeout(300)
    print('del opts:', pg.evaluate("[...document.querySelectorAll('input[name=delGate]')].map(i=>i.parentElement.textContent.trim())"))
    pg.click('#structSave'); pg.wait_for_timeout(900)
    print('err:', pg.evaluate("document.querySelector('#structError')?.textContent||''"))
    for x in pg.evaluate(ROWS,'01.01'): print('  ',x)
    print('check', pg.evaluate(CHECK,'01.01'))
    print('panel B count ok:', pg.evaluate("state.diagram.nodes.filter(n=>n.type==='act').length"))
    # master roundtrip
    pg.click('#btnExport'); pg.wait_for_timeout(400); pg.click('label:has(input[name=fmt][value=xlsx])'); pg.wait_for_timeout(300)
    pg.click('#exportDlg label:has-text("Simpan Master Workbook")'); pg.wait_for_timeout(300)
    with pg.expect_download(timeout=60000) as d: pg.click('#exGo')
    d.value.save_as('/home/claude/master_gw.xlsx')
    p3=b.new_context(viewport={'width':1600,'height':1100}).new_page(); p3.on('dialog',lambda d: d.accept()); p3.on('pageerror',lambda e: errs.append('p3 '+str(e)))
    p3.goto(URL); p3.wait_for_timeout(1200); p3.set_input_files('#fileIn','/home/claude/master_gw.xlsx'); p3.wait_for_timeout(5000)
    p3.evaluate("select('01.01')"); p3.select_option('#processEditMode','tobe'); p3.wait_for_timeout(600)
    print('pulih dec:', p3.evaluate("state.diagram.nodes.filter(n=>n.type==='dec').map(n=>n.act.l5no)"))
    print('errors',errs); b.close()
