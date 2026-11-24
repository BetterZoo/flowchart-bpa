from playwright.sync_api import sync_playwright
URL='file:///home/claude/out/Flowchart_BPA_v1.8.0_BPMN.html'
CHECK=open('/home/claude/check2.py').read().split('JS = r"""')[1].split('"""')[0]
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1600,'height':1100}); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('dialog',lambda d: d.accept())
    pg.goto(URL); pg.wait_for_timeout(1200); pg.set_input_files('#fileIn','/home/claude/wb2/BPA_Master_Workbook_PELNI_RR_SOP_Revisi_2_3_perbaikan.xlsx'); pg.wait_for_timeout(3500)
    pg.select_option('select[aria-label="Sheet SOP"]','SOP (3)'); pg.click('button:has-text("Import sheet")'); pg.wait_for_timeout(5000)
    pg.evaluate("select('12.01')"); pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(600)
    def op_add(code, where_idx, name, lane_idx=None):
        pg.evaluate(f"directEditor('{code}')"); pg.wait_for_timeout(300); pg.click('#structAdd'); pg.wait_for_timeout(300)
        opts=pg.evaluate("[...document.querySelectorAll('#structWhere option')].map(o=>o.value+' = '+o.textContent)")
        pg.select_option('#structWhere', index=where_idx); pg.fill('#structName',name)
        if lane_idx is not None: pg.select_option('#structLane', index=lane_idx)
        pg.click('#structSave'); pg.wait_for_timeout(900)
        print('ADD',code,'opsi:',opts,'| toast:',pg.evaluate("document.querySelector('#toast')?.textContent||''")[:80], '| err:', pg.evaluate("document.querySelector('#structError')?.textContent||''"))
    def op_del(code, gate_val=None):
        pg.evaluate(f"directEditor('{code}')"); pg.wait_for_timeout(300); pg.click('#structDel'); pg.wait_for_timeout(300)
        radios=pg.evaluate("[...document.querySelectorAll('input[name=delGate]')].map(i=>i.value+' = '+i.parentElement.textContent.trim())")
        if gate_val: pg.check(f'input[name=delGate][value="{gate_val}"]')
        pg.click('#structSave'); pg.wait_for_timeout(900)
        print('DEL',code,'pilihan:',radios,'| toast:',pg.evaluate("document.querySelector('#toast')?.textContent||''")[:80], '| err:', pg.evaluate("document.querySelector('#structError')?.textContent||''"))
    op_add('12.01.01',1,'Mencatat permintaan masuk')
    op_add('12.01.04',2,'Meminta kelengkapan dokumen', 1)
    op_add('12.01.10',0,'Mengarsipkan dokumen')
    op_del('12.01.09','keep:1')
    op_del('12.01.06')
    print('lanes:', pg.evaluate("state.diagram.lanes.filter(l=>l.kind==='act').map(l=>l.key)"))
    print('check', pg.evaluate(CHECK,'12.01'))
    svg=pg.evaluate("document.querySelector('#canvas svg').outerHTML")
    p2=b.new_page(viewport={'width':1600,'height':900}); p2.set_content('<body style="margin:0">'+svg.replace('<svg','<svg style="width:1600px;height:auto"',1)+'</body>')
    p2.screenshot(path='/home/claude/ui_12.01.png',full_page=True)
    # undo
    pg.evaluate("undoStructural()"); pg.wait_for_timeout(700)
    print('undo ok, acts:', pg.evaluate("structuralRows(state.l4s.get('12.01'),'tobe').length"))
    print('errors',errs[:5]); b.close()
