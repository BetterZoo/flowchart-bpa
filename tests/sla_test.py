from playwright.sync_api import sync_playwright
W='/home/claude/wb2/BPA_Master_Workbook_PELNI_RR_SOP_Revisi_2_3_perbaikan.xlsx'
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1600,'height':1100}); pg.on('dialog',lambda d:d.accept())
    pg.goto('file:///home/claude/out/Flowchart_BPA_v1.7.0_BPMN.html'); pg.wait_for_timeout(1000)
    pg.set_input_files('#fileIn',W); pg.wait_for_timeout(3000)
    pg.select_option('select[aria-label="Sheet SOP"]','SOP (3)'); pg.click('button:has-text("Import sheet")'); pg.wait_for_timeout(5000)
    pg.evaluate("""for(const c of ['01.01.01','01.01.02','01.01.07','01.01.08','12.01.01','12.01.02','12.01.03','01.03.06']){ state.extra[c]=Object.assign(state.extra[c]||{}, {sla:'3 hari kerja sejak dokumen lengkap diterima'}); } saveExtra&&saveExtra(); select('01.01');""")
    pg.wait_for_timeout(800)
    svg=pg.evaluate("document.querySelector('#canvas svg').outerHTML")
    for c in ['01.01','12.01','01.03']:
        pg.evaluate(f"select('{c}')"); pg.wait_for_timeout(700)
        svg=pg.evaluate("document.querySelector('#canvas svg').outerHTML")
        p2=b.new_page(viewport={'width':1600,'height':900}); p2.set_content('<body style="margin:0">'+svg.replace('<svg','<svg style="width:1600px;height:auto"',1)+'</body>')
        p2.screenshot(path=f'/home/claude/sla_{c}.png',full_page=True); p2.close()
    print(pg.evaluate("Object.keys(state.extra).filter(k=>state.extra[k].sla).length"))
    b.close()
