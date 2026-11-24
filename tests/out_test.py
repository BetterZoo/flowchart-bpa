from playwright.sync_api import sync_playwright
URL='file:///home/claude/out/Flowchart_BPA_v1.7.0_BPMN.html'
def kinds(pg, code):
    return pg.evaluate(f"(()=>{{const n=state.diagram.nodes.find(n=>n.type==='act'&&n.code==='{code}');return n&&n.tag?n.tag.kind+':'+n.tag.text:'-'}})()")
def legend(pg):
    return pg.evaluate("[...document.querySelectorAll('#canvas svg text')].map(t=>t.textContent).filter(t=>/Dokumen|Data di sistem|Data Object|Data Store|Arsip/.test(t)&&t.length<60).slice(-4)")
with sync_playwright() as p:
    b=p.chromium.launch(); ctx=b.new_context(viewport={'width':1600,'height':1100},accept_downloads=True); pg=ctx.new_page(); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('dialog',lambda d: d.accept())
    pg.add_init_script("delete window.showSaveFilePicker; delete window.showDirectoryPicker;")
    pg.goto(URL); pg.wait_for_timeout(1200)
    pg.set_input_files('#fileIn','/home/claude/wb/pelni.xlsx'); pg.wait_for_timeout(4000)
    pg.evaluate("select('12.01')"); pg.wait_for_timeout(400); pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(500)
    print('awal 12.01.07:', kinds(pg,'12.01.07'))
    # TO-BE: set Dokumen
    pg.click('#canvas [data-code="12.01.07"] rect', force=True); pg.wait_for_timeout(400)
    pg.select_option('#directOutKind','object'); pg.fill('#directOutName','Arsip Dokumen Kepegawaian'); pg.click('#directSave'); pg.wait_for_timeout(800)
    print('TO-BE dokumen :', kinds(pg,'12.01.07'), '| legenda:', legend(pg))
    # switch to Data
    pg.click('#canvas [data-code="12.01.07"] rect', force=True); pg.wait_for_timeout(400)
    print('form terisi   :', pg.eval_on_selector('#directOutKind','e=>e.value'), pg.eval_on_selector('#directOutName','e=>e.value'))
    pg.select_option('#directOutKind','store'); pg.fill('#directOutName','Data Induk Pegawai (SAP HR)'); pg.click('#directSave'); pg.wait_for_timeout(800)
    print('TO-BE data    :', kinds(pg,'12.01.07'))
    pg.select_option('#processEditMode','asis'); pg.wait_for_timeout(500)
    print('AS-IS tetap   :', kinds(pg,'12.01.07'))
    # AS-IS: set with empty name -> uses Output
    pg.click('#canvas [data-code="12.01.10"] rect', force=True); pg.wait_for_timeout(400)
    pg.select_option('#directOutKind','object'); pg.fill('#directOutName',''); pg.click('#directSave'); pg.wait_for_timeout(1000)
    pg.evaluate("select('12.01')"); pg.wait_for_timeout(500)
    print('AS-IS 12.01.10 (nama kosong -> Output):', kinds(pg,'12.01.10'))
    # remove symbol
    pg.click('#canvas [data-code="12.01.10"] rect', force=True); pg.wait_for_timeout(400)
    pg.eval_on_selector('#directOutKind',"e=>{e.value='';}"); print('  pilih:', repr(pg.eval_on_selector('#directOutKind','e=>e.value'))); pg.click('#directSave'); pg.wait_for_timeout(1000); pg.evaluate("select('12.01')"); pg.wait_for_timeout(500)
    print('dbg', pg.evaluate("(()=>{const g=state.l4s.get('12.01');const a=g.acts.find(x=>x.l5no==='12.01.10');const d=buildDiagram(g,'',{items:[]});const n=d.nodes.find(n=>n.type==='act'&&n.code==='12.01.10');return JSON.stringify([a.docX,a.systemX,bpmnDataKind(a),n.tag&&n.tag.kind, document.getElementById('processEditMode').value])})()")); print('AS-IS hapus simbol:', kinds(pg,'12.01.10'), pg.evaluate("JSON.stringify([state.byCode.get('12.01.10').docX, state.byCode.get('12.01.10').systemX])"))
    # legend customization
    pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(400)
    pg.evaluate("const e=document.getElementById('legendData'); e.value='Data (tersimpan di aplikasi)'; e.dispatchEvent(new Event('input'))"); pg.wait_for_timeout(800)
    print('legenda custom:', pg.evaluate("[...document.querySelectorAll('#canvas svg text')].some(t=>t.textContent==='Data (tersimpan di aplikasi)')"))
    # master roundtrip
    pg.click('#btnExport'); pg.wait_for_timeout(400); pg.click('label:has(input[name=fmt][value=xlsx])'); pg.wait_for_timeout(300)
    pg.click('#exportDlg label:has-text("Simpan Master Workbook")'); pg.wait_for_timeout(300)
    with pg.expect_download(timeout=60000) as d: pg.click('#exGo')
    d.value.save_as('/home/claude/master_out.xlsx')
    p2=b.new_context(viewport={'width':1600,'height':1100}).new_page(); p2.on('dialog',lambda d: d.accept()); p2.on('pageerror',lambda e: errs.append(str(e)))
    p2.goto(URL); p2.wait_for_timeout(1200); p2.set_input_files('#fileIn','/home/claude/master_out.xlsx'); p2.wait_for_timeout(5000)
    p2.evaluate("select('12.01')"); p2.wait_for_timeout(400); p2.select_option('#processEditMode','tobe'); p2.wait_for_timeout(600)
    print('pulihkan TO-BE:', kinds(p2,'12.01.07'))
    print('errors',errs); b.close()
