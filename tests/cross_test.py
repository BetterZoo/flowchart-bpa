from playwright.sync_api import sync_playwright
URL='file:///home/claude/out/Flowchart_BPA_v1.7.0_BPMN.html'
CI="(c)=>{const n=state.diagram.nodes.find(n=>n.type==='act'&&n.code===c);return n&&n.crossInputs&&n.crossInputs.length?n.crossInputs.map(x=>x.source+':'+x.kind+':'+x.label).join(' | '):'-'}"
OUT="(c)=>{const n=state.diagram.nodes.find(n=>n.type==='act'&&n.code===c);return n&&n.tag?n.tag.kind+':'+n.tag.text:'-'}"
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1600,'height':1100}); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('dialog',lambda d: d.accept())
    pg.goto(URL); pg.wait_for_timeout(1200); pg.set_input_files('#fileIn','/home/claude/wb/pelni.xlsx'); pg.wait_for_timeout(4000)
    pg.evaluate("select('12.03')"); pg.wait_for_timeout(400); pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(400)
    print('12.03.01 input awal (TO-BE):', pg.evaluate(CI,'12.03.01'))
    pg.evaluate("select('12.01')"); pg.wait_for_timeout(400)
    print('12.01.10 output awal       :', pg.evaluate(OUT,'12.01.10'))
    pg.click('#canvas [data-code="12.01.10"] rect', force=True); pg.wait_for_timeout(400)
    pg.select_option('#directOutKind','store'); pg.fill('#directOutName','Data Induk Pegawai'); pg.click('#directSave'); pg.wait_for_timeout(800)
    print('12.01.10 output baru       :', pg.evaluate(OUT,'12.01.10'))
    pg.evaluate("select('12.03')"); pg.wait_for_timeout(500)
    print('12.03.01 input TO-BE       :', pg.evaluate(CI,'12.03.01'))
    pg.select_option('#processEditMode','asis'); pg.wait_for_timeout(500)
    print('12.03.01 input AS-IS       :', pg.evaluate(CI,'12.03.01'))
    pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(400)
    pg.evaluate("select('12.01')"); pg.wait_for_timeout(400)
    pg.click('#canvas [data-code="12.01.10"] rect', force=True); pg.wait_for_timeout(400)
    pg.eval_on_selector('#directOutKind',"e=>{e.value='';}"); pg.click('#directSave'); pg.wait_for_timeout(800)
    pg.evaluate("select('12.03')"); pg.wait_for_timeout(500)
    print('12.03.01 input stl "tidak ada":', pg.evaluate(CI,'12.03.01'))
    print('errors',errs); b.close()
