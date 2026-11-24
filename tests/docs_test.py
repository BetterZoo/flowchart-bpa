from playwright.sync_api import sync_playwright
URL='file:///home/claude/out/Flowchart_BPA_v1.8.0_BPMN.html'
CHECK=open('/home/claude/check2.py').read().split('JS = r"""')[1].split('"""')[0]
with sync_playwright() as p:
    b=p.chromium.launch(); ctx=b.new_context(viewport={'width':1600,'height':1100},accept_downloads=True); pg=ctx.new_page(); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('dialog',lambda d: d.accept())
    pg.add_init_script("delete window.showSaveFilePicker; delete window.showDirectoryPicker;")
    pg.goto(URL); pg.wait_for_timeout(1200); pg.set_input_files('#fileIn','/home/claude/wb/pelni.xlsx'); pg.wait_for_timeout(4500)
    pg.evaluate("select('12.01')"); pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(500)
    def add(code, kind_btn, name, k='object'):
        pg.evaluate(f"directEditor('{code}')"); pg.wait_for_timeout(300)
        pg.click(f'#docsSec [data-df={kind_btn}]'); pg.wait_for_timeout(150)
        rows=pg.query_selector_all('#docsSec .docrow'); r=rows[-1]
        r.query_selector('select[data-df=k]').select_option(k)
        r.query_selector('input[data-df=n]').fill(name); pg.wait_for_timeout(500)
    add('12.01.02','addin','Formulir Permintaan Perubahan Data')
    add('12.01.02','addin','Data Pegawai di SAP','store')
    add('12.01.02','addout','Checklist Kelengkapan Dokumen')
    add('12.01.10','addout','Laporan Pemutakhiran Data')   # has main output already
    add('12.01.01','addout','Tanda Terima Permintaan')      # has cross inputs
    print('extra:', pg.evaluate("JSON.stringify({a:state.extra['12.01.02'].docs,b:state.extra['12.01.10'].docs})"))
    print('stored:', pg.evaluate("localStorage.getItem('flowbpa:extra')")[:300])
    print('check', pg.evaluate(CHECK,'12.01'))
    # delete one
    pg.evaluate("directEditor('12.01.02')"); pg.wait_for_timeout(300); pg.click('#docsSec .docrow >> nth=1 >> [data-df=del]'); pg.wait_for_timeout(500)
    print('after del:', pg.evaluate("JSON.stringify(state.extra['12.01.02'].docs)"))
    print('panel:', [l for l in pg.inner_text('#canvas').split('\n') if l.startswith(('Input 12','Output 12'))])
    svg=pg.evaluate("document.querySelector('#canvas svg').outerHTML")
    p2=ctx.new_page(); p2.set_content('<body style="margin:0">'+svg.replace('<svg','<svg style="width:1600px;height:auto"',1)+'</body>'); p2.screenshot(path='/home/claude/docs.png',full_page=True); p2.close()
    # master roundtrip
    pg.click('#btnExport'); pg.wait_for_timeout(400); pg.click('label:has(input[name=fmt][value=xlsx])'); pg.wait_for_timeout(300)
    pg.click('#exportDlg label:has-text("Simpan Master Workbook")'); pg.wait_for_timeout(300)
    with pg.expect_download(timeout=60000) as d: pg.click('#exGo')
    d.value.save_as('/home/claude/master_docs.xlsx')
    p3=b.new_context(viewport={'width':1600,'height':1100}).new_page(); p3.on('dialog',lambda d: d.accept())
    p3.goto(URL); p3.wait_for_timeout(1200); p3.set_input_files('#fileIn','/home/claude/master_docs.xlsx'); p3.wait_for_timeout(5000)
    print('pulih:', p3.evaluate("JSON.stringify([state.extra['12.01.02']?.docs,state.extra['12.01.10']?.docs])"))
    print('errors',errs); b.close()
