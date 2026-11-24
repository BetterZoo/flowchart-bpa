import sys
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    b=p.chromium.launch(); ctx=b.new_context(accept_downloads=True,viewport={'width':1600,'height':1100}); pg=ctx.new_page(); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('console',lambda m: errs.append(m.text) if m.type=='error' else None)
    pg.add_init_script("delete window.showSaveFilePicker; delete window.showDirectoryPicker;")
    pg.goto('file:///home/claude/out/Flowchart_BPA_v1.8.0_BPMN.html'); pg.wait_for_timeout(1200)
    pg.set_input_files('#fileIn','/home/claude/wb/pelni.xlsx'); pg.wait_for_timeout(4000)
    pg.evaluate("select('12.01')"); pg.wait_for_timeout(800)
    for fmt in sys.argv[1:]:
        pg.click('#btnExport'); pg.wait_for_timeout(400)
        pg.click(f'label:has(input[name=fmt][value={fmt}])')
        try:
            with pg.expect_download(timeout=60000) as d: pg.click('#exGo')
            path=f'/home/claude/exp_test.{fmt}'; d.value.save_as(path); print(fmt,'saved',d.value.suggested_filename)
        except Exception as e:
            print(fmt,'no download',str(e)[:200]); print(pg.inner_text('#prog')[:300], pg.inner_text('#fallback')[:300])
        pg.wait_for_timeout(500)
        if pg.is_visible('#exCancel'): pg.click('#exCancel')
    print('errors',errs); b.close()
