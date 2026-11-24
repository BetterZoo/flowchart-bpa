import sys
from playwright.sync_api import sync_playwright
url, wb, mode = sys.argv[1], sys.argv[2], sys.argv[3]
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1600,'height':1100}); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('console',lambda m: errs.append(m.text) if m.type=='error' else None)
    pg.goto('file://'+url); pg.wait_for_timeout(1200)
    pg.set_input_files('#fileIn',wb); pg.wait_for_timeout(4000)
    if pg.query_selector('select[aria-label="Sheet SOP"]'):
        pg.select_option('select[aria-label="Sheet SOP"]','SOP (3)'); pg.click('button:has-text("Import sheet")'); pg.wait_for_timeout(5000)

    pg.select_option('#processEditMode',mode); pg.wait_for_timeout(500)
    if sys.argv[4]=='ALL':
        codes=pg.evaluate("[...state.l4s.keys()]"); bad=[]
        for c in codes:
            try:
                pg.evaluate(f"select('{c}')")
                ok=pg.evaluate("!!document.querySelector('#canvas svg')")
                if not ok: bad.append(c)
            except Exception as e: bad.append((c,str(e)[:120]))
        print('L4 count',len(codes),'bad',bad[:10])
        print('warnings', pg.evaluate("(state.issues||state.warnings||[]).length") if True else '')
    else:
        for code in sys.argv[4:]:
            pg.evaluate(f"select('{code}')"); pg.wait_for_timeout(800)
            svg=pg.evaluate("document.querySelector('#canvas svg').outerHTML")
            p2=b.new_page(viewport={'width':1600,'height':900}); p2.set_content('<body style="margin:0">'+svg.replace('<svg','<svg style="width:1600px;height:auto"',1)+'</body>')
            p2.screenshot(path=f'/home/claude/pel_{code}_{sys.argv[5] if False else "x"}.png',full_page=True); p2.close()
    print('errors',errs[:5]); b.close()
