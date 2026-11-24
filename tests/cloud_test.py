import sys
from playwright.sync_api import sync_playwright
H='/home/claude/flowchart-bpa/dist/Flowchart_BPA_v1.9.2_BPMN.html'
MOCK=open('/home/claude/flowchart-bpa/tests/sb_mock.js').read()
WB='/home/claude/wb/pelni.xlsx'
def page(b, port_ctx=None):
    ctx=b.new_context(viewport={'width':1500,'height':1000}); pg=ctx.new_page(); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('dialog',lambda d: d.accept(d.default_value or 'PELNI baru') if d.type=='prompt' else d.accept())
    pg.add_init_script(MOCK); return ctx,pg,errs
with sync_playwright() as p:
    b=p.chromium.launch()
    import http.server, threading, functools, os
    os.chdir('/home/claude/flowchart-bpa'); srv=http.server.ThreadingHTTPServer(('127.0.0.1',8765),functools.partial(http.server.SimpleHTTPRequestHandler,directory='/home/claude/flowchart-bpa')); threading.Thread(target=srv.serve_forever,daemon=True).start()
    URL='http://127.0.0.1:8765/dist/Flowchart_BPA_v1.9.2_BPMN.html'
    ctx,pg,errs=page(b); pg.goto(URL); pg.wait_for_timeout(1500)
    print('button:', pg.inner_text('#btnCloud'))
    pg.click('#btnCloud'); pg.wait_for_timeout(500); print('panel1:', pg.inner_text('#cloudDlg h3'))
    pg.fill('#cloudEmail','dea@better.co'); pg.click('[data-c=login]'); pg.wait_for_timeout(300); print('otp:', pg.evaluate('window.__otp'), '|', pg.inner_text('#cloudMsg')[:60])
    # simulate the magic-link session
    pg.evaluate("localStorage.setItem('__mockuser','dea@better.co')"); pg.reload(); pg.wait_for_timeout(1500)
    pg.set_input_files('#fileIn',WB); pg.wait_for_timeout(4500)
    pg.click('#btnCloud'); pg.wait_for_timeout(600); print('panel2:', pg.inner_text('#cloudDlg .dlg')[:120].replace('\n',' | '))
    pg.fill('#cloudNewName','PELNI SOP HC'); pg.click('[data-c=new]'); pg.wait_for_timeout(800)
    print('state:', pg.inner_text('#btnCloud'), '| db:', pg.evaluate("JSON.parse(localStorage.__mockdb).projects.map(p=>[p.name,p.version,Object.keys(p.data).length])"))
    pg.click('[data-c=close]')
    # make an edit (TO-BE lane rename via SLA quick edit) and wait for autosave
    pg.evaluate("state.extra['12.01.02']={sla:'2 HK'}; saveExtra();"); pg.wait_for_timeout(9500)
    print('autosave ->', pg.evaluate("JSON.parse(localStorage.__mockdb).projects[0].version"), pg.inner_text('#btnCloud'))
    # second browser (other teammate) opens the project from cloud
    ctx2,pg2,errs2=page(b); pg2.goto(URL); pg2.wait_for_timeout(800)
    pg2.evaluate("localStorage.setItem('__mockuser','dea@better.co'); localStorage.setItem('__mockdb', %s)" % repr(pg.evaluate("localStorage.__mockdb")))
    pg2.reload(); pg2.wait_for_timeout(1500); pg2.click('#btnCloud'); pg2.wait_for_timeout(600)
    pg2.click('[data-c=open]'); pg2.wait_for_timeout(5000)
    print('browser2 loaded:', pg2.evaluate("state.l4s.size"), pg2.evaluate("(state.extra['12.01.02']||{}).sla"), '|', pg2.inner_text('#btnCloud'))
    # browser2 saves a change -> browser1 conflict
    pg2.evaluate("state.extra['12.01.03']={sla:'5 HK'}; saveExtra();"); pg2.evaluate("cloudSave(true)"); pg2.wait_for_timeout(500)
    db2=pg2.evaluate("localStorage.__mockdb"); pg.evaluate("db=>localStorage.setItem('__mockdb',db)", db2)
    pg.evaluate("state.extra['12.01.04']={sla:'1 HK'}; saveExtra();"); pg.wait_for_timeout(9500)
    print('conflict dlg:', pg.evaluate("document.querySelector('#cloudDlg:not(.hidden) h3')?.textContent"), '|', pg.inner_text('#btnCloud'))
    pg.click('[data-c=overwrite]'); pg.wait_for_timeout(800)
    print('after overwrite v:', pg.evaluate("JSON.parse(localStorage.__mockdb).projects[0].version"), pg.inner_text('#btnCloud'))
    # history + restore
    pg.click('#btnCloud'); pg.wait_for_timeout(600); pg.click('[data-c=history]'); pg.wait_for_timeout(500)
    hist=pg.evaluate("[...document.querySelectorAll('#cloudDlg [data-c=restore]')].length"); print('history items:', hist)
    pg.click('#cloudDlg [data-c=restore] >> nth=-1'); pg.wait_for_timeout(5000)
    print('restored first version, 12.01.02 sla:', pg.evaluate("(state.extra['12.01.02']||{}).sla"), '| btn', pg.inner_text('#btnCloud'))
    # non-member
    ctx3,pg3,errs3=page(b); pg3.goto(URL); pg3.wait_for_timeout(500); pg3.evaluate("localStorage.setItem('__mockuser','orang@lain.com')"); pg3.reload(); pg3.wait_for_timeout(1200)
    pg3.click('#btnCloud'); pg3.wait_for_timeout(600); print('non-member:', pg3.inner_text('#cloudDlg h3'))
    print('errors', errs, errs2, errs3); srv.shutdown(); b.close()
