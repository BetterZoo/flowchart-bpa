from playwright.sync_api import sync_playwright
URL='file:///home/claude/out/Flowchart_BPA_v1.7.0_BPMN.html'
lanes="state.diagram.lanes.filter(l=>l.kind==='act').map(l=>l.key.replace(/^pos::/,''))"
with sync_playwright() as p:
    b=p.chromium.launch(); ctx=b.new_context(viewport={'width':1600,'height':1100}); pg=ctx.new_page(); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('dialog',lambda d: d.accept())
    pg.goto(URL); pg.wait_for_timeout(1200)
    pg.set_input_files('#fileIn','/home/claude/wb/pelni.xlsx'); pg.wait_for_timeout(4000)
    pg.evaluate("select('12.01')"); pg.wait_for_timeout(500); pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(500)
    print('awal TO-BE:', pg.evaluate(lanes))
    # dblclick header of 2nd lane
    d=pg.evaluate("(()=>{const d=state.diagram,ln=d.lanes.filter(l=>l.kind==='act')[1];const r=document.querySelector('#canvas svg').getBoundingClientRect();return [r.left+(ln.x+d.laneW/2)*state.zoom, r.top+((d.top+d.bodyTop)/2)*state.zoom]})()")
    pg.mouse.dblclick(d[0],d[1]); pg.wait_for_timeout(600)
    print('input muncul:', bool(pg.query_selector('#laneNewName')), pg.input_value('#laneNewName') if pg.query_selector('#laneNewName') else '')
    pg.fill('#laneNewName','Supervisor Layanan SDM'); pg.press('#laneNewName','Enter'); pg.wait_for_timeout(900)
    print('TO-BE sesudah:', pg.evaluate(lanes))
    print('svg berisi nama baru:', 'Supervisor Layanan SDM' in pg.evaluate("document.querySelector('#canvas svg').textContent"))
    print('order:', pg.evaluate("JSON.stringify(state.roleEdit.order['12.01'])"))
    pg.select_option('#processEditMode','asis'); pg.wait_for_timeout(600)
    print('AS-IS tetap:', pg.evaluate(lanes))
    # rename via button in AS-IS -> blocked because draft exists
    pg.click('#workspaceOrg'); pg.wait_for_timeout(400)
    pg.click('.rcard >> nth=0 >> [data-a=laneRename]'); pg.wait_for_timeout(300)
    pg.fill('#laneNewName','Staff Personalia'); pg.click('[data-a=laneRenameGo]'); pg.wait_for_timeout(800)
    print('AS-IS (ada draft TO-BE):', pg.evaluate(lanes), '| toast:', pg.evaluate("document.querySelector('.toast,#toast')?.textContent||''")[:90])
    # undo in TO-BE
    pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(500)
    pg.reload(); pg.wait_for_timeout(2500); pg.evaluate("select('12.01')"); pg.wait_for_timeout(500); pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(600)
    print('setelah reload TO-BE:', pg.evaluate(lanes))
    # AS-IS rename on another L4 without draft
    pg.evaluate("select('01.01')"); pg.wait_for_timeout(500); pg.select_option('#processEditMode','asis'); pg.wait_for_timeout(500)
    print('01.01 AS-IS awal:', pg.evaluate(lanes))
    pg.click('#workspaceOrg'); pg.wait_for_timeout(400)
    pg.click('.rcard >> nth=0 >> [data-a=laneRename]'); pg.wait_for_timeout(300)
    pg.fill('#laneNewName','Analis Organisasi'); pg.click('[data-a=laneRenameGo]'); pg.wait_for_timeout(1000)
    print('01.01 AS-IS sesudah:', pg.evaluate(lanes))
    print('rows posX:', pg.evaluate("state.rows.filter(r=>r.l4no==='01.01').map(r=>r.posX).join(' | ')")[:200])
    print('errors', errs[:5]); b.close()
