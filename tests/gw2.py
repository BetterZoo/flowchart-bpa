from playwright.sync_api import sync_playwright
URL='file:///home/claude/out/Flowchart_BPA_v1.8.0_BPMN.html'
CHECK=open('/home/claude/check2.py').read().split('JS = r"""')[1].split('"""')[0]
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1600,'height':1100}); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('dialog',lambda d: d.accept())
    pg.goto(URL); pg.wait_for_timeout(1200); pg.set_input_files('#fileIn','/home/claude/wb/pelni.xlsx'); pg.wait_for_timeout(4500)
    # 1) re-save gateway with parallel branch unchanged (AS-IS)
    pg.evaluate("select('01.01')"); pg.select_option('#processEditMode','asis'); pg.wait_for_timeout(500)
    pg.evaluate("structuralEditor('01.01.08','connector','decision')"); pg.wait_for_timeout(300)
    print('hint:', pg.evaluate("[...document.querySelectorAll('#branchRows .ws-field-hint')].map(x=>x.textContent)"))
    pg.fill('#decisionQuestion','Apakah rancangan disetujui Direksi?'); pg.click('#structSave'); pg.wait_for_timeout(900)
    print('err:', pg.evaluate("document.querySelector('#structError')?.textContent||''"), '|', pg.evaluate("state.rows.find(r=>r.l5no==='01.01.08').next"))
    # 2) AS-IS nested gateway, sub-branch to cross 01.02.01
    pg.evaluate("structuralEditor('01.01.08','connector','decision')"); pg.wait_for_timeout(300)
    r=pg.query_selector_all('#branchRows .branchRow')[0]
    print('branch0', r.query_selector('.branchLabel').input_value())
    r.query_selector('.branchTarget').select_option('__NEWGATE__'); r.query_selector('.newGateQ').fill('Perlu revisi minor?')
    subs=r.query_selector_all('.subBranch'); subs[0].query_selector('.subTarget').select_option('01.01.07'); subs[1].query_selector('.subTarget').select_option('01.01.04')
    pg.click('#structSave'); pg.wait_for_timeout(1000)
    print('err:', pg.evaluate("document.querySelector('#structError')?.textContent||''"))
    print(pg.evaluate("state.rows.filter(r=>r.l4no==='01.01'&&/08/.test(r.l5no)).map(r=>r.l5no+' | '+r.prev+' -> '+r.next)"))
    print('01.01.04 prev:', pg.evaluate("state.rows.find(r=>r.l5no==='01.01.04').prev"))
    print('check', pg.evaluate(CHECK,'01.01'), 'warn', pg.evaluate("state.warnings.filter(w=>w.l4==='01.01').map(w=>w.msg).slice(0,3)"))
    print('fileInfo:', pg.evaluate("$('fileInfo').textContent"))
    print('errors',errs); b.close()
