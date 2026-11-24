from playwright.sync_api import sync_playwright
URL='file:///home/claude/out/Flowchart_BPA_v1.9.2_BPMN.html'
CHECK=open('/home/claude/check2.py').read().split('JS = r"""')[1].split('"""')[0]
ROWS="(l4)=>structuralRows(state.l4s.get(l4),$('processEditMode').value).map(r=>r.l5no+' | '+r.prev+' -> '+r.next)"
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1600,'height':1100}); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('dialog',lambda d: d.accept())
    pg.goto(URL); pg.wait_for_timeout(1200); pg.set_input_files('#fileIn','/home/claude/wb/pelni.xlsx'); pg.wait_for_timeout(4500)
    def err(): return pg.evaluate("document.querySelector('#evError')?.textContent||''")
    # --- trigger in the middle: 03.04.02 triggered by 03.03.07
    pg.evaluate("select('03.04')"); pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(600)
    pg.click('[data-ev=trig]'); pg.wait_for_timeout(300); print('panel:', pg.inner_text('#detail h3'), '|', pg.inner_text('#detail .muted')[:60])
    pg.click('#evRemove'); pg.wait_for_timeout(800); print('hapus pemicu tengah -> err:', err())
    print([x for x in pg.evaluate(ROWS,'03.04') if x.startswith('03.04.02')], 'trig left:', pg.evaluate("document.querySelectorAll('[data-ev=trig]').length"))
    print('check', pg.evaluate(CHECK,'03.04'))
    # --- start: remove Mulai on 01.01.01 (also triggered by 01.03.06 -> still reachable)
    pg.evaluate("select('01.01')"); pg.wait_for_timeout(500)
    pg.click('[data-symbol=start]'); pg.wait_for_timeout(300); print('start panel:', pg.inner_text('#detail h3'))
    pg.click('#evRemove'); pg.wait_for_timeout(800); print('hapus Mulai 01.01 -> err:', err()); print([x for x in pg.evaluate(ROWS,'01.01') if x.startswith('01.01.01')])
    # --- start remove where it would be unreachable: 12.02 (only Mulai)
    pg.evaluate("select('12.02')"); pg.wait_for_timeout(500); pg.click('[data-symbol=start]'); pg.wait_for_timeout(300)
    pg.click('#evRemove'); pg.wait_for_timeout(500); print('hapus Mulai 12.02 (tanpa pengganti) -> err:', err())
    pg.select_option('#evAddTo','12.02.02'); pg.click('#evRemove'); pg.wait_for_timeout(800); print('pindah Mulai ke 12.02.02 -> err:', err())
    print(pg.evaluate(ROWS,'12.02')[:3]); print('check', pg.evaluate(CHECK,'12.02'))
    # --- end: 12.02 Selesai after 12.02.05 branch "Tidak ada temuan" -> remove with nextTo
    pg.click('[data-symbol=end]'); pg.wait_for_timeout(300); print('end panel:', pg.inner_text('#detail .muted')[:60])
    pg.click('#evRemove'); pg.wait_for_timeout(500); print('hapus Selesai tanpa lanjut ->', err())
    pg.select_option('#evNext','12.02.04'); pg.click('#evRemove'); pg.wait_for_timeout(800); print('hapus Selesai, lanjut ke 12.02.04 ->', err(), [x for x in pg.evaluate(ROWS,'12.02') if x.startswith('12.02.05')])
    print('ends left:', pg.evaluate("document.querySelectorAll('[data-symbol=end]').length"), 'check', pg.evaluate(CHECK,'12.02'))
    # --- handoff: 12.01.10 -> 12.03.01 remove (AS-IS) and check other L4 updated
    pg.evaluate("select('12.01')"); pg.select_option('#processEditMode','asis'); pg.wait_for_timeout(600)
    print('12.03.01 prev before:', pg.evaluate("state.rows.find(r=>r.l5no==='12.03.01').prev"))
    pg.click('[data-ev=hand]'); pg.wait_for_timeout(300); print('hand panel:', pg.inner_text('#detail h3'))
    pg.click('#evRemove'); pg.wait_for_timeout(900); print('hapus serah terima -> err:', err())
    print('12.01.10 next:', pg.evaluate("state.rows.find(r=>r.l5no==='12.01.10').next"), '| 12.03.01 prev:', pg.evaluate("state.rows.find(r=>r.l5no==='12.03.01').prev"))
    print('check', pg.evaluate(CHECK,'12.01'))
    # --- label edit
    pg.evaluate("select('01.02')"); pg.wait_for_timeout(400); pg.click('[data-symbol=end] >> nth=0'); pg.wait_for_timeout(300)
    pg.fill('#evLabel','Selesai — perubahan dijalankan'); pg.click('#evSaveLabel'); pg.wait_for_timeout(700)
    print('label:', pg.evaluate("[...document.querySelectorAll('[data-symbol=end] text')].map(t=>t.textContent).join(' ')")[:60], err())
    # --- unit rename
    pg.evaluate("select('12.01')"); pg.select_option('#processEditMode','asis'); pg.wait_for_timeout(500)
    pg.click('#workspaceOrg'); pg.wait_for_timeout(400); pg.click('.rcard >> nth=0 >> [data-a=laneUnit]'); pg.wait_for_timeout(300)
    print('unit lama:', pg.input_value('#laneNewUnit')); pg.fill('#laneNewUnit','Pelayanan SDM'); pg.check('#laneUnitAll'); pg.click('[data-a=laneUnitGo]'); pg.wait_for_timeout(1200)
    print('toast:', pg.evaluate("document.querySelector('#toast').textContent"))
    print('units 12.01:', pg.evaluate("[...new Set(state.rows.filter(r=>r.l4no==='12.01').map(r=>r.unitX))]"), '| 12.03:', pg.evaluate("[...new Set(state.rows.filter(r=>r.l4no==='12.03').map(r=>r.unitX))]"))
    print('svg has new unit:', 'Pelayanan SDM' in pg.evaluate("document.querySelector('#canvas svg').textContent"))
    # dblclick on unit subtitle
    pg.evaluate("select('01.01')"); pg.wait_for_timeout(500)
    pos=pg.evaluate("(()=>{const d=state.diagram;const ln=d.lanes.filter(l=>l.kind==='act')[0];const u=laneUnitOf(d.lanes.indexOf(ln));const t=[...document.querySelectorAll('#canvas svg text')].find(t=>t.textContent.trim()===u);const r=t.getBoundingClientRect();return [r.x+r.width/2,r.y+r.height/2,u]})()")
    pg.mouse.dblclick(pos[0],pos[1]); pg.wait_for_timeout(500); print('dblclick unit ->', pg.evaluate("$('laneNewUnit')?.value"), 'expected', pos[2])
    print('errors', errs); b.close()
