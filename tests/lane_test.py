from playwright.sync_api import sync_playwright
URL='file:///home/claude/out/Flowchart_BPA_v1.8.0_BPMN.html'
lanes="state.diagram.lanes.filter(l=>l.kind==='act').map(l=>l.key.replace(/^pos::/,''))"
with sync_playwright() as p:
    b=p.chromium.launch(); ctx=b.new_context(viewport={'width':1600,'height':1100},accept_downloads=True); pg=ctx.new_page(); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e))); pg.on('dialog',lambda d: d.accept())
    pg.add_init_script("delete window.showSaveFilePicker; delete window.showDirectoryPicker;")
    pg.goto(URL); pg.wait_for_timeout(1200)
    pg.set_input_files('#fileIn','/home/claude/wb/pelni.xlsx'); pg.wait_for_timeout(4000)
    pg.evaluate("select('12.01')"); pg.wait_for_timeout(500); pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(500)
    print('awal      :', pg.evaluate(lanes))
    pg.click('#workspaceOrg'); pg.wait_for_timeout(500)
    print('panel     :', pg.inner_text('#detail')[:160].replace('\n',' | '))
    # add lane after first
    pg.fill('#newLanePos','Manager Layanan SDM'); pg.fill('#newLaneUnit','Layanan SDM'); pg.select_option('#newLaneAt','1')
    pg.click('[data-a=laneAdd]'); pg.wait_for_timeout(700)
    print('tambah@2  :', pg.evaluate(lanes))
    # move new lane to right end
    for _ in range(3):
        btn=pg.query_selector('.rcard[data-lane="pos::Manager Layanan SDM"] [data-a=laneDown]')
        if btn and btn.is_enabled(): btn.click(); pg.wait_for_timeout(500)
    print('geser kanan:', pg.evaluate(lanes))
    # move Supervisor to left
    pg.click('.rcard[data-lane="pos::Supervisor Pengelolaan Payroll dan Personalia"] [data-a=laneUp]'); pg.wait_for_timeout(500)
    print('geser kiri :', pg.evaluate(lanes))
    svg=pg.evaluate("document.querySelector('#canvas svg').textContent")
    print('lane baru di canvas:', 'Manager Layanan SDM' in svg)
    # remove lane with activities (Officer) -> move to Staff in TO-BE
    pg.click('.rcard[data-lane="pos::Officer Administrasi Personalia"] [data-a=laneRemove]'); pg.wait_for_timeout(400)
    pg.select_option('.rcard[data-lane="pos::Officer Administrasi Personalia"] [data-rf=target]','pos::Staff Administrasi Personalia')
    pg.click('[data-a=laneRemoveGo]'); pg.wait_for_timeout(900)
    print('hapus Officer (TO-BE):', pg.evaluate(lanes))
    pg.select_option('#processEditMode','asis'); pg.wait_for_timeout(600)
    print('AS-IS tetap :', pg.evaluate(lanes))
    pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(600)
    # remove empty added lane
    pg.click('#workspaceOrg'); pg.wait_for_timeout(400)
    pg.click('.rcard[data-lane="pos::Manager Layanan SDM"] [data-a=laneRemove]'); pg.wait_for_timeout(600)
    print('hapus lane kosong:', pg.evaluate(lanes))
    # auto reset
    if pg.query_selector('[data-a=laneAuto]'): pg.click('[data-a=laneAuto]'); pg.wait_for_timeout(500)
    print('reset otomatis:', pg.evaluate(lanes))
    # manual order + apply to SOP
    pg.click('.rcard[data-lane="pos::Supervisor Pengelolaan Payroll dan Personalia"] [data-a=laneUp]'); pg.wait_for_timeout(500)
    pg.click('[data-a=laneApplySop]'); pg.wait_for_timeout(500)
    print('order tersimpan:', pg.evaluate("JSON.stringify(Object.keys(state.roleEdit.order))"))
    # master roundtrip
    pg.click('#btnExport'); pg.wait_for_timeout(400); pg.click('label:has(input[name=fmt][value=xlsx])'); pg.wait_for_timeout(300)
    pg.click('#exportDlg label:has-text("Simpan Master Workbook")'); pg.wait_for_timeout(300)
    with pg.expect_download(timeout=60000) as d: pg.click('#exGo')
    d.value.save_as('/home/claude/master_lane.xlsx')
    before=pg.evaluate(lanes)
    p2=b.new_context(viewport={'width':1600,'height':1100}).new_page(); p2.on('dialog',lambda d: d.accept()); p2.on('pageerror',lambda e: errs.append(str(e)))
    p2.goto(URL); p2.wait_for_timeout(1200); p2.set_input_files('#fileIn','/home/claude/master_lane.xlsx'); p2.wait_for_timeout(5000)
    p2.evaluate("select('12.01')"); p2.wait_for_timeout(500); p2.select_option('#processEditMode','tobe'); p2.wait_for_timeout(600)
    print('pulihkan sama:', p2.evaluate(lanes)==before, p2.evaluate(lanes))
    print('errors',errs); b.close()
