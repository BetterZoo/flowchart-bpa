import sys, os, json
from playwright.sync_api import sync_playwright
JS = r"""
(code)=>{ select(code); const svg=document.querySelector('#canvas svg'); if(!svg) return {code,err:'nosvg'};
 const d=state.diagram; const poolR=Math.max(...d.lanes.map(l=>l.x))+d.laneW;
 const acts=d.nodes.map(n=>({x:n.cx-n.w/2,y:n.cy-n.h/2,w:n.w,h:n.h,n}));
 // 1) data associations must end on a node edge
 let dangling=0; const dd=[];
 for(const p of svg.querySelectorAll('path[marker-end="url(#arA)"]')){
   const nums=(p.getAttribute('d').match(/-?[\d.]+/g)||[]).map(Number); if(nums.length<2) continue;
   // rebuild absolute points from M/H/V commands
   const cmds=p.getAttribute('d').match(/[MLHV][^MLHV]*/g); let x=0,y=0;
   for(const c of cmds){ const v=(c.slice(1).match(/-?[\d.]+/g)||[]).map(Number); if(c[0]==='M'||c[0]==='L'){x=v[0];y=v[1];} else if(c[0]==='H'){x=v[0];} else if(c[0]==='V'){y=v[0];} }
   if(p.getTotalLength()<40) continue; // legend sample
   const ok=acts.some(b=>x>=b.x-8&&x<=b.x+b.w+8&&y>=b.y-8&&y<=b.y+b.h+8) || [...svg.querySelectorAll('path')].some(()=>false);
   if(!ok){ // allow ends on a data glyph (output associations)
     const near=[...svg.querySelectorAll('path,rect')].some(e=>{ if(e===p) return false; const bb=e.getBBox(); return bb.width<40&&bb.height<45&&bb.width>10&&x>=bb.x-8&&x<=bb.x+bb.width+8&&y>=bb.y-6&&y<=bb.y+bb.height+6; });
     if(!near){dangling++; dd.push(Math.round(x)+','+Math.round(y));}
   }
 }
 // 2) overlapping texts inside the pool
 const T=[...svg.querySelectorAll('text')].map(t=>{const b=t.getBBox();return {t:t.textContent,x:b.x,y:b.y,w:b.width,h:b.height};}).filter(b=>b.x<poolR&&b.w>0&&b.y>140);
 let tov=0; const td=[];
 for(let i=0;i<T.length;i++)for(let j=i+1;j<T.length;j++){const a=T[i],b=T[j];const ox=Math.min(a.x+a.w,b.x+b.w)-Math.max(a.x,b.x), oy=Math.min(a.y+a.h,b.y+b.h)-Math.max(a.y,b.y); if(ox>3&&oy>3){tov++; td.push(a.t.slice(0,18)+' | '+b.t.slice(0,18));}}
 // 3) text over a task box it does not belong to
 let tin=0; const ti=[];
 for(const a of T) for(const b of acts.filter(b=>b.n.type==='act')){ const ox=Math.min(a.x+a.w,b.x+b.w)-Math.max(a.x,b.x), oy=Math.min(a.y+a.h,b.y+b.h)-Math.max(a.y,b.y);
   if(ox>3&&oy>3 && !(a.x>=b.x-2&&a.x+a.w<=b.x+b.w+2&&a.y>=b.y-12&&a.y+a.h<=b.y+b.h+2)){tin++; ti.push(a.t.slice(0,20)+' @'+b.n.code);} }
 // 4) sequence flows crossing text
 let ex=0; const ed=[];
 for(const e of d.edges){ if(!e.pts||e.side) continue; for(let k=1;k<e.pts.length;k++){const [x1,y1]=e.pts[k-1],[x2,y2]=e.pts[k];
   for(const a of T){ if(e.label && a.t && e.label.includes(a.t.slice(0,6))) continue;
     const hit = x1===x2 ? (x1>a.x+1&&x1<a.x+a.w-1&&Math.max(y1,y2)>a.y+2&&Math.min(y1,y2)<a.y+a.h-2) : (y1===y2 && y1>a.y+2&&y1<a.y+a.h-2&&Math.max(x1,x2)>a.x+1&&Math.min(x1,x2)<a.x+a.w-1);
     if(hit){ex++; ed.push((e.s.code||e.s.type)+'>'+(e.t.code||e.t.type)+' x "'+a.t.slice(0,20)+'"');} } } }
 return {code,dangling,dd:dd.slice(0,4),tov,td:td.slice(0,4),tin,ti:ti.slice(0,4),ex,ed:ed.slice(0,6)};
}
"""
url, wb = sys.argv[1], sys.argv[2]
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1600,'height':1100}); errs=[]
    pg.on('pageerror',lambda e: errs.append(str(e)))
    pg.goto('file://'+url); pg.wait_for_timeout(1200)
    pg.set_input_files('#fileIn',wb); pg.wait_for_timeout(4500)
    if pg.query_selector('select[aria-label="Sheet SOP"]'):
        pg.select_option('select[aria-label="Sheet SOP"]','SOP (3)'); pg.click('button:has-text("Import sheet")'); pg.wait_for_timeout(5000)

    if os.environ.get('DOCS_ALL'): pg.evaluate("state.rows.forEach((r,i)=>{const x=state.extra[r.l5no]=state.extra[r.l5no]||{}; x.docs=[{d:'in',k:'object',n:'Formulir uji '+i},{d:'out',k:i%2?'store':'object',n:'Laporan hasil uji '+i}]})")
    if os.environ.get('SLA_ALL'): pg.evaluate("state.rows.forEach(r=>{state.extra[r.l5no]=Object.assign(state.extra[r.l5no]||{},{sla:'2 HK sejak dokumen lengkap'})})")
    for mode in sys.argv[3:] or ['tobe']:
        pg.select_option('#processEditMode',mode); pg.wait_for_timeout(300)
        res=[pg.evaluate(JS,c) for c in pg.evaluate("[...state.l4s.keys()]")]
        tot={k:sum(r.get(k,0) for r in res) for k in ['dangling','tov','tin','ex']}
        print(mode,'TOTAL',tot)
        for r in res:
            if r.get('dangling') or r.get('tov') or r.get('tin') or r.get('ex'): print(' ',r)
    print('errors',errs[:3]); b.close()
