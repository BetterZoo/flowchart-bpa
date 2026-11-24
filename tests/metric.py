import sys, os, json
from playwright.sync_api import sync_playwright
METRIC = r"""
(code)=>{ select(code); const d=state.diagram; if(!d) return null;
 const E=d.edges.filter(e=>e.pts&&e.pts.length>1), N=d.nodes;
 const seg=e=>e.pts.slice(1).map((b,i)=>[e.pts[i],b]);
 let overlap=0, hits=0, cross=0, railHits=0, labelHits=0, near=0; const det=[];
 const boxes=N.map(n=>({n,x:n.cx-n.w/2,y:n.cy-n.h/2,w:n.w,h:n.h}));
 const rails=N.filter(n=>n.type==='act'&&(n.tag||n.sla||(n.crossInputs&&n.crossInputs.length))).map(n=>{const ln=d.lanes[n.lane];const c=ln.x+d.laneW-84;return {n,x:c-48,y:n.cy-6,w:96,h:80};});
 const labs=N.filter(n=>n._evBox||n._qBox).flatMap(n=>[n._evBox,n._qBox].filter(Boolean).map(b=>({n,...b})));
 const hitBox=(a,b,B)=>{const [x,y]=a,[u,v]=b; if(x===u) return x>B.x+1&&x<B.x+B.w-1&&Math.max(y,v)>B.y+1&&Math.min(y,v)<B.y+B.h-1; if(y===v) return y>B.y+1&&y<B.y+B.h-1&&Math.max(x,u)>B.x+1&&Math.min(x,u)<B.x+B.w-1; return false;};
 for(const e of E){ for(const [a,b] of seg(e)){
   for(const B of boxes){ if(B.n===e.s||B.n===e.t) continue; if(hitBox(a,b,B)){hits++; det.push('hit '+(e.s.code||e.s.type)+'>'+(e.t.code||e.t.type)+' x '+(B.n.code||B.n.type));} }
   for(const B of rails){ if(B.n===e.s||B.n===e.t) continue; if(hitBox(a,b,B)){railHits++; det.push('rail '+(e.s.code||e.s.type)+'>'+(e.t.code||e.t.type)+' x '+B.n.code);} }
   if(!e.side) for(const B of labs){ if(B.n===e.s||B.n===e.t) continue; if(hitBox(a,b,B)){labelHits++; det.push('label '+(e.s.code||e.s.type)+'>'+(e.t.code||e.t.type)+' x '+(B.n.code||B.n.type));} }
 }}
 for(let i=0;i<E.length;i++)for(let j=i+1;j<E.length;j++){ const e=E[i],f=E[j]; const sameT=e.t===f.t, sameS=e.s===f.s;
   for(const [[x,y],[u,v]] of seg(e)) for(const [[a,b],[c,dd]] of seg(f)){
     if(!sameT&&!sameS){
       if(x===u&&a===c&&Math.abs(x-a)<2){const o=Math.min(Math.max(y,v),Math.max(b,dd))-Math.max(Math.min(y,v),Math.min(b,dd)); if(o>3){overlap+=o; det.push('ovl '+(e.s.code||e.s.type)+'>'+(e.t.code||e.t.type)+' & '+(f.s.code||f.s.type)+'>'+(f.t.code||f.t.type)+' '+Math.round(o));}}
       if(y===v&&b===dd&&Math.abs(y-b)<2){const o=Math.min(Math.max(x,u),Math.max(a,c))-Math.max(Math.min(x,u),Math.min(a,c)); if(o>3){overlap+=o; det.push('ovl '+(e.s.code||e.s.type)+'>'+(e.t.code||e.t.type)+' & '+(f.s.code||f.s.type)+'>'+(f.t.code||f.t.type)+' '+Math.round(o));}}
     }
     if(!sameT&&!sameS){
       if(x===u&&a===c&&Math.abs(x-a)>=2&&Math.abs(x-a)<9){const o=Math.min(Math.max(y,v),Math.max(b,dd))-Math.max(Math.min(y,v),Math.min(b,dd)); if(o>25){near++; det.push('near '+(e.s.code||e.s.type)+'>'+(e.t.code||e.t.type)+' & '+(f.s.code||f.s.type)+'>'+(f.t.code||f.t.type)+' '+Math.round(o));}}
       if(y===v&&b===dd&&Math.abs(y-b)>=2&&Math.abs(y-b)<9){const o=Math.min(Math.max(x,u),Math.max(a,c))-Math.max(Math.min(x,u),Math.min(a,c)); if(o>25){near++; det.push('near '+(e.s.code||e.s.type)+'>'+(e.t.code||e.t.type)+' & '+(f.s.code||f.s.type)+'>'+(f.t.code||f.t.type)+' '+Math.round(o));}}
     }
     if(x===u&&b===dd){ if(x>Math.min(a,c)+1&&x<Math.max(a,c)-1&&b>Math.min(y,v)+1&&b<Math.max(y,v)-1) cross++; }
     else if(y===v&&a===c){ if(a>Math.min(x,u)+1&&a<Math.max(x,u)-1&&y>Math.min(b,dd)+1&&y<Math.max(b,dd)-1) cross++; }
   }}
 return {code,overlap:Math.round(overlap),hits,railHits,labelHits,cross,near,det:det.slice(0,12)};
}
"""
url, wb = sys.argv[1], sys.argv[2]
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1600,'height':1100})
    pg.goto('file://'+url); pg.wait_for_timeout(1200)
    pg.set_input_files('#fileIn',wb); pg.wait_for_timeout(4000)
    if os.environ.get('DOCS_ALL'): pg.evaluate("state.rows.forEach((r,i)=>{const x=state.extra[r.l5no]=state.extra[r.l5no]||{}; x.docs=[{d:'in',k:'object',n:'Formulir uji '+i},{d:'out',k:i%2?'store':'object',n:'Laporan hasil uji '+i}]})")
    if os.environ.get('SLA_ALL'): pg.evaluate("state.rows.forEach(r=>{state.extra[r.l5no]=Object.assign(state.extra[r.l5no]||{},{sla:'2 HK sejak dokumen lengkap'})})")
    pg.select_option('#processEditMode','tobe'); pg.wait_for_timeout(300)
    codes=pg.evaluate("[...state.l4s.keys()]"); res=[]
    for c in codes: res.append(pg.evaluate(METRIC, c))
    tot={k:sum(r[k] for r in res) for k in ['overlap','hits','railHits','labelHits','cross','near']}
    print('TOTAL',tot)
    bad=sorted(res,key=lambda r:-(r['overlap']/10+r['hits']*20+r['railHits']*10+r['labelHits']*10+r['cross']+r['near']*15))[:int(sys.argv[3]) if len(sys.argv)>3 else 8]
    for r in bad: print(r['code'],{k:r[k] for k in tot}, r['det'][:6])
    json.dump(res,open('/home/claude/metric_last.json','w'))
    b.close()
