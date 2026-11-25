import re
s = open('src/app.html', encoding='utf-8').read()
app = open('src/app.js', encoding='utf-8').read() + '\n' + open('src/cloud.js', encoding='utf-8').read()
sblib = open('vendor/supabase-2.119.2.js', encoding='utf-8').read().replace('</script', '<\\/script')
blocks = list(re.finditer(r'(<script[^>]*>)(.*?)(</script>)', s, re.S))
b = blocks[4]
assert b.group(2).lstrip().startswith('"use strict";')
s = s[:b.start(2)] + app + s[b.end(2):]
# v1.9.0: Supabase client library (UMD) right before the app script
s = s[:b.start(1)] + '<script>/* @supabase/supabase-js 2.119.2 (MIT) */\n' + sblib + '\n</script>' + s[b.start(1):]
VER = 'v1.11.1'
s = re.sub(r'<title>.*?</title>', '<title>Flowchart BPA '+VER+' — Notasi BPMN 2.0</title>', s, count=1, flags=re.S)
s = s.replace('Professional Workspace · v1.3.8', 'Professional Workspace · '+VER, 1)
old = '<label class="status-in">Label status <input id="statusLabel" value="DRAFT — Konsultan"></label>'
assert s.count(old) == 1
s = s.replace(old, old + '<label class="status-in">Nama pool BPMN <input id="poolLabel" value="Organisasi" title="Nama partisipan/organisasi yang tampil di pita pool diagram"></label><label class="status-in" title="BPMN ketat: beberapa pemicu digabung lewat gateway XOR sebelum aktivitas pertama"><input type="checkbox" id="mergeGw"> Tampilkan gateway penggabung di awal</label><label class="status-in" title="Semua jalur yang selesai bertemu di satu titik Selesai di bagian bawah"><input type="checkbox" id="oneEndCk" checked> Satu titik Selesai</label><label class="status-in">Legenda dokumen <input id="legendDoc" placeholder="Dokumen (Data Object)" title="Teks legenda untuk ikon kertas"></label><label class="status-in">Legenda data <input id="legendData" placeholder="Data di sistem (Data Store)" title="Teks legenda untuk ikon silinder"></label>')
s = s.replace('</head>', '<style>/* v1.6.6 */ .modal .dlg .row{flex-wrap:wrap;gap:8px;justify-content:flex-end} .modal .dlg .row .btn{white-space:nowrap}</style></head>', 1)
out = 'dist/Flowchart_BPA_'+VER+'_BPMN.html'
import os; os.makedirs('dist', exist_ok=True)
open(out, 'w', encoding='utf-8').write(s)
print(out, len(s))

# GitHub Pages: the latest build is also served as index.html
import shutil; shutil.copyfile(out, 'index.html'); print('index.html updated')
