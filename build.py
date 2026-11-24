import re
s = open('src/app.html', encoding='utf-8').read()
app = open('src/app.js', encoding='utf-8').read()
blocks = list(re.finditer(r'(<script[^>]*>)(.*?)(</script>)', s, re.S))
b = blocks[4]
assert b.group(2).lstrip().startswith('"use strict";')
s = s[:b.start(2)] + app + s[b.end(2):]
VER = 'v1.8.0'
s = re.sub(r'<title>.*?</title>', '<title>Flowchart BPA '+VER+' — Notasi BPMN 2.0</title>', s, count=1, flags=re.S)
s = s.replace('Professional Workspace · v1.3.8', 'Professional Workspace · '+VER, 1)
old = '<label class="status-in">Label status <input id="statusLabel" value="DRAFT — Konsultan"></label>'
assert s.count(old) == 1
s = s.replace(old, old + '<label class="status-in">Nama pool BPMN <input id="poolLabel" value="Organisasi" title="Nama partisipan/organisasi yang tampil di pita pool diagram"></label><label class="status-in" title="BPMN ketat: beberapa pemicu digabung lewat gateway XOR sebelum aktivitas pertama"><input type="checkbox" id="mergeGw"> Tampilkan gateway penggabung di awal</label><label class="status-in">Legenda dokumen <input id="legendDoc" placeholder="Dokumen (Data Object)" title="Teks legenda untuk ikon kertas"></label><label class="status-in">Legenda data <input id="legendData" placeholder="Data di sistem (Data Store)" title="Teks legenda untuk ikon silinder"></label>')
s = s.replace('</head>', '<style>/* v1.6.6 */ .modal .dlg .row{flex-wrap:wrap;gap:8px;justify-content:flex-end} .modal .dlg .row .btn{white-space:nowrap}</style></head>', 1)
out = 'dist/Flowchart_BPA_'+VER+'_BPMN.html'
import os; os.makedirs('dist', exist_ok=True)
open(out, 'w', encoding='utf-8').write(s)
print(out, len(s))
