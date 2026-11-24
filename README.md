# Flowchart BPA

Aplikasi internal (satu file HTML, tanpa server) untuk membaca workbook BPA (Excel) dan membuat flowchart swimlane bernotasi BPMN 2.0, lengkap dengan edit AS-IS/TO-BE untuk workshop dan ekspor PDF / Word / Excel / Master Workbook.

## Struktur
- `src/app.html` — kerangka HTML + library (SheetJS, jsPDF, AutoTable, docx) yang sudah di-inline.
- `src/app.js` — kode aplikasi.
- `build.py` — menyuntikkan `src/app.js` ke kerangka dan menulis `dist/Flowchart_BPA_<versi>_BPMN.html`. Versi diatur di variabel `VER`.
- `dist/` — file siap pakai (buka langsung di browser).
- `tests/` — skrip uji Playwright (render semua L4, cek tumpang tindih, uji acak edit struktur, ekspor/impor). Path workbook uji di dalam skrip perlu disesuaikan.
- `docs/CHANGELOG.md` — riwayat perubahan.

## Build
```
python3 build.py
```
