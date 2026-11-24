# Changelog Flowchart BPA

## v1.8.0 — Panel gateway & gateway setelah gateway (2026-11-24)
- Panel gateway: klik dua kali belah ketupat di canvas (mode AS-IS / TO-BE) → "Edit gateway": ubah pertanyaan, label & tujuan tiap cabang, tambah/hapus kondisi.
- Gateway setelah gateway: tujuan cabang "◇ Gateway baru setelah cabang ini…" → pertanyaan + kondisi (min. 2). Bisa bertingkat.
  - Disimpan sebagai baris khusus berkode turunan (mis. 12.01.08-G1), Level 5 = "Gateway: pertanyaan". Tidak dihitung sebagai aktivitas.
  - Panel gateway bertingkat: "＋ Tambah aktivitas di cabang / sebelum gateway" dan "Hapus gateway ini".
- Cabang paralel ("Disetujui: SELESAI + 01.02.01") bisa diedit tanpa terpecah; gateway yang terhubung ke L4 lain bisa diedit.
- Master Workbook berisi gateway bertingkat harus dibuka dengan v1.8.0+.

## v1.7.1 — Tambah & hapus aktivitas, input/output tambahan (2026-11-24)
- Tambah aktivitas: sesudah / sebelum (termasuk jadi pertama/terakhir) / sebelum gateway / di cabang gateway; pilih lane di form.
- Dokumen dari L4 lain tetap di aktivitas yang memakainya saat aktivitas baru disisipkan di depannya (kolom tersembunyi "Dokumen dari L4 lain dipakai di").
- Hapus aktivitas di posisi mana pun; aktivitas bergateway: pindahkan gateway ke pendahulu atau pertahankan satu cabang.
- Hubungan lintas L4 disesuaikan otomatis di AS-IS; catatan di TO-BE.
- Input / output tambahan per aktivitas (Detail Aktivitas), tersimpan di sheet "SLA & Output" kolom "Input/Output Tambahan".

## v1.7.0 — Tahap 1 untuk workshop (2026-11-24)
- Ganti nama lane langsung (klik dua kali judul lane / tombol di Pengaturan Swimlane).
- Start/End Event lintas L4 di lane yang sesuai, tanpa amplop; berdampingan dengan Mulai.
- Gateway XOR tanpa tanda X.
- SLA sebagai Text Annotation di atas jalur ikon.
- Perbaikan celah kosong dan panel Keterangan.

## v1.6.7 — Tanda "X" untuk dokumen input (2026-11-23)
## v1.6.6 — Impor file (2026-11-23)
## v1.6.5 — Input lintas-L4 mengikuti output sumber (2026-11-23)
## v1.6.4 — Jenis output di aplikasi & legenda (2026-11-23)
## v1.6.3 — Ikon dokumen/data & pengecekan menyeluruh (2026-11-23)
## v1.6.2 — Garis balik di aktivitas berikon data (2026-11-23)
## v1.6.1 — Penanda "X" pada Dokumen/Sistem (2026-11-23)
## v1.6.0 — Pengaturan Swimlane (2026-11-23)
## v1.5.1 — Perbaikan backup Master Workbook (2026-11-23)
## v1.5.0 — Label, garis alur, jabatan (2026-11-23)
## v1.4.x — Notasi BPMN 2.0 (2026-11-23)

## Cara kerja ikon dokumen/data
1. Kolom U "Dokumen" berisi "-" → tidak ada simbol.
2. U terisi (nama atau "X") → Dokumen; W terisi → Data di sistem.
3. "X" mengambil nama dari Output (T) → ikon OUTPUT; T kosong tapi Input (S) terisi → ikon INPUT.
4. U dan W kosong tetapi T terisi → otomatis Dokumen output.
5. Input/output lain → Detail Aktivitas → "Input / output tambahan".
