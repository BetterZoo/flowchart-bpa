# Changelog Flowchart BPA

## v1.9.5 — Nomor runut di sheet "SOP TO-BE" (2026-11-24)
- L4 yang punya rancangan TO-BE dinomori ulang mengikuti urutan alur (12.01.01, 12.01.02, … tanpa bolong); gateway lanjutan diberi kode <aktivitas sebelumnya>-G1.
- Proses Sebelum/Sesudah di sheet ikut nomor baru, termasuk rujukan dari L4 lain.
- Kolom baru "No L5 lama" berisi nomor asal bila berubah. Baris "Dihapus di TO-BE": No L5 kosong, nomor asal di "No L5 lama".
- Data di aplikasi tidak diubah; penomoran ulang hanya di sheet hasil ekspor.

## v1.9.4 — Sheet "SOP TO-BE" di Master Workbook (2026-11-24)
- Export → Excel → Simpan Master Workbook sekarang berisi sheet "SOP TO-BE": kolom sama dengan sheet SOP, isi versi TO-BE terbaru per L4 (L4 tanpa rancangan TO-BE diisi AS-IS), plus kolom "Status TO-BE": Tetap / Berubah / Baru / Dihapus di TO-BE (gateway lanjutan ditandai "(gateway)").
- Sheet SOP tetap AS-IS. Saat Master Workbook diimpor lagi, "SOP TO-BE" tidak dipakai sebagai sumber (rancangan TO-BE tetap dipulihkan dari Elaborasi Flowchart).
- v1.9.3: panel Cloud menampilkan detail pengecekan anggota + tombol "Cek lagi"; docs/supabase_fix_members.sql.

## v1.9.2 — Satu titik Selesai (2026-11-24)
- Semua jalur yang berakhir sekarang bertemu di satu titik Selesai di bagian bawah, di lane jalur terakhir yang selesai (sebelumnya satu Selesai per aktivitas/gateway).
- Bisa dikembalikan ke cara lama: Pengaturan → hilangkan centang "Satu titik Selesai".
- Klik titik Selesai bersama → pilih jalur mana yang mau diubah ("Selesai setelah"); label Selesai berlaku untuk semua jalur.
- Serah terima ke L4 lain ("Ke 12.03.01") tetap berupa titik akhir terpisah.

## v1.9.1 — Edit Mulai/Selesai & ganti nama unit (2026-11-24)
- Klik lingkaran Mulai, Selesai, "Dari L4 lain", atau "Ke L4 lain" di canvas → panel Detail Simbol:
  - Mulai: ubah label, hapus, atau pindahkan ke aktivitas lain; opsi "sambungkan dari" bila aktivitasnya jadi tanpa pendahulu.
  - Pemicu dari L4 lain: hapus atau pindahkan ke aktivitas lain (AS-IS: Proses Sesudah di L4 asal ikut disesuaikan).
  - Selesai: ubah label, hapus dengan "lanjut ke" (label cabang dipertahankan), atau "tambahkan Selesai setelah" aktivitas lain.
  - Serah terima ke L4 lain: hapus, opsional "lanjut ke" (AS-IS: Proses Sebelum di L4 tujuan ikut disesuaikan).
  - Pengaman: perubahan yang memutus alur ditolak dengan petunjuk.
- Ganti nama unit: klik dua kali nama unit di judul lane, atau tombol "✎ Ganti unit" di Pengaturan Swimlane; opsi "Terapkan ke semua L4 yang memakai unit ini" (TO-BE: hanya L4 yang sudah punya rancangan; AS-IS: L4 yang punya rancangan TO-BE dilewati).

## v1.9.0 — Simpan ke cloud (Supabase) (2026-11-24)
- Tombol "☁ Cloud": login email tanpa password (magic link), hanya email yang terdaftar di tabel `members`.
- Simpan pekerjaan di browser sebagai proyek cloud tim; buka proyek tim dari laptop mana pun.
- Simpan otomatis (cek tiap 8 detik bila ada perubahan) atau manual; status simpan tampil di tombol.
- Riwayat versi: saat dibuat, saat simpan manual, dan otomatis paling sering tiap 15 menit; 30 versi terakhir; bisa dipulihkan.
- Pengaman: bila anggota lain menyimpan lebih dulu, simpan otomatis berhenti dan muncul pilihan "Muat versi terbaru" / "Timpa dengan versi saya".
- Tanpa login aplikasi tetap jalan seperti biasa (tersimpan di browser).

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
