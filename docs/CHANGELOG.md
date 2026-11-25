# Changelog Flowchart BPA

## v1.12.0 — Gabungkan dengan mapping per L3 (2026-11-25)
- Layar Gabungkan proyek kini punya dua metode: "Bandingkan otomatis" (diff, seperti v1.11) dan "Mapping per L3" (tanpa diff).
- Mapping per L3: pilih dokumen utama, lalu dokumen sumber untuk tiap L3 (default: dokumen utama bila punya L3 itu; ada pintasan "Semua L3 dari …"). Keterangan "diubah / sama dengan awal" hanya informasi.
- Untuk tiap L4 di bawah L3, semua isi diambil utuh dari dokumen yang dipilih: data AS-IS, rancangan TO-BE, SLA/output/catatan (termasuk aktivitas baru TO-BE), pemetaan jabatan/unit, mode lane, pengaturan lane per L4, analisis dan register. Data umum (alias peran, pemetaan keputusan, tampilan jabatan, pengaturan simbol BPMN, nama file) dari dokumen utama. Nomor kode terpakai digabung dari semua dokumen plus kode di hasil.
- Sebelum dibuat ada ringkasan per dokumen dan daftar rujukan Proses Sebelum/Sesudah/Data dari yang mengarah ke aktivitas yang tidak ada di hasil.
- Catatan versi 1 proyek baru mencatat mapping-nya (mis. "L3A ← Idur v3; lainnya ← Dea v3").

## v1.11.1 — Perbaikan gabung proyek yang sumbernya hasil gabungan (2026-11-25)
- Bug: menggabungkan proyek hasil gabungan (mis. "Gabungan — Idur_Dea") dengan proyek lain membuat SLA/catatan dan rancangan TO-BE yang hanya ada di proyek gabungan itu hilang. Penyebab: snapshot awal proyek gabungan sudah berisi semuanya, sehingga isiannya dianggap "tidak diubah", sedangkan proyek lain yang belum punya isian dianggap "menghapus" dengan log lebih baru.
- Perbaikan: "diubah" kini dinilai terhadap satu leluhur bersama, yaitu snapshot awal tertua di antara proyek yang dipilih.

## v1.11.0 — Gabungkan proyek cloud (2026-11-25)
- Panel Cloud → "⇄ Gabungkan proyek…": pilih 2 proyek atau lebih → Bandingkan → tinjau → "Buat proyek gabungan". Hasilnya proyek cloud BARU; proyek asal tidak diubah. Versi 1 proyek baru mencatat sumbernya ("Hasil gabungan: …").
- Perbandingan per bagian kecil: data AS-IS per L4, rancangan TO-BE per L4, SLA/output/catatan per aktivitas, pemetaan jabatan/unit/keputusan/mode lane per L4, pengaturan peran, analisis dan register per entri. Nomor kode terpakai digabung semua.
- Aturan default: bagian yang sama → diambil. Bagian yang berbeda → hanya nilai yang benar-benar diubah (berbeda dari snapshot awal proyek) yang dipertimbangkan; bila diubah di beberapa proyek, diambil yang log simpannya (project_saves) paling baru untuk bagian itu. Bila tidak ada log, dipakai waktu simpan terakhir proyek.
- SLA/catatan untuk aktivitas baru TO-BE ikut proyek yang rancangan TO-BE L4-nya dipakai.
- Pratinjau menampilkan "Perlu perhatian" (diubah di lebih dari satu proyek) dan "Diambil otomatis"; setiap pilihan bisa diganti lewat dropdown.
- Ketepatan waktu per bagian mengikuti log simpan per jenis data (mis. "SLA/Output"), bukan per aktivitas.

## v1.10.1 — Nomor TO-BE runut & Catatan untuk aktivitas baru (2026-11-25)
- Mode TO-BE: nomor aktivitas yang tampil (badge flowchart, panel detail, Catatan, toast) mengikuti urutan alur. Aktivitas yang disisipkan setelah 07.02.01 tampil sebagai 07.02.02 dan nomor berikutnya bergeser; gateway <aktivitas>-G1. Penomoran sama dengan sheet "SOP TO-BE" di Master Workbook.
- Kode internal tetap (mis. 07.02.11), ditampilkan kecil sebagai "id" di panel detail, agar perbandingan AS-IS vs TO-BE, SLA/catatan, dan riwayat tidak rusak. AS-IS tidak dinomori ulang.
- Catatan (SLA/dokumen/catatan): di luar mode AS-IS, daftar diambil dari rancangan TO-BE sehingga aktivitas yang ditambahkan di aplikasi ikut muncul. Isian pada aktivitas AS-IS yang dihapus di TO-BE tetap terlihat dengan tanda "(tidak ada di TO-BE)".
- Sheet "SLA & Output": nama aktivitas untuk aktivitas baru TO-BE kini terisi.

## v1.10.0 — Hapus proyek cloud (2026-11-25)
- Daftar "Proyek tim" di panel ☁ Cloud punya tombol **Hapus** per proyek. Konfirmasi dengan mengetik nama proyek persis.
- Menghapus permanen: riwayat versi dan log simpan proyek itu ikut terhapus. Bila proyek yang dihapus sedang aktif, data di browser tetap ada tetapi dilepas dari cloud.

## v1.9.9 — Versi saat ini di Riwayat (2026-11-25)
- Riwayat versi menampilkan baris "Versi N · saat ini" di paling atas bila isi proyek sekarang belum punya snapshot sendiri (mis. disimpan kurang dari 15 menit setelah snapshot terakhir).

## v1.9.8 — Log simpan di Riwayat (2026-11-25)
- Riwayat versi kini juga menampilkan **Log simpan**: setiap simpan (otomatis/manual) dengan versi, waktu, siapa, dan bagian yang berubah (mis. Rancangan TO-BE, Mode swimlane). Sebelumnya hanya snapshot tiap ±15 menit yang terlihat, sehingga simpanan di antaranya tampak tidak tercatat dan snapshot berikutnya tercatat atas nama orang yang kebetulan menyimpan.
- Database: bila penyunting berganti, versi terakhir penyunting sebelumnya otomatis disimpan ke Riwayat (bisa dipulihkan).
- Perlu dijalankan sekali: docs/supabase_save_log.sql. Tanpa itu aplikasi tetap jalan, hanya Log simpan tidak muncul.
- Nama pengguna ditampilkan tanpa @flowbpa.local.

## v1.9.7 — Login username + password, pengaman simpan cloud (2026-11-25)
- Login cloud memakai username + password (tidak lagi link via email, jadi tidak kena batas kirim email Supabase). Username `dea` = akun `dea@flowbpa.local` di Supabase; mengetik alamat email lengkap juga bisa.
- Panel Cloud: tautan "Ganti password" (minimal 8 karakter).
- Akun dibuat admin, lihat docs/supabase_users.md.
- Simpan otomatis tidak lagi menimpa proyek cloud bila data di browser kosong atau tinggal kurang dari separuh baris (tombol menampilkan "tidak disimpan — cek data"); Simpan manual meminta konfirmasi dulu.

## v1.9.6 — Impor "SOP TO-BE" sebagai data utama (2026-11-24)
- Saat membuka Master Workbook yang punya sheet SOP dan SOP TO-BE, aplikasi menanyakan sumbernya:
  - "SOP — AS-IS + rancangan TO-BE seperti terakhir disimpan" (perilaku lama), atau
  - "SOP TO-BE — jadikan TO-BE sebagai data utama (AS-IS baru)": nomor runut dari sheet dipakai, baris "Dihapus di TO-BE" diabaikan, rancangan TO-BE lama tidak dimuat lagi (sudah jadi AS-IS), SLA/input-output dan jenis gateway ikut pindah ke nomor baru lewat kolom "No L5 lama".

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
