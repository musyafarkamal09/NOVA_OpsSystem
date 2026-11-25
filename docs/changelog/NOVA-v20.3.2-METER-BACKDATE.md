# NOVA v20.3.2 — Pencatatan Meteran: input terlambat (backdated)

## File yang diubah
| File | Perubahan |
|---|---|
| `Backend/code.gs.txt` (Code.gs) | `mCmp_` (urutan tunggal), `meterUse_` (tanpa bagi rata), `meterGaps_` (baru), `meterDiff_`, `refreshMeterMirrors_`, `meter_prev` (+`next`), `meter_stats` (+`gaps`) |
| `app.js` | peringatan tanggal terlewat / pencatatan berikutnya di formulir, panel "Pencatatan belum tersedia", penanda anomali di detail, panel statistik langsung segar setelah simpan/edit/arsip |
| `index.html`, `sw.js` | `app.js?v=185`, cache `nova-v20.3.2-meter-backdate` agar perangkat memuat versi baru |
| `tests/test_meter_backdate.js` (baru), `tests/browser/meter-backdate.js` (baru), `tests/test_backend.js` | tes skenario + tes lama yang mengharapkan bagi-rata diperbarui |

Tidak diubah: desain, navigasi, modul lain, struktur sheet, ID laporan, foto/Drive, data yang sudah ada.

## Penyebab bug
1. **Bagi rata ke tanggal kosong** (`meterUse_`): selisih 4→6 dipecah menjadi angka "fiktif" untuk tanggal 5 di grafik, tabel, dan statistik. Setelah tanggal 5 diisi, angka itu baru berubah.
2. **Selisih laporan sesudahnya tidak tersimpan** (`refreshMeterMirrors_`): laporan tanggal 6 hanya diperbarui di sheet `METER_*`, tidak di data `REPORTS`. Data tersimpan tetap memuat selisih lama (30, bukan 20).
3. **Dua urutan berbeda**: `meterRecs_` memakai tanggal+jam+id, `meterDiff_` memakai tanggal+jam+createdAt. Dua pencatatan dengan tanggal dan jam identik saling mengabaikan, sehingga yang kedua tidak dibandingkan dengan yang pertama.
4. **Tidak ada status tanggal kosong**, dan formulir tidak memberi tahu bahwa input susulan mengubah pencatatan berikutnya.
5. **Penyegaran berat**: setiap simpan menulis ulang semua laporan sesudah tanggal terkait ke Spreadsheet (4+ panggilan sheet per baris).

## Perilaku sekarang
- Urutan selalu tanggal + jam meter yang dicatat (`mCmp_`); `createdAt`/id hanya pemecah seri.
- Pemakaian = angka sekarang − pencatatan aktual sebelumnya, per lokasi dan per jenis (LWBP, WBP, Genset, Air). Total listrik = LWBP + WBP; Genset terpisah.
- Pemakaian dicatat pada tanggal pencatatan aktual yang lebih baru. Tanggal tanpa pencatatan tidak diberi angka; `meter_stats.gaps` mendaftarkannya (dari pencatatan pertama sampai kemarin WIB, per lokasi, listrik dan air) dan panel menampilkannya.
- Angka turun (reset/ganti meter): tidak negatif, tidak diakumulasi, ditandai ⚠ di detail laporan dan panel pemeriksaan data.
- Simpan/edit/arsip: hanya laporan yang selisihnya benar-benar berubah ditulis ulang (data + sheet), biasanya laporan itu dan satu sesudahnya.

## Hasil uji
- `node tests/test_meter_backdate.js`: **24/24 lulus** (pada backend asli 8 gagal: tanggal fiktif, selisih tersimpan basi, urutan jam identik).
- `node tests/browser/meter-backdate.js`: **8/8 lulus** (desktop dan ponsel 390 px, tanpa overflow).
- Tes lain tidak berubah: audience 36/36, auth 14/14, ft_laporan 46/46, photoflow 47/47, dll.; browser `run.js` 18 modul tanpa masalah di desktop dan ponsel; `v16` 14/14.
- Gagal yang sudah ada sebelum perubahan ini (sama persis pada kode asli): `test_dcp` 3, `test_dcp_batch` 4, `test_ft_dok` 1, `test_backend.js` berhenti di bagian Data Penonton, browser `v14`.

## Cara menerapkan
1. Ganti isi Code.gs dengan `Backend/code.gs.txt`. Deploy > Kelola deployment > Edit > **Versi baru**.
2. Unggah `app.js`, `index.html`, `sw.js` ke GitHub. Muat ulang aplikasi (Ctrl+F5) agar cache lama diganti.
3. Tidak perlu migrasi: selisih lama dikoreksi otomatis saat ada simpan/edit/arsip di lokasi terkait.

## Yang masih perlu diperhatikan
- Backend diuji pada tiruan Spreadsheet dalam memori dan UI pada backend tiruan. Belum dijalankan di Apps Script/Spreadsheet asli.
- Grafik harian tetap menggambar nol di tanggal kosong (desain tidak diubah); peringatan tanggal kosong tampil di atasnya. Selama tanggal 5 kosong, pemakaian 4→6 muncul di tanggal 6.
- Total bulanan bisa bergeser bila selisih melewati pergantian bulan: dulu dibagi ke hari-hari sebelumnya, sekarang masuk ke bulan pencatatan yang lebih baru. Total keseluruhan tetap sama.
- Peringatan tanggal kosong mengasumsikan pencatatan setiap hari. Lokasi yang libur tetap akan muncul sebagai kosong. Daftar dibatasi 31 tanggal per lokasi dan jenis.
- Setelah reset/ganti meter, pemakaian pada interval itu tidak diperkirakan. Angka baru menjadi pembanding berikutnya.
- Jam identik di lokasi yang sama diurutkan menurut waktu input (tidak ada cara lain yang pasti).
- Batas edit 30 menit untuk non-Admin tetap berlaku untuk mengedit laporan lama. Membuat laporan susulan tidak dibatasi (tanggal ≥ 2020-01-01 dan tidak lebih dari 36 jam ke depan).
- Penulisan bersamaan oleh dua pengguna tidak saya audit.
- Baris sheet lama yang selisihnya sudah benar tidak ditulis ulang. Gunakan `rebuildMirrors_` bila ingin menyegarkan seluruh sheet sekaligus.
