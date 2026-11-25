# NOVA v20.3.0 — Penyesuaian tampilan mengikuti desain Figma

Hanya frontend (tampilan). Fitur, backend (`Backend/*.gs.txt`), ID, dan kelas yang dipakai JS tidak diubah.

- Baru: `nova-figma.css`, dimuat paling akhir (setelah `nova-ds.css`). Menimpa token warna dan komponen.
- Palet: sidebar navy `#071b3c`, aksen biru `#0879f9`, kanvas `#f5f7fb`, label grup menu cyan, menu aktif biru tua tanpa cincin emas.
- Kartu putih bergaris tipis tanpa bayangan; tombol utama biru, sekunder putih bergaris; kolom isian fokus biru lembut; kepala tabel abu muda.
- Mode gelap ikut palet navy.
- `index.html`: tautan CSS baru + theme-color; `manifest.json`: theme/background color; `sw.js`: CSS baru masuk cache, versi `nova-v20.3.0-figma`.

Penerapan: unggah isi ZIP, lalu Ctrl+Shift+R (atau tutup-buka PWA agar cache baru terpakai).
Untuk kembali ke tampilan lama, hapus tautan `nova-figma.css` di `index.html`.

## v20.3.1 — hapus Cafe, perbaikan pencarian dashboard
- Dihapus: menu/halaman Cafe, `cafe.js`, `Backend/cafe.gs.txt`, role Cashier Cafe (opsi akun, hak akses, gerbang backend), aksi `cafe_*`, pintasan "Omzet & Bill Cafe" di tombol Buat, ikon dan deskripsi menu. Label "Outlet (Cafe)" di form akun menjadi "Outlet".
- Data lama aman: sheet `CAFE_SALES` dan `CAFE_TARGET` di Spreadsheet tidak dihapus dan tidak ditandai "tak terpakai"; hapus manual bila sudah tidak perlu.
- Akun lama ber-role CASHIER_CAFE tidak punya menu lagi: ubah role-nya lewat Manajemen Pengguna.
- Pencarian dashboard (`nova-figma.css` bagian 12): kotak seragam radius 10, fokus biru, tombol hapus dan kbd "/" rapi, panel hasil (header, chip kategori, baris hasil, sorotan kata) kini bertata gaya, 16px di HP agar tidak zoom iOS.
- Penerapan: tempel ulang `Backend/code.gs.txt` ke Apps Script (hapus berkas cafe.gs lama dari proyek Apps Script) → Terapkan > Versi baru; unggah isi ZIP; Ctrl+Shift+R.
