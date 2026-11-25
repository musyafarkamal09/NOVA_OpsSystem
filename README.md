# NOVA — Operational Management System

Aplikasi web (PWA) untuk pencatatan pekerjaan, meteran, dan pembelian Tim Teknikal **Sukaramai Trade Center XXI**.
*NOVA — An Innovation by MK.*

## Modul utama

- **DCP** dan **Pelaporan DCP** (banyak item per laporan, foto, PDF)
- **Daily Film Testing** (beberapa studio dalam satu laporan)
- **Data Penonton** per show
- **Monitoring AC** (termasuk grafik suhu)
- **Pencatatan Meteran** (listrik dan air, termasuk input terlambat/backdated)
- Draft laporan otomatis, ekspor PDF dan Excel, notifikasi push

## Teknologi

| Bagian | Yang dipakai |
|---|---|
| Frontend | HTML, CSS, dan JavaScript statis (tanpa bundler), PWA dengan service worker |
| Gaya | Tailwind CSS v4 (hasil build sudah ikut di repo) + lapisan CSS desain NOVA |
| Backend | Google Apps Script (Web App) dengan data di Google Sheets dan Drive |
| Pustaka via CDN | jsPDF, jsPDF-AutoTable, SheetJS, ExcelJS, PDF.js, OneSignal |

## Struktur repo

```
.
├── index.html            # halaman utama (harus di akar)
├── manifest.json         # manifest PWA (harus di akar)
├── sw.js                 # service worker (harus di akar agar scope mencakup seluruh situs)
├── assets/
│   ├── css/              # tailwind.css, legacy-modules.css, oos.css, figma.css, nova*.css
│   ├── js/               # app.js dan modul (modul-dcp, pelaporan, penonton, pdfkit, ...)
│   ├── img/              # foto proyektor/auditorium
│   │   └── logo/         # logo NOVA dan XXI
│   └── icons/            # ikon PWA
├── backend/              # Google Apps Script: Code.gs, dcp.gs, appsscript.json
├── src/                  # sumber CSS Tailwind (tailwind.input.css)
├── tools/                # build-css.mjs
├── tests/                # tes Node (backend/logika) dan tes browser (Playwright)
├── docs/
│   ├── changelog/        # catatan perubahan per versi
│   ├── design/           # DESIGN.md dan catatan redesign
│   └── audit/            # audit, verifikasi, hasil tes
└── archive/              # snapshot frontend lama (tidak dipakai, aman dihapus)
```

## Menjalankan di lokal

```bash
python3 -m http.server 8765
# buka http://localhost:8765/index.html
```

URL backend diatur lewat konstanta `API_URL` di `assets/js/app.js`.

## Backend (Apps Script)

1. Tempel isi `backend/Code.gs` dan `backend/dcp.gs` ke proyek Apps Script (`appsscript.json` mengatur zona waktu dan akses Web App).
2. **Terapkan → Kelola penerapan → ikon pensil → Versi baru → Terapkan**. URL Web App tetap sama.

## Build ulang CSS (hanya bila mengubah `src/tailwind.input.css` atau kelas Tailwind di HTML/JS)

```bash
npm i -D tailwindcss@4
node tools/build-css.mjs      # menulis assets/css/tailwind.css
```

## Pengujian

```bash
# tes Node (backend dan logika)
for t in tests/test_*.js; do node "$t"; done

# tes browser: butuh Playwright dan server statis di :8765
sh tests/run-all.sh
```

## Merilis versi baru

- Naikkan `CACHE_VERSION` di `sw.js`.
- Naikkan parameter `?v=` pada berkas yang berubah di `index.html`.
- Setelah deploy, hard-refresh sekali agar service worker memuat cache baru.

## Dokumentasi

Lihat [`docs/`](docs/README.md): catatan perubahan, sistem desain, dan laporan audit.
