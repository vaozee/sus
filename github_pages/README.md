# KopSUS App — GitHub Pages Demo

Aplikasi manajemen inventori & keuangan koperasi berbasis web (single-file, standalone).

## Status

✅ Semua 10 halaman berfungsi penuh (bukan dummy): Dashboard, Inventory, Suppliers, Customers,
Purchases (PO + Purchase Details), Sales (SO + Sales Details), Receipts, Payments, Reports
(4 sub-tab), Settings (Dimensions, Company Profile, Display Preferences, Users).

✅ Skema data 1:1 dengan struktur Google Sheets asli (RANGEINVENTORYITEMS, RANGESUPPLIERS,
RANGECUSTOMERS, RANGEPO, RANGEPD, RANGESO, RANGESD, RANGERECEIPTS, RANGEPAYMENTS,
RANGEDIMENSIONS, Users sheet).

✅ CRUD penuh (Tambah/Edit/Hapus) tervalidasi otomatis dengan 58 automated test (headless
browser via jsdom) — 0 gagal, 0 runtime error. Mencakup: validasi role-based access (Admin/
Manager/Staff), proteksi hapus data yang masih punya saldo/stok, auto-fill lokasi dari
master data, perhitungan balance otomatis, multi-item Sales Order builder, filter/search,
pagination, dan logout/login flow.

## Cara Deploy ke GitHub Pages

1. Buat repository baru di GitHub (misal: `kopsus-app`)
2. Upload file `index.html` ini ke repository
3. Buka **Settings → Pages**
4. Pilih **Source: Deploy from a branch**
5. Pilih branch `main`, folder `/ (root)`
6. Klik **Save**
7. Tunggu beberapa menit, lalu akses di: `https://username.github.io/kopsus-app`

## Akun Demo

| Username | Password | Role |
|----------|----------|------|
| admin | admin123 | Admin (akses penuh, termasuk Settings) |
| manager | manager123 | Manager (akses penuh, termasuk Settings) |
| staff | staff123 | Staff (Settings tersembunyi, hanya bisa edit data milik sendiri) |

## Catatan Penting

> ⚠️ Versi GitHub Pages ini adalah **demo dengan data di memori browser** — semua perubahan
> (tambah/edit/hapus) akan **hilang saat halaman di-refresh**. Ini bukan keterbatasan/bug,
> melainkan desain: GitHub Pages hanya hosting file statis, tidak punya database backend.
>
> Untuk versi production dengan data permanen tersimpan di Google Sheets, deploy kode `.GS`
> asli sebagai Web App di Google Apps Script (Extensions → Apps Script → Deploy → New deployment
> → Web app).

