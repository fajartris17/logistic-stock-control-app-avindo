# LOGISTIC STOCK CONTROL APP AVINDO

Repository untuk aplikasi Logistic Stock Control App AVINDO.

Branch pengembangan awal: `rev02-database-excel`.

## Perbaikan Print APK_DAILY_REPORT

Status: sudah ditambahkan file CSS perbaikan print layout.

File:

`fixes/apk_daily_report_print_fix.css`

Target perbaikan:

- Tampilan Form, Print Preview, PDF, dan JPG dibuat sama.
- Area laporan dikunci untuk A4 landscape.
- Print area difokuskan hanya ke form laporan utama.
- Tombol dan kontrol aplikasi tidak ikut tercetak.
- Foto dokumentasi tetap berada di dalam frame.
- Tabel dibuat stabil agar tidak melebar atau bergeser saat print.

Cara pemasangan pada file HTML utama:

Tambahkan CSS ini pada bagian head HTML utama APK Daily Report:

`<link rel="stylesheet" href="fixes/apk_daily_report_print_fix.css">`

Pastikan area laporan utama memakai salah satu penanda berikut:

- `printArea`
- `formDailyReport`
- `print-area`
- `daily-report-form`

Catatan penting: file HTML utama dari screenshot masih berada di lokal komputer dan belum tersedia di repository ini. Setelah file HTML utama di-upload ke GitHub, CSS ini bisa langsung dipasang ke file tersebut.
