# LOGISTIC STOCK CONTROL APP AVINDO

Repository untuk aplikasi Logistic Stock Control App AVINDO.

Branch pengembangan awal: rev02-database-excel.

## Perbaikan Print APK_DAILY_REPORT

Status: sudah ditambahkan file CSS perbaikan print layout.

File:
fixes/apk_daily_report_print_fix.css

Target perbaikan:

- Tampilan Form, Print Preview, PDF, dan JPG dibuat sama.
- Area laporan dikunci untuk A4 landscape.
- Print area difokuskan hanya ke form laporan utama.
- Tombol dan kontrol aplikasi tidak ikut tercetak.
- Foto dokumentasi tetap berada di dalam frame.
- Tabel dibuat stabil agar tidak melebar atau bergeser saat print.

## Langkah Kerja Lanjutan

1. Upload file HTML utama APK_DAILY_REPORT dari laptop ke repository.
2. Setelah file HTML tersedia di repository, pasang CSS print fix ke bagian head HTML utama.
3. Area form utama harus memakai penanda printArea atau formDailyReport.
4. Frame foto dokumentasi harus memakai penanda photo-frame atau foto-dokumentasi.
5. Lakukan test Print Preview di browser.
6. Setting printer harus A4 landscape.
7. Gunakan scale fit to page atau fit to printable area.
8. Pastikan hasil preview tidak lagi menjadi 3 lembar.
9. Pastikan foto dokumentasi tidak keluar frame.
10. Pastikan PDF dan JPG sama dengan tampilan form.

## Catatan Penting

File HTML utama dari screenshot masih berada di lokal komputer dan belum tersedia di repository ini. Karena itu perbaikan CSS sudah disiapkan, README sudah diperbarui, dan Issue sudah diberi status, tetapi pemasangan langsung ke HTML utama belum bisa dilakukan sampai file tersebut diupload.

## Status Issue

Issue pembahasan: Pembahasan Perbaikan APK_DAILY_REPORT.
Prioritas: High Priority / Wajib Dicek.
