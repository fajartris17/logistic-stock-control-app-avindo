LOGISTIC STOCK CONTROL APP AVINDO - REV04 FINAL

Status:
Aplikasi logistic sudah dilengkapi untuk penggunaan lokal/PWA dengan database Excel otomatis, input barang masuk, barang keluar, stok, lampiran foto/nota/bukti, laporan filter, edit transaksi, log edit, dan koreksi manual.

Fitur utama REV04:
1. Dashboard AVINDO responsif untuk laptop dan HP.
2. Menu Barang Masuk dengan:
   - Code Number
   - Nama Barang
   - Jenis Barang
   - QTY
   - Satuan
   - Harga Satuan
   - Total Harga otomatis
   - Minimum Stok
   - Supplier / Asal Barang
   - Upload Foto Barang
   - Upload Foto Nota
3. Menu Barang Keluar dengan:
   - Validasi stok agar tidak minus.
   - Upload Foto Barang.
   - Upload Bukti Keluar.
4. Database Excel otomatis di:
   /database/logistic_stock_database.xlsx
5. Jika database lama sudah ada, aplikasi menambah sheet/kolom yang kurang tanpa menghapus data lama.
6. Upload file tersimpan di folder:
   /uploads
7. Format upload yang diizinkan:
   JPG, JPEG, PNG, WEBP, PDF.
8. Maksimal ukuran file:
   10 MB per file dan 32 MB per submit.
9. Export Excel dari tombol Download Database Excel.
10. Laporan logistic dengan filter tanggal dan pencarian.
11. Print laporan dari halaman Laporan.
12. Edit transaksi berdasarkan ID transaksi.
13. Semua perubahan edit tersimpan otomatis di sheet LOG_EDIT_DATA.
14. OCR / Nota manual disiapkan di sheet DATABASE_NOTA_OCR.
15. Log koreksi manual disimpan di sheet LOG_KOREKSI_DATA.
16. PWA manifest dan service worker tetap tersedia.

Sheet Excel:
1. DATABASE_BARANG_MASUK
2. DATABASE_BARANG_KELUAR
3. DATABASE_STOK
4. MASTER_BARANG
5. LOG_EDIT_DATA
6. DATABASE_NOTA_OCR
7. LOG_KOREKSI_DATA

Format ID otomatis:
- Barang Masuk  : BM-YYYYMMDD-0001
- Barang Keluar : BK-YYYYMMDD-0001
- Edit Data     : EDIT-YYYYMMDD-0001
- OCR Nota      : OCR-YYYYMMDD-0001
- Koreksi       : KOREKSI-YYYYMMDD-0001

Cara menjalankan di Windows:
1. Pastikan Python sudah terinstal.
2. Buka folder repository.
3. Klik RUN_APP_WINDOWS.bat.
4. Tunggu sampai dependency selesai terinstal.
5. Buka browser ke:
   http://127.0.0.1:5000

Catatan penting:
- Jangan hapus folder /database jika ingin mempertahankan data lama.
- Jangan rename file logistic_stock_database.xlsx.
- Jangan hapus folder /uploads jika ingin lampiran tetap bisa dibuka.
- File Excel database tidak perlu dibuat manual; aplikasi membuatnya otomatis.
- Jika file database hilang, aplikasi akan membuat database baru kosong.
