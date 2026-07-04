LOGISTIC STOCK CONTROL APP AVINDO - REV02 DATABASE EXCEL OTOMATIS

Fokus REV02:
1. Sistem otomatis membuat database Excel jika belum ada.
2. File database tersimpan di:
   /database/logistic_stock_database.xlsx
3. Jika file database sudah ada, aplikasi membaca file lama dan menambahkan sheet/kolom yang kurang tanpa menghapus data.
4. Semua transaksi barang masuk/keluar memakai ID otomatis berurutan per tanggal.
5. Tombol Export Excel mendownload database aktif.

Sheet Excel yang dibuat otomatis:
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

Cara menjalankan:
1. Extract ZIP.
2. Buka folder LOGISTIC_STOCK_CONTROL_APP_AVINDO_REV02.
3. Klik RUN_APP_WINDOWS.bat.
4. Buka browser ke http://127.0.0.1:5000.

Catatan penting:
- Jangan hapus folder /database.
- Jangan rename file logistic_stock_database.xlsx jika ingin aplikasi membaca database yang sama.
- Jika database terhapus, aplikasi akan membuat database baru kosong otomatis.
- Data tidak hilang saat aplikasi ditutup karena tersimpan di file Excel lokal.

Tahap berikutnya yang bisa dikembangkan:
- Upload foto barang dan nota.
- OCR nota/kwitansi otomatis.
- Sistem koreksi selisih antara input dan nota.
- Edit data dengan log edit otomatis.
- Laporan filter tanggal, proyek, supplier, dan export PDF.
