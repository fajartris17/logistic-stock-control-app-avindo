from flask import Flask, render_template, jsonify, request, send_file
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import re

BASE_DIR = Path(__file__).resolve().parent
DATABASE_DIR = BASE_DIR / 'database'
UPLOAD_DIR = BASE_DIR / 'uploads'
DB_PATH = DATABASE_DIR / 'logistic_stock_database.xlsx'

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

DAY_NAMES = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']

SHEET_HEADERS = {
    'DATABASE_BARANG_MASUK': ['ID Transaksi','Tanggal Input','Jam Input','Code Number','Nama Barang','Jenis Barang','QTY Masuk','Satuan','Total QTY','Harga Satuan','Total Harga','Supplier / Asal Barang','Nama File Foto Barang','Path Foto Barang','Nama File Nota','Path Foto Nota','Status Koreksi','Keterangan','User Input'],
    'DATABASE_BARANG_KELUAR': ['ID Transaksi','Tanggal Input','Jam Input','Code Number','Nama Barang','Jenis Barang','QTY Keluar','Satuan','Total QTY','Harga Satuan','Total Harga','Tujuan / Pengguna Barang','Nama File Foto Barang','Path Foto Barang','Nama File Bukti Keluar','Path Foto Bukti Keluar','Status Koreksi','Keterangan','User Input'],
    'DATABASE_STOK': ['Code Number','Nama Barang','Jenis Barang','Total QTY Masuk','Total QTY Keluar','Stok Akhir','Satuan','Harga Satuan Terakhir','Estimasi Nilai Stok','Minimum Stok','Status Stok','Foto Barang Terakhir','Keterangan'],
    'MASTER_BARANG': ['Code Number','Nama Barang','Jenis Barang','Satuan','Harga Satuan Terakhir','Minimum Stok','Keterangan'],
    'LOG_EDIT_DATA': ['ID Edit','Tanggal Edit','Jam Edit','ID Transaksi','Field yang Diedit','Data Lama','Data Baru','User Edit','Keterangan Edit'],
    'DATABASE_NOTA_OCR': ['ID OCR','Tanggal Upload','Jam Upload','Nama File Nota','Path Foto Nota','Tanggal Nota','Nama Supplier / Toko','Nomor Nota','Nama Barang OCR','QTY OCR','Satuan OCR','Harga Satuan OCR','Total Harga OCR','Total Nota OCR','Status Pembacaan','Keterangan OCR'],
    'LOG_KOREKSI_DATA': ['ID Koreksi','Tanggal Koreksi','Jam Koreksi','ID Transaksi','Nama Barang Input','Nama Barang Nota','QTY Input','QTY Nota','Harga Satuan Input','Harga Satuan Nota','Total Harga Input','Total Harga Nota','Status Koreksi','Tindakan User','Keterangan']
}


def now_date(): return datetime.now().strftime('%Y-%m-%d')
def now_time(): return datetime.now().strftime('%H:%M:%S')


def to_float(value):
    try:
        if value is None or value == '': return 0.0
        return float(str(value).replace('.', '').replace(',', '.')) if isinstance(value, str) and ',' in value else float(value)
    except Exception:
        return 0.0


def rupiah(value): return 'Rp {:,.0f}'.format(to_float(value)).replace(',', '.')


def form_or_json(): return request.form if request.form else (request.get_json(silent=True) or {})
def file_name(field):
    f = request.files.get(field)
    return f.filename if f and f.filename else ''


def style_sheet(ws):
    fill = PatternFill('solid', fgColor='0B1F3A'); font = Font(bold=True, color='FFFFFF')
    thin = Side(style='thin', color='D9E2EC'); border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill = fill; cell.font = font; cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); cell.border = border
    ws.freeze_panes = 'A2'; ws.row_dimensions[1].height = 34
    for idx, cell in enumerate(ws[1], start=1):
        h = str(cell.value or '').lower(); width = 14
        if any(k in h for k in ['nama','keterangan','supplier','tujuan','path','field','foto','nota','bukti']): width = 24
        if 'id' in h or 'code' in h: width = 18
        ws.column_dimensions[get_column_letter(idx)].width = width


def add_table_if_needed(ws):
    table_name = re.sub(r'[^A-Za-z0-9_]', '', ws.title) + 'Table'
    if table_name in ws.tables: return
    try:
        tab = Table(displayName=table_name, ref=f'A1:{get_column_letter(ws.max_column)}{max(ws.max_row, 2)}')
        tab.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        ws.add_table(tab)
    except Exception: pass


def ensure_database():
    DATABASE_DIR.mkdir(parents=True, exist_ok=True); UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(DB_PATH) if DB_PATH.exists() else Workbook()
    if 'Sheet' in wb.sheetnames and len(wb.sheetnames) == 1 and not wb['Sheet']['A1'].value: wb.remove(wb['Sheet'])
    for sheet, headers in SHEET_HEADERS.items():
        if sheet not in wb.sheetnames:
            ws = wb.create_sheet(sheet); ws.append(headers)
        else:
            ws = wb[sheet]; existing = [c.value for c in ws[1]] if ws.max_row >= 1 else []
            if not any(existing): ws.append(headers); existing = headers
            for h in headers:
                if h not in existing: ws.cell(row=1, column=ws.max_column + 1, value=h)
        style_sheet(wb[sheet]); add_table_if_needed(wb[sheet])
    wb.save(DB_PATH)


def read_sheet(sheet):
    ensure_database(); wb = load_workbook(DB_PATH, data_only=True); ws = wb[sheet]; rows = list(ws.iter_rows(values_only=True))
    if not rows: return []
    headers = [str(h or '') for h in rows[0]]; data = []
    for row in rows[1:]:
        if any(row): data.append({headers[i]: row[i] if i < len(row) else None for i in range(len(headers))})
    return data


def append_dict_row(sheet, data):
    ensure_database(); wb = load_workbook(DB_PATH); ws = wb[sheet]; headers = [c.value for c in ws[1]]
    ws.append([data.get(h, '') for h in headers]); style_sheet(ws); add_table_if_needed(ws); wb.save(DB_PATH)


def next_id(sheet, prefix):
    date_key = datetime.now().strftime('%Y%m%d'); pat = re.compile(rf'^{re.escape(prefix)}-{date_key}-(\d{{4}})$'); max_no = 0
    for row in read_sheet(sheet):
        val = str(row.get('ID Transaksi') or row.get('ID Edit') or row.get('ID OCR') or row.get('ID Koreksi') or '')
        m = pat.match(val)
        if m: max_no = max(max_no, int(m.group(1)))
    return f'{prefix}-{date_key}-{max_no + 1:04d}'


def upsert_master_barang(code, nama, jenis, satuan, harga, minimum=0, keterangan=''):
    if not code: return
    wb = load_workbook(DB_PATH); ws = wb['MASTER_BARANG']; headers = [c.value for c in ws[1]]; code_idx = headers.index('Code Number') + 1; found = None
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, code_idx).value or '').strip().upper() == code: found = r; break
    row_data = {'Code Number': code, 'Nama Barang': nama, 'Jenis Barang': jenis, 'Satuan': satuan, 'Harga Satuan Terakhir': harga, 'Minimum Stok': minimum, 'Keterangan': keterangan}
    if found:
        for h, v in row_data.items():
            if v not in [None, '']: ws.cell(found, headers.index(h) + 1, value=v)
    else: ws.append([row_data.get(h, '') for h in headers])
    style_sheet(ws); add_table_if_needed(ws); wb.save(DB_PATH)


def rebuild_stock():
    ensure_database(); masuk = read_sheet('DATABASE_BARANG_MASUK'); keluar = read_sheet('DATABASE_BARANG_KELUAR'); stock = {}
    def item(code, row):
        return stock.setdefault(code, {'Code Number': code,'Nama Barang': row.get('Nama Barang') or '','Jenis Barang': row.get('Jenis Barang') or '','Total QTY Masuk': 0.0,'Total QTY Keluar': 0.0,'Stok Akhir': 0.0,'Satuan': row.get('Satuan') or '','Harga Satuan Terakhir': to_float(row.get('Harga Satuan')),'Estimasi Nilai Stok': 0.0,'Minimum Stok': 0,'Status Stok': 'TERSEDIA','Foto Barang Terakhir': row.get('Path Foto Barang') or row.get('Nama File Foto Barang') or '','Keterangan': ''})
    for r in masuk:
        code = str(r.get('Code Number') or '').strip().upper()
        if code:
            it = item(code, r); it['Total QTY Masuk'] += to_float(r.get('Total QTY') or r.get('QTY Masuk')); it['Harga Satuan Terakhir'] = to_float(r.get('Harga Satuan')) or it['Harga Satuan Terakhir']; it['Foto Barang Terakhir'] = r.get('Path Foto Barang') or r.get('Nama File Foto Barang') or it['Foto Barang Terakhir']
    for r in keluar:
        code = str(r.get('Code Number') or '').strip().upper()
        if code: item(code, r)['Total QTY Keluar'] += to_float(r.get('Total QTY') or r.get('QTY Keluar'))
    for it in stock.values():
        it['Stok Akhir'] = it['Total QTY Masuk'] - it['Total QTY Keluar']; it['Estimasi Nilai Stok'] = it['Stok Akhir'] * it['Harga Satuan Terakhir']; it['Status Stok'] = 'STOK HABIS' if it['Stok Akhir'] <= 0 else 'TERSEDIA'
    wb = load_workbook(DB_PATH); ws = wb['DATABASE_STOK']
    if ws.max_row > 1: ws.delete_rows(2, ws.max_row - 1)
    headers = [c.value for c in ws[1]]
    for _, it in sorted(stock.items()): ws.append([it.get(h, '') for h in headers])
    style_sheet(ws); add_table_if_needed(ws); wb.save(DB_PATH)


def build_laporan():
    rows = []
    for r in read_sheet('DATABASE_BARANG_MASUK'):
        rows.append({'Tanggal': r.get('Tanggal Input'),'Jam': r.get('Jam Input'),'Jenis Transaksi': 'BARANG MASUK','ID Transaksi': r.get('ID Transaksi'),'Code Number': r.get('Code Number'),'Nama Barang': r.get('Nama Barang'),'Qty': r.get('Total QTY') or r.get('QTY Masuk'),'Satuan': r.get('Satuan'),'Nilai': r.get('Total Harga'),'Pihak': r.get('Supplier / Asal Barang'),'Foto': r.get('Path Foto Barang') or r.get('Nama File Foto Barang'),'Bukti': r.get('Path Foto Nota') or r.get('Nama File Nota')})
    for r in read_sheet('DATABASE_BARANG_KELUAR'):
        rows.append({'Tanggal': r.get('Tanggal Input'),'Jam': r.get('Jam Input'),'Jenis Transaksi': 'BARANG KELUAR','ID Transaksi': r.get('ID Transaksi'),'Code Number': r.get('Code Number'),'Nama Barang': r.get('Nama Barang'),'Qty': r.get('Total QTY') or r.get('QTY Keluar'),'Satuan': r.get('Satuan'),'Nilai': r.get('Total Harga'),'Pihak': r.get('Tujuan / Pengguna Barang'),'Foto': r.get('Path Foto Barang') or r.get('Nama File Foto Barang'),'Bukti': r.get('Path Foto Bukti Keluar') or r.get('Nama File Bukti Keluar')})
    return sorted(rows, key=lambda x: str(x.get('Tanggal') or '') + ' ' + str(x.get('Jam') or ''))

@app.route('/')
def index(): ensure_database(); return render_template('index.html')

@app.route('/api/dashboard')
def dashboard():
    rebuild_stock(); today = datetime.now(); masuk = read_sheet('DATABASE_BARANG_MASUK'); keluar = read_sheet('DATABASE_BARANG_KELUAR'); stok = read_sheet('DATABASE_STOK')
    nilai_masuk = sum(to_float(r.get('Total Harga')) for r in masuk); nilai_keluar = sum(to_float(r.get('Total Harga')) for r in keluar)
    return jsonify({'tanggal': today.strftime('%d-%m-%Y'),'hari': DAY_NAMES[today.weekday()],'total_jenis_barang': len(stok),'total_barang_masuk': sum(to_float(r.get('Total QTY') or r.get('QTY Masuk')) for r in masuk),'total_barang_keluar': sum(to_float(r.get('Total QTY') or r.get('QTY Keluar')) for r in keluar),'total_stok_tersedia': sum(to_float(r.get('Stok Akhir')) for r in stok),'total_nilai_masuk_fmt': rupiah(nilai_masuk),'total_nilai_keluar_fmt': rupiah(nilai_keluar),'recent_stock': stok[-6:]})

@app.route('/api/stok')
def api_stok(): rebuild_stock(); return jsonify(read_sheet('DATABASE_STOK'))

@app.route('/api/barang-masuk', methods=['POST'])
def api_barang_masuk():
    p = form_or_json(); qty = to_float(p.get('qty')); harga = to_float(p.get('harga_satuan')); total = qty * harga; code = str(p.get('kode_barang') or '').strip().upper(); tx = next_id('DATABASE_BARANG_MASUK', 'BM')
    foto = file_name('foto_barang'); nota = file_name('foto_nota')
    row = {'ID Transaksi': tx,'Tanggal Input': p.get('tanggal') or now_date(),'Jam Input': now_time(),'Code Number': code,'Nama Barang': p.get('nama_barang') or '','Jenis Barang': p.get('kategori') or p.get('jenis_barang') or '','QTY Masuk': qty,'Satuan': p.get('satuan') or '','Total QTY': qty,'Harga Satuan': harga,'Total Harga': total,'Supplier / Asal Barang': p.get('supplier') or '','Nama File Foto Barang': foto,'Path Foto Barang': '','Nama File Nota': nota,'Path Foto Nota': '','Status Koreksi': 'BELUM DICEK','Keterangan': p.get('keterangan') or '','User Input': p.get('user_input') or 'Admin'}
    append_dict_row('DATABASE_BARANG_MASUK', row); upsert_master_barang(code, row['Nama Barang'], row['Jenis Barang'], row['Satuan'], harga); rebuild_stock()
    return jsonify({'ok': True, 'message': f'Barang masuk tersimpan dengan ID {tx}', 'id': tx})

@app.route('/api/barang-keluar', methods=['POST'])
def api_barang_keluar():
    p = form_or_json(); qty = to_float(p.get('qty')); harga = to_float(p.get('harga_satuan')); total = qty * harga; code = str(p.get('kode_barang') or '').strip().upper(); tx = next_id('DATABASE_BARANG_KELUAR', 'BK')
    foto = file_name('foto_barang'); bukti = file_name('foto_bukti_keluar')
    row = {'ID Transaksi': tx,'Tanggal Input': p.get('tanggal') or now_date(),'Jam Input': now_time(),'Code Number': code,'Nama Barang': p.get('nama_barang') or '','Jenis Barang': p.get('kategori') or p.get('jenis_barang') or '','QTY Keluar': qty,'Satuan': p.get('satuan') or '','Total QTY': qty,'Harga Satuan': harga,'Total Harga': total,'Tujuan / Pengguna Barang': p.get('penerima') or p.get('tujuan') or '','Nama File Foto Barang': foto,'Path Foto Barang': '','Nama File Bukti Keluar': bukti,'Path Foto Bukti Keluar': '','Status Koreksi': 'BELUM DICEK','Keterangan': p.get('keterangan') or '','User Input': p.get('user_input') or 'Admin'}
    append_dict_row('DATABASE_BARANG_KELUAR', row); upsert_master_barang(code, row['Nama Barang'], row['Jenis Barang'], row['Satuan'], harga); rebuild_stock()
    return jsonify({'ok': True, 'message': f'Barang keluar tersimpan dengan ID {tx}', 'id': tx})

@app.route('/api/laporan')
def api_laporan(): rebuild_stock(); return jsonify(build_laporan())

@app.route('/api/database-info')
def api_database_info():
    ensure_database(); wb = load_workbook(DB_PATH, read_only=True)
    return jsonify({'path': str(DB_PATH).replace(str(BASE_DIR), ''),'exists': DB_PATH.exists(),'sheet_count': len(wb.sheetnames),'sheets': wb.sheetnames,'file_name': DB_PATH.name,'file_size_kb': round(DB_PATH.stat().st_size / 1024, 2)})

@app.route('/api/export-excel')
def api_export_excel(): ensure_database(); rebuild_stock(); return send_file(DB_PATH, as_attachment=True, download_name='logistic_stock_database.xlsx')

if __name__ == '__main__': ensure_database(); app.run(host='127.0.0.1', port=5000, debug=True)
