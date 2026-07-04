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

DAY_NAMES = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']

SHEET_HEADERS = {
    'DATABASE_BARANG_MASUK': [
        'ID Transaksi', 'Tanggal Input', 'Jam Input', 'Code Number', 'Nama Barang', 'Jenis Barang',
        'QTY Masuk', 'Satuan', 'Total QTY', 'Harga Satuan', 'Total Harga', 'Supplier / Asal Barang',
        'Nama File Foto Barang', 'Path Foto Barang', 'Nama File Nota', 'Path Foto Nota',
        'Status Koreksi', 'Keterangan', 'User Input'
    ],
    'DATABASE_BARANG_KELUAR': [
        'ID Transaksi', 'Tanggal Input', 'Jam Input', 'Code Number', 'Nama Barang', 'Jenis Barang',
        'QTY Keluar', 'Satuan', 'Total QTY', 'Harga Satuan', 'Total Harga', 'Tujuan / Pengguna Barang',
        'Nama File Foto Barang', 'Path Foto Barang', 'Nama File Bukti Keluar', 'Path Foto Bukti Keluar',
        'Status Koreksi', 'Keterangan', 'User Input'
    ],
    'DATABASE_STOK': [
        'Code Number', 'Nama Barang', 'Jenis Barang', 'Total QTY Masuk', 'Total QTY Keluar', 'Stok Akhir',
        'Satuan', 'Harga Satuan Terakhir', 'Estimasi Nilai Stok', 'Minimum Stok', 'Status Stok',
        'Foto Barang Terakhir', 'Keterangan'
    ],
    'MASTER_BARANG': [
        'Code Number', 'Nama Barang', 'Jenis Barang', 'Satuan', 'Harga Satuan Terakhir', 'Minimum Stok', 'Keterangan'
    ],
    'LOG_EDIT_DATA': [
        'ID Edit', 'Tanggal Edit', 'Jam Edit', 'ID Transaksi', 'Field yang Diedit', 'Data Lama', 'Data Baru',
        'User Edit', 'Keterangan Edit'
    ],
    'DATABASE_NOTA_OCR': [
        'ID OCR', 'Tanggal Upload', 'Jam Upload', 'Nama File Nota', 'Path Foto Nota', 'Tanggal Nota',
        'Nama Supplier / Toko', 'Nomor Nota', 'Nama Barang OCR', 'QTY OCR', 'Satuan OCR',
        'Harga Satuan OCR', 'Total Harga OCR', 'Total Nota OCR', 'Status Pembacaan', 'Keterangan OCR'
    ],
    'LOG_KOREKSI_DATA': [
        'ID Koreksi', 'Tanggal Koreksi', 'Jam Koreksi', 'ID Transaksi', 'Nama Barang Input', 'Nama Barang Nota',
        'QTY Input', 'QTY Nota', 'Harga Satuan Input', 'Harga Satuan Nota', 'Total Harga Input',
        'Total Harga Nota', 'Status Koreksi', 'Tindakan User', 'Keterangan'
    ]
}

ID_PREFIX_BY_SHEET = {
    'DATABASE_BARANG_MASUK': 'BM',
    'DATABASE_BARANG_KELUAR': 'BK',
    'LOG_EDIT_DATA': 'EDIT',
    'DATABASE_NOTA_OCR': 'OCR',
    'LOG_KOREKSI_DATA': 'KOREKSI'
}


def now_date():
    return datetime.now().strftime('%Y-%m-%d')


def now_time():
    return datetime.now().strftime('%H:%M:%S')


def to_float(value):
    try:
        if value is None or value == '':
            return 0.0
        return float(str(value).replace('.', '').replace(',', '.')) if isinstance(value, str) and ',' in value else float(value)
    except Exception:
        return 0.0


def rupiah(value):
    return 'Rp {:,.0f}'.format(to_float(value)).replace(',', '.')


def style_sheet(ws):
    header_fill = PatternFill('solid', fgColor='0B1F3A')
    header_font = Font(bold=True, color='FFFFFF')
    thin = Side(style='thin', color='D9E2EC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    ws.freeze_panes = 'A2'
    ws.row_dimensions[1].height = 34
    for idx, cell in enumerate(ws[1], start=1):
        header = str(cell.value or '')
        width = 14
        h = header.lower()
        if any(k in h for k in ['nama', 'keterangan', 'supplier', 'tujuan', 'path', 'field']):
            width = 24
        if 'id' in h or 'code' in h:
            width = 18
        ws.column_dimensions[get_column_letter(idx)].width = width


def add_table_if_needed(ws):
    table_name = re.sub(r'[^A-Za-z0-9_]', '', ws.title) + 'Table'
    if table_name in ws.tables:
        return
    end_col = get_column_letter(ws.max_column)
    end_row = max(ws.max_row, 2)
    ref = f'A1:{end_col}{end_row}'
    try:
        tab = Table(displayName=table_name, ref=ref)
        style = TableStyleInfo(name='TableStyleMedium2', showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        ws.add_table(tab)
    except Exception:
        pass


def ensure_database():
    """Membuat database Excel jika belum ada, dan menambah sheet/kolom yang hilang tanpa menghapus data lama."""
    DATABASE_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

    if DB_PATH.exists():
        wb = load_workbook(DB_PATH)
    else:
        wb = Workbook()
        default_ws = wb.active
        wb.remove(default_ws)

    changed = False
    for sheet_name, headers in SHEET_HEADERS.items():
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
            ws.append(headers)
            changed = True
        else:
            ws = wb[sheet_name]
            existing_headers = [c.value for c in ws[1]] if ws.max_row >= 1 else []
            if not any(existing_headers):
                ws.append(headers)
                existing_headers = headers
                changed = True
            for header in headers:
                if header not in existing_headers:
                    ws.cell(row=1, column=ws.max_column + 1, value=header)
                    changed = True
        style_sheet(ws)
        add_table_if_needed(ws)

    if changed or not DB_PATH.exists():
        wb.save(DB_PATH)
    else:
        # Tetap simpan style bila file lama belum rapi.
        wb.save(DB_PATH)


def read_sheet(sheet_name):
    ensure_database()
    wb = load_workbook(DB_PATH, data_only=True)
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h or '') for h in rows[0]]
    data = []
    for row in rows[1:]:
        if not any(row):
            continue
        data.append({headers[i]: row[i] if i < len(row) else None for i in range(len(headers))})
    return data


def append_dict_row(sheet_name, data):
    ensure_database()
    wb = load_workbook(DB_PATH)
    ws = wb[sheet_name]
    headers = [cell.value for cell in ws[1]]
    ws.append([data.get(h, '') for h in headers])
    style_sheet(ws)
    add_table_if_needed(ws)
    wb.save(DB_PATH)


def next_id(sheet_name, prefix=None):
    ensure_database()
    prefix = prefix or ID_PREFIX_BY_SHEET.get(sheet_name, 'ID')
    date_key = datetime.now().strftime('%Y%m%d')
    pattern = re.compile(rf'^{re.escape(prefix)}-{date_key}-(\d{{4}})$')
    rows = read_sheet(sheet_name)
    max_no = 0
    for row in rows:
        value = str(row.get('ID Transaksi') or row.get('ID Edit') or row.get('ID OCR') or row.get('ID Koreksi') or '')
        match = pattern.match(value)
        if match:
            max_no = max(max_no, int(match.group(1)))
    return f'{prefix}-{date_key}-{max_no + 1:04d}'


def upsert_master_barang(code, nama, jenis, satuan, harga, minimum=0, keterangan=''):
    if not code:
        return
    ensure_database()
    wb = load_workbook(DB_PATH)
    ws = wb['MASTER_BARANG']
    headers = [c.value for c in ws[1]]
    code_idx = headers.index('Code Number') + 1
    found_row = None
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, code_idx).value or '').strip().upper() == code:
            found_row = r
            break
    row_data = {
        'Code Number': code,
        'Nama Barang': nama,
        'Jenis Barang': jenis,
        'Satuan': satuan,
        'Harga Satuan Terakhir': harga,
        'Minimum Stok': minimum,
        'Keterangan': keterangan
    }
    if found_row:
        for h, v in row_data.items():
            if v not in [None, '']:
                ws.cell(found_row, headers.index(h) + 1, value=v)
    else:
        ws.append([row_data.get(h, '') for h in headers])
    style_sheet(ws)
    add_table_if_needed(ws)
    wb.save(DB_PATH)


def get_master_map():
    return {str(r.get('Code Number') or '').strip().upper(): r for r in read_sheet('MASTER_BARANG')}


def rebuild_stock():
    ensure_database()
    masuk = read_sheet('DATABASE_BARANG_MASUK')
    keluar = read_sheet('DATABASE_BARANG_KELUAR')
    master = get_master_map()
    stock = {}

    def get_item(code, fallback=None):
        fallback = fallback or {}
        m = master.get(code, {})
        return stock.setdefault(code, {
            'Code Number': code,
            'Nama Barang': m.get('Nama Barang') or fallback.get('Nama Barang') or '',
            'Jenis Barang': m.get('Jenis Barang') or fallback.get('Jenis Barang') or '',
            'Total QTY Masuk': 0.0,
            'Total QTY Keluar': 0.0,
            'Stok Akhir': 0.0,
            'Satuan': m.get('Satuan') or fallback.get('Satuan') or '',
            'Harga Satuan Terakhir': to_float(m.get('Harga Satuan Terakhir') or fallback.get('Harga Satuan') or 0),
            'Estimasi Nilai Stok': 0.0,
            'Minimum Stok': to_float(m.get('Minimum Stok') or 0),
            'Status Stok': 'TERSEDIA',
            'Foto Barang Terakhir': '',
            'Keterangan': m.get('Keterangan') or ''
        })

    for row in masuk:
        code = str(row.get('Code Number') or '').strip().upper()
        if not code:
            continue
        item = get_item(code, row)
        qty = to_float(row.get('Total QTY') or row.get('QTY Masuk'))
        harga = to_float(row.get('Harga Satuan'))
        item['Total QTY Masuk'] += qty
        if harga:
            item['Harga Satuan Terakhir'] = harga
        item['Nama Barang'] = row.get('Nama Barang') or item['Nama Barang']
        item['Jenis Barang'] = row.get('Jenis Barang') or item['Jenis Barang']
        item['Satuan'] = row.get('Satuan') or item['Satuan']
        item['Foto Barang Terakhir'] = row.get('Path Foto Barang') or row.get('Nama File Foto Barang') or item['Foto Barang Terakhir']

    for row in keluar:
        code = str(row.get('Code Number') or '').strip().upper()
        if not code:
            continue
        item = get_item(code, row)
        qty = to_float(row.get('Total QTY') or row.get('QTY Keluar'))
        item['Total QTY Keluar'] += qty
        item['Nama Barang'] = row.get('Nama Barang') or item['Nama Barang']
        item['Jenis Barang'] = row.get('Jenis Barang') or item['Jenis Barang']
        item['Satuan'] = row.get('Satuan') or item['Satuan']

    for item in stock.values():
        item['Stok Akhir'] = item['Total QTY Masuk'] - item['Total QTY Keluar']
        item['Estimasi Nilai Stok'] = item['Stok Akhir'] * item['Harga Satuan Terakhir']
        if item['Stok Akhir'] <= 0:
            item['Status Stok'] = 'STOK HABIS'
        elif item['Minimum Stok'] and item['Stok Akhir'] <= item['Minimum Stok']:
            item['Status Stok'] = 'STOK MINIMUM'
        else:
            item['Status Stok'] = 'TERSEDIA'

    wb = load_workbook(DB_PATH)
    ws = wb['DATABASE_STOK']
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    headers = [c.value for c in ws[1]]
    for _, item in sorted(stock.items()):
        ws.append([item.get(h, '') for h in headers])
    style_sheet(ws)
    add_table_if_needed(ws)
    wb.save(DB_PATH)


def build_laporan():
    rows = []
    for r in read_sheet('DATABASE_BARANG_MASUK'):
        rows.append({
            'Tanggal': r.get('Tanggal Input'), 'Jam': r.get('Jam Input'), 'Jenis Transaksi': 'BARANG MASUK',
            'ID Transaksi': r.get('ID Transaksi'), 'Code Number': r.get('Code Number'), 'Nama Barang': r.get('Nama Barang'),
            'Jenis Barang': r.get('Jenis Barang'), 'Qty': r.get('Total QTY') or r.get('QTY Masuk'), 'Satuan': r.get('Satuan'),
            'Nilai': r.get('Total Harga'), 'Pihak': r.get('Supplier / Asal Barang'), 'Status Koreksi': r.get('Status Koreksi'),
            'Keterangan': r.get('Keterangan')
        })
    for r in read_sheet('DATABASE_BARANG_KELUAR'):
        rows.append({
            'Tanggal': r.get('Tanggal Input'), 'Jam': r.get('Jam Input'), 'Jenis Transaksi': 'BARANG KELUAR',
            'ID Transaksi': r.get('ID Transaksi'), 'Code Number': r.get('Code Number'), 'Nama Barang': r.get('Nama Barang'),
            'Jenis Barang': r.get('Jenis Barang'), 'Qty': r.get('Total QTY') or r.get('QTY Keluar'), 'Satuan': r.get('Satuan'),
            'Nilai': r.get('Total Harga'), 'Pihak': r.get('Tujuan / Pengguna Barang'), 'Status Koreksi': r.get('Status Koreksi'),
            'Keterangan': r.get('Keterangan')
        })
    rows.sort(key=lambda x: str(x.get('Tanggal') or '') + ' ' + str(x.get('Jam') or ''))
    return rows


@app.route('/')
def index():
    ensure_database()
    return render_template('index.html')


@app.route('/api/dashboard')
def dashboard():
    ensure_database()
    rebuild_stock()
    today = datetime.now()
    masuk = read_sheet('DATABASE_BARANG_MASUK')
    keluar = read_sheet('DATABASE_BARANG_KELUAR')
    stok = read_sheet('DATABASE_STOK')
    total_masuk_qty = sum(to_float(r.get('Total QTY') or r.get('QTY Masuk')) for r in masuk)
    total_keluar_qty = sum(to_float(r.get('Total QTY') or r.get('QTY Keluar')) for r in keluar)
    total_stok_qty = sum(to_float(r.get('Stok Akhir')) for r in stok)
    nilai_masuk = sum(to_float(r.get('Total Harga')) for r in masuk)
    nilai_keluar = sum(to_float(r.get('Total Harga')) for r in keluar)
    return jsonify({
        'tanggal': today.strftime('%d-%m-%Y'),
        'tanggal_iso': today.strftime('%Y-%m-%d'),
        'hari': DAY_NAMES[today.weekday()],
        'total_jenis_barang': len(stok),
        'total_barang_masuk': total_masuk_qty,
        'total_barang_keluar': total_keluar_qty,
        'total_stok_tersedia': total_stok_qty,
        'total_nilai_masuk': nilai_masuk,
        'total_nilai_keluar': nilai_keluar,
        'total_nilai_masuk_fmt': rupiah(nilai_masuk),
        'total_nilai_keluar_fmt': rupiah(nilai_keluar),
        'recent_stock': stok[-6:]
    })


@app.route('/api/stok')
def api_stok():
    rebuild_stock()
    return jsonify(read_sheet('DATABASE_STOK'))


@app.route('/api/barang-masuk', methods=['POST'])
def api_barang_masuk():
    payload = request.get_json(force=True)
    qty = to_float(payload.get('qty'))
    harga = to_float(payload.get('harga_satuan'))
    total = qty * harga
    code = str(payload.get('kode_barang') or '').strip().upper()
    tanggal = payload.get('tanggal') or now_date()
    row = {
        'ID Transaksi': next_id('DATABASE_BARANG_MASUK', 'BM'),
        'Tanggal Input': tanggal,
        'Jam Input': now_time(),
        'Code Number': code,
        'Nama Barang': payload.get('nama_barang') or '',
        'Jenis Barang': payload.get('kategori') or payload.get('jenis_barang') or '',
        'QTY Masuk': qty,
        'Satuan': payload.get('satuan') or '',
        'Total QTY': qty,
        'Harga Satuan': harga,
        'Total Harga': total,
        'Supplier / Asal Barang': payload.get('supplier') or '',
        'Nama File Foto Barang': '',
        'Path Foto Barang': '',
        'Nama File Nota': '',
        'Path Foto Nota': '',
        'Status Koreksi': 'BELUM DICEK',
        'Keterangan': payload.get('keterangan') or '',
        'User Input': payload.get('user_input') or 'Admin'
    }
    append_dict_row('DATABASE_BARANG_MASUK', row)
    upsert_master_barang(code, row['Nama Barang'], row['Jenis Barang'], row['Satuan'], harga, 0, '')
    rebuild_stock()
    return jsonify({'ok': True, 'message': f"Barang masuk tersimpan dengan ID {row['ID Transaksi']}", 'id': row['ID Transaksi'], 'total': total, 'total_fmt': rupiah(total)})


@app.route('/api/barang-keluar', methods=['POST'])
def api_barang_keluar():
    payload = request.get_json(force=True)
    qty = to_float(payload.get('qty'))
    harga = to_float(payload.get('harga_satuan'))
    total = qty * harga
    code = str(payload.get('kode_barang') or '').strip().upper()
    tanggal = payload.get('tanggal') or now_date()
    row = {
        'ID Transaksi': next_id('DATABASE_BARANG_KELUAR', 'BK'),
        'Tanggal Input': tanggal,
        'Jam Input': now_time(),
        'Code Number': code,
        'Nama Barang': payload.get('nama_barang') or '',
        'Jenis Barang': payload.get('kategori') or payload.get('jenis_barang') or '',
        'QTY Keluar': qty,
        'Satuan': payload.get('satuan') or '',
        'Total QTY': qty,
        'Harga Satuan': harga,
        'Total Harga': total,
        'Tujuan / Pengguna Barang': payload.get('penerima') or payload.get('tujuan') or '',
        'Nama File Foto Barang': '',
        'Path Foto Barang': '',
        'Nama File Bukti Keluar': '',
        'Path Foto Bukti Keluar': '',
        'Status Koreksi': 'BELUM DICEK',
        'Keterangan': payload.get('keterangan') or '',
        'User Input': payload.get('user_input') or 'Admin'
    }
    append_dict_row('DATABASE_BARANG_KELUAR', row)
    upsert_master_barang(code, row['Nama Barang'], row['Jenis Barang'], row['Satuan'], harga, 0, '')
    rebuild_stock()
    return jsonify({'ok': True, 'message': f"Barang keluar tersimpan dengan ID {row['ID Transaksi']}", 'id': row['ID Transaksi'], 'total': total, 'total_fmt': rupiah(total)})


@app.route('/api/laporan')
def api_laporan():
    ensure_database()
    rebuild_stock()
    return jsonify(build_laporan())


@app.route('/api/database-info')
def api_database_info():
    ensure_database()
    wb = load_workbook(DB_PATH, read_only=True)
    return jsonify({
        'path': str(DB_PATH).replace(str(BASE_DIR), ''),
        'exists': DB_PATH.exists(),
        'sheet_count': len(wb.sheetnames),
        'sheets': wb.sheetnames,
        'file_name': DB_PATH.name,
        'file_size_kb': round(DB_PATH.stat().st_size / 1024, 2)
    })


@app.route('/api/export-excel')
def api_export_excel():
    ensure_database()
    rebuild_stock()
    return send_file(DB_PATH, as_attachment=True, download_name='logistic_stock_database.xlsx')


if __name__ == '__main__':
    ensure_database()
    app.run(host='127.0.0.1', port=5000, debug=True)
