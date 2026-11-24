import hmac
import os
import re
from datetime import datetime

import psycopg2
import psycopg2.extras

from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, send_file
)
from openpyxl import load_workbook, Workbook

app = Flask(__name__)
# SECRET_KEY dipakai untuk menandatangani cookie sesi login. Wajib diset di
# environment (Vercel/Render); kalau tidak ada, dibuat acak tiap server start.
app.secret_key = os.environ.get('SECRET_KEY') or os.urandom(32)

# ------------------------------------------------------------------
# KONFIGURASI
# ------------------------------------------------------------------
PIN_AKSES = os.environ.get('PIN_AKSES')  # PIN login dashboard, diset lewat environment variable
DATABASE_URL = os.environ.get('DATABASE_URL')  # connection string Postgres (Supabase pooler)
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

BULAN_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
            'Agustus', 'September', 'Oktober', 'November', 'Desember']

# Kata kunci untuk mengenali header kolom pada file Excel yang diunggah.
# Pencarian tidak peka huruf besar/kecil & mengabaikan spasi/garis bawah.
# Setiap field punya beberapa "tingkat" keyword: tingkat paling spesifik
# dicoba dulu, baru turun ke yang lebih umum. Ini penting supaya misalnya
# 'jumlah_berkas_selesai_dpmptsp' (selesai final di DPMPTSP) tidak salah
# tertukar dengan 'jumlah_berkas_selesai_tim_teknis' (selesai tahap teknis).
FIELD_GROUPS = {
    'dinas': [['pdpelaksana'], ['dinas'], ['instansi'], ['opd']],
    'jenis_pelayanan': [['jenispelayanan']],
    'nama_layanan': [['namalayanan']],
    'jenis_izin_fallback': [['jenisizin'], ['jenis'], ['izin', 'perizinan']],
    'status_usaha': [['statuspelakuusaha'], ['statususaha']],
    'sistem_pelayanan': [['sistempelayanan']],
    'masuk': [['jumlahberkasmasuk'], ['berkasmasuk'], ['masuk']],
    'selesai': [['berkasselesaidpmptsp'], ['selesaidpmptsp'], ['berkasselesai'], ['selesai']],
    'tepat_waktu': [['berkasselesaitepatwaktudpmptsp'], ['tepatwaktudpmptsp'], ['tepatwaktu'], ['tepat']],
    'proses': [['proses']],
    'pending': [['pending'], ['tunda']],
    # Tahap "lolos PD Teknis" -- dipakai sebagai sumber pemisahan Proses/Pending
    # kalau file TIDAK punya kolom Proses/Pending eksplisit.
    'selesai_pd_teknis': [['perstekselesaipdteknis'], ['selesaipdteknis']],
    # Jumlah SK/izin yang benar-benar diterbitkan (bukan sekadar selesai diproses).
    'penerbitan_ijin': [['penerbitanijinpdteknis'], ['penerbitanijin'], ['penerbitansk'], ['skterbit']],
}


# ------------------------------------------------------------------
# PENYIMPANAN DATA -- Postgres (Supabase), 2 tabel:
#   periode  : satu baris per periode/bulan (label, nama file, dst)
#   berkas   : satu baris per baris data Excel, terhubung ke periode
#              lewat kolom periode_id (foreign key)
# ------------------------------------------------------------------
def get_db():
    if not DATABASE_URL:
        raise RuntimeError(
            "Environment variable DATABASE_URL belum di-set. Isi dengan "
            "connection string Postgres (pakai connection pooler Supabase, "
            "bukan direct connection -- lihat README.md bagian Supabase)."
        )
    return psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS periode (
            id SERIAL PRIMARY KEY,
            slug TEXT UNIQUE NOT NULL,
            label TEXT NOT NULL,
            file_name TEXT,
            uploaded_at TEXT,
            pending_method TEXT NOT NULL DEFAULT 'tidak_ada'
        );

        CREATE TABLE IF NOT EXISTS berkas (
            id SERIAL PRIMARY KEY,
            periode_id INTEGER NOT NULL REFERENCES periode(id) ON DELETE CASCADE,
            dinas TEXT,
            jenis_izin TEXT,
            kategori TEXT,
            status_usaha TEXT,
            masuk REAL,
            selesai REAL,
            selesai_asli REAL NOT NULL DEFAULT 0,
            tepat_waktu REAL,
            proses REAL,
            pending REAL,
            sk_terbit REAL NOT NULL DEFAULT 0
        );

        CREATE INDEX IF NOT EXISTS idx_berkas_periode ON berkas(periode_id);
        CREATE INDEX IF NOT EXISTS idx_berkas_dinas ON berkas(dinas);
    ''')
    conn.commit()
    cur.close()
    conn.close()


def slugify(label):
    slug = re.sub(r'[^a-z0-9]+', '-', label.strip().lower()).strip('-')
    return slug or 'periode'


def save_period(label, rows, file_name, pending_method='tidak_ada', filter_segmen=None):
    """
    filter_segmen: None -> replace SELURUH data periode ini (perilaku lama,
                   dipakai kalau upload 1 file berisi semua data).
                   'pelaku-usaha' / 'non-pelaku-usaha' -> hanya isi/timpa
                   baris dengan status itu saja; baris status lain yang
                   sudah ada di periode yang sama (dari upload sebelumnya)
                   TIDAK ikut terhapus -- dipakai untuk menggabungkan 2 file
                   berbeda jadi satu periode (mis. Pelaku Usaha dari file A,
                   Non Pelaku Usaha dari file MONEV B).
    """
    key = slugify(label)
    now = datetime.now().strftime('%d %b %Y %H:%M')
    conn = get_db()
    cur = conn.cursor()

    if filter_segmen:
        rows = filter_rows_by_segmen(rows, filter_segmen)

    cur.execute('SELECT * FROM periode WHERE slug = %s', (key,))
    existing = cur.fetchone()

    if filter_segmen and existing:
        # Mode gabung: hapus HANYA baris berstatus sama dari upload
        # sebelumnya, baris status lain (dari file berbeda) tetap aman.
        periode_id = existing['id']
        status_target = 'pelaku usaha' if filter_segmen == 'pelaku-usaha' else 'non pelaku usaha'
        cur.execute(
            "DELETE FROM berkas WHERE periode_id = %s AND LOWER(TRIM(status_usaha)) = %s",
            (periode_id, status_target)
        )
        nama_gabungan = existing['file_name'] or ''
        daftar_file = [n.strip() for n in nama_gabungan.split('+')] if nama_gabungan else []
        if file_name not in daftar_file:
            daftar_file.append(file_name)
        metode_gabungan = pending_method if existing['pending_method'] == pending_method else 'campuran'
        cur.execute(
            'UPDATE periode SET file_name = %s, uploaded_at = %s, pending_method = %s WHERE id = %s',
            (' + '.join(daftar_file), now, metode_gabungan, periode_id)
        )
    else:
        # Upload biasa (tanpa filter_segmen) = ganti SELURUH periode itu.
        cur.execute('DELETE FROM periode WHERE slug = %s', (key,))
        cur.execute(
            '''INSERT INTO periode (slug, label, file_name, uploaded_at, pending_method)
               VALUES (%s, %s, %s, %s, %s) RETURNING id''',
            (key, label, file_name, now, pending_method)
        )
        periode_id = cur.fetchone()['id']

    psycopg2.extras.execute_batch(
        cur,
        '''INSERT INTO berkas (periode_id, dinas, jenis_izin, kategori, status_usaha, masuk, selesai, selesai_asli, tepat_waktu, proses, pending, sk_terbit)
           VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)''',
        [(periode_id, r['dinas'], r['jenis_izin'], r['kategori'], r['status_usaha'],
          r['masuk'], r['selesai'], r.get('selesai_asli', r['selesai']), r['tepat_waktu'],
          r['proses'], r['pending'], r.get('sk_terbit', 0)) for r in rows]
    )
    conn.commit()
    cur.close()
    conn.close()
    return key


def list_periods():
    """Daftar {key,label} semua periode, terbaru duluan."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT slug, label FROM periode ORDER BY id DESC')
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return [{'key': r['slug'], 'label': r['label']} for r in rows]


def get_period_meta(key=None):
    """Ambil metadata 1 periode (tanpa baris berkas-nya). Kalau key kosong
    atau tidak ditemukan, otomatis ambil periode paling baru."""
    conn = get_db()
    cur = conn.cursor()
    row = None
    if key:
        cur.execute('SELECT * FROM periode WHERE slug = %s', (key,))
        row = cur.fetchone()
    if row is None:
        cur.execute('SELECT * FROM periode ORDER BY id DESC LIMIT 1')
        row = cur.fetchone()
    cur.close()
    conn.close()
    return dict(row) if row else None


def get_period_rows(periode_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        'SELECT dinas, jenis_izin, kategori, status_usaha, masuk, selesai, selesai_asli, tepat_waktu, proses, pending, sk_terbit '
        'FROM berkas WHERE periode_id = %s', (periode_id,)
    )
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return [dict(r) for r in rows]


def delete_period(key):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('DELETE FROM periode WHERE slug = %s', (key,))  # cascade menghapus baris berkas terkait
    conn.commit()
    cur.close()
    conn.close()


def parse_period_label(label):
    """Ubah label seperti 'Maret 2026' jadi (tahun, index_bulan) supaya bisa
    diurutkan secara kronologis (bukan urutan upload). Label yang tidak
    berformat 'Bulan Tahun' (mis. 'Data Contoh') ditaruh paling akhir."""
    parts = label.strip().split()
    if len(parts) >= 2 and parts[0] in BULAN_ID and parts[-1].isdigit():
        return (int(parts[-1]), BULAN_ID.index(parts[0]))
    return (9999, 99)


def get_monthly_trend(segmen='semua'):
    """Total Berkas Masuk/Selesai/SLA per periode, diurutkan kronologis --
    dipakai untuk grafik tren naik/turun bulanan di halaman Ringkasan."""
    where = ''
    if segmen == 'pelaku-usaha':
        where = "AND LOWER(TRIM(b.status_usaha)) = 'pelaku usaha'"
    elif segmen == 'non-pelaku-usaha':
        where = "AND LOWER(TRIM(b.status_usaha)) = 'non pelaku usaha'"

    conn = get_db()
    cur = conn.cursor()
    cur.execute(f'''
        SELECT p.label AS label,
               COALESCE(SUM(b.masuk), 0) AS masuk,
               COALESCE(SUM(b.selesai), 0) AS selesai,
               COALESCE(SUM(b.tepat_waktu), 0) AS tepat_waktu
        FROM periode p
        LEFT JOIN berkas b ON b.periode_id = p.id {where}
        GROUP BY p.id
    ''')
    rows = cur.fetchall()
    cur.close()
    conn.close()

    data = [dict(r) for r in rows]
    data.sort(key=lambda d: parse_period_label(d['label']))
    for d in data:
        d['sla'] = (d['tepat_waktu'] / d['selesai'] * 100) if d['selesai'] else 0
    return data


init_db()


def normalize(text):
    return ''.join(ch for ch in str(text or '').lower() if ch.isalpha())


def find_column_priority(headers, groups):
    """Try each keyword tier in order; return the first header that matches
    any keyword within the first tier that finds a hit."""
    for tier in groups:
        for h in headers:
            if any(kw in normalize(h) for kw in tier):
                return h
    return None


def to_number(value):
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return value
    try:
        cleaned = ''.join(ch for ch in str(value) if ch.isdigit() or ch in '.-')
        return float(cleaned) if cleaned else 0
    except ValueError:
        return 0


def bersihkan_teks(s):
    """Normalisasi teks untuk pencocokan (hapus newline, spasi ganda,
    lowercase) -- dipakai untuk menyamakan nama dinas/jenis izin yang
    penulisannya sedikit beda antar sheet."""
    if not s:
        return ''
    s = str(s).replace('\n', ' ').replace('\r', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s.lower()


def detect_laporan_ptsp(sheet_list):
    """Deteksi sheet gaya 'laporan cetak PTSP' -- punya kolom eksplisit
    Proses & Pending, tapi nama dinasnya di merged cell (cuma muncul di
    baris pertama tiap kelompok, kolom sebelah kiri jenis izin). Beda dari
    format MONEV kode-kolom biasa yang ditangani find_best_sheet.
    Mengembalikan (nama_sheet, rows, header_idx) atau (None, None, None)."""
    for name, rows in sheet_list:
        for i, row in enumerate(rows[:15]):
            texts = [normalize(str(v)) for v in row if v]
            if 'proses' in texts and 'pending' in texts and any('masuk' in t for t in texts):
                return name, rows, i
    return None, None, None


def build_kategori_lookup(sheet_list, kecuali_sheet):
    """Cari sheet lain (selain kecuali_sheet) dengan format MONEV kode-kolom
    standar yang punya kolom dinas/nama_layanan/jenis_pelayanan, lalu
    bangun lookup {(dinas, nama_layanan) -> jenis_pelayanan} untuk
    referensi kategori. Dipakai supaya sheet laporan PTSP (yang tidak
    punya kolom kategori sendiri) tetap bisa dikelompokkan dengan benar."""
    kandidat = [(n, r) for n, r in sheet_list if n != kecuali_sheet]
    nama, rows = find_best_sheet(kandidat)
    if not rows:
        return {}
    headers = list(rows[0])
    data_rows = rows[1:]
    if data_rows and looks_like_code_row(data_rows[0]):
        data_rows = data_rows[1:]

    col_dinas = find_column_priority(headers, FIELD_GROUPS['dinas'])
    col_nama = find_column_priority(headers, FIELD_GROUPS['nama_layanan'])
    col_jenis = find_column_priority(headers, FIELD_GROUPS['jenis_pelayanan'])
    if not (col_dinas and col_nama and col_jenis):
        return {}
    i_dinas, i_nama, i_jenis = headers.index(col_dinas), headers.index(col_nama), headers.index(col_jenis)

    lookup = {}
    for row in data_rows:
        d = bersihkan_teks(row[i_dinas]) if i_dinas < len(row) else ''
        n = bersihkan_teks(row[i_nama]) if i_nama < len(row) else ''
        j = str(row[i_jenis]).strip() if i_jenis < len(row) and row[i_jenis] else None
        if d and n and j:
            lookup[(d, n)] = j
    return lookup


def parse_laporan_ptsp(rows, header_idx, kategori_lookup):
    """Parser khusus untuk sheet laporan PTSP: Proses/Pending/Berkas Selesai/
    Berkas Tepat Waktu/Berkas Masuk eksplisit, nama dinas di merged cell
    (forward-fill dari baris terakhir yang terisi). Sheet ini judulnya
    'LAPORAN PELAYANAN PERIZINAN NON BERUSAHA...' -- seluruh barisnya
    otomatis dianggap Non Pelaku Usaha."""
    headers = list(rows[header_idx])

    def cari(*kata_kunci):
        for i, h in enumerate(headers):
            if h and any(k in normalize(str(h)) for k in kata_kunci):
                return i
        return None

    i_dinas = cari('namaperizinan', 'namalayanan')
    i_jenis = i_dinas + 1 if i_dinas is not None else None
    i_proses = cari('proses')
    i_pending = cari('pending')
    i_selesai = cari('berkasselesai')
    i_tepat = cari('tepatwaktu')
    i_masuk = cari('berkasmasuk')
    if i_jenis is None or i_masuk is None:
        return []

    parsed = []
    dinas_sekarang = None
    for row in rows[header_idx + 1:]:
        if row is None or all(v is None for v in row):
            continue
        dinas_cell = row[i_dinas] if i_dinas is not None and i_dinas < len(row) else None
        if dinas_cell and str(dinas_cell).strip():
            dinas_sekarang = str(dinas_cell).strip()
        jenis_cell = row[i_jenis] if i_jenis < len(row) else None
        if not jenis_cell or not str(jenis_cell).strip() or not dinas_sekarang:
            continue
        jenis = str(jenis_cell).strip()

        masuk = to_number(row[i_masuk]) if i_masuk < len(row) else 0
        selesai = to_number(row[i_selesai]) if i_selesai is not None and i_selesai < len(row) else 0
        tepat_waktu = to_number(row[i_tepat]) if i_tepat is not None and i_tepat < len(row) else 0
        proses = to_number(row[i_proses]) if i_proses is not None and i_proses < len(row) else 0
        pending = to_number(row[i_pending]) if i_pending is not None and i_pending < len(row) else 0

        kategori = kategori_lookup.get((bersihkan_teks(dinas_sekarang), bersihkan_teks(jenis)), jenis)

        parsed.append({
            'dinas': dinas_sekarang,
            'jenis_izin': jenis,
            'kategori': kategori,
            'status_usaha': 'Non Pelaku Usaha',
            'masuk': masuk,
            'selesai': selesai,
            'selesai_asli': selesai,
            'tepat_waktu': tepat_waktu,
            'proses': proses,
            'pending': pending,
            'sk_terbit': selesai,
        })
    return parsed


def find_best_sheet(sheet_list):
    """Pilih sheet yang paling cocok jadi data utama -- yaitu yang salah
    satu barisnya (baris 1 atau baris 2, untuk jaga-jaga ada baris kode)
    mengandung kolom Dinas & Berkas Masuk yang bisa dikenali. Banyak file
    MONEV punya beberapa sheet (cover laporan, rekap per bulan, dll) --
    cuma satu yang datanya asli. Kalau ada beberapa kandidat cocok,
    prioritaskan nama sheet yang mengandung 'fix' (konvensi penamaan versi
    final), lalu yang skor kecocokan kolomnya paling tinggi."""
    kandidat = []
    for name, rows in sheet_list:
        if not rows:
            continue
        for baris_ke in (0, 1):
            if baris_ke >= len(rows):
                break
            headers = list(rows[baris_ke])
            if all(v in (None, '') for v in headers):
                continue
            col = {key: find_column_priority(headers, groups) for key, groups in FIELD_GROUPS.items()}
            if col['dinas'] and col['masuk']:
                skor = sum(1 for v in col.values() if v)
                kandidat.append((name, rows[baris_ke:], skor))
                break  # sheet ini sudah ketemu header-nya, tak perlu cek baris lain
    if not kandidat:
        return None, None
    kandidat.sort(key=lambda x: ('fix' not in x[0].lower(), -x[2]))
    return kandidat[0][0], kandidat[0][1]


def read_all_sheets(filepath):
    """Baca SEMUA sheet dari file Excel (.xlsx/.xlsm via openpyxl, .xls lama
    via xlrd). Mengembalikan list of (nama_sheet, rows)."""
    ext = os.path.splitext(filepath)[1].lower()
    sheet_list = []
    if ext == '.xls':
        import xlrd
        wb = xlrd.open_workbook(filepath)
        for sn in wb.sheet_names():
            ws = wb.sheet_by_name(sn)
            rows = []
            for r in range(ws.nrows):
                row = ws.row_values(r)
                # xlrd mengisi sel kosong dengan '' -- samakan dengan openpyxl (None)
                rows.append(tuple(v if v != '' else None for v in row))
            sheet_list.append((sn, rows))
    else:
        wb = load_workbook(filepath, data_only=True)
        for ws in wb.worksheets:
            sheet_list.append((ws.title, list(ws.iter_rows(values_only=True))))
    return sheet_list


def read_excel_rows(filepath):
    """Baca baris mentah dari file Excel. Kalau file punya banyak sheet,
    otomatis pilih sheet yang datanya paling cocok (lihat find_best_sheet).
    Mengembalikan list of tuple, baris pertama adalah header."""
    sheet_list = read_all_sheets(filepath)
    if not sheet_list:
        return []

    nama, rows = find_best_sheet(sheet_list)
    if rows is not None:
        return rows
    # Fallback: tidak ada sheet yang jelas cocok -- pakai sheet pertama
    # seperti perilaku lama, biar pesan error di parse_excel yang menjelaskan.
    return sheet_list[0][1]


def looks_like_code_row(row):
    """Deteksi baris 'kode kolom' seperti pd_pelaksana, jenis_pelayanan, dst
    -- beberapa template resmi (mis. MONEV) punya baris kedua berisi nama
    kode snake_case sebelum data sungguhan dimulai. Baris seperti ini harus
    dilewati, bukan dibaca sebagai baris data."""
    text_vals = [str(v).strip() for v in row if isinstance(v, str) and str(v).strip()]
    if not text_vals:
        return False
    snake_like = [v for v in text_vals if re.fullmatch(r'[a-z][a-z0-9_]*', v)]
    return len(snake_like) >= max(1, len(text_vals) // 2)


def parse_excel(filepath):
    sheet_list = read_all_sheets(filepath)
    if not sheet_list:
        raise ValueError('File kosong atau tidak bisa dibaca.')

    # Coba deteksi format 'laporan cetak PTSP' dulu (kolom Proses/Pending
    # eksplisit, nama dinas di merged cell) -- kalau ketemu, pakai parser
    # khusus dan kategori-nya dicari dari sheet MONEV lain di file yang sama.
    nama_laporan, rows_laporan, header_idx = detect_laporan_ptsp(sheet_list)
    if rows_laporan is not None:
        kategori_lookup = build_kategori_lookup(sheet_list, nama_laporan)
        parsed = parse_laporan_ptsp(rows_laporan, header_idx, kategori_lookup)
        if parsed:
            return parsed, 'eksplisit'
        # Kalau ternyata gagal (0 baris), lanjut ke jalur normal di bawah.

    all_rows = read_excel_rows(filepath)
    if not all_rows:
        raise ValueError('File kosong atau tidak bisa dibaca.')
    headers = list(all_rows[0])
    data_rows = all_rows[1:]

    # Sebagian template (mis. MONEV) punya baris kedua berisi kode kolom
    # snake_case sebelum data sungguhan -- lewati kalau terdeteksi.
    if data_rows and looks_like_code_row(data_rows[0]):
        data_rows = data_rows[1:]

    col = {key: find_column_priority(headers, groups) for key, groups in FIELD_GROUPS.items()}

    if not col['dinas'] or not col['masuk']:
        raise ValueError(
            "Kolom 'Dinas' (atau 'pd_pelaksana') dan 'Berkas Masuk' tidak ditemukan. "
            "Periksa header kolom pada file Excel Anda."
        )

    # Kolom nama jenis izin: prioritaskan nama layanan yang paling spesifik,
    # lalu turun ke jenis pelayanan (kategori), lalu ke nama kolom lama.
    jenis_display_col = col['nama_layanan'] or col['jenis_pelayanan'] or col['jenis_izin_fallback']
    kategori_col = col['jenis_pelayanan'] or jenis_display_col

    idx = {key: headers.index(val) if val else None for key, val in col.items()}
    idx['jenis_display'] = headers.index(jenis_display_col) if jenis_display_col else None
    idx['kategori'] = headers.index(kategori_col) if kategori_col else None

    # Cara menentukan Proses/Pending, dari yang paling akurat ke paling umum:
    #  1. 'eksplisit'  -- file punya kolom Proses/Pending sendiri, pakai langsung.
    #  2. 'tahapan'    -- file punya kolom 'selesai PD Teknis' (tahap sebelum
    #                     penerbitan DPMPTSP). Proses = sudah lolos PD Teknis
    #                     tapi belum terbit DPMPTSP; Pending = belum lolos
    #                     tahap PD Teknis sama sekali.
    #  3. 'tidak_ada'  -- tidak ada sumber data untuk memisahkan; semua yang
    #                     belum selesai dihitung sebagai "Proses", Pending = 0.
    if col['proses'] is not None or col['pending'] is not None:
        pending_method = 'eksplisit'
    elif col['selesai_pd_teknis'] is not None:
        pending_method = 'tahapan'
    else:
        pending_method = 'tidak_ada'

    def cell(row, key):
        i = idx.get(key)
        return row[i] if i is not None else None

    parsed = []
    for row in data_rows:
        if row is None or all(v is None for v in row):
            continue
        dinas_val = cell(row, 'dinas')
        if dinas_val is None or not str(dinas_val).strip():
            # Baris ringkasan/total di bagian bawah file biasanya tidak
            # punya nama dinas -- lewati baris seperti ini.
            continue

        dinas = str(dinas_val).strip()
        jenis = str(cell(row, 'jenis_display') or 'Tidak Diketahui').strip() or 'Tidak Diketahui'
        kategori = str(cell(row, 'kategori') or jenis).strip() or jenis
        status_usaha = str(cell(row, 'status_usaha') or '').strip()

        masuk = to_number(cell(row, 'masuk'))
        selesai_asli = to_number(cell(row, 'selesai')) if idx['selesai'] is not None else 0
        tepat_waktu = to_number(cell(row, 'tepat_waktu')) if idx['tepat_waktu'] is not None else 0

        # SK Terbit = Berkas Selesai DPMPTSP, berlaku untuk SEMUA baris
        # (Pelaku Usaha maupun Non Pelaku Usaha) -- kolom
        # jumlah_penerbitan_ijin_pd_teknis TIDAK dipakai lagi.
        sk_terbit = selesai_asli
        selesai = selesai_asli

        if pending_method == 'eksplisit':
            proses = to_number(cell(row, 'proses')) if idx['proses'] is not None else 0
            pending = to_number(cell(row, 'pending')) if idx['pending'] is not None else 0
        elif pending_method == 'tahapan':
            selesai_pd_teknis = to_number(cell(row, 'selesai_pd_teknis'))
            # Sisa = bagian yang BELUM tercatat selesai di DPMPTSP -- ini yang
            # dibagi ke Proses/Pending, supaya berkas yang sudah selesai
            # (dihitung terpisah di KPI "Total Berkas Selesai") tidak ikut
            # dobel dihitung lagi sebagai "Sedang Diproses".
            sisa = max(0, masuk - selesai)
            if sisa == 0:
                proses = 0
                pending = 0
            elif selesai_pd_teknis == 0 and selesai == 0:
                # Kedua tahap masih 0/kosong -- belum tersentuh sama sekali.
                pending = sisa
                proses = 0
            else:
                # Minimal salah satu tahap (PD Teknis atau DPMPTSP) sudah
                # tersentuh -- sisanya dianggap sedang berjalan di pipeline.
                proses = sisa
                pending = 0
        else:
            # Data sumber tidak memisahkan status "proses" vs "pending" --
            # seluruh berkas yang belum selesai dihitung sebagai "proses".
            proses = max(0, masuk - selesai)
            pending = 0

        if status_usaha.strip().lower() == 'pelaku usaha':
            # Pelaku Usaha tidak punya konsep Proses/Pending -- begitu
            # selesai, SK-nya otomatis terbit (biasanya lewat OSS/sistem
            # otomatis), jadi tidak ada tahap "menggantung" yang perlu
            # dilacak seperti Non Pelaku Usaha.
            proses = 0
            pending = 0

        parsed.append({
            'dinas': dinas,
            'jenis_izin': jenis,
            'kategori': kategori,
            'status_usaha': status_usaha,
            'masuk': masuk,
            'selesai': selesai,
            'selesai_asli': selesai_asli,
            'tepat_waktu': tepat_waktu,
            'proses': proses,
            'pending': pending,
            'sk_terbit': sk_terbit,
        })

    if not parsed:
        raise ValueError('Tidak ada baris data yang bisa dibaca dari file ini.')

    return parsed, pending_method


# ------------------------------------------------------------------
# PERHITUNGAN & AGREGASI UNTUK DASHBOARD
# ------------------------------------------------------------------
def build_dashboard_context(period_data, period_key, all_periods):
    rows = period_data['rows'] if period_data else []

    totals = {'masuk': 0, 'selesai': 0, 'selesai_asli': 0, 'tepat_waktu': 0, 'proses': 0, 'pending': 0, 'sk_terbit': 0}
    for r in rows:
        for k in totals:
            totals[k] += r.get(k, 0) or 0
    totals['backlog'] = totals['proses'] + totals['pending']
    # SLA sengaja dihitung dari selesai_asli (murni kolom "selesai DPMPTSP"),
    # BUKAN dari 'selesai' yang sudah termasuk fallback SK Terbit -- supaya
    # dinas yang tidak melacak ketepatan waktu (mis. retribusi/reklame, yang
    # langsung lapor SK Terbit) tidak ikut menambah penyebut dan menurunkan
    # rasio SLA secara tidak adil.
    totals['sla'] = (totals['tepat_waktu'] / totals['selesai_asli'] * 100) if totals['selesai_asli'] else 0

    # Agregasi per dinas
    by_dinas = {}
    for r in rows:
        d = by_dinas.setdefault(r['dinas'], {
            'dinas': r['dinas'], 'masuk': 0, 'selesai': 0, 'selesai_asli': 0,
            'tepat_waktu': 0, 'proses': 0, 'pending': 0
        })
        d['masuk'] += r['masuk']
        d['selesai'] += r['selesai']
        d['selesai_asli'] += r.get('selesai_asli', r['selesai'])
        d['tepat_waktu'] += r['tepat_waktu']
        d['proses'] += r['proses']
        d['pending'] += r['pending']

    dinas_list = []
    for d in by_dinas.values():
        d = dict(d)
        d['backlog_ratio'] = ((d['proses'] + d['pending']) / d['masuk']) if d['masuk'] else 0
        # Sama seperti totals['sla'] -- pakai selesai_asli supaya dinas yang
        # menyelesaikan lewat SK Terbit (tanpa tracking tepat waktu) tidak
        # ikut menurunkan rasio SLA mereka secara tidak adil.
        d['sla_ratio'] = (d['tepat_waktu'] / d['selesai_asli']) if d['selesai_asli'] else 0
        dinas_list.append(d)

    produktif = sorted(dinas_list, key=lambda x: x['selesai'], reverse=True)[:5]
    bottleneck = sorted(
        [d for d in dinas_list if d['masuk'] > 0],
        key=lambda x: x['backlog_ratio'], reverse=True
    )[:5]
    tepat_waktu_top = sorted(
        [d for d in dinas_list if d['selesai'] > 0],
        key=lambda x: x['sla_ratio'], reverse=True
    )[:5]
    # Daftar LENGKAP semua dinas (bukan cuma top 5), diurutkan nama --
    # dipakai untuk tabel referensi penuh di halaman Peringkat Dinas.
    dinas_all = sorted(dinas_list, key=lambda x: x['dinas'])

    # Agregasi per jenis izin (untuk pie chart volume permohonan)
    by_jenis = {}
    for r in rows:
        key = r.get('kategori') or r['jenis_izin']
        by_jenis[key] = by_jenis.get(key, 0) + r['masuk']
    jenis_sorted = sorted(by_jenis.items(), key=lambda x: x[1], reverse=True)
    if len(jenis_sorted) > 6:
        top6 = jenis_sorted[:6]
        sisanya = sum(v for _, v in jenis_sorted[6:])
        pie_data = top6 + [('Lainnya', sisanya)]
    else:
        pie_data = jenis_sorted
    total_jenis_masuk = sum(v for _, v in jenis_sorted) or 1
    # Ranking lengkap semua jenis izin (bukan cuma top 6) -- untuk tabel
    # daftar peringkat volume di samping pie chart.
    jenis_ranked = [
        {'nama': nama, 'masuk': v, 'persen': v / total_jenis_masuk * 100}
        for nama, v in jenis_sorted
    ]

    # Daftar peringatan dini (overdue) -- pakai selesai_asli, bukan selesai,
    # supaya baris yang selesai lewat SK Terbit (tidak melacak tepat waktu)
    # tidak salah muncul sebagai "sangat terlambat".
    warning_list = []
    for r in rows:
        overdue = r.get('selesai_asli', r['selesai']) - r['tepat_waktu']
        if overdue > 0:
            warning_list.append({**r, 'overdue': overdue})
    warning_list.sort(key=lambda x: x['overdue'], reverse=True)

    # Ringkasan masalah per dinas -- dinas mana yang paling sering muncul
    # di daftar "perlu perhatian khusus".
    warning_by_dinas_map = {}
    for w in warning_list:
        d = warning_by_dinas_map.setdefault(w['dinas'], {'dinas': w['dinas'], 'jumlah_jenis': 0, 'total_overdue': 0})
        d['jumlah_jenis'] += 1
        d['total_overdue'] += w['overdue']
    warning_by_dinas = sorted(warning_by_dinas_map.values(), key=lambda x: x['total_overdue'], reverse=True)[:5]
    total_overdue = sum(w['overdue'] for w in warning_list)

    return {
        'rows': rows,
        'totals': totals,
        'produktif': produktif,
        'bottleneck': bottleneck,
        'tepat_waktu_top': tepat_waktu_top,
        'dinas_all': dinas_all,
        'jumlah_dinas': len(dinas_list),
        'pie_labels': [x[0] for x in pie_data],
        'pie_values': [x[1] for x in pie_data],
        'jenis_ranked': jenis_ranked,
        'jumlah_jenis_izin': len(jenis_sorted),
        'warning_list': warning_list,
        'warning_by_dinas': warning_by_dinas,
        'total_overdue': total_overdue,
        'file_name': period_data.get('file_name') if period_data else None,
        'uploaded_at': period_data.get('uploaded_at') if period_data else None,
        'pending_method': period_data.get('pending_method', 'tidak_ada') if period_data else 'tidak_ada',
        'period_label': period_data.get('label') if period_data else None,
        'period_key': period_key,
        'all_periods': all_periods,
    }


def sample_rows():
    import random
    random.seed(42)
    dinas_jenis = {
        'Dinas Perindustrian dan Tenaga Kerja': [
            'Permohonan Baru Kartu AK-1', 'Perpanjangan Kartu AK-1', 'Izin Operasional LPK'],
        'Dinas Penanaman Modal dan PTSP': [
            'Izin Usaha Mikro Kecil', 'Izin Mendirikan Bangunan', 'Tanda Daftar Perusahaan'],
        'Dinas Kependudukan dan Catatan Sipil': [
            'Akta Kelahiran', 'Kartu Keluarga', 'KTP Elektronik'],
        'Dinas Perhubungan': ['Izin Trayek Angkutan', 'Izin Usaha Bengkel'],
        'Dinas Lingkungan Hidup': ['Izin Lingkungan (UKL-UPL)', 'Izin Pembuangan Limbah'],
        'Dinas Kesehatan': ['Izin Praktik Tenaga Kesehatan', 'Sertifikat Laik Higiene'],
        'Dinas Pariwisata': ['Tanda Daftar Usaha Pariwisata'],
    }
    rows = []
    for dinas, jenis_list in dinas_jenis.items():
        for jenis in jenis_list:
            masuk = random.randint(20, 100)
            selesai = round(masuk * random.uniform(0.55, 0.95))
            tepat_waktu = round(selesai * random.uniform(0.5, 0.98))
            proses = round((masuk - selesai) * random.uniform(0.4, 0.9))
            pending = max(0, masuk - selesai - proses)
            sk_terbit = round(selesai * random.uniform(0.85, 1.0))
            rows.append({
                'dinas': dinas, 'jenis_izin': jenis, 'kategori': jenis, 'status_usaha': 'Non Pelaku Usaha',
                'masuk': masuk, 'selesai': selesai,
                'tepat_waktu': tepat_waktu, 'proses': proses, 'pending': pending,
                'sk_terbit': sk_terbit, 'selesai_asli': selesai,
            })
    return rows


# ------------------------------------------------------------------
# ROUTES
# ------------------------------------------------------------------
def is_logged_in():
    return 'sudah_login' in session


@app.route('/')
def index():
    if not is_logged_in():
        return redirect(url_for('halaman_login'))
    periode = request.args.get('periode')
    return redirect(url_for('halaman_ringkasan', periode=periode) if periode else url_for('halaman_ringkasan'))


CUSTOM_ALL_KEY = 'semua-periode'


SEGMEN_VALID = ('semua', 'pelaku-usaha', 'non-pelaku-usaha')


def filter_rows_by_segmen(rows, segmen):
    """Saring baris berdasarkan status_pelaku_usaha. 'semua' = tidak difilter
    (gabungan Pelaku Usaha + Non Pelaku Usaha, seperti sebelumnya)."""
    if segmen == 'pelaku-usaha':
        return [r for r in rows if (r.get('status_usaha') or '').strip().lower() == 'pelaku usaha']
    if segmen == 'non-pelaku-usaha':
        return [r for r in rows if (r.get('status_usaha') or '').strip().lower() == 'non pelaku usaha']
    return rows


def build_combined_period():
    """Gabungkan rows dari SEMUA periode yang tersimpan jadi satu 'periode
    virtual' -- query langsung ke tabel berkas TANPA filter periode_id."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT COUNT(*) c FROM periode')
    jumlah_periode = cur.fetchone()['c']
    if jumlah_periode == 0:
        cur.close()
        conn.close()
        return None
    cur.execute(
        'SELECT dinas, jenis_izin, kategori, status_usaha, masuk, selesai, selesai_asli, tepat_waktu, proses, pending, sk_terbit FROM berkas'
    )
    rows = cur.fetchall()
    cur.execute('SELECT DISTINCT pending_method FROM periode')
    metode_dipakai = {r['pending_method'] for r in cur.fetchall()}
    cur.close()
    conn.close()
    pending_method = metode_dipakai.pop() if len(metode_dipakai) == 1 else 'campuran'
    return {
        'label': 'Gabungan Semua Periode',
        'rows': [dict(r) for r in rows],
        'file_name': f'{jumlah_periode} periode digabung',
        'uploaded_at': '-',
        'pending_method': pending_method,
    }


def build_usaha_comparison(rows):
    """Bandingkan total Masuk/Selesai/Proses/Pending antara Pelaku Usaha vs
    Non Pelaku Usaha, dihitung dari rows SEBELUM difilter oleh toggle segmen
    -- supaya panel perbandingan di Ringkasan tetap informatif walau user
    sedang melihat salah satu segmen saja."""
    hasil = {
        'pelaku_usaha': {'masuk': 0, 'selesai': 0, 'proses': 0, 'pending': 0},
        'non_pelaku_usaha': {'masuk': 0, 'selesai': 0, 'proses': 0, 'pending': 0},
    }
    for r in rows:
        status = (r.get('status_usaha') or '').strip().lower()
        if status == 'pelaku usaha':
            key = 'pelaku_usaha'
        elif status == 'non pelaku usaha':
            key = 'non_pelaku_usaha'
        else:
            continue
        hasil[key]['masuk'] += r['masuk']
        hasil[key]['selesai'] += r['selesai']
        hasil[key]['proses'] += r['proses']
        hasil[key]['pending'] += r['pending']
    return hasil


def get_page_context(active_page):
    requested_key = request.args.get('periode')
    segmen = request.args.get('segmen', 'semua')
    if segmen not in SEGMEN_VALID:
        segmen = 'semua'
    all_periods = list_periods()

    if requested_key == CUSTOM_ALL_KEY and all_periods:
        key = CUSTOM_ALL_KEY
        period_data = build_combined_period()
    else:
        meta = get_period_meta(requested_key)
        if meta:
            key = meta['slug']
            period_data = dict(meta)
            period_data['rows'] = get_period_rows(meta['id'])
        else:
            key, period_data = None, None

    period_had_data = bool(period_data and period_data['rows'])
    usaha_comparison = build_usaha_comparison(period_data['rows']) if period_data else None
    if period_data:
        period_data = dict(period_data)
        period_data['rows'] = filter_rows_by_segmen(period_data['rows'], segmen)

    context = build_dashboard_context(period_data, key, all_periods)
    context['period_had_data'] = period_had_data
    context['usaha_comparison'] = usaha_comparison
    tahun_sekarang = datetime.now().year
    context['bulan_list'] = BULAN_ID
    context['tahun_list'] = list(range(tahun_sekarang - 2, tahun_sekarang + 2))
    context['bulan_default'] = BULAN_ID[datetime.now().month - 1]
    context['tahun_default'] = tahun_sekarang
    context['active_page'] = active_page
    context['combined_available'] = len(all_periods) > 1
    context['segmen'] = segmen

    trend = get_monthly_trend(segmen)
    context['trend'] = trend
    context['trend_delta'] = None
    if len(trend) >= 2:
        now_m, prev_m = trend[-1], trend[-2]
        def pct(now_val, prev_val):
            return ((now_val - prev_val) / prev_val * 100) if prev_val else None
        context['trend_delta'] = {
            'label_now': now_m['label'],
            'label_prev': prev_m['label'],
            'masuk': pct(now_m['masuk'], prev_m['masuk']),
            'selesai': pct(now_m['selesai'], prev_m['selesai']),
        }
    return context


@app.route('/ringkasan')
def halaman_ringkasan():
    if not is_logged_in():
        return redirect(url_for('halaman_login'))
    return render_template('ringkasan.html', **get_page_context('ringkasan'))


@app.route('/peringkat-dinas')
def halaman_peringkat_dinas():
    if not is_logged_in():
        return redirect(url_for('halaman_login'))
    return render_template('peringkat_dinas.html', **get_page_context('peringkat-dinas'))


@app.route('/beban-izin')
def halaman_beban_izin():
    if not is_logged_in():
        return redirect(url_for('halaman_login'))
    return render_template('beban_izin.html', **get_page_context('beban-izin'))


@app.route('/tabel-monitoring')
def halaman_tabel_monitoring():
    if not is_logged_in():
        return redirect(url_for('halaman_login'))
    return render_template('tabel_monitoring.html', **get_page_context('tabel-monitoring'))


@app.route('/peringatan')
def halaman_peringatan():
    if not is_logged_in():
        return redirect(url_for('halaman_login'))
    return render_template('peringatan.html', **get_page_context('peringatan'))


@app.route('/login', methods=['GET', 'POST'])
def halaman_login():
    if request.method == 'POST':
        pin = request.form.get('pin', '')
        if not PIN_AKSES:
            flash('PIN akses belum dikonfigurasi di server. Hubungi admin.')
            return redirect(url_for('halaman_login'))
        if hmac.compare_digest(pin, PIN_AKSES):
            session['sudah_login'] = True
            return redirect(url_for('halaman_ringkasan'))
        else:
            flash('PIN salah. Silakan coba lagi.')
            return redirect(url_for('halaman_login'))
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.pop('sudah_login', None)
    return redirect(url_for('halaman_login'))


@app.route('/upload', methods=['POST'])
def upload_excel():
    if not is_logged_in():
        return redirect(url_for('halaman_login'))

    file = request.files.get('file_excel')
    if not file or file.filename == '':
        flash('Tidak ada file yang dipilih.')
        return redirect(url_for('halaman_ringkasan'))

    bulan_pilihan = (request.form.get('bulan') or '').strip()
    tahun_pilihan = (request.form.get('tahun') or '').strip()
    now = datetime.now()
    if bulan_pilihan and tahun_pilihan:
        periode_label = f'{bulan_pilihan} {tahun_pilihan}'
    elif bulan_pilihan:
        periode_label = f'{bulan_pilihan} {now.year}'
    else:
        periode_label = f'{BULAN_ID[now.month - 1]} {now.year}'

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    filter_segmen = (request.form.get('sumber_segmen') or '').strip()
    if filter_segmen not in ('pelaku-usaha', 'non-pelaku-usaha'):
        filter_segmen = None

    try:
        rows, pending_method = parse_excel(filepath)
        if filter_segmen:
            rows_cocok = filter_rows_by_segmen(rows, filter_segmen)
            if not rows_cocok:
                label_segmen = 'Pelaku Usaha' if filter_segmen == 'pelaku-usaha' else 'Non Pelaku Usaha'
                flash(f'Tidak ada baris berstatus "{label_segmen}" di file ini -- tidak ada yang disimpan.')
                return redirect(url_for('halaman_ringkasan'))
        key = save_period(periode_label, rows, file.filename, pending_method, filter_segmen)
        if filter_segmen:
            label_segmen = 'Pelaku Usaha' if filter_segmen == 'pelaku-usaha' else 'Non Pelaku Usaha'
            jumlah = len(filter_rows_by_segmen(rows, filter_segmen))
            flash(f'Berhasil menggabungkan {jumlah} baris "{label_segmen}" ke periode "{periode_label}".')
        else:
            flash(f'Berhasil memuat {len(rows)} baris data untuk periode "{periode_label}".')
        return redirect(url_for('halaman_ringkasan', periode=key))
    except Exception as e:
        flash(f'Gagal membaca file: {e}')
        return redirect(url_for('halaman_ringkasan'))


@app.route('/muat-contoh')
def muat_data_contoh():
    if not is_logged_in():
        return redirect(url_for('halaman_login'))
    key = save_period('Data Contoh', sample_rows(), 'data-contoh.xlsx', 'eksplisit')
    flash('Data contoh berhasil dimuat.')
    return redirect(url_for('halaman_ringkasan', periode=key))


@app.route('/hapus-periode/<key>')
def hapus_periode(key):
    if not is_logged_in():
        return redirect(url_for('halaman_login'))
    meta = get_period_meta(key)
    label = meta['label'] if meta and meta['slug'] == key else key
    delete_period(key)
    flash(f'Periode "{label}" telah dihapus.')
    return redirect(url_for('halaman_ringkasan'))


@app.route('/template-excel')
def unduh_template():
    if not is_logged_in():
        return redirect(url_for('halaman_login'))
    wb = Workbook()
    ws = wb.active
    ws.title = 'Data Perizinan'
    ws.append([
        'pd_pelaksana', 'jenis_pelayanan', 'nama_layanan', 'status_pelaku_usaha',
        'jumlah_berkas_masuk', 'jumlah_berkas_selesai_tim_teknis',
        'jumlah_perstek_selesai_pd_teknis', 'jumlah_perstek_selesai_tepat_waktu_pd_teknis',
        'jumlah_penerbitan_ijin_pd_teknis', 'jumlah_berkas_selesai_dpmptsp',
        'jumlah_berkas_selesai_tepat_waktu_dpmptsp', 'jumlah_pelaku_usaha',
        'peraturan_mengatur_ketepatan_waktu', 'sistem_pelayanan', 'kesesuaian_dengan_ketentuan',
    ])
    contoh_baris = [
        ['Dinas Perindustrian dan Tenaga Kerja', 'Kartu Tanda Pencari Kerja (Kartu AK-1)',
         'Permohonan Baru Kartu Tanda Pencari Kerja (Kartu AK-1)', 'Non Pelaku Usaha',
         0, 0, 0, 0, 0, 0, 0, 0, 'Peraturan Walikota Surabaya Tahun 2021', 'SSW Alfa', 'Ya'],
        ['Dinas Perindustrian dan Tenaga Kerja', 'Kartu Tanda Pencari Kerja (Kartu AK-1)',
         'Permohonan Perpanjangan Kartu Tanda Pencari Kerja (Kartu AK-1)', 'Non Pelaku Usaha',
         0, 0, 0, 0, 0, 0, 0, 0, 'Peraturan Walikota Surabaya Tahun 2021', 'SSW Alfa', 'Ya'],
        ['Dinas Perindustrian dan Tenaga Kerja', 'Kartu Tanda Pencari Kerja (Kartu AK-1)',
         'Pemutakhiran Data Kartu Tanda Pencari Kerja (Kartu AK-1)', 'Non Pelaku Usaha',
         0, 0, 0, 0, 0, 0, 0, 0, 'Peraturan Walikota Surabaya Tahun 2021', 'SSW Alfa', 'Ya'],
        ['Dinas Perindustrian dan Tenaga Kerja', 'Lembaga Pelatihan Kerja (LPK) Pemerintah',
         'Izin Baru Lembaga Pelatihan Kerja (LPK) Pemerintah', 'Non Pelaku Usaha',
         0, 0, 0, 0, 0, 0, 0, 0, 'Peraturan Walikota Surabaya Tahun 2021', 'SSW Alfa', 'Ya'],
        ['Dinas Perindustrian dan Tenaga Kerja', 'Lembaga Pelatihan Kerja (LPK) Pemerintah',
         'Perubahan Lembaga Pelatihan Kerja (LPK) Pemerintah', 'Non Pelaku Usaha',
         0, 0, 0, 0, 0, 0, 0, 0, 'Peraturan Walikota Surabaya Tahun 2021', 'SSW Alfa', 'Ya'],
        ['Dinas Perindustrian dan Tenaga Kerja', 'Lembaga Pelatihan Kerja (LPK) Perusahaan',
         'Izin Baru Lembaga Pelatihan Kerja (LPK) Perusahaan', 'Non Pelaku Usaha',
         0, 0, 0, 0, 0, 0, 0, 0, 'Peraturan Walikota Surabaya Tahun 2021', 'SSW Alfa', 'Ya'],
        ['Dinas Perindustrian dan Tenaga Kerja', 'Lembaga Pelatihan Kerja (LPK) Perusahaan',
         'Perubahan Lembaga Pelatihan Kerja (LPK) Perusahaan', 'Non Pelaku Usaha',
         0, 0, 0, 0, 0, 0, 0, 0, 'Peraturan Walikota Surabaya Tahun 2021', 'SSW Alfa', 'Ya'],
    ]
    for baris in contoh_baris:
        ws.append(baris)
    out_path = os.path.join(UPLOAD_FOLDER, 'template-monitoring-perizinan.xlsx')
    wb.save(out_path)
    return send_file(out_path, as_attachment=True, download_name='template-monitoring-perizinan.xlsx')


if __name__ == '__main__':
    app.run(debug=True)
