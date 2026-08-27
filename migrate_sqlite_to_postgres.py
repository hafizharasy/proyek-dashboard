"""
Migrasi data dari dashboard.db (SQLite lama) ke Postgres (Supabase).

Cara pakai:
    1. Set environment variable DATABASE_URL ke connection string POOLER
       Supabase (bukan direct connection db.<ref>.supabase.co -- itu akan
       gagal DNS di banyak jaringan). Contoh:
           postgresql://postgres.<project-ref>:<password>@aws-0-<region>.pooler.supabase.com:5432/postgres
    2. Taruh file dashboard.db lama di folder yang sama dengan script ini
       (atau ubah SQLITE_PATH di bawah).
    3. Jalankan: python migrate_sqlite_to_postgres.py

Script ini AMAN dijalankan berkali-kali (idempotent) -- tabel dibuat
dengan CREATE TABLE IF NOT EXISTS, dan data lama di Postgres akan
dikosongkan dulu (TRUNCATE) sebelum diisi ulang, supaya tidak dobel
kalau di-run lebih dari sekali.
"""

import os
import sqlite3
import sys

import psycopg2
import psycopg2.extras

SQLITE_PATH = os.path.join(os.path.dirname(__file__), 'dashboard.db')
DATABASE_URL = os.environ.get('DATABASE_URL')

SCHEMA_SQL = '''
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
'''


def main():
    if not DATABASE_URL:
        print('ERROR: environment variable DATABASE_URL belum di-set.')
        print('Contoh (PowerShell):')
        print('  $env:DATABASE_URL = "postgresql://postgres.xxxx:password@aws-0-xxxx.pooler.supabase.com:5432/postgres"')
        sys.exit(1)

    if not os.path.exists(SQLITE_PATH):
        print(f'ERROR: {SQLITE_PATH} tidak ditemukan. Taruh dashboard.db di folder yang sama.')
        sys.exit(1)

    print(f'Menyambung ke SQLite lama: {SQLITE_PATH}')
    src = sqlite3.connect(SQLITE_PATH)
    src.row_factory = sqlite3.Row

    print('Menyambung ke Postgres (Supabase)...')
    dst = psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
    dst_cur = dst.cursor()

    print('Membuat skema tabel (kalau belum ada)...')
    dst_cur.execute(SCHEMA_SQL)
    dst.commit()

    periods = src.execute('SELECT * FROM periode ORDER BY id').fetchall()
    print(f'Ditemukan {len(periods)} periode di dashboard.db lama. Memindahkan...')

    # Kosongkan dulu supaya aman dijalankan berkali-kali tanpa duplikat.
    dst_cur.execute('TRUNCATE TABLE berkas, periode RESTART IDENTITY CASCADE;')
    dst.commit()

    total_berkas = 0
    for p in periods:
        dst_cur.execute(
            '''INSERT INTO periode (slug, label, file_name, uploaded_at, pending_method)
               VALUES (%s, %s, %s, %s, %s) RETURNING id''',
            (p['slug'], p['label'], p['file_name'], p['uploaded_at'], p['pending_method'])
        )
        new_periode_id = dst_cur.fetchone()['id']

        rows = src.execute('SELECT * FROM berkas WHERE periode_id = ?', (p['id'],)).fetchall()
        psycopg2.extras.execute_batch(
            dst_cur,
            '''INSERT INTO berkas
               (periode_id, dinas, jenis_izin, kategori, status_usaha,
                masuk, selesai, selesai_asli, tepat_waktu, proses, pending, sk_terbit)
               VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)''',
            [
                (new_periode_id, r['dinas'], r['jenis_izin'], r['kategori'], r['status_usaha'],
                 r['masuk'], r['selesai'], r['selesai_asli'], r['tepat_waktu'],
                 r['proses'], r['pending'], r['sk_terbit'])
                for r in rows
            ]
        )
        dst.commit()
        total_berkas += len(rows)
        print(f'  -> {p["label"]:20s} : {len(rows)} baris dipindahkan')

    print(f'\nSELESAI. Total {len(periods)} periode, {total_berkas} baris berkas berhasil dipindahkan ke Postgres.')

    # Verifikasi cepat
    dst_cur.execute('SELECT COUNT(*) AS c FROM periode')
    n_periode = dst_cur.fetchone()['c']
    dst_cur.execute('SELECT COUNT(*) AS c FROM berkas')
    n_berkas = dst_cur.fetchone()['c']
    print(f'Verifikasi di Postgres -> periode: {n_periode}, berkas: {n_berkas}')

    src.close()
    dst_cur.close()
    dst.close()


if __name__ == '__main__':
    main()
