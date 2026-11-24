# Dashboard Kinerja Pelayanan Perizinan

## Cara menjalankan
1. Buat project Supabase (gratis) di https://supabase.com kalau belum punya,
   lalu ambil **connection string pooler**-nya (Project Settings -> Database
   -> Connect -> pilih "Session pooler", BUKAN "Direct connection" -- lihat
   bagian "Database: Supabase" di bawah untuk alasannya).
2. Install dependensi:
   pip install -r requirements.txt
3. Set environment variable DATABASE_URL ke connection string itu, plus
   PIN_AKSES (PIN login) dan SECRET_KEY (string acak panjang).
   PowerShell:  $env:DATABASE_URL = "postgresql://postgres.xxxx:password@aws-0-xxxx.pooler.supabase.com:5432/postgres"
                $env:PIN_AKSES = "<PIN pilihan Anda>"
                $env:SECRET_KEY = "<string acak panjang>"
   macOS/Linux: export DATABASE_URL="postgresql://postgres.xxxx:password@aws-0-xxxx.pooler.supabase.com:5432/postgres"
                export PIN_AKSES="<PIN pilihan Anda>"
                export SECRET_KEY="<string acak panjang>"
   (contoh membuat SECRET_KEY: python -c "import secrets; print(secrets.token_hex(32))")
4. Jalankan aplikasi:
   python app.py
5. Buka http://127.0.0.1:5000 di browser.
6. Masuk dengan PIN yang Anda set di PIN_AKSES.

## Struktur
- app.py                -> logika backend Flask (login PIN, upload & parsing Excel, agregasi data, riwayat per periode, routing multi-halaman)
- templates/layout.html -> kerangka bersama: sidebar, header, pemilih periode, form impor, empty-state
- templates/login.html  -> halaman login PIN
- templates/ringkasan.html       -> halaman /ringkasan (KPI cards)
- templates/peringkat_dinas.html -> halaman /peringkat-dinas (leaderboard 3 tab)
- templates/beban_izin.html      -> halaman /beban-izin (pie chart jenis izin)
- templates/tabel_monitoring.html-> halaman /tabel-monitoring (tabel cari & sort)
- templates/peringatan.html      -> halaman /peringatan (daftar perlu perhatian khusus)
- migrate_sqlite_to_postgres.py  -> script sekali-pakai untuk memindahkan data lama dari dashboard.db (SQLite) ke Postgres/Supabase
- Database: **Postgres di Supabase** (bukan SQLite/file lokal lagi), berisi tabel `periode` dan `berkas`; skema dibuat otomatis saat aplikasi pertama kali jalan (`CREATE TABLE IF NOT EXISTS`). Karena Postgres-nya di Supabase (server terpisah dari tempat aplikasi Flask di-deploy), data **tidak hilang** lagi setiap kali aplikasi di-restart/redeploy.
- uploads/               -> folder penyimpanan file Excel yang diunggah (sementara, tidak dipakai untuk penyimpanan data permanen)

Tiap menu di sidebar sekarang adalah halaman (URL) tersendiri, bukan lagi
scroll ke bagian tertentu di satu halaman -- periode yang sedang dipilih
tetap terbawa saat Anda pindah antar halaman.

## Riwayat data per bulan (multi-periode)
Setiap kali Anda mengimpor file Excel, pilih bulan & tahun lewat dropdown
(misal "Januari" + "2026") sebelum memilih file. Data akan tersimpan sebagai
periode tersendiri dan TIDAK menimpa periode lain yang sudah ada -- jadi
rekap Januari, Februari, Maret dst bisa dikirim satu per satu, termasuk
mengisi data bulan-bulan lama (backlog) meskipun aplikasi baru mulai
dipakai di bulan yang lebih baru.

- Pilih periode yang ingin dilihat lewat dropdown di pojok kanan atas, atau
  lewat daftar "Periode Tersimpan" di sidebar.
- Dropdown bulan & tahun defaultnya mengikuti bulan berjalan, tapi bebas
  diganti ke bulan/tahun mana saja -- termasuk untuk mengisi data bulan lalu
  yang belum sempat diimpor.
- Mengupload ulang dengan nama periode yang SAMA akan menimpa (update) data
  periode itu saja -- periode lain tetap aman.
- Tanda silang (&times;) di sebelah nama periode di sidebar untuk menghapus
  periode tersebut kalau perlu.

## Gabung 2 file jadi 1 periode (Pelaku Usaha + Non Pelaku Usaha)
Kalau data Pelaku Usaha dan Non Pelaku Usaha datang dari 2 sumber file
yang berbeda (mis. rekap biasa untuk Pelaku Usaha, file MONEV terpisah
untuk Non Pelaku Usaha), pakai dropdown **"Ambil dari file ini"** di form
upload:
1. Upload file pertama, pilih **"Hanya ambil baris Pelaku Usaha"**.
2. Upload file kedua ke bulan & tahun yang SAMA, pilih **"Hanya ambil
   baris Non Pelaku Usaha"**.

Kedua file akan digabung jadi satu periode -- upload ulang salah satu
segmen hanya mengganti baris segmen itu saja, segmen lainnya tetap aman.
Nama kedua file akan tercatat di banner atas ("sumber: file1.xlsx +
file2.xls").

Kalau file yang diupload punya banyak sheet (umum di file MONEV), sistem
otomatis memilih sheet yang datanya paling cocok (header-nya mengandung
kolom Dinas & Berkas Masuk) -- prioritas ke sheet yang namanya mengandung
"FIX" kalau ada beberapa yang cocok.

### Format "laporan cetak PTSP" (kolom Proses/Pending eksplisit)
Beberapa file MONEV (mis. sheet bernama sesuai nama bulan, seperti
"Februari") berupa laporan cetak dengan kolom **Proses** dan **Pending**
eksplisit langsung, dan nama dinas ditulis di sel gabungan (cuma muncul
sekali di baris pertama tiap kelompok jenis izin). Sistem otomatis
mendeteksi format ini (header punya kolom "Proses" DAN "Pending" DAN
"Berkas Masuk") dan membacanya dengan cara khusus:
- Nama dinas di-"forward-fill" dari baris terakhir yang terisi.
- Seluruh baris dari sheet ini otomatis dianggap **Non Pelaku Usaha**
  (sesuai judul laporannya: "LAPORAN PELAYANAN PERIZINAN NON BERUSAHA...").
- Kategori (jenis pelayanan) dicari dengan mencocokkan nama dinas + nama
  layanan ke sheet MONEV lain di file yang sama (biasanya yang namanya
  mengandung "FIX") -- kalau tidak ketemu pasangannya, kategori jatuh
  balik ke nama jenis izin itu sendiri.
- Karena Proses/Pending diambil langsung dari file (bukan diturunkan),
  `pending_method` untuk periode ini otomatis "eksplisit".

## Format kolom Excel yang dikenali
Aplikasi ini mendukung file **.xlsx maupun .xls (format lama)** -- keduanya
dibaca otomatis sesuai ekstensi file (pakai openpyxl untuk .xlsx, xlrd untuk
.xls). Deteksi kolom juga fleksibel terhadap variasi penulisan header, mis.
"pd_pelaksana" atau "Nama PD pelaksana pelayanan" sama-sama dikenali sebagai
kolom Dinas. Sebagian template (mis. MONEV) punya baris kedua berisi kode
kolom snake_case (pd_pelaksana, jenis_pelayanan, dst) sebelum data sungguhan
dimulai -- baris ini otomatis terdeteksi dan dilewati.

Nama kolom yang biasa dipakai di rekap MONEV Surabaya:
pd_pelaksana / Nama PD pelaksana pelayanan, jenis_pelayanan / Jenis pelayanan,
nama_layanan / Nama Layanan, status_pelaku_usaha / Status Pelaku Usaha,
jumlah_berkas_masuk / Jumlah berkas masuk, jumlah_perstek_selesai_pd_teknis /
Jumlah perstek selesai PD Teknis, jumlah_berkas_selesai_dpmptsp / Jumlah
berkas selesai DPMPTSP, jumlah_berkas_selesai_tepat_waktu_dpmptsp / Jumlah
berkas selesai tepat waktu DPMPTSP, sistem_pelayanan / Sistem Pelayanan.

Pemetaan yang dipakai:
- Dinas          -> pd_pelaksana / Nama PD pelaksana pelayanan
- Jenis Izin     -> nama_layanan / Nama Layanan (fallback: jenis_pelayanan)
- Kategori (pie) -> jenis_pelayanan / Jenis pelayanan
- Berkas Masuk   -> jumlah_berkas_masuk / Jumlah berkas masuk
- Berkas Selesai -> jumlah_berkas_selesai_dpmptsp / Jumlah berkas selesai DPMPTSP (selesai final di DPMPTSP)
- Tepat Waktu    -> jumlah_berkas_selesai_tepat_waktu_dpmptsp / Jumlah berkas selesai tepat waktu DPMPTSP

## Pemisahan Pelaku Usaha / Non Pelaku Usaha
Kalau file punya kolom status_pelaku_usaha / Status Pelaku Usaha (nilai
"Pelaku Usaha" atau "Non Pelaku Usaha"), dashboard otomatis punya toggle di
pojok kanan atas header: **Semua | Pelaku Usaha | Non Pelaku Usaha**. Semua
KPI, chart, dan tabel di semua halaman ikut menyesuaikan sesuai pilihan --
tampilan gabungan ("Semua") tetap tersedia sebagai default.

## Proses vs Pending
Ada 3 cara aplikasi menentukan angka Proses/Pending, dari yang paling
akurat ke paling umum:
1. **eksplisit** -- file punya kolom Proses/Pending sendiri, dipakai langsung.
2. **tahapan** -- file tidak punya kolom Proses/Pending langsung, tapi punya
   kolom "selesai PD Teknis" (jumlah_perstek_selesai_pd_teknis). Dihitung:
   - Proses  = sudah lolos PD Teknis, tapi belum terbit DPMPTSP
   - Pending = belum lolos tahap PD Teknis sama sekali
3. **tidak_ada** -- tidak ada sumber data sama sekali untuk memisahkan;
   seluruh berkas yang belum selesai dihitung sebagai "Proses", Pending = 0.

Dashboard menampilkan catatan otomatis di bagian atas halaman sesuai mode
yang terdeteksi, supaya jelas dasar perhitungan angka Proses/Pending yang
ditampilkan. Kalau tampilan "Gabungan Semua Periode" mencampur periode
dengan mode berbeda-beda, munculnya catatan "campuran" (angka bersifat
perkiraan).

**Pengecualian untuk Pelaku Usaha:** berapa pun mode di atas yang
terdeteksi, baris berstatus **Pelaku Usaha selalu Proses = 0 dan
Pending = 0** -- karena begitu selesai, SK Pelaku Usaha otomatis terbit
(biasanya lewat OSS), jadi tidak ada tahap "menggantung" yang perlu
dilacak seperti Non Pelaku Usaha. Konsekuensinya: kalau ada baris Pelaku
Usaha dengan Berkas Masuk lebih besar dari Berkas Selesai, selisihnya
TIDAK dihitung sebagai backlog di manapun.

## Jumlah SK Terbit
Kartu KPI "Jumlah SK Terbit" = **Berkas Selesai DPMPTSP**, berlaku untuk
SEMUA baris (Pelaku Usaha maupun Non Pelaku Usaha). Kolom
`jumlah_penerbitan_ijin_pd_teknis` di file sumber TIDAK dipakai untuk
metrik ini.

Catatan: sebelumnya sempat dicoba pendekatan "kalau selesai DPMPTSP
kosong tapi kolom penerbitan_ijin terisi, pakai itu sebagai sinyal
selesai" (berguna untuk layanan retribusi/reklame yang kadang tidak
mengisi kolom selesai DPMPTSP) -- tapi pendekatan itu sudah tidak
dipakai lagi sesuai keputusan terbaru: definisi SK Terbit disederhanakan
supaya selalu sama dengan Berkas Selesai DPMPTSP, apa adanya.

Format template lama yang lebih sederhana (Dinas, Jenis Izin, Berkas Masuk,
Berkas Selesai, Berkas Tepat Waktu, Proses, Pending) tetap didukung sebagai
fallback jika kolom-kolom di atas tidak ditemukan.

Baris ringkasan/total di bagian bawah file (baris tanpa nilai pd_pelaksana)
otomatis dilewati saat parsing.

## Catatan
Data disimpan di **Postgres di Supabase** -- ada 2 tabel: `periode` (info
tiap bulan yang diupload) dan `berkas` (baris data Excel, terhubung ke
periode masing-masing lewat foreign key, otomatis ikut terhapus kalau
periode-nya dihapus). Karena Supabase adalah server database terpisah
dari tempat aplikasi Flask-nya berjalan, **data tetap ada** meskipun
aplikasi web-nya di-restart, redeploy, atau bahkan dipindah ke hosting
lain -- selama `DATABASE_URL` yang dipakai tetap menunjuk ke project
Supabase yang sama.

## Database: Supabase

### Kenapa bukan "Direct connection"?
Saat mengambil connection string dari Supabase (Project Settings ->
Database -> Connect), akan ada beberapa pilihan. **Jangan pakai "Direct
connection"** (`db.<project-ref>.supabase.co`) -- hostname itu sering
cuma bisa diakses lewat IPv6, dan banyak jaringan/ISP tidak mendukung
itu, sehingga muncul error `could not translate host name`. Pakai
**"Session pooler"** sebagai gantinya -- hostnya `aws-0-<region>.pooler.
supabase.com`, dan usernamenya berformat `postgres.<project-ref>` (bukan
cuma `postgres`).

### Migrasi data lama (kalau sebelumnya masih pakai dashboard.db SQLite)
Kalau Anda punya `dashboard.db` dari versi sebelum pakai Supabase, jalankan
sekali:
```bash
export DATABASE_URL="<connection string pooler Supabase Anda>"
python migrate_sqlite_to_postgres.py
```
Script ini membuat skema tabel di Postgres (kalau belum ada) lalu
menyalin semua periode & baris data dari `dashboard.db` lama. Aman
dijalankan berkali-kali (data lama di Postgres akan ditimpa bersih dulu,
bukan menumpuk dobel).

## Deploy ke Internet (supaya bisa diakses semua orang)

GitHub sendiri cuma tempat menyimpan kode (version control) -- BUKAN
tempat menjalankan aplikasi web. Supaya dapat URL publik, kodenya perlu
dihubungkan ke platform hosting yang bisa menjalankan Python. Cara
termudah & gratis: **Render.com**.

### 1. Push kode ke GitHub
```bash
cd Proyek_Dashboard
git init
git add .
git commit -m "Dashboard kinerja pelayanan perizinan"
git branch -M main
git remote add origin https://github.com/<username-anda>/<nama-repo>.git
git push -u origin main
```
(Buat dulu repo kosong di github.com/new sebelum menjalankan perintah di atas.
`.gitignore` sudah mengatur supaya `dashboard.db` lama dan isi folder
`uploads/` tidak ikut ter-commit.)

### 2. Hubungkan ke Render
1. Daftar/masuk ke https://render.com (bisa pakai akun GitHub langsung)
2. Klik **New +** -> **Web Service**
3. Pilih repo GitHub yang barusan di-push
4. Render otomatis mendeteksi `render.yaml` untuk Build/Start Command --
   kalau tidak, isi manual:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT`
5. **Wajib**: buka tab **Environment** di service tersebut, tambahkan
   environment variable:
   - Key: `DATABASE_URL`
   - Value: connection string pooler Supabase Anda (sama seperti yang
     dipakai untuk migrasi di atas)
   - Key: `PIN_AKSES` -- Value: PIN login dashboard
   - Key: `SECRET_KEY` -- Value: string acak panjang (lihat contoh di atas)
6. Klik **Create Web Service** -- tunggu beberapa menit sampai status "Live"
7. URL publiknya muncul di bagian atas (format `https://nama-app.onrender.com`)

Karena datanya sekarang di Supabase (bukan file lokal di server Render),
paket gratis Render yang pakai disk sementara (ephemeral) **tidak lagi
jadi masalah** -- server boleh restart/redeploy/tidur kapan saja, data
Anda tetap aman di Supabase.

### PIN & SECRET_KEY
PIN login dan kunci sesi **tidak** ditulis di kode. Keduanya dibaca dari
environment variable `PIN_AKSES` dan `SECRET_KEY` di hosting (Vercel/Render).
Kalau `PIN_AKSES` belum diset, login akan ditolak. Kalau `SECRET_KEY` belum
diset, sesi login bisa sering terputus. Jangan tulis nilainya di repo ini.
