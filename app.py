from flask import (Flask, render_template, request, redirect, url_for,
                   flash, jsonify, Response, session, send_file, abort)
import sqlite3
from datetime import datetime, date, timedelta, timezone
import json, os, io, calendar, math, time
from urllib.parse import urlencode
from hashlib import sha256
from functools import wraps
import logging
from logging.handlers import RotatingFileHandler

# Zona waktu proses (WIB/WITA/WIT) -- lihat apply_tz_app(). Dipanggil sekali di
# sini dgn default aman SEBELUM database ada (perlu terutama di hosting bersama
# spt PythonAnywhere yg server-nya UTC & tak bisa diubah timezone sistemnya
# tanpa akses root), lalu dipanggil ULANG setelah init_db() dengan nilai
# tersimpan di settings (bisa WITA/WIT kalau usaha di luar WIB) -- lihat
# dekat pemanggilan init_db() di bagian bawah file, dan route /settings.
TZ_APP_PILIHAN = {
    'Asia/Jakarta':  'WIB - Jakarta, Jawa, Sumatera, Kalimantan Barat/Tengah',
    'Asia/Makassar': 'WITA - Bali, NTB, NTT, Kalimantan Selatan/Timur/Utara, Sulawesi',
    'Asia/Jayapura': 'WIT - Maluku, Papua',
}
TZ_APP_DEFAULT = 'Asia/Jakarta'

def apply_tz_app(tz_name=None):
    """Terapkan zona waktu proses via TZ env + time.tzset(), berlaku langsung
    utk datetime.now()/date.today() di seluruh app DAN modifier SQL 'localtime'
    (dipakai baca kolom CURRENT_TIMESTAMP yg selalu UTC) -- satu titik ini
    cukup membetulkan semuanya, tanpa perlu restart server. time.tzset() cuma
    ada di POSIX (Linux/Mac); di Windows fungsi ini no-op aman (ikut jam
    Windows apa adanya, seperti sebelum fitur ini ada)."""
    if not hasattr(time, 'tzset'):
        return
    os.environ['TZ'] = tz_name if tz_name in TZ_APP_PILIHAN else TZ_APP_DEFAULT
    time.tzset()

apply_tz_app(TZ_APP_DEFAULT)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

# Resolve base paths, works in dev, --onedir, and --onefile PyInstaller
import sys as _sys
if getattr(_sys, 'frozen', False):
    _BUNDLE = _sys._MEIPASS                        # read-only temp extraction dir
    _DATA   = os.path.dirname(_sys.executable)     # next to EXE, writable, persistent
else:
    _BUNDLE = os.path.dirname(os.path.abspath(__file__))
    _DATA   = _BUNDLE

app = Flask(__name__,
    template_folder=os.path.join(_BUNDLE, 'templates'),
    static_folder=os.path.join(_BUNDLE, 'static'))
app.config['TEMPLATES_AUTO_RELOAD'] = True
_key_file = os.path.join(_DATA, '.secret_key')
if not os.path.exists(_key_file):
    with open(_key_file, 'wb') as _f:
        _f.write(os.urandom(32))
with open(_key_file, 'rb') as _f:
    app.secret_key = _f.read()
DB = os.path.join(_DATA, 'finansial.db')
UPLOAD_FOLDER = os.path.join(_DATA, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
ALLOWED_IMG = {'png', 'jpg', 'jpeg', 'gif', 'webp', 'svg'}

# ── MODE DEMO (kunci password) ────────────────────────────────────────
# Untuk instance demo publik (mis. finance.cliffbiz.com): password akun
# TIDAK dapat diubah, supaya login demo (admin/admin123) tetap stabil.
# Aktif bila env DEMO_LOCK=1 ATAU ada file '.demo_lock' di folder data.
# Default OFF, jadi distribusi/instance lain tidak terpengaruh.
DEMO_LOCK = (os.environ.get('DEMO_LOCK', '').strip() == '1'
             or os.path.exists(os.path.join(_DATA, '.demo_lock')))

# ── PEMULIHAN PASSWORD (khusus aplikasi desktop) ──────────────────────
# Banyak pengguna desktop lupa password dan datanya jadi tidak bisa dibuka
# sama sekali, padahal file finansial.db ada di komputer mereka sendiri.
# Halaman login menyediakan "Lupa password?" yang mengembalikan SEMUA login
# ke bawaan, TANPA menyentuh satu pun transaksi.
#
# HANYA aktif saat berjalan sebagai aplikasi desktop (EXE hasil PyInstaller).
# Ini WAJIB, karena app.py yang sama dipakai semua tenant online di server:
# tanpa gerbang ini, siapa pun yang membuka halaman login tenant bisa
# mereset password semua orang. Gerbangnya sengaja "mati kecuali terbukti
# desktop", bukan sebaliknya.
#
# Env RESET_PASSWORD_LOKAL=1 hanya untuk pengujian otomatis.
DESKTOP_APP = bool(getattr(_sys, 'frozen', False))
RESET_LOKAL_AKTIF = (DESKTOP_APP
                     or os.environ.get('RESET_PASSWORD_LOKAL', '').strip() == '1')

# Simpan traceback error runtime di sebelah app/EXE agar 500 di build dist mudah dilacak.
_log_file = os.path.join(_DATA, 'error.log')
_file_handler = RotatingFileHandler(_log_file, maxBytes=1_000_000, backupCount=3, encoding='utf-8')
_file_handler.setLevel(logging.ERROR)
_file_handler.setFormatter(logging.Formatter(
    '%(asctime)s %(levelname)s [%(pathname)s:%(lineno)d] %(message)s'
))
if not any(isinstance(h, RotatingFileHandler) and getattr(h, 'baseFilename', None) == _log_file for h in app.logger.handlers):
    app.logger.addHandler(_file_handler)
app.logger.setLevel(logging.INFO)

# Serve user-uploaded files from persistent DATA dir (supports --onefile bundle)
from flask import send_from_directory as _send_from_dir
@app.route('/static/uploads/<path:filename>')
def _serve_upload(filename):
    return _send_from_dir(UPLOAD_FOLDER, filename)

BULAN = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agt','Sep','Okt','Nov','Des']

# ---------- DB ----------
def db():
    conn = sqlite3.connect(DB, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

# ── Bagan Akun bawaan (COA default) ───────────────────────────────────────────
# Dipakai init_db() untuk seed & backfill, DAN sebagai daftar akun yang tidak boleh
# dihapus / diubah kodenya lewat menu Daftar Akun. Alasannya: kode-kode ini dirujuk
# langsung oleh mesin jurnal (mis. 1130 Persediaan dipakai semua transaksi barang).
# Kalau salah satunya hilang, penyimpanan transaksi gagal dengan 'Akun tidak ditemukan'.
AKUN_DEFAULT = [
            ('1100','Kas','ASET','Aset Lancar','DEBIT'),
            ('1110','Bank','ASET','Aset Lancar','DEBIT'),
            ('1120','Piutang Usaha','ASET','Aset Lancar','DEBIT'),
            ('1130','Persediaan Barang','ASET','Aset Lancar','DEBIT'),
            ('1140','Biaya Dibayar Dimuka','ASET','Aset Lancar','DEBIT'),
            ('1150','Cadangan Kerugian Piutang','ASET','Aset Lancar','KREDIT'),
            ('1160','Piutang Karyawan','ASET','Aset Lancar','DEBIT'),
            ('1180','PPN Masukan (Dapat Dikreditkan)','ASET','Aset Lancar','DEBIT'),
            ('1181','PPh 23 Dibayar Dimuka','ASET','Aset Lancar','DEBIT'),
            ('1182','Angsuran PPh Pasal 25','ASET','Aset Lancar','DEBIT'),
            ('1183','PPh 22 Dibayar Dimuka','ASET','Aset Lancar','DEBIT'),
            ('1200','Peralatan','ASET','Aset Tetap','DEBIT'),
            ('1210','Kendaraan','ASET','Aset Tetap','DEBIT'),
            ('1220','Gedung & Bangunan','ASET','Aset Tetap','DEBIT'),
            ('1290','Akumulasi Penyusutan','ASET','Aset Tetap','KREDIT'),
            ('2100','Hutang Usaha','LIABILITAS','Liabilitas Lancar','KREDIT'),
            ('2110','Hutang Pajak','LIABILITAS','Liabilitas Lancar','KREDIT'),
            ('2111','Hutang PPN','LIABILITAS','Liabilitas Lancar','KREDIT'),
            ('2112','Hutang PPh Pasal 23','LIABILITAS','Liabilitas Lancar','KREDIT'),
            ('2120','Hutang Gaji','LIABILITAS','Liabilitas Lancar','KREDIT'),
            ('2130','Pendapatan Diterima Dimuka','LIABILITAS','Liabilitas Lancar','KREDIT'),
            ('2200','Hutang Bank','LIABILITAS','Liabilitas Jangka Panjang','KREDIT'),
            ('3100','Modal Pemilik','EKUITAS','Ekuitas','KREDIT'),
            ('3200','Laba Ditahan','EKUITAS','Ekuitas','KREDIT'),
            ('3300','Prive / Penarikan Owner','EKUITAS','Ekuitas','DEBIT'),
            ('4100','Pendapatan Penjualan','PENDAPATAN','Pendapatan Usaha','KREDIT'),
            ('4150','Retur Penjualan','PENDAPATAN','Pendapatan Usaha','DEBIT'),
            ('4200','Pendapatan Jasa','PENDAPATAN','Pendapatan Usaha','KREDIT'),
            ('4300','Pendapatan Lain-lain','PENDAPATAN','Pendapatan Lain','KREDIT'),
            ('5100','Harga Pokok Penjualan','BEBAN','HPP','DEBIT'),
            ('5150','Retur Pembelian','BEBAN','HPP','KREDIT'),
            ('5190','Koreksi Persediaan','BEBAN','HPP','DEBIT'),
            ('6100','Beban Gaji','BEBAN','Beban Operasional','DEBIT'),
            ('6110','Beban Sewa','BEBAN','Beban Operasional','DEBIT'),
            ('6120','Beban Listrik & Air','BEBAN','Beban Operasional','DEBIT'),
            ('6130','Beban Penyusutan','BEBAN','Beban Penyusutan','DEBIT'),
            ('6140','Beban Pemasaran','BEBAN','Beban Operasional','DEBIT'),
            ('6150','Beban Administrasi','BEBAN','Beban Operasional','DEBIT'),
            ('6160','Beban Bunga','BEBAN','Beban Bunga','DEBIT'),
            ('6170','Beban Pajak','BEBAN','Beban Pajak','DEBIT'),
            ('6180','Beban Lainnya','BEBAN','Beban Operasional','DEBIT'),
            ('6190','Beban Kerugian Piutang','BEBAN','Beban Operasional','DEBIT'),
]
AKUN_DEFAULT_KODE = frozenset(a[0] for a in AKUN_DEFAULT)

def _migrasi_fk_penjualan_resep_bahan(conn):
    """Pasang foreign key jurnal_id pada penjualan_resep_bahan (database LAMA).

    Database baru sudah mendapat FK ini lewat CREATE TABLE. Database lama dibuat
    sebelum FK itu ada, dan SQLite TIDAK punya `ALTER TABLE ADD CONSTRAINT`, jadi
    satu-satunya cara adalah membangun ulang tabelnya: buat tabel baru ber-FK,
    salin isinya, buang yang lama, ganti nama.

    SENGAJA DILEWATI kalau masih ada baris yatim (jurnal induknya sudah tidak ada).
    Alasannya: menyalin baris yatim ke tabel ber-FK pasti gagal, dan satu-satunya
    cara memaksanya lewat adalah MENGHAPUS baris tersebut. Baris yatim di database
    yang sudah berjalan bukan sekadar sampah -- sebagian bisa menempel ke jurnal
    aktif dan perlu ditinjau manusia lebih dulu. Migrasi tidak boleh diam-diam
    menghapus data pembukuan. Jadi: bersihkan dulu, FK menyusul otomatis di
    startup berikutnya.

    Idempoten: kalau FK sudah ada, langsung keluar.
    """
    try:
        # Sudah punya FK? (cek daftar foreign key tabel)
        for fk in conn.execute("PRAGMA foreign_key_list(penjualan_resep_bahan)"):
            if str(fk['table']).lower() == 'jurnal':
                return
        yatim = conn.execute(
            "SELECT COUNT(*) c FROM penjualan_resep_bahan p "
            "WHERE NOT EXISTS (SELECT 1 FROM jurnal j WHERE j.id = p.jurnal_id)"
        ).fetchone()['c']
        if yatim:
            app.logger.warning(
                'FK penjualan_resep_bahan DILEWATI: masih ada %s baris yatim. '
                'Bersihkan dulu, FK akan dipasang otomatis saat startup berikutnya.', yatim)
            return
        # Rebuild. foreign_keys harus OFF selama rename supaya FK tabel LAIN tidak
        # ikut divalidasi setengah jalan; dinyalakan lagi setelah selesai.
        # conn.commit() WAJIB duluan: init_db sudah menjalankan DDL/DML sebelum ini,
        # jadi masih ada transaksi implisit yang terbuka. Tanpa menutupnya, BEGIN
        # gagal dengan "cannot start a transaction within a transaction", dan PRAGMA
        # foreign_keys juga diabaikan diam-diam kalau dipanggil di dalam transaksi.
        conn.commit()
        conn.execute('PRAGMA foreign_keys = OFF')
        conn.execute('BEGIN')
        conn.execute("""
            CREATE TABLE penjualan_resep_bahan_baru (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                jurnal_id INTEGER NOT NULL,
                template_id INTEGER,
                line_idx INTEGER DEFAULT 0,
                produk_id INTEGER,
                deskripsi TEXT DEFAULT '',
                menu_qty REAL DEFAULT 0,
                qty REAL DEFAULT 0,
                harga REAL DEFAULT 0,
                subtotal REAL DEFAULT 0,
                qty_returned REAL DEFAULT 0,
                FOREIGN KEY (jurnal_id) REFERENCES jurnal(id) ON DELETE CASCADE
            )""")
        conn.execute("""
            INSERT INTO penjualan_resep_bahan_baru
                (id,jurnal_id,template_id,line_idx,produk_id,deskripsi,
                 menu_qty,qty,harga,subtotal,qty_returned)
            SELECT id,jurnal_id,template_id,line_idx,produk_id,deskripsi,
                   menu_qty,qty,harga,subtotal,qty_returned
              FROM penjualan_resep_bahan""")
        conn.execute('DROP TABLE penjualan_resep_bahan')
        conn.execute('ALTER TABLE penjualan_resep_bahan_baru RENAME TO penjualan_resep_bahan')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_prb_jurnal ON penjualan_resep_bahan(jurnal_id)')
        conn.commit()
        conn.execute('PRAGMA foreign_keys = ON')
        app.logger.info('FK penjualan_resep_bahan.jurnal_id terpasang.')
    except Exception as ex:
        # Gagal = biarkan tabel lama apa adanya. App tetap jalan seperti sebelumnya.
        try:
            conn.rollback()
            conn.execute('PRAGMA foreign_keys = ON')
        except Exception:
            pass
        app.logger.warning('Migrasi FK penjualan_resep_bahan gagal (dilewati): %s', ex)


def init_db():
    conn = db()
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY, value TEXT
    );
    CREATE TABLE IF NOT EXISTS tx_sequence (
        periode TEXT PRIMARY KEY,
        last_number INTEGER NOT NULL DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        nama TEXT,
        role TEXT NOT NULL DEFAULT 'VIEWER',
        aktif INTEGER DEFAULT 1,
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS akun (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kode TEXT UNIQUE NOT NULL,
        nama TEXT NOT NULL,
        tipe TEXT NOT NULL,
        subtipe TEXT,
        saldo_normal TEXT NOT NULL DEFAULT 'DEBIT'
    );
    CREATE TABLE IF NOT EXISTS jurnal (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tanggal DATE NOT NULL,
        keterangan TEXT NOT NULL,
        referensi TEXT,
        kategori TEXT DEFAULT 'OPERASIONAL',
        tipe_tx TEXT DEFAULT 'JURNAL',
        nomor_tx TEXT,
        proyek_id INTEGER,
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS detail_jurnal (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jurnal_id INTEGER NOT NULL,
        akun_id INTEGER NOT NULL,
        debit REAL DEFAULT 0,
        kredit REAL DEFAULT 0,
        FOREIGN KEY (jurnal_id) REFERENCES jurnal(id) ON DELETE CASCADE,
        FOREIGN KEY (akun_id) REFERENCES akun(id)
    );
    CREATE TABLE IF NOT EXISTS produk (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kode TEXT UNIQUE NOT NULL,
        nama TEXT NOT NULL,
        varian TEXT,
        satuan TEXT DEFAULT 'pcs',
        harga_beli REAL DEFAULT 0,
        harga_jual REAL DEFAULT 0,
        stok REAL DEFAULT 0,
        min_stok REAL DEFAULT 0,
        aktif INTEGER DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS produk_harga (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        produk_id INTEGER NOT NULL,
        label TEXT NOT NULL,
        harga REAL DEFAULT 0,
        FOREIGN KEY (produk_id) REFERENCES produk(id)
    );
    CREATE INDEX IF NOT EXISTS idx_produk_harga_pid ON produk_harga(produk_id);
    CREATE TABLE IF NOT EXISTS produksi_template_harga (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        template_id INTEGER NOT NULL,
        label TEXT NOT NULL,
        harga REAL DEFAULT 0,
        FOREIGN KEY (template_id) REFERENCES produksi_template(id)
    );
    CREATE INDEX IF NOT EXISTS idx_produksi_template_harga_tid ON produksi_template_harga(template_id);
    CREATE TABLE IF NOT EXISTS pergerakan_stok (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        produk_id INTEGER NOT NULL,
        tanggal DATE NOT NULL,
        jenis TEXT NOT NULL,
        qty REAL NOT NULL,
        harga REAL DEFAULT 0,
        keterangan TEXT,
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (produk_id) REFERENCES produk(id)
    );
    CREATE TABLE IF NOT EXISTS pergerakan_persediaan_non_sku (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jurnal_id INTEGER NOT NULL,
        tanggal DATE NOT NULL,
        deskripsi TEXT NOT NULL,
        nilai_delta REAL NOT NULL,
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (jurnal_id) REFERENCES jurnal(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS piutang (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tanggal DATE NOT NULL,
        jatuh_tempo DATE,
        pelanggan TEXT NOT NULL,
        keterangan TEXT,
        jumlah REAL NOT NULL,
        terbayar REAL DEFAULT 0,
        status TEXT DEFAULT 'BELUM LUNAS',
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS hutang (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tanggal DATE NOT NULL,
        jatuh_tempo DATE,
        pemasok TEXT NOT NULL,
        keterangan TEXT,
        jumlah REAL NOT NULL,
        terbayar REAL DEFAULT 0,
        status TEXT DEFAULT 'BELUM LUNAS',
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS customer (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nama TEXT UNIQUE NOT NULL,
        telepon TEXT DEFAULT '',
        alamat TEXT DEFAULT '',
        catatan TEXT DEFAULT '',
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS vendor (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nama TEXT UNIQUE NOT NULL,
        telepon TEXT DEFAULT '',
        alamat TEXT DEFAULT '',
        catatan TEXT DEFAULT '',
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS transaksi_item (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jurnal_id INTEGER NOT NULL,
        produk_id INTEGER,
        deskripsi TEXT DEFAULT '',
        qty REAL DEFAULT 1,
        satuan TEXT DEFAULT '',
        harga REAL DEFAULT 0,
        diskon REAL DEFAULT 0,
        subtotal REAL DEFAULT 0,
        arah TEXT DEFAULT 'NONE',
        FOREIGN KEY (jurnal_id) REFERENCES jurnal(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS pos_shift (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        username TEXT NOT NULL,
        dibuka DATETIME DEFAULT CURRENT_TIMESTAMP,
        ditutup DATETIME,
        saldo_awal REAL DEFAULT 0,
        kas_sistem REAL DEFAULT 0,
        kas_fisik REAL DEFAULT 0,
        selisih REAL DEFAULT 0,
        status TEXT DEFAULT 'OPEN',
        catatan TEXT DEFAULT '',
        FOREIGN KEY (user_id) REFERENCES users(id)
    );
    CREATE TABLE IF NOT EXISTS pos_sale (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jurnal_id INTEGER NOT NULL,
        shift_id INTEGER,
        nomor_struk TEXT UNIQUE NOT NULL,
        tanggal DATE NOT NULL,
        customer TEXT DEFAULT '',
        metode_bayar TEXT DEFAULT 'TUNAI',
        akun_kas TEXT DEFAULT '1100',
        total REAL DEFAULT 0,
        bayar REAL DEFAULT 0,
        kembalian REAL DEFAULT 0,
        status TEXT DEFAULT 'SELESAI',
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (jurnal_id) REFERENCES jurnal(id) ON DELETE CASCADE,
        FOREIGN KEY (shift_id) REFERENCES pos_shift(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS voucher (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kode TEXT UNIQUE NOT NULL,
        tipe TEXT DEFAULT 'NOMINAL',
        nilai REAL DEFAULT 0,
        min_belanja REAL DEFAULT 0,
        maks_potongan REAL DEFAULT 0,
        kuota INTEGER DEFAULT 0,
        terpakai INTEGER DEFAULT 0,
        berlaku_dari DATE,
        berlaku_sampai DATE,
        aktif INTEGER DEFAULT 1,
        catatan TEXT DEFAULT '',
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS voucher_pakai (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        voucher_id INTEGER NOT NULL,
        jurnal_id INTEGER,
        tanggal DATE NOT NULL,
        potongan REAL DEFAULT 0,
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (voucher_id) REFERENCES voucher(id) ON DELETE CASCADE,
        FOREIGN KEY (jurnal_id) REFERENCES jurnal(id) ON DELETE SET NULL
    );
    -- Scope per-produk: voucher hanya menghitung potongan dari subtotal produk berikut.
    CREATE TABLE IF NOT EXISTS voucher_produk (
        voucher_id INTEGER NOT NULL,
        produk_id INTEGER NOT NULL,
        PRIMARY KEY (voucher_id, produk_id),
        FOREIGN KEY (voucher_id) REFERENCES voucher(id) ON DELETE CASCADE,
        FOREIGN KEY (produk_id) REFERENCES produk(id) ON DELETE CASCADE
    );
    -- Pembatasan per-customer (orthogonal terhadap scope): voucher hanya berlaku
    -- untuk pelanggan dengan nama berikut (cocok case-insensitive). Kosong = semua.
    CREATE TABLE IF NOT EXISTS voucher_customer (
        voucher_id INTEGER NOT NULL,
        customer_nama TEXT NOT NULL,
        PRIMARY KEY (voucher_id, customer_nama),
        FOREIGN KEY (voucher_id) REFERENCES voucher(id) ON DELETE CASCADE
    );
    -- Scope multi-kategori: voucher KATEGORI bisa mencakup beberapa kategori sekaligus
    -- (potongan dihitung dari subtotal produk di SEMUA kategori terpilih).
    CREATE TABLE IF NOT EXISTS voucher_kategori (
        voucher_id INTEGER NOT NULL,
        kategori TEXT NOT NULL,
        PRIMARY KEY (voucher_id, kategori),
        FOREIGN KEY (voucher_id) REFERENCES voucher(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS bayar_piutang (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        piutang_id INTEGER NOT NULL,
        tanggal DATE NOT NULL,
        jumlah REAL NOT NULL,
        catatan TEXT,
        FOREIGN KEY (piutang_id) REFERENCES piutang(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS bayar_hutang (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hutang_id INTEGER NOT NULL,
        tanggal DATE NOT NULL,
        jumlah REAL NOT NULL,
        catatan TEXT,
        FOREIGN KEY (hutang_id) REFERENCES hutang(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS penyesuaian_tagihan (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jurnal_id INTEGER NOT NULL,
        jenis TEXT NOT NULL,
        record_id INTEGER NOT NULL,
        jumlah_delta REAL NOT NULL,
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (jurnal_id) REFERENCES jurnal(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS writeoff_piutang (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        piutang_id INTEGER NOT NULL,
        jurnal_id INTEGER NOT NULL,
        tanggal DATE NOT NULL,
        jumlah REAL NOT NULL,
        metode TEXT DEFAULT 'LANGSUNG',
        alasan TEXT DEFAULT '',
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (piutang_id) REFERENCES piutang(id) ON DELETE CASCADE,
        FOREIGN KEY (jurnal_id) REFERENCES jurnal(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS aset_tetap (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nama TEXT NOT NULL,
        kategori TEXT,
        harga_beli REAL NOT NULL,
        tanggal_beli DATE NOT NULL,
        masa_pakai INTEGER NOT NULL,
        penyusutan_bulan REAL,
        aktif INTEGER DEFAULT 1,
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    """)
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS invoice (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jurnal_id INTEGER,
        nomor TEXT UNIQUE NOT NULL,
        tanggal DATE NOT NULL,
        jatuh_tempo DATE,
        top_hari INTEGER,
        top_note TEXT DEFAULT '',
        pelanggan TEXT NOT NULL,
        alamat_pelanggan TEXT DEFAULT '',
        telepon_pelanggan TEXT DEFAULT '',
        diskon REAL DEFAULT 0,
        ongkir REAL DEFAULT 0,
        biaya_lain REAL DEFAULT 0,
        catatan TEXT DEFAULT '',
        status TEXT DEFAULT 'DRAFT',
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (jurnal_id) REFERENCES jurnal(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS invoice_item (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_id INTEGER NOT NULL,
        deskripsi TEXT NOT NULL,
        qty REAL DEFAULT 1,
        satuan TEXT DEFAULT 'pcs',
        harga_satuan REAL DEFAULT 0,
        diskon_item REAL DEFAULT 0,
        subtotal REAL DEFAULT 0,
        FOREIGN KEY (invoice_id) REFERENCES invoice(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS purchase_order (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nomor TEXT UNIQUE NOT NULL,
        tanggal DATE NOT NULL,
        expected_date DATE,
        vendor TEXT NOT NULL,
        alamat_vendor TEXT DEFAULT '',
        telepon_vendor TEXT DEFAULT '',
        diskon REAL DEFAULT 0,
        ongkir REAL DEFAULT 0,
        biaya_lain REAL DEFAULT 0,
        catatan TEXT DEFAULT '',
        status TEXT DEFAULT 'DRAFT',
        hutang_id INTEGER,
        jurnal_id INTEGER,
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS po_item (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        po_id INTEGER NOT NULL,
        produk_id INTEGER,
        deskripsi TEXT NOT NULL,
        qty REAL DEFAULT 1,
        satuan TEXT DEFAULT 'pcs',
        harga_satuan REAL DEFAULT 0,
        diskon_item REAL DEFAULT 0,
        subtotal REAL DEFAULT 0,
        FOREIGN KEY (po_id) REFERENCES purchase_order(id) ON DELETE CASCADE,
        FOREIGN KEY (produk_id) REFERENCES produk(id)
    );
    CREATE TABLE IF NOT EXISTS hpp_produk (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nama TEXT NOT NULL,
        jenis TEXT DEFAULT 'FNB_MENU',
        satuan TEXT DEFAULT 'porsi',
        harga_jual REAL DEFAULT 0,
        bahan TEXT DEFAULT '[]',
        updated_at TEXT DEFAULT (date('now'))
    );
    CREATE TABLE IF NOT EXISTS produksi_template (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nama TEXT NOT NULL,
        produk_jadi_id INTEGER,
        qty_hasil REAL DEFAULT 0,
        biaya_tambahan REAL DEFAULT 0,
        bahan_sku TEXT DEFAULT '[]',
        bahan_non_sku TEXT DEFAULT '[]',
        updated_at TEXT DEFAULT (date('now'))
    );
    CREATE TABLE IF NOT EXISTS log_aktivitas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        waktu DATETIME DEFAULT CURRENT_TIMESTAMP,
        user_id INTEGER,
        username TEXT,
        aksi TEXT NOT NULL,
        detail TEXT
    );
    CREATE TABLE IF NOT EXISTS produksi (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tanggal DATE NOT NULL,
        produk_jadi_id INTEGER NOT NULL,
        qty_hasil REAL NOT NULL,
        biaya_tambahan REAL DEFAULT 0,
        total_hpp REAL DEFAULT 0,
        hpp_per_unit REAL DEFAULT 0,
        keterangan TEXT DEFAULT '',
        jurnal_id INTEGER,
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (produk_jadi_id) REFERENCES produk(id),
        FOREIGN KEY (jurnal_id) REFERENCES jurnal(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS produksi_bahan (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        produksi_id INTEGER NOT NULL,
        produk_id INTEGER NOT NULL,
        qty REAL NOT NULL,
        harga REAL DEFAULT 0,
        subtotal REAL DEFAULT 0,
        FOREIGN KEY (produksi_id) REFERENCES produksi(id) ON DELETE CASCADE,
        FOREIGN KEY (produk_id) REFERENCES produk(id)
    );
    CREATE TABLE IF NOT EXISTS produksi_non_sku (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        produksi_id INTEGER NOT NULL,
        deskripsi TEXT NOT NULL,
        nilai REAL DEFAULT 0,
        qty REAL DEFAULT 0,
        satuan TEXT DEFAULT '',
        FOREIGN KEY (produksi_id) REFERENCES produksi(id) ON DELETE CASCADE
    );
    -- Backflush resep saat Penjualan Umum: subledger AUDIT konsumsi bahan per baris menu.
    -- BUKAN ledger reversal (stok tetap dibalik via pergerakan_stok). Key by jurnal_id.
    -- produk_id NULL = bahan non-SKU (nilai); terisi = bahan SKU berstok (diskriminator).
    CREATE TABLE IF NOT EXISTS penjualan_resep_bahan (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jurnal_id INTEGER NOT NULL,
        template_id INTEGER,
        line_idx INTEGER DEFAULT 0,
        produk_id INTEGER,
        deskripsi TEXT DEFAULT '',
        menu_qty REAL DEFAULT 0,
        qty REAL DEFAULT 0,
        harga REAL DEFAULT 0,
        subtotal REAL DEFAULT 0,
        qty_returned REAL DEFAULT 0,
        -- Dulu kolom jurnal_id ini TIDAK punya foreign key sama sekali. Akibatnya barisnya
        -- bisa tertinggal saat jurnal induknya hilang, lalu "menempel" ke jurnal baru yang
        -- kebetulan ber-ID sama dan terbaca sebagai milik transaksi yang tak bersangkutan.
        -- CASCADE membuat baris ini selalu ikut terhapus bersama jurnalnya.
        FOREIGN KEY (jurnal_id) REFERENCES jurnal(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS anggaran (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        periode_bulan INTEGER NOT NULL,
        periode_tahun INTEGER NOT NULL,
        akun_kode TEXT NOT NULL,
        nominal REAL DEFAULT 0,
        catatan TEXT DEFAULT '',
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(periode_bulan, periode_tahun, akun_kode)
    );
    CREATE TABLE IF NOT EXISTS target_pendapatan (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        periode_bulan INTEGER NOT NULL,
        periode_tahun INTEGER NOT NULL,
        nominal REAL DEFAULT 0,
        catatan TEXT DEFAULT '',
        UNIQUE(periode_bulan, periode_tahun)
    );
    CREATE TABLE IF NOT EXISTS karyawan (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        no_karyawan TEXT DEFAULT '',
        nama TEXT NOT NULL,
        jabatan TEXT DEFAULT '',
        status_kerja TEXT DEFAULT 'TETAP',
        status_pajak TEXT DEFAULT 'TK0',
        gaji_pokok REAL DEFAULT 0,
        tunjangan_tetap REAL DEFAULT 0,
        bpjs_kesehatan INTEGER DEFAULT 1,
        bpjs_ketenagakerjaan INTEGER DEFAULT 1,
        npwp TEXT DEFAULT '',
        aktif INTEGER DEFAULT 1,
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS payroll (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        karyawan_id INTEGER NOT NULL,
        periode_bulan INTEGER NOT NULL,
        periode_tahun INTEGER NOT NULL,
        gaji_pokok REAL DEFAULT 0,
        tunjangan REAL DEFAULT 0,
        penghasilan_bruto REAL DEFAULT 0,
        bpjs_kes_kar REAL DEFAULT 0,
        bpjs_tk_kar REAL DEFAULT 0,
        bpjs_kes_prs REAL DEFAULT 0,
        bpjs_tk_prs REAL DEFAULT 0,
        pph21 REAL DEFAULT 0,
        gaji_bersih REAL DEFAULT 0,
        jurnal_id INTEGER,
        catatan TEXT DEFAULT '',
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (karyawan_id) REFERENCES karyawan(id),
        FOREIGN KEY (jurnal_id) REFERENCES jurnal(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS proyek (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kode TEXT UNIQUE NOT NULL,
        nama TEXT NOT NULL,
        pelanggan TEXT DEFAULT '',
        nilai_kontrak REAL DEFAULT 0,
        anggaran_biaya REAL DEFAULT 0,
        tanggal_mulai DATE,
        target_selesai DATE,
        status TEXT DEFAULT 'AKTIF',
        catatan TEXT DEFAULT '',
        dibuat DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    """)
    # migrations
    for col_sql in [
        "ALTER TABLE users ADD COLUMN permissions TEXT",
        "ALTER TABLE jurnal ADD COLUMN tipe_tx TEXT DEFAULT 'JURNAL'",
        "ALTER TABLE jurnal ADD COLUMN nomor_tx TEXT",
        "ALTER TABLE produk ADD COLUMN varian TEXT",
        "ALTER TABLE akun ADD COLUMN is_rekening INTEGER DEFAULT 0",
        "ALTER TABLE akun ADD COLUMN no_rekening TEXT",
        "ALTER TABLE aset_tetap ADD COLUMN bulan_penyusutan_dicatat INTEGER DEFAULT 0",
        "ALTER TABLE aset_tetap ADD COLUMN akumulasi_penyusutan REAL DEFAULT 0",
        "ALTER TABLE aset_tetap ADD COLUMN jurnal_id INTEGER",
        "ALTER TABLE jurnal ADD COLUMN aset_id INTEGER",
        "ALTER TABLE log_aktivitas ADD COLUMN kategori TEXT DEFAULT 'LAINNYA'",
        "ALTER TABLE log_aktivitas ADD COLUMN role TEXT",
        "ALTER TABLE produk ADD COLUMN kategori TEXT DEFAULT ''",
        "ALTER TABLE produk ADD COLUMN pos_label TEXT DEFAULT ''",
        "ALTER TABLE invoice ADD COLUMN top_hari INTEGER",
        "ALTER TABLE invoice ADD COLUMN top_note TEXT DEFAULT ''",
        "ALTER TABLE invoice ADD COLUMN jurnal_id INTEGER",
        "ALTER TABLE jurnal ADD COLUMN pihak TEXT DEFAULT ''",
        "ALTER TABLE jurnal ADD COLUMN tx_meta TEXT DEFAULT ''",
        "ALTER TABLE pergerakan_stok ADD COLUMN jurnal_id INTEGER",
        "ALTER TABLE piutang ADD COLUMN jurnal_id INTEGER",
        "ALTER TABLE hutang ADD COLUMN jurnal_id INTEGER",
        "ALTER TABLE pergerakan_stok ADD COLUMN hpp_effect TEXT DEFAULT ''",
        "ALTER TABLE pergerakan_stok ADD COLUMN stok_sebelum REAL",
        "ALTER TABLE pergerakan_stok ADD COLUMN hpp_sebelum REAL",
        "ALTER TABLE pergerakan_stok ADD COLUMN stok_sesudah REAL",
        "ALTER TABLE pergerakan_stok ADD COLUMN hpp_sesudah REAL",
        "ALTER TABLE bayar_piutang ADD COLUMN jurnal_id INTEGER",
        "ALTER TABLE bayar_hutang ADD COLUMN jurnal_id INTEGER",
        "ALTER TABLE hutang ADD COLUMN akun_kode TEXT DEFAULT '2100'",
        "ALTER TABLE hpp_produk ADD COLUMN jenis TEXT DEFAULT 'FNB_MENU'",
        "ALTER TABLE hpp_produk ADD COLUMN satuan TEXT DEFAULT 'porsi'",
        "ALTER TABLE produksi_non_sku ADD COLUMN qty REAL DEFAULT 0",
        "ALTER TABLE produksi_non_sku ADD COLUMN satuan TEXT DEFAULT ''",
        "ALTER TABLE invoice ADD COLUMN pajak_persen REAL DEFAULT 0",
        "ALTER TABLE invoice ADD COLUMN pajak_nominal REAL DEFAULT 0",
        "ALTER TABLE invoice ADD COLUMN pph23_persen REAL DEFAULT 0",
        "ALTER TABLE invoice ADD COLUMN pph23_nominal REAL DEFAULT 0",
        # PPh 22 (1,5% pungutan instansi/bendaharawan saat membayar penjual)
        "ALTER TABLE invoice ADD COLUMN pph22_persen REAL DEFAULT 0",
        "ALTER TABLE invoice ADD COLUMN pph22_nominal REAL DEFAULT 0",
        "ALTER TABLE piutang ADD COLUMN invoice_id INTEGER",
        "ALTER TABLE customer ADD COLUMN telepon TEXT DEFAULT ''",
        "ALTER TABLE customer ADD COLUMN alamat TEXT DEFAULT ''",
        "ALTER TABLE customer ADD COLUMN catatan TEXT DEFAULT ''",
        "ALTER TABLE vendor ADD COLUMN telepon TEXT DEFAULT ''",
        "ALTER TABLE vendor ADD COLUMN alamat TEXT DEFAULT ''",
        "ALTER TABLE vendor ADD COLUMN catatan TEXT DEFAULT ''",
        "ALTER TABLE hpp_produk ADD COLUMN updated_at TEXT DEFAULT (date('now'))",
        # Gaji expanded: jenis upah, tarif non-tetap, tanggungan (info), bonus
        "ALTER TABLE karyawan ADD COLUMN jenis_upah TEXT DEFAULT 'BULANAN'",
        "ALTER TABLE karyawan ADD COLUMN tarif_harian REAL DEFAULT 0",
        "ALTER TABLE karyawan ADD COLUMN jumlah_tanggungan INTEGER DEFAULT 0",
        "ALTER TABLE payroll ADD COLUMN jenis_upah TEXT DEFAULT 'BULANAN'",
        "ALTER TABLE payroll ADD COLUMN bonus REAL DEFAULT 0",
        "ALTER TABLE payroll ADD COLUMN hari_kerja REAL DEFAULT 0",
        "ALTER TABLE payroll ADD COLUMN qty_output REAL DEFAULT 0",
        # Variable cost per periode: tunjangan tidak tetap (teratur) + lain-lain (tidak teratur)
        "ALTER TABLE payroll ADD COLUMN uang_makan REAL DEFAULT 0",
        "ALTER TABLE payroll ADD COLUMN uang_transport REAL DEFAULT 0",
        "ALTER TABLE payroll ADD COLUMN lain_lain REAL DEFAULT 0",
        # Proyek (job costing): tag transaksi ke proyek
        "ALTER TABLE jurnal ADD COLUMN proyek_id INTEGER",
        # POS: produk favorit (akses cepat di kasir)
        "ALTER TABLE produk ADD COLUMN favorit INTEGER DEFAULT 0",
        # Arsip produk: produk dengan riwayat tidak dihapus fisik, hanya disembunyikan dari list aktif.
        "ALTER TABLE produk ADD COLUMN aktif INTEGER DEFAULT 1",
        # PO item dapat tertaut ke SKU agar penerimaan PO menambah stok produk, bukan selalu non-SKU.
        "ALTER TABLE po_item ADD COLUMN produk_id INTEGER",
        # Voucher: mode scope (NOTA = seluruh nota; PRODUK = subtotal produk terpilih; KATEGORI = subtotal kategori)
        "ALTER TABLE voucher ADD COLUMN scope TEXT DEFAULT 'NOTA'",
        "ALTER TABLE voucher ADD COLUMN syarat_kategori TEXT DEFAULT ''",
        "ALTER TABLE voucher ADD COLUMN min_qty REAL DEFAULT 0",
        # Voucher PRODUK/KATEGORI: cara potong TOTAL (sekali ke subtotal subset) atau PER_PRODUK (nilai per unit).
        "ALTER TABLE voucher ADD COLUMN potongan_mode TEXT DEFAULT 'TOTAL'",
        # Resep produksi bisa dijual langsung di POS (backflush bahan saat checkout, Model B)
        "ALTER TABLE produksi_template ADD COLUMN pos_aktif INTEGER DEFAULT 0",
        # Harga jual resep di POS. Dulu harga hanya diambil dari produk jadi yang tertaut,
        # padahal resep made-to-order sering TIDAK punya produk jadi -> harga selalu 0.
        # Default 0 = belum diisi, sistem jatuh balik ke harga produk jadi (perilaku lama).
        "ALTER TABLE produksi_template ADD COLUMN harga_jual REAL DEFAULT 0",
        # Kategori resep untuk dikelompokkan/di-filter di grid POS (chip kategori kasir).
        # Default '' = ikut kategori produk jadi tertaut, fallback 'Racikan'.
        "ALTER TABLE produksi_template ADD COLUMN kategori TEXT DEFAULT ''",
        # POS Lanjutan: pilih produk mana yang tampil di kasir. Default 1 = tampil (perilaku lama,
        # semua produk aktif muncul). Di-uncheck lewat POS Master utk menyembunyikan dari grid kasir.
        "ALTER TABLE produk ADD COLUMN pos_tampil INTEGER DEFAULT 1",
        # V6.5: idempotency key checkout POS — cegah transaksi dobel dari klik ganda /
        # koneksi lambat (retry). NULL utk baris lama, aman (unique index di bawah partial).
        "ALTER TABLE pos_sale ADD COLUMN idempotency_key TEXT",
        # V7: PERAN akun di laporan. Dulu Laba Rugi & label transaksi mengunci makna ke NOMOR
        # akun (6130 pasti penyusutan, 6160 pasti bunga, 6170 pasti pajak). Padahal nama akun
        # bawaan BOLEH diganti user, dan banyak yang memakainya ulang (mis. 6170 dijadikan
        # "Beban Pengiriman"). Akibatnya biaya operasional tampil sebagai Pajak/Bunga dan
        # keluar dari Beban Operasional, sehingga Laba Usaha kelihatan lebih besar dari
        # seharusnya. Sekarang maknanya disimpan eksplisit di sini, bukan ditebak dari nomor.
        # NULL = belum pernah diisi (akan di-backfill sekali di bawah); '' = sengaja tanpa peran.
        "ALTER TABLE akun ADD COLUMN peran TEXT",
        # V7: barcode produk (scanner USB/Bluetooth). SENGAJA kolom sendiri, BUKAN
        # menumpang `kode`: `kode` itu SKU internal buatan pemilik ("P001") yang
        # dipakai untuk urutan alami (KODE-2 < KODE-10) dan label ikon POS,
        # sedangkan barcode pabrik itu 13 digit angka. Menumpangkannya merusak
        # keduanya. TEXT, bukan angka: barcode boleh berawalan nol ("0012345")
        # dan nol itu hilang kalau disimpan sebagai bilangan.
        "ALTER TABLE produk ADD COLUMN barcode TEXT DEFAULT ''",
    ]:
        try:
            conn.execute(col_sql)
        except sqlite3.OperationalError as _e:
            _msg = str(_e).lower()
            if 'duplicate column' in _msg or 'already exists' in _msg:
                continue
            app.logger.warning('Migration ALTER gagal: %s | SQL: %s', _e, col_sql)
    # Indeks barcode dibuat SETELAH loop migrasi, karena kolomnya baru ada di situ
    # untuk database lama. Sengaja indeks biasa, BUKAN UNIQUE: barcode kosong itu
    # normal (kebanyakan UMKM tidak punya barcode untuk semua barang) dan keunikan
    # divalidasi di aplikasi supaya pesannya bisa menyebut produk mana yang bentrok.
    conn.execute("CREATE INDEX IF NOT EXISTS idx_produk_barcode ON produk(barcode)")
    conn.execute("UPDATE produk SET barcode='' WHERE barcode IS NULL")
    conn.execute("UPDATE produk SET aktif=1 WHERE aktif IS NULL")
    conn.execute("UPDATE hutang SET akun_kode='2100' WHERE akun_kode IS NULL OR TRIM(akun_kode)=''")
    conn.execute("UPDATE akun SET is_rekening=1 WHERE kode IN ('1100','1110')")
    # Backfill invoice_id ke piutang lama via jurnal_id yang sama
    conn.execute("""
        UPDATE piutang SET invoice_id = (
            SELECT id FROM invoice WHERE invoice.jurnal_id = piutang.jurnal_id LIMIT 1
        )
        WHERE invoice_id IS NULL AND jurnal_id IS NOT NULL
    """)
    # invoice_id yang menunjuk invoice milik jurnal LAIN. Sumbernya Restore "tambah data"
    # versi lama yang menyalin invoice_id mentah dari DB sumber, padahal ID invoice di DB
    # ini berbeda. Akibatnya invoice tunai pelanggan lain ikut terbaca "masih ada pembayaran"
    # dan terkunci Lunas. Kebenarannya jurnal_id: diarahkan ke invoice berjurnal sama, atau
    # dikosongkan kalau tidak ada. Idempoten.
    conn.execute("""
        UPDATE piutang SET invoice_id = (
            SELECT i.id FROM invoice i WHERE i.jurnal_id = piutang.jurnal_id ORDER BY i.id LIMIT 1
        )
        WHERE jurnal_id IS NOT NULL AND invoice_id IS NOT NULL
          AND NOT EXISTS (SELECT 1 FROM invoice i
                          WHERE i.id = piutang.invoice_id AND i.jurnal_id = piutang.jurnal_id)
    """)
    # Voucher KATEGORI lama (single syarat_kategori) -> tabel multi-kategori (idempoten)
    conn.execute("""
        INSERT OR IGNORE INTO voucher_kategori(voucher_id, kategori)
        SELECT id, TRIM(syarat_kategori) FROM voucher
        WHERE scope='KATEGORI' AND TRIM(COALESCE(syarat_kategori,'')) <> ''
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_voucher_produk_vid ON voucher_produk(voucher_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_voucher_customer_vid ON voucher_customer(voucher_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_voucher_kategori_vid ON voucher_kategori(voucher_id)")
    # V6.5: idempotency checkout POS. Partial unique -> baris lama (NULL/'') tak ikut
    # dibatasi, hanya key yang benar-benar dikirim browser yang dijaga unik.
    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_pos_sale_idem ON pos_sale(idempotency_key)
        WHERE idempotency_key IS NOT NULL AND idempotency_key <> ''
    """)

    akun_default = AKUN_DEFAULT
    if conn.execute('SELECT COUNT(*) FROM akun').fetchone()[0] == 0:
        conn.executemany('INSERT INTO akun(kode,nama,tipe,subtipe,saldo_normal) VALUES(?,?,?,?,?)', akun_default)
    else:
        # DB lama: pastikan semua akun default tetap ada (mis. 3300 Prive yang ditambahkan
        # di versi baru). Tanpa ini, pemakaian akun yg belum ter-seed → "Akun tidak ditemukan" → 500.
        conn.executemany('INSERT OR IGNORE INTO akun(kode,nama,tipe,subtipe,saldo_normal) VALUES(?,?,?,?,?)', akun_default)

    # ── Backfill PERAN akun (SEKALI saja: hanya baris yang peran-nya masih NULL) ──
    # Sadar-nama & konservatif: peran khusus hanya diberikan bila akunnya MASIH BERNAMA
    # bawaan. Kalau user sudah mengganti namanya, berarti akun itu dipakai untuk hal lain,
    # jadi diperlakukan sebagai beban operasional biasa. Dengan begitu:
    #  - tenant yang tak pernah mengubah COA: angka laporannya PERSIS sama seperti sebelumnya
    #  - tenant yang memakai ulang akun (mis. 6170 -> "Beban Pengiriman"): otomatis pulih
    # Dipakai WHERE peran IS NULL supaya user yang nanti sengaja mengosongkan peran ('')
    # tidak ditimpa balik setiap aplikasi start.
    conn.execute("""
        UPDATE akun SET peran = CASE
            WHEN kode='6130' AND LOWER(TRIM(nama))='beban penyusutan' THEN 'PENYUSUTAN'
            WHEN kode='6160' AND LOWER(TRIM(nama))='beban bunga'      THEN 'BUNGA'
            WHEN kode='6170' AND LOWER(TRIM(nama))='beban pajak'      THEN 'PAJAK'
            ELSE ''
        END
        WHERE peran IS NULL
    """)

    if conn.execute('SELECT COUNT(*) FROM settings').fetchone()[0] == 0:
        conn.executemany('INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)', [
            ('account_type','single'), ('modal_awal','0'), ('nama_usaha','Usaha Saya'),
        ])
    conn.executemany('INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)', [
        ('pos_struk_subtitle', 'Struk POS'),
        ('pos_struk_footer', 'Terima kasih.'),
        ('pos_struk_label_kasir', 'Kasir'),
        ('pos_struk_show_customer', '1'),
        ('pos_struk_show_payment', '1'),
        ('pos_struk_kertas', POS_STRUK_KERTAS_DEFAULT),
        ('pos_menu_size', 'medium'),
        ('pos_icon_mode', 'initial'),
        ('pos_default_payment', 'TUNAI'),
        ('tz_app', TZ_APP_DEFAULT),
    ])

    if conn.execute('SELECT COUNT(*) FROM users').fetchone()[0] == 0:
        ph = sha256('admin123'.encode()).hexdigest()
        conn.execute("INSERT INTO users(username,password_hash,nama,role) VALUES(?,?,?,?)",
                     ('admin', ph, 'Administrator', 'ADMIN'))
    ph_demo = sha256('demo123'.encode()).hexdigest()
    conn.execute(
        "INSERT OR IGNORE INTO users(username,password_hash,nama,role) VALUES(?,?,?,?)",
        ('demo', ph_demo, 'Demo User', 'DEMO')
    )

    # Migrasi: selaraskan subtipe akun beban dengan pengelompokan calc_profitability
    conn.execute("UPDATE akun SET subtipe='Beban Penyusutan'  WHERE kode='6130' AND subtipe='Beban Operasional'")
    conn.execute("UPDATE akun SET subtipe='Beban Bunga'       WHERE kode='6160' AND subtipe IN ('Beban Lain','Beban Operasional')")
    conn.execute("UPDATE akun SET subtipe='Beban Pajak'       WHERE kode='6170' AND subtipe IN ('Beban Operasional','Beban Lain')")
    conn.execute("UPDATE akun SET subtipe='Beban Operasional' WHERE kode='6180' AND subtipe='Beban Lain'")
    # Migrasi: tambah akun yang mungkin belum ada di database lama
    missing = [
        ('6170', 'Beban Pajak',       'BEBAN',      'Beban Pajak',        'DEBIT'),
        ('6180', 'Beban Lainnya',     'BEBAN',      'Beban Operasional',  'DEBIT'),
        ('4150', 'Retur Penjualan',   'PENDAPATAN', 'Pendapatan Usaha',   'DEBIT'),
        ('5150', 'Retur Pembelian',   'BEBAN',      'HPP',                'KREDIT'),
        ('5190', 'Koreksi Persediaan','BEBAN',      'HPP',                'DEBIT'),
        ('2140', 'Hutang kepada Owner','LIABILITAS','Liabilitas Lancar',  'KREDIT'),
        ('1150', 'Cadangan Kerugian Piutang','ASET','Aset Lancar',        'KREDIT'),
        ('6190', 'Beban Kerugian Piutang','BEBAN',  'Beban Operasional',  'DEBIT'),
        # Piutang Karyawan: kasbon/pinjaman karyawan (non-beban, jalur Pengeluaran → Pinjaman Karyawan)
        ('1160', 'Piutang Karyawan',      'ASET',   'Aset Lancar',        'DEBIT'),
        # Akun pajak (PKP), selalu di-seed agar tersedia kapan saja user butuh.
        # Kode dipilih di range 1180+ dan 2111+ agar tidak konflik dengan akun bank custom user.
        ('1180', 'PPN Masukan (Dapat Dikreditkan)','ASET',      'Aset Lancar',       'DEBIT'),
        ('1181', 'PPh 23 Dibayar Dimuka',         'ASET',       'Aset Lancar',       'DEBIT'),
        ('1182', 'Angsuran PPh Pasal 25',         'ASET',       'Aset Lancar',       'DEBIT'),
        ('1183', 'PPh 22 Dibayar Dimuka',         'ASET',       'Aset Lancar',       'DEBIT'),
        ('2111', 'Hutang PPN',                    'LIABILITAS', 'Liabilitas Lancar', 'KREDIT'),
        ('2112', 'Hutang PPh Pasal 23',           'LIABILITAS', 'Liabilitas Lancar', 'KREDIT'),
        # Akun Gaji & BPJS (Modul PRO Gaji)
        ('2115', 'Hutang PPh Pasal 21',           'LIABILITAS', 'Liabilitas Lancar', 'KREDIT'),
        ('2116', 'Hutang BPJS',                   'LIABILITAS', 'Liabilitas Lancar', 'KREDIT'),
        ('6101', 'Beban BPJS Perusahaan',         'BEBAN',      'Beban Operasional', 'DEBIT'),
    ]
    for kode, nama, tipe, subtipe, saldo_normal in missing:
        conn.execute(
            "INSERT OR IGNORE INTO akun(kode,nama,tipe,subtipe,saldo_normal) VALUES(?,?,?,?,?)",
            (kode, nama, tipe, subtipe, saldo_normal)
        )
    conn.execute("UPDATE akun SET is_rekening=1 WHERE kode IN ('1100','1110')")

    # ── Cleanup akun zombie 2113/2114 (PPN/PPh Disetor) ─────────────────────
    # Akun ini ter-seed di versi lama tapi tidak pernah dijurnal (logic-nya
    # di-skip). Saldo selalu 0 → tampil membingungkan di Neraca & Hub Pajak.
    # Hapus hanya jika benar-benar belum ada mutasi.
    for _zk in ('2113', '2114'):
        _row = conn.execute(
            "SELECT a.id FROM akun a WHERE a.kode=? "
            "AND NOT EXISTS (SELECT 1 FROM detail_jurnal dj WHERE dj.akun_id=a.id)",
            (_zk,)
        ).fetchone()
        if _row:
            conn.execute("DELETE FROM akun WHERE id=?", (_row['id'],))

    # Arsip transaksi yang dihapus (soft-delete + restore). Snapshot JSON penuh
    # supaya bisa dipulihkan kapan saja. Tidak ada FK ke jurnal (jurnalnya sudah dihapus).
    conn.execute("""CREATE TABLE IF NOT EXISTS arsip_transaksi (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jurnal_id_lama INTEGER,
        nomor_tx TEXT,
        tanggal DATE,
        keterangan TEXT,
        tipe_tx TEXT,
        pihak TEXT,
        total REAL DEFAULT 0,
        snapshot TEXT NOT NULL,
        dihapus_oleh TEXT,
        dihapus_pada DATETIME DEFAULT CURRENT_TIMESTAMP,
        restored INTEGER DEFAULT 0,
        restored_pada DATETIME,
        restored_jurnal_id INTEGER
    )""")
    backfill_nomor_tx(conn)
    backfill_payment_journal_links(conn)
    backfill_asset_journal_links(conn)
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_jurnal_nomor_tx ON jurnal(nomor_tx)")
    conn.execute("""CREATE UNIQUE INDEX IF NOT EXISTS idx_invoice_jurnal_id
                    ON invoice(jurnal_id) WHERE jurnal_id IS NOT NULL""")
    # Performance indexes
    conn.execute("CREATE INDEX IF NOT EXISTS idx_jurnal_tanggal ON jurnal(tanggal)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_jurnal_tipe_tgl ON jurnal(tipe_tx, tanggal)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_dj_jid ON detail_jurnal(jurnal_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_dj_akun ON detail_jurnal(akun_id, jurnal_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_ps_tanggal ON pergerakan_stok(tanggal, produk_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_ps_jid ON pergerakan_stok(jurnal_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_pos_sale_tgl ON pos_sale(tanggal)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_piutang_status ON piutang(status)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_hutang_status ON hutang(status)")
    # jurnal_id ditambah via ALTER TABLE (bukan CREATE TABLE) → belum ter-index; dibutuhkan
    # kolom Sisa Tagihan di /daftar-transaksi (WHERE jurnal_id IN (...) tanpa full-scan tabel).
    conn.execute("CREATE INDEX IF NOT EXISTS idx_piutang_jurnal_id ON piutang(jurnal_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_hutang_jurnal_id ON hutang(jurnal_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_prb_jurnal ON penjualan_resep_bahan(jurnal_id)")
    _migrasi_fk_penjualan_resep_bahan(conn)
    # stok_list() menghitung produk 'terpakai' (lihat komentar di route) dengan
    # SELECT DISTINCT produk_id dari tabel-tabel ini. idx_ps_tanggal punya produk_id di posisi
    # KEDUA (bukan leading column) jadi tak dipakai sbg covering index utk DISTINCT produk_id saja;
    # 4 tabel lain sama sekali belum ada index di kolom ini. Tanpa index ini, /stok memicu full
    # table scan di transaksi_item/pergerakan_stok setiap kali dibuka — makin lambat seiring
    # riwayat transaksi bertambah (terverifikasi via stress-test 1,5 juta baris: 0,3 detik/tabel
    # tanpa index -> 0,05 detik dgn index, EXPLAIN QUERY PLAN konfirmasi jadi covering index scan).
    conn.execute("CREATE INDEX IF NOT EXISTS idx_ps_pid ON pergerakan_stok(produk_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_ti_pid ON transaksi_item(produk_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_poi_pid ON po_item(produk_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_produksi_pjid ON produksi(produk_jadi_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_prodbahan_pid ON produksi_bahan(produk_id)")
    # stok_list() juga hitung 'terakhir dibeli' per produk lewat
    # SELECT produk_id, MAX(tanggal) ... WHERE jenis='MASUK' GROUP BY produk_id. idx_ps_tanggal
    # (tanggal, produk_id) & idx_ps_pid (produk_id) tak cocok utk filter+group ini -> SQLite jatuh
    # ke scan seluruh index produk_id (proporsional total baris pergerakan_stok, BUKAN jumlah
    # produk). Index komposit (jenis, produk_id, tanggal) biarkan SQLite langsung loncat ke
    # baris jenis='MASUK' + pakai optimisasi MAX()-lewat-index per grup. Terverifikasi stress-test
    # 1,5 juta baris: 6,8 detik -> 0,04 detik (bottleneck dominan /stok pada DB riwayat panjang).
    conn.execute("CREATE INDEX IF NOT EXISTS idx_ps_jenis_pid_tgl ON pergerakan_stok(jenis, produk_id, tanggal)")

    conn.commit(); conn.close()


def parse_rp(v):
    """Parse Rupiah dari format Indonesia maupun angka desimal ternormalisasi."""
    if v is None or v == '': return 0.0
    text = str(v).strip().replace(' ', '')
    if ',' in text:
        # Tampilan Indonesia: 1.234,56 -> 1234.56
        text = text.replace('.', '').replace(',', '.')
    elif text.count('.') > 1:
        # Tampilan Indonesia tanpa desimal: 1.234.567 -> 1234567
        text = text.replace('.', '')
    elif text.count('.') == 1:
        # Selain format ribuan 1.234, titik diterima sebagai desimal frontend.
        # Pola ribuan yang sah: 1-3 digit di depan & TIDAK diawali "0". Tanpa syarat ini,
        # nilai desimal 3 angka ikut tertelan: "0.125" jadi 125 dan "1234.125" jadi 1234125
        # (1000x). Aturan ini disamakan dengan parseMoney() di base.html.
        kiri, kanan = text.rsplit('.', 1)
        _k = kiri.lstrip('-')
        if (len(kanan) == 3 and _k.isdigit() and kanan.isdigit()
                and 1 <= len(_k) <= 3 and _k != '0'):
            text = kiri + kanan
    return float(text) or 0.0

def parse_qty(v, default=0):
    """Parse qty dari form: koma adalah desimal, titik hasil normalisasi tetap desimal."""
    if v is None or str(v).strip() == '':
        return float(default)
    text = str(v).strip()
    if ',' in text:
        text = text.replace('.', '').replace(',', '.')
    return float(text)

def parse_barcode(v):
    """Rapikan barcode dari form/Excel.

    SENGAJA TIDAK lewat parse_rp()/parse_qty(): barcode itu IDENTITAS berupa teks,
    bukan nominal. Diperlakukan sebagai angka, "0012345" kehilangan nol depannya
    dan barcode 13 digit bisa berubah jadi notasi ilmiah. Yang dibuang cuma spasi
    (scanner kadang menyisipkan spasi/tab) dan karakter kendali dari akhir kiriman.
    Excel juga sering menyerahkan angka sebagai float (8991002103016.0), jadi ekor
    ".0" itu dipangkas.
    """
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    text = str(v).strip()
    if text.endswith('.0') and text[:-2].isdigit():
        text = text[:-2]
    return ''.join(text.split())


def barcode_bentrok(conn, barcode, exclude_id=None):
    """Kembalikan produk lain yang sudah memakai barcode ini, atau None.

    Barcode kosong TIDAK pernah dianggap bentrok: mayoritas produk memang tak
    punya barcode, dan menganggapnya bentrok akan mengunci penyimpanan produk.
    """
    barcode = parse_barcode(barcode)
    if not barcode:
        return None
    sql = 'SELECT id, kode, nama FROM produk WHERE barcode=?'
    params = [barcode]
    if exclude_id:
        sql += ' AND id<>?'
        params.append(exclude_id)
    return conn.execute(sql, params).fetchone()


def parse_percentage(v, label='Persentase', default=0):
    """Parse persen dari form lokal Indonesia dan pastikan angkanya valid."""
    try:
        value = parse_qty(v, default)
    except (ValueError, TypeError):
        raise ValueError(f'{label} harus berupa angka yang valid.')
    if not math.isfinite(value):
        raise ValueError(f'{label} harus berupa angka yang valid.')
    return value

def parse_nonnegative_rp(v, label='Nominal'):
    """Parse angka Rupiah dari form dan tolak nilai invalid atau negatif."""
    try:
        value = parse_rp(v)
    except (ValueError, TypeError):
        raise ValueError(f'{label} harus berupa angka yang valid.')
    if not math.isfinite(value) or value < 0:
        raise ValueError(f'{label} tidak boleh negatif.')
    return value

def _round_rp(x):
    """Bulatkan nominal ke rupiah utuh (half-up) — samakan dengan JS Math.round di frontend.
    Dipakai saat app dalam mode rupiah bulat (pref_decimal OFF) agar tidak ada sisa
    pecahan sub-rupiah yang bocor ke jurnal (mis. hutang hantu Rp 0,5)."""
    try:
        x = float(x)
    except (TypeError, ValueError):
        return 0.0
    if not math.isfinite(x):
        return 0.0
    return math.floor(x + 0.5) if x >= 0 else math.ceil(x - 0.5)

def parse_positive_qty(v, label='Qty'):
    """Parse kuantitas dari form dan pastikan nilainya benar-benar positif."""
    try:
        value = parse_qty(v)
    except (ValueError, TypeError):
        raise ValueError(f'{label} harus berupa angka yang valid.')
    if not math.isfinite(value) or value <= 0:
        raise ValueError(f'{label} harus lebih dari 0.')
    return value

def is_valid_iso_date(v):
    try:
        date.fromisoformat(str(v))
        return True
    except (TypeError, ValueError):
        return False

# ── Paginasi sisi-server (dipakai semua list) ────────────────────────────
PER_PAGE = 50

def _page_arg(default=1):
    """Nomor halaman dari ?page=N, minimal 1."""
    try:
        return max(1, int(request.args.get('page', default) or default))
    except (ValueError, TypeError):
        return 1

def paginate_meta(page, total, per_page=PER_PAGE):
    """Metadata paginasi: clamp halaman, hitung offset & rentang tampilan."""
    per_page = max(1, int(per_page))
    total = max(0, int(total or 0))
    pages = max(1, (total + per_page - 1) // per_page)
    page = min(max(1, int(page or 1)), pages)
    offset = (page - 1) * per_page
    return {
        'page': page, 'per_page': per_page, 'total': total, 'pages': pages,
        'offset': offset, 'has_prev': page > 1, 'has_next': page < pages,
        'prev': page - 1, 'next': page + 1,
        'start': 0 if total == 0 else offset + 1,
        'end': min(offset + per_page, total),
    }

def is_valid_optional_iso_date(v):
    return not v or is_valid_iso_date(v)

def is_rekening_kode(conn, kode):
    return bool(conn.execute(
        "SELECT 1 FROM akun WHERE kode=? AND is_rekening=1",
        ((kode or '').strip(),)
    ).fetchone())

def _reject_form(conn, message, endpoint):
    conn.rollback()
    conn.close()
    flash(message, 'danger')
    return redirect(url_for(endpoint))

# ---------- APP PREFERENCES (global toggles) ----------
# Default preferences, overridden by DB settings table at request time.
_APP_PREFS = {
    'custom_coa_aktif':  False,   # Enable COA override selectors di pemasukan/pengeluaran
    'decimal_aktif':     False,   # Tampilkan 2 angka desimal di filter rp
    'ratio_widgets':     ['current_ratio', 'der', 'roe', 'npm'],
    'simple_mode':       False,   # Sembunyikan menu advanced dari sidebar
    'simple_features':   [],
    'hide_login_hint':   False,   # Sembunyikan kotak "Login default: admin/admin123" di halaman login
}

# Daftar semua rasio yang bisa dipilih untuk widget dashboard.
RATIO_WIDGETS_AVAIL = [
    ('current_ratio',        'Current Ratio',        'Likuiditas',     'fas fa-droplet',       '#0ea5e9'),
    ('quick_ratio',          'Quick Ratio',          'Likuiditas',     'fas fa-bolt',          '#06b6d4'),
    ('cash_ratio',           'Cash Ratio',           'Likuiditas',     'fas fa-coins',         '#14b8a6'),
    ('der',                  'DER',                  'Solvabilitas',   'fas fa-scale-balanced','#f97316'),
    ('debt_to_asset',        'Debt to Asset',        'Solvabilitas',   'fas fa-link',          '#fb923c'),
    ('equity_ratio',         'Equity Ratio',         'Solvabilitas',   'fas fa-shield',        '#fbbf24'),
    ('gpm',                  'Gross Margin',         'Profitabilitas', 'fas fa-percent',       '#22c55e'),
    ('opm',                  'Operating Margin',     'Profitabilitas', 'fas fa-percent',       '#16a34a'),
    ('npm',                  'Net Margin',           'Profitabilitas', 'fas fa-percent',       '#8b5cf6'),
    ('roa',                  'ROA',                  'Profitabilitas', 'fas fa-chart-pie',     '#ec4899'),
    ('roe',                  'ROE',                  'Profitabilitas', 'fas fa-chart-line',    '#10b981'),
    ('inv_turnover',         'Inventory Turnover',   'Aktivitas',      'fas fa-boxes-stacked', '#0284c7'),
    ('asset_turnover',       'Asset Turnover',       'Aktivitas',      'fas fa-rotate',        '#7c3aed'),
    ('receivables_turnover', 'Receivables Turnover', 'Aktivitas',      'fas fa-hand-holding-dollar', '#d97706'),
]

SIMPLE_FEATURES_AVAIL = [
    ('transfer', 'Transfer Rekening', 'Pindah saldo antar kas/bank.'),
    ('retur_penjualan', 'Retur Penjualan', 'Retur dari customer.'),
    ('retur_pembelian', 'Retur Pembelian', 'Retur barang ke vendor.'),
    ('pendanaan', 'Pendanaan', 'Setoran modal dan pinjaman.'),
    ('pelunasan', 'Pelunasan', 'Bayar/terima piutang dan hutang.'),
    ('invoice', 'Invoice', 'Dokumen tagihan penjualan.'),
    ('po', 'Purchase Order', 'Dokumen pesanan pembelian sebelum jadi transaksi.'),
    ('produksi', 'Produksi', 'Ubah bahan baku menjadi produk jadi.'),
    ('hpp_kalkulator', 'Master HPP', 'Hitung resep/menu dan HPP detail.'),
    ('database', 'Database Customer & Vendor', 'Master data customer dan vendor.'),
    ('aset_tetap', 'Aset Tetap', 'Aset, penyusutan, dan investasi aset.'),
    ('grafik', 'Grafik', 'Visualisasi chart laporan.'),
    ('performa_penjualan', 'Performa Penjualan', 'Top produk dan metrik penjualan.'),
    ('perubahan_ekuitas', 'Perubahan Ekuitas', 'Laporan modal/ekuitas.'),
    ('analisis_rasio', 'Analisis Rasio', 'Rasio keuangan advanced.'),
    ('buku_besar', 'Buku Besar', 'Mutasi detail per akun.'),
    ('jurnal', 'Jurnal', 'Halaman jurnal manual/detail.'),
    ('bagan_akun', 'Bagan Akun', 'Kelola COA/akun akuntansi.'),
    ('setup_saldo_awal', 'Setup Saldo Awal', 'Input saldo awal rekening, piutang, hutang, stok.'),
    ('pajak', 'Modul Pajak', 'Hub pajak, PPN, PPh23, bukti potong/bayar.'),
    ('gaji', 'Modul Gaji', 'Data karyawan, proses gaji, slip gaji.'),
    ('anggaran', 'Anggaran & Target', 'Budgeting dan target usaha.'),
    ('proyek', 'Proyek', 'Job costing: laba rugi per proyek.'),
]
SIMPLE_FEATURE_KEYS = {k for k, _, _ in SIMPLE_FEATURES_AVAIL}

USER_PERMISSIONS = [
    ('sales_input', 'Input Penjualan', 'Bisa input transaksi Pemasukan/Penjualan non-POS.'),
    ('pos_sales', 'Akses POS Kasir', 'Bisa masuk POS Kasir dan input order/transaksi POS.'),
    ('expense_input', 'Input Pengeluaran', 'Bisa input transaksi Pengeluaran.'),
    ('inventory_input', 'Input Inventory', 'Bisa membuka inventory, tambah produk, edit data produk, dan koreksi stok.'),
    ('inventory_hpp', 'Lihat/Edit HPP Inventory', 'Bisa melihat dan mengubah harga beli/HPP, valuasi persediaan, dan nilai non-SKU.'),
    ('show_hpp_margin', 'Tampilkan HPP & Gross Profit Margin', 'Bisa melihat HPP, laba kotor, dan gross margin di ringkasan.'),
    ('show_journal', 'Tampilkan Jurnal', 'Bisa membuka daftar transaksi dan detail jurnal.'),
    ('manage_master_pos', 'Kelola Master POS', 'Bisa membuka Master POS (dashboard kasir, grafik penjualan POS, formasi kasir, tambah user kasir, jadwal shift).'),
]
PERMISSION_KEYS = {k for k, _, _ in USER_PERMISSIONS}
ROLE_DEFAULT_PERMISSIONS = {
    'ADMIN': set(PERMISSION_KEYS),
    'FINANCE': set(PERMISSION_KEYS),
    'MANAJER': {'manage_master_pos'},
    'OPERATOR': {'sales_input', 'expense_input', 'inventory_input', 'show_journal'},
    'KASIR': {'pos_sales'},
    'INVESTOR': {'show_hpp_margin', 'inventory_hpp', 'show_journal'},
    'DEMO': set(PERMISSION_KEYS),
}

def _decode_permissions(raw):
    if raw is None:
        return None
    try:
        data = json.loads(raw or '[]')
    except Exception:
        data = []
    return {p for p in data if p in PERMISSION_KEYS}

def effective_permissions(role, raw_permissions=None):
    if role == 'ADMIN':
        return set(PERMISSION_KEYS)
    custom = _decode_permissions(raw_permissions)
    if custom is not None:
        if role != 'KASIR' and 'pos_sales' in custom:
            custom.add('sales_input')
        return custom
    return set(ROLE_DEFAULT_PERMISSIONS.get(role or '', set()))

def current_permissions(conn=None):
    role = session.get('role', '')
    if role == 'ADMIN':
        return set(PERMISSION_KEYS)
    uid = session.get('user_id')
    raw = None
    if uid:
        close_after = False
        if conn is None:
            conn = db()
            close_after = True
        row = conn.execute("SELECT permissions FROM users WHERE id=?", (uid,)).fetchone()
        raw = row['permissions'] if row else None
        if close_after:
            conn.close()
    return effective_permissions(role, raw)

def has_permission(key, conn=None):
    return key in current_permissions(conn)

def _permission_denied(message='Akses ditolak. Anda tidak memiliki hak akses untuk fitur ini.'):
    flash(message, 'danger')
    return redirect(url_for('dashboard'))

def _permissions_from_form():
    return sorted({p for p in request.form.getlist('permissions') if p in PERMISSION_KEYS})

def reload_app_prefs():
    """Refresh global _APP_PREFS dari settings table."""
    try:
        conn = db()
        _APP_PREFS['custom_coa_aktif'] = get_setting(conn, 'pref_custom_coa', '0') == '1'
        _APP_PREFS['decimal_aktif']    = get_setting(conn, 'pref_decimal',    '0') == '1'
        _APP_PREFS['hide_login_hint']  = get_setting(conn, 'pref_hide_login_hint', '0') == '1'
        raw_w = get_setting(conn, 'pref_ratio_widgets', '')
        if raw_w:
            picks = [w.strip() for w in raw_w.split(',') if w.strip()]
            valid_keys = {k for (k,*_rest) in RATIO_WIDGETS_AVAIL}
            picks = [w for w in picks if w in valid_keys]
            _APP_PREFS['ratio_widgets'] = picks if picks else ['current_ratio','der','roe','npm']
        else:
            _APP_PREFS['ratio_widgets'] = ['current_ratio','der','roe','npm']
        _APP_PREFS['simple_mode'] = get_setting(conn, 'pref_simple_mode', '0') == '1'
        raw_simple = get_setting(conn, 'pref_simple_features', '')
        simple_picks = [w.strip() for w in raw_simple.split(',') if w.strip()]
        _APP_PREFS['simple_features'] = [w for w in simple_picks if w in SIMPLE_FEATURE_KEYS]
        conn.close()
    except Exception:
        pass  # DB belum siap → pakai default

# ---------- FILTERS ----------
@app.template_filter('rp')
def rp_filter(v):
    if v is None: v = 0
    v = float(v)
    prefix = '-Rp' if v < 0 else 'Rp'
    if _APP_PREFS.get('decimal_aktif'):
        # Format desimal: "1.500.000,23"
        s = "{:,.2f}".format(abs(v))                  # "1,500,000.23"
        s = s.replace(',', '§').replace('.', ',').replace('§', '.')
        return "{} {}".format(prefix, s)
    return "{} {:,.0f}".format(prefix, abs(v)).replace(',', '.')

@app.template_filter('qty')
def qty_filter(v):
    """Format kuantitas desimal: buang nol di belakang, desimal pakai koma (gaya
    Indonesia), TANPA pemisah ribuan (cegah ambigu '1.500' = 1,5 vs 1500).
    Contoh: 2.0 -> '2', 2.5 -> '2,5', 0.25 -> '0,25', 1500 -> '1500'.
    Dipakai HANYA untuk tampilan; jangan dipakai di value input number / argumen JS
    (yang butuh titik desimal mentah)."""
    if v is None:
        return '0'
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    if not math.isfinite(f):
        return '0'
    # Bulatkan ke 6 desimal untuk membuang artefak float (mis. 0.1+0.2), lalu
    # rapikan nol di belakang. Input app maksimal 3 desimal, jadi tak ada presisi hilang.
    s = "{:.6f}".format(f).rstrip('0').rstrip('.')
    if s in ('', '-0'):
        s = '0'
    return s.replace('.', ',')

@app.template_filter('inputnum')
def inputnum_filter(v):
    """Nilai numerik untuk atribut value= pada input nominal (type=number) & literal angka JS.

    Kolom uang di DB bertipe REAL dan bisa membawa ekor presisi penuh hasil pembagian,
    mis. 465283.3333333333 (= 1.395.850 / 3) atau 17800.545454545454. Kalau ekor itu
    dirender apa adanya ke value=, dua hal buruk terjadi:
      1) user melihat "angka aneh" 17.800,54545454 saat mengedit, dan
      2) kalau field belum sempat diformat, titik desimalnya bisa ikut dibuang sebagai
         pemisah ribuan sehingga nilai membengkak 10^n (total pernah jadi Rp 9,3 kuadriliun).
    Jadi bulatkan ke 2 desimal & buang nol di belakang. Tetap memakai TITIK sebagai
    desimal karena input type=number menolak koma."""
    if v is None:
        return '0'
    try:
        f = float(v)
    except (TypeError, ValueError):
        return '0'
    if not math.isfinite(f):
        return '0'
    # Half-up (bukan pembulatan bankir bawaan round()) supaya sama dengan _round_rp &
    # Math.round di frontend: 0,125 -> 0,13. Decimal(str(f)) menghindari kejutan biner.
    from decimal import Decimal, ROUND_HALF_UP
    try:
        d = Decimal(str(f)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    except Exception:
        return '0'
    s = format(d, 'f').rstrip('0').rstrip('.')
    return s if s not in ('', '-', '-0') else '0'

@app.template_filter('tgl')
def tgl_filter(v):
    if not v: return '-'
    try:
        d = datetime.strptime(str(v)[:10], '%Y-%m-%d')
        return f"{d.day} {BULAN[d.month-1]} {d.year}"
    except: return str(v)

@app.template_filter('jam')
def jam_filter(v, mode='penuh'):
    """Ubah cap waktu otomatis dari UTC ke waktu setempat sebelum ditampilkan.

    Kolom ber-DEFAULT CURRENT_TIMESTAMP (log_aktivitas.waktu, jurnal.dibuat,
    pos_shift.dibuka/ditutup, arsip_transaksi.dihapus_pada, dsb) SELALU disimpan
    SQLite dalam UTC. Itu perilaku baku SQLite, bukan kerusakan data. Yang dulu
    keliru adalah menampilkannya mentah, jadi pemakai WIB melihat jam mundur 7 jam.

    astimezone() tanpa argumen memakai zona waktu sistem, jadi satu kode ini benar
    di dua tempat sekaligus: di Windows/Mac ikut jam perangkat, di server ikut TZ
    yang dipasang apply_tz_app() dari setelan Zona Waktu.

    SENGAJA satu filter untuk semua layar. Konversi yang ditulis ulang per halaman
    cepat atau lambat akan saling menyimpang.
    """
    if not v:
        return ''
    s = str(v).strip().replace('T', ' ')[:19]
    d = None
    for pola in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M'):
        try:
            d = datetime.strptime(s, pola)
            break
        except ValueError:
            pass
    if d is None:
        return str(v)   # bentuk tak dikenal: tampilkan apa adanya, jangan bikin halaman error
    lokal = d.replace(tzinfo=timezone.utc).astimezone()
    return lokal.strftime('%H:%M' if mode == 'jam' else '%Y-%m-%d %H:%M')

@app.context_processor
def inject_page_url():
    """page_url(n): URL halaman-n untuk endpoint sekarang, mempertahankan filter/search lain."""
    def page_url(page):
        args = request.args.to_dict(flat=True)
        args['page'] = page
        qs = urlencode(args)
        return request.path + (('?' + qs) if qs else '')
    return dict(page_url=page_url)

@app.context_processor
def inject_globals():
    conn = db()
    # Batch semua settings dalam satu query
    _KEYS = ('nama_usaha','pref_custom_coa','pref_decimal','pref_hide_login_hint',
             'pref_ratio_widgets','pref_simple_mode','pref_simple_features','pos_default_payment')
    placeholders = ','.join('?' for _ in _KEYS)
    rows = conn.execute(f"SELECT key,value FROM settings WHERE key IN ({placeholders})", _KEYS).fetchall()
    _s = {r['key']: r['value'] for r in rows}
    def _g(k, default=''):
        return _s.get(k, default) or default

    nama_usaha = _g('nama_usaha', 'FinansialApp')
    # Refresh global prefs supaya rp filter & toggles sinkron
    _APP_PREFS['custom_coa_aktif'] = _g('pref_custom_coa', '0') == '1'
    _APP_PREFS['decimal_aktif']    = _g('pref_decimal',    '0') == '1'
    _APP_PREFS['hide_login_hint']  = _g('pref_hide_login_hint', '0') == '1'
    raw_w = _g('pref_ratio_widgets', '')
    if raw_w:
        picks = [w.strip() for w in raw_w.split(',') if w.strip()]
        valid_keys = {k for (k,*_rest) in RATIO_WIDGETS_AVAIL}
        picks = [w for w in picks if w in valid_keys]
        _APP_PREFS['ratio_widgets'] = picks if picks else ['current_ratio','der','roe','npm']
    else:
        _APP_PREFS['ratio_widgets'] = ['current_ratio','der','roe','npm']
    _APP_PREFS['simple_mode'] = _g('pref_simple_mode', '0') == '1'
    raw_simple = _g('pref_simple_features', '')
    simple_picks = [w.strip() for w in raw_simple.split(',') if w.strip()]
    _APP_PREFS['simple_features'] = [w for w in simple_picks if w in SIMPLE_FEATURE_KEYS]
    pos_default_payment = _g('pos_default_payment', 'TUNAI')
    if pos_default_payment not in ('TUNAI', 'REKENING', 'PIUTANG'):
        pos_default_payment = 'TUNAI'
    show_feature = {
        k: ((not _APP_PREFS['simple_mode']) or (k in _APP_PREFS['simple_features']))
        for k in SIMPLE_FEATURE_KEYS
    }
    perms = current_permissions(conn)
    conn.close()
    return {'now': datetime.now(), 'today': date.today(),
            'session': session, 'nama_usaha': nama_usaha,
            'app_pref': dict(_APP_PREFS),
            'ratio_widgets_avail': RATIO_WIDGETS_AVAIL,
            'simple_features_avail': SIMPLE_FEATURES_AVAIL,
            'show_feature': show_feature,
            'pos_default_payment': pos_default_payment,
            'demo_lock': DEMO_LOCK,
            'reset_lokal_aktif': RESET_LOKAL_AKTIF and not DEMO_LOCK,
            'can': {k: (k in perms) for k in PERMISSION_KEYS}}


# ---------- AUTH ----------
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('login'))
        if session.get('role') != 'ADMIN':
            flash('Akses ditolak. Hanya Admin.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated

def investor_required(f):
    """ADMIN + FINANCE + INVESTOR, read-only financial views."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('login'))
        if session.get('role') not in ('ADMIN', 'FINANCE', 'INVESTOR', 'DEMO'):
            flash('Akses ditolak. Halaman ini hanya untuk Admin, Finance, dan Investor.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated

def invoice_read_required(f):
    """Invoice juga dapat dibaca operator yang membuat transaksi penjualan."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('login'))
        if session.get('role') not in ('ADMIN', 'FINANCE', 'INVESTOR', 'OPERATOR', 'DEMO'):
            flash('Akses ditolak. Anda tidak memiliki akses ke invoice.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated

def transaction_read_required(f):
    """Daftar transaksi operasional dapat dibaca operator, tetapi bukan role legacy VIEWER."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('login'))
        if session.get('role') not in ('ADMIN', 'FINANCE', 'INVESTOR', 'DEMO') and not has_permission('show_journal'):
            flash('Akses ditolak. Anda tidak memiliki akses ke transaksi.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated

def finance_required(f):
    """ADMIN + FINANCE, full financial write access."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('login'))
        if session.get('role') not in ('ADMIN', 'FINANCE', 'DEMO'):
            flash('Akses ditolak. Halaman ini hanya untuk Admin dan Finance.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated

def operator_required(f):
    """ADMIN + FINANCE + OPERATOR, input pages."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('login'))
        if session.get('role') not in ('ADMIN', 'FINANCE', 'OPERATOR', 'DEMO'):
            flash('Akses ditolak. Halaman ini hanya untuk Admin, Finance, dan Operator.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated

def sales_input_required(f):
    """Input penjualan non-POS permission."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('login'))
        if not has_permission('sales_input'):
            return _permission_denied('Akses ditolak. Anda tidak memiliki hak input penjualan/pemasukan.')
        return f(*args, **kwargs)
    return decorated

def expense_input_required(f):
    """Input pengeluaran permission."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('login'))
        if not has_permission('expense_input'):
            return _permission_denied('Akses ditolak. Anda tidak memiliki hak input pengeluaran.')
        return f(*args, **kwargs)
    return decorated

def hpp_margin_required(f):
    """HPP, laba kotor, dan gross margin hanya untuk user yang diberi izin."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            if request.is_json or request.path.startswith('/api/'):
                return jsonify({'error': 'login required'}), 401
            return redirect(url_for('login'))
        if not has_permission('show_hpp_margin'):
            if request.is_json or request.path.startswith('/api/'):
                return jsonify({'error': 'forbidden'}), 403
            return _permission_denied('Akses ditolak. Anda tidak memiliki hak melihat HPP dan gross profit margin.')
        return f(*args, **kwargs)
    return decorated

def inventory_required(f):
    """Inventory section permission."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('login'))
        if not has_permission('inventory_input'):
            return _permission_denied('Akses ditolak. Anda tidak memiliki hak input inventory.')
        return f(*args, **kwargs)
    return decorated

def inventory_hpp_required(f):
    """Harga beli/HPP inventory dan valuasi persediaan hanya untuk user berizin."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            if request.is_json or request.path.startswith('/api/'):
                return jsonify({'error': 'login required'}), 401
            return redirect(url_for('login'))
        if not has_permission('inventory_hpp'):
            if request.is_json or request.path.startswith('/api/'):
                return jsonify({'error': 'forbidden'}), 403
            return _permission_denied('Akses ditolak. Anda tidak memiliki hak melihat atau mengubah HPP inventory.')
        return f(*args, **kwargs)
    return decorated

def pos_required(f):
    """POS portal: module must be active and user must be a cashier-capable role."""
    @wraps(f)
    def decorated(*args, **kwargs):
        conn = db()
        aktif = is_module_active(conn, 'POS_KASIR')
        conn.close()
        if not aktif:
            return render_template('pos_locked.html'), 403
        if not session.get('user_id'):
            return redirect(url_for('pos_login'))
        if not has_permission('pos_sales'):
            flash('Akses POS hanya untuk user yang diberi hak Akses POS Kasir.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated

def master_pos_required(f):
    """Master POS dashboard: module aktif + user punya hak 'manage_master_pos'.
    Berbeda dgn pos_required: ini utk MANAGER/OWNER yang memantau performa POS,
    bukan utk kasir yang input transaksi. Default dimiliki ADMIN, FINANCE & MANAJER."""
    @wraps(f)
    def decorated(*args, **kwargs):
        conn = db()
        aktif = is_module_active(conn, 'POS_KASIR')
        conn.close()
        if not aktif:
            flash('Modul POS Kasir belum aktif di sistem.', 'warning')
            return redirect(url_for('dashboard'))
        if not session.get('user_id'):
            return redirect(url_for('login'))
        if not has_permission('manage_master_pos'):
            return _permission_denied('Akses Master POS hanya untuk user yang diberi hak "Kelola Master POS".')
        return f(*args, **kwargs)
    return decorated

@app.before_request
def demo_readonly():
    if session.get('role') != 'DEMO':
        return
    if request.method in ('POST', 'PUT', 'DELETE', 'PATCH'):
        if request.is_json or request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            from flask import jsonify
            return jsonify({'error': 'User Demo tidak bisa menggunakan fitur ini'}), 403
        flash('User Demo tidak bisa menggunakan fitur ini', 'warning')
        return redirect(request.referrer or url_for('dashboard'))


def _operator_date_ok(tanggal_str):
    """Return True if date is in current month, or user is not OPERATOR."""
    try:
        tgl = datetime.strptime(tanggal_str, '%Y-%m-%d').date()
        if session.get('role') != 'OPERATOR':
            return True
        today = date.today()
        return tgl.year == today.year and tgl.month == today.month
    except (TypeError, ValueError):
        return False

@app.route('/login', methods=['GET','POST'])
def login():
    if session.get('user_id'):
        perms = current_permissions()
        if session.get('role') == 'KASIR' and perms <= {'pos_sales'}:
            return redirect(url_for('pos_home'))
        if session.get('role') == 'MANAJER' and perms <= {'manage_master_pos'} and _pos_module_active():
            return redirect(url_for('pos_master'))
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form.get('username','').strip()
        password = request.form.get('password','')
        ph = sha256(password.encode()).hexdigest()
        conn = db()
        user = conn.execute(
            "SELECT * FROM users WHERE username=? AND password_hash=? AND aktif=1",
            (username, ph)
        ).fetchone()
        conn.close()
        if user:
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['nama'] = user['nama'] or user['username']
            session['role'] = user['role']
            conn2 = db()
            add_log(conn2, 'Login', f"Role: {user['role']}", 'ADMIN')
            conn2.commit(); conn2.close()
            user_perms = effective_permissions(user['role'], user['permissions'] if 'permissions' in user.keys() else None)
            if user['role'] == 'KASIR' and user_perms <= {'pos_sales'}:
                return redirect(url_for('pos_home'))
            if user['role'] == 'MANAJER' and user_perms <= {'manage_master_pos'} and _pos_module_active():
                return redirect(url_for('pos_master'))
            return redirect(url_for('dashboard'))
        flash('Username atau password salah.', 'danger')
    return render_template('login.html')


def _password_bawaan(username):
    """Password bawaan untuk satu user saat pemulihan.

    admin & demo kembali ke bawaan pabrik yang sudah dikenal; user buatan sendiri
    (mis. 'della', 'kasir1') tidak punya bawaan, jadi dipakai pola username+123.
    Semuanya ditampilkan ke layar sesudah reset, jadi tidak ada yang perlu ditebak.
    """
    u = (username or '').strip()
    if u.lower() == 'admin':
        return 'admin123'
    if u.lower() == 'demo':
        return 'demo123'
    return f'{u}123'


@app.route('/lupa-password', methods=['GET', 'POST'])
def lupa_password():
    """Kembalikan SEMUA login ke bawaan. Transaksi TIDAK disentuh sama sekali.

    Sengaja tanpa syarat apa pun: yang bisa membuka halaman ini sudah duduk di
    depan komputer tempat finansial.db berada, jadi dia sebenarnya sudah bisa
    melakukan hal yang sama lewat DB Browser. Yang dilakukan di sini cuma
    membuat yang memang sudah mungkin jadi tidak menyakitkan.

    SEMUA user direset, bukan cuma admin. Ini disengaja sebagai pengaman: akses
    admin tidak bisa diambil diam-diam, karena semua orang akan sadar
    passwordnya ikut berubah. Setiap reset juga dicatat di Log Aktivitas.
    """
    if DEMO_LOCK:
        flash('Mode demo aktif: password tidak dapat diubah.', 'warning')
        return redirect(url_for('login'))

    conn = db()
    if request.method == 'POST':
        # Hanya aplikasi desktop yang boleh benar-benar mereset. Di tenant online
        # halaman ini cuma panduan; tanpa gerbang ini siapa pun yang membuka
        # halaman login tenant bisa mengganti password semua orang.
        if not RESET_LOKAL_AKTIF:
            conn.close()
            abort(404)
        users = conn.execute(
            "SELECT id, username, nama, role FROM users ORDER BY id").fetchall()
        hasil = []
        for u in users:
            pw = _password_bawaan(u['username'])
            conn.execute("UPDATE users SET password_hash=?, aktif=1 WHERE id=?",
                         (sha256(pw.encode()).hexdigest(), u['id']))
            hasil.append({'username': u['username'], 'nama': u['nama'] or u['username'],
                          'role': u['role'], 'password': pw})
        # Jejak permanen: pemilik tetap melihat bahwa reset pernah terjadi & kapan.
        add_log(conn, 'Pemulihan password',
                f'Semua login ({len(hasil)} user) dikembalikan ke bawaan lewat halaman Lupa Password',
                'PASSWORD')
        conn.commit()
        conn.close()
        session.clear()     # sesi lama tidak boleh ikut terbawa
        return render_template('lupa_password.html', hasil=hasil, selesai=True)

    # Daftar pengguna dan jumlah transaksi HANYA untuk aplikasi desktop. Di pemasangan
    # online halaman ini terbuka untuk pengunjung tanpa login, jadi keduanya tidak boleh
    # ikut terkirim (dulu username, nama, dan peran semua pengguna tampil ke siapa saja).
    users, n_jurnal = [], 0
    if RESET_LOKAL_AKTIF:
        users = conn.execute(
            "SELECT username, nama, role, COALESCE(aktif,1) AS aktif FROM users ORDER BY id"
        ).fetchall()
        n_jurnal = conn.execute("SELECT COUNT(*) c FROM jurnal").fetchone()['c']
    conn.close()
    return render_template('lupa_password.html', users=users, n_jurnal=n_jurnal,
                           selesai=False)


@app.route('/ganti-password', methods=['POST'])
@admin_required
def ganti_password():
    if DEMO_LOCK:
        flash('Mode demo aktif: password tidak dapat diubah.', 'warning')
        return redirect(request.referrer or url_for('dashboard'))
    uid = session['user_id']
    pw_lama    = request.form.get('pw_lama', '')
    pw_baru    = request.form.get('pw_baru', '')
    pw_konfirm = request.form.get('pw_konfirm', '')

    if len(pw_baru) < 6:
        flash('Password baru minimal 6 karakter.', 'danger')
        return redirect(request.referrer or url_for('dashboard'))
    if pw_baru != pw_konfirm:
        flash('Konfirmasi password baru tidak cocok.', 'danger')
        return redirect(request.referrer or url_for('dashboard'))

    conn = db()
    user = conn.execute("SELECT password_hash FROM users WHERE id=?", (uid,)).fetchone()
    if not user or sha256(pw_lama.encode()).hexdigest() != user['password_hash']:
        conn.close()
        flash('Password lama tidak benar.', 'danger')
        return redirect(request.referrer or url_for('dashboard'))

    conn.execute("UPDATE users SET password_hash=? WHERE id=?",
                 (sha256(pw_baru.encode()).hexdigest(), uid))
    add_log(conn, 'Ubah password sendiri', f"User: {session.get('username')}", 'PASSWORD')
    conn.commit()
    conn.close()
    flash('Password berhasil diubah!', 'success')
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ---------- HELPERS ----------
def _qv(conn, where_sql, params, normal='DEBIT'):
    """Total nilai akun ber-WHERE, dijumlah SATU ARAH sesuai sifat alami TIPE-nya,
    BUKAN per-saldo_normal tiap akun. Tujuannya supaya akun KONTRA ikut net negatif.
      normal='DEBIT'  -> SUM(debit-kredit)  (total BEBAN/ASET-debit/Prive)
      normal='KREDIT' -> SUM(kredit-debit)  (total PENDAPATAN/LIABILITAS)
    Dengan begini Retur Penjualan (PENDAPATAN, saldo_normal DEBIT) otomatis MENGURANGI
    pendapatan, dan Retur Pembelian (BEBAN, saldo_normal KREDIT) otomatis MENGURANGI HPP.
    (Versi lama netting per-saldo_normal -> akun kontra salah-tanda -> laba & neraca rusak saat retur.)"""
    expr = "d.debit-d.kredit" if normal == 'DEBIT' else "d.kredit-d.debit"
    r = conn.execute(f"""
        SELECT COALESCE(SUM({expr}), 0) as v
        FROM akun a
        JOIN detail_jurnal d ON d.akun_id=a.id
        JOIN jurnal j ON j.id=d.jurnal_id
        WHERE {where_sql}
    """, params).fetchone()
    return float(r['v']) if r else 0.0

def persediaan_ledger_value(conn):
    """Saldo ledger semua akun persediaan, termasuk custom COA seperti 1131."""
    return _qv(conn, "a.tipe='ASET' AND a.saldo_normal='DEBIT' AND (a.kode LIKE '113%' OR a.subtipe LIKE '%ersediaan%')", [])

def calc_profitability(conn, sd, ed):
    def q(where, p, normal='DEBIT'): return _qv(conn, where, p, normal)
    rev      = q("a.tipe='PENDAPATAN' AND j.tanggal>=? AND j.tanggal<=?", [sd,ed], 'KREDIT')
    # HPP termasuk koreksi persediaan (5190) sesuai PSAK 14 ¶38:
    # kerugian persediaan & koreksi turun-nilai diakui sebagai beban (bagian HPP).
    hpp      = q("a.kode IN ('5100','5150','5190') AND j.tanggal>=? AND j.tanggal<=?", [sd,ed])
    # Biaya operasional: semua beban kecuali HPP/retur-beli, penyusutan, bunga, dan pajak.
    # Penyusutan/bunga/pajak dikenali dari kolom akun.peran (lihat backfill di init_db),
    # BUKAN dari nomor akun. Nama akun bawaan boleh diganti user, jadi 6170 belum tentu
    # pajak; kalau maknanya ditebak dari nomor, biaya operasional bisa salah keluar dari
    # Laba Usaha dan membuat laba tampak lebih besar dari kenyataan.
    op_exp   = q("a.tipe='BEBAN' AND a.kode NOT IN ('5100','5150','5190') "
                 "AND COALESCE(a.peran,'') NOT IN ('PENYUSUTAN','BUNGA','PAJAK') "
                 "AND j.tanggal>=? AND j.tanggal<=?", [sd,ed])
    depr     = q("COALESCE(a.peran,'')='PENYUSUTAN' AND j.tanggal>=? AND j.tanggal<=?", [sd,ed])
    interest = q("COALESCE(a.peran,'')='BUNGA' AND j.tanggal>=? AND j.tanggal<=?", [sd,ed])
    tax      = q("COALESCE(a.peran,'')='PAJAK' AND j.tanggal>=? AND j.tanggal<=?", [sd,ed])
    prive    = q("a.kode='3300' AND j.tanggal>=? AND j.tanggal<=?", [sd,ed])

    laba_kotor  = rev - hpp
    laba_op     = laba_kotor - op_exp        # Margin Operasional (sebelum penyusutan)
    ebit        = laba_op - depr             # EBIT (setelah penyusutan, sebelum bunga & pajak)
    laba_bersih = ebit - interest - tax
    laba_tahan  = laba_bersih - prive

    def pct(v): return round(v/rev*100,1) if rev else 0
    return dict(
        rev=rev, hpp=hpp, op_exp=op_exp, depr=depr, interest=interest,
        tax=tax, prive=prive,
        laba_kotor=laba_kotor, laba_op=laba_op, ebit=ebit,
        laba_bersih=laba_bersih, laba_tahan=laba_tahan,
        other_e=interest,
        pct_hpp=pct(hpp), pct_op=pct(op_exp), pct_depr=pct(depr),
        pct_tax=pct(tax), pct_prive=pct(prive),
        pct_laba_kotor=pct(laba_kotor), pct_laba_op=pct(laba_op),
        pct_ebit=pct(ebit), pct_laba_bersih=pct(laba_bersih),
        pct_laba_tahan=pct(laba_tahan),
    )

def _bs_snapshot(conn, ed):
    """Snapshot Neraca pada tanggal ed (kumulatif s/d ed).
    Mengikuti pola route `neraca`: contra-asset (saldo_normal=KREDIT pada tipe ASET)
    mengurangi total aset; contra-equity (saldo_normal=DEBIT pada tipe EKUITAS,
    mis. Prive 3300) mengurangi total ekuitas.
    """
    rows = conn.execute("""
        SELECT a.kode, a.tipe, a.subtipe, a.saldo_normal,
               COALESCE(SUM(CASE WHEN j.tanggal <= ? THEN d.debit ELSE 0 END),0) as td,
               COALESCE(SUM(CASE WHEN j.tanggal <= ? THEN d.kredit ELSE 0 END),0) as tk
        FROM akun a
        LEFT JOIN detail_jurnal d ON d.akun_id = a.id
        LEFT JOIN jurnal j ON j.id = d.jurnal_id
        GROUP BY a.id
    """, (ed, ed)).fetchall()
    out = dict(
        aset_lancar=0.0, aset_tetap=0.0, total_aset=0.0,
        liab_lancar=0.0, liab_jp=0.0, total_liab=0.0,
        ekuitas_inti=0.0, persediaan=0.0, piutang=0.0, kas_bank=0.0,
    )
    for r in rows:
        saldo = (r['td'] - r['tk']) if r['saldo_normal']=='DEBIT' else (r['tk'] - r['td'])
        kode = r['kode'] or ''
        if r['tipe'] == 'ASET':
            net = -saldo if r['saldo_normal'] == 'KREDIT' else saldo
            out['total_aset'] += net
            sub = (r['subtipe'] or 'Aset Lancar').lower()
            if 'lancar' in sub and 'tidak' not in sub:
                out['aset_lancar'] += net
            else:
                out['aset_tetap'] += net
            if kode.startswith('113'):
                out['persediaan'] += net
            if kode.startswith('112'):
                out['piutang'] += net
        elif r['tipe'] == 'LIABILITAS':
            out['total_liab'] += saldo
            sub = (r['subtipe'] or 'Liabilitas Lancar').lower()
            if 'lancar' in sub and 'panjang' not in sub and 'tidak' not in sub:
                out['liab_lancar'] += saldo
            else:
                out['liab_jp'] += saldo
        elif r['tipe'] == 'EKUITAS':
            net = -saldo if r['saldo_normal'] == 'DEBIT' else saldo
            out['ekuitas_inti'] += net

    fiscal_start = f"{ed[:4]}-01-01"
    laba_tahun   = calc_profitability(conn, fiscal_start, ed)['laba_bersih']
    laba_all     = calc_profitability(conn, '2000-01-01', ed)['laba_bersih']
    laba_ditahan = laba_all - laba_tahun
    out['laba_tahun_berjalan'] = laba_tahun
    out['laba_ditahan'] = laba_ditahan
    out['total_ekuitas'] = out['ekuitas_inti'] + laba_tahun + laba_ditahan

    kas_row = conn.execute("""
        SELECT COALESCE(SUM(CASE WHEN j2.tanggal <= ? THEN dj.debit - dj.kredit ELSE 0 END), 0) as v
        FROM akun a
        LEFT JOIN detail_jurnal dj ON dj.akun_id = a.id
        LEFT JOIN jurnal j2 ON j2.id = dj.jurnal_id
        WHERE a.is_rekening = 1
    """, (ed,)).fetchone()
    out['kas_bank'] = float(kas_row['v'] or 0)
    return out

def _balance_warning(conn, ed=None):
    """Bila Neraca tidak seimbang (A != L + E) pada tanggal ed, kembalikan pesan
    peringatan KERAS + saran perbaikan; selain itu None. Toleransi Rp1 untuk pembulatan.
    Dipakai setelah edit/hapus transaksi (V5.3) agar ketidakseimbangan langsung terlihat."""
    ed = ed or date.today().strftime('%Y-%m-%d')
    s = _bs_snapshot(conn, ed)
    selisih = s['total_aset'] - (s['total_liab'] + s['total_ekuitas'])
    if abs(selisih) <= 1:
        return None
    nominal = "Rp " + format(int(round(abs(selisih))), ',d').replace(',', '.')
    arah = 'lebih besar' if selisih > 0 else 'lebih kecil'
    return ('NERACA TIDAK SEIMBANG: total Aset ' + arah + ' dari Liabilitas + Ekuitas sebesar '
            + nominal + '. Saran perbaikan: buka menu Neraca lalu lihat panel Diagnostik untuk '
            'jurnal yang tidak seimbang; periksa transaksi ber-HPP yang baru diedit/dihapus; '
            'perbaiki lewat Edit Transaksi atau buat jurnal koreksi. Bila ragu, urungkan perubahan terakhir.')

def calc_financial_ratios(conn, sd, ed):
    """Hitung rasio keuangan: Likuiditas, Solvabilitas, Profitabilitas, Aktivitas.
    - sd..ed: periode untuk rasio flow (P&L based: revenue, HPP, laba bersih).
    - ed: tanggal snapshot untuk pos neraca.
    Mengembalikan dict berisi nilai numerik + interpretasi singkat per rasio.
    """
    bs  = _bs_snapshot(conn, ed)
    pnl = calc_profitability(conn, sd, ed)

    def sdiv(a, b):
        return (a / b) if b else 0.0

    rev         = pnl['rev']
    hpp         = pnl['hpp']
    laba_kotor  = pnl['laba_kotor']
    laba_op     = pnl['laba_op']
    laba_bersih = pnl['laba_bersih']

    al = bs['aset_lancar']; ll = bs['liab_lancar']
    ta = bs['total_aset'];  tl = bs['total_liab']; te = bs['total_ekuitas']

    current_ratio = sdiv(al, ll)
    quick_ratio   = sdiv(al - bs['persediaan'], ll)
    cash_ratio    = sdiv(bs['kas_bank'], ll)

    der           = sdiv(tl, te)
    debt_to_asset = sdiv(tl, ta)
    equity_ratio  = sdiv(te, ta)

    gpm = sdiv(laba_kotor, rev)  * 100
    opm = sdiv(laba_op,    rev)  * 100
    npm = sdiv(laba_bersih, rev) * 100
    roa = sdiv(laba_bersih, ta)  * 100
    roe = sdiv(laba_bersih, te)  * 100

    inv_turnover         = sdiv(hpp, bs['persediaan'])
    asset_turnover       = sdiv(rev, ta)
    receivables_turnover = sdiv(rev, bs['piutang'])

    def hi(value, lo, hi_):
        """Higher is better."""
        if value >= hi_: return 'Sehat'
        if value >= lo:  return 'Cukup'
        return 'Perlu Perhatian'

    def lo(value, good_max, ok_max):
        """Lower is better."""
        if value <= good_max: return 'Sehat'
        if value <= ok_max:   return 'Cukup'
        return 'Perlu Perhatian'

    return dict(
        bs=bs,
        pnl=dict(rev=rev, hpp=hpp, laba_kotor=laba_kotor,
                 laba_op=laba_op, laba_bersih=laba_bersih),
        # Likuiditas
        current_ratio=current_ratio, current_int=hi(current_ratio, 1.0, 1.5),
        quick_ratio=quick_ratio,     quick_int=hi(quick_ratio,   0.7, 1.0),
        cash_ratio=cash_ratio,       cash_int=hi(cash_ratio,     0.2, 0.5),
        # Solvabilitas
        der=der,                     der_int=lo(der,           1.0, 2.0),
        debt_to_asset=debt_to_asset, dta_int=lo(debt_to_asset, 0.4, 0.6),
        equity_ratio=equity_ratio,   eqr_int=hi(equity_ratio,  0.4, 0.6),
        # Profitabilitas (dalam %)
        gpm=gpm, gpm_int=hi(gpm, 10, 25),
        opm=opm, opm_int=hi(opm, 5,  15),
        npm=npm, npm_int=hi(npm, 3,  10),
        roa=roa, roa_int=hi(roa, 3,  8),
        roe=roe, roe_int=hi(roe, 8,  15),
        # Aktivitas
        inv_turnover=inv_turnover,                 inv_int=hi(inv_turnover, 3, 6),
        asset_turnover=asset_turnover,             at_int=hi(asset_turnover, 0.5, 1.0),
        receivables_turnover=receivables_turnover, recv_int=hi(receivables_turnover, 5, 10),
    )

def get_rekening_ids(conn):
    rows = conn.execute("SELECT id FROM akun WHERE is_rekening=1").fetchall()
    return [r['id'] for r in rows]

def get_rekening_saldo(conn):
    rows = conn.execute("""
        SELECT a.id, a.kode, a.nama, a.no_rekening,
               COALESCE(SUM(d.debit - d.kredit), 0) as saldo
        FROM akun a
        LEFT JOIN detail_jurnal d ON d.akun_id = a.id
        WHERE a.is_rekening = 1
        GROUP BY a.id ORDER BY a.kode
    """).fetchall()
    return [dict(r) for r in rows]

def calc_cashflow(conn, sd, ed):
    kas_ids = get_rekening_ids(conn)
    if not kas_ids:
        return {'masuk':0,'keluar':0,'saldo':0}
    ph = ','.join('?'*len(kas_ids))
    rows = conn.execute(f"""
        SELECT j.id, COALESCE(j.tipe_tx,'') AS tipe_tx,
               COALESCE(SUM(d.debit),0) AS debit,
               COALESCE(SUM(d.kredit),0) AS kredit
        FROM detail_jurnal d
        JOIN jurnal j ON j.id=d.jurnal_id
        WHERE d.akun_id IN ({ph}) AND j.tanggal>=? AND j.tanggal<=?
        GROUP BY j.id, j.tipe_tx
    """, kas_ids+[sd,ed]).fetchall()
    masuk = keluar = 0.0
    for row in rows:
        debit = float(row['debit'] or 0)
        kredit = float(row['kredit'] or 0)
        if (row['tipe_tx'] or '').upper() == 'TRANSFER':
            net = debit - kredit
            if net > 0:
                masuk += net
            elif net < 0:
                keluar += abs(net)
        else:
            masuk += debit
            keluar += kredit
    saldo = masuk - keluar
    return {'masuk':masuk,'keluar':keluar,'saldo':saldo}

def month_end(y, m):
    last = calendar.monthrange(y, m)[1]
    return date(y, m, last)

def get_akun_id(conn, kode):
    r = conn.execute("SELECT id FROM akun WHERE kode=?", (kode,)).fetchone()
    return r['id'] if r else None

def next_nomor_tx(conn, tanggal):
    """Nomor transaksi internal bulanan yang stabil dan tidak dipakai ulang."""
    try:
        ym = datetime.strptime(str(tanggal)[:10], '%Y-%m-%d').strftime('%Y%m')
    except ValueError:
        ym = date.today().strftime('%Y%m')
    prefix = f'TRX-{ym}-'
    max_existing = conn.execute(
        "SELECT COALESCE(MAX(CAST(SUBSTR(nomor_tx, 12) AS INTEGER)),0) "
        "FROM jurnal WHERE nomor_tx LIKE ?",
        (prefix + '%',)
    ).fetchone()[0] or 0
    row = conn.execute("SELECT last_number FROM tx_sequence WHERE periode=?", (ym,)).fetchone()
    last = max(int(max_existing), int(row['last_number']) if row else 0)
    next_number = last + 1
    conn.execute(
        "INSERT OR REPLACE INTO tx_sequence(periode,last_number) VALUES(?,?)",
        (ym, next_number)
    )
    return f'{prefix}{next_number:04d}'

def backfill_nomor_tx(conn):
    """Isi ID transaksi database lama secara idempotent setelah startup atau restore."""
    rows = conn.execute(
        "SELECT id,tanggal FROM jurnal WHERE nomor_tx IS NULL OR TRIM(nomor_tx)='' "
        "ORDER BY tanggal,id"
    ).fetchall()
    for row in rows:
        conn.execute("UPDATE jurnal SET nomor_tx=? WHERE id=?",
                     (next_nomor_tx(conn, row['tanggal']), row['id']))

def backfill_payment_journal_links(conn):
    """Hubungkan riwayat pelunasan lama ke jurnalnya bila pasangan datanya unik."""
    configs = [
        ('bayar_piutang', 'piutang', 'piutang_id', 'pelanggan', '1120', 'kredit', 'Pelunasan piutang: '),
        ('bayar_hutang', 'hutang', 'hutang_id', 'pemasok', '2100', 'debit', 'Pelunasan hutang: '),
    ]
    for payment_table, tracker_table, fk_col, party_col, akun_kode, amount_col, prefix in configs:
        rows = conn.execute(
            f"""SELECT b.id,b.tanggal,b.jumlah,t.{party_col} AS pihak
                FROM {payment_table} b
                JOIN {tracker_table} t ON t.id=b.{fk_col}
                WHERE b.jurnal_id IS NULL"""
        ).fetchall()
        for row in rows:
            matches = conn.execute(
                f"""SELECT j.id
                    FROM jurnal j
                    JOIN detail_jurnal d ON d.jurnal_id=j.id
                    JOIN akun a ON a.id=d.akun_id
                    WHERE j.tipe_tx='PELUNASAN' AND j.tanggal=?
                      AND j.keterangan=? AND a.kode=?
                      AND ABS(d.{amount_col}-?)<0.01
                      AND NOT EXISTS (
                          SELECT 1 FROM {payment_table} used WHERE used.jurnal_id=j.id
                      )
                    ORDER BY j.id""",
                (row['tanggal'], prefix + row['pihak'], akun_kode, row['jumlah'])
            ).fetchall()
            if len(matches) == 1:
                conn.execute(f"UPDATE {payment_table} SET jurnal_id=? WHERE id=?",
                             (matches[0]['id'], row['id']))

def backfill_asset_journal_links(conn):
    """Hubungkan aset_tetap lama ke jurnal sumber bila pasangan datanya unik."""
    kategori_kode = {'Peralatan': '1200', 'Kendaraan': '1210', 'Gedung': '1220'}
    rows = conn.execute(
        "SELECT id,nama,kategori,harga_beli,tanggal_beli,dibuat FROM aset_tetap WHERE jurnal_id IS NULL"
    ).fetchall()
    linked = 0

    def unique_id(match_rows):
        ids = []
        for row in match_rows:
            if row['id'] not in ids:
                ids.append(row['id'])
        return ids[0] if len(ids) == 1 else None

    for aset in rows:
        kode = kategori_kode.get((aset['kategori'] or '').strip())
        if not kode:
            continue
        nama = (aset['nama'] or '').strip()
        harga = float(aset['harga_beli'] or 0)
        tanggal = str(aset['tanggal_beli'] or '')[:10]
        jid = None

        if nama:
            matches = conn.execute("""
                SELECT DISTINCT j.id
                FROM jurnal j
                JOIN detail_jurnal d ON d.jurnal_id=j.id
                JOIN akun a ON a.id=d.akun_id
                WHERE j.tipe_tx='PENGELUARAN' AND j.kategori='INVESTASI'
                  AND j.tanggal=? AND a.kode=? AND d.debit>0 AND ABS(d.debit-?)<0.01
                  AND (j.keterangan LIKE ? OR j.aset_id=?)
                ORDER BY j.id
            """, (tanggal, kode, harga, f"%{nama}%", aset['id'])).fetchall()
            jid = unique_id(matches)

        if not jid:
            matches = conn.execute("""
                SELECT DISTINCT j.id
                FROM jurnal j
                JOIN detail_jurnal d ON d.jurnal_id=j.id
                JOIN akun a ON a.id=d.akun_id
                WHERE j.tipe_tx='PENGELUARAN' AND j.kategori='INVESTASI'
                  AND j.tanggal=? AND a.kode=? AND d.debit>0 AND ABS(d.debit-?)<0.01
                  AND NOT EXISTS (SELECT 1 FROM aset_tetap used WHERE used.jurnal_id=j.id)
                ORDER BY j.id
            """, (tanggal, kode, harga)).fetchall()
            jid = unique_id(matches)

        if not jid and (aset['dibuat'] or ''):
            matches = conn.execute("""
                SELECT DISTINCT j.id
                FROM jurnal j
                JOIN detail_jurnal d ON d.jurnal_id=j.id
                JOIN akun a ON a.id=d.akun_id
                WHERE j.tipe_tx='SALDO_AWAL_AKUN'
                  AND a.kode=? AND d.debit>0 AND ABS(d.debit-?)<0.01
                  AND ABS(CAST(strftime('%s', j.dibuat) AS INTEGER) - CAST(strftime('%s', ?) AS INTEGER)) <= 120
                ORDER BY j.id
            """, (kode, harga, aset['dibuat'])).fetchall()
            jid = unique_id(matches)

        if jid:
            conn.execute("UPDATE aset_tetap SET jurnal_id=? WHERE id=?", (jid, aset['id']))
            conn.execute(
                "UPDATE jurnal SET aset_id=COALESCE(aset_id, ?) WHERE id=? AND kategori='INVESTASI'",
                (aset['id'], jid)
            )
            linked += 1
    return linked

def insert_jurnal(conn, tanggal, keterangan, kategori, tipe_tx, entries, pihak=''):
    if not is_valid_iso_date(tanggal):
        raise ValueError('Tanggal jurnal tidak valid.')
    prepared = []
    total_debit = 0.0
    total_kredit = 0.0
    for akun_kode, debit, kredit in entries:
        try:
            debit = float(debit or 0)
            kredit = float(kredit or 0)
        except (ValueError, TypeError):
            raise ValueError('Nominal jurnal harus berupa angka yang valid.')
        if not math.isfinite(debit) or not math.isfinite(kredit) or debit < 0 or kredit < 0:
            raise ValueError('Nominal debit dan kredit jurnal tidak boleh negatif.')
        aid = get_akun_id(conn, akun_kode)
        if not aid:
            raise ValueError(f"Akun jurnal '{akun_kode}' tidak ditemukan.")
        if debit > 0 or kredit > 0:
            prepared.append((aid, debit, kredit))
            total_debit += debit
            total_kredit += kredit
    if total_debit <= 0 or total_kredit <= 0:
        raise ValueError('Jurnal harus memiliki nilai debit dan kredit lebih dari 0.')
    if abs(total_debit - total_kredit) > 0.01:
        raise ValueError('Total debit dan kredit jurnal harus seimbang.')

    nomor_tx = next_nomor_tx(conn, tanggal)
    cur = conn.execute(
        "INSERT INTO jurnal(tanggal,keterangan,kategori,tipe_tx,pihak,nomor_tx) VALUES(?,?,?,?,?,?)",
        (tanggal, keterangan, kategori, tipe_tx, (pihak or '').strip(), nomor_tx)
    )
    jid = cur.lastrowid
    conn.executemany(
        "INSERT INTO detail_jurnal(jurnal_id,akun_id,debit,kredit) VALUES(?,?,?,?)",
        [(jid, aid, debit, kredit) for aid, debit, kredit in prepared]
    )
    return jid

# Akun kontrol (punya buku pembantu sendiri) -> dilarang di-posting via jurnal manual
# supaya buku besar tetap rekonsiliasi dgn buku pembantunya:
#   1120 Piutang Usaha  <-> tabel piutang
#   2100 Hutang Usaha   <-> tabel hutang
#   1130 Persediaan     <-> stok produk + pergerakan_stok
CONTROL_ACCOUNTS = {'1120', '2100', '1130'}


def _akun_beban_berperan(conn, peran):
    """Kode akun BEBAN yang perannya di Laba Rugi = `peran` (PAJAK / BUNGA /
    PENYUSUTAN), atau None kalau belum ada.

    Satu-satunya cara yang boleh dipakai untuk menemukan akun pajak/bunga saat
    membuat jurnal. Jangan pernah lagi menulis '6170' atau '6160' mentah: user
    boleh memakai ulang nomor itu untuk hal lain (mis. Beban Pengiriman), dan
    kalau nomornya dipaksa, beban pajak/bunga tercampur ke beban yang bukan
    pajak/bunga tanpa ketahuan, lalu Laba Rugi menggolongkannya salah.
    Kalau hasilnya None, pemanggil WAJIB menolak dengan pesan yang mengarahkan
    ke Daftar Akun, bukan menebak akun lain.
    """
    r = conn.execute(
        "SELECT kode FROM akun WHERE COALESCE(peran,'')=? AND tipe='BEBAN' ORDER BY kode LIMIT 1",
        (peran,)).fetchone()
    return r['kode'] if r else None


def _pesan_tanpa_akun_berperan(label):
    return (f'Belum ada akun beban yang berperan sebagai {label}. Buka Daftar Akun, '
            f'buat atau pilih akun beban {label.lower()}, lalu set Peran di laporan menjadi "{label}".')


def _akun_hutang_pilihan(conn):
    """Akun kewajiban yang boleh dipilih sebagai akun hutang saat belanja.

    Semua akun bertipe LIABILITAS ditawarkan, supaya usaha bisa memisahkan hutang
    (mis. per pemasok besar, atau hutang usaha vs hutang jangka panjang) tanpa perlu
    minta kode program diubah.
    """
    return conn.execute(
        "SELECT kode, nama FROM akun WHERE tipe='LIABILITAS' ORDER BY kode"
    ).fetchall()


def control_accounts(conn):
    """CONTROL_ACCOUNTS + akun hutang yang BENAR-BENAR dipakai buku pembantu.

    Akun kontrol dilarang dijurnal manual supaya buku besar tetap cocok dengan buku
    pembantunya. Sejak akun hutang bisa dipilih user, daftar statis tidak lagi cukup:
    hutang yang tercatat di akun kustom (mis. 2105) juga punya buku pembantu di tabel
    `hutang`, jadi akun itu harus ikut terkunci. Kalau tidak, jurnal manual bisa
    menggeser saldo buku besar tanpa mengubah daftar Hutang, dan keduanya berselisih
    diam-diam. Diturunkan dari data (bukan dihardcode) supaya tidak perlu dirawat.
    """
    kode = set(CONTROL_ACCOUNTS)
    try:
        for r in conn.execute(
                "SELECT DISTINCT akun_kode FROM hutang "
                "WHERE akun_kode IS NOT NULL AND TRIM(akun_kode) <> ''"):
            kode.add(str(r['akun_kode']).strip())
    except Exception:
        pass          # tabel/kolom belum ada (DB lama) -> pakai daftar statis saja
    return kode

# Akun sistem bawaan yang kebetulan ASET bersaldo normal Debit tapi JELAS bukan kas/bank:
# piutang, persediaan, biaya dibayar dimuka, pajak dibayar dimuka, dan aset tetap.
# Semuanya punya logika sendiri di app (piutang, inventory, kredit PPN, penyusutan), jadi
# menandainya sebagai rekening akan merusak angka Kas & Bank serta Arus Kas.
AKUN_SISTEM_NON_REKENING = {'1120', '1130', '1140', '1150', '1160',
                            '1180', '1181', '1182', '1183',
                            '1200', '1210', '1220', '1290'}

def boleh_jadi_rekening(kode, tipe, saldo_normal):
    """Syarat akun boleh dijadikan rekening kas/bank. Satu aturan dipakai bersama oleh
    server (Bagan Akun, Setup Saldo Awal) dan tampilan, supaya tak pernah beda."""
    return (tipe == 'ASET' and saldo_normal == 'DEBIT'
            and str(kode) not in AKUN_SISTEM_NON_REKENING)

app.jinja_env.globals['boleh_jadi_rekening'] = boleh_jadi_rekening
# versi list (JSON-able) supaya form Bagan Akun bisa memakai aturan yang sama di sisi JS
app.jinja_env.globals['AKUN_SISTEM_NON_REKENING'] = sorted(AKUN_SISTEM_NON_REKENING)
# Kode akun bawaan: dipakai Daftar Akun untuk menandai "BAWAAN" & mengunci tombol hapus.
app.jinja_env.globals['AKUN_DEFAULT_KODE'] = AKUN_DEFAULT_KODE

def parse_manual_journal_lines(conn, akun_ids, debits, kredits):
    if len(akun_ids) != len(debits) or len(akun_ids) != len(kredits):
        raise ValueError('Rincian jurnal tidak lengkap.')
    lines = []
    total_debit = 0.0
    total_kredit = 0.0
    for i, akun_id_raw in enumerate(akun_ids, start=1):
        try:
            debit = parse_rp(debits[i - 1] or 0)
            kredit = parse_rp(kredits[i - 1] or 0)
        except (TypeError, ValueError):
            raise ValueError(f'Nominal jurnal baris {i} tidak valid.')
        if not math.isfinite(debit) or not math.isfinite(kredit) or debit < 0 or kredit < 0:
            raise ValueError(f'Nominal jurnal baris {i} tidak boleh negatif.')
        if debit > 0 and kredit > 0:
            raise ValueError(f'Baris jurnal {i} hanya boleh berisi debit atau kredit.')
        if debit <= 0 and kredit <= 0:
            continue
        try:
            akun_id = int(akun_id_raw)
        except (TypeError, ValueError):
            raise ValueError(f'Akun jurnal baris {i} tidak valid.')
        arow = conn.execute("SELECT kode FROM akun WHERE id=?", (akun_id,)).fetchone()
        if not arow:
            raise ValueError(f'Akun jurnal baris {i} tidak ditemukan.')
        if arow['kode'] in control_accounts(conn):
            raise ValueError(
                f'Baris {i}: akun {arow["kode"]} adalah akun kontrol (buku pembantu '
                f'piutang/hutang/persediaan) dan tidak boleh diisi lewat jurnal manual. '
                f'Pakai menu Pemasukan / Pengeluaran / Piutang / Hutang / Koreksi Stok '
                f'agar buku pembantu tetap sinkron dengan buku besar.')
        lines.append((akun_id, debit, kredit))
        total_debit += debit
        total_kredit += kredit
    if total_debit <= 0 or total_kredit <= 0:
        raise ValueError('Jurnal harus memiliki debit dan kredit lebih dari 0.')
    if abs(total_debit - total_kredit) > 0.01:
        raise ValueError('Total debit dan kredit harus seimbang.')
    return lines

def adjust_inventory_position(conn, produk_id, target_qty, target_hpp, tanggal, alasan, mode):
    """Set posisi stok/HPP produk dan jurnal selisih nilainya tanpa mengubah jurnal lama."""
    p = conn.execute("SELECT nama,stok,harga_beli FROM produk WHERE id=?", (produk_id,)).fetchone()
    if not p:
        return None, 0.0, 0.0
    target_qty = float(target_qty or 0)
    target_hpp = float(target_hpp or 0)
    stok_lama = float(p['stok'] or 0)
    hpp_lama = float(p['harga_beli'] or 0)
    delta_qty = round(target_qty - stok_lama, 6)
    delta_nilai = round((target_qty * target_hpp) - (stok_lama * hpp_lama), 2)
    if abs(delta_qty) < 0.000001 and abs(delta_nilai) < 0.01:
        return None, 0.0, 0.0

    is_opening = mode == 'SALDO_AWAL'
    akun_lawan = '3100' if is_opening else '5190'
    tipe_tx = 'SALDO_AWAL' if is_opening else 'KOREKSI_STOK'
    kategori = 'PENDANAAN' if is_opening else 'OPERASIONAL'
    label = 'Saldo awal persediaan' if is_opening else 'Koreksi persediaan'
    entries = []
    if delta_nilai > 0:
        entries = [('1130', delta_nilai, 0), (akun_lawan, 0, delta_nilai)]
    elif delta_nilai < 0:
        entries = [(akun_lawan, abs(delta_nilai), 0), ('1130', 0, abs(delta_nilai))]
    jid = insert_jurnal(conn, tanggal, f"{label}: {p['nama']} - {alasan}",
                        kategori, tipe_tx, entries) if entries else None

    conn.execute("UPDATE produk SET stok=?,harga_beli=? WHERE id=?",
                 (target_qty, target_hpp, produk_id))
    if abs(delta_qty) >= 0.000001 or abs(target_hpp - hpp_lama) >= 0.01:
        jenis = 'MASUK' if delta_qty > 0 else ('KELUAR' if delta_qty < 0 else 'OPNAME')
        record_stock_movement(conn, produk_id, tanggal, jenis, abs(delta_qty), target_hpp,
                              f"{label}: {alasan}", jid, 'SET_POSITION',
                              stok_lama, hpp_lama)
    return jid, delta_qty, delta_nilai

def update_average_cost(conn, produk_id, qty_masuk, harga_baru):
    """Perbarui HPP produk dengan metode rata-rata tertimbang sebelum stok ditambah."""
    p = conn.execute("SELECT stok,harga_beli FROM produk WHERE id=?", (produk_id,)).fetchone()
    if not p or qty_masuk <= 0 or harga_baru < 0:
        return
    stok_lama = max(float(p['stok'] or 0), 0)
    hpp_lama = float(p['harga_beli'] or 0)
    total_qty = stok_lama + qty_masuk
    if total_qty > 0:
        hpp_baru = ((stok_lama * hpp_lama) + (qty_masuk * harga_baru)) / total_qty
        conn.execute("UPDATE produk SET harga_beli=? WHERE id=?", (hpp_baru, produk_id))

def remove_from_average_cost(conn, produk_id, qty_keluar, harga_keluar):
    """Kurangi lapisan nilai stok dan hitung ulang HPP rata-rata untuk sisa barang."""
    p = conn.execute("SELECT stok,harga_beli FROM produk WHERE id=?", (produk_id,)).fetchone()
    if not p or qty_keluar <= 0:
        return
    stok_lama = max(float(p['stok'] or 0), 0)
    sisa_qty = stok_lama - qty_keluar
    if sisa_qty <= 0:
        return
    sisa_nilai = (stok_lama * float(p['harga_beli'] or 0)) - (qty_keluar * float(harga_keluar or 0))
    if sisa_nilai >= 0:
        conn.execute("UPDATE produk SET harga_beli=? WHERE id=?", (sisa_nilai / sisa_qty, produk_id))


# ── Rincian item transaksi: simpan, reverse, recompute ──────────────────────
def record_stock_movement(conn, produk_id, tanggal, jenis, qty, harga, keterangan, jurnal_id,
                          hpp_effect='NONE', stok_sebelum=None, hpp_sebelum=None):
    """Simpan posisi stok sebelum/sesudah agar reversal dapat diterapkan dengan tepat."""
    p = conn.execute("SELECT stok,harga_beli FROM produk WHERE id=?", (produk_id,)).fetchone()
    if not p:
        return
    conn.execute(
        """INSERT INTO pergerakan_stok(
               produk_id,tanggal,jenis,qty,harga,keterangan,jurnal_id,hpp_effect,
               stok_sebelum,hpp_sebelum,stok_sesudah,hpp_sesudah
           ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
        (produk_id, tanggal, jenis, qty, harga, keterangan, jurnal_id, hpp_effect,
         stok_sebelum, hpp_sebelum, p['stok'], p['harga_beli'])
    )

def record_non_sku_movement(conn, jurnal_id, tanggal, deskripsi, nilai_delta):
    """Catat valuasi persediaan yang tidak dilacak sebagai produk berkuantitas."""
    nilai_delta = float(nilai_delta or 0)
    if abs(nilai_delta) < 0.01:
        return
    conn.execute(
        """INSERT INTO pergerakan_persediaan_non_sku(jurnal_id,tanggal,deskripsi,nilai_delta)
           VALUES(?,?,?,?)""",
        (jurnal_id, tanggal, (deskripsi or 'Persediaan non-SKU').strip(), nilai_delta)
    )


def current_non_sku_value(conn):
    row = conn.execute(
        "SELECT COALESCE(SUM(nilai_delta),0) AS total FROM pergerakan_persediaan_non_sku"
    ).fetchone()
    return float((row['total'] if row and 'total' in row.keys() else 0) or 0)


def negative_stock_notice(conn, produk_ids):
    """Pesan peringatan bila ada produk yang stoknya menjadi minus setelah transaksi.

    Mengembalikan string peringatan, atau None bila semua stok masih >= 0. Stok minus
    diizinkan (mis. jual/pakai dulu, barang menyusul saat diterima) — ini sekadar
    pengingat agar pembelian/koreksi stok yang tertinggal dilengkapi."""
    ids = [int(i) for i in dict.fromkeys(produk_ids) if i]
    if not ids:
        return None
    rows = conn.execute(
        "SELECT nama,stok,satuan FROM produk WHERE id IN (%s) AND stok < -0.000001"
        % ','.join('?' * len(ids)),
        ids
    ).fetchall()
    if not rows:
        return None
    detail = ', '.join(
        "%s (%g %s)" % (r['nama'], float(r['stok'] or 0), (r['satuan'] or '').strip())
        for r in rows
    )
    return ('Stok minus tercatat: ' + detail +
            '. Transaksi tetap disimpan. Segera lengkapi pembelian atau koreksi stok '
            'agar nilai persediaan akurat.')


def returnable_non_sku_value(conn):
    """Nilai persediaan non-SKU yang sudah dikonsumsi (sale) tapi belum dikembalikan
    via retur. Dipakai untuk membatasi hpp_balik di retur penjualan mode 'umum',
    supaya 1130 tidak digelembungkan tanpa stok fisik.

    'returned' HANYA dihitung dari jurnal retur penjualan. Dulu semua nilai_delta
    positif dijumlahkan, padahal PEMBELIAN bahan baku non-SKU & penerimaan PO non-SKU
    juga positif. Akibatnya begitu usaha membeli bahan non-SKU lebih besar dari yang
    terpakai (kondisi normal, pasti ada sisa bahan), angka ini jatuh ke 0 dan retur
    penjualan ber-HPP jadi selalu ditolak "melebihi persediaan yang dapat dikembalikan".
    """
    row = conn.execute(
        """SELECT COALESCE(SUM(CASE WHEN nilai_delta < 0 THEN -nilai_delta ELSE 0 END),0) AS consumed
             FROM pergerakan_persediaan_non_sku"""
    ).fetchone()
    consumed = float((row['consumed'] if row and 'consumed' in row.keys() else 0) or 0)
    returned = sum(float(r['nilai'] or 0) for r in non_sku_return_jurnals(conn))
    return max(0.0, consumed - returned)

def sale_hpp_credit_breakdown(conn, jurnal_id):
    """Komposisi kredit HPP pada transaksi jual sumber.

    Dr 5100 bisa diimbangi Cr 1130 (persediaan) dan/atau Cr akun beban
    (reclass tenaga kerja/overhead/komisi). Dipakai retur agar pembalikan
    mengikuti sumbernya, bukan memaksa semua balik ke 1130.
    """
    if not jurnal_id:
        return {'total': 0.0, 'credits': {}}
    hpp_row = conn.execute(
        """SELECT COALESCE(SUM(d.debit),0) AS total
           FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id
           WHERE d.jurnal_id=? AND a.kode='5100'""",
        (jurnal_id,)
    ).fetchone()
    total_hpp = float((hpp_row['total'] if hpp_row else 0) or 0)
    if total_hpp <= 0:
        return {'total': 0.0, 'credits': {}}
    credit_codes = ['1130'] + sorted(set(HPP_RECLASS_ACCOUNTS.values()))
    rows = conn.execute(
        """SELECT a.kode, COALESCE(SUM(d.kredit),0) AS nilai
           FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id
           WHERE d.jurnal_id=? AND d.kredit>0 AND a.kode IN (%s)
           GROUP BY a.kode""" % ','.join('?' * len(credit_codes)),
        [jurnal_id] + credit_codes
    ).fetchall()
    credits = {r['kode']: float(r['nilai'] or 0) for r in rows if float(r['nilai'] or 0) > 0}
    credit_total = sum(credits.values())
    if credit_total <= 0:
        return {'total': 0.0, 'credits': {}}
    return {'total': min(total_hpp, credit_total), 'credits': credits}

def non_sku_return_jurnals(conn):
    """Daftar jurnal RETUR PENJUALAN yang mengembalikan nilai persediaan non-SKU.

    Hanya retur penjualan (tipe_tx RETUR_JUAL) yang menambah balik pool non-SKU
    tanpa tautan tracker ke transaksi asalnya. Pemasukan positif LAIN ke pool ini
    -- pembelian bahan baku non-SKU, penerimaan PO non-SKU, koreksi stok bertambah
    -- BUKAN retur dan tidak boleh ikut dihitung (lihat non_sku_return_overdraw_on_delete).
    """
    return conn.execute(
        """SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan,
                  COALESCE(SUM(n.nilai_delta),0) AS nilai
             FROM pergerakan_persediaan_non_sku n
             JOIN jurnal j ON j.id = n.jurnal_id
            WHERE n.nilai_delta > 0 AND j.tipe_tx = 'RETUR_JUAL'
            GROUP BY j.id
           HAVING nilai > 0
            ORDER BY j.tanggal DESC, j.id DESC"""
    ).fetchall()


def non_sku_return_overdraw_on_delete(conn, jurnal_id):
    """True bila MENGHAPUS transaksi ini akan menyisakan retur non-SKU tanpa konsumsi penopang
    (nilai Persediaan 1130 menggantung). Retur penjualan mode 'umum' sumber 'kas' membukukan
    Dr1130/Cr5100 tanpa tautan tracker ke sale, jadi transaction_dependency_count TIDAK
    memblokir hapus sale-nya. Bila total retur > (total 'consumed' - konsumsi non-SKU transaksi
    ini), hapus transaksi ini bikin retur>konsumsi → 1130 hantu. Blokir.

    PENTING: 'returned' dihitung HANYA dari jurnal retur penjualan. Dulu di sini semua
    nilai_delta positif dijumlahkan, padahal pembelian bahan baku non-SKU & penerimaan PO
    non-SKU juga positif. Akibatnya alur normal (beli bahan non-SKU lalu sebagian dipakai
    produksi) ikut terblokir, dan user disuruh menghapus 'retur' yang memang tidak pernah ada
    sehingga transaksinya buntu tak bisa dihapus sama sekali.
    """
    row = conn.execute(
        """SELECT COALESCE(SUM(CASE WHEN nilai_delta < 0 THEN -nilai_delta ELSE 0 END),0) AS consumed
             FROM pergerakan_persediaan_non_sku"""
    ).fetchone()
    consumed = float((row['consumed'] if row else 0) or 0)
    returned = sum(float(r['nilai'] or 0) for r in non_sku_return_jurnals(conn))
    this_row = conn.execute(
        "SELECT COALESCE(SUM(CASE WHEN nilai_delta<0 THEN -nilai_delta ELSE 0 END),0) c "
        "FROM pergerakan_persediaan_non_sku WHERE jurnal_id=?", (jurnal_id,)
    ).fetchone()
    this_consumed = float((this_row['c'] if this_row else 0) or 0)
    return this_consumed > 0.01 and returned > (consumed - this_consumed) + 0.01

def is_inventory_account(conn, kode):
    kode = (kode or '').strip()
    if not kode:
        return False
    row = conn.execute(
        """SELECT kode,nama,subtipe
           FROM akun
           WHERE kode=? AND tipe='ASET' AND saldo_normal='DEBIT'
           LIMIT 1""",
        (kode,)
    ).fetchone()
    if not row:
        return False
    label = f"{row['kode']} {row['nama'] or ''} {row['subtipe'] or ''}".lower()
    return kode.startswith('113') or 'persediaan' in label or 'inventory' in label

def find_open_tracker(conn, jenis, pihak, required_amount):
    """Cari tagihan aktif terbaru dengan sisa cukup untuk retur."""
    if jenis == 'PIUTANG':
        table, party_col = 'piutang', 'pelanggan'
    else:
        table, party_col = 'hutang', 'pemasok'
    return conn.execute(
        f"""SELECT * FROM {table}
            WHERE {party_col}=? AND status!='LUNAS'
              AND jumlah-terbayar>=?
            ORDER BY tanggal DESC,id DESC LIMIT 1""",
        ((pihak or '').strip(), float(required_amount or 0) - 0.01)
    ).fetchone()

def apply_tracker_adjustment(conn, jenis, record_id, jumlah_delta, jurnal_id):
    """Ubah nilai pokok piutang/hutang dan simpan relasi reversal-nya."""
    table = 'piutang' if jenis == 'PIUTANG' else 'hutang'
    conn.execute(f"UPDATE {table} SET jumlah=jumlah+? WHERE id=?", (jumlah_delta, record_id))
    if jenis == 'PIUTANG':
        update_piutang_status(conn, record_id)
    else:
        update_hutang_status(conn, record_id)
    conn.execute(
        "INSERT INTO penyesuaian_tagihan(jurnal_id,jenis,record_id,jumlah_delta) VALUES(?,?,?,?)",
        (jurnal_id, jenis, record_id, jumlah_delta)
    )

def _select_return_tracker(conn, jenis, record_id_raw, pihak, required_amount):
    table = 'piutang' if jenis == 'PIUTANG' else 'hutang'
    party_col = 'pelanggan' if jenis == 'PIUTANG' else 'pemasok'
    label = 'Piutang' if jenis == 'PIUTANG' else 'Hutang'
    record_id_raw = (record_id_raw or '').strip()
    if record_id_raw:
        try:
            record_id = int(record_id_raw)
        except (TypeError, ValueError):
            raise ValueError(f'{label} sumber retur tidak valid.')
        tracker = conn.execute(f"SELECT * FROM {table} WHERE id=?", (record_id,)).fetchone()
        if not tracker:
            raise ValueError(f'{label} sumber retur tidak ditemukan.')
        if tracker['status'] == 'LUNAS':
            raise ValueError(f'{label} sumber retur sudah lunas.')
        if pihak and (tracker[party_col] or '').strip() != (pihak or '').strip():
            raise ValueError(f'{label} sumber retur tidak cocok dengan nama pihak yang diisi.')
        return tracker
    return find_open_tracker(conn, jenis, pihak, required_amount)

def _tracker_tax_context(conn, jenis, tracker):
    if not tracker or not tracker['jurnal_id']:
        return {'ppn': 0.0, 'pph': 0.0, 'pph22': 0.0,
                'net': float((tracker['jumlah'] if tracker else 0) or 0), 'dpp': 0.0}
    delta = conn.execute(
        "SELECT COALESCE(SUM(jumlah_delta),0) FROM penyesuaian_tagihan WHERE jenis=? AND record_id=?",
        (jenis, tracker['id'])
    ).fetchone()[0] or 0
    original_net = max(0.0, float(tracker['jumlah'] or 0) - float(delta or 0))
    pph22 = 0.0
    if jenis == 'PIUTANG':
        ppn = conn.execute(
            """SELECT COALESCE(SUM(dj.kredit),0)
                 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id
                WHERE dj.jurnal_id=? AND a.kode='2111'""",
            (tracker['jurnal_id'],)
        ).fetchone()[0] or 0
        pph = conn.execute(
            """SELECT COALESCE(SUM(dj.debit),0)
                 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id
                WHERE dj.jurnal_id=? AND a.kode='1181'""",
            (tracker['jurnal_id'],)
        ).fetchone()[0] or 0
        pph22 = conn.execute(
            """SELECT COALESCE(SUM(dj.debit),0)
                 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id
                WHERE dj.jurnal_id=? AND a.kode='1183'""",
            (tracker['jurnal_id'],)
        ).fetchone()[0] or 0
    else:
        ppn = conn.execute(
            """SELECT COALESCE(SUM(dj.debit),0)
                 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id
                WHERE dj.jurnal_id=? AND a.kode='1180'""",
            (tracker['jurnal_id'],)
        ).fetchone()[0] or 0
        pph = conn.execute(
            """SELECT COALESCE(SUM(dj.kredit),0)
                 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id
                WHERE dj.jurnal_id=? AND a.kode='2112'""",
            (tracker['jurnal_id'],)
        ).fetchone()[0] or 0
    ppn = float(ppn or 0)
    pph = float(pph or 0)
    pph22 = float(pph22 or 0)
    # net = nilai pokok hutang/piutang yang sudah memperhitungkan pemotongan pajak penghasilan
    # (PPh23 + PPh22 untuk sisi PIUTANG). dpp = nilai sebelum PPN tapi sebelum potong PPh.
    dpp = max(0.0, original_net - ppn + pph + pph22)
    return {'ppn': ppn, 'pph': pph, 'pph22': pph22, 'net': original_net, 'dpp': dpp}

def _tracker_return_tax_parts(conn, jenis, tracker, amount, basis='NET'):
    amount = float(amount or 0)
    ctx = _tracker_tax_context(conn, jenis, tracker)
    if amount <= 0 or (ctx['ppn'] <= 0 and ctx['pph'] <= 0 and ctx.get('pph22', 0) <= 0):
        return {'dpp': amount, 'ppn': 0.0, 'pph': 0.0, 'pph22': 0.0, 'net': amount}
    if basis == 'DPP':
        ratio = min(1.0, amount / ctx['dpp']) if ctx['dpp'] > 0 else 0.0
        ppn = round(ctx['ppn'] * ratio, 2)
        pph = round(ctx['pph'] * ratio, 2)
        pph22 = round(ctx.get('pph22', 0) * ratio, 2)
        return {'dpp': round(amount, 2), 'ppn': ppn, 'pph': pph, 'pph22': pph22,
                'net': round(amount + ppn - pph - pph22, 2)}
    ratio = min(1.0, amount / ctx['net']) if ctx['net'] > 0 else 0.0
    ppn = round(ctx['ppn'] * ratio, 2)
    pph = round(ctx['pph'] * ratio, 2)
    pph22 = round(ctx.get('pph22', 0) * ratio, 2)
    dpp = round(max(0.0, amount - ppn + pph + pph22), 2)
    return {'dpp': dpp, 'ppn': ppn, 'pph': pph, 'pph22': pph22, 'net': round(amount, 2)}

def transaction_dependency_count(conn, jurnal_id):
    """Hitung pembayaran/retur/write-off yang masih bergantung pada transaksi sumber.
    Write-off piutang (writeoff_piutang) ikut dihitung: jurnal write-off (Dr 6190/1150,
    Cr 1120) hidup terpisah dari transaksi penjualannya, jadi menghapus/mengedit penjualan
    tanpa membatalkan write-off dulu akan meninggalkan jurnal write-off menggantung dan
    buku besar 1120 selisih dengan kartu piutang (senyap, karena tetap seimbang)."""
    return conn.execute(
        """SELECT
             (SELECT COUNT(*) FROM bayar_piutang b
              JOIN piutang p ON p.id=b.piutang_id WHERE p.jurnal_id=?)
           + (SELECT COUNT(*) FROM bayar_hutang b
              JOIN hutang h ON h.id=b.hutang_id WHERE h.jurnal_id=?)
           + (SELECT COUNT(*) FROM penyesuaian_tagihan x
              JOIN piutang p ON x.jenis='PIUTANG' AND p.id=x.record_id WHERE p.jurnal_id=?)
           + (SELECT COUNT(*) FROM penyesuaian_tagihan x
              JOIN hutang h ON x.jenis='HUTANG' AND h.id=x.record_id WHERE h.jurnal_id=?)
           + (SELECT COUNT(*) FROM writeoff_piutang w
              JOIN piutang p ON p.id=w.piutang_id WHERE p.jurnal_id=?)""",
        (jurnal_id, jurnal_id, jurnal_id, jurnal_id, jurnal_id)
    ).fetchone()[0]

def pinjaman_karyawan_repayment_count(conn, jurnal_id):
    """Jml jurnal PENGEMBALIAN (Cr 1160, tipe PELUNASAN) utk pihak yg sama dengan
    jurnal PEMBERIAN pinjaman (Dr 1160) ini. Pemberian & pengembalian sama-sama jurnal
    polos tanpa tracker; saldo pinjaman per karyawan diturunkan dari ledger 1160 per pihak.
    Menghapus pemberian yg sudah punya pengembalian akan meninggalkan kredit 1160
    menggantung (saldo karyawan jadi negatif, understate aset — senyap). Guard ini
    mewajibkan hapus pengembalian dulu (pola sama dgn transaction_dependency_count)."""
    src = conn.execute(
        "SELECT COALESCE(NULLIF(TRIM(j.pihak),''),'') AS pihak "
        "FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id JOIN jurnal j ON j.id=d.jurnal_id "
        "WHERE d.jurnal_id=? AND a.kode='1160' AND d.debit>0 LIMIT 1", (jurnal_id,)
    ).fetchone()
    if not src or not src['pihak']:
        return 0
    return conn.execute(
        "SELECT COUNT(*) FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id JOIN jurnal j ON j.id=d.jurnal_id "
        "WHERE a.kode='1160' AND d.kredit>0 AND j.tipe_tx='PELUNASAN' "
        "AND COALESCE(NULLIF(TRIM(j.pihak),''),'')=? AND j.id!=?",
        (src['pihak'], jurnal_id)
    ).fetchone()[0]

def has_later_stock_movements(conn, jurnal_id):
    """Deteksi adanya mutasi stok lebih baru utk produk yang sama (informasi, bukan blokir)."""
    return bool(conn.execute(
        """SELECT 1
           FROM pergerakan_stok current
           JOIN pergerakan_stok later
             ON later.produk_id=current.produk_id
            AND later.id>current.id
            AND later.jurnal_id<>current.jurnal_id
           WHERE current.jurnal_id=?
           LIMIT 1""",
        (jurnal_id,)
    ).fetchone())

def has_stock_movement_after_date(conn, jurnal_id, produk_ids, tanggal):
    """Tolak backdate yang melewati mutasi SKU lain karena basis HPP akan berubah."""
    produk_ids = sorted({int(pid) for pid in produk_ids if pid})
    if not produk_ids:
        return False
    placeholders = ','.join('?' * len(produk_ids))
    return bool(conn.execute(
        f"""SELECT 1 FROM pergerakan_stok
            WHERE produk_id IN ({placeholders})
              AND jurnal_id<>?
              AND tanggal>?
            LIMIT 1""",
        produk_ids + [jurnal_id, tanggal]
    ).fetchone())

def has_untracked_legacy_return_adjustment(conn, jurnal_id):
    """Retur lama berbasis tagihan tidak aman dibalik tanpa relasi tracker eksplisit."""
    j = conn.execute("SELECT tipe_tx FROM jurnal WHERE id=?", (jurnal_id,)).fetchone()
    if not j or j['tipe_tx'] not in ('RETUR_JUAL', 'RETUR_BELI'):
        return False
    if conn.execute("SELECT 1 FROM penyesuaian_tagihan WHERE jurnal_id=? LIMIT 1",
                    (jurnal_id,)).fetchone():
        return False
    return bool(conn.execute(
        """SELECT 1 FROM detail_jurnal d
           JOIN akun a ON a.id=d.akun_id
           WHERE d.jurnal_id=? AND a.kode IN ('1120','2100')
           LIMIT 1""",
        (jurnal_id,)
    ).fetchone())

def sync_tracker_due_date(conn, jurnal_id, jatuh_tempo):
    """Simpan perubahan JT tracker tanpa mengubah kronologi atau deskripsi transaksi sumber."""
    if not jurnal_id:
        return
    j = conn.execute("SELECT tx_meta FROM jurnal WHERE id=?", (jurnal_id,)).fetchone()
    if not j:
        return
    try:
        meta = json.loads((j['tx_meta'] if 'tx_meta' in j.keys() else '') or '{}')
    except Exception:
        meta = {}
    if meta:
        meta['jt'] = jatuh_tempo
        conn.execute("UPDATE jurnal SET tx_meta=? WHERE id=?", (json.dumps(meta), jurnal_id))
    conn.execute(
        "UPDATE invoice SET jatuh_tempo=? WHERE jurnal_id=? AND status='DRAFT'",
        (jatuh_tempo, jurnal_id)
    )

def tracker_has_structured_source(conn, jurnal_id):
    if not jurnal_id:
        return False
    row = conn.execute("SELECT tx_meta FROM jurnal WHERE id=?", (jurnal_id,)).fetchone()
    return bool(row and ((row['tx_meta'] if 'tx_meta' in row.keys() else '') or ''))

def _append_varian(deskripsi, varian):
    """Sisipkan varian ke deskripsi item jadi 'Nama (Varian)'. Idempoten walau varian PRODUK
    berubah kemudian: bila deskripsi sudah berakhir dengan tanda kurung '(...)' (artinya sudah
    memuat snapshot varian), TIDAK menempel lagi, supaya edit/restore transaksi setelah varian
    diganti tak menghasilkan 'Nama (Lama) (Baru)'."""
    deskripsi = (deskripsi or '').strip()
    varian = (varian or '').strip()
    if varian and not deskripsi.endswith(')'):
        return (deskripsi + ' (' + varian + ')').strip()
    return deskripsi

def _store_items(conn, jid, items):
    for it in items:
        qty = float(it.get('qty', 0) or 0)
        harga_raw = float(it.get('harga', 0) or 0)
        dis_raw = float(it.get('diskon', 0) or 0)
        sub = it.get('subtotal')
        if sub is None:
            sub = qty * harga_raw - dis_raw
        # Bulatkan ke 2 desimal SEBELUM disimpan. Harga per unit sering lahir dari pembagian
        # (mis. total 1.395.850 / qty 3 = 465283.3333333333) atau dari HPP rata-rata bergerak,
        # dan ekor presisi penuh itu ikut terbawa saat transaksi dibuka lagi untuk diedit:
        # tampil sebagai "angka aneh" 17.800,54545454 dan, kalau field belum sempat diformat,
        # bisa salah dibaca 10^n kali lebih besar. subtotal sengaja dibulatkan dari nilai MENTAH
        # (bukan dari harga yang sudah dibulatkan) supaya totalnya tetap sedekat mungkin ke nilai
        # asli. Kami TIDAK membulatkan produk.harga_beli (HPP rata-rata bergerak) di sini, karena
        # itu dipakai menilai persediaan dan pembulatan berulang bisa membuatnya melenceng dari
        # saldo buku besar 1130.
        harga = round(harga_raw, 2)
        dis = round(dis_raw, 2)
        sub = round(float(sub or 0), 2)
        deskripsi = it.get('deskripsi', '')
        pid = it.get('produk_id')
        # Snapshot varian produk ke deskripsi item saat transaksi disimpan, supaya struk &
        # detail transaksi menampilkan "Nama (Varian)" dan tetap akurat walau varian produk
        # diedit/dihapus di kemudian hari.
        if pid:
            _pv = conn.execute("SELECT varian FROM produk WHERE id=?", (pid,)).fetchone()
            if _pv:
                deskripsi = _append_varian(deskripsi, _pv['varian'])
        conn.execute(
            """INSERT INTO transaksi_item(jurnal_id,produk_id,deskripsi,qty,satuan,harga,diskon,subtotal,arah)
               VALUES(?,?,?,?,?,?,?,?,?)""",
            (jid, pid, deskripsi, qty, it.get('satuan', ''),
             harga, dis, sub, it.get('arah', 'NONE'))
        )

def _enrich_item_varian(conn, items):
    """Untuk TAMPILAN (struk/detail transaksi): tambahkan varian dari data produk saat ini bila
    deskripsi item (data lama, sebelum snapshot) belum memuatnya. Tidak menyentuh DB. Data baru
    yang sudah bervarian tidak berubah (idempoten via _append_varian)."""
    out = []
    for it in items:
        d = dict(it)
        pid = d.get('produk_id')
        if pid:
            pv = conn.execute("SELECT varian FROM produk WHERE id=?", (pid,)).fetchone()
            if pv:
                d['deskripsi'] = _append_varian(d.get('deskripsi', ''), pv['varian'])
        out.append(d)
    return out

def _resync_product_cost(conn, pid):
    """Hitung ulang produk.harga_beli sebagai (nilai persediaan tersisa / stok) dari
    pergerakan_stok yang TERSISA. Model 'COGS terkunci + sisa moving average':
    COGS tiap transaksi terkunci pada harga yang tercatat di gerakannya (tidak diutak-atik
    saat transaksi lain dihapus), dan rata-rata hanya mengatur nilai persediaan yang masih
    ada. Karena nilai = penjumlahan, hasilnya ORDER-INDEPENDENT → aman untuk hapus/edit
    transaksi lama tanpa merusak rantai. Nilai di-anchor pada SET_POSITION (opname) terakhir.
    TIDAK menyentuh produk.stok (qty dikelola terpisah & sudah benar)."""
    pid = int(pid)
    rows = conn.execute(
        "SELECT jenis,qty,harga,hpp_effect,stok_sesudah,hpp_sesudah "
        "FROM pergerakan_stok WHERE produk_id=? ORDER BY id", (pid,)
    ).fetchall()
    value = 0.0
    for m in rows:
        if (m['hpp_effect'] or '').strip() == 'SET_POSITION':
            value = float(m['stok_sesudah'] or 0) * float(m['hpp_sesudah'] or 0)   # anchor opname
        elif (m['jenis'] or '') == 'MASUK':
            value += float(m['qty'] or 0) * float(m['harga'] or 0)
        elif (m['jenis'] or '') == 'KELUAR':
            value -= float(m['qty'] or 0) * float(m['harga'] or 0)
    row = conn.execute("SELECT stok FROM produk WHERE id=?", (pid,)).fetchone()
    if not row:
        return
    stok = float(row['stok'] or 0)
    # Hanya set harga_beli bila stok positif & nilai sisa wajar. Saat stok<=0 atau nilai
    # negatif (jual > beli), biarkan harga_beli lama (sama spt update_average_cost yg skip).
    if stok > 1e-9 and value >= -1e-6:
        conn.execute("UPDATE produk SET harga_beli=? WHERE id=?", (max(0.0, value) / stok, pid))

def _reverse_tx(conn, jid):
    """Balik semua efek transaksi terstruktur (stok, piutang/hutang, jurnal, item)."""
    # Kembalikan kuota voucher yang dipakai transaksi ini (POS reversal/edit)
    for vp in conn.execute("SELECT voucher_id FROM voucher_pakai WHERE jurnal_id=?", (jid,)).fetchall():
        conn.execute("UPDATE voucher SET terpakai=MAX(0,COALESCE(terpakai,0)-1) WHERE id=?", (vp['voucher_id'],))
    conn.execute("DELETE FROM voucher_pakai WHERE jurnal_id=?", (jid,))
    for b in conn.execute("SELECT * FROM bayar_piutang WHERE jurnal_id=?", (jid,)).fetchall():
        conn.execute("DELETE FROM bayar_piutang WHERE id=?", (b['id'],))
        conn.execute("UPDATE piutang SET terbayar=MAX(0,terbayar-?) WHERE id=?",
                     (b['jumlah'], b['piutang_id']))
        update_piutang_status(conn, b['piutang_id'])
        _sync_invoice_from_piutang(conn, b['piutang_id'])
    # Write-off piutang yang dibuat jurnal ini (mis. saat jurnal WRITE_OFF_PIUTANG dihapus/di-
    # reverse): kembalikan porsi 'terbayar' yang berasal dari write-off supaya kartu piutang
    # hidup lagi dan tidak salah tetap berstatus LUNAS. Baris piutang-nya sendiri milik jurnal
    # penjualan (jurnal_id beda), jadi tak ikut terhapus di sini.
    for w in conn.execute("SELECT * FROM writeoff_piutang WHERE jurnal_id=?", (jid,)).fetchall():
        conn.execute("UPDATE piutang SET terbayar=MAX(0,terbayar-?) WHERE id=?",
                     (w['jumlah'], w['piutang_id']))
        update_piutang_status(conn, w['piutang_id'])
        _sync_invoice_from_piutang(conn, w['piutang_id'])
    conn.execute("DELETE FROM writeoff_piutang WHERE jurnal_id=?", (jid,))
    for b in conn.execute("SELECT * FROM bayar_hutang WHERE jurnal_id=?", (jid,)).fetchall():
        conn.execute("DELETE FROM bayar_hutang WHERE id=?", (b['id'],))
        conn.execute("UPDATE hutang SET terbayar=MAX(0,terbayar-?) WHERE id=?",
                     (b['jumlah'], b['hutang_id']))
        update_hutang_status(conn, b['hutang_id'])
    for x in conn.execute("SELECT * FROM penyesuaian_tagihan WHERE jurnal_id=?", (jid,)).fetchall():
        table = 'piutang' if x['jenis'] == 'PIUTANG' else 'hutang'
        conn.execute(f"UPDATE {table} SET jumlah=jumlah-? WHERE id=?",
                     (x['jumlah_delta'], x['record_id']))
        if x['jenis'] == 'PIUTANG':
            update_piutang_status(conn, x['record_id'])
        else:
            update_hutang_status(conn, x['record_id'])
    conn.execute("DELETE FROM penyesuaian_tagihan WHERE jurnal_id=?", (jid,))
    movements = conn.execute(
        """SELECT ps.*,j.tipe_tx
           FROM pergerakan_stok ps LEFT JOIN jurnal j ON j.id=ps.jurnal_id
           WHERE ps.jurnal_id=? ORDER BY ps.id DESC""",
        (jid,)
    ).fetchall()
    # Kembalikan QTY stok (operasi penjumlahan → order-independent). HPP/rata-rata TIDAK
    # disesuaikan inkremental di sini (dulu pakai update/remove_average_cost yg cuma benar
    # kalau membalik transaksi paling baru). Sebagai gantinya, setelah movement transaksi
    # ini dihapus, harga_beli dihitung ulang dari SISA pergerakan (model COGS-terkunci),
    # sehingga hapus/edit transaksi lama tetap akurat tanpa peduli urutan.
    _resync_pids = set()
    _needs_review_pids = set()
    for m in movements:
        if m['produk_id']:
            effect = (m['hpp_effect'] or '').strip()
            if effect == 'SET_POSITION' and m['stok_sebelum'] is not None:
                # Opname (penetapan posisi absolut) → pulihkan stok & HPP eksplisit ke sebelum.
                current = conn.execute("SELECT stok,harga_beli FROM produk WHERE id=?", (m['produk_id'],)).fetchone()
                if (current and m['stok_sesudah'] is not None
                        and abs(float(current['stok'] or 0) - float(m['stok_sesudah'] or 0)) < 0.000001
                        and abs(float(current['harga_beli'] or 0) - float(m['hpp_sesudah'] or 0)) < 0.01):
                    conn.execute("UPDATE produk SET stok=?,harga_beli=? WHERE id=?",
                                 (m['stok_sebelum'], m['hpp_sebelum'], m['produk_id']))
                elif m['stok_sesudah'] is not None:
                    conn.execute("UPDATE produk SET stok=stok-? WHERE id=?",
                                 (float(m['stok_sesudah'] or 0) - float(m['stok_sebelum'] or 0),
                                  m['produk_id']))
            elif m['jenis'] == 'KELUAR':
                conn.execute("UPDATE produk SET stok=stok+? WHERE id=?", (m['qty'], m['produk_id']))
                # Baris ini yang PERSIS menghabiskan stok produk ke 0 (stok_sesudah≈0). Kalau ada
                # pergerakan lain SESUDAHNYA (mis. restock), keputusan HPP restock itu kemungkinan
                # sudah mengasumsikan stok mulai dari nol. Menghapus baris ini "menghidupkan lagi"
                # unit lama yang TIDAK ADA cara aman untuk menebak otomatis: mau dianggap tetap
                # terpisah dari batch sesudahnya, atau dicampur jadi satu rata-rata baru, dua-duanya
                # sama-sama bisa salah tergantung niat sebenarnya (lihat memory hpp_zero_stock_
                # crossing_bug). Jadi resync OTOMATIS dilewati, biar tidak diam-diam salah hitung;
                # user diminta memverifikasi manual lewat Edit Produk.
                if m['stok_sesudah'] is not None and abs(float(m['stok_sesudah'])) < 0.000001:
                    later = conn.execute(
                        "SELECT 1 FROM pergerakan_stok WHERE produk_id=? AND id>? LIMIT 1",
                        (m['produk_id'], m['id'])
                    ).fetchone()
                    if later:
                        _needs_review_pids.add(int(m['produk_id']))
                        continue
                _resync_pids.add(int(m['produk_id']))
            elif m['jenis'] == 'MASUK':
                conn.execute("UPDATE produk SET stok=stok-? WHERE id=?", (m['qty'], m['produk_id']))
                _resync_pids.add(int(m['produk_id']))
    _resync_pids -= _needs_review_pids
    conn.execute("DELETE FROM pergerakan_stok WHERE jurnal_id=?", (jid,))
    for _pid in _resync_pids:
        _resync_product_cost(conn, _pid)
    for _pid in _needs_review_pids:
        p = conn.execute("SELECT kode,nama FROM produk WHERE id=?", (_pid,)).fetchone()
        if p:
            add_log(conn, 'Perlu cek HPP',
                    f"{p['kode']} - {p['nama']}: transaksi yang dihapus tadinya menghabiskan stok "
                    f"produk ini sampai 0, padahal ada transaksi lain sesudahnya (mis. restock). "
                    f"HPP produk ini TIDAK dihitung ulang otomatis, mohon periksa dan betulkan "
                    f"manual lewat Edit Produk bila perlu.", 'INPUT')
            flash(f'Perhatian: HPP untuk "{p["nama"]}" ({p["kode"]}) tidak dihitung ulang otomatis '
                  f'karena transaksi yang dihapus tadinya menghabiskan stok produk ini ke 0 dan ada '
                  f'transaksi lain (mis. restock) sesudahnya. Cek Log Aktivitas, lalu periksa HPP '
                  f'produk ini manual lewat Edit Produk kalau perlu.', 'warning')
    conn.execute("DELETE FROM pergerakan_persediaan_non_sku WHERE jurnal_id=?", (jid,))
    # Subledger audit backflush resep (Penjualan Umum): stok bahan sudah di-rollback via
    # pergerakan_stok di atas; di sini cukup bersihkan baris auditnya agar tak yatim.
    conn.execute("DELETE FROM penjualan_resep_bahan WHERE jurnal_id=?", (jid,))
    conn.execute("DELETE FROM piutang WHERE jurnal_id=?", (jid,))
    conn.execute("DELETE FROM hutang WHERE jurnal_id=?", (jid,))
    # Produksi sederhana: hapus header + bahan (stok sudah di-rollback via pergerakan_stok di atas)
    for _pr in conn.execute("SELECT id FROM produksi WHERE jurnal_id=?", (jid,)).fetchall():
        conn.execute("DELETE FROM produksi_bahan WHERE produksi_id=?", (_pr['id'],))
        conn.execute("DELETE FROM produksi_non_sku WHERE produksi_id=?", (_pr['id'],))
    conn.execute("DELETE FROM produksi WHERE jurnal_id=?", (jid,))
    # Aset tetap yang dibuat oleh jurnal ini → bersihkan penyusutannya lalu hapus baris aset,
    # supaya tidak ada aset "hantu" yang tertinggal saat transaksi pembeliannya dihapus.
    backfill_asset_journal_links(conn)
    for _a in conn.execute("SELECT id, nama FROM aset_tetap WHERE jurnal_id=?", (jid,)).fetchall():
        _purge_asset_depreciation(conn, _a['id'], _a['nama'])
        conn.execute("DELETE FROM aset_tetap WHERE id=?", (_a['id'],))
    conn.execute("DELETE FROM detail_jurnal WHERE jurnal_id=?", (jid,))
    conn.execute("DELETE FROM transaksi_item WHERE jurnal_id=?", (jid,))

def _purge_asset_depreciation(conn, aset_id, nama):
    """Batalkan & hapus semua jurnal penyusutan milik satu aset.
    Penyusutan baru tertaut via jurnal.aset_id; data lama dicocokkan via keterangan."""
    rows = conn.execute(
        """SELECT id FROM jurnal
           WHERE tipe_tx='PENYUSUTAN'
             AND (aset_id=? OR (aset_id IS NULL AND keterangan LIKE ?))""",
        (aset_id, f"Penyusutan {nama}:%")
    ).fetchall()
    for r in rows:
        _reverse_tx(conn, r['id'])
        conn.execute("DELETE FROM jurnal WHERE id=?", (r['id'],))
    return len(rows)

def _je(conn, jid, kode, debit, kredit):
    debit = float(debit or 0)
    kredit = float(kredit or 0)
    if not math.isfinite(debit) or not math.isfinite(kredit) or debit < 0 or kredit < 0:
        raise ValueError('Nominal debit dan kredit jurnal tidak boleh negatif.')
    aid = get_akun_id(conn, kode)
    if not aid:
        raise ValueError(f"Akun jurnal '{kode}' tidak ditemukan.")
    if debit > 0 or kredit > 0:
        conn.execute("INSERT INTO detail_jurnal(jurnal_id,akun_id,debit,kredit) VALUES(?,?,?,?)",
                     (jid, aid, debit, kredit))

def _assert_journal_balanced(conn, jid):
    row = conn.execute(
        "SELECT COALESCE(SUM(debit),0) debit,COALESCE(SUM(kredit),0) kredit "
        "FROM detail_jurnal WHERE jurnal_id=?",
        (jid,)
    ).fetchone()
    debit = float(row['debit'] or 0)
    kredit = float(row['kredit'] or 0)
    if debit <= 0 or kredit <= 0 or abs(debit - kredit) > 0.01:
        raise ValueError('Total debit dan kredit jurnal harus seimbang.')

HPP_RECLASS_ACCOUNTS = {
    'TENAGA_KERJA': '6100',
    'OVERHEAD': '6180',
    'KOMISI': '6140',
    'LAINNYA': '6180',
}
HPP_INVENTORY_TYPES = {'BAHAN', 'KEMASAN'}

def _add_amount(bucket, key, amount):
    amount = float(amount or 0)
    if amount > 0 and math.isfinite(amount):
        bucket[key] = bucket.get(key, 0.0) + amount

def _normalise_reclass_map(value):
    """Return {akun_kode: nilai} untuk biaya HPP yang mereklas beban, bukan persediaan."""
    if not value:
        return {}
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except (TypeError, ValueError):
            return {}
    if not isinstance(value, dict):
        return {}
    out = {}
    for kode, amount in value.items():
        kode = str(kode or '').strip()
        try:
            amount = float(amount or 0)
        except (TypeError, ValueError):
            continue
        if kode and amount > 0 and math.isfinite(amount):
            out[kode] = out.get(kode, 0.0) + amount
    return out

def _build_sale(conn, jid, tanggal, keterangan, items, meta):
    """Bangun ulang penjualan dari items + meta. Asumsi _reverse_tx sudah dipanggil."""
    akun_kas = meta.get('akun_kas', '1100')
    total_sub = 0.0; total_hpp_produk = 0.0; jasa_total = 0.0
    stock_need = {}
    stock_info = {}
    for it in items:
        qty = float(it.get('qty', 0) or 0); harga = float(it.get('harga', 0) or 0); dis = float(it.get('diskon', 0) or 0)
        if min(qty, harga, dis) < 0 or qty <= 0 or not all(math.isfinite(v) for v in (qty, harga, dis)):
            raise ValueError('Qty, harga, dan diskon item penjualan tidak valid.')
        if dis > qty * harga:
            raise ValueError('Diskon item penjualan tidak boleh melebihi subtotal bruto.')
        it['subtotal'] = qty * harga - dis
        if it.get('produk_id'):
            pid = int(it['produk_id'])
            p = conn.execute(
                "SELECT nama,harga_beli,satuan,stok FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
                (pid,)
            ).fetchone()
            if not p:
                raise ValueError('Produk penjualan tidak ditemukan atau sudah diarsipkan.')
            stock_need[pid] = stock_need.get(pid, 0.0) + qty
            stock_info[pid] = p
            total_sub += it['subtotal']
            total_hpp_produk += qty * p['harga_beli']
            if not it.get('satuan'): it['satuan'] = p['satuan']
            it['arah'] = 'KELUAR'
        elif (it.get('satuan', '') == 'jasa') or (it.get('arah') == 'JASA'):
            # Pendapatan jasa non-SKU → akun 4200, tidak punya HPP
            jasa_total += it['subtotal']
            it['arah'] = 'JASA'; it['satuan'] = 'jasa'
        else:
            total_sub += it['subtotal']
            it['arah'] = 'NONE'
    # Stok minus DIIZINKAN (jual dulu, barang menyusul). Tidak ada blokir di sini;
    # pengingat stok minus ditampilkan di halaman inventory dan setelah input.
    hpp_generik = float(meta.get('hpp_generik', 0) or 0)
    hpp_reclass = _normalise_reclass_map(meta.get('hpp_reclass'))
    hpp_reclass_total = sum(hpp_reclass.values())
    inventory_hpp = total_hpp_produk + hpp_generik
    total_hpp = inventory_hpp + hpp_reclass_total
    diskon = float(meta.get('diskon', 0) or 0); ongkir = float(meta.get('ongkir', 0) or 0); biaya = float(meta.get('biaya', 0) or 0)
    biaya_mp = float(meta.get('biaya_mp', 0) or 0)   # potongan/biaya marketplace (beban); 0 utk transaksi lama (tanpa field ini)
    uang_masuk_raw = float(meta.get('uang_masuk', 0) or 0)
    if (min(hpp_generik, hpp_reclass_total, diskon, ongkir, biaya, biaya_mp, uang_masuk_raw) < 0
            or not all(math.isfinite(v) for v in (hpp_generik, hpp_reclass_total, diskon, ongkir, biaya, biaya_mp, uang_masuk_raw))):
        raise ValueError('Nilai penjualan tidak boleh negatif.')
    total_sebelum_diskon = total_sub + jasa_total
    if diskon > total_sebelum_diskon + 0.01:
        raise ValueError('Diskon tidak boleh melebihi subtotal penjualan.')
    # Distribusikan diskon proporsional ke barang vs jasa agar keduanya tereduksi
    if total_sebelum_diskon > 0 and diskon > 0:
        ratio_barang = total_sub / total_sebelum_diskon
        diskon_barang = round(diskon * ratio_barang, 2)
        diskon_jasa   = diskon - diskon_barang
    else:
        diskon_barang = diskon
        diskon_jasa   = 0.0
    pendapatan_barang = max(0, total_sub - diskon_barang + ongkir + biaya)
    grand = pendapatan_barang + max(0, jasa_total - diskon_jasa)
    if grand <= 0:
        raise ValueError('Total penjualan harus lebih dari 0.')
    try:
        ppn_pct = parse_percentage(meta.get('ppn_persen', 0), 'Persen PPN')
        pph_pct = parse_percentage(meta.get('pph23_persen', 0), 'Persen PPh 23')
        pph22_pct = parse_percentage(meta.get('pph22_persen', 0), 'Persen PPh 22')
    except ValueError as ex:
        raise ValueError(str(ex))
    ppn_pct = max(0, min(100, ppn_pct))
    pph_pct = max(0, min(100, pph_pct))
    pph22_pct = max(0, min(100, pph22_pct))
    ppn_nom = round(grand * ppn_pct / 100.0, 2)
    pph_nom = round(grand * pph_pct / 100.0, 2)
    pph22_nom = round(grand * pph22_pct / 100.0, 2)
    gross_kewajiban = grand + ppn_nom - pph_nom - pph22_nom
    if biaya_mp > gross_kewajiban + 0.01:
        raise ValueError('Biaya marketplace tidak boleh melebihi nilai transaksi.')
    total_kewajiban = gross_kewajiban - biaya_mp   # potongan marketplace kurangi yg cair → kas net
    uang_masuk_basis = min(uang_masuk_raw, grand)
    ratio_bayar = (uang_masuk_basis / grand) if grand > 0 else 0
    uang_masuk = min(total_kewajiban, round(total_kewajiban * ratio_bayar, 2))
    piutang_nm = max(0, total_kewajiban - uang_masuk)
    if piutang_nm <= 0.01:
        piutang_nm = 0
    if uang_masuk > 0:
        if not is_rekening_kode(conn, akun_kas):
            raise ValueError('Rekening kas/bank tidak valid.')
        _je(conn, jid, akun_kas, uang_masuk, 0)
    if piutang_nm > 0: _je(conn, jid, '1120', piutang_nm, 0)
    if pph_nom > 0:
        _je(conn, jid, '1181', pph_nom, 0)
    if pph22_nom > 0:
        _je(conn, jid, '1183', pph22_nom, 0)
    if biaya_mp > 0:
        _je(conn, jid, '6150', biaya_mp, 0)   # Dr Beban Administrasi (potongan/biaya marketplace)
    pendapatan_jasa = max(0.0, jasa_total - diskon_jasa)
    # Akun pendapatan COA kustom (disimpan di meta saat input mode Umum): seluruh pendapatan
    # (barang + jasa) ke satu akun itu, konsisten dengan jalur input. Tanpa ini, edit rincian
    # akan mereklas pendapatan dari COA kustom ke 4100/4200 standar (salah klasifikasi jurnal).
    akun_pend_custom = (meta.get('akun_pendapatan_kode') or '').strip()
    if akun_pend_custom and conn.execute(
            "SELECT 1 FROM akun WHERE kode=? AND tipe='PENDAPATAN'", (akun_pend_custom,)).fetchone():
        total_pend = round(pendapatan_barang + pendapatan_jasa, 2)
        if total_pend > 0:
            _je(conn, jid, akun_pend_custom, 0, total_pend)
    else:
        if pendapatan_barang > 0:
            _je(conn, jid, '4100', 0, pendapatan_barang)
        if pendapatan_jasa > 0:
            _je(conn, jid, '4200', 0, pendapatan_jasa)
    if ppn_nom > 0:
        _je(conn, jid, '2111', 0, ppn_nom)
    if total_hpp > 0:
        _je(conn, jid, '5100', total_hpp, 0)
        if inventory_hpp > 0:
            _je(conn, jid, '1130', 0, inventory_hpp)
        for akun_beban, nilai in sorted(hpp_reclass.items()):
            _je(conn, jid, akun_beban, 0, nilai)
    _assert_journal_balanced(conn, jid)
    for it in items:
        if it.get('produk_id'):
            pid = int(it['produk_id'])
            p = conn.execute(
                "SELECT stok,harga_beli,nama FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
                (pid,)
            ).fetchone()
            if not p:
                raise ValueError('Produk penjualan tidak ditemukan atau sudah diarsipkan.')
            # Stok minus diizinkan: decrement langsung tanpa blokir (jual dulu, barang menyusul).
            conn.execute(
                "UPDATE produk SET stok=stok-? WHERE id=?",
                (float(it['qty']), pid)
            )
            record_stock_movement(conn, pid, tanggal, 'KELUAR', float(it['qty']),
                                  float(p['harga_beli'] if p else 0), keterangan, jid, 'REMOVE_LAYER',
                                  p['stok'] if p else None, p['harga_beli'] if p else None)
    if hpp_generik > 0:
        record_non_sku_movement(conn, jid, tanggal, f"HPP non-SKU: {keterangan}", -hpp_generik)
    _store_items(conn, jid, items)
    if piutang_nm > 0:
        conn.execute("INSERT INTO piutang(tanggal,jatuh_tempo,pelanggan,keterangan,jumlah,jurnal_id) VALUES(?,?,?,?,?,?)",
                     (tanggal, meta.get('jt') or None, meta.get('pihak') or 'Customer', keterangan, piutang_nm, jid))
    return grand, total_hpp

def _build_purchase(conn, jid, tanggal, keterangan, items, meta):
    """Bangun ulang pembelian/beban dari items + meta. Asumsi _reverse_tx sudah dipanggil."""
    akun_debit = meta.get('akun_debit', '1130')
    akun_kas = meta.get('akun_kas', '1100')
    akun_hutang = meta.get('akun_hutang') or ('2110' if meta.get('kategori') == 'PAJAK' else '2100')
    inventory_debit = is_inventory_account(conn, akun_debit)
    total = 0.0
    for it in items:
        qty = float(it.get('qty', 0) or 0); harga = float(it.get('harga', 0) or 0); dis = float(it.get('diskon', 0) or 0)
        if min(qty, harga, dis) < 0 or qty <= 0 or not all(math.isfinite(v) for v in (qty, harga, dis)):
            raise ValueError('Qty, harga, dan diskon item pembelian tidak valid.')
        if dis > qty * harga:
            raise ValueError('Diskon item pembelian tidak boleh melebihi subtotal bruto.')
        it['subtotal'] = qty * harga - dis
        total += it['subtotal']
        if it.get('produk_id'):
            p = conn.execute(
                "SELECT satuan FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
                (int(it['produk_id']),)
            ).fetchone()
            if p and not it.get('satuan'):
                it['satuan'] = p['satuan']
            if not p:
                raise ValueError('Produk pembelian tidak ditemukan atau sudah diarsipkan.')
        it['arah'] = 'MASUK' if it.get('produk_id') else ('NON_SKU' if inventory_debit else 'NONE')
    if total <= 0:
        raise ValueError('Total pembelian harus lebih dari 0.')
    # ── Pajak PPN Masukan (Dr 1180) & PPh 23 dipotong (Cr 2112) ──
    # Dibangun ulang dari PERSEN yang disimpan di meta saat input (_proses_pengeluaran, with_tax).
    # Tanpa ini, edit rincian pembelian ber-pajak akan MENGHILANGKAN baris pajak dan salah hitung
    # hutang. Persen dipakai ulang atas `total` (DPP) baru — pajak ikut mengikuti DPP hasil edit.
    try:
        _ppn_pct = max(0, min(100, parse_percentage(meta.get('ppn_persen', 0), 'Persen PPN')))
        _pph_pct = max(0, min(100, parse_percentage(meta.get('pph23_persen', 0), 'Persen PPh 23')))
    except ValueError as ex:
        raise ValueError(str(ex))
    splits = [
        {'kode': str(x.get('kode', '')).strip(), 'nominal': float(x.get('nominal', 0) or 0)}
        for x in (meta.get('kas_splits') or [])
        if str(x.get('kode', '')).strip() and float(x.get('nominal', 0) or 0) > 0
    ]
    _decimal_off = get_setting(conn, 'pref_decimal', '0') != '1'
    if _decimal_off:
        total = _round_rp(total)   # mode rupiah bulat: samakan dgn input (hindari sisa sub-rupiah)
    ppn_nom = round(total * _ppn_pct / 100.0, 2)
    pph_nom = round(total * _pph_pct / 100.0, 2)
    if _decimal_off:
        ppn_nom = _round_rp(ppn_nom)
        pph_nom = _round_rp(pph_nom)
    total_kewajiban = total + ppn_nom - pph_nom
    uang_keluar_raw = float(meta.get('uang_keluar', 0) or 0)
    if uang_keluar_raw < 0 or not math.isfinite(uang_keluar_raw):
        raise ValueError('Uang keluar tidak boleh negatif.')
    uang_keluar = min(uang_keluar_raw, total_kewajiban)   # cap ke kewajiban SETELAH pajak
    if _decimal_off and not splits:
        uang_keluar = _round_rp(uang_keluar)
    hutang_nm = max(0, total_kewajiban - uang_keluar)
    if hutang_nm <= 0.01:
        hutang_nm = 0
    _je(conn, jid, akun_debit, total, 0)
    if ppn_nom > 0:
        _je(conn, jid, '1180', ppn_nom, 0)   # Dr PPN Masukan (kredit pajak)
    split_total = sum(x['nominal'] for x in splits)
    if any(not math.isfinite(x['nominal']) for x in splits):
        raise ValueError('Nominal rekening pembayaran tidak valid.')
    if any(not is_rekening_kode(conn, x['kode']) for x in splits):
        raise ValueError('Rekening pembayaran tidak valid.')
    if splits and split_total > 0 and uang_keluar > 0:
        remaining = uang_keluar
        scaled = []
        for i, split in enumerate(splits):
            nominal = remaining if i == len(splits) - 1 else round(uang_keluar * split['nominal'] / split_total, 2)
            remaining -= nominal
            if nominal > 0:
                scaled.append({'kode': split['kode'], 'nominal': nominal})
                _je(conn, jid, split['kode'], 0, nominal)
        meta['kas_splits'] = scaled
    elif uang_keluar > 0:
        if not is_rekening_kode(conn, akun_kas):
            raise ValueError('Rekening kas/bank tidak valid.')
        _je(conn, jid, akun_kas, 0, uang_keluar)
    if pph_nom > 0:
        _je(conn, jid, '2112', 0, pph_nom)   # Cr Hutang PPh 23 dipotong (kita sbg pemotong)
    if hutang_nm > 0: _je(conn, jid, akun_hutang, 0, hutang_nm)
    _assert_journal_balanced(conn, jid)
    for it in items:
        if it.get('produk_id'):
            pid = int(it['produk_id'])
            p = conn.execute("SELECT stok,harga_beli FROM produk WHERE id=?", (pid,)).fetchone()
            update_average_cost(conn, pid, float(it['qty']), float(it['harga']))
            conn.execute("UPDATE produk SET stok=stok+? WHERE id=?", (float(it['qty']), pid))
            record_stock_movement(conn, pid, tanggal, 'MASUK', float(it['qty']),
                                  float(it['harga']), keterangan, jid, 'ADD_LAYER',
                                  p['stok'] if p else None, p['harga_beli'] if p else None)
        elif inventory_debit:
            record_non_sku_movement(conn, jid, tanggal, it.get('deskripsi'), it['subtotal'])
    _store_items(conn, jid, items)
    if hutang_nm > 0:
        conn.execute("INSERT INTO hutang(tanggal,jatuh_tempo,pemasok,keterangan,jumlah,jurnal_id,akun_kode) VALUES(?,?,?,?,?,?,?)",
                     (tanggal, meta.get('jt') or None, meta.get('pihak') or 'Pemasok',
                      keterangan, hutang_nm, jid, akun_hutang))
    return total


def update_piutang_status(conn, pid):
    p = conn.execute("SELECT jumlah,terbayar FROM piutang WHERE id=?", (pid,)).fetchone()
    if not p: return
    if p['jumlah'] - p['terbayar'] <= 0.01:
        s = 'LUNAS'
        conn.execute("UPDATE piutang SET terbayar=jumlah WHERE id=?", (pid,))
    elif p['terbayar'] > 0:
        s = 'SEBAGIAN'
    else:
        s = 'BELUM LUNAS'
    conn.execute("UPDATE piutang SET status=? WHERE id=?", (s, pid))

def _sync_invoice_from_piutang(conn, pid):
    """Samakan label status invoice dengan kondisi bayar piutangnya (untuk piutang yang berasal
    dari invoice). Piutang invoice lunas penuh -> invoice jadi PAID (Lunas); kalau balik belum
    lunas -> kembali ke SENT. HANYA mengubah label status invoice; tidak menyentuh kas/jurnal
    (penerimaan kas sudah ditangani alur Pelunasan). Dipanggil dari rute pembayaran, BUKAN dari
    update_piutang_status, supaya tidak rekursi dengan _settle_invoice_on_paid."""
    p = conn.execute("SELECT id,status,invoice_id,jurnal_id FROM piutang WHERE id=?", (pid,)).fetchone()
    if not p:
        return
    # Tautan utamanya jurnal_id; invoice_id hanya untuk piutang tanpa jurnal. invoice_id
    # hasil Restore "tambah data" versi lama bisa menunjuk invoice pelanggan LAIN, dan dulu
    # dipakai lebih dulu sehingga invoice orang lain ikut berubah jadi Lunas.
    inv = None
    if p['jurnal_id']:
        inv = conn.execute("SELECT id,status FROM invoice WHERE jurnal_id=? ORDER BY id DESC LIMIT 1",
                           (p['jurnal_id'],)).fetchone()
    elif p['invoice_id']:
        inv = conn.execute("SELECT id,status FROM invoice WHERE id=?", (p['invoice_id'],)).fetchone()
    if not inv:
        return
    if p['status'] == 'LUNAS' and inv['status'] in ('SENT', 'DRAFT'):
        conn.execute("UPDATE invoice SET status='PAID' WHERE id=?", (inv['id'],))
    elif p['status'] != 'LUNAS' and inv['status'] == 'PAID':
        conn.execute("UPDATE invoice SET status='SENT' WHERE id=?", (inv['id'],))

def update_hutang_status(conn, hid):
    h = conn.execute("SELECT jumlah,terbayar FROM hutang WHERE id=?", (hid,)).fetchone()
    if not h: return
    if h['jumlah'] - h['terbayar'] <= 0.01:
        s = 'LUNAS'
        conn.execute("UPDATE hutang SET terbayar=jumlah WHERE id=?", (hid,))
    elif h['terbayar'] > 0:
        s = 'SEBAGIAN'
    else:
        s = 'BELUM LUNAS'
    conn.execute("UPDATE hutang SET status=? WHERE id=?", (s, hid))

def piutang_status_label(row, today):
    if row['status'] == 'LUNAS':
        return 'LUNAS', 'success'
    sisa = row['jumlah'] - row['terbayar']
    if not row['jatuh_tempo']:
        return 'BELUM LUNAS', 'secondary'
    try:
        jt = datetime.strptime(str(row['jatuh_tempo'])[:10], '%Y-%m-%d').date()
    except (TypeError, ValueError):
        return 'TANGGAL TIDAK VALID', 'danger'
    delta = (jt - today).days
    if delta < 0:
        return 'LEWAT JATUH TEMPO', 'danger'
    elif delta <= 7:
        return 'MENDEKATI JATUH TEMPO', 'warning'
    return 'BELUM JATUH TEMPO', 'info'

def get_setting(conn, key, default=''):
    r = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return r['value'] if r else default

# ─── Modul PRO (locked by serial number) ─────────────────────────────────────
# Daftar modul yang bisa diaktivasi & serial number yang valid.
# Untuk produksi: serial sebaiknya per-customer (signed), saat ini placeholder.
MODUL_PRO = {
    'PAJAK_OTOMATIS': {
        'nama': 'Pajak & Gaji',
        'deskripsi': (
            'PPN & PPh 23: field pajak di invoice, otomasi jurnal saat posting, SPT Masa PPN, '
            'Bukti Potong, Bukti Bayar. '
            'Gaji & PPh 21: master karyawan, proses payroll bulanan, kalkulasi PPh 21 progresif, '
            'BPJS Kesehatan & Ketenagakerjaan, slip gaji cetak, jurnal otomatis.'
        ),
    },
    'ANGGARAN': {
        'nama': 'Anggaran & Target',
        'deskripsi': (
            'Tetapkan anggaran beban per akun COA dan target pendapatan setiap bulan. '
            'Pantau realisasi vs anggaran secara real-time dengan progress bar dan analisis selisih. '
            'Cocok untuk bisnis yang ingin kontrol pengeluaran dan presentasi ke investor atau bank.'
        ),
    },
    'POS_KASIR': {
        'nama': 'POS Kasir',
        'deskripsi': (
            'Portal kasir di /pos dengan login khusus, layar penjualan cepat, shift kasir, '
            'struk, pembayaran tunai/bank/piutang, dan posting otomatis ke jurnal, stok, HPP, '
            'serta laporan performa produk.'
        ),
    },
}

# ─── Serial number lisensi (disimpan sebagai HASH, bukan plaintext) ─────────
# Kode asli (16-digit, format XXXX-XXXX-XXXX-XXXX) TIDAK disimpan di source ini
# — hanya SHA-256 hash-nya. Ini mencegah siapa pun yang membaca/membongkar file
# ini (termasuk lewat AI coding agent / text editor) langsung melihat kode yang
# valid. Untuk menebak satu kode yang valid, penyerang harus brute-force ruang
# 16-digit (10 kuadriliun kombinasi) — praktis mustahil.
#
# Model lisensi MIXED:
#   1) _ALL_ACCESS_HASHES   → kode "bundle": sekali pakai, membuka SEMUA modul
#                             PRO (cocok utk customer yang beli paket lengkap).
#   2) _PER_MODULE_HASHES   → kode "satuan": hanya membuka SATU modul tertentu
#                             (cocok utk customer yang beli modul terpisah —
#                             kode utk modul A tidak bisa dipakai buka modul B).
_ALL_ACCESS_HASHES = frozenset([
    '6ecc8edca06f665ef2962a7f258b8cc7934ab66903998d7d9a477a9a22f7dd52',
    '6c00c1143e6dd975157bc89e9cbdbcada230af4b7ccb845354132ba548bb56e9',
    '5fb7403600349ec04bbc2a4f026c42cb9ce8e32f25f7b9523f579463b1d56f3e',
    '46d4e158a874f6c507b5b887e79445ccd9023bad6869fe7929be83609539ed69',
])

_PER_MODULE_HASHES = {
    'PAJAK_OTOMATIS': frozenset([
        '2a218ed869651cac59e6dfd6b92ba983a6b144462af5801573bf0654161f7ea9',
        'f41e1739bb65ebc4a35aae9f20f2f384a7da1f3213d035004fc797d5ea7a8dcc',
    ]),
    'ANGGARAN': frozenset([
        '8edde44499356041cac0fa2e94af2cbb41174032d23738a9bf90c7a4246fad59',
        'd07b99f58429a1fd7b3e8c8aef3cf4bc2c280b395426981cfb8c6ef4b4d3c4ef',
    ]),
    'POS_KASIR': frozenset([
        '22894f664ee1f3966625313d425670e6202723c4c1f5c353d7575884f4902dd7',
        'fe13a3b4298282cee97770d22676f53f8337153808592f2cf3bb7397f0ee597e',
    ]),
}

def _serial_hash(serial):
    """Normalisasi (strip + uppercase) lalu hash SHA-256 sebuah serial."""
    return sha256(str(serial).strip().upper().encode('utf-8')).hexdigest()

def is_module_active(conn, key):
    """Cek apakah modul PRO sudah diaktivasi."""
    return get_setting(conn, f'modul_{key}_aktif', '0') == '1'

def activate_module(conn, key, serial):
    """Coba aktivasi modul dengan serial number. Return (ok, message).
    Validasi dilakukan dengan mencocokkan HASH dari serial yang dimasukkan
    terhadap hash yang valid — serial asli tidak pernah dibandingkan atau
    disimpan dalam bentuk plaintext di source code.
    Sebuah serial dianggap valid utk modul `key` jika hash-nya ada di
    _ALL_ACCESS_HASHES (kode bundle, berlaku utk modul apa pun) ATAU di
    _PER_MODULE_HASHES[key] (kode satuan khusus modul ini)."""
    m = MODUL_PRO.get(key)
    if not m:
        return False, 'Modul tidak dikenali.'
    h = _serial_hash(serial)
    valid = (h in _ALL_ACCESS_HASHES) or (h in _PER_MODULE_HASHES.get(key, frozenset()))
    if not valid:
        return False, 'Serial number salah. Hubungi penjual untuk mendapatkan serial yang valid.'
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                 (f'modul_{key}_aktif', '1'))
    # PENTING: jangan simpan serial asli (plaintext) ke database — tabel `settings`
    # ikut ter-backup/ter-ekspor/ter-share, sehingga itu akan jadi celah kebocoran
    # baru meski source code sudah aman. Cukup simpan "petunjuk" 4 digit terakhir
    # (utk referensi dukungan pelanggan, mis. "kode yang dipakai berakhiran ...0338").
    raw = str(serial).strip()
    hint = ('•'*max(0, len(raw)-4)) + raw[-4:] if len(raw) >= 4 else '••••'
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                 (f'modul_{key}_serial_hint', hint))
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                 (f'modul_{key}_aktif_tgl', date.today().strftime('%Y-%m-%d')))
    return True, f'Modul "{m["nama"]}" berhasil diaktivasi.'

def deactivate_module(conn, key):
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                 (f'modul_{key}_aktif', '0'))

@app.context_processor
def inject_modul_pro():
    """Sediakan dict status modul PRO ke semua template untuk conditional UI."""
    try:
        conn = db()
        active = {k: is_module_active(conn, k) for k in MODUL_PRO}
        conn.close()
    except Exception:
        active = {k: False for k in MODUL_PRO}
    return {'modul_aktif': active}

def add_log(conn, aksi, detail='', kategori='LAINNYA'):
    uid   = session.get('user_id')
    uname = session.get('username', 'system')
    role  = session.get('role', '') or ''
    conn.execute(
        "INSERT INTO log_aktivitas(user_id, username, role, aksi, detail, kategori) VALUES(?,?,?,?,?,?)",
        (uid, uname, role, aksi, detail, kategori)
    )

def _edit_detail(ket, old_row, new_tanggal, old_total, new_total):
    """Susun detail log untuk transaksi yang diedit: tampilkan nilai LAMA → BARU,
    hanya untuk bagian yang benar-benar berubah, supaya audit log mudah dibaca.
    Contoh hasil: 'Bayar listrik | nominal Rp100.000 → Rp150.000'."""
    chg = []
    try:
        if abs(float(old_total or 0) - float(new_total or 0)) > 0.5:
            chg.append("nominal {} → {}".format(rp_filter(old_total), rp_filter(new_total)))
    except (TypeError, ValueError):
        pass
    old_ket = (old_row['keterangan'] if old_row else '') or ''
    if old_ket != (ket or ''):
        chg.append('ket: "{}" → "{}"'.format(old_ket, ket or ''))
    old_tgl = (old_row['tanggal'] if old_row else '') or ''
    if old_tgl != (new_tanggal or ''):
        chg.append("tgl: {} → {}".format(old_tgl, new_tanggal))
    return (ket or '') + (" | " + "; ".join(chg) if chg else " | (nilai tidak berubah)")

def _pos_module_active():
    conn = db()
    aktif = is_module_active(conn, 'POS_KASIR')
    conn.close()
    return aktif

def _active_pos_shift(conn):
    return conn.execute(
        """SELECT * FROM pos_shift
           WHERE user_id=? AND status='OPEN'
           ORDER BY id DESC LIMIT 1""",
        (session.get('user_id'),)
    ).fetchone()

def _next_pos_receipt(conn, tanggal):
    ym = str(tanggal).replace('-', '')[:8]
    prefix = f"POS-{ym}-"
    row = conn.execute(
        "SELECT nomor_struk FROM pos_sale WHERE nomor_struk LIKE ? ORDER BY nomor_struk DESC LIMIT 1",
        (prefix + '%',)
    ).fetchone()
    next_number = 1
    if row and row['nomor_struk']:
        try:
            next_number = int(str(row['nomor_struk']).rsplit('-', 1)[-1]) + 1
        except ValueError:
            next_number = 1
    return f"{prefix}{next_number:04d}"

# Ukuran kertas thermal yang didukung -> lebar AREA CETAK efektifnya (mm).
# Kertas 58mm hanya bisa dicetak selebar ~48mm dan 80mm selebar ~76mm karena ada
# margin mekanis di kiri-kanan printer. Struk selama ini dipaku 76mm, jadi 80 = default
# supaya pengguna lama tidak berubah hasil cetaknya.
POS_STRUK_AREA_MM = {'58': 48, '80': 76}
POS_STRUK_KERTAS_DEFAULT = '80'


def _pos_struk_kertas(conn):
    """Ukuran kertas struk pilihan user ('58' / '80'), divalidasi ke daftar yang didukung."""
    v = (get_setting(conn, 'pos_struk_kertas', POS_STRUK_KERTAS_DEFAULT) or '').strip()
    return v if v in POS_STRUK_AREA_MM else POS_STRUK_KERTAS_DEFAULT


def _pos_receipt_settings(conn):
    menu_size = get_setting(conn, 'pos_menu_size', 'medium')
    if menu_size not in ('small', 'medium', 'large'):
        menu_size = 'medium'
    icon_mode = get_setting(conn, 'pos_icon_mode', 'initial')
    if icon_mode not in ('custom', 'initial', 'code'):
        icon_mode = 'initial'
    default_payment = get_setting(conn, 'pos_default_payment', 'TUNAI')
    if default_payment not in ('TUNAI', 'REKENING', 'PIUTANG'):
        default_payment = 'TUNAI'
    d = {
        'pos_struk_nama': get_setting(conn, 'pos_struk_nama', '') or get_setting(conn, 'nama_usaha', 'Usaha Saya'),
        'pos_struk_subtitle': get_setting(conn, 'pos_struk_subtitle', 'Struk POS'),
        'pos_struk_alamat': get_setting(conn, 'pos_struk_alamat', ''),
        'pos_struk_telepon': get_setting(conn, 'pos_struk_telepon', ''),
        'pos_struk_header': get_setting(conn, 'pos_struk_header', ''),
        'pos_struk_footer': get_setting(conn, 'pos_struk_footer', 'Terima kasih.'),
        'pos_struk_label_kasir': get_setting(conn, 'pos_struk_label_kasir', 'Kasir'),
        'pos_struk_show_customer': get_setting(conn, 'pos_struk_show_customer', '1') == '1',
        'pos_struk_show_payment': get_setting(conn, 'pos_struk_show_payment', '1') == '1',
        # Lebar kertas printer thermal. Yang disimpan UKURAN KERTAS (58/80), sedangkan
        # yang dipakai mencetak adalah lebar AREA CETAK-nya (48/76 mm) -- kertas 80mm
        # tidak bisa dicetak penuh 80mm karena ada margin mekanis di kiri-kanan.
        # Default 80 supaya struk yang selama ini tercetak 76mm tidak berubah.
        'pos_struk_kertas': _pos_struk_kertas(conn),
        'pos_struk_lebar_mm': POS_STRUK_AREA_MM[_pos_struk_kertas(conn)],
        'pos_menu_size': menu_size,
        'pos_icon_mode': icon_mode,
        'pos_default_payment': default_payment,
        'pos_diskon_lock': get_setting(conn, 'pos_diskon_lock', '0') == '1',
        'pos_diskon_has_pin': bool((get_setting(conn, 'pos_diskon_pin', '') or '').strip()),
    }
    return d

def _save_voucher_scope(conn, vid, scope, produk_ids, customer_list, kategori_list=None):
    """Tulis ulang baris scope produk/kategori & pembatasan customer voucher (DELETE lalu INSERT).
    Pembatasan customer berlaku orthogonal (untuk scope apa pun)."""
    conn.execute("DELETE FROM voucher_produk WHERE voucher_id=?", (vid,))
    conn.execute("DELETE FROM voucher_customer WHERE voucher_id=?", (vid,))
    conn.execute("DELETE FROM voucher_kategori WHERE voucher_id=?", (vid,))
    if scope == 'PRODUK':
        for pid in (produk_ids or []):
            conn.execute("INSERT OR IGNORE INTO voucher_produk(voucher_id, produk_id) VALUES(?,?)", (vid, int(pid)))
    if scope == 'KATEGORI':
        for kat in (kategori_list or []):
            kat = (kat or '').strip()
            if kat:
                conn.execute("INSERT OR IGNORE INTO voucher_kategori(voucher_id, kategori) VALUES(?,?)", (vid, kat))
    for cn in (customer_list or []):
        cn = (cn or '').strip()
        if cn:
            conn.execute("INSERT OR IGNORE INTO voucher_customer(voucher_id, customer_nama) VALUES(?,?)", (vid, cn))

def _evaluate_voucher(conn, kode, subtotal, today_str, items=None, customer=None):
    """Validasi voucher & hitung potongan. Return dict {ok, error, voucher, potongan, label}.
    Terpusat supaya AJAX-cek (kasir) & enforce-checkout pakai logika sama.
    items   : list keranjang [{produk_id|id, qty, harga}] — dipakai utk scope PRODUK.
    customer: nama pelanggan — dipakai utk pembatasan per-customer (case-insensitive)."""
    kode = (kode or '').strip().upper()
    if not kode:
        return {'ok': False, 'error': 'Kode voucher kosong.'}
    v = conn.execute("SELECT * FROM voucher WHERE UPPER(kode)=?", (kode,)).fetchone()
    if not v:
        return {'ok': False, 'error': 'Kode voucher tidak ditemukan.'}
    if not v['aktif']:
        return {'ok': False, 'error': 'Voucher tidak aktif.'}
    if v['berlaku_dari'] and today_str < v['berlaku_dari']:
        return {'ok': False, 'error': 'Voucher belum berlaku.'}
    if v['berlaku_sampai'] and today_str > v['berlaku_sampai']:
        return {'ok': False, 'error': 'Voucher sudah kedaluwarsa.'}
    kuota = int(v['kuota'] or 0)
    if kuota > 0 and int(v['terpakai'] or 0) >= kuota:
        return {'ok': False, 'error': 'Kuota voucher sudah habis.'}
    # ── Pembatasan per-customer (opsional, orthogonal terhadap scope) ──
    cust_rows = conn.execute(
        "SELECT customer_nama FROM voucher_customer WHERE voucher_id=?", (v['id'],)).fetchall()
    if cust_rows:
        allowed_cust = {(r['customer_nama'] or '').strip().lower() for r in cust_rows}
        cust_norm = (customer or '').strip().lower()
        if not cust_norm:
            return {'ok': False, 'error': 'Voucher khusus customer tertentu — isi nama pelanggan dulu.'}
        if cust_norm not in allowed_cust:
            return {'ok': False, 'error': 'Voucher tidak berlaku untuk pelanggan ini.'}
    # ── Subset efektif sesuai BASIS (scope): NOTA / KATEGORI / PRODUK ──
    # Subset menentukan di mana minimum (nilai/qty) diukur DAN di mana potongan dihitung.
    scope = (v['scope'] if 'scope' in v.keys() else None) or 'NOTA'
    eff_subtotal = float(subtotal or 0)
    eff_qty = 0.0
    scope_label = ''
    if scope in ('PRODUK', 'KATEGORI'):
        if scope == 'PRODUK':
            prod_rows = conn.execute(
                "SELECT produk_id FROM voucher_produk WHERE voucher_id=?", (v['id'],)).fetchall()
            allowed_pid = {int(r['produk_id']) for r in prod_rows}
            if not allowed_pid:
                return {'ok': False, 'error': 'Voucher produk belum punya daftar produk.'}
            scope_label = ' khusus produk'
            empty_msg = 'Voucher khusus produk tertentu — produknya belum ada di keranjang.'
        else:  # KATEGORI (bisa beberapa kategori sekaligus)
            kat_rows = [(r['kategori'] or '').strip() for r in conn.execute(
                "SELECT kategori FROM voucher_kategori WHERE voucher_id=?", (v['id'],)).fetchall()
                if (r['kategori'] or '').strip()]
            if not kat_rows:
                # Fallback voucher lama (single syarat_kategori, belum ke-backfill).
                _legacy = (v['syarat_kategori'] if 'syarat_kategori' in v.keys() else '') or ''
                if _legacy.strip():
                    kat_rows = [_legacy.strip()]
            if not kat_rows:
                return {'ok': False, 'error': 'Voucher kategori belum punya kategori.'}
            kat_set = {k.lower() for k in kat_rows}
            ids = []
            for it in (items or []):
                try:
                    ids.append(int(it.get('produk_id', it.get('id'))))
                except (TypeError, ValueError, AttributeError):
                    continue
            allowed_pid = set()
            if ids:
                ph = ','.join('?' * len(ids))
                for r in conn.execute(
                        f"SELECT id, COALESCE(NULLIF(TRIM(kategori),''),'Produk') AS k FROM produk WHERE id IN ({ph})",
                        ids).fetchall():
                    if (r['k'] or '').strip().lower() in kat_set:
                        allowed_pid.add(int(r['id']))
            _kat_disp = ', '.join(kat_rows)
            scope_label = f' khusus kategori {_kat_disp}'
            empty_msg = f'Voucher khusus kategori "{_kat_disp}" — belum ada produk kategori itu di keranjang.'
        eff_subtotal = 0.0
        for it in (items or []):
            try:
                pid = int(it.get('produk_id', it.get('id')))
            except (TypeError, ValueError, AttributeError):
                continue
            if pid in allowed_pid:
                q = float(it.get('qty', 0) or 0)
                eff_subtotal += q * float(it.get('harga', 0) or 0)
                eff_qty += q
        eff_subtotal = round(eff_subtotal, 2)
        if eff_subtotal <= 0:
            return {'ok': False, 'error': empty_msg}
    else:  # NOTA → subset = seluruh keranjang
        for it in (items or []):
            try:
                eff_qty += float(it.get('qty', 0) or 0)
            except (TypeError, ValueError, AttributeError):
                continue
    # ── Ambang minimum (nilai dan/atau qty) pada subset ──
    min_belanja = float(v['min_belanja'] or 0)
    if eff_subtotal < min_belanja - 0.01:
        return {'ok': False, 'error': f'Min. belanja Rp {min_belanja:,.0f} belum tercapai.'}
    min_qty = float(v['min_qty'] if 'min_qty' in v.keys() else 0) or 0
    if min_qty > 0 and items is not None and eff_qty < min_qty - 1e-9:
        return {'ok': False, 'error': f'Min. {("%g" % min_qty)} qty belum tercapai.'}
    # ── Cara potong: TOTAL (sekali ke subtotal subset) atau PER_PRODUK (nilai per unit).
    #    PER_PRODUK hanya bermakna pd subset produk/kategori (butuh qty subset = eff_qty). ──
    pmode = (v['potongan_mode'] if 'potongan_mode' in v.keys() else 'TOTAL') or 'TOTAL'
    per_produk = (pmode == 'PER_PRODUK' and scope in ('PRODUK', 'KATEGORI'))
    nilai = float(v['nilai'] or 0)
    maks = float(v['maks_potongan'] or 0)
    if (v['tipe'] or 'NOMINAL') == 'PERSEN':
        pot = eff_subtotal * nilai / 100.0
        if maks > 0:
            # PER_PRODUK → batas potongan berlaku per unit (skala dgn qty subset).
            pot = min(pot, maks * eff_qty if per_produk else maks)
    else:
        # PER_PRODUK → Rp nilai dikali jumlah unit pada subset (per item dibeli).
        pot = nilai * eff_qty if per_produk else nilai
    pot = round(min(pot, eff_subtotal), 2)
    if pot <= 0:
        return {'ok': False, 'error': 'Potongan voucher tidak valid (0).'}
    _pp = ' /produk' if per_produk else ''
    if (v['tipe'] or 'NOMINAL') == 'PERSEN':
        label = f"{v['kode']} (-{('%g' % nilai)}%{_pp}{scope_label})"
    else:
        label = f"{v['kode']} (-Rp {nilai:,.0f}{_pp}{scope_label})"
    # ── Alokasi 'pot' ke tiap item IN-SCOPE (utk tampil harga per item di kasir & struk).
    #    Proporsional thd qty*harga dalam subset; item terakhir dapat sisa pembulatan; item
    #    di LUAR scope = 0 (voucher kategori/produk TIDAK mendiskon produk lain). Display-only,
    #    TIDAK mengubah akuntansi (jurnal tetap pakai diskon nota meta['diskon']). ──
    alloc = {}
    _inscope = []
    for it in (items or []):
        try:
            _pid = int(it.get('produk_id', it.get('id')))
        except (TypeError, ValueError, AttributeError):
            continue
        _g = float(it.get('qty', 0) or 0) * float(it.get('harga', 0) or 0)
        if _g <= 0:
            continue
        if scope == 'NOTA' or _pid in allowed_pid:
            _inscope.append((_pid, _g))
    _base = sum(g for _, g in _inscope)
    if _base > 0:
        _acc = 0.0
        for _i, (_pid, _g) in enumerate(_inscope):
            if _i == len(_inscope) - 1:
                _a = round(pot - _acc, 2)
            else:
                _a = round(pot * _g / _base, 2); _acc = round(_acc + _a, 2)
            alloc[str(_pid)] = round(alloc.get(str(_pid), 0.0) + max(0.0, _a), 2)
    return {'ok': True, 'error': '', 'voucher': v, 'potongan': pot, 'label': label, 'alloc': alloc}

@app.route('/pos')
def pos_home():
    if not _pos_module_active():
        return render_template('pos_locked.html'), 403
    if not session.get('user_id'):
        return redirect(url_for('pos_login'))
    if not has_permission('pos_sales'):
        flash('Akses POS hanya untuk user yang diberi hak Akses POS Kasir.', 'danger')
        return redirect(url_for('dashboard'))
    return redirect(url_for('pos_kasir'))

@app.route('/pos/login', methods=['GET', 'POST'])
def pos_login():
    if not _pos_module_active():
        return render_template('pos_locked.html'), 403
    if session.get('user_id') and has_permission('pos_sales'):
        return redirect(url_for('pos_kasir'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        ph = sha256(password.encode()).hexdigest()
        conn = db()
        user = conn.execute(
            "SELECT * FROM users WHERE username=? AND password_hash=? AND aktif=1",
            (username, ph)
        ).fetchone()
        user_perms = effective_permissions(user['role'], user['permissions'] if 'permissions' in user.keys() else None) if user else set()
        if user and 'pos_sales' in user_perms:
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['nama'] = user['nama'] or user['username']
            session['role'] = user['role']
            add_log(conn, 'Login POS', f"Role: {user['role']}", 'POS')
            conn.commit(); conn.close()
            return redirect(url_for('pos_kasir'))
        conn.close()
        flash('Username/password salah atau user tidak punya akses POS.', 'danger')
    return render_template('pos_login.html')

@app.route('/pos/logout')
def pos_logout():
    session.clear()
    return redirect(url_for('pos_login'))

# ── Backflush resep produksi (dipakai POS "jual item resep", Model B) ──────────
# Formula & pemetaan akun IDENTIK dengan jalur backflush di pemasukan() umum
# (bahan_sku -> Cr 1130 di-back pergerakan_stok; bahan_non_sku -> Cr 1130 pool non-SKU;
#  biaya_tambahan -> reclass Cr 6180 beban). Lihat pemasukan() cabang `if tmpl`.
def _recipe_consumption(conn, tmpl, qty, line_idx=0):
    """PURE (baca produk.harga_beli, TAK menulis). Hitung konsumsi bahan 1 baris resep
    utk `qty` menu. Return {cogs_sku, cogs_nonsku, reclass:{akun:nilai}, recs:[...], bahan_pids}."""
    qty = float(qty or 0)
    qty_hasil = float(tmpl['qty_hasil'] or 0)
    if qty_hasil <= 0 or not math.isfinite(qty_hasil):
        raise ValueError(f'Template resep "{tmpl["nama"]}" belum punya jumlah hasil yang valid. Perbaiki template dulu.')
    if qty <= 0 or not math.isfinite(qty):
        raise ValueError('Qty resep harus lebih dari 0.')
    scale = qty / qty_hasil
    cogs_sku = 0.0; cogs_nonsku = 0.0; reclass = {}; recs = []; bahan_pids = []
    for b in json.loads(tmpl['bahan_sku'] or '[]'):
        try:
            bpid = int(b.get('produk_id'))
        except (TypeError, ValueError):
            continue
        qb = float(b.get('qty', 0) or 0) * scale
        if qb <= 0:
            continue
        prow = conn.execute("SELECT id,nama,harga_beli FROM produk WHERE id=? AND COALESCE(aktif,1)=1", (bpid,)).fetchone()
        if not prow:
            raise ValueError(f'Bahan resep "{tmpl["nama"]}" tidak ditemukan atau sudah diarsipkan. Perbaiki template resep.')
        hb = float(prow['harga_beli'] or 0)
        st = qb * hb
        cogs_sku += st
        bahan_pids.append(bpid)
        recs.append({'produk_id': bpid, 'deskripsi': prow['nama'], 'menu_qty': qty,
                     'line_idx': line_idx, 'template_id': tmpl['id'], 'qty': qb,
                     'harga': hb, 'subtotal': st})
    for b in json.loads(tmpl['bahan_non_sku'] or '[]'):
        nvv = float(b.get('nilai', 0) or 0) * scale
        if nvv <= 0:
            continue
        cogs_nonsku += nvv
        recs.append({'produk_id': None, 'deskripsi': (b.get('deskripsi', '') or '').strip() or 'Bahan non-SKU',
                     'menu_qty': qty, 'line_idx': line_idx, 'template_id': tmpl['id'],
                     'qty': float(b.get('qty', 0) or 0) * scale, 'harga': 0.0, 'subtotal': nvv})
    biaya_konversi = float(tmpl['biaya_tambahan'] or 0) * scale
    if biaya_konversi > 0:
        _add_amount(reclass, '6180', biaya_konversi)
        recs.append({'produk_id': None, 'deskripsi': 'Biaya tambahan produksi (tenaga kerja/overhead)',
                     'menu_qty': qty, 'line_idx': line_idx, 'template_id': tmpl['id'],
                     'qty': qty, 'harga': biaya_konversi / qty if qty else 0.0, 'subtotal': biaya_konversi})
    return {'cogs_sku': cogs_sku, 'cogs_nonsku': cogs_nonsku, 'reclass': reclass,
            'recs': recs, 'bahan_pids': bahan_pids}

def _apply_recipe_backflush(conn, jid, tanggal, keterangan, recs):
    """Terapkan konsumsi bahan resep (recs dari _recipe_consumption) ke DB: potong stok bahan SKU
    + pergerakan_stok KELUAR + subledger penjualan_resep_bahan. TIDAK menulis jurnal (caller emit
    Dr5100/Cr1130/beban) & TIDAK record_non_sku_movement (caller agregasi). Reversal otomatis via
    _reverse_tx (pergerakan_stok + penjualan_resep_bahan sudah dibersihkan generik)."""
    for rec in recs:
        if rec.get('produk_id'):
            _cur = conn.execute("SELECT stok FROM produk WHERE id=?", (rec['produk_id'],)).fetchone()
            _stok_lama = float(_cur['stok'] or 0) if _cur else 0.0
            conn.execute("UPDATE produk SET stok=stok-? WHERE id=?", (rec['qty'], rec['produk_id']))
            record_stock_movement(conn, rec['produk_id'], tanggal, 'KELUAR', rec['qty'], rec['harga'],
                                  f"Bahan resep: {keterangan}", jid, 'REMOVE_LAYER', _stok_lama, rec['harga'])
        conn.execute(
            """INSERT INTO penjualan_resep_bahan(jurnal_id,template_id,line_idx,produk_id,deskripsi,menu_qty,qty,harga,subtotal,qty_returned)
               VALUES(?,?,?,?,?,?,?,?,?,0)""",
            (jid, rec['template_id'], rec['line_idx'], rec['produk_id'], rec['deskripsi'],
             rec['menu_qty'], rec['qty'], rec['harga'], rec['subtotal'])
        )

def _pos_stock_shortage(conn, items, recipe_lines):
    """Pre-flight cek stok utk POS bila toggle 'pos_block_minus' aktif.
    Hitung total kebutuhan stok per produk SKU (item langsung + bahan SKU resep) lalu
    bandingkan dengan stok sekarang. Return list pesan kekurangan (kosong = aman).
    Non-SKU (produk_id None) TIDAK dicek: valuasi persediaannya boleh minus by design.
    PURE: tidak menulis DB (memanggil _recipe_consumption yang juga pure)."""
    demand = {}   # produk_id -> total qty diminta transaksi ini
    for it in items:
        pid = it.get('produk_id')
        if pid:
            demand[int(pid)] = demand.get(int(pid), 0.0) + float(it.get('qty') or 0)
    for rl in recipe_lines:
        cons = _recipe_consumption(conn, rl['template'], rl['qty'], rl['line_idx'])
        for rec in cons['recs']:
            pid = rec.get('produk_id')
            if pid:
                demand[int(pid)] = demand.get(int(pid), 0.0) + float(rec.get('qty') or 0)
    short = []
    for pid, qneed in demand.items():
        row = conn.execute("SELECT nama, stok, satuan FROM produk WHERE id=?", (pid,)).fetchone()
        if not row:
            continue
        stok = float(row['stok'] or 0)
        if stok - qneed < -1e-6:
            sat = (row['satuan'] or '').strip()
            short.append("%s (stok %g%s, butuh %g)" % (
                row['nama'], stok, (' ' + sat) if sat else '', qneed))
    return short

def _pos_hidden_kategori(conn):
    """Kategori yang disembunyikan dari grid kasir (set huruf kecil, utk perbandingan).

    Disimpan sebagai JSON array nama kategori di setting 'pos_hide_kategori'. Dipakai
    supaya kategori non-dagangan (mis. "Bahan Baku") bisa disembunyikan sekali jalan,
    tanpa perlu meng-uncheck produknya satu per satu lewat POS Lanjutan — dan produk
    baru di kategori itu ikut tersembunyi otomatis."""
    try:
        raw = json.loads(get_setting(conn, 'pos_hide_kategori', '[]') or '[]')
        if not isinstance(raw, list):
            raw = []
    except (ValueError, TypeError):
        raw = []
    return {str(k).strip().lower() for k in raw if str(k).strip()}

def _pos_kategori_terpakai(conn):
    """Semua kategori yang muncul di grid kasir (barang jadi + racikan) beserta jumlah item.
    Dipakai panel "Kategori di Kasir" pada POS Master. Nama kategori diturunkan dgn aturan
    yang SAMA persis dengan pos_kasir(), supaya daftar di setelan tidak beda dengan layar kasir."""
    rows = conn.execute("""
        SELECT kategori, SUM(n) AS n FROM (
            SELECT COALESCE(NULLIF(TRIM(kategori),''), 'Produk') AS kategori, COUNT(*) AS n
            FROM produk WHERE COALESCE(aktif,1)=1
            GROUP BY 1
            UNION ALL
            SELECT COALESCE(NULLIF(TRIM(pt.kategori),''), NULLIF(TRIM(pj.kategori),''), 'Racikan') AS kategori,
                   COUNT(*) AS n
            FROM produksi_template pt
            LEFT JOIN produk pj ON pj.id = pt.produk_jadi_id
            WHERE COALESCE(pt.pos_aktif,0) = 1
            GROUP BY 1
        )
        GROUP BY kategori ORDER BY kategori COLLATE NOCASE
    """).fetchall()
    return [{'nama': r['kategori'], 'n': int(r['n'] or 0)} for r in rows]

@app.route('/pos/kasir')
@pos_required
def pos_kasir():
    conn = db()
    products = [dict(r) for r in conn.execute("""
        SELECT id,kode,nama,varian,satuan,harga_beli,harga_jual,stok,
               COALESCE(favorit,0) AS favorit,
               COALESCE(barcode,'') AS barcode,
               COALESCE(NULLIF(TRIM(kategori),''), 'Produk') AS kategori
        FROM produk
        WHERE COALESCE(aktif,1)=1 AND COALESCE(pos_tampil,1)=1
        ORDER BY kategori COLLATE NOCASE, nama COLLATE NOCASE
    """).fetchall()]
    # Mode blokir stok minus (toggle di POS Master). Default OFF = "jual dulu,
    # barang menyusul". Nilainya WAJIB ikut ke layar kasir: dulu tidak dikirim,
    # sehingga tampilan kasir memblokir stok 0 tanpa syarat (kartu di-disable,
    # qty dipangkas) padahal server sudah mengizinkan — setelan jadi tak berefek.
    pos_block_minus = get_setting(conn, 'pos_block_minus', '0') == '1'
    # POS Lanjutan: multi-harga produk. Saat aktif, tempelkan daftar level harga tambahan
    # (produk_harga) ke tiap produk. Harga "Utama" tetap = harga_jual (harga patokan di Stok).
    pos_multi_harga = get_setting(conn, 'pos_multi_harga_aktif', '0') == '1'
    if pos_multi_harga:
        _hg = {}
        for r in conn.execute("SELECT produk_id, label, harga FROM produk_harga WHERE harga > 0 ORDER BY id"):
            _hg.setdefault(r['produk_id'], []).append({'label': r['label'], 'harga': float(r['harga'] or 0)})
        for p in products:
            extra = _hg.get(p['id'], [])
            # Level Utama (default) + level tambahan. Hanya kirim >1 kalau ada harga tambahan.
            p['prices'] = ([{'label': 'Utama', 'harga': float(p['harga_jual'] or 0)}] + extra) if extra else []
    akun_kas = conn.execute("SELECT kode,nama,is_rekening FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    customer_names = [r['nama'] for r in conn.execute("SELECT nama FROM customer ORDER BY nama").fetchall()]
    shift = _active_pos_shift(conn)
    recent_sales = conn.execute("""
        SELECT ps.*, j.keterangan
        FROM pos_sale ps
        JOIN jurnal j ON j.id=ps.jurnal_id
        WHERE ps.shift_id=?
        ORDER BY ps.id DESC LIMIT 8
    """, (shift['id'],)).fetchall() if shift else []
    # Ringkasan per metode untuk modal tutup shift
    shift_summary = {}
    if shift:
        _rows = conn.execute("""
            SELECT metode_bayar, COUNT(*) cnt, COALESCE(SUM(total),0) total
            FROM pos_sale WHERE shift_id=? AND status='SELESAI'
            GROUP BY metode_bayar
        """, (shift['id'],)).fetchall()
        shift_summary = {r['metode_bayar']: {'cnt': r['cnt'], 'total': r['total']} for r in _rows}
        shift_summary['_total_cnt'] = sum(r['cnt'] for r in _rows)
        shift_summary['_total_rp']  = sum(r['total'] for r in _rows)
    pos_s = _pos_receipt_settings(conn)
    # Voucher aktif & masih bisa dipakai (tanggal valid + kuota sisa) untuk dropdown kasir.
    # Min. belanja TIDAK difilter di sini (tergantung keranjang) — divalidasi server saat dipakai.
    today_str = date.today().strftime('%Y-%m-%d')
    # Tampilkan SEMUA voucher aktif (tanpa filter tanggal/kuota) — ikut customprint:
    # validitas penuh (tanggal/kuota/min-belanja/scope) tetap dicek _evaluate_voucher saat dipakai,
    # supaya voucher berlaku masa depan tak hilang & bikin bingung kasir.
    voucher_options = conn.execute("""
        SELECT kode, tipe, nilai, min_belanja, maks_potongan
        FROM voucher
        WHERE aktif=1
        ORDER BY kode
    """).fetchall()
    # Resep produksi yang diaktifkan untuk POS (Model B: dijual = bahan otomatis terpotong).
    # Urutan harga: harga jual resep sendiri -> harga produk jadi yang tertaut -> 0.
    # Resep made-to-order sering tidak punya produk jadi, jadi harga resep yang dipakai.
    # Kategori: kategori resep sendiri -> kategori produk jadi tertaut -> 'Racikan'.
    recipes = [dict(r) for r in conn.execute("""
        SELECT pt.id, pt.nama,
               COALESCE(NULLIF(pt.harga_jual, 0), pj.harga_jual, 0) AS harga_jual,
               COALESCE(NULLIF(TRIM(pt.kategori), ''), NULLIF(TRIM(pj.kategori), ''), 'Racikan') AS kategori
        FROM produksi_template pt
        LEFT JOIN produk pj ON pj.id = pt.produk_jadi_id
        WHERE COALESCE(pt.pos_aktif, 0) = 1
        ORDER BY pt.nama COLLATE NOCASE
    """).fetchall()]
    # POS Lanjutan multi-harga untuk RACIKAN (analog produk): Utama = harga_jual resep,
    # + level tambahan dari produksi_template_harga. Hanya kirim >1 kalau ada harga tambahan.
    if pos_multi_harga:
        _rg = {}
        for r in conn.execute("SELECT template_id, label, harga FROM produksi_template_harga WHERE harga > 0 ORDER BY id"):
            _rg.setdefault(r['template_id'], []).append({'label': r['label'], 'harga': float(r['harga'] or 0)})
        for rc in recipes:
            extra = _rg.get(rc['id'], [])
            rc['prices'] = ([{'label': 'Utama', 'harga': float(rc['harga_jual'] or 0)}] + extra) if extra else []
    # Mode tampilan item di kasir (diatur di POS Master): 'both' (default) | 'produk' | 'resep'.
    # Beberapa usaha (mis. FnB made-to-order) cuma perlu menampilkan racikan, tanpa barang jadi.
    pos_show_mode = get_setting(conn, 'pos_show_mode', 'both')
    if pos_show_mode not in ('both', 'produk', 'resep'):
        pos_show_mode = 'both'
    hide_kategori = _pos_hidden_kategori(conn)
    conn.close()
    if pos_show_mode == 'produk':
        recipes = []
    elif pos_show_mode == 'resep':
        products = []
    # Sembunyikan kategori yang dimatikan di POS Master (mis. "Bahan Baku"). Dilakukan
    # SEBELUM `categories` dihitung supaya chip kategorinya ikut hilang, bukan cuma isinya.
    # Ini murni TAMPILAN — checkout tetap menerima produknya, sama seperti perilaku
    # pos_tampil per produk, supaya pesanan tertahan/rujukan lama tidak mendadak ditolak.
    if hide_kategori:
        products = [p for p in products if (p['kategori'] or '').strip().lower() not in hide_kategori]
        recipes = [r for r in recipes if (r['kategori'] or '').strip().lower() not in hide_kategori]
    categories = sorted({p['kategori'] for p in products} | {r['kategori'] for r in recipes})
    # Kelompokkan varian (key=nama) untuk kartu induk di grid kasir. products tetap
    # flat (dipakai PRODUCTS js + popup varian); product_groups menjaga urutan asli
    # (kategori, nama) lewat first-seen, tiap grup = {nama, kategori, members:[...]}.
    product_groups = []
    _grp_index = {}
    for p in products:
        key = (p['nama'] or '').strip().lower()
        gi = _grp_index.get(key)
        if gi is None:
            _grp_index[key] = len(product_groups)
            product_groups.append({'nama': p['nama'], 'kategori': p['kategori'], 'members': [p]})
        else:
            product_groups[gi]['members'].append(p)
    return render_template('pos_kasir.html',
        products=products, product_groups=product_groups, recipes=recipes,
        categories=categories, akun_kas=akun_kas,
        customer_names=customer_names, shift=shift, recent_sales=recent_sales,
        pos_s=pos_s, today=today_str,
        voucher_options=voucher_options,
        pos_multi_harga=pos_multi_harga, pos_block_minus=pos_block_minus,
        shift_summary=shift_summary)

@app.route('/pos/shift/open', methods=['POST'])
@pos_required
def pos_shift_open():
    conn = db()
    if _active_pos_shift(conn):
        conn.close()
        flash('Shift kasir masih terbuka.', 'info')
        return redirect(url_for('pos_kasir'))
    conn.execute(
        "INSERT INTO pos_shift(user_id,username,saldo_awal,status) VALUES(?,?,0,'OPEN')",
        (session['user_id'], session.get('username', ''))
    )
    add_log(conn, 'Buka shift POS', f"Dibuka oleh {session.get('username','')}", 'POS')
    conn.commit(); conn.close()
    flash('Shift POS dibuka. Kasir siap dipakai.', 'success')
    return redirect(url_for('pos_kasir'))

@app.route('/pos/shift/close', methods=['POST'])
@pos_required
def pos_shift_close():
    conn = db()
    shift = _active_pos_shift(conn)
    if not shift:
        conn.close()
        flash('Tidak ada shift POS yang sedang terbuka.', 'warning')
        return redirect(url_for('pos_kasir'))
    # Hitung total tunai otomatis — rekonsiliasi kas ada di Buku Besar
    tunai = conn.execute(
        "SELECT COALESCE(SUM(total),0) FROM pos_sale WHERE shift_id=? AND metode_bayar='TUNAI' AND status='SELESAI'",
        (shift['id'],)
    ).fetchone()[0] or 0
    total_trx = conn.execute(
        "SELECT COUNT(*) FROM pos_sale WHERE shift_id=? AND status='SELESAI'", (shift['id'],)
    ).fetchone()[0]
    conn.execute(
        """UPDATE pos_shift
           SET ditutup=CURRENT_TIMESTAMP, kas_sistem=?, kas_fisik=0, selisih=0,
               status='CLOSED', catatan=?
           WHERE id=?""",
        (float(tunai), request.form.get('catatan', '').strip(), shift['id'])
    )
    add_log(conn, 'Tutup shift POS', f"Shift #{shift['id']} | {total_trx} transaksi | Tunai: {tunai}", 'POS')
    conn.commit(); conn.close()
    flash('Shift POS ditutup.', 'success')
    return redirect(url_for('pos_kasir'))

@app.route('/pos/checkout', methods=['POST'])
@pos_required
def pos_checkout():
    conn = db()
    # ── Idempotency guard (cegah transaksi dobel dari klik ganda / koneksi lambat) ──
    # Browser membuat idem_key sekali tiap modal Bayar dibuka & mengirimnya di setiap
    # percobaan submit. Kalau key ini sudah pernah berhasil tersimpan sebagai penjualan,
    # request ini adalah pengulangan (dobel klik / retry jaringan) -> jangan diproses ulang,
    # cukup arahkan ke struk yang sudah ada. Murni SELECT, tanpa efek samping ke stok/jurnal.
    # SENGAJA dicek PALING AWAL, SEBELUM cek shift aktif: replay yang sah (retry dari
    # transaksi yang sudah sukses tersimpan) harus tetap dikenali biarpun shift sudah
    # ditutup di antara request asli & request ulang. Kalau dicek SETELAH shift check,
    # replay yang datang setelah shift ditutup malah nyasar ke pesan "buka shift dulu"
    # padahal transaksinya sudah tercatat -- pas di skenario yang harusnya dilindungi.
    idem_key = (request.form.get('idem_key', '') or '').strip()[:64]
    if idem_key:
        _dup = conn.execute(
            "SELECT jurnal_id FROM pos_sale WHERE idempotency_key=?", (idem_key,)
        ).fetchone()
        if _dup:
            conn.close()
            return redirect(url_for('pos_struk', id=_dup['jurnal_id']))
    shift = _active_pos_shift(conn)
    if not shift:
        conn.close()
        flash('Buka shift kasir dulu sebelum checkout.', 'warning')
        return redirect(url_for('pos_kasir'))
    try:
        cart = json.loads(request.form.get('cart_json', '[]') or '[]')
    except json.JSONDecodeError:
        cart = []
    pos_s = _pos_receipt_settings(conn)
    metode = request.form.get('metode_bayar', pos_s['pos_default_payment'])
    if metode not in ('TUNAI', 'REKENING', 'PIUTANG'):
        metode = 'TUNAI'
    akun_kas = request.form.get('akun_kas', '1100').strip() or '1100'
    if metode in ('TUNAI', 'REKENING') and not is_rekening_kode(conn, akun_kas):
        conn.close()
        flash('Rekening pembayaran POS tidak valid.', 'danger')
        return redirect(url_for('pos_kasir'))
    try:
        diskon = parse_nonnegative_rp(request.form.get('diskon_nota', '0'), 'Diskon nota')
        bayar = parse_nonnegative_rp(request.form.get('bayar', '0'), 'Nominal bayar')
    except ValueError as ex:
        conn.close()
        flash(str(ex), 'danger')
        return redirect(url_for('pos_kasir'))

    # POS Lanjutan multi-harga: bila aktif, kasir boleh memilih level harga per baris.
    # Harga yang dikirim browser HANYA diterima kalau cocok dengan salah satu level yang
    # sudah dikonfigurasi owner (tabel produk_harga) — di luar itu tetap pakai harga
    # patokan (produk.harga_jual). Ini menjaga guard anti-manipulasi payload tetap ketat.
    pos_multi_harga = get_setting(conn, 'pos_multi_harga_aktif', '0') == '1'
    _allowed_harga = {}
    _allowed_harga_resep = {}
    if pos_multi_harga:
        for _r in conn.execute("SELECT produk_id, harga FROM produk_harga WHERE harga > 0"):
            _allowed_harga.setdefault(_r['produk_id'], set()).add(round(float(_r['harga']), 2))
        for _r in conn.execute("SELECT template_id, harga FROM produksi_template_harga WHERE harga > 0"):
            _allowed_harga_resep.setdefault(_r['template_id'], set()).add(round(float(_r['harga']), 2))

    items = []
    recipe_lines = []   # baris resep produksi (Model B: bahan backflush saat checkout)
    for idx, raw in enumerate(cart, start=1):
        if (raw.get('type') or '') == 'resep':
            try:
                tid = int(raw.get('rid') or raw.get('id'))
                qty = parse_positive_qty(raw.get('qty', 1), f'Qty item {idx}')
            except (TypeError, ValueError) as ex:
                conn.close(); flash(str(ex), 'danger'); return redirect(url_for('pos_kasir'))
            tmpl = conn.execute(
                "SELECT * FROM produksi_template WHERE id=? AND COALESCE(pos_aktif,0)=1", (tid,)
            ).fetchone()
            if not tmpl:
                conn.close()
                flash(f'Resep item {idx} tidak tersedia untuk POS (mungkin dinonaktifkan). Muat ulang halaman kasir.', 'danger')
                return redirect(url_for('pos_kasir'))
            # Harga resep SELALU dihitung server (harga jual resep -> harga produk jadi), TIDAK
            # dari cart_json. UI kasir tak punya field edit harga (kartu dirender dari harga
            # server yang sama), jadi nilai 'harga' kiriman browser tidak dipercaya — cegah
            # manipulasi payload lewat DevTools/API untuk jual di bawah harga.
            harga = float((tmpl['harga_jual'] if 'harga_jual' in tmpl.keys() else 0) or 0)
            if harga <= 0 and tmpl['produk_jadi_id']:
                _pj = conn.execute("SELECT harga_jual FROM produk WHERE id=?", (tmpl['produk_jadi_id'],)).fetchone()
                if _pj:
                    harga = float(_pj['harga_jual'] or 0)
            # Multi-harga: terima level harga pilihan kasir HANYA bila termasuk level yang
            # dikonfigurasi utk racikan ini (produksi_template_harga). Level "Utama" = harga_jual
            # (sudah jadi default di atas). Di luar itu tetap pakai harga server (anti-manipulasi).
            if pos_multi_harga:
                try:
                    _want = round(float(raw.get('harga') or 0), 2)
                except (TypeError, ValueError):
                    _want = 0
                if _want > 0 and _want != round(harga, 2) and _want in _allowed_harga_resep.get(tid, set()):
                    harga = _want
            if harga < 0 or not math.isfinite(harga):
                conn.close(); flash(f'Harga item {idx} tidak valid.', 'danger'); return redirect(url_for('pos_kasir'))
            recipe_lines.append({'template': tmpl, 'qty': qty, 'harga': harga, 'line_idx': idx})
            # Baris resep = pendapatan non-SKU utk _build_sale (Cr 4100); COGS + backflush ditambah setelahnya.
            items.append({'produk_id': None, 'deskripsi': tmpl['nama'], 'qty': qty,
                          'satuan': 'porsi', 'harga': harga, 'diskon': 0, 'arah': 'NONE'})
            continue
        try:
            pid = int(raw.get('id'))
            qty = parse_positive_qty(raw.get('qty', 1), f'Qty item {idx}')
        except (TypeError, ValueError) as ex:
            conn.close()
            flash(str(ex), 'danger')
            return redirect(url_for('pos_kasir'))
        p = conn.execute(
            "SELECT * FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
            (pid,)
        ).fetchone()
        if not p:
            conn.close()
            flash(f'Produk item {idx} tidak ditemukan atau sudah diarsipkan.', 'danger')
            return redirect(url_for('pos_kasir'))
        # Stok minus diizinkan (jual dulu, barang menyusul) — tidak ada blokir di sini.
        # Harga jual SELALU dari server (produk.harga_jual), TIDAK dari cart_json. UI kasir tak
        # punya field edit harga, jadi 'harga' kiriman browser tak dipercaya — cegah manipulasi
        # payload (jual di bawah harga, sekaligus tak bisa dipakai membobol kontrol diskon PIN).
        harga = float(p['harga_jual'] or 0)
        # Multi-harga: terima harga pilihan kasir hanya bila termasuk level yang dikonfigurasi
        # untuk produk ini. Level "Utama" = harga_jual (sudah jadi default di atas).
        if pos_multi_harga:
            try:
                _want = round(float(raw.get('harga') or 0), 2)
            except (TypeError, ValueError):
                _want = 0
            if _want > 0 and _want != round(harga, 2) and _want in _allowed_harga.get(pid, set()):
                harga = _want
        if harga < 0 or not math.isfinite(harga):
            conn.close()
            flash(f'Harga item {idx} tidak valid.', 'danger')
            return redirect(url_for('pos_kasir'))
        items.append({
            'produk_id': pid,
            'deskripsi': p['nama'],
            'qty': qty,
            'satuan': p['satuan'],
            'harga': harga,
            'diskon': 0,
        })
    if not items:
        conn.close()
        flash('Keranjang POS masih kosong.', 'warning')
        return redirect(url_for('pos_kasir'))

    # ── Guard stok minus (opsional, toggle 'pos_block_minus' di POS Master) ──
    # Default OFF: jual dulu, barang menyusul (stok boleh minus, hanya diperingatkan setelah simpan).
    # ON: tolak checkout bila stok produk SKU / bahan SKU resep tidak mencukupi. Cek dilakukan
    # SEBELUM jurnal dibuat / stok dipotong sehingga tidak ada mutasi yang perlu di-rollback.
    if get_setting(conn, 'pos_block_minus', '0') == '1':
        try:
            _short = _pos_stock_shortage(conn, items, recipe_lines)
        except ValueError as ex:
            conn.close(); flash(str(ex), 'danger'); return redirect(url_for('pos_kasir'))
        if _short:
            conn.close()
            flash('Checkout ditolak karena stok tidak mencukupi (mode blokir stok minus aktif): '
                  + '; '.join(_short) + '. Tambah stok lewat menu Pembelian/Produksi, atau matikan '
                  'mode blokir di POS Master.', 'danger')
            return redirect(url_for('pos_kasir'))

    subtotal = sum(float(i['qty']) * float(i['harga']) for i in items)
    tanggal = date.today().strftime('%Y-%m-%d')

    # ── Voucher: sumber diskon TEROTORISASI (kode = izinnya, bypass PIN lock) ──
    # Server re-validasi & re-hitung potongan (jangan percaya nilai dari klien).
    customer = request.form.get('customer', '').strip()
    voucher_row = None
    voucher_alloc = {}
    voucher_kode = (request.form.get('voucher_kode', '') or '').strip().upper()
    if voucher_kode:
        ev = _evaluate_voucher(conn, voucher_kode, subtotal, tanggal, items=items, customer=customer)
        if not ev['ok']:
            conn.close()
            flash(f'Voucher {voucher_kode}: {ev["error"]}', 'danger')
            return redirect(url_for('pos_kasir'))
        voucher_row = ev['voucher']
        voucher_alloc = ev.get('alloc', {})
        diskon = ev['potongan']          # voucher menentukan diskon (override input)

    # ── Lock diskon: HANYA untuk diskon manual (voucher sudah terotorisasi) ──
    # Enforce server-side (UI bisa dibypass via devtools).
    if pos_s['pos_diskon_lock'] and diskon > 0 and voucher_row is None:
        pin_hash = (get_setting(conn, 'pos_diskon_pin', '') or '').strip()
        pin_in = (request.form.get('diskon_pin', '') or '').strip()
        if not pin_hash:
            conn.close()
            flash('Diskon nota dikunci. Hubungi admin untuk mengatur PIN diskon.', 'danger')
            return redirect(url_for('pos_kasir'))
        if not pin_in or sha256(pin_in.encode()).hexdigest() != pin_hash:
            conn.close()
            flash('PIN diskon salah atau belum dibuka. Diskon nota dibatalkan.', 'danger')
            return redirect(url_for('pos_kasir'))

    if diskon > subtotal:
        conn.close()
        flash('Diskon nota tidak boleh melebihi subtotal.', 'danger')
        return redirect(url_for('pos_kasir'))
    total = subtotal - diskon
    if total <= 0:
        conn.close()
        flash('Total checkout harus lebih dari 0.', 'danger')
        return redirect(url_for('pos_kasir'))
    if metode == 'PIUTANG':
        uang_masuk = 0
        bayar = 0
        kembalian = 0
    else:
        if metode == 'TUNAI' and bayar + 0.01 < total:
            conn.close()
            flash('Nominal bayar tunai kurang dari total.', 'danger')
            return redirect(url_for('pos_kasir'))
        if metode == 'REKENING':
            bayar = total
        uang_masuk = total
        kembalian = max(0, bayar - total)

    customer = request.form.get('customer', '').strip()
    nomor = _next_pos_receipt(conn, tanggal)
    keterangan = f"POS {nomor}" + (f" - {customer}" if customer else "")
    jid = conn.execute(
        "INSERT INTO jurnal(tanggal,keterangan,kategori,tipe_tx,pihak,nomor_tx) VALUES(?,?,?,?,?,?)",
        (tanggal, keterangan, 'OPERASIONAL', 'PEMASUKAN', customer, next_nomor_tx(conn, tanggal))
    ).lastrowid
    try:
        grand, _total_hpp = _build_sale(conn, jid, tanggal, keterangan, items, {
            'akun_kas': akun_kas,
            'uang_masuk': uang_masuk,
            'diskon': diskon,
            'ongkir': 0,
            'biaya': 0,
            'hpp_generik': 0,
            'pihak': customer,
        })
        # ── Baris resep (Model B): backflush bahan + COGS DI ATAS penjualan _build_sale ──
        # _build_sale sudah membalans pendapatan + kas (resep = baris non-SKU -> Cr 4100).
        # Di sini tambahkan HPP resep: Dr 5100 / Cr 1130 (bahan SKU+non-SKU) + Cr beban
        # (biaya tambahan). Blok ini self-balanced (Dr == Cr) sehingga jurnal tetap seimbang.
        resep_bahan_pids = []
        if recipe_lines:
            cogs_sku_t = cogs_nonsku_t = 0.0
            reclass_t = {}
            recs_all = []
            for rl in recipe_lines:
                cons = _recipe_consumption(conn, rl['template'], rl['qty'], rl['line_idx'])
                cogs_sku_t += cons['cogs_sku']
                cogs_nonsku_t += cons['cogs_nonsku']
                for _ak, _nv in cons['reclass'].items():
                    _add_amount(reclass_t, _ak, _nv)
                recs_all.extend(cons['recs'])
                resep_bahan_pids.extend(cons['bahan_pids'])
            inventory_cogs = cogs_sku_t + cogs_nonsku_t
            total_cogs = inventory_cogs + sum(reclass_t.values())
            if total_cogs > 0:
                _je(conn, jid, '5100', total_cogs, 0)
                if inventory_cogs > 0:
                    _je(conn, jid, '1130', 0, inventory_cogs)
                for _ak, _nv in sorted(reclass_t.items()):
                    _je(conn, jid, _ak, 0, _nv)
            _apply_recipe_backflush(conn, jid, tanggal, keterangan, recs_all)
            if cogs_nonsku_t > 0:
                record_non_sku_movement(conn, jid, tanggal, f"HPP non-SKU resep: {keterangan}", -cogs_nonsku_t)
            _assert_journal_balanced(conn, jid)
    except ValueError as ex:
        conn.rollback(); conn.close()
        flash(str(ex), 'danger')
        return redirect(url_for('pos_kasir'))
    try:
        conn.execute(
            """INSERT INTO pos_sale(jurnal_id,shift_id,nomor_struk,tanggal,customer,metode_bayar,akun_kas,total,bayar,kembalian,idempotency_key)
               VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
            (jid, shift['id'], nomor, tanggal, customer, metode, akun_kas, grand, bayar, kembalian, idem_key or None)
        )
    except sqlite3.IntegrityError:
        # Nabrak UNIQUE (nomor_struk atau idempotency_key) -> checkout kembar yang lolos
        # bareng dari pengecekan awal (race murni, dua request nyaris bersamaan). Batalkan
        # SELURUH mutasi request ini (jurnal/stok/backflush belum di-commit sama sekali).
        conn.rollback()
        _dup2 = None
        if idem_key:
            _dup2 = conn.execute(
                "SELECT jurnal_id FROM pos_sale WHERE idempotency_key=?", (idem_key,)
            ).fetchone()
        conn.close()
        if _dup2:
            return redirect(url_for('pos_struk', id=_dup2['jurnal_id']))
        flash('Checkout gagal karena bentrok nomor struk dengan transaksi lain di waktu hampir '
              'bersamaan. Silakan coba checkout ulang.', 'danger')
        return redirect(url_for('pos_kasir'))
    # Catat pemakaian voucher (increment kuota terpakai + log untuk audit).
    # UPDATE bersyarat + cek rowcount supaya atomik: kalau kuota baru saja habis dipakai
    # transaksi lain di antara pengecekan awal (_evaluate_voucher) & titik ini -- dua
    # checkout voucher kuota-terbatas yang nyaris bersamaan -- batalkan seluruh transaksi
    # ini alih-alih diam-diam melebihi kuota yang sudah ditetapkan.
    if voucher_row is not None:
        # Cek ULANG kuota DAN status aktif LANGSUNG dari kolom tabel (bukan snapshot
        # voucher_row yang direkam _evaluate_voucher di awal fungsi) di titik commit ini.
        # Kalau admin menonaktifkan / mengetatkan kuota voucher ini SAAT checkout sedang
        # diproses (jendela antara evaluasi awal & titik ini bisa berarti, bukan instan),
        # perubahan itu ikut dihormati -- bukan diabaikan karena kita masih pegang nilai lama.
        _vcur = conn.execute(
            "UPDATE voucher SET terpakai = COALESCE(terpakai,0) + 1 "
            "WHERE id=? AND COALESCE(aktif,0)=1 "
            "AND (COALESCE(kuota,0) <= 0 OR COALESCE(terpakai,0) < COALESCE(kuota,0))",
            (voucher_row['id'],)
        )
        if _vcur.rowcount == 0:
            conn.rollback()
            conn.close()
            flash(f'Voucher {voucher_row["kode"]}: sudah tidak berlaku (dinonaktifkan atau kuota '
                  'baru saja habis) saat transaksi diproses. Checkout dibatalkan, silakan ulangi '
                  'tanpa voucher atau pakai voucher lain.', 'danger')
            return redirect(url_for('pos_kasir'))
        conn.execute(
            "INSERT INTO voucher_pakai(voucher_id, jurnal_id, tanggal, potongan) VALUES(?,?,?,?)",
            (voucher_row['id'], jid, tanggal, diskon)
        )
    conn.execute("UPDATE jurnal SET tx_meta=? WHERE id=?", (json.dumps({
        'tipe': 'JUAL', 'sumber': 'POS', 'pos_nomor': nomor, 'pos_shift_id': shift['id'],
        'metode_bayar': metode, 'akun_kas': akun_kas, 'uang_masuk': uang_masuk,
        'diskon': diskon, 'ongkir': 0, 'biaya': 0, 'hpp_generik': 0,
        'pihak': customer, 'has_resep': bool(recipe_lines),
        'voucher_kode': voucher_row['kode'] if voucher_row is not None else '',
        'item_disc': voucher_alloc,
    }), jid))
    add_log(conn, 'Checkout POS', f"{nomor} | {grand}", 'POS')
    neg_notice = negative_stock_notice(conn, [i.get('produk_id') for i in items] + resep_bahan_pids)
    conn.commit(); conn.close()
    flash(f'Checkout berhasil. Struk {nomor} tersimpan.', 'success')
    if neg_notice:
        flash(neg_notice, 'warning')
    return redirect(url_for('pos_struk', id=jid))

@app.route('/pos/produk/<int:id>/favorit', methods=['POST'])
@pos_required
def pos_toggle_favorit(id):
    """Toggle status favorit produk untuk akses cepat di kasir (AJAX)."""
    conn = db()
    p = conn.execute(
        "SELECT id, COALESCE(favorit,0) AS favorit FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
        (id,)
    ).fetchone()
    if not p:
        conn.close()
        return jsonify({'ok': False, 'error': 'Produk tidak ditemukan.'}), 404
    baru = 0 if p['favorit'] else 1
    conn.execute("UPDATE produk SET favorit=? WHERE id=?", (baru, id))
    conn.commit()
    total_fav = conn.execute(
        "SELECT COUNT(*) c FROM produk WHERE COALESCE(favorit,0)=1 AND COALESCE(aktif,1)=1"
    ).fetchone()['c']
    conn.close()
    return jsonify({'ok': True, 'favorit': baru, 'total_favorit': total_fav})

@app.route('/pos/voucher/cek', methods=['POST'])
@pos_required
def pos_voucher_cek():
    """Validasi voucher + hitung potongan untuk subtotal saat ini (AJAX, dipakai kasir)."""
    conn = db()
    try:
        subtotal = parse_nonnegative_rp(request.form.get('subtotal', '0'), 'Subtotal')
    except ValueError:
        subtotal = 0.0
    try:
        items = json.loads(request.form.get('items', '[]') or '[]')
        if not isinstance(items, list):
            items = []
    except (json.JSONDecodeError, TypeError):
        items = []
    customer = (request.form.get('customer', '') or '').strip()
    ev = _evaluate_voucher(conn, request.form.get('kode', ''), subtotal,
                           date.today().strftime('%Y-%m-%d'), items=items, customer=customer)
    conn.close()
    if not ev['ok']:
        return jsonify({'ok': False, 'error': ev['error']})
    return jsonify({'ok': True, 'potongan': ev['potongan'], 'label': ev['label'],
                    'kode': ev['voucher']['kode'], 'alloc': ev.get('alloc', {})})

@app.route('/pos/voucher', methods=['GET', 'POST'])
@master_pos_required
def pos_voucher():
    """Kelola voucher diskon POS (manajer/admin)."""
    conn = db()
    if request.method == 'POST':
        action = request.form.get('action', 'tambah')
        if action in ('tambah', 'edit'):
            kode = (request.form.get('kode') or '').strip().upper()
            tipe = (request.form.get('tipe') or 'NOMINAL').strip().upper()
            if tipe not in ('NOMINAL', 'PERSEN'):
                tipe = 'NOMINAL'
            catatan = (request.form.get('catatan') or '').strip()
            berlaku_dari = (request.form.get('berlaku_dari') or '').strip() or None
            berlaku_sampai = (request.form.get('berlaku_sampai') or '').strip() or None
            if not kode:
                conn.close(); flash('Kode voucher wajib diisi.', 'danger')
                return redirect(url_for('pos_voucher'))
            if not is_valid_optional_iso_date(berlaku_dari) or not is_valid_optional_iso_date(berlaku_sampai):
                conn.close(); flash('Tanggal berlaku voucher tidak valid.', 'danger')
                return redirect(url_for('pos_voucher'))
            try:
                nilai = parse_nonnegative_rp(request.form.get('nilai', '0'), 'Nilai voucher')
                min_belanja = parse_nonnegative_rp(request.form.get('min_belanja', '0'), 'Minimal belanja')
                maks_potongan = parse_nonnegative_rp(request.form.get('maks_potongan', '0'), 'Maks potongan')
                kuota = int(parse_qty(request.form.get('kuota', '0')))
            except (ValueError, TypeError) as ex:
                conn.close(); flash(str(ex), 'danger')
                return redirect(url_for('pos_voucher'))
            if tipe == 'PERSEN' and nilai > 100:
                conn.close(); flash('Voucher persen tidak boleh lebih dari 100%.', 'danger')
                return redirect(url_for('pos_voucher'))
            if nilai <= 0:
                conn.close(); flash('Nilai voucher harus lebih dari 0.', 'danger')
                return redirect(url_for('pos_voucher'))
            # Basis diskon (NOTA/KATEGORI/PRODUK) + ambang qty + pembatasan customer
            scope = (request.form.get('scope') or 'NOTA').strip().upper()
            if scope not in ('NOTA', 'PRODUK', 'KATEGORI'):
                scope = 'NOTA'
            # Cara potong: TOTAL (sekali) / PER_PRODUK (nilai per unit). Hanya utk subset produk/kategori.
            potongan_mode = (request.form.get('potongan_mode') or 'TOTAL').strip().upper()
            if potongan_mode not in ('TOTAL', 'PER_PRODUK') or scope == 'NOTA':
                potongan_mode = 'TOTAL'
            # Multi-kategori: checkbox kirim banyak nilai 'syarat_kategori'. Dedup + strip.
            _seen_kat = set(); kategori_list = []
            for _kt in request.form.getlist('syarat_kategori'):
                _kt = (_kt or '').strip()
                if _kt and _kt.lower() not in _seen_kat:
                    _seen_kat.add(_kt.lower()); kategori_list.append(_kt)
            syarat_kategori = ', '.join(kategori_list)   # kolom legacy = ringkas (badge/fallback)
            try:
                min_qty = float(parse_qty(request.form.get('min_qty', '0')))
            except (ValueError, TypeError):
                min_qty = 0.0
            if min_qty < 0:
                min_qty = 0.0
            produk_ids = []
            for _pid in request.form.getlist('produk_id'):
                try:
                    produk_ids.append(int(_pid))
                except (TypeError, ValueError):
                    pass
            produk_ids = sorted(set(produk_ids))
            # Saring ke produk yang masih ada (cegah FOREIGN KEY error 500 saat
            # form basi / produk dihapus setelah halaman terbuka).
            if produk_ids:
                _ph = ','.join('?' * len(produk_ids))
                _exist = {r['id'] for r in conn.execute(
                    f"SELECT id FROM produk WHERE id IN ({_ph})", produk_ids).fetchall()}
                produk_ids = [p for p in produk_ids if p in _exist]
            _seen_cust = set(); customer_list = []
            for _cn in request.form.getlist('customer_nama'):
                _cn = (_cn or '').strip()
                if _cn and _cn.lower() not in _seen_cust:
                    _seen_cust.add(_cn.lower()); customer_list.append(_cn)
            if scope == 'PRODUK' and not produk_ids:
                conn.close(); flash('Voucher per-produk wajib memilih minimal 1 produk.', 'danger')
                return redirect(url_for('pos_voucher'))
            if scope == 'KATEGORI' and not kategori_list:
                conn.close(); flash('Voucher per-kategori wajib memilih minimal 1 kategori.', 'danger')
                return redirect(url_for('pos_voucher'))
            if scope != 'KATEGORI':
                syarat_kategori = ''; kategori_list = []
            if action == 'tambah':
                try:
                    vid = conn.execute(
                        "INSERT INTO voucher(kode,tipe,nilai,min_belanja,maks_potongan,kuota,"
                        "berlaku_dari,berlaku_sampai,catatan,scope,syarat_kategori,min_qty,potongan_mode) "
                        "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (kode, tipe, nilai, min_belanja, maks_potongan, kuota,
                         berlaku_dari, berlaku_sampai, catatan, scope, syarat_kategori, min_qty, potongan_mode)).lastrowid
                except sqlite3.IntegrityError:
                    conn.close(); flash(f'Kode voucher {kode} sudah ada.', 'danger')
                    return redirect(url_for('pos_voucher'))
                _save_voucher_scope(conn, vid, scope, produk_ids, customer_list, kategori_list)
                add_log(conn, 'Voucher dibuat', f'{kode} {tipe} {nilai} {scope}', 'POS')
                flash('Voucher baru dibuat!', 'success')
            else:
                vid = request.form.get('id', type=int)
                if not conn.execute("SELECT id FROM voucher WHERE id=?", (vid,)).fetchone():
                    conn.close(); flash('Voucher tidak ditemukan.', 'danger')
                    return redirect(url_for('pos_voucher'))
                try:
                    conn.execute(
                        "UPDATE voucher SET kode=?,tipe=?,nilai=?,min_belanja=?,maks_potongan=?,kuota=?,"
                        "berlaku_dari=?,berlaku_sampai=?,catatan=?,scope=?,syarat_kategori=?,min_qty=?,potongan_mode=? WHERE id=?",
                        (kode, tipe, nilai, min_belanja, maks_potongan, kuota,
                         berlaku_dari, berlaku_sampai, catatan, scope, syarat_kategori, min_qty, potongan_mode, vid))
                except sqlite3.IntegrityError:
                    conn.close(); flash(f'Kode voucher {kode} sudah dipakai.', 'danger')
                    return redirect(url_for('pos_voucher'))
                _save_voucher_scope(conn, vid, scope, produk_ids, customer_list, kategori_list)
                flash('Voucher diperbarui!', 'success')
        elif action == 'toggle':
            vid = request.form.get('id', type=int)
            v = conn.execute("SELECT aktif FROM voucher WHERE id=?", (vid,)).fetchone()
            if v:
                conn.execute("UPDATE voucher SET aktif=? WHERE id=?", (0 if v['aktif'] else 1, vid))
                flash('Status voucher diubah.', 'success')
        elif action == 'hapus':
            vid = request.form.get('id', type=int)
            n_pakai = conn.execute("SELECT COUNT(*) c FROM voucher_pakai WHERE voucher_id=?", (vid,)).fetchone()['c']
            if n_pakai:
                conn.execute("UPDATE voucher SET aktif=0 WHERE id=?", (vid,))
                flash(f'Voucher sudah pernah dipakai ({n_pakai}x) — dinonaktifkan saja (riwayat dijaga).', 'info')
            else:
                conn.execute("DELETE FROM voucher WHERE id=?", (vid,))
                flash('Voucher dihapus.', 'success')
        conn.commit(); conn.close()
        return redirect(url_for('pos_voucher'))

    page = _page_arg()
    n_rows = conn.execute("SELECT COUNT(*) FROM voucher").fetchone()[0]
    pg = paginate_meta(page, n_rows)
    rows = conn.execute(
        "SELECT * FROM voucher ORDER BY aktif DESC, dibuat DESC LIMIT ? OFFSET ?",
        (pg['per_page'], pg['offset'])
    ).fetchall()
    # Data untuk picker scope (produk) + pembatasan customer di form voucher.
    products = [dict(r) for r in conn.execute(
        "SELECT id, kode, nama, varian FROM produk WHERE COALESCE(aktif,1)=1 "
        "ORDER BY nama COLLATE NOCASE, varian COLLATE NOCASE"
    ).fetchall()]
    customer_names = [r['nama'] for r in conn.execute(
        "SELECT nama FROM customer ORDER BY nama COLLATE NOCASE"
    ).fetchall()]
    kategori_produk = [row['kategori'] for row in conn.execute(
        "SELECT DISTINCT TRIM(kategori) AS kategori FROM produk "
        "WHERE COALESCE(aktif,1)=1 AND TRIM(COALESCE(kategori,''))<>'' ORDER BY kategori COLLATE NOCASE"
    ).fetchall()]
    vp_map = {}
    for r in conn.execute("SELECT voucher_id, produk_id FROM voucher_produk").fetchall():
        vp_map.setdefault(r['voucher_id'], []).append(r['produk_id'])
    vc_map = {}
    for r in conn.execute("SELECT voucher_id, customer_nama FROM voucher_customer").fetchall():
        vc_map.setdefault(r['voucher_id'], []).append(r['customer_nama'])
    vk_map = {}
    for r in conn.execute("SELECT voucher_id, kategori FROM voucher_kategori").fetchall():
        vk_map.setdefault(r['voucher_id'], []).append(r['kategori'])
    conn.close()
    return render_template('voucher_list.html', voucher_rows=rows,
                           today=date.today().strftime('%Y-%m-%d'), pg=pg,
                           products=products, customer_names=customer_names,
                           kategori_produk=kategori_produk,
                           vp_map=vp_map, vc_map=vc_map, vk_map=vk_map)

@app.route('/pos/struk/<int:id>')
@pos_required
def pos_struk(id):
    conn = db()
    sale = conn.execute(
        "SELECT ps.*,j.keterangan,j.nomor_tx,u.nama AS kasir_nama FROM pos_sale ps "
        "JOIN jurnal j ON j.id=ps.jurnal_id "
        "LEFT JOIN pos_shift sh ON sh.id=ps.shift_id "
        "LEFT JOIN users u ON u.id=sh.user_id "
        "WHERE ps.jurnal_id=?",
        (id,)
    ).fetchone()
    if not sale:
        conn.close()
        flash('Struk POS tidak ditemukan.', 'danger')
        return redirect(url_for('pos_kasir'))
    items = _enrich_item_varian(conn, conn.execute(
        "SELECT * FROM transaksi_item WHERE jurnal_id=? ORDER BY id",
        (id,)
    ).fetchall())
    # HARGA PER ITEM di struk = harga SETELAH diskon. Sumber alokasi:
    #  (1) voucher -> pakai alokasi tersimpan di tx_meta.item_disc (AKURAT per scope: produk
    #      di luar kategori/scope voucher TIDAK ikut terdiskon);
    #  (2) diskon manual nota -> sebar proporsional ke semua item.
    # Blok ringkasan Subtotal/Diskon/Total tetap pakai subtotal kotor (it.subtotal) -> angka tak berubah.
    _item_disc = {}
    try:
        _mrow = conn.execute("SELECT tx_meta FROM jurnal WHERE id=?", (id,)).fetchone()
        if _mrow and _mrow['tx_meta']:
            _item_disc = (json.loads(_mrow['tx_meta']) or {}).get('item_disc') or {}
    except Exception:
        _item_disc = {}
    _gross = sum(float(it.get('subtotal') or 0) for it in items)
    _total = float(sale['total'] or 0)
    if _item_disc:
        for it in items:
            g = float(it.get('subtotal') or 0)
            d = float(_item_disc.get(str(it.get('produk_id')), 0) or 0)
            net = max(0.0, round(g - d, 2))
            q = float(it.get('qty') or 0)
            it['subtotal_net'] = net
            it['harga_net'] = round(net / q, 2) if q else net
    elif _gross > _total + 0.01 and _total > 0:
        _acc = 0.0
        for i, it in enumerate(items):
            g = float(it.get('subtotal') or 0)
            if i == len(items) - 1:
                net = round(_total - _acc, 2)            # item terakhir = sisa (rapikan pembulatan)
            else:
                net = round(g * _total / _gross, 2)
                _acc = round(_acc + net, 2)
            net = max(0.0, net)
            q = float(it.get('qty') or 0)
            it['subtotal_net'] = net
            it['harga_net'] = round(net / q, 2) if q else net
    else:
        for it in items:
            it['subtotal_net'] = float(it.get('subtotal') or 0)
            it['harga_net'] = float(it.get('harga') or 0)
    pos_s = _pos_receipt_settings(conn)
    conn.close()
    return render_template('pos_struk.html', sale=sale, items=items, pos_s=pos_s)

# ═══════════════════════════════════════════════════════════════════════════
#   MASTER POS — Dashboard pengawasan utk Manager/Finance/Admin
#   (bukan utk kasir input transaksi — itu di /pos)
# ═══════════════════════════════════════════════════════════════════════════
@app.route('/pos/master')
@master_pos_required
def pos_master():
    conn = db()
    today    = date.today()
    today_s  = today.strftime('%Y-%m-%d')
    week_ago = (today - timedelta(days=6)).strftime('%Y-%m-%d')
    period   = request.args.get('period', 'today')  # today | week | month | day | range
    if period not in ('today', 'week', 'month', 'day', 'range'):
        period = 'today'

    # Pilih hari (single-day mode)
    sel_raw  = request.args.get('tanggal', today_s)
    try:
        _sel = date.fromisoformat(sel_raw)
        if _sel > today: _sel = today
        selected_date_s = _sel.strftime('%Y-%m-%d')
    except ValueError:
        _sel = today
        selected_date_s = today_s
    prev_date_s = (_sel - timedelta(days=1)).strftime('%Y-%m-%d')
    next_date_s = min(_sel + timedelta(days=1), today).strftime('%Y-%m-%d')

    # Rentang kustom (range mode)
    try:
        _rsd = date.fromisoformat(request.args.get('sd', today_s))
        _red = date.fromisoformat(request.args.get('ed', today_s))
        if _rsd > today: _rsd = today
        if _red > today: _red = today
        if _rsd > _red: _rsd, _red = _red, _rsd
    except ValueError:
        _rsd = _red = today
    range_sd_s = _rsd.strftime('%Y-%m-%d')
    range_ed_s = _red.strftime('%Y-%m-%d')

    if period == 'week':
        date_from, date_to = week_ago, today_s
    elif period == 'month':
        date_from, date_to = today.replace(day=1).strftime('%Y-%m-%d'), today_s
    elif period == 'day':
        date_from = date_to = selected_date_s
    elif period == 'range':
        date_from, date_to = range_sd_s, range_ed_s
    else:
        date_from = date_to = today_s

    # ── KPI utama ─────────────────────────────────────────────────────────
    kpi_day_date = selected_date_s if period == 'day' else today_s
    kpi_today = conn.execute(
        "SELECT COUNT(*) c, COALESCE(SUM(total),0) v FROM pos_sale WHERE tanggal=?",
        (kpi_day_date,)
    ).fetchone()
    kpi_period = conn.execute(
        "SELECT COUNT(*) c, COALESCE(SUM(total),0) v FROM pos_sale "
        "WHERE tanggal BETWEEN ? AND ?", (date_from, date_to)
    ).fetchone()
    avg_trx = float(kpi_period['v'] or 0) / kpi_period['c'] if kpi_period['c'] else 0
    open_shifts_cnt = conn.execute(
        "SELECT COUNT(*) c FROM pos_shift WHERE status='OPEN'"
    ).fetchone()['c']

    # ── Orders per hour (hari ini atau hari terpilih) ────────────────────
    hourly_date = selected_date_s if period == 'day' else today_s
    rows_hour = conn.execute(
        "SELECT CAST(strftime('%H', dibuat, 'localtime') AS INTEGER) AS jam, "
        "       COUNT(*) AS c, COALESCE(SUM(total),0) AS v "
        "FROM pos_sale WHERE tanggal=? GROUP BY jam ORDER BY jam",
        (hourly_date,)
    ).fetchall()
    hourly = {int(r['jam']): {'c': r['c'], 'v': float(r['v'])} for r in rows_hour}
    hourly_labels = list(range(7, 23))  # 07:00 – 22:00 (window operasional umum)
    hourly_count = [hourly.get(h, {'c': 0})['c'] for h in hourly_labels]
    hourly_value = [hourly.get(h, {'v': 0})['v'] for h in hourly_labels]

    # ── Top produk (day + range), daftar transaksi (day only) ──────────────
    if period in ('day', 'range'):
        rows_produk = conn.execute(
            "SELECT pr.nama, pr.satuan, COALESCE(SUM(ABS(ps.qty)),0) AS qty, "
            "       COALESCE(SUM(ABS(ps.qty) * pr.harga_jual),0) AS est_revenue "
            "FROM pergerakan_stok ps "
            "JOIN produk pr ON pr.id = ps.produk_id "
            "JOIN pos_sale psa ON psa.jurnal_id = ps.jurnal_id "
            "WHERE ps.tanggal BETWEEN ? AND ? AND ps.jenis='KELUAR' "
            "GROUP BY ps.produk_id ORDER BY qty DESC LIMIT 10",
            (date_from, date_to)
        ).fetchall()
        top_produk_day = [dict(r) for r in rows_produk]
    else:
        top_produk_day = []

    # Daftar transaksi: ikut periode (semua mode) — biar konsisten dgn KPI
    rows_txn = conn.execute(
        "SELECT ps.nomor_struk, ps.customer, ps.metode_bayar, ps.total, ps.dibuat, ps.tanggal, "
        "       COALESCE(NULLIF(u.nama,''), psh.username, '-') AS kasir_nama "
        "FROM pos_sale ps "
        "LEFT JOIN pos_shift psh ON psh.id = ps.shift_id "
        "LEFT JOIN users u ON u.id = psh.user_id "
        "WHERE ps.tanggal BETWEEN ? AND ? ORDER BY ps.dibuat DESC LIMIT 200",
        (date_from, date_to)
    ).fetchall()
    daily_transactions = [dict(r) for r in rows_txn]

    # Tanggal terlama di POS (untuk min date di range picker)
    _ep = conn.execute("SELECT MIN(tanggal) FROM pos_sale").fetchone()
    earliest_pos_s = _ep[0] if _ep and _ep[0] else today_s

    # ── Top kasir (periode) ───────────────────────────────────────────────
    rows_kasir = conn.execute(
        "SELECT COALESCE(NULLIF(u.nama,''), psh.username) AS kasir, "
        "       COUNT(ps.id) AS jml, COALESCE(SUM(ps.total),0) AS omzet "
        "FROM pos_sale ps "
        "JOIN pos_shift psh ON psh.id = ps.shift_id "
        "LEFT JOIN users u  ON u.id = psh.user_id "
        "WHERE ps.tanggal BETWEEN ? AND ? "
        "GROUP BY kasir ORDER BY omzet DESC LIMIT 8",
        (date_from, date_to)
    ).fetchall()
    top_kasir = [dict(r) for r in rows_kasir]

    # ── Payment-method split (periode) ────────────────────────────────────
    rows_pay = conn.execute(
        "SELECT metode_bayar, COUNT(*) AS c, COALESCE(SUM(total),0) AS v "
        "FROM pos_sale WHERE tanggal BETWEEN ? AND ? "
        "GROUP BY metode_bayar", (date_from, date_to)
    ).fetchall()
    pay_split = {r['metode_bayar']: {'c': r['c'], 'v': float(r['v'])} for r in rows_pay}

    # ── Trend omzet: SELALU mengikuti periode yang dipilih (intuitif) ──
    # Tidak ada lagi "context window" yang membingungkan. Apa yang dipilih
    # user di DRP itulah yang ditampilkan grafiknya, persis sama dengan KPI.
    _tsd = date.fromisoformat(date_from)
    _ted = date.fromisoformat(date_to)
    trend_from_s = date_from
    trend_to_s   = date_to

    rows_trend = conn.execute(
        "SELECT tanggal, COUNT(*) AS c, COALESCE(SUM(total),0) AS v "
        "FROM pos_sale WHERE tanggal BETWEEN ? AND ? "
        "GROUP BY tanggal ORDER BY tanggal",
        (trend_from_s, trend_to_s)
    ).fetchall()
    trend_map = {r['tanggal']: {'c': r['c'], 'v': float(r['v'])} for r in rows_trend}

    trend_labels, trend_count, trend_value = [], [], []
    span_days = (_ted - _tsd).days + 1
    HARI_NAMA = ['Sen','Sel','Rab','Kam','Jum','Sab','Min']  # weekday(): 0=Senin

    if span_days <= 62:
        # Mode harian (per tanggal)
        _d = _tsd
        while _d <= _ted:
            ds = _d.strftime('%Y-%m-%d')
            trend_labels.append(f"{HARI_NAMA[_d.weekday()]} {_d.day}/{_d.month}")
            trend_count.append(trend_map.get(ds, {'c': 0})['c'])
            trend_value.append(trend_map.get(ds, {'v': 0})['v'])
            _d += timedelta(days=1)
    else:
        # Mode mingguan (agregasi 7 hari) supaya chart tetap terbaca
        _d = _tsd
        while _d <= _ted:
            _we = min(_d + timedelta(days=6), _ted)
            wk_c = wk_v = 0
            _wd = _d
            while _wd <= _we:
                wd_s = _wd.strftime('%Y-%m-%d')
                bucket = trend_map.get(wd_s)
                if bucket:
                    wk_c += bucket['c']
                    wk_v += bucket['v']
                _wd += timedelta(days=1)
            trend_labels.append(f"{_d.day}/{_d.month}–{_we.day}/{_we.month}")
            trend_count.append(wk_c)
            trend_value.append(wk_v)
            _d = _we + timedelta(days=1)

    trend_from_disp = trend_from_s
    trend_to_disp   = trend_to_s
    trend_is_weekly = span_days > 62

    # ── Ringkasan shift (daftar lengkap pindah ke /pos/master/shift) ──────
    # Daftar penuh dulu ditampilkan inline di halaman ini; di kedai yang ramai
    # tabelnya jadi sangat panjang dan menutupi panel setelan di bawahnya.
    shifts_date = selected_date_s if period == 'day' else today_s
    shift_hari_ini = conn.execute(
        "SELECT COUNT(*) FROM pos_shift WHERE DATE(dibuka,'localtime')=?", (shifts_date,)
    ).fetchone()[0]
    shift_total = conn.execute("SELECT COUNT(*) FROM pos_shift").fetchone()[0]

    # ── Daftar kasir yang sudah terdaftar (utk panel "Formasi Kasir") ─────
    kasir_users = conn.execute(
        "SELECT id, username, nama, role, aktif FROM users "
        "WHERE role='KASIR' ORDER BY aktif DESC, username"
    ).fetchall()
    kasir_users = [dict(u) for u in kasir_users]

    # ── Setting tampilan kasir (diatur oleh manajer dari Master POS) ──────
    pos_menu_size = get_setting(conn, 'pos_menu_size', 'medium')
    if pos_menu_size not in ('small', 'medium', 'large'):
        pos_menu_size = 'medium'
    pos_icon_mode = get_setting(conn, 'pos_icon_mode', 'initial')
    if pos_icon_mode not in ('initial', 'code', 'custom'):
        pos_icon_mode = 'initial'
    pos_diskon_lock = get_setting(conn, 'pos_diskon_lock', '0') == '1'
    pos_diskon_has_pin = bool((get_setting(conn, 'pos_diskon_pin', '') or '').strip())
    pos_block_minus = get_setting(conn, 'pos_block_minus', '0') == '1'
    pos_show_mode = get_setting(conn, 'pos_show_mode', 'both')
    if pos_show_mode not in ('both', 'produk', 'resep'):
        pos_show_mode = 'both'
    # ── Kategori di kasir (sembunyikan sekaligus, bukan per produk) ──
    # Kategori yang SEDANG disembunyikan tetap ikut didaftar walau kebetulan sudah tak
    # punya item (mis. semua produknya diarsipkan sementara), supaya statusnya tidak
    # hilang diam-diam dan bisa ditampilkan lagi kapan pun.
    _hide_set = _pos_hidden_kategori(conn)
    kategori_list = _pos_kategori_terpakai(conn)
    _ada = {k['nama'].strip().lower() for k in kategori_list}
    try:
        _hide_raw = json.loads(get_setting(conn, 'pos_hide_kategori', '[]') or '[]')
        if not isinstance(_hide_raw, list):
            _hide_raw = []
    except (ValueError, TypeError):
        _hide_raw = []
    for _k in _hide_raw:
        _k = str(_k).strip()
        if _k and _k.lower() not in _ada:
            _ada.add(_k.lower())
            kategori_list.append({'nama': _k, 'n': 0})
    kategori_list.sort(key=lambda k: k['nama'].lower())
    for k in kategori_list:
        k['tampil'] = k['nama'].strip().lower() not in _hide_set
    n_kategori_hidden = sum(1 for k in kategori_list if not k['tampil'])

    conn.close()
    return render_template('pos_master.html',
        period=period, date_from=date_from, date_to=date_to,
        today_s=today_s, selected_date_s=selected_date_s,
        prev_date_s=prev_date_s, next_date_s=next_date_s,
        range_sd_s=range_sd_s, range_ed_s=range_ed_s,
        earliest_pos_s=earliest_pos_s,
        kpi_today=dict(kpi_today), kpi_period=dict(kpi_period),
        avg_trx=avg_trx, open_shifts_cnt=open_shifts_cnt,
        hourly_labels=hourly_labels, hourly_count=hourly_count, hourly_value=hourly_value,
        hourly_date=hourly_date,
        top_kasir=top_kasir, pay_split=pay_split,
        trend_labels=trend_labels, trend_count=trend_count, trend_value=trend_value,
        trend_from_disp=trend_from_disp, trend_to_disp=trend_to_disp, trend_is_weekly=trend_is_weekly,
        shift_hari_ini=shift_hari_ini, shift_total=shift_total, kasir_users=kasir_users,
        top_produk_day=top_produk_day, daily_transactions=daily_transactions,
        pos_menu_size=pos_menu_size, pos_icon_mode=pos_icon_mode,
        pos_diskon_lock=pos_diskon_lock, pos_diskon_has_pin=pos_diskon_has_pin,
        pos_block_minus=pos_block_minus,
        pos_show_mode=pos_show_mode,
        kategori_list=kategori_list, n_kategori_hidden=n_kategori_hidden,
    )

@app.route('/pos/master/shift')
@master_pos_required
def pos_master_shift():
    """Riwayat jadwal shift kasir, berpaginasi.

    Dipisah dari halaman POS Master karena daftarnya bisa panjang sekali
    (satu baris per buka-tutup kasir) sehingga menutupi panel setelan.
    Filter: status (semua/open/closed), kasir, dan rentang tanggal.
    """
    conn = db()
    status = (request.args.get('status', '') or '').strip().upper()
    kasir  = (request.args.get('kasir', '') or '').strip()
    dari   = (request.args.get('dari', '') or '').strip()
    sampai = (request.args.get('sampai', '') or '').strip()

    where, params = [], []
    if status in ('OPEN', 'CLOSED'):
        where.append("psh.status=?"); params.append(status)
    if kasir:
        where.append("psh.username=?"); params.append(kasir)
    if is_valid_iso_date(dari):
        where.append("DATE(psh.dibuka,'localtime')>=?"); params.append(dari)
    if is_valid_iso_date(sampai):
        where.append("DATE(psh.dibuka,'localtime')<=?"); params.append(sampai)
    w = (' WHERE ' + ' AND '.join(where)) if where else ''

    n_rows = conn.execute(f"SELECT COUNT(*) FROM pos_shift psh{w}", params).fetchone()[0]
    page = _page_arg()
    pg = paginate_meta(page, n_rows)
    shifts = [dict(r) for r in conn.execute(
        "SELECT psh.*, COALESCE(NULLIF(u.nama,''), psh.username) AS kasir_nama, "
        "       (SELECT COUNT(*) FROM pos_sale ps WHERE ps.shift_id=psh.id) AS jml_trx, "
        "       (SELECT COALESCE(SUM(total),0) FROM pos_sale ps WHERE ps.shift_id=psh.id) AS omzet "
        "FROM pos_shift psh LEFT JOIN users u ON u.id=psh.user_id"
        f"{w} ORDER BY psh.dibuka DESC, psh.id DESC LIMIT ? OFFSET ?",
        params + [pg['per_page'], pg['offset']]
    ).fetchall()]

    # Ringkasan dihitung dari SELURUH hasil filter, bukan cuma halaman ini.
    tot = conn.execute(
        "SELECT COUNT(*) n, "
        "  COALESCE(SUM((SELECT COUNT(*) FROM pos_sale ps WHERE ps.shift_id=psh.id)),0) trx, "
        "  COALESCE(SUM((SELECT COALESCE(SUM(total),0) FROM pos_sale ps WHERE ps.shift_id=psh.id)),0) omzet "
        f"FROM pos_shift psh{w}", params).fetchone()
    open_cnt = conn.execute("SELECT COUNT(*) FROM pos_shift WHERE status='OPEN'").fetchone()[0]
    kasir_list = [r['username'] for r in conn.execute(
        "SELECT DISTINCT username FROM pos_shift ORDER BY username").fetchall()]
    conn.close()
    return render_template('pos_master_shift.html', shifts=shifts, pg=pg,
                           status=status, kasir=kasir, dari=dari, sampai=sampai,
                           kasir_list=kasir_list, open_cnt=open_cnt,
                           tot_shift=tot['n'], tot_trx=tot['trx'], tot_omzet=tot['omzet'])

@app.route('/pos/master/hourly-compare')
@master_pos_required
def pos_master_hourly_compare():
    """Orderan per jam untuk 1–7 tanggal (perbandingan). Kembalikan JSON."""
    from flask import jsonify
    conn = db()
    today = date.today()
    raw = (request.args.get('dates', '') or '').strip()
    parsed, seen = [], set()
    for tok in raw.split(','):
        tok = tok.strip()
        if not tok:
            continue
        try:
            d = date.fromisoformat(tok)
        except ValueError:
            continue
        if d > today:
            d = today
        ds = d.strftime('%Y-%m-%d')
        if ds in seen:
            continue
        seen.add(ds)
        parsed.append(ds)
        if len(parsed) >= 7:
            break
    if not parsed:
        parsed = [today.strftime('%Y-%m-%d')]
    labels = list(range(7, 23))  # 07:00 – 22:00
    series = []
    for ds in parsed:
        rows = conn.execute(
            "SELECT CAST(strftime('%H', dibuat, 'localtime') AS INTEGER) AS jam, "
            "       COUNT(*) AS c, COALESCE(SUM(total),0) AS v "
            "FROM pos_sale WHERE tanggal=? GROUP BY jam ORDER BY jam",
            (ds,)
        ).fetchall()
        m = {int(r['jam']): {'c': r['c'], 'v': float(r['v'])} for r in rows}
        count = [m.get(h, {'c': 0})['c'] for h in labels]
        value = [m.get(h, {'v': 0})['v'] for h in labels]
        series.append({
            'date': ds, 'count': count, 'value': value,
            'total_c': sum(count), 'total_v': sum(value),
        })
    conn.close()
    return jsonify({'labels': labels, 'series': series})

@app.route('/pos/master/tambah-kasir', methods=['POST'])
@master_pos_required
def pos_master_tambah_kasir():
    """Tambah user KASIR-ONLY dari Master POS (manager bisa, tanpa harus jadi ADMIN)."""
    username = (request.form.get('username') or '').strip()
    nama     = (request.form.get('nama') or '').strip()
    password = (request.form.get('password') or '').strip()
    if not username or not password:
        flash('Username & password wajib diisi.', 'danger')
        return redirect(url_for('pos_master'))
    if len(password) < 4:
        flash('Password minimal 4 karakter.', 'danger')
        return redirect(url_for('pos_master'))
    conn = db()
    existing = conn.execute("SELECT id FROM users WHERE username=?", (username,)).fetchone()
    if existing:
        conn.close()
        flash(f'Username "{username}" sudah dipakai.', 'danger')
        return redirect(url_for('pos_master'))
    # Permission KASIR-only: tidak bisa eskalasi ke role lain dari sini
    perms_json = json.dumps(['pos_sales'])
    conn.execute(
        "INSERT INTO users(username, password_hash, role, nama, aktif, permissions) "
        "VALUES(?, ?, 'KASIR', ?, 1, ?)",
        (username, sha256(password.encode()).hexdigest(), nama or username, perms_json)
    )
    add_log(conn, 'Tambah Kasir (Master POS)',
            f'username={username} | role=KASIR', 'POS')
    conn.commit()
    conn.close()
    flash(f'Kasir "{username}" berhasil ditambahkan.', 'success')
    return redirect(url_for('pos_master'))

@app.route('/pos/master/tampilan', methods=['POST'])
@master_pos_required
def pos_master_tampilan():
    """Update setting tampilan kasir (ukuran kartu + mode ikon) dari Master POS.
    Bisa diatur oleh siapapun yang punya hak Master POS, tidak hanya ADMIN."""
    menu_size = (request.form.get('pos_menu_size', 'medium') or 'medium').strip()
    if menu_size not in ('small', 'medium', 'large'):
        menu_size = 'medium'
    icon_mode = (request.form.get('pos_icon_mode', 'initial') or 'initial').strip()
    if icon_mode not in ('initial', 'code', 'custom'):
        icon_mode = 'initial'
    conn = db()
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_menu_size', ?)", (menu_size,))
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_icon_mode', ?)", (icon_mode,))
    add_log(conn, 'Update Tampilan Kasir',
            f'ukuran={menu_size} | mode_ikon={icon_mode}', 'POS')
    conn.commit()
    conn.close()
    flash('Tampilan kasir berhasil disimpan.', 'success')
    return redirect(url_for('pos_master') + '#card-tampilan-kasir')

@app.route('/pos/master/diskon', methods=['POST'])
@master_pos_required
def pos_master_diskon():
    """Atur penguncian diskon nota POS + PIN supervisor untuk membuka.
    Diatur oleh pemegang hak Master POS."""
    conn = db()
    lock = '1' if request.form.get('pos_diskon_lock') else '0'
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_diskon_lock', ?)", (lock,))
    # PIN: hanya update kalau diisi. Field kosong = biarkan PIN lama.
    pin_baru = (request.form.get('pos_diskon_pin', '') or '').strip()
    hapus_pin = request.form.get('pos_diskon_pin_hapus')
    if hapus_pin:
        conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_diskon_pin', '')")
        pin_note = 'PIN dihapus'
    elif pin_baru:
        if not pin_baru.isdigit() or not (4 <= len(pin_baru) <= 8):
            conn.close()
            flash('PIN diskon harus 4-8 digit angka.', 'danger')
            return redirect(url_for('pos_master') + '#card-kunci-diskon')
        conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_diskon_pin', ?)",
                     (sha256(pin_baru.encode()).hexdigest(),))
        pin_note = 'PIN diperbarui'
    else:
        pin_note = 'PIN tidak diubah'
    add_log(conn, 'Update Kunci Diskon POS', f'lock={lock} | {pin_note}', 'POS')
    conn.commit()
    conn.close()
    flash(f'Pengaturan kunci diskon disimpan ({pin_note}).', 'success')
    return redirect(url_for('pos_master') + '#card-kunci-diskon')

@app.route('/pos/master/stok', methods=['POST'])
@master_pos_required
def pos_master_stok():
    """Toggle mode blokir stok minus di POS. Default OFF (jual dulu, barang menyusul).
    ON = checkout ditolak bila stok produk SKU / bahan resep tidak cukup."""
    conn = db()
    block = '1' if request.form.get('pos_block_minus') else '0'
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_block_minus', ?)", (block,))
    add_log(conn, 'Update Blokir Stok Minus POS', f'block={block}', 'POS')
    conn.commit()
    conn.close()
    flash('Pengaturan stok POS disimpan (' + ('blokir stok minus AKTIF' if block == '1'
          else 'stok minus diizinkan') + ').', 'success')
    return redirect(url_for('pos_master') + '#card-stok-minus')

@app.route('/pos/master/tampilan-item', methods=['POST'])
@master_pos_required
def pos_master_show_mode():
    """Atur item apa yang tampil di grid kasir: keduanya / barang jadi saja / racikan saja.
    Berguna utk usaha (mis. FnB made-to-order) yang tak perlu menampilkan barang jadi."""
    conn = db()
    mode = (request.form.get('pos_show_mode', 'both') or 'both').strip()
    if mode not in ('both', 'produk', 'resep'):
        mode = 'both'
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_show_mode', ?)", (mode,))
    add_log(conn, 'Update Mode Tampilan Item POS', f'mode={mode}', 'POS')
    conn.commit()
    conn.close()
    _label = {'both': 'barang jadi + racikan', 'produk': 'barang jadi saja', 'resep': 'racikan saja'}[mode]
    flash(f'Tampilan item kasir disimpan ({_label}).', 'success')
    return redirect(url_for('pos_master') + '#card-show-mode')

@app.route('/pos/master/kategori', methods=['POST'])
@master_pos_required
def pos_master_kategori():
    """Sembunyikan / tampilkan KATEGORI di grid kasir, sekali jalan (bukan per produk).

    Praktis untuk kategori yang bukan barang dagangan (mis. "Bahan Baku"): tak perlu
    meng-uncheck produknya satu per satu di POS Lanjutan, dan produk baru yang masuk ke
    kategori itu ikut tersembunyi otomatis. Murni tampilan — data produk tidak disentuh
    sama sekali, jadi mengaktifkan kembali cukup satu klik."""
    conn = db()
    # Form mengirim kategori yang TAMPIL (checkbox tercentang) + daftar acuan lengkap
    # ('kategori_all'). Acuan itu dipakai supaya kategori yang baru muncul SETELAH halaman
    # ini dibuka tidak ikut tersembunyi diam-diam hanya karena tidak ada di form.
    tampil = {(k or '').strip().lower() for k in request.form.getlist('kategori_tampil')}
    hide, seen = [], set()
    for k in request.form.getlist('kategori_all'):
        k = (k or '').strip()
        if not k or k.lower() in seen:
            continue
        seen.add(k.lower())
        if k.lower() not in tampil:
            hide.append(k)
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_hide_kategori', ?)",
                 (json.dumps(hide),))
    add_log(conn, 'Update Kategori Tampil POS',
            (', '.join(hide) if hide else 'semua kategori tampil'), 'POS')
    conn.commit()
    conn.close()
    if hide:
        flash(f'Pengaturan kategori kasir disimpan. {len(hide)} kategori disembunyikan: '
              + ', '.join(hide) + '.', 'success')
    else:
        flash('Pengaturan kategori kasir disimpan. Semua kategori tampil di kasir.', 'success')
    return redirect(url_for('pos_master') + '#card-kategori')

def _pos_active_levels(conn):
    """Daftar nama level harga aktif (maks 5, tanpa 'Utama')."""
    try:
        levels = json.loads(get_setting(conn, 'pos_harga_levels', '[]') or '[]')
        if not isinstance(levels, list):
            levels = []
    except (ValueError, TypeError):
        levels = []
    return [str(x).strip()[:30] for x in levels if str(x).strip()][:5]

def _pos_item_levels(conn, tipe, iid, levels):
    """Harga per level {label: harga} untuk satu item (produk/resep), dibatasi level aktif."""
    tbl, col = ('produksi_template_harga', 'template_id') if tipe == 'resep' else ('produk_harga', 'produk_id')
    raw = {}
    for r in conn.execute(f"SELECT label, harga FROM {tbl} WHERE {col}=?", (iid,)):
        raw[r['label']] = float(r['harga'] or 0)
    return {lv: raw.get(lv, 0) for lv in levels}

@app.route('/pos/master/produk')
@master_pos_required
def pos_master_produk():
    """Halaman Lanjutan POS Master: atur multi-harga + tampil/harga per PRODUK & RACIKAN.
    Skala ribuan: daftar penuh TIDAK dirender — pemilihan item lewat pencarian (AJAX).
    Halaman Stok / Produk / Produksi TIDAK disentuh."""
    conn = db()
    multi_harga = get_setting(conn, 'pos_multi_harga_aktif', '0') == '1'
    levels = _pos_active_levels(conn)
    n_produk = conn.execute("SELECT COUNT(*) c FROM produk WHERE COALESCE(aktif,1)=1").fetchone()['c']
    n_resep = conn.execute("SELECT COUNT(*) c FROM produksi_template").fetchone()['c']
    conn.close()
    return render_template('pos_master_produk.html',
                           levels=levels, multi_harga=multi_harga,
                           n_produk=n_produk, n_resep=n_resep)

@app.route('/pos/master/produk/cari')
@master_pos_required
def pos_master_produk_cari():
    """AJAX: cari produk dan/atau racikan by nama/kode. Return item + status tampil di kasir
    + harga utama + harga per level (untuk kartu edit). LIMIT supaya ringan walau ribuan."""
    conn = db()
    q = (request.args.get('q', '') or '').strip()
    tipe = (request.args.get('tipe', 'all') or 'all').strip()
    if tipe not in ('all', 'produk', 'resep'):
        tipe = 'all'
    levels = _pos_active_levels(conn)
    like = f"%{q}%"
    items = []
    if tipe in ('all', 'produk'):
        sql = ("SELECT id, kode, nama, varian, kategori, COALESCE(harga_jual,0) AS harga_jual, "
               "COALESCE(pos_tampil,1) AS tampil FROM produk WHERE COALESCE(aktif,1)=1 ")
        params = []
        if q:
            sql += "AND (nama LIKE ? OR kode LIKE ? OR IFNULL(varian,'') LIKE ?) "
            params += [like, like, like]
        sql += "ORDER BY nama COLLATE NOCASE LIMIT 25"
        for r in conn.execute(sql, params):
            items.append({'tipe': 'produk', 'id': r['id'], 'kode': r['kode'] or '',
                          'nama': r['nama'] + (f" ({r['varian']})" if r['varian'] else ''),
                          'kategori': r['kategori'] or '',
                          'harga_utama': float(r['harga_jual'] or 0), 'harga_utama_editable': False,
                          'tampil': int(r['tampil']),
                          'prices': _pos_item_levels(conn, 'produk', r['id'], levels)})
    if tipe in ('all', 'resep'):
        sql = ("SELECT id, nama, kategori, COALESCE(harga_jual,0) AS harga_jual, "
               "COALESCE(pos_aktif,0) AS tampil FROM produksi_template ")
        params = []
        if q:
            sql += "WHERE nama LIKE ? "
            params += [like]
        sql += "ORDER BY nama COLLATE NOCASE LIMIT 25"
        for r in conn.execute(sql, params):
            items.append({'tipe': 'resep', 'id': r['id'], 'kode': 'RACIKAN',
                          'nama': r['nama'], 'kategori': r['kategori'] or '',
                          'harga_utama': float(r['harga_jual'] or 0), 'harga_utama_editable': True,
                          'tampil': int(r['tampil']),
                          'prices': _pos_item_levels(conn, 'resep', r['id'], levels)})
    conn.close()
    return jsonify(ok=True, items=items, levels=levels)

@app.route('/pos/master/produk/simpan-item', methods=['POST'])
@master_pos_required
def pos_master_item_simpan():
    """AJAX: simpan SATU item (produk/resep): tampil di kasir + harga per level (+ harga jual
    untuk resep). Produk: tampil=pos_tampil, harga Utama read-only (dari Stok). Resep:
    tampil=pos_aktif, harga Utama = harga_jual (editable). Tidak menyentuh akuntansi."""
    data = request.get_json(silent=True) or {}
    tipe = (data.get('tipe') or '').strip()
    try:
        iid = int(data.get('id') or 0)
    except (ValueError, TypeError):
        iid = 0
    if tipe not in ('produk', 'resep') or iid <= 0:
        return jsonify(ok=False, error='Item tidak valid.'), 400
    tampil = 1 if data.get('tampil') else 0
    conn = db()
    levels = _pos_active_levels(conn)
    lv_in = data.get('levels') or {}
    def _h(v):
        try:
            return parse_nonnegative_rp(v) if str(v).strip() else 0
        except ValueError:
            return 0
    if tipe == 'produk':
        if not conn.execute("SELECT 1 FROM produk WHERE id=? AND COALESCE(aktif,1)=1", (iid,)).fetchone():
            conn.close(); return jsonify(ok=False, error='Produk tidak ditemukan.'), 404
        conn.execute("UPDATE produk SET pos_tampil=? WHERE id=?", (tampil, iid))
        conn.execute("DELETE FROM produk_harga WHERE produk_id=?", (iid,))
        for lv in levels:
            h = _h(lv_in.get(lv, ''))
            if h > 0:
                conn.execute("INSERT INTO produk_harga(produk_id,label,harga) VALUES(?,?,?)", (iid, lv, h))
    else:
        if not conn.execute("SELECT 1 FROM produksi_template WHERE id=?", (iid,)).fetchone():
            conn.close(); return jsonify(ok=False, error='Racikan tidak ditemukan.'), 404
        hj = _h(data.get('harga_utama', ''))
        conn.execute("UPDATE produksi_template SET pos_aktif=?, harga_jual=? WHERE id=?", (tampil, hj, iid))
        conn.execute("DELETE FROM produksi_template_harga WHERE template_id=?", (iid,))
        for lv in levels:
            h = _h(lv_in.get(lv, ''))
            if h > 0:
                conn.execute("INSERT INTO produksi_template_harga(template_id,label,harga) VALUES(?,?,?)", (iid, lv, h))
    add_log(conn, 'Atur item POS (Lanjutan)', f'{tipe}#{iid} tampil={tampil}', 'POS')
    conn.commit(); conn.close()
    return jsonify(ok=True)

@app.route('/pos/master/produk/simpan', methods=['POST'])
@master_pos_required
def pos_master_produk_simpan():
    """Simpan pengaturan GLOBAL: toggle multi-harga + nama level harga. Tampil & harga per
    item disimpan terpisah lewat /pos/master/produk/simpan-item (AJAX)."""
    conn = db()
    multi = '1' if request.form.get('pos_multi_harga_aktif') else '0'
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_multi_harga_aktif', ?)", (multi,))
    raw_levels = (request.form.get('harga_levels', '') or '').split(',')
    levels, seen = [], set()
    for lv in raw_levels:
        lv = lv.strip()[:30]
        if lv and lv.lower() != 'utama' and lv.lower() not in seen:
            seen.add(lv.lower()); levels.append(lv)
        if len(levels) >= 5:
            break
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_harga_levels', ?)", (json.dumps(levels),))
    # Bersihkan harga level yang labelnya sudah tak ada lagi (level di-rename/hapus) supaya
    # tak jadi harga hantu di kasir. Kalau level dikosongkan semua, hapus seluruh harga level.
    if levels:
        ph = ','.join('?' * len(levels))
        conn.execute(f"DELETE FROM produk_harga WHERE label NOT IN ({ph})", levels)
        conn.execute(f"DELETE FROM produksi_template_harga WHERE label NOT IN ({ph})", levels)
    else:
        conn.execute("DELETE FROM produk_harga")
        conn.execute("DELETE FROM produksi_template_harga")
    add_log(conn, 'Update Multi-Harga POS', f'aktif={multi} | level={len(levels)}', 'POS')
    conn.commit()
    conn.close()
    flash('Pengaturan multi-harga disimpan.', 'success')
    return redirect(url_for('pos_master_produk'))

@app.route('/pos/diskon/verify', methods=['POST'])
@pos_required
def pos_diskon_verify():
    """Validasi PIN supervisor untuk membuka diskon nota terkunci (AJAX)."""
    conn = db()
    pin_hash = (get_setting(conn, 'pos_diskon_pin', '') or '').strip()
    conn.close()
    pin_in = (request.form.get('pin', '') or '').strip()
    ok = bool(pin_hash) and bool(pin_in) and sha256(pin_in.encode()).hexdigest() == pin_hash
    return jsonify({'ok': ok})

@app.route('/pos/master/toggle-kasir/<int:uid>', methods=['POST'])
@master_pos_required
def pos_master_toggle_kasir(uid):
    """Aktifkan/non-aktifkan kasir (tidak menghapus, agar history shift tetap utuh)."""
    conn = db()
    u = conn.execute("SELECT * FROM users WHERE id=? AND role='KASIR'", (uid,)).fetchone()
    if not u:
        conn.close()
        flash('Kasir tidak ditemukan.', 'danger')
        return redirect(url_for('pos_master'))
    new_state = 0 if u['aktif'] else 1
    conn.execute("UPDATE users SET aktif=? WHERE id=?", (new_state, uid))
    add_log(conn, 'Toggle Status Kasir',
            f"{u['username']} -> {'AKTIF' if new_state else 'NONAKTIF'}", 'POS')
    conn.commit()
    conn.close()
    flash(f"Kasir {u['username']} {'diaktifkan' if new_state else 'di-nonaktifkan'}.", 'info')
    return redirect(url_for('pos_master'))


def auto_penyusutan(conn):
    """Catat penyusutan yang belum dijurnal untuk semua aset aktif hingga bulan berjalan."""
    today = date.today()
    asets = conn.execute(
        "SELECT * FROM aset_tetap WHERE aktif=1 AND penyusutan_bulan > 0"
    ).fetchall()
    changed = False
    for aset in asets:
        try:
            tgl_beli = datetime.strptime(str(aset['tanggal_beli'])[:10], '%Y-%m-%d').date()
            months_elapsed = (today.year - tgl_beli.year) * 12 + (today.month - tgl_beli.month) + 1
            months_to_record = min(months_elapsed, aset['masa_pakai'])
            months_recorded  = aset['bulan_penyusutan_dicatat'] or 0
            if months_to_record <= months_recorded:
                continue
            months_new = months_to_record - months_recorded
            amount     = round(months_new * aset['penyusutan_bulan'], 2)
            if amount <= 0:
                continue
            ket = (f"Penyusutan {aset['nama']}: "
                   f"bulan {months_recorded+1}–{months_to_record} / {aset['masa_pakai']}")
            jid_dep = insert_jurnal(conn, str(today), ket, 'OPERASIONAL', 'PENYUSUTAN', [
                ('6130', amount, 0),
                ('1290', 0, amount),
            ])
            conn.execute("UPDATE jurnal SET aset_id=? WHERE id=?", (aset['id'], jid_dep))
            conn.execute(
                "UPDATE aset_tetap SET bulan_penyusutan_dicatat=?, akumulasi_penyusutan=akumulasi_penyusutan+? WHERE id=?",
                (months_to_record, amount, aset['id'])
            )
            changed = True
        except Exception:
            continue
    return changed


# ---------- DASHBOARD ----------
@app.route('/')
@login_required
def dashboard():
    perms = current_permissions()
    if session.get('role') == 'KASIR' and perms <= {'pos_sales'}:
        return redirect(url_for('pos_home'))
    if session.get('role') == 'MANAJER' and perms <= {'manage_master_pos'} and _pos_module_active():
        return redirect(url_for('pos_master'))
    today = date.today()
    sd = request.args.get('sd', today.replace(day=1).strftime('%Y-%m-%d'))
    ed = request.args.get('ed', today.strftime('%Y-%m-%d'))

    conn = db()
    if auto_penyusutan(conn):
        conn.commit()
    pnl   = calc_profitability(conn, sd, ed)
    cf    = calc_cashflow(conn, sd, ed)
    ratios = calc_financial_ratios(conn, sd, ed)

    total_piutang = conn.execute(
        "SELECT COALESCE(SUM(jumlah-terbayar),0) FROM piutang WHERE status!='LUNAS'"
    ).fetchone()[0]
    total_hutang = conn.execute(
        "SELECT COALESCE(SUM(jumlah-terbayar),0) FROM hutang WHERE status!='LUNAS'"
    ).fetchone()[0]

    piutang_count = conn.execute(
        "SELECT COUNT(*) FROM piutang WHERE status!='LUNAS'"
    ).fetchone()[0]
    hutang_count = conn.execute(
        "SELECT COUNT(*) FROM hutang WHERE status!='LUNAS'"
    ).fetchone()[0]
    piutang_overdue = conn.execute(
        "SELECT COUNT(*) FROM piutang WHERE status!='LUNAS' AND jatuh_tempo < ?", (today,)
    ).fetchone()[0]
    hutang_overdue = conn.execute(
        "SELECT COUNT(*) FROM hutang WHERE status!='LUNAS' AND jatuh_tempo < ?", (today,)
    ).fetchone()[0]

    piutang_lunas = conn.execute(
        "SELECT COUNT(*) FROM piutang WHERE status='LUNAS'"
    ).fetchone()[0]
    piutang_total = conn.execute(
        "SELECT COUNT(*) FROM piutang"
    ).fetchone()[0]
    hutang_lunas = conn.execute(
        "SELECT COUNT(*) FROM hutang WHERE status='LUNAS'"
    ).fetchone()[0]
    hutang_total = conn.execute(
        "SELECT COUNT(*) FROM hutang"
    ).fetchone()[0]
    hutang_jatuh = conn.execute(
        "SELECT * FROM hutang WHERE status!='LUNAS' AND jatuh_tempo <= ? ORDER BY jatuh_tempo LIMIT 5",
        ((today + timedelta(days=7)),)
    ).fetchall()

    transaksi_baru = conn.execute("""
        SELECT j.id, j.tanggal, j.keterangan, j.referensi, j.kategori, j.tipe_tx,
               COALESCE(SUM(d.debit),0) as total
        FROM jurnal j LEFT JOIN detail_jurnal d ON d.jurnal_id=j.id
        GROUP BY j.id ORDER BY j.tanggal DESC, j.id DESC LIMIT 10
    """).fetchall()

    stok_rendah = conn.execute(
        "SELECT * FROM produk WHERE COALESCE(aktif,1)=1 "
        "AND (stok<0 OR (stok<=min_stok AND min_stok>0)) ORDER BY stok LIMIT 5"
    ).fetchall()
    piutang_jatuh = conn.execute(
        "SELECT * FROM piutang WHERE status!='LUNAS' AND jatuh_tempo <= ? ORDER BY jatuh_tempo LIMIT 5",
        ((today + timedelta(days=7)),)
    ).fetchall()

    nama_usaha = get_setting(conn, 'nama_usaha', 'Usaha Saya')
    saldo_rekening = get_rekening_saldo(conn)
    earliest_row = conn.execute("SELECT MIN(tanggal) FROM jurnal").fetchone()
    earliest_date = earliest_row[0] if earliest_row and earliest_row[0] else today.strftime('%Y-%m-%d')

    available_years = [row[0] for row in conn.execute(
        "SELECT DISTINCT CAST(strftime('%Y', tanggal) AS INTEGER) FROM jurnal ORDER BY 1 DESC"
    ).fetchall()]
    if not available_years:
        available_years = [today.year]

    # Mini insight: pendapatan tertinggi bulanan tahun ini
    year = today.year
    monthly_rev = []
    for m in range(1, today.month + 1):
        last_day = calendar.monthrange(year, m)[1]
        ms = f'{year}-{m:02d}-01'
        me = f'{year}-{m:02d}-{last_day:02d}'
        v = _qv(conn, "a.tipe='PENDAPATAN' AND j.tanggal>=? AND j.tanggal<=?", [ms, me], 'KREDIT')
        monthly_rev.append(v)
    max_rev_year = max(monthly_rev) if monthly_rev else 0

    # ── Breakdown cashflow: masuk per rekening, keluar per kategori ──────────
    cf_masuk_rek = [
        {'nama': r['nama'], 'total': r['total']}
        for r in conn.execute("""
            SELECT a.nama, COALESCE(SUM(d.debit),0) AS total
            FROM detail_jurnal d
            JOIN jurnal j ON j.id = d.jurnal_id
            JOIN akun a ON a.id = d.akun_id
            WHERE a.is_rekening=1 AND d.debit>0
              AND j.tanggal>=? AND j.tanggal<=?
              AND COALESCE(j.tipe_tx,'')!='TRANSFER'
            GROUP BY a.id, a.nama
            HAVING total>0 ORDER BY total DESC
        """, [sd, ed]).fetchall()
    ]
    cf_keluar_kat = [
        {'label': r['label'], 'total': r['total']}
        for r in conn.execute("""
            SELECT
              CASE
                WHEN a.kode = '4150'                  THEN 'Retur Penjualan'
                WHEN a.kode IN ('2115','2116','2120') THEN 'Gaji & Tenaga Kerja'
                WHEN a.kode LIKE '113%'               THEN 'Bahan Baku'
                WHEN a.kode = '1160'                  THEN 'Pinjaman Karyawan'
                WHEN a.kode LIKE '12%'                THEN 'Investasi Aset'
                WHEN COALESCE(a.peran,'') = 'PAJAK'    THEN 'Pajak'
                WHEN a.kode LIKE '61%'                THEN 'Beban Operasional'
                WHEN a.kode = '3300'                  THEN 'Prive'
                WHEN a.kode LIKE '2%'                 THEN 'Pelunasan Hutang'
                ELSE 'Lainnya'
              END AS label,
              COALESCE(SUM(d.debit),0) AS total
            FROM detail_jurnal d
            JOIN jurnal j ON j.id = d.jurnal_id
            JOIN akun a ON a.id = d.akun_id
            WHERE a.is_rekening=0 AND d.debit>0
              AND j.tanggal>=? AND j.tanggal<=?
              AND EXISTS (
                  SELECT 1 FROM detail_jurnal d2
                  JOIN akun a2 ON a2.id = d2.akun_id
                  WHERE d2.jurnal_id = d.jurnal_id
                    AND a2.is_rekening=1 AND d2.kredit>0
              )
              -- Kecualikan reklasifikasi non-kas: retur penjualan membalik HPP lewat
              -- Dr 1130 (Persediaan) / Cr 5100-5150-5190 (HPP) dalam jurnal yang SAMA.
              -- Pasangan ini tidak pernah menyentuh kas, jadi tidak boleh ikut terhitung
              -- sebagai "uang keluar" (dulu bikin donut overcount + salah kategori Bahan Baku).
              AND NOT (
                  a.kode = '1130'
                  AND EXISTS (
                      SELECT 1 FROM detail_jurnal d3
                      JOIN akun a3 ON a3.id = d3.akun_id
                      WHERE d3.jurnal_id = d.jurnal_id
                        AND a3.kode IN ('5100','5150','5190') AND d3.kredit>0
                  )
              )
            GROUP BY 1 HAVING total>0 ORDER BY total DESC
        """, [sd, ed]).fetchall()
    ]

    conn.close()

    session['dash_sd'] = sd
    session['dash_ed'] = ed

    return render_template('dashboard.html',
        sd=sd, ed=ed, pnl=pnl, cf=cf, ratios=ratios,
        max_rev_year=max_rev_year,
        total_piutang=total_piutang, total_hutang=total_hutang,
        piutang_count=piutang_count, hutang_count=hutang_count,
        piutang_overdue=piutang_overdue, hutang_overdue=hutang_overdue,
        piutang_lunas=piutang_lunas, piutang_total=piutang_total,
        hutang_lunas=hutang_lunas, hutang_total=hutang_total,
        hutang_jatuh=hutang_jatuh,
        transaksi_baru=transaksi_baru,
        stok_rendah=stok_rendah, piutang_jatuh=piutang_jatuh,
        nama_usaha=nama_usaha, saldo_rekening=saldo_rekening,
        earliest_date=earliest_date,
        available_years=available_years,
        cf_masuk_rek=cf_masuk_rek, cf_keluar_kat=cf_keluar_kat,
    )


# ---------- PEMASUKAN ----------
def set_inventory_value_alert(required_value, available_value):
    session['inventory_value_alert'] = {
        'required': round(float(required_value or 0), 2),
        'available': round(float(available_value or 0), 2),
        'shortage': round(max(0, float(required_value or 0) - float(available_value or 0)), 2),
        'resulting': round(float(available_value or 0) - float(required_value or 0), 2),
    }

# ---------- PROYEK (JOB COSTING) ----------
PROYEK_STATUS = ('AKTIF', 'SELESAI', 'BATAL')

def _proyek_summary(conn, proyek_id):
    """Ringkasan finansial satu proyek dari semua jurnal ber-tag proyek_id."""
    row = conn.execute("""
        SELECT
          COALESCE(SUM(CASE WHEN a.tipe='PENDAPATAN' THEN dj.kredit - dj.debit ELSE 0 END), 0) AS pendapatan,
          COALESCE(SUM(CASE WHEN a.tipe='BEBAN' THEN dj.debit - dj.kredit ELSE 0 END), 0) AS biaya
        FROM detail_jurnal dj
        JOIN jurnal j ON j.id = dj.jurnal_id
        JOIN akun a ON a.id = dj.akun_id
        WHERE j.proyek_id = ?
    """, (proyek_id,)).fetchone()
    pendapatan = float(row['pendapatan'] or 0)
    biaya = float(row['biaya'] or 0)
    laba = pendapatan - biaya
    margin = (laba / pendapatan * 100.0) if pendapatan > 0 else 0.0
    return {'pendapatan': pendapatan, 'biaya': biaya, 'laba': laba, 'margin': margin}

def _proyek_arap(conn, proyek_id):
    """Posisi penagihan proyek (DISPLAY-ONLY, tidak menulis jurnal):
    piutang belum tertagih & hutang vendor belum dibayar — di-derive dari tabel
    piutang/hutang yang jurnal sumbernya ber-tag proyek_id. Outstanding =
    jumlah - terbayar (independen dari tag jurnal pelunasan)."""
    ar_rows = conn.execute("""
        SELECT p.id, p.tanggal, p.jatuh_tempo, p.pelanggan AS pihak, p.keterangan,
               p.jumlah, COALESCE(p.terbayar,0) AS terbayar,
               (p.jumlah - COALESCE(p.terbayar,0)) AS sisa, p.status
          FROM piutang p
          JOIN jurnal j ON j.id = p.jurnal_id
         WHERE j.proyek_id = ? AND (p.jumlah - COALESCE(p.terbayar,0)) > 0.005
         ORDER BY p.jatuh_tempo IS NULL, p.jatuh_tempo, p.id
    """, (proyek_id,)).fetchall()
    ap_rows = conn.execute("""
        SELECT h.id, h.tanggal, h.jatuh_tempo, h.pemasok AS pihak, h.keterangan,
               h.jumlah, COALESCE(h.terbayar,0) AS terbayar,
               (h.jumlah - COALESCE(h.terbayar,0)) AS sisa, h.status
          FROM hutang h
          JOIN jurnal j ON j.id = h.jurnal_id
         WHERE j.proyek_id = ? AND (h.jumlah - COALESCE(h.terbayar,0)) > 0.005
         ORDER BY h.jatuh_tempo IS NULL, h.jatuh_tempo, h.id
    """, (proyek_id,)).fetchall()
    piutang_total = round(sum(float(r['sisa']) for r in ar_rows), 2)
    hutang_total  = round(sum(float(r['sisa']) for r in ap_rows), 2)
    return {'ar_rows': ar_rows, 'ap_rows': ap_rows,
            'piutang_outstanding': piutang_total, 'hutang_outstanding': hutang_total}

@app.route('/proyek', methods=['GET', 'POST'])
@finance_required
def proyek_list():
    conn = db()
    if request.method == 'POST':
        action = request.form.get('action', 'tambah')
        if action in ('tambah', 'edit'):
            kode = (request.form.get('kode') or '').strip().upper()
            nama = (request.form.get('nama') or '').strip()
            pelanggan = (request.form.get('pelanggan') or '').strip()
            catatan = (request.form.get('catatan') or '').strip()
            tanggal_mulai = (request.form.get('tanggal_mulai') or '').strip() or None
            target_selesai = (request.form.get('target_selesai') or '').strip() or None
            if not kode or not nama:
                conn.close(); flash('Kode dan nama proyek wajib diisi.', 'danger')
                return redirect(url_for('proyek_list'))
            if not is_valid_optional_iso_date(tanggal_mulai) or not is_valid_optional_iso_date(target_selesai):
                conn.close(); flash('Tanggal proyek tidak valid.', 'danger')
                return redirect(url_for('proyek_list'))
            try:
                nilai_kontrak = parse_nonnegative_rp(request.form.get('nilai_kontrak', '0'), 'Nilai kontrak')
                anggaran_biaya = parse_nonnegative_rp(request.form.get('anggaran_biaya', '0'), 'Anggaran biaya')
            except ValueError as ex:
                conn.close(); flash(str(ex), 'danger')
                return redirect(url_for('proyek_list'))
            if action == 'tambah':
                try:
                    conn.execute(
                        "INSERT INTO proyek(kode,nama,pelanggan,nilai_kontrak,anggaran_biaya,"
                        "tanggal_mulai,target_selesai,catatan) VALUES(?,?,?,?,?,?,?,?)",
                        (kode, nama, pelanggan, nilai_kontrak, anggaran_biaya,
                         tanggal_mulai, target_selesai, catatan))
                except sqlite3.IntegrityError:
                    conn.close(); flash(f'Kode proyek {kode} sudah dipakai.', 'danger')
                    return redirect(url_for('proyek_list'))
                add_log(conn, 'Proyek dibuat', f'{kode} - {nama}', 'INPUT')
                flash('Proyek baru dibuat!', 'success')
            else:
                pid = request.form.get('id', type=int)
                if not conn.execute("SELECT id FROM proyek WHERE id=?", (pid,)).fetchone():
                    conn.close(); flash('Proyek tidak ditemukan.', 'danger')
                    return redirect(url_for('proyek_list'))
                try:
                    conn.execute(
                        "UPDATE proyek SET kode=?,nama=?,pelanggan=?,nilai_kontrak=?,anggaran_biaya=?,"
                        "tanggal_mulai=?,target_selesai=?,catatan=? WHERE id=?",
                        (kode, nama, pelanggan, nilai_kontrak, anggaran_biaya,
                         tanggal_mulai, target_selesai, catatan, pid))
                except sqlite3.IntegrityError:
                    conn.close(); flash(f'Kode proyek {kode} sudah dipakai.', 'danger')
                    return redirect(url_for('proyek_list'))
                flash('Proyek diperbarui!', 'success')
        elif action == 'status':
            pid = request.form.get('id', type=int)
            status = request.form.get('status', '')
            if status not in PROYEK_STATUS:
                conn.close(); flash('Status proyek tidak valid.', 'danger')
                return redirect(url_for('proyek_list'))
            conn.execute("UPDATE proyek SET status=? WHERE id=?", (status, pid))
            flash(f'Status proyek diubah ke {status}.', 'success')
        elif action == 'hapus':
            pid = request.form.get('id', type=int)
            n_tx = conn.execute("SELECT COUNT(*) c FROM jurnal WHERE proyek_id=?", (pid,)).fetchone()['c']
            if n_tx:
                conn.close()
                flash(f'Proyek tidak bisa dihapus: masih ada {n_tx} transaksi ber-tag proyek ini. '
                      'Ubah status ke SELESAI atau BATAL saja.', 'warning')
                return redirect(url_for('proyek_list'))
            conn.execute("DELETE FROM proyek WHERE id=?", (pid,))
            flash('Proyek dihapus.', 'success')
        conn.commit(); conn.close()
        return redirect(url_for('proyek_list'))

    page = _page_arg()
    n_rows = conn.execute("SELECT COUNT(*) FROM proyek").fetchone()[0]
    pg = paginate_meta(page, n_rows)
    rows = conn.execute(
        "SELECT * FROM proyek ORDER BY CASE status WHEN 'AKTIF' THEN 0 WHEN 'SELESAI' THEN 1 ELSE 2 END, dibuat DESC LIMIT ? OFFSET ?",
        (pg['per_page'], pg['offset'])
    ).fetchall()
    proyek_rows = []
    for p in rows:
        s = _proyek_summary(conn, p['id'])
        arap = _proyek_arap(conn, p['id'])
        n_tx = conn.execute("SELECT COUNT(*) c FROM jurnal WHERE proyek_id=?", (p['id'],)).fetchone()['c']
        proyek_rows.append({**dict(p), **s, 'n_tx': n_tx,
                            'piutang_outstanding': arap['piutang_outstanding'],
                            'hutang_outstanding': arap['hutang_outstanding']})
    conn.close()
    return render_template('proyek_list.html', proyek_rows=proyek_rows,
                           today=date.today().strftime('%Y-%m-%d'), pg=pg)

@app.route('/proyek/<int:id>')
@finance_required
def proyek_detail(id):
    conn = db()
    p = conn.execute("SELECT * FROM proyek WHERE id=?", (id,)).fetchone()
    if not p:
        conn.close(); flash('Proyek tidak ditemukan.', 'danger')
        return redirect(url_for('proyek_list'))
    s = _proyek_summary(conn, id)
    tx_rows = conn.execute("""
        SELECT j.id, j.tanggal, j.keterangan, j.tipe_tx, j.pihak, j.nomor_tx,
          COALESCE(SUM(CASE WHEN a.tipe='PENDAPATAN' THEN dj.kredit - dj.debit ELSE 0 END), 0) AS pendapatan,
          COALESCE(SUM(CASE WHEN a.tipe='BEBAN' THEN dj.debit - dj.kredit ELSE 0 END), 0) AS biaya
        FROM jurnal j
        JOIN detail_jurnal dj ON dj.jurnal_id = j.id
        JOIN akun a ON a.id = dj.akun_id
        WHERE j.proyek_id = ?
        GROUP BY j.id ORDER BY j.tanggal DESC, j.id DESC
    """, (id,)).fetchall()
    biaya_per_akun = conn.execute("""
        SELECT a.kode, a.nama, SUM(dj.debit - dj.kredit) AS total
        FROM detail_jurnal dj
        JOIN jurnal j ON j.id = dj.jurnal_id
        JOIN akun a ON a.id = dj.akun_id
        WHERE j.proyek_id = ? AND a.tipe = 'BEBAN'
        GROUP BY a.kode HAVING ABS(total) > 0.005 ORDER BY total DESC
    """, (id,)).fetchall()
    arap = _proyek_arap(conn, id)
    conn.close()
    return render_template('proyek_detail.html', p=p, s=s, tx_rows=tx_rows,
                           biaya_per_akun=biaya_per_akun, arap=arap)

# ─────────────────────────────────────────────────────────────────────────
# IMPORT PENJUALAN MARKETPLACE (Shopee / TikTok-Tokopedia) — V6
# Baca laporan analytics omzet HARIAN (.xlsx, data saja, file tak disimpan) →
# preview editable → jurnal per-hari Dr Kas / Cr 4100. Idempotency per-hari.
# ─────────────────────────────────────────────────────────────────────────
MP_PLATFORMS = {'SHOPEE': 'Shopee', 'TIKTOK': 'TikTok/Tokopedia'}

def _mp_num(v):
    """Angka dari sel: dukung numeric langsung & format Indonesia ('.'ribuan, ','desimal)."""
    if v is None or isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s in ('-', '—', '–'):
        return 0.0
    s = s.replace('Rp', '').replace('IDR', '').replace(' ', '').replace('%', '')
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')     # 140.723,26 -> 140723.26
        else:
            s = s.replace(',', '')                         # 1,792,883.00 -> 1792883.00
    elif ',' in s:
        s = s.replace(',', '.')                            # desimal koma
    else:
        # hanya titik: pemisah ribuan, KECUALI satu titik dgn ekor != 3 digit = desimal
        parts = s.split('.')
        if len(parts) == 2 and len(parts[1]) != 3:
            pass                                           # 1234.56 -> desimal, biarkan
        else:
            s = s.replace('.', '')                         # 6.613.993 -> 6613993 (ribuan)
    try:
        return float(s)
    except ValueError:
        return 0.0

def _mp_date(v):
    """ISO 'YYYY-MM-DD' bila sel berisi SATU tanggal. Dukung: date/datetime, ISO
    'YYYY-MM-DD[ HH:MM]', dan DD-MM-YYYY / DD/MM/YYYY (opsional diikuti jam). Else None."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, date):
        return v.strftime('%Y-%m-%d')
    import re
    s = str(v).strip()
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?$', s)      # ISO, mis. '2026-05-01 02:30'
    if m:
        y, mo, d = m.groups()
        try:
            return date(int(y), int(mo), int(d)).strftime('%Y-%m-%d')
        except ValueError:
            return None
    m = re.match(r'^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})(?:[ T].*)?$', s)   # DD-MM-YYYY (opsional jam)
    if not m:
        return None
    d, mo, y = m.groups()
    try:
        return date(int(y), int(mo), int(d)).strftime('%Y-%m-%d')   # validasi kalender (tolak 31-02)
    except ValueError:
        return None

def _mp_detect(wb):
    """Deteksi platform dari struktur workbook."""
    if any('pesanan' in (s or '').lower() for s in wb.sheetnames):
        return 'SHOPEE'
    ws = wb.worksheets[0]
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i > 12:
            break
        joined = ' '.join(str(c) for c in row if c is not None).lower()
        if 'tanggal analisis' in joined or ('gmv' in joined and 'pendapatan bruto' in joined):
            return 'TIKTOK'
    return None

_MP_OMZET_EXCLUDE = ('dibatalkan', 'dikembalikan', 'batal', 'retur', 'refund', 'cancel')

def _mp_collect_days(ws, omzet_headers, date_col=0):
    """Temukan kolom omzet (exact-match diutamakan, hindari kolom batal/retur),
    lalu kumpulkan baris ber-tanggal (distinct, dibatasi MAX_ROWS). Omzet -> rupiah penuh."""
    MAX_ROWS = 20000
    rows = []
    for r in ws.iter_rows(values_only=True):
        rows.append(r)
        if len(rows) >= MAX_ROWS:
            break

    def _find_col(exact):
        for row in rows:
            for ci, cell in enumerate(row):
                if cell is None:
                    continue
                label = str(cell).strip().lower()
                if any(x in label for x in _MP_OMZET_EXCLUDE):
                    continue
                if exact and any(label == h for h in omzet_headers):
                    return ci
                if not exact and any(label.startswith(h) for h in omzet_headers):
                    return ci
        return None

    omzet_col = _find_col(True)
    if omzet_col is None:
        omzet_col = _find_col(False)
    if omzet_col is None:
        return []
    days, seen = [], set()
    for row in rows:
        if date_col >= len(row):
            continue
        iso = _mp_date(row[date_col])
        if not iso or iso in seen:
            continue
        omz = _mp_num(row[omzet_col]) if omzet_col < len(row) else 0.0
        seen.add(iso)
        days.append({'tgl': iso, 'omzet': round(omz)})
    days.sort(key=lambda d: d['tgl'])
    return days

# ── Export ORDER per-pesanan (mis. Shopee "Order.all") ──────────────────────
# Beda dari laporan Statistik/Analytics (1 baris = 1 hari): file ini 1 baris =
# 1 item pesanan, banyak baris per hari. Jadi omzet dijumlah per-tanggal, baris
# batal/retur dibuang. Kolom dicocokkan via NAMA header (adaptif ke urutan &
# variasi kolom), bukan posisi tetap.
_MP_ORDER_FIELDS = {
    'order':   ('no. pesanan', 'no pesanan', 'nomor pesanan', 'no. order', 'order id', 'order sn', 'order number'),
    'status':  ('status pesanan', 'order status', 'status'),
    'date':    ('waktu pesanan dibuat', 'waktu pembayaran dilakukan', 'waktu pesanan selesai',
                'order creation date', 'order creation time', 'order paid time', 'order paid date',
                'order complete time', 'created time', 'created at', 'paid time', 'tanggal pesanan', 'tanggal'),
    'revenue': ('subtotal pesanan', 'subtotal produk', 'product subtotal', 'order subtotal', 'subtotal',
                'total penjualan (idr)', 'total penjualan'),
    'qty':     ('jumlah', 'quantity', 'qty', 'kuantitas'),
    'price':   ('harga setelah diskon', 'deal price', 'price after discount', 'unit price', 'harga satuan', 'harga'),
}

def _mp_norm(s):
    return ' '.join(str(s).strip().lower().split())

def _mp_find_order_header(ws, max_scan=15):
    """Cari baris header export ORDER + petakan field->indeks kolom (0-based).
    Return (header_row_1based, colmap) kalau ini export order, else None."""
    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), start=1):
        labels = {}
        for ci, cell in enumerate(row):
            if cell is None:
                continue
            lab = _mp_norm(cell)
            if lab and lab not in labels:
                labels[lab] = ci
        if not labels:
            continue
        colmap = {}
        for field, cands in _MP_ORDER_FIELDS.items():
            hit = None
            for cand in cands:                       # exact match dulu (urut prioritas)
                if cand in labels:
                    hit = labels[cand]; break
            if hit is None:                          # fallback: header diawali kandidat
                for cand in cands:
                    for lab, ci in labels.items():
                        if lab.startswith(cand):
                            hit = ci; break
                    if hit is not None:
                        break
            if hit is not None:
                colmap[field] = hit
        # Ciri export order: ada kolom nomor pesanan + tanggal + (omzet | qty&harga)
        if 'order' in colmap and 'date' in colmap and (
                'revenue' in colmap or ('qty' in colmap and 'price' in colmap)):
            return (ri, colmap)
    return None

def _mp_collect_orders(ws, header_row, colmap):
    """Agregasi omzet per-tanggal dari baris pesanan. Baris batal/retur dibuang;
    omzet per-baris = kolom 'Subtotal Pesanan' (sudah per-item, tidak double-count)
    atau qty*harga bila subtotal tak ada."""
    MAX_ROWS = 100000
    ci_status = colmap.get('status'); ci_date = colmap.get('date')
    ci_rev = colmap.get('revenue'); ci_qty = colmap.get('qty'); ci_price = colmap.get('price')
    daily = {}
    n = 0
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        n += 1
        if n > MAX_ROWS:
            break
        if ci_status is not None and ci_status < len(row):
            st = _mp_norm(row[ci_status]) if row[ci_status] is not None else ''
            if st and any(x in st for x in _MP_OMZET_EXCLUDE):
                continue
        iso = _mp_date(row[ci_date]) if (ci_date is not None and ci_date < len(row)) else None
        if not iso:
            continue
        if ci_rev is not None and ci_rev < len(row):
            rev = _mp_num(row[ci_rev])
        elif ci_qty is not None and ci_price is not None and ci_qty < len(row) and ci_price < len(row):
            rev = _mp_num(row[ci_qty]) * _mp_num(row[ci_price])
        else:
            rev = 0.0
        if rev <= 0:
            continue
        daily[iso] = daily.get(iso, 0.0) + rev
    days = [{'tgl': k, 'omzet': round(v)} for k, v in daily.items()]
    days.sort(key=lambda d: d['tgl'])
    return days

def _parse_marketplace_xlsx(f):
    """Parse xlsx marketplace -> dict. Baca byte ke memori (dibatasi), file TIDAK disimpan."""
    MAXB = 15 * 1024 * 1024
    raw = f.read(MAXB + 1)
    if not raw:
        raise ValueError('File kosong.')
    if len(raw) > MAXB:
        raise ValueError('File terlalu besar (maksimal 15 MB).')
    try:
        wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
    except Exception as ex:
        raise ValueError('File Excel tidak bisa dibaca: %s' % ex)
    try:
        if len(wb.sheetnames) > 60:
            raise ValueError('File punya terlalu banyak sheet (tidak wajar untuk laporan marketplace).')
        # 1) Export ORDER per-pesanan (paling spesifik; deteksi via header) — mis. Shopee "Order.all"
        order_ws = order_hdr = None
        for ws in wb.worksheets:
            _oh = _mp_find_order_header(ws)
            if _oh:
                order_ws, order_hdr = ws, _oh
                break
        if order_ws is not None:
            platform = 'SHOPEE'
            sheet_used = order_ws.title
            days = _mp_collect_orders(order_ws, order_hdr[0], order_hdr[1])
            if not days:
                raise ValueError('Tidak ada baris pesanan valid (di luar yang dibatalkan/'
                                 'dikembalikan) yang terbaca dari file export order Shopee.')
        else:
            # 2) Laporan Statistik/Analytics (1 baris = 1 hari)
            platform = _mp_detect(wb)
            if platform == 'SHOPEE':
                target = None
                for pref in ('pesanan dibayar', 'pesanan dibuat', 'pesanan'):
                    for ws in wb.worksheets:
                        if (ws.title or '').strip().lower().startswith(pref):
                            target = ws
                            break
                    if target is not None:
                        break
                if target is None:
                    raise ValueError('Sheet pesanan Shopee tidak ditemukan.')
                sheet_used = target.title
                days = _mp_collect_days(target, ['total penjualan (idr)', 'total penjualan'])
                if not days:
                    raise ValueError("Kolom 'Total Penjualan' tidak ditemukan di sheet '%s'. "
                                     "Pastikan ini file Statistik Toko Shopee yang benar." % sheet_used)
            elif platform == 'TIKTOK':
                ws = wb.worksheets[0]
                sheet_used = ws.title
                days = _mp_collect_days(ws, ['gmv'])
                if not days:
                    raise ValueError("Kolom 'GMV' harian tidak ditemukan. Pastikan ini file "
                                     "Shop Analytics / Key metrics TikTok-Tokopedia.")
            else:
                raise ValueError('Format file tidak dikenali. Didukung: export order Shopee '
                                 '(mis. "Order.all"), Statistik Toko Shopee, atau Shop Analytics / '
                                 'Key metrics TikTok-Tokopedia.')
    finally:
        wb.close()
    days = [d for d in days if d['omzet'] > 0]
    if not days:
        raise ValueError('Tidak ada baris penjualan harian (> 0) yang terbaca dari file.')
    return {
        'platform': platform,
        'platform_label': MP_PLATFORMS.get(platform, platform),
        'sheet': sheet_used,
        'is_dibayar': platform != 'SHOPEE' or 'dibayar' in (sheet_used or '').lower(),
        'days': days,
        'total': round(sum(d['omzet'] for d in days)),
        'n_hari': len(days),
        'start': days[0]['tgl'],
        'end': days[-1]['tgl'],
    }

def _mp_day_sig(platform, tgl_iso):
    """Idempotency per (platform, tanggal) -- bukan per nominal, supaya edit nominal
    tidak menembus deteksi duplikat (1 hari = 1 impor per platform; pakai 'force' utk ulang)."""
    raw = '%s|%s' % (platform, tgl_iso)
    return sha256(raw.encode('utf-8')).hexdigest()[:16]

@app.route('/import-marketplace', methods=['GET', 'POST'])
@sales_input_required
def import_marketplace():
    if not HAS_XLSX:
        flash('Library openpyxl tidak tersedia untuk membaca Excel.', 'danger')
        return redirect(url_for('dashboard'))
    conn = db()
    akun_kas = conn.execute("SELECT * FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    if request.method == 'GET':
        conn.close()
        return render_template('import_marketplace.html', state='form', akun_kas=akun_kas)
    f = request.files.get('file')
    if not f or not f.filename:
        conn.close()
        flash('Pilih file Excel (.xlsx) laporan marketplace dulu.', 'warning')
        return redirect(url_for('import_marketplace'))
    if not f.filename.lower().endswith('.xlsx'):
        conn.close()
        flash('Format harus .xlsx (export Shopee / TikTok-Tokopedia).', 'danger')
        return redirect(url_for('import_marketplace'))
    try:
        parsed = _parse_marketplace_xlsx(f)
    except ValueError as ex:
        conn.close()
        flash(str(ex), 'danger')
        return redirect(url_for('import_marketplace'))
    except Exception:
        conn.close()
        app.logger.exception('import_marketplace: gagal parse file')
        flash('File tidak bisa diproses (format tak terduga). Pastikan file asli export marketplace (.xlsx).', 'danger')
        return redirect(url_for('import_marketplace'))
    dup_dates = set()
    for d in parsed['days']:
        sig = _mp_day_sig(parsed['platform'], d['tgl'])
        if conn.execute("SELECT id FROM jurnal WHERE tx_meta LIKE ? LIMIT 1",
                        ('%"mp_sig": "' + sig + '"%',)).fetchone():
            dup_dates.add(d['tgl'])
    kas_default = '1100' if any(a['kode'] == '1100' for a in akun_kas) else (akun_kas[0]['kode'] if akun_kas else '1100')
    produk_list = conn.execute(
        "SELECT id,nama,stok,harga_beli FROM produk WHERE COALESCE(aktif,1)=1 ORDER BY nama").fetchall()
    conn.close()
    payload = json.dumps({
        'platform': parsed['platform'], 'platform_label': parsed['platform_label'],
        'sheet': parsed['sheet'], 'src': f.filename,
        'start': parsed['start'], 'end': parsed['end'],
    })
    return render_template('import_marketplace.html', state='preview',
                           parsed=parsed, payload=payload, akun_kas=akun_kas, kas_default=kas_default,
                           default_ket='Penjualan ' + parsed['platform_label'], produk_list=produk_list,
                           dup_dates=dup_dates, n_dup=len(dup_dates), src_name=f.filename)

@app.route('/import-marketplace/posting', methods=['POST'])
@sales_input_required
def import_marketplace_posting():
    try:
        meta = json.loads(request.form.get('payload', '') or '{}')
    except (ValueError, TypeError):
        flash('Data import tidak valid. Ulangi upload.', 'danger')
        return redirect(url_for('import_marketplace'))
    platform = meta.get('platform', '')
    if platform not in MP_PLATFORMS:
        flash('Data import tidak valid (platform tak dikenali). Ulangi upload.', 'danger')
        return redirect(url_for('import_marketplace'))
    platform_label = MP_PLATFORMS[platform]
    kas_kode = (request.form.get('akun_kas', '1100') or '1100').strip()
    force = request.form.get('force') == '1'
    try:
        n = int(request.form.get('row_count', '0') or 0)
    except ValueError:
        n = 0
    conn = db()
    if not is_rekening_kode(conn, kas_kode):
        conn.close()
        flash('Rekening kas/bank tidak valid.', 'danger')
        return redirect(url_for('import_marketplace'))
    posted, skipped_dup, total_posted, hpp_info = 0, 0, 0.0, ''
    try:
        for i in range(n):
            if request.form.get('inc_%d' % i) != '1':
                continue
            tgl = (request.form.get('tgl_%d' % i, '') or '').strip()
            ket = (request.form.get('ket_%d' % i, '') or '').strip()[:200] or ('Penjualan ' + platform_label)
            try:
                # pakai parser rupiah app (strip titik ribuan dari numFmt: "53.000" -> 53000)
                omz = round(parse_nonnegative_rp(request.form.get('nom_%d' % i, '0') or '0', 'Omzet'))
            except ValueError:
                continue
            if not is_valid_iso_date(tgl) or omz <= 0:
                continue
            if not _operator_date_ok(tgl):
                continue
            sig = _mp_day_sig(platform, tgl)
            if not force and conn.execute(
                    "SELECT id FROM jurnal WHERE tx_meta LIKE ? LIMIT 1",
                    ('%"mp_sig": "' + sig + '"%',)).fetchone():
                skipped_dup += 1
                continue
            jid = insert_jurnal(conn, tgl, ket, 'OPERASIONAL', 'PEMASUKAN',
                                [(kas_kode, omz, 0), ('4100', 0, omz)], pihak=platform_label)
            conn.execute("UPDATE jurnal SET tx_meta=? WHERE id=?", (json.dumps({
                'tipe': 'JUAL', 'sumber': 'IMPORT_MARKETPLACE', 'platform': platform,
                'akun_kas': kas_kode, 'mp_sig': sig, 'tgl': tgl, 'omzet': omz,
                'src': meta.get('src', ''),
            }), jid))
            posted += 1
            total_posted += omz
        if posted == 0 and skipped_dup == 0:
            conn.rollback(); conn.close()
            flash('Tidak ada baris yang diposting. Centang minimal satu hari dengan nominal > 0.', 'warning')
            return redirect(url_for('import_marketplace'))
        # Opsional: kurangi stok SKU + catat HPP (1 jurnal terpisah Dr 5100 / Cr 1130).
        # Hanya jika ada penjualan baru diposting (posted>0), supaya re-import (semua dup) tak dobel-potong stok.
        if posted > 0:
            sku_pairs = []
            ids = request.form.getlist('prod_id[]')
            qtys = request.form.getlist('prod_qty[]')
            for pid_s, q_s in zip(ids, qtys):
                pid_s = (pid_s or '').strip()
                if not pid_s:
                    continue
                try:
                    pid = int(pid_s)
                    q = parse_positive_qty(q_s, 'Qty')
                except (ValueError, TypeError):
                    continue
                sku_pairs.append((pid, q))
            if sku_pairs:
                hpp_tgl = (request.form.get('hpp_tgl', '') or '').strip()
                if not is_valid_iso_date(hpp_tgl):
                    hpp_tgl = meta.get('end') if is_valid_iso_date(meta.get('end', '')) else tgl
                prepared, total_hpp = [], 0.0
                for pid, q in sku_pairs:
                    pr = conn.execute("SELECT id,nama,stok,harga_beli FROM produk WHERE id=?", (pid,)).fetchone()
                    if not pr:
                        raise ValueError('Produk untuk pengurangan stok tidak ditemukan.')
                    cost = float(pr['harga_beli'] or 0)
                    if cost <= 0:
                        raise ValueError('Produk "%s" belum punya HPP/harga beli. Isi dulu di Inventory sebelum kurangi stok.' % pr['nama'])
                    if float(pr['stok'] or 0) + 1e-9 < q:
                        raise ValueError('Stok "%s" tidak cukup: tersedia %g, diminta %g.' % (pr['nama'], float(pr['stok'] or 0), q))
                    line = round(q * cost)
                    total_hpp += line
                    prepared.append((pr, q, cost))
                if prepared and total_hpp > 0:
                    hpp_ket = 'HPP Penjualan %s' % platform_label
                    hpp_jid = insert_jurnal(conn, hpp_tgl, hpp_ket, 'OPERASIONAL', 'PEMASUKAN',
                                            [('5100', total_hpp, 0), ('1130', 0, total_hpp)], pihak=platform_label)
                    for pr, q, cost in prepared:
                        cur = conn.execute("SELECT stok,harga_beli FROM produk WHERE id=?", (pr['id'],)).fetchone()
                        sb, hb = cur['stok'], cur['harga_beli']
                        conn.execute("UPDATE produk SET stok=stok-? WHERE id=?", (q, pr['id']))
                        record_stock_movement(conn, pr['id'], hpp_tgl, 'KELUAR', q, hb, hpp_ket,
                                              hpp_jid, 'REMOVE_LAYER', sb, hb)
                    conn.execute("UPDATE jurnal SET tx_meta=? WHERE id=?", (json.dumps({
                        'tipe': 'HPP', 'sumber': 'IMPORT_MARKETPLACE_HPP', 'platform': platform,
                        'periode': [meta.get('start', ''), meta.get('end', '')],
                        'total_hpp': total_hpp, 'n_sku': len(prepared),
                    }), hpp_jid))
                    hpp_info = ' Stok dikurangi untuk %d produk, HPP %s.' % (len(prepared), rp_filter(total_hpp))
        add_log(conn, 'Import Penjualan Marketplace',
                '%s | %d hari | %s%s' % (platform_label, posted, rp_filter(total_posted),
                 (' | %d hari dilewati (duplikat)' % skipped_dup) if skipped_dup else ''), 'INPUT')
        conn.commit()
    except ValueError as ex:
        conn.rollback(); conn.close()
        flash('Gagal posting: %s' % ex, 'danger')
        return redirect(url_for('import_marketplace'))
    except Exception as ex:
        conn.rollback(); conn.close()
        flash('Gagal posting: %s' % ex, 'danger')
        return redirect(url_for('import_marketplace'))
    conn.close()
    msg = 'Berhasil import %s: %d jurnal harian, total %s.' % (platform_label, posted, rp_filter(total_posted))
    if skipped_dup:
        msg += ' %d hari dilewati karena sudah pernah diimport.' % skipped_dup
    msg += hpp_info
    flash(msg, 'success')
    return redirect(url_for('daftar_transaksi', tab='semua'))

@app.route('/pemasukan', methods=['GET','POST'])
@sales_input_required
def pemasukan():
    conn = db()
    produk_list = conn.execute("SELECT * FROM produk WHERE COALESCE(aktif,1)=1 ORDER BY nama").fetchall()
    akun_kas = conn.execute("SELECT * FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()

    if request.method == 'POST':
        mode = request.form.get('mode','umum')
        neg_stock_notice = None  # peringatan stok minus (diisi di mode produk)
        value_alert_set = False  # True jika transaksi ini memicu alert nilai persediaan minus
        tanggal = request.form['tanggal']
        keterangan = request.form.get('keterangan','Penjualan')
        akun_kas_kode = request.form.get('akun_kas','1100')
        # Nama pelanggan (untuk tag omzet/profit) + opsi simpan ke database customer
        nama_pelanggan = (request.form.get('pelanggan','') or '').strip()
        simpan_customer = request.form.get('simpan_customer')
        jt = request.form.get('jatuh_tempo') or None

        if not _operator_date_ok(tanggal):
            conn.close()
            flash('Tanggal transaksi tidak valid. Operator hanya bisa input transaksi untuk bulan berjalan.', 'warning')
            return redirect(url_for('pemasukan'))
        if mode not in ('umum', 'produk'):
            return _reject_form(conn, 'Mode pemasukan tidak valid.', 'pemasukan')
        if not is_rekening_kode(conn, akun_kas_kode):
            return _reject_form(conn, 'Rekening kas/bank tidak valid.', 'pemasukan')
        if not is_valid_optional_iso_date(jt):
            return _reject_form(conn, 'Tanggal jatuh tempo tidak valid.', 'pemasukan')
        # Tag proyek (opsional, job costing) — hanya proyek AKTIF yang valid
        proyek_id = request.form.get('proyek_id', type=int) or None
        if proyek_id and not conn.execute(
                "SELECT id FROM proyek WHERE id=? AND status='AKTIF'", (proyek_id,)).fetchone():
            return _reject_form(conn, 'Proyek tidak valid atau sudah tidak aktif.', 'pemasukan')

        if simpan_customer and nama_pelanggan:
            conn.execute("INSERT OR IGNORE INTO customer(nama) VALUES(?)", (nama_pelanggan,))

        if mode == 'umum':
            try:
                diskon  = parse_nonnegative_rp(request.form.get('diskon','0'), 'Diskon')
                ongkir  = parse_nonnegative_rp(request.form.get('ongkir','0'), 'Ongkir')
                biaya_lain = parse_nonnegative_rp(request.form.get('biaya_lain','0'), 'Biaya lain')
                # Biaya/potongan marketplace = BEBAN (mengurangi kas cair), bukan tagihan ke pembeli.
                biaya_mp = parse_nonnegative_rp(request.form.get('biaya_mp','0'), 'Biaya marketplace')
                uang_masuk = parse_nonnegative_rp(request.form.get('uang_masuk','0'), 'Uang masuk')
            except ValueError as ex:
                return _reject_form(conn, str(ex), 'pemasukan')
            # Rincian multi-baris "Penjualan Umum": tiap baris bebas ATAU diambil dari Master HPP.
            # Semua baris non-SKU (tanpa potong stok); nominal = Σ(qty×harga). Struktur jurnal
            # identik dgn versi 1-baris (1 kredit pendapatan utk total), cuma transaksi_item jadi
            # banyak baris — konsumen lain (edit via _build_sale, retur, invoice) sudah baca sbg list.
            um_desc = request.form.getlist('item_deskripsi[]')
            um_qty  = request.form.getlist('item_qty[]')
            um_hrg  = request.form.getlist('item_harga[]')
            um_mst  = request.form.getlist('item_hpp_master[]')
            um_resep = request.form.getlist('item_resep_id[]')
            um_pid  = request.form.getlist('item_produk_id[]')   # baris = produk ber-SKU (potong stok langsung)
            um_dis  = request.form.getlist('item_diskon[]')      # diskon per baris (opsional)
            um_hpp  = request.form.getlist('item_hpp[]')         # HPP/modal per-unit baris BEBAS (SKU/master/resep dihitung server)
            line_items = []
            nominal = 0.0
            manual_hpp = 0.0          # HPP manual per-baris (baris bebas): item_hpp[i] × qty → Persediaan 1130
            jasa_subtotal = 0.0       # subtotal baris jasa (Master HPP jenis JASA) → Pendapatan Jasa 4200
            auto_hpp_total = 0.0      # HPP total dari Master HPP (kalkulator biaya)
            auto_hpp_inventory = 0.0  # porsi Master HPP yang benar-benar mengurangi persediaan
            auto_hpp_reclass = {}     # porsi Master HPP tenaga kerja/overhead/komisi -> reclass beban
            cogs_sku = 0.0        # COGS backflush bahan SKU (potong stok, moving-average)
            cogs_nonsku = 0.0     # COGS backflush bahan non-SKU resep (nilai)
            cogs_sku_direct = 0.0 # COGS produk ber-SKU dijual langsung (baris umum menunjuk produk)
            cogs_reclass = {}      # COGS non-persediaan (tenaga kerja/overhead/komisi) -> Cr akun beban
            backflush_recs = []   # rencana konsumsi bahan (dieksekusi PASCA-jurnal supaya reversible)
            sku_direct_recs = []  # rencana potong stok produk-jadi (dieksekusi PASCA-jurnal, reversible)
            bahan_pids = []       # id bahan SKU / produk-jadi utk peringatan stok minus pasca-commit
            all_jasa = True
            for i in range(len(um_desc)):
                desc  = (um_desc[i] or '').strip()
                q_raw = um_qty[i] if i < len(um_qty) else ''
                h_raw = um_hrg[i] if i < len(um_hrg) else ''
                m_raw = (um_mst[i].strip() if i < len(um_mst) and um_mst[i] else '')
                r_raw = (um_resep[i].strip() if i < len(um_resep) and um_resep[i] else '')
                p_raw = (um_pid[i].strip() if i < len(um_pid) and um_pid[i] else '')
                # Prioritas sumber baris (mutual-exclusive supaya 1130 tak dikredit dobel):
                # Resep (backflush) > Produk ber-SKU (potong stok langsung) > Master HPP > bebas.
                tmpl = None
                if r_raw:
                    try:
                        tmpl = conn.execute("SELECT * FROM produksi_template WHERE id=?", (int(r_raw),)).fetchone()
                    except (ValueError, TypeError):
                        tmpl = None
                    # Resep dipilih tapi tak resolve (id rusak / template dihapus antara GET dan POST):
                    # TOLAK — jangan diam-diam booking pendapatan tanpa COGS & tanpa potong stok.
                    if not tmpl:
                        return _reject_form(conn, f'Resep produksi pada baris {i+1} tidak ditemukan (mungkin sudah dihapus/diubah). Muat ulang halaman lalu pilih ulang resep.', 'pemasukan')
                prod_direct = None
                if not tmpl and p_raw:
                    try:
                        prod_direct = conn.execute(
                            "SELECT * FROM produk WHERE id=? AND COALESCE(aktif,1)=1", (int(p_raw),)).fetchone()
                    except (ValueError, TypeError):
                        prod_direct = None
                    # Produk dipilih tapi tak resolve (id rusak / produk diarsipkan antara GET dan POST):
                    # TOLAK — jangan booking pendapatan tanpa COGS & tanpa potong stok.
                    if not prod_direct:
                        return _reject_form(conn, f'Produk pada baris {i+1} tidak ditemukan atau sudah diarsipkan. Muat ulang halaman lalu pilih ulang produk.', 'pemasukan')
                master = None
                if not tmpl and not prod_direct and m_raw:
                    try:
                        master = conn.execute("SELECT * FROM hpp_produk WHERE id=?", (int(m_raw),)).fetchone()
                    except (ValueError, TypeError):
                        master = None
                # Lewati baris benar-benar kosong (tanpa deskripsi, produk, master, resep, qty, harga)
                if (not desc and not master and not tmpl and not prod_direct
                        and not str(q_raw).strip() and not str(h_raw).strip()):
                    continue
                try:
                    qty   = parse_positive_qty(q_raw or '1', f'Qty baris {i+1}')
                    harga = parse_nonnegative_rp(h_raw or '0', f'Harga baris {i+1}')
                    dis_line = parse_nonnegative_rp(um_dis[i] if i < len(um_dis) else '0', f'Diskon baris {i+1}')
                except (ValueError, TypeError) as ex:
                    return _reject_form(conn, str(ex) if isinstance(ex, ValueError) else f'Qty baris {i+1} tidak valid.', 'pemasukan')
                # Default baris = non-SKU (bebas/Master HPP/resep). Baris produk-jadi override di bawah.
                _li_pid, _li_satuan, _li_arah = None, '', 'NONE'
                if tmpl:
                    qty_hasil = float(tmpl['qty_hasil'] or 0)
                    if qty_hasil <= 0 or not math.isfinite(qty_hasil):
                        return _reject_form(conn, f'Template resep "{tmpl["nama"]}" (baris {i+1}) belum punya jumlah hasil yang valid. Perbaiki template dulu.', 'pemasukan')
                    scale = qty / qty_hasil
                    if not desc:
                        desc = tmpl['nama']
                    # Urutan harga resep: input user -> harga jual resep -> harga produk jadi -> 0.
                    if harga <= 0:
                        harga = float((tmpl['harga_jual'] if 'harga_jual' in tmpl.keys() else 0) or 0)
                    if harga <= 0 and tmpl['produk_jadi_id']:
                        _pj = conn.execute("SELECT harga_jual FROM produk WHERE id=?", (tmpl['produk_jadi_id'],)).fetchone()
                        if _pj:
                            harga = float(_pj['harga_jual'] or 0)
                    for b in json.loads(tmpl['bahan_sku'] or '[]'):
                        try:
                            bpid = int(b.get('produk_id'))
                        except (TypeError, ValueError):
                            continue
                        qb = float(b.get('qty', 0) or 0) * scale
                        if qb <= 0:
                            continue
                        prow = conn.execute("SELECT id,nama,harga_beli FROM produk WHERE id=? AND COALESCE(aktif,1)=1", (bpid,)).fetchone()
                        if not prow:
                            return _reject_form(conn, f'Bahan resep pada baris {i+1} tidak ditemukan atau sudah diarsipkan. Perbaiki template resep.', 'pemasukan')
                        hb = float(prow['harga_beli'] or 0)
                        st = qb * hb
                        cogs_sku += st
                        bahan_pids.append(bpid)
                        backflush_recs.append({'produk_id': bpid, 'deskripsi': prow['nama'],
                                               'menu_qty': qty, 'line_idx': i, 'template_id': tmpl['id'],
                                               'qty': qb, 'harga': hb, 'subtotal': st})
                    for b in json.loads(tmpl['bahan_non_sku'] or '[]'):
                        nvv = float(b.get('nilai', 0) or 0) * scale
                        if nvv <= 0:
                            continue
                        cogs_nonsku += nvv
                        backflush_recs.append({'produk_id': None, 'deskripsi': (b.get('deskripsi', '') or '').strip() or 'Bahan non-SKU',
                                               'menu_qty': qty, 'line_idx': i, 'template_id': tmpl['id'],
                                               'qty': float(b.get('qty', 0) or 0) * scale, 'harga': 0.0, 'subtotal': nvv})
                    biaya_konversi = float(tmpl['biaya_tambahan'] or 0) * scale
                    if biaya_konversi > 0:
                        # Biaya tambahan template = tenaga kerja/overhead. Saat backflush jual,
                        # ini mereklas beban ke HPP, bukan mengurangi Persediaan 1130.
                        _add_amount(cogs_reclass, '6180', biaya_konversi)
                        backflush_recs.append({'produk_id': None, 'deskripsi': 'Biaya tambahan produksi (tenaga kerja/overhead)',
                                               'menu_qty': qty, 'line_idx': i, 'template_id': tmpl['id'],
                                               'qty': qty, 'harga': biaya_konversi / qty if qty else 0.0,
                                               'subtotal': biaya_konversi})
                    all_jasa = False
                elif prod_direct:
                    # Produk ber-SKU dijual langsung dalam satu dokumen penjualan umum.
                    # COGS = qty × harga_beli (moving-average saat ini) → Persediaan 1130,
                    # persis jalur mode "per produk". Potong stok dieksekusi PASCA-jurnal.
                    if not desc:
                        desc = prod_direct['nama']
                    if harga <= 0:
                        harga = float(prod_direct['harga_jual'] or 0)
                    hb = float(prod_direct['harga_beli'] or 0)
                    st = qty * hb
                    cogs_sku_direct += st
                    bahan_pids.append(int(prod_direct['id']))
                    sku_direct_recs.append({'produk_id': int(prod_direct['id']), 'qty': qty})
                    _li_pid = int(prod_direct['id'])
                    _li_satuan = prod_direct['satuan'] or ''
                    _li_arah = 'KELUAR'
                    all_jasa = False
                elif master:
                    if not desc:
                        desc = master['nama']
                    if harga <= 0:
                        harga = float(master['harga_jual'] or 0)
                    _bd = hpp_bahan_breakdown(json.loads(master['bahan'] or '[]'))
                    auto_hpp_total += qty * _bd['total']
                    auto_hpp_inventory += qty * _bd['inventory']
                    for _akun_beban, _nilai in _bd['reclass'].items():
                        _add_amount(auto_hpp_reclass, _akun_beban, qty * _nilai)
                    if (master['jenis'] or '').upper() == 'JASA':
                        # Baris jasa → Pendapatan Jasa (4200). Simpan penanda arah/satuan agar
                        # _build_sale (edit) juga merutekannya ke 4200 secara konsisten.
                        _li_arah, _li_satuan = 'JASA', 'jasa'
                    else:
                        all_jasa = False
                else:
                    # Baris BEBAS (ketik manual): HPP/modal diisi per baris (item_hpp[], per-unit).
                    # SKU/master/resep TIDAK baca field ini (HPP-nya dihitung server) → tak dobel.
                    all_jasa = False
                    try:
                        _hpp_unit = parse_nonnegative_rp(um_hpp[i] if i < len(um_hpp) else '0', f'HPP baris {i+1}')
                    except (ValueError, TypeError) as ex:
                        return _reject_form(conn, str(ex) if isinstance(ex, ValueError) else f'HPP baris {i+1} tidak valid.', 'pemasukan')
                    manual_hpp += _hpp_unit * qty
                if not desc:
                    return _reject_form(conn, f'Deskripsi baris {i+1} wajib diisi.', 'pemasukan')
                gross = qty * harga
                if dis_line > gross + 0.01:
                    return _reject_form(conn, f'Diskon baris {i+1} tidak boleh melebihi subtotal (Rp {gross:,.0f}).', 'pemasukan')
                sub = max(0.0, gross - dis_line)
                nominal += sub
                if _li_arah == 'JASA':
                    jasa_subtotal += sub
                line_items.append({'produk_id': _li_pid, 'deskripsi': desc, 'qty': qty,
                                   'satuan': _li_satuan, 'harga': harga, 'diskon': dis_line,
                                   'subtotal': sub, 'arah': _li_arah})
            if not line_items:
                return _reject_form(conn, 'Tambahkan minimal 1 baris penjualan.', 'pemasukan')
            # HPP per baris, tiap sumber TERISOLASI (tak ada lagi scaling/override global):
            # - baris bebas → manual_hpp (item_hpp/unit × qty) langsung Persediaan 1130.
            # - Master HPP → auto (bahan→1130 inventory; tenaga kerja/overhead/komisi→reclass beban).
            # - SKU/resep → dihitung terpisah di cogs_sku_direct/cogs_sku/cogs_nonsku/cogs_reclass.
            hpp_inventory = manual_hpp + auto_hpp_inventory
            hpp_reclass = {k: v for k, v in auto_hpp_reclass.items() if v > 0}
            hpp = hpp_inventory + sum(hpp_reclass.values())   # total HPP non-SKU (utk tx_meta)
            for _akun_beban, _nilai in hpp_reclass.items():
                _add_amount(cogs_reclass, _akun_beban, _nilai)
            # Bucket COGS:
            # - nonsku_cogs: porsi pool persediaan non-SKU untuk movement/value-alert (TANPA produk
            #   ber-SKU langsung — itu dipantau via pergerakan_stok/negative_stock_notice).
            # - inventory_cogs: seluruh kredit 1130 (manual/Master inventory + bahan SKU resep +
            #   non-SKU resep + produk-jadi ber-SKU dijual langsung).
            # - reclass_cogs: kredit akun beban, bukan persediaan.
            nonsku_cogs = hpp_inventory + cogs_nonsku
            inventory_cogs = hpp_inventory + cogs_sku + cogs_nonsku + cogs_sku_direct
            reclass_cogs = sum(cogs_reclass.values())
            total_cogs = inventory_cogs + reclass_cogs
            hpp_master = None       # tidak ada master tunggal lagi (kompat variabel di bawah)
            hpp_master_id = None
            if not keterangan or keterangan.strip() in ('', 'Penjualan'):
                _nm = [li['deskripsi'] for li in line_items]
                keterangan = 'Penjualan ' + ', '.join(_nm[:3]) + ('+' if len(_nm) > 3 else '')
            if nominal <= 0:
                return _reject_form(conn, 'Nominal penjualan harus lebih dari 0.', 'pemasukan')
            if diskon > nominal:
                return _reject_form(conn, 'Diskon tidak boleh melebihi nominal penjualan.', 'pemasukan')
            total = max(0, nominal - diskon + ongkir + biaya_lain)
            uang_masuk = min(uang_masuk, total)
            piutang_nm = max(0, total - uang_masuk)

            if nonsku_cogs > 0:
                saldo_1130 = persediaan_ledger_value(conn)
                if nonsku_cogs > saldo_1130 + 0.01:
                    # Nilai persediaan non-SKU boleh minus (jual/pakai dulu, barang menyusul).
                    # Tidak diblokir — peringatan tampil sebagai modal di halaman pemasukan.
                    # Porsi SKU backflush dipantau terpisah lewat negative_stock_notice pasca-commit.
                    set_inventory_value_alert(nonsku_cogs, saldo_1130)
                    value_alert_set = True

            if total <= 0:
                conn.close()
                flash('Nominal penjualan harus lebih dari 0.', 'danger')
                return redirect(url_for('pemasukan'))

            # Parse pajak (PRO only, silently 0 jika modul tidak aktif)
            pjk_aktif = is_module_active(conn, 'PAJAK_OTOMATIS')
            try:
                ppn_pct = parse_percentage(request.form.get('ppn_persen', '0'), 'Persen PPN') if pjk_aktif else 0
                pph_pct = parse_percentage(request.form.get('pph23_persen', '0'), 'Persen PPh 23') if pjk_aktif else 0
                pph22_pct = parse_percentage(request.form.get('pph22_persen', '0'), 'Persen PPh 22') if pjk_aktif else 0
            except ValueError as ex:
                return _reject_form(conn, str(ex), 'pemasukan')
            ppn_pct = max(0, min(100, ppn_pct))
            pph_pct = max(0, min(100, pph_pct))
            pph22_pct = max(0, min(100, pph22_pct))
            ppn_nom = round(total * ppn_pct / 100.0, 2)
            pph_nom = round(total * pph_pct / 100.0, 2)
            pph22_nom = round(total * pph22_pct / 100.0, 2)
            # Total kewajiban customer = total + PPN - PPh23 - PPh22 (PPh dipotong customer di sumber).
            gross_kewajiban = total + ppn_nom - pph_nom - pph22_nom
            # Biaya/potongan marketplace = BEBAN yg dipotong dari yg cair ke penjual (mirip PPh):
            # kurangi kewajiban → kas net + tambah Dr 6150 Beban Administrasi. Revenue tetap kotor.
            if biaya_mp > 0:
                if biaya_mp > gross_kewajiban + 0.01:
                    return _reject_form(conn, f'Biaya marketplace (Rp {biaya_mp:,.0f}) tidak boleh melebihi nilai transaksi (Rp {gross_kewajiban:,.0f}).', 'pemasukan')
                if not conn.execute("SELECT 1 FROM akun WHERE kode='6150'").fetchone():
                    return _reject_form(conn, 'Akun 6150 (Beban Administrasi) tidak ada, tidak bisa mencatat biaya marketplace.', 'pemasukan')
            total_kewajiban = gross_kewajiban - biaya_mp
            # Kas/piutang adjustment: alokasi pertahankan rasio uang_masuk vs piutang
            if uang_masuk > 0 or piutang_nm > 0:
                # Scale uang_masuk dari pre-pajak ke post-pajak proporsional
                rasio = uang_masuk / total if total > 0 else 0
                uang_masuk_adj = round(total_kewajiban * rasio, 2)
                piutang_adj    = max(0, total_kewajiban - uang_masuk_adj)
            else:
                uang_masuk_adj, piutang_adj = 0, total_kewajiban

            entries = []
            if uang_masuk_adj > 0:
                entries.append((akun_kas_kode, uang_masuk_adj, 0))
            if piutang_adj > 0:
                entries.append(('1120', piutang_adj, 0))
            if pph_nom > 0:
                entries.append(('1181', pph_nom, 0))   # Dr PPh 23 Dibayar Dimuka
            if pph22_nom > 0:
                entries.append(('1183', pph22_nom, 0)) # Dr PPh 22 Dibayar Dimuka
            if biaya_mp > 0:
                entries.append(('6150', biaya_mp, 0))  # Dr Beban Administrasi (potongan/biaya marketplace)
            # Akun pendapatan: custom COA (semua ke sana) ATAU split otomatis 4100 barang / 4200 jasa.
            akun_pend_custom = request.form.get('akun_pendapatan_kode','').strip()
            if akun_pend_custom:
                _ap = conn.execute("SELECT kode FROM akun WHERE kode=? AND tipe='PENDAPATAN'", (akun_pend_custom,)).fetchone()
                if not _ap:
                    return _reject_form(conn, f'Akun pendapatan COA "{akun_pend_custom}" tidak valid.', 'pemasukan')
                entries.append((akun_pend_custom, 0, total))
            else:
                # Baris jasa (Master HPP jenis JASA) → 4200; sisanya → 4100. Diskon transaksi
                # dialokasikan proporsional barang vs jasa; ongkir & biaya lain masuk barang (4100).
                # Koreksi pembulatan dibebankan ke barang agar barang+jasa == total (jurnal balance).
                jasa_sub  = min(jasa_subtotal, nominal)
                barang_sub = max(0.0, nominal - jasa_sub)
                if nominal > 0 and diskon > 0:
                    diskon_barang = round(diskon * (barang_sub / nominal), 2)
                    diskon_jasa   = diskon - diskon_barang
                else:
                    diskon_barang, diskon_jasa = diskon, 0.0
                pendapatan_jasa   = max(0.0, round(jasa_sub - diskon_jasa, 2))
                pendapatan_barang = max(0.0, round(total - pendapatan_jasa, 2))
                if pendapatan_barang > 0:
                    entries.append(('4100', 0, pendapatan_barang))
                if pendapatan_jasa > 0:
                    entries.append(('4200', 0, pendapatan_jasa))
            if ppn_nom > 0:
                entries.append(('2111', 0, ppn_nom))   # Cr Hutang PPN
            # HPP agregat: non-SKU (manual/Master HPP) + backflush bahan SKU + non-SKU resep.
            # SATU pasang Dr 5100 / Cr 1130 (JANGAN emit 1130 per-bahan) → balance otomatis.
            if total_cogs > 0:
                entries.append(('5100', total_cogs, 0))
                if inventory_cogs > 0:
                    entries.append(('1130', 0, inventory_cogs))
                for _akun_beban, _nilai in sorted(cogs_reclass.items()):
                    entries.append((_akun_beban, 0, _nilai))

            jid = insert_jurnal(conn, tanggal, keterangan, 'OPERASIONAL', 'PEMASUKAN', entries, pihak=nama_pelanggan)
            # Update piutang_nm untuk insert ke tabel piutang dgn nilai adjusted
            piutang_nm = piutang_adj

            if piutang_nm > 0:
                pelanggan = nama_pelanggan or 'Customer'
                conn.execute(
                    "INSERT INTO piutang(tanggal,jatuh_tempo,pelanggan,keterangan,jumlah,jurnal_id) VALUES(?,?,?,?,?,?)",
                    (tanggal, jt, pelanggan, keterangan, piutang_nm, jid)
                )
            # Pergerakan non-SKU: gabung porsi manual/Master HPP + non-SKU resep (JANGAN masukkan
            # cogs_sku — itu di-back oleh pergerakan_stok, bukan pool non-SKU).
            if nonsku_cogs > 0:
                record_non_sku_movement(conn, jid, tanggal, f"HPP non-SKU: {keterangan}", -nonsku_cogs)
            # Rincian item multi-baris (generik/non-SKU) — HANYA baris menu, bahan TIDAK masuk sini.
            _store_items(conn, jid, line_items)
            # Backflush resep: potong stok bahan SKU + catat pergerakan (reversal via pergerakan_stok)
            # + subledger audit penjualan_resep_bahan. Non-SKU sudah masuk record_non_sku_movement di atas.
            for rec in backflush_recs:
                if rec['produk_id']:
                    _cur = conn.execute("SELECT stok FROM produk WHERE id=?", (rec['produk_id'],)).fetchone()
                    _stok_lama = float(_cur['stok'] or 0) if _cur else 0.0
                    conn.execute("UPDATE produk SET stok=stok-? WHERE id=?", (rec['qty'], rec['produk_id']))
                    record_stock_movement(conn, rec['produk_id'], tanggal, 'KELUAR', rec['qty'], rec['harga'],
                                          f"Bahan resep: {keterangan}", jid, 'REMOVE_LAYER', _stok_lama, rec['harga'])
                conn.execute(
                    """INSERT INTO penjualan_resep_bahan(jurnal_id,template_id,line_idx,produk_id,deskripsi,menu_qty,qty,harga,subtotal,qty_returned)
                       VALUES(?,?,?,?,?,?,?,?,?,0)""",
                    (jid, rec['template_id'], rec['line_idx'], rec['produk_id'], rec['deskripsi'],
                     rec['menu_qty'], rec['qty'], rec['harga'], rec['subtotal'])
                )
            # Produk ber-SKU dijual langsung: potong stok + pergerakan (reversal via pergerakan_stok,
            # sama seperti mode "per produk"). harga_beli dibaca ulang (stabil di dalam 1 transaksi),
            # stok dibaca segar tiap baris agar stok_sebelum akurat bila SKU sama muncul >1 baris.
            for rec in sku_direct_recs:
                _cur = conn.execute("SELECT stok,harga_beli FROM produk WHERE id=?", (rec['produk_id'],)).fetchone()
                _stok_lama = float(_cur['stok'] or 0) if _cur else 0.0
                _hb = float(_cur['harga_beli'] or 0) if _cur else 0.0
                conn.execute("UPDATE produk SET stok=stok-? WHERE id=?", (rec['qty'], rec['produk_id']))
                record_stock_movement(conn, rec['produk_id'], tanggal, 'KELUAR', rec['qty'], _hb,
                                      keterangan, jid, 'REMOVE_LAYER', _stok_lama, _hb)
            if bahan_pids:
                _nn = negative_stock_notice(conn, bahan_pids)
                if _nn:
                    neg_stock_notice = _nn
            conn.execute("UPDATE jurnal SET tx_meta=? WHERE id=?", (json.dumps({
                'tipe': 'JUAL', 'akun_kas': akun_kas_kode, 'uang_masuk': uang_masuk,
                'diskon': diskon, 'ongkir': ongkir, 'biaya': biaya_lain, 'biaya_mp': biaya_mp,
                'hpp_generik': round(hpp_inventory, 2),
                'hpp_reclass': {k: round(v, 2) for k, v in hpp_reclass.items() if v > 0},
                'hpp_generik_total': round(hpp, 2),
                'hpp_reclass_all': {k: round(v, 2) for k, v in cogs_reclass.items() if v > 0},
                'hpp_produk_id': hpp_master_id if hpp_master else None,
                'jenis_master_hpp': hpp_master['jenis'] if hpp_master else '',
                'has_resep': bool(backflush_recs),
                'pihak': nama_pelanggan, 'jt': jt,
                # Akun pendapatan COA kustom (bila dipilih) DISIMPAN supaya edit rincian tidak
                # mereklas balik ke 4100/4200 (lihat _build_sale). None = split otomatis 4100/4200.
                'akun_pendapatan_kode': akun_pend_custom or None,
                'ppn_persen': ppn_pct, 'pph23_persen': pph_pct,
                'pph22_persen': pph22_pct}), jid))

        else:  # per produk, multi-item
            produk_ids   = request.form.getlist('produk_id[]')
            qtys         = request.form.getlist('qty[]')
            harga_juals  = request.form.getlist('harga_jual[]')
            diskon_items = request.form.getlist('diskon_item[]')
            try:
                ongkir     = parse_nonnegative_rp(request.form.get('ongkir','0'), 'Ongkir')
                biaya_lain = parse_nonnegative_rp(request.form.get('biaya_lain','0'), 'Biaya lain')
                biaya_jasa = parse_nonnegative_rp(request.form.get('biaya_jasa','0'), 'Biaya jasa')
                uang_masuk = parse_nonnegative_rp(request.form.get('uang_masuk','0'), 'Uang masuk')
            except ValueError as ex:
                return _reject_form(conn, str(ex), 'pemasukan')
            pelanggan  = request.form.get('pelanggan','Customer')
            items = []
            stock_need = {}
            stock_info = {}
            for i, pid_str in enumerate(produk_ids):
                if not pid_str: continue
                try:
                    pid   = int(pid_str)
                    qty   = parse_positive_qty(qtys[i] if i < len(qtys) else 1, f'Qty item baris {i + 1}')
                    harga = parse_nonnegative_rp(harga_juals[i] if i < len(harga_juals) else '0',
                                                 f'Harga jual item baris {i + 1}')
                    dis   = parse_nonnegative_rp(diskon_items[i] if i < len(diskon_items) else '0',
                                                 f'Diskon item baris {i + 1}')
                except (ValueError, TypeError):
                    return _reject_form(conn, f'Rincian produk baris {i + 1} tidak valid.', 'pemasukan')
                produk = conn.execute(
                    "SELECT * FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
                    (pid,)
                ).fetchone()
                if not produk:
                    return _reject_form(conn, f'Produk pada baris {i + 1} tidak ditemukan.', 'pemasukan')
                if dis > qty * harga:
                    return _reject_form(conn, f'Diskon item baris {i + 1} tidak boleh melebihi subtotal bruto.',
                                        'pemasukan')
                stock_need[pid] = stock_need.get(pid, 0.0) + qty
                stock_info[pid] = produk
                subtotal  = qty * harga - dis
                hpp_item  = qty * produk['harga_beli']
                items.append({'produk': produk, 'qty': qty, 'harga': harga,
                              'subtotal': subtotal, 'hpp': hpp_item})

            if not items:
                return _reject_form(conn, 'Pilih minimal 1 produk.', 'pemasukan')
            # Stok fisik minus diizinkan (jual dulu, barang menyusul). Peringatan tampil
            # setelah simpan (flash) dan di halaman inventory — bukan blokir di sini.
            total_sub = sum(it['subtotal'] for it in items)
            total_hpp = sum(it['hpp'] for it in items)
            grand_total = total_sub + ongkir + biaya_lain + biaya_jasa
            pjk_aktif = is_module_active(conn, 'PAJAK_OTOMATIS')
            try:
                ppn_pct = parse_percentage(request.form.get('ppn_persen', '0'), 'Persen PPN') if pjk_aktif else 0
                pph_pct = parse_percentage(request.form.get('pph23_persen', '0'), 'Persen PPh 23') if pjk_aktif else 0
                pph22_pct = parse_percentage(request.form.get('pph22_persen', '0'), 'Persen PPh 22') if pjk_aktif else 0
            except ValueError as ex:
                return _reject_form(conn, str(ex), 'pemasukan')
            ppn_pct = max(0, min(100, ppn_pct))
            pph_pct = max(0, min(100, pph_pct))
            pph22_pct = max(0, min(100, pph22_pct))
            ppn_nom = round(grand_total * ppn_pct / 100.0, 2)
            pph_nom = round(grand_total * pph_pct / 100.0, 2)
            pph22_nom = round(grand_total * pph22_pct / 100.0, 2)
            total_kewajiban = grand_total + ppn_nom - pph_nom - pph22_nom
            uang_masuk_basis = min(uang_masuk, grand_total)
            ratio_bayar = (uang_masuk_basis / grand_total) if grand_total > 0 else 0
            uang_masuk_adj = min(total_kewajiban, round(total_kewajiban * ratio_bayar, 2))
            piutang_nm  = max(0, total_kewajiban - uang_masuk_adj)
            if piutang_nm <= 0.01:
                piutang_nm = 0

            if not keterangan:
                names = [it['produk']['nama'] for it in items]
                keterangan = 'Penjualan ' + ', '.join(names[:3]) + ('+' if len(names) > 3 else '')

            if total_hpp > 0:
                saldo_1130 = persediaan_ledger_value(conn)
                if total_hpp > saldo_1130 + 0.01:
                    # Nilai persediaan boleh minus (jual dulu, barang menyusul).
                    # Tidak diblokir — peringatan tampil sebagai modal di halaman pemasukan.
                    set_inventory_value_alert(total_hpp, saldo_1130)
                    value_alert_set = True

            if grand_total <= 0:
                conn.close()
                flash('Total penjualan harus lebih dari 0.', 'danger')
                return redirect(url_for('pemasukan'))

            entries = []
            if uang_masuk_adj > 0:
                entries.append((akun_kas_kode, uang_masuk_adj, 0))
            if piutang_nm > 0:
                entries.append(('1120', piutang_nm, 0))
            if pph_nom > 0:
                entries.append(('1181', pph_nom, 0))
            if pph22_nom > 0:
                entries.append(('1183', pph22_nom, 0))
            # Pendapatan barang (4100) dipisah dari pendapatan jasa non-SKU (4200)
            entries.append(('4100', 0, grand_total - biaya_jasa))
            if biaya_jasa > 0:
                entries.append(('4200', 0, biaya_jasa))
            if ppn_nom > 0:
                entries.append(('2111', 0, ppn_nom))
            if total_hpp > 0:
                entries += [('5100', total_hpp, 0), ('1130', 0, total_hpp)]

            jid = insert_jurnal(conn, tanggal, keterangan, 'OPERASIONAL', 'PEMASUKAN', entries, pihak=nama_pelanggan)

            _items_store = []
            for item in items:
                current = conn.execute("SELECT stok,harga_beli FROM produk WHERE id=?",
                                       (item['produk']['id'],)).fetchone()
                stok_sebelum = current['stok']
                hpp_sebelum = current['harga_beli']
                conn.execute("UPDATE produk SET stok=stok-? WHERE id=?",
                             (item['qty'], item['produk']['id']))
                record_stock_movement(conn, item['produk']['id'], tanggal, 'KELUAR',
                                      item['qty'], hpp_sebelum, keterangan, jid, 'REMOVE_LAYER',
                                      stok_sebelum, hpp_sebelum)
                _items_store.append({'produk_id': item['produk']['id'], 'deskripsi': item['produk']['nama'],
                                     'qty': item['qty'], 'satuan': item['produk']['satuan'],
                                     'harga': item['harga'], 'diskon': (item['qty']*item['harga'] - item['subtotal']),
                                     'subtotal': item['subtotal'], 'arah': 'KELUAR'})

            if piutang_nm > 0:
                conn.execute(
                    "INSERT INTO piutang(tanggal,jatuh_tempo,pelanggan,keterangan,jumlah,jurnal_id) VALUES(?,?,?,?,?,?)",
                    (tanggal, jt, pelanggan, keterangan, piutang_nm, jid)
                )

            if biaya_jasa > 0:
                _items_store.append({'produk_id': None, 'deskripsi': '(Jasa) ' + (request.form.get('jasa_ket','').strip() or 'Biaya Jasa'),
                                     'qty': 1, 'satuan': 'jasa', 'harga': biaya_jasa,
                                     'diskon': 0, 'subtotal': biaya_jasa, 'arah': 'JASA'})
            _store_items(conn, jid, _items_store)
            conn.execute("UPDATE jurnal SET tx_meta=? WHERE id=?", (json.dumps({
                'tipe': 'JUAL', 'akun_kas': akun_kas_kode, 'uang_masuk': uang_masuk,
                'diskon': 0, 'ongkir': ongkir, 'biaya': biaya_lain, 'jasa': biaya_jasa, 'hpp_generik': 0,
                'pihak': nama_pelanggan, 'jt': jt,
                'ppn_persen': ppn_pct, 'pph23_persen': pph_pct,
                'pph22_persen': pph22_pct}), jid))
            neg_stock_notice = negative_stock_notice(conn, stock_info.keys())

        buat_invoice = request.form.get('buat_invoice') == 'on'
        inv_id = None
        if buat_invoice:
            try:
                inv_id = _ensure_transaction_invoice(conn, jid, {
                    'nomor': request.form.get('inv_nomor', ''),
                    'pelanggan': request.form.get('inv_pelanggan', ''),
                    'telepon': request.form.get('inv_telepon_pelanggan', ''),
                    'alamat': request.form.get('inv_alamat_pelanggan', ''),
                    'jatuh_tempo': request.form.get('inv_jatuh_tempo', ''),
                    'catatan': request.form.get('inv_catatan_custom', ''),
                })
            except ValueError as ex:
                conn.rollback(); conn.close()
                flash(str(ex), 'danger')
                return redirect(url_for('pemasukan'))

        if proyek_id:
            conn.execute("UPDATE jurnal SET proyek_id=? WHERE id=?", (proyek_id, jid))
        add_log(conn, 'Pemasukan dicatat', f"{keterangan} | {tanggal}", 'INPUT')
        saldo_persediaan_setelah = persediaan_ledger_value(conn)
        if saldo_persediaan_setelah >= -0.01:
            session.pop('inventory_value_alert', None)
        elif value_alert_set:
            # Lampirkan jurnal_id transaksi ini agar modal bisa menawarkan "Batalkan Transaksi".
            _alert = dict(session.get('inventory_value_alert') or {})
            _alert['jid'] = jid
            session['inventory_value_alert'] = _alert
        conn.commit(); conn.close()
        flash('Penjualan berhasil dicatat!', 'success')
        if neg_stock_notice:
            flash(neg_stock_notice, 'warning')
        if inv_id:
            return redirect(url_for('pemasukan', new_inv=inv_id))
        return redirect(url_for('pemasukan'))

    saldo_persediaan = persediaan_ledger_value(conn)
    inventory_value_alert = session.get('inventory_value_alert')
    if inventory_value_alert and saldo_persediaan < -0.01:
        # Modal hanya muncul jika ADA alert dari transaksi pemicu (PRG setelah POST).
        # Persediaan minus tanpa alert = state "jual dulu" yang diizinkan → jangan paksa modal.
        inventory_value_alert = dict(inventory_value_alert)
        inventory_value_alert['resulting'] = round(float(saldo_persediaan), 2)
        inventory_value_alert['shortage'] = round(abs(float(saldo_persediaan)), 2)
        inventory_value_alert.setdefault('available', 0)
        inventory_value_alert.setdefault('required', inventory_value_alert['shortage'])
        session['inventory_value_alert'] = inventory_value_alert
    else:
        # Tidak ada alert, atau persediaan sudah >= 0 → bersihkan & jangan tampilkan modal.
        if inventory_value_alert:
            session.pop('inventory_value_alert', None)
        inventory_value_alert = None
    customer_names = [r['nama'] for r in conn.execute("SELECT nama FROM customer ORDER BY nama").fetchall()]
    hpp_rows = conn.execute(
        "SELECT id,nama,jenis,satuan,harga_jual,bahan FROM hpp_produk ORDER BY updated_at DESC,nama"
    ).fetchall()
    hpp_master = []
    for row in hpp_rows:
        bahan = json.loads(row['bahan'] or '[]')
        _bd = hpp_bahan_breakdown(bahan)
        hpp_master.append({
            'id': row['id'], 'nama': row['nama'], 'jenis': row['jenis'] or 'FNB_MENU',
            'satuan': row['satuan'] or 'porsi', 'harga_jual': row['harga_jual'] or 0,
            'total_hpp': _bd['total'],
            'inventory_hpp': _bd['inventory'],
            'reclass_hpp': round(sum(_bd['reclass'].values()), 2),
            'reclass_accounts': _bd['reclass'],
        })
    # Resep Produksi (backflush saat jual): potong stok bahan + COGS moving-average.
    # hpp_per_unit dihitung dari harga_beli bahan SAAT INI (preview JS); server tetap re-snapshot saat jual.
    resep_templates = []
    for t in conn.execute(
        "SELECT id,nama,produk_jadi_id,qty_hasil,biaya_tambahan,harga_jual,bahan_sku,bahan_non_sku FROM produksi_template ORDER BY nama"
    ).fetchall():
        qh = float(t['qty_hasil'] or 0)
        cost_inventory = 0.0
        for b in (json.loads(t['bahan_sku'] or '[]') or []):
            _pr = conn.execute("SELECT harga_beli FROM produk WHERE id=?", (b.get('produk_id'),)).fetchone()
            cost_inventory += float(b.get('qty', 0) or 0) * float((_pr['harga_beli'] if _pr else 0) or 0)
        for b in (json.loads(t['bahan_non_sku'] or '[]') or []):
            cost_inventory += float(b.get('nilai', 0) or 0)
        cost_reclass = float(t['biaya_tambahan'] or 0)
        cost = cost_inventory + cost_reclass
        _pj = conn.execute("SELECT harga_jual FROM produk WHERE id=?", (t['produk_jadi_id'],)).fetchone()
        resep_templates.append({
            'id': t['id'], 'nama': t['nama'], 'qty_hasil': qh,
            'hpp_per_unit': round(cost / qh, 2) if qh > 0 else 0,
            'inventory_hpp_per_unit': round(cost_inventory / qh, 2) if qh > 0 else 0,
            'reclass_hpp_per_unit': round(cost_reclass / qh, 2) if qh > 0 else 0,
            # Harga jual resep sendiri diutamakan; kalau belum diisi pakai harga produk jadi.
            'harga_jual': float((t['harga_jual'] if 'harga_jual' in t.keys() else 0) or 0)
                          or float((_pj['harga_jual'] if _pj else 0) or 0),
        })
    akun_pendapatan_list = conn.execute(
        "SELECT kode, nama FROM akun WHERE tipe='PENDAPATAN' AND saldo_normal='KREDIT' ORDER BY kode"
    ).fetchall()
    proyek_aktif = conn.execute(
        "SELECT id, kode, nama FROM proyek WHERE status='AKTIF' ORDER BY nama"
    ).fetchall()
    conn.close()
    return render_template('pemasukan.html', produk_list=produk_list, akun_kas=akun_kas,
                           today=date.today().strftime('%Y-%m-%d'),
                           saldo_persediaan=saldo_persediaan, customer_names=customer_names,
                           hpp_master=hpp_master, resep_templates=resep_templates,
                           akun_pendapatan_list=akun_pendapatan_list,
                           proyek_aktif=proyek_aktif,
                           inventory_value_alert=inventory_value_alert)

@app.route('/api/produk/<int:pid>')
@operator_required
def api_produk(pid):
    conn = db()
    p = conn.execute(
        "SELECT * FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
        (pid,)
    ).fetchone()
    conn.close()
    if not p:
        return jsonify({})
    harga_beli = p['harga_beli'] if (has_permission('show_hpp_margin') or has_permission('inventory_hpp')) else 0
    return jsonify({'harga_beli': harga_beli, 'harga_jual': p['harga_jual'],
                    'stok': p['stok'], 'satuan': p['satuan']})


# ---------- PENDANAAN (Funding: Setoran Modal / Pinjaman) ----------
@app.route('/pendanaan', methods=['GET','POST'])
@finance_required
def pendanaan():
    conn = db()
    akun_kas = conn.execute("SELECT * FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    akun_hutang = conn.execute(
        "SELECT * FROM akun WHERE tipe='LIABILITAS' ORDER BY kode"
    ).fetchall()

    if request.method == 'POST':
        sumber        = request.form.get('sumber', 'modal')
        tanggal       = request.form.get('tanggal', '')
        keterangan    = (request.form.get('keterangan', '') or '').strip()
        akun_kas_kode = request.form.get('akun_kas', '1100')

        if sumber not in ('modal', 'hutang'):
            return _reject_form(conn, 'Sumber pendanaan tidak valid.', 'pendanaan')
        if not is_valid_iso_date(tanggal):
            return _reject_form(conn, 'Tanggal pendanaan tidak valid.', 'pendanaan')
        if not is_rekening_kode(conn, akun_kas_kode):
            return _reject_form(conn, 'Rekening kas/bank tujuan tidak valid.', 'pendanaan')
        try:
            nominal = parse_nonnegative_rp(request.form.get('nominal', '0'), 'Nominal pendanaan')
        except ValueError as ex:
            return _reject_form(conn, str(ex), 'pendanaan')
        if nominal <= 0:
            return _reject_form(conn, 'Nominal pendanaan harus lebih dari 0.', 'pendanaan')

        if sumber == 'modal':
            # Setoran pemilik: Dr Kas/Bank, Cr Modal Pemilik (3100)
            ket = keterangan or 'Setoran Modal Pemilik'
            jid = insert_jurnal(conn, tanggal, ket, 'PENDANAAN', 'PENDANAAN', [
                (akun_kas_kode, nominal, 0), ('3100', 0, nominal)
            ])
            add_log(conn, 'Pendanaan: setoran modal', f"Rp {nominal:,.0f} | {tanggal}", 'MODAL')
            conn.commit(); conn.close()
            flash('Setoran modal pemilik dicatat. Kas & Modal Pemilik bertambah.', 'success')
            return redirect(url_for('pendanaan'))

        # sumber == 'hutang' → Pinjaman: Dr Kas/Bank, Cr Liabilitas + tracker hutang
        akun_liab = request.form.get('akun_hutang', '2200')
        if not conn.execute("SELECT 1 FROM akun WHERE kode=? AND tipe='LIABILITAS'", (akun_liab,)).fetchone():
            return _reject_form(conn, 'Akun liabilitas (sumber pinjaman) tidak valid.', 'pendanaan')
        kreditur = (request.form.get('kreditur', '') or '').strip() or 'Kreditur'
        jt  = request.form.get('jatuh_tempo') or None
        if not is_valid_optional_iso_date(jt):
            return _reject_form(conn, 'Tanggal jatuh tempo tidak valid.', 'pendanaan')
        ket = keterangan or f'Pinjaman dari {kreditur}'
        jid = insert_jurnal(conn, tanggal, ket, 'PENDANAAN', 'PENDANAAN', [
            (akun_kas_kode, nominal, 0), (akun_liab, 0, nominal)
        ], pihak=kreditur)
        conn.execute(
            "INSERT INTO hutang(tanggal,jatuh_tempo,pemasok,keterangan,jumlah,jurnal_id,akun_kode) VALUES(?,?,?,?,?,?,?)",
            (tanggal, jt, kreditur, ket, nominal, jid, akun_liab)
        )
        add_log(conn, 'Pendanaan: pinjaman', f"Rp {nominal:,.0f} | {kreditur} | {tanggal}", 'INPUT')
        conn.commit(); conn.close()
        flash('Pinjaman dicatat. Kas bertambah & hutang masuk ke tracker untuk dilunasi nanti.', 'success')
        return redirect(url_for('pendanaan'))

    # GET, riwayat pendanaan via menu ini
    page = _page_arg()
    n_rows = conn.execute("""SELECT COUNT(DISTINCT j.id) FROM jurnal j
        JOIN detail_jurnal d ON d.jurnal_id=j.id JOIN akun a ON a.id=d.akun_id
        WHERE j.tipe_tx='PENDANAAN' AND a.is_rekening=1 AND d.debit>0""").fetchone()[0]
    pg = paginate_meta(page, n_rows)
    riwayat = conn.execute("""
        SELECT j.id, j.tanggal, j.keterangan, j.pihak,
               COALESCE(SUM(d.debit),0) AS nominal,
               CASE WHEN EXISTS(
                   SELECT 1 FROM detail_jurnal d2 JOIN akun a2 ON a2.id=d2.akun_id
                   WHERE d2.jurnal_id=j.id AND a2.kode='3100' AND d2.kredit>0
               ) THEN 'modal' ELSE 'hutang' END AS sumber
        FROM jurnal j
        JOIN detail_jurnal d ON d.jurnal_id=j.id
        JOIN akun a ON a.id=d.akun_id
        WHERE j.tipe_tx='PENDANAAN' AND a.is_rekening=1 AND d.debit>0
        GROUP BY j.id ORDER BY j.tanggal DESC, j.id DESC LIMIT ? OFFSET ?
    """, (pg['per_page'], pg['offset'])).fetchall()
    today = date.today().strftime('%Y-%m-%d')
    conn.close()
    return render_template('pendanaan.html', akun_kas=akun_kas, akun_hutang=akun_hutang,
                           riwayat=riwayat, today=today, pg=pg)


@app.route('/pendanaan/<int:id>/hapus', methods=['POST'])
@finance_required
def pendanaan_hapus(id):
    conn = db()
    j = conn.execute("SELECT id FROM jurnal WHERE id=? AND tipe_tx='PENDANAAN'", (id,)).fetchone()
    if not j:
        conn.close(); flash('Data pendanaan tidak ditemukan.', 'danger')
        return redirect(url_for('pendanaan'))
    # Guard integritas: pendanaan tipe pinjaman membuat baris hutang. Bila hutang itu sudah
    # dicicil (bayar_hutang), menghapus pendanaan akan mereverse hutangnya dan FK CASCADE
    # menghapus bukti pembayaran, meninggalkan jurnal kas pembayaran menggantung tanpa hutang
    # asal. Wajibkan hapus cicilan dulu — pola sama dengan transaksi_hapus/hutang_hapus.
    if transaction_dependency_count(conn, id):
        conn.close()
        flash('Pendanaan ini sudah punya pembayaran/pelunasan terkait. Hapus dulu transaksi pembayaran hutangnya sebelum menghapus data pendanaan.', 'warning')
        return redirect(url_for('pendanaan'))
    _reverse_tx(conn, id)
    conn.execute('DELETE FROM jurnal WHERE id=?', (id,))
    add_log(conn, 'Hapus pendanaan', f"Jurnal #{id}", 'MODAL')
    conn.commit(); conn.close()
    flash('Data pendanaan dihapus & efek kas/hutang dibatalkan.', 'warning')
    return redirect(url_for('pendanaan'))


# ---------- TRANSFER REKENING ----------
@app.route('/transfer-rekening', methods=['GET', 'POST'])
@operator_required
def transfer_rekening():
    conn = db()
    rekening = get_rekening_saldo(conn)

    if request.method == 'POST':
        tanggal = request.form.get('tanggal', str(date.today()))
        akun_asal = (request.form.get('akun_asal', '') or '').strip()
        akun_tujuan = (request.form.get('akun_tujuan', '') or '').strip()
        keterangan = (request.form.get('keterangan', '') or '').strip()
        try:
            nominal = parse_nonnegative_rp(request.form.get('nominal', '0'), 'Nominal transfer')
            biaya_admin = parse_nonnegative_rp(request.form.get('biaya_admin', '0'), 'Biaya admin')
        except ValueError as ex:
            return _reject_form(conn, str(ex), 'transfer_rekening')

        if not _operator_date_ok(tanggal):
            conn.close()
            flash('Tanggal transfer tidak valid. Operator hanya bisa input transaksi untuk bulan berjalan.', 'warning')
            return redirect(url_for('transfer_rekening'))
        if nominal <= 0:
            return _reject_form(conn, 'Nominal transfer harus lebih dari 0.', 'transfer_rekening')
        if not is_rekening_kode(conn, akun_asal):
            return _reject_form(conn, 'Rekening asal tidak valid.', 'transfer_rekening')
        if not is_rekening_kode(conn, akun_tujuan):
            return _reject_form(conn, 'Rekening tujuan tidak valid.', 'transfer_rekening')
        if akun_asal == akun_tujuan:
            return _reject_form(conn, 'Rekening asal dan tujuan tidak boleh sama.', 'transfer_rekening')

        asal = conn.execute("SELECT kode,nama FROM akun WHERE kode=?", (akun_asal,)).fetchone()
        tujuan = conn.execute("SELECT kode,nama FROM akun WHERE kode=?", (akun_tujuan,)).fetchone()
        ket = keterangan or f"Transfer rekening: {asal['nama']} ke {tujuan['nama']}"
        entries = [(akun_tujuan, nominal, 0)]
        if biaya_admin > 0:
            entries.append(('6150', biaya_admin, 0))
        entries.append((akun_asal, 0, nominal + biaya_admin))

        jid = insert_jurnal(conn, tanggal, ket, 'OPERASIONAL', 'TRANSFER', entries)
        conn.execute("UPDATE jurnal SET tx_meta=? WHERE id=?", (json.dumps({
            'tipe': 'TRANSFER_REKENING',
            'akun_asal': akun_asal,
            'akun_tujuan': akun_tujuan,
            'nominal': nominal,
            'biaya_admin': biaya_admin,
        }), jid))
        add_log(conn, 'Transfer rekening',
                f"{akun_asal} -> {akun_tujuan} | Rp {nominal:,.0f} | biaya Rp {biaya_admin:,.0f}",
                'INPUT')
        conn.commit(); conn.close()
        flash('Transfer rekening berhasil dicatat.', 'success')
        return redirect(url_for('transfer_rekening'))

    page = _page_arg()
    n_rows = conn.execute("SELECT COUNT(*) FROM jurnal WHERE tipe_tx='TRANSFER'").fetchone()[0]
    pg = paginate_meta(page, n_rows)
    riwayat = conn.execute("""
        SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan,
               (SELECT a.kode || ' - ' || a.nama
                FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id
                WHERE d.jurnal_id=j.id AND a.is_rekening=1 AND d.kredit>0
                ORDER BY d.kredit DESC LIMIT 1) AS rekening_asal,
               (SELECT a.kode || ' - ' || a.nama
                FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id
                WHERE d.jurnal_id=j.id AND a.is_rekening=1 AND d.debit>0
                ORDER BY d.debit DESC LIMIT 1) AS rekening_tujuan,
               (SELECT COALESCE(SUM(d.debit),0)
                FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id
                WHERE d.jurnal_id=j.id AND a.is_rekening=1) AS nominal,
               (SELECT COALESCE(SUM(d.debit),0)
                FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id
                WHERE d.jurnal_id=j.id AND a.is_rekening=0) AS biaya_admin,
               (SELECT COALESCE(SUM(d.kredit),0)
                FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id
                WHERE d.jurnal_id=j.id AND a.is_rekening=1) AS total_keluar
        FROM jurnal j
        WHERE j.tipe_tx='TRANSFER'
        ORDER BY j.tanggal DESC, j.id DESC
        LIMIT ? OFFSET ?
    """, (pg['per_page'], pg['offset'])).fetchall()
    today = date.today().strftime('%Y-%m-%d')
    conn.close()
    return render_template('transfer_rekening.html',
                           rekening=rekening, riwayat=riwayat, today=today, pg=pg)


@app.route('/transfer-rekening/<int:id>/hapus', methods=['POST'])
@finance_required
def transfer_rekening_hapus(id):
    conn = db()
    j = conn.execute("SELECT id FROM jurnal WHERE id=? AND tipe_tx='TRANSFER'", (id,)).fetchone()
    if not j:
        conn.close()
        flash('Data transfer rekening tidak ditemukan.', 'danger')
        return redirect(url_for('transfer_rekening'))
    _reverse_tx(conn, id)
    conn.execute("DELETE FROM jurnal WHERE id=?", (id,))
    add_log(conn, 'Hapus transfer rekening', f"Jurnal #{id}", 'HAPUS')
    conn.commit(); conn.close()
    flash('Transfer rekening dihapus.', 'warning')
    return redirect(url_for('transfer_rekening'))


# ---------- PENGELUARAN ----------
OPERASIONAL_MAP = {
    'Gaji':'6100','Sewa':'6110','Utilitas':'6120','Pemasaran':'6140',
    'Administrasi':'6150','Bunga':'6160','Lainnya':'6180'
}
PENGELUARAN_KATEGORI = {'OPERASIONAL', 'PAJAK', 'PENARIKAN_OWNER', 'PINJAMAN_KARYAWAN', 'BAHAN_BAKU', 'INVESTASI_ASET'}

def _proses_pengeluaran(conn, ep):
    """Proses POST input Pengeluaran & Pembelian (handler bersama, logika akuntansi identik).
    ep = endpoint halaman asal ('pengeluaran' / 'pembelian') agar redirect error kembali
    ke form asal. Blok `if True:` disengaja: mempertahankan indentasi body asli 1:1
    saat diekstrak dari route pengeluaran() supaya diff logika = nol."""
    if True:
        kategori   = request.form['kategori']
        tanggal    = request.form['tanggal']
        keterangan = request.form.get('keterangan','')
        kas_mode   = request.form.get('kas_mode','single')

        if not _operator_date_ok(tanggal):
            conn.close()
            flash('Tanggal transaksi tidak valid. Operator hanya bisa input transaksi untuk bulan berjalan.', 'warning')
            return redirect(url_for(ep))
        if kategori not in PENGELUARAN_KATEGORI:
            return _reject_form(conn, 'Kategori pengeluaran tidak valid.', ep)
        if kas_mode not in ('single', 'multi'):
            return _reject_form(conn, 'Mode pembayaran tidak valid.', ep)
        try:
            nominal = parse_nonnegative_rp(request.form.get('nominal','0'), 'Nominal pengeluaran')
        except ValueError as ex:
            return _reject_form(conn, str(ex), ep)
        if nominal <= 0:
            return _reject_form(conn, 'Nominal pengeluaran harus lebih dari 0.', ep)
        # Tag proyek (opsional, job costing) — hanya proyek AKTIF yang valid
        proyek_id = request.form.get('proyek_id', type=int) or None
        if proyek_id and not conn.execute(
                "SELECT id FROM proyek WHERE id=? AND status='AKTIF'", (proyek_id,)).fetchone():
            return _reject_form(conn, 'Proyek tidak valid atau sudah tidak aktif.', ep)
        # Akun hutang: default mengikuti kategori, tapi user boleh memilih akun
        # kewajiban lain (mis. memisahkan hutang per pemasok atau per jenis).
        # Divalidasi ke COA supaya kode ngawur dari form tidak bisa membuat jurnal
        # nyasar; kalau tidak valid, diam-diam kembali ke default kategori.
        # Nilai ini SATU-SATUNYA sumber: mengalir ke jurnal, kolom hutang.akun_kode,
        # dan tx_meta['akun_hutang'] (dipakai _build_purchase saat transaksi diedit),
        # jadi pelunasan & retur otomatis memakai akun yang sama.
        akun_hutang_kode = {
            'PAJAK': '2110',
            'PENARIKAN_OWNER': '2140',
        }.get(kategori, '2100')
        _ah_pilih = (request.form.get('akun_hutang', '') or '').strip()
        if _ah_pilih and _ah_pilih != akun_hutang_kode:
            if conn.execute("SELECT 1 FROM akun WHERE kode=? AND tipe='LIABILITAS'",
                            (_ah_pilih,)).fetchone():
                akun_hutang_kode = _ah_pilih
            else:
                return _reject_form(conn, 'Akun hutang yang dipilih tidak valid.', ep)

        # ── Vendor (tag belanja) + opsi simpan ke database vendor ────────────────
        nama_vendor = (request.form.get('vendor','') or '').strip()
        simpan_vendor = request.form.get('simpan_vendor')
        if simpan_vendor and nama_vendor:
            conn.execute("INSERT OR IGNORE INTO vendor(nama) VALUES(?)", (nama_vendor,))

        # ── Hutang fields (semua kategori sekarang support hutang) ───────────────
        # pemasok default ke nama vendor bila kosong (untuk tracker hutang)
        pemasok_raw = (request.form.get('pemasok','') or '').strip() or nama_vendor
        jt          = request.form.get('jatuh_tempo') or None
        if not is_valid_optional_iso_date(jt):
            return _reject_form(conn, 'Tanggal jatuh tempo tidak valid.', ep)

        # ── Build kas credit entries + tentukan uang_keluar ──────────────────────
        if kas_mode == 'multi':
            kas_kodens = request.form.getlist('akun_kas_multi[]')
            kas_noms   = request.form.getlist('nominal_kas_multi[]')
            if len(kas_kodens) != len(kas_noms):
                return _reject_form(conn, 'Rincian multi-rekening tidak lengkap.', ep)
            kas_splits = []
            for i, (kode, raw_nom) in enumerate(zip(kas_kodens, kas_noms), start=1):
                if not kode and not str(raw_nom or '').strip():
                    continue
                if not is_rekening_kode(conn, kode):
                    return _reject_form(conn, f'Rekening pembayaran baris {i} tidak valid.', ep)
                try:
                    nom = parse_nonnegative_rp(raw_nom, f'Nominal rekening baris {i}')
                except ValueError as ex:
                    return _reject_form(conn, str(ex), ep)
                if nom <= 0:
                    return _reject_form(conn, f'Nominal rekening baris {i} harus lebih dari 0.', ep)
                kas_splits.append((kode, nom))
            total_kas  = sum(n for _, n in kas_splits)
            # Validasi total_kas vs nominal dilakukan setelah pajak diketahui (di bawah).
            uang_keluar = total_kas
            def kas_credits(amount):
                return [(kode, 0, nom) for kode, nom in kas_splits]
        else:
            akun_kas_kode = request.form.get('akun_kas','1100')
            if not is_rekening_kode(conn, akun_kas_kode):
                return _reject_form(conn, 'Rekening kas/bank tidak valid.', ep)
            raw_uk = request.form.get('uang_keluar')
            if raw_uk is None:
                # Field tidak ada di form (form versi lama) → default full cash semua kategori
                uang_keluar = nominal
            else:
                # Field ada (mungkin kosong) → kosong berarti user memang ingin uang_keluar=0
                try:
                    uang_keluar = parse_nonnegative_rp(raw_uk, 'Uang keluar')
                except ValueError as ex:
                    return _reject_form(conn, str(ex), ep)
            # Validasi uang_keluar, dilakukan SETELAH pajak diketahui (lihat blok pajak di bawah).
            def kas_credits(amount):
                return [(akun_kas_kode, 0, amount)] if amount > 0 else []

        # ── Pajak PPN/PPh (PRO) ──
        pjk_aktif = is_module_active(conn, 'PAJAK_OTOMATIS')
        try:
            x_ppn_pct = parse_percentage(request.form.get('ppn_persen', '0'), 'Persen PPN') if pjk_aktif else 0
            x_pph_pct = parse_percentage(request.form.get('pph23_persen', '0'), 'Persen PPh 23') if pjk_aktif else 0
        except ValueError as ex:
            return _reject_form(conn, str(ex), ep)
        x_ppn_pct = max(0, min(100, x_ppn_pct))
        x_pph_pct = max(0, min(100, x_pph_pct))
        x_ppn_nom = round(nominal * x_ppn_pct / 100.0, 2)
        x_pph_nom = round(nominal * x_pph_pct / 100.0, 2)
        # ── Mode rupiah bulat: bulatkan nominal ke rupiah utuh supaya sisa pecahan sub-rupiah
        # (mis. qty desimal 0,5 x harga = ,50) tidak bocor jadi Hutang Usaha / kas hantu.
        # Saat pref_decimal ON, biarkan pecahan karena bisnisnya memang pakai sen.
        # uang_keluar hanya dibulatkan di mode single-kas: di multi-kas kredit kas diambil dari
        # kas_splits mentah, jadi membulatkan totalnya bisa merusak keseimbangan jurnal.
        if get_setting(conn, 'pref_decimal', '0') != '1':
            nominal      = _round_rp(nominal)
            x_ppn_nom    = _round_rp(x_ppn_nom)
            x_pph_nom    = _round_rp(x_pph_nom)
            if kas_mode != 'multi':
                uang_keluar = _round_rp(uang_keluar)
        # Total kewajiban kita ke supplier = DPP + PPN - PPh23 (PPh23 dipotong = kita bayar lebih sedikit)
        total_kewajiban_pjk = nominal + x_ppn_nom - x_pph_nom
        if uang_keluar > total_kewajiban_pjk + 0.5:
            return _reject_form(conn, f'Uang keluar (Rp {uang_keluar:,.0f}) melebihi total kewajiban setelah pajak (Rp {total_kewajiban_pjk:,.0f}).', ep)
        hutang_nm = max(0, total_kewajiban_pjk - uang_keluar)

        def with_tax(entries):
            """Inject Dr PPN Masukan + Cr Hutang PPh 23 ke entries list."""
            out = list(entries)
            if x_ppn_nom > 0:
                out.append(('1180', x_ppn_nom, 0))
            if x_pph_nom > 0:
                out.append(('2112', 0, x_pph_nom))
            return out

        jid = None                 # jurnal id transaksi ini
        _akun_debit = None         # akun debit utama (untuk recompute edit)
        _peng_items = []           # rincian item
        _peng_editable = True      # apakah dapat tombol Edit Item (recompute)
        _single_kas = akun_kas_kode if kas_mode != 'multi' else '1100'
        _kas_splits_meta = [
            {'kode': kode, 'nominal': nom} for kode, nom in (kas_splits if kas_mode == 'multi' else [])
        ]

        if kategori == 'OPERASIONAL':
            sub = request.form.get('subkategori','Lainnya')
            # Akun beban override: COA kustom user menggantikan OPERASIONAL_MAP
            akun_beban_override = request.form.get('akun_beban_kode','').strip()
            if akun_beban_override:
                _ab = conn.execute("SELECT kode, nama FROM akun WHERE kode=? AND tipe='BEBAN'", (akun_beban_override,)).fetchone()
                if not _ab:
                    return _reject_form(conn, f'Akun beban COA "{akun_beban_override}" tidak valid.', ep)
                beban_kode = akun_beban_override
                keterangan = keterangan or f"Beban {_ab['nama']}"
            elif sub == 'Bunga':
                # Bunga punya PERAN sendiri di Laba Rugi (baris Beban Bunga terpisah),
                # jadi akunnya dicari lewat peran, bukan nomor 6160 di OPERASIONAL_MAP.
                # Aturannya disamakan dengan Pajak: kalau perannya belum ada, tolak dan
                # arahkan ke Daftar Akun. Menebak nomor akan mendaratkan bunga ke akun
                # yang mungkin sudah dipakai ulang untuk hal lain.
                beban_kode = _akun_beban_berperan(conn, 'BUNGA')
                if not beban_kode:
                    return _reject_form(conn, _pesan_tanpa_akun_berperan('Bunga'), ep)
                keterangan = keterangan or f"Beban {sub}"
            else:
                beban_kode = OPERASIONAL_MAP.get(sub, '6180')
                keterangan = keterangan or f"Beban {sub}"
            entries = [(beban_kode, nominal, 0)] + kas_credits(uang_keluar)
            if hutang_nm > 0: entries.append((akun_hutang_kode, 0, hutang_nm))
            jid = insert_jurnal(conn, tanggal, keterangan, 'OPERASIONAL', 'PENGELUARAN', with_tax(entries), pihak=nama_vendor)
            _akun_debit = beban_kode
            _peng_items = [{'produk_id': None, 'deskripsi': keterangan, 'qty': 1, 'harga': nominal, 'arah': 'NONE'}]

        elif kategori == 'PAJAK':
            keterangan = keterangan or 'Pembayaran Pajak'
            # Akun beban pajak override (default 6170 Beban Pajak)
            akun_pajak_override = request.form.get('akun_pajak_kode','').strip()
            if akun_pajak_override:
                _ap = conn.execute("SELECT kode FROM akun WHERE kode=? AND tipe='BEBAN'", (akun_pajak_override,)).fetchone()
                if not _ap:
                    return _reject_form(conn, f'Akun beban pajak COA "{akun_pajak_override}" tidak valid.', ep)
                pajak_kode = akun_pajak_override
            else:
                # Akun berperan PAJAK, bukan asal nomor 6170 (lihat _akun_beban_berperan).
                pajak_kode = _akun_beban_berperan(conn, 'PAJAK')
                if not pajak_kode:
                    return _reject_form(conn, _pesan_tanpa_akun_berperan('Pajak'), ep)
            entries = [(pajak_kode, nominal, 0)] + kas_credits(uang_keluar)
            if hutang_nm > 0: entries.append((akun_hutang_kode, 0, hutang_nm))
            jid = insert_jurnal(conn, tanggal, keterangan, 'OPERASIONAL', 'PENGELUARAN', with_tax(entries), pihak=nama_vendor)
            _akun_debit = pajak_kode
            _peng_items = [{'produk_id': None, 'deskripsi': keterangan, 'qty': 1, 'harga': nominal, 'arah': 'NONE'}]

        elif kategori == 'PENARIKAN_OWNER':
            keterangan = keterangan or 'Penarikan Owner / Prive'
            # Akun ekuitas (Prive) override, default 3300
            akun_prive_override = request.form.get('akun_prive_kode','').strip()
            if akun_prive_override:
                _ae = conn.execute("SELECT kode FROM akun WHERE kode=? AND tipe='EKUITAS'", (akun_prive_override,)).fetchone()
                if not _ae:
                    return _reject_form(conn, f'Akun ekuitas / prive COA "{akun_prive_override}" tidak valid.', ep)
                prive_kode = akun_prive_override
            else:
                prive_kode = '3300'
            entries = [(prive_kode, nominal, 0)] + kas_credits(uang_keluar)
            if hutang_nm > 0: entries.append((akun_hutang_kode, 0, hutang_nm))
            jid = insert_jurnal(conn, tanggal, keterangan, 'PENDANAAN', 'PENGELUARAN', with_tax(entries), pihak=nama_vendor)
            _peng_editable = False   # prive: tampil item saja, edit lewat jurnal
            _peng_items = [{'produk_id': None, 'deskripsi': keterangan, 'qty': 1, 'harga': nominal, 'arah': 'NONE'}]

        elif kategori == 'PINJAMAN_KARYAWAN':
            # NON-BEBAN: pinjaman/kasbon karyawan dicatat sebagai ASET (1160 Piutang Karyawan),
            # TIDAK menyentuh Laba Rugi. Pengembalian dicatat lewat halaman Pelunasan
            # (Dr kas / Cr 1160). Saldo per karyawan diturunkan dari ledger via jurnal.pihak,
            # jadi pihak WAJIB diisi nama karyawan yang konsisten.
            karyawan_nama = (request.form.get('karyawan_nama','') or '').strip()
            if not karyawan_nama:
                return _reject_form(conn, 'Nama karyawan wajib diisi untuk pinjaman karyawan.', ep)
            if x_ppn_nom > 0 or x_pph_nom > 0:
                return _reject_form(conn, 'Pinjaman karyawan tidak dikenai PPN/PPh.', ep)
            if abs(uang_keluar - nominal) > 0.01:
                return _reject_form(conn, 'Pinjaman karyawan harus dibayar penuh dari kas/bank '
                                          '(tidak bisa dicatat sebagai hutang).', ep)
            hutang_nm = 0            # tegas: tidak ada cabang hutang utk pemberian pinjaman
            keterangan = keterangan or f'Pinjaman karyawan: {karyawan_nama}'
            entries = [('1160', nominal, 0)] + kas_credits(uang_keluar)
            # kategori arus kas 'OPERASIONAL' (bukan PENDANAAN/INVESTASI: label & export
            # akan salah baca; lihat sub_label daftar transaksi + filter export arus kas)
            jid = insert_jurnal(conn, tanggal, keterangan, 'OPERASIONAL', 'PENGELUARAN', entries, pihak=karyawan_nama)
            _peng_editable = False   # koreksi lewat hapus-catat-ulang / jurnal manual
            _peng_items = [{'produk_id': None, 'deskripsi': keterangan, 'qty': 1, 'harga': nominal, 'arah': 'NONE'}]

        elif kategori in ('BAHAN_BAKU', 'INVESTASI_ASET'):
            if kategori == 'BAHAN_BAKU':
                keterangan = keterangan or 'Belanja Bahan Baku'
                # Akun persediaan override, default 1130 Persediaan Barang
                akun_persediaan_override = request.form.get('akun_persediaan_kode','').strip()
                if akun_persediaan_override:
                    _ai = conn.execute(
                        "SELECT kode FROM akun WHERE kode=? AND tipe='ASET' AND saldo_normal='DEBIT' "
                        "AND (subtipe LIKE '%ersediaan%' OR kode LIKE '113%')",
                        (akun_persediaan_override,)
                    ).fetchone()
                    if not _ai:
                        return _reject_form(conn, f'Akun persediaan COA "{akun_persediaan_override}" tidak valid.', ep)
                    persediaan_kode = akun_persediaan_override
                else:
                    persediaan_kode = '1130'
                entries = [(persediaan_kode, nominal, 0)] + kas_credits(uang_keluar)
                if hutang_nm > 0: entries.append((akun_hutang_kode, 0, hutang_nm))
                jid = insert_jurnal(conn, tanggal, keterangan, 'OPERASIONAL', 'PENGELUARAN', with_tax(entries), pihak=nama_vendor)
                _akun_debit = persediaan_kode
                # ── Tabel terpadu: tiap baris = produk lama (produk_id) ATAU produk baru (new_nama) ──
                ids     = request.form.getlist('produk_id_bb[]')
                names   = request.form.getlist('new_nama_bb[]')
                satuans = request.form.getlist('satuan_bb[]')
                qtys    = request.form.getlist('qty_bb_pilih[]')
                hrgs    = request.form.getlist('harga_bb_pilih[]')
                n_rows  = max(len(ids), len(names), len(qtys))
                import re as _re
                for i in range(n_rows):
                    pid_str = (ids[i].strip() if i < len(ids) and ids[i] else '')
                    nm      = (names[i].strip() if i < len(names) and names[i] else '')
                    raw_qty = qtys[i] if i < len(qtys) else ''
                    try:
                        qn = parse_qty(raw_qty)
                    except (ValueError, TypeError):
                        return _reject_form(conn, f'Qty bahan baku baris {i + 1} tidak valid.', ep)
                    try:
                        hn = parse_nonnegative_rp(hrgs[i] if i < len(hrgs) else '0',
                                                  f'Harga bahan baku baris {i + 1}')
                    except ValueError as ex:
                        return _reject_form(conn, str(ex), ep)
                    sat = (satuans[i].strip() if i < len(satuans) and satuans[i] else 'pcs') or 'pcs'
                    if qn < 0:
                        return _reject_form(conn, f'Qty bahan baku baris {i + 1} tidak boleh negatif.', ep)
                    if qn == 0:
                        if pid_str or nm:
                            return _reject_form(conn, f'Qty bahan baku baris {i + 1} harus lebih dari 0.', ep)
                        continue

                    if pid_str:
                        # Produk yang sudah ada
                        try:
                            pid = int(pid_str)
                        except ValueError:
                            return _reject_form(conn, f'Produk bahan baku baris {i + 1} tidak valid.', ep)
                        _prow = conn.execute(
                            'SELECT nama,satuan,stok,harga_beli FROM produk WHERE id=? AND COALESCE(aktif,1)=1',
                            (pid,)
                        ).fetchone()
                        if not _prow:
                            return _reject_form(conn, f'Produk bahan baku baris {i + 1} tidak ditemukan atau sudah diarsipkan.', ep)
                        if hn <= 0:
                            hn = _prow['harga_beli']
                        _desc = _prow['nama']
                        _sat  = _prow['satuan']
                    elif nm:
                        # Produk baru → daftarkan ke inventory
                        kode_base = _re.sub(r'[^A-Z0-9]', '-', nm.upper())[:14].strip('-') or 'BRG'
                        kode_new = kode_base; _sfx = 1
                        while conn.execute('SELECT id FROM produk WHERE kode=?', (kode_new,)).fetchone():
                            kode_new = f'{kode_base}-{_sfx}'; _sfx += 1
                        conn.execute(
                            'INSERT INTO produk(kode,nama,satuan,harga_beli,harga_jual,stok,min_stok) VALUES(?,?,?,?,?,0,0)',
                            (kode_new, nm, sat, hn, 0)
                        )
                        pid = conn.execute('SELECT last_insert_rowid()').fetchone()[0]
                        _desc, _sat = nm, sat
                        _prow = {'stok': 0, 'harga_beli': hn}
                    else:
                        return _reject_form(conn, f'Nama produk bahan baku baris {i + 1} wajib diisi.', ep)
                    if hn <= 0:
                        return _reject_form(conn, f'Harga bahan baku baris {i + 1} harus lebih dari 0.', ep)

                    update_average_cost(conn, pid, qn, hn)
                    conn.execute('UPDATE produk SET stok=stok+? WHERE id=?', (qn, pid))
                    record_stock_movement(conn, pid, tanggal, 'MASUK', qn, hn,
                                          keterangan, jid, 'ADD_LAYER',
                                          _prow['stok'], _prow['harga_beli'])
                    _peng_items.append({'produk_id': pid, 'deskripsi': _desc, 'qty': qn,
                                        'satuan': _sat, 'harga': hn, 'arah': 'MASUK'})
                non_sku_desc = (request.form.get('non_sku_deskripsi', '') or '').strip()
                try:
                    non_sku_nilai = parse_nonnegative_rp(request.form.get('non_sku_nilai', '0'),
                                                         'Valuasi persediaan non-SKU')
                except ValueError as ex:
                    return _reject_form(conn, str(ex), ep)
                if non_sku_nilai > 0:
                    non_sku_desc = non_sku_desc or 'Persediaan bahan baku non-SKU'
                    _peng_items.append({'produk_id': None, 'deskripsi': non_sku_desc, 'qty': 1,
                                        'satuan': 'nilai', 'harga': non_sku_nilai,
                                        'subtotal': non_sku_nilai, 'arah': 'NON_SKU'})
                    record_non_sku_movement(conn, jid, tanggal, non_sku_desc, non_sku_nilai)
            else:
                nama_aset  = request.form.get('nama_aset','Aset Baru')
                kat_aset   = request.form.get('kategori_aset','Peralatan')
                aset_kode_map = {'Peralatan':'1200','Kendaraan':'1210','Gedung':'1220'}
                if kat_aset not in aset_kode_map:
                    return _reject_form(conn, 'Kategori aset tidak valid.', ep)
                try:
                    masa_pakai = int(request.form.get('masa_pakai',12) or 12)
                except (TypeError, ValueError):
                    return _reject_form(conn, 'Masa pakai aset harus berupa bilangan bulat.', ep)
                if masa_pakai <= 0:
                    return _reject_form(conn, 'Masa pakai aset harus lebih dari 0 bulan.', ep)
                peny_bln   = round(nominal / masa_pakai, 2)
                # Akun aset tetap override, default 1200/1210/1220 dari kategori
                akun_aset_override = request.form.get('akun_aset_kode','').strip()
                if akun_aset_override:
                    _at = conn.execute("SELECT kode FROM akun WHERE kode=? AND tipe='ASET' AND saldo_normal='DEBIT'", (akun_aset_override,)).fetchone()
                    if not _at:
                        return _reject_form(conn, f'Akun aset tetap COA "{akun_aset_override}" tidak valid.', ep)
                    aset_kode = akun_aset_override
                else:
                    aset_kode = aset_kode_map[kat_aset]
                keterangan = keterangan or f"Investasi {kat_aset}: {nama_aset}"
                entries = [(aset_kode, nominal, 0)] + kas_credits(uang_keluar)
                if hutang_nm > 0: entries.append((akun_hutang_kode, 0, hutang_nm))
                jid = insert_jurnal(conn, tanggal, keterangan, 'INVESTASI', 'PENGELUARAN', with_tax(entries), pihak=nama_vendor)
                conn.execute(
                    "INSERT INTO aset_tetap(nama,kategori,harga_beli,tanggal_beli,masa_pakai,penyusutan_bulan,jurnal_id) VALUES(?,?,?,?,?,?,?)",
                    (nama_aset, kat_aset, nominal, tanggal, masa_pakai, peny_bln, jid)
                )
                _aset_id_new = conn.execute('SELECT last_insert_rowid()').fetchone()[0]
                conn.execute("UPDATE jurnal SET aset_id=? WHERE id=?", (_aset_id_new, jid))
                _peng_editable = False   # aset: edit lewat jurnal (terkait penyusutan)
                _peng_items = [{'produk_id': None, 'deskripsi': nama_aset, 'qty': 1, 'harga': nominal, 'arah': 'NONE'}]

        if kategori == 'BAHAN_BAKU' and not _peng_items:
            conn.rollback(); conn.close()
            flash('Isi minimal satu produk terlacak atau satu valuasi persediaan non-SKU.', 'danger')
            return redirect(url_for(ep))
        if kategori == 'BAHAN_BAKU' and _peng_items:
            total_item = sum(float(it.get('qty', 0) or 0) * float(it.get('harga', 0) or 0)
                             for it in _peng_items)
            # Di mode rupiah bulat, nominal sudah dibulatkan ke rupiah utuh sementara total
            # rincian bisa berakhiran pecahan (qty desimal x harga). Bandingkan sesudah
            # dibulatkan supaya selisih sub-rupiah tidak salah ditolak.
            if get_setting(conn, 'pref_decimal', '0') != '1':
                cocok = abs(_round_rp(total_item) - nominal) <= 0.01
            else:
                cocok = abs(total_item - nominal) <= 0.01
            if not cocok:
                conn.rollback(); conn.close()
                flash('Total rincian produk harus sama dengan Nominal Invoice bahan baku.', 'danger')
                return redirect(url_for(ep))

        # ── Insert ke tabel hutang (berlaku untuk SEMUA kategori sekarang) ───────
        if hutang_nm > 0:
            _default_pemasok = {
                'BAHAN_BAKU':       'Pemasok',
                'INVESTASI_ASET':   'Vendor Aset',
                'OPERASIONAL':      'Kreditur Operasional',
                'PAJAK':            'Kantor Pajak',
                'PENARIKAN_OWNER':  'Hutang ke Owner',
            }.get(kategori, 'Kreditur')
            nama_kreditur = pemasok_raw or _default_pemasok
            conn.execute(
                "INSERT INTO hutang(tanggal,jatuh_tempo,pemasok,keterangan,jumlah,jurnal_id,akun_kode) VALUES(?,?,?,?,?,?,?)",
                (tanggal, jt, nama_kreditur, keterangan, hutang_nm, jid, akun_hutang_kode)
            )

        # ── Simpan rincian item + metadata recompute ─────────────────────────────
        if jid:
            if _peng_items:
                _store_items(conn, jid, _peng_items)
            if _peng_editable and _akun_debit:
                conn.execute("UPDATE jurnal SET tx_meta=? WHERE id=?", (json.dumps({
                    'tipe': 'BELI', 'akun_debit': _akun_debit, 'akun_kas': _single_kas,
                    'uang_keluar': uang_keluar, 'pihak': nama_vendor, 'jt': jt,
                    'kategori': kategori, 'akun_hutang': akun_hutang_kode,
                    'kas_splits': _kas_splits_meta,
                    'ppn_persen': x_ppn_pct, 'pph23_persen': x_pph_pct}), jid))
            if proyek_id:
                conn.execute("UPDATE jurnal SET proyek_id=? WHERE id=?", (proyek_id, jid))
            # ── Ref dokumen pembelian (opsional, non-akuntansi): No. Invoice Supplier + No. Faktur Pajak ──
            # Disimpan di tx_meta sebagai catatan. Merge (read-modify-write) supaya berlaku untuk
            # kategori editable (BAHAN_BAKU, tx_meta sudah ada) maupun non-editable (INVESTASI_ASET).
            _no_invoice = (request.form.get('no_invoice_supplier','') or '').strip()[:60]
            _no_faktur  = (request.form.get('no_faktur_pajak','') or '').strip()[:60]
            if _no_invoice or _no_faktur:
                _row = conn.execute("SELECT tx_meta FROM jurnal WHERE id=?", (jid,)).fetchone()
                try:
                    _m = json.loads(_row['tx_meta']) if (_row and _row['tx_meta']) else {}
                    if not isinstance(_m, dict): _m = {}
                except Exception:
                    _m = {}
                if _no_invoice: _m['no_invoice_supplier'] = _no_invoice
                if _no_faktur:  _m['no_faktur_pajak'] = _no_faktur
                conn.execute("UPDATE jurnal SET tx_meta=? WHERE id=?", (json.dumps(_m), jid))

        add_log(conn, 'Pengeluaran dicatat', f"{keterangan} | {tanggal} | {kategori}", 'INPUT')
        conn.commit(); conn.close()
        flash('Pengeluaran berhasil dicatat!', 'success')
        return redirect(url_for('dashboard'))


@app.route('/pengeluaran', methods=['GET','POST'])
@expense_input_required
def pengeluaran():
    """Halaman Pengeluaran: beban operasional, pajak, prive & pinjaman karyawan.
    (Belanja bahan baku & aset tetap pindah ke halaman Pembelian — handler POST sama.)"""
    conn = db()
    if request.method == 'POST':
        return _proses_pengeluaran(conn, 'pengeluaran')
    akun_kas = conn.execute("SELECT * FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    vendor_names = [r['nama'] for r in conn.execute("SELECT nama FROM vendor ORDER BY nama").fetchall()]
    # Nama karyawan utk kategori Pinjaman Karyawan (tabel karyawan + pihak pinjaman berjalan)
    karyawan_names = [r['nama'] for r in conn.execute(
        "SELECT nama FROM karyawan WHERE COALESCE(aktif,1)=1 ORDER BY nama"
    ).fetchall()]
    # ── Akun COA per kategori (untuk override default mapping) ──
    akun_beban_list = conn.execute(
        "SELECT kode, nama FROM akun WHERE tipe='BEBAN' AND saldo_normal='DEBIT' "
        "AND kode NOT IN ('5100','5150','5190','6130') ORDER BY kode"
    ).fetchall()
    # Beban pajak (untuk PAJAK override), biasanya 6170 + akun pajak lain yang user tambah
    akun_beban_pajak_list = conn.execute(
        "SELECT kode, nama FROM akun WHERE tipe='BEBAN' AND saldo_normal='DEBIT' ORDER BY kode"
    ).fetchall()
    # Ekuitas debit (untuk PENARIKAN_OWNER override), biasanya 3300 Prive
    akun_ekuitas_debit_list = conn.execute(
        "SELECT kode, nama FROM akun WHERE tipe='EKUITAS' AND saldo_normal='DEBIT' ORDER BY kode"
    ).fetchall()
    proyek_aktif = conn.execute(
        "SELECT id, kode, nama FROM proyek WHERE status='AKTIF' ORDER BY nama"
    ).fetchall()
    akun_hutang_list = _akun_hutang_pilihan(conn)
    conn.close()
    return render_template('pengeluaran.html', akun_kas=akun_kas,
                           today=date.today().strftime('%Y-%m-%d'),
                           subkategori_list=list(OPERASIONAL_MAP.keys()),
                           vendor_names=vendor_names,
                           karyawan_names=karyawan_names,
                           akun_beban_list=akun_beban_list,
                           akun_beban_pajak_list=akun_beban_pajak_list,
                           akun_ekuitas_debit_list=akun_ekuitas_debit_list,
                           akun_hutang_list=akun_hutang_list,
                           proyek_aktif=proyek_aktif)


@app.route('/pembelian', methods=['GET','POST'])
@expense_input_required
def pembelian():
    """Halaman Pembelian: belanja bahan baku / persediaan & pembelian aset tetap.
    Logika posting = _proses_pengeluaran (kategori BAHAN_BAKU / INVESTASI_ASET),
    identik dengan jalur lama di halaman Pengeluaran — hanya pintu masuknya yang pindah."""
    conn = db()
    if request.method == 'POST':
        if request.form.get('kategori') not in ('BAHAN_BAKU', 'INVESTASI_ASET'):
            return _reject_form(conn, 'Kategori tidak valid untuk halaman Pembelian.', 'pembelian')
        return _proses_pengeluaran(conn, 'pembelian')
    akun_kas = conn.execute("SELECT * FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    produk_list_bb = conn.execute(
        "SELECT id, kode, nama, varian, satuan, stok, harga_beli FROM produk "
        "WHERE COALESCE(aktif,1)=1 ORDER BY nama"
    ).fetchall()
    vendor_names = [r['nama'] for r in conn.execute("SELECT nama FROM vendor ORDER BY nama").fetchall()]
    # Persediaan (untuk BAHAN_BAKU override), semua aset lancar tipe persediaan
    akun_persediaan_list = conn.execute(
        "SELECT kode, nama FROM akun WHERE tipe='ASET' AND saldo_normal='DEBIT' "
        "AND (subtipe LIKE '%ersediaan%' OR kode LIKE '113%') ORDER BY kode"
    ).fetchall()
    # Aset tetap (untuk INVESTASI_ASET override)
    akun_aset_tetap_list = conn.execute(
        "SELECT kode, nama FROM akun WHERE tipe='ASET' AND saldo_normal='DEBIT' "
        "AND (subtipe LIKE '%set Tetap%' OR kode LIKE '12%') AND kode != '1290' ORDER BY kode"
    ).fetchall()
    prefill_produk_id = request.args.get('produk_id', type=int)
    proyek_aktif = conn.execute(
        "SELECT id, kode, nama FROM proyek WHERE status='AKTIF' ORDER BY nama"
    ).fetchall()
    akun_hutang_list = _akun_hutang_pilihan(conn)
    conn.close()
    return render_template('pembelian.html', akun_kas=akun_kas,
                           today=date.today().strftime('%Y-%m-%d'),
                           produk_list_bb=produk_list_bb, vendor_names=vendor_names,
                           akun_persediaan_list=akun_persediaan_list,
                           akun_aset_tetap_list=akun_aset_tetap_list,
                           akun_hutang_list=akun_hutang_list,
                           proyek_aktif=proyek_aktif,
                           prefill_produk_id=prefill_produk_id)


# ---------- RETUR PENJUALAN ----------
@app.route('/retur-penjualan', methods=['GET','POST'])
@operator_required
def retur_penjualan():
    conn = db()
    produk_list = conn.execute("SELECT * FROM produk WHERE COALESCE(aktif,1)=1 ORDER BY nama").fetchall()
    akun_kas = conn.execute("SELECT * FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()

    if request.method == 'POST':
        mode       = request.form.get('mode', 'umum')
        tanggal    = request.form['tanggal']
        keterangan = request.form.get('keterangan', 'Retur Penjualan')
        sumber     = request.form.get('sumber', 'kas')   # 'kas' atau 'piutang'
        akun_kas_kode = request.form.get('akun_kas', '1100')

        if not _operator_date_ok(tanggal):
            conn.close()
            flash('Tanggal transaksi tidak valid. Operator hanya bisa input transaksi untuk bulan berjalan.', 'warning')
            return redirect(url_for('retur_penjualan'))
        if mode not in ('umum', 'produk'):
            return _reject_form(conn, 'Mode retur penjualan tidak valid.', 'retur_penjualan')
        if sumber not in ('kas', 'piutang'):
            return _reject_form(conn, 'Sumber retur penjualan tidak valid.', 'retur_penjualan')
        if sumber == 'kas' and not is_rekening_kode(conn, akun_kas_kode):
            return _reject_form(conn, 'Rekening kas/bank tidak valid.', 'retur_penjualan')

        if mode == 'umum':
            try:
                nominal = parse_nonnegative_rp(request.form.get('nominal', '0'), 'Nominal retur')
                hpp_balik = parse_nonnegative_rp(request.form.get('hpp_balik', '0'), 'HPP yang dibalik')
            except ValueError as ex:
                return _reject_form(conn, str(ex), 'retur_penjualan')
            if nominal <= 0:
                conn.close()
                flash('Nominal retur harus lebih dari 0.', 'danger')
                return redirect(url_for('retur_penjualan'))
            pelanggan = request.form.get('pelanggan', '').strip()
            try:
                tracker = _select_return_tracker(conn, 'PIUTANG', request.form.get('piutang_id'), pelanggan, nominal) if sumber == 'piutang' else None
            except ValueError as ex:
                return _reject_form(conn, str(ex), 'retur_penjualan')
            if sumber == 'piutang':
                if not tracker:
                    conn.close()
                    flash('Piutang aktif dengan sisa yang cukup tidak ditemukan. Periksa pelanggan atau nominal retur.', 'danger')
                    return redirect(url_for('retur_penjualan'))
                if nominal > (tracker['jumlah'] - tracker['terbayar']) + 0.01:
                    return _reject_form(conn, f'Nominal retur melebihi sisa piutang Rp {tracker["jumlah"] - tracker["terbayar"]:,.0f}.', 'retur_penjualan')

            tax = _tracker_return_tax_parts(conn, 'PIUTANG', tracker, nominal, 'NET') if tracker else {
                'dpp': nominal, 'ppn': 0.0, 'pph': 0.0, 'pph22': 0.0, 'net': nominal
            }
            entries = [('4150', tax['dpp'], 0)]   # debit kontra-pendapatan
            if tax['ppn'] > 0:
                entries.append(('2111', tax['ppn'], 0))  # balik Hutang PPN
            if sumber == 'piutang':
                entries.append(('1120', 0, tax['net']))   # kurangi piutang
            else:
                entries.append((akun_kas_kode, 0, tax['net']))   # uang keluar
            if tax['pph'] > 0:
                entries.append(('1181', 0, tax['pph']))  # balik PPh 23 dimuka
            if tax.get('pph22', 0) > 0:
                entries.append(('1183', 0, tax['pph22']))  # balik PPh 22 dimuka
            # HPP reversal opsional — wajib dibatasi: tidak boleh melebihi
            # persediaan non-SKU yang sudah dikonsumsi dan belum dikembalikan.
            hpp_inventory_balik = 0.0
            hpp_reclass_balik = {}
            if hpp_balik > 0:
                src_hpp = sale_hpp_credit_breakdown(conn, tracker['jurnal_id']) if tracker and tracker['jurnal_id'] else None
                if src_hpp and src_hpp['total'] > 0:
                    if hpp_balik > src_hpp['total'] + 0.01:
                        return _reject_form(
                            conn,
                            f'HPP retur ({hpp_balik:,.0f}) melebihi HPP transaksi sumber '
                            f'({src_hpp["total"]:,.0f}).',
                            'retur_penjualan'
                        )
                    ratio_balik = hpp_balik / src_hpp['total'] if src_hpp['total'] > 0 else 0
                    hpp_inventory_balik = src_hpp['credits'].get('1130', 0.0) * ratio_balik
                    hpp_reclass_balik = {
                        k: v * ratio_balik for k, v in src_hpp['credits'].items()
                        if k != '1130' and v > 0
                    }
                else:
                    hpp_inventory_balik = hpp_balik
                cap_non_sku = returnable_non_sku_value(conn)
                if hpp_inventory_balik > cap_non_sku + 0.01:
                    return _reject_form(
                        conn,
                        f'HPP persediaan yang dibalik ({hpp_inventory_balik:,.0f}) melebihi persediaan non-SKU yang dapat dikembalikan '
                        f'({cap_non_sku:,.0f}). Untuk retur produk ber-SKU gunakan mode "per produk", '
                        'atau lakukan koreksi persediaan lebih dulu.',
                        'retur_penjualan'
                    )
                if hpp_inventory_balik > 0:
                    entries.append(('1130', hpp_inventory_balik, 0))
                for _akun_beban, _nilai in sorted(hpp_reclass_balik.items()):
                    entries.append((_akun_beban, _nilai, 0))
                entries.append(('5100', 0, hpp_balik))

            jid = insert_jurnal(conn, tanggal, keterangan, 'OPERASIONAL', 'RETUR_JUAL', entries)
            # Inherit proyek_id dari sale sumber (via tracker piutang) supaya P&L proyek konsisten.
            if tracker and tracker['jurnal_id']:
                src_pid = conn.execute("SELECT proyek_id FROM jurnal WHERE id=?",
                                       (tracker['jurnal_id'],)).fetchone()
                if src_pid and src_pid['proyek_id']:
                    conn.execute("UPDATE jurnal SET proyek_id=? WHERE id=?",
                                 (src_pid['proyek_id'], jid))
            if hpp_inventory_balik > 0:
                record_non_sku_movement(conn, jid, tanggal, f"Retur penjualan non-SKU: {keterangan}", hpp_inventory_balik)
            if tracker:
                apply_tracker_adjustment(conn, 'PIUTANG', tracker['id'], -tax['net'], jid)

        else:  # per produk
            produk_ids   = request.form.getlist('produk_id[]')
            qtys         = request.form.getlist('qty[]')
            harga_juals  = request.form.getlist('harga_jual[]')

            items = []
            for i, pid_str in enumerate(produk_ids):
                if not pid_str: continue
                try:
                    pid   = int(pid_str)
                    qty   = parse_positive_qty(qtys[i] if i < len(qtys) else 1, f'Qty item baris {i + 1}')
                    harga = parse_nonnegative_rp(harga_juals[i] if i < len(harga_juals) else '0',
                                                 f'Harga jual item baris {i + 1}')
                except (ValueError, TypeError):
                    return _reject_form(conn, f'Rincian produk baris {i + 1} tidak valid.', 'retur_penjualan')
                produk = conn.execute(
                    "SELECT * FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
                    (pid,)
                ).fetchone()
                if not produk:
                    return _reject_form(conn, f'Produk pada baris {i + 1} tidak ditemukan atau sudah diarsipkan.', 'retur_penjualan')
                subtotal = qty * harga
                hpp_item = qty * produk['harga_beli']
                items.append({'produk': produk, 'qty': qty, 'harga': harga,
                              'subtotal': subtotal, 'hpp': hpp_item})

            total_sub = sum(it['subtotal'] for it in items)
            total_hpp = sum(it['hpp'] for it in items)

            if total_sub <= 0 or not items:
                conn.close()
                flash('Pilih minimal 1 produk dengan nominal > 0.', 'danger')
                return redirect(url_for('retur_penjualan'))
            pelanggan = request.form.get('pelanggan', '').strip()
            try:
                tracker = _select_return_tracker(conn, 'PIUTANG', request.form.get('piutang_id'), pelanggan, total_sub) if sumber == 'piutang' else None
            except ValueError as ex:
                return _reject_form(conn, str(ex), 'retur_penjualan')
            if sumber == 'piutang' and not tracker:
                conn.close()
                flash('Piutang aktif dengan sisa yang cukup tidak ditemukan. Periksa pelanggan atau nominal retur.', 'danger')
                return redirect(url_for('retur_penjualan'))
            tax = _tracker_return_tax_parts(conn, 'PIUTANG', tracker, total_sub, 'DPP') if tracker else {
                'dpp': total_sub, 'ppn': 0.0, 'pph': 0.0, 'pph22': 0.0, 'net': total_sub
            }
            if tracker and tax['net'] > (tracker['jumlah'] - tracker['terbayar']) + 0.01:
                return _reject_form(conn, f'Nominal retur melebihi sisa piutang Rp {tracker["jumlah"] - tracker["terbayar"]:,.0f}.', 'retur_penjualan')

            entries = [('4150', tax['dpp'], 0)]
            if tax['ppn'] > 0:
                entries.append(('2111', tax['ppn'], 0))
            if sumber == 'piutang':
                entries.append(('1120', 0, tax['net']))
            else:
                entries.append((akun_kas_kode, 0, tax['net']))
            if tax['pph'] > 0:
                entries.append(('1181', 0, tax['pph']))
            if tax.get('pph22', 0) > 0:
                entries.append(('1183', 0, tax['pph22']))
            if total_hpp > 0:
                entries += [('1130', total_hpp, 0), ('5100', 0, total_hpp)]

            jid = insert_jurnal(conn, tanggal, keterangan, 'OPERASIONAL', 'RETUR_JUAL', entries)
            # Inherit proyek_id dari sale sumber (via tracker piutang) supaya P&L proyek konsisten.
            if tracker and tracker['jurnal_id']:
                src_pid = conn.execute("SELECT proyek_id FROM jurnal WHERE id=?",
                                       (tracker['jurnal_id'],)).fetchone()
                if src_pid and src_pid['proyek_id']:
                    conn.execute("UPDATE jurnal SET proyek_id=? WHERE id=?",
                                 (src_pid['proyek_id'], jid))

            # Stok masuk kembali
            return_items = []
            for item in items:
                current = conn.execute("SELECT stok,harga_beli FROM produk WHERE id=?",
                                       (item['produk']['id'],)).fetchone()
                stok_sebelum = current['stok']
                hpp_sebelum = current['harga_beli']
                update_average_cost(conn, item['produk']['id'], item['qty'], hpp_sebelum)
                conn.execute("UPDATE produk SET stok=stok+? WHERE id=?",
                             (item['qty'], item['produk']['id']))
                record_stock_movement(conn, item['produk']['id'], tanggal, 'MASUK',
                                      item['qty'], hpp_sebelum, f"Retur: {keterangan}",
                                      jid, 'ADD_LAYER', stok_sebelum, hpp_sebelum)
                return_items.append({'produk_id': item['produk']['id'], 'deskripsi': item['produk']['nama'],
                                     'qty': item['qty'], 'satuan': item['produk']['satuan'],
                                     'harga': item['harga'], 'subtotal': item['subtotal'], 'arah': 'MASUK'})
            _store_items(conn, jid, return_items)
            if tracker:
                apply_tracker_adjustment(conn, 'PIUTANG', tracker['id'], -tax['net'], jid)

        add_log(conn, 'Retur Penjualan', f"{keterangan} | {tanggal}", 'INPUT')
        conn.commit(); conn.close()
        flash('Retur penjualan berhasil dicatat!', 'success')
        return redirect(url_for('retur_penjualan'))

    saldo_persediaan = persediaan_ledger_value(conn)
    piutang_aktif = conn.execute(
        """SELECT *, (jumlah-terbayar) AS sisa
           FROM piutang
           WHERE status!='LUNAS'
           ORDER BY tanggal DESC,id DESC"""
    ).fetchall()
    customer_names = [r['nama'] for r in conn.execute("SELECT nama FROM customer ORDER BY nama").fetchall()]
    conn.close()
    return render_template('retur_penjualan.html',
                           produk_list=produk_list, akun_kas=akun_kas,
                           today=date.today().strftime('%Y-%m-%d'),
                           saldo_persediaan=saldo_persediaan,
                           piutang_aktif=piutang_aktif,
                           customer_names=customer_names)


# ---------- RETUR PEMBELIAN ----------
@app.route('/retur-pembelian', methods=['GET','POST'])
@operator_required
def retur_pembelian():
    conn = db()
    produk_list = conn.execute("SELECT * FROM produk WHERE COALESCE(aktif,1)=1 ORDER BY nama").fetchall()
    akun_kas = conn.execute("SELECT * FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()

    if request.method == 'POST':
        mode       = request.form.get('mode', 'umum')
        tanggal    = request.form['tanggal']
        keterangan = request.form.get('keterangan', 'Retur Pembelian')
        sumber     = request.form.get('sumber', 'kas')   # 'kas' atau 'hutang'
        akun_kas_kode = request.form.get('akun_kas', '1100')

        if not _operator_date_ok(tanggal):
            conn.close()
            flash('Tanggal transaksi tidak valid. Operator hanya bisa input transaksi untuk bulan berjalan.', 'warning')
            return redirect(url_for('retur_pembelian'))
        if mode not in ('umum', 'produk'):
            return _reject_form(conn, 'Mode retur pembelian tidak valid.', 'retur_pembelian')
        if sumber not in ('kas', 'hutang'):
            return _reject_form(conn, 'Sumber retur pembelian tidak valid.', 'retur_pembelian')
        if sumber == 'kas' and not is_rekening_kode(conn, akun_kas_kode):
            return _reject_form(conn, 'Rekening kas/bank tidak valid.', 'retur_pembelian')

        if mode == 'umum':
            try:
                nominal = parse_nonnegative_rp(request.form.get('nominal', '0'), 'Nominal retur')
            except ValueError as ex:
                return _reject_form(conn, str(ex), 'retur_pembelian')
            if nominal <= 0:
                conn.close()
                flash('Nominal retur harus lebih dari 0.', 'danger')
                return redirect(url_for('retur_pembelian'))
            pemasok = request.form.get('pemasok', '').strip()
            try:
                tracker = _select_return_tracker(conn, 'HUTANG', request.form.get('hutang_id'), pemasok, nominal) if sumber == 'hutang' else None
            except ValueError as ex:
                return _reject_form(conn, str(ex), 'retur_pembelian')
            if sumber == 'hutang':
                if not tracker:
                    conn.close()
                    flash('Hutang aktif dengan sisa yang cukup tidak ditemukan. Periksa pemasok atau nominal retur.', 'danger')
                    return redirect(url_for('retur_pembelian'))
                if nominal > (tracker['jumlah'] - tracker['terbayar']) + 0.01:
                    return _reject_form(conn, f'Nominal retur melebihi sisa hutang Rp {tracker["jumlah"] - tracker["terbayar"]:,.0f}.', 'retur_pembelian')
            akun_hutang_kode = (tracker['akun_kode'] or '2100') if tracker else '2100'
            tax = _tracker_return_tax_parts(conn, 'HUTANG', tracker, nominal, 'NET') if tracker else {
                'dpp': nominal, 'ppn': 0.0, 'pph': 0.0, 'net': nominal
            }
            saldo_non_sku = current_non_sku_value(conn)
            if tax['dpp'] > saldo_non_sku + 0.01:
                conn.close()
                flash(f'Retur pembelian umum melebihi persediaan Non-SKU tersedia Rp {saldo_non_sku:,.0f}. Gunakan retur per produk untuk SKU atau lakukan koreksi persediaan lebih dulu.', 'danger')
                return redirect(url_for('retur_pembelian'))

            # App ini pakai perpetual inventory: pembelian langsung ke 1130, jadi
            # retur langsung mengurangi 1130. Akun 5150 tersedia di COA untuk
            # entry manual periodic-style bila diperlukan.
            if sumber == 'hutang':
                entries = [(akun_hutang_kode, tax['net'], 0)]
            else:
                entries = [(akun_kas_kode, tax['net'], 0)]
            if tax['pph'] > 0:
                entries.append(('2112', tax['pph'], 0))
            entries.append(('1130', 0, tax['dpp']))
            if tax['ppn'] > 0:
                entries.append(('1180', 0, tax['ppn']))

            jid = insert_jurnal(conn, tanggal, keterangan, 'OPERASIONAL', 'RETUR_BELI', entries)
            # Inherit proyek_id dari pembelian sumber (via tracker hutang) supaya biaya proyek konsisten.
            if tracker and tracker['jurnal_id']:
                src_pid = conn.execute("SELECT proyek_id FROM jurnal WHERE id=?",
                                       (tracker['jurnal_id'],)).fetchone()
                if src_pid and src_pid['proyek_id']:
                    conn.execute("UPDATE jurnal SET proyek_id=? WHERE id=?",
                                 (src_pid['proyek_id'], jid))
            record_non_sku_movement(conn, jid, tanggal, f"Retur pembelian non-SKU: {keterangan}", -tax['dpp'])
            if tracker:
                apply_tracker_adjustment(conn, 'HUTANG', tracker['id'], -tax['net'], jid)

        else:  # per produk
            produk_ids = request.form.getlist('produk_id[]')
            qtys       = request.form.getlist('qty[]')
            hargas     = request.form.getlist('harga_beli[]')

            items = []
            requested_qty = {}
            for i, pid_str in enumerate(produk_ids):
                if not pid_str: continue
                try:
                    pid   = int(pid_str)
                    qty   = parse_positive_qty(qtys[i] if i < len(qtys) else 1, f'Qty item baris {i + 1}')
                    harga = parse_nonnegative_rp(hargas[i] if i < len(hargas) else '0',
                                                 f'Harga beli item baris {i + 1}')
                except (ValueError, TypeError):
                    return _reject_form(conn, f'Rincian produk baris {i + 1} tidak valid.', 'retur_pembelian')
                produk = conn.execute(
                    "SELECT * FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
                    (pid,)
                ).fetchone()
                if not produk:
                    return _reject_form(conn, f'Produk pada baris {i + 1} tidak ditemukan atau sudah diarsipkan.', 'retur_pembelian')
                if harga <= 0:
                    return _reject_form(conn, f'Harga beli item baris {i + 1} harus lebih dari 0.',
                                        'retur_pembelian')
                requested_qty[pid] = requested_qty.get(pid, 0) + qty
                if requested_qty[pid] > produk['stok']:
                    conn.close()
                    flash(f'Stok {produk["nama"]} tidak cukup untuk retur ({produk["stok"]} < {requested_qty[pid]}).', 'danger')
                    return redirect(url_for('retur_pembelian'))
                items.append({'produk': produk, 'qty': qty, 'harga': harga,
                              'subtotal': qty * harga})

            total_sub = sum(it['subtotal'] for it in items)
            if total_sub <= 0 or not items:
                conn.close()
                flash('Pilih minimal 1 produk dengan nominal > 0.', 'danger')
                return redirect(url_for('retur_pembelian'))
            pemasok = request.form.get('pemasok', '').strip()
            try:
                tracker = _select_return_tracker(conn, 'HUTANG', request.form.get('hutang_id'), pemasok, total_sub) if sumber == 'hutang' else None
            except ValueError as ex:
                return _reject_form(conn, str(ex), 'retur_pembelian')
            if sumber == 'hutang' and not tracker:
                conn.close()
                flash('Hutang aktif dengan sisa yang cukup tidak ditemukan. Periksa pemasok atau nominal retur.', 'danger')
                return redirect(url_for('retur_pembelian'))
            akun_hutang_kode = (tracker['akun_kode'] or '2100') if tracker else '2100'
            tax = _tracker_return_tax_parts(conn, 'HUTANG', tracker, total_sub, 'DPP') if tracker else {
                'dpp': total_sub, 'ppn': 0.0, 'pph': 0.0, 'net': total_sub
            }
            if tracker and tax['net'] > (tracker['jumlah'] - tracker['terbayar']) + 0.01:
                return _reject_form(conn, f'Nominal retur melebihi sisa hutang Rp {tracker["jumlah"] - tracker["terbayar"]:,.0f}.', 'retur_pembelian')

            if sumber == 'hutang':
                entries = [(akun_hutang_kode, tax['net'], 0)]
            else:
                entries = [(akun_kas_kode, tax['net'], 0)]
            if tax['pph'] > 0:
                entries.append(('2112', tax['pph'], 0))
            entries.append(('1130', 0, tax['dpp']))
            if tax['ppn'] > 0:
                entries.append(('1180', 0, tax['ppn']))

            jid = insert_jurnal(conn, tanggal, keterangan, 'OPERASIONAL', 'RETUR_BELI', entries)
            # Inherit proyek_id dari pembelian sumber (via tracker hutang) supaya biaya proyek konsisten.
            if tracker and tracker['jurnal_id']:
                src_pid = conn.execute("SELECT proyek_id FROM jurnal WHERE id=?",
                                       (tracker['jurnal_id'],)).fetchone()
                if src_pid and src_pid['proyek_id']:
                    conn.execute("UPDATE jurnal SET proyek_id=? WHERE id=?",
                                 (src_pid['proyek_id'], jid))

            # Stok keluar (kembali ke supplier)
            return_items = []
            for item in items:
                current = conn.execute("SELECT stok,harga_beli FROM produk WHERE id=?",
                                       (item['produk']['id'],)).fetchone()
                stok_sebelum = current['stok']
                hpp_sebelum = current['harga_beli']
                remove_from_average_cost(conn, item['produk']['id'], item['qty'], item['harga'])
                conn.execute("UPDATE produk SET stok=stok-? WHERE id=?",
                             (item['qty'], item['produk']['id']))
                record_stock_movement(conn, item['produk']['id'], tanggal, 'KELUAR',
                                      item['qty'], item['harga'], f"Retur ke supplier: {keterangan}",
                                      jid, 'REMOVE_LAYER', stok_sebelum, hpp_sebelum)
                return_items.append({'produk_id': item['produk']['id'], 'deskripsi': item['produk']['nama'],
                                     'qty': item['qty'], 'satuan': item['produk']['satuan'],
                                     'harga': item['harga'], 'subtotal': item['subtotal'], 'arah': 'KELUAR'})
            _store_items(conn, jid, return_items)
            if tracker:
                apply_tracker_adjustment(conn, 'HUTANG', tracker['id'], -tax['net'], jid)

        add_log(conn, 'Retur Pembelian', f"{keterangan} | {tanggal}", 'INPUT')
        conn.commit(); conn.close()
        flash('Retur pembelian berhasil dicatat!', 'success')
        return redirect(url_for('retur_pembelian'))

    saldo_persediaan = persediaan_ledger_value(conn)
    hutang_aktif = conn.execute(
        """SELECT *, (jumlah-terbayar) AS sisa
           FROM hutang
           WHERE status!='LUNAS'
           ORDER BY tanggal DESC,id DESC"""
    ).fetchall()
    vendor_names = [r['nama'] for r in conn.execute("SELECT nama FROM vendor ORDER BY nama").fetchall()]
    conn.close()
    return render_template('retur_pembelian.html',
                           produk_list=produk_list, akun_kas=akun_kas,
                           today=date.today().strftime('%Y-%m-%d'),
                           saldo_persediaan=saldo_persediaan,
                           hutang_aktif=hutang_aktif,
                           vendor_names=vendor_names)


# ---------- PELUNASAN ----------
@app.route('/pelunasan', methods=['GET','POST'])
@finance_required
def pelunasan():
    conn = db()
    piutang_aktif = conn.execute(
        "SELECT * FROM piutang WHERE status!='LUNAS' ORDER BY jatuh_tempo, tanggal"
    ).fetchall()
    hutang_aktif = conn.execute(
        "SELECT * FROM hutang WHERE status!='LUNAS' ORDER BY jatuh_tempo, tanggal"
    ).fetchall()
    # Lampirkan sumber dokumen (invoice asal / transaksi asal) ke tiap baris utk label dropdown
    def _enrich_sumber(rows, with_invoice):
        jur_ids = list({r['jurnal_id'] for r in rows if r['jurnal_id']})
        tx_map = {}
        if jur_ids:
            ph = ','.join('?' * len(jur_ids))
            for jr in conn.execute(f"SELECT id,nomor_tx FROM jurnal WHERE id IN ({ph})", jur_ids):
                tx_map[jr['id']] = jr['nomor_tx']
        inv_map = {}
        if with_invoice:
            inv_ids = list({r['invoice_id'] for r in rows if r['invoice_id']})
            if inv_ids:
                ph = ','.join('?' * len(inv_ids))
                for ir in conn.execute(f"SELECT id,nomor FROM invoice WHERE id IN ({ph})", inv_ids):
                    inv_map[ir['id']] = ir['nomor']
        out = []
        for r in rows:
            d = dict(r)
            d['tx_nomor'] = tx_map.get(r['jurnal_id'])
            d['inv_nomor'] = inv_map.get(r['invoice_id']) if with_invoice else None
            out.append(d)
        return out
    piutang_aktif = _enrich_sumber(piutang_aktif, True)
    hutang_aktif = _enrich_sumber(hutang_aktif, False)
    akun_kas = conn.execute("SELECT * FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    # Saldo pinjaman karyawan outstanding (ledger-derived per pihak, akun 1160 Piutang Karyawan)
    pinjaman_aktif = conn.execute(
        """SELECT COALESCE(NULLIF(TRIM(j.pihak),''),'Tanpa Nama') AS nama,
                  COALESCE(SUM(dj.debit - dj.kredit),0) AS sisa
           FROM detail_jurnal dj
           JOIN akun a ON a.id = dj.akun_id
           JOIN jurnal j ON j.id = dj.jurnal_id
           WHERE a.kode='1160'
           GROUP BY 1 HAVING sisa > 0.005 ORDER BY nama"""
    ).fetchall()

    if request.method == 'POST':
        jenis   = request.form['jenis']
        if jenis not in ('PIUTANG', 'HUTANG', 'PINJAMAN_KARYAWAN'):
            return _reject_form(conn, 'Jenis pelunasan tidak valid.', 'pelunasan')

        if jenis == 'PINJAMAN_KARYAWAN':
            # Pengembalian pinjaman/kasbon karyawan: Dr kas / Cr 1160 Piutang Karyawan.
            # Saldo per karyawan diturunkan dari ledger via jurnal.pihak (tanpa tabel tracker),
            # jadi jurnal pengembalian WAJIB mengisi pihak dgn nama yang sama persis.
            nama_kar = (request.form.get('pinjaman_nama','') or '').strip()
            if not nama_kar:
                return _reject_form(conn, 'Pilih karyawan yang mengembalikan pinjaman.', 'pelunasan')
            try:
                nominal = parse_nonnegative_rp(request.form.get('nominal',0), 'Nominal pengembalian')
            except (ValueError, TypeError):
                return _reject_form(conn, 'Nominal pengembalian tidak valid.', 'pelunasan')
            tanggal = request.form.get('tanggal', str(date.today()))
            akun_kas_kode = request.form.get('akun_kas','1100')
            if not _operator_date_ok(tanggal):
                return _reject_form(conn, 'Tanggal pelunasan tidak valid.', 'pelunasan')
            if not is_rekening_kode(conn, akun_kas_kode):
                return _reject_form(conn, 'Rekening kas/bank tidak valid.', 'pelunasan')
            row = conn.execute(
                """SELECT COALESCE(SUM(dj.debit - dj.kredit),0) AS sisa
                   FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id
                   JOIN jurnal j ON j.id=dj.jurnal_id
                   WHERE a.kode='1160' AND COALESCE(NULLIF(TRIM(j.pihak),''),'Tanpa Nama')=?""",
                (nama_kar,)
            ).fetchone()
            sisa = float(row['sisa'] or 0)
            if sisa <= 0.005:
                conn.close(); flash('Pinjaman karyawan ini sudah lunas.', 'warning')
                return redirect(url_for('pelunasan'))
            if nominal <= 0:
                return _reject_form(conn, 'Nominal pengembalian harus lebih dari 0.', 'pelunasan')
            if nominal > sisa + 0.01:
                return _reject_form(conn, f'Nominal melebihi sisa pinjaman Rp {sisa:,.0f}.', 'pelunasan')
            # Keterangan JANGAN berprefix "Pelunasan piutang: " — prefix itu dipakai
            # backfill_payment_journal_links utk mencocokkan bayar_piutang lama.
            insert_jurnal(conn, tanggal, f"Pengembalian pinjaman karyawan: {nama_kar}",
                          'OPERASIONAL', 'PELUNASAN',
                          [(akun_kas_kode, nominal, 0), ('1160', 0, nominal)], pihak=nama_kar)
            add_log(conn, 'Pengembalian pinjaman karyawan',
                    f"{nama_kar} | Rp {nominal:,.0f} | {tanggal}", 'INPUT')
            conn.commit(); conn.close()
            flash('Pengembalian pinjaman karyawan dicatat!', 'success')
            return redirect(url_for('pelunasan'))
        # getlist karena ada 2 select dengan name=record_id (piutang + hutang)
        rid_candidates = [v.strip() for v in request.form.getlist('record_id') if v.strip()]
        if not rid_candidates:
            conn.close()
            flash('Pilih hutang atau piutang terlebih dahulu.', 'warning')
            return redirect(url_for('pelunasan'))
        rid_str = rid_candidates[-1]  # ambil yang terakhir (select yang visible)
        try:
            rid = int(rid_str)
            nominal = parse_nonnegative_rp(request.form.get('nominal',0), 'Nominal pelunasan')
        except (ValueError, TypeError):
            return _reject_form(conn, 'Rincian pelunasan tidak valid.', 'pelunasan')
        tanggal= request.form.get('tanggal', str(date.today()))
        akun_kas_kode = request.form.get('akun_kas','1100')
        catatan= request.form.get('catatan','')
        if not _operator_date_ok(tanggal):
            return _reject_form(conn, 'Tanggal pelunasan tidak valid.', 'pelunasan')
        if not is_rekening_kode(conn, akun_kas_kode):
            return _reject_form(conn, 'Rekening kas/bank tidak valid.', 'pelunasan')

        if jenis == 'PIUTANG':
            p = conn.execute("SELECT * FROM piutang WHERE id=?", (rid,)).fetchone()
            if not p:
                return _reject_form(conn, 'Piutang tidak ditemukan.', 'pelunasan')
            sisa = p['jumlah'] - p['terbayar']
            if nominal > sisa + 0.01:
                return _reject_form(conn, f'Nominal pelunasan melebihi sisa piutang Rp {sisa:,.0f}.', 'pelunasan')
            if nominal <= 0:
                conn.close(); flash('Piutang sudah lunas atau jumlah tidak valid.', 'warning')
                return redirect(url_for('pelunasan'))
            try:
                payment_jid = insert_jurnal(conn, tanggal, f"Pelunasan piutang: {p['pelanggan']}", 'OPERASIONAL', 'PELUNASAN', [
                    (akun_kas_kode, nominal, 0), ('1120', 0, nominal)
                ])
            except ValueError as ex:
                conn.rollback(); conn.close()
                flash(f'Gagal mencatat pelunasan piutang: {ex}', 'danger')
                return redirect(url_for('pelunasan'))
            conn.execute("INSERT INTO bayar_piutang(piutang_id,tanggal,jumlah,catatan,jurnal_id) VALUES(?,?,?,?,?)",
                         (rid, tanggal, nominal, catatan, payment_jid))
            conn.execute("UPDATE piutang SET terbayar=terbayar+? WHERE id=?", (nominal, rid))
            update_piutang_status(conn, rid)
            _sync_invoice_from_piutang(conn, rid)
            add_log(conn, 'Pelunasan piutang', f"{p['pelanggan']} | Rp {nominal:,.0f} | {tanggal}", 'INPUT')
            flash('Pelunasan piutang dicatat!', 'success')
        else:
            h = conn.execute("SELECT * FROM hutang WHERE id=?", (rid,)).fetchone()
            if not h:
                return _reject_form(conn, 'Hutang tidak ditemukan.', 'pelunasan')
            sisa = h['jumlah'] - h['terbayar']
            if nominal > sisa + 0.01:
                return _reject_form(conn, f'Nominal pelunasan melebihi sisa hutang Rp {sisa:,.0f}.', 'pelunasan')
            if nominal <= 0:
                conn.close(); flash('Hutang sudah lunas atau jumlah tidak valid.', 'warning')
                return redirect(url_for('pelunasan'))
            try:
                payment_jid = insert_jurnal(conn, tanggal, f"Pelunasan hutang: {h['pemasok']}", 'OPERASIONAL', 'PELUNASAN', [
                    (h['akun_kode'] or '2100', nominal, 0), (akun_kas_kode, 0, nominal)
                ])
            except ValueError as ex:
                conn.rollback(); conn.close()
                flash(f'Gagal mencatat pelunasan hutang: {ex}. Pastikan akun COA untuk hutang ini masih ada di Bagan Akun.', 'danger')
                return redirect(url_for('pelunasan'))
            conn.execute("INSERT INTO bayar_hutang(hutang_id,tanggal,jumlah,catatan,jurnal_id) VALUES(?,?,?,?,?)",
                         (rid, tanggal, nominal, catatan, payment_jid))
            conn.execute("UPDATE hutang SET terbayar=terbayar+? WHERE id=?", (nominal, rid))
            update_hutang_status(conn, rid)
            add_log(conn, 'Pelunasan hutang', f"{h['pemasok']} | Rp {nominal:,.0f} | {tanggal}", 'INPUT')
            flash('Pelunasan hutang dicatat!', 'success')

        conn.commit(); conn.close()
        return redirect(url_for('pelunasan'))

    conn.close()
    return render_template('pelunasan.html',
                           piutang_aktif=piutang_aktif, hutang_aktif=hutang_aktif,
                           pinjaman_aktif=pinjaman_aktif,
                           akun_kas=akun_kas, today=date.today().strftime('%Y-%m-%d'))


# ---------- PIUTANG ----------
@app.route('/piutang')
@investor_required
def piutang_list():
    conn = db()
    today = date.today()
    sf = request.args.get('status','')
    q = request.args.get('q','')
    where, params = [], []
    if sf: where.append("status=?"); params.append(sf)
    if q:  where.append("(pelanggan LIKE ? OR keterangan LIKE ?)"); params += [f'%{q}%',f'%{q}%']
    ws = ('WHERE '+' AND '.join(where)) if where else ''
    page = _page_arg()
    n_rows = conn.execute(f'SELECT COUNT(*) FROM piutang {ws}', params).fetchone()[0]
    pg = paginate_meta(page, n_rows)
    rows = conn.execute(
        f'SELECT * FROM piutang {ws} ORDER BY CASE status WHEN "LUNAS" THEN 1 ELSE 0 END, jatuh_tempo, tanggal DESC LIMIT ? OFFSET ?',
        params + [pg['per_page'], pg['offset']]).fetchall()
    total = conn.execute("SELECT COALESCE(SUM(jumlah-terbayar),0) FROM piutang WHERE status!='LUNAS'").fetchone()[0]
    # KPI dihitung via SQL (lintas semua halaman, bukan cuma halaman ini)
    today_s = today.strftime('%Y-%m-%d')
    soon_s = (today + timedelta(days=7)).strftime('%Y-%m-%d')
    n_overdue = conn.execute(
        "SELECT COUNT(*) FROM piutang WHERE status!='LUNAS' AND jatuh_tempo IS NOT NULL "
        "AND jatuh_tempo!='' AND substr(jatuh_tempo,1,10) < ?", (today_s,)).fetchone()[0]
    n_duesoon = conn.execute(
        "SELECT COUNT(*) FROM piutang WHERE status!='LUNAS' AND jatuh_tempo IS NOT NULL AND jatuh_tempo!='' "
        "AND substr(jatuh_tempo,1,10) >= ? AND substr(jatuh_tempo,1,10) <= ?", (today_s, soon_s)).fetchone()[0]
    writeoff_rows = conn.execute(
        "SELECT piutang_id, COALESCE(SUM(jumlah),0) AS jumlah FROM writeoff_piutang GROUP BY piutang_id"
    ).fetchall()
    writeoff_map = {r['piutang_id']: float(r['jumlah'] or 0) for r in writeoff_rows}
    # Sumber dokumen tiap baris (invoice asal / transaksi asal) -- halaman ini saja
    inv_ids = list({r['invoice_id'] for r in rows if r['invoice_id']})
    inv_map = {}
    if inv_ids:
        ph = ','.join('?' * len(inv_ids))
        for ir in conn.execute(f"SELECT id,nomor FROM invoice WHERE id IN ({ph})", inv_ids):
            inv_map[ir['id']] = ir['nomor']
    jur_ids = list({r['jurnal_id'] for r in rows if r['jurnal_id']})
    tx_map = {}
    saldo_awal_jids = set()
    if jur_ids:
        ph = ','.join('?' * len(jur_ids))
        for jr in conn.execute(f"SELECT id,nomor_tx,tipe_tx FROM jurnal WHERE id IN ({ph})", jur_ids):
            tx_map[jr['id']] = jr['nomor_tx']
            # Baris yang lahir dari Setup Saldo Awal ditandai supaya bisa dibedakan
            # dari tagihan transaksi biasa (bawaan pembukuan lama, bukan belanja baru).
            if jr['tipe_tx'] == 'SALDO_AWAL_AKUN':
                saldo_awal_jids.add(jr['id'])
    rows_with_status = []
    for r in rows:
        row = dict(r)
        row['writeoff_jumlah'] = writeoff_map.get(r['id'], 0)
        row['inv_nomor'] = inv_map.get(r['invoice_id'])
        row['tx_nomor'] = tx_map.get(r['jurnal_id'])
        row['dari_saldo_awal'] = r['jurnal_id'] in saldo_awal_jids
        lbl, cls = piutang_status_label(r, today)
        if row['writeoff_jumlah'] > 0 and row['jumlah'] - row['terbayar'] <= 0.01:
            lbl, cls = 'DIHAPUS / WRITE OFF', 'dark'
        rows_with_status.append((row, lbl, cls))
    akun_kas = conn.execute("SELECT kode,nama FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    conn.close()
    return render_template('piutang.html', rows=rows_with_status, total=total, sf=sf, q=q,
                           akun_kas=akun_kas, today=date.today(), pg=pg,
                           n_overdue=n_overdue, n_duesoon=n_duesoon)

@app.route('/piutang/baru', methods=['POST'])
@operator_required
def piutang_baru():
    flash('Piutang tidak bisa ditambah langsung. Catat melalui form Pemasukan agar jurnal terbentuk.', 'warning')
    return redirect(url_for('piutang_list'))

@app.route('/piutang/<int:id>/edit', methods=['GET','POST'])
@operator_required
def piutang_edit(id):
    conn = db()
    p = conn.execute('SELECT * FROM piutang WHERE id=?', (id,)).fetchone()
    if not p:
        conn.close(); flash('Tidak ditemukan.','danger')
        return redirect(url_for('piutang_list'))
    if request.method == 'POST':
        tanggal = request.form['tanggal']
        jt = request.form.get('jatuh_tempo') or None
        pelanggan = request.form['pelanggan'].strip()
        keterangan = request.form.get('keterangan','')
        if not _operator_date_ok(tanggal):
            conn.close()
            flash('Operator hanya bisa mengedit piutang untuk bulan berjalan.', 'warning')
            return redirect(url_for('piutang_list'))
        if not is_valid_optional_iso_date(jt):
            conn.close()
            flash('Tanggal jatuh tempo piutang tidak valid.', 'danger')
            return redirect(url_for('piutang_list'))
        if (tracker_has_structured_source(conn, p['jurnal_id'])
                and (tanggal != p['tanggal']
                     or pelanggan != (p['pelanggan'] or '').strip()
                     or keterangan != (p['keterangan'] or ''))):
            conn.close()
            flash('Tanggal, pelanggan, dan keterangan piutang ini mengikuti transaksi sumber. Ubah melalui menu Edit Transaksi & Item.', 'warning')
            return redirect(url_for('piutang_list'))
        conn.execute("""UPDATE piutang SET tanggal=?,jatuh_tempo=?,pelanggan=?,keterangan=? WHERE id=?""",
            (tanggal, jt, pelanggan, keterangan, id))
        sync_tracker_due_date(conn, p['jurnal_id'], jt)
        conn.commit(); conn.close()
        flash('Piutang diperbarui!', 'success')
        return redirect(url_for('piutang_list'))
    conn.close()
    return render_template('piutang_edit.html', p=p)

@app.route('/piutang/<int:id>/bayar', methods=['POST'])
@operator_required
def piutang_bayar(id):
    conn = db()
    try:
        jumlah = parse_nonnegative_rp(request.form.get('jumlah','0'), 'Jumlah pembayaran')
    except ValueError as ex:
        return _reject_form(conn, str(ex), 'piutang_list')
    tanggal = request.form.get('tanggal', str(date.today()))
    catatan = request.form.get('catatan','')
    akun_kas_kode = request.form.get('akun_kas','1100')
    if not _operator_date_ok(tanggal):
        conn.close()
        flash('Tanggal pembayaran tidak valid. Operator hanya bisa mencatat pembayaran untuk bulan berjalan.', 'warning')
        return redirect(url_for('piutang_list'))
    if not is_rekening_kode(conn, akun_kas_kode):
        return _reject_form(conn, 'Rekening kas/bank tidak valid.', 'piutang_list')
    p = conn.execute('SELECT * FROM piutang WHERE id=?', (id,)).fetchone()
    if not p:
        conn.close(); flash('Piutang tidak ditemukan.', 'danger')
        return redirect(url_for('piutang_list'))
    sisa = p['jumlah'] - p['terbayar']
    if jumlah > sisa + 0.01:
        return _reject_form(conn, f'Jumlah pembayaran melebihi sisa piutang Rp {sisa:,.0f}.', 'piutang_list')
    if jumlah <= 0:
        conn.close(); flash('Piutang sudah lunas atau jumlah tidak valid.', 'warning')
        return redirect(url_for('piutang_list'))
    payment_jid = insert_jurnal(conn, tanggal, f"Pelunasan piutang: {p['pelanggan']}", 'OPERASIONAL', 'PELUNASAN', [
        (akun_kas_kode, jumlah, 0), ('1120', 0, jumlah)
    ])
    conn.execute('INSERT INTO bayar_piutang(piutang_id,tanggal,jumlah,catatan,jurnal_id) VALUES(?,?,?,?,?)',
                 (id, tanggal, jumlah, catatan, payment_jid))
    conn.execute('UPDATE piutang SET terbayar=terbayar+? WHERE id=?', (jumlah, id))
    update_piutang_status(conn, id)
    _sync_invoice_from_piutang(conn, id)
    add_log(conn, 'Pelunasan piutang', f"{p['pelanggan']} | Rp {jumlah:,.0f} | {tanggal}", 'INPUT')
    conn.commit(); conn.close()
    flash('Pembayaran piutang dicatat!', 'success')
    return redirect(url_for('piutang_list'))

@app.route('/piutang/<int:id>/writeoff', methods=['POST'])
@finance_required
def piutang_writeoff(id):
    conn = db()
    p = conn.execute('SELECT * FROM piutang WHERE id=?', (id,)).fetchone()
    if not p:
        conn.close(); flash('Piutang tidak ditemukan.', 'danger')
        return redirect(url_for('piutang_list'))
    tanggal = request.form.get('tanggal', str(date.today()))
    metode = (request.form.get('metode') or 'LANGSUNG').upper()
    alasan = (request.form.get('alasan') or '').strip()
    if not is_valid_iso_date(tanggal):
        conn.close(); flash('Tanggal write-off piutang tidak valid.', 'danger')
        return redirect(url_for('piutang_list'))
    if metode not in ('LANGSUNG', 'CADANGAN'):
        conn.close(); flash('Metode write-off piutang tidak valid.', 'danger')
        return redirect(url_for('piutang_list'))
    try:
        jumlah = parse_nonnegative_rp(request.form.get('jumlah','0'), 'Jumlah write-off')
    except ValueError as ex:
        return _reject_form(conn, str(ex), 'piutang_list')
    sisa = float(p['jumlah'] or 0) - float(p['terbayar'] or 0)
    if jumlah <= 0:
        jumlah = sisa
    if jumlah <= 0.01:
        conn.close(); flash('Piutang sudah tidak memiliki sisa saldo.', 'warning')
        return redirect(url_for('piutang_list'))
    if jumlah > sisa + 0.01:
        return _reject_form(conn, f'Jumlah write-off melebihi sisa piutang Rp {sisa:,.0f}.', 'piutang_list')

    if metode == 'CADANGAN':
        entries = [('1150', jumlah, 0), ('1120', 0, jumlah)]
        ket_metode = 'pakai cadangan'
    else:
        entries = [('6190', jumlah, 0), ('1120', 0, jumlah)]
        ket_metode = 'langsung ke beban'
    jid = insert_jurnal(
        conn, tanggal, f"Write-off piutang: {p['pelanggan']} ({ket_metode})",
        'OPERASIONAL', 'WRITE_OFF_PIUTANG', entries, pihak=p['pelanggan']
    )
    conn.execute(
        "INSERT INTO writeoff_piutang(piutang_id,jurnal_id,tanggal,jumlah,metode,alasan) VALUES(?,?,?,?,?,?)",
        (id, jid, tanggal, jumlah, metode, alasan)
    )
    conn.execute('UPDATE piutang SET terbayar=MIN(jumlah,terbayar+?) WHERE id=?', (jumlah, id))
    update_piutang_status(conn, id)
    add_log(conn, 'Write-off piutang', f"{p['pelanggan']} | Rp {jumlah:,.0f} | {metode} | {tanggal}", 'INPUT')
    conn.commit(); conn.close()
    flash('Piutang di-write-off dan jurnal kerugian piutang dibuat.', 'warning')
    return redirect(url_for('piutang_list'))

@app.route('/piutang/<int:id>/riwayat')
@investor_required
def piutang_riwayat(id):
    conn = db()
    p = conn.execute('SELECT * FROM piutang WHERE id=?', (id,)).fetchone()
    if not p:
        conn.close(); flash('Piutang tidak ditemukan.', 'danger')
        return redirect(url_for('piutang_list'))
    riwayat = conn.execute('SELECT * FROM bayar_piutang WHERE piutang_id=? ORDER BY tanggal DESC', (id,)).fetchall()
    conn.close()
    return render_template('piutang_riwayat.html', p=p, riwayat=riwayat)

@app.route('/piutang/<int:id>/hapus', methods=['POST'])
@finance_required
def piutang_hapus(id):
    conn = db()
    p = conn.execute('SELECT * FROM piutang WHERE id=?', (id,)).fetchone()
    if p:
        if tracker_has_structured_source(conn, p['jurnal_id']):
            conn.close()
            flash('Piutang ini berasal dari transaksi sumber. Hapus atau edit transaksi sumber agar jurnal, invoice, dan stok tetap sinkron.', 'warning')
            return redirect(url_for('piutang_list'))
        related = conn.execute(
            """SELECT (SELECT COUNT(*) FROM bayar_piutang WHERE piutang_id=?)
                    + (SELECT COUNT(*) FROM penyesuaian_tagihan WHERE jenis='PIUTANG' AND record_id=?)
                    + (SELECT COUNT(*) FROM writeoff_piutang WHERE piutang_id=?)""",
            (id, id, id)
        ).fetchone()[0]
        if related:
            conn.close()
            flash('Piutang memiliki pembayaran, retur, atau write-off terkait. Hapus jurnal terkait lebih dulu sebelum menghapus piutang.', 'warning')
            return redirect(url_for('piutang_list'))
        sisa_piutang = p['jumlah'] - p['terbayar']
        # Reverse jurnal sisa piutang yang belum terbayar: Dr 4100 (undo revenue), Cr 1120 (hapus aset).
        # Pakai tanggal piutang asli (bukan hari ini) supaya pembatalan revenue net di periode
        # yang sama dengan pengakuannya — tidak menggerus laba periode berjalan secara semu.
        if sisa_piutang > 0:
            tgl_koreksi = p['tanggal'] if is_valid_iso_date(p['tanggal']) else str(date.today())
            insert_jurnal(conn, tgl_koreksi, f"Koreksi hapus piutang: {p['pelanggan']}",
                          'KOREKSI', 'KOREKSI', [
                ('4100', sisa_piutang, 0),   # Dr Pendapatan, batalkan revenue yang belum diterima
                ('1120', 0, sisa_piutang),   # Cr Piutang, hapus saldo aset piutang
            ])
        conn.execute('DELETE FROM piutang WHERE id=?', (id,))
    conn.commit(); conn.close()
    flash('Piutang dihapus dan jurnal dikoreksi.', 'warning')
    return redirect(url_for('piutang_list'))


# ---------- HUTANG ----------
@app.route('/hutang')
@investor_required
def hutang_list():
    conn = db()
    today = date.today()
    sf = request.args.get('status','')
    q = request.args.get('q','')
    where, params = [], []
    if sf: where.append("status=?"); params.append(sf)
    if q:  where.append("(pemasok LIKE ? OR keterangan LIKE ?)"); params += [f'%{q}%',f'%{q}%']
    ws = ('WHERE '+' AND '.join(where)) if where else ''
    page = _page_arg()
    n_rows = conn.execute(f'SELECT COUNT(*) FROM hutang {ws}', params).fetchone()[0]
    pg = paginate_meta(page, n_rows)
    rows = conn.execute(
        f'SELECT * FROM hutang {ws} ORDER BY CASE status WHEN "LUNAS" THEN 1 ELSE 0 END, jatuh_tempo, tanggal DESC LIMIT ? OFFSET ?',
        params + [pg['per_page'], pg['offset']]).fetchall()
    total = conn.execute("SELECT COALESCE(SUM(jumlah-terbayar),0) FROM hutang WHERE status!='LUNAS'").fetchone()[0]
    # KPI jatuh tempo dihitung via SQL (lintas semua halaman, bukan cuma halaman ini)
    today_s = today.strftime('%Y-%m-%d')
    soon_s = (today + timedelta(days=7)).strftime('%Y-%m-%d')
    n_overdue = conn.execute(
        "SELECT COUNT(*) FROM hutang WHERE status!='LUNAS' AND jatuh_tempo IS NOT NULL "
        "AND jatuh_tempo!='' AND substr(jatuh_tempo,1,10) < ?", (today_s,)).fetchone()[0]
    n_duesoon = conn.execute(
        "SELECT COUNT(*) FROM hutang WHERE status!='LUNAS' AND jatuh_tempo IS NOT NULL AND jatuh_tempo!='' "
        "AND substr(jatuh_tempo,1,10) >= ? AND substr(jatuh_tempo,1,10) <= ?", (today_s, soon_s)).fetchone()[0]
    jur_ids = list({r['jurnal_id'] for r in rows if r['jurnal_id']})
    tx_map = {}
    saldo_awal_jids = set()
    if jur_ids:
        ph = ','.join('?' * len(jur_ids))
        for jr in conn.execute(f"SELECT id,nomor_tx,tipe_tx FROM jurnal WHERE id IN ({ph})", jur_ids):
            tx_map[jr['id']] = jr['nomor_tx']
            # Baris yang lahir dari Setup Saldo Awal ditandai supaya bisa dibedakan
            # dari tagihan transaksi biasa (bawaan pembukuan lama, bukan belanja baru).
            if jr['tipe_tx'] == 'SALDO_AWAL_AKUN':
                saldo_awal_jids.add(jr['id'])
    rows_with_status = []
    for r in rows:
        row = dict(r)
        row['tx_nomor'] = tx_map.get(r['jurnal_id'])
        row['dari_saldo_awal'] = r['jurnal_id'] in saldo_awal_jids
        lbl, cls = piutang_status_label(r, today)
        rows_with_status.append((row, lbl, cls))
    akun_kas = conn.execute("SELECT kode,nama FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    conn.close()
    return render_template('hutang.html', rows=rows_with_status, total=total, sf=sf, q=q,
                           akun_kas=akun_kas, today=date.today(), pg=pg,
                           n_overdue=n_overdue, n_duesoon=n_duesoon)

@app.route('/hutang/baru', methods=['POST'])
@operator_required
def hutang_baru():
    flash('Hutang tidak bisa ditambah langsung. Catat melalui form Pengeluaran agar jurnal terbentuk.', 'warning')
    return redirect(url_for('hutang_list'))

@app.route('/hutang/<int:id>/edit', methods=['GET','POST'])
@operator_required
def hutang_edit(id):
    conn = db()
    h = conn.execute('SELECT * FROM hutang WHERE id=?', (id,)).fetchone()
    if not h:
        conn.close(); flash('Tidak ditemukan.','danger')
        return redirect(url_for('hutang_list'))
    if request.method == 'POST':
        tanggal = request.form['tanggal']
        jt = request.form.get('jatuh_tempo') or None
        pemasok = request.form['pemasok'].strip()
        keterangan = request.form.get('keterangan','')
        if not _operator_date_ok(tanggal):
            conn.close()
            flash('Operator hanya bisa mengedit hutang untuk bulan berjalan.', 'warning')
            return redirect(url_for('hutang_list'))
        if not is_valid_optional_iso_date(jt):
            conn.close()
            flash('Tanggal jatuh tempo hutang tidak valid.', 'danger')
            return redirect(url_for('hutang_list'))
        if (tracker_has_structured_source(conn, h['jurnal_id'])
                and (tanggal != h['tanggal']
                     or pemasok != (h['pemasok'] or '').strip()
                     or keterangan != (h['keterangan'] or ''))):
            conn.close()
            flash('Tanggal, pemasok, dan keterangan hutang ini mengikuti transaksi sumber. Ubah melalui menu Edit Transaksi & Item.', 'warning')
            return redirect(url_for('hutang_list'))
        conn.execute("""UPDATE hutang SET tanggal=?,jatuh_tempo=?,pemasok=?,keterangan=? WHERE id=?""",
            (tanggal, jt, pemasok, keterangan, id))
        sync_tracker_due_date(conn, h['jurnal_id'], jt)
        conn.commit(); conn.close()
        flash('Hutang diperbarui!', 'success')
        return redirect(url_for('hutang_list'))
    conn.close()
    return render_template('hutang_edit.html', h=h)

@app.route('/hutang/<int:id>/bayar', methods=['POST'])
@operator_required
def hutang_bayar(id):
    conn = db()
    try:
        jumlah = parse_nonnegative_rp(request.form.get('jumlah','0'), 'Jumlah pembayaran')
    except ValueError as ex:
        return _reject_form(conn, str(ex), 'hutang_list')
    tanggal = request.form.get('tanggal', str(date.today()))
    catatan = request.form.get('catatan','')
    akun_kas_kode = request.form.get('akun_kas','1100')
    if not _operator_date_ok(tanggal):
        conn.close()
        flash('Tanggal pembayaran tidak valid. Operator hanya bisa mencatat pembayaran untuk bulan berjalan.', 'warning')
        return redirect(url_for('hutang_list'))
    if not is_rekening_kode(conn, akun_kas_kode):
        return _reject_form(conn, 'Rekening kas/bank tidak valid.', 'hutang_list')
    h = conn.execute('SELECT * FROM hutang WHERE id=?', (id,)).fetchone()
    if not h:
        conn.close(); flash('Hutang tidak ditemukan.', 'danger')
        return redirect(url_for('hutang_list'))
    sisa = h['jumlah'] - h['terbayar']
    if jumlah > sisa + 0.01:
        return _reject_form(conn, f'Jumlah pembayaran melebihi sisa hutang Rp {sisa:,.0f}.', 'hutang_list')
    if jumlah <= 0:
        conn.close(); flash('Hutang sudah lunas atau jumlah tidak valid.', 'warning')
        return redirect(url_for('hutang_list'))
    try:
        payment_jid = insert_jurnal(conn, tanggal, f"Pelunasan hutang: {h['pemasok']}", 'OPERASIONAL', 'PELUNASAN', [
            (h['akun_kode'] or '2100', jumlah, 0), (akun_kas_kode, 0, jumlah)
        ])
    except ValueError as ex:
        conn.rollback(); conn.close()
        flash(str(ex), 'danger')
        return redirect(url_for('hutang_list'))
    conn.execute('INSERT INTO bayar_hutang(hutang_id,tanggal,jumlah,catatan,jurnal_id) VALUES(?,?,?,?,?)',
                 (id, tanggal, jumlah, catatan, payment_jid))
    conn.execute('UPDATE hutang SET terbayar=terbayar+? WHERE id=?', (jumlah, id))
    update_hutang_status(conn, id)
    add_log(conn, 'Pelunasan hutang', f"{h['pemasok']} | Rp {jumlah:,.0f} | {tanggal}", 'INPUT')
    conn.commit(); conn.close()
    flash('Pembayaran hutang dicatat!', 'success')
    return redirect(url_for('hutang_list'))

@app.route('/hutang/<int:id>/riwayat')
@investor_required
def hutang_riwayat(id):
    conn = db()
    h = conn.execute('SELECT * FROM hutang WHERE id=?', (id,)).fetchone()
    if not h:
        conn.close(); flash('Hutang tidak ditemukan.', 'danger')
        return redirect(url_for('hutang_list'))
    riwayat = conn.execute('SELECT * FROM bayar_hutang WHERE hutang_id=? ORDER BY tanggal DESC', (id,)).fetchall()
    conn.close()
    return render_template('hutang_riwayat.html', h=h, riwayat=riwayat)

@app.route('/hutang/<int:id>/hapus', methods=['POST'])
@finance_required
def hutang_hapus(id):
    conn = db()
    h = conn.execute('SELECT * FROM hutang WHERE id=?', (id,)).fetchone()
    if h:
        if tracker_has_structured_source(conn, h['jurnal_id']):
            conn.close()
            flash('Hutang ini berasal dari transaksi sumber. Hapus atau edit transaksi sumber agar jurnal, stok, dan tagihan tetap sinkron.', 'warning')
            return redirect(url_for('hutang_list'))
        related = conn.execute(
            """SELECT (SELECT COUNT(*) FROM bayar_hutang WHERE hutang_id=?)
                    + (SELECT COUNT(*) FROM penyesuaian_tagihan WHERE jenis='HUTANG' AND record_id=?)""",
            (id, id)
        ).fetchone()[0]
        if related:
            conn.close()
            flash('Hutang memiliki pembayaran atau retur terkait. Hapus jurnal terkait lebih dulu sebelum menghapus hutang.', 'warning')
            return redirect(url_for('hutang_list'))
        sisa_hutang = h['jumlah'] - h['terbayar']
        # Hapus sisa hutang: Dr 2100 (hapus liability), Cr 4300 (pendapatan lain, hutang dibebaskan)
        if sisa_hutang > 0:
            insert_jurnal(conn, str(date.today()), f"Koreksi hapus hutang: {h['pemasok']}",
                          'KOREKSI', 'KOREKSI', [
                (h['akun_kode'] or '2100', sisa_hutang, 0),  # Dr akun hutang terkait
                ('4300', 0, sisa_hutang),  # Cr Pendapatan Lain-lain, hutang yang dibebaskan
            ])
        conn.execute('DELETE FROM hutang WHERE id=?', (id,))
    conn.commit(); conn.close()
    flash('Hutang dihapus dan jurnal dikoreksi.', 'warning')
    return redirect(url_for('hutang_list'))


# ---------- INVENTORY / STOK ----------
def _kode_natural_key(s):
    """Kunci sorting 'natural' untuk kode barang: pecah bagian angka jadi integer
    supaya urut secara manusiawi, mis. KODE-2 < KODE-10 (bukan urutan teks yang
    menaruh '10' sebelum '2'). Bagian huruf dibandingkan case-insensitive. re.split
    selalu mulai dengan potongan non-angka (boleh kosong), jadi posisi ganjil/genap
    konsisten str/int dan aman dibandingkan antar-kode."""
    import re
    return [int(t) if t.isdigit() else t.lower()
            for t in re.split(r'(\d+)', str(s or ''))]


@app.route('/stok')
@inventory_required
def stok_list():
    conn = db()
    q = request.args.get('q','').strip()
    kategori = request.args.get('kategori','').strip()
    show_arsip = request.args.get('arsip') == '1'
    where = ["COALESCE(aktif,1)=?"]
    params = [0 if show_arsip else 1]
    if q:
        # Barcode ikut dicari supaya hasil scan langsung ketemu produknya,
        # sama seperti mengetik kode atau nama.
        where.append("(nama LIKE ? OR kode LIKE ? OR kategori LIKE ? OR COALESCE(barcode,'') LIKE ?)")
        params.extend((f'%{q}%', f'%{q}%', f'%{q}%', f'%{q}%'))
    if kategori == '__tanpa__':
        where.append("TRIM(COALESCE(kategori,''))=''")
    elif kategori:
        where.append("kategori=?")
        params.append(kategori)
    where_sql = ' WHERE ' + ' AND '.join(where)

    sort = request.args.get('sort', 'nama').strip()
    if sort not in ('nama', 'kode', 'stok', 'tgl_beli'):
        sort = 'nama'

    # Ukuran halaman: pilihan tetap (cegah ?per_page=<angka bebas> dipakai iseng
    # buat minta seluruh katalog sekaligus lagi -- itu justru sumber lemotnya).
    STOK_PAGE_SIZES = (50, 100, 500, 1000)
    try:
        per_page = int(request.args.get('per_page', 50))
    except (TypeError, ValueError):
        per_page = 50
    if per_page not in STOK_PAGE_SIZES:
        per_page = 50

    # Tanggal pembelian terakhir per produk (dari mutasi stok MASUK)
    last_beli = {row['produk_id']: row['tgl'] for row in conn.execute(
        "SELECT produk_id, MAX(tanggal) AS tgl FROM pergerakan_stok WHERE jenis='MASUK' GROUP BY produk_id"
    ).fetchall()}

    # ── PAGINASI PER GRUP (satu grup = satu nama produk, varian menyatu di dalamnya).
    # Katalog bisa ratusan ribu SKU (terukur: 190rb SKU -> render 448MB HTML, 11 detik,
    # browser hang), jadi grup halaman ini WAJIB dipilih dulu lewat SQL -- baru ambil
    # baris ANGGOTA grup itu saja -- bukan tarik semua baris lalu dipotong belakangan. ──
    page = _page_arg()
    total_grup = conn.execute(
        f"SELECT COUNT(*) FROM (SELECT nama FROM produk{where_sql} GROUP BY nama)", params
    ).fetchone()[0]
    pg = paginate_meta(page, total_grup, per_page=per_page)

    if sort == 'kode':
        # Natural sort (KODE-2 < KODE-10) tidak bisa dinyatakan lewat ORDER BY SQL biasa.
        # Ambil dulu (nama, kode) SEMUA baris yang lolos filter -- ringan, cuma 2 kolom,
        # bukan SELECT * -- urutkan grup di Python pakai _kode_natural_key (sama persis
        # dgn logika lama), baru potong nama grup utk halaman ini.
        _rows_kode = conn.execute(f"SELECT nama, kode FROM produk{where_sql}", params).fetchall()
        _kode_by_nama = {}
        for r in _rows_kode:
            _kode_by_nama.setdefault(r['nama'], []).append(r['kode'])
        _semua_nama = sorted(
            _kode_by_nama.keys(),
            key=lambda n: min(_kode_natural_key(k) for k in _kode_by_nama[n])
        )
        nama_hal = _semua_nama[pg['offset']:pg['offset'] + pg['per_page']]
    elif sort == 'stok':            # stok terbanyak dulu (total per grup)
        grp_sql = (f"SELECT nama FROM produk{where_sql} GROUP BY nama "
                   "ORDER BY SUM(COALESCE(stok,0)) DESC, LOWER(nama) LIMIT ? OFFSET ?")
        nama_hal = [r['nama'] for r in conn.execute(
            grp_sql, params + [pg['per_page'], pg['offset']]).fetchall()]
    elif sort == 'tgl_beli':        # pembelian terbaru dulu; tanpa pembelian → terakhir
        grp_sql = ("SELECT p.nama AS nama FROM produk p "
                   "LEFT JOIN (SELECT produk_id, MAX(tanggal) tgl FROM pergerakan_stok "
                   "           WHERE jenis='MASUK' GROUP BY produk_id) lb ON lb.produk_id=p.id"
                   f"{where_sql} GROUP BY p.nama "
                   "ORDER BY MAX(COALESCE(lb.tgl,'')) DESC, LOWER(p.nama) LIMIT ? OFFSET ?")
        nama_hal = [r['nama'] for r in conn.execute(
            grp_sql, params + [pg['per_page'], pg['offset']]).fetchall()]
    else:                          # nama A-Z
        grp_sql = (f"SELECT nama FROM produk{where_sql} GROUP BY nama "
                   "ORDER BY LOWER(nama) LIMIT ? OFFSET ?")
        nama_hal = [r['nama'] for r in conn.execute(
            grp_sql, params + [pg['per_page'], pg['offset']]).fetchall()]

    produk = []
    if nama_hal:
        ph = ','.join('?' for _ in nama_hal)
        produk = [dict(r) for r in conn.execute(
            f"SELECT * FROM produk{where_sql} AND nama IN ({ph}) ORDER BY nama, varian",
            params + nama_hal).fetchall()]
    for p in produk:
        p['last_beli'] = last_beli.get(p['id'])
    # Produk yang punya jejak transaksi TIDAK bisa dihapus permanen (cuma diarsipkan),
    # supaya tombol di UI bisa jujur: "Arsipkan" utk yang terpakai, "Hapus" utk yang belum.
    # Dihitung dari tabel LAIN (bukan dari `produk`), jadi tetap benar walau `produk` di
    # sini cuma isi satu halaman -- logikanya sama persis dgn _produk_terpakai() yang
    # dipakai route hapus, jadi label = perilaku.
    terpakai_ids = set()
    for tbl, col in (('pergerakan_stok', 'produk_id'), ('transaksi_item', 'produk_id'),
                     ('po_item', 'produk_id'), ('produksi', 'produk_jadi_id'),
                     ('produksi_bahan', 'produk_id'), ('produksi_template', 'produk_jadi_id')):
        try:
            for r in conn.execute(f"SELECT DISTINCT {col} AS pid FROM {tbl} WHERE {col} IS NOT NULL"):
                terpakai_ids.add(r['pid'])
        except Exception:
            continue
    try:
        for row in conn.execute("SELECT bahan_sku FROM produksi_template").fetchall():
            for b in (json.loads(row['bahan_sku'] or '[]') or []):
                if isinstance(b, dict) and b.get('produk_id'):
                    terpakai_ids.add(int(b['produk_id']))
    except Exception:
        pass
    for p in produk:
        p['terpakai'] = p['id'] in terpakai_ids
    # Kelompokkan HANYA baris halaman ini per nama, urut sesuai urutan grup yg sudah
    # ditentukan (nama_hal) -- bukan sort ulang dari nol (grupnya sudah diurut server).
    _grp = {}
    for p in produk:
        _grp.setdefault(p['nama'], []).append(p)
    _urut = {n: i for i, n in enumerate(nama_hal)}
    grouped = sorted(_grp.items(), key=lambda g: _urut.get(g[0], len(_urut)))
    if sort == 'kode':
        # Varian dalam satu grup ikut urut kode (natural sort) supaya konsisten
        # dengan urutan grup di mode ini.
        for _nama, _members in grouped:
            _members.sort(key=lambda m: _kode_natural_key(m['kode']))
    # Peringatan stok minimum dihitung dari SELURUH data (bukan cuma halaman ini),
    # supaya angkanya tidak ikut mengecil saat dipaginasi -- dulu ini dihitung di
    # template dari `produk` penuh, jadi kalau tak dipindah ke sini angkanya akan
    # diam-diam salah begitu paginasi aktif (cuma menghitung produk di halaman ini).
    low_list = [dict(r) for r in conn.execute(
        f"SELECT * FROM produk{where_sql} AND stok>=0 AND COALESCE(min_stok,0)>0 "
        "AND stok<=min_stok ORDER BY nama, varian", params).fetchall()]
    kategori_produk = [
        row['kategori'] for row in conn.execute(
            "SELECT DISTINCT TRIM(kategori) AS kategori FROM produk "
            "WHERE COALESCE(aktif,1)=? AND TRIM(COALESCE(kategori,''))<>'' "
            "ORDER BY kategori COLLATE NOCASE",
            (0 if show_arsip else 1,)
        ).fetchall()
    ]
    nilai_produk = conn.execute(
        "SELECT COALESCE(SUM(stok * harga_beli),0) FROM produk WHERE COALESCE(aktif,1)=1"
    ).fetchone()[0] or 0
    nilai_non_sku = current_non_sku_value(conn)
    nilai_inventory = float(nilai_produk) + float(nilai_non_sku)
    non_sku_rows = conn.execute(
        """SELECT n.*,j.nomor_tx,j.keterangan
           FROM pergerakan_persediaan_non_sku n
           LEFT JOIN jurnal j ON j.id=n.jurnal_id
           ORDER BY n.tanggal DESC,n.id DESC LIMIT 10"""
    ).fetchall()
    negative_produk = conn.execute(
        "SELECT * FROM produk WHERE COALESCE(aktif,1)=1 AND stok<0 ORDER BY stok,nama"
    ).fetchall()
    saldo_persediaan = persediaan_ledger_value(conn)
    selisih_persediaan = float(nilai_inventory) - saldo_persediaan
    conn.close()
    return render_template('stok.html', produk=produk, grouped=grouped, sort=sort, q=q,
                           pg=pg, low_list=low_list,
                           show_arsip=show_arsip,
                           kategori=kategori, kategori_produk=kategori_produk,
                           nilai_inventory=nilai_inventory, nilai_produk=nilai_produk,
                           nilai_non_sku=nilai_non_sku, non_sku_rows=non_sku_rows,
                           negative_produk=negative_produk,
                           saldo_persediaan=saldo_persediaan,
                           selisih_persediaan=selisih_persediaan)

@app.route('/stok/perbaiki-hpp', methods=['POST'])
@inventory_hpp_required
def stok_perbaiki_hpp():
    """Tutup selisih 'Total Valuasi' vs 'Ledger Persediaan 113*' tanpa menyentuh jurnal.

    Penyebab paling umum: transaksi diinput tidak urut tanggal (mis. migrasi data
    lama), sehingga HPP rata-rata bergerak (produk.harga_beli) ketinggalan dari
    urutan input, bukan urutan tanggal. Ledger 113* sendiri sudah benar karena
    dibangun dari nilai riil yang diposting; yang dikoreksi cuma master data
    harga_beli supaya cocok dengan riwayat pergerakan_stok yang sebenarnya.
    Produk yang qty-nya sendiri tidak cocok dgn riwayat (bukan cuma nilainya)
    SENGAJA dilewati -- itu wilayah Upload Excel mode Koreksi Stok Opname supaya
    ada jejak jurnalnya, bukan diam-diam ditimpa di sini.

    Produk yang pernah direstock persis setelah stoknya pas 0 di TENGAH riwayat
    (bukan pengisian pertama) JUGA dilewati, dilaporkan terpisah bukan ditimpa
    diam-diam. Kalau ada transaksi di sekitar titik nol itu yang sempat dihapus/
    diedit, jumlah polos dari seluruh riwayat bisa tampak "cocok" padahal
    sebenarnya sudah tercampur antara batch lama dan batch baru, dan tidak ada
    cara aman menghitung ulang otomatis mana yang benar tanpa tahu niat aslinya.
    """
    conn = db()
    berisiko = {r['produk_id'] for r in conn.execute("""
        SELECT DISTINCT ps.produk_id FROM pergerakan_stok ps
        WHERE ps.jenis='MASUK' AND ps.stok_sebelum IS NOT NULL AND ABS(ps.stok_sebelum) < 0.000001
          AND EXISTS (SELECT 1 FROM pergerakan_stok ps2
                      WHERE ps2.produk_id = ps.produk_id AND ps2.id < ps.id)
    """)}

    agg = {}
    for r in conn.execute("SELECT produk_id, jenis, qty, harga, hpp_sebelum FROM pergerakan_stok"):
        a = agg.setdefault(r['produk_id'], {'qty': 0.0, 'nilai': 0.0})
        if r['jenis'] == 'MASUK':
            a['qty'] += r['qty']
            a['nilai'] += r['qty'] * (r['harga'] or 0)
        else:
            biaya = r['hpp_sebelum'] if r['hpp_sebelum'] is not None else (r['harga'] or 0)
            a['qty'] -= r['qty']
            a['nilai'] -= r['qty'] * biaya

    diperbaiki = []
    perlu_cek = []
    for p in conn.execute("SELECT id, kode, nama, stok, harga_beli FROM produk WHERE COALESCE(aktif,1)=1"):
        a = agg.get(p['id'])
        stok = p['stok'] or 0
        if not a or abs(stok) < 1e-9:
            continue
        if abs(stok - a['qty']) > 1e-6:
            continue                                    # beda qty -> bukan wilayah fitur ini
        hpp_benar = a['nilai'] / stok
        beda = (stok * (p['harga_beli'] or 0)) - a['nilai']
        if abs(beda) > 0.005:
            if p['id'] in berisiko:
                perlu_cek.append(p)
                continue
            conn.execute('UPDATE produk SET harga_beli=? WHERE id=?', (hpp_benar, p['id']))
            diperbaiki.append((p, hpp_benar, beda))

    if perlu_cek:
        ringkas_cek = '; '.join(f"{p['kode']} ({p['nama']})" for p in perlu_cek[:20])
        add_log(conn, 'Perlu cek HPP (dilewati Perbaiki HPP)',
                f"{len(perlu_cek)} produk pernah restock setelah stok pas 0, dilewati karena "
                f"tidak aman dihitung otomatis: {ringkas_cek}", 'INPUT')

    if not diperbaiki:
        conn.commit()
        conn.close()
        if perlu_cek:
            flash(f'Tidak ada HPP yang diperbaiki otomatis. {len(perlu_cek)} produk (cek Log '
                  f'Aktivitas) dilewati karena pernah restock setelah stok pas 0, ini butuh '
                  f'verifikasi manual lewat Edit Produk, bukan ditimpa otomatis.', 'warning')
        else:
            flash('Tidak ada HPP yang perlu diperbaiki. Selisih (jika masih ada) berasal dari '
                  'produk dengan stok tidak cocok riwayat -- gunakan Upload Excel mode '
                  'Koreksi Stok Terkini Hasil Opname untuk itu.', 'info')
        return redirect(url_for('stok_list'))

    total_tutup = sum(b for _, _, b in diperbaiki)
    ringkas = '; '.join(f"{p['kode']} {p['harga_beli']:,.0f}->{hpp:,.0f}" for p, hpp, _ in diperbaiki[:20])
    add_log(conn, 'Perbaiki HPP persediaan',
            f"{len(diperbaiki)} produk | selisih ditutup Rp {total_tutup:,.0f} | {ringkas}", 'INPUT')
    conn.commit()
    conn.close()
    pesan_sukses = (f'HPP {len(diperbaiki)} produk diperbaiki, menutup selisih Rp {total_tutup:,.0f}. '
                    'Jurnal tidak diubah -- Neraca dan Laba Rugi tetap sama.')
    if perlu_cek:
        pesan_sukses += (f' {len(perlu_cek)} produk lain dilewati (perlu verifikasi manual, '
                          f'cek Log Aktivitas).')
    flash(pesan_sukses, 'success')
    return redirect(url_for('stok_list'))

@app.route('/stok/produk/baru', methods=['POST'])
@inventory_required
def produk_baru():
    conn = db()
    kode_base = request.form.get('kode','').strip()
    nama      = request.form.get('nama','').strip()
    kategori  = request.form.get('kategori','').strip()

    # multi-variant arrays
    varian_list    = request.form.getlist('varian[]')
    satuan_list    = request.form.getlist('satuan[]')
    # Barcode hanya ditawarkan pada form varian TUNGGAL (tabel multi-varian sudah
    # sangat lebar). Saat mode multi-varian, list ini kosong dan tiap varian
    # mengisi barcodenya lewat Edit Produk masing-masing.
    barcode_list   = request.form.getlist('barcode[]')
    stok_list      = request.form.getlist('stok_awal[]')
    harga_beli_list= request.form.getlist('harga_beli[]')
    harga_jual_list= request.form.getlist('harga_jual[]')
    min_stok_list  = request.form.getlist('min_stok[]')

    count = len(varian_list)
    added = 0
    errors = []
    today = date.today()
    stok_awal_rows = []

    for i in range(count):
        varian     = varian_list[i].strip() if i < len(varian_list) else ''
        satuan     = satuan_list[i] if i < len(satuan_list) else 'pcs'
        try:
            stok_awal = parse_qty(stok_list[i] if i < len(stok_list) else '0')
        except (ValueError, TypeError):
            errors.append(f'Stok awal {kode_base or nama} harus berupa angka yang valid')
            continue
        harga_beli = parse_rp(harga_beli_list[i] if i < len(harga_beli_list) else '0')
        if not has_permission('inventory_hpp'):
            stok_awal = 0
            harga_beli = 0
        harga_jual = parse_rp(harga_jual_list[i] if i < len(harga_jual_list) else '0')
        try:
            min_stok = parse_qty(min_stok_list[i] if i < len(min_stok_list) else '0')
        except (ValueError, TypeError):
            errors.append(f'Min stok {kode_base or nama} harus berupa angka yang valid')
            continue
        if stok_awal > 0 and harga_beli <= 0:
            errors.append(f'Harga beli wajib diisi jika stok awal {kode_base or nama} lebih dari 0')
            continue

        # kode: base for single variant, base-1/base-2/... for multiple
        kode = kode_base if count == 1 else f'{kode_base}-{i+1}'

        barcode = parse_barcode(barcode_list[i] if i < len(barcode_list) else '')
        bentrok = barcode_bentrok(conn, barcode)
        if bentrok:
            errors.append(f'Barcode {barcode} sudah dipakai produk '
                          f'{bentrok["kode"]} - {bentrok["nama"]}')
            continue

        try:
            conn.execute(
                'INSERT INTO produk(kode,nama,varian,satuan,harga_beli,harga_jual,stok,min_stok,kategori,barcode) '
                'VALUES(?,?,?,?,?,?,?,?,?,?)',
                (kode, nama, varian, satuan, harga_beli, harga_jual, stok_awal, min_stok, kategori, barcode)
            )
            if stok_awal > 0:
                pid = conn.execute('SELECT id FROM produk WHERE kode=?', (kode,)).fetchone()['id']
                stok_awal_rows.append((pid, stok_awal, harga_beli))
            added += 1
        except sqlite3.IntegrityError:
            errors.append(f'Kode {kode} sudah ada')

    total_stok_awal = sum(qty * harga for _, qty, harga in stok_awal_rows)
    if total_stok_awal > 0:
        jid = insert_jurnal(conn, str(today), 'Saldo awal persediaan', 'PENDANAAN', 'SALDO_AWAL', [
            ('1130', total_stok_awal, 0), ('3100', 0, total_stok_awal)
        ])
        for pid, qty, harga in stok_awal_rows:
            record_stock_movement(conn, pid, today, 'MASUK', qty, harga,
                                  'Saldo awal persediaan', jid, 'ADD_LAYER', 0, 0)
        add_log(conn, 'Saldo awal persediaan', f'Rp {total_stok_awal:,.0f} | {len(stok_awal_rows)} produk', 'INPUT')
    conn.commit()
    conn.close()

    if added:
        flash(f'{added} varian produk berhasil ditambahkan!', 'success')
    for e in errors:
        flash(e, 'danger')
    return redirect(url_for('stok_list'))


# ---------- TEMPLATE & IMPORT EXCEL INVENTORY ----------
# 'Barcode' sengaja ditaruh di URUTAN TERAKHIR dan tanpa tanda wajib. File Excel
# pelanggan lama yang belum punya kolom ini tetap terimpor normal, karena
# stok_import() mencocokkan kolom berdasarkan NAMA header, bukan posisinya.
INV_TEMPLATE_HEADERS = ['Kode*', 'Nama Produk*', 'Kategori', 'Varian', 'Satuan',
                        'Harga Beli', 'Harga Jual', 'Stok Saat Ini', 'Min Stok',
                        'Barcode']

@app.route('/stok/template')
@inventory_required
def stok_template():
    if not has_permission('inventory_hpp'):
        flash('Template import stok berisi kolom HPP/valuasi dan hanya bisa diunduh oleh user dengan hak Lihat/Edit HPP Inventory.', 'danger')
        return redirect(url_for('stok_list'))
    if not HAS_XLSX:
        flash('Library openpyxl tidak tersedia.', 'danger')
        return redirect(url_for('stok_list'))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Produk'
    # header
    for c, h in enumerate(INV_TEMPLATE_HEADERS, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True, color='FFFFFF', size=10)
        cell.fill = PatternFill(patternType='solid', fgColor='1E293B')
        cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 20
    # contoh baris
    contoh = [
        ['BRG-001', 'Kaos Polos', 'Pakaian', 'Hitam - M', 'pcs', 25000, 50000, 100, 10],
        ['BRG-002', 'Celana Jeans', 'Pakaian', '', 'pcs', 80000, 150000, 30, 5],
        ['BHN-001', 'Tepung Terigu', 'Bahan Baku', '', 'kg', 12000, 0, 50, 10],
    ]
    for r, row in enumerate(contoh, start=2):
        for c, val in enumerate(row, start=1):
            cc = ws.cell(row=r, column=c, value=val)
            if c >= 6:
                cc.number_format = '#,##0'
            cc.font = Font(color='94A3B8', italic=True, size=10)
    widths = [14, 22, 16, 16, 10, 14, 14, 12, 10]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w

    # sheet petunjuk
    ws2 = wb.create_sheet('Petunjuk')
    petunjuk = [
        'CARA PAKAI TEMPLATE IMPORT INVENTORY',
        '',
        '1. Isi data produk mulai BARIS 2 di sheet "Produk".',
        '2. Hapus 3 baris contoh (BRG-001 dst) sebelum upload, ganti dengan data Anda.',
        '3. Kolom dengan tanda * WAJIB diisi: Kode dan Nama Produk.',
        '4. Kode adalah kunci produk. Produk lama dapat diperbarui melalui mode koreksi stok.',
        '5. Angka (harga/stok) cukup ketik angka biasa, tanpa titik/Rp. Contoh: 25000',
        '6. Satuan contoh: pcs, kg, liter, meter, dus, lusin.',
        '7. Kategori dan satuan boleh diisi bebas sesuai kebutuhan usaha.',
        '8. Simpan file, lalu upload kembali di menu Inventory > Upload Excel.',
        '',
        'PENTING soal HARGA BELI:',
        '- Jika mengisi Stok Saat Ini untuk mode SALDO AWAL / OPNAME, kolom',
        '  HARGA BELI WAJIB diisi. Tanpa harga beli, persediaan TIDAK BISA dinilai',
        '  (valuasi = 0) dan baris akan dilewati saat upload.',
        '- Untuk stok lama yang harga belinya tidak diingat: isi PERKIRAAN harga',
        '  kulakan saat ini. Yang penting tidak nol.',
        '',
        'Pilih mode upload di aplikasi:',
        '- Tambah Master Produk: buat/update katalog SAJA. Stok & persediaan TIDAK',
        '  dicatat (qty diabaikan walau diisi).',
        '- Input Saldo Awal: untuk barang yang SUDAH DIMILIKI sebelum pakai app ini.',
        '  TIDAK dicatat sebagai pengeluaran (sudah dibeli di masa lalu), TAPI menambah',
        '  valuasi Persediaan. Jurnal: Persediaan (debit) vs Modal Pemilik (kredit).',
        '- Koreksi Hasil Opname: set stok fisik terkini, selisih dijurnal ke',
        '  Koreksi Persediaan.',
    ]
    for r, line in enumerate(petunjuk, start=1):
        cell = ws2.cell(row=r, column=1, value=line)
        if r == 1:
            cell.font = Font(bold=True, size=12, color='1E293B')
    ws2.column_dimensions['A'].width = 80

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return send_file(buf, download_name='template_inventory.xlsx', as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/stok/export')
@inventory_required
def stok_export():
    """Export stok produk aktif (mengikuti filter yang sedang dibuka) ke Excel sebagai
    lembar kerja opname fisik. Kolomnya SAMA dengan INV_TEMPLATE_HEADERS supaya bisa
    langsung diupload ulang lewat /stok/import mode 'Koreksi Hasil Opname' tanpa perlu
    endpoint import baru — kolom "Stok Sistem" tambahan hanya referensi (diabaikan saat
    import karena header_map di stok_import() mencari kolom berdasarkan NAMA, bukan posisi)."""
    if not has_permission('inventory_hpp'):
        flash('Export stok opname berisi kolom HPP/valuasi dan hanya bisa diunduh oleh user dengan hak Lihat/Edit HPP Inventory.', 'danger')
        return redirect(url_for('stok_list'))
    if not HAS_XLSX:
        flash('Library openpyxl tidak tersedia.', 'danger')
        return redirect(url_for('stok_list'))

    conn = db()
    q = request.args.get('q', '').strip()
    kategori = request.args.get('kategori', '').strip()
    show_arsip = request.args.get('arsip') == '1'
    where = ["COALESCE(aktif,1)=?"]
    params = [0 if show_arsip else 1]
    if q:
        # Barcode ikut dicari supaya hasil scan langsung ketemu produknya,
        # sama seperti mengetik kode atau nama.
        where.append("(nama LIKE ? OR kode LIKE ? OR kategori LIKE ? OR COALESCE(barcode,'') LIKE ?)")
        params.extend((f'%{q}%', f'%{q}%', f'%{q}%', f'%{q}%'))
    if kategori == '__tanpa__':
        where.append("TRIM(COALESCE(kategori,''))=''")
    elif kategori:
        where.append("kategori=?")
        params.append(kategori)
    rows = conn.execute(
        'SELECT * FROM produk WHERE ' + ' AND '.join(where) + ' ORDER BY nama, varian',
        params
    ).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Produk'
    headers = ['Kode*', 'Nama Produk*', 'Kategori', 'Varian', 'Satuan',
               'Harga Beli', 'Harga Jual', 'Stok Sistem (Referensi)',
               'Stok Saat Ini', 'Min Stok', 'Barcode']
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True, color='FFFFFF', size=10)
        cell.fill = PatternFill(patternType='solid', fgColor='1E293B')
        cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = 'A2'

    for r, p in enumerate(rows, start=2):
        stok = float(p['stok'] or 0)
        vals = [p['kode'], p['nama'], p['kategori'] or '', p['varian'] or '',
                p['satuan'] or 'pcs', float(p['harga_beli'] or 0), float(p['harga_jual'] or 0),
                stok, stok, float(p['min_stok'] or 0),
                (p['barcode'] if 'barcode' in p.keys() else '') or '']
        for c, val in enumerate(vals, start=1):
            cc = ws.cell(row=r, column=c, value=val)
            if c in (6, 7, 8, 9, 10):
                cc.number_format = '#,##0'
            elif c == 11:
                # Format TEKS. Tanpa ini Excel memperlakukan barcode sebagai bilangan
                # saat file disimpan ulang: nol depan hilang dan 13 digit berubah
                # jadi notasi ilmiah, lalu rusak saat diimpor balik.
                cc.number_format = '@'
        # Kolom "Stok Sistem" abu-abu miring: penanda visual "acuan, jangan diisi ulang"
        ws.cell(row=r, column=8).font = Font(color='94A3B8', italic=True, size=10)
    widths = [14, 24, 16, 16, 10, 13, 13, 16, 13, 10, 18]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w

    ws2 = wb.create_sheet('Petunjuk Opname')
    petunjuk = [
        'CARA PAKAI EXPORT STOK OPNAME',
        '',
        f'Diekspor: {date.today().strftime("%d-%m-%Y")}  ·  {len(rows)} produk',
        '',
        '1. Lakukan hitung fisik barang di gudang/toko.',
        '2. Isi kolom "Stok Saat Ini" dengan HASIL HITUNG FISIK (bukan kolom Stok Sistem).',
        '3. Kolom "Stok Sistem (Referensi)" adalah acuan pembanding saja, TIDAK dibaca',
        '   saat upload — boleh dibiarkan apa adanya.',
        '4. Produk yang stoknya SAMA dengan sistem, biarkan kolom "Stok Saat Ini" apa adanya',
        '   (tidak perlu diubah, tidak akan menimbulkan koreksi).',
        '5. Simpan file ini, lalu buka menu Inventory > Upload Excel,',
        '   pilih mode "Koreksi Hasil Opname", isi alasan (mis. "Opname bulanan Juli 2026").',
        '6. Sistem otomatis menghitung selisih dan mencatat jurnal ke Koreksi Persediaan.',
        '   Selisih lebih banyak dari sistem menambah nilai persediaan; lebih sedikit mengurangi.',
    ]
    for r, line in enumerate(petunjuk, start=1):
        cell = ws2.cell(row=r, column=1, value=line)
        if r == 1:
            cell.font = Font(bold=True, size=12, color='1E293B')
    ws2.column_dimensions['A'].width = 85

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    fname = f"stok_opname_{date.today().strftime('%Y%m%d')}.xlsx"
    return send_file(buf, download_name=fname, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/stok/import', methods=['POST'])
@inventory_required
def stok_import():
    if not HAS_XLSX:
        flash('Library openpyxl tidak tersedia untuk membaca Excel.', 'danger')
        return redirect(url_for('stok_list'))
    f = request.files.get('file')
    if not f or not f.filename:
        flash('Pilih file Excel (.xlsx) dulu.', 'warning')
        return redirect(url_for('stok_list'))
    if not f.filename.lower().endswith('.xlsx'):
        flash('Format harus .xlsx (hasil dari template). File lain tidak didukung.', 'danger')
        return redirect(url_for('stok_list'))

    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        ws = wb['Produk'] if 'Produk' in wb.sheetnames else wb.active
    except Exception as ex:
        flash(f'Gagal membaca file Excel: {ex}', 'danger')
        return redirect(url_for('stok_list'))

    def _num(v):
        """Balikan angka, atau None kalau isinya benar benar bukan angka.
        parse_rp() sendiri boleh melempar ValueError untuk teks acak (dipakai
        di banyak tempat lain dengan asumsi pemanggilnya menangkap sendiri),
        jadi di sini percobaan keduanya wajib ikut dibungkus. Tanpa ini, satu
        sel salah ketik (misal "1 kali" di kolom Harga Beli) bisa menjatuhkan
        seluruh proses upload jadi 500, bukan cuma melewati baris itu."""
        if v is None or v == '':
            return 0.0
        try:
            return float(v)
        except (ValueError, TypeError):
            pass
        try:
            return parse_rp(str(v))
        except (ValueError, TypeError):
            return None

    mode = request.form.get('mode', 'MASTER')
    if mode not in ('MASTER', 'SALDO_AWAL', 'KOREKSI'):
        flash('Mode import inventory tidak valid.', 'danger')
        return redirect(url_for('stok_list'))
    can_inventory_hpp = has_permission('inventory_hpp')
    if mode != 'MASTER' and not can_inventory_hpp:
        flash('Upload stok awal/opname membutuhkan hak Lihat/Edit HPP Inventory karena nilai persediaan harus dihitung.', 'danger')
        return redirect(url_for('stok_list'))
    alasan = (request.form.get('alasan', '') or '').strip()
    if mode == 'KOREKSI' and not alasan:
        flash('Alasan koreksi hasil opname wajib diisi.', 'warning')
        return redirect(url_for('stok_list'))

    def _header(v):
        return ' '.join(str(v or '').replace('*', '').strip().lower().split())

    header_map = {
        _header(cell.value): idx for idx, cell in enumerate(ws[1]) if _header(cell.value)
    }
    if 'kode' not in header_map or 'nama produk' not in header_map:
        flash('Header Excel wajib memiliki kolom Kode dan Nama Produk.', 'danger')
        return redirect(url_for('stok_list'))

    def _cell(row, name, default=None):
        idx = header_map.get(name)
        return row[idx] if idx is not None and idx < len(row) else default

    conn = db()
    today = date.today()
    added = 0
    updated = 0
    errors = []
    adjusted = 0
    total_valuasi = 0.0          # akumulasi nilai persediaan yang tercatat
    skipped_no_price = []        # produk dilewati karena Harga Beli kosong
    master_stok_ignored = 0      # produk berstok tapi dipilih mode MASTER (stok diabaikan)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        kode = str(_cell(row, 'kode', '') or '').strip()
        nama = str(_cell(row, 'nama produk', '') or '').strip()
        if not kode or not nama:
            continue  # baris kosong / tidak lengkap → lewati diam-diam
        kategori = str(_cell(row, 'kategori', '') or '').strip()
        varian = str(_cell(row, 'varian', '') or '').strip()
        satuan = str(_cell(row, 'satuan', '') or '').strip() or 'pcs'
        harga_beli = _num(_cell(row, 'harga beli', 0)) if can_inventory_hpp else 0
        harga_jual = _num(_cell(row, 'harga jual', 0))
        stok_target = _num(_cell(row, 'stok saat ini', 0))
        min_stok = _num(_cell(row, 'min stok', 0))
        # Kolom yang isinya bukan angka sama sekali (misal "1 kali" atau "1 Tim"
        # ke kolom Harga Beli) ditandai None oleh _num(). Baris ini dilewati
        # dengan pesan yang menunjuk kolom dan nilainya, bukan dibiarkan
        # menjatuhkan seluruh upload.
        if None in (harga_beli, harga_jual, stok_target, min_stok):
            for nama_field, label, nilai in (
                ('harga beli', 'Harga Beli', harga_beli),
                ('harga jual', 'Harga Jual', harga_jual),
                ('stok saat ini', 'Stok Saat Ini', stok_target),
                ('min stok', 'Min Stok', min_stok),
            ):
                if nilai is None:
                    mentah = _cell(row, nama_field)
                    errors.append(f'{kode}: kolom {label} berisi "{mentah}" yang bukan angka. Baris ini dilewati.')
            continue
        if min(harga_beli, harga_jual, stok_target, min_stok) < 0:
            errors.append(f'{kode}: angka harga dan stok tidak boleh negatif')
            continue
        # Mode MASTER hanya bikin katalog, stok diabaikan. Ingatkan bila kolom stok terisi.
        if mode == 'MASTER' and stok_target > 0:
            master_stok_ignored += 1
        # Stok awal/opname WAJIB punya harga beli, kalau tidak persediaan tak bisa dinilai.
        if mode != 'MASTER' and stok_target > 0 and harga_beli <= 0:
            skipped_no_price.append(kode)
            continue

        # Barcode HANYA disentuh kalau filenya memang punya kolomnya. File lama
        # tanpa kolom Barcode tidak boleh menghapus barcode yang sudah tersimpan.
        punya_kolom_barcode = 'barcode' in header_map
        barcode_baris = parse_barcode(_cell(row, 'barcode', '')) if punya_kolom_barcode else ''
        if barcode_baris:
            _bentrok = conn.execute(
                'SELECT kode FROM produk WHERE barcode=? AND kode<>?', (barcode_baris, kode)
            ).fetchone()
            if _bentrok:
                # Barisnya tetap diimpor (harga/stoknya sah), yang dibatalkan cuma
                # barcodenya, dan itu dilaporkan supaya tidak berubah diam-diam.
                errors.append(f'{kode}: barcode {barcode_baris} sudah dipakai {_bentrok["kode"]}. '
                              f'Data lain baris ini tetap diimpor, barcodenya saja yang dilewati.')
                barcode_baris = ''

        existing = conn.execute('SELECT * FROM produk WHERE kode=?', (kode,)).fetchone()
        try:
            if existing:
                pid = existing['id']
                hpp_master = existing['harga_beli']
                if can_inventory_hpp:
                    hpp_master = harga_beli if float(existing['stok'] or 0) <= 0 else existing['harga_beli']
                if 'kategori' in header_map:
                    conn.execute(
                        'UPDATE produk SET nama=?,kategori=?,varian=?,satuan=?,harga_beli=?,harga_jual=?,min_stok=? WHERE id=?',
                        (nama, kategori, varian, satuan, hpp_master, harga_jual, min_stok, pid)
                    )
                else:
                    conn.execute(
                        'UPDATE produk SET nama=?,varian=?,satuan=?,harga_beli=?,harga_jual=?,min_stok=? WHERE id=?',
                        (nama, varian, satuan, hpp_master, harga_jual, min_stok, pid)
                    )
                # Barcode di-update terpisah, bukan digabung ke UPDATE di atas, supaya
                # tidak melahirkan kombinasi SQL (ada/tidak kategori x ada/tidak barcode).
                if punya_kolom_barcode:
                    conn.execute('UPDATE produk SET barcode=? WHERE id=?', (barcode_baris, pid))
                updated += 1
            else:
                conn.execute(
                    'INSERT INTO produk(kode,nama,varian,satuan,harga_beli,harga_jual,stok,min_stok,kategori,barcode) '
                    'VALUES(?,?,?,?,?,?,0,?,?,?)',
                    (kode, nama, varian, satuan, harga_beli, harga_jual, min_stok, kategori, barcode_baris)
                )
                pid = conn.execute('SELECT last_insert_rowid()').fetchone()[0]
                added += 1
            if mode != 'MASTER':
                jid, delta_qty, delta_nilai = adjust_inventory_position(
                    conn, pid, stok_target, harga_beli, str(today),
                    alasan or ('Import saldo awal inventory' if mode == 'SALDO_AWAL' else 'Import hasil opname'),
                    'SALDO_AWAL' if mode == 'SALDO_AWAL' else 'KOREKSI'
                )
                if abs(delta_qty) >= 0.000001 or abs(delta_nilai) >= 0.01:
                    adjusted += 1
                    total_valuasi += delta_nilai
        except sqlite3.IntegrityError:
            errors.append(f'{kode}: gagal disimpan')

    if added or updated:
        add_log(conn, 'Import inventory Excel',
                f'Mode {mode} | {added} baru | {updated} diperbarui | {adjusted} stok disesuaikan', 'INPUT')
    conn.commit(); conn.close()

    # ── Ringkasan hasil ──────────────────────────────────────────────
    def _fmt_rp(n):
        return 'Rp ' + f'{abs(n):,.0f}'.replace(',', '.')

    msg = f'{added} produk baru, {updated} produk diperbarui'
    if mode != 'MASTER' and adjusted:
        arah = 'bertambah' if total_valuasi >= 0 else 'berkurang'
        lawan = 'Modal Pemilik' if mode == 'SALDO_AWAL' else 'Koreksi Persediaan'
        msg += f'. Valuasi persediaan {arah} {_fmt_rp(total_valuasi)} (jurnal lawan: {lawan})'
    flash(msg, 'success' if added or updated else 'warning')

    # Peringatan: mode MASTER mengabaikan stok yang sudah diisi user
    if master_stok_ignored:
        flash(f'⚠️ {master_stok_ignored} produk punya kolom "Stok Saat Ini" terisi, '
              f'tetapi mode "Tambah / Update Master Produk" TIDAK mencatat stok maupun valuasi '
              f'persediaan. Untuk mencatat stok awal, pilih mode "Input Saldo Awal / Stok Saat Ini" '
              f'lalu upload ulang file yang sama.', 'warning')

    # Peringatan: produk dilewati karena Harga Beli kosong (persediaan tak bisa dinilai)
    if skipped_no_price:
        contoh = ', '.join(skipped_no_price[:5]) + ('…' if len(skipped_no_price) > 5 else '')
        flash(f'⚠️ {len(skipped_no_price)} produk dilewati karena Harga Beli kosong: {contoh}. '
              f'Persediaan tidak bisa dinilai tanpa harga beli. Isi kolom Harga Beli '
              f'(boleh perkiraan harga kulakan saat ini), lalu upload ulang.', 'danger')

    for e in errors:
        flash(e, 'danger')
    return redirect(url_for('stok_list'))


@app.route('/stok/<int:id>/edit', methods=['GET','POST'])
@inventory_required
def produk_edit(id):
    conn = db()
    p = conn.execute(
        'SELECT * FROM produk WHERE id=? AND COALESCE(aktif,1)=1',
        (id,)
    ).fetchone()
    if not p:
        conn.close(); flash('Produk tidak ditemukan atau sudah diarsipkan.','danger')
        return redirect(url_for('stok_list'))
    can_inventory_hpp = has_permission('inventory_hpp')
    if request.method == 'POST':
        try:
            harga_beli = parse_nonnegative_rp(request.form.get('harga_beli',0), 'Harga beli') if can_inventory_hpp else float(p['harga_beli'] or 0)
            harga_jual = parse_nonnegative_rp(request.form.get('harga_jual',0), 'Harga jual')
            min_stok = parse_qty(request.form.get('min_stok',0))
        except (TypeError, ValueError):
            conn.close(); flash('Harga dan minimum stok harus berupa angka valid.', 'danger')
            return redirect(url_for('produk_edit', id=id))
        if not math.isfinite(min_stok) or min_stok < 0:
            conn.close(); flash('Minimum stok tidak boleh negatif.', 'danger')
            return redirect(url_for('produk_edit', id=id))
        stok_kini = float(p['stok'] or 0)
        if can_inventory_hpp and stok_kini > 0 and harga_beli <= 0:
            conn.close()
            flash('HPP wajib lebih dari 0 jika produk masih memiliki stok.', 'warning')
            return redirect(url_for('produk_edit', id=id))
        pos_label = (request.form.get('pos_label') or '').strip()[:5].upper()
        barcode = parse_barcode(request.form.get('barcode'))
        bentrok = barcode_bentrok(conn, barcode, exclude_id=id)
        if bentrok:
            conn.close()
            flash(f'Barcode {barcode} sudah dipakai produk {bentrok["kode"]} - {bentrok["nama"]}. '
                  f'Satu barcode hanya boleh untuk satu produk, kalau tidak hasil scan jadi ambigu.',
                  'danger')
            return redirect(url_for('produk_edit', id=id))
        conn.execute("""UPDATE produk SET nama=?,varian=?,satuan=?,kategori=?,
                        harga_jual=?,min_stok=?,pos_label=?,barcode=? WHERE id=?""",
            (request.form['nama'], request.form.get('varian',''),
             request.form.get('satuan','pcs'),
             request.form.get('kategori','').strip(),
             harga_jual, min_stok, pos_label, barcode, id))
        if can_inventory_hpp and abs(harga_beli - float(p['harga_beli'] or 0)) >= 0.01:
            if abs(stok_kini) >= 0.000001:
                # Ada stok (positif/negatif): jurnal selisih valuasi + set HPP baru.
                adjust_inventory_position(conn, id, stok_kini, harga_beli, str(date.today()),
                                          'Perubahan HPP dari edit produk', 'KOREKSI')
            else:
                # Stok 0: valuasi tak berubah, tapi HPP tetap harus tersimpan untuk transaksi berikutnya.
                conn.execute("UPDATE produk SET harga_beli=? WHERE id=?", (harga_beli, id))
        conn.commit(); conn.close()
        flash('Produk diperbarui!','success')
        return redirect(url_for('stok_list'))
    kategori_produk = [
        row['kategori'] for row in conn.execute(
            "SELECT DISTINCT TRIM(kategori) AS kategori FROM produk "
            "WHERE TRIM(COALESCE(kategori,''))<>'' ORDER BY kategori COLLATE NOCASE"
        ).fetchall()
    ]
    conn.close()
    return render_template('produk_edit.html', p=p, kategori_produk=kategori_produk)

@app.route('/stok/<int:id>/gerakan', methods=['POST'])
@inventory_required
def stok_gerakan(id):
    conn = db()
    jenis = request.form.get('jenis', 'OPNAME')   # OPNAME | SALDO_AWAL
    try:
        qty = parse_qty(request.form.get('qty', 0))
    except (TypeError, ValueError):
        conn.close(); flash('Jumlah stok harus berupa angka valid.', 'danger')
        return redirect(url_for('stok_list'))
    if not math.isfinite(qty) or qty < 0:
        conn.close(); flash('Jumlah stok tidak boleh negatif.', 'danger')
        return redirect(url_for('stok_list'))
    ket   = (request.form.get('keterangan', '') or '').strip()
    tgl   = request.form.get('tanggal', str(date.today()))
    p = conn.execute(
        'SELECT nama,stok,harga_beli FROM produk WHERE id=? AND COALESCE(aktif,1)=1',
        (id,)
    ).fetchone()
    if not p:
        conn.close()
        flash('Produk tidak ditemukan atau sudah diarsipkan.', 'danger')
        return redirect(url_for('stok_list'))
    if not _operator_date_ok(tgl):
        conn.close()
        flash('Operator hanya bisa mencatat stok untuk bulan berjalan.', 'warning')
        return redirect(url_for('stok_list'))

    is_saldo_awal = (jenis == 'SALDO_AWAL')
    can_inventory_hpp = has_permission('inventory_hpp')

    if is_saldo_awal:
        if not can_inventory_hpp:
            conn.close()
            flash('Stok awal membutuhkan hak Lihat/Edit HPP Inventory karena nilai persediaan harus dihitung.', 'danger')
            return redirect(url_for('stok_list'))
        # Stok Awal, pakai harga dari form, keterangan opsional
        try:
            harga_input = parse_nonnegative_rp(request.form.get('harga_satuan', '0'), 'Harga beli')
        except ValueError as ex:
            conn.close(); flash(str(ex), 'danger')
            return redirect(url_for('stok_list'))
        if qty <= 0:
            conn.close()
            flash('Jumlah stok harus lebih dari 0.', 'warning')
            return redirect(url_for('stok_list'))
        if harga_input <= 0:
            conn.close()
            flash('Harga beli wajib diisi untuk menghitung valuasi persediaan.', 'warning')
            return redirect(url_for('stok_list'))
        ket = ket or 'Saldo awal stok'
        # Gunakan adjust_inventory_position agar jurnal Dr 1130 / Cr 3100 terbentuk
        jid, delta_qty, delta_nilai = adjust_inventory_position(
            conn, id, qty, harga_input, tgl, ket, 'SALDO_AWAL'
        )
        if abs(delta_qty) < 0.000001 and abs(delta_nilai) < 0.01:
            conn.close()
            flash('Tidak ada perubahan stok atau nilai yang perlu dicatat.', 'info')
            return redirect(url_for('stok_list'))
        add_log(conn, 'Stok awal', f"{p['nama']} | qty={qty} | hpp={harga_input:,.0f}", 'INPUT')
        conn.commit(); conn.close()
        flash(f'Stok awal {p["nama"]} berhasil dicatat. Persediaan +Rp {delta_nilai:,.0f}.', 'success')
        return redirect(url_for('stok_list'))

    # ── OPNAME (koreksi), perilaku lama ──────────────────────────────────────
    if not ket:
        conn.close()
        flash('Alasan koreksi stok wajib diisi.', 'warning')
        return redirect(url_for('stok_list'))
    stok_lama = float(p['stok'] or 0)
    delta = qty - stok_lama
    if float(p['harga_beli'] or 0) <= 0 and delta != 0:
        conn.close()
        if can_inventory_hpp:
            flash('Harga beli produk masih 0. Isi HPP produk terlebih dahulu sebelum koreksi stok, '
                  'agar nilai persediaan dapat dihitung dengan benar.', 'warning')
        else:
            flash('Produk ini belum punya HPP. Minta user dengan hak Lihat/Edit HPP Inventory '
                  'mengisi HPP terlebih dahulu sebelum koreksi stok.', 'warning')
        return redirect(url_for('stok_list'))
    if stok_lama + delta < 0:
        conn.close()
        flash(f"Stok tidak mencukupi. Stok tersedia: {stok_lama:.0f}", 'danger')
        return redirect(url_for('stok_list'))
    if abs(delta) < 0.000001:
        conn.close()
        flash('Tidak ada perubahan stok yang perlu dicatat.', 'info')
        return redirect(url_for('stok_list'))
    harga = float(p['harga_beli'] or 0)
    nilai = abs(delta) * harga
    arah = 'MASUK' if delta > 0 else 'KELUAR'
    entries = [('1130', nilai, 0), ('5190', 0, nilai)] if delta > 0 else [
        ('5190', nilai, 0), ('1130', 0, nilai)
    ]
    jid = insert_jurnal(conn, tgl, f"Koreksi stok: {p['nama']} - {ket}", 'OPERASIONAL', 'KOREKSI_STOK', entries)
    conn.execute('UPDATE produk SET stok=stok+? WHERE id=?', (delta, id))
    hpp_effect = 'ADD_LAYER' if arah == 'MASUK' else 'REMOVE_LAYER'
    record_stock_movement(conn, id, tgl, arah, abs(delta), harga, f'Koreksi stok: {ket}',
                          jid, hpp_effect, stok_lama, harga)
    add_log(conn, 'Koreksi stok', f"{p['nama']} | {delta:+g} | {ket}", 'INPUT')
    conn.commit(); conn.close()
    flash('Pergerakan stok dicatat!', 'success')
    return redirect(url_for('stok_list'))

@app.route('/stok/non-sku/gerakan', methods=['POST'])
@inventory_required
@inventory_hpp_required
def stok_non_sku_gerakan():
    conn = db()
    mode = request.form.get('mode', 'KOREKSI_MINUS')
    tanggal = request.form.get('tanggal', str(date.today()))
    deskripsi = (request.form.get('deskripsi', '') or '').strip()
    alasan = (request.form.get('alasan', '') or '').strip()
    try:
        nilai = parse_nonnegative_rp(request.form.get('nilai', '0'), 'Nilai persediaan non-SKU')
    except ValueError as ex:
        return _reject_form(conn, str(ex), 'stok_list')
    if mode not in ('SALDO_AWAL', 'KOREKSI_PLUS', 'KOREKSI_MINUS'):
        conn.close()
        flash('Mode persediaan non-SKU tidak valid.', 'danger')
        return redirect(url_for('stok_list'))
    if not _operator_date_ok(tanggal):
        conn.close()
        flash('Tanggal tidak valid. Operator hanya bisa mencatat stok untuk bulan berjalan.', 'warning')
        return redirect(url_for('stok_list'))
    if nilai <= 0:
        conn.close()
        flash('Nilai persediaan non-SKU harus lebih dari 0.', 'warning')
        return redirect(url_for('stok_list'))
    if not deskripsi:
        conn.close()
        flash('Deskripsi persediaan non-SKU wajib diisi.', 'warning')
        return redirect(url_for('stok_list'))
    if mode != 'SALDO_AWAL' and not alasan:
        conn.close()
        flash('Alasan koreksi non-SKU wajib diisi.', 'warning')
        return redirect(url_for('stok_list'))

    saldo_non_sku = current_non_sku_value(conn)
    if mode == 'KOREKSI_MINUS' and nilai > float(saldo_non_sku) + 0.01:
        conn.close()
        flash(f'Nilai koreksi melebihi persediaan non-SKU tersedia Rp {saldo_non_sku:,.0f}.', 'danger')
        return redirect(url_for('stok_list'))

    if mode == 'SALDO_AWAL':
        ket = f"Saldo awal persediaan non-SKU: {deskripsi}"
        entries = [('1130', nilai, 0), ('3100', 0, nilai)]
        kategori, tipe_tx, delta = 'PENDANAAN', 'SALDO_AWAL', nilai
        log_title = 'Saldo awal non-SKU'
    elif mode == 'KOREKSI_PLUS':
        ket = f"Koreksi persediaan non-SKU bertambah: {deskripsi} - {alasan}"
        entries = [('1130', nilai, 0), ('5190', 0, nilai)]
        kategori, tipe_tx, delta = 'OPERASIONAL', 'KOREKSI_STOK', nilai
        log_title = 'Koreksi non-SKU bertambah'
    else:
        ket = f"Koreksi persediaan non-SKU berkurang: {deskripsi} - {alasan}"
        entries = [('5190', nilai, 0), ('1130', 0, nilai)]
        kategori, tipe_tx, delta = 'OPERASIONAL', 'KOREKSI_STOK', -nilai
        log_title = 'Koreksi non-SKU berkurang'

    jid = insert_jurnal(conn, tanggal, ket, kategori, tipe_tx, entries)
    record_non_sku_movement(conn, jid, tanggal, deskripsi, delta)
    add_log(conn, log_title, f"{deskripsi} | Rp {nilai:,.0f} | {tanggal}", 'INPUT')
    conn.commit()
    conn.close()
    flash('Persediaan non-SKU berhasil dicatat dan jurnal otomatis dibuat.', 'success')
    return redirect(url_for('stok_list'))

@app.route('/stok/<int:id>/riwayat')
@inventory_required
def stok_riwayat(id):
    conn = db()
    produk = conn.execute('SELECT * FROM produk WHERE id=?',(id,)).fetchone()
    riwayat= conn.execute("""
        SELECT ps.*,j.nomor_tx
        FROM pergerakan_stok ps LEFT JOIN jurnal j ON j.id=ps.jurnal_id
        WHERE ps.produk_id=? ORDER BY ps.tanggal DESC, ps.id DESC
    """,(id,)).fetchall()
    conn.close()
    return render_template('stok_riwayat.html', produk=produk, riwayat=riwayat)

def _produk_terpakai(conn, pid):
    """True bila produk masih dipakai di data lain sehingga TIDAK aman di-hard-delete
    (menghindari 'FOREIGN KEY constraint failed' + referensi yatim). Cek semua tabel
    ber-FK ke produk + jejak transaksi + resep produksi. Kalau terpakai -> arsipkan saja."""
    for tbl, col in (('pergerakan_stok', 'produk_id'), ('transaksi_item', 'produk_id'),
                     ('po_item', 'produk_id'), ('produksi', 'produk_jadi_id'),
                     ('produksi_bahan', 'produk_id'), ('produksi_template', 'produk_jadi_id')):
        try:
            if conn.execute(f"SELECT 1 FROM {tbl} WHERE {col}=? LIMIT 1", (pid,)).fetchone():
                return True
        except Exception:
            continue
    # Resep produksi menyimpan bahan SKU sebagai JSON [{'produk_id':.., 'qty':..}, ...]
    try:
        for row in conn.execute("SELECT bahan_sku FROM produksi_template").fetchall():
            for b in (json.loads(row['bahan_sku'] or '[]') or []):
                if isinstance(b, dict) and int(b.get('produk_id') or 0) == pid:
                    return True
    except Exception:
        pass
    return False

@app.route('/stok/<int:id>/hapus', methods=['POST'])
@finance_required
@inventory_required
def produk_hapus(id):
    conn = db()
    try:
        p = conn.execute(
            "SELECT id,kode,nama,stok,COALESCE(aktif,1) AS aktif FROM produk WHERE id=?",
            (id,)
        ).fetchone()
        if not p:
            flash('Produk tidak ditemukan.', 'danger')
            return redirect(url_for('stok_list'))
        # Produk yang masih terpakai di mana pun -> ARSIPKAN (bukan hard-delete), supaya
        # tidak FK-crash & tidak meninggalkan referensi yatim. Hard-delete hanya untuk
        # produk yang belum pernah dipakai sama sekali.
        if _produk_terpakai(conn, id):
            if abs(float(p['stok'] or 0)) >= 0.000001:
                flash('Produk masih memiliki stok. Koreksi/opname stok menjadi 0 dulu sebelum diarsipkan.', 'warning')
                return redirect(url_for('stok_list'))
            if not p['aktif']:
                flash('Produk sudah berada di arsip.', 'info')
                return redirect(url_for('stok_list', arsip=1))
            conn.execute('UPDATE produk SET aktif=0 WHERE id=?', (id,))
            add_log(conn, 'Arsip produk', f"{p['kode']} - {p['nama']}", 'INPUT')
            conn.commit()
            flash('Produk masih terpakai (transaksi/PO/produksi/resep), jadi diarsipkan dan disembunyikan dari list aktif.', 'success')
            return redirect(url_for('stok_list'))
        # Belum pernah dipakai di mana pun -> aman dihapus permanen.
        conn.execute('DELETE FROM produk WHERE id=?', (id,))
        add_log(conn, 'Hapus produk', f"{p['kode']} - {p['nama']}", 'INPUT')
        conn.commit()
        flash('Produk dihapus.', 'warning')
        return redirect(url_for('stok_list'))
    except Exception:
        conn.rollback()
        app.logger.exception('produk_hapus gagal (id=%s)', id)
        flash('Gagal menghapus produk karena masih terkait data lain. Produk tidak diubah; coba arsipkan saja.', 'danger')
        return redirect(url_for('stok_list'))
    finally:
        conn.close()

@app.route('/stok/<int:id>/restore', methods=['POST'])
@finance_required
@inventory_required
def produk_restore(id):
    conn = db()
    p = conn.execute(
        "SELECT id,kode,nama FROM produk WHERE id=?",
        (id,)
    ).fetchone()
    if not p:
        conn.close()
        flash('Produk tidak ditemukan.', 'danger')
        return redirect(url_for('stok_list', arsip=1))
    conn.execute('UPDATE produk SET aktif=1 WHERE id=?', (id,))
    add_log(conn, 'Pulihkan produk arsip', f"{p['kode']} - {p['nama']}", 'INPUT')
    conn.commit()
    conn.close()
    flash('Produk dipulihkan ke list aktif.', 'success')
    return redirect(url_for('stok_list', arsip=1))


# ---------- PRODUKSI SEDERHANA ----------
@app.route('/produksi')
@inventory_required
@inventory_hpp_required
def produksi_list():
    conn = db()
    produk = conn.execute(
        "SELECT id,kode,nama,varian,satuan,harga_beli,harga_jual,stok FROM produk "
        "WHERE COALESCE(aktif,1)=1 ORDER BY nama,varian"
    ).fetchall()
    riwayat = conn.execute("""
        SELECT pr.*, p.nama AS produk_nama, p.varian AS produk_varian, p.satuan AS produk_satuan,
               j.nomor_tx
        FROM produksi pr
        LEFT JOIN produk p ON p.id = pr.produk_jadi_id
        LEFT JOIN jurnal j ON j.id = pr.jurnal_id
        ORDER BY pr.tanggal DESC, pr.id DESC
        LIMIT 50
    """).fetchall()
    # Untuk tiap riwayat, siapkan ringkas bahan
    bahan_map = {}
    for b in conn.execute("""
        SELECT pb.produksi_id, p.nama, p.varian, pb.qty, p.satuan
        FROM produksi_bahan pb LEFT JOIN produk p ON p.id = pb.produk_id
        ORDER BY pb.id
    """).fetchall():
        bahan_map.setdefault(b['produksi_id'], []).append(b)
    non_sku_map = {}
    for n in conn.execute("""
        SELECT produksi_id, deskripsi, nilai, qty, satuan
        FROM produksi_non_sku
        ORDER BY id
    """).fetchall():
        non_sku_map.setdefault(n['produksi_id'], []).append(n)
    nilai_non_sku = current_non_sku_value(conn)
    # produk untuk dropdown JS (list of dict)
    produk_js = [dict(id=p['id'], kode=p['kode'], nama=p['nama'], varian=p['varian'],
                      satuan=p['satuan'], harga_beli=float(p['harga_beli'] or 0),
                      stok=float(p['stok'] or 0)) for p in produk]
    rekening = conn.execute(
        "SELECT kode,nama FROM akun WHERE is_rekening=1 ORDER BY kode"
    ).fetchall()
    # Template batch produksi (untuk produksi berulang)
    templates_js = []
    for t in conn.execute(
            "SELECT id,nama,produk_jadi_id,qty_hasil,biaya_tambahan,bahan_sku,bahan_non_sku,"
            "COALESCE(pos_aktif,0) AS pos_aktif, COALESCE(harga_jual,0) AS harga_jual, "
            "COALESCE(kategori,'') AS kategori "
            "FROM produksi_template ORDER BY nama").fetchall():
        try:
            bsku = json.loads(t['bahan_sku'] or '[]')
        except (ValueError, TypeError):
            bsku = []
        try:
            bnon = json.loads(t['bahan_non_sku'] or '[]')
        except (ValueError, TypeError):
            bnon = []
        templates_js.append(dict(id=t['id'], nama=t['nama'], produk_jadi_id=t['produk_jadi_id'],
                                 qty_hasil=float(t['qty_hasil'] or 0),
                                 biaya_tambahan=float(t['biaya_tambahan'] or 0),
                                 pos_aktif=int(t['pos_aktif'] or 0),
                                 harga_jual=float(t['harga_jual'] or 0),
                                 kategori=(t['kategori'] or ''),
                                 bahan_sku=bsku, bahan_non_sku=bnon))
    # Saran kategori (datalist form template) — gabungan kategori produk + kategori resep yang ada,
    # supaya konsisten dengan chip kategori di POS.
    kategori_list = [r['k'] for r in conn.execute(
        "SELECT DISTINCT TRIM(kategori) AS k FROM ("
        "  SELECT kategori FROM produk WHERE TRIM(COALESCE(kategori,''))<>'' "
        "  UNION SELECT kategori FROM produksi_template WHERE TRIM(COALESCE(kategori,''))<>'' "
        ") ORDER BY k COLLATE NOCASE").fetchall()]
    conn.close()
    today = date.today()
    return render_template('produksi.html', produk=produk, produk_js=produk_js,
                           riwayat=riwayat, bahan_map=bahan_map,
                           non_sku_map=non_sku_map, nilai_non_sku=nilai_non_sku,
                           rekening=rekening, templates_js=templates_js,
                           kategori_list=kategori_list, today=today)

@app.route('/produksi/template/simpan', methods=['POST'])
@inventory_required
@inventory_hpp_required
def produksi_template_simpan():
    data = request.get_json(silent=True) or {}
    nama = (data.get('nama') or '').strip()[:80]
    if not nama:
        return jsonify(ok=False, error='Nama template wajib diisi.'), 400
    try:
        produk_jadi_id = int(data.get('produk_jadi_id') or 0) or None
    except (ValueError, TypeError):
        produk_jadi_id = None
    try:
        qty_hasil = float(data.get('qty_hasil') or 0)
    except (ValueError, TypeError):
        qty_hasil = 0
    try:
        biaya = float(data.get('biaya_tambahan') or 0)
    except (ValueError, TypeError):
        biaya = 0
    # Harga jual resep di POS. 0 = belum diisi (nanti jatuh balik ke harga produk jadi).
    try:
        harga_jual = float(data.get('harga_jual') or 0)
    except (ValueError, TypeError):
        harga_jual = 0
    if harga_jual < 0 or not math.isfinite(harga_jual):
        harga_jual = 0
    # Kategori resep utk grid POS (chip kategori kasir). Kosong = ikut kategori produk jadi.
    kategori = (data.get('kategori') or '').strip()[:40]
    pos_aktif = 1 if data.get('pos_aktif') else 0   # jual resep ini langsung di POS (backflush bahan)
    bsku = []
    for b in (data.get('bahan_sku') or []):
        try:
            pid = int(b.get('produk_id'))
            q = float(b.get('qty') or 0)
        except (ValueError, TypeError, AttributeError):
            continue
        if pid and q > 0:
            bsku.append({'produk_id': pid, 'qty': q})
    bnon = []
    for n in (data.get('bahan_non_sku') or []):
        try:
            desc = (n.get('deskripsi') or '').strip()[:120]
            nilai = float(n.get('nilai') or 0)
            q = float(n.get('qty') or 0)
            sat = (n.get('satuan') or '').strip()[:20]
        except (ValueError, TypeError, AttributeError):
            continue
        if desc:
            bnon.append({'deskripsi': desc, 'nilai': nilai, 'qty': q, 'satuan': sat})
    if not bsku and not bnon:
        return jsonify(ok=False, error='Template kosong: isi minimal satu bahan dulu.'), 400
    # id opsional: kalau dikirim = mode EDIT template existing (update by id, boleh ganti nama).
    # Tanpa id = perilaku lama (cocokkan by nama: update kalau ada, else insert).
    try:
        tmpl_id = int(data.get('id') or 0) or None
    except (ValueError, TypeError):
        tmpl_id = None
    conn = db()
    if tmpl_id is not None:
        row = conn.execute("SELECT id FROM produksi_template WHERE id=?", (tmpl_id,)).fetchone()
        if not row:
            conn.close()
            return jsonify(ok=False, error='Template tidak ditemukan (mungkin sudah dihapus). Muat ulang halaman.'), 404
        clash = conn.execute("SELECT id FROM produksi_template WHERE nama=? AND id<>?", (nama, tmpl_id)).fetchone()
        if clash:
            conn.close()
            return jsonify(ok=False, error='Nama itu sudah dipakai template lain. Pakai nama berbeda.'), 400
        conn.execute(
            "UPDATE produksi_template SET nama=?,produk_jadi_id=?,qty_hasil=?,biaya_tambahan=?,harga_jual=?,kategori=?,"
            "bahan_sku=?,bahan_non_sku=?,pos_aktif=?,updated_at=date('now') WHERE id=?",
            (nama, produk_jadi_id, qty_hasil, biaya, harga_jual, kategori, json.dumps(bsku), json.dumps(bnon), pos_aktif, tmpl_id))
        tid = tmpl_id
    else:
        existing = conn.execute("SELECT id FROM produksi_template WHERE nama=?", (nama,)).fetchone()
        if existing:
            conn.execute(
                "UPDATE produksi_template SET produk_jadi_id=?,qty_hasil=?,biaya_tambahan=?,harga_jual=?,kategori=?,"
                "bahan_sku=?,bahan_non_sku=?,pos_aktif=?,updated_at=date('now') WHERE id=?",
                (produk_jadi_id, qty_hasil, biaya, harga_jual, kategori, json.dumps(bsku), json.dumps(bnon), pos_aktif, existing['id']))
            tid = existing['id']
        else:
            cur = conn.execute(
                "INSERT INTO produksi_template(nama,produk_jadi_id,qty_hasil,biaya_tambahan,harga_jual,kategori,bahan_sku,bahan_non_sku,pos_aktif) "
                "VALUES(?,?,?,?,?,?,?,?,?)",
                (nama, produk_jadi_id, qty_hasil, biaya, harga_jual, kategori, json.dumps(bsku), json.dumps(bnon), pos_aktif))
            tid = cur.lastrowid
    add_log(conn, 'Simpan template produksi',
            '%s | %d bahan SKU, %d non-SKU' % (nama, len(bsku), len(bnon)), 'INPUT')
    # Peringatan (bukan blokir): resep dijual di POS tapi harganya masih 0 -> bisa kejual Rp 0.
    warn = ''
    if pos_aktif and harga_jual <= 0:
        _hj_pj = 0.0
        if produk_jadi_id:
            _pjr = conn.execute("SELECT harga_jual FROM produk WHERE id=?", (produk_jadi_id,)).fetchone()
            _hj_pj = float((_pjr['harga_jual'] if _pjr else 0) or 0)
        if _hj_pj <= 0:
            warn = ('Resep ini aktif di POS tapi harga jualnya masih 0. '
                    'Isi Harga Jual (POS) supaya tidak terjual Rp 0.')
    conn.commit()
    conn.close()
    return jsonify(ok=True, warning=warn,
                   template=dict(id=tid, nama=nama, produk_jadi_id=produk_jadi_id,
                   qty_hasil=qty_hasil, biaya_tambahan=biaya, pos_aktif=pos_aktif,
                   harga_jual=harga_jual, kategori=kategori, bahan_sku=bsku, bahan_non_sku=bnon))

@app.route('/produksi/template/<int:id>/hapus', methods=['POST'])
@inventory_required
@inventory_hpp_required
def produksi_template_hapus(id):
    conn = db()
    row = conn.execute("SELECT nama FROM produksi_template WHERE id=?", (id,)).fetchone()
    if row:
        conn.execute("DELETE FROM produksi_template WHERE id=?", (id,))
        add_log(conn, 'Hapus template produksi', row['nama'], 'HAPUS')
        conn.commit()
    conn.close()
    if request.is_json or request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return jsonify(ok=True)
    flash('Template produksi dihapus.' if row else 'Template tidak ditemukan.', 'info')
    return redirect(url_for('produksi_list'))

@app.route('/produksi', methods=['POST'])
@inventory_required
@inventory_hpp_required
def produksi_simpan():
    conn = db()
    tanggal = request.form.get('tanggal', str(date.today()))
    keterangan = (request.form.get('keterangan', '') or '').strip()
    if not _operator_date_ok(tanggal):
        conn.close()
        flash('Operator hanya bisa mencatat produksi untuk bulan berjalan.', 'warning')
        return redirect(url_for('produksi_list'))

    # ── Produk jadi (pilih yang ada ATAU buat baru) ──
    try:
        qty_hasil = parse_qty(request.form.get('qty_hasil', 0))
    except (TypeError, ValueError):
        conn.close()
        flash('Jumlah hasil produksi harus valid.', 'danger')
        return redirect(url_for('produksi_list'))
    if qty_hasil <= 0 or not math.isfinite(qty_hasil):
        conn.close()
        flash('Jumlah hasil produksi harus lebih dari 0.', 'warning')
        return redirect(url_for('produksi_list'))

    mode_jadi = request.form.get('produk_jadi_mode', 'existing')
    produk_jadi_id = 0
    if mode_jadi == 'baru':
        nama_baru = (request.form.get('produk_jadi_nama', '') or '').strip()
        satuan_baru = (request.form.get('produk_jadi_satuan', '') or '').strip() or 'porsi'
        try:
            jual_baru = parse_nonnegative_rp(request.form.get('produk_jadi_harga_jual', '0'),
                                             'Harga jual produk jadi')
        except ValueError as ex:
            conn.close(); flash(str(ex), 'danger')
            return redirect(url_for('produksi_list'))
        if not nama_baru:
            conn.close()
            flash('Nama produk jadi baru wajib diisi.', 'warning')
            return redirect(url_for('produksi_list'))
        import re as _re
        kode_base = _re.sub(r'[^A-Z0-9]', '-', nama_baru.upper())[:14].strip('-') or 'PRODUK'
        kode_new = kode_base; _sfx = 1
        while conn.execute('SELECT id FROM produk WHERE kode=?', (kode_new,)).fetchone():
            kode_new = f'{kode_base}-{_sfx}'; _sfx += 1
        conn.execute(
            'INSERT INTO produk(kode,nama,satuan,harga_beli,harga_jual,stok,min_stok) VALUES(?,?,?,0,?,0,0)',
            (kode_new, nama_baru, satuan_baru, jual_baru)
        )
        produk_jadi_id = conn.execute('SELECT last_insert_rowid()').fetchone()[0]
    else:
        try:
            produk_jadi_id = int(request.form.get('produk_jadi_id', 0) or 0)
        except (TypeError, ValueError):
            conn.close()
            flash('Produk jadi harus valid.', 'danger')
            return redirect(url_for('produksi_list'))
    produk_jadi = conn.execute(
        "SELECT id,nama,varian,satuan,stok,harga_beli FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
        (produk_jadi_id,)
    ).fetchone()
    if not produk_jadi:
        conn.rollback(); conn.close()
        flash('Produk jadi tidak ditemukan atau sudah diarsipkan.', 'danger')
        return redirect(url_for('produksi_list'))

    # ── Biaya tambahan (tenaga kerja / overhead), opsional ──
    try:
        biaya_tambahan = parse_nonnegative_rp(request.form.get('biaya_tambahan', '0'),
                                              'Biaya produksi tambahan')
    except ValueError as ex:
        conn.close()
        flash(str(ex), 'danger')
        return redirect(url_for('produksi_list'))
    akun_kas = request.form.get('akun_kas', '1100')
    if biaya_tambahan > 0 and not is_rekening_kode(conn, akun_kas):
        conn.close()
        flash('Rekening kas/bank untuk biaya produksi tidak valid.', 'danger')
        return redirect(url_for('produksi_list'))

    # ── Bahan baku (multi-row) ──
    bahan_ids = request.form.getlist('bahan_id[]')
    bahan_qtys = request.form.getlist('bahan_qty[]')
    bahan_items = []  # (produk_row, qty, harga_beli, subtotal)
    total_bahan = 0.0
    for bid, bq in zip(bahan_ids, bahan_qtys):
        if not bid:
            continue
        try:
            pid = int(bid)
            qb = parse_qty(bq)
        except (TypeError, ValueError):
            conn.close()
            flash('Qty bahan baku tidak valid.', 'danger')
            return redirect(url_for('produksi_list'))
        if qb <= 0:
            continue
        if pid == produk_jadi_id:
            conn.close()
            flash('Bahan baku tidak boleh sama dengan produk jadi.', 'warning')
            return redirect(url_for('produksi_list'))
        prow = conn.execute(
            "SELECT id,nama,varian,satuan,stok,harga_beli FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
            (pid,)
        ).fetchone()
        if not prow:
            conn.close()
            flash('Salah satu bahan baku tidak ditemukan atau sudah diarsipkan.', 'danger')
            return redirect(url_for('produksi_list'))
        # Stok bahan minus diizinkan (pakai dulu, bahan menyusul saat diterima).
        # Peringatan tampil setelah simpan dan di halaman inventory — bukan blokir.
        harga = float(prow['harga_beli'] or 0)
        subtotal = qb * harga
        total_bahan += subtotal
        bahan_items.append((prow, qb, harga, subtotal))

    non_sku_descs = request.form.getlist('non_sku_deskripsi[]')
    non_sku_nilais = request.form.getlist('non_sku_nilai[]')
    non_sku_qtys = request.form.getlist('non_sku_qty[]')
    non_sku_satuans = request.form.getlist('non_sku_satuan[]')
    non_sku_items = []  # (deskripsi, nilai, qty, satuan)
    total_non_sku = 0.0
    n_non_sku = max(len(non_sku_descs), len(non_sku_nilais))
    for i in range(n_non_sku):
        desc = (non_sku_descs[i] if i < len(non_sku_descs) else '').strip()
        raw_nilai = non_sku_nilais[i] if i < len(non_sku_nilais) else '0'
        raw_qty = (non_sku_qtys[i] if i < len(non_sku_qtys) else '').strip()
        satuan = (non_sku_satuans[i] if i < len(non_sku_satuans) else '').strip()
        try:
            nilai = parse_nonnegative_rp(raw_nilai, f'Nilai bahan non-SKU baris {i + 1}')
        except ValueError as ex:
            conn.close()
            flash(str(ex), 'danger')
            return redirect(url_for('produksi_list'))
        # Qty opsional — kalau diisi, harus angka >= 0; kosong dianggap 0
        try:
            qty = float(raw_qty) if raw_qty else 0.0
        except (TypeError, ValueError):
            conn.close()
            flash(f'Qty bahan non-SKU baris {i + 1} harus berupa angka.', 'warning')
            return redirect(url_for('produksi_list'))
        if qty < 0:
            conn.close()
            flash(f'Qty bahan non-SKU baris {i + 1} tidak boleh negatif.', 'warning')
            return redirect(url_for('produksi_list'))
        if nilai <= 0:
            if desc:
                conn.close()
                flash(f'Nilai bahan non-SKU baris {i + 1} harus lebih dari 0.', 'warning')
                return redirect(url_for('produksi_list'))
            continue
        if not desc:
            conn.close()
            flash(f'Deskripsi bahan non-SKU baris {i + 1} wajib diisi.', 'warning')
            return redirect(url_for('produksi_list'))
        total_non_sku += nilai
        non_sku_items.append((desc, nilai, qty, satuan))

    # Nilai persediaan non-SKU boleh minus (pakai dulu, bahan menyusul) — tidak diblokir.

    if not bahan_items and not non_sku_items:
        conn.close()
        flash('Isi minimal satu bahan baku SKU atau satu bahan non-SKU.', 'warning')
        return redirect(url_for('produksi_list'))

    total_hpp = total_bahan + total_non_sku + biaya_tambahan
    if total_hpp <= 0:
        conn.close()
        flash('Total nilai produksi (bahan + biaya) harus lebih dari 0. Pastikan bahan baku punya HPP.', 'warning')
        return redirect(url_for('produksi_list'))
    hpp_per_unit = total_hpp / qty_hasil

    # ── Jurnal ──
    # Bahan baku SKU, non-SKU, dan produk jadi sama-sama di akun 1130 Persediaan,
    # jadi nilai bahan "berpindah" di dalam 1130. Biaya tambahan (kas) menambah 1130.
    #   Dr 1130 (total_hpp)  [nilai produk jadi masuk]
    #   Cr 1130 (total_bahan)[nilai bahan SKU keluar]
    #   Cr 1130 (total_non_sku)[nilai bahan non-SKU keluar]
    #   Cr Kas (biaya_tambahan) [jika ada]
    entries = [('1130', total_hpp, 0)]
    if total_bahan > 0:
        entries.append(('1130', 0, total_bahan))
    if total_non_sku > 0:
        entries.append(('1130', 0, total_non_sku))
    if biaya_tambahan > 0:
        entries.append((akun_kas, 0, biaya_tambahan))
    ket_jurnal = keterangan or f"Produksi: {produk_jadi['nama']}" + (f" ({produk_jadi['varian']})" if produk_jadi['varian'] else "")
    try:
        jid = insert_jurnal(conn, tanggal, ket_jurnal, 'OPERASIONAL', 'PRODUKSI', entries)
    except ValueError as ex:
        conn.rollback(); conn.close()
        flash(f'Gagal membuat jurnal produksi: {ex}', 'danger')
        return redirect(url_for('produksi_list'))

    # ── Insert header produksi ──
    conn.execute(
        """INSERT INTO produksi(tanggal,produk_jadi_id,qty_hasil,biaya_tambahan,total_hpp,hpp_per_unit,keterangan,jurnal_id)
           VALUES(?,?,?,?,?,?,?,?)""",
        (tanggal, produk_jadi_id, qty_hasil, biaya_tambahan, total_hpp, hpp_per_unit, keterangan, jid)
    )
    produksi_id = conn.execute('SELECT last_insert_rowid()').fetchone()[0]

    for desc, nilai, qty, satuan in non_sku_items:
        # Kalau qty/satuan diisi, sisipkan ke keterangan pergerakan agar histori jelas
        qty_label = ''
        if qty > 0:
            qty_label = f" ({qty:g}{(' ' + satuan) if satuan else ''})"
        record_non_sku_movement(conn, jid, tanggal,
                                f"Dipakai produksi: {produk_jadi['nama']} - {desc}{qty_label}",
                                -nilai)
        conn.execute(
            "INSERT INTO produksi_non_sku(produksi_id,deskripsi,nilai,qty,satuan) VALUES(?,?,?,?,?)",
            (produksi_id, desc, nilai, qty, satuan)
        )

    # ── Bahan baku keluar (kurangi stok, catat pergerakan) ──
    for prow, qb, harga, subtotal in bahan_items:
        stok_lama = float(prow['stok'] or 0)
        conn.execute("UPDATE produk SET stok=stok-? WHERE id=?", (qb, prow['id']))
        record_stock_movement(conn, prow['id'], tanggal, 'KELUAR', qb, harga,
                              f"Dipakai produksi: {produk_jadi['nama']}", jid,
                              'REMOVE_LAYER', stok_lama, harga)
        conn.execute(
            "INSERT INTO produksi_bahan(produksi_id,produk_id,qty,harga,subtotal) VALUES(?,?,?,?,?)",
            (produksi_id, prow['id'], qb, harga, subtotal)
        )

    # ── Produk jadi masuk (tambah stok + update HPP rata-rata) ──
    stok_jadi_lama = float(produk_jadi['stok'] or 0)
    hpp_jadi_lama = float(produk_jadi['harga_beli'] or 0)
    update_average_cost(conn, produk_jadi_id, qty_hasil, hpp_per_unit)
    conn.execute("UPDATE produk SET stok=stok+? WHERE id=?", (qty_hasil, produk_jadi_id))
    record_stock_movement(conn, produk_jadi_id, tanggal, 'MASUK', qty_hasil, hpp_per_unit,
                          f"Hasil produksi (HPP/unit Rp {hpp_per_unit:,.0f})", jid,
                          'ADD_LAYER', stok_jadi_lama, hpp_jadi_lama)

    add_log(conn, 'Produksi',
            f"{produk_jadi['nama']} +{qty_hasil:g} {produk_jadi['satuan']} | HPP Rp {total_hpp:,.0f}",
            'INPUT')
    neg_notice = negative_stock_notice(conn, [b[0]['id'] for b in bahan_items])
    # Kalimat "menjadi minus" hanya jujur kalau produksi INI memakai bahan non-SKU.
    # Produksi murni SKU tidak mengubah saldo pool, jadi mengulang peringatan itu
    # di setiap produksi SKU menyesatkan user yang memang tidak memakai non-SKU.
    # Saldo minusnya tetap terlihat di lencana "Saldo non-SKU" halaman Produksi.
    nonsku_minus = total_non_sku > 0 and current_non_sku_value(conn) < -0.01
    conn.commit(); conn.close()
    flash(f"Produksi dicatat: {produk_jadi['nama']} +{qty_hasil:g} {produk_jadi['satuan']} "
          f"(HPP/unit Rp {hpp_per_unit:,.0f}).", 'success')
    if neg_notice:
        flash(neg_notice, 'warning')
    if nonsku_minus:
        flash('Nilai persediaan non-SKU menjadi minus. Segera lengkapi pembelian/koreksi '
              'persediaan non-SKU agar valuasi akurat.', 'warning')
    return redirect(url_for('produksi_list'))

@app.route('/produksi/<int:id>/hapus', methods=['POST'])
@finance_required
@inventory_hpp_required
def produksi_hapus(id):
    conn = db()
    pr = conn.execute("SELECT * FROM produksi WHERE id=?", (id,)).fetchone()
    if not pr:
        conn.close()
        flash('Data produksi tidak ditemukan.', 'danger')
        return redirect(url_for('produksi_list'))
    if pr['jurnal_id']:
        if has_later_stock_movements(conn, pr['jurnal_id']):
            conn.close()
            flash('Produksi ini memiliki mutasi stok yang lebih baru. Hapus mutasi stok terbaru lebih dulu agar stok tetap akurat.', 'warning')
            return redirect(url_for('produksi_list'))
        # _reverse_tx mengembalikan stok bahan & menghapus stok jadi, lalu hapus produksi
        _reverse_tx(conn, pr['jurnal_id'])
        conn.execute("DELETE FROM jurnal WHERE id=?", (pr['jurnal_id'],))
    else:
        conn.execute("DELETE FROM produksi_bahan WHERE produksi_id=?", (id,))
        conn.execute("DELETE FROM produksi_non_sku WHERE produksi_id=?", (id,))
        conn.execute("DELETE FROM produksi WHERE id=?", (id,))
    add_log(conn, 'Hapus produksi', f"Produksi #{id} dibatalkan", 'HAPUS')
    conn.commit(); conn.close()
    flash('Produksi dibatalkan. Stok bahan dikembalikan dan stok hasil dihapus.', 'warning')
    return redirect(url_for('produksi_list'))


# ---------- INVOICE ----------

# Format cetak invoice. SATU daftar untuk setelan, pemilih di halaman cetak, dan
# ukuran kertas yang dipakai. Kuncinya yang disimpan di settings `inv_format_cetak`.
#  - layout 'invoice' = lembar penuh; 'struk' = tata letak struk sempit.
#  - A4 SENGAJA tanpa page/margin di sini: nilainya hidup di @page stylesheet
#    invoice_print.html. Kalau ditulis dua kali, mengubah satu diam-diam kalah.
#  - A4 Landscape membawa page/margin sendiri (seperti dot matrix: menimpa @page
#    stylesheet lewat style suntikan JS). Penataan blok bawahnya ada di CSS
#    body[data-format="a4_landscape"] invoice_print.html; ganti kunci = ganti di sana.
#  - Thermal memakai lebar AREA CETAK (POS_STRUK_AREA_MM), bukan lebar kertas:
#    kertas 58/80 mm tak bisa dicetak penuh karena margin mekanis printer.
#  - Dot matrix dicetak browser dalam mode GRAFIS, jadi lambat & tak bisa dipercepat
#    dari halaman web. Yang diberikan: tata letak ramah dot matrix (monospace,
#    hitam pekat, tanpa latar, tanpa logo), bukan kecepatan mode teks.
# Didaftarkan sebagai Jinja global karena invoice_print.html dirender dari DUA rute
# (cetak & preview); kalau dikirim lewat argumen, satu rute pasti suatu saat terlewat.
INV_FORMAT_CETAK = {
    'a4':        {'label': 'A4 Portrait (printer laser / inkjet)', 'layout': 'invoice'},
    'a4_landscape': {'label': 'A4 Landscape (printer laser / inkjet)', 'layout': 'invoice',
                     'page': 'A4 landscape', 'margin': '10mm 16mm'},
    'thermal58': {'label': 'Thermal 58 mm', 'layout': 'struk',
                  'lebar_mm': POS_STRUK_AREA_MM['58']},
    'thermal80': {'label': 'Thermal 80 mm', 'layout': 'struk',
                  'lebar_mm': POS_STRUK_AREA_MM['80']},
    'dm_full':   {'label': 'Dot matrix, continuous 9,5 x 11 inci', 'layout': 'invoice',
                  'page': '9.5in 11in', 'margin': '10mm 12mm', 'dot_matrix': True},
    'dm_half':   {'label': 'Dot matrix, setengah halaman 9,5 x 5,5 inci', 'layout': 'invoice',
                  'page': '9.5in 5.5in', 'margin': '7mm 10mm', 'dot_matrix': True},
}
INV_FORMAT_CETAK_DEFAULT = 'a4'
app.jinja_env.globals['INV_FORMAT_CETAK'] = INV_FORMAT_CETAK


def _inv_format_cetak(nilai):
    """Kunci format cetak yang sah, atau default untuk nilai kosong/ngawur."""
    nilai = (nilai or '').strip()
    return nilai if nilai in INV_FORMAT_CETAK else INV_FORMAT_CETAK_DEFAULT


def _inv_settings(conn):
    keys = ['inv_nama','inv_alamat','inv_telepon','inv_email',
            'inv_rek','inv_top','inv_top_note','inv_catatan','inv_prefix',
            'inv_logo','inv_tagline','inv_terms']
    d = {k: get_setting(conn, k, '') for k in keys}
    d['inv_nama']   = d['inv_nama']   or get_setting(conn, 'nama_usaha', 'Usaha Saya')
    d['inv_prefix'] = d['inv_prefix'] or 'INV'
    d['inv_top']    = d['inv_top']    or '30'
    d['inv_format_cetak'] = _inv_format_cetak(get_setting(conn, 'inv_format_cetak', ''))
    try:
        d['inv_rek_list'] = json.loads(d['inv_rek']) if d['inv_rek'] else []
    except Exception:
        d['inv_rek_list'] = []
    return d

def _calc_top(inv):
    """Return actual ToP days: derived from date diff if both dates exist, else stored top_hari."""
    try:
        if inv['jatuh_tempo'] and inv['tanggal']:
            return (date.fromisoformat(inv['jatuh_tempo']) - date.fromisoformat(inv['tanggal'])).days
    except Exception:
        pass
    return inv['top_hari']

def _next_inv_number(conn, prefix):
    ym = date.today().strftime('%Y%m')
    last = conn.execute(
        "SELECT nomor FROM invoice WHERE nomor LIKE ? ORDER BY id DESC LIMIT 1",
        (f"{prefix}-{ym}-%",)
    ).fetchone()
    seq = int(last['nomor'].rsplit('-', 1)[-1]) + 1 if last else 1
    return f"{prefix}-{ym}-{seq:03d}"

def _po_settings(conn):
    keys = ['po_nama','po_alamat','po_telepon','po_email',
            'po_prefix','po_tagline','po_terms','po_catatan','po_logo']
    d = {k: get_setting(conn, k, '') for k in keys}
    d['po_nama']   = d['po_nama']   or get_setting(conn, 'nama_usaha', 'Usaha Saya')
    d['po_prefix'] = d['po_prefix'] or 'PO'
    return d

def _next_po_number(conn, prefix):
    ym = date.today().strftime('%Y%m')
    last = conn.execute(
        "SELECT nomor FROM purchase_order WHERE nomor LIKE ? ORDER BY id DESC LIMIT 1",
        (f"{prefix}-{ym}-%",)
    ).fetchone()
    seq = int(last['nomor'].rsplit('-', 1)[-1]) + 1 if last else 1
    return f"{prefix}-{ym}-{seq:03d}"

def _save_po_items(conn, po_id, form):
    conn.execute("DELETE FROM po_item WHERE po_id=?", (po_id,))
    produk_ids = form.getlist('produk_id[]')
    desks = form.getlist('deskripsi[]')
    qtys  = form.getlist('qty[]')
    sats  = form.getlist('satuan[]')
    hargs = form.getlist('harga_satuan[]')
    diss  = form.getlist('diskon_item[]')
    saved = 0
    row_count = max(len(produk_ids), len(desks), len(qtys), len(sats), len(hargs), len(diss))
    for i in range(row_count):
        pid = None
        prod = None
        raw_pid = (produk_ids[i] if i < len(produk_ids) else '').strip()
        if raw_pid:
            try:
                pid = int(raw_pid)
            except (TypeError, ValueError):
                raise ValueError(f'Produk item PO baris {i+1} tidak valid.')
            prod = conn.execute(
                "SELECT id,nama,varian,satuan,harga_beli FROM produk WHERE id=?",
                (pid,)
            ).fetchone()
            if not prod:
                raise ValueError(f'Produk item PO baris {i+1} tidak ditemukan.')

        desk = (desks[i] if i < len(desks) else '').strip()
        if prod and not desk:
            desk = prod['nama'] + (f" - {prod['varian']}" if prod['varian'] else '')
        if not desk:
            continue
        try:
            qty  = parse_qty(qtys[i] if i < len(qtys) else 1, 1)
            default_harga = prod['harga_beli'] if prod else 0
            harg = parse_rp(hargs[i] if i < len(hargs) and hargs[i] != '' else default_harga)
            dis_i= parse_rp(diss[i]  if i < len(diss)  else 0)
        except (TypeError, ValueError):
            raise ValueError(f'Angka item PO baris {i+1} tidak valid.')
        if qty <= 0:
            raise ValueError(f'Qty item PO baris {i+1} harus lebih dari 0.')
        if harg < 0:
            raise ValueError(f'Harga item PO baris {i+1} tidak boleh negatif.')
        if dis_i < 0:
            raise ValueError(f'Diskon item PO baris {i+1} tidak boleh negatif.')
        bruto = qty * harg
        if dis_i > bruto + 0.01:
            raise ValueError(f'Diskon item PO baris {i+1} tidak boleh melebihi nilai barang.')
        sat = (sats[i] if i < len(sats) else (prod['satuan'] if prod else 'pcs')).strip() or 'pcs'
        subtotal = max(0, bruto - dis_i)
        conn.execute(
            """INSERT INTO po_item(
                   po_id,produk_id,deskripsi,qty,satuan,harga_satuan,diskon_item,subtotal
               ) VALUES(?,?,?,?,?,?,?,?)""",
            (po_id, pid, desk.strip(), qty, sat, harg, dis_i, subtotal))
        saved += 1
    if not saved:
        raise ValueError('Minimal 1 item PO harus diisi.')

def _po_produk_list(conn, po_id=None):
    if po_id:
        return conn.execute(
            """SELECT DISTINCT p.*
               FROM produk p
               LEFT JOIN po_item pi ON pi.produk_id=p.id AND pi.po_id=?
               WHERE COALESCE(p.aktif,1)=1 OR pi.id IS NOT NULL
               ORDER BY COALESCE(p.aktif,1) DESC, p.nama, COALESCE(p.varian,'')""",
            (po_id,)
        ).fetchall()
    return conn.execute(
        "SELECT * FROM produk WHERE COALESCE(aktif,1)=1 ORDER BY nama, COALESCE(varian,'')"
    ).fetchall()

@app.route('/invoice')
@invoice_read_required
def invoice_list():
    conn = db()
    q  = request.args.get('q','')
    sf = request.args.get('status','')
    where = " WHERE 1=1"
    params = []
    if q:
        where += " AND (pelanggan LIKE ? OR nomor LIKE ?)"
        params += [f'%{q}%', f'%{q}%']
    if sf:
        where += " AND status=?"
        params.append(sf)
    page = _page_arg()
    n_rows = conn.execute("SELECT COUNT(*) FROM invoice" + where, params).fetchone()[0]
    pg = paginate_meta(page, n_rows)
    invoices = conn.execute(
        "SELECT * FROM invoice" + where + " ORDER BY tanggal DESC, id DESC LIMIT ? OFFSET ?",
        params + [pg['per_page'], pg['offset']]).fetchall()
    conn.close()
    return render_template('invoice_list.html', invoices=invoices, q=q, sf=sf, pg=pg)

@app.route('/invoice/pengaturan', methods=['GET','POST'])
@admin_required
def invoice_pengaturan():
    conn = db()
    if request.method == 'POST':
        for k in ['inv_nama','inv_alamat','inv_telepon','inv_email',
                  'inv_top','inv_top_note','inv_catatan','inv_prefix',
                  'inv_tagline','inv_terms']:
            conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                         (k, request.form.get(k,'')))
        # Format cetak HANYA disimpan kalau field-nya memang dikirim. Form lain yang
        # kebetulan POST ke sini tanpa field ini tidak boleh mereset pilihan ke A4.
        _fmt = request.form.get('inv_format_cetak')
        if _fmt is not None:
            conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                         ('inv_format_cetak', _inv_format_cetak(_fmt)))
        rek_list = []
        for i in range(1, 4):
            bank = request.form.get(f'rek_bank_{i}','').strip()
            norek= request.form.get(f'rek_no_{i}','').strip()
            an   = request.form.get(f'rek_an_{i}','').strip()
            if bank or norek:
                rek_list.append({'bank': bank, 'norek': norek, 'an': an})
        conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                     ('inv_rek', json.dumps(rek_list)))
        # Logo upload
        logo_file = request.files.get('inv_logo_file')
        if logo_file and logo_file.filename:
            ext = logo_file.filename.rsplit('.', 1)[-1].lower()
            if ext in ALLOWED_IMG:
                for f in os.listdir(UPLOAD_FOLDER):
                    if f.startswith('inv_logo.'):
                        try: os.remove(os.path.join(UPLOAD_FOLDER, f))
                        except: pass
                logo_fn = f'inv_logo.{ext}'
                logo_file.save(os.path.join(UPLOAD_FOLDER, logo_fn))
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                             ('inv_logo', logo_fn))
        elif request.form.get('inv_logo_delete') == '1':
            old = get_setting(conn, 'inv_logo', '')
            if old:
                try: os.remove(os.path.join(UPLOAD_FOLDER, old))
                except: pass
            conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                         ('inv_logo', ''))
        conn.commit(); conn.close()
        flash('Template invoice disimpan!', 'success')
        return redirect(url_for('invoice_pengaturan'))
    inv_s = _inv_settings(conn)
    conn.close()
    return render_template('invoice_pengaturan.html', inv_s=inv_s)

@app.route('/pengaturan/modul', methods=['GET','POST'])
@admin_required
def pengaturan_modul():
    conn = db()
    if request.method == 'POST':
        action = request.form.get('action','').strip()
        key    = request.form.get('modul_key','').strip()
        if action == 'aktivasi':
            serial = request.form.get('serial','').strip()
            ok, msg = activate_module(conn, key, serial)
            if ok:
                add_log(conn, 'Aktivasi Modul PRO', f'{key} | aktif', 'PENGATURAN')
                conn.commit()
                flash(msg, 'success')
            else:
                flash(msg, 'danger')
        elif action == 'nonaktif':
            deactivate_module(conn, key)
            add_log(conn, 'Non-aktifkan Modul PRO', f'{key}', 'PENGATURAN')
            conn.commit()
            flash('Modul dinon-aktifkan.', 'info')
        conn.close()
        return redirect(url_for('pengaturan_modul'))
    # GET
    modul_info = []
    for k, m in MODUL_PRO.items():
        modul_info.append({
            'key': k,
            'nama': m['nama'],
            'deskripsi': m['deskripsi'],
            'aktif': is_module_active(conn, k),
            'tgl_aktif': get_setting(conn, f'modul_{k}_aktif_tgl', ''),
        })
    conn.close()
    return render_template('modul_pro.html', modul_list=modul_info)

# ═══════════════════════════════════════════════════════════════════════════
#   MODUL PRO: PPN & FAKTUR OTOMATIS, Routes
# ═══════════════════════════════════════════════════════════════════════════

def _require_pajak_module():
    """Pastikan modul PAJAK_OTOMATIS aktif sebelum akses route PRO."""
    conn = db()
    aktif = is_module_active(conn, 'PAJAK_OTOMATIS')
    conn.close()
    if not aktif:
        flash('Modul Lanjutan "Pajak & Gaji" belum aktif. Aktivasi dulu di Pengaturan > Modul Lanjutan.', 'warning')
        return redirect(url_for('pengaturan_modul'))
    return None

def _akun_saldo(conn, kode):
    """Saldo akun (Dr − Cr untuk akun debit, Cr − Dr untuk akun kredit). Selalu positif kalau wajar."""
    row = conn.execute("""
        SELECT a.saldo_normal,
               COALESCE(SUM(dj.debit),0)  AS dt,
               COALESCE(SUM(dj.kredit),0) AS kt
          FROM akun a LEFT JOIN detail_jurnal dj ON dj.akun_id=a.id
         WHERE a.kode=? GROUP BY a.id
    """, (kode,)).fetchone()
    if not row: return 0.0
    return float(row['dt'] - row['kt']) if row['saldo_normal'] == 'DEBIT' else float(row['kt'] - row['dt'])

def _akun_aktivitas_periode(conn, kode, from_d, to_d):
    """Total debit & kredit ke akun antara from_d..to_d. Return dict {'debit':x,'kredit':y}."""
    row = conn.execute("""
        SELECT COALESCE(SUM(dj.debit),0) AS dt, COALESCE(SUM(dj.kredit),0) AS kt
          FROM detail_jurnal dj
          JOIN akun a ON a.id=dj.akun_id
          JOIN jurnal j ON j.id=dj.jurnal_id
         WHERE a.kode=? AND j.tanggal BETWEEN ? AND ?
    """, (kode, from_d, to_d)).fetchone()
    return {'debit': float(row['dt'] or 0), 'kredit': float(row['kt'] or 0)}

@app.route('/pajak')
@finance_required
def pajak_hub():
    r = _require_pajak_module()
    if r: return r
    conn = db()
    saldo = {
        'hutang_ppn':    _akun_saldo(conn, '2111'),
        'hutang_pph23':  _akun_saldo(conn, '2112'),
        'hutang_pph_bd': _akun_saldo(conn, '2110'),
        'ppn_masukan':   _akun_saldo(conn, '1180'),
        'pph23_dimuka':  _akun_saldo(conn, '1181'),
        'angsuran25':    _akun_saldo(conn, '1182'),
    }
    # PPN bersih bulan ini
    today = date.today()
    bln_awal = today.replace(day=1).strftime('%Y-%m-%d')
    bln_akhir = today.strftime('%Y-%m-%d')
    ppn_keluaran = _akun_aktivitas_periode(conn, '2111', bln_awal, bln_akhir)['kredit']
    ppn_masukan_periode = _akun_aktivitas_periode(conn, '1180', bln_awal, bln_akhir)['debit']
    ppn_kurang_bayar = max(0, ppn_keluaran - ppn_masukan_periode)
    # Count bukti potong outstanding
    n_bp = conn.execute("SELECT COUNT(*) FROM invoice WHERE COALESCE(pph23_nominal,0) > 0").fetchone()[0]
    # Gaji bulan ini
    saldo['hutang_pph21']  = _akun_saldo(conn, '2115')
    saldo['hutang_bpjs']   = _akun_saldo(conn, '2116')
    saldo['hutang_gaji']   = _akun_saldo(conn, '2120')
    gaji_bulan = conn.execute("""
        SELECT COUNT(*) AS n_slip,
               COALESCE(SUM(penghasilan_bruto),0) AS total_bruto,
               COALESCE(SUM(gaji_bersih),0)      AS total_bersih,
               COALESCE(SUM(pph21),0)             AS total_pph21
        FROM payroll
        WHERE periode_bulan=? AND periode_tahun=?
    """, (today.month, today.year)).fetchone()
    n_karyawan_aktif = conn.execute("SELECT COUNT(*) FROM karyawan WHERE aktif=1").fetchone()[0]
    conn.close()
    return render_template('pajak_hub.html', saldo=saldo,
                           ppn_keluaran=ppn_keluaran, ppn_masukan_periode=ppn_masukan_periode,
                           ppn_kurang_bayar=ppn_kurang_bayar, bln=today.strftime('%B %Y'),
                           n_bukti_potong=n_bp,
                           gaji_bulan=gaji_bulan, n_karyawan_aktif=n_karyawan_aktif)

@app.route('/pajak/setor', methods=['GET','POST'])
@finance_required
def pajak_setor():
    r = _require_pajak_module()
    if r: return r
    conn = db()
    AKUN_PAJAK_LIAB = ['2111','2112','2110','2115','2116','2120']  # Pajak, BPJS, dan gaji terutang
    if request.method == 'POST':
        tanggal = request.form.get('tanggal', date.today().strftime('%Y-%m-%d'))
        if not is_valid_iso_date(tanggal):
            flash('Tanggal tidak valid.', 'danger'); conn.close()
            return redirect(url_for('pajak_setor'))
        akun_hutang_kode = request.form.get('akun_hutang','').strip()
        akun_kas_kode    = request.form.get('akun_kas','').strip()
        try:
            nominal = parse_nonnegative_rp(request.form.get('nominal', '0'), 'Nominal pembayaran')
        except ValueError as ex:
            flash(str(ex), 'danger'); conn.close()
            return redirect(url_for('pajak_setor'))
        if akun_hutang_kode not in AKUN_PAJAK_LIAB:
            flash('Akun hutang pajak/gaji tidak valid.', 'danger'); conn.close()
            return redirect(url_for('pajak_setor'))
        if not is_rekening_kode(conn, akun_kas_kode):
            flash('Rekening kas/bank tidak valid.', 'danger'); conn.close()
            return redirect(url_for('pajak_setor'))
        if nominal <= 0:
            flash('Nominal pembayaran harus > 0.', 'danger'); conn.close()
            return redirect(url_for('pajak_setor'))
        saldo_h = _akun_saldo(conn, akun_hutang_kode)

        # ── Akrual otomatis untuk PPh Badan (2110) jika belum dicatat ───────
        # Tanpa akrual, setor langsung membuat saldo 2110 negatif & beban pajak
        # tidak pernah masuk Laba Rugi. Kita akrual dulu sesuai best practice.
        if akun_hutang_kode == '2110':
            akrual_nm = max(0.0, nominal - max(0.0, saldo_h))
            if akrual_nm > 0:
                # Akun beban pajak dicari lewat PERAN, persis jalur Pengeluaran > Pajak.
                # Dulu di sini ditulis '6170' mentah: kalau nomor itu sudah dipakai ulang
                # (mis. Beban Pengiriman), akrual PPh Badan diam-diam masuk ke situ dan
                # Laba Rugi menghitungnya sebagai beban operasional.
                pajak_kode = _akun_beban_berperan(conn, 'PAJAK')
                if not pajak_kode:
                    flash(_pesan_tanpa_akun_berperan('Pajak'), 'danger'); conn.close()
                    return redirect(url_for('pajak_setor'))
                try:
                    insert_jurnal(conn, tanggal,
                        f"Akrual Beban PPh Badan (otomatis pra-setor) {akrual_nm:,.0f}",
                        'OPERASIONAL', 'PENYESUAIAN',
                        [(pajak_kode, akrual_nm, 0), ('2110', 0, akrual_nm)])
                    saldo_h = _akun_saldo(conn, akun_hutang_kode)
                except ValueError as ex:
                    conn.rollback()
                    flash(f'Gagal akrual PPh Badan: {ex}', 'danger'); conn.close()
                    return redirect(url_for('pajak_setor'))

        if akun_hutang_kode == '2120' and saldo_h <= 0:
            flash('Belum ada Hutang Gaji tercatat. Proses gaji dulu di modul Gaji → Proses Payroll, '
                  'baru kembali ke sini untuk mencatat pembayaran.', 'warning')
            conn.close()
            return redirect(url_for('gaji_proses'))
        if nominal > saldo_h + 1:
            flash(f'Nominal melebihi saldo hutang (Rp {saldo_h:,.0f}). Periksa lagi.', 'warning'); conn.close()
            return redirect(url_for('pajak_setor'))

        # ── Auto-offset PPN Masukan (1180) saat setor Hutang PPN (2111) ─────
        # PSAK: PPN Masukan dapat dikreditkan terhadap PPN Keluaran. Saat setor PPN,
        # saldo 1180 harus di-Cr untuk menutup posisinya. Tanpa ini, 1180 numpuk di
        # Neraca selamanya → Aset Lancar overstated.
        ppn_masukan_offset = 0.0
        kreditkan_masukan  = request.form.get('kreditkan_masukan') == '1'
        if akun_hutang_kode == '2111' and kreditkan_masukan:
            saldo_1180 = _akun_saldo(conn, '1180')
            # Offset = sisa hutang PPN setelah kas dibayar (yang ditutup PPN Masukan)
            ppn_masukan_offset = round(min(saldo_1180, max(0.0, saldo_h - nominal)), 2)

        akun_nama = conn.execute("SELECT nama FROM akun WHERE kode=?", (akun_hutang_kode,)).fetchone()['nama']
        keterangan = f"Pembayaran {akun_nama}"
        if ppn_masukan_offset > 0:
            keterangan += f" (kredit PPN Masukan Rp {ppn_masukan_offset:,.0f})"

        entries = [(akun_hutang_kode, nominal + ppn_masukan_offset, 0)]
        if ppn_masukan_offset > 0:
            entries.append(('1180', 0, ppn_masukan_offset))
        entries.append((akun_kas_kode, 0, nominal))

        try:
            jid = insert_jurnal(conn, tanggal, keterangan, 'OPERASIONAL', 'BAYAR', entries)
            add_log(conn, 'Pembayaran Hutang Pajak/Gaji',
                    f'{akun_nama} kas {nominal:,.0f}' +
                    (f' + offset 1180 {ppn_masukan_offset:,.0f}' if ppn_masukan_offset > 0 else ''),
                    'INPUT')
            conn.commit()
            msg = f'Pembayaran {akun_nama} berhasil dicatat (Jurnal #{jid}). Kas keluar Rp {nominal:,.0f}'
            if ppn_masukan_offset > 0:
                msg += f', PPN Masukan dikreditkan Rp {ppn_masukan_offset:,.0f}'
            flash(msg + '.', 'success')
        except ValueError as ex:
            conn.rollback()
            flash(str(ex), 'danger')
        conn.close()
        return redirect(url_for('pajak_hub'))
    # GET
    akun_hutang = conn.execute(
        "SELECT kode,nama FROM akun WHERE kode IN ('2110','2111','2112','2115','2116','2120') ORDER BY kode"
    ).fetchall()
    akun_hutang_with_saldo = [{'kode':a['kode'],'nama':a['nama'],'saldo':_akun_saldo(conn, a['kode'])} for a in akun_hutang]
    akun_kas = conn.execute("SELECT kode,nama FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    # Saldo PPN Masukan untuk preview auto-offset di UI
    saldo_ppn_masukan = _akun_saldo(conn, '1180')
    # Pre-fill dari query string (SPT Masa redirect)
    prefill_akun  = request.args.get('akun', '').strip()
    prefill_nomi  = request.args.get('nominal', '').strip()
    conn.close()
    return render_template('pajak_setor.html',
                           akun_hutang=akun_hutang_with_saldo, akun_kas=akun_kas,
                           saldo_ppn_masukan=saldo_ppn_masukan,
                           prefill_akun=prefill_akun, prefill_nomi=prefill_nomi,
                           today=date.today().strftime('%Y-%m-%d'))

@app.route('/pajak/pembelian-pkp', methods=['GET','POST'])
@finance_required
def pajak_pembelian_pkp():
    r = _require_pajak_module()
    if r: return r
    conn = db()
    if request.method == 'POST':
        tanggal = request.form.get('tanggal', date.today().strftime('%Y-%m-%d'))
        if not is_valid_iso_date(tanggal):
            flash('Tanggal tidak valid.','danger'); conn.close()
            return redirect(url_for('pajak_pembelian_pkp'))
        supplier      = request.form.get('supplier','').strip() or 'Supplier'
        keterangan    = request.form.get('keterangan','').strip() or 'Pembelian PKP'
        akun_debit    = request.form.get('akun_debit','').strip()      # ke mana DPP didebit (persediaan/beban)
        dpp           = parse_rp(request.form.get('dpp','0'))
        try:
            ppn_persen = parse_percentage(request.form.get('ppn_persen', '11'), 'Persen PPN')
            pph23_persen = parse_percentage(request.form.get('pph23_persen', '0'), 'Persen PPh 23')
        except ValueError as ex:
            flash(str(ex), 'danger'); conn.close()
            return redirect(url_for('pajak_pembelian_pkp'))
        ppn_persen = max(0, min(100, ppn_persen))
        pph23_persen = max(0, min(100, pph23_persen))  # perusahaan sebagai pemotong PPh 23
        sumber        = request.form.get('sumber','tunai')  # 'tunai' / 'hutang'
        akun_kas_kode = request.form.get('akun_kas','').strip()
        if dpp <= 0:
            flash('DPP harus > 0.','danger'); conn.close()
            return redirect(url_for('pajak_pembelian_pkp'))
        if not conn.execute("SELECT 1 FROM akun WHERE kode=?", (akun_debit,)).fetchone():
            flash('Akun debit (Persediaan/Beban) tidak valid.','danger'); conn.close()
            return redirect(url_for('pajak_pembelian_pkp'))
        if sumber == 'tunai' and not is_rekening_kode(conn, akun_kas_kode):
            flash('Rekening kas/bank tidak valid.','danger'); conn.close()
            return redirect(url_for('pajak_pembelian_pkp'))
        ppn_nom   = round(dpp * ppn_persen / 100.0, 2)
        pph_nom   = round(dpp * pph23_persen / 100.0, 2)
        total_kewajiban = dpp + ppn_nom - pph_nom
        entries = [(akun_debit, dpp, 0)]
        if ppn_nom > 0:
            entries.append(('1180', ppn_nom, 0))  # Dr PPN Masukan
        if sumber == 'tunai':
            entries.append((akun_kas_kode, 0, total_kewajiban))  # Cr Kas
        else:
            entries.append(('2100', 0, total_kewajiban))         # Cr Hutang Usaha
        if pph_nom > 0:
            entries.append(('2112', 0, pph_nom))  # Cr Hutang PPh 23 (kita sbg pemotong)
        try:
            jid = insert_jurnal(conn, tanggal, f"{keterangan} - {supplier}", 'OPERASIONAL', 'BELI',
                                entries, pihak=supplier)
            if sumber == 'hutang':
                jt = (date.fromisoformat(tanggal) + timedelta(days=30)).strftime('%Y-%m-%d')
                conn.execute(
                    "INSERT INTO hutang(tanggal,jatuh_tempo,pemasok,keterangan,jumlah,jurnal_id,akun_kode) "
                    "VALUES(?,?,?,?,?,?,'2100')",
                    (tanggal, jt, supplier, f"Pembelian PKP - {keterangan}", total_kewajiban, jid)
                )
            add_log(conn, 'Pembelian PKP', f"{supplier} | DPP {dpp:,.0f}", 'INPUT')
            conn.commit()
            flash(f'Pembelian PKP berhasil dicatat (Jurnal #{jid}). PPN Masukan +{ppn_nom:,.0f}'
                  + (f', Hutang PPh 23 +{pph_nom:,.0f}' if pph_nom > 0 else ''), 'success')
        except ValueError as ex:
            conn.rollback()
            flash(str(ex), 'danger')
        conn.close()
        return redirect(url_for('pajak_hub'))
    # GET
    akun_debit_opts = conn.execute(
        "SELECT kode,nama,tipe FROM akun WHERE tipe IN ('ASET','BEBAN') AND kode NOT IN ('1180','1181','1182','1100','1110','1120','1150') AND is_rekening=0 ORDER BY kode"
    ).fetchall()
    akun_kas = conn.execute("SELECT kode,nama FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    vendor_list = [dict(r) for r in conn.execute(
        "SELECT nama,telepon,alamat FROM vendor ORDER BY nama").fetchall()]
    conn.close()
    return render_template('pajak_pembelian.html', akun_debit_opts=akun_debit_opts,
                           akun_kas=akun_kas, today=date.today().strftime('%Y-%m-%d'),
                           vendor_list=vendor_list)

@app.route('/pajak/spt-masa')
@finance_required
def pajak_spt_masa():
    r = _require_pajak_module()
    if r: return r
    today = date.today()
    tahun = int(request.args.get('tahun', today.year))
    bulan = int(request.args.get('bulan', today.month))
    from_d = date(tahun, bulan, 1).strftime('%Y-%m-%d')
    # Akhir bulan
    if bulan == 12:
        to_d = date(tahun, 12, 31).strftime('%Y-%m-%d')
    else:
        to_d = (date(tahun, bulan+1, 1) - timedelta(days=1)).strftime('%Y-%m-%d')
    conn = db()
    # PPN Keluaran (kredit ke 2111 sepanjang bulan)
    ppn_keluaran = _akun_aktivitas_periode(conn, '2111', from_d, to_d)['kredit']
    # PPN Masukan (debit ke 1180 sepanjang bulan)
    ppn_masukan  = _akun_aktivitas_periode(conn, '1180', from_d, to_d)['debit']
    # Pelunasan: kredit ke 1180 (return atau koreksi), diabaikan untuk ringkasan
    # Setoran PPN bulan ini = debit ke 2111 (pelunasan hutang ke kas)
    setoran_ppn = _akun_aktivitas_periode(conn, '2111', from_d, to_d)['debit']
    selisih = ppn_keluaran - ppn_masukan
    kurang_bayar = max(0, selisih)
    lebih_bayar  = max(0, -selisih)
    # Detail rincian transaksi
    rincian_keluaran = conn.execute("""
        SELECT j.tanggal, j.keterangan, j.pihak, dj.kredit AS nominal
          FROM detail_jurnal dj
          JOIN jurnal j ON j.id=dj.jurnal_id
          JOIN akun a ON a.id=dj.akun_id
         WHERE a.kode='2111' AND dj.kredit > 0 AND j.tanggal BETWEEN ? AND ?
         ORDER BY j.tanggal, j.id
    """, (from_d, to_d)).fetchall()
    rincian_masukan = conn.execute("""
        SELECT j.tanggal, j.keterangan, j.pihak, dj.debit AS nominal
          FROM detail_jurnal dj
          JOIN jurnal j ON j.id=dj.jurnal_id
          JOIN akun a ON a.id=dj.akun_id
         WHERE a.kode='1180' AND dj.debit > 0 AND j.tanggal BETWEEN ? AND ?
         ORDER BY j.tanggal, j.id
    """, (from_d, to_d)).fetchall()
    conn.close()
    bulan_nama = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus',
                  'September','Oktober','November','Desember'][bulan-1]
    return render_template('pajak_spt_masa.html',
                           tahun=tahun, bulan=bulan, bulan_nama=bulan_nama,
                           from_d=from_d, to_d=to_d,
                           ppn_keluaran=ppn_keluaran, ppn_masukan=ppn_masukan,
                           setoran_ppn=setoran_ppn,
                           kurang_bayar=kurang_bayar, lebih_bayar=lebih_bayar,
                           rincian_keluaran=rincian_keluaran, rincian_masukan=rincian_masukan,
                           today=today)

@app.route('/pajak/spt-masa/export')
@finance_required
def pajak_spt_masa_export():
    """Export SPT Masa PPN (ringkasan + rincian PPN Keluaran/Masukan) 1 bulan ke Excel.
    Perhitungan IDENTIK dengan tampilan /pajak/spt-masa (periode bulan/tahun)."""
    r = _require_pajak_module()
    if r: return r
    today = date.today()
    tahun = int(request.args.get('tahun', today.year))
    bulan = int(request.args.get('bulan', today.month))
    if not (1 <= bulan <= 12):
        bulan = today.month
    from_d = date(tahun, bulan, 1).strftime('%Y-%m-%d')
    to_d = (date(tahun, 12, 31) if bulan == 12 else date(tahun, bulan + 1, 1) - timedelta(days=1)).strftime('%Y-%m-%d')
    conn = db()
    ppn_keluaran = _akun_aktivitas_periode(conn, '2111', from_d, to_d)['kredit']
    ppn_masukan  = _akun_aktivitas_periode(conn, '1180', from_d, to_d)['debit']
    setoran_ppn  = _akun_aktivitas_periode(conn, '2111', from_d, to_d)['debit']
    rincian_keluaran = conn.execute(
        "SELECT j.tanggal, j.keterangan, j.pihak, dj.kredit AS nominal "
        "FROM detail_jurnal dj JOIN jurnal j ON j.id=dj.jurnal_id JOIN akun a ON a.id=dj.akun_id "
        "WHERE a.kode='2111' AND dj.kredit>0 AND j.tanggal BETWEEN ? AND ? ORDER BY j.tanggal,j.id",
        (from_d, to_d)).fetchall()
    rincian_masukan = conn.execute(
        "SELECT j.tanggal, j.keterangan, j.pihak, dj.debit AS nominal "
        "FROM detail_jurnal dj JOIN jurnal j ON j.id=dj.jurnal_id JOIN akun a ON a.id=dj.akun_id "
        "WHERE a.kode='1180' AND dj.debit>0 AND j.tanggal BETWEEN ? AND ? ORDER BY j.tanggal,j.id",
        (from_d, to_d)).fetchall()
    conn.close()
    selisih = ppn_keluaran - ppn_masukan
    bulan_nama = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus',
                  'September','Oktober','November','Desember'][bulan - 1]
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    import io
    def _hdr(ws, rownum, ncol):
        for i in range(1, ncol + 1):
            c = ws.cell(row=rownum, column=i)
            c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='305496')
    wb = Workbook()
    ws = wb.active; ws.title = 'Ringkasan'
    ws.append([f'SPT Masa PPN - {bulan_nama} {tahun}']); ws.append([])
    ws.append(['PPN Keluaran', ppn_keluaran])
    ws.append(['PPN Masukan', ppn_masukan])
    ws.append(['Kurang Bayar' if selisih >= 0 else 'Lebih Bayar', abs(selisih)])
    ws.append(['Sudah Disetor Bulan Ini', setoran_ppn])
    for rr in range(3, 7):
        ws.cell(row=rr, column=2).number_format = '#,##0'
    ws.column_dimensions['A'].width = 26; ws.column_dimensions['B'].width = 18
    for title, data in (('PPN Keluaran', rincian_keluaran), ('PPN Masukan', rincian_masukan)):
        wsx = wb.create_sheet(title)
        wsx.append(['Tanggal', 'Keterangan', 'Pihak', 'Nominal PPN']); _hdr(wsx, 1, 4)
        tot = 0.0
        for row in data:
            n = float(row['nominal'] or 0); tot += n
            wsx.append([row['tanggal'], row['keterangan'] or '', row['pihak'] or '', n])
        wsx.append(['', '', 'TOTAL', tot]); wsx.cell(row=wsx.max_row, column=3).font = Font(bold=True)
        for i, w in enumerate([12, 44, 24, 18], 1):
            wsx.column_dimensions[chr(64 + i)].width = w
        for rr in range(2, wsx.max_row + 1):
            wsx.cell(row=rr, column=4).number_format = '#,##0'
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, download_name=f"spt_masa_ppn_{tahun}_{bulan:02d}.xlsx", as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/pajak/pph22')
@finance_required
def pajak_pph22():
    """Rekap pemotongan PPh 22 oleh instansi/bendaharawan (sisi penjual).
    PPh 22 di-claim sebagai kredit pajak di SPT Tahunan PPh badan."""
    r = _require_pajak_module()
    if r: return r
    today = date.today()
    tahun = int(request.args.get('tahun', today.year))
    bulan = int(request.args.get('bulan', today.month))
    from_d = date(tahun, bulan, 1).strftime('%Y-%m-%d')
    if bulan == 12:
        to_d = date(tahun, 12, 31).strftime('%Y-%m-%d')
    else:
        to_d = (date(tahun, bulan + 1, 1) - timedelta(days=1)).strftime('%Y-%m-%d')
    conn = db()
    # Periode ini
    aktivitas = _akun_aktivitas_periode(conn, '1183', from_d, to_d)
    pph22_dipotong = aktivitas['debit']    # Pemotongan oleh instansi (= kredit pajak baru)
    pph22_dibalik  = aktivitas['kredit']   # Pembalikan (retur/koreksi)
    pph22_netto    = pph22_dipotong - pph22_dibalik
    saldo_akumulatif = _akun_saldo(conn, '1183')  # akumulasi kredit pajak s.d. saat ini
    # Rincian transaksi pemotongan bulan ini
    rincian = conn.execute("""
        SELECT j.id, j.tanggal, j.keterangan, j.pihak, j.nomor_tx, dj.debit AS nominal
          FROM detail_jurnal dj
          JOIN jurnal j ON j.id = dj.jurnal_id
          JOIN akun a   ON a.id = dj.akun_id
         WHERE a.kode = '1183' AND dj.debit > 0
           AND j.tanggal BETWEEN ? AND ?
         ORDER BY j.tanggal, j.id
    """, (from_d, to_d)).fetchall()
    rincian_balik = conn.execute("""
        SELECT j.id, j.tanggal, j.keterangan, j.pihak, j.nomor_tx, dj.kredit AS nominal
          FROM detail_jurnal dj
          JOIN jurnal j ON j.id = dj.jurnal_id
          JOIN akun a   ON a.id = dj.akun_id
         WHERE a.kode = '1183' AND dj.kredit > 0
           AND j.tanggal BETWEEN ? AND ?
         ORDER BY j.tanggal, j.id
    """, (from_d, to_d)).fetchall()
    conn.close()
    bulan_nama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
                  'Agustus', 'September', 'Oktober', 'November', 'Desember'][bulan - 1]
    return render_template('pajak_pph22.html',
                           tahun=tahun, bulan=bulan, bulan_nama=bulan_nama,
                           from_d=from_d, to_d=to_d,
                           pph22_dipotong=pph22_dipotong,
                           pph22_dibalik=pph22_dibalik,
                           pph22_netto=pph22_netto,
                           saldo_akumulatif=saldo_akumulatif,
                           rincian=rincian, rincian_balik=rincian_balik,
                           today=today)

@app.route('/pajak/bukti-potong')
@finance_required
def pajak_bukti_potong_list():
    r = _require_pajak_module()
    if r: return r
    conn = db()
    page = _page_arg()
    n_rows = conn.execute("SELECT COUNT(*) FROM invoice WHERE COALESCE(pph23_nominal,0) > 0").fetchone()[0]
    pg = paginate_meta(page, n_rows)
    rows = conn.execute("""
        SELECT id, nomor, tanggal, pelanggan, alamat_pelanggan, pph23_persen, pph23_nominal, jurnal_id
          FROM invoice
         WHERE COALESCE(pph23_nominal,0) > 0
         ORDER BY tanggal DESC, id DESC LIMIT ? OFFSET ?
    """, (pg['per_page'], pg['offset'])).fetchall()
    conn.close()
    return render_template('pajak_bukti_potong_list.html', rows=rows, pg=pg)

@app.route('/pajak/bukti-potong/<int:inv_id>/cetak')
@finance_required
def pajak_bukti_potong_cetak(inv_id):
    r = _require_pajak_module()
    if r: return r
    conn = db()
    inv = conn.execute("SELECT * FROM invoice WHERE id=?", (inv_id,)).fetchone()
    if not inv:
        conn.close(); flash('Invoice tidak ditemukan.', 'danger')
        return redirect(url_for('pajak_bukti_potong_list'))
    if not (inv['pph23_nominal'] or 0) > 0:
        conn.close(); flash('Invoice ini tidak punya PPh 23.', 'warning')
        return redirect(url_for('pajak_bukti_potong_list'))
    inv_s = _inv_settings(conn)
    items = conn.execute("SELECT * FROM invoice_item WHERE invoice_id=? ORDER BY id", (inv_id,)).fetchall()
    subtotal = sum(it['subtotal'] for it in items)
    dpp = _calc_invoice_dpp(subtotal, inv['diskon'], inv['ongkir'], inv['biaya_lain'])
    # Nomor bukti potong (BP-YYYYMM-XXX based on invoice id)
    no_bp = f"BP-{inv['tanggal'].replace('-','')[:6]}-{inv_id:03d}"
    conn.close()
    return render_template('pajak_bukti_potong_print.html', inv=inv, inv_s=inv_s,
                           items=items, dpp=dpp, no_bp=no_bp)

# ── Bukti Potong KELUAR (kita sebagai pemotong PPh 23 ke supplier) ──────────
@app.route('/pajak/bukti-potong-keluar')
@finance_required
def pajak_bukti_potong_keluar_list():
    r = _require_pajak_module()
    if r: return r
    conn = db()
    page = _page_arg()
    n_rows = conn.execute("""SELECT COUNT(*) FROM detail_jurnal dj
          JOIN jurnal j ON j.id = dj.jurnal_id JOIN akun a ON a.id = dj.akun_id
         WHERE a.kode = '2112' AND dj.kredit > 0""").fetchone()[0]
    pg = paginate_meta(page, n_rows)
    # Cari jurnal yang punya kredit ke akun 2112 (Hutang PPh Pasal 23)
    rows = conn.execute("""
        SELECT j.id, j.tanggal, j.keterangan, j.pihak,
               dj.kredit AS pph_nom
          FROM detail_jurnal dj
          JOIN jurnal j ON j.id = dj.jurnal_id
          JOIN akun a   ON a.id = dj.akun_id
         WHERE a.kode = '2112' AND dj.kredit > 0
         ORDER BY j.tanggal DESC, j.id DESC LIMIT ? OFFSET ?
    """, (pg['per_page'], pg['offset'])).fetchall()
    conn.close()
    return render_template('pajak_bukti_potong_keluar_list.html', rows=rows, pg=pg)

@app.route('/pajak/bukti-potong-keluar/<int:jurnal_id>/cetak')
@finance_required
def pajak_bukti_potong_keluar_cetak(jurnal_id):
    r = _require_pajak_module()
    if r: return r
    conn = db()
    j = conn.execute("SELECT * FROM jurnal WHERE id=?", (jurnal_id,)).fetchone()
    if not j:
        conn.close(); flash('Transaksi tidak ditemukan.', 'danger')
        return redirect(url_for('pajak_bukti_potong_keluar_list'))
    # Cek ada kredit ke 2112
    pph_row = conn.execute("""
        SELECT dj.kredit FROM detail_jurnal dj
          JOIN akun a ON a.id=dj.akun_id
         WHERE dj.jurnal_id=? AND a.kode='2112' AND dj.kredit>0
    """, (jurnal_id,)).fetchone()
    if not pph_row:
        conn.close(); flash('Transaksi ini tidak memiliki pemotongan PPh 23.', 'warning')
        return redirect(url_for('pajak_bukti_potong_keluar_list'))
    pph_nom = pph_row['kredit']
    # Cari DPP: debit ke akun non-pajak non-kas (aset/beban)
    detail_rows = conn.execute("""
        SELECT a.kode, a.nama, a.tipe, dj.debit, dj.kredit
          FROM detail_jurnal dj
          JOIN akun a ON a.id=dj.akun_id
         WHERE dj.jurnal_id=?
         ORDER BY dj.id
    """, (jurnal_id,)).fetchall()
    # Hitung DPP (total debit ke akun beban/aset non-PPN)
    dpp = sum(r['debit'] for r in detail_rows
              if r['debit'] > 0 and r['kode'] not in ('1180','1181','1182'))
    ppn_row = conn.execute("""
        SELECT dj.debit FROM detail_jurnal dj
          JOIN akun a ON a.id=dj.akun_id
         WHERE dj.jurnal_id=? AND a.kode='1180' AND dj.debit>0
    """, (jurnal_id,)).fetchone()
    ppn_nom = ppn_row['debit'] if ppn_row else 0
    dpp_bersih = dpp  # DPP sudah tanpa PPN Masukan
    pph_persen = round(pph_nom / dpp * 100, 2) if dpp > 0 else 2.0
    inv_s = _inv_settings(conn)
    no_bpk = f"BPK-{j['tanggal'].replace('-','')[:6]}-{jurnal_id:03d}"
    conn.close()
    return render_template('pajak_bukti_potong_keluar_print.html',
                           j=j, inv_s=inv_s, detail_rows=detail_rows,
                           dpp=dpp_bersih, ppn_nom=ppn_nom, pph_nom=pph_nom,
                           pph_persen=pph_persen, no_bpk=no_bpk)

# ── Bukti Bayar Pajak (internal BPN setelah Setor Pajak) ────────────────────
@app.route('/pajak/bukti-bayar')
@finance_required
def pajak_bukti_bayar_list():
    r = _require_pajak_module()
    if r: return r
    conn = db()
    page = _page_arg()
    n_rows = conn.execute("""SELECT COUNT(*) FROM detail_jurnal dj
          JOIN jurnal j ON j.id = dj.jurnal_id JOIN akun a ON a.id = dj.akun_id
         WHERE a.kode IN ('2111','2112','2110','2115','2116','2120') AND dj.debit > 0""").fetchone()[0]
    pg = paginate_meta(page, n_rows)
    # Cari jurnal setor pajak/gaji: punya debit ke akun kewajiban pajak, BPJS, atau gaji.
    rows = conn.execute("""
        SELECT j.id, j.tanggal, j.keterangan, j.pihak,
               a.kode AS akun_kode, a.nama AS akun_nama, dj.debit AS nominal
          FROM detail_jurnal dj
          JOIN jurnal j ON j.id = dj.jurnal_id
          JOIN akun a   ON a.id = dj.akun_id
         WHERE a.kode IN ('2111','2112','2110','2115','2116','2120') AND dj.debit > 0
         ORDER BY j.tanggal DESC, j.id DESC LIMIT ? OFFSET ?
    """, (pg['per_page'], pg['offset'])).fetchall()
    conn.close()
    return render_template('pajak_bukti_bayar_list.html', rows=rows, pg=pg)

@app.route('/pajak/bukti-bayar/<int:jurnal_id>/cetak')
@finance_required
def pajak_bukti_bayar_cetak(jurnal_id):
    r = _require_pajak_module()
    if r: return r
    conn = db()
    j = conn.execute("SELECT * FROM jurnal WHERE id=?", (jurnal_id,)).fetchone()
    if not j:
        conn.close(); flash('Transaksi tidak ditemukan.', 'danger')
        return redirect(url_for('pajak_bukti_bayar_list'))
    detail_rows = conn.execute("""
        SELECT a.kode, a.nama, a.tipe, dj.debit, dj.kredit
          FROM detail_jurnal dj
          JOIN akun a ON a.id=dj.akun_id
         WHERE dj.jurnal_id=?
         ORDER BY dj.id
    """, (jurnal_id,)).fetchall()
    # Akun pajak/gaji yang dibayar (debit ke hutang terkait)
    pajak_rows = [r for r in detail_rows if r['kode'] in ('2110','2111','2112','2115','2116','2120') and r['debit'] > 0]
    if not pajak_rows:
        conn.close(); flash('Jurnal ini bukan transaksi bayar pajak/gaji/BPJS.', 'warning')
        return redirect(url_for('pajak_bukti_bayar_list'))
    # Akun sumber pembayaran (kredit dari kas/bank)
    kas_rows = [r for r in detail_rows if r['kredit'] > 0 and r['kode'] not in ('2110','2111','2112','2115','2116','2120')]
    total_setor = sum(r['debit'] for r in pajak_rows)
    inv_s = _inv_settings(conn)
    no_bb = f"BB-{j['tanggal'].replace('-','')[:6]}-{jurnal_id:03d}"
    conn.close()
    return render_template('pajak_bukti_bayar_print.html',
                           j=j, inv_s=inv_s,
                           pajak_rows=pajak_rows, kas_rows=kas_rows,
                           total_setor=total_setor, no_bb=no_bb)

# ═══════════════════════════════════════════════════════════════════════════
#   MODUL PRO: GAJI & PPh 21
# ═══════════════════════════════════════════════════════════════════════════

# PTKP 2024 (per tahun)
PTKP = {
    'TK0': 54_000_000,
    'TK1': 58_500_000,
    'K0':  58_500_000,
    'K1':  63_000_000,
    'K2':  67_500_000,
    'K3':  72_000_000,
}

# Tarif progresif PPh 21 Pasal 17 (batas = lebar layer, bukan akumulasi)
TARIF_PPH21 = [
    (60_000_000,    0.05),
    (190_000_000,   0.15),
    (250_000_000,   0.25),
    (4_500_000_000, 0.30),
    (float('inf'),  0.35),
]

def hitung_pph21_bulanan(penghasilan_neto_bulan, status_pajak='TK0'):
    """Annualize -> kurangi PTKP -> tarif progresif -> bagi 12. Return integer rupiah."""
    ptkp = PTKP.get(status_pajak, 54_000_000)
    pkp = max(0, round(penghasilan_neto_bulan * 12 / 1_000) * 1_000 - ptkp)
    pph_tahunan = 0.0
    sisa = pkp
    for batas, tarif in TARIF_PPH21:
        if sisa <= 0:
            break
        kena = min(sisa, batas)
        pph_tahunan += kena * tarif
        sisa -= kena
    return round(pph_tahunan / 12)

# Ambang PPh 21 pegawai tidak tetap (harian/borongan) — metode klasik
PPH21_HARIAN_BEBAS   = 450_000      # upah/hari s.d. nilai ini tidak dipotong
PPH21_HARIAN_BULANAN = 4_500_000    # akumulasi sebulan; di atasnya pakai PTKP harian

def _pph21_tarif_progresif(pkp):
    """PPh 21 SETAHUN dari PKP (sudah dikurangi PTKP) pakai tarif Pasal 17."""
    sisa = max(0.0, float(pkp or 0)); pph = 0.0
    for batas, tarif in TARIF_PPH21:
        if sisa <= 0:
            break
        kena = min(sisa, batas); pph += kena * tarif; sisa -= kena
    return pph

def hitung_pph21_harian(total_upah, hari_kerja, status_pajak='TK0'):
    """PPh 21 pegawai tidak tetap (harian/borongan) — metode AMBANG HARIAN klasik:
      - upah/hari <= 450rb DAN total sebulan <= 4,5jt  -> 0
      - upah/hari >  450rb DAN total sebulan <= 4,5jt  -> 5% x (total - 450rb x hari)
      - total sebulan > 4,5jt                          -> 5% x (total - (PTKP_setahun/360) x hari)
    Catatan: utk akumulasi sangat besar (>10,2jt) idealnya disetahunkan (Pasal 17);
    di sini disederhanakan ke tarif 5% tier harian — memadai utk kasus UMKM.
    Return integer rupiah (>=0)."""
    total = max(0.0, float(total_upah or 0))
    hari  = float(hari_kerja or 0)
    if total <= 0:
        return 0
    if hari <= 0:
        hari = 1
    upah_harian = total / hari
    if total <= PPH21_HARIAN_BULANAN:
        if upah_harian <= PPH21_HARIAN_BEBAS:
            return 0
        return max(0, round(0.05 * (total - PPH21_HARIAN_BEBAS * hari)))
    ptkp_harian = PTKP.get(status_pajak, 54_000_000) / 360.0
    return max(0, round(0.05 * (total - ptkp_harian * hari)))

def hitung_bpjs(gaji_pokok, tunjangan=0, pakai_bpjs_kes=True, pakai_bpjs_tk=True):
    """
    Hitung iuran BPJS per bulan sesuai regulasi 2024.
    BPJS Kes: kar 1%, prs 4% (max dasar 12 jt).
    BPJS TK JHT: kar 2%, prs 3.7%; JP: kar 1%, prs 2% (max dasar JP 10,547,400).
    Return dict: bpjs_kes_kar, bpjs_kes_prs, bpjs_tk_kar, bpjs_tk_prs.
    """
    gaji = gaji_pokok + tunjangan
    dasar_kes = min(gaji, 12_000_000)
    bpjs_kes_kar = round(dasar_kes * 0.01) if pakai_bpjs_kes else 0
    bpjs_kes_prs = round(dasar_kes * 0.04) if pakai_bpjs_kes else 0
    if pakai_bpjs_tk:
        bpjs_jht_kar = round(gaji_pokok * 0.02)
        bpjs_jht_prs = round(gaji_pokok * 0.037)
        dasar_jp     = min(gaji_pokok, 10_547_400)
        bpjs_jp_kar  = round(dasar_jp * 0.01)
        bpjs_jp_prs  = round(dasar_jp * 0.02)
        bpjs_tk_kar  = bpjs_jht_kar + bpjs_jp_kar
        bpjs_tk_prs  = bpjs_jht_prs + bpjs_jp_prs
    else:
        bpjs_tk_kar = bpjs_tk_prs = 0
    return {
        'bpjs_kes_kar': bpjs_kes_kar,
        'bpjs_kes_prs': bpjs_kes_prs,
        'bpjs_tk_kar':  bpjs_tk_kar,
        'bpjs_tk_prs':  bpjs_tk_prs,
    }

def _kalkulasi_payroll(k):
    """Hitung komponen gaji dari row karyawan. Return dict siap pakai."""
    bruto = float(k['gaji_pokok']) + float(k['tunjangan_tetap'])
    b = hitung_bpjs(float(k['gaji_pokok']), float(k['tunjangan_tetap']),
                    bool(k['bpjs_kesehatan']), bool(k['bpjs_ketenagakerjaan']))
    biaya_jabatan = min(bruto * 0.05, 500_000)
    neto_bulan    = bruto - b['bpjs_kes_kar'] - b['bpjs_tk_kar'] - biaya_jabatan
    pph21         = hitung_pph21_bulanan(neto_bulan, k['status_pajak'])
    gaji_bersih   = bruto - b['bpjs_kes_kar'] - b['bpjs_tk_kar'] - pph21
    bpjs_total_prs = b['bpjs_kes_prs'] + b['bpjs_tk_prs']
    bpjs_total_kar = b['bpjs_kes_kar'] + b['bpjs_tk_kar']
    return {
        'bruto': bruto,
        'pph21': pph21,
        'gaji_bersih': gaji_bersih,
        'bpjs_total_prs': bpjs_total_prs,
        'bpjs_total_kar': bpjs_total_kar,
        **b,
    }

def kalkulasi_payroll(k, jenis_upah=None, bonus=0, hari_kerja=0, qty_output=0, tarif=None,
                      uang_makan=0, uang_transport=0, lain_lain=0):
    """Hitung payroll satu periode utk SEMUA jenis upah + komponen variabel.
    Komponen variabel (per periode, dari form Proses Gaji):
      - uang_makan, uang_transport : tunjangan tidak tetap TERATUR
          -> digabung ke penghasilan rutin bulanan, ikut biaya jabatan & anualisasi.
      - bonus, lain_lain           : penghasilan TIDAK TERATUR (THR/bonus/dll)
          -> PPh dihitung sbg selisih progresif setahun.
    BULANAN : gaji_pokok+tunjangan dari master + tunjangan variabel, BPJS aktif.
    HARIAN  : tarif_harian x hari_kerja, TANPA BPJS, PPh21 ambang harian.
    BORONGAN: tarif(satuan) x qty_output, TANPA BPJS, PPh21 ambang harian (pakai hari_kerja).
    Return dict komponen siap simpan + jurnal."""
    cols = set(k.keys())
    jenis = (jenis_upah or (k['jenis_upah'] if 'jenis_upah' in cols else None) or 'BULANAN').upper()
    bonus      = max(0.0, float(bonus or 0))
    uang_makan = max(0.0, float(uang_makan or 0))
    uang_trans = max(0.0, float(uang_transport or 0))
    lain_lain  = max(0.0, float(lain_lain or 0))
    teratur_var   = uang_makan + uang_trans       # tunjangan tidak tetap, sifat teratur
    tidak_teratur = bonus + lain_lain             # bonus/THR/lain-lain, sifat tidak teratur
    status_pajak = k['status_pajak'] or 'TK0'
    hari = float(hari_kerja or 0)
    qty  = float(qty_output or 0)

    if jenis == 'BULANAN':
        gaji_pokok = float(k['gaji_pokok'] or 0)
        tunjangan  = float(k['tunjangan_tetap'] or 0)
        upah = gaji_pokok + tunjangan                 # bagian tetap (utk slip & BPJS)
        penghasilan_teratur = upah + teratur_var      # dasar PPh teratur
        # BPJS dihitung dari bagian tetap saja (tunjangan tidak tetap bukan dasar BPJS)
        b = hitung_bpjs(gaji_pokok, tunjangan,
                        bool(k['bpjs_kesehatan']), bool(k['bpjs_ketenagakerjaan']))
        neto_reg = (penghasilan_teratur - b['bpjs_kes_kar'] - b['bpjs_tk_kar']
                    - min(penghasilan_teratur * 0.05, 500_000))
        pph_reg  = hitung_pph21_bulanan(neto_reg, status_pajak)
        if tidak_teratur > 0:
            ptkp = PTKP.get(status_pajak, 54_000_000)
            pkp_reg  = max(0, round(neto_reg * 12 / 1_000) * 1_000 - ptkp)
            pkp_with = max(0, round((neto_reg * 12 + tidak_teratur) / 1_000) * 1_000 - ptkp)
            pph_tt = max(0, round(_pph21_tarif_progresif(pkp_with) - _pph21_tarif_progresif(pkp_reg)))
        else:
            pph_tt = 0
        pph21 = pph_reg + pph_tt
        bkk, btk = b['bpjs_kes_kar'], b['bpjs_tk_kar']
        bkp, btp = b['bpjs_kes_prs'], b['bpjs_tk_prs']
        bruto = penghasilan_teratur + tidak_teratur
        gaji_bersih = bruto - bkk - btk - pph21
    else:
        t = float(tarif if tarif is not None
                  else ((k['tarif_harian'] if 'tarif_harian' in cols else 0) or 0))
        upah = t * (hari if jenis == 'HARIAN' else qty)
        # Pegawai tidak tetap: semua komponen variabel masuk bruto & ambang harian
        bruto = upah + teratur_var + tidak_teratur
        pph21 = hitung_pph21_harian(bruto, hari, status_pajak)
        bkk = btk = bkp = btp = 0
        gaji_bersih = bruto - pph21

    return {
        'jenis_upah': jenis, 'upah': round(upah, 2), 'bonus': round(bonus, 2),
        'uang_makan': round(uang_makan, 2), 'uang_transport': round(uang_trans, 2),
        'lain_lain': round(lain_lain, 2),
        'bruto': round(bruto, 2), 'pph21': pph21, 'gaji_bersih': round(gaji_bersih, 2),
        'bpjs_kes_kar': bkk, 'bpjs_tk_kar': btk, 'bpjs_kes_prs': bkp, 'bpjs_tk_prs': btp,
        'bpjs_total_kar': bkk + btk, 'bpjs_total_prs': bkp + btp,
        'hari_kerja': hari, 'qty_output': qty,
        'tarif': float(tarif) if tarif is not None else 0.0,
    }

# ── Karyawan CRUD ─────────────────────────────────────────────────────────

@app.route('/gaji/karyawan', methods=['GET','POST'])
@finance_required
def gaji_karyawan():
    r = _require_pajak_module()
    if r: return r
    conn = db()
    if request.method == 'POST':
        action = request.form.get('action','')
        if action == 'tambah':
            nama   = request.form.get('nama','').strip()
            jabatan= request.form.get('jabatan','').strip()
            st_kerja = request.form.get('status_kerja','TETAP')
            st_pajak = request.form.get('status_pajak','TK0')
            no_kar   = request.form.get('no_karyawan','').strip()
            npwp     = request.form.get('npwp','').strip()
            bpjs_kes = 1 if request.form.get('bpjs_kesehatan') else 0
            bpjs_tk  = 1 if request.form.get('bpjs_ketenagakerjaan') else 0
            jenis_upah = (request.form.get('jenis_upah','BULANAN') or 'BULANAN').upper()
            if jenis_upah not in ('BULANAN','HARIAN','BORONGAN'):
                jenis_upah = 'BULANAN'
            try:
                tanggungan = max(0, min(20, int(request.form.get('jumlah_tanggungan',0) or 0)))
            except (TypeError, ValueError):
                tanggungan = 0
            try:
                gaji_pokok  = parse_nonnegative_rp(request.form.get('gaji_pokok',0),  'Gaji Pokok')
                tunjangan   = parse_nonnegative_rp(request.form.get('tunjangan',0),    'Tunjangan')
                tarif_harian= parse_nonnegative_rp(request.form.get('tarif_harian',0), 'Tarif Harian/Satuan')
            except ValueError as ex:
                flash(str(ex), 'danger')
                conn.close(); return redirect(url_for('gaji_karyawan'))
            if not nama:
                flash('Nama karyawan wajib diisi.', 'danger')
            else:
                conn.execute("""
                    INSERT INTO karyawan(no_karyawan,nama,jabatan,status_kerja,status_pajak,
                                        gaji_pokok,tunjangan_tetap,bpjs_kesehatan,
                                        bpjs_ketenagakerjaan,npwp,
                                        jenis_upah,tarif_harian,jumlah_tanggungan)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (no_kar, nama, jabatan, st_kerja, st_pajak,
                      gaji_pokok, tunjangan, bpjs_kes, bpjs_tk, npwp,
                      jenis_upah, tarif_harian, tanggungan))
                conn.commit()
                flash(f'Karyawan {nama} berhasil ditambahkan.', 'success')
        elif action == 'edit':
            kid = int(request.form.get('karyawan_id',0) or 0)
            nama   = request.form.get('nama','').strip()
            jabatan= request.form.get('jabatan','').strip()
            st_kerja = request.form.get('status_kerja','TETAP')
            st_pajak = request.form.get('status_pajak','TK0')
            no_kar   = request.form.get('no_karyawan','').strip()
            npwp     = request.form.get('npwp','').strip()
            bpjs_kes = 1 if request.form.get('bpjs_kesehatan') else 0
            bpjs_tk  = 1 if request.form.get('bpjs_ketenagakerjaan') else 0
            aktif    = 1 if request.form.get('aktif') else 0
            jenis_upah = (request.form.get('jenis_upah','BULANAN') or 'BULANAN').upper()
            if jenis_upah not in ('BULANAN','HARIAN','BORONGAN'):
                jenis_upah = 'BULANAN'
            try:
                tanggungan = max(0, min(20, int(request.form.get('jumlah_tanggungan',0) or 0)))
            except (TypeError, ValueError):
                tanggungan = 0
            try:
                gaji_pokok = parse_nonnegative_rp(request.form.get('gaji_pokok',0), 'Gaji Pokok')
                tunjangan  = parse_nonnegative_rp(request.form.get('tunjangan',0),  'Tunjangan')
                tarif_harian= parse_nonnegative_rp(request.form.get('tarif_harian',0),'Tarif Harian/Satuan')
            except ValueError as ex:
                flash(str(ex), 'danger')
                conn.close(); return redirect(url_for('gaji_karyawan'))
            if nama and kid:
                conn.execute("""
                    UPDATE karyawan SET no_karyawan=?,nama=?,jabatan=?,status_kerja=?,
                        status_pajak=?,gaji_pokok=?,tunjangan_tetap=?,bpjs_kesehatan=?,
                        bpjs_ketenagakerjaan=?,npwp=?,aktif=?,
                        jenis_upah=?,tarif_harian=?,jumlah_tanggungan=?
                    WHERE id=?
                """, (no_kar, nama, jabatan, st_kerja, st_pajak, gaji_pokok, tunjangan,
                      bpjs_kes, bpjs_tk, npwp, aktif,
                      jenis_upah, tarif_harian, tanggungan, kid))
                conn.commit()
                flash('Data karyawan diperbarui.', 'success')
        elif action == 'hapus':
            kid = int(request.form.get('karyawan_id',0) or 0)
            n_payroll = conn.execute("SELECT COUNT(*) FROM payroll WHERE karyawan_id=?", (kid,)).fetchone()[0]
            if n_payroll > 0:
                flash('Tidak bisa menghapus karyawan yang sudah punya riwayat gaji. Nonaktifkan saja.', 'warning')
            elif kid:
                conn.execute("DELETE FROM karyawan WHERE id=?", (kid,))
                conn.commit()
                flash('Karyawan dihapus.', 'warning')
        conn.close()
        return redirect(url_for('gaji_karyawan'))

    karyawan_list = [dict(r) for r in conn.execute("""
        SELECT k.*,
               (SELECT COUNT(*) FROM payroll p WHERE p.karyawan_id=k.id) AS total_payroll
        FROM karyawan k ORDER BY k.aktif DESC, k.nama
    """).fetchall()]
    conn.close()
    return render_template('gaji_karyawan.html', karyawan_list=karyawan_list,
                           PTKP=PTKP)

# ── Preview kalkulasi (AJAX) ───────────────────────────────────────────────

@app.route('/gaji/preview-kalkulasi')
@finance_required
def gaji_preview_kalkulasi():
    r = _require_pajak_module()
    if r: return jsonify({'ok': False, 'msg': 'Modul belum aktif.'})
    conn = db()
    kid = int(request.args.get('karyawan_id', 0) or 0)
    k   = conn.execute("SELECT * FROM karyawan WHERE id=?", (kid,)).fetchone()
    conn.close()
    if not k:
        return jsonify({'ok': False, 'msg': 'Karyawan tidak ditemukan.'})
    def _num(v):
        try:
            return float(str(v or '').replace('.', '').replace(',', '.') or 0)
        except (TypeError, ValueError):
            return 0.0
    _t = (request.args.get('tarif') or '').strip()
    d = kalkulasi_payroll(
        k,
        jenis_upah=request.args.get('jenis_upah') or None,
        bonus=_num(request.args.get('bonus')),
        hari_kerja=_num(request.args.get('hari_kerja')),
        qty_output=_num(request.args.get('qty_output')),
        tarif=_num(_t) if _t else None,
        uang_makan=_num(request.args.get('uang_makan')),
        uang_transport=_num(request.args.get('uang_transport')),
        lain_lain=_num(request.args.get('lain_lain')),
    )
    return jsonify({'ok': True, **d, 'nama': k['nama']})

# ── Proses Payroll ─────────────────────────────────────────────────────────

@app.route('/gaji/proses', methods=['GET','POST'])
@finance_required
def gaji_proses():
    r = _require_pajak_module()
    if r: return r
    conn = db()
    today = date.today()
    karyawan_list = conn.execute("SELECT * FROM karyawan WHERE aktif=1 ORDER BY nama").fetchall()

    if request.method == 'POST':
        kid    = int(request.form.get('karyawan_id', 0) or 0)
        bulan  = int(request.form.get('bulan', today.month))
        tahun  = int(request.form.get('tahun', today.year))
        catatan = request.form.get('catatan','').strip()

        k = conn.execute("SELECT * FROM karyawan WHERE id=?", (kid,)).fetchone()
        if not k:
            flash('Karyawan tidak ditemukan.', 'danger')
            conn.close(); return redirect(url_for('gaji_proses'))

        dup = conn.execute(
            "SELECT id FROM payroll WHERE karyawan_id=? AND periode_bulan=? AND periode_tahun=?",
            (kid, bulan, tahun)
        ).fetchone()
        if dup:
            flash(f'Gaji {k["nama"]} bulan {bulan}/{tahun} sudah pernah diproses.', 'warning')
            conn.close(); return redirect(url_for('gaji_proses'))

        cols_k = set(k.keys())
        jenis_upah = (request.form.get('jenis_upah')
                      or (k['jenis_upah'] if 'jenis_upah' in cols_k else 'BULANAN')
                      or 'BULANAN').upper()
        try:
            bonus      = parse_nonnegative_rp(request.form.get('bonus', 0) or 0, 'Bonus')
            uang_makan = parse_nonnegative_rp(request.form.get('uang_makan', 0) or 0, 'Uang Makan')
            uang_trans = parse_nonnegative_rp(request.form.get('uang_transport', 0) or 0, 'Uang Transport')
            lain_lain  = parse_nonnegative_rp(request.form.get('lain_lain', 0) or 0, 'Lain-lain')
            # hari_kerja dulu bare float(): field ini data-nf="1" (namanya tak match
            # regex qty di base.html), jadi wire-nya sekarang koma-desimal ("15,5")
            # sejak perbaikan initForms(). float() Python menolak koma -> WAJIB lewat
            # parse_nonnegative_rp() spt field lain di sini, bukan cuma field uang.
            hari_kerja = parse_nonnegative_rp(request.form.get('hari_kerja', 0) or 0, 'Jumlah hari kerja')
            qty_output = float(request.form.get('qty_output') or 0)
            _tarif_in  = (request.form.get('tarif') or '').strip()
            tarif      = parse_nonnegative_rp(_tarif_in, 'Tarif') if _tarif_in else None
        except ValueError as ex:
            flash(str(ex), 'danger')
            conn.close(); return redirect(url_for('gaji_proses'))

        calc = kalkulasi_payroll(k, jenis_upah=jenis_upah, bonus=bonus,
                                 hari_kerja=hari_kerja, qty_output=qty_output, tarif=tarif,
                                 uang_makan=uang_makan, uang_transport=uang_trans,
                                 lain_lain=lain_lain)
        jenis_upah     = calc['jenis_upah']
        bruto          = calc['bruto']
        pph21          = calc['pph21']
        gaji_bersih    = calc['gaji_bersih']
        bpjs_total_prs = calc['bpjs_total_prs']
        bpjs_total_kar = calc['bpjs_total_kar']
        bpjs_total     = bpjs_total_kar + bpjs_total_prs

        if bruto <= 0:
            flash('Nilai upah/gaji nol. Periksa tarif, hari kerja, atau qty output.', 'warning')
            conn.close(); return redirect(url_for('gaji_proses'))

        label_gaji = 'Gaji' if jenis_upah == 'BULANAN' else 'Upah'
        nm_bln = calendar.month_name[bulan]
        ket    = f"{label_gaji} {k['nama']} - {nm_bln} {tahun}"
        tgl_gajian = date(tahun, bulan,
                          min(25, calendar.monthrange(tahun, bulan)[1])).strftime('%Y-%m-%d')

        # Jurnal:
        # Dr 6100 Beban Gaji/Upah     = bruto
        # Dr 6101 Beban BPJS Persh.   = bpjs_total_prs  (hanya BULANAN)
        # Cr 2115 Hutang PPh 21       = pph21            (jika ada)
        # Cr 2116 Hutang BPJS         = bpjs_total       (hanya BULANAN)
        # Cr 2120 Hutang Gaji         = gaji_bersih
        entries = [('6100', bruto, 0)]
        if bpjs_total_prs > 0:
            entries.append(('6101', bpjs_total_prs, 0))
        if pph21 > 0:
            entries.append(('2115', 0, pph21))
        if bpjs_total > 0:
            entries.append(('2116', 0, bpjs_total))
        entries.append(('2120', 0, gaji_bersih))

        try:
            j_id = insert_jurnal(conn, tgl_gajian, ket, 'GAJI', 'JURNAL', entries, pihak=k['nama'])
        except ValueError as ex:
            flash(str(ex), 'danger')
            conn.close(); return redirect(url_for('gaji_proses'))

        conn.execute("""
            INSERT INTO payroll(karyawan_id,periode_bulan,periode_tahun,
                gaji_pokok,tunjangan,penghasilan_bruto,
                bpjs_kes_kar,bpjs_tk_kar,bpjs_kes_prs,bpjs_tk_prs,
                pph21,gaji_bersih,jurnal_id,catatan,
                jenis_upah,bonus,hari_kerja,qty_output,
                uang_makan,uang_transport,lain_lain)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (kid, bulan, tahun,
              (float(k['gaji_pokok']) if jenis_upah == 'BULANAN' else calc['upah']),
              (float(k['tunjangan_tetap']) if jenis_upah == 'BULANAN' else 0),
              bruto, calc['bpjs_kes_kar'], calc['bpjs_tk_kar'],
              calc['bpjs_kes_prs'], calc['bpjs_tk_prs'],
              pph21, gaji_bersih, j_id, catatan,
              jenis_upah, calc['bonus'], calc['hari_kerja'], calc['qty_output'],
              calc['uang_makan'], calc['uang_transport'], calc['lain_lain']))
        conn.commit()
        add_log(conn, 'Proses gaji', f"{k['nama']} | {nm_bln} {tahun} | Rp {gaji_bersih:,.0f}", 'GAJI')
        conn.commit()
        flash(f'Gaji {k["nama"]} bulan {nm_bln} {tahun} diproses. Jurnal #{j_id} dibuat.', 'success')
        conn.close()
        return redirect(url_for('gaji_riwayat'))

    conn.close()
    return render_template('gaji_proses.html',
                           karyawan_list=karyawan_list,
                           today=today,
                           bulan_ini=today.month,
                           tahun_ini=today.year,
                           PTKP=PTKP)

# ── Riwayat Payroll ────────────────────────────────────────────────────────

@app.route('/gaji/riwayat')
@finance_required
def gaji_riwayat():
    r = _require_pajak_module()
    if r: return r
    conn = db()
    rows = conn.execute("""
        SELECT p.*, k.nama AS nama_karyawan, k.jabatan
        FROM payroll p
        JOIN karyawan k ON k.id=p.karyawan_id
        ORDER BY p.periode_tahun DESC, p.periode_bulan DESC, k.nama
    """).fetchall()
    # Group by (tahun, bulan)
    from collections import OrderedDict
    periodes = OrderedDict()
    for row in rows:
        key = (row['periode_tahun'], row['periode_bulan'])
        if key not in periodes:
            periodes[key] = {
                'tahun':  row['periode_tahun'],
                'bulan':  row['periode_bulan'],
                'nm_bln': calendar.month_name[row['periode_bulan']],
                'rows':   [],
                'total_bruto':   0,
                'total_pph21':   0,
                'total_bersih':  0,
            }
        periodes[key]['rows'].append(row)
        periodes[key]['total_bruto']  += row['penghasilan_bruto']
        periodes[key]['total_pph21']  += row['pph21']
        periodes[key]['total_bersih'] += row['gaji_bersih']
    conn.close()
    return render_template('gaji_riwayat.html', periodes=list(periodes.values()))

# ── Slip Gaji (print) ──────────────────────────────────────────────────────

@app.route('/gaji/<int:pid>/slip')
@finance_required
def gaji_slip(pid):
    r = _require_pajak_module()
    if r: return r
    conn = db()
    p = conn.execute("""
        SELECT p.*, k.nama AS nama_karyawan, k.jabatan, k.no_karyawan,
               k.npwp, k.status_pajak, k.status_kerja
        FROM payroll p
        JOIN karyawan k ON k.id=p.karyawan_id
        WHERE p.id=?
    """, (pid,)).fetchone()
    if not p:
        conn.close()
        flash('Data payroll tidak ditemukan.', 'danger')
        return redirect(url_for('gaji_riwayat'))
    inv_s = _inv_settings(conn)
    conn.close()
    nm_bln = calendar.month_name[p['periode_bulan']]
    return render_template('gaji_slip.html', p=p, inv_s=inv_s,
                           nm_bln=nm_bln, PTKP=PTKP)

# ── Hapus Payroll Entry ────────────────────────────────────────────────────

@app.route('/gaji/<int:pid>/hapus', methods=['POST'])
@finance_required
def gaji_hapus(pid):
    r = _require_pajak_module()
    if r: return r
    conn = db()
    p = conn.execute("SELECT * FROM payroll WHERE id=?", (pid,)).fetchone()
    if p:
        if p['jurnal_id']:
            kewajiban_rows = conn.execute("""
                SELECT a.kode, a.nama, COALESCE(SUM(dj.kredit),0) AS nominal
                  FROM detail_jurnal dj
                  JOIN akun a ON a.id=dj.akun_id
                 WHERE dj.jurnal_id=? AND a.kode IN ('2115','2116','2120') AND dj.kredit>0
                 GROUP BY a.kode, a.nama
            """, (p['jurnal_id'],)).fetchall()
            blocked = []
            for row in kewajiban_rows:
                saldo = _akun_saldo(conn, row['kode'])
                if saldo + 0.01 < float(row['nominal'] or 0):
                    blocked.append(row['nama'])
            if blocked:
                conn.close()
                flash('Payroll ini sudah terkait pembayaran kewajiban. Hapus/batalkan pembayaran '
                      + ', '.join(blocked) + ' lebih dulu.', 'warning')
                return redirect(url_for('gaji_riwayat'))
            _reverse_tx(conn, p['jurnal_id'])
            conn.execute("DELETE FROM jurnal WHERE id=?", (p['jurnal_id'],))
        conn.execute("DELETE FROM payroll WHERE id=?", (pid,))
        conn.commit()
        flash('Data gaji dan jurnalnya berhasil dihapus.', 'warning')
    conn.close()
    return redirect(url_for('gaji_riwayat'))


# ═══════════════════════════════════════════════════════════════════════════
#   MODUL LANJUTAN: ANGGARAN & TARGET
# ═══════════════════════════════════════════════════════════════════════════

def _require_anggaran_module():
    conn = db()
    aktif = is_module_active(conn, 'ANGGARAN')
    conn.close()
    if not aktif:
        flash('Modul Lanjutan "Anggaran & Target" belum aktif. Aktivasi dulu di Pengaturan > Modul Lanjutan.', 'warning')
        return redirect(url_for('pengaturan_modul'))
    return None

def _realisasi_akun(conn, akun_kode, from_d, to_d):
    """Realisasi beban: debit - kredit untuk akun BEBAN periode tertentu."""
    row = conn.execute("""
        SELECT COALESCE(SUM(dj.debit) - SUM(dj.kredit), 0) AS net
        FROM detail_jurnal dj
        JOIN akun a ON a.id=dj.akun_id
        JOIN jurnal j ON j.id=dj.jurnal_id
        WHERE a.kode=? AND j.tanggal BETWEEN ? AND ?
    """, (akun_kode, from_d, to_d)).fetchone()
    return float(row['net'] or 0)

@app.route('/anggaran', methods=['GET','POST'])
@finance_required
def anggaran_home():
    r = _require_anggaran_module()
    if r: return r
    conn = db()
    today = date.today()
    try:
        bulan = int(request.args.get('bulan', today.month))
        tahun = int(request.args.get('tahun', today.year))
    except (ValueError, TypeError):
        bulan, tahun = today.month, today.year
    bulan = max(1, min(12, bulan))

    if request.method == 'POST':
        action = request.form.get('action','')
        b = int(request.form.get('bulan', bulan))
        t = int(request.form.get('tahun', tahun))
        if action == 'set_target':
            try:
                nominal = parse_nonnegative_rp(request.form.get('nominal',0), 'Target Pendapatan')
            except ValueError as ex:
                flash(str(ex), 'danger')
                conn.close()
                return redirect(url_for('anggaran_home', bulan=b, tahun=t))
            catatan = request.form.get('catatan','').strip()
            conn.execute("""
                INSERT INTO target_pendapatan(periode_bulan,periode_tahun,nominal,catatan)
                VALUES(?,?,?,?)
                ON CONFLICT(periode_bulan,periode_tahun)
                DO UPDATE SET nominal=excluded.nominal, catatan=excluded.catatan
            """, (b, t, nominal, catatan))
            conn.commit()
            flash(f'Target pendapatan {calendar.month_name[b]} {t} disimpan.', 'success')
        elif action == 'set_anggaran':
            akun_kode = request.form.get('akun_kode','').strip()
            try:
                nominal = parse_nonnegative_rp(request.form.get('nominal',0), 'Nominal Anggaran')
            except ValueError as ex:
                flash(str(ex), 'danger')
                conn.close()
                return redirect(url_for('anggaran_home', bulan=b, tahun=t))
            if akun_kode and nominal > 0:
                conn.execute("""
                    INSERT INTO anggaran(periode_bulan,periode_tahun,akun_kode,nominal)
                    VALUES(?,?,?,?)
                    ON CONFLICT(periode_bulan,periode_tahun,akun_kode)
                    DO UPDATE SET nominal=excluded.nominal
                """, (b, t, akun_kode, nominal))
                conn.commit()
                flash('Pos anggaran disimpan.', 'success')
            else:
                flash('Pilih akun dan masukkan nominal anggaran.', 'warning')
        elif action == 'hapus_anggaran':
            aid = int(request.form.get('anggaran_id',0) or 0)
            if aid:
                conn.execute("DELETE FROM anggaran WHERE id=?", (aid,))
                conn.commit()
                flash('Pos anggaran dihapus.', 'warning')
        elif action == 'salin_bulan_lalu':
            # Salin anggaran dari bulan sebelumnya
            prev_b = bulan - 1 if bulan > 1 else 12
            prev_t = tahun if bulan > 1 else tahun - 1
            prev_rows = conn.execute(
                "SELECT akun_kode, nominal FROM anggaran WHERE periode_bulan=? AND periode_tahun=?",
                (prev_b, prev_t)
            ).fetchall()
            count = 0
            for pr in prev_rows:
                conn.execute("""
                    INSERT OR IGNORE INTO anggaran(periode_bulan,periode_tahun,akun_kode,nominal)
                    VALUES(?,?,?,?)
                """, (bulan, tahun, pr['akun_kode'], pr['nominal']))
                count += 1
            conn.commit()
            flash(f'{count} pos anggaran disalin dari {calendar.month_name[prev_b]} {prev_t}.', 'success')
        conn.close()
        return redirect(url_for('anggaran_home', bulan=bulan, tahun=tahun))

    # GET — build data
    last_day = calendar.monthrange(tahun, bulan)[1]
    from_d   = f"{tahun}-{bulan:02d}-01"
    to_d     = f"{tahun}-{bulan:02d}-{last_day:02d}"

    # Target & aktual pendapatan
    target_row = conn.execute(
        "SELECT * FROM target_pendapatan WHERE periode_bulan=? AND periode_tahun=?",
        (bulan, tahun)
    ).fetchone()
    target_nominal = float(target_row['nominal']) if target_row else 0.0
    aktual_pendapatan = float(conn.execute("""
        SELECT COALESCE(SUM(dj.kredit - dj.debit), 0)
        FROM detail_jurnal dj
        JOIN akun a ON a.id=dj.akun_id
        JOIN jurnal j ON j.id=dj.jurnal_id
        WHERE a.tipe='PENDAPATAN' AND j.tanggal BETWEEN ? AND ?
    """, (from_d, to_d)).fetchone()[0] or 0)

    # Anggaran beban
    ang_rows = conn.execute("""
        SELECT an.id, an.akun_kode, an.nominal, ak.nama AS akun_nama, ak.subtipe
        FROM anggaran an
        JOIN akun ak ON ak.kode=an.akun_kode
        WHERE an.periode_bulan=? AND an.periode_tahun=?
        ORDER BY ak.subtipe, ak.kode
    """, (bulan, tahun)).fetchall()

    anggaran_data = []
    total_anggaran = total_realisasi = 0.0
    for row in ang_rows:
        real = _realisasi_akun(conn, row['akun_kode'], from_d, to_d)
        nom  = float(row['nominal'])
        sel  = nom - real
        pct  = min(200, round(real / nom * 100)) if nom > 0 else 0
        anggaran_data.append({
            'id': row['id'], 'akun_kode': row['akun_kode'],
            'akun_nama': row['akun_nama'], 'subtipe': row['subtipe'],
            'nominal': nom, 'realisasi': real,
            'selisih': sel, 'pct': pct, 'over': real > nom,
        })
        total_anggaran += nom
        total_realisasi += real

    # Daftar akun BEBAN yang belum dianggarkan (untuk picker)
    budgeted = {d['akun_kode'] for d in anggaran_data}
    akun_beban_avail = conn.execute(
        "SELECT kode, nama, subtipe FROM akun WHERE tipe='BEBAN' ORDER BY kode"
    ).fetchall()
    akun_beban_avail = [a for a in akun_beban_avail if a['kode'] not in budgeted]

    nm_bln = calendar.month_name[bulan]
    conn.close()
    return render_template('anggaran.html',
        bulan=bulan, tahun=tahun, nm_bln=nm_bln, today=today,
        target_nominal=target_nominal, target_row=target_row,
        aktual_pendapatan=aktual_pendapatan,
        anggaran_data=anggaran_data,
        total_anggaran=total_anggaran, total_realisasi=total_realisasi,
        akun_beban_avail=akun_beban_avail,
    )


def _save_invoice_items(conn, inv_id, form):
    conn.execute("DELETE FROM invoice_item WHERE invoice_id=?", (inv_id,))
    desks = form.getlist('deskripsi[]')
    qtys  = form.getlist('qty[]')
    sats  = form.getlist('satuan[]')
    hargs = form.getlist('harga_satuan[]')
    diss  = form.getlist('diskon_item[]')
    saved = 0
    for i, desk in enumerate(desks):
        if not desk.strip(): continue
        try:
            qty = parse_qty(qtys[i] if i < len(qtys) else 1, 1)
            harg = parse_rp(hargs[i] if i < len(hargs) else 0)
            dis_i = parse_rp(diss[i] if i < len(diss) else 0)
        except (TypeError, ValueError):
            raise ValueError(f'Angka item invoice baris {i + 1} tidak valid.')
        if qty <= 0:
            raise ValueError(f'Qty item invoice baris {i + 1} harus lebih dari 0.')
        sat = (sats[i] if i < len(sats) else 'pcs').strip() or 'pcs'
        subtotal = max(0, qty * harg - dis_i)
        conn.execute(
            "INSERT INTO invoice_item(invoice_id,deskripsi,qty,satuan,harga_satuan,diskon_item,subtotal) VALUES(?,?,?,?,?,?,?)",
            (inv_id, desk.strip(), qty, sat, harg, dis_i, subtotal))
        saved += 1
    if not saved:
        raise ValueError('Minimal 1 item invoice harus diisi.')

def _invoice_subtotal(conn, inv_id):
    """Subtotal = jumlah subtotal semua item invoice (sebelum diskon/ongkir/biaya/pajak)."""
    row = conn.execute("SELECT COALESCE(SUM(subtotal),0) FROM invoice_item WHERE invoice_id=?", (inv_id,)).fetchone()
    return float(row[0] or 0)

def _calc_invoice_pajak(subtotal, diskon, ongkir, biaya_lain, pajak_persen):
    """DPP pajak dihitung dari (subtotal - diskon + ongkir + biaya_lain). Pajak persen=0 → 0."""
    dpp = max(0.0, float(subtotal) - float(diskon or 0) + float(ongkir or 0) + float(biaya_lain or 0))
    return round(dpp * (float(pajak_persen or 0) / 100.0), 2)

def _calc_invoice_dpp(subtotal, diskon, ongkir, biaya_lain):
    return max(0.0, float(subtotal) - float(diskon or 0) + float(ongkir or 0) + float(biaya_lain or 0))

def _recalc_invoice_pajak(conn, inv_id):
    """Hitung ulang pajak_nominal + pph23_nominal + pph22_nominal & simpan. Aman walau persen=0 (hasil=0)."""
    inv = conn.execute("SELECT diskon,ongkir,biaya_lain,pajak_persen,pph23_persen,pph22_persen FROM invoice WHERE id=?", (inv_id,)).fetchone()
    if not inv: return
    sub = _invoice_subtotal(conn, inv_id)
    dpp = _calc_invoice_dpp(sub, inv['diskon'], inv['ongkir'], inv['biaya_lain'])
    pjk = round(dpp * (float(inv['pajak_persen'] or 0) / 100.0), 2)
    pph = round(dpp * (float(inv['pph23_persen'] or 0) / 100.0), 2)
    pph22 = round(dpp * (float(inv['pph22_persen'] or 0) / 100.0), 2)
    conn.execute("UPDATE invoice SET pajak_nominal=?, pph23_nominal=?, pph22_nominal=? WHERE id=?",
                 (pjk, pph, pph22, inv_id))

def _post_invoice_to_jurnal(conn, inv_id):
    """
    Catat jurnal otomatis untuk invoice PKP (manual invoice).
    Format sesuai standar PKP yang diminta customer:
        Dr Piutang Usaha       = DPP + PPN - PPh 23
        Dr PPh 23 Dibayar Dimuka = PPh 23 nominal (kalau dipotong)
            Cr Pendapatan Penjualan = DPP
            Cr Hutang PPN           = PPN nominal (kalau kena PPN)
    Hanya untuk invoice manual (jurnal_id IS NULL). Modul PAJAK_OTOMATIS harus aktif.
    """
    inv = conn.execute("SELECT * FROM invoice WHERE id=?", (inv_id,)).fetchone()
    if not inv:
        raise ValueError('Invoice tidak ditemukan.')
    if inv['jurnal_id']:
        raise ValueError('Invoice ini sudah tertaut ke jurnal sumber.')
    if not is_module_active(conn, 'PAJAK_OTOMATIS'):
        raise ValueError('Modul PPN & Faktur Otomatis belum aktif.')
    _recalc_invoice_pajak(conn, inv_id)
    inv = conn.execute("SELECT * FROM invoice WHERE id=?", (inv_id,)).fetchone()
    sub = _invoice_subtotal(conn, inv_id)
    dpp = _calc_invoice_dpp(sub, inv['diskon'], inv['ongkir'], inv['biaya_lain'])
    ppn = float(inv['pajak_nominal'] or 0)
    pph = float(inv['pph23_nominal'] or 0)
    pph22 = float((inv['pph22_nominal'] if 'pph22_nominal' in inv.keys() else 0) or 0)
    piutang = dpp + ppn - pph - pph22
    if dpp <= 0:
        raise ValueError('DPP invoice harus lebih dari 0 untuk dijurnal.')
    entries = [
        ('1120', piutang, 0),   # Dr Piutang Usaha
        ('4100', 0, dpp),       # Cr Pendapatan Penjualan
    ]
    if ppn > 0:
        entries.append(('2111', 0, ppn))  # Cr Hutang PPN
    if pph > 0:
        entries.append(('1181', pph, 0))  # Dr PPh 23 Dibayar Dimuka
    if pph22 > 0:
        entries.append(('1183', pph22, 0))  # Dr PPh 22 Dibayar Dimuka
    keterangan = f"Penjualan: Invoice {inv['nomor']} - {inv['pelanggan']}"
    jid = insert_jurnal(conn, inv['tanggal'], keterangan, 'OPERASIONAL', 'JUAL',
                        entries, pihak=inv['pelanggan'])
    conn.execute("UPDATE invoice SET jurnal_id=?, status=? WHERE id=?", (jid, 'SENT', inv_id))
    # Catat piutang juga (agar dashboard piutang & aging sinkron)
    conn.execute(
        "INSERT INTO piutang(tanggal,jatuh_tempo,pelanggan,keterangan,jumlah,terbayar,status,jurnal_id,invoice_id) "
        "VALUES(?,?,?,?,?,0,'BELUM LUNAS',?,?)",
        (inv['tanggal'], inv['jatuh_tempo'], inv['pelanggan'],
         f"Invoice {inv['nomor']}", piutang, jid, inv_id)
    )
    return jid

def _invoice_tx_meta(jurnal):
    try:
        return json.loads((jurnal['tx_meta'] if 'tx_meta' in jurnal.keys() else '') or '{}')
    except Exception:
        return {}

def _sync_invoice_from_transaction(conn, inv_id, jurnal_id):
    """Salin ulang bagian finansial invoice yang bersumber dari transaksi penjualan."""
    inv = conn.execute("SELECT * FROM invoice WHERE id=?", (inv_id,)).fetchone()
    jurnal = conn.execute("SELECT * FROM jurnal WHERE id=?", (jurnal_id,)).fetchone()
    if not inv or not jurnal:
        raise ValueError('Invoice atau transaksi sumber tidak ditemukan.')
    meta = _invoice_tx_meta(jurnal)
    if meta.get('tipe') != 'JUAL':
        raise ValueError('Invoice transaksi hanya dapat dibuat dari penjualan.')
    items = _enrich_item_varian(conn, conn.execute(
        "SELECT * FROM transaksi_item WHERE jurnal_id=? ORDER BY id", (jurnal_id,)
    ).fetchall())
    if not items:
        raise ValueError('Transaksi penjualan tidak memiliki rincian item.')
    pelanggan = (meta.get('pihak') or jurnal['pihak'] or inv['pelanggan'] or 'Customer').strip()
    # Pajak ikut tersinkron kalau transaksi sumber menyimpan ppn/pph23/pph22 di tx_meta.
    ppn_pct = float(meta.get('ppn_persen', 0) or 0)
    pph_pct = float(meta.get('pph23_persen', 0) or 0)
    pph22_pct = float(meta.get('pph22_persen', 0) or 0)
    conn.execute("""UPDATE invoice SET tanggal=?,pelanggan=?,diskon=?,ongkir=?,biaya_lain=?,
                    pajak_persen=?,pph23_persen=?,pph22_persen=?
                    WHERE id=?""",
                 (jurnal['tanggal'], pelanggan or 'Customer',
                  float(meta.get('diskon', 0) or 0),
                  float(meta.get('ongkir', 0) or 0),
                  float(meta.get('biaya', 0) or 0),
                  ppn_pct, pph_pct, pph22_pct, inv_id))
    conn.execute("DELETE FROM invoice_item WHERE invoice_id=?", (inv_id,))
    for item in items:
        qty = float(item['qty'] or 0)
        harga = float(item['harga'] or 0)
        diskon = float(item['diskon'] or 0)
        subtotal = max(0, qty * harga - diskon)
        conn.execute(
            """INSERT INTO invoice_item(invoice_id,deskripsi,qty,satuan,harga_satuan,diskon_item,subtotal)
               VALUES(?,?,?,?,?,?,?)""",
            (inv_id, item['deskripsi'] or jurnal['keterangan'], qty,
             item['satuan'] or 'unit', harga, diskon, subtotal)
        )
    # Recompute pajak_nominal & pph23_nominal berdasarkan item yang baru disinkron.
    _recalc_invoice_pajak(conn, inv_id)

def _ensure_transaction_invoice(conn, jurnal_id, overrides=None):
    """Ambil invoice transaksi atau buat draft baru saat pertama kali dicetak."""
    overrides = overrides or {}
    existing = conn.execute("SELECT * FROM invoice WHERE jurnal_id=?", (jurnal_id,)).fetchone()
    if existing:
        if existing['status'] == 'DRAFT':
            _sync_invoice_from_transaction(conn, existing['id'], jurnal_id)
        return existing['id']
    jurnal = conn.execute("SELECT * FROM jurnal WHERE id=?", (jurnal_id,)).fetchone()
    if not jurnal:
        raise ValueError('Transaksi tidak ditemukan.')
    meta = _invoice_tx_meta(jurnal)
    if meta.get('tipe') != 'JUAL':
        raise ValueError('Invoice hanya tersedia untuk transaksi penjualan.')
    inv_s = _inv_settings(conn)
    try:
        top_hari = int(inv_s['inv_top'] or 30)
    except (TypeError, ValueError):
        top_hari = 30
    nomor = (overrides.get('nomor') or '').strip() or _next_inv_number(conn, inv_s['inv_prefix'])
    if conn.execute("SELECT id FROM invoice WHERE nomor=?", (nomor,)).fetchone():
        raise ValueError(f'Nomor invoice {nomor} sudah digunakan.')
    tanggal = jurnal['tanggal']
    jatuh_tempo = (overrides.get('jatuh_tempo') or '').strip() or meta.get('jt') or \
        (date.fromisoformat(tanggal) + timedelta(days=top_hari)).strftime('%Y-%m-%d')
    if not is_valid_iso_date(jatuh_tempo):
        raise ValueError('Tanggal jatuh tempo invoice tidak valid.')
    pelanggan = (meta.get('pihak') or jurnal['pihak'] or overrides.get('pelanggan') or 'Customer').strip()
    inv_id = conn.execute(
        """INSERT INTO invoice(jurnal_id,nomor,tanggal,jatuh_tempo,top_hari,top_note,pelanggan,
                    alamat_pelanggan,telepon_pelanggan,catatan,status,dibuat)
           VALUES(?,?,?,?,?,?,?,?,?,?,'DRAFT',?)""",
        (jurnal_id, nomor, tanggal, jatuh_tempo, top_hari, inv_s['inv_top_note'],
         pelanggan or 'Customer', overrides.get('alamat', ''), overrides.get('telepon', ''),
         overrides.get('catatan', ''), tanggal)
    ).lastrowid
    _sync_invoice_from_transaction(conn, inv_id, jurnal_id)
    return inv_id

def _locked_transaction_invoice(conn, jurnal_id):
    return conn.execute(
        "SELECT * FROM invoice WHERE jurnal_id=? AND status!='DRAFT' LIMIT 1", (jurnal_id,)
    ).fetchone()

_LABEL_STATUS_INVOICE = {'DRAFT': 'Draft', 'SENT': 'Terkirim', 'PAID': 'Lunas'}


def _gabung_kalimat(bagian):
    """['a'] -> 'a'; ['a','b','c'] -> 'a, b dan c'."""
    return bagian[0] if len(bagian) == 1 else ', '.join(bagian[:-1]) + ' dan ' + bagian[-1]


def _info_kunci_transaksi(conn, jid):
    """Apa yang mengunci transaksi dari edit & hapus di sisi invoice/piutang, dan apakah bisa
    dibuka lewat rute buka kunci. Dict:
      invoice    : invoice transaksi yang statusnya bukan Draft, atau None
      pembayaran : bayar_piutang atas piutang transaksi ini (nomor & rekening jurnalnya ikut).
                   Inilah yang dibatalkan saat kunci dibuka.
      penghalang : yang TIDAK dibatalkan otomatis karena ikut mengubah hal lain (retur atau
                   penyesuaian tagihan, write-off, pembayaran tanpa jurnal tertaut, jurnal
                   pembayaran yang bercampur transaksi lain). Harus dihapus manual dulu.
      alasan     : kalimat pendek untuk layar ("invoice INV-1 berstatus Lunas dan ...")
      total_bayar, terkunci, bisa_dibuka
    Sisi hutang (pembelian yang sudah dibayar) sengaja tidak ditangani di sini."""
    inv = _locked_transaction_invoice(conn, jid)
    pembayaran, penghalang = [], []
    for b in conn.execute(
            """SELECT b.tanggal, b.jumlah, b.catatan, j.id AS jid, j.nomor_tx, j.keterangan
               FROM bayar_piutang b JOIN piutang p ON p.id=b.piutang_id
               LEFT JOIN jurnal j ON j.id=b.jurnal_id
               WHERE p.jurnal_id=? ORDER BY b.tanggal, b.id""", (jid,)).fetchall():
        if not b['jid'] or b['jid'] == jid:
            penghalang.append({'jenis': 'Pembayaran lama tanpa jurnal', 'jurnal_id': None,
                               'nomor': b['catatan'] or 'tanpa catatan', 'tanggal': b['tanggal'],
                               'jumlah': float(b['jumlah'] or 0)})
            continue
        rek = conn.execute(
            """SELECT a.kode, a.nama FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id
               WHERE d.jurnal_id=? AND d.debit>0 ORDER BY d.id LIMIT 1""", (b['jid'],)).fetchone()
        item = {'jurnal_id': b['jid'], 'nomor': b['nomor_tx'] or ('#' + str(b['jid'])),
                'tanggal': b['tanggal'], 'jumlah': float(b['jumlah'] or 0),
                'keterangan': b['keterangan'] or '',
                'rekening': (rek['kode'] + ' ' + rek['nama']) if rek else ''}
        # Jurnal pembayaran normal hanya memuat SATU baris bayar_piutang. Kalau memuat hal lain
        # (stok, tagihan lain), menghapusnya ikut membatalkan hal itu: jangan otomatis.
        lain = conn.execute(
            """SELECT (SELECT COUNT(*) FROM bayar_piutang WHERE jurnal_id=:j)
                    + (SELECT COUNT(*) FROM bayar_hutang WHERE jurnal_id=:j)
                    + (SELECT COUNT(*) FROM penyesuaian_tagihan WHERE jurnal_id=:j)
                    + (SELECT COUNT(*) FROM writeoff_piutang WHERE jurnal_id=:j)
                    + (SELECT COUNT(*) FROM piutang WHERE jurnal_id=:j)
                    + (SELECT COUNT(*) FROM hutang WHERE jurnal_id=:j)
                    + (SELECT COUNT(*) FROM pergerakan_stok WHERE jurnal_id=:j)
                    + (SELECT COUNT(*) FROM transaksi_item WHERE jurnal_id=:j)""",
            {'j': b['jid']}).fetchone()[0]
        if lain > 1:
            item['jenis'] = 'Pembayaran yang jurnalnya bercampur transaksi lain'
            penghalang.append(item)
        else:
            pembayaran.append(item)
    for x in conn.execute(
            """SELECT x.jurnal_id, x.jumlah_delta, j.nomor_tx, j.tanggal
               FROM penyesuaian_tagihan x JOIN piutang p ON x.jenis='PIUTANG' AND p.id=x.record_id
               LEFT JOIN jurnal j ON j.id=x.jurnal_id
               WHERE p.jurnal_id=? ORDER BY x.id""", (jid,)).fetchall():
        penghalang.append({'jenis': 'Retur atau penyesuaian tagihan', 'jurnal_id': x['jurnal_id'],
                           'nomor': x['nomor_tx'] or ('#' + str(x['jurnal_id'])), 'tanggal': x['tanggal'],
                           'jumlah': abs(float(x['jumlah_delta'] or 0))})
    for w in conn.execute(
            """SELECT w.jurnal_id, w.jumlah, w.tanggal, j.nomor_tx
               FROM writeoff_piutang w JOIN piutang p ON p.id=w.piutang_id
               LEFT JOIN jurnal j ON j.id=w.jurnal_id
               WHERE p.jurnal_id=? ORDER BY w.id""", (jid,)).fetchall():
        penghalang.append({'jenis': 'Write-off piutang', 'jurnal_id': w['jurnal_id'],
                           'nomor': w['nomor_tx'] or ('#' + str(w['jurnal_id'])), 'tanggal': w['tanggal'],
                           'jumlah': float(w['jumlah'] or 0)})
    total_bayar = sum(b['jumlah'] for b in pembayaran)
    alasan = []
    if inv:
        alasan.append('invoice %s berstatus %s' % (inv['nomor'], _LABEL_STATUS_INVOICE.get(inv['status'], inv['status'])))
    if pembayaran:
        alasan.append('sudah ada %d pembayaran %s atas tagihannya' % (len(pembayaran), rp_filter(total_bayar)))
    if penghalang:
        jenis = []
        for x in penghalang:
            if x['jenis'].lower() not in jenis:
                jenis.append(x['jenis'].lower())
        alasan.append('ada ' + _gabung_kalimat(jenis))
    terkunci = bool(inv or pembayaran or penghalang)
    return {'invoice': inv, 'pembayaran': pembayaran, 'penghalang': penghalang,
            'total_bayar': total_bayar, 'alasan': _gabung_kalimat(alasan) if alasan else '',
            'terkunci': terkunci, 'bisa_dibuka': terkunci and not penghalang}


def _pesan_transaksi_terkunci(conn, jid):
    """Pesan kenapa transaksi belum bisa diedit (invoice bukan Draft, atau sudah ada
    pembayaran/retur/write-off), atau None kalau bebas. Syarat blokirnya sama persis dengan
    dua pemeriksaan lama (invoice terkunci ATAU transaction_dependency_count)."""
    li = _locked_transaction_invoice(conn, jid)
    if li:
        return ('Transaksi ini terkunci karena invoice %s berstatus %s. Petunjuk membuka kuncinya '
                'ada di bawah.' % (li['nomor'], _LABEL_STATUS_INVOICE.get(li['status'], li['status'])))
    if transaction_dependency_count(conn, jid):
        if _info_kunci_transaksi(conn, jid)['terkunci']:
            return ('Transaksi ini terkunci karena sudah punya pembayaran, retur, atau write-off. '
                    'Petunjuk membuka kuncinya ada di bawah.')
        return ('Transaksi memiliki pembayaran atau retur terkait. Hapus jurnal pelunasan/retur '
                'terkait lebih dulu agar saldo tetap sinkron.')
    return None

@app.route('/invoice/preview')
@investor_required
def invoice_preview():
    conn = db()
    inv_s = _inv_settings(conn)
    conn.close()
    today_str = date.today().strftime('%Y-%m-%d')
    top = int(inv_s['inv_top'] or 30)
    jt  = (date.today() + timedelta(days=top)).strftime('%Y-%m-%d')
    inv = {
        'nomor': f"{inv_s['inv_prefix']}-{date.today().strftime('%Y%m')}-001",
        'tanggal': today_str,
        'jatuh_tempo': jt,
        'top_hari': top,
        'top_note': inv_s['inv_top_note'],
        'pelanggan': 'Contoh Pelanggan / PT Maju Bersama',
        'alamat_pelanggan': 'Jl. Contoh No. 123, Jakarta',
        'telepon_pelanggan': '0812-3456-7890',
        'diskon': 50000,
        'ongkir': 25000,
        'biaya_lain': 0,
        'catatan': '',
        'status': 'DRAFT',
    }
    items = [
        {'deskripsi': 'Produk / Jasa Contoh A', 'qty': 2, 'satuan': 'pcs',
         'harga_satuan': 500000, 'diskon_item': 0, 'subtotal': 1000000},
        {'deskripsi': 'Produk / Jasa Contoh B', 'qty': 1, 'satuan': 'unit',
         'harga_satuan': 750000, 'diskon_item': 25000, 'subtotal': 725000},
    ]
    subtotal = sum(it['subtotal'] for it in items)
    dpp      = _calc_invoice_dpp(subtotal, inv['diskon'], inv['ongkir'], inv['biaya_lain'])
    pajak_pct = float(inv.get('pajak_persen') or 0)
    pph23_pct = float(inv.get('pph23_persen') or 0)
    pph22_pct = float(inv.get('pph22_persen') or 0)
    pajak_nom = round(dpp * pajak_pct / 100.0, 2)
    pph23_nom = round(dpp * pph23_pct / 100.0, 2)
    pph22_nom = round(dpp * pph22_pct / 100.0, 2)
    total    = dpp + pajak_nom
    return render_template('invoice_print.html', inv=inv, items=items, inv_s=inv_s,
                           subtotal=subtotal, total=total, autoprint=False, preview=True,
                           pajak_pct=pajak_pct, pajak_nom=pajak_nom,
                           pph23_pct=pph23_pct, pph23_nom=pph23_nom,
                           pph22_pct=pph22_pct, pph22_nom=pph22_nom,
                           actual_top=top)

@app.route('/invoice/baru', methods=['GET','POST'])
@operator_required
def invoice_baru():
    conn = db()
    inv_s = _inv_settings(conn)
    produk_list = conn.execute("SELECT * FROM produk WHERE COALESCE(aktif,1)=1 ORDER BY nama").fetchall()
    if request.method == 'POST':
        prefix = inv_s['inv_prefix']
        nomor  = request.form.get('nomor','').strip() or _next_inv_number(conn, prefix)
        if conn.execute("SELECT id FROM invoice WHERE nomor=?", (nomor,)).fetchone():
            flash(f'Nomor invoice {nomor} sudah digunakan.', 'danger')
            conn.close()
            return redirect(url_for('invoice_baru'))
        top_hari = int(request.form.get('top_hari') or inv_s['inv_top'] or 30)
        top_note = request.form.get('top_note', inv_s['inv_top_note']).strip()
        tanggal  = request.form['tanggal']
        if not is_valid_iso_date(tanggal):
            conn.close(); flash('Tanggal invoice tidak valid.', 'danger')
            return redirect(url_for('invoice_baru'))
        jt = request.form.get('jatuh_tempo','').strip() or \
             (date.fromisoformat(tanggal) + timedelta(days=top_hari)).strftime('%Y-%m-%d')
        if not is_valid_iso_date(jt):
            conn.close(); flash('Tanggal jatuh tempo invoice tidak valid.', 'danger')
            return redirect(url_for('invoice_baru'))
        try:
            pajak_persen = parse_percentage(request.form.get('pajak_persen', '0'), 'Persen pajak')
            pph23_persen = parse_percentage(request.form.get('pph23_persen', '0'), 'Persen PPh 23') if is_module_active(conn, 'PAJAK_OTOMATIS') else 0.0
            pph22_persen = parse_percentage(request.form.get('pph22_persen', '0'), 'Persen PPh 22') if is_module_active(conn, 'PAJAK_OTOMATIS') else 0.0
        except ValueError as ex:
            flash(str(ex), 'danger'); conn.close()
            return redirect(url_for('invoice_baru'))
        if pajak_persen < 0: pajak_persen = 0
        if pajak_persen > 100: pajak_persen = 100
        # PPh 23 & PPh 22 hanya tersedia kalau modul PAJAK_OTOMATIS aktif
        if pph23_persen < 0: pph23_persen = 0
        if pph23_persen > 100: pph23_persen = 100
        if pph22_persen < 0: pph22_persen = 0
        if pph22_persen > 100: pph22_persen = 100
        inv_id = conn.execute(
            "INSERT INTO invoice(nomor,tanggal,jatuh_tempo,top_hari,top_note,pelanggan,alamat_pelanggan,telepon_pelanggan,diskon,ongkir,biaya_lain,catatan,pajak_persen,pph23_persen,pph22_persen) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (nomor, tanggal, jt, top_hari, top_note,
             request.form.get('pelanggan',''),
             request.form.get('alamat_pelanggan',''),
             request.form.get('telepon_pelanggan',''),
             parse_rp(request.form.get('diskon','0')),
             parse_rp(request.form.get('ongkir','0')),
             parse_rp(request.form.get('biaya_lain','0')),
             request.form.get('catatan',''),
             pajak_persen, pph23_persen, pph22_persen)
        ).lastrowid
        try:
            _save_invoice_items(conn, inv_id, request.form)
            _recalc_invoice_pajak(conn, inv_id)
        except ValueError as ex:
            conn.rollback(); conn.close()
            flash(str(ex), 'danger')
            return redirect(url_for('invoice_baru'))
        add_log(conn, 'Invoice dibuat', f"{nomor} | {request.form.get('pelanggan','')}", 'INPUT')
        conn.commit(); conn.close()
        flash(f'Invoice {nomor} berhasil dibuat!', 'success')
        return redirect(url_for('invoice_view', id=inv_id))
    next_num = _next_inv_number(conn, inv_s['inv_prefix'])
    customer_list = [dict(r) for r in conn.execute(
        "SELECT nama,telepon,alamat FROM customer ORDER BY nama").fetchall()]
    conn.close()
    return render_template('invoice_form.html', inv_s=inv_s, produk_list=produk_list,
                           next_num=next_num, today=date.today().strftime('%Y-%m-%d'),
                           edit=None, items=[], customer_list=customer_list)

@app.route('/invoice/<int:id>/edit', methods=['GET','POST'])
@operator_required
def invoice_edit(id):
    conn = db()
    inv = conn.execute("SELECT * FROM invoice WHERE id=?", (id,)).fetchone()
    if not inv:
        conn.close(); flash('Invoice tidak ditemukan.','danger')
        return redirect(url_for('invoice_list'))
    inv_s = _inv_settings(conn)
    produk_list = conn.execute("SELECT * FROM produk WHERE COALESCE(aktif,1)=1 ORDER BY nama").fetchall()
    if request.method == 'POST':
        if inv['status'] == 'PAID' and session.get('role') not in ('ADMIN', 'FINANCE'):
            conn.close(); flash('Invoice yang sudah lunas hanya bisa diedit oleh Finance/Admin.', 'danger')
            return redirect(url_for('invoice_view', id=id))
        nomor_e = request.form.get('nomor', inv['nomor']).strip() or inv['nomor']
        if nomor_e != inv['nomor'] and conn.execute("SELECT id FROM invoice WHERE nomor=? AND id!=?", (nomor_e, id)).fetchone():
            conn.close(); flash(f'Nomor invoice {nomor_e} sudah digunakan.', 'danger')
            return redirect(url_for('invoice_edit', id=id))
        top_hari_e = int(request.form.get('top_hari') or inv_s['inv_top'] or 30)
        top_note_e = request.form.get('top_note', '').strip()
        tanggal_e  = request.form.get('tanggal', inv['tanggal'])
        if not is_valid_iso_date(tanggal_e):
            conn.close(); flash('Tanggal invoice tidak valid.', 'danger')
            return redirect(url_for('invoice_edit', id=id))
        jt_e = request.form.get('jatuh_tempo','').strip() or \
               (date.fromisoformat(tanggal_e) + timedelta(days=top_hari_e)).strftime('%Y-%m-%d')
        if not is_valid_iso_date(jt_e):
            conn.close(); flash('Tanggal jatuh tempo invoice tidak valid.', 'danger')
            return redirect(url_for('invoice_edit', id=id))
        if inv['jurnal_id']:
            conn.execute("""UPDATE invoice SET nomor=?,jatuh_tempo=?,top_hari=?,top_note=?,
                            alamat_pelanggan=?,telepon_pelanggan=?,catatan=? WHERE id=?""",
                         (nomor_e, jt_e, top_hari_e, top_note_e,
                          request.form.get('alamat_pelanggan',''),
                          request.form.get('telepon_pelanggan',''),
                          request.form.get('catatan',''), id))
            # Hanya invoice yang benar-benar berasal dari transaksi penjualan (tx_meta.tipe=JUAL)
            # yang boleh disinkron ulang dari item transaksi. Invoice manual yang sudah di-post
            # ke jurnal (Dr Piutang/Cr Pendapatan) tidak punya transaksi sumber, jadi cukup update
            # kolom non-finansial di atas tanpa _sync (mencegah HTTP 500 - ValueError tak tertangkap).
            _jr = conn.execute("SELECT tx_meta FROM jurnal WHERE id=?", (inv['jurnal_id'],)).fetchone()
            if _jr and _invoice_tx_meta(_jr).get('tipe') == 'JUAL':
                _sync_invoice_from_transaction(conn, id, inv['jurnal_id'])
        else:
            try:
                pajak_persen_e = parse_percentage(request.form.get('pajak_persen', '0'), 'Persen pajak')
                pph23_persen_e = parse_percentage(request.form.get('pph23_persen', '0'), 'Persen PPh 23') if is_module_active(conn, 'PAJAK_OTOMATIS') else 0.0
                pph22_persen_e = parse_percentage(request.form.get('pph22_persen', '0'), 'Persen PPh 22') if is_module_active(conn, 'PAJAK_OTOMATIS') else 0.0
            except ValueError as ex:
                flash(str(ex), 'danger'); conn.close()
                return redirect(url_for('invoice_edit', id=id))
            if pajak_persen_e < 0: pajak_persen_e = 0
            if pajak_persen_e > 100: pajak_persen_e = 100
            if pph23_persen_e < 0: pph23_persen_e = 0
            if pph23_persen_e > 100: pph23_persen_e = 100
            if pph22_persen_e < 0: pph22_persen_e = 0
            if pph22_persen_e > 100: pph22_persen_e = 100
            conn.execute("""UPDATE invoice SET nomor=?,tanggal=?,jatuh_tempo=?,top_hari=?,top_note=?,
                            pelanggan=?,alamat_pelanggan=?,telepon_pelanggan=?,
                            diskon=?,ongkir=?,biaya_lain=?,catatan=?,pajak_persen=?,pph23_persen=?,pph22_persen=? WHERE id=?""",
                         (nomor_e, tanggal_e, jt_e, top_hari_e, top_note_e,
                          request.form.get('pelanggan',''),
                          request.form.get('alamat_pelanggan',''),
                          request.form.get('telepon_pelanggan',''),
                          parse_rp(request.form.get('diskon','0')),
                          parse_rp(request.form.get('ongkir','0')),
                          parse_rp(request.form.get('biaya_lain','0')),
                          request.form.get('catatan',''), pajak_persen_e, pph23_persen_e, pph22_persen_e, id))
            try:
                _save_invoice_items(conn, id, request.form)
                _recalc_invoice_pajak(conn, id)
            except ValueError as ex:
                conn.rollback(); conn.close()
                flash(str(ex), 'danger')
                return redirect(url_for('invoice_edit', id=id))
        conn.commit(); conn.close()
        flash('Invoice diperbarui!', 'success')
        return redirect(url_for('invoice_view', id=id))
    items = conn.execute("SELECT * FROM invoice_item WHERE invoice_id=? ORDER BY id", (id,)).fetchall()
    customer_list = [dict(r) for r in conn.execute(
        "SELECT nama,telepon,alamat FROM customer ORDER BY nama").fetchall()]
    conn.close()
    return render_template('invoice_form.html', inv_s=inv_s, produk_list=produk_list,
                           today=date.today().strftime('%Y-%m-%d'), edit=inv, items=items,
                           customer_list=customer_list)

@app.route('/transaksi/<int:id>/invoice/cetak')
@operator_required
def transaksi_invoice_cetak(id):
    if session.get('role') == 'DEMO':
        flash('User Demo tidak bisa membuat invoice.', 'warning')
        return redirect(url_for('transaksi_detail', id=id))
    conn = db()
    try:
        inv_id = _ensure_transaction_invoice(conn, id)
        conn.commit()
    except ValueError as ex:
        conn.rollback(); conn.close()
        flash(str(ex), 'danger')
        return redirect(url_for('transaksi_detail', id=id))
    conn.close()
    return redirect(url_for('invoice_cetak', id=inv_id))

@app.route('/invoice/<int:id>')
@invoice_read_required
def invoice_view(id):
    conn = db()
    inv   = conn.execute("SELECT * FROM invoice WHERE id=?", (id,)).fetchone()
    if not inv:
        conn.close(); flash('Invoice tidak ditemukan.','danger')
        return redirect(url_for('invoice_list'))
    # Sinkron ulang HANYA untuk invoice-transaksi (sumber penjualan, tx_meta.tipe=JUAL).
    # Invoice manual yang di-post ke jurnal punya jurnal_id tapi bukan penjualan -> jangan
    # di-_sync (mencegah HTTP 500 - ValueError tak tertangkap).
    if inv['jurnal_id'] and inv['status'] == 'DRAFT':
        _jr = conn.execute("SELECT tx_meta FROM jurnal WHERE id=?", (inv['jurnal_id'],)).fetchone()
        if _jr and _invoice_tx_meta(_jr).get('tipe') == 'JUAL':
            _sync_invoice_from_transaction(conn, id, inv['jurnal_id'])
            conn.commit()
            inv = conn.execute("SELECT * FROM invoice WHERE id=?", (id,)).fetchone()
    items = conn.execute("SELECT * FROM invoice_item WHERE invoice_id=? ORDER BY id", (id,)).fetchall()
    inv_s = _inv_settings(conn)
    # Piutang terbuka + rekening kas (untuk form Terima DP/Pembayaran di invoice_view).
    _inv_piutang = _invoice_open_piutang(conn, inv)
    akun_kas = conn.execute("SELECT * FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    # Tombol "Buka Kunci" bila transaksi sumbernya terkunci (invoice bukan Draft / sudah dibayar).
    kunci_tx = _info_kunci_transaksi(conn, inv['jurnal_id']) if inv['jurnal_id'] else None
    conn.close()
    if _inv_piutang:
        inv_terbayar = float(_inv_piutang['terbayar'] or 0)
        inv_sisa = max(0.0, float(_inv_piutang['jumlah'] or 0) - inv_terbayar)
        inv_has_piutang = True
    else:
        inv_terbayar = 0.0; inv_sisa = 0.0; inv_has_piutang = False
    subtotal  = sum(it['subtotal'] for it in items)
    dpp       = _calc_invoice_dpp(subtotal, inv['diskon'], inv['ongkir'], inv['biaya_lain'])
    pajak_pct = float(inv['pajak_persen'] if 'pajak_persen' in inv.keys() else 0) or 0
    pph23_pct = float(inv['pph23_persen'] if 'pph23_persen' in inv.keys() else 0) or 0
    pph22_pct = float(inv['pph22_persen'] if 'pph22_persen' in inv.keys() else 0) or 0
    pajak_nom = round(dpp * pajak_pct / 100.0, 2)
    pph23_nom = round(dpp * pph23_pct / 100.0, 2)
    pph22_nom = round(dpp * pph22_pct / 100.0, 2)
    total     = dpp + pajak_nom
    actual_top = _calc_top(inv)
    return render_template('invoice_view.html', inv=inv, items=items, inv_s=inv_s,
                           subtotal=subtotal, total=total, actual_top=actual_top,
                           pajak_pct=pajak_pct, pajak_nom=pajak_nom,
                           pph23_pct=pph23_pct, pph23_nom=pph23_nom,
                           pph22_pct=pph22_pct, pph22_nom=pph22_nom,
                           inv_has_piutang=inv_has_piutang, inv_sisa=inv_sisa,
                           inv_terbayar=inv_terbayar, akun_kas=akun_kas, kunci_tx=kunci_tx)

@app.route('/invoice/<int:id>/cetak')
@invoice_read_required
def invoice_cetak(id):
    conn = db()
    inv   = conn.execute("SELECT * FROM invoice WHERE id=?", (id,)).fetchone()
    if not inv:
        conn.close(); return redirect(url_for('invoice_list'))
    # Sinkron ulang HANYA untuk invoice-transaksi (sumber penjualan, tx_meta.tipe=JUAL);
    # invoice manual yang di-post ke jurnal jangan di-_sync (mencegah HTTP 500).
    if inv['jurnal_id'] and inv['status'] == 'DRAFT':
        _jr = conn.execute("SELECT tx_meta FROM jurnal WHERE id=?", (inv['jurnal_id'],)).fetchone()
        if _jr and _invoice_tx_meta(_jr).get('tipe') == 'JUAL':
            _sync_invoice_from_transaction(conn, id, inv['jurnal_id'])
            conn.commit()
            inv = conn.execute("SELECT * FROM invoice WHERE id=?", (id,)).fetchone()
    items = conn.execute("SELECT * FROM invoice_item WHERE invoice_id=? ORDER BY id", (id,)).fetchall()
    inv_s = _inv_settings(conn)
    subtotal = sum(it['subtotal'] for it in items)
    dpp      = _calc_invoice_dpp(subtotal, inv['diskon'], inv['ongkir'], inv['biaya_lain'])
    pajak_pct = float(inv['pajak_persen'] if 'pajak_persen' in inv.keys() else 0) or 0
    pph23_pct = float(inv['pph23_persen'] if 'pph23_persen' in inv.keys() else 0) or 0
    pph22_pct = float(inv['pph22_persen'] if 'pph22_persen' in inv.keys() else 0) or 0
    pajak_nom = round(dpp * pajak_pct / 100.0, 2)
    pph23_nom = round(dpp * pph23_pct / 100.0, 2)
    pph22_nom = round(dpp * pph22_pct / 100.0, 2)
    total    = dpp + pajak_nom

    # Invoice tertaut memakai jurnal_id. Invoice manual lama memakai fallback historis.
    if inv['jurnal_id']:
        piutang_row = conn.execute(
            "SELECT * FROM piutang WHERE jurnal_id=? ORDER BY id DESC LIMIT 1",
            (inv['jurnal_id'],)
        ).fetchone()
    else:
        piutang_row = conn.execute(
            "SELECT * FROM piutang WHERE pelanggan=? AND keterangan LIKE ? ORDER BY id DESC LIMIT 1",
            (inv['pelanggan'], f"%{inv['nomor']}%")
        ).fetchone()
    # Prioritas 2: tanggal + pelanggan sama, jumlah <= total invoice (piutang partial/full)
    if not inv['jurnal_id'] and not piutang_row:
        piutang_row = conn.execute(
            """SELECT * FROM piutang WHERE pelanggan=? AND tanggal=?
               AND jumlah <= ? AND jumlah > 0
               ORDER BY ABS(jumlah - ?) LIMIT 1""",
            (inv['pelanggan'], inv['tanggal'], total + 1, total)
        ).fetchone()
    # Prioritas 3: pelanggan + jumlah tepat sama (tanpa batasan tanggal)
    if not inv['jurnal_id'] and not piutang_row:
        piutang_row = conn.execute(
            "SELECT * FROM piutang WHERE pelanggan=? AND ROUND(jumlah,0)=ROUND(?,0) ORDER BY id DESC LIMIT 1",
            (inv['pelanggan'], total)
        ).fetchone()

    # Tampilkan pembayaran yang sudah tercatat dan sisa tagihan saat ini.
    if piutang_row:
        p_jumlah     = float(piutang_row['jumlah'] or 0)
        p_terbayar   = float(piutang_row['terbayar'] or 0)
        pay_sisa     = max(0, p_jumlah - p_terbayar)
        pay_dp       = max(0, float(total) - pay_sisa)
        pay_stamp    = piutang_row['status']
        show_payment = True
    elif inv['jurnal_id'] or inv['status'] == 'PAID':
        pay_dp       = float(total)
        pay_sisa     = 0.0
        pay_stamp    = 'LUNAS'
        show_payment = True
    elif inv['jatuh_tempo']:
        pay_dp       = 0.0
        pay_sisa     = float(total)
        pay_stamp    = 'BELUM LUNAS'
        show_payment = True
    else:
        pay_dp = pay_sisa = 0.0
        pay_stamp    = None
        show_payment = False

    conn.close()
    autoprint  = request.args.get('autoprint','0') == '1'
    actual_top = _calc_top(inv)
    return render_template('invoice_print.html', inv=inv, items=items, inv_s=inv_s,
                           subtotal=subtotal, total=total, autoprint=autoprint,
                           actual_top=actual_top,
                           pajak_pct=pajak_pct, pajak_nom=pajak_nom,
                           pph23_pct=pph23_pct, pph23_nom=pph23_nom,
                           pph22_pct=pph22_pct, pph22_nom=pph22_nom,
                           show_payment=show_payment,
                           pay_dp=pay_dp, pay_sisa=pay_sisa, pay_stamp=pay_stamp)

def _invoice_open_piutang(conn, inv):
    """Piutang yang terkait invoice. Tautannya jurnal_id: invoice transaksi maupun invoice
    yang di-post ke jurnal berbagi jurnal_id dengan piutangnya (cetak invoice juga begitu).
    invoice_id hanya dipakai untuk piutang tanpa jurnal. Dulu `jurnal_id=? OR invoice_id=?`,
    sehingga piutang milik transaksi LAIN yang invoice_id-nya nyasar (Restore "tambah data"
    versi lama) ikut terbaca: invoice tunai tanpa piutang pun terkunci Lunas, tak bisa kembali
    ke Draft, dan tandai Lunas bisa melunasi piutang pelanggan lain."""
    if not inv['jurnal_id']:
        return None
    p = conn.execute(
        "SELECT * FROM piutang WHERE jurnal_id=? ORDER BY id DESC LIMIT 1", (inv['jurnal_id'],)
    ).fetchone()
    if p:
        return p
    return conn.execute(
        "SELECT * FROM piutang WHERE invoice_id=? AND jurnal_id IS NULL ORDER BY id DESC LIMIT 1",
        (inv['id'],)
    ).fetchone()

def _settle_invoice_on_paid(conn, inv):
    """Saat invoice ditandai LUNAS: catat penerimaan kas & lunasi piutang terkait.
    Hanya untuk invoice yang punya piutang terbuka (invoice manual yang sudah di-post ke
    jurnal, atau invoice transaksi penjualan kredit). Idempoten: kalau sudah lunas, no-op."""
    piutang = _invoice_open_piutang(conn, inv)
    if not piutang:
        return
    sisa = float(piutang['jumlah'] or 0) - float(piutang['terbayar'] or 0)
    if sisa <= 0.01:
        return
    tanggal = date.today().strftime('%Y-%m-%d')
    akun_kas = '1100' if is_rekening_kode(conn, '1100') else None
    if not akun_kas:
        row = conn.execute("SELECT kode FROM akun WHERE is_rekening=1 ORDER BY kode LIMIT 1").fetchone()
        akun_kas = row['kode'] if row else '1100'
    pay_jid = insert_jurnal(conn, tanggal,
                            f"Pelunasan piutang via invoice {inv['nomor']}: {piutang['pelanggan']}",
                            'OPERASIONAL', 'PELUNASAN',
                            [(akun_kas, sisa, 0), ('1120', 0, sisa)])
    conn.execute("INSERT INTO bayar_piutang(piutang_id,tanggal,jumlah,catatan,jurnal_id) VALUES(?,?,?,?,?)",
                 (piutang['id'], tanggal, sisa,
                  f"Otomatis: invoice {inv['nomor']} ditandai LUNAS", pay_jid))
    conn.execute("UPDATE piutang SET terbayar=terbayar+? WHERE id=?", (sisa, piutang['id']))
    update_piutang_status(conn, piutang['id'])
    add_log(conn, 'Pelunasan piutang via invoice', f"{inv['nomor']} | Rp {sisa:,.0f}", 'INPUT')

def _unsettle_invoice_on_unpaid(conn, inv):
    """Balik pelunasan otomatis yang dibuat saat invoice ditandai LUNAS (mis. status
    dikembalikan ke Terkirim/Draft). Hanya membatalkan pembayaran yang dibuat otomatis."""
    piutang = _invoice_open_piutang(conn, inv)
    if not piutang:
        return
    rows = conn.execute(
        "SELECT id, jurnal_id FROM bayar_piutang WHERE piutang_id=? AND catatan LIKE ?",
        (piutang['id'], f"Otomatis: invoice {inv['nomor']} ditandai LUNAS%")
    ).fetchall()
    for b in rows:
        if b['jurnal_id']:
            _reverse_tx(conn, b['jurnal_id'])                       # hapus bayar_piutang + kurangi terbayar + update status
            conn.execute("DELETE FROM jurnal WHERE id=?", (b['jurnal_id'],))  # cascade hapus detail_jurnal
        else:
            conn.execute("DELETE FROM bayar_piutang WHERE id=?", (b['id'],))
            update_piutang_status(conn, piutang['id'])

def _apply_invoice_payment(conn, inv, piutang, amount, akun_kas, catatan):
    """Catat pembayaran (mis. DP/uang muka) sebesar `amount` atas piutang invoice:
    Dr Kas/Bank, Cr Piutang Usaha (1120). Mengurangi sisa tagihan + update status invoice
    (SEBAGIAN/LUNAS mengikuti piutang). Return jurnal_id pembayaran."""
    tanggal = date.today().strftime('%Y-%m-%d')
    pay_jid = insert_jurnal(conn, tanggal, catatan, 'OPERASIONAL', 'PELUNASAN',
                            [(akun_kas, amount, 0), ('1120', 0, amount)],
                            pihak=(piutang['pelanggan'] or ''))
    conn.execute("INSERT INTO bayar_piutang(piutang_id,tanggal,jumlah,catatan,jurnal_id) VALUES(?,?,?,?,?)",
                 (piutang['id'], tanggal, amount, catatan, pay_jid))
    conn.execute("UPDATE piutang SET terbayar=terbayar+? WHERE id=?", (amount, piutang['id']))
    update_piutang_status(conn, piutang['id'])
    _sync_invoice_from_piutang(conn, piutang['id'])
    return pay_jid

@app.route('/invoice/<int:id>/bayar', methods=['POST'])
@finance_required
def invoice_bayar(id):
    """Terima pembayaran / DP atas invoice yang sudah tertaut piutang: Dr Kas/Bank, Cr Piutang.
    Mengurangi Sisa Tagihan (terintegrasi dengan 'Sudah Dibayar' di cetak invoice)."""
    conn = db()
    inv = conn.execute("SELECT * FROM invoice WHERE id=?", (id,)).fetchone()
    if not inv:
        conn.close(); flash('Invoice tidak ditemukan.', 'danger')
        return redirect(url_for('invoice_list'))
    piutang = _invoice_open_piutang(conn, inv)
    if not piutang:
        conn.close()
        flash('Invoice ini belum tertaut piutang, jadi pembayaran/DP belum bisa dicatat. '
              'Untuk invoice manual, klik "Catat ke Jurnal" dulu supaya piutang terbentuk.', 'warning')
        return redirect(url_for('invoice_view', id=id))
    sisa = float(piutang['jumlah'] or 0) - float(piutang['terbayar'] or 0)
    if sisa <= 0.01:
        conn.close()
        flash('Invoice ini sudah lunas, tidak ada sisa tagihan.', 'info')
        return redirect(url_for('invoice_view', id=id))
    akun_kas = request.form.get('akun_kas', '1100')
    if not is_rekening_kode(conn, akun_kas):
        conn.close(); flash('Rekening kas/bank tidak valid.', 'danger')
        return redirect(url_for('invoice_view', id=id))
    try:
        amount = parse_nonnegative_rp(request.form.get('nominal', '0'), 'Nominal DP/pembayaran')
    except ValueError as ex:
        conn.close(); flash(str(ex), 'danger')
        return redirect(url_for('invoice_view', id=id))
    if amount <= 0:
        conn.close(); flash('Nominal DP/pembayaran harus lebih dari 0.', 'danger')
        return redirect(url_for('invoice_view', id=id))
    if amount > sisa + 0.01:
        conn.close()
        flash(f'Nominal (Rp {amount:,.0f}) melebihi sisa tagihan Rp {sisa:,.0f}. '
              'Masukkan maksimal sebesar sisa.', 'danger')
        return redirect(url_for('invoice_view', id=id))
    try:
        _apply_invoice_payment(conn, inv, piutang, amount, akun_kas,
                               f"DP/Pembayaran invoice {inv['nomor']}: {piutang['pelanggan']}")
        add_log(conn, 'Terima DP/pembayaran invoice', f"{inv['nomor']} | Rp {amount:,.0f}", 'INPUT')
        conn.commit()
    except Exception as ex:
        conn.rollback(); conn.close()
        flash(f'Gagal mencatat pembayaran: {ex}', 'danger')
        return redirect(url_for('invoice_view', id=id))
    conn.close()
    flash(f'Pembayaran/DP Rp {amount:,.0f} tercatat (Dr Kas, Cr Piutang). '
          f'Sisa tagihan: Rp {max(0, sisa - amount):,.0f}.', 'success')
    return redirect(url_for('invoice_view', id=id))

@app.route('/invoice/<int:id>/status', methods=['POST'])
@operator_required
def invoice_status_update(id):
    new_status = request.form.get('status', 'SENT')
    if new_status not in ('DRAFT', 'SENT', 'PAID'):
        flash('Status invoice tidak valid.', 'danger')
        return redirect(url_for('invoice_view', id=id))
    # Hanya ADMIN dan FINANCE yang boleh set status PAID
    if new_status == 'PAID' and session.get('role') not in ('ADMIN', 'FINANCE'):
        flash('Hanya Admin/Finance yang bisa menandai invoice sebagai Lunas.', 'danger')
        return redirect(url_for('invoice_view', id=id))
    conn = db()
    inv = conn.execute("SELECT * FROM invoice WHERE id=?", (id,)).fetchone()
    if not inv:
        conn.close(); flash('Invoice tidak ditemukan.', 'danger')
        return redirect(url_for('invoice_list'))
    old_status = inv['status']
    try:
        if new_status == 'PAID' and old_status != 'PAID':
            if not inv['jurnal_id']:
                # Invoice manual yang belum di-post ke jurnal: tidak ada piutang/pendapatan/kas tercatat.
                pajak_pct = float(inv['pajak_persen'] if 'pajak_persen' in inv.keys() else 0) or 0
                if pajak_pct > 0 and is_module_active(conn, 'PAJAK_OTOMATIS'):
                    # Invoice PKP (kena PPN) sebaiknya dicatat ke jurnal dulu supaya piutang & pendapatan terbentuk.
                    conn.close()
                    flash('Invoice PKP ini belum dicatat ke jurnal. Klik "Post ke Jurnal" dulu agar '
                          'piutang dan pendapatan tercatat, baru tandai Lunas.', 'warning')
                    return redirect(url_for('invoice_view', id=id))
                # Invoice dokumen biasa: boleh ditandai Lunas, tapi beri tahu tidak ada entri otomatis.
                flash('Invoice ditandai Lunas. Catatan: invoice ini hanya dokumen (belum tertaut jurnal), '
                      'jadi tidak membuat entri kas/pendapatan otomatis. Catat pemasukannya lewat menu '
                      'Pemasukan bila penjualan ini belum dibukukan.', 'info')
            else:
                _settle_invoice_on_paid(conn, inv)
        elif old_status == 'PAID' and new_status != 'PAID':
            _unsettle_invoice_on_unpaid(conn, inv)
            # _unsettle hanya membalik pelunasan OTOMATIS (marker 'Otomatis: ... LUNAS').
            # Kalau piutang masih LUNAS (ada DP/pembayaran nyata, retur, atau write-off), label
            # invoice tidak boleh berubah. Pembalikan otomatis di atas ikut dibatalkan (rollback)
            # supaya tak berhenti setengah jalan, lalu user dibawa ke halaman buka kunci yang
            # menyebut pembayaran mana yang harus dibatalkan, dengan satu tombol. Dulu cuma pesan
            # "hapus dulu pembayaran itu di menu Transaksi" tanpa menyebut yang mana.
            _p = _invoice_open_piutang(conn, inv)
            if _p and _p['status'] == 'LUNAS':
                conn.rollback(); conn.close()
                flash(f'Invoice {inv["nomor"]} masih Lunas karena ada pembayaran yang tercatat atas '
                      'tagihannya. Untuk mengubah statusnya, pembayaran itu perlu dibatalkan dulu. '
                      'Rinciannya ada di halaman ini.', 'info')
                return redirect(url_for('transaksi_buka_kunci', id=inv['jurnal_id']))
        conn.execute("UPDATE invoice SET status=? WHERE id=?", (new_status, id))
        conn.commit()
    except ValueError as ex:
        conn.rollback(); conn.close()
        flash(f'Gagal memperbarui status invoice: {ex}', 'danger')
        return redirect(url_for('invoice_view', id=id))
    conn.close()
    return redirect(url_for('invoice_view', id=id))

@app.route('/invoice/<int:id>/post-jurnal', methods=['POST'])
@finance_required
def invoice_post_jurnal(id):
    """Catat invoice manual PKP ke jurnal (otomatis): Dr Piutang, Cr Pendapatan/Hutang PPN/PPh Dimuka."""
    conn = db()
    try:
        jid = _post_invoice_to_jurnal(conn, id)
        add_log(conn, 'Invoice di-post ke Jurnal', f'Inv #{id} -> Jurnal #{jid}', 'INPUT')
        conn.commit()
        flash(f'Invoice berhasil dicatat ke jurnal (#{jid}). Saldo Piutang & Hutang PPN sudah diupdate.', 'success')
    except ValueError as ex:
        conn.rollback()
        flash(str(ex), 'danger')
    finally:
        conn.close()
    return redirect(url_for('invoice_view', id=id))

@app.route('/invoice/<int:id>/hapus', methods=['POST'])
@finance_required
def invoice_hapus(id):
    conn = db()
    inv = conn.execute("SELECT * FROM invoice WHERE id=?", (id,)).fetchone()
    if not inv:
        conn.close()
        flash('Invoice tidak ditemukan.', 'danger')
        return redirect(url_for('invoice_list'))

    try:
        j = conn.execute("SELECT * FROM jurnal WHERE id=?", (inv['jurnal_id'],)).fetchone() if inv['jurnal_id'] else None
        has_source_meta = bool(j and ((j['tx_meta'] if 'tx_meta' in j.keys() else '') or ''))
        if has_source_meta:
            if inv['status'] != 'DRAFT':
                conn.close()
                flash('Invoice dari transaksi sumber sudah tidak Draft. Kembalikan ke Draft sebelum menghapus dokumen invoice.', 'warning')
                return redirect(url_for('invoice_view', id=id))
            conn.execute("DELETE FROM invoice WHERE id=?", (id,))
            flash('Invoice draft dari transaksi sumber dihapus. Transaksi penjualan tetap aman.', 'info')
        else:
            if inv['jurnal_id']:
                if transaction_dependency_count(conn, inv['jurnal_id']):
                    conn.close()
                    flash('Invoice ini sudah memiliki pembayaran atau retur terkait. Hapus jurnal terkait lebih dulu.', 'warning')
                    return redirect(url_for('invoice_view', id=id))
                _reverse_tx(conn, inv['jurnal_id'])
                conn.execute("DELETE FROM jurnal WHERE id=?", (inv['jurnal_id'],))
            conn.execute("DELETE FROM invoice WHERE id=?", (id,))
            flash('Invoice dihapus dan jurnal/piutang terkait dibatalkan.', 'info')
        conn.commit()
    except ValueError as ex:
        conn.rollback()
        flash(str(ex), 'danger')
    conn.close()
    return redirect(url_for('invoice_list'))


# ─────────────────────────────────────────────────────────────
# PURCHASE ORDER
# ─────────────────────────────────────────────────────────────

@app.route('/po')
@operator_required
def po_list():
    conn = db()
    q  = request.args.get('q', '')
    sf = request.args.get('status', '')
    where = " WHERE 1=1"
    params = []
    if q:
        where += " AND (vendor LIKE ? OR nomor LIKE ?)"
        params += [f'%{q}%', f'%{q}%']
    if sf:
        where += " AND status=?"
        params.append(sf)
    page = _page_arg()
    n_rows = conn.execute("SELECT COUNT(*) FROM purchase_order" + where, params).fetchone()[0]
    pg = paginate_meta(page, n_rows)
    pos = conn.execute(
        "SELECT * FROM purchase_order" + where + " ORDER BY tanggal DESC, id DESC LIMIT ? OFFSET ?",
        params + [pg['per_page'], pg['offset']]).fetchall()
    # hitung total per PO
    po_data = []
    for p in pos:
        items = conn.execute("SELECT subtotal FROM po_item WHERE po_id=?", (p['id'],)).fetchall()
        subtotal = sum(it['subtotal'] for it in items)
        total = subtotal - float(p['diskon'] or 0) + float(p['ongkir'] or 0) + float(p['biaya_lain'] or 0)
        po_data.append({'po': p, 'total': total})
    conn.close()
    return render_template('po_list.html', po_data=po_data, q=q, sf=sf, pg=pg)

@app.route('/po/pengaturan', methods=['GET','POST'])
@admin_required
def po_pengaturan():
    conn = db()
    if request.method == 'POST':
        for k in ['po_nama','po_alamat','po_telepon','po_email',
                  'po_prefix','po_tagline','po_terms','po_catatan']:
            conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                         (k, request.form.get(k, '')))
        logo_file = request.files.get('po_logo_file')
        if logo_file and logo_file.filename:
            ext = logo_file.filename.rsplit('.', 1)[-1].lower()
            if ext in ALLOWED_IMG:
                for f in os.listdir(UPLOAD_FOLDER):
                    if f.startswith('po_logo.'):
                        try: os.remove(os.path.join(UPLOAD_FOLDER, f))
                        except: pass
                logo_fn = f'po_logo.{ext}'
                logo_file.save(os.path.join(UPLOAD_FOLDER, logo_fn))
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                             ('po_logo', logo_fn))
        elif request.form.get('po_logo_delete') == '1':
            old = get_setting(conn, 'po_logo', '')
            if old:
                try: os.remove(os.path.join(UPLOAD_FOLDER, old))
                except: pass
            conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", ('po_logo', ''))
        conn.commit(); conn.close()
        flash('Template PO disimpan!', 'success')
        return redirect(url_for('po_pengaturan'))
    po_s = _po_settings(conn)
    conn.close()
    return render_template('po_pengaturan.html', po_s=po_s)

@app.route('/po/preview')
@operator_required
def po_preview():
    conn = db()
    po_s = _po_settings(conn)
    conn.close()
    po = {
        'nomor': f"{po_s['po_prefix']}-{date.today().strftime('%Y%m')}-001",
        'tanggal': date.today().strftime('%Y-%m-%d'),
        'expected_date': (date.today() + timedelta(days=7)).strftime('%Y-%m-%d'),
        'vendor': 'Contoh Vendor / PT Supplier Jaya',
        'alamat_vendor': 'Jl. Vendor No. 99, Surabaya',
        'telepon_vendor': '031-xxx-xxxx',
        'diskon': 0, 'ongkir': 25000, 'biaya_lain': 0,
        'catatan': '', 'status': 'DRAFT',
    }
    items = [
        {'deskripsi': 'Bahan Baku / Produk Contoh A', 'qty': 10, 'satuan': 'pcs',
         'harga_satuan': 150000, 'diskon_item': 0, 'subtotal': 1500000},
        {'deskripsi': 'Bahan Baku / Produk Contoh B', 'qty': 5, 'satuan': 'kg',
         'harga_satuan': 80000, 'diskon_item': 0, 'subtotal': 400000},
    ]
    subtotal = sum(it['subtotal'] for it in items)
    total = subtotal - float(po['diskon']) + float(po['ongkir']) + float(po['biaya_lain'])
    return render_template('po_print.html', po=po, items=items, po_s=po_s,
                           subtotal=subtotal, total=total, autoprint=False, preview=True)

@app.route('/po/baru', methods=['GET','POST'])
@operator_required
def po_baru():
    conn = db()
    po_s = _po_settings(conn)
    produk_list = _po_produk_list(conn)
    vendor_list = [dict(r) for r in conn.execute(
        "SELECT nama,telepon,alamat,catatan FROM vendor ORDER BY nama"
    ).fetchall()]
    if request.method == 'POST':
        prefix = po_s['po_prefix']
        nomor  = request.form.get('nomor','').strip() or _next_po_number(conn, prefix)
        if conn.execute("SELECT id FROM purchase_order WHERE nomor=?", (nomor,)).fetchone():
            flash(f'Nomor PO {nomor} sudah digunakan.', 'danger')
            conn.close()
            return redirect(url_for('po_baru'))
        tanggal = request.form.get('tanggal','')
        if not is_valid_iso_date(tanggal):
            conn.close(); flash('Tanggal PO tidak valid.', 'danger')
            return redirect(url_for('po_baru'))
        expected = request.form.get('expected_date','').strip() or None
        if expected and not is_valid_iso_date(expected):
            conn.close(); flash('Tanggal pengiriman tidak valid.', 'danger')
            return redirect(url_for('po_baru'))
        vendor = request.form.get('vendor','').strip()
        if not vendor:
            conn.close(); flash('Nama vendor wajib diisi.', 'danger')
            return redirect(url_for('po_baru'))
        po_id = conn.execute(
            "INSERT INTO purchase_order(nomor,tanggal,expected_date,vendor,alamat_vendor,telepon_vendor,diskon,ongkir,biaya_lain,catatan) VALUES(?,?,?,?,?,?,?,?,?,?)",
            (nomor, tanggal, expected,
             vendor,
             request.form.get('alamat_vendor','').strip(),
             request.form.get('telepon_vendor','').strip(),
             parse_rp(request.form.get('diskon','0')),
             parse_rp(request.form.get('ongkir','0')),
             parse_rp(request.form.get('biaya_lain','0')),
             request.form.get('catatan','').strip())
        ).lastrowid
        try:
            _save_po_items(conn, po_id, request.form)
        except ValueError as ex:
            conn.rollback(); conn.close()
            flash(str(ex), 'danger')
            return redirect(url_for('po_baru'))
        add_log(conn, 'PO dibuat', f"{nomor} | {request.form.get('vendor','')}", 'INPUT')
        conn.commit(); conn.close()
        flash(f'Purchase Order {nomor} berhasil dibuat!', 'success')
        return redirect(url_for('po_view', id=po_id))
    next_num = _next_po_number(conn, po_s['po_prefix'])
    conn.close()
    return render_template('po_form.html', po_s=po_s, produk_list=produk_list,
                           vendor_list=vendor_list,
                           next_num=next_num, today=date.today().strftime('%Y-%m-%d'),
                           edit=None, items=[])

@app.route('/po/<int:id>', methods=['GET'])
@operator_required
def po_view(id):
    conn = db()
    po = conn.execute("SELECT * FROM purchase_order WHERE id=?", (id,)).fetchone()
    if not po:
        conn.close(); flash('PO tidak ditemukan.', 'danger')
        return redirect(url_for('po_list'))
    items = conn.execute("SELECT * FROM po_item WHERE po_id=? ORDER BY id", (id,)).fetchall()
    po_s  = _po_settings(conn)
    akun_kas = conn.execute(
        "SELECT kode,nama FROM akun WHERE is_rekening=1 ORDER BY kode"
    ).fetchall()
    hutang = conn.execute("SELECT * FROM hutang WHERE id=?", (po['hutang_id'],)).fetchone() if po['hutang_id'] else None
    conn.close()
    subtotal = sum(it['subtotal'] for it in items)
    total    = subtotal - float(po['diskon'] or 0) + float(po['ongkir'] or 0) + float(po['biaya_lain'] or 0)
    return render_template('po_view.html', po=po, items=items, po_s=po_s,
                           subtotal=subtotal, total=total, today=date.today(),
                           akun_kas=akun_kas, hutang=hutang)

@app.route('/po/<int:id>/edit', methods=['GET','POST'])
@operator_required
def po_edit(id):
    conn = db()
    po = conn.execute("SELECT * FROM purchase_order WHERE id=?", (id,)).fetchone()
    if not po:
        conn.close(); flash('PO tidak ditemukan.', 'danger')
        return redirect(url_for('po_list'))
    if po['status'] in ('DITERIMA','DIBATALKAN') or po['hutang_id'] or po['jurnal_id']:
        conn.close(); flash('PO yang sudah diterima/dibatalkan tidak bisa diedit langsung. Ubah status dulu agar jurnal, hutang, dan stok dibalik dengan rapi.', 'danger')
        return redirect(url_for('po_view', id=id))
    po_s = _po_settings(conn)
    produk_list = _po_produk_list(conn, id)
    vendor_list = [dict(r) for r in conn.execute(
        "SELECT nama,telepon,alamat,catatan FROM vendor ORDER BY nama"
    ).fetchall()]
    if request.method == 'POST':
        nomor_e = request.form.get('nomor', po['nomor']).strip() or po['nomor']
        if nomor_e != po['nomor'] and conn.execute(
                "SELECT id FROM purchase_order WHERE nomor=? AND id!=?", (nomor_e, id)).fetchone():
            conn.close(); flash(f'Nomor PO {nomor_e} sudah digunakan.', 'danger')
            return redirect(url_for('po_edit', id=id))
        tanggal_e = request.form.get('tanggal', po['tanggal'])
        if not is_valid_iso_date(tanggal_e):
            conn.close(); flash('Tanggal PO tidak valid.', 'danger')
            return redirect(url_for('po_edit', id=id))
        expected_e = request.form.get('expected_date','').strip() or None
        if expected_e and not is_valid_iso_date(expected_e):
            conn.close(); flash('Tanggal pengiriman tidak valid.', 'danger')
            return redirect(url_for('po_edit', id=id))
        vendor_e = request.form.get('vendor','').strip()
        if not vendor_e:
            conn.close(); flash('Nama vendor wajib diisi.', 'danger')
            return redirect(url_for('po_edit', id=id))
        conn.execute("""UPDATE purchase_order SET nomor=?,tanggal=?,expected_date=?,
                        vendor=?,alamat_vendor=?,telepon_vendor=?,
                        diskon=?,ongkir=?,biaya_lain=?,catatan=? WHERE id=?""",
                     (nomor_e, tanggal_e, expected_e,
                      vendor_e,
                      request.form.get('alamat_vendor','').strip(),
                      request.form.get('telepon_vendor','').strip(),
                      parse_rp(request.form.get('diskon','0')),
                      parse_rp(request.form.get('ongkir','0')),
                      parse_rp(request.form.get('biaya_lain','0')),
                      request.form.get('catatan','').strip(), id))
        try:
            _save_po_items(conn, id, request.form)
        except ValueError as ex:
            conn.rollback(); conn.close()
            flash(str(ex), 'danger')
            return redirect(url_for('po_edit', id=id))
        conn.commit(); conn.close()
        flash('PO diperbarui!', 'success')
        return redirect(url_for('po_view', id=id))
    items = conn.execute("SELECT * FROM po_item WHERE po_id=? ORDER BY id", (id,)).fetchall()
    conn.close()
    return render_template('po_form.html', po_s=po_s, produk_list=produk_list,
                           vendor_list=vendor_list,
                           today=date.today().strftime('%Y-%m-%d'), edit=po, items=items)

@app.route('/po/<int:id>/cetak')
@operator_required
def po_cetak(id):
    conn = db()
    po = conn.execute("SELECT * FROM purchase_order WHERE id=?", (id,)).fetchone()
    if not po:
        conn.close(); return redirect(url_for('po_list'))
    items = conn.execute("SELECT * FROM po_item WHERE po_id=? ORDER BY id", (id,)).fetchall()
    po_s  = _po_settings(conn)
    conn.close()
    subtotal = sum(it['subtotal'] for it in items)
    total    = subtotal - float(po['diskon'] or 0) + float(po['ongkir'] or 0) + float(po['biaya_lain'] or 0)
    autoprint = request.args.get('autoprint','0') == '1'
    return render_template('po_print.html', po=po, items=items, po_s=po_s,
                           subtotal=subtotal, total=total, autoprint=autoprint, preview=False)

def _po_total(conn, po):
    items = conn.execute("SELECT subtotal FROM po_item WHERE po_id=?", (po['id'],)).fetchall()
    subtotal = sum(float(it['subtotal'] or 0) for it in items)
    total = subtotal - float(po['diskon'] or 0) + float(po['ongkir'] or 0) + float(po['biaya_lain'] or 0)
    return round(max(0.0, total), 2)

def _po_item_allocations(conn, po):
    items = conn.execute("SELECT * FROM po_item WHERE po_id=? ORDER BY id", (po['id'],)).fetchall()
    if not items:
        raise ValueError('PO belum memiliki item.')
    subtotal = round(sum(float(it['subtotal'] or 0) for it in items), 2)
    total = _po_total(conn, po)
    if subtotal <= 0:
        raise ValueError('Subtotal item PO harus lebih dari 0 sebelum barang diterima.')
    if total <= 0:
        raise ValueError('Total PO harus lebih dari 0 sebelum barang diterima.')

    allocated = 0.0
    rows = []
    for idx, it in enumerate(items):
        qty = float(it['qty'] or 0)
        if qty <= 0:
            raise ValueError(f'Qty item PO "{it["deskripsi"]}" harus lebih dari 0.')
        if idx == len(items) - 1:
            landed = round(total - allocated, 2)
        else:
            landed = round(float(it['subtotal'] or 0) * total / subtotal, 2)
            allocated = round(allocated + landed, 2)
        landed = max(0.0, landed)
        rows.append({
            'item': it,
            'landed_subtotal': landed,
            'unit_cost': round(landed / qty, 6) if qty else 0.0,
        })
    return subtotal, total, rows

def _po_hutang_has_payments(conn, hutang_id):
    if not hutang_id:
        return False
    h = conn.execute("SELECT terbayar FROM hutang WHERE id=?", (hutang_id,)).fetchone()
    if h and float(h['terbayar'] or 0) > 0.01:
        return True
    return bool(conn.execute("SELECT 1 FROM bayar_hutang WHERE hutang_id=? LIMIT 1", (hutang_id,)).fetchone())

def _po_remove_accounting(conn, po):
    if po['hutang_id'] and _po_hutang_has_payments(conn, po['hutang_id']):
        raise ValueError('PO ini sudah memiliki pembayaran hutang. Hapus pembayaran hutangnya lebih dulu.')

    jid = po['jurnal_id']
    if jid and conn.execute("SELECT 1 FROM jurnal WHERE id=?", (jid,)).fetchone():
        _reverse_tx(conn, jid)
        conn.execute("DELETE FROM jurnal WHERE id=?", (jid,))
    if po['hutang_id']:
        conn.execute("DELETE FROM hutang WHERE id=? AND COALESCE(terbayar,0)=0", (po['hutang_id'],))

    conn.execute("UPDATE purchase_order SET hutang_id=NULL,jurnal_id=NULL WHERE id=?", (po['id'],))

def _po_receive_form(conn, po, form=None):
    form = form or {}
    tanggal = (form.get('tanggal_terima') or date.today().strftime('%Y-%m-%d')).strip()
    if not is_valid_iso_date(tanggal):
        raise ValueError('Tanggal terima PO tidak valid.')
    if po['tanggal'] and tanggal < po['tanggal']:
        raise ValueError('Tanggal terima tidak boleh lebih awal dari tanggal PO.')
    if not _operator_date_ok(tanggal):
        raise ValueError('Operator hanya boleh mencatat penerimaan pada bulan berjalan.')

    jatuh_tempo = (form.get('jatuh_tempo') or po['expected_date'] or '').strip() or None
    if jatuh_tempo and not is_valid_iso_date(jatuh_tempo):
        raise ValueError('Tanggal jatuh tempo hutang tidak valid.')
    if jatuh_tempo and jatuh_tempo < tanggal:
        raise ValueError('Tanggal jatuh tempo tidak boleh lebih awal dari tanggal terima.')

    uang_keluar = round(parse_nonnegative_rp(form.get('uang_keluar', '0'), 'Uang dibayar saat terima'), 2)
    akun_kas = (form.get('akun_kas') or '1100').strip()
    if uang_keluar > 0 and not is_rekening_kode(conn, akun_kas):
        raise ValueError('Akun kas/bank pembayaran PO tidak valid.')
    return {'tanggal': tanggal, 'jatuh_tempo': jatuh_tempo, 'uang_keluar': uang_keluar, 'akun_kas': akun_kas}

def _po_ensure_accounting(conn, po, form=None):
    _, total, allocations = _po_item_allocations(conn, po)
    receive = _po_receive_form(conn, po, form)
    uang_keluar = receive['uang_keluar']
    if uang_keluar > total + 0.01:
        raise ValueError('Uang dibayar saat terima tidak boleh melebihi total PO.')
    if abs(uang_keluar - total) <= 0.01:
        uang_keluar = total
    hutang_nilai = round(max(0.0, total - uang_keluar), 2)

    sku_count = 0
    non_sku_count = 0
    for alloc in allocations:
        it = alloc['item']
        if it['produk_id']:
            prod = conn.execute(
                "SELECT id FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
                (it['produk_id'],)
            ).fetchone()
            if not prod:
                raise ValueError(f'Produk "{it["deskripsi"]}" sudah tidak aktif/tidak ditemukan. Aktifkan produk atau ubah item PO sebelum menerima barang.')

    if po['hutang_id'] or po['jurnal_id']:
        _po_remove_accounting(conn, po)

    entries = [('1130', total, 0)]
    if uang_keluar > 0:
        entries.append((receive['akun_kas'], 0, uang_keluar))
    if hutang_nilai > 0:
        entries.append(('2100', 0, hutang_nilai))

    jid = insert_jurnal(
        conn, receive['tanggal'], f"Terima PO {po['nomor']} - {po['vendor']}", 'OPERASIONAL', 'PO_DITERIMA',
        entries, pihak=po['vendor']
    )
    conn.execute("UPDATE jurnal SET tx_meta=? WHERE id=?", (json.dumps({
        'tipe': 'PO',
        'po_id': po['id'],
        'nomor': po['nomor'],
        'pihak': po['vendor'],
        'tanggal_terima': receive['tanggal'],
        'akun_kas': receive['akun_kas'] if uang_keluar > 0 else '',
        'uang_keluar': uang_keluar,
        'hutang': hutang_nilai,
        'jatuh_tempo': receive['jatuh_tempo'],
    }), jid))

    for alloc in allocations:
        it = alloc['item']
        qty = float(it['qty'] or 0)
        if it['produk_id']:
            p = conn.execute(
                "SELECT id,nama,stok,harga_beli FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
                (it['produk_id'],)
            ).fetchone()
            stok_sebelum = float(p['stok'] or 0)
            hpp_sebelum = float(p['harga_beli'] or 0)
            update_average_cost(conn, p['id'], qty, alloc['unit_cost'])
            conn.execute("UPDATE produk SET stok=stok+? WHERE id=?", (qty, p['id']))
            record_stock_movement(
                conn, p['id'], receive['tanggal'], 'MASUK', qty, alloc['unit_cost'],
                f"PO diterima: {po['nomor']} - {it['deskripsi']}", jid,
                'ADD_LAYER', stok_sebelum, hpp_sebelum
            )
            sku_count += 1
        else:
            record_non_sku_movement(
                conn, jid, receive['tanggal'],
                f"PO diterima: {po['nomor']} - {it['deskripsi']}", alloc['landed_subtotal']
            )
            non_sku_count += 1

    hutang_id = None
    if hutang_nilai > 0:
        hutang_id = conn.execute(
            """INSERT INTO hutang(tanggal,jatuh_tempo,pemasok,keterangan,jumlah,jurnal_id,akun_kode)
               VALUES(?,?,?,?,?,?,?)""",
            (receive['tanggal'], receive['jatuh_tempo'], po['vendor'],
             f"Pembelian PO {po['nomor']}", hutang_nilai, jid, '2100')
        ).lastrowid
    conn.execute("UPDATE purchase_order SET hutang_id=?,jurnal_id=? WHERE id=?", (hutang_id, jid, po['id']))
    return {
        'total': total,
        'uang_keluar': uang_keluar,
        'hutang': hutang_nilai,
        'jurnal_id': jid,
        'hutang_id': hutang_id,
        'sku_count': sku_count,
        'non_sku_count': non_sku_count,
    }

def _fmt_rupiah_plain(value):
    return 'Rp ' + f"{float(value or 0):,.0f}".replace(',', '.')

@app.route('/po/<int:id>/status', methods=['POST'])
@operator_required
def po_status_update(id):
    new_status = request.form.get('status','DIKIRIM')
    if new_status not in ('DRAFT','DIKIRIM','DITERIMA','DIBATALKAN'):
        flash('Status PO tidak valid.', 'danger')
        return redirect(url_for('po_view', id=id))
    conn = db()
    po = conn.execute("SELECT * FROM purchase_order WHERE id=?", (id,)).fetchone()
    if not po:
        conn.close(); flash('PO tidak ditemukan.', 'danger')
        return redirect(url_for('po_list'))
    try:
        if new_status == 'DITERIMA':
            if po['status'] == 'DITERIMA' and not any(k in request.form for k in ('tanggal_terima', 'uang_keluar', 'akun_kas', 'jatuh_tempo')):
                flash('Status PO tetap DITERIMA.', 'info')
            else:
                hasil = _po_ensure_accounting(conn, po, request.form)
                if hasil['uang_keluar'] > 0 and hasil['hutang'] > 0:
                    msg = (f'Barang diterima. Persediaan/stok dicatat sebesar {_fmt_rupiah_plain(hasil["total"])}, '
                           f'kas/bank keluar {_fmt_rupiah_plain(hasil["uang_keluar"])}, '
                           f'dan sisa hutang {_fmt_rupiah_plain(hasil["hutang"])} ke {po["vendor"]}.')
                elif hasil['uang_keluar'] > 0:
                    msg = (f'Barang diterima dan dibayar penuh. Persediaan/stok dicatat sebesar '
                           f'{_fmt_rupiah_plain(hasil["total"])}.')
                else:
                    msg = (f'Barang diterima. Persediaan/stok dicatat sebesar {_fmt_rupiah_plain(hasil["total"])} '
                           f'dan hutang {_fmt_rupiah_plain(hasil["hutang"])} ke {po["vendor"]} dibuat.')
                flash(msg, 'success')
        elif po['status'] == 'DITERIMA' or po['hutang_id'] or po['jurnal_id']:
            _po_remove_accounting(conn, po)
            flash(f'Status PO diubah ke {new_status}. Jurnal dan hutang PO dibatalkan.', 'info')
        else:
            flash(f'Status PO diubah ke {new_status}.', 'info')
        conn.execute("UPDATE purchase_order SET status=? WHERE id=?", (new_status, id))
        conn.commit()
    except ValueError as ex:
        conn.rollback()
        flash(str(ex), 'danger')
    conn.close()
    return redirect(url_for('po_view', id=id))

@app.route('/po/<int:id>/hapus', methods=['POST'])
@finance_required
def po_hapus(id):
    conn = db()
    po = conn.execute("SELECT * FROM purchase_order WHERE id=?", (id,)).fetchone()
    if not po:
        conn.close()
        flash('PO tidak ditemukan.', 'danger')
        return redirect(url_for('po_list'))
    try:
        if po['hutang_id'] or po['jurnal_id']:
            _po_remove_accounting(conn, po)
        conn.execute("DELETE FROM purchase_order WHERE id=?", (id,))
        conn.commit()
        flash('PO dihapus.', 'info')
    except ValueError as ex:
        conn.rollback()
        flash(str(ex), 'danger')
    conn.close()
    return redirect(url_for('po_list'))

# ---------- LAPORAN ----------
@app.route('/neraca')
@investor_required
def neraca():
    conn = db()
    ed = request.args.get('ed', date.today().strftime('%Y-%m-%d'))
    rows = conn.execute("""
        SELECT a.kode, a.nama, a.tipe, a.subtipe, a.saldo_normal,
               COALESCE(SUM(CASE WHEN j.tanggal <= ? THEN d.debit ELSE 0 END),0) as td,
               COALESCE(SUM(CASE WHEN j.tanggal <= ? THEN d.kredit ELSE 0 END),0) as tk
        FROM akun a
        LEFT JOIN detail_jurnal d ON d.akun_id = a.id
        LEFT JOIN jurnal j ON j.id = d.jurnal_id
        GROUP BY a.id ORDER BY a.kode
    """, (ed, ed)).fetchall()
    aset={};liab={};ekuitas_list=[]
    total_aset=total_liab=total_ekuitas=0
    for r in rows:
        saldo = (r['td']-r['tk']) if r['saldo_normal']=='DEBIT' else (r['tk']-r['td'])
        item = dict(kode=r['kode'],nama=r['nama'],saldo=saldo)
        if r['tipe']=='ASET':
            # Contra-asset (saldo normal KREDIT, mis. Akumulasi Penyusutan) mengurangi aset
            aset_saldo = -saldo if r['saldo_normal'] == 'KREDIT' else saldo
            aset.setdefault(r['subtipe'] or 'Aset Lancar',[]).append(
                dict(kode=r['kode'], nama=r['nama'], saldo=aset_saldo))
            total_aset += aset_saldo
        elif r['tipe']=='LIABILITAS':
            liab.setdefault(r['subtipe'] or 'Liabilitas Lancar',[]).append(item)
            total_liab += saldo
        elif r['tipe']=='EKUITAS':
            # Akun ekuitas berdebet (contoh: 3300 Prive) MENGURANGI ekuitas
            ekuitas_saldo = -saldo if r['saldo_normal'] == 'DEBIT' else saldo
            ekuitas_list.append(dict(kode=r['kode'], nama=r['nama'], saldo=ekuitas_saldo))
            total_ekuitas += ekuitas_saldo
    fiscal_start  = f"{ed[:4]}-01-01"
    laba_tahun    = calc_profitability(conn, fiscal_start, ed)['laba_bersih']
    laba_all      = calc_profitability(conn, '2000-01-01', ed)['laba_bersih']
    laba_ditahan  = laba_all - laba_tahun
    if abs(laba_ditahan) > 0.01:
        ekuitas_list.append(dict(kode='-', nama='Laba Ditahan (Akumulasi)', saldo=laba_ditahan))
        total_ekuitas += laba_ditahan
    ekuitas_list.append(dict(kode='-', nama=f'Laba Bersih Tahun {ed[:4]}', saldo=laba_tahun))
    total_ekuitas += laba_tahun

    # Diagnostik 1: jurnal yang tidak seimbang (debit ≠ kredit)
    unbalanced = conn.execute("""
        SELECT j.id, j.tanggal, j.keterangan, j.tipe_tx,
               ROUND(SUM(dj.debit),2)  AS total_debit,
               ROUND(SUM(dj.kredit),2) AS total_kredit,
               ROUND(SUM(dj.debit) - SUM(dj.kredit), 2) AS selisih
        FROM jurnal j
        JOIN detail_jurnal dj ON dj.jurnal_id = j.id
        GROUP BY j.id
        HAVING ABS(SUM(dj.debit) - SUM(dj.kredit)) > 0.01
        ORDER BY ABS(SUM(dj.debit) - SUM(dj.kredit)) DESC
        LIMIT 20
    """).fetchall()
    # Diagnostik 2: entri jurnal yang merujuk akun yang sudah dihapus (yatim)
    orphaned = conn.execute("""
        SELECT j.id, j.tanggal, j.keterangan, j.tipe_tx,
               ROUND(SUM(dj.debit),2)  AS orphan_debit,
               ROUND(SUM(dj.kredit),2) AS orphan_kredit,
               GROUP_CONCAT(dj.akun_id) AS akun_ids_hilang
        FROM detail_jurnal dj
        LEFT JOIN akun a ON a.id = dj.akun_id
        JOIN jurnal j ON j.id = dj.jurnal_id
        WHERE a.id IS NULL
        GROUP BY j.id
        ORDER BY j.tanggal DESC
        LIMIT 20
    """).fetchall()
    # Cek total global
    global_check = conn.execute(
        "SELECT ROUND(SUM(debit)-SUM(kredit),2) as diff FROM detail_jurnal"
    ).fetchone()
    global_diff = float(global_check['diff'] or 0)

    conn.close()
    return render_template('neraca.html',
        aset=aset,liab=liab,ekuitas_list=ekuitas_list,
        total_aset=total_aset,total_liab=total_liab,total_ekuitas=total_ekuitas,ed=ed,
        unbalanced=unbalanced, orphaned=orphaned, global_diff=global_diff)

def calc_equity_movement(conn, sd, ed):
    """
    Hitung perubahan ekuitas selama periode [sd, ed].
    Mengembalikan dict berisi:
      - rows: list per akun ekuitas {kode, nama, saldo_normal, awal, tambah, kurang, akhir}
      - laba_sebelum: akumulasi laba ditahan sampai tanggal sebelum sd
      - laba_periode: laba bersih periode [sd, ed]
      - total_awal, total_tambah, total_kurang, total_akhir
    """
    rows = conn.execute("""
        SELECT a.id, a.kode, a.nama, a.saldo_normal,
               COALESCE(SUM(CASE WHEN j.tanggal <  ? THEN dj.debit  ELSE 0 END),0) AS d_pre,
               COALESCE(SUM(CASE WHEN j.tanggal <  ? THEN dj.kredit ELSE 0 END),0) AS k_pre,
               COALESCE(SUM(CASE WHEN j.tanggal >= ? AND j.tanggal <= ? THEN dj.debit  ELSE 0 END),0) AS d_per,
               COALESCE(SUM(CASE WHEN j.tanggal >= ? AND j.tanggal <= ? THEN dj.kredit ELSE 0 END),0) AS k_per
        FROM akun a
        LEFT JOIN detail_jurnal dj ON dj.akun_id = a.id
        LEFT JOIN jurnal j ON j.id = dj.jurnal_id
        WHERE a.tipe = 'EKUITAS'
        GROUP BY a.id ORDER BY a.kode
    """, (sd, sd, sd, ed, sd, ed)).fetchall()
    out_rows = []
    total_awal = total_tambah = total_kurang = total_akhir = 0.0
    for r in rows:
        d_pre, k_pre = float(r['d_pre']), float(r['k_pre'])
        d_per, k_per = float(r['d_per']), float(r['k_per'])
        if r['saldo_normal'] == 'KREDIT':
            awal   = k_pre - d_pre
            tambah = k_per
            kurang = d_per
        else:  # DEBIT (contoh: Prive 3300), saldo positif berarti pengurang ekuitas
            # Tampilkan sebagai pengurang: nilai disajikan dengan tanda minus
            awal   = -(d_pre - k_pre)
            tambah = -d_per  # debit di akun prive = pengurang ekuitas
            kurang = -k_per  # kredit di akun prive = penambah ekuitas (jarang, reversal)
        akhir = awal + tambah - kurang
        out_rows.append(dict(
            kode=r['kode'], nama=r['nama'], saldo_normal=r['saldo_normal'],
            awal=awal, tambah=tambah, kurang=kurang, akhir=akhir
        ))
        total_awal   += awal
        total_tambah += tambah
        total_kurang += kurang
        total_akhir  += akhir
    # Akumulasi laba ditahan sampai akhir tanggal sebelum sd
    # (gunakan calc_profitability untuk semua tanggal sebelum sd)
    from datetime import datetime, timedelta
    sd_dt = datetime.strptime(sd, '%Y-%m-%d').date()
    day_before = (sd_dt - timedelta(days=1)).strftime('%Y-%m-%d')
    laba_sebelum = calc_profitability(conn, '2000-01-01', day_before)['laba_bersih'] if day_before >= '2000-01-01' else 0.0
    laba_periode = calc_profitability(conn, sd, ed)['laba_bersih']
    # Tambahkan laba ditahan akumulasi (sebelum sd) sebagai baris virtual di saldo awal
    total_awal  += laba_sebelum
    total_akhir += laba_sebelum + laba_periode
    return dict(
        rows=out_rows,
        laba_sebelum=laba_sebelum, laba_periode=laba_periode,
        total_awal=total_awal, total_tambah=total_tambah,
        total_kurang=total_kurang, total_akhir=total_akhir,
    )

@app.route('/perubahan-ekuitas')
@investor_required
def perubahan_ekuitas():
    conn = db()
    today = date.today()
    sd = request.args.get('sd', f"{today.year}-01-01")
    ed = request.args.get('ed', today.strftime('%Y-%m-%d'))
    data = calc_equity_movement(conn, sd, ed)
    conn.close()
    return render_template('perubahan_ekuitas.html', sd=sd, ed=ed, data=data, today=today)

@app.route('/analisis-rasio')
@investor_required
def analisis_rasio():
    conn = db()
    today = date.today()
    sd = request.args.get('sd', f"{today.year}-01-01")
    ed = request.args.get('ed', today.strftime('%Y-%m-%d'))
    cur = calc_financial_ratios(conn, sd, ed)
    # Periode pembanding: rentang yang sama panjangnya, berakhir sehari sebelum sd
    try:
        sd_dt = datetime.strptime(sd, '%Y-%m-%d').date()
        ed_dt = datetime.strptime(ed, '%Y-%m-%d').date()
        span_days = (ed_dt - sd_dt).days
        prev_ed_dt = sd_dt - timedelta(days=1)
        prev_sd_dt = prev_ed_dt - timedelta(days=span_days)
        prev_sd = prev_sd_dt.strftime('%Y-%m-%d')
        prev_ed = prev_ed_dt.strftime('%Y-%m-%d')
        prev = calc_financial_ratios(conn, prev_sd, prev_ed)
    except Exception:
        prev_sd = prev_ed = ''
        prev = None
    conn.close()
    return render_template('analisis_rasio.html',
        sd=sd, ed=ed, cur=cur, prev=prev,
        prev_sd=prev_sd, prev_ed=prev_ed, today=today)

@app.route('/laba-rugi')
@investor_required
def laba_rugi():
    conn = db()
    today = date.today()
    sd = request.args.get('sd', today.replace(day=1).strftime('%Y-%m-%d'))
    ed = request.args.get('ed', today.strftime('%Y-%m-%d'))
    pnl = calc_profitability(conn, sd, ed)
    rows = conn.execute("""
        SELECT a.kode, a.nama, a.tipe, a.subtipe,
               COALESCE(SUM(CASE WHEN j.tanggal >= ? AND j.tanggal <= ?
                                 THEN (CASE WHEN a.saldo_normal='DEBIT' THEN d.debit-d.kredit
                                            ELSE d.kredit-d.debit END) ELSE 0 END),0) as saldo
        FROM akun a
        LEFT JOIN detail_jurnal d ON d.akun_id = a.id
        LEFT JOIN jurnal j ON j.id = d.jurnal_id
        WHERE a.tipe IN ('PENDAPATAN','BEBAN')
        GROUP BY a.id ORDER BY a.kode
    """, (sd, ed)).fetchall()
    pendapatan={}; beban={}
    for r in rows:
        item = dict(kode=r['kode'],nama=r['nama'],saldo=r['saldo'])
        if r['tipe']=='PENDAPATAN':
            pendapatan.setdefault(r['subtipe'] or 'Pendapatan',[]).append(item)
        else:
            beban.setdefault(r['subtipe'] or 'Beban',[]).append(item)
    conn.close()
    return render_template('laba_rugi.html',
        pendapatan=pendapatan, beban=beban, pnl=pnl, sd=sd, ed=ed)

def _sales_performance_report(conn, sd, ed):
    product_rows = []
    rows = conn.execute("""
        WITH sale AS (
            SELECT ti.produk_id,
                   COALESCE(SUM(ti.qty),0) AS qty_sold,
                   COALESCE(SUM(ti.qty * ti.harga),0) AS gross_sales,
                   COALESCE(SUM(ti.diskon),0) AS item_discount,
                   COALESCE(SUM(ti.subtotal),0) AS sales_net,
                   COUNT(DISTINCT j.id) AS trx_count
            FROM transaksi_item ti
            JOIN jurnal j ON j.id = ti.jurnal_id
            WHERE j.tipe_tx='PEMASUKAN'
              AND ti.produk_id IS NOT NULL
              AND ti.arah='KELUAR'
              AND j.tanggal>=? AND j.tanggal<=?
            GROUP BY ti.produk_id
        ),
        retur AS (
            SELECT ti.produk_id,
                   COALESCE(SUM(ti.qty),0) AS qty_returned,
                   COALESCE(SUM(ti.subtotal),0) AS return_net,
                   COUNT(DISTINCT j.id) AS return_trx_count
            FROM transaksi_item ti
            JOIN jurnal j ON j.id = ti.jurnal_id
            WHERE j.tipe_tx='RETUR_JUAL'
              AND ti.produk_id IS NOT NULL
              AND ti.arah='MASUK'
              AND j.tanggal>=? AND j.tanggal<=?
            GROUP BY ti.produk_id
        ),
        hpp_sale AS (
            SELECT ps.produk_id,
                   COALESCE(SUM(ps.qty * ps.harga),0) AS hpp_sold
            FROM pergerakan_stok ps
            JOIN jurnal j ON j.id = ps.jurnal_id
            WHERE j.tipe_tx='PEMASUKAN'
              AND ps.jenis='KELUAR'
              AND ps.tanggal>=? AND ps.tanggal<=?
            GROUP BY ps.produk_id
        ),
        hpp_retur AS (
            SELECT ps.produk_id,
                   COALESCE(SUM(ps.qty * ps.harga),0) AS hpp_returned
            FROM pergerakan_stok ps
            JOIN jurnal j ON j.id = ps.jurnal_id
            WHERE j.tipe_tx='RETUR_JUAL'
              AND ps.jenis='MASUK'
              AND ps.tanggal>=? AND ps.tanggal<=?
            GROUP BY ps.produk_id
        )
        SELECT p.id, p.kode, p.nama, p.varian, p.satuan,
               COALESCE(NULLIF(TRIM(p.kategori),''), 'Tanpa Kategori') AS kategori,
               COALESCE(s.qty_sold,0) AS qty_sold,
               COALESCE(s.gross_sales,0) AS gross_sales,
               COALESCE(s.item_discount,0) AS item_discount,
               COALESCE(s.sales_net,0) AS sales_net,
               COALESCE(s.trx_count,0) AS trx_count,
               COALESCE(r.qty_returned,0) AS qty_returned,
               COALESCE(r.return_net,0) AS return_net,
               COALESCE(r.return_trx_count,0) AS return_trx_count,
               COALESCE(hs.hpp_sold,0) AS hpp_sold,
               COALESCE(hr.hpp_returned,0) AS hpp_returned
        FROM produk p
        LEFT JOIN sale s ON s.produk_id=p.id
        LEFT JOIN retur r ON r.produk_id=p.id
        LEFT JOIN hpp_sale hs ON hs.produk_id=p.id
        LEFT JOIN hpp_retur hr ON hr.produk_id=p.id
        WHERE COALESCE(s.qty_sold,0)<>0
           OR COALESCE(r.qty_returned,0)<>0
           OR COALESCE(s.sales_net,0)<>0
           OR COALESCE(r.return_net,0)<>0
        ORDER BY (COALESCE(s.sales_net,0)-COALESCE(r.return_net,0)) DESC, p.nama
    """, (sd, ed, sd, ed, sd, ed, sd, ed)).fetchall()

    for r in rows:
        item = dict(r)
        item['qty_net'] = float(item['qty_sold'] or 0) - float(item['qty_returned'] or 0)
        item['omzet_net'] = float(item['sales_net'] or 0) - float(item['return_net'] or 0)
        item['hpp_net'] = float(item['hpp_sold'] or 0) - float(item['hpp_returned'] or 0)
        item['margin'] = item['omzet_net'] - item['hpp_net']
        item['margin_pct'] = round((item['margin'] / item['omzet_net'] * 100), 1) if item['omzet_net'] else 0
        item['avg_price'] = item['omzet_net'] / item['qty_net'] if item['qty_net'] else 0
        item['return_rate'] = round((float(item['qty_returned'] or 0) / float(item['qty_sold'] or 1) * 100), 1) if item['qty_sold'] else 0
        product_rows.append(item)

    totals = {
        'qty_sold': sum(r['qty_sold'] for r in product_rows),
        'qty_returned': sum(r['qty_returned'] for r in product_rows),
        'qty_net': sum(r['qty_net'] for r in product_rows),
        'gross_sales': sum(r['gross_sales'] for r in product_rows),
        'item_discount': sum(r['item_discount'] for r in product_rows),
        'sales_net': sum(r['sales_net'] for r in product_rows),
        'return_net': sum(r['return_net'] for r in product_rows),
        'omzet_net': sum(r['omzet_net'] for r in product_rows),
        'hpp_net': sum(r['hpp_net'] for r in product_rows),
        'margin': sum(r['margin'] for r in product_rows),
        'trx_count': conn.execute(
            "SELECT COUNT(*) FROM jurnal WHERE tipe_tx='PEMASUKAN' AND tanggal>=? AND tanggal<=?",
            (sd, ed)
        ).fetchone()[0],
        'return_trx_count': conn.execute(
            "SELECT COUNT(*) FROM jurnal WHERE tipe_tx='RETUR_JUAL' AND tanggal>=? AND tanggal<=?",
            (sd, ed)
        ).fetchone()[0],
    }
    totals['margin_pct'] = round(totals['margin'] / totals['omzet_net'] * 100, 1) if totals['omzet_net'] else 0
    totals['discount_rate'] = round(totals['item_discount'] / totals['gross_sales'] * 100, 1) if totals['gross_sales'] else 0
    totals['return_rate'] = round(totals['return_net'] / totals['sales_net'] * 100, 1) if totals['sales_net'] else 0
    totals['avg_order'] = totals['sales_net'] / totals['trx_count'] if totals['trx_count'] else 0
    totals['avg_unit_price'] = totals['omzet_net'] / totals['qty_net'] if totals['qty_net'] else 0
    totals['active_products'] = len([r for r in product_rows if r['sales_net'] or r['return_net']])

    for r in product_rows:
        r['share'] = round(r['omzet_net'] / totals['omzet_net'] * 100, 1) if totals['omzet_net'] else 0

    category_map = {}
    for r in product_rows:
        c = category_map.setdefault(r['kategori'], {
            'kategori': r['kategori'], 'produk_count': 0, 'qty_net': 0.0,
            'omzet_net': 0.0, 'hpp_net': 0.0, 'margin': 0.0, 'return_net': 0.0
        })
        c['produk_count'] += 1
        c['qty_net'] += r['qty_net']
        c['omzet_net'] += r['omzet_net']
        c['hpp_net'] += r['hpp_net']
        c['margin'] += r['margin']
        c['return_net'] += r['return_net']
    category_rows = sorted(category_map.values(), key=lambda x: x['omzet_net'], reverse=True)
    for c in category_rows:
        c['margin_pct'] = round(c['margin'] / c['omzet_net'] * 100, 1) if c['omzet_net'] else 0
        c['share'] = round(c['omzet_net'] / totals['omzet_net'] * 100, 1) if totals['omzet_net'] else 0

    customer_rows = [dict(r) for r in conn.execute("""
        SELECT COALESCE(NULLIF(TRIM(j.pihak),''), 'Tanpa Nama') AS nama,
               COUNT(DISTINCT j.id) AS trx_count,
               COALESCE(SUM(CASE WHEN a.tipe='PENDAPATAN' THEN d.kredit-d.debit ELSE 0 END),0) AS omzet,
               COALESCE(SUM(CASE WHEN a.kode='5100' THEN d.debit-d.kredit ELSE 0 END),0) AS hpp
        FROM jurnal j
        JOIN detail_jurnal d ON d.jurnal_id=j.id
        JOIN akun a ON a.id=d.akun_id
        WHERE j.tipe_tx='PEMASUKAN'
          AND j.tanggal>=? AND j.tanggal<=?
        GROUP BY COALESCE(NULLIF(TRIM(j.pihak),''), 'Tanpa Nama')
        HAVING omzet<>0 OR hpp<>0
        ORDER BY omzet DESC
        LIMIT 10
    """, (sd, ed)).fetchall()]
    for c in customer_rows:
        c['margin'] = c['omzet'] - c['hpp']
        c['margin_pct'] = round(c['margin'] / c['omzet'] * 100, 1) if c['omzet'] else 0
        c['avg_order'] = c['omzet'] / c['trx_count'] if c['trx_count'] else 0

    trend_rows = [dict(r) for r in conn.execute("""
        SELECT substr(j.tanggal,1,7) AS periode,
               COALESCE(SUM(CASE WHEN j.tipe_tx='PEMASUKAN' AND ti.produk_id IS NOT NULL THEN ti.subtotal ELSE 0 END),0) AS omzet_produk,
               COALESCE(SUM(CASE WHEN j.tipe_tx='RETUR_JUAL' AND ti.produk_id IS NOT NULL THEN ti.subtotal ELSE 0 END),0) AS retur_produk,
               COUNT(DISTINCT CASE WHEN j.tipe_tx='PEMASUKAN' THEN j.id END) AS trx_count
        FROM jurnal j
        LEFT JOIN transaksi_item ti ON ti.jurnal_id=j.id
        WHERE j.tipe_tx IN ('PEMASUKAN','RETUR_JUAL')
          AND j.tanggal>=? AND j.tanggal<=?
        GROUP BY substr(j.tanggal,1,7)
        ORDER BY periode DESC
        LIMIT 12
    """, (sd, ed)).fetchall()]
    for t in trend_rows:
        t['omzet_net'] = t['omzet_produk'] - t['retur_produk']
        t['return_rate'] = round(t['retur_produk'] / t['omzet_produk'] * 100, 1) if t['omzet_produk'] else 0

    insights = {
        'top_omzet': product_rows[0] if product_rows else None,
        'top_margin': max(product_rows, key=lambda r: r['margin']) if product_rows else None,
        'top_qty': max(product_rows, key=lambda r: r['qty_net']) if product_rows else None,
        'top_category': category_rows[0] if category_rows else None,
        'best_customer': customer_rows[0] if customer_rows else None,
    }
    return {
        'products': product_rows,
        'categories': category_rows,
        'customers': customer_rows,
        'trends': trend_rows,
        'totals': totals,
        'insights': insights,
    }

@app.route('/performa-penjualan')
@investor_required
def performa_penjualan():
    conn = db()
    today = date.today()
    sd = request.args.get('sd', today.replace(day=1).strftime('%Y-%m-%d'))
    ed = request.args.get('ed', today.strftime('%Y-%m-%d'))
    report = _sales_performance_report(conn, sd, ed)
    pnl = calc_profitability(conn, sd, ed)
    conn.close()
    return render_template('performa_penjualan.html',
        sd=sd, ed=ed, report=report, pnl=pnl, today=today.strftime('%Y-%m-%d'))

@app.route('/buku-besar')
@investor_required
def buku_besar():
    conn  = db()
    today = date.today()
    sd    = request.args.get('sd', today.replace(day=1).strftime('%Y-%m-%d'))
    ed    = request.args.get('ed', today.strftime('%Y-%m-%d'))
    kode  = request.args.get('akun', '')

    # Semua akun untuk dropdown (grouped by tipe)
    semua_akun = conn.execute(
        "SELECT kode, nama, tipe, saldo_normal FROM akun ORDER BY kode"
    ).fetchall()

    entri = []
    akun_info   = None
    saldo_awal  = 0
    total_debit = 0
    total_kredit = 0

    if kode:
        akun_row = conn.execute(
            "SELECT id, kode, nama, tipe, saldo_normal FROM akun WHERE kode=?", (kode,)
        ).fetchone()

        if akun_row:
            akun_info = dict(akun_row)
            aid = akun_row['id']
            sn  = akun_row['saldo_normal']   # 'DEBIT' atau 'KREDIT'

            # ── Saldo awal: semua entry SEBELUM periode ──────────────────
            r = conn.execute("""
                SELECT COALESCE(SUM(d.debit),0) as tot_d,
                       COALESCE(SUM(d.kredit),0) as tot_k
                FROM detail_jurnal d
                JOIN jurnal j ON j.id = d.jurnal_id
                WHERE d.akun_id = ? AND j.tanggal < ?
            """, [aid, sd]).fetchone()
            tot_d_awal  = r['tot_d']
            tot_k_awal  = r['tot_k']
            if sn == 'DEBIT':
                saldo_awal = tot_d_awal - tot_k_awal
            else:
                saldo_awal = tot_k_awal - tot_d_awal

            # ── Entri dalam periode ──────────────────────────────────────
            rows = conn.execute("""
                SELECT j.id, j.tanggal, j.keterangan, j.referensi, j.tipe_tx,
                       d.debit, d.kredit
                FROM detail_jurnal d
                JOIN jurnal j ON j.id = d.jurnal_id
                WHERE d.akun_id = ? AND j.tanggal >= ? AND j.tanggal <= ?
                ORDER BY j.tanggal ASC, j.id ASC
            """, [aid, sd, ed]).fetchall()

            saldo_berjalan = saldo_awal
            for row in rows:
                dbt = row['debit']
                krd = row['kredit']
                if sn == 'DEBIT':
                    saldo_berjalan += dbt - krd
                else:
                    saldo_berjalan += krd - dbt
                total_debit  += dbt
                total_kredit += krd
                entri.append({
                    'jurnal_id'  : row['id'],
                    'tanggal'    : row['tanggal'],
                    'keterangan' : row['keterangan'],
                    'referensi'  : row['referensi'] or '',
                    'tipe_tx'    : row['tipe_tx'],
                    'debit'      : dbt,
                    'kredit'     : krd,
                    'saldo'      : saldo_berjalan,
                })

    saldo_akhir = saldo_awal + (total_debit - total_kredit) if (akun_info and akun_info['saldo_normal']=='DEBIT') \
                  else saldo_awal + (total_kredit - total_debit)

    conn.close()
    return render_template('buku_besar.html',
        semua_akun=semua_akun, akun_info=akun_info,
        entri=entri, sd=sd, ed=ed, kode=kode,
        saldo_awal=saldo_awal, saldo_akhir=saldo_akhir,
        total_debit=total_debit, total_kredit=total_kredit,
        today=today.strftime('%Y-%m-%d'),
    )


@app.route('/buku-besar/export')
@investor_required
def buku_besar_export():
    """Export Buku Besar 1 akun ke Excel (mutasi + saldo berjalan), memakai
    perhitungan yang IDENTIK dengan tampilan layar (saldo awal < sd, entri sd..ed)."""
    conn  = db()
    today = date.today()
    sd    = request.args.get('sd', today.replace(day=1).strftime('%Y-%m-%d'))
    ed    = request.args.get('ed', today.strftime('%Y-%m-%d'))
    kode  = request.args.get('akun', '')
    akun_row = conn.execute("SELECT id,kode,nama,tipe,saldo_normal FROM akun WHERE kode=?", (kode,)).fetchone() if kode else None
    if not akun_row:
        conn.close()
        flash('Pilih akun dulu sebelum export Buku Besar.', 'warning')
        return redirect(url_for('buku_besar', sd=sd, ed=ed))
    aid = akun_row['id']; sn = akun_row['saldo_normal']
    r = conn.execute("SELECT COALESCE(SUM(d.debit),0) td, COALESCE(SUM(d.kredit),0) tk "
                     "FROM detail_jurnal d JOIN jurnal j ON j.id=d.jurnal_id "
                     "WHERE d.akun_id=? AND j.tanggal < ?", [aid, sd]).fetchone()
    saldo = (r['td'] - r['tk']) if sn == 'DEBIT' else (r['tk'] - r['td'])
    saldo_awal = saldo
    rows = conn.execute("SELECT j.tanggal, j.nomor_tx, j.keterangan, d.debit, d.kredit "
                        "FROM detail_jurnal d JOIN jurnal j ON j.id=d.jurnal_id "
                        "WHERE d.akun_id=? AND j.tanggal>=? AND j.tanggal<=? "
                        "ORDER BY j.tanggal ASC, j.id ASC", [aid, sd, ed]).fetchall()
    conn.close()
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    import io
    wb = Workbook(); ws = wb.active; ws.title = 'Buku Besar'
    ws.append([f"Buku Besar: {akun_row['kode']} - {akun_row['nama']}"])
    ws.append([f"Periode: {sd} s/d {ed}"])
    ws.append(['Saldo Awal', '', '', '', '', saldo_awal])
    ws.append([])
    cols = ['Tanggal', 'No. Transaksi', 'Keterangan', 'Debit', 'Kredit', 'Saldo']
    ws.append(cols)
    hr = ws.max_row
    for i in range(1, len(cols) + 1):
        c = ws.cell(row=hr, column=i)
        c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='305496')
    total_d = total_k = 0.0
    for row in rows:
        dbt = float(row['debit'] or 0); krd = float(row['kredit'] or 0)
        saldo += (dbt - krd) if sn == 'DEBIT' else (krd - dbt)
        total_d += dbt; total_k += krd
        ws.append([row['tanggal'], row['nomor_tx'] or '', row['keterangan'] or '', dbt, krd, saldo])
    ws.append(['', '', 'TOTAL', total_d, total_k, saldo])
    ws.cell(row=ws.max_row, column=3).font = Font(bold=True)
    for i, w in enumerate([12, 18, 46, 16, 16, 18], 1):
        ws.column_dimensions[chr(64 + i)].width = w
    for rr in range(hr + 1, ws.max_row + 1):
        for cc in (4, 5, 6):
            ws.cell(row=rr, column=cc).number_format = '#,##0'
    ws.cell(row=3, column=6).number_format = '#,##0'
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    fname = f"buku_besar_{akun_row['kode']}_{sd}_{ed}.xlsx"
    return send_file(buf, download_name=fname, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/arus-kas')
@investor_required
def arus_kas():
    conn = db()
    today = date.today()
    sd = request.args.get('sd', today.replace(day=1).strftime('%Y-%m-%d'))
    ed = request.args.get('ed', today.strftime('%Y-%m-%d'))
    kas_ids = get_rekening_ids(conn)
    saldo_awal = 0; operasional=[]; investasi=[]; pendanaan=[]; total_op=total_inv=total_pend=0
    if kas_ids:
        ph = ','.join('?'*len(kas_ids))
        saldo_awal = conn.execute(f"""
            SELECT COALESCE(SUM(d.debit-d.kredit),0)
            FROM detail_jurnal d JOIN jurnal j ON j.id=d.jurnal_id
            WHERE d.akun_id IN ({ph}) AND j.tanggal<?
        """, kas_ids+[sd]).fetchone()[0]
        flows = conn.execute(f"""
            SELECT j.id,j.tanggal,j.keterangan,j.kategori,
                   SUM(CASE WHEN d.akun_id IN ({ph}) THEN d.debit-d.kredit ELSE 0 END) as net_kas
            FROM jurnal j JOIN detail_jurnal d ON d.jurnal_id=j.id
            WHERE j.tanggal>=? AND j.tanggal<=?
            GROUP BY j.id HAVING net_kas!=0 ORDER BY j.tanggal,j.id
        """, kas_ids+[sd,ed]).fetchall()
        # Operasional = catch-all: semua kategori selain INVESTASI/PENDANAAN
        # (mis. GAJI, PAJAK, KOREKSI) ikut operasional supaya tidak ada arus kas
        # yang hilang dari laporan padahal tetap mengubah saldo akhir.
        investasi   = [r for r in flows if r['kategori']=='INVESTASI']
        pendanaan   = [r for r in flows if r['kategori']=='PENDANAAN']
        operasional = [r for r in flows if r['kategori'] not in ('INVESTASI','PENDANAAN')]
        total_op   = sum(r['net_kas'] for r in operasional)
        total_inv  = sum(r['net_kas'] for r in investasi)
        total_pend = sum(r['net_kas'] for r in pendanaan)
    saldo_akhir = conn.execute(f"""
        SELECT COALESCE(SUM(d.debit-d.kredit),0)
        FROM detail_jurnal d JOIN jurnal j ON j.id=d.jurnal_id
        WHERE d.akun_id IN ({ph}) AND j.tanggal<=?
    """, kas_ids+[ed]).fetchone()[0] if kas_ids else 0
    cf = calc_cashflow(conn, sd, ed)
    conn.close()
    return render_template('arus_kas.html',
        operasional=operasional,investasi=investasi,pendanaan=pendanaan,
        saldo_awal=saldo_awal,total_op=total_op,total_inv=total_inv,
        total_pend_fin=total_pend,saldo_akhir=saldo_akhir,sd=sd,ed=ed,
        uang_masuk=cf['masuk'],uang_keluar=cf['keluar'])


# ---------- CHARTS ----------
@app.route('/charts')
@investor_required
def charts():
    months = int(request.args.get('months', 6))
    return render_template('charts.html', months=months)

@app.route('/api/chart-data')
@investor_required
def api_chart_data():
    n = int(request.args.get('months', 12))
    today = date.today()
    conn = db()
    data = []
    kas_ids = get_rekening_ids(conn)
    kas_ph  = ','.join('?'*len(kas_ids)) if kas_ids else '0'

    for i in range(n-1, -1, -1):
        m = today.month - i
        y = today.year
        while m <= 0: m += 12; y -= 1
        ms = date(y, m, 1)
        me = month_end(y, m)
        sd = ms.strftime('%Y-%m-%d')
        ed = me.strftime('%Y-%m-%d')
        label = f"{BULAN[m-1]} {y}"

        pnl = calc_profitability(conn, sd, ed)
        cf  = calc_cashflow(conn, sd, ed)

        saldo_kum = conn.execute(f"""
            SELECT COALESCE(SUM(d.debit-d.kredit),0)
            FROM detail_jurnal d JOIN jurnal j ON j.id=d.jurnal_id
            WHERE d.akun_id IN ({kas_ph}) AND j.tanggal<=?
        """, kas_ids+[ed]).fetchone()[0] if kas_ids else 0

        data.append({
            'label': label,
            'pendapatan': pnl['rev'], 'hpp': pnl['hpp'],
            'laba_kotor': pnl['laba_kotor'], 'op_exp': pnl['op_exp'],
            'laba_op': pnl['laba_op'], 'depr': pnl['depr'],
            'ebit': pnl['ebit'], 'tax': pnl['tax'],
            'laba_bersih': pnl['laba_bersih'], 'prive': pnl['prive'],
            'laba_tahan': pnl['laba_tahan'],
            'uang_masuk': cf['masuk'], 'uang_keluar': cf['keluar'],
            'saldo_kumulatif': saldo_kum,
        })
    conn.close()
    return jsonify(data)


@app.route('/api/chart-year')
@investor_required
def api_chart_year():
    """12 bulan untuk tahun yang dipilih, data lengkap untuk 5 metrik chart."""
    today = date.today()
    y     = int(request.args.get('year', today.year))
    conn  = db()

    kas_ids = get_rekening_ids(conn)
    kas_ph  = ','.join('?'*len(kas_ids)) if kas_ids else '0'

    data = []
    for m in range(1, 13):
        ms  = date(y, m, 1).strftime('%Y-%m-%d')
        me  = month_end(y, m).strftime('%Y-%m-%d')
        pnl = calc_profitability(conn, ms, me)
        total_beban = pnl['hpp'] + pnl['op_exp'] + pnl['depr'] + pnl['interest'] + pnl['tax']

        # Arus kas: debit/kredit pada akun kas per bulan
        if kas_ids:
            kas_masuk  = conn.execute(
                f"SELECT COALESCE(SUM(d.debit-d.kredit),0) FROM detail_jurnal d "
                f"JOIN jurnal j ON j.id=d.jurnal_id "
                f"WHERE d.akun_id IN ({kas_ph}) AND j.tanggal>=? AND j.tanggal<=? AND j.tipe_tx='PEMASUKAN'",
                kas_ids+[ms, me]).fetchone()[0] or 0
            kas_keluar = conn.execute(
                f"SELECT COALESCE(SUM(d.kredit-d.debit),0) FROM detail_jurnal d "
                f"JOIN jurnal j ON j.id=d.jurnal_id "
                f"WHERE d.akun_id IN ({kas_ph}) AND j.tanggal>=? AND j.tanggal<=? AND j.tipe_tx='PENGELUARAN'",
                kas_ids+[ms, me]).fetchone()[0] or 0
        else:
            kas_masuk = kas_keluar = 0

        rev = pnl['rev'] or 0
        gm_pct = round(pnl['laba_kotor'] / rev * 100, 1) if rev else 0
        nm_pct = round(pnl['laba_bersih'] / rev * 100, 1) if rev else 0

        data.append({
            'label':       BULAN[m-1],
            'month':       m,
            'is_future':   (y == today.year and m > today.month) or y > today.year,
            # Chart 1 – Pendapatan vs Beban
            'pendapatan':  rev,
            'total_beban': total_beban,
            'laba_bersih': pnl['laba_bersih'],
            'laba_tahan':  pnl['laba_tahan'],
            # Chart 2 – Laba Bersih (uses laba_bersih)
            # Chart 3 – Arus Kas
            'kas_masuk':   kas_masuk,
            'kas_keluar':  kas_keluar,
            'net_kas':     kas_masuk - kas_keluar,
            # Chart 4 – Margin %
            'gm_pct':      gm_pct,
            'nm_pct':      nm_pct,
            # Chart 5 – Breakdown Biaya
            'hpp':         pnl['hpp'],
            'op_exp':      pnl['op_exp'],
            'depr':        pnl['depr'],
            'tax':         pnl['tax'],
        })
    conn.close()
    return jsonify(data)

@app.route('/api/chart-daily')
@investor_required
def api_chart_daily():
    """Tren harian (pendapatan vs laba bersih) untuk 1 bulan yang dipilih.
    Contoh: ?year=2025&month=6 -> data 1 Juni s/d 30 Juni 2025 per hari.
    Memakai 1 query agregat per-hari (replikasi logika calc_profitability),
    jadi tetap ringan walau data besar."""
    today = date.today()
    try:
        y = int(request.args.get('year', today.year))
        m = int(request.args.get('month', today.month))
    except (TypeError, ValueError):
        y, m = today.year, today.month
    if m < 1 or m > 12:
        m = today.month

    conn = db()
    last_day = calendar.monthrange(y, m)[1]
    ms = f"{y}-{m:02d}-01"
    me = f"{y}-{m:02d}-{last_day:02d}"

    rows = conn.execute("""
        SELECT j.tanggal AS tgl,
          SUM(CASE WHEN a.tipe='PENDAPATAN' THEN d.kredit-d.debit ELSE 0 END) AS rev,
          SUM(CASE WHEN a.kode IN ('5100','5150','5190') THEN d.debit-d.kredit ELSE 0 END) AS hpp,
          SUM(CASE WHEN a.tipe='BEBAN' AND a.kode NOT IN ('5100','5150','5190')
                        AND COALESCE(a.peran,'') NOT IN ('PENYUSUTAN','BUNGA','PAJAK')
                   THEN d.debit-d.kredit ELSE 0 END) AS op_exp,
          SUM(CASE WHEN COALESCE(a.peran,'')='PENYUSUTAN' THEN d.debit-d.kredit ELSE 0 END) AS depr,
          SUM(CASE WHEN COALESCE(a.peran,'')='BUNGA' THEN d.debit-d.kredit ELSE 0 END) AS interest,
          SUM(CASE WHEN COALESCE(a.peran,'')='PAJAK' THEN d.debit-d.kredit ELSE 0 END) AS tax
        FROM jurnal j
        JOIN detail_jurnal d ON d.jurnal_id=j.id
        JOIN akun a ON a.id=d.akun_id
        WHERE j.tanggal>=? AND j.tanggal<=?
        GROUP BY j.tanggal
    """, (ms, me)).fetchall()
    conn.close()

    by_day = {}
    for r in rows:
        tgl = r['tgl'] or ''
        try:
            dd = int(tgl[8:10])
        except (TypeError, ValueError):
            continue
        rev  = float(r['rev'] or 0)
        laba = rev - float(r['hpp'] or 0) - float(r['op_exp'] or 0) \
                   - float(r['depr'] or 0) - float(r['interest'] or 0) - float(r['tax'] or 0)
        by_day[dd] = (rev, laba)

    days = []
    for dnum in range(1, last_day + 1):
        rev, laba = by_day.get(dnum, (0.0, 0.0))
        is_future = date(y, m, dnum) > today
        days.append({
            'day': dnum, 'label': str(dnum),
            'pendapatan': round(rev, 2),
            'laba_bersih': round(laba, 2),
            'is_future': is_future,
        })

    total_rev    = sum(d['pendapatan'] for d in days)
    total_profit = sum(d['laba_bersih'] for d in days)
    active_days  = sum(1 for d in days
                       if not d['is_future'] and (d['pendapatan'] != 0 or d['laba_bersih'] != 0))
    avg_profit   = (total_profit / active_days) if active_days else 0.0

    return jsonify({
        'year': y, 'month': m, 'bulan': BULAN[m-1],
        'days': days,
        'total_pendapatan': round(total_rev, 2),
        'total_laba_bersih': round(total_profit, 2),
        'avg_laba_bersih': round(avg_profit, 2),
        'active_days': active_days,
    })

@app.route('/api/chart-tahunan')
@investor_required
def api_chart_tahunan():
    n = int(request.args.get('years', 5))
    today = date.today()
    conn = db()
    kas_ids = get_rekening_ids(conn)
    kas_ph  = ','.join('?'*len(kas_ids)) if kas_ids else '0'
    data = []

    for i in range(n-1, -1, -1):
        y = today.year - i
        sd = f"{y}-01-01"
        ed = f"{y}-12-31"

        pnl = calc_profitability(conn, sd, ed)
        cf  = calc_cashflow(conn, sd, ed)

        saldo_akhir = conn.execute(f"""
            SELECT COALESCE(SUM(d.debit-d.kredit),0)
            FROM detail_jurnal d JOIN jurnal j ON j.id=d.jurnal_id
            WHERE d.akun_id IN ({kas_ph}) AND j.tanggal<=?
        """, kas_ids+[ed]).fetchone()[0] if kas_ids else 0

        total_beban = pnl['hpp'] + pnl['op_exp'] + pnl['depr'] + pnl['tax'] + pnl['other_e']

        data.append({
            'label': str(y),
            'pendapatan': pnl['rev'], 'hpp': pnl['hpp'],
            'total_beban': total_beban, 'op_exp': pnl['op_exp'],
            'laba_kotor': pnl['laba_kotor'], 'laba_op': pnl['laba_op'],
            'laba_bersih': pnl['laba_bersih'], 'laba_tahan': pnl['laba_tahan'],
            'depr': pnl['depr'], 'tax': pnl['tax'], 'prive': pnl['prive'],
            'uang_masuk': cf['masuk'], 'uang_keluar': cf['keluar'],
            'saldo_akhir': saldo_akhir,
            'pct_laba_bersih': pnl['pct_laba_bersih'],
        })

    # Hitung pertumbuhan pendapatan YoY (%)
    for i in range(len(data)):
        prev = data[i-1]['pendapatan'] if i > 0 else 0
        curr = data[i]['pendapatan']
        data[i]['growth_pct'] = round((curr - prev) / prev * 100, 1) if prev else 0

    conn.close()
    return jsonify(data)


# ---------- DAFTAR TRANSAKSI (Semua / Pemasukan / Pengeluaran / Pelunasan / Transfer) ----------
@app.route('/daftar-transaksi')
@transaction_read_required
def daftar_transaksi():
    conn = db()
    tab       = request.args.get('tab', 'semua')
    if tab not in ('semua', 'pemasukan', 'pengeluaran', 'retur', 'pelunasan', 'transfer'):
        tab = 'semua'
    beban     = request.args.get('beban', '')
    q         = request.args.get('q', '')
    _today    = date.today()
    _def_from = session.get('dash_sd', _today.replace(day=1).strftime('%Y-%m-%d'))
    _def_to   = session.get('dash_ed', _today.strftime('%Y-%m-%d'))
    date_from = request.args.get('date_from', _def_from)
    date_to   = request.args.get('date_to',   _def_to)

    q_where  = ["(j.keterangan LIKE ? OR j.nomor_tx LIKE ?)"] if q else []
    q_params = [f'%{q}%', f'%{q}%'] if q else []
    if date_from:
        q_where.append("j.tanggal >= ?"); q_params.append(date_from)
    if date_to:
        q_where.append("j.tanggal <= ?"); q_params.append(date_to)

    # Sub-label SQL: tampilkan kategori bahasa UMKM bukan akuntansi
    _sub_label_sql = """
        CASE
          WHEN j.tipe_tx='PEMASUKAN' THEN 'Penjualan'
          WHEN j.kategori='INVESTASI' THEN 'Investasi Aset'
          WHEN j.kategori='PENDANAAN' THEN 'Penarikan Owner'
          WHEN EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun ax ON ax.id=dj.akun_id
                      WHERE dj.jurnal_id=j.id AND ax.kode='1130' AND dj.debit>0)
               THEN 'Bahan Baku'
          WHEN EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun ax ON ax.id=dj.akun_id
                      WHERE dj.jurnal_id=j.id AND COALESCE(ax.peran,'')='PAJAK' AND dj.debit>0)
               THEN 'Pajak'
          WHEN EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun ax ON ax.id=dj.akun_id
                      WHERE dj.jurnal_id=j.id AND ax.kode='1160' AND dj.debit>0)
               THEN 'Pinjaman Karyawan'
          ELSE 'Operasional'
        END as sub_label
    """

    def _fetch_masuk(extra_where=[], extra_params=[]):
        w = ["j.tipe_tx='PEMASUKAN'"] + extra_where
        ws = 'WHERE ' + ' AND '.join(w)
        return conn.execute(f"""
            SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan, j.kategori, 'PEMASUKAN' as tipe_tx,
                   'MASUK' as flow_direction, 'USAHA' as flow_group,
                   (SELECT COALESCE(SUM(dj2.kredit),0)
                    FROM detail_jurnal dj2 JOIN akun a2 ON a2.id=dj2.akun_id
                    WHERE dj2.jurnal_id=j.id AND a2.kode LIKE '4%') as total,
                   (SELECT COALESCE(SUM(CASE WHEN ak.is_rekening=1 THEN dk.debit-dk.kredit ELSE 0 END),0)
                    FROM detail_jurnal dk JOIN akun ak ON ak.id=dk.akun_id
                    WHERE dk.jurnal_id=j.id) as net_kas,
                   (SELECT COALESCE(SUM(dh.debit),0)
                    FROM detail_jurnal dh JOIN akun ah ON ah.id=dh.akun_id
                    WHERE dh.jurnal_id=j.id AND ah.kode='5100') as hpp,
                   (SELECT COALESCE(SUM(p.jumlah-p.terbayar),0)
                    FROM piutang p WHERE p.jurnal_id=j.id) as sisa_piutang,
                   {_sub_label_sql},
                   ps.nomor_struk AS pos_nomor,
                   COALESCE(NULLIF(u.nama,''), psh.username) AS pos_kasir
            FROM jurnal j
            LEFT JOIN pos_sale ps   ON ps.jurnal_id = j.id
            LEFT JOIN pos_shift psh ON psh.id = ps.shift_id
            LEFT JOIN users u       ON u.id  = psh.user_id
            {ws} ORDER BY j.tanggal DESC, j.id DESC
        """, extra_params).fetchall()

    def _fetch_keluar(extra_where=[], extra_params=[]):
        w = ["j.tipe_tx='PENGELUARAN'"] + extra_where
        ws = 'WHERE ' + ' AND '.join(w)
        return conn.execute(f"""
            SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan, j.kategori, 'PENGELUARAN' as tipe_tx,
                   'KELUAR' as flow_direction, 'USAHA' as flow_group,
                   COALESCE(SUM(d.debit),0) as total,
                   (SELECT COALESCE(SUM(CASE WHEN ak.is_rekening=1 THEN dk.debit-dk.kredit ELSE 0 END),0)
                    FROM detail_jurnal dk JOIN akun ak ON ak.id=dk.akun_id
                    WHERE dk.jurnal_id=j.id) as net_kas,
                   (SELECT COALESCE(SUM(db.debit),0)
                    FROM detail_jurnal db JOIN akun ab ON ab.id=db.akun_id
                    WHERE db.jurnal_id=j.id AND ab.kode LIKE '6%') as beban,
                   (SELECT COALESCE(SUM(h.jumlah-h.terbayar),0)
                    FROM hutang h WHERE h.jurnal_id=j.id) as sisa_hutang,
                   {_sub_label_sql}
            FROM jurnal j LEFT JOIN detail_jurnal d ON d.jurnal_id=j.id
            {ws} GROUP BY j.id ORDER BY j.tanggal DESC, j.id DESC
        """, extra_params).fetchall()

    def _fetch_pelunasan(extra_where=[], extra_params=[]):
        w = ["j.tipe_tx='PELUNASAN'"] + extra_where
        ws = 'WHERE ' + ' AND '.join(w)
        return conn.execute(f"""
            SELECT x.id, x.nomor_tx, x.tanggal, x.keterangan, x.kategori, 'PELUNASAN' as tipe_tx,
                   CASE WHEN x.net_kas>=0 THEN 'MASUK' ELSE 'KELUAR' END as flow_direction,
                   'PELUNASAN' as flow_group, ABS(x.net_kas) as total, x.net_kas as net_kas,
                   CASE WHEN EXISTS(SELECT 1 FROM detail_jurnal dj3 JOIN akun a3 ON a3.id=dj3.akun_id
                                    WHERE dj3.jurnal_id=x.id AND a3.kode='1160' AND dj3.kredit>0)
                             THEN 'Pengembalian Pinjaman Karyawan'
                        WHEN x.net_kas>=0 THEN 'Penerimaan Piutang' ELSE 'Pembayaran Hutang' END as sub_label
            FROM (
                SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan, j.kategori,
                       (SELECT COALESCE(SUM(CASE WHEN a2.is_rekening=1 THEN dj2.debit-dj2.kredit ELSE 0 END),0)
                        FROM detail_jurnal dj2 JOIN akun a2 ON a2.id=dj2.akun_id
                        WHERE dj2.jurnal_id=j.id) as net_kas
                FROM jurnal j {ws}
            ) x
            ORDER BY x.tanggal DESC, x.id DESC
        """, extra_params).fetchall()

    def _fetch_transfer(extra_where=[], extra_params=[]):
        w = ["j.tipe_tx='TRANSFER'"] + extra_where
        ws = 'WHERE ' + ' AND '.join(w)
        return conn.execute(f"""
            SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan, j.kategori, 'TRANSFER' as tipe_tx,
                   'TRANSFER' as flow_direction, 'TRANSFER' as flow_group,
                   (SELECT COALESCE(SUM(dj2.debit),0)
                    FROM detail_jurnal dj2 JOIN akun a2 ON a2.id=dj2.akun_id
                    WHERE dj2.jurnal_id=j.id AND a2.is_rekening=1) as total,
                   0 as net_kas,
                   'Transfer Rekening' as sub_label
            FROM jurnal j {ws} ORDER BY j.tanggal DESC, j.id DESC
        """, extra_params).fetchall()

    def _fetch_retur(extra_where=[], extra_params=[]):
        w = ["j.tipe_tx IN ('RETUR_JUAL','RETUR_BELI')"] + extra_where
        ws = 'WHERE ' + ' AND '.join(w)
        return conn.execute(f"""
            SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan, j.kategori, j.tipe_tx,
                   CASE WHEN j.tipe_tx='RETUR_JUAL' THEN 'KELUAR' ELSE 'MASUK' END as flow_direction,
                   'RETUR' as flow_group,
                   -- Nominal retur murni (DPP): 4150 utk retur jual, 1130 utk retur beli.
                   -- TIDAK pakai SUM semua debit jurnal (jebakan spt bug donut Cashflow —
                   -- baris reklasifikasi HPP 1130/5100 bukan bagian dari "nilai retur").
                   CASE WHEN j.tipe_tx='RETUR_JUAL'
                        THEN (SELECT COALESCE(SUM(dj.debit),0) FROM detail_jurnal dj
                              JOIN akun a2 ON a2.id=dj.akun_id WHERE dj.jurnal_id=j.id AND a2.kode='4150')
                        ELSE (SELECT COALESCE(SUM(dj.kredit),0) FROM detail_jurnal dj
                              JOIN akun a2 ON a2.id=dj.akun_id WHERE dj.jurnal_id=j.id AND a2.kode='1130')
                   END as total,
                   (SELECT COALESCE(SUM(CASE WHEN ak.is_rekening=1 THEN dk.debit-dk.kredit ELSE 0 END),0)
                    FROM detail_jurnal dk JOIN akun ak ON ak.id=dk.akun_id
                    WHERE dk.jurnal_id=j.id) as net_kas,
                   -- Dampak ke laba rugi: netting per-tipe (spt _qv), bukan raw debit.
                   -- Retur beli TIDAK sentuh akun PENDAPATAN/HPP -> otomatis 0 (bukan P&L).
                   ((SELECT COALESCE(SUM(CASE WHEN a3.tipe='PENDAPATAN' THEN dj3.kredit-dj3.debit ELSE 0 END),0)
                     FROM detail_jurnal dj3 JOIN akun a3 ON a3.id=dj3.akun_id WHERE dj3.jurnal_id=j.id)
                    -
                    (SELECT COALESCE(SUM(CASE WHEN a4.kode IN ('5100','5150','5190') THEN dj4.debit-dj4.kredit ELSE 0 END),0)
                     FROM detail_jurnal dj4 JOIN akun a4 ON a4.id=dj4.akun_id WHERE dj4.jurnal_id=j.id)
                   ) as laba_impact,
                   CASE WHEN j.tipe_tx='RETUR_JUAL' THEN 'Retur Penjualan' ELSE 'Retur Pembelian' END as sub_label
            FROM jurnal j
            {ws} ORDER BY j.tanggal DESC, j.id DESC
        """, extra_params).fetchall()

    # ── Efisien utk jutaan baris: total via AGREGAT SQL (bukan sum baris Python),
    #    count + id halaman via id-first (cuma sentuh index jurnal), lalu perkaya
    #    HANYA 50 baris halaman via _fetch_* + `j.id IN (...)` (output tetap identik). ──
    def _extra(wl):
        return (' AND ' + ' AND '.join(wl)) if wl else ''
    def _agg_masuk(wl, pp):
        return conn.execute(
            "SELECT COALESCE(SUM(dj.kredit),0) FROM detail_jurnal dj "
            "JOIN akun a2 ON a2.id=dj.akun_id JOIN jurnal j ON j.id=dj.jurnal_id "
            "WHERE j.tipe_tx='PEMASUKAN' AND a2.kode LIKE '4%'" + _extra(wl), pp).fetchone()[0]
    def _agg_keluar(wl, pp):
        return conn.execute(
            "SELECT COALESCE(SUM(d.debit),0) FROM detail_jurnal d "
            "JOIN jurnal j ON j.id=d.jurnal_id WHERE j.tipe_tx='PENGELUARAN'" + _extra(wl), pp).fetchone()[0]
    def _agg_transfer(wl, pp):
        return conn.execute(
            "SELECT COALESCE(SUM(dj.debit),0) FROM detail_jurnal dj "
            "JOIN akun a2 ON a2.id=dj.akun_id JOIN jurnal j ON j.id=dj.jurnal_id "
            "WHERE j.tipe_tx='TRANSFER' AND a2.is_rekening=1" + _extra(wl), pp).fetchone()[0]
    def _agg_pelunasan(wl, pp):
        r = conn.execute(
            "SELECT COALESCE(SUM(CASE WHEN nk>=0 THEN nk ELSE 0 END),0) AS m, "
            "       COALESCE(SUM(CASE WHEN nk<0 THEN -nk ELSE 0 END),0) AS k FROM ("
            "  SELECT (SELECT COALESCE(SUM(CASE WHEN a2.is_rekening=1 THEN dj2.debit-dj2.kredit ELSE 0 END),0) "
            "          FROM detail_jurnal dj2 JOIN akun a2 ON a2.id=dj2.akun_id WHERE dj2.jurnal_id=j.id) AS nk "
            "  FROM jurnal j WHERE j.tipe_tx='PELUNASAN'" + _extra(wl) + ")", pp).fetchone()
        return float(r['m'] or 0), float(r['k'] or 0)
    def _agg_retur(wl, pp):
        jual = conn.execute(
            "SELECT COALESCE(SUM(dj.debit),0) FROM detail_jurnal dj "
            "JOIN akun a2 ON a2.id=dj.akun_id JOIN jurnal j ON j.id=dj.jurnal_id "
            "WHERE j.tipe_tx='RETUR_JUAL' AND a2.kode='4150'" + _extra(wl), pp).fetchone()[0]
        beli = conn.execute(
            "SELECT COALESCE(SUM(dj.kredit),0) FROM detail_jurnal dj "
            "JOIN akun a2 ON a2.id=dj.akun_id JOIN jurnal j ON j.id=dj.jurnal_id "
            "WHERE j.tipe_tx='RETUR_BELI' AND a2.kode='1130'" + _extra(wl), pp).fetchone()[0]
        return float(jual or 0), float(beli or 0)
    def _page_ids(tipe_where, wl, pp):
        wsql = "WHERE " + tipe_where + _extra(wl)
        n = conn.execute(f"SELECT COUNT(*) FROM jurnal j {wsql}", pp).fetchone()[0]
        _pg = paginate_meta(_page_arg(), n)
        _ids = [row['id'] for row in conn.execute(
            f"SELECT j.id FROM jurnal j {wsql} ORDER BY j.tanggal DESC, j.id DESC LIMIT ? OFFSET ?",
            pp + [_pg['per_page'], _pg['offset']]).fetchall()]
        return _pg, _ids
    def _enrich(fetch_fn, ids):
        if not ids:
            return []
        idw = ["j.id IN (%s)" % ','.join('?' for _ in ids)]
        return sorted((dict(r) for r in fetch_fn(idw, ids)),
                      key=lambda r: (r['tanggal'], r['id']), reverse=True)

    total_masuk = total_keluar = total_pelunasan_masuk = total_pelunasan_keluar = total_transfer = 0
    total_retur_jual = total_retur_beli = 0
    if tab == 'semua':
        pg, ids = _page_ids("j.tipe_tx IN ('PEMASUKAN','PENGELUARAN','PELUNASAN','TRANSFER','RETUR_JUAL','RETUR_BELI')", q_where, q_params)
        if ids:
            idw = ["j.id IN (%s)" % ','.join('?' for _ in ids)]
            rows = ([dict(r) for r in _fetch_masuk(idw, ids)]
                    + [dict(r) for r in _fetch_keluar(idw, ids)]
                    + [dict(r) for r in _fetch_pelunasan(idw, ids)]
                    + [dict(r) for r in _fetch_transfer(idw, ids)]
                    + [dict(r) for r in _fetch_retur(idw, ids)])
            rows.sort(key=lambda r: (r['tanggal'], r['id']), reverse=True)
        else:
            rows = []
        total_masuk  = _agg_masuk(q_where, q_params)
        total_keluar = _agg_keluar(q_where, q_params)
        total_transfer = _agg_transfer(q_where, q_params)
        total_pelunasan_masuk, total_pelunasan_keluar = _agg_pelunasan(q_where, q_params)
        total_retur_jual, total_retur_beli = _agg_retur(q_where, q_params)
        total_all    = total_masuk - total_keluar
    elif tab == 'pemasukan':
        pg, ids = _page_ids("j.tipe_tx='PEMASUKAN'", q_where, q_params)
        rows = _enrich(_fetch_masuk, ids)
        total_masuk = total_all = _agg_masuk(q_where, q_params)
    elif tab == 'pengeluaran':
        beban_w = []
        if beban == 'bahan_baku':
            # Bahan baku → Dr 1130 (Persediaan), bukan 5100
            beban_w.append("EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=j.id AND a.kode='1130' AND dj.debit>0)")
        elif beban == 'operasional':
            beban_w += ["j.kategori='OPERASIONAL'",
                        "NOT EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=j.id AND (a.kode IN('1130','1160') OR COALESCE(a.peran,'')='PAJAK') AND dj.debit>0)"]
        elif beban == 'investasi':
            beban_w.append("j.kategori='INVESTASI'")
        elif beban == 'pajak':
            beban_w.append("EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=j.id AND COALESCE(a.peran,'')='PAJAK' AND dj.debit>0)")
        elif beban == 'penarikan':
            beban_w.append("j.kategori='PENDANAAN'")
        elif beban == 'pinjaman_karyawan':
            # Pinjaman karyawan → Dr 1160 (Piutang Karyawan), non-beban
            beban_w.append("EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=j.id AND a.kode='1160' AND dj.debit>0)")
        pg, ids = _page_ids("j.tipe_tx='PENGELUARAN'", q_where + beban_w, q_params)
        rows = _enrich(_fetch_keluar, ids)
        total_keluar = total_all = _agg_keluar(q_where + beban_w, q_params)
    elif tab == 'retur':
        pg, ids = _page_ids("j.tipe_tx IN ('RETUR_JUAL','RETUR_BELI')", q_where, q_params)
        rows = _enrich(_fetch_retur, ids)
        total_retur_jual, total_retur_beli = _agg_retur(q_where, q_params)
        total_all = total_retur_beli - total_retur_jual
    elif tab == 'pelunasan':
        pg, ids = _page_ids("j.tipe_tx='PELUNASAN'", q_where, q_params)
        rows = _enrich(_fetch_pelunasan, ids)
        total_pelunasan_masuk, total_pelunasan_keluar = _agg_pelunasan(q_where, q_params)
        total_all = total_pelunasan_masuk - total_pelunasan_keluar
    else:
        pg, ids = _page_ids("j.tipe_tx='TRANSFER'", q_where, q_params)
        rows = _enrich(_fetch_transfer, ids)
        total_transfer = total_all = _agg_transfer(q_where, q_params)

    conn.close()
    return render_template('daftar_transaksi.html',
                           rows=rows, tab=tab, beban=beban, q=q,
                           date_from=date_from, date_to=date_to,
                           total_all=total_all,
                           total_masuk=total_masuk, total_keluar=total_keluar,
                           total_pelunasan_masuk=total_pelunasan_masuk,
                           total_pelunasan_keluar=total_pelunasan_keluar,
                           total_transfer=total_transfer,
                           total_retur_jual=total_retur_jual,
                           total_retur_beli=total_retur_beli, pg=pg)


@app.route('/daftar-transaksi/export')
@transaction_read_required
def daftar_transaksi_export():
    if not HAS_XLSX:
        flash('Library openpyxl tidak tersedia.', 'danger')
        return redirect(url_for('daftar_transaksi'))

    conn = db()
    tab       = request.args.get('tab', 'semua')
    if tab not in ('semua', 'pemasukan', 'pengeluaran', 'retur', 'pelunasan', 'transfer'):
        tab = 'semua'
    beban     = request.args.get('beban', '')
    q         = request.args.get('q', '')
    date_from = request.args.get('date_from', '')
    date_to   = request.args.get('date_to', '')

    q_where  = ["(j.keterangan LIKE ? OR j.nomor_tx LIKE ?)"] if q else []
    q_params = [f'%{q}%', f'%{q}%'] if q else []
    if date_from:
        q_where.append("j.tanggal >= ?"); q_params.append(date_from)
    if date_to:
        q_where.append("j.tanggal <= ?"); q_params.append(date_to)

    _sub_label_sql = """
        CASE
          WHEN j.tipe_tx='PEMASUKAN' THEN 'Penjualan'
          WHEN j.kategori='INVESTASI' THEN 'Investasi Aset'
          WHEN j.kategori='PENDANAAN' THEN 'Penarikan Owner'
          WHEN EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun ax ON ax.id=dj.akun_id
                      WHERE dj.jurnal_id=j.id AND ax.kode='1130' AND dj.debit>0)
               THEN 'Bahan Baku'
          WHEN EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun ax ON ax.id=dj.akun_id
                      WHERE dj.jurnal_id=j.id AND COALESCE(ax.peran,'')='PAJAK' AND dj.debit>0)
               THEN 'Pajak'
          WHEN EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun ax ON ax.id=dj.akun_id
                      WHERE dj.jurnal_id=j.id AND ax.kode='1160' AND dj.debit>0)
               THEN 'Pinjaman Karyawan'
          ELSE 'Operasional'
        END as sub_label
    """

    def _fetch_masuk(ew=[], ep=[]):
        w = ["j.tipe_tx='PEMASUKAN'"] + ew
        ws = 'WHERE ' + ' AND '.join(w)
        return conn.execute(f"""
            SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan, 'PEMASUKAN' as tipe_tx,
                   'MASUK' as flow_direction, 'USAHA' as flow_group,
                   (SELECT COALESCE(SUM(dj2.kredit),0)
                    FROM detail_jurnal dj2 JOIN akun a2 ON a2.id=dj2.akun_id
                    WHERE dj2.jurnal_id=j.id AND a2.kode LIKE '4%') as total,
                   {_sub_label_sql}
            FROM jurnal j {ws} ORDER BY j.tanggal, j.id
        """, ep).fetchall()

    def _fetch_keluar(ew=[], ep=[]):
        w = ["j.tipe_tx='PENGELUARAN'"] + ew
        ws = 'WHERE ' + ' AND '.join(w)
        beban_w = list(ew)
        if beban == 'bahan_baku':
            beban_w.append("EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=j.id AND a.kode='1130' AND dj.debit>0)")
        elif beban == 'operasional':
            beban_w += ["j.kategori='OPERASIONAL'",
                        "NOT EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=j.id AND (a.kode IN('1130','1160') OR COALESCE(a.peran,'')='PAJAK') AND dj.debit>0)"]
        elif beban == 'investasi':
            beban_w.append("j.kategori='INVESTASI'")
        elif beban == 'pajak':
            beban_w.append("EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=j.id AND COALESCE(a.peran,'')='PAJAK' AND dj.debit>0)")
        elif beban == 'penarikan':
            beban_w.append("j.kategori='PENDANAAN'")
        elif beban == 'pinjaman_karyawan':
            beban_w.append("EXISTS(SELECT 1 FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=j.id AND a.kode='1160' AND dj.debit>0)")
        w2 = ["j.tipe_tx='PENGELUARAN'"] + beban_w
        ws2 = 'WHERE ' + ' AND '.join(w2)
        return conn.execute(f"""
            SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan, 'PENGELUARAN' as tipe_tx,
                   'KELUAR' as flow_direction, 'USAHA' as flow_group,
                   COALESCE(SUM(d.debit),0) as total,
                   {_sub_label_sql}
            FROM jurnal j LEFT JOIN detail_jurnal d ON d.jurnal_id=j.id
            {ws2} GROUP BY j.id ORDER BY j.tanggal, j.id
        """, ep).fetchall()

    def _fetch_pelunasan(ew=[], ep=[]):
        w = ["j.tipe_tx='PELUNASAN'"] + ew
        ws = 'WHERE ' + ' AND '.join(w)
        return conn.execute(f"""
            SELECT x.id, x.nomor_tx, x.tanggal, x.keterangan, 'PELUNASAN' as tipe_tx,
                   CASE WHEN x.net_kas>=0 THEN 'MASUK' ELSE 'KELUAR' END as flow_direction,
                   'PELUNASAN' as flow_group, ABS(x.net_kas) as total,
                   CASE WHEN EXISTS(SELECT 1 FROM detail_jurnal dj3 JOIN akun a3 ON a3.id=dj3.akun_id
                                    WHERE dj3.jurnal_id=x.id AND a3.kode='1160' AND dj3.kredit>0)
                             THEN 'Pengembalian Pinjaman Karyawan'
                        WHEN x.net_kas>=0 THEN 'Penerimaan Piutang' ELSE 'Pembayaran Hutang' END as sub_label
            FROM (
                SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan,
                       (SELECT COALESCE(SUM(CASE WHEN a2.is_rekening=1 THEN dj2.debit-dj2.kredit ELSE 0 END),0)
                        FROM detail_jurnal dj2 JOIN akun a2 ON a2.id=dj2.akun_id
                        WHERE dj2.jurnal_id=j.id) as net_kas
                FROM jurnal j {ws}
            ) x
            ORDER BY x.tanggal, x.id
        """, ep).fetchall()

    def _fetch_transfer(ew=[], ep=[]):
        w = ["j.tipe_tx='TRANSFER'"] + ew
        ws = 'WHERE ' + ' AND '.join(w)
        return conn.execute(f"""
            SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan, 'TRANSFER' as tipe_tx,
                   'TRANSFER' as flow_direction, 'TRANSFER' as flow_group,
                   (SELECT COALESCE(SUM(dj2.debit),0)
                    FROM detail_jurnal dj2 JOIN akun a2 ON a2.id=dj2.akun_id
                    WHERE dj2.jurnal_id=j.id AND a2.is_rekening=1) as total,
                   'Transfer Rekening' as sub_label
            FROM jurnal j {ws} ORDER BY j.tanggal, j.id
        """, ep).fetchall()

    def _fetch_retur(ew=[], ep=[]):
        w = ["j.tipe_tx IN ('RETUR_JUAL','RETUR_BELI')"] + ew
        ws = 'WHERE ' + ' AND '.join(w)
        return conn.execute(f"""
            SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan, j.tipe_tx,
                   CASE WHEN j.tipe_tx='RETUR_JUAL' THEN 'KELUAR' ELSE 'MASUK' END as flow_direction,
                   'RETUR' as flow_group,
                   CASE WHEN j.tipe_tx='RETUR_JUAL'
                        THEN (SELECT COALESCE(SUM(dj.debit),0) FROM detail_jurnal dj
                              JOIN akun a2 ON a2.id=dj.akun_id WHERE dj.jurnal_id=j.id AND a2.kode='4150')
                        ELSE (SELECT COALESCE(SUM(dj.kredit),0) FROM detail_jurnal dj
                              JOIN akun a2 ON a2.id=dj.akun_id WHERE dj.jurnal_id=j.id AND a2.kode='1130')
                   END as total,
                   CASE WHEN j.tipe_tx='RETUR_JUAL' THEN 'Retur Penjualan' ELSE 'Retur Pembelian' END as sub_label
            FROM jurnal j {ws} ORDER BY j.tanggal, j.id
        """, ep).fetchall()

    if tab == 'pemasukan':
        rows = _fetch_masuk(q_where, q_params)
    elif tab == 'pengeluaran':
        rows = _fetch_keluar(q_where, q_params)
    elif tab == 'retur':
        rows = _fetch_retur(q_where, q_params)
    elif tab == 'pelunasan':
        rows = _fetch_pelunasan(q_where, q_params)
    elif tab == 'transfer':
        rows = _fetch_transfer(q_where, q_params)
    else:
        rp = [dict(r) for r in _fetch_masuk(q_where, q_params)]
        rk = [dict(r) for r in _fetch_keluar(q_where, q_params)]
        rl = [dict(r) for r in _fetch_pelunasan(q_where, q_params)]
        rt = [dict(r) for r in _fetch_transfer(q_where, q_params)]
        rr = [dict(r) for r in _fetch_retur(q_where, q_params)]
        rows = sorted(rp + rk + rl + rt + rr, key=lambda r: (r['tanggal'], r['id']))

    conn.close()

    # ── Build Excel ───────────────────────────────────────────────────
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    tab_label = {'semua': 'Semua', 'pemasukan': 'Penjualan', 'pengeluaran': 'Pengeluaran',
                 'retur': 'Retur', 'pelunasan': 'Pelunasan', 'transfer': 'Transfer'}.get(tab, tab)
    ws.title = f'Transaksi {tab_label}'

    # judul
    ws.merge_cells('A1:F1')
    ws['A1'] = f'Daftar Transaksi, {tab_label}'
    ws['A1'].font = Font(bold=True, size=13)
    ws['A2'] = f'Periode: {date_from or "-"}  s/d  {date_to or "-"}'
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.append([])

    # header
    hdr = ['ID Transaksi', 'Tanggal', 'Keterangan', 'Jenis', 'Tipe', 'Nominal (Rp)']
    ws.append(hdr)
    hrow = ws.max_row
    hfill = PatternFill('solid', fgColor='1E293B')
    hfont = Font(bold=True, color='FFFFFF', size=10)
    for col in range(1, 7):
        cell = ws.cell(hrow, col)
        cell.fill = hfill; cell.font = hfont
        cell.alignment = Alignment(horizontal='center')

    # data
    total_masuk = total_keluar = total_pelunasan_masuk = total_pelunasan_keluar = total_transfer = 0
    total_retur_jual = total_retur_beli = 0
    for r in rows:
        is_masuk = r['flow_direction'] == 'MASUK'
        is_pelunasan = r['flow_group'] == 'PELUNASAN'
        is_transfer = r['flow_group'] == 'TRANSFER'
        is_retur = r['flow_group'] == 'RETUR'
        nominal = r['total']
        if is_transfer:
            total_transfer += nominal
        elif is_retur:
            if r['tipe_tx'] == 'RETUR_JUAL':
                total_retur_jual += nominal
            else:
                total_retur_beli += nominal
        elif is_pelunasan and is_masuk:
            total_pelunasan_masuk += nominal
        elif is_pelunasan:
            total_pelunasan_keluar += nominal
        elif is_masuk:
            total_masuk += nominal
        else:
            total_keluar += nominal
        tipe_text = ('Transfer' if is_transfer else
                     'Retur' if is_retur else
                     'Pelunasan' if is_pelunasan else
                     'Pemasukan' if is_masuk else 'Pengeluaran')
        ws.append([
            r['nomor_tx'],
            r['tanggal'],
            r['keterangan'],
            r['sub_label'],
            tipe_text,
            nominal if (is_masuk or is_transfer) else -nominal,
        ])
        drow = ws.max_row
        ws.cell(drow, 6).number_format = '#,##0'
        ws.cell(drow, 6).font = Font(
            color='0E7490' if is_transfer else ('16A34A' if is_masuk else 'DC2626'), bold=True)

    # footer
    ws.append([])
    frow = ws.max_row + 1
    if tab == 'semua':
        ws.cell(frow, 5).value = 'Total Pemasukan'
        ws.cell(frow, 6).value = total_masuk
        ws.cell(frow, 6).number_format = '#,##0'
        ws.cell(frow, 6).font = Font(bold=True, color='16A34A')
        ws.cell(frow+1, 5).value = 'Total Pengeluaran'
        ws.cell(frow+1, 6).value = -total_keluar
        ws.cell(frow+1, 6).number_format = '#,##0'
        ws.cell(frow+1, 6).font = Font(bold=True, color='DC2626')
        ws.cell(frow+2, 5).value = 'Selisih'
        ws.cell(frow+2, 6).value = total_masuk - total_keluar
        ws.cell(frow+2, 6).number_format = '#,##0'
        ws.cell(frow+2, 6).font = Font(bold=True, size=12)
        ws.cell(frow+4, 5).value = 'Penerimaan Piutang'
        ws.cell(frow+4, 6).value = total_pelunasan_masuk
        ws.cell(frow+4, 6).number_format = '#,##0'
        ws.cell(frow+4, 6).font = Font(bold=True, color='0F766E')
        ws.cell(frow+5, 5).value = 'Pembayaran Hutang'
        ws.cell(frow+5, 6).value = -total_pelunasan_keluar
        ws.cell(frow+5, 6).number_format = '#,##0'
        ws.cell(frow+5, 6).font = Font(bold=True, color='7C3AED')
        ws.cell(frow+7, 5).value = 'Retur Penjualan'
        ws.cell(frow+7, 6).value = -total_retur_jual
        ws.cell(frow+7, 6).number_format = '#,##0'
        ws.cell(frow+7, 6).font = Font(bold=True, color='C2410C')
        ws.cell(frow+8, 5).value = 'Retur Pembelian'
        ws.cell(frow+8, 6).value = total_retur_beli
        ws.cell(frow+8, 6).number_format = '#,##0'
        ws.cell(frow+8, 6).font = Font(bold=True, color='0369A1')
        ws.cell(frow+10, 5).value = 'Transfer Rekening'
        ws.cell(frow+10, 6).value = total_transfer
        ws.cell(frow+10, 6).number_format = '#,##0'
        ws.cell(frow+10, 6).font = Font(bold=True, color='0E7490')
    elif tab == 'pemasukan':
        ws.cell(frow, 5).value = 'Total Pemasukan'
        ws.cell(frow, 6).value = total_masuk
        ws.cell(frow, 6).number_format = '#,##0'
        ws.cell(frow, 6).font = Font(bold=True, color='16A34A')
    elif tab == 'pengeluaran':
        ws.cell(frow, 5).value = 'Total Pengeluaran'
        ws.cell(frow, 6).value = -total_keluar
        ws.cell(frow, 6).number_format = '#,##0'
        ws.cell(frow, 6).font = Font(bold=True, color='DC2626')
    elif tab == 'retur':
        ws.cell(frow, 5).value = 'Retur Penjualan'
        ws.cell(frow, 6).value = -total_retur_jual
        ws.cell(frow, 6).number_format = '#,##0'
        ws.cell(frow, 6).font = Font(bold=True, color='C2410C')
        ws.cell(frow+1, 5).value = 'Retur Pembelian'
        ws.cell(frow+1, 6).value = total_retur_beli
        ws.cell(frow+1, 6).number_format = '#,##0'
        ws.cell(frow+1, 6).font = Font(bold=True, color='0369A1')
        ws.cell(frow+2, 5).value = 'Selisih Retur'
        ws.cell(frow+2, 6).value = total_retur_beli - total_retur_jual
        ws.cell(frow+2, 6).number_format = '#,##0'
        ws.cell(frow+2, 6).font = Font(bold=True, size=12)
    elif tab == 'pelunasan':
        ws.cell(frow, 5).value = 'Penerimaan Piutang'
        ws.cell(frow, 6).value = total_pelunasan_masuk
        ws.cell(frow, 6).number_format = '#,##0'
        ws.cell(frow, 6).font = Font(bold=True, color='0F766E')
        ws.cell(frow+1, 5).value = 'Pembayaran Hutang'
        ws.cell(frow+1, 6).value = -total_pelunasan_keluar
        ws.cell(frow+1, 6).number_format = '#,##0'
        ws.cell(frow+1, 6).font = Font(bold=True, color='7C3AED')
        ws.cell(frow+2, 5).value = 'Selisih Pelunasan'
        ws.cell(frow+2, 6).value = total_pelunasan_masuk - total_pelunasan_keluar
        ws.cell(frow+2, 6).number_format = '#,##0'
        ws.cell(frow+2, 6).font = Font(bold=True, size=12)
    else:
        ws.cell(frow, 5).value = 'Total Transfer Rekening'
        ws.cell(frow, 6).value = total_transfer
        ws.cell(frow, 6).number_format = '#,##0'
        ws.cell(frow, 6).font = Font(bold=True, color='0E7490')

    # lebar kolom
    for col, w in zip(range(1, 7), [19, 13, 42, 18, 14, 18]):
        ws.column_dimensions[get_column_letter(col)].width = w

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    fname_parts = ['transaksi', tab]
    if date_from: fname_parts.append(date_from)
    if date_to:   fname_parts.append(date_to)
    fname = '_'.join(fname_parts) + '.xlsx'
    return send_file(buf, download_name=fname, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ---------- TRANSAKSI (Advanced Journal) ----------
MANUAL_JOURNAL_CATEGORIES = {'OPERASIONAL', 'INVESTASI', 'PENDANAAN'}

@app.route('/transaksi')
@transaction_read_required
def transaksi_list():
    conn = db()
    q = request.args.get('q',''); kat = request.args.get('kat','')
    params=[]; where=[]
    if q: where.append("(j.nomor_tx LIKE ? OR j.keterangan LIKE ? OR j.referensi LIKE ?)"); params+=[f'%{q}%',f'%{q}%',f'%{q}%']
    if kat: where.append("j.kategori=?"); params.append(kat)
    ws = ('WHERE '+' AND '.join(where)) if where else ''
    # Paginasi server-side (skala 1,1 jt+ baris): hitung total, pilih HANYA id
    # halaman ini (cuma sentuh index jurnal), lalu join detail utk ~50 baris saja.
    page = _page_arg()
    n_rows = conn.execute(f"SELECT COUNT(*) FROM jurnal j {ws}", params).fetchone()[0]
    pg = paginate_meta(page, n_rows)
    id_rows = conn.execute(
        f"SELECT j.id FROM jurnal j {ws} ORDER BY j.tanggal DESC, j.id DESC LIMIT ? OFFSET ?",
        params + [pg['per_page'], pg['offset']]
    ).fetchall()
    ids = [r['id'] for r in id_rows]
    if ids:
        ph = ','.join('?' for _ in ids)
        rows = conn.execute(f"""
            SELECT j.id,j.nomor_tx,j.tanggal,j.keterangan,j.referensi,j.kategori,j.tipe_tx,
                   COALESCE(SUM(d.debit),0) as total,
                   ps.nomor_struk AS pos_nomor,
                   COALESCE(NULLIF(u.nama,''), psh.username) AS pos_kasir
            FROM jurnal j
            LEFT JOIN detail_jurnal d ON d.jurnal_id=j.id
            LEFT JOIN pos_sale ps   ON ps.jurnal_id = j.id
            LEFT JOIN pos_shift psh ON psh.id = ps.shift_id
            LEFT JOIN users u       ON u.id  = psh.user_id
            WHERE j.id IN ({ph})
            GROUP BY j.id ORDER BY j.tanggal DESC, j.id DESC
        """, ids).fetchall()
    else:
        rows = []
    conn.close()
    return render_template('transaksi.html', rows=rows, q=q, kat=kat, pg=pg)

@app.route('/transaksi/baru', methods=['GET','POST'])
@finance_required
def transaksi_baru_form():
    conn = db()
    if request.method == 'POST':
        tanggal    = request.form['tanggal']
        keterangan = request.form['keterangan']
        referensi  = request.form.get('referensi','')
        kategori   = request.form.get('kategori','OPERASIONAL')
        akun_ids   = request.form.getlist('akun_id[]')
        debits     = request.form.getlist('debit[]')
        kredits    = request.form.getlist('kredit[]')
        if not is_valid_iso_date(tanggal):
            flash('Tanggal jurnal tidak valid.', 'danger')
            akun_list = conn.execute('SELECT * FROM akun ORDER BY kode').fetchall()
            conn.close()
            return render_template('transaksi_form.html', akun_list=akun_list)
        if kategori not in MANUAL_JOURNAL_CATEGORIES:
            flash('Kategori arus kas jurnal tidak valid.', 'danger')
            akun_list = conn.execute('SELECT * FROM akun ORDER BY kode').fetchall()
            conn.close()
            return render_template('transaksi_form.html', akun_list=akun_list)
        try:
            lines = parse_manual_journal_lines(conn, akun_ids, debits, kredits)
        except ValueError as ex:
            flash(str(ex), 'danger')
            akun_list = conn.execute('SELECT * FROM akun ORDER BY kode').fetchall()
            conn.close()
            return render_template('transaksi_form.html', akun_list=akun_list)
        cur = conn.execute(
            'INSERT INTO jurnal(tanggal,keterangan,referensi,kategori,tipe_tx,nomor_tx) VALUES(?,?,?,?,?,?)',
            (tanggal, keterangan, referensi, kategori, 'JURNAL', next_nomor_tx(conn, tanggal))
        )
        jid = cur.lastrowid
        conn.executemany(
            'INSERT INTO detail_jurnal(jurnal_id,akun_id,debit,kredit) VALUES(?,?,?,?)',
            [(jid, akun_id, debit, kredit) for akun_id, debit, kredit in lines]
        )
        conn.commit(); conn.close()
        flash('Jurnal berhasil disimpan!','success')
        return redirect(url_for('transaksi_list'))
    akun_list = conn.execute('SELECT * FROM akun ORDER BY kode').fetchall()
    conn.close()
    return render_template('transaksi_form.html', akun_list=akun_list)

@app.route('/transaksi/<int:id>')
@transaction_read_required
def transaksi_detail(id):
    conn = db()
    j = conn.execute('SELECT * FROM jurnal WHERE id=?',(id,)).fetchone()
    if not j:
        conn.close()
        flash('Tidak ditemukan','danger')
        return redirect(url_for('transaksi_list'))
    lines = conn.execute("""
        SELECT dj.*,a.kode,a.nama,a.tipe FROM detail_jurnal dj
        JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=?
    """,(id,)).fetchall()
    items = _enrich_item_varian(conn, conn.execute("SELECT * FROM transaksi_item WHERE jurnal_id=? ORDER BY id", (id,)).fetchall())
    has_meta = bool((j['tx_meta'] if 'tx_meta' in j.keys() else '') or '')
    meta = _invoice_tx_meta(j)
    linked_invoice = conn.execute("SELECT * FROM invoice WHERE jurnal_id=?", (id,)).fetchone()
    # Info POS (kasir & nomor struk) jika transaksi ini berasal dari POS
    pos_info = conn.execute(
        "SELECT ps.nomor_struk, ps.metode_bayar, ps.dibuat, "
        "       COALESCE(NULLIF(u.nama,''), psh.username) AS kasir_nama, psh.username AS kasir_username "
        "FROM pos_sale ps "
        "LEFT JOIN pos_shift psh ON psh.id = ps.shift_id "
        "LEFT JOIN users u       ON u.id  = psh.user_id "
        "WHERE ps.jurnal_id = ?", (id,)
    ).fetchone()
    proyek_info = None
    _pid = j['proyek_id'] if 'proyek_id' in j.keys() else None
    if _pid:
        proyek_info = conn.execute("SELECT id, kode, nama FROM proyek WHERE id=?", (_pid,)).fetchone()
    # Konsumsi bahan backflush resep (Penjualan Umum) — ditampilkan terpisah dari rincian menu.
    resep_bahan = conn.execute(
        "SELECT * FROM penjualan_resep_bahan WHERE jurnal_id=? ORDER BY line_idx, id", (id,)
    ).fetchall()
    kunci = _info_kunci_transaksi(conn, id)
    conn.close()
    # Peran pihak SENGAJA disimpulkan di sini, bukan di template. Kolom
    # `jurnal.pihak` cuma menyimpan NAMA; arah transaksinya ada di meta/tipe_tx.
    # Kalau arahnya tidak jelas (pelunasan, transfer, jurnal umum), dipakai label
    # netral 'PIHAK'. Lebih baik tidak menyebut peran daripada salah menyebut
    # pelanggan sebagai vendor.
    pihak_nama = ((j['pihak'] if 'pihak' in j.keys() else '') or '').strip()
    _mt = meta.get('tipe')
    if _mt == 'JUAL' or j['tipe_tx'] in ('PEMASUKAN', 'RETUR_JUAL'):
        pihak_peran = 'PELANGGAN'
    elif _mt in ('BELI', 'PO') or j['tipe_tx'] in ('PENGELUARAN', 'RETUR_BELI'):
        pihak_peran = 'VENDOR'
    else:
        pihak_peran = 'PIHAK'
    return render_template('transaksi_detail.html', j=j, lines=lines, items=items,
                           has_meta=has_meta, is_sale=meta.get('tipe') == 'JUAL',
                           pihak_nama=pihak_nama, pihak_peran=pihak_peran,
                           linked_invoice=linked_invoice, pos_info=pos_info,
                           proyek_info=proyek_info, resep_bahan=resep_bahan,
                           ref_no_invoice=meta.get('no_invoice_supplier'),
                           ref_no_faktur=meta.get('no_faktur_pajak'), kunci=kunci)


@app.route('/transaksi/<int:id>/edit-item', methods=['GET','POST'])
@operator_required
def transaksi_edit_item(id):
    conn = db()
    j = conn.execute("SELECT * FROM jurnal WHERE id=?", (id,)).fetchone()
    if not j:
        conn.close(); flash('Transaksi tidak ditemukan.', 'danger')
        return redirect(url_for('transaksi_list'))
    try:
        meta = json.loads((j['tx_meta'] if 'tx_meta' in j.keys() else '') or '{}')
    except Exception:
        meta = {}
    if not meta:
        conn.close()
        flash('Transaksi ini tidak punya rincian item yang dapat diedit. Gunakan Edit Jurnal.', 'warning')
        return redirect(url_for('transaksi_edit', id=id))
    if session.get('role') == 'OPERATOR' and not _operator_date_ok(j['tanggal']):
        conn.close()
        flash('Operator hanya bisa mengedit transaksi bulan berjalan.', 'warning')
        return redirect(url_for('transaksi_detail', id=id))
    # GUARD backflush resep (GET+POST): _build_sale saat rebuild HANYA baca transaksi_item
    # (menu-only) + tx_meta.hpp_generik yang tak memuat resep → edit inline akan membuang konsumsi
    # bahan & meng-nol-kan COGS. Blokir sejak GET agar form tak muncul; ubah lewat hapus + input ulang.
    if conn.execute("SELECT 1 FROM penjualan_resep_bahan WHERE jurnal_id=? LIMIT 1", (id,)).fetchone():
        conn.close()
        flash('Transaksi ini memakai resep produksi (backflush bahan). Untuk mengubahnya, hapus lalu input ulang agar konsumsi bahan & HPP tetap konsisten.', 'warning')
        return redirect(url_for('transaksi_detail', id=id))
    # Kunci invoice/pembayaran dicek sejak GET. Dulu hanya saat Simpan: form tetap terbuka,
    # user mengisi semuanya, lalu baru ditolak.
    if request.method == 'GET':
        _pesan_kunci = _pesan_transaksi_terkunci(conn, id)
        if _pesan_kunci:
            conn.close(); flash(_pesan_kunci, 'warning')
            return redirect(url_for('transaksi_detail', id=id))

    def _pf(v):
        try: return parse_qty(v)
        except (ValueError, TypeError): return 0.0

    if request.method == 'POST':
        tanggal    = request.form.get('tanggal', j['tanggal'])
        keterangan = request.form.get('keterangan', j['keterangan']) or j['keterangan']
        pihak      = (request.form.get('pihak', '') or '').strip()
        if not _operator_date_ok(tanggal):
            conn.close()
            flash('Tanggal transaksi tidak valid. Operator hanya bisa mengedit transaksi untuk bulan berjalan.', 'warning')
            return redirect(url_for('transaksi_detail', id=id))
        pids  = request.form.getlist('item_produk_id[]')
        descs = request.form.getlist('item_deskripsi[]')
        qtys  = request.form.getlist('item_qty[]')
        sats  = request.form.getlist('item_satuan[]')
        hrgs  = request.form.getlist('item_harga[]')
        diss  = request.form.getlist('item_diskon[]')
        items = []
        for i in range(len(qtys)):
            q = _pf(qtys[i])
            raw_h = hrgs[i] if i < len(hrgs) else ''
            raw_d = diss[i] if i < len(diss) else '0'
            pid = (pids[i].strip() if i < len(pids) and pids[i] else '')
            desc = (descs[i].strip() if i < len(descs) and descs[i] else '')
            if q <= 0 and not desc and not pid:
                continue
            if not str(raw_h).strip():
                conn.close(); flash(f'Harga item baris {i + 1} wajib diisi.', 'danger')
                return redirect(url_for('transaksi_edit_item', id=id))
            try:
                h = parse_nonnegative_rp(raw_h, f'Harga item baris {i + 1}')
                d = parse_nonnegative_rp(raw_d, f'Diskon item baris {i + 1}')
            except ValueError as ex:
                conn.close(); flash(str(ex), 'danger')
                return redirect(url_for('transaksi_edit_item', id=id))
            if q <= 0:
                conn.close(); flash(f'Qty item baris {i + 1} harus lebih dari 0.', 'danger')
                return redirect(url_for('transaksi_edit_item', id=id))
            if d > q * h:
                conn.close(); flash(f'Diskon item baris {i + 1} tidak boleh melebihi subtotal bruto.', 'danger')
                return redirect(url_for('transaksi_edit_item', id=id))
            try:
                produk_id = int(pid) if pid else None
            except ValueError:
                conn.close(); flash(f'Produk item baris {i + 1} tidak valid.', 'danger')
                return redirect(url_for('transaksi_edit_item', id=id))
            if produk_id and not conn.execute(
                    "SELECT 1 FROM produk WHERE id=? AND COALESCE(aktif,1)=1",
                    (produk_id,)
            ).fetchone():
                conn.close(); flash(f'Produk item baris {i + 1} tidak ditemukan atau sudah diarsipkan.', 'danger')
                return redirect(url_for('transaksi_edit_item', id=id))
            items.append({'produk_id': produk_id, 'deskripsi': desc,
                          'qty': q, 'satuan': (sats[i].strip() if i < len(sats) else ''),
                          'harga': h, 'diskon': d})
        if not items:
            conn.close(); flash('Minimal 1 item harus diisi.', 'danger')
            return redirect(url_for('transaksi_edit_item', id=id))

        new_meta = dict(meta)
        new_meta['akun_kas'] = request.form.get('akun_kas', meta.get('akun_kas', '1100'))
        if not is_rekening_kode(conn, new_meta['akun_kas']):
            conn.close(); flash('Rekening kas/bank tidak valid.', 'danger')
            return redirect(url_for('transaksi_edit_item', id=id))
        new_meta['pihak'] = pihak
        new_meta['jt'] = request.form.get('jatuh_tempo') or None
        if not is_valid_optional_iso_date(new_meta['jt']):
            conn.close(); flash('Tanggal jatuh tempo tidak valid.', 'danger')
            return redirect(url_for('transaksi_edit_item', id=id))
        try:
            if meta.get('tipe') == 'JUAL':
                new_meta['uang_masuk']  = parse_nonnegative_rp(request.form.get('uang_masuk', '0'), 'Uang masuk')
                new_meta['diskon']      = parse_nonnegative_rp(request.form.get('diskon', '0'), 'Diskon')
                new_meta['ongkir']      = parse_nonnegative_rp(request.form.get('ongkir', '0'), 'Ongkir')
                new_meta['biaya']       = parse_nonnegative_rp(request.form.get('biaya', '0'), 'Biaya lain')
                new_meta['hpp_generik'] = parse_nonnegative_rp(request.form.get('hpp_generik', '0'), 'HPP non-SKU')
            else:
                new_meta['uang_keluar'] = parse_nonnegative_rp(request.form.get('uang_keluar', '0'), 'Uang keluar')
        except ValueError as ex:
            conn.close(); flash(str(ex), 'danger')
            return redirect(url_for('transaksi_edit_item', id=id))

        # Reverse efek lama → bangun ulang dari item baru (jurnal+stok+saldo sinkron)
        _pesan_kunci = _pesan_transaksi_terkunci(conn, id)
        if _pesan_kunci:
            conn.close(); flash(_pesan_kunci, 'warning')
            return redirect(url_for('transaksi_detail', id=id))
        # (Guard backflush resep sudah ditegakkan di awal fungsi, mencakup GET+POST.)
        # Guard HPP DILONGGARKAN (V5.3): koreksi historis TIDAK diblokir, hanya mengumpulkan
        # info adanya mutasi stok lebih baru. Dengan engine COGS-terkunci, COGS tiap transaksi
        # tetap di jurnalnya; stok & rata-rata HPP barang sisa otomatis disesuaikan (_reverse_tx
        # -> _resync_product_cost). Guard pembayaran & invoice di atas TETAP memblokir (integritas tagihan).
        hpp_warnings = []
        if has_later_stock_movements(conn, id):
            hpp_warnings.append('ada mutasi stok yang lebih baru untuk produk yang sama')
        if has_stock_movement_after_date(
                conn, id, [it.get('produk_id') for it in items], tanggal):
            hpp_warnings.append('tanggal transaksi melewati mutasi stok SKU lain yang sudah tercatat')
        old_total = conn.execute(
            "SELECT COALESCE(SUM(debit),0) t FROM detail_jurnal WHERE jurnal_id=?", (id,)
        ).fetchone()['t']
        try:
            _reverse_tx(conn, id)
            conn.execute("UPDATE jurnal SET tanggal=?, keterangan=?, pihak=? WHERE id=?",
                         (tanggal, keterangan, pihak, id))
            if meta.get('tipe') == 'JUAL':
                _build_sale(conn, id, tanggal, keterangan, items, new_meta)
            else:
                _build_purchase(conn, id, tanggal, keterangan, items, new_meta)
        except ValueError as ex:
            conn.rollback(); conn.close()
            flash(str(ex), 'danger')
            return redirect(url_for('transaksi_edit_item', id=id))
        conn.execute("UPDATE jurnal SET tx_meta=? WHERE id=?", (json.dumps(new_meta), id))
        new_total = conn.execute(
            "SELECT COALESCE(SUM(debit),0) t FROM detail_jurnal WHERE jurnal_id=?", (id,)
        ).fetchone()['t']
        linked_invoice = conn.execute("SELECT * FROM invoice WHERE jurnal_id=?", (id,)).fetchone()
        if linked_invoice:
            _sync_invoice_from_transaction(conn, linked_invoice['id'], id)
        log_detail = _edit_detail(keterangan, j, tanggal, old_total, new_total)
        if hpp_warnings:
            log_detail += ' [PERINGATAN HPP: koreksi historis]'
        add_log(conn, 'Edit rincian transaksi', log_detail, 'INPUT')
        bal_warn = _balance_warning(conn)   # baca sebelum close (lihat state belum-commit)
        conn.commit(); conn.close()
        flash('Rincian diperbarui, jurnal, stok, dan saldo otomatis disesuaikan.', 'success')
        if hpp_warnings:
            flash('Catatan: koreksi diproses meski ' + ' dan '.join(hpp_warnings) +
                  '. COGS transaksi lain tetap terkunci di jurnalnya; stok dan rata-rata HPP '
                  'barang yang tersisa otomatis disesuaikan dari sisa.', 'info')
        if bal_warn:
            flash(bal_warn, 'danger')
        return redirect(url_for('transaksi_detail', id=id))

    items = conn.execute("SELECT * FROM transaksi_item WHERE jurnal_id=? ORDER BY id", (id,)).fetchall()
    produk_list = conn.execute(
        "SELECT id,kode,nama,varian,satuan,stok,harga_beli,harga_jual FROM produk "
        "WHERE COALESCE(aktif,1)=1 ORDER BY nama").fetchall()
    akun_kas = conn.execute("SELECT * FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    warn_hpp = has_later_stock_movements(conn, id)   # V5.3: peringatan pra-edit (bukan blokir)
    conn.close()
    return render_template('transaksi_item_form.html', j=j, items=items, meta=meta,
                           produk_list=produk_list, akun_kas=akun_kas, warn_hpp=warn_hpp,
                           today=date.today().strftime('%Y-%m-%d'))

@app.route('/transaksi/<int:id>/edit', methods=['GET','POST'])
@finance_required
def transaksi_edit(id):
    conn = db()
    j = conn.execute('SELECT * FROM jurnal WHERE id=?', (id,)).fetchone()
    if not j:
        conn.close()
        flash('Jurnal tidak ditemukan.', 'danger')
        return redirect(url_for('transaksi_list'))

    has_meta = bool((j['tx_meta'] if 'tx_meta' in j.keys() else '') or '')
    if has_meta:
        conn.close()
        flash('Gunakan Edit Transaksi & Item agar rincian barang, stok, dan jurnal tetap sinkron.', 'info')
        return redirect(url_for('transaksi_edit_item', id=id))
    if (j['tipe_tx'] or 'JURNAL') != 'JURNAL':
        conn.close()
        flash('Transaksi terstruktur tidak dapat diedit sebagai jurnal manual. Hapus lalu catat ulang agar stok dan saldo tetap sinkron.', 'warning')
        return redirect(url_for('transaksi_detail', id=id))

    akun_list = conn.execute('SELECT * FROM akun ORDER BY kode').fetchall()

    if request.method == 'POST':
        tanggal    = request.form['tanggal']
        keterangan = request.form['keterangan']
        referensi  = request.form.get('referensi', '')
        kategori   = request.form.get('kategori', 'OPERASIONAL')
        akun_ids   = request.form.getlist('akun_id[]')
        debits     = request.form.getlist('debit[]')
        kredits    = request.form.getlist('kredit[]')

        if not is_valid_iso_date(tanggal):
            flash('Tanggal jurnal tidak valid.', 'warning')
            lines = conn.execute("""
                SELECT dj.*,a.kode,a.nama FROM detail_jurnal dj
                JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=?
            """, (id,)).fetchall()
            conn.close()
            return render_template('transaksi_form.html', akun_list=akun_list,
                                   edit=j, lines=lines, items=[],
                                   today=date.today().strftime('%Y-%m-%d'))
        if kategori not in MANUAL_JOURNAL_CATEGORIES:
            flash('Kategori arus kas jurnal tidak valid.', 'danger')
            lines = conn.execute("""
                SELECT dj.*,a.kode,a.nama FROM detail_jurnal dj
                JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=?
            """, (id,)).fetchall()
            conn.close()
            return render_template('transaksi_form.html', akun_list=akun_list,
                                   edit=j, lines=lines, items=[],
                                   today=date.today().strftime('%Y-%m-%d'))

        try:
            new_lines = parse_manual_journal_lines(conn, akun_ids, debits, kredits)
        except ValueError as ex:
            flash(str(ex), 'danger')
            lines = conn.execute("""
                SELECT dj.*,a.kode,a.nama FROM detail_jurnal dj
                JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=?
            """, (id,)).fetchall()
            conn.close()
            return render_template('transaksi_form.html', akun_list=akun_list,
                                   edit=j, lines=lines, items=[],
                                   today=date.today().strftime('%Y-%m-%d'))

        old_total = conn.execute(
            "SELECT COALESCE(SUM(debit),0) t FROM detail_jurnal WHERE jurnal_id=?", (id,)
        ).fetchone()['t']
        new_total = sum(debit for _akun, debit, _kredit in new_lines)
        conn.execute(
            'UPDATE jurnal SET tanggal=?,keterangan=?,referensi=?,kategori=? WHERE id=?',
            (tanggal, keterangan, referensi, kategori, id)
        )
        conn.execute('DELETE FROM detail_jurnal WHERE jurnal_id=?', (id,))
        conn.executemany(
            'INSERT INTO detail_jurnal(jurnal_id,akun_id,debit,kredit) VALUES(?,?,?,?)',
            [(id, akun_id, debit, kredit) for akun_id, debit, kredit in new_lines]
        )
        add_log(conn, 'Edit jurnal', _edit_detail(keterangan, j, tanggal, old_total, new_total), 'INPUT')
        conn.commit(); conn.close()
        flash('Jurnal berhasil diupdate!', 'success')
        return redirect(url_for('transaksi_detail', id=id))

    lines = conn.execute("""
        SELECT dj.*,a.kode,a.nama FROM detail_jurnal dj
        JOIN akun a ON a.id=dj.akun_id WHERE dj.jurnal_id=?
    """, (id,)).fetchall()
    items = conn.execute(
        "SELECT * FROM transaksi_item WHERE jurnal_id=? ORDER BY id", (id,)
    ).fetchall()
    conn.close()
    return render_template('transaksi_form.html', akun_list=akun_list,
                           edit=j, lines=lines, items=items,
                           today=date.today().strftime('%Y-%m-%d'))


def _archive_transaction(conn, jid):
    """Snapshot lengkap transaksi (jurnal + detail + item + meta) ke arsip_transaksi
    SEBELUM dihapus, supaya bisa dipulihkan. detail_jurnal disimpan pakai KODE akun
    (bukan id) agar restore tahan walau id akun berubah."""
    j = conn.execute("SELECT * FROM jurnal WHERE id=?", (jid,)).fetchone()
    if not j:
        return
    jd = {k: j[k] for k in j.keys()}
    dj = conn.execute("""SELECT a.kode AS akun_kode, d.debit, d.kredit
                         FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id
                         WHERE d.jurnal_id=? ORDER BY d.id""", (jid,)).fetchall()
    items = conn.execute("SELECT * FROM transaksi_item WHERE jurnal_id=? ORDER BY id", (jid,)).fetchall()
    total = conn.execute("SELECT COALESCE(SUM(debit),0) t FROM detail_jurnal WHERE jurnal_id=?", (jid,)).fetchone()['t']
    snap = {'jurnal': jd,
            'detail_jurnal': [dict(r) for r in dj],
            'items': [dict(r) for r in items]}
    conn.execute("""INSERT INTO arsip_transaksi
        (jurnal_id_lama, nomor_tx, tanggal, keterangan, tipe_tx, pihak, total, snapshot, dihapus_oleh)
        VALUES(?,?,?,?,?,?,?,?,?)""",
        (jid, jd.get('nomor_tx'), jd.get('tanggal'), jd.get('keterangan'),
         jd.get('tipe_tx'), jd.get('pihak'), float(total or 0),
         json.dumps(snap, default=str), session.get('username', '')))


def _archive_restorable(tipe_tx, snapshot_json):
    """Apakah arsip bisa dipulihkan OTOMATIS dengan setia (jurnal + stok + relasi)?
    Hanya tipe yang dibangun lewat jalur _build_sale/_build_purchase (penjualan &
    pembelian/beban terstruktur), transfer rekening, dan jurnal manual murni.
    Tipe lain (retur, PO, aset, gaji, produksi, pengeluaran non-editable) → pulihkan
    manual, karena stok/relasi pendukungnya tidak ikut ter-rebuild dan bisa bikin
    stok/Neraca desync. Return (bool, alasan_jika_tidak)."""
    try:
        snap = json.loads(snapshot_json or '{}')
    except Exception:
        return False, 'Snapshot rusak.'
    jr = snap.get('jurnal') or {}
    try:
        meta = json.loads(jr.get('tx_meta') or '{}')
    except Exception:
        meta = {}
    t = (meta.get('tipe') or '').upper()
    # Penjualan ber-resep (backflush bahan): snapshot TANPA subledger penjualan_resep_bahan →
    # rebuild via _build_sale akan menghasilkan revenue TANPA konsumsi bahan (HPP=0, stok tak turun)
    # → Neraca/persediaan desync. Blokir pemulihan otomatis; input ulang manual.
    if t == 'JUAL' and meta.get('has_resep'):
        return False, ('Penjualan dengan resep produksi (backflush bahan) belum bisa dipulihkan '
                       'otomatis dari arsip karena konsumsi bahan tidak ikut ter-rebuild. Catat ulang manual bila perlu.')
    if t in ('JUAL', 'BELI', 'TRANSFER_REKENING'):
        return True, ''
    if not meta and (tipe_tx or 'JURNAL') == 'JURNAL':
        return True, ''
    label = {'RETUR_JUAL': 'retur penjualan', 'RETUR_BELI': 'retur pembelian',
             'PRODUKSI': 'produksi'}.get((tipe_tx or '').upper(), (tipe_tx or 'tipe khusus'))
    return False, ('Pemulihan otomatis untuk ' + label + ' belum didukung '
                   '(stok/relasi pendukung tidak ikut ter-restore). Catat ulang manual bila perlu.')


def _restore_transaction(conn, arsip_id):
    """Pulihkan transaksi dari arsip. Transaksi terstruktur (punya tx_meta + item)
    dibangun ulang via _build_sale/_build_purchase (HPP memakai rata-rata SAAT INI,
    konsisten dgn perilaku edit). Jurnal manual di-insert ulang dari snapshot.
    Return (jurnal_id_baru, error_str)."""
    a = conn.execute("SELECT * FROM arsip_transaksi WHERE id=?", (arsip_id,)).fetchone()
    if not a:
        return None, 'Arsip tidak ditemukan.'
    if a['restored']:
        return None, 'Arsip ini sudah dipulihkan sebelumnya.'
    ok, why = _archive_restorable(a['tipe_tx'], a['snapshot'])
    if not ok:
        return None, why
    try:
        snap = json.loads(a['snapshot'] or '{}')
    except Exception:
        return None, 'Snapshot arsip rusak, tidak bisa dipulihkan.'
    jr = snap.get('jurnal', {}) or {}
    tanggal = jr.get('tanggal'); keterangan = jr.get('keterangan') or '(tanpa keterangan)'
    pihak = jr.get('pihak') or ''; kategori = jr.get('kategori') or 'OPERASIONAL'
    tipe_tx = jr.get('tipe_tx') or 'JURNAL'
    try:
        meta = json.loads(jr.get('tx_meta') or '{}')
    except Exception:
        meta = {}
    items = snap.get('items', []) or []
    if meta and items:
        # Transaksi terstruktur → bangun ulang lewat jalur build (stok + jurnal sinkron).
        nomor = next_nomor_tx(conn, tanggal)
        jid = conn.execute(
            "INSERT INTO jurnal(tanggal,keterangan,kategori,tipe_tx,pihak,nomor_tx) VALUES(?,?,?,?,?,?)",
            (tanggal, keterangan, kategori, tipe_tx, pihak, nomor)).lastrowid
        bitems = [dict(produk_id=it.get('produk_id'), deskripsi=it.get('deskripsi'),
                       qty=it.get('qty'), satuan=it.get('satuan'), harga=it.get('harga'),
                       diskon=it.get('diskon') or 0, arah=it.get('arah')) for it in items]
        if meta.get('tipe') == 'JUAL':
            _build_sale(conn, jid, tanggal, keterangan, bitems, meta)
        else:
            _build_purchase(conn, jid, tanggal, keterangan, bitems, meta)
        conn.execute("UPDATE jurnal SET tx_meta=? WHERE id=?", (json.dumps(meta), jid))
        if jr.get('proyek_id'):
            try:
                conn.execute("UPDATE jurnal SET proyek_id=? WHERE id=?", (jr.get('proyek_id'), jid))
            except Exception:
                pass
    else:
        # Jurnal manual → insert ulang baris dari snapshot (tanpa stok).
        entries = [(d.get('akun_kode'), d.get('debit', 0), d.get('kredit', 0))
                   for d in snap.get('detail_jurnal', []) if d.get('akun_kode')]
        if not entries:
            return None, 'Snapshot tidak punya rincian jurnal untuk dipulihkan.'
        jid = insert_jurnal(conn, tanggal, keterangan, kategori, tipe_tx, entries, pihak=pihak)
    conn.execute("UPDATE arsip_transaksi SET restored=1, restored_pada=CURRENT_TIMESTAMP, restored_jurnal_id=? WHERE id=?",
                 (jid, arsip_id))
    return jid, None


@app.route('/transaksi/<int:id>/hapus', methods=['GET', 'POST'])
@finance_required
def transaksi_hapus(id):
    conn = db()
    j = conn.execute("SELECT * FROM jurnal WHERE id=?", (id,)).fetchone()
    if not j:
        conn.close(); flash('Transaksi tidak ditemukan.', 'danger')
        return redirect(url_for('transaksi_list'))
    detail = f"{j['keterangan']} | {j['tanggal']}"
    # Blokir KERAS yang DIPERTAHANKAN (integritas tagihan, bukan presisi HPP).
    block = None
    if transaction_dependency_count(conn, id):
        block = ('Transaksi ini punya pembayaran atau retur terkait. Hapus jurnal pelunasan/retur '
                 'terkait lebih dulu agar saldo tetap sinkron.')
    elif pinjaman_karyawan_repayment_count(conn, id):
        block = ('Pinjaman karyawan ini sudah punya catatan pengembalian. Hapus dulu transaksi '
                 'Pengembalian pinjaman karyawan terkait di Daftar Transaksi agar saldo '
                 'Piutang Karyawan tetap sinkron.')
    elif has_untracked_legacy_return_adjustment(conn, id):
        block = ('Retur historis ini belum memiliki relasi tracker yang lengkap. Lakukan koreksi manual '
                 'agar piutang/hutang tetap sinkron.')
    elif non_sku_return_overdraw_on_delete(conn, id):
        # Sebutkan retur MANA yang menghalangi. Tanpa ini user cuma dapat perintah
        # "hapus retur terkait" tanpa tahu harus mencari ke mana.
        _returs = non_sku_return_jurnals(conn)
        _daftar = ', '.join(
            f"{(r['nomor_tx'] or ('#' + str(r['id'])))} ({r['tanggal']})" for r in _returs[:5]
        )
        if len(_returs) > 5:
            _daftar += f", dan {len(_returs) - 5} lainnya"
        block = ('Konsumsi persediaan non-SKU transaksi ini sudah dikembalikan lewat retur penjualan '
                 '(nilai HPP). Hapus retur tersebut lebih dulu lewat menu Retur atau Daftar Transaksi '
                 'agar nilai Persediaan (1130) tidak menggantung.'
                 + (f' Retur yang dimaksud: {_daftar}.' if _daftar else ''))
    else:
        li = _locked_transaction_invoice(conn, id)
        if li:
            block = (f'Invoice {li["nomor"]} berstatus {_LABEL_STATUS_INVOICE.get(li["status"], li["status"])}. '
                     'Invoice harus kembali ke Draft dulu sebelum transaksinya dihapus.')
    hpp_warn = has_later_stock_movements(conn, id)

    if request.method == 'GET':
        # HALAMAN KONFIRMASI: tampilkan ringkasan + peringatan SEBELUM benar-benar dihapus.
        items = _enrich_item_varian(conn, conn.execute("SELECT * FROM transaksi_item WHERE jurnal_id=? ORDER BY id", (id,)).fetchall())
        total = conn.execute("SELECT COALESCE(SUM(debit),0) t FROM detail_jurnal WHERE jurnal_id=?", (id,)).fetchone()['t']
        kunci = _info_kunci_transaksi(conn, id) if block else None
        conn.close()
        return render_template('transaksi_hapus_konfirmasi.html', j=j, items=items, total=total,
                               block=block, hpp_warn=hpp_warn, next=request.args.get('next', ''),
                               kunci=kunci)

    # POST → arsipkan + reverse + hapus (hanya jika tidak diblokir)
    if block:
        conn.close(); flash(block, 'warning')
        return redirect(url_for('transaksi_detail', id=id))
    _archive_transaction(conn, id)
    conn.execute("DELETE FROM invoice WHERE jurnal_id=? AND status='DRAFT'", (id,))
    _reverse_tx(conn, id)
    conn.execute('DELETE FROM jurnal WHERE id=?', (id,))
    if hpp_warn:
        detail += ' [PERINGATAN HPP: koreksi historis]'
    add_log(conn, 'Hapus transaksi (arsip)', detail, 'INPUT')
    bal_warn = _balance_warning(conn)   # baca sebelum close (lihat state belum-commit)
    conn.commit(); conn.close()
    next_pemasukan = (request.form.get('next') or '').strip() == 'pemasukan'
    flash('Transaksi dibatalkan & diarsipkan. Stok dikembalikan & jurnal dihapus.' if next_pemasukan
          else 'Transaksi dihapus & dipindahkan ke Arsip. Bisa dipulihkan kapan saja.', 'warning')
    if hpp_warn:
        flash('Catatan: transaksi yang dihapus punya mutasi stok yang lebih baru. COGS transaksi '
              'lain tetap terkunci; stok dan rata-rata HPP barang sisa otomatis disesuaikan. '
              'Ingat: menghapus pembelian yang barangnya sudah terjual dapat membuat stok minus.', 'warning')
    if bal_warn:
        flash(bal_warn, 'danger')
    return redirect(url_for('pemasukan', restore=1) if next_pemasukan else url_for('transaksi_list'))


@app.route('/transaksi/<int:id>/buka-kunci', methods=['GET', 'POST'])
@finance_required
def transaksi_buka_kunci(id):
    """Buka kunci transaksi supaya bisa dikoreksi: invoice kembali ke Draft dan pembayaran atas
    piutangnya dibatalkan (jurnal diarsipkan lalu dihapus, sama dengan menghapusnya satu per
    satu dari Daftar Transaksi). Retur/penyesuaian dan write-off TIDAK dibatalkan otomatis
    karena ikut mengubah stok atau beban; selama masih ada, kunci tidak dibuka.
    GET = halaman konfirmasi yang merinci semuanya, POST = jalankan."""
    conn = db()
    j = conn.execute("SELECT * FROM jurnal WHERE id=?", (id,)).fetchone()
    if not j:
        conn.close(); flash('Transaksi tidak ditemukan.', 'danger')
        return redirect(url_for('transaksi_list'))
    info = _info_kunci_transaksi(conn, id)
    if request.method == 'GET':
        total = conn.execute("SELECT COALESCE(SUM(debit),0) t FROM detail_jurnal WHERE jurnal_id=?",
                             (id,)).fetchone()['t']
        conn.close()
        return render_template('transaksi_buka_kunci.html', j=j, info=info, total=total,
                               label_status=_LABEL_STATUS_INVOICE)
    if not info['terkunci']:
        conn.close(); flash('Transaksi ini tidak sedang terkunci invoice atau pembayaran.', 'info')
        return redirect(url_for('transaksi_detail', id=id))
    if info['penghalang']:
        conn.close()
        flash('Kunci belum bisa dibuka. Hapus dulu catatan yang tercantum di kotak merah, '
              'lalu coba lagi.', 'warning')
        return redirect(url_for('transaksi_buka_kunci', id=id))
    inv = info['invoice']
    try:
        for b in info['pembayaran']:
            _archive_transaction(conn, b['jurnal_id'])
            _reverse_tx(conn, b['jurnal_id'])   # bayar_piutang, terbayar, status, label invoice
            conn.execute("DELETE FROM jurnal WHERE id=?", (b['jurnal_id'],))
        conn.execute("UPDATE invoice SET status='DRAFT' WHERE jurnal_id=? AND status!='DRAFT'", (id,))
        catat = []
        if inv:
            catat.append(f"invoice {inv['nomor']} {inv['status']} -> DRAFT")
        if info['pembayaran']:
            catat.append(f"{len(info['pembayaran'])} pembayaran dibatalkan {rp_filter(info['total_bayar'])} ("
                         + ', '.join(b['nomor'] for b in info['pembayaran']) + ')')
        add_log(conn, 'Buka kunci transaksi',
                f"{j['nomor_tx'] or ('#' + str(id))} | " + ' | '.join(catat), 'INPUT')
        bal_warn = _balance_warning(conn)
        conn.commit()
    except Exception as ex:
        conn.rollback(); conn.close()
        flash(f'Gagal membuka kunci, tidak ada yang diubah: {ex}', 'danger')
        return redirect(url_for('transaksi_buka_kunci', id=id))
    conn.close()
    pesan = ['Kunci dibuka.']
    if inv:
        pesan.append(f"Invoice {inv['nomor']} kembali ke Draft.")
    if info['pembayaran']:
        pesan.append(f"{len(info['pembayaran'])} pembayaran senilai {rp_filter(info['total_bayar'])} "
                     'dibatalkan, jejaknya ada di Arsip Transaksi.')
    pesan.append('Transaksi ini sekarang bisa diedit atau dihapus.')
    if info['pembayaran']:
        pesan.append('Kalau pembayarannya memang sudah diterima, catat lagi sesudah koreksi selesai.')
    flash(' '.join(pesan), 'success')
    if bal_warn:
        flash(bal_warn, 'danger')
    return redirect(url_for('transaksi_detail', id=id))


@app.route('/transaksi/arsip')
@finance_required
def transaksi_arsip():
    conn = db()
    page = _page_arg()
    n_rows = conn.execute("SELECT COUNT(*) FROM arsip_transaksi").fetchone()[0]
    pg = paginate_meta(page, n_rows)
    raw = conn.execute("SELECT * FROM arsip_transaksi ORDER BY dihapus_pada DESC, id DESC LIMIT ? OFFSET ?",
                       (pg['per_page'], pg['offset'])).fetchall()
    conn.close()
    rows = []
    for r in raw:
        d = dict(r)
        d['can_restore'], d['restore_reason'] = _archive_restorable(r['tipe_tx'], r['snapshot'])
        rows.append(d)
    return render_template('transaksi_arsip.html', rows=rows, pg=pg)


@app.route('/transaksi/arsip/<int:id>/restore', methods=['POST'])
@finance_required
def transaksi_arsip_restore(id):
    conn = db()
    try:
        jid, err = _restore_transaction(conn, id)
    except ValueError as ex:
        conn.rollback(); conn.close()
        flash('Gagal memulihkan: ' + str(ex), 'danger')
        return redirect(url_for('transaksi_arsip'))
    if err:
        conn.close(); flash(err, 'warning')
        return redirect(url_for('transaksi_arsip'))
    add_log(conn, 'Pulihkan transaksi dari arsip', f'arsip #{id} -> jurnal #{jid}', 'INPUT')
    bal_warn = _balance_warning(conn)
    conn.commit(); conn.close()
    flash('Transaksi dipulihkan dari arsip (dibangun ulang dengan nomor baru; HPP memakai rata-rata saat ini).', 'success')
    if bal_warn:
        flash(bal_warn, 'danger')
    return redirect(url_for('transaksi_detail', id=jid))


@app.route('/transaksi/arsip/<int:id>/hapus-permanen', methods=['POST'])
@admin_required
def transaksi_arsip_hapus(id):
    conn = db()
    conn.execute("DELETE FROM arsip_transaksi WHERE id=?", (id,))
    add_log(conn, 'Hapus permanen arsip transaksi', f'arsip #{id}', 'INPUT')
    conn.commit(); conn.close()
    flash('Arsip dihapus permanen.', 'warning')
    return redirect(url_for('transaksi_arsip'))


# ---------- DATABASE CUSTOMER & VENDOR ----------
def _customer_metrics(conn, sd=None, ed=None):
    """dict nama→{omzet, hpp, profit} dari jurnal tipe PEMASUKAN, opsional filter tanggal."""
    where = "j.tipe_tx='PEMASUKAN' AND IFNULL(j.pihak,'')<>''"
    params = []
    if sd:
        where += " AND j.tanggal >= ?"; params.append(sd)
    if ed:
        where += " AND j.tanggal <= ?"; params.append(ed)
    rows = conn.execute(f"""
        SELECT j.pihak AS nama,
               COALESCE(SUM(CASE WHEN a.tipe='PENDAPATAN' THEN d.kredit ELSE 0 END),0) AS omzet,
               COALESCE(SUM(CASE WHEN a.kode='5100' THEN d.debit ELSE 0 END),0) AS hpp
        FROM jurnal j
        JOIN detail_jurnal d ON d.jurnal_id=j.id
        JOIN akun a ON a.id=d.akun_id
        WHERE {where}
        GROUP BY j.pihak
    """, params).fetchall()
    return {r['nama']: {'omzet': r['omzet'], 'hpp': r['hpp'],
                        'profit': r['omzet'] - r['hpp']} for r in rows}

def _piutang_metrics(conn, today):
    """dict nama→{jt (overdue), sisa_total}."""
    rows = conn.execute("""
        SELECT pelanggan AS nama,
               COALESCE(SUM(CASE WHEN status!='LUNAS' AND jatuh_tempo<? THEN jumlah-terbayar ELSE 0 END),0) AS jt,
               COALESCE(SUM(CASE WHEN status!='LUNAS' THEN jumlah-terbayar ELSE 0 END),0) AS sisa
        FROM piutang GROUP BY pelanggan
    """, (today,)).fetchall()
    return {r['nama']: {'jt': r['jt'], 'sisa': r['sisa']} for r in rows}

def _vendor_metrics(conn):
    """dict nama→{belanja} dari jurnal tipe PENGELUARAN."""
    # 1160 (Piutang Karyawan) dikecualikan: pinjaman karyawan bukan "belanja" dari vendor,
    # dan pihak-nya bisa kebetulan sama nama dgn vendor -> jangan cemari metrik belanja.
    rows = conn.execute("""
        SELECT j.pihak AS nama,
               COALESCE(SUM(CASE WHEN a.is_rekening=0 AND a.tipe!='LIABILITAS' AND a.kode!='1160' THEN d.debit ELSE 0 END),0) AS belanja
        FROM jurnal j
        JOIN detail_jurnal d ON d.jurnal_id=j.id
        JOIN akun a ON a.id=d.akun_id
        WHERE j.tipe_tx IN ('PENGELUARAN','PO_DITERIMA') AND IFNULL(j.pihak,'')<>''
        GROUP BY j.pihak
    """).fetchall()
    return {r['nama']: {'belanja': r['belanja']} for r in rows}

def _hutang_metrics(conn, today):
    rows = conn.execute("""
        SELECT pemasok AS nama,
               COALESCE(SUM(CASE WHEN status!='LUNAS' AND jatuh_tempo<? THEN jumlah-terbayar ELSE 0 END),0) AS jt,
               COALESCE(SUM(CASE WHEN status!='LUNAS' THEN jumlah-terbayar ELSE 0 END),0) AS sisa
        FROM hutang GROUP BY pemasok
    """, (today,)).fetchall()
    return {r['nama']: {'jt': r['jt'], 'sisa': r['sisa']} for r in rows}


@app.route('/database/customer')
@investor_required
def db_customer():
    conn = db()
    today = date.today().strftime('%Y-%m-%d')
    sort = request.args.get('sort', 'abjad')
    sd   = request.args.get('sd', '')
    ed   = request.args.get('ed', '')
    custs = [dict(r) for r in conn.execute("SELECT * FROM customer ORDER BY nama").fetchall()]
    met = _customer_metrics(conn, sd or None, ed or None)
    piu = _piutang_metrics(conn, today)
    for c in custs:
        m = met.get(c['nama'], {})
        p = piu.get(c['nama'], {})
        c['omzet']  = m.get('omzet', 0)
        c['profit'] = m.get('profit', 0)
        c['piutang_jt'] = p.get('jt', 0)
        c['piutang_sisa'] = p.get('sisa', 0)
    keyf = {'abjad': lambda c: c['nama'].lower(),
            'omzet': lambda c: -c['omzet'],
            'profit': lambda c: -c['profit'],
            'jt': lambda c: -c['piutang_jt']}.get(sort, lambda c: c['nama'].lower())
    custs.sort(key=keyf)
    rank_key = 'profit' if sort == 'profit' else 'omzet'
    ranked = sorted([c for c in custs if c.get(rank_key, 0) > 0], key=lambda c: -c[rank_key])
    medal_rank = {c['id']: i+1 for i, c in enumerate(ranked[:3])}
    for c in custs:
        c['medal_rank'] = medal_rank.get(c['id'], 0)
    earliest = conn.execute(
        "SELECT MIN(tanggal) AS e FROM jurnal WHERE tipe_tx='PEMASUKAN'"
    ).fetchone()['e'] or today
    page = _page_arg()
    pg = paginate_meta(page, len(custs))           # sort/medal dari full set, slice utk tampilan
    custs = custs[pg['offset']:pg['offset'] + pg['per_page']]
    conn.close()
    return render_template('db_customer.html', custs=custs, sort=sort,
                           sd=sd, ed=ed, earliest_date=earliest, pg=pg)


@app.route('/database/customer/baru', methods=['POST'])
@finance_required
def db_customer_baru():
    conn = db()
    nama = (request.form.get('nama','') or '').strip()
    if not nama:
        conn.close(); flash('Nama customer wajib diisi.', 'warning')
        return redirect(url_for('db_customer'))
    try:
        conn.execute("INSERT INTO customer(nama,telepon,alamat,catatan) VALUES(?,?,?,?)",
                     (nama, request.form.get('telepon','').strip(),
                      request.form.get('alamat','').strip(), request.form.get('catatan','').strip()))
        conn.commit(); flash(f'Customer "{nama}" ditambahkan.', 'success')
    except sqlite3.IntegrityError:
        flash(f'Customer "{nama}" sudah ada.', 'warning')
    conn.close()
    return redirect(url_for('db_customer'))


@app.route('/database/customer/<int:id>')
@investor_required
def db_customer_detail(id):
    conn = db()
    today = date.today().strftime('%Y-%m-%d')
    c = conn.execute("SELECT * FROM customer WHERE id=?", (id,)).fetchone()
    if not c:
        conn.close(); flash('Customer tidak ditemukan.', 'danger')
        return redirect(url_for('db_customer'))
    nama = c['nama']
    met = _customer_metrics(conn).get(nama, {'omzet':0,'hpp':0,'profit':0})
    transaksi = conn.execute("""
        SELECT j.id, j.tanggal, j.keterangan,
               COALESCE(SUM(CASE WHEN a.tipe='PENDAPATAN' THEN d.kredit ELSE 0 END),0) AS total,
               COALESCE(SUM(CASE WHEN a.kode='5100' THEN d.debit ELSE 0 END),0) AS hpp
        FROM jurnal j
        LEFT JOIN detail_jurnal d ON d.jurnal_id=j.id
        LEFT JOIN akun a ON a.id=d.akun_id
        WHERE j.tipe_tx='PEMASUKAN' AND j.pihak=?
        GROUP BY j.id ORDER BY j.tanggal DESC, j.id DESC
    """, (nama,)).fetchall()
    belum = conn.execute("SELECT * FROM piutang WHERE pelanggan=? AND status!='LUNAS' ORDER BY jatuh_tempo", (nama,)).fetchall()
    lunas = conn.execute("SELECT * FROM piutang WHERE pelanggan=? AND status='LUNAS' ORDER BY tanggal DESC", (nama,)).fetchall()

    # monthly trend data
    monthly_raw = conn.execute("""
        SELECT strftime('%Y', j.tanggal) AS yr,
               strftime('%m', j.tanggal) AS mo,
               COALESCE(SUM(CASE WHEN a.tipe='PENDAPATAN' THEN d.kredit ELSE 0 END),0) AS omzet,
               COUNT(DISTINCT j.id) AS cnt
        FROM jurnal j
        LEFT JOIN detail_jurnal d ON d.jurnal_id=j.id
        LEFT JOIN akun a ON a.id=d.akun_id
        WHERE j.tipe_tx='PEMASUKAN' AND j.pihak=?
        GROUP BY yr, mo ORDER BY yr, mo
    """, (nama,)).fetchall()

    # build per-year dict: {year: [omzet_jan..omzet_dec]}, [cnt_jan..cnt_dec]}
    from collections import defaultdict
    year_omzet = defaultdict(lambda: [0]*12)
    year_cnt   = defaultdict(lambda: [0]*12)
    for r in monthly_raw:
        y, m = r['yr'], int(r['mo']) - 1
        year_omzet[r['yr']][m] = r['omzet']
        year_cnt[r['yr']][m]   = r['cnt']

    trend_years = sorted(year_omzet.keys(), reverse=True)

    # loyalty metrics
    from datetime import datetime as _dt
    loyalty = {'first': None, 'last': None, 'active_months': 0,
                'total_months': 0, 'consistency': 0, 'recency_days': 9999,
                'freq': len(transaksi), 'tier': 'Baru'}
    if monthly_raw:
        first_str = monthly_raw[0]['yr'] + '-' + monthly_raw[0]['mo']
        last_raw  = monthly_raw[-1]
        last_str  = last_raw['yr'] + '-' + last_raw['mo']
        loyalty['first'] = first_str
        loyalty['last']  = last_str
        active_months = len(monthly_raw)
        first_dt = _dt.strptime(first_str, '%Y-%m')
        today_dt = _dt.strptime(today, '%Y-%m-%d').replace(day=1)
        total_months = max(1, (today_dt.year - first_dt.year)*12 + (today_dt.month - first_dt.month) + 1)
        consistency = round(active_months / total_months * 100)
        # recency: last transaction date
        last_trx_date = transaksi[0]['tanggal'] if transaksi else None
        recency_days = 9999
        if last_trx_date:
            recency_days = (date.today() - date.fromisoformat(last_trx_date)).days
        loyalty.update({
            'active_months': active_months,
            'total_months': total_months,
            'consistency': consistency,
            'recency_days': recency_days,
        })
        if consistency >= 60 and recency_days <= 60:
            loyalty['tier'] = 'Pelanggan Setia'
        elif consistency >= 30 or recency_days <= 120:
            loyalty['tier'] = 'Pelanggan Reguler'
        elif recency_days > 365:
            loyalty['tier'] = 'Tidak Aktif'
        else:
            loyalty['tier'] = 'Jarang Beli'

    import json as _json
    trend_data = _json.dumps({
        yr: {'omzet': year_omzet[yr], 'cnt': year_cnt[yr]}
        for yr in trend_years
    })

    conn.close()
    return render_template('db_customer_detail.html', c=c, met=met,
                           transaksi=transaksi, belum=belum, lunas=lunas, today=today,
                           trend_years=trend_years, trend_data=trend_data, loyalty=loyalty)


@app.route('/database/customer/<int:id>/edit', methods=['POST'])
@finance_required
def db_customer_edit(id):
    conn = db()
    nama = (request.form.get('nama', '') or '').strip()
    if not nama:
        conn.close(); flash('Nama customer wajib diisi.', 'warning')
        return redirect(url_for('db_customer_detail', id=id))
    old = conn.execute("SELECT nama FROM customer WHERE id=?", (id,)).fetchone()
    if not old:
        conn.close(); flash('Customer tidak ditemukan.', 'danger')
        return redirect(url_for('db_customer'))
    try:
        conn.execute(
            "UPDATE customer SET nama=?,telepon=?,alamat=?,catatan=? WHERE id=?",
            (nama,
             (request.form.get('telepon', '') or '').strip(),
             (request.form.get('alamat', '') or '').strip(),
             (request.form.get('catatan', '') or '').strip(),
             id)
        )
        if nama != old['nama']:
            conn.execute("UPDATE jurnal SET pihak=? WHERE pihak=? AND tipe_tx='PEMASUKAN'", (nama, old['nama']))
            conn.execute("UPDATE jurnal SET pihak=? WHERE pihak=? AND tipe_tx IN ('WRITE_OFF_PIUTANG','JUAL')", (nama, old['nama']))
            # PELUNASAN dipakai bareng oleh piutang & hutang, jadi disamakan lewat jurnal_id
            # pembayaran piutang milik customer ini, bukan tebak dari teks pihak semata.
            conn.execute("""UPDATE jurnal SET pihak=? WHERE tipe_tx='PELUNASAN' AND id IN (
                SELECT jurnal_id FROM bayar_piutang WHERE jurnal_id IS NOT NULL
                  AND piutang_id IN (SELECT id FROM piutang WHERE pelanggan=?)
            )""", (nama, old['nama']))
            conn.execute("UPDATE jurnal SET pihak=? WHERE pihak=? AND tipe_tx='KOREKSI' AND keterangan LIKE 'Koreksi hapus piutang:%'", (nama, old['nama']))
            conn.execute("UPDATE piutang SET pelanggan=? WHERE pelanggan=?", (nama, old['nama']))
            conn.execute("UPDATE invoice SET pelanggan=? WHERE pelanggan=?", (nama, old['nama']))
        conn.commit()
        flash('Data customer diperbarui.', 'success')
    except sqlite3.IntegrityError:
        flash(f'Customer "{nama}" sudah ada.', 'warning')
    conn.close()
    return redirect(url_for('db_customer_detail', id=id))


@app.route('/database/customer/<int:id>/hapus', methods=['POST'])
@finance_required
def db_customer_hapus(id):
    conn = db()
    conn.execute("DELETE FROM customer WHERE id=?", (id,))
    conn.commit(); conn.close()
    flash('Customer dihapus dari database (data transaksi tetap utuh).', 'warning')
    return redirect(url_for('db_customer'))


@app.route('/database/vendor')
@investor_required
def db_vendor():
    conn = db()
    today = date.today().strftime('%Y-%m-%d')
    sort = request.args.get('sort', 'abjad')
    vends = [dict(r) for r in conn.execute("SELECT * FROM vendor ORDER BY nama").fetchall()]
    met = _vendor_metrics(conn)
    hut = _hutang_metrics(conn, today)
    for v in vends:
        m = met.get(v['nama'], {})
        h = hut.get(v['nama'], {})
        v['belanja'] = m.get('belanja', 0)
        v['hutang_jt'] = h.get('jt', 0)
        v['hutang_sisa'] = h.get('sisa', 0)
    keyf = {'abjad': lambda v: v['nama'].lower(),
            'belanja': lambda v: -v['belanja'],
            'jt': lambda v: -v['hutang_jt']}.get(sort, lambda v: v['nama'].lower())
    vends.sort(key=keyf)
    page = _page_arg()
    pg = paginate_meta(page, len(vends))
    vends = vends[pg['offset']:pg['offset'] + pg['per_page']]
    conn.close()
    return render_template('db_vendor.html', vends=vends, sort=sort, pg=pg)


@app.route('/database/vendor/baru', methods=['POST'])
@finance_required
def db_vendor_baru():
    conn = db()
    nama = (request.form.get('nama','') or '').strip()
    if not nama:
        conn.close(); flash('Nama vendor wajib diisi.', 'warning')
        return redirect(url_for('db_vendor'))
    try:
        conn.execute("INSERT INTO vendor(nama,telepon,alamat,catatan) VALUES(?,?,?,?)",
                     (nama, request.form.get('telepon','').strip(),
                      request.form.get('alamat','').strip(), request.form.get('catatan','').strip()))
        conn.commit(); flash(f'Vendor "{nama}" ditambahkan.', 'success')
    except sqlite3.IntegrityError:
        flash(f'Vendor "{nama}" sudah ada.', 'warning')
    conn.close()
    return redirect(url_for('db_vendor'))


@app.route('/database/vendor/<int:id>')
@investor_required
def db_vendor_detail(id):
    conn = db()
    today = date.today().strftime('%Y-%m-%d')
    v = conn.execute("SELECT * FROM vendor WHERE id=?", (id,)).fetchone()
    if not v:
        conn.close(); flash('Vendor tidak ditemukan.', 'danger')
        return redirect(url_for('db_vendor'))
    nama = v['nama']
    met = _vendor_metrics(conn).get(nama, {'belanja':0})
    # Belanja = Pengeluaran manual (PENGELUARAN) + penerimaan PO (PO_DITERIMA).
    # Nilai per baris memakai rumus yang sama dgn _vendor_metrics (hanya debit ke akun
    # non-kas / non-liabilitas / non-1160) supaya total daftar == "Total Belanja" di header.
    belanja_rows = [dict(r) for r in conn.execute("""
        SELECT j.id, j.id AS sort_id, j.tanggal, j.keterangan, j.kategori,
               'Pembelian' AS jenis,
               COALESCE(SUM(CASE WHEN a.is_rekening=0 AND a.tipe!='LIABILITAS' AND a.kode!='1160' THEN d.debit ELSE 0 END),0) AS total
        FROM jurnal j LEFT JOIN detail_jurnal d ON d.jurnal_id=j.id
                      LEFT JOIN akun a ON a.id=d.akun_id
        WHERE j.tipe_tx IN ('PENGELUARAN','PO_DITERIMA') AND j.pihak=?
          AND NOT EXISTS (SELECT 1 FROM detail_jurnal dx JOIN akun ax ON ax.id=dx.akun_id
                          WHERE dx.jurnal_id=j.id AND ax.kode='1160' AND dx.debit>0)
        GROUP BY j.id ORDER BY j.tanggal DESC, j.id DESC
    """, (nama,)).fetchall()]
    pembayaran_rows = [dict(r) for r in conn.execute("""
        SELECT j.id, b.id AS sort_id, COALESCE(j.tanggal,b.tanggal) AS tanggal,
               COALESCE(NULLIF(j.keterangan,''), NULLIF(b.catatan,''), 'Pelunasan hutang: ' || h.pemasok) AS keterangan,
               'PELUNASAN' AS kategori, 'Pembayaran Hutang' AS jenis, COALESCE(b.jumlah,0) AS total
        FROM bayar_hutang b
        JOIN hutang h ON h.id=b.hutang_id
        LEFT JOIN jurnal j ON j.id=b.jurnal_id
        WHERE h.pemasok=?
    """, (nama,)).fetchall()]
    transaksi = sorted(
        belanja_rows + pembayaran_rows,
        key=lambda r: (r.get('tanggal') or '', r.get('sort_id') or 0),
        reverse=True
    )
    belum = conn.execute("SELECT * FROM hutang WHERE pemasok=? AND status!='LUNAS' ORDER BY jatuh_tempo", (nama,)).fetchall()
    lunas = conn.execute("SELECT * FROM hutang WHERE pemasok=? AND status='LUNAS' ORDER BY tanggal DESC", (nama,)).fetchall()
    conn.close()
    return render_template('db_vendor_detail.html', v=v, met=met,
                           transaksi=transaksi, belanja_rows=belanja_rows,
                           belum=belum, lunas=lunas, today=today)


@app.route('/database/vendor/<int:id>/edit', methods=['POST'])
@finance_required
def db_vendor_edit(id):
    conn = db()
    nama = (request.form.get('nama', '') or '').strip()
    if not nama:
        conn.close(); flash('Nama vendor wajib diisi.', 'warning')
        return redirect(url_for('db_vendor_detail', id=id))
    old = conn.execute("SELECT nama FROM vendor WHERE id=?", (id,)).fetchone()
    if not old:
        conn.close(); flash('Vendor tidak ditemukan.', 'danger')
        return redirect(url_for('db_vendor'))
    try:
        conn.execute(
            "UPDATE vendor SET nama=?,telepon=?,alamat=?,catatan=? WHERE id=?",
            (nama,
             (request.form.get('telepon', '') or '').strip(),
             (request.form.get('alamat', '') or '').strip(),
             (request.form.get('catatan', '') or '').strip(),
             id)
        )
        if nama != old['nama']:
            # 'BELI' = jurnal Pajak > Pembelian PKP, juga membawa nama vendor. Tanpa ini
            # hutangnya berganti nama tapi jurnalnya tidak: satu vendor terpecah dua.
            conn.execute("UPDATE jurnal SET pihak=? WHERE pihak=? AND tipe_tx IN ('PENGELUARAN','PO_DITERIMA','BELI')", (nama, old['nama']))
            # PELUNASAN dipakai bareng oleh piutang & hutang, jadi disamakan lewat jurnal_id
            # pembayaran hutang milik vendor ini, bukan tebak dari teks pihak semata.
            conn.execute("""UPDATE jurnal SET pihak=? WHERE tipe_tx='PELUNASAN' AND id IN (
                SELECT jurnal_id FROM bayar_hutang WHERE jurnal_id IS NOT NULL
                  AND hutang_id IN (SELECT id FROM hutang WHERE pemasok=?)
            )""", (nama, old['nama']))
            conn.execute("UPDATE jurnal SET pihak=? WHERE pihak=? AND tipe_tx='KOREKSI' AND keterangan LIKE 'Koreksi hapus hutang:%'", (nama, old['nama']))
            conn.execute("UPDATE hutang SET pemasok=? WHERE pemasok=?", (nama, old['nama']))
            conn.execute("UPDATE purchase_order SET vendor=? WHERE vendor=?", (nama, old['nama']))
        conn.commit()
        flash('Data vendor diperbarui.', 'success')
    except sqlite3.IntegrityError:
        flash(f'Vendor "{nama}" sudah ada.', 'warning')
    conn.close()
    return redirect(url_for('db_vendor_detail', id=id))


@app.route('/database/vendor/<int:id>/hapus', methods=['POST'])
@finance_required
def db_vendor_hapus(id):
    conn = db()
    conn.execute("DELETE FROM vendor WHERE id=?", (id,))
    conn.commit(); conn.close()
    flash('Vendor dihapus dari database (data transaksi tetap utuh).', 'warning')
    return redirect(url_for('db_vendor'))


# ---------- MASTER HPP ----------
def hpp_component_value(b):
    tipe = (b.get('tipe') or 'BAHAN').upper()
    harga = float(b.get('harga') or 0)
    frekuensi = float(b.get('frekuensi') or 0)
    takaran = float(b.get('takaran') or 0)
    if tipe != 'BAHAN' and frekuensi <= 0:
        return harga
    if frekuensi > 0:
        return (harga / frekuensi) * takaran
    return 0.0

def hpp_bahan_breakdown(bahan):
    inventory = 0.0
    reclass = {}
    total = 0.0
    for b in bahan or []:
        try:
            nilai = hpp_component_value(b)
        except (TypeError, ValueError):
            continue
        if nilai <= 0 or not math.isfinite(nilai):
            continue
        tipe = (b.get('tipe') or 'BAHAN').upper()
        total += nilai
        akun_reclass = HPP_RECLASS_ACCOUNTS.get(tipe)
        if akun_reclass:
            _add_amount(reclass, akun_reclass, nilai)
        else:
            inventory += nilai
    return {
        'total': round(total, 2),
        'inventory': round(inventory, 2),
        'reclass': {k: round(v, 2) for k, v in reclass.items() if v > 0},
    }

def hpp_bahan_total(bahan):
    return hpp_bahan_breakdown(bahan)['total']

@app.route('/kalkulator-hpp')
@investor_required
@hpp_margin_required
def kalkulator_hpp():
    conn = db()
    produk_list = conn.execute(
        "SELECT id, nama, jenis, satuan, harga_jual, bahan FROM hpp_produk ORDER BY updated_at DESC"
    ).fetchall()
    conn.close()
    import json as _json
    items = []
    for p in produk_list:
        bahan = _json.loads(p['bahan'] or '[]')
        _bd = hpp_bahan_breakdown(bahan)
        items.append({'id': p['id'], 'nama': p['nama'],
                      'jenis': p['jenis'] or 'FNB_MENU', 'satuan': p['satuan'] or 'porsi',
                      'harga_jual': p['harga_jual'], 'total_hpp': _bd['total'],
                      'inventory_hpp': _bd['inventory'],
                      'reclass_hpp': round(sum(_bd['reclass'].values()), 2),
                      'reclass_accounts': _bd['reclass']})
    return render_template('hpp_kalkulator.html', produk_list=items)

@app.route('/api/hpp-produk', methods=['GET'])
@investor_required
@hpp_margin_required
def hpp_produk_list():
    import json as _json
    conn = db()
    rows = conn.execute(
        "SELECT id, nama, jenis, satuan, harga_jual, bahan FROM hpp_produk ORDER BY updated_at DESC"
    ).fetchall()
    conn.close()
    result = []
    for p in rows:
        bahan = _json.loads(p['bahan'] or '[]')
        _bd = hpp_bahan_breakdown(bahan)
        result.append({'id': p['id'], 'nama': p['nama'],
                       'jenis': p['jenis'] or 'FNB_MENU', 'satuan': p['satuan'] or 'porsi',
                       'harga_jual': p['harga_jual'], 'total_hpp': _bd['total'],
                       'inventory_hpp': _bd['inventory'],
                       'reclass_hpp': round(sum(_bd['reclass'].values()), 2),
                       'reclass_accounts': _bd['reclass']})
    return jsonify(result)

@app.route('/api/hpp-produk/<int:pid>', methods=['GET'])
@investor_required
@hpp_margin_required
def hpp_produk_get(pid):
    conn = db()
    p = conn.execute("SELECT * FROM hpp_produk WHERE id=?", (pid,)).fetchone()
    conn.close()
    if not p:
        return jsonify({'error': 'not found'}), 404
    import json as _json
    bahan = _json.loads(p['bahan'] or '[]')
    _bd = hpp_bahan_breakdown(bahan)
    return jsonify({'id': p['id'], 'nama': p['nama'],
                    'jenis': p['jenis'] or 'FNB_MENU',
                    'satuan': p['satuan'] or 'porsi',
                    'harga_jual': p['harga_jual'],
                    'total_hpp': _bd['total'],
                    'inventory_hpp': _bd['inventory'],
                    'reclass_hpp': round(sum(_bd['reclass'].values()), 2),
                    'reclass_accounts': _bd['reclass'],
                    'bahan': bahan})

@app.route('/api/hpp-produk', methods=['POST'])
@finance_required
@hpp_margin_required
def hpp_produk_save():
    import json as _json
    data = request.get_json()
    nama = (data.get('nama') or '').strip()
    if not nama:
        return jsonify({'error': 'nama wajib diisi'}), 400
    jenis = (data.get('jenis') or 'FNB_MENU').strip().upper()
    if jenis not in ('FNB_MENU', 'JASA', 'PAKET', 'NON_SKU'):
        jenis = 'FNB_MENU'
    satuan = (data.get('satuan') or 'porsi').strip() or 'porsi'
    harga_jual = float(data.get('harga_jual') or 0)
    bahan = _json.dumps(data.get('bahan') or [], ensure_ascii=False)
    conn = db()
    cur = conn.execute(
        "INSERT INTO hpp_produk(nama, jenis, satuan, harga_jual, bahan, updated_at) VALUES(?,?,?,?,?,date('now'))",
        (nama, jenis, satuan, harga_jual, bahan)
    )
    pid = cur.lastrowid
    conn.commit(); conn.close()
    return jsonify({'id': pid, 'nama': nama})

@app.route('/api/hpp-produk/<int:pid>', methods=['PUT'])
@finance_required
@hpp_margin_required
def hpp_produk_update(pid):
    import json as _json
    data = request.get_json()
    nama = (data.get('nama') or '').strip()
    if not nama:
        return jsonify({'error': 'nama wajib diisi'}), 400
    jenis = (data.get('jenis') or 'FNB_MENU').strip().upper()
    if jenis not in ('FNB_MENU', 'JASA', 'PAKET', 'NON_SKU'):
        jenis = 'FNB_MENU'
    satuan = (data.get('satuan') or 'porsi').strip() or 'porsi'
    harga_jual = float(data.get('harga_jual') or 0)
    bahan = _json.dumps(data.get('bahan') or [], ensure_ascii=False)
    conn = db()
    conn.execute(
        "UPDATE hpp_produk SET nama=?, jenis=?, satuan=?, harga_jual=?, bahan=?, updated_at=date('now') WHERE id=?",
        (nama, jenis, satuan, harga_jual, bahan, pid)
    )
    conn.commit(); conn.close()
    return jsonify({'id': pid, 'nama': nama})

@app.route('/api/hpp-produk/<int:pid>', methods=['DELETE'])
@finance_required
@hpp_margin_required
def hpp_produk_delete(pid):
    conn = db()
    conn.execute("DELETE FROM hpp_produk WHERE id=?", (pid,))
    conn.commit(); conn.close()
    return jsonify({'ok': True})


# ---------- BACKUP & RESTORE ----------
def _table_columns(conn, table):
    return {r['name'] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()}


def _row_value(row, key, default=None):
    return row[key] if key in row.keys() else default


def _insert_common_row(conn, table, row, exclude=('id',), overrides=None):
    overrides = overrides or {}
    columns = [
        col for col in _table_columns(conn, table)
        if col not in exclude and (col in row.keys() or col in overrides)
    ]
    values = [overrides[col] if col in overrides else row[col] for col in columns]
    placeholders = ','.join('?' * len(columns))
    conn.execute(
        f"INSERT INTO {table}({','.join(columns)}) VALUES({placeholders})",
        values
    )
    return conn.execute("SELECT last_insert_rowid()").fetchone()[0]


def _journal_signature(conn, row, jurnal_id):
    details = conn.execute("""
        SELECT a.kode,d.debit,d.kredit
        FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id
        WHERE d.jurnal_id=? ORDER BY a.kode,d.id
    """, (jurnal_id,)).fetchall()
    lines = tuple(sorted(
        (d['kode'], round(float(d['debit'] or 0), 2), round(float(d['kredit'] or 0), 2))
        for d in details
    ))
    return (
        str(_row_value(row, 'tanggal', '') or ''),
        str(_row_value(row, 'keterangan', '') or ''),
        str(_row_value(row, 'referensi', '') or ''),
        str(_row_value(row, 'kategori', 'OPERASIONAL') or 'OPERASIONAL'),
        str(_row_value(row, 'tipe_tx', 'JURNAL') or 'JURNAL'),
        str(_row_value(row, 'pihak', '') or ''),
        lines,
    )


def _merge_restore_database(source_path):
    """Tambahkan data transaksi baru tanpa menimpa settings dan user database aktif."""
    from collections import Counter

    target = db()
    source = sqlite3.connect(source_path)
    source.row_factory = sqlite3.Row
    source_tables = {
        r['name'] for r in source.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()
    }
    summary = {
        'jurnal': 0, 'duplikat': 0, 'konflik_id': 0, 'produk': 0,
        'customer': 0, 'vendor': 0, 'invoice': 0,
        'payroll': 0, 'po': 0, 'produksi': 0, 'pos': 0,
    }
    try:
        target.execute('BEGIN')

        # Akun diperlukan agar setiap baris jurnal sumber tetap dapat dipetakan.
        for row in source.execute("SELECT * FROM akun ORDER BY id").fetchall():
            existing = target.execute("SELECT id FROM akun WHERE kode=?", (row['kode'],)).fetchone()
            if not existing:
                _insert_common_row(target, 'akun', row)

        product_map = {}
        new_products = set()
        if 'produk' in source_tables:
            for row in source.execute("SELECT * FROM produk ORDER BY id").fetchall():
                existing = target.execute("SELECT id FROM produk WHERE kode=?", (row['kode'],)).fetchone()
                if existing:
                    product_map[row['id']] = existing['id']
                else:
                    pid = _insert_common_row(target, 'produk', row, overrides={'stok': 0})
                    product_map[row['id']] = pid
                    new_products.add(pid)
                    summary['produk'] += 1

        for table in ('customer', 'vendor'):
            if table not in source_tables:
                continue
            for row in source.execute(f"SELECT * FROM {table} ORDER BY id").fetchall():
                existing = target.execute(
                    f"SELECT id FROM {table} WHERE nama=?", (row['nama'],)
                ).fetchone()
                if not existing:
                    _insert_common_row(target, table, row)
                    summary[table] += 1

        existing_signatures = Counter()
        for row in target.execute("SELECT * FROM jurnal ORDER BY id").fetchall():
            existing_signatures[_journal_signature(target, row, row['id'])] += 1

        imported_journals = {}
        for row in source.execute("SELECT * FROM jurnal ORDER BY tanggal,id").fetchall():
            signature = _journal_signature(source, row, row['id'])
            nomor_tx = (_row_value(row, 'nomor_tx', '') or '').strip()
            existing_by_number = None
            if nomor_tx:
                existing_by_number = target.execute(
                    "SELECT * FROM jurnal WHERE nomor_tx=?", (nomor_tx,)
                ).fetchone()
            if existing_by_number:
                if _journal_signature(target, existing_by_number, existing_by_number['id']) != signature:
                    summary['konflik_id'] += 1
                else:
                    summary['duplikat'] += 1
                continue
            if existing_signatures[signature] > 0:
                existing_signatures[signature] -= 1
                summary['duplikat'] += 1
                continue

            entries = []
            for detail in source.execute("""
                SELECT a.kode,d.debit,d.kredit
                FROM detail_jurnal d JOIN akun a ON a.id=d.akun_id
                WHERE d.jurnal_id=? ORDER BY d.id
            """, (row['id'],)).fetchall():
                entries.append((detail['kode'], detail['debit'], detail['kredit']))
            jid = insert_jurnal(
                target, row['tanggal'], row['keterangan'],
                _row_value(row, 'kategori', 'OPERASIONAL') or 'OPERASIONAL',
                _row_value(row, 'tipe_tx', 'JURNAL') or 'JURNAL',
                entries, pihak=_row_value(row, 'pihak', '') or ''
            )
            generated_number = target.execute(
                "SELECT nomor_tx FROM jurnal WHERE id=?", (jid,)
            ).fetchone()['nomor_tx']
            preserved_number = nomor_tx or generated_number
            target.execute("""
                UPDATE jurnal SET referensi=?, dibuat=?, pihak=?, tx_meta=?, nomor_tx=?
                WHERE id=?
            """, (
                _row_value(row, 'referensi', ''),
                _row_value(row, 'dibuat', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
                _row_value(row, 'pihak', ''),
                _row_value(row, 'tx_meta', ''),
                preserved_number, jid
            ))
            imported_journals[row['id']] = jid
            existing_signatures[signature] += 1
            summary['jurnal'] += 1

        if 'transaksi_item' in source_tables:
            for row in source.execute("SELECT * FROM transaksi_item ORDER BY id").fetchall():
                jid = imported_journals.get(row['jurnal_id'])
                if not jid:
                    continue
                _insert_common_row(target, 'transaksi_item', row, overrides={
                    'jurnal_id': jid,
                    'produk_id': product_map.get(row['produk_id']) if row['produk_id'] else None,
                })

        moved_products = set()
        existing_unlinked_movements = Counter()
        for row in target.execute("""
            SELECT ps.*,p.kode FROM pergerakan_stok ps
            JOIN produk p ON p.id=ps.produk_id WHERE ps.jurnal_id IS NULL
        """).fetchall():
            key = (row['kode'], row['tanggal'], row['jenis'], round(row['qty'], 6),
                   round(row['harga'] or 0, 2), row['keterangan'] or '')
            existing_unlinked_movements[key] += 1

        if 'pergerakan_stok' in source_tables and 'produk' in source_tables:
            for row in source.execute("""
                SELECT ps.*,p.kode FROM pergerakan_stok ps
                JOIN produk p ON p.id=ps.produk_id ORDER BY ps.id
            """).fetchall():
                source_jid = _row_value(row, 'jurnal_id')
                jid = imported_journals.get(source_jid) if source_jid else None
                if source_jid and not jid:
                    continue
                if not source_jid:
                    key = (row['kode'], row['tanggal'], row['jenis'], round(row['qty'], 6),
                           round(row['harga'] or 0, 2), row['keterangan'] or '')
                    if existing_unlinked_movements[key] > 0:
                        existing_unlinked_movements[key] -= 1
                        continue
                pid = product_map[row['produk_id']]
                _insert_common_row(target, 'pergerakan_stok', row, overrides={
                    'produk_id': pid, 'jurnal_id': jid,
                })
                qty = float(row['qty'] or 0)
                delta = qty if row['jenis'] == 'MASUK' else (-qty if row['jenis'] == 'KELUAR' else 0)
                if delta:
                    target.execute("UPDATE produk SET stok=stok+? WHERE id=?", (delta, pid))
                moved_products.add(pid)

        # Harga pokok akhir sumber dipakai setelah seluruh mutasi baru diterapkan.
        if 'produk' in source_tables:
            for row in source.execute("SELECT id,harga_beli,stok FROM produk ORDER BY id").fetchall():
                pid = product_map[row['id']]
                if pid in moved_products:
                    target.execute("UPDATE produk SET harga_beli=? WHERE id=?", (row['harga_beli'], pid))
                elif pid in new_products:
                    target.execute("UPDATE produk SET stok=?,harga_beli=? WHERE id=?",
                                   (row['stok'], row['harga_beli'], pid))

        tracker_maps = {'piutang': {}, 'hutang': {}}
        for table, child, fk in (
            ('piutang', 'bayar_piutang', 'piutang_id'),
            ('hutang', 'bayar_hutang', 'hutang_id'),
        ):
            if table not in source_tables:
                continue
            for row in source.execute(f"SELECT * FROM {table} ORDER BY id").fetchall():
                source_jid = _row_value(row, 'jurnal_id')
                jid = imported_journals.get(source_jid) if source_jid else None
                if source_jid and not jid:
                    continue
                values = (
                    row['tanggal'], _row_value(row, 'jatuh_tempo'), row['jumlah'],
                    _row_value(row, 'terbayar', 0), _row_value(row, 'status', ''),
                    _row_value(row, 'keterangan', '')
                )
                party_col = 'pelanggan' if table == 'piutang' else 'pemasok'
                existing = target.execute(
                    f"""SELECT id FROM {table}
                        WHERE tanggal=? AND {party_col}=? AND jumlah=? AND terbayar=?
                          AND status=? AND COALESCE(keterangan,'')=COALESCE(?,'')""",
                    (values[0], row[party_col], values[2], values[3], values[4], values[5])
                ).fetchone()
                if existing:
                    tracker_maps[table][row['id']] = existing['id']
                    continue
                _ov = {'jurnal_id': jid}
                if table == 'piutang':
                    # invoice_id sumber adalah ID invoice di DB SUMBER; di DB ini ID yang sama bisa
                    # milik invoice lain (mis. invoice tunai pelanggan lain), yang lalu terkunci
                    # Lunas. Dikosongkan, lalu ditautkan ulang lewat jurnal_id sesudah invoice masuk.
                    _ov['invoice_id'] = None
                rid = _insert_common_row(target, table, row, overrides=_ov)
                tracker_maps[table][row['id']] = rid
                if child in source_tables:
                    for payment in source.execute(
                        f"SELECT * FROM {child} WHERE {fk}=? ORDER BY id", (row['id'],)
                    ).fetchall():
                        payment_jid = _row_value(payment, 'jurnal_id')
                        mapped_payment_jid = imported_journals.get(payment_jid) if payment_jid else None
                        if payment_jid and not mapped_payment_jid:
                            continue
                        overrides = {fk: rid}
                        if 'jurnal_id' in _table_columns(target, child):
                            overrides['jurnal_id'] = mapped_payment_jid
                        _insert_common_row(target, child, payment, overrides=overrides)

        if 'penyesuaian_tagihan' in source_tables:
            for row in source.execute("SELECT * FROM penyesuaian_tagihan ORDER BY id").fetchall():
                jid = imported_journals.get(row['jurnal_id'])
                if not jid:
                    continue
                jenis = row['jenis']
                table = 'piutang' if jenis == 'PIUTANG' else 'hutang'
                rid = tracker_maps.get(table, {}).get(row['record_id'])
                if not rid:
                    continue
                existing = target.execute(
                    """SELECT id FROM penyesuaian_tagihan
                       WHERE jurnal_id=? AND jenis=? AND record_id=? AND jumlah_delta=?""",
                    (jid, jenis, rid, row['jumlah_delta'])
                ).fetchone()
                if not existing:
                    _insert_common_row(target, 'penyesuaian_tagihan', row, overrides={
                        'jurnal_id': jid, 'record_id': rid,
                    })

        if 'pergerakan_persediaan_non_sku' in source_tables:
            for row in source.execute("SELECT * FROM pergerakan_persediaan_non_sku ORDER BY id").fetchall():
                jid = imported_journals.get(row['jurnal_id'])
                if not jid:
                    continue
                existing = target.execute(
                    """SELECT id FROM pergerakan_persediaan_non_sku
                       WHERE jurnal_id=? AND tanggal=? AND deskripsi=? AND nilai_delta=?""",
                    (jid, row['tanggal'], row['deskripsi'], row['nilai_delta'])
                ).fetchone()
                if not existing:
                    _insert_common_row(target, 'pergerakan_persediaan_non_sku', row, overrides={'jurnal_id': jid})

        karyawan_map = {}
        if 'karyawan' in source_tables:
            for row in source.execute("SELECT * FROM karyawan ORDER BY id").fetchall():
                no_karyawan = (_row_value(row, 'no_karyawan', '') or '').strip()
                if no_karyawan:
                    existing = target.execute(
                        "SELECT id FROM karyawan WHERE no_karyawan=?", (no_karyawan,)
                    ).fetchone()
                else:
                    existing = target.execute(
                        "SELECT id FROM karyawan WHERE nama=? AND jabatan=?",
                        (row['nama'], _row_value(row, 'jabatan', ''))
                    ).fetchone()
                if existing:
                    karyawan_map[row['id']] = existing['id']
                else:
                    kid = _insert_common_row(target, 'karyawan', row)
                    karyawan_map[row['id']] = kid

        if 'payroll' in source_tables:
            for row in source.execute("SELECT * FROM payroll ORDER BY id").fetchall():
                kid = karyawan_map.get(row['karyawan_id'])
                if not kid:
                    continue
                source_jid = _row_value(row, 'jurnal_id')
                jid = imported_journals.get(source_jid) if source_jid else None
                if source_jid and not jid:
                    continue
                existing = target.execute(
                    """SELECT id FROM payroll
                       WHERE karyawan_id=? AND periode_bulan=? AND periode_tahun=?""",
                    (kid, row['periode_bulan'], row['periode_tahun'])
                ).fetchone()
                if existing:
                    continue
                _insert_common_row(target, 'payroll', row, overrides={'karyawan_id': kid, 'jurnal_id': jid})
                summary['payroll'] += 1

        po_map = {}
        if 'purchase_order' in source_tables:
            for row in source.execute("SELECT * FROM purchase_order ORDER BY id").fetchall():
                existing = target.execute("SELECT id FROM purchase_order WHERE nomor=?", (row['nomor'],)).fetchone()
                if existing:
                    po_map[row['id']] = existing['id']
                    continue
                source_jid = _row_value(row, 'jurnal_id')
                source_hutang_id = _row_value(row, 'hutang_id')
                jid = imported_journals.get(source_jid) if source_jid else None
                hid = tracker_maps['hutang'].get(source_hutang_id) if source_hutang_id else None
                if source_jid and not jid:
                    continue
                if source_hutang_id and not hid:
                    continue
                po_id = _insert_common_row(target, 'purchase_order', row, overrides={
                    'jurnal_id': jid, 'hutang_id': hid,
                })
                po_map[row['id']] = po_id
                summary['po'] += 1
            if 'po_item' in source_tables:
                for item in source.execute("SELECT * FROM po_item ORDER BY id").fetchall():
                    po_id = po_map.get(item['po_id'])
                    if po_id:
                        _insert_common_row(target, 'po_item', item, overrides={'po_id': po_id})

        produksi_map = {}
        if 'produksi' in source_tables:
            for row in source.execute("SELECT * FROM produksi ORDER BY id").fetchall():
                pid = product_map.get(row['produk_jadi_id'])
                if not pid:
                    continue
                source_jid = _row_value(row, 'jurnal_id')
                jid = imported_journals.get(source_jid) if source_jid else None
                if source_jid and not jid:
                    continue
                existing = target.execute("SELECT id FROM produksi WHERE jurnal_id=?", (jid,)).fetchone() if jid else None
                if not existing:
                    existing = target.execute(
                        """SELECT id FROM produksi
                           WHERE tanggal=? AND produk_jadi_id=? AND qty_hasil=? AND total_hpp=?
                             AND COALESCE(keterangan,'')=COALESCE(?,'')""",
                        (row['tanggal'], pid, row['qty_hasil'], row['total_hpp'], _row_value(row, 'keterangan', ''))
                    ).fetchone()
                if existing:
                    produksi_map[row['id']] = existing['id']
                    continue
                produksi_id = _insert_common_row(target, 'produksi', row, overrides={
                    'produk_jadi_id': pid, 'jurnal_id': jid,
                })
                produksi_map[row['id']] = produksi_id
                summary['produksi'] += 1
            if 'produksi_bahan' in source_tables:
                for item in source.execute("SELECT * FROM produksi_bahan ORDER BY id").fetchall():
                    produksi_id = produksi_map.get(item['produksi_id'])
                    pid = product_map.get(item['produk_id'])
                    if produksi_id and pid:
                        _insert_common_row(target, 'produksi_bahan', item, overrides={
                            'produksi_id': produksi_id, 'produk_id': pid,
                        })
            if 'produksi_non_sku' in source_tables:
                for item in source.execute("SELECT * FROM produksi_non_sku ORDER BY id").fetchall():
                    produksi_id = produksi_map.get(item['produksi_id'])
                    if produksi_id:
                        _insert_common_row(target, 'produksi_non_sku', item, overrides={'produksi_id': produksi_id})

        pos_shift_map = {}
        if 'pos_shift' in source_tables:
            for row in source.execute("SELECT * FROM pos_shift ORDER BY id").fetchall():
                username = _row_value(row, 'username', '') or ''
                user = target.execute("SELECT id FROM users WHERE username=?", (username,)).fetchone()
                if not user:
                    continue
                existing = target.execute(
                    "SELECT id FROM pos_shift WHERE username=? AND dibuka=?",
                    (username, _row_value(row, 'dibuka'))
                ).fetchone()
                if existing:
                    pos_shift_map[row['id']] = existing['id']
                else:
                    shift_id = _insert_common_row(target, 'pos_shift', row, overrides={'user_id': user['id']})
                    pos_shift_map[row['id']] = shift_id

        if 'pos_sale' in source_tables:
            for row in source.execute("SELECT * FROM pos_sale ORDER BY id").fetchall():
                if target.execute("SELECT id FROM pos_sale WHERE nomor_struk=?", (row['nomor_struk'],)).fetchone():
                    continue
                source_jid = _row_value(row, 'jurnal_id')
                jid = imported_journals.get(source_jid) if source_jid else None
                if not jid:
                    continue
                shift_id = pos_shift_map.get(_row_value(row, 'shift_id')) if _row_value(row, 'shift_id') else None
                _insert_common_row(target, 'pos_sale', row, overrides={
                    'jurnal_id': jid, 'shift_id': shift_id,
                })
                summary['pos'] += 1

        if 'invoice' in source_tables:
            for row in source.execute("SELECT * FROM invoice ORDER BY id").fetchall():
                if target.execute("SELECT id FROM invoice WHERE nomor=?", (row['nomor'],)).fetchone():
                    continue
                source_jid = _row_value(row, 'jurnal_id')
                invoice_id = _insert_common_row(target, 'invoice', row, overrides={
                    'jurnal_id': imported_journals.get(source_jid) if source_jid else None,
                })
                if 'invoice_item' in source_tables:
                    for item in source.execute(
                        "SELECT * FROM invoice_item WHERE invoice_id=? ORDER BY id", (row['id'],)
                    ).fetchall():
                        _insert_common_row(target, 'invoice_item', item, overrides={'invoice_id': invoice_id})
                summary['invoice'] += 1
        # Tautkan ulang piutang hasil impor ke invoice-nya lewat jurnal_id (lihat impor piutang).
        target.execute("""
            UPDATE piutang SET invoice_id = (
                SELECT id FROM invoice WHERE invoice.jurnal_id = piutang.jurnal_id LIMIT 1
            )
            WHERE invoice_id IS NULL AND jurnal_id IS NOT NULL
        """)

        if 'aset_tetap' in source_tables:
            for row in source.execute("SELECT * FROM aset_tetap ORDER BY id").fetchall():
                existing = target.execute("""
                    SELECT id FROM aset_tetap WHERE nama=? AND tanggal_beli=? AND harga_beli=?
                """, (row['nama'], row['tanggal_beli'], row['harga_beli'])).fetchone()
                if not existing:
                    _insert_common_row(target, 'aset_tetap', row)

        if 'hpp_produk' in source_tables:
            for row in source.execute("SELECT * FROM hpp_produk ORDER BY id").fetchall():
                existing = target.execute("""
                    SELECT id FROM hpp_produk WHERE nama=? AND harga_jual=? AND bahan=?
                """, (row['nama'], row['harga_jual'], row['bahan'])).fetchone()
                if not existing:
                    _insert_common_row(target, 'hpp_produk', row)

        backfill_nomor_tx(target)
        target.commit()
        return summary
    except Exception:
        target.rollback()
        raise
    finally:
        source.close()
        target.close()


@app.route('/backup')
@finance_required
def backup():
    if session.get('role') == 'DEMO':
        flash('User Demo tidak boleh mengunduh backup database.', 'danger')
        return redirect(url_for('dashboard'))
    import shutil, tempfile
    today = date.today().strftime('%Y%m%d')
    tmp = tempfile.NamedTemporaryFile(suffix='.db', delete=False)
    tmp.close()
    src_conn = sqlite3.connect(DB)
    dst_conn = sqlite3.connect(tmp.name)
    src_conn.backup(dst_conn)
    src_conn.close()
    dst_conn.close()
    conn2 = db()
    add_log(conn2, 'Backup database', f"finansial_backup_{today}.db", 'MODAL')
    conn2.commit(); conn2.close()
    return send_file(tmp.name, as_attachment=True,
                     download_name=f'finansial_backup_{today}.db',
                     mimetype='application/octet-stream')

@app.route('/restore', methods=['POST'])
@finance_required
def restore():
    f = request.files.get('backup_file')
    restore_mode = request.form.get('restore_mode', 'replace')
    if restore_mode not in ('replace', 'merge'):
        flash('Mode restore tidak valid.', 'danger')
        return redirect('/settings')
    if not f or not f.filename.endswith('.db'):
        flash('File tidak valid. Gunakan file .db hasil backup.', 'danger')
        return redirect('/settings')
    import tempfile, shutil
    tmp = tempfile.NamedTemporaryFile(suffix='.db', delete=False)
    f.save(tmp.name)
    tmp.close()
    # Validate: try opening and checking tables exist
    try:
        test = sqlite3.connect(tmp.name)
        tables = {r[0] for r in test.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
        test.close()
        required = {'jurnal', 'detail_jurnal', 'akun', 'users'}
        if not required.issubset(tables):
            raise ValueError('Bukan file database FinansialApp yang valid.')
    except Exception as e:
        os.unlink(tmp.name)
        flash(f'Restore gagal: {e}', 'danger')
        return redirect('/settings')
    if restore_mode == 'replace':
        # Replace database
        shutil.copy2(tmp.name, DB)
        os.unlink(tmp.name)
        init_db()  # apply any missing migrations to the restored database
        conn2 = db()
        add_log(conn2, 'Restore database', f"Mode ganti penuh | File: {f.filename}", 'MODAL')
        conn2.commit(); conn2.close()
        flash('Database berhasil diganti penuh. Silakan login ulang jika diperlukan.', 'success')
    else:
        try:
            summary = _merge_restore_database(tmp.name)
        except Exception as e:
            os.unlink(tmp.name)
            flash(f'Restore tambah data gagal: {e}', 'danger')
            return redirect('/settings')
        os.unlink(tmp.name)
        conn2 = db()
        add_log(conn2, 'Restore database', (
            f"Mode tambah | File: {f.filename} | "
            f"Jurnal baru: {summary['jurnal']} | Duplikat dilewati: {summary['duplikat']} | "
            f"Konflik ID: {summary['konflik_id']}"
        ), 'MODAL')
        conn2.commit(); conn2.close()
        flash(
            f"Restore tambah data selesai. {summary['jurnal']} transaksi baru ditambahkan; "
            f"{summary['duplikat']} transaksi duplikat dilewati; "
            f"{summary['konflik_id']} konflik ID dilewati.",
            'success'
        )
    return redirect('/settings')


# ---------- CATATAN RILIS ----------
@app.route('/changelog')
@login_required
def changelog():
    """Halaman 'Apa yang Baru'. Isinya statis di template, jadi tak perlu query.

    Dulu ini pop-up kecil di base.html yang harus dibuka satu versi per satu lewat
    tombol maju-mundur. Setelah daftar versinya panjang, pop-up jadi tidak nyaman:
    tidak bisa di-scroll bebas, tidak bisa dicari dengan Ctrl+F, dan alamatnya tidak
    bisa dibagikan. Sengaja @login_required saja (bukan per-peran) supaya semua
    pengguna, termasuk kasir, bisa membaca apa yang berubah di aplikasi yang dipakainya.
    """
    return render_template('changelog.html')


# ---------- LOG AKTIVITAS ----------
@app.route('/log')
@finance_required
def log_aktivitas():
    import math
    conn = db()
    page     = int(request.args.get('page', 1))
    per      = 50
    q        = request.args.get('q', '').strip()
    kategori = request.args.get('kategori', '').strip()
    where, params = [], []
    if q:
        where.append("(detail LIKE ? OR username LIKE ? OR aksi LIKE ? OR role LIKE ?)")
        params += [f'%{q}%', f'%{q}%', f'%{q}%', f'%{q}%']
    if kategori:
        where.append("kategori=?")
        params.append(kategori)
    where_sql = ('WHERE ' + ' AND '.join(where)) if where else ''
    total = conn.execute(f"SELECT COUNT(*) FROM log_aktivitas {where_sql}", params).fetchone()[0]
    logs  = conn.execute(
        f"SELECT * FROM log_aktivitas {where_sql} ORDER BY id DESC LIMIT ? OFFSET ?",
        params + [per, (page-1)*per]
    ).fetchall()
    # counts per category for stats row
    counts = {}
    for row in conn.execute("SELECT kategori, COUNT(*) as c FROM log_aktivitas GROUP BY kategori"):
        counts[row['kategori']] = row['c']
    conn.close()
    total_pages = max(1, math.ceil(total / per))
    return render_template('log.html',
        logs=logs, page=page, total_pages=total_pages,
        total=total, q=q, kategori=kategori, counts=counts)


# ---------- SETTINGS ----------
@app.route('/settings', methods=['GET','POST'])
@admin_required
def settings():
    conn = db()
    if request.method == 'POST':
        action = request.form.get('action','')
        if action in ('save_settings', 'save_prefs', 'save_all'):
            # ── Unified save: handle Nama Usaha + semua preferensi sekaligus ──
            # Dukung 3 action name untuk backward compat. Section disimpan kalau
            # field-nya ada di form (idempotent, section yg tidak dikirim tidak
            # tersentuh).

            # 1. Pengaturan Umum: Nama Usaha + Zona Waktu
            if 'nama_usaha' in request.form:
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('nama_usaha', ?)",
                             (request.form.get('nama_usaha',''),))
            if 'tz_app' in request.form:
                # Whitelist murni dari <select> -- jangan simpan/terapkan nilai
                # sembarangan krn ini langsung jadi os.environ['TZ'] proses.
                _tz_pick = request.form.get('tz_app', '')
                if _tz_pick in TZ_APP_PILIHAN:
                    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('tz_app', ?)", (_tz_pick,))
                    apply_tz_app(_tz_pick)  # langsung berlaku, tanpa restart server

            # 2. Preferensi Tampilan: Custom COA, Desimal, Widget Rasio.
            # Marker '_save_prefs_marker' menandakan section prefs ikut disubmit
            # (krn checkbox unchecked tidak muncul di form). save_prefs/save_all
            # juga selalu trigger section ini.
            if '_save_prefs_marker' in request.form or action in ('save_prefs', 'save_all'):
                v_coa = '1' if request.form.get('pref_custom_coa') == '1' else '0'
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pref_custom_coa', ?)", (v_coa,))
                v_dec = '1' if request.form.get('pref_decimal') == '1' else '0'
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pref_decimal', ?)", (v_dec,))
                v_hint = '1' if request.form.get('pref_hide_login_hint') == '1' else '0'
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pref_hide_login_hint', ?)", (v_hint,))
                picks = request.form.getlist('pref_ratio_widgets')
                valid_keys = {k for (k,*_rest) in RATIO_WIDGETS_AVAIL}
                picks = [w for w in picks if w in valid_keys]
                csv_picks = ','.join(picks) if picks else 'current_ratio,der,roe,npm'
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pref_ratio_widgets', ?)", (csv_picks,))
                simple_picks = request.form.getlist('pref_simple_features')
                simple_picks = [w for w in simple_picks if w in SIMPLE_FEATURE_KEYS]
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pref_simple_features', ?)",
                             (','.join(simple_picks),))
                reload_app_prefs()

            if '_save_pos_receipt_marker' in request.form or action == 'save_all':
                for k in ['pos_struk_nama', 'pos_struk_subtitle', 'pos_struk_alamat',
                          'pos_struk_telepon', 'pos_struk_header', 'pos_struk_footer',
                          'pos_struk_label_kasir']:
                    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                                 (k, request.form.get(k, '').strip()))
                menu_size = request.form.get('pos_menu_size', 'medium')
                if menu_size not in ('small', 'medium', 'large'):
                    menu_size = 'medium'
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_menu_size', ?)",
                             (menu_size,))
                icon_mode = request.form.get('pos_icon_mode', 'initial')
                if icon_mode not in ('custom', 'initial', 'code'):
                    icon_mode = 'initial'
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_icon_mode', ?)",
                             (icon_mode,))
                if 'pos_default_payment' in request.form:
                    default_payment = request.form.get('pos_default_payment', 'TUNAI')
                    if default_payment not in ('TUNAI', 'REKENING', 'PIUTANG'):
                        default_payment = 'TUNAI'
                    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_default_payment', ?)",
                                 (default_payment,))
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_struk_show_customer', ?)",
                             ('1' if request.form.get('pos_struk_show_customer') == '1' else '0',))
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_struk_show_payment', ?)",
                             ('1' if request.form.get('pos_struk_show_payment') == '1' else '0',))
                # Ukuran kertas printer thermal. Divalidasi ke daftar yang didukung supaya
                # nilai ngawur dari form tidak membuat struk tercetak dengan lebar aneh.
                kertas = (request.form.get('pos_struk_kertas', '') or '').strip()
                if kertas not in POS_STRUK_AREA_MM:
                    kertas = POS_STRUK_KERTAS_DEFAULT
                conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_struk_kertas', ?)",
                             (kertas,))

            conn.commit()
            flash('Pengaturan tersimpan!', 'success')
        elif action == 'add_user':
            username = request.form.get('username','').strip()
            password = request.form.get('password','')
            nama     = request.form.get('nama','')
            role     = request.form.get('role','KASIR')
            if role not in ('ADMIN', 'FINANCE', 'MANAJER', 'INVESTOR', 'OPERATOR', 'KASIR'):
                role = 'KASIR'
            permissions = json.dumps(_permissions_from_form())
            ph = sha256(password.encode()).hexdigest()
            try:
                conn.execute("INSERT INTO users(username,password_hash,nama,role,permissions) VALUES(?,?,?,?,?)",
                             (username, ph, nama, role, permissions))
                add_log(conn, 'Tambah user', f"Username: {username} | Role: {role}", 'ADMIN')
                conn.commit()
                flash(f'User {username} berhasil ditambahkan!','success')
            except sqlite3.IntegrityError:
                flash('Username sudah digunakan!','danger')
        elif action == 'toggle_user':
            uid = int(request.form.get('user_id',0))
            u = conn.execute("SELECT username, aktif FROM users WHERE id=?", (uid,)).fetchone()
            conn.execute("UPDATE users SET aktif=CASE WHEN aktif=1 THEN 0 ELSE 1 END WHERE id=? AND username!='admin'", (uid,))
            if u:
                status_baru = 'dinonaktifkan' if u['aktif'] else 'diaktifkan'
                add_log(conn, 'Toggle status user', f"Username: {u['username']} → {status_baru}", 'ADMIN')
            conn.commit()
            flash('Status user diperbarui.','info')
        elif action == 'reset_password':
            if DEMO_LOCK:
                flash('Mode demo aktif: password tidak dapat diubah.', 'warning')
            else:
                uid = int(request.form.get('user_id',0))
                new_pw = request.form.get('new_password','')
                if new_pw:
                    ph = sha256(new_pw.encode()).hexdigest()
                    conn.execute("UPDATE users SET password_hash=? WHERE id=?", (ph, uid))
                    u = conn.execute("SELECT username FROM users WHERE id=?", (uid,)).fetchone()
                    add_log(conn, 'Reset password user', f"Username: {u['username'] if u else uid}", 'PASSWORD')
                    conn.commit()
                    flash('Password berhasil direset.','success')
        elif action == 'delete_user':
            uid = int(request.form.get('user_id',0))
            u = conn.execute("SELECT username FROM users WHERE id=?", (uid,)).fetchone()
            conn.execute("DELETE FROM users WHERE id=? AND username!='admin'", (uid,))
            add_log(conn, 'Hapus user', f"Username: {u['username'] if u else uid}", 'ADMIN')
            conn.commit()
            flash('User dihapus.','warning')
        elif action == 'change_role':
            uid      = int(request.form.get('user_id', 0))
            new_role = request.form.get('new_role', '').strip()
            if new_role in ('ADMIN', 'FINANCE', 'MANAJER', 'INVESTOR', 'OPERATOR', 'KASIR') and uid:
                u = conn.execute("SELECT username, role FROM users WHERE id=?", (uid,)).fetchone()
                default_permissions = sorted(ROLE_DEFAULT_PERMISSIONS.get(new_role, set()))
                conn.execute(
                    "UPDATE users SET role=?, permissions=? WHERE id=? AND username!='admin'",
                    (new_role, json.dumps(default_permissions), uid)
                )
                if u:
                    add_log(conn, 'Ubah role user', f"Username: {u['username']} | {u['role']} → {new_role}", 'ADMIN')
                conn.commit()
                flash(f'Role berhasil diubah ke {new_role} dan hak akses diset ke default role.', 'success')
        elif action == 'update_permissions':
            uid = int(request.form.get('user_id', 0))
            u = conn.execute("SELECT username, role FROM users WHERE id=?", (uid,)).fetchone()
            if u and u['username'] != 'admin':
                permissions = _permissions_from_form()
                conn.execute("UPDATE users SET permissions=? WHERE id=?", (json.dumps(permissions), uid))
                add_log(conn, 'Ubah hak akses user', f"Username: {u['username']} | {', '.join(permissions) or 'tanpa hak custom'}", 'ADMIN')
                conn.commit()
                flash(f'Hak akses {u["username"]} diperbarui.', 'success')
        elif action == 'add_rekening':
            rek_nama = request.form.get('rek_nama','').strip()
            rek_no   = request.form.get('rek_no','').strip()
            rek_jenis= request.form.get('rek_jenis','Bank')
            if rek_nama:
                # Cari kode kosong pertama di blok rekening kas/bank (1101-1119).
                # Dulu polanya '11[0-9][0-9]' + ambil yang TERBESAR, jadi ikut menghitung
                # akun pajak 1180-1183 dan rekening baru dapat kode 1184 (nempel di blok
                # pajak), sekaligus menyisakan kuota belasan saja. Sekarang cari slot kosong
                # supaya kode juga tidak bentrok dgn akun yang dibuat manual lewat Bagan Akun.
                terpakai = {r['kode'] for r in conn.execute(
                    "SELECT kode FROM akun WHERE kode GLOB '11[0-9][0-9]'").fetchall()}
                next_kode = next((str(n) for n in range(1101, 1120) if str(n) not in terpakai), None)
                if next_kode is None:
                    flash('Blok kode rekening kas/bank (1101-1119) sudah penuh. '
                          'Hapus rekening yang tidak terpakai, atau buat akun baru lewat '
                          'Bagan Akun lalu centang "Jadikan rekening kas/bank".', 'danger')
                else:
                    nama_akun = f"{rek_jenis} {rek_nama}"
                    conn.execute(
                        "INSERT INTO akun(kode,nama,tipe,subtipe,saldo_normal,is_rekening,no_rekening) VALUES(?,?,?,?,?,?,?)",
                        (next_kode, nama_akun, 'ASET', 'Aset Lancar', 'DEBIT', 1, rek_no or None)
                    )
                    conn.commit()
                    flash(f'Rekening {nama_akun} berhasil ditambahkan!','success')
        elif action == 'hapus_rekening':
            aid = int(request.form.get('akun_id',0))
            a = conn.execute("SELECT * FROM akun WHERE id=?", (aid,)).fetchone()
            if a and a['kode'] in ('1100','1110'):
                flash('Rekening Kas dan Bank utama tidak bisa dihapus.','danger')
            elif a:
                used = conn.execute("SELECT COUNT(*) FROM detail_jurnal WHERE akun_id=?", (aid,)).fetchone()[0]
                if used > 0:
                    flash('Rekening sudah digunakan dalam transaksi, tidak bisa dihapus.','danger')
                else:
                    conn.execute("DELETE FROM akun WHERE id=?", (aid,))
                    conn.commit()
                    flash('Rekening dihapus.','warning')
        elif action == 'edit_rekening':
            aid      = int(request.form.get('akun_id', 0))
            rek_nama = request.form.get('rek_nama', '').strip()
            rek_no   = request.form.get('rek_no', '').strip()
            if aid and rek_nama:
                conn.execute("UPDATE akun SET nama=?, no_rekening=? WHERE id=? AND is_rekening=1",
                             (rek_nama, rek_no or None, aid))
                conn.commit()
                flash('Rekening berhasil diperbarui!', 'success')
        elif action == 'toggle_rekening':
            aid = int(request.form.get('akun_id',0))
            a = conn.execute("SELECT * FROM akun WHERE id=?", (aid,)).fetchone()
            if a and a['kode'] not in ('1100','1110'):
                new_flag = 0 if a['is_rekening'] else 1
                conn.execute("UPDATE akun SET is_rekening=? WHERE id=?", (new_flag, aid))
                conn.commit()
                flash('Status rekening diperbarui.','info')
        conn.close()
        return redirect(url_for('settings'))

    s_nama_usaha   = get_setting(conn, 'nama_usaha','Usaha Saya')
    s_tz_app = get_setting(conn, 'tz_app', TZ_APP_DEFAULT)
    if s_tz_app not in TZ_APP_PILIHAN:
        s_tz_app = TZ_APP_DEFAULT
    pos_struk = _pos_receipt_settings(conn)
    users = []
    for row in conn.execute("SELECT * FROM users ORDER BY role, username").fetchall():
        u = dict(row)
        custom = _decode_permissions(u.get('permissions'))
        u['permissions_customized'] = custom is not None
        u['effective_permissions'] = sorted(effective_permissions(u.get('role'), u.get('permissions')))
        users.append(u)
    akun_kas = conn.execute("SELECT * FROM akun WHERE is_rekening=1 ORDER BY kode").fetchall()
    rekening_list = get_rekening_saldo(conn)
    conn.close()
    return render_template('settings.html',
        s_nama_usaha=s_nama_usaha, s_tz_app=s_tz_app, tz_app_pilihan=TZ_APP_PILIHAN,
        tz_bisa_diatur=hasattr(time, 'tzset'),
        users=users, akun_kas=akun_kas,
        rekening_list=rekening_list, pos_struk=pos_struk,
        permissions_avail=USER_PERMISSIONS)

@app.route('/settings/simple-mode', methods=['POST'])
@admin_required
def toggle_simple_mode():
    conn = db()
    requested_mode = request.form.get('mode', '').strip().lower()
    if requested_mode in ('simple', 'detail'):
        new_value = '1' if requested_mode == 'simple' else '0'
    else:
        current = get_setting(conn, 'pref_simple_mode', '0') == '1'
        new_value = '0' if current else '1'
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pref_simple_mode', ?)", (new_value,))
    conn.commit()
    conn.close()
    reload_app_prefs()
    flash(f"Mode tampilan diubah ke {'Simple' if new_value == '1' else 'Detail'}.", 'success')
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/settings/pos-default-payment', methods=['POST'])
@admin_required
def update_pos_default_payment():
    default_payment = request.form.get('pos_default_payment', 'TUNAI')
    if default_payment not in ('TUNAI', 'REKENING', 'PIUTANG'):
        default_payment = 'TUNAI'
    conn = db()
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pos_default_payment', ?)",
                 (default_payment,))
    conn.commit()
    conn.close()
    label = {'TUNAI': 'Tunai', 'REKENING': 'Transfer / QRIS / EDC', 'PIUTANG': 'Piutang Customer'}[default_payment]
    flash(f'Default pembayaran POS diubah ke {label}.', 'success')
    return redirect(request.referrer or url_for('dashboard'))


# ---------- ASET TETAP ----------
@app.route('/aset-tetap')
@investor_required
def aset_tetap_list():
    conn = db()
    today = date.today()
    linked_legacy = backfill_asset_journal_links(conn)
    raw = conn.execute("""
        SELECT a.*, CASE WHEN a.jurnal_id IS NULL OR j.id IS NULL THEN 1 ELSE 0 END AS source_missing
        FROM aset_tetap a
        LEFT JOIN jurnal j ON j.id=a.jurnal_id
        ORDER BY a.tanggal_beli DESC, a.id DESC
    """).fetchall()
    aset_list = []
    for a in raw:
        tgl_beli = datetime.strptime(str(a['tanggal_beli'])[:10], '%Y-%m-%d').date()
        months_elapsed = (today.year - tgl_beli.year) * 12 + (today.month - tgl_beli.month) + 1
        bulan_real = min(months_elapsed, a['masa_pakai'])
        peny_real  = round(bulan_real * (a['penyusutan_bulan'] or 0), 0)
        bulan_ctt  = a['bulan_penyusutan_dicatat'] or 0
        row = dict(a)
        row['source_missing'] = bool(row.get('source_missing'))
        row['bulan_real']  = bulan_real
        row['peny_real']   = peny_real
        row['bulan_ctt']   = bulan_ctt
        aset_list.append(row)
    orphan_count = sum(1 for row in aset_list if row.get('source_missing'))
    if linked_legacy:
        conn.commit()
    conn.close()
    return render_template('aset_tetap.html', aset_list=aset_list, today=today,
                           orphan_count=orphan_count)


@app.route('/aset-tetap/bersihkan-yatim', methods=['POST'])
@finance_required
def aset_tetap_bersihkan_yatim():
    conn = db()
    backfill_asset_journal_links(conn)
    orphans = conn.execute("""
        SELECT a.id,a.nama
        FROM aset_tetap a
        LEFT JOIN jurnal j ON j.id=a.jurnal_id
        WHERE a.jurnal_id IS NULL OR j.id IS NULL
        ORDER BY a.id
    """).fetchall()
    if not orphans:
        conn.commit(); conn.close()
        flash('Tidak ada aset tanpa transaksi sumber yang perlu dibersihkan.', 'info')
        return redirect(url_for('aset_tetap_list'))

    dep_deleted = 0
    for aset in orphans:
        dep_deleted += _purge_asset_depreciation(conn, aset['id'], aset['nama'])
        conn.execute("DELETE FROM aset_tetap WHERE id=?", (aset['id'],))
    add_log(conn, 'Bersihkan aset tanpa transaksi',
            f"{len(orphans)} aset dihapus dari daftar aset tetap; {dep_deleted} jurnal penyusutan dibersihkan",
            'HAPUS')
    conn.commit(); conn.close()
    flash(f'{len(orphans)} aset tanpa transaksi sumber dibersihkan dari Daftar Aset Tetap.', 'warning')
    return redirect(url_for('aset_tetap_list'))


# ---------- AKUN ----------
@app.route('/akun')
@investor_required
def akun_list():
    conn = db()
    rows = conn.execute("""
        SELECT a.*, COALESCE((SELECT COUNT(*) FROM detail_jurnal dj WHERE dj.akun_id=a.id),0) AS pakai
        FROM akun a ORDER BY a.kode
    """).fetchall()
    conn.close()
    return render_template('akun.html', rows=rows)

@app.route('/akun/baru', methods=['POST'])
@finance_required
def akun_baru():
    conn = db()
    tipe = (request.form.get('tipe','') or '').strip().upper()
    if tipe not in ('ASET','LIABILITAS','EKUITAS','PENDAPATAN','BEBAN'):
        conn.close()
        flash('Tipe akun tidak valid. Pilih: ASET, LIABILITAS, EKUITAS, PENDAPATAN, atau BEBAN.', 'danger')
        return redirect(url_for('akun_list'))
    saldo_normal = request.form['saldo_normal']
    # Penanda rekening kas/bank. Inilah yang menentukan akun muncul di pilihan
    # "rekening penerimaan/pembayaran" di seluruh form (is_rekening=1). Dulu akun
    # buatan user selalu 0 tanpa cara menyalakannya, jadi akun bank baru bisa diisi
    # saldo awal tapi tak pernah bisa dipilih saat terima/bayar uang.
    # Dibatasi lewat boleh_jadi_rekening() supaya Piutang/PPN Masukan tak salah ditandai.
    is_rek = 1 if (request.form.get('is_rekening') == '1'
                   and boleh_jadi_rekening(request.form['kode'], tipe, saldo_normal)) else 0
    try:
        conn.execute(
            'INSERT INTO akun(kode,nama,tipe,subtipe,saldo_normal,is_rekening) VALUES(?,?,?,?,?,?)',
            (request.form['kode'], request.form['nama'], tipe,
             request.form.get('subtipe',''), saldo_normal, is_rek)
        )
        conn.commit()
        flash('Akun ditambahkan!' + (' Akun ini aktif sebagai rekening kas/bank.' if is_rek else ''), 'success')
    except sqlite3.IntegrityError:
        flash('Kode akun sudah ada!','danger')
    finally:
        conn.close()
    return redirect(url_for('akun_list'))

@app.route('/setup/saldo-awal', methods=['GET'])
@finance_required
def saldo_awal_form():
    """Tampilkan wizard input saldo awal untuk semua akun (kas, piutang, hutang, aset tetap)."""
    conn = db()
    # Akun ASET (kecuali persediaan 1130 yang sudah punya alur sendiri, dan akun aset tetap 12xx yang masuk section terpisah)
    aset_rows = conn.execute("""
        SELECT * FROM akun
        WHERE tipe='ASET' AND kode NOT IN ('1130','1200','1210','1220','1290')
        ORDER BY kode
    """).fetchall()
    # Akun aset tetap kontra (kode 1290 Akumulasi Penyusutan), preview saja
    aset_tetap_akun = conn.execute("""
        SELECT * FROM akun WHERE kode IN ('1200','1210','1220','1290') ORDER BY kode
    """).fetchall()
    # Akun LIABILITAS
    liab_rows = conn.execute("""
        SELECT * FROM akun WHERE tipe='LIABILITAS' ORDER BY kode
    """).fetchall()
    # Cek apakah sudah pernah ada saldo awal akun → tampilkan warning
    sudah_ada = conn.execute(
        "SELECT COUNT(*) FROM jurnal WHERE tipe_tx='SALDO_AWAL_AKUN'"
    ).fetchone()[0]
    # Saldo terkini per akun (untuk preview)
    saldo_now = {}
    for r in (list(aset_rows) + list(liab_rows) + list(aset_tetap_akun)):
        row = conn.execute("""
            SELECT COALESCE(SUM(dj.debit),0) AS td, COALESCE(SUM(dj.kredit),0) AS tk
            FROM detail_jurnal dj JOIN akun a ON a.id=dj.akun_id
            WHERE a.kode=?
        """, (r['kode'],)).fetchone()
        saldo = (row['td']-row['tk']) if r['saldo_normal']=='DEBIT' else (row['tk']-row['td'])
        saldo_now[r['kode']] = float(saldo or 0)
    conn.close()
    today = date.today()
    default_tgl = f"{today.year}-01-01"
    return render_template('saldo_awal.html',
                           aset_rows=aset_rows, liab_rows=liab_rows,
                           aset_tetap_akun=aset_tetap_akun,
                           saldo_now=saldo_now, sudah_ada=sudah_ada,
                           default_tgl=default_tgl, today=today)

@app.route('/setup/saldo-awal', methods=['POST'])
@finance_required
def saldo_awal_simpan():
    """Buat 1 jurnal SALDO_AWAL_AKUN yang berisi semua entry saldo awal + counter ke Modal Pemilik."""
    conn = db()
    tanggal = request.form.get('tanggal', '') or str(date.today())
    if not is_valid_iso_date(tanggal):
        conn.close(); flash('Tanggal saldo awal tidak valid.', 'danger')
        return redirect(url_for('saldo_awal_form'))
    # Form ini SELALU membuat jurnal baru, tidak pernah menimpa. Selama saldo awal
    # cuma menyentuh jurnal, isi ulang "hanya" menggandakan jurnal. Sejak saldo awal
    # ikut membuat baris hutang/piutang, isi ulang akan menggandakan hutang di
    # Dashboard & Pelunasan juga, jadi harus dicegah di depan. Koreksi dilakukan
    # dengan menghapus jurnal saldo awal lama (yang sekaligus membersihkan trackernya
    # lewat _reverse_tx), baru mengisi ulang.
    _lama = conn.execute(
        "SELECT id, tanggal FROM jurnal WHERE tipe_tx='SALDO_AWAL_AKUN' ORDER BY id LIMIT 1"
    ).fetchone()
    if _lama:
        conn.close()
        flash('Saldo awal sudah pernah dicatat pada '
              f"{_lama['tanggal']} (Jurnal #{_lama['id']}). Untuk mengoreksinya, hapus "
              'dulu transaksi saldo awal itu lewat Daftar Transaksi, baru isi form ini '
              'lagi. Kalau diisi ulang tanpa menghapus, saldo dan hutangnya jadi dobel.',
              'warning')
        return redirect(url_for('saldo_awal_form'))
    entries = []
    total_d = 0.0
    total_k = 0.0
    # ── Aktifkan akun sbg rekening kas/bank (centang inline di baris wizard) ──
    # Dikirim sbg daftar KODE yang dicentang (bukan per-indeks) supaya tetap sejajar
    # walau checkbox yang tak dicentang tidak ikut ter-submit.
    rek_aktifkan = [k.strip() for k in request.form.getlist('aktifkan_rekening[]') if k.strip()]
    rek_sukses = []
    for kode in rek_aktifkan:
        a = conn.execute(
            "SELECT kode,nama,tipe,saldo_normal,is_rekening FROM akun WHERE kode=?", (kode,)
        ).fetchone()
        # Batas sama persis dgn Bagan Akun (satu aturan bersama).
        if a and not a['is_rekening'] and boleh_jadi_rekening(a['kode'], a['tipe'], a['saldo_normal']):
            conn.execute("UPDATE akun SET is_rekening=1 WHERE kode=?", (kode,))
            rek_sukses.append(f"{a['kode']} {a['nama']}")
    if rek_sukses:
        # Commit terpisah & lebih dulu: aktivasi rekening harus tetap tersimpan walau
        # bagian saldo awal di bawah gagal validasi atau memang sengaja dikosongkan
        # (user cuma ingin mengaktifkan rekening, tanpa mengisi saldo pembukaan).
        conn.commit()
        flash('Sekarang bisa dipilih saat terima atau bayar uang: ' + ', '.join(rek_sukses) + '.', 'success')
    # Nominal WAJIB lewat parse_rp(), jangan bikin parser sendiri di sini.
    # Parser lama `float(x.replace('.','').replace(',','.'))` membuang SEMUA titik,
    # padahal normalizer JS di base.html mengirim desimal memakai TITIK ("10000.05").
    # Akibatnya isian 10.000,05 tersimpan jadi 1.000.005 alias 100x lipat, dan hanya
    # di form ini karena form lain memang sudah memakai parse_rp().
    def _nilai(x):
        try:
            return parse_rp(x)
        except (ValueError, TypeError):
            return 0.0

    # ── Rincian per pihak (OPSIONAL) ────────────────────────────────────────
    # Saldo awal hutang/piutang dulu HANYA masuk jurnal, tanpa baris di tabel
    # `hutang`/`piutang`. Akibatnya: muncul di Neraca (baca jurnal) tapi HILANG
    # dari Dashboard, halaman Hutang/Piutang, dan Pelunasan (baca tabel tracker),
    # sehingga tidak pernah bisa dilunasi dan total hutang tampil lebih kecil dari
    # kenyataan. Di sini rinciannya dipanen dulu supaya baris tracker bisa dibuat
    # setelah jurnal sukses.
    #
    # Rincian ini adalah PENJABARAN dari nominal per akun, bukan tambahan. Kalau
    # dirinci, jumlahnya wajib sama dengan nominal akunnya; kalau nominal akun
    # dikosongkan, nominal itu diambil dari jumlah rincian. Tanpa aturan ini,
    # buku besar dan buku pembantu bisa berselisih sejak hari pertama.
    def _panen_rincian(prefix, akun_default):
        """[(nama, nilai, jatuh_tempo, akun_kode)] dari baris rincian di form."""
        nama_l = request.form.getlist(f'{prefix}_nama[]')
        nilai_l = request.form.getlist(f'{prefix}_nilai[]')
        jt_l = request.form.getlist(f'{prefix}_jt[]')
        akun_l = request.form.getlist(f'{prefix}_akun[]')
        keluar = []
        for i, nm in enumerate(nama_l):
            nm = (nm or '').strip()
            v = _nilai(nilai_l[i] if i < len(nilai_l) else 0)
            if not nm and v <= 0:
                continue                      # baris kosong, lewati
            if not nm:
                raise ValueError(f'Nama pihak pada rincian baris {i + 1} wajib diisi.')
            if v <= 0:
                raise ValueError(f'Nominal rincian "{nm}" harus lebih dari 0.')
            jt = (jt_l[i] if i < len(jt_l) else '') or None
            if not is_valid_optional_iso_date(jt):
                raise ValueError(f'Tanggal jatuh tempo rincian "{nm}" tidak valid.')
            kd = ((akun_l[i] if i < len(akun_l) else '') or '').strip() or akun_default
            keluar.append((nm, v, jt, kd))
        return keluar

    try:
        rinci_hutang = _panen_rincian('rh', '2100')
        rinci_piutang = _panen_rincian('rp', '1120')
    except ValueError as ex:
        conn.close(); flash(str(ex), 'danger')
        return redirect(url_for('saldo_awal_form'))

    # Jumlah rincian per kode akun, dipakai memvalidasi/mengisi nominal akun.
    rinci_total = {}
    for _nm, _v, _jt, _kd in (rinci_hutang + rinci_piutang):
        rinci_total[_kd] = rinci_total.get(_kd, 0.0) + _v

    def _nilai_akun(kode, nilai_form, label):
        """Nominal akun: dari form, atau dari jumlah rincian bila form kosong."""
        v = _nilai(nilai_form)
        rt = rinci_total.get(kode)
        if rt is None:
            return v
        if v <= 0:
            return rt                      # nominal akun diturunkan dari rincian
        if abs(v - rt) > 0.01:
            raise ValueError(
                f'{label} {kode}: nominal akun Rp {v:,.0f} tidak sama dengan jumlah '
                f'rinciannya Rp {rt:,.0f}. Samakan dulu, atau kosongkan nominal akun '
                f'supaya diambil dari rincian.')
        return v

    try:
        # ── Akun ASET (saldo Debit) ──
        aset_kode = request.form.getlist('aset_kode[]')
        aset_nilai = request.form.getlist('aset_nilai[]')
        for kode, nilai in zip(aset_kode, aset_nilai):
            v = _nilai_akun(kode, nilai, 'Aset')
            if v > 0:
                entries.append((kode, v, 0))
                total_d += v
        # ── Akun LIABILITAS (saldo Kredit) ──
        liab_kode = request.form.getlist('liab_kode[]')
        liab_nilai = request.form.getlist('liab_nilai[]')
        for kode, nilai in zip(liab_kode, liab_nilai):
            v = _nilai_akun(kode, nilai, 'Liabilitas')
            if v > 0:
                entries.append((kode, 0, v))
                total_k += v
    except ValueError as ex:
        conn.close(); flash(str(ex), 'danger')
        return redirect(url_for('saldo_awal_form'))

    # Tiap akun yang dirinci WAJIB ikut terbentuk barisnya di jurnal. Kalau tidak,
    # baris tracker akan lahir tanpa penopang di buku besar (hutang hantu).
    _akun_di_jurnal = {e[0] for e in entries}
    _yatim = sorted(set(rinci_total) - _akun_di_jurnal)
    if _yatim:
        conn.close()
        flash(f'Akun {", ".join(_yatim)} dirinci per pihak tapi tidak ada di daftar '
              f'saldo awal. Periksa kembali akun yang dipilih pada rincian.', 'danger')
        return redirect(url_for('saldo_awal_form'))
    # ── Aset Tetap (multi-row) ──
    at_nama = request.form.getlist('at_nama[]')
    at_kat  = request.form.getlist('at_kat[]')
    at_harga = request.form.getlist('at_harga[]')
    at_tgl  = request.form.getlist('at_tgl[]')
    at_masa = request.form.getlist('at_masa[]')
    at_akum = request.form.getlist('at_akum[]')
    kat_kode = {'Peralatan': '1200', 'Kendaraan': '1210', 'Gedung': '1220'}
    aset_tetap_inserts = []  # tunda insert sampai jurnal sukses
    for nm, kat, hrg, tg, masa, akum in zip(at_nama, at_kat, at_harga, at_tgl, at_masa, at_akum):
        nm = (nm or '').strip()
        if not nm: continue
        try:
            hrg_v  = parse_rp(hrg or '0')
            akum_v = parse_rp(akum or '0')
            masa_v = int(masa or 12)
        except (ValueError, TypeError):
            conn.close(); flash(f'Format nominal aset tetap "{nm}" tidak valid.', 'danger')
            return redirect(url_for('saldo_awal_form'))
        if hrg_v <= 0: continue
        if kat not in kat_kode:
            conn.close(); flash(f'Kategori aset "{kat}" tidak valid.', 'danger')
            return redirect(url_for('saldo_awal_form'))
        if not is_valid_iso_date(tg or tanggal):
            tg = tanggal
        if akum_v > hrg_v:
            conn.close(); flash(f'Akumulasi penyusutan untuk "{nm}" tidak boleh melebihi harga perolehan.', 'danger')
            return redirect(url_for('saldo_awal_form'))
        peny_bln = round(hrg_v / max(masa_v, 1), 2)
        entries.append((kat_kode[kat], hrg_v, 0))
        total_d += hrg_v
        if akum_v > 0:
            entries.append(('1290', 0, akum_v))
            total_k += akum_v
        # bulan_penyusutan_dicatat = round(akum_v / peny_bln) supaya tidak double-counted nanti
        bulan_dicatat = round(akum_v / peny_bln) if peny_bln > 0 else 0
        aset_tetap_inserts.append((nm, kat, hrg_v, tg or tanggal, masa_v, peny_bln, bulan_dicatat, akum_v))
    if not entries:
        conn.close(); flash('Tidak ada nilai yang diisi. Isi minimal satu akun saldo awal.', 'warning')
        return redirect(url_for('saldo_awal_form'))
    # ── Auto-balance ke 3100 Modal Pemilik ──
    selisih = total_d - total_k
    if abs(selisih) > 0.01:
        if selisih > 0:
            entries.append(('3100', 0, selisih))  # Aset > Liabilitas → modal masuk kredit
        else:
            entries.append(('3100', -selisih, 0))  # rare: Liabilitas > Aset → modal negatif (defisit)
    keterangan = 'Saldo Awal Akun (setup pembukuan)'
    try:
        jid = insert_jurnal(conn, tanggal, keterangan, 'PENDANAAN', 'SALDO_AWAL_AKUN', entries)
        # Insert ke tabel aset_tetap
        for nm, kat, hrg, tg, masa, peny, bulan_dicatat, akum in aset_tetap_inserts:
            conn.execute(
                "INSERT INTO aset_tetap(nama,kategori,harga_beli,tanggal_beli,masa_pakai,penyusutan_bulan,bulan_penyusutan_dicatat,akumulasi_penyusutan,jurnal_id) VALUES(?,?,?,?,?,?,?,?,?)",
                (nm, kat, hrg, tg, masa, peny, bulan_dicatat, akum, jid)
            )
        # ── Baris tracker hutang/piutang ────────────────────────────────────
        # Ditandai jurnal_id saldo awal, jadi _reverse_tx (yang sudah menghapus
        # piutang/hutang + riwayat bayarnya per jurnal_id) otomatis membersihkan
        # ini kalau jurnal saldo awalnya dihapus. Tidak perlu pembersih khusus.
        n_tracker = 0
        for nm, v, jt, kd in rinci_hutang:
            conn.execute(
                "INSERT INTO hutang(tanggal,jatuh_tempo,pemasok,keterangan,jumlah,terbayar,"
                "status,jurnal_id,akun_kode) VALUES(?,?,?,?,?,0,'BELUM LUNAS',?,?)",
                (tanggal, jt, nm, 'Saldo awal hutang', v, jid, kd))
            n_tracker += 1
        for nm, v, jt, kd in rinci_piutang:
            conn.execute(
                "INSERT INTO piutang(tanggal,jatuh_tempo,pelanggan,keterangan,jumlah,terbayar,"
                "status,jurnal_id) VALUES(?,?,?,?,?,0,'BELUM LUNAS',?)",
                (tanggal, jt, nm, 'Saldo awal piutang', v, jid))
            n_tracker += 1
        # Tanpa rincian, saldo akun tracker tetap dibuatkan SATU baris gelondongan
        # supaya tetap bisa dilunasi lewat menu Pelunasan (itu keluhan aslinya).
        _sudah_dirinci = set(rinci_total)
        for kode, d, k in entries:
            if kode in _sudah_dirinci:
                continue
            if kode == '2100' and k > 0:
                conn.execute(
                    "INSERT INTO hutang(tanggal,jatuh_tempo,pemasok,keterangan,jumlah,terbayar,"
                    "status,jurnal_id,akun_kode) VALUES(?,NULL,?,?,?,0,'BELUM LUNAS',?,?)",
                    (tanggal, 'Saldo Awal', 'Saldo awal hutang (belum dirinci per pemasok)',
                     k, jid, kode))
                n_tracker += 1
            elif kode == '1120' and d > 0:
                conn.execute(
                    "INSERT INTO piutang(tanggal,jatuh_tempo,pelanggan,keterangan,jumlah,terbayar,"
                    "status,jurnal_id) VALUES(?,NULL,?,?,?,0,'BELUM LUNAS',?)",
                    (tanggal, 'Saldo Awal', 'Saldo awal piutang (belum dirinci per pelanggan)',
                     d, jid))
                n_tracker += 1
        add_log(conn, 'Saldo Awal Akun',
                f"{tanggal} | {len(entries)} baris | total Rp {max(total_d, total_k):,.0f}"
                + (f" | {n_tracker} baris hutang/piutang" if n_tracker else ''),
                'INPUT')
        conn.commit()
        _pesan_tracker = (f' {n_tracker} baris hutang/piutang ikut dibuat dan sudah bisa '
                          f'dilunasi lewat menu Pelunasan.') if n_tracker else ''
        flash(f'Saldo awal berhasil dicatat. {len(entries)} entri jurnal dibuat. '
              f'Modal Pemilik auto-balance Rp {abs(selisih):,.0f}.{_pesan_tracker}', 'success')
    except ValueError as ex:
        conn.rollback()
        flash(f'Gagal menyimpan saldo awal: {ex}', 'danger')
    finally:
        conn.close()
    return redirect(url_for('neraca'))

@app.route('/akun/<int:id>/edit', methods=['POST'])
@finance_required
def akun_edit(id):
    conn = db()
    akun_row = conn.execute('SELECT * FROM akun WHERE id=?', (id,)).fetchone()
    if not akun_row:
        conn.close(); flash('Akun tidak ditemukan.', 'danger')
        return redirect(url_for('akun_list'))
    kode_baru = (request.form.get('kode', '') or '').strip()
    nama_baru = (request.form.get('nama', '') or '').strip()
    subtipe_baru = (request.form.get('subtipe', '') or '').strip()
    tipe_baru = request.form.get('tipe', akun_row['tipe'])
    saldo_normal_baru = request.form.get('saldo_normal', akun_row['saldo_normal'])
    if not kode_baru or not nama_baru:
        conn.close(); flash('Kode dan Nama akun wajib diisi.', 'warning')
        return redirect(url_for('akun_list'))
    if tipe_baru not in ('ASET', 'LIABILITAS', 'EKUITAS', 'PENDAPATAN', 'BEBAN'):
        conn.close(); flash('Tipe akun tidak valid.', 'warning')
        return redirect(url_for('akun_list'))
    if saldo_normal_baru not in ('DEBIT', 'KREDIT'):
        conn.close(); flash('Saldo normal tidak valid.', 'warning')
        return redirect(url_for('akun_list'))
    # Cek apakah akun sudah dipakai di jurnal
    dipakai = conn.execute('SELECT COUNT(*) FROM detail_jurnal WHERE akun_id=?', (id,)).fetchone()[0]
    # Akun bawaan: KODE dikunci selamanya (walau belum pernah dipakai di jurnal). Mesin jurnal
    # merujuk kode-kode ini secara harfiah, jadi mengganti kodenya sama merusaknya dengan
    # menghapusnya. Nama & subtipe tetap boleh diubah supaya user bisa menyesuaikan istilah.
    akun_bawaan = akun_row['kode'] in AKUN_DEFAULT_KODE
    if akun_bawaan:
        _ubah = (kode_baru != akun_row['kode'] or tipe_baru != akun_row['tipe']
                 or saldo_normal_baru != akun_row['saldo_normal'])
        kode_baru = akun_row['kode']
        tipe_baru = akun_row['tipe']
        saldo_normal_baru = akun_row['saldo_normal']
        if _ubah:
            flash(f"Akun bawaan {akun_row['kode']}: kode, tipe, dan saldo normal dikunci karena "
                  "dipakai otomatis oleh sistem saat mencatat transaksi dan menyusun Neraca. "
                  "Nama akun tetap bisa Anda sesuaikan.", 'warning')
    # Jika sudah dipakai, kode & tipe & saldo_normal dikunci agar laporan historis tidak rusak
    if dipakai > 0:
        kode_baru = akun_row['kode']
        tipe_baru = akun_row['tipe']
        saldo_normal_baru = akun_row['saldo_normal']
    else:
        # Cek bentrok kode
        bentrok = conn.execute('SELECT id FROM akun WHERE kode=? AND id<>?', (kode_baru, id)).fetchone()
        if bentrok:
            conn.close(); flash(f'Kode {kode_baru} sudah dipakai akun lain.', 'danger')
            return redirect(url_for('akun_list'))
    # Penanda rekening kas/bank (lihat komentar di akun_baru). Kas 1100 & Bank 1110
    # dikunci tetap rekening supaya form pembayaran tak pernah kehabisan pilihan.
    if akun_row['kode'] in ('1100', '1110'):
        is_rek_baru = 1
    else:
        is_rek_baru = 1 if (request.form.get('is_rekening') == '1'
                            and boleh_jadi_rekening(kode_baru, tipe_baru, saldo_normal_baru)) else 0
    rek_berubah = int(akun_row['is_rekening'] or 0) != is_rek_baru
    # Peran akun di Laba Rugi. Hanya berlaku utk BEBAN; akun lain selalu kosong supaya
    # tidak ada aset/pendapatan yang tanpa sengaja ikut dihitung sbg pajak/bunga/penyusutan.
    peran_lama = (akun_row['peran'] if 'peran' in akun_row.keys() else '') or ''
    peran_baru = (request.form.get('peran', '') or '').strip().upper()
    if peran_baru not in ('', 'PENYUSUTAN', 'BUNGA', 'PAJAK') or tipe_baru != 'BEBAN':
        peran_baru = '' if tipe_baru != 'BEBAN' else peran_lama
    try:
        conn.execute(
            'UPDATE akun SET kode=?, nama=?, tipe=?, subtipe=?, saldo_normal=?, is_rekening=?, peran=? WHERE id=?',
            (kode_baru, nama_baru, tipe_baru, subtipe_baru, saldo_normal_baru, is_rek_baru, peran_baru, id)
        )
        conn.commit()
        if peran_baru != peran_lama:
            _nm = {'PENYUSUTAN': 'Penyusutan', 'BUNGA': 'Bunga', 'PAJAK': 'Pajak'}
            if peran_baru:
                flash(f'Akun ini sekarang dihitung sebagai {_nm[peran_baru]} di Laba Rugi, '
                      'jadi tidak lagi masuk Beban Operasional.', 'info')
            else:
                flash('Akun ini sekarang dihitung sebagai Beban Operasional biasa di Laba Rugi.', 'info')
        if rek_berubah:
            flash(('Akun ini sekarang aktif sebagai rekening kas/bank dan sudah bisa dipilih '
                   'saat terima atau bayar uang.') if is_rek_baru else
                  ('Akun ini dinonaktifkan sebagai rekening kas/bank, jadi tidak lagi muncul '
                   'di pilihan pembayaran dan tidak dihitung sebagai Kas & Bank.'),
                  'success' if is_rek_baru else 'warning')
        if dipakai > 0:
            flash(f'Akun diperbarui. Kode/tipe/saldo normal dikunci karena akun sudah dipakai di {dipakai} entri jurnal.', 'info')
        else:
            flash('Akun berhasil diperbarui.', 'success')
        add_log(conn, 'Edit akun', f"{akun_row['kode']} {akun_row['nama']} → {kode_baru} {nama_baru}", 'EDIT')
        conn.commit()
    except sqlite3.IntegrityError as ex:
        flash(f'Gagal memperbarui akun: {ex}', 'danger')
    finally:
        conn.close()
    return redirect(url_for('akun_list'))

@app.route('/akun/<int:id>/hapus', methods=['POST'])
@finance_required
def akun_hapus(id):
    conn = db()
    row = conn.execute('SELECT kode,nama FROM akun WHERE id=?', (id,)).fetchone()
    if not row:
        conn.close(); flash('Akun tidak ditemukan.', 'danger')
        return redirect(url_for('akun_list'))
    used = conn.execute('SELECT COUNT(*) FROM detail_jurnal WHERE akun_id=?',(id,)).fetchone()[0]
    # Akun bawaan dipakai langsung oleh mesin jurnal (mis. 1130 Persediaan untuk semua
    # transaksi barang). Kalau dihapus, penyimpanan transaksi gagal "Akun tidak ditemukan".
    # Cek ini WAJIB di depan: akun bawaan yang belum pernah terpakai pun tidak boleh hilang.
    if row['kode'] in AKUN_DEFAULT_KODE:
        flash(f"Akun {row['kode']} {row['nama']} adalah akun bawaan sistem dan tidak bisa dihapus. "
              "Akun ini dipakai otomatis saat mencatat transaksi. Kalau tidak terpakai, biarkan saja.",
              'danger')
    elif used > 0:
        flash('Akun sudah digunakan, tidak bisa dihapus!','danger')
    else:
        conn.execute('DELETE FROM akun WHERE id=?',(id,))
        add_log(conn, 'Hapus akun', f"{row['kode']} - {row['nama']}", 'EDIT')
        conn.commit(); flash('Akun dihapus.','warning')
    conn.close()
    return redirect(url_for('akun_list'))


# ---------- EXPORT ----------
@app.route('/export/dashboard')
@investor_required
def export_dashboard():
    today = date.today()
    sd = request.args.get('sd', today.replace(day=1).strftime('%Y-%m-%d'))
    ed = request.args.get('ed', today.strftime('%Y-%m-%d'))
    conn = db()
    pnl = calc_profitability(conn, sd, ed)
    cf  = calc_cashflow(conn, sd, ed)
    conn.close()

    if HAS_XLSX:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Dashboard"
        hdr = Font(bold=True)
        ws.append(["Laporan Keuangan", f"{sd} s/d {ed}"])
        ws.append([])
        ws.append(["CASHFLOW"])
        ws.append(["Uang Masuk", cf['masuk']])
        ws.append(["Uang Keluar", cf['keluar']])
        ws.append(["Saldo", cf['saldo']])
        ws.append([])
        ws.append(["PROFITABILITAS", "Nominal", "%"])
        items = [
            ("Pendapatan", pnl['rev'], 100),
            ("HPP", pnl['hpp'], pnl['pct_hpp']),
            ("Laba Kotor", pnl['laba_kotor'], pnl['pct_laba_kotor']),
            ("Biaya Operasional", pnl['op_exp'], pnl['pct_op']),
            ("Laba Operasional", pnl['laba_op'], pnl['pct_laba_op']),
            ("Penyusutan", pnl['depr'], pnl['pct_depr']),
            ("EBIT", pnl['ebit'], pnl['pct_ebit']),
            ("Pajak", pnl['tax'], pnl['pct_tax']),
            ("Laba Bersih", pnl['laba_bersih'], pnl['pct_laba_bersih']),
            ("Penarikan Owner", pnl['prive'], pnl['pct_prive']),
            ("Laba Ditahan", pnl['laba_tahan'], pnl['pct_laba_tahan']),
        ]
        for row in items:
            ws.append(list(row))
        buf = io.BytesIO()
        wb.save(buf); buf.seek(0)
        return send_file(buf, download_name=f"dashboard_{sd}_{ed}.xlsx",
                         as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # CSV fallback
    lines = [f"Laporan Keuangan {sd} s/d {ed}"]
    lines += ["","CASHFLOW",f"Uang Masuk,{cf['masuk']}",f"Uang Keluar,{cf['keluar']}",f"Saldo,{cf['saldo']}"]
    lines += ["","PROFITABILITAS,Nominal,%"]
    for name, val, pct in [("Pendapatan",pnl['rev'],100),("HPP",pnl['hpp'],pnl['pct_hpp']),
                            ("Laba Kotor",pnl['laba_kotor'],pnl['pct_laba_kotor']),
                            ("Biaya Operasional",pnl['op_exp'],pnl['pct_op']),
                            ("Laba Operasional",pnl['laba_op'],pnl['pct_laba_op']),
                            ("Penyusutan",pnl['depr'],pnl['pct_depr']),
                            ("EBIT",pnl['ebit'],pnl['pct_ebit']),
                            ("Pajak",pnl['tax'],pnl['pct_tax']),
                            ("Laba Bersih",pnl['laba_bersih'],pnl['pct_laba_bersih']),
                            ("Penarikan Owner",pnl['prive'],pnl['pct_prive']),
                            ("Laba Ditahan",pnl['laba_tahan'],pnl['pct_laba_tahan'])]:
        lines.append(f"{name},{val},{pct}%")
    csv_data = '\n'.join(lines)
    return Response(csv_data, mimetype='text/csv',
                    headers={'Content-Disposition': f'attachment; filename=dashboard_{sd}_{ed}.csv'})

@app.route('/export/piutang')
@investor_required
def export_piutang():
    conn = db()
    rows = conn.execute("SELECT * FROM piutang ORDER BY status, jatuh_tempo").fetchall()
    conn.close()
    if HAS_XLSX:
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Piutang"
        ws.append(["Tanggal","Jatuh Tempo","Pelanggan","Keterangan","Jumlah","Terbayar","Sisa","Status"])
        for r in rows:
            ws.append([r['tanggal'],r['jatuh_tempo'],r['pelanggan'],r['keterangan'],
                       r['jumlah'],r['terbayar'],r['jumlah']-r['terbayar'],r['status']])
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, download_name='piutang.xlsx', as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    lines = ["Tanggal,Jatuh Tempo,Pelanggan,Keterangan,Jumlah,Terbayar,Sisa,Status"]
    for r in rows:
        lines.append(f"{r['tanggal']},{r['jatuh_tempo']},{r['pelanggan']},{r['keterangan']},"
                     f"{r['jumlah']},{r['terbayar']},{r['jumlah']-r['terbayar']},{r['status']}")
    return Response('\n'.join(lines), mimetype='text/csv',
                    headers={'Content-Disposition':'attachment; filename=piutang.csv'})

@app.route('/export/hutang')
@investor_required
def export_hutang():
    conn = db()
    rows = conn.execute("SELECT * FROM hutang ORDER BY status, jatuh_tempo").fetchall()
    conn.close()
    if HAS_XLSX:
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Hutang"
        ws.append(["Tanggal","Jatuh Tempo","Pemasok","Keterangan","Jumlah","Terbayar","Sisa","Status"])
        for r in rows:
            ws.append([r['tanggal'],r['jatuh_tempo'],r['pemasok'],r['keterangan'],
                       r['jumlah'],r['terbayar'],r['jumlah']-r['terbayar'],r['status']])
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, download_name='hutang.xlsx', as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    lines = ["Tanggal,Jatuh Tempo,Pemasok,Keterangan,Jumlah,Terbayar,Sisa,Status"]
    for r in rows:
        lines.append(f"{r['tanggal']},{r['jatuh_tempo']},{r['pemasok']},{r['keterangan']},"
                     f"{r['jumlah']},{r['terbayar']},{r['jumlah']-r['terbayar']},{r['status']}")
    return Response('\n'.join(lines), mimetype='text/csv',
                    headers={'Content-Disposition':'attachment; filename=hutang.csv'})


@app.route('/export/laporan-keuangan')
@investor_required
def export_laporan_keuangan():
    today = date.today()
    sd = request.args.get('sd', today.replace(day=1).strftime('%Y-%m-%d'))
    ed = request.args.get('ed', today.strftime('%Y-%m-%d'))

    conn = db()

    # ── Data Laba Rugi ──────────────────────────────────────────────────────
    pnl = calc_profitability(conn, sd, ed)
    lr_rows = conn.execute("""
        SELECT a.kode, a.nama, a.tipe, a.subtipe,
               COALESCE(SUM(CASE WHEN j.tanggal >= ? AND j.tanggal <= ?
                                 THEN (CASE WHEN a.saldo_normal='DEBIT' THEN d.debit-d.kredit
                                            ELSE d.kredit-d.debit END) ELSE 0 END),0) as saldo
        FROM akun a
        LEFT JOIN detail_jurnal d ON d.akun_id = a.id
        LEFT JOIN jurnal j ON j.id = d.jurnal_id
        WHERE a.tipe IN ('PENDAPATAN','BEBAN')
        GROUP BY a.id ORDER BY a.kode
    """, (sd, ed)).fetchall()
    pendapatan = {}; beban = {}
    for r in lr_rows:
        item = dict(kode=r['kode'], nama=r['nama'], saldo=r['saldo'])
        if r['tipe'] == 'PENDAPATAN':
            pendapatan.setdefault(r['subtipe'] or 'Pendapatan', []).append(item)
        else:
            beban.setdefault(r['subtipe'] or 'Beban', []).append(item)

    # ── Data Arus Kas ───────────────────────────────────────────────────────
    kas_ids = get_rekening_ids(conn)
    saldo_awal = 0; operasional = []; investasi = []; pendanaan_list = []
    if kas_ids:
        ph = ','.join('?' * len(kas_ids))
        saldo_awal = conn.execute(f"""
            SELECT COALESCE(SUM(d.debit-d.kredit),0)
            FROM detail_jurnal d JOIN jurnal j ON j.id=d.jurnal_id
            WHERE d.akun_id IN ({ph}) AND j.tanggal<?
        """, kas_ids + [sd]).fetchone()[0]
        flows = conn.execute(f"""
            SELECT j.id, j.tanggal, j.keterangan, j.kategori,
                   SUM(CASE WHEN d.akun_id IN ({ph}) THEN d.debit-d.kredit ELSE 0 END) as net_kas
            FROM jurnal j JOIN detail_jurnal d ON d.jurnal_id=j.id
            WHERE j.tanggal>=? AND j.tanggal<=?
            GROUP BY j.id HAVING net_kas!=0 ORDER BY j.tanggal, j.id
        """, kas_ids + [sd, ed]).fetchall()
        operasional   = [r for r in flows if r['kategori'] == 'OPERASIONAL']
        investasi     = [r for r in flows if r['kategori'] == 'INVESTASI']
        pendanaan_list = [r for r in flows if r['kategori'] == 'PENDANAAN']
    total_op   = sum(r['net_kas'] for r in operasional)
    total_inv  = sum(r['net_kas'] for r in investasi)
    total_pend = sum(r['net_kas'] for r in pendanaan_list)
    saldo_akhir = saldo_awal + total_op + total_inv + total_pend

    # ── Data Neraca (per ed) ────────────────────────────────────────────────
    nr_rows = conn.execute("""
        SELECT a.kode, a.nama, a.tipe, a.subtipe, a.saldo_normal,
               COALESCE(SUM(CASE WHEN j.tanggal <= ? THEN d.debit ELSE 0 END),0) as td,
               COALESCE(SUM(CASE WHEN j.tanggal <= ? THEN d.kredit ELSE 0 END),0) as tk
        FROM akun a
        LEFT JOIN detail_jurnal d ON d.akun_id = a.id
        LEFT JOIN jurnal j ON j.id = d.jurnal_id
        GROUP BY a.id ORDER BY a.kode
    """, (ed, ed)).fetchall()
    aset = {}; liab = {}; ekuitas_list = []
    total_aset = total_liab = total_ekuitas = 0
    for r in nr_rows:
        saldo = (r['td'] - r['tk']) if r['saldo_normal'] == 'DEBIT' else (r['tk'] - r['td'])
        item = dict(kode=r['kode'], nama=r['nama'], saldo=saldo)
        if r['tipe'] == 'ASET':
            aset_saldo = -saldo if r['saldo_normal'] == 'KREDIT' else saldo
            aset.setdefault(r['subtipe'] or 'Aset Lancar', []).append(
                dict(kode=r['kode'], nama=r['nama'], saldo=aset_saldo))
            total_aset += aset_saldo
        elif r['tipe'] == 'LIABILITAS':
            liab.setdefault(r['subtipe'] or 'Liabilitas Lancar', []).append(item)
            total_liab += saldo
        elif r['tipe'] == 'EKUITAS':
            eks = -saldo if r['saldo_normal'] == 'DEBIT' else saldo
            ekuitas_list.append(dict(kode=r['kode'], nama=r['nama'], saldo=eks))
            total_ekuitas += eks
    fiscal_start = f"{ed[:4]}-01-01"
    laba_tahun   = calc_profitability(conn, fiscal_start, ed)['laba_bersih']
    laba_all     = calc_profitability(conn, '2000-01-01', ed)['laba_bersih']
    laba_ditahan = laba_all - laba_tahun
    if abs(laba_ditahan) > 0.01:
        ekuitas_list.append(dict(kode='-', nama='Laba Ditahan (Akumulasi)', saldo=laba_ditahan))
        total_ekuitas += laba_ditahan
    ekuitas_list.append(dict(kode='-', nama=f'Laba Bersih Tahun {ed[:4]}', saldo=laba_tahun))
    total_ekuitas += laba_tahun

    conn.close()

    if not HAS_XLSX:
        flash('Library openpyxl tidak tersedia. Install dengan: pip install openpyxl', 'danger')
        return redirect(request.referrer or url_for('laba_rugi'))

    wb = openpyxl.Workbook()

    # ── Style helpers ───────────────────────────────────────────────────────
    def _hdr1(ws, text, row):
        c = ws.cell(row=row, column=1, value=text)
        c.font = Font(bold=True, size=14, color='FFFFFF')
        c.fill = PatternFill(patternType='solid', fgColor='1E293B')
        c.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[row].height = 22

    def _hdr2(ws, text, row):
        c = ws.cell(row=row, column=1, value=text)
        c.font = Font(bold=True, size=10, color='FFFFFF')
        c.fill = PatternFill(patternType='solid', fgColor='334155')
        ws.row_dimensions[row].height = 16

    def _subtipe_row(ws, text, row, ncols=3):
        c = ws.cell(row=row, column=1, value=text.upper())
        c.font = Font(bold=True, size=9, color='475569')
        c.fill = PatternFill(patternType='solid', fgColor='F1F5F9')
        for col in range(2, ncols + 1):
            ws.cell(row=row, column=col).fill = PatternFill(patternType='solid', fgColor='F1F5F9')

    def _total_row(ws, label, value, row, ncols=3, color='E2E8F0'):
        c = ws.cell(row=row, column=1, value=label)
        c.font = Font(bold=True, size=10)
        c.fill = PatternFill(patternType='solid', fgColor=color)
        v = ws.cell(row=row, column=ncols, value=value)
        v.font = Font(bold=True, size=10)
        v.number_format = '#,##0'
        v.fill = PatternFill(patternType='solid', fgColor=color)
        for col in range(2, ncols):
            ws.cell(row=row, column=col).fill = PatternFill(patternType='solid', fgColor=color)

    def _data_row(ws, kode, nama, value, row):
        ws.cell(row=row, column=1, value=kode).font = Font(color='94A3B8', size=9)
        ws.cell(row=row, column=2, value=nama).font = Font(size=10)
        v = ws.cell(row=row, column=3, value=value)
        v.number_format = '#,##0'
        v.font = Font(size=10, color='DC2626' if value < 0 else '000000')

    def _info_row(ws, label, value, row, bold=False):
        c = ws.cell(row=row, column=1, value=label)
        v = ws.cell(row=row, column=3, value=value)
        if bold:
            c.font = Font(bold=True, size=10)
            v.font = Font(bold=True, size=10)
        v.number_format = '#,##0'

    # ══════════════════════════════════════════════════════════════════════
    # SHEET 1: LABA RUGI
    # ══════════════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = 'Laba Rugi'
    ws1.column_dimensions['A'].width = 12
    ws1.column_dimensions['B'].width = 36
    ws1.column_dimensions['C'].width = 18

    r = 1
    _hdr1(ws1, 'LAPORAN LABA RUGI', r); r += 1
    ws1.cell(row=r, column=1, value=f'Periode: {sd}  s/d  {ed}').font = Font(italic=True, size=9, color='64748B'); r += 2

    # Pendapatan
    _hdr2(ws1, 'PENDAPATAN', r); r += 1
    for subtipe, items in pendapatan.items():
        _subtipe_row(ws1, subtipe, r); r += 1
        sub = 0
        for it in items:
            _data_row(ws1, it['kode'], it['nama'], it['saldo'], r); r += 1
            sub += it['saldo']
        _total_row(ws1, f'  Total {subtipe}', sub, r); r += 1
    _total_row(ws1, 'TOTAL PENDAPATAN', pnl['rev'], r, color='DBEAFE'); r += 2

    # Beban
    _hdr2(ws1, 'BEBAN', r); r += 1
    for subtipe, items in beban.items():
        _subtipe_row(ws1, subtipe, r); r += 1
        sub = 0
        for it in items:
            _data_row(ws1, it['kode'], it['nama'], it['saldo'], r); r += 1
            sub += it['saldo']
        _total_row(ws1, f'  Total {subtipe}', sub, r); r += 1
    _total_row(ws1, 'TOTAL BEBAN', pnl['hpp'] + pnl['op_exp'] + pnl['depr'] + pnl['interest'] + pnl['tax'], r, color='FEE2E2'); r += 2

    # Ringkasan P&L
    _hdr2(ws1, 'RINGKASAN', r); r += 1
    summary = [
        ('Pendapatan',          pnl['rev'],          False),
        ('HPP / Beban Pokok',   -pnl['hpp'],          False),
        ('Laba Kotor',          pnl['laba_kotor'],    True),
        ('Biaya Operasional',   -pnl['op_exp'],       False),
        ('Laba Operasional',    pnl['laba_op'],       True),
        ('Penyusutan',          -pnl['depr'],         False),
        ('EBIT',                pnl['ebit'],          True),
        ('Pajak',               -pnl['tax'],          False),
        ('Laba Bersih',         pnl['laba_bersih'],   True),
        ('Penarikan Owner',     -pnl['prive'],        False),
        ('Laba Ditahan',        pnl['laba_tahan'],    True),
    ]
    for label, val, bold in summary:
        _info_row(ws1, label, val, r, bold=bold); r += 1

    # ══════════════════════════════════════════════════════════════════════
    # SHEET 2: ARUS KAS
    # ══════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet('Arus Kas')
    ws2.column_dimensions['A'].width = 12
    ws2.column_dimensions['B'].width = 40
    ws2.column_dimensions['C'].width = 18

    r = 1
    _hdr1(ws2, 'LAPORAN ARUS KAS', r); r += 1
    ws2.cell(row=r, column=1, value=f'Periode: {sd}  s/d  {ed}').font = Font(italic=True, size=9, color='64748B'); r += 2

    ws2.cell(row=r, column=1, value='Saldo Kas Awal').font = Font(bold=True, size=10)
    v = ws2.cell(row=r, column=3, value=saldo_awal)
    v.number_format = '#,##0'; v.font = Font(bold=True, size=10)
    r += 2

    def _ak_section(ws, title, flows, total, row, hdr_color):
        c = ws.cell(row=row, column=1, value=title)
        c.font = Font(bold=True, size=10, color='FFFFFF')
        c.fill = PatternFill(patternType='solid', fgColor=hdr_color)
        ws.cell(row=row, column=2).fill = PatternFill(patternType='solid', fgColor=hdr_color)
        ws.cell(row=row, column=3).fill = PatternFill(patternType='solid', fgColor=hdr_color)
        row += 1
        for flow in flows:
            ws.cell(row=row, column=1, value=str(flow['tanggal'])[:10]).font = Font(size=9, color='94A3B8')
            ws.cell(row=row, column=2, value=flow['keterangan']).font = Font(size=9)
            v = ws.cell(row=row, column=3, value=flow['net_kas'])
            v.number_format = '#,##0'
            v.font = Font(size=9, color='15803D' if flow['net_kas'] >= 0 else 'DC2626')
            row += 1
        _total_row(ws, f'  Total {title}', total, row)
        row += 2
        return row

    r = _ak_section(ws2, 'Aktivitas Operasional', operasional, total_op, r, '0369A1')
    r = _ak_section(ws2, 'Aktivitas Investasi',   investasi,   total_inv, r, '0369A1')
    r = _ak_section(ws2, 'Aktivitas Pendanaan',   pendanaan_list, total_pend, r, '0369A1')

    ws2.cell(row=r, column=1, value='Saldo Kas Akhir').font = Font(bold=True, size=11)
    v = ws2.cell(row=r, column=3, value=saldo_akhir)
    v.number_format = '#,##0'
    v.font = Font(bold=True, size=11, color='1D4ED8')

    # ══════════════════════════════════════════════════════════════════════
    # SHEET 3: NERACA
    # ══════════════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet('Neraca')
    ws3.column_dimensions['A'].width = 12
    ws3.column_dimensions['B'].width = 36
    ws3.column_dimensions['C'].width = 18

    try:
        r = 1
        _hdr1(ws3, 'NERACA (BALANCE SHEET)', r); r += 1
        ws3.cell(row=r, column=1, value=f'Per tanggal: {ed}').font = Font(italic=True, size=9, color='64748B'); r += 2

        for tipe_label, sections, grand_total, hdr_color, total_label in [
            ('ASET',       aset, total_aset, '1D4ED8', 'TOTAL ASET'),
            ('LIABILITAS', liab, total_liab, 'DC2626', 'TOTAL LIABILITAS'),
        ]:
            c = ws3.cell(row=r, column=1, value=tipe_label)
            c.font = Font(bold=True, size=11, color='FFFFFF')
            c.fill = PatternFill(patternType='solid', fgColor=hdr_color)
            ws3.cell(row=r, column=2).fill = PatternFill(patternType='solid', fgColor=hdr_color)
            ws3.cell(row=r, column=3).fill = PatternFill(patternType='solid', fgColor=hdr_color)
            r += 1
            for subtipe, items in sections.items():
                _subtipe_row(ws3, subtipe, r); r += 1
                sub = 0
                for it in items:
                    _data_row(ws3, it['kode'], it['nama'], it['saldo'], r); r += 1
                    sub += it['saldo']
                _total_row(ws3, f'  Total {subtipe}', sub, r); r += 1
            tot_color = 'BFDBFE' if hdr_color == '1D4ED8' else 'FECACA'
            _total_row(ws3, total_label, grand_total, r, color=tot_color); r += 2

        # Ekuitas
        c = ws3.cell(row=r, column=1, value='EKUITAS')
        c.font = Font(bold=True, size=11, color='FFFFFF')
        c.fill = PatternFill(patternType='solid', fgColor='15803D')
        ws3.cell(row=r, column=2).fill = PatternFill(patternType='solid', fgColor='15803D')
        ws3.cell(row=r, column=3).fill = PatternFill(patternType='solid', fgColor='15803D')
        r += 1
        for it in ekuitas_list:
            _data_row(ws3, it['kode'], it['nama'], it['saldo'], r); r += 1
        _total_row(ws3, 'TOTAL EKUITAS', total_ekuitas, r, color='BBF7D0'); r += 2

        # Cek seimbang
        selisih = total_aset - (total_liab + total_ekuitas)
        _total_row(ws3, 'TOTAL LIABILITAS + EKUITAS', total_liab + total_ekuitas, r, color='BFDBFE'); r += 1
        if abs(selisih) > 1:
            ws3.cell(row=r, column=1, value=f'Selisih: {selisih:,.0f}').font = Font(color='DC2626', bold=True)

    except Exception as e:
        ws3.cell(row=1, column=1, value=f'Error membangun sheet Neraca: {e}').font = Font(color='DC2626', bold=True)

    # ══════════════════════════════════════════════════════════════════════
    # SHEET 4: DAFTAR TRANSAKSI  (Nominal | Piutang | Hutang)
    # ══════════════════════════════════════════════════════════════════════
    ws4 = wb.create_sheet('Daftar Transaksi')
    ws4.column_dimensions['A'].width = 19
    ws4.column_dimensions['B'].width = 13
    ws4.column_dimensions['C'].width = 42
    ws4.column_dimensions['D'].width = 20
    ws4.column_dimensions['E'].width = 18
    ws4.column_dimensions['F'].width = 18
    ws4.column_dimensions['G'].width = 18

    _hdr1(ws4, 'DAFTAR TRANSAKSI', 1)
    ws4.cell(row=2, column=1,
             value=f'Periode: {sd}  s/d  {ed}').font = Font(italic=True, size=9, color='64748B')

    hdr_row = 4
    for col4, txt4 in enumerate(
            ['ID Transaksi', 'Tanggal', 'Keterangan', 'Jenis', 'Nominal (Rp)', 'Piutang (Rp)', 'Hutang (Rp)'], 1):
        ch = ws4.cell(row=hdr_row, column=col4, value=txt4)
        ch.font = Font(bold=True, color='FFFFFF', size=10)
        ch.fill = PatternFill(patternType='solid', fgColor='334155')
        ch.alignment = Alignment(horizontal='center')

    try:
        conn2 = db()

        # Pre-fetch piutang delta per jurnal (account 1120)
        piutang_map = {}
        for row_ in conn2.execute("""
            SELECT dj.jurnal_id, SUM(dj.debit) - SUM(dj.kredit) as delta
            FROM detail_jurnal dj
            JOIN akun a ON a.id = dj.akun_id
            WHERE a.kode = '1120'
            GROUP BY dj.jurnal_id
        """).fetchall():
            piutang_map[row_['jurnal_id']] = row_['delta'] or 0

        # Pre-fetch hutang delta per jurnal (account 2100)
        hutang_map = {}
        for row_ in conn2.execute("""
            SELECT dj.jurnal_id, SUM(dj.kredit) - SUM(dj.debit) as delta
            FROM detail_jurnal dj
            JOIN akun a ON a.id = dj.akun_id
            WHERE a.kode = '2100'
            GROUP BY dj.jurnal_id
        """).fetchall():
            hutang_map[row_['jurnal_id']] = row_['delta'] or 0

        # Main transaction rows
        tx_rows = conn2.execute("""
            SELECT j.id, j.nomor_tx, j.tanggal, j.keterangan, j.tipe_tx, j.kategori,
                   COALESCE((
                       SELECT SUM(dj2.kredit)
                       FROM detail_jurnal dj2
                       JOIN akun a2 ON a2.id = dj2.akun_id
                       WHERE dj2.jurnal_id = j.id AND a2.kode LIKE '4%'
                   ), 0) as rev_kredit,
                   COALESCE((
                       SELECT SUM(d2.debit)
                       FROM detail_jurnal d2
                       WHERE d2.jurnal_id = j.id
                   ), 0) as total_debit
            FROM jurnal j
            WHERE j.tanggal >= ? AND j.tanggal <= ?
            ORDER BY j.tanggal, j.id
        """, (sd, ed)).fetchall()
        # Jurnal yang menyentuh 1160 Piutang Karyawan (label 'Pinjaman Karyawan' di Sheet Jenis)
        pinj_ids4 = {r[0] for r in conn2.execute("""
            SELECT DISTINCT dj.jurnal_id FROM detail_jurnal dj
            JOIN akun a ON a.id = dj.akun_id
            JOIN jurnal j ON j.id = dj.jurnal_id
            WHERE a.kode='1160' AND j.tanggal >= ? AND j.tanggal <= ?
        """, (sd, ed)).fetchall()}
        conn2.close()

        r4 = hdr_row + 1
        total_nominal4 = total_piutang4 = total_hutang4 = 0

        for tx in tx_rows:
            jid  = tx['id']
            pd4  = piutang_map.get(jid, 0)
            hd4  = hutang_map.get(jid, 0)
            is_masuk = tx['tipe_tx'] == 'PEMASUKAN'

            # Jenis label
            if pd4 > 0.005:
                jenis4 = 'Penjualan Kredit'
            elif pd4 < -0.005:
                jenis4 = 'Penerimaan Piutang'
            elif hd4 > 0.005:
                jenis4 = 'Pembelian Kredit'
            elif hd4 < -0.005:
                jenis4 = 'Pembayaran Hutang'
            elif jid in pinj_ids4:
                jenis4 = ('Pengembalian Pinjaman Karyawan'
                          if tx['tipe_tx'] == 'PELUNASAN' else 'Pinjaman Karyawan')
            elif is_masuk:
                jenis4 = 'Penjualan'
            elif tx['kategori'] == 'INVESTASI':
                jenis4 = 'Investasi Aset'
            elif tx['kategori'] == 'PENDANAAN':
                jenis4 = 'Penarikan Owner'
            else:
                jenis4 = 'Operasional'

            # Column values (mutually exclusive)
            if abs(pd4) > 0.005:
                nominal_v, piutang_v, hutang_v = None, pd4, None
                total_piutang4 += pd4
            elif abs(hd4) > 0.005:
                nominal_v, piutang_v, hutang_v = None, None, hd4
                total_hutang4 += hd4
            else:
                # Pengembalian pinjaman karyawan (Dr kas / Cr 1160, tipe PELUNASAN) = kas MASUK,
                # jadi bertanda positif seperti penjualan (bukan outflow negatif).
                is_repay_pinjaman = (jid in pinj_ids4 and tx['tipe_tx'] == 'PELUNASAN')
                raw4 = (tx['rev_kredit'] if is_masuk else tx['total_debit']) or 0
                nominal_v = raw4 if (is_masuk or is_repay_pinjaman) else -raw4
                piutang_v, hutang_v = None, None
                total_nominal4 += nominal_v

            ws4.cell(row=r4, column=1, value=tx['nomor_tx'])
            ws4.cell(row=r4, column=2, value=str(tx['tanggal']))
            ws4.cell(row=r4, column=3, value=tx['keterangan'])
            ws4.cell(row=r4, column=4, value=jenis4)

            for col4v, val4 in [(5, nominal_v), (6, piutang_v), (7, hutang_v)]:
                if val4 is not None:
                    cv = ws4.cell(row=r4, column=col4v, value=val4)
                    cv.number_format = '#,##0'
                    cv.font = Font(color='16A34A' if val4 >= 0 else 'DC2626', bold=True)
            r4 += 1

        # Footer total row
        r4 += 1
        fill_foot = PatternFill(patternType='solid', fgColor='F1F5F9')
        ct = ws4.cell(row=r4, column=4, value='TOTAL')
        ct.font = Font(bold=True, size=10)
        ct.fill = fill_foot
        for col4v, val4 in [(5, total_nominal4), (6, total_piutang4), (7, total_hutang4)]:
            cv = ws4.cell(row=r4, column=col4v, value=val4)
            cv.number_format = '#,##0'
            cv.font = Font(bold=True, color='16A34A' if val4 >= 0 else 'DC2626')
            cv.fill = fill_foot

    except Exception as e4:
        ws4.cell(row=5, column=1,
                 value=f'Error Sheet 4: {e4}').font = Font(color='DC2626', bold=True)

    # ── Output ──────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    fname = f"laporan_keuangan_{sd}_{ed}.xlsx"
    return send_file(buf, download_name=fname, as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ---------- CLEAR DATA ----------
@app.route('/admin/clear-data', methods=['POST'])
@admin_required
def clear_data():
    konfirmasi = request.form.get('konfirmasi', '').strip()
    password   = request.form.get('password', '').strip()

    if konfirmasi != 'HAPUS SEMUA DATA':
        flash('Teks konfirmasi tidak sesuai. Data tidak dihapus.', 'danger')
        return redirect(url_for('settings'))

    conn = db()
    user = conn.execute("SELECT password_hash FROM users WHERE id=?", (session['user_id'],)).fetchone()
    conn.close()
    if not user or sha256(password.encode()).hexdigest() != user['password_hash']:
        flash('Password salah. Data tidak dihapus.', 'danger')
        return redirect(url_for('settings'))

    conn = db()
    # ── Urutan hapus: ANAK dulu, baru INDUK ──────────────────────────────────
    # Dijalankan satu per satu di DALAM satu transaksi (BUKAN executescript).
    # executescript melakukan COMMIT implisit lalu menjalankan tiap pernyataan
    # berurutan, jadi kalau satu pernyataan gagal di tengah, pernyataan
    # sebelumnya SUDAH tersimpan permanen. Itu yang dulu terjadi: penghapusan
    # berhenti di `produk` karena produk_harga menahannya lewat foreign key,
    # sementara jurnal, piutang, invoice, dan lainnya sudah telanjur hilang.
    # Database jadi separuh terhapus dan tidak bisa dikembalikan.
    #
    # Dua tabel di bawah ini dulu TIDAK ikut dihapus dan itu sumber masalahnya:
    #   produk_harga          : punya FK ke produk tanpa cascade -> memblokir
    #                           DELETE FROM produk (inilah penyebab error 500)
    #   penjualan_resep_bahan : tidak punya FK sama sekali -> barisnya tertinggal
    #                           sebagai yatim, dan karena jurnal_id-nya angka
    #                           biasa, baris lama bisa "menempel" ke jurnal baru
    #                           yang kebetulan ber-ID sama
    #
    # Tabel ber-FK CASCADE ke jurnal (transaksi_item, writeoff_piutang,
    # penyesuaian_tagihan, detail_jurnal, pos_sale, pergerakan_persediaan_non_sku)
    # ikut terhapus sendiri saat jurnal dihapus, jadi tidak perlu disebut.
    TABEL_HAPUS = [
        'penjualan_resep_bahan',           # anak jurnal tanpa FK, wajib manual
        'detail_jurnal', 'penyesuaian_tagihan', 'pergerakan_persediaan_non_sku',
        'jurnal',
        'bayar_piutang', 'bayar_hutang', 'piutang', 'hutang',
        'invoice_item', 'invoice',
        'po_item', 'purchase_order',
        'pos_sale', 'pos_shift',
        'payroll', 'karyawan',
        'produksi_non_sku', 'produksi_bahan', 'produksi',
        'pergerakan_stok', 'hpp_produk',
        'produk_harga',                    # anak produk, HARUS sebelum produk
        'produk',
        'aset_tetap', 'log_aktivitas',
    ]
    # sqlite_sequence SENGAJA TIDAK dihapus. Menghapusnya mereset AUTOINCREMENT
    # sehingga ID jurnal mulai dari 1 lagi dan bisa bertabrakan dengan baris anak
    # yang (karena bug apa pun di masa depan) sempat tertinggal. Nomor transaksi
    # yang dilihat user adalah `nomor_tx`, bukan ID internal ini, jadi tidak ada
    # bedanya di layar. Membiarkannya = ID tidak pernah dipakai ulang selamanya.
    reset_keys = ['nama_usaha','inv_nama','inv_tagline','inv_alamat','inv_email',
                  'inv_logo','inv_telepon','inv_catatan','inv_terms','inv_top_note',
                  'inv_rek','modal_awal']
    try:
        conn.execute('BEGIN')
        for t in TABEL_HAPUS:
            conn.execute(f'DELETE FROM {t}')
        for k in reset_keys:
            conn.execute("UPDATE settings SET value='' WHERE key=?", (k,))
        conn.execute("UPDATE settings SET value='[]' WHERE key='inv_rek'")
        conn.commit()
    except Exception as ex:
        # Gagal di tengah = BATAL SEMUANYA. Lebih baik user mengulang dengan data
        # utuh daripada meninggalkan database separuh terhapus yang tak bisa dipulihkan.
        conn.rollback()
        conn.close()
        app.logger.error('Bersihkan Data dibatalkan: %s', ex)
        flash('Pembersihan data dibatalkan karena ada kesalahan, dan TIDAK ADA data yang '
              f'terhapus. Silakan coba lagi atau hubungi admin. Rincian: {ex}', 'danger')
        return redirect(url_for('settings'))
    conn.close()
    flash('Semua data berhasil dihapus. Aplikasi siap digunakan dari awal.', 'success')
    return redirect(url_for('dashboard'))


# ─────────────────────────────────────────────────────────────────────────────
# Auto-init DB pada module load, penting untuk deploy WSGI (PythonAnywhere,
# gunicorn, dll) yang load app via `from app import app`, bukan via __main__.
# init_db() idempotent (pakai CREATE TABLE IF NOT EXISTS + INSERT OR IGNORE),
# jadi aman dipanggil setiap proses start. Tanpa ini, first request akan kena
# "sqlite3.OperationalError: no such table: settings" → 500 Internal Server Error.
# ─────────────────────────────────────────────────────────────────────────────
try:
    init_db()
    _tzconn = db()
    apply_tz_app(get_setting(_tzconn, 'tz_app', TZ_APP_DEFAULT))
    _tzconn.close()
except Exception as _e:
    # Jangan crash module load, biar Flask tetap up & error muncul di log.
    import traceback as _tb
    print('[WARN] init_db() gagal pada startup:', _e)
    _tb.print_exc()

if __name__ == '__main__':
    print('\n' + '='*55)
    print('   FINANSIAL APP v5.2 - Sistem Akuntansi Lokal')
    print('   Buka browser: http://localhost:5000')
    print('   Login: admin / admin123')
    print('='*55 + '\n')
    app.run(debug=False, host='127.0.0.1', port=5000)
