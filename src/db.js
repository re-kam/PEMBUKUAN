// In-memory data store for PEMBUKUAN (Finance by Cliff)
// Mocked in-memory store according to AI Studio migration guidelines

export const store = {
  settings: {
    nama_usaha: 'Finance by Cliff',
    pref_decimal: false,
    pref_custom_coa: false,
    theme: 'light'
  },

  users: [
    { id: 1, username: 'admin', password: 'admin123', nama: 'Cliff (Owner)', role: 'ADMIN' },
    { id: 2, username: 'operator', password: 'operator123', nama: 'Siti Operator', role: 'OPERATOR' },
    { id: 3, username: 'kasir', password: 'kasir123', nama: 'Budi Kasir', role: 'KASIR' }
  ],

  accounts: [
    { id: 1, kode: '101', nama: 'Kas Utama', kategori: 'ASET', tipe: 'DEBIT', saldo: 25450000 },
    { id: 2, kode: '102', nama: 'Bank BCA Operasional', kategori: 'ASET', tipe: 'DEBIT', saldo: 142800000 },
    { id: 3, kode: '103', nama: 'Bank Mandiri Cadangan', kategori: 'ASET', tipe: 'DEBIT', saldo: 68250000 },
    { id: 4, kode: '104', nama: 'Piutang Usaha', kategori: 'ASET', tipe: 'DEBIT', saldo: 34600000 },
    { id: 5, kode: '105', nama: 'Persediaan Barang Dagang', kategori: 'ASET', tipe: 'DEBIT', saldo: 45000000 },
    { id: 6, kode: '121', nama: 'Peralatan & Mesin Usaha', kategori: 'ASET_TETAP', tipe: 'DEBIT', saldo: 35000000 },
    { id: 7, kode: '122', nama: 'Akumulasi Penyusutan Peralatan', kategori: 'ASET_TETAP', tipe: 'KREDIT', saldo: -5000000 },
    { id: 8, kode: '201', nama: 'Hutang Usaha (Supplier)', kategori: 'KEWAJIBAN', tipe: 'KREDIT', saldo: 28500000 },
    { id: 9, kode: '202', nama: 'Hutang Pajak', kategori: 'KEWAJIBAN', tipe: 'KREDIT', saldo: 2150000 },
    { id: 10, kode: '301', nama: 'Modal Pemilik', kategori: 'EKUITAS', tipe: 'KREDIT', saldo: 250000000 },
    { id: 11, kode: '302', nama: 'Laba Ditahan', kategori: 'EKUITAS', tipe: 'KREDIT', saldo: 32500000 },
    { id: 12, kode: '401', nama: 'Pendapatan Penjualan Produk', kategori: 'PENDAPATAN', tipe: 'KREDIT', saldo: 95400000 },
    { id: 13, kode: '402', nama: 'Pendapatan Jasa & Layanan', kategori: 'PENDAPATAN', tipe: 'KREDIT', saldo: 18000000 },
    { id: 14, kode: '501', nama: 'Harga Pokok Penjualan (HPP)', kategori: 'HPP', tipe: 'DEBIT', saldo: 42100000 },
    { id: 15, kode: '601', nama: 'Beban Gaji & Upah', kategori: 'BEBAN', tipe: 'DEBIT', saldo: 15500000 },
    { id: 16, kode: '602', nama: 'Beban Sewa Gedung', kategori: 'BEBAN', tipe: 'DEBIT', saldo: 6000000 },
    { id: 17, kode: '603', nama: 'Beban Listrik, Air, & WiFi', kategori: 'BEBAN', tipe: 'DEBIT', saldo: 2450000 },
    { id: 18, kode: '604', nama: 'Beban Promosi & Iklan', kategori: 'BEBAN', tipe: 'DEBIT', saldo: 4200000 },
    { id: 19, kode: '605', nama: 'Beban Perlengkapan & Umum', kategori: 'BEBAN', tipe: 'DEBIT', saldo: 1800000 }
  ],

  transactions: [
    {
      id: 1,
      tanggal: '2026-10-08',
      tipe: 'PEMASUKAN',
      referensi: 'TX-20261008-001',
      kategori: 'Penjualan Retail',
      keterangan: 'Penjualan paket kopi dan pastry kafe batch 1',
      kontak: 'Pelanggan Umum',
      rekening: 'Kas Utama',
      nominal: 3450000,
      akun_debit: 'Kas Utama',
      akun_kredit: 'Pendapatan Penjualan Produk'
    },
    {
      id: 2,
      tanggal: '2026-10-07',
      tipe: 'PEMASUKAN',
      referensi: 'INV-2026-089',
      kategori: 'Penjualan Grosir',
      keterangan: 'Penjualan biji kopi wholesale 50kg ke Resto Harmony',
      kontak: 'Resto Harmony Nusantara',
      rekening: 'Bank BCA Operasional',
      nominal: 8500000,
      akun_debit: 'Bank BCA Operasional',
      akun_kredit: 'Pendapatan Penjualan Produk'
    },
    {
      id: 3,
      tanggal: '2026-10-06',
      tipe: 'PENGELUARAN',
      referensi: 'EXP-20261006-002',
      kategori: 'Bahan Baku & Stok',
      keterangan: 'Pembelian green beans Java Arabica 100kg',
      kontak: 'Koperasi Tani Java',
      rekening: 'Bank Mandiri Cadangan',
      nominal: 9200000,
      akun_debit: 'Harga Pokok Penjualan (HPP)',
      akun_kredit: 'Bank Mandiri Cadangan'
    },
    {
      id: 4,
      tanggal: '2026-10-05',
      tipe: 'PENGELUARAN',
      referensi: 'EXP-20261005-001',
      kategori: 'Utilitas & Operasional',
      keterangan: 'Pembayaran tagihan listrik PLN dan Biznet WiFi',
      kontak: 'PLN & Biznet',
      rekening: 'Bank BCA Operasional',
      nominal: 1850000,
      akun_debit: 'Beban Listrik, Air, & WiFi',
      akun_kredit: 'Bank BCA Operasional'
    },
    {
      id: 5,
      tanggal: '2026-10-04',
      tipe: 'PEMASUKAN',
      referensi: 'TX-20261004-003',
      kategori: 'Jasa Konsultasi',
      keterangan: 'Jasa kurasi racikan kopi & pelatihan barista event hotel',
      kontak: 'Grand Mercure Jakarta',
      rekening: 'Bank BCA Operasional',
      nominal: 6000000,
      akun_debit: 'Bank BCA Operasional',
      akun_kredit: 'Pendapatan Jasa & Layanan'
    },
    {
      id: 6,
      tanggal: '2026-10-03',
      tipe: 'PENGELUARAN',
      referensi: 'EXP-20261003-001',
      kategori: 'Pemasaran & Digital',
      keterangan: 'Kampanye promosi digital Meta Ads dan TikTok Ads',
      kontak: 'Meta Platforms Inc',
      rekening: 'Bank BCA Operasional',
      nominal: 2500000,
      akun_debit: 'Beban Promosi & Iklan',
      akun_kredit: 'Bank BCA Operasional'
    },
    {
      id: 7,
      tanggal: '2026-10-02',
      tipe: 'PEMASUKAN',
      referensi: 'TX-POS-100201',
      kategori: 'Penjualan Kasir POS',
      keterangan: 'Penerimaan harian kasir POS shift pagi & malam',
      kontak: 'Pelanggan Walk-In',
      rekening: 'Kas Utama',
      nominal: 4820000,
      akun_debit: 'Kas Utama',
      akun_kredit: 'Pendapatan Penjualan Produk'
    },
    {
      id: 8,
      tanggal: '2026-10-01',
      tipe: 'PENGELUARAN',
      referensi: 'GAJI-202610',
      kategori: 'Gaji Karyawan',
      keterangan: 'Gaji pokok barista, roaster, dan kasir periode awal Oktober',
      kontak: 'Karyawan Toko (5 Orang)',
      rekening: 'Bank BCA Operasional',
      nominal: 15500000,
      akun_debit: 'Beban Gaji & Upah',
      akun_kredit: 'Bank BCA Operasional'
    }
  ],

  products: [
    { id: 1, kode: 'PRD-001', nama: 'Kopi Arabika Specialty 250g', varian: 'Medium Roast', harga: 85000, stok: 48, satuan: 'Pck', kategori: 'Coffee Beans' },
    { id: 2, kode: 'PRD-002', nama: 'Kopi Robusta Flores 250g', varian: 'Dark Roast', harga: 55000, stok: 65, satuan: 'Pck', kategori: 'Coffee Beans' },
    { id: 3, kode: 'PRD-003', nama: 'Matcha Uji Premium 500g', varian: 'Powder', harga: 120000, stok: 30, satuan: 'Can', kategori: 'Beverage' },
    { id: 4, kode: 'PRD-004', nama: 'Sirup Vanilla Gourmet 750ml', varian: 'Artisan', harga: 95000, stok: 22, satuan: 'Btl', kategori: 'Syrup' },
    { id: 5, kode: 'PRD-005', nama: 'Paper Cup & Lid 12oz (Pack 50)', varian: 'Eco Brown', harga: 45000, stok: 80, satuan: 'Pack', kategori: 'Packaging' },
    { id: 6, kode: 'PRD-006', nama: 'Brownies Fudgy Signature', varian: 'Almond Dark', harga: 35000, stok: 25, satuan: 'Pcs', kategori: 'Pastry' },
    { id: 7, kode: 'PRD-007', nama: 'Croissant Butter Flaky', varian: 'French Butter', harga: 28000, stok: 40, satuan: 'Pcs', kategori: 'Pastry' },
    { id: 8, kode: 'PRD-008', nama: 'Cold Brew Signature 250ml', varian: 'Ready To Drink', harga: 32000, stok: 35, satuan: 'Btl', kategori: 'Beverage' }
  ],

  piutang: [
    { id: 1, kontak: 'PT Berkah Mandiri Sejahtera', jatuh_tempo: '2026-10-25', jumlah: 18500000, terbayar: 0, status: 'BELUM LUNAS', ref: 'INV-2026-074' },
    { id: 2, kontak: 'CV Sinar Abadi Cafe', jatuh_tempo: '2026-10-18', jumlah: 12100000, terbayar: 0, status: 'BELUM LUNAS', ref: 'INV-2026-079' },
    { id: 3, kontak: 'Toko Rezeki Segar', jatuh_tempo: '2026-10-12', jumlah: 4000000, terbayar: 0, status: 'BELUM LUNAS', ref: 'INV-2026-081' }
  ],

  hutang: [
    { id: 1, kontak: 'Java Coffee Farm Syndicate', jatuh_tempo: '2026-10-28', jumlah: 15000000, terbayar: 0, status: 'BELUM LUNAS', ref: 'PO-2026-041' },
    { id: 2, kontak: 'Packaging Jaya Nusantara', jatuh_tempo: '2026-10-20', jumlah: 8500000, terbayar: 0, status: 'BELUM LUNAS', ref: 'PO-2026-044' },
    { id: 3, kontak: 'Distributor Susu Segar Murni', jatuh_tempo: '2026-10-14', jumlah: 5000000, terbayar: 0, status: 'BELUM LUNAS', ref: 'PO-2026-047' }
  ],

  posOrders: []
};

// Formatting helpers
export function formatRp(num) {
  const n = Number(num) || 0;
  const sign = n < 0 ? '-Rp ' : 'Rp ';
  return sign + Math.abs(n).toLocaleString('id-ID');
}

export function formatTgl(dateStr) {
  if (!dateStr) return '-';
  try {
    const d = new Date(dateStr);
    const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
  } catch {
    return dateStr;
  }
}

// Accounting calculations
export function getFinancialSummary() {
  const totalKasBank = store.accounts
    .filter(a => ['Kas Utama', 'Bank BCA Operasional', 'Bank Mandiri Cadangan'].includes(a.nama))
    .reduce((sum, a) => sum + a.saldo, 0);

  const totalPendapatan = store.accounts
    .filter(a => a.kategori === 'PENDAPATAN')
    .reduce((sum, a) => sum + a.saldo, 0);

  const totalHpp = store.accounts
    .filter(a => a.kategori === 'HPP')
    .reduce((sum, a) => sum + a.saldo, 0);

  const totalBeban = store.accounts
    .filter(a => a.kategori === 'BEBAN')
    .reduce((sum, a) => sum + a.saldo, 0);

  const labaKotor = totalPendapatan - totalHpp;
  const labaBersih = labaKotor - totalBeban;

  const totalPiutang = store.piutang.reduce((sum, p) => sum + (p.jumlah - p.terbayar), 0);
  const totalHutang = store.hutang.reduce((sum, h) => sum + (h.jumlah - h.terbayar), 0);

  return {
    totalKasBank,
    totalPendapatan,
    totalHpp,
    totalBeban,
    labaKotor,
    labaBersih,
    totalPiutang,
    totalHutang
  };
}
