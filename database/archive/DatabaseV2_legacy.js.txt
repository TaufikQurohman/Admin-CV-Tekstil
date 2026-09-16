function setupDatabaseV2() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // =====================================================
  // RENAME SHEET LAMA AGAR STRUKTURNYA RAPI
  // DATA TIDAK DIHAPUS
  // =====================================================

  const renameMap = {
    "03_HPP": "05_HPP",
    "04_PENGELUARAN": "06_PENGELUARAN",
    "05_SURAT_JALAN": "07_SURAT_JALAN",
    "06_DETAIL_SURAT_JALAN": "08_DETAIL_SURAT_JALAN",
    "07_MASTER_PABRIK": "09_MASTER_PABRIK",
    "08_MASTER_KATEGORI": "10_MASTER_KATEGORI"
  };

  for (const oldName in renameMap) {

    const newName = renameMap[oldName];

    const oldSheet = ss.getSheetByName(oldName);
    const newSheet = ss.getSheetByName(newName);

    // Kalau sheet lama ada dan nama baru belum ada,
    // rename tanpa menghapus isi.
    if (oldSheet && !newSheet) {
      oldSheet.setName(newName);
    }
  }


  // =====================================================
  // PEMBUKUAN KAIN GREY
  // =====================================================

  createSheetIfMissing_(
    ss,
    "03_PEMBUKUAN_GREY",
    [
      "ID Pembukuan",
      "ID Invoice",
      "No Invoice",
      "Tanggal",
      "Nama Pembeli",
      "Jumlah Roll",
      "Jumlah Kg",
      "Jenis Kain",
      "Warna",
      "Harga",
      "Jenis Pembayaran",
      "Bukti Pembayaran",
      "Keterangan",
      "Total Penjualan"
    ]
  );


  // =====================================================
  // PEMBUKUAN KAIN JADI
  // =====================================================

  createSheetIfMissing_(
    ss,
    "04_PEMBUKUAN_KAIN_JADI",
    [
      "ID Pembukuan",
      "ID Invoice",
      "No Invoice",
      "Tanggal",
      "Nama Pembeli",
      "Jumlah Roll",
      "Jumlah Kg",
      "Jenis Kain",
      "Warna",
      "Harga",
      "Jenis Pembayaran",
      "Bukti Pembayaran",
      "Keterangan",
      "Total Penjualan"
    ]
  );


  // =====================================================
  // PASTIKAN SHEET LAIN ADA
  // =====================================================

  createSheetIfMissing_(
    ss,
    "05_HPP",
    [
      "ID HPP",
      "Tanggal",
      "Supplier",
      "Nama Barang",
      "Warna",
      "Jumlah",
      "Harga Satuan",
      "Pajak",
      "Total",
      "Jenis Transaksi",
      "Keterangan"
    ]
  );


  createSheetIfMissing_(
    ss,
    "06_PENGELUARAN",
    [
      "ID Pengeluaran",
      "Tanggal",
      "Nama",
      "Kebutuhan",
      "Jumlah",
      "Jenis Pembayaran",
      "Kategori",
      "Keterangan"
    ]
  );


  createSheetIfMissing_(
    ss,
    "07_SURAT_JALAN",
    [
      "ID Surat Jalan",
      "No Surat Jalan",
      "Tanggal",
      "Pabrik Tujuan",
      "Alamat",
      "No PO",
      "Kendaraan",
      "Nama Supir",
      "Tujuan",
      "Jenis Pengiriman",
      "Keterangan"
    ]
  );


  createSheetIfMissing_(
    ss,
    "08_DETAIL_SURAT_JALAN",
    [
      "ID Detail",
      "ID Surat Jalan",
      "No",
      "Nama Barang",
      "Jenis Kain",
      "Warna",
      "No Roll",
      "Berat Kg",
      "Keterangan"
    ]
  );


  createSheetIfMissing_(
    ss,
    "09_MASTER_PABRIK",
    [
      "ID Pabrik",
      "Nama Pabrik",
      "Alamat",
      "Keterangan"
    ]
  );


  createSheetIfMissing_(
    ss,
    "10_MASTER_KATEGORI",
    [
      "ID Kategori",
      "Nama Kategori",
      "Keterangan"
    ]
  );


  SpreadsheetApp.getUi().alert(
    "Database V2 berhasil dibuat!"
  );
}


/* =========================================================
   FUNGSI MEMBUAT SHEET JIKA BELUM ADA
   ========================================================= */

function createSheetIfMissing_(ss, sheetName, headers) {

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {

    sheet = ss.insertSheet(sheetName);

    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers]);

    sheet
      .getRange(1, 1, 1, headers.length)
      .setFontWeight("bold");

    sheet.setFrozenRows(1);
  }
}