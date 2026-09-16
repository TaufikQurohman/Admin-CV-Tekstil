function setupDatabase() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const database = {
    "01_INVOICE": [
      "ID Invoice",
      "No Invoice",
      "Tanggal",
      "Nama Pembeli",
      "Alamat Pembeli",
      "No PO",
      "Jenis Kain",
      "Jenis Pembayaran",
      "DP",
      "Sisa Tagihan",
      "Keterangan",
      "Total Invoice",
      "Status"
    ],

    "02_DETAIL_INVOICE": [
      "ID Detail",
      "ID Invoice",
      "No",
      "Nama Kain",
      "Warna",
      "No Roll",
      "Berat",
      "Satuan",
      "Harga Satuan",
      "Jumlah"
    ],

    "03_HPP": [
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
    ],

    "04_PENGELUARAN": [
      "ID Pengeluaran",
      "Tanggal",
      "Nama",
      "Kebutuhan",
      "Jumlah",
      "Jenis Pembayaran",
      "Kategori",
      "Keterangan"
    ],

    "05_SURAT_JALAN": [
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
    ],

    "06_DETAIL_SURAT_JALAN": [
      "ID Detail",
      "ID Surat Jalan",
      "No",
      "Nama Barang",
      "Jenis Kain",
      "Warna",
      "No Roll",
      "Berat Kg",
      "Keterangan"
    ],

    "07_MASTER_PABRIK": [
      "ID Pabrik",
      "Nama Pabrik",
      "Alamat",
      "Keterangan"
    ],

    "08_MASTER_KATEGORI": [
      "ID Kategori",
      "Nama Kategori",
      "Keterangan"
    ]
  };

  for (const sheetName in database) {

    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    const headers = database[sheetName];

    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers]);

    sheet
      .getRange(1, 1, 1, headers.length)
      .setFontWeight("bold");

    sheet.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert("Database V1 berhasil dibuat!");
}