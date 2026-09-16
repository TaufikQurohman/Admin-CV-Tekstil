/************************************************************
 * DATABASE V3 FINAL
 * CV PANDS TEKSTIL
 *
 * DATABASE:
 *
 * 01_INVOICE
 * 02_DETAIL_INVOICE
 * 03_PEMBUKUAN_GREY
 * 04_PEMBUKUAN_KAIN_JADI
 * 05_HPP
 * 06_PENGELUARAN
 * 07_SURAT_JALAN
 * 08_DETAIL_SURAT_JALAN
 * 09_MASTER_PABRIK
 * 10_LOG_AKTIVITAS
 *
 * FITUR YANG DIPERSIAPKAN:
 * - Pembukuan Kg / Meter
 * - Surat Jalan tanpa Tujuan
 * - Detail Surat Jalan tanpa Nama Barang
 * - Log aktivitas
 * - Status Invoice untuk Lunas / DP / Belum Lunas / Dibatalkan
 ************************************************************/


/* =========================================================
   KONFIGURASI
========================================================= */

const DB_CONFIG = {

  SHEETS: {

    INVOICE:
      "01_INVOICE",

    DETAIL_INVOICE:
      "02_DETAIL_INVOICE",

    PEMBUKUAN_GREY:
      "03_PEMBUKUAN_GREY",

    PEMBUKUAN_KAIN_JADI:
      "04_PEMBUKUAN_KAIN_JADI",

    HPP:
      "05_HPP",

    PENGELUARAN:
      "06_PENGELUARAN",

    SURAT_JALAN:
      "07_SURAT_JALAN",

    DETAIL_SURAT_JALAN:
      "08_DETAIL_SURAT_JALAN",

    MASTER_PABRIK:
      "09_MASTER_PABRIK",

    LOG_AKTIVITAS:
      "10_LOG_AKTIVITAS"

  }

};


/* =========================================================
   HEADER DATABASE
========================================================= */

const DB_HEADERS = {

  /* -------------------------------------------------------
     01 INVOICE
  ------------------------------------------------------- */

  INVOICE: [

    "ID Invoice",
    "No Invoice",
    "Tanggal",
    "Nama Pembeli",
    "Alamat Pembeli",
    "No PO",
    "Kategori Penjualan",
    "Jenis Pembayaran",
    "DP",
    "Sisa Tagihan",
    "Keterangan",
    "Total Invoice",
    "Status"

  ],


  /* -------------------------------------------------------
     02 DETAIL INVOICE
  ------------------------------------------------------- */

  DETAIL_INVOICE: [

    "ID Invoice",
    "No",
    "Jenis Kain",
    "Warna",
    "No Roll",
    "Qty",
    "Satuan",
    "Harga Satuan",
    "Jumlah"

  ],


  /* -------------------------------------------------------
     03 & 04 PEMBUKUAN
  ------------------------------------------------------- */

  PEMBUKUAN: [

    "ID Pembukuan",
    "ID Invoice",
    "No Invoice",
    "Tanggal",
    "Nama Pembeli",
    "Jumlah",
    "Satuan",
    "Berat Kg",
    "Jenis Kain",
    "Warna",
    "Harga Satuan",
    "Jenis Pembayaran",
    "Bukti Pembayaran",
    "Keterangan",
    "Total Penjualan"

  ],


  /* -------------------------------------------------------
     05 HPP
  ------------------------------------------------------- */

  HPP: [

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


  /* -------------------------------------------------------
     06 PENGELUARAN
  ------------------------------------------------------- */

  PENGELUARAN: [

    "ID Pengeluaran",
    "Tanggal",
    "Nama",
    "Kebutuhan",
    "Jumlah",
    "Jenis Pembayaran",
    "Kategori",
    "Keterangan"

  ],


  /* -------------------------------------------------------
     07 SURAT JALAN
     TUJUAN SUDAH DIHAPUS
  ------------------------------------------------------- */

  SURAT_JALAN: [

    "ID Surat Jalan",
    "No Surat Jalan",
    "Tanggal",
    "ID Pabrik",
    "Nama Pabrik",
    "Alamat Pabrik",
    "No PO",
    "Kendaraan",
    "Nama Supir",
    "Jenis Pengiriman",
    "Keterangan",
    "Total Roll",
    "Total Kg"

  ],


  /* -------------------------------------------------------
     08 DETAIL SURAT JALAN
     NAMA BARANG SUDAH DIHAPUS
  ------------------------------------------------------- */

  DETAIL_SURAT_JALAN: [

    "ID Surat Jalan",
    "No",
    "Jenis Kain",
    "Warna",
    "No Roll",
    "Berat Kg",
    "Keterangan"

  ],


  /* -------------------------------------------------------
     09 MASTER PABRIK
  ------------------------------------------------------- */

  MASTER_PABRIK: [

    "ID Pabrik",
    "Nama Pabrik",
    "Alamat",
    "Kontak",
    "Keterangan"

  ],


  /* -------------------------------------------------------
     10 LOG AKTIVITAS
  ------------------------------------------------------- */

  LOG_AKTIVITAS: [

    "ID Log",
    "Tanggal & Waktu",
    "User",
    "Aktivitas",
    "Modul",
    "ID Data",
    "Keterangan"

  ]

};


/* =========================================================
   MENU DATABASE
========================================================= */

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu(
      "Database CV Pands Tekstil"
    )

    .addItem(
      "Setup / Perbaiki Database V3",
      "setupDatabaseV3"
    )

    .addItem(
      "Migrasi Pembukuan",
      "migratePembukuanV3"
    )

    .addItem(
      "Migrasi Surat Jalan",
      "migrateSuratJalanV3"
    )

    .addItem(
      "Cek Database",
      "cekDatabaseV3"
    )

    .addItem(
      "Cek Struktur Pembukuan",
      "cekStrukturPembukuanV3"
    )

    .addItem(
      "Cek Struktur Surat Jalan",
      "cekStrukturSuratJalanV3"
    )

    .addToUi();

}


/* =========================================================
   SETUP DATABASE V3
========================================================= */

function setupDatabaseV3() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const definitions = [

    {
      name:
        DB_CONFIG.SHEETS.INVOICE,

      headers:
        DB_HEADERS.INVOICE
    },


    {
      name:
        DB_CONFIG.SHEETS.DETAIL_INVOICE,

      headers:
        DB_HEADERS.DETAIL_INVOICE
    },


    {
      name:
        DB_CONFIG.SHEETS.PEMBUKUAN_GREY,

      headers:
        DB_HEADERS.PEMBUKUAN
    },


    {
      name:
        DB_CONFIG.SHEETS.PEMBUKUAN_KAIN_JADI,

      headers:
        DB_HEADERS.PEMBUKUAN
    },


    {
      name:
        DB_CONFIG.SHEETS.HPP,

      headers:
        DB_HEADERS.HPP
    },


    {
      name:
        DB_CONFIG.SHEETS.PENGELUARAN,

      headers:
        DB_HEADERS.PENGELUARAN
    },


    {
      name:
        DB_CONFIG.SHEETS.SURAT_JALAN,

      headers:
        DB_HEADERS.SURAT_JALAN
    },


    {
      name:
        DB_CONFIG.SHEETS.DETAIL_SURAT_JALAN,

      headers:
        DB_HEADERS.DETAIL_SURAT_JALAN
    },


    {
      name:
        DB_CONFIG.SHEETS.MASTER_PABRIK,

      headers:
        DB_HEADERS.MASTER_PABRIK
    },


    {
      name:
        DB_CONFIG.SHEETS.LOG_AKTIVITAS,

      headers:
        DB_HEADERS.LOG_AKTIVITAS
    }

  ];


  definitions.forEach(
    function(definition) {

      let sheet =
        ss.getSheetByName(
          definition.name
        );


      /*
       * Jika belum ada,
       * buat sheet baru.
       */

      if (!sheet) {

        sheet =
          ss.insertSheet(
            definition.name
          );

      }


      /*
       * PEMBUKUAN
       */

      if (

        definition.name ===
          DB_CONFIG.SHEETS.PEMBUKUAN_GREY

        ||

        definition.name ===
          DB_CONFIG.SHEETS.PEMBUKUAN_KAIN_JADI

      ) {

        migratePembukuanSheet_(
          sheet
        );

      }


      /*
       * SURAT JALAN
       */

      else if (

        definition.name ===
          DB_CONFIG.SHEETS.SURAT_JALAN

      ) {

        migrateSuratJalanHeader_(
          sheet
        );

      }


      /*
       * DETAIL SURAT JALAN
       */

      else if (

        definition.name ===
          DB_CONFIG.SHEETS.DETAIL_SURAT_JALAN

      ) {

        migrateDetailSuratJalanSheet_(
          sheet
        );

      }


      /*
       * SHEET LAIN
       */

      else {

        ensureHeader_(
          sheet,
          definition.headers
        );

      }


      formatDatabaseSheet_(
        sheet,
        definition.headers
      );

    }
  );


  /*
   * Master Kategori sudah tidak digunakan.
   *
   * Kalau masih ada, arsipkan.
   */

  const oldMasterKategori =
    ss.getSheetByName(
      "10_MASTER_KATEGORI"
    );


  if (oldMasterKategori) {

    const archiveName =
      "ARSIP_MASTER_KATEGORI";


    const archiveSheet =
      ss.getSheetByName(
        archiveName
      );


    if (!archiveSheet) {

      oldMasterKategori.setName(
        archiveName
      );

    }

  }


  return {

    success:
      true,

    message:
      "Database V3 berhasil disiapkan dan diperbarui."

  };

}


/* =========================================================
   MIGRASI PEMBUKUAN V3
========================================================= */

function migratePembukuanV3() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheets = [

    DB_CONFIG.SHEETS.PEMBUKUAN_GREY,

    DB_CONFIG.SHEETS.PEMBUKUAN_KAIN_JADI

  ];


  sheets.forEach(
    function(sheetName) {

      const sheet =
        ss.getSheetByName(
          sheetName
        );


      if (!sheet) {

        return;

      }


      migratePembukuanSheet_(
        sheet
      );


      formatDatabaseSheet_(
        sheet,
        DB_HEADERS.PEMBUKUAN
      );

    }
  );


  return {

    success:
      true,

    message:
      "Pembukuan Grey dan Kain Jadi berhasil diperbarui."

  };

}


/* =========================================================
   MIGRASI SHEET PEMBUKUAN
========================================================= */

function migratePembukuanSheet_(
  sheet
) {

  if (!sheet) {

    return;

  }


  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  /*
   * Sheet kosong.
   */

  if (

    lastRow === 0

    ||

    lastColumn === 0

  ) {

    ensureHeader_(
      sheet,
      DB_HEADERS.PEMBUKUAN
    );

    return;

  }


  const currentHeaders =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0]
      .map(
        function(value) {

          return cleanString_(
            value
          );

        }
      );


  /*
   * Sudah menggunakan struktur baru.
   */

  if (

    arraysEqual_(
      currentHeaders,
      DB_HEADERS.PEMBUKUAN
    )

  ) {

    return;

  }


  const dataCount =
    Math.max(
      0,
      lastRow - 1
    );


  let oldData = [];


  if (dataCount > 0) {

    oldData =
      sheet
        .getRange(
          2,
          1,
          dataCount,
          lastColumn
        )
        .getValues();

  }


  const indexMap =
    createHeaderIndexMap_(
      currentHeaders
    );


  const newData = [];


  oldData.forEach(
    function(row) {

      /*
       * Abaikan baris kosong.
       */

      if (
        row.every(
          function(value) {

            return isBlank_(
              value
            );

          }
        )
      ) {

        return;

      }


      const idPembukuan =
        getMappedValue_(
          row,
          indexMap,
          [
            "ID Pembukuan"
          ]
        );


      const idInvoice =
        getMappedValue_(
          row,
          indexMap,
          [
            "ID Invoice"
          ]
        );


      const noInvoice =
        getMappedValue_(
          row,
          indexMap,
          [
            "No Invoice"
          ]
        );


      const tanggal =
        getMappedValue_(
          row,
          indexMap,
          [
            "Tanggal"
          ]
        );


      const namaPembeli =
        getMappedValue_(
          row,
          indexMap,
          [
            "Nama Pembeli"
          ]
        );


      let jumlah =
        getMappedValue_(
          row,
          indexMap,
          [
            "Jumlah"
          ]
        );


      let satuan =
        getMappedValue_(
          row,
          indexMap,
          [
            "Satuan"
          ]
        );


      let beratKg =
        getMappedValue_(
          row,
          indexMap,
          [
            "Berat Kg",
            "Jumlah Kg"
          ]
        );


      /*
       * Data lama hanya memiliki
       * Jumlah Kg.
       */

      if (

        isBlank_(jumlah)

        &&

        !isBlank_(beratKg)

      ) {

        jumlah =
          toNumber_(
            beratKg
          );

      }


      /*
       * Default data lama = Kg.
       */

      if (
        isBlank_(satuan)
      ) {

        satuan =
          "Kg";

      }


      /*
       * Jika Kg,
       * Berat Kg = Jumlah.
       */

      if (

        cleanString_(
          satuan
        ).toLowerCase() === "kg"

        &&

        isBlank_(beratKg)

      ) {

        beratKg =
          toNumber_(
            jumlah
          );

      }


      /*
       * Jika Meter dan tidak ada
       * berat Kg, biarkan kosong.
       */

      if (

        cleanString_(
          satuan
        ).toLowerCase() === "meter"

        &&

        isBlank_(beratKg)

      ) {

        beratKg =
          "";

      }


      const jenisKain =
        getMappedValue_(
          row,
          indexMap,
          [
            "Jenis Kain"
          ]
        );


      const warna =
        getMappedValue_(
          row,
          indexMap,
          [
            "Warna"
          ]
        );


      let hargaSatuan =
        getMappedValue_(
          row,
          indexMap,
          [
            "Harga Satuan",
            "Harga"
          ]
        );


      hargaSatuan =
        toNumber_(
          hargaSatuan
        );


      const jenisPembayaran =
        getMappedValue_(
          row,
          indexMap,
          [
            "Jenis Pembayaran"
          ]
        );


      const buktiPembayaran =
        getMappedValue_(
          row,
          indexMap,
          [
            "Bukti Pembayaran"
          ]
        );


      const keterangan =
        getMappedValue_(
          row,
          indexMap,
          [
            "Keterangan"
          ]
        );


      let totalPenjualan =
        getMappedValue_(
          row,
          indexMap,
          [
            "Total Penjualan"
          ]
        );


      if (
        isBlank_(
          totalPenjualan
        )
      ) {

        totalPenjualan =
          toNumber_(
            jumlah
          )
          *
          toNumber_(
            hargaSatuan
          );

      }


      totalPenjualan =
        toNumber_(
          totalPenjualan
        );


      newData.push([

        idPembukuan,

        idInvoice,

        noInvoice,

        tanggal,

        namaPembeli,

        toNumber_(
          jumlah
        ),

        satuan,

        isBlank_(
          beratKg
        )
          ? ""
          : toNumber_(
              beratKg
            ),

        jenisKain,

        warna,

        hargaSatuan,

        jenisPembayaran,

        buktiPembayaran,

        keterangan,

        totalPenjualan

      ]);

    }
  );


  /*
   * Bersihkan isi.
   */

  sheet.clearContents();


  /*
   * Header baru.
   */

  sheet
    .getRange(
      1,
      1,
      1,
      DB_HEADERS.PEMBUKUAN.length
    )
    .setValues([
      DB_HEADERS.PEMBUKUAN
    ]);


  /*
   * Tulis data hasil migrasi.
   */

  if (
    newData.length > 0
  ) {

    sheet
      .getRange(
        2,
        1,
        newData.length,
        DB_HEADERS.PEMBUKUAN.length
      )
      .setValues(
        newData
      );

  }

}


/* =========================================================
   MIGRASI SURAT JALAN V3
========================================================= */

function migrateSuratJalanV3() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const suratJalanSheet =
    ss.getSheetByName(
      DB_CONFIG.SHEETS.SURAT_JALAN
    );


  const detailSheet =
    ss.getSheetByName(
      DB_CONFIG.SHEETS.DETAIL_SURAT_JALAN
    );


  if (suratJalanSheet) {

    migrateSuratJalanHeader_(
      suratJalanSheet
    );


    formatDatabaseSheet_(
      suratJalanSheet,
      DB_HEADERS.SURAT_JALAN
    );

  }


  if (detailSheet) {

    migrateDetailSuratJalanSheet_(
      detailSheet
    );


    formatDatabaseSheet_(
      detailSheet,
      DB_HEADERS.DETAIL_SURAT_JALAN
    );

  }


  return {

    success:
      true,

    message:
      "Surat Jalan berhasil diperbarui."

  };

}


/* =========================================================
   MIGRASI HEADER SURAT JALAN
========================================================= */

function migrateSuratJalanHeader_(
  sheet
) {

  if (!sheet) {

    return;

  }


  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  /*
   * Kosong.
   */

  if (

    lastRow === 0

    ||

    lastColumn === 0

  ) {

    ensureHeader_(
      sheet,
      DB_HEADERS.SURAT_JALAN
    );

    return;

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0]
      .map(
        function(value) {

          return cleanString_(
            value
          );

        }
      );


  /*
   * Sudah format baru.
   */

  if (

    arraysEqual_(
      headers,
      DB_HEADERS.SURAT_JALAN
    )

  ) {

    return;

  }


  const dataCount =
    Math.max(
      0,
      lastRow - 1
    );


  let oldData = [];


  if (dataCount > 0) {

    oldData =
      sheet
        .getRange(
          2,
          1,
          dataCount,
          lastColumn
        )
        .getValues();

  }


  const indexMap =
    createHeaderIndexMap_(
      headers
    );


  const newData = [];


  oldData.forEach(
    function(row) {

      if (
        row.every(
          function(value) {

            return isBlank_(
              value
            );

          }
        )
      ) {

        return;

      }


      const idSuratJalan =
        getMappedValue_(
          row,
          indexMap,
          [
            "ID Surat Jalan"
          ]
        );


      const noSuratJalan =
        getMappedValue_(
          row,
          indexMap,
          [
            "No Surat Jalan"
          ]
        );


      const tanggal =
        getMappedValue_(
          row,
          indexMap,
          [
            "Tanggal"
          ]
        );


      const idPabrik =
        getMappedValue_(
          row,
          indexMap,
          [
            "ID Pabrik"
          ]
        );


      /*
       * Data lama mungkin memakai
       * "Pabrik Tujuan".
       */

      const namaPabrik =
        getMappedValue_(
          row,
          indexMap,
          [
            "Nama Pabrik",
            "Pabrik Tujuan"
          ]
        );


      const alamatPabrik =
        getMappedValue_(
          row,
          indexMap,
          [
            "Alamat Pabrik",
            "Alamat"
          ]
        );


      const noPO =
        getMappedValue_(
          row,
          indexMap,
          [
            "No PO",
            "No PO / Pesanan",
            "No Pesanan"
          ]
        );


      const kendaraan =
        getMappedValue_(
          row,
          indexMap,
          [
            "Kendaraan"
          ]
        );


      const namaSupir =
        getMappedValue_(
          row,
          indexMap,
          [
            "Nama Supir"
          ]
        );


      const jenisPengiriman =
        getMappedValue_(
          row,
          indexMap,
          [
            "Jenis Pengiriman"
          ]
        );


      const keterangan =
        getMappedValue_(
          row,
          indexMap,
          [
            "Keterangan"
          ]
        );


      const totalRoll =
        toNumber_(
          getMappedValue_(
            row,
            indexMap,
            [
              "Total Roll"
            ]
          )
        );


      const totalKg =
        toNumber_(
          getMappedValue_(
            row,
            indexMap,
            [
              "Total Kg"
            ]
          )
        );


      /*
       * TUJUAN TIDAK DIPINDAHKAN.
       */


      newData.push([

        idSuratJalan,

        noSuratJalan,

        tanggal,

        idPabrik,

        namaPabrik,

        alamatPabrik,

        noPO,

        kendaraan,

        namaSupir,

        jenisPengiriman,

        keterangan,

        totalRoll,

        totalKg

      ]);

    }
  );


  sheet.clearContents();


  sheet
    .getRange(
      1,
      1,
      1,
      DB_HEADERS.SURAT_JALAN.length
    )
    .setValues([
      DB_HEADERS.SURAT_JALAN
    ]);


  if (
    newData.length > 0
  ) {

    sheet
      .getRange(
        2,
        1,
        newData.length,
        DB_HEADERS.SURAT_JALAN.length
      )
      .setValues(
        newData
      );

  }

}


/* =========================================================
   MIGRASI DETAIL SURAT JALAN
========================================================= */

function migrateDetailSuratJalanSheet_(
  sheet
) {

  if (!sheet) {

    return;

  }


  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  if (

    lastRow === 0

    ||

    lastColumn === 0

  ) {

    ensureHeader_(
      sheet,
      DB_HEADERS.DETAIL_SURAT_JALAN
    );

    return;

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0]
      .map(
        function(value) {

          return cleanString_(
            value
          );

        }
      );


  /*
   * Sudah format baru.
   */

  if (

    arraysEqual_(
      headers,
      DB_HEADERS.DETAIL_SURAT_JALAN
    )

  ) {

    return;

  }


  const dataCount =
    Math.max(
      0,
      lastRow - 1
    );


  let oldData = [];


  if (dataCount > 0) {

    oldData =
      sheet
        .getRange(
          2,
          1,
          dataCount,
          lastColumn
        )
        .getValues();

  }


  const indexMap =
    createHeaderIndexMap_(
      headers
    );


  const newData = [];


  oldData.forEach(
    function(row) {

      if (
        row.every(
          function(value) {

            return isBlank_(
              value
            );

          }
        )
      ) {

        return;

      }


      const idSuratJalan =
        getMappedValue_(
          row,
          indexMap,
          [
            "ID Surat Jalan"
          ]
        );


      const no =
        getMappedValue_(
          row,
          indexMap,
          [
            "No"
          ]
        );


      const jenisKain =
        getMappedValue_(
          row,
          indexMap,
          [
            "Jenis Kain"
          ]
        );


      const warna =
        getMappedValue_(
          row,
          indexMap,
          [
            "Warna"
          ]
        );


      const noRoll =
        getMappedValue_(
          row,
          indexMap,
          [
            "No Roll"
          ]
        );


      const beratKg =
        getMappedValue_(
          row,
          indexMap,
          [
            "Berat Kg"
          ]
        );


      const keterangan =
        getMappedValue_(
          row,
          indexMap,
          [
            "Keterangan"
          ]
        );


      /*
       * Nama Barang sengaja tidak
       * dipindahkan karena sudah
       * dihapus dari struktur.
       */


      newData.push([

        idSuratJalan,

        toNumber_(
          no
        ),

        jenisKain,

        warna,

        noRoll,

        toNumber_(
          beratKg
        ),

        keterangan

      ]);

    }
  );


  sheet.clearContents();


  sheet
    .getRange(
      1,
      1,
      1,
      DB_HEADERS.DETAIL_SURAT_JALAN.length
    )
    .setValues([
      DB_HEADERS.DETAIL_SURAT_JALAN
    ]);


  if (
    newData.length > 0
  ) {

    sheet
      .getRange(
        2,
        1,
        newData.length,
        DB_HEADERS.DETAIL_SURAT_JALAN.length
      )
      .setValues(
        newData
      );

  }

}


/* =========================================================
   PASTIKAN HEADER
========================================================= */

function ensureHeader_(
  sheet,
  headers
) {

  if (

    !sheet ||

    !headers ||

    !headers.length

  ) {

    return;

  }


  /*
   * Tambahkan kolom kalau jumlah
   * kolom sheet belum cukup.
   */

  const maxColumns =
    sheet.getMaxColumns();


  if (
    maxColumns <
    headers.length
  ) {

    sheet.insertColumnsAfter(
      maxColumns,
      headers.length - maxColumns
    );

  }


  sheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([
      headers
    ]);

}


/* =========================================================
   FORMAT DATABASE SHEET
========================================================= */

function formatDatabaseSheet_(
  sheet,
  headers
) {

  if (!sheet) {

    return;

  }


  sheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setFontWeight(
      "bold"
    )
    .setBackground(
      "#1f4e78"
    )
    .setFontColor(
      "#ffffff"
    )
    .setHorizontalAlignment(
      "center"
    )
    .setVerticalAlignment(
      "middle"
    );


  sheet.setRowHeight(
    1,
    28
  );


  sheet.setFrozenRows(
    1
  );


  /*
   * Auto resize.
   */

  for (
    let col = 1;
    col <= headers.length;
    col++
  ) {

    try {

      sheet.autoResizeColumn(
        col
      );

    } catch (error) {

      // Abaikan.

    }

  }


  /*
   * Batasi lebar kolom
   * agar tidak terlalu besar.
   */

  for (
    let col = 1;
    col <= headers.length;
    col++
  ) {

    try {

      const width =
        sheet.getColumnWidth(
          col
        );


      if (
        width > 280
      ) {

        sheet.setColumnWidth(
          col,
          280
        );

      }

    } catch (error) {

      // Abaikan.

    }

  }

}


/* =========================================================
   HEADER INDEX MAP
========================================================= */

function createHeaderIndexMap_(
  headers
) {

  const map = {};


  headers.forEach(
    function(header, index) {

      const key =
        cleanString_(
          header
        )
        .toLowerCase();


      if (key) {

        map[key] =
          index;

      }

    }
  );


  return map;

}


/* =========================================================
   AMBIL VALUE BERDASARKAN HEADER
========================================================= */

function getMappedValue_(
  row,
  indexMap,
  possibleHeaders
) {

  for (
    let i = 0;
    i < possibleHeaders.length;
    i++
  ) {

    const key =
      cleanString_(
        possibleHeaders[i]
      )
      .toLowerCase();


    if (
      Object.prototype
        .hasOwnProperty
        .call(
          indexMap,
          key
        )
    ) {

      const index =
        indexMap[key];


      if (

        index >= 0 &&

        index < row.length

      ) {

        return row[index];

      }

    }

  }


  return "";

}


/* =========================================================
   BANDING ARRAY
========================================================= */

function arraysEqual_(
  a,
  b
) {

  if (

    !Array.isArray(a) ||

    !Array.isArray(b)

  ) {

    return false;

  }


  if (
    a.length !==
    b.length
  ) {

    return false;

  }


  for (
    let i = 0;
    i < a.length;
    i++
  ) {

    if (

      cleanString_(
        a[i]
      )

      !==

      cleanString_(
        b[i]
      )

    ) {

      return false;

    }

  }


  return true;

}


/* =========================================================
   CEK KOSONG
========================================================= */

function isBlank_(
  value
) {

  return (

    value === null ||

    value === undefined ||

    cleanString_(
      value
    ) === ""

  );

}


/* =========================================================
   CLEAN STRING
========================================================= */

function cleanString_(
  value
) {

  if (

    value === null ||

    value === undefined

  ) {

    return "";

  }


  return String(
    value
  ).trim();

}


/* =========================================================
   CONVERT NUMBER
========================================================= */

function toNumber_(
  value
) {

  if (

    value === null ||

    value === undefined ||

    value === ""

  ) {

    return 0;

  }


  if (
    typeof value === "number"
  ) {

    return isNaN(value)
      ? 0
      : value;

  }


  let text =
    cleanString_(
      value
    );


  /*
   * Hapus Rp dan spasi.
   */

  text =
    text
      .replace(
        /rp/gi,
        ""
      )
      .replace(
        /\s/g,
        ""
      );


  /*
   * Format:
   * 1.500.000,50
   */

  if (

    text.indexOf(".") !== -1 &&

    text.indexOf(",") !== -1

  ) {

    text =
      text
        .replace(
          /\./g,
          ""
        )
        .replace(
          ",",
          "."
        );

  }


  /*
   * Format:
   * 15.000
   */

  else if (

    text.indexOf(".") !== -1

  ) {

    const parts =
      text.split(".");


    if (

      parts.length === 2 &&

      parts[1].length === 3

    ) {

      text =
        text.replace(
          ".",
          ""
        );

    }

  }


  /*
   * Format:
   * 15,5
   */

  else if (

    text.indexOf(",") !== -1

  ) {

    text =
      text.replace(
        ",",
        "."
      );

  }


  const result =
    Number(
      text
    );


  return isNaN(result)
    ? 0
    : result;

}


/* =========================================================
   CEK DATABASE
========================================================= */

function cekDatabaseV3() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const result = [];


  Object.keys(
    DB_CONFIG.SHEETS
  )
  .forEach(
    function(key) {

      const sheetName =
        DB_CONFIG.SHEETS[key];


      const sheet =
        ss.getSheetByName(
          sheetName
        );


      result.push({

        key:
          key,

        sheet:
          sheetName,

        tersedia:
          !!sheet,

        baris:
          sheet
            ? sheet.getLastRow()
            : 0,

        kolom:
          sheet
            ? sheet.getLastColumn()
            : 0

      });

    }
  );


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;

}


/* =========================================================
   CEK PEMBUKUAN
========================================================= */

function cekStrukturPembukuanV3() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const result = {};


  [

    DB_CONFIG.SHEETS.PEMBUKUAN_GREY,

    DB_CONFIG.SHEETS.PEMBUKUAN_KAIN_JADI

  ]
  .forEach(
    function(sheetName) {

      const sheet =
        ss.getSheetByName(
          sheetName
        );


      if (!sheet) {

        result[sheetName] = {

          exists:
            false

        };

        return;

      }


      const headers =
        sheet
          .getRange(
            1,
            1,
            1,
            sheet.getLastColumn()
          )
          .getValues()[0];


      result[sheetName] = {

        exists:
          true,

        jumlahKolom:
          headers.length,

        headers:
          headers,

        compatible:
          arraysEqual_(
            headers,
            DB_HEADERS.PEMBUKUAN
          )

      };

    }
  );


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;

}


/* =========================================================
   CEK SURAT JALAN
========================================================= */

function cekStrukturSuratJalanV3() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const result = {};


  /*
   * HEADER SURAT JALAN
   */

  const suratJalanSheet =
    ss.getSheetByName(
      DB_CONFIG.SHEETS.SURAT_JALAN
    );


  if (suratJalanSheet) {

    const headers =
      suratJalanSheet
        .getRange(
          1,
          1,
          1,
          suratJalanSheet.getLastColumn()
        )
        .getValues()[0];


    result.suratJalan = {

      exists:
        true,

      jumlahKolom:
        headers.length,

      headers:
        headers,

      compatible:
        arraysEqual_(
          headers,
          DB_HEADERS.SURAT_JALAN
        )

    };

  }

  else {

    result.suratJalan = {

      exists:
        false

    };

  }


  /*
   * DETAIL SURAT JALAN
   */

  const detailSheet =
    ss.getSheetByName(
      DB_CONFIG.SHEETS.DETAIL_SURAT_JALAN
    );


  if (detailSheet) {

    const headers =
      detailSheet
        .getRange(
          1,
          1,
          1,
          detailSheet.getLastColumn()
        )
        .getValues()[0];


    result.detailSuratJalan = {

      exists:
        true,

      jumlahKolom:
        headers.length,

      headers:
        headers,

      compatible:
        arraysEqual_(
          headers,
          DB_HEADERS.DETAIL_SURAT_JALAN
        )

    };

  }

  else {

    result.detailSuratJalan = {

      exists:
        false

    };

  }


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;

}


/* =========================================================
   CEK LOG AKTIVITAS
========================================================= */

function cekStrukturLogAktivitasV3() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      DB_CONFIG.SHEETS.LOG_AKTIVITAS
    );


  if (!sheet) {

    return {

      exists:
        false

    };

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0];


  const result = {

    exists:
      true,

    jumlahKolom:
      headers.length,

    headers:
      headers,

    compatible:
      arraysEqual_(
        headers,
        DB_HEADERS.LOG_AKTIVITAS
      )

  };


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;

}