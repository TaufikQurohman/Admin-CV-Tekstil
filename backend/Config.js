/**
 * ============================================================
 * CV PANDS TEKSTIL
 * BACKEND WEB APP V3
 * ============================================================
 *
 * Database V3:
 * 01_INVOICE
 * 02_DETAIL_INVOICE
 * 03_PEMBUKUAN_GREY
 * 04_PEMBUKUAN_KAIN_JADI
 * 05_HPP
 * 06_PENGELUARAN
 * 07_SURAT_JALAN
 * 08_DETAIL_SURAT_JALAN
 * 09_MASTER_PABRIK
 *
 * Login:
 * Username : admin
 * Password : ~
 *
 * Catatan:
 * - databasev3.gs tetap dipakai untuk setup struktur database.
 * - Pembukuan Grey/Kain Jadi mendukung Qty dalam Kg maupun Meter.
 * - Kode ini TIDAK berisi setupDatabaseV3() agar tidak duplikat.
 * - Master Kategori tidak digunakan lagi.
 * ============================================================
 */


/* ============================================================
   KONFIGURASI
   ============================================================ */

const CONFIG = {

  SHEETS: {
    INVOICE: "01_INVOICE",
    DETAIL_INVOICE: "02_DETAIL_INVOICE",
    PEMBUKUAN_GREY: "03_PEMBUKUAN_GREY",
    PEMBUKUAN_JADI: "04_PEMBUKUAN_KAIN_JADI",
    HPP: "05_HPP",
    PENGELUARAN: "06_PENGELUARAN",
    SURAT_JALAN: "07_SURAT_JALAN",
    DETAIL_SURAT_JALAN: "08_DETAIL_SURAT_JALAN",
    MASTER_PABRIK: "09_MASTER_PABRIK",
    LOG_AKTIVITAS: "10_LOG_AKTIVITAS"
  },

  PEMBAYARAN: [
    "Cash",
    "Transfer",
    "QRIS",
    "Debit",
    "Lainnya"
  ],

  KATEGORI_PENJUALAN: [
    "Grey",
    "Kain Jadi"
  ],

  SATUAN: [
    "Kg",
    "Meter"
  ],

  JENIS_HPP: [
    "Pembelian",
    "Retur"
  ],

  SESSION_SECONDS: 21600

};


/* ============================================================
   WEB APP
   ============================================================ */

function doGet() {

  return HtmlService
    .createTemplateFromFile("Index").evaluate()
    .setTitle("CV Pands Tekstil")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/* ============================================================
   INCLUDE HTML
   Dipakai jika Index dipisah menjadi beberapa file.
   ============================================================ */

function include(filename) {

  return HtmlService
    .createTemplateFromFile(filename)
    .getRawContent();

}



