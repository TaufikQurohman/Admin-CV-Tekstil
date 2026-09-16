/* ============================================================
   PEMBUKUAN
   ============================================================ */

function getPembukuanGrey(token) {

  requireAuth_(token);


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  return getPembukuan_(
    ss,
    CONFIG.SHEETS.PEMBUKUAN_GREY
  );

}

function getPembukuanKainJadi(
  token
) {

  requireAuth_(token);


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  return getPembukuan_(
    ss,
    CONFIG.SHEETS.PEMBUKUAN_JADI
  );

}

function getPembukuan_(ss, sheetName) {
  const sheet = getSheet_(ss, sheetName);
  return getDataRows_(sheet).map(function(row){
    return {
      idPembukuan: row[0],
      idInvoice: row[1],
      noInvoice: row[2],
      tanggal: formatDateForClient_(row[3]),
      namaPembeli: row[4],
      jumlah: row[5],
      satuan: row[6],
      beratKg: row[7],
      jumlahKg: row[7],
      jenisKain: row[8],
      warna: row[9],
      hargaSatuan: row[10],
      harga: row[10],
      jenisPembayaran: row[11],
      buktiPembayaran: row[12],
      keterangan: row[13],
      totalPenjualan: row[14]
    };
  });
}


