/* ============================================================
   MASTER PABRIK
   ============================================================ */

function getMasterPabrik(token) {
  requireAuth_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss, CONFIG.SHEETS.MASTER_PABRIK);
  return getDataRows_(sheet).map(function(row){
    return {
      idPabrik: row[0],
      namaPabrik: row[1],
      alamat: row[2],
      kontak: row[3],
      keterangan: row[4]
    };
  });
}

function saveMasterPabrik(data, token) {
  const username = requireAuth_(token);
  if (!data) throw new Error("Data pabrik tidak ditemukan.");
  if (!cleanString_(data.namaPabrik)) throw new Error("Nama pabrik wajib diisi.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss, CONFIG.SHEETS.MASTER_PABRIK);
  const idPabrik = getNextId_(sheet, "FAB");

  sheet.appendRow([
    idPabrik,
    cleanString_(data.namaPabrik),
    cleanString_(data.alamat),
    cleanString_(data.kontak),
    cleanString_(data.keterangan)
  ]);

  writeActivityLog_(username, "Menambah Master Pabrik", "Master Pabrik", idPabrik, cleanString_(data.namaPabrik));
  return { success:true, idPabrik:idPabrik };
}


