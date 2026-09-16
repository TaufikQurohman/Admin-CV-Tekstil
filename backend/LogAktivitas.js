/* ============================================================
   LOG AKTIVITAS
   ============================================================ */

function writeActivityLog_(user, aktivitas, modul, idData, keterangan) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_AKTIVITAS);
    if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.LOG_AKTIVITAS);
    if (sheet.getLastRow()===0) {
      sheet.getRange(1,1,1,7).setValues([["ID Log","Tanggal & Waktu","User","Aktivitas","Modul","ID Data","Keterangan"]]);
      sheet.setFrozenRows(1);
    }
    const idLog = getNextId_(sheet,"LOG");
    sheet.appendRow([idLog,new Date(),cleanString_(user),cleanString_(aktivitas),cleanString_(modul),cleanString_(idData),cleanString_(keterangan)]);
  } catch (e) {
    // Log aktivitas tidak boleh menggagalkan transaksi utama.
  }
}

function getActivityLog(token) {
  requireAuth_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_AKTIVITAS);
  if (!sheet) return [];
  return getDataRows_(sheet).map(function(row){
    return {idLog:row[0],tanggalWaktu: formatDateTimeForClient_(row[1]),user:row[2],aktivitas:row[3],modul:row[4],idData:row[5],keterangan:row[6]};
  });
}

function recordPrintActivity(modul, idData, token) {
  const username = requireAuth_(token);
  writeActivityLog_(username,"Cetak / Print Ulang",modul,idData,"");
  return {success:true};
}


