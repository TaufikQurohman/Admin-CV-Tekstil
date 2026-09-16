/* ============================================================
   SURAT JALAN
   ============================================================ */

function saveSuratJalan(data, token) {
  const username = requireAuth_(token);
  if (!data) throw new Error("Data surat jalan tidak ditemukan.");
  if (!data.tanggal) throw new Error("Tanggal surat jalan wajib diisi.");
  if (!data.idPabrik) throw new Error("Pabrik tujuan wajib dipilih.");
  if (!cleanString_(data.kendaraan)) throw new Error("Kendaraan wajib diisi.");
  if (!cleanString_(data.namaSupir)) throw new Error("Nama supir wajib diisi.");
  if (!Array.isArray(data.detail) || !data.detail.length) throw new Error("Detail surat jalan belum diisi.");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sjSheet = getSheet_(ss, CONFIG.SHEETS.SURAT_JALAN);
    const detailSheet = getSheet_(ss, CONFIG.SHEETS.DETAIL_SURAT_JALAN);
    const pabrikSheet = getSheet_(ss, CONFIG.SHEETS.MASTER_PABRIK);

    const pabrikRows = getDataRows_(pabrikSheet);
    const pabrik = pabrikRows.find(function(row){ return String(row[0]) === String(data.idPabrik); });
    if (!pabrik) throw new Error("Pabrik tidak ditemukan di Master Pabrik.");

    const idSuratJalan = getNextId_(sjSheet, "SJ");
    const noSuratJalan = generateSuratJalanNumber_();

    const detailRows = [];
    let totalRoll = 0;
    let totalKg = 0;

    data.detail.forEach(function(item, index){
      const jenisKain = cleanString_(item.jenisKain);
      const warna = cleanString_(item.warna);
      const noRoll = cleanString_(item.noRoll);
      const beratKg = toNumber_(item.beratKg);
      const keterangan = cleanString_(item.keterangan);

      if (!jenisKain && !warna && !noRoll && beratKg === 0 && !keterangan) return;
      if (!jenisKain) throw new Error("Jenis kain pada detail surat jalan wajib diisi.");
      if (beratKg <= 0) throw new Error("Berat setiap roll wajib lebih dari 0 Kg.");

      totalRoll++;
      totalKg += beratKg;

      detailRows.push([
        idSuratJalan,
        index + 1,
        jenisKain,
        warna,
        noRoll,
        beratKg,
        keterangan
      ]);
    });

    if (!detailRows.length) throw new Error("Minimal harus ada satu detail kain.");

    sjSheet.appendRow([
      idSuratJalan,
      noSuratJalan,
      parseDate_(data.tanggal),
      pabrik[0],
      pabrik[1],
      pabrik[2],
      cleanString_(data.noPO || data.noPo),
      cleanString_(data.kendaraan),
      cleanString_(data.namaSupir),
      cleanString_(data.jenisPengiriman),
      cleanString_(data.keterangan),
      totalRoll,
      totalKg
    ]);

    detailSheet.getRange(detailSheet.getLastRow()+1,1,detailRows.length,7).setValues(detailRows);

    writeActivityLog_(username, "Membuat Surat Jalan", "Surat Jalan", idSuratJalan, noSuratJalan);

    return { success:true, idSuratJalan, noSuratJalan, totalRoll, totalKg };
  } finally {
    lock.releaseLock();
  }
}

function generateSuratJalanNumber_() {

  const props =
    PropertiesService
      .getScriptProperties();


  const now =
    new Date();


  const timezone =
    Session.getScriptTimeZone();


  const periode =
    Utilities.formatDate(
      now,
      timezone,
      "yyyyMM"
    );


  const key =
    "SJ_COUNTER_" +
    periode;


  let counter =
    Number(
      props.getProperty(key)
    ) || 0;


  counter++;


  props.setProperty(
    key,
    String(counter)
  );


  return (
    "SJ-" +
    periode +
    "-" +
    String(counter)
      .padStart(3, "0")
  );

}

function getSuratJalan(token) {
  requireAuth_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss, CONFIG.SHEETS.SURAT_JALAN);
  return getDataRows_(sheet).map(function(row){
    return {
      idSuratJalan: row[0],
      noSuratJalan: row[1],
      tanggal: formatDateForClient_(row[2]),
      idPabrik: row[3],
      namaPabrik: row[4],
      pabrikTujuan: row[4],
      alamatPabrik: row[5],
      alamat: row[5],
      noPO: row[6],
      kendaraan: row[7],
      namaSupir: row[8],
      jenisPengiriman: row[9],
      keterangan: row[10],
      totalRoll: row[11],
      totalKg: row[12]
    };
  });
}

function getSuratJalanById(idSuratJalan, token) {
  requireAuth_(token);
  if (!idSuratJalan) throw new Error("ID Surat Jalan tidak ditemukan.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sjSheet = getSheet_(ss, CONFIG.SHEETS.SURAT_JALAN);
  const detailSheet = getSheet_(ss, CONFIG.SHEETS.DETAIL_SURAT_JALAN);

  const sj = getDataRows_(sjSheet).find(function(row){ return String(row[0]) === String(idSuratJalan); });
  if (!sj) throw new Error("Surat Jalan tidak ditemukan.");

  const detail = getDataRows_(detailSheet).filter(function(row){
    return String(row[0]) === String(idSuratJalan);
  }).map(function(row){
    return {
      idSuratJalan: row[0],
      no: row[1],
      jenisKain: row[2],
      warna: row[3],
      noRoll: row[4],
      beratKg: row[5],
      keterangan: row[6]
    };
  });

  let totalRoll = detail.length;
  let totalKg = 0;
  detail.forEach(function(item){ totalKg += toNumber_(item.beratKg); });

  return {
    idSuratJalan: sj[0],
    noSuratJalan: sj[1],
    tanggal: formatDateForClient_(sj[2]),
    idPabrik: sj[3],
    namaPabrik: sj[4],
    pabrikTujuan: sj[4],
    alamatPabrik: sj[5],
    alamat: sj[5],
    noPO: sj[6],
    kendaraan: sj[7],
    namaSupir: sj[8],
    jenisPengiriman: sj[9],
    keterangan: sj[10],
    totalRoll: sj[11] || totalRoll,
    totalKg: sj[12] || totalKg,
    detail: detail
  };
}

function searchSuratJalan(keyword, year, month, token) {
  requireAuth_(token);
  const q = cleanString_(keyword).toLowerCase();
  const y = Number(year)||0;
  const m = Number(month)||0;
  return getSuratJalan(token).filter(function(row){
    if (q && [row.noSuratJalan,row.namaPabrik,row.noPO,row.kendaraan,row.namaSupir].join(" ").toLowerCase().indexOf(q)===-1) return false;
    if (y && m && !isSameMonth_(row.tanggal,y,m)) return false;
    return true;
  });
}


