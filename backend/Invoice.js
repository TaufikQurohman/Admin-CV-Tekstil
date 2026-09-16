/* ============================================================
   INVOICE
   ============================================================ */

function saveInvoice(data, token) {
  const username = requireAuth_(token);
  if (!data) throw new Error("Data invoice tidak ditemukan.");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    validateInvoice_(data);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invoiceSheet = getSheet_(ss, CONFIG.SHEETS.INVOICE);
    const detailSheet = getSheet_(ss, CONFIG.SHEETS.DETAIL_INVOICE);
    const greySheet = getSheet_(ss, CONFIG.SHEETS.PEMBUKUAN_GREY);
    const jadiSheet = getSheet_(ss, CONFIG.SHEETS.PEMBUKUAN_JADI);

    const kategori = normalizeKategoriPenjualan_(data.jenisPenjualan || data.jenisKain);
    if (!kategori) throw new Error("Kategori penjualan harus Grey atau Kain Jadi.");

    const idInvoice = getNextId_(invoiceSheet, "INV");
    const noInvoice = generateInvoiceNumber_();

    const built = buildInvoiceRows_(
      data,
      idInvoice,
      noInvoice,
      greySheet,
      jadiSheet
    );

    const totalInvoice = built.totalInvoice;
    const dp = Math.max(0, toNumber_(data.dp));
    if (dp > totalInvoice) throw new Error("DP tidak boleh lebih besar dari total invoice.");

    const sisaTagihan = totalInvoice - dp;
    const status = sisaTagihan <= 0 ? "Lunas" : (dp > 0 ? "DP" : "Belum Lunas");

    invoiceSheet.appendRow([
      idInvoice,
      noInvoice,
      parseDate_(data.tanggal),
      cleanString_(data.namaPembeli),
      cleanString_(data.alamatPembeli),
      cleanString_(data.noPO || data.noPo),
      kategori,
      cleanString_(data.jenisPembayaran),
      dp,
      sisaTagihan,
      cleanString_(data.keterangan),
      totalInvoice,
      status
    ]);

    if (built.detailRows.length) {
      detailSheet.getRange(detailSheet.getLastRow()+1,1,built.detailRows.length,9).setValues(built.detailRows);
    }
    if (built.greyRows.length) {
      greySheet.getRange(greySheet.getLastRow()+1,1,built.greyRows.length,15).setValues(built.greyRows);
    }
    if (built.jadiRows.length) {
      jadiSheet.getRange(jadiSheet.getLastRow()+1,1,built.jadiRows.length,15).setValues(built.jadiRows);
    }

    writeActivityLog_(username, "Membuat Invoice", "Invoice", idInvoice, noInvoice);

    return {
      success: true,
      idInvoice,
      noInvoice,
      tanggal: data.tanggal,
      namaPembeli: cleanString_(data.namaPembeli),
      jenisKain: kategori,
      jenisPenjualan: kategori,
      totalInvoice,
      dp,
      sisaTagihan,
      status,
      jumlahRoll: built.jumlahRoll,
      totalKgPembukuan: built.totalKg,
      totalMeterPembukuan: built.totalMeter
    };
  } finally {
    lock.releaseLock();
  }
}

function validateInvoice_(data) {
  if (!data) throw new Error("Data invoice tidak ditemukan.");
  if (!data.tanggal) throw new Error("Tanggal invoice wajib diisi.");
  if (!cleanString_(data.namaPembeli)) throw new Error("Nama pembeli wajib diisi.");
  const kategori = normalizeKategoriPenjualan_(data.jenisPenjualan || data.jenisKain);
  if (!kategori) throw new Error("Kategori penjualan harus Grey atau Kain Jadi.");
  if (!Array.isArray(data.detail) || data.detail.length === 0) throw new Error("Detail kain belum diisi.");
  if (data.jenisPembayaran && CONFIG.PEMBAYARAN.indexOf(cleanString_(data.jenisPembayaran)) === -1) {
    throw new Error("Jenis pembayaran tidak valid.");
  }
}

function generateInvoiceNumber_() {

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
    "INVOICE_COUNTER_" +
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
    "INV-" +
    periode +
    "-" +
    String(counter)
      .padStart(3, "0")
  );

}

function getInvoices(token) {
  requireAuth_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss, CONFIG.SHEETS.INVOICE);
  return getDataRows_(sheet).map(function(row) {
    return {
      idInvoice: row[0],
      noInvoice: row[1],
      tanggal: formatDateForClient_(row[2]),
      namaPembeli: row[3],
      alamatPembeli: row[4],
      noPO: row[5],
      jenisKain: row[6],
      jenisPenjualan: row[6],
      jenisPembayaran: row[7],
      dp: row[8],
      sisaTagihan: row[9],
      keterangan: row[10],
      totalInvoice: row[11],
      status: row[12] || ""
    };
  });
}

function getInvoiceById(idInvoice, token) {
  requireAuth_(token);
  if (!idInvoice) throw new Error("ID Invoice tidak ditemukan.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoiceSheet = getSheet_(ss, CONFIG.SHEETS.INVOICE);
  const detailSheet = getSheet_(ss, CONFIG.SHEETS.DETAIL_INVOICE);
  const invoice = getDataRows_(invoiceSheet).find(function(row){ return String(row[0]) === String(idInvoice); });
  if (!invoice) throw new Error("Invoice tidak ditemukan.");

  const detail = getDataRows_(detailSheet).filter(function(row){
    return String(row[0]) === String(idInvoice);
  }).map(function(row){
    return {
      idInvoice: row[0],
      no: row[1],
      jenisKain: row[2],
      namaKain: row[2],
      warna: row[3],
      noRoll: row[4],
      qty: row[5],
      berat: row[5],
      satuan: row[6],
      hargaSatuan: row[7],
      jumlah: row[8]
    };
  });

  return {
    idInvoice: invoice[0],
    noInvoice: invoice[1],
    tanggal: formatDateForClient_(invoice[2]),
    namaPembeli: invoice[3],
    alamatPembeli: invoice[4],
    noPO: invoice[5],
    jenisKain: invoice[6],
    jenisPenjualan: invoice[6],
    jenisPembayaran: invoice[7],
    dp: invoice[8],
    sisaTagihan: invoice[9],
    keterangan: invoice[10],
    totalInvoice: invoice[11],
    status: invoice[12],
    detail: detail
  };
}

function buildInvoiceRows_(data, idInvoice, noInvoice, greySheet, jadiSheet) {
  const detailRows = [];
  const greyRows = [];
  const jadiRows = [];
  let totalInvoice = 0;
  let jumlahRoll = 0;
  let totalKg = 0;
  let totalMeter = 0;
  const kategori = normalizeKategoriPenjualan_(data.jenisPenjualan || data.jenisKain);
  const isGrey = kategori === "Grey";
  let greyCounter = getNextSequenceStart_(greySheet, "PEM");
  let jadiCounter = getNextSequenceStart_(jadiSheet, "PEM");

  data.detail.forEach(function(item, index){
    const jenisKain = cleanString_(item.jenisKain || item.namaKain);
    const warna = cleanString_(item.warna);
    const noRoll = cleanString_(item.noRoll);
    const satuan = cleanString_(item.satuan || "Kg");
    const qty = toNumber_(item.qty !== undefined ? item.qty : item.berat);
    const hargaSatuan = toNumber_(item.hargaSatuan);

    if (!jenisKain && !warna && !noRoll && qty===0 && hargaSatuan===0) return;
    if (!jenisKain) throw new Error("Jenis kain pada detail wajib diisi.");
    if (qty <= 0) throw new Error("Qty pada detail harus lebih dari 0.");
    if (hargaSatuan < 0) throw new Error("Harga satuan tidak valid.");
    if (CONFIG.SATUAN.indexOf(satuan) === -1) throw new Error("Satuan hanya boleh Kg atau Meter.");

    const jumlah = qty * hargaSatuan;
    totalInvoice += jumlah;
    jumlahRoll++;
    if (satuan.toLowerCase()==="kg") totalKg += qty;
    if (satuan.toLowerCase()==="meter") totalMeter += qty;

    detailRows.push([idInvoice,index+1,jenisKain,warna,noRoll,qty,satuan,hargaSatuan,jumlah]);

    const pembukuanId = isGrey
      ? formatSequentialId_("PEM", greyCounter++)
      : formatSequentialId_("PEM", jadiCounter++);
    const row = [
      pembukuanId,
      idInvoice,
      noInvoice,
      parseDate_(data.tanggal),
      cleanString_(data.namaPembeli),
      qty,
      satuan,
      satuan.toLowerCase()==="kg" ? qty : "",
      jenisKain,
      warna,
      hargaSatuan,
      cleanString_(data.jenisPembayaran),
      cleanString_(data.buktiPembayaran),
      cleanString_(data.keterangan),
      jumlah
    ];
    if (isGrey) greyRows.push(row); else jadiRows.push(row);
  });

  if (!detailRows.length) throw new Error("Minimal harus ada satu detail kain.");
  return {detailRows,greyRows,jadiRows,totalInvoice,jumlahRoll,totalKg,totalMeter};
}

function getInvoiceStatusById_(idInvoice, token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss, CONFIG.SHEETS.INVOICE);
  const row = getDataRows_(sheet).find(function(r){ return String(r[0])===String(idInvoice); });
  return row ? String(row[12]||"") : "";
}

function searchInvoices(keyword, year, month, status, token) {
  requireAuth_(token);
  const q = cleanString_(keyword).toLowerCase();
  const y = Number(year)||0;
  const m = Number(month)||0;
  const s = cleanString_(status).toLowerCase();
  return getInvoices(token).filter(function(row){
    if (q && [row.noInvoice,row.namaPembeli,row.noPO,row.jenisPenjualan].join(" ").toLowerCase().indexOf(q)===-1) return false;
    if (y && m && !isSameMonth_(row.tanggal,y,m)) return false;
    if (s && String(row.status||"").toLowerCase()!==s) return false;
    return true;
  });
}

function updateInvoice(data, token) {
  const username = requireAuth_(token);
  if (!data || !data.idInvoice) throw new Error("ID Invoice untuk edit tidak ditemukan.");
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    validateInvoice_(data);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invoiceSheet = getSheet_(ss,CONFIG.SHEETS.INVOICE);
    const detailSheet = getSheet_(ss,CONFIG.SHEETS.DETAIL_INVOICE);
    const greySheet = getSheet_(ss,CONFIG.SHEETS.PEMBUKUAN_GREY);
    const jadiSheet = getSheet_(ss,CONFIG.SHEETS.PEMBUKUAN_JADI);
    const dataRows = getDataRows_(invoiceSheet);
    const idx = dataRows.findIndex(function(r){ return String(r[0])===String(data.idInvoice); });
    if (idx<0) throw new Error("Invoice tidak ditemukan.");
    if (String(dataRows[idx][12]||"")==="Dibatalkan") throw new Error("Invoice yang sudah dibatalkan tidak dapat diedit.");

    const idInvoice = dataRows[idx][0];
    const noInvoice = dataRows[idx][1];
    const built = buildInvoiceRows_(data,idInvoice,noInvoice,greySheet,jadiSheet);
    const dp = Math.max(0,toNumber_(data.dp));
    if (dp>built.totalInvoice) throw new Error("DP tidak boleh lebih besar dari total invoice.");
    const sisa = built.totalInvoice-dp;
    const status = sisa<=0 ? "Lunas" : (dp>0 ? "DP" : "Belum Lunas");
    const kategori = normalizeKategoriPenjualan_(data.jenisPenjualan || data.jenisKain);

    invoiceSheet.getRange(idx+2,1,1,13).setValues([[
      idInvoice,noInvoice,parseDate_(data.tanggal),cleanString_(data.namaPembeli),cleanString_(data.alamatPembeli),cleanString_(data.noPO||data.noPo),kategori,cleanString_(data.jenisPembayaran),dp,sisa,cleanString_(data.keterangan),built.totalInvoice,status
    ]]);

    deleteRowsByKey_(detailSheet,1,idInvoice);
    deleteRowsByKey_(greySheet,2,idInvoice);
    deleteRowsByKey_(jadiSheet,2,idInvoice);

    if (built.detailRows.length) detailSheet.getRange(detailSheet.getLastRow()+1,1,built.detailRows.length,9).setValues(built.detailRows);
    if (built.greyRows.length) greySheet.getRange(greySheet.getLastRow()+1,1,built.greyRows.length,15).setValues(built.greyRows);
    if (built.jadiRows.length) jadiSheet.getRange(jadiSheet.getLastRow()+1,1,built.jadiRows.length,15).setValues(built.jadiRows);

    writeActivityLog_(username,"Mengedit Invoice","Invoice",idInvoice,noInvoice);
    return {success:true,idInvoice,noInvoice,totalInvoice:built.totalInvoice,dp,sisaTagihan:sisa,status};
  } finally {
    lock.releaseLock();
  }
}

function voidInvoice(idInvoice, token) {
  const username = requireAuth_(token);
  if (!idInvoice) throw new Error("ID Invoice tidak ditemukan.");
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invoiceSheet = getSheet_(ss,CONFIG.SHEETS.INVOICE);
    const rows = getDataRows_(invoiceSheet);
    const idx = rows.findIndex(function(r){ return String(r[0])===String(idInvoice); });
    if (idx<0) throw new Error("Invoice tidak ditemukan.");
    if (String(rows[idx][12]||"")==="Dibatalkan") throw new Error("Invoice sudah dibatalkan.");
    invoiceSheet.getRange(idx+2,13).setValue("Dibatalkan");
    writeActivityLog_(username,"Membatalkan Invoice","Invoice",idInvoice,rows[idx][1]||"");
    return {success:true,idInvoice,noInvoice:rows[idx][1],status:"Dibatalkan"};
  } finally {
    lock.releaseLock();
  }
}


