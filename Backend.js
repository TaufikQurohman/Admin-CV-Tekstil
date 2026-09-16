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
    .createHtmlOutputFromFile("Index")
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
    .evaluate()
    .getContent();

}


/* ============================================================
   LOGIN / SESSION
   ============================================================ */

function login(username, password) {

  const user =
    cleanString_(username);

  const pass =
    String(
      password == null
        ? ""
        : password
    );


  if (
    user !== "admin" ||
    pass !== "Bismillah123"
  ) {

    throw new Error(
      "Username atau password salah."
    );

  }


  const token =
    Utilities.getUuid() +
    "-" +
    Utilities.getUuid();


  CacheService
    .getScriptCache()
    .put(
      "AUTH_" + token,
      "admin",
      CONFIG.SESSION_SECONDS
    );


  return {
    success: true,
    token: token,
    username: "admin"
  };

}


function checkLogin(token) {

  if (!token) {

    return {
      loggedIn: false,
      username: ""
    };

  }


  const username =
    CacheService
      .getScriptCache()
      .get(
        "AUTH_" + String(token)
      );


  return {
    loggedIn: username === "admin",
    username:
      username || ""
  };

}


function logout(token) {

  if (token) {

    CacheService
      .getScriptCache()
      .remove(
        "AUTH_" + String(token)
      );

  }


  return {
    success: true
  };

}


function requireAuth_(token) {

  if (!token) {

    throw new Error(
      "Sesi login tidak ditemukan. Silakan login kembali."
    );

  }


  const username =
    CacheService
      .getScriptCache()
      .get(
        "AUTH_" + String(token)
      );


  if (username !== "admin") {

    throw new Error(
      "Sesi login tidak valid atau sudah kedaluwarsa. Silakan login kembali."
    );

  }


  return username;

}


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
/* ============================================================
   VALIDASI INVOICE
   ============================================================ */

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
/* ============================================================
   NOMOR INVOICE
   Contoh: INV-202609-001
   ============================================================ */

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


/* ============================================================
   SEMUA INVOICE
   ============================================================ */

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
/* ============================================================
   DETAIL INVOICE
   ============================================================ */

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
/* ============================================================
   PEMBUKUAN GREY
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


/* ============================================================
   PEMBUKUAN KAIN JADI
   ============================================================ */

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


/* ============================================================
   PEMBUKUAN UMUM
   ============================================================ */

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
/* ============================================================
   HPP
   ============================================================ */

function saveHpp(data, token) {
  const username = requireAuth_(token);
  if (!data) throw new Error("Data HPP tidak ditemukan.");
  if (!data.tanggal) throw new Error("Tanggal HPP wajib diisi.");
  const jenisTransaksi = cleanString_(data.jenisTransaksi) || "Pembelian";
  if (CONFIG.JENIS_HPP.indexOf(jenisTransaksi) === -1) throw new Error("Jenis transaksi HPP harus Pembelian atau Retur.");

  const jumlah = toNumber_(data.jumlah);
  const hargaSatuan = toNumber_(data.hargaSatuan);
  const pajak = toNumber_(data.pajak);
  if (jumlah <= 0) throw new Error("Jumlah HPP harus lebih dari 0.");
  if (hargaSatuan < 0) throw new Error("Harga satuan HPP tidak valid.");
  const total = (jumlah * hargaSatuan) + pajak;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss, CONFIG.SHEETS.HPP);
  const idHpp = getNextId_(sheet, "HPP");

  sheet.appendRow([
    idHpp,
    parseDate_(data.tanggal),
    cleanString_(data.supplier),
    cleanString_(data.namaBarang),
    cleanString_(data.warna),
    jumlah,
    hargaSatuan,
    pajak,
    total,
    jenisTransaksi,
    cleanString_(data.keterangan)
  ]);

  writeActivityLog_(username, "Menambah HPP", "HPP", idHpp, cleanString_(data.namaBarang));
  return { success:true, idHpp, total };
}
/* ============================================================
   RINGKASAN HPP BULANAN
   ============================================================ */

function getHpp(token) {

  requireAuth_(token);


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    getSheet_(
      ss,
      CONFIG.SHEETS.HPP
    );


  return getDataRows_(
    sheet
  ).map(
    function(row) {

      return {

        idHpp:
          row[0],

        tanggal:
          formatDateForClient_(
            row[1]
          ),

        supplier:
          row[2],

        namaBarang:
          row[3],

        warna:
          row[4],

        jumlah:
          row[5],

        hargaSatuan:
          row[6],

        pajak:
          row[7],

        total:
          row[8],

        jenisTransaksi:
          row[9],

        keterangan:
          row[10]

      };

    }
  );

}



function getHppSummary(
  year,
  month,
  token
) {

  requireAuth_(token);


  const targetYear =
    Number(year);


  const targetMonth =
    Number(month);


  if (
    !targetYear ||
    !targetMonth ||
    targetMonth < 1 ||
    targetMonth > 12
  ) {

    throw new Error(
      "Tahun atau bulan HPP tidak valid."
    );

  }


  const rows =
    getHpp(token);


  let totalPembelian =
    0;


  let totalRetur =
    0;


  rows.forEach(
    function(row) {

      const date =
        parseClientDate_(
          row.tanggal
        );


      if (!date) {
        return;
      }


      if (
        date.getFullYear() !==
        targetYear
      ) {
        return;
      }


      if (
        date.getMonth() + 1 !==
        targetMonth
      ) {
        return;
      }


      const total =
        toNumber_(
          row.total
        );


      if (
        row.jenisTransaksi ===
        "Pembelian"
      ) {

        totalPembelian +=
          total;

      } else if (
        row.jenisTransaksi ===
        "Retur"
      ) {

        totalRetur +=
          total;

      }

    }
  );


  return {

    tahun:
      targetYear,

    bulan:
      targetMonth,

    totalPembelian:
      totalPembelian,

    totalRetur:
      totalRetur,

    hppBersih:
      totalPembelian -
      totalRetur

  };

}


/* ============================================================
   MASUKKAN HPP BULANAN KE PENGELUARAN
   ============================================================ */

function masukkanHppKePengeluaran(year, month, token) {
  const username = requireAuth_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const summary = getHppSummary(year, month, token);
    if (summary.hppBersih <= 0) throw new Error("HPP bersih bulan tersebut tidak lebih dari 0.");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheet_(ss,CONFIG.SHEETS.PENGELUARAN);
    const rows = getDataRows_(sheet);
    const label = "HPP "+monthName_(month)+" "+year;
    if (rows.some(function(row){ return String(row[3]).trim().toLowerCase()===label.trim().toLowerCase(); })) {
      throw new Error("HPP bulan tersebut sudah dimasukkan ke Pengeluaran.");
    }
    const id = getNextId_(sheet,"EXP");
    const tanggal = new Date(Number(year),Number(month),0);
    sheet.appendRow([id,tanggal,"CV Pands Tekstil",label,summary.hppBersih,"Lainnya","HPP","Otomatis dari HPP bulanan"]);
    writeActivityLog_(username,"Memasukkan HPP ke Pengeluaran","HPP",id,label);
    return {success:true,idPengeluaran:id,label:label,jumlah:summary.hppBersih};
  } finally { lock.releaseLock(); }
}
/* ============================================================
   PENGELUARAN
   ============================================================ */

function savePengeluaran(data, token) {
  const username = requireAuth_(token);
  if (!data) throw new Error("Data pengeluaran tidak ditemukan.");
  if (!data.tanggal) throw new Error("Tanggal pengeluaran wajib diisi.");
  const jumlah = toNumber_(data.jumlah);
  if (jumlah <= 0) throw new Error("Jumlah pengeluaran harus lebih dari 0.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss, CONFIG.SHEETS.PENGELUARAN);
  const id = getNextId_(sheet, "EXP");

  sheet.appendRow([
    id,
    parseDate_(data.tanggal),
    cleanString_(data.nama),
    cleanString_(data.kebutuhan),
    jumlah,
    cleanString_(data.jenisPembayaran),
    cleanString_(data.kategori),
    cleanString_(data.keterangan)
  ]);

  writeActivityLog_(username, "Menambah Pengeluaran", "Pengeluaran", id, cleanString_(data.kebutuhan));
  return { success:true, idPengeluaran:id };
}
/* ============================================================
   RINGKASAN PENGELUARAN
   ============================================================ */

function getPengeluaran(
  token
) {

  requireAuth_(token);


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    getSheet_(
      ss,
      CONFIG.SHEETS.PENGELUARAN
    );


  return getDataRows_(
    sheet
  ).map(
    function(row) {

      return {

        idPengeluaran:
          row[0],

        tanggal:
          formatDateForClient_(
            row[1]
          ),

        nama:
          row[2],

        kebutuhan:
          row[3],

        jumlah:
          row[4],

        jenisPembayaran:
          row[5],

        kategori:
          row[6],

        keterangan:
          row[7]

      };

    }
  );

}



function getPengeluaranSummary(
  year,
  month,
  token
) {

  requireAuth_(token);


  const targetYear =
    Number(year);


  const targetMonth =
    Number(month);


  const rows =
    getPengeluaran(token);


  let total =
    0;


  let jumlahTransaksi =
    0;


  rows.forEach(
    function(row) {

      const date =
        parseClientDate_(
          row.tanggal
        );


      if (!date) {
        return;
      }


      if (
        date.getFullYear() !==
        targetYear
      ) {
        return;
      }


      if (
        date.getMonth() + 1 !==
        targetMonth
      ) {
        return;
      }


      total +=
        toNumber_(
          row.jumlah
        );


      jumlahTransaksi++;

    }
  );


  return {

    tahun:
      targetYear,

    bulan:
      targetMonth,

    jumlahTransaksi:
      jumlahTransaksi,

    total:
      total

  };

}


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
/* ============================================================
   NOMOR SURAT JALAN
   Contoh: SJ-202609-001
   ============================================================ */

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


/* ============================================================
   SEMUA SURAT JALAN
   ============================================================ */

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
/* ============================================================
   DETAIL SURAT JALAN
   ============================================================ */

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
/* ============================================================
   SIMPAN MASTER PABRIK
   ============================================================ */

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
/* ============================================================
   DASHBOARD
   ============================================================ */

function getDashboardSummary(year, month, token) {
  requireAuth_(token);
  const now = new Date();
  const targetYear = Number(year) || now.getFullYear();
  const targetMonth = Number(month) || (now.getMonth()+1);

  const invoices = getInvoices(token).filter(function(row){
    return row.status !== "Dibatalkan" && isSameMonth_(row.tanggal, targetYear, targetMonth);
  });
  const grey = getPembukuanGrey(token).filter(function(row){ return getInvoiceStatusById_(row.idInvoice, token) !== "Dibatalkan"; });
  const jadi = getPembukuanKainJadi(token).filter(function(row){ return getInvoiceStatusById_(row.idInvoice, token) !== "Dibatalkan"; });
  const pengeluaran = getPengeluaran(token);

  let penjualanGrey=0, penjualanJadi=0, kgGrey=0, kgJadi=0, meterGrey=0, meterJadi=0;
  grey.forEach(function(row){
    if (isSameMonth_(row.tanggal,targetYear,targetMonth)) {
      penjualanGrey += toNumber_(row.totalPenjualan);
      kgGrey += toNumber_(row.beratKg);
      if (String(row.satuan||"").toLowerCase()==="meter") meterGrey += toNumber_(row.jumlah);
    }
  });
  jadi.forEach(function(row){
    if (isSameMonth_(row.tanggal,targetYear,targetMonth)) {
      penjualanJadi += toNumber_(row.totalPenjualan);
      kgJadi += toNumber_(row.beratKg);
      if (String(row.satuan||"").toLowerCase()==="meter") meterJadi += toNumber_(row.jumlah);
    }
  });

  let totalPengeluaran=0;
  pengeluaran.forEach(function(row){
    if (isSameMonth_(row.tanggal,targetYear,targetMonth)) totalPengeluaran += toNumber_(row.jumlah);
  });

  const hpp = getHppSummary(targetYear,targetMonth,token);
  const totalPenjualan = penjualanGrey + penjualanJadi;
  const totalSuratJalan = getSuratJalan(token).filter(function(row){ return isSameMonth_(row.tanggal,targetYear,targetMonth); }).length;

  return {
    tahun:targetYear,
    bulan:targetMonth,
    totalInvoice:invoices.length,
    penjualanGrey,
    penjualanKainJadi:penjualanJadi,
    totalGrey:penjualanGrey,
    totalJadi:penjualanJadi,
    totalPenjualan,
    penjualanBersih:totalPenjualan-totalPengeluaran,
    kgGrey,
    kgKainJadi:kgJadi,
    meterGrey,
    meterKainJadi:meterJadi,
    totalPengeluaran,
    pengeluaran:totalPengeluaran,
    totalHpp:hpp.hppBersih,
    hpp:hpp.hppBersih,
    totalSuratJalan,
    suratJalan:totalSuratJalan
  };
}

/* ============================================================
   7 FITUR TAMBAHAN
   ============================================================ */

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

function deleteRowsByKey_(sheet, keyColumn, keyValue) {
  const lastRow = sheet.getLastRow();
  if (lastRow<=1) return;
  const rows = sheet.getRange(2,1,lastRow-1,sheet.getLastColumn()).getValues();
  for (let i=rows.length-1;i>=0;i--) {
    if (String(rows[i][keyColumn-1])===String(keyValue)) sheet.deleteRow(i+2);
  }
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

function backupDatabase(token) {
  const username = requireAuth_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const folder = DriveApp.getFileById(file.getId()).getParents().hasNext() ? DriveApp.getFileById(file.getId()).getParents().next() : DriveApp.getRootFolder();
  const stamp = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMdd-HHmmss");
  const copy = file.makeCopy("BACKUP_CV_PANDS_TEKSTIL_"+stamp,folder);
  writeActivityLog_(username,"Backup Database","Sistem",copy.getId(),copy.getName());
  return {success:true,fileId:copy.getId(),fileName:copy.getName(),fileUrl:copy.getUrl()};
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

/* ============================================================
   HELPER
   ============================================================ */

function getSheet_(
  ss,
  sheetName
) {

  const sheet =
    ss.getSheetByName(
      sheetName
    );


  if (!sheet) {

    throw new Error(
      "Sheet '" +
      sheetName +
      "' tidak ditemukan."
    );

  }


  return sheet;

}


function getDataRows_(
  sheet
) {

  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  if (
    lastRow <= 1 ||
    lastColumn <= 0
  ) {

    return [];

  }


  return sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      lastColumn
    )
    .getValues();

}


/* ============================================================
   ID PENDEK
   Contoh:
   INV0001
   PEM0001
   HPP0001
   EXP0001
   SJ0001
   FAB0001
   ============================================================ */

function getNextSequenceStart_(
  sheet,
  prefix
) {

  const rows =
    getDataRows_(
      sheet
    );

  let maxNumber =
    0;

  rows.forEach(function(row) {

    const value =
      cleanString_(row[0]);

    const match =
      value.match(
        new RegExp("^" + prefix + "(\\d+)$", "i")
      );

    if (match) {
      const number = Number(match[1]);
      if (number > maxNumber) {
        maxNumber = number;
      }
    }

  });

  const props =
    PropertiesService
      .getScriptProperties();

  const key =
    "ID_COUNTER_" + prefix;

  const propertyCounter =
    Number(
      props.getProperty(key)
    ) || 0;

  return Math.max(
    maxNumber,
    propertyCounter
  ) + 1;

}


function formatSequentialId_(
  prefix,
  number
) {

  return (
    prefix +
    String(number)
      .padStart(4, "0")
  );

}


function getNextId_(
  sheet,
  prefix
) {

  const rows =
    getDataRows_(
      sheet
    );


  let maxNumber =
    0;


  rows.forEach(
    function(row) {

      const value =
        cleanString_(
          row[0]
        );


      const match =
        value.match(
          new RegExp(
            "^" +
            prefix +
            "(\\d+)$",
            "i"
          )
        );


      if (match) {

        const number =
          Number(
            match[1]
          );


        if (
          number >
          maxNumber
        ) {

          maxNumber =
            number;

        }

      }

    }
  );


  /*
   * Simpan juga counter di Script Properties
   * agar tetap cepat dan aman dari duplikasi.
   */

  const props =
    PropertiesService
      .getScriptProperties();


  const key =
    "ID_COUNTER_" +
    prefix;


  const propertyCounter =
    Number(
      props.getProperty(key)
    ) || 0;


  const nextNumber =
    Math.max(
      maxNumber,
      propertyCounter
    ) + 1;


  props.setProperty(
    key,
    String(nextNumber)
  );


  return (
    prefix +
    String(nextNumber)
      .padStart(4, "0")
  );

}


/* ============================================================
   NORMALISASI KATEGORI PENJUALAN
   ============================================================ */

function normalizeKategoriPenjualan_(
  value
) {

  const text =
    cleanString_(
      value
    )
    .toLowerCase();


  if (
    text === "grey"
  ) {

    return "Grey";

  }


  if (
    text === "kain jadi"
  ) {

    return "Kain Jadi";

  }


  return "";

}


/* ============================================================
   STRING
   ============================================================ */

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


/* ============================================================
   ANGKA
   ============================================================ */

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
    typeof value ===
    "number"
  ) {

    return isNaN(value)
      ? 0
      : value;

  }


  let text =
    String(value)
      .trim();


  text =
    text
      .replace(
        /Rp/gi,
        ""
      )
      .replace(
        /\s/g,
        ""
      );


  if (
    text.includes(".") &&
    text.includes(",")
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

  } else if (
    text.includes(".")
  ) {

    const parts =
      text.split(".");


    if (
      parts.length === 2 &&
      parts[1].length === 3
    ) {

      text =
        text.replace(
          /\./g,
          ""
        );

    }

  } else if (
    text.includes(",")
  ) {

    text =
      text.replace(
        ",",
        "."
      );

  }


  const number =
    Number(text);


  return isNaN(number)
    ? 0
    : number;

}


/* ============================================================
   TANGGAL
   ============================================================ */

function parseDate_(
  value
) {

  if (
    value instanceof Date &&
    !isNaN(
      value.getTime()
    )
  ) {

    return value;

  }


  if (!value) {

    throw new Error(
      "Tanggal tidak valid."
    );

  }


  const text =
    String(value)
      .trim();


  const iso =
    text.match(
      /^(\d{4})-(\d{2})-(\d{2})$/
    );


  if (iso) {

    return new Date(
      Number(iso[1]),
      Number(iso[2]) - 1,
      Number(iso[3])
    );

  }


  const indo =
    text.match(
      /^(\d{2})\/(\d{2})\/(\d{4})$/
    );


  if (indo) {

    return new Date(
      Number(indo[3]),
      Number(indo[2]) - 1,
      Number(indo[1])
    );

  }


  const date =
    new Date(text);


  if (
    isNaN(
      date.getTime()
    )
  ) {

    throw new Error(
      "Format tanggal tidak valid: " +
      text
    );

  }


  return date;

}


/* ============================================================
   FORMAT TANGGAL UNTUK FRONTEND
   ============================================================ */

function formatDateForClient_(
  value
) {

  if (!value) {

    return "";

  }


  let date;


  if (
    value instanceof Date
  ) {

    date =
      value;

  } else {

    date =
      new Date(value);

  }


  if (
    isNaN(
      date.getTime()
    )
  ) {

    return "";

  }


  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

}


/* ============================================================
   PARSE TANGGAL DARI FRONTEND
   ============================================================ */

function parseClientDate_(
  value
) {

  if (!value) {

    return null;

  }


  try {

    return parseDate_(
      value
    );

  } catch (error) {

    return null;

  }

}


/* ============================================================
   CEK BULAN
   ============================================================ */

function isSameMonth_(
  dateValue,
  year,
  month
) {

  const date =
    parseClientDate_(
      dateValue
    );


  if (!date) {

    return false;

  }


  return (
    date.getFullYear() ===
    Number(year)
  ) &&
  (
    date.getMonth() + 1 ===
    Number(month)
  );

}


/* ============================================================
   NAMA BULAN
   ============================================================ */

function formatDateTimeForClient_(value) {
  if (!value) return "";
  let date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function monthName_(
  month
) {

  const names = [

    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember"

  ];


  return (
    names[
      Number(month) - 1
    ] || ""
  );

}
