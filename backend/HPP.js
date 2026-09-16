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


