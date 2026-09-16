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


