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


