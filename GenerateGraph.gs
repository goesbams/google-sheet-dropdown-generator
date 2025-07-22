function rekapHutangPiutang() {
  const sheetName = "Input Transaksi Bulanan";
  const outputSheetName = "Rekap Hutang Piutang";
  
  const masukKategori = ["Hutang", "Gadai", "Pelunasan Piutang", "Cicilan Piutang"];
  const keluarKategori = ["Piutang", "Pelunasan Hutang", "Cicilan Hutang", "Penebusan Gadai"];
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues(); // all data including headers

  const headers = data[0];
  const rows = data.slice(1);

  const kategoriTransaksiCol = headers.indexOf("Kategori Transaksi");
  const kategoriUtamaCol = headers.indexOf("Kategori Utama");
  const jumlahCol = headers.indexOf("Jumlah Transaksi");

  const result = {};

  for (let row of rows) {
    const kategoriTransaksi = row[kategoriTransaksiCol];
    const kategoriUtama = row[kategoriUtamaCol];
    const jumlah = row[jumlahCol];

    if (kategoriTransaksi !== "Hutang Piutang" || typeof jumlah !== "number") continue;

    if (!(kategoriUtama in result)) result[kategoriUtama] = 0;

    if (masukKategori.includes(kategoriUtama)) {
      result[kategoriUtama] += jumlah;
    } else if (keluarKategori.includes(kategoriUtama)) {
      result[kategoriUtama] -= jumlah;
    }
  }

  // Output to new sheet
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  } else {
    outputSheet.clearContents();
  }

  outputSheet.getRange(1, 1, 1, 2).setValues([["Kategori Utama", "Saldo Akhir"]]);

  const outputData = Object.entries(result).map(([kategori, saldo]) => [kategori, saldo]);
  outputSheet.getRange(2, 1, outputData.length, 2).setValues(outputData);

  SpreadsheetApp.flush();
}
