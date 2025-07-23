function generateHutangPiutangChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Setup Dashboard');

  const lastRow = sheet.getLastRow();

  // Ambil data range (kategori dan sisa hutang)
  const kategoriRange = sheet.getRange(`A27:A${lastRow - 1}`);  // Exclude "Total"
  const sisaHutangRange = sheet.getRange(`D27:D${lastRow - 1}`);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(kategoriRange)
    .addRange(sisaHutangRange)
    .setPosition(lastRow + 2, 1, 0, 0) // Chart akan ditempatkan 2 baris di bawah data
    .setOption('title', 'Grafik Hutang Piutang')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: 'Sisa Hutang', format: 'short' })
    .setOption('vAxis', { title: 'Kategori' })
    .setOption('colors', ['#2E86AB']) // kamu bisa ganti ke gradient/warna dinamis
    .build();

  sheet.insertChart(chart);
}
