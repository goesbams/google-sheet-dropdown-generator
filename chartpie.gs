function generatePengeluaranPieChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sheet1"); // Ganti dengan nama sheet Anda

  const dataRange = sheet.getRange("A1:B15"); // Asumsikan data dimulai di A1:B15 (Kategori & Total)
  
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataRange)
    .setOption("title", "Persentase Pengeluaran")
    .setPosition(sheet.getRange(2, 5)) // Posisi grafik di kolom E baris 2
    .build();

  sheet.insertChart(chart);
}
