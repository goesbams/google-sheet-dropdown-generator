function generatePendapatanPieChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Setup Dashboard"); // Sesuaikan nama sheet Anda

  const dataRange = sheet.getRange("A2:B8"); 
  
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataRange)
    .setOption("title", "Persentase Pemasukan")
    .setPosition(2, 5, 0, 0) // Baris 2, kolom 5 (Kolom E), offset 0px
    .build();

  sheet.insertChart(chart);
}
