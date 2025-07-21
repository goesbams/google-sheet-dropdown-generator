function generateMonthlyIncomeChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input Transaksi Bulanan");
  const dashboardSheet = ss.getSheetByName("Setup Dashboard");

  const dataRange = inputSheet.getDataRange().getValues();
  const headers = dataRange[0];

  const tanggalIdx = headers.indexOf("Tanggal");
  const kategoriTransaksiIdx = headers.indexOf("Kategori Transaksi");
  const kategoriUtamaIdx = headers.indexOf("Kategori Utama");
  const jumlahTransaksiIdx = headers.indexOf("Jumlah Transaksi");

  const today = new Date();
  const thisMonth = today.getMonth();
  const thisYear = today.getFullYear();

  const incomeData = {};

  for (let i = 1; i < dataRange.length; i++) {
    const row = dataRange[i];
    const tanggal = new Date(row[tanggalIdx]);
    const kategoriTransaksi = row[kategoriTransaksiIdx];
    const kategoriUtama = row[kategoriUtamaIdx];
    let jumlah = row[jumlahTransaksiIdx];

    if (
      kategoriTransaksi === "Pemasukan" &&
      tanggal.getMonth() === thisMonth &&
      tanggal.getFullYear() === thisYear
    ) {
      // Clean jumlah (remove "Rp", dots, etc.)
      if (typeof jumlah === "string") {
        jumlah = parseFloat(jumlah.replace(/[^\d,-]/g, "").replace(",", "."));
      }

      if (!isNaN(jumlah)) {
        if (!incomeData[kategoriUtama]) {
          incomeData[kategoriUtama] = 0;
        }
        incomeData[kategoriUtama] += jumlah;
      }
    }
  }

  // Clear dashboard area
  dashboardSheet.getRange("A1:B20").clearContent();

  // Write data to dashboard
  const labels = Object.keys(incomeData);
  const values = Object.values(incomeData);

  dashboardSheet.getRange(1, 1).setValue("Kategori Utama");
  dashboardSheet.getRange(1, 2).setValue("Total Jumlah");

  for (let i = 0; i < labels.length; i++) {
    dashboardSheet.getRange(i + 2, 1).setValue(labels[i]);
    dashboardSheet.getRange(i + 2, 2).setValue(values[i]);
  }

  // Delete existing charts
  const charts = dashboardSheet.getCharts();
  charts.forEach(chart => dashboardSheet.removeChart(chart));

  // Create new chart
  const chart = dashboardSheet.newChart()
    .asColumnChart()
    .addRange(dashboardSheet.getRange(1, 1, labels.length + 1, 2))
    .setPosition(1, 4, 0, 0)
    .setOption("title", "Pendapatan Bersih Bulan ini (Bulan Berjalan)")
    .setOption("legend", { position: "none" })
    .setOption("colors", ["#4caf50"])
    .build();

  dashboardSheet.insertChart(chart);
}
