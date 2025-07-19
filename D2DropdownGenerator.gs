function setDynamicDropdownKolomD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Input Transaksi Bulanan");
  const startRow = 2;
  const endRow = 501;
  const targetColumn = 4; // Column D
  const sourceStartColumn = 27; // Column AA (27), sampai AJ (36)
  const numColumns = 10; // AJ - AA + 1

  for (let row = startRow; row <= endRow; row++) {
    const ruleRange = sheet.getRange(row, targetColumn); // Dx
    const sourceRange = sheet.getRange(row, sourceStartColumn, 1, numColumns); // AAx:AJx

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(false)
      .build();

    ruleRange.setDataValidation(rule);
  }

  SpreadsheetApp.flush();
  Logger.log("Dropdown untuk kolom D2:D501 sudah diterapkan.");
}
