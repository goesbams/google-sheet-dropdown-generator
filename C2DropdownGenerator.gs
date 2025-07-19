function setDynamicDropdownPerRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Input Transaksi Bulanan");
  const startRow = 2;
  const endRow = 501;
  const targetColumn = 3; // Column C

  for (let row = startRow; row <= endRow; row++) {
    const ruleRange = sheet.getRange(row, targetColumn); // Cx
    const sourceRange = sheet.getRange(row, 9, 1, 18); // Ix:Zx → 18 columns
    
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(false)
      .build();

    ruleRange.setDataValidation(rule);
  }

  SpreadsheetApp.flush();
  Logger.log("Dynamic dropdowns applied to C2:C501.");
}
