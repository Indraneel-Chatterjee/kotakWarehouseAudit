function onEdit(e) {
  if (e.range.columnEnd === 121 && e.value === "Yes") {
    const rowNum = Number(e.range.rowStart) - 1;
    fillFormAnswers(rowNum);
  } else {
    const sheet = SpreadsheetApp.openById(
      RESPONSES_SPREADSHEET_ID
    ).getSheetByName(RESPONSES_SHEET_NAME);
    const rowNum = e.range.rowStart;
    const cell = sheet.getRange(rowNum, 122);
    cell.setValue("");
  }
}
