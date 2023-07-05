async function objectToAzureEndPoint(obj, rowNumber) {
  try {
    const url =
      "https://inspection-storage.azurewebsites.net/api/httpTrigger1?code=UDtASrlys7_M3h-lvkMEfJghSkzvvUCbPs3A0kIHeTE7AzFu1NsaGQ==";

    const data = {
      name: "Indraneel",
      data: obj,
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(data),
    };
    const response = await UrlFetchApp.fetch(url, options);

    Browser.msgBox("Successfull");
    rowNumber = rowNumber + 1;
    const sheet = SpreadsheetApp.openById(
      RESPONSES_SPREADSHEET_ID
    ).getSheetByName(RESPONSES_SHEET_NAME);
    const cell = sheet.getRange(rowNumber, 122);
    cell.setValue("Data Added to azure.");
    Browser.msgBox(response);
  } catch (error) {
    const sheet = SpreadsheetApp.openById(
      RESPONSES_SPREADSHEET_ID
    ).getSheetByName(RESPONSES_SHEET_NAME);
    const cell = sheet.getRange(rowNumber, 122);
    cell.setValue(error.message);
  }
}
