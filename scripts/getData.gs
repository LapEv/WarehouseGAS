function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const wsData = SpreadsheetApp.openById(params.ssID).getSheetByName(
    params.sheet
  );
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const getLastColumn = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.NEXT)
    .getColumn();
  const data = wsData.getRange(2, 1, getLastRow - 1, getLastColumn).getValues();
  return ContentService.createTextOutput(data);
}
