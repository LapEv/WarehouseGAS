function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const wsData = SpreadsheetApp.openById(params.ssID).getSheetByName(
    params.ssSheet
  );
  if (params.log) {
    const getLastColumn = wsData
      .getRange(1, params.startColumn)
      .getNextDataCell(SpreadsheetApp.Direction.NEXT)
      .getColumn();
    const getLog = wsData
      .getRange(params.startRow, getLastColumn, 1, 1)
      .getValue();
    const curDate = Utilities.formatDate(
      new Date(),
      'GMT+3',
      'dd.MM.yyyy HH:mm'
    );
    const logText = `${curDate}:: ${params.user} add line "${params.data}"`;
    const log = getLog ? `${getLog}\n${logText}` : logText;
    params.data.push(log);
    wsData
      .getRange(
        params.startRow,
        params.startColumn,
        params.rows,
        params.columns + 1
      )
      .setValues([params.data]);
    return;
  }
  wsData
    .getRange(params.startRow, params.startColumn, params.rows, params.columns)
    .setValues(params.data);
}
