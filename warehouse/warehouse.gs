const urlWriteToSheetScript =
  'https://script.google.com/macros/s/AKfycbyczQsUNQWUyaTO5FS1-sqr1ZBRBFvzo_ezlXjnd3L0YMuul1iC5J4V8IkFxaKCcW9H/exec';
const urlWriteToDriveScript =
  'https://script.google.com/macros/s/AKfycbyGguDmQK0-SulZ4X1YH9BnCIYBROJ0hMUe7XXylnKjIkpgDrTmRLw8xAApcSKcqeDF/exec';
const urlGetDataScript =
  'https://script.google.com/macros/s/AKfycbxCcFhNhdtmepMMHGbjwetXU_8GyEOAAjYnROGixnUzgvdGJ_LVxGuv7iXMs_MU9F_Tqw/exec';
const ssIDWarehouse = '1QHenRbqyifedRX-XN4uLMOnOiBFm5XYMzzjuksL2yyQ';
const ssIDAdmin = '11x7Mh-OCuuC1WOC26WEDoVdx5jpZqTCCQyEvJks1OrQ';
const ssNameUsers = 'Пользователи';
const ssNameDataZIP = 'ЗИП Данные';
const ssNameDataPodmena = 'Подмена Данные';
const ssNameClients = 'Клиенты';

const ssNameClassifierZIP = 'Классификатор ЗИП';
const ssNameClassifierPodmena = 'Классификатор подмены';
const ssNameClassifier = 'Классификатор';
const ssNameModel = 'Модели';
const ssNameData = 'Данные';

const folderUploadArrivalZIP = '1XIq0WIheYUmlj44c7nvA-8ZftRmWmXqV';
const folderUploadArrivalPodmena = '1z2i4H2zcmHpIQh7gF-4Umc68XaNe5kH5';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Warehouse').setXFrameOptionsMode(
    HtmlService.XFrameOptionsMode.ALLOWALL
  );
}

function checkAccess(func) {
  const user = Session.getActiveUser().getEmail().toLowerCase();
  const options = {
    ssID: ssIDAdmin,
    sheet: ssNameUsers,
  };
  const token = ScriptApp.getOAuthToken();
  const params = {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(options),
    contentType: 'application/json',
    muteHttpExceptions: true,
  };
  const res = UrlFetchApp.fetch(urlGetDataScript, params);
  const resDataArr = res.getContentText().split(',');
  const resData = resDataArr.reduce((acc, _, index) => {
    if (index % 6 === 0) {
      acc.push(resDataArr.slice(index, index + 6));
    }
    return acc;
  }, []);
  const data = resData.filter((account) => account[1] === user)[0];
  return { result: 'success', func, data };
}

function writeNewArrivalZIP(data, files) {
  const optionsUpload = {
    folderUpload: folderUploadArrivalZIP,
    files: files,
  };
  const resUpload = writeToDrive(optionsUpload);
  const urls = resUpload.getContentText();

  const user = Session.getActiveUser().getEmail().toLowerCase();
  let newArr = [];
  const today = getToday();
  const entrance = `${today}:: ${user}:: Поступление: ${data[0]} -> ${data[4]}`;
  if (data[6]) {
    const snArr = data[6]
      .split(/[ .:;?!~,"&|()<>{}\[\]\r\n/\\]+/)
      .filter((item) => item !== '');
    newArr = snArr.map((item) => [
      data[3],
      data[1],
      data[2],
      item,
      data[4],
      entrance,
      data[7],
      urls,
    ]);
  } else {
    let k = 1;
    while (k <= data[5]) {
      newArr.push([
        data[3],
        data[1],
        data[2],
        '',
        data[4],
        entrance,
        data[7],
        urls,
      ]);
      k++;
    }
  }
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameDataZIP);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const options = {
    log: false,
    ssID: ssIDWarehouse,
    ssSheet: ssNameDataZIP,
    startRow: getLastRow + 1,
    startColumn: 1,
    rows: newArr.length,
    columns: 8,
    data: newArr,
  };

  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func: 'newArrivalZIP', data };
  }
  return { result: 'error', func: 'newArrivalZIP', data };
}

function writeNewArrivalPodmena(data, files) {
  const optionsUpload = {
    folderUpload: folderUploadArrivalPodmena,
    files: files,
  };
  const resUpload = writeToDrive(optionsUpload);
  const urls = resUpload.getContentText();

  const user = Session.getActiveUser().getEmail().toLowerCase();
  let newArr = [];
  const today = getToday();
  const entrance = `${today}:: ${user}:: Поступление: ${data[0]} -> ${data[3]}`;
  const snArr = data[4]
    .split(/[ .:;?!~,"&|()<>{}\[\]\r\n/\\]+/)
    .filter((item) => item !== '');
  newArr = snArr.map((item) => [
    data[2],
    data[1],
    item,
    data[3],
    entrance,
    data[5],
    urls,
  ]);
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameDataPodmena);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const options = {
    log: false,
    ssID: ssIDWarehouse,
    ssSheet: ssNameDataPodmena,
    startRow: getLastRow + 1,
    startColumn: 1,
    rows: newArr.length,
    columns: 7,
    data: newArr,
  };

  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func: 'newArrivalPodmena', data };
  }
  return { result: 'error', func: 'newArrivalPodmena', data };
}

function writeNewMovingZIP(data) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameDataZIP);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const user = Session.getActiveUser().getEmail().toLowerCase();
  let numRows = [];
  const today = getToday();
  const entrance = `${today}:: ${user}:: Перемещение: ${data[0]} -> ${data[4]}`;
  if (data[3]) {
    const snArr = data[3]
      .split(/[ .:;?!~,"&|()<>{}\[\]\r\n/\\]+/)
      .filter((item) => item !== '');
    const rangesUnit = wsData
      .getRange(1, 1, getLastRow, 1)
      .createTextFinder(data[1])
      .findAll()
      .map((item) => item.getRow());
    const rangesWarehouse = wsData
      .getRange(1, 5, getLastRow, 1)
      .createTextFinder(data[0])
      .findAll()
      .map((item) => item.getRow());
    const numUW = rangesUnit.filter((value) => rangesWarehouse.includes(value));
    const numSN = [].concat(
      ...snArr.map((itemSN) =>
        wsData
          .getRange(1, 4, getLastRow, 1)
          .createTextFinder(itemSN)
          .findAll()
          .map((item) => item.getRow())
      )
    );
    numRows = numSN.filter((value) => numUW.includes(value));
  } else {
    const rangesUnit = wsData
      .getRange(1, 1, getLastRow, 1)
      .createTextFinder(data[1])
      .findAll()
      .map((item) => item.getRowIndex());
    const rangesWarehouse = wsData
      .getRange(1, 5, getLastRow, 1)
      .createTextFinder(data[0])
      .findAll()
      .map((item) => item.getRowIndex());
    numRows = rangesUnit
      .filter((value) => rangesWarehouse.includes(value))
      .slice(0, data[2]);
  }

  if (!numRows.length) return;
  let result = [];
  numRows.map((item) => {
    const [
      unit,
      classifier,
      model,
      serial,
      warehouse,
      arrivalDate,
      arrivalComment,
      arrivalActs,
      moving,
      comments,
    ] = wsData.getRange(item, 1, 1, 10).getValues()[0];
    const newMoving = moving ? `${moving}\n${entrance}` : entrance;
    const newComments = moving ? `${comments}\n${data[5]}` : data[5];
    const newArr = [
      unit,
      classifier,
      model,
      serial,
      data[4],
      arrivalDate,
      arrivalComment,
      arrivalActs,
      newMoving,
      newComments,
    ];
    const options = {
      log: false,
      ssID: ssIDWarehouse,
      ssSheet: ssNameDataZIP,
      startRow: item,
      startColumn: 1,
      rows: 1,
      columns: 10,
      data: [newArr],
    };
    const res = writeToSheet(options);
    result = [res, ...result];
  });

  if (!result.includes(200)) {
    return { result: 'error', func: 'newMovingZIP', data };
  }
  return { result: 'success', func: 'newMovingZIP', data };
}

function writeNewMovingPodmena(data) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameDataPodmena);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const user = Session.getActiveUser().getEmail().toLowerCase();
  let numRows = [];
  const today = getToday();
  const entrance = `${today}:: ${user}:: Перемещение: ${data[0]} -> ${data[3]}`;
  if (data[2]) {
    const snArr = data[2]
      .split(/[ .:;?!~,"&|()<>{}\[\]\r\n/\\]+/)
      .filter((item) => item !== '');
    const rangesUnit = wsData
      .getRange(1, 1, getLastRow, 1)
      .createTextFinder(data[1])
      .findAll()
      .map((item) => item.getRow());
    const rangesWarehouse = wsData
      .getRange(1, 4, getLastRow, 1)
      .createTextFinder(data[0])
      .findAll()
      .map((item) => item.getRow());
    const numUW = rangesUnit.filter((value) => rangesWarehouse.includes(value));
    const numSN = [].concat(
      ...snArr.map((itemSN) =>
        wsData
          .getRange(1, 3, getLastRow, 1)
          .createTextFinder(itemSN)
          .findAll()
          .map((item) => item.getRow())
      )
    );
    numRows = numSN.filter((value) => numUW.includes(value));
  }
  if (!numRows.length) return;
  let result = [];
  numRows.map((item) => {
    const [
      unit,
      classifier,
      serial,
      warehouse,
      arrivalDate,
      arrivalComment,
      arrivalActs,
      moving,
      comments,
    ] = wsData.getRange(item, 1, 1, 9).getValues()[0];
    const newMoving = moving ? `${moving}\n${entrance}` : entrance;
    const newComments = moving ? `${comments}\n${data[4]}` : data[4];
    const newArr = [
      unit,
      classifier,
      serial,
      data[3],
      arrivalDate,
      arrivalComment,
      arrivalActs,
      newMoving,
      newComments,
    ];
    const options = {
      log: false,
      ssID: ssIDWarehouse,
      ssSheet: ssNameDataPodmena,
      startRow: item,
      startColumn: 1,
      rows: 1,
      columns: 9,
      data: [newArr],
    };
    const res = writeToSheet(options);
    result = [res, ...result];
  });

  if (!result.includes(200)) {
    return { result: 'error', func: 'newMovingPodmena', data };
  }
  return { result: 'success', func: 'newMovingPodmena', data };
}

function getСlassifier(func, type) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameClassifier);
  const data = wsData
    .getRange(2, 1, wsData.getLastRow() - 1, 2)
    .getValues()
    .filter((item) => item[1] === 'Активный')
    .map((item) => item[0])
    .sort();
  return { result: 'success', func, data, type };
}

function getModel(func, type) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameModel);
  const data = wsData
    .getRange(2, 1, wsData.getLastRow() - 1, 3)
    .getValues()
    .filter((item) => item[2] === 'Активный')
    .map((item) => [item[0], item[1]])
    .sort();
  return { result: 'success', func, data, type };
}

function getClients(func, type) {
  const options = {
    ssID: ssIDAdmin,
    sheet: ssNameClients,
  };
  const token = ScriptApp.getOAuthToken();
  const params = {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(options),
    contentType: 'application/json',
    muteHttpExceptions: true,
  };
  const res = UrlFetchApp.fetch(urlGetDataScript, params);
  const resDataArr = res.getContentText().split(',');
  const resData = resDataArr.reduce((acc, _, index) => {
    if (index % 3 === 0) {
      acc.push(resDataArr.slice(index, index + 3));
    }
    return acc;
  }, []);
  const data = resData
    .filter((item) => item[2] === 'Активный')
    .map((item) => item[1])
    .sort();
  return { result: 'success', func, data, type };
}

function getWarehouses(func, type) {
  const options = {
    ssID: ssIDAdmin,
    sheet: ssNameUsers,
  };
  const token = ScriptApp.getOAuthToken();
  const params = {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(options),
    contentType: 'application/json',
    muteHttpExceptions: true,
  };
  const res = UrlFetchApp.fetch(urlGetDataScript, params);
  const resDataArr = res.getContentText().split(',');
  const resData = resDataArr.reduce((acc, _, index) => {
    if (index % 6 === 0) {
      acc.push(resDataArr.slice(index, index + 6));
    }
    return acc;
  }, []);
  const data = resData
    .filter((item) => item[5] === 'Активный')
    .map((item) => item[0])
    .sort();
  return { result: 'success', func, data, type };
}

function getZIP(func, type) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameClassifierZIP);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const data = wsData
    .getRange(2, 1, getLastRow - 1, 4)
    .getValues()
    .filter((item) => item[3] === 'Активный')
    .map((item) => [item[0], item[1], item[2]])
    .sort();
  return { result: 'success', func, data, type };
}

function getZIPonWarehouses(func, type) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameDataZIP);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const data = wsData
    .getRange(2, 1, getLastRow - 1, 5)
    .getValues()
    .sort();
  return { result: 'success', func, data, type };
}

function getZIPForPrint(func, type) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameDataZIP);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const getLastColumn = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.NEXT)
    .getColumn();
  const today = getTodayDDMMYYY();
  const data = wsData
    .getRange(2, 1, getLastRow - 1, getLastColumn)
    .getValues()
    .filter((item) => item[8].includes(today))
    .sort();
  return { result: 'success', func, data, type };
}

function getPodmenaonWarehouses(func, type) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameDataPodmena);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const data = wsData
    .getRange(2, 1, getLastRow - 1, 4)
    .getValues()
    .sort();
  return { result: 'success', func, data, type };
}

function getPodmenaForPrint(func, type) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameDataPodmena);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const getLastColumn = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.NEXT)
    .getColumn();
  const today = getTodayDDMMYYY();
  const data = wsData
    .getRange(2, 1, getLastRow - 1, getLastColumn)
    .getValues()
    .filter((item) => item[7].includes(today))
    .sort();
  return { result: 'success', func, data, type };
}

function getPodmena(func, type) {
  const wsData = SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(
    ssNameClassifierPodmena
  );
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const data = wsData
    .getRange(2, 1, getLastRow - 1, 3)
    .getValues()
    .filter((item) => item[2] === 'Активный')
    .map((item) => [item[0], item[1]])
    .sort();
  return { result: 'success', func, data, type };
}

function getActNumber(func, type) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameData);
  const data = wsData.getRange(1, 2).getValue();
  return { result: 'success', func, data, type };
}

function addToAct(func, actNumber) {
  const options = {
    log: false,
    ssID: ssIDWarehouse,
    ssSheet: ssNameData,
    startRow: 1,
    startColumn: 2,
    rows: 1,
    columns: 1,
    data: [[actNumber]],
  };

  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func };
  }
  return { result: 'error', func };
}

function writeToDrive(optionsUpload) {
  const token = ScriptApp.getOAuthToken();
  const params = {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(optionsUpload),
    contentType: 'application/json',
    muteHttpExceptions: true,
  };
  return UrlFetchApp.fetch(urlWriteToDriveScript, params);
}

function writeToSheet(options) {
  const user = Session.getActiveUser().getEmail().toLowerCase();
  const token = ScriptApp.getOAuthToken();
  const params = {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ ...options, user: user }),
    contentType: 'application/json',
    muteHttpExceptions: true,
  };
  return UrlFetchApp.fetch(urlWriteToSheetScript, params).getResponseCode();
}

function getToday() {
  const today = new Date();
  const dd = today.getDate() < 10 ? `0${today.getDate()}` : today.getDate();
  const mm =
    today.getMonth() + 1 < 10
      ? `0${today.getMonth() + 1}`
      : today.getMonth() + 1;
  const yyyy = today.getFullYear();
  const hh = today.getHours() < 10 ? `0${today.getHours()}` : today.getHours();
  const min =
    today.getMinutes() < 10 ? `0${today.getMinutes()}` : today.getMinutes();
  return `${dd}.${mm}.${yyyy} ${hh}:${min}`;
}

function getTodayDDMMYYY() {
  const today = new Date();
  const dd = today.getDate() < 10 ? `0${today.getDate()}` : today.getDate();
  const mm =
    today.getMonth() + 1 < 10
      ? `0${today.getMonth() + 1}`
      : today.getMonth() + 1;
  const yyyy = today.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}
