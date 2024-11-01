const urlWriteToSheetScript =
  'https://script.google.com/macros/s/AKfycbzyPNucegxoejhJNI2kJUjsfvqjGaBVP3bDafh3yUw/dev';
const ssID = '11x7Mh-OCuuC1WOC26WEDoVdx5jpZqTCCQyEvJks1OrQ';
const ssIDWarehouse = '1QHenRbqyifedRX-XN4uLMOnOiBFm5XYMzzjuksL2yyQ';
const ssNameUsers = 'Пользователи';
const ssNamePosts = 'Должности';
const ssNameRoles = 'Роли';
const ssNameClients = 'Клиенты';
const ssNameClassifier = 'Классификатор';
const ssNameModel = 'Модели';
const ssNameZIP = 'Классификатор ЗИП';
const ssNamePodmena = 'Классификатор подмены';
const ssNameObjects = 'Объекты';
const ssNameRegions = 'Регионы';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('administration');
}

function checkAccess(func) {
  const user = Session.getActiveUser().getEmail().toLowerCase();
  const wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssNameUsers);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const dataAccess = wsData
    .getRange(2, 1, getLastRow - 1, wsData.getLastColumn())
    .getValues();
  const data = dataAccess.filter((account) => account[1] === user)[0];
  return { result: 'success', func, data };
}

function writeNewСlassifier(data) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameClassifier);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const options = {
    log: true,
    ssID: ssIDWarehouse,
    ssSheet: ssNameClassifier,
    startRow: getLastRow + 1,
    startColumn: 1,
    rows: 1,
    columns: 2,
    data: data,
  };
  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func: 'newСlassifier', data };
  }
  return { result: 'error', func: 'newСlassifier', data };
}

function writeNewModel(data) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameModel);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const options = {
    log: true,
    ssID: ssIDWarehouse,
    ssSheet: ssNameModel,
    startRow: getLastRow + 1,
    startColumn: 1,
    rows: 1,
    columns: 3,
    data: data,
  };
  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func: 'newModel', data };
  }
  return { result: 'error', func: 'newModel', data };
}

function writeNewZIP(data) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNameZIP);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const options = {
    log: true,
    ssID: ssIDWarehouse,
    ssSheet: ssNameZIP,
    startRow: getLastRow + 1,
    startColumn: 1,
    rows: 1,
    columns: 4,
    data: data,
  };
  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func: 'newZIP', data };
  }
  return { result: 'error', func: 'newZIP', data };
}

function writeNewPodmena(data) {
  const wsData =
    SpreadsheetApp.openById(ssIDWarehouse).getSheetByName(ssNamePodmena);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const options = {
    log: true,
    ssID: ssIDWarehouse,
    ssSheet: ssNamePodmena,
    startRow: getLastRow + 1,
    startColumn: 1,
    rows: 1,
    columns: 3,
    data: data,
  };
  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func: 'newPodmena', data };
  }
  return { result: 'error', func: 'newPodmena', data };
}

function writeNewEmployee(data) {
  const wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssNameUsers);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const options = {
    log: true,
    ssID: ssID,
    ssSheet: ssNameUsers,
    startRow: getLastRow + 1,
    startColumn: 1,
    rows: 1,
    columns: 6,
    data: data,
  };
  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func: 'newEmployee', data };
  }
  return { result: 'error', func: 'newEmployee', data };
}

function writeNewPost(data) {
  const wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssNamePosts);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getLastRow();
  const options = {
    log: true,
    ssID: ssID,
    ssSheet: ssNamePosts,
    startRow: getLastRow + 1,
    startColumn: 1,
    rows: 1,
    columns: 2,
    data: data,
  };
  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func: 'newPost', data };
  }
  return { result: 'error', func: 'newPost', data };
}

function writeNewRole(data) {
  const wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssNameRoles);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const options = {
    log: true,
    ssID: ssID,
    ssSheet: ssNameRoles,
    startRow: getLastRow + 1,
    startColumn: 1,
    rows: 1,
    columns: 2,
    data: data,
  };
  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func: 'newRole', data };
  }
  return { result: 'error', func: 'newRole', data };
}

function writeNewClient(data) {
  const wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssNameClients);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const options = {
    log: true,
    ssID: ssID,
    ssSheet: ssNameClients,
    startRow: getLastRow + 1,
    startColumn: 1,
    rows: 1,
    columns: 3,
    data: data,
  };
  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func: 'newClient', data };
  }
  return { result: 'error', func: 'newClient', data };
}

function writeNewObject(data) {
  const wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssNameObjects);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const options = {
    log: true,
    ssID: ssID,
    ssSheet: ssNameObjects,
    startRow: getLastRow + 1,
    startColumn: 1,
    rows: 1,
    columns: 8,
    data: data,
  };
  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func: 'newObject', data };
  }
  return { result: 'error', func: 'newObject', data };
}

function writeNewRegion(data) {
  const wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssNameRegions);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const options = {
    log: true,
    ssID: ssID,
    ssSheet: ssNameRegions,
    startRow: getLastRow + 1,
    startColumn: 1,
    rows: 1,
    columns: 2,
    data: data,
  };
  const res = writeToSheet(options);
  if (res === 200) {
    return { result: 'success', func: 'newRegion', data };
  }
  return { result: 'error', func: 'newRegion', data };
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

function getPosts(func, type) {
  const wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssNamePosts);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const data = wsData
    .getRange(2, 1, getLastRow - 1, 2)
    .getValues()
    .filter((item) => item[1] === 'Активный')
    .map((item) => item[0])
    .sort();
  return { result: 'success', func, data, type };
}

function getRoles(func, type) {
  const wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssNameRoles);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const data = wsData
    .getRange(2, 1, getLastRow - 1, 2)
    .getValues()
    .filter((item) => item[1] === 'Активный')
    .map((item) => item[0])
    .sort();
  return { result: 'success', func, data, type };
}

function getClients(func, type) {
  const wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssNameClients);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const data = wsData
    .getRange(2, 1, getLastRow - 1, 3)
    .getValues()
    .filter((item) => item[2] === 'Активный')
    .map((item) => item[1])
    .sort();
  return { result: 'success', func, data, type };
}

function getRegions(func, type) {
  const wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssNameRegions);
  const getLastRow = wsData
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const data = wsData
    .getRange(2, 1, getLastRow - 1, 2)
    .getValues()
    .filter((item) => item[1] === 'Активный')
    .map((item) => item[0])
    .sort();
  return { result: 'success', func, data, type };
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
