/**
 * 蛋白質日記 - Google Apps Script 後端（v4-sugar-update）
 */

var FOODS_SHEET_NAME = 'Foods';
var LOGS_SHEET_NAME = 'Logs';
var FOODS_HEADERS = ['id', 'name', 'base', 'protein100', 'fat100', 'sugar100', 'cal100'];
var LOGS_HEADERS = ['id', 'person', 'date', 'time', 'type', 'foodId', 'foodName', 'grams', 'protein', 'fat', 'sugar', 'cal'];
var APP_BACKEND_VERSION = 'v4-sugar-update';

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'getData') {
    return jsonResponse(getData());
  }
  return jsonResponse({ error: 'unknown action: ' + action });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'invalid JSON body' });
  }
  var action = body.action;
  var payload = body.payload || {};
  var result;
  try {
    switch (action) {
      case 'addFood': result = addFood(payload); break;
      case 'updateFood': result = updateFood(payload); break;
      case 'deleteFood': result = deleteFood(payload); break;
      case 'addLog': result = addLog(payload); break;
      case 'deleteLog': result = deleteLog(payload); break;
      default: result = { error: 'unknown action: ' + action };
    }
  } catch (err) {
    result = { error: String(err) };
  }
  return jsonResponse(result);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- sheet + header helpers ----------

function getFoodsSheet() { return getSheet(FOODS_SHEET_NAME, FOODS_HEADERS); }
function getLogsSheet() { return getSheet(LOGS_SHEET_NAME, LOGS_HEADERS); }

function getSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }
  ensureHeaders(sheet, headers);
  return sheet;
}

function ensureHeaders(sheet, headers) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (!headerRow[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  headers.forEach(function (h) {
    if (headerRow.indexOf(h) === -1) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(h);
      headerRow.push(h);
    }
  });
}

function headerIndexMap(sheet) {
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headerRow.forEach(function (h, i) { if (h) map[h] = i; });
  return map;
}

function appendRowByHeader(sheet, headers, valuesObj) {
  var map = headerIndexMap(sheet);
  var lastCol = Math.max(sheet.getLastColumn(), headers.length);
  var row = new Array(lastCol).fill('');
  headers.forEach(function (h) {
    if (map.hasOwnProperty(h) && valuesObj.hasOwnProperty(h)) {
      row[map[h]] = valuesObj[h];
    }
  });
  sheet.appendRow(row);
}

// ---------- read ----------

function getData() {
  return { foods: readFoods(), logs: readLogs() };
}

function readFoods() {
  var sheet = getFoodsSheet();
  var map = headerIndexMap(sheet);
  var rows = sheet.getDataRange().getValues();
  var foods = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[map['id']]) continue;
    var sugarVal = map['sugar100'] !== undefined ? r[map['sugar100']] : (map['carb100'] !== undefined ? r[map['carb100']] : 0);
    foods.push({
      id: r[map['id']],
      name: r[map['name']],
      base: Number(r[map['base']]) || 100,
      protein100: Number(r[map['protein100']]) || 0,
      fat100: Number(r[map['fat100']]) || 0,
      sugar100: Number(sugarVal) || 0,
      cal100: Number(r[map['cal100']]) || 0
    });
  }
  return foods;
}

function readLogs() {
  var sheet = getLogsSheet();
  var map = headerIndexMap(sheet);
  var rows = sheet.getDataRange().getValues();
  var logs = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[map['id']]) continue;
    var gramsRaw = r[map['grams']];
    var sugarVal = map['sugar'] !== undefined ? r[map['sugar']] : (map['carb'] !== undefined ? r[map['carb']] : 0);
    logs.push({
      id: r[map['id']],
      person: r[map['person']] || 'A',
      date: r[map['date']],
      time: r[map['time']],
      type: r[map['type']],
      foodId: r[map['foodId']] || null,
      foodName: r[map['foodName']] || '',
      grams: (gramsRaw === '' || gramsRaw === undefined) ? null : Number(gramsRaw),
      protein: Number(r[map['protein']]) || 0,
      fat: Number(r[map['fat']]) || 0,
      sugar: Number(sugarVal) || 0,
      cal: Number(r[map['cal']]) || 0
    });
  }
  return logs;
}

// ---------- write ----------

function addFood(payload) {
  var sheet = getFoodsSheet();
  var id = payload.id || ('f_' + new Date().getTime());
  var sugarVal = payload.sugar100 !== undefined ? payload.sugar100 : payload.carb100;
  appendRowByHeader(sheet, FOODS_HEADERS, {
    id: id,
    name: payload.name || '',
    base: Number(payload.base) || 100,
    protein100: Number(payload.protein100) || 0,
    fat100: Number(payload.fat100) || 0,
    sugar100: Number(sugarVal) || 0,
    cal100: Number(payload.cal100) || 0
  });
  SpreadsheetApp.flush();
  return { success: true, id: id };
}

function updateFood(payload) {
  var sheet = getFoodsSheet();
  var map = headerIndexMap(sheet);
  var data = sheet.getDataRange().getValues();
  var sugarVal = payload.sugar100 !== undefined ? payload.sugar100 : payload.carb100;
  var sugarCol = map['sugar100'] !== undefined ? map['sugar100'] : map['carb100'];

  for (var i = 1; i < data.length; i++) {
    if (data[i][map['id']] === payload.id) {
      var rowNum = i + 1;
      sheet.getRange(rowNum, map['name'] + 1).setValue(payload.name || '');
      sheet.getRange(rowNum, map['base'] + 1).setValue(Number(payload.base) || 100);
      sheet.getRange(rowNum, map['protein100'] + 1).setValue(Number(payload.protein100) || 0);
      sheet.getRange(rowNum, map['fat100'] + 1).setValue(Number(payload.fat100) || 0);
      if (sugarCol !== undefined) {
        sheet.getRange(rowNum, sugarCol + 1).setValue(Number(sugarVal) || 0);
      }
      sheet.getRange(rowNum, map['cal100'] + 1).setValue(Number(payload.cal100) || 0);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { error: 'food not found' };
}

function deleteFood(payload) {
  var sheet = getFoodsSheet();
  var map = headerIndexMap(sheet);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][map['id']] === payload.id) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { error: 'food not found' };
}

function addLog(payload) {
  var sheet = getLogsSheet();
  var id = 'l_' + new Date().getTime();
  var sugarVal = payload.sugar !== undefined ? payload.sugar : payload.carb;
  appendRowByHeader(sheet, LOGS_HEADERS, {
    id: id,
    person: payload.person || 'A',
    date: payload.date,
    time: payload.time,
    type: payload.type,
    foodId: payload.foodId || '',
    foodName: payload.foodName || '',
    grams: payload.grams == null ? '' : payload.grams,
    protein: Number(payload.protein) || 0,
    fat: Number(payload.fat) || 0,
    sugar: Number(sugarVal) || 0,
    cal: Number(payload.cal) || 0
  });
  SpreadsheetApp.flush();
  return { success: true, id: id };
}

function deleteLog(payload) {
  var sheet = getLogsSheet();
  var map = headerIndexMap(sheet);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][map['id']] === payload.id) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { error: 'log not found' };
}