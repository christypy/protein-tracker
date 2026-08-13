/**
 * 蛋白質日記 - Google Apps Script 後端（v2）
 *
 * 這一版做了兩件事：
 * 1. 讀寫欄位改成「用欄位名稱對照」，不再用固定的欄位順序（index）。
 *    這樣即使你的試算表是舊版本建立的、欄位數量跟現在的程式碼對不上，
 *    程式也會在每次執行時自動把缺少的欄位（例如「熱量」cal100 / cal）補到表格最後一欄，
 *    不會動到既有資料。這就是為什麼舊試算表可能「看起來」沒有熱量欄位 ——
 *    因為那個試算表是在程式碼還沒加上熱量欄位時就建立的，之後程式改版並不會自動幫舊表加欄位。
 * 2. Logs 多了一個 person 欄位，用來標記這筆紀錄屬於「身份 A」還是「身份 B」，
 *    這樣兩個人共用同一份試算表時，彼此的今日紀錄跟達成率才不會混在一起。
 *
 * 部署方式跟之前一樣：
 * 1. 開一份新的（或你原本那份）Google 試算表
 * 2. 上方選單「擴充功能」→「Apps Script」
 * 3. 刪除預設內容，貼上這整份程式碼，按儲存
 * 4. 右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    - 執行身分：我
 *    - 誰可以存取：任何人
 * 5. 複製部署後產生的網址，貼到「蛋白質日記」App 的設定頁面
 *
 * 如果是「更新既有部署」（沿用同一支網址），選「管理部署作業」→ 編輯 → 版本選「新版本」→ 部署，
 * 不需要重新複製網址。
 */

var FOODS_SHEET_NAME = 'Foods';
var LOGS_SHEET_NAME = 'Logs';
var FOODS_HEADERS = ['id', 'name', 'base', 'protein100', 'fat100', 'carb100', 'cal100'];
var LOGS_HEADERS = ['id', 'person', 'date', 'time', 'type', 'foodId', 'foodName', 'grams', 'protein', 'fat', 'carb', 'cal'];

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

// 確保表頭包含所有必要欄位；缺少的欄位（例如舊表沒有的「熱量」cal100/cal，
// 或新加入的「person」）會自動補到最後一欄，不影響既有資料。
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

// 回傳 {欄位名稱: 0-based 欄位索引}
function headerIndexMap(sheet) {
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headerRow.forEach(function (h, i) { if (h) map[h] = i; });
  return map;
}

// 依欄位名稱把一筆資料組成正確順序的陣列後 appendRow
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
    foods.push({
      id: r[map['id']],
      name: r[map['name']],
      base: Number(r[map['base']]) || 100,
      protein100: Number(r[map['protein100']]) || 0,
      fat100: Number(r[map['fat100']]) || 0,
      carb100: Number(r[map['carb100']]) || 0,
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
      carb: Number(r[map['carb']]) || 0,
      cal: Number(r[map['cal']]) || 0
    });
  }
  return logs;
}

// ---------- write ----------

function addFood(payload) {
  var sheet = getFoodsSheet();
  var id = 'f_' + new Date().getTime();
  appendRowByHeader(sheet, FOODS_HEADERS, {
    id: id,
    name: payload.name,
    base: payload.base || 100,
    protein100: payload.protein100 || 0,
    fat100: payload.fat100 || 0,
    carb100: payload.carb100 || 0,
    cal100: payload.cal100 || 0
  });
  return { id: id };
}

function updateFood(payload) {
  var sheet = getFoodsSheet();
  var map = headerIndexMap(sheet);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][map['id']] === payload.id) {
      var rowNum = i + 1;
      sheet.getRange(rowNum, map['name'] + 1).setValue(payload.name);
      sheet.getRange(rowNum, map['base'] + 1).setValue(payload.base || 100);
      sheet.getRange(rowNum, map['protein100'] + 1).setValue(payload.protein100 || 0);
      sheet.getRange(rowNum, map['fat100'] + 1).setValue(payload.fat100 || 0);
      sheet.getRange(rowNum, map['carb100'] + 1).setValue(payload.carb100 || 0);
      sheet.getRange(rowNum, map['cal100'] + 1).setValue(payload.cal100 || 0);
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
      return { success: true };
    }
  }
  return { error: 'food not found' };
}

function addLog(payload) {
  var sheet = getLogsSheet();
  var id = 'l_' + new Date().getTime();
  appendRowByHeader(sheet, LOGS_HEADERS, {
    id: id,
    person: payload.person || 'A',
    date: payload.date,
    time: payload.time,
    type: payload.type,
    foodId: payload.foodId || '',
    foodName: payload.foodName || '',
    grams: payload.grams == null ? '' : payload.grams,
    protein: payload.protein || 0,
    fat: payload.fat || 0,
    carb: payload.carb || 0,
    cal: payload.cal || 0
  });
  return { id: id };
}

function deleteLog(payload) {
  var sheet = getLogsSheet();
  var map = headerIndexMap(sheet);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][map['id']] === payload.id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'log not found' };
}
