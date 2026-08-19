/**
 * 蛋白質日記 - Google Apps Script 後端（v6-date-fix）
 *
 * 這版修正三個問題：
 * 1. 舊試算表如果欄位還是 carb100 / carb，過去的版本會「多新增一欄」sugar100 / sugar，
 *    導致資料分裂在兩個欄位，加總、顯示都會對不到正確的值（這就是糖量顯示異常的根本原因）。
 *    現在改成：偵測到舊欄位時會自動把資料「搬進」新欄位（或直接改名），並刪掉舊欄位。
 * 2. appendRowByHeader() 以前如果有欄位在表格裡找不到對應的表頭，會「默默把該筆資料丟掉」，
 *    完全不會報錯 —— 這也是熱量(cal)有時候沒有同步進試算表的原因。
 *    現在改成：找不到欄位就自動新增該欄，絕不會再默默漏資料。
 * 3.【本次新增】連上試算表之後「今日紀錄」整批消失、但重新整理成本機模式又看得到資料，
 *    根本原因是 Google 試算表會把 "2026-08-13" 這種字串自動判斷成日期物件；讀回網頁後
 *    變成帶時間的 ISO 字串，跟前端用來篩選「今天」的純日期字串永遠對不起來，紀錄因此
 *    像是憑空消失。現在做兩件事修好它：
 *      a) date / time 欄位整欄強制設成「純文字」格式，之後新寫入的資料不會再被誤判。
 *      b) 讀取時如果偵測到儲存格仍是日期物件（例如這次修復前就已經寫入的舊資料），
 *         會就地轉回 yyyy-MM-dd / HH:mm 純文字字串再回傳給前端，舊資料也一併修好，
 *         不需要另外手動搬移。
 */

var FOODS_SHEET_NAME = 'Foods';
var LOGS_SHEET_NAME = 'Logs';
var FOODS_HEADERS = ['id', 'name', 'base', 'unit', 'category', 'protein100', 'fat100', 'sugar100', 'cal100'];
var LOGS_HEADERS = ['id', 'person', 'date', 'time', 'type', 'foodId', 'foodName', 'grams', 'unit', 'protein', 'fat', 'sugar', 'cal', 'order'];
// 【v9 新增】Logs 新增了 order 欄位，用來記錄「今日紀錄」使用者手動拖移排序後的順序
// （單純一個數字，同一天、同一身份的紀錄依這個數字由小到大顯示）。舊試算表沒有這欄時
// ensureHeaders() 會自動補上；讀取舊資料時 order 欄位若是空的，前端會用時間排序當作預設值。
// 舊欄位名稱 -> 新欄位名稱。ensureHeaders() 會自動把舊欄位的資料合併進新欄位。
var HEADER_RENAME_MAP = { 'carb100': 'sugar100', 'carb': 'sugar' };
var APP_BACKEND_VERSION = 'v9-manual-order';
// 【v7 新增】Foods / Logs 都新增了 unit 欄位（例如 g、顆、盒、碗）。
// 舊試算表沒有這欄時，ensureHeaders() 會自動補上，不需要手動搬移；
// 讀取舊資料時 unit 欄位若是空的，一律視為 'g'，行為與升級前完全一致。
// 【v8 新增】Foods 新增了 category 欄位（食材類別，例如肉類、蔬菜、乳製品…），
// 用來讓「新增這一餐」時可以用類別快速篩選食材。舊試算表沒有這欄時 ensureHeaders()
// 會自動補上；讀取舊資料時 category 欄位若是空的，一律視為「未分類」。

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
      case 'updateLog': result = updateLog(payload); break;
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

// 這些欄位一定要用「純文字」格式儲存，否則 Google 試算表會自動把
// "2026-08-13" 這種字串認成日期物件，讀回來的時候前端拿字串比對
// (l.date === state.currentDate) 就永遠對不上，紀錄因此「連了試算表反而不見」。
var TEXT_FORMAT_COLUMNS = ['date', 'time'];

function getSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    forceTextFormatColumns(sheet, headers);
    return sheet;
  }
  ensureHeaders(sheet, headers);
  forceTextFormatColumns(sheet, headers);
  return sheet;
}

// 把 date / time 欄位整欄設成純文字格式，避免 Google 試算表自動轉型成日期。
function forceTextFormatColumns(sheet, headers) {
  var map = headerIndexMap(sheet);
  var maxRows = Math.max(sheet.getMaxRows(), 1000);
  TEXT_FORMAT_COLUMNS.forEach(function (h) {
    if (headers.indexOf(h) === -1) return; // 這張表沒有這個欄位
    if (!map.hasOwnProperty(h)) return;
    sheet.getRange(1, map[h] + 1, maxRows, 1).setNumberFormat('@');
  });
}

function ensureHeaders(sheet, headers) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (!headerRow[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  // Step 1：把舊欄位（例如 carb100 / carb）合併或改名成新欄位（sugar100 / sugar）
  Object.keys(HEADER_RENAME_MAP).forEach(function (oldName) {
    var newName = HEADER_RENAME_MAP[oldName];
    if (headers.indexOf(newName) === -1) return; // 這張表本來就不需要這個新欄位
    mergeLegacyColumn(sheet, oldName, newName);
  });

  // Step 2：重新讀一次目前欄位（上面可能已經改動過欄位數），再補上真的還缺少的欄位
  lastCol = sheet.getLastColumn();
  headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  headers.forEach(function (h) {
    if (headerRow.indexOf(h) === -1) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(h);
      headerRow.push(h);
    }
  });
}

/**
 * 把「舊欄位」(oldName) 的資料合併進「新欄位」(newName)：
 * - 兩欄都存在：新欄位是空白的儲存格，就用舊欄位的值補上，然後把舊欄位整欄刪除。
 * - 只有舊欄位存在：直接把該欄標題改成新名稱，資料原地保留（最安全、也不用搬資料）。
 * - 只有新欄位或兩者都沒有：不用處理。
 */
function mergeLegacyColumn(sheet, oldName, newName) {
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var oldIdx = headerRow.indexOf(oldName);
  var newIdx = headerRow.indexOf(newName);

  if (oldIdx === -1) return; // 沒有舊欄位，不用處理

  if (newIdx === -1) {
    sheet.getRange(1, oldIdx + 1).setValue(newName);
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var oldVals = sheet.getRange(2, oldIdx + 1, lastRow - 1, 1).getValues();
    var newVals = sheet.getRange(2, newIdx + 1, lastRow - 1, 1).getValues();
    var changed = false;
    for (var i = 0; i < newVals.length; i++) {
      var newEmpty = (newVals[i][0] === '' || newVals[i][0] === null || newVals[i][0] === undefined);
      var oldHasVal = !(oldVals[i][0] === '' || oldVals[i][0] === null || oldVals[i][0] === undefined);
      if (newEmpty && oldHasVal) {
        newVals[i][0] = oldVals[i][0];
        changed = true;
      }
    }
    if (changed) {
      sheet.getRange(2, newIdx + 1, newVals.length, 1).setValues(newVals);
    }
  }
  sheet.deleteColumn(oldIdx + 1);
}

function headerIndexMap(sheet) {
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headerRow.forEach(function (h, i) { if (h) map[h] = i; });
  return map;
}

// 寫入一列新資料。若某欄位在試算表裡還找不到對應欄位，會自動新增該欄再寫入，
// 不會再像以前一樣默默把資料丟掉（這是熱量偶爾沒同步進試算表的根本原因）。
function appendRowByHeader(sheet, headers, valuesObj) {
  var map = headerIndexMap(sheet);
  headers.forEach(function (h) {
    if (!map.hasOwnProperty(h) && valuesObj.hasOwnProperty(h)) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(h);
      map[h] = newCol - 1;
    }
  });
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
      unit: (map.hasOwnProperty('unit') && r[map['unit']]) ? String(r[map['unit']]) : 'g',
      category: (map.hasOwnProperty('category') && r[map['category']]) ? String(r[map['category']]) : '',
      protein100: Number(r[map['protein100']]) || 0,
      fat100: Number(r[map['fat100']]) || 0,
      sugar100: Number(r[map['sugar100']]) || 0,
      cal100: Number(r[map['cal100']]) || 0
    });
  }
  return foods;
}

// 如果儲存格已經被 Google 試算表誤判成日期/時間物件（例如修復這個 bug 之前
// 就已經寫入的舊資料），讀取時就地轉回純文字字串，前端比對日期才會正常。
function formatDateCell(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return v;
}
function formatTimeCell(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'HH:mm');
  }
  return v;
}

function readLogs() {
  var sheet = getLogsSheet();
  var map = headerIndexMap(sheet);
  var rows = sheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  var logs = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[map['id']]) continue;
    var gramsRaw = r[map['grams']];
    var orderRaw = map.hasOwnProperty('order') ? r[map['order']] : '';
    var logObj = {
      id: r[map['id']],
      person: r[map['person']] || 'A',
      date: formatDateCell(r[map['date']], tz),
      time: formatTimeCell(r[map['time']], tz),
      type: r[map['type']],
      foodId: r[map['foodId']] || null,
      foodName: r[map['foodName']] || '',
      grams: (gramsRaw === '' || gramsRaw === undefined) ? null : Number(gramsRaw),
      unit: (map.hasOwnProperty('unit') && r[map['unit']]) ? String(r[map['unit']]) : 'g',
      protein: Number(r[map['protein']]) || 0,
      fat: Number(r[map['fat']]) || 0,
      sugar: Number(r[map['sugar']]) || 0,
      cal: Number(r[map['cal']]) || 0
    };
    // order 欄位若是空的（例如舊資料、還沒被拖移過），刻意不帶這個屬性，
    // 讓前端自己用時間排序當作預設順序，而不是把每筆都塞一個 0 造成排序錯亂。
    if (orderRaw !== '' && orderRaw !== undefined && orderRaw !== null && !isNaN(Number(orderRaw))) {
      logObj.order = Number(orderRaw);
    }
    logs.push(logObj);
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
    unit: payload.unit ? String(payload.unit) : 'g',
    category: payload.category ? String(payload.category) : '',
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

  for (var i = 1; i < data.length; i++) {
    if (data[i][map['id']] === payload.id) {
      var rowNum = i + 1;
      sheet.getRange(rowNum, map['name'] + 1).setValue(payload.name || '');
      sheet.getRange(rowNum, map['base'] + 1).setValue(Number(payload.base) || 100);
      if (!map.hasOwnProperty('unit')) {
        var newUnitCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newUnitCol).setValue('unit');
        map['unit'] = newUnitCol - 1;
      }
      sheet.getRange(rowNum, map['unit'] + 1).setValue(payload.unit ? String(payload.unit) : 'g');
      if (!map.hasOwnProperty('category')) {
        var newCatCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCatCol).setValue('category');
        map['category'] = newCatCol - 1;
      }
      sheet.getRange(rowNum, map['category'] + 1).setValue(payload.category ? String(payload.category) : '');
      sheet.getRange(rowNum, map['protein100'] + 1).setValue(Number(payload.protein100) || 0);
      sheet.getRange(rowNum, map['fat100'] + 1).setValue(Number(payload.fat100) || 0);
      sheet.getRange(rowNum, map['sugar100'] + 1).setValue(Number(sugarVal) || 0);
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
    unit: payload.unit ? String(payload.unit) : 'g',
    protein: Number(payload.protein) || 0,
    fat: Number(payload.fat) || 0,
    sugar: Number(sugarVal) || 0,
    cal: Number(payload.cal) || 0,
    order: payload.order != null ? Number(payload.order) : ''
  });
  SpreadsheetApp.flush();
  return { success: true, id: id };
}

// 局部更新一筆已存在的紀錄。目前主要用途是「今日紀錄」手動拖移排序後同步
// order 欄位，但也支援一併更新其他欄位（例如未來若要編輯紀錄本身的數值），
// 沒有帶到的欄位維持原值不動。
function updateLog(payload) {
  if (!payload || !payload.id) return { error: 'missing id' };
  var sheet = getLogsSheet();
  var map = headerIndexMap(sheet);
  var data = sheet.getDataRange().getValues();
  var sugarVal = payload.sugar !== undefined ? payload.sugar : payload.carb;

  for (var i = 1; i < data.length; i++) {
    if (data[i][map['id']] === payload.id) {
      var rowNum = i + 1;
      if (payload.order !== undefined) {
        if (!map.hasOwnProperty('order')) {
          var newOrderCol = sheet.getLastColumn() + 1;
          sheet.getRange(1, newOrderCol).setValue('order');
          map['order'] = newOrderCol - 1;
        }
        sheet.getRange(rowNum, map['order'] + 1).setValue(Number(payload.order) || 0);
      }
      if (payload.foodName !== undefined) sheet.getRange(rowNum, map['foodName'] + 1).setValue(payload.foodName);
      if (payload.grams !== undefined) sheet.getRange(rowNum, map['grams'] + 1).setValue(payload.grams == null ? '' : payload.grams);
      if (payload.protein !== undefined) sheet.getRange(rowNum, map['protein'] + 1).setValue(Number(payload.protein) || 0);
      if (payload.fat !== undefined) sheet.getRange(rowNum, map['fat'] + 1).setValue(Number(payload.fat) || 0);
      if (sugarVal !== undefined) sheet.getRange(rowNum, map['sugar'] + 1).setValue(Number(sugarVal) || 0);
      if (payload.cal !== undefined) sheet.getRange(rowNum, map['cal'] + 1).setValue(Number(payload.cal) || 0);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { error: 'log not found' };
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

/**
 * 手動修復小工具：如果你想在不呼叫 API 的情況下，馬上把 Foods / Logs 分頁的
 * carb100 / carb 欄位合併成 sugar100 / sugar，可以在 Apps Script 編輯器裡
 * 選這個函式，按「執行」跑一次。只會動欄位名稱與搬資料，不會刪除任何一列。
 */
function migrateHeaders() {
  getFoodsSheet();
  getLogsSheet();
  return 'done';
}