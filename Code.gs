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
var FOODS_HEADERS = ['id', 'name', 'base', 'unit', 'servings', 'category', 'categories', 'outOfStock', 'protein100', 'fat100', 'sugar100', 'cal100'];
var LOGS_HEADERS = ['id', 'person', 'date', 'time', 'type', 'foodId', 'foodName', 'grams', 'unit', 'protein', 'fat', 'sugar', 'cal', 'order', 'groupId'];
// 【v9 新增】Logs 新增了 order 欄位，用來記錄「今日紀錄」使用者手動拖移排序後的順序
// （單純一個數字，同一天、同一身份的紀錄依這個數字由小到大顯示）。舊試算表沒有這欄時
// ensureHeaders() 會自動補上；讀取舊資料時 order 欄位若是空的，前端會用時間排序當作預設值。
// 【v12 新增】Logs 新增了 groupId 欄位：「新增這一餐」一次選好幾樣食材、按「全部加入紀錄」
// 送出時，同一批送出的紀錄會共用同一個 groupId，前端「今日紀錄」會把同一個 groupId 的
// 紀錄自動包成一張「這一餐」卡片，顯示這一餐吃了哪些東西、加總多少蛋白質/熱量等。
// 舊試算表沒有這欄時 ensureHeaders() 會自動補上；讀取舊資料時 groupId 是空的就當作沒有
// 分組（維持原本一筆一筆顯示的行為）。
// 舊欄位名稱 -> 新欄位名稱。ensureHeaders() 會自動把舊欄位的資料合併進新欄位。
var HEADER_RENAME_MAP = { 'carb100': 'sugar100', 'carb': 'sugar' };
var APP_BACKEND_VERSION = 'v14-cross-device-delete-fix';
// 【v14 新增】修「重新整理／立即重新同步會讓已刪除的食材、紀錄又跑出來，
// 而且在試算表裡重複顯示」這個問題。根本原因有兩個，這版一次修掉：
//
// 1) 以前「這筆是不是已經被刪除」只記在刪除當下那一台裝置的本機（tombstone），
//    試算表本身完全不知道。只要有任何一台裝置、或同一台裝置沒關掉的舊分頁，
//    本機還留著早就被刪掉的食材／紀錄，它一重新整理或按「立即重新同步」，
//    就會把「試算表上找不到」誤判成「我這筆還沒同步上去」，自動把它當成
//    新資料補傳（addFood／addLog）回試算表——刪掉的東西因此陰魂不散地
//    又跑回來。
//    修法：新增一個 DeletedIds 分頁，deleteFood／deleteLog 時把 id 一併記錄
//    在這裡，getData() 回傳時把這份清單也一起給前端；前端不論用哪一台裝置
//    連上試算表，都會把伺服器端的刪除紀錄併進自己的本機 tombstone，之後
//    自然不會再被誤判成「待補傳」。
// 2) addFood／addLog 以前是「無條件 appendRow」，就算真的被誤判補傳了同一筆
//    （同一個 id）也只會傻傻地在試算表多長一列出來，這就是你會在試算表
//    看到兩列一模一樣資料的原因。
//    修法：改成 upsert——寫入前先檢查試算表裡是不是已經有相同 id 的列，
//    有的話直接更新那一列，不會再新增出重複列；沒有才新增一列。
//    如此一來，就算前面第 1 點的情境仍然發生（例如舊版前端快取還沒更新），
//    也不會再產生重複資料，最多只是把同一筆內容原地覆寫一次。
// 另外 readFoods()／readLogs() 也加了一層防呆：讀取時如果試算表裡剛好還
// 殘留著「同一個 id 出現兩次以上」的舊資料（例如這次修復之前就已經重複），
// 一律只取第一筆，畫面上不會再顯示成兩筆一樣的紀錄；並提供
// dedupeFoodsAndLogsSheets() 這個一次性工具，可以在 Apps Script 編輯器手動
// 執行，把試算表裡「已經存在」的重複列直接清乾淨。
// 【v13 新增】deleteFood() 修好「食材刪不掉」的問題：以前比對到第一筆符合的
// id 就刪除、馬上結束，如果同一個 id 因故在表格裡留下不只一列（例如同步時
// 按太快、或曾經同步失敗又重試），只會刪掉其中一列，另一列還在，重新整理
// 後那項食材就會像沒刪掉一樣又跑出來。現在改成用寬鬆比對（去頭尾空白、
// 轉成字串再比）把所有符合的列一次刪光。
// 另外新增 renameBrownRiceToGermRice() 一次性工具函式，可以在 Apps Script
// 編輯器手動執行，把「糙米飯」這個食材與過去所有「今日紀錄」裡吃過
// 「糙米飯」的歷史紀錄，一起改名成「黃金胚芽」，不需要透過刪除（也就不受
// 上面那個 bug 影響），也保留原本的食材 id 與既有設定。
// 【v7 新增】Foods / Logs 都新增了 unit 欄位（例如 g、顆、盒、碗）。
// 舊試算表沒有這欄時，ensureHeaders() 會自動補上，不需要手動搬移；
// 讀取舊資料時 unit 欄位若是空的，一律視為 'g'。
// 【v10 新增】前端現在強制單位一律是公克（g），並改用「基礎一份克數（base）＋總份數（servings）」
// 的資料模型：Foods 新增了 servings 欄位（這個食材整包/整份含幾份），跟 base（一份幾克）
// 搭配起來就能算出「整包總重量、總營養」= base × servings。舊試算表沒有這欄時 ensureHeaders()
// 會自動補上；讀取舊資料時 servings 欄位若是空的或不是有效數字，一律視為 1 份，行為與升級前一致。
// 【v8 新增】Foods 新增了 category 欄位；【v11】改為 categories 多選。
// 舊試算表仍保留 category 作為相容欄位，新增 categories 欄位後可讓同一食材同時屬於多個類別。
// categories 在試算表中以「、」分隔，例如「肉類、低脂、高蛋白」；讀取舊資料時會把 category
// 自動轉成單一元素的 categories 陣列。

// 【v14 新增】伺服器端的「刪除紀錄」分頁：不管哪一台裝置刪除食材或紀錄，
// 都會把 type（'food' 或 'log'）與 id 寫進這張表，getData() 會把這份清單
// 回傳給所有連上來的裝置，讓「這筆已經被刪掉了」變成試算表這邊共同的
// 事實，而不是只有刪除當下那一台裝置自己知道。
var DELETED_SHEET_NAME = 'DeletedIds';
var DELETED_HEADERS = ['type', 'id', 'deletedAt'];
// 這份清單只要「夠用」就好，不需要無限成長，超過上限時只保留最新的部分，
// 跟前端本機 tombstone 的 TOMBSTONE_LIMIT 概念一致。
var DELETED_SHEET_LIMIT = 2000;

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
function getDeletedSheet() { return getSheet(DELETED_SHEET_NAME, DELETED_HEADERS); }

// 把一批「已刪除」的 id 記進 DeletedIds 分頁。同一個 type+id 已經記錄過
// 就不會重複再寫一次，避免使用者反覆刪同一筆（理論上不會發生，但保險起見）
// 讓這張表無意義地一直變大。
function recordDeletion(type, ids) {
  if (!ids || !ids.length) return;
  var sheet = getDeletedSheet();
  var map = headerIndexMap(sheet);
  var lastRow = sheet.getLastRow();
  var existing = {};
  if (lastRow > 1) {
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][map['type']]) + '|' + String(data[i][map['id']]);
      existing[key] = true;
    }
  }
  var now = new Date();
  var rowsToAdd = [];
  ids.forEach(function (id) {
    var idStr = String(id == null ? '' : id).trim();
    if (!idStr) return;
    var key = type + '|' + idStr;
    if (existing[key]) return;
    existing[key] = true;
    rowsToAdd.push([type, idStr, now]);
  });
  if (rowsToAdd.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, DELETED_HEADERS.length).setValues(rowsToAdd);
    trimDeletedSheet(sheet);
  }
}

// 保持 DeletedIds 分頁不要無限成長：超過上限時，只保留最新刪除的那些列。
function trimDeletedSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var dataRowCount = lastRow - 1;
  if (dataRowCount <= DELETED_SHEET_LIMIT) return;
  var removeCount = dataRowCount - DELETED_SHEET_LIMIT;
  sheet.deleteRows(2, removeCount);
}

// 讀出某個 type（'food' 或 'log'）目前記錄到的所有已刪除 id，回傳給前端
// 跟本機 tombstone 合併。
function readDeletedIds(type) {
  var sheet = getDeletedSheet();
  var map = headerIndexMap(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var ids = [];
  var seen = {};
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][map['type']]) !== type) continue;
    var id = String(data[i][map['id']] == null ? '' : data[i][map['id']]).trim();
    if (!id || seen[id]) continue;
    seen[id] = true;
    ids.push(id);
  }
  return ids;
}

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

// 【v14 新增】跟 appendRowByHeader 幾乎一樣，差別是寫入前會先檢查表格裡
// 是不是已經有相同 id 的那一列：
// - 有的話，直接原地覆寫那一列，不會再多新增一列出來。
// - 沒有的話，才新增一列（行為跟 appendRowByHeader 完全一樣）。
// 這是「就算前端的補傳判斷誤判、把同一筆資料重複送過來，試算表也不會
// 生出重複列」的最後一道防線。idField 是這個 headers 裡代表 id 的欄位名稱
// （目前 Foods／Logs 都叫 'id'）。
function upsertRowByHeader(sheet, headers, idField, valuesObj) {
  var map = headerIndexMap(sheet);
  headers.forEach(function (h) {
    if (!map.hasOwnProperty(h) && valuesObj.hasOwnProperty(h)) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(h);
      map[h] = newCol - 1;
    }
  });

  var targetId = String(valuesObj[idField] == null ? '' : valuesObj[idField]).trim();
  var existingRowNum = -1;
  if (targetId && map.hasOwnProperty(idField)) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var idCol = sheet.getRange(2, map[idField] + 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < idCol.length; i++) {
        var rowId = String(idCol[i][0] == null ? '' : idCol[i][0]).trim();
        if (rowId === targetId) { existingRowNum = i + 2; break; }
      }
    }
  }

  var lastCol = Math.max(sheet.getLastColumn(), headers.length);
  var row = new Array(lastCol).fill('');
  headers.forEach(function (h) {
    if (map.hasOwnProperty(h) && valuesObj.hasOwnProperty(h)) {
      row[map[h]] = valuesObj[h];
    }
  });

  if (existingRowNum === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(existingRowNum, 1, 1, row.length).setValues([row]);
  }
}

// ---------- read ----------

function getData() {
  return {
    foods: readFoods(),
    logs: readLogs(),
    // 【v14 新增】把試算表這邊「大家共同知道的刪除紀錄」也回傳給前端，
    // 讓任何一台裝置連上來都能學到「這些 id 已經被刪掉了」，不用只靠
    // 自己本機那份 tombstone，才不會把別的裝置刪掉的東西誤判成新資料
    // 又補傳回去。
    deletedFoodIds: readDeletedIds('food'),
    deletedLogIds: readDeletedIds('log')
  };
}

function parseFoodCategories(value, legacyCategory) {
  if (Array.isArray(value)) {
    return value.map(function(v) { return String(v).trim(); }).filter(Boolean);
  }
  var s = value === null || value === undefined ? '' : String(value).trim();
  if (!s) s = legacyCategory === null || legacyCategory === undefined ? '' : String(legacyCategory).trim();
  if (!s) return [];
  // 兼容 JSON 陣列、中文頓號、逗號等既有資料格式。
  try {
    var parsed = JSON.parse(s);
    if (Array.isArray(parsed)) {
      return parsed.map(function(v) { return String(v).trim(); }).filter(Boolean);
    }
  } catch (err) {}
  return s.split(/[、,，]/).map(function(v) { return v.trim(); }).filter(Boolean);
}

function serializeFoodCategories(categories) {
  var list = parseFoodCategories(categories, '');
  var seen = {};
  return list.filter(function(c) {
    if (seen[c]) return false;
    seen[c] = true;
    return true;
  }).join('、');
}

function readFoods() {
  var sheet = getFoodsSheet();
  var map = headerIndexMap(sheet);
  var rows = sheet.getDataRange().getValues();
  var foods = [];
  var seenIds = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[map['id']]) continue;
    // 【v14 新增】防呆：如果試算表裡剛好還殘留著「同一個 id 出現不只一次」
    // 的舊資料（例如這次修復之前就已經重複寫入），這裡只取第一筆，畫面上
    // 就不會再顯示成一模一樣的兩筆。真正把試算表裡的重複列刪掉，
    // 請手動執行 dedupeFoodsAndLogsSheets()。
    var idKey = String(r[map['id']]);
    if (seenIds[idKey]) continue;
    seenIds[idKey] = true;
    foods.push({
      id: r[map['id']],
      name: r[map['name']],
      base: Number(r[map['base']]) || 100,
      unit: (map.hasOwnProperty('unit') && r[map['unit']]) ? String(r[map['unit']]) : 'g',
      servings: (map.hasOwnProperty('servings') && Number(r[map['servings']]) > 0) ? Number(r[map['servings']]) : 1,
      category: (map.hasOwnProperty('category') && r[map['category']]) ? String(r[map['category']]) : '',
      categories: parseFoodCategories(map.hasOwnProperty('categories') ? r[map['categories']] : '', map.hasOwnProperty('category') ? r[map['category']] : ''),
      outOfStock: map.hasOwnProperty('outOfStock') ? (r[map['outOfStock']] === true || String(r[map['outOfStock']]).toLowerCase() === 'true') : false,
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
  var seenLogIds = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[map['id']]) continue;
    // 【v14 新增】跟 readFoods() 一樣的防呆：同一個 id 只取第一筆。
    var logIdKey = String(r[map['id']]);
    if (seenLogIds[logIdKey]) continue;
    seenLogIds[logIdKey] = true;
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
    var groupIdRaw = map.hasOwnProperty('groupId') ? r[map['groupId']] : '';
    if (groupIdRaw !== '' && groupIdRaw !== undefined && groupIdRaw !== null) {
      logObj.groupId = String(groupIdRaw);
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
  // 【v14】改用 upsertRowByHeader：如果這個 id 在表格裡已經存在（例如前端
  // 補傳邏輯誤判、把同一筆又送了一次），會直接原地覆寫，不會再多長出一列
  // 重複資料。
  upsertRowByHeader(sheet, FOODS_HEADERS, 'id', {
    id: id,
    name: payload.name || '',
    base: Number(payload.base) || 100,
    unit: payload.unit ? String(payload.unit) : 'g',
    servings: Number(payload.servings) > 0 ? Number(payload.servings) : 1,
    category: serializeFoodCategories((payload.categories && parseFoodCategories(payload.categories, '')).length ? payload.categories : payload.category).split('、')[0] || '',
    categories: serializeFoodCategories((payload.categories && parseFoodCategories(payload.categories, '')).length ? payload.categories : payload.category),
    outOfStock: payload.outOfStock === true || String(payload.outOfStock).toLowerCase() === 'true',
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
      if (!map.hasOwnProperty('servings')) {
        var newServingsCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newServingsCol).setValue('servings');
        map['servings'] = newServingsCol - 1;
      }
      sheet.getRange(rowNum, map['servings'] + 1).setValue(Number(payload.servings) > 0 ? Number(payload.servings) : 1);
      var categoryText = serializeFoodCategories((payload.categories && parseFoodCategories(payload.categories, '')).length ? payload.categories : payload.category);
      if (!map.hasOwnProperty('category')) {
        var newCatCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCatCol).setValue('category');
        map['category'] = newCatCol - 1;
      }
      sheet.getRange(rowNum, map['category'] + 1).setValue(categoryText ? categoryText.split('、')[0] : '');
      if (!map.hasOwnProperty('categories')) {
        var newCategoriesCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCategoriesCol).setValue('categories');
        map['categories'] = newCategoriesCol - 1;
      }
      sheet.getRange(rowNum, map['categories'] + 1).setValue(categoryText);
      if (!map.hasOwnProperty('outOfStock')) {
        var newStockCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newStockCol).setValue('outOfStock');
        map['outOfStock'] = newStockCol - 1;
      }
      sheet.getRange(rowNum, map['outOfStock'] + 1).setValue(payload.outOfStock === true || String(payload.outOfStock).toLowerCase() === 'true');
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

// 【修復】以前這裡比對到「第一筆」符合的 id 就刪除、馬上 return，如果同一個
// id 因為之前同步時的競爭情況（例如按太快、或曾經同步失敗又重試）在表格裡
// 留下不只一列，就只會刪掉其中一列，另一列還在，重新整理後那項食材就會
// 「陰魂不散」地又出現。現在改成：用寬鬆比對（去除頭尾空白、轉成字串再比）
// 把所有符合的列「全部」刪除（從最後一列往前刪，避免刪除時 index 位移），
// 才不會再有殘留的重複列。
function deleteFood(payload) {
  var sheet = getFoodsSheet();
  var map = headerIndexMap(sheet);
  var data = sheet.getDataRange().getValues();
  var targetId = String(payload.id == null ? '' : payload.id).trim();
  var deletedCount = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    var rowId = String(data[i][map['id']] == null ? '' : data[i][map['id']]).trim();
    if (rowId === targetId) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }
  if (deletedCount > 0) {
    // 【v14 新增】把這個 id 記進 DeletedIds 分頁，讓其他裝置下次連上試算表
    // 時也能知道「這筆已經被刪掉了」，不會再把它當成自己漏同步的資料補傳。
    recordDeletion('food', [targetId]);
    SpreadsheetApp.flush();
    return { success: true, deletedCount: deletedCount };
  }
  return { error: 'food not found' };
}

function addLog(payload) {
  var sheet = getLogsSheet();
  // 【v14】以前這裡完全不理會前端送來的 id、永遠自己重新產生一個新的，
  // 前端還得在收到回應後把本機那筆的 id 換成試算表回傳的 id，多一層容易
  // 出錯的步驟（例如那次的回應剛好在網路不穩時遺失，前端就永遠以為
  // 沒同步成功，下次補傳又用回原本的 id 送一次、變成兩筆內容一樣但 id
  // 不同的重複紀錄）。現在改成：前端有帶 id 就沿用前端的 id，沒帶才自己
  // 產生；同時改用 upsertRowByHeader，這樣同一個 id 不管被送幾次，
  // 試算表裡永遠只會有一列，不會再重複。
  var id = payload.id || ('l_' + new Date().getTime());
  var sugarVal = payload.sugar !== undefined ? payload.sugar : payload.carb;
  upsertRowByHeader(sheet, LOGS_HEADERS, 'id', {
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
    order: payload.order != null ? Number(payload.order) : '',
    groupId: payload.groupId ? String(payload.groupId) : ''
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
      if (payload.groupId !== undefined) {
        if (!map.hasOwnProperty('groupId')) {
          var newGroupCol = sheet.getLastColumn() + 1;
          sheet.getRange(1, newGroupCol).setValue('groupId');
          map['groupId'] = newGroupCol - 1;
        }
        sheet.getRange(rowNum, map['groupId'] + 1).setValue(payload.groupId || '');
      }
      if (payload.date !== undefined) sheet.getRange(rowNum, map['date'] + 1).setValue(payload.date);
      if (payload.person !== undefined) sheet.getRange(rowNum, map['person'] + 1).setValue(payload.person || 'A');
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
      // 【v14 新增】同 deleteFood，把這個 id 記進 DeletedIds 分頁。
      recordDeletion('log', [payload.id]);
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

/**
 * 【本次新增】手動修復小工具：把「糙米飯」這個食材、以及「今日紀錄」裡
 * 過去所有吃「糙米飯」的紀錄，統一改名成「黃金胚芽」。
 *
 * 使用方式：在 Apps Script 編輯器上方的函式下拉選單選
 * renameBrownRiceToGermRice，按「執行」跑一次即可，不需要透過網頁前端。
 *
 * 這個做法直接把 Foods 分頁裡名字是「糙米飯」的那一列改名，而不是先刪除
 * 再新增——這樣就不會受到「刪除食材」目前這個 bug 影響（如果同一個 id
 * 在表格裡意外留下重複列，單純刪除可能只刪掉其中一列，改名則沒有這個問題），
 * 也保留了這項食材原本的 id、既有的每份克數／營養設定，不用重新輸入。
 * Logs 分頁裡舊的「今日紀錄」是各自存了當時吃的 foodName 文字快照，
 * 所以要另外把 foodName 是「糙米飯」的每一列，也一起改成「黃金胚芽」，
 * 之後「今日紀錄」畫面上看到的歷史紀錄才會跟著顯示新名字。
 * 如果你想改別的食材名稱，把下面兩個字串換掉即可。
 */
function renameBrownRiceToGermRice() {
  return renameFoodEverywhere('糙米飯', '黃金胚芽');
}

function renameFoodEverywhere(oldName, newName) {
  var renamedFoods = renameInSheet(getFoodsSheet(), 'name', oldName, newName);
  var renamedLogs = renameInSheet(getLogsSheet(), 'foodName', oldName, newName);
  return { renamedFoods: renamedFoods, renamedLogs: renamedLogs };
}

function renameInSheet(sheet, columnName, oldValue, newValue) {
  var map = headerIndexMap(sheet);
  if (!map.hasOwnProperty(columnName)) return 0;
  var col = map[columnName] + 1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var range = sheet.getRange(2, col, lastRow - 1, 1);
  var values = range.getValues();
  var changed = 0;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === oldValue) {
      values[i][0] = newValue;
      changed++;
    }
  }
  if (changed > 0) {
    range.setValues(values);
    SpreadsheetApp.flush();
  }
  return changed;
}

/**
 * 【本次新增】手動修復小工具：直接在試算表這一端把「糙米飯」這個食材、
 * 以及「澱粉」這個分類標籤，徹底清乾淨。
 *
 * 背景：前端已經加了「刪除食材時記一份 tombstone（我確實刪過的 id 清單），
 * 同步時排除這些 id」的修復，理論上之後用 App 裡「刪除食材」按鈕刪掉的東西
 * 不會再自己跑回來。但如果某一台裝置/瀏覽器分頁的本機快取，是在套用這個
 * 修復「之前」就已經把「糙米飯」記下來、且從來沒有點過「刪除食材」（單純
 * 編輯試算表、或用舊版前端刪過），那台裝置的本機清單裡就沒有對應的
 * tombstone 紀錄——下次那台裝置一連上試算表，還是會把「本機有、試算表沒有」
 * 的「糙米飯」誤判成「還沒同步的新資料」，自動又補傳回試算表一次。
 * 這個小工具直接在試算表原始資料這端做「終結」處理，不透過前端、也不受
 * 任何一台裝置的本機快取影響：
 *   1) 把 Foods 分頁裡名字是「糙米飯」的每一列（不只比對 id，比對「名字」，
 *      這樣不管過去復活過幾次、累積了幾個不同 id 的重複列，都會一次殺光）
 *      整列刪除。
 *   2) 把「澱粉」這個分類標籤，從 Foods 分頁的 category／categories 欄位
 *      整個移除（只拿掉這個分類文字，不會動到食材本身的其他資料；如果某個
 *      食材除了「澱粉」還有別的分類，只會把「澱粉」拿掉，其他分類不受影響）。
 *
 * 使用方式：在 Apps Script 編輯器上方的函式下拉選單選
 * purgeBrownRiceFoodAndStarchCategory，按「執行」跑一次即可。
 * 這個工具只動 Foods 分頁，不會刪除 Logs 分頁裡過去吃「糙米飯」的歷史紀錄
 * ——那些紀錄本身是你已經吃過的事實，只是紀錄裡存的食材名稱文字快照，
 * 不影響「糙米飯」不會再出現在食材庫、也不會再被拿來新增紀錄。
 * 如果你也想清掉別的食材／分類，把下面兩個字串換掉即可；不需要用到的那個
 * （例如只想刪食材、不想動分類）可以把對應那一行呼叫拿掉或把參數傳空字串。
 */
function purgeBrownRiceFoodAndStarchCategory() {
  var deletedFoodRows = deleteFoodsByName('糙米飯');
  var strippedCategoryFrom = removeCategoryEverywhere('澱粉');
  return { deletedFoodRows: deletedFoodRows, strippedCategoryFrom: strippedCategoryFrom };
}

// 依「名字」（而不是 id）刪除 Foods 分頁裡所有符合的列——同一個名字不管
// 因為過去的同步/復活問題累積了幾個不同 id 的重複列，這裡一律當作同一種
// 食材，全部刪光，才不會殺了一個 id 卻漏掉另一個 id 的重複列。
function deleteFoodsByName(name) {
  var sheet = getFoodsSheet();
  var map = headerIndexMap(sheet);
  var data = sheet.getDataRange().getValues();
  var target = String(name == null ? '' : name).trim();
  if (!target) return 0;
  var deleted = 0;
  var deletedIds = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var rowName = String(data[i][map['name']] == null ? '' : data[i][map['name']]).trim();
    if (rowName === target) {
      var rowId = String(data[i][map['id']] == null ? '' : data[i][map['id']]).trim();
      if (rowId) deletedIds.push(rowId);
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  if (deleted > 0) {
    // 【v14 新增】跟 deleteFood() 一樣，把刪掉的 id 記進 DeletedIds 分頁，
    // 避免其他裝置的本機舊快取之後又把它當成新資料補傳回來。
    if (deletedIds.length) recordDeletion('food', deletedIds);
    SpreadsheetApp.flush();
  }
  return deleted;
}

// 把某個分類標籤從 Foods 分頁的 category（單一）／categories（多選，用「、」
// 分隔）兩個欄位裡整個拿掉。只移除這個分類文字本身，其餘分類、其餘欄位資料
// 都不會被動到；如果某個食材原本只有這一個分類，該欄位會變成空字串
// （代表這個食材現在沒有分類），不會連食材本身一起刪掉。
function removeCategoryEverywhere(categoryName) {
  var target = String(categoryName == null ? '' : categoryName).trim();
  if (!target) return 0;
  var sheet = getFoodsSheet();
  var map = headerIndexMap(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var changed = 0;

  if (map.hasOwnProperty('categories')) {
    var range = sheet.getRange(2, map['categories'] + 1, lastRow - 1, 1);
    var values = range.getValues();
    var anyChanged = false;
    for (var i = 0; i < values.length; i++) {
      var raw = String(values[i][0] || '');
      if (!raw) continue;
      var parts = raw.split('、').map(function (s) { return s.trim(); }).filter(Boolean);
      var idx = parts.indexOf(target);
      if (idx > -1) {
        parts.splice(idx, 1);
        values[i][0] = parts.join('、');
        changed++;
        anyChanged = true;
      }
    }
    if (anyChanged) range.setValues(values);
  }

  if (map.hasOwnProperty('category')) {
    var range2 = sheet.getRange(2, map['category'] + 1, lastRow - 1, 1);
    var values2 = range2.getValues();
    var anyChanged2 = false;
    for (var j = 0; j < values2.length; j++) {
      if (String(values2[j][0] || '').trim() === target) {
        values2[j][0] = '';
        anyChanged2 = true;
      }
    }
    if (anyChanged2) range2.setValues(values2);
  }

  if (changed > 0) SpreadsheetApp.flush();
  return changed;
}

/**
 * 【v14 新增】一次性清理工具：把 Foods／Logs 分頁裡「同一個 id 出現不只
 * 一次」的重複列直接刪掉，只保留最先出現的那一列（其餘欄位資料不會被
 * 改動，只是把多餘的重複列移除）。
 *
 * 這是用來清掉「這次修復之前」就已經因為補傳誤判而產生的重複資料
 * （例如你貼的那兩列一模一樣的「糙米飯」）。修復之後，正常操作理論上
 * 不會再產生新的重複列，這個工具只需要手動執行這一次即可。
 *
 * 使用方式：在 Apps Script 編輯器上方的函式下拉選單選
 * dedupeFoodsAndLogsSheets，按「執行」跑一次即可。
 */
function dedupeFoodsAndLogsSheets() {
  var removedFoods = dedupeSheetById(getFoodsSheet());
  var removedLogs = dedupeSheetById(getLogsSheet());
  return { removedFoodRows: removedFoods, removedLogRows: removedLogs };
}

function dedupeSheetById(sheet) {
  var map = headerIndexMap(sheet);
  if (!map.hasOwnProperty('id')) return 0;
  var data = sheet.getDataRange().getValues();
  var seen = {};
  var rowsToDelete = [];
  // 先由上往下掃過一遍，找出「第一次以外」的重複列該刪掉的實際列號，
  // 這樣才能保證留下來的是最先出現的那一列。
  for (var i = 1; i < data.length; i++) {
    var idVal = String(data[i][map['id']] == null ? '' : data[i][map['id']]).trim();
    if (!idVal) continue;
    if (seen[idVal]) {
      rowsToDelete.push(i + 1); // 試算表的實際列號（第 1 列是表頭）
    } else {
      seen[idVal] = true;
    }
  }
  // 再由下往上刪，避免刪除當下讓還沒處理到的列號跟著位移跑掉。
  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(rowsToDelete[j]);
  }
  if (rowsToDelete.length > 0) SpreadsheetApp.flush();
  return rowsToDelete.length;
}