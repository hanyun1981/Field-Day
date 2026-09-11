/**
 * Field Day秩序冊 · Google Sheets 後端
 *
 * === 部署步驟 ===
 * 1. 在 Google Sheets 建立一份新的試算表（隨便取名，例如「Field Day資料庫」）
 * 2. 點 [Extensions 擴充功能] > [Apps Script]
 * 3. 把預設的 code.gs 內容全刪除，貼入這整段程式碼
 * 4. 存檔（Ctrl/⌘ + S）
 * 5. 上方選單選 [執行] > 選函式 setup > 點 [執行]
 *    - 第一次會要求授權，同意即可
 *    - 授權後會自動建立 8 個工作表 (events, students, ...)
 * 6. 點右上 [部署] > [新增部署作業]
 *    - 類型：[網頁應用程式 Web app]
 *    - 說明：填「Field Day API」
 *    - 執行身分：[我]
 *    - 誰可以存取：[任何人]（或 [組織內所有人] 若在教育版）
 *    - 點 [部署]
 * 7. 授權後會顯示 [Web app URL]，複製這個網址
 * 8. 在 field-day.html 系統的「班級與權限 > 資料後端」貼入這個網址
 *
 * === 更新 ===
 * 若之後修改了此檔案，要點 [部署] > [管理部署作業]
 * > 對現有部署點鉛筆圖示 > 版本改為 [新版本] > [部署]
 * （URL 不會變）
 */

const COLLECTIONS = ["events", "students", "registrations", "results", "teams", "classes", "years", "settings"];

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "listAll";
  return handle(action, e.parameter || {});
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return handle(body.action, body);
}

function handle(action, params) {
  try {
    if (action === "listAll") return json({ data: listAll() });
    if (action === "list") return json({ data: listCollection(params.collection) });
    if (action === "set") return json({ ok: true, id: setDoc(params.collection, params.id, params.data) });
    if (action === "add") return json({ ok: true, id: addDoc(params.collection, params.data) });
    if (action === "delete") return json({ ok: true, deleted: deleteDoc(params.collection, params.id) });
    if (action === "batch") return json({ ok: true, results: batchOps(params.ops) });
    if (action === "init") return json({ ok: true, sheets: initSheets() });
    if (action === "ping") return json({ ok: true, version: "1.0" });
    return json({ error: "unknown action: " + action });
  } catch (err) {
    return json({ error: String(err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, 2).setValues([["id", "data"]]);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 200);
    sh.setColumnWidth(2, 800);
    sh.getRange("A1:B1").setFontWeight("bold").setBackground("#f5f2ea");
  }
  return sh;
}

function listCollection(name) {
  const sh = getSheet(name);
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const id = String(values[i][0] || "").trim();
    if (!id) continue;
    let data = {};
    try { data = JSON.parse(values[i][1] || "{}"); } catch (e) {}
    out.push(Object.assign({ id: id }, data));
  }
  return out;
}

function listAll() {
  const out = {};
  COLLECTIONS.forEach(c => { out[c] = listCollection(c); });
  return out;
}

function setDoc(collection, id, data) {
  const sh = getSheet(collection);
  const dataJson = JSON.stringify(data || {});
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sh.getRange(i + 1, 2).setValue(dataJson);
      return id;
    }
  }
  sh.appendRow([id, dataJson]);
  return id;
}

function addDoc(collection, data) {
  const id = "id_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  setDoc(collection, id, data);
  return id;
}

function deleteDoc(collection, id) {
  const sh = getSheet(collection);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return id;
    }
  }
  return null;
}

function batchOps(ops) {
  const results = [];
  (ops || []).forEach(op => {
    if (op.action === "set") results.push({ id: setDoc(op.collection, op.id, op.data) });
    else if (op.action === "add") results.push({ id: addDoc(op.collection, op.data) });
    else if (op.action === "delete") results.push({ deleted: deleteDoc(op.collection, op.id) });
  });
  return results;
}

function initSheets() {
  COLLECTIONS.forEach(c => getSheet(c));
  return COLLECTIONS;
}

// 手動執行一次以初始化工作表 + 授權（部署前必做）
function setup() {
  initSheets();
  try {
    SpreadsheetApp.getUi().alert(
      'Field Day資料庫初始化完成\n\n已建立 8 個工作表：' + COLLECTIONS.join(", ") +
      '\n\n下一步：\n右上角 [部署] > [新增部署作業] > 類型選 [網頁應用程式]\n' +
      '執行身分：我\n誰可以存取：任何人\n\n部署後複製 Web app URL 貼到系統。'
    );
  } catch (e) {}
}
