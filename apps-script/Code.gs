/**
 * B払い計算ツール - 商品リスト共有用バックエンド
 * このファイルは Google スプレッドシートの「拡張機能 > Apps Script」に貼り付けて使います。
 *
 * セットアップ手順は README.md を参照してください。
 */

const SHEET_NAME = "products";
const HEADERS = ["id", "name", "price", "quantity", "unit"];

// ここを必ず自分だけが知っている文字列に変更してください（第三者による書き換え防止用）
const API_TOKEN = "REPLACE_WITH_YOUR_OWN_SECRET";

function doGet(e) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(r => r[0] !== "" && r[0] !== null);
  const products = rows.map(r => ({
    id: r[0],
    name: r[1],
    price: r[2],
    quantity: r[3] || 0,
    unit: r[4] || ""
  }));
  return jsonOutput_({ products });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ error: "invalid_json" });
  }

  if (body.token !== API_TOKEN) {
    return jsonOutput_({ error: "unauthorized" });
  }
  if (!Array.isArray(body.products)) {
    return jsonOutput_({ error: "invalid_payload" });
  }

  // 複数端末からほぼ同時に書き込まれても内容が混ざらないよう、書き込み中は他の
  // 書き込みをロックで待たせる（これがないと行が重複・増殖することがある）。
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const rows = [HEADERS].concat(
      body.products.map(p => [p.id, p.name, p.price, p.quantity || 0, p.unit || ""])
    );
    const prevLastRow = sheet.getLastRow();
    // 先に新データを書き込んでから余った古い行を消すことで、読み取り側が
    // 一瞬「空」の状態を見てしまう(＝誤って再シードしてしまう)のを防ぐ。
    sheet.getRange(1, 1, rows.length, HEADERS.length).setValues(rows);
    if (prevLastRow > rows.length) {
      sheet.getRange(rows.length + 1, 1, prevLastRow - rows.length, HEADERS.length).clearContent();
    }
  } finally {
    lock.releaseLock();
  }

  return jsonOutput_({ ok: true });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
