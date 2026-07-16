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

  const sheet = getSheet_();
  sheet.clearContents();
  sheet.appendRow(HEADERS);
  body.products.forEach(p => {
    sheet.appendRow([p.id, p.name, p.price, p.quantity || 0, p.unit || ""]);
  });

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
