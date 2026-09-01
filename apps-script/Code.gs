/**
 * B払い計算ツール - 商品リスト・計算書共有用バックエンド
 * このファイルは Google スプレッドシートの「拡張機能 > Apps Script」に貼り付けて使います。
 *
 * セットアップ手順は README.md を参照してください。
 */

const PRODUCTS_SHEET_NAME = "products";
const PRODUCTS_HEADERS = ["id", "name", "price", "quantity", "unit", "d", "w", "h"];

const QUOTES_SHEET_NAME = "quotes";
const QUOTES_HEADERS = ["id", "owner", "name", "savedAt", "stateJson", "status", "confirmedAt"];

// 計算書が確定されたときに通知するメールアドレス
const ADMIN_NOTIFY_EMAIL = "umeki.seitaro@gmail.com";

const SHIPPING_RATES_SHEET_NAME = "shipping_rates";
const SHIPPING_RATES_HEADERS = ["id", "origin", "size", "region", "fee"];
const SHIPPING_AREAS_SHEET_NAME = "shipping_areas";
const SHIPPING_AREAS_HEADERS = ["name"];

// ここを必ず自分だけが知っている文字列に変更してください（第三者による書き換え防止用）
const API_TOKEN = "YGfcWJDnrN8fXEs2m4ru0AOGFR-z63KW";

function doGet(e) {
  const type = (e.parameter && e.parameter.type) || "products";
  if (type === "quotes") {
    return jsonOutput_({ quotes: readQuotes_() });
  }
  if (type === "shipping") {
    return jsonOutput_({ areaNames: readShippingAreas_(), rates: readShippingRates_() });
  }
  return jsonOutput_({ products: readProducts_() });
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

  if (body.type === "quotes") {
    return handleQuotesPost_(body);
  }
  if (body.type === "shipping") {
    return handleShippingPost_(body);
  }
  return handleProductsPost_(body);
}

// ---------- products（既存の一括置き換え方式。変更なし） ----------
function readProducts_() {
  const sheet = getSheet_(PRODUCTS_SHEET_NAME, PRODUCTS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(r => r[0] !== "" && r[0] !== null);
  return rows.map(r => ({
    id: r[0],
    name: r[1],
    price: r[2],
    quantity: r[3] || 0,
    unit: r[4] || "",
    d: r[5] || 0,
    w: r[6] || 0,
    h: r[7] || 0
  }));
}

function handleProductsPost_(body) {
  if (!Array.isArray(body.products)) {
    return jsonOutput_({ error: "invalid_payload" });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_(PRODUCTS_SHEET_NAME, PRODUCTS_HEADERS);
    const rows = [PRODUCTS_HEADERS].concat(
      body.products.map(p => [p.id, p.name, p.price, p.quantity || 0, p.unit || "", p.d || 0, p.w || 0, p.h || 0])
    );
    const prevLastRow = sheet.getLastRow();
    // 先に新データを書き込んでから余った古い行を消すことで、読み取り側が
    // 一瞬「空」の状態を見てしまう(＝誤って再シードしてしまう)のを防ぐ。
    sheet.getRange(1, 1, rows.length, PRODUCTS_HEADERS.length).setValues(rows);
    if (prevLastRow > rows.length) {
      sheet.getRange(rows.length + 1, 1, prevLastRow - rows.length, PRODUCTS_HEADERS.length).clearContent();
    }
  } finally {
    lock.releaseLock();
  }
  return jsonOutput_({ ok: true });
}

// ---------- quotes（複数人が同時に保存しても壊れないよう、1件ずつ追加・更新・削除する） ----------
function readQuotes_() {
  const sheet = getSheet_(QUOTES_SHEET_NAME, QUOTES_HEADERS);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(r => r[0] !== "" && r[0] !== null);
  return rows.map(r => {
    let state = {};
    try { state = JSON.parse(r[4] || "{}"); } catch (err) {}
    return {
      id: r[0],
      owner: r[1],
      name: r[2],
      savedAt: r[3],
      state: state,
      status: r[5] || "draft",
      confirmedAt: r[6] || null
    };
  });
}

function handleQuotesPost_(body) {
  const action = body.action;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_(QUOTES_SHEET_NAME, QUOTES_HEADERS);

    if (action === "add") {
      const q = body.quote;
      if (!q || !q.id) return jsonOutput_({ error: "invalid_payload" });
      sheet.appendRow([
        q.id, q.owner || "", q.name || "", q.savedAt || Date.now(), JSON.stringify(q.state || {}),
        q.status || "draft", q.confirmedAt || ""
      ]);
      return jsonOutput_({ ok: true });
    }

    if (action === "update") {
      const q = body.quote;
      if (!q || !q.id) return jsonOutput_({ error: "invalid_payload" });
      const rowIndex = findQuoteRow_(sheet, q.id);
      if (rowIndex === -1) return jsonOutput_({ error: "not_found" });
      sheet.getRange(rowIndex, 1, 1, QUOTES_HEADERS.length).setValues(
        [[q.id, q.owner || "", q.name || "", q.savedAt || Date.now(), JSON.stringify(q.state || {}),
          q.status || "draft", q.confirmedAt || ""]]
      );
      return jsonOutput_({ ok: true });
    }

    if (action === "confirm") {
      const id = body.id;
      if (!id) return jsonOutput_({ error: "invalid_payload" });
      const rowIndex = findQuoteRow_(sheet, id);
      if (rowIndex === -1) return jsonOutput_({ error: "not_found" });
      const confirmedAt = Date.now();
      sheet.getRange(rowIndex, 6, 1, 2).setValues([["confirmed", confirmedAt]]);
      const row = sheet.getRange(rowIndex, 1, 1, QUOTES_HEADERS.length).getValues()[0];
      notifyQuoteConfirmed_({ id: row[0], owner: row[1], name: row[2] });
      return jsonOutput_({ ok: true });
    }

    if (action === "delete") {
      const id = body.id;
      if (!id) return jsonOutput_({ error: "invalid_payload" });
      const rowIndex = findQuoteRow_(sheet, id);
      if (rowIndex === -1) return jsonOutput_({ error: "not_found" });
      sheet.deleteRow(rowIndex);
      return jsonOutput_({ ok: true });
    }

    return jsonOutput_({ error: "unknown_action" });
  } finally {
    lock.releaseLock();
  }
}

function notifyQuoteConfirmed_(q) {
  if (!ADMIN_NOTIFY_EMAIL) return;
  try {
    MailApp.sendEmail({
      to: ADMIN_NOTIFY_EMAIL,
      subject: "【B払い計算ツール】計算書が確定されました：" + q.name,
      body:
        "計算書が確定されました。\n\n" +
        "担当者：" + q.owner + "\n" +
        "計算書名：" + q.name + "\n" +
        "確定日時：" + new Date().toLocaleString("ja-JP") + "\n\n" +
        "管理者ページから内容をご確認ください。"
    });
  } catch (err) {
    // メール送信に失敗しても確定処理自体は成功させる
  }
}

function findQuoteRow_(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-indexed row number
  }
  return -1;
}

// ---------- shipping（送料設定。管理者のみ編集する想定のため products と同じ一括置き換え方式） ----------
function readShippingRates_() {
  const sheet = getSheet_(SHIPPING_RATES_SHEET_NAME, SHIPPING_RATES_HEADERS);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(r => r[0] !== "" && r[0] !== null);
  return rows.map(r => ({ id: r[0], origin: r[1], size: r[2], region: r[3], fee: r[4] || 0 }));
}

function readShippingAreas_() {
  const sheet = getSheet_(SHIPPING_AREAS_SHEET_NAME, SHIPPING_AREAS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(r => r[0] !== "" && r[0] !== null);
  return rows.map(r => r[0]);
}

function handleShippingPost_(body) {
  if (!Array.isArray(body.rates) || !Array.isArray(body.areaNames)) {
    return jsonOutput_({ error: "invalid_payload" });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const rateSheet = getSheet_(SHIPPING_RATES_SHEET_NAME, SHIPPING_RATES_HEADERS);
    const rateRows = [SHIPPING_RATES_HEADERS].concat(
      body.rates.map(r => [r.id, r.origin, r.size, r.region, r.fee || 0])
    );
    const prevRateLastRow = rateSheet.getLastRow();
    rateSheet.getRange(1, 1, rateRows.length, SHIPPING_RATES_HEADERS.length).setValues(rateRows);
    if (prevRateLastRow > rateRows.length) {
      rateSheet.getRange(rateRows.length + 1, 1, prevRateLastRow - rateRows.length, SHIPPING_RATES_HEADERS.length).clearContent();
    }

    const areaSheet = getSheet_(SHIPPING_AREAS_SHEET_NAME, SHIPPING_AREAS_HEADERS);
    const areaRows = [SHIPPING_AREAS_HEADERS].concat(body.areaNames.map(name => [name]));
    const prevAreaLastRow = areaSheet.getLastRow();
    areaSheet.getRange(1, 1, areaRows.length, SHIPPING_AREAS_HEADERS.length).setValues(areaRows);
    if (prevAreaLastRow > areaRows.length) {
      areaSheet.getRange(areaRows.length + 1, 1, prevAreaLastRow - areaRows.length, SHIPPING_AREAS_HEADERS.length).clearContent();
    }
  } finally {
    lock.releaseLock();
  }
  return jsonOutput_({ ok: true });
}

// ---------- common ----------
function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
