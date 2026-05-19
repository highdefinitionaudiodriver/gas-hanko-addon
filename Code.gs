/**
 * 電子印鑑（データ印）— Google Workspace 用 GAS アドオン
 * スプレッドシート / ドキュメント / スライドに対応
 * サイドバーで印鑑画像を生成し、各アプリの適切な位置に挿入する
 */

// ==========================================================
//  メールアドレス → 苗字 マッピング
//  ※ 実際の運用時はここにユーザーを追加してください
// ==========================================================
var NAME_MAP = {
  "taro.yamada@example.com": "山田",
  "hanako.suzuki@example.com": "鈴木",
  "ichiro.tanaka@example.com": "田中",
  "kindaichi@example.com": "金田一"
};

// ==========================================================
//  ホストアプリ判定ヘルパー
// ==========================================================

/**
 * 現在のホストアプリケーションを判定する
 * @return {string} "spreadsheet" | "document" | "slides" | "unknown"
 */
function getHostApp_() {
  try { SpreadsheetApp.getActiveSpreadsheet(); return "spreadsheet"; } catch (e) {}
  try { DocumentApp.getActiveDocument();       return "document";    } catch (e) {}
  try { SlidesApp.getActivePresentation();     return "slides";      } catch (e) {}
  return "unknown";
}

/**
 * ホストアプリに応じた UI オブジェクトを返す
 * @return {Ui}
 */
function getUi_() {
  var host = getHostApp_();
  switch (host) {
    case "spreadsheet": return SpreadsheetApp.getUi();
    case "document":    return DocumentApp.getUi();
    case "slides":      return SlidesApp.getUi();
    default:
      throw new Error("未対応のアプリケーションです");
  }
}

// ==========================================================
//  メニュー & サイドバー
// ==========================================================

/**
 * ファイルを開いた時にメニューを追加
 * （スプレッドシート / ドキュメント / スライド 共通）
 */
function onOpen() {
  getUi_()
    .createMenu("電子印鑑")
    .addItem("印鑑パネルを開く", "showSidebar")
    .addToUi();
}

/**
 * サイドバーを表示
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("電子印鑑（データ印）")
    .setWidth(300);
  getUi_().showSidebar(html);
}

// ==========================================================
//  苗字取得（メール → 名前マッピング）
// ==========================================================

/**
 * 実行ユーザーのメールアドレスから苗字を返す
 * マッピングに一致すればその苗字、なければ @ より前の文字列を返す
 * @return {string}
 */
function getUserLastName() {
  var email = Session.getActiveUser().getEmail();
  if (!email) return "山田"; // メール取得不可の場合のフォールバック

  if (NAME_MAP[email]) {
    return NAME_MAP[email];
  }

  // @ より前の部分をデフォルト値として返す
  return email.split("@")[0];
}

// ==========================================================
//  画像挿入（クロスアプリ対応）
// ==========================================================

/**
 * Base64 PNG 画像を現在のアプリの適切な位置に挿入する
 * @param {string} base64Data - data:image/png;base64, プレフィックス付き
 * @return {string} 結果メッセージ
 */
function insertStampImage(base64Data) {
  var raw = base64Data.replace(/^data:image\/png;base64,/, "");
  var blob = Utilities.newBlob(Utilities.base64Decode(raw), "image/png", "hanko.png");

  var host = getHostApp_();
  switch (host) {
    case "spreadsheet":
      return insertIntoSpreadsheet_(blob);
    case "document":
      return insertIntoDocument_(blob);
    case "slides":
      return insertIntoSlides_(blob);
    default:
      throw new Error("未対応のアプリケーションです: " + host);
  }
}

/**
 * スプレッドシート: アクティブセルの位置に挿入
 */
function insertIntoSpreadsheet_(blob) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();
  var row = cell.getRow();
  var col = cell.getColumn();

  var image = sheet.insertImage(blob, col, row);
  var targetSize = 60;
  image.setWidth(targetSize);
  image.setHeight(targetSize);

  return "スプレッドシートに押印しました (行:" + row + " 列:" + col + ")";
}

/**
 * ドキュメント: カーソル位置にインライン画像として挿入
 */
function insertIntoDocument_(blob) {
  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();

  if (cursor) {
    var element = cursor.insertInlineImage(blob);
    // サイズ調整（ポイント単位: 60px ≈ 45pt）
    var targetPt = 45;
    element.setWidth(targetPt);
    element.setHeight(targetPt);
    return "ドキュメントのカーソル位置に押印しました";
  }

  // カーソルが取得できない場合（範囲選択中など）→ Body末尾に挿入
  var body = doc.getBody();
  var paragraph = body.appendParagraph("");
  var image = paragraph.appendInlineImage(blob);
  var targetPt = 45;
  image.setWidth(targetPt);
  image.setHeight(targetPt);
  return "ドキュメント末尾に押印しました（カーソル位置を取得できなかったため）";
}

// ==========================================================
//  証跡 (Audit Log) — UserProperties で個人ごとに永続化
//  法人プラン想定: Drive Sheet への一括同期・部署別集計を別途提供
// ==========================================================

var AUDIT_LOG_KEY = "hanko_audit_log_v1";
var AUDIT_LOG_MAX_BYTES = 8 * 1024;  // UserProperties は 9KB/key が上限。8KB で警告

/**
 * 押印イベントを証跡に追記する
 * @param {Object} entry { trackingId, timestamp, topText, midText, bottomText }
 * @return {Object} { ok, count, warning? }
 */
function appendAuditEntry(entry) {
  if (!entry || typeof entry !== "object") throw new Error("invalid entry");
  var props = PropertiesService.getUserProperties();
  var raw = props.getProperty(AUDIT_LOG_KEY);
  var log = [];
  if (raw) {
    try { log = JSON.parse(raw); if (!Array.isArray(log)) log = []; } catch (e) { log = []; }
  }
  var record = {
    trackingId: String(entry.trackingId || ""),
    timestamp:  entry.timestamp || new Date().toISOString(),
    host:       getHostApp_(),
    user:       Session.getActiveUser().getEmail() || "",
    topText:    String(entry.topText || ""),
    midText:    String(entry.midText || ""),
    bottomText: String(entry.bottomText || "")
  };
  log.push(record);

  var serialized = JSON.stringify(log);
  var warning = null;
  if (serialized.length > AUDIT_LOG_MAX_BYTES) {
    // 古いエントリから消して 80% 以下まで縮める
    while (log.length > 0 && JSON.stringify(log).length > AUDIT_LOG_MAX_BYTES * 0.8) {
      log.shift();
    }
    serialized = JSON.stringify(log);
    warning = "ローカル証跡容量上限に達したため、古いエントリを自動削除しました（法人プランでクラウド集約推奨）";
  }
  props.setProperty(AUDIT_LOG_KEY, serialized);
  return { ok: true, count: log.length, warning: warning };
}

/**
 * 証跡の件数を返す
 * @return {number}
 */
function getAuditCount() {
  var props = PropertiesService.getUserProperties();
  var raw = props.getProperty(AUDIT_LOG_KEY);
  if (!raw) return 0;
  try {
    var log = JSON.parse(raw);
    return Array.isArray(log) ? log.length : 0;
  } catch (e) { return 0; }
}

/**
 * 証跡を全削除
 * @return {Object}
 */
function clearAuditLog() {
  PropertiesService.getUserProperties().deleteProperty(AUDIT_LOG_KEY);
  return { ok: true };
}

/**
 * CSV 1 セルを RFC4180 風にエスケープ
 */
function csvCell_(v) {
  if (v == null) return "";
  var s = String(v);
  if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

/**
 * 証跡を CSV 文字列として返す（UTF-8 BOM 付き）
 * クライアントが Blob + <a download> でダウンロードする
 * @return {Object} { ok, csv, count, filename }
 */
function exportAuditCsv() {
  var props = PropertiesService.getUserProperties();
  var raw = props.getProperty(AUDIT_LOG_KEY);
  var log = [];
  if (raw) {
    try { log = JSON.parse(raw); if (!Array.isArray(log)) log = []; } catch (e) { log = []; }
  }
  var headers = ["timestamp_iso8601","tracking_id","host_app","user_email","top_text","middle_text","bottom_text"];
  var lines = [headers.join(",")];
  for (var i = 0; i < log.length; i++) {
    var e = log[i];
    lines.push([
      csvCell_(e.timestamp), csvCell_(e.trackingId), csvCell_(e.host),
      csvCell_(e.user), csvCell_(e.topText), csvCell_(e.midText), csvCell_(e.bottomText)
    ].join(","));
  }
  var csv = "﻿" + lines.join("\r\n") + "\r\n";

  var d = new Date();
  var pad = function (n) { return ("0" + n).slice(-2); };
  var fname = "hanko_audit_" +
    d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + "_" +
    pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds()) + ".csv";
  return { ok: true, csv: csv, count: log.length, filename: fname };
}

/**
 * スライド: 現在選択中のスライドに挿入
 */
function insertIntoSlides_(blob) {
  var presentation = SlidesApp.getActivePresentation();
  var selection = presentation.getSelection();
  var slide;

  // 選択中のスライドを取得
  var currentPage = selection.getCurrentPage();
  if (currentPage && currentPage.getPageType() === SlidesApp.PageType.SLIDE) {
    slide = currentPage.asSlide();
  } else {
    // フォールバック: 最初のスライド
    var slides = presentation.getSlides();
    if (slides.length === 0) {
      throw new Error("スライドが存在しません");
    }
    slide = slides[0];
  }

  // スライド中央付近に挿入（単位: ポイント）
  var image = slide.insertImage(blob);
  var targetPt = 60;
  image.setWidth(targetPt);
  image.setHeight(targetPt);
  // 中央寄せ
  var pageWidth = presentation.getPageWidth();
  var pageHeight = presentation.getPageHeight();
  image.setLeft((pageWidth - targetPt) / 2);
  image.setTop((pageHeight - targetPt) / 2);

  return "スライドに押印しました";
}
