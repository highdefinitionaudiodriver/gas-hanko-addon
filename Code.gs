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
