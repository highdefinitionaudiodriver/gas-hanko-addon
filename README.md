# GAS Hanko Add-on (電子印鑑 for Google Workspace)

日本のビジネスシーンで使われる**データ印（電子印鑑）**を、Google スプレッドシート・ドキュメント・スライドに挿入する Google Apps Script アドオンです。

---

## 🎯 これは何？（30秒で）

- **誰のため**：Google Workspace を導入しているが日本式押印文化が残っている中小企業／在宅勤務環境の総務・経理
- **何が解決される**：Google ドキュメント／スプレッドシート／スライド内で **ワンクリック電子押印**。社内承認フローを Google Workspace で完結
- **なぜ既存ツールではダメか**：商用電子印鑑サービスは大半が Microsoft Office 向け、または独自プラットフォーム。本ツールは **Google Workspace Marketplace 配布可能**で、組織展開も個人利用もすぐ可能
- **使う条件**：Google アカウント（個人・Workspace どちらも可）／ブラウザのみ・インストール不要

## 💰 想定ユースケース・価格帯

| 用途 | 形態 |
|---|---|
| 個人利用 | 無料（MIT） |
| 法人利用（証跡 CSV エクスポート・部署一括展開） | 今後の有料機能として検討 |
| カスタマイズ・社内テンプレ統一・Google SSO 連携 | 応相談 |

![Google Workspace](https://img.shields.io/badge/Google%20Workspace-Sheets%20%7C%20Docs%20%7C%20Slides-34A853)
![License](https://img.shields.io/badge/license-MIT-blue)
![GAS](https://img.shields.io/badge/runtime-Google%20Apps%20Script%20V8-4285F4)

---

## 特徴

- **3アプリ対応** — スプレッドシート / ドキュメント / スライドで共通利用
- **Canvas で動的生成** — 丸型・赤色・3段構成（苗字・日付・部署名）のデータ印をリアルタイムプレビュー
- **アカウント連携** — Googleアカウントのメールアドレスから苗字を自動入力
- **トラッキングID** — 押印ごとに8桁のユニークIDを印影に埋め込み（簡易証跡）
- **インストール不要** — ブラウザだけで動作。Node.js やローカルサーバーは不要

## スクリーンショット

```
    ┌─────────────┐
    │  電子印鑑    │  ← サイドバー
    │             │
    │ 上段: 山田   │
    │ 日付: 2026/  │
    │ 下段: 営業部  │
    │             │
    │  [押 印]     │
    │             │
    │   ◯ 印影    │  ← リアルタイムプレビュー
    │  プレビュー   │
    └─────────────┘
```

## セットアップ

### 方法1: スクリプトエディタに直接貼り付け（最も簡単）

1. Google スプレッドシート（またはドキュメント / スライド）を開く
2. **拡張機能** → **Apps Script** をクリック
3. デフォルトの `コード.gs` を開き、`Code.gs` の内容で上書き
4. 左サイドバーの **＋** → **HTML** → ファイル名を `Sidebar` にして `Sidebar.html` の内容を貼り付け
5. **プロジェクトの設定**（歯車アイコン）→ 「`appsscript.json` マニフェスト ファイルをエディタで表示する」にチェック → `appsscript.json` の内容で上書き
6. **保存**（Ctrl+S）してスプレッドシートを**リロード**
7. メニューバーに「**電子印鑑**」が追加される

### 方法2: clasp を使ったデプロイ（開発者向け）

```bash
# clasp をインストール（未導入の場合）
npm install -g @google/clasp

# ログイン
clasp login

# プロジェクトを作成してプッシュ
clasp create --type sheets --title "電子印鑑"
clasp push
```

## 使い方

1. メニューバーの **「電子印鑑」** → **「印鑑パネルを開く」** をクリック
2. サイドバーが表示される（苗字はアカウント情報から自動入力）
3. 各欄を入力:
   | 欄 | 内容 | 例 |
   |---|---|---|
   | 上段 | 苗字など | `山田` |
   | 中段 | 日付（自動入力済み） | `2026/04/06` |
   | 下段 | 部署名など（空欄可） | `営業部` |
4. プレビューで印影を確認
5. 挿入したい位置を選択（スプレッドシートならセル、ドキュメントならカーソル位置）
6. **「押印」** ボタンをクリック

### アプリ別の挿入動作

| アプリ | 挿入位置 | 画像形式 |
|---|---|---|
| スプレッドシート | アクティブセルの位置 | OverGridImage (60px) |
| ドキュメント | カーソル位置（インライン画像） | InlineImage (45pt) |
| スライド | 選択中のスライド中央 | Image (60pt) |

## ファイル構成

```
gas-hanko-addon/
├── Code.gs           # サーバー側ロジック（メニュー、画像挿入、苗字取得）
├── Sidebar.html      # サイドバーUI（Canvas描画、入力フォーム）
├── appsscript.json   # プロジェクト設定（スコープ、タイムゾーン）
├── README.md         # 本ドキュメント
└── dist/
    └── GASHankoAddon説明書.txt   # 配布用説明書
```

## カスタマイズ

### 苗字マッピングの編集

`Code.gs` 内の `NAME_MAP` にメールアドレスと苗字の対応を追加できます:

```javascript
var NAME_MAP = {
  "taro.yamada@example.com": "山田",
  "hanako.suzuki@example.com": "鈴木",
  // ↓ ここに追加
  "your.email@company.com": "あなたの苗字"
};
```

マッピングにないアドレスは、`@` より前の文字列が自動的にデフォルト値として使われます。

### 印鑑のサイズ・色の変更

`Sidebar.html` 内の定数を編集:

```javascript
var STAMP_SIZE = 200;          // Canvas描画解像度 (px)
var STAMP_COLOR = "#c62828";   // 印鑑の色（赤）
var TRACKING_COLOR = "#aaaaaa"; // トラッキングIDの色
var TRACKING_FONT_SIZE = 8;    // トラッキングIDのサイズ
```

`Code.gs` 内の挿入サイズ:

```javascript
var targetSize = 60;   // スプレッドシート挿入サイズ (px)
var targetPt = 45;     // ドキュメント挿入サイズ (pt)
var targetPt = 60;     // スライド挿入サイズ (pt)
```

## 必要な OAuth スコープ

| スコープ | 用途 |
|---|---|
| `spreadsheets.currentonly` | スプレッドシートへの画像挿入 |
| `documents.currentonly` | ドキュメントへの画像挿入 |
| `presentations.currentonly` | スライドへの画像挿入 |
| `script.container.ui` | サイドバーUI表示 |
| `userinfo.email` | アカウントメール取得（苗字自動入力） |

## テクノロジースタック

- **Google Apps Script (V8)** — サーバー側ロジック
- **HTML Service** — サイドバーUI
- **HTML5 Canvas** — 印鑑画像の動的生成
- **Web Crypto API** — トラッキングIDの乱数生成
- **SpreadsheetApp / DocumentApp / SlidesApp** — 各アプリへの画像挿入

## 関連プロジェクト

- [excel-hanko-addin](https://github.com/highdefinitionaudiodriver/excel-hanko-addin) — Microsoft Office (Excel / Word / PowerPoint) 版

## ライセンス

MIT License - Copyright (c) 2026 highdefinitionaudiodriver

---

## 🤝 商用利用・カスタマイズ依頼

- 個人利用は無料（MIT ライセンス）
- 法人導入支援、カスタマイズ、業務テンプレ整備、追加機能開発は応相談
- 連絡先：highdefinitionaudiodriver@gmail.com
