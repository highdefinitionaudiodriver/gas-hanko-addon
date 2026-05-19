# Google Workspace Marketplace 提出用テンプレート（日本語）

GAS Hanko Add-on を Google Workspace Marketplace に提出する際に使用する説明文・スクリーンショット要件・メタデータの雛形です。

---

## アプリ名

```
電子印鑑 for Google Workspace
```

英語版：`Hanko Add-on - Japanese Digital Stamp for Workspace`

---

## 短い説明（最大 80 文字）

```
Sheets / Docs / Slides にワンクリックで日本式電子印鑑（データ印）を挿入
```

---

## 詳細説明（最大 16,000 文字）

```
■ 概要
日本のビジネスシーンで使われる「データ印（電子印鑑）」を、Google スプレッドシート・
ドキュメント・スライドにワンクリックで挿入できる Google Workspace アドオンです。
社内承認フローを Google Workspace 上で完結させ、押印のために出社・郵送する
非効率を解消します。

■ こんな方におすすめ
・Google Workspace を導入しているが日本式押印文化が残っている中小企業
・在宅勤務環境の総務・経理・人事担当
・脱ハンコと業務効率化を同時に進めたい組織

■ 主な機能
・3アプリ対応 — Sheets / Docs / Slides で共通利用
・Canvas で動的生成 — リアルタイムプレビュー付き作業ウィンドウ
・3段構成（苗字・日付・部署名）の丸型・赤色データ印
・Google アカウント連携 — メールアドレスから苗字を自動入力
・トラッキング ID — 押印ごとに 8 桁のユニーク ID を印影に埋込（簡易証跡）
・インストール不要 — ブラウザだけで動作。Node.js・ローカルサーバー不要

■ 動作環境
・Google アカウント（個人・Workspace どちらも可）
・対応ブラウザ：Chrome / Edge / Safari / Firefox 最新版

■ プライバシー・セキュリティ
・氏名・部署名等の入力データは Google Apps Script の実行コンテキスト内で
　のみ処理され、外部第三者のサーバーへ送信されることはありません。
・印影 PNG の生成は Canvas API でブラウザ内完結。
・Google アカウント情報の利用は OAuth スコープ最小限の範囲。

■ ライセンス
個人利用は無料（MIT ライセンス）。
法人向けのカスタマイズ・証跡 CSV エクスポート・SSO 連携は別途ご相談ください。

■ サポート
GitHub: https://github.com/highdefinitionaudiodriver/gas-hanko-addon
お問い合わせ: highdefinitionaudiodriver@gmail.com
```

---

## カテゴリ

- 主：**Business Tools**
- 副：**Productivity**

---

## キーワード

```
hanko, 電子印鑑, デジタル印鑑, ハンコ, 脱ハンコ, 日本式印鑑, 押印,
approval, workflow, japanese stamp, digital signature
```

---

## OAuth スコープ（最小限）

```
https://www.googleapis.com/auth/spreadsheets.currentonly
https://www.googleapis.com/auth/documents.currentonly
https://www.googleapis.com/auth/presentations.currentonly
https://www.googleapis.com/auth/userinfo.email
```

`.currentonly` を使用することで、現在開いているドキュメントのみへのアクセスに限定 → 審査が通りやすい。

---

## スクリーンショット要件

Google Workspace Marketplace 提出には **5 枚以上** のスクリーンショットが必要です。

1. **スプレッドシートでの押印例**（セルに印影が挿入された画面）
2. **ドキュメントでの押印例**（本文末尾の押印箇所）
3. **スライドでの押印例**（プレゼン資料への押印）
4. **サイドバーのフォーム入力中**（苗字・日付・部署名）
5. **印影プレビュー拡大**（トラッキング ID も視認できる状態）

推奨サイズ：**1280×800**、PNG または JPG。

---

## アイコン

- **96×96 PNG**（アドオンメニュー用、透過背景）
- **128×128 PNG**（Marketplace 一覧用）
- **220×140 PNG**（バナー用）

`appsscript.json` の `addOns.common.logoUrl` に Google Drive 公開 URL or GitHub Pages URL を指定。

---

## サポート URL / プライバシーポリシー URL

GitHub Pages を有効化して以下を準備：

- サポート URL：`https://highdefinitionaudiodriver.github.io/gas-hanko-addon/support`
- プライバシーポリシー URL：`https://highdefinitionaudiodriver.github.io/gas-hanko-addon/privacy`
- 利用規約 URL：`https://highdefinitionaudiodriver.github.io/gas-hanko-addon/terms`

（テンプレート HTML は `docs/store-listing/pages/` に配置予定）

---

## 提出時のチェックリスト

- [ ] Google Cloud Platform プロジェクトを作成し、OAuth 同意画面を構成
- [ ] OAuth スコープを `.currentonly` 系に限定（**機密スコープ審査を回避**）
- [ ] `appsscript.json` の `addOns.common.name` / `logoUrl` を本番値に
- [ ] プライバシーポリシー URL が有効でアクセス可能
- [ ] スクリーンショットを Google Cloud Console にアップロード
- [ ] **Google Cloud Console での「OAuth 検証」が完了するまで配布できない**（数日〜数週間）
- [ ] 認定されると Marketplace に公開リスティング可能
