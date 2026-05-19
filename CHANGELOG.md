# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- README に「これは何？（30秒で）」「想定ユースケース・価格帯」セクションを追加
- SECURITY.md を追加（脆弱性報告フロー）
- 商用利用・カスタマイズ依頼の連絡先を README 末尾に明記
- **証跡（押印履歴）機能**：押印ごとに `timestamp / trackingId / host_app / user_email / 上中下段テキスト` を UserProperties に永続化（Google アカウントごと）
- **CSV エクスポート機能**：サーバ側で CSV を生成 → クライアントが Blob ダウンロード（`hanko_audit_YYYYMMDD_HHMMSS.csv`）
- 履歴クリアボタン（誤操作防止の確認ダイアログ付き）
- UserProperties 容量上限（9KB/key）に近づいた場合は古いエントリを自動削除＋警告表示
- docs/store-listing/MARKETPLACE_JA.md — Google Workspace Marketplace 提出用テンプレート
- **印影バリエーション 3 形状**（丸 / 角 / スタンプ）
  - サイドバーに 3 形状のラジオセレクタを追加
  - drawStamp() を形状ディスパッチに刷新
  - 角印は正方形外枠 + 水平 3 分割
  - スタンプ風は角丸長方形（横長）で部署印・会社印を想定
  - 証跡 CSV に `shape` 列を追加（UserProperties + クライアント側）

## [0.1.0]

### Added
- 初版リリース
