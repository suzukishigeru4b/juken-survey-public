# 受験校調査アプリ (Preferred School Survey System)

[![Version](https://img.shields.io/badge/version-2.3.0-blue.svg)](./VERSION_CHANGES.md)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Google%20Apps%20Script-4285F4.svg)](https://developers.google.com/apps-script)

Google Apps Script (GAS) と Google スプレッドシートを使用した、**受験校調査および調査書発行願作成システム**です。

---

## 📖 目次

- [概要](#概要)
- [v2.3.0 新機能ハイライト](#v230-新機能ハイライト)
- [機能一覧](#機能一覧)
- [システム要件](#システム要件)
- [セットアップ手順](#セットアップ手順)
- [ファイル構成](#ファイル構成)
- [ドキュメント](#ドキュメント)
- [ライセンス](#ライセンス)

---

## 概要

生徒が自身の受験予定校や合否結果を入力し、教員がその状況を一元管理するためのWebアプリケーションです。

**主な特徴:**

- 📱 **マルチデバイス対応** - PC、タブレット、スマートフォンで快適操作
- 🌙 **ダークモード** - OSの設定に連動した自動切り替え
- 🔍 **大学検索機能** - Benesseデータベース対応のインクリメンタルサーチ
- 📄 **PDF自動発行** - 調査書交付願をワンクリックでメール送信
- 🔒 **アクセス制御** - 生徒・教員・管理者の権限管理
- ⚡ **高速化** - サーバー・クライアント両面でのキャッシュ戦略

---

## ✨ v2.3.0 新機能ハイライト

> **2025年12月30日リリース** - データ取得の一貫性向上とキャッシュ管理の自動化！

| 改善項目 | 詳細 |
|:---|:---|
| 🔄 **startRowパラメータ廃止** | すべてのシート取得を1行目から統一し、コードの一貫性と保守性を向上 |
| ⚡ **自動トリガー設定** | メニューから1クリックでキャッシュ更新トリガーを設定できる機能を追加 |
| 📊 **管理者メニュー刷新** | 「校内DB用」から「管理者メニュー」に改名し、キャッシュ管理機能を追加 |
| 🔄 **定期キャッシュ更新** | 1時間ごとに全キャッシュを自動更新するタイムベーストリガーを実装 |
| 🛡️ **キャッシュキー簡素化** | `sheetName_startRow`から`sheetName`に変更し、管理が容易に |
| 📝 **ドキュメント刷新** | すべてのマニュアルを最新の実装に合わせて更新 |

詳しくは [VERSION_CHANGES.md](./VERSION_CHANGES.md) をご覧ください。

---

## 機能一覧

### 👨‍🎓 生徒用機能

| 機能 | 説明 |
|:---|:---|
| 受験校入力 | 志望大学、学部、試験形態、合否結果、進学先を入力・保存 |
| 大学検索 | コードまたは名称でインクリメンタルサーチ (localStorage キャッシュ) |
| 調査書交付願 | 入力データに基づきPDFを作成し、メール送信 |

### 👨‍🏫 教員用機能

| 機能 | 説明 |
|:---|:---|
| 生徒データ閲覧 | クラス・番号を選択して入力状況を確認 |
| 代理入力 | 必要に応じて生徒のデータを編集 |

### ⚙️ 管理者用機能

| 機能 | 説明 |
|:---|:---|
| マスタ管理 | 生徒・職員・大学・選択肢データの管理 |
| 設定変更 | タイトル、入力許可期間、メール文面の制御 |
| データエクスポート | 校内DB用データの出力 |
| Benesseインポート | Benesse形式大学データの一括取り込み |
| キャッシュ管理 | 自動更新トリガーの設定と手動更新機能 |

---

## システム要件

### 必須要件

- ✅ **Google アカウント** (Google Workspace 推奨)
- ✅ **Google Chrome** (最新版)

### 推奨環境

- Google Workspace for Education
- スプレッドシートの編集権限
- Apps Script の高度なサービス有効化 (Google Sheets API)

---

## セットアップ手順

### 1. スプレッドシートの準備

以下のシートを作成してください（**名称変更不可**）:

| シート名 | 用途 |
|:---|:---|
| `学籍データ` | 生徒のメールアドレス、氏名、クラスなど |
| `職員データ` | 教員のメールアドレス、氏名 |
| `受験校DB` | 入力された受験データの蓄積先 |
| `大学データ` | 大学コードマスタ |
| `試験形態` | ドロップダウン選択肢 |
| `合否選択肢` | ドロップダウン選択肢 |
| `受験形態選択肢` | ドロップダウン選択肢 |
| `設定` | システム全体の設定 |
| `調査書交付願` | PDF出力用テンプレート |
| `校内DB用` | データエクスポート用一時シート |

### 2. スクリプトのデプロイ

1. スプレッドシートの **拡張機能** > **Apps Script** を開く
2. 各ファイル (`main.js`, `index.html`, `css.html`, `script.html`) をコピー
3. **サービス** から `Google Sheets API` を追加
4. **デプロイ** > **新しいデプロイ** を選択
5. 以下の設定でデプロイ:

   | 項目 | 設定値 |
   |:---|:---|
   | 種類 | ウェブアプリ |
   | 次のユーザーとして実行 | `ウェブアプリケーションにアクセスしているユーザー` |
   | アクセスできるユーザー | `ドメイン内の全員` または `全員` |

6. 発行されたURLを共有

### 3. トリガーの設定（推奨）

**自動キャッシュ更新のためにトリガーを設定してください:**

#### 方法1: メニューから自動設定（推奨）

1. スプレッドシートを開く
2. 上部メニューの **「管理者メニュー」** → **「キャッシュ更新トリガー設定」** を選択
3. 確認メッセージが表示されたら完了

これで以下のトリガーが自動的に設定されます：
- `onEdit` - 編集時に対象シートのキャッシュを更新
- `onChange` - 変更時に対象シートのキャッシュを更新
- `warmUpAllCache` - 1時間ごとに全キャッシュを更新

#### 方法2: 手動設定（上級者向け）

1. Apps Script エディタで **トリガー** アイコンをクリック
2. 以下のトリガーを手動で追加:

   | 関数名 | イベントタイプ | イベントソース |
   |:---|:---|:---|
   | `onEdit` | 編集時 | スプレッドシートから |
   | `onChange` | 変更時 | スプレッドシートから |
   | `warmUpAllCache` | 時間主導型 | 1時間ごと |

> [!TIP]
> これらのトリガーにより、設定や選択肢マスタの変更が即座にキャッシュに反映されます。

### 4. Benesseデータのインポート（オプション）

> [!IMPORTANT]
> Benesseの大学データCSVファイルは **ShiftJIS** でエンコードされています。インポート前に **UTF-8** に変換してください。

**変換手順:**

1. CSVファイルをメモ帳などで開く
2. **名前を付けて保存** > **エンコード: UTF-8** を選択して保存
3. スプレッドシートにインポート後、`importUniversityData()` を実行

---

## ファイル構成

```
juken-survey/
├── main.js              # サーバーサイドロジック (GAS)
├── index.html           # メイン画面のHTML構造
├── css.html             # スタイルシート (CSS)
├── script.html          # クライアントサイドロジック (JavaScript)
├── appsscript.json      # GASマニフェストファイル
├── README.md            # 本ドキュメント
├── PROGRAM_SPECIFICATION.md  # プログラム詳細仕様
├── TEACHER_MANUAL.md    # 教員向け操作マニュアル
├── STUDENT_MANUAL.md    # 生徒向け操作マニュアル
└── VERSION_CHANGES.md   # バージョン変更履歴
```

---

## ドキュメント

| ドキュメント | 対象 | 内容 |
|:---|:---|:---|
| [PROGRAM_SPECIFICATION.md](./PROGRAM_SPECIFICATION.md) | 開発者 | システム設計・API仕様・キャッシュ戦略 |
| [TEACHER_MANUAL.md](./TEACHER_MANUAL.md) | 教員・管理者 | 操作手順・管理機能の使い方 |
| [STUDENT_MANUAL.md](./STUDENT_MANUAL.md) | 生徒 | 受験校入力・調査書発行の手順 |
| [VERSION_CHANGES.md](./VERSION_CHANGES.md) | 全員 | v2.3.0の変更点一覧 |

---

## パフォーマンス最適化

### キャッシュ戦略

| データ種別 | キャッシュ場所 | 有効期限 | 特徴 |
|:---|:---|:---|:---|
| 設定・選択肢 | CacheService | 6時間 | 自動更新 (トリガー) |
| 学籍・職員 | CacheService | 6時間 | 自動更新 (トリガー) |
| 大学データ | localStorage | 24時間 | クライアント側 |

### API呼び出し最適化

- **初期データ取得**: `batchGet` による複数シート一括取得
- **受験データ保存**: `batchUpdate` + `append` による効率的な書き込み
- **キャッシュヒット率**: トリガーによる自動更新で常に最新状態を維持

---

## トラブルシューティング

### キャッシュ関連

| 問題 | 対処法 |
|:---|:---|
| 設定変更が反映されない | トリガーが正しく設定されているか確認 |
| 大学検索が遅い | ブラウザの localStorage をクリア |
| データが古い | `warmUpAllCache()` を手動実行 |

---

## ライセンス

[MIT License](https://opensource.org/licenses/MIT)

```
Copyright (c) 2025 Shigeru Suzuki

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
```

---

<div align="center">

**Made for Education with ❤️**

*v2.3.0 - Unified Data Retrieval & Automated Cache Management*

</div>
