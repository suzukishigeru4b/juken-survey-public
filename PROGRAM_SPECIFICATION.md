# プログラム仕様書

[![Version](https://img.shields.io/badge/version-2.3.0-blue.svg)](./VERSION_CHANGES.md)
[![Platform](https://img.shields.io/badge/platform-Google%20Apps%20Script-4285F4.svg)](https://developers.google.com/apps-script)
[![For](https://img.shields.io/badge/対象-開発者-red.svg)](#)

> 本ドキュメントは、受験校調査アプリの技術仕様を記載した開発者向けリファレンスです。

---

## 📖 目次

- [システム概要](#1-システム概要)
- [ファイル構成](#2-ファイル構成)
- [スプレッドシート構造](#3-スプレッドシート構造)
- [サーバーサイド仕様](#4-サーバーサイド仕様)
- [クライアントサイド仕様](#5-クライアントサイド仕様)
- [セキュリティと権限](#6-セキュリティと権限)
- [API リファレンス](#7-api-リファレンス)

---

## 1. システム概要

### アーキテクチャ

```mermaid
flowchart TB
    subgraph Client[クライアント]
        Browser[ブラウザ]
        JS[script.html<br/>JavaScript]
        LocalCache[localStorage<br/>大学データキャッシュ]
    end
    
    subgraph GAS[Google Apps Script]
        Main[main.js<br/>サーバーロジック]
        Cache[CacheService<br/>設定・マスタキャッシュ]
    end
    
    subgraph Google[Google Services]
        Sheets[(スプレッドシート)]
        Mail[Gmail API]
        Drive[Drive API]
    end
    
    Browser --> JS
    JS <-->|google.script.run| Main
    JS <--> LocalCache
    Main <--> Cache
    Main <--> Sheets
    Main --> Mail
    Main --> Drive
```

### 技術スタック

| レイヤー | 技術 |
|:---|:---|
| フロントエンド | HTML5, CSS3, Vanilla JavaScript |
| バックエンド | Google Apps Script (V8 Runtime) |
| データベース | Google Spreadsheet |
| 認証 | Google OAuth 2.0 |
| キャッシュ | CacheService (サーバー), localStorage (クライアント) |

---

## 2. ファイル構成

```
juken-survey/
├── main.js              # サーバーサイド処理 (GAS)
├── index.html           # メイン画面レイアウト
├── css.html             # スタイルシート (CSS)
├── script.html          # クライアントサイド処理
└── appsscript.json      # マニフェストファイル
```

### ファイル詳細

| ファイル | 役割 | 備考 |
|:---|:---|:---|
| `main.js` | サーバーサイド処理 | API, DB操作, メール送信, キャッシュ管理 |
| `index.html` | メイン画面レイアウト | HTML構造, 静的コンテンツ |
| `css.html` | スタイルシート | CSS変数, レスポンシブ対応 |
| `script.html` | クライアントサイド処理 | UI操作, 非同期通信, バリデーション |
| `appsscript.json` | プロジェクト設定 | マニフェストファイル |

---

## 3. スプレッドシート構造

システムは単一のスプレッドシートでデータベースと設定を管理します。

### 3.1 シート一覧

| シート名 | 定数名 | 用途 |
|:---|:---|:---|
| **学籍データ** | `STUDENTS` | 生徒マスタ |
| **職員データ** | `TEACHERS` | 教員マスタ |
| **受験校DB** | `JUKEN_DB` | トランザクションデータ |
| **大学データ** | `DAIGAKU` | 大学コードマスタ |
| **試験形態** | `KEITAI` | 選択肢マスタ |
| **合否選択肢** | `SEL_GOUHI` | 選択肢マスタ |
| **受験形態選択肢** | `SEL_KEITAI` | 選択肢マスタ |
| **設定** | `SETTINGS` | システム設定 |
| **調査書交付願** | `PDF_TEMPLATE` | 帳票テンプレート |
| **校内DB用** | `KOUNAI_DB` | エクスポート用 |

#### 設定 (`SETTINGS`)

| 行 | カラム名 | 説明 | 備考 |
|:---:|:---|:---|:---|
| 1 | ページタイトル | アプリのタイトル | |
| 2 | 最大登録件数 | 生徒1人あたりの最大件数 | |
| 3 | 入力許可 | 生徒の入力を制御 | `TRUE` / `FALSE` |
| 4 | 大学シリアル | 大学データのキャッシュシリアル番号 | 更新するとクライアントキャッシュを強制更新 |
| 5 | メール件名 | PDF送信時のメール件名 | |
| 6 | メール本文 | PDF送信時のメール本文 | |

### 3.2 主要シートのカラム定義

#### 学籍データ (`STUDENTS`)

| 列 | カラム名 | 型 | 必須 |
|:---:|:---|:---|:---:|
| A | メールアドレス | String | ✅ |
| B | 学年 | Number | ✅ |
| C | クラス | String | ✅ |
| D | 出席番号 | Number | ✅ |
| E | 氏名 | String | ✅ |

#### 受験校DB (`JUKEN_DB`)

| 列 | カラム名 | 型 | 備考 |
|:---:|:---|:---|:---|
| A | タイムスタンプ | Date | 自動設定 |
| B | メールアドレス | String | 学籍データと紐付け |
| C | 削除フラグ | Boolean | 論理削除用 |
| D | 大学コード | String | 大学データと紐付け |
| E | 大学名 | String | |
| F | 合否 | String | 選択肢マスタ参照 |
| G | 受験形態 | String | 選択肢マスタ参照 |
| H | 進学 | Boolean | |

### 3.3 ER図

```mermaid
erDiagram
    STUDENTS ||--o{ JUKEN_DB : "メールアドレス"
    DAIGAKU ||--o{ JUKEN_DB : "大学コード"
    SEL_KEITAI ||--o{ JUKEN_DB : "試験形態"
    SEL_GOUHI ||--o{ JUKEN_DB : "合否"
    
    STUDENTS {
        string email PK
        int grade
        string class
        int number
        string name
    }
    
    JUKEN_DB {
        date timestamp
        string email FK
        boolean deleted
        string universityCode FK
        string universityName
        string result FK
        string examType FK
        boolean enrollment
    }
    
    DAIGAKU {
        string code PK
        string name
    }
```

---

## 4. サーバーサイド仕様

> **ファイル**: `main.js`

### 4.1 エントリポイント・初期化

#### `doGet(e)`

Webアプリへのアクセス時に呼び出されるエントリポイント。

```javascript
function doGet(e) {
  const settings = getSettings();
  const htmlTemplate = HtmlService.createTemplateFromFile('index');
  const html = htmlTemplate.evaluate();
  html.setTitle(settings.pageTitle);
  return html;
}
```

| 項目 | 内容 |
|:---|:---|
| **戻り値** | `HtmlOutput` |
| **セキュリティ** | X-Frame-Options: SAMEORIGIN (デフォルト) |

---

#### `getInitialData()`

クライアント起動時に必要なデータを一括取得。

```mermaid
sequenceDiagram
    participant C as Client
    participant S as Server
    participant Cache as CacheService
    participant SS as Spreadsheet
    
    C->>S: getInitialData()
    S->>Cache: キャッシュ確認
    alt キャッシュヒット
        Cache-->>S: キャッシュデータ
    else キャッシュミス
        S->>SS: batchGet (複数シート)
        SS-->>S: データ返却
        S->>Cache: キャッシュ保存
    end
    S->>S: ユーザー判定
    S-->>C: JSON Response
```

| 項目 | 内容 |
|:---|:---|
| **戻り値** | `Object` (ユーザーロールに応じたJSON) |
| **取得内容** | 選択肢マスタ, ユーザー情報, 設定値 |
| **ユーザーロール** | `student` / `teacher` / `guest` |
| **キャッシュ** | CacheService (6時間) |

---

### 4.2 キャッシュ戦略

#### キャッシュ対象シート

```javascript
const CACHE_TARGET_SHEETS = [
  SHEET_NAMES.SETTINGS,
  SHEET_NAMES.SEL_GOUHI,
  SHEET_NAMES.SEL_KEITAI,
  SHEET_NAMES.TEACHERS,
  SHEET_NAMES.STUDENTS
];
```

#### キャッシュ更新トリガー

| トリガー | 動作 |
|:---|:---|
| `onEdit` | セル編集時に自動キャッシュ更新 |
| `onChange` | シート構造変更時に自動キャッシュ更新 |
| 手動 | `warmUpCache()` / `warmUpAllCache()` |

> [!IMPORTANT]
> `onEdit` トリガーは **インストーラブルトリガー** として設定する必要があります。

#### キャッシュ関連関数

| 関数名 | 説明 |
|:---|:---|
| `getSheetDataApiWithCache(sheetName)` | キャッシュ付きシートデータ取得 |
| `getBatchSheetDataWithCache(requests)` | 複数シートの一括取得（キャッシュ対応） |
| `warmUpCache(sheetName)` | 指定シートのキャッシュ強制更新 |
| `warmUpAllCache()` | 全対象シートのキャッシュ更新 |
| `checkAndUpdateCache(sheetName)` | トリガー経由でのキャッシュ更新判定 |
| `setupTriggers()` | キャッシュ更新トリガーの一括設定 |

---

### 4.3 データ操作

#### `saveExamDataList(strJuken, mailAddr)`

受験データを保存する。

| パラメータ | 型 | 説明 |
|:---|:---|:---|
| `strJuken` | `String` | JSON文字列 (受験データ配列) |
| `mailAddr` | `String` | 対象のメールアドレス |

**処理フロー:**

```mermaid
flowchart TD
    A[データ受信] --> B{権限チェック}
    B -->|OK| C[ロック取得]
    B -->|NG| X[エラー返却]
    C --> D{バリデーション}
    D -->|OK| E[既存データ取得<br/>batchGet]
    D -->|NG| X
    E --> F[差分計算]
    F --> G[batchUpdate<br/>更新・削除]
    G --> H[append<br/>新規追加]
    H --> I[ロック解放]
    I --> J[最新データ返却]
```

| 機能 | 説明 |
|:---|:---|
| **排他制御** | `LockService` で同時書き込みを防止 (待機時間10秒) |
| **バリデーション** | 入力件数、選択肢の正当性チェック (設定・マスタはキャッシュ利用) |
| **更新方式** | 差分更新（変更行のみ `batchUpdate`） |
| **追加方式** | 新規データは `append` で一括追加 |
| **削除方式** | 論理削除（削除フラグ） |
| **重複対策** | 同一大学コードの重複レコードを自動的に論理削除 |

---

#### `getExamDataList(mailAddr)`

指定されたメールアドレスの受験データを取得。

| パラメータ | 型 | 説明 |
|:---|:---|:---|
| `mailAddr` | `String` | 対象のメールアドレス |

| 項目 | 内容 |
|:---|:---|
| **戻り値** | `Array` (受験データ配列、ヘッダー含む) |
| **フィルタ** | 削除フラグ `TRUE` のデータは除外 |
| **API** | `Sheets.Spreadsheets.Values.get()` |

---

### 4.4 帳票・メール

#### `sendPdf(mailAddr)`

調査書交付願PDFを生成してメール送信。

```mermaid
sequenceDiagram
    participant C as Client
    participant S as Server
    participant SS as Spreadsheet
    participant D as Drive
    participant M as Gmail
    
    C->>S: sendPdf(mailAddr)
    S->>SS: テンプレート取得
    S->>SS: 一時シート作成
    S->>SS: データ埋め込み
    S->>D: PDF変換
    D-->>S: PDF Blob
    S->>M: メール送信 (添付)
    S->>D: 一時ファイル削除
    S-->>C: 完了通知
```

---

### 4.5 管理・ユーティリティ

| 関数名 | 説明 |
|:---|:---|
| `createData()` | 校内DB用データを生成 |
| `deleteMarkedRows()` | 論理削除レコードを物理削除 |
| `queryDeleteMarkedRows()` | 削除前に確認ダイアログを表示 |
| `importUniversityData()` | Benesseデータをインポート |
| `clearUniversityData()` | 大学データシートをクリア |
| `getUniversityDataApi()` | 大学コードマスタを取得（キャッシュなし） |
| `warmUpAllCache()` | 全キャッシュを手動更新（メニューから実行可） |
| `setupTriggers()` | キャッシュ更新トリガーを自動設定（メニューから実行可） |

---

## 5. クライアントサイド仕様

> **ファイル**: `script.html`

### 5.1 初期化・状態管理

#### `pageLoaded()`

ページ読み込み時の初期化処理。

```javascript
async function pageLoaded() {
  cacheDomElements();
  showLoading();
  try {
    const data = await new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject)
        .getInitialData();
    });
    initializeUI(data);
  } finally {
    hideLoading();
  }
}
```

#### `cacheDomElements()`

頻繁にアクセスするDOM要素をキャッシュ。

| グローバル変数 | 型 | 用途 |
|:---|:---|:---|
| `dom` | `Object` | DOM要素の参照を格納 |

---

### 5.2 UI操作

#### `setInputTable()`

受験データに基づき入力フォームを動的生成。

| 特徴 | 説明 |
|:---|:---|
| **テンプレート使用** | `<template id="examRowTemplate">` |
| **パフォーマンス** | DOM生成コストを低減 |

---

#### 大学検索機能

| 関数 | 説明 |
|:---|:---|
| `searchUniversityCode(row)` | 検索ダイアログを表示 |
| `searchKeyword()` | インクリメンタルサーチを実行 |
| `getUniversityDataList()` | 大学データを取得（キャッシュ対応） |

**キャッシュ仕様:**

| 項目 | 内容 |
|:---|:---|
| 保存先 | `localStorage` |
| キー | `universityDataCache` |
| タイムスタンプ | `universityDataTimestamp` |
| 有効期限 | 24時間 |

---

### 5.3 データ送信

#### `sendExamData()`

フォームの入力値をサーバーに送信。

```mermaid
flowchart TD
    A[送信ボタンクリック] --> B{変更あり?}
    B -->|No| C[処理終了]
    B -->|Yes| D[確認ダイアログ]
    D --> E{ユーザー確認}
    E -->|キャンセル| C
    E -->|OK| F[ローディング表示]
    F --> G[フォーム値を収集]
    G --> H[JSON変換]
    H --> I[サーバー送信]
    I --> J{成功?}
    J -->|Yes| K[データ更新・成功表示]
    J -->|No| L[エラーメッセージ]
```

---

## 6. セキュリティと権限

### 認証・認可

| 項目 | 説明 |
|:---|:---|
| **認証** | Google OAuth 2.0 |
| **ユーザー識別** | `Session.getActiveUser().getEmail()` |
| **権限チェック** | `isValidUser()` 関数 |

### アクセス制御マトリクス

| 操作 | 生徒 | 教員 | ゲスト |
|:---|:---:|:---:|:---:|
| 自分のデータ閲覧 | ✅ | ✅ | ❌ |
| 自分のデータ編集 | ✅ | ✅ | ❌ |
| 他生徒のデータ閲覧 | ❌ | ✅ | ❌ |
| 他生徒のデータ編集 | ❌ | ✅ | ❌ |
| PDF発行 | ✅ | ✅ | ❌ |

---

## 7. API リファレンス

### クライアント → サーバー

| 関数名 | 引数 | 戻り値 | 説明 |
|:---|:---|:---|:---|
| `getInitialData()` | なし | `Object` | 初期データ取得 (キャッシュ活用) |
| `getExamDataList(mailAddr)` | `String` | `Array` | 受験データ取得 |
| `sendExamData(strJuken, mailAddr)` | `String`, `String` | `Object` | 受験データ保存 |
| `sendPdf(mailAddr)` | `String` | `Object` | PDF発行 |
| `getUniversityDataList()` | なし | `Array` | 大学データ取得 |
| `getStudentsList()` | なし | `Array` | 生徒一覧取得 (教員用) |

### レスポンス形式

```javascript
// 成功時
{
  "success": true,
  "message": "保存しました",
  "data": { ... }
}

// エラー時
{
  "success": false,
  "message": "エラーメッセージ",
  "error": "詳細情報"
}
```

---

## 8. パフォーマンス最適化

### 8.1 キャッシュ戦略まとめ

| データ種別 | キャッシュ場所 | 有効期限 | 更新方法 |
|:---|:---|:---|:---|
| 設定・選択肢 | CacheService | 6時間 | トリガー自動更新 |
| 学籍・職員 | CacheService | 6時間 | トリガー自動更新 |
| 大学データ | localStorage | 24時間 | `大学シリアル` 更新時に強制更新 |
| 受験データ | キャッシュなし | - | 毎回取得 |

### 8.2 API呼び出し最適化

| 最適化項目 | 手法 |
|:---|:---|
| 初期データ取得 | `batchGet` による一括取得 |
| 受験データ保存 | `batchUpdate` + `append` |
| 大学データ検索 | クライアント側でフィルタリング |

---

<div align="center">

📚 **関連ドキュメント**

[README](./README.md) ｜ [変更履歴](./VERSION_CHANGES.md) ｜ [教員マニュアル](./TEACHER_MANUAL.md)

</div>
