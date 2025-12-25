# バージョン変更履歴 (VERSION_CHANGES.md)

## Ver 2.0.0 (2025年12月)

### 主な変更点

#### 1. パフォーマンス最適化 (Google Sheets API 対応)
- **初期データ取得 (`getInitialData`)**: 従来の `getValues()` から `Sheets.Spreadsheets.Values.batchGet()` に変更し、複数シートのデータを1回のAPI呼び出しで取得するように改善しました。
- **受験データ取得 (`getExamDataList`)**: Google Sheets API (`Sheets.Spreadsheets.Values.get()`) を使用するようになりました。
- **受験データ保存 (`saveExamDataList`)**: 大規模リファクタリングを実施。
    - **検索処理**: 従来の全件取得 (`getDataRange`) を廃止し、対象ユーザーのみをAPI経由で検索・取得する方式 (`batchGet`) に変更。
    - **書き込み処理**: 個別の `setValue/setValues` を廃止し、`batchUpdate` (更新/削除) と `append` (追加) に統合。これによりスケーラビリティが大幅に向上しました。
    - **戻り値追加**: 保存完了後に最新の受験データリストを直接返すようになり、クライアント側での再取得が不要になりました。

#### 2. iFrame対応
- `doGet` 関数内に `html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)` を追加し、iFrame 内での表示を可能にしました。

#### 3. キャッシュ戦略の変更
- **サーバーサイドキャッシュ廃止**: v1.0.0 で実装されていた `CacheService` によるサーバーサイドキャッシュ関連コード (`putCache`, `getCache`, `clearCache`, `getSheetData` のキャッシュ処理) を廃止しました。キャッシュ管理の複雑さとAPIパフォーマンス向上のトレードオフとして決定されました。
- **クライアントサイドキャッシュ追加**: `script.html` において、大学データを `localStorage` に24時間キャッシュする機能を追加しました (`getUniversityDataList` 関数内)。これにより、大学検索時のロード時間を短縮しています。

#### 4. 教員モード改善
- `getStudentsList` 関数を新規追加し、教員モード時に生徒一覧を非同期で取得するように変更しました。初期ページロードの高速化に貢献しています。
- `getFilteredUniversityDataList` 関数を追加し、大学データのサーバーサイド検索・ページネーションをサポートしました（将来拡張用、現時点ではクライアント検索がデフォルト）。

#### 5. クライアントサイド改善 (`script.html`)
- **DOMキャッシュ導入**: `dom` オブジェクトにDOM要素をキャッシュし、`document.getElementById()` の呼び出し回数を削減しました。
- **is-hiddenクラス**: 要素の表示/非表示制御を `style.display` から `.is-hidden` クラスのトグルに変更し、CSS との一貫性を向上させました。
- **プログレスダイアログ改善**: よりユーザーフレンドリーなスピナー表示に変更（ようこそメッセージの表示など）。
- **データ補正ヘルパー追加**: `padRowsWithHeader` 関数を追加し、Sheets API からのデータ（行末の空セルが欠損する可能性）を安全に扱うようになりました。
- **Null安全対策**: 各種データ処理において `|| []` や `|| {}` による初期値設定を徹底し、予期せぬ `undefined` エラーを防止しました。

#### 6. カスタムメニュー変更
- `onOpen` 関数内のカスタムメニュー項目を改訂しました。
    - 「キャッシュクリア」メニューは廃止。
    - 「削除レコード完全削除」の呼び出し関数名を `deleteMarkedRows` から `queryDeleteMarkedRows` に変更（ダイアログ確認を追加したため）。

---

## Ver 1.0.0 (初期リリース)

- 基本的な受験校調査機能を実装。
- サーバーサイドキャッシュ (`CacheService`) を採用。
- 従来のGAS組み込み関数 (`getValues`, `setValues`) を使用。
