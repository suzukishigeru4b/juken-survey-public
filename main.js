/**
 * 受験校調査システム (Preferred school survey system)
 * Version 2.4.0
 * * Copyright (c) 2026 Shigeru Suzuki
 * * Released under the MIT License.
 * https://opensource.org/licenses/MIT
 * * [免責事項]
 * 本ソフトウェアの使用により生じた、いかなる損害（データの損失、業務の停止、
 * 授業運営への支障などを含むがこれに限らない）について、著作者は一切の責任を負いません。
 * 利用者は自己の責任において本ソフトウェアを使用するものとします。
 */

//================================================================================
// 1. 定数・グローバル変数定義
//================================================================================
// const appUrl = ScriptApp.getService().getUrl();
// const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
// const myMailAddress = Session.getActiveUser().getEmail();

// シート名定義
// シートデータ定義
const DATA_SHEETS = {
  STUDENTS: { SHEET: '学籍データ', COLS: 6 }, // A-F
  TEACHERS: { SHEET: '職員データ', COLS: 2 }, // A-B
  JUKEN_DB: { SHEET: '受験校DB', COLS: 8 }, // A-H
  DAIGAKU: { SHEET: '大学データ', COLS: 8 }, // A-H (PDF作成用・全データ)
  DAIGAKU_SEARCH: { SHEET: '大学データ', COLS: 2 }, // A-B (検索用・キャッシュ用 軽量データ)
  SEL_GOUHI: { SHEET: '合否選択肢', COLS: 1 }, // A
  SEL_KEITAI: { SHEET: '受験形態選択肢', COLS: 1 }, // A
  SETTINGS: { SHEET: '設定', COLS: 2 }, // A-B
  PDF_TEMPLATE: { SHEET: '調査書交付願', COLS: 0 },
  KOUNAI_DB: { SHEET: '校内DB用', COLS: 9 }, // A-I
  ERROR_LOG: { SHEET: 'エラーログ', COLS: 5 }  // A-E
};

// 受験データ列定義
const EXAM_DATA = {
  TIME_STAMP: 0,
  MAIL_ADDR: 1,
  DEL_FLAG: 2,
  DAIGAKU_CODE: 3,
  DAIGAKU_NAME: 4,
  GOUHI: 5,
  KEITAI: 6,
  SHINGAKU: 7
};

// 生徒データ列定義
const STUDENT_DATA = {
  MAIL_ADDR: 0,
  GRADE: 1,
  CLASS: 2,
  NUMBER: 3,
  NAME: 4,
  REG_COUNT: 5
};

// 教員データ列定義
const TEACHER_DATA = {
  MAIL_ADDR: 0,
  NAME: 1
};

// Benesseデータ列定義 (インポート用)
const BENESSE_DATA = {
  CODE_DAIGAKU: 0, // Benesseデータをこのカラムのヘッダでチェック
  CODE_GAKUBU: 1,  // Benesseデータをこのカラムのヘッダでチェック
  CODE_GAKKA: 2,
  CODE_NITTEI: 3,
  CODE_HOUSHIKI: 4,
  NAME_DAIGAKU: 6,
  NAME_GAKUBU: 7,
  NAME_GAKKA: 8,
  NAME_NITTEI: 9,
  NAME_HOUSHIKI: 10,
  DAY_SHIMEKIRI_WEB: 36,
  DAY_SHIMEKIRI_YUBIN: 37,
  DAY_SHIMEKIRI_MADOGUCHI: 38,
  DAY_NYUSHI: 40, // Benesseデータがカレンダーかどうかをこのカラムのヘッダでチェック
  DAY_HAPPYOU: 41,
  DAY_TETSUZUKI: 42,
};

//================================================================================
// 1. ヘルパー関数
//================================================================================
// 文字列をブール値に変換する(Sheet APIでbooleanの戻り値は文字列なので)
const isTrue = (val) => {
  return String(val).toUpperCase() === 'TRUE';
}

// Benesseインポート用：基準年度計算
// 4月〜12月は同年、1月〜3月は前年を基準とする（年度の開始年）
function calculateBaseYear(date) {
  const year = date.getFullYear();
  const month = date.getMonth(); // 0-11
  // 4月(3)より後は同年、それ以外(1-3月)は前年
  return (month > 3) ? year : year - 1;
}

// Benesseインポート用：日付フォーマット変換
// Benesseインポート用：日付フォーマット変換
function formatBenesseDate(mmdd, baseYear) {
  if (!mmdd || String(mmdd) === "0000" || String(mmdd).length !== 4) {
    return '';
  }
  const str = String(mmdd);
  const month = parseInt(str.substring(0, 2), 10);
  const day = parseInt(str.substring(2), 10);

  // 4月〜12月は基準年度、1月〜3月は翌年
  const year = (month > 3) ? baseYear : baseYear + 1;
  return `${year}/${month}/${day}`;
}

/* カラム数から列文字(A, B, ...)を取得する簡易ヘルパー */
function getColLetter(colIndex) {
  // 簡易実装: 1=A, 2=B, ... 26=Z
  const letters = ['', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
  return letters[colIndex] || '';
}

/**
 * シート定義からAPIリクエスト情報を生成
 * @param {Object} sheetDef - DATA_SHEETS のエントリ
 * @returns {Object} { id: spreadsheetId, range: rangeString }
 */
function makeSheetApiRequest(sheetDef) {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const id = activeSpreadsheet.getId();
  const name = sheetDef.SHEET;
  const cols = sheetDef.COLS;

  let range = `'${name}'`;
  if (cols > 0) {
    const letter = getColLetter(cols);
    range += `!A:${letter}`;
  }
  return { id: id, range: range };
}

//================================================================================
// 2. トリガー・エントリポイント (Triggers)
//================================================================================
// Webアプリとしてのエントリポイント (doGet)
function doGet(e) {
  const settings = getSettings();
  const htmlTemplate = HtmlService.createTemplateFromFile('index');
  const html = htmlTemplate.evaluate();
  html.setTitle(settings.pageTitle);
  //html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  //html.addMetaTag('viewport', 'width=device-width, user-scalable=no, initial-scale=1');
  return html;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// スプレッドシートを開いたときのメニュー追加 (onOpen)
function onOpen() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const entries = [{
    name: "校内DB用データ生成",
    functionName: "createData"
  },
  {
    name: "削除レコード完全削除",
    functionName: "queryDeleteMarkedRows"
  },
  {
    name: "Bennese大学データインポート",
    functionName: "importUniversityData"
  },
  {
    name: "大学データクリア",
    functionName: "clearUniversityData"
  },
    null, // セパレータ
  {
    name: "全キャッシュ更新",
    functionName: "warmUpAllCache"
  },
  {
    name: "キャッシュ更新トリガー設定",
    functionName: "setupTriggers"
  }
  ];
  spreadsheet.addMenu("管理者メニュー", entries);
}

// 変更監視トリガー (キャッシュ更新)
// 変更監視トリガー (キャッシュ更新)
const CACHE_TARGET_SHEETS = [
  DATA_SHEETS.SETTINGS.SHEET,
  DATA_SHEETS.SEL_GOUHI.SHEET,
  DATA_SHEETS.SEL_KEITAI.SHEET,
  DATA_SHEETS.TEACHERS.SHEET,
  DATA_SHEETS.STUDENTS.SHEET
];

function onEdit(e) {
  if (!e) return;
  // Simple trigger では Advanced Google Services(Sheets API) を利用できないため、
  // インストール型トリガー(=FULL) のときのみキャッシュ更新を実行する。
  if (e.authMode !== ScriptApp.AuthMode.FULL) return;
  checkAndUpdateCache(e.range.getSheet().getName());
}

function onChange(e) {
  if (!e) return;
  // Simple trigger では Advanced Google Services(Sheets API) を利用できないため、
  // インストール型トリガー(=FULL) のときのみキャッシュ更新を実行する。
  if (e.authMode !== ScriptApp.AuthMode.FULL) return;
  const sheet = e.source.getActiveSheet();
  if (sheet) {
    checkAndUpdateCache(sheet.getName());
  }
}

function checkAndUpdateCache(sheetName) {
  if (CACHE_TARGET_SHEETS.includes(sheetName)) {
    // ※ Sheets APIを使用しているため、onEditはインストーラブルトリガーとして設定する必要があります
    try {
      warmUpCache(sheetName);
    } catch (e) {
      logErrorToSheet('checkAndUpdateCache', e.message, e.stack);
      console.error(`キャッシュ更新エラー(${sheetName}): ${e.message}`);
    }
  }
  // 大学データシートが変更された場合、シリアル番号を自動インクリメント
  // 大学データシートが変更された場合、シリアル番号を自動インクリメント
  if (sheetName === DATA_SHEETS.DAIGAKU.SHEET) {
    try {
      incrementUniversitySerial();
    } catch (e) {
      logErrorToSheet('incrementUniversitySerial', e.message, e.stack);
      console.error(`シリアル番号更新エラー: ${e.message}`);
    }
  }
}

// 大学データシリアル番号のインクリメント（設定シートセルB5の整数値をインクリメントする）
function incrementUniversitySerial() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEETS.SETTINGS.SHEET);
  if (!sheet) {
    throw new Error('設定シートが見つかりません');
  }
  const serialCell = sheet.getRange('B5');
  const currentSerial = parseInt(serialCell.getValue()) || 0;
  const newSerial = currentSerial + 1;
  serialCell.setValue(newSerial);
  SpreadsheetApp.flush();
  try {
    warmUpCache(DATA_SHEETS.SETTINGS.SHEET);
  } catch (e) {
    logErrorToSheet('warmUpCache(SETTINGS)', e.message, e.stack);
    console.error('設定シートキャッシュ更新エラー: ' + e);
  }
  return newSerial;
}


//================================================================================
// 3. クライアント連携関数 ※google.script.run から呼び出される関数群
//================================================================================

//--------------------------------------------------------------------------------
// 3-1. 初期データ取得
//--------------------------------------------------------------------------------
// 初期データの取得 (アプリ起動時) - Google Sheets API (batchGet)を使用
// 初期データの取得 (アプリ起動時) - Google Sheets API (batchGet)を使用
function getInitialData() {
  const myMailAddress = Session.getActiveUser().getEmail();
  try {
    // 取得する対象の定義
    const targetSheets = [
      DATA_SHEETS.SEL_KEITAI,
      DATA_SHEETS.SEL_GOUHI,
      DATA_SHEETS.STUDENTS,
      DATA_SHEETS.TEACHERS,
      DATA_SHEETS.SETTINGS
    ];

    // キャッシュを活用してデータを取得
    // 定義オブジェクトの配列を直接渡す
    const dataList = getBatchSheetDataWithCache(targetSheets);

    // データの抽出 (値がない場合は空配列)
    // 1行目はヘッダーなので、2行目以降のデータを取得
    const examTypeOptions = (dataList[0] || []).slice(1).map(row => row[0] || '');
    const resultOptions = (dataList[1] || []).slice(1).map(row => row[0] || '');
    const allStudentsData = dataList[2] || []; // 直接配列として扱う
    const studentData = allStudentsData.find(row => row[STUDENT_DATA.MAIL_ADDR] === myMailAddress);
    const allTeachersData = dataList[3] || [];  // 直接配列として扱う
    const teacherData = allTeachersData.find(row => row[TEACHER_DATA.MAIL_ADDR] === myMailAddress);
    // const settingsRaw = (dataList[4] || []).slice(1); // 不要

    // 設定データの取得 (getBatchSheetDataWithCacheでキャッシュ済み)
    const settings = getSettings();
    let allData = {};
    if (studentData) {
      // 生徒モード
      allData = {
        headerTitle: settings.pageTitle,
        inputEnable: settings.inputEnable,
        myMailAddress: myMailAddress,
        userRole: 'student',
        studentStructure: STUDENT_DATA,
        examStructure: EXAM_DATA,
        examMaxCount: settings.inputMax,
        examTypeOptions: examTypeOptions,
        resultOptions: resultOptions,
        teacherData: [],
        studentsData: [],
        studentData: studentData, // データを送信（1次元配列）
        examData: [], // 空配列：後で取得
        universityCodeSerial: settings.daigakuSerial
      }
    } else if (teacherData) {
      // 教員モード
      allData = {
        headerTitle: settings.pageTitle,
        inputEnable: true,
        myMailAddress: myMailAddress,
        userRole: 'teacher',
        studentStructure: STUDENT_DATA,
        examStructure: EXAM_DATA,
        examMaxCount: settings.inputMax,
        examTypeOptions: examTypeOptions,
        resultOptions: resultOptions,
        teacherData: teacherData, // データを送信（1次元配列）
        studentsData: [], // 空配列：後で取得
        studentData: [],
        examData: [],
        universityCodeSerial: settings.daigakuSerial
      }
    } else {
      // ゲスト/権限なし
      allData = {
        headerTitle: settings.pageTitle,
        inputEnable: false,
        myMailAddress: myMailAddress,
        userRole: '',
        studentStructure: [],
        examStructure: [],
        examMaxCount: 0,
        examTypeOptions: [],
        resultOptions: [],
        teacherData: [],
        studentsData: [],
        studentData: [],
        examData: [],
        universityCodeSerial: settings.daigakuSerial
      }
    }
    return JSON.stringify(allData);
  } catch (e) {
    logErrorToSheet('getInitialData', e.message, e.stack);
    console.error('getInitialDataでエラー: ' + e);
    //    throw new Error('作成者に連絡してください。' + e.message); // デバッグ用
    throw new Error('作成者に連絡してください。');
  }
}

//--------------------------------------------------------------------------------
// 3-2. マスタデータ取得
//--------------------------------------------------------------------------------
// 生徒一覧取得（教員モード用） - API使用
function getStudentsList() {
  try {
    // キャッシュ確認
    const allStudentsData = getSheetDataApiWithCache(DATA_SHEETS.STUDENTS) || [];
    return JSON.stringify(allStudentsData.slice(1)); // ヘッダーを除去して返却
  } catch (e) {
    logErrorToSheet('getStudentsList', e.message, e.stack);
    console.error('getStudentsListでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  }
}

// 大学データの送信 (検索用)
function getUniversityDataList() {
  try {
    const settings = getSettings();
    // 検索用（軽量）データを取得
    const data = getSheetDataApi(DATA_SHEETS.DAIGAKU_SEARCH).slice(1);
    const jsonString = JSON.stringify(data);
    // JSON文字列をBlobに変換してgzip圧縮
    const blob = Utilities.newBlob(jsonString, 'text/plain', 'data.json');
    const compressed = Utilities.gzip(blob);
    const compressedBytes = compressed.getBytes();

    // Base64エンコード
    const base64 = Utilities.base64Encode(compressedBytes);

    return base64;
  } catch (e) {
    logErrorToSheet('getUniversityDataList', e.message, e.stack);
    console.error('getUniversityDataListでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  }
}
/*
function getFilteredUniversityDataList(searchWord, start, limit) {
  const data = getUniversityDataApi() || [];
  const keywords = searchWord.replace(/　/g, ' ').split(' ').filter(kw => kw);
  const filteredData = [];
  for (let row of data) {
    if (keywords.every(kw => (row[1] || "").includes(kw) || String(row[0] || "").includes(kw))) {
      filteredData.push(row);
      continue;
    }
  }
  return JSON.stringify({
    total: filteredData.length,
    data: filteredData.slice(start, start + limit)
  });
}
*/

//--------------------------------------------------------------------------------
// 3-3. 受験データ操作 (取得・保存)
//--------------------------------------------------------------------------------
// 受験データのみの送信 (更新後など) - API使用
function getExamDataList(mailAddr = Session.getActiveUser().getEmail()) {
  try {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = activeSpreadsheet.getId();
    // 受験データ全体を取得 (A1からH列まで)

    const rangeName = `'${DATA_SHEETS.JUKEN_DB.SHEET}'!A1:H`;
    const response = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName);
    const allExamData = response.values || [];

    // フィルタリング
    const examData = allExamData.filter(row => row[EXAM_DATA.MAIL_ADDR] === mailAddr && isTrue(row[EXAM_DATA.DEL_FLAG]) !== true);
    // 注意：DEL_FLAGとSHINGAKUは文字列のままフロントエンドに渡される
    return JSON.stringify(examData);
  } catch (e) {
    logErrorToSheet('getExamDataList', e.message, e.stack);
    console.error('getExamDataListでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  }
}

//
// 受験データの保存 (バリデーション付き・最適化版)
//
function sendExamData(strJuken, mailAddr = Session.getActiveUser().getEmail()) {
  if (!isValidUser(mailAddr)) { // バリデーション
    throw new Error("保存権限がありません。不正なアクセスの可能性があります。");
  }
  const examInputData = JSON.parse(strJuken) || [];
  const activeExamData = examInputData.filter(row => isTrue(row[EXAM_DATA.DEL_FLAG]) !== true); // 削除フラグが立っていないデータのみを抽出してバリデーション
  validateInputData(activeExamData);
  const currentDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const lock = LockService.getDocumentLock();
  let hasLock = false;
  try {
    hasLock = lock.tryLock(10000);
    if (!hasLock) {
      throw new Error('他のユーザーが編集中です。少し待ってから再度お試しください。');
    }
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = activeSpreadsheet.getId();
    // ===== STEP 1: 既存データの取得 =====
    // メールアドレスでフィルタリングして対象レコードを取得
    const indexRange = `'${DATA_SHEETS.JUKEN_DB.SHEET}'!B2:D`;
    const mailResponse = Sheets.Spreadsheets.Values.get(spreadsheetId, indexRange);
    const mailValues = mailResponse.values || [];

    const matchedRowIndices = [];
    mailValues.forEach((row, index) => {
      if (row[0] === mailAddr) {
        matchedRowIndices.push(index + 2); // 実際の行番号
      }
    });
    // 既存データをMapで管理: Map<大学コード, Array<{rowIndex, data}>>
    const existingDataMap = new Map();
    if (matchedRowIndices.length > 0) {
      const sheetName = DATA_SHEETS.JUKEN_DB.SHEET;
      const ranges = matchedRowIndices.map(rowIndex =>
        `'${sheetName}'!A${rowIndex}:H${rowIndex}`
      );
      const batchResponse = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, {
        ranges: ranges
      });
      const valueRanges = batchResponse.valueRanges || [];
      valueRanges.forEach((valueRange, i) => {
        if (valueRange.values && valueRange.values.length > 0) {
          const rowValues = valueRange.values[0];
          const rowIndex = matchedRowIndices[i];
          const daigakuCode = String(rowValues[EXAM_DATA.DAIGAKU_CODE] || '');

          if (daigakuCode) {
            if (!existingDataMap.has(daigakuCode)) {
              existingDataMap.set(daigakuCode, []);
            }
            existingDataMap.get(daigakuCode).push({
              rowIndex: rowIndex,
              data: rowValues
            });
          }
        }
      });
    }
    // ===== STEP 2: 入力データの処理 =====
    const dataToAdd = [];      // 新規追加データ
    const dataToUpdate = [];   // 更新データ (ValueRange objects)
    const processedCodes = new Set(); // 処理済み大学コード
    examInputData.forEach(inputRow => {
      const daigakuCode = String(inputRow[EXAM_DATA.DAIGAKU_CODE] || '').trim();
      if (!daigakuCode) return;
      const isDeleteRequest = isTrue(inputRow[EXAM_DATA.DEL_FLAG]);
      const existingRecords = existingDataMap.get(daigakuCode);
      if (existingRecords && existingRecords.length > 0) { // ---- 既存データがある場合 ----
        if (isDeleteRequest) { // A. 削除リクエスト → すべて論理削除
          existingRecords.forEach(record => {
            const range = `'${DATA_SHEETS.JUKEN_DB.SHEET}'!A${record.rowIndex}:C${record.rowIndex}`;
            dataToUpdate.push({
              range: range,
              values: [[currentDate, mailAddr, true]]
            });
          });
        } else { // B. 更新リクエスト → 最初の1件を更新、残りは論理削除
          const targetRecord = existingRecords[0];
          for (let i = 1; i < existingRecords.length; i++) { // 重複レコード(2件目以降)を論理削除
            const dupRecord = existingRecords[i];
            const range = `'${DATA_SHEETS.JUKEN_DB.SHEET}'!A${dupRecord.rowIndex}:C${dupRecord.rowIndex}`;
            dataToUpdate.push({
              range: range,
              values: [[currentDate, mailAddr, true]]
            });
          }
          // メインレコードの更新チェック
          const existingData = targetRecord.data;
          const isChanged =
            String(existingData[EXAM_DATA.GOUHI] || '') !== String(inputRow[EXAM_DATA.GOUHI] || '') ||
            String(existingData[EXAM_DATA.KEITAI] || '') !== String(inputRow[EXAM_DATA.KEITAI] || '') ||
            isTrue(existingData[EXAM_DATA.SHINGAKU]) !== isTrue(inputRow[EXAM_DATA.SHINGAKU]) ||
            isTrue(existingData[EXAM_DATA.DEL_FLAG]) !== false; // 削除済みからの復帰
          if (isChanged) {
            const updateRow = [
              currentDate,                                      // TIME_STAMP
              mailAddr,                                         // MAIL_ADDR
              false,                                            // DEL_FLAG
              daigakuCode,                                      // DAIGAKU_CODE
              inputRow[EXAM_DATA.DAIGAKU_NAME] || '',          // DAIGAKU_NAME
              inputRow[EXAM_DATA.GOUHI] || '',                 // GOUHI
              inputRow[EXAM_DATA.KEITAI] || '',                // KEITAI
              isTrue(inputRow[EXAM_DATA.SHINGAKU])             // SHINGAKU
            ];
            const range = `'${DATA_SHEETS.JUKEN_DB.SHEET}'!A${targetRecord.rowIndex}:H${targetRecord.rowIndex}`;
            dataToUpdate.push({
              range: range,
              values: [updateRow]
            });
          }
        }
        processedCodes.add(daigakuCode);
      } else {
        if (!isDeleteRequest) { // ---- 既存データがない場合 ----
          const newRow = [ // C. 新規追加
            currentDate,                                      // TIME_STAMP
            mailAddr,                                         // MAIL_ADDR
            false,                                            // DEL_FLAG
            daigakuCode,                                      // DAIGAKU_CODE
            inputRow[EXAM_DATA.DAIGAKU_NAME] || '',          // DAIGAKU_NAME
            inputRow[EXAM_DATA.GOUHI] || '',                 // GOUHI
            inputRow[EXAM_DATA.KEITAI] || '',                // KEITAI
            isTrue(inputRow[EXAM_DATA.SHINGAKU])             // SHINGAKU
          ];
          dataToAdd.push(newRow);
        }
      }
    });
    // ===== STEP 3: 入力に含まれない既存データを論理削除 =====
    existingDataMap.forEach((records, daigakuCode) => {
      if (!processedCodes.has(daigakuCode)) { // 入力データに含まれていない既存レコードはすべて論理削除
        records.forEach(record => { // 既に論理削除されている場合はスキップ
          if (isTrue(record.data[EXAM_DATA.DEL_FLAG]) !== true) {
            const range = `'${DATA_SHEETS.JUKEN_DB.SHEET}'!A${record.rowIndex}:C${record.rowIndex}`;
            dataToUpdate.push({
              range: range,
              values: [[currentDate, mailAddr, true]]
            });
          }
        });
      }
    });
    // ===== STEP 4: データベースへの反映 =====
    // 4-1. 更新の一括実行 (batchUpdate)
    if (dataToUpdate.length > 0) {
      Sheets.Spreadsheets.Values.batchUpdate({
        valueInputOption: 'USER_ENTERED',
        data: dataToUpdate
      }, spreadsheetId);
    }
    // 4-2. 新規データの追加 (append)
    if (dataToAdd.length > 0) {
      const range = `'${DATA_SHEETS.JUKEN_DB.SHEET}'!A1`;
      Sheets.Spreadsheets.Values.append({
        range: range,
        majorDimension: 'ROWS',
        values: dataToAdd
      }, spreadsheetId, range, { valueInputOption: 'USER_ENTERED' });
    }
    // ===== STEP 5: 登録数の更新 =====
    const activeCount = examInputData.filter(row => isTrue(row[EXAM_DATA.DEL_FLAG]) !== true).length;
    updateStudentRegCount(spreadsheetId, mailAddr, activeCount);
    // ===== STEP 6: 最新データを返却 =====
    return getExamDataList(mailAddr);
  } catch (e) {
    logErrorToSheet('sendExamData', e.message, e.stack);
    console.error('sendExamDataでエラー: ' + e);
    if (e.message && e.message.includes('他のユーザーが編集中')) {
      throw e;
    } else if (e.message && e.message.includes('保存権限')) {
      throw e;
    } else {
      throw new Error('データの保存に失敗しました。時間をおいて再度お試しください。');
    }
  } finally {
    if (hasLock) {
      try {
        lock.releaseLock();
      } catch (e) {
        console.warn('ロックの解放に失敗しました:', e);
      }
    }
    SpreadsheetApp.flush();
  }
}
//--------------------------------------------------------------------------------
// 3-4. PDF作成・メール送信
//--------------------------------------------------------------------------------
// PDF作成・メール送信 (調査書交付願)
function sendPdf(mailAddr = Session.getActiveUser().getEmail()) {
  let ssOutput = null;
  let ssOutputID = null;
  try {
    // 受験データ
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const crrDate = new Date();
    const suffix = Utilities.formatDate(crrDate, 'Asia/Tokyo', 'yyyyMMddHHmmss');
    const ssName = '調査書交付願' + suffix; // 一時ファイル名
    // 一時ファイルを作成
    ssOutput = SpreadsheetApp.create(ssName);
    ssOutputID = ssOutput.getId();
    const srcsheet = activeSpreadsheet.getSheetByName(DATA_SHEETS.PDF_TEMPLATE.SHEET);
    const sheetOutput = srcsheet.copyTo(ssOutput);
    const sheetOutputID = sheetOutput.getSheetId();
    // 日時埋め込み
    const dateRange = sheetOutput.getRange("P2");
    dateRange.setNumberFormat('@');
    dateRange.setHorizontalAlignment('right');
    dateRange.setValue(Utilities.formatDate(crrDate, 'Asia/Tokyo', 'yyyy年MM月dd日 HH時mm分'));
    // 生徒データ埋め込み
    const sheetStudents = activeSpreadsheet.getSheetByName(DATA_SHEETS.STUDENTS.SHEET);
    const mailList = sheetStudents.getRange(1, 1, sheetStudents.getLastRow(), 1).getValues().flat();
    const idx = mailList.indexOf(mailAddr);
    const studentData = [];
    const colsStudent = sheetStudents.getLastColumn();
    studentData.push(sheetStudents.getRange(1, 1, 1, colsStudent).getValues().flat());
    studentData.push(sheetStudents.getRange(idx + 1, 1, 1, colsStudent).getValues().flat());
    const pRange = sheetOutput.getRange('G4:P8');
    let setDataList = pRange.getValues();
    for (let i = 0; i < studentData[0].length; i++) {
      setDataList = setDataList.map(record => record.map(value => value.replace('${' + studentData[0][i] + '}', convertIfDate(studentData[1][i], 'yyyy/MM/dd'))));
    }
    pRange.setNumberFormat('@');
    pRange.setValues(setDataList);
    // 受験データ埋め込み
    const examDataSheet = activeSpreadsheet.getSheetByName(DATA_SHEETS.JUKEN_DB.SHEET);
    const allExamData = examDataSheet.getDataRange().getValues();
    const examData = allExamData.filter(row => row[EXAM_DATA.MAIL_ADDR] === mailAddr && isTrue(row[EXAM_DATA.DEL_FLAG]) !== true);
    /* 
      // PDF作成時は全データが必要なので DAIGAKU (COLS:8) を使用する
      // getSheetDataApi を使って最適化可能だが、ここでは元のロジックに合わせてシート取得する
      // ただし、getSheetDataApi経由の方がキャッシュ効く？ -> いや、この処理は即時性が重要かも
      // 元コードは getSheetByName しているので踏襲
    */
    const universitySheet = activeSpreadsheet.getSheetByName(DATA_SHEETS.DAIGAKU.SHEET);
    const allUniversityData = universitySheet.getRange(2, 1, universitySheet.getLastRow() - 1, universitySheet.getLastColumn()).getValues();
    const jRange = sheetOutput.getRange("A11:F40");
    const arrOutput = jRange.getValues();
    for (let i = 0; i < examData.length; i++) {
      let universityCode = examData[i][EXAM_DATA.DAIGAKU_CODE];
      let universityData = allUniversityData.find(arr => String(arr[0]) === String(universityCode)); // unversityCodeを文字列に
      if (universityData) {
        arrOutput[i][0] = universityCode;
        arrOutput[i][1] = universityData[1];
        arrOutput[i][2] = examData[i][EXAM_DATA.KEITAI];
        arrOutput[i][3] = universityData[5];
        arrOutput[i][4] = universityData[6];
        arrOutput[i][5] = universityData[7];
      }
    }
    jRange.setValues(arrOutput);
    SpreadsheetApp.flush();
    // メール送信
    const settings = getSettings()
    const title = `${settings.pageTitle}-${studentData[1][STUDENT_DATA.NAME]}`;
    const message = settings.mailMessage;
    const attachmentfiles = [];
    attachmentfiles.push(createSheetPDF(ssName, ssOutputID, sheetOutputID, true));
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), title, message, {
      name: '開智学園',
      attachments: attachmentfiles
    });
  } catch (e) {
    logErrorToSheet('sendPdf', e.message, e.stack);
    console.error('sendPdfでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  } finally {
    if (ssOutputID) {
      try {
        DriveApp.getFileById(ssOutputID).setTrashed(true); // setTrashed(trashed: Boolean)
      } catch (e) {
        console.warn("一時ファイルの削除に失敗しました: " + ssOutputID);
      }
    }
  }
}

//================================================================================
// 4. バリデーション・設定・ヘルパー関数 (Business Logic & Helpers)
//================================================================================
// 設定データの取得
function getSettings() {
  const values = (getSheetDataApiWithCache(DATA_SHEETS.SETTINGS) || []).slice(1);
  return {
    pageTitle: values[0] ? values[0][1] : "",
    inputMax: values[1] ? values[1][1] : 0,
    inputEnable: values[2] ? isTrue(values[2][1]) : false, // APIの戻り値は文字列
    daigakuSerial: values[3] ? values[3][1] : "",
    mailTitle: values[4] ? values[4][1] : "",
    mailMessage: values[5] ? values[5][1] : ""
  };
}
// 権限チェック用ヘルパー (本人が操作しているか、または教員が操作しているかを確認)
function isValidUser(targetMailAddr) {
  const currentUser = Session.getActiveUser().getEmail();
  // 1. 自分自身のデータを保存しようとしている場合はOK
  if (currentUser === targetMailAddr) {
    return true;
  }
  // 2. 他人のデータを保存しようとしている場合、実行者が「教員」かチェック
  const allTeachersData = (getSheetDataApiWithCache(DATA_SHEETS.TEACHERS) || []).slice(1);
  const isTeacher = allTeachersData.some(row => row[TEACHER_DATA.MAIL_ADDR] === currentUser);
  if (isTeacher) {
    return true; // 教員なら他人のデータも保存OK
  }
  return false;
}
// 入力データ整合性チェック用ヘルパー
function validateInputData(data) {

  const maxCount = getSettings().inputMax;
  const allowedResults = (getSheetDataApiWithCache(DATA_SHEETS.SEL_GOUHI) || []).slice(1).map(row => row[0] || "");
  const allowedTypes = (getSheetDataApiWithCache(DATA_SHEETS.SEL_KEITAI) || []).slice(1).map(row => row[0] || "");
  if (!Array.isArray(data)) {
    throw new Error("データ形式が不正です。");
  }
  if (data.length > maxCount) {
    throw new Error(`登録可能件数(${maxCount}件)を超えています。`);
  }
  data.forEach((row, index) => {
    const gouhi = row[EXAM_DATA.GOUHI];
    if (gouhi && !allowedResults.includes(gouhi)) {
      throw new Error(`No.${index + 1} の合否選択肢「${gouhi}」は不正です。`);
    }
    const keitai = row[EXAM_DATA.KEITAI];
    if (keitai && !allowedTypes.includes(keitai)) {
      throw new Error(`No.${index + 1} の受験形態「${keitai}」は不正です。`);
    }
  });
}
// 生徒データの登録数を更新するヘルパー
function updateStudentRegCount(spreadsheetId, mailAddr, count) {
  try {
    const rangeName = `'${DATA_SHEETS.STUDENTS.SHEET}'!A:A`;
    const response = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName);
    const mailCol = (response.values || []).flat();
    const rowIndex = mailCol.indexOf(mailAddr);
    if (rowIndex > 0) {
      const regCountCol = STUDENT_DATA.REG_COUNT + 1; // 1-indexed (F列 = 6)
      const updateRange = `'${DATA_SHEETS.STUDENTS.SHEET}'!F${rowIndex + 1}`;
      Sheets.Spreadsheets.Values.update(
        { values: [[count]] },
        spreadsheetId,
        updateRange,
        { valueInputOption: 'USER_ENTERED' }
      );
    }
  } catch (e) {
    console.warn('登録数の更新に失敗しました:', e);
  }
}
// 日付フォーマット変換ヘルパー
function convertIfDate(data, format) {
  if (Object.prototype.toString.call(data) === '[object Date]') {
    data = Utilities.formatDate(data, 'Asia/Tokyo', format);
  }
  return data
}
// スプレッドシートのPDF化ヘルパー
function createSheetPDF(ssName, ssID, sheetID, portrait) {
  const token = ScriptApp.getOAuthToken();
  const pdfoptions = `&format=pdf&portrait=${portrait}&size=A4&fitw=true&gridlines=false`;
  const url = `https://docs.google.com/spreadsheets/d/${ssID}/export?gid=${sheetID}${pdfoptions}`;
  return UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`
    }
  }).getBlob().setName(`${ssName}.pdf`);
}

//================================================================================
// 5. データアクセス (Data Access)
//================================================================================



// 複数シートデータの取得（キャッシュ付き・バッチ取得）
function getBatchSheetDataWithCache(sheetDefs) {
  const cache = CacheService.getScriptCache();
  const results = new Array(sheetDefs.length);
  const fetchIndices = [];
  const fetchRanges = [];

  // 1. キャッシュの確認
  sheetDefs.forEach((def, index) => {
    const cacheKey = def.SHEET;
    const cached = cache.get(cacheKey);
    if (cached) {
      const data = JSON.parse(cached);
      // キャッシュデータが要求カラム数より多い場合、切り詰めて返す
      if (def.COLS > 0 && data.length > 0 && data[0].length > def.COLS) {
        results[index] = data.map(row => row.slice(0, def.COLS));
      } else {
        results[index] = data;
      }
    } else {
      fetchIndices.push(index);
      // APIリクエスト情報を生成
      const req = makeSheetApiRequest(def);
      fetchRanges.push(req.range);
    }
  });

  // 2. キャッシュにないデータをバッチ取得
  if (fetchIndices.length > 0) {
    try {
      const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const spreadsheetId = activeSpreadsheet.getId();
      const response = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, { ranges: fetchRanges });
      const valueRanges = response.valueRanges;

      valueRanges.forEach((vr, i) => {
        const originalIndex = fetchIndices[i];
        const data = vr.values || [];
        results[originalIndex] = data;

        // 取得したデータをキャッシュに保存
        // 取得時に指定カラム数で絞っている(A:F等)ので、返却データはCOLS以下になっているはず
        // そのままキャッシュしてOK
        const def = sheetDefs[originalIndex];
        const cacheKey = def.SHEET;
        cache.put(cacheKey, JSON.stringify(data), 21600); // 6時間
      });
    } catch (e) {
      logErrorToSheet('getBatchSheetDataWithCache', e.message, e.stack);
      console.error('getBatchSheetDataWithCacheでエラー: ' + e);
      throw new Error("データの一括取得に失敗しました。");
    }
  }

  return results;
}


// シートデータの取得キャッシュ付き（getBatchSheetDataWithCacheに統合を検討すべきだが一旦維持）
// ※ 引数を sheetName から sheetDef に変更
// シートデータの取得キャッシュ付き
// ※ 引数は sheetDef オブジェクト必須
function getSheetDataApiWithCache(sheetDef) {
  const cache = CacheService.getScriptCache();
  const cacheKey = sheetDef.SHEET; // 単純化のためシート名をキーにする
  const cached = cache.get(cacheKey);
  if (cached) {
    const data = JSON.parse(cached);
    // キャッシュデータが要求カラム数より多い場合、切り詰めて返す (Superset Cache -> Subset Request)
    if (sheetDef.COLS > 0 && data.length > 0 && data[0].length > sheetDef.COLS) {
      return data.map(row => row.slice(0, sheetDef.COLS));
    }
    return data;
  }
  const data = getSheetDataApi(sheetDef);
  // データがあればキャッシュ（空配列はキャッシュしない方が安全？）
  if (data && data.length > 0) {
    cache.put(cacheKey, JSON.stringify(data), 21600);
  }
  return data;
}

// シートデータの取得（共通化）
function getSheetDataApi(sheetDef) {
  try {
    const req = makeSheetApiRequest(sheetDef);
    // Values.get は範囲指定がないと全範囲、A:Bならその列を全行取得
    const response = Sheets.Spreadsheets.Values.get(req.id, req.range);
    return response.values || [];
  } catch (e) {
    logErrorToSheet('getSheetDataApi', e.message, e.stack);
    console.error(e);
    // エラー時は空配列を返すか、エラーを投げるか
    // 呼び出し元の期待値に合わせる（空配列）
    return [];
  }
}

// キャッシュの強制更新
function warmUpCache(sheetName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = sheetName;
  /* sheetName string -> sheetDef object conversion for cache updates */
  let sheetDef = Object.values(DATA_SHEETS).find(d => d.SHEET === sheetName);
  if (!sheetDef) {
    sheetDef = { SHEET: sheetName, COLS: 0 };
  }
  const data = getSheetDataApi(sheetDef);
  cache.put(cacheKey, JSON.stringify(data), 21600); // 6時間
  return data;
}

// 全キャッシュの強制更新
function warmUpAllCache() {
  const targetSheets = [

    DATA_SHEETS.SETTINGS.SHEET,
    DATA_SHEETS.TEACHERS.SHEET,
    DATA_SHEETS.STUDENTS.SHEET,
    DATA_SHEETS.SEL_GOUHI.SHEET,
    DATA_SHEETS.SEL_KEITAI.SHEET
  ];
  targetSheets.forEach(sheetName => {
    warmUpCache(sheetName);
  });
}

// キャッシュ更新トリガーの設定
function setupTriggers() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 既存のトリガーを削除（重複防止）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    const handlerFunction = trigger.getHandlerFunction();
    if (handlerFunction === 'onEdit' || handlerFunction === 'onChange' || handlerFunction === 'warmUpAllCache' || handlerFunction === 'deleteMarkedRows') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // onEdit トリガー（編集時）
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  // onChange トリガー（変更時）
  ScriptApp.newTrigger('onChange')
    .forSpreadsheet(spreadsheet)
    .onChange()
    .create();

  // 1時間ごとのキャッシュ更新トリガー
  ScriptApp.newTrigger('warmUpAllCache')
    .timeBased()
    .everyHours(3)
    .create();

  // deleteMarkedRows
  ScriptApp.newTrigger('deleteMarkedRows')
    .timeBased()
    .everyHours(3)
    .create();

  Browser.msgBox("キャッシュ更新トリガーの設定が完了しました。\\n\\n設定されたトリガー:\\n- 編集時（onEdit）\\n- 変更時（onChange）\\n- 3時間ごと（warmUpAllCache）\\n- 3時間ごと（deleteMarkedRows）");
}

//================================================================================
// 6. 管理者・メンテナンス用関数 ※ メニューバーから実行される関数群
//================================================================================
// 校内DB用データの生成
function createData() { // API不使用のためboolean対策なし
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetDB = activeSpreadsheet.getSheetByName(DATA_SHEETS.KOUNAI_DB.SHEET);
  const universitySheet = activeSpreadsheet.getSheetByName(DATA_SHEETS.DAIGAKU.SHEET);
  const sheetJuken = activeSpreadsheet.getSheetByName(DATA_SHEETS.JUKEN_DB.SHEET);
  const sheetStudents = activeSpreadsheet.getSheetByName(DATA_SHEETS.STUDENTS.SHEET);
  const allUniversityData = universitySheet.getRange(2, 1, universitySheet.getLastRow() - 1, universitySheet.getLastColumn()).getValues();
  const mapAllDaigaku = new Map;
  allUniversityData.forEach(drow => mapAllDaigaku.set(String(drow[0]), drow[1]));
  const allExamData = sheetJuken.getRange(2, 1, sheetJuken.getLastRow() - 1, sheetJuken.getLastColumn()).getValues();
  const allStudentsData = sheetStudents.getRange(2, 1, sheetStudents.getLastRow() - 1, sheetStudents.getLastColumn()).getValues();
  const sheetData = []

  // 1. 受験データを生徒メールアドレスをキーにしてMap化 (高速化)
  const examMap = new Map();
  allExamData.forEach(jrow => {
    // 削除フラグが立っているデータは除外
    if (isTrue(jrow[EXAM_DATA.DEL_FLAG]) === true) return;

    const mail = jrow[EXAM_DATA.MAIL_ADDR];
    if (!examMap.has(mail)) {
      examMap.set(mail, []);
    }
    examMap.get(mail).push(jrow);
  });

  allStudentsData.forEach(strow => {
    const outst = [strow[STUDENT_DATA.GRADE], strow[STUDENT_DATA.CLASS], strow[STUDENT_DATA.NUMBER], strow[STUDENT_DATA.NAME]]
    // Mapから取得 (O(1))
    const examData = examMap.get(strow[STUDENT_DATA.MAIL_ADDR]) || [];

    examData.forEach(jrow => {
      const outj = [
        jrow[EXAM_DATA.DAIGAKU_CODE],
        mapAllDaigaku.get(String(jrow[EXAM_DATA.DAIGAKU_CODE])),
        jrow[EXAM_DATA.KEITAI],
        jrow[EXAM_DATA.GOUHI],
        jrow[EXAM_DATA.SHINGAKU]
      ];
      sheetData.push(outst.concat(outj));
    });
  });

  sheetDB.clear();
  sheetDB.appendRow(['学年', 'クラス', '出席番号', '氏名', '大学コード', '大学学部学科名', '試験形態', '合否', '進学'])
  if (sheetData.length > 0) {
    sheetDB.getRange(2, 1, sheetData.length, sheetData[0].length).setValues(sheetData);
  }
}
// 削除マークのついた行を完全削除
function queryDeleteMarkedRows() {
  const dialogResult = Browser.msgBox("削除マークの付いた行を完全に削除します。よろしいですか？", Browser.Buttons.YES_NO);
  if (dialogResult !== 'yes') return;
  deleteMarkedRows();
}
function deleteMarkedRows() { // 削除フラグ付きを一括削除（再書き込み）
  const lock = LockService.getDocumentLock(); // スプレッドシート単位でロック
  try {
    lock.waitLock(120000); // 120秒待機
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const examDataSheet = activeSpreadsheet.getSheetByName(DATA_SHEETS.JUKEN_DB.SHEET);

    // データ範囲全体を取得
    const values = examDataSheet.getDataRange().getValues();

    // 削除対象でない行だけをフィルタリング (ヘッダーは残す)
    // isTrue関数を使用して boolean / string 両対応
    const newValues = values.filter((row, i) => {
      if (i === 0) return true; // ヘッダー行保持
      return isTrue(row[EXAM_DATA.DEL_FLAG]) !== true; // 削除フラグが TRUE でないものを残す
    });

    // 行数が減っている場合のみ書き換え実行
    if (newValues.length < values.length) {
      examDataSheet.clearContents();
      examDataSheet.getRange(1, 1, newValues.length, newValues[0].length).setValues(newValues);
    }
  } catch (e) {
    logErrorToSheet('deleteMarkedRows', e.message, e.stack);
    console.error('deleteMarkedRowsでエラーまたはロックタイムアウト: ' + e);
    throw new Error('削除処理に失敗しました。時間をおいて再度お試しください。' + e.message);
  } finally {
    lock.releaseLock();
    SpreadsheetApp.flush();
  }
}
// 大学データのクリア
function clearUniversityData() {
  const dialogResult = Browser.msgBox("大学データを消去します。\\n\\nよろしいですか？", Browser.Buttons.YES_NO);
  if (dialogResult === 'no') return;

  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const universitySheet = activeSpreadsheet.getSheetByName(DATA_SHEETS.DAIGAKU.SHEET);
  universitySheet.clear();
  const outputHeader = ['大学コード', '大学名', 'Web締切', '窓口締切', '郵送締切', '入試日', '発表日', '手続き締切'];
  universitySheet.appendRow(outputHeader);
  Browser.msgBox("消去が完了しました。");
}
// Benesseデータのインポート
function importUniversityData() {
  const dialogResult = Browser.msgBox("現在のシートから大学データを追記ます。\\n既存のデータは上書きされます。\\n\\nよろしいですか？", Browser.Buttons.YES_NO);
  if (dialogResult === 'no') return;
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const benesseSheet = activeSpreadsheet.getActiveSheet();
  const benesseDataRows = benesseSheet.getDataRange().getValues();
  if (benesseDataRows?.[0]?.[BENESSE_DATA.CODE_DAIGAKU] !== '大学ｺｰﾄﾞ' || benesseDataRows?.[0]?.[BENESSE_DATA.CODE_GAKUBU] !== '学部ｺｰﾄﾞ') {
    Browser.msgBox("現在のシートはBeneeseのデータではないようです。", Browser.Buttons.OK);
    return;
  }
  const universitySheet = activeSpreadsheet.getSheetByName(DATA_SHEETS.DAIGAKU.SHEET);
  const universityDataRows = universitySheet.getDataRange().getValues();
  universityDataRows.shift();
  const universityMap = new Map();
  universityDataRows.forEach(row => {
    universityMap.set(String(row[0]), row);
  })
  const isCalendar = benesseDataRows[0][BENESSE_DATA.DAY_NYUSHI] === '入試日' ? true : false
  benesseDataRows.shift();
  // 基準年度の計算
  const baseYear = calculateBaseYear(new Date());

  const appendData = [];
  benesseDataRows.forEach(row => {
    const rowData = [];
    let universityCode = row[BENESSE_DATA.CODE_DAIGAKU] + row[BENESSE_DATA.CODE_GAKUBU] + row[BENESSE_DATA.CODE_GAKKA] + row[BENESSE_DATA.CODE_NITTEI] + row[BENESSE_DATA.CODE_HOUSHIKI]
    let universityName = (row[BENESSE_DATA.NAME_DAIGAKU] || "").trim();
    let facultyName = (row[BENESSE_DATA.NAME_GAKUBU] || "").trim();
    if (facultyName.length > 0) universityName += `・${facultyName}`;
    let gakka = (row[BENESSE_DATA.NAME_GAKKA] || "").trim();
    if (gakka.length > 0) universityName += `・${gakka}`;
    let schedule = (row[BENESSE_DATA.NAME_NITTEI] || "").trim();
    let method = (row[BENESSE_DATA.NAME_HOUSHIKI] || "").trim();
    if (schedule.length > 0 && method.length > 0) universityName += `[${schedule}]${method}`;
    else if (schedule.length > 0) universityName += `[${schedule}]`;
    else if (method.length > 0) universityName += `・${method}`;
    rowData.push(String(universityCode), universityName);

    if (isCalendar) {
      let scheduleArray = row.splice(BENESSE_DATA.DAY_SHIMEKIRI_WEB, 7);
      scheduleArray.splice(3, 1); // 不要な列(Window?)を除去
      scheduleArray.forEach(str => {
        rowData.push(formatBenesseDate(str, baseYear));
      });
    } else {
      rowData.push(...Array(6))
    }
    const existRow = universityMap.get(rowData[0])
    if (existRow) {
      existRow.length = 0;
      existRow.push(...rowData);
    } else {
      universityDataRows.push(rowData);
    }
  })
  universityDataRows.sort((a, b) => String(a[0]).localeCompare(String(b[0])));
  universitySheet.clear();
  const outputHeader = ['大学コード', '大学名', 'Web締切', '窓口締切', '郵送締切', '入試日', '発表日', '手続き締切'];
  universitySheet.appendRow(outputHeader);
  universitySheet.getRange(2, 1, universityDataRows.length, outputHeader.length).setValues(universityDataRows);
  // シリアル番号の更新
  incrementUniversitySerial();
  Browser.msgBox("インポートが完了しました。");
}

// エラーログの記録
function logErrorToSheet(type, message, detail) {
  const lock = LockService.getScriptLock();
  let hasLock = false;
  try {
    // ログ記録も同時書き込みを避けるためロックを取得
    lock.waitLock(5000);
    hasLock = true;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DATA_SHEETS.ERROR_LOG.SHEET);
    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = ss.insertSheet(DATA_SHEETS.ERROR_LOG.SHEET);
      sheet.appendRow(['日時', 'ユーザー', '種類', 'メッセージ', '詳細（スタックトレース等）']);
      sheet.setFrozenRows(1);
    }
    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date();
    sheet.appendRow([
      timestamp,
      userEmail,
      type,
      message,
      detail
    ]);
  } catch (e) {
    // ログ記録自体が失敗した場合はStackdriverにのみ記録
    console.error('Failed to log error to sheet: ' + e.toString());
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}
