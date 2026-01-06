/**
 * 受験校調査システム (Preferred school survey system)
 * Version 2.3.1
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
const SHEET_NAMES = {
  STUDENTS: '学籍データ',
  TEACHERS: '職員データ',
  JUKEN_DB: '受験校DB',
  KEITAI: '試験形態',
  DAIGAKU: '大学データ',
  SEL_GOUHI: '合否選択肢',
  SEL_KEITAI: '受験形態選択肢',
  SETTINGS: '設定',
  PDF_TEMPLATE: '調査書交付願',
  KOUNAI_DB: '校内DB用'
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
  NAME: 4
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
const CACHE_TARGET_SHEETS = [
  SHEET_NAMES.SETTINGS,
  SHEET_NAMES.SEL_GOUHI,
  SHEET_NAMES.SEL_KEITAI,
  SHEET_NAMES.TEACHERS,
  SHEET_NAMES.STUDENTS
];

function onEdit(e) {
  if (!e) return;
  checkAndUpdateCache(e.range.getSheet().getName());
}

function onChange(e) {
  if (!e) return;
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
    } catch (err) {
      console.error(`キャッシュ更新エラー(${sheetName}): ${err.message}`);
    }
  }
  // 大学データシートが変更された場合、シリアル番号を自動インクリメント
  if (sheetName === SHEET_NAMES.DAIGAKU) {
    try {
      incrementUniversitySerial();
    } catch (err) {
      console.error(`シリアル番号更新エラー: ${err.message}`);
    }
  }
}

// 大学データシリアル番号のインクリメント（設定シートセルB5の整数値をインクリメントする）
function incrementUniversitySerial() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    throw new Error('設定シートが見つかりません');
  }
  const serialCell = sheet.getRange('B5');
  const currentSerial = parseInt(serialCell.getValue()) || 0;
  serialCell.setValue(currentSerial + 1);
  return newSerial;
}


//================================================================================
// 3. クライアント連携関数 ※google.script.run から呼び出される関数群
//================================================================================

//--------------------------------------------------------------------------------
// 3-1. 初期データ取得
//--------------------------------------------------------------------------------
// 初期データの取得 (アプリ起動時) - Google Sheets API (batchGet)を使用
function getInitialData() {
  const myMailAddress = Session.getActiveUser().getEmail();
  try {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = activeSpreadsheet.getId();
    // 取得する範囲の定義
    const requests = [
      { sheetName: SHEET_NAMES.SEL_KEITAI, range: `'${SHEET_NAMES.SEL_KEITAI}'!A1:A` },
      { sheetName: SHEET_NAMES.SEL_GOUHI, range: `'${SHEET_NAMES.SEL_GOUHI}'!A1:A` },
      { sheetName: SHEET_NAMES.STUDENTS, range: `'${SHEET_NAMES.STUDENTS}'!A1:E` },
      { sheetName: SHEET_NAMES.TEACHERS, range: `'${SHEET_NAMES.TEACHERS}'!A1:B` },
      { sheetName: SHEET_NAMES.SETTINGS, range: `'${SHEET_NAMES.SETTINGS}'!A1:B` }
    ];

    // キャッシュを活用してデータを取得
    const dataList = getBatchSheetDataWithCache(requests);

    // データの抽出 (値がない場合は空配列)
    // 1行目はヘッダーなので、2行目以降のデータを取得
    const examTypeOptions = (dataList[0] || []).slice(1).map(row => row[0] || '');
    const resultOptions = (dataList[1] || []).slice(1).map(row => row[0] || '');
    const allStudentsData = dataList[2] || []; // 直接配列として扱う
    const studentData = allStudentsData.find(row => row[STUDENT_DATA.MAIL_ADDR] === myMailAddress);
    const arrAllTeachers = dataList[3] || [];  // 直接配列として扱う
    const arrTeacher = arrAllTeachers.find(row => row[TEACHER_DATA.MAIL_ADDR] === myMailAddress);
    const settingsRaw = (dataList[4] || []).slice(1);

    // 設定データのオブジェクト化 (getSettings相当の処理)
    const settings = {
      pageTitle: settingsRaw[0] ? settingsRaw[0][1] : "",
      inputMax: settingsRaw[1] ? settingsRaw[1][1] : 0,
      inputEnable: settingsRaw[2] ? isTrue(settingsRaw[2][1]) : false, // APIの戻り値は文字列
      daigakuSerial: settingsRaw[3] ? settingsRaw[3][1] : "",
      mailTitle: settingsRaw[4] ? settingsRaw[4][1] : "",
      mailMessage: settingsRaw[5] ? settingsRaw[5][1] : ""
    };
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
    } else if (arrTeacher) {
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
        teacherData: arrTeacher, // データを送信（1次元配列）
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
    const allStudentsData = getSheetDataApiWithCache(SHEET_NAMES.STUDENTS) || [];
    return JSON.stringify(allStudentsData.slice(1)); // ヘッダーを除去して返却
  } catch (e) {
    console.error('getStudentsListでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  }
}

// 大学データの送信 (検索用)
function getUniversityDataList() {
  try {
    const settings = getSettings();
    const data = getUniversityDataApi();
    const jsonString = JSON.stringify(data);
    // JSON文字列をBlobに変換してgzip圧縮
    const blob = Utilities.newBlob(jsonString, 'text/plain', 'data.json');
    const compressed = Utilities.gzip(blob);
    const compressedBytes = compressed.getBytes();

    // Base64エンコード
    const base64 = Utilities.base64Encode(compressedBytes);

    return base64;
  } catch (e) {
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
    const rangeName = `'${SHEET_NAMES.JUKEN_DB}'!A1:H`;
    const response = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName);
    const allExamData = response.values || [];

    // フィルタリング
    const examData = allExamData.filter(row => row[EXAM_DATA.MAIL_ADDR] === mailAddr && isTrue(row[EXAM_DATA.DEL_FLAG]) !== true);
    // 注意：DEL_FLAGとSHINGAKUは文字列のままフロントエンドに渡される
    return JSON.stringify(examData);
  } catch (e) {
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
    const indexRange = `'${SHEET_NAMES.JUKEN_DB}'!B2:D`;
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
      const sheetName = SHEET_NAMES.JUKEN_DB;
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
            const range = `'${SHEET_NAMES.JUKEN_DB}'!A${record.rowIndex}:C${record.rowIndex}`;
            dataToUpdate.push({
              range: range,
              values: [[currentDate, mailAddr, true]]
            });
          });
        } else { // B. 更新リクエスト → 最初の1件を更新、残りは論理削除
          const targetRecord = existingRecords[0];
          for (let i = 1; i < existingRecords.length; i++) { // 重複レコード(2件目以降)を論理削除
            const dupRecord = existingRecords[i];
            const range = `'${SHEET_NAMES.JUKEN_DB}'!A${dupRecord.rowIndex}:C${dupRecord.rowIndex}`;
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
            const range = `'${SHEET_NAMES.JUKEN_DB}'!A${targetRecord.rowIndex}:H${targetRecord.rowIndex}`;
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
            const range = `'${SHEET_NAMES.JUKEN_DB}'!A${record.rowIndex}:C${record.rowIndex}`;
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
      const range = `'${SHEET_NAMES.JUKEN_DB}'!A1`;
      Sheets.Spreadsheets.Values.append({
        range: range,
        majorDimension: 'ROWS',
        values: dataToAdd
      }, spreadsheetId, range, { valueInputOption: 'USER_ENTERED' });
    }
    // ===== STEP 5: 最新データを返却 =====
    return getExamDataList(mailAddr);
  } catch (e) {
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
    const srcsheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.PDF_TEMPLATE);
    const sheetOutput = srcsheet.copyTo(ssOutput);
    const sheetOutputID = sheetOutput.getSheetId();
    // 日時埋め込み
    const dateRange = sheetOutput.getRange("P2");
    dateRange.setNumberFormat('@');
    dateRange.setHorizontalAlignment('right');
    dateRange.setValue(Utilities.formatDate(crrDate, 'Asia/Tokyo', 'yyyy年MM月dd日 HH時mm分'));
    // 生徒データ埋め込み
    const sheetStudents = activeSpreadsheet.getSheetByName(SHEET_NAMES.STUDENTS);
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
    const examDataSheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.JUKEN_DB);
    const allExamData = examDataSheet.getDataRange().getValues();
    const examData = allExamData.filter(row => row[EXAM_DATA.MAIL_ADDR] === mailAddr && isTrue(row[EXAM_DATA.DEL_FLAG]) !== true);
    const universitySheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.DAIGAKU);
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
  const values = (getSheetDataApiWithCache(SHEET_NAMES.SETTINGS) || []).slice(1);
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
  const arrAllTeachers = (getSheetDataApiWithCache(SHEET_NAMES.TEACHERS) || []).slice(1);
  const isTeacher = arrAllTeachers.some(row => row[0] === currentUser);
  if (isTeacher) {
    return true; // 教員なら他人のデータも保存OK
  }
  return false;
}
// 入力データ整合性チェック用ヘルパー
function validateInputData(data) {
  const maxCount = getSettings().inputMax;
  const allowedResults = (getSheetDataApiWithCache(SHEET_NAMES.SEL_GOUHI) || []).slice(1).map(row => row[0] || "");
  const allowedTypes = (getSheetDataApiWithCache(SHEET_NAMES.SEL_KEITAI) || []).slice(1).map(row => row[0] || "");
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
// シートデータの取得API使用（大学データ用）
function getUniversityDataApi() {
  const sheetName = SHEET_NAMES.DAIGAKU
  try {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = activeSpreadsheet.getId();
    const sheet = activeSpreadsheet.getSheetByName(sheetName);
    const lastRow = sheet.getLastRow();
    // 実際のデータ行数のみを取得（ヘッダー行をスキップ）
    if (lastRow <= 1) {
      return []; // データがない場合
    }
    const rangeName = `'${sheetName}'!A2:B${lastRow}`; // 範囲指定：A2からB列の最終行まで
    const response = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName); // Sheets API を直接叩いて値を取得
    return response.values || [];
  } catch (e) {
    console.error(e);
    throw new Error("API取得エラー: " + e.message);
  }
}

// 複数シートデータの取得（キャッシュ付き・バッチ取得）
function getBatchSheetDataWithCache(requests) {
  const cache = CacheService.getScriptCache();
  const results = new Array(requests.length);
  const fetchIndices = [];
  const fetchRanges = [];

  // 1. キャッシュの確認
  requests.forEach((req, index) => {
    const cacheKey = req.sheetName;
    const cached = cache.get(cacheKey);
    if (cached) {
      results[index] = JSON.parse(cached);
    } else {
      fetchIndices.push(index);
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
        const req = requests[originalIndex];
        const cacheKey = req.sheetName;
        cache.put(cacheKey, JSON.stringify(data), 21600); // 6時間
      });
    } catch (e) {
      console.error('getBatchSheetDataWithCacheでエラー: ' + e);
      throw new Error("データの一括取得に失敗しました。");
    }
  }

  return results;
}

// シートデータの取得キャッシュ付き
function getSheetDataApiWithCache(sheetName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = sheetName;
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
  const data = getSheetDataApi(sheetName);
  cache.put(cacheKey, JSON.stringify(data), 21600);
  return data;
}
// シートデータの取得
function getSheetDataApi(sheetName) {
  try {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = activeSpreadsheet.getId();
    const sheet = activeSpreadsheet.getSheetByName(sheetName);
    const lastRow = sheet.getLastRow();
    // 実際のデータ行数のみを取得
    if (lastRow < 1) {
      return []; // データがない場合
    }
    const lastColumn = sheet.getLastColumn();
    const a1Notation = sheet.getRange(1, 1, lastRow, lastColumn).getA1Notation();
    const rangeName = `'${sheetName}'!${a1Notation}`;
    // Sheets API を直接叩いて値を取得
    const response = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName);
    return response.values || [];
  } catch (e) {
    console.error(e);
    throw new Error("API取得エラー: " + e.message);
  }
}

// キャッシュの強制更新
function warmUpCache(sheetName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = sheetName;
  const data = getSheetDataApi(sheetName);
  cache.put(cacheKey, JSON.stringify(data), 21600); // 6時間
  return data;
}

// 全キャッシュの強制更新
function warmUpAllCache() {
  const targetSheets = [
    SHEET_NAMES.SETTINGS,
    SHEET_NAMES.TEACHERS,
    SHEET_NAMES.STUDENTS,
    SHEET_NAMES.SEL_GOUHI,
    SHEET_NAMES.SEL_KEITAI
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
  const sheetDB = activeSpreadsheet.getSheetByName(SHEET_NAMES.KOUNAI_DB);
  const universitySheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.DAIGAKU);
  const sheetJuken = activeSpreadsheet.getSheetByName(SHEET_NAMES.JUKEN_DB);
  const sheetStudents = activeSpreadsheet.getSheetByName(SHEET_NAMES.STUDENTS);
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
    const examDataSheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.JUKEN_DB);

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
  const universitySheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.DAIGAKU);
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
  const universitySheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.DAIGAKU);
  const universityDataRows = universitySheet.getDataRange().getValues();
  universityDataRows.shift();
  const universityMap = new Map();
  universityDataRows.forEach(row => {
    universityMap.set(String(row[0]), row);
  })
  const isCalendar = benesseDataRows[0][BENESSE_DATA.DAY_NYUSHI] === '入試日' ? true : false
  benesseDataRows.shift();
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
    let examDate = [];
    if (isCalendar) {
      const now = new Date();
      let oldYear
      let newYear
      let currentYear = now.getFullYear();
      let currentMonth = now.getMonth();
      if (currentMonth > 3) {
        oldYear = currentYear;
        newYear = currentYear + 1;
      } else {
        oldYear = currentYear - 1;
        newYear = currentYear;
      }
      let scheduleArray = row.splice(BENESSE_DATA.DAY_SHIMEKIRI_WEB, 7);
      scheduleArray.splice(3, 1);
      scheduleArray.forEach(str => {
        if (str === "0000" || str.length !== 4) {
          examDate = '';
        } else {
          const dataMonth = Number(str.substring(0, 2))
          const dataDay = Number(str.substring(2))
          if (dataMonth > 3) {
            examDate = `${oldYear}/${dataMonth}/${dataDay}`
          } else {
            examDate = `${newYear}/${dataMonth}/${dataDay}`
          }
        }
        rowData.push(examDate);
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
