/**
 * 受験校調査システム (Preferred school survey system)
 * Version 2.0.0
 * * Copyright (c) 2025 Shigeru Suzuki
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
const appUrl = ScriptApp.getService().getUrl();
const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const myMailAddress = Session.getActiveUser().getEmail();

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
// 2. トリガー・エントリポイント (Triggers)
//================================================================================
// Webアプリとしてのエントリポイント (doGet)
function doGet(e) {
  const settings = getSettings();
  const htmlTemplate = HtmlService.createTemplateFromFile('index');
  const html = htmlTemplate.evaluate();
  html.setTitle(settings.pageTitle);
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
  }
  ];
  spreadsheet.addMenu("校内DB用", entries);
}


//================================================================================
// 3. クライアント連携関数 ※google.script.run から呼び出される関数群
//================================================================================
// 初期データの取得 (アプリ起動時) - Google Sheets API (batchGet)を使用
function getInitialData() {
  try {
    const spreadsheetId = activeSpreadsheet.getId();
    // 取得する範囲のリスト
    const ranges = [
      `'${SHEET_NAMES.SEL_KEITAI}'!A2:A`,  // 受験形態選択肢
      `'${SHEET_NAMES.SEL_GOUHI}'!A2:A`,   // 合否選択肢
      `'${SHEET_NAMES.STUDENTS}'!A1:E`,    // 生徒データ 1行目は列数確認のためクライアントに渡す
      `'${SHEET_NAMES.TEACHERS}'!A1:B`,    // 職員データ 1行目は列数確認のためクライアントに渡す
      `'${SHEET_NAMES.SETTINGS}'!A2:B`     // 設定
    ];

    // batchGetで一括取得
    const response = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, { ranges: ranges });
    const valueRanges = response.valueRanges;

    // データの抽出 (値がない場合は空配列)
    const examTypeOptions = (valueRanges[0].values || []).map(row => {
      let examType = row[0]
      if (!examType) examType = '';
      return examType;
    });
    const resultOptions = (valueRanges[1].values || []).map(row => {
      let result = row[0];
      if (!result) result = '';
      return result;
    });
    const allStudentsData = valueRanges[2].values || []; // 直接配列として扱う
    const studentData = allStudentsData.find(row => row[STUDENT_DATA.MAIL_ADDR] == myMailAddress);
    const arrAllTeachers = valueRanges[3].values || [];  // 直接配列として扱う
    const arrTeacher = arrAllTeachers.find(row => row[TEACHER_DATA.MAIL_ADDR] == myMailAddress);
    const settingsRaw = valueRanges[4].values || [];

    // 設定データのオブジェクト化 (getSettings相当の処理)
    const settings = {
      pageTitle: settingsRaw[0] ? settingsRaw[0][1] : "",
      inputMax: settingsRaw[1] ? settingsRaw[1][1] : 0,
      inputEnable: settingsRaw[2] ? settingsRaw[2][1] : false, // デフォルトfalse
      mailTitle: settingsRaw[3] ? settingsRaw[3][1] : "",
      mailMessage: settingsRaw[4] ? settingsRaw[4][1] : ""
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
        studentData: [allStudentsData[0], studentData], // ヘッダーとデータをセットで送信
        examData: [], // 空配列：後で取得
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
        teacherData: [arrAllTeachers[0], arrTeacher], // ヘッダーとデータをセットで送信
        studentsData: [], // 空配列：後で取得
        studentData: [],
        examData: [],
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
      }
    }
    return JSON.stringify(allData);
  } catch (e) {
    console.error('getInitialDataでエラー: ' + e);
    //    throw new Error('作成者に連絡してください。' + e.message); // デバッグ用
    throw new Error('作成者に連絡してください。');
  }
}

// 生徒一覧取得（教員モード用） - API使用
function getStudentsList() {
  try {
    const spreadsheetId = activeSpreadsheet.getId();
    // ヘッダー(A1)を含むデータ範囲を取得
    const rangeName = `'${SHEET_NAMES.STUDENTS}'!A1:E`;
    const response = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName);
    const allStudentsData = response.values || [];
    return JSON.stringify(allStudentsData);
  } catch (e) {
    console.error('getStudentsListでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  }
}

// 大学データの送信 (検索用)
function getUniversityDataList() {
  try {
    // キャッシュ確認
    return JSON.stringify(getUniversityDataApi());
  } catch (e) {
    console.error('getUniversityDataListでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  }
}

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


// 受験データのみの送信 (更新後など) - API使用
function getExamDataList(mailAddr = myMailAddress) {
  try {
    const spreadsheetId = activeSpreadsheet.getId();
    // 受験データ全体を取得 (A1からZ列まで)
    const rangeName = `'${SHEET_NAMES.JUKEN_DB}'!A1:H`;
    const response = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName);
    const allExamData = response.values || [];

    // フィルタリング
    const examData = allExamData.filter(row => row[EXAM_DATA.MAIL_ADDR] == mailAddr && String(row[EXAM_DATA.DEL_FLAG]).toUpperCase() !== "TRUE");

    // クライアント側の padRowsWithHeader で1行目がヘッダーとして扱われ削除されるため、
    // ここでヘッダー行を先頭に追加する
    if (allExamData.length > 0) {
      examData.unshift(allExamData[0]);
    }

    return JSON.stringify(examData);
  } catch (e) {
    console.error('getExamDataListでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  }
}

// 受験データの保存 (バリデーション付き・最適化版)
function saveExamDataList(strJuken, mailAddr = myMailAddress) {
  // バリデーション
  if (!isValidUser(mailAddr)) {
    throw new Error("保存権限がありません。不正なアクセスの可能性があります。");
  }
  const examInputData = JSON.parse(strJuken) || [];
  // 設定値（最大件数や選択肢）を取得してチェック
  const settings = getSettings();
  const allowedResults = (getSheetDataApiWithCache(SHEET_NAMES.SEL_GOUHI) || []).map(row => row[0] || "");
  const allowedTypes = (getSheetDataApiWithCache(SHEET_NAMES.SEL_KEITAI) || []).map(row => row[0] || "");
  // 削除フラグが立っていないデータのみを抽出してバリデーション
  const activeExamData = examInputData.filter(row => row[EXAM_DATA.DEL_FLAG] != true);
  validateInputData(activeExamData, settings.inputMax, allowedResults, allowedTypes);
  const currentDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  // 1. 既存データの取得とマップ化 (Sheets APIによる最適化された取得)
  const existingDataMap = new Map();
  const spreadsheetId = activeSpreadsheet.getId();

  // A. メールアドレス列(Col B)を一括取得して検索
  // B2:B の範囲を取得
  const indexRange = `'${SHEET_NAMES.JUKEN_DB}'!B2:D`; // row[0]:mailAddr, row[1]:delFlag, row[2]:daigakuCode;
  const mailResponse = Sheets.Spreadsheets.Values.get(spreadsheetId, indexRange);
  const mailValues = mailResponse.values || [];
  const matchedRowIndices = [];
  mailValues.forEach((row, index) => {
    // row[0] がメールアドレス
    if (row[0] === mailAddr) {
      // index は 0-based で B2 スタートなので、実際の行番号は index + 2
      matchedRowIndices.push(index + 2);
    }
  });

  // B. 対象行のデータ取得 (BatchGet)
  if (matchedRowIndices.length > 0) {
    // 取得する範囲のリストを作成 (例: '受験校DB!A10:H10')
    // 最終列は H (index 7) なので、H列まで取得すれば十分
    // ※ EXAM_DATA.SHINGAKU が 7 なので Column H corresponds to index 7 (A=0, H=7 ? No. A=1 in A1 notation, index 0 in array)
    // EXAM_DATA values are 0-based indices for array access.
    // Column H is the 8th column.
    const sheetName = SHEET_NAMES.JUKEN_DB;
    const ranges = matchedRowIndices.map(rowIndex => `'${sheetName}'!A${rowIndex}:H${rowIndex}`);

    const batchResponse = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, { ranges: ranges });
    const valueRanges = batchResponse.valueRanges || [];

    valueRanges.forEach((valueRange, i) => {
      if (valueRange.values && valueRange.values.length > 0) {
        const rowValues = valueRange.values[0];
        const rowIndex = matchedRowIndices[i];

        // 削除フラグがTrueでないもののみマップに追加
        if (String(rowValues[EXAM_DATA.DEL_FLAG]).toLowerCase() !== "true") {
          existingDataMap.set(String(rowValues[EXAM_DATA.DAIGAKU_CODE]), {
            rowIndex: rowIndex,
            data: rowValues
          });
        }
      }
    });
  }

  const dataToAdd = [];
  const dataToUpdate = []; // 更新リクエストのリスト (ValueRange objects)

  // 2. 入力データの処理 (更新・新規・論理削除)
  examInputData.forEach(inputRow => {
    const daigakuCode = String(inputRow[EXAM_DATA.DAIGAKU_CODE]);
    if (!daigakuCode) return;

    const existingRecord = existingDataMap.get(daigakuCode);

    // A. クライアント側で削除指示 (DEL_FLAG = true)
    if (String(inputRow[EXAM_DATA.DEL_FLAG]).toLowerCase() === "true") {
      if (existingRecord) {
        // 既存にあれば削除フラグを立てる (Sheets API batchUpdate用)
        // Range: シート名!C<行番号>
        const range = `'${SHEET_NAMES.JUKEN_DB}'!A${existingRecord.rowIndex}:C${existingRecord.rowIndex}`;
        dataToUpdate.push({
          range: range,
          values: [[currentDate, mailAddr, "TRUE"]]
        });
        existingDataMap.delete(daigakuCode); // 処理済みとしてMapから削除
      }
      return; // 新規追加もしない
    }
    // B. 既存データの更新 または 変更なし
    if (existingRecord) {
      // 比較対象のカラム: 合否, 受験形態, 進学, (大学コードはキーなので一致), (大学名はマスタ依存だが一応確認?)
      // ここでは編集可能な主要項目を比較
      const isChanged =
        String(existingRecord.data[EXAM_DATA.GOUHI]) !== String(inputRow[EXAM_DATA.GOUHI]) ||
        String(existingRecord.data[EXAM_DATA.KEITAI]) !== String(inputRow[EXAM_DATA.KEITAI]) ||
        String(existingRecord.data[EXAM_DATA.SHINGAKU]).toLowerCase() !== String(inputRow[EXAM_DATA.SHINGAKU]).toLowerCase(); // boolean
      if (isChanged) {
        const updateRow = [...existingRecord.data]; // 既存データをコピー
        updateRow[EXAM_DATA.TIME_STAMP] = currentDate; // 更新日時
        updateRow[EXAM_DATA.DAIGAKU_NAME] = inputRow[EXAM_DATA.DAIGAKU_NAME];
        updateRow[EXAM_DATA.GOUHI] = inputRow[EXAM_DATA.GOUHI];
        updateRow[EXAM_DATA.KEITAI] = inputRow[EXAM_DATA.KEITAI];
        updateRow[EXAM_DATA.SHINGAKU] = String(inputRow[EXAM_DATA.SHINGAKU]).toUpperCase();
        const range = `'${SHEET_NAMES.JUKEN_DB}'!A${existingRecord.rowIndex}:H${existingRecord.rowIndex}`;
        dataToUpdate.push({
          range: range,
          values: [updateRow]
        });
      }
      // 変更なしの場合は何もしないで処理済みにする
      existingDataMap.delete(daigakuCode); // 処理済みとしてMapから削除
    } else {
      // C. 新規追加
      const newRow = [];
      newRow[EXAM_DATA.TIME_STAMP] = currentDate;
      newRow[EXAM_DATA.MAIL_ADDR] = mailAddr;
      newRow[EXAM_DATA.DEL_FLAG] = "FALSE";
      newRow[EXAM_DATA.DAIGAKU_CODE] = inputRow[EXAM_DATA.DAIGAKU_CODE];
      newRow[EXAM_DATA.DAIGAKU_NAME] = inputRow[EXAM_DATA.DAIGAKU_NAME];
      newRow[EXAM_DATA.GOUHI] = inputRow[EXAM_DATA.GOUHI];
      newRow[EXAM_DATA.KEITAI] = inputRow[EXAM_DATA.KEITAI];
      newRow[EXAM_DATA.SHINGAKU] = String(inputRow[EXAM_DATA.SHINGAKU]).toUpperCase();
      dataToAdd.push(newRow);
    }
  });
  // 3. 入力データに存在しなかった既存データ (削除されたとみなす)
  // existingDataMap に残っているデータは、クライアントからのリストに含まれていなかったもの
  if (existingDataMap.size > 0) {
    existingDataMap.forEach((record) => {
      // Range: シート名!A<行番号>:C<行番号>
      const range = `'${SHEET_NAMES.JUKEN_DB}'!A${record.rowIndex}:C${record.rowIndex}`;
      dataToUpdate.push({
        range: range,
        values: [[currentDate, mailAddr, true]]
      });
    });
  }

  // 4-1. 更新の一括実行 (batchUpdate)
  const lock = LockService.getDocumentLock(); // スプレッドシート単位でロック
  let hasLock = false; // フラグの初期化
  try {
    // 待機時間を10秒に延長
    hasLock = lock.tryLock(10000);
    if (!hasLock) {
      throw new Error('他のユーザーが編集中です。少し待ってから再度お試しください。');
    }
    if (dataToUpdate.length > 0) {
      Sheets.Spreadsheets.Values.batchUpdate({
        valueInputOption: 'USER_ENTERED',
        data: dataToUpdate
      }, spreadsheetId);
    }

    // 4-2. 新規データの追加 (append)
    if (dataToAdd.length > 0) {
      // シート名のみを指定してappend (最終行に追加される)
      const range = `'${SHEET_NAMES.JUKEN_DB}'!A1`;
      Sheets.Spreadsheets.Values.append({
        range: range,
        majorDimension: 'ROWS',
        values: dataToAdd
      }, spreadsheetId, range, { valueInputOption: 'USER_ENTERED' });
    }

    // 変更点：保存完了後に最新のデータリストを直接返す
    return getExamDataList(mailAddr);

  } catch (e) {
    console.error('saveExamDataListでエラー: ' + e);
    // エラーメッセージを詳細化
    if (e.message && e.message.includes('他のユーザーが編集中')) {
      throw e; // タイムアウトメッセージをそのまま返す
    } else if (e.message && e.message.includes('保存権限')) {
      throw e; // 権限エラーメッセージをそのまま返す
    } else {
      throw new Error('データの保存に失敗しました。時間をおいて再度お試しください。');
    }
  } finally {
    if (hasLock) {
      try {
        lock.releaseLock();
      } catch (e) {
        console.warn('ロックの解放に失敗しました:', e);
        // ロック解放エラーは無視（自動的に解放される）
      }
    }
    SpreadsheetApp.flush();
  }
}
/*
// setValuesによる書き込み
function saveExamDataList(strJuken, mailAddr = myMailAddress) {
  const lock = LockService.getDocumentLock(); // スプレッドシート単位でロック
  let hasLock = false;
  try {
    // より短い待機時間で試行（5秒）
    hasLock = lock.tryLock(5000);
    if (!hasLock) {
      throw new Error('他のユーザーが編集中です。少し待ってから再度お試しください。');
    }

    if (!isValidUser(mailAddr)) {
      throw new Error("保存権限がありません。不正なアクセスの可能性があります。");
    }
    const examInputData = JSON.parse(strJuken) || [];

    // 設定値（最大件数や選択肢）を取得してチェック
    const settings = getSettings();
    const allowedResults = (JSON.parse(getSheetData(SHEET_NAMES.SEL_GOUHI)) || []).map(row => row[0]);
    const allowedTypes = (JSON.parse(getSheetData(SHEET_NAMES.SEL_KEITAI)) || []).map(row => row[0]);

    validateInputData(examInputData, settings.inputMax, allowedResults, allowedTypes);

    const examDataSheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.JUKEN_DB);

    // 1. 既存データの論理削除 (DEL_FLAG = TRUE)
    if (examDataSheet.getLastRow() > 1) {
      const searchRange = examDataSheet.getRange(2, EXAM_DATA.MAIL_ADDR + 1, examDataSheet.getLastRow() - 1, 1);
      const finder = searchRange.createTextFinder(mailAddr).matchEntireCell(true);
      const ranges = finder.findAll();

      if (ranges.length > 0) {
        const delFlagRanges = ranges.map(range => range.offset(0, 1).getA1Notation());
        examDataSheet.getRangeList(delFlagRanges).setValue(true);
      }
    }

    // 2. 新規データの追加
    const currentDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const dataToAdd = [];

    examInputData.forEach(inputRow => {
      const daigakuCode = String(inputRow[EXAM_DATA.DAIGAKU_CODE]);
      // クライアント側で削除済み、または無効なデータは追加しない
      if (!daigakuCode || daigakuCode.length == 0 || inputRow[EXAM_DATA.DEL_FLAG] == true) return;

      inputRow[EXAM_DATA.TIME_STAMP] = currentDate;
      inputRow[EXAM_DATA.MAIL_ADDR] = mailAddr;
      inputRow[EXAM_DATA.DEL_FLAG] = false; // 新規データはFalse

      dataToAdd.push(inputRow);
    });

    if (dataToAdd.length > 0) {
      examDataSheet.getRange(examDataSheet.getLastRow() + 1, 1, dataToAdd.length, dataToAdd[0].length).setValues(dataToAdd);
    }

  } catch (e) {
    console.error('saveExamDataListでエラー: ' + e);
    // エラーメッセージを詳細化
    if (e.message && e.message.includes('他のユーザーが編集中')) {
      throw e; // タイムアウトメッセージをそのまま返す
    } else if (e.message && e.message.includes('保存権限')) {
      throw e; // 権限エラーメッセージをそのまま返す
    } else {
      throw new Error('データの保存に失敗しました。時間をおいて再度お試しください。');
    }
  } finally {
    if (hasLock) {
      try {
        lock.releaseLock();
      } catch (e) {
        console.warn('ロックの解放に失敗しました:', e);
        // ロック解放エラーは無視（自動的に解放される）
      }
    }
    SpreadsheetApp.flush();
  }
}
*/
/*
// 全データ書き換えバージョン
function saveExamDataList(strJuken, mailAddr = myMailAddress) {
  const lock = LockService.getDocumentLock(); // スプレッドシート単位でロック
  try {
    lock.waitLock(30000); // 30秒待機
    if (!isValidUser(mailAddr)) {
      throw new Error("保存権限がありません。不正なアクセスの可能性があります。");
    }
    const examInputData = JSON.parse(strJuken);
    
    // 設定値（最大件数や選択肢）を取得してチェック
    const settings = getSettings();
    const allowedResults = JSON.parse(getSheetData(SHEET_NAMES.SEL_GOUHI)).map(row => row[0]);
    const allowedTypes = JSON.parse(getSheetData(SHEET_NAMES.SEL_KEITAI)).map(row => row[0]);
    
    validateInputData(examInputData, settings.inputMax, allowedResults, allowedTypes);
    
    // データ保存処理
    const examDataSheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.JUKEN_DB);
    const allData = examDataSheet.getDataRange().getValues();
    const header = allData.shift(); // ヘッダーを分離（未使用）
    const savedJukenData = allData.filter(row => row[EXAM_DATA.MAIL_ADDR] == mailAddr);
    const otherStudentsData = allData.filter(row => row[EXAM_DATA.MAIL_ADDR] != mailAddr);
    const finalData = [];
    const arrCompareItems = [EXAM_DATA.GOUHI, EXAM_DATA.KEITAI, EXAM_DATA.SHINGAKU];
    const currentDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    
    examInputData.forEach(inputRow => {
      const daigakuCode = String(inputRow[EXAM_DATA.DAIGAKU_CODE]);
      if (!daigakuCode || daigakuCode.length == 0 || inputRow[EXAM_DATA.DEL_FLAG] == true) return;
      let strDate = currentDate;
      const savedRow = savedJukenData.find(row => String(inputRow[EXAM_DATA.DAIGAKU_CODE]) == String(row[EXAM_DATA.DAIGAKU_CODE]))
      if (savedRow) {
        if (arrCompareItems.every(item => String(inputRow[item]) == String(savedRow[item]))) {
          strDate = savedRow[EXAM_DATA.TIME_STAMP];
        }
      }
      inputRow[EXAM_DATA.TIME_STAMP] = strDate;
      inputRow[EXAM_DATA.MAIL_ADDR] = mailAddr;
      finalData.push(inputRow);
    });
    
    const dataToWrite = otherStudentsData.concat(finalData);
    if (examDataSheet.getLastRow() > 1) {
      examDataSheet.getRange(2, 1, examDataSheet.getLastRow() - 1, examDataSheet.getLastColumn()).clearContent();
    }
    if (dataToWrite.length > 0) {
      examDataSheet.getRange(2, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
    }
  } catch (e) {
    console.error('saveExamDataListでエラーまたはロックタイムアウト: ' + e);
    throw new Error('データの保存に失敗しました。時間をおいて再度お試しください。');
  } finally {
    lock.releaseLock();
    SpreadsheetApp.flush();
  }
}
*/
// PDF作成・メール送信 (調査書交付願)
function sendPdf(mailAddr = myMailAddress) {
  let ssOutput = null;
  let ssOutputID = null;
  try {
    // 受験データ
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
    const examData = allExamData.filter(row => row[EXAM_DATA.MAIL_ADDR] == mailAddr && row[EXAM_DATA.DEL_FLAG] != true);
    const universitySheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.DAIGAKU);
    const allUniversityData = universitySheet.getRange(2, 1, universitySheet.getLastRow() - 1, universitySheet.getLastColumn()).getValues();
    const jRange = sheetOutput.getRange("A11:F40");
    const arrOutput = jRange.getValues();

    for (let i = 0; i < examData.length; i++) {
      let universityCode = examData[i][EXAM_DATA.DAIGAKU_CODE];
      let universityData = allUniversityData.find(arr => arr[0] == universityCode);
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
    attachmentfiles.push(createSheetPDF(ssName, ssOutputID, sheetOutputID, "true"));

    GmailApp.sendEmail(myMailAddress, title, message, {
      name: '開智学園',
      attachments: attachmentfiles
    });
  } catch (e) {
    console.error('sendPdfでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  } finally {
    if (ssOutputID) {
      try {
        DriveApp.getFileById(ssOutputID).setTrashed(true);
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
  const values = getSheetDataApiWithCache(SHEET_NAMES.SETTINGS) || []
  return {
    pageTitle: values[0][1] || "",
    inputMax: values[1][1] || 0,
    inputEnable: values[2][1],
    mailTitle: values[3][1] || "",
    mailMessage: values[4][1] || ""
  };
}

// 権限チェック用ヘルパー (本人が操作しているか、または教員が操作しているかを確認)
function isValidUser(targetMailAddr) {
  const currentUser = Session.getActiveUser().getEmail();
  // 1. 自分自身のデータを保存しようとしている場合はOK
  if (currentUser == targetMailAddr) {
    return true;
  }
  // 2. 他人のデータを保存しようとしている場合、実行者が「教員」かチェック
  const arrAllTeachers = getSheetDataApiWithCache(SHEET_NAMES.TEACHERS) || [];
  const isTeacher = arrAllTeachers.some(row => row[0] == currentUser);
  if (isTeacher) {
    return true; // 教員なら他人のデータも保存OK
  }
  return false;
}

// 入力データ整合性チェック用ヘルパー
function validateInputData(data, maxCount, allowedResults, allowedTypes) {
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
  if (Object.prototype.toString.call(data) == '[object Date]') {
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
    const spreadsheetId = activeSpreadsheet.getId();
    const sheet = activeSpreadsheet.getSheetByName(sheetName);
    const lastRow = sheet.getLastRow();

    // 実際のデータ行数のみを取得（ヘッダー行をスキップ）
    if (lastRow <= 1) {
      return []; // データがない場合
    }

    // 範囲指定：A2からB列の最終行まで
    const rangeName = `'${sheetName}'!A2:B${lastRow}`;

    // Sheets API を直接叩いて値を取得
    const response = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName);
    const data = response.values || [];

    return data;
  } catch (e) {
    console.error(e);
    throw new Error("API取得エラー: " + e.message);
  }
}

// シートデータの取得
function getSheetData(sheetName) {
  try {
    const objSheet = activeSpreadsheet.getSheetByName(sheetName);
    const allData = objSheet.getRange(2, 1, objSheet.getLastRow() - 1, objSheet.getLastColumn()).getValues();
    const jsonString = JSON.stringify(allData);
    return jsonString;
  } catch (e) {
    console.error('getSheetDataでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  }
}

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

function getSheetDataApi(sheetName) {
  try {
    const spreadsheetId = activeSpreadsheet.getId();
    const sheet = activeSpreadsheet.getSheetByName(sheetName);
    const lastRow = sheet.getLastRow();
    // 実際のデータ行数のみを取得（ヘッダー行をスキップ）
    if (lastRow <= 1) {
      return []; // データがない場合
    }
    const lastColumn = sheet.getLastColumn();
    const a1Notation = sheet.getRange(2, 1, lastRow - 1, lastColumn).getA1Notation()
    const rangeName = `'${sheetName}'!${a1Notation}`;
    // Sheets API を直接叩いて値を取得
    const response = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName);
    return response.values || [];
  } catch (e) {
    console.error(e);
    throw new Error("API取得エラー: " + e.message);
  }
}

//================================================================================
// 6. 管理者・メンテナンス用関数 ※ メニューバーから実行される関数群
//================================================================================
// 校内DB用データの生成
function createData() {
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
  allStudentsData.forEach(strow => {
    const outst = [strow[STUDENT_DATA.GRADE], strow[STUDENT_DATA.CLASS], strow[STUDENT_DATA.NUMBER], strow[STUDENT_DATA.NAME]]
    // 変数宣言 const を追加しました
    const examData = allExamData.filter(jrow => jrow[EXAM_DATA.MAIL_ADDR] == strow[STUDENT_DATA.MAIL_ADDR] && jrow[EXAM_DATA.DEL_FLAG] == false)
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
  if (dialogResult != 'yes') return;
  deleteMarkedRows();
}

function deleteMarkedRows() {
  const lock = LockService.getDocumentLock(); // スプレッドシート単位でロック
  try {
    lock.waitLock(120000); // 120秒待機
    const examDataSheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.JUKEN_DB);
    const arrDeleteFlag = examDataSheet.getRange(1, EXAM_DATA.DEL_FLAG + 1, examDataSheet.getLastRow(), 1).getValues().flat();
    for (let i = arrDeleteFlag.length - 1; i > 0; i--) { // [0] はタイトル行 行番号の大きい方から消す
      if (arrDeleteFlag[i] == true) {
        examDataSheet.deleteRow(i + 1);
      }
    }
  } catch (e) {
    console.error('saveExamDataListでエラーまたはロックタイムアウト: ' + e);
    throw new Error('データの保存に失敗しました。時間をおいて再度お試しください。');
  } finally {
    lock.releaseLock();
    SpreadsheetApp.flush();
  }
}


// 大学データのクリア
function clearUniversityData() {
  const dialogResult = Browser.msgBox("大学データを消去します。\\n\\nよろしいですか？", Browser.Buttons.YES_NO);
  if (dialogResult == 'no') return;

  const universitySheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.DAIGAKU);
  universitySheet.clear();
  const outputHeader = ['大学コード', '大学名', 'Web締切', '窓口締切', '郵送締切', '入試日', '発表日', '手続き締切'];
  universitySheet.appendRow(outputHeader);
  Browser.msgBox("消去が完了しました。");
}

// Benesseデータのインポート
function importUniversityData() {
  const dialogResult = Browser.msgBox("現在のシートから大学データを追記ます。\\n既存のデータは上書きされます。\\n\\nよろしいですか？", Browser.Buttons.YES_NO);
  if (dialogResult == 'no') return;
  const benesseSheet = activeSpreadsheet.getActiveSheet();
  const benesseDataRows = benesseSheet.getDataRange().getValues();
  if (benesseDataRows?.[0]?.[BENESSE_DATA.CODE_DAIGAKU] != '大学ｺｰﾄﾞ' || benesseDataRows?.[0]?.[BENESSE_DATA.CODE_GAKUBU] != '学部ｺｰﾄﾞ') {
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
  const isCalendar = benesseDataRows[0][BENESSE_DATA.DAY_NYUSHI] == '入試日' ? true : false
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
        if (str == "0000" || str.length != 4) {
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
  Browser.msgBox("インポートが完了しました。");
}
