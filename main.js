//================================================================================
//
// 受験校調査システム main.gs
//
//================================================================================

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
console.log(settings)
  const htmlTemplate = HtmlService.createTemplateFromFile('index');
  const html = htmlTemplate.evaluate();
  html.setTitle(settings.pageTitle);
  //html.addMetaTag('viewport', 'width=device-width, user-scalable=no, initial-scale=1');
  return html;
}

// スプレッドシートを開いたときのメニュー追加 (onOpen)
function onOpen() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const entries = [{
      name: "校内DB用データ生成",
      functionName: "createData"
    },
    {
      name: "削除レコード完全削除(後方互換性)",
      functionName: "deleteMarkedRows"
    },
    {
      name: "Bennese大学データインポート",
      functionName: "importUniversityData"
    },
    {
      name: "大学データクリア",
      functionName: "clearUniversityData"
    },
    {
      name: "キャッシュクリア",
      functionName: "clearAllCache"
    }
  ];
  spreadsheet.addMenu("校内DB用", entries);
}

// 編集時トリガー (デバウンス対応)
function onEdit(e) {
  if (!e || !e.source) return;
  const sheetName = e.source.getActiveSheet().getName();
  // 3秒のデバウンスを実行
  runDebouncedCacheClear(sheetName, e.source, "編集");
}
// 変更時トリガー (デバウンス対応)
function onChange(e) {
  if (!e || !e.source) return;
  // 行削除・挿入などを検知
  if (e.changeType == 'REMOVE_ROW' || e.changeType == 'INSERT_ROW' || e.changeType == 'OTHER') {
    const sheetName = e.source.getActiveSheet().getName();
    // 3秒のデバウンスを実行
    runDebouncedCacheClear(sheetName, e.source, "構成変更");
  }
}
// デバウンス制御付きキャッシュクリア実行関数
// @param {string} sheetName - シート名
// @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
// @param {string} changeTypeStr - 変更の種類（表示用）
function runDebouncedCacheClear(sheetName, spreadsheet, changeTypeStr) {
  const DEBOUNCE_TIME_MS = 10000; // 10秒待機
  const PROP_KEY = 'LAST_TRIGGER_' + sheetName; // シートごとにキーを分ける
  const now = new Date().getTime();
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, now.toString());
  Utilities.sleep(DEBOUNCE_TIME_MS);
  const lastTriggerTime = PropertiesService.getScriptProperties().getProperty(PROP_KEY);
  if (lastTriggerTime == now.toString()) {
    processCacheClear(sheetName, spreadsheet, changeTypeStr);
  } else {
    console.log(`Debounced: ${sheetName} の処理をスキップしました`);
  }
}
// キャッシュクリア共通処理 (実処理)
function processCacheClear(sheetName, spreadsheet, changeTypeStr) {
  const cacheKeys = [SHEET_NAMES.SEL_KEITAI, SHEET_NAMES.SEL_GOUHI, SHEET_NAMES.STUDENTS, SHEET_NAMES.TEACHERS, SHEET_NAMES.DAIGAKU, SHEET_NAMES.SETTINGS];
  if (cacheKeys.includes(sheetName)) {
    clearCache(sheetName);
    // トースト通知もここで出す（10秒後に1回だけ出る）
    spreadsheet.toast(`シート【${sheetName}】の${changeTypeStr}を反映しました。`, "システム通知");
  }
}

//================================================================================
// 3. クライアント連携関数 ※google.script.run から呼び出される関数群
//================================================================================
// 初期データの取得 (アプリ起動時)
function getInitialData() {
  try {
    const examTypeOptions = JSON.parse(getSheetData(SHEET_NAMES.SEL_KEITAI)).map(row => row[0]);
    const resultOptions = JSON.parse(getSheetData(SHEET_NAMES.SEL_GOUHI)).map(row => row[0]);
    const allStudentsData = JSON.parse(getSheetData(SHEET_NAMES.STUDENTS))
    const studentData = allStudentsData.find(row => row[STUDENT_DATA.MAIL_ADDR] == myMailAddress);
    const arrAllTeachers = JSON.parse(getSheetData(SHEET_NAMES.TEACHERS));
    const arrTeacher = arrAllTeachers.find(row => row[TEACHER_DATA.MAIL_ADDR] == myMailAddress);
    const settings = getSettings()
    let allData = {};
    if (studentData) {
      const examDataSheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.JUKEN_DB);
      const allExamData = examDataSheet.getDataRange().getValues();
      const examData = allExamData.filter(row => row[EXAM_DATA.MAIL_ADDR] == myMailAddress && row[EXAM_DATA.DEL_FLAG] != true);
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
        studentData: studentData,
        examData: examData,
      }
    } else if (arrTeacher) {
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
        teacherData: arrTeacher,
        studentsData: allStudentsData,
        studentData: [],
        examData: [],
      }
    } else {
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
    console.log('getInitialDataでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  }
}

// 大学データの送信 (検索用)
function getUniversityDataList() {
  try {
    // キャッシュ確認
    return getSheetData(SHEET_NAMES.DAIGAKU)
  } catch (e) {
    console.log('getUniversityDataListでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  }
}

// 受験データのみの送信 (更新後など)
function getExamDataList(mailAddr = myMailAddress) {
  try {
    const examDataSheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.JUKEN_DB);
    const allExamData = examDataSheet.getDataRange().getValues();
    const examData = allExamData.filter(row => row[EXAM_DATA.MAIL_ADDR] == mailAddr && row[EXAM_DATA.DEL_FLAG] == false);
    return JSON.stringify(examData);
  } catch (e) {
    console.log('getExamDataListでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  }
}

// 受験データの保存 (バリデーション付き)
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
    console.log('saveExamDataListでエラーまたはロックタイムアウト: ' + e);
    throw new Error('データの保存に失敗しました。時間をおいて再度お試しください。');
  } finally {
    lock.releaseLock();
    SpreadsheetApp.flush();
  }
}

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
    console.log('sendPdfでエラー: ' + e);
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
  const values = JSON.parse(getSheetData(SHEET_NAMES.SETTINGS))
  return {
    pageTitle: values[0][1],
    inputMax: values[1][1],
    inputEnable: values[2][1],
    mailTitle: values[3][1],
    mailMessage: values[4][1]
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
  const arrAllTeachers = JSON.parse(getSheetData(SHEET_NAMES.TEACHERS));
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
// 5. データアクセス・キャッシュ制御 (Data Access & Cache)
//================================================================================
// シートデータの取得 (キャッシュ対応)
function getSheetData(sheetName) {
  try {
    const cacheKey = sheetName;
    const cachedData = getCache(cacheKey);
    if (cachedData) { // キャッシュ確認
      return cachedData;
    }
    const objSheet = activeSpreadsheet.getSheetByName(sheetName);
    const allData = objSheet.getRange(2, 1, objSheet.getLastRow() - 1, objSheet.getLastColumn()).getValues();
    const jsonString = JSON.stringify(allData);
    // キャッシュ保存 (6時間)
    putCache(cacheKey, jsonString, 21600);
    return jsonString;
  } catch (e) {
    console.log('getSheetDataでエラー: ' + e);
    throw new Error('作成者に連絡してください。');
  }
}
// キャッシュ保存 (チャンク分割対応)
function putCache(key, value, expirationInSeconds) {
  const cache = CacheService.getScriptCache();
  const chunkSize = 100000; // 100KB safe margin
  const jsonStr = value;
  if (jsonStr.length <= chunkSize) {
    cache.put(key, jsonStr, expirationInSeconds);
    return;
  }
  let chunkCount = 0;
  for (let i = 0; i < jsonStr.length; i += chunkSize) {
    const chunk = jsonStr.substring(i, i + chunkSize);
    cache.put(key + "_" + chunkCount, chunk, expirationInSeconds);
    chunkCount++;
  }
  cache.put(key, "CHUNKED_" + chunkCount, expirationInSeconds);
}
// キャッシュ取得 (チャンク結合対応)
function getCache(key) {
  const cache = CacheService.getScriptCache();
  const cachedValue = cache.get(key);
  if (!cachedValue) return null;
  if (cachedValue.startsWith("CHUNKED_")) {
    const chunkCount = parseInt(cachedValue.split("_")[1], 10);
    let fullString = "";
    for (let i = 0; i < chunkCount; i++) {
      const chunk = cache.get(key + "_" + i);
      if (!chunk) return null;
      fullString += chunk;
    }
    return fullString;
  }
  return cachedValue;
}
// 特定キーのキャッシュクリア
function clearCache(key) {
  const cache = CacheService.getScriptCache();
  const cachedValue = cache.get(key);
  if (!cachedValue) return null;
  cache.remove(key);
  if (cachedValue.startsWith("CHUNKED_")) {
    const chunkCount = parseInt(cachedValue.split("_")[1], 10);
    for (let i = 0; i < chunkCount; i++) {
      cache.remove(`${key}_${i}`);
    }
  }
}

// 全キャッシュクリア (メニュー実行用)
function clearAllCache() {
  const cacheKeys = [SHEET_NAMES.SEL_KEITAI, SHEET_NAMES.SEL_GOUHI, SHEET_NAMES.STUDENTS, SHEET_NAMES.TEACHERS, SHEET_NAMES.DAIGAKU]
  cacheKeys.forEach(key => {
    clearCache(key);
  })
  Browser.msgBox("すべてのキャッシュをクリアしました。");
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
      const outj = [jrow[EXAM_DATA.DAIGAKU_CODE], mapAllDaigaku.get(String(jrow[EXAM_DATA.DAIGAKU_CODE]))];
      sheetData.push(outst.concat(outj));
    });
  });
  sheetDB.clear();
  sheetDB.appendRow(['学年', 'クラス', '出席番号', '氏名', '大学コード', '大学学部学科名'])
  sheetDB.getRange(2, 1, sheetData.length, sheetData[0].length).setValues(sheetData);
}

// 削除マークのついた行を完全削除
function deleteMarkedRows() {
  const dialogResult = Browser.msgBox("削除マークの付いた行を完全に削除します。よろしいですか？", Browser.Buttons.YES_NO);
  if (dialogResult != 'yes') return;
  const examDataSheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.JUKEN_DB);
  const arrDeleteFlag = examDataSheet.getRange(1, EXAM_DATA.DEL_FLAG + 1, examDataSheet.getLastRow(), 1).getValues().flat();
  for (let i = arrDeleteFlag.length - 1; i > 0; i--) { // [0] はタイトル行
    if (arrDeleteFlag[i] == true) {
      examDataSheet.deleteRow(i + 1);
    }
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
  // キャッシュをクリア
  clearCache("UNIVERSITY_DATA");
  Browser.msgBox("消去が完了しました。");
}

// Benesseデータのインポート
function importUniversityData() {
  const dialogResult = Browser.msgBox("現在のシートから大学データを追記ます。\\n既存のデータは上書きされます。\\n\\nよろしいですか？", Browser.Buttons.YES_NO);
  if (dialogResult == 'no') return;
  const benesseSheet = activeSpreadsheet.getActiveSheet();
  const benesseDataRows = benesseSheet.getDataRange().getValues();
  if (benesseDataRows[0][BENESSE_DATA.CODE_DAIGAKU] != '大学ｺｰﾄﾞ' || benesseDataRows[0][BENESSE_DATA.CODE_GAKUBU] != '学部ｺｰﾄﾞ') {
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
    let universityName = row[BENESSE_DATA.NAME_DAIGAKU].trim();
    let facultyName = row[BENESSE_DATA.NAME_GAKUBU].trim();
    if (facultyName.length > 0) universityName += `・${facultyName}`;
    let gakka = row[BENESSE_DATA.NAME_GAKKA].trim();
    if (gakka.length > 0) universityName += `・${gakka}`;
    let schedule = row[BENESSE_DATA.NAME_NITTEI].trim();
    let method = row[BENESSE_DATA.NAME_HOUSHIKI].trim();
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
