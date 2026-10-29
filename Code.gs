/**
 * 設定：シート名と列の定義
 */
const CONFIG = {
  PRIVATE_SHEET_NAME: "Private",
  BIZ_SHEET_NAME: "Biz",
  OLD_SHEET_NAME: "タスク", // 移行用・後方互換
  DONE_SHEET_NAME: "完了タスク",
  PROJECT_SHEET_NAME: "プロジェクト",
  CONTEXT_SHEET_NAME: "コンテキスト",

  // 新しい列配置
  COL_ID: 1,           // A列: タスクID
  COL_STATUS: 2,       // B列: 完了
  COL_REPEAT: 3,       // C列: Repeat
  COL_TYPE: 4,         // D列: 項目 (Private/Biz)
  COL_PROJECT_NAME: 5, // E列: プロジェクト名 (ARRAYFORMULA数式)
  COL_PROJECT: 6,      // F列: プロジェクトID
  COL_TASK: 7,         // G列: タスク
  COL_DATE_E: 8,       // H列: 期日
  COL_DESCRIPTION: 9,  // I列: Context (コンテキスト)
  COL_START_TIME: 10,  // J列: 開始時間
  COL_END_TIME: 11,    // K列: 終了時間
};

/**
 * 対象がタスクシート（Private, Biz, 旧タスク）かどうかを判定
 */
function isTaskSheet(sheetName) {
  return sheetName === CONFIG.PRIVATE_SHEET_NAME || 
         sheetName === CONFIG.BIZ_SHEET_NAME || 
         sheetName === CONFIG.OLD_SHEET_NAME;
}

/**
 * 編集を検知して実行するトリガー用関数
 */
let cachedProjectTypes = null; // プロジェクトキャッシュ

function onEditTrigger(e) {
  const ss = e.source;
  const sheet = ss.getActiveSheet();
  const range = e.range;
  const sheetName = sheet.getName();
  const row = range.getRow();
  const col = range.getColumn();
  const val = range.getValue();

  if (row < 2) return; // ヘッダー行の編集は無視

  // --- 0. タスクID等のデフォルト値自動割り当て ---
  if (isTaskSheet(sheetName) && val !== "") {
    const idRange = sheet.getRange(row, CONFIG.COL_ID);
    if (!idRange.getValue()) {
      const lastCol = Math.max(sheet.getLastColumn(), CONFIG.COL_END_TIME);
      const rowRange = sheet.getRange(row, 1, 1, lastCol);
      const rowValues = rowRange.getValues()[0];
      
      let needsCheckbox = false;

      rowValues[CONFIG.COL_ID - 1] = Utilities.getUuid().split('-')[0];

      if (rowValues[CONFIG.COL_STATUS - 1] === "") {
        needsCheckbox = true;
      }

      if (rowValues[CONFIG.COL_REPEAT - 1] === "") {
        rowValues[CONFIG.COL_REPEAT - 1] = "None";
      }

      if (rowValues[CONFIG.COL_TYPE - 1] === "") {
        if (sheetName === CONFIG.BIZ_SHEET_NAME) {
          rowValues[CONFIG.COL_TYPE - 1] = "Biz";
        } else if (sheetName === CONFIG.PRIVATE_SHEET_NAME) {
          rowValues[CONFIG.COL_TYPE - 1] = "Private";
        } else {
          rowValues[CONFIG.COL_TYPE - 1] = getProjectTypeById(ss, rowValues[CONFIG.COL_PROJECT - 1]);
        }
      }

      // ARRAYFORMULAの展開を邪魔しないよう、E列の値を空にしてからセットする
      rowValues[CONFIG.COL_PROJECT_NAME - 1] = "";

      rowRange.setValues([rowValues]);

      if (needsCheckbox) {
        sheet.getRange(row, CONFIG.COL_STATUS).insertCheckboxes();
      }
    }
  }

  // --- 1. プロジェクトID(F列)変更時に項目(D列)を自動同期 ---
  if (isTaskSheet(sheetName) && col === CONFIG.COL_PROJECT) {
    const projectType = getProjectTypeById(ss, val);
    sheet.getRange(row, CONFIG.COL_TYPE).setValue(projectType);
  }

  // --- 2. タスク完了（B列チェック）時の処理 ---
  if (isTaskSheet(sheetName) && col === CONFIG.COL_STATUS) {
    if (val === true || String(val).toUpperCase() === "TRUE") {
      if (e.oldValue && String(e.oldValue).toUpperCase() === "TRUE") return;
      processTaskCompletion(ss, sheet, row);
      return;
    }
  }

  // --- 3. プロジェクト名変更の同期処理 ---
  if (sheetName === CONFIG.PROJECT_SHEET_NAME && col === 2) {
    const oldName = e.oldValue;
    const newName = e.value;
    if (oldName && newName && oldName !== newName) {
      syncProjectNameChange(ss, oldName, newName);
    }
  }
}

/**
 * プロジェクトIDから対応する項目 (Private/Biz) を取得するヘルパー関数
 */
function getProjectTypeById(ss, projectId) {
  if (!projectId) return "Private";
  
  if (!cachedProjectTypes) {
    cachedProjectTypes = new Map();
    const projectSheet = ss.getSheetByName(CONFIG.PROJECT_SHEET_NAME);
    if (projectSheet) {
      const values = projectSheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        const id = String(values[i][0]);
        if (id) {
          cachedProjectTypes.set(id, values[i][2] || "Private");
        }
      }
    }
  }
  
  return cachedProjectTypes.get(String(projectId)) || "Private";
}

/**
 * プロジェクトIDからプロジェクト名を取得するヘルパー関数
 */
function getProjectNameById(ss, projectId) {
  if (!projectId) return "";
  const projectSheet = ss.getSheetByName(CONFIG.PROJECT_SHEET_NAME);
  if (!projectSheet) return "";
  const values = projectSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(projectId).trim()) {
      return String(values[i][1] || "").trim();
    }
  }
  return "";
}

/**
 * タスク完了時のメイン処理
 */
function processTaskCompletion(ss, sheet, row) {
  // A列から一番右の列まで全てのデータを取得する
  const lastCol = Math.max(sheet.getLastColumn(), CONFIG.COL_END_TIME);
  const rowRange = sheet.getRange(row, 1, 1, lastCol);
  const rowData = rowRange.getValues()[0];

  // E列（プロジェクト名）が空の場合はプロジェクトIDから補完
  if (!rowData[CONFIG.COL_PROJECT_NAME - 1] && rowData[CONFIG.COL_PROJECT - 1]) {
    rowData[CONFIG.COL_PROJECT_NAME - 1] = getProjectNameById(ss, rowData[CONFIG.COL_PROJECT - 1]);
  }

  const repeatType = rowData[CONFIG.COL_REPEAT - 1];
  const sheetName = sheet.getName();

  // カレンダー登録とリピート作成を分離し、一方のエラーで止まらないように安全に処理
  // --- カレンダー＆リピート作成 ---
  try {
    syncToCalendarLogic(rowData, sheetName);
  } catch (err) {
    console.error("カレンダー作成エラー: " + err.message);
  }

  let repeatedTaskObj = null;
  try {
    const rtTypeStr = String(repeatType || "").trim();
    if (rtTypeStr !== "" && rtTypeStr.toUpperCase() !== "NONE") {
      repeatedTaskObj = handleRepeatTask(sheet, rowData, rtTypeStr);
    }
  } catch (err) {
    console.error("リピート作成エラー: " + err.message);
  }

  // --- 完了転記と削除 ---
  try {
    let doneSheet = ss.getSheetByName(CONFIG.DONE_SHEET_NAME);
    if (!doneSheet) {
      doneSheet = ss.insertSheet(CONFIG.DONE_SHEET_NAME);
      setupDoneSheetFormat(doneSheet);
    }

    // 常に2行目の位置に新しい空行を自動挿入（最新の完了タスクを一番上に配置するため）
    doneSheet.insertRowBefore(2);

    // ARRAYFORMULAの展開を妨げないようE列(COL_PROJECT_NAME)の値を空にする
    const newDoneRow = [...rowData];
    newDoneRow[CONFIG.COL_PROJECT_NAME - 1] = "";
    newDoneRow[CONFIG.COL_STATUS - 1] = true;

    // 2行目に一括転記
    doneSheet.getRange(2, 1, 1, newDoneRow.length).setValues([newDoneRow]);

    // 元のシートから該当行を削除
    sheet.deleteRow(row);

    // リピートタスクが新規追加された場合のみ並び替えを実行（単なる行削除時は既存の並び順が維持されるため不要）
    if (repeatedTaskObj) {
      sortAll(sheet);
    }
  } catch (err) {
    console.error("転記/削除エラー: " + err.message);
  }

  return repeatedTaskObj;
}

/**
 * リピートタスクを末尾に新規作成
 */
function handleRepeatTask(sheet, rowData, repeatType) {
  const rType = String(repeatType).toUpperCase();
  const rawDate = rowData[CONFIG.COL_DATE_E - 1];
  let nextDate = rawDate;

  if (rType === "COPY") {
    // Copyの場合は期日を変更なし（そのまま維持）
    nextDate = rawDate;
  } else {
    const currentDate = new Date(rawDate);
    if (isNaN(currentDate.getTime())) {
      return null;
    }
    const calculatedDate = new Date(currentDate.getTime());
    if (rType === "DAILY") calculatedDate.setDate(calculatedDate.getDate() + 1);
    else if (rType === "WEEKLY") calculatedDate.setDate(calculatedDate.getDate() + 7);
    else if (rType === "BIWEEKLY") calculatedDate.setDate(calculatedDate.getDate() + 14);
    else if (rType === "MONTHLY") calculatedDate.setMonth(calculatedDate.getMonth() + 1);
    nextDate = calculatedDate;
  }

  const newRowData = [...rowData];
  newRowData[CONFIG.COL_ID - 1] = Utilities.getUuid().split('-')[0]; // 新しいIDを発行
  newRowData[CONFIG.COL_PROJECT_NAME - 1] = ""; 
  newRowData[CONFIG.COL_DATE_E - 1] = nextDate; // 期日
  newRowData[CONFIG.COL_STATUS - 1] = false;  // 完了フラグを外す (false)

  // Copyの場合は新タスクのRepeat設定を "None" に戻す
  if (rType === "COPY") {
    newRowData[CONFIG.COL_REPEAT - 1] = "None";
  }

  // 開始時間と終了時間のコピーは不要のためクリアする
  newRowData[CONFIG.COL_START_TIME - 1] = ""; 
  newRowData[CONFIG.COL_END_TIME - 1] = "";

  // 本当にデータが有る行の下に追加
  const targetRow = getRealLastRow(sheet) + 1;
  sheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);
  sheet.getRange(targetRow, CONFIG.COL_DATE_E).setNumberFormat("M/d(ddd)");

  const taskObj = parseRowToTask(newRowData);
  taskObj.projectName = String(rowData[CONFIG.COL_PROJECT_NAME - 1] || "");
  return taskObj;
}

/**
 * ARRAYFORMULAなどで空白に見えるセルを除き、A列またはG列で本当に値が入っている最後の行を取得する関数
 */
function getRealLastRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 1;

  const values = sheet.getRange(1, 1, lastRow, CONFIG.COL_TASK).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const idVal = values[i][CONFIG.COL_ID - 1];
    const taskVal = values[i][CONFIG.COL_TASK - 1];
    if (idVal !== "" || taskVal !== "") {
      return i + 1; // 値が入っている最後の行番号を返す
    }
  }
  return 1; // データがなければ1行目を返す
}

/**
 * カレンダー登録ロジック（PrivateはGoogleカレンダー、BizはOutlookへICSメール送信）
 */
function syncToCalendarLogic(rowData, sheetName) {
  const projectName = rowData[CONFIG.COL_PROJECT_NAME - 1];
  const taskName = rowData[CONFIG.COL_TASK - 1];
  const startDate = rowData[CONFIG.COL_DATE_E - 1]; // カレンダー基準日は「期日」
  const startTime = rowData[CONFIG.COL_START_TIME - 1];
  const endTime = rowData[CONFIG.COL_END_TIME - 1];
  const description = rowData[CONFIG.COL_DESCRIPTION - 1];
  const itemType = String(rowData[CONFIG.COL_TYPE - 1] || "").trim();

  const title = [projectName, taskName].filter(Boolean).join(" ");
  const start = combineDateTime(startDate, startTime);
  const end = combineDateTime(startDate, endTime);

  if (!start || !end) return;

  // Bizタスク判定 (シート名がBiz または 項目がBiz)
  const isBiz = (sheetName === CONFIG.BIZ_SHEET_NAME) || (itemType.toUpperCase() === "BIZ");

  if (isBiz) {
    // 会社用Outlook宛にICS招待メールを送信
    sendIcsToOutlook(title, start, end, description);
  } else {
    // プライベートは従来通り Google カレンダーへ登録
    CalendarApp.getDefaultCalendar().createEvent(title, start, end, { description: description });
  }
}

/**
 * 会社のOutlookメール宛に iCalendar (ICS) 形式の予定メールを送信
 */
function sendIcsToOutlook(title, start, end, description) {
  const outlookEmail = PropertiesService.getScriptProperties().getProperty("OUTLOOK_EMAIL");

  if (!outlookEmail) {
    console.warn("【警告】スクリプトプロパティ 'OUTLOOK_EMAIL' が設定されていないため、Outlookへの予定送信をスキップしました。");
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast("Outlookメールアドレス(OUTLOOK_EMAIL)が未設定のため予定メールは送信されませんでした。", "通知");
    } catch (e) {}
    return;
  }

  try {
    const icsContent = createIcsString(title, start, end, description);
    const blob = Utilities.newBlob(icsContent, "text/calendar; charset=UTF-8; method=REQUEST", "event.ics");

    const formattedStart = Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
    const formattedEnd = Utilities.formatDate(end, Session.getScriptTimeZone(), "HH:mm");

    const body = [
      "タスクが完了しました。Outlookカレンダーに予定を登録してください。",
      "",
      `■ タスク名: ${title}`,
      `■ 日時: ${formattedStart} 〜 ${formattedEnd}`,
      `■ メモ: ${description || "（なし）"}`,
      "",
      "※添付の .ics ファイルを開くか、メール内の「承諾/カレンダーに追加」から登録できます。"
    ].join("\n");

    GmailApp.sendEmail(outlookEmail, `【完了タスク予定】${title}`, body, {
      attachments: [blob],
      name: "(todo)"
    });

    console.log(`Outlook (${outlookEmail}) 宛にICS予定メールを送信しました: ${title}`);
  } catch (e) {
    console.error("Outlookメール送信エラー: " + e.message);
  }
}

/**
 * iCalendar (RFC 5545) 形式の文字列を生成
 */
function createIcsString(title, start, end, description) {
  const dtStamp = formatIcsDate(new Date());
  const dtStart = formatIcsDate(start);
  const dtEnd = formatIcsDate(end);
  const uid = Utilities.getUuid() + "@gas-task-manager";

  return [
    "BEGIN:VCALENDAR",
    "PRODID:-//(todo)//JA",
    "VERSION:2.0",
    "CALSCALE:GREGORIAN",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    `UID:${uid}`,
    `DTSTAMP:${dtStamp}`,
    `DTSTART:${dtStart}`,
    `DTEND:${dtEnd}`,
    `SUMMARY:${escapeIcsText(title)}`,
    `DESCRIPTION:${escapeIcsText(description || "")}`,
    "ORGANIZER;CN=(todo):mailto:no-reply@gas-todo",
    "STATUS:CONFIRMED",
    "SEQUENCE:0",
    "END:VEVENT",
    "END:VCALENDAR"
  ].join("\r\n");
}

function formatIcsDate(date) {
  return Utilities.formatDate(date, "UTC", "yyyyMMdd'T'HHmmss'Z'");
}

function escapeIcsText(text) {
  return String(text || "")
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\n/g, "\\n");
}

/**
 * プロジェクト名変更時の同期処理
 * E列はE1セルのARRAYFORMULA（VLOOKUP）で自動参照されるため、E2以降の値クリアのみで安全に最新化される
 */
function syncProjectNameChange(ss, oldName, newName) {
  const sheetsToUpdate = [CONFIG.PRIVATE_SHEET_NAME, CONFIG.BIZ_SHEET_NAME, CONFIG.OLD_SHEET_NAME, CONFIG.DONE_SHEET_NAME];

  sheetsToUpdate.forEach(name => {
    const targetSheet = ss.getSheetByName(name);
    if (!targetSheet) return;

    const lastRow = targetSheet.getLastRow();
    if (lastRow < 2) return;

    targetSheet.getRange(2, CONFIG.COL_PROJECT_NAME, lastRow - 1, 1).clearContent();
  });
}

/**
 * 補助関数：日付と時間を結合
 */
function combineDateTime(date, time) {
  if (!(date instanceof Date)) return null;
  const dateTime = new Date(date.getTime());
  if (time instanceof Date) {
    dateTime.setHours(time.getHours(), time.getMinutes(), 0);
  } else {
    dateTime.setHours(0, 0, 0);
  }
  return dateTime;
}

// --- メニューとソート機能 ---
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('タスク管理')
    .addItem('並び替え（Bizシート）', 'sortBizSheet')
    .addItem('並び替え（Privateシート）', 'sortPrivateSheet')
    .addItem('並び替え（全シート）', 'sortAllSheets')
    .addToUi();
}

/**
 * 各タスクシートのヘッダーとE1セル(ARRAYFORMULA)のフォーマットを保証する関数
 */
function setupTaskSheetFormat(sheet) {
  if (!sheet) return;

  const headers = [
    ["タスクID", "完了", "Repeat", "項目", "プロジェクト名", "プロ...", "タスク", "期日", "Context", "開始", "終了"]
  ];
  if (sheet.getLastRow() < 1 || sheet.getRange(1, 1).getValue() === "") {
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  }

  const e1Range = sheet.getRange(1, CONFIG.COL_PROJECT_NAME);
  if (!e1Range.getFormula()) {
    e1Range.setFormula(`={"プロジェクト名"; ARRAYFORMULA(IF(F2:F="","",IFERROR(VLOOKUP(F2:F, 'プロジェクト'!A:B, 2, FALSE),"")))}`);
  }
}

/**
 * 完了タスクシートのヘッダーとE1セル(ARRAYFORMULA)のフォーマットを保証する関数
 */
function setupDoneSheetFormat(doneSheet) {
  setupTaskSheetFormat(doneSheet);
}

/**
 * Privateシートのみを並び替え
 */
function sortPrivateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.PRIVATE_SHEET_NAME);
  if (sheet) {
    sortByDateAndContext(sheet);
  } else {
    ss.toast("Privateシートが見つかりません。", "エラー");
  }
}

/**
 * Bizシートのみを並び替え
 */
function sortBizSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.BIZ_SHEET_NAME);
  if (sheet) {
    sortByDateAndContext(sheet);
  } else {
    ss.toast("Bizシートが見つかりません。", "エラー");
  }
}

/**
 * 指定シート（またはアクティブシート）を「完了状態 → 期日 (昇順) → 開始時間 → コンテキスト (ID/マスター定義順)」で高速に並び替える関数
 * ※ 毎回フィルターを削除・再作成すると1〜2秒の遅延が発生するため、既存フィルターを保持したままデータ領域のみを一括更新する
 */
let cachedContextOrderMap = null;

function getContextOrderMap(ss) {
  if (cachedContextOrderMap) return cachedContextOrderMap;
  cachedContextOrderMap = new Map([
    ["最優先", 100],
    ["優先", 101],
    ["通常", 102],
    ["翌日以降可", 103],
    ["連絡待ち", 104],
    ["簡単", 105],
    ["隙間時間", 106]
  ]);
  const contextSheet = ss.getSheetByName(CONFIG.CONTEXT_SHEET_NAME);
  if (contextSheet) {
    const cValues = contextSheet.getDataRange().getValues();
    for (let i = 1; i < cValues.length; i++) {
      const cId = Number(cValues[i][0]) || (i * 10);
      const cName = String(cValues[i][1]).trim();
      if (cName) {
        cachedContextOrderMap.set(cName, cId);
      }
    }
  }
  return cachedContextOrderMap;
}

function sortTaskRowsInMemory(values, contextOrderMap) {
  values.sort((a, b) => {
    // 1. 完了状態 (B列)
    const statusA = (a[CONFIG.COL_STATUS - 1] === true || String(a[CONFIG.COL_STATUS - 1]).toUpperCase() === "TRUE") ? 1 : 0;
    const statusB = (b[CONFIG.COL_STATUS - 1] === true || String(b[CONFIG.COL_STATUS - 1]).toUpperCase() === "TRUE") ? 1 : 0;
    if (statusA !== statusB) return statusA - statusB;

    // 2. 日付 / 期日 (H列)
    const dateA = a[CONFIG.COL_DATE_E - 1] instanceof Date && !isNaN(a[CONFIG.COL_DATE_E - 1]) ? a[CONFIG.COL_DATE_E - 1].getTime() : Infinity;
    const dateB = b[CONFIG.COL_DATE_E - 1] instanceof Date && !isNaN(b[CONFIG.COL_DATE_E - 1]) ? b[CONFIG.COL_DATE_E - 1].getTime() : Infinity;
    if (dateA !== dateB) return dateA - dateB;

    // 3. 開始時間 (J列)
    const timeA = a[CONFIG.COL_START_TIME - 1] instanceof Date && !isNaN(a[CONFIG.COL_START_TIME - 1]) ? a[CONFIG.COL_START_TIME - 1].getTime() : Infinity;
    const timeB = b[CONFIG.COL_START_TIME - 1] instanceof Date && !isNaN(b[CONFIG.COL_START_TIME - 1]) ? b[CONFIG.COL_START_TIME - 1].getTime() : Infinity;
    if (timeA !== timeB) return timeA - timeB;

    // 4. コンテキスト (I列: ID/マスター定義順)
    const ctxNameA = String(a[CONFIG.COL_DESCRIPTION - 1] || "").trim();
    const ctxNameB = String(b[CONFIG.COL_DESCRIPTION - 1] || "").trim();
    const orderA = contextOrderMap.has(ctxNameA) ? contextOrderMap.get(ctxNameA) : Infinity;
    const orderB = contextOrderMap.has(ctxNameB) ? contextOrderMap.get(ctxNameB) : Infinity;
    return orderA - orderB;
  });

  // ARRAYFORMULAの展開を邪魔しないよう、E列(COL_PROJECT_NAME)の値をすべて空文字にする
  for (let i = 0; i < values.length; i++) {
    values[i][CONFIG.COL_PROJECT_NAME - 1] = "";
  }
  return values;
}

function sortByDateAndContext(targetSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = targetSheet || ss.getActiveSheet();
  if (!sheet || !isTaskSheet(sheet.getName())) return;

  const sheetLastRow = sheet.getLastRow();
  if (sheetLastRow < 2) return;
  const lastCol = Math.max(sheet.getLastColumn(), CONFIG.COL_END_TIME);

  const allValues = sheet.getRange(2, 1, sheetLastRow - 1, lastCol).getValues();
  let realCount = 0;
  for (let i = allValues.length - 1; i >= 0; i--) {
    if (allValues[i][CONFIG.COL_ID - 1] !== "" || allValues[i][CONFIG.COL_TASK - 1] !== "") {
      realCount = i + 1;
      break;
    }
  }
  if (realCount < 2) return;

  const values = allValues.slice(0, realCount);
  const contextOrderMap = getContextOrderMap(ss);
  sortTaskRowsInMemory(values, contextOrderMap);

  sheet.getRange(2, 1, realCount, lastCol).setValues(values);
}

function sortAll(sheet) {
  try {
    if (sheet) {
      sortByDateAndContext(sheet);
    } else {
      sortAllSheets();
    }
  } catch (err) {
    console.warn("並び替え処理スキップ: " + err.message);
  }
}

function sortAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [CONFIG.PRIVATE_SHEET_NAME, CONFIG.BIZ_SHEET_NAME].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) sortByDateAndContext(s);
  });
}

/**
 * 外部Webアプリ（GitHub Pages等）からのAPI処理共通ディスパッチャー
 * 更新系リクエストは LockService で排他制御し、同時リクエストによるデータの先祖返り・競合を完全防止する
 */
function handleApiAction(action, payload) {
  payload = payload || {};

  if (action === 'getInitialData') {
    return getInitialData();
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    let result;
    if (action === 'saveTask') {
      result = upsertTaskFromApp(payload.taskData);
    } else if (action === 'toggleTaskStatus' || action === 'completeTask') {
      result = completeTaskFromApp(payload.taskId, payload.taskName, payload.taskDate);
    } else if (action === 'uncompleteTask') {
      result = uncompleteTaskFromApp(payload.taskId, payload.taskData);
    } else if (action === 'deleteTask') {
      result = deleteTaskFromApp(payload.taskId, payload.taskName, payload.taskDate);
    } else if (action === 'saveProject') {
      if (payload.projectData && payload.projectData.id && !payload.projectData.isNew && !String(payload.projectData.id).startsWith('p_')) {
        result = updateProjectFromApp(payload.projectData.id, payload.projectData.name, payload.projectData.type);
      } else {
        result = addNewProjectFromApp(payload.projectData);
      }
    } else if (action === 'deleteProject') {
      result = deleteProjectFromApp(payload.projectId, payload.projectName);
    } else if (action === 'saveContext') {
      result = saveContextFromApp(payload.contextData);
    } else if (action === 'deleteContext') {
      result = deleteContextFromApp(payload.contextId);
    } else if (action === 'updateSortOrder') {
      result = updateSortOrderFromApp(payload.idList);
    } else {
      throw new Error('未定義のアクションです: ' + action);
    }
    return result;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Webアプリのエントリポイント（HTML配信 & Web API GETハンドラー）
 */
function doGet(e) {
  // 1. 外部Webアプリ（GitHub Pages等）からのAPIリクエスト処理 (GET)
  if (e && e.parameter && e.parameter.action) {
    try {
      const action = e.parameter.action;
      let payload = {};
      if (e.parameter.payload) {
        try {
          payload = JSON.parse(e.parameter.payload);
        } catch (jsonErr) {
          payload = {};
        }
      } else {
        payload = { ...e.parameter };
      }

      const result = handleApiAction(action, payload);
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 2. 従来のGASウェブアプリとして直接開かれた場合はHTMLを配信
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('プロジェクト・タスク管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 外部WebアプリからのPOSTリクエストハンドラー（追加・更新・削除・並び替え）
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('リクエストボディが空です');
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const result = handleApiAction(action, payload);

    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * アプリ起動時の初期データ（タスク、プロジェクト、完了履歴）を一括取得
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. プロジェクトデータ取得（名前が入力されている行のみ対象）
  const projectSheet = ss.getSheetByName(CONFIG.PROJECT_SHEET_NAME);
  const projects = [];
  const idToProjectName = {};
  const idToProjectType = {};
  if (projectSheet) {
    const values = projectSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const pId = String(values[i][0] || "").trim();
      const pName = String(values[i][1] || "").trim();
      const pType = String(values[i][2] || "Private").trim();
      if (pId !== "" && pName !== "") {
        projects.push({
          id: pId,
          name: pName,
          type: pType
        });
        idToProjectName[pId] = pName;
        idToProjectType[pId] = pType;
      }
    }
  }
  
  // 1-2. コンテキストデータ取得（名前が入力されている行のみ対象）
  const contextSheet = ss.getSheetByName(CONFIG.CONTEXT_SHEET_NAME);
  const contexts = [];
  const idToContextName = {};
  if (contextSheet) {
    const values = contextSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const idVal = String(values[i][0] || "").trim();
      const nameVal = String(values[i][1] || "").trim();
      if (nameVal !== "") {
        contexts.push({
          id: idVal || nameVal,
          name: nameVal
        });
        if (idVal) {
          idToContextName[idVal] = nameVal;
        }
      }
    }
  }

  // 2. タスクデータ取得 (Private, Bizシートからタスクを取得)
  const tasks = [];
  const taskSheetNames = [CONFIG.PRIVATE_SHEET_NAME, CONFIG.BIZ_SHEET_NAME];
  const lingeringDoneTasks = [];
  const seenActiveIds = new Set();
  
  taskSheetNames.forEach(name => {
    const taskSheet = ss.getSheetByName(name);
    if (taskSheet) {
      const lastRow = taskSheet.getLastRow();
      if (lastRow < 2) return;
      const lastCol = Math.max(taskSheet.getLastColumn(), CONFIG.COL_END_TIME);
      const values = taskSheet.getRange(1, 1, lastRow, lastCol).getValues();
      for (let i = 1; i < values.length; i++) {
        let currentId = String(values[i][CONFIG.COL_ID - 1] || "").trim();
        const taskName = String(values[i][CONFIG.COL_TASK - 1] || "").trim();
        
        // タスク名が存在する有効な行を対象とする
        if (taskName !== "") {
          // IDが空欄またはtmp_で始まっている場合、正規の8桁UUIDを発行してシートのA列を恒久修復
          if (!currentId || currentId.startsWith("tmp_")) {
            currentId = Utilities.getUuid().split("-")[0];
            taskSheet.getRange(i + 1, CONFIG.COL_ID).setValue(currentId);
            values[i][CONFIG.COL_ID - 1] = currentId;
          }

          const taskObj = parseRowToTask(values[i]);
          // ARRAYFORMULAの再計算前でも確実にプロジェクト名を解決
          if (!taskObj.projectName && taskObj.projectId && idToProjectName[taskObj.projectId]) {
            taskObj.projectName = idToProjectName[taskObj.projectId];
          }
          // コンテキストIDが入っている場合は名前に解決
          if (taskObj.description && idToContextName[taskObj.description]) {
            taskObj.description = idToContextName[taskObj.description];
          }
          // シート名に基づくデフォルト補正（未指定の場合）
          if (name === CONFIG.BIZ_SHEET_NAME && !taskObj.type) {
            taskObj.type = "Biz";
          } else if (taskObj.projectId && idToProjectType[taskObj.projectId] && !values[i][CONFIG.COL_TYPE - 1]) {
            taskObj.type = idToProjectType[taskObj.projectId];
          }
          // 未完了タスクシート内に完了済み行が残っていた場合は未完了一覧に含めず完了側に退避
          if (taskObj.status) {
            lingeringDoneTasks.push(taskObj);
          } else if (!seenActiveIds.has(currentId)) {
            tasks.push(taskObj);
            seenActiveIds.add(currentId);
          }
        }
      }
    }
  });
  
  // 3. 完了タスクデータ取得 (先頭2行目以降の最新履歴から最大500件を高速取得)
  const doneSheet = ss.getSheetByName(CONFIG.DONE_SHEET_NAME);
  const doneTasks = [];
  const seenDoneIds = new Set();

  // 未完了シートに残っていた完了タスクがあれば優先追加
  lingeringDoneTasks.forEach(dt => {
    doneTasks.push(dt);
    seenDoneIds.add(dt.id);
  });

  if (doneSheet) {
    const doneLastRow = doneSheet.getLastRow();
    if (doneLastRow >= 2) {
      const maxRows = Math.min(doneLastRow, 501);
      const lastCol = Math.max(doneSheet.getLastColumn(), CONFIG.COL_END_TIME);
      const values = doneSheet.getRange(1, 1, maxRows, lastCol).getValues();
      for (let i = 1; i < values.length; i++) {
        let rowId = String(values[i][CONFIG.COL_ID - 1] || "").trim();
        const rowTaskName = String(values[i][CONFIG.COL_TASK - 1] || "").trim();
        if (rowTaskName !== "") {
          if (!rowId || rowId.startsWith("tmp_")) {
            rowId = Utilities.getUuid().split("-")[0];
            doneSheet.getRange(i + 1, CONFIG.COL_ID).setValue(rowId);
            values[i][CONFIG.COL_ID - 1] = rowId;
          }
          if (!seenDoneIds.has(rowId) && !seenActiveIds.has(rowId)) {
            const taskObj = parseRowToTask(values[i]);
            taskObj.status = true;
            if (!taskObj.projectName && taskObj.projectId && idToProjectName[taskObj.projectId]) {
              taskObj.projectName = idToProjectName[taskObj.projectId];
            }
            if (taskObj.description && idToContextName[taskObj.description]) {
              taskObj.description = idToContextName[taskObj.description];
            }
            doneTasks.push(taskObj);
            seenDoneIds.add(rowId);
          }
        }
      }
    }
  }
  
  return {
    tasks: tasks,
    projects: projects,
    contexts: contexts,
    doneTasks: doneTasks
  };
}

/**
 * 行データをフロントエンド用のタスクオブジェクトに変換
 */
function parseRowToTask(row) {
  return {
    id: String(row[CONFIG.COL_ID - 1]),
    status: row[CONFIG.COL_STATUS - 1] === true || String(row[CONFIG.COL_STATUS - 1]).toUpperCase() === "TRUE",
    repeat: String(row[CONFIG.COL_REPEAT - 1] || "None"),
    type: String(row[CONFIG.COL_TYPE - 1] || "Private"),
    projectName: String(row[CONFIG.COL_PROJECT_NAME - 1] || ""),
    projectId: String(row[CONFIG.COL_PROJECT - 1] || ""),
    task: String(row[CONFIG.COL_TASK - 1] || ""),
    date: formatDate(row[CONFIG.COL_DATE_E - 1]),
    startTime: formatTime(row[CONFIG.COL_START_TIME - 1]),
    endTime: formatTime(row[CONFIG.COL_END_TIME - 1]),
    description: String(row[CONFIG.COL_DESCRIPTION - 1] || "")
  };
}

/**
 * 日付オブジェクトを yyyy-MM-dd フォーマットの文字列に超高速変換（V8ネイティブ処理・IPC呼び出しゼロ）
 */
function formatDate(dateVal) {
  if (dateVal instanceof Date) {
    const t = dateVal.getTime();
    if (isNaN(t)) return "";
    const y = dateVal.getFullYear();
    const m = String(dateVal.getMonth() + 1).padStart(2, "0");
    const d = String(dateVal.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  return String(dateVal || "");
}

/**
 * 時間オブジェクトを HH:mm フォーマットの文字列に超高速変換（V8ネイティブ処理・IPC呼び出しゼロ）
 */
function formatTime(timeVal) {
  if (timeVal instanceof Date) {
    const t = timeVal.getTime();
    if (isNaN(t)) return "";
    const hh = String(timeVal.getHours()).padStart(2, "0");
    const mm = String(timeVal.getMinutes()).padStart(2, "0");
    return `${hh}:${mm}`;
  }
  return String(timeVal || "");
}

/**
 * 指定のタスクタイプに応じたシートを取得（なければ自動作成・フォーマット）
 */
function getOrCreateTaskSheetByType(ss, type) {
  const isBiz = String(type || "").toUpperCase() === "BIZ";
  const targetSheetName = isBiz ? CONFIG.BIZ_SHEET_NAME : CONFIG.PRIVATE_SHEET_NAME;
  
  let sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(targetSheetName);
    setupTaskSheetFormat(sheet);
  }
  return sheet;
}

/**
 * タスクオブジェクトからスプレッドシート1行分の配列を生成する内部関数
 */
function buildTaskRowArray(taskData, taskId, taskType) {
  const dateVal = taskData.date ? new Date(taskData.date.replace(/-/g, "/")) : "";
  const baseDateStr = taskData.date || formatDate(new Date());
  const startTimeVal = taskData.startTime ? parseTimeStringToDate(baseDateStr, taskData.startTime) : "";
  const endTimeVal = taskData.endTime ? parseTimeStringToDate(baseDateStr, taskData.endTime) : "";

  const rowArr = Array(CONFIG.COL_END_TIME).fill("");
  rowArr[CONFIG.COL_ID - 1] = taskId;
  rowArr[CONFIG.COL_STATUS - 1] = Boolean(taskData.status);
  rowArr[CONFIG.COL_REPEAT - 1] = taskData.repeat || "None";
  rowArr[CONFIG.COL_TYPE - 1] = taskType;
  rowArr[CONFIG.COL_PROJECT_NAME - 1] = ""; // ARRAYFORMULA数式との衝突防止のため空文字
  rowArr[CONFIG.COL_PROJECT - 1] = taskData.projectId || "";
  rowArr[CONFIG.COL_TASK - 1] = taskData.task || "";
  rowArr[CONFIG.COL_DATE_E - 1] = dateVal;
  rowArr[CONFIG.COL_DESCRIPTION - 1] = taskData.description || "";
  rowArr[CONFIG.COL_START_TIME - 1] = startTimeVal;
  rowArr[CONFIG.COL_END_TIME - 1] = endTimeVal;
  return rowArr;
}

/**
 * 指定シートの有効データ領域を1回のgetValuesで取得し、メモリ上で追加/更新＆ソートして1回のsetValuesで書き戻す超高速関数
 * （Write→Read→Writeによるスプレッドシート全体再計算フラッシュをゼロにする）
 */
function upsertAndSortInSheetMemory(ss, sheet, rowArr, targetId, allowInsert) {
  const sheetLastRow = sheet.getLastRow();
  const lastCol = CONFIG.COL_END_TIME;
  let activeRows = [];

  if (sheetLastRow >= 2) {
    const allValues = sheet.getRange(2, 1, sheetLastRow - 1, lastCol).getValues();
    let realCount = 0;
    for (let i = allValues.length - 1; i >= 0; i--) {
      if (allValues[i][CONFIG.COL_ID - 1] !== "" || allValues[i][CONFIG.COL_TASK - 1] !== "") {
        realCount = i + 1;
        break;
      }
    }
    if (realCount > 0) {
      activeRows = allValues.slice(0, realCount).map(r => r.slice(0, lastCol));
    }
  }

  let foundIndex = -1;
  if (targetId) {
    for (let i = 0; i < activeRows.length; i++) {
      if (String(activeRows[i][CONFIG.COL_ID - 1]).trim() === targetId) {
        foundIndex = i;
        break;
      }
    }
  }

  if (foundIndex === -1 && !allowInsert) {
    return false;
  }

  if (foundIndex >= 0) {
    activeRows[foundIndex] = rowArr;
  } else {
    activeRows.push(rowArr);
  }

  const contextOrderMap = getContextOrderMap(ss);
  sortTaskRowsInMemory(activeRows, contextOrderMap);

  sheet.getRange(2, 1, activeRows.length, lastCol).setValues(activeRows);
  return foundIndex >= 0 ? "updated" : "created";
}

/**
 * タスクのUpsert（メモリ上一括更新・追加・ソートによりRead-after-Writeフラッシュを排除した超高速版）
 */
function upsertTaskFromApp(taskData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!taskData || !taskData.task || String(taskData.task).trim() === "") {
    return { ok: false };
  }

  const rawId = taskData.id ? String(taskData.id).trim() : "";
  const hasValidId = rawId !== "" && !rawId.startsWith("tmp_");
  const targetId = hasValidId ? rawId : Utilities.getUuid().split("-")[0];
  const taskType = taskData.type || getProjectTypeById(ss, taskData.projectId);
  const targetSheet = getOrCreateTaskSheetByType(ss, taskType);
  const rowArr = buildTaskRowArray(taskData, targetId, taskType);

  // 1. まず別シート（Private <-> Biz）に同一IDが存在していないか確認（区分変更時の移動対応）
  if (hasValidId) {
    const otherSheetName = targetSheet.getName() === CONFIG.BIZ_SHEET_NAME ? CONFIG.PRIVATE_SHEET_NAME : CONFIG.BIZ_SHEET_NAME;
    const otherSheet = ss.getSheetByName(otherSheetName);
    if (otherSheet) {
      const otherLastRow = otherSheet.getLastRow();
      if (otherLastRow >= 2) {
        const otherIds = otherSheet.getRange(2, CONFIG.COL_ID, otherLastRow - 1, 1).getValues();
        for (let i = otherIds.length - 1; i >= 0; i--) {
          if (String(otherIds[i][0]).trim() === targetId) {
            otherSheet.deleteRow(i + 2);
          }
        }
      }
    }
  }

  // 2. 対象シート上でメモリ内Upsert＋ソート＋1回の一括書き込みを実行
  const mode = upsertAndSortInSheetMemory(ss, targetSheet, rowArr, targetId, true);
  if (mode === "updated") {
    return { ok: true, updatedTaskId: targetId };
  }
  return { ok: true, createdTaskId: targetId };
}

/**
 * アプリから新規タスクを追加（Private/Bizシートへ自動振り分け）
 */
function addNewTaskFromApp(taskData) {
  return upsertTaskFromApp(taskData);
}

/**
 * 既存行を1回のsetValuesで高速に一括更新する内部関数
 */
function updateTaskRowFast(ss, sheet, row, taskData) {
  return upsertTaskFromApp(taskData);
}

/**
 * アプリからタスクを編集（後方互換ラッパー）
 */
function updateTaskFromApp(taskData) {
  return upsertTaskFromApp(taskData);
}

/**
 * 日付文字列と時間文字列からDateオブジェクトを生成
 */
function parseTimeStringToDate(dateString, timeString) {
  if (!timeString) return "";
  const base = dateString ? dateString.replace(/-/g, "/") : formatDate(new Date()).replace(/-/g, "/");
  const d = new Date(base);
  if (isNaN(d.getTime())) return "";
  const parts = String(timeString).split(":");
  if (parts.length < 2) return "";
  d.setHours(parseInt(parts[0], 10), parseInt(parts[1], 10), 0, 0);
  return d;
}

/**
 * タスクIDから該当シートと行番号を検索（preferredTypeが指定されていれば該当シートを最優先で検索）
 */
function findTaskLocationById(ss, taskId, fallbackTaskName, fallbackDate, includeDoneSheet, preferredType) {
  const isBiz = String(preferredType || "").toUpperCase() === "BIZ";
  const candidateNames = isBiz
    ? [CONFIG.BIZ_SHEET_NAME, CONFIG.PRIVATE_SHEET_NAME]
    : [CONFIG.PRIVATE_SHEET_NAME, CONFIG.BIZ_SHEET_NAME];
  if (includeDoneSheet) {
    candidateNames.push(CONFIG.DONE_SHEET_NAME);
  }
  
  // 1. まず taskId で完全一致検索
  if (taskId) {
    const targetIdStr = String(taskId).trim();
    for (let name of candidateNames) {
      const sheet = ss.getSheetByName(name);
      if (!sheet) continue;

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) continue;

      const ids = sheet.getRange(2, CONFIG.COL_ID, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === targetIdStr) {
          return { sheet: sheet, row: i + 2 };
        }
      }
    }
  }

  // 2. taskId で見つからなかった場合、タスク名（および期日）でフォールバック検索（A列空欄行等の救済）
  if (fallbackTaskName) {
    const targetName = String(fallbackTaskName).trim();
    let bestMatch = null;

    for (let name of candidateNames) {
      const sheet = ss.getSheetByName(name);
      if (!sheet) continue;

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) continue;

      // A列(ID), G列(Task), H列(Date) をまとめて取得
      const rangeValues = sheet.getRange(2, 1, lastRow - 1, CONFIG.COL_DATE_E).getValues();
      for (let i = 0; i < rangeValues.length; i++) {
        const rowData = rangeValues[i];
        const rowTaskName = String(rowData[CONFIG.COL_TASK - 1] || "").trim();

        if (rowTaskName === targetName) {
          const row = i + 2;
          const currentId = String(rowData[CONFIG.COL_ID - 1] || "").trim();
          const rowDateStr = formatDate(rowData[CONFIG.COL_DATE_E - 1]);
          const dateMatches = !fallbackDate || (rowDateStr === String(fallbackDate).trim());

          // 日付も一致するものを最優先（IDが空欄なら新IDを採番して返す）
          if (dateMatches) {
            if (!currentId) {
              sheet.getRange(row, CONFIG.COL_ID).setValue(taskId || Utilities.getUuid().split('-')[0]);
            }
            return { sheet: sheet, row: row };
          }

          // 日付が合わなくてもタスク名が一致するものを第2候補として保持
          if (!bestMatch) {
            bestMatch = { sheet: sheet, row: row, needId: !currentId };
          }
        }
      }
    }

    if (bestMatch) {
      if (bestMatch.needId) {
        bestMatch.sheet.getRange(bestMatch.row, CONFIG.COL_ID).setValue(taskId || Utilities.getUuid().split('-')[0]);
      }
      return { sheet: bestMatch.sheet, row: bestMatch.row };
    }
  }

  return null;
}

/**
 * アプリからタスクを完了処理（IDで見つからない場合はタスク名と期日でフォールバック検索）
 */
function completeTaskFromApp(taskId, taskName, taskDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const location = findTaskLocationById(ss, taskId, taskName, taskDate, false);
  
  if (!location) {
    return { ok: true, completedTaskId: taskId, alreadyCompleted: true };
  }
  
  const { sheet, row } = location;
  const repeatedTask = processTaskCompletion(ss, sheet, row);
  
  return { ok: true, completedTaskId: taskId, repeatedTask: repeatedTask || null };
}

/**
 * アプリから完了タスクを未完了に戻す処理
 */
function uncompleteTaskFromApp(taskId, taskData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 完了タスクシートから優先検索
  const doneSheet = ss.getSheetByName(CONFIG.DONE_SHEET_NAME);
  if (doneSheet && taskId) {
    const lastRow = doneSheet.getLastRow();
    if (lastRow >= 2) {
      const maxSearchRows = Math.min(lastRow - 1, 500);
      const ids = doneSheet.getRange(2, CONFIG.COL_ID, maxSearchRows, 1).getValues();
      const targetIdStr = String(taskId).trim();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === targetIdStr) {
          const foundRow = i + 2;
          const lastCol = Math.max(doneSheet.getLastColumn(), CONFIG.COL_END_TIME);
          const rowData = doneSheet.getRange(foundRow, 1, 1, lastCol).getValues()[0];
          doneSheet.deleteRow(foundRow);

          const restoredTask = taskData || parseRowToTask(rowData);
          restoredTask.id = targetIdStr;
          restoredTask.status = false;
          return addNewTaskFromApp(restoredTask);
        }
      }
    }
  }

  if (taskData) {
    taskData.status = false;
    return upsertTaskFromApp(taskData);
  }

  return { ok: true, uncompletedTaskId: taskId };
}

/**
 * アプリから新規プロジェクトを追加
 */
function addNewProjectFromApp(projectData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.PROJECT_SHEET_NAME);
  if (!sheet) throw new Error("プロジェクトシートが見つかりません。");
  
  const lastRow = sheet.getLastRow();
  const newName = (projectData.name || "").trim();
  if (!newName) throw new Error("プロジェクト名が空です。");
  
  // 既存の同名プロジェクト重複チェック
  if (lastRow >= 2) {
    const existingNames = sheet.getRange(2, 2, lastRow - 1, 1).getValues().map(r => String(r[0]).trim());
    if (existingNames.includes(newName)) {
      return getInitialData(); // 既に存在する場合は二重追加せず最新データを返す
    }
  }

  let newId = 100;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(row => Number(row[0])).filter(val => !isNaN(val));
    if (ids.length > 0) {
      newId = Math.max(...ids) + 1;
    }
  }
  
  const targetRow = lastRow + 1;
  sheet.getRange(targetRow, 1, 1, 3).setValues([[newId, newName, projectData.type || "Biz"]]);
  cachedProjectTypes = null;
  
  return getInitialData();
}

/**
 * アプリからプロジェクト（フォルダ）名を更新
 */
function updateProjectFromApp(projectId, newName, newType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.PROJECT_SHEET_NAME);
  if (!sheet) throw new Error("プロジェクトシートが見つかりません。");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("プロジェクトが存在しません。");

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  let foundRow = -1;
  let oldName = "";

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(projectId).trim()) {
      foundRow = i + 2;
      oldName = String(values[i][1]).trim();
      break;
    }
  }

  // IDで見つからない場合は名前で検索
  if (foundRow === -1) {
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][1]).trim() === String(projectId).trim()) {
        foundRow = i + 2;
        oldName = String(values[i][1]).trim();
        break;
      }
    }
  }

  if (foundRow !== -1) {
    if (newName && String(newName).trim() !== "") {
      const trimmedNewName = String(newName).trim();
      sheet.getRange(foundRow, 2).setValue(trimmedNewName);
      if (oldName && oldName !== trimmedNewName) {
        syncProjectNameChange(ss, oldName, trimmedNewName);
      }
    }
    if (newType && (newType === "Biz" || newType === "Private")) {
      sheet.getRange(foundRow, 3).setValue(newType);
    }
  } else if (newName) {
    return addNewProjectFromApp({ id: projectId, name: newName, type: newType || "Biz" });
  }
  cachedProjectTypes = null;

  return getInitialData();
}

/**
 * アプリからプロジェクト（フォルダ）を削除
 */
function deleteProjectFromApp(projectId, projectName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.PROJECT_SHEET_NAME);
  if (!sheet) throw new Error("プロジェクトシートが見つかりません。");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return getInitialData();

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() === String(projectId).trim() || String(values[i][1]).trim() === String(projectName).trim()) {
      sheet.deleteRow(i + 2);
      break;
    }
  }
  cachedProjectTypes = null;

  return getInitialData();
}

/**
 * アプリからタスクを高速削除する（まずPrivate/Bizを走査し、未発見時のみ完了シートを走査。ソート済みシートからの行削除は並び順が保たれるため再ソート不要）
 */
function deleteTaskFromApp(taskId, taskName, fallbackDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let deletedCount = 0;

  if (taskId && !String(taskId).startsWith("tmp_")) {
    const targetIdStr = String(taskId).trim();

    // 1. まずアクティブな未完了シート（Private, Biz）から検索・削除
    const activeNames = [CONFIG.PRIVATE_SHEET_NAME, CONFIG.BIZ_SHEET_NAME];
    activeNames.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) return;
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;

      const ids = sheet.getRange(2, CONFIG.COL_ID, lastRow - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]).trim() === targetIdStr) {
          sheet.deleteRow(i + 2);
          deletedCount++;
        }
      }
    });

    // 2. アクティブシートになかった場合のみ、完了タスクシートを検索・削除
    if (deletedCount === 0) {
      const doneSheet = ss.getSheetByName(CONFIG.DONE_SHEET_NAME);
      if (doneSheet) {
        const lastRow = doneSheet.getLastRow();
        if (lastRow >= 2) {
          const maxSearchRows = Math.min(lastRow - 1, 500);
          const ids = doneSheet.getRange(2, CONFIG.COL_ID, maxSearchRows, 1).getValues();
          for (let i = ids.length - 1; i >= 0; i--) {
            if (String(ids[i][0]).trim() === targetIdStr) {
              doneSheet.deleteRow(i + 2);
              deletedCount++;
            }
          }
        }
      }
    }
  }

  // 3. IDで見つからなかった場合のみ、タスク名と期日でフォールバック検索して削除
  if (deletedCount === 0 && taskName) {
    const location = findTaskLocationById(ss, null, taskName, fallbackDate, true);
    if (location) {
      location.sheet.deleteRow(location.row);
      deletedCount++;
    }
  }

  return {
    ok: true,
    deletedTaskId: taskId,
    deleteSuccess: deletedCount > 0,
    deleteNotFound: deletedCount === 0
  };
}

/**
 * コンテキストシートを取得（存在しなければ自動作成・ヘッダー初期化）
 */
function getOrCreateContextSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.CONTEXT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.CONTEXT_SHEET_NAME);
    sheet.getRange(1, 1).setValue("ID");
    sheet.getRange(1, 2).setValue("コンテキスト名");
    sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#e2e8f0");
  }
  return sheet;
}

/**
 * アプリからコンテキストを追加または更新
 */
function saveContextFromApp(contextData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateContextSheet(ss);
  
  const lastRow = sheet.getLastRow();
  const cId = contextData.id ? String(contextData.id) : "";
  const cName = contextData.name ? String(contextData.name).trim() : "";
  
  if (!cName) throw new Error("コンテキスト名を入力してください");

  // 既存IDの行を検索して更新
  let foundRow = -1;
  if (cId && lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === cId) {
        foundRow = i + 2;
        break;
      }
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 2).setValue(cName);
  } else {
    // 新規採番
    let newId = 1;
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => Number(r[0])).filter(v => !isNaN(v));
      if (ids.length > 0) newId = Math.max(...ids) + 1;
    }
    const targetRow = lastRow + 1;
    sheet.getRange(targetRow, 1, 1, 2).setValues([[newId, cName]]);
  }

  return getInitialData();
}

/**
 * アプリからコンテキストを削除
 */
function deleteContextFromApp(contextId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.CONTEXT_SHEET_NAME);
  if (!sheet) return getInitialData();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return getInitialData();

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]).trim() === String(contextId).trim()) {
      sheet.deleteRow(i + 2);
      break;
    }
  }

  return getInitialData();
}

/**
 * Gmail送信権限を強制承認＆テスト送信する関数
 */
function forceAuthorizeGmail() {
  GmailApp.getAliases(); // try-catchなしで直接呼ぶことで承認ダイアログを強制表示
  testSendOutlook();
}

/**
 * Outlook送信の権限承認＆テスト用関数
 */
function testSendOutlook() {
  const now = new Date();
  const end = new Date(now.getTime() + 30 * 60 * 1000); // 30分後
  sendIcsToOutlook("【テスト】タスク完了通知", now, end, "Outlook連携のテスト送信です。");
}

/**
 * アプリからのドラッグ＆ドロップ並び順更新リクエストの受信関数
 */
function updateSortOrderFromApp(idList) {
  // 並び順更新を受け取り、最新のタスクデータを返却
  return getInitialData();
}
