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
      const lastCol = sheet.getLastColumn();
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
 * タスク完了時のメイン処理
 */
function processTaskCompletion(ss, sheet, row) {
  // A列から一番右の列まで全てのデータを取得する
  const lastCol = sheet.getLastColumn();
  const rowRange = sheet.getRange(row, 1, 1, lastCol);
  const rowData = rowRange.getValues()[0];

  const repeatType = rowData[CONFIG.COL_REPEAT - 1];
  const sheetName = sheet.getName();

  // カレンダー登録とリピート作成を分離し、一方のエラーで止まらないように安全に処理
  // --- カレンダー＆リピート作成 ---
  try {
    syncToCalendarLogic(rowData, sheetName);
  } catch (err) {
    console.error("カレンダー作成エラー: " + err.message);
  }

  try {
    const rtTypeStr = String(repeatType || "").trim();
    if (rtTypeStr !== "" && rtTypeStr.toUpperCase() !== "NONE") {
      handleRepeatTask(sheet, rowData, rtTypeStr);
    }
  } catch (err) {
    console.error("リピート作成エラー: " + err.message);
  }

  // --- 完了転記と削除 ---
  try {
    const doneSheet = ss.getSheetByName(CONFIG.DONE_SHEET_NAME) || ss.insertSheet(CONFIG.DONE_SHEET_NAME);

    // ヘッダーとARRAYFORMULAの初期確認・設定
    setupDoneSheetFormat(doneSheet);

    // 常に2行目の位置に新しい空行を自動挿入（最新の完了タスクを一番上に配置するため）
    doneSheet.insertRowBefore(2);

    // ARRAYFORMULAの展開を妨げないようE列(COL_PROJECT_NAME)の値を空にする
    const newDoneRow = [...rowData];
    newDoneRow[CONFIG.COL_PROJECT_NAME - 1] = "";
    newDoneRow[CONFIG.COL_STATUS - 1] = true;

    // 2行目に転記
    doneSheet.getRange(2, 1, 1, newDoneRow.length).setValues([newDoneRow]);

    // 書式・チェックボックスの設定
    doneSheet.getRange(2, CONFIG.COL_STATUS).insertCheckboxes();
    doneSheet.getRange(2, CONFIG.COL_STATUS).setValue(true);
    doneSheet.getRange(2, CONFIG.COL_DATE_E).setNumberFormat("M/d(ddd)");
    doneSheet.getRange(2, CONFIG.COL_START_TIME).setNumberFormat("hh:mm");
    doneSheet.getRange(2, CONFIG.COL_END_TIME).setNumberFormat("hh:mm");

    SpreadsheetApp.flush();
    sheet.deleteRow(row);

    ss.toast("タスクを完了し、完了タスクの一番上へ移動しました。", "処理完了");

    // 完了タスクが移動・削除された後、該当シートの並び替えを実行する
    sortAll(sheet);

  } catch (err) {
    ss.toast("転記/削除エラー: " + err.message, "エラー");
  }
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
      SpreadsheetApp.getActiveSpreadsheet().toast("【警告】期日が日付形式ではないか、入力されていないため、次回のタスクが作られませんでした。", "エラー");
      return;
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

  // ARRAYFORMULAの自動展開を妨げないように、追加された行のプロジェクト名(E列)のセルを空っぽに戻します
  sheet.getRange(targetRow, CONFIG.COL_PROJECT_NAME).clearContent();

  sheet.getRange(targetRow, CONFIG.COL_STATUS).insertCheckboxes();
  sheet.getRange(targetRow, CONFIG.COL_DATE_E).setNumberFormat("M/d(ddd)");
  sheet.getRange(targetRow, CONFIG.COL_START_TIME).setNumberFormat("hh:mm");
  sheet.getRange(targetRow, CONFIG.COL_END_TIME).setNumberFormat("hh:mm");
}

/**
 * ARRAYFORMULAなどで空白に見えるセルを除き、A列で本当に値が入っている最後の行を取得する関数
 */
function getRealLastRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 1;

  const values = sheet.getRange(1, 1, lastRow, 1).getValues(); // A列を取得
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
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
    SpreadsheetApp.getActiveSpreadsheet().toast("Outlookメールアドレス(OUTLOOK_EMAIL)が未設定のため予定メールは送信されませんでした。", "通知");
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
 * プロジェクト名の一括置換
 */
function syncProjectNameChange(ss, oldName, newName) {
  const sheetsToUpdate = [CONFIG.PRIVATE_SHEET_NAME, CONFIG.BIZ_SHEET_NAME, CONFIG.OLD_SHEET_NAME, CONFIG.DONE_SHEET_NAME];

  sheetsToUpdate.forEach(name => {
    const targetSheet = ss.getSheetByName(name);
    if (!targetSheet) return;

    const lastRow = targetSheet.getLastRow();
    if (lastRow < 2) return;

    // プロジェクト名が存在する列の一括置換
    const range = targetSheet.getRange(2, CONFIG.COL_PROJECT_NAME, lastRow - 1, 1);
    const values = range.getValues();
    const updatedValues = values.map(row => [row[0] === oldName ? newName : row[0]]);
    range.setValues(updatedValues);
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
 * 指定シート（またはアクティブシート）を「完了状態 → 期日 (昇順) → コンテキスト (ID/マスター定義順)」で並び替える関数
 */
function sortByDateAndContext(targetSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = targetSheet || ss.getActiveSheet();
  if (!sheet || !isTaskSheet(sheet.getName())) return;

  // コンテキストシートから優先度マップ（コンテキスト名 -> ID/順序）を作成
  const contextSheet = ss.getSheetByName(CONFIG.CONTEXT_SHEET_NAME);
  const contextOrderMap = new Map();
  if (contextSheet) {
    const cValues = contextSheet.getDataRange().getValues();
    for (let i = 1; i < cValues.length; i++) {
      const cId = Number(cValues[i][0]) || (i * 10); // A列: コンテキストID (100, 101...)
      const cName = String(cValues[i][1]).trim();     // B列: コンテキスト名
      if (cName) {
        contextOrderMap.set(cName, cId);
      }
    }
  }

  // 既存のフィルターを一時解除（安全に実行）
  try {
    const existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
  } catch (e) {
    console.warn("フィルター解除スキップ: " + e.message);
  }

  const lastRow = getRealLastRow(sheet);
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const values = range.getValues();

  // 多重ソート（1. 完了状態 -> 2. 日付 -> 3. 開始時間 -> 4. コンテキストID）
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

  range.setValues(values);

  // フィルターを再作成（安全に実行）
  try {
    const filterRange = sheet.getRange(1, 1, lastRow, lastCol);
    filterRange.createFilter();
  } catch (e) {
    console.warn("フィルター再作成スキップ: " + e.message);
  }

  try {
    ss.toast(`「${sheet.getName()}」シートを並び替えました。`, "処理完了");
  } catch (e) {
    // Web APIからの実行時はtoastが使用できないため無視
  }
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
  [CONFIG.PRIVATE_SHEET_NAME, CONFIG.BIZ_SHEET_NAME, CONFIG.OLD_SHEET_NAME].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) sortByDateAndContext(s);
  });
}

/**
 * 外部Webアプリ（GitHub Pages等）からのAPI処理共通ディスパッチャー
 */
function handleApiAction(action, payload) {
  payload = payload || {};
  let result;

  if (action === 'getInitialData') {
    result = getInitialData();
  } else if (action === 'saveTask') {
    const data = payload.taskData;
    const isTempId = data && data.id && String(data.id).startsWith('tmp_');
    if (data && data.id && !isTempId) {
      result = updateTaskFromApp(data);
    } else {
      result = addNewTaskFromApp(data);
    }
  } else if (action === 'toggleTaskStatus' || action === 'completeTask') {
    result = completeTaskFromApp(payload.taskId);
  } else if (action === 'deleteTask') {
    result = deleteTaskFromApp(payload.taskId, payload.taskName, payload.taskDate);
  } else if (action === 'saveProject') {
    result = addNewProjectFromApp(payload.projectData);
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
  if (projectSheet) {
    const values = projectSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const pId = String(values[i][0] || "").trim();
      const pName = String(values[i][1] || "").trim();
      if (pId !== "" && pName !== "") {
        projects.push({
          id: pId,
          name: pName,
          type: String(values[i][2] || "Private")
        });
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

  // 2. タスクデータ取得 (Private, Biz, 旧タスクシートから未完了タスクを取得)
  const tasks = [];
  const taskSheetNames = [CONFIG.PRIVATE_SHEET_NAME, CONFIG.BIZ_SHEET_NAME, CONFIG.OLD_SHEET_NAME];
  
  taskSheetNames.forEach(name => {
    const taskSheet = ss.getSheetByName(name);
    if (taskSheet) {
      const values = taskSheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][CONFIG.COL_ID - 1] !== "") {
          const taskObj = parseRowToTask(values[i]);
          // コンテキストIDが入っている場合は名前に解決
          if (taskObj.description && idToContextName[taskObj.description]) {
            taskObj.description = idToContextName[taskObj.description];
          }
          // シート名に基づくデフォルト補正（未指定の場合）
          if (name === CONFIG.BIZ_SHEET_NAME && !taskObj.type) {
            taskObj.type = "Biz";
          }
          tasks.push(taskObj);
        }
      }
    }
  });
  
  // 3. 完了タスクデータ取得 (直近50件)
  const doneSheet = ss.getSheetByName(CONFIG.DONE_SHEET_NAME);
  const doneTasks = [];
  if (doneSheet) {
    const values = doneSheet.getDataRange().getValues();
    const startRow = Math.max(1, values.length - 50);
    for (let i = values.length - 1; i >= startRow; i--) {
      if (values[i][CONFIG.COL_ID - 1] !== "") {
        const taskObj = parseRowToTask(values[i]);
        if (taskObj.description && idToContextName[taskObj.description]) {
          taskObj.description = idToContextName[taskObj.description];
        }
        doneTasks.push(taskObj);
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
    status: String(row[CONFIG.COL_STATUS - 1]).toUpperCase() === "TRUE",
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
 * 日付オブジェクトを yyyy-MM-dd フォーマットの文字列に変換
 */
function formatDate(dateVal) {
  if (dateVal instanceof Date) {
    if (isNaN(dateVal.getTime())) return "";
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(dateVal || "");
}

/**
 * 時間オブジェクトを HH:mm フォーマットの文字列に変換
 */
function formatTime(timeVal) {
  if (timeVal instanceof Date) {
    if (isNaN(timeVal.getTime())) return "";
    return Utilities.formatDate(timeVal, Session.getScriptTimeZone(), "HH:mm");
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
 * アプリから新規タスクを追加（Private/Bizシートへ自動振り分け）
 */
function addNewTaskFromApp(taskData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskType = taskData.type || getProjectTypeById(ss, taskData.projectId);
  const sheet = getOrCreateTaskSheetByType(ss, taskType);
  
  // 既存のIDがあれば引き継ぎ、なければ新規発行
  const newId = taskData.id ? String(taskData.id) : Utilities.getUuid().split('-')[0];
  const row = getRealLastRow(sheet) + 1;
  
  const dateVal = taskData.date ? new Date(taskData.date.replace(/-/g, "/")) : "";
  const startTimeVal = taskData.startTime ? parseTimeStringToDate(taskData.date, taskData.startTime) : "";
  const endTimeVal = taskData.endTime ? parseTimeStringToDate(taskData.date, taskData.endTime) : "";
  
  const newRow = Array(CONFIG.COL_END_TIME).fill("");
  newRow[CONFIG.COL_ID - 1] = newId;
  newRow[CONFIG.COL_STATUS - 1] = "FALSE";
  newRow[CONFIG.COL_REPEAT - 1] = taskData.repeat || "None";
  newRow[CONFIG.COL_TYPE - 1] = taskType;
  newRow[CONFIG.COL_PROJECT_NAME - 1] = ""; // ARRAYFORMULA用（空文字）
  newRow[CONFIG.COL_PROJECT - 1] = taskData.projectId || "";
  newRow[CONFIG.COL_TASK - 1] = taskData.task || "";
  newRow[CONFIG.COL_DATE_E - 1] = dateVal;
  newRow[CONFIG.COL_START_TIME - 1] = startTimeVal;
  newRow[CONFIG.COL_END_TIME - 1] = endTimeVal;
  newRow[CONFIG.COL_DESCRIPTION - 1] = taskData.description || "";
  
  sheet.getRange(row, 1, 1, newRow.length).setValues([newRow]);
  
  sheet.getRange(row, CONFIG.COL_DATE_E).setNumberFormat("M/d(ddd)");
  sheet.getRange(row, CONFIG.COL_START_TIME).setNumberFormat("hh:mm");
  sheet.getRange(row, CONFIG.COL_END_TIME).setNumberFormat("hh:mm");
  
  SpreadsheetApp.flush();
  sortAll(sheet);
  
  return getInitialData();
}

/**
 * 日付文字列と時間文字列からDateオブジェクトを生成
 */
function parseTimeStringToDate(dateString, timeString) {
  if (!dateString || !timeString) return "";
  const d = new Date(dateString.replace(/-/g, "/"));
  const parts = timeString.split(":");
  d.setHours(parseInt(parts[0], 10), parseInt(parts[1], 10), 0, 0);
  return d;
}

/**
 * タスクIDから該当シートと行番号を検索（見つからない場合はタスク名でフォールバック検索）
 */
function findTaskLocationById(ss, taskId, fallbackTaskName, fallbackDate, includeDoneSheet) {
  const candidateNames = [CONFIG.PRIVATE_SHEET_NAME, CONFIG.BIZ_SHEET_NAME, CONFIG.OLD_SHEET_NAME];
  if (includeDoneSheet) {
    candidateNames.push(CONFIG.DONE_SHEET_NAME);
  }
  
  // 1. まず taskId で完全一致検索
  if (taskId) {
    for (let name of candidateNames) {
      const sheet = ss.getSheetByName(name);
      if (!sheet) continue;

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) continue;

      const ids = sheet.getRange(2, CONFIG.COL_ID, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === String(taskId).trim()) {
          return { sheet: sheet, row: i + 2 };
        }
      }
    }
  }

  // 2. taskId で見つからなかった場合、タスク名（および期日）でフォールバック検索（A列空欄行等の救済）
  if (fallbackTaskName) {
    const targetName = String(fallbackTaskName).trim();
    for (let name of candidateNames) {
      const sheet = ss.getSheetByName(name);
      if (!sheet) continue;

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) continue;

      const tasks = sheet.getRange(2, CONFIG.COL_TASK, lastRow - 1, 1).getValues();
      for (let i = 0; i < tasks.length; i++) {
        if (String(tasks[i][0]).trim() === targetName) {
          const row = i + 2;
          // ID列が空なら自動的にIDを採番して書き込む
          const currentId = sheet.getRange(row, CONFIG.COL_ID).getValue();
          if (!currentId) {
            sheet.getRange(row, CONFIG.COL_ID).setValue(taskId || Utilities.getUuid().split('-')[0]);
          }
          return { sheet: sheet, row: row };
        }
      }
    }
  }

  return null;
}

/**
 * アプリからタスクを編集
 */
function updateTaskFromApp(taskData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const location = findTaskLocationById(ss, taskData.id, taskData.task, taskData.date);
  
  if (!location) throw new Error("指定されたタスクが見つかりません。ID: " + taskData.id);
  
  const { sheet, row } = location;
  const newType = taskData.type || getProjectTypeById(ss, taskData.projectId);
  const isBizType = String(newType).toUpperCase() === "BIZ";
  const expectedSheetName = isBizType ? CONFIG.BIZ_SHEET_NAME : CONFIG.PRIVATE_SHEET_NAME;

  // もしタスクの種類（Private/Biz）が変更されて別シートへ移動が必要な場合
  if (sheet.getName() !== expectedSheetName && (sheet.getName() === CONFIG.PRIVATE_SHEET_NAME || sheet.getName() === CONFIG.BIZ_SHEET_NAME || sheet.getName() === CONFIG.OLD_SHEET_NAME)) {
    // 既存シートから削除し、新しいシートへ追加（IDは保持）
    sheet.deleteRow(row);
    addNewTaskFromApp(taskData);
    return getInitialData();
  }

  const dateVal = taskData.date ? new Date(taskData.date.replace(/-/g, "/")) : "";
  const startTimeVal = taskData.startTime ? parseTimeStringToDate(taskData.date, taskData.startTime) : "";
  const endTimeVal = taskData.endTime ? parseTimeStringToDate(taskData.date, taskData.endTime) : "";
  
  sheet.getRange(row, CONFIG.COL_REPEAT).setValue(taskData.repeat || "None");
  sheet.getRange(row, CONFIG.COL_TYPE).setValue(newType);
  
  // ARRAYFORMULA数式との衝突防止のため、E列(プロジェクト名)のセルをクリア
  sheet.getRange(row, CONFIG.COL_PROJECT_NAME).clearContent();
  
  sheet.getRange(row, CONFIG.COL_PROJECT).setValue(taskData.projectId || "");
  sheet.getRange(row, CONFIG.COL_TASK).setValue(taskData.task || "");
  sheet.getRange(row, CONFIG.COL_DATE_E).setValue(dateVal);
  sheet.getRange(row, CONFIG.COL_START_TIME).setValue(startTimeVal);
  sheet.getRange(row, CONFIG.COL_END_TIME).setValue(endTimeVal);
  sheet.getRange(row, CONFIG.COL_DESCRIPTION).setValue(taskData.description || "");
  
  sheet.getRange(row, CONFIG.COL_DATE_E).setNumberFormat("M/d(ddd)");
  sheet.getRange(row, CONFIG.COL_START_TIME).setNumberFormat("hh:mm");
  sheet.getRange(row, CONFIG.COL_END_TIME).setNumberFormat("hh:mm");
  
  SpreadsheetApp.flush();
  sortAll(sheet);
  
  return getInitialData();
}

/**
 * アプリからタスクを完了処理
 */
function completeTaskFromApp(taskId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const location = findTaskLocationById(ss, taskId);
  
  if (!location) throw new Error("指定されたタスクが見つかりません。ID: " + taskId);
  
  const { sheet, row } = location;
  sheet.getRange(row, CONFIG.COL_STATUS).setValue(true);
  processTaskCompletion(ss, sheet, row);
  
  return getInitialData();
}

/**
 * アプリから新規プロジェクトを追加
 */
function addNewProjectFromApp(projectData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.PROJECT_SHEET_NAME);
  if (!sheet) throw new Error("プロジェクトシートが見つかりません。");
  
  const lastRow = sheet.getLastRow();
  let newId = 100;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(row => Number(row[0])).filter(val => !isNaN(val));
    if (ids.length > 0) {
      newId = Math.max(...ids) + 1;
    }
  }
  
  const targetRow = lastRow + 1;
  sheet.getRange(targetRow, 1).setValue(newId);
  sheet.getRange(targetRow, 2).setValue(projectData.name);
  sheet.getRange(targetRow, 3).setValue(projectData.type || "Private");
  
  SpreadsheetApp.flush();
  return getInitialData();
}

/**
 * アプリからタスクを削除する（カレンダー同期やリピートは行わず、行を削除）
 */
function deleteTaskFromApp(taskId, taskName, fallbackDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 完了シートも含めて検索
  const location = findTaskLocationById(ss, taskId, taskName, fallbackDate, true);
  
  if (!location) {
    // 一時ID（tmp_...）または見つからない場合でも、すでに削除されたとみなしてエラーにせず同期
    console.warn("指定されたタスクが見つかりませんでした（すでに削除済みの可能性があります）。ID: " + taskId);
    return { status: 'success', deletedId: taskId };
  }
  
  const { sheet, row } = location;
  sheet.deleteRow(row);
  SpreadsheetApp.flush();
  sortAll(sheet);
  
  return { status: 'success', deletedId: taskId };
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
    sheet.getRange(targetRow, 1).setValue(newId);
    sheet.getRange(targetRow, 2).setValue(cName);
  }

  SpreadsheetApp.flush();
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
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(contextId).trim()) {
      sheet.deleteRow(i + 2);
      break;
    }
  }

  SpreadsheetApp.flush();
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

