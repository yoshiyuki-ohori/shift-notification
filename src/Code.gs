/**
 * Code.gs - エントリポイント・カスタムメニュー
 * シフト通知システムのメイン制御
 */

/**
 * スプレッドシート起動時にカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('シフト通知')
    .addItem('シフトデータ取込 (練馬)', 'importNerimaShifts')
    .addItem('シフトデータ取込 (世田谷)', 'importSetagayaShifts')
    .addSeparator()
    .addItem('名寄せ実行', 'runNameMatching')
    .addItem('未マッチ確認', 'showUnmatchedReport')
    .addSeparator()
    .addItem('テスト送信 (管理者のみ)', 'sendTestNotification')
    .addItem('一括送信 (本番)', 'sendAllNotifications')
    .addSeparator()
    .addItem('送信ログ確認', 'showSendLog')
    .addToUi();
}

/**
 * 練馬エリアのシフトデータ取込
 */
function importNerimaShifts() {
  const ui = SpreadsheetApp.getUi();
  try {
    const targetMonth = getSettingValue(SETTING_KEYS.TARGET_MONTH);
    const ssId = getSettingValue(SETTING_KEYS.NERIMA_SS_ID);

    ui.alert('取込開始', '練馬エリアのシフトデータを取り込みます。\n対象年月: ' + targetMonth, ui.ButtonSet.OK);

    const sourceSS = SpreadsheetApp.openById(ssId);
    const sheets = sourceSS.getSheets();
    let totalRecords = 0;

    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      // シフトシートを判定（施設名を含むシートのみ処理）
      const records = parseNerimaShift(sheet, targetMonth);
      if (records.length > 0) {
        writeShiftData(records);
        totalRecords += records.length;
      }
    }

    ui.alert('取込完了', '練馬エリア: ' + totalRecords + '件のシフトデータを取り込みました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', '練馬シフト取込でエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
    Logger.log('importNerimaShifts error: ' + e.toString());
  }
}

/**
 * 世田谷エリアのシフトデータ取込
 */
function importSetagayaShifts() {
  const ui = SpreadsheetApp.getUi();
  try {
    const targetMonth = getSettingValue(SETTING_KEYS.TARGET_MONTH);
    const ssId = getSettingValue(SETTING_KEYS.SETAGAYA_SS_ID);

    ui.alert('取込開始', '世田谷エリアのシフトデータを取り込みます。\n対象年月: ' + targetMonth, ui.ButtonSet.OK);

    const sourceSS = SpreadsheetApp.openById(ssId);
    const sheets = sourceSS.getSheets();
    let totalRecords = 0;

    for (const sheet of sheets) {
      const records = parseSetagayaShift(sheet, targetMonth);
      if (records.length > 0) {
        writeShiftData(records);
        totalRecords += records.length;
      }
    }

    ui.alert('取込完了', '世田谷エリア: ' + totalRecords + '件のシフトデータを取り込みました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', '世田谷シフト取込でエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
    Logger.log('importSetagayaShifts error: ' + e.toString());
  }
}

/**
 * 名寄せ実行
 */
function runNameMatching() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shiftSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_DATA);
    if (!shiftSheet || shiftSheet.getLastRow() <= 1) {
      ui.alert('エラー', 'シフトデータがありません。先にシフトデータを取り込んでください。', ui.ButtonSet.OK);
      return;
    }

    const employeeMaster = loadEmployeeMaster();
    const matcher = new NameMatcherEngine(employeeMaster);

    const shiftData = shiftSheet.getDataRange().getValues();
    let matchedCount = 0;
    let unmatchedCount = 0;
    const unmatchedNames = [];

    // ヘッダー行をスキップ (i=1から)
    for (let i = 1; i < shiftData.length; i++) {
      const originalName = String(shiftData[i][SHIFT_COLS.ORIGINAL_NAME - 1]).trim();
      if (!originalName) continue;

      // 既にマッチ済みの場合はスキップ
      if (shiftData[i][SHIFT_COLS.EMPLOYEE_NO - 1]) continue;

      const facility = String(shiftData[i][SHIFT_COLS.FACILITY - 1]).trim();
      const result = matcher.match(originalName, facility);

      if (result) {
        shiftSheet.getRange(i + 1, SHIFT_COLS.EMPLOYEE_NO).setValue(result.employeeNo);
        shiftSheet.getRange(i + 1, SHIFT_COLS.FORMAL_NAME).setValue(result.formalName);
        matchedCount++;
      } else {
        const candidates = matcher.findCandidates(originalName);
        unmatchedNames.push({
          facility: facility,
          originalName: originalName,
          candidates: candidates.join(', ')
        });
        unmatchedCount++;
      }
    }

    // 未マッチデータ書き込み
    if (unmatchedNames.length > 0) {
      writeUnmatchedData(unmatchedNames);
    }

    ui.alert('名寄せ完了',
      'マッチ成功: ' + matchedCount + '件\n' +
      '未マッチ: ' + unmatchedCount + '件\n\n' +
      (unmatchedCount > 0 ? '「未マッチ確認」から未マッチ一覧を確認してください。' : ''),
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', '名寄せでエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
    Logger.log('runNameMatching error: ' + e.toString());
  }
}

/**
 * 未マッチレポート表示
 */
function showUnmatchedReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.UNMATCHED);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('情報', '未マッチデータはありません。', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * テスト送信（管理者のみ）
 */
function sendTestNotification() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('テスト送信',
    'テスト送信を実行します。管理者のLINEにのみ送信されます。\n実行しますか？',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  try {
    setSettingValue(SETTING_KEYS.SEND_MODE, 'テスト');
    startBatchSend();
    ui.alert('テスト送信完了', 'テスト送信が完了しました。送信ログを確認してください。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', 'テスト送信でエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
    Logger.log('sendTestNotification error: ' + e.toString());
  }
}

/**
 * 一括送信（本番）
 */
function sendAllNotifications() {
  const ui = SpreadsheetApp.getUi();
  const targetMonth = getSettingValue(SETTING_KEYS.TARGET_MONTH);

  const response = ui.alert('一括送信確認',
    '【本番送信】\n対象年月: ' + targetMonth + '\n\n' +
    '全職員にシフト通知を送信します。\n本当に実行しますか？',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  try {
    setSettingValue(SETTING_KEYS.SEND_MODE, '本番');
    startBatchSend();
    ui.alert('送信開始', '一括送信を開始しました。\n送信ログで進捗を確認してください。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', '一括送信でエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
    Logger.log('sendAllNotifications error: ' + e.toString());
  }
}

/**
 * 送信ログ表示
 */
function showSendLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SEND_LOG);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('情報', '送信ログはまだありません。', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * シフトデータをシートに書き込む
 * @param {Array<Object>} records - シフトレコード配列
 */
function writeShiftData(records) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_DATA);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.SHIFT_DATA);
    sheet.getRange(1, 1, 1, 8).setValues([[
      '年月', '日付', 'エリア', '施設名', '時間帯', '担当者名(原文)', '社員No', '氏名(正式)'
    ]]);
  }

  const rows = records.map(r => [
    r.yearMonth,
    r.date,
    r.area,
    r.facility,
    r.timeSlot,
    r.originalName,
    r.employeeNo || '',
    r.formalName || ''
  ]);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 8).setValues(rows);
}

/**
 * 未マッチデータをシートに書き込む
 * @param {Array<Object>} unmatchedList - 未マッチリスト
 */
function writeUnmatchedData(unmatchedList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.UNMATCHED);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.UNMATCHED);
    sheet.getRange(1, 1, 1, 4).setValues([[
      '施設名', '原文名前', '候補', '解決済'
    ]]);
  }

  // 既存の未マッチから重複除去
  const existing = new Set();
  if (sheet.getLastRow() > 1) {
    const existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    existingData.forEach(row => existing.add(row[0] + '|' + row[1]));
  }

  const newRows = unmatchedList
    .filter(item => !existing.has(item.facility + '|' + item.originalName))
    .map(item => [item.facility, item.originalName, item.candidates, 'FALSE']);

  if (newRows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, 4).setValues(newRows);
  }
}

/**
 * 従業員マスタを読み込む
 * @return {Array<Object>} 従業員データ配列
 */
function loadEmployeeMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEE_MASTER);
  if (!sheet) throw new Error('従業員マスタシートが見つかりません');

  const data = sheet.getDataRange().getValues();
  const employees = [];

  // ヘッダー行スキップ (i=1から)
  for (let i = 1; i < data.length; i++) {
    const no = String(data[i][MASTER_COLS.NO - 1]).trim();
    const name = String(data[i][MASTER_COLS.NAME - 1]).trim();
    if (!no || !name) continue;

    employees.push({
      employeeNo: no,
      name: name,
      furigana: String(data[i][MASTER_COLS.FURIGANA - 1] || '').trim(),
      lineUserId: String(data[i][MASTER_COLS.LINE_USER_ID - 1] || '').trim(),
      area: String(data[i][MASTER_COLS.AREA - 1] || '').trim(),
      facility: String(data[i][MASTER_COLS.FACILITY - 1] || '').trim(),
      status: String(data[i][MASTER_COLS.STATUS - 1] || '在職').trim(),
      notifyEnabled: data[i][MASTER_COLS.NOTIFY - 1] !== false && data[i][MASTER_COLS.NOTIFY - 1] !== 'FALSE',
      aliases: String(data[i][MASTER_COLS.ALIASES - 1] || '').split(',').map(a => a.trim()).filter(a => a)
    });
  }

  return employees;
}

/**
 * 初期セットアップ: 必要なシートを作成
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 従業員マスタ
  if (!ss.getSheetByName(SHEET_NAMES.EMPLOYEE_MASTER)) {
    const sheet = ss.insertSheet(SHEET_NAMES.EMPLOYEE_MASTER);
    sheet.getRange(1, 1, 1, 9).setValues([[
      'No', '氏名', 'フリガナ', 'LINE_UserId', 'エリア', '主担当施設', 'ステータス', '通知有効', '別名リスト'
    ]]);
    sheet.setFrozenRows(1);
  }

  // シフトデータ
  if (!ss.getSheetByName(SHEET_NAMES.SHIFT_DATA)) {
    const sheet = ss.insertSheet(SHEET_NAMES.SHIFT_DATA);
    sheet.getRange(1, 1, 1, 8).setValues([[
      '年月', '日付', 'エリア', '施設名', '時間帯', '担当者名(原文)', '社員No', '氏名(正式)'
    ]]);
    sheet.setFrozenRows(1);
  }

  // 送信ログ
  if (!ss.getSheetByName(SHEET_NAMES.SEND_LOG)) {
    const sheet = ss.insertSheet(SHEET_NAMES.SEND_LOG);
    sheet.getRange(1, 1, 1, 6).setValues([[
      '送信日時', '対象年月', '社員No', '氏名', 'ステータス', 'エラー詳細'
    ]]);
    sheet.setFrozenRows(1);
  }

  // 設定
  if (!ss.getSheetByName(SHEET_NAMES.SETTINGS)) {
    const sheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    sheet.getRange(1, 1, 1, 3).setValues([['キー', '値', '説明']]);
    sheet.getRange(2, 1, 5, 3).setValues([
      [SETTING_KEYS.TARGET_MONTH, '', '送信対象月 (例: 2026-03)'],
      [SETTING_KEYS.NERIMA_SS_ID, '', '練馬シフトのSpreadsheet ID'],
      [SETTING_KEYS.SETAGAYA_SS_ID, '', '世田谷シフトのSpreadsheet ID'],
      [SETTING_KEYS.SEND_MODE, 'テスト', 'テスト/本番'],
      [SETTING_KEYS.TEST_USER_ID, '', 'テスト用LINE UserId']
    ]);
    sheet.setFrozenRows(1);
  }

  // 名寄せ未マッチ
  if (!ss.getSheetByName(SHEET_NAMES.UNMATCHED)) {
    const sheet = ss.insertSheet(SHEET_NAMES.UNMATCHED);
    sheet.getRange(1, 1, 1, 4).setValues([[
      '施設名', '原文名前', '候補', '解決済'
    ]]);
    sheet.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert('セットアップ完了', '全シートの初期設定が完了しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}
