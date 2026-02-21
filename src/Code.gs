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
    .addItem('シフトデータ取込 (マスタExcel)', 'importMasterExcelShifts')
    .addSeparator()
    .addItem('名寄せ実行', 'runNameMatching')
    .addItem('未マッチ確認', 'showUnmatchedReport')
    .addSeparator()
    .addItem('テスト送信 (管理者のみ)', 'sendTestNotification')
    .addItem('一括送信 (本番)', 'sendAllNotifications')
    .addSeparator()
    .addItem('送信ログ確認', 'showSendLog')
    .addSeparator()
    .addItem('Googleカレンダー一括登録', 'syncAllToGoogleCalendar')
    .addSeparator()
    .addItem('一斉メッセージ送信 (全友だち)', 'sendBroadcastMessage')
    .addToUi();

  ui.createMenu('社員マスタ')
    .addItem('外部マスタから同期 (手動)', 'runManualSync')
    .addItem('毎日自動同期トリガー設定', 'setupDailySyncTrigger')
    .addToUi();

  ui.createMenu('シフト希望')
    .addItem('希望収集を開始', 'startPreferenceCollection')
    .addItem('希望収集を締切', 'closePreferenceCollection')
    .addSeparator()
    .addItem('提出状況確認', 'showSubmissionStatus')
    .addItem('未提出者にリマインド送信', 'sendPreferenceReminders')
    .addSeparator()
    .addItem('希望一覧シート表示', 'showPreferenceSheet')
    .addToUi();

  ui.createMenu('労基チェック')
    .addItem('労基チェック実行', 'runComplianceCheckMenu')
    .addItem('配置充足率チェック', 'runStaffingCheckMenu')
    .addSeparator()
    .addItem('労基チェック結果表示', 'showComplianceResults')
    .addItem('労基ルール設定表示', 'showLaborRulesSheet')
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

    // 自動コンプライアンスチェック
    runAutoComplianceCheck_(targetMonth);
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

    // 自動コンプライアンスチェック
    runAutoComplianceCheck_(targetMonth);
  } catch (e) {
    ui.alert('エラー', '世田谷シフト取込でエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
    Logger.log('importSetagayaShifts error: ' + e.toString());
  }
}

/**
 * マスタExcelからのシフトデータ取込
 * Node.jsツール (parse-master-shift.js --write-ss) でデータ投入済みの場合に使用
 */
function importMasterExcelShifts() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_DATA);
  const rowCount = sheet ? Math.max(sheet.getLastRow() - 1, 0) : 0;

  ui.alert('マスタExcel取込',
    'マスタExcelからのシフトデータ取込は、Node.jsツールで実行します。\n\n' +
    '手順:\n' +
    '1. ターミナルで以下を実行:\n' +
    '   node tools/parse-master-shift.js --write-ss\n\n' +
    '2. 実行後、「名寄せ実行」で未マッチを解決してください。\n\n' +
    '現在のシフトデータ: ' + rowCount + '件',
    ui.ButtonSet.OK);
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
    sheet.getRange(1, 1, 1, 9).setValues([[
      '年月', '日付', 'エリア', '施設名(正式)', '施設コード', '時間帯', '担当者名(原文)', '社員No', '氏名(正式)'
    ]]);
  }

  const rows = records.map(r => {
    const officialName = getOfficialFacilityName(r.facility);
    const facilityId = getFacilityId(r.facility) || '';
    return [
      r.yearMonth,
      r.date,
      r.area,
      officialName,
      facilityId,
      r.timeSlot,
      r.originalName,
      r.employeeNo || '',
      r.formalName || ''
    ];
  });

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 9).setValues(rows);
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
      status: String(data[i][MASTER_COLS.STATUS - 1] || '在職').trim(),
      notifyEnabled: data[i][MASTER_COLS.NOTIFY - 1] !== false && data[i][MASTER_COLS.NOTIFY - 1] !== 'FALSE',
      aliases: String(data[i][MASTER_COLS.ALIASES - 1] || '').split(',').map(a => a.trim()).filter(a => a)
    });
  }

  return employees;
}

/**
 * Googleカレンダーにシフトを一括登録
 */
function syncAllToGoogleCalendar() {
  const ui = SpreadsheetApp.getUi();
  const targetMonth = getSettingValue(SETTING_KEYS.TARGET_MONTH);

  const response = ui.alert('Googleカレンダー登録',
    '対象年月: ' + targetMonth + '\n\n' +
    '全員分のシフトをGoogleカレンダーに登録します。\n' +
    '「シフト予定」カレンダーが自動作成されます。\n\n' +
    '実行しますか？',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  try {
    const eventCount = syncToGoogleCalendar(targetMonth);
    ui.alert('登録完了',
      'Googleカレンダーに ' + eventCount + ' 件のシフトイベントを登録しました。\n\n' +
      'Googleカレンダーの「シフト予定」カレンダーを確認してください。',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', 'Googleカレンダー登録でエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
    Logger.log('syncAllToGoogleCalendar error: ' + e.toString());
  }
}

/**
 * 一斉メッセージ送信（全友だちにブロードキャスト）
 */
function sendBroadcastMessage() {
  const ui = SpreadsheetApp.getUi();

  // メッセージ入力ダイアログ
  const result = ui.prompt('一斉メッセージ送信',
    '全友だちに送信するメッセージを入力してください。\n\n' +
    '例: 社員番号登録のお願いメッセージ',
    ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const messageText = result.getResponseText().trim();
  if (!messageText) {
    ui.alert('エラー', 'メッセージが入力されていません。', ui.ButtonSet.OK);
    return;
  }

  // 確認ダイアログ
  const confirm = ui.alert('送信確認',
    '以下のメッセージを全友だちに送信します。\n\n' +
    '---\n' + messageText + '\n---\n\n' +
    '本当に送信しますか？',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  const sendResult = broadcastMessage([createTextMessage(messageText)]);

  if (sendResult.success) {
    ui.alert('送信完了', '一斉メッセージを送信しました。', ui.ButtonSet.OK);
  } else {
    ui.alert('送信エラー', 'エラー: ' + sendResult.error, ui.ButtonSet.OK);
  }
}

/**
 * 社員番号登録お願いメッセージを一斉送信
 */
function sendRegistrationBroadcast() {
  const ui = SpreadsheetApp.getUi();

  const message = 'シフト通知システムのご案内です。\n\n' +
    '今後、毎月のシフト予定をこのLINEでお届けします。\n\n' +
    '【登録方法】\n' +
    'このトーク画面に「社員番号」を数字で送信してください。\n' +
    '（例: 072）\n\n' +
    '※社員番号が分からない場合は管理者にお問い合わせください。';

  const confirm = ui.alert('社員番号登録お願い送信',
    '以下のメッセージを全友だちに送信します。\n\n---\n' + message + '\n---\n\n送信しますか？',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  const result = broadcastMessage([createTextMessage(message)]);

  if (result.success) {
    ui.alert('送信完了', '登録お願いメッセージを全友だちに送信しました。', ui.ButtonSet.OK);
  } else {
    ui.alert('送信エラー', 'エラー: ' + result.error, ui.ButtonSet.OK);
  }
}

/**
 * LINE連携確認メッセージを一斉送信 (トリガーから呼び出し用)
 */
function sendLineVerificationBroadcast() {
  var message = 'シフト通知システムのLINE連携確認です。\n\n' +
    'お手数ですが、社員番号を数字で送信してください。\n' +
    '（例: 072）\n\n' +
    '送信いただくとLINE連携状況を確認できます。\n' +
    '既に登録済みの方には「連携済み」と表示されます。\n\n' +
    '※社員番号が分からない場合は管理者にお問い合わせください。';

  var result = broadcastMessage([createTextMessage(message)]);

  if (result.success) {
    Logger.log('LINE連携確認メッセージを送信しました');
  } else {
    Logger.log('LINE連携確認メッセージ送信エラー: ' + result.error);
  }

  // 使い捨てトリガーを削除
  deleteScheduledBroadcastTriggers_();
}

/**
 * 指定時刻にブロードキャストを予約
 * @param {Date} scheduledTime - 送信予定時刻
 */
function scheduleBroadcast_(scheduledTime) {
  deleteScheduledBroadcastTriggers_();
  ScriptApp.newTrigger('sendLineVerificationBroadcast')
    .timeBased()
    .at(scheduledTime)
    .create();
}

/**
 * 予約ブロードキャストのトリガーを削除
 */
function deleteScheduledBroadcastTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendLineVerificationBroadcast') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * 初期セットアップ: 必要なシートを作成
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 従業員マスタ
  if (!ss.getSheetByName(SHEET_NAMES.EMPLOYEE_MASTER)) {
    const sheet = ss.insertSheet(SHEET_NAMES.EMPLOYEE_MASTER);
    sheet.getRange(1, 1, 1, 7).setValues([[
      'No', '氏名', 'フリガナ', 'LINE_UserId', 'ステータス', '通知有効', '別名リスト'
    ]]);
    sheet.setFrozenRows(1);
  }

  // シフトデータ
  if (!ss.getSheetByName(SHEET_NAMES.SHIFT_DATA)) {
    const sheet = ss.insertSheet(SHEET_NAMES.SHIFT_DATA);
    sheet.getRange(1, 1, 1, 9).setValues([[
      '年月', '日付', 'エリア', '施設名(正式)', '施設コード', '時間帯', '担当者名(原文)', '社員No', '氏名(正式)'
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

  // シフト希望
  if (!ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE)) {
    const sheet = ss.insertSheet(SHEET_NAMES.SHIFT_PREFERENCE);
    sheet.getRange(1, 1, 1, 8).setValues([[
      '対象年月', '社員No', '氏名', '日付', '種別', '時間帯', '理由', '提出日時'
    ]]);
    sheet.setFrozenRows(1);
  }

  // 必要配置
  if (!ss.getSheetByName(SHEET_NAMES.STAFFING_REQUIREMENT)) {
    const sheet = ss.insertSheet(SHEET_NAMES.STAFFING_REQUIREMENT);
    sheet.getRange(1, 1, 1, 5).setValues([[
      '施設ID', '時間帯', '曜日種別', '最低人数', '推奨人数'
    ]]);
    sheet.setFrozenRows(1);
  }

  // 労基ルール
  setupLaborRulesSheet();

  // 労基チェック結果
  setupComplianceResultSheet();

  SpreadsheetApp.getUi().alert('セットアップ完了', '全シートの初期設定が完了しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===== シフト希望管理 =====

/**
 * 希望収集を開始
 */
function startPreferenceCollection() {
  const ui = SpreadsheetApp.getUi();

  // 対象年月を入力
  const monthResult = ui.prompt('希望収集開始',
    '対象年月を入力してください (例: 2026-04)',
    ui.ButtonSet.OK_CANCEL);
  if (monthResult.getSelectedButton() !== ui.Button.OK) return;
  const targetMonth = monthResult.getResponseText().trim();

  if (!targetMonth.match(/^\d{4}-\d{2}$/)) {
    ui.alert('エラー', '年月の形式が正しくありません。例: 2026-04', ui.ButtonSet.OK);
    return;
  }

  // 締切日を入力
  const deadlineResult = ui.prompt('締切日設定',
    '希望提出の締切日を入力してください (例: 2026-03-20)',
    ui.ButtonSet.OK_CANCEL);
  if (deadlineResult.getSelectedButton() !== ui.Button.OK) return;
  const deadline = deadlineResult.getResponseText().trim();

  if (!deadline.match(/^\d{4}-\d{2}-\d{2}$/)) {
    ui.alert('エラー', '日付の形式が正しくありません。例: 2026-03-20', ui.ButtonSet.OK);
    return;
  }

  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  // 収集期間を設定
  setCollectionPeriod(targetMonth, today, deadline);

  // リマインドトリガー設定（締切2日前に自動リマインド）
  setupPreferenceReminderTrigger_(deadline);

  // 全職員に希望入力開始通知を送信
  var confirm = ui.alert('通知送信',
    targetMonth + 'のシフト希望収集を開始します。\n' +
    '締切: ' + deadline + '\n\n' +
    '全職員にLINE通知を送信しますか？',
    ui.ButtonSet.YES_NO);

  if (confirm === ui.Button.YES) {
    sendPreferenceStartNotification_(targetMonth, deadline);
  }

  ui.alert('開始完了',
    'シフト希望収集を開始しました。\n\n' +
    '対象月: ' + targetMonth + '\n' +
    '締切: ' + deadline + '\n' +
    'リマインド: 締切2日前に自動送信',
    ui.ButtonSet.OK);
}

/**
 * 希望収集を締切
 */
function closePreferenceCollection() {
  const ui = SpreadsheetApp.getUi();
  var period = getCollectionPeriod();

  if (!period) {
    ui.alert('情報', '収集期間が設定されていません。', ui.ButtonSet.OK);
    return;
  }

  var status = getSubmissionStatus(period.targetMonth);
  var confirm = ui.alert('収集締切',
    '対象月: ' + period.targetMonth + '\n' +
    '提出済み: ' + status.submitted.length + '人\n' +
    '未提出: ' + status.notSubmitted.length + '人\n\n' +
    '希望収集を締め切りますか？',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  // 締切日を今日に設定（実質クローズ）
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  setSettingValue(SETTING_KEYS.PREF_COLLECTION_END, today);

  // リマインドトリガー削除
  deletePreferenceReminderTriggers_();

  ui.alert('締切完了', 'シフト希望収集を締め切りました。', ui.ButtonSet.OK);
}

/**
 * 提出状況確認
 */
function showSubmissionStatus() {
  const ui = SpreadsheetApp.getUi();
  var period = getCollectionPeriod();

  if (!period) {
    ui.alert('情報', '収集期間が設定されていません。', ui.ButtonSet.OK);
    return;
  }

  var status = getSubmissionStatus(period.targetMonth);
  var submittedNames = status.submitted.map(function(s) {
    return s.name + ' (' + s.count + '件)';
  }).join('\n');

  var notSubmittedNames = status.notSubmitted.map(function(s) {
    return s.name + (s.lineUserId ? '' : ' [LINE未登録]');
  }).join('\n');

  ui.alert('提出状況 - ' + period.targetMonth,
    '【提出済み: ' + status.submitted.length + '人】\n' +
    (submittedNames || '(なし)') + '\n\n' +
    '【未提出: ' + status.notSubmitted.length + '人】\n' +
    (notSubmittedNames || '(なし)'),
    ui.ButtonSet.OK);
}

/**
 * 未提出者にリマインド送信
 */
function sendPreferenceReminders() {
  const ui = SpreadsheetApp.getUi();
  var period = getCollectionPeriod();

  if (!period) {
    ui.alert('情報', '収集期間が設定されていません。', ui.ButtonSet.OK);
    return;
  }

  var status = getSubmissionStatus(period.targetMonth);

  if (status.notSubmitted.length === 0) {
    ui.alert('情報', '全員提出済みです。', ui.ButtonSet.OK);
    return;
  }

  var confirm = ui.alert('リマインド送信',
    '未提出者 ' + status.notSubmitted.length + '人にリマインドを送信しますか？',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  var result = sendPreferenceReminder(period.targetMonth);

  ui.alert('送信完了',
    '送信成功: ' + result.sent + '人\n' +
    '送信失敗: ' + result.failed + '人',
    ui.ButtonSet.OK);
}

/**
 * 希望一覧シート表示
 */
function showPreferenceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('情報', 'シフト希望データはまだありません。', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 希望入力開始通知を全職員に送信
 */
function sendPreferenceStartNotification_(targetMonth, deadline) {
  var employees = loadEmployeeMaster();
  var parts = targetMonth.split('-');
  var displayMonth = parts[0] + '年' + parseInt(parts[1], 10) + '月';

  employees.forEach(function(emp) {
    if (emp.status !== '在職' || !emp.lineUserId) return;

    var message = {
      type: 'flex',
      altText: displayMonth + ' シフト希望入力のお知らせ',
      contents: {
        type: 'bubble',
        size: 'kilo',
        header: {
          type: 'box',
          layout: 'vertical',
          contents: [{
            type: 'text',
            text: 'シフト希望入力のお知らせ',
            weight: 'bold',
            size: 'md',
            color: '#FFFFFF'
          }],
          backgroundColor: '#1DB446',
          paddingAll: '12px'
        },
        body: {
          type: 'box',
          layout: 'vertical',
          contents: [
            {
              type: 'text',
              text: emp.name + 'さん',
              weight: 'bold',
              size: 'md'
            },
            {
              type: 'text',
              text: displayMonth + 'のシフト希望入力を受付中です。',
              wrap: true,
              size: 'sm',
              margin: 'md',
              color: '#555555'
            },
            {
              type: 'text',
              text: '締切: ' + deadline,
              weight: 'bold',
              size: 'sm',
              margin: 'md',
              color: '#E74C3C'
            }
          ],
          paddingAll: '15px'
        },
        footer: {
          type: 'box',
          layout: 'vertical',
          contents: [{
            type: 'button',
            action: {
              type: 'postback',
              label: 'シフト希望を入力する',
              data: 'action=pref_start&month=' + targetMonth,
              displayText: 'シフト希望入力'
            },
            style: 'primary',
            color: '#1DB446',
            height: 'sm'
          }],
          paddingAll: '12px'
        }
      }
    };

    pushMessage(emp.lineUserId, [message]);
    Utilities.sleep(LINE_API.RATE_LIMIT_DELAY_MS);
  });
}

/**
 * リマインドトリガーを設定（締切2日前）
 */
function setupPreferenceReminderTrigger_(deadline) {
  deletePreferenceReminderTriggers_();

  var deadlineDate = new Date(deadline + 'T10:00:00+09:00');
  var reminderDate = new Date(deadlineDate.getTime() - 2 * 24 * 60 * 60 * 1000);

  // 過去でなければトリガー設定
  if (reminderDate > new Date()) {
    ScriptApp.newTrigger('autoSendPreferenceReminder')
      .timeBased()
      .at(reminderDate)
      .create();
  }
}

/**
 * リマインドトリガー削除
 */
function deletePreferenceReminderTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoSendPreferenceReminder') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * 自動リマインド実行（トリガーから呼び出し）
 */
function autoSendPreferenceReminder() {
  var period = getCollectionPeriod();
  if (!period) return;

  var result = sendPreferenceReminder(period.targetMonth);
  Logger.log('Auto reminder sent: ' + result.sent + ' success, ' + result.failed + ' failed');

  deletePreferenceReminderTriggers_();
}

// ===== 労基チェック関連 =====

/**
 * メニューから労基チェックを実行
 */
function runComplianceCheckMenu() {
  var ui = SpreadsheetApp.getUi();
  try {
    var targetMonth = getSettingValue(SETTING_KEYS.TARGET_MONTH);
    var confirm = ui.alert('労基チェック実行',
      '対象年月: ' + targetMonth + '\n\n' +
      '労働基準法に基づくコンプライアンスチェックを実行します。\n' +
      '結果は「労基チェック結果」シートに出力されます。',
      ui.ButtonSet.OK_CANCEL);

    if (confirm !== ui.Button.OK) return;

    var result = runComplianceCheck(targetMonth);
    writeComplianceResults(targetMonth, result);

    var msg = '【チェック完了】\n' +
      '対象: ' + result.summary.checkedEmployees + '名\n' +
      '違反: ' + result.summary.totalViolations + '件\n' +
      '警告: ' + result.summary.totalWarnings + '件\n' +
      '問題のある職員: ' + result.summary.employeesWithIssues + '名';

    if (result.summary.totalViolations > 0) {
      msg += '\n\n【重大違反あり】\n';
      for (var i = 0; i < Math.min(result.violations.length, 5); i++) {
        var v = result.violations[i];
        msg += '- ' + v.employeeName + ': ' + v.detail + '\n';
      }
      if (result.violations.length > 5) {
        msg += '...他' + (result.violations.length - 5) + '件';
      }
    }

    ui.alert('労基チェック結果', msg, ui.ButtonSet.OK);

    // 結果シートを表示
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.COMPLIANCE_RESULT);
    if (sheet) ss.setActiveSheet(sheet);
  } catch (e) {
    ui.alert('エラー', '労基チェックでエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
    Logger.log('runComplianceCheckMenu error: ' + e.toString());
  }
}

/**
 * メニューから配置充足率チェックを実行
 */
function runStaffingCheckMenu() {
  var ui = SpreadsheetApp.getUi();
  try {
    var targetMonth = getSettingValue(SETTING_KEYS.TARGET_MONTH);
    var result = checkStaffingLevels(targetMonth);

    var msg = '【配置充足率チェック結果】\n対象年月: ' + targetMonth + '\n\n';

    // エリア別サマリー
    var areas = Object.keys(result.areaSummary);
    for (var i = 0; i < areas.length; i++) {
      var area = areas[i];
      var s = result.areaSummary[area];
      msg += '■ ' + area + 'エリア\n';
      msg += '  充足率: ' + s.avgFillRate + '%\n';
      msg += '  不足: ' + s.shortage + '件 / 注意: ' + s.caution + '件 / OK: ' + s.ok + '件\n\n';
    }

    if (result.alerts.length > 0) {
      msg += '【要対応（人員不足）】\n';
      for (var j = 0; j < Math.min(result.alerts.length, 10); j++) {
        var a = result.alerts[j];
        msg += '- ' + a.facilityName + ' ' + formatShortDate_(a.date) + ' ' + a.timeSlot +
               ' (' + a.assigned + '/' + a.required + '名)\n';
      }
      if (result.alerts.length > 10) {
        msg += '...他' + (result.alerts.length - 10) + '件';
      }
    }

    ui.alert('配置充足率', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', '配置充足率チェックでエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
    Logger.log('runStaffingCheckMenu error: ' + e.toString());
  }
}

/**
 * 労基チェック結果シートを表示
 */
function showComplianceResults() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.COMPLIANCE_RESULT);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('情報', '労基チェック結果はまだありません。', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 労基ルール設定シートを表示
 */
function showLaborRulesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.LABOR_RULES);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    // シートが無ければ作成
    setupLaborRulesSheet();
    sheet = ss.getSheetByName(SHEET_NAMES.LABOR_RULES);
    if (sheet) ss.setActiveSheet(sheet);
  }
}

/**
 * シフト取込後の自動コンプライアンスチェック
 * @param {string} targetMonth - 対象年月
 * @private
 */
function runAutoComplianceCheck_(targetMonth) {
  try {
    var result = runComplianceCheck(targetMonth);
    writeComplianceResults(targetMonth, result);

    if (result.summary.totalViolations > 0) {
      // 重大違反があればアラート表示
      var ui = SpreadsheetApp.getUi();
      var msg = '【労基チェック自動実行結果】\n' +
        '違反: ' + result.summary.totalViolations + '件\n' +
        '警告: ' + result.summary.totalWarnings + '件\n\n';

      for (var i = 0; i < Math.min(result.violations.length, 5); i++) {
        var v = result.violations[i];
        msg += '- ' + v.employeeName + ': ' + v.detail + '\n';
      }

      ui.alert('労基違反検知', msg, ui.ButtonSet.OK);

      // 管理者へLINE Push通知
      notifyComplianceViolation_(targetMonth, result);
    }

    Logger.log('Auto compliance check: ' + result.summary.totalViolations + ' violations, ' +
               result.summary.totalWarnings + ' warnings');
  } catch (e) {
    Logger.log('runAutoComplianceCheck_ error: ' + e.toString());
  }
}

/**
 * 重大違反時の管理者LINE通知
 * @param {string} targetMonth - 対象年月
 * @param {Object} result - チェック結果
 * @private
 */
function notifyComplianceViolation_(targetMonth, result) {
  try {
    var adminUserId = getSettingValue(SETTING_KEYS.ADMIN_NOTIFY_USER_ID);
    if (!adminUserId) return;

    var parts = targetMonth.split('-');
    var displayMonth = parts[0] + '年' + parseInt(parts[1], 10) + '月';

    var text = '【労基違反検知】' + displayMonth + '\n' +
      '違反: ' + result.summary.totalViolations + '件\n' +
      '警告: ' + result.summary.totalWarnings + '件\n\n';

    for (var i = 0; i < Math.min(result.violations.length, 3); i++) {
      var v = result.violations[i];
      text += v.employeeName + ': ' + v.detail + '\n';
    }
    if (result.violations.length > 3) {
      text += '...他' + (result.violations.length - 3) + '件';
    }

    text += '\n\n管理者ダッシュボードで詳細を確認してください。';

    pushMessage(adminUserId, [createTextMessage(text)]);
  } catch (e) {
    Logger.log('notifyComplianceViolation_ error: ' + e.toString());
  }
}
