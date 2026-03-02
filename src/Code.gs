/**
 * Code.gs - エントリポイント・カスタムメニュー
 * シフト通知システムのメイン制御
 */

/**
 * スプレッドシート起動時にカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // メニュー1: シフト管理（カレンダー・メッセージ・同期）
  ui.createMenu('シフト管理')
    .addItem('Googleカレンダー登録', 'syncAllToGoogleCalendar')
    .addItem('一斉メッセージ送信', 'sendBroadcastMessage')
    .addSeparator()
    .addItem('社員マスタ同期 (手動)', 'runManualSync')
    .addItem('毎日自動同期トリガー設定', 'setupDailySyncTrigger')
    .addToUi();

  // メニュー2: 希望・出勤（希望収集 + 出勤確認）
  ui.createMenu('希望・出勤')
    .addItem('希望収集を開始', 'startPreferenceCollection')
    .addItem('希望収集を締切', 'closePreferenceCollection')
    .addItem('提出状況確認', 'showSubmissionStatus')
    .addItem('未提出者にリマインド送信', 'sendPreferenceReminders')
    .addItem('テスト用リンク生成', 'generatePreferenceTestLink')
    .addItem('希望一覧シート表示', 'showPreferenceSheet')
    .addSeparator()
    .addItem('出勤確認テスト送信 (今日)', 'testSendAttendanceConfirm')
    .addItem('未確認者チェック (手動)', 'testCheckUnconfirmed')
    .addItem('出勤確認トリガー設定', 'setupAttendanceTriggers')
    .addItem('出勤確認トリガー削除', 'deleteAttendanceTriggers')
    .addToUi();

  // メニュー3: チェック（労基 + 配置充足率）
  ui.createMenu('チェック')
    .addItem('労基チェック実行', 'runComplianceCheckMenu')
    .addItem('配置充足率チェック', 'runStaffingCheckMenu')
    .addSeparator()
    .addItem('労基チェック結果表示', 'showComplianceResults')
    .addItem('労基ルール設定表示', 'showLaborRulesSheet')
    .addToUi();

  // メニュー4: 配置管理（仮配置→確定ワークフロー）
  ui.createMenu('配置管理')
    .addItem('仮配置を確定', 'confirmAllocationsMenu')
    .addItem('仮配置をクリア', 'clearTentativeMenu')
    .addSeparator()
    .addItem('仮配置シート表示', 'showTentativeSheet')
    .addToUi();
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
      employeeNo: no.padStart(3, '0'),
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
    sheet.getRange(2, 1, 3, 3).setValues([
      [SETTING_KEYS.TARGET_MONTH, '', '送信対象月 (例: 2026-03)'],
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
    sheet.getRange(1, 1, 1, 9).setValues([[
      '対象年月', '社員No', '氏名', '日付', '種別', '時間帯', '理由', '提出日時', '希望施設'
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

  // 出勤確認
  if (!ss.getSheetByName(SHEET_NAMES.ATTENDANCE_CONFIRM)) {
    const sheet = ss.insertSheet(SHEET_NAMES.ATTENDANCE_CONFIRM);
    sheet.getRange(1, 1, 1, 6).setValues([[
      '日付', '社員No', '氏名', 'ステータス', '確認日時', '通知日時'
    ]]);
    sheet.setFrozenRows(1);
  }

  // 仮配置
  if (!ss.getSheetByName(SHEET_NAMES.TENTATIVE_ASSIGNMENT)) {
    const sheet = ss.insertSheet(SHEET_NAMES.TENTATIVE_ASSIGNMENT);
    sheet.getRange(1, 1, 1, 13).setValues([[
      '年月', '日付', 'エリア', '施設名', '施設ID', '時間帯',
      '社員No', '氏名', 'ステータス', '配置理由', '希望一致', '配置日時', '配置者'
    ]]);
    sheet.setFrozenRows(1);
  }

  // 労基ルール
  setupLaborRulesSheet();

  // 労基チェック結果
  setupComplianceResultSheet();

  SpreadsheetApp.getUi().alert('セットアップ完了', '全シートの初期設定が完了しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 仮配置シート表示
 */
function showTentativeSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TENTATIVE_ASSIGNMENT);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('情報', '仮配置データはまだありません。', SpreadsheetApp.getUi().ButtonSet.OK);
  }
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

  // 期間区分を選択
  var periodResult = ui.prompt('期間区分の選択',
    '収集する期間区分を入力してください:\n\n' +
    '  全日 → 1日～末日\n' +
    '  前半 → 1日～15日\n' +
    '  後半 → 16日～末日\n\n' +
    '「全日」「前半」「後半」のいずれかを入力:',
    ui.ButtonSet.OK_CANCEL);
  if (periodResult.getSelectedButton() !== ui.Button.OK) return;
  var periodLabel = periodResult.getResponseText().trim();

  if (['全日', '前半', '後半'].indexOf(periodLabel) === -1) {
    ui.alert('エラー', '期間区分は「全日」「前半」「後半」のいずれかで入力してください。', ui.ButtonSet.OK);
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

  // 収集期間を設定（periodLabel付き）
  setCollectionPeriod(targetMonth, today, deadline, periodLabel);

  // リマインドトリガー設定（締切2日前に自動リマインド）
  setupPreferenceReminderTrigger_(deadline);

  // 期間表示テキスト生成
  var periodDisplay = periodLabel === '全日' ? '' : ' ' + periodLabel;

  // 全職員に希望入力開始通知を送信
  var confirm = ui.alert('通知送信',
    targetMonth + periodDisplay + 'のシフト希望収集を開始します。\n' +
    '締切: ' + deadline + '\n\n' +
    '全職員にLINE通知を送信しますか？',
    ui.ButtonSet.YES_NO);

  if (confirm === ui.Button.YES) {
    var sendResults = sendPreferenceStartNotification_(targetMonth, deadline, periodLabel);

    // 送信結果を表示
    ui.alert('送信結果',
      '送信成功: ' + sendResults.sent + '人\n' +
      '送信失敗: ' + sendResults.failed + '人\n' +
      'LINE未登録: ' + sendResults.noLine + '人',
      ui.ButtonSet.OK);
  }

  ui.alert('開始完了',
    'シフト希望収集を開始しました。\n\n' +
    '対象月: ' + targetMonth + periodDisplay + '\n' +
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
 * 希望入力開始通知を送信
 * @param {string} targetMonth - 対象年月
 * @param {string} deadline - 締切日
 * @param {string} [periodLabel] - 期間区分 ('全日' | '前半' | '後半')
 * @param {string} [targetEmpNo] - 指定時はこの社員番号のみに送信（テスト用）
 * @return {Object} { sent, failed, noLine, details }
 */
function sendPreferenceStartNotification_(targetMonth, deadline, periodLabel, targetEmpNo) {
  var employees = loadEmployeeMaster();
  if (targetEmpNo) {
    targetEmpNo = targetEmpNo.padStart(3, '0');
    employees = employees.filter(function(emp) { return emp.employeeNo === targetEmpNo; });
  }
  var parts = targetMonth.split('-');
  var displayMonth = parts[0] + '年' + parseInt(parts[1], 10) + '月';
  var periodDisplay = (periodLabel && periodLabel !== '全日') ? ' ' + periodLabel : '';
  var results = { sent: 0, failed: 0, noLine: 0, details: [] };

  // LIFF URLを構築（LIFF_IDが未設定の場合はpostbackにフォールバック）
  var liffId = null;
  try {
    liffId = getSettingValue(SETTING_KEYS.LIFF_ID);
  } catch (e) {
    Logger.log('LIFF_ID not set, using postback fallback');
  }
  var liffUrl = liffId ? 'https://liff.line.me/' + liffId + '/preference/?month=' + targetMonth : '';

  employees.forEach(function(emp) {
    if (emp.status !== '在職') return;
    if (!emp.lineUserId) {
      results.noLine++;
      results.details.push({ name: emp.name, status: 'LINE未登録' });
      return;
    }

    // ボタンアクション: LIFF URLがあればuri、なければpostback
    var buttonAction = liffUrl
      ? { type: 'uri', label: 'シフト希望を入力する', uri: liffUrl }
      : { type: 'postback', label: 'シフト希望を入力する', data: 'action=pref_start&month=' + targetMonth, displayText: 'シフト希望入力' };

    var message = {
      type: 'flex',
      altText: displayMonth + periodDisplay + ' シフト希望入力のお知らせ',
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
              text: displayMonth + periodDisplay + 'のシフト希望入力を受付中です。',
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
            action: buttonAction,
            style: 'primary',
            color: '#1DB446',
            height: 'sm'
          }],
          paddingAll: '12px'
        }
      }
    };

    var result = pushMessage(emp.lineUserId, [message]);
    if (result.success) {
      results.sent++;
      results.details.push({ name: emp.name, status: '送信成功' });
    } else {
      results.failed++;
      results.details.push({ name: emp.name, status: '送信失敗', error: result.error });
    }
    Utilities.sleep(LINE_API.RATE_LIMIT_DELAY_MS);
  });

  return results;
}

/**
 * テスト用LIFF希望入力リンクを生成
 */
function generatePreferenceTestLink() {
  var ui = SpreadsheetApp.getUi();
  var period = getCollectionPeriod();
  var targetMonth = period ? period.targetMonth : '';

  if (!targetMonth) {
    var result = ui.prompt('テスト用リンク生成',
      '対象年月を入力してください (例: 2026-04)',
      ui.ButtonSet.OK_CANCEL);
    if (result.getSelectedButton() !== ui.Button.OK) return;
    targetMonth = result.getResponseText().trim();

    if (!targetMonth.match(/^\d{4}-\d{2}$/)) {
      ui.alert('エラー', '年月の形式が正しくありません。例: 2026-04', ui.ButtonSet.OK);
      return;
    }
  }

  var liffId;
  try {
    liffId = getSettingValue(SETTING_KEYS.LIFF_ID);
  } catch (e) {
    ui.alert('エラー', 'LIFF_IDが設定シートに登録されていません。', ui.ButtonSet.OK);
    return;
  }

  var liffUrl = 'https://liff.line.me/' + liffId + '/preference/?month=' + targetMonth;
  var periodLabel = period ? (period.periodLabel || '全日') : '全日';

  ui.alert('テスト用リンク',
    '以下のURLをLINEアプリで開いてください:\n\n' +
    liffUrl + '\n\n' +
    '対象月: ' + targetMonth + '\n' +
    '期間区分: ' + periodLabel + '\n\n' +
    '※ LINE連携済みのアカウントでアクセスしてください\n' +
    '※ LINEアプリ内ブラウザで開く必要があります',
    ui.ButtonSet.OK);
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

// ===== 出勤確認関連 =====

/**
 * 出勤確認テスト送信（メニューから手動実行）
 */
function testSendAttendanceConfirm() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert('出勤確認テスト',
    '当日シフトのある職員に出勤確認メッセージを送信します。\n実行しますか？',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  sendDailyAttendanceConfirm();
  ui.alert('完了', '出勤確認メッセージの送信が完了しました。', ui.ButtonSet.OK);
}

/**
 * 未確認者チェック手動実行（メニューから）
 */
function testCheckUnconfirmed() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert('未確認者チェック',
    '未確認者に再通知し、管理者にアラートを送信します。\n実行しますか？',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  checkUnconfirmedAttendance();
  ui.alert('完了', '未確認者チェックが完了しました。', ui.ButtonSet.OK);
}

/**
 * 出勤確認の定期トリガーを設定
 * 毎朝7:00 → sendDailyAttendanceConfirm
 * 毎日12:00 → checkUnconfirmedAttendance
 */
function setupAttendanceTriggers() {
  // 既存トリガーを削除
  deleteAttendanceTriggers();

  // 毎朝7時: 出勤確認送信
  ScriptApp.newTrigger('sendDailyAttendanceConfirm')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .nearMinute(0)
    .create();

  // 毎日12時: 未確認チェック
  ScriptApp.newTrigger('checkUnconfirmedAttendance')
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .nearMinute(0)
    .create();

  SpreadsheetApp.getUi().alert('トリガー設定完了',
    '出勤確認トリガーを設定しました。\n\n' +
    '- 毎朝 7:00 出勤確認送信\n' +
    '- 毎日 12:00 未確認者チェック',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 出勤確認トリガーを削除
 */
function deleteAttendanceTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'sendDailyAttendanceConfirm' || fn === 'checkUnconfirmedAttendance') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
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
