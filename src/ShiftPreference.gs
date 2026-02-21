/**
 * ShiftPreference.gs - シフト希望データ管理
 * 職員のシフト希望を収集・保存・集計する
 */

/**
 * シフト希望を保存
 * @param {Object} pref - 希望データ
 * @param {string} pref.yearMonth - 対象年月 (2026-03)
 * @param {string} pref.employeeNo - 社員番号
 * @param {string} pref.name - 氏名
 * @param {string} pref.date - 日付 (YYYY/MM/DD)
 * @param {string} pref.type - 希望/NG/どちらでも
 * @param {string} [pref.timeSlot] - 時間帯
 * @param {string} [pref.reason] - 理由
 * @return {boolean} 成功/失敗
 */
function savePreference(pref) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAMES.SHIFT_PREFERENCE);
      sheet.getRange(1, 1, 1, 8).setValues([[
        '対象年月', '社員No', '氏名', '日付', '種別', '時間帯', '理由', '提出日時'
      ]]);
      sheet.setFrozenRows(1);
    }

    // 既存の同一希望を削除（上書き）
    deletePrefIfExists_(sheet, pref.yearMonth, pref.employeeNo, pref.date, pref.type);

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, 8).setValues([[
      pref.yearMonth,
      pref.employeeNo,
      pref.name,
      pref.date,
      pref.type,
      pref.timeSlot || '',
      pref.reason || '',
      new Date()
    ]]);

    return true;
  } catch (e) {
    Logger.log('savePreference error: ' + e.toString());
    return false;
  }
}

/**
 * 複数の希望を一括保存
 * @param {Array<Object>} prefs - 希望データ配列
 * @return {number} 保存件数
 */
function savePreferences(prefs) {
  if (!prefs || prefs.length === 0) return 0;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAMES.SHIFT_PREFERENCE);
      sheet.getRange(1, 1, 1, 8).setValues([[
        '対象年月', '社員No', '氏名', '日付', '種別', '時間帯', '理由', '提出日時'
      ]]);
      sheet.setFrozenRows(1);
    }

    // 同一職員・同一月の既存データを削除
    deletePrefsForEmployee_(sheet, prefs[0].yearMonth, prefs[0].employeeNo);

    const now = new Date();
    const rows = prefs.map(function(p) {
      return [
        p.yearMonth,
        p.employeeNo,
        p.name,
        p.date,
        p.type,
        p.timeSlot || '',
        p.reason || '',
        now
      ];
    });

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, 8).setValues(rows);

    return rows.length;
  } catch (e) {
    Logger.log('savePreferences error: ' + e.toString());
    return 0;
  }
}

/**
 * 特定職員の希望を取得
 * @param {string} yearMonth - 対象年月
 * @param {string} employeeNo - 社員番号
 * @return {Array<Object>} 希望配列
 */
function getPreferencesForEmployee(yearMonth, employeeNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][PREF_COLS.YEAR_MONTH - 1]).trim() === yearMonth &&
        String(data[i][PREF_COLS.EMPLOYEE_NO - 1]).trim().padStart(3, '0') === employeeNo) {
      results.push({
        date: String(data[i][PREF_COLS.DATE - 1]).trim(),
        type: String(data[i][PREF_COLS.TYPE - 1]).trim(),
        timeSlot: String(data[i][PREF_COLS.TIME_SLOT - 1] || '').trim(),
        reason: String(data[i][PREF_COLS.REASON - 1] || '').trim()
      });
    }
  }

  return results;
}

/**
 * 全職員の希望提出状況を取得
 * @param {string} yearMonth - 対象年月
 * @return {Object} { submitted: [{empNo, name, count, submittedAt}], notSubmitted: [{empNo, name}] }
 */
function getSubmissionStatus(yearMonth) {
  const employees = loadEmployeeMaster();
  const activeEmployees = employees.filter(function(e) { return e.status === '在職'; });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);

  const submittedMap = {};

  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][PREF_COLS.YEAR_MONTH - 1]).trim() !== yearMonth) continue;
      var empNo = String(data[i][PREF_COLS.EMPLOYEE_NO - 1]).trim().padStart(3, '0');
      if (!submittedMap[empNo]) {
        submittedMap[empNo] = {
          count: 0,
          submittedAt: data[i][PREF_COLS.SUBMITTED_AT - 1]
        };
      }
      submittedMap[empNo].count++;
      // 最新の提出日時を保持
      var ts = data[i][PREF_COLS.SUBMITTED_AT - 1];
      if (ts && ts > submittedMap[empNo].submittedAt) {
        submittedMap[empNo].submittedAt = ts;
      }
    }
  }

  var submitted = [];
  var notSubmitted = [];

  activeEmployees.forEach(function(emp) {
    var empNo = emp.employeeNo.padStart(3, '0');
    if (submittedMap[empNo]) {
      submitted.push({
        employeeNo: empNo,
        name: emp.name,
        count: submittedMap[empNo].count,
        submittedAt: submittedMap[empNo].submittedAt
      });
    } else {
      notSubmitted.push({
        employeeNo: empNo,
        name: emp.name,
        lineUserId: emp.lineUserId
      });
    }
  });

  return { submitted: submitted, notSubmitted: notSubmitted };
}

/**
 * 希望収集期間を設定
 * @param {string} targetMonth - 対象年月 (2026-03)
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 締切日 (YYYY-MM-DD)
 */
function setCollectionPeriod(targetMonth, startDate, endDate) {
  setSettingValue(SETTING_KEYS.PREF_TARGET_MONTH, targetMonth);
  setSettingValue(SETTING_KEYS.PREF_COLLECTION_START, startDate);
  setSettingValue(SETTING_KEYS.PREF_COLLECTION_END, endDate);
}

/**
 * 希望収集期間を取得
 * @return {Object|null} {targetMonth, startDate, endDate} or null
 */
function getCollectionPeriod() {
  try {
    return {
      targetMonth: getSettingValue(SETTING_KEYS.PREF_TARGET_MONTH),
      startDate: getSettingValue(SETTING_KEYS.PREF_COLLECTION_START),
      endDate: getSettingValue(SETTING_KEYS.PREF_COLLECTION_END)
    };
  } catch (e) {
    return null;
  }
}

/**
 * 希望収集中かどうか判定
 * @return {boolean}
 */
function isCollectionOpen() {
  var period = getCollectionPeriod();
  if (!period || !period.startDate || !period.endDate) return false;

  var now = new Date();
  var start = new Date(period.startDate + 'T00:00:00+09:00');
  var end = new Date(period.endDate + 'T23:59:59+09:00');

  return now >= start && now <= end;
}

/**
 * 日別の希望集計（ヒートマップ用）
 * @param {string} yearMonth - 対象年月
 * @return {Object} { "2026/03/01": { want: 5, ng: 2, either: 1 }, ... }
 */
function getPreferenceSummaryByDate(yearMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);
  if (!sheet || sheet.getLastRow() <= 1) return {};

  const data = sheet.getDataRange().getValues();
  var summary = {};

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][PREF_COLS.YEAR_MONTH - 1]).trim() !== yearMonth) continue;
    var date = String(data[i][PREF_COLS.DATE - 1]).trim();
    var type = String(data[i][PREF_COLS.TYPE - 1]).trim();

    if (!summary[date]) {
      summary[date] = { want: 0, ng: 0, either: 0 };
    }

    if (type === PREF_TYPE.WANT) summary[date].want++;
    else if (type === PREF_TYPE.NG) summary[date].ng++;
    else summary[date].either++;
  }

  return summary;
}

/**
 * 未提出者にリマインド送信
 * @param {string} yearMonth - 対象年月
 * @return {Object} { sent: number, failed: number }
 */
function sendPreferenceReminder(yearMonth) {
  var status = getSubmissionStatus(yearMonth);
  var period = getCollectionPeriod();
  var deadline = period ? period.endDate : '未設定';

  var sent = 0;
  var failed = 0;

  status.notSubmitted.forEach(function(emp) {
    if (!emp.lineUserId) {
      failed++;
      return;
    }

    var message = buildPreferenceReminderMessage_(yearMonth, emp.name, deadline);
    var result = pushMessage(emp.lineUserId, [message]);
    if (result.success) {
      sent++;
    } else {
      failed++;
      Logger.log('Reminder failed for ' + emp.name + ': ' + result.error);
    }
  });

  return { sent: sent, failed: failed };
}

/**
 * 提出済みとしてマーク（スクリプトプロパティに記録）
 * @param {string} employeeNo - 社員番号
 * @param {string} yearMonth - 対象年月 (YYYY-MM)
 * @return {string} ISO日時文字列
 */
function markAsSubmitted(employeeNo, yearMonth) {
  var key = 'PREF_SUBMITTED_' + employeeNo + '_' + yearMonth;
  var now = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(key, now);
  return now;
}

/**
 * 最終提出日時を取得
 * @param {string} employeeNo - 社員番号
 * @param {string} yearMonth - 対象年月 (YYYY-MM)
 * @return {string|null} ISO日時文字列 or null
 */
function getSubmittedAt(employeeNo, yearMonth) {
  var key = 'PREF_SUBMITTED_' + employeeNo + '_' + yearMonth;
  return PropertiesService.getScriptProperties().getProperty(key) || null;
}

/**
 * 提出時に管理者へLINE Push通知を送信
 * @param {string} employeeName - 従業員名
 * @param {string} yearMonth - 対象年月 (YYYY-MM)
 * @param {number} wantCount - 希望日数
 * @param {number} ngCount - NG日数
 */
function notifyAdminOnSubmission_(employeeName, yearMonth, wantCount, ngCount) {
  try {
    var adminUserId = getSettingValue(SETTING_KEYS.ADMIN_NOTIFY_USER_ID);
    if (!adminUserId) return;

    var parts = yearMonth.split('-');
    var displayMonth = parts[0] + '年' + parseInt(parts[1], 10) + '月';

    var text = employeeName + 'さんが' + displayMonth + 'のシフト希望を提出しました' +
      ' (希望:' + wantCount + '日, NG:' + ngCount + '日)';

    pushMessage(adminUserId, [createTextMessage(text)]);
  } catch (e) {
    Logger.log('notifyAdminOnSubmission_ error: ' + e.toString());
  }
}

// ===== Private helpers =====

/**
 * 既存の同一希望を削除
 */
function deletePrefIfExists_(sheet, yearMonth, employeeNo, date, type) {
  if (sheet.getLastRow() <= 1) return;

  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === yearMonth &&
        String(data[i][1]).trim().padStart(3, '0') === employeeNo &&
        String(data[i][3]).trim() === date &&
        String(data[i][4]).trim() === type) {
      sheet.deleteRow(i + 1);
    }
  }
}

/**
 * 特定職員の月の全希望を削除
 */
function deletePrefsForEmployee_(sheet, yearMonth, employeeNo) {
  if (sheet.getLastRow() <= 1) return;

  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === yearMonth &&
        String(data[i][1]).trim().padStart(3, '0') === employeeNo) {
      sheet.deleteRow(i + 1);
    }
  }
}

/**
 * リマインドメッセージ生成
 */
function buildPreferenceReminderMessage_(yearMonth, name, deadline) {
  var parts = yearMonth.split('-');
  var displayMonth = parts[0] + '年' + parseInt(parts[1], 10) + '月';

  return {
    type: 'flex',
    altText: displayMonth + 'シフト希望 提出のお願い',
    contents: {
      type: 'bubble',
      size: 'kilo',
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [{
          type: 'text',
          text: 'シフト希望 提出のお願い',
          weight: 'bold',
          size: 'md',
          color: '#FFFFFF'
        }],
        backgroundColor: '#E74C3C',
        paddingAll: '12px'
      },
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'text',
            text: name + 'さん',
            weight: 'bold',
            size: 'md'
          },
          {
            type: 'text',
            text: displayMonth + 'のシフト希望がまだ提出されていません。',
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
            data: 'action=pref_start&month=' + yearMonth
          },
          style: 'primary',
          color: '#1DB446',
          height: 'sm'
        }],
        paddingAll: '12px'
      }
    }
  };
}
