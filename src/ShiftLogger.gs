/**
 * ShiftLogger.gs - 送信ログ記録
 * シフト通知の送信結果を記録・管理
 */

/**
 * 送信結果をログに記録
 * @param {string} targetMonth - 対象年月
 * @param {string} employeeNo - 社員番号
 * @param {string} name - 氏名
 * @param {string} status - ステータス (成功/失敗/スキップ/LINE未登録/無効)
 * @param {string} errorDetail - エラー詳細
 */
function logSendResult(targetMonth, employeeNo, name, status, errorDetail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SEND_LOG);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAMES.SEND_LOG);
      sheet.getRange(1, 1, 1, 6).setValues([[
        '送信日時', '対象年月', '社員No', '氏名', 'ステータス', 'エラー詳細'
      ]]);
      sheet.setFrozenRows(1);
    }

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const lastRow = sheet.getLastRow();

    sheet.getRange(lastRow + 1, 1, 1, 6).setValues([[
      timestamp,
      targetMonth,
      employeeNo,
      name,
      status,
      errorDetail || ''
    ]]);

  } catch (e) {
    // ログ記録のエラーは握りつぶさずGASログに記録
    Logger.log('logSendResult error: ' + e.toString() +
      ' [' + employeeNo + ' ' + name + ' ' + status + ']');
  }
}

/**
 * 送信ログのサマリーを取得
 * @param {string} targetMonth - 対象年月
 * @return {Object} サマリー情報
 */
function getSendLogSummary(targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SEND_LOG);

  const summary = {
    total: 0,
    success: 0,
    failed: 0,
    skipped: 0,
    noLineId: 0,
    inactive: 0
  };

  if (!sheet || sheet.getLastRow() <= 1) return summary;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][LOG_COLS.TARGET_MONTH - 1]).trim() !== targetMonth) continue;

    summary.total++;
    const status = String(data[i][LOG_COLS.STATUS - 1]).trim();

    switch (status) {
      case SEND_STATUS.SUCCESS:
        summary.success++;
        break;
      case SEND_STATUS.FAILED:
        summary.failed++;
        break;
      case SEND_STATUS.SKIPPED:
        summary.skipped++;
        break;
      case SEND_STATUS.NO_LINE_ID:
        summary.noLineId++;
        break;
      case SEND_STATUS.INACTIVE:
        summary.inactive++;
        break;
    }
  }

  return summary;
}

/**
 * 特定月の送信ログをクリア
 * @param {string} targetMonth - 対象年月
 */
function clearSendLog(targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SEND_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return;

  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][LOG_COLS.TARGET_MONTH - 1]).trim() === targetMonth) {
      rowsToDelete.push(i + 1);
    }
  }

  for (const row of rowsToDelete) {
    sheet.deleteRow(row);
  }
}

/**
 * 送信失敗者の一覧を取得
 * @param {string} targetMonth - 対象年月
 * @return {Array<Object>} 失敗者リスト
 */
function getFailedSendList(targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SEND_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const failed = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][LOG_COLS.TARGET_MONTH - 1]).trim() !== targetMonth) continue;
    if (String(data[i][LOG_COLS.STATUS - 1]).trim() !== SEND_STATUS.FAILED) continue;

    failed.push({
      employeeNo: String(data[i][LOG_COLS.EMPLOYEE_NO - 1]).trim(),
      name: String(data[i][LOG_COLS.NAME - 1]).trim(),
      error: String(data[i][LOG_COLS.ERROR_DETAIL - 1]).trim()
    });
  }

  return failed;
}
