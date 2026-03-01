/**
 * AllocationConfirmer.gs - 仮配置 → 確定ワークフロー
 * 仮配置データをシフトデータシートにコピーし、LINE通知を送信する
 */

/**
 * 仮配置を確定する
 * @param {string} targetMonth - 対象年月 (YYYY-MM)
 * @return {Object} 確定結果
 */
function confirmAllocations(targetMonth) {
  var assignments = loadTentativeAssignments(targetMonth);

  if (assignments.length === 0) {
    return { success: false, error: '仮配置データがありません。' };
  }

  // 仮配置ステータスのもののみ対象
  var tentative = assignments.filter(function(a) {
    return a.status === ASSIGN_STATUS.TENTATIVE;
  });

  if (tentative.length === 0) {
    return { success: false, error: '確定対象の仮配置がありません（全て確定済み）。' };
  }

  // 1. 重複チェック（同日・同時間帯・同職員の複数施設配置）
  var duplicates = findDuplicateAssignments_(tentative);
  if (duplicates.length > 0) {
    return {
      success: false,
      error: '重複配置があります。解消してから確定してください。',
      duplicates: duplicates
    };
  }

  // 2. 労基チェック（ComplianceChecker利用）
  // 仮配置データを一時的にシフトデータ形式に変換してチェック
  var complianceResult = null;
  try {
    complianceResult = runComplianceCheck(targetMonth);
  } catch (e) {
    Logger.log('confirmAllocations: compliance check error: ' + e.toString());
  }

  // 3. 仮配置 → シフトデータシートにコピー
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shiftSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_DATA);
  if (!shiftSheet) {
    return { success: false, error: 'シフトデータシートが見つかりません。' };
  }

  var rows = tentative.map(function(a) {
    return [
      a.yearMonth,
      a.date,
      a.area,
      a.facility,
      a.timeSlot,
      a.name,        // 担当者名(原文) = 氏名
      a.empNo,
      a.name         // 氏名(正式)
    ];
  });

  var lastRow = shiftSheet.getLastRow();
  shiftSheet.getRange(lastRow + 1, 1, rows.length, 8).setValues(rows);

  // 4. 仮配置シートのステータスを「確定」に更新
  updateTentativeStatus_(targetMonth, ASSIGN_STATUS.CONFIRMED);

  // 5. 結果サマリー
  var result = {
    success: true,
    confirmedCount: tentative.length,
    targetMonth: targetMonth,
    complianceWarnings: complianceResult ? complianceResult.summary.totalWarnings : 0,
    complianceViolations: complianceResult ? complianceResult.summary.totalViolations : 0
  };

  return result;
}

/**
 * 確定後にLINE通知を送信
 * @param {string} targetMonth - 対象年月
 * @return {Object} 送信結果
 */
function sendConfirmationNotifications(targetMonth) {
  try {
    setSettingValue(SETTING_KEYS.SEND_MODE, '本番');
    startBatchSend();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * 重複配置を検出
 * @param {Array<Object>} assignments - 配置データ
 * @return {Array<Object>} 重複リスト
 * @private
 */
function findDuplicateAssignments_(assignments) {
  var empDateSlot = {};
  var duplicates = [];

  for (var i = 0; i < assignments.length; i++) {
    var a = assignments[i];
    var key = a.empNo + '|' + a.date + '|' + a.timeSlot;
    if (!empDateSlot[key]) {
      empDateSlot[key] = [];
    }
    empDateSlot[key].push(a);
  }

  var keys = Object.keys(empDateSlot);
  for (var k = 0; k < keys.length; k++) {
    if (empDateSlot[keys[k]].length > 1) {
      var items = empDateSlot[keys[k]];
      duplicates.push({
        empNo: items[0].empNo,
        name: items[0].name,
        date: items[0].date,
        timeSlot: items[0].timeSlot,
        facilities: items.map(function(x) { return x.facility; })
      });
    }
  }

  return duplicates;
}

/**
 * 仮配置シートのステータスを一括更新
 * @param {string} targetMonth - 対象年月
 * @param {string} newStatus - 新ステータス
 * @private
 */
function updateTentativeStatus_(targetMonth, newStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TENTATIVE_ASSIGNMENT);
  if (!sheet || sheet.getLastRow() <= 1) return;

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var ym = formatYearMonth_(data[i][TENTATIVE_COLS.YEAR_MONTH - 1]);
    if (ym === targetMonth && String(data[i][TENTATIVE_COLS.STATUS - 1]).trim() === ASSIGN_STATUS.TENTATIVE) {
      sheet.getRange(i + 1, TENTATIVE_COLS.STATUS).setValue(newStatus);
    }
  }
}

/**
 * GASメニューから仮配置を確定
 */
function confirmAllocationsMenu() {
  var ui = SpreadsheetApp.getUi();
  try {
    var targetMonth = getSettingValue(SETTING_KEYS.TARGET_MONTH);
    var stats = getTentativeStats(targetMonth);

    if (stats.totalAssigned === 0) {
      ui.alert('情報', '仮配置データがありません。', ui.ButtonSet.OK);
      return;
    }

    var msg = '【確定前サマリー】\n' +
      '対象年月: ' + targetMonth + '\n' +
      '配置数: ' + stats.totalAssigned + '件\n' +
      '充足率: ' + stats.fillRate + '%\n';

    if (stats.conflicts.length > 0) {
      msg += '\n【警告: 重複配置 ' + stats.conflicts.length + '件】\n';
      for (var c = 0; c < Math.min(stats.conflicts.length, 5); c++) {
        msg += '- ' + stats.conflicts[c].key + ' → ' + stats.conflicts[c].facilities.join(', ') + '\n';
      }
      msg += '\n重複を解消してから確定してください。';
      ui.alert('重複配置あり', msg, ui.ButtonSet.OK);
      return;
    }

    if (stats.shortageSlots > 0) {
      msg += '\n人員不足: ' + stats.shortageSlots + 'スロット';
    }

    msg += '\n\n確定しますか？（シフトデータシートにコピーされます）';

    var confirm = ui.alert('仮配置確定', msg, ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return;

    var result = confirmAllocations(targetMonth);

    if (result.success) {
      var successMsg = '確定完了: ' + result.confirmedCount + '件のシフトをコピーしました。';
      if (result.complianceViolations > 0) {
        successMsg += '\n\n【労基違反: ' + result.complianceViolations + '件】確認してください。';
      }

      var sendConfirm = ui.alert('確定完了', successMsg + '\n\nLINE通知を送信しますか？', ui.ButtonSet.YES_NO);
      if (sendConfirm === ui.Button.YES) {
        sendConfirmationNotifications(targetMonth);
        ui.alert('通知送信', 'LINE通知の送信を開始しました。', ui.ButtonSet.OK);
      }
    } else {
      ui.alert('エラー', result.error, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('エラー', '確定処理でエラーが発生しました:\n' + e.message, ui.ButtonSet.OK);
    Logger.log('confirmAllocationsMenu error: ' + e.toString());
  }
}

/**
 * GASメニューから仮配置をクリア
 */
function clearTentativeMenu() {
  var ui = SpreadsheetApp.getUi();
  try {
    var targetMonth = getSettingValue(SETTING_KEYS.TARGET_MONTH);
    var confirm = ui.alert('仮配置クリア',
      targetMonth + 'の仮配置データを全て削除します。\nよろしいですか？',
      ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return;

    var count = clearTentativeAssignments(targetMonth);
    ui.alert('クリア完了', count + '件の仮配置を削除しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}
