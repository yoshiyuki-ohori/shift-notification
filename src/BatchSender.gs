/**
 * BatchSender.gs - バッチ送信制御
 * GAS 6分制限対策のバッチ分割・Trigger連鎖
 */

// PropertiesServiceのキー
const BATCH_PROPS = {
  BATCH_INDEX: 'SHIFT_BATCH_INDEX',
  BATCH_TARGET_MONTH: 'SHIFT_BATCH_TARGET_MONTH',
  BATCH_SEND_MODE: 'SHIFT_BATCH_SEND_MODE',
  BATCH_IN_PROGRESS: 'SHIFT_BATCH_IN_PROGRESS'
};

/**
 * バッチ送信を開始
 * カスタムメニューから呼び出されるエントリポイント
 */
function startBatchSend() {
  const targetMonth = getSettingValue(SETTING_KEYS.TARGET_MONTH);
  const sendMode = getSettingValue(SETTING_KEYS.SEND_MODE);

  // 既存のバッチ処理が走っていないか確認
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(BATCH_PROPS.BATCH_IN_PROGRESS) === 'true') {
    throw new Error('バッチ送信が既に実行中です。完了をお待ちください。');
  }

  // シフトデータを集約
  const aggregated = aggregateByEmployee(targetMonth);
  if (aggregated.length === 0) {
    throw new Error('送信対象のシフトデータがありません。名寄せ済みのデータを確認してください。');
  }

  const summary = getAggregationSummary(aggregated);
  Logger.log('バッチ送信開始: ' + JSON.stringify(summary));

  // バッチ状態を初期化
  props.setProperty(BATCH_PROPS.BATCH_INDEX, '0');
  props.setProperty(BATCH_PROPS.BATCH_TARGET_MONTH, targetMonth);
  props.setProperty(BATCH_PROPS.BATCH_SEND_MODE, sendMode);
  props.setProperty(BATCH_PROPS.BATCH_IN_PROGRESS, 'true');

  // 最初のバッチを実行
  processBatch();
}

/**
 * バッチ処理実行
 * 1バッチ50名ずつ送信、時間超過時はTriggerで続行
 */
function processBatch() {
  const props = PropertiesService.getScriptProperties();
  const startTime = Date.now();

  try {
    const batchIndex = parseInt(props.getProperty(BATCH_PROPS.BATCH_INDEX) || '0', 10);
    const targetMonth = props.getProperty(BATCH_PROPS.BATCH_TARGET_MONTH);
    const sendMode = props.getProperty(BATCH_PROPS.BATCH_SEND_MODE);

    if (!targetMonth) {
      cleanupBatch();
      return;
    }

    // シフトデータを再集約
    const aggregated = aggregateByEmployee(targetMonth);
    const startIdx = batchIndex * BATCH.SEND_BATCH_SIZE;

    if (startIdx >= aggregated.length) {
      // 全件送信完了
      Logger.log('バッチ送信完了: 全 ' + aggregated.length + ' 名の処理が完了');
      cleanupBatch();
      return;
    }

    const endIdx = Math.min(startIdx + BATCH.SEND_BATCH_SIZE, aggregated.length);
    const batch = aggregated.slice(startIdx, endIdx);

    Logger.log('バッチ ' + (batchIndex + 1) + ' 処理中: ' + startIdx + '～' + (endIdx - 1) + ' / ' + aggregated.length + '名');

    for (const emp of batch) {
      // 実行時間チェック
      if (Date.now() - startTime > BATCH.MAX_EXECUTION_MS) {
        Logger.log('実行時間上限に達したため次バッチに繰り越し');
        props.setProperty(BATCH_PROPS.BATCH_INDEX, String(batchIndex));
        scheduleNextBatch();
        return;
      }

      sendShiftToEmployee(emp, targetMonth, sendMode);

      // レート制限対策
      Utilities.sleep(LINE_API.RATE_LIMIT_DELAY_MS);
    }

    // 次のバッチがあるか確認
    if (endIdx < aggregated.length) {
      props.setProperty(BATCH_PROPS.BATCH_INDEX, String(batchIndex + 1));
      scheduleNextBatch();
    } else {
      Logger.log('バッチ送信完了: 全 ' + aggregated.length + ' 名の処理が完了');
      cleanupBatch();
    }

  } catch (e) {
    Logger.log('バッチ処理エラー: ' + e.toString());
    cleanupBatch();
    throw e;
  }
}

/**
 * 個別従業員にシフト通知を送信
 * @param {Object} emp - 従業員集約データ
 * @param {string} targetMonth - 対象年月
 * @param {string} sendMode - 送信モード (テスト/本番)
 */
function sendShiftToEmployee(emp, targetMonth, sendMode) {
  // 送信可否チェック
  if (emp.status === '退職') {
    logSendResult(targetMonth, emp.employeeNo, emp.name, SEND_STATUS.SKIPPED, '退職者');
    return;
  }

  if (!emp.notifyEnabled) {
    logSendResult(targetMonth, emp.employeeNo, emp.name, SEND_STATUS.INACTIVE, '通知無効');
    return;
  }

  // 送信先を決定
  let targetUserId;
  if (sendMode === 'テスト') {
    targetUserId = getSettingValue(SETTING_KEYS.TEST_USER_ID);
    if (!targetUserId) {
      logSendResult(targetMonth, emp.employeeNo, emp.name, SEND_STATUS.SKIPPED, 'テスト送信先未設定');
      return;
    }
  } else {
    targetUserId = emp.lineUserId;
    if (!targetUserId) {
      logSendResult(targetMonth, emp.employeeNo, emp.name, SEND_STATUS.NO_LINE_ID, 'LINE UserId未登録');
      return;
    }
  }

  // Flex Message生成
  const message = buildShiftFlexMessage(targetMonth, emp.name, emp.shifts, emp.employeeNo);

  // 送信
  const result = pushMessage(targetUserId, [message]);

  if (result.success) {
    logSendResult(targetMonth, emp.employeeNo, emp.name, SEND_STATUS.SUCCESS, '');
  } else {
    logSendResult(targetMonth, emp.employeeNo, emp.name, SEND_STATUS.FAILED, result.error);
  }
}

/**
 * 次バッチのTriggerをスケジュール
 */
function scheduleNextBatch() {
  // 既存のprocessBatch Triggerを削除
  deleteProcessBatchTriggers();

  // 1分後にprocessBatchを実行するTriggerを作成
  ScriptApp.newTrigger('processBatch')
    .timeBased()
    .after(BATCH.TRIGGER_DELAY_MINUTES * 60 * 1000)
    .create();

  Logger.log('次バッチを ' + BATCH.TRIGGER_DELAY_MINUTES + ' 分後にスケジュール');
}

/**
 * processBatch関連のTriggerを全て削除
 */
function deleteProcessBatchTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processBatch') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * バッチ状態をクリーンアップ
 */
function cleanupBatch() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(BATCH_PROPS.BATCH_INDEX);
  props.deleteProperty(BATCH_PROPS.BATCH_TARGET_MONTH);
  props.deleteProperty(BATCH_PROPS.BATCH_SEND_MODE);
  props.deleteProperty(BATCH_PROPS.BATCH_IN_PROGRESS);

  deleteProcessBatchTriggers();
}

/**
 * バッチ送信を強制停止（緊急時用）
 */
function forceStopBatch() {
  cleanupBatch();
  Logger.log('バッチ送信を強制停止しました');
  SpreadsheetApp.getUi().alert('強制停止', 'バッチ送信を強制停止しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}
