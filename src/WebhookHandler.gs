/**
 * WebhookHandler.gs - LINE Webhook (UserID登録用)
 * 職員がLINE公式アカウントに社員番号を送信→UserID自動登録
 *
 * フロー:
 * 1. 職員が「社員番号」(例: "072") を送信
 * 2. 社員番号を従業員マスタと照合
 * 3. マッチしたらLINE_UserIdを自動登録
 * 4. 登録完了メッセージを返信
 */

/**
 * WebApp エントリポイント - POSTリクエスト処理
 * @param {Object} e - イベントオブジェクト
 * @return {ContentService} レスポンス
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // LINE署名検証
    if (!verifySignature(e)) {
      return ContentService.createTextOutput('Unauthorized')
        .setMimeType(ContentService.MimeType.TEXT);
    }

    // イベント処理
    if (body.events && body.events.length > 0) {
      body.events.forEach(function(event) {
        processWebhookEvent(event);
      });
    }

    return ContentService.createTextOutput('OK')
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    Logger.log('Webhook Error: ' + error.toString());
    return ContentService.createTextOutput('Error')
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * GETリクエスト処理（Webhook URL検証用）
 */
function doGet(e) {
  return ContentService.createTextOutput('Shift Notification Webhook Active')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * LINE署名検証
 * @param {Object} e - イベントオブジェクト
 * @return {boolean} 検証結果
 */
function verifySignature(e) {
  try {
    const channelSecret = getLineChannelSecret();
    const signature = e.parameter ? e.parameter['x-line-signature'] : null;

    // 署名が無い場合（開発環境等）
    if (!signature) {
      Logger.log('Warning: No signature provided');
      return true; // 開発中は通す。本番では false にすること
    }

    const body = e.postData.contents;
    const hash = Utilities.computeHmacSha256Signature(body, channelSecret);
    const base64Hash = Utilities.base64Encode(hash);

    return signature === base64Hash;
  } catch (error) {
    Logger.log('Signature verification error: ' + error.toString());
    return false;
  }
}

/**
 * Webhookイベント処理
 * @param {Object} event - LINEイベント
 */
function processWebhookEvent(event) {
  const userId = event.source.userId;
  const replyToken = event.replyToken;

  switch (event.type) {
    case 'message':
      if (event.message.type === 'text') {
        handleRegistrationMessage(event.message.text, userId, replyToken);
      }
      break;

    case 'follow':
      handleFollowEvent_(userId, replyToken);
      break;

    default:
      break;
  }
}

/**
 * テキストメッセージ処理 - 社員番号登録
 * @param {string} text - メッセージテキスト
 * @param {string} userId - LINE UserId
 * @param {string} replyToken - 返信トークン
 */
function handleRegistrationMessage(text, userId, replyToken) {
  const trimmedText = text.trim();

  // 社員番号パターン判定 (001-999)
  const empNoMatch = trimmedText.match(/^(\d{1,3})$/);
  if (!empNoMatch) {
    // 社員番号以外のメッセージ
    replyLineMessage(replyToken, [
      createTextMessage(
        '社員番号を入力してください。\n' +
        '例: 072\n\n' +
        '※社員番号が分からない場合は管理者にお問い合わせください。'
      )
    ]);
    return;
  }

  // 社員番号を3桁にゼロパディング
  const employeeNo = String(parseInt(empNoMatch[1], 10)).padStart(3, '0');

  // 従業員マスタから検索
  const employee = findEmployeeByNo(employeeNo);
  if (!employee) {
    replyLineMessage(replyToken, [
      createTextMessage(
        '社員番号「' + employeeNo + '」は見つかりませんでした。\n' +
        '正しい社員番号を入力してください。\n\n' +
        '※社員番号が分からない場合は管理者にお問い合わせください。'
      )
    ]);
    return;
  }

  // 既に登録済みか確認
  if (employee.lineUserId === userId) {
    replyLineMessage(replyToken, [
      createTextMessage(
        employee.name + 'さん、既にLINE連携済みです。\n' +
        'シフト通知は自動的に届きます。'
      )
    ]);
    return;
  }

  // LINE UserIdをマスタに登録
  const success = registerLineUserId(employeeNo, userId);

  if (success) {
    replyLineMessage(replyToken, [
      createTextMessage(
        employee.name + 'さん、LINE連携が完了しました！\n\n' +
        'シフトが確定すると、こちらのLINEに通知が届きます。'
      )
    ]);
    Logger.log('LINE UserId registered: empNo=' + employeeNo);
  } else {
    replyLineMessage(replyToken, [
      createTextMessage(
        '登録処理でエラーが発生しました。\n' +
        'お手数ですが管理者にご連絡ください。'
      )
    ]);
  }
}

/**
 * フォローイベント処理（友だち追加時）
 * @param {string} userId - LINE UserId
 * @param {string} replyToken - 返信トークン
 */
function handleFollowEvent_(userId, replyToken) {
  replyLineMessage(replyToken, [
    createTextMessage(
      'シフト通知システムへようこそ！\n\n' +
      'LINE連携を行うため、社員番号を入力してください。\n' +
      '例: 072\n\n' +
      '※社員番号が分からない場合は管理者にお問い合わせください。'
    )
  ]);
}

/**
 * 従業員マスタから社員番号で検索
 * @param {string} employeeNo - 社員番号
 * @return {Object|null} 従業員データ
 */
function findEmployeeByNo(employeeNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEE_MASTER);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const no = String(data[i][MASTER_COLS.NO - 1]).trim();
    // ゼロパディングして比較
    const paddedNo = no.padStart(3, '0');
    if (paddedNo === employeeNo) {
      return {
        employeeNo: paddedNo,
        name: String(data[i][MASTER_COLS.NAME - 1]).trim(),
        lineUserId: String(data[i][MASTER_COLS.LINE_USER_ID - 1] || '').trim(),
        rowIndex: i + 1 // シート上の行番号(1-indexed)
      };
    }
  }
  return null;
}

/**
 * 従業員マスタにLINE UserIdを登録
 * @param {string} employeeNo - 社員番号
 * @param {string} userId - LINE UserId
 * @return {boolean} 成功/失敗
 */
function registerLineUserId(employeeNo, userId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEE_MASTER);
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const no = String(data[i][MASTER_COLS.NO - 1]).trim().padStart(3, '0');
      if (no === employeeNo) {
        sheet.getRange(i + 1, MASTER_COLS.LINE_USER_ID).setValue(userId);
        return true;
      }
    }
    return false;
  } catch (e) {
    Logger.log('registerLineUserId error: ' + e.toString());
    return false;
  }
}
