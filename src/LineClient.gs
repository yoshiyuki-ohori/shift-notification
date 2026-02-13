/**
 * LineClient.gs - LINE Messaging API呼び出し
 * Push Message送信機能
 */

/**
 * LINE Push Messageを送信
 * @param {string} userId - LINE UserID
 * @param {Array<Object>} messages - メッセージ配列（最大5件）
 * @return {Object} 送信結果 {success: boolean, error: string}
 */
function pushMessage(userId, messages) {
  if (!userId || !messages || messages.length === 0) {
    return { success: false, error: 'userId またはメッセージが空です' };
  }

  const url = LINE_API.PUSH_URL;
  const token = getLineToken();

  const payload = {
    to: userId,
    messages: messages.slice(0, 5) // LINE APIの制限: 最大5メッセージ
  };

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode === 200) {
      return { success: true, error: '' };
    }

    const responseBody = response.getContentText();
    let errorMessage = 'HTTP ' + statusCode;
    try {
      const errorJson = JSON.parse(responseBody);
      errorMessage += ': ' + (errorJson.message || responseBody);
    } catch (e) {
      errorMessage += ': ' + responseBody;
    }

    return { success: false, error: errorMessage };

  } catch (e) {
    return { success: false, error: 'UrlFetchApp error: ' + e.toString() };
  }
}

/**
 * LINE Reply Messageを送信
 * @param {string} replyToken - 返信トークン
 * @param {Array<Object>} messages - メッセージ配列
 * @return {Object} 送信結果
 */
function replyLineMessage(replyToken, messages) {
  if (!replyToken || !messages || messages.length === 0) {
    return { success: false, error: 'replyToken またはメッセージが空です' };
  }

  const url = LINE_API.REPLY_URL;
  const token = getLineToken();

  const payload = {
    replyToken: replyToken,
    messages: messages.slice(0, 5)
  };

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode === 200) {
      return { success: true, error: '' };
    }

    return { success: false, error: 'HTTP ' + statusCode + ': ' + response.getContentText() };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * LINE Broadcast Messageを送信（全友だちに一斉送信）
 * @param {Array<Object>} messages - メッセージ配列（最大5件）
 * @return {Object} 送信結果 {success: boolean, error: string}
 */
function broadcastMessage(messages) {
  if (!messages || messages.length === 0) {
    return { success: false, error: 'メッセージが空です' };
  }

  const url = LINE_API.BROADCAST_URL;
  const token = getLineToken();

  const payload = {
    messages: messages.slice(0, 5)
  };

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode === 200) {
      return { success: true, error: '' };
    }

    const responseBody = response.getContentText();
    let errorMessage = 'HTTP ' + statusCode;
    try {
      const errorJson = JSON.parse(responseBody);
      errorMessage += ': ' + (errorJson.message || responseBody);
    } catch (e) {
      errorMessage += ': ' + responseBody;
    }

    return { success: false, error: errorMessage };

  } catch (e) {
    return { success: false, error: 'UrlFetchApp error: ' + e.toString() };
  }
}

/**
 * テキストメッセージオブジェクトを作成
 * @param {string} text - メッセージテキスト
 * @return {Object} テキストメッセージ
 */
function createTextMessage(text) {
  return {
    type: 'text',
    text: text
  };
}

/**
 * LINE APIの送信可能残数を確認
 * @return {Object} クォータ情報
 */
function checkMessageQuota() {
  const token = getLineToken();
  const url = 'https://api.line.me/v2/bot/message/quota/consumption';

  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + token
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log('checkMessageQuota error: ' + e.toString());
    return null;
  }
}
