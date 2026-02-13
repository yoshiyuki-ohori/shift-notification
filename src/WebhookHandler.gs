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

    // Admin API (POST書き込み) - actionフィールドで判別
    if (body.action) {
      var expectedKey = PropertiesService.getScriptProperties().getProperty('ADMIN_API_KEY') || '';
      if (!expectedKey || body.key !== expectedKey) {
        return jsonResponse_({ error: 'Unauthorized' }, 401);
      }
      return handlePostAction_(body);
    }

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
 * GETリクエスト処理 - 管理API + Webhook URL検証
 * クエリパラメータ action で操作を分岐
 *
 * ?action=info           → シート一覧
 * ?action=read&sheet=XX  → シートデータ読み取り
 * ?action=read&sheet=XX&range=A1:Z10 → 範囲指定読み取り
 * ?action=status         → システムステータス
 * (パラメータなし)        → Webhook URL検証
 */
function doGet(e) {
  const params = e ? (e.parameter || {}) : {};
  const action = params.action || '';
  const adminKey = params.key || '';

  // 管理APIキー検証（Script Propertiesに ADMIN_API_KEY を設定）
  // ADMIN_API_KEYが未設定の場合はアクション実行を拒否（デフォルト拒否）
  if (action) {
    const expectedKey = PropertiesService.getScriptProperties().getProperty('ADMIN_API_KEY') || '';
    if (!expectedKey || adminKey !== expectedKey) {
      return jsonResponse_({ error: 'Unauthorized' }, 401);
    }
  }

  try {
    switch (action) {
      case 'info':
        return handleGetInfo_();
      case 'read':
        return handleGetRead_(params);
      case 'status':
        return handleGetStatus_();
      case 'write':
        return handleGetWrite_(params);
      case 'formatText':
        return handleFormatText_(params);
      case 'runFunction':
        return handleRunFunction_(params);
      case 'readExternal':
        return handleReadExternal_(params);
      default:
        return ContentService.createTextOutput('Shift Notification Webhook Active')
          .setMimeType(ContentService.MimeType.TEXT);
    }
  } catch (error) {
    return jsonResponse_({ error: error.toString() }, 500);
  }
}

/**
 * シート一覧を返す
 */
function handleGetInfo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(function(s) {
    return {
      name: s.getName(),
      rows: s.getLastRow(),
      cols: s.getLastColumn()
    };
  });
  return jsonResponse_({ spreadsheetId: ss.getId(), sheets: sheets });
}

/**
 * シートデータを読み取る
 */
function handleGetRead_(params) {
  const sheetName = params.sheet;
  if (!sheetName) return jsonResponse_({ error: 'sheet parameter required' });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse_({ error: 'Sheet not found: ' + sheetName });

  let data;
  if (params.range) {
    data = sheet.getRange(params.range).getValues();
  } else {
    if (sheet.getLastRow() === 0) {
      data = [];
    } else {
      data = sheet.getDataRange().getValues();
    }
  }

  return jsonResponse_({ sheet: sheetName, rows: data.length, data: data });
}

/**
 * システムステータスを返す
 */
function handleGetStatus_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  const masterSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEE_MASTER);
  const shiftSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_DATA);
  const logSheet = ss.getSheetByName(SHEET_NAMES.SEND_LOG);

  return jsonResponse_({
    spreadsheetName: ss.getName(),
    sheets: ss.getSheets().map(function(s) { return s.getName(); }),
    employeeMasterRows: masterSheet ? masterSheet.getLastRow() - 1 : 0,
    shiftDataRows: shiftSheet ? shiftSheet.getLastRow() - 1 : 0,
    sendLogRows: logSheet ? logSheet.getLastRow() - 1 : 0,
    lineTokenSet: !!props.getProperty('LINE_CHANNEL_ACCESS_TOKEN'),
    batchInProgress: props.getProperty('SHIFT_BATCH_IN_PROGRESS') === 'true'
  });
}

/**
 * シートにデータを書き込む（GETパラメータ経由）
 * ?action=write&sheet=XX&range=A2&data=[[...]]
 */
function handleGetWrite_(params) {
  const sheetName = params.sheet;
  const range = params.range;
  const dataJson = params.data;

  if (!sheetName || !range || !dataJson) {
    return jsonResponse_({ error: 'sheet, range, data parameters required' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse_({ error: 'Sheet not found: ' + sheetName });

  const data = JSON.parse(dataJson);
  sheet.getRange(range).setValues(data);

  return jsonResponse_({ success: true, sheet: sheetName, range: range, rowsWritten: data.length });
}

/**
 * 列を書式なしテキストに設定してデータを再書き込み
 * ?action=formatText&sheet=XX&col=A
 */
function handleFormatText_(params) {
  const sheetName = params.sheet;
  const col = params.col || 'A';

  if (!sheetName) return jsonResponse_({ error: 'sheet parameter required' });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse_({ error: 'Sheet not found: ' + sheetName });

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return jsonResponse_({ error: 'No data rows' });

  // 列のフォーマットを書式なしテキストに設定
  const colIndex = col.charCodeAt(0) - 64; // A=1, B=2...
  const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
  range.setNumberFormat('@'); // テキスト形式

  // 現在の値を3桁ゼロパディングで再設定
  const values = range.getValues();
  const paddedValues = values.map(function(row) {
    const val = String(row[0]).trim();
    if (val && !isNaN(val)) {
      return [val.padStart(3, '0')];
    }
    return [val];
  });
  range.setValues(paddedValues);

  return jsonResponse_({ success: true, sheet: sheetName, column: col, rowsFixed: paddedValues.length });
}

/**
 * GAS関数をリモート実行
 * ?action=runFunction&name=setupSheets
 */
function handleRunFunction_(params) {
  const funcName = params.name;
  if (!funcName) return jsonResponse_({ error: 'name parameter required' });

  // 許可された関数のみ実行
  const allowedFunctions = {
    'setupSheets': setupSheets,
    'clearShiftDataForMonth': clearShiftDataForMonth,
    'runAllTests': typeof runAllTests === 'function' ? runAllTests : null
  };

  const func = allowedFunctions[funcName];
  if (!func) return jsonResponse_({ error: 'Function not allowed: ' + funcName });

  try {
    var arg = params.arg || null;
    var result = arg ? func(arg) : func();
    return jsonResponse_({ success: true, function: funcName, result: result || null });
  } catch (e) {
    return jsonResponse_({ error: e.toString(), function: funcName });
  }
}

/**
 * 外部スプレッドシート読み取り
 * ?action=readExternal&ssId=XXXXX&sheet=シート名
 * ?action=readExternal&ssId=XXXXX (シート一覧)
 */
function handleReadExternal_(params) {
  const ssId = params.ssId;
  if (!ssId) return jsonResponse_({ error: 'ssId parameter required' });

  try {
    const ss = SpreadsheetApp.openById(ssId);

    // シート名指定なし → シート一覧
    if (!params.sheet) {
      const sheets = ss.getSheets().map(function(s) {
        return { name: s.getName(), rows: s.getLastRow(), cols: s.getLastColumn() };
      });
      return jsonResponse_({ ssId: ssId, title: ss.getName(), sheets: sheets });
    }

    // シート名指定あり → データ読み取り
    const sheet = ss.getSheetByName(params.sheet);
    if (!sheet) {
      // シート名が見つからない場合、gidで探す
      if (params.gid) {
        const allSheets = ss.getSheets();
        for (var i = 0; i < allSheets.length; i++) {
          if (String(allSheets[i].getSheetId()) === params.gid) {
            var gidSheet = allSheets[i];
            var data = params.range ? gidSheet.getRange(params.range).getValues() : gidSheet.getDataRange().getValues();
            return jsonResponse_({ ssId: ssId, sheet: gidSheet.getName(), gid: params.gid, rows: data.length, cols: data[0] ? data[0].length : 0, data: data });
          }
        }
      }
      return jsonResponse_({ error: 'Sheet not found: ' + params.sheet });
    }

    var data;
    if (params.range) {
      data = sheet.getRange(params.range).getValues();
    } else {
      if (sheet.getLastRow() === 0) {
        data = [];
      } else {
        data = sheet.getDataRange().getValues();
      }
    }

    return jsonResponse_({ ssId: ssId, sheet: params.sheet, rows: data.length, cols: data[0] ? data[0].length : 0, data: data });
  } catch (e) {
    return jsonResponse_({ error: 'Cannot open spreadsheet: ' + e.toString() });
  }
}

/**
 * JSON レスポンスヘルパー
 */
function jsonResponse_(obj, statusCode) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST Admin API ルーティング
 * @param {Object} body - リクエストボディ
 * @return {ContentService} レスポンス
 */
function handlePostAction_(body) {
  switch (body.action) {
    case 'write':
      return handlePostWrite_(body);
    default:
      return jsonResponse_({ error: 'Unknown action: ' + body.action });
  }
}

/**
 * POST経由のシート書き込み（大量データ対応）
 * @param {Object} body - {sheet, range, data}
 * @return {ContentService} レスポンス
 */
function handlePostWrite_(body) {
  var sheetName = body.sheet;
  var range = body.range;
  var data = body.data;

  if (!sheetName || !range || !data) {
    return jsonResponse_({ error: 'sheet, range, data parameters required' });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse_({ error: 'Sheet not found: ' + sheetName });

  sheet.getRange(range).setValues(data);
  return jsonResponse_({ success: true, sheet: sheetName, range: range, rowsWritten: data.length });
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

    // 署名が無い場合
    if (!signature) {
      // Script Properties の SKIP_SIGNATURE_VERIFY が 'true' の場合のみ通す（開発用）
      const skipVerify = PropertiesService.getScriptProperties().getProperty('SKIP_SIGNATURE_VERIFY');
      if (skipVerify === 'true') {
        Logger.log('Warning: Signature verification skipped (dev mode)');
        return true;
      }
      Logger.log('Rejected: No signature provided');
      return false;
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
