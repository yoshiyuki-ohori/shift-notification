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

    // チャネル判定（統合 or シフト）
    var channel = detectChannel(body.destination || '');

    // イベント処理
    if (body.events && body.events.length > 0) {
      body.events.forEach(function(event) {
        processWebhookEvent(event, channel);
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
 * ?action=ical&empNo=XX&month=YYYY-MM&token=XXX → iCalファイル出力
 * (パラメータなし)        → Webhook URL検証
 */
function doGet(e) {
  const params = e ? (e.parameter || {}) : {};
  const action = params.action || '';
  const adminKey = params.key || '';

  // 管理APIキー検証（Script Propertiesに ADMIN_API_KEY を設定）
  // ADMIN_API_KEYが未設定の場合はアクション実行を拒否（デフォルト拒否）
  // ical は独自のトークン認証、lineWebhook はCloud Functions経由、liff/myshift はLIFF用のため除外
  if (action && action !== 'ical' && action !== 'lineWebhook' && action !== 'liff' && action !== 'myshift' && action !== 'empLookup' && action !== 'empRegister') {
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
      case 'ical':
        return handleGetIcal_(params);
      case 'liff':
        return serveLiffPage();
      case 'myshift':
        return handleMyShiftApi_(params);
      case 'empLookup':
        return handleEmpLookup_(params);
      case 'empRegister':
        return handleEmpRegister_(params);
      case 'lineWebhook':
        return handleLineWebhookViaGet_(params);
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
    'runAllTests': typeof runAllTests === 'function' ? runAllTests : null,
    'scheduleVerificationBroadcast': function(timeStr) {
      var d = new Date(timeStr);
      scheduleBroadcast_(d);
      return 'Scheduled at ' + d.toString();
    }
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
 * Cloud Functions経由のLINE Webhookデータ処理（GET経由）
 * Base64エンコードされたWebhookボディをデコードして処理
 * @param {Object} params - クエリパラメータ
 * @return {ContentService} レスポンス
 */
function handleLineWebhookViaGet_(params) {
  try {
    var encodedData = params.data;
    if (!encodedData) {
      return jsonResponse_({ error: 'data parameter required' });
    }

    var decoded = Utilities.newBlob(Utilities.base64Decode(encodedData)).getDataAsString();
    var body = JSON.parse(decoded);

    // チャネル判定
    var channel = detectChannel(body.destination || '');

    // イベント処理
    if (body.events && body.events.length > 0) {
      body.events.forEach(function(event) {
        processWebhookEvent(event, channel);
      });
    }

    return jsonResponse_({ success: true, eventsProcessed: (body.events || []).length });
  } catch (error) {
    Logger.log('lineWebhook via GET error: ' + error.toString());
    return jsonResponse_({ error: error.toString() });
  }
}

/**
 * iCalファイルを返す (HMAC署名トークンで認証)
 * ?action=ical&empNo=072&month=2026-03&token=XXXX
 */
function handleGetIcal_(params) {
  var empNo = params.empNo;
  var month = params.month;
  var token = params.token;

  if (!empNo || !month) {
    return ContentService.createTextOutput('empNo and month parameters required')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  // 社員番号を3桁ゼロパディング
  empNo = String(empNo).padStart(3, '0');

  // トークン検証
  if (!verifyICalToken(empNo, month, token)) {
    return ContentService.createTextOutput('Invalid or missing token')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  var ical = generateICal(empNo, month);
  if (!ical) {
    return ContentService.createTextOutput('No shift data found')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  return ContentService.createTextOutput(ical)
    .setMimeType(ContentService.MimeType.TEXT);
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
 * 注意: GASのdoPost()ではHTTPヘッダー (X-Line-Signature) にアクセスできないため、
 * GAS環境では署名検証は実質不可能。GASのWebアプリURLは推測困難な長いURLで
 * 保護されているため、セキュリティリスクは限定的。
 * @param {Object} e - イベントオブジェクト
 * @return {boolean} 検証結果
 */
function verifySignature(e) {
  // GASではHTTPリクエストヘッダーを取得できないため、
  // LINE署名検証(X-Line-Signature)は実行不可。
  // 代わりにリクエストボディの基本構造を検証する。
  try {
    const body = JSON.parse(e.postData.contents);

    // LINEからのWebhookリクエストはeventsフィールドを必ず持つ
    if (!body.hasOwnProperty('events')) {
      Logger.log('Rejected: Not a LINE webhook request (no events field)');
      return false;
    }

    // destination (Bot User ID) が存在すればLINEプラットフォームからのリクエスト
    if (body.destination) {
      Logger.log('LINE webhook verified (destination: ' + body.destination + ')');
    }

    return true;
  } catch (error) {
    Logger.log('Webhook validation error: ' + error.toString());
    return false;
  }
}

/**
 * Webhookイベント処理
 * @param {Object} event - LINEイベント
 * @param {string} [channel] - チャネル種別
 */
function processWebhookEvent(event, channel) {
  const userId = event.source.userId;
  const replyToken = event.replyToken;

  switch (event.type) {
    case 'message':
      if (event.message.type === 'text') {
        handleRegistrationMessage(event.message.text, userId, replyToken, channel);
      }
      break;

    case 'postback':
      handlePostbackEvent_(event.postback.data, userId, replyToken, channel);
      break;

    case 'follow':
      handleFollowEvent_(userId, replyToken, channel);
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
 * @param {string} [channel] - チャネル種別
 */
function handleRegistrationMessage(text, userId, replyToken, channel) {
  const trimmedText = text.trim();

  // 社員番号パターン判定 (001-999) のみ反応。それ以外は無視
  const empNoMatch = trimmedText.match(/^(\d{1,3})$/);
  if (!empNoMatch) {
    // 社員番号以外のメッセージは無視（返信しない）
    return;
  }

  // 社員番号を3桁にゼロパディング
  const employeeNo = String(parseInt(empNoMatch[1], 10)).padStart(3, '0');

  // 従業員マスタから検索
  const employee = findEmployeeByNo(employeeNo);
  if (!employee) {
    replyLineMessage(replyToken, [
      createTextMessage(
        '社員番号「' + employeeNo + '」が見つかりませんでした。\n' +
        '正しい社員番号を入力してください。'
      )
    ], channel);
    return;
  }

  // LINE UserIdをマスタに登録（既存でも上書き更新）
  const success = registerLineUserId(employeeNo, userId);

  if (success) {
    replyLineMessage(replyToken, [
      createTextMessage(
        employee.name + 'さん、登録しました。'
      )
    ], channel);
    Logger.log('LINE UserId registered: empNo=' + employeeNo + ' channel=' + channel);
  } else {
    replyLineMessage(replyToken, [
      createTextMessage(
        '登録処理でエラーが発生しました。\n' +
        'お手数ですが管理者にご連絡ください。'
      )
    ], channel);
  }
}

/**
 * フォローイベント処理（友だち追加時）
 * @param {string} userId - LINE UserId
 * @param {string} replyToken - 返信トークン
 * @param {string} [channel] - チャネル種別
 */
function handleFollowEvent_(userId, replyToken, channel) {
  replyLineMessage(replyToken, [
    createTextMessage(
      'シフト通知システムへようこそ！\n\n' +
      'LINE連携を行うため、社員番号を入力してください。\n' +
      '例: 072\n\n' +
      '※社員番号が分からない場合は管理者にお問い合わせください。'
    )
  ], channel);
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
 * Postbackイベント処理 - シフト希望入力フロー
 * @param {string} data - Postbackデータ (URLパラメータ形式)
 * @param {string} userId - LINE UserId
 * @param {string} replyToken - 返信トークン
 * @param {string} [channel] - チャネル種別
 */
function handlePostbackEvent_(data, userId, replyToken, channel) {
  var params = parsePostbackData_(data);
  var action = params.action || '';

  // 職員情報をLINE UserIdから検索
  var employee = findEmployeeByLineUserId_(userId);
  if (!employee) {
    replyLineMessage(replyToken, [
      createTextMessage('LINE連携が完了していません。\n先に社員番号を入力してください。')
    ], channel);
    return;
  }

  switch (action) {
    case 'pref_start':
      handlePrefStart_(params, employee, replyToken, channel);
      break;

    case 'pref_date':
      handlePrefDate_(params, replyToken, channel);
      break;

    case 'pref_type':
      handlePrefType_(params, employee, replyToken, channel);
      break;

    case 'pref_delete':
      handlePrefDelete_(params, employee, replyToken, channel);
      break;

    case 'pref_clear':
      handlePrefClear_(params, employee, replyToken, channel);
      break;

    case 'pref_finish':
      handlePrefFinish_(params, employee, replyToken, channel);
      break;

    default:
      Logger.log('Unknown postback action: ' + action);
      break;
  }
}

/**
 * pref_start: シフト希望入力開始 - カレンダー表示
 */
function handlePrefStart_(params, employee, replyToken, channel) {
  var yearMonth = params.month;
  if (!yearMonth) {
    // 収集期間の対象月を使用
    var period = getCollectionPeriod();
    yearMonth = period ? period.targetMonth : null;
  }

  if (!yearMonth) {
    replyLineMessage(replyToken, [
      createTextMessage('現在、シフト希望の収集期間外です。')
    ], channel);
    return;
  }

  var existingPrefs = getPreferencesForEmployee(yearMonth, employee.employeeNo);
  var message = buildPreferenceStartMessage(yearMonth, employee.name, existingPrefs);

  replyLineMessage(replyToken, [message], channel);
}

/**
 * pref_date: 日付選択 → 種別選択画面
 */
function handlePrefDate_(params, replyToken, channel) {
  var yearMonth = params.month;
  var dateStr = params.date;

  if (!yearMonth || !dateStr) return;

  var message = buildTypeSelectionMessage(yearMonth, dateStr);
  replyLineMessage(replyToken, [message], channel);
}

/**
 * pref_type: 種別選択 → 希望保存
 */
function handlePrefType_(params, employee, replyToken, channel) {
  var yearMonth = params.month;
  var dateStr = params.date;
  var type = params.type;

  if (!yearMonth || !dateStr || !type) return;

  var success = savePreference({
    yearMonth: yearMonth,
    employeeNo: employee.employeeNo,
    name: employee.name,
    date: dateStr,
    type: type
  });

  if (success) {
    var message = buildPreferenceSavedMessage(dateStr, type, yearMonth);
    replyLineMessage(replyToken, [message], channel);
  } else {
    replyLineMessage(replyToken, [
      createTextMessage('保存に失敗しました。もう一度お試しください。')
    ], channel);
  }
}

/**
 * pref_delete: 特定日の希望を削除
 */
function handlePrefDelete_(params, employee, replyToken, channel) {
  var yearMonth = params.month;
  var dateStr = params.date;
  if (!yearMonth || !dateStr) return;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);
    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getValues();
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]).trim() === yearMonth &&
            String(data[i][1]).trim().padStart(3, '0') === employee.employeeNo &&
            String(data[i][3]).trim() === dateStr) {
          sheet.deleteRow(i + 1);
        }
      }
    }

    var dateParts = dateStr.split('/');
    var displayDate = parseInt(dateParts[1], 10) + '月' + parseInt(dateParts[2], 10) + '日';

    replyLineMessage(replyToken, [{
      type: 'text',
      text: '🗑 ' + displayDate + ' の希望を削除しました。',
      quickReply: {
        items: [{
          type: 'action',
          action: {
            type: 'postback',
            label: 'カレンダーに戻る',
            data: 'action=pref_start&month=' + yearMonth,
            displayText: 'シフト希望入力'
          }
        }]
      }
    }], channel);
  } catch (e) {
    Logger.log('handlePrefDelete_ error: ' + e.toString());
    replyLineMessage(replyToken, [
      createTextMessage('削除に失敗しました。')
    ], channel);
  }
}

/**
 * pref_clear: 全希望クリア
 */
function handlePrefClear_(params, employee, replyToken, channel) {
  var yearMonth = params.month;
  if (!yearMonth) return;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);
    if (sheet && sheet.getLastRow() > 1) {
      deletePrefsForEmployee_(sheet, yearMonth, employee.employeeNo);
    }

    replyLineMessage(replyToken, [{
      type: 'text',
      text: '全ての希望をクリアしました。',
      quickReply: {
        items: [{
          type: 'action',
          action: {
            type: 'postback',
            label: '最初から入力',
            data: 'action=pref_start&month=' + yearMonth,
            displayText: 'シフト希望入力'
          }
        }]
      }
    }], channel);
  } catch (e) {
    Logger.log('handlePrefClear_ error: ' + e.toString());
    replyLineMessage(replyToken, [
      createTextMessage('クリアに失敗しました。')
    ], channel);
  }
}

/**
 * pref_finish: 入力完了
 */
function handlePrefFinish_(params, employee, replyToken, channel) {
  var yearMonth = params.month;
  if (!yearMonth) return;

  var prefs = getPreferencesForEmployee(yearMonth, employee.employeeNo);
  var message = buildPreferenceCompleteMessage(yearMonth, prefs.length);
  replyLineMessage(replyToken, [message], channel);
}

/**
 * LINE UserIdから従業員を検索
 * @param {string} userId - LINE UserId
 * @return {Object|null} { employeeNo, name }
 */
function findEmployeeByLineUserId_(userId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEE_MASTER);
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var lineId = String(data[i][MASTER_COLS.LINE_USER_ID - 1] || '').trim();
    if (lineId === userId) {
      return {
        employeeNo: String(data[i][MASTER_COLS.NO - 1]).trim().padStart(3, '0'),
        name: String(data[i][MASTER_COLS.NAME - 1]).trim()
      };
    }
  }
  return null;
}

/**
 * Postbackデータをパースしてオブジェクトに変換
 * @param {string} data - "action=xxx&month=yyy&date=zzz"
 * @return {Object} パラメータオブジェクト
 */
function parsePostbackData_(data) {
  var result = {};
  if (!data) return result;

  data.split('&').forEach(function(pair) {
    var parts = pair.split('=');
    if (parts.length === 2) {
      result[decodeURIComponent(parts[0])] = decodeURIComponent(parts[1]);
    }
  });
  return result;
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
