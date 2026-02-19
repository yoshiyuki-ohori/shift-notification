/**
 * LiffApp.gs - LIFF マイシフトビューア サーバー側処理
 * LINE Front-end Framework で職員が自分のシフトを確認する機能
 *
 * フロー:
 * 1. LIFF → GitHub Pages の HTML を開く
 * 2. クライアント: LIFF SDK init → accessToken 取得
 * 3. クライアント: fetch(GAS_URL?action=myshift&token=xxx&month=xxx)
 * 4. サーバー: LINE API で accessToken 検証 → userId → 社員番号 → データ返却
 */

/**
 * マイシフト JSON API エンドポイント (doGet から呼ばれる)
 * @param {Object} params - クエリパラメータ { token, month }
 * @return {ContentService.TextOutput} JSON レスポンス
 */
function handleMyShiftApi_(params) {
  var token = params.token || '';
  var month = params.month || '';
  var empNo = params.empNo || '';
  var data = getMyShiftData(token, month, empNo);
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * クライアントから呼ばれるメイン API
 * LINE アクセストークンを検証し、該当職員のシフトデータを返す
 * @param {string} accessToken - LIFF から取得したアクセストークン
 * @param {string} [targetMonth] - 対象年月 (YYYY-MM)。省略時は現在月
 * @param {string} [overrideEmpNo] - 管理者テスト用: 指定社員番号のデータを表示
 * @return {Object} シフトデータ
 */
function getMyShiftData(accessToken, targetMonth, overrideEmpNo) {
  try {
    // 1. アクセストークンからLINE userId を取得（サーバー側検証）
    var userId = verifyLineAccessToken_(accessToken);
    if (!userId) {
      return { error: 'LINE認証に失敗しました。再度お試しください。' };
    }

    // 2. userId から社員番号を取得
    var employee = findEmployeeByLineUserId_(userId);
    if (!employee) {
      return { needsRegistration: true };
    }

    // 2b. 管理者テスト: empNo 指定があれば該当者のデータを表示
    if (overrideEmpNo) {
      var targetEmpNo = overrideEmpNo.padStart(3, '0');
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var masterSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEE_MASTER);
      if (masterSheet) {
        var data = masterSheet.getDataRange().getValues();
        for (var idx = 1; idx < data.length; idx++) {
          if (String(data[idx][MASTER_COLS.NO - 1]).trim().padStart(3, '0') === targetEmpNo) {
            employee = {
              employeeNo: targetEmpNo,
              name: String(data[idx][MASTER_COLS.NAME - 1]).trim()
            };
            break;
          }
        }
      }
    }

    // 3. 対象年月の決定
    if (!targetMonth) {
      var now = new Date();
      var y = now.getFullYear();
      var m = ('0' + (now.getMonth() + 1)).slice(-2);
      targetMonth = y + '-' + m;
    }

    // 4. シフトデータ取得（全員分から該当者を抽出）
    var allAggregated = aggregateByEmployee(targetMonth);
    var myShifts = [];
    for (var i = 0; i < allAggregated.length; i++) {
      if (allAggregated[i].employeeNo === employee.employeeNo) {
        myShifts = allAggregated[i].shifts;
        break;
      }
    }

    // 5. 希望データ取得
    var preferences = getPreferencesForEmployee(targetMonth, employee.employeeNo);

    // 6. 希望収集期間情報
    var collectionPeriod = getCollectionPeriod();
    var collectionOpen = isCollectionOpen();

    // 7. 統計サマリー
    var facilitySet = {};
    for (var j = 0; j < myShifts.length; j++) {
      facilitySet[myShifts[j].facility] = true;
    }

    return {
      employee: {
        name: employee.name,
        employeeNo: employee.employeeNo
      },
      targetMonth: targetMonth,
      shifts: myShifts,
      preferences: preferences,
      collectionPeriod: collectionPeriod,
      collectionOpen: collectionOpen,
      stats: {
        totalShifts: myShifts.length,
        facilityCount: Object.keys(facilitySet).length
      }
    };
  } catch (e) {
    Logger.log('getMyShiftData error: ' + e.toString());
    return { error: 'データの取得中にエラーが発生しました。' };
  }
}

/**
 * 社員番号で従業員を検索 (登録前の確認用)
 * @param {Object} params - { token, empNo }
 * @return {ContentService.TextOutput} JSON
 */
function handleEmpLookup_(params) {
  var token = params.token || '';
  var empNo = params.empNo || '';

  var userId = verifyLineAccessToken_(token);
  if (!userId) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'LINE認証に失敗しました。' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (!empNo) {
    return ContentService.createTextOutput(JSON.stringify({ error: '社員番号を入力してください。' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var targetEmpNo = empNo.padStart(3, '0');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEE_MASTER);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ error: '従業員マスタが見つかりません。' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var no = String(data[i][MASTER_COLS.NO - 1]).trim().padStart(3, '0');
    if (no === targetEmpNo) {
      var existingLineId = String(data[i][MASTER_COLS.LINE_USER_ID - 1] || '').trim();
      if (existingLineId && existingLineId !== userId) {
        return ContentService.createTextOutput(JSON.stringify({
          error: 'この社員番号は既に別のLINEアカウントで登録されています。'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({
        found: true,
        empNo: targetEmpNo,
        name: String(data[i][MASTER_COLS.NAME - 1]).trim()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    error: '社員番号「' + empNo + '」は見つかりませんでした。\n正しい番号を入力してください。'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * LINE userId を従業員マスタに登録
 * @param {Object} params - { token, empNo }
 * @return {ContentService.TextOutput} JSON
 */
function handleEmpRegister_(params) {
  var token = params.token || '';
  var empNo = params.empNo || '';

  var userId = verifyLineAccessToken_(token);
  if (!userId) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'LINE認証に失敗しました。' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var targetEmpNo = empNo.padStart(3, '0');
  var success = registerLineUserId(targetEmpNo, userId);

  if (success) {
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } else {
    return ContentService.createTextOutput(JSON.stringify({ error: '登録に失敗しました。' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * LINE アクセストークンを検証し userId を返す
 * @param {string} accessToken - LIFF のアクセストークン
 * @return {string|null} LINE userId (検証失敗時は null)
 * @private
 */
function verifyLineAccessToken_(accessToken) {
  if (!accessToken) return null;

  try {
    var response = UrlFetchApp.fetch('https://api.line.me/v2/profile', {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + accessToken
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      Logger.log('LINE token verification failed: ' + response.getContentText());
      return null;
    }

    var profile = JSON.parse(response.getContentText());
    return profile.userId || null;
  } catch (e) {
    Logger.log('verifyLineAccessToken_ error: ' + e.toString());
    return null;
  }
}
