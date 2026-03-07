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
 * 管理者キー認証で社員情報を取得する共通ヘルパー
 * @param {string} adminKey - 管理者APIキー
 * @param {string} empNo - 社員番号
 * @return {Object|null} { employeeNo, name } または null（認証失敗/社員未発見）
 * @private
 */
function authenticateByAdminKey_(adminKey, empNo) {
  var expectedKey = PropertiesService.getScriptProperties().getProperty('ADMIN_API_KEY') || '';
  if (!expectedKey || adminKey !== expectedKey) {
    return null;
  }
  var targetEmpNo = empNo.padStart(3, '0');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEE_MASTER);
  if (!masterSheet) return null;

  var data = masterSheet.getDataRange().getValues();
  for (var idx = 1; idx < data.length; idx++) {
    if (String(data[idx][MASTER_COLS.NO - 1]).trim().padStart(3, '0') === targetEmpNo) {
      return {
        employeeNo: targetEmpNo,
        name: String(data[idx][MASTER_COLS.NAME - 1]).trim()
      };
    }
  }
  return null;
}

/**
 * マイシフト JSON API エンドポイント (doGet から呼ばれる)
 * @param {Object} params - クエリパラメータ { token, month }
 * @return {ContentService.TextOutput} JSON レスポンス
 */
function handleMyShiftApi_(params) {
  var token = params.token || '';
  var month = params.month || '';
  var empNo = params.empNo || '';
  var adminKey = params.key || '';
  var data = getMyShiftData(token, month, empNo, adminKey);
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
function getMyShiftData(accessToken, targetMonth, overrideEmpNo, adminKey) {
  try {
    var employee = null;

    // 管理者キー認証: LINE トークン不要で社員番号指定のデータを返す
    if (adminKey && overrideEmpNo) {
      employee = authenticateByAdminKey_(adminKey, overrideEmpNo);
      if (!employee) {
        return { error: '管理者認証に失敗、または社員番号「' + overrideEmpNo + '」が見つかりません。' };
      }
    } else {
      // 1. アクセストークンからLINE userId を取得（サーバー側検証）
      var userId = verifyLineAccessToken_(accessToken);
      if (!userId) {
        return { error: 'LINE認証に失敗しました。再度お試しください。' };
      }

      // 2. userId から社員番号を取得
      employee = findEmployeeByLineUserId_(userId);
      if (!employee) {
        return { needsRegistration: true };
      }

      // 2b. 管理者テスト: empNo 指定があれば該当者のデータを表示
      if (overrideEmpNo) {
        var targetEmpNo2 = overrideEmpNo.padStart(3, '0');
        var ss2 = SpreadsheetApp.getActiveSpreadsheet();
        var masterSheet2 = ss2.getSheetByName(SHEET_NAMES.EMPLOYEE_MASTER);
        if (masterSheet2) {
          var data2 = masterSheet2.getDataRange().getValues();
          for (var idx2 = 1; idx2 < data2.length; idx2++) {
            if (String(data2[idx2][MASTER_COLS.NO - 1]).trim().padStart(3, '0') === targetEmpNo2) {
              employee = {
                employeeNo: targetEmpNo2,
                name: String(data2[idx2][MASTER_COLS.NAME - 1]).trim()
              };
              break;
            }
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

    // 4b. 同一建物の同僚情報を追加
    myShifts = addCoworkerInfo_(myShifts, allAggregated, employee.employeeNo);

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

  // 全角数字→半角変換
  empNo = empNo.replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });

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
      var status = String(data[i][MASTER_COLS.STATUS - 1] || '在職').trim();
      if (status === '退職') {
        return ContentService.createTextOutput(JSON.stringify({
          error: 'この社員番号は現在使用できません。\n管理者にお問い合わせください。'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var existingLineId = String(data[i][MASTER_COLS.LINE_USER_ID - 1] || '').trim();
      var result = {
        found: true,
        empNo: targetEmpNo,
        name: String(data[i][MASTER_COLS.NAME - 1]).trim()
      };
      if (existingLineId && existingLineId !== userId) {
        result.reRegister = true;
        result.warning = 'この社員番号は別の端末で登録済みです。\nこの端末で再登録すると、以前の端末ではシフトを確認できなくなります。';
      }
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
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
 * 全施設配置状況 JSON API エンドポイント (doGet から呼ばれる)
 * ADMIN_API_KEY 認証済み前提 (doGet で検証済み)
 * @param {Object} params - クエリパラメータ { month }
 * @return {ContentService.TextOutput} JSON レスポンス
 */
function handleFacilityOverview_(params) {
  try {
    // 対象年月の決定
    var targetMonth = params.month || '';
    if (!targetMonth) {
      var now = new Date();
      targetMonth = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2);
    }

    // シフトデータ全行取得
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shiftSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_DATA);
    if (!shiftSheet || shiftSheet.getLastRow() <= 1) {
      return jsonResponse_({
        targetMonth: targetMonth,
        facilities: [],
        summary: { totalFacilities: 0, totalShifts: 0, totalStaff: 0 }
      });
    }

    var shiftData = shiftSheet.getDataRange().getValues();

    // 施設×日付×時間帯でグルーピング
    var facilityMap = {};  // { facilityName: { dates: { day: [ shifts ] }, area: '', id: '' } }
    var allStaffSet = {};

    for (var i = 1; i < shiftData.length; i++) {
      var rawYM = shiftData[i][SHIFT_COLS.YEAR_MONTH - 1];
      var yearMonth;
      if (rawYM instanceof Date) {
        yearMonth = rawYM.getFullYear() + '-' + ('0' + (rawYM.getMonth() + 1)).slice(-2);
      } else {
        yearMonth = String(rawYM).trim();
      }
      if (yearMonth !== targetMonth) continue;

      var facility = String(shiftData[i][SHIFT_COLS.FACILITY - 1]).trim();
      var rawDate = shiftData[i][SHIFT_COLS.DATE - 1];
      var dateStr;
      if (rawDate instanceof Date) {
        dateStr = rawDate.getFullYear() + '/' + ('0' + (rawDate.getMonth() + 1)).slice(-2) + '/' + ('0' + rawDate.getDate()).slice(-2);
      } else {
        dateStr = String(rawDate).trim();
      }
      var area = String(shiftData[i][SHIFT_COLS.AREA - 1]).trim();
      var facilityId = getFacilityId(facility) || '';
      var timeSlot = String(shiftData[i][SHIFT_COLS.TIME_SLOT - 1]).trim();
      var empNo = String(shiftData[i][SHIFT_COLS.EMPLOYEE_NO - 1]).trim().padStart(3, '0');
      var empName = String(shiftData[i][SHIFT_COLS.FORMAL_NAME - 1]).trim();

      if (!facility) continue;

      // 日付から日を抽出 (例: "2026/03/01" → "1")
      var day = '';
      if (dateStr.indexOf('/') !== -1) {
        var parts = dateStr.split('/');
        day = String(parseInt(parts[parts.length - 1], 10));
      } else if (rawDate instanceof Date) {
        day = String(rawDate.getDate());
      } else {
        day = dateStr;
      }

      if (!facilityMap[facility]) {
        facilityMap[facility] = { dates: {}, area: area, id: facilityId, staffSet: {} };
      }
      if (!facilityMap[facility].dates[day]) {
        facilityMap[facility].dates[day] = [];
      }

      facilityMap[facility].dates[day].push({
        timeSlot: timeSlot,
        empNo: empNo,
        name: empName
      });

      if (empNo) {
        facilityMap[facility].staffSet[empNo] = true;
        allStaffSet[empNo] = true;
      }
    }

    // 施設一覧を配列に変換
    var facilities = [];
    var totalShifts = 0;

    var facilityNames = Object.keys(facilityMap).sort();
    for (var f = 0; f < facilityNames.length; f++) {
      var fname = facilityNames[f];
      var fdata = facilityMap[fname];
      var facilityShiftCount = 0;

      // 各日の時間帯をソート
      var sortedDates = {};
      var days = Object.keys(fdata.dates);
      for (var d = 0; d < days.length; d++) {
        var dayShifts = fdata.dates[days[d]];
        dayShifts.sort(function(a, b) {
          return getTimeSlotOrder(a.timeSlot) - getTimeSlotOrder(b.timeSlot);
        });
        sortedDates[days[d]] = dayShifts;
        facilityShiftCount += dayShifts.length;
      }

      totalShifts += facilityShiftCount;

      facilities.push({
        name: fname,
        id: fdata.id,
        area: fdata.area,
        dates: sortedDates,
        stats: {
          totalShifts: facilityShiftCount,
          uniqueStaff: Object.keys(fdata.staffSet).length
        }
      });
    }

    return jsonResponse_({
      targetMonth: targetMonth,
      facilities: facilities,
      summary: {
        totalFacilities: facilities.length,
        totalShifts: totalShifts,
        totalStaff: Object.keys(allStaffSet).length
      }
    });

  } catch (e) {
    Logger.log('handleFacilityOverview_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
  }
}

/**
 * シフト希望データ取得 API エンドポイント (doGet から呼ばれる)
 * ?action=prefData&token=xxx&month=YYYY-MM
 * @param {Object} params - クエリパラメータ { token, month }
 * @return {ContentService.TextOutput} JSON レスポンス
 */
function handlePrefDataApi_(params) {
  var token = params.token || '';
  var month = params.month || '';
  var adminKey = params.key || '';
  var empNo = params.empNo || '';

  var employee = null;

  // 管理者キー認証: LIFF不要で社員番号指定
  if (adminKey && empNo) {
    employee = authenticateByAdminKey_(adminKey, empNo);
    if (!employee) {
      return ContentService.createTextOutput(JSON.stringify({ error: '管理者認証に失敗、または社員番号「' + empNo + '」が見つかりません。' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } else {
    var userId = verifyLineAccessToken_(token);
    if (!userId) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'LINE認証に失敗しました。' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    employee = findEmployeeByLineUserId_(userId);
    if (!employee) {
      return ContentService.createTextOutput(JSON.stringify({ needsRegistration: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 対象年月の決定
  if (!month) {
    var period = getCollectionPeriod();
    month = period ? period.targetMonth : null;
    if (!month) {
      var now = new Date();
      month = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2);
    }
  }

  try {
    var preferences = getPreferencesForEmployee(month, employee.employeeNo);
    var collectionPeriod = getCollectionPeriod();
    var collectionOpen = isCollectionOpen();
    var submittedAt = getSubmittedAt(employee.employeeNo, month);

    // 期間区分と日付範囲を算出
    var periodLabel = collectionPeriod ? (collectionPeriod.periodLabel || '全日') : '全日';
    var monthParts = month.split('-');
    var daysInMonth = new Date(parseInt(monthParts[0], 10), parseInt(monthParts[1], 10), 0).getDate();
    var dateRange = {
      startDay: periodLabel === '後半' ? 16 : 1,
      endDay: periodLabel === '前半' ? 15 : daysInMonth
    };

    var facilities = getFacilityList_();

    var result = {
      employee: {
        name: employee.name,
        employeeNo: employee.employeeNo
      },
      targetMonth: month,
      preferences: preferences,
      collectionPeriod: collectionPeriod,
      collectionOpen: collectionOpen,
      submittedAt: submittedAt,
      periodLabel: periodLabel,
      dateRange: dateRange,
      facilities: facilities
    };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    Logger.log('handlePrefDataApi_ error: ' + e.toString());
    return ContentService.createTextOutput(JSON.stringify({ error: 'データ取得中にエラーが発生しました。' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * シフト希望一括保存 API エンドポイント (doPost から呼ばれる)
 * @param {Object} body - リクエストボディ { action, token, month, preferences }
 * @return {ContentService.TextOutput} JSON レスポンス
 */
function handleSavePrefBatch_(body) {
  var token = body.token || '';
  var month = body.month || '';
  var preferences = body.preferences || [];
  var adminKey = body.key || '';
  var empNo = body.empNo || '';

  var employee = null;

  // 管理者キー認証: LIFF不要で社員番号指定
  if (adminKey && empNo) {
    employee = authenticateByAdminKey_(adminKey, empNo);
    if (!employee) {
      return ContentService.createTextOutput(JSON.stringify({ error: '管理者認証に失敗、または社員番号「' + empNo + '」が見つかりません。' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } else {
    var userId = verifyLineAccessToken_(token);
    if (!userId) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'LINE認証に失敗しました。' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    employee = findEmployeeByLineUserId_(userId);
    if (!employee) {
      return ContentService.createTextOutput(JSON.stringify({ error: '従業員情報が見つかりません。' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (!month) {
    return ContentService.createTextOutput(JSON.stringify({ error: '対象年月が指定されていません。' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    // 希望データを整形
    var prefs = preferences.map(function(p) {
      return {
        yearMonth: month,
        employeeNo: employee.employeeNo,
        name: employee.name,
        date: p.date,
        type: p.type,
        timeSlot: p.timeSlot || '',
        reason: p.reason || '',
        facility: p.facility || ''
      };
    });

    // 一括保存（既存データは削除して上書き）
    var savedCount = 0;
    if (prefs.length > 0) {
      savedCount = savePreferences(prefs);
    } else {
      // 空の場合は既存データを削除
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);
      if (sheet && sheet.getLastRow() > 1) {
        deletePrefsForEmployee_(sheet, month, employee.employeeNo);
      }
    }

    // 提出済みマーク
    var submittedAt = markAsSubmitted(employee.employeeNo, month);

    // 種別ごとのカウント
    var wantCount = 0, ngCount = 0;
    prefs.forEach(function(p) {
      if (p.type === PREF_TYPE.WANT) wantCount++;
      else if (p.type === PREF_TYPE.NG) ngCount++;
    });

    // 管理者通知
    notifyAdminOnSubmission_(employee.name, month, wantCount, ngCount);

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      savedCount: savedCount,
      submittedAt: submittedAt
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    Logger.log('handleSavePrefBatch_ error: ' + e.toString());
    return ContentService.createTextOutput(JSON.stringify({ error: '保存中にエラーが発生しました。' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 全職員のシフト希望一覧 API エンドポイント (doGet から呼ばれる)
 * ADMIN_API_KEY 認証済み前提
 * ?action=allPreferences&month=YYYY-MM
 * @param {Object} params - クエリパラメータ
 * @return {ContentService.TextOutput} JSON レスポンス
 */
function handleAllPreferencesApi_(params) {
  try {
    var targetMonth = params.month || '';
    if (!targetMonth) {
      var now = new Date();
      targetMonth = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // シフト希望データ取得
    var prefSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);
    var preferences = [];
    if (prefSheet && prefSheet.getLastRow() > 1) {
      var prefData = prefSheet.getDataRange().getValues();
      for (var i = 1; i < prefData.length; i++) {
        var ym = formatYearMonth_(prefData[i][PREF_COLS.YEAR_MONTH - 1]);
        if (ym !== targetMonth) continue;

        preferences.push({
          employeeNo: String(prefData[i][PREF_COLS.EMPLOYEE_NO - 1]).trim().padStart(3, '0'),
          name: String(prefData[i][PREF_COLS.NAME - 1]).trim(),
          date: formatDateValue_(prefData[i][PREF_COLS.DATE - 1]),
          type: String(prefData[i][PREF_COLS.TYPE - 1]).trim(),
          timeSlot: String(prefData[i][PREF_COLS.TIME_SLOT - 1]).trim(),
          reason: String(prefData[i][PREF_COLS.REASON - 1] || '').trim(),
          facility: String(prefData[i][PREF_COLS.FACILITY - 1] || '').trim()
        });
      }
    }

    // 日付×人でグループ化
    var byDate = {};
    for (var j = 0; j < preferences.length; j++) {
      var p = preferences[j];
      if (!byDate[p.date]) byDate[p.date] = [];
      byDate[p.date].push(p);
    }

    // 提出状況を取得
    var employees = loadEmployeeMaster();
    var submittedSet = {};
    preferences.forEach(function(p) { submittedSet[p.employeeNo] = true; });

    var submitted = [];
    var notSubmitted = [];
    employees.forEach(function(emp) {
      if (emp.status !== '在職') return;
      if (submittedSet[emp.employeeNo]) {
        submitted.push({ employeeNo: emp.employeeNo, name: emp.name });
      } else {
        notSubmitted.push({ employeeNo: emp.employeeNo, name: emp.name });
      }
    });

    return jsonResponse_({
      targetMonth: targetMonth,
      preferences: preferences,
      byDate: byDate,
      submissionStatus: {
        submitted: submitted,
        notSubmitted: notSubmitted,
        submittedCount: submitted.length,
        notSubmittedCount: notSubmitted.length
      }
    });
  } catch (e) {
    Logger.log('handleAllPreferencesApi_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
  }
}

/**
 * コンプライアンスチェック API エンドポイント (doGet から呼ばれる)
 * ADMIN_API_KEY 認証済み前提
 * ?action=complianceCheck&month=YYYY-MM
 * @param {Object} params - クエリパラメータ
 * @return {ContentService.TextOutput} JSON レスポンス
 */
function handleComplianceCheck_(params) {
  try {
    var targetMonth = params.month || '';
    if (!targetMonth) {
      var now = new Date();
      targetMonth = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2);
    }

    var result = runComplianceCheck(targetMonth);
    result.targetMonth = targetMonth;

    return jsonResponse_(result);
  } catch (e) {
    Logger.log('handleComplianceCheck_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
  }
}

/**
 * 配置充足率チェック API エンドポイント (doGet から呼ばれる)
 * ADMIN_API_KEY 認証済み前提
 * ?action=staffingCheck&month=YYYY-MM
 * @param {Object} params - クエリパラメータ
 * @return {ContentService.TextOutput} JSON レスポンス
 */
function handleStaffingCheck_(params) {
  try {
    var targetMonth = params.month || '';
    if (!targetMonth) {
      var now = new Date();
      targetMonth = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2);
    }

    var result = checkStaffingLevels(targetMonth);

    return jsonResponse_(result);
  } catch (e) {
    Logger.log('handleStaffingCheck_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
  }
}

/**
 * 配置比率 API エンドポイント (doGet から呼ばれる)
 * ADMIN_API_KEY 認証済み前提
 * ?action=prefCoverage&month=YYYY-MM
 * @param {Object} params - クエリパラメータ
 * @return {ContentService.TextOutput} JSON レスポンス
 */
function handlePrefCoverage_(params) {
  try {
    var targetMonth = params.month || '';
    if (!targetMonth) {
      var now = new Date();
      targetMonth = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2);
    }

    var result = calculatePrefCoverage_(targetMonth);
    return jsonResponse_(result);
  } catch (e) {
    Logger.log('handlePrefCoverage_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
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

/**
 * LIFF ページを GAS HtmlService で返す
 * doGet の case 'liff' から呼ばれる
 * @return {HtmlService.HtmlOutput}
 */
function serveLiffPage() {
  return HtmlService.createHtmlOutputFromFile('LiffView')
    .setTitle('マイシフト')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * エリア別施設リストを取得（PREF_FACILITY_LIST をそのまま返す）
 * @return {Object} { "練馬区": ["施設A", ...], "世田谷区": [...], ... }
 * @private
 */
function getFacilityList_() {
  return PREF_FACILITY_LIST;
}

/**
 * 同一建物の同僚情報をシフトデータに付与
 * @param {Array} myShifts - 対象職員のシフト配列
 * @param {Array} allAggregated - 全職員の集約データ
 * @param {string} myEmpNo - 対象職員の社員番号
 * @return {Array} coworkers プロパティを追加したシフト配列
 * @private
 */
function addCoworkerInfo_(myShifts, allAggregated, myEmpNo) {
  // 建物グループ+日付 → [{empNo, name, facility, timeSlot}] のルックアップを構築
  var buildingShiftMap = {};

  for (var a = 0; a < allAggregated.length; a++) {
    var emp = allAggregated[a];
    for (var s = 0; s < emp.shifts.length; s++) {
      var shift = emp.shifts[s];
      var group = BUILDING_GROUPS[shift.facility];
      if (!group) continue;

      var key = group + '|' + shift.date;
      if (!buildingShiftMap[key]) buildingShiftMap[key] = [];
      buildingShiftMap[key].push({
        empNo: emp.employeeNo,
        name: emp.name,
        facility: shift.facility,
        timeSlot: shift.timeSlot
      });
    }
  }

  // 自分のシフトに同僚情報を付与
  for (var i = 0; i < myShifts.length; i++) {
    var myGroup = BUILDING_GROUPS[myShifts[i].facility];
    if (!myGroup) continue;

    var lookupKey = myGroup + '|' + myShifts[i].date;
    var workers = buildingShiftMap[lookupKey];
    if (!workers) continue;

    var coworkers = [];
    for (var w = 0; w < workers.length; w++) {
      if (workers[w].empNo === myEmpNo) continue;
      coworkers.push({
        name: workers[w].name,
        facility: workers[w].facility,
        timeSlot: workers[w].timeSlot
      });
    }

    if (coworkers.length > 0) {
      myShifts[i].coworkers = coworkers;
    }
  }

  return myShifts;
}

// ===== 配置管理 API ハンドラー =====

/**
 * 仮配置データ取得 (GET)
 * ?action=tentativeData&month=YYYY-MM
 */
function handleTentativeData_(params) {
  try {
    var targetMonth = params.month || '';
    if (!targetMonth) {
      var now = new Date();
      targetMonth = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2);
    }

    var assignments = loadTentativeAssignments(targetMonth);
    var stats = getTentativeStats(targetMonth);

    // 施設×日付×時間帯でグループ化
    var byFacility = {};
    for (var i = 0; i < assignments.length; i++) {
      var a = assignments[i];
      var key = a.facilityId || a.facility;
      if (!byFacility[key]) {
        byFacility[key] = { facility: a.facility, facilityId: a.facilityId, area: a.area, slots: {} };
      }
      var slotKey = a.date + '|' + a.timeSlot;
      if (!byFacility[key].slots[slotKey]) {
        byFacility[key].slots[slotKey] = [];
      }
      byFacility[key].slots[slotKey].push({
        empNo: a.empNo,
        name: a.name,
        prefMatch: a.prefMatch,
        source: a.source
      });
    }

    return jsonResponse_({
      targetMonth: targetMonth,
      assignments: assignments,
      byFacility: byFacility,
      stats: stats
    });
  } catch (e) {
    Logger.log('handleTentativeData_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
  }
}

/**
 * 配置可能職員リスト取得 (GET)
 * ?action=availableStaff&month=YYYY-MM&date=YYYY/MM/DD&timeSlot=...
 */
function handleAvailableStaff_(params) {
  try {
    var targetMonth = params.month || '';
    var date = params.date || '';
    var timeSlot = params.timeSlot || '';
    var facilityId = params.facilityId || '';

    if (!targetMonth || !date || !timeSlot) {
      return jsonResponse_({ error: 'month, date, timeSlot パラメータが必要です。' });
    }

    var candidates = getAvailableStaff(targetMonth, date, timeSlot, facilityId);

    return jsonResponse_({
      targetMonth: targetMonth,
      date: date,
      timeSlot: timeSlot,
      candidates: candidates
    });
  } catch (e) {
    Logger.log('handleAvailableStaff_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
  }
}

/**
 * 配置統計サマリー取得 (GET)
 * ?action=allocationSummary&month=YYYY-MM
 */
function handleAllocationSummary_(params) {
  try {
    var targetMonth = params.month || '';
    if (!targetMonth) {
      var now = new Date();
      targetMonth = now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2);
    }

    var stats = getTentativeStats(targetMonth);
    var requirements = loadStaffingRequirements_();

    // 施設ごとの必要配置情報
    var facReqs = {};
    for (var r = 0; r < requirements.length; r++) {
      var req = requirements[r];
      var key = req.facilityId + '|' + req.timeSlot + '|' + req.dayType;
      facReqs[key] = { minStaff: req.minStaff, preferredStaff: req.preferredStaff };
    }

    return jsonResponse_({
      targetMonth: targetMonth,
      stats: stats,
      requirements: facReqs
    });
  } catch (e) {
    Logger.log('handleAllocationSummary_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
  }
}

/**
 * 配置追加 (POST)
 * body: { action: 'addAssignment', key, month, date, facility, facilityId, area, timeSlot, empNo, name }
 */
function handleAddAssignment_(body) {
  try {
    var result = addTentativeAssignment({
      yearMonth: body.month || '',
      date: body.date || '',
      area: body.area || '',
      facility: body.facility || '',
      facilityId: body.facilityId || '',
      timeSlot: body.timeSlot || '',
      empNo: body.empNo || '',
      name: body.name || '',
      assignedBy: '管理者'
    });

    return jsonResponse_(result);
  } catch (e) {
    Logger.log('handleAddAssignment_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
  }
}

/**
 * 配置削除 (POST)
 * body: { action: 'removeAssignment', key, month, date, facilityId, timeSlot, empNo }
 */
function handleRemoveAssignment_(body) {
  try {
    var success = removeTentativeAssignment(
      body.month || '',
      body.date || '',
      body.facilityId || '',
      body.timeSlot || '',
      body.empNo || ''
    );
    return jsonResponse_({ success: success });
  } catch (e) {
    Logger.log('handleRemoveAssignment_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
  }
}

/**
 * 配置確定 (POST)
 * body: { action: 'confirmAllocations', key, month }
 */
function handleConfirmAllocations_(body) {
  try {
    var targetMonth = body.month || '';
    if (!targetMonth) {
      return jsonResponse_({ error: '対象年月が指定されていません。' });
    }

    var result = confirmAllocations(targetMonth);
    return jsonResponse_(result);
  } catch (e) {
    Logger.log('handleConfirmAllocations_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
  }
}

/**
 * 一括仮配置追加 (POST)
 * body: { action: 'bulkAddAssignments', key, assignments: [...] }
 */
function handleBulkAddAssignments_(body) {
  try {
    var result = bulkAddTentativeAssignments(body.assignments || []);
    return jsonResponse_(result);
  } catch (e) {
    Logger.log('handleBulkAddAssignments_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
  }
}

/**
 * 仮配置クリア (POST)
 * body: { action: 'clearAllocations', key, month }
 */
function handleClearAllocations_(body) {
  try {
    var targetMonth = body.month || '';
    if (!targetMonth) {
      return jsonResponse_({ error: '対象年月が指定されていません。' });
    }

    var count = clearTentativeAssignments(targetMonth);
    return jsonResponse_({ success: true, clearedCount: count });
  } catch (e) {
    Logger.log('handleClearAllocations_ error: ' + e.toString());
    return jsonResponse_({ error: e.toString() }, 500);
  }
}
