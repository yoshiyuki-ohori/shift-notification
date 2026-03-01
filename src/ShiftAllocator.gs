/**
 * ShiftAllocator.gs - 仮配置 CRUD + 候補者一覧
 * 仮配置シートの読み書き・候補者検索を提供する
 */

/**
 * 仮配置データを読み込む
 * @param {string} targetMonth - 対象年月 (YYYY-MM)
 * @return {Array<Object>} 仮配置データ配列
 */
function loadTentativeAssignments(targetMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TENTATIVE_ASSIGNMENT);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var ym = formatYearMonth_(data[i][TENTATIVE_COLS.YEAR_MONTH - 1]);
    if (ym !== targetMonth) continue;

    result.push({
      yearMonth: ym,
      date: formatDateValue_(data[i][TENTATIVE_COLS.DATE - 1]),
      area: String(data[i][TENTATIVE_COLS.AREA - 1]).trim(),
      facility: String(data[i][TENTATIVE_COLS.FACILITY - 1]).trim(),
      facilityId: String(data[i][TENTATIVE_COLS.FACILITY_ID - 1]).trim(),
      timeSlot: String(data[i][TENTATIVE_COLS.TIME_SLOT - 1]).trim(),
      empNo: String(data[i][TENTATIVE_COLS.EMPLOYEE_NO - 1]).trim().padStart(3, '0'),
      name: String(data[i][TENTATIVE_COLS.NAME - 1]).trim(),
      status: String(data[i][TENTATIVE_COLS.STATUS - 1]).trim(),
      source: String(data[i][TENTATIVE_COLS.SOURCE - 1]).trim(),
      prefMatch: String(data[i][TENTATIVE_COLS.PREF_MATCH - 1]).trim(),
      assignedAt: String(data[i][TENTATIVE_COLS.ASSIGNED_AT - 1]).trim(),
      assignedBy: String(data[i][TENTATIVE_COLS.ASSIGNED_BY - 1]).trim(),
      rowIndex: i + 1  // 1-indexed sheet row
    });
  }

  return result;
}

/**
 * 仮配置データを一括書き込み
 * @param {Array<Object>} assignments - 仮配置データ配列
 */
function writeTentativeAssignments(assignments) {
  if (!assignments || assignments.length === 0) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureTentativeSheet_(ss);

  var rows = assignments.map(function(a) {
    return [
      a.yearMonth,
      a.date,
      a.area || '',
      a.facility,
      a.facilityId || '',
      a.timeSlot,
      a.empNo,
      a.name,
      a.status || ASSIGN_STATUS.TENTATIVE,
      a.source || 'manual',
      a.prefMatch || '',
      a.assignedAt || new Date().toISOString(),
      a.assignedBy || '管理者'
    ];
  });

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 13).setValues(rows);
}

/**
 * 指定年月の仮配置をクリア
 * @param {string} targetMonth - 対象年月 (YYYY-MM)
 * @return {number} 削除件数
 */
function clearTentativeAssignments(targetMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TENTATIVE_ASSIGNMENT);
  if (!sheet || sheet.getLastRow() <= 1) return 0;

  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];

  for (var i = data.length - 1; i >= 1; i--) {
    var ym = formatYearMonth_(data[i][TENTATIVE_COLS.YEAR_MONTH - 1]);
    if (ym === targetMonth) {
      rowsToDelete.push(i + 1);
    }
  }

  // 下から削除（行番号がずれないように）
  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j]);
  }

  return rowsToDelete.length;
}

/**
 * 仮配置を1件追加
 * @param {Object} assignment - 配置データ
 * @return {Object} 追加結果 { success, conflict }
 */
function addTentativeAssignment(assignment) {
  // 同日・同時間帯の重複チェック（同じ職員が別施設に配置されていないか）
  var conflict = checkDuplicateAssignment_(
    assignment.yearMonth || '',
    assignment.date,
    assignment.timeSlot,
    assignment.empNo
  );

  // 希望一致チェック
  var prefMatch = checkPreferenceMatch_(
    assignment.yearMonth || '',
    assignment.date,
    assignment.timeSlot,
    assignment.empNo
  );

  var now = new Date().toISOString();
  var entry = {
    yearMonth: assignment.yearMonth,
    date: assignment.date,
    area: assignment.area || '',
    facility: assignment.facility,
    facilityId: assignment.facilityId || '',
    timeSlot: assignment.timeSlot,
    empNo: assignment.empNo,
    name: assignment.name,
    status: ASSIGN_STATUS.TENTATIVE,
    source: 'manual',
    prefMatch: prefMatch,
    assignedAt: now,
    assignedBy: assignment.assignedBy || '管理者'
  };

  writeTentativeAssignments([entry]);

  return {
    success: true,
    conflict: conflict,
    prefMatch: prefMatch
  };
}

/**
 * 仮配置を1件削除
 * @param {string} yearMonth - 年月
 * @param {string} date - 日付
 * @param {string} facilityId - 施設ID
 * @param {string} timeSlot - 時間帯
 * @param {string} empNo - 社員No
 * @return {boolean} 削除成功
 */
function removeTentativeAssignment(yearMonth, date, facilityId, timeSlot, empNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TENTATIVE_ASSIGNMENT);
  if (!sheet || sheet.getLastRow() <= 1) return false;

  var data = sheet.getDataRange().getValues();
  var targetEmpNo = String(empNo).padStart(3, '0');

  for (var i = data.length - 1; i >= 1; i--) {
    var ym = formatYearMonth_(data[i][TENTATIVE_COLS.YEAR_MONTH - 1]);
    var d = formatDateValue_(data[i][TENTATIVE_COLS.DATE - 1]);
    var fid = String(data[i][TENTATIVE_COLS.FACILITY_ID - 1]).trim();
    var facility = String(data[i][TENTATIVE_COLS.FACILITY - 1]).trim();
    var ts = String(data[i][TENTATIVE_COLS.TIME_SLOT - 1]).trim();
    var en = String(data[i][TENTATIVE_COLS.EMPLOYEE_NO - 1]).trim().padStart(3, '0');

    if (ym === yearMonth && d === date && (fid === facilityId || facility === facilityId) && ts === timeSlot && en === targetEmpNo) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }

  return false;
}

/**
 * 仮配置の統計を取得
 * @param {string} targetMonth - 対象年月
 * @return {Object} 統計データ
 */
function getTentativeStats(targetMonth) {
  var assignments = loadTentativeAssignments(targetMonth);
  var requirements = loadStaffingRequirements_();

  var parts = targetMonth.split('-');
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var daysInMonth = new Date(year, month, 0).getDate();

  // 配置数をカウント: facilityId|date|timeSlot → count
  var assignedCounts = {};
  var totalAssigned = 0;
  var empSlots = {};  // empNo → count
  var conflicts = [];

  for (var i = 0; i < assignments.length; i++) {
    var a = assignments[i];
    var key = a.facilityId + '|' + a.date + '|' + a.timeSlot;
    assignedCounts[key] = (assignedCounts[key] || 0) + 1;
    totalAssigned++;

    empSlots[a.empNo] = (empSlots[a.empNo] || 0) + 1;
  }

  // 重複チェック: 同一職員が同日同時間帯に複数施設
  var empDateSlot = {};
  for (var j = 0; j < assignments.length; j++) {
    var b = assignments[j];
    var dupKey = b.empNo + '|' + b.date + '|' + b.timeSlot;
    if (!empDateSlot[dupKey]) {
      empDateSlot[dupKey] = [];
    }
    empDateSlot[dupKey].push(b.facility);
  }
  var dupKeys = Object.keys(empDateSlot);
  for (var dk = 0; dk < dupKeys.length; dk++) {
    if (empDateSlot[dupKeys[dk]].length > 1) {
      conflicts.push({
        key: dupKeys[dk],
        facilities: empDateSlot[dupKeys[dk]]
      });
    }
  }

  // 必要配置に対する充足率
  var totalSlots = 0;
  var filledSlots = 0;
  var shortageSlots = 0;

  for (var d = 1; d <= daysInMonth; d++) {
    var dateStr = year + '/' + ('0' + month).slice(-2) + '/' + ('0' + d).slice(-2);
    var dateObj = new Date(year, month - 1, d);
    var dayType = getDayType_(dateObj);

    for (var r = 0; r < requirements.length; r++) {
      var req = requirements[r];
      if (req.dayType !== dayType) continue;

      totalSlots++;
      var slotKey = req.facilityId + '|' + dateStr + '|' + req.timeSlot;
      var assigned = assignedCounts[slotKey] || 0;
      if (assigned >= req.minStaff) filledSlots++;
      else shortageSlots++;
    }
  }

  return {
    totalAssigned: totalAssigned,
    totalSlots: totalSlots,
    filledSlots: filledSlots,
    shortageSlots: shortageSlots,
    fillRate: totalSlots > 0 ? Math.round(filledSlots / totalSlots * 100) : 0,
    conflicts: conflicts,
    empSlots: empSlots
  };
}

/**
 * 指定日付・時間帯に配置可能な職員リストを取得
 * 希望データと照合して優先度付きで返す
 * @param {string} targetMonth - 対象年月
 * @param {string} date - 日付 (YYYY/MM/DD)
 * @param {string} timeSlot - 時間帯
 * @param {string} [facilityId] - 施設ID (省略可)
 * @return {Array<Object>} 候補職員リスト
 */
function getAvailableStaff(targetMonth, date, timeSlot, facilityId) {
  var employees = loadEmployeeMaster();
  var assignments = loadTentativeAssignments(targetMonth);

  // 全シフト希望をロード
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var prefSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);
  var prefMap = {};  // empNo|date → { type, timeSlot }
  if (prefSheet && prefSheet.getLastRow() > 1) {
    var prefData = prefSheet.getDataRange().getValues();
    for (var p = 1; p < prefData.length; p++) {
      var pYM = formatYearMonth_(prefData[p][PREF_COLS.YEAR_MONTH - 1]);
      if (pYM !== targetMonth) continue;
      var pEmpNo = String(prefData[p][PREF_COLS.EMPLOYEE_NO - 1]).trim().padStart(3, '0');
      var pDate = formatDateValue_(prefData[p][PREF_COLS.DATE - 1]);
      var pType = String(prefData[p][PREF_COLS.TYPE - 1]).trim();
      var pTimeSlot = String(prefData[p][PREF_COLS.TIME_SLOT - 1]).trim();

      var prefKey = pEmpNo + '|' + pDate;
      if (!prefMap[prefKey]) prefMap[prefKey] = [];
      prefMap[prefKey].push({ type: pType, timeSlot: pTimeSlot });
    }
  }

  // 同日・同時間帯に既に配置済みの職員を特定
  var assignedSameSlot = {};  // empNo → facility
  var empMonthlyCount = {};   // empNo → count
  for (var a = 0; a < assignments.length; a++) {
    var asn = assignments[a];
    empMonthlyCount[asn.empNo] = (empMonthlyCount[asn.empNo] || 0) + 1;
    if (asn.date === date && asn.timeSlot === timeSlot) {
      assignedSameSlot[asn.empNo] = asn.facility;
    }
  }

  // 候補者リスト生成
  var candidates = [];
  for (var e = 0; e < employees.length; e++) {
    var emp = employees[e];
    if (emp.status !== '在職') continue;

    var empNo = emp.employeeNo;
    var prefs = prefMap[empNo + '|' + date] || [];

    // 希望マッチ判定
    var prefType = 'none';  // none / want / either / ng
    for (var pi = 0; pi < prefs.length; pi++) {
      var pref = prefs[pi];
      // 時間帯が一致 or 時間帯指定なし
      if (!pref.timeSlot || pref.timeSlot === timeSlot || pref.timeSlot === '') {
        if (pref.type === PREF_TYPE.WANT) prefType = 'want';
        else if (pref.type === PREF_TYPE.NG) prefType = 'ng';
        else if (pref.type === PREF_TYPE.EITHER && prefType !== 'want' && prefType !== 'ng') prefType = 'either';
      }
    }

    var alreadyAssigned = !!assignedSameSlot[empNo];
    var assignedFacility = assignedSameSlot[empNo] || null;

    candidates.push({
      empNo: empNo,
      name: emp.name,
      prefType: prefType,
      alreadyAssigned: alreadyAssigned,
      assignedFacility: assignedFacility,
      monthlyCount: empMonthlyCount[empNo] || 0,
      sortOrder: prefType === 'want' ? 0 : prefType === 'either' ? 1 : prefType === 'none' ? 2 : 3
    });
  }

  // ソート: 希望者→どちらでも→希望なし→NG、同じカテゴリ内は月間配置数少ない順
  candidates.sort(function(a, b) {
    if (a.sortOrder !== b.sortOrder) return a.sortOrder - b.sortOrder;
    return a.monthlyCount - b.monthlyCount;
  });

  return candidates;
}

// ===== Private helpers =====

/**
 * 仮配置シートを確保（なければ作成）
 * @private
 */
function ensureTentativeSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.TENTATIVE_ASSIGNMENT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.TENTATIVE_ASSIGNMENT);
    sheet.getRange(1, 1, 1, 13).setValues([[
      '年月', '日付', 'エリア', '施設名', '施設ID', '時間帯',
      '社員No', '氏名', 'ステータス', '配置理由', '希望一致', '配置日時', '配置者'
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 同日・同時間帯に同じ職員の重複配置をチェック
 * @return {Object|null} 衝突情報 or null
 * @private
 */
function checkDuplicateAssignment_(yearMonth, date, timeSlot, empNo) {
  var assignments = loadTentativeAssignments(yearMonth);
  var targetEmpNo = String(empNo).padStart(3, '0');

  for (var i = 0; i < assignments.length; i++) {
    var a = assignments[i];
    if (a.empNo === targetEmpNo && a.date === date && a.timeSlot === timeSlot) {
      return {
        facility: a.facility,
        facilityId: a.facilityId,
        date: a.date,
        timeSlot: a.timeSlot
      };
    }
  }
  return null;
}

/**
 * 希望一致をチェック
 * @return {string} 'want' / 'either' / 'ng' / 'none'
 * @private
 */
function checkPreferenceMatch_(yearMonth, date, timeSlot, empNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);
  if (!sheet || sheet.getLastRow() <= 1) return 'none';

  var data = sheet.getDataRange().getValues();
  var targetEmpNo = String(empNo).padStart(3, '0');

  for (var i = 1; i < data.length; i++) {
    var ym = formatYearMonth_(data[i][PREF_COLS.YEAR_MONTH - 1]);
    if (ym !== yearMonth) continue;
    var en = String(data[i][PREF_COLS.EMPLOYEE_NO - 1]).trim().padStart(3, '0');
    if (en !== targetEmpNo) continue;
    var d = formatDateValue_(data[i][PREF_COLS.DATE - 1]);
    if (d !== date) continue;
    var ts = String(data[i][PREF_COLS.TIME_SLOT - 1]).trim();
    if (ts && ts !== timeSlot) continue;

    var type = String(data[i][PREF_COLS.TYPE - 1]).trim();
    if (type === PREF_TYPE.WANT) return 'want';
    if (type === PREF_TYPE.NG) return 'ng';
    if (type === PREF_TYPE.EITHER) return 'either';
  }

  return 'none';
}

/**
 * 短縮日付フォーマット
 * @private
 */
function formatShortDate_(dateStr) {
  if (!dateStr) return '';
  var parts = String(dateStr).split('/');
  if (parts.length >= 3) {
    return parseInt(parts[1], 10) + '/' + parseInt(parts[2], 10);
  }
  return dateStr;
}
