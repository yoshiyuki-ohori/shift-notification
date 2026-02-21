/**
 * StaffingChecker.gs - 配置充足率チェック・エリア別管理
 * 「必要配置」シートの基準値と実配置を比較し、充足状況を判定する
 *
 * 3段階判定:
 *   shortage (不足): 実配置 < 最低人数 → 即座に対応必要
 *   caution (注意): 最低人数 <= 実配置 < 目標値 → バッファ不足
 *   ok (目標達成): 実配置 >= 目標値 → 安定運用可能
 */

/**
 * 配置充足率チェックを実行
 * @param {string} targetMonth - 対象年月 (YYYY-MM)
 * @return {Object} { facilities: [...], areaSummary: {...}, alerts: [...] }
 */
function checkStaffingLevels(targetMonth) {
  var requirements = loadStaffingRequirements_();
  var shiftData = loadShiftDataForStaffing_(targetMonth);
  var prefData = loadPreferenceDataForStaffing_(targetMonth);

  var parts = targetMonth.split('-');
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var daysInMonth = new Date(year, month, 0).getDate();

  var facilities = [];
  var alerts = [];
  var areaTotals = {};

  // 施設×日付×時間帯ごとにチェック
  for (var d = 1; d <= daysInMonth; d++) {
    var dateStr = year + '/' + ('0' + month).slice(-2) + '/' + ('0' + d).slice(-2);
    var dateObj = new Date(year, month - 1, d);
    var dayType = getDayType_(dateObj);

    // 全施設の要件をチェック
    for (var r = 0; r < requirements.length; r++) {
      var req = requirements[r];

      // 曜日種別が一致するもののみ
      if (req.dayType !== dayType) continue;

      var facilityId = req.facilityId;
      var timeSlot = req.timeSlot;

      // 実配置数を取得
      var assigned = countAssigned_(shiftData, facilityId, dateStr, timeSlot);
      // 希望数を取得
      var prefWant = countPreferences_(prefData, dateStr, timeSlot);

      var status;
      var fillRate;
      if (assigned < req.minStaff) {
        status = 'shortage';
      } else if (assigned < req.preferredStaff) {
        status = 'caution';
      } else {
        status = 'ok';
      }
      fillRate = req.minStaff > 0 ? Math.round(assigned / req.minStaff * 100) : 100;

      var facilityName = getFacilityNameById_(facilityId);
      var area = getFacilityAreaById_(facilityId);

      var entry = {
        facilityId: facilityId,
        facilityName: facilityName,
        date: dateStr,
        timeSlot: timeSlot,
        dayType: dayType,
        required: req.minStaff,
        target: req.preferredStaff,
        assigned: assigned,
        prefWant: prefWant,
        status: status,
        fillRate: fillRate
      };

      facilities.push(entry);

      if (status === 'shortage') {
        alerts.push(entry);
      }

      // エリア別集計
      if (!areaTotals[area]) {
        areaTotals[area] = { totalSlots: 0, shortage: 0, caution: 0, ok: 0, totalFillRate: 0 };
      }
      areaTotals[area].totalSlots++;
      areaTotals[area][status]++;
      areaTotals[area].totalFillRate += fillRate;
    }
  }

  // エリア別サマリー生成
  var areaSummary = {};
  var areaKeys = Object.keys(areaTotals);
  for (var a = 0; a < areaKeys.length; a++) {
    var areaName = areaKeys[a];
    var totals = areaTotals[areaName];
    areaSummary[areaName] = {
      totalSlots: totals.totalSlots,
      shortage: totals.shortage,
      caution: totals.caution,
      ok: totals.ok,
      avgFillRate: totals.totalSlots > 0 ? Math.round(totals.totalFillRate / totals.totalSlots) : 0
    };
  }

  return {
    targetMonth: targetMonth,
    facilities: facilities,
    areaSummary: areaSummary,
    alerts: alerts
  };
}

/**
 * 「必要配置」シートから要件を読み込む
 * @return {Array<Object>} [{ facilityId, timeSlot, dayType, minStaff, preferredStaff }]
 * @private
 */
function loadStaffingRequirements_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.STAFFING_REQUIREMENT);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var facilityId = String(data[i][STAFFING_COLS.FACILITY_ID - 1]).trim();
    var timeSlot = String(data[i][STAFFING_COLS.TIME_SLOT - 1]).trim();
    var dayType = String(data[i][STAFFING_COLS.DAY_TYPE - 1]).trim();
    var minStaff = Number(data[i][STAFFING_COLS.MIN_STAFF - 1]) || 0;
    var preferredStaff = Number(data[i][STAFFING_COLS.PREFERRED_STAFF - 1]) || 0;

    if (!facilityId) continue;

    result.push({
      facilityId: facilityId,
      timeSlot: timeSlot,
      dayType: dayType,
      minStaff: minStaff,
      preferredStaff: preferredStaff
    });
  }

  return result;
}

/**
 * シフトデータを配置チェック用に読み込む
 * @param {string} targetMonth - 対象年月
 * @return {Object} { "facilityId|dateStr|timeSlot": count }
 * @private
 */
function loadShiftDataForStaffing_(targetMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shiftSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_DATA);
  if (!shiftSheet || shiftSheet.getLastRow() <= 1) return {};

  var data = shiftSheet.getDataRange().getValues();
  var counts = {};

  for (var i = 1; i < data.length; i++) {
    var rawYM = data[i][SHIFT_COLS.YEAR_MONTH - 1];
    var yearMonth;
    if (rawYM instanceof Date) {
      yearMonth = rawYM.getFullYear() + '-' + ('0' + (rawYM.getMonth() + 1)).slice(-2);
    } else {
      yearMonth = String(rawYM).trim();
    }
    if (yearMonth !== targetMonth) continue;

    var facilityId = String(data[i][SHIFT_COLS.FACILITY_ID - 1]).trim();
    var rawDate = data[i][SHIFT_COLS.DATE - 1];
    var dateStr;
    if (rawDate instanceof Date) {
      dateStr = rawDate.getFullYear() + '/' + ('0' + (rawDate.getMonth() + 1)).slice(-2) + '/' + ('0' + rawDate.getDate()).slice(-2);
    } else {
      dateStr = String(rawDate).trim();
    }
    var timeSlot = String(data[i][SHIFT_COLS.TIME_SLOT - 1]).trim();
    var empNo = String(data[i][SHIFT_COLS.EMPLOYEE_NO - 1]).trim();

    if (!facilityId || !empNo) continue;

    var key = facilityId + '|' + dateStr + '|' + timeSlot;
    counts[key] = (counts[key] || 0) + 1;
  }

  return counts;
}

/**
 * シフト希望データを配置チェック用に読み込む
 * @param {string} targetMonth - 対象年月
 * @return {Object} { "dateStr|timeSlot": count }
 * @private
 */
function loadPreferenceDataForStaffing_(targetMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_PREFERENCE);
  if (!sheet || sheet.getLastRow() <= 1) return {};

  var data = sheet.getDataRange().getValues();
  var counts = {};

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][PREF_COLS.YEAR_MONTH - 1]).trim() !== targetMonth) continue;
    var type = String(data[i][PREF_COLS.TYPE - 1]).trim();
    if (type !== PREF_TYPE.WANT) continue;

    var dateStr = String(data[i][PREF_COLS.DATE - 1]).trim();
    var timeSlot = String(data[i][PREF_COLS.TIME_SLOT - 1]).trim();

    var key = dateStr + '|' + timeSlot;
    counts[key] = (counts[key] || 0) + 1;
  }

  return counts;
}

/**
 * 実配置数をカウント
 * @private
 */
function countAssigned_(shiftData, facilityId, dateStr, timeSlot) {
  var key = facilityId + '|' + dateStr + '|' + timeSlot;
  return shiftData[key] || 0;
}

/**
 * 希望数をカウント
 * @private
 */
function countPreferences_(prefData, dateStr, timeSlot) {
  var key = dateStr + '|' + timeSlot;
  return prefData[key] || 0;
}

/**
 * 曜日種別を判定
 * @param {Date} dateObj - 日付
 * @return {string} '平日' / '土曜' / '日祝'
 * @private
 */
function getDayType_(dateObj) {
  var dow = dateObj.getDay();
  if (dow === 0) return '日祝';
  if (dow === 6) return '土曜';
  return '平日';
}

/**
 * 施設IDから施設名を取得
 * @param {string} facilityId - 施設ID
 * @return {string} 施設名
 * @private
 */
function getFacilityNameById_(facilityId) {
  var keys = Object.keys(FACILITY_MAP);
  for (var i = 0; i < keys.length; i++) {
    if (FACILITY_MAP[keys[i]].id === facilityId) {
      return FACILITY_MAP[keys[i]].name;
    }
  }
  return facilityId;
}

/**
 * 施設IDからエリアを判定
 * @param {string} facilityId - 施設ID
 * @return {string} エリア名
 * @private
 */
function getFacilityAreaById_(facilityId) {
  // 世田谷エリアの施設ID
  var setagayaIds = ['GH6', 'GH25', 'GH36', 'GH37'];
  if (setagayaIds.indexOf(facilityId) >= 0) {
    return '世田谷';
  }
  return '練馬';
}
