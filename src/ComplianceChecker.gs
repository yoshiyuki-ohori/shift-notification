/**
 * ComplianceChecker.gs - 労基法コンプライアンスチェックエンジン
 * シフトデータに対して労働基準法チェックを実行する
 *
 * チェック項目:
 *   C-01: 月間労働時間超過 (変形労働時間制: 40h×暦日/7)
 *   C-02: 連続勤務日数超過
 *   C-03: 勤務間インターバル不足
 *   C-04: 月間夜勤回数超過
 *   C-05: 週間休日不足
 *   C-06: 4週休日不足
 *   C-07: 残業月45時間超過
 *   C-08: 残業月100時間到達
 */

/**
 * コンプライアンスチェックを実行
 * @param {string} targetMonth - 対象年月 (YYYY-MM)
 * @return {Object} { violations: [...], warnings: [...], summary: {...} }
 */
function runComplianceCheck(targetMonth) {
  var rules = loadLaborRules();
  var aggregated = aggregateByEmployee(targetMonth);
  var shiftsByEmployee = buildShiftMap_(targetMonth);

  var violations = [];
  var warnings = [];
  var employeesWithIssues = {};

  for (var i = 0; i < aggregated.length; i++) {
    var emp = aggregated[i];
    var empShifts = shiftsByEmployee[emp.employeeNo] || [];

    // C-01: 月間労働時間超過
    checkMonthlyHours_(emp, empShifts, targetMonth, rules, violations, warnings, employeesWithIssues);

    // C-02: 連続勤務日数超過
    checkConsecutiveDays_(emp, empShifts, targetMonth, rules, violations, warnings, employeesWithIssues);

    // C-03: 勤務間インターバル不足
    checkInterval_(emp, empShifts, rules, violations, warnings, employeesWithIssues);

    // C-04: 月間夜勤回数超過
    checkNightShiftCount_(emp, empShifts, rules, violations, warnings, employeesWithIssues);

    // C-05: 週間休日不足
    checkWeeklyHoliday_(emp, empShifts, targetMonth, rules, violations, warnings, employeesWithIssues);

    // C-06: 4週休日不足
    checkFourWeekHoliday_(emp, empShifts, targetMonth, rules, violations, warnings, employeesWithIssues);

    // C-07: 残業月45時間超過
    // C-08: 残業月100時間到達
    checkOvertime_(emp, empShifts, targetMonth, rules, violations, warnings, employeesWithIssues);
  }

  var summary = {
    totalViolations: violations.length,
    totalWarnings: warnings.length,
    employeesWithIssues: Object.keys(employeesWithIssues).length,
    checkedEmployees: aggregated.length
  };

  return {
    violations: violations,
    warnings: warnings,
    summary: summary
  };
}

/**
 * シフトデータを社員別に詳細マップとして構築
 * @param {string} targetMonth - 対象年月
 * @return {Object} { empNo: [{ date: Date, dateStr, timeSlot, startHour, endHour, hours }] }
 * @private
 */
function buildShiftMap_(targetMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shiftSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_DATA);
  if (!shiftSheet || shiftSheet.getLastRow() <= 1) return {};

  var data = shiftSheet.getDataRange().getValues();
  var result = {};

  for (var i = 1; i < data.length; i++) {
    var rawYM = data[i][SHIFT_COLS.YEAR_MONTH - 1];
    var yearMonth;
    if (rawYM instanceof Date) {
      yearMonth = rawYM.getFullYear() + '-' + ('0' + (rawYM.getMonth() + 1)).slice(-2);
    } else {
      yearMonth = String(rawYM).trim();
    }
    if (yearMonth !== targetMonth) continue;

    var empNo = String(data[i][SHIFT_COLS.EMPLOYEE_NO - 1]).trim();
    if (!empNo) continue;

    var rawDate = data[i][SHIFT_COLS.DATE - 1];
    var dateObj;
    var dateStr;
    if (rawDate instanceof Date) {
      dateObj = rawDate;
      dateStr = rawDate.getFullYear() + '/' + ('0' + (rawDate.getMonth() + 1)).slice(-2) + '/' + ('0' + rawDate.getDate()).slice(-2);
    } else {
      dateStr = String(rawDate).trim();
      var parts = dateStr.split('/');
      dateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    }

    var timeSlot = String(data[i][SHIFT_COLS.TIME_SLOT - 1]).trim();
    var slotInfo = TIME_SLOT_HOURS[timeSlot] || { start: 0, end: 0, hours: 0 };

    if (!result[empNo]) result[empNo] = [];
    result[empNo].push({
      date: dateObj,
      dateStr: dateStr,
      timeSlot: timeSlot,
      startHour: slotInfo.start,
      endHour: slotInfo.end,
      hours: slotInfo.hours
    });
  }

  // 日付順にソート
  var empNos = Object.keys(result);
  for (var j = 0; j < empNos.length; j++) {
    result[empNos[j]].sort(function(a, b) {
      return a.date.getTime() - b.date.getTime() || a.startHour - b.startHour;
    });
  }

  return result;
}

/**
 * C-01: 月間労働時間超過チェック
 * 変形労働時間制: 上限 = 40h × 暦日数 / 7
 */
function checkMonthlyHours_(emp, shifts, targetMonth, rules, violations, warnings, issueMap) {
  var parts = targetMonth.split('-');
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var daysInMonth = new Date(year, month, 0).getDate();

  var weeklyHours = rules['労働時間|法定週労働時間'] || 40;
  var monthlyLimit = weeklyHours * daysInMonth / 7;

  var totalHours = 0;
  for (var i = 0; i < shifts.length; i++) {
    totalHours += shifts[i].hours;
  }

  if (totalHours > monthlyLimit) {
    violations.push({
      type: 'C-01',
      severity: 'violation',
      employeeNo: emp.employeeNo,
      employeeName: emp.name,
      detail: '月間労働時間' + totalHours.toFixed(1) + '時間（上限' + monthlyLimit.toFixed(1) + '時間）',
      dates: [targetMonth]
    });
    issueMap[emp.employeeNo] = true;
  }
}

/**
 * C-02: 連続勤務日数超過チェック
 */
function checkConsecutiveDays_(emp, shifts, targetMonth, rules, violations, warnings, issueMap) {
  var limit = rules['連勤|連続勤務上限'] || 6;
  var absLimit = rules['連勤|連続勤務絶対上限'] || 12;

  // 勤務日を抽出（ユニーク日付）
  var workDays = {};
  for (var i = 0; i < shifts.length; i++) {
    var key = shifts[i].date.getFullYear() + '/' +
              ('0' + (shifts[i].date.getMonth() + 1)).slice(-2) + '/' +
              ('0' + shifts[i].date.getDate()).slice(-2);
    workDays[key] = true;
  }

  var parts = targetMonth.split('-');
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var daysInMonth = new Date(year, month, 0).getDate();

  var consecutive = 0;
  var startDate = null;
  var dates = [];

  for (var d = 1; d <= daysInMonth; d++) {
    var dateStr = year + '/' + ('0' + month).slice(-2) + '/' + ('0' + d).slice(-2);
    if (workDays[dateStr]) {
      consecutive++;
      dates.push(month + '/' + d);
      if (!startDate) startDate = dateStr;
    } else {
      if (consecutive > limit) {
        var severity = consecutive >= absLimit ? 'violation' : 'warning';
        var item = {
          type: 'C-02',
          severity: severity,
          employeeNo: emp.employeeNo,
          employeeName: emp.name,
          detail: '連続勤務' + consecutive + '日（上限' + limit + '日）',
          dates: dates.slice()
        };
        if (severity === 'violation') {
          violations.push(item);
        } else {
          warnings.push(item);
        }
        issueMap[emp.employeeNo] = true;
      }
      consecutive = 0;
      startDate = null;
      dates = [];
    }
  }

  // 月末まで連続している場合
  if (consecutive > limit) {
    var sev = consecutive >= absLimit ? 'violation' : 'warning';
    var lastItem = {
      type: 'C-02',
      severity: sev,
      employeeNo: emp.employeeNo,
      employeeName: emp.name,
      detail: '連続勤務' + consecutive + '日（上限' + limit + '日）',
      dates: dates.slice()
    };
    if (sev === 'violation') {
      violations.push(lastItem);
    } else {
      warnings.push(lastItem);
    }
    issueMap[emp.employeeNo] = true;
  }
}

/**
 * C-03: 勤務間インターバル不足チェック
 */
function checkInterval_(emp, shifts, rules, violations, warnings, issueMap) {
  var intervalHours = rules['インターバル|勤務間インターバル'] || 11;

  if (shifts.length < 2) return;

  // 日付×時間帯ごとにグループ化し、各日の最終終了時刻と翌日の最初開始時刻を比較
  var dayShifts = {};
  for (var i = 0; i < shifts.length; i++) {
    var dayKey = shifts[i].dateStr;
    if (!dayShifts[dayKey]) dayShifts[dayKey] = [];
    dayShifts[dayKey].push(shifts[i]);
  }

  var days = Object.keys(dayShifts).sort();
  for (var d = 0; d < days.length - 1; d++) {
    var todayShifts = dayShifts[days[d]];
    var tomorrowShifts = dayShifts[days[d + 1]];

    // 今日の最終終了時刻
    var latestEnd = 0;
    for (var t = 0; t < todayShifts.length; t++) {
      if (todayShifts[t].endHour > latestEnd) {
        latestEnd = todayShifts[t].endHour;
      }
    }

    // 翌日の最早開始時刻
    var earliestStart = 48;
    for (var n = 0; n < tomorrowShifts.length; n++) {
      var start = tomorrowShifts[n].startHour;
      if (start < earliestStart) {
        earliestStart = start;
      }
    }

    // endHour が 24超（翌日跨ぎ）の場合、翌日の開始に24を足して比較
    var gap;
    if (latestEnd > 24) {
      // 22時～翌6時の場合: endHour=30, 翌日startが6時だとgap=0
      gap = (earliestStart + 24) - latestEnd;
    } else {
      gap = (earliestStart + 24) - latestEnd;
    }

    if (gap < intervalHours) {
      warnings.push({
        type: 'C-03',
        severity: 'warning',
        employeeNo: emp.employeeNo,
        employeeName: emp.name,
        detail: '勤務間インターバル' + gap + '時間（推奨' + intervalHours + '時間）',
        dates: [formatShortDate_(days[d]) + ' → ' + formatShortDate_(days[d + 1])]
      });
      issueMap[emp.employeeNo] = true;
    }
  }
}

/**
 * C-04: 月間夜勤回数超過チェック
 */
function checkNightShiftCount_(emp, shifts, rules, violations, warnings, issueMap) {
  var nightStartHour = rules['深夜|深夜開始時刻'] || 22;
  var maxNightShifts = rules['深夜|月間夜勤回数上限'] || 8;

  var nightCount = 0;
  for (var i = 0; i < shifts.length; i++) {
    if (shifts[i].startHour >= nightStartHour || shifts[i].endHour > 24) {
      nightCount++;
    }
  }

  // 夜勤は22時～のシフトのみカウント（17時～22時は含まない）
  var nightShiftCount = 0;
  var nightDates = [];
  var countedDays = {};
  for (var j = 0; j < shifts.length; j++) {
    if (shifts[j].timeSlot === '22時～' && !countedDays[shifts[j].dateStr]) {
      nightShiftCount++;
      nightDates.push(formatShortDate_(shifts[j].dateStr));
      countedDays[shifts[j].dateStr] = true;
    }
  }

  if (nightShiftCount > maxNightShifts) {
    warnings.push({
      type: 'C-04',
      severity: 'warning',
      employeeNo: emp.employeeNo,
      employeeName: emp.name,
      detail: '月間夜勤' + nightShiftCount + '回（上限' + maxNightShifts + '回）',
      dates: nightDates
    });
    issueMap[emp.employeeNo] = true;
  }
}

/**
 * C-05: 週間休日不足チェック（7日間で休日0日）
 */
function checkWeeklyHoliday_(emp, shifts, targetMonth, rules, violations, warnings, issueMap) {
  var minWeeklyHoliday = rules['休日|週最低休日'] || 1;

  var parts = targetMonth.split('-');
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var daysInMonth = new Date(year, month, 0).getDate();

  // 勤務日セット
  var workDays = {};
  for (var i = 0; i < shifts.length; i++) {
    workDays[shifts[i].dateStr] = true;
  }

  // 7日スライディングウィンドウ
  for (var startDay = 1; startDay <= daysInMonth - 6; startDay++) {
    var holidays = 0;
    var weekDates = [];
    for (var d = 0; d < 7; d++) {
      var day = startDay + d;
      if (day > daysInMonth) break;
      var dateStr = year + '/' + ('0' + month).slice(-2) + '/' + ('0' + day).slice(-2);
      weekDates.push(month + '/' + day);
      if (!workDays[dateStr]) {
        holidays++;
      }
    }

    if (holidays < minWeeklyHoliday && weekDates.length === 7) {
      violations.push({
        type: 'C-05',
        severity: 'violation',
        employeeNo: emp.employeeNo,
        employeeName: emp.name,
        detail: '7日間連続勤務（週1日の休日なし）',
        dates: weekDates
      });
      issueMap[emp.employeeNo] = true;
      break; // 1つ見つかれば十分
    }
  }
}

/**
 * C-06: 4週休日不足チェック（28日間で休日 < 4日）
 */
function checkFourWeekHoliday_(emp, shifts, targetMonth, rules, violations, warnings, issueMap) {
  var minFourWeekHoliday = rules['休日|4週最低休日'] || 4;

  var parts = targetMonth.split('-');
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var daysInMonth = new Date(year, month, 0).getDate();

  // 勤務日セット
  var workDays = {};
  for (var i = 0; i < shifts.length; i++) {
    workDays[shifts[i].dateStr] = true;
  }

  // 月全体での休日数をカウント
  var totalHolidays = 0;
  for (var d = 1; d <= daysInMonth; d++) {
    var dateStr = year + '/' + ('0' + month).slice(-2) + '/' + ('0' + d).slice(-2);
    if (!workDays[dateStr]) {
      totalHolidays++;
    }
  }

  // 暦日数に比例した必要休日数（28日で4日 = 4週基準）
  var requiredHolidays = Math.floor(daysInMonth / 7) * minFourWeekHoliday / 4;
  // 簡易的に月全体で4日未満なら違反
  if (totalHolidays < minFourWeekHoliday) {
    violations.push({
      type: 'C-06',
      severity: 'violation',
      employeeNo: emp.employeeNo,
      employeeName: emp.name,
      detail: '月間休日' + totalHolidays + '日（4週' + minFourWeekHoliday + '日以上必要）',
      dates: [targetMonth]
    });
    issueMap[emp.employeeNo] = true;
  }
}

/**
 * C-07 & C-08: 残業時間チェック
 * C-07: 月間時間外 > 36協定月上限 (45h) → warning
 * C-08: 月間時間外(+休日労働) >= 特別条項月上限 (100h) → violation
 */
function checkOvertime_(emp, shifts, targetMonth, rules, violations, warnings, issueMap) {
  var dailyLimit = rules['労働時間|法定日労働時間'] || 8;
  var monthlyOvertimeLimit = rules['残業|36協定月上限'] || 45;
  var specialLimit = rules['残業|特別条項月上限'] || 100;

  var parts = targetMonth.split('-');
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var daysInMonth = new Date(year, month, 0).getDate();

  // 日ごとの労働時間を集計
  var dailyHours = {};
  for (var i = 0; i < shifts.length; i++) {
    var dayKey = shifts[i].dateStr;
    if (!dailyHours[dayKey]) dailyHours[dayKey] = 0;
    dailyHours[dayKey] += shifts[i].hours;
  }

  // 時間外労働を計算（日ごとに法定労働時間を超えた分）
  var totalOvertime = 0;
  var days = Object.keys(dailyHours);
  for (var j = 0; j < days.length; j++) {
    var dayHours = dailyHours[days[j]];
    if (dayHours > dailyLimit) {
      totalOvertime += dayHours - dailyLimit;
    }
  }

  // C-08: 特別条項月上限 (100h)
  if (totalOvertime >= specialLimit) {
    violations.push({
      type: 'C-08',
      severity: 'violation',
      employeeNo: emp.employeeNo,
      employeeName: emp.name,
      detail: '月間時間外労働' + totalOvertime.toFixed(1) + '時間（上限' + specialLimit + '時間）',
      dates: [targetMonth]
    });
    issueMap[emp.employeeNo] = true;
  }
  // C-07: 36協定月上限 (45h)
  else if (totalOvertime > monthlyOvertimeLimit) {
    warnings.push({
      type: 'C-07',
      severity: 'warning',
      employeeNo: emp.employeeNo,
      employeeName: emp.name,
      detail: '月間時間外労働' + totalOvertime.toFixed(1) + '時間（36協定上限' + monthlyOvertimeLimit + '時間）',
      dates: [targetMonth]
    });
    issueMap[emp.employeeNo] = true;
  }
}

/**
 * チェック結果をシートに書き出す
 * @param {string} targetMonth - 対象年月
 * @param {Object} result - runComplianceCheck() の戻り値
 */
function writeComplianceResults(targetMonth, result) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.COMPLIANCE_RESULT);

  if (!sheet) {
    setupComplianceResultSheet();
    sheet = ss.getSheetByName(SHEET_NAMES.COMPLIANCE_RESULT);
  }

  var now = new Date();
  var rows = [];

  // violations
  for (var i = 0; i < result.violations.length; i++) {
    var v = result.violations[i];
    rows.push([
      now, targetMonth, v.type, v.severity,
      v.employeeNo, v.employeeName, v.detail,
      v.dates.join(', ')
    ]);
  }

  // warnings
  for (var j = 0; j < result.warnings.length; j++) {
    var w = result.warnings[j];
    rows.push([
      now, targetMonth, w.type, w.severity,
      w.employeeNo, w.employeeName, w.detail,
      w.dates.join(', ')
    ]);
  }

  if (rows.length > 0) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, 8).setValues(rows);
  }
}

/**
 * 日付文字列を短い形式に変換
 * @param {string} dateStr - "YYYY/MM/DD"
 * @return {string} "M/D"
 * @private
 */
function formatShortDate_(dateStr) {
  if (!dateStr) return '';
  var parts = dateStr.split('/');
  if (parts.length >= 3) {
    return parseInt(parts[1], 10) + '/' + parseInt(parts[2], 10);
  }
  return dateStr;
}
