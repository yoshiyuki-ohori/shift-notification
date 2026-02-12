/**
 * ShiftAggregator.gs - 従業員ごとのシフト集約
 * シフトデータを従業員単位にまとめて送信用データを生成
 */

/**
 * シフトデータを従業員ごとに集約
 * @param {string} targetMonth - 対象年月 (例: "2026-03")
 * @return {Array<Object>} 従業員ごとの集約データ
 *   [{
 *     employeeNo: "072",
 *     name: "石井 祐一",
 *     lineUserId: "Uxxxx",
 *     notifyEnabled: true,
 *     shifts: [
 *       { date: "2026/03/01", timeSlot: "17時～22時", facility: "南大泉" },
 *       ...
 *     ]
 *   }]
 */
function aggregateByEmployee(targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shiftSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_DATA);
  if (!shiftSheet || shiftSheet.getLastRow() <= 1) return [];

  const shiftData = shiftSheet.getDataRange().getValues();
  const employeeMaster = loadEmployeeMaster();

  // 従業員マスタをMapに変換
  const empMap = new Map();
  for (const emp of employeeMaster) {
    empMap.set(emp.employeeNo, emp);
  }

  // 社員Noでグルーピング
  const grouped = new Map();

  for (let i = 1; i < shiftData.length; i++) {
    const yearMonth = String(shiftData[i][SHIFT_COLS.YEAR_MONTH - 1]).trim();
    if (yearMonth !== targetMonth) continue;

    const employeeNo = String(shiftData[i][SHIFT_COLS.EMPLOYEE_NO - 1]).trim();
    if (!employeeNo) continue;

    const shift = {
      date: String(shiftData[i][SHIFT_COLS.DATE - 1]).trim(),
      timeSlot: String(shiftData[i][SHIFT_COLS.TIME_SLOT - 1]).trim(),
      facility: String(shiftData[i][SHIFT_COLS.FACILITY - 1]).trim()
    };

    if (!grouped.has(employeeNo)) {
      grouped.set(employeeNo, []);
    }
    grouped.get(employeeNo).push(shift);
  }

  // 従業員情報と結合
  const result = [];
  for (const [empNo, shifts] of grouped) {
    const emp = empMap.get(empNo);
    if (!emp) continue;

    // 日付・時間帯でソート
    shifts.sort((a, b) => {
      const dateCompare = a.date.localeCompare(b.date);
      if (dateCompare !== 0) return dateCompare;
      return getTimeSlotOrder(a.timeSlot) - getTimeSlotOrder(b.timeSlot);
    });

    result.push({
      employeeNo: empNo,
      name: emp.name,
      lineUserId: emp.lineUserId,
      notifyEnabled: emp.notifyEnabled,
      status: emp.status,
      shifts: shifts
    });
  }

  // 社員番号順にソート
  result.sort((a, b) => a.employeeNo.localeCompare(b.employeeNo));

  return result;
}

/**
 * 時間帯の表示順を返す
 * @param {string} timeSlot - 時間帯文字列
 * @return {number} ソート順
 */
function getTimeSlotOrder(timeSlot) {
  if (timeSlot.startsWith('6')) return 1;
  if (timeSlot.startsWith('9')) return 2;
  if (timeSlot.startsWith('17')) return 3;
  if (timeSlot.startsWith('22')) return 4;
  return 5;
}

/**
 * 集約結果のサマリーを生成
 * @param {Array<Object>} aggregated - 集約データ
 * @return {Object} サマリー情報
 */
function getAggregationSummary(aggregated) {
  let totalShifts = 0;
  let withLineId = 0;
  let withoutLineId = 0;
  let notifyDisabled = 0;

  for (const emp of aggregated) {
    totalShifts += emp.shifts.length;
    if (emp.lineUserId) {
      withLineId++;
    } else {
      withoutLineId++;
    }
    if (!emp.notifyEnabled) {
      notifyDisabled++;
    }
  }

  return {
    totalEmployees: aggregated.length,
    totalShifts: totalShifts,
    withLineId: withLineId,
    withoutLineId: withoutLineId,
    notifyDisabled: notifyDisabled,
    sendableCount: withLineId - notifyDisabled
  };
}
