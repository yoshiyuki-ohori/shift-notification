/**
 * ShiftParser.gs - シフトデータ解析
 * 練馬パターン（施設別シート）と世田谷パターン（統合シート）を解析
 */

/**
 * 練馬パターンのシフト解析
 * 形式:
 *   Row1: 施設名,南大泉
 *   Row2: (空行)
 *   Row3: 日付,6時～9時,17時～22時,22時～
 *   Row4+: 1日,,石井 祐一,石井 祐一
 *
 * @param {Sheet} sheet - Google Spreadsheetのシートオブジェクト
 * @param {string} targetMonth - 対象年月 (例: "2026-03")
 * @return {Array<Object>} シフトレコード配列
 */
function parseNerimaShift(sheet, targetMonth) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 4) return [];

  // Row1: 施設名の取得
  const facilityName = extractFacilityName(data[0]);
  if (!facilityName) return [];

  // Row3: 時間帯ヘッダーの取得
  const timeSlots = extractTimeSlots(data);
  if (timeSlots.length === 0) return [];

  // ヘッダー行のインデックスを特定
  const headerRowIndex = findHeaderRow(data);
  if (headerRowIndex < 0) return [];

  const [year, month] = targetMonth.split('-').map(Number);
  const records = [];

  // データ行を処理
  for (let i = headerRowIndex + 1; i < data.length; i++) {
    const dayStr = String(data[i][0]).trim();
    const dayNum = extractDayNumber(dayStr);
    if (!dayNum) continue;

    const dateStr = formatDate(year, month, dayNum);

    for (let col = 1; col < data[i].length && col <= timeSlots.length; col++) {
      const staffName = String(data[i][col]).trim();
      if (!staffName) continue;

      records.push({
        yearMonth: targetMonth,
        date: dateStr,
        area: '練馬',
        facility: facilityName,
        timeSlot: timeSlots[col - 1],
        originalName: staffName,
        employeeNo: '',
        formalName: ''
      });
    }
  }

  return records;
}

/**
 * 世田谷パターンのシフト解析
 * 形式:
 *   Row1: 10月,砧①107,,,砧②207,,,松原,,,夜
 *   Row2: 月日,6時～9時,17時～,22時～,6時～9時,17時～,22時～,6時～9時,17時～,22時～,
 *   Row3+: 1日(火),,吉﨑,吉﨑,,旭,旭,,高橋百合,高橋百合,柳幸子
 *
 * @param {Sheet} sheet - Google Spreadsheetのシートオブジェクト
 * @param {string} targetMonth - 対象年月 (例: "2026-03")
 * @return {Array<Object>} シフトレコード配列
 */
function parseSetagayaShift(sheet, targetMonth) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 3) return [];

  // Row1から施設情報を解析
  const facilityMap = parseSetagayaFacilityHeader(data[0]);
  if (Object.keys(facilityMap).length === 0) return [];

  // Row2から各列の時間帯を取得
  const columnTimeSlots = parseSetagayaTimeHeader(data[1], facilityMap);

  // 対象月かチェック
  const monthStr = String(data[0][0]).trim();
  const sheetMonth = extractMonthNumber(monthStr);
  const [year, month] = targetMonth.split('-').map(Number);
  if (sheetMonth && sheetMonth !== month) return [];

  const records = [];

  // データ行を処理 (Row3以降)
  for (let i = 2; i < data.length; i++) {
    const dayStr = String(data[i][0]).trim();
    const dayNum = extractDayNumber(dayStr);
    if (!dayNum) continue;

    const dateStr = formatDate(year, month, dayNum);

    // 各列を処理
    for (let col = 1; col < data[i].length; col++) {
      const staffName = String(data[i][col]).trim();
      if (!staffName) continue;

      const colInfo = columnTimeSlots[col];
      if (!colInfo) continue;

      records.push({
        yearMonth: targetMonth,
        date: dateStr,
        area: '世田谷',
        facility: colInfo.facility,
        timeSlot: colInfo.timeSlot,
        originalName: staffName,
        employeeNo: '',
        formalName: ''
      });
    }
  }

  return records;
}

/**
 * 練馬シフトの施設名を抽出
 * @param {Array} headerRow - 1行目のデータ
 * @return {string} 施設名
 */
function extractFacilityName(headerRow) {
  // パターン: "施設名,南大泉" or 1列目が施設名で2列目が値
  if (String(headerRow[0]).trim() === '施設名' && headerRow[1]) {
    return String(headerRow[1]).trim();
  }
  // シート名から取得する場合のフォールバック
  return '';
}

/**
 * 練馬シフトの施設名をシート名から抽出
 * @param {string} sheetName - シート名 (例: "10月_南大泉_シフト")
 * @return {string} 施設名
 */
function extractFacilityFromSheetName(sheetName) {
  // パターン: "XX月_施設名_シフト" or "施設名"
  const match = sheetName.match(/\d+月[_\s]*(.+?)[_\s]*シフト/);
  if (match) return match[1];
  return sheetName;
}

/**
 * ヘッダー行（日付,時間帯...）のインデックスを特定
 * @param {Array<Array>} data - シートデータ
 * @return {number} ヘッダー行インデックス (-1 if not found)
 */
function findHeaderRow(data) {
  for (let i = 0; i < Math.min(data.length, 10); i++) {
    const firstCell = String(data[i][0]).trim();
    if (firstCell === '日付' || firstCell === '月日') {
      return i;
    }
  }
  return -1;
}

/**
 * 時間帯ヘッダーを抽出（練馬パターン）
 * @param {Array<Array>} data - シートデータ
 * @return {Array<string>} 時間帯配列
 */
function extractTimeSlots(data) {
  const headerIdx = findHeaderRow(data);
  if (headerIdx < 0) return [];

  const slots = [];
  for (let col = 1; col < data[headerIdx].length; col++) {
    const cell = String(data[headerIdx][col]).trim();
    if (cell) {
      slots.push(cell);
    }
  }
  return slots;
}

/**
 * 世田谷パターンの施設ヘッダーを解析
 * Row1: 10月,砧①107,,,砧②207,,,松原,,,夜
 * → { 1: {name: "砧①107", endCol: 3}, 4: {name: "砧②207", endCol: 6}, ... }
 *
 * @param {Array} headerRow - 1行目のデータ
 * @return {Object} 施設マップ {startCol: {name, endCol}}
 */
function parseSetagayaFacilityHeader(headerRow) {
  const facilityMap = {};
  let currentFacility = null;
  let currentStartCol = -1;

  for (let col = 1; col < headerRow.length; col++) {
    const cell = String(headerRow[col]).trim();
    if (cell && cell !== '') {
      // 前の施設の終了列を記録
      if (currentFacility !== null) {
        facilityMap[currentStartCol] = {
          name: currentFacility,
          endCol: col - 1
        };
      }
      currentFacility = cell;
      currentStartCol = col;
    }
  }

  // 最後の施設
  if (currentFacility !== null) {
    facilityMap[currentStartCol] = {
      name: currentFacility,
      endCol: headerRow.length - 1
    };
  }

  return facilityMap;
}

/**
 * 世田谷パターンの時間帯ヘッダーを解析
 * Row2の各列に対して、所属施設と時間帯をマッピング
 *
 * @param {Array} timeRow - 2行目のデータ
 * @param {Object} facilityMap - 施設マップ
 * @return {Object} 列→{facility, timeSlot}のマッピング
 */
function parseSetagayaTimeHeader(timeRow, facilityMap) {
  const columnMap = {};
  const facilityEntries = Object.entries(facilityMap).map(([col, info]) => ({
    startCol: Number(col),
    ...info
  })).sort((a, b) => a.startCol - b.startCol);

  for (let col = 1; col < timeRow.length; col++) {
    const timeSlot = String(timeRow[col]).trim();
    if (!timeSlot) continue;

    // この列がどの施設に属するかを判定
    let facility = '不明';
    for (const entry of facilityEntries) {
      if (col >= entry.startCol && col <= entry.endCol) {
        facility = entry.name;
        break;
      }
    }

    columnMap[col] = {
      facility: facility,
      timeSlot: normalizeTimeSlot(timeSlot)
    };
  }

  return columnMap;
}

/**
 * 日付文字列から日番号を抽出
 * @param {string} dayStr - "1日", "1日(火)", "1" etc.
 * @return {number|null} 日番号
 */
function extractDayNumber(dayStr) {
  if (!dayStr) return null;
  const match = String(dayStr).match(/^(\d{1,2})/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * 月文字列から月番号を抽出
 * @param {string} monthStr - "10月" etc.
 * @return {number|null} 月番号
 */
function extractMonthNumber(monthStr) {
  if (!monthStr) return null;
  const match = String(monthStr).match(/(\d{1,2})月/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * 年月日からフォーマットされた日付文字列を生成
 * @param {number} year - 年
 * @param {number} month - 月
 * @param {number} day - 日
 * @return {string} "YYYY/MM/DD" 形式
 */
function formatDate(year, month, day) {
  return year + '/' +
    String(month).padStart(2, '0') + '/' +
    String(day).padStart(2, '0');
}

/**
 * 時間帯表記を正規化
 * "17時～" → "17時～22時", "22時～" → "22時～翌朝"
 * @param {string} timeSlot - 時間帯文字列
 * @return {string} 正規化された時間帯
 */
function normalizeTimeSlot(timeSlot) {
  if (!timeSlot) return '';
  let normalized = timeSlot.trim();
  // すでに完全な形式の場合はそのまま
  if (normalized.match(/\d+時～\d+時/)) return normalized;
  // "17時～" のような末尾が空の場合
  if (normalized === '17時～') return '17時～22時';
  if (normalized === '22時～') return '22時～';
  return normalized;
}

/**
 * シフトデータシートを対象年月でクリアして再取込用に準備
 * @param {string} targetMonth - 対象年月
 */
function clearShiftDataForMonth(targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_DATA);
  if (!sheet || sheet.getLastRow() <= 1) return;

  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === targetMonth) {
      rowsToDelete.push(i + 1);
    }
  }

  // 逆順で削除（行番号がずれないように）
  for (const row of rowsToDelete) {
    sheet.deleteRow(row);
  }
}
