/**
 * ShiftParser.gs - シフトデータ解析ユーティリティ
 */

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
    const cell = data[i][0];
    let cellYM = '';
    if (cell instanceof Date) {
      // Googleスプレッドシートが "2026-02" をDate型に変換した場合
      const y = cell.getFullYear();
      const m = String(cell.getMonth() + 1).padStart(2, '0');
      cellYM = y + '-' + m;
    } else {
      cellYM = String(cell).trim();
    }
    if (cellYM === targetMonth) {
      rowsToDelete.push(i + 1);
    }
  }

  // 逆順で削除（行番号がずれないように）
  for (const row of rowsToDelete) {
    sheet.deleteRow(row);
  }
}
