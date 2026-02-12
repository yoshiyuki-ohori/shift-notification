/**
 * Config.gs - 定数・設定管理
 * シフト通知システムの設定値を一元管理
 */

// ===== シート名定数 =====
const SHEET_NAMES = {
  EMPLOYEE_MASTER: '従業員マスタ',
  SHIFT_DATA: 'シフトデータ',
  SEND_LOG: '送信ログ',
  SETTINGS: '設定',
  UNMATCHED: '名寄せ未マッチ'
};

// ===== 従業員マスタ列定数 (1-indexed) =====
const MASTER_COLS = {
  NO: 1,          // A: 社員番号
  NAME: 2,        // B: 氏名
  FURIGANA: 3,    // C: フリガナ
  LINE_USER_ID: 4,// D: LINE_UserId
  AREA: 5,        // E: エリア
  FACILITY: 6,    // F: 主担当施設
  STATUS: 7,      // G: ステータス (在職/退職)
  NOTIFY: 8,      // H: 通知有効 (TRUE/FALSE)
  ALIASES: 9      // I: 別名リスト (カンマ区切り)
};

// ===== シフトデータ列定数 =====
const SHIFT_COLS = {
  YEAR_MONTH: 1,  // A: 年月
  DATE: 2,        // B: 日付
  AREA: 3,        // C: エリア
  FACILITY: 4,    // D: 施設名
  TIME_SLOT: 5,   // E: 時間帯
  ORIGINAL_NAME: 6,// F: 担当者名(原文)
  EMPLOYEE_NO: 7, // G: 社員No
  FORMAL_NAME: 8  // H: 氏名(正式)
};

// ===== 送信ログ列定数 =====
const LOG_COLS = {
  TIMESTAMP: 1,   // A: 送信日時
  TARGET_MONTH: 2,// B: 対象年月
  EMPLOYEE_NO: 3, // C: 社員No
  NAME: 4,        // D: 氏名
  STATUS: 5,      // E: ステータス
  ERROR_DETAIL: 6 // F: エラー詳細
};

// ===== 未マッチ列定数 =====
const UNMATCHED_COLS = {
  FACILITY: 1,    // A: 施設名
  ORIGINAL_NAME: 2,// B: 原文名前
  CANDIDATES: 3,  // C: 候補
  RESOLVED: 4     // D: 解決済
};

// ===== 設定キー =====
const SETTING_KEYS = {
  TARGET_MONTH: '対象年月',
  NERIMA_SS_ID: '練馬シフトSS_ID',
  SETAGAYA_SS_ID: '世田谷シフトSS_ID',
  SEND_MODE: '送信モード',
  TEST_USER_ID: 'テスト送信先UserId'
};

// ===== LINE API定数 =====
const LINE_API = {
  PUSH_URL: 'https://api.line.me/v2/bot/message/push',
  REPLY_URL: 'https://api.line.me/v2/bot/message/reply',
  RATE_LIMIT_DELAY_MS: 100
};

// ===== バッチ定数 =====
const BATCH = {
  SEND_BATCH_SIZE: 50,
  TRIGGER_DELAY_MINUTES: 1,
  MAX_EXECUTION_MS: 300000  // 5分 (6分制限に余裕を持たせる)
};

// ===== 送信ステータス =====
const SEND_STATUS = {
  SUCCESS: '成功',
  FAILED: '失敗',
  SKIPPED: 'スキップ',
  NO_LINE_ID: 'LINE未登録',
  INACTIVE: '無効'
};

// ===== 練馬エリア 時間帯定数 =====
const NERIMA_TIME_SLOTS = ['6時～9時', '17時～22時', '22時～'];

/**
 * 設定シートから値を取得
 * @param {string} key - 設定キー
 * @return {string} 設定値
 */
function getSettingValue(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) throw new Error('設定シートが見つかりません');

  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      return String(data[i][1]).trim();
    }
  }
  throw new Error('設定キーが見つかりません: ' + key);
}

/**
 * 設定シートに値を書き込む
 * @param {string} key - 設定キー
 * @param {string} value - 設定値
 */
function setSettingValue(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) throw new Error('設定シートが見つかりません');

  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  // キーが無い場合は追記
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
}

/**
 * Script PropertiesからLINEトークンを取得
 * @return {string} LINE Channel Access Token
 */
function getLineToken() {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) throw new Error('LINE_CHANNEL_ACCESS_TOKENが設定されていません。Script Propertiesを確認してください。');
  return token;
}

/**
 * Script PropertiesからLINEチャネルシークレットを取得
 * @return {string} LINE Channel Secret
 */
function getLineChannelSecret() {
  const secret = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_SECRET');
  if (!secret) throw new Error('LINE_CHANNEL_SECRETが設定されていません。Script Propertiesを確認してください。');
  return secret;
}
