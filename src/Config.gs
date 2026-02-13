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
  FACILITY: 4,    // D: 施設名(正式)
  FACILITY_ID: 5, // E: 施設コード
  TIME_SLOT: 6,   // F: 時間帯
  ORIGINAL_NAME: 7,// G: 担当者名(原文)
  EMPLOYEE_NO: 8, // H: 社員No
  FORMAL_NAME: 9  // I: 氏名(正式)
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
  BROADCAST_URL: 'https://api.line.me/v2/bot/message/broadcast',
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

// ===== 施設マッピング (CSV施設名 → Firestore施設ID/正式名) =====
// safe-rise-prod Firestore facilities コレクションと同期
const FACILITY_MAP = {
  'グリーンビレッジB': { id: 'GH3', name: 'グリーンビレッジＢ' },
  'グリーンビレッジE': { id: 'GH7', name: 'グリーンビレッジＥ' },
  'ビレッジE102': { id: 'GH29', name: 'グリーンビレッジE102' },
  '中町': { id: 'GH6', name: '中町' },
  '南大泉': { id: 'GH8', name: '南大泉' },
  '南大泉３丁目': { id: 'GH28', name: '南大泉３丁目' },
  '大泉町': { id: 'GH1', name: '大泉町' },
  '春日町同一①': { id: 'GH9', name: '春日町 (B103)' },
  '春日町２同一①': { id: 'GH12', name: '春日町2 (B203)' },
  '東大泉': { id: 'GH10', name: '東大泉' },
  '松原': { id: 'GH37', name: '松原' },
  '石神井公園': { id: 'GH41', name: '石神井公園' },
  '砧①107': { id: 'GH25', name: '砧 (107)' },
  '砧②207': { id: 'GH36', name: '砧2 (207)' },
  '西長久保': { id: 'GH11', name: '西長久保' },
  '都民農園': { id: 'GH24', name: '都民農園' },
  '長久保': { id: 'GH4', name: '長久保' },
  '関町南2F同一②': { id: 'GH14', name: '関町南2' },
  '関町南3F同一②': { id: 'GH15', name: '関町南3' },
  '関町南4F同一②': { id: 'GH15', name: '関町南3' }
};

/**
 * CSV施設名をFirestoreの正式施設名に変換
 * @param {string} csvFacilityName - CSV上の施設名
 * @return {string} 正式施設名 (マッピングがない場合はそのまま返す)
 */
function getOfficialFacilityName(csvFacilityName) {
  const entry = FACILITY_MAP[csvFacilityName];
  return entry ? entry.name : csvFacilityName;
}

/**
 * CSV施設名からFirestore施設IDを取得
 * @param {string} csvFacilityName - CSV上の施設名
 * @return {string|null} 施設ID (マッピングがない場合はnull)
 */
function getFacilityId(csvFacilityName) {
  const entry = FACILITY_MAP[csvFacilityName];
  return entry ? entry.id : null;
}

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
