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
  UNMATCHED: '名寄せ未マッチ',
  SHIFT_PREFERENCE: 'シフト希望',
  STAFFING_REQUIREMENT: '必要配置',
  LABOR_RULES: '労基ルール',
  COMPLIANCE_RESULT: '労基チェック結果',
  ATTENDANCE_CONFIRM: '出勤確認',
  TENTATIVE_ASSIGNMENT: '仮配置'
};

// ===== 従業員マスタ列定数 (1-indexed) =====
const MASTER_COLS = {
  NO: 1,          // A: 社員番号
  NAME: 2,        // B: 氏名
  FURIGANA: 3,    // C: フリガナ
  LINE_USER_ID: 4,// D: LINE_UserId
  STATUS: 5,      // E: ステータス (在職/退職)
  NOTIFY: 6,      // F: 通知有効 (TRUE/FALSE)
  ALIASES: 7      // G: 別名リスト (カンマ区切り)
};

// ===== シフトデータ列定数 =====
// シート: 年月,日付,エリア,施設名,時間帯,担当者名(原文),社員No,氏名(正式)
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

// ===== シフト希望列定数 =====
const PREF_COLS = {
  YEAR_MONTH: 1,    // A: 対象年月
  EMPLOYEE_NO: 2,   // B: 社員番号
  NAME: 3,          // C: 氏名
  DATE: 4,          // D: 日付
  TYPE: 5,          // E: 希望/NG/どちらでも
  TIME_SLOT: 6,     // F: 希望時間帯
  REASON: 7,        // G: 理由
  SUBMITTED_AT: 8,  // H: 提出日時
  FACILITY: 9       // I: 希望施設（カンマ区切り）
};

// ===== 必要配置列定数 =====
const STAFFING_COLS = {
  FACILITY_ID: 1,      // A: 施設ID
  TIME_SLOT: 2,        // B: 時間帯
  DAY_TYPE: 3,         // C: 平日/土曜/日祝
  MIN_STAFF: 4,        // D: 最低人数
  PREFERRED_STAFF: 5   // E: 推奨人数
};

// ===== 出勤確認列定数 =====
const ATTEND_COLS = {
  DATE: 1,            // A: 日付
  EMPLOYEE_NO: 2,     // B: 社員No
  NAME: 3,            // C: 氏名
  STATUS: 4,          // D: ステータス (未確認/出勤/欠勤連絡)
  CONFIRMED_AT: 5,    // E: 確認日時
  NOTIFIED_AT: 6      // F: 通知日時
};

// ===== 仮配置シート名・列定数 =====
// SHEET_NAMES に追加
// (下記 SHEET_NAMES オブジェクトの末尾に TENTATIVE_ASSIGNMENT を追加済み)

const TENTATIVE_COLS = {
  YEAR_MONTH: 1,     // A: 年月
  DATE: 2,           // B: 日付
  AREA: 3,           // C: エリア
  FACILITY: 4,       // D: 施設名
  FACILITY_ID: 5,    // E: 施設ID
  TIME_SLOT: 6,      // F: 時間帯
  EMPLOYEE_NO: 7,    // G: 社員No
  NAME: 8,           // H: 氏名
  STATUS: 9,         // I: ステータス (仮配置/確定)
  SOURCE: 10,        // J: 配置理由 (manual)
  PREF_MATCH: 11,    // K: 希望一致
  ASSIGNED_AT: 12,   // L: 配置日時
  ASSIGNED_BY: 13    // M: 配置者
};

const ASSIGN_STATUS = {
  TENTATIVE: '仮配置',
  CONFIRMED: '確定'
};

// ===== 希望種別 =====
const PREF_TYPE = {
  WANT: '希望',
  NG: 'NG',
  EITHER: 'どちらでも'
};

// ===== 設定キー =====
const SETTING_KEYS = {
  TARGET_MONTH: '対象年月',
  SEND_MODE: '送信モード',
  TEST_USER_ID: 'テスト送信先UserId',
  PREF_COLLECTION_START: '希望収集開始日',
  PREF_COLLECTION_END: '希望収集締切日',
  PREF_TARGET_MONTH: '希望対象年月',
  PREF_PERIOD_LABEL: '希望収集期間区分',
  LIFF_ID: 'LIFF_ID',
  EMPLOYEE_MASTER_SS_ID: 'EMPLOYEE_MASTER_SS_ID',
  ADMIN_NOTIFY_USER_ID: 'テスト送信先UserId'
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

// ===== 時間帯→労働時間マッピング =====
const TIME_SLOT_HOURS = {
  '6時～9時':   { start: 6, end: 9, hours: 3 },
  '9時～17時':  { start: 9, end: 17, hours: 8 },   // 旧選択肢（既存データ互換）
  '17時～22時': { start: 17, end: 22, hours: 5 },
  '22時～':     { start: 22, end: 30, hours: 8 },   // 旧選択肢（既存データ互換）
  '22時～6時':  { start: 22, end: 30, hours: 8 },
  '22時～9時':  { start: 22, end: 33, hours: 11 },
  '17時～翌9時': { start: 17, end: 33, hours: 16 }  // 夜勤通し
};

// ===== 労基チェック結果列定数 =====
const COMPLIANCE_COLS = {
  CHECK_DATE: 1,    // A: チェック日時
  TARGET_MONTH: 2,  // B: 対象年月
  CHECK_ID: 3,      // C: チェックID
  SEVERITY: 4,      // D: 重大度
  EMPLOYEE_NO: 5,   // E: 社員No
  NAME: 6,          // F: 氏名
  DETAIL: 7,        // G: 詳細
  DATES: 8          // H: 対象日
};

// ===== 希望入力用 施設リスト（エリア別・住所ベース） =====
const PREF_FACILITY_LIST = {
  '練馬区': [
    'グリーンビレッジＢ', 'グリーンビレッジＥ103', 'グリーンビレッジE102',
    '大泉町', '南大泉', '南大泉３丁目', '東大泉', '西大泉',
    '長久保', '西長久保', '都民農園',
    '春日町 (B103)', '春日町2 (B203)', '豊島園',
    '石神井公園',
    '関町南（UPG101）', '関町南（UPG306）',
    '練馬1 (203)', '練馬2 (303)',
    '江古田 (part1 201)', '江古田2 (part2 205)'
  ],
  '世田谷区': [
    '砧 (107)', '砧2 (207)',
    '松原',
    '芦花公園（ﾚｼﾞｵﾝ）', '芦花公園2（ｴﾘｰｾﾞ）'
  ],
  '中野区': [
    '若宮', '野方', '沼袋'
  ],
  '杉並区': [
    '下井草', '方南'
  ],
  '板橋区': [
    '成増3 (池ミュハイツ)',
    '赤塚2 (アーバン)', '赤塚3 (第7)',
    '富士見町'
  ],
  '西東京市': [
    '中町',
    '芝久保（センチュリー201）', '芝久保（センチュリー301）'
  ],
  '立川市': [
    '立川', 'エルウィング', 'SSエルウィング'
  ],
  '大阪': [
    '天満①', '天満②',
    '寿', '淡路'
  ],
  '千葉': [
    '船橋', '津田沼'
  ],
  '埼玉': [
    '所沢'
  ]
};

// ===== 同一建物グループ (施設名 → 建物グループID) =====
// CSV名と正式名（FACILITY_MAP変換後）の両方を登録
// 住所ベース: 同一住所の施設を同一グループにまとめる
const BUILDING_GROUPS = {
  // 春日町: 練馬区春日町2-15-8 ティーケー平和台
  '春日町同一①': 'kasugacho', '春日町 (B103)': 'kasugacho',
  '春日町２同一①': 'kasugacho', '春日町2 (B203)': 'kasugacho',
  // 関町南: 練馬区関町南4-21-12 UPGレジデンス関町南
  '関町南2F同一②': 'sekimachi', '関町南2階': 'sekimachi',
  '関町南3F同一②': 'sekimachi', '関町南4F同一②': 'sekimachi',
  '関町南1': 'sekimachi', '関町南2': 'sekimachi', '関町南3': 'sekimachi',
  '関町南（UPG101）': 'sekimachi', '関町南（UPG306）': 'sekimachi',
  // 砧: 世田谷区砧2-10-10 ティーケー砧
  '砧①107': 'kinuta', '砧 (107)': 'kinuta',
  '砧②207': 'kinuta', '砧2 (207)': 'kinuta',
  // 練馬: 練馬区練馬1-6-16 サンパレス練馬
  '練馬203': 'nerima', '練馬1 (203)': 'nerima', '練馬1': 'nerima',
  '練馬2 (303)': 'nerima', '練馬2': 'nerima',
  // 芦花公園: 世田谷区北烏山1-12-8/9
  'レジオン': 'roka', '芦花公園': 'roka', '芦花公園（ﾚｼﾞｵﾝ）': 'roka',
  '江リーザ': 'roka', 'エリーザ': 'roka', '芦花公園2': 'roka', '芦花公園2（ｴﾘｰｾﾞ）': 'roka',
  // 芝久保: 西東京市芝久保町3-4-29 センチュリー芝久保
  '芝久保1': 'shibakubo', '芝久保２': 'shibakubo',
  '芝久保2': 'shibakubo', '芝久保１': 'shibakubo',
  '芝久保3': 'shibakubo', '芝久保３': 'shibakubo',
  '芝久保（センチュリー201）': 'shibakubo', '芝久保（センチュリー301）': 'shibakubo',
  // 江古田: 練馬区旭丘2-22 ユエヴィ江古田
  '江古田': 'ekoda', '江古田 (part1 201)': 'ekoda',
  '江古田2': 'ekoda', '江古田2 (part2 205)': 'ekoda',
  // 天満: 大阪市北区与力町4-4 ヴェルディオクト
  '天満①': 'tenma', '天満①102': 'tenma',
  '天満②': 'tenma', '天満②302': 'tenma',
  // グリーンビレッジ: 練馬区大泉町3-26 (B棟14, E棟47)
  'グリーンビレッジB': 'green_village', 'グリーンビレッジＢ': 'green_village',
  'グリーンビレッジE': 'green_village', 'グリーンビレッジＥ': 'green_village',
  'グリーンビレッジＥ103': 'green_village',
  'ビレッジE102': 'green_village', 'グリーンビレッジE102': 'green_village',
  // エルウィング: 立川市錦町1-15-31
  'エルウィング': 'elwing', 'SSエルウィング': 'elwing'
};

// ===== 施設マッピング (CSV施設名 → Firestore施設ID/正式名) =====
// safe-rise-prod Firestore facilities コレクションと同期
const FACILITY_MAP = {
  'グリーンビレッジB': { id: 'GH3', name: 'グリーンビレッジＢ' },
  'グリーンビレッジE': { id: 'GH7', name: 'グリーンビレッジＥ103' },
  'グリーンビレッジＥ103': { id: 'GH7', name: 'グリーンビレッジＥ103' },
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
  '関町南2F同一②': { id: 'GH14', name: '関町南（UPG101）' },
  '関町南3F同一②': { id: 'GH15', name: '関町南（UPG306）' },
  '関町南4F同一②': { id: 'GH15', name: '関町南（UPG306）' },
  '関町南1': { id: 'GH14', name: '関町南（UPG101）' },
  '関町南2': { id: 'GH14', name: '関町南（UPG101）' },
  '関町南3': { id: 'GH15', name: '関町南（UPG306）' },
  '関町南（UPG101）': { id: 'GH14', name: '関町南（UPG101）' },
  '関町南（UPG306）': { id: 'GH15', name: '関町南（UPG306）' }
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
      var val = data[i][1];
      // Date型はYYYY-MM-DD形式に変換（Sheetsが自動変換するため）
      if (val instanceof Date) {
        var y = val.getFullYear();
        var m = ('0' + (val.getMonth() + 1)).slice(-2);
        var d = ('0' + val.getDate()).slice(-2);
        return y + '-' + m + '-' + d;
      }
      return String(val).trim();
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
 * 日付値を "YYYY/MM/DD" 形式の文字列に変換
 * Google Sheets が Date オブジェクトを返す場合に対応
 * @param {Date|string} rawDate - セルの日付値
 * @return {string} "YYYY/MM/DD" 形式
 */
function formatDateValue_(rawDate) {
  if (rawDate instanceof Date) {
    var y = rawDate.getFullYear();
    var m = ('0' + (rawDate.getMonth() + 1)).slice(-2);
    var d = ('0' + rawDate.getDate()).slice(-2);
    return y + '/' + m + '/' + d;
  }
  return String(rawDate).trim();
}

/**
 * 年月値を "YYYY-MM" 形式の文字列に変換
 * Google Sheets が Date オブジェクトを返す場合に対応
 * @param {Date|string} rawYM - セルの年月値
 * @return {string} "YYYY-MM" 形式
 */
function formatYearMonth_(rawYM) {
  if (rawYM instanceof Date) {
    return rawYM.getFullYear() + '-' + ('0' + (rawYM.getMonth() + 1)).slice(-2);
  }
  return String(rawYM).trim();
}

// ===== LINEチャネル種別 =====
const LINE_CHANNEL = {
  SHIFT: 'shift',
  UNIFIED: 'unified'
};

/**
 * Script PropertiesからLINEトークンを取得
 * @param {string} [channel] - チャネル種別 ('shift' or 'unified')。省略時はshift
 * @return {string} LINE Channel Access Token
 */
function getLineToken(channel) {
  const props = PropertiesService.getScriptProperties();
  if (channel === LINE_CHANNEL.UNIFIED) {
    const token = props.getProperty('LINE_UNIFIED_CHANNEL_ACCESS_TOKEN');
    if (!token) throw new Error('LINE_UNIFIED_CHANNEL_ACCESS_TOKENが設定されていません。');
    return token;
  }
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) throw new Error('LINE_CHANNEL_ACCESS_TOKENが設定されていません。');
  return token;
}

/**
 * Script PropertiesからLINEチャネルシークレットを取得
 * @param {string} [channel] - チャネル種別 ('shift' or 'unified')。省略時はshift
 * @return {string} LINE Channel Secret
 */
function getLineChannelSecret(channel) {
  const props = PropertiesService.getScriptProperties();
  if (channel === LINE_CHANNEL.UNIFIED) {
    const secret = props.getProperty('LINE_UNIFIED_CHANNEL_SECRET');
    if (!secret) throw new Error('LINE_UNIFIED_CHANNEL_SECRETが設定されていません。');
    return secret;
  }
  const secret = props.getProperty('LINE_CHANNEL_SECRET');
  if (!secret) throw new Error('LINE_CHANNEL_SECRETが設定されていません。');
  return secret;
}

/**
 * 日付文字列を短い形式に変換
 * @param {string} dateStr - "YYYY/MM/DD"
 * @return {string} "M/D"
 */
function formatShortDate_(dateStr) {
  if (!dateStr) return '';
  var parts = String(dateStr).split('/');
  if (parts.length >= 3) {
    return parseInt(parts[1], 10) + '/' + parseInt(parts[2], 10);
  }
  return dateStr;
}

/**
 * Webhook受信時にdestinationからチャネル種別を判定
 * @param {string} destination - Bot User ID
 * @return {string} チャネル種別
 */
function detectChannel(destination) {
  const props = PropertiesService.getScriptProperties();
  const unifiedBotId = props.getProperty('LINE_UNIFIED_BOT_USER_ID');
  if (unifiedBotId && destination === unifiedBotId) {
    return LINE_CHANNEL.UNIFIED;
  }
  return LINE_CHANNEL.SHIFT;
}
