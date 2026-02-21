/**
 * LaborRules.gs - 労基ルール設定シート管理
 * 労働基準法の定数をスプレッドシートで管理し、法改正時に管理者が更新できるようにする
 */

/**
 * 労基ルールのデフォルト値定義
 */
var LABOR_RULE_DEFAULTS = [
  ['労働時間', '法定日労働時間', 8, '時間', '2026-02'],
  ['労働時間', '法定週労働時間', 40, '時間', '2026-02'],
  ['残業', '36協定月上限', 45, '時間', '2026-02'],
  ['残業', '36協定年上限', 360, '時間', '2026-02'],
  ['残業', '特別条項月上限', 100, '時間（休日労働含む）', '2026-02'],
  ['残業', '特別条項年上限', 720, '時間', '2026-02'],
  ['残業', '複数月平均上限', 80, '時間（2-6ヶ月平均）', '2026-02'],
  ['残業', '特別条項適用月数上限', 6, '月/年', '2026-02'],
  ['休憩', '6時間超休憩', 45, '分', '2026-02'],
  ['休憩', '8時間超休憩', 60, '分', '2026-02'],
  ['休日', '週最低休日', 1, '日', '2026-02'],
  ['休日', '4週最低休日', 4, '日', '2026-02'],
  ['連勤', '連続勤務上限', 6, '日（推奨）', '2026-02'],
  ['連勤', '連続勤務絶対上限', 12, '日（法定）', '2026-02'],
  ['インターバル', '勤務間インターバル', 11, '時間（努力義務）', '2026-02'],
  ['深夜', '深夜開始時刻', 22, '時', '2026-02'],
  ['深夜', '深夜終了時刻', 5, '時', '2026-02'],
  ['深夜', '月間夜勤回数上限', 8, '回（業界推奨）', '2026-02'],
  ['変形労働', '変形期間', '1ヶ月', '', '2026-02'],
  ['メタ', 'ルールバージョン', '2026-02', 'YYYY-MM', '2026-02'],
  ['メタ', '施行基準日', '2024-04-01', '', '2026-02']
];

/**
 * 「労基ルール」シートを初期セットアップ
 * setupSheets()から呼び出される
 */
function setupLaborRulesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.LABOR_RULES);

  if (sheet) return; // 既に存在する場合はスキップ

  sheet = ss.insertSheet(SHEET_NAMES.LABOR_RULES);

  // ヘッダー
  sheet.getRange(1, 1, 1, 5).setValues([[
    'カテゴリ', 'ルール名', '値', '備考', 'バージョン'
  ]]);
  sheet.setFrozenRows(1);

  // デフォルト値投入
  if (LABOR_RULE_DEFAULTS.length > 0) {
    sheet.getRange(2, 1, LABOR_RULE_DEFAULTS.length, 5).setValues(LABOR_RULE_DEFAULTS);
  }

  // 列幅調整
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 100);
}

/**
 * 「労基チェック結果」シートを初期セットアップ
 */
function setupComplianceResultSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.COMPLIANCE_RESULT);

  if (sheet) return;

  sheet = ss.insertSheet(SHEET_NAMES.COMPLIANCE_RESULT);
  sheet.getRange(1, 1, 1, 8).setValues([[
    'チェック日時', '対象年月', 'チェックID', '重大度', '社員No', '氏名', '詳細', '対象日'
  ]]);
  sheet.setFrozenRows(1);
}

/**
 * 労基ルールをシートから全て読み込む
 * @return {Object} { "カテゴリ|ルール名": value, ... } 形式のオブジェクト
 */
function loadLaborRules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.LABOR_RULES);

  if (!sheet || sheet.getLastRow() <= 1) {
    // シートが無い場合はデフォルト値から生成
    return loadLaborRulesFromDefaults_();
  }

  var data = sheet.getDataRange().getValues();
  var rules = {};

  for (var i = 1; i < data.length; i++) {
    var category = String(data[i][0]).trim();
    var name = String(data[i][1]).trim();
    var value = data[i][2];

    if (!category || !name) continue;

    var key = category + '|' + name;
    // 数値に変換可能であれば数値として保持
    if (typeof value === 'number') {
      rules[key] = value;
    } else {
      var numVal = Number(value);
      rules[key] = isNaN(numVal) ? String(value).trim() : numVal;
    }
  }

  return rules;
}

/**
 * 個別のルール値を取得
 * @param {string} category - カテゴリ名
 * @param {string} name - ルール名
 * @return {number|string} ルール値
 */
function getLaborRule(category, name) {
  var rules = loadLaborRules();
  var key = category + '|' + name;
  if (rules.hasOwnProperty(key)) {
    return rules[key];
  }
  throw new Error('労基ルールが見つかりません: ' + category + ' / ' + name);
}

/**
 * デフォルト値からルールオブジェクトを生成（フォールバック）
 * @return {Object} ルールオブジェクト
 * @private
 */
function loadLaborRulesFromDefaults_() {
  var rules = {};
  for (var i = 0; i < LABOR_RULE_DEFAULTS.length; i++) {
    var row = LABOR_RULE_DEFAULTS[i];
    var key = row[0] + '|' + row[1];
    var value = row[2];
    if (typeof value === 'number') {
      rules[key] = value;
    } else {
      var numVal = Number(value);
      rules[key] = isNaN(numVal) ? String(value).trim() : numVal;
    }
  }
  return rules;
}
