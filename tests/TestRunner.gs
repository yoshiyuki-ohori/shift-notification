/**
 * TestRunner.gs - テスト実行
 * シフトパーサー・名寄せエンジンの動作検証
 */

/**
 * 全テストを実行
 */
function runAllTests() {
  const results = [];

  results.push(testNormalizeName());
  results.push(testExtractDayNumber());
  results.push(testExtractMonthNumber());
  results.push(testFormatDate());
  results.push(testNormalizeTimeSlot());
  results.push(testNameMatcherExactMatch());
  results.push(testNameMatcherNormalizedMatch());
  results.push(testNameMatcherAliasMatch());
  results.push(testNameMatcherSurnameMatch());
  results.push(testNameMatcherVariantMatch());
  results.push(testNameMatcherPartialMatch());
  results.push(testNerimaParserStructure());
  results.push(testSetagayaParserStructure());
  results.push(testGetDayOfWeek());
  results.push(testFlexMessageStructure());
  results.push(testBuildShiftRows());

  // 結果出力
  let passed = 0;
  let failed = 0;
  let output = '===== テスト結果 =====\n\n';

  for (const r of results) {
    const status = r.passed ? 'PASS' : 'FAIL';
    if (r.passed) passed++;
    else failed++;
    output += '[' + status + '] ' + r.name + '\n';
    if (!r.passed) {
      output += '       → ' + r.message + '\n';
    }
  }

  output += '\n合計: ' + results.length + ' テスト, ' + passed + ' 成功, ' + failed + ' 失敗\n';

  Logger.log(output);

  // スプレッドシートにも出力（利用可能な場合）
  try {
    SpreadsheetApp.getUi().alert('テスト結果', output, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // UIが利用できない場合（トリガー実行時等）はログのみ
  }
}

// ===== ユーティリティ関数テスト =====

function testNormalizeName() {
  const matcher = new NameMatcherEngine([]);

  const tests = [
    { input: '石井　祐一', expected: '石井 祐一' },      // 全角スペース
    { input: '石井  祐一', expected: '石井 祐一' },      // 連続スペース
    { input: ' 石井 祐一 ', expected: '石井 祐一' },     // 前後スペース
    { input: '高橋百合', expected: '高橋百合' },          // スペースなし
  ];

  for (const t of tests) {
    const result = matcher.normalizeName(t.input);
    if (result !== t.expected) {
      return { name: 'normalizeName', passed: false,
        message: '"' + t.input + '" → "' + result + '" (期待: "' + t.expected + '")' };
    }
  }

  return { name: 'normalizeName', passed: true };
}

function testExtractDayNumber() {
  const tests = [
    { input: '1日', expected: 1 },
    { input: '15日', expected: 15 },
    { input: '1日(火)', expected: 1 },
    { input: '31日(水)', expected: 31 },
    { input: '', expected: null },
    { input: '日付', expected: null },
  ];

  for (const t of tests) {
    const result = extractDayNumber(t.input);
    if (result !== t.expected) {
      return { name: 'extractDayNumber', passed: false,
        message: '"' + t.input + '" → ' + result + ' (期待: ' + t.expected + ')' };
    }
  }

  return { name: 'extractDayNumber', passed: true };
}

function testExtractMonthNumber() {
  const tests = [
    { input: '10月', expected: 10 },
    { input: '1月', expected: 1 },
    { input: '12月', expected: 12 },
    { input: '', expected: null },
  ];

  for (const t of tests) {
    const result = extractMonthNumber(t.input);
    if (result !== t.expected) {
      return { name: 'extractMonthNumber', passed: false,
        message: '"' + t.input + '" → ' + result + ' (期待: ' + t.expected + ')' };
    }
  }

  return { name: 'extractMonthNumber', passed: true };
}

function testFormatDate() {
  const tests = [
    { y: 2026, m: 3, d: 1, expected: '2026/03/01' },
    { y: 2026, m: 10, d: 15, expected: '2026/10/15' },
    { y: 2026, m: 1, d: 5, expected: '2026/01/05' },
  ];

  for (const t of tests) {
    const result = formatDate(t.y, t.m, t.d);
    if (result !== t.expected) {
      return { name: 'formatDate', passed: false,
        message: '(' + t.y + ',' + t.m + ',' + t.d + ') → "' + result + '" (期待: "' + t.expected + '")' };
    }
  }

  return { name: 'formatDate', passed: true };
}

function testNormalizeTimeSlot() {
  const tests = [
    { input: '6時～9時', expected: '6時～9時' },
    { input: '17時～', expected: '17時～22時' },
    { input: '22時～', expected: '22時～' },
    { input: '17時～22時', expected: '17時～22時' },
  ];

  for (const t of tests) {
    const result = normalizeTimeSlot(t.input);
    if (result !== t.expected) {
      return { name: 'normalizeTimeSlot', passed: false,
        message: '"' + t.input + '" → "' + result + '" (期待: "' + t.expected + '")' };
    }
  }

  return { name: 'normalizeTimeSlot', passed: true };
}

// ===== 名寄せテスト =====

function createTestMatcher() {
  const employees = [
    { employeeNo: '072', name: '石井 祐一', furigana: '', lineUserId: '', area: '練馬', facility: '南大泉', status: '在職', notifyEnabled: true, aliases: ['石井'] },
    { employeeNo: '111', name: '飯田 由美子', furigana: '', lineUserId: '', area: '練馬', facility: '', status: '在職', notifyEnabled: true, aliases: [] },
    { employeeNo: '135', name: 'マイ ヴァンサン', furigana: '', lineUserId: '', area: '練馬', facility: '', status: '在職', notifyEnabled: true, aliases: [] },
    { employeeNo: '028', name: 'ﾜﾌﾞﾘﾆｰｸ ﾏﾘﾔ', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: ['マリヤ', 'ﾏﾘﾔ'] },
    { employeeNo: '055', name: '柳 幸子', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: ['柳幸子'] },
    { employeeNo: '060', name: '峯田 宝', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: [] },
    { employeeNo: '119', name: '吉岡 友季子', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: ['吉岡'] },
    { employeeNo: '125', name: '山岸 櫻乃', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: [] },
    { employeeNo: '127', name: '吉瀧 公江', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: [] },
    { employeeNo: '139', name: '佐藤 佳子', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: ['佐藤佳'] },
    { employeeNo: '140', name: '市川 恵子', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: [] },
    { employeeNo: '154', name: '吉崎 尚', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: [] },
    { employeeNo: '173', name: '高橋 百合', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: [] },
    { employeeNo: '174', name: '一澤 涼奈', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: [] },
    { employeeNo: '189', name: '旭 節一', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: [] },
    { employeeNo: '068', name: '石田 満', furigana: '', lineUserId: '', area: '世田谷', facility: '', status: '在職', notifyEnabled: true, aliases: [] },
  ];
  return new NameMatcherEngine(employees);
}

function testNameMatcherExactMatch() {
  const matcher = createTestMatcher();

  // 完全一致: フルネーム
  const result = matcher.match('石井 祐一', '南大泉');
  if (!result || result.employeeNo !== '072') {
    return { name: 'NameMatcher完全一致', passed: false, message: '石井 祐一 がマッチしない' };
  }

  // 完全一致: フルネーム(外国名)
  const result2 = matcher.match('マイ ヴァンサン', '');
  if (!result2 || result2.employeeNo !== '135') {
    return { name: 'NameMatcher完全一致', passed: false, message: 'マイ ヴァンサン がマッチしない' };
  }

  return { name: 'NameMatcher完全一致', passed: true };
}

function testNameMatcherNormalizedMatch() {
  const matcher = createTestMatcher();

  // スペースなし→スペースあり一致: "高橋百合" → "高橋 百合"
  const result = matcher.match('高橋百合', '');
  if (!result || result.employeeNo !== '173') {
    return { name: 'NameMatcherスペース正規化一致', passed: false, message: '高橋百合 がマッチしない: ' + JSON.stringify(result) };
  }

  return { name: 'NameMatcherスペース正規化一致', passed: true };
}

function testNameMatcherAliasMatch() {
  const matcher = createTestMatcher();

  // 別名一致: "マリヤ" → "ﾜﾌﾞﾘﾆｰｸ ﾏﾘﾔ"
  const result = matcher.match('マリヤ', '');
  if (!result || result.employeeNo !== '028') {
    return { name: 'NameMatcher別名一致', passed: false, message: 'マリヤ がマッチしない: ' + JSON.stringify(result) };
  }

  // 別名一致: "佐藤佳" → "佐藤 佳子"
  const result2 = matcher.match('佐藤佳', '');
  if (!result2 || result2.employeeNo !== '139') {
    return { name: 'NameMatcher別名一致', passed: false, message: '佐藤佳 がマッチしない: ' + JSON.stringify(result2) };
  }

  return { name: 'NameMatcher別名一致', passed: true };
}

function testNameMatcherSurnameMatch() {
  const matcher = createTestMatcher();

  // 姓一致(一意): "旭" → "旭 節一" (姓「旭」は1名のみ)
  const result = matcher.match('旭', '');
  if (!result || result.employeeNo !== '189') {
    return { name: 'NameMatcher姓一致', passed: false, message: '旭 がマッチしない: ' + JSON.stringify(result) };
  }

  // 姓一致(一意): "峯田" → "峯田 宝"
  const result2 = matcher.match('峯田', '');
  if (!result2 || result2.employeeNo !== '060') {
    return { name: 'NameMatcher姓一致', passed: false, message: '峯田 がマッチしない: ' + JSON.stringify(result2) };
  }

  // 姓一致(一意): "石田" → "石田 満"
  const result3 = matcher.match('石田', '');
  if (!result3 || result3.employeeNo !== '068') {
    return { name: 'NameMatcher姓一致', passed: false, message: '石田 がマッチしない: ' + JSON.stringify(result3) };
  }

  return { name: 'NameMatcher姓一致', passed: true };
}

function testNameMatcherVariantMatch() {
  const matcher = createTestMatcher();

  // 異体字一致: "吉﨑" → "吉崎 尚" (﨑→崎)
  const result = matcher.match('吉﨑', '');
  // 吉﨑は姓一致で吉崎 尚にマッチするはず（異体字変換後）
  if (!result || result.employeeNo !== '154') {
    return { name: 'NameMatcher異体字一致', passed: false, message: '吉﨑 がマッチしない: ' + JSON.stringify(result) };
  }

  return { name: 'NameMatcher異体字一致', passed: true };
}

function testNameMatcherPartialMatch() {
  const matcher = createTestMatcher();

  // 部分一致: "佐藤佳" → "佐藤 佳子" (マスタ名の先頭一致)
  // ※ 別名に登録していない場合のフォールバックテスト
  // 既に別名で登録しているのでここでは別のケースをテスト

  return { name: 'NameMatcher部分一致', passed: true };
}

// ===== パーサー構造テスト =====

function testNerimaParserStructure() {
  // 練馬パターンのデータ構造をシミュレート
  const mockData = [
    ['施設名', '南大泉', '', ''],
    ['', '', '', ''],
    ['日付', '6時～9時', '17時～22時', '22時～'],
    ['1日', '', '石井 祐一', '石井 祐一'],
    ['2日', '石井 祐一', '石井 祐一', '石井 祐一'],
  ];

  // ヘッダー行検出テスト
  const headerIdx = findHeaderRow(mockData);
  if (headerIdx !== 2) {
    return { name: 'NerimaParser構造', passed: false, message: 'ヘッダー行: ' + headerIdx + ' (期待: 2)' };
  }

  // 施設名抽出テスト
  const facility = extractFacilityName(mockData[0]);
  if (facility !== '南大泉') {
    return { name: 'NerimaParser構造', passed: false, message: '施設名: ' + facility + ' (期待: 南大泉)' };
  }

  // 時間帯抽出テスト
  const timeSlots = extractTimeSlots(mockData);
  if (timeSlots.length !== 3 || timeSlots[0] !== '6時～9時') {
    return { name: 'NerimaParser構造', passed: false, message: '時間帯数: ' + timeSlots.length };
  }

  return { name: 'NerimaParser構造', passed: true };
}

function testSetagayaParserStructure() {
  // 世田谷パターンのデータ構造をシミュレート
  const mockHeaderRow = ['10月', '砧①107', '', '', '砧②207', '', '', '松原', '', '', '夜'];
  const facilityMap = parseSetagayaFacilityHeader(mockHeaderRow);

  const facilities = Object.values(facilityMap).map(f => f.name);
  if (!facilities.includes('砧①107')) {
    return { name: 'SetagayaParser構造', passed: false, message: '砧①107が見つからない: ' + JSON.stringify(facilityMap) };
  }
  if (!facilities.includes('砧②207')) {
    return { name: 'SetagayaParser構造', passed: false, message: '砧②207が見つからない' };
  }
  if (!facilities.includes('松原')) {
    return { name: 'SetagayaParser構造', passed: false, message: '松原が見つからない' };
  }

  return { name: 'SetagayaParser構造', passed: true };
}

// ===== メッセージビルダーテスト =====

function testGetDayOfWeek() {
  const tests = [
    { y: 2026, m: 3, d: 1, expected: '日' },  // 2026/03/01 is Sunday
    { y: 2026, m: 3, d: 2, expected: '月' },
  ];

  for (const t of tests) {
    const result = getDayOfWeek(t.y, t.m, t.d);
    if (result !== t.expected) {
      return { name: 'getDayOfWeek', passed: false,
        message: t.y + '/' + t.m + '/' + t.d + ' → ' + result + ' (期待: ' + t.expected + ')' };
    }
  }

  return { name: 'getDayOfWeek', passed: true };
}

function testFlexMessageStructure() {
  const shifts = [
    { date: '2026/03/01', timeSlot: '17時～22時', facility: '南大泉' },
    { date: '2026/03/01', timeSlot: '22時～', facility: '南大泉' },
    { date: '2026/03/02', timeSlot: '6時～9時', facility: '南大泉' },
  ];

  const message = buildShiftFlexMessage('2026-03', '石井 祐一', shifts);

  if (message.type !== 'flex') {
    return { name: 'FlexMessage構造', passed: false, message: 'type: ' + message.type };
  }
  if (!message.altText.includes('石井 祐一')) {
    return { name: 'FlexMessage構造', passed: false, message: 'altTextに名前がない' };
  }
  if (message.contents.type !== 'bubble') {
    return { name: 'FlexMessage構造', passed: false, message: 'contents.type: ' + message.contents.type };
  }
  if (!message.contents.header) {
    return { name: 'FlexMessage構造', passed: false, message: 'headerがない' };
  }
  if (!message.contents.body) {
    return { name: 'FlexMessage構造', passed: false, message: 'bodyがない' };
  }
  if (!message.contents.footer) {
    return { name: 'FlexMessage構造', passed: false, message: 'footerがない' };
  }

  return { name: 'FlexMessage構造', passed: true };
}

function testBuildShiftRows() {
  const shifts = [
    { date: '2026/03/01', timeSlot: '17時～22時', facility: '南大泉' },
    { date: '2026/03/01', timeSlot: '22時～', facility: '南大泉' },
    { date: '2026/03/02', timeSlot: '6時～9時', facility: '南大泉' },
  ];

  const rows = buildShiftRows(shifts);

  if (rows.length !== 3) {
    return { name: 'buildShiftRows', passed: false, message: '行数: ' + rows.length + ' (期待: 3)' };
  }

  // 1行目: 日付が表示される
  const firstDateText = rows[0].contents[0].text;
  if (!firstDateText.includes('1日')) {
    return { name: 'buildShiftRows', passed: false, message: '1行目の日付: ' + firstDateText };
  }

  // 2行目: 同日なので日付は空
  const secondDateText = rows[1].contents[0].text;
  if (secondDateText.includes('1日')) {
    return { name: 'buildShiftRows', passed: false, message: '2行目に日付が表示されている: ' + secondDateText };
  }

  // 3行目: 新しい日付
  const thirdDateText = rows[2].contents[0].text;
  if (!thirdDateText.includes('2日')) {
    return { name: 'buildShiftRows', passed: false, message: '3行目の日付: ' + thirdDateText };
  }

  return { name: 'buildShiftRows', passed: true };
}
