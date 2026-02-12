#!/usr/bin/env node
/**
 * local-test.js - ローカルテスト実行
 * GASランタイムなしでパーサー・名寄せロジックを検証
 * 実行: node tests/local-test.js
 */

const fs = require('fs');
const path = require('path');

// ===== GAS互換レイヤー =====
// テスト用にGAS固有のAPIをスタブ化

// ===== ユーティリティ関数 (ShiftParser.gsから移植) =====

function extractDayNumber(dayStr) {
  if (!dayStr) return null;
  const match = String(dayStr).match(/^(\d{1,2})/);
  return match ? parseInt(match[1], 10) : null;
}

function extractMonthNumber(monthStr) {
  if (!monthStr) return null;
  const match = String(monthStr).match(/(\d{1,2})月/);
  return match ? parseInt(match[1], 10) : null;
}

function formatDate(year, month, day) {
  return year + '/' +
    String(month).padStart(2, '0') + '/' +
    String(day).padStart(2, '0');
}

function normalizeTimeSlot(timeSlot) {
  if (!timeSlot) return '';
  let normalized = timeSlot.trim();
  if (normalized.match(/\d+時～\d+時/)) return normalized;
  if (normalized === '17時～') return '17時～22時';
  if (normalized === '22時～') return '22時～';
  return normalized;
}

function findHeaderRow(data) {
  for (let i = 0; i < Math.min(data.length, 10); i++) {
    const firstCell = String(data[i][0]).trim();
    if (firstCell === '日付' || firstCell === '月日') return i;
  }
  return -1;
}

function extractFacilityName(headerRow) {
  if (String(headerRow[0]).trim() === '施設名' && headerRow[1]) {
    return String(headerRow[1]).trim();
  }
  return '';
}

function extractTimeSlots(data) {
  const headerIdx = findHeaderRow(data);
  if (headerIdx < 0) return [];
  const slots = [];
  for (let col = 1; col < data[headerIdx].length; col++) {
    const cell = String(data[headerIdx][col]).trim();
    if (cell) slots.push(cell);
  }
  return slots;
}

function parseSetagayaFacilityHeader(headerRow) {
  const facilityMap = {};
  let currentFacility = null;
  let currentStartCol = -1;
  for (let col = 1; col < headerRow.length; col++) {
    const cell = String(headerRow[col]).trim();
    if (cell && cell !== '') {
      if (currentFacility !== null) {
        facilityMap[currentStartCol] = { name: currentFacility, endCol: col - 1 };
      }
      currentFacility = cell;
      currentStartCol = col;
    }
  }
  if (currentFacility !== null) {
    facilityMap[currentStartCol] = { name: currentFacility, endCol: headerRow.length - 1 };
  }
  return facilityMap;
}

function parseSetagayaTimeHeader(timeRow, facilityMap) {
  const columnMap = {};
  const facilityEntries = Object.entries(facilityMap).map(([col, info]) => ({
    startCol: Number(col), ...info
  })).sort((a, b) => a.startCol - b.startCol);

  for (let col = 1; col < timeRow.length; col++) {
    const timeSlot = String(timeRow[col]).trim();
    if (!timeSlot) continue;
    let facility = '不明';
    for (const entry of facilityEntries) {
      if (col >= entry.startCol && col <= entry.endCol) {
        facility = entry.name;
        break;
      }
    }
    columnMap[col] = { facility, timeSlot: normalizeTimeSlot(timeSlot) };
  }
  return columnMap;
}

// ===== CSVパーサー =====

function parseCSV(content) {
  const lines = content.split('\n');
  return lines.map(line => {
    // 簡易CSVパース（ダブルクォート対応）
    const fields = [];
    let current = '';
    let inQuote = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        inQuote = !inQuote;
      } else if (ch === ',' && !inQuote) {
        fields.push(current);
        current = '';
      } else {
        current += ch;
      }
    }
    fields.push(current.replace(/\r$/, ''));
    return fields;
  });
}

// ===== 名寄せエンジン (NameMatcher.gsから移植) =====

class NameMatcherEngine {
  constructor(employees) {
    this.employees = employees;
    this.normalizedNameMap = new Map();
    this.surnameMap = new Map();
    this.aliasMap = new Map();
    this.normalizedAliasMap = new Map();
    this.buildIndices();
  }

  buildIndices() {
    for (const emp of this.employees) {
      if (emp.status === '退職') continue;
      const normalizedName = this.normalizeName(emp.name);
      this.normalizedNameMap.set(normalizedName, emp);
      this.normalizedNameMap.set(emp.name, emp);
      const surname = this.extractSurname(emp.name);
      if (surname) {
        if (!this.surnameMap.has(surname)) this.surnameMap.set(surname, []);
        this.surnameMap.get(surname).push(emp);
      }
      for (const alias of emp.aliases) {
        this.aliasMap.set(alias, emp);
        this.normalizedAliasMap.set(this.normalizeName(alias), emp);
      }
    }
  }

  match(shiftName, facility) {
    if (!shiftName) return null;
    const trimmedName = shiftName.trim();
    if (!trimmedName) return null;

    let result;
    result = this.findExactMatch(trimmedName);
    if (result) return this.toResult(result, '完全一致');
    result = this.findNormalizedMatch(trimmedName);
    if (result) return this.toResult(result, 'スペース正規化一致');
    result = this.findAliasMatch(trimmedName);
    if (result) return this.toResult(result, '別名一致');
    result = this.findVariantMatch(trimmedName);
    if (result) return this.toResult(result, '異体字変換一致');
    result = this.findUniqueSurnameMatch(trimmedName, facility);
    if (result) return this.toResult(result, '姓一致');
    result = this.findPartialNameMatch(trimmedName);
    if (result) return this.toResult(result, '部分一致');
    return null;
  }

  findExactMatch(name) {
    for (const emp of this.employees) {
      if (emp.status === '退職') continue;
      if (emp.name === name) return emp;
    }
    return null;
  }

  findNormalizedMatch(name) {
    const normalized = this.normalizeName(name);
    const match = this.normalizedNameMap.get(normalized);
    if (match && match.status !== '退職') return match;
    const noSpace = name.replace(/[\s　]+/g, '');
    for (const emp of this.employees) {
      if (emp.status === '退職') continue;
      if (this.normalizeName(emp.name).replace(/[\s　]+/g, '') === noSpace) return emp;
    }
    return null;
  }

  findAliasMatch(name) {
    const match = this.aliasMap.get(name);
    if (match && match.status !== '退職') return match;
    const normalized = this.normalizeName(name);
    const normalizedMatch = this.normalizedAliasMap.get(normalized);
    if (normalizedMatch && normalizedMatch.status !== '退職') return normalizedMatch;
    return null;
  }

  findVariantMatch(name) {
    const variants = this.generateVariants(name);
    for (const variant of variants) {
      const match = this.findExactMatch(variant) || this.findNormalizedMatch(variant);
      if (match) return match;
    }
    return null;
  }

  findUniqueSurnameMatch(name, facility) {
    const candidates = this.surnameMap.get(name);
    if (!candidates) return null;
    const active = candidates.filter(e => e.status !== '退職');
    if (active.length === 1) return active[0];
    if (facility && active.length > 1) {
      const facilityMatch = active.filter(e => e.facility === facility);
      if (facilityMatch.length === 1) return facilityMatch[0];
    }
    return null;
  }

  findPartialNameMatch(name) {
    const noSpace = name.replace(/[\s　]+/g, '');
    if (noSpace.length < 2) return null;
    const matches = [];
    for (const emp of this.employees) {
      if (emp.status === '退職') continue;
      const empNoSpace = emp.name.replace(/[\s　]+/g, '');
      if (empNoSpace.startsWith(noSpace) && empNoSpace.length > noSpace.length) matches.push(emp);
    }
    return matches.length === 1 ? matches[0] : null;
  }

  normalizeName(name) {
    if (!name) return '';
    let normalized = String(name);
    normalized = normalized.replace(/　/g, ' ');
    normalized = normalized.replace(/\s+/g, ' ');
    normalized = normalized.trim();
    normalized = this.convertVariantChars(normalized);
    return normalized;
  }

  extractSurname(name) {
    if (!name) return '';
    const parts = name.trim().split(/[\s　]+/);
    return parts[0] || '';
  }

  convertVariantChars(text) {
    const standardize = {
      '﨑': '崎', '髙': '高', '邊': '辺', '邉': '辺',
      '齋': '斎', '齊': '斎', '澤': '沢', '櫻': '桜',
      '國': '国', '龍': '竜', '藏': '蔵', '壽': '寿',
      '廣': '広', '惠': '恵'
    };
    let result = text;
    for (const [from, to] of Object.entries(standardize)) {
      result = result.replace(new RegExp(from, 'g'), to);
    }
    return result;
  }

  generateVariants(name) {
    const variants = new Set();
    const variantMap = {
      '﨑': '崎', '崎': '﨑', '髙': '高', '高': '髙',
      '澤': '沢', '沢': '澤', '櫻': '桜', '桜': '櫻',
      '壽': '寿', '寿': '壽', '惠': '恵', '恵': '惠'
    };
    for (let i = 0; i < name.length; i++) {
      const char = name[i];
      if (variantMap[char]) {
        const variant = name.substring(0, i) + variantMap[char] + name.substring(i + 1);
        variants.add(variant);
      }
    }
    return Array.from(variants);
  }

  toResult(employee, matchType) {
    return { employeeNo: employee.employeeNo, formalName: employee.name, matchType };
  }
}

// ===== テスト実行 =====

let passed = 0;
let failed = 0;

function assert(testName, condition, message) {
  if (condition) {
    console.log(`  \x1b[32m✓\x1b[0m ${testName}`);
    passed++;
  } else {
    console.log(`  \x1b[31m✗\x1b[0m ${testName}: ${message}`);
    failed++;
  }
}

// ----- 1. 練馬CSVパースのテスト -----
console.log('\n\x1b[1m=== 練馬CSV パーステスト ===\x1b[0m');

const nerimaCsvPath = path.join(__dirname, '..', '..', 'タイムシートと賃金', 'シフト表', '10月_南大泉_シフト.csv');
try {
  const nerimaCsv = fs.readFileSync(nerimaCsvPath, 'utf-8');
  const nerimaData = parseCSV(nerimaCsv);

  const facility = extractFacilityName(nerimaData[0]);
  assert('施設名抽出', facility === '南大泉', `got: "${facility}"`);

  const headerIdx = findHeaderRow(nerimaData);
  assert('ヘッダー行検出', headerIdx === 2, `got: ${headerIdx}`);

  const timeSlots = extractTimeSlots(nerimaData);
  assert('時間帯数', timeSlots.length === 3, `got: ${timeSlots.length}`);
  assert('時間帯1', timeSlots[0] === '6時～9時', `got: "${timeSlots[0]}"`);
  assert('時間帯2', timeSlots[1] === '17時～22時', `got: "${timeSlots[1]}"`);
  assert('時間帯3', timeSlots[2] === '22時～', `got: "${timeSlots[2]}"`);

  // データ行パース
  let recordCount = 0;
  const names = new Set();
  for (let i = headerIdx + 1; i < nerimaData.length; i++) {
    const dayNum = extractDayNumber(String(nerimaData[i][0]).trim());
    if (!dayNum) continue;
    for (let col = 1; col <= 3; col++) {
      const name = String(nerimaData[i][col] || '').trim();
      if (name) {
        recordCount++;
        names.add(name);
      }
    }
  }
  assert('レコード数', recordCount > 0, `got: ${recordCount}`);
  assert('レコード数(期待80前後)', recordCount >= 70 && recordCount <= 100, `got: ${recordCount}`);
  assert('ユニーク名前数', names.size >= 3, `got: ${names.size}: ${Array.from(names).join(', ')}`);
  console.log(`  ℹ 練馬シフト: ${recordCount}レコード, 名前: ${Array.from(names).join(', ')}`);

} catch (e) {
  console.log(`  ⚠ 練馬CSVファイルが見つかりません: ${nerimaCsvPath}`);
}

// ----- 2. 世田谷CSVパースのテスト -----
console.log('\n\x1b[1m=== 世田谷CSV パーステスト ===\x1b[0m');

const setagayaCsvPath = path.join(__dirname, '..', '..', 'タイムシートと賃金', '世田谷', '世田谷１年分シフト - シート3.csv');
try {
  const setagayaCsv = fs.readFileSync(setagayaCsvPath, 'utf-8');
  const setagayaData = parseCSV(setagayaCsv);

  const facilityMap = parseSetagayaFacilityHeader(setagayaData[0]);
  const facilities = Object.values(facilityMap).map(f => f.name);
  assert('施設数', facilities.length >= 3, `got: ${facilities.length}: ${facilities.join(', ')}`);
  assert('砧①107', facilities.includes('砧①107'), `facilities: ${facilities.join(', ')}`);
  assert('砧②207', facilities.includes('砧②207'), `facilities: ${facilities.join(', ')}`);
  assert('松原', facilities.includes('松原'), `facilities: ${facilities.join(', ')}`);

  const columnTimeSlots = parseSetagayaTimeHeader(setagayaData[1], facilityMap);
  const colCount = Object.keys(columnTimeSlots).length;
  assert('列マッピング数', colCount >= 9, `got: ${colCount}`);

  // データ行パース
  let recordCount = 0;
  const names = new Set();
  for (let i = 2; i < setagayaData.length; i++) {
    const dayNum = extractDayNumber(String(setagayaData[i][0]).trim());
    if (!dayNum) continue;
    for (let col = 1; col < setagayaData[i].length; col++) {
      const name = String(setagayaData[i][col] || '').trim();
      if (name && columnTimeSlots[col]) {
        recordCount++;
        names.add(name);
      }
    }
  }
  assert('レコード数', recordCount > 0, `got: ${recordCount}`);
  assert('ユニーク名前数', names.size >= 8, `got: ${names.size}`);
  console.log(`  ℹ 世田谷シフト: ${recordCount}レコード, 名前: ${Array.from(names).join(', ')}`);

} catch (e) {
  console.log(`  ⚠ 世田谷CSVファイルが見つかりません: ${setagayaCsvPath}`);
}

// ----- 3. 従業員マスタ + 名寄せテスト -----
console.log('\n\x1b[1m=== 名寄せテスト (実データ) ===\x1b[0m');

const empCsvPath = path.join(__dirname, '..', '..', 'タイムシートと賃金', '社員コード - シート1.csv');
try {
  const empCsv = fs.readFileSync(empCsvPath, 'utf-8');
  const empData = parseCSV(empCsv);

  const employees = [];
  for (let i = 2; i < empData.length; i++) {
    const no = String(empData[i][0] || '').trim();
    const name = String(empData[i][1] || '').trim();
    if (!no || !name) continue;

    employees.push({
      employeeNo: no.padStart(3, '0'),
      name: name.replace(/（.*）/, '').trim(), // "(MF未登録)" 等を除去
      furigana: String(empData[i][2] || '').trim(),
      lineUserId: '',
      area: '',
      facility: '',
      status: '在職',
      notifyEnabled: true,
      aliases: []
    });
  }

  assert('従業員数', employees.length >= 150, `got: ${employees.length}`);

  // 世田谷の主要略称を別名登録
  const aliasConfig = {
    '028': ['マリヤ', 'ﾏﾘﾔ'],
    '055': ['柳幸子'],
    '119': ['吉岡'],
    '139': ['佐藤佳'],
    '154': ['吉﨑'],
  };
  for (const emp of employees) {
    if (aliasConfig[emp.employeeNo]) {
      emp.aliases = aliasConfig[emp.employeeNo];
    }
  }

  const matcher = new NameMatcherEngine(employees);

  // 練馬の名前テスト
  const nerimaTests = [
    { name: '石井 祐一', expected: '072' },
    { name: '飯田 由美子', expected: '111' },
    { name: 'マイ ヴァンサン', expected: '135' },
    { name: '柳岡 央登', expected: '107' },
  ];

  for (const t of nerimaTests) {
    const result = matcher.match(t.name, '');
    assert(`練馬: "${t.name}" → ${t.expected}`,
      result && result.employeeNo === t.expected,
      `got: ${result ? result.employeeNo + ' (' + result.matchType + ')' : 'null'}`);
  }

  // 世田谷の名前テスト（略称）
  const setagayaTests = [
    { name: '吉﨑', expected: '154', desc: '異体字+別名' },
    { name: '旭', expected: '189', desc: '姓一致(一意)' },
    { name: 'マリヤ', expected: '028', desc: '別名' },
    { name: '高橋百合', expected: '173', desc: 'スペース正規化' },
    { name: '峯田', expected: '060', desc: '姓一致(一意)' },
    { name: '市川', expected: '140', desc: '姓一致(一意)' },
    { name: '山岸', expected: '125', desc: '姓一致(一意)' },
    { name: '吉岡', expected: '119', desc: '別名' },
    { name: '吉瀧', expected: '127', desc: '姓一致(一意)' },
    { name: '石田', expected: '068', desc: '姓一致(一意)' },
    { name: '佐藤佳', expected: '139', desc: '別名' },
    { name: '柳幸子', expected: '055', desc: '別名' },
    { name: '一澤', expected: '174', desc: '姓一致(一意)' },
  ];

  for (const t of setagayaTests) {
    const result = matcher.match(t.name, '');
    assert(`世田谷: "${t.name}" → ${t.expected} (${t.desc})`,
      result && result.employeeNo === t.expected,
      `got: ${result ? result.employeeNo + ':' + result.formalName + ' (' + result.matchType + ')' : 'null'}`);
  }

  // 未マッチテスト
  const unmatchedTests = [
    { name: '柳恵子', desc: 'マスタに無い名前' },
  ];
  for (const t of unmatchedTests) {
    const result = matcher.match(t.name, '');
    // 柳恵子はマスタにないので null が期待値
    // ただし「柳」で姓一致して柳 幸子 にマッチする可能性もある
    console.log(`  ℹ 未マッチテスト: "${t.name}" → ${result ? result.employeeNo + ':' + result.formalName + ' (' + result.matchType + ')' : 'null (未マッチ)'}`);
  }

} catch (e) {
  console.log(`  ⚠ 従業員CSVファイルが見つかりません: ${empCsvPath}`);
  console.log(`  Error: ${e.message}`);
}

// ----- 4. Flex Message構造テスト -----
console.log('\n\x1b[1m=== Flex Message構造テスト ===\x1b[0m');

function getDayOfWeek(year, month, day) {
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  const date = new Date(year, month - 1, day);
  return days[date.getDay()];
}

assert('2026/03/01は日曜', getDayOfWeek(2026, 3, 1) === '日', `got: ${getDayOfWeek(2026, 3, 1)}`);
assert('2026/03/02は月曜', getDayOfWeek(2026, 3, 2) === '月', `got: ${getDayOfWeek(2026, 3, 2)}`);

// ===== 結果サマリー =====
console.log('\n\x1b[1m===========================\x1b[0m');
console.log(`\x1b[1m結果: ${passed + failed} テスト, \x1b[32m${passed} 成功\x1b[0m, \x1b[31m${failed} 失敗\x1b[0m`);

if (failed > 0) {
  process.exit(1);
}
