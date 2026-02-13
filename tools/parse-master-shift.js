#!/usr/bin/env node
/**
 * parse-master-shift.js
 * 夜間職員シフト.xlsm を解析 → 全施設のシフトデータを抽出
 *
 * Excel構造:
 *   行1-2: メタデータ（色範囲、NG職員等）
 *   行3: 施設名ヘッダー（4列ごとに: [日付シリアル, 施設名, (空), (空)] or [施設名, 時間1, 時間2, 時間3]）
 *   行4: 時間帯ヘッダー（月日, 6時～9時, 17時～, 22時～）
 *   行5+: シフトデータ（日付シリアル, 職員名, 職員名, 職員名）
 *
 *   左端col0-4: 職員候補リスト（解析対象外）
 *   col5以降: 4列ずつ施設データ [月日, 6時～, 17時～, 22時～]
 */

const path = require('path');
const fs = require('fs');
const https = require('https');
const XLSX = require('xlsx');

const envPath = path.join(__dirname, '..', '.env');
const envContent = fs.readFileSync(envPath, 'utf-8');
const env = {};
envContent.split('\n').forEach(line => {
  const [key, ...vals] = line.split('=');
  if (key && vals.length) env[key.trim()] = vals.join('=').trim();
});
const WEBAPP_URL = env.WEBAPP_URL;

// 施設マッピング (Excel施設名 → Firestore施設ID/正式名)
const FACILITY_MAP = {
  '大泉町': { id: 'GH1', name: '大泉町' },
  'グリーンビレッジB': { id: 'GH3', name: 'グリーンビレッジＢ' },
  '長久保': { id: 'GH4', name: '長久保' },
  '中町': { id: 'GH6', name: '中町' },
  'グリーンビレッジE': { id: 'GH7', name: 'グリーンビレッジＥ' },
  '南大泉': { id: 'GH8', name: '南大泉' },
  '春日町同一①': { id: 'GH9', name: '春日町 (B103)' },
  '春日町': { id: 'GH9', name: '春日町 (B103)' },
  '東大泉': { id: 'GH10', name: '東大泉' },
  '西長久保': { id: 'GH11', name: '西長久保' },
  '春日町２同一①': { id: 'GH12', name: '春日町2 (B203)' },
  '関町南2F同一②': { id: 'GH14', name: '関町南2' },
  '関町南3F同一②': { id: 'GH15', name: '関町南3' },
  '関町南4F同一②': { id: 'GH15', name: '関町南3' },
  '若宮': { id: 'GH18', name: '若宮' },
  '成増5丁目(池ミュー)': { id: 'GH19', name: '成増3 (池ミュハイツ)' },
  '野方': { id: 'GH21', name: '野方' },
  '練馬203同一③': { id: 'GH22', name: '練馬1 (203)' },
  '練馬303同一③': { id: 'GH23', name: '練馬2 (303)' },
  '都民農園': { id: 'GH24', name: '都民農園' },
  '砧①107同一④': { id: 'GH25', name: '砧 (107)' },
  '砧①107': { id: 'GH25', name: '砧 (107)' },
  '天満①(102)同一⑥': { id: 'GH26', name: '天満①' },
  '天満②(302)同一⑥': { id: 'GH27', name: '天満②' },
  '南大泉３丁目': { id: 'GH28', name: '南大泉３丁目' },
  'ビレッジE102': { id: 'GH29', name: 'グリーンビレッジE102' },
  '赤塚３丁目（アーバン）': { id: 'GH30', name: '赤塚3 (第7)' },
  '赤塚５丁目（シティハイム第７）': { id: 'GH31', name: '赤塚2 (アーバン)' },
  '寿': { id: 'GH32', name: '寿' },
  '下井草Ｂ（プラザ阿佐ヶ谷）': { id: 'GH33', name: '下井草' },
  '方南（ﾀﾏｷﾊｲﾂ）': { id: 'GH34', name: '方南' },
  '江古田（ユエヴィ江古田）': { id: 'GH35', name: '江古田 (part1 201)' },
  '砧②207同一④': { id: 'GH36', name: '砧2 (207)' },
  '砧②207': { id: 'GH36', name: '砧2 (207)' },
  '松原': { id: 'GH37', name: '松原' },
  '立川205': { id: 'GH38', name: '立川' },
  '淡路': { id: 'GH39', name: '淡路' },
  '豊島園': { id: 'GH40', name: '豊島園' },
  '石神井公園': { id: 'GH41', name: '石神井公園' },
  'エルウィング208': { id: 'GHT02', name: 'エルウィング' },
  '沼袋': { id: 'GH43', name: '沼袋' },
  '船橋': { id: 'GH44', name: '船橋' },
  '富士見町102': { id: 'GH45', name: '富士見町' },
  '所沢': { id: 'GH46', name: '所沢' },
  'ﾚｼﾞｵﾝ106同一⑤': { id: 'GHsetagaya04', name: '芦花公園' },
  'ｴﾘｰｾﾞ206同一⑤': { id: 'GHsetagaya05', name: '芦花公園2' },
  '津田沼': { id: 'GHF02', name: '津田沼' },
  '芝久保2-201同一⑥': { id: 'GHnishitokyo02', name: '芝久保2' },
  '芝久保3-301同一⑥': { id: 'GHnishitokyo03', name: '芝久保3' },
};

// 施設ベース同姓判別ヒント（CSV検証済み）
// 同姓の複数社員がいる場合、施設名から特定する
const FACILITY_SURNAME_HINTS = {
  'グリーンビレッジＢ': { '青木': '011' },   // 青木 康陽 (CSV検証)
  '芦花公園2': { '青木': '149' },            // 青木 弓恵 (消去法)
};

function excelDateToDate(serial) {
  if (!serial || isNaN(serial)) return null;
  const num = Number(serial);
  if (num < 40000 || num > 50000) return null;
  const d = new Date((num - 25569) * 86400 * 1000);
  return d;
}

function formatDate(d) {
  return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
}

function normalizeTimeSlot(t) {
  if (!t) return '';
  t = t.trim();
  if (t === '17時～') return '17時～22時';
  if (t === '17時～22時') return t;
  if (t === '22時～') return t;
  if (t === '6時～9時') return t;
  return t;
}

function cleanName(name) {
  if (!name) return '';
  let n = String(name).trim();
  // スキップ対象
  if (['空き', '職員配置不要', '配置不要', '欠員', '募集中'].some(s => n.includes(s))) return '';
  // 時間表記を除去 (例: "沼田-8", "青木19-", "マイ16半", "田代-20", "細谷18-", "小川20半-", "伊藤園19-", "向井7半", "星野⁻8半")
  n = n.replace(/[⁻\-ー]\d+半?$/, '');   // 末尾の -数字(半)
  n = n.replace(/^\d+半?[⁻\-ー]/, '');   // 先頭の 数字(半)-
  n = n.replace(/\d+半?[⁻\-ー]$/, '');   // 末尾の 数字半-
  n = n.replace(/\d+半?[⁻\-ー]/, '');    // 中間の 数字半-
  n = n.replace(/\d+半$/, '');           // 末尾の 数字半 (例: "武田16半", "金井16半")
  n = n.replace(/★/g, '');
  n = n.trim();
  if (!n || n === '空き') return '';
  return n;
}

function parseSheet(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) return [];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (data.length < 5) return [];

  // 行3: 施設名ヘッダー
  const row3 = data[2] || [];
  // 行4: 時間帯ヘッダー
  const row4 = data[3] || [];

  // 施設ブロックを検出（4列ごと: 月日列(日付シリアル), タイムスロット1, タイムスロット2, タイムスロット3）
  const facilityBlocks = [];
  for (let c = 5; c < row3.length; c++) {
    const val = String(row3[c]).trim();
    // 日付シリアル番号をスキップ（施設名は数値でない）
    if (!val || !isNaN(val)) continue;
    // 施設名候補
    if (val.length >= 2) {
      // この施設のタイムスロットを確認
      const timeSlots = [];
      for (let t = c; t < Math.min(c + 4, row4.length); t++) {
        const ts = String(row4[t]).trim();
        if (ts && (ts.includes('時') || ts === '月日')) {
          timeSlots.push({ col: t, slot: ts });
        }
      }
      // 月日列の位置を特定（施設名の1列前が月日のはず）
      const dateCol = c - 1;
      facilityBlocks.push({
        name: val,
        dateCol: dateCol,
        timeSlots: timeSlots.filter(t => t.slot !== '月日'),
        startCol: c,
      });
    }
  }

  // シート名から年月を判定
  const [yearStr, monthStr] = sheetName.split('.');
  const year = parseInt(yearStr);
  const month = parseInt(monthStr);
  const yearMonth = `${year}-${String(month).padStart(2, '0')}`;

  // データ行解析
  const records = [];
  for (let i = 4; i < data.length; i++) {
    const row = data[i] || [];

    for (const block of facilityBlocks) {
      // 日付列の値
      const dateSerial = row[block.dateCol];
      const date = excelDateToDate(dateSerial);
      if (!date) continue;
      if (date.getMonth() + 1 !== month) continue; // 対象月以外はスキップ

      const mapping = FACILITY_MAP[block.name];
      const officialName = mapping ? mapping.name : block.name;
      const facilityId = mapping ? mapping.id : '';

      for (const ts of block.timeSlots) {
        const rawName = String(row[ts.col] || '').trim();
        const name = cleanName(rawName);
        if (!name) continue;

        records.push({
          yearMonth,
          date: formatDate(date),
          facility: block.name,
          officialFacility: officialName,
          facilityId,
          timeSlot: normalizeTimeSlot(ts.slot),
          originalName: name,
          rawName: rawName,
        });
      }
    }
  }

  return records;
}

// HTTP helper for spreadsheet API
function fetch(url, n) {
  return new Promise((resolve, reject) => {
    if (n > 5) { reject(new Error('Too many redirects')); return; }
    const u = new URL(url);
    https.get({ hostname: u.hostname, path: u.pathname + u.search }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        resolve(fetch(res.headers.location, n + 1)); return;
      }
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => { try { resolve(JSON.parse(d)); } catch (e) { reject(new Error(d.substring(0, 200))); } });
    }).on('error', reject);
  });
}

// POST JSON to GAS Web App (for large data writes)
function postJSON(url, body) {
  return new Promise((resolve, reject) => {
    const payload = JSON.stringify(body);
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path: u.pathname + u.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload)
      }
    };
    const req = https.request(options, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        // GAS redirects POST responses to a result URL - follow as GET
        resolve(fetch(res.headers.location, 0));
        return;
      }
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => {
        try { resolve(JSON.parse(d)); }
        catch (e) { reject(new Error(d.substring(0, 200))); }
      });
    });
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

// Write shift records to spreadsheet via GAS Web App
async function writeToSpreadsheet(records) {
  const BATCH_SIZE = 50;
  const SHEET_NAME = 'シフトデータ';

  if (!records.length) {
    console.log('書き込み対象のレコードがありません');
    return;
  }

  const targetMonth = records[0].yearMonth;
  console.log(`\nスプレッドシート書き込み開始 (${records.length}件, 対象月: ${targetMonth})...`);

  // Step 1: Clear existing data for target month
  console.log('  既存データをクリア中...');
  try {
    const clearUrl = `${WEBAPP_URL}?action=runFunction&name=clearShiftDataForMonth&arg=${encodeURIComponent(targetMonth)}`;
    await fetch(clearUrl, 0);
    console.log('  クリア完了');
  } catch (e) {
    console.error('  クリアエラー (続行します):', e.message);
  }

  // Step 2: Get current sheet state to determine starting row
  let startRow = 2;
  try {
    const readUrl = `${WEBAPP_URL}?action=read&sheet=${encodeURIComponent(SHEET_NAME)}`;
    const readResult = await fetch(readUrl, 0);
    if (readResult.rows > 0) {
      startRow = readResult.rows + 1;
    }
  } catch (e) {
    console.log('  シート状態取得失敗、行2から開始');
  }

  // Step 3: Write in batches via POST
  const totalBatches = Math.ceil(records.length / BATCH_SIZE);
  let written = 0;
  let errors = 0;

  for (let batch = 0; batch < totalBatches; batch++) {
    const start = batch * BATCH_SIZE;
    const end = Math.min(start + BATCH_SIZE, records.length);
    const batchRecords = records.slice(start, end);

    const rows = batchRecords.map(r => [
      r.yearMonth,
      r.date,
      '',  // エリア (マスタExcel統合のため空)
      r.officialFacility,
      r.facilityId,
      r.timeSlot,
      r.originalName,
      r.employeeNo || '',
      r.formalName || ''
    ]);

    const range = `A${startRow}:I${startRow + rows.length - 1}`;

    try {
      const result = await postJSON(WEBAPP_URL, {
        action: 'write',
        sheet: SHEET_NAME,
        range: range,
        data: rows
      });
      if (result.error) {
        console.error(`\n  バッチ ${batch + 1}/${totalBatches} エラー: ${result.error}`);
        errors++;
      } else {
        written += rows.length;
      }
    } catch (e) {
      console.error(`\n  バッチ ${batch + 1}/${totalBatches} 通信エラー: ${e.message}`);
      errors++;
    }

    startRow += rows.length;
    process.stdout.write(`\r  書き込み中... ${written}/${records.length}件 (バッチ ${batch + 1}/${totalBatches})`);

    // Rate limiting
    if (batch < totalBatches - 1) {
      await new Promise(resolve => setTimeout(resolve, 200));
    }
  }

  console.log(`\n  書き込み完了: ${written}件成功${errors > 0 ? ', ' + errors + '件エラー' : ''}`);
}

// NameMatcher (reused)
class NameMatcher {
  constructor(employees) {
    this.employees = employees.filter(e => e.status !== '退職');
    this.surnameMap = new Map();
    this.aliasMap = new Map();
    for (const emp of this.employees) {
      const surname = emp.name.trim().split(/[\s　]+/)[0];
      if (surname) {
        if (!this.surnameMap.has(surname)) this.surnameMap.set(surname, []);
        this.surnameMap.get(surname).push(emp);
        // Also register normalized surname for variant matching (e.g. 髙松→高松)
        const nSurname = this.normalize(surname);
        if (nSurname !== surname) {
          if (!this.surnameMap.has(nSurname)) this.surnameMap.set(nSurname, []);
          this.surnameMap.get(nSurname).push(emp);
        }
      }
      for (const a of emp.aliases) { this.aliasMap.set(a, emp); this.aliasMap.set(this.normalize(a), emp); }
    }
  }
  normalize(n) {
    return String(n).replace(/　/g, ' ').replace(/\s+/g, ' ').trim()
      .replace(/﨑/g, '崎').replace(/髙/g, '高').replace(/澤/g, '沢')
      .replace(/櫻/g, '桜').replace(/壽/g, '寿').replace(/惠/g, '恵')
      // CJK simplified ↔ traditional variants
      .replace(/张/g, '張').replace(/单/g, '単').replace(/單/g, '単')
      .replace(/华/g, '華').replace(/華/g, '華')
      .replace(/云/g, '雲').replace(/雲/g, '雲')
      .replace(/艳/g, '艶');
  }
  match(name) {
    const t = name.trim(); if (!t) return null;
    for (const e of this.employees) { if (e.name === t) return e; }
    const ns = this.normalize(t).replace(/\s/g, '');
    for (const e of this.employees) { if (this.normalize(e.name).replace(/\s/g, '') === ns) return e; }
    let m = this.aliasMap.get(t) || this.aliasMap.get(this.normalize(t));
    if (m) return m;
    const vmap = { '﨑': '崎', '崎': '﨑', '髙': '高', '高': '髙', '澤': '沢', '沢': '澤', '惠': '恵', '恵': '惠' };
    for (let i = 0; i < t.length; i++) { if (vmap[t[i]]) { const v = t.substring(0, i) + vmap[t[i]] + t.substring(i + 1); for (const e of this.employees) { if (e.name === v) return e; } } }
    const sc = this.surnameMap.get(t) || this.surnameMap.get(this.normalize(t));
    if (sc && sc.length === 1) return sc[0];
    const noSp = t.replace(/\s/g, '');
    if (noSp.length >= 2) { const partial = this.employees.filter(e => e.name.replace(/\s/g, '').startsWith(noSp) && e.name.replace(/\s/g, '').length > noSp.length); if (partial.length === 1) return partial[0]; }
    return null;
  }
  // 施設情報を使った同姓判別（2パス目用）
  matchWithFacility(name, facility, facilityEmployeeMap) {
    const t = name.trim(); if (!t) return null;
    // 静的ヒントで判別
    const hints = FACILITY_SURNAME_HINTS[facility];
    if (hints && hints[t]) {
      const emp = this.employees.find(e => e.employeeNo === hints[t]);
      if (emp) return emp;
    }
    const sc = this.surnameMap.get(t) || this.surnameMap.get(this.normalize(t));
    if (!sc || sc.length <= 1) return null; // 1名以下は通常matchで処理済み
    // この施設で過去にマッチした社員がいるか
    const facilityEmps = facilityEmployeeMap.get(facility) || new Set();
    const candidates = sc.filter(e => facilityEmps.has(e.employeeNo));
    if (candidates.length === 1) return candidates[0];
    return null;
  }
}

async function main() {
  const targetSheet = process.argv[2] || '2026.2';
  const xlsPath = path.join(__dirname, '..', 'data', 'shift-master.xlsm');

  if (!fs.existsSync(xlsPath)) {
    console.error('先に download-shift-excel.js を実行してください');
    process.exit(1);
  }

  console.log(`Excelファイルを読み込み中: ${xlsPath}`);
  const workbook = XLSX.readFile(xlsPath);

  console.log(`シート「${targetSheet}」を解析中...\n`);
  const records = parseSheet(workbook, targetSheet);
  console.log(`  ${records.length}件のシフトレコード\n`);

  // 施設別集計
  const byFacility = new Map();
  for (const r of records) {
    if (!byFacility.has(r.officialFacility)) byFacility.set(r.officialFacility, { id: r.facilityId, records: [] });
    byFacility.get(r.officialFacility).records.push(r);
  }
  console.log(`施設別:`);
  for (const [name, data] of [...byFacility.entries()].sort()) {
    const uniqueNames = new Set(data.records.map(r => r.originalName));
    console.log(`  ${name.padEnd(25)} [${data.id.padEnd(12)}] ${data.records.length}件 (${uniqueNames.size}名)`);
  }

  // 未マッピング施設
  const unmappedFacilities = new Set();
  for (const r of records) { if (!r.facilityId) unmappedFacilities.add(r.facility); }
  if (unmappedFacilities.size > 0) {
    console.log(`\n未マッピング施設: ${[...unmappedFacilities].join(', ')}`);
  }

  // 名寄せ
  console.log('\n従業員マスタを読み込み中...');
  const masterResp = await fetch(`${WEBAPP_URL}?action=read&sheet=${encodeURIComponent('従業員マスタ')}`, 0);
  const employees = [];
  for (let i = 1; i < masterResp.data.length; i++) {
    const r = masterResp.data[i];
    const no = String(r[0]).trim(); const name = String(r[1]).trim();
    if (!no || !name) continue;
    employees.push({ employeeNo: no.padStart(3, '0'), name, status: String(r[6] || '在職'), aliases: String(r[8] || '').split(',').map(a => a.trim()).filter(a => a) });
  }

  const matcher = new NameMatcher(employees);

  // === 1パス目: 通常の名寄せ ===
  let matched = 0, unmatched = 0;
  const unmatchedRecs = [];
  for (const rec of records) {
    const emp = matcher.match(rec.originalName);
    if (emp) { rec.employeeNo = emp.employeeNo; rec.formalName = emp.name; matched++; }
    else { rec.employeeNo = ''; rec.formalName = ''; unmatchedRecs.push(rec); }
  }

  // === 2パス目: 施設ベース判別（同姓複数候補の解決） ===
  // マッチ済みレコードから施設→社員マップを構築
  const facilityEmployeeMap = new Map();
  for (const rec of records) {
    if (!rec.employeeNo) continue;
    if (!facilityEmployeeMap.has(rec.officialFacility)) facilityEmployeeMap.set(rec.officialFacility, new Set());
    facilityEmployeeMap.get(rec.officialFacility).add(rec.employeeNo);
  }

  let pass2Matched = 0;
  for (const rec of unmatchedRecs) {
    const emp = matcher.matchWithFacility(rec.originalName, rec.officialFacility, facilityEmployeeMap);
    if (emp) { rec.employeeNo = emp.employeeNo; rec.formalName = emp.name; matched++; pass2Matched++; }
  }
  unmatched = unmatchedRecs.length - pass2Matched;

  const unmatchedMap = new Map();
  for (const rec of unmatchedRecs) {
    if (rec.employeeNo) continue; // 2パス目で解決済み
    if (!unmatchedMap.has(rec.originalName)) unmatchedMap.set(rec.originalName, new Set());
    unmatchedMap.get(rec.originalName).add(rec.officialFacility);
  }

  console.log(`\n名寄せ結果: マッチ ${matched}件, 未マッチ ${unmatched}件`);
  if (pass2Matched > 0) console.log(`  (施設ベース判別で ${pass2Matched}件追加マッチ)`);
  if (unmatchedMap.size > 0) {
    console.log('未マッチ一覧:');
    for (const [name, facs] of [...unmatchedMap.entries()].sort()) {
      console.log(`  「${name}」 → ${[...facs].join(', ')}`);
    }
  }

  // テキスト出力
  const byEmployee = new Map();
  for (const rec of records) {
    if (!rec.employeeNo) continue;
    if (!byEmployee.has(rec.employeeNo)) byEmployee.set(rec.employeeNo, { name: rec.formalName, shifts: [] });
    byEmployee.get(rec.employeeNo).shifts.push(rec);
  }
  const days = ['日', '月', '火', '水', '木', '金', '土'];

  let output = '';
  const [y, m] = targetSheet.split('.');
  output += `========================================\n`;
  output += `  ${y}年${m}月 シフト一覧（全施設・Excelマスタ）\n`;
  output += `  ${byFacility.size}施設, ${byEmployee.size}名, ${matched}件\n`;
  output += `========================================\n\n`;

  for (const [empNo, data] of [...byEmployee.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
    data.shifts.sort((a, b) => a.date.localeCompare(b.date) || a.timeSlot.localeCompare(b.timeSlot));
    output += `■ ${data.name} (No.${empNo})\n`;
    output += '─'.repeat(65) + '\n';
    let prevDate = '';
    for (const s of data.shifts) {
      const [yy, mm, dd] = s.date.split('/').map(Number);
      const dow = days[new Date(yy, mm - 1, dd).getDay()];
      const dateLabel = s.date !== prevDate ? `${dd}日(${dow})` : '';
      prevDate = s.date;
      output += `  ${dateLabel.padEnd(10)} ${s.timeSlot.padEnd(12)} ${s.officialFacility} [${s.facilityId}]\n`;
    }
    output += `  合計: ${data.shifts.length}件\n\n`;
  }

  output += '========================================\n';
  output += '  施設別 配置人数\n';
  output += '========================================\n';
  for (const [name, data] of [...byFacility.entries()].sort()) {
    const uniqueEmp = new Set(data.records.filter(r => r.employeeNo).map(r => r.employeeNo));
    output += `  ${name.padEnd(25)} [${data.id.padEnd(12)}] ${uniqueEmp.size}名\n`;
  }
  output += `\n総合: ${byEmployee.size}名, ${records.length}件 (未マッチ: ${unmatched}件)\n`;

  const outputPath = path.join(__dirname, '..', 'data', `shift-${targetSheet.replace('.', '-')}.txt`);
  fs.writeFileSync(outputPath, output, 'utf-8');
  console.log(`\nテキスト出力: ${outputPath}`);

  // スプレッドシート書き込み (--write-ss フラグ)
  if (process.argv.includes('--write-ss')) {
    await writeToSpreadsheet(records);
  }
}

main().catch(e => { console.error('Error:', e.message); process.exit(1); });
