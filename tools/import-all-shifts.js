#!/usr/bin/env node
/**
 * import-all-shifts.js
 * 全施設のCSVを一括解析→名寄せ→シフトデータ投入→テキスト出力
 */

const fs = require('fs');
const path = require('path');
const https = require('https');

const envPath = path.join(__dirname, '..', '.env');
const envContent = fs.readFileSync(envPath, 'utf-8');
const env = {};
envContent.split('\n').forEach(line => {
  const [key, ...vals] = line.split('=');
  if (key && vals.length) env[key.trim()] = vals.join('=').trim();
});
const WEBAPP_URL = env.WEBAPP_URL;
const TARGET_MONTH = '2025-10';

// 施設マッピング (CSV施設名 → Firestore施設ID/正式名)
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

// ===== HTTP helper =====
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

async function writeSheet(sheetName, range, data) {
  const url = `${WEBAPP_URL}?action=write&sheet=${encodeURIComponent(sheetName)}&range=${encodeURIComponent(range)}&data=${encodeURIComponent(JSON.stringify(data))}`;
  if (url.length > 7500) {
    const half = Math.floor(data.length / 2);
    const startRow = parseInt(range.match(/(\d+)/)[1]);
    await writeSheet(sheetName, `A${startRow}:H${startRow + half - 1}`, data.slice(0, half));
    await writeSheet(sheetName, `A${startRow + half}:H${startRow + half + data.length - half - 1}`, data.slice(half));
    return;
  }
  return fetch(url, 0);
}

// ===== CSV parser =====
function parseCSV(content) {
  return content.split('\n').map(line => {
    const fields = []; let current = ''; let inQuote = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') inQuote = !inQuote;
      else if (ch === ',' && !inQuote) { fields.push(current); current = ''; }
      else current += ch;
    }
    fields.push(current.replace(/\r$/, ''));
    return fields;
  });
}

function extractDayNumber(s) { const m = String(s).match(/^(\d{1,2})/); return m ? parseInt(m[1], 10) : null; }
function formatDate(y, m, d) { return `${y}/${String(m).padStart(2, '0')}/${String(d).padStart(2, '0')}`; }
function normalizeTimeSlot(t) {
  if (!t) return ''; t = t.trim();
  if (t.match(/\d+時～\d+時/)) return t;
  if (t === '17時～') return '17時～22時';
  return t;
}

// ===== 練馬パーサー =====
function parseNerima(csvPath, targetMonth) {
  const data = parseCSV(fs.readFileSync(csvPath, 'utf-8'));
  if (data.length < 4) return [];
  const facility = (String(data[0][0]).trim() === '施設名' && data[0][1]) ? String(data[0][1]).trim() : '';
  if (!facility) return [];
  let headerIdx = -1;
  for (let i = 0; i < Math.min(data.length, 10); i++) {
    if (String(data[i][0]).trim() === '日付') { headerIdx = i; break; }
  }
  if (headerIdx < 0) return [];
  const timeSlots = [];
  for (let c = 1; c < data[headerIdx].length; c++) { const v = String(data[headerIdx][c]).trim(); if (v) timeSlots.push(v); }
  const [year, month] = targetMonth.split('-').map(Number);
  const records = [];
  for (let i = headerIdx + 1; i < data.length; i++) {
    const dayNum = extractDayNumber(data[i][0]);
    if (!dayNum) continue;
    for (let c = 1; c < data[i].length && c <= timeSlots.length; c++) {
      const name = String(data[i][c]).trim();
      if (!name) continue;
      records.push({ yearMonth: targetMonth, date: formatDate(year, month, dayNum), area: '練馬', facility, timeSlot: timeSlots[c - 1], originalName: name });
    }
  }
  return records;
}

// ===== 世田谷パーサー =====
function parseSetagaya(csvPath, targetMonth) {
  const data = parseCSV(fs.readFileSync(csvPath, 'utf-8'));
  if (data.length < 3) return [];
  const facilityMap = {};
  let curFac = null, curStart = -1;
  for (let c = 1; c < data[0].length; c++) {
    const v = String(data[0][c]).trim();
    if (v) { if (curFac !== null) facilityMap[curStart] = { name: curFac, endCol: c - 1 }; curFac = v; curStart = c; }
  }
  if (curFac !== null) facilityMap[curStart] = { name: curFac, endCol: data[0].length - 1 };
  const entries = Object.entries(facilityMap).map(([c, info]) => ({ startCol: Number(c), ...info })).sort((a, b) => a.startCol - b.startCol);
  const colMap = {};
  for (let c = 1; c < data[1].length; c++) {
    const ts = String(data[1][c]).trim();
    if (!ts) continue;
    let fac = '不明';
    for (const e of entries) { if (c >= e.startCol && c <= e.endCol) { fac = e.name; break; } }
    colMap[c] = { facility: fac, timeSlot: normalizeTimeSlot(ts) };
  }
  const [year, month] = targetMonth.split('-').map(Number);
  const records = [];
  for (let i = 2; i < data.length; i++) {
    const dayNum = extractDayNumber(data[i][0]);
    if (!dayNum) continue;
    for (let c = 1; c < data[i].length; c++) {
      const name = String(data[i][c]).trim();
      if (!name || !colMap[c]) continue;
      records.push({ yearMonth: targetMonth, date: formatDate(year, month, dayNum), area: '世田谷', facility: colMap[c].facility, timeSlot: colMap[c].timeSlot, originalName: name });
    }
  }
  return records;
}

// ===== Name matcher =====
class NameMatcher {
  constructor(employees) {
    this.employees = employees.filter(e => e.status !== '退職');
    this.surnameMap = new Map();
    this.aliasMap = new Map();
    for (const emp of this.employees) {
      const surname = emp.name.trim().split(/[\s　]+/)[0];
      if (surname) { if (!this.surnameMap.has(surname)) this.surnameMap.set(surname, []); this.surnameMap.get(surname).push(emp); }
      for (const a of emp.aliases) { this.aliasMap.set(a, emp); this.aliasMap.set(this.normalize(a), emp); }
    }
  }
  normalize(n) { return String(n).replace(/　/g, ' ').replace(/\s+/g, ' ').trim().replace(/﨑/g, '崎').replace(/髙/g, '高').replace(/澤/g, '沢').replace(/櫻/g, '桜').replace(/壽/g, '寿').replace(/惠/g, '恵'); }
  match(name) {
    const t = name.trim(); if (!t) return null;
    for (const e of this.employees) { if (e.name === t) return e; }
    const ns = this.normalize(t).replace(/\s/g, '');
    for (const e of this.employees) { if (this.normalize(e.name).replace(/\s/g, '') === ns) return e; }
    let m = this.aliasMap.get(t) || this.aliasMap.get(this.normalize(t));
    if (m) return m;
    const vmap = { '﨑': '崎', '崎': '﨑', '髙': '高', '高': '髙', '澤': '沢', '沢': '澤', '惠': '恵', '恵': '惠' };
    for (let i = 0; i < t.length; i++) { if (vmap[t[i]]) { const v = t.substring(0, i) + vmap[t[i]] + t.substring(i + 1); for (const e of this.employees) { if (e.name === v) return e; } } }
    const sc = this.surnameMap.get(t);
    if (sc && sc.length === 1) return sc[0];
    const noSp = t.replace(/\s/g, '');
    if (noSp.length >= 2) { const partial = this.employees.filter(e => e.name.replace(/\s/g, '').startsWith(noSp) && e.name.replace(/\s/g, '').length > noSp.length); if (partial.length === 1) return partial[0]; }
    return null;
  }
}

// ===== Main =====
async function main() {
  // 1. Load employee master
  console.log('従業員マスタを読み込み中...');
  const masterResp = await fetch(`${WEBAPP_URL}?action=read&sheet=${encodeURIComponent('従業員マスタ')}`, 0);
  const employees = [];
  for (let i = 1; i < masterResp.data.length; i++) {
    const r = masterResp.data[i];
    const no = String(r[0]).trim(); const name = String(r[1]).trim();
    if (!no || !name) continue;
    employees.push({ employeeNo: no.padStart(3, '0'), name, furigana: String(r[2] || ''), lineUserId: String(r[3] || ''), area: String(r[4] || ''), facility: String(r[5] || ''), status: String(r[6] || '在職'), notifyEnabled: r[7] !== false && r[7] !== 'FALSE', aliases: String(r[8] || '').split(',').map(a => a.trim()).filter(a => a) });
  }
  console.log(`  ${employees.length}名読み込み`);

  // 2. Parse ALL CSVs
  console.log('\nシフトCSVを解析中...');
  const shiftDir = path.join(__dirname, '..', '..', 'タイムシートと賃金', 'シフト表');
  const setagayaPath = path.join(__dirname, '..', '..', 'タイムシートと賃金', '世田谷', '世田谷１年分シフト - シート3.csv');

  let allRecords = [];

  // 練馬: 全施設
  const nerimaFiles = fs.readdirSync(shiftDir).filter(f => f.endsWith('.csv'));
  console.log(`  練馬: ${nerimaFiles.length}施設`);
  for (const file of nerimaFiles) {
    const records = parseNerima(path.join(shiftDir, file), TARGET_MONTH);
    const facility = records.length > 0 ? records[0].facility : file;
    console.log(`    ${facility}: ${records.length}件`);
    allRecords.push(...records);
  }

  // 世田谷
  const setagayaRecords = parseSetagaya(setagayaPath, TARGET_MONTH);
  console.log(`  世田谷: ${setagayaRecords.length}件`);
  allRecords.push(...setagayaRecords);

  console.log(`  合計: ${allRecords.length}件のシフトレコード`);

  // 3. Name matching
  console.log('\n名寄せ実行中...');
  const matcher = new NameMatcher(employees);
  let matched = 0, unmatched = 0;
  const unmatchedMap = new Map(); // name -> facilities

  for (const rec of allRecords) {
    const emp = matcher.match(rec.originalName);
    if (emp) { rec.employeeNo = emp.employeeNo; rec.formalName = emp.name; matched++; }
    else {
      rec.employeeNo = ''; rec.formalName = ''; unmatched++;
      if (!unmatchedMap.has(rec.originalName)) unmatchedMap.set(rec.originalName, new Set());
      unmatchedMap.get(rec.originalName).add(rec.facility);
    }
  }
  console.log(`  マッチ: ${matched}件, 未マッチ: ${unmatched}件`);
  if (unmatchedMap.size > 0) {
    console.log(`  未マッチ一覧:`);
    for (const [name, facilities] of [...unmatchedMap.entries()].sort()) {
      console.log(`    「${name}」 → ${[...facilities].join(', ')}`);
    }
  }

  // 4. 施設名を正式名に変換
  for (const rec of allRecords) {
    const mapping = FACILITY_MAP[rec.facility];
    if (mapping) {
      rec.facilityId = mapping.id;
      rec.officialFacility = mapping.name;
    } else {
      rec.facilityId = '';
      rec.officialFacility = rec.facility;
    }
  }

  // 5. Write to spreadsheet
  console.log('\nシフトデータをスプレッドシートに書き込み中...');
  const rows = allRecords.map(r => [r.yearMonth, r.date, r.area, r.officialFacility, r.facilityId, r.timeSlot, r.originalName, r.employeeNo, r.formalName]);

  const BATCH = 12;
  for (let i = 0; i < rows.length; i += BATCH) {
    const batch = rows.slice(i, i + BATCH);
    const range = `A${i + 2}:I${i + 2 + batch.length - 1}`;
    await writeSheet('シフトデータ', range, batch);
    process.stdout.write('.');
    await new Promise(r => setTimeout(r, 300));
  }
  console.log(` ${rows.length}件書き込み完了`);

  // 5. Aggregate and output text
  const byEmployee = new Map();
  for (const rec of allRecords) {
    if (!rec.employeeNo) continue;
    if (!byEmployee.has(rec.employeeNo)) byEmployee.set(rec.employeeNo, { name: rec.formalName, shifts: [] });
    byEmployee.get(rec.employeeNo).shifts.push(rec);
  }

  const days = ['日', '月', '火', '水', '木', '金', '土'];

  // テキストファイルにも出力
  let output = '';
  output += `========================================\n`;
  output += `  ${TARGET_MONTH.replace('-', '年')}月 シフト一覧（全施設）\n`;
  output += `  練馬${nerimaFiles.length}施設 + 世田谷\n`;
  output += `  ${byEmployee.size}名, ${allRecords.filter(r => r.employeeNo).length}件\n`;
  output += `========================================\n\n`;

  for (const [empNo, data] of [...byEmployee.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
    data.shifts.sort((a, b) => { const dc = a.date.localeCompare(b.date); return dc !== 0 ? dc : a.timeSlot.localeCompare(b.timeSlot); });

    output += `■ ${data.name} (No.${empNo})\n`;
    output += '─'.repeat(55) + '\n';

    let prevDate = '';
    for (const s of data.shifts) {
      const [y, m, d] = s.date.split('/').map(Number);
      const dow = days[new Date(y, m - 1, d).getDay()];
      const dateLabel = s.date !== prevDate ? `${d}日(${dow})` : '';
      prevDate = s.date;
      output += `  ${dateLabel.padEnd(10)} ${s.timeSlot.padEnd(12)} ${s.officialFacility} (${s.facilityId})\n`;
    }
    output += `  合計: ${data.shifts.length}件\n\n`;
  }

  if (unmatchedMap.size > 0) {
    output += `\n===== 未マッチ =====\n`;
    for (const [name, facilities] of [...unmatchedMap.entries()].sort()) {
      output += `  「${name}」 → ${[...facilities].join(', ')}\n`;
    }
  }

  const outputPath = path.join(__dirname, '..', 'data', 'shift-text-output.txt');
  fs.writeFileSync(outputPath, output, 'utf-8');
  console.log(`\nテキスト出力: ${outputPath}`);

  // サマリーを表示
  console.log('\n' + output.split('\n').slice(0, 6).join('\n'));

  // 施設別・人数サマリー
  const facilityCount = new Map();
  for (const rec of allRecords) {
    const key = `${rec.officialFacility} (${rec.facilityId})`;
    if (!facilityCount.has(key)) facilityCount.set(key, new Set());
    if (rec.employeeNo) facilityCount.get(key).add(rec.employeeNo);
  }
  console.log('\n施設別 人数:');
  for (const [fac, emps] of [...facilityCount.entries()].sort()) {
    console.log(`  ${fac.padEnd(30)} ${emps.size}名`);
  }

  // 人数サマリー
  console.log(`\n総合: ${byEmployee.size}名, ${allRecords.length}件 (未マッチ: ${unmatched}件)`);
}

main().catch(e => { console.error('Error:', e.message); process.exit(1); });
