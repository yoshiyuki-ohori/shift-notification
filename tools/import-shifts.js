#!/usr/bin/env node
/**
 * import-shifts.js
 * 練馬・世田谷のCSVを解析→名寄せ→シフトデータシートに投入→テキスト出力
 */

const fs = require('fs');
const path = require('path');
const https = require('https');

// ===== 設定 =====
const envPath = path.join(__dirname, '..', '.env');
const envContent = fs.readFileSync(envPath, 'utf-8');
const env = {};
envContent.split('\n').forEach(line => {
  const [key, ...vals] = line.split('=');
  if (key && vals.length) env[key.trim()] = vals.join('=').trim();
});
const WEBAPP_URL = env.WEBAPP_URL;
const TARGET_MONTH = '2025-10';

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
      res.on('end', () => { try { resolve(JSON.parse(d)); } catch (e) { reject(new Error(d.substring(0, 300))); } });
    }).on('error', reject);
  });
}

async function writeSheet(sheetName, range, data) {
  const url = `${WEBAPP_URL}?action=write&sheet=${encodeURIComponent(sheetName)}&range=${encodeURIComponent(range)}&data=${encodeURIComponent(JSON.stringify(data))}`;
  if (url.length > 8000) {
    const half = Math.floor(data.length / 2);
    const r1Start = parseInt(range.match(/\d+/)[0]);
    await writeSheet(sheetName, range, data.slice(0, half));
    const r2Start = r1Start + half;
    const colEnd = range.match(/:([A-Z]+)/)[1];
    await writeSheet(sheetName, `A${r2Start}:${colEnd}${r2Start + data.length - half - 1}`, data.slice(half));
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

// ===== Parser functions =====
function extractDayNumber(s) { const m = String(s).match(/^(\d{1,2})/); return m ? parseInt(m[1], 10) : null; }
function formatDate(y, m, d) { return `${y}/${String(m).padStart(2,'0')}/${String(d).padStart(2,'0')}`; }
function normalizeTimeSlot(t) {
  if (!t) return '';
  t = t.trim();
  if (t.match(/\d+時～\d+時/)) return t;
  if (t === '17時～') return '17時～22時';
  return t;
}

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

function parseSetagaya(csvPath, targetMonth) {
  const data = parseCSV(fs.readFileSync(csvPath, 'utf-8'));
  if (data.length < 3) return [];
  // Parse facility header
  const facilityMap = {};
  let curFac = null, curStart = -1;
  for (let c = 1; c < data[0].length; c++) {
    const v = String(data[0][c]).trim();
    if (v) { if (curFac !== null) facilityMap[curStart] = { name: curFac, endCol: c - 1 }; curFac = v; curStart = c; }
  }
  if (curFac !== null) facilityMap[curStart] = { name: curFac, endCol: data[0].length - 1 };
  // Parse time slots
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
    this.employees = employees;
    this.surnameMap = new Map();
    this.aliasMap = new Map();
    for (const emp of employees) {
      if (emp.status === '退職') continue;
      const surname = emp.name.trim().split(/[\s　]+/)[0];
      if (surname) { if (!this.surnameMap.has(surname)) this.surnameMap.set(surname, []); this.surnameMap.get(surname).push(emp); }
      for (const a of emp.aliases) { this.aliasMap.set(a, emp); this.aliasMap.set(this.normalize(a), emp); }
    }
  }
  normalize(n) { return String(n).replace(/　/g,' ').replace(/\s+/g,' ').trim().replace(/﨑/g,'崎').replace(/髙/g,'高').replace(/澤/g,'沢').replace(/櫻/g,'桜').replace(/壽/g,'寿').replace(/惠/g,'恵'); }
  match(name) {
    const t = name.trim(); if (!t) return null;
    // exact
    for (const e of this.employees) { if (e.status!=='退職' && e.name===t) return e; }
    // normalized (no space)
    const ns = this.normalize(t).replace(/\s/g,'');
    for (const e of this.employees) { if (e.status!=='退職' && this.normalize(e.name).replace(/\s/g,'')===ns) return e; }
    // alias
    let m = this.aliasMap.get(t) || this.aliasMap.get(this.normalize(t));
    if (m && m.status!=='退職') return m;
    // variant (﨑→崎 etc)
    const variants = []; const vmap = {'﨑':'崎','崎':'﨑','髙':'高','高':'髙','澤':'沢','沢':'澤'};
    for (let i=0;i<t.length;i++) { if(vmap[t[i]]) variants.push(t.substring(0,i)+vmap[t[i]]+t.substring(i+1)); }
    for (const v of variants) { for (const e of this.employees) { if(e.status!=='退職' && e.name===v) return e; } }
    // surname unique
    const sc = this.surnameMap.get(t);
    if (sc) { const act = sc.filter(e=>e.status!=='退職'); if(act.length===1) return act[0]; }
    // partial
    for (const e of this.employees) { if(e.status!=='退職' && e.name.replace(/\s/g,'').startsWith(ns) && e.name.replace(/\s/g,'').length>ns.length) return e; }
    return null;
  }
}

// ===== Main =====
async function main() {
  // 1. Load employee master from spreadsheet
  console.log('従業員マスタを読み込み中...');
  const masterResp = await fetch(`${WEBAPP_URL}?action=read&sheet=${encodeURIComponent('従業員マスタ')}`, 0);
  const employees = [];
  for (let i = 1; i < masterResp.data.length; i++) {
    const r = masterResp.data[i];
    const no = String(r[0]).trim(); const name = String(r[1]).trim();
    if (!no || !name) continue;
    employees.push({ employeeNo: no.padStart(3, '0'), name, furigana: String(r[2]||''), lineUserId: String(r[3]||''), area: String(r[4]||''), facility: String(r[5]||''), status: String(r[6]||'在職'), notifyEnabled: r[7] !== false && r[7] !== 'FALSE', aliases: String(r[8]||'').split(',').map(a=>a.trim()).filter(a=>a) });
  }
  console.log(`  ${employees.length}名読み込み`);

  // 2. Parse CSVs
  console.log('シフトCSVを解析中...');
  const nerimaPath = path.join(__dirname, '..', '..', 'タイムシートと賃金', 'シフト表', '10月_南大泉_シフト.csv');
  const setagayaPath = path.join(__dirname, '..', '..', 'タイムシートと賃金', '世田谷', '世田谷１年分シフト - シート3.csv');

  const nerimaRecords = parseNerima(nerimaPath, TARGET_MONTH);
  console.log(`  練馬(南大泉): ${nerimaRecords.length}件`);

  const setagayaRecords = parseSetagaya(setagayaPath, TARGET_MONTH);
  console.log(`  世田谷: ${setagayaRecords.length}件`);

  const allRecords = [...nerimaRecords, ...setagayaRecords];

  // 3. Name matching
  console.log('名寄せ実行中...');
  const matcher = new NameMatcher(employees);
  let matched = 0, unmatched = 0;
  const unmatchedNames = new Set();

  for (const rec of allRecords) {
    const emp = matcher.match(rec.originalName);
    if (emp) { rec.employeeNo = emp.employeeNo; rec.formalName = emp.name; matched++; }
    else { rec.employeeNo = ''; rec.formalName = ''; unmatched++; unmatchedNames.add(rec.originalName); }
  }
  console.log(`  マッチ: ${matched}件, 未マッチ: ${unmatched}件`);
  if (unmatchedNames.size > 0) console.log(`  未マッチ名: ${Array.from(unmatchedNames).join(', ')}`);

  // 4. Write to spreadsheet
  console.log('シフトデータをスプレッドシートに書き込み中...');
  const rows = allRecords.map(r => [r.yearMonth, r.date, r.area, r.facility, r.timeSlot, r.originalName, r.employeeNo, r.formalName]);

  const BATCH = 20;
  for (let i = 0; i < rows.length; i += BATCH) {
    const batch = rows.slice(i, i + BATCH);
    const range = `A${i + 2}:H${i + 2 + batch.length - 1}`;
    await writeSheet('シフトデータ', range, batch);
    process.stdout.write('.');
    await new Promise(r => setTimeout(r, 300));
  }
  console.log(` ${rows.length}件書き込み完了`);

  // 5. Aggregate by employee and output text
  console.log('\n========================================');
  console.log(`  ${TARGET_MONTH.replace('-','年')}月 シフト一覧`);
  console.log('========================================\n');

  const byEmployee = new Map();
  for (const rec of allRecords) {
    if (!rec.employeeNo) continue;
    if (!byEmployee.has(rec.employeeNo)) byEmployee.set(rec.employeeNo, { name: rec.formalName, shifts: [] });
    byEmployee.get(rec.employeeNo).shifts.push(rec);
  }

  const days = ['日','月','火','水','木','金','土'];

  for (const [empNo, data] of [...byEmployee.entries()].sort((a,b) => a[0].localeCompare(b[0]))) {
    data.shifts.sort((a,b) => { const dc = a.date.localeCompare(b.date); return dc !== 0 ? dc : a.timeSlot.localeCompare(b.timeSlot); });

    console.log(`■ ${data.name} (No.${empNo})`);
    console.log('─'.repeat(50));

    let prevDate = '';
    for (const s of data.shifts) {
      const [y,m,d] = s.date.split('/').map(Number);
      const dow = days[new Date(y, m-1, d).getDay()];
      const dateLabel = s.date !== prevDate ? `${d}日(${dow})` : '       ';
      prevDate = s.date;
      console.log(`  ${dateLabel.padEnd(10)} ${s.timeSlot.padEnd(12)} ${s.facility}`);
    }
    console.log(`  合計: ${data.shifts.length}件`);
    console.log('');
  }

  // Summary
  console.log('========================================');
  console.log(`  合計: ${byEmployee.size}名, ${allRecords.filter(r=>r.employeeNo).length}件のシフト`);
  if (unmatchedNames.size > 0) {
    console.log(`  未マッチ: ${Array.from(unmatchedNames).join(', ')}`);
  }
  console.log('========================================');
}

main().catch(e => { console.error('Error:', e.message); process.exit(1); });
