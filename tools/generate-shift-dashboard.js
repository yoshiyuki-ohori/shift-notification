#!/usr/bin/env node
/**
 * generate-shift-dashboard.js
 * Excelシフトデータをパース → 集計 → 静的HTMLダッシュボード生成
 *
 * 変更点:
 *   - 利用者人数の表示なし
 *   - 施設詳細ページ（facilities/GHxx.html）を自動生成しリンク
 *   - 職員選択で自分の出勤施設のみ表示
 *
 * Usage: node tools/generate-shift-dashboard.js [sheetName]
 *   例: node tools/generate-shift-dashboard.js 2026.3
 */
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const TARGET_SHEET = process.argv[2] || '2026.3';

// ================================================================
// 施設マッピング (test-line-send.js と同一)
// ================================================================
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

const ID_TO_FACILITY = {};
for (const [, v] of Object.entries(FACILITY_MAP)) {
  ID_TO_FACILITY[v.id] = v.name;
}

let FACILITY_INFO = {};
try {
  FACILITY_INFO = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'data', 'facility-info.json'), 'utf-8'));
} catch (e) { /* なければ空 */ }

// ================================================================
// Excel パーサー
// ================================================================
function excelDateToDate(serial) {
  if (!serial || isNaN(serial)) return null;
  const num = Number(serial);
  if (num < 40000 || num > 50000) return null;
  return new Date((num - 25569) * 86400 * 1000);
}
function formatDate(d) {
  return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
}
function normalizeTimeSlot(t) {
  if (!t) return '';
  t = t.trim();
  if (t === '17時～') return '17時～22時';
  return t;
}
function cleanName(name) {
  if (!name) return '';
  let n = String(name).trim();
  if (['空き', '職員配置不要', '配置不要', '欠員', '募集中'].some(s => n.includes(s))) return '';
  n = n.replace(/[⁻\-ー]\d+半?$/, '').replace(/^\d+半?[⁻\-ー]/, '').replace(/\d+半?[⁻\-ー]$/, '')
       .replace(/\d+半?[⁻\-ー]/, '').replace(/\d+半$/, '').replace(/★/g, '').trim();
  if (!n || n === '空き') return '';
  return n;
}

function parseSheet(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) { console.error(`シート「${sheetName}」が見つかりません`); process.exit(1); }
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (data.length < 5) return [];
  const row3 = data[2] || [], row4 = data[3] || [];

  const facilityBlocks = [];
  for (let c = 5; c < row3.length; c++) {
    const val = String(row3[c]).trim();
    if (!val || !isNaN(val) || val.length < 2) continue;
    const timeSlots = [];
    for (let t = c; t < Math.min(c + 4, row4.length); t++) {
      const ts = String(row4[t]).trim();
      if (ts && (ts.includes('時') || ts === '月日')) timeSlots.push({ col: t, slot: ts });
    }
    facilityBlocks.push({ name: val, dateCol: c - 1, timeSlots: timeSlots.filter(t => t.slot !== '月日'), startCol: c });
  }

  const [yearStr, monthStr] = sheetName.split('.');
  const year = parseInt(yearStr), month = parseInt(monthStr);
  const yearMonth = `${year}-${String(month).padStart(2, '0')}`;

  const records = [];
  for (let i = 4; i < data.length; i++) {
    const row = data[i] || [];
    for (const block of facilityBlocks) {
      const date = excelDateToDate(row[block.dateCol]);
      if (!date || date.getMonth() + 1 !== month) continue;
      const mapping = FACILITY_MAP[block.name];
      const officialName = mapping ? mapping.name : block.name;
      const facilityId = mapping ? mapping.id : '';
      for (const ts of block.timeSlots) {
        const name = cleanName(String(row[ts.col] || ''));
        if (!name) continue;
        records.push({ yearMonth, date: formatDate(date), facility: officialName, facilityId, timeSlot: normalizeTimeSlot(ts.slot), originalName: name, excelFacility: block.name });
      }
    }
  }
  return records;
}

// ================================================================
// 名寄せ
// ================================================================
function loadMasterTSV() {
  const tsvPath = path.join(__dirname, '..', 'data', 'employee-master.tsv');
  const lines = fs.readFileSync(tsvPath, 'utf-8').split('\n');
  const employees = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split('\t');
    if (cols.length < 2) continue;
    const no = cols[0].trim(), name = cols[1].trim();
    if (!no || !name) continue;
    employees.push({
      employeeNo: no.padStart(3, '0'), name,
      status: (cols[4] || '在職').trim(),
      aliases: (cols[6] || '').split(',').map(a => a.trim()).filter(a => a)
    });
  }
  return employees;
}
function matchName(name, employees) {
  const t = name.trim(); if (!t) return null;
  for (const e of employees) { if (e.status !== '退職' && e.name === t) return e; }
  const ns = t.replace(/　/g, ' ').replace(/\s+/g, '').replace(/﨑/g, '崎').replace(/髙/g, '高');
  for (const e of employees) {
    const en = e.name.replace(/　/g, ' ').replace(/\s+/g, '').replace(/﨑/g, '崎').replace(/髙/g, '高');
    if (e.status !== '退職' && en === ns) return e;
  }
  for (const e of employees) { if (e.status !== '退職' && e.aliases.includes(t)) return e; }
  const surnameMatches = employees.filter(e => e.status !== '退職' && e.name.split(/[\s　]+/)[0] === t);
  if (surnameMatches.length === 1) return surnameMatches[0];
  return null;
}

// ================================================================
// 時間帯統合
// ================================================================
function mergeOvernightShifts(shifts) {
  shifts.sort((a, b) => a.date.localeCompare(b.date) || getTimeOrder(a.timeSlot) - getTimeOrder(b.timeSlot));
  const merged = [];
  const consumed = new Set();
  for (let i = 0; i < shifts.length; i++) {
    if (consumed.has(i)) continue;
    const s = shifts[i];
    if (s.timeSlot === '17時～22時') {
      const nightIdx = shifts.findIndex((x, j) => j > i && !consumed.has(j) &&
        x.date === s.date && x.facility === s.facility &&
        (x.timeSlot === '22時～' || x.timeSlot === '21時～'));
      if (nightIdx >= 0) {
        const nextDate = getNextDate(s.date);
        const morningIdx = shifts.findIndex((x, j) => j > i && !consumed.has(j) &&
          x.date === nextDate && x.facility === s.facility &&
          x.timeSlot === '6時～9時');
        consumed.add(i);
        consumed.add(nightIdx);
        if (morningIdx >= 0) {
          consumed.add(morningIdx);
          merged.push({ ...s, timeSlot: '17時～翌9時', isMerged: true });
        } else {
          merged.push({ ...s, timeSlot: '17時～翌朝', isMerged: true });
        }
        continue;
      }
    }
    if (s.timeSlot === '22時～' || s.timeSlot === '21時～') {
      const nextDate = getNextDate(s.date);
      const morningIdx = shifts.findIndex((x, j) => j > i && !consumed.has(j) &&
        x.date === nextDate && x.facility === s.facility &&
        x.timeSlot === '6時～9時');
      if (morningIdx >= 0) {
        consumed.add(i);
        consumed.add(morningIdx);
        const startHour = s.timeSlot.match(/^(\d+)時/)[1];
        merged.push({ ...s, timeSlot: `${startHour}時～翌9時`, isMerged: true });
        continue;
      }
    }
    if (!consumed.has(i)) {
      merged.push({ ...s, isMerged: false });
    }
  }
  merged.sort((a, b) => a.date.localeCompare(b.date) || getTimeOrder(a.timeSlot) - getTimeOrder(b.timeSlot));
  return merged;
}
function getTimeOrder(ts) {
  if (ts.startsWith('6')) return 1;
  if (ts.startsWith('17')) return 3;
  if (ts.startsWith('21')) return 4;
  if (ts.startsWith('22')) return 5;
  return 9;
}
function getNextDate(dateStr) {
  const [y, m, d] = dateStr.split('/').map(Number);
  const dt = new Date(y, m - 1, d + 1);
  return `${dt.getFullYear()}/${String(dt.getMonth() + 1).padStart(2, '0')}/${String(dt.getDate()).padStart(2, '0')}`;
}

// ================================================================
// HTML生成ヘルパー
// ================================================================
function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function getDayOfWeek(y, m, d) {
  return ['日', '月', '火', '水', '木', '金', '土'][new Date(y, m - 1, d).getDay()];
}

function getTimeSlotColor(ts) {
  if (ts.includes('翌')) return '#C0392B';
  if (ts.match(/^6時/)) return '#E67E22';
  if (ts.match(/^17時/)) return '#8E44AD';
  if (ts.match(/^22時|^21時/)) return '#2C3E80';
  return '#555555';
}

function getTimeSlotBg(ts) {
  if (ts.includes('翌')) return '#FDEDEC';
  if (ts.match(/^6時/)) return '#FEF5E7';
  if (ts.match(/^17時/)) return '#F4ECF7';
  if (ts.match(/^22時|^21時/)) return '#EBF5FB';
  return '#F8F9FA';
}

// ================================================================
// 施設詳細ページ生成（モバイル対応）
// ================================================================
function generateFacilityPage(fid, fname, info, calData, year, month, daysInMonth, records, generatedAt) {
  const mapUrl = info && info.mapQuery
    ? `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(info.mapQuery)}`
    : null;
  const room = info ? info.room : '';

  const facilityRecords = records.filter(r => r.facilityId === fid);
  const byDay = {};
  for (const r of facilityRecords) {
    const day = parseInt(r.date.split('/')[2]);
    if (!byDay[day]) byDay[day] = [];
    byDay[day].push(r);
  }
  const staffNames = [...new Set(facilityRecords.map(r => r.formalName || r.originalName))].sort();

  let shiftCards = '';
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = getDayOfWeek(year, month, d);
    const dowIdx = new Date(year, month - 1, d).getDay();
    const dayColor = dowIdx === 0 ? '#E74C3C' : dowIdx === 6 ? '#2980B9' : '#333';
    const shifts = byDay[d] || [];
    const bgClass = shifts.length === 0 ? ' style="background:#FFF5F5;"' : '';
    let shiftHtml = '';
    if (shifts.length === 0) {
      shiftHtml = '<div style="color:#CCC;font-size:13px;padding:4px 0;">配置なし</div>';
    } else {
      for (const s of shifts) {
        shiftHtml += `<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid #F5F5F5;">
          <span style="color:${getTimeSlotColor(s.timeSlot)};font-weight:700;font-size:14px;min-width:90px;">${esc(s.timeSlot)}</span>
          <span style="font-size:14px;">${esc(s.formalName || s.originalName)}</span>
        </div>`;
      }
    }
    shiftCards += `<div class="day-card"${bgClass}>
      <div class="day-card-header" style="color:${dayColor};">${month}/${d}(${dow})</div>
      ${shiftHtml}
    </div>\n`;
  }

  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>${esc(fname)} - 施設詳細</title>
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Hiragino Sans', 'Noto Sans JP', sans-serif; background: #F0F2F5; color: #333; font-size: 15px; line-height: 1.6; -webkit-text-size-adjust: 100%; }
.header { background: linear-gradient(135deg, #1DB446, #00B900); color: #fff; padding: 14px 16px; box-shadow: 0 2px 8px rgba(0,0,0,.15); position: sticky; top: 0; z-index: 10; }
.header h1 { font-size: 18px; font-weight: 700; }
.header .sub { font-size: 12px; opacity: .85; margin-top: 2px; }
.back-link { display: block; margin: 12px; padding: 14px 16px; background: #fff; border-radius: 10px; text-decoration: none; color: #1DB446; font-weight: 700; font-size: 15px; box-shadow: 0 1px 4px rgba(0,0,0,.08); text-align: center; -webkit-tap-highlight-color: transparent; }
.container { max-width: 600px; margin: 0 auto; padding: 0 12px 40px; }
.card { background: #fff; border-radius: 12px; padding: 16px; margin-bottom: 12px; box-shadow: 0 1px 4px rgba(0,0,0,.08); }
.card h2 { font-size: 15px; color: #555; margin-bottom: 10px; border-bottom: 2px solid #E0E0E0; padding-bottom: 6px; }
.info-row { display: flex; padding: 6px 0; border-bottom: 1px solid #F5F5F5; }
.info-row dt { font-weight: 700; color: #888; font-size: 13px; width: 100px; flex-shrink: 0; }
.info-row dd { font-size: 15px; }
.map-btn { display: block; margin-top: 12px; padding: 14px; background: #1DB446; color: #fff; border-radius: 10px; text-decoration: none; font-weight: 700; font-size: 15px; text-align: center; -webkit-tap-highlight-color: transparent; }
.staff-chips { display: flex; flex-wrap: wrap; gap: 8px; }
.staff-chip { background: #F0F2F5; padding: 6px 14px; border-radius: 20px; font-size: 14px; }
.day-card { background: #fff; border-radius: 10px; padding: 12px; margin-bottom: 8px; box-shadow: 0 1px 3px rgba(0,0,0,.06); }
.day-card-header { font-weight: 800; font-size: 16px; padding-bottom: 6px; border-bottom: 2px solid #F0F0F0; margin-bottom: 4px; }
.placeholder { background: #FFFDE7; border: 1px dashed #FBC02D; border-radius: 8px; padding: 14px; color: #F57F17; font-size: 13px; }
</style>
</head>
<body>
<div class="header">
  <h1>${esc(fname)}</h1>
  <div class="sub">${esc(fid)} | ${year}年${month}月</div>
</div>
<a class="back-link" href="../shift-dashboard-${year}-${String(month).padStart(2,'0')}.html">&larr; ダッシュボードに戻る</a>
<div class="container">

<div class="card">
  <h2>施設情報</h2>
  <div class="info-row"><dt>施設ID</dt><dd>${esc(fid)}</dd></div>
  <div class="info-row"><dt>施設名</dt><dd>${esc(fname)}</dd></div>
  ${room ? `<div class="info-row"><dt>部屋</dt><dd>${esc(room)}</dd></div>` : ''}
  ${mapUrl ? `<a class="map-btn" href="${mapUrl}" target="_blank" rel="noopener">Google Maps で開く</a>` : ''}
</div>

<div class="card">
  <h2>利用者特性</h2>
  <div class="placeholder">
    この施設の利用者特性情報は今後追加予定です。<br>
    Firestoreの care_users コレクションからデータを取得して表示します。
  </div>
</div>

<div class="card">
  <h2>配置職員（${staffNames.length}名）</h2>
  <div class="staff-chips">
    ${staffNames.map(n => `<span class="staff-chip">${esc(n)}</span>`).join('')}
  </div>
</div>

<div class="card">
  <h2>シフト一覧</h2>
</div>
${shiftCards}

</div>
<div style="text-align:center;padding:20px;color:#999;font-size:11px;">生成: ${esc(generatedAt)}</div>
</body>
</html>`;
}

// ================================================================
// Main
// ================================================================
function main() {
  const [yearStr, monthStr] = TARGET_SHEET.split('.');
  const year = parseInt(yearStr), month = parseInt(monthStr);
  const yearMonth = `${year}-${String(month).padStart(2, '0')}`;
  const daysInMonth = new Date(year, month, 0).getDate();

  // 1. Parse Excel
  const xlsPath = path.join(__dirname, '..', 'data', 'shift-master.xlsm');
  console.log(`Excelパース中 (${TARGET_SHEET})...`);
  const workbook = XLSX.readFile(xlsPath);
  const records = parseSheet(workbook, TARGET_SHEET);
  console.log(`  ${records.length}件のシフトレコード`);

  // 2. Name matching
  const employees = loadMasterTSV();
  console.log(`  従業員マスタ: ${employees.length}名`);
  const unmatchedNames = new Set();
  const unmappedFacilities = new Set();
  let matched = 0;
  for (const rec of records) {
    const emp = matchName(rec.originalName, employees);
    if (emp) {
      rec.employeeNo = emp.employeeNo;
      rec.formalName = emp.name;
      matched++;
    } else {
      unmatchedNames.add(rec.originalName);
    }
    if (!rec.facilityId) {
      unmappedFacilities.add(rec.excelFacility || rec.facility);
    }
  }
  const matchRate = records.length > 0 ? ((matched / records.length) * 100).toFixed(1) : 0;
  console.log(`  名寄せ: ${matched}/${records.length}件マッチ (${matchRate}%)`);

  // 3. Group by employee and merge overnight shifts
  const byEmployee = new Map();
  for (const rec of records) {
    const key = rec.employeeNo || `_unmatched_${rec.originalName}`;
    const name = rec.formalName || rec.originalName;
    if (!byEmployee.has(key)) byEmployee.set(key, { name, employeeNo: rec.employeeNo || '', shifts: [] });
    byEmployee.get(key).shifts.push(rec);
  }
  for (const [, emp] of byEmployee) {
    emp.mergedShifts = mergeOvernightShifts([...emp.shifts]);
  }

  // 4. Aggregate data

  // Facility IDs & names
  const facilityIds = [...new Set(records.filter(r => r.facilityId).map(r => r.facilityId))].sort();
  const facilityNames = {};
  for (const r of records) {
    if (r.facilityId) facilityNames[r.facilityId] = r.facility;
  }

  // Facility calendar: facilityId → day → [{timeSlot, name}]
  const facilityCalendar = {};
  for (const r of records) {
    if (!r.facilityId) continue;
    if (!facilityCalendar[r.facilityId]) facilityCalendar[r.facilityId] = {};
    const day = parseInt(r.date.split('/')[2]);
    if (!facilityCalendar[r.facilityId][day]) facilityCalendar[r.facilityId][day] = [];
    facilityCalendar[r.facilityId][day].push({
      timeSlot: r.timeSlot,
      name: r.formalName || r.originalName
    });
  }

  // Staff list
  const staffList = [];
  // staffFacilityIds: empNo → Set of facilityIds (for JS filtering)
  const staffFacilityMap = {};
  for (const [key, emp] of byEmployee) {
    if (key.startsWith('_unmatched_')) continue;
    const facIds = new Set(emp.shifts.filter(s => s.facilityId).map(s => s.facilityId));
    const facilities = new Set(emp.mergedShifts.map(s => s.facility));
    const days = new Set(emp.mergedShifts.map(s => s.date));
    staffList.push({
      employeeNo: emp.employeeNo,
      name: emp.name,
      shiftCount: emp.mergedShifts.length,
      dayCount: days.size,
      facilityCount: facilities.size,
      facilities: [...facilities].sort(),
      facilityIds: [...facIds]
    });
    staffFacilityMap[emp.employeeNo] = [...facIds];
  }
  staffList.sort((a, b) => a.employeeNo.localeCompare(b.employeeNo));

  // Daily coverage
  const dailyCoverage = {};
  for (let d = 1; d <= daysInMonth; d++) {
    dailyCoverage[d] = { shifts: 0, facilities: new Set(), staff: new Set() };
  }
  for (const r of records) {
    const day = parseInt(r.date.split('/')[2]);
    if (dailyCoverage[day]) {
      dailyCoverage[day].shifts++;
      if (r.facilityId) dailyCoverage[day].facilities.add(r.facilityId);
      dailyCoverage[day].staff.add(r.formalName || r.originalName);
    }
  }

  // KPIs
  const totalShifts = records.length;
  const uniqueStaff = new Set(records.map(r => r.formalName || r.originalName)).size;
  const uniqueFacilities = facilityIds.length;
  const avgPerDay = (totalShifts / daysInMonth).toFixed(1);
  const totalFacilitiesInInfo = Object.keys(FACILITY_INFO).length;

  // ================================================================
  // 5. 施設詳細ページ生成
  // ================================================================
  const generatedAt = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });
  const facilitiesDir = path.join(__dirname, '..', 'data', 'facilities');
  if (!fs.existsSync(facilitiesDir)) fs.mkdirSync(facilitiesDir, { recursive: true });

  let facilityPagesGenerated = 0;
  for (const fid of Object.keys(FACILITY_INFO)) {
    const info = FACILITY_INFO[fid];
    const fname = info.name || fid;
    const calData = facilityCalendar[fid] || {};
    const pageHtml = generateFacilityPage(fid, fname, info, calData, year, month, daysInMonth, records, generatedAt);
    fs.writeFileSync(path.join(facilitiesDir, `${fid}.html`), pageHtml, 'utf-8');
    facilityPagesGenerated++;
  }
  console.log(`  施設詳細ページ: ${facilityPagesGenerated}件生成`);

  // ================================================================
  // 6. ダッシュボードHTML生成（モバイルファースト）
  // ================================================================

  // 6. Load preference data if available
  let prefData = [];
  const prefTsvPath = path.join(__dirname, '..', 'data', 'shift-preferences.tsv');
  if (fs.existsSync(prefTsvPath)) {
    const prefLines = fs.readFileSync(prefTsvPath, 'utf-8').split('\n');
    for (let i = 1; i < prefLines.length; i++) {
      const cols = prefLines[i].split('\t');
      if (cols.length < 5) continue;
      const ym = cols[0].trim();
      if (ym !== yearMonth) continue;
      prefData.push({
        yearMonth: ym,
        employeeNo: cols[1].trim().padStart(3, '0'),
        name: cols[2].trim(),
        date: cols[3].trim(),
        type: cols[4].trim(),
        timeSlot: (cols[5] || '').trim(),
        reason: (cols[6] || '').trim(),
        submittedAt: (cols[7] || '').trim()
      });
    }
    console.log(`  シフト希望: ${prefData.length}件読み込み`);
  }

  // Aggregate preference data
  const prefByEmployee = {};
  const prefByDate = {};
  for (const p of prefData) {
    if (!prefByEmployee[p.employeeNo]) {
      prefByEmployee[p.employeeNo] = { name: p.name, prefs: [], submittedAt: p.submittedAt };
    }
    prefByEmployee[p.employeeNo].prefs.push(p);
    if (p.submittedAt > prefByEmployee[p.employeeNo].submittedAt) {
      prefByEmployee[p.employeeNo].submittedAt = p.submittedAt;
    }

    if (!prefByDate[p.date]) prefByDate[p.date] = { want: 0, ng: 0, either: 0 };
    if (p.type === '希望') prefByDate[p.date].want++;
    else if (p.type === 'NG') prefByDate[p.date].ng++;
    else prefByDate[p.date].either++;
  }

  // Submission status
  const submittedEmpNos = new Set(Object.keys(prefByEmployee));
  const activeEmployees = employees.filter(e => e.status !== '退職');
  const notSubmitted = activeEmployees.filter(e => !submittedEmpNos.has(e.employeeNo));

  // Staff shift JSON for client-side personal view
  const staffShiftsData = {};
  for (const [key, emp] of byEmployee) {
    if (key.startsWith('_unmatched_')) continue;
    staffShiftsData[emp.employeeNo] = {
      name: emp.name,
      facilityIds: [...new Set(emp.shifts.filter(s => s.facilityId).map(s => s.facilityId))],
      shifts: emp.mergedShifts.map(s => ({
        d: s.date, t: s.timeSlot, f: s.facility, fid: s.facilityId || ''
      }))
    };
  }
  const staffShiftsJson = JSON.stringify(staffShiftsData);
  const staffFacilityJson = JSON.stringify(staffFacilityMap);

  const staffOptions = staffList.map(s =>
    `<option value="${esc(s.employeeNo)}">${esc(s.name)}</option>`
  ).join('\n        ');

  const html = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>${year}年${month}月 シフト管理ダッシュボード</title>
<style>
/* ===== Reset & Base (Mobile First) ===== */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Hiragino Sans', 'Noto Sans JP', sans-serif; background: #F0F2F5; color: #333; font-size: 15px; line-height: 1.5; -webkit-text-size-adjust: 100%; }

/* ===== Header ===== */
.header { background: linear-gradient(135deg, #1DB446, #00B900); color: #fff; padding: 12px 16px; box-shadow: 0 2px 8px rgba(0,0,0,.15); position: sticky; top: 0; z-index: 100; }
.header h1 { font-size: 17px; font-weight: 700; }
.header .meta { font-size: 11px; opacity: .8; }

/* ===== Staff Selector ===== */
.staff-selector { background: linear-gradient(135deg, #15912f, #0a7a1e); padding: 10px 16px; display: flex; align-items: center; gap: 10px; position: sticky; top: 46px; z-index: 100; }
.staff-selector select { flex: 1; padding: 10px 12px; border: none; border-radius: 10px; font-size: 15px; background: #fff; color: #333; -webkit-appearance: none; appearance: none; background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%23666' d='M6 8L1 3h10z'/%3E%3C/svg%3E"); background-repeat: no-repeat; background-position: right 12px center; min-height: 44px; }
.staff-selector .mode-label { font-size: 11px; padding: 4px 10px; border-radius: 12px; font-weight: 700; white-space: nowrap; }
.mode-admin { background: #FFD54F; color: #5D4037; }
.mode-staff { background: #fff; color: #1DB446; }

/* ===== Tabs ===== */
.tab-bar { display: flex; background: #fff; border-bottom: 2px solid #E0E0E0; overflow-x: auto; -webkit-overflow-scrolling: touch; position: sticky; top: 112px; z-index: 99; }
.tab-btn { flex-shrink: 0; padding: 12px 16px; cursor: pointer; border: none; background: none; font-size: 14px; font-weight: 600; color: #888; border-bottom: 3px solid transparent; transition: all .2s; white-space: nowrap; -webkit-tap-highlight-color: transparent; min-height: 44px; }
.tab-btn.active { color: #1DB446; border-bottom-color: #1DB446; }

/* ===== Tab Content ===== */
.tab-content { display: none; padding: 12px; max-width: 900px; margin: 0 auto; }
.tab-content.active { display: block; }

/* ===== KPI Cards ===== */
.kpi-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; margin-bottom: 16px; }
.kpi-card { background: #fff; border-radius: 12px; padding: 14px; box-shadow: 0 1px 4px rgba(0,0,0,.08); text-align: center; }
.kpi-card .value { font-size: 28px; font-weight: 800; color: #1DB446; }
.kpi-card .label { font-size: 12px; color: #888; margin-top: 2px; }
.kpi-card.warn .value { color: #E67E22; }
.kpi-card.danger .value { color: #E74C3C; }

/* ===== Personal Shift View (マイシフト) ===== */
#my-shift-view { padding: 12px; max-width: 900px; margin: 0 auto; }
.my-header { background: #fff; border-radius: 12px; padding: 16px; margin-bottom: 12px; box-shadow: 0 1px 4px rgba(0,0,0,.08); text-align: center; }
.my-header .my-name { font-size: 20px; font-weight: 800; color: #1DB446; }
.my-header .my-stats { display: flex; justify-content: center; gap: 20px; margin-top: 8px; }
.my-header .my-stat { text-align: center; }
.my-header .my-stat-val { font-size: 24px; font-weight: 800; color: #333; }
.my-header .my-stat-label { font-size: 11px; color: #888; }
.my-day-card { background: #fff; border-radius: 12px; margin-bottom: 8px; box-shadow: 0 1px 3px rgba(0,0,0,.06); overflow: hidden; }
.my-day-header { padding: 10px 16px; font-weight: 800; font-size: 15px; border-bottom: 1px solid #F0F0F0; display: flex; justify-content: space-between; align-items: center; }
.my-day-header.sun { background: #FFF5F5; color: #E74C3C; }
.my-day-header.sat { background: #F0F0FF; color: #2980B9; }
.my-shift-item { padding: 12px 16px; display: flex; align-items: center; gap: 12px; border-bottom: 1px solid #F8F8F8; -webkit-tap-highlight-color: transparent; }
.my-shift-item:last-child { border-bottom: none; }
.my-time-badge { padding: 4px 10px; border-radius: 8px; font-weight: 700; font-size: 13px; white-space: nowrap; }
.my-facility-name { font-size: 15px; font-weight: 600; flex: 1; }
.my-facility-link { color: #1DB446; text-decoration: none; font-size: 13px; font-weight: 600; padding: 8px 12px; border: 1px solid #1DB446; border-radius: 8px; white-space: nowrap; min-height: 36px; display: flex; align-items: center; -webkit-tap-highlight-color: transparent; }

/* ===== Filter ===== */
.filter-bar { margin-bottom: 12px; }
.filter-bar input { width: 100%; padding: 12px 14px; border: 1px solid #D0D0D0; border-radius: 10px; font-size: 15px; min-height: 44px; }
.filter-bar input:focus { outline: none; border-color: #1DB446; box-shadow: 0 0 0 2px rgba(29,180,70,.2); }

/* ===== Facility Section (Admin) ===== */
.facility-section { background: #fff; border-radius: 12px; margin-bottom: 10px; box-shadow: 0 1px 4px rgba(0,0,0,.08); overflow: hidden; }
.facility-section summary { padding: 14px 16px; cursor: pointer; font-weight: 700; font-size: 15px; background: #FAFAFA; border-bottom: 1px solid #EEE; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; min-height: 52px; -webkit-tap-highlight-color: transparent; }
.facility-section summary .badge { background: #1DB446; color: #fff; font-size: 11px; padding: 2px 8px; border-radius: 10px; font-weight: 600; }
.facility-section summary .info { font-size: 12px; color: #888; font-weight: 400; }
.facility-toolbar { padding: 8px 16px; display: flex; justify-content: flex-end; background: #FAFAFA; border-bottom: 1px solid #EEE; }
.facility-toolbar .detail-link { font-size: 13px; color: #1DB446; text-decoration: none; font-weight: 600; padding: 6px 12px; border: 1px solid #1DB446; border-radius: 8px; white-space: nowrap; -webkit-tap-highlight-color: transparent; }

/* ===== Calendar Grid (Desktop) ===== */
.cal-grid { display: none; }
@media (min-width: 768px) {
  .cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 1px; background: #E0E0E0; padding: 1px; }
  .cal-list { display: none !important; }
}
.cal-header { background: #F5F5F5; padding: 6px 4px; text-align: center; font-weight: 700; font-size: 12px; color: #666; }
.cal-header.sun { color: #E74C3C; }
.cal-header.sat { color: #2980B9; }
.cal-cell { background: #fff; min-height: 80px; padding: 4px; font-size: 11px; }
.cal-cell.empty { background: #F8F8F8; }
.cal-cell.no-shift { background: #FFF5F5; }
.cal-cell .day-num { font-weight: 700; font-size: 13px; margin-bottom: 2px; }
.cal-cell .day-num.sun { color: #E74C3C; }
.cal-cell .day-num.sat { color: #2980B9; }
.cal-cell .shift-entry { padding: 1px 3px; margin: 1px 0; border-radius: 3px; font-size: 10px; line-height: 1.3; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }

/* ===== Calendar List (Mobile) ===== */
.cal-list { padding: 8px; }
@media (min-width: 768px) { .cal-list { display: none; } }
.cal-list-day { padding: 10px 12px; border-bottom: 1px solid #F0F0F0; }
.cal-list-day:last-child { border-bottom: none; }
.cal-list-date { font-weight: 800; font-size: 14px; margin-bottom: 4px; }
.cal-list-date.sun { color: #E74C3C; }
.cal-list-date.sat { color: #2980B9; }
.cal-list-shift { display: flex; align-items: center; gap: 8px; padding: 3px 0; font-size: 13px; }
.cal-list-empty { color: #CCC; font-size: 13px; }

/* ===== Table ===== */
.data-table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 12px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.08); }
.data-table th { background: #F5F5F5; padding: 10px 12px; text-align: left; font-size: 13px; font-weight: 700; color: #555; border-bottom: 2px solid #E0E0E0; cursor: pointer; user-select: none; white-space: nowrap; }
.data-table th .sort-arrow { font-size: 10px; margin-left: 4px; opacity: .4; }
.data-table th.sorted-asc .sort-arrow, .data-table th.sorted-desc .sort-arrow { opacity: 1; color: #1DB446; }
.data-table td { padding: 10px 12px; border-bottom: 1px solid #F0F0F0; font-size: 14px; }
.data-table tr:hover td { background: #F0FFF4; }
.data-table .num { text-align: right; font-variant-numeric: tabular-nums; }

/* ===== Staff Cards (Mobile Alt) ===== */
.staff-card-list { }
.staff-card { background: #fff; border-radius: 12px; padding: 14px 16px; margin-bottom: 8px; box-shadow: 0 1px 3px rgba(0,0,0,.06); }
.staff-card-name { font-weight: 700; font-size: 16px; }
.staff-card-no { color: #888; font-size: 12px; }
.staff-card-stats { display: flex; gap: 16px; margin-top: 6px; font-size: 13px; color: #555; }
.staff-card-stats b { color: #1DB446; }
.staff-card-facs { margin-top: 6px; font-size: 12px; color: #888; }
.staff-card-facs a { color: #1DB446; text-decoration: none; }

/* ===== Issue List ===== */
.issue-section { background: #fff; border-radius: 12px; padding: 16px; margin-bottom: 12px; box-shadow: 0 1px 4px rgba(0,0,0,.08); }
.issue-section h3 { font-size: 15px; margin-bottom: 10px; color: #555; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.issue-section h3 .count { background: #E74C3C; color: #fff; font-size: 11px; padding: 2px 8px; border-radius: 10px; }
.issue-list { list-style: none; }
.issue-list li { padding: 10px 0; border-bottom: 1px solid #F0F0F0; font-size: 14px; }
.issue-list li:last-child { border-bottom: none; }
.issue-list li .tag { display: inline-block; font-size: 10px; padding: 2px 6px; border-radius: 4px; margin-right: 6px; }
.tag-warn { background: #FEF3E2; color: #E67E22; }
.tag-error { background: #FDEDEC; color: #E74C3C; }

/* ===== Coverage Calendar ===== */
.cov-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 2px; }
.cov-header { text-align: center; font-weight: 700; font-size: 11px; padding: 6px 2px; color: #666; }
.cov-header.sun { color: #E74C3C; }
.cov-header.sat { color: #2980B9; }
.cov-cell { background: #fff; border-radius: 6px; padding: 6px 2px; text-align: center; min-height: 60px; border: 1px solid #EEE; }
.cov-cell.empty { background: transparent; border: none; }
.cov-cell .cov-day { font-weight: 700; font-size: 13px; }
.cov-cell .cov-day.sun { color: #E74C3C; }
.cov-cell .cov-day.sat { color: #2980B9; }
.cov-cell .cov-num { font-size: 18px; font-weight: 800; margin: 1px 0; }
.cov-cell .cov-detail { font-size: 9px; color: #888; }
.cov-level-0 { background: #FFF5F5; border-color: #FFCCCC; }
.cov-level-1 { background: #FEF9E7; border-color: #F9E79F; }
.cov-level-2 { background: #EAFAF1; border-color: #A9DFBF; }
.cov-level-3 { background: #D5F5E3; border-color: #82E0AA; }

.hidden { display: none !important; }

/* ===== Desktop ===== */
@media (min-width: 768px) {
  body { font-size: 14px; }
  .header { padding: 16px 24px; }
  .header h1 { font-size: 20px; }
  .staff-selector { padding: 10px 24px; }
  .tab-content { padding: 20px; max-width: 1200px; }
  .kpi-grid { grid-template-columns: repeat(3, 1fr); }
  .kpi-card .value { font-size: 36px; }
  #my-shift-view { padding: 20px; max-width: 1200px; }
}
</style>
</head>
<body>

<div class="header">
  <h1>${esc(`${year}年${month}月`)} シフトダッシュボード</h1>
  <div class="meta">生成: ${esc(generatedAt)}</div>
</div>

<div class="staff-selector">
  <select id="staff-select" onchange="onStaffChange()">
    <option value="">管理者モード（全施設表示）</option>
    ${staffOptions}
  </select>
  <span class="mode-label mode-admin" id="mode-badge">管理者</span>
</div>

<div class="tab-bar" id="admin-tabs">
  <button class="tab-btn active" data-tab="overview">概要</button>
  <button class="tab-btn" data-tab="facility">施設別</button>
  <button class="tab-btn" data-tab="staff">職員</button>
  <button class="tab-btn" data-tab="issues">問題</button>
  <button class="tab-btn" data-tab="coverage">カバレッジ</button>
  <button class="tab-btn" data-tab="preference">希望一覧</button>
</div>

<!-- ===== 職員モード: マイシフト ===== -->
<div id="my-shift-view" class="hidden"></div>

<!-- ===== Tab 1: 概要 ===== -->
<div class="tab-content active" id="tab-overview">
  <div class="kpi-grid">
    <div class="kpi-card"><div class="value">${totalShifts.toLocaleString()}</div><div class="label">総シフト数</div></div>
    <div class="kpi-card"><div class="value">${uniqueStaff}</div><div class="label">配置職員数</div></div>
    <div class="kpi-card"><div class="value">${uniqueFacilities}/${totalFacilitiesInInfo}</div><div class="label">稼働施設</div></div>
    <div class="kpi-card ${parseFloat(matchRate) < 90 ? 'warn' : ''}"><div class="value">${matchRate}%</div><div class="label">マッチ率</div></div>
    <div class="kpi-card"><div class="value">${avgPerDay}</div><div class="label">1日平均</div></div>
    <div class="kpi-card ${unmatchedNames.size > 0 ? 'danger' : ''}"><div class="value">${unmatchedNames.size}</div><div class="label">未マッチ</div></div>
  </div>

  <div class="issue-section">
    <h3>時間帯別シフト分布</h3>
    <table class="data-table" style="margin-top:8px;">
      <thead><tr><th>時間帯</th><th class="num">件数</th><th class="num">割合</th></tr></thead>
      <tbody>
${(() => {
  const slotCounts = {};
  for (const r of records) { slotCounts[r.timeSlot] = (slotCounts[r.timeSlot] || 0) + 1; }
  return Object.entries(slotCounts).sort((a,b) => b[1]-a[1]).map(([slot, cnt]) =>
    `        <tr><td><span style="color:${getTimeSlotColor(slot)};font-weight:700;">${esc(slot)}</span></td><td class="num">${cnt}</td><td class="num">${(cnt/totalShifts*100).toFixed(1)}%</td></tr>`
  ).join('\n');
})()}
      </tbody>
    </table>
  </div>
</div>

<!-- ===== Tab 2: 施設別 ===== -->
<div class="tab-content" id="tab-facility">
  <div class="filter-bar">
    <input type="text" id="facility-filter" placeholder="施設名で絞り込み..." oninput="filterFacilities()">
  </div>
${facilityIds.map(fid => {
  const fname = facilityNames[fid] || fid;
  const info = FACILITY_INFO[fid];
  const room = info && info.room ? info.room : '';
  const cal = facilityCalendar[fid] || {};
  const totalForFacility = Object.values(cal).reduce((s, arr) => s + arr.length, 0);
  const firstDow = new Date(year, month - 1, 1).getDay();

  // Desktop grid
  let calCells = '';
  const dowNames = ['日', '月', '火', '水', '木', '金', '土'];
  for (let i = 0; i < 7; i++) {
    const cls = i === 0 ? ' sun' : i === 6 ? ' sat' : '';
    calCells += `<div class="cal-header${cls}">${dowNames[i]}</div>`;
  }
  for (let i = 0; i < firstDow; i++) calCells += '<div class="cal-cell empty"></div>';
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(year, month - 1, d).getDay();
    const dayShifts = cal[d] || [];
    const cellClass = dayShifts.length === 0 ? 'cal-cell no-shift' : 'cal-cell';
    const dayClass = dow === 0 ? ' sun' : dow === 6 ? ' sat' : '';
    let entries = '';
    for (const s of dayShifts) {
      entries += `<div class="shift-entry" style="background:${getTimeSlotBg(s.timeSlot)};color:${getTimeSlotColor(s.timeSlot)}" title="${esc(s.timeSlot + ' ' + s.name)}">${esc(s.timeSlot.replace('時～', '-').replace('時', ''))} ${esc(s.name)}</div>`;
    }
    calCells += `<div class="${cellClass}"><div class="day-num${dayClass}">${d}</div>${entries}</div>`;
  }
  const lastDow = new Date(year, month - 1, daysInMonth).getDay();
  for (let i = lastDow + 1; i < 7; i++) calCells += '<div class="cal-cell empty"></div>';

  // Mobile list
  let calList = '';
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = getDayOfWeek(year, month, d);
    const dowIdx = new Date(year, month - 1, d).getDay();
    const dayClass = dowIdx === 0 ? ' sun' : dowIdx === 6 ? ' sat' : '';
    const dayShifts = cal[d] || [];
    let shiftsHtml = '';
    if (dayShifts.length === 0) {
      shiftsHtml = '<div class="cal-list-empty">-</div>';
    } else {
      for (const s of dayShifts) {
        shiftsHtml += `<div class="cal-list-shift"><span style="color:${getTimeSlotColor(s.timeSlot)};font-weight:700;min-width:80px;font-size:12px;">${esc(s.timeSlot)}</span><span>${esc(s.name)}</span></div>`;
      }
    }
    calList += `<div class="cal-list-day"><div class="cal-list-date${dayClass}">${month}/${d}(${dow})</div>${shiftsHtml}</div>`;
  }

  return `  <details class="facility-section" data-facility-name="${esc(fname)}" data-facility-id="${esc(fid)}">
    <summary>
      ${esc(fname)}
      <span class="badge">${totalForFacility}件</span>
      ${room ? `<span class="info">${esc(room)}</span>` : ''}
    </summary>
    <div class="facility-toolbar"><a class="detail-link" href="facilities/${esc(fid)}.html">施設詳細 &rarr;</a></div>
    <div class="cal-grid">${calCells}</div>
    <div class="cal-list">${calList}</div>
  </details>`;
}).join('\n')}
</div>

<!-- ===== Tab 3: 職員 ===== -->
<div class="tab-content" id="tab-staff">
  <div class="filter-bar">
    <input type="text" id="staff-filter" placeholder="名前で検索..." oninput="filterStaff()">
  </div>
  <div class="staff-card-list">
${staffList.map(s =>
  `    <div class="staff-card" data-emp-no="${esc(s.employeeNo)}" data-name="${esc(s.name.toLowerCase())}">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <div><span class="staff-card-name">${esc(s.name)}</span> <span class="staff-card-no">No.${esc(s.employeeNo)}</span></div>
      </div>
      <div class="staff-card-stats">
        <span>勤務 <b>${s.shiftCount}</b>回</span>
        <span><b>${s.dayCount}</b>日</span>
        <span><b>${s.facilityCount}</b>施設</span>
      </div>
      <div class="staff-card-facs">${s.facilities.map(f => {
        const fidEntry = Object.entries(facilityNames).find(([, name]) => name === f);
        return fidEntry ? `<a href="facilities/${esc(fidEntry[0])}.html">${esc(f)}</a>` : esc(f);
      }).join(', ')}</div>
    </div>`
).join('\n')}
  </div>
</div>

<!-- ===== Tab 4: 問題 ===== -->
<div class="tab-content" id="tab-issues">
  <div class="issue-section">
    <h3>未マッチ名 <span class="count">${unmatchedNames.size}件</span></h3>
    ${unmatchedNames.size > 0 ? `<ul class="issue-list">
${[...unmatchedNames].sort().map(n => {
  const cnt = records.filter(r => r.originalName === n && !r.employeeNo).length;
  return `      <li><span class="tag tag-error">未マッチ</span>${esc(n)} <span style="color:#999;font-size:12px;">(${cnt}件)</span></li>`;
}).join('\n')}
    </ul>` : '<p style="color:#1DB446;font-weight:700;padding:12px;">全員マッチ済み</p>'}
  </div>
  <div class="issue-section">
    <h3>未マッピング施設 <span class="count">${unmappedFacilities.size}件</span></h3>
    ${unmappedFacilities.size > 0 ? `<ul class="issue-list">
${[...unmappedFacilities].sort().map(f => {
  const cnt = records.filter(r => !r.facilityId && (r.excelFacility === f || r.facility === f)).length;
  return `      <li><span class="tag tag-warn">未マッピング</span>${esc(f)} <span style="color:#999;font-size:12px;">(${cnt}件)</span></li>`;
}).join('\n')}
    </ul>` : '<p style="color:#1DB446;font-weight:700;padding:12px;">全施設マッピング済み</p>'}
  </div>
  <div class="issue-section">
    <h3>未稼働施設</h3>
    ${(() => {
      const infoIds = new Set(Object.keys(FACILITY_INFO));
      const activeIds = new Set(facilityIds);
      const inactive = [...infoIds].filter(id => !activeIds.has(id)).sort();
      if (inactive.length === 0) return '<p style="color:#1DB446;font-weight:700;padding:12px;">全施設にシフト配置あり</p>';
      return `<ul class="issue-list">${inactive.map(id => `<li><span class="tag tag-warn">未稼働</span>${esc(id)} - ${esc(FACILITY_INFO[id]?.name || '不明')}</li>`).join('')}</ul>`;
    })()}
  </div>
</div>

<!-- ===== Tab 5: カバレッジ ===== -->
<div class="tab-content" id="tab-coverage">
  <div class="kpi-grid" style="margin-bottom:16px;">
    <div class="kpi-card"><div class="value">${Math.max(...Object.values(dailyCoverage).map(d => d.shifts))}</div><div class="label">最大/日</div></div>
    <div class="kpi-card"><div class="value">${Math.min(...Object.values(dailyCoverage).map(d => d.shifts))}</div><div class="label">最小/日</div></div>
  </div>
  <div class="cov-grid">
    ${['日','月','火','水','木','金','土'].map((d, i) => {
      const cls = i === 0 ? ' sun' : i === 6 ? ' sat' : '';
      return `<div class="cov-header${cls}">${d}</div>`;
    }).join('')}
    ${(() => {
      const firstDow = new Date(year, month - 1, 1).getDay();
      let cells = '';
      for (let i = 0; i < firstDow; i++) cells += '<div class="cov-cell empty"></div>';
      for (let d = 1; d <= daysInMonth; d++) {
        const dow = new Date(year, month - 1, d).getDay();
        const cov = dailyCoverage[d];
        const shifts = cov.shifts;
        const facCount = cov.facilities.size;
        const staffCount = cov.staff.size;
        const avg = totalShifts / daysInMonth;
        let level = shifts === 0 ? 0 : shifts < avg * 0.5 ? 1 : shifts < avg * 1.2 ? 2 : 3;
        const dayClass = dow === 0 ? ' sun' : dow === 6 ? ' sat' : '';
        cells += `<div class="cov-cell cov-level-${level}"><div class="cov-day${dayClass}">${d}</div><div class="cov-num">${shifts}</div><div class="cov-detail">${facCount}施設 ${staffCount}名</div></div>`;
      }
      const lastDow = new Date(year, month - 1, daysInMonth).getDay();
      for (let i = lastDow + 1; i < 7; i++) cells += '<div class="cov-cell empty"></div>';
      return cells;
    })()}
  </div>
</div>

<!-- ===== Tab 6: 希望一覧 ===== -->
<div class="tab-content" id="tab-preference">
${prefData.length > 0 ? `
  <div class="kpi-grid" style="margin-bottom:16px;">
    <div class="kpi-card"><div class="value">${prefData.length}</div><div class="label">希望総数</div></div>
    <div class="kpi-card"><div class="value">${submittedEmpNos.size}</div><div class="label">提出済み</div></div>
    <div class="kpi-card ${notSubmitted.length > 0 ? 'warn' : ''}"><div class="value">${notSubmitted.length}</div><div class="label">未提出</div></div>
    <div class="kpi-card"><div class="value">${prefData.filter(p => p.type === 'NG').length}</div><div class="label">NG件数</div></div>
  </div>

  <div class="issue-section">
    <h3>提出状況</h3>
    <div class="filter-bar"><input type="text" id="pref-filter" placeholder="名前で検索..." oninput="filterPref()"></div>
    <div class="staff-card-list" id="pref-list">
${Object.entries(prefByEmployee).sort(([a],[b]) => a.localeCompare(b)).map(([empNo, data]) => {
  const wantCnt = data.prefs.filter(p => p.type === '希望').length;
  const ngCnt = data.prefs.filter(p => p.type === 'NG').length;
  const eitherCnt = data.prefs.filter(p => p.type === 'どちらでも').length;
  const ngDates = data.prefs.filter(p => p.type === 'NG').map(p => {
    const parts = p.date.split('/');
    return parseInt(parts[2]) + '日';
  }).join(', ');
  return `      <div class="staff-card" data-name="${esc(data.name.toLowerCase())}">
        <div style="display:flex;justify-content:space-between;align-items:center;">
          <div><span class="staff-card-name">${esc(data.name)}</span> <span class="staff-card-no">No.${esc(empNo)}</span></div>
          <span style="font-size:11px;color:#999;">${esc(data.submittedAt || '')}</span>
        </div>
        <div class="staff-card-stats">
          <span style="color:#27AE60;">希望 <b>${wantCnt}</b></span>
          <span style="color:#E74C3C;">NG <b>${ngCnt}</b></span>
          <span style="color:#F39C12;">他 <b>${eitherCnt}</b></span>
        </div>
        ${ngCnt > 0 ? `<div style="font-size:12px;color:#E74C3C;margin-top:4px;">NG日: ${esc(ngDates)}</div>` : ''}
      </div>`;
}).join('\n')}
    </div>
  </div>

  ${notSubmitted.length > 0 ? `
  <div class="issue-section" style="margin-top:12px;">
    <h3>未提出者 <span class="count">${notSubmitted.length}名</span></h3>
    <ul class="issue-list">
${notSubmitted.map(e => `      <li><span class="tag tag-warn">未提出</span>${esc(e.name)} (No.${esc(e.employeeNo)})</li>`).join('\n')}
    </ul>
  </div>` : ''}

  <div class="issue-section" style="margin-top:12px;">
    <h3>日別 希望/NG分布</h3>
    <div class="cov-grid">
      ${['日','月','火','水','木','金','土'].map((d, i) => {
        const cls = i === 0 ? ' sun' : i === 6 ? ' sat' : '';
        return `<div class="cov-header${cls}">${d}</div>`;
      }).join('')}
      ${(() => {
        const firstDow = new Date(year, month - 1, 1).getDay();
        let cells = '';
        for (let i = 0; i < firstDow; i++) cells += '<div class="cov-cell empty"></div>';
        for (let d = 1; d <= daysInMonth; d++) {
          const dow = new Date(year, month - 1, d).getDay();
          const dateStr = year + '/' + String(month).padStart(2,'0') + '/' + String(d).padStart(2,'0');
          const dp = prefByDate[dateStr] || { want: 0, ng: 0, either: 0 };
          const dayClass = dow === 0 ? ' sun' : dow === 6 ? ' sat' : '';
          const bgColor = dp.ng > 3 ? '#FADBD8' : dp.want > 3 ? '#D5F5E3' : '';
          cells += `<div class="cov-cell" ${bgColor ? `style="background:${bgColor};"` : ''}><div class="cov-day${dayClass}">${d}</div><div style="font-size:10px;"><span style="color:#27AE60;">${dp.want}</span>/<span style="color:#E74C3C;">${dp.ng}</span></div></div>`;
        }
        const lastDow = new Date(year, month - 1, daysInMonth).getDay();
        for (let i = lastDow + 1; i < 7; i++) cells += '<div class="cov-cell empty"></div>';
        return cells;
      })()}
    </div>
  </div>
` : `
  <div class="issue-section">
    <h3>シフト希望データ</h3>
    <div style="padding:24px;text-align:center;color:#999;">
      <p style="font-size:15px;">希望データがまだ読み込まれていません。</p>
      <p style="font-size:13px;margin-top:8px;">
        Google Sheetsの「シフト希望」シートからTSVエクスポートして<br>
        <code style="background:#F0F2F5;padding:2px 8px;border-radius:4px;">data/shift-preferences.tsv</code> に配置してください。
      </p>
      <p style="font-size:12px;margin-top:12px;color:#BBB;">
        列: 対象年月 / 社員No / 氏名 / 日付 / 種別 / 時間帯 / 理由 / 提出日時
      </p>
    </div>
  </div>
`}
</div>

<script>
var STAFF_SHIFTS = ${staffShiftsJson};
var STAFF_FACILITIES = ${staffFacilityJson};
var DOW_NAMES = ['日','月','火','水','木','金','土'];
var TS_COLORS = {'翌':'#C0392B','6':'#E67E22','17':'#8E44AD','22':'#2C3E80','21':'#2C3E80'};
var TS_BGS = {'翌':'#FDEDEC','6':'#FEF5E7','17':'#F4ECF7','22':'#EBF5FB','21':'#EBF5FB'};
function tsColor(t){for(var k in TS_COLORS)if(t.indexOf(k)>=0)return TS_COLORS[k];return '#555';}
function tsBg(t){for(var k in TS_BGS)if(t.indexOf(k)>=0)return TS_BGS[k];return '#F8F9FA';}

// Tab switching
document.querySelectorAll('.tab-btn').forEach(function(btn){
  btn.addEventListener('click',function(){
    document.querySelectorAll('.tab-btn').forEach(function(b){b.classList.remove('active');});
    document.querySelectorAll('.tab-content').forEach(function(c){c.classList.remove('active');});
    btn.classList.add('active');
    document.getElementById('tab-'+btn.dataset.tab).classList.add('active');
  });
});

// Staff mode switch
function onStaffChange(){
  var empNo=document.getElementById('staff-select').value;
  // Save to sessionStorage for back-button restore
  if(empNo){sessionStorage.setItem('selectedStaff',empNo);}else{sessionStorage.removeItem('selectedStaff');}
  applyStaffMode(empNo);
}
function applyStaffMode(empNo){
  var badge=document.getElementById('mode-badge');
  var isAdmin=!empNo;
  var adminTabs=document.getElementById('admin-tabs');
  var myView=document.getElementById('my-shift-view');

  if(isAdmin){
    badge.textContent='管理者';badge.className='mode-label mode-admin';
    adminTabs.classList.remove('hidden');
    myView.classList.add('hidden');
    document.querySelectorAll('.tab-content').forEach(function(c){c.classList.remove('active');});
    // Reactivate first tab
    var activeBtn=document.querySelector('.tab-btn.active');
    if(!activeBtn){activeBtn=document.querySelector('.tab-btn');activeBtn.classList.add('active');}
    document.getElementById('tab-'+activeBtn.dataset.tab).classList.add('active');
  } else {
    badge.textContent='マイシフト';badge.className='mode-label mode-staff';
    adminTabs.classList.add('hidden');
    document.querySelectorAll('.tab-content').forEach(function(c){c.classList.remove('active');});
    myView.classList.remove('hidden');
    renderMyShifts(empNo);
  }
}
// Restore staff selection on page load (back-button support)
(function(){
  var saved=sessionStorage.getItem('selectedStaff');
  if(saved){
    var sel=document.getElementById('staff-select');
    if(sel){sel.value=saved;applyStaffMode(saved);}
  }
})();

function renderMyShifts(empNo){
  var data=STAFF_SHIFTS[empNo];
  var el=document.getElementById('my-shift-view');
  if(!data||!data.shifts.length){
    el.innerHTML='<div class="my-header"><div class="my-name">シフトデータがありません</div></div>';
    return;
  }
  var facCount=data.facilityIds.length;
  var shiftCount=data.shifts.length;
  var days={};
  data.shifts.forEach(function(s){if(!days[s.d])days[s.d]=[];days[s.d].push(s);});
  var dayCount=Object.keys(days).length;

  var html='<div class="my-header">';
  html+='<div class="my-name">'+esc(data.name)+' さん</div>';
  html+='<div class="my-stats">';
  html+='<div class="my-stat"><div class="my-stat-val">'+shiftCount+'</div><div class="my-stat-label">勤務回数</div></div>';
  html+='<div class="my-stat"><div class="my-stat-val">'+dayCount+'</div><div class="my-stat-label">勤務日数</div></div>';
  html+='<div class="my-stat"><div class="my-stat-val">'+facCount+'</div><div class="my-stat-label">施設数</div></div>';
  html+='</div></div>';

  var sortedDates=Object.keys(days).sort();
  sortedDates.forEach(function(dateStr){
    var parts=dateStr.split('/').map(Number);
    var dow=DOW_NAMES[new Date(parts[0],parts[1]-1,parts[2]).getDay()];
    var dowIdx=new Date(parts[0],parts[1]-1,parts[2]).getDay();
    var dayClass=dowIdx===0?' sun':dowIdx===6?' sat':'';
    html+='<div class="my-day-card">';
    html+='<div class="my-day-header'+dayClass+'"><span>'+parts[1]+'/'+parts[2]+' ('+dow+')</span><span style="font-size:12px;font-weight:400;">'+days[dateStr].length+'件</span></div>';
    days[dateStr].forEach(function(s){
      html+='<div class="my-shift-item">';
      html+='<span class="my-time-badge" style="background:'+tsBg(s.t)+';color:'+tsColor(s.t)+';">'+esc(s.t)+'</span>';
      html+='<span class="my-facility-name">'+esc(s.f)+'</span>';
      if(s.fid){html+='<a class="my-facility-link" href="facilities/'+esc(s.fid)+'.html">詳細</a>';}
      html+='</div>';
    });
    html+='</div>';
  });
  el.innerHTML=html;
}

function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

// Facility filter
function filterFacilities(){
  var q=document.getElementById('facility-filter').value.toLowerCase();
  document.querySelectorAll('.facility-section').forEach(function(el){
    var name=(el.dataset.facilityName||'').toLowerCase();
    el.style.display=!q||name.indexOf(q)>=0?'':'none';
  });
}

// Staff filter
function filterStaff(){
  var q=document.getElementById('staff-filter').value.toLowerCase();
  document.querySelectorAll('.staff-card').forEach(function(el){
    var name=(el.dataset.name||'');
    el.style.display=!q||name.indexOf(q)>=0?'':'none';
  });
}

// Preference filter
function filterPref(){
  var q=document.getElementById('pref-filter').value.toLowerCase();
  document.querySelectorAll('#pref-list .staff-card').forEach(function(el){
    var name=(el.dataset.name||'');
    el.style.display=!q||name.indexOf(q)>=0?'':'none';
  });
}

// Table sort
document.querySelectorAll('.data-table th[data-sort]').forEach(function(th){
  th.addEventListener('click',function(){
    var table=th.closest('table');var tbody=table.querySelector('tbody');
    var col=parseInt(th.dataset.col);var type=th.dataset.sort;var isAsc=th.classList.contains('sorted-asc');
    table.querySelectorAll('th').forEach(function(h){h.classList.remove('sorted-asc','sorted-desc');});
    th.classList.add(isAsc?'sorted-desc':'sorted-asc');
    var rows=Array.from(tbody.querySelectorAll('tr'));
    rows.sort(function(a,b){
      var va=a.children[col].textContent.trim();var vb=b.children[col].textContent.trim();
      if(type==='number'){va=parseFloat(va)||0;vb=parseFloat(vb)||0;}
      var cmp=type==='number'?va-vb:va.localeCompare(vb,'ja');return isAsc?-cmp:cmp;
    });
    rows.forEach(function(r){tbody.appendChild(r);});
  });
});
</script>
</body>
</html>`;

  // Write dashboard HTML
  const outPath = path.join(__dirname, '..', 'data', `shift-dashboard-${yearMonth}.html`);
  fs.writeFileSync(outPath, html, 'utf-8');
  const sizeKB = (Buffer.byteLength(html, 'utf-8') / 1024).toFixed(1);
  console.log(`\nダッシュボード生成完了!`);
  console.log(`  出力: ${outPath}`);
  console.log(`  サイズ: ${sizeKB} KB`);
  console.log(`  施設詳細: ${facilitiesDir}/ (${facilityPagesGenerated}ファイル)`);
  console.log(`  施設数: ${uniqueFacilities}, 職員数: ${uniqueStaff}, シフト数: ${totalShifts}`);
  console.log(`  未マッチ名: ${unmatchedNames.size}件, 未マッピング施設: ${unmappedFacilities.size}件`);
}

main();
