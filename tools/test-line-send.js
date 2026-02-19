#!/usr/bin/env node
/**
 * test-line-send.js
 * シフトをパース → 時間帯統合 → LINE Flex Message送信
 *
 * 改善点:
 *   1. 17-22時 + 22時～ + 翌6-9時 → 「17時～翌9時」に統合
 *   2. 施設名フル表示（途切れない）
 *   3. 施設タップで地図表示、補足情報（部屋・利用者数）表示
 */
const path = require('path');
const fs = require('fs');
const https = require('https');
const XLSX = require('xlsx');

// ===== 設定 =====
const LINE_TOKEN = process.env.LINE_TOKEN || process.argv[3] || '';
const TEST_USER_ID = process.argv[2] || 'Uba564e6a9fdc8616b146f03f1c4f22f9';
const TARGET_SHEET = process.argv[4] || '2026.3';

// 施設マッピング (Excel施設名 → 施設ID/正式名)
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

// 施設ID → 正式名の逆引き
const ID_TO_FACILITY = {};
for (const [, v] of Object.entries(FACILITY_MAP)) {
  ID_TO_FACILITY[v.id] = v.name;
}

// 施設情報 (Firestoreから出力済み)
let FACILITY_INFO = {};
try {
  FACILITY_INFO = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'data', 'facility-info.json'), 'utf-8'));
} catch (e) { /* なければ空 */ }

// ================================================================
// Excel parser (parse-master-shift.js と同等)
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
        records.push({ yearMonth, date: formatDate(date), facility: officialName, facilityId, timeSlot: normalizeTimeSlot(ts.slot), originalName: name });
      }
    }
  }
  return records;
}

// ================================================================
// Name matcher (ローカル TSV)
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
      status: (cols[6] || '在職').trim(),
      aliases: (cols[8] || '').split(',').map(a => a.trim()).filter(a => a)
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
// ★ 時間帯統合ロジック
// 同一人物・同一施設で連続する時間帯を統合:
//   N日 17時～22時 + N日 22時～(or 21時～) → 「17時～翌朝」
//   + N+1日 6時～9時 (同一施設)            → 「17時～翌9時」
// ================================================================
function mergeOvernightShifts(shifts) {
  // 日付+施設でソート済み前提
  shifts.sort((a, b) => a.date.localeCompare(b.date) || getTimeOrder(a.timeSlot) - getTimeOrder(b.timeSlot));

  const merged = [];
  const consumed = new Set(); // 統合で消費されたインデックス

  for (let i = 0; i < shifts.length; i++) {
    if (consumed.has(i)) continue;
    const s = shifts[i];

    // 17時～22時 を起点に統合を試みる
    if (s.timeSlot === '17時～22時') {
      // 同日・同施設の 22時～ or 21時～ を探す
      const nightIdx = shifts.findIndex((x, j) => j > i && !consumed.has(j) &&
        x.date === s.date && x.facility === s.facility &&
        (x.timeSlot === '22時～' || x.timeSlot === '21時～'));

      if (nightIdx >= 0) {
        // 翌日の 6時～9時 を探す
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

    // 22時～ or 21時～ を起点（17時台がなかった場合）
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

  // 再ソート
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
// 施設情報ヘルパー
// ================================================================
function getFacilityNote(facilityId) {
  const info = FACILITY_INFO[facilityId];
  if (!info) return null;
  const parts = [];
  if (info.room) parts.push(info.room);
  if (info.userCount > 0) parts.push(`利用者${info.userCount}名`);
  return parts.length > 0 ? parts.join(' / ') : null;
}

// ================================================================
// ★ Flex Message 生成（改善版）
// ================================================================
function getDayOfWeek(y, m, d) {
  return ['日', '月', '火', '水', '木', '金', '土'][new Date(y, m - 1, d).getDay()];
}

function getTimeSlotColor(ts) {
  if (ts.includes('翌')) return '#C0392B';     // 夜勤通し: 赤系
  if (ts.match(/^6時/)) return '#E67E22';       // 早朝: オレンジ
  if (ts.match(/^17時/)) return '#8E44AD';       // 夕方: パープル
  if (ts.match(/^22時|^21時/)) return '#2C3E80'; // 夜間: ダークブルー
  return '#555555';
}

function buildFlexMessage(targetMonth, employeeName, shifts) {
  const [year, month] = targetMonth.split('-').map(Number);
  const displayMonth = `${year}年${month}月`;

  const bodyContents = [];
  let prevDate = '';

  for (let idx = 0; idx < shifts.length; idx++) {
    const s = shifts[idx];
    const [y, m, d] = s.date.split('/').map(Number);
    const dow = getDayOfWeek(y, m, d);
    const showDate = s.date !== prevDate;
    prevDate = s.date;

    const dateText = showDate ? `${d}日(${dow})` : '';
    const dateColor = dow === '日' ? '#FF0000' : dow === '土' ? '#0000FF' : '#333333';
    const isSunday = dow === '日', isSaturday = dow === '土';

    // 罫線（各シフトエントリの前にセパレーター、ヘッダー直後は除く）
    if (idx > 0) {
      bodyContents.push({
        type: 'separator', margin: showDate ? 'md' : 'xs',
        color: showDate ? '#CCCCCC' : '#EEEEEE'
      });
    }

    // メイン行: 日付 + 時間帯
    const mainRow = {
      type: 'box', layout: 'horizontal',
      contents: [
        {
          type: 'text', text: dateText || ' ', size: 'sm',
          color: showDate ? dateColor : '#333333',
          flex: 3, weight: showDate ? 'bold' : 'regular'
        },
        {
          type: 'text', text: s.timeSlot, size: 'sm',
          color: getTimeSlotColor(s.timeSlot),
          weight: 'bold', flex: 4
        }
      ],
      margin: 'sm',
      paddingTop: showDate ? '6px' : '2px'
    };
    if (isSunday) { mainRow.backgroundColor = '#FFF0F0'; mainRow.cornerRadius = '4px'; mainRow.paddingAll = '4px'; }
    if (isSaturday) { mainRow.backgroundColor = '#F0F0FF'; mainRow.cornerRadius = '4px'; mainRow.paddingAll = '4px'; }
    bodyContents.push(mainRow);

    // 施設名行
    const note = s.facilityId ? getFacilityNote(s.facilityId) : null;

    const facilityRow = {
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'text', text: ' ', flex: 1, size: 'xxs' },
        {
          type: 'box', layout: 'horizontal', flex: 6,
          contents: [
            {
              type: 'text', text: s.facility, size: 'xs',
              color: '#555555', flex: 4, wrap: true
            },
            {
              type: 'text', text: note || '', size: 'xxs',
              color: '#999999', flex: 3, wrap: true, gravity: 'center'
            }
          ]
        }
      ],
      margin: 'xs', paddingBottom: '4px'
    };
    bodyContents.push(facilityRow);
  }

  // 勤務回数（統合後の件数）
  const shiftCount = shifts.length;

  return {
    type: 'flex',
    altText: `${displayMonth} シフト予定 - ${employeeName}さん`,
    contents: {
      type: 'bubble', size: 'mega',
      header: {
        type: 'box', layout: 'vertical',
        contents: [
          { type: 'text', text: `${displayMonth} シフト予定`, weight: 'bold', size: 'lg', color: '#FFFFFF' },
          { type: 'text', text: `${employeeName} さん`, size: 'md', color: '#FFFFFF', margin: 'sm' }
        ],
        backgroundColor: '#1DB446', paddingAll: '15px'
      },
      body: {
        type: 'box', layout: 'vertical',
        contents: [
          // テーブルヘッダー
          {
            type: 'box', layout: 'horizontal',
            contents: [
              { type: 'text', text: '日付', size: 'xs', color: '#888888', weight: 'bold', flex: 3 },
              { type: 'text', text: '時間帯', size: 'xs', color: '#888888', weight: 'bold', flex: 4 }
            ], margin: 'md'
          },
          { type: 'separator', margin: 'sm' },
          ...bodyContents
        ],
        paddingAll: '12px', spacing: 'none'
      },
      footer: {
        type: 'box', layout: 'vertical',
        contents: [
          { type: 'separator' },
          { type: 'text', text: `合計: ${shiftCount}回の勤務`, size: 'sm', color: '#666666', margin: 'md', weight: 'bold' },
          { type: 'text', text: '変更がある場合は管理者にご連絡ください', size: 'xxs', color: '#999999', margin: 'sm', wrap: true }
        ],
        paddingAll: '15px'
      }
    }
  };
}

// ================================================================
// LINE Push API
// ================================================================
function pushMessage(userId, messages, token) {
  return new Promise((resolve, reject) => {
    const payload = JSON.stringify({ to: userId, messages });
    const options = {
      hostname: 'api.line.me', path: '/v2/bot/message/push', method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${token}`, 'Content-Length': Buffer.byteLength(payload) }
    };
    const req = https.request(options, res => {
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => {
        if (res.statusCode === 200) resolve({ success: true });
        else resolve({ success: false, status: res.statusCode, body: d });
      });
    });
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

// ================================================================
// Main
// ================================================================
async function main() {
  if (!LINE_TOKEN) {
    console.error('Usage: LINE_TOKEN=xxx node tools/test-line-send.js [userId] [token] [sheet]');
    process.exit(1);
  }

  // 1. Parse Excel
  const xlsPath = path.join(__dirname, '..', 'data', 'shift-master.xlsm');
  console.log(`Excelパース中 (${TARGET_SHEET})...`);
  const workbook = XLSX.readFile(xlsPath);
  const records = parseSheet(workbook, TARGET_SHEET);
  console.log(`  ${records.length}件のシフトレコード`);

  // 2. Name matching
  const employees = loadMasterTSV();
  console.log(`  従業員マスタ: ${employees.length}名`);
  let matched = 0;
  for (const rec of records) {
    const emp = matchName(rec.originalName, employees);
    if (emp) { rec.employeeNo = emp.employeeNo; rec.formalName = emp.name; matched++; }
  }
  console.log(`  名寄せ: ${matched}/${records.length}件マッチ`);

  // 3. Group by employee
  const byEmployee = new Map();
  for (const rec of records) {
    if (!rec.employeeNo) continue;
    if (!byEmployee.has(rec.employeeNo)) byEmployee.set(rec.employeeNo, { name: rec.formalName, shifts: [] });
    byEmployee.get(rec.employeeNo).shifts.push(rec);
  }

  // Pick first employee (or specify via env)
  const targetEmpNo = process.env.EMP_NO;
  let empNo, data;
  if (targetEmpNo && byEmployee.has(targetEmpNo)) {
    empNo = targetEmpNo;
    data = byEmployee.get(targetEmpNo);
  } else {
    [empNo, data] = [...byEmployee.entries()].sort((a, b) => a[0].localeCompare(b[0]))[0];
  }

  // 4. ★ 時間帯統合
  const before = data.shifts.length;
  const mergedShifts = mergeOvernightShifts(data.shifts);
  const after = mergedShifts.length;

  console.log(`\n対象: ${data.name} (No.${empNo})`);
  console.log(`  統合前: ${before}件 → 統合後: ${after}回の勤務`);
  for (const s of mergedShifts) {
    const mark = s.isMerged ? '★' : ' ';
    console.log(`  ${mark} ${s.date} ${s.timeSlot.padEnd(12)} ${s.facility}`);
  }

  console.log(`\n送信先: ${TEST_USER_ID}`);

  // 5. Build & send
  const [ys, ms] = TARGET_SHEET.split('.');
  const flexMsg = buildFlexMessage(`${ys}-${ms.padStart(2, '0')}`, data.name, mergedShifts);

  const result = await pushMessage(TEST_USER_ID, [flexMsg], LINE_TOKEN);
  if (result.success) {
    console.log('送信成功！LINEを確認してください。');
  } else {
    console.error(`送信失敗: HTTP ${result.status}`);
    console.error(result.body);
  }
}

main().catch(e => { console.error('Error:', e.message); process.exit(1); });
