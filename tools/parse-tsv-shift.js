#!/usr/bin/env node
/**
 * parse-tsv-shift.js
 * TSV形式のシフトデータ(マスターExcelからコピーしたもの)を解析
 * → スプレッドシートに書き込み
 *
 * Usage:
 *   node tools/parse-tsv-shift.js [data/feb2026-raw.tsv] [--write-ss] [--year-month 2026-02]
 */

const path = require('path');
const fs = require('fs');
const https = require('https');

const envPath = path.join(__dirname, '..', '.env');
const envContent = fs.readFileSync(envPath, 'utf-8');
const env = {};
envContent.split('\n').forEach(line => {
  const [key, ...vals] = line.split('=');
  if (key && vals.length) env[key.trim()] = vals.join('=').trim();
});
const WEBAPP_URL = env.WEBAPP_URL;
const ADMIN_KEY = env.ADMIN_API_KEY || '';

// === 施設マッピング (TSV施設名 → 正式名/ID) ===
const FACILITY_MAP = {
  '大泉町': { id: 'GH1', name: '大泉町' },
  'グリーンビレッジB': { id: 'GH3', name: 'グリーンビレッジＢ' },
  '長久保': { id: 'GH4', name: '長久保' },
  '中町': { id: 'GH6', name: '中町' },
  'グリーンビレッジE': { id: 'GH7', name: 'グリーンビレッジＥ' },
  '南大泉': { id: 'GH8', name: '南大泉' },
  '春日町同一①': { id: 'GH9', name: '春日町 (B103)' },
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
  '天満①(102)同一⑥': { id: 'GH26', name: '天満①' },
  '天満②(302)同一⑥': { id: 'GH27', name: '天満②' },
  '南大泉３丁目': { id: 'GH28', name: '南大泉３丁目' },
  'ビレッジE102': { id: 'GH29', name: 'グリーンビレッジE102' },
  '赤塚３丁目（アーバン）': { id: 'GH30', name: '赤塚3 (アーバン)' },
  '赤塚５丁目（シティハイム第７）': { id: 'GH31', name: '赤塚5 (シティハイム第7)' },
  '寿': { id: 'GH32', name: '寿' },
  '下井草Ｂ（プラザ阿佐ヶ谷）': { id: 'GH33', name: '下井草' },
  '方南（ﾀﾏｷﾊｲﾂ）': { id: 'GH34', name: '方南' },
  '江古田（ユエヴィ江古田）': { id: 'GH35', name: '江古田 (part1 201)' },
  '砧②207同一④': { id: 'GH36', name: '砧2 (207)' },
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

// 施設ベース同姓判別ヒント (officialFacility名 → { cleanName: employeeNo })
const FACILITY_SURNAME_HINTS = {
  'グリーンビレッジＢ': { '青木': '011' },
  '芦花公園2': { '青木': '149' },
  '西長久保': { '菊池': '058' },            // 菊池明子 (≠菊池昇#078)
  'グリーンビレッジE102': { '菊池': '078' }, // 菊池昇
  'グリーンビレッジＥ': { '菊池': '078', '上野': '042' }, // 菊池昇, 上野蘭湖
  '大泉町': { '小川': '137' },              // 小川直子 (≠小川真弓#018)
  '春日町 (B103)': { '上野': '042' },       // 上野蘭湖
  '春日町2 (B203)': { '上野': '042' },      // 上野蘭湖
  '富士見町': { '上野': '042' },            // 上野蘭湖
  '船橋': { '劉': '071' },                  // 劉建国 (≠劉玲#141)
};

// 手動名寄せオーバーライド (cleanName後の名前 → {employeeNo, formalName})
// ※ 施設ベース判別 (FACILITY_SURNAME_HINTS) が優先される
const MANUAL_NAME_MAP = {
  '前田か': { employeeNo: '045', formalName: '前田 柑菜' },
  '岩本': { employeeNo: '150', formalName: '岩本 亮紀' },
};

function cleanName(name) {
  if (!name) return '';
  let n = String(name).trim();
  if (['空き', '職員配置不要', '配置不要', '欠員', '募集中'].some(s => n.includes(s))) return '';
  n = n.replace(/[⁻\-ー]\d+半?$/, '');
  n = n.replace(/^\d+半?[⁻\-ー]/, '');
  n = n.replace(/\d+半?[⁻\-ー]$/, '');
  n = n.replace(/\d+半?[⁻\-ー]/, '');
  n = n.replace(/\d+半$/, '');
  n = n.replace(/★/g, '');
  n = n.trim();
  if (!n || n === '空き') return '';
  return n;
}

function normalizeTimeSlot(t) {
  if (!t) return '';
  t = t.trim();
  if (t.match(/\d+時～\d+時/)) return t;
  if (t === '17時～') return '17時～22時';
  if (t === '21時～') return '21時～';
  if (t === '22時～') return '22時～';
  if (t === '6時～9時') return t;
  return t;
}

function formatDate(year, month, day) {
  return `${year}/${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}`;
}

// === TSVパーサー ===
function parseTSV(tsvPath, yearMonth) {
  const content = fs.readFileSync(tsvPath, 'utf-8');
  const lines = content.split('\n').filter(l => l.trim());

  if (lines.length < 3) {
    console.error('TSVファイルのデータが不足しています');
    return [];
  }

  const [year, month] = yearMonth.split('-').map(Number);

  // Row 0: 施設名ヘッダー
  // 形式: 2月\t大泉町\t\t\t2月\t春日町同一①\t\t\t...
  const headerCols = lines[0].split('\t');

  // 施設ブロックを検出 (4列ずつ: 月, 施設名, 空, 空)
  const facilityBlocks = [];
  for (let c = 0; c < headerCols.length; c++) {
    const val = headerCols[c].trim();
    if (val.match(/^\d+月$/)) {
      // 次の列が施設名
      const facName = (headerCols[c + 1] || '').trim();
      if (facName) {
        facilityBlocks.push({
          name: facName,
          startCol: c,  // 月日列
          col1: c + 1,  // 6時～9時
          col2: c + 2,  // 17時～ (or 17時～22時)
          col3: c + 3,  // 22時～ (or 21時～)
        });
      }
      c += 3; // 次の施設ブロックへ
    }
  }

  // Row 1: 時間帯ヘッダー
  const timeCols = lines[1].split('\t');
  // 各施設の時間帯を取得
  for (const block of facilityBlocks) {
    block.timeSlots = [];
    // col1, col2, col3 の時間帯を読み取る
    for (const col of [block.col1, block.col2, block.col3]) {
      if (col < timeCols.length) {
        const ts = timeCols[col].trim();
        if (ts && ts !== '月日') {
          block.timeSlots.push({ col, slot: normalizeTimeSlot(ts) });
        }
      }
    }
  }

  console.log(`施設数: ${facilityBlocks.length}`);
  for (const b of facilityBlocks) {
    const mapping = FACILITY_MAP[b.name];
    const id = mapping ? mapping.id : '???';
    console.log(`  ${b.name.padEnd(30)} [${id.padEnd(15)}] slots: ${b.timeSlots.map(t => t.slot).join(', ')}`);
  }

  // Row 2+: データ行
  const records = [];
  for (let i = 2; i < lines.length; i++) {
    const cols = lines[i].split('\t');
    const dayStr = cols[0].trim();

    // 日番号を抽出
    const dayMatch = dayStr.match(/^(\d{1,2})日/);
    if (!dayMatch) {
      // overflow行 (28日の翌朝分) - col0が空の場合
      // この行は6時～9時列のみにデータがある = March 1 morning
      if (!dayStr && i === lines.length - 2) {
        // Feb 28 → March 1 morning shift
        const nextMonth = month === 12 ? 1 : month + 1;
        const nextYear = month === 12 ? year + 1 : year;
        const nextYM = `${nextYear}-${String(nextMonth).padStart(2, '0')}`;
        const nextDate = formatDate(nextYear, nextMonth, 1);

        for (const block of facilityBlocks) {
          // Only 6時～9時 column (col1) for overflow
          const morningSlot = block.timeSlots.find(t => t.slot === '6時～9時');
          if (!morningSlot) continue;
          const rawName = (cols[morningSlot.col] || '').trim();
          const name = cleanName(rawName);
          if (!name) continue;

          const mapping = FACILITY_MAP[block.name];
          records.push({
            yearMonth: nextYM,
            date: nextDate,
            facility: block.name,
            officialFacility: mapping ? mapping.name : block.name,
            facilityId: mapping ? mapping.id : '',
            timeSlot: '6時～9時',
            originalName: name,
            rawName: rawName,
          });
        }
        continue;
      }
      continue;
    }

    const dayNum = parseInt(dayMatch[1], 10);
    const dateStr = formatDate(year, month, dayNum);

    for (const block of facilityBlocks) {
      const mapping = FACILITY_MAP[block.name];
      const officialName = mapping ? mapping.name : block.name;
      const facilityId = mapping ? mapping.id : '';

      for (const ts of block.timeSlots) {
        const rawName = (cols[ts.col] || '').trim();
        const name = cleanName(rawName);
        if (!name) continue;

        records.push({
          yearMonth: yearMonth,
          date: dateStr,
          facility: block.name,
          officialFacility: officialName,
          facilityId: facilityId,
          timeSlot: ts.slot,
          originalName: name,
          rawName: rawName,
        });
      }
    }
  }

  return records;
}

// === NameMatcher (parse-master-shift.jsから流用) ===
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
        const nSurname = this.normalize(surname);
        if (nSurname !== surname) {
          if (!this.surnameMap.has(nSurname)) this.surnameMap.set(nSurname, []);
          this.surnameMap.get(nSurname).push(emp);
        }
      }
      for (const a of emp.aliases) {
        this.aliasMap.set(a, emp);
        this.aliasMap.set(this.normalize(a), emp);
      }
    }
  }
  normalize(n) {
    return String(n).replace(/　/g, ' ').replace(/\s+/g, ' ').trim()
      .replace(/﨑/g, '崎').replace(/髙/g, '高').replace(/澤/g, '沢')
      .replace(/櫻/g, '桜').replace(/壽/g, '寿').replace(/惠/g, '恵')
      .replace(/张/g, '張').replace(/单/g, '単').replace(/單/g, '単')
      .replace(/华/g, '華').replace(/云/g, '雲').replace(/艳/g, '艶');
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
    if (noSp.length >= 2) {
      const partial = this.employees.filter(e => e.name.replace(/\s/g, '').startsWith(noSp) && e.name.replace(/\s/g, '').length > noSp.length);
      if (partial.length === 1) return partial[0];
    }
    return null;
  }
  matchWithFacility(name, facility, facilityEmployeeMap) {
    const t = name.trim(); if (!t) return null;
    const hints = FACILITY_SURNAME_HINTS[facility];
    if (hints && hints[t]) {
      const emp = this.employees.find(e => e.employeeNo === hints[t]);
      if (emp) return emp;
    }
    const sc = this.surnameMap.get(t) || this.surnameMap.get(this.normalize(t));
    if (!sc || sc.length <= 1) return null;
    const facilityEmps = facilityEmployeeMap.get(facility) || new Set();
    const candidates = sc.filter(e => facilityEmps.has(e.employeeNo));
    if (candidates.length === 1) return candidates[0];
    return null;
  }
}

// === HTTP helpers ===
function fetchJSON(url, n) {
  return new Promise((resolve, reject) => {
    if (n > 5) { reject(new Error('Too many redirects')); return; }
    const u = new URL(url);
    https.get({ hostname: u.hostname, path: u.pathname + u.search }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        resolve(fetchJSON(res.headers.location, n + 1)); return;
      }
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => { try { resolve(JSON.parse(d)); } catch (e) { reject(new Error(d.substring(0, 200))); } });
    }).on('error', reject);
  });
}

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
        resolve(fetchJSON(res.headers.location, 0)); return;
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

async function writeToSpreadsheet(records) {
  const BATCH_SIZE = 100;
  const SHEET_NAME = 'シフトデータ';
  const MAX_RETRIES = 3;

  if (!records.length) {
    console.log('書き込み対象のレコードがありません');
    return;
  }

  const targetMonth = records[0].yearMonth;
  console.log(`\nスプレッドシート書き込み開始 (${records.length}件, 対象月: ${targetMonth})...`);

  // Step 1: Node.js側でクリア（GAS Date型問題を回避）
  console.log('  既存データを読み込み中...');
  try {
    const readUrl = `${WEBAPP_URL}?action=read&key=${encodeURIComponent(ADMIN_KEY)}&sheet=${encodeURIComponent(SHEET_NAME)}`;
    const readResp = await fetchJSON(readUrl, 0);
    const allData = readResp.data || [];
    console.log(`  現在の行数: ${allData.length}`);

    // ヘッダー行(row 0)と対象月以外の行を保持
    const header = allData[0] || [];
    const keepRows = [];
    let removedCount = 0;
    for (let i = 1; i < allData.length; i++) {
      const cell = allData[i][0];
      let cellYM = '';
      if (cell && typeof cell === 'string') {
        // Date文字列 "2026-01-31T15:00:00.000Z" → 年月変換
        if (cell.match(/^\d{4}-\d{2}-\d{2}T/)) {
          const d = new Date(cell);
          // JSTに変換 (+9時間)
          const jst = new Date(d.getTime() + 9 * 60 * 60 * 1000);
          cellYM = jst.getFullYear() + '-' + String(jst.getMonth() + 1).padStart(2, '0');
        } else if (cell.match(/^\d{4}-\d{2}$/)) {
          cellYM = cell;
        }
      }
      if (cellYM === targetMonth) {
        removedCount++;
      } else if (cell) {
        // 空行でないものだけ保持
        keepRows.push(allData[i]);
      }
    }
    console.log(`  ${targetMonth} データ: ${removedCount}行削除, ${keepRows.length}行保持`);

    if (removedCount > 0) {
      // シート全体をクリアして書き直す
      // まずヘッダー+保持行を書き込み
      const writeData = [header, ...keepRows];
      const range = `A1:H${writeData.length}`;
      console.log('  非対象月データを書き戻し中...');
      const writeResult = await postJSON(WEBAPP_URL, {
        action: 'write',
        key: ADMIN_KEY,
        sheet: SHEET_NAME,
        range: range,
        data: writeData
      });
      if (writeResult.error) {
        console.error('  書き戻しエラー:', writeResult.error);
      } else {
        console.log(`  書き戻し完了: ${writeData.length}行`);
      }

      // 余分な行をクリア（旧データの残り部分）
      if (allData.length > writeData.length) {
        const clearStart = writeData.length + 1;
        const clearEnd = allData.length;
        const emptyRows = Array(clearEnd - clearStart + 1).fill(['','','','','','','','']);
        const clearRange = `A${clearStart}:H${clearEnd}`;
        console.log(`  余分な行をクリア中: ${clearStart}～${clearEnd}`);
        await postJSON(WEBAPP_URL, {
          action: 'write',
          key: ADMIN_KEY,
          sheet: SHEET_NAME,
          range: clearRange,
          data: emptyRows
        });
      }

      await new Promise(resolve => setTimeout(resolve, 2000));
    }
  } catch (e) {
    console.error('  クリアエラー:', e.message);
    // フォールバック: 旧GAS関数を試す
    try {
      const clearUrl = `${WEBAPP_URL}?action=runFunction&key=${encodeURIComponent(ADMIN_KEY)}&name=clearShiftDataForMonth&arg=${encodeURIComponent(targetMonth)}`;
      await fetchJSON(clearUrl, 0);
    } catch (e2) {}
    await new Promise(resolve => setTimeout(resolve, 2000));
  }

  // Step 2: Write in batches via POST append with retry
  const totalBatches = Math.ceil(records.length / BATCH_SIZE);
  let written = 0;
  let errors = 0;

  for (let batch = 0; batch < totalBatches; batch++) {
    const start = batch * BATCH_SIZE;
    const end = Math.min(start + BATCH_SIZE, records.length);
    const batchRecords = records.slice(start, end);

    const rows = batchRecords.map(r => [
      r.yearMonth,       // A: 年月
      r.date,            // B: 日付
      '',                // C: エリア
      r.officialFacility,// D: 施設名
      r.timeSlot,        // E: 時間帯
      r.originalName,    // F: 担当者名(原文)
      r.employeeNo || '',// G: 社員No
      r.formalName || '' // H: 氏名(正式)
    ]);

    let success = false;
    for (let retry = 0; retry < MAX_RETRIES; retry++) {
      try {
        const result = await postJSON(WEBAPP_URL, {
          action: 'append',
          key: ADMIN_KEY,
          sheet: SHEET_NAME,
          data: rows
        });
        if (result.error) {
          if (retry < MAX_RETRIES - 1) {
            await new Promise(resolve => setTimeout(resolve, 2000 * (retry + 1)));
            continue;
          }
          console.error(`\n  バッチ ${batch + 1}/${totalBatches} エラー: ${result.error}`);
          errors++;
        } else {
          written += rows.length;
          success = true;
        }
        break;
      } catch (e) {
        if (retry < MAX_RETRIES - 1) {
          await new Promise(resolve => setTimeout(resolve, 2000 * (retry + 1)));
          continue;
        }
        console.error(`\n  バッチ ${batch + 1}/${totalBatches} 通信エラー (${MAX_RETRIES}回失敗): ${e.message}`);
        errors++;
      }
    }

    process.stdout.write(`\r  書き込み中... ${written}/${records.length}件 (バッチ ${batch + 1}/${totalBatches})`);

    if (batch < totalBatches - 1) {
      await new Promise(resolve => setTimeout(resolve, 500));
    }
  }

  console.log(`\n  書き込み完了: ${written}件成功${errors > 0 ? ', ' + errors + '件エラー' : ''}`);
}

// === Main ===
async function main() {
  const args = process.argv.slice(2);
  const tsvPath = args.find(a => !a.startsWith('--')) || path.join(__dirname, '..', 'data', 'feb2026-raw.tsv');
  const writeSS = args.includes('--write-ss');

  // Auto-detect year-month from filename or arg
  let yearMonth = '2026-02';
  const ymArg = args.indexOf('--year-month');
  if (ymArg >= 0 && args[ymArg + 1]) yearMonth = args[ymArg + 1];

  const fullPath = path.resolve(tsvPath);
  if (!fs.existsSync(fullPath)) {
    console.error(`ファイルが見つかりません: ${fullPath}`);
    process.exit(1);
  }

  console.log(`TSV解析: ${fullPath}`);
  console.log(`対象年月: ${yearMonth}\n`);

  const records = parseTSV(fullPath, yearMonth);
  console.log(`\n合計: ${records.length}件のシフトレコード\n`);

  // 施設別集計
  const byFacility = new Map();
  for (const r of records) {
    if (!byFacility.has(r.officialFacility)) byFacility.set(r.officialFacility, { id: r.facilityId, records: [] });
    byFacility.get(r.officialFacility).records.push(r);
  }
  console.log('施設別件数:');
  for (const [name, data] of [...byFacility.entries()].sort()) {
    const uniqueNames = new Set(data.records.map(r => r.originalName));
    console.log(`  ${name.padEnd(25)} [${(data.id || '???').padEnd(15)}] ${data.records.length}件 (${uniqueNames.size}名)`);
  }

  // 未マッピング施設
  const unmappedFacilities = new Set();
  for (const r of records) { if (!r.facilityId) unmappedFacilities.add(r.facility); }
  if (unmappedFacilities.size > 0) {
    console.log(`\n未マッピング施設: ${[...unmappedFacilities].join(', ')}`);
  }

  // 名寄せ
  console.log('\n従業員マスタを読み込み中...');
  let matcher = null;
  let employees = [];
  try {
    const masterResp = await fetchJSON(
      `${WEBAPP_URL}?action=read&key=${encodeURIComponent(ADMIN_KEY)}&sheet=${encodeURIComponent('従業員マスタ')}`, 0
    );
    for (let i = 1; i < masterResp.data.length; i++) {
      const r = masterResp.data[i];
      const no = String(r[0]).trim(); const name = String(r[1]).trim();
      if (!no || !name) continue;
      employees.push({
        employeeNo: no.padStart(3, '0'),
        name,
        status: String(r[6] || '在職'),
        aliases: String(r[8] || '').split(',').map(a => a.trim()).filter(a => a)
      });
    }
    console.log(`  従業員: ${employees.length}名 (在職: ${employees.filter(e => e.status !== '退職').length}名)`);
    matcher = new NameMatcher(employees);
  } catch (e) {
    console.error(`  従業員マスタ読み込み失敗: ${e.message}`);
    console.log('  名寄せをスキップします');
  }

  // 0パス目: 施設ベース同姓判別 → 手動オーバーライド (matcher不要)
  let manualMatched = 0;
  for (const rec of records) {
    // 施設ベース判別を最優先（同姓問題の解決）
    const hints = FACILITY_SURNAME_HINTS[rec.officialFacility];
    if (hints && hints[rec.originalName]) {
      const empNo = hints[rec.originalName];
      if (employees.length > 0) {
        const emp = employees.find(e => e.employeeNo === empNo);
        if (emp) { rec.employeeNo = empNo; rec.formalName = emp.name; manualMatched++; continue; }
      }
      rec.employeeNo = empNo; rec.formalName = ''; manualMatched++; continue;
    }
    // グローバル手動マッピング
    const manual = MANUAL_NAME_MAP[rec.originalName];
    if (manual) {
      rec.employeeNo = manual.employeeNo;
      rec.formalName = manual.formalName;
      manualMatched++;
    }
  }
  if (manualMatched > 0) console.log(`手動マッピング: ${manualMatched}件`);

  if (matcher) {
    // 1パス目: 通常の名寄せ (手動マッチ済みはスキップ)
    let matched = manualMatched;
    const unmatchedRecs = [];
    for (const rec of records) {
      if (rec.employeeNo) { continue; } // 手動マッチ済み
      const emp = matcher.match(rec.originalName);
      if (emp) { rec.employeeNo = emp.employeeNo; rec.formalName = emp.name; matched++; }
      else { rec.employeeNo = ''; rec.formalName = ''; unmatchedRecs.push(rec); }
    }

    // 2パス目: 施設ベース判別
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
    const unmatched = unmatchedRecs.length - pass2Matched;

    const unmatchedMap = new Map();
    for (const rec of unmatchedRecs) {
      if (rec.employeeNo) continue;
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
  }

  // テキスト出力
  const [y, m] = yearMonth.split('-');
  const outputPath = path.join(__dirname, '..', 'data', `shift-tsv-${y}-${m}.txt`);
  let output = `========================================\n`;
  output += `  ${y}年${m}月 シフト一覧（TSV解析）\n`;
  output += `  ${byFacility.size}施設, ${records.length}件\n`;
  output += `========================================\n\n`;

  const days = ['日', '月', '火', '水', '木', '金', '土'];
  for (const [name, data] of [...byFacility.entries()].sort()) {
    output += `■ ${name} [${data.id || '???'}]\n`;
    output += '─'.repeat(60) + '\n';
    data.records.sort((a, b) => a.date.localeCompare(b.date) || a.timeSlot.localeCompare(b.timeSlot));
    let prevDate = '';
    for (const s of data.records) {
      const [yy, mm, dd] = s.date.split('/').map(Number);
      const dow = days[new Date(yy, mm - 1, dd).getDay()];
      const dateLabel = s.date !== prevDate ? `${dd}日(${dow})` : '';
      prevDate = s.date;
      const empInfo = s.employeeNo ? ` [${s.employeeNo}]` : '';
      output += `  ${dateLabel.padEnd(10)} ${s.timeSlot.padEnd(12)} ${s.originalName}${empInfo}\n`;
    }
    output += '\n';
  }

  fs.writeFileSync(outputPath, output, 'utf-8');
  console.log(`\nテキスト出力: ${outputPath}`);

  // スプレッドシート書き込み
  if (writeSS) {
    await writeToSpreadsheet(records);
  } else {
    console.log('\n--write-ss フラグでスプレッドシートに書き込みます');
  }
}

main().catch(e => { console.error('Error:', e.message); process.exit(1); });
