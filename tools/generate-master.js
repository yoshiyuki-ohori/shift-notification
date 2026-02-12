#!/usr/bin/env node
/**
 * generate-master.js
 * 社員コードCSVから従業員マスタの初期データ(TSV)を生成
 * 世田谷シフトの略称から別名リストも自動生成
 *
 * 実行: node tools/generate-master.js
 * 出力: data/employee-master.tsv (スプレッドシートに貼り付け可能)
 */

const fs = require('fs');
const path = require('path');

// ===== 設定 =====

// 世田谷シフトで使われている略称 → 対応社員番号のマッピング
// (実データ解析結果から作成)
const ALIAS_CONFIG = {
  '028': { aliases: ['マリヤ', 'ﾏﾘﾔ'], area: '世田谷' },
  '055': { aliases: ['柳幸子'], area: '世田谷' },
  '060': { aliases: ['峯田'], area: '世田谷' },
  '068': { aliases: ['石田'], area: '世田谷' },
  '072': { aliases: ['石井'], area: '練馬', facility: '南大泉' },
  '107': { aliases: ['柳岡'], area: '練馬' },
  '111': { aliases: ['飯田'], area: '練馬' },
  '119': { aliases: ['吉岡'], area: '世田谷' },
  '125': { aliases: ['山岸'], area: '世田谷' },
  '127': { aliases: ['吉瀧'], area: '世田谷' },
  '135': { aliases: [], area: '練馬' },
  '139': { aliases: ['佐藤佳'], area: '世田谷' },
  '140': { aliases: ['市川'], area: '世田谷' },
  '154': { aliases: ['吉﨑', '吉崎'], area: '世田谷' },
  '173': { aliases: ['高橋百合'], area: '世田谷' },
  '174': { aliases: ['一澤'], area: '世田谷' },
  '189': { aliases: ['旭'], area: '世田谷' },
};

// ===== CSVパーサー =====
function parseCSV(content) {
  const lines = content.split('\n');
  return lines.map(line => {
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

// ===== メイン処理 =====
const empCsvPath = path.join(__dirname, '..', '..', 'タイムシートと賃金', '社員コード - シート1.csv');
const outputDir = path.join(__dirname, '..', 'data');
const outputPath = path.join(outputDir, 'employee-master.tsv');

if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

const empCsv = fs.readFileSync(empCsvPath, 'utf-8');
const empData = parseCSV(empCsv);

// ヘッダー
const header = ['No', '氏名', 'フリガナ', 'LINE_UserId', 'エリア', '主担当施設', 'ステータス', '通知有効', '別名リスト'];
const rows = [header.join('\t')];

let validCount = 0;
let aliasCount = 0;

// ヘッダー行(2行)をスキップ
for (let i = 2; i < empData.length; i++) {
  const no = String(empData[i][0] || '').trim();
  let name = String(empData[i][1] || '').trim();
  const furigana = String(empData[i][2] || '').trim();

  if (!no || !name) continue;

  // "(MF未登録)" 等の注記を除去
  name = name.replace(/（.*?）/g, '').replace(/\(.*?\)/g, '').trim();

  const paddedNo = no.padStart(3, '0');

  // 別名・エリア情報
  const config = ALIAS_CONFIG[paddedNo] || {};
  const aliases = (config.aliases || []).join(',');
  const area = config.area || '';
  const facility = config.facility || '';

  if (aliases) aliasCount++;

  rows.push([
    paddedNo,
    name,
    furigana,
    '',           // LINE_UserId (後から登録)
    area,
    facility,
    '在職',
    'TRUE',
    aliases
  ].join('\t'));

  validCount++;
}

fs.writeFileSync(outputPath, rows.join('\n'), 'utf-8');

console.log(`従業員マスタ生成完了:`);
console.log(`  出力先: ${outputPath}`);
console.log(`  有効レコード数: ${validCount}`);
console.log(`  別名登録済み: ${aliasCount}`);
console.log(`\nスプレッドシートの「従業員マスタ」シートにTSVを貼り付けてください。`);

// ===== 別名一覧レポート出力 =====
const aliasReportPath = path.join(outputDir, 'alias-report.txt');
let report = '===== 別名リスト登録一覧 =====\n\n';
report += '以下の略称がシフト表で使われています。マスタの別名リストに事前登録済みです。\n\n';

for (const [no, config] of Object.entries(ALIAS_CONFIG).sort()) {
  if (config.aliases.length === 0) continue;
  // 社員名を取得
  let empName = '';
  for (let i = 2; i < empData.length; i++) {
    const csvNo = String(empData[i][0] || '').trim().padStart(3, '0');
    if (csvNo === no) {
      empName = String(empData[i][1] || '').trim().replace(/（.*?）/g, '').trim();
      break;
    }
  }
  report += `  ${no}: ${empName}\n`;
  report += `    別名: ${config.aliases.join(', ')}\n`;
  report += `    エリア: ${config.area}\n\n`;
}

report += '\n===== 追加が必要な可能性がある略称 =====\n\n';
report += '  柳恵子 → 該当者不明 (マスタに「柳」姓は 柳 幸子(055) のみ)\n';
report += '  ※ 世田谷シフトで「柳恵子」が出現しますが、マスタに一致する従業員がいません。\n';
report += '  ※ 確認の上、新規従業員の追加 or 別名登録が必要です。\n';

fs.writeFileSync(aliasReportPath, report, 'utf-8');
console.log(`\n別名レポート: ${aliasReportPath}`);
