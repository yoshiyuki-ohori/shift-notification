#!/usr/bin/env node
/**
 * Google Driveから夜間職員シフト.xlsmをダウンロード→全シート解析
 */
const path = require('path');
const fs = require('fs');
const { google } = require('googleapis');
const XLSX = require('xlsx');

const SA_PATH = path.join(
  __dirname, '..', '..', 'expense-management-system', 'credentials', 'service-account.json'
);
const serviceAccount = require(SA_PATH);
const FILE_ID = '1U_YTNiyz2fsolHhBF5GTPPxSWrek0frO';

async function main() {
  const auth = new google.auth.GoogleAuth({
    credentials: serviceAccount,
    scopes: ['https://www.googleapis.com/auth/drive.readonly'],
  });
  const drive = google.drive({ version: 'v3', auth });

  // ダウンロード
  console.log('夜間職員シフト.xlsm をダウンロード中...');
  const resp = await drive.files.get(
    { fileId: FILE_ID, alt: 'media' },
    { responseType: 'arraybuffer' }
  );
  const outputPath = path.join(__dirname, '..', 'data', 'shift-master.xlsm');
  fs.writeFileSync(outputPath, Buffer.from(resp.data));
  console.log(`  保存: ${outputPath} (${(Buffer.from(resp.data).length / 1024).toFixed(1)}KB)\n`);

  // 解析
  const workbook = XLSX.readFile(outputPath);
  const action = process.argv[2] || 'list';

  console.log(`=== シート一覧 (${workbook.SheetNames.length}シート) ===`);
  for (const name of workbook.SheetNames) {
    const ws = workbook.Sheets[name];
    if (!ws['!ref']) {
      console.log(`  ${name.padEnd(30)} (空)`);
      continue;
    }
    const range = XLSX.utils.decode_range(ws['!ref']);
    const rows = range.e.r - range.s.r + 1;
    const cols = range.e.c - range.s.c + 1;
    console.log(`  ${name.padEnd(30)} ${rows}行 x ${cols}列`);
  }

  if (action === 'detail') {
    // 各シートの先頭を表示
    for (const name of workbook.SheetNames) {
      const ws = workbook.Sheets[name];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (data.length === 0) continue;
      console.log(`\n--- ${name} (${data.length}行) ---`);
      const showRows = Math.min(data.length, 8);
      for (let i = 0; i < showRows; i++) {
        const row = data[i] || [];
        const display = row.slice(0, 12).map(c =>
          c === undefined || c === null ? '' : String(c).substring(0, 12)
        ).join('\t');
        console.log(`${String(i + 1).padStart(3)}: ${display}`);
      }
      if (data.length > showRows) console.log(`  ... (残り${data.length - showRows}行)`);
    }
  } else if (action === 'sheet') {
    // 特定シートの全データ
    const sheetName = process.argv[3];
    if (!sheetName) { console.error('Usage: node download-shift-excel.js sheet "シート名"'); process.exit(1); }
    const ws = workbook.Sheets[sheetName];
    if (!ws) { console.error('シートが見つかりません: ' + sheetName); process.exit(1); }
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    console.log(`\n--- ${sheetName} (${data.length}行) ---`);
    for (let i = 0; i < data.length; i++) {
      const row = data[i] || [];
      const display = row.slice(0, 20).map(c =>
        c === undefined || c === null ? '' : String(c).substring(0, 15)
      ).join('\t');
      console.log(`${String(i + 1).padStart(3)}: ${display}`);
    }
  } else if (action === 'facilities') {
    // 全シートから施設名を抽出
    console.log('\n=== 施設名抽出 ===');
    const facilitySet = new Set();
    for (const name of workbook.SheetNames) {
      const ws = workbook.Sheets[name];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (data.length < 2) continue;
      // シート名自体が施設名の場合も
      facilitySet.add(name);
      // 1行目に施設名がある場合
      if (data[0]) {
        for (const cell of data[0]) {
          const v = String(cell).trim();
          if (v && v.length > 1 && v.length < 30) facilitySet.add(v);
        }
      }
    }
    console.log(`  シート名/ヘッダーから抽出: ${facilitySet.size}件`);
    for (const f of [...facilitySet].sort()) {
      console.log(`  ${f}`);
    }
  }
}

main().catch(e => { console.error('Error:', e.message); process.exit(1); });
