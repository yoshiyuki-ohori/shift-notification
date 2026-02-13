#!/usr/bin/env node
/**
 * Google Drive APIでシフト表ファイルをダウンロード→解析
 * Excel(.xlsx)ファイルにも対応
 */
const path = require('path');
const fs = require('fs');
const { google } = require('googleapis');

const SA_PATH = path.join(
  __dirname, '..', '..', 'expense-management-system', 'credentials', 'service-account.json'
);
const serviceAccount = require(SA_PATH);

const FILE_ID = '1U_YTNiyz2fsolHhBF5GTPPxSWrek0frO';

async function getAuth() {
  return new google.auth.GoogleAuth({
    credentials: serviceAccount,
    scopes: [
      'https://www.googleapis.com/auth/drive.readonly',
      'https://www.googleapis.com/auth/spreadsheets.readonly',
    ],
  });
}

async function main() {
  const auth = await getAuth();
  const drive = google.drive({ version: 'v3', auth });

  // 1. ファイル情報取得
  console.log('ファイル情報を取得中...');
  const fileMeta = await drive.files.get({
    fileId: FILE_ID,
    fields: 'id,name,mimeType,size,modifiedTime',
  });
  const meta = fileMeta.data;
  console.log(`  名前: ${meta.name}`);
  console.log(`  MIME: ${meta.mimeType}`);
  console.log(`  サイズ: ${meta.size ? (meta.size / 1024).toFixed(1) + 'KB' : '不明'}`);
  console.log(`  更新日: ${meta.modifiedTime}\n`);

  const isGoogleSheet = meta.mimeType === 'application/vnd.google-apps.spreadsheet';
  const isExcel = meta.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    || meta.mimeType === 'application/vnd.ms-excel';

  if (isGoogleSheet) {
    // Google Sheets → export as xlsx
    console.log('Google Sheetsとしてエクスポート中...');
    const resp = await drive.files.export({
      fileId: FILE_ID,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }, { responseType: 'arraybuffer' });
    const outputPath = path.join(__dirname, '..', 'data', 'shift-master.xlsx');
    fs.writeFileSync(outputPath, Buffer.from(resp.data));
    console.log(`  保存: ${outputPath}`);
    parseExcel(outputPath);
  } else if (isExcel || meta.name.endsWith('.xlsx') || meta.name.endsWith('.xls')) {
    // Excel → download directly
    console.log('Excelファイルをダウンロード中...');
    const resp = await drive.files.get({
      fileId: FILE_ID,
      alt: 'media',
    }, { responseType: 'arraybuffer' });
    const outputPath = path.join(__dirname, '..', 'data', 'shift-master.xlsx');
    fs.writeFileSync(outputPath, Buffer.from(resp.data));
    console.log(`  保存: ${outputPath} (${(Buffer.from(resp.data).length / 1024).toFixed(1)}KB)`);
    parseExcel(outputPath);
  } else {
    console.log('未対応のファイル形式です: ' + meta.mimeType);

    // Google Sheetsとして開けるか試す
    try {
      console.log('\nGoogle Sheets APIで試行中...');
      const sheets = google.sheets({ version: 'v4', auth });
      const sheetMeta = await sheets.spreadsheets.get({ spreadsheetId: FILE_ID });
      console.log(`シート一覧:`);
      for (const s of sheetMeta.data.sheets) {
        console.log(`  ${s.properties.title} (gid=${s.properties.sheetId})`);
      }
    } catch (e) {
      console.log('Sheets APIでもアクセスできません: ' + e.message);
      // xlsx エクスポートを試す
      try {
        console.log('\nxlsxエクスポートを試行中...');
        const resp = await drive.files.export({
          fileId: FILE_ID,
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }, { responseType: 'arraybuffer' });
        const outputPath = path.join(__dirname, '..', 'data', 'shift-master.xlsx');
        fs.writeFileSync(outputPath, Buffer.from(resp.data));
        console.log(`  保存: ${outputPath}`);
        parseExcel(outputPath);
      } catch (e2) {
        console.log('エクスポートも失敗: ' + e2.message);
      }
    }
  }
}

function parseExcel(filePath) {
  const XLSX = require('xlsx');
  const workbook = XLSX.readFile(filePath);

  console.log(`\n=== シート一覧 (${workbook.SheetNames.length}シート) ===`);
  for (const name of workbook.SheetNames) {
    const ws = workbook.Sheets[name];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const rows = range.e.r - range.s.r + 1;
    const cols = range.e.c - range.s.c + 1;
    console.log(`  ${name.padEnd(30)} ${rows}行 x ${cols}列`);
  }

  // 各シートの先頭数行を表示
  const action = process.argv[2] || 'list';
  if (action === 'detail' || action === 'all') {
    for (const name of workbook.SheetNames) {
      const ws = workbook.Sheets[name];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      console.log(`\n--- ${name} (${data.length}行) ---`);
      const showRows = action === 'all' ? data.length : Math.min(data.length, 5);
      for (let i = 0; i < showRows; i++) {
        const row = data[i] || [];
        const display = row.slice(0, 15).map(c => c === undefined || c === null ? '' : String(c).substring(0, 15)).join('\t');
        console.log(`${String(i + 1).padStart(3)}: ${display}`);
      }
      if (data.length > showRows) console.log(`  ... (残り${data.length - showRows}行)`);
    }
  }
}

main().catch(e => { console.error('Error:', e.message); process.exit(1); });
