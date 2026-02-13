#!/usr/bin/env node
/**
 * import-master.js
 * 従業員マスタTSVをスプレッドシートに投入
 * GAS WebアプリのAPIを使ってバッチ書き込み
 *
 * 実行: node tools/import-master.js
 */

const fs = require('fs');
const path = require('path');
const https = require('https');
const http = require('http');

// .envから読み込み
const envPath = path.join(__dirname, '..', '.env');
const envContent = fs.readFileSync(envPath, 'utf-8');
const env = {};
envContent.split('\n').forEach(line => {
  const [key, ...vals] = line.split('=');
  if (key && vals.length) env[key.trim()] = vals.join('=').trim();
});

const WEBAPP_URL = env.WEBAPP_URL;
if (!WEBAPP_URL) {
  console.error('WEBAPP_URL が .env に設定されていません');
  process.exit(1);
}

const TSV_PATH = path.join(__dirname, '..', 'data', 'employee-master.tsv');

// TSV読み込み
const tsvContent = fs.readFileSync(TSV_PATH, 'utf-8');
const lines = tsvContent.split('\n').filter(l => l.trim());
const header = lines[0].split('\t');
const dataRows = lines.slice(1).map(line => line.split('\t'));

console.log(`従業員マスタ: ${dataRows.length} 行を投入します`);

// バッチサイズ（URLの長さ制限対策）
const BATCH_SIZE = 20;

async function fetchUrl(url) {
  return new Promise((resolve, reject) => {
    const get = url.startsWith('https') ? https.get : http.get;
    // GASのリダイレクトに対応
    const options = {
      headers: { 'User-Agent': 'shift-notification-importer' }
    };

    function doRequest(requestUrl, redirectCount) {
      if (redirectCount > 5) {
        reject(new Error('Too many redirects'));
        return;
      }

      const urlObj = new URL(requestUrl);
      const reqOptions = {
        hostname: urlObj.hostname,
        path: urlObj.pathname + urlObj.search,
        headers: { 'User-Agent': 'shift-notification-importer' }
      };

      https.get(reqOptions, (res) => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
          doRequest(res.headers.location, redirectCount + 1);
          return;
        }

        let data = '';
        res.on('data', chunk => data += chunk);
        res.on('end', () => {
          try {
            resolve(JSON.parse(data));
          } catch (e) {
            reject(new Error(`JSON parse error: ${data.substring(0, 200)}`));
          }
        });
      }).on('error', reject);
    }

    doRequest(url, 0);
  });
}

async function writeBatch(startRow, rows) {
  // A列:B列... の範囲を計算
  const range = `A${startRow + 2}:I${startRow + 2 + rows.length - 1}`;
  const dataJson = JSON.stringify(rows);

  const url = `${WEBAPP_URL}?action=write&sheet=${encodeURIComponent('従業員マスタ')}&range=${encodeURIComponent(range)}&data=${encodeURIComponent(dataJson)}`;

  if (url.length > 8000) {
    // URL長すぎる場合はさらに分割
    const half = Math.floor(rows.length / 2);
    await writeBatch(startRow, rows.slice(0, half));
    await writeBatch(startRow + half, rows.slice(half));
    return;
  }

  const result = await fetchUrl(url);
  if (result.error) {
    throw new Error(`Write error: ${result.error}`);
  }
  return result;
}

async function main() {
  let imported = 0;

  for (let i = 0; i < dataRows.length; i += BATCH_SIZE) {
    const batch = dataRows.slice(i, i + BATCH_SIZE);
    // 各行を9列に正規化
    const normalizedBatch = batch.map(row => {
      const padded = [...row];
      while (padded.length < 9) padded.push('');
      return padded.slice(0, 9);
    });

    process.stdout.write(`  バッチ ${Math.floor(i / BATCH_SIZE) + 1}: 行 ${i + 1}-${i + batch.length} ... `);

    try {
      const result = await writeBatch(i, normalizedBatch);
      console.log(`OK (${result.rowsWritten} 行書き込み)`);
      imported += batch.length;
    } catch (e) {
      console.log(`ERROR: ${e.message}`);
    }

    // レート制限対策
    await new Promise(r => setTimeout(r, 500));
  }

  console.log(`\n完了: ${imported}/${dataRows.length} 行を投入しました`);
}

main().catch(e => {
  console.error('エラー:', e.message);
  process.exit(1);
});
