#!/usr/bin/env node
/**
 * firestore-facilities.js
 * safe-rise-prod Firestoreのfacilitiesコレクションから施設マスタを取得
 * シフトCSVの施設名とのマッピングテーブルを生成
 */

const path = require('path');
const { initializeApp, cert } = require('firebase-admin/app');
const { getFirestore } = require('firebase-admin/firestore');

// ===== Firebase初期化 =====
// サービスアカウント (環境変数 or デフォルトパス)
const SERVICE_ACCOUNT_PATH = process.env.FIREBASE_SA_PATH || path.join(
  __dirname, '..', '..', 'safe-rise-api', 'credentials.json'
);

const serviceAccount = require(SERVICE_ACCOUNT_PATH);

const app = initializeApp({
  credential: cert(serviceAccount),
  projectId: 'safe-rise-prod',
});

const db = getFirestore(app);

async function main() {
  const action = process.argv[2] || 'list';

  if (action === 'list') {
    // 全施設一覧を取得
    console.log('=== safe-rise-prod Firestore: facilities コレクション ===\n');

    const snapshot = await db.collection('facilities').get();

    if (snapshot.empty) {
      console.log('施設データが見つかりません');
      process.exit(0);
    }

    const facilities = [];
    snapshot.forEach(doc => {
      facilities.push({ id: doc.id, ...doc.data() });
    });

    // ソート: name順
    facilities.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ja'));

    console.log(`${facilities.length}施設\n`);
    console.log('ID'.padEnd(30) + 'コード'.padEnd(10) + '施設名'.padEnd(25) + 'Active  スタッフ数  利用者数');
    console.log('─'.repeat(100));

    for (const f of facilities) {
      const staffCount = Array.isArray(f.staff_ids) ? f.staff_ids.length : 0;
      const userCount = Array.isArray(f.care_user_ids) ? f.care_user_ids.length : 0;
      const active = f.is_active !== false ? '○' : '×';
      console.log(
        `${(f.id || '').padEnd(30)}${(f.code || '').padEnd(10)}${(f.name || '').padEnd(25)}${active.padEnd(8)}${String(staffCount).padEnd(12)}${userCount}`
      );
    }

    // JSON出力
    console.log('\n\n=== JSON形式 ===');
    console.log(JSON.stringify(facilities.map(f => ({
      id: f.id,
      name: f.name,
      code: f.code || null,
      is_active: f.is_active !== false,
      address: f.address || null,
      phone: f.phone_number || f.phone || null,
      staff_count: Array.isArray(f.staff_ids) ? f.staff_ids.length : 0,
      care_user_count: Array.isArray(f.care_user_ids) ? f.care_user_ids.length : 0,
    })), null, 2));

  } else if (action === 'mapping') {
    // シフトCSVの施設名 → Firestore施設名のマッピング生成
    console.log('=== シフトCSV施設名 → Firestore施設 マッピング ===\n');

    const snapshot = await db.collection('facilities').get();
    const firestoreFacilities = [];
    snapshot.forEach(doc => {
      firestoreFacilities.push({ id: doc.id, ...doc.data() });
    });

    // シフトCSVから抽出される施設名一覧
    const fs = require('fs');
    const shiftDir = path.join(__dirname, '..', '..', 'タイムシートと賃金', 'シフト表');
    const setagayaPath = path.join(__dirname, '..', '..', 'タイムシートと賃金', '世田谷', '世田谷１年分シフト - シート3.csv');

    const csvFacilities = new Set();

    // 練馬: CSVの1行目から施設名取得
    const nerimaFiles = fs.readdirSync(shiftDir).filter(f => f.endsWith('.csv'));
    for (const file of nerimaFiles) {
      const content = fs.readFileSync(path.join(shiftDir, file), 'utf-8');
      const firstLine = content.split('\n')[0];
      const fields = firstLine.split(',');
      if (fields[0].trim() === '施設名' && fields[1]) {
        csvFacilities.add(fields[1].trim());
      }
    }

    // 世田谷: CSVの1行目からヘッダー取得
    const setagayaContent = fs.readFileSync(setagayaPath, 'utf-8');
    const setagayaFirstLine = setagayaContent.split('\n')[0].split(',');
    for (let i = 1; i < setagayaFirstLine.length; i++) {
      const v = setagayaFirstLine[i].trim();
      if (v && v !== '夜') csvFacilities.add(v);
    }

    console.log(`シフトCSVの施設名: ${csvFacilities.size}件`);
    console.log(`Firestore施設: ${firestoreFacilities.length}件\n`);

    // 手動マッピング（自動マッチ不可能なもの）
    const MANUAL_MAP = {
      'グリーンビレッジB': 'GH3',       // 半角B→全角Ｂ
      '春日町同一①': 'GH9',             // 春日町 (B103)
      '春日町２同一①': 'GH12',           // 春日町2 (B203)
      '砧①107': 'GH25',               // 砧 (107)
      '砧②207': 'GH36',               // 砧2 (207)
      '関町南4F同一②': 'GH15',           // 関町南3 (4Fだが施設DBでは関町南3相当)
    };

    // 全角半角変換ヘルパー
    function toFullWidth(s) {
      return s.replace(/[A-Za-z0-9]/g, ch =>
        String.fromCharCode(ch.charCodeAt(0) + 0xFEE0)
      );
    }

    // マッチング
    const mapping = [];
    for (const csvName of [...csvFacilities].sort()) {
      let match = null;
      let matchType = '未マッチ';

      // 手動マッピング優先
      if (MANUAL_MAP[csvName]) {
        match = firestoreFacilities.find(f => f.id === MANUAL_MAP[csvName]);
        if (match) matchType = '手動設定';
      }

      // 完全一致
      if (!match) {
        match = firestoreFacilities.find(f => f.name === csvName);
        if (match) matchType = '完全一致';
      }

      // 全角変換一致
      if (!match) {
        const fw = toFullWidth(csvName);
        match = firestoreFacilities.find(f => f.name === fw);
        if (match) matchType = '全角変換';
      }

      // 部分一致 (CSVの名前がFirestoreの名前に含まれる、または逆)
      if (!match) {
        match = firestoreFacilities.find(f =>
          f.name && (f.name.includes(csvName) || csvName.includes(f.name))
        );
        if (match) matchType = '部分一致';
      }

      // コードで一致
      if (!match) {
        match = firestoreFacilities.find(f =>
          f.code && csvName.includes(f.code)
        );
        if (match) matchType = 'コード一致';
      }

      mapping.push({
        csvName,
        firestoreId: match ? match.id : null,
        firestoreName: match ? match.name : null,
        matchType,
      });
    }

    console.log('CSV施設名'.padEnd(25) + '→ Firestore施設名'.padEnd(30) + 'マッチ');
    console.log('─'.repeat(80));
    for (const m of mapping) {
      const arrow = m.firestoreName ? `→ ${m.firestoreName}` : '→ ???';
      console.log(`${m.csvName.padEnd(25)}${arrow.padEnd(30)}${m.matchType}`);
    }

    const unmatchedCount = mapping.filter(m => !m.firestoreId).length;
    console.log(`\nマッチ: ${mapping.length - unmatchedCount}/${mapping.length}件`);
    if (unmatchedCount > 0) {
      console.log(`未マッチ: ${mapping.filter(m => !m.firestoreId).map(m => m.csvName).join(', ')}`);
    }

    // マッピングJSONを出力
    console.log('\n=== マッピングJSON ===');
    console.log(JSON.stringify(
      Object.fromEntries(mapping.map(m => [m.csvName, { firestoreId: m.firestoreId, firestoreName: m.firestoreName }])),
      null, 2
    ));

  } else if (action === 'collections') {
    // コレクション一覧
    console.log('=== Firestoreルートコレクション一覧 ===\n');
    const collections = await db.listCollections();
    for (const col of collections) {
      const snapshot = await col.limit(1).get();
      const count = snapshot.size > 0 ? '(データあり)' : '(空)';
      console.log(`  ${col.id} ${count}`);
    }
  }

  process.exit(0);
}

main().catch(e => {
  console.error('Error:', e.message);
  process.exit(1);
});
