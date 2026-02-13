#!/usr/bin/env node
/**
 * Firestore全施設一覧取得 + シフトCSVとの差分確認
 */
const path = require('path');
const { initializeApp, cert } = require('firebase-admin/app');
const { getFirestore } = require('firebase-admin/firestore');

const SA_PATH = process.env.FIREBASE_SA_PATH || path.join(
  __dirname, '..', '..', 'expense-management-system', 'credentials', 'service-account.json'
);
const app = initializeApp({ credential: cert(require(SA_PATH)), projectId: 'safe-rise-prod' });
const db = getFirestore(app);

async function main() {
  const snapshot = await db.collection('facilities').get();
  const facilities = [];
  snapshot.forEach(doc => {
    const d = doc.data();
    facilities.push({
      id: doc.id,
      name: d.name || '',
      code: d.code || '',
      is_active: d.is_active !== false,
      staff_count: Array.isArray(d.staff_ids) ? d.staff_ids.length : 0,
      user_count: Array.isArray(d.care_user_ids) ? d.care_user_ids.length : 0,
      address: d.address || '',
    });
  });
  facilities.sort((a, b) => a.name.localeCompare(b.name, 'ja'));

  const active = facilities.filter(f => f.is_active);
  const inactive = facilities.filter(f => !f.is_active);

  console.log(`=== Firestore全施設一覧 (${facilities.length}施設) ===\n`);
  console.log(`■ アクティブ施設 (${active.length}件)`);
  console.log('  ' + 'ID'.padEnd(22) + '施設名'.padEnd(22) + 'コード'.padEnd(12) + 'Staff  利用者');
  console.log('  ' + '─'.repeat(70));
  for (const f of active) {
    console.log(`  ${f.id.padEnd(22)}${f.name.padEnd(22)}${f.code.padEnd(12)}${String(f.staff_count).padEnd(7)}${f.user_count}`);
  }

  if (inactive.length) {
    console.log(`\n■ 非アクティブ施設 (${inactive.length}件)`);
    for (const f of inactive) {
      console.log(`  ${f.id.padEnd(22)}${f.name.padEnd(22)}${f.code.padEnd(12)}${String(f.staff_count).padEnd(7)}${f.user_count}`);
    }
  }

  // 今回シフトCSVに登場した施設(マッピング済み)
  const csvMapped = new Set([
    'グリーンビレッジＢ', 'グリーンビレッジＥ', 'グリーンビレッジE102',
    '中町', '南大泉', '南大泉３丁目', '大泉町',
    '春日町 (B103)', '春日町2 (B203)',
    '東大泉', '松原', '石神井公園',
    '砧 (107)', '砧2 (207)',
    '西長久保', '都民農園', '長久保',
    '関町南2', '関町南3'
  ]);

  const notInCsv = active.filter(f =>
    !csvMapped.has(f.name) &&
    !f.name.includes('テスト') &&
    f.name !== '◆退去◆' &&
    f.name !== 'SSエルウィング'
  );

  console.log(`\n■ 今回のシフトCSVに未登場のアクティブ施設 (${notInCsv.length}件)`);
  console.log('  これらの施設のシフトデータが別のCSVファイルにある可能性があります:');
  for (const f of notInCsv) {
    console.log(`  ${f.id.padEnd(22)}${f.name}`);
  }

  console.log(`\n=== サマリー ===`);
  console.log(`  Firestore全施設: ${facilities.length}件`);
  console.log(`  アクティブ: ${active.length}件`);
  console.log(`  今回CSVマッピング済: ${csvMapped.size}件`);
  console.log(`  CSV未登場(要確認): ${notInCsv.length}件`);

  process.exit(0);
}

main().catch(e => { console.error('Error:', e.message); process.exit(1); });
