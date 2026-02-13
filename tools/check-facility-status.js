#!/usr/bin/env node
/**
 * safe-rise-prod Firestoreから施設の稼働状況を確認
 * - facilities: 利用者数・スタッフ数
 * - staff: 各施設に配置されているスタッフ
 * - care_users: 各施設の利用者
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
  // 1. 全施設取得
  console.log('施設データを取得中...');
  const facSnap = await db.collection('facilities').get();
  const facilities = new Map();
  facSnap.forEach(doc => {
    const d = doc.data();
    facilities.set(doc.id, {
      id: doc.id,
      name: d.name || '',
      code: d.code || '',
      is_active: d.is_active !== false,
      staff_ids: d.staff_ids || [],
      care_user_ids: d.care_user_ids || [],
      address: d.address || '',
      phone: d.phone_number || d.phone || '',
      google_account_email: d.google_account_email || '',
    });
  });
  console.log(`  ${facilities.size}施設\n`);

  // 2. スタッフ取得
  console.log('スタッフデータを取得中...');
  const staffSnap = await db.collection('staff').get();
  const staffByFacility = new Map();
  const allStaff = [];
  staffSnap.forEach(doc => {
    const d = doc.data();
    allStaff.push({
      id: doc.id,
      name: d.name || d.staff_name || '',
      facility_id: d.facility_id || '',
      facility_name: d.facility_name || '',
      is_active: d.is_active !== false,
    });
    const fid = d.facility_id || 'unknown';
    if (!staffByFacility.has(fid)) staffByFacility.set(fid, []);
    staffByFacility.get(fid).push({ id: doc.id, name: d.name || d.staff_name || '', is_active: d.is_active !== false });
  });
  console.log(`  ${allStaff.length}名\n`);

  // 3. 利用者取得
  console.log('利用者データを取得中...');
  const userSnap = await db.collection('care_users').get();
  const usersByFacility = new Map();
  let totalUsers = 0;
  userSnap.forEach(doc => {
    const d = doc.data();
    totalUsers++;
    const fid = d.facility_id || 'unknown';
    if (!usersByFacility.has(fid)) usersByFacility.set(fid, []);
    usersByFacility.get(fid).push({ id: doc.id, name: d.name || '', is_active: d.is_active !== false });
  });
  console.log(`  ${totalUsers}名\n`);

  // 4. コレクション一覧確認（シフト関連がないか）
  console.log('コレクション一覧を確認中...');
  const collections = await db.listCollections();
  const colNames = [];
  for (const col of collections) {
    colNames.push(col.id);
  }
  console.log(`  ${colNames.length}コレクション: ${colNames.join(', ')}\n`);

  // シフト関連コレクションを探す
  const shiftRelated = colNames.filter(c =>
    c.includes('shift') || c.includes('シフト') || c.includes('schedule') ||
    c.includes('attendance') || c.includes('勤務')
  );
  if (shiftRelated.length > 0) {
    console.log(`  シフト関連コレクション: ${shiftRelated.join(', ')}`);
    for (const col of shiftRelated) {
      const snap = await db.collection(col).limit(3).get();
      console.log(`  ${col}: ${snap.size}件 (サンプル)`);
      snap.forEach(doc => {
        console.log(`    ${doc.id}: ${JSON.stringify(doc.data()).substring(0, 200)}`);
      });
    }
  } else {
    console.log('  シフト関連コレクションは見つかりませんでした');
  }

  // 5. 施設別サマリー
  console.log('\n========================================');
  console.log('  施設別 稼働状況（Firestoreデータ）');
  console.log('========================================\n');

  const sortedFacilities = [...facilities.values()]
    .filter(f => !f.name.includes('テスト') && f.name !== '◆退去◆')
    .sort((a, b) => a.name.localeCompare(b.name, 'ja'));

  console.log('  ' + '施設名'.padEnd(25) + 'コード'.padEnd(10) + '利用者'.padEnd(8) + 'Staff(DB)'.padEnd(12) + 'Staff(ref)'.padEnd(12) + '住所');
  console.log('  ' + '─'.repeat(100));

  for (const f of sortedFacilities) {
    const dbStaff = staffByFacility.get(f.id) || [];
    const activeDbStaff = dbStaff.filter(s => s.is_active);
    const dbUsers = usersByFacility.get(f.id) || [];
    const activeDbUsers = dbUsers.filter(u => u.is_active);
    const refUsers = f.care_user_ids.length;

    console.log(
      `  ${f.name.padEnd(25)}${f.code.padEnd(10)}${String(activeDbUsers.length || refUsers).padEnd(8)}${String(activeDbStaff.length).padEnd(12)}${String(f.staff_ids.length).padEnd(12)}${f.address.substring(0, 30)}`
    );
  }

  // 6. 利用者がいる施設（実稼働中と推定）
  const operatingFacilities = sortedFacilities.filter(f => {
    const users = usersByFacility.get(f.id) || [];
    return users.filter(u => u.is_active).length > 0 || f.care_user_ids.length > 0;
  });

  console.log(`\n■ 利用者がいる施設（実稼働中）: ${operatingFacilities.length}件`);
  for (const f of operatingFacilities) {
    const users = usersByFacility.get(f.id) || [];
    const activeUsers = users.filter(u => u.is_active).length || f.care_user_ids.length;
    console.log(`  ${f.name.padEnd(25)}${f.code.padEnd(10)}利用者${activeUsers}名`);
  }

  const noUsersFacilities = sortedFacilities.filter(f => {
    const users = usersByFacility.get(f.id) || [];
    return users.filter(u => u.is_active).length === 0 && f.care_user_ids.length === 0;
  });
  if (noUsersFacilities.length > 0) {
    console.log(`\n■ 利用者0の施設: ${noUsersFacilities.length}件`);
    for (const f of noUsersFacilities) {
      console.log(`  ${f.name.padEnd(25)}${f.code}`);
    }
  }

  process.exit(0);
}

main().catch(e => { console.error('Error:', e.message); process.exit(1); });
