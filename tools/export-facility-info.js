#!/usr/bin/env node
const path = require('path');
const fs = require('fs');
const { initializeApp, cert } = require('firebase-admin/app');
const { getFirestore } = require('firebase-admin/firestore');

const SA_PATH = process.env.FIREBASE_SA_PATH || path.join(
  __dirname, '..', '..', 'expense-management-system', 'credentials', 'service-account.json'
);
const sa = require(SA_PATH);
const app = initializeApp({ credential: cert(sa), projectId: 'safe-rise-prod' });
const db = getFirestore(app);

async function main() {
  const snap = await db.collection('facilities').get();
  const info = {};
  snap.forEach(doc => {
    const d = doc.data();
    if (doc.id === 'GH000' || doc.id.startsWith('facility_test')) return;
    info[doc.id] = {
      name: d.name || '',
      reading: d.address || '',
      room: d.phone_number || d.phone || '',
      userCount: (d.care_user_ids || []).length,
      isActive: d.is_active !== false,
    };
  });
  const outPath = path.join(__dirname, '..', 'data', 'facility-info.json');
  fs.writeFileSync(outPath, JSON.stringify(info, null, 2));
  console.log('facility-info.json saved:', Object.keys(info).length, 'facilities');
  process.exit(0);
}

main().catch(e => { console.error(e.message); process.exit(1); });
