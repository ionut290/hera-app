// Test con dati sintetici. Non avviare contro servizi reali.
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { initializeTestEnvironment, assertSucceeds, assertFails } = require('@firebase/rules-unit-testing');
const { doc, collection, getDoc, getDocs, setDoc, updateDoc } = require('firebase/firestore');
const projectId = 'demo-varga-multi-tenant';
async function main() {
  if (process.env.FIRESTORE_EMULATOR_HOST !== '127.0.0.1:8180') {
    throw new Error('Avviare esclusivamente il Firestore Emulator locale 127.0.0.1:8180.');
  }
  const env = await initializeTestEnvironment({
    projectId,
    firestore: { host: '127.0.0.1', port: 8180, rules: readFileSync(join(__dirname, 'firestore.rules'), 'utf8') }
  });
  try {
    await env.clearFirestore();
    await env.withSecurityRulesDisabled(async ctx => {
      const db = ctx.firestore();
      const seed = (path, data) => setDoc(doc(db, path), data);
      await Promise.all([
        seed('tenants/a', { name: 'Demo A', status: 'active' }),
        seed('tenants/b', { name: 'Demo B', status: 'active' }),
        seed('tenants/c', { name: 'Demo C', status: 'suspended' }),
        seed('tenants/a/members/alice', { status: 'active', role: 'worker' }),
        seed('tenants/b/members/bob', { status: 'active', role: 'owner' }),
        seed('tenants/b/members/alice', { status: 'revoked', role: 'worker' }),
        seed('tenants/c/members/alice', { status: 'active', role: 'worker' }),
        seed('users/alice/memberships/a', { status: 'active' }),
        seed('users/bob/memberships/b', { status: 'active' }),
        seed('tenants/a/members/adminA', { status: 'active', role: 'admin' }),
        seed('tenants/c/members/adminA', { status: 'active', role: 'admin' }),
        seed('tenants/a/commesse/example', { name: 'Synthetic job A' }),
        seed('tenants/b/commesse/example', { name: 'Synthetic job B' }),
        seed('tenants/c/commesse/example', { name: 'Synthetic job C' }),
        seed('tenants/a/commesse/example/impianti/secret', { name: 'Protected nested record' }),
        seed('commesse/legacy', { name: 'Legacy record must remain inaccessible' })
      ]);
    });
    const alice = env.authenticatedContext('alice', { email_verified: true }).firestore();
    const bob = env.authenticatedContext('bob', { email_verified: true }).firestore();
    const review = env.authenticatedContext('alice', { email_verified: true, apple_review: true }).firestore();
    const unverified = env.authenticatedContext('alice', { email_verified: false }).firestore();
    const anon = env.unauthenticatedContext().firestore();
    const read = (db, path) => getDoc(doc(db, path));
    await assertSucceeds(read(alice, 'tenants/a'));
    await assertSucceeds(read(alice, 'tenants/a/members/alice'));
    await assertSucceeds(read(bob, 'tenants/b'));
    await assertSucceeds(getDocs(collection(alice, 'users/alice/memberships')));
    await assertFails(read(alice, 'tenants/b'));
    await assertFails(read(bob, 'tenants/a'));
    await assertFails(read(alice, 'tenants/c'));
    await assertFails(read(alice, 'tenants/b/members/bob'));
    await assertFails(read(alice, 'users/bob/memberships/b'));
    await assertFails(getDocs(collection(alice, 'tenants')));
    await assertFails(read(anon, 'tenants/a'));
    await assertFails(read(review, 'tenants/a'));
    await assertFails(read(unverified, 'tenants/a'));
    await assertFails(updateDoc(doc(alice, 'tenants/a/members/alice'), { role: 'owner' }));
    await assertFails(setDoc(doc(alice, 'tenants/b/members/alice'), { status: 'active', role: 'owner' }));
    await assertFails(setDoc(doc(alice, 'users/alice/memberships/b'), { status: 'active' }));
    await assertFails(read(alice, 'tenants/a/commesse/example')); // Worker permission not enabled in this pilot.
    const adminA = env.authenticatedContext('adminA', { email_verified: true }).firestore();
    const superAdmin = env.authenticatedContext('platform', { email_verified: true, platform_admin: true }).firestore();
    const adminReview = env.authenticatedContext('adminA', { email_verified: true, apple_review: true }).firestore();
    await assertSucceeds(read(adminA, 'tenants/a/commesse/example'));
    await assertSucceeds(read(bob, 'tenants/b/commesse/example'));
    const own = await assertSucceeds(getDocs(collection(adminA, 'tenants/a/commesse')));
    if (own.size !== 1 || own.docs[0].data().name !== 'Synthetic job A') throw Error('Unexpected tenant A data');
    await assertFails(read(adminA, 'tenants/b/commesse/example'));
    await assertFails(read(bob, 'tenants/a/commesse/example'));
    await assertFails(getDocs(collection(adminA, 'tenants/b/commesse')));
    await assertFails(read(adminA, 'tenants/c/commesse/example'));
    await assertFails(read(superAdmin, 'tenants/a/commesse/example'));
    await assertFails(read(adminReview, 'tenants/a/commesse/example'));
    await assertFails(read(anon, 'tenants/a/commesse/example'));
    await assertFails(read(adminA, 'tenants/a/commesse/example/impianti/secret'));
    await assertFails(read(adminA, 'commesse/legacy'));
    await assertFails(setDoc(doc(adminA, 'tenants/a/commesse/new'), { name: 'Forbidden write' }));
    await assertFails(updateDoc(doc(adminA, 'tenants/a/commesse/example'), { name: 'Forbidden update' }));
    console.log('Tenant membership and read-only commesse pilot assertions passed.');
  } finally {
    await env.cleanup();
  }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
