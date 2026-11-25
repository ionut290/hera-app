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
        seed('tenants/a/commesse/example', { name: 'Synthetic job' })
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
    await assertFails(read(alice, 'tenants/a/commesse/example')); // Not implemented, fail closed.
    console.log('17 tenant-boundary assertions passed.');
  } finally {
    await env.cleanup();
  }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
