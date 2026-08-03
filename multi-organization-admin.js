import { initializeApp } from 'https://www.gstatic.com/firebasejs/11.10.0/firebase-app.js';
import { getAuth, GoogleAuthProvider, onAuthStateChanged, signInWithPopup, signOut } from 'https://www.gstatic.com/firebasejs/11.10.0/firebase-auth.js';
import { arrayUnion, collection, doc, getDoc, getDocs, getFirestore, limit, query, serverTimestamp, setDoc, writeBatch, where } from 'https://www.gstatic.com/firebasejs/11.10.0/firebase-firestore.js';

const firebaseConfig = {
  apiKey: 'AIzaSyBHZG9D1H5YOT9QUzG-cdSdftlreDJNa_k',
  authDomain: 'hera-app-6cd2b.firebaseapp.com',
  projectId: 'hera-app-6cd2b',
  storageBucket: 'hera-app-6cd2b.firebasestorage.app',
  messagingSenderId: '645390631375',
  appId: '1:645390631375:web:df3659a23812560e4012ba'
};

const BOOTSTRAP_SUPER_ADMIN_EMAIL = 'ionut29019@gmail.com';
const DEFAULT_ORGANIZATION_ID = 'varga';
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const state = { user: null, superAdmin: false, organizations: [], migrationPending: [] };

const $ = (id) => document.getElementById(id);
const normalizeId = (value) => String(value || '').trim().toLowerCase().replace(/[^a-z0-9_-]+/g, '-').replace(/^-+|-+$/g, '');
const message = (text, error = false) => { $('feedback').textContent = text; $('feedback').classList.toggle('error', error); };

async function isSuperAdmin(user) {
  if (!user) return false;
  const snapshot = await getDoc(doc(db, 'platformSuperAdmins', user.uid));
  return snapshot.exists() && snapshot.data().active !== false;
}

async function bootstrapSuperAdmin(user) {
  if (!user || String(user.email || '').toLowerCase() !== BOOTSTRAP_SUPER_ADMIN_EMAIL) return false;
  const ref = doc(db, 'platformSuperAdmins', user.uid);
  const existing = await getDoc(ref);
  if (!existing.exists()) {
    await setDoc(ref, { uid: user.uid, email: user.email, active: true, createdAt: serverTimestamp(), bootstrap: true });
  }
  return true;
}

async function loadOrganizations() {
  const snapshot = await getDocs(collection(db, 'organizations'));
  state.organizations = snapshot.docs.map((item) => ({ id: item.id, ...item.data() })).sort((a, b) => String(a.name || a.id).localeCompare(String(b.name || b.id), 'it'));
  renderOrganizations();
}

function renderOrganizations() {
  const root = $('organizations-list');
  root.replaceChildren();
  if (!state.organizations.length) {
    root.textContent = 'Nessuna organizzazione configurata.';
    return;
  }
  state.organizations.forEach((organization) => {
    const card = document.createElement('article');
    card.className = 'organization-card';
    const title = document.createElement('strong');
    title.textContent = organization.name || organization.id;
    const meta = document.createElement('span');
    meta.textContent = `${organization.id} · ${organization.status || 'attiva'}`;
    card.append(title, meta);
    root.appendChild(card);
  });
}

async function createOrganization(event) {
  event.preventDefault();
  if (!state.superAdmin) throw new Error('Operazione riservata al Super Admin.');
  const name = $('organization-name').value.trim();
  const id = normalizeId($('organization-id').value || name);
  const adminEmail = $('organization-admin-email').value.trim().toLowerCase();
  if (!name || !id) throw new Error('Nome e ID organizzazione sono obbligatori.');

  const orgRef = doc(db, 'organizations', id);
  if ((await getDoc(orgRef)).exists()) throw new Error('Esiste già un’organizzazione con questo ID.');

  let adminProfile = null;
  if (adminEmail) {
    const result = await getDocs(query(collection(db, 'platformUsers'), where('email', '==', adminEmail), limit(1)));
    if (result.empty) throw new Error('L’amministratore deve prima avere un account nell’app.');
    adminProfile = { id: result.docs[0].id, ...result.docs[0].data() };
  }

  const batch = writeBatch(db);
  batch.set(orgRef, {
    id,
    name,
    status: 'attiva',
    createdAt: serverTimestamp(),
    createdByUid: state.user.uid,
    createdByEmail: state.user.email,
    settings: { legacyDataMode: id === DEFAULT_ORGANIZATION_ID }
  });
  if (adminProfile) {
    batch.set(doc(db, 'organizations', id, 'members', adminProfile.id), {
      uid: adminProfile.id,
      email: adminEmail,
      role: 'admin',
      active: true,
      addedAt: serverTimestamp(),
      addedByUid: state.user.uid
    });
    batch.update(doc(db, 'platformUsers', adminProfile.id), {
      organizationMemberships: arrayUnion({ organizationId: id, organizationName: name, role: 'admin', active: true })
    });
  }
  await batch.commit();
  event.target.reset();
  message(`Organizzazione “${name}” creata.`);
  await loadOrganizations();
}

async function buildMigrationDryRun() {
  if (!state.superAdmin) throw new Error('Operazione riservata al Super Admin.');
  message('Analisi utenti in corso…');
  const snapshot = await getDocs(collection(db, 'platformUsers'));
  const pending = [];
  snapshot.forEach((userDoc) => {
    const data = userDoc.data();
    const memberships = Array.isArray(data.organizationMemberships) ? data.organizationMemberships : [];
    if (memberships.length) return;
    const admin = data.isAdmin === true || data.admin === true || ['admin', 'amministratore'].includes(String(data.role || data.ruolo || '').toLowerCase());
    pending.push({ uid: userDoc.id, email: data.email || '', role: admin ? 'admin' : 'operatore' });
  });
  state.migrationPending = pending;
  $('migration-result').textContent = `Utenti analizzati: ${snapshot.size}\nDa assegnare a Varga: ${pending.length}\nScritture previste: ${pending.length * 2}`;
  $('migration-apply').disabled = pending.length === 0;
  message('Dry-run completato. Nessuna scrittura eseguita.');
}

async function applyMigration() {
  const pending = state.migrationPending;
  if (!pending.length) return;
  if (!window.confirm(`Confermi la migrazione di ${pending.length} utenti verso Varga?`)) return;
  let completed = 0;
  for (let offset = 0; offset < pending.length; offset += 200) {
    const batch = writeBatch(db);
    pending.slice(offset, offset + 200).forEach((item) => {
      batch.set(doc(db, 'organizations', DEFAULT_ORGANIZATION_ID, 'members', item.uid), {
        uid: item.uid,
        email: item.email,
        role: item.role,
        active: true,
        migratedAt: serverTimestamp(),
        migratedByUid: state.user.uid
      }, { merge: true });
      batch.update(doc(db, 'platformUsers', item.uid), {
        organizationMemberships: arrayUnion({ organizationId: DEFAULT_ORGANIZATION_ID, organizationName: 'Varga', role: item.role, active: true })
      });
    });
    await batch.commit();
    completed += Math.min(200, pending.length - offset);
    message(`Migrazione utenti: ${completed}/${pending.length}`);
  }
  state.migrationPending = [];
  $('migration-apply').disabled = true;
  await buildMigrationDryRun();
}

function bindEvents() {
  $('login').addEventListener('click', () => signInWithPopup(auth, new GoogleAuthProvider()).catch((error) => message(error.message, true)));
  $('logout').addEventListener('click', () => signOut(auth));
  $('organization-form').addEventListener('submit', (event) => createOrganization(event).catch((error) => message(error.message, true)));
  $('refresh-organizations').addEventListener('click', () => loadOrganizations().catch((error) => message(error.message, true)));
  $('migration-dry-run').addEventListener('click', () => buildMigrationDryRun().catch((error) => message(error.message, true)));
  $('migration-apply').addEventListener('click', () => applyMigration().catch((error) => message(error.message, true)));
}

bindEvents();
onAuthStateChanged(auth, async (user) => {
  state.user = user;
  $('login').hidden = Boolean(user);
  $('logout').hidden = !user;
  $('session').textContent = user ? `${user.displayName || user.email} · ${user.uid}` : 'Non autenticato';
  document.querySelectorAll('[data-protected]').forEach((element) => { element.hidden = !user; });
  if (!user) return;
  try {
    state.superAdmin = await isSuperAdmin(user);
    if (!state.superAdmin && await bootstrapSuperAdmin(user)) state.superAdmin = await isSuperAdmin(user);
    $('access-state').textContent = state.superAdmin ? 'Super Admin attivo' : 'Accesso negato: account non Super Admin';
    document.querySelectorAll('[data-super-admin]').forEach((element) => { element.hidden = !state.superAdmin; });
    if (state.superAdmin) await loadOrganizations();
  } catch (error) {
    message(error.message, true);
  }
});
