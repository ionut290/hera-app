const fs = require('fs');
const admin = require('firebase-admin');

function loadServiceAccount() {
  const raw = process.env.FIREBASE_SERVICE_ACCOUNT_HERA_APP_6CD2B;
  if (!raw) {
    throw new Error('Secret FIREBASE_SERVICE_ACCOUNT_HERA_APP_6CD2B non configurato.');
  }
  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new Error('Il secret FIREBASE_SERVICE_ACCOUNT_HERA_APP_6CD2B non contiene JSON valido.');
  }
}

function normalizeEmail(value) {
  return String(value || '').trim().toLowerCase();
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

async function collectFromFirestore(db) {
  const emails = new Set();
  const collectionNames = ['users', 'utenti', 'profiles'];

  for (const collectionName of collectionNames) {
    try {
      const snapshot = await db.collection(collectionName).get();
      snapshot.forEach((doc) => {
        const data = doc.data() || {};
        const candidates = [
          data.email,
          data.mail,
          data.userEmail,
          data.emailAddress,
          data.profilo && data.profilo.email,
          data.profile && data.profile.email,
        ];
        for (const candidate of candidates) {
          const email = normalizeEmail(candidate);
          if (isValidEmail(email)) emails.add(email);
        }
      });
      console.log(`${collectionName}: ${snapshot.size} documenti analizzati`);
    } catch (error) {
      console.warn(`${collectionName}: non leggibile (${error.message})`);
    }
  }

  return emails;
}

async function collectFromAuthentication() {
  const emails = new Set();
  let pageToken;
  do {
    const result = await admin.auth().listUsers(1000, pageToken);
    for (const user of result.users) {
      const email = normalizeEmail(user.email);
      if (isValidEmail(email)) emails.add(email);
    }
    pageToken = result.pageToken;
  } while (pageToken);
  return emails;
}

async function main() {
  admin.initializeApp({ credential: admin.credential.cert(loadServiceAccount()) });

  const allEmails = new Set();
  const authEmails = await collectFromAuthentication();
  const firestoreEmails = await collectFromFirestore(admin.firestore());

  for (const email of authEmails) allEmails.add(email);
  for (const email of firestoreEmails) allEmails.add(email);

  const emails = [...allEmails].sort((a, b) => a.localeCompare(b));
  fs.mkdirSync('exports', { recursive: true });
  fs.writeFileSync('exports/google-play-testers.csv', emails.join('\n') + (emails.length ? '\n' : ''), 'utf8');

  console.log(`Esportazione completata: ${emails.length} email uniche.`);
  if (!emails.length) process.exitCode = 2;
}

main().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
