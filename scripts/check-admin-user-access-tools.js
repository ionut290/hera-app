const fs = require('fs');

const read = (file) => fs.readFileSync(file, 'utf8');
const assert = (condition, message) => {
  if (!condition) throw new Error(message);
  console.log(`✓ ${message}`);
};

const client = read('admin-user-access-tools.js');
const backend = read('functions/admin-password-private-requests.js');
const firebaseConfig = read('firebase-config.js');
const functionsMain = read('functions/main.js');
const rules = read('firestore.rules');
const deploy = read('.github/workflows/deploy-firebase-functions.yml');

assert(client.includes('user-management-search-input'), 'Gestione utenti include ricerca per nome o email');
assert(client.includes('type="text"') && client.includes('inputmode="search"'), 'La ricerca usa un campo testo compatibile con iPhone');
assert(client.includes('CAMBIA PASSWORD') && client.includes('user-ban-list'), 'Il pulsante CAMBIA PASSWORD viene aggiunto agli utenti');
assert(client.includes('.collection("privateDocuments")') && client.includes('.collection("adminPasswordRequests")'), 'Il client invia la richiesta nel percorso privato dell’admin');
assert(client.includes('RSA-OAEP') && client.includes('crypto.subtle.generateKey') && client.includes('decryptTemporaryPassword'), 'La password temporanea usa cifratura RSA-OAEP end-to-end');
assert(!client.includes('httpsCallable("adminSetUserPassword")'), 'Il client non dipende più dal callable HTTP bloccato da IAM');
assert(backend.includes('privateDocuments/{userId}/adminPasswordRequests/{requestId}'), 'Il trigger ascolta solo richieste private');
assert(backend.includes('resolveAdministrator') && backend.includes('getAdminEmails'), 'Il backend riverifica il ruolo amministratore');
assert(backend.includes('admin.auth().updateUser') && backend.includes('revokeRefreshTokens'), 'Il backend aggiorna Firebase Auth e revoca le vecchie sessioni');
assert(backend.includes('mustChangePassword: true'), 'Il backend forza il cambio password al primo accesso');
assert(backend.includes('crypto.publicEncrypt') && backend.includes('encryptedTemporaryPassword'), 'Firestore riceve soltanto la password cifrata');
assert(rules.includes('match /privateDocuments/{userId}/{document=**}') && rules.includes('request.auth.uid == userId'), 'Le regole isolano le richieste private per proprietario');
assert(firebaseConfig.includes('admin-user-access-tools.js') && firebaseConfig.includes('HeraAdminUserAccessTools'), 'Firebase config carica gli strumenti gestione utenti');
assert(functionsMain.includes('admin-password-private-requests'), 'functions/main.js esporta il trigger privato');
assert(deploy.includes('functions:processAdminPasswordPrivateRequest'), 'Il workflow distribuisce il trigger privato');

console.log('Controllo Gestione utenti e password privata completato.');
