const fs = require('fs');

const read = (file) => fs.readFileSync(file, 'utf8');
const assert = (condition, message) => {
  if (!condition) throw new Error(message);
  console.log(`✓ ${message}`);
};

const client = read('admin-password-private-bridge.js');
const backend = read('functions/admin-password-private-requests.js');
const firebaseConfig = read('firebase-config.js');
const rules = read('firestore.rules');
const deploy = read('.github/workflows/deploy-firebase-functions.yml');
const functionsMain = read('functions/main.js');

assert(client.includes('.collection("privateDocuments")') && client.includes('.collection("adminPasswordRequests")'), 'Il client usa la sottocollezione privata dell’amministratore');
assert(client.includes('RSA-OAEP') && client.includes('window.crypto.subtle.generateKey') && client.includes('decryptTemporaryPassword'), 'La password temporanea viaggia cifrata con chiave effimera RSA-OAEP');
assert(client.includes('[data-admin-manage-password]') && client.includes('stopImmediatePropagation'), 'Il bridge sostituisce il vecchio pulsante callable senza doppie chiamate');
assert(backend.includes('privateDocuments/{userId}/adminPasswordRequests/{requestId}'), 'Il trigger ascolta soltanto richieste nel percorso privato dell’utente');
assert(backend.includes('resolveAdministrator') && backend.includes('getAdminEmails'), 'Il backend riverifica il ruolo amministratore');
assert(backend.includes('admin.auth().updateUser') && backend.includes('revokeRefreshTokens'), 'Il backend aggiorna Firebase Auth e revoca le vecchie sessioni');
assert(backend.includes('crypto.publicEncrypt') && backend.includes('encryptedTemporaryPassword'), 'Il backend restituisce solo la password cifrata');
assert(rules.includes('match /privateDocuments/{userId}/{document=**}') && rules.includes('request.auth.uid == userId'), 'Le regole esistenti isolano le richieste nel percorso privato dell’admin');
assert(firebaseConfig.includes('admin-password-private-bridge.js') && firebaseConfig.includes('HeraAdminPasswordPrivateBridge'), 'Firebase config carica il nuovo bridge password');
assert(functionsMain.includes('admin-password-private-requests'), 'functions/main.js esporta il nuovo trigger');
assert(deploy.includes('functions:processAdminPasswordPrivateRequest'), 'Il workflow distribuisce il trigger privato');

console.log('Controllo cambio password privato completato.');
