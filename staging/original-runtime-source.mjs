// Candidate generation only. Never load this candidate in a live browser or native app.
// The original app.js remains byte-for-byte unchanged on disk. Replace only its
// SDK initialization statements, not the protected application functions.
export const originalSdkInit=[
 '  firebaseApp = firebase.apps && firebase.apps.length',
 '    ? firebase.app()',
 '    : firebase.initializeApp(firebaseConfig);',
 '  auth = typeof firebase.auth === "function" ? firebase.auth() : null;',
 '  db = typeof firebase.firestore === "function" ? firebase.firestore() : null;',
 '  functions = typeof firebase.functions === "function" ? firebase.functions() : null;'
].join('\n');
export const scopedSdkInit=[
 '  const prepared = window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__;',
 '  if (!prepared || prepared.projectId !== "varga-cantieri-v2-staging" || !prepared.session?.firestore || !prepared.auth || prepared.auth.currentUser?.uid !== prepared.session.context?.uid) {',
 '    throw new Error("Sessione aziendale verificata obbligatoria.");',
 '  }',
 '  if (firebaseConfig.projectId !== prepared.projectId || prepared.auth.app?.options?.projectId !== prepared.projectId) {',
 '    throw new Error("Progetto staging non valido.");',
 '  }',
 '  firebaseApp = prepared.auth.app;',
 '  auth = prepared.auth;',
 '  db = prepared.session.firestore;',
 '  functions = null; // External callable services are disabled in this candidate.'
].join('\n');
function replaceExactly(source,from,to){
 if(source.split(from).length!==2)throw Error('Original SDK initialization changed; review required.');
 return source.replace(from,to);
}
export function buildScopedOriginalAppCandidate(source){
 if(typeof source!=='string'||source.length<100000||!source.includes('enableFirestorePersistence();'))throw Error('Complete original app source required');
 const code=replaceExactly(source,originalSdkInit,scopedSdkInit);
 // Verify transformation reversible. The entirety of the remaining original
 // functions, including FATTO/WhatsApp, is not touched.
 if(replaceExactly(code,scopedSdkInit,originalSdkInit)!==source)throw Error('Unexpected source change');
 return code;
}
