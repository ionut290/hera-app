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

// The original runtime candidate's auth gate is replaced only in memory.
// Exact source anchors stop the build on unexpected upstream changes.
const startupStart="  if (loggedIn) {\n    try {\n      const savedSession = savedStartupSession || await readPersistedSession();";
const startupEnd="    if (!(await requireBiometricAtStartup())) return;";
const oldAdmin="function canManageData() {\n  const email = normalizeEmail(currentUser?.email || \"\");\n  return isBuiltInSuperAdminEmail(email) || adminEmails.has(email);\n}";
const newAdmin="function canManageData() {\n  const prepared = window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__;\n  return Boolean(prepared?.isActive?.() && currentUser?.uid === prepared.session.context.uid && prepared.session.context.role === \"admin\");\n}";
const newStartup="  if (loggedIn) {\n    try {\n      const prepared = window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__;\n      if (!prepared?.isActive?.() || user.uid !== prepared.session.context.uid) throw new Error(\"Sessione aziendale non valida.\");\n      const authorization = await prepared.approval.verify(user);\n      if (!authorization.allowed || !prepared.isActive()) throw new Error(\"Autorizzazione aziendale non valida.\");\n      await savePersistedSession(user, authorization.profile);\n    } catch (error) {\n      currentUser = null;\n      setAuthenticationGateState(\"required\", \"Accesso aziendale non autorizzato.\");\n      hideStartupLoading();\n      return;\n    }\n";
export function buildIntegratedRuntimeCandidate(source){
 let code=buildScopedOriginalAppCandidate(source);
 const from=code.indexOf(startupStart),end=code.indexOf(startupEnd,from);
 if(from<0||end<0||code.indexOf(startupStart,from+1)>=0)throw Error('Startup gate changed');
 const originalGate=code.slice(from,end);
 if(!originalGate.includes('verifyPersistedSessionAgainstDatabase(user, savedSession)')||!originalGate.includes('window.HeraAccessApproval.verify(user)'))throw Error('Legacy approval markers changed');
 code=replaceExactly(code,originalGate,newStartup);
 code=replaceExactly(code,oldAdmin,newAdmin);
 const builtIn=code.match(/^const BUILT_IN_SUPER_ADMIN_EMAILS = [^\r\n]+;$/gm);
 if(!builtIn||builtIn.length!==1)throw Error('Legacy administrator list changed');
 code=replaceExactly(code,builtIn[0],'const BUILT_IN_SUPER_ADMIN_EMAILS = [];');
 // Add an independent photo archive action; never replace WhatsApp/FATTO attachments.
 const photoAnchor='    secondaryActionsRow.appendChild(safetyQuickBtn);';
 const photoAction=[photoAnchor,'    if (window.parent.__vargaCompanyDocumentsEnabled?.()) {','      const ids = getImpiantoDocIds(impianto);','      const photoJob = String(selectedCommessaId || ""), photoArchive = getCommesseCollectionName();','      const photoButton = document.createElement("button");','      photoButton.type = "button";','      photoButton.className = "btn action-icon-btn";','      photoButton.textContent = "📷";','      photoButton.setAttribute("aria-label", "Archivio fotografie impianto");','      photoButton.dataset.v2PlantPhotos = "1";','      photoButton.disabled = ids.length !== 1 || !photoJob;','      photoButton.title = photoButton.disabled ? "Apri un impianto non aggregato per associare le foto" : "Archivio fotografie impianto";','      photoButton.addEventListener("click", event => {','        event.preventDefault(); event.stopPropagation();','        if (ids.length !== 1 || !photoJob || String(selectedCommessaId || "") !== photoJob || getCommesseCollectionName() !== photoArchive) return;','        window.parent.__vargaOpenCompanyDocuments?.({archive:photoArchive,jobId:photoJob,plantId:String(ids[0])});','      });','      secondaryActionsRow.appendChild(photoButton);','    }'].join('\n');
 code=replaceExactly(code,photoAnchor,photoAction);
 // Bind only the app's unqualified cache identifiers; other modules and window/native
 // APIs are not covered by this transform and must remain in quarantine.
 const binding=['const localStorage = window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__.localStorage;','const sessionStorage = window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__.sessionStorage;',''].join('\n');
 code=binding+code;
 return code;
}
export const integratedStartupGate=newStartup;
export const integratedAdminGate=newAdmin;
