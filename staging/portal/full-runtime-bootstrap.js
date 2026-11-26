// Must run before all original script tags. The root source stays quarantined.
(() => {
 if(window.parent===window||!['localhost','127.0.0.1'].includes(location.hostname))throw Error('Local company launcher required');
 const {session,namespace,apiKey}=window.parent.__vargaLocalRuntime();
 if(!session.isActive())throw Error('Company closed');
 Object.defineProperty(window,'__VARGA_ISOLATED_RUNTIME_CANDIDATE__',{value:session,writable:false});
 Object.defineProperty(window,'firebase',{value:namespace,writable:false});
 Object.defineProperty(window,'localStorage',{value:session.localStorage,configurable:true});
 Object.defineProperty(window,'sessionStorage',{value:session.sessionStorage,configurable:true});
 window.__VARGA_STAGING_FIREBASE_CONFIG__={apiKey};
})();