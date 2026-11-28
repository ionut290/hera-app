// Must run before all original script tags. The root source stays quarantined.
(() => {
 if(window.parent===window||!['localhost','127.0.0.1'].includes(location.hostname))throw Error('Local company launcher required');
 const {session,namespace,apiKey}=window.parent.__vargaLocalRuntime();
 if(!session.isActive())throw Error('Company closed');
 // The original runtime is loaded inside the launcher iframe. On iOS the
 // iframe can report zero safe-area insets even though the native shell is
 // drawing below the Dynamic Island/notch. Mark the frame explicitly so the
 // shared ios-safe-area.css fallback is applied to every original screen.
 const parentCap=window.parent.Capacitor;
 const nativeIOS=location.protocol==='capacitor:'
  || Boolean(parentCap?.isNativePlatform?.()&&parentCap?.getPlatform?.()==='ios')
  || /iPhone|iPad|iPod/i.test(navigator.userAgent||'');
 if(nativeIOS)document.documentElement.classList.add('varga-ios-native-safe-area');
 Object.defineProperty(window,'__VARGA_ISOLATED_RUNTIME_CANDIDATE__',{value:session,writable:false});
 Object.defineProperty(window,'firebase',{value:namespace,writable:false});
 Object.defineProperty(window,'localStorage',{value:session.localStorage,configurable:true});
 Object.defineProperty(window,'sessionStorage',{value:session.sessionStorage,configurable:true});
 window.__VARGA_STAGING_FIREBASE_CONFIG__={apiKey};
 document.addEventListener('click',event=>{if(event.target?.closest?.('#open-private-docs-btn')&&window.parent.__vargaOpenCompanyDocuments?.()){event.preventDefault();event.stopImmediatePropagation();}},true);
})();
