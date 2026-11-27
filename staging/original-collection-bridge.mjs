// Compatibility boundary for the seven original normal/snow archive selectors.
// Not installed in app.js; no direct Firebase operations or protected function changes.
const names=Object.freeze({
 getCommesseCollectionName:['commesse','neve_commesse'],
 getSquadreCurrentCollectionName:['squadreCommesse','neve_squadre'],
 getSquadreHistoryCollectionName:['squadreStorico','neve_squadre_storico'],
 getOreReportsCollectionName:['oreReports','neve_ore'],
 getOreApprovalRequestsCollectionName:['oreApprovalRequests','neve_ore_richieste'],
 getPersonaleCollectionName:['personale','neve_personale'],
 getMezziCollectionName:['mezzi','neve_mezzi']
});
export function createOriginalCollectionBridge(router){
 if(typeof router?.path!=='function'||typeof router?.close!=='function')throw Error('Router required');
 const selectors=Object.fromEntries(Object.entries(names).map(([key,pair])=>[key,(snow=false)=>{
 if(typeof snow!=='boolean')throw Error('Explicit service context required');
 return router.path(pair[snow?1:0]);
 }]));
 return Object.freeze({...selectors,close:()=>router.close()});
}
