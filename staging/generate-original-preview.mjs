// Extract only approved original UI and handlers; never bundle app.js or company datasets.
import {readFileSync,writeFileSync} from 'node:fs';
import {fileURLToPath} from 'node:url';
const root=new URL('../',import.meta.url);
const read=name=>readFileSync(new URL(name,root),'utf8');
function exactly(source,needle){const i=source.indexOf(needle);if(i<0||source.indexOf(needle,i+needle.length)>=0)throw Error('Expected one source marker: '+needle);return i;}
export function extractOriginalPreview(){
 const html=read('index.html'),source=read('app.js');
 const start=exactly(html,'<div id="panel-commesse" class="management-panel hidden">');
 const last='<div id="commesse-manage-list" class="commesse-manage-grid"></div>';
 const end=exactly(html,'<section id="prezziario-management-screen"');
 if(exactly(html,last)<start||exactly(html,last)>end)throw Error('Original list missing');
 if(end<=start)throw Error('Invalid original panel');
 const markup=html.slice(start,end)+'</div>';
 if(/<(?:script|iframe|object|embed|link|base|meta)\b|\bon[a-z]+\s*=|\b(?:src|href|action|formaction)\s*=/i.test(markup))throw Error('Unexpected active markup');
 const names=['createCommessa','openNewCommessaDialog','closeNewCommessaDialog','updateCommessaParentField','renderCommesseManagementList','sortCommesseByCreatedAtDesc','getSubcommesse','escapeHTML','createButton','normalizeCommessaDocument','openImpiantiManagement','closeImpiantiManagement','addManualImpianto','getTargetCommessaId','getTargetCommessaName','buildImpiantoKey','mergeMultiValue','splitCodes','hasOrdinario','hasStraordinario','classifyTipoManutenzione','getCommessaCachedImpianti','normalizeManagementSearch','getManagementPlantStatus','getManagementExecution','hasValidManagementCoordinates','getManagementSearchRank','managementFieldValue','populateManagementFilter','renderImpiantiManagementTable','renderManagementPlantRow','updateManagementBulkBar'];
 const constantsStart=exactly(source,'const MANAGEMENT_IMPIANTI_COLUMNS = [');
 const constantsEnd=exactly(source,'function normalizeManagementSearch(value)');
 const constants=source.slice(constantsStart,constantsEnd);
 const handlers=constants+'\n'+names.map(name=>{
 const token='function '+name+'(';const pos=exactly(source,token);const from=source.slice(pos-6,pos)==='async '?pos-6:pos;
 const end=source.indexOf('\n}\n',pos);if(end<pos)throw Error('Invalid handler '+name);
 return source.slice(from,end+2);
 }).join('\n');
 const css=['style.css','management-v2.css','accounting-v2.css'].map(read).join('\n');
 return {markup,handlers,css};
}
if(process.argv[1]===fileURLToPath(import.meta.url)){
 const result=extractOriginalPreview();
 writeFileSync(new URL('staging/portal/src/original-preview.generated.mjs',root),'// Generated from current original source; do not edit.\nexport default '+JSON.stringify(result)+';\n');
 console.log('Original commessa preview extracted');
}
