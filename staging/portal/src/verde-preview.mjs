import { installVerdeLevatoPilot } from '../../verde-levato-pilot.mjs';
const records = Object.freeze([
 {id:'demo-c1',tipoRecord:'cantiere',denominazione:'Cantiere dimostrativo A',codice:'DEMO-C01',comune:'Comune fittizio',via:'Via Esempio 1'},
 {id:'demo-a1',tipoRecord:'albero',denominazione:'Albero dimostrativo',codice:'DEMO-A01',comune:'Comune fittizio',specie:'Acero'},
 {id:'demo-s1',tipoRecord:'siepe',denominazione:'Siepe dimostrativa',codice:'DEMO-S01',comune:'Comune fittizio',specie:'Lauro'}
]);
const readStore = Object.freeze({collection(name){
 if (!['verdeLevatoRecords','verdeLevatoCommesse','verdeLevatoConfig'].includes(name)) throw new Error('Collection denied');
 return Object.freeze({get:async()=>({docs:name==='verdeLevatoRecords'?records.map(({id,...data})=>({id,data:()=>({...data})})):[]}),doc:()=>Object.freeze({get:async()=>({exists:false})})});
}});
const pilot=installVerdeLevatoPilot({readStore,user:{uid:'preview-only',email:'preview@example.test'},onClose:()=>location.reload()});
await pilot.open();
