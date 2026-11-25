import { installVerdeLevatoPilot } from '../../verde-levato-pilot.mjs';
const empty = Object.freeze({docs:[]});
const readStore = Object.freeze({collection(name){
  if (!['verdeLevatoRecords','verdeLevatoCommesse','verdeLevatoConfig'].includes(name)) throw new Error('Collection denied');
  return Object.freeze({get:async()=>empty,doc:()=>Object.freeze({get:async()=>({exists:false})})});
}});
const pilot=installVerdeLevatoPilot({readStore,user:{uid:'preview-only',email:'preview@example.test'},onClose:()=>location.reload()});
await pilot.open();
