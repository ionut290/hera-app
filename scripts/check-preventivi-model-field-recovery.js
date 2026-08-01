'use strict';
const assert=require('assert');
const fs=require('fs');
const path=require('path');
const cp=require('child_process');
const root=path.join(__dirname,'..');
const recoveryPath=path.join(root,'preventivi-model-field-recovery.js');
const followupPath=path.join(root,'preventivi-registry-model-followup.js');
const swPath=path.join(root,'sw.js');
[recoveryPath,followupPath,swPath].forEach(file=>cp.execFileSync(process.execPath,['--check',file],{stdio:'pipe'}));
const recovery=fs.readFileSync(recoveryPath,'utf8');
const followup=fs.readFileSync(followupPath,'utf8');
const sw=fs.readFileSync(swPath,'utf8');
[
  'nome impianto',
  'id sap',
  'comune',
  'odl',
  'descrizione intervento',
  'richiedente intervento',
  'ditta esecutrice se diversa da avola',
  'data richiesta',
  'data esecuzione',
  'competenza bologna ovest',
  'competenza bologna est',
  'cod prest est',
  'importo prestazione',
  "add({key:'commessa'",
  'fieldDetectionVersion',
  'recoverModel(model,form=null)',
  'M.analyzeFile=async file',
  'input[type="checkbox"][data-pvm-field]'
].forEach(token=>assert(recovery.includes(token),`Manca nel recupero campi: ${token}`));
assert(followup.includes('preventivi-model-field-recovery.js?v=20260801c'),'Modulo recupero campi non caricato.');
assert(followup.indexOf('preventivi-model-field-recovery.js')<followup.indexOf('preventivi-model-driven-form.js'),'Il recupero campi deve caricarsi prima del form dinamico.');
assert(sw.includes('varga-cantieri-shell-v87'),'Cache PWA non aggiornata a v87.');
assert(sw.includes('preventivi-model-field-recovery.js?v=20260801c'),'Modulo recupero campi assente dalla cache.');
console.log('check-preventivi-model-field-recovery: OK');
