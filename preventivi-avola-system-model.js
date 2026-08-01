(()=>{
  'use strict';
  const M=window.HeraPreventiviModels,P=M?.PV;
  if(!M||!P||M.runtime.avolaSystemModel)return;
  M.runtime.avolaSystemModel=true;
  const VERSION='20260801d',NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main',RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships',PRNS='http://schemas.openxmlformats.org/package/2006/relationships',XNS='http://www.w3.org/XML/1998/namespace';
  const parser=new DOMParser(),serializer=new XMLSerializer(),C=v=>String(v??'').trim(),N=v=>C(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
  const fields=[
    {key:'numero_preventivo',label:'Numero preventivo',type:'text',source:'number',automatic:true,calculated:true},
    {key:'data_documento',label:'Data preventivo',type:'date',source:'date',automatic:true,calculated:true},
    {key:'cliente',label:'Cliente / Ragione sociale',type:'textarea',source:'clientName',required:true},
    {key:'richiedente',label:'C.A. / Richiedente',type:'text',source:'requester'},
    {key:'oggetto',label:'Oggetto',type:'textarea',source:'subject',required:true},
    {key:'lavorazioni',label:'Lavorazioni',type:'repeater',source:'lavorazioni',required:true,automatic:true,calculated:true}
  ];
  const isAvola=model=>{const text=N(`${model?.name||''} ${model?.originalName||''}`);return model?.format==='xlsx'&&(text.includes('standard avola')||text.includes('matrice preventivo avola fedele'))};
  function updateExisting(){
    const models=P.state.models||[],model=models.find(isAvola);if(!model)return null;
    const signature=JSON.stringify(model.fields||[]),next=JSON.stringify(fields);
    Object.assign(model,{documentType:'preventivo',active:true,systemModel:true,fieldDetectionVersion:VERSION,fields:fields.map(x=>({...x})),compatibility:'XLSX originale compilato, logo e impaginazione invariati',analysis:{...(model.analysis||{}),fieldDetectionVersion:VERSION,recoveredFields:fields.length,avolaTemplate:true}});
    if(signature!==next){model.version=Math.max(1,Number(model.version)||1)+1;model.updatedAt=P.nowIso();model.updatedBy=P.currentUser();model.syncPending=true;P.persistLocal?.();P.scheduleSync?.();}
    return model;
  }
  const parse=text=>{const doc=parser.parseFromString(text,'application/xml');if(doc.querySelector('parsererror'))throw Error('La matrice Avola non è leggibile.');return doc};
  const sheetPath=zip=>{const wb=parse(zip.file('xl/workbook.xml').asText()),rels=parse(zip.file('xl/_rels/workbook.xml.rels').asText()),sheet=wb.getElementsByTagNameNS(NS,'sheet')[0],id=sheet?.getAttributeNS(RNS,'id')||sheet?.getAttribute('r:id'),rel=[...rels.getElementsByTagNameNS(PRNS,'Relationship')].find(x=>x.getAttribute('Id')===id),target=rel?.getAttribute('Target')?.replace(/^\/+/, '');if(!target)throw Error('Foglio Preventivo Avola non trovato.');return target.startsWith('xl/')?target:`xl/${target}`};
  const colN=s=>[...String(s||'').toUpperCase()].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0),point=r=>{const m=String(r||'').match(/^([A-Z]+)(\d+)$/i);return m?{c:colN(m[1]),r:Number(m[2])}:null};
  const row=(doc,n)=>{let r=[...doc.getElementsByTagNameNS(NS,'row')].find(x=>Number(x.getAttribute('r'))===n);if(r)return r;r=doc.createElementNS(NS,'row');r.setAttribute('r',String(n));const data=doc.getElementsByTagNameNS(NS,'sheetData')[0],next=[...data.children].find(x=>Number(x.getAttribute('r'))>n);data.insertBefore(r,next||null);return r};
  const cell=(doc,ref)=>{let c=[...doc.getElementsByTagNameNS(NS,'c')].find(x=>x.getAttribute('r')===ref);if(c)return c;const p=point(ref),rw=row(doc,p.r);c=doc.createElementNS(NS,'c');c.setAttribute('r',ref);const next=[...rw.children].find(x=>colN((x.getAttribute('r')||'').replace(/\d+/g,''))>p.c);rw.insertBefore(c,next||null);return c};
  const clear=c=>{[...c.children].forEach(x=>x.remove());c.removeAttribute('t')};
  const text=(doc,ref,value)=>{const c=cell(doc,ref),s=c.getAttribute('s');clear(c);if(s!==null)c.setAttribute('s',s);c.setAttribute('t','inlineStr');const is=doc.createElementNS(NS,'is'),t=doc.createElementNS(NS,'t');t.setAttributeNS(XNS,'xml:space','preserve');t.textContent=String(value??'');is.appendChild(t);c.appendChild(is)};
  const number=(doc,ref,value)=>{const c=cell(doc,ref),s=c.getAttribute('s');clear(c);if(s!==null)c.setAttribute('s',s);const v=doc.createElementNS(NS,'v');v.textContent=String(Number(value)||0);c.appendChild(v)};
  const formula=(doc,ref,f,cached=0)=>{const c=cell(doc,ref),s=c.getAttribute('s');clear(c);if(s!==null)c.setAttribute('s',s);const fn=doc.createElementNS(NS,'f'),v=doc.createElementNS(NS,'v');fn.textContent=f;v.textContent=String(Number(cached)||0);c.append(fn,v)};
  const dateIt=v=>{const m=C(v).match(/^(\d{4})-(\d{2})-(\d{2})/);return m?`${m[3]}/${m[2]}/${m[1]}`:C(v)};
  async function exportAvola(model,stored,data){
    await M.ensureScript('PizZip',M.constants.ZIP_LIB,'pizzip');const zip=new PizZip(stored.buffer),path=sheetPath(zip),entry=zip.file(path);if(!entry)throw Error('File originale STANDARD AVOLA non disponibile.');
    const doc=parse(entry.asText()),lines=Array.isArray(data.lavorazioni)?data.lavorazioni:[];if(lines.length>12)throw Error('La matrice STANDARD AVOLA contiene 12 righe lavorazioni.');
    text(doc,'H2',`Offerta n° ${C(data.numero_preventivo||data.number)} del ${dateIt(data.data_documento||data.date)}`);text(doc,'H6',C(data.cliente||data.clientName));text(doc,'H11',`C.A.: ${C(data.richiedente||data.requester)}`);text(doc,'D17',C(data.oggetto||data.subject));
    for(let i=0;i<12;i+=1){const r=24+i,line=lines[i];text(doc,`B${r}`,line?.codice||'');text(doc,`D${r}`,line?.descrizione||'');text(doc,`J${r}`,line?.unita_misura||'');if(line){number(doc,`K${r}`,line.quantita);number(doc,`L${r}`,line.prezzo_unitario);formula(doc,`N${r}`,`K${r}*L${r}`,line.totale_riga)}else{text(doc,`K${r}`,'');text(doc,`L${r}`,'');text(doc,`N${r}`,'')}}
    formula(doc,'N37',`SUM(N24:N${Math.max(24,23+lines.length)})`,data.totale_imponibile);zip.file(path,serializer.serializeToString(doc));zip.remove('xl/calcChain.xml');M.download(zip.generate({type:'blob',mimeType:stored.type||'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',compression:'DEFLATE'}),'Matrice_Preventivo_Avola-compilato.xlsx');
  }
  const previous=M.exportSheet?.bind(M);if(previous)M.exportSheet=(model,stored,data,doc)=>isAvola(model)?exportAvola(model,stored,data):previous(model,stored,data,doc);
  updateExisting();[800,2500,6000].forEach(ms=>setTimeout(updateExisting,ms));window.addEventListener('focus',()=>setTimeout(updateExisting,200));
})();
