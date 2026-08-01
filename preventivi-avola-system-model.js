(()=>{
  'use strict';
  const M=window.HeraPreventiviModels,P=M?.PV;
  if(!M||!P||M.runtime.avolaSystemModel)return;
  M.runtime.avolaSystemModel=true;

  const MODEL_ID='model-standard-avola-app';
  const MODEL_NAME='STANDARD AVOLA – MODELLO APP';
  const ORIGINAL_NAME='Matrice_Preventivo_Avola_Modello_App.xlsx';
  const FILE_URL='./models/Matrice_Preventivo_Avola_Modello_App.xlsx?v=20260801a';
  const DETECTION_VERSION='20260801c';
  const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PRNS='http://schemas.openxmlformats.org/package/2006/relationships';
  const XNS='http://www.w3.org/XML/1998/namespace';
  const parser=new DOMParser(),serializer=new XMLSerializer();
  const C=v=>String(v??'').trim();
  const N=v=>C(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();

  const fields=[
    {key:'numero_preventivo',label:'Numero preventivo',type:'text',source:'number',required:false,automatic:true,calculated:true,token:'{{numero_preventivo}}'},
    {key:'data_documento',label:'Data preventivo',type:'date',source:'date',required:false,automatic:true,calculated:true,token:'{{data_documento}}'},
    {key:'cliente',label:'Cliente / Ragione sociale',type:'textarea',source:'clientName',required:true,automatic:false,token:'{{cliente}}'},
    {key:'richiedente',label:'C.A. / Richiedente',type:'text',source:'requester',required:false,automatic:false,token:'{{richiedente}}'},
    {key:'oggetto',label:'Oggetto',type:'textarea',source:'subject',required:true,automatic:false,token:'{{oggetto}}'},
    {key:'lavorazioni',label:'Lavorazioni',type:'repeater',source:'lavorazioni',required:true,automatic:true,calculated:true,detectedFromLabel:'tabella lavorazioni'}
  ];

  function ensureModel(){
    P.state.models=P.state.models||[];
    let model=P.state.models.find(item=>item.id===MODEL_ID)||P.state.models.find(item=>{
      const text=N(`${item.name||''} ${item.originalName||''}`);
      return text.includes('standard avola')||text.includes('matrice preventivo avola fedele pdf');
    });
    const now=P.nowIso?.()||new Date().toISOString();
    const data={
      id:model?.id||MODEL_ID,
      name:MODEL_NAME,
      documentType:'preventivo',
      description:'Matrice Avola ufficiale integrata nell’app. Compila cliente, richiedente, oggetto e lavorazioni mantenendo il foglio originale.',
      format:'xlsx',
      originalName:ORIGINAL_NAME,
      fileUrl:FILE_URL,
      active:true,
      systemModel:true,
      builtIn:false,
      version:Math.max(1,Number(model?.version)||1),
      fields:fields.map(field=>({...field})),
      fieldDetectionVersion:DETECTION_VERSION,
      analysis:{...(model?.analysis||{}),sheets:['Preventivo Avola'],fieldDetectionVersion:DETECTION_VERSION,recoveredFields:fields.length,systemTemplate:true},
      compatibility:'XLSX originale compilato, logo e impaginazione invariati',
      createdAt:model?.createdAt||now,
      updatedAt:now,
      updatedBy:P.currentUser?.()||'',
      syncPending:true
    };
    if(model)Object.assign(model,data);else P.state.models.unshift(data);
    P.persistLocal?.();
    P.scheduleSync?.();
    return model||data;
  }

  const parse=text=>{const doc=parser.parseFromString(text,'application/xml');if(doc.querySelector('parsererror'))throw new Error('La matrice Avola non è leggibile.');return doc};
  const child=(node,name)=>[...node.children].find(item=>item.localName===name)||null;
  const sheetPath=zip=>{
    const wb=parse(zip.file('xl/workbook.xml').asText());
    const rels=parse(zip.file('xl/_rels/workbook.xml.rels').asText());
    const sheet=wb.getElementsByTagNameNS(NS,'sheet')[0];
    const id=sheet?.getAttributeNS(RNS,'id')||sheet?.getAttribute('r:id');
    const rel=[...rels.getElementsByTagNameNS(PRNS,'Relationship')].find(item=>item.getAttribute('Id')===id);
    const target=rel?.getAttribute('Target')?.replace(/^\/+/, '');
    if(!target)throw new Error('Foglio Preventivo Avola non trovato.');
    return target.startsWith('xl/')?target:`xl/${target}`;
  };
  const sheetData=doc=>doc.getElementsByTagNameNS(NS,'sheetData')[0];
  const row=(doc,n)=>{
    let result=[...doc.getElementsByTagNameNS(NS,'row')].find(item=>Number(item.getAttribute('r'))===n);
    if(result)return result;
    result=doc.createElementNS(NS,'row');result.setAttribute('r',String(n));
    const data=sheetData(doc),next=[...data.children].find(item=>Number(item.getAttribute('r'))>n);
    data.insertBefore(result,next||null);return result;
  };
  const colNumber=letters=>[...String(letters||'').toUpperCase()].reduce((total,ch)=>total*26+ch.charCodeAt(0)-64,0);
  const point=ref=>{const match=String(ref||'').match(/^([A-Z]+)(\d+)$/i);return match?{col:colNumber(match[1]),row:Number(match[2])}:null};
  const cell=(doc,ref)=>{
    let result=[...doc.getElementsByTagNameNS(NS,'c')].find(item=>item.getAttribute('r')===ref);
    if(result)return result;
    const p=point(ref),targetRow=row(doc,p.row);result=doc.createElementNS(NS,'c');result.setAttribute('r',ref);
    const next=[...targetRow.children].find(item=>colNumber((item.getAttribute('r')||'').replace(/\d+/g,''))>p.col);
    targetRow.insertBefore(result,next||null);return result;
  };
  const clear=target=>{[...target.children].forEach(item=>item.remove());target.removeAttribute('t')};
  const setText=(doc,ref,value)=>{
    const target=cell(doc,ref),style=target.getAttribute('s');clear(target);if(style!==null)target.setAttribute('s',style);
    target.setAttribute('t','inlineStr');const is=doc.createElementNS(NS,'is'),text=doc.createElementNS(NS,'t');
    text.setAttributeNS(XNS,'xml:space','preserve');text.textContent=String(value??'');is.appendChild(text);target.appendChild(is);
  };
  const setNumber=(doc,ref,value)=>{
    const target=cell(doc,ref),style=target.getAttribute('s');clear(target);if(style!==null)target.setAttribute('s',style);
    const node=doc.createElementNS(NS,'v');node.textContent=String(Number(value)||0);target.appendChild(node);
  };
  const setFormula=(doc,ref,formula,cached=0)=>{
    const target=cell(doc,ref),style=target.getAttribute('s');clear(target);if(style!==null)target.setAttribute('s',style);
    const f=doc.createElementNS(NS,'f'),v=doc.createElementNS(NS,'v');f.textContent=formula;v.textContent=String(Number(cached)||0);target.append(f,v);
  };
  const dateIt=value=>{const raw=C(value),m=raw.match(/^(\d{4})-(\d{2})-(\d{2})/);return m?`${m[3]}/${m[2]}/${m[1]}`:raw};
  const forceRecalc=zip=>{
    zip.remove('xl/calcChain.xml');
    const entry=zip.file('xl/workbook.xml');if(!entry)return;
    const doc=parse(entry.asText());let calc=doc.getElementsByTagNameNS(NS,'calcPr')[0];
    if(!calc){calc=doc.createElementNS(NS,'calcPr');doc.documentElement.appendChild(calc)}
    calc.setAttribute('calcMode','auto');calc.setAttribute('fullCalcOnLoad','1');calc.setAttribute('forceFullCalc','1');
    zip.file('xl/workbook.xml',serializer.serializeToString(doc));
  };
  const isAvola=model=>model&&(model.id===MODEL_ID||model.systemModel===true&&N(model.name).includes('standard avola'));

  async function exportAvola(model,stored,data){
    await M.ensureScript('PizZip',M.constants.ZIP_LIB,'pizzip');
    const zip=new PizZip(stored.buffer),path=sheetPath(zip),entry=zip.file(path);
    if(!entry)throw new Error('File matrice Avola non disponibile.');
    const doc=parse(entry.asText()),lines=Array.isArray(data.lavorazioni)?data.lavorazioni:[];
    if(lines.length>12)throw new Error('La matrice STANDARD AVOLA contiene 12 righe lavorazioni. Riduci le lavorazioni oppure crea un secondo preventivo.');

    setText(doc,'H2',`Offerta n° ${C(data.numero_preventivo||data.numero_documento||data.number)} del ${dateIt(data.data_documento||data.date)}`);
    setText(doc,'H6',C(data.cliente||data.clientName));
    setText(doc,'H11',`C.A.: ${C(data.richiedente||data.requester)}`);
    setText(doc,'D17',C(data.oggetto||data.subject));

    for(let index=0;index<12;index+=1){
      const r=24+index,line=lines[index]||{};
      setText(doc,`B${r}`,C(line.codice));
      setText(doc,`D${r}`,C(line.descrizione));
      setText(doc,`J${r}`,C(line.unita_misura));
      if(lines[index]){
        setNumber(doc,`K${r}`,Number(line.quantita)||0);
        setNumber(doc,`L${r}`,Number(line.prezzo_unitario)||0);
        setFormula(doc,`N${r}`,`K${r}*L${r}`,Number(line.totale_riga)||((Number(line.quantita)||0)*(Number(line.prezzo_unitario)||0)));
      }else{
        setText(doc,`K${r}`,'');setText(doc,`L${r}`,'');setText(doc,`N${r}`,'');
      }
    }
    const end=Math.max(24,23+lines.length);
    setFormula(doc,'N37',`SUM(N24:N${end})`,Number(data.totale_imponibile)||0);
    zip.file(path,serializer.serializeToString(doc));forceRecalc(zip);
    const output=zip.generate({type:'blob',mimeType:stored.type||'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',compression:'DEFLATE'});
    M.download(output,'Matrice_Preventivo_Avola-compilato.xlsx');
  }

  const previous=M.exportSheet?.bind(M);
  if(previous)M.exportSheet=(model,stored,data,doc)=>isAvola(model)?exportAvola(model,stored,data):previous(model,stored,data,doc);

  ensureModel();
  [800,2500,6000].forEach(delay=>setTimeout(ensureModel,delay));
  window.addEventListener('focus',()=>setTimeout(ensureModel,200));
})();
