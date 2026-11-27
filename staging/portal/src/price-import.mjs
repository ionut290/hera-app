// CSV transport only; the original accounting aliases and price calculator are injected.
export function parsePriceCsv(input) {
 const text=String(input).replace(/^\uFEFF/,'');
 if(text.length>1000000)throw Error('File troppo grande (massimo 1 MB).');
 function parse(separator){
  const rows=[];let row=[],field='',quoted=false,closed=false;
  function finishField(){row.push(field);field='';closed=false;}
  function finishRow(){finishField();if(row.some(v=>v.trim()))rows.push(row);row=[];if(rows.length>501)throw Error('Massimo 500 voci per importazione.');}
  for(let i=0;i<text.length;i++){
   const c=text[i];
   if(quoted){if(c==='"'){if(text[i+1]==='"'){field+='"';i++;}else{quoted=false;closed=true;}}else field+=c;continue;}
   if(c===separator){finishField();continue;}
   if(c==='\r'||c==='\n'){if(c==='\r'&&text[i+1]==='\n')i++;finishRow();continue;}
   if(closed)throw Error('CSV non valido dopo virgolette.');
   if(c==='"'){if(field)throw Error('Virgolette CSV non valide.');quoted=true;}else field+=c;
  }
  if(quoted)throw Error('Virgolette CSV non chiuse.');
  if(field||row.length||closed)finishRow();
  if(rows.length<2||rows[0].length<2)throw Error('CSV senza intestazioni o righe.');
  const headers=rows.shift().map(h=>h.trim());
  if(new Set(headers).size!==headers.length||headers.some(h=>!h))throw Error('Intestazioni CSV non valide.');
  if(rows.some(r=>r.length!==headers.length))throw Error('Numero colonne CSV non coerente.');
  return {columns:headers.length,rows:rows.map(row=>Object.fromEntries(headers.map((h,i)=>[h,row[i]])))};
 }
 const results=[];for(const separator of [';',','])try{results.push(parse(separator));}catch{}
 if(!results.length)throw Error('CSV non valido o limite superato.');
 results.sort((a,b)=>b.columns-a.columns);
 return results[0].rows;
}
export function planPriceImport(rows,existing,helpers){
 const {mapped,num,norm,discounted}=helpers;
 if(!Array.isArray(rows)||rows.length>500)throw Error('Massimo 500 voci per importazione.');
 const existingByCode=new Map();
 for(const item of existing){const code=norm(item.codiceVoce);if(!code)continue;if(existingByCode.has(code))throw Error('Prezziario esistente con codici duplicati.');existingByCode.set(code,item);}
 const seen=new Set(),plan=[];let next=Math.max(0,...existing.map(p=>num(p.numeroProgressivo)||0));
 const read=(row,k)=>String(mapped(row,k)??'').trim();
 for(const [index,row] of rows.entries()){
  if(Object.values(row).some(v=>norm(v).includes('esempio - non importare')))continue;
  const codiceVoce=read(row,'codiceVocePrezzo').toUpperCase(),code=norm(codiceVoce);
  const descrizione=read(row,'tipologiaLavorazione'),unitaMisura=read(row,'unitaMisura'),note=read(row,'note');
  const base=read(row,'prezzoBase'),raw=read(row,'percentualeRibasso');
  const prezzoBase=base?num(base):null;
  const numeric=raw?num(raw.endsWith('%')?raw.slice(0,-1).trim():raw):null;
  const percentualeRibasso=numeric===null?null:raw.endsWith('%')||numeric>1?numeric/100:numeric;
  if(!code||seen.has(code)||!unitaMisura||prezzoBase===null||!Number.isFinite(prezzoBase)||prezzoBase<0||prezzoBase>100000000||raw&&numeric===null||percentualeRibasso!==null&&(!Number.isFinite(percentualeRibasso)||percentualeRibasso<0||percentualeRibasso>1)||codiceVoce.length>150||descrizione.length>2000||unitaMisura.length>100||note.length>3000)throw Error('Riga '+(index+2)+': codice duplicato o dati non validi.');
  seen.add(code);
  const previous=existingByCode.get(code);
  const data={numeroProgressivo:previous?previous.numeroProgressivo:++next,codiceVoce,descrizione,unitaMisura,prezzoBase,percentualeRibasso,note};
  data.prezzoRibassato=discounted(data);
  plan.push({id:previous?.id??null,data});
 }
 if(!plan.length)throw Error('Il file non contiene righe importabili.');
 return plan;
}
