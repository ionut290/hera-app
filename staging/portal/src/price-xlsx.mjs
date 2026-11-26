// Bounded local .xlsx reader. No network and no formula evaluation.
export function inspectXlsxZip(buffer){
 const d=new DataView(buffer),length=d.byteLength;
 if(length<22||length>5000000)throw Error('File Excel non valido o troppo grande (max 5 MB).');
 let end=-1;
 for(let i=length-22;i>=Math.max(0,length-65557);i--){if(d.getUint32(i,true)===0x06054b50&&i+22+d.getUint16(i+20,true)===length){end=i;break;}}
 if(end<0||d.getUint16(end+4,true)||d.getUint16(end+6,true))throw Error('Archivio Excel non valido.');
 const count=d.getUint16(end+10,true),bytes=d.getUint32(end+12,true),start=d.getUint32(end+16,true);
 if(count!==d.getUint16(end+8,true)||count<1||count>600||start+bytes!==end)throw Error('Archivio Excel non supportato.');
 let at=start,total=0;const names=new Set(),entries=[];
 for(let i=0;i<count;i++){
  if(at+46>end||d.getUint32(at,true)!==0x02014b50)throw Error('Archivio Excel danneggiato.');
  const flags=d.getUint16(at+8,true),method=d.getUint16(at+10,true),packed=d.getUint32(at+20,true),size=d.getUint32(at+24,true),n=d.getUint16(at+28,true),extra=d.getUint16(at+30,true),comment=d.getUint16(at+32,true),local=d.getUint32(at+42,true);
  if(flags&1||![0,8].includes(method)||at+46+n+extra+comment>end||local+30>start||d.getUint32(local,true)!==0x04034b50)throw Error('Archivio Excel non supportato.');
  const localNameLength=d.getUint16(local+26,true),dataAt=local+30+localNameLength+d.getUint16(local+28,true);
  if(d.getUint16(local+6,true)!==flags||d.getUint16(local+8,true)!==method||localNameLength!==n||dataAt>start)throw Error('Intestazioni ZIP incoerenti.');
  for(let k=0;k<n;k++)if(d.getUint8(local+30+k)!==d.getUint8(at+46+k))throw Error('Nome ZIP incoerente.');
  if(!(flags&8)&&(d.getUint32(local+18,true)!==packed||d.getUint32(local+22,true)!==size))throw Error('Dimensioni ZIP incoerenti.');
  if(dataAt+packed>start||size>8000000||(total+=size)>20000000)throw Error('Archivio Excel troppo grande.');
  const name=new TextDecoder().decode(new Uint8Array(buffer,at+46,n));
  if(names.has(name)||name.includes('\\')||name.startsWith('/')||name.split('/').includes('..')||/vbaProject|externalLinks/i.test(name))throw Error('Archivio Excel non supportato.');
  names.add(name);entries.push({local,dataAt,packed,size,method});at+=46+n+extra+comment;
 }
 if(at!==end||!names.has('[Content_Types].xml')||!names.has('xl/workbook.xml'))throw Error('File XLSX non valido.');
 entries.sort((a,b)=>a.local-b.local);
 for(let i=1;i<entries.length;i++)if(entries[i-1].dataAt+entries[i-1].packed>entries[i].local)throw Error('Dati ZIP sovrapposti.');
 return entries;
}
// Verify actual inflated length, not only sizes claimed in ZIP metadata.
export async function verifyXlsxInflation(buffer,entries){
 let total=0;
 for(const {dataAt,packed,size,method} of entries){
  if(method===0){if(packed!==size)throw Error('Dimensioni ZIP incoerenti.');total+=size;continue;}
  if(typeof DecompressionStream!=='function')throw Error('Importazione Excel non supportata da questo browser.');
  const reader=new Blob([new Uint8Array(buffer,dataAt,packed)]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();
  let actual=0;
  try{
   for(;;){
    const {done,value}=await reader.read();if(done)break;
    actual+=value.byteLength;
    if(actual>size||actual>8000000||total+actual>20000000)throw Error('Archivio Excel oltre il limite di decompressione.');
   }
   if(actual!==size)throw Error('Dimensioni ZIP incoerenti.');
  }finally{try{await reader.cancel();}catch{}reader.releaseLock();}
  total+=actual;
 }
}

export async function readPriceXlsx(buffer,ExcelJS,aliases,key){
 const entries=inspectXlsxZip(buffer);await verifyXlsxInflation(buffer,entries);
 const wb=new ExcelJS.Workbook();await wb.xlsx.load(buffer);
 const ws=wb.getWorksheet('Elenco prezzi')||wb.worksheets[0];
 if(!ws||ws.rowCount>501||ws.columnCount>64)throw Error('Massimo 500 voci e 64 colonne.');
 const allowed=new Set(['numeroProgressivoRiga','tipologiaLavorazione','codiceVocePrezzo','unitaMisura','prezzoBase','percentualeRibasso','note'].flatMap(f=>aliases[f].map(key)));
 const headers=[],seen=new Set();
 const scalar=cell=>{
  const v=cell.value;
  if(v==null)return '';
  if(typeof v==='string'||typeof v==='number')return v;
  if(v.richText)return v.richText.map(t=>t.text||'').join('');
  throw Error('Le celle del prezziario devono contenere valori, non formule o collegamenti.');
 };
 for(let col=1;col<=ws.columnCount;col++){
  const value=ws.getRow(1).getCell(col).value;
  if(typeof value!=='string')continue;
  const k=key(value);if(!allowed.has(k))continue;
  if(seen.has(k))throw Error('Intestazione Excel duplicata.');
  seen.add(k);headers.push([col,value]);
 }
 if(!headers.length)throw Error('Intestazioni del prezziario non trovate.');
 const rows=[];
 for(let r=2;r<=ws.rowCount;r++){
  const row=Object.fromEntries(headers.map(([col,name])=>[name,scalar(ws.getRow(r).getCell(col))]));
  if(Object.values(row).some(v=>String(v).trim()))rows.push(row);
 }
 return rows;
}
