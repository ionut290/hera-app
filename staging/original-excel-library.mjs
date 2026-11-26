// Same ExcelJS version as the original app. Build-time only; no CDN in the preview.
import {createHash} from 'node:crypto';
import {gunzipSync} from 'node:zlib';
const URL='https://registry.npmjs.org/exceljs/-/exceljs-4.4.0.tgz';
const INTEGRITY='sha512-XctvKaEMaj1Ii9oDOqbW/6e1gXknSY4g/aLCDicOXqBE4M0nRWkUu0PTp++UPNzoFY12BNHMfs/VadKIS6llvg==';
export async function loadOriginalExcelLibrary(){
 const response=await fetch(URL,{signal:AbortSignal.timeout(45000),redirect:'error'});
 if(!response.ok)throw Error('Cannot download pinned ExcelJS: '+response.status);
 const chunks=[];let size=0;
 for await(const chunk of response.body){size+=chunk.length;if(size>16000000)throw Error('ExcelJS package too large');chunks.push(Buffer.from(chunk));}
 const bytes=Buffer.concat(chunks);
 if('sha512-'+createHash('sha512').update(bytes).digest('base64')!==INTEGRITY)throw Error('ExcelJS integrity mismatch');
 const tar=gunzipSync(bytes,{maxOutputLength:32000000});
 const files=new Map();
 for(let at=0;at+512<=tar.length;){
  const name=tar.toString('utf8',at,at+100).split('\0')[0];if(!name)break;
  const n=parseInt(tar.toString('ascii',at+124,at+136).trim().replace(/\0/g,''),8);
  if(!Number.isSafeInteger(n)||n<0||at+512+n>tar.length)throw Error('Invalid package archive');
  if(name==='package/dist/exceljs.min.js'||name==='package/LICENSE')files.set(name,tar.toString('utf8',at+512,at+512+n));
  at+=512+Math.ceil(n/512)*512;
 }
 const code=files.get('package/dist/exceljs.min.js'),license=files.get('package/LICENSE');
 if(!code||!license)throw Error('ExcelJS distribution or license missing');
 return '/* ExcelJS 4.4.0 - '+license.replace(/\*\//g,'* /')+' */\n'+code.replace(/\/\/# sourceMappingURL=.*$/m,'');
}
