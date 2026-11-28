// Local emulator HTTP adapter only; must never be mounted on a live web server.
import {MAX_FILE_BYTES} from './storage-quota-service.mjs';
const cap=Math.ceil(MAX_FILE_BYTES/3)*4+4096;
const allowed=new Set(['list','upload','download','remove']);
export function createLocalStorageHandler(service){
 if(process.env.FIREBASE_STORAGE_EMULATOR_HOST!=='127.0.0.1:9199')throw Error('EMULATORS_REQUIRED');
 return async(req,res)=>{
  const reply=(status,data)=>{res.writeHead(status,{'Content-Type':'application/json','Cache-Control':'no-store','X-Content-Type-Options':'nosniff'});res.end(JSON.stringify(data));};
  if(req.method!=='POST'||req.headers.host!=='127.0.0.1:4173'||req.headers.origin!=='http://127.0.0.1:4173'||!String(req.headers['content-type']).startsWith('application/json'))return reply(403,{code:'FORBIDDEN'});
  const action=req.url.split('/').at(-1);
  if(!allowed.has(action))return reply(404,{code:'NOT_FOUND'});
  const token=String(req.headers.authorization||'').match(/^Bearer ([A-Za-z0-9._-]+)$/)?.[1];
  if(!token)return reply(401,{code:'UNAUTHENTICATED'});
  try{
   let size=0;const chunks=[];
   for await(const chunk of req){size+=chunk.length;if(size>cap){reply(413,{code:'INVALID_FILE'});return;}chunks.push(chunk);}
   const data=JSON.parse(Buffer.concat(chunks).toString('utf8'));
   if(!data||typeof data!=='object'||Array.isArray(data))return reply(400,{code:'INVALID_REQUEST'});
   const result=await service[action]({...data,token});
   return reply(200,result);
  }catch(e){const code=typeof e?.code==='string'&&/^[A-Z_]{3,48}$/.test(e.code)?e.code:'REQUEST_FAILED';return reply(code==='FORBIDDEN'?403:code==='UNAUTHENTICATED'?401:400,{code});}
 };
}
