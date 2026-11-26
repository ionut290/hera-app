'use strict';
// Local-only transport for synthetic emulator tests. Never export from live Functions.
const {createServer}=require('node:http');
const ACTIONS=new Set(['listCompanies','createCompany','setCompanyStatus','listMyCompanies','openCompany','listCompanyCommesse']);
function localBoundary(){
  if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||process.env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9199')throw Error('emulator-only');
}
async function startLocalCompanyHttp(dispatch){
  localBoundary();
  let origin;
  const server=createServer(async(req,res)=>{
    const send=(status,body)=>{res.writeHead(status,{'content-type':'application/json','cache-control':'no-store','x-content-type-options':'nosniff'});res.end(JSON.stringify(body));};
    try{
      localBoundary();
      if(req.method!=='POST'||req.url!=='/api/v2/companies')return send(404,{ok:false,code:'not-found'});
      if(req.headers.host!==new URL(origin).host||req.headers.origin!==origin)return send(403,{ok:false,code:'forbidden'});
      if(req.headers['content-type']!=='application/json')return send(415,{ok:false,code:'invalid-payload'});
      const token=/^Bearer ([^\s]{1,16000})$/.exec(req.headers.authorization||'')?.[1];
      if(!token)return send(401,{ok:false,code:'unauthenticated'});
      let size=0;const chunks=[];
      for await(const chunk of req){size+=chunk.length;if(size>8192)return send(413,{ok:false,code:'invalid-payload'});chunks.push(chunk);}
      let body;try{body=JSON.parse(Buffer.concat(chunks).toString('utf8'));}catch{return send(400,{ok:false,code:'invalid-payload'});}
      if(!body||!ACTIONS.has(body.action)||!body.payload||typeof body.payload!=='object'||Array.isArray(body.payload))return send(400,{ok:false,code:'invalid-payload'});
      const result=await dispatch(token,body.action,body.payload);
      return send(200,{ok:true,result});
    }catch(error){
      const code=error?.code;
      if(code==='unauthenticated')return send(401,{ok:false,code});
      if(code==='forbidden')return send(403,{ok:false,code});
      if(typeof code==='string'&&/^(invalid-[a-z-]+|not-found|unknown-action|already-exists)$/.test(code))return send(400,{ok:false,code});
      return send(500,{ok:false,code:'request-failed'});
    }
  });
  server.requestTimeout=15000;
  await new Promise((resolve,reject)=>{server.once('error',reject);server.listen(0,'127.0.0.1',resolve);});
  origin='http://127.0.0.1:'+server.address().port;
  return {origin,close:()=>new Promise((resolve,reject)=>server.close(e=>e?reject(e):resolve()))};
}
module.exports={startLocalCompanyHttp};
