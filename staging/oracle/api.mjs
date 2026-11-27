import http from 'node:http';
import {randomUUID,randomBytes} from 'node:crypto';
class ApiError extends Error {constructor(status,code){super(code);this.status=status;}}
const fail=(status,code)=>{throw new ApiError(status,code)};
const uuid=v=>{if(typeof v!=='string'||!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(v))fail(400,'INVALID_ID');return v.toLowerCase()};
const name=(v,min=1,max=160)=>{if(typeof v!=='string'||v.trim().length<min||v.trim().length>max)fail(400,'INVALID_NAME');return v.trim()};
const code=v=>{if(typeof v!=='string'||!/^VC-[A-Z0-9]{8}$/.test(v.trim().toUpperCase()))fail(400,'INVALID_CODE');return v.trim().toUpperCase()};
async function body(req){let size=0;const chunks=[];for await(const chunk of req){size+=chunk.length;if(size>16384)fail(413,'BODY_TOO_LARGE');chunks.push(chunk)}try{const data=JSON.parse(Buffer.concat(chunks).toString());if(!data||Array.isArray(data)||typeof data!=='object')fail(400,'INVALID_JSON');return data}catch{fail(400,'INVALID_JSON')}}
async function transaction(pool,work){const db=await pool.connect();try{await db.query('BEGIN');const result=await work(db);await db.query('COMMIT');return result}catch(e){try{await db.query('ROLLBACK')}catch{}throw e}finally{db.release()}}
async function scope(db,tenant,uid,admin=false,lock=false){
 const result=await db.query('SELECT t.*,m.role FROM tenants t JOIN members m ON m.tenant_id=t.id WHERE t.id=$1 AND m.uid=$2 AND m.status=\'active\' AND t.status=\'active\''+(lock?' FOR UPDATE OF t':''),[tenant,uid]);
 const row=result.rows[0];if(!row)fail(404,'COMPANY_NOT_FOUND');if(admin&&row.role!=='admin')fail(403,'ADMIN_REQUIRED');
 await db.query("SELECT set_config('varga.tenant_id',$1,true)",[tenant]);return row;
}
const count=async(db,sql,values)=>Number((await db.query(sql,values)).rows[0].n);
function coordinate(v,min,max){if(v===null||v===undefined)return null;if(typeof v!=='number'||!Number.isFinite(v)||v<min||v>max)fail(400,'INVALID_COORDINATES');return v}
function publicTenant(t){return {id:t.id,name:t.name,status:t.status,plan:t.plan,role:t.role,limits:{commesse:t.max_jobs,impianti:t.max_plants,admin:t.max_admins,operatori:t.max_operators},...(t.role==='admin'?{joinCode:t.join_code}:{})}}
export function createApi({pool,verifyToken}){
 if(!pool||typeof verifyToken!=='function')throw Error('Required database and verifier');
 return http.createServer(async(req,res)=>{
  res.setHeader('Content-Type','application/json; charset=utf-8');res.setHeader('Cache-Control','no-store');res.setHeader('X-Content-Type-Options','nosniff');
  const reply=(status,data)=>{res.statusCode=status;res.end(JSON.stringify(data))};
  try{
   const url=new URL(req.url,'http://localhost'),p=url.pathname.split('/').filter(Boolean),method=req.method;
   if(method==='GET'&&url.pathname==='/health')return reply(200,{ok:true});
   const match=/^Bearer ([^\s]+)$/.exec(req.headers.authorization||'');if(!match)fail(401,'LOGIN_REQUIRED');
   let user;try{user=await verifyToken(match[1])}catch{fail(401,'INVALID_TOKEN')}
   if(!user?.uid||typeof user.uid!=='string'||user.uid.length>128||user.email_verified!==true||typeof user.email!=='string')fail(403,'VERIFIED_EMAIL_REQUIRED');
   const uid=user.uid;
   if(p[0]!=='v1')fail(404,'NOT_FOUND');
   if(p.length===2&&p[1]==='companies'&&method==='GET'){
    const q=await pool.query("SELECT t.*,m.role FROM tenants t JOIN members m ON t.id=m.tenant_id WHERE m.uid=$1 AND m.status='active' AND t.status='active' ORDER BY t.created_at",[uid]);return reply(200,q.rows.map(publicTenant));
   }
   if(p.length===2&&p[1]==='companies'&&method==='POST'){
    const b=await body(req),label=name(b.name,2,120),requestId=uuid(b.requestId);
    const out=await transaction(pool,async db=>{
     const id=randomUUID(),joinCode='VC-'+randomBytes(4).toString('hex').toUpperCase();
     const q=await db.query('INSERT INTO tenants(id,owner_uid,request_id,name,join_code) VALUES($1,$2,$3,$4,$5) ON CONFLICT(owner_uid,request_id) DO NOTHING RETURNING *',[id,uid,requestId,label,joinCode]);
     let t=q.rows[0];if(t)await db.query("INSERT INTO members(tenant_id,uid,role) VALUES($1,$2,'admin')",[id,uid]);
     else {t=(await db.query('SELECT * FROM tenants WHERE owner_uid=$1 AND request_id=$2',[uid,requestId])).rows[0];if(!t||t.name!==label)fail(409,'REQUEST_CONFLICT')}
     return publicTenant({...t,role:'admin'});
    });return reply(200,out);
   }
   if(p.length===2&&p[1]==='join'&&method==='POST'){
    const b=await body(req),joinCode=code(b.code);
    const out=await transaction(pool,async db=>{
     const t=(await db.query("SELECT id FROM tenants WHERE join_code=$1 AND status='active' FOR UPDATE",[joinCode])).rows[0];if(!t)fail(404,'CODE_NOT_FOUND');
     const m=(await db.query("SELECT status FROM members WHERE tenant_id=$1 AND uid=$2",[t.id,uid])).rows[0];if(m) return {status:m.status==='active'?'accepted':'rejected'};
     await db.query("INSERT INTO join_requests(tenant_id,uid,email,display_name) VALUES($1,$2,$3,$4) ON CONFLICT DO NOTHING",[t.id,uid,user.email.slice(0,320),String(user.name||'Operatore').slice(0,120)]);
     return {status:(await db.query('SELECT status FROM join_requests WHERE tenant_id=$1 AND uid=$2',[t.id,uid])).rows[0].status};
    });return reply(200,out);
   }
   if(p[1]!=='companies'||p.length<4)fail(404,'NOT_FOUND');
   const tenant=uuid(p[2]),resource=p[3];
   const input=method==='POST'?await body(req):null;
   const out=await transaction(pool,async db=>{
    const t=await scope(db,tenant,uid,resource==='requests'||method==='POST',method==='POST');
    if(resource==='requests'&&p.length===4&&method==='GET')return (await db.query("SELECT uid,email,display_name,status,requested_at FROM join_requests WHERE tenant_id=$1 AND status='pending' ORDER BY requested_at",[tenant])).rows;
    if(resource==='requests'&&p.length===5&&method==='POST'){
     const target=decodeURIComponent(p[4]);if(!target||target.length>128||!['accepted','rejected'].includes(input.decision))fail(400,'INVALID_DECISION');
     const q=(await db.query('SELECT status FROM join_requests WHERE tenant_id=$1 AND uid=$2',[tenant,target])).rows[0];if(!q)fail(404,'REQUEST_NOT_FOUND');if(q.status!=='pending')fail(409,'ALREADY_REVIEWED');
     if(input.decision==='accepted'){
      if(await count(db,"SELECT count(*) n FROM members WHERE tenant_id=$1 AND status='active' AND role='employee'",[tenant])>=t.max_operators)fail(409,'OPERATOR_LIMIT');
      await db.query("INSERT INTO members(tenant_id,uid,role) VALUES($1,$2,'employee')",[tenant,target]);
     }
     await db.query('UPDATE join_requests SET status=$3,reviewed_at=now(),reviewed_by=$4 WHERE tenant_id=$1 AND uid=$2',[tenant,target,input.decision,uid]);return {status:input.decision};
    }
    if(resource==='jobs'&&p.length===4&&method==='GET')return (await db.query('SELECT id,name,created_at FROM jobs WHERE tenant_id=$1 ORDER BY created_at',[tenant])).rows;
    if(resource==='jobs'&&p.length===4&&method==='POST'){
     if(await count(db,'SELECT count(*) n FROM jobs WHERE tenant_id=$1',[tenant])>=t.max_jobs)fail(409,'JOB_LIMIT');
     return (await db.query('INSERT INTO jobs(tenant_id,id,name,created_by) VALUES($1,$2,$3,$4) RETURNING id,name,created_at',[tenant,randomUUID(),name(input.name),uid])).rows[0];
    }
    if(resource==='jobs'&&p.length===6&&p[5]==='plants'){
     const job=uuid(p[4]);if(!(await db.query('SELECT id FROM jobs WHERE tenant_id=$1 AND id=$2',[tenant,job])).rowCount)fail(404,'JOB_NOT_FOUND');
     if(method==='GET')return (await db.query('SELECT id,job_id,name,lat,lng,created_at FROM plants WHERE tenant_id=$1 AND job_id=$2 ORDER BY created_at',[tenant,job])).rows;
     if(method==='POST'){
      if(await count(db,'SELECT count(*) n FROM plants WHERE tenant_id=$1',[tenant])>=t.max_plants)fail(409,'PLANT_LIMIT');
      const lat=coordinate(input.lat,-90,90),lng=coordinate(input.lng,-180,180);if((lat===null)!==(lng===null))fail(400,'INVALID_COORDINATES');
      return (await db.query('INSERT INTO plants(tenant_id,id,job_id,name,lat,lng,created_by) VALUES($1,$2,$3,$4,$5,$6,$7) RETURNING id,job_id,name,lat,lng',[tenant,randomUUID(),job,name(input.name),lat,lng,uid])).rows[0];
     }
    }
    fail(404,'NOT_FOUND');
   });return reply(200,out);
  }catch(e){if(!res.headersSent)reply(e instanceof ApiError?e.status:500,{error:e instanceof ApiError?e.message:'INTERNAL_ERROR'});else res.end()}
 });
}
