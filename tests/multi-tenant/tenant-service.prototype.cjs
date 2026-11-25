'use strict';
// Backend prototype ONLY. Not exported by production Functions or called by the app.
// The future HTTP/callable adapter MUST supply verified Auth claims, never client claims.
// No production endpoint, credentials, database or emails are used.
const { randomBytes, randomUUID, createHash } = require('node:crypto');
const roles = new Set(['admin','manager','team_lead','worker','client']);
const adminRoles = new Set(['owner','admin']);
const DAY = 86400000;
function fail(code) { const err = new Error(code); err.code = code; throw err; }
function id(value) {
  if (typeof value !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(value) || /^__.*__$/.test(value)) fail('invalid-id');
  return value;
}
function email(value) {
  if (typeof value !== 'string') fail('invalid-email');
  const normalized = value.trim().toLowerCase();
  if (normalized.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized)) fail('invalid-email');
  return normalized;
}
function actor(auth) {
  if (!auth || auth.email_verified !== true || auth.apple_review === true || auth.disabled === true) fail('unauthenticated');
  return { uid:id(auth.uid), email:email(auth.email), platformAdmin:auth.platform_admin === true };
}
function digest(token) { return createHash('sha256').update(token).digest('hex'); }
function tenantPath(tenantId) { return 'tenants/'+id(tenantId); }
function memberPath(tenantId,uid) { return tenantPath(tenantId)+'/members/'+id(uid); }
function indexPath(tenantId,uid) { return 'users/'+id(uid)+'/memberships/'+id(tenantId); }
// store.runTransaction(fn): tx.get(path) -> data/null, tx.create(path,data), tx.set(path,data).
// All reads precede writes; implementation must be atomic, retry-safe and fail on create conflicts.
function makeTenantService(store, clock=Date.now) {
  async function createCompany(auth, {name,ownerUid,ownerEmail}) {
    const caller = actor(auth);
    if (!caller.platformAdmin) fail('forbidden');
    if (typeof name !== 'string' || !name.trim() || name.trim().length > 100) fail('invalid-name');
    // ownerUid/email MUST be resolved server-side from Firebase Auth by future adapter.
    const uid=id(ownerUid), normalized=email(ownerEmail), tenantId=randomUUID(), now=clock();
    const root=tenantPath(tenantId), member={role:'owner',status:'active',email:normalized,createdAt:now};
    await store.runTransaction(async tx=>{
      tx.create(root,{name:name.trim(),status:'active',inviteGeneration:0,createdAt:now,createdBy:caller.uid});
      tx.create(memberPath(tenantId,uid),member);
      tx.create(indexPath(tenantId,uid),{role:'owner',status:'active'});
      tx.create(root+'/audit/'+randomUUID(),{event:'company-created',actorUid:caller.uid,subjectUid:uid,at:now});
    });
    return {tenantId};
  }
  async function invite(auth, {tenantId,recipientEmail,role}) {
    const caller=actor(auth), root=tenantPath(tenantId), target=email(recipientEmail);
    if (!roles.has(role)) fail('invalid-role');
    const token=randomBytes(32).toString('hex'), inviteId=digest(token), now=clock();
    await store.runTransaction(async tx=>{
      const tenant=await tx.get(root), membership=await tx.get(memberPath(tenantId,caller.uid));
      if (tenant?.status!=='active' || membership?.status!=='active' || !adminRoles.has(membership.role)) fail('forbidden');
      // An admin may invite operational roles, not another admin.
      if (role==='admin' && membership.role!=='owner') fail('forbidden');
      tx.create(root+'/invites/'+inviteId,{email:target,role,status:'pending',generation:tenant.inviteGeneration??0,createdBy:caller.uid,createdAt:now,expiresAt:now+7*DAY});
      tx.create(root+'/audit/'+randomUUID(),{event:'invite-created',actorUid:caller.uid,role,at:now});
    });
    // Only the caller receives the raw token; never store or log it.
    return {tenantId,token,expiresAt:now+7*DAY};
  }
  async function accept(auth,{tenantId,token}) {
    const caller=actor(auth),root=tenantPath(tenantId);
    if (typeof token!=='string'||!/^[a-f0-9]{64}$/.test(token)) fail('invalid-invite');
    const path=root+'/invites/'+digest(token);
    await store.runTransaction(async tx=>{
      const tenant=await tx.get(root), invitation=await tx.get(path), member=await tx.get(memberPath(tenantId,caller.uid)), now=clock();
      if (tenant?.status!=='active'|| invitation?.status!=='pending'|| (invitation.generation??0)!==(tenant.inviteGeneration??0)|| !Number.isFinite(invitation.expiresAt)|| invitation.expiresAt<=now|| invitation.email!==caller.email||!roles.has(invitation.role)) fail('invalid-invite');
      // Recheck issuer membership; revoked issuers cannot confer access.
      const issuer=await tx.get(memberPath(tenantId,invitation.createdBy));
      if (issuer?.status!=='active'|| !adminRoles.has(issuer.role)|| (invitation.role==='admin'&&issuer.role!=='owner')) fail('invalid-invite');
      // Never overwrite active or historical/revoked memberships. Reactivation requires separate review.
      if (member) fail('membership-exists');
      tx.create(memberPath(tenantId,caller.uid),{role:invitation.role,status:'active',email:caller.email,createdAt:now});
      tx.create(indexPath(tenantId,caller.uid),{role:invitation.role,status:'active'});
      tx.set(path,{...invitation,status:'used',usedBy:caller.uid,usedAt:now});
      tx.create(root+'/audit/'+randomUUID(),{event:'invite-accepted',actorUid:caller.uid,at:now});
    });
    return {tenantId};
  }
  return Object.freeze({createCompany,invite,accept});
}
module.exports={makeTenantService};
