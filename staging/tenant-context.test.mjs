import test from "node:test";
import assert from "node:assert/strict";
import { resolveTenantContext, tenantCollectionPath } from "./tenant-context.mjs";
const base = { uid:"demo-user", claims:{email_verified:true}, tenantId:"demo-001", tenant:{demo:true,status:"active"}, member:{status:"active",role:"admin"}, membership:{tenantId:"demo-001",status:"active",role:"admin"} };
for (const role of ["admin","supervisor","employee"]) {
  test("accepts active "+role,()=>assert.equal(resolveTenantContext({...base,member:{...base.member,role},membership:{...base.membership,role}}).role,role));
}
for (const [name, patch] of Object.entries({
  anonymous:{uid:null}, unverified:{claims:{email_verified:false}}, appleReview:{claims:{email_verified:true,apple_review:true}},
  realTenant:{tenant:{demo:false,status:"active"}}, suspendedTenant:{tenant:{demo:true,status:"suspended"}},
  suspendedMember:{member:{...base.member,status:"suspended"}}, suspendedMembership:{membership:{...base.membership,status:"suspended"}},
  wrongTenant:{membership:{...base.membership,tenantId:"demo-002"}}, roleMismatch:{member:{...base.member,role:"employee"}},
  invalidRole:{member:{...base.member,role:"owner"},membership:{...base.membership,role:"owner"}}, invalidId:{tenantId:"../production"}
})) test("rejects "+name,()=>assert.equal(resolveTenantContext({...base,...patch}),null));
test("tenant path",()=>assert.deepEqual(tenantCollectionPath(resolveTenantContext(base),"sites"),["tenants","demo-001","sites"]));
test("rejects unapproved collection",()=>assert.throws(()=>tenantCollectionPath(resolveTenantContext(base),"impianti")));
test("rejects missing context",()=>assert.throws(()=>tenantCollectionPath(null,"sites")));
test("context is immutable",()=>assert.ok(Object.isFrozen(resolveTenantContext(base))));
