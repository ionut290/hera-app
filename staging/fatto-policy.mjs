// Pure staging policy; no operational effects or production integration.
const actions=Object.freeze(['saveCompletion','moveToCompleted','openWhatsApp','exportSheet']);
export function validateFattoPolicy({tenant,context,policy}){
 if(!tenant || tenant.demo!==true || tenant.status!=='active' || tenant.id!==context?.tenantId || !context.uid || context.role!=='admin') throw new Error('Access denied');
 // Only explicitly provisioned configurable demo tenants are eligible.
 if(tenant.fattoMode!=='configurable-demo') throw new Error('Protected or unknown workflow');
 if(!policy || !Number.isSafeInteger(policy.version) || policy.version<1 || !Array.isArray(policy.actions)) throw new Error('Invalid policy');
 const steps=[...policy.actions];
 if(steps[0]!=='saveCompletion' || new Set(steps).size!==steps.length || steps.some(x=>!actions.includes(x))) throw new Error('Invalid action sequence');
 if(steps.includes('exportSheet') && tenant.sheetConfigured!==true) throw new Error('Sheet integration not configured');
 if(steps.includes('openWhatsApp') && tenant.whatsappConfigured!==true) throw new Error('WhatsApp not configured');
 return Object.freeze({tenantId:tenant.id,version:policy.version,actions:Object.freeze(steps)});
}
