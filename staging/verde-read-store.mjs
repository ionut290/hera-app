// Read-only compat adapter for the isolated Verde Levato pilot.
import { tenantCollectionPath } from './tenant-context.mjs';
const allowed = new Set(['verdeLevatoRecords','verdeLevatoCommesse','verdeLevatoConfig']);
export function createVerdeReadStore(db, context) {
  if (!db || db.app?.options?.projectId !== 'varga-cantieri-v2-staging') throw new Error('Staging Firebase required');
  const [root, tenant] = tenantCollectionPath(context, 'sites');
  return Object.freeze({
    collection(name) {
      if (!allowed.has(name)) throw new Error('Collection not approved');
      const ref = db.collection(root).doc(tenant).collection(name);
      return Object.freeze({
        get: () => ref.get(),
        doc(id) {
          if (typeof id !== 'string' || !id || id.includes('/')) throw new Error('Document ID required');
          return Object.freeze({get: () => ref.doc(id).get()});
        }
      });
    }
  });
}
