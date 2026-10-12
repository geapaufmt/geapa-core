import {exact,need,input,revisions,hex} from './command-contract.mjs';
import {id} from './errors.mjs';
// SDK batch for isolated demo only. Existing Rules intentionally keep this denied
// until the separately owned D admission/request-use gates are implemented.
export function createRequestUsePort(db,sdk,auth,config){
 exact(config,['environment','projectId']);need(config.environment==='SYNTHETIC_LOCAL'&&config.projectId==='demo-geapa-m13-local');
 const target=db?._delegate||db;
 need(target?.app?.options?.projectId===config.projectId&&target?._databaseId?.database==='(default)'&&target?._settings?.host==='127.0.0.1:8184'&&target?._settings?.ssl===false);
 return async pair=>{
  exact(pair,['request','use']);const r=pair.request,u=pair.use,t=auth.capture();
  exact(r,['schemaVersion','operationId','ownerUid','commandType','target','expectedRevisions','payload','evidenceRefs','decisionRef','digest','ticketId','state']);
  input(Object.fromEntries(['schemaVersion','operationId','commandType','target','payload','evidenceRefs','decisionRef'].map(k=>[k,r[k]])));revisions(r.expectedRevisions);
  need(r.ownerUid===t.uid&&r.state==='RECEIVED'&&id(r.ticketId)&&hex(r.digest));
  exact(u,['schemaVersion','ticketId','operationId','ownerUid']);need(u.schemaVersion==='member-admission-use-v13'&&u.ticketId===r.ticketId&&u.operationId===r.operationId&&u.ownerUid===t.uid);
  const batch=sdk.writeBatch(db);batch.set(sdk.doc(db,'commandRequests/'+r.operationId),{...r,receivedAt:sdk.serverTimestamp()});batch.set(sdk.doc(db,'memberAdmissionUses/'+r.ticketId),{...u,reservedAt:sdk.serverTimestamp()});auth.assert(t);await batch.commit();auth.assert(t);
 };
}
