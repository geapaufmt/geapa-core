'use strict';
// Bounded historical reconciliation: index existing approved synthetic records.
// Never creates approvals, identities, grants, enrollments or invitations.
const a=require('node:assert/strict'),crypto=require('node:crypto');
const PROJECT='geapa-convergencia-dev',hash=v=>crypto.createHash('sha256').update(JSON.stringify(v)).digest('hex');
async function indexApproval({db,uid,approvalIds,expectedGrantRevision,expectedIndexRevision,expectedApprovalRevisions,operator,apply=false,clock=Date.now}){
 a.equal(db.projectId,PROJECT,'DEV_ONLY');a.match(uid,/^[A-Za-z0-9_-]{1,48}$/);a(typeof operator==='string'&&operator.length>0);a(Array.isArray(approvalIds)&&approvalIds.length>0&&approvalIds.length<=16&&new Set(approvalIds).size===approvalIds.length);a(approvalIds.every(id=>/^[A-Za-z0-9_-]{1,48}$/.test(id)));a.equal(expectedApprovalRevisions.length,approvalIds.length);
 return db.runTransaction(async tx=>{
  const get=async p=>{const s=await tx.get(db.doc(p));return s.exists?s.data():null;},now=clock(),grant=await get('coreDevAdministrativeGrants/'+uid);
  a(grant?.active&&grant.uid===uid&&grant.projectId===PROJECT&&grant.environment==='DEV'&&grant.exportable===false&&grant.fieldMasks.MEMBERS_INGRESS_EXECUTE.includes('preRegister'),'GRANT_REQUIRED');a.equal(grant.revision,expectedGrantRevision,'GRANT_CHANGED');a((grant.expiresAt.toMillis?.()||+grant.expiresAt)>now,'GRANT_EXPIRED');
  const approvals=[];for(const [n,id]of approvalIds.entries()){a(grant.admissionApprovalIds.includes(id),'OUTSIDE_GRANT');const v=await get('memberAdmissionApprovals/'+id);a(v?.synthetic===true&&v.devTechnicalOwnerUid===uid&&grant.syntheticPersonIds.includes(v.devPersonId),'SYNTHETIC_ONLY');a.equal(v.revision,expectedApprovalRevisions[n],'APPROVAL_CHANGED');a.equal(v.registrationDigest,hash({email:v.devRegistration.email,rga:v.devRegistration.rga}),'PAIR_MISMATCH');approvals.push(v);}
  const digest=approvals[0].registrationDigest;a(approvals.every(v=>v.registrationDigest===digest),'PAIR_MISMATCH');
  const id='AP_'+digest.slice(0,40),path='memberAdmissionApprovalIndexes/'+id,old=await get(path),legacy=await get('memberAdmissionApprovals/'+id);a.equal(old?.revision||0,expectedIndexRevision,'INDEX_CHANGED');
  if(old)a(old.schemaVersion==='member-admission-approval-index-v1'&&old.registrationDigest===digest&&Array.isArray(old.approvalIds),'INDEX_INVALID');
  const ids=[...new Set([...(old?.approvalIds||[]),...approvalIds,...(legacy?[id]:[])])].sort();a(ids.length<=16,'INDEX_LIMIT');
  const unchanged=old&&JSON.stringify(old.approvalIds)===JSON.stringify(ids);if(unchanged)return {applied:false,unchanged:true,approvalCount:ids.length,ambiguous:ids.length!==1};
  const next={schemaVersion:'member-admission-approval-index-v1',id,revision:(old?.revision||0)+1,registrationDigest:digest,approvalIds:ids},auditId=hash({uid,id,revision:next.revision}).slice(0,32),auditPath='coreDevAdministrativeAudit/'+auditId;a.equal(await get(auditPath),null,'AUDIT_EXISTS');
  if(apply){tx.set(db.doc(path),next);tx.set(db.doc(auditPath),{schemaVersion:'dev-admission-approval-index-audit-v1',action:'INDEX_EXISTING_APPROVAL',uid,operator,at:new Date(now),before:old,after:next,approvalRevisions:approvals.map(v=>({id:v.id,revision:v.revision})),grantRevision:grant.revision});}
  return {applied:apply,unchanged:false,indexRevision:next.revision,approvalCount:ids.length,ambiguous:ids.length!==1,changedCollections:['memberAdmissionApprovalIndexes','coreDevAdministrativeAudit']};
 });
}
module.exports={indexApproval};
