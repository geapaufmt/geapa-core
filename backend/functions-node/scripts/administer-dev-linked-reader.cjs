'use strict';
// Operator-only, CAS and audit. No Auth, institutional policy or route changes.
const a=require('node:assert/strict'),{createHash}=require('node:crypto');
async function administer({db,uid,personId,approvalRef,operator,expectedRevision,hours=4,apply=false,revoke=false,clock=Date.now}){
 a.equal(db.projectId,'geapa-convergencia-dev','DEV_ONLY');a.match(uid,/^[A-Za-z0-9_-]{1,128}$/);a.match(personId,/^DEVADM_[a-f0-9]{12}_[12]$/);a.match(approvalRef,/^[A-Za-z0-9_-]{1,80}$/);a.equal(typeof operator,'string');a(operator.length>0&&operator.length<=128);a(Number.isInteger(expectedRevision)&&expectedRevision>=0);a(Number.isInteger(hours)&&hours>0&&hours<=4);
 const now=clock(),path='coreDevLinkedReaderGrants/'+uid;
 return db.runTransaction(async tx=>{const get=async p=>{const s=await tx.get(db.doc(p));return s.exists?s.data():null;},old=await get(path),b=await get('memberAuthBindings/'+uid),p=await get('memberIdentities/'+personId);a.equal(old?.revision||0,expectedRevision,'STALE_REVISION');a.equal(b?.active,true,'BINDING_REQUIRED');a.equal(b.idPessoa,personId);a(Number.isSafeInteger(b.bindingVersion)&&b.bindingVersion>0);a.equal(p?.origin,'DEV_ADMIN_SYNTHETIC_V1');a.equal(p.active,true);a.equal(p.idPessoa,personId);
 const next={schemaVersion:'core-dev-linked-reader-grant-v1',projectId:'geapa-convergencia-dev',uid,personId,bindingVersion:b.bindingVersion,revision:expectedRevision+1,active:!revoke,startsAt:now,expiresAt:now+hours*3600000,approvalRef};
 const id=createHash('sha256').update(JSON.stringify({kind:'DEV_LINKED_READER',uid,revision:next.revision,approvalRef})).digest('hex').slice(0,32),audit='coreDevAdministrativeAudit/'+id;a.equal(await get(audit),null,'AUDIT_EXISTS');
 if(apply){tx.set(db.doc(path),next);tx.create(db.doc(audit),{schemaVersion:'core-dev-linked-reader-audit-v1',at:now,operator,approvalRef,personId,grantRevision:next.revision,action:revoke?'REVOKE':'ISSUE',expiresAt:next.expiresAt});}
 return {applied:apply,revision:next.revision,active:next.active,expiresAt:next.expiresAt,personId,changedCollections:['coreDevLinkedReaderGrants','coreDevAdministrativeAudit'],institutionalPermissionsChanged:false};
 });
}
module.exports={administer};

