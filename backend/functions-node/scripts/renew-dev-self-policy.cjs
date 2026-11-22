'use strict';
// Operator-only renewal of the existing DEV SELF policy; never a grant issuer.
const a=require('node:assert/strict'),{createHash}=require('node:crypto'),{createDomain}=require('../src/domain.cjs');
const POLICY='MEMBERS_PROFILE_SELF_DEV',PROJECT='geapa-convergencia-dev';
const ms=v=>v?.toMillis?.()??(v instanceof Date?v.getTime():NaN);
function validate(p){
 const domain=createDomain(()=>{throw Error('POLICY_VALIDATION_IO_DENIED');});
 domain.run('(p)=>{p.issuedAt=codec.Timestamp.fromMillis(p.issuedAt);p.expiresAt=codec.Timestamp.fromMillis(p.expiresAt);f4MemberAuthzContracts_().policy(p);return true;}',{...p,issuedAt:ms(p.issuedAt),expiresAt:ms(p.expiresAt)});
 a.equal(p.id,POLICY);a.equal(p.active,true);a.equal(p.proofKind,'SELF');a.deepEqual(p.capabilities,['MEMBERS_SELF_UPDATE']);a.deepEqual(p.actorSelector,{kind:'SELF'});a.equal(p.targets.ownPerson,true);a.deepEqual(p.targets.commandScopes,['OWN_MEMBER_RESOURCES']);
}
async function renew({db,expectedRevision,approvalRef,operator,hours=4,apply=false,clock=Date.now}){
 a.equal(db.projectId,PROJECT,'DEV_ONLY');a(Number.isSafeInteger(expectedRevision)&&expectedRevision>0);a(Number.isInteger(hours)&&hours>=1&&hours<=24);
 for(const v of [approvalRef,operator])a.match(v,/^[A-Za-z0-9_-]{1,48}$/);
 const now=clock();a(Number.isSafeInteger(now));
 return db.runTransaction(async tx=>{
  const path='memberAuthzPolicies/'+POLICY,old=(await tx.get(db.doc(path))).data();a.equal(old?.revision,expectedRevision,'STALE_REVISION');validate(old);a(ms(old.expiresAt)<=now,'POLICY_NOT_EXPIRED');
  const next={...old,revision:old.revision+1,issuedAt:new Date(now),expiresAt:new Date(now+hours*3600000)};validate(next);
  const id=createHash('sha256').update(JSON.stringify({kind:'DEV_SELF_POLICY_RENEWAL',policyId:POLICY,revision:next.revision,approvalRef})).digest('hex').slice(0,32),audit='coreDevAdministrativeAudit/'+id;
  a.equal((await tx.get(db.doc(audit))).exists,false,'AUDIT_EXISTS');
  if(apply){tx.set(db.doc(path),next);tx.create(db.doc(audit),{schemaVersion:'core-dev-self-policy-renewal-v1',action:'RENEW',operator,approvalRef,policyId:POLICY,previousRevision:old.revision,revision:next.revision,at:new Date(now),previousIssuedAt:old.issuedAt,previousExpiresAt:old.expiresAt,expiresAt:next.expiresAt,capabilities:old.capabilities,targets:old.targets,fieldMasks:old.fieldMasks});}
  return {projectId:PROJECT,policyId:POLICY,applied:apply,revision:next.revision,expiresAt:next.expiresAt.toISOString(),scopePreserved:true,fieldsPreserved:true,identitiesChanged:false,documents:2};
 });
}
module.exports={renew};
