'use strict';
const {test}=require('node:test'),a=require('node:assert/strict'),{Timestamp}=require('@google-cloud/firestore');
const {administer}=require('../scripts/administer-dev-enrollment.cjs');
function fixture(){
 const now=Date.now(),uid='SYN_OPERATOR',candidateUid='SYN_CANDIDATE',personId='SYN_PERSON',enrollmentId='d'.repeat(32),approvalId='SYN_APPROVAL',projectId='geapa-convergencia-dev',registration={email:'candidate@example.invalid'},docs={};
 docs['coreDevAdministrativeGrants/'+uid]={id:uid,uid,revision:3,active:true,environment:'DEV',projectId,exportable:false,personId:'SYN_OPERATOR_PERSON',bindingVersion:1,subjectEpoch:1,expiresAt:Timestamp.fromMillis(now+86400000),fieldMasks:{MEMBERS_INGRESS_EXECUTE:['preRegister']},syntheticPersonIds:[personId],admissionApprovalIds:[approvalId]};
 docs['memberAuthBindings/'+uid]={active:true,idPessoa:'SYN_OPERATOR_PERSON',bindingVersion:1};docs['memberAuthzSubjects/SYN_OPERATOR_PERSON']={active:true,authzEpoch:1};
 docs['memberEnrollments/'+enrollmentId]={id:enrollmentId,personId,createdBy:uid,approvalId,registration,status:'PRE_REGISTERED'};
 docs['memberAdmissionApprovals/'+approvalId]={active:true,synthetic:true,devTechnicalOwnerUid:uid,devPersonId:personId,devRegistration:registration};docs['memberIdentities/'+personId]={origin:'DEV_ADMIN_SYNTHETIC_V1'};
 docs['memberOperationalConfig/notifications']={enabled:true,mode:'DRY_RUN',projectId,environment:'DEV'};
 docs['memberExecutorScopes/SYN_SCOPE']={synthetic:true,projectId,personIds:[personId]};docs['portalReaderControl/current']={projectId,generation:8};
 docs['commandExecutorRoutes/MEMBERS_ONBOARDING_COMPLETE']={revision:4,projectId,state:'ACTIVE',mode:'DEV_SCOPED_OPERATIONAL',syntheticOnly:true,scopeId:'SYN_SCOPE',actorUids:[uid,'OTHER_SYN'],expiresAt:now-1,operationIds:['a'.repeat(32)]};
 const db={projectId,doc:path=>({path}),runTransaction:async fn=>{const writes=[];const out=await fn({get:async ref=>({exists:!!docs[ref.path],data:()=>docs[ref.path]}),set:(ref,v)=>writes.push([ref.path,v])});for(const [p,v]of writes)docs[p]=v;return out;}};
 const account={uid:candidateUid,email:registration.email,emailVerified:true,disabled:false,providerData:[{providerId:'password'}]},auth={getUser:async()=>account};
 const args={db,auth,uid,candidateUid,enrollmentId,expectedGrantRevision:3,expectedRouteRevision:4,operator:'SYN_OPERATOR',clock:()=>now};return {args,docs,account,now};
}
test('operator grant is dry-run first, bounded to one enrollment, audited and closed without broadening other actors',async()=>{
 const {args,docs,now}=fixture(),original=JSON.stringify(docs),gp='coreDevAdministrativeGrants/'+args.uid,rp='commandExecutorRoutes/MEMBERS_ONBOARDING_COMPLETE';
 a.equal((await administer(args)).applied,false);a.equal(JSON.stringify(docs),original);
 await administer({...args,apply:true});a.equal(docs[gp].invitationWindow.enrollmentId,args.enrollmentId);a.equal(docs[rp].devActorWindow.fallbackExpiresAt,now-1);a.deepEqual(docs[rp].operationIds,['a'.repeat(32)]);a.deepEqual(docs[gp].fieldMasks.MEMBERS_INGRESS_EXECUTE,['preRegister']);
 const open=JSON.stringify(docs);await a.rejects(administer({...args,apply:true}),/GRANT_CHANGED/);a.equal(JSON.stringify(docs),open);
 await administer({...args,apply:true,close:true,expectedGrantRevision:4,expectedRouteRevision:5});a.equal(docs[gp].invitationWindow,undefined);a.equal(docs[rp].devActorWindow.actors[0].expiresAt,now);a.equal(docs[rp].devActorWindow.fallbackExpiresAt,now-1);a.equal(Object.keys(docs).filter(p=>p.startsWith('coreDevAdministrativeAudit/')).length,2);
});
test('issuer rejects production, different binding, real identity, other candidate, unverified account and live notifications without writes',async()=>{
 for(const mutate of [f=>f.args.db.projectId='portal-geapa',f=>f.docs['memberAuthBindings/SYN_OPERATOR'].bindingVersion++,f=>f.docs['memberIdentities/SYN_PERSON'].origin='REAL',f=>f.account.email='other@example.invalid',f=>f.account.emailVerified=false,f=>f.docs['memberOperationalConfig/notifications'].mode='LIVE',f=>f.docs['memberAdmissionApprovals/SYN_APPROVAL'].synthetic=false]){const f=fixture();mutate(f);const before=JSON.stringify(f.docs);await a.rejects(administer({...f.args,apply:true}));a.equal(JSON.stringify(f.docs),before);}
});
