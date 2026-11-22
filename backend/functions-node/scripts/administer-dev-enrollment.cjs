'use strict';
// Operator-only, nominal rehearsal. No roles, approvals or member data are issued.
const a=require('node:assert/strict'),crypto=require('node:crypto');
const PROJECT='geapa-convergencia-dev';
const hash=v=>crypto.createHash('sha256').update(JSON.stringify(v)).digest('hex');
async function administer({db,auth,uid,enrollmentId,candidateUid,expectedGrantRevision,expectedRouteRevision,operator,apply=false,close=false,clock=Date.now}){
 a.equal(db.projectId,PROJECT,'DEV_ONLY');for(const x of [uid,candidateUid])a.match(x,/^[A-Za-z0-9_-]{1,48}$/);a.notEqual(uid,candidateUid);a.match(enrollmentId,/^[a-f0-9]{32}$/);a.equal(typeof operator,'string');a(operator.length>0);
 const candidate=await auth.getUser(candidateUid);a.equal(candidate.uid,candidateUid);a.equal(candidate.emailVerified,true);a(candidate.email.endsWith('@example.invalid'));a(candidate.providerData.length>0&&candidate.providerData.every(p=>p.providerId==='password'));if(!close)a.equal(candidate.disabled,false);
 const now=clock(),end=now+2*3600000,gp='coreDevAdministrativeGrants/'+uid,rp='commandExecutorRoutes/MEMBERS_ONBOARDING_COMPLETE';
 return db.runTransaction(async tx=>{const reads=new Map(),get=async p=>{if(!reads.has(p)){const s=await tx.get(db.doc(p));reads.set(p,s.exists?s.data():null);}return reads.get(p);};
  const g=await get(gp),r=await get(rp),e=await get('memberEnrollments/'+enrollmentId),b=await get('memberAuthBindings/'+uid),subject=b&&await get('memberAuthzSubjects/'+b.idPessoa);
  a.equal(g?.revision,expectedGrantRevision,'GRANT_CHANGED');a.equal(r?.revision,expectedRouteRevision,'ROUTE_CHANGED');a.equal(g.uid,uid);a(g.active&&g.environment==='DEV'&&g.projectId===PROJECT&&g.exportable===false);a(g.fieldMasks.MEMBERS_INGRESS_EXECUTE.includes('preRegister'));a(b?.active&&b.idPessoa===g.personId&&b.bindingVersion===g.bindingVersion&&subject?.active&&subject.authzEpoch===g.subjectEpoch,'BINDING_CHANGED');a(g.expiresAt.toMillis()>end);
  a(e&&e.createdBy===uid&&g.syntheticPersonIds.includes(e.personId)&&g.admissionApprovalIds.includes(e.approvalId),'OUTSIDE_GRANT');a.equal(e.registration.email,candidate.email);a(['PRE_REGISTERED','INVITED','ACCEPTED',...(close?['COMPLETED']:[])].includes(e.status));
  const approval=await get('memberAdmissionApprovals/'+e.approvalId),person=await get('memberIdentities/'+e.personId),candidateBinding=await get('memberAuthBindings/'+candidateUid),notification=await get('memberOperationalConfig/notifications');
  a(approval?.active&&approval.synthetic===true&&approval.devTechnicalOwnerUid===uid&&approval.devPersonId===e.personId);a.equal(hash(approval.devRegistration),hash(e.registration));a.equal(person?.origin,'DEV_ADMIN_SYNTHETIC_V1');a(!candidateBinding||candidateBinding.idPessoa===e.personId,'CANDIDATE_BINDING_CONFLICT');a(notification?.enabled&&notification.mode==='DRY_RUN'&&notification.projectId===PROJECT&&notification.environment==='DEV');
  a(r.projectId===PROJECT&&r.state==='ACTIVE'&&r.mode==='DEV_SCOPED_OPERATIONAL'&&r.syntheticOnly===true);const sp='memberExecutorScopes/'+r.scopeId,s=await get(sp);a(s?.synthetic&&s.projectId===PROJECT);a(s.personIds.includes(e.personId),'TARGET_NOT_IN_EXECUTOR_SCOPE');
  const gatePath='portalReaderControl/current',gate=await get(gatePath);a.equal(gate?.projectId,PROJECT);
  if(close)a.equal(g.invitationWindow?.enrollmentId,enrollmentId,'WINDOW_CHANGED');else a(!g.invitationWindow,'WINDOW_ALREADY_EXISTS');
  const {invitationWindow:previousWindow,...base}=g,nextGrant={...base,revision:g.revision+1,...(close?{}:{invitationWindow:{enrollmentId,issuedAt:new Date(now),expiresAt:new Date(end)}})};
  const prior=r.devActorWindow||{fallbackExpiresAt:r.expiresAt,actors:[]};
  const actors=prior.actors||Object.entries(prior.expiresAtByUid||{}).map(([uid,expiresAt])=>({uid,expiresAt}));
  if(!close)a(!actors.some(x=>x.uid===candidateUid),'CANDIDATE_WINDOW_EXISTS');
  const nextRoute={...r,revision:r.revision+1,actorUids:[...new Set([...r.actorUids,candidateUid])],expiresAt:Math.max(r.expiresAt,end),devActorWindow:{fallbackExpiresAt:prior.fallbackExpiresAt,actors:[...actors.filter(x=>x.uid!==candidateUid),{uid:candidateUid,expiresAt:close?now:end}]}};a(nextRoute.actorUids.length<=16);
  const auditPath='coreDevAdministrativeAudit/'+hash({uid,enrollmentId,revision:nextGrant.revision}).slice(0,32);a.equal(await get(auditPath),null);
  const writes={[gp]:nextGrant,[rp]:nextRoute,[gatePath]:{...gate,generation:gate.generation+1},[auditPath]:{schemaVersion:'core-dev-enrollment-rehearsal-audit-v1',action:close?'CLOSE':'OPEN',operator,uid,candidateUid,enrollmentId,at:new Date(now),expiresAt:new Date(close?now:end),grantRevision:nextGrant.revision,before:{grant:g,route:r},notificationMode:'DRY_RUN'}};
  if(apply)for(const [p,v]of Object.entries(writes))tx.set(db.doc(p),v);
  return {projectId:PROJECT,applied:apply,action:close?'CLOSE':'OPEN',grantRevision:nextGrant.revision,routeRevision:nextRoute.revision,expiresAt:new Date(close?now:end).toISOString(),documents:Object.keys(writes).length,rolesChanged:false,sourceRecordsChanged:false,notificationMode:'DRY_RUN'};
 });
}
module.exports={administer};
