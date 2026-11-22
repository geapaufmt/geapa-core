'use strict';
// An explicit synthetic decision authorizes one operation, never an institutional role.
const a=require('node:assert/strict'),crypto=require('node:crypto');
const PROJECT='geapa-convergencia-dev',hash=v=>crypto.createHash('sha256').update(JSON.stringify(v)).digest('hex');
async function administer({db,uid,enrollmentId,recognizedOn,reference,expectedGrantRevision,expectedPersonRevision,operator,apply=false,close=false,clock=Date.now}){
 a.equal(db.projectId,PROJECT,'DEV_ONLY');a.match(uid,/^[A-Za-z0-9_-]{1,48}$/);a.match(enrollmentId,/^[a-f0-9]{32}$/);a.match(reference,/^TESTE_[A-Z0-9_]{1,120}$/);a.equal(typeof operator,'string');a(operator.length>0);
 const now=clock(),today=new Intl.DateTimeFormat('en-CA',{timeZone:'America/Cuiaba',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(now));if(!close)a.equal(recognizedOn,today,'SYNTHETIC_TODAY_ONLY');
 return db.runTransaction(async tx=>{const get=async p=>{const s=await tx.get(db.doc(p));return s.exists?s.data():null;};
  const gp='coreDevAdministrativeGrants/'+uid,g=await get(gp),e=await get('memberEnrollments/'+enrollmentId),b=await get('memberAuthBindings/'+uid),subject=b&&await get('memberAuthzSubjects/'+b.idPessoa);
  a.equal(g?.revision,expectedGrantRevision,'GRANT_CHANGED');a(g.uid===uid&&g.active&&g.environment==='DEV'&&g.projectId===PROJECT&&g.exportable===false);a(g.capabilities.includes('MEMBERS_INGRESS_EXECUTE'));a(b?.active&&b.idPessoa===g.personId&&b.bindingVersion===g.bindingVersion&&subject?.active&&subject.authzEpoch===g.subjectEpoch,'BINDING_CHANGED');
  a(e?.status==='COMPLETED'&&e.createdBy===uid&&g.syntheticPersonIds.includes(e.personId)&&g.admissionApprovalIds.includes(e.approvalId),'OUTSIDE_GRANT');a.notEqual(e.personId,g.personId);
  const approval=await get('memberAdmissionApprovals/'+e.approvalId),person=await get('memberIdentities/'+e.personId),head=await get('memberAggregateHeads/'+e.personId),state=await get('memberCurrentStates/'+e.personId);
  a.equal(head?.revision,expectedPersonRevision,'PERSON_CHANGED');a.equal(state?.sourceRevision,head.revision);if(!close)a.equal(state.state,'MEMBRO_INGRESSANTE');a.equal(person?.origin,'DEV_ADMIN_SYNTHETIC_V1');a(approval?.active&&approval.synthetic===true&&approval.devTechnicalOwnerUid===uid&&approval.devPersonId===e.personId);a.equal(hash(approval.devRegistration),hash(e.registration));
  const r=await get('commandExecutorRoutes/MEMBERS_INGRESS_EXECUTE'),s=r&&await get('memberExecutorScopes/'+r.scopeId),gate=await get('portalReaderControl/current');a(r?.projectId===PROJECT&&r.state==='ACTIVE'&&r.syntheticOnly&&r.actorUids.includes(uid));a(s?.synthetic&&s.projectId===PROJECT&&s.personIds.includes(e.personId));a.equal(gate?.projectId,PROJECT);
  const actor=r.devActorWindow?.actors?.find(v=>v.uid===uid),routeEnd=r.devActorWindow?(actor?.expiresAt??r.devActorWindow.fallbackExpiresAt):r.expiresAt;
  const end=Math.min(now+2*3600000,g.expiresAt.toMillis(),routeEnd);if(!close)a(end>now+60000,'EXECUTOR_WINDOW_EXPIRED');
  const window={enrollmentId,personId:e.personId,recognizedOn,reference,issuedAt:new Date(now),expiresAt:new Date(end)};
  if(close){for(const key of ['enrollmentId','personId','recognizedOn','reference'])a.equal(g.effectivationWindow?.[key],window[key],'WINDOW_CHANGED');}else a(!g.effectivationWindow,'WINDOW_ALREADY_EXISTS');
  const {effectivationWindow:prior,...base}=g,next={...base,revision:g.revision+1,...(close?{}:{effectivationWindow:window})};
  const ap='coreDevAdministrativeAudit/'+hash({uid,enrollmentId,reference,revision:next.revision}).slice(0,32);a.equal(await get(ap),null);
  if(apply){tx.set(db.doc(gp),next);tx.set(db.doc('portalReaderControl/current'),{...gate,generation:gate.generation+1});tx.set(db.doc(ap),{schemaVersion:'core-dev-effectivation-rehearsal-audit-v1',action:close?'CLOSE':'OPEN',uid,enrollmentId,personId:e.personId,recognizedOn,reference,at:new Date(now),operator,before:g,grantRevision:next.revision});}
  return {projectId:PROJECT,applied:apply,action:close?'CLOSE':'OPEN',grantRevision:next.revision,expiresAt:close?null:new Date(end).toISOString(),documents:3,rolesChanged:false,executorChanged:false,sourceRecordsChanged:false};
 });
}
module.exports={administer};
