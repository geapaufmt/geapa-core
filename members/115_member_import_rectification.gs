// Trusted maintenance only. No browser route, arbitrary patch, policy grant or import.
// Protected documentary approvals select exact people, revisions and operations.
function f4MemberImportRectification_(){
 const C=f4Core_(),F=f4Firestore_(),N=C.need,K=f4MemberAuthzContracts_(),A=f4MemberAuthz_();
 const hash=v=>C.hash(F.encode(v));
 function resolve(get,input,now){
  C.exact(input,['approvalId','personId','operationId']);N(C.id(input.approvalId)&&C.id(input.personId)&&C.op(input.operationId),'COMMAND_PAYLOAD_INVALID');
  const batch=get('memberImportRectificationApprovals/'+input.approvalId);
  C.exact(batch,['schemaVersion','id','revision','active','projectId','environment','approvalRef','approverUid','policyId','issuedAt','expiresAt','cases']);
  N(batch.schemaVersion==='member-import-rectification-approval-v1'&&batch.id===input.approvalId&&batch.revision===1&&batch.active===true&&batch.projectId===f4ExecutionConfig_().projectId&&batch.environment===f4ExecutionConfig_().environment&&C.id(batch.approvalRef)&&C.id(batch.approverUid)&&C.id(batch.policyId),'ACCESS_REVOKED');
  N(C.ms(batch.issuedAt)<=now&&now<C.ms(batch.expiresAt)&&Array.isArray(batch.cases)&&batch.cases.length>0&&batch.cases.length<=100,'ACCESS_REVOKED');
  N(new Set(batch.cases.map(x=>x.personId)).size===batch.cases.length&&new Set(batch.cases.map(x=>x.operationId)).size===batch.cases.length,'SCHEMA_INVALID');
  const item=batch.cases.find(x=>x.personId===input.personId);N(item&&item.operationId===input.operationId,'ACCESS_REVOKED');
  C.exact(item,['personId','operationId','personRevision','identityRevision','identityHash','aggregateHash','sourceSnapshotHash','sourceRow','sourceActive','sourceStatus','cause','evidenceId','eligibleState']);
  N(C.id(item.personId)&&C.op(item.operationId)&&K.positive(item.personRevision)&&item.identityRevision===1&&[item.identityHash,item.aggregateHash,item.sourceSnapshotHash].every(v=>typeof v==='string'&&/^[a-f0-9]{64}$/.test(v))&&Number.isInteger(item.sourceRow)&&item.sourceRow>=2&&item.sourceActive===null&&item.sourceStatus===null&&item.cause==='INCOMPLETE_SOURCE_ACTIVE'&&C.id(item.evidenceId)&&['MEMBRO_INGRESSANTE','MEMBRO_EFETIVO'].includes(item.eligibleState),'SCHEMA_INVALID');
  const actor=A.identify(get,batch.approverUid,now),policy=K.policy(get('memberAuthzPolicies/'+batch.policyId));
  N(policy.proofKind==='ROLE'&&policy.actorSelector.roleKey==='DIRETOR_ADMINISTRATIVO'&&policy.capabilities.includes('MEMBERS_PROFILE_CORRECT'),'AUTHORIZATION_DENIED');
  const fn=actor.derived.currentFunctions.find(f=>f.roleKey===policy.actorSelector.roleKey&&f.scopeKey===policy.actorSelector.actorScope);N(fn,'AUTHORIZATION_DENIED');
  const target={personId:item.personId,scopeKey:'ADMINISTRACAO',boardId:null,roleKey:null,resources:['memberIdentities/'+item.personId],caseId:batch.id,purpose:'IMPORTED_INACTIVITY_RECTIFICATION',fieldMask:[],linkId:null,effect:'PROFILE_CORRECTION',effectiveDate:Utilities.formatDate(new Date(C.ms(batch.issuedAt)),'America/Cuiaba','yyyy-MM-dd')};
  const end=Math.min(now+60000,C.ms(batch.expiresAt)),q={schemaVersion:'member-authz-request-v13',uid:batch.approverUid,kind:'ROLE',sourceId:policy.id,functionId:fn.idVigencia,actorScope:fn.scopeKey,capability:'MEMBERS_PROFILE_CORRECT',expected:{authorityEpoch:actor.control.epoch,policyEpoch:actor.control.policyEpoch,readAuthzEpoch:actor.control.readAuthzEpoch,subjectEpoch:actor.subject.authzEpoch,bindingVersion:actor.binding.bindingVersion,personRevision:actor.actor.h.revision,authorizationRevision:policy.revision},tokenExpiresAt:F.Timestamp.fromMillis(end),admissionExpiresAt:F.Timestamp.fromMillis(end),target,evidenceRefs:[item.evidenceId]};
  const access=A.authorize(get,q,now),e=f4MemberEvidence_().validate(get('memberEvidenceAttestations/'+item.evidenceId));N(e.approvalRef===batch.approvalRef&&e.contentHash===hash(item),'AUTH_BINDING_INVALID');
  const digest=hash({approvalId:batch.id,approvalRevision:batch.revision,item}),receiptPath='memberAudit/'+input.operationId,old=get(receiptPath);
  if(old){N(old.schemaVersion==='member-import-rectification-receipt-v1'&&old.state==='COMPLETED'&&old.digest===digest&&old.personId===item.personId&&old.approverUid===batch.approverUid,'OPERATION_CONFLICT');return {replay:true,receipt:old};}
  const loaded=access.target,person=loaded.d['memberIdentities/'+item.personId];
  N(person.origin==='PESSOAS_BASE'&&person.active===false&&!person.technicalAccount&&person.revision===item.identityRevision&&loaded.h.revision===item.personRevision&&hash(person)===item.identityHash&&hash(loaded.h)===item.aggregateHash,'REVISION_CONFLICT');
  const before=loaded.M.derive(loaded.d,item.personId,loaded.h.revision,now).state;N(before.state===item.eligibleState&&before.blocked===true,'AUTHORIZATION_DENIED');
  // Counterfactual derivation uses the same complete canonical graph. Active is
  // the only prospective business-field change; suspensions/exits still block.
  const next={...person,active:true,revision:person.revision+1},docs={...loaded.d,['memberIdentities/'+item.personId]:next},after=loaded.M.derive(docs,item.personId,loaded.h.revision+1,now).state;
  N(after.state===before.state&&after.blocked===false,'AUTHORIZATION_DENIED');
  const changes=[{path:'memberIdentities/'+item.personId,expectedRevision:person.revision,data:next}];
  return {replay:false,mutation:{id:item.personId,expected:loaded.h.revision,changes,now,boardExpected:{},mode:'M13',deadline:C.ms(access.proof.expiresAt)},receiptPath,receipt:{schemaVersion:'member-import-rectification-receipt-v1',event:'IMPORTED_INACTIVITY_RECTIFIED',state:'COMPLETED',operationId:input.operationId,digest,personId:item.personId,approvalId:batch.id,approvalRef:batch.approvalRef,approverUid:batch.approverUid,actorPersonId:actor.binding.idPessoa,capability:q.capability,policyId:policy.id,policyRevision:policy.revision,evidenceRefs:[item.evidenceId],sourceSnapshotHash:item.sourceSnapshotHash,before:{active:false,identityRevision:person.revision,personRevision:loaded.h.revision,state:before.state,identityHash:hash(person)},after:{active:true,identityRevision:next.revision,personRevision:loaded.h.revision+1,state:after.state,blocked:false,identityHash:hash(next)},createdAt:F.Timestamp.fromMillis(now)}};
 }
 return {resolve};
}
