// Invitation lifecycle inside members.ingress.execute. Pending enrollment is
// never a membership, an Auth binding or an institutional permission.
function f4MemberInvitation_(){
 const C=f4Core_(),N=C.need,U=f4MemberFamilyInputs_(),E=f4MemberEnrollmentContract_();
 const reviewActions=['ACKNOWLEDGE','REQUEST_CORRECTION'];
 const admissionFields=['membershipType','approvedOn','confirmed'];
 const rh=actor=>actor?.derived?.currentFunctions?.some(f=>f.roleKey==='DIRETOR_RECURSOS_HUMANOS'&&f.scopeKey==='RECURSOS_HUMANOS'&&f.group==='DIRETORIA')===true;
 function reviewer(actor,policy){
  const rules=policy?.presidencyReview;
  return actor?.derived?.currentFunctions?.some(f=>f.roleKey==='PRESIDENTE'||f.roleKey==='VICE_PRESIDENTE'&&rules?.viceAuthorized===true&&C.id(rules.decisionRef))===true;
 }
 function requiresReview(policy){
  const rules=policy?.presidencyReview;
  if(rules?.required!==true)return false;
  N(C.id(rules.decisionRef),'ADMISSION_APPROVAL_REQUIRED');return true;
 }
 function validate(r){
  C.exact(r.target,['personId','linkId']);N(C.id(r.target.personId)&&r.target.linkId===null&&(r.decisionRef===null||C.id(r.decisionRef))&&r.evidenceRefs.length>0,'COMMAND_PAYLOAD_INVALID');
  C.exact(r.payload,['operation','enrollmentId','action','reason',...(r.payload.action==='AMEND'?['field','value']:r.payload.action==='CONFIRM_ADMISSION'?admissionFields:[])]);N(r.payload.operation==='MANAGE_INVITATION'&&C.op(r.payload.enrollmentId)&&['SEND','RESEND','CANCEL',...reviewActions,'RESOLVE_CORRECTION','AMEND','CONFIRM_ADMISSION'].includes(r.payload.action),'COMMAND_PAYLOAD_INVALID');
  N(typeof r.payload.reason==='string'&&r.payload.reason.trim().length>0&&r.payload.reason.length<=500&&!/[\x00-\x1f<>]/.test(r.payload.reason),'COMMAND_PAYLOAD_INVALID');return true;
 }
 function decision(get,old,r,now){
  const p=r.payload,actor=f4MemberAuthz_().identify(get,r.ownerUid,now);
  const D=f4DevAdministrativeAccess_(),technical=!rh(actor)&&D.source(get,{uid:r.ownerUid,expiresAt:now+60000},now,'MEMBERS_INGRESS_EXECUTE');
  N((rh(actor)||technical&&D.decisionAllowed(get,technical,old,now,p))&&actor.binding.idPessoa!==old.personId,'ACCESS_REVOKED');
  N(old.draft===true&&old.approval===null&&['PRE_REGISTERED','INVITED','ACCEPTED'].includes(old.status),'DOMAIN_STATE_INVALID');
  N(p.confirmed===true&&typeof p.reason==='string'&&p.reason.trim().length>=10,'COMMAND_PAYLOAD_INVALID');
  N(['MEMBRO_INGRESSANTE','MEMBRO_EFETIVO'].includes(p.membershipType),'ADMISSION_CATEGORY_INVALID');
  const day=t=>Utilities.formatDate(new Date(t),'America/Cuiaba','yyyy-MM-dd');C.civilStart(p.approvedOn);N(p.approvedOn>=day(old.createdAt)&&p.approvedOn<=day(now),'COMMAND_PAYLOAD_INVALID');
  const id='AA_'+r.operationId,path='memberAdmissionApprovals/'+id,indexId=f4MemberPreRegistration_().key(old.registration),indexPath='memberAdmissionApprovalIndexes/'+indexId,index=get(indexPath),legacy=get('memberAdmissionApprovals/'+indexId);
  N(!get(path)&&!legacy,'OPERATION_CONFLICT');
  const retire=[];
  if(index){
   N(index.schemaVersion==='member-admission-approval-index-v1'&&index.id===indexId&&index.registrationDigest===C.hash({email:old.registration.email,rga:old.registration.rga})&&C.rev(index.revision)&&index.revision>0&&index.personId===old.personId&&Array.isArray(index.approvalIds)&&index.approvalIds.length===1,'ADMISSION_APPROVAL_AMBIGUOUS');
   const prior=get('memberAdmissionApprovals/'+index.approvalIds[0]);N(prior&&typeof prior.active==='boolean'&&prior.id===index.approvalIds[0]&&prior.personId===old.personId&&prior.enrollmentId===index.enrollmentId&&prior.registrationDigest===index.registrationDigest,'ADMISSION_APPROVAL_AMBIGUOUS');
   if(index.enrollmentId!==old.id){
    // Terminal enrollments retain their history. Only their exact decision is
    // retired, atomically with the new decision/index, under its prior revision.
    const previous=E.validate(get('memberEnrollments/'+index.enrollmentId));
    N(previous.personId===old.personId&&['CANCELLED','DECLINED'].includes(previous.status),'OPERATION_CONFLICT');
    if(prior.active){N(previous.approvalId===prior.id&&f4MemberCommandCatalog_().equal(previous.approval,f4MemberPreRegistration_().snapshot(prior)),'ADMISSION_APPROVAL_AMBIGUOUS');retire.push({path:'memberAdmissionApprovals/'+prior.id,expectedRevision:prior.revision,data:{...prior,revision:prior.revision+1,active:false}});}
   }else N(prior.active===false,'ADMISSION_APPROVAL_AMBIGUOUS');
  }
  const registrationDigest=C.hash({email:old.registration.email,rga:old.registration.rga}),approval={schemaVersion:'member-admission-approval-v1',id,personId:old.personId,enrollmentId:old.id,revision:1,active:true,registrationDigest,reference:'DEC_'+r.operationId,approvedBy:actor.binding.idPessoa,approvedOn:p.approvedOn,membershipType:p.membershipType,decisionOperationId:r.operationId,reason:p.reason};
  return [{path,expectedRevision:0,data:approval},{path:indexPath,expectedRevision:index?.revision||0,data:{schemaVersion:'member-admission-approval-index-v1',id:indexId,personId:old.personId,enrollmentId:old.id,revision:(index?.revision||0)+1,registrationDigest,approvalIds:[id]}},...retire];
 }
 function approvalChange(c,id,prior){
  const d=c.data;N(d.personId===id&&C.op(d.enrollmentId)&&C.rev(c.expectedRevision)&&d.revision===c.expectedRevision+1&&/^[a-f0-9]{64}$/.test(d.registrationDigest),'COMMAND_PAYLOAD_INVALID');
  if(c.path.startsWith('memberAdmissionApprovalIndexes/')){
   C.exact(d,['schemaVersion','id','personId','enrollmentId','revision','registrationDigest','approvalIds']);N(d.schemaVersion==='member-admission-approval-index-v1'&&d.id==='AP_'+d.registrationDigest.slice(0,40)&&c.path==='memberAdmissionApprovalIndexes/'+d.id&&d.approvalIds.length===1&&C.id(d.approvalIds[0]),'COMMAND_PAYLOAD_INVALID');
  }else{
   C.exact(d,['schemaVersion','id','personId','enrollmentId','revision','active','registrationDigest','reference','approvedBy','approvedOn','membershipType','decisionOperationId','reason']);
   N(d.schemaVersion==='member-admission-approval-v1'&&C.op(d.decisionOperationId)&&d.id==='AA_'+d.decisionOperationId&&d.reference==='DEC_'+d.decisionOperationId&&c.path==='memberAdmissionApprovals/'+d.id&&C.id(d.approvedBy)&&['MEMBRO_INGRESSANTE','MEMBRO_EFETIVO'].includes(d.membershipType)&&typeof d.reason==='string'&&d.reason.trim().length>=10&&d.reason.length<=500,'COMMAND_PAYLOAD_INVALID');C.civilStart(d.approvedOn);
   N(c.expectedRevision===0?d.active===true:prior?.active===true&&f4MemberCommandCatalog_().equal(d,{...prior,revision:prior.revision+1,active:false}),'COMMAND_PAYLOAD_INVALID');
  }return true;
 }
 function amendment(get,old,p,now){
  const field=p.field,id=old.personId;
  N(['fullName','phone','rga','email'].includes(field),'COMMAND_PAYLOAD_INVALID');
  const registration=E.quick({...old.registration,[field]:p.value});
  N(registration[field]!==old.registration[field],'COMMAND_PAYLOAD_INVALID');
  if(field==='phone')return {registration};
  if(field==='fullName'){
   const path='memberIdentities/'+id,prior=get(path);
   N(prior?.idPessoa===id&&prior.active,'AUTH_BINDING_INVALID');
   return {registration,path,prior,data:{...prior,displayName:registration.fullName,revision:prior.revision+1}};
  }
  N(['PRE_REGISTERED','INVITED'].includes(old.status)||field==='rga'&&old.status==='ACCEPTED','DOMAIN_STATE_INVALID');
  const kind=field.toUpperCase(),claim=get('memberIdentifierClaims/'+C.hash({kind,value:old.registration[field]}));
  N(claim?.active&&claim.idPessoa===id&&C.id(claim.identifierId),'AUTH_BINDING_INVALID');
  const path='memberIdentifiers/'+claim.identifierId,prior=get(path),head=get('memberAggregateHeads/'+id);
  N(prior?.active&&prior.kind===kind&&prior.idPessoa===id&&prior.value===old.registration[field]&&head?.paths.includes(path),'AUTH_BINDING_INVALID');
  const target=get('memberIdentifierClaims/'+C.hash({kind,value:registration[field]}));
  N(!target?.active,'OPERATION_CONFLICT');
  const revokedApproval=old.draft===true&&old.approval!==null?f4MemberPreRegistration_().approval(get,old.approvalId,old.registration,now):null;
  const approval=old.approval===null||old.draft===true?null:f4MemberPreRegistration_().locate(get,registration,old.approval.reference,now);
  if(approval)N(approval.reference===old.approval.reference&&approval.membershipType===old.approval.membershipType,'ADMISSION_APPROVAL_REQUIRED');
  return {registration,path,prior,data:{...prior,value:registration[field],revision:prior.revision+1},approval,revokedApproval};
 }
 function proposal(get,r,now){
  validate(r);const p=r.payload,id=r.target.personId,path='memberEnrollments/'+p.enrollmentId,old=E.validate(get(path)),indexPath='memberEnrollmentIndexes/'+id,index=get(indexPath),policy=get('memberOperationalConfig/registration');
  N(old.personId===id&&index?.schemaVersion==='member-enrollment-index-v1'&&index.personId===id&&index.activeEnrollmentId===old.id,'AUTH_BINDING_INVALID');
  const facts=f4MemberActContent_().act(get,r,now).facts;C.exact(facts,['schemaVersion','personId','enrollmentId','enrollmentRevision','policyRevision']);
  N(facts.schemaVersion==='member-invitation-action-v1'&&facts.personId===id&&facts.enrollmentId===old.id&&facts.enrollmentRevision===old.revision,'M13_ACT_INVALID');
  N(policy?.schemaVersion==='member-registration-policy-v1'&&policy.enabled===true&&policy.stagedEnrollment===true&&facts.policyRevision===policy.revision&&Number.isSafeInteger(policy.invitationLifetimeHours)&&policy.invitationLifetimeHours>=1&&policy.invitationLifetimeHours<=720,'ACCESS_REVOKED');
  N((p.action==='CONFIRM_ADMISSION'?'DEC_'+r.operationId:old.approval?.reference||null)===r.decisionRef,'AUTH_BINDING_INVALID');
  if(old.approval!==null&&!(p.action==='AMEND'&&['rga','email'].includes(p.field))){
   const ref=old.approvalId||f4MemberPreRegistration_().key(old.registration);
   const approval=f4MemberPreRegistration_().approval(get,ref,old.registration,now);
   N(approval.reference===old.approval.reference&&approval.membershipType===old.approval.membershipType,'ACCESS_REVOKED');
  }
  const reviewing=reviewActions.includes(p.action),resolving=p.action==='RESOLVE_CORRECTION';
  if(reviewing){
   const actor=f4MemberAuthz_().identify(get,r.ownerUid,now);
   N(reviewer(actor,policy)&&actor.binding.idPessoa!==id,'ACCESS_REVOKED');
   N(p.action!=='ACKNOWLEDGE'||old.review?.state!=='CORRECTION_REQUESTED','DOMAIN_STATE_INVALID');
  }
  if(resolving)N(old.review?.state==='CORRECTION_REQUESTED','DOMAIN_STATE_INVALID');
  if(['SEND','RESEND'].includes(p.action)&&requiresReview(policy))N(old.review?.state==='ACKNOWLEDGED','ADMISSION_APPROVAL_REQUIRED');
  N(p.action==='SEND'?old.status==='PRE_REGISTERED':p.action==='RESEND'?['INVITED','ACCEPTED'].includes(old.status):['PRE_REGISTERED','INVITED','ACCEPTED'].includes(old.status),'DOMAIN_STATE_INVALID');
  const revision=U.revision(r,path),expected=U.revision(r,'memberAggregateHeads/'+id),h=get('memberAggregateHeads/'+id);N(old.revision===revision&&h?.revision===expected,'REVISION_CONFLICT');
  const next={...old,revision:old.revision+1,updatedAt:now};
  const amend=p.action==='AMEND'?amendment(get,old,p,now):null,decisions=p.action==='CONFIRM_ADMISSION'?decision(get,old,r,now):[];
  if(amend){
   next.registration=amend.registration;
   if(amend.approval){const a=amend.approval;next.approvalId=a.id;next.approval=f4MemberPreRegistration_().snapshot(a);}
   if(amend.revokedApproval){next.approval=null;delete next.approvalId;}
   next.review={state:'RESOLVED',at:now,actorUid:r.ownerUid,note:p.reason};
  }else if(decisions.length){const a=decisions[0].data;next.approvalId=a.id;next.approval=f4MemberPreRegistration_().snapshot(a);}
  else if(reviewing||resolving)next.review={state:p.action==='ACKNOWLEDGE'?'ACKNOWLEDGED':resolving?'RESOLVED':'CORRECTION_REQUESTED',at:now,actorUid:r.ownerUid,note:p.reason};
  else if(p.action==='CANCEL')next.status='CANCELLED';
  else {if(p.action==='SEND')next.status='INVITED';next.invitedAt=old.invitedAt||now;next.expiresAt=now+policy.invitationLifetimeHours*3600000;}
  E.validate(next);const changes=[{path,expectedRevision:revision,data:next}];
  if(amend?.path){N(U.revision(r,amend.path)===amend.prior.revision,'REVISION_CONFLICT');changes.push({path:amend.path,expectedRevision:amend.prior.revision,data:amend.data});}
  if(amend?.revokedApproval){const a=amend.revokedApproval;decisions.push({path:'memberAdmissionApprovals/'+a.id,expectedRevision:a.revision,data:{...a,revision:a.revision+1,active:false}});}
  for(const c of decisions){N(U.revision(r,c.path)===c.expectedRevision,'REVISION_CONFLICT');changes.push(c);}
  if(p.action==='CANCEL'){N(U.revision(r,indexPath)===index.revision,'REVISION_CONFLICT');changes.push({path:indexPath,expectedRevision:index.revision,data:{...index,revision:index.revision+1,activeEnrollmentId:null}});}
  const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),target=f4MemberAuthz_().load(get,id,control);
  return {changes,target,effect:'INVITATION',effectiveDate:Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'),expected};
 }
 function notification(get,r){
  validate(r);if(!['SEND','RESEND'].includes(r.payload.action))return null;
  const p=get('memberOperationalConfig/notifications');
  N(p?.schemaVersion==='member-notification-policy-v1'&&p.projectId===f4ExecutionEnvironment_().project&&p.environment===f4ExecutionEnvironment_().environment&&p.enabled===true&&p.mode==='DRY_RUN'&&C.id(p.approvalRef)&&C.rev(p.revision)&&p.revision>0,'ACCESS_REVOKED');
  // References only. Recipient is resolved from the canonical enrollment by
  // the controlled provider; no browser-selected email reaches an outbox.
  return {schemaVersion:'member-notification-outbox-v1',operationId:r.operationId,commandType:r.commandType,kind:'MEMBERS_NOTIFICATION',event:'ENROLLMENT_INVITATION',state:'PENDING',idempotencyKey:r.operationId,attempts:0,personId:r.target.personId,enrollmentId:r.payload.enrollmentId,policyRevision:p.revision,projectId:p.projectId,environment:p.environment,mode:p.mode,sourceDigest:r.digest};
 }
 function outbox(v,r){const expected=notification(path=>path==='memberOperationalConfig/notifications'?{schemaVersion:'member-notification-policy-v1',projectId:f4ExecutionEnvironment_().project,environment:f4ExecutionEnvironment_().environment,enabled:true,mode:'DRY_RUN',approvalRef:'PROTECTED_POLICY',revision:v?.policyRevision}:null,r);N(expected&&f4MemberCommandCatalog_().equal(expected,v),'COMMAND_PAYLOAD_INVALID');return true;}
 return {validate,proposal,notification,outbox,reviewer,requiresReview,amendment,rh,admissionFields,decision,approvalChange};
}
