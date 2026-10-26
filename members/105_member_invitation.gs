// Invitation lifecycle inside members.ingress.execute. Pending enrollment is
// never a membership, an Auth binding or an institutional permission.
function f4MemberInvitation_(){
 const C=f4Core_(),N=C.need,U=f4MemberFamilyInputs_(),E=f4MemberEnrollmentContract_();
 function validate(r){
  C.exact(r.target,['personId','linkId']);N(C.id(r.target.personId)&&r.target.linkId===null&&C.id(r.decisionRef)&&r.evidenceRefs.length>0,'COMMAND_PAYLOAD_INVALID');
  C.exact(r.payload,['operation','enrollmentId','action','reason']);N(r.payload.operation==='MANAGE_INVITATION'&&C.op(r.payload.enrollmentId)&&['SEND','RESEND','CANCEL'].includes(r.payload.action),'COMMAND_PAYLOAD_INVALID');
  N(typeof r.payload.reason==='string'&&r.payload.reason.trim().length>0&&r.payload.reason.length<=500&&!/[\x00-\x1f<>]/.test(r.payload.reason),'COMMAND_PAYLOAD_INVALID');return true;
 }
 function proposal(get,r,now){
  validate(r);const p=r.payload,id=r.target.personId,path='memberEnrollments/'+p.enrollmentId,old=E.validate(get(path)),indexPath='memberEnrollmentIndexes/'+id,index=get(indexPath),policy=get('memberOperationalConfig/registration');
  N(old.personId===id&&index?.schemaVersion==='member-enrollment-index-v1'&&index.personId===id&&index.activeEnrollmentId===old.id,'AUTH_BINDING_INVALID');
  const facts=f4MemberActContent_().act(get,r,now).facts;C.exact(facts,['schemaVersion','personId','enrollmentId','enrollmentRevision','policyRevision']);
  N(facts.schemaVersion==='member-invitation-action-v1'&&facts.personId===id&&facts.enrollmentId===old.id&&facts.enrollmentRevision===old.revision,'M13_ACT_INVALID');
  N(policy?.schemaVersion==='member-registration-policy-v1'&&policy.enabled===true&&policy.stagedEnrollment===true&&facts.policyRevision===policy.revision&&Number.isSafeInteger(policy.invitationLifetimeHours)&&policy.invitationLifetimeHours>=1&&policy.invitationLifetimeHours<=720,'ACCESS_REVOKED');
  N(old.approval.reference===r.decisionRef,'AUTH_BINDING_INVALID');
  const registration=old.registration,ref='AP_'+C.hash({email:registration.email,rga:registration.rga}).slice(0,40),approval=f4MemberPreRegistration_().approval(get,ref,registration,now);
  N(approval.reference===old.approval.reference&&approval.membershipType===old.approval.membershipType,'ACCESS_REVOKED');
  N(p.action==='SEND'?old.status==='PRE_REGISTERED':p.action==='RESEND'?['INVITED','ACCEPTED'].includes(old.status):['PRE_REGISTERED','INVITED','ACCEPTED'].includes(old.status),'DOMAIN_STATE_INVALID');
  const revision=U.revision(r,path),expected=U.revision(r,'memberAggregateHeads/'+id),h=get('memberAggregateHeads/'+id);N(old.revision===revision&&h?.revision===expected,'REVISION_CONFLICT');
  const next={...old,revision:old.revision+1,updatedAt:now};
  if(p.action==='CANCEL')next.status='CANCELLED';
  else {if(p.action==='SEND')next.status='INVITED';next.invitedAt=old.invitedAt||now;next.expiresAt=now+policy.invitationLifetimeHours*3600000;}
  E.validate(next);const changes=[{path,expectedRevision:revision,data:next}];
  if(p.action==='CANCEL'){N(U.revision(r,indexPath)===index.revision,'REVISION_CONFLICT');changes.push({path:indexPath,expectedRevision:index.revision,data:{...index,revision:index.revision+1,activeEnrollmentId:null}});}
  const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),target=f4MemberAuthz_().load(get,id,control);
  return {changes,target,effect:'INVITATION',effectiveDate:Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'),expected};
 }
 function notification(get,r){
  validate(r);if(r.payload.action==='CANCEL')return null;
  const p=get('memberOperationalConfig/notifications');
  N(p?.schemaVersion==='member-notification-policy-v1'&&p.projectId===f4ExecutionEnvironment_().project&&p.environment===f4ExecutionEnvironment_().environment&&p.enabled===true&&p.mode==='DRY_RUN'&&C.id(p.approvalRef)&&C.rev(p.revision)&&p.revision>0,'ACCESS_REVOKED');
  // References only. Recipient is resolved from the canonical enrollment by
  // the controlled provider; no browser-selected email reaches an outbox.
  return {schemaVersion:'member-notification-outbox-v1',operationId:r.operationId,commandType:r.commandType,kind:'MEMBERS_NOTIFICATION',event:'ENROLLMENT_INVITATION',state:'PENDING',idempotencyKey:r.operationId,attempts:0,personId:r.target.personId,enrollmentId:r.payload.enrollmentId,policyRevision:p.revision,projectId:p.projectId,environment:p.environment,mode:p.mode,sourceDigest:r.digest};
 }
 function outbox(v,r){const expected=notification(path=>path==='memberOperationalConfig/notifications'?{schemaVersion:'member-notification-policy-v1',projectId:f4ExecutionEnvironment_().project,environment:f4ExecutionEnvironment_().environment,enabled:true,mode:'DRY_RUN',approvalRef:'PROTECTED_POLICY',revision:v?.policyRevision}:null,r);N(expected&&f4MemberCommandCatalog_().equal(expected,v),'COMMAND_PAYLOAD_INVALID');return true;}
 return {validate,proposal,notification,outbox};
}
