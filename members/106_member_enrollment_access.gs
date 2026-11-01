// Narrow candidate authorization. It neither creates nor substitutes a member
// binding, subject, role, SELF policy or login permission for institutional UI.
// establish is called only with the result of verifyEnrollment server-side.
function f4MemberEnrollmentAccess_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),K=f4MemberCommandCatalog_(),E=f4MemberEnrollmentContract_(),T=F.Timestamp.fromMillis;
 const cap='MEMBERS_ENROLLMENT_SELF',schema='member-enrollment-authz-request-v1';
 function environment(config){C.destination(config);N(f4ExecutionEnvironment_().destination(config),'DESTINATION_INVALID');}
 function source(get,uid,now){
  N(f4ExecutionEnvironment_().destination(f4ExecutionConfig_()),'DESTINATION_INVALID');N(C.id(uid),'AUTH_BINDING_INVALID');const session=get('memberEnrollmentSessions/'+uid);
  C.exact(session,['schemaVersion','uid','personId','enrollmentId','revision','emailClaimId','projectId','environment','issuedAt','expiresAt']);
  N(session.schemaVersion==='member-enrollment-session-v1'&&session.uid===uid&&C.id(session.personId)&&C.op(session.enrollmentId)&&C.rev(session.revision)&&session.revision>0&&K.hex(session.emailClaimId)&&session.projectId===f4ExecutionEnvironment_().project&&session.environment===f4ExecutionEnvironment_().environment&&C.time(session.issuedAt,session.expiresAt,now),'AUTH_BINDING_INVALID');
  const enrollment=E.validate(get('memberEnrollments/'+session.enrollmentId)),claim=get('memberIdentifierClaims/'+session.emailClaimId),identifier=claim&&get('memberIdentifiers/'+claim.identifierId),head=get('memberAggregateHeads/'+session.personId),person=get('memberIdentities/'+session.personId),binding=get('memberAuthBindings/'+uid);
  N(enrollment.personId===session.personId&&(enrollment.uid===null||enrollment.uid===uid)&&claim?.active===true&&claim.idPessoa===session.personId&&identifier?.active===true&&identifier.idPessoa===session.personId&&identifier.kind==='EMAIL'&&identifier.value===enrollment.registration.email&&C.hash({kind:'EMAIL',value:identifier.value})===session.emailClaimId,'AUTH_BINDING_INVALID');
  N(person?.active===true&&person.technicalAccount===false&&head?.id===person.id&&head.paths.includes('memberIdentities/'+person.id)&&head.paths.includes('memberIdentifiers/'+identifier.id)&&C.rev(head.revision)&&head.revision>0,'AUTH_BINDING_INVALID');
  N(!binding||binding.active===true&&binding.uid===uid&&binding.idPessoa===session.personId,'AUTH_BINDING_INVALID');
  const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS'));N(control.mode==='CANONICAL'&&control.activeWriterVersion==='M13','ACCESS_REVOKED');
  const policy=get('memberOperationalConfig/registration');N(policy?.enabled===true&&policy.stagedEnrollment===true&&C.rev(policy.revision)&&policy.revision>0,'ACCESS_REVOKED');const approvalRef='AP_'+C.hash({email:enrollment.registration.email,rga:enrollment.registration.rga}).slice(0,40),approval=f4MemberPreRegistration_().approval(get,approvalRef,enrollment.registration,now);N(approval.reference===enrollment.approval.reference&&approval.membershipType===enrollment.approval.membershipType,'ACCESS_REVOKED');
  return {session,enrollment,head,control,policy,approval};
 }
 function establish(tx,verified,now,config){
  environment(config);C.exact(verified,['uid','expiresAt','projectId','email','emailVerified']);N(verified.projectId===config.projectId&&verified.emailVerified===true&&C.id(verified.uid)&&typeof verified.email==='string'&&Number.isSafeInteger(verified.expiresAt)&&now<verified.expiresAt,'AUTH_BINDING_INVALID');
  const get=p=>tx.get(p).data(),email=verified.email.trim().toLowerCase(),emailClaimId=C.hash({kind:'EMAIL',value:email}),claim=get('memberIdentifierClaims/'+emailClaimId);N(claim?.active===true&&C.id(claim.idPessoa),'AUTH_BINDING_INVALID');
  const index=get('memberEnrollmentIndexes/'+claim.idPessoa);N(index?.schemaVersion==='member-enrollment-index-v1'&&index.personId===claim.idPessoa&&C.op(index.activeEnrollmentId),'ACCESS_REVOKED');
  const enrollment=E.validate(get('memberEnrollments/'+index.activeEnrollmentId));N(['INVITED','ACCEPTED'].includes(enrollment.status)&&enrollment.invitedAt!==null,'ACCESS_REVOKED');E.identity(enrollment,verified,now,config.projectId);
  const path='memberEnrollmentSessions/'+verified.uid,old=get(path);N(!old||old.schemaVersion==='member-enrollment-session-v1'&&old.uid===verified.uid&&old.personId===claim.idPessoa,'AUTH_BINDING_INVALID');
  const data={schemaVersion:'member-enrollment-session-v1',uid:verified.uid,personId:claim.idPessoa,enrollmentId:enrollment.id,revision:(old?.revision||0)+1,emailClaimId,projectId:config.projectId,environment:config.environment,issuedAt:T(now),expiresAt:T(Math.min(verified.expiresAt,enrollment.expiresAt))};
  source(p=>p===path?data:get(p),verified.uid,now);
  // Refreshing login never extends an already prepared proof; revision makes
  // old admission fail closed, while the shared receipt recovers committed work.
  tx.set(path,data);return {personId:enrollment.personId,enrollmentId:enrollment.id,revision:enrollment.revision,expiresAt:C.ms(data.expiresAt)};
 }
 function request(get,identity,r,now,deadline){
  const v=source(get,identity.uid,now);N(r.commandType==='members.onboarding.complete'&&r.payload?.operation==='ENROLLMENT'&&r.target.personId===v.session.personId&&r.target.enrollmentId===v.enrollment.id,'AUTH_BINDING_INVALID');
  return {schemaVersion:schema,kind:'INVITATION',uid:identity.uid,sourceId:identity.uid,capability:cap,expected:{authorityEpoch:v.control.epoch,policyEpoch:v.control.policyEpoch,enrollmentRevision:v.enrollment.revision,personRevision:v.head.revision,sessionRevision:v.session.revision,registrationRevision:v.policy.revision,approvalRevision:v.approval.revision},tokenExpiresAt:T(identity.expiresAt),admissionExpiresAt:T(deadline),target:{personId:v.session.personId,enrollmentId:v.enrollment.id,action:r.payload.action}};
 }
 function authorize(get,q,now){
  C.exact(q,['schemaVersion','kind','uid','sourceId','capability','expected','tokenExpiresAt','admissionExpiresAt','target']);N(q.schemaVersion===schema&&q.kind==='INVITATION'&&q.capability===cap&&q.sourceId===q.uid,'AUTH_BINDING_INVALID');
  C.exact(q.target,['personId','enrollmentId','action']);N(['ACCEPT','DECLINE','SAVE','ACKNOWLEDGE','IMAGE','COMPLETE'].includes(q.target.action),'COMMAND_PAYLOAD_INVALID');
  const v=source(get,q.uid,now),e=v.enrollment;N(e.personId===q.target.personId&&e.id===q.target.enrollmentId&&['INVITED','ACCEPTED'].includes(e.status)&&now<e.expiresAt,'ACCESS_REVOKED');
  N(get('memberEnrollmentIndexes/'+e.personId)?.activeEnrollmentId===e.id,'ACCESS_REVOKED');
  const expected={authorityEpoch:v.control.epoch,policyEpoch:v.control.policyEpoch,enrollmentRevision:e.revision,personRevision:v.head.revision,sessionRevision:v.session.revision,registrationRevision:v.policy.revision,approvalRevision:v.approval.revision};N(K.equal(q.expected,expected),'REVISION_CONFLICT');
  const end=Math.min(C.ms(q.tokenExpiresAt),C.ms(q.admissionExpiresAt),C.ms(v.session.expiresAt),e.expiresAt);N(Number.isSafeInteger(end)&&now<end,'ACCESS_REVOKED');
  return {request:q,proof:{schemaVersion:'member-enrollment-authz-proof-v1',kind:'INVITATION',uid:q.uid,idPessoa:e.personId,capability:cap,authorityEpoch:v.control.epoch,policyEpoch:v.control.policyEpoch,subjectEpoch:0,bindingVersion:0,personRevision:v.head.revision,policyId:'ENROLLMENT_'+e.id,policyRevision:e.revision,sessionRevision:v.session.revision,registrationRevision:v.policy.revision,approvalRevision:v.approval.revision,targetsDigest:C.hash(q.target),issuedAt:T(now),expiresAt:T(end)},...v};
 }
 function recheck(get,q,proof,now){N(C.time(proof.issuedAt,proof.expiresAt,now),'ACCESS_REVOKED');const next=authorize(get,q,now).proof,stable=p=>Object.fromEntries(Object.entries(p).filter(([k])=>!['issuedAt','expiresAt'].includes(k)));N(K.equal(stable(next),stable(proof))&&C.ms(next.expiresAt)>=C.ms(proof.expiresAt),'ACCESS_REVOKED');return next;}
 function statusAccess(get,identity,key,now){
  N(f4ExecutionEnvironment_().destination(f4ExecutionConfig_())&&now<identity.expiresAt,'ACCESS_REVOKED');
  const r=get('commandRequests/'+key.operationId);N(r&&r.ownerUid===identity.uid&&r.commandType===key.commandType&&r.digest===key.digest&&f4MemberEnrollmentCommands_().is(r),'ACCESS_REVOKED');
  K.envelope(f4ExecutionEnvironment_().project,r,f4MemberEnrollmentCommands_().validate);
  const t=get('memberAdmissionTickets/'+r.ticketId),proof=t?.authProofs?.[0],binding=get('memberAuthBindings/'+identity.uid);
  N(t?.ownerUid===identity.uid&&t.operationId===r.operationId&&t.digest===r.digest&&t.idPessoa===r.target.personId&&t.authProofs.length===1&&K.equal(t.requestEnvelope,K.template(r))&&proof?.schemaVersion==='member-enrollment-authz-proof-v1'&&proof.uid===identity.uid&&proof.idPessoa===r.target.personId&&proof.capability===cap,'ACCESS_REVOKED');
  N(!binding||binding.active===true&&binding.idPessoa===r.target.personId,'ACCESS_REVOKED');
  // Receipt-only access survives completion/decline and token renewal. It is
  // never used to authorize execution: claim/finalize still recheck the source.
  return {proof};
 }
 return {cap,schema,source,establish,request,authorize,recheck,statusAccess};
}
