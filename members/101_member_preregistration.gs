// PRE_REGISTER extends the existing ingress command. It reserves canonical
// identity and records a draft only; no memberLink or INGRESSION event exists.
function f4MemberPreRegistration_(){
 const C=f4Core_(),N=C.need,U=f4MemberFamilyInputs_(),E=f4MemberEnrollmentContract_(),K=f4MemberCommandCatalog_();
 function validate(r){
  C.exact(r.target,['personId','linkId']);N(C.id(r.target.personId)&&r.target.linkId===null&&r.decisionRef!==null&&r.evidenceRefs.length>0,'COMMAND_PAYLOAD_INVALID');
  C.exact(r.payload,['operation','identity','identifiers','registration','profile','approvalRef']);N(r.payload.operation==='PRE_REGISTER'&&C.id(r.payload.approvalRef)&&Array.isArray(r.payload.identifiers),'COMMAND_PAYLOAD_INVALID');
  if(r.payload.identity!==null){C.exact(r.payload.identity,['displayName']);N(typeof r.payload.identity.displayName==='string','COMMAND_PAYLOAD_INVALID');}
  for(const v of r.payload.identifiers){C.exact(v,['id','kind','value']);N(C.id(v.id)&&['EMAIL','RGA'].includes(v.kind),'COMMAND_PAYLOAD_INVALID');}
  E.quick(r.payload.registration);return true;
 }
 function approval(get,ref,registration,now){
  const d=get('memberAdmissionApprovals/'+ref);N(d?.schemaVersion==='member-admission-approval-v1'&&d.id===ref&&d.active===true&&C.rev(d.revision)&&d.revision>0,'ACCESS_REVOKED');
  N(d.registrationDigest===C.hash({email:registration.email,rga:registration.rga})&&C.id(d.reference)&&C.id(d.approvedBy),'AUTH_BINDING_INVALID');
  C.civilStart(d.approvedOn);N(C.civilStart(d.approvedOn)<=now&&['MEMBRO_INGRESSANTE','MEMBRO_EFETIVO'].includes(d.membershipType),'COMMAND_PAYLOAD_INVALID');
  return d;
 }
 function proposal(get,r,now){
  validate(r);const p=r.payload,id=r.target.personId,body=f4MemberActContent_().act(get,r,now),facts=body.facts;
  C.exact(facts,['schemaVersion','personId','approvalRef','approvalRevision','policyRevision']);
  N(facts.schemaVersion==='member-preregistration-v1'&&facts.personId===id&&facts.approvalRef===p.approvalRef,'M13_ACT_INVALID');
  const policy=get('memberOperationalConfig/registration');N(policy?.schemaVersion==='member-registration-policy-v1'&&policy.enabled===true&&policy.stagedEnrollment===true&&policy.revision===facts.policyRevision&&Number.isSafeInteger(policy.invitationLifetimeHours)&&policy.invitationLifetimeHours>0&&policy.invitationLifetimeHours<=720,'ACCESS_REVOKED');
  const today=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'),registration=E.quick(p.registration),a=approval(get,p.approvalRef,registration,now);
  N(a.revision===facts.approvalRevision&&a.reference===r.decisionRef&&K.equal(registration,p.registration),'REVISION_CONFLICT');
  const resolution=f4MemberRegistrationIdentity_().resolve(get,{email:registration.email,rga:registration.rga});N(resolution.personId===null||resolution.personId===id,'AUTH_BINDING_INVALID');
  const identity=get('memberIdentities/'+id),h=get('memberAggregateHeads/'+id),expected=U.revision(r,'memberAggregateHeads/'+id),control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS'));
  N((h?.revision||0)===expected,'REVISION_CONFLICT');N(identity?p.identity===null:!h&&p.identity?.displayName===registration.fullName,'COMMAND_PAYLOAD_INVALID');
  if(identity){const loaded=f4MemberAuthz_().load(get,id,control),s=loaded.M.derive(loaded.d,id,loaded.h.revision,now).state;N(['SEM_VINCULO','DESLIGADO','EGRESSO','MEMBRO_EM_ESPERA'].includes(s.state),'DOMAIN_STATE_INVALID');}
  const changes=[],entity=(col,key,fields)=>({schemaVersion:'m1-'+col+'-v1',id:key,revision:1,origin:'M13_COMMAND',...fields});
  function add(path,data){const old=get(path),revision=U.revision(r,path);N((old?.revision||0)===revision&&data.revision===revision+1,'REVISION_CONFLICT');changes.push({path,expectedRevision:revision,data});}
  if(!identity)add('memberIdentities/'+id,entity('memberIdentities',id,{idPessoa:id,displayName:registration.fullName,technicalAccount:false,active:true}));
  for(const item of p.identifiers){N(!get('memberIdentifiers/'+item.id)&&resolution.identifiers.some(v=>v.kind===item.kind&&v.value===item.value),'COMMAND_PAYLOAD_INVALID');add('memberIdentifiers/'+item.id,entity('memberIdentifiers',item.id,{idPessoa:id,kind:item.kind,value:item.value,active:true}));}
  for(const item of resolution.identifiers){const claim=get('memberIdentifierClaims/'+C.hash(item));N(claim?.active&&claim.idPessoa===id||p.identifiers.filter(v=>v.kind===item.kind&&v.value===item.value).length===1,'COMMAND_PAYLOAD_INVALID');}
  const indexPath='memberEnrollmentIndexes/'+id,index=get(indexPath);N(!index||index.schemaVersion==='member-enrollment-index-v1'&&index.personId===id&&index.activeEnrollmentId===null,'OPERATION_CONFLICT');
  const enrollment=E.create({id:r.operationId,personId:id,registration,approval:{reference:a.reference,approvedBy:a.approvedBy,approvedOn:a.approvedOn,membershipType:a.membershipType},actorUid:r.ownerUid,now,expiresAt:now+policy.invitationLifetimeHours*3600000});
  enrollment.profile=E.profile(p.profile,today);N(!get('memberEnrollments/'+r.operationId),'OPERATION_CONFLICT');add('memberEnrollments/'+r.operationId,enrollment);
  add(indexPath,{schemaVersion:'member-enrollment-index-v1',personId:id,revision:(index?.revision||0)+1,activeEnrollmentId:r.operationId});
  const overlay=Object.fromEntries(changes.map(v=>[v.path,v.data])),paths=[...new Set([...(h?.paths||[]),...changes.filter(v=>/^member(Identities|Identifiers)\//.test(v.path)).map(v=>v.path)])].sort();
  const future=path=>path==='memberAggregateHeads/'+id?{schemaVersion:'m11-head-v1',id,kind:'PERSON',revision:Math.max(1,expected+1),paths}:Object.hasOwn(overlay,path)?overlay[path]:get(path);
  return {changes,target:f4MemberAuthz_().load(future,id,control),effect:'PRE_REGISTRATION',effectiveDate:today,expected};
 }
 function change(c,id){
  const d=c.data;N(C.rev(c.expectedRevision)&&d.revision===c.expectedRevision+1&&d.personId===id,'REVISION_CONFLICT');
  if(c.path.startsWith('memberEnrollments/')){
   E.validate(d);N(c.path==='memberEnrollments/'+d.id,'COMMAND_PAYLOAD_INVALID');
   if(c.expectedRevision===0)N(d.status==='PRE_REGISTERED'&&d.revision===1&&d.uid===null&&d.invitedAt===null&&d.respondedAt===null&&d.completedAt===null&&d.normAcknowledgement===null&&d.imageManifestation===null,'COMMAND_PAYLOAD_INVALID');
  }else{C.exact(d,['schemaVersion','personId','revision','activeEnrollmentId']);N(c.path==='memberEnrollmentIndexes/'+id&&d.schemaVersion==='member-enrollment-index-v1'&&(d.activeEnrollmentId===null||C.op(d.activeEnrollmentId)),'COMMAND_PAYLOAD_INVALID');}
  return true;
 }
 return {validate,approval,proposal,change};
}
