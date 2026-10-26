// Candidate actions extend the existing onboarding command. The lifecycle
// remains separate from the administrative C/J orchestration already in use.
function f4MemberEnrollmentCommands_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),E=f4MemberEnrollmentContract_(),A=f4MemberEnrollmentAccess_(),U=f4MemberFamilyInputs_(),K=f4MemberCommandCatalog_();
 const is=r=>r?.commandType==='members.onboarding.complete'&&r.payload?.operation==='ENROLLMENT';
 function validate(r){
  N(is(r),'COMMAND_TYPE_INVALID');C.exact(r.target,['personId','enrollmentId']);N(C.id(r.target.personId)&&C.op(r.target.enrollmentId)&&r.evidenceRefs.length===0&&r.decisionRef===null,'COMMAND_PAYLOAD_INVALID');
  const p=r.payload;N(['ACCEPT','DECLINE','SAVE','ACKNOWLEDGE','IMAGE','COMPLETE'].includes(p.action),'COMMAND_PAYLOAD_INVALID');
  C.exact(p,['operation','action',...(p.action==='SAVE'?['profile']:p.action==='ACKNOWLEDGE'?['accepted','termVersion','documents']:p.action==='IMAGE'?['choice','termVersion']:[])]);
  if(p.action==='SAVE')E.profile(p.profile,'9999-12-31');
  if(p.action==='ACKNOWLEDGE'){N(p.accepted===true&&C.id(p.termVersion)&&Array.isArray(p.documents)&&p.documents.length>0&&p.documents.length<=20,'COMMAND_PAYLOAD_INVALID');for(const d of p.documents){C.exact(d,['id','version']);N(C.id(d.id)&&C.id(d.version),'COMMAND_PAYLOAD_INVALID');}N(new Set(p.documents.map(d=>d.id)).size===p.documents.length,'COMMAND_PAYLOAD_INVALID');}
  if(p.action==='IMAGE')N(['GRANTED','DECLINED','REVOKED'].includes(p.choice)&&C.id(p.termVersion),'COMMAND_PAYLOAD_INVALID');return true;
 }
 function catalog(get,now){
  const c=get('memberOperationalConfig/enrollmentTerms');N(c?.schemaVersion==='member-enrollment-terms-v1'&&c.enabled===true&&C.rev(c.revision)&&c.revision>0&&C.id(c.normTermVersion)&&Array.isArray(c.normVersionIds)&&c.normVersionIds.length>0&&c.normVersionIds.length<=20&&c.normVersionIds.every(C.id),'ACCESS_REVOKED');
  N(new Set(c.normVersionIds).size===c.normVersionIds.length,'SCHEMA_INVALID');
  const norms=c.normVersionIds.map(id=>{
   const v=get('institutionalNormVersions/'+id);N(v?.schemaVersion==='institutional-norm-version-v1'&&v.versionId===id&&v.status==='CURRENT'&&v.appliesToEnrollment===true&&C.id(v.approvalRef),'ACCESS_REVOKED');
   const a=get('institutionalNormApprovals/'+v.approvalRef);N(a?.schemaVersion==='institutional-norm-approval-v1'&&a.active===true&&a.normId===v.id&&a.version===v.version&&a.approvedOn===v.approvedOn&&a.effectiveFrom===v.effectiveFrom&&a.immutableVersionRef===v.immutableVersionRef&&C.id(a.formalReference),'ACCESS_REVOKED');
   N(C.civilStart(v.approvedOn)<=now&&C.civilStart(v.effectiveFrom)<=now,'ACCESS_REVOKED');return v;
  });
  // This also rejects drafts and validates the immutable version references.
  E.normReceipt({personId:'VALIDATE',enrollmentId:'0'.repeat(32),uid:'VALIDATE',catalog:norms,accepted:true,now,termVersion:c.normTermVersion});
  return {configuration:c,norms};
 }
 function proposal(get,r,now){
  validate(r);const s=A.source(get,r.ownerUid,now),old=s.enrollment,p=r.payload,id=old.personId,path='memberEnrollments/'+old.id,expected=U.revision(r,'memberAggregateHeads/'+id),revision=U.revision(r,path);
  N(r.target.personId===id&&r.target.enrollmentId===old.id&&expected===s.head.revision&&revision===old.revision,'REVISION_CONFLICT');
  const verified={uid:r.ownerUid,projectId:s.session.projectId,emailVerified:true,email:old.registration.email,expiresAt:C.ms(s.session.expiresAt)},today=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd');
  let next,receipt=null;
  if(['ACCEPT','DECLINE'].includes(p.action))next=E.response(old,revision,p.action,verified,now,s.session.projectId);
  else{
   E.identity(old,verified,now,s.session.projectId);N(old.uid===r.ownerUid&&old.status==='ACCEPTED','ACCESS_REVOKED');
   if(p.action==='SAVE'){
    for(const id of p.profile.interestAxisIds||[]){const axis=get('presentationPolicyReferences/'+id);N(axis?.id===id&&axis.approved===true&&axis.active===true,'COMMAND_PAYLOAD_INVALID');}
    next=E.save(old,revision,p.profile,verified,now,s.session.projectId,today);
   }
   else if(p.action==='ACKNOWLEDGE'){
    const c=catalog(get,now),shown=c.norms.map(v=>({id:v.id,version:v.version})).sort((a,b)=>a.id.localeCompare(b.id));N(p.termVersion===c.configuration.normTermVersion&&K.equal(p.documents.slice().sort((a,b)=>a.id.localeCompare(b.id)),shown),'REVISION_CONFLICT');
    receipt=E.normReceipt({personId:id,enrollmentId:old.id,uid:r.ownerUid,catalog:c.norms,accepted:p.accepted,now,termVersion:p.termVersion});next={...old,normAcknowledgement:receipt,revision:revision+1,updatedAt:now};
   }else if(p.action==='IMAGE'){
    const c=get('memberOperationalConfig/enrollmentTerms');N(c?.schemaVersion==='member-enrollment-terms-v1'&&c.enabled===true&&c.imageTerm?.version===p.termVersion&&c.imageTerm.optional===true&&c.imageTerm.activationConsequence==='NONE','ACCESS_REVOKED');
    receipt=E.imageReceipt({personId:id,enrollmentId:old.id,uid:r.ownerUid,choice:p.choice,term:c.imageTerm,now,previous:old.imageManifestation});next={...old,imageManifestation:receipt,revision:revision+1,updatedAt:now};
   }else return completion(get,r,now,s,catalog(get,now));
  }
  E.validate(next);const changes=[{path,expectedRevision:revision,data:next}];
  if(p.action==='DECLINE'){const ip='memberEnrollmentIndexes/'+id,index=get(ip);N(index?.activeEnrollmentId===old.id&&index.personId===id&&U.revision(r,ip)===index.revision,'REVISION_CONFLICT');changes.push({path:ip,expectedRevision:index.revision,data:{...index,revision:index.revision+1,activeEnrollmentId:null}});}
  if(receipt){const fp='memberEnrollmentFacts/'+r.operationId;N(!get(fp)&&U.revision(r,fp)===0,'REVISION_CONFLICT');changes.push({path:fp,expectedRevision:0,data:{schemaVersion:'member-enrollment-fact-v1',id:r.operationId,personId:id,enrollmentId:old.id,revision:1,kind:p.action,receipt}});}
  return {id,expected,changes};
 }
 function completion(get,r,now,s,catalog){
  const old=s.enrollment,id=old.personId,today=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'),verified={uid:r.ownerUid,projectId:s.session.projectId,emailVerified:true,email:old.registration.email,expiresAt:C.ms(s.session.expiresAt)};
  const facts=E.ready(get,old,U.revision(r,'memberEnrollments/'+old.id),verified,now,s.session.projectId,today,catalog.norms);
  N(facts.normAcknowledgement.termVersion===catalog.configuration.normTermVersion,'REVISION_CONFLICT');
  const approvalRef='AP_'+C.hash({email:old.registration.email,rga:old.registration.rga}).slice(0,40),approved=f4MemberPreRegistration_().approval(get,approvalRef,old.registration,now);
  N(K.equal(old.approval,{reference:approved.reference,approvedBy:approved.approvedBy,approvedOn:approved.approvedOn,membershipType:approved.membershipType}),'REVISION_CONFLICT');
  // The canonical prior approval determines the type. Candidate input cannot
  // promote an incoming member or substitute the institutional decision.
  N(['MEMBRO_INGRESSANTE','MEMBRO_EFETIVO'].includes(facts.membershipType),'DOMAIN_STATE_INVALID');
  const loaded=f4MemberAuthz_().load(get,id,s.control),M=loaded.M,derived=M.derive(loaded.d,id,loaded.h.revision,now),changes=[];
  N(['SEM_VINCULO','DESLIGADO','EGRESSO','MEMBRO_EM_ESPERA'].includes(derived.state.state),'DOMAIN_STATE_INVALID');
  function add(path,data){const previous=get(path),revision=U.revision(r,path);N((previous?.revision||0)===revision&&data.revision===revision+1,'REVISION_CONFLICT');changes.push({path,expectedRevision:revision,data});}
  const entity=(col,key,fields)=>({schemaVersion:'m1-'+col+'-v1',id:key,revision:1,origin:'M13_COMMAND',...fields});
  const cpfs=M.rows(loaded.d,'memberIdentifiers').filter(v=>v.kind==='CPF'&&v.active);N(cpfs.length<=1&&cpfs.every(v=>v.value===facts.profile.cpf),'RECONCILIATION_REQUIRED');
  if(!cpfs.length){const key='I_CPF_'+r.operationId;N(!get('memberIdentifiers/'+key),'REVISION_CONFLICT');add('memberIdentifiers/'+key,entity('memberIdentifiers',key,{idPessoa:id,kind:'CPF',value:facts.profile.cpf,active:true}));}
  const profilePath='memberPrivateProfiles/'+id,previousProfile=get(profilePath),profile={...(previousProfile||{id,idPessoa:id,origin:'M13_COMMAND',academicReference:''}),schemaVersion:'m1-memberPrivateProfiles-v2',revision:(previousProfile?.revision||0)+1,contactEmail:old.registration.email,phone:old.registration.phone};
  for(const [key,value]of Object.entries(facts.profile))if(key!=='cpf')profile[key==='displayName'?'preferredName':key]=value;
  f4MemberProfileContract_().validate(profile);add(profilePath,profile);
  const linkId='V_'+old.id,eventId='EVT_'+r.operationId;
  N(!get('memberLinks/'+linkId)&&!get('memberEvents/'+eventId),'REVISION_CONFLICT');
  const event=(key,link,type,effect,date,extra={})=>entity('memberEvents',key,{idPessoa:id,idVinculo:link,type,effect,effectiveDate:date,endDate:null,supersedes:null,correctionPath:null,prior:null,...extra});
  const at=C.civilStart(today),overlap=M.rows(loaded.d,'memberLinks').filter(v=>{const interval=M.interval(v.startDate,v.endDate);return interval.a<=at&&at<interval.z;});N(overlap.length<=1,'RECONCILIATION_REQUIRED');
  if(overlap.length){const prior=overlap[0];N(prior.startDate<today&&derived.currentFunctions.length===0&&(['MEMBRO_EM_ESPERA','EGRESSO'].includes(prior.type)||derived.state.state==='DESLIGADO'),'RECONCILIATION_REQUIRED');
   const end=new Date(Date.parse(today+'T12:00:00Z')-86400000).toISOString().slice(0,10),correctionId='COR_'+r.operationId;N(!get('memberEvents/'+correctionId),'REVISION_CONFLICT');
   add('memberLinks/'+prior.id,{...prior,revision:prior.revision+1,endDate:end});add('memberEvents/'+correctionId,event(correctionId,prior.id,'RETIFICACAO',null,today,{correctionPath:'memberLinks/'+prior.id,prior}));
  }
  add('memberLinks/'+linkId,entity('memberLinks',linkId,{idPessoa:id,type:facts.membershipType,startDate:today,endDate:null,exitDate:null}));
  const entryEffect=facts.membershipType==='MEMBRO_EFETIVO'?'EFETIVACAO':'INGRESSO';
  add('memberEvents/'+eventId,event(eventId,linkId,entryEffect,entryEffect,today));
  const bindingPath='memberAuthBindings/'+r.ownerUid,binding=get(bindingPath);
  N(!binding||binding.active===true&&binding.idPessoa===id&&binding.uid===r.ownerUid,'AUTH_BINDING_INVALID');
  if(!binding)add(bindingPath,{schemaVersion:'binding-v1',id:r.ownerUid,revision:1,origin:'M13_COMMAND',uid:r.ownerUid,idPessoa:id,active:true,bindingVersion:1,policyVersion:1,sourceRevision:1,createdAt:now,updatedAt:now,domainVersion:'m11'});
  const subjectPath='memberAuthzSubjects/'+id,subject=get(subjectPath);N(!subject||subject.active===true&&subject.idPessoa===id,'ACCESS_REVOKED');
  if(!subject)add(subjectPath,{schemaVersion:'member-authz-subject-v13',idPessoa:id,revision:1,authzEpoch:1,active:true,approvalRef:old.approval.reference,updatedAt:F.Timestamp.fromMillis(now)});
  const next={...old,status:'COMPLETED',completedAt:now,updatedAt:now,revision:old.revision+1};E.validate(next);add('memberEnrollments/'+old.id,next);
  const ip='memberEnrollmentIndexes/'+id,index=get(ip);N(index?.personId===id&&index.activeEnrollmentId===old.id,'AUTH_BINDING_INVALID');add(ip,{...index,revision:index.revision+1,activeEnrollmentId:null});
  return {id,expected:s.head.revision,changes};
 }
 function resolve(get,r,now){
  const p=proposal(get,r,now),saved=get('commandRequests/'+r.operationId),prov=get('commandProvenance/'+r.operationId),ticket=get('memberAdmissionTickets/'+r.ticketId);
  N(K.equal(saved,r)&&prov?.ownerUid===r.ownerUid&&prov.operationId===r.operationId&&prov.ticketId===r.ticketId&&prov.digest===r.digest&&ticket?.ownerUid===r.ownerUid&&ticket.digest===r.digest&&K.equal(ticket.requestEnvelope,K.template(r))&&prov.authRequests.length===1&&ticket.authProofs.length===1,'AUTH_BINDING_INVALID');
  const q=prov.authRequests[0],proof=ticket.authProofs[0];N(q.uid===r.ownerUid&&q.target.personId===r.target.personId&&q.target.enrollmentId===r.target.enrollmentId&&q.target.action===r.payload.action,'AUTH_BINDING_INVALID');A.recheck(get,q,proof,now);
  return {...p,now,boardExpected:{},mode:'M13',authorization:{request:q,proof,command:r},deadline:C.ms(proof.expiresAt)};
 }
 function change(c,id){
  const d=c.data;N(c.path==='memberEnrollmentFacts/'+d.id&&C.op(d.id)&&d.schemaVersion==='member-enrollment-fact-v1'&&d.personId===id&&C.op(d.enrollmentId)&&c.expectedRevision===0&&d.revision===1&&['ACKNOWLEDGE','IMAGE'].includes(d.kind),'COMMAND_PAYLOAD_INVALID');
  N(d.receipt?.personId===id&&d.receipt.enrollmentId===d.enrollmentId&&C.id(d.receipt.uid),'COMMAND_PAYLOAD_INVALID');return true;
 }
 function paths(get,r,now){
  validate(r);const id=r.target.personId,action=r.payload.action,result=['memberAggregateHeads/'+id,'memberEnrollments/'+r.target.enrollmentId];
  if(['DECLINE','COMPLETE'].includes(action))result.push('memberEnrollmentIndexes/'+id);
  if(['ACKNOWLEDGE','IMAGE'].includes(action))result.push('memberEnrollmentFacts/'+r.operationId);
  if(action==='COMPLETE'){
   const s=A.source(get,r.ownerUid,now);N(id===s.enrollment.personId&&r.target.enrollmentId===s.enrollment.id,'AUTH_BINDING_INVALID');
   const loaded=f4MemberAuthz_().load(get,id,s.control),M=loaded.M,today=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'),at=C.civilStart(today);
   result.push('memberPrivateProfiles/'+id,'memberLinks/V_'+s.enrollment.id,'memberEvents/EVT_'+r.operationId);
   if(!M.rows(loaded.d,'memberIdentifiers').some(v=>v.kind==='CPF'&&v.active))result.push('memberIdentifiers/I_CPF_'+r.operationId);
   if(!get('memberAuthBindings/'+r.ownerUid))result.push('memberAuthBindings/'+r.ownerUid);
   if(!get('memberAuthzSubjects/'+id))result.push('memberAuthzSubjects/'+id);
   for(const v of M.rows(loaded.d,'memberLinks')){const iv=M.interval(v.startDate,v.endDate);if(iv.a<=at&&at<iv.z)result.push('memberLinks/'+v.id,'memberEvents/COR_'+r.operationId);}
  }
  return [...new Set(result)].sort();
 }
 return {is,validate,catalog,proposal,resolve,change,paths};
}
