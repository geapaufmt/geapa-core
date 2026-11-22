// Technical DEV access is a separate, protected UID record. It is never a role,
// institutional grant, membership, or portable production authorization.
function f4DevAdmin_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),K=f4MemberAuthzContracts_();
 const project='geapa-convergencia-dev',path=uid=>'coreDevAdmins/'+uid;
 function profile(get,uid,now,required=false){
  N(C.id(uid)&&Number.isSafeInteger(now),'AUTH_BINDING_INVALID');
  const p=get(path(uid));if(!p){N(!required,'AUTHORIZATION_DENIED');return null;}
  // Runtime identity and the immutable DEV destination must agree even when a
  // copied record claims DEV. No client field can select this environment.
  f4DevBoundary_().check();
  N(f4DevBoundary_().permit({environment:'DEV',projectId:project,database:'(default)'}),'DESTINATION_INVALID');
  C.exact(p,['schemaVersion','projectId','environment','uid','active','revision','issuedAt','expiresAt','permissions','exportable']);
  N(p.schemaVersion==='core-dev-admin-v1'&&p.projectId===project&&p.environment==='DEV'&&p.uid===uid&&p.exportable===false&&p.permissions==='FUNCTIONAL_ALL'&&C.rev(p.revision)&&typeof p.active==='boolean','AUTH_BINDING_INVALID');
  const live=p.active&&C.ms(p.issuedAt)<=now&&now<C.ms(p.expiresAt);
  N(C.ms(p.issuedAt)<C.ms(p.expiresAt),'AUTH_BINDING_INVALID');
  if(!live){N(!required,'ACCESS_REVOKED');return null;}return p;
 }
 function identify(get,uid,now){
  const p=profile(get,uid,now,true),control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS'));
  N(control.mode==='CANONICAL'&&control.activeWriterVersion==='M13','ACCESS_REVOKED');
  const binding=get('memberAuthBindings/'+uid);f4MemberModel_(true).validate('memberAuthBindings',binding);
  N(binding.uid===uid&&binding.active&&binding.domainVersion==='m11'&&binding.sourceRevision===binding.revision,'AUTH_BINDING_INVALID');
  const actor=f4MemberAuthz_().load(get,binding.idPessoa,control),person=actor.d['memberIdentities/'+binding.idPessoa];
  N(K.digest(get('portalUsers/'+uid))===K.digest(f4MemberRepository_().bindingView(binding,person)),'AUTH_BINDING_INVALID');
  const subject=K.subject(get('memberAuthzSubjects/'+binding.idPessoa)),state=get('memberCurrentStates/'+binding.idPessoa);
  N(state?.idPessoa===binding.idPessoa&&state.sourceRevision===actor.h.revision,'AUTH_BINDING_INVALID');
  return {profile:p,control,binding,actor,person,subject,state,derived:actor.M.derive(actor.d,binding.idPessoa,actor.h.revision,now)};
 }
 function expected(a){return {authorityEpoch:a.control.epoch,policyEpoch:a.control.policyEpoch,readAuthzEpoch:a.control.readAuthzEpoch,subjectEpoch:a.subject.authzEpoch,bindingVersion:a.binding.bindingVersion,personRevision:a.actor.h.revision,authorizationRevision:a.profile.revision};}
 function authorize(get,q,now){
  N(q.kind==='DEV_ADMIN','AUTH_BINDING_INVALID');
  f4ActivitiesAuthzDomain_().contracts().request(q);
  const a=identify(get,q.uid,now),e=expected(a);
  N(q.sourceId===q.uid&&K.digest(q.expected)===K.digest(e)&&q.target.personId===a.binding.idPessoa,'AUTH_BINDING_INVALID');
  const end=Math.min(C.ms(q.tokenExpiresAt),C.ms(q.admissionExpiresAt),C.ms(a.profile.expiresAt));N(now<end,'ACCESS_REVOKED');
  const proof={schemaVersion:'act-authz-proof-v13',kind:'DEV_ADMIN',uid:q.uid,idPessoa:a.binding.idPessoa,capability:q.capability,...e,stateRevision:a.state.sourceRevision,targetsDigest:K.digest(q.target),evidenceRevisions:[],projectId:project,profilePath:path(q.uid),profileRevision:a.profile.revision,issuedAt:F.Timestamp.fromMillis(now),expiresAt:F.Timestamp.fromMillis(end)};
  return {...a,proof,request:q,target:a.actor};
 }
 function recheck(get,q,proof,now){
  N(proof?.kind==='DEV_ADMIN'&&C.ms(proof.issuedAt)<=now&&now<C.ms(proof.expiresAt),'ACCESS_REVOKED');
  const next=authorize(get,q,now).proof,stable=p=>Object.fromEntries(Object.entries(p).filter(([k])=>!['issuedAt','expiresAt'].includes(k)));
  N(K.digest(stable(next))===K.digest(stable(proof))&&C.ms(next.expiresAt)>=C.ms(proof.expiresAt),'ACCESS_REVOKED');return next;
 }
 return {profile,identify,expected,authorize,recheck,path};
}

// Named technical grants are independent of FUNCTIONAL_ALL and institutional
// roles. Administrative tooling owns this document; browsers cannot issue it.
function f4DevAdministrativeAccess_(){
 const C=f4Core_(),K=f4MemberAuthzContracts_(),N=K.need;
 const path=uid=>'coreDevAdministrativeGrants/'+uid;
 const fields={PEOPLE_DIRECTORY_READ:['idPessoa','displayLabel','state','categories','functions','relations','history'],MEMBERS_PROFILE_CORRECT:['displayName','cpf','rga','birthDate','countryCode','courseId'],INSTITUTIONAL_DIRECTORY_READ:['summary','fields','provenance','issues','canonical'],INSTITUTIONAL_GOVERNANCE_DRAFT:['draft'],MEMBERS_INGRESS_EXECUTE:['preRegister']};
 function source(get,identity,now,cap){
  // Check the server destination BEFORE reading any copied DEV grant in PROD.
  if(f4ExecutionEnvironment_().production||!f4DevBoundary_().permit({environment:'DEV',projectId:'geapa-convergencia-dev',database:'(default)'}))return null;
  const g=get(path(identity.uid));if(!g)return null;
  C.exact(g,['schemaVersion','id','uid','personId','bindingVersion','subjectEpoch','projectId','environment','exportable','active','revision','issuedAt','expiresAt','approvalRef','capabilities','fieldMasks','syntheticPersonIds','sections',...(Object.hasOwn(g,'admissionApprovalIds')?['admissionApprovalIds']:[])]);
  if(g.admissionApprovalIds)N(Array.isArray(g.admissionApprovalIds)&&g.admissionApprovalIds.length<=10&&g.admissionApprovalIds.every(C.id)&&new Set(g.admissionApprovalIds).size===g.admissionApprovalIds.length,'AUTHORIZATION_DENIED');
  N(g.schemaVersion==='core-dev-administrative-grant-v1'&&g.id===identity.uid&&g.uid===identity.uid&&g.projectId==='geapa-convergencia-dev'&&g.environment==='DEV'&&g.exportable===false&&typeof g.active==='boolean'&&C.rev(g.revision)&&g.revision>0&&C.id(g.personId)&&C.id(g.approvalRef),'AUTHORIZATION_DENIED');
  N(Array.isArray(g.capabilities)&&g.capabilities.length>0&&new Set(g.capabilities).size===g.capabilities.length&&g.capabilities.every(c=>Object.hasOwn(fields,c)),'AUTHORIZATION_DENIED');
  C.exact(g.fieldMasks,g.capabilities);for(const c of g.capabilities)N(Array.isArray(g.fieldMasks[c])&&g.fieldMasks[c].length>0&&new Set(g.fieldMasks[c]).size===g.fieldMasks[c].length&&g.fieldMasks[c].every(f=>fields[c].includes(f)),'AUTHORIZATION_DENIED');
  for(const [xs,valid,max]of [[g.syntheticPersonIds,C.id,100],[g.sections,s=>['identity','parameters','norms','governance'].includes(s),4]])N(Array.isArray(xs)&&xs.length<=max&&new Set(xs).size===xs.length&&xs.every(valid),'AUTHORIZATION_DENIED');
  const start=C.ms(g.issuedAt),end=C.ms(g.expiresAt);N(start<end&&end-start<=30*86400000,'AUTHORIZATION_DENIED');
  if(!g.active||now<start||now>=end||!g.capabilities.includes(cap))return null;
  N(now<identity.expiresAt,'ACCESS_REVOKED');
  const actor=f4MemberAuthz_().identify(get,identity.uid,now);
  N(actor.binding.idPessoa===g.personId&&actor.binding.bindingVersion===g.bindingVersion&&actor.subject.authzEpoch===g.subjectEpoch,'ACCESS_REVOKED');
  return {actor,source:g,policies:[g],mask:g.fieldMasks[cap],technical:true,validUntil:Math.min(end,identity.expiresAt,now+60000,C.ms(actor.state.validUntil),C.ms(actor.derived.state.validUntil))};
 }
 function institutional(get,identity,now,cap,resource){const a=source(get,identity,now,cap);if(!a)return null;
  N(cap==='INSTITUTIONAL_DIRECTORY_READ'?a.source.sections.some(s=>resource==='institutionalDirectory/'+s):cap==='INSTITUTIONAL_GOVERNANCE_DRAFT'&&resource==='institutionalGovernanceDrafts/current','AUTHORIZATION_DENIED');return a;
 }
 function correction(get,identity,now,target){const a=source(get,identity,now,'MEMBERS_PROFILE_CORRECT');N(a&&a.source.syntheticPersonIds.includes(target.personId),'AUTHORIZATION_DENIED');
  N(target.scopeKey==='ADMINISTRACAO'&&target.purpose==='PROFILE_CORRECTION'&&target.effect==='PROFILE_CORRECTION'&&target.fieldMask.every(f=>a.mask.includes(f)),'AUTHORIZATION_DENIED');
  f4MemberOperationalScope_(get,{capability:'MEMBERS_PROFILE_CORRECT',target},{ownPerson:false,allowedPersons:a.source.syntheticPersonIds,allowedTargetScopes:['ADMINISTRACAO'],allowedBoards:[],allowedRoles:[],allowedResources:[],commandScopes:['PROFILE_CORRECTION']},a.actor.binding.idPessoa);return a;
 }
 function command(get,identity,now,cap,target,r){
  if(cap==='MEMBERS_PROFILE_CORRECT')return correction(get,identity,now,target);
  N(cap==='MEMBERS_INGRESS_EXECUTE'&&r?.commandType==='members.ingress.execute'&&r.payload.operation==='PRE_REGISTER','AUTHORIZATION_DENIED');
  const a=source(get,identity,now,cap);N(a&&a.mask.includes('preRegister')&&a.source.admissionApprovalIds?.includes(r.payload.approvalRef)&&a.source.syntheticPersonIds.includes(target.personId),'AUTHORIZATION_DENIED');
  const approval=f4MemberPreRegistration_().approval(get,r.payload.approvalRef,r.payload.registration,now);
  N(approval.devTechnicalOwnerUid===identity.uid&&approval.synthetic===true&&approval.devPersonId===target.personId&&C.hash(approval.devRegistration)===C.hash(r.payload.registration)&&Object.keys(r.payload.profile).length===0,'AUTHORIZATION_DENIED');
  N(target.purpose==='ADMINISTRATIVE_INGRESS'&&target.effect==='PRE_REGISTRATION'&&target.fieldMask.length===0,'AUTHORIZATION_DENIED');
  f4MemberOperationalScope_(get,{kind:'DEV_TECHNICAL',capability:cap,target},{ownPerson:false,allowedPersons:a.source.syntheticPersonIds,allowedTargetScopes:['ADMINISTRACAO'],allowedBoards:[],allowedRoles:[],allowedResources:[],commandScopes:['ADMINISTRATIVE_INGRESS']},a.actor.binding.idPessoa,r,now);return a;
 }
 function governanceInput(auth,input,old){if(!auth.technical)return;
  N(!input.action&&!input.minutes&&!input.documentaryEvidence&&!input.attachments?.length&&input.sourceRecordId===null&&input.references?.length===0&&input.documentaryState==='DRAFT'&&(!old||old.devTechnicalOwnerUid===auth.source.uid&&old.documentaryState==='DRAFT'&&!old.kind),'AUTHORIZATION_DENIED');
 }
 // A stale browser journal may refer to a completed pre-migration command.
 // Read only its canonical terminal state; never claim it or change executor.
 function terminalRegistration(get,identity,now,key){
  const a=source(get,identity,now,'MEMBERS_INGRESS_EXECUTE');N(a&&a.mask.includes('preRegister'),'AUTHORIZATION_DENIED');
  f4SharedCommandCatalog_().selector(key);N(key.commandType==='members.ingress.execute','AUTHORIZATION_DENIED');
  const result=f4MemberLedgerM13_().recover(get,key,identity.uid,now);
  N(result&&['COMPLETED','FAILED'].includes(result.state),'RECONCILIATION_REQUIRED');
  return {operationId:key.operationId,receipt:{state:result.state,code:result.code}};
 }
 // Recover the exact already committed synthetic registration before a new
 // preparation. This grants no invitation, candidate or activation authority.
 function registrationStatus(get,identity,now,input){
  const a=source(get,identity,now,'MEMBERS_INGRESS_EXECUTE');if(!a)return null;
  C.exact(input,['operationId','route','payload']);N(C.op(input.operationId)&&input.route==='/admin/ingressos-membros/cadastrar','COMMAND_PAYLOAD_INVALID');
  const p=input.payload;C.exact(p,['nomeCompleto','emailPrincipal','rga','telefone','categoriaIngresso','referenciaAprovacao',...(Object.hasOwn(p,'dadosComplementares')?['dadosComplementares']:[])]);
  const registration=f4MemberEnrollmentContract_().quick({fullName:p.nomeCompleto,email:p.emailPrincipal,rga:p.rga,phone:p.telefone});
  const approval=f4MemberPreRegistration_().locate(get,registration,p.referenciaAprovacao,now);
  N(approval.reference===p.referenciaAprovacao&&approval.membershipType===p.categoriaIngresso,'ADMISSION_APPROVAL_REQUIRED');
  N(a.mask.includes('preRegister')&&a.source.admissionApprovalIds?.includes(approval.id)&&approval.synthetic===true&&approval.devTechnicalOwnerUid===identity.uid&&a.source.syntheticPersonIds.includes(approval.devPersonId)&&C.hash(approval.devRegistration)===C.hash(registration)&&Object.keys(p.dadosComplementares||{}).length===0,'AUTHORIZATION_DENIED');
  const resolved=f4MemberRegistrationIdentity_().resolve(get,{email:registration.email,rga:registration.rga});
  N(resolved.personId===approval.devPersonId,'AUTHORIZATION_DENIED');
  const index=get('memberEnrollmentIndexes/'+resolved.personId);if(!index?.activeEnrollmentId)return null;
  N(index.schemaVersion==='member-enrollment-index-v1'&&index.personId===resolved.personId&&C.op(index.activeEnrollmentId),'RECONCILIATION_REQUIRED');
  const enrollment=f4MemberEnrollmentContract_().validate(get('memberEnrollments/'+index.activeEnrollmentId)),r=get('commandRequests/'+index.activeEnrollmentId);
  N(enrollment.personId===resolved.personId&&C.hash(enrollment.registration)===C.hash(registration)&&enrollment.approval.reference===approval.reference&&enrollment.approval.membershipType===approval.membershipType,'OPERATION_CONFLICT');
  N(r?.ownerUid===identity.uid&&r.commandType==='members.ingress.execute'&&r.target.personId===resolved.personId&&r.payload.operation==='PRE_REGISTER'&&r.payload.approvalRef===approval.id&&C.hash(r.payload.registration)===C.hash(registration)&&Object.keys(r.payload.profile).length===0,'OPERATION_CONFLICT');
  const result=f4MemberLedgerM13_().recover(get,{operationId:r.operationId,commandType:r.commandType,digest:r.digest},identity.uid,now);
  N(result?.state==='COMPLETED','RECONCILIATION_REQUIRED');
  return {operationId:r.operationId,receipt:{state:result.state,code:result.code}};
 }
 return {source,institutional,correction,command,governanceInput,terminalRegistration,registrationStatus,path,fields};
}
