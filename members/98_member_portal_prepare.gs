// Protected preparation for the existing command transport. IDs, revisions,
// evidence and selectors are produced here; the Portal gets a command intent.
function f4MemberPortalPrepare_({config,identity,clock=Date.now}){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),K=f4MemberCommandCatalog_(),db=F.connect(config),T=F.Timestamp.fromMillis;
 const allowed=['members.self.update','members.profile.correct','members.ingress.execute','members.voluntary.submit','members.voluntary.review','members.membership.voluntaryExit'];
 const material=r=>Object.fromEntries(['commandType','target','payload','evidenceRefs','decisionRef'].map(k=>[k,r[k]]));
 function current(get,now){return f4DevAdmin_().profile(get,identity.uid,now)?f4DevAdmin_().identify(get,identity.uid,now):f4MemberAuthz_().identify(get,identity.uid,now);}
 function configuration(get){const s=get('memberOperationalConfig/portal');N(s?.schemaVersion==='member-operational-portal-v1'&&s.enabled===true&&C.id(s.approvalRef)&&Array.isArray(s.commands)&&s.commands.every(c=>allowed.includes(c))&&Array.isArray(s.policyIds)&&s.policyIds.length<=16&&s.policyIds.every(C.id),'ACCESS_REVOKED');return s;}
 function requestFor(get,r,resources,cap,now,end,settings){
  const technical=f4DevAdmin_().profile(get,identity.uid,now),a=technical?f4DevAdmin_().identify(get,identity.uid,now):f4MemberAuthz_().identify(get,identity.uid,now),self=['MEMBERS_SELF_UPDATE','MEMBERS_VOLUNTARY_REQUEST'].includes(cap);
  N(!self||r.target.personId===a.binding.idPessoa,'ACCESS_REVOKED');
  let source=technical,kind=technical?'DEV_ADMIN':null,functionId=null,actorScope=null;
  if(!technical){for(const id of settings.policyIds){const p=get('memberAuthzPolicies/'+id);if(!p?.active||!p.capabilities?.includes(cap))continue;f4MemberAuthzContracts_().policy(p);if(self&&p.proofKind==='SELF'){source=p;kind='SELF';break;}if(!self&&p.proofKind==='ROLE'){const f=a.derived.currentFunctions.find(f=>f.roleKey===p.actorSelector.roleKey&&f.scopeKey===p.actorSelector.actorScope);if(f){source=p;kind='ROLE';functionId=f.idVigencia;actorScope=f.scopeKey;break;}}}}
  N(source&&kind,'ACCESS_REVOKED');
  const target={personId:r.target.personId,scopeKey:self?null:cap==='MEMBERS_INGRESS_EXECUTE'?'ADMINISTRACAO':'RECURSOS_HUMANOS',boardId:null,roleKey:null,resources,caseId:self?null:r.target.requestId||'ADM_'+r.operationId,purpose:self?null:cap==='MEMBERS_INGRESS_EXECUTE'?'ADMINISTRATIVE_INGRESS':cap==='MEMBERS_VOLUNTARY_SUSPENSION_EXECUTE'?'VOLUNTARY_SUSPENSION':'VOLUNTARY_EXIT',fieldMask:[],linkId:self?null:r.target.linkId,effect:self?null:cap==='MEMBERS_INGRESS_EXECUTE'?(r.payload.operation==='PRE_REGISTER'?'PRE_REGISTRATION':r.payload.operation==='MANAGE_INVITATION'?'INVITATION':'INGRESSO'):cap==='MEMBERS_VOLUNTARY_SUSPENSION_EXECUTE'?(r.payload.action==='RETURN'?'RETORNO':'SUSPENSAO'):'DESLIGAMENTO',effectiveDate:self?null:cap==='MEMBERS_VOLUNTARY_SUSPENSION_EXECUTE'?suspensionDate(get,r):r.payload.startDate||r.payload.effectiveDate||Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd')};
  if(['MEMBERS_SELF_UPDATE','MEMBERS_PROFILE_CORRECT'].includes(cap)){target.fieldMask=f4MemberProfileCorrections_().is(r)?(self?[]:[r.payload.field]):Object.keys(r.payload.patch||r.payload.corrections).sort();target.linkId=null;if(!self){target.scopeKey='ADMINISTRACAO';target.purpose='PROFILE_CORRECTION';target.effect='PROFILE_CORRECTION';if(f4MemberProfileCorrections_().is(r))target.caseId=r.payload.requestId;}}
  return {schemaVersion:'member-authz-request-v13',uid:identity.uid,kind,sourceId:technical?identity.uid:source.id,functionId,actorScope,capability:cap,expected:{authorityEpoch:a.control.epoch,policyEpoch:a.control.policyEpoch,readAuthzEpoch:a.control.readAuthzEpoch,subjectEpoch:a.subject.authzEpoch,bindingVersion:a.binding.bindingVersion,personRevision:a.actor.h.revision,authorizationRevision:source.revision},tokenExpiresAt:T(identity.expiresAt),admissionExpiresAt:T(end),target,evidenceRefs:r.evidenceRefs};
 }
 function suspensionDate(get,r){const v=f4MemberVoluntarySuspension_().validate(get('memberVoluntaryRequests/'+r.target.requestId));return r.payload.action==='RETURN'?new Date(Date.parse(v.endDate+'T12:00:00Z')+86400000).toISOString().slice(0,10):v.startDate;}
 function prepare(input){
  C.exact(input,['operationId','route','payload']);N(C.op(input.operationId),'COMMAND_PAYLOAD_INVALID');
  return db.transaction(tx=>{
   const get=p=>tx.get(p).data(),settings=configuration(get),now=clock(),end=Math.min(now+3600000,identity.expiresAt),binding=get('memberAuthBindings/'+identity.uid);
   N(binding?.active&&now<end,'ACCESS_REVOKED');
   const intentPath='memberPreparedIntents/'+input.operationId,inputHash=C.hash(input),prior=get(intentPath);
   if(prior){N(prior.ownerUid===identity.uid&&prior.inputHash===inputHash,'OPERATION_CONFLICT');current(get,now);return prior.input;}
   const p=input.payload,op=input.operationId,r={schemaVersion:'geapa-members-command-v13',operationId:op,ownerUid:identity.uid,commandType:null,target:null,payload:null,evidenceRefs:[],decisionRef:null};
   let paths=[],resources=[],facts=null,requestId,cap;
   if(['/meu-perfil/correcoes/solicitar','/meu-perfil/correcoes/acao','/admin/correcoes-cadastrais/acao'].includes(input.route)){
    const W=f4MemberProfileCorrections_(),self=input.route.startsWith('/meu-perfil/'),submit=input.route.endsWith('/solicitar');
    C.exact(p,submit?['field','value','reason','personRevision']:['requestId','revision','action','reason']);
    let old=null;if(!submit){N(C.op(p.requestId)&&C.rev(p.revision),'COMMAND_PAYLOAD_INVALID');old=W.record(get('memberProfileCorrections/'+p.requestId));N(old.revision===p.revision,'REVISION_CONFLICT');}
    const personId=submit?binding.idPessoa:old.personId;N(!self||personId===binding.idPessoa,'ACCESS_REVOKED');
    if(submit)N(C.rev(p.personRevision)&&get('memberAggregateHeads/'+personId)?.revision===p.personRevision,'REVISION_CONFLICT');
    r.commandType=self?'members.self.update':'members.profile.correct';r.target={personId};r.payload={operation:'CORRECTION_REQUEST',action:submit?'SUBMIT':p.action,requestId:submit?op:p.requestId,field:submit?p.field:old.field,value:submit?p.value:null,reason:p.reason};
    cap=self?'MEMBERS_SELF_UPDATE':'MEMBERS_PROFILE_CORRECT';paths=W.paths(get,r);resources=paths.slice();if(!self){facts={};r.decisionRef='CORR_'+p.requestId;}
   }else if(input.route==='/meu-perfil/salvar'||input.route==='/admin/membros/corrigir-perfil'){
    const self=input.route==='/meu-perfil/salvar';C.exact(p,self?['patch','profileRevision','personRevision']:['personId','patch','profileRevision','personRevision','justification','documentReference']);
    const personId=self?binding.idPessoa:p.personId;N(C.id(personId)&&C.rev(p.profileRevision)&&C.rev(p.personRevision),'COMMAND_PAYLOAD_INVALID');
    N((get('memberPrivateProfiles/'+personId)?.revision||0)===p.profileRevision&&get('memberAggregateHeads/'+personId)?.revision===p.personRevision,'REVISION_CONFLICT');
    const patch=f4MemberProfileContract_().patch(p.patch,self,Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'));
    r.commandType=self?'members.self.update':'members.profile.correct';r.target={personId};r.payload=self?{patch}:{corrections:patch,justification:p.justification};cap=self?'MEMBERS_SELF_UPDATE':'MEMBERS_PROFILE_CORRECT';
    paths=[...(Object.hasOwn(patch,'displayName')?['memberIdentities/'+personId]:[]),...(Object.keys(patch).some(k=>k!=='displayName')?['memberPrivateProfiles/'+personId]:[])];resources=paths.slice();
    if(!self){N(C.id(p.documentReference),'COMMAND_PAYLOAD_INVALID');r.decisionRef=p.documentReference;facts={};}
   }else if(input.route==='/admin/ingressos-membros/cadastrar'){
    const keys=['nomeCompleto','emailPrincipal','rga','telefone'];C.exact(p,[...keys,...(Object.hasOwn(p,'dadosComplementares')?['dadosComplementares']:[])]);
    const E=f4MemberEnrollmentContract_(),registration=E.quick({fullName:p.nomeCompleto,email:p.emailPrincipal,rga:p.rga,phone:p.telefone});
    const approvalRef='AP_'+C.hash({email:registration.email,rga:registration.rga}).slice(0,40),approval=f4MemberPreRegistration_().approval(get,approvalRef,registration,now),policy=get('memberOperationalConfig/registration');
    N(policy?.enabled===true&&policy.stagedEnrollment===true,'ACCESS_REVOKED');
    const resolution=f4MemberRegistrationIdentity_().resolve(get,{email:registration.email,rga:registration.rga}),personId=resolution.personId||'P_'+op;
    const identifiers=resolution.identifiers.filter(v=>!get('memberIdentifierClaims/'+C.hash(v))).map((v,i)=>({id:'I'+i+'_'+op,...v}));
    r.commandType='members.ingress.execute';r.target={personId,linkId:null};r.payload={operation:'PRE_REGISTER',identity:resolution.personId?null:{displayName:registration.fullName},identifiers,registration,profile:p.dadosComplementares||{},approvalRef};r.decisionRef=approval.reference;cap='MEMBERS_INGRESS_EXECUTE';
    facts={schemaVersion:'member-preregistration-v1',personId,approvalRef,approvalRevision:approval.revision,policyRevision:policy.revision};
    paths=['memberEnrollments/'+op,'memberEnrollmentIndexes/'+personId,...identifiers.map(v=>'memberIdentifiers/'+v.id)];if(!resolution.personId)paths.push('memberIdentities/'+personId);resources=paths.slice();
   }else if(input.route==='/admin/ingressos-membros/convite'){
    C.exact(p,['enrollmentId','revision','action','reason']);N(C.op(p.enrollmentId)&&C.rev(p.revision),'COMMAND_PAYLOAD_INVALID');
    const old=f4MemberEnrollmentContract_().validate(get('memberEnrollments/'+p.enrollmentId)),policy=get('memberOperationalConfig/registration');N(old.revision===p.revision,'REVISION_CONFLICT');
    r.commandType='members.ingress.execute';r.target={personId:old.personId,linkId:null};r.payload={operation:'MANAGE_INVITATION',enrollmentId:old.id,action:p.action,reason:p.reason};r.decisionRef=old.approval.reference;cap='MEMBERS_INGRESS_EXECUTE';
    facts={schemaVersion:'member-invitation-action-v1',personId:old.personId,enrollmentId:old.id,enrollmentRevision:old.revision,policyRevision:policy?.revision};
    paths=['memberEnrollments/'+old.id,...(p.action==='CANCEL'?['memberEnrollmentIndexes/'+old.personId]:[])];resources=paths.slice();
   }else if(input.route==='/admin/ingressos-membros/admitir-excepcional'){
    N(get('memberOperationalConfig/registration')?.exceptionalDirectAdmission===true,'ACCESS_REVOKED');
    const fields=['modalidadeCadastro','nomeCompleto','nomeExibicao','emailPrincipal','telefone','rga','cursoId','cursoNomeOutro','dataIngresso','semestreEntrada','formaIngresso','origemCadastro','documentoReferencia','observacaoAdministrativa'];
    C.exact(p,[...fields,...(Object.hasOwn(p,'dadosComplementares')?['dadosComplementares']:[]),...(Object.hasOwn(p,'chaveIdempotencia')?['chaveIdempotencia']:[])]);
    const complementary=['dataNascimento','instagram','paisOrigemCodigo','municipioOrigemCodigo','cidadeNatal','ufOrigem','regiaoOrigem'],raw=Object.fromEntries(fields.map(k=>[k,p[k]]));
    if(p.dadosComplementares)N(Object.keys(p.dadosComplementares).every(k=>complementary.includes(k)),'COMMAND_PAYLOAD_INVALID');raw.dadosComplementares=Object.fromEntries(complementary.map(k=>[k,p.dadosComplementares?.[k]||'']));
    const today=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'),registration=f4MemberAdministrativeIngress_().normalize(raw,today),resolution=f4MemberRegistrationIdentity_().resolve(get,{email:registration.emailPrincipal,rga:registration.rga}),personId=resolution.personId||'P_'+op,linkId='V_'+op;
    const policy=get('memberOperationalConfig/registration');N(policy?.enabled===true&&Array.isArray(policy.semesterIds)&&policy.semesterIds.every(C.id),'ACCESS_REVOKED');
    const matches=policy.semesterIds.map(id=>get('academicSemesters/'+id)).filter(s=>s?.label===registration.semestreEntrada);N(matches.length===1,'COMMAND_PAYLOAD_INVALID');const semester=matches[0],course=get('memberCourseCatalog/'+registration.cursoId);N(course?.active,'COMMAND_PAYLOAD_INVALID');
    const identifiers=resolution.identifiers.filter(v=>!get('memberIdentifierClaims/'+C.hash(v))).map((v,i)=>({id:'I'+i+'_'+op,...v}));
    r.commandType='members.ingress.execute';r.target={personId,linkId};r.payload={operation:'RECORD_INGRESS',identity:resolution.personId?null:{displayName:registration.nomeExibicao||registration.nomeCompleto},identifiers,startDate:registration.dataIngresso,endDate:null,originalEventId:null};r.decisionRef='ADM_'+op;cap='MEMBERS_INGRESS_EXECUTE';
    facts={schemaVersion:'member-administrative-ingress-v1',personId,policyRevision:policy.revision,courseRevision:course.revision,semesterId:semester.id,semesterRevision:semester.revision,registration,approval:{reference:r.decisionRef,confirmedByUid:identity.uid,previouslyApproved:true}};
    paths=['memberLinks/'+linkId,'memberEvents/EVT_'+op,...identifiers.map(v=>'memberIdentifiers/'+v.id)];if(!resolution.personId)paths.push('memberIdentities/'+personId);
    else{
     const loaded=f4MemberAuthz_().load(get,personId,f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS'))),at=C.civilStart(registration.dataIngresso),overlap=loaded.M.rows(loaded.d,'memberLinks').filter(l=>{const iv=loaded.M.interval(l.startDate,l.endDate);return iv.a<=at&&at<iv.z;});
     N(overlap.length<=1,'RECONCILIATION_REQUIRED');if(overlap.length)paths.push('memberLinks/'+overlap[0].id,'memberEvents/COR_'+op);
    }
    resources=paths.slice();
   }else if(['/meu-vinculo/solicitar-desligamento','/meu-vinculo/solicitar-suspensao'].includes(input.route)){
    C.exact(p,['modality','reason','semesterId',...(input.route.endsWith('/solicitar-suspensao')?['startDate','endDate']:[]),...(Object.hasOwn(p,'submission')?['submission']:[])]);N(input.route.endsWith('/solicitar-suspensao')?(p.modality==='SUSPENSAO_TEMPORARIA'):(p.modality!=='SUSPENSAO_TEMPORARIA'),'COMMAND_PAYLOAD_INVALID');const a=f4MemberAuthz_().identify(get,identity.uid,now),link=a.actor.M.rows(a.actor.d,'memberLinks').find(l=>l.type===a.state.state&&a.actor.M.interval(l.startDate,l.endDate).a<=now&&now<a.actor.M.interval(l.startDate,l.endDate).z);N(link,'ACCESS_REVOKED');
    r.commandType='members.voluntary.submit';r.target={personId:binding.idPessoa,linkId:link.id,requestId:op};r.payload=p;cap='MEMBERS_VOLUNTARY_REQUEST';
    paths=['memberVoluntaryRequests/'+op,'memberVoluntaryIndexes/'+binding.idPessoa];resources=paths.slice();
   }else if(input.route==='/meu-vinculo/solicitacoes/acao'||input.route==='/admin/solicitacoes-vinculo/acao'){
    C.exact(p,['requestId','revision','action','reason','minutesReference','confirmed',...(Object.hasOwn(p,'externalReview')?['externalReview']:[]),...(Object.hasOwn(p,'graveException')?['graveException']:[])]);N(C.op(p.requestId)&&C.rev(p.revision),'COMMAND_PAYLOAD_INVALID');requestId=p.requestId;
    const old=f4MemberVoluntaryWorkflow_().validate(get('memberVoluntaryRequests/'+requestId));N(old.revision===p.revision,'REVISION_CONFLICT');
    const suspension=old.modality==='SUSPENSAO_TEMPORARIA',self=['CANCEL_SELF','COMPLEMENT'].includes(p.action);N(suspension||!Object.hasOwn(p,'graveException'),'COMMAND_PAYLOAD_INVALID');N(input.route.startsWith('/meu-vinculo/')?self:!self,'ACCESS_REVOKED');N(!self||old.personId===binding.idPessoa,'ACCESS_REVOKED');
    r.commandType=p.action==='APPROVE'&&!suspension?'members.membership.voluntaryExit':'members.voluntary.review';r.target={personId:old.personId,linkId:old.linkId,...(r.commandType==='members.membership.voluntaryExit'?{}:{requestId})};
    const review={action:p.action,reason:p.reason,minutesReference:p.minutesReference,confirmed:p.confirmed,...(suspension?{requestKind:old.modality}:{}),...(Object.hasOwn(p,'graveException')?{graveException:p.graveException}:{})};r.payload=review;cap=self?'MEMBERS_VOLUNTARY_REQUEST':suspension?'MEMBERS_VOLUNTARY_SUSPENSION_EXECUTE':'MEMBERS_VOLUNTARY_EXIT_EXECUTE';
    paths=['memberVoluntaryRequests/'+requestId,'memberVoluntaryIndexes/'+old.personId];resources=paths.slice();if(!self)resources.push('memberLinks/'+old.linkId);
    if(p.action==='APPROVE'&&!suspension){
     const policy=get('memberOperationalConfig/voluntaryExit'),semester=old.semesterId?get('academicSemesters/'+old.semesterId):null,today=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd');
     const decision=f4MemberVoluntaryRules_().finalDecision(f4MemberVoluntaryWorkflow_().legacy(old),{confirmacaoReforcada:p.confirmed,ataReferencia:p.minutesReference},policy?.minutes,today,semester);
     r.payload={effectiveDate:decision.effectiveDate};paths.push('memberEvents/EVT_'+op);resources.push('memberEvents/EVT_'+op);
     facts={schemaVersion:'member-voluntary-decision-v1',requestId,requestRevision:p.revision,policyRevision:policy.revision,review,externalReview:p.externalReview};
    }else if(!self){facts={};if(suspension&&['APPROVE','RETURN'].includes(p.action)){paths.push('memberEvents/EVT_'+op);resources.push('memberEvents/EVT_'+op);}}
   }else N(false,'COMMAND_TYPE_INVALID');
   N(settings.commands.includes(r.commandType),'ACCESS_REVOKED');
   paths.push('memberAggregateHeads/'+r.target.personId);r.expectedRevisions=[...new Set(paths)].sort().map(path=>({resourceType:path.split('/')[0],id:path.split('/')[1],revision:get(path)?.revision||0}));
   if(facts!==null)r.evidenceRefs=['MVE_'+op];
   const q=requestFor(get,r,resources,cap,now,end,settings);if(requestId&&cap!=='MEMBERS_VOLUNTARY_REQUEST')q.target.caseId=requestId;
   const pending={};
   const operationalChecks=r.commandType==='members.voluntary.submit'||r.commandType==='members.membership.voluntaryExit'||r.commandType==='members.voluntary.review'&&p.action==='APPROVE'?f4MemberActivityChecks_().capture(get,db,r.target.personId,now):null;
   const browser=Object.fromEntries(['schemaVersion','operationId','commandType','target','payload','evidenceRefs','decisionRef'].map(k=>[k,r[k]]));
   pending[intentPath]={schemaVersion:'member-prepared-intent-v1',operationId:op,ownerUid:identity.uid,inputHash,input:browser,createdAt:T(now),expiresAt:T(end),operationalChecks};
   if(facts!==null){const body={schemaVersion:'member-act-content-v13',...f4MemberActContent_().material(r),facts},hash=C.hash(body),ref=r.evidenceRefs[0];
    pending['memberEvidenceContents/'+hash]=body;pending['memberEvidenceAttestations/'+ref]={schemaVersion:'member-evidence-attestation-v13',id:ref,revision:1,active:true,kind:cap==='MEMBERS_PROFILE_CORRECT'?'CADASTRAL_EVIDENCE':cap==='MEMBERS_INGRESS_EXECUTE'?'BOARD_DECISION':'VOLUNTARY_REQUEST',holderId:r.target.personId,linkId:r.target.linkId||null,targetResourceIds:resources,effect:q.target.effect,effectiveDate:q.target.effectiveDate,issuerRef:identity.uid,approvalRef:r.decisionRef||'REQUEST_'+requestId,contentHash:hash,verifiedAt:T(now),expiresAt:T(end)};
   }
   const overlay=path=>Object.hasOwn(pending,path)?pending[path]:get(path),auth=r.commandType==='members.ingress.execute'?f4MemberAuthz_().authorizeCommand(overlay,q,r,now):f4MemberAuthz_().authorize(overlay,q,now);
   const simulated={...r,ticketId:'PREPARE_'+op,digest:K.digest(config.projectId,r),receivedAt:T(now),state:'RECEIVED'};
   const provenance={operationId:op,ownerUid:identity.uid,digest:simulated.digest,ticketId:simulated.ticketId,authRequests:[q]},ticket={ownerUid:identity.uid,digest:simulated.digest,requestEnvelope:K.template(simulated),authProofs:[auth.proof]};
   const preflight=path=>path==='commandRequests/'+op?simulated:path==='commandProvenance/'+op?provenance:path==='memberAdmissionTickets/'+simulated.ticketId?ticket:overlay(path);
   f4MemberFamilyCatalog_(config)[r.commandType].resolve(preflight,simulated,now);
   pending['devProtectedInputs/'+op]={schemaVersion:'core-dev-protected-input-v1',operationId:op,ownerUid:identity.uid,materialHash:C.hash(material(r)),expectedRevisions:r.expectedRevisions,authRequests:[q]};
   // All reads and validation precede the first write.
   for(const path of Object.keys(pending))N(!get(path),'OPERATION_CONFLICT');for(const [path,data]of Object.entries(pending))tx.set(path,data);return browser;
  });
 }
 function resume(key){K.selector(key);return db.transaction(tx=>{const get=p=>tx.get(p).data();configuration(get);current(get,clock());const v=get('memberPreparedIntents/'+key.operationId),p=get('devProtectedInputs/'+key.operationId);N(v?.ownerUid===identity.uid&&p?.ownerUid===identity.uid&&v.input.commandType===key.commandType,'ACCESS_REVOKED');N(K.digest(config.projectId,{...v.input,ownerUid:identity.uid,expectedRevisions:p.expectedRevisions})===key.digest,'OPERATION_CONFLICT');return v.input;});}
 return {prepare,resume,configuration,allowed};
}
