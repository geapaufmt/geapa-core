'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GEAPA_PILOT_MODE:'SHARED_DEV',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
if(process.env.GEAPA_TARGET_ENVIRONMENT==='PROD')delete process.env.GEAPA_PILOT_MODE;
const {test}=require('node:test'),assert=require('node:assert/strict');
const {randomUUID}=require('node:crypto'),{Timestamp}=require('firebase-admin/firestore');
const {seed}=require('./fixtures.cjs'),{decode}=require('../src/firestore.cjs');
const {createDomain}=require('../src/domain.cjs'),{executeMember}=require('../src/member-executor.cjs');
const {PROJECT,DEMO,EXECUTOR,PRODUCTION,profile}=require('../src/environment.cjs');

const {database}=require('./fixtures/member-memory-db.cjs');
test('ordinary absence limits follow the effective membership category',()=>{
 const domain=createDomain(()=>{});
 const counts=(state,countsAbsence)=>JSON.parse(domain.run('(v)=>f4MemberCountsAbsence_(v.state,v.activity)',{state,activity:{countsAbsence}}));
 assert.equal(counts('MEMBRO_INGRESSANTE',true),false);assert.equal(counts('MEMBRO_EM_ESPERA',true),false);
 assert.equal(counts('MEMBRO_EFETIVO',true),true);assert.equal(counts('MEMBRO_EFETIVO',false),false);
});
const op=()=>randomUUID().replaceAll('-','');
async function fixture(category,t){
 const {docs:raw,now}=seed(),docs=Object.fromEntries(Object.entries(raw).map(([key,v])=>[key,decode({mapValue:v})]));
 const own=docs['memberAuthorityControl/MEMBERS'].selfReadPolicy;
 docs['memberAuthzPolicies/'+own.policyId]={schemaVersion:'member-authz-policy-v13',id:own.policyId,revision:own.policyRevision,active:true,origin:'F4_M621_INSTITUTIONAL',approvalRef:'SYN_SELF',issuedAt:own.issuedAt,expiresAt:own.expiresAt,proofKind:'SELF',capabilities:['MEMBERS_SELF_READ'],actorSelector:{kind:'SELF'},targets:{ownPerson:true,allowedTargetScopes:[],allowedBoards:[],allowedRoles:[],allowedPersons:[],allowedResources:['memberCurrentStates/PES-000001']},fieldMasks:own.fieldMasks};
 const real=process.env.GEAPA_ENROLLMENT_EMULATOR_TEST==='1';
 const director=docs['memberRoleCatalog/CARGO-001'],rolePolicy=docs['memberAuthzPolicies/FAST_MEMBER_ROLE'];
 director.roleKey='DIRETOR_RECURSOS_HUMANOS';docs['memberFunctions/VIG-001'].scopeKey='RECURSOS_HUMANOS';
 rolePolicy.actorSelector={...rolePolicy.actorSelector,roleKey:director.roleKey,actorScope:'RECURSOS_HUMANOS'};
 const db=real?new (require('@google-cloud/firestore').Firestore)({projectId:DEMO,databaseId:'admission-'+op()}):database(docs);
 if(real)t.after(()=>db.terminate());
 const email='convite@example.invalid',rga='202610345';
 const hash=JSON.parse(createDomain(()=>{}).run('(v)=>f4Core_().hash(v)',{email,rga}));
 const ref=category==='MEMBRO_EFETIVO'?'FAST_PREVIOUS_APPROVAL':'AP_'+hash.slice(0,40),approval=Object.values(docs).find(v=>v.schemaVersion==='member-admission-approval-v1');
 docs['memberAdmissionApprovals/'+ref]={...approval,id:ref,membershipType:category,registrationDigest:hash,courseId:'AGRONOMIA_UFMT_SINOP'};
 const indexId='AP_'+hash.slice(0,40);docs['memberAdmissionApprovalIndexes/'+indexId]={schemaVersion:'member-admission-approval-index-v1',id:indexId,revision:1,registrationDigest:hash,approvalIds:[ref]};
 if(ref!==indexId)delete docs['memberAdmissionApprovals/'+indexId];
 if(category==='MEMBRO_EFETIVO'){
  docs['memberOperationalConfig/registration'].recognizePastAdmission=true;
  docs['memberAdmissionApprovals/'+ref].recognizedEffects={on:approval.approvedOn,reference:'SYN_PAST_EFFECT',approvedBy:approval.approvedBy};
 }
 docs['memberOperationalConfig/notifications']={...docs['memberOperationalConfig/notifications'],projectId:PROJECT,environment:profile.environment};
 const control={schemaVersion:'member-executor-control-v1',projectId:PROJECT,revision:1,enabled:true,actorUids:['CF_ADMIN'],scopeIds:['SYN_ADMISSION'],startsAt:now-1000,expiresAt:now+3600000,browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['SYN_APP']}};
 docs['memberExecutorControl/current']=control;
 docs['memberExecutorScopes/SYN_ADMISSION']={schemaVersion:'member-executor-scope-v1',id:'SYN_ADMISSION',projectId:PROJECT,synthetic:!PRODUCTION,revision:1,personIds:['PES-000001'],approvalRefs:[ref]};
 for(const type of ['ingress.execute','onboarding.complete']){
  const id='MEMBERS_'+type.toUpperCase().replace('.','_');
  docs['commandExecutorRoutes/'+id]={schemaVersion:'member-executor-route-v1',projectId:PROJECT,routeId:id,commandType:'members.'+type,scopeId:'SYN_ADMISSION',syntheticOnly:!PRODUCTION,executor:EXECUTOR,state:'ACTIVE',epoch:1,revision:1,actorUids:['CF_ADMIN'],operationIds:[],maxOperations:50,startsAt:now-1000,expiresAt:now+3600000};
 }
 if(real){
  const batch=db.batch();for(const [p,value] of Object.entries(docs))batch.set(db.doc(p),value);await batch.commit();
 }
 async function set(path,value){docs[path]=value;if(real){if(value===null)await db.doc(path).delete();else await db.doc(path).set(value);}}
 const appCheck={verifyToken:async()=>({appId:'SYN_APP'})};
 const call=async(uid,action,request,opts={})=>{
  const auth={verifyIdToken:async()=>({uid,sub:uid,aud:DEMO,iss:'https://securetoken.google.com/'+DEMO,exp:Math.floor(now/1000)+3600,iat:Math.floor(now/1000),auth_time:Math.floor(now/1000),email,email_verified:true,...opts.claims})};
  try{return await executeMember({db,auth,appCheck,token:'synthetic',appCheckToken:'synthetic',origin:profile.origin,body:{action,request},hooks:opts.hooks});}
  finally{
   if(real)for(const collection of ['memberEnrollments','memberEnrollmentIndexes','memberLinks','memberAuthBindings','memberCurrentStates','portalUsers','jobOutbox','memberAdmissionApprovals','memberAdmissionApprovalIndexes','memberIdentifiers','memberIdentifierClaims','memberNotificationIndexes','memberNotifications']){
    const snap=await db.collection(collection).get();for(const row of snap.docs)docs[row.ref.path]=row.data();
   }
  }
 };
 async function submit(uid,action,intent){
  const prepared=await call(uid,action,{operationId:op(),...intent});
  const response=await call(uid,'memberExecute',prepared.request);
  assert.equal(response.dto.receipt.state,'COMPLETED',JSON.stringify(response.dto.receipt));
  return prepared;
 }
 const rh=(route,payload)=>submit('CF_ADMIN','memberPrepare',{route,payload});
 const registration={nomeCompleto:'Pessoa Sintética',emailPrincipal:email,rga,telefone:'66999990000',categoriaIngresso:category,referenciaAprovacao:approval.reference};
 return {docs,call,submit,rh,registration,set};
}
for(const category of ['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE','DECLINE']){
 test('canonical enrollment from minimal pre-registration to '+category,async t=>{
  const type=category==='DECLINE'?'MEMBRO_INGRESSANTE':category;
  const {docs,call,rh,submit,registration,set}=await fixture(type,t);
  const route='/admin/ingressos-membros/cadastrar';
  await assert.rejects(rh(route,{...registration,referenciaAprovacao:'OUTRA'}),/ADMISSION_APPROVAL_REQUIRED/);
  const confirmation=await call('CF_ADMIN','memberRegistrationApproval',{email:registration.emailPrincipal.toUpperCase(),rga:registration.rga});assert.equal(confirmation.result.data.membershipType,type);assert.equal(confirmation.metrics.writes,0);
  const {categoriaIngresso,referenciaAprovacao,...basic}=registration;
  const registered=await rh(route,registration),id=registered.request.operationId,person=registered.request.target.personId;
  assert.equal(registered.request.decisionRef,registration.referenciaAprovacao);
  const enrollment=()=>docs['memberEnrollments/'+id];
  if(type==='MEMBRO_EFETIVO')assert.equal(enrollment().approval.recognizedEffects.reference,'SYN_PAST_EFFECT');
  assert(!docs['memberLinks/V_'+id]);assert(!docs['memberAuthBindings/SYN_INVITEE']);
  await assert.rejects(rh(route,registration),/OPERATION_CONFLICT/);
  const manage=(action,reason)=>rh('/admin/ingressos-membros/convite',{enrollmentId:id,revision:enrollment().revision,action,reason});
  await assert.rejects(manage('ACKNOWLEDGE','Teste'),/ACCESS_REVOKED|M13_AUTHZ_DENIED/);
  assert.equal(docs['jobOutbox/'+id]?.event,'ENROLLMENT_REGISTERED');
  if(category==='MEMBRO_INGRESSANTE'){
   const policy='memberOperationalConfig/registration',role='memberRoleCatalog/CARGO-001',position='memberFunctions/VIG-001',grant='memberAuthzPolicies/FAST_MEMBER_ROLE';
   const original=Object.fromEntries([role,position,grant].map(path=>[path,docs[path]]));
   async function presidency(on){
    await set(role,on?{...original[role],roleKey:'PRESIDENTE'}:original[role]);
    await set(position,on?{...original[position],scopeKey:'PRESIDENCIA'}:original[position]);
    await set(grant,on?{...original[grant],actorSelector:{...original[grant].actorSelector,roleKey:'PRESIDENTE',actorScope:'PRESIDENCIA'},targets:{...original[grant].targets,commandScopes:['ENROLLMENT_SUPERVISION']}}:original[grant]);
   }
   await set(policy,{...docs[policy],revision:2,presidencyReview:{required:true,decisionRef:'SYN_APPROVED_RULE'}});
   await assert.rejects(manage('SEND','Teste'),/ADMISSION_APPROVAL_REQUIRED/);
   await presidency(true);await manage('REQUEST_CORRECTION','Conferir telefone');
   await assert.rejects(manage('SEND','Teste'));
   await presidency(false);
   await rh('/admin/ingressos-membros/convite',{enrollmentId:id,revision:enrollment().revision,action:'AMEND',field:'phone',value:'66988880000',reason:'Contato corrigido pelo RH'});
   assert.equal(enrollment().registration.phone,'+5566988880000');
   await presidency(true);await manage('ACKNOWLEDGE','Conferido pela Presidência');await presidency(false);
  }
  assert.equal(docs['memberNotificationIndexes/'+person],undefined);
  await manage('SEND','Convite sintético');
  const inbox=docs['memberNotificationIndexes/'+person];
  assert.equal(inbox.unreadCount,1);
  assert.equal(docs['memberNotifications/'+inbox.notificationIds[0]].event,'ENROLLMENT_INVITATION');
  const open=(opts)=>call('SYN_INVITEE','memberEnrollmentOpen',{},opts);
  await assert.rejects(open({claims:{email_verified:false}}),/AUTH_BINDING_INVALID/);
  await assert.rejects(open({claims:{email:'diferente@example.invalid'}}),/AUTH_BINDING_INVALID/);
  if(category==='DECLINE'){
   const binding={...docs['memberAuthBindings/CF_ADMIN'],uid:'OTHER_SYNTHETIC',id:'OTHER_SYNTHETIC',idPessoa:person};
   await set('memberAuthBindings/OTHER_SYNTHETIC',binding);
   await assert.rejects(open(),/BINDING_CONFLICT/);
   await set('memberAuthBindings/OTHER_SYNTHETIC',{...binding,idPessoa:'OTHER_SYNTHETIC_PERSON'});
  }
  const view=(await open()).result.data;
  assert.equal(view.membershipType,type);assert.equal(view.entrySemester.semester,1);assert(!view.profile);
  if(category==='DECLINE'){
   const value='202620345',hash=JSON.parse(createDomain(()=>{}).run('(rga)=>f4MemberPreRegistration_().key({email:"convite@example.invalid",rga})',value));
   const amend=()=>rh('/admin/ingressos-membros/convite',{enrollmentId:id,revision:enrollment().revision,action:'AMEND',field:'rga',value,reason:'Correção documental sintética'});
   await assert.rejects(amend(),/ADMISSION_APPROVAL_REQUIRED/);
   const source=Object.values(docs).find(v=>v.schemaVersion==='member-admission-approval-v1'&&v.reference===registration.referenciaAprovacao);
   const digest=JSON.parse(createDomain(()=>{}).run('(rga)=>f4Core_().hash({email:"convite@example.invalid",rga})',value));
   await set('memberAdmissionApprovals/'+hash,{...source,id:hash,registrationDigest:digest,membershipType:type});
   const scope=docs['memberExecutorScopes/SYN_ADMISSION'];
   await set('memberExecutorScopes/SYN_ADMISSION',{...scope,approvalRefs:[...scope.approvalRefs,hash]});
   await amend();assert.equal((await open()).result.data.entrySemester.semester,2);
  }
  const candidate=(action,payload={})=>submit('SYN_INVITEE','memberEnrollmentPrepare',{revision:enrollment().revision,action,payload});
  if(category==='DECLINE'){await candidate('DECLINE');assert.equal(enrollment().status,'DECLINED');assert(!docs['memberAuthBindings/SYN_INVITEE']);return;}
  await candidate('ACCEPT');await candidate('SAVE',{profile:{city:'Sinop'}});
  assert.equal((await open()).result.data.profile.city,'Sinop');
  await assert.rejects(candidate('SAVE',{profile:{rga:'202610000'}}),/COMMAND_PAYLOAD_INVALID/);
  await assert.rejects(candidate('SAVE',{profile:{courseId:'UNLISTED'}}),/COMMAND_PAYLOAD_INVALID/);
  await candidate('SAVE',{profile:{cpf:'52998224725',birthDate:'2001-05-01',countryCode:'BR',stateCode:'MT'}});
  const termsPath='memberOperationalConfig/enrollmentTerms',publishedTerms=docs[termsPath];
  await set(termsPath,{...publishedTerms,enabled:false});
  assert.equal((await open()).result.data.termsUnavailable,true);await candidate('SAVE',{profile:{city:'Sinop'}});
  await assert.rejects(candidate('COMPLETE'));await set(termsPath,publishedTerms);
  const terms=(await call('SYN_INVITEE','memberEnrollmentRead',{})).result.data;
  await candidate('ACKNOWLEDGE',{accepted:true,termVersion:terms.normTermVersion,documents:terms.norms.map(({id,version})=>({id,version}))});
  await candidate('IMAGE',{choice:'DECLINED',termVersion:terms.imageTerm.version});
  const prepared=await call('SYN_INVITEE','memberEnrollmentPrepare',{operationId:op(),revision:enrollment().revision,action:'COMPLETE',payload:{}});
  await open(); // Reopening a valid invitation preserves the prepared proof.
  await assert.rejects(call('SYN_INVITEE','memberExecute',prepared.request,{hooks:{afterCommit(){throw Error('LOST_RESPONSE');}}}),/LOST_RESPONSE/);
  const receipt=await call('SYN_INVITEE','memberRecover',prepared.key);
  assert.equal(receipt.dto.receipt.state,'COMPLETED');assert.equal(enrollment().status,'COMPLETED');
  assert.equal(docs['portalUsers/SYN_INVITEE'].idPessoa,person);assert.equal(docs['memberCurrentStates/'+person].state,type);
  const link=docs['memberLinks/V_'+id];assert.equal(link.type,type);assert.equal(docs['memberAuthBindings/SYN_INVITEE'].idPessoa,person);
  const localDay=new Intl.DateTimeFormat('en-CA',{timeZone:'America/Cuiaba',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(enrollment().completedAt));
  assert.equal(link.startDate,localDay);
  assert(enrollment().createdAt<=enrollment().respondedAt&&enrollment().respondedAt<=enrollment().completedAt);
  assert.equal(Object.values(docs).filter(v=>v.schemaVersion==='m1-memberLinks-v1'&&v.idPessoa===person).length,1);
  if(type==='MEMBRO_INGRESSANTE'){
   const input={personId:person,recognizedOn:link.startDate,reference:''};
   await assert.rejects(call('SYN_INVITEE','memberPrepare',{operationId:op(),route:'/admin/membros/efetivar',payload:input}));
   await rh('/admin/membros/efetivar',input);
   assert.equal(docs['memberLinks/V_'+id].type,'MEMBRO_EFETIVO');
   await assert.rejects(rh('/admin/membros/efetivar',input),/DOMAIN_STATE_INVALID/);
  }
 });
}

test('RH effectivates a prior incoming link and preserves its interval and event',async t=>{
 const {docs,set,rh}=await fixture('MEMBRO_INGRESSANTE',t);
 const id='SYN_PRIOR',link='SYN_PRIOR_LINK',event='SYN_PRIOR_ENTRY';
 const day=delta=>new Intl.DateTimeFormat('en-CA',{timeZone:'America/Cuiaba',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(Date.now()+delta*86400000));
 const scope=docs['memberExecutorScopes/SYN_ADMISSION'];
 await set('memberExecutorScopes/SYN_ADMISSION',{...scope,personIds:[...scope.personIds,id]});
 await set('memberIdentities/'+id,{...docs['memberIdentities/PES-000001'],id,idPessoa:id,displayName:'Pessoa sintética anterior'});
 await set('memberLinks/'+link,{...docs['memberLinks/VINC-001'],id:link,idPessoa:id,type:'MEMBRO_INGRESSANTE',startDate:day(-10)});
 await set('memberEvents/'+event,{schemaVersion:'m1-memberEvents-v1',id:event,revision:1,origin:'SYN_TEST',idPessoa:id,idVinculo:link,type:'INGRESSO',effect:'INGRESSO',effectiveDate:day(-10),endDate:null,supersedes:null,correctionPath:null,prior:null});
 await set('memberAggregateHeads/'+id,{schemaVersion:'m11-head-v1',id,kind:'PERSON',revision:1,paths:['memberIdentities/'+id,'memberLinks/'+link,'memberEvents/'+event]});
 const promote=date=>rh('/admin/membros/efetivar',{personId:id,recognizedOn:date,reference:''});
 await assert.rejects(promote(day(1)));await assert.rejects(promote(day(-1)));
 const result=await promote(day(0));
 assert.equal(docs['memberLinks/'+link].endDate,day(-1));assert.equal(docs['memberLinks/'+link].type,'MEMBRO_INGRESSANTE');
 assert.equal(docs['memberLinks/'+result.request.target.linkId].startDate,day(0));assert.equal(docs['memberCurrentStates/'+id].state,'MEMBRO_EFETIVO');
});

test('draft without any approval reserves one identity, recovers, invites and waits for a genuine approval before activation',async t=>{
 const {docs,call,rh,submit,registration,set}=await fixture('MEMBRO_INGRESSANTE',t);
 const approvals=Object.entries(docs).filter(([p])=>/^memberAdmissionApproval(s|Indexes)\//.test(p));
 for(const [path]of approvals)await set(path,null);
 const {categoriaIngresso,referenciaAprovacao,...basic}=registration,route='/admin/ingressos-membros/cadastrar';
 // The old scope does not silently gain access to unapproved registrations.
 await assert.rejects(rh(route,basic),/PILOT_RESOURCE_DENIED|DOMAIN_STATE_INVALID/);
 await set('memberExecutorScopes/SYN_ADMISSION',{...docs['memberExecutorScopes/SYN_ADMISSION'],personIds:[],approvalRefs:[],preregistrationDrafts:true});
 const policy=docs['memberAuthzPolicies/FAST_MEMBER_ROLE'];
 await set('memberAuthzPolicies/FAST_MEMBER_ROLE',{...policy,active:false});
 await assert.rejects(rh(route,basic),/ACCESS_REVOKED|M13_AUTHZ_DENIED/);
 await set('memberAuthzPolicies/FAST_MEMBER_ROLE',policy);
 const input={operationId:op(),route,payload:basic},registered=await call('CF_ADMIN','memberPrepare',input),id=registered.request.operationId,person=registered.request.target.personId;
 assert.equal(registered.request.payload.approvalRef,null);assert.equal(registered.request.decisionRef,null);
 assert.deepEqual((await call('CF_ADMIN','memberPrepare',input)).key,registered.key);
 await assert.rejects(call('CF_ADMIN','memberExecute',registered.request,{hooks:{afterCommit(){throw Error('LOST_RESPONSE');}}}),/LOST_RESPONSE/);
 assert.equal((await call('CF_ADMIN','memberRecover',registered.key)).dto.receipt.state,'COMPLETED');
 const enrollment=()=>docs['memberEnrollments/'+id];
 assert.equal(enrollment().approval,null);assert.equal(enrollment().draft,true);
 assert(!docs['memberLinks/V_'+id]);assert(!docs['memberAuthBindings/SYN_INVITEE']);
 assert.equal(docs['memberCurrentStates/'+person].state,'SEM_VINCULO');
 await assert.rejects(rh(route,basic),/OPERATION_CONFLICT/);
 assert.equal(Object.values(docs).filter(v=>v?.schemaVersion==='member-enrollment-v1'&&v.personId===person).length,1);
 const manage=(action,extra={})=>rh('/admin/ingressos-membros/convite',{enrollmentId:id,revision:enrollment().revision,action,reason:'Decisão humana sintética para homologação',...extra});
 const approvedOn=new Intl.DateTimeFormat('en-CA',{timeZone:'America/Cuiaba',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
 const decision={membershipType:'MEMBRO_INGRESSANTE',approvedOn,confirmed:true};
 await assert.rejects(manage('CONFIRM_ADMISSION',{...decision,confirmed:false}),/COMMAND_PAYLOAD_INVALID/);
 await assert.rejects(manage('CONFIRM_ADMISSION',{...decision,membershipType:'ADMIN'}),/ADMISSION_CATEGORY_INVALID/);
 const role=docs['memberRoleCatalog/CARGO-001'],fn=docs['memberFunctions/VIG-001'],pol=docs['memberAuthzPolicies/FAST_MEMBER_ROLE'];
 await set('memberRoleCatalog/CARGO-001',{...role,roleKey:'DIRETOR_ADMINISTRATIVO'});await set('memberFunctions/VIG-001',{...fn,scopeKey:'ADMINISTRACAO'});await set('memberAuthzPolicies/FAST_MEMBER_ROLE',{...pol,actorSelector:{...pol.actorSelector,roleKey:'DIRETOR_ADMINISTRATIVO',actorScope:'ADMINISTRACAO'}});
 await assert.rejects(manage('CONFIRM_ADMISSION',decision),/ACCESS_REVOKED/);
 await set('memberRoleCatalog/CARGO-001',role);await set('memberFunctions/VIG-001',fn);await set('memberAuthzPolicies/FAST_MEMBER_ROLE',pol);
 const prepareDecision=()=>call('CF_ADMIN','memberPrepare',{operationId:op(),route:'/admin/ingressos-membros/convite',payload:{enrollmentId:id,revision:enrollment().revision,action:'CONFIRM_ADMISSION',reason:'Confirmação humana sintética de admissão',...decision}});
 const first=await prepareDecision(),concurrent=await prepareDecision();
 await assert.rejects(call('CF_ADMIN','memberExecute',first.request,{hooks:{afterCommit(){throw Error('LOST_RESPONSE');}}}),/LOST_RESPONSE/);
 assert.equal((await call('CF_ADMIN','memberRecover',first.key)).dto.receipt.state,'COMPLETED');
 await assert.rejects(call('CF_ADMIN','memberExecute',concurrent.request),/REVISION_CONFLICT/);
 const firstApproval=enrollment().approvalId;assert.equal(enrollment().status,'PRE_REGISTERED');assert(!docs['memberLinks/V_'+id]);
 await manage('AMEND',{field:'email',value:'draft.changed@example.invalid'});assert.equal(enrollment().approval,null);assert.equal(docs['memberAdmissionApprovals/'+firstApproval].active,false);
 await manage('AMEND',{field:'email',value:basic.emailPrincipal});
 await manage('SEND');
 const open=()=>call('SYN_INVITEE','memberEnrollmentOpen',{});
 const invited=(await open()).result.data;assert.equal(invited.membershipType,null);assert.equal(invited.admissionApprovalPending,true);
 const candidate=(action,payload={})=>submit('SYN_INVITEE','memberEnrollmentPrepare',{revision:enrollment().revision,action,payload});
 await candidate('ACCEPT');
 await candidate('SAVE',{profile:{cpf:'52998224725',birthDate:'2001-05-01',countryCode:'BR',stateCode:'MT',city:'Sinop',courseId:'AGRONOMIA_UFMT_SINOP'}});
 const terms=(await open()).result.data;
 await candidate('ACKNOWLEDGE',{accepted:true,termVersion:terms.normTermVersion,documents:terms.norms.map(({id,version})=>({id,version}))});
 await assert.rejects(candidate('COMPLETE'),/ADMISSION_APPROVAL_REQUIRED|ADMISSION_APPROVAL_INACTIVE/);
 await assert.rejects(candidate('COMPLETE',{approval:{membershipType:'MEMBRO_EFETIVO'}}),/SCHEMA_INVALID/);
 assert.equal(enrollment().status,'ACCEPTED');assert.equal(enrollment().approval,null);assert(!docs['memberAuthBindings/SYN_INVITEE']);
 // RH records a separate human decision; changing RGA revokes it atomically.
 await manage('CONFIRM_ADMISSION',decision);const priorApproval=enrollment().approvalId;
 await manage('AMEND',{field:'rga',value:'202620999'});assert.equal(enrollment().approval,null);assert.equal(docs['memberAdmissionApprovals/'+priorApproval].active,false);
 assert.equal((await open()).result.data.admissionApprovalPending,true);await assert.rejects(candidate('COMPLETE'),/ADMISSION_APPROVAL_REQUIRED|ADMISSION_APPROVAL_INACTIVE/);
 await manage('CONFIRM_ADMISSION',decision);
 assert.equal((await open()).result.data.admissionApprovalPending,false);
 const ready=await call('SYN_INVITEE','memberEnrollmentPrepare',{operationId:op(),revision:enrollment().revision,action:'COMPLETE',payload:{}});
 const activeApproval=['memberAdmissionApprovals/'+enrollment().approvalId,docs['memberAdmissionApprovals/'+enrollment().approvalId]];
 const invalidated=await call('SYN_INVITEE','memberExecute',ready.request,{hooks:{afterClaim:async()=>set(activeApproval[0],{...activeApproval[1],membershipType:'MEMBRO_EFETIVO'})}});assert.equal(invalidated.dto.receipt.state,'FAILED');
 assert(!docs['memberAuthBindings/SYN_INVITEE']);
 await set(activeApproval[0],{...activeApproval[1],active:false});
 await assert.rejects(candidate('COMPLETE'),/ADMISSION_APPROVAL_INACTIVE|ACCESS_REVOKED/);
 assert(!docs['memberAuthBindings/SYN_INVITEE']);
 await set(activeApproval[0],activeApproval[1]);
 assert.equal((await call('SYN_INVITEE','memberRecover',ready.key)).dto.receipt.state,'FAILED');
 const completed=await candidate('COMPLETE');
 assert.equal((await call('SYN_INVITEE','memberRecover',completed.key)).dto.receipt.state,'COMPLETED');
 assert.equal(enrollment().status,'COMPLETED');assert.equal(enrollment().approval.reference,activeApproval[1].reference);
 assert.equal(docs['memberCurrentStates/'+person].state,'MEMBRO_INGRESSANTE');
 assert.equal(Object.values(docs).filter(v=>v?.schemaVersion==='m1-memberLinks-v1'&&v.idPessoa===person).length,1);
});

for(const terminal of ['CANCELLED','DECLINED'])test('terminal '+terminal+' admission preserves history and permits a separately confirmed new enrollment',async t=>{
 const {docs,call,rh,submit,registration,set}=await fixture('MEMBRO_INGRESSANTE',t);
 for(const path of Object.keys(docs).filter(p=>/^memberAdmissionApproval(s|Indexes)\//.test(p)))await set(path,null);
 await set('memberExecutorScopes/SYN_ADMISSION',{...docs['memberExecutorScopes/SYN_ADMISSION'],personIds:[],approvalRefs:[],preregistrationDrafts:true});
 const {categoriaIngresso,referenciaAprovacao,...basic}=registration,route='/admin/ingressos-membros/cadastrar';
 const approvedOn=new Intl.DateTimeFormat('en-CA',{timeZone:'America/Cuiaba',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
 const decision={membershipType:'MEMBRO_INGRESSANTE',approvedOn,confirmed:true};
 const first=await rh(route,basic),oldId=first.request.operationId,person=first.request.target.personId;
 let id=oldId;
 const enrollment=()=>docs['memberEnrollments/'+id];
 const manage=(action,extra={})=>rh('/admin/ingressos-membros/convite',{enrollmentId:id,revision:enrollment().revision,action,reason:'Decisão sintética explícita para novo ingresso',...extra});
 const candidate=(action,payload={})=>submit('SYN_INVITEE','memberEnrollmentPrepare',{revision:enrollment().revision,action,payload});
 await manage('CONFIRM_ADMISSION',decision);await manage('SEND');
 if(terminal==='CANCELLED')await manage('CANCEL');
 else {await call('SYN_INVITEE','memberEnrollmentOpen',{});await candidate('DECLINE');}
 const previous=structuredClone(enrollment()),approvalPath='memberAdmissionApprovals/'+previous.approvalId,oldApproval=structuredClone(docs[approvalPath]);
 assert.equal(previous.status,terminal);assert.equal(oldApproval.active,true);
 const second=await rh(route,basic);id=second.request.operationId;
 assert.equal(second.request.target.personId,person);assert.equal(enrollment().approval,null);
 // A stale approval index must never authorize replacing a still-open enrollment.
 await set('memberEnrollments/'+oldId,{...previous,status:terminal==='CANCELLED'?'INVITED':'ACCEPTED'});
 await assert.rejects(manage('CONFIRM_ADMISSION',decision),/OPERATION_CONFLICT/);
 assert.deepEqual(docs[approvalPath],oldApproval);assert.equal(enrollment().approval,null);
 await set('memberEnrollments/'+oldId,previous);
 const ready=await call('CF_ADMIN','memberPrepare',{operationId:op(),route:'/admin/ingressos-membros/convite',payload:{enrollmentId:id,revision:enrollment().revision,action:'CONFIRM_ADMISSION',reason:'Confirmação sintética posterior ao encerramento',...decision}});
 await assert.rejects(call('CF_ADMIN','memberExecute',ready.request,{hooks:{afterCommit(){throw Error('LOST_RESPONSE');}}}),/LOST_RESPONSE/);
 assert.equal((await call('CF_ADMIN','memberRecover',ready.key)).dto.receipt.state,'COMPLETED');
 assert.deepEqual(docs[approvalPath],{...oldApproval,revision:oldApproval.revision+1,active:false});
 assert.deepEqual(docs['memberEnrollments/'+oldId],previous);
 const index=Object.values(docs).find(v=>v?.schemaVersion==='member-admission-approval-index-v1'&&v.personId===person);
 assert.equal(index.enrollmentId,id);assert.deepEqual(index.approvalIds,[enrollment().approvalId]);
 assert.equal(docs['memberAdmissionApprovals/'+enrollment().approvalId].active,true);
 await manage('SEND');await call('SYN_INVITEE','memberEnrollmentOpen',{});await candidate('ACCEPT');
 await candidate('SAVE',{profile:{cpf:'52998224725',birthDate:'2001-05-01',countryCode:'BR',stateCode:'MT',city:'Sinop',courseId:'AGRONOMIA_UFMT_SINOP'}});
 const terms=(await call('SYN_INVITEE','memberEnrollmentRead',{})).result.data;
 await candidate('ACKNOWLEDGE',{accepted:true,termVersion:terms.normTermVersion,documents:terms.norms.map(({id,version})=>({id,version}))});
 await candidate('COMPLETE');assert.equal(enrollment().status,'COMPLETED');
 assert.equal((await call('CF_ADMIN','memberRecover',first.key)).dto.receipt.state,'COMPLETED');
 assert.deepEqual(docs['memberEnrollments/'+oldId],previous);
 assert.equal(Object.values(docs).filter(v=>v?.schemaVersion==='m1-memberLinks-v1'&&v.idPessoa===person).length,1);
 assert.equal(Object.values(docs).filter(v=>v?.schemaVersion==='member-enrollment-v1'&&v.personId===person).length,2);
});

test('DEV technical draft window permits only its exact synthetic fixture and never grants invitation or membership',{skip:PRODUCTION},async t=>{
 const {docs,set,call}=await fixture('MEMBRO_INGRESSANTE',t),{draftFixture,planDevDraft}=require('../scripts/administer-rh-registration.cjs'),now=Date.now(),f=draftFixture('CF_ADMIN',now);
 for(const [p,v]of Object.entries(f.docs))await set(p,v);
 const actor=docs['memberAuthBindings/CF_ADMIN'],subject=docs['memberAuthzSubjects/'+actor.idPessoa],gp='coreDevAdministrativeGrants/CF_ADMIN';
 await set(gp,{schemaVersion:'core-dev-administrative-grant-v1',id:'CF_ADMIN',uid:'CF_ADMIN',personId:actor.idPessoa,bindingVersion:actor.bindingVersion,subjectEpoch:subject.authzEpoch,projectId:PROJECT,environment:'DEV',exportable:false,active:true,revision:1,issuedAt:new Date(now-1000),expiresAt:new Date(now+3600000),approvalRef:'SYN_DRAFT_REHEARSAL',capabilities:['MEMBERS_INGRESS_EXECUTE'],fieldMasks:{MEMBERS_INGRESS_EXECUTE:['preRegister']},syntheticPersonIds:[f.personId],sections:[]});
 const head=docs['memberAggregateHeads/'+actor.idPessoa];await set('memberAggregateHeads/'+actor.idPessoa,{...head,paths:head.paths.filter(p=>!p.startsWith('memberFunctions/'))});
 const scope=docs['memberExecutorScopes/SYN_ADMISSION'];await set('memberExecutorScopes/SYN_ADMISSION',{...scope,personIds:[f.personId],approvalRefs:[],preregistrationDrafts:true});
 const plan=planDevDraft({docs,uid:'CF_ADMIN',personId:f.personId,registration:f.registration,expectedGrantRevision:1,now});for(const [p,v]of Object.entries(plan.writes))await set(p,v);
 const payload={nomeCompleto:f.registration.fullName,emailPrincipal:f.registration.email,rga:f.registration.rga,telefone:f.registration.phone},route='/admin/ingressos-membros/cadastrar';
 const prepare=value=>call('CF_ADMIN','memberPrepare',{operationId:op(),route,payload:value});
 await assert.rejects(prepare({...payload,telefone:'+5566988880000'}),/AUTHORIZATION_DENIED/);
 await assert.rejects(prepare({...payload,dadosComplementares:{courseId:'AGRONOMIA_UFMT_SINOP'}}),/AUTHORIZATION_DENIED/);
 const ready=await prepare(payload),id=ready.request.operationId;
 assert.equal((await call('CF_ADMIN','memberExecute',ready.request)).dto.receipt.state,'COMPLETED');
 const repeat=await prepare(payload);assert.equal(repeat.recoveredExisting,true);assert.equal(repeat.dto.operationId,id);
 const e=docs['memberEnrollments/'+id];assert.equal(e.approval,null);assert.equal(e.personId,f.personId);assert(!docs['memberLinks/V_'+id]);
 await assert.rejects(call('CF_ADMIN','memberPrepare',{operationId:op(),route:'/admin/ingressos-membros/convite',payload:{enrollmentId:id,revision:e.revision,action:'SEND',reason:'Negativa de janela sem convite'}}),/AUTHORIZATION_DENIED/);
 const {planDevAdmissionTest}=require('../scripts/administer-rh-registration.cjs');
 const admission=planDevAdmissionTest({docs,uid:'CF_ADMIN',enrollmentId:id,expectedGrantRevision:2,now});for(const [p,v]of Object.entries(admission.writes))await set(p,v);
 const window=docs[gp].admissionTestWindow,original=structuredClone(docs[gp]);
 const confirm=(extra={})=>call('CF_ADMIN','memberPrepare',{operationId:op(),route:'/admin/ingressos-membros/convite',payload:{enrollmentId:id,revision:docs['memberEnrollments/'+id].revision,action:'CONFIRM_ADMISSION',membershipType:window.membershipType,approvedOn:window.approvedOn,reason:window.reason,confirmed:true,...extra}});
 await assert.rejects(confirm({membershipType:'MEMBRO_EFETIVO'}),/ACCESS_REVOKED|AUTHORIZATION_DENIED/);
 await assert.rejects(confirm({reason:'Outra decisão que não foi autorizada'}),/ACCESS_REVOKED|AUTHORIZATION_DENIED/);
 await set(gp,{...original,admissionTestWindow:{...window,expiresAt:new Date(now-1),issuedAt:new Date(now-3600000)}});
 await assert.rejects(confirm(),/ACCESS_REVOKED|AUTHORIZATION_DENIED/);await set(gp,original);
 const identityPath='memberIdentities/'+f.personId,identity=docs[identityPath];await set(identityPath,{...identity,origin:'REAL_SOURCE'});
 await assert.rejects(confirm(),/ACCESS_REVOKED|AUTHORIZATION_DENIED/);await set(identityPath,identity);
 const decision=await confirm();assert.equal((await call('CF_ADMIN','memberExecute',decision.request)).dto.receipt.state,'COMPLETED');
 assert.equal(docs['memberEnrollments/'+id].status,'PRE_REGISTERED');assert(!docs['memberLinks/V_'+id]);
 const reader=createDomain(raw=>{const q=JSON.parse(raw);if(q.kind==='get')return JSON.stringify(docs[q.path]?require('../src/firestore.cjs').encode(docs[q.path]):null);if(['guard','inspect','getAll'].includes(q.kind))return 'null';throw Error('UNEXPECTED_READER_IO_'+q.kind);});
 const read=()=>JSON.parse(reader.run('(now)=>f4MemberOperationalReader_({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:now+60000},clock:()=>now}).read({route:"/admin/ingressos-membros",params:{}})',now));
 assert.equal(read().data.items[0].enrollmentId,id);assert.equal(read().data.items[0].canSend,true);assert.equal(read().data.items[0].canConfirmAdmission,false);
 const sent=await call('CF_ADMIN','memberPrepare',{operationId:op(),route:'/admin/ingressos-membros/convite',payload:{enrollmentId:id,revision:docs['memberEnrollments/'+id].revision,action:'SEND',reason:'Convite sintético autorizado'}});
 assert.equal((await call('CF_ADMIN','memberExecute',sent.request)).dto.receipt.state,'COMPLETED');
 const closeAdmission=planDevAdmissionTest({docs,uid:'CF_ADMIN',enrollmentId:id,expectedGrantRevision:3,close:true,now});for(const [p,v]of Object.entries(closeAdmission.writes))await set(p,v);
 assert.equal(docs[gp].admissionTestWindow,undefined);
 assert.equal(read().data.items.length,0);
 const closed=planDevDraft({docs,uid:'CF_ADMIN',personId:f.personId,registration:f.registration,expectedGrantRevision:4,close:true,now});for(const [p,v]of Object.entries(closed.writes))await set(p,v);
 await assert.rejects(prepare(payload),/ADMISSION_APPROVAL_REQUIRED|AUTHORIZATION_DENIED/);
 assert.equal(docs[gp].preregistrationWindow,undefined);assert.equal(Object.values(docs).filter(v=>v?.schemaVersion==='m1-memberLinks-v1'&&v.idPessoa===f.personId).length,0);
});
