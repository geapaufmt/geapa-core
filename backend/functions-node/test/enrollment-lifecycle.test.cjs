'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GEAPA_PILOT_MODE:'SHARED_DEV',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),assert=require('node:assert/strict');
const {randomUUID}=require('node:crypto'),{Timestamp}=require('firebase-admin/firestore');
const {seed}=require('./fixtures.cjs'),{decode}=require('../src/firestore.cjs');
const {createDomain}=require('../src/domain.cjs'),{executeMember}=require('../src/member-executor.cjs');
const {PROJECT,DEMO,EXECUTOR}=require('../src/environment.cjs');

// Snapshot and atomic commit port for offline regression. No external recipient or database.
function database(docs){
 let version=Date.now();
 const doc=path=>({path});
 const snap=ref=>({exists:!!docs[ref.path],ref,data:()=>docs[ref.path],updateTime:Timestamp.fromMillis(version)});
 const getAll=async(...refs)=>refs.map(snap);
 function collection(name){
  assert.equal(name,'memberAuthBindings');
  let person;
  return {where(field,op,value){assert.equal(field,'idPessoa');assert.equal(op,'==');person=value;return this;},
   orderBy(){return this;},limit(n){assert.equal(n,2);return this;},
   async get(){const rows=Object.keys(docs).filter(p=>p.startsWith(name+'/')&&docs[p].idPessoa===person).sort().slice(0,2).map(p=>snap(doc(p)));return {docs:rows,size:rows.length,empty:!rows.length};}};
 }
 return {projectId:DEMO,doc,getAll,collection,async runTransaction(fn){
  const writes=[];
  const result=await fn({getAll,get:async ref=>ref.get?ref.get():snap(ref),set:(r,v)=>writes.push([r.path,v]),update:(r,v)=>writes.push([r.path,v]),delete:r=>writes.push([r.path,null])});
  for(const [p,v] of writes){if(v)docs[p]=v;else delete docs[p];}version++;return result;
 }};
}
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
 if(category==='MEMBRO_EFETIVO'){
  docs['memberOperationalConfig/registration'].recognizePastAdmission=true;
  docs['memberAdmissionApprovals/'+ref].recognizedEffects={on:approval.approvedOn,reference:'SYN_PAST_EFFECT',approvedBy:approval.approvedBy};
 }
 const control={schemaVersion:'member-executor-control-v1',projectId:PROJECT,revision:1,enabled:true,actorUids:['CF_ADMIN'],scopeIds:['SYN_ADMISSION'],startsAt:now-1000,expiresAt:now+3600000,browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['SYN_APP']}};
 docs['memberExecutorControl/current']=control;
 docs['memberExecutorScopes/SYN_ADMISSION']={schemaVersion:'member-executor-scope-v1',id:'SYN_ADMISSION',projectId:PROJECT,synthetic:true,revision:1,personIds:['PES-000001'],approvalRefs:[ref]};
 for(const type of ['ingress.execute','onboarding.complete']){
  const id='MEMBERS_'+type.toUpperCase().replace('.','_');
  docs['commandExecutorRoutes/'+id]={schemaVersion:'member-executor-route-v1',projectId:PROJECT,routeId:id,commandType:'members.'+type,scopeId:'SYN_ADMISSION',syntheticOnly:true,executor:EXECUTOR,state:'ACTIVE',epoch:1,revision:1,actorUids:['CF_ADMIN'],operationIds:[],maxOperations:50,startsAt:now-1000,expiresAt:now+3600000};
 }
 if(real){
  const batch=db.batch();for(const [p,value] of Object.entries(docs))batch.set(db.doc(p),value);await batch.commit();
 }
 async function set(path,value){docs[path]=value;if(real)await db.doc(path).set(value);}
 const appCheck={verifyToken:async()=>({appId:'SYN_APP'})};
 const call=async(uid,action,request,opts={})=>{
  const auth={verifyIdToken:async()=>({uid,sub:uid,aud:DEMO,iss:'https://securetoken.google.com/'+DEMO,exp:Math.floor(now/1000)+3600,iat:Math.floor(now/1000),auth_time:Math.floor(now/1000),email,email_verified:true,...opts.claims})};
  try{return await executeMember({db,auth,appCheck,token:'synthetic',appCheckToken:'synthetic',origin:'https://geapa-convergencia-dev.web.app',body:{action,request},hooks:opts.hooks});}
  finally{
   if(real)for(const collection of ['memberEnrollments','memberLinks','memberAuthBindings','memberCurrentStates','portalUsers','jobOutbox']){
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
  const registered=await rh(route,registration),id=registered.request.operationId,person=registered.request.target.personId;
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
  await manage('SEND','Convite sintético');
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
  const candidate=(action,payload={})=>submit('SYN_INVITEE','memberEnrollmentPrepare',{revision:enrollment().revision,action,payload});
  if(category==='DECLINE'){await candidate('DECLINE');assert.equal(enrollment().status,'DECLINED');assert(!docs['memberAuthBindings/SYN_INVITEE']);return;}
  await candidate('ACCEPT');await candidate('SAVE',{profile:{city:'Sinop'}});
  assert.equal((await open()).result.data.profile.city,'Sinop');
  await assert.rejects(candidate('SAVE',{profile:{rga:'202610000'}}),/COMMAND_PAYLOAD_INVALID/);
  await assert.rejects(candidate('SAVE',{profile:{courseId:'UNLISTED'}}),/COMMAND_PAYLOAD_INVALID/);
  await candidate('SAVE',{profile:{cpf:'52998224725',birthDate:'2001-05-01',countryCode:'BR',stateCode:'MT'}});
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
