'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GEAPA_PILOT_MODE:'SHARED_DEV',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),a=require('node:assert/strict'),crypto=require('node:crypto');
const {Timestamp}=require('firebase-admin/firestore');
const {createDomain}=require('../src/domain.cjs');
const {decode,encode}=require('../src/firestore.cjs');

function setupTestDomain(initialDocs = {}){
 const docs={...initialDocs};
 const reads=[],writes=[];
 const domain=createDomain(raw=>{
  const x=JSON.parse(raw);
  if(x.kind==='get'){
   reads.push(x.path);
   return JSON.stringify(docs[x.path]!==undefined?encode(docs[x.path]):null);
  }
  if(x.kind==='write'){
   writes.push(x);
   if(x.mode==='delete')delete docs[x.path];
   else docs[x.path]=decode(x.data);
   return 'null';
  }
  if(x.kind==='getAll')return 'null';
  if(['inspect','guard'].includes(x.kind))return 'null';
  if(x.kind==='pilotConfig')return JSON.stringify({root:'',executor:'CANONICAL_TEST',settings:{commands:[],policyIds:[]}});
  throw Error('UNEXPECTED_IO_'+x.kind);
 });
 return {domain,docs,reads,writes};
}

test('Central de Notificações - e-mail simulado: prévia institucional e recibo DRY_RUN sem envio real',()=>{
 const now=1700000000000;
 const person='PES-TEST-003';
 const opId='e'.repeat(32);
 const sourceDigest=crypto.createHash('sha256').update(JSON.stringify({opId})).digest('hex');

 const initialDocs={
  'coreDevAdmins/CF_ADMIN':{schemaVersion:'core-dev-admin-v1',projectId:'geapa-convergencia-dev',environment:'DEV',uid:'CF_ADMIN',active:true,revision:1,issuedAt:Timestamp.fromMillis(now-1000),expiresAt:Timestamp.fromMillis(now+60000),permissions:'FUNCTIONAL_ALL',exportable:false},
  'memberOperationalConfig/notifications':{schemaVersion:'member-notification-policy-v1',enabled:true,mode:'DRY_RUN',projectId:'geapa-convergencia-dev',environment:'DEV',revision:1,approvalRef:'SYN_APPROVAL',presidencyEmail:'presidencia@geapa.invalid'},
  'memberOperationalConfig/mailBrand':{schemaVersion:'member-mail-brand-v1',revision:1,approvalRef:'SYN_APPROVAL',values:{orgName:'Grupo de Estudos de Apoio à Produção Agrícola',shortName:'GEAPA',quote:'"Cultivar o Conhecimento Para Colher Sabedoria"',officialEmail:'contato@geapa.invalid',universityName:'Universidade Federal de Mato Grosso',universityShortName:'UFMT',instituteName:'Instituto de Ciências Agrárias e Ambientais',instituteShortName:'ICAA',courseName:'Agronomia',greenColor:'#167d0a',brownColor:'#5c371a',blackColor:'#163226'}},
  ['memberAggregateHeads/'+person]:{schemaVersion:'m11-head-v1',id:person,kind:'PERSON',revision:1,paths:['memberIdentifiers/ID_EMAIL']},
  'memberIdentifiers/ID_EMAIL':{schemaVersion:'m1-memberIdentifiers-v1',id:'ID_EMAIL',idPessoa:person,revision:1,kind:'EMAIL',value:'membro@geapa.invalid',active:true},
  ['jobOutbox/'+opId]:{schemaVersion:'member-notification-outbox-v1',operationId:opId,commandType:'members.voluntary.submit',kind:'MEMBERS_NOTIFICATION',event:'VOLUNTARY_RECEIVED',state:'PENDING',idempotencyKey:opId,attempts:0,personId:person,enrollmentId:null,requestId:null,policyRevision:1,projectId:'geapa-convergencia-dev',environment:'DEV',mode:'DRY_RUN',sourceDigest},
  ['commandRequests/'+opId]:{operationId:opId,digest:sourceDigest,commandType:'members.voluntary.submit',target:{personId:person}},
  ['commandLedger/'+opId]:{operationId:opId,state:'COMPLETED',digest:sourceDigest},
  ['memberAudit/'+opId]:{operationId:opId,digest:sourceDigest}
 };

 const {domain,docs}=setupTestDomain(initialDocs);

 // 1. Preview generation (renders HTML template via legacy/19_core_mail_renderer.gs)
 const previewResult=JSON.parse(domain.run('(input)=>{const nodeConfig={projectId:"geapa-convergencia-dev",environment:"DEV",database:"(default)",host:"firestore.googleapis.com",port:443};const M=f4MemberNotifications_();const c=M.consumer({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:input.now+60000},clock:()=>input.now});return c.preview(input.opId);}',{opId,now}));

 a.equal(previewResult.operationId,opId);
 a.equal(previewResult.recipient,'membro@geapa.invalid');
 a.equal(previewResult.mode,'DRY_RUN');
 a.equal(previewResult.isRealMail,false);
 a(typeof previewResult.rendered.htmlBody==='string');
 a(previewResult.rendered.htmlBody.includes('GEAPA'));
 a(previewResult.rendered.htmlBody.includes('Solicitação recebida'));
 a(previewResult.rendered.htmlBody.includes('Cultivar o Conhecimento Para Colher Sabedoria'));
 a(previewResult.rendered.htmlBody.includes('Mensagem de teste. Nenhum e-mail foi enviado.'));

 // 2. Simulation run (generates receipt without real network send)
 const runResult=JSON.parse(domain.run('(input)=>{const nodeConfig={projectId:"geapa-convergencia-dev",environment:"DEV",database:"(default)",host:"firestore.googleapis.com",port:443};const M=f4MemberNotifications_();const c=M.consumer({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:input.now+60000},clock:()=>input.now});return c.run(input.opId);}',{opId,now}));

 a.equal(runResult.mode,'DRY_RUN');
 a.equal(runResult.sent,false);
 a.equal(runResult.operationId,opId);
 a.equal(runResult.event,'VOLUNTARY_RECEIVED');
 a.equal(runResult.channel,'SIMULATED_EMAIL');
 a.equal(runResult.replay,false);
 a(runResult.contentDigest);
 a(runResult.recipientDigest);

 // 3. Repeated run: Idempotent replay
 const replayResult=JSON.parse(domain.run('(input)=>{const nodeConfig={projectId:"geapa-convergencia-dev",environment:"DEV",database:"(default)",host:"firestore.googleapis.com",port:443};const M=f4MemberNotifications_();const c=M.consumer({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:input.now+60000},clock:()=>input.now});return c.run(input.opId);}',{opId,now}));
 a.equal(replayResult.replay,true);
 a.equal(replayResult.sent,false);
 a.equal(replayResult.mode,'DRY_RUN');
});

test('Central de Notificações - simulação controlada de falha do canal externo',()=>{
 const now=1700000000000;
 const person='PES-TEST-004';
 const opId='f'.repeat(32);
 const sourceDigest=crypto.createHash('sha256').update(JSON.stringify({opId})).digest('hex');

 const initialDocs={
  'coreDevAdmins/CF_ADMIN':{schemaVersion:'core-dev-admin-v1',projectId:'geapa-convergencia-dev',environment:'DEV',uid:'CF_ADMIN',active:true,revision:1,issuedAt:Timestamp.fromMillis(now-1000),expiresAt:Timestamp.fromMillis(now+60000),permissions:'FUNCTIONAL_ALL',exportable:false},
  'memberOperationalConfig/notifications':{schemaVersion:'member-notification-policy-v1',enabled:true,mode:'DRY_RUN',projectId:'geapa-convergencia-dev',environment:'DEV',revision:1,approvalRef:'SYN_APPROVAL',presidencyEmail:'presidencia@geapa.invalid'},
  ['memberAggregateHeads/'+person]:{schemaVersion:'m11-head-v1',id:person,kind:'PERSON',revision:1,paths:['memberIdentifiers/ID_EMAIL']},
  'memberIdentifiers/ID_EMAIL':{schemaVersion:'m1-memberIdentifiers-v1',id:'ID_EMAIL',idPessoa:person,revision:1,kind:'EMAIL',value:'membro@geapa.invalid',active:true},
  ['jobOutbox/'+opId]:{schemaVersion:'member-notification-outbox-v1',operationId:opId,commandType:'members.voluntary.submit',kind:'MEMBERS_NOTIFICATION',event:'VOLUNTARY_RECEIVED',state:'PENDING',idempotencyKey:opId,attempts:0,personId:person,enrollmentId:null,requestId:null,policyRevision:1,projectId:'geapa-convergencia-dev',environment:'DEV',mode:'DRY_RUN',sourceDigest},
  ['commandRequests/'+opId]:{operationId:opId,digest:sourceDigest,commandType:'members.voluntary.submit',target:{personId:person}},
  ['commandLedger/'+opId]:{operationId:opId,state:'COMPLETED',digest:sourceDigest},
  ['memberAudit/'+opId]:{operationId:opId,digest:sourceDigest}
 };

 const {domain,docs}=setupTestDomain(initialDocs);

 // Run with simulated failure
 const failureResult=JSON.parse(domain.run('(input)=>{const nodeConfig={projectId:"geapa-convergencia-dev",environment:"DEV",database:"(default)",host:"firestore.googleapis.com",port:443};const M=f4MemberNotifications_();const c=M.consumer({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:input.now+60000},clock:()=>input.now});return c.run(input.opId,{simulateFailure:true,failureCode:"SIMULATED_TRANSIENT_OUTAGE",failureMessage:"Queda temporária simulada no gateway de e-mail."});}',{opId,now}));

 a.equal(failureResult.state,'RECONCILIATION_REQUIRED');
 a.equal(failureResult.error,'SIMULATED_TRANSIENT_OUTAGE');
 a.equal(failureResult.errorMessage,'Queda temporária simulada no gateway de e-mail.');
 a.equal(failureResult.mode,'DRY_RUN');
 a.equal(failureResult.sent,false);
 const resolution='LOCAL_SIMULATION_CONFIRMED_NO_RESEND';
 Object.assign(docs['jobOutbox/'+opId],{state:'COMPLETED',resolution});
 const replay=JSON.parse(domain.run('x=>f4MemberNotifications_().consumer({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:x.now+60000},clock:()=>x.now}).run(x.opId)',{now,opId}));
 a.equal(replay.resolution,resolution);a.equal(replay.sent,false);a.equal(replay.replay,true);
});

function notificationFixture(){
 const {seed}=require('./fixtures.cjs'),{now,docs:raw}=seed();
 const initial=Object.fromEntries(Object.entries(raw).map(([key,v])=>[key,decode({mapValue:v})]));
 const uid='CF_ADMIN',person=initial['memberAuthBindings/'+uid].idPessoa;
 delete initial['coreDevAdmins/'+uid];
 const policy=initial['memberAuthorityControl/MEMBERS'].selfReadPolicy;
 initial['memberAuthzPolicies/'+policy.policyId]={schemaVersion:'member-authz-policy-v13',id:policy.policyId,revision:policy.policyRevision,active:true,origin:'F4_M621_INSTITUTIONAL',approvalRef:'SYN_SELF',issuedAt:policy.issuedAt,expiresAt:policy.expiresAt,proofKind:'SELF',capabilities:['MEMBERS_SELF_READ'],actorSelector:{kind:'SELF'},targets:{ownPerson:true,allowedTargetScopes:[],allowedBoards:[],allowedRoles:[],allowedPersons:[],allowedResources:['memberCurrentStates/'+person],commandScopes:['OWN_MEMBER_RESOURCES']},fieldMasks:policy.fieldMasks};
 const {PROJECT,DEMO}=require('../src/environment.cjs'),route='/meu-vinculo/notificacoes',writeRoute=route+'/leitura';
 initial['projectionRecoveryControl/current']={schemaVersion:'activity-projection-recovery-v1',projectId:PROJECT,namespace:'',mode:'DEV_READERS_OPERATIONAL',enabled:true,revision:1,startsAt:now-1000,expiresAt:now+3600000,actorUids:[uid],policyIds:[policy.policyId],readModelCommands:[],browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['SYN_APP']},specializedReaders:{enabled:true,actorUids:[uid],routes:[route,writeRoute]}};
 const {domain,docs}=setupTestDomain(initial);
 const emit=(n,to=person,title='Teste')=>JSON.parse(domain.run('x=>f4MemberNotifications_().emit(nodeTx,x.payload,x.now)',{payload:{personId:to,title,event:'SYNTHETIC_TEST',operationId:n.toString(16).padStart(32,'0')},now}));
 const seconds=Math.floor(now/1000);
 const auth={verifyIdToken:async(token,revoked)=>{a.equal(revoked,true);return {uid,sub:uid,aud:DEMO,iss:'https://securetoken.google.com/'+DEMO,iat:seconds,exp:seconds+3600,auth_time:seconds}}};
 const appCheck={verifyToken:async()=>({appId:'SYN_APP'})};
 const read=async(db,params={},target=route,appCheckToken='synthetic')=>(await require('../src/act-readers.cjs').readAct({db,auth,appCheck,token:'synthetic',appCheckToken,origin:'https://geapa-convergencia-dev.web.app',body:{action:'actRead',request:{route:target,params}}})).result.data;
 return {docs,person,uid,policy,emit,read,route,writeRoute};
}

test('recipient-scoped delivery, stable replay, retention and legacy counter repair',()=>{
 const {emit,person,docs}=notificationFixture();
 const first=emit(1);
 a.equal(emit(1).replay,true);
 a.throws(()=>emit(1,person,'Conteúdo alterado'),/OPERATION_CONFLICT/);
 a.notEqual(emit(1,'SYN_OTHER').notification.id,first.notification.id);
 for(let n=2;n<=100;n++)emit(n);
 const ix=docs['memberNotificationIndexes/'+person];
 ix.schemaVersion='member-notification-index-v1';ix.unreadCount=124;
 const last=emit(101);
 a.equal(last.index.unreadCount,100);a.equal(last.index.notificationIds.length,100);
 a(!last.index.notificationIds.includes(first.notification.id));a.equal(emit(1).replay,true);
});

test('reader and write guard enforce SELF, atomic batch, response contract and stable individual replay',async()=>{
 const {docs,uid,person,emit,read,writeRoute,route,policy}=notificationFixture();
 const db=require('./fixtures/member-memory-db.cjs').database(docs);
 const first=emit(1).notification;emit(2);emit(3);
 const single={action:'MARK_READ',notificationId:first.id},bulk={action:'MARK_ALL_READ',expectedRevision:3};
 a.equal((await read(db)).unreadCount,3);
 await a.rejects(read(db,single,route),/SCHEMA_INVALID/);
 await a.rejects(read(db,{},route,null),/APP_CHECK_REQUIRED/);
 a.equal((await read(db,single,writeRoute)).unreadCount,2);
 const at=docs['memberNotifications/'+first.id].readAt;
 a.equal((await read(db,single,writeRoute)).unreadCount,2);a.equal(docs['memberNotifications/'+first.id].readAt,at);
 await a.rejects(read(db,bulk,writeRoute),/REVISION_CONFLICT/);
 bulk.expectedRevision=(await read(db)).revision;
 a.equal((await read(db,bulk,writeRoute)).unreadCount,0);
 a.equal((await read(db,bulk,writeRoute)).unreadCount,0);
 a((await read(db)).items.every(row=>row.read&&row.targetRoute==='minha-situacao'));
 const foreign=emit(1,'SYN_OTHER').notification;
 await a.rejects(read(db,{action:'MARK_READ',notificationId:foreign.id},writeRoute),/RESOURCE_NOT_FOUND/);
 docs['memberNotificationIndexes/'+person].notificationIds.push(foreign.id);
 await a.rejects(read(db,bulk,writeRoute),/RECONCILIATION_REQUIRED/);
 a.equal(docs['memberNotifications/'+foreign.id].read,false);
 docs['memberNotificationIndexes/'+person].notificationIds.pop();
 docs['memberAuthzPolicies/'+policy.policyId].active=false;
 await a.rejects(read(db));
 docs['memberAuthzPolicies/'+policy.policyId].active=true;
 docs['memberAuthBindings/'+uid].active=false;
 await a.rejects(read(db));await a.rejects(read(db,bulk,writeRoute));
});

test('concurrent notification receipts in a native Firestore emulator transaction',{skip:process.env.GEAPA_NOTIFICATION_EMULATOR_TEST!=='1'},async t=>{
 const {docs,emit,read,writeRoute}=notificationFixture(),id=emit(1).notification.id;emit(2);
 const {Firestore}=require('@google-cloud/firestore'),{DEMO}=require('../src/environment.cjs');
 const db=new Firestore({projectId:DEMO,databaseId:'notice-'+crypto.randomUUID().slice(0,8)});
 t.after(()=>db.terminate());
 const batch=db.batch();for(const [path,value] of Object.entries(docs))batch.set(db.doc(path),value);await batch.commit();
 const mark=()=>read(db,{action:'MARK_READ',notificationId:id},writeRoute);
 await Promise.all([mark(),mark()]);a.equal((await read(db)).unreadCount,1);
 a.equal((await read(db,{action:'MARK_ALL_READ',expectedRevision:(await read(db)).revision},writeRoute)).unreadCount,0);
});
