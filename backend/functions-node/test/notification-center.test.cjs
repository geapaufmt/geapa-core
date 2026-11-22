'use strict';
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

test('Central de Notificações - emissão sintética e persistência de notificação interna',()=>{
 const {domain,docs}=setupTestDomain();
 const now=1700000000000;
 const personId='PES-ALICE-001';
 const opId='a'.repeat(32);

 const emitResult=JSON.parse(domain.run('(input)=>{const tx={get:p=>({data:()=>nodeDb.get(p)}),set:(p,v)=>{nodeTx.set(p,v);}};const M=f4MemberNotifications_();const res=M.emit(tx,{personId:input.personId,event:"SYNTHETIC_TEST",title:"Teste Local",body:"Mensagem sintética de teste local.",targetRoute:"/app/minha-situacao",category:"OPERACIONAL",operationId:input.opId},input.now);return res;}',{personId,opId,now}));

 a.equal(emitResult.replay,false);
 a.equal(emitResult.notification.id,opId);
 a.equal(emitResult.notification.personId,personId);
 a.equal(emitResult.notification.read,false);
 a.equal(emitResult.notification.readAt,null);
 a.equal(emitResult.notification.title,'Teste Local');
 a.equal(emitResult.notification.body,'Mensagem sintética de teste local.');
 a.equal(emitResult.notification.targetRoute,'/app/minha-situacao');
 a.equal(emitResult.notification.category,'OPERACIONAL');
 a.equal(emitResult.index.unreadCount,1);
 a.deepEqual(emitResult.index.notificationIds,[opId]);

 // Persisted documents in docs
 a(docs['memberNotifications/'+opId]);
 a(docs['memberNotificationIndexes/'+personId]);
 a.equal(docs['memberNotifications/'+opId].read,false);
 a.equal(docs['memberNotificationIndexes/'+personId].unreadCount,1);
});

test('Central de Notificações - isolamento rigoroso entre usuários (Alice vs Bob)',()=>{
 const now=1700000000000;
 const alice='PES-ALICE-001',bob='PES-BOB-002';
 const notifAlice='a'.repeat(32);
 const notifBob='b'.repeat(32);

 const initialDocs={
  'memberAuthBindings/UID_ALICE':{idPessoa:alice,active:true},
  'memberAuthBindings/UID_BOB':{idPessoa:bob,active:true}
 };

 const {domain,docs}=setupTestDomain(initialDocs);

 // Emit notification for Alice and for Bob
 domain.run('(input)=>{const tx={get:p=>({data:()=>nodeDb.get(p)}),set:(p,v)=>{nodeTx.set(p,v);}};const M=f4MemberNotifications_();M.emit(tx,{personId:input.alice,event:"SYNTHETIC_TEST",title:"Alice Notif",body:"Para Alice",operationId:input.notifAlice},input.now);M.emit(tx,{personId:input.bob,event:"SYNTHETIC_TEST",title:"Bob Notif",body:"Para Bob",operationId:input.notifBob},input.now);}',{alice,bob,notifAlice,notifBob,now});

 // Alice reads own notifications
 const aliceRead=JSON.parse(domain.run('(input)=>{const get=p=>nodeDb.get(p);const M=f4MemberNotifications_();return M.readAuthorized(get,{uid:"UID_ALICE",expiresAt:input.now+60000},input.now);}',{now}));
 a.equal(aliceRead.personId,alice);
 a.equal(aliceRead.unreadCount,1);
 a.equal(aliceRead.items.length,1);
 a.equal(aliceRead.items[0].id,notifAlice);
 a.equal(aliceRead.items[0].title,'Alice Notif');

 // Bob reads own notifications
 const bobRead=JSON.parse(domain.run('(input)=>{const get=p=>nodeDb.get(p);const M=f4MemberNotifications_();return M.readAuthorized(get,{uid:"UID_BOB",expiresAt:input.now+60000},input.now);}',{now}));
 a.equal(bobRead.personId,bob);
 a.equal(bobRead.unreadCount,1);
 a.equal(bobRead.items.length,1);
 a.equal(bobRead.items[0].id,notifBob);
 a.equal(bobRead.items[0].title,'Bob Notif');

 // Regular user trying to read Alice's notifications throws ACCESS_REVOKED
 a.throws(()=>{
  domain.run('(input)=>{const get=p=>nodeDb.get(p);const M=f4MemberNotifications_();return M.readAuthorized(get,{uid:"UID_BOB",expiresAt:input.now+60000},input.now,{personId:input.alice});}',{now,alice});
 },/ACCESS_REVOKED/);
});

test('Central de Notificações - marcação de leitura individual e idempotência',()=>{
 const now=1700000000000;
 const person='PES-TEST-001';
 const notifId='c'.repeat(32);
 const initialDocs={
  'memberAuthBindings/UID_USER':{idPessoa:person,active:true}
 };

 const {domain,docs}=setupTestDomain(initialDocs);

 // 1. Emit notification
 domain.run('(input)=>{const tx={get:p=>({data:()=>nodeDb.get(p)}),set:(p,v)=>{nodeTx.set(p,v);}};const M=f4MemberNotifications_();M.emit(tx,{personId:input.person,event:"SYNTHETIC_TEST",title:"Leitura",body:"Corpo",operationId:input.notifId},input.now);}',{person,notifId,now});
 a.equal(docs['memberNotificationIndexes/'+person].unreadCount,1);
 a.equal(docs['memberNotifications/'+notifId].read,false);

 // 2. First markRead: updates read to true, sets readAt, decrements unreadCount to 0
 const mark1=JSON.parse(domain.run('(input)=>{const tx={get:p=>({data:()=>nodeDb.get(p)}),set:(p,v)=>{nodeTx.set(p,v);}};const M=f4MemberNotifications_();return M.markRead(tx,{uid:"UID_USER",expiresAt:input.now+60000},input.notifId,input.now+500);}',{notifId,now}));

 a.equal(mark1.ok,true);
 a.equal(mark1.replay,false);
 a.equal(mark1.unreadCount,0);
 a.equal(mark1.notification.read,true);
 a.equal(mark1.notification.readAt,now+500);
 a.equal(docs['memberNotificationIndexes/'+person].unreadCount,0);
 a.equal(docs['memberNotifications/'+notifId].read,true);

 // 3. Second markRead on same notification: IDEMPOTENT (replay: true, unreadCount stays 0)
 const mark2=JSON.parse(domain.run('(input)=>{const tx={get:p=>({data:()=>nodeDb.get(p)}),set:(p,v)=>{nodeTx.set(p,v);}};const M=f4MemberNotifications_();return M.markRead(tx,{uid:"UID_USER",expiresAt:input.now+60000},input.notifId,input.now+1000);}',{notifId,now}));

 a.equal(mark2.ok,true);
 a.equal(mark2.replay,true);
 a.equal(mark2.unreadCount,0);
 a.equal(mark2.notification.read,true);
 a.equal(mark2.notification.readAt,now+500); // preserves original readAt
 a.equal(docs['memberNotificationIndexes/'+person].unreadCount,0);

 // 4. Attempt to mark other person's notification throws ACCESS_REVOKED
 docs['memberAuthBindings/UID_INTRUDER']={idPessoa:'PES-OTHER',active:true};
 a.throws(()=>{
  domain.run('(input)=>{const tx={get:p=>({data:()=>nodeDb.get(p)}),set:(p,v)=>{nodeTx.set(p,v);}};const M=f4MemberNotifications_();return M.markRead(tx,{uid:"UID_INTRUDER",expiresAt:input.now+60000},input.notifId,input.now+2000);}',{notifId,now});
 },/ACCESS_REVOKED/);
});

test('Central de Notificações - marcação em lote (markAllRead) e zera contador',()=>{
 const now=1700000000000;
 const person='PES-TEST-002';
 const initialDocs={
  'memberAuthBindings/UID_USER2':{idPessoa:person,active:true}
 };

 const {domain,docs}=setupTestDomain(initialDocs);

 const id1='1'.repeat(32);
 const id2='2'.repeat(32);
 const id3='3'.repeat(32);

 domain.run('(input)=>{const tx={get:p=>({data:()=>nodeDb.get(p)}),set:(p,v)=>{nodeTx.set(p,v);}};const M=f4MemberNotifications_();M.emit(tx,{personId:input.person,event:"SYNTHETIC_TEST",title:"N1",operationId:input.id1},input.now);M.emit(tx,{personId:input.person,event:"SYNTHETIC_TEST",title:"N2",operationId:input.id2},input.now);M.emit(tx,{personId:input.person,event:"SYNTHETIC_TEST",title:"N3",operationId:input.id3},input.now);}',{person,id1,id2,id3,now});

 a.equal(docs['memberNotificationIndexes/'+person].unreadCount,3);

 // Mark all read
 const markAllResult=JSON.parse(domain.run('(input)=>{const tx={get:p=>({data:()=>nodeDb.get(p)}),set:(p,v)=>{nodeTx.set(p,v);}};const M=f4MemberNotifications_();return M.markAllRead(tx,{uid:"UID_USER2",expiresAt:input.now+60000},input.now+100);}',{now}));

 a.equal(markAllResult.ok,true);
 a.equal(markAllResult.unreadCount,0);
 a.equal(markAllResult.updated,3);
 a.equal(docs['memberNotificationIndexes/'+person].unreadCount,0);
 a.equal(docs['memberNotifications/'+id1].read,true);
 a.equal(docs['memberNotifications/'+id2].read,true);
 a.equal(docs['memberNotifications/'+id3].read,true);

 // Idempotency of markAllRead: calling again returns replay: true
 const markAllAgain=JSON.parse(domain.run('(input)=>{const tx={get:p=>({data:()=>nodeDb.get(p)}),set:(p,v)=>{nodeTx.set(p,v);}};const M=f4MemberNotifications_();return M.markAllRead(tx,{uid:"UID_USER2",expiresAt:input.now+60000},input.now+200);}',{now}));
 a.equal(markAllAgain.replay,true);
 a.equal(markAllAgain.unreadCount,0);
 a.equal(markAllAgain.updated,0);
});

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
});

test('Central de Notificações - rota operacional /meu-vinculo/notificacoes do Portal',()=>{
 const now=1700000000000;
 const person='PES-TEST-005';
 const notifId='d'.repeat(32);
 const initialDocs={
  'memberAuthBindings/UID_PORTAL_USER':{idPessoa:person,active:true}
 };

 const {domain,docs}=setupTestDomain(initialDocs);

 // 1. Emit notification
 domain.run('(input)=>{const tx={get:p=>({data:()=>nodeDb.get(p)}),set:(p,v)=>{nodeTx.set(p,v);}};const M=f4MemberNotifications_();M.emit(tx,{personId:input.person,event:"SYNTHETIC_TEST",title:"Portal Route Test",body:"Testando rota do portal",operationId:input.notifId},input.now);}',{person,notifId,now});

 // 2. Read via operational reader /meu-vinculo/notificacoes
 const readResult=JSON.parse(domain.run('(input)=>{const nodeConfig={projectId:"geapa-convergencia-dev",environment:"DEV",database:"(default)",host:"firestore.googleapis.com",port:443};const reader=f4MemberOperationalReader_({config:nodeConfig,identity:{uid:"UID_PORTAL_USER",expiresAt:input.now+60000},clock:()=>input.now});return reader.read({route:"/meu-vinculo/notificacoes",params:{}});}',{now}));

 a.equal(readResult.ok,true);
 a.equal(readResult.data.personId,person);
 a.equal(readResult.data.unreadCount,1);
 a.equal(readResult.data.items.length,1);
 a.equal(readResult.data.items[0].id,notifId);
 a.equal(readResult.data.items[0].read,false);

 // 3. Mark read via operational reader /meu-vinculo/notificacoes
 const markResult=JSON.parse(domain.run('(input)=>{const nodeConfig={projectId:"geapa-convergencia-dev",environment:"DEV",database:"(default)",host:"firestore.googleapis.com",port:443};const reader=f4MemberOperationalReader_({config:nodeConfig,identity:{uid:"UID_PORTAL_USER",expiresAt:input.now+60000},clock:()=>input.now});return reader.read({route:"/meu-vinculo/notificacoes",params:{action:"MARK_READ",notificationId:input.notifId}});}',{notifId,now}));

 a.equal(markResult.ok,true);
 a.equal(markResult.unreadCount,0);
 a.equal(markResult.notification.read,true);

 // 4. Read again: verified unreadCount is 0
 const readAfter=JSON.parse(domain.run('(input)=>{const nodeConfig={projectId:"geapa-convergencia-dev",environment:"DEV",database:"(default)",host:"firestore.googleapis.com",port:443};const reader=f4MemberOperationalReader_({config:nodeConfig,identity:{uid:"UID_PORTAL_USER",expiresAt:input.now+60000},clock:()=>input.now});return reader.read({route:"/meu-vinculo/notificacoes",params:{}});}',{now}));
 a.equal(readAfter.data.unreadCount,0);
 a.equal(readAfter.data.items[0].read,true);
});

