'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GEAPA_PILOT_MODE:'SHARED_DEV',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),a=require('assert/strict'),{Firestore}=require('@google-cloud/firestore'),crypto=require('node:crypto');
const {manageEditorial}=require('../src/editorial-commands.cjs'),{readAct}=require('../src/act-readers.cjs'),{PROJECT,DEMO,EXECUTOR,ROOT}=require('../src/environment.cjs');
const {fixture}=require('./institutional-fixture.cjs');

test('editorial CMS end-to-end on Firestore Emulator: authz, drafts, preview, publish, history, idempotency, concurrency, revocation and unpublish',{skip:process.env.GEAPA_EDITORIAL_EMULATOR_TEST!=='1'},async t=>{
 const db=new Firestore({projectId:DEMO}),ref=p=>db.doc((ROOT?ROOT+'/':'')+p),now=Date.now(),metrics=[];
 let revoked=false,badApp=false,currentUid='CF_ADMIN';
 const auth={verifyIdToken:async(token,check)=>{a.equal(check,true);if(revoked)throw Object.assign(Error('revoked'),{code:'auth/id-token-revoked'});return {uid:currentUid,sub:currentUid,aud:DEMO,iss:'https://securetoken.google.com/'+DEMO,exp:Math.floor(now/1000)+3600,iat:Math.floor(now/1000),auth_time:Math.floor(now/1000)}}};
 const appCheck={verifyToken:async()=>{if(badApp)throw Error('INVALID');return {appId:'SYN_APP'}}};
 const call=(action,request,hooks={})=>manageEditorial({db,auth,appCheck,token:'synthetic-emulator-only',appCheckToken:'synthetic',origin:'https://geapa-convergencia-dev.web.app',body:{action:'editorial'+action,request},hooks});
 const prepare=async input=>(await call('Preview',{operationId:crypto.randomBytes(16).toString('hex'),input})).result;

 async function seed(role='DIRETOR_COMUNICACAO',group='DIRETORIA',scope='COMUNICACAO'){
  const h=fixture('DIRETOR_ADMINISTRATIVO');
  h.docs['memberRoleCatalog/CARGO-001'].roleKey=role;
  h.docs['memberRoleCatalog/CARGO-001'].group=group;
  h.docs['memberFunctions/VIG-001'].scopeKey=scope;
  const edPolicy={
   ...h.policy,
   id:'EDITORIAL_POLICY',
   actorSelector:{kind:'ROLE',roleKey:role,group,actorScope:scope},
   capabilities:['PUBLIC_CONTENT_MANAGE'],
   fieldMasks:{PUBLIC_CONTENT_MANAGE:['pages','timeline']},
   targets:{...h.policy.targets,allowedTargetScopes:['COMUNICACAO','ADMINISTRACAO','PRESIDENCIA'],allowedResources:['canonicalEditorial/pages','canonicalEditorial/timeline']}
  };
  const edControl={
   schemaVersion:'editorial-control-v1',
   projectId:PROJECT,
   commandType:'editorial.public.manage',
   enabled:true,
   executor:EXECUTOR,
   epoch:1,
   revision:1,
   actorUids:['CF_ADMIN'],
   policyIds:['EDITORIAL_POLICY'],
   availabilityMode:'TIMEBOXED',
   startsAt:now-1000,
   expiresAt:now+3600000,
   rateLimit:{windowMs:3600000,maxAdmissions:60},
   browserEnabled:true,
   appCheck:{mode:'REQUIRED',appIds:['SYN_APP']}
  };
  const recControl={
   schemaVersion:'activity-projection-recovery-v1',
   projectId:PROJECT,
   namespace:ROOT,
   enabled:true,
   revision:1,
   startsAt:now-1000,
   expiresAt:now+3600000,
   policyIds:['EDITORIAL_POLICY'],
   readModelCommands:[],
   browserEnabled:true,
   appCheck:{mode:'REQUIRED',appIds:['SYN_APP']},
   specializedReaders:{enabled:true,actorUids:['CF_ADMIN'],routes:['/act/context','/admin/conteudo-publico']}
  };
  const b=db.batch();
  for(const [path,v]of Object.entries({...h.docs,'memberAuthzPolicies/EDITORIAL_POLICY':edPolicy,'editorialControl/current':edControl,'projectionRecoveryControl/current':recControl}))b.set(ref(path),v);
  for(const del of ['canonicalEditorial/home','canonicalEditorial/about','canonicalEditorial/timeline','canonicalEditorialHistory/home','canonicalEditorialHistory/about','canonicalEditorialHistory/timeline','publicPages/home','publicPages/about','editorialRate/'+crypto.createHash('sha256').update(JSON.stringify('CF_ADMIN')).digest('hex').slice(0,32)])b.delete(ref(del));
  const oldTimeline=await db.collection('canonicalEditorial').get();
  for(const d of oldTimeline.docs)if(d.id.startsWith('timeline_'))b.delete(d.ref);
  const oldPubTimeline=await db.collection('publicTimeline').get();
  for(const d of oldPubTimeline.docs)b.delete(d.ref);
  await b.commit();
 }

 try{
  await t.test('1 & 2. authorized Assessor de Comunicacao bootstraps context and reads /admin/conteudo-publico without writes',async()=>{
   await seed('ASSESSOR_COMUNICACAO_1','ASSESSORIA','COMUNICACAO');
   const ctx=await readAct({db,auth,appCheck,token:'synthetic-emulator-only',appCheckToken:'synthetic',origin:'https://geapa-convergencia-dev.web.app',body:{action:'actRead',request:{route:'/act/context',params:{}}}});
   a.equal(ctx.result.institutionalOnly,true);
   a.equal(ctx.result.navigation.routes['admin-conteudo-publico'],true);
   a.equal(ctx.metrics.writes,0);
   const view=await readAct({db,auth,appCheck,token:'synthetic-emulator-only',appCheckToken:'synthetic',origin:'https://geapa-convergencia-dev.web.app',body:{action:'actRead',request:{route:'/admin/conteudo-publico',params:{}}}});
   a.equal(view.result.data.schemaVersion,'editorial-admin-view-v1');
   a.equal(view.result.data.canManagePages,true);
   a.equal(view.result.data.canManageTimeline,true);
   a.equal(view.metrics.writes,0);
  });

  await t.test('3-8. save draft (private), preview (zero writes), publish, update public, keep history, idempotent replay, unpublish',async()=>{
   await seed();
   const draftPrep=await prepare({action:'PAGE_SAVE_DRAFT',pageId:'about',expectedRevision:0,content:{title:'Sobre o GEAPA',subtitle:'Subtítulo oficial',sections:[{id:'sec_hero',type:'HERO',title:'Nossa Missão',subtitle:'UFMT',text:'Atuação em agroecologia <script>alert(1)</script>',imageUrl:'https://firebasestorage.googleapis.com/v0/b/portal-geapa/o/foto.png',buttonText:'Ver Eixos',buttonUrl:'#/eixos-tematicos',order:1,published:true},{id:'sec_hidden',type:'TEXT',title:'Nota interna rascunho',subtitle:'',text:'Não publicar ainda',imageUrl:'',buttonText:'',buttonUrl:'',order:2,published:false}]}});
   a.equal(draftPrep.preview.publicUpdate,false);
   const draftExec=await call('Execute',draftPrep.request);
   metrics.push(draftExec.metrics);
   a.equal(draftExec.result.state,'COMPLETED');
   a.equal(draftExec.result.revision,1);
   a.equal(draftExec.result.hasDraft,true);
   a.equal(draftExec.result.published,false);
   // Public collection MUST remain absent after draft save
   a.equal((await ref('publicPages/about').get()).exists,false);

   // Preview of PAGE_PUBLISH must execute zero writes and keep publicPages absent
   const pubPrep=await prepare({action:'PAGE_PUBLISH',pageId:'about',expectedRevision:1});
   a.equal(pubPrep.preview.publicUpdate,true);
   a.equal((await ref('publicPages/about').get()).exists,false);

   // Publish v1 -> revision 2
   const pubExec=await call('Execute',pubPrep.request);
   metrics.push(pubExec.metrics);
   a.equal(pubExec.result.state,'COMPLETED');
   a.equal(pubExec.result.published,true);
   a.equal(pubExec.result.revision,2);
   const pubSnap=(await ref('publicPages/about').get()).data();
   a.equal(pubSnap.published,true);
   a.equal(pubSnap.sanitized,true);
   a.equal(pubSnap.revision,2);
   a.equal(pubSnap.sections.length,1); // unpublished section excluded!
   a.equal(pubSnap.sections[0].text,'Atuação em agroecologia');

   // Idempotent replay with same request -> zero writes!
   const replay=await call('Execute',pubPrep.request);
   a.equal(replay.metrics.writes,0);
   a.equal(replay.result.state,'COMPLETED');
   a.equal(replay.result.revision,2);

   // Save draft v2 and publish v2 -> history preserves v1 + v2!
   const d2=await prepare({action:'PAGE_SAVE_DRAFT',pageId:'about',expectedRevision:2,content:{title:'Sobre o GEAPA v2',subtitle:'Atualizado',sections:[{id:'sec_hero',type:'HERO',title:'Missão v2',subtitle:'',text:'Texto v2',imageUrl:'',buttonText:'',buttonUrl:'',order:1,published:true}]}});
   await call('Execute',d2.request);
   const p2=await prepare({action:'PAGE_PUBLISH',pageId:'about',expectedRevision:3});
   await call('Execute',p2.request);
   a.equal((await ref('publicPages/about').get()).data().title,'Sobre o GEAPA v2');
   const hist=(await ref('canonicalEditorialHistory/about').get()).data();
   a.equal(hist.events.length,4);
   a.equal(hist.events[1].revisionAfter,2);
   a.equal(hist.events[1].snapshot.title,'Sobre o GEAPA');
   a.equal(hist.events[3].revisionAfter,4);

   // Unpublish removes publicPages/about completely while retaining draft & history!
   const unpub=await prepare({action:'PAGE_UNPUBLISH',pageId:'about',expectedRevision:4});
   const unpubExec=await call('Execute',unpub.request);
   a.equal(unpubExec.result.state,'COMPLETED');
   a.equal(unpubExec.result.published,false);
   a.equal((await ref('publicPages/about').get()).exists,false);
   a.equal((await ref('canonicalEditorial/about').get()).data().published,false);
   a.equal((await ref('canonicalEditorial/about').get()).data().draft.title,'Sobre o GEAPA v2');
  });

  await t.test('timeline canonical workflow: create, publish, reorder published only, unpublish',async()=>{
   await seed();
   // Create first item (published later)
   const t1=await prepare({action:'TIMELINE_SAVE_DRAFT',itemId:null,expectedRevision:0,expectedIndexRevision:0,content:{year:2020,date:'10/03/2020',title:'Criação',description:'Início do grupo',imageUrl:'',order:1}});
   const r1=await call('Execute',t1.request);
   const id1=r1.result.targetId;
   // Create second item (kept strictly as draft)
   const t2=await prepare({action:'TIMELINE_SAVE_DRAFT',itemId:null,expectedRevision:0,expectedIndexRevision:1,content:{year:2024,date:'2024-08-01',title:'Rascunho Sigiloso',description:'Não publicado',imageUrl:'',order:2}});
   const r2=await call('Execute',t2.request);
   const id2=r2.result.targetId;
   a.equal((await ref('publicTimeline/'+id1).get()).exists,false);
   a.equal((await ref('publicTimeline/'+id2).get()).exists,false);

   // Publish item 1 only
   const pub1=await prepare({action:'TIMELINE_PUBLISH',itemId:id1,expectedRevision:1});
   await call('Execute',pub1.request);
   a.equal((await ref('publicTimeline/'+id1).get()).exists,true);
   a.equal((await ref('publicTimeline/'+id1).get()).data().date,'2020-03-10');

   // Reorder items [id2, id1] -> id1 gets order 2 in publicTimeline, id2 stays unpublished!
   const reord=await prepare({action:'TIMELINE_REORDER',expectedIndexRevision:2,items:[{id:id2,expectedRevision:1},{id:id1,expectedRevision:2}]});
   await call('Execute',reord.request);
   a.equal((await ref('publicTimeline/'+id1).get()).data().order,2);
   a.equal((await ref('publicTimeline/'+id2).get()).exists,false);

   // Unpublish id1 removes publicTimeline/id1
   const unpub=await prepare({action:'TIMELINE_UNPUBLISH',itemId:id1,expectedRevision:3});
   await call('Execute',unpub.request);
   a.equal((await ref('publicTimeline/'+id1).get()).exists,false);
  });

  await t.test('9 & 10. concurrency, lost response, pre-commit abort, revocation and invalid URLs',async()=>{
   await seed();
   // Concurrency on revision 0
   const pA=await prepare({action:'PAGE_SAVE_DRAFT',pageId:'home',expectedRevision:0,content:{title:'Home A',subtitle:'',sections:[{id:'s1',type:'HERO',title:'A',subtitle:'',text:'T',imageUrl:'',buttonText:'',buttonUrl:'',order:1,published:true}]}});
   const pB=await prepare({action:'PAGE_SAVE_DRAFT',pageId:'home',expectedRevision:0,content:{title:'Home B',subtitle:'',sections:[{id:'s1',type:'HERO',title:'B',subtitle:'',text:'T',imageUrl:'',buttonText:'',buttonUrl:'',order:1,published:true}]}});
   const raced=await Promise.allSettled([call('Execute',pA.request),call('Execute',pB.request)]);
   a.equal(raced.filter(x=>x.status==='fulfilled').length,1);
   a.equal(raced.filter(x=>x.status==='rejected').length,1);

   // Lost response after commit + closed admission -> status recovers COMPLETED with 0 writes
   const pLost=await prepare({action:'PAGE_SAVE_DRAFT',pageId:'home',expectedRevision:1,content:{title:'Home Lost',subtitle:'',sections:[{id:'s1',type:'HERO',title:'C',subtitle:'',text:'T',imageUrl:'',buttonText:'',buttonUrl:'',order:1,published:true}]}});
   await a.rejects(call('Execute',pLost.request,{afterCommit:()=>{throw Error('TIMEOUT_LOST_RESPONSE')}}),/TIMEOUT_LOST_RESPONSE/);
   await ref('editorialControl/current').update({enabled:false});
   const rec=await call('Status',{operationId:pLost.request.operationId,digest:pLost.request.digest});
   a.equal(rec.result.state,'COMPLETED');
   a.equal(rec.metrics.writes,0);

   // Failure before commit -> zero state, status NOT_ADMITTED
   await ref('editorialControl/current').update({enabled:true});
   const pPre=await prepare({action:'PAGE_SAVE_DRAFT',pageId:'home',expectedRevision:2,content:{title:'Abort',subtitle:'',sections:[{id:'s1',type:'HERO',title:'D',subtitle:'',text:'T',imageUrl:'',buttonText:'',buttonUrl:'',order:1,published:true}]}});
   await a.rejects(call('Execute',pPre.request,{beforeCommit:()=>{throw Error('ABORT_PRE')}}),/ABORT_PRE/);
   a.equal((await ref('canonicalEditorial/home').get()).data().revision,2);
   a.equal((await call('Status',{operationId:pPre.request.operationId,digest:pPre.request.digest})).result.state,'NOT_ADMITTED');

   // Reject unsafe URL & unauthorized role & policy revocation mid-flight
   await a.rejects(prepare({action:'PAGE_SAVE_DRAFT',pageId:'home',expectedRevision:2,content:{title:'Bad',subtitle:'',sections:[{id:'s1',type:'HERO',title:'D',subtitle:'',text:'T',imageUrl:'javascript:alert(1)',buttonText:'',buttonUrl:'',order:1,published:true}]}}),/COMMAND_PAYLOAD_INVALID/);

   const pRev=await prepare({action:'PAGE_SAVE_DRAFT',pageId:'home',expectedRevision:2,content:{title:'Revoked',subtitle:'',sections:[{id:'s1',type:'HERO',title:'E',subtitle:'',text:'T',imageUrl:'',buttonText:'',buttonUrl:'',order:1,published:true}]}});
   await ref('memberAuthzPolicies/EDITORIAL_POLICY').update({active:false});
   await a.rejects(call('Execute',pRev.request),/AUTHORIZATION_DENIED/);

   await ref('memberAuthzPolicies/EDITORIAL_POLICY').update({active:true});
   revoked=true;await a.rejects(prepare({action:'PAGE_SAVE_DRAFT',pageId:'home',expectedRevision:2,content:{title:'X',subtitle:'',sections:[{id:'s1',type:'HERO',title:'X',subtitle:'',text:'T',imageUrl:'',buttonText:'',buttonUrl:'',order:1,published:true}]}}),/AUTH_BINDING_INVALID/);
   revoked=false;badApp=true;await a.rejects(prepare({action:'PAGE_SAVE_DRAFT',pageId:'home',expectedRevision:2,content:{title:'X',subtitle:'',sections:[{id:'s1',type:'HERO',title:'X',subtitle:'',text:'T',imageUrl:'',buttonText:'',buttonUrl:'',order:1,published:true}]}}),/APP_CHECK_INVALID/);
   badApp=false;

   // Different user cannot recover another owner's status
   currentUid='OTHER_UID';
   await ref('editorialControl/current').update({actorUids:['CF_ADMIN','OTHER_UID']});
   await a.rejects(call('Status',{operationId:pLost.request.operationId,digest:pLost.request.digest}),/AUTHORIZATION_DENIED|ACCESS_REVOKED|M13_AUTHZ_DENIED/);
   currentUid='CF_ADMIN';
  });

  console.log(JSON.stringify({emulatorOnly:true,metrics}));
 }finally{await db.terminate();}
});
