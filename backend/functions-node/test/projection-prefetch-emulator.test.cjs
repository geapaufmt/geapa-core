'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),a=require('node:assert/strict'),{Firestore}=require('@google-cloud/firestore'),crypto=require('node:crypto');
const {recoverProjection,RecoveryStore}=require('../src/projection-recovery.cjs'),{ROOT,PROJECT,DEMO}=require('../src/environment.cjs'),{agenda}=require('./agenda-matrix-fixture.cjs');
test('projection path prefetch: fresh transactions, parity, recovery and isolation',{skip:process.env.GEAPA_PROJECTION_EMULATOR_TEST!=='1'},async t=>{
 const db=new Firestore({projectId:DEMO}),ref=p=>db.doc(ROOT+'/'+p),now=Date.now(),hash=x=>crypto.createHash('sha256').update(x).digest('hex');
 const view='portalReaderViews/CF_ADMIN/items/activities',lease='projectionRecoveryLeases/'+hash('CF_ADMIN|activities'),budget='projectionRecoveryBudgets/'+hash('CF_ADMIN');
 let revoked=false,badApp=false;
 const auth={verifyIdToken:async(token,check)=>{a.equal(check,true);if(revoked)throw Error('revoked');return {uid:'CF_ADMIN',sub:'CF_ADMIN',aud:DEMO,iss:'https://securetoken.google.com/'+DEMO,iat:Math.floor(now/1000),auth_time:Math.floor(now/1000),exp:Math.floor(now/1000)+3600}}};
 const appCheck={verifyToken:async()=>{if(badApp)throw Error('invalid');return {appId:'SYN_APP'}}};
 const call=(hooks={},action='projectionRecover')=>recoverProjection({db,auth,appCheck,token:'synthetic-emulator-only',appCheckToken:'synthetic',origin:'https://geapa-convergencia-dev.web.app',body:{action,request:{viewId:'activities'}},hooks});
 async function seed(){const docs=agenda('MEMBRO_EFETIVO').docs,base=docs['canonicalActivities/FUTURE'];for(let i=0;i<60;i++){const id='PERF_'+String(i).padStart(3,'0');docs['canonicalActivities/'+id]={...base,id};}
  docs['projectionRecoveryControl/current']={schemaVersion:'activity-projection-recovery-v1',projectId:PROJECT,namespace:ROOT,enabled:true,revision:1,startsAt:now-1000,expiresAt:now+3600000,policyIds:['CF_DEV_ACT_POLICY'],readModelCommands:[],browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['SYN_APP']}};
  const b=db.batch();for(const [p,d]of Object.entries(docs))b.set(ref(p),d);for(const p of [view,lease,budget,'portalReaderResourceFences/activities'])b.delete(ref(p));await b.commit();return docs;
 }
 const facts=async()=>JSON.stringify((await db.collection(ROOT+'/canonicalActivities').orderBy('__name__').get()).docs.map(d=>[d.id,d.data()]));
 const data=async()=>{const d=(await ref(view).get()).data().data;delete d.bundle.data.ultimaAtualizacao;return d;};
 try{
 await t.test('same source, same output and reads; fewer sequential SDK calls',async()=>{
  await seed();const oldGuard=RecoveryStore.prototype.guard;
  let before;try{RecoveryStore.prototype.guard=function(){this.publishReadPaths=[];return oldGuard.call(this)};before=await call();}finally{RecoveryStore.prototype.guard=oldGuard;}
  const expected=await data(),beforeFacts=await facts();await ref(view).delete();await ref(lease).delete();await ref(budget).delete();
  const after=await call();a.deepEqual(await data(),expected);a.equal(await facts(),beforeFacts);a.equal(after.metrics.reads,before.metrics.reads);a.equal(after.metrics.writes,4);a(after.metrics.sdkCalls<=before.metrics.sdkCalls-5);a(after.metrics.discoveryRounds<before.metrics.discoveryRounds);
  console.log(JSON.stringify({emulatorOnly:true,before:before.metrics,after:after.metrics,rows:expected.bundle.data.calendario.length}));
 });
 await t.test('current view and lost response resolve without a new build',async()=>{await seed();await a.rejects(call({afterPublish:()=>{throw Error('LOST_RESPONSE')}}),/LOST_RESPONSE/);const r=await call({},'projectionStatus');a.equal(r.state,'CURRENT');a.equal(r.metrics.writes,0);a.equal((await ref(budget).get()).data().count,1);});
 await t.test('failure after claim keeps the identifiable lease, expiry recovers',async()=>{await seed();await a.rejects(call({afterClaim:()=>{throw Error('STOP_AFTER_CLAIM')}}),/STOP_AFTER_CLAIM/);a.equal((await call()).state,'REBUILDING');a.equal((await ref(view).get()).exists,false);await ref(lease).update({expiresAt:Date.now()-1});a.equal((await call()).state,'CURRENT');a.equal((await ref(budget).get()).data().count,2);});
 await t.test('simultaneous recovery has one publisher',async()=>{await seed();const results=await Promise.all([call(),call()]);a(results.some(r=>r.state==='CURRENT'));a(results.every(r=>['CURRENT','REBUILDING'].includes(r.state)));a.equal((await ref(budget).get()).data().count,1);a.equal((await ref(lease).get()).data().state,'CURRENT');});
 await t.test('policy revoked between transactions is not reused',async()=>{await seed();await a.rejects(call({afterClaim:()=>ref('memberAuthzPolicies/CF_DEV_ACT_POLICY').update({active:false})}),/AUTHORIZATION_DENIED|ACCESS_REVOKED/);a.equal((await ref(view).get()).exists,false);});
 await t.test('changed binding between transactions is not reused',async()=>{await seed();await a.rejects(call({afterClaim:()=>ref('memberAuthBindings/CF_ADMIN').delete()}));a.equal((await ref(view).get()).exists,false);});
 await t.test('changed source is freshly read and source fence published',async()=>{await seed();await call({afterClaim:async()=>{await ref('canonicalActivities/PERF_000').update({title:'Updated synthetic title',revision:2});await ref('portalReaderResourceFences/activities').set({revision:2});}});const v=(await ref(view).get()).data();a.equal(v.basis.dataFence.revision,2);a.equal(v.data.bundle.data.calendario.find(r=>r.idAtividade==='PERF_000').tituloPublico,'Updated synthetic title');});
 await t.test('App Check/Auth denial, private fields and other UID remain protected',async()=>{await seed();badApp=true;await a.rejects(call(),/APP_CHECK_INVALID/);badApp=false;revoked=true;await a.rejects(call(),/AUTH_BINDING_INVALID/);revoked=false;await call();const v=await data();a(!JSON.stringify(v).includes('PRIVATE_SENTINEL'));a.equal((await ref('portalReaderViews/OTHER/items/activities').get()).exists,false);});
 }finally{await db.terminate();}
});

