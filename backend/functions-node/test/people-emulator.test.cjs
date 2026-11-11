'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),a=require('assert/strict'),{Firestore}=require('@google-cloud/firestore'),{readAct}=require('../src/act-readers.cjs'),{ROOT,PROJECT,DEMO}=require('../src/environment.cjs'),{setup}=require('./people-fixture.cjs');
const enabled=process.env.GEAPA_PEOPLE_EMULATOR_TEST==='1';
test('Pessoas Firestore emulator: concurrent pages, current masks, revoke before return and recovery',{skip:!enabled},async()=>{
 const db=new Firestore({projectId:DEMO}),h=setup('DIRETOR_RECURSOS_HUMANOS',99),now=Date.now(),ref=p=>db.doc(ROOT+'/'+p);
 const control={schemaVersion:'activity-projection-recovery-v1',projectId:PROJECT,namespace:ROOT,enabled:true,revision:1,startsAt:now-1000,expiresAt:now+3600000,policyIds:['PEOPLE_POLICY'],readModelCommands:[],browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['SYN_APP']},specializedReaders:{enabled:true,actorUids:['CF_ADMIN'],routes:['/admin/pessoas','/admin/pessoas/detalhe']}};
 let revoked=false;const auth={verifyIdToken:async(t,check)=>{a.equal(check,true);if(revoked)throw Object.assign(Error('revoked'),{code:'auth/id-token-revoked'});return {uid:'CF_ADMIN',sub:'CF_ADMIN',aud:DEMO,iss:'https://securetoken.google.com/'+DEMO,iat:Math.floor(now/1000),auth_time:Math.floor(now/1000),exp:Math.floor(now/1000)+3600}}},appCheck={verifyToken:async()=>({appId:'SYN_APP'})};
 const read=(params={},hooks={})=>readAct({db,auth,appCheck,token:'synthetic-emulator-only',appCheckToken:'emulator',origin:'https://geapa-convergencia-dev.web.app',body:{action:'actRead',request:{route:'/admin/pessoas',params:{query:'',category:'',state:'',role:'',order:'name-asc',page:1,pageSize:25,...params}}},hooks});
 try{const entries=Object.entries({...h.docs,'projectionRecoveryControl/current':control});for(let i=0;i<entries.length;i+=300){const b=db.batch();for(const [p,d]of entries.slice(i,i+300))b.set(ref(p),d);await b.commit()}
 const first=await read();a.equal(first.result.data.total,99);a.equal(first.metrics.writes,0);const concurrent=await Promise.all([read({page:2,snapshot:first.result.data.snapshot}),read({page:2,snapshot:first.result.data.snapshot})]);a.deepEqual(concurrent[0].result.data.items,concurrent[1].result.data.items);
 await a.rejects(read({}, {beforeReturn:async()=>{revoked=true}}),/AUTH_BINDING_INVALID/);revoked=false;
 await a.rejects(read({}, {beforeReturn:async()=>{await ref('memberAuthzPolicies/PEOPLE_POLICY').update({active:false})}}),/ACCESS_REVOKED/);
 await ref('memberAuthzPolicies/PEOPLE_POLICY').update({active:true});const recovered=await read();a.equal(recovered.result.data.total,99);a.equal(recovered.metrics.writes,0);
 console.log(JSON.stringify({emulatorOnly:true,first:first.metrics,concurrent:concurrent.map(r=>r.metrics),recovered:recovered.metrics}));
 }finally{await db.terminate()}
});
