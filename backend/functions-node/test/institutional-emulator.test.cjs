'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),a=require('assert/strict'),{Firestore}=require('@google-cloud/firestore'),{readAct}=require('../src/act-readers.cjs'),{ROOT,PROJECT,DEMO}=require('../src/environment.cjs'),{fixture}=require('./institutional-fixture.cjs');
test('institutional emulator: concurrent reads, revoke, data change fence and recovery',{skip:process.env.GEAPA_INSTITUTIONAL_EMULATOR_TEST!=='1'},async()=>{
 const db=new Firestore({projectId:DEMO}),h=fixture(),now=Date.now(),ref=p=>db.doc(ROOT+'/'+p);let revoked=false;
 const control={schemaVersion:'activity-projection-recovery-v1',projectId:PROJECT,namespace:ROOT,enabled:true,revision:1,startsAt:now-1000,expiresAt:now+3600000,policyIds:['PEOPLE_POLICY'],readModelCommands:[],browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['SYN_APP']},specializedReaders:{enabled:true,actorUids:['CF_ADMIN'],routes:['/admin/gestao-institucional']}};
 const auth={verifyIdToken:async(t,check)=>{a.equal(check,true);if(revoked)throw Object.assign(Error('revoked'),{code:'auth/id-token-revoked'});return {uid:'CF_ADMIN',sub:'CF_ADMIN',aud:DEMO,iss:'https://securetoken.google.com/'+DEMO,iat:Math.floor(now/1000),auth_time:Math.floor(now/1000),exp:Math.floor(now/1000)+3600}}},appCheck={verifyToken:async()=>({appId:'SYN_APP'})};
 const read=(hooks={})=>readAct({db,auth,appCheck,token:'synthetic-emulator-only',appCheckToken:'emulator',origin:'https://geapa-convergencia-dev.web.app',body:{action:'actRead',request:{route:'/admin/gestao-institucional',params:{section:'identity',query:'',kind:'',page:1}}},hooks});
 try{const b=db.batch();for(const [p,d]of Object.entries({...h.docs,'projectionRecoveryControl/current':control}))b.set(ref(p),d);await b.commit();
 const first=await read();a.equal(first.result.data.total,1);a.equal(first.metrics.writes,0);const both=await Promise.all([read(),read()]);a.deepEqual(both[0].result.data.items,both[1].result.data.items);
 await a.rejects(read({beforeReturn:async()=>{revoked=true}}),/AUTH_BINDING_INVALID/);revoked=false;
 await a.rejects(read({beforeReturn:async()=>{await ref('memberAuthzPolicies/PEOPLE_POLICY').update({active:false})}}),/ACCESS_REVOKED/);await ref('memberAuthzPolicies/PEOPLE_POLICY').update({active:true});
 await a.rejects(read({beforeReturn:async()=>{await ref('institutionalDirectory/identity').update({revision:2})}}),/ACCESS_REVOKED/);
 const recovered=await read();a.equal(recovered.result.data.revision,2);a.equal(recovered.metrics.writes,0);
 await a.rejects(readAct({db,auth,appCheck:{verifyToken:async()=>{throw Error('invalid')}},token:'synthetic-emulator-only',appCheckToken:'invalid',origin:'https://geapa-convergencia-dev.web.app',body:{action:'actRead',request:{route:'/admin/gestao-institucional',params:{section:'identity',query:'',kind:'',page:1}}}}),/APP_CHECK/);
 console.log(JSON.stringify({emulatorOnly:true,first:first.metrics,concurrent:both.map(r=>r.metrics),recovered:recovered.metrics}));
 }finally{await db.terminate()}
});
