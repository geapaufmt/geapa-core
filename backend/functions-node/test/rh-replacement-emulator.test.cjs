'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GEAPA_PILOT_MODE:'SHARED_DEV',GEAPA_ACTIVITY_COMMANDS_ENABLED:'true',GEAPA_ACT_FLOWS_ENABLED:'true',GEAPA_ACT_EXTENSIONS_ENABLED:'true',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),a=require('node:assert/strict'),crypto=require('node:crypto'),{Firestore}=require('@google-cloud/firestore'),{execute}=require('../src/executor.cjs'),{PROJECT,DEMO,EXECUTOR}=require('../src/environment.cjs'),{setup}=require('./activity-profile-fixture.cjs');
test('replacement through Portal adapter and shared executor preserves facts, planning, receipts and both revisions',{skip:process.env.GEAPA_RH_EMULATOR_TEST!=='1'},async()=>{
 const db=new Firestore({projectId:DEMO}),now=Date.now(),docs=setup('PRESIDENTE').docs,ref=(kind,id)=>({kind,id});
 const resources=['presentationCycles/SYN_REPLACE','activityAttendanceRegistry/current'];
 for(const [id,occurred]of [['SYN_ORIGINAL',false],['SYN_SUBSTITUTE',true]]){
  docs['canonicalActivities/'+id]={id,revision:1,recordMode:'NEW_V2',status:occurred?'OCCURRED':'NOT_OCCURRED',cycleId:'SYN_REPLACE',cycleRef:ref('cycle','SYN_REPLACE'),civilDate:'2026-11-01',realizationRef:ref('activityRealization','R_'+id),typeSnapshot:{countsAttendance:true}};
  docs['activityRealizations/R_'+id]={id:'R_'+id,revision:1,recordMode:'NEW_V2',recognized:true,occurred,occurrenceTemporalObject:{precision:'DAY',civilDate:'2026-11-01'}};
  resources.push('canonicalActivities/'+id,'activityRealizations/R_'+id);
 }
 docs['presentationCycles/SYN_REPLACE']={id:'SYN_REPLACE',revision:1,recordMode:'NEW_V2',status:'OPEN',frequencyPlanning:{periods:[{planned:25,absenceLimit:5}]}};
 docs['activityAttendance/SYN_REPLACE_FACT']={id:'SYN_REPLACE_FACT',revision:1,activityRef:ref('activity','SYN_SUBSTITUTE'),status:'PRESENT'};
 const p=docs['memberAuthzPolicies/CF_DEV_ACT_POLICY'];p.capabilities=['ACTIVITY_REALIZATION_MANAGE'];p.fieldMasks={ACTIVITY_REALIZATION_MANAGE:[]};p.targets.allowedTargetScopes=['ACTIVITIES'];p.targets.allowedResources=resources;delete p.targets.resourceScope;
 docs['activityExecutorControl/current']={schemaVersion:'activity-executor-control-v1',projectId:PROJECT,enabled:true,revision:1,actorUids:['CF_ADMIN'],policyIds:['CF_DEV_ACT_POLICY'],startsAt:now-1000,expiresAt:now+3600000,browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['SYN_APP']}};
 docs['actFlowExecutorScopes/SYN_REPLACE']={id:'SYN_REPLACE',projectId:PROJECT,synthetic:true,revision:1,resourcePaths:resources};
 docs['commandExecutorRoutes/ACT_FLOW_ACTIVITY_REPLACE']={schemaVersion:'geapa-act-flow-route-v1',projectId:PROJECT,routeId:'ACT_FLOW_ACTIVITY_REPLACE',commandType:'ACT.ACTIVITY_REPLACE',scopeId:'SYN_REPLACE',syntheticOnly:true,executor:EXECUTOR,state:'ACTIVE',epoch:1,revision:1,actorUids:['CF_ADMIN'],operationIds:[],maxOperations:8,availabilityMode:'TIMEBOXED',startsAt:now-1000,expiresAt:now+3600000};
 const auth={verifyIdToken:async()=>({uid:'CF_ADMIN',sub:'CF_ADMIN',aud:DEMO,iss:'https://securetoken.google.com/'+DEMO,iat:Math.floor(now/1000),auth_time:Math.floor(now/1000),exp:Math.floor(now/1000)+3600})},appCheck={verifyToken:async()=>({appId:'SYN_APP'})};
 const call=(action,request,hooks={})=>execute({db,auth,appCheck,token:'synthetic-emulator-only',appCheckToken:'synthetic',origin:'https://geapa-convergencia-dev.web.app',body:{action,request},hooks});
 try{
  const b=db.batch();for(const [path,value]of Object.entries(docs))b.set(db.doc(path),value);await b.commit();
  const input={operationId:crypto.randomBytes(16).toString('hex'),route:'/admin/atividades/substituicao',payload:{activityId:'SYN_ORIGINAL',activityRevision:1,replacementId:'SYN_SUBSTITUTE',replacementRevision:1,reason:'Synthetic documented replacement'}};
  const q=(await call('flowPortalPrepare',input)).request;
  await a.rejects(call('execute',q,{afterCommit(){throw Error('LOST_RESPONSE')}}),/LOST_RESPONSE/);
  await call('execute',q);
  for(const id of ['SYN_ORIGINAL','SYN_SUBSTITUTE'])a.equal((await db.doc('canonicalActivities/'+id).get()).data().revision,2);
  a.equal((await db.doc('canonicalActivities/SYN_SUBSTITUTE').get()).data().frequencySlotRef.id,'SYN_ORIGINAL');
  a.deepEqual((await db.doc('presentationCycles/SYN_REPLACE').get()).data(),docs['presentationCycles/SYN_REPLACE']);
  a.deepEqual((await db.doc('activityAttendance/SYN_REPLACE_FACT').get()).data(),docs['activityAttendance/SYN_REPLACE_FACT']);
  a.equal((await db.collection('activityHistory').where('operationId','==',q.operationId).get()).size,2);
  await a.rejects(call('flowPortalPrepare',{...input,operationId:crypto.randomBytes(16).toString('hex')}),/STALE_REVISION/);
 }finally{await db.terminate();}
});
