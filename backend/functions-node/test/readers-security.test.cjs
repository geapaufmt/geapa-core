'use strict';
require('../src/local-network.cjs');
const {test}=require('node:test'),a=require('node:assert/strict');
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {control,contract,ReaderStore}=require('../src/act-readers.cjs'),{ROOT,PROJECT,DEMO}=require('../src/environment.cjs'),{encode}=require('../src/firestore.cjs'),{readerQuery}=require('../src/reader-query.cjs'),{queryPlan}=require('../src/query.cjs'),{handler,ORIGIN}=require('../src/http.cjs');
const q={route:'/act/context',params:{}};
const good=()=>({schemaVersion:'activity-projection-recovery-v1',projectId:PROJECT,namespace:ROOT,enabled:true,revision:1,startsAt:Date.now()-1000,expiresAt:Date.now()+60000,readModelCommands:[],policyIds:['P'],browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['local']},specializedReaders:{enabled:true,actorUids:['A'],routes:[q.route]}});
test('operational read rollout retains environment, actor, switch and conditional route fences',()=>{
 const c={...good(),mode:'DEV_READERS_OPERATIONAL',availabilityMode:'PERMANENT',expiresAt:null,actorUids:['A']};
 a(control(c,'A',q.route));
 for(const patch of [{projectId:'prod'},{enabled:false},{expiresAt:Date.now()+31*86400000},{actorUids:[]},{actorUids:['A','A']},{mode:'ANY'},{actorUids:['B']}])a.throws(()=>control({...c,...patch},'A',q.route),/PILOT_DISABLED/);
 for(const route of ['/admin/apresentacoes/cotas','/admin/apresentacoes/incidencia'])a.throws(()=>control({...c,specializedReaders:{...c.specializedReaders,routes:[route]}},'A',route),/CONFIG_CAPABILITY_DISABLED/);
 const old=process.env.GCLOUD_PROJECT;try{process.env.GCLOUD_PROJECT='prod';a.throws(()=>control(c,'A',q.route),/DESTINATION_INVALID/);}finally{process.env.GCLOUD_PROJECT=old;}
});
test('reader protected control pins environment actor route window and App Check',()=>{
 a(control(good(),'A',q.route));for(const patch of [{projectId:'prod'},{namespace:'wrong'},{expiresAt:0},{specializedReaders:{enabled:false,actorUids:['A'],routes:[q.route]}},{appCheck:{mode:'OBSERVE',appIds:[]}}])a.throws(()=>control({...good(),...patch},'A',q.route),/PILOT_DISABLED/);
 a.throws(()=>control(good(),'B',q.route),/PILOT_DISABLED/);a.throws(()=>control(good(),'A','/other'),/PILOT_DISABLED/);
});
test('reader contract has no caller authority force projection or command fields',()=>{
 a.deepEqual(contract({action:'actRead',request:q}),q);
 for(const body of [{action:'execute',request:q},{action:'actRead',request:{...q,uid:'A'}},{action:'actRead',request:q,force:true},{action:'actRead',request:{route:'/atividades/chamada',params:{}}}])a.throws(()=>contract(body),/COMMAND_PAYLOAD_INVALID/);
});
test('all derived and canonical writes and preparation denied by reader store',()=>{
 const s=new ReaderStore({projectId:DEMO},{uid:'A'},q,good());
 for(const kind of ['write','preview','preparationClock'])a.throws(()=>s.io(JSON.stringify({kind,path:'commandRequests/x'})),/PILOT_WRITE_DENIED/);
 a.throws(()=>s.validateWrites({writes:[{path:'portalReaderViews/A/items/self'}]}),/PILOT_WRITE_DENIED/);
});
const db={projectId:DEMO,collection(){return {where(){return this},orderBy(){return this},limit(){return this},startAfter(){return this}}},doc:p=>p};
const simple={collection:'activityJustifications',filters:[{field:'state',op:'EQUAL',value:encode('PENDING')}],limit:11,order:[{field:'__name__',direction:'ASCENDING'}]};
const decode=require('../src/firestore.cjs').decode;

test('cycle configuration history permits one bounded child collection and isolated cursors',()=>{
 const plan={collection:'presentationCycles/SYN_CYCLE/configRevisions',filters:[],limit:21,order:[{field:'__name__',direction:'ASCENDING'}]};
 a(readerQuery(db,plan,decode));
 for(const patch of [{limit:22},{collection:'presentationCycles/SYN_CYCLE/private'},{collection:'presentationCycles/../configRevisions'},{filters:[{field:'cycleRef',op:'EQUAL',value:encode({kind:'cycle',id:'OTHER'})}]}])a.throws(()=>readerQuery(db,{...plan,...patch},decode),/QUERY_NOT_IN_PILOT/);
 a.throws(()=>readerQuery(db,{...plan,cursor:[{referenceValue:'projects/'+PROJECT+'/databases/(default)/documents/'+(ROOT?ROOT+'/':'')+'presentationCycles/OTHER/configRevisions/x'}]},decode),/DESTINATION_INVALID/);
});

test('board transparency queries reject unbounded, cross-board and cross-project plans',()=>{
 const plan={collection:'memberFunctions',filters:[{field:'idDiretoria',op:'EQUAL',value:encode('DIR-2026')}],limit:20,order:[{field:'__name__',direction:'ASCENDING'}]};
 a(readerQuery(db,plan,decode));a(readerQuery(db,{...plan,collection:'commandExecutorRoutes',filters:[]},decode));
 for(const patch of [{limit:21},{filters:[]},{filters:[{...plan.filters[0],op:'NOT_EQUAL'}]},{filters:[{...plan.filters[0],value:encode('../OTHER')}]},{filters:[{...plan.filters[0],value:encode({id:'DIR-2026'})}]}])a.throws(()=>readerQuery(db,{...plan,...patch},decode),/QUERY_NOT_IN_PILOT/);
 a.throws(()=>readerQuery(db,{...plan,cursor:[{referenceValue:'projects/prod/databases/(default)/documents/memberFunctions/x'}]},decode),/DESTINATION_INVALID/);
});
test('typed reader queries are bounded and do not broaden command query adapter',()=>{
 a(readerQuery(db,simple,decode));a.throws(()=>queryPlan(db,simple,decode),/QUERY_NOT_IN_PILOT/);
 for(const patch of [{limit:21},{limit:0},{collection:'memberPrivateProfiles'},{filters:[]},{filters:[{field:'state',op:'NOT_EQUAL',value:encode('PENDING')}]},{filters:[{field:'state',op:'EQUAL',value:encode('ALL')}]},{order:[{field:'private',direction:'ASCENDING'}]}])a.throws(()=>readerQuery(db,{...simple,...patch},decode),/QUERY_NOT_IN_PILOT/);
 for(const referenceValue of ['projects/prod/databases/(default)/documents/activityJustifications/a','projects/'+PROJECT+'/databases/(default)/documents/'+(ROOT?ROOT+'/':'')+'memberPrivateProfiles/a'])a.throws(()=>readerQuery(db,{...simple,cursor:[{referenceValue}]},decode),/DESTINATION_INVALID/);
});
function res(){return {set(){},status(n){this.code=n;return this},json(v){this.value=v;return this},send(){return this}};}
test('cycle-scoped planning queries keep bounded typed equality and cursor isolation',()=>{
 for(const collection of ['canonicalActivities','presentationPolicyReferences']){
  const plan={...simple,collection,filters:[{field:'cycleRef',op:'EQUAL',value:encode({kind:'cycle',id:'GEAPA_2026'})}]};
  a(readerQuery(db,plan,decode));
  for(const patch of [{limit:collection==='canonicalActivities'?101:21},{filters:[{...plan.filters[0],op:'NOT_EQUAL'}]},{filters:[{...plan.filters[0],value:encode({kind:'person',id:'PES-000001'})}]},{filters:[{...plan.filters[0],value:encode({kind:'cycle',id:'../OTHER'})}]},{filters:[{...plan.filters[0],value:encode({kind:'cycle',id:'GEAPA_2026',authority:true})}]}])a.throws(()=>readerQuery(db,{...plan,...patch},decode),/QUERY_NOT_IN_PILOT/);
  a.throws(()=>readerQuery(db,{...plan,cursor:[{referenceValue:'projects/prod/databases/(default)/documents/'+collection+'/x'}]},decode),/DESTINATION_INVALID/);
 }
});
test('read action HTTP is closed by default and keeps CORS and authentication boundary',async()=>{
 let calls=0;const execute=async x=>{a.equal(x.token,'local');a.equal(x.appCheckToken,'attest');calls++;return {route:q.route,result:{ok:true}}},runtime=async()=>({});
 const req={method:'POST',headers:{origin:ORIGIN,authorization:'Bearer local','x-firebase-appcheck':'attest','content-type':'application/json'},body:{action:'actRead',request:q}};
 let r=res();await handler({runtime,execute,log:()=>{}})(req,r);a.equal(r.code,403);a.equal(calls,0);
 r=res();await handler({runtime,execute,actions:['actRead'],log:()=>{}})(req,r);a.equal(r.value.route,q.route);a.equal(calls,1);
 r=res();await handler({runtime,execute,actions:['actRead'],log:()=>{}})({...req,headers:{...req.headers,origin:'https://evil.invalid'}},r);a.equal(r.code,403);a.equal(calls,1);
});
