'use strict';
require('../src/local-network.cjs');
const {test}=require('node:test'),a=require('node:assert/strict');
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {RecoveryStore,validateControl,contract,validateReadModelCommands}=require('../src/projection-recovery.cjs'),{encode}=require('../src/firestore.cjs'),{ROOT,PROJECT,DEMO}=require('../src/environment.cjs'),{handler,ORIGIN}=require('../src/http.cjs');
const good=()=>({schemaVersion:'activity-projection-recovery-v1',projectId:PROJECT,namespace:ROOT,enabled:true,revision:1,startsAt:Date.now()-1000,expiresAt:Date.now()+60000,readModelCommands:[],policyIds:['EXISTING_POLICY'],browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['local-app']}});
test('operational recovery pins the actor and preserves the read-only write boundary',()=>{
 const c={...good(),mode:'DEV_READERS_OPERATIONAL',availabilityMode:'PERMANENT',expiresAt:null,actorUids:['A']};
 const s=new RecoveryStore({projectId:DEMO},{uid:'A'},'activities');s.hostRead=()=>c;a(s.guard());
 const stranger=new RecoveryStore({projectId:DEMO},{uid:'B'},'activities');stranger.hostRead=()=>c;a.throws(()=>stranger.guard(),/PILOT_DISABLED/);
 s.pinned={...c,revision:c.revision-1};a.throws(()=>s.guard(),/ACCESS_REVOKED/);
 s.buildId='a'.repeat(32);a.throws(()=>s.validateWrites({writes:[{mode:'set',path:'canonicalActivities/X',data:encode({})}]}),/PILOT_WRITE_DENIED/);
});
test('recovery control is DEV only, finite and explicitly enabled',()=>{a(validateControl(good()));for(const patch of [{enabled:false},{projectId:'geapa-prod'},{namespace:'wrong'},{revision:0},{expiresAt:0},{expiresAt:Date.now()+86400001},{policyIds:[]},{policyIds:['../../x']}])a.throws(()=>validateControl({...good(),...patch}),/PILOT_DISABLED/);});
test('read availability hints cannot select arbitrary commands or admissions',()=>{
 for(const readModelCommands of [undefined,['../ADMIN'],['ACTIVITY_SAVE','ACTIVITY_SAVE'],Array(129).fill('ACTIVITY_SAVE')])a.throws(()=>validateControl({...good(),readModelCommands}),/PILOT_DISABLED/);
 const s=new RecoveryStore({projectId:DEMO},{uid:'A'},'activities');s.control=validateControl({...good(),readModelCommands:['ACTIVITY_SAVE','ATTENDANCE_FINALIZE']});
 const c=JSON.parse(s.io(JSON.stringify({kind:'pilotConfig'})));a.deepEqual(c.settings.commands,s.control.readModelCommands);a.equal(c.settings.admissionsPerHour,0);a.deepEqual(c.settings.jobs,[]);a.equal(c.executor,'DERIVED_PROJECTION_ONLY');
 a.throws(()=>contract({action:'execute',request:{viewId:'activities'}}),/COMMAND_PAYLOAD_INVALID/);
 s.buildId='a'.repeat(32);a.throws(()=>s.validateWrites({writes:[{mode:'set',path:'canonicalActivities/X',data:encode({ownerUid:'A'})}]}),/PILOT_WRITE_DENIED/);
});
test('derived adapter cannot write facts, owners, receipts, context or another UID',()=>{
 const s=new RecoveryStore({projectId:DEMO},{uid:'A'},'activities');s.buildId='a'.repeat(32);
 for(const path of ['canonicalActivities/X','commandRequests/X','commandReceipts/X','commandExecutorOwners/X','portalReaderViews/B/items/activities','portalReaderViews/A/items/context'])a.throws(()=>s.validateWrites({writes:[{mode:'set',path,data:encode({ownerUid:'A'})}]}),/PILOT_WRITE_DENIED/);
 a.throws(()=>s.validateWrites({writes:[{mode:'delete',path:s.viewPath}]}),/PILOT_WRITE_DENIED/);
 a.throws(()=>s.validateWrites({writes:[{mode:'set',path:s.viewPath,data:encode({schemaVersion:'portal-reader-view-v1',viewId:'activities',basis:{ownerUid:'B'}})}]}),/PILOT_WRITE_DENIED/);
});
test('recovery accepts exactly one list/detail selector, no client authority or force',()=>{a.equal(contract({action:'projectionStatus',request:{viewId:'activities'}}),'activities');for(const request of [{viewId:'activities',force:true},{viewId:'activities',ownerUid:'A'},{viewId:'activities',capabilities:['ADMIN']},{viewIds:['activities']},{viewId:'activity_'+'a'.repeat(129)}])a.throws(()=>contract({action:'projectionRecover',request}),/COMMAND_PAYLOAD_INVALID/);});
function res(){return {headers:{},set(k,v){this.headers[k]=v;},status(s){this.code=s;return this;},json(v){this.value=v;return this;},send(v){this.value=v;return this;}};}
test('HTTP projection action is closed by default and explicit whitelist reuses Auth/CORS boundary',async()=>{
 let calls=0;const execute=async x=>{a.equal(x.token,'synthetic');a.equal(x.appCheckToken,'attestation');calls++;return {state:'CURRENT'};},runtime=async()=>({});
 const req={method:'POST',headers:{origin:ORIGIN,authorization:'Bearer synthetic','content-type':'application/json','x-firebase-appcheck':'attestation'},body:{action:'projectionRecover',request:{viewId:'activities'}}};
 let r=res();await handler({runtime,execute,log:()=>{}})(req,r);a.equal(r.code,403);a.equal(calls,0);
 r=res();await handler({runtime,execute,actions:['projectionRecover','projectionStatus'],log:()=>{}})(req,r);a.equal(r.value.state,'CURRENT');a.equal(calls,1);
 r=res();await handler({runtime,execute,actions:['projectionRecover'],log:()=>{}})({...req,headers:{...req.headers,origin:'https://evil.invalid'}},r);a.equal(r.code,403);a.equal(calls,1);
});

test('read hint catalog is the canonical Core catalog',()=>{const d=require('../src/domain.cjs').createDomain(()=>{throw Error('UNEXPECTED_IO')});const check=x=>validateReadModelCommands(d,x);a.doesNotThrow(()=>check(['ACTIVITY_PUBLICATION_SET','THEME_VALIDATE','MATERIAL_REVIEW']));a.throws(()=>check(['ADMIN']),/PILOT_DISABLED/);});

test('finite policy registry accommodates scoped extensions without dropping existing sources',()=>{const policies=Array.from({length:32},(_,i)=>'POLICY_'+i);a(validateControl({...good(),policyIds:policies}));a.throws(()=>validateControl({...good(),policyIds:[...policies,'EXTRA']}),/PILOT_DISABLED/);a.throws(()=>validateControl({...good(),policyIds:['P','P']}),/PILOT_DISABLED/);});
