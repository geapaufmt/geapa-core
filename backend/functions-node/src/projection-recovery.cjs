'use strict';
// Derived-only adapter. It cannot admit commands or change canonical facts.
const {createHash,randomUUID}=require('node:crypto');
const {Store,encode,decode}=require('./firestore.cjs');
const {createDomain}=require('./domain.cjs');
const {authenticate,appCheckGate}=require('./auth.cjs');
const {PROJECT,ROOT,fence,fail,PRODUCTION}=require('./environment.cjs');
const CONTROL='projectionRecoveryControl/current',LEASE_MS=90000,WINDOW_MS=600000,MAX_BUILDS=6;
const hash=s=>createHash('sha256').update(s).digest('hex');
// Availability hints for the existing Portal controls, not executable commands.
// Values are copied from the published Core configuration into protected control.
// The HTTP action allowlist and derived-only write guard remain independent.
const MAX_READ_COMMAND_HINTS=128;
// Finite registry accommodates cycle-scoped extensions without replacing live policies.
const MAX_READ_POLICIES=32;
function validateReadModelCommands(domain,commands){
 if(JSON.parse(domain.run('(commands)=>commands.every(c=>Object.prototype.hasOwnProperty.call(f4ActivitiesContracts_(),c))',commands))!==true)fail('PILOT_DISABLED');
}
function validateControl(c){
 const now=Date.now();
 const operational=c?.mode===(PRODUCTION?'PROD_READERS_OPERATIONAL':'DEV_READERS_OPERATIONAL');
 if(PRODUCTION&&(!operational||!Number.isSafeInteger(c.expiresAt)))fail('PILOT_DISABLED');
 // Legacy null remains readable during rollout/rollback; new windows are finite.
 // Institutional policies still expire independently; this is not an auth grant.
 if(operational){fence();if((c.expiresAt!==null&&!require('./execution-window.cjs').windowLive({...c,mode:PRODUCTION?'PROD_SCOPED_OPERATIONAL':'DEV_SCOPED_OPERATIONAL'},now))||!Array.isArray(c.actorUids)||!c.actorUids.length||c.actorUids.length>(PRODUCTION?128:16)||new Set(c.actorUids).size!==c.actorUids.length||c.actorUids.some(x=>typeof x!=='string'||!/^[-A-Za-z0-9_]{1,128}$/.test(x)))fail('PILOT_DISABLED');}
 else if(c?.mode!==undefined||!Number.isSafeInteger(c?.expiresAt)||c.expiresAt<=now||c.expiresAt-c.startsAt>86400000)fail('PILOT_DISABLED');
 if(c?.schemaVersion!=='activity-projection-recovery-v1'||c.projectId!==PROJECT||c.namespace!==ROOT||c.enabled!==true||!Number.isSafeInteger(c.revision)||c.revision<1||!Number.isSafeInteger(c.startsAt)||c.startsAt>now||!Array.isArray(c.policyIds)||!c.policyIds.length||c.policyIds.length>MAX_READ_POLICIES||new Set(c.policyIds).size!==c.policyIds.length||c.policyIds.some(x=>!/^[-A-Za-z0-9_]{1,128}$/.test(x)))fail('PILOT_DISABLED');
 if(c.browserEnabled!==true||c.appCheck?.mode!=='REQUIRED'||!Array.isArray(c.appCheck.appIds)||!c.appCheck.appIds.length||c.appCheck.appIds.length>4||c.appCheck.appIds.some(x=>typeof x!=='string'||!x||x.length>200))fail('PILOT_DISABLED');
 if(!Array.isArray(c.readModelCommands)||c.readModelCommands.length>MAX_READ_COMMAND_HINTS||new Set(c.readModelCommands).size!==c.readModelCommands.length||c.readModelCommands.some(x=>typeof x!=='string'||!/^[_A-Z][_A-Z0-9]{1,63}$/.test(x)))fail('PILOT_DISABLED');
 return c;
}
class RecoveryStore extends Store{
 constructor(db,identity,viewId,control,commandKey){super(db,identity);this.queryConcurrency=4;this.viewId=viewId;this.pinned=control;this.commandKey=commandKey;this.viewPath='portalReaderViews/'+identity.uid+'/items/'+viewId;this.leasePath='projectionRecoveryLeases/'+hash(identity.uid+'|'+viewId);this.budgetPath='projectionRecoveryBudgets/'+hash(identity.uid+(commandKey?'|command|'+commandKey.operationId+'|'+viewId:''));}
 guard(){
  fence();const c=validateControl(this.hostRead(CONTROL));
  if(['DEV_READERS_OPERATIONAL','PROD_READERS_OPERATIONAL'].includes(c.mode)&&!c.actorUids.includes(this.identity.uid))fail('PILOT_DISABLED');
  if(this.pinned&&JSON.stringify(c)!==JSON.stringify(this.pinned))fail('ACCESS_REVOKED');
  this.control=c;this.ownershipEpoch=c.revision;
  if(this.commandKey){
   const k=this.commandKey;if(!/^[a-f0-9]{32}$/.test(k.operationId)||!/^[a-f0-9]{64}$/.test(k.digest))fail('ACCESS_REVOKED');
   const ledger=this.hostRead('commandLedger/'+k.operationId),request=this.hostRead('commandRequests/'+k.operationId);
   const invalidation=this.hostRead('commandProjectionInvalidations/'+k.operationId),binding=this.hostRead('memberAuthBindings/'+this.identity.uid);
   const flowAllowed=process.env.GEAPA_ACT_FLOWS_ENABLED==='true'&&require('./act-flows.cjs').types.includes(k.commandType)&&(invalidation?.fences||[]).includes(this.viewId==='frequency'?'frequency_'+binding?.idPessoa:this.viewId);
   const activityAllowed=['ACT.ACTIVITY_SAVE','ACT.ACTIVITY_PUBLICATION_SET','ACT.ACTIVITY_REALIZE'].includes(k.commandType)&&['activities','activity_'+request?.payload?.activityRef?.id].includes(this.viewId);
   if(ledger?.state!=='COMPLETED'||ledger.ownerUid!==this.identity.uid||ledger.digest!==k.digest||request?.ownerUid!==this.identity.uid||request.digest!==k.digest||request.commandType!==k.commandType||!(flowAllowed||activityAllowed))fail('ACCESS_REVOKED');
  }
  if(this.buildId){const l=this.hostRead(this.leasePath);if(l?.ownerUid!==this.identity.uid||l.viewId!==this.viewId||l.recoveryId!==this.buildId||l.state!=='REBUILDING'||l.expiresAt<=Date.now())fail('OPERATION_CONFLICT');this.leaseDeadline=l.expiresAt;}
  return c;
 }
 io(raw){const x=JSON.parse(raw);if(x.kind==='pilotConfig')return JSON.stringify({root:ROOT,shared:false,executor:'DERIVED_PROJECTION_ONLY',epoch:this.control.revision,settings:{commands:this.control.readModelCommands,policyIds:this.control.policyIds,admissionsPerHour:0,jobs:[],activityAgendaDirect:false}});return super.io(raw);}
 validateWrites(s){
  const allowed=this.buildId?[this.viewPath,this.leasePath]:[this.leasePath,this.budgetPath];
  for(const w of s.writes){
   if(w.mode!=='set'||!allowed.includes(w.path))fail('PILOT_WRITE_DENIED');
   const d=decode(w.data);if(d.ownerUid!==this.identity.uid&&d.basis?.ownerUid!==this.identity.uid)fail('PILOT_WRITE_DENIED');
   if(w.path===this.viewPath&&(d.viewId!==this.viewId||d.schemaVersion!=='portal-reader-view-v1'||Buffer.byteLength(JSON.stringify(w.data))>=700000))fail('PILOT_WRITE_DENIED');
  }
 }
 set(path,value){this.active.writes.push({path,mode:'set',data:encode(value)});}
}
function contract(body){
 if(!body||!['projectionRecover','projectionStatus'].includes(body.action)||Object.keys(body).sort().join()!=='action,request'||!body.request||Object.keys(body.request).join()!=='viewId')fail('COMMAND_PAYLOAD_INVALID');
 const id=body.request.viewId;if(!(id==='frequency'&&process.env.GEAPA_ACT_FLOWS_ENABLED==='true')&&id!=='activities'&&!/^activity_[A-Za-z0-9_-]{1,128}$/.test(id))fail('COMMAND_PAYLOAD_INVALID');return id;
}
async function recoverProjection({db,auth,appCheck,token,appCheckToken,origin,body,hooks={},commandKey}){
 const started=performance.now(),viewId=contract(body),timing={authCalls:0,authMs:0,appCheckMs:0};
 const verify=async()=>{const t=performance.now();timing.authCalls++;try{return await authenticate(auth,token);}finally{timing.authMs+=performance.now()-t;}};
 const identity=await verify();
 if(!fence().emulator&&Object.keys(hooks).length)fail('DESTINATION_INVALID');
 // Control is also read and pinned in every SDK transaction attempt.
 const probe=new RecoveryStore(db,identity,viewId),control=await probe.read(()=>probe.guard());
 const appStart=performance.now();await appCheckGate(appCheck,appCheckToken,control,origin);timing.appCheckMs=performance.now()-appStart;
 // Internal post-commit credit is backed by a terminal canonical command.
 // HTTP cannot supply it. Each operation/view has two attempts per window,
 // separate from the unchanged manual recovery budget and sharing its lease.
 const store=new RecoveryStore(db,identity,viewId,control,commandKey),domain=createDomain(x=>store.io(x));
 const run=(fn,x)=>{const v=domain.run(fn,x);return v===undefined?undefined:JSON.parse(v);};run('(x)=>nodeIdentity(x)',identity);
 // Validate hint names against the unique Core catalog, not a second command list.
 validateReadModelCommands(domain,control.readModelCommands);
 const recoveryId=randomUUID().replaceAll('-','');
 store.beforeCommit=async()=>{await verify();validateControl(control);};
 const decision=await store.transaction(()=>{
  const observed=run('(id)=>nodeProjectionInspect(id)',viewId),now=Date.now();
  const lease=store.hostRead(store.leasePath),budget=store.hostRead(store.budgetPath);
  if(lease&&(lease.schemaVersion!=='projection-recovery-lease-v1'||lease.ownerUid!==identity.uid||lease.viewId!==viewId||!['REBUILDING','CURRENT'].includes(lease.state)||!/^[a-f0-9]{32}$/.test(lease.recoveryId)||!Number.isSafeInteger(lease.expiresAt)))fail('SCHEMA_INVALID');
  if(budget&&(budget.ownerUid!==identity.uid||!Number.isSafeInteger(budget.count)||budget.count<0||!Number.isSafeInteger(budget.windowUntil)))fail('SCHEMA_INVALID');
  if(observed.state==='CURRENT')return {...observed,recoveryId:lease?.recoveryId||null};
  if(lease?.state==='REBUILDING'&&lease.expiresAt>now)return {state:'REBUILDING',recoveryId:lease.recoveryId,retryAfterMs:2000};
  if(body.action==='projectionStatus')return {...observed,recoveryId:lease?.recoveryId||null,retryAfterMs:2000};
  const count=budget?.windowUntil>now?budget.count:0;
  if(count>=(commandKey?2:MAX_BUILDS))fail('PILOT_LIMIT_REACHED');
  store.set(store.budgetPath,{ownerUid:identity.uid,count:count+1,windowUntil:budget?.windowUntil>now?budget.windowUntil:now+WINDOW_MS});
  store.set(store.leasePath,{schemaVersion:'projection-recovery-lease-v1',ownerUid:identity.uid,viewId,recoveryId,state:'REBUILDING',startedAt:now,expiresAt:now+LEASE_MS});
  return {state:'CLAIMED',priorState:observed.state,recoveryId};
 });
 if(decision.state==='CLAIMED'){
  await hooks.afterClaim?.();store.buildId=recoveryId;
  store.beforeCommit=async()=>{await verify();await hooks.beforePublish?.();validateControl(control);if(store.leaseDeadline<=Date.now())fail('OPERATION_CONFLICT');};
  await store.transaction(()=>{
   // Lease and the complete Core source/proof read set share the publish CAS.
   run('(ids)=>nodeProject(ids)',[viewId]);
   store.set(store.leasePath,{schemaVersion:'projection-recovery-lease-v1',ownerUid:identity.uid,viewId,recoveryId,state:'CURRENT',expiresAt:Date.now(),completedAt:Date.now()});
  });
  await hooks.afterPublish?.();decision.state='CURRENT';
 }
 // A lost response is resolved by the same UID+view status selector, not a new
 // command ID. No projection data (especially private data) leaves this route.
 return {...decision,viewId,metrics:{...store.metrics,...timing,controlReads:probe.metrics.reads,controlSdkCalls:probe.metrics.sdkCalls,totalMs:performance.now()-started}};
}
module.exports={recoverProjection,RecoveryStore,validateControl,contract,CONTROL,LEASE_MS,MAX_BUILDS,validateReadModelCommands};
