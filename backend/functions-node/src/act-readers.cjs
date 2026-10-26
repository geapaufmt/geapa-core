'use strict';
// Execute canonical Portal producers without business admissions or mutations.
// Bootstrap may repair only the authenticated caller's fenced SELF projections.
const {Store}=require('./firestore.cjs');
const {createDomain}=require('./domain.cjs');
const {authenticate,appCheckGate}=require('./auth.cjs');
const {CONTROL,validateControl}=require('./projection-recovery.cjs');
const {ROOT,fence,fail}=require('./environment.cjs');
const {routes:bootstrapRoutes}=require('./member-readers.cjs');
function control(c,uid,route){
 validateControl(c);
 if(['DEV_READERS_OPERATIONAL','PROD_READERS_OPERATIONAL'].includes(c.mode)&&!c.actorUids.includes(uid))fail('PILOT_DISABLED');
 if(['DEV_READERS_OPERATIONAL','PROD_READERS_OPERATIONAL'].includes(c.mode)&&['/admin/apresentacoes/cotas','/admin/apresentacoes/incidencia'].includes(route))fail('CONFIG_CAPABILITY_DISABLED');
 const r=c.specializedReaders;
 if(r?.enabled!==true||!Array.isArray(r.actorUids)||!r.actorUids.length||r.actorUids.length>(require('./environment.cjs').PRODUCTION?128:16)||!r.actorUids.includes(uid)||
 !Array.isArray(r.routes)||!r.routes.length||r.routes.length>48||!r.routes.includes(route))fail('PILOT_DISABLED');
 return c;
}
class ReaderStore extends Store{
 constructor(db,identity,request,pinned){super(db,identity);this.request=request;this.pinned=pinned;this.specializedReaders=true;this.queryConcurrency=4;}
 guard(){
  fence();const c=control(this.hostRead(CONTROL),this.identity.uid,this.request.route);
  if(this.pinned&&JSON.stringify(c)!==JSON.stringify(this.pinned))fail('ACCESS_REVOKED');
  this.control=c;return c;
 }
 io(raw){
  const x=JSON.parse(raw);
  if(x.kind==='auditClock'){if(!this.active?.native)fail('TRANSACTION_REQUIRED');return JSON.stringify(this.active.preparationAt);}
  if(['preview','preparationClock'].includes(x.kind)||x.kind==='write'&&!bootstrapRoutes.includes(this.request.route))fail('PILOT_WRITE_DENIED');
  if(x.kind==='pilotConfig')return JSON.stringify({root:ROOT,shared:false,executor:'CANONICAL_READ_ONLY',epoch:this.control.revision,settings:{commands:this.control.readModelCommands,policyIds:this.control.policyIds,admissionsPerHour:0,jobs:[],activityAgendaDirect:false}});
  // Resolve the directory's prefetched states before deciding which aggregates
  // need derivation. Unknown discovery placeholders are not missing projections.
  if(x.kind==='getAll'&&this.request.route==='/admin/membros'){
   const result=super.io(raw);
   if(x.paths.some(p=>!this.active.cache.has(p)))fail('DISCOVERY_READ_PENDING');
   return result;
  }
  return super.io(raw);
 }
 validateWrites(s){
  if(!bootstrapRoutes.includes(this.request.route)){if(s.writes.length)fail('PILOT_WRITE_DENIED');return;}
  require('./member-readers.cjs').validateWrites(s,this.request,this.identity.uid);
 }
 async transaction(fn){
  this.metrics.transactions++;const started=performance.now();
  try{return await this.db.runTransaction(async native=>{this.metrics.transactionAttempts++;return this.evaluate(fn,native);},bootstrapRoutes.includes(this.request.route)?{}:{readOnly:true});}
  finally{this.metrics.transactionWallMs=(this.metrics.transactionWallMs||0)+performance.now()-started;}
 }
}
function contract(body){
 if(!body||Object.keys(body).sort().join()!=='action,request'||body.action!=='actRead'||!body.request||Array.isArray(body.request))fail('COMMAND_PAYLOAD_INVALID');
 const q=body.request,keys=q.route==='/atividades/chamada'?['activityId','route']:['params','route'];
 if(Object.keys(q).sort().join()!==keys.join()||typeof q.route!=='string'||q.route.length>128)fail('COMMAND_PAYLOAD_INVALID');
 return q;
}
async function readAct({db,auth,appCheck,token,appCheckToken,origin,body,hooks={},workspacePhase=null}){
 const started=performance.now(),request=contract(body),timing={authMs:0,authCalls:0};
 if(!fence().emulator&&Object.keys(hooks).length)fail('DESTINATION_INVALID');
 const verify=async()=>{const t=performance.now();timing.authCalls++;try{return await authenticate(auth,token);}finally{timing.authMs+=performance.now()-t;}};
 const identity=await verify(),probe=new ReaderStore(db,identity,request),pinned=await probe.read(()=>probe.guard());
 const at=performance.now();await appCheckGate(appCheck,appCheckToken,pinned,origin);timing.appCheckMs=performance.now()-at;
 const store=new ReaderStore(db,identity,request,pinned),domain=createDomain(x=>store.io(x));
 if(hooks.queryConcurrency===1)store.queryConcurrency=1;
 hooks.configureDomain?.(domain);
 const run=(fn,x)=>{const v=domain.run(fn,x);return v===undefined?undefined:JSON.parse(v);};
 run('(x)=>nodeIdentity(x)',identity);
 if(!workspacePhase&&!bootstrapRoutes.includes(request.route)&&!run('(r)=>f4ActivitiesPortalReaderRoutes_().includes(r)',request.route))fail('COMMAND_TYPE_INVALID');
 if(!run('(commands)=>commands.every(c=>Object.prototype.hasOwnProperty.call(f4ActivitiesContracts_(),c))',pinned.readModelCommands))fail('PILOT_DISABLED');
 // All domain reads, policies and epochs use one consistent Firestore snapshot.
 // Token revocation is checked again before returning any private response.
 store.beforeCommit=async()=>{
  await hooks.beforeReturn?.();await verify();control(pinned,identity.uid,request.route);
  // A read-only SDK transaction need not issue a commit. Explicitly recheck the
  // observed authorization/epoch versions after producing the DTO.
  const paths=[...store.active.touched].filter(p=>/^(member|coreDevAdmins\/|projectionRecoveryControl\/|actStorageControl\/|actHistoricalFile|presentationArtifacts\/|canonicalPresentations\/|canonicalActivities\/|activityExecutorControl\/|portalReaderControl\/|portalReaderResourceFences\/|activityAttendanceRegistry\/|presentationCapacityHeads\/registry$)/.test(p));
  for(let i=0;i<paths.length;i+=180){
   const part=paths.slice(i,i+180),docs=await store.measure(()=>db.getAll(...part.map(p=>store.ref(p))));store.metrics.reads+=part.length;
   docs.forEach((d,j)=>{const old=store.active.cache.get(part[j]);if(d.exists!==old.exists||d.exists&&require('./firestore.cjs').encode(d.updateTime).timestampValue!==old.stamp)fail('ACCESS_REVOKED');});
  }
 };
 const result=await store.transaction(()=>{
  let storage;
  if(workspacePhase?.storage&&!/^APR-/.test(request.params?.idApresentacao||''))storage=require('./storage-control.cjs').check(p=>store.hostRead(p),identity.uid,store.hostRead('activityExecutorControl/current')?.expiresAt);
  const result=workspacePhase?run('(x)=>nodeWorkspaceRead(x)',{...workspacePhase,request}):run('(q)=>nodeActRead(q)',request);
  if(storage&&result.plan.settings.provider==='GOOGLE_CLOUD_STORAGE')require('./storage-control.cjs').effect(p=>store.hostRead(p),storage,result.plan.request.effectId,identity.uid,false);
  return result;
 });
 if(Buffer.byteLength(JSON.stringify(result))>(workspacePhase?36*1024*1024:700000))fail('SCOPE_TOO_LARGE');
 return {route:request.route,result,metrics:{...store.metrics,...timing,controlReads:probe.metrics.reads,controlSdkCalls:probe.metrics.sdkCalls,totalMs:Math.round(performance.now()-started)}};
}
module.exports={readAct,ReaderStore,contract,control,bootstrapRoutes};
