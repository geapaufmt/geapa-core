'use strict';
// Preserve synchronous domain functions in a worker; actual I/O runs through the
// official asynchronous SDK in the parent. There are no REST calls to Firestore.
const {parentPort,workerData}=require('node:worker_threads'),{PROJECT,ROOT}=require('./sdk.cjs');
const control=new Int32Array(workerData.shared,0,4),bytes=new Uint8Array(workerData.shared,16),input=workerData.input;
function sdk(value){Atomics.store(control,0,0);parentPort.postMessage({type:'sdk',value:{project:PROJECT,environment:'DEV',...value}});if(Atomics.wait(control,0,0,60000)==='timed-out')throw Error('SDK_BRIDGE_TIMEOUT');const result=JSON.parse(Buffer.from(bytes.slice(0,Atomics.load(control,1))).toString());if(!result.ok)throw Object.assign(Error('SDK_OPERATION_FAILED'),{code:result.code});return result.value;}
try{
 const r=require('./local.cjs').runtime({'coreDevValidation/ACT_FUNCTIONAL_20261008':ROOT});
 r.context.__b0Now=()=>input.clock||Date.now();r.context.__b0Verify=token=>sdk({kind:'verify',token});
 const memoryFetch=r.context.UrlFetchApp.fetch;
 r.context.UrlFetchApp.fetch=(url,options)=>{
  try{
  const base='https://firestore.googleapis.com/v1/projects/'+PROJECT+'/databases/(default)/documents',body=options.payload?JSON.parse(options.payload):{};
  if(!url.startsWith(base))throw Error('B0_NETWORK_DENIED');let v;
  const stage=r.context.f4DevReadTiming_.active_?.frame?.label;
  if(options.method==='get'){const u=new URL(url);v=sdk({kind:'get',name:decodeURIComponent(u.pathname.slice('/v1/'.length)),body:{transaction:u.searchParams.get('transaction')||undefined},stage});}
  else{const pos=url.lastIndexOf(':'),kind=url.slice(pos+1),parent=url.slice(base.length,pos);v=sdk({kind,body,...(kind==='runQuery'&&parent&&parent!=='/'+ROOT?{parent:base.slice('https://firestore.googleapis.com/v1/'.length)+parent}:{}),stage});}
  return {getResponseCode:()=>v===null?404:200,getContentText:()=>JSON.stringify(v===null?{error:{status:'NOT_FOUND'}}:v)};
  }catch(e){const status={6:[409,'ALREADY_EXISTS'],9:[400,'FAILED_PRECONDITION'],10:[409,'ABORTED']}[e.code];if(!status)throw e;return {getResponseCode:()=>status[0],getContentText:()=>JSON.stringify({error:{status:status[1]}})};}
 };
 if(input.memoryDocs){for(const [path,d]of Object.entries(input.memoryDocs))if(d)r.documents.set(path,d);r.context.UrlFetchApp.fetch=memoryFetch;require('./test-support/functional/compat-harness.cjs').queries(r);}
 const raw=r.run;r.run=(fn,arg)=>raw('f4DevBoundary_().run(()=>('+fn.toString()+')('+JSON.stringify(arg??null)+'))');
 r.run(x=>{
  Date.now=__b0Now;globalThis.tick=Date.now();globalThis.cfg=f4aTarget_();globalThis.idToken=x.token;
  // Same 120/hour setting as current DEV; confined to the worker's in-memory
  // Script Properties and B0 namespace. No published configuration is changed.
  PropertiesService.getScriptProperties().setProperty('ACT_CORE_DEV_CONFIG',JSON.stringify({project:'geapa-convergencia-dev',root:x.namespace,mode:'SYNTHETIC_FUNCTIONAL',provider:'DEV_FAKE',uids:x.allowedUids,functional:{commands:Object.keys(f4ActivitiesContracts_()),policyIds:x.policyIds||['ACT_POLICY'],admissionsPerHour:120,jobs:[],activityAgendaDirect:false}}));
  const verify=token=>JSON.parse(JSON.stringify(__b0Verify(token)));globalThis.identity=verify(idToken);
  globalThis.pipe=f4CorePipeline_({config:cfg,clock:Date.now,verifyIdentity:verify,membersAuthorization:()=>[],membersRevisions:()=>[]});
  globalThis.ports=f4ActivitiesFunctionalPorts_({config:cfg,identity,clock:Date.now,settings:{commands:Object.keys(f4ActivitiesContracts_()),policyIds:x.policyIds||['ACT_POLICY']}});
  globalThis.compat=f4ActivitiesActivityCompat_({config:cfg,ports,clock:Date.now});
  f4DevReadTiming_.active_={metrics:{},stages:{},counts:{reads:0,writes:0,firestoreCalls:0}};
 },{...input,namespace:ROOT,allowedUids:process.env.B0_ACTOR_UIDS?.split(',').filter(Boolean)||[]});
 const result=r.run(x=>{
  const F=f4Firestore_(),M=f4ActivitiesModel_(),db=F.connect(cfg),started=Date.now();let request=x.request;
  if(x.mode==='projection'){const docs=f4PortalProjectionRebuild_({config:cfg,identity,clock:Date.now,settings:{commands:Object.keys(f4ActivitiesContracts_()),policyIds:x.policyIds||['ACT_POLICY']}},x.viewIds);return {viewPaths:Object.keys(docs),projectionMs:Date.now()-started};}
  if(x.mode==='roster'){const rosterRef=f4ActivitiesAttendanceRoster_({config:cfg,ports,clock:Date.now}).prepare({kind:'activity',id:x.activityId});return {rosterRef,rosterMs:Date.now()-started};}
  if(!request){
   if(x.mode==='edit'){
    const a=db.get('canonicalActivities/'+x.activityId);M.need(a,'NOT_FOUND');
    request=compat.prepare({operationId:x.operationId,route:'/admin/atividades/salvar',payload:{idAtividade:a.id,revision:x.expectedRevision,campos:{tituloPublico:x.title}}});
   }else request=ports.prepare({operationId:x.operationId,commandType:x.commandType,payload:x.payload});
  }
  const preparedAt=Date.now();if(x.prepareOnly)return {request};
  const prior=db.get('commandRequests/'+request.operationId);
  let key,ticket;if(prior){
   // Original immutable intent is required for recovery; changed payload cannot
   // borrow a previous operation ID, even when its receipt is already terminal.
   f4Core_().exact(request,['schemaVersion','operationId','commandType','target','payload','evidenceRefs','decisionRef']);
   M.need(f4SharedCommandCatalog_().digest(cfg.projectId,{...prior,...request,ownerUid:identity.uid})===prior.digest,'OPERATION_CONFLICT');
   key={operationId:prior.operationId,commandType:prior.commandType,digest:prior.digest};
  }else{
   ticket=pipe.endpoint.post({method:'POST',projectId:cfg.projectId,scriptId:f4DevBoundary_().script,body:JSON.stringify({action:'core.admit.v13',idToken,request})});
   M.need(ticket.requestEnvelope,'AUTH_BINDING_INVALID');
   const q={...ticket.requestEnvelope,ticketId:ticket.ticketId,receivedAt:F.Timestamp.fromMillis(Date.now()),state:'RECEIVED'};
   db.transaction(tx=>{const old=tx.get('commandRequests/'+q.operationId).data(),used=tx.get('memberAdmissionUses/'+q.ticketId).data();M.need(!old&&!used,'OPERATION_CONFLICT');tx.set('commandRequests/'+q.operationId,q);tx.set('memberAdmissionUses/'+q.ticketId,{schemaVersion:'member-admission-use-v13',ticketId:q.ticketId,operationId:q.operationId,ownerUid:q.ownerUid,reservedAt:q.receivedAt});});
   key={operationId:q.operationId,commandType:q.commandType,digest:q.digest};
  }
  const admittedAt=Date.now(),done=pipe.endpoint.attempt(idToken,key,'B0_NODE_WORKER'),committedAt=Date.now();
  const metrics={...f4DevReadTiming_.active_};delete metrics.frame;
  return {request,key,result:done,phases:{prepareMs:preparedAt-started,admissionMs:admittedAt-preparedAt,processMs:committedAt-admittedAt,commandMs:committedAt-started},metrics};
 },input);
 if(input.memoryDocs)result.memoryAfter=Object.fromEntries(r.documents);
 parentPort.postMessage({type:'result',value:result});
}catch(e){parentPort.postMessage({type:'error',code:e.code||e.message,detail:e.detail,stack:e.stack?.split('\n').filter(l=>l.trim().startsWith('at ')).slice(0,12).join('\n')});}
