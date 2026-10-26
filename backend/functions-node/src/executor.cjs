'use strict';
const {Store}=require('./firestore.cjs'),{createDomain}=require('./domain.cjs'),{fence,fail,ACTIVITY,EXECUTOR}=require('./environment.cjs');
const pilot=require('./pilot.cjs'),{authenticate:verifyActor,appCheckGate}=require('./auth.cjs');
async function execute({db,auth,token,body,appCheck,appCheckToken,origin,hooks={}}){
 const env=fence();if(!env.emulator&&Object.keys(hooks).length)fail('DESTINATION_INVALID');
 const started=performance.now(),phases={};let actor;
 let priorOwner=null;let expanded=process.env.GEAPA_ACTIVITY_COMMANDS_ENABLED==='true';
 // A prior pilot receipt keeps its original adapter after the wider lane opens.
 // Never transfer an uncertain operation merely because Hosting changed routes.
 if(expanded){
  if(env.root!=='')fail('DESTINATION_INVALID');
  if(body?.request?.payload?.activityRef?.id===ACTIVITY||body?.request?.payload?.idAtividade===ACTIVITY)expanded=false;
  else if(['recover','status'].includes(body?.action)&&/^[a-f0-9]{32}$/.test(body.request?.operationId)){
   const identity=await verifyActor(auth,token);if(db.projectId!==env.projectId)fail('DESTINATION_INVALID');
   const owner=(await db.doc('commandExecutorOwners/'+body.request.operationId).get()).data();priorOwner=owner;
   if(owner&&owner.ownerUid!==identity.uid)fail('ACCESS_REVOKED');
   if(owner&&!owner.routeId&&owner.activityId===ACTIVITY)expanded=false;
  }
 }
 const flowEnabled=process.env.GEAPA_ACT_FLOWS_ENABLED==='true',flowType=body?.request?.commandType;
 let flow=flowEnabled&&(body?.action==='flowRoute'||body?.action==='flowPreview'||body?.action==='flowPrepare'||body?.action==='flowPortalPrepare'||require('./act-flows.cjs').types.includes(flowType)&&!['ACT.ACTIVITY_SAVE','ACT.ACTIVITY_PUBLICATION_SET','ACT.ACTIVITY_REALIZE'].includes(flowType));
 if(flowEnabled&&['ACT.ACTIVITY_SAVE','ACT.ACTIVITY_PUBLICATION_SET','ACT.ACTIVITY_REALIZE'].includes(flowType)){if(priorOwner)flow=!!priorOwner.routeId?.startsWith('ACT_FLOW_');else if(body.request.payload?.activityRef){await verifyActor(auth,token);if(db.projectId!==env.projectId)fail('DESTINATION_INVALID');const id=body.request.payload.activityRef.id;if(!/^[A-Za-z0-9_-]{1,128}$/.test(id))fail('COMMAND_PAYLOAD_INVALID');const a=(await db.doc('canonicalActivities/'+id).get()).data();flow=!!a?.presentationSessionRef;}}
 if(flow)expanded=true;
 const adapter=flow?require('./act-flows.cjs'):expanded?require('./activity-commands.cjs'):pilot;
 const store=flow?new adapter.FlowStore(db):expanded?new adapter.ActivityStore(db):new Store(db),domain=createDomain(x=>store.io(x));
 const run=(fn,x)=>{const v=domain.run(fn,x);return v===undefined?undefined:JSON.parse(v);};
 const phase=async(name,fn)=>{const t=performance.now();try{return await fn();}finally{phases[name]=(phases[name]||0)+performance.now()-t;}};
 const authenticate=()=>phase('authMs',async()=>{actor=await verifyActor(auth,token);store.identity=actor;run('(x)=>nodeIdentity(x)',actor);});
 await authenticate();
 if(!body||Object.keys(body).some(k=>!['action','request'].includes(k))||!body.request)fail('COMMAND_PAYLOAD_INVALID');
 let request,key,result;
 store.action=body.action;
 await store.read(()=>null);
 const appCheckSpec=JSON.stringify(store.control.appCheck);
 const appCheckState=await appCheckGate(appCheck,appCheckToken,store.control,origin);
 const checked=fn=>()=>{const c=store.guard();if(JSON.stringify(c.appCheck)!==appCheckSpec)fail('ACCESS_REVOKED');if(c.appCheck.mode==='REQUIRED'&&appCheckState!=='VERIFIED')fail('APP_CHECK_REQUIRED');return fn();};
 const guarded=fn=>checked(()=>{if(key){const saved=store.hostRead('commandRequests/'+key.operationId);if(saved){adapter.commandInput(saved,expanded?undefined:store.hostRead('canonicalActivities/'+ACTIVITY));if(expanded)adapter.select(store,run,saved);}run('(k)=>nodePilotOperation(k,false)',key);}return fn();});
 if(body.action==='activityRoute'){
  if(!expanded)fail('COMMAND_TYPE_INVALID');
  const routing=await store.read(checked(()=>run('(x)=>nodeActivityRoute(x)',body.request)));
  return {...routing,metrics:{...store.metrics,...phases,totalMs:performance.now()-started}};
 }
 if(body.action==='flowPreview'){if(!flow||body.request?.dryRun!==true)fail('COMMAND_TYPE_INVALID');const preview=await phase('previewMs',()=>store.read(checked(()=>run('(x)=>nodePreview(x)',body.request))));return {preview,metrics:{...store.metrics,...phases,totalMs:performance.now()-started}};}
 if(body.action==='preview'){
  if(!expanded||body.request?.dryRun!==true)fail('COMMAND_TYPE_INVALID');
  const preview=await phase('previewMs',()=>store.read(checked(()=>run('(x)=>nodePreview(x)',body.request))));
  return {preview,metrics:{...store.metrics,...phases,totalMs:performance.now()-started}};
 }
 if(body.action==='flowRoute'){
  const routed=await store.read(checked(()=>run('(x)=>{const t=nodeFlowPortalType(x,true);if(!t)return {executor:"APPS_SCRIPT_DEV"};const id="ACT_FLOW_"+t.slice(4),f=f4ActFlowExecutor_();if(!nodeRead("commandExecutorRoutes/"+id))return {executor:"APPS_SCRIPT_DEV"};const v=f.route(nodeRead,nodeConfig,id),n=Date.now();if(t==="ACT.MEMBER_PRESENTATION_ACTIVITY_CREATE"&&!nodeRead("activityTypes/"+x.payload?.idConfig)?.portalModel?.memberPresentation)return {executor:"APPS_SCRIPT_DEV"};if(["ACT.ACTIVITY_SAVE","ACT.ACTIVITY_PUBLICATION_SET","ACT.ACTIVITY_REALIZE"].includes(t)&&!nodeRead("canonicalActivities/"+x.payload?.idAtividade)?.presentationSessionRef)return {executor:"APPS_SCRIPT_DEV"};f4Core_().need(v.state==="ACTIVE"&&v.actorUids.includes(nodeActor.uid)&&v.startsAt<=n&&n<v.expiresAt,"ACCESS_REVOKED");return {executor:v.executor,routeId:id};}',body.request)));
  if(routed.routeId&&require('./act-flows.cjs').extensions.includes(routed.routeId.slice(9))&&process.env.GEAPA_ACT_EXTENSIONS_ENABLED!=='true')fail('COMMAND_TYPE_INVALID');
  return {...routed,metrics:{...store.metrics,...phases,totalMs:performance.now()-started}};
 }
 if(body.action==='flowPrepare'||body.action==='flowPortalPrepare'){
  if(!flow)fail('COMMAND_TYPE_INVALID');
  await store.read(checked(()=>adapter.preparation(store,run,body.request,body.action==='flowPortalPrepare')));
  store.preparation=true;try{request=await phase('prepareMs',()=>store.transaction(checked(()=>run('(x)=>nodeFlowPrepare(x)',{input:body.request,portal:body.action==='flowPortalPrepare'}))));}finally{store.preparation=false;}
  await store.read(checked(()=>adapter.select(store,run,request,true)));
  return {request,metrics:{...store.metrics,...phases,totalMs:performance.now()-started}};
 }
 if(body.action==='prepare'){
  adapter.prepareInput(body.request);
  request=await phase('prepareMs',()=>store.read(checked(()=>run('(x)=>nodePrepare(x)',body.request))));
  await store.read(checked(()=>{adapter.commandInput(request,expanded?undefined:store.hostRead('canonicalActivities/'+ACTIVITY));if(expanded)adapter.select(store,run,request,body.action==='prepare');}));
  return {request,metrics:{...store.metrics,...phases,totalMs:performance.now()-started}};
 }
 await store.read(checked(()=>run('()=>{globalThis.pipeline=nodePipeline();return null;}')));
 if(body.action==='execute'){
  request=body.request;await store.read(checked(()=>{adapter.commandInput(request,expanded?undefined:store.hostRead('canonicalActivities/'+ACTIVITY));if(expanded)adapter.select(store,run,request,body.action==='prepare');}));if(!expanded&&request.commandType!=='ACT.ACTIVITY_SAVE')fail('COMMAND_TYPE_INVALID');
  // Recovery precedes admission: replay must not require a fresh ticket or effect.
  key=await store.read(checked(()=>run('(r)=>{const C=f4Core_();C.exact(r,["schemaVersion","operationId","commandType","target","payload","evidenceRefs","decisionRef"]);f4ActivitiesPipeline_().validate(r);C.need((f4ActFlowExecutor_().supports(r)&&r.target.expected.length>0)||r.target.expected.some(e=>e.kind==="activity"&&e.id===r.payload.activityRef?.id&&(Number.isSafeInteger(e.revision)&&e.revision>0||e.exists===false)),"COMMAND_TYPE_INVALID");return {operationId:r.operationId,commandType:r.commandType,digest:f4SharedCommandCatalog_().digest(nodeConfig.projectId,{...r,ownerUid:nodeActor.uid,expectedRevisions:[]})};}',request)));
  const exists=await store.read(()=>run('(k)=>!!nodeRead("commandRequests/"+k.operationId)',key));
  if(!exists){store.action='admit';const ticket=await phase('admissionMs',()=>store.transaction(checked(()=>run('(r)=>pipeline.admission.issueBrowser(nodeTx,pipeline.admission.authenticate("verified"),r)',request))));
   key=await phase('reserveMs',()=>store.transaction(checked(()=>run('(t)=>{const F=f4Firestore_(),C=f4Core_(),q={...t.requestEnvelope,ticketId:t.ticketId,receivedAt:F.Timestamp.fromMillis(Date.now()),state:"RECEIVED"},old=nodeTx.get("commandRequests/"+q.operationId).data(),use=nodeTx.get("memberAdmissionUses/"+q.ticketId).data();const key={operationId:q.operationId,commandType:q.commandType,digest:q.digest};nodePilotOperation(key,!old);if(old){C.need(old.digest===q.digest&&old.ownerUid===q.ownerUid,"OPERATION_CONFLICT");}else{C.need(!use,"OPERATION_CONFLICT");nodeTx.set("commandRequests/"+q.operationId,q);nodeTx.set("memberAdmissionUses/"+q.ticketId,{schemaVersion:"member-admission-use-v13",ticketId:q.ticketId,operationId:q.operationId,ownerUid:q.ownerUid,reservedAt:q.receivedAt});}return {operationId:q.operationId,commandType:q.commandType,digest:q.digest};}',ticket))));store.action='recover';
  }
 }else if(['recover','status'].includes(body.action)){key=body.request;if(!(expanded?adapter.types.includes(key.commandType):key.commandType==='ACT.ACTIVITY_SAVE'))fail('COMMAND_TYPE_INVALID');}
 else fail('COMMAND_TYPE_INVALID');
 await authenticate();
 if(!key||!Object.keys(key).every(k=>['operationId','commandType','digest'].includes(k))||!/^([a-f0-9]{32})$/.test(key.operationId)||!/^[a-f0-9]{64}$/.test(key.digest))fail('COMMAND_PAYLOAD_INVALID');
 if(body.action==='status')result=await phase('statusMs',()=>store.transaction(guarded(()=>run('(k)=>pipeline.recovery.readStatus(nodeTx,pipeline.recovery.authenticate("verified"),k)',key))));
 else{
 const claim=await phase('claimMs',()=>store.transaction(guarded(()=>run('(k)=>pipeline.recovery.claim(nodeTx,pipeline.recovery.authenticate("verified"),k,f4CommandExecutorId_())',key))));
 if(claim.state!=='CLAIMED')result=claim;
 else {await hooks.afterClaim?.(key);await authenticate();store.beforeCommit=hooks.beforeCommit;try{result=await phase('finalizeMs',()=>store.transaction(guarded(()=>run('(x)=>pipeline.recovery.finalize(nodeTx,pipeline.recovery.authenticate("verified"),x.key,x.handle)',{key,handle:claim.handle}))));}catch(e){store.beforeCommit=null;try{await store.transaction(guarded(()=>run('(x)=>pipeline.recovery.release(nodeTx,pipeline.recovery.authenticate("verified"),x.key,x.handle)',{key,handle:claim.handle})));}catch{}throw e;}store.beforeCommit=null;await hooks.afterCommit?.(key);}
 }
 let projection={state:'NOT_REFRESHED',rebuildRequired:true};
 if(body.action!=='status'&&['COMPLETED','RECOVERED'].includes(result.state)){
  // The committing actor receives both existing direct-reader projections.
  // The command receipt stays terminal if rebuilding fails; recovery retries
  // only derived work, and Rules keep the earlier list/detail fenced.
  try{await hooks.beforeProjection?.();projection=await phase('projectionMs',()=>expanded?adapter.refresh({db,auth,token,appCheck,appCheckToken,origin,identityUid:actor.uid},store.activityId,key):require('./projection.cjs').rebuildActivityViews({db,auth,token,viewIds:['activity_'+ACTIVITY,'activities']}));}
  catch{projection={state:'STALE',rebuildRequired:true,retryAction:'recover',operationId:key.operationId};}
 }
 return {key,result,projection,metrics:{...store.metrics,...phases,totalMs:performance.now()-started}};
}
module.exports={execute};
