'use strict';
// The canonical catalog is loaded from the versioned Core, never a browser hint.
const {ActivityStore,control,CONTROL}=require('./activity-commands.cjs');
const {encode,decode}=require('./firestore.cjs');
const {createDomain}=require('./domain.cjs');
const {ROOT,EXECUTOR,fail,PROJECT,PRODUCTION}=require('./environment.cjs');
const catalog=JSON.parse(createDomain(()=>{throw Error('UNEXPECTED_IO');}).run('()=>f4ActFlowExecutor_().commands'));
const extensions=JSON.parse(createDomain(()=>{throw Error('UNEXPECTED_IO');}).run('()=>f4ActRemainingConsumers_().commands'));
const types=catalog.map(c=>'ACT.'+c);
class FlowStore extends ActivityStore{
 guard(){const c=super.guard();if(this.routeId){const r=this.hostRead('commandExecutorRoutes/'+this.routeId);if(r){this.hostRead('actFlowExecutorScopes/'+r.scopeId);if(this.preparation&&(r.state!=='ACTIVE'||r.executor!==EXECUTOR||!r.actorUids?.includes(this.identity.uid)||r.startsAt>Date.now()||r.expiresAt<=Date.now()))fail('ACCESS_REVOKED');}}return c;}
 io(raw){
  const input=JSON.parse(raw);
  if(input.kind==='pilotConfig')return JSON.stringify({root:ROOT,shared:true,executor:EXECUTOR,settings:{commands:catalog,policyIds:this.control.policyIds,admissionsPerHour:20,jobs:[],activityAgendaDirect:false}});
  // A mixed DEV database must not add real people to synthetic attendance.
  // Only this protected executor scope selects the finite fixture population;
  // canonical membership derivation and eligibility remain unchanged.
  if(input.kind==='query'&&input.plan?.collection==='memberAggregateHeads'&&this.routeId){
   const r=this.hostRead('commandExecutorRoutes/'+this.routeId),scope=r&&this.hostRead('actFlowExecutorScopes/'+r.scopeId);
   if(this.active.missing.size)return '[]';
   if(scope?.fixturePersonIds){
    const ids=scope.fixturePersonIds,p=input.plan;
    if(scope.projectId!=='geapa-convergencia-dev'||scope.synthetic!==true||!Array.isArray(ids)||!ids.length||ids.length>30||new Set(ids).size!==ids.length||ids.some(x=>!/^PES-99[0-9]{4}$/.test(x))||p.filters?.length!==0||p.limit!==20||JSON.stringify(p.order)!==JSON.stringify([{field:'__name__',direction:'ASCENDING'}]))fail('PILOT_POPULATION_DENIED');
    const prefix='projects/geapa-convergencia-dev/databases/(default)/documents/memberAggregateHeads/',cursor=p.cursor?.[0]?.referenceValue;
    if(p.cursor&&(!Array.isArray(p.cursor)||p.cursor.length!==1||!cursor?.startsWith(prefix)||!ids.includes(cursor.slice(prefix.length))))fail('PILOT_POPULATION_DENIED');
    const selected=[...ids].sort().filter(id=>!cursor||id>cursor.slice(prefix.length)).slice(0,p.limit),rows=[];
    for(const id of selected){const path='memberAggregateHeads/'+id,head=this.hostRead(path),person=this.hostRead('memberIdentities/'+id);if(this.active.missing.size)continue;
     if(head?.id!==id||head.kind!=='PERSON'||person?.idPessoa!==id||person.origin!=='SYNTHETIC_M1'||!head.paths?.includes('memberIdentities/'+id))fail('PILOT_POPULATION_DENIED');rows.push({path,data:encode(head)});}
    return JSON.stringify(rows);
   }
  }
  return super.io(raw);
 }
 validateWrites(s){
  const route=this.hostRead('commandExecutorRoutes/'+this.routeId),scopePath='actFlowExecutorScopes/'+route?.scopeId,scope=route&&this.hostRead(scopePath);
  if(s.missing.size)return;
  if(scope?.synthetic!==!PRODUCTION||scope.projectId!==PROJECT)fail('PILOT_RESOURCE_DENIED');
  const technical=['memberAdmissionTickets','memberAdmissionCounters','memberAdmissionUses','commandProvenance','commandRequests','commandReceipts','commandLedger','commandTombstones','commandLeases','commandReconciliation','memberAudit','commandProjectionInvalidations','commandExecutorOwners','commandExecutorRoutes'];
  const roots=new Set(scope.resourcePaths),added=[];
  // Optional cycle boundary is protected config. No browser-provided cycle or
  // person can extend it. New canonical outputs resolve through this atomic
  // write set; existing facts must qualify before AND after the mutation.
  const boundary=scope.operational,domain=boundary&&(this.boundaryDomain||(this.boundaryDomain=createDomain(()=>{throw Error('UNEXPECTED_IO');}))); 
  const allowed=(path,writes=false,readOnly=false)=>{
   const overlay=writes?new Map(s.writes.map(w=>[w.path,decode(w.data)])):new Map(),docs={};
   // Domain runs synchronously; collect only concrete dependency reads. Missing
   // reads participate in Store discovery and transaction CAS as usual.
   for(let round=0;round<96;round++){
    const out=JSON.parse(domain.run(x=>{const missing=[];const get=p=>{if(!(p in x.docs)){missing.push(p);return null;}return x.docs[p];};const ok=f4ActCycleBoundary_().contains(get,x.path,x.boundary,x.readOnly);return {ok,missing};},{docs,path,boundary,readOnly}));
    if(!out.missing.length)return out.ok;
    for(const p of out.missing)docs[p]=overlay.has(p)?overlay.get(p):this.hostRead(p)||null;
   }return false;
  };
  for(const w of s.writes){
   const [col,id]=w.path.split('/'),root=col+'/'+id,d=decode(w.data),v=d.snapshot||d;
   if(w.mode==='delete')fail('PILOT_WRITE_DENIED');
   if(this.preparation&&!['activityAttendanceRosters','presentationCapacityAssessments','presentationPolicyReferences'].includes(col))fail('PREPARATION_WRITE_DENIED');
   if(scope.fixturePersonIds&&col==='activityAttendance'&&!scope.fixturePersonIds.includes(v.personRef?.id)){
    // An explicitly scoped, existing file-test fact may receive a SELF request.
    // This never expands the server-resolved roster or changes attendance.
    const list=scope.fixtureAttachmentAttendancePaths,old=this.hostRead(root),activity=old?.activityRef&&this.hostRead('canonicalActivities/'+old.activityRef.id);if(s.missing.size)return;
    if(this.routeId!=='ACT_FLOW_ATTENDANCE_JUSTIFICATION_REQUEST'||!Array.isArray(list)||list.length>20||!list.includes(root)||!(roots.has(root)||scope.operational&&allowed(root))||!old||activity?.originRef?.kind!=='DEV_FIXTURE'||activity.recordMode==='HISTORICAL_PARTIAL'||!(roots.has('canonicalActivities/'+old.activityRef.id)||scope.operational&&allowed('canonicalActivities/'+old.activityRef.id)))fail('PILOT_POPULATION_DENIED');
    const attachmentFields=['pendingJustificationRef','lastJustificationRef','revision','updatedAt','operationId','originRef'];
    if(v.revision!==old.revision+1||v.operationId!==this.operationId||v.originRef?.kind!=='CORE_COMMAND'||v.originRef.operationId!==this.operationId||v.pendingJustificationRef?.kind!=='justification'||v.lastJustificationRef?.id!==v.pendingJustificationRef.id)fail('PILOT_POPULATION_DENIED');
    for(const k of new Set([...Object.keys(old),...Object.keys(v)]))if(!attachmentFields.includes(k)&&JSON.stringify(encode(old[k]??null))!==JSON.stringify(encode(v[k]??null)))fail('PILOT_POPULATION_DENIED');
   }
   if(technical.includes(col)){
    if(['commandRequests','commandProvenance','commandReceipts','commandLedger','commandLeases','commandProjectionInvalidations','commandExecutorOwners'].includes(col)&&id!==this.operationId)fail('PILOT_WRITE_DENIED');
    if(col==='commandExecutorRoutes'&&id!==this.routeId)fail('PILOT_WRITE_DENIED');continue;
   }
   if(col==='portalReaderResourceFences')continue;
   if(w.path==='portalReaderControl/current')continue; // monotonic fallback installed by Store
   // Creating a scoped cycle advances the shared registry, not an institutional
   // cycle. Preserve every existing field except the canonical commit envelope
   // and the single monotonic cycle epoch. Never add the registry to a scope.
   if(w.path==='presentationCapacityHeads/registry'){
    const old=this.hostRead(w.path);if(s.missing.size)return;
    const envelope=['revision','updatedAt','operationId','originRef','schemaVersion','recordMode'];
    if(!['ACT_FLOW_CYCLE_CREATE','ACT_FLOW_CYCLE_OPEN'].includes(this.routeId)||!old||d.id!=='registry'||d.kind!=='DOMAIN_FENCE'||d.cycleRegistryEpoch!==old.cycleRegistryEpoch+1||d.membersEpoch!==old.membersEpoch||d.revision!==old.revision+1||d.operationId!==this.operationId||d.originRef?.kind!=='CORE_COMMAND'||d.originRef.operationId!==this.operationId||d.schemaVersion!=='act-v2-candidate'||d.recordMode!=='NEW_V2')fail('PILOT_WRITE_DENIED');
    for(const k of new Set([...Object.keys(old),...Object.keys(d)]))if(k!=='cycleRegistryEpoch'&&!envelope.includes(k)&&JSON.stringify(encode(old[k]??null))!==JSON.stringify(encode(d[k]??null)))fail('PILOT_WRITE_DENIED');
    continue;
   }
   // Shared invalidation marker is not a fixture entity. The canonical handler
   // may only advance it once; it cannot adopt or rewrite institutional facts.
   if(w.path==='activityAttendanceRegistry/current'){
    const old=this.hostRead(w.path);if(s.missing.size)return;
    const keys=['epoch','id','schemaVersion','recordMode','revision','createdAt','updatedAt','operationId','originRef'];
    if(Object.keys(d).some(k=>!keys.includes(k))||d.id!=='current'||d.schemaVersion!=='act-v2-candidate'||d.recordMode!=='NEW_V2'||d.epoch!==(old?.epoch||0)+1||d.revision!==(old?.revision||0)+1||d.operationId!==this.operationId||d.originRef?.kind!=='CORE_COMMAND'||d.originRef.operationId!==this.operationId||old&&JSON.stringify(encode(d.createdAt))!==JSON.stringify(encode(old.createdAt)))fail('PILOT_WRITE_DENIED');
    continue;
   }
   if(!this.validPath(w.path)||/^(member|migration)/.test(col))fail('PILOT_WRITE_DENIED');
   const extension=extensions.includes(this.routeId?.slice(9)),proofAssociation=col==='jobOutbox'&&this.routeId==='ACT_FLOW_ATTENDANCE_JUSTIFICATION_REQUEST';
   if(['jobOutbox','presentationPolicyReferences','activityTypes','activityPolicies'].includes(col)&&(!(extension||proofAssociation)||process.env.GEAPA_ACT_EXTENSIONS_ENABLED!=='true'&&!this.hostRead('commandExecutorOwners/'+this.operationId)))fail('PILOT_WRITE_DENIED');
   if(col==='jobOutbox'){
    const prior=this.hostRead(root);if(s.missing.size)return;
    if(!['DRIVE_CREATE','FILE_CREATE'].includes(d.effectKind))fail('PILOT_WRITE_DENIED');
    if(!prior){if(proofAssociation||d.status!=='PENDING'||d.createdByOperation!==this.operationId)fail('PILOT_WRITE_DENIED');}
    else{const command=this.routeId.slice(9),mutableFields=['submissionRef','photoSubmitted','revision','createdAt','updatedAt','operationId','originRef','schemaVersion','recordMode'];
     if(proofAssociation)mutableFields.push('justificationRef');
     if(!(boundary?allowed(root):roots.has(root))||!['MATERIAL_SUBMIT','MATERIAL_CORRECTION_SUBMIT','PHOTO_SUBMIT','ATTENDANCE_JUSTIFICATION_REQUEST'].includes(command)||prior.status!=='CONFIRMED'||d.status!=='CONFIRMED')fail('PILOT_WRITE_DENIED');
     if(proofAssociation){if(prior.justificationRef||d.justificationRef?.kind!=='justification'||prior.details?.purpose!=='ATTENDANCE_PROOF'||d.submissionRef!==prior.submissionRef||d.photoSubmitted!==prior.photoSubmitted)fail('PILOT_WRITE_DENIED');}
     else if(command==='PHOTO_SUBMIT'?(prior.photoSubmitted||d.photoSubmitted!==true||prior.details?.purpose!=='PRESENTATION_PHOTO'||d.submissionRef!==prior.submissionRef):(prior.submissionRef||!d.submissionRef||d.submissionRef.kind!=='submission'||d.photoSubmitted!==prior.photoSubmitted))fail('PILOT_WRITE_DENIED');
     if(d.schemaVersion!=='act-v2-candidate'||d.recordMode!=='NEW_V2'||d.revision!==prior.revision+1||d.operationId!==this.operationId||prior.createdAt&&JSON.stringify(encode(prior.createdAt))!==JSON.stringify(encode(d.createdAt)))fail('PILOT_WRITE_DENIED');
     for(const k of new Set([...Object.keys(prior),...Object.keys(d)]))if(!mutableFields.includes(k)&&JSON.stringify(encode(prior[k]??null))!==JSON.stringify(encode(d[k]??null)))fail('PILOT_WRITE_DENIED');
    }
   }
   if(this.preparation&&col==='presentationPolicyReferences'){
    const prior=this.hostRead(root);if(s.missing.size)return;
    if(prior&&JSON.stringify(encode(prior))!==JSON.stringify(encode(d))||d.state!=='SEALED'||d.complete!==true||!['UPLOAD_ARTIFACT_MANIFEST','PORTAL_MATERIAL_COVERAGE','PORTAL_MATERIAL_RECEIPTS','PORTAL_MATERIAL_EXCEPTION_BASIS','PORTAL_DYNAMIC_MATERIALS'].includes(d.kind)||!d.presentationRef||!(boundary?allowed('canonicalPresentations/'+d.presentationRef.id):roots.has('canonicalPresentations/'+d.presentationRef.id)))fail('PILOT_WRITE_DENIED');
    if(boundary)continue;
    roots.add(root);added.push(root);continue;
   }
   if(col==='publicPresentations'){
    const allowed=['publicId','schemaVersion','published','sanitized','cycleId','sourceRevision','updatedAt','title','civilDate','approvedFormat','authorizedAxis','authorizedAxisLabel','permittedDisplay','sourceSnapshotRef','revision'];
    if(!['PUBLICATION_APPROVE','PUBLICATION_WITHDRAW'].includes(this.routeId.slice(9))||!roots.has('canonicalPresentations/'+id)||d.publicId!==id||d.sanitized!==true||Object.keys(d).some(k=>!allowed.includes(k)))fail('PILOT_WRITE_DENIED');
    roots.add(root);if(!scope.resourcePaths.includes(root))added.push(root);continue;
   }
   if(boundary&&(roots.has(root)||roots.has(w.path)))continue; // Preserve approved exact roots and recovery compatibility.
   if(boundary){
    const prior=this.hostRead(root);if(s.missing.size)return;
    const readBefore=prior?allowed(w.path):true,readAfter=allowed(w.path,true);
    if(s.missing.size)return;
    if(!readBefore||!readAfter||v.recordMode==='HISTORICAL_PARTIAL'||!prior&&col!=='activityAgendaViews'&&(v.operationId!==this.operationId||v.originRef?.kind!=='CORE_COMMAND'||v.originRef.operationId!==this.operationId)&&!(this.preparation&&['activityAttendanceRosters','presentationCapacityAssessments'].includes(col)&&d.state==='SEALED'&&d.complete===true))fail('PILOT_RESOURCE_DENIED');
    continue;
   }
   if(roots.has(root)||roots.has(w.path))continue;
   if(col==='activityAgendaViews'&&roots.has('canonicalActivities/'+id)){roots.add(root);added.push(root);continue;}
   // New outputs originate only in this canonical operation. Existing unrelated
   // documents can never be adopted into a scope through a command.
   const old=this.hostRead(root);if(s.missing.size)return;
   if(this.preparation&&!old&&['activityAttendanceRosters','presentationCapacityAssessments'].includes(col)&&d.state==='SEALED'&&d.complete===true&&[d.activityRef,d.cycleRef].some(r=>r&&roots.has(({activity:'canonicalActivities',cycle:'presentationCycles'})[r.kind]+'/'+r.id))){roots.add(root);added.push(root);continue;}
   if(old||v.operationId!==this.operationId||v.originRef?.kind!=='CORE_COMMAND'||v.originRef.operationId!==this.operationId)fail('PILOT_RESOURCE_DENIED');
   roots.add(root);added.push(root);
  }
  if(added.length){if(roots.size>480)fail('SCOPE_TOO_LARGE');s.writes.push({path:scopePath,mode:'set',data:encode({...scope,revision:scope.revision+1,resourcePaths:[...roots].sort()})});}
 }
}
function commandInput(r){if(!types.includes(r?.commandType)||!/^[a-f0-9]{32}$/.test(r.operationId))fail('COMMAND_TYPE_INVALID');}
function select(store,run,r,admit=false){
 commandInput(r);store.operationId=r.operationId;
 if(extensions.includes(r.commandType.slice(4))&&process.env.GEAPA_ACT_EXTENSIONS_ENABLED!=='true'&&!store.hostRead('commandExecutorOwners/'+r.operationId))fail('COMMAND_TYPE_INVALID');
 store.routeId=run('(r)=>{const f=f4ActFlowExecutor_(),o=nodeRead("commandExecutorOwners/"+r.operationId),id=o?.routeId||f.select(nodeRead,r,nodeConfig);f4Core_().need(id,"ACCESS_REVOKED");return id;}',r);
 store.validPath=p=>run('(p)=>f4ActivitiesModel_().validPath(p)',p);
 if(admit)run('(x)=>{const f=f4ActFlowExecutor_(),v=f.route(nodeRead,nodeConfig,x.id),n=Date.now();f4Core_().need(v.state==="ACTIVE"&&v.executor===f4CommandExecutorId_()&&v.actorUids.includes(nodeActor.uid)&&v.startsAt<=n&&n<v.expiresAt,"ACCESS_REVOKED");f.fields(nodeRead,x.r,v);}',{id:store.routeId,r});
}
function preparation(store,run,input,portal){
 const type=portal?run('(x)=>nodeFlowPortalType(x)',input):input.commandType;
 commandInput({commandType:type,operationId:input.operationId});if(extensions.includes(type.slice(4))&&process.env.GEAPA_ACT_EXTENSIONS_ENABLED!=='true')fail('COMMAND_TYPE_INVALID');store.operationId=input.operationId;store.routeId='ACT_FLOW_'+type.slice(4);store.validPath=p=>run('(p)=>f4ActivitiesModel_().validPath(p)',p);
 run('(id)=>{const f=f4ActFlowExecutor_(),v=f.route(nodeRead,nodeConfig,id),n=Date.now();f4Core_().need(v.state==="ACTIVE"&&v.executor===f4CommandExecutorId_()&&v.actorUids.includes(nodeActor.uid)&&v.startsAt<=n&&n<v.expiresAt,"ACCESS_REVOKED");}',store.routeId);
}
// Frequency/presentation families are deliberately not reported CURRENT until
// the canonical producer has published them. The receipt remains recoverable.
async function refresh(options,unused,key){
 const {recoverProjection}=require('./projection-recovery.cjs');
 const snapshots=await options.db.getAll(options.db.doc('commandProjectionInvalidations/'+key.operationId),options.db.doc('memberAuthBindings/'+options.identityUid));
 const invalidation=snapshots[0].data(),binding=snapshots[1].data();
 const ids=(invalidation?.fences||[]).filter(id=>!id.startsWith('frequency_')||id==='frequency_'+binding?.idPessoa).map(id=>id.startsWith('frequency_')?'frequency':id);
 const supported=ids.filter(id=>id==='frequency'||id==='activities'||/^activity_[A-Za-z0-9_-]+$/.test(id));
 // Independent UID+view leases and publication CAS allow two derived views to
 // rebuild concurrently. Bound fan-out; settle failures without losing the
 // successful view or changing the terminal business receipt.
 const views=new Array(supported.length);let cursor=0;
 await Promise.all(Array.from({length:Math.min(2,supported.length)},async()=>{while(cursor<supported.length){const i=cursor++,viewId=supported[i];try{views[i]=await recoverProjection({...options,commandKey:key,body:{action:'projectionRecover',request:{viewId}}});}catch{views[i]={viewId,state:'STALE',rebuildRequired:true,retryAction:'recover'};}}}));
 return {state:views.length&&supported.length===ids.length&&views.every(v=>v.state==='CURRENT')?'CURRENT':'STALE',views,viewIds:ids,deferredViewCount:(invalidation?.fences||[]).length-ids.length,lookupMetrics:{reads:2,sdkCalls:1},rebuildRequired:!views.length||supported.length!==ids.length||views.some(v=>v.state!=='CURRENT'),retryAction:'recover'};
}
module.exports={FlowStore,types,catalog,extensions,commandInput,select,preparation,refresh,control,CONTROL};
