'use strict';
const {isDeepStrictEqual}=require('node:util');
const {SHARED,PROJECT,ROOT,ACTIVITY,EXECUTOR,fail}=require('./environment.cjs');
const CONTROL=SHARED?'pilotControl/portalIntegration':'pilotControl/runtime',OWNER=SHARED?'commandExecutorRoutes/'+ACTIVITY:'pilotExecutorOwnership/current';
function validateControl(c,owner,uid,action,now=Date.now()){
 if(!c||c.schemaVersion!=='geapa-functions-edit-pilot-v1'||c.projectId!==PROJECT||c.namespace!==ROOT||c.activityId!==ACTIVITY||!Number.isSafeInteger(c.revision)||c.revision<1||!['ACTIVE','DRAINING'].includes(c.state)||!Array.isArray(c.actorUids)||!c.actorUids.includes(uid)||c.actorUids.length>4||!Array.isArray(c.policyIds)||!c.policyIds.length||c.policyIds.length>4||c.policyIds.some(x=>!(SHARED?/^[A-Za-z0-9_-]{1,128}$/:/^CF_DEV_[A-Z0-9_]{1,35}$/).test(x))||!Number.isSafeInteger(c.startsAt)||!Number.isSafeInteger(c.expiresAt)||!(c.startsAt<=now&&now<c.expiresAt)||c.expiresAt-c.startsAt>86400000||!Number.isInteger(c.maxOperations)||c.maxOperations<1||c.maxOperations>100)fail('PILOT_DISABLED');
 if(SHARED){if(!owner||owner.schemaVersion!=='geapa-executor-route-v1'||owner.projectId!==PROJECT||owner.activityId!==ACTIVITY||owner.commandType!=='ACT.ACTIVITY_SAVE'||!['ACTIVE','DRAINING'].includes(owner.state)||!owner.actorUids?.includes(uid)||now>=owner.expiresAt||!Number.isSafeInteger(owner.epoch)||owner.epoch<1)fail('EXECUTOR_FENCE_CONFLICT');if(['prepare','admit'].includes(action)&&(owner.executor!==EXECUTOR||owner.state!=='ACTIVE'))fail('EXECUTOR_FENCE_CONFLICT');}
 else if(!owner||owner.schemaVersion!=='geapa-executor-ownership-v1'||owner.executor!==EXECUTOR||owner.namespace!==ROOT||owner.commandType!=='ACT.ACTIVITY_SAVE'||owner.mode!=='ISOLATED'||owner.sharedCutover!==false||!Number.isSafeInteger(owner.epoch)||owner.epoch<1)fail('EXECUTOR_FENCE_CONFLICT');
 if(['prepare','admit'].includes(action)&&c.state!=='ACTIVE')fail('PILOT_DRAINING');return c;
}
function validateResource(marker,activity){
 if(!marker||marker.schemaVersion!=='geapa-pilot-resource-v1'||marker.projectId!==PROJECT||marker.namespace!==ROOT||marker.synthetic!==true||marker.activityId!==ACTIVITY||!activity||activity.id!==ACTIVITY||!Number.isSafeInteger(activity.revision)||activity.revision<1||activity.recordMode==='HISTORICAL_PARTIAL')fail('PILOT_RESOURCE_DENIED');
}
function prepareInput(r){
 if(!r||r.route!=='/admin/atividades/salvar'||r.payload?.idAtividade!==ACTIVITY||!Number.isSafeInteger(r.payload.revision)||r.payload.revision<1||!r.payload.campos||Object.keys(r.payload.campos).length<1||(!SHARED&&Object.keys(r.payload.campos).some(x=>!['tituloPublico','descricaoPublica'].includes(x))))fail('PILOT_RESOURCE_DENIED');
}
function commandInput(r,activity){
 if(r?.commandType!=='ACT.ACTIVITY_SAVE'||r.payload?.activityRef?.id!==ACTIVITY||r.payload.activityRef.kind!=='activity'||!r.target?.expected?.some(e=>e.kind==='activity'&&e.id===ACTIVITY&&Number.isSafeInteger(e.revision)&&e.revision>0))fail('COMMAND_TYPE_INVALID');
 const p=r.payload,allowed=['activityRef','typeRef','title','location','startsAt','endsAt','status','details','decision'];
 if(Object.keys(p).some(k=>!allowed.includes(k))||typeof p.title!=='string'||p.title.length>300||p.decision?.reasonCode!=='PORTAL_ADMIN_EDIT')fail('PILOT_FIELD_DENIED');
 const millis=v=>v?.toMillis?v.toMillis():v;
 for(const k of ['typeRef','location','status'])if(!isDeepStrictEqual(p[k],activity[k]))fail('PILOT_FIELD_DENIED');
 for(const k of ['startsAt','endsAt'])if(p[k]!==millis(activity[k]))fail('PILOT_FIELD_DENIED');
 const rest=d=>Object.fromEntries(Object.entries(d||{}).filter(([k,v])=>k!=='descricaoPublica'&&v!==''));
 if(!isDeepStrictEqual(rest(p.details),rest(activity.details)))fail('PILOT_FIELD_DENIED');
}
function validateWrites(writes,read,decode,uid,projectionViews=[]){
 const allowed=/^(canonicalActivities|activityAgendaViews|activityAttendanceRegistry|activityHistory|memberAdmissionTickets|memberAdmissionCounters|memberAdmissionUses|commandProvenance|commandRequests|commandReceipts|commandLedger|commandTombstones|commandLeases|commandReconciliation|memberAudit|commandProjectionInvalidations|portalReaderResourceFences|pilotOperationOwners|pilotBudget|portalReaderViews|commandExecutorOwners|commandExecutorRoutes)\//;
 for(const w of writes){
  if(!allowed.test(w.path))fail('PILOT_WRITE_DENIED');
  if(w.path.startsWith('activityAgendaViews/')&&w.path!=='activityAgendaViews/'+ACTIVITY)fail('PILOT_RESOURCE_DENIED');
  if(w.path.startsWith('portalReaderViews/')&&w.path!=='portalReaderViews/'+uid+'/items/activity_'+ACTIVITY&&!(projectionViews.includes('activities')&&w.path==='portalReaderViews/'+uid+'/items/activities'))fail('PILOT_RESOURCE_DENIED');
  if(w.path.startsWith('portalReaderResourceFences/')&&!['portalReaderResourceFences/activities','portalReaderResourceFences/activity_'+ACTIVITY].includes(w.path))fail('PILOT_RESOURCE_DENIED');
  if(w.path.startsWith('canonicalActivities/')){
   if(w.path!=='canonicalActivities/'+ACTIVITY&&!new RegExp('^canonicalActivities/'+ACTIVITY+'/revisions/[0-9]+$').test(w.path))fail('PILOT_RESOURCE_DENIED');
   if(w.path==='canonicalActivities/'+ACTIVITY){const before=read(w.path),after=decode(w.data);validateResource(read('pilotResources/'+ACTIVITY),before);
    const omit=x=>Object.fromEntries(Object.entries(x).filter(([k])=>!['title','details','revision','updatedAt','decision','operationId','originRef'].includes(k)));
    if(!isDeepStrictEqual(omit(before),omit(after))||after.revision!==before.revision+1)fail('PILOT_FIELD_DENIED');
    if(!isDeepStrictEqual(after.originRef,{kind:'CORE_COMMAND',operationId:after.operationId}))fail('PILOT_FIELD_DENIED');
    const rest=x=>Object.fromEntries(Object.entries(x||{}).filter(([k,v])=>k!=='descricaoPublica'&&v!==''));
    if(!isDeepStrictEqual(rest(before.details),rest(after.details)))fail('PILOT_FIELD_DENIED');
   }
  }
 }
}
module.exports={CONTROL,OWNER,validateControl,validateResource,prepareInput,commandInput,validateWrites};
