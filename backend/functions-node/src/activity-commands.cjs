'use strict';
// Adapter limits only. Preparation, authorization and effects remain in Core.
const {Store,decode}=require('./firestore.cjs');
const {SHARED,ROOT,PROJECT,EXECUTOR,fail,fence}=require('./environment.cjs');
const {windowLive}=require('./execution-window.cjs');
const {isDeepStrictEqual}=require('node:util');
const CONTROL='activityExecutorControl/current';
const routes={'/atividades/modelo/criar':'CREATE','/admin/atividades/salvar':'EDIT','/admin/atividades/publicar':'PUBLISH','/admin/atividades/ocultar':'HIDE','/admin/atividades/cancelar':'CANCEL','/admin/atividades/reabrir':'REOPEN','/admin/atividades/realizacao':'REALIZE'};
const types=['ACT.ACTIVITY_SAVE','ACT.ACTIVITY_PUBLICATION_SET','ACT.ACTIVITY_REALIZE'];
function control(c,uid){
 const now=Date.now(),ids=x=>Array.isArray(x)&&x.length>0&&x.length<=(require('./environment.cjs').PRODUCTION?128:16)&&new Set(x).size===x.length&&x.every(s=>typeof s==='string'&&/^[A-Za-z0-9_-]{1,128}$/.test(s));
 if(!SHARED||ROOT!==''||process.env.GEAPA_ACTIVITY_COMMANDS_ENABLED!=='true'||c?.schemaVersion!=='activity-executor-control-v1'||c.projectId!==PROJECT||c.enabled!==true||!Number.isSafeInteger(c.revision)||c.revision<1||!ids(c.actorUids)||!c.actorUids.includes(uid)||!ids(c.policyIds)||!windowLive(c,now)||c.browserEnabled!==true||c.appCheck?.mode!=='REQUIRED'||!Array.isArray(c.appCheck.appIds)||!c.appCheck.appIds.length||c.appCheck.appIds.length>4)fail('PILOT_DISABLED');
 if(c.appCheck.appIds.some(x=>typeof x!=='string'||!x||x.length>200))fail('PILOT_DISABLED');return c;
}
class ActivityStore extends Store{
 guard(){fence();const c=control(this.hostRead(CONTROL),this.identity?.uid);if(this.pinned&&JSON.stringify(c)!==this.pinned)fail('ACCESS_REVOKED');this.pinned=JSON.stringify(c);this.control=c;return c;}
 io(raw){if(JSON.parse(raw).kind==='pilotConfig')return JSON.stringify({root:ROOT,shared:true,executor:EXECUTOR,settings:{commands:types.map(t=>t.slice(4)),policyIds:this.control.policyIds,admissionsPerHour:20,jobs:[],activityAgendaDirect:false}});return super.io(raw);}
 validateWrites(s){
  const target=this.activityId,operation=this.operationId;
  const commandCollections=['memberAdmissionTickets','memberAdmissionCounters','memberAdmissionUses','commandProvenance','commandRequests','commandReceipts','commandLedger','commandTombstones','commandLeases','commandReconciliation','memberAudit','commandProjectionInvalidations','commandExecutorOwners','commandExecutorRoutes'];
  for(const w of s.writes){
   const [collection,id]=w.path.split('/'),d=w.data&&decode(w.data);
   if(w.mode==='delete')fail('PILOT_WRITE_DENIED');
   if(collection==='activityRealExecutorScopes'){
    const route=this.hostRead('commandExecutorRoutes/'+this.routeId),old=this.hostRead(w.path);
    if(!this.routeId?.startsWith('ACTIVITY_REAL_')||id!==route?.scopeId||!old||old.operationIds.includes(operation)||
      !isDeepStrictEqual(d,{...old,revision:old.revision+1,operationIds:old.operationIds.concat(operation)}))fail('PILOT_WRITE_DENIED');
    continue;
   }
   if(commandCollections.includes(collection)){
    if(['commandRequests','commandReceipts','commandLedger','commandLeases','commandProjectionInvalidations','commandExecutorOwners'].includes(collection)&&id!==operation)fail('PILOT_WRITE_DENIED');
    if(collection==='commandExecutorRoutes'&&id!==this.routeId)fail('PILOT_WRITE_DENIED');continue;
   }
   if(['canonicalActivities','activityAgendaViews','activityAttendanceCollections'].includes(collection)){
    if(id!==target)fail('PILOT_RESOURCE_DENIED');continue;
   }
   if(collection==='activityHistory'||collection==='activityRealizations'){
    const value=d?.snapshot||d;if((collection==='activityHistory'?value?.subjectRef:value?.activityRef)?.id!==target)fail('PILOT_RESOURCE_DENIED');continue;
   }
   if(w.path==='activityAttendanceRegistry/current')continue;
   if(collection==='portalReaderResourceFences'&&['activities','activity_'+target].includes(id))continue;
   fail('PILOT_WRITE_DENIED');
  }
 }
}
function prepareInput(r){
 if(!r||!routes[r.route]||!/^[a-f0-9]{32}$/.test(r.operationId)||!r.payload)fail('COMMAND_TYPE_INVALID');
 if(routes[r.route]!=='CREATE'&&(!/^[A-Za-z0-9_-]{1,128}$/.test(r.payload.idAtividade)||!Number.isSafeInteger(r.payload.revision)||r.payload.revision<1))fail('COMMAND_PAYLOAD_INVALID');
}
function commandInput(r){
 if(!types.includes(r?.commandType)||!/^[a-f0-9]{32}$/.test(r.operationId)||r.payload?.activityRef?.kind!=='activity'||!/^[A-Za-z0-9_-]{1,128}$/.test(r.payload.activityRef.id))fail('COMMAND_TYPE_INVALID');
}
function select(store,run,r,admit=false){
 commandInput(r);store.activityId=r.payload.activityRef.id;store.operationId=r.operationId;
 const selected=run('(r)=>{const o=nodeRead("commandExecutorOwners/"+r.operationId),routes=f4ActivityExecutorRoutes_(),id=o?.routeId||routes.select(nodeRead,r,nodeConfig);f4Core_().need(id,"ACCESS_REVOKED");return id;}',r);
 store.routeId=selected;
 if(admit)run('(x)=>{const routes=f4ActivityExecutorRoutes_(),v=routes.route(nodeRead,nodeConfig,x.id),now=Date.now();f4Core_().need(v.state==="ACTIVE"&&v.executor===f4CommandExecutorId_()&&v.actorUids.includes(nodeActor.uid)&&v.startsAt<=now&&now<v.expiresAt,"ACCESS_REVOKED");routes.fields(nodeRead,x.r,v);}',{id:selected,r});
}
async function refresh(options,activityId,commandKey){
 const {recoverProjection}=require('./projection-recovery.cjs'),views=[];
 for(const viewId of ['activity_'+activityId,'activities'])views.push(await recoverProjection({...options,commandKey,body:{action:'projectionRecover',request:{viewId}}}));
 return {state:views.every(v=>v.state==='CURRENT')?'CURRENT':'REBUILDING',viewIds:views.map(v=>v.viewId),views,rebuildRequired:views.some(v=>v.state!=='CURRENT')};
}
module.exports={ActivityStore,CONTROL,control,routes,types,prepareInput,commandInput,select,refresh};
