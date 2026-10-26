// Execution availability is distinct from institutional policy validity.
// Operational DEV scopes remain finite; this never grants a capability.
function f4ExecutorWindow_(){
 const valid=v=>!!v&&Number.isSafeInteger(v.startsAt)&&Number.isSafeInteger(v.expiresAt)&&v.expiresAt>v.startsAt&&
  (v.mode===undefined||v.mode==='DEV_SCOPED_OPERATIONAL'||v.mode==='PROD_SCOPED_OPERATIONAL')&&v.expiresAt-v.startsAt<=(v.mode?30*86400000:86400000)&&
  (v.mode!=='DEV_SCOPED_OPERATIONAL'||v.projectId==='geapa-convergencia-dev')&&
  (v.mode!=='PROD_SCOPED_OPERATIONAL'||v.projectId==='portal-geapa'&&f4ExecutionEnvironment_().production);
 return {valid,live:(v,now)=>valid(v)&&v.startsAt<=now&&now<v.expiresAt};
}
// Shared DEV pilot fence. The executor identity comes from its adapter, never HTTP.
function f4CommandExecutorId_(){return 'APPS_SCRIPT_DEV';}
function f4ExecutorOwnership_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),ID='CF_DEV_EDIT_ACTIVITY_V2',TYPE='ACT.ACTIVITY_SAVE',ROUTE='commandExecutorRoutes/'+ID;
 const executor=()=>f4CommandExecutorId_();
 function selected(r){return r?.commandType===TYPE&&r.payload?.activityRef?.id===ID;}
 function route(get,config,routeId){
  if(routeId?.startsWith('MEMBERS_'))return f4MemberExecutor_().route(get,config,routeId);
  if(routeId?.startsWith('ACT_FLOW_'))return f4ActFlowExecutor_().route(get,config,routeId);
  if(routeId&&routeId!==ID)return f4ActivityExecutorRoutes_().route(get,config,routeId);
  N(config.projectId==='geapa-convergencia-dev'&&['DEV','SYNTHETIC_LOCAL'].includes(config.environment),'DESTINATION_INVALID');
  const v=get(ROUTE);N(v?.schemaVersion==='geapa-executor-route-v1'&&v.projectId===config.projectId&&v.activityId===ID&&v.commandType===TYPE&&['APPS_SCRIPT_DEV','CF_GEN2_DEV_EDIT_V1'].includes(v.executor)&&Number.isSafeInteger(v.epoch)&&v.epoch>0&&Number.isSafeInteger(v.revision)&&v.revision>0&&['ACTIVE','DRAINING','CLOSED'].includes(v.state)&&Array.isArray(v.actorUids)&&v.actorUids.length>0&&v.actorUids.length<=4&&Array.isArray(v.operationIds)&&v.operationIds.length<=100&&new Set(v.operationIds).size===v.operationIds.length,'ACCESS_REVOKED');return v;
 }
 function fields(get,r){
  const a=get('canonicalActivities/'+ID),marker=get('pilotResources/'+ID);
  N(marker?.synthetic===true&&marker.projectId==='geapa-convergencia-dev'&&marker.activityId===ID&&marker.namespace===''&&a?.id===ID&&a.revision>0&&a.recordMode!=='HISTORICAL_PARTIAL','ACCESS_REVOKED');
  N(r.payload.activityRef.kind==='activity'&&r.target.expected.some(x=>x.kind==='activity'&&x.id===ID&&Number.isSafeInteger(x.revision)&&x.revision>0),'COMMAND_PAYLOAD_INVALID');
  const p=r.payload,equal=(a,b)=>C.hash(F.encode(a))===C.hash(F.encode(b));
  N(Object.keys(p).every(k=>['activityRef','typeRef','title','location','startsAt','endsAt','status','details','decision'].includes(k))&&typeof p.title==='string'&&p.title.length<=300&&p.decision?.reasonCode==='PORTAL_ADMIN_EDIT','COMMAND_PAYLOAD_INVALID');
  for(const k of ['typeRef','location','status'])N(equal(p[k],a[k]),'COMMAND_PAYLOAD_INVALID');
  for(const k of ['startsAt','endsAt'])N(p[k]===C.ms(a[k]),'COMMAND_PAYLOAD_INVALID');
  const rest=d=>Object.fromEntries(Object.entries(d||{}).filter(([k,v])=>k!=='descricaoPublica'&&v!==''));
  N(equal(rest(p.details),rest(a.details)),'COMMAND_PAYLOAD_INVALID');
 }
 function admission(get,r,config,now){
  if(!f4ActFlowExecutor_().supports(r)&&!f4MemberExecutor_().supports(r))return [];
  const previous=get('commandExecutorOwners/'+r.operationId),routeId=previous?.routeId||(selected(r)?ID:f4MemberExecutor_().select(get,r,config)||f4ActFlowExecutor_().select(get,r,config)||f4ActivityExecutorRoutes_().select(get,r,config));if(!routeId){N(!f4ExecutionEnvironment_().production,'ACCESS_REVOKED');return [];}const routes=routeId.startsWith('MEMBERS_')?f4MemberExecutor_():routeId.startsWith('ACT_FLOW_')?f4ActFlowExecutor_():f4ActivityExecutorRoutes_();
  const v=route(get,config,routeId),path='commandExecutorOwners/'+r.operationId,old=get(path);
  N(v.state==='ACTIVE'&&f4ExecutorWindow_().live(v,now)&&v.actorUids.includes(r.ownerUid)&&v.executor===executor(),'ACCESS_REVOKED');if(routeId===ID)fields(get,r);else routes.fields(get,r,v);
  if(old){N(old.executor===executor()&&old.epoch===v.epoch&&old.digest===r.digest&&old.ownerUid===r.ownerUid&&old.ticketId===r.ticketId&&v.operationIds.includes(r.operationId),'OPERATION_CONFLICT');return [];}
  N(v.operationIds.length<Math.min(100,v.maxOperations)&&!v.operationIds.includes(r.operationId),'ACCESS_REVOKED');
  const scoped=routeId.startsWith('ACTIVITY_REAL_')?f4ActivityRealExecutor_().admission(get,r,v,now):[];
  return [...scoped,{path,data:{schemaVersion:'geapa-command-executor-owner-v1',...(routeId===ID?{}:{routeId}),operationId:r.operationId,commandType:r.commandType,activityId:routes.resource?routes.resource(r):r.payload.activityRef.id,ownerUid:r.ownerUid,digest:r.digest,ticketId:r.ticketId,executor:executor(),epoch:v.epoch,createdAt:F.Timestamp.fromMillis(now)}},{path:'commandExecutorRoutes/'+routeId,data:{...v,revision:v.revision+1,operationIds:v.operationIds.concat(r.operationId)}}];
 }
 function recover(get,key,uid,config,now){
  if(!f4MemberExecutor_().supports(key)&&!f4ActFlowExecutor_().supports(key)&&!['ACT.ACTIVITY_SAVE','ACT.ACTIVITY_PUBLICATION_SET','ACT.ACTIVITY_REALIZE'].includes(key?.commandType))return;
  const r=key&&get('commandRequests/'+key.operationId);if(!r)return;
  const existingMemberOwner=get('commandExecutorOwners/'+key.operationId);if(f4MemberExecutor_().supports(key)&&!existingMemberOwner){N(executor()==='APPS_SCRIPT_DEV','ACCESS_REVOKED');return;}
  const o=existingMemberOwner,routeId=o?.routeId||(selected(r)?ID:f4MemberExecutor_().select(get,r,config)||f4ActFlowExecutor_().select(get,r,config)||f4ActivityExecutorRoutes_().select(get,r,config));if(!routeId)return;const routes=routeId.startsWith('MEMBERS_')?f4MemberExecutor_():routeId.startsWith('ACT_FLOW_')?f4ActFlowExecutor_():f4ActivityExecutorRoutes_();
  const v=route(get,config,routeId);
  N(o?.schemaVersion==='geapa-command-executor-owner-v1'&&o.executor===executor()&&o.ownerUid===uid&&o.digest===key.digest&&o.ticketId===r.ticketId&&o.activityId===(routes.resource?routes.resource(r):r.payload.activityRef.id)&&o.commandType===r.commandType&&key.commandType===r.commandType,'ACCESS_REVOKED');
  // A terminal receipt remains pinned to its executor after rollback. It cannot execute again.
  const l=get('commandLedger/'+key.operationId);if(l)return;
  N(['ACTIVE','DRAINING'].includes(v.state)&&v.epoch===o.epoch&&v.executor===o.executor&&v.operationIds.includes(key.operationId)&&v.actorUids.includes(uid)&&now<v.expiresAt,'ACCESS_REVOKED');if(routeId===ID)fields(get,r);else routes.fields(get,r,v);
 }
 // Operator uses a CAS transaction: draining first, then settle every admission.
 function switchPlan(get,config,expectedRevision,nextExecutor,now,routeId){
  routeId=routeId||ID;const v=route(get,config,routeId);N(v.state==='DRAINING'&&v.revision===expectedRevision&&f4ExecutionEnvironment_().executors.includes(nextExecutor),'REVISION_CONFLICT');
  for(const id of v.operationIds){
   const o=get('commandExecutorOwners/'+id),r=get('commandRequests/'+id),l=get('commandLedger/'+id),receipt=get('commandReceipts/'+id),lease=get('commandLeases/'+id),ticket=o&&get('memberAdmissionTickets/'+o.ticketId),use=o&&get('memberAdmissionUses/'+o.ticketId);
   N(o&&o.epoch===v.epoch&&o.executor===v.executor&&ticket,'RECONCILIATION_REQUIRED');
   if(r){N(l&&receipt&&['COMPLETED','FAILED'].includes(l.state)&&l.digest===o.digest&&receipt.operationId===id&&receipt.ownerUid===o.ownerUid&&receipt.state===l.state&&C.hash(F.encode(receipt))===C.hash(F.encode(f4MemberLedgerM13_().toReceipt(l)))&&(!lease||lease.state==='TERMINAL'||C.ms(lease.deadline)<=now),'RECONCILIATION_REQUIRED');}
   else N(!use&&!lease&&!l&&!receipt&&C.ms(ticket.expiresAt)<=now,'RECONCILIATION_REQUIRED');
  }
  return {path:'commandExecutorRoutes/'+routeId,data:{...v,executor:nextExecutor,epoch:v.epoch+1,revision:v.revision+1,state:'ACTIVE',operationIds:[]}};
 }
 return Object.freeze({selected,admission,recover,switchPlan,routePath:ROUTE,activityId:ID});
}

// Routing is an execution fence, not an institutional capability. Both adapters
// use this selector before admission; the original contracts still authorize
// every field, state transition and transactional effect.
function f4ActivityExecutorRoutes_(){
 const C=f4Core_(),N=C.need,ids=['CREATE','EDIT','PUBLISH','HIDE','CANCEL','REOPEN','REALIZE'];
 function operation(r){const p=r?.payload;
  if(r?.commandType==='ACT.ACTIVITY_SAVE')return r.target?.expected?.some(e=>e.kind==='activity'&&e.id===p?.activityRef?.id&&e.exists===false)?'CREATE':'EDIT';
  if(r?.commandType==='ACT.ACTIVITY_PUBLICATION_SET'&&['PUBLISH','HIDE','CANCEL','REOPEN'].includes(p?.action))return p.action;
  return r?.commandType==='ACT.ACTIVITY_REALIZE'?'REALIZE':null;
 }
 function route(get,config,id){
  N(f4ExecutionEnvironment_().destination(config),'DESTINATION_INVALID');
  if(id?.startsWith('ACTIVITY_REAL_'))return f4ActivityRealExecutor_().route(get,config,id);
  N(ids.some(x=>'ACTIVITY_'+x===id),'ACCESS_REVOKED');const v=get('commandExecutorRoutes/'+id);
  const list=(x,max)=>Array.isArray(x)&&x.length>0&&x.length<=max&&new Set(x).size===x.length&&x.every(s=>typeof s==='string'&&/^[A-Za-z0-9_-]{1,128}$/.test(s));
  N(v?.schemaVersion==='geapa-activity-executor-route-v2'&&v.projectId===config.projectId&&v.routeId===id&&v.operation===id.slice(9)&&f4ExecutionEnvironment_().route(v)&&f4ExecutionEnvironment_().executors.includes(v.executor)&&['ACTIVE','DRAINING','CLOSED'].includes(v.state)&&Number.isSafeInteger(v.epoch)&&v.epoch>0&&Number.isSafeInteger(v.revision)&&v.revision>0&&list(v.actorUids,f4ExecutionEnvironment_().production?128:16)&&list(v.cycleIds,8)&&list(v.typeIds,16)&&Array.isArray(v.resourceIds)&&v.resourceIds.length<=100&&v.resourceIds.every(x=>/^[A-Za-z0-9_-]{1,128}$/.test(x))&&Array.isArray(v.operationIds)&&v.operationIds.length<=100&&new Set(v.operationIds).size===v.operationIds.length&&Number.isInteger(v.maxOperations)&&v.maxOperations>0&&v.maxOperations<=100&&f4ExecutorWindow_().valid(v),'ACCESS_REVOKED');return v;
 }
 function matches(get,r,v){
  const op=operation(r),id=r.payload?.activityRef?.id,a=get('canonicalActivities/'+id),create=op==='CREATE',type=create?r.payload.typeRef:a?.typeRef;
  const template=create&&!r.payload.cycleRef&&type?get('activityTypes/'+type.id):null,policy=template?.attendancePolicyRef?get('activityPolicies/'+template.attendancePolicyRef.id):null,cycle=create?r.payload.cycleRef||policy?.cycleRef:a?.cycleRef;
  if(a?.presentationSessionRef||a?.recordMode==='HISTORICAL_PARTIAL')return false;
  if(!cycle||!type||!v.cycleIds.includes(cycle.id)||!v.typeIds.includes(type.id)||v.resourceIds.length&&!v.resourceIds.includes(id))return false;
  const scope=get('activityExecutorScopes/'+cycle.id);
  N(f4ExecutionEnvironment_().scope(scope)&&scope.cycleId===cycle.id,'ACCESS_REVOKED');return true;
 }
 function select(get,r,config){const op=operation(r);if(!op)return null;const id='ACTIVITY_'+op,real=f4ExecutionEnvironment_().production?null:f4ActivityRealExecutor_().select(get,r,config);
  if(!get('commandExecutorRoutes/'+id))return real;const v=route(get,config,id),synthetic=matches(get,r,v);
  N(!(real&&synthetic),'ACCESS_REVOKED');return real||(synthetic?id:null);
 }
 function fields(get,r,v){
  const op=operation(r),id=r.payload?.activityRef?.id,a=get('canonicalActivities/'+id);
  N(v.operation===op&&r.payload.activityRef.kind==='activity','ACCESS_REVOKED');
  if(v.routeId.startsWith('ACTIVITY_REAL_'))f4ActivityRealExecutor_().fields(get,r,v);else N(matches(get,r,v),'ACCESS_REVOKED');
  // Linked presentation sessions have their own capacity/reorganization writes.
  // They remain on the existing executor until that command family is migrated.
  N(!a?.presentationSessionRef&&a?.recordMode!=='HISTORICAL_PARTIAL','UNRESOLVED_DEPENDENCY');
  const expected=r.target?.expected?.find(x=>x.kind==='activity'&&x.id===id);
  if(op==='CREATE')N(!a&&id==='act_'+r.operationId&&expected&&expected.exists===false,'COMMAND_PAYLOAD_INVALID');
  else N(a&&Number.isSafeInteger(expected?.revision)&&expected.revision>0,'COMMAND_PAYLOAD_INVALID');
  const p=r.payload;
  if(op==='PUBLISH'&&v.publicationVisibilities)publication(v.publicationVisibilities,p.visibility||a?.publishedVisibility||'BOARD');
  if(op==='EDIT'){
   editFields(r,a);
  }else if(op==='CREATE'){
   N(p.decision?.reasonCode==='PORTAL_MODEL_CREATION'&&Object.keys(p).every(k=>['activityRef','typeRef','cycleRef','academicSemesterRef','title','location','startsAt','endsAt','status','details','modelContent','decision'].includes(k))&&p.status==='SCHEDULED','COMMAND_PAYLOAD_INVALID');
  }
 }
 function editFields(r,a){const p=r.payload;N(p.decision?.reasonCode==='PORTAL_ADMIN_EDIT'&&Object.keys(p).every(k=>['activityRef','typeRef','title','location','startsAt','endsAt','status','details','decision',...(a?.presentationSessionRef?['sessionRef','calendarAssessmentRef']:[])].includes(k))&&p.typeRef?.id===a?.typeRef?.id&&p.status===a?.status,'COMMAND_PAYLOAD_INVALID');}
 function publication(list,visibility){N(Array.isArray(list)&&list.length>0&&list.length<=3&&new Set(list).size===list.length&&list.every(v=>['BOARD','MEMBERS','PUBLIC'].includes(v))&&list.includes(visibility),'ACCESS_REVOKED');}
 return Object.freeze({operation,route,select,fields,editFields,publication,ids});
}

// Protected execution configuration only; no institutional grants or business facts.
// Intent IDs are approved before admission. Eligibility of the created aggregate
// derives from the existing atomic canonical commit + ledger, never a second registry.
function f4ActivityRealExecutor_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),E=f4ExecutionEnvironment_(),P=E.project,MODE=E.production?'PROD_SCOPED_OPERATIONAL':'DEV_SCOPED_OPERATIONAL',INDEX='activityRealExecutorIndex/current',PREFIX='activityRealExecutorScopes/';
 const ops=['CREATE','EDIT','PUBLISH','HIDE','CANCEL','REOPEN'],id=x=>typeof x==='string'&&/^[A-Za-z0-9_-]{1,64}$/.test(x),opId=x=>typeof x==='string'&&/^[a-f0-9]{32}$/.test(x),hex=x=>typeof x==='string'&&/^[a-f0-9]{64}$/.test(x),rev=x=>Number.isSafeInteger(x)&&x>0;
 const hash=x=>C.hash(F.encode(x)),unique=x=>Array.isArray(x)&&new Set(x).size===x.length;
 const routeId=(s,op)=>'ACTIVITY_REAL_'+s+'_'+op;
 function environment(config){N(E.destination(config),'DESTINATION_INVALID');}
 function scope(get,key){const s=get(PREFIX+key);
  N(id(key)&&s?.schemaVersion==='geapa-activity-real-scope-v1'&&s.id===key&&s.projectId===P&&s.synthetic===false&&rev(s.revision)&&rev(s.configRevision)&&['ACTIVE','DRAINING','CLOSED'].includes(s.state),'ACCESS_REVOKED');
  N(s.mode===MODE&&f4ExecutorWindow_().valid(s)&&s.expiresAt-s.startsAt<=48*3600000&&Number.isSafeInteger(s.authorizationExpiresAt)&&s.expiresAt<=s.authorizationExpiresAt,'ACCESS_REVOKED');
  N(unique(s.actorUids)&&s.actorUids.length===1&&s.actorUids.every(id)&&Number.isInteger(s.maxOperations)&&s.maxOperations>0&&s.maxOperations<=20&&unique(s.operationIds)&&s.operationIds.length<=s.maxOperations&&s.operationIds.every(opId),'ACCESS_REVOKED');
  N(Array.isArray(s.intents)&&s.intents.length<=3,'ACCESS_REVOKED');
  for(const i of s.intents)N(opId(i.operationId)&&i.activityId==='act_'+i.operationId&&hex(i.payloadHash)&&hex(i.portalPayloadHash)&&id(i.cycleId)&&id(i.typeId)&&rev(i.cycleRevision)&&rev(i.typeRevision)&&(i.policyId===null&&i.policyRevision===null||id(i.policyId)&&rev(i.policyRevision)),'ACCESS_REVOKED');
  for(const k of ['operationId','activityId','payloadHash','portalPayloadHash'])N(new Set(s.intents.map(i=>i[k])).size===s.intents.length,'ACCESS_REVOKED');return s;
 }
 function all(get,config){const index=get(INDEX);if(!index)return [];environment(config);
  N(index.schemaVersion==='geapa-activity-real-index-v1'&&index.projectId===P&&rev(index.revision)&&unique(index.scopeIds)&&index.scopeIds.length<=4&&index.scopeIds.every(id),'ACCESS_REVOKED');
  const scopes=index.scopeIds.map(key=>scope(get,key)),intents=scopes.flatMap(s=>s.intents);
  for(const k of ['activityId','operationId','portalPayloadHash','payloadHash'])N(new Set(intents.map(i=>i[k])).size===intents.length,'ACCESS_REVOKED');return scopes;
 }
 function route(get,config,key){environment(config);const v=get('commandExecutorRoutes/'+key);
  N(v?.schemaVersion==='geapa-activity-real-route-v1'&&v.projectId===P&&id(v.scopeId)&&ops.includes(v.operation)&&key===routeId(v.scopeId,v.operation)&&v.routeId===key&&v.syntheticOnly===false&&rev(v.scopeConfigRevision)&&rev(v.epoch)&&rev(v.revision)&&['ACTIVE','DRAINING','CLOSED'].includes(v.state)&&E.executors.includes(v.executor),'ACCESS_REVOKED');
  N(unique(v.actorUids)&&v.actorUids.length===1&&v.actorUids.every(id)&&unique(v.operationIds)&&v.operationIds.every(opId)&&v.operationIds.length<=20&&Number.isInteger(v.maxOperations)&&v.maxOperations>0&&v.maxOperations<=20&&v.mode===MODE&&f4ExecutorWindow_().valid(v)&&v.expiresAt-v.startsAt<=48*3600000,'ACCESS_REVOKED');return v;
 }
 function bound(get,v){const s=scope(get,v.scopeId);N(s.configRevision===v.scopeConfigRevision&&hash(s.actorUids)===hash(v.actorUids)&&v.startsAt>=s.startsAt&&v.expiresAt<=s.expiresAt,'ACCESS_REVOKED');return s;}
 function select(get,r,config){const scopes=all(get,config),claims=scopes.filter(s=>s.intents.some(i=>i.activityId===r.payload?.activityRef?.id||i.operationId===r.operationId));
  if(!scopes.length)return null;
  N(claims.length<=1,'ACCESS_REVOKED');if(!claims.length){
   // An unapproved CREATE for a reserved cycle/model must not escape to the
   // Apps Script admission path merely by generating another operation ID.
   if(f4ActivityExecutorRoutes_().operation(r)==='CREATE'){
    const type=r.payload?.typeRef,model=type&&get('activityTypes/'+type.id),policy=model?.attendancePolicyRef&&get('activityPolicies/'+model.attendancePolicyRef.id),cycle=r.payload?.cycleRef||policy?.cycleRef;
    N(!scopes.some(s=>s.intents.some(i=>i.typeId===type?.id&&i.cycleId===cycle?.id)),'ACCESS_REVOKED');
   }return null;
  }
  const op=f4ActivityExecutorRoutes_().operation(r);N(ops.includes(op),'ACCESS_REVOKED');const key=routeId(claims[0].id,op);bound(get,route(get,config,key));return key;
 }
 function fields(get,r,v){const s=bound(get,v),i=s.intents.find(i=>i.activityId===r.payload?.activityRef?.id),a=get('canonicalActivities/'+r.payload?.activityRef?.id);
  N(i&&['ACTIVE','DRAINING'].includes(s.state),'ACCESS_REVOKED');
  N(a?.recordMode!=='HISTORICAL_PARTIAL'&&!a?.presentationSessionRef,'ACCESS_REVOKED');
  const cycle=get('presentationCycles/'+i.cycleId),type=get('activityTypes/'+i.typeId),policy=i.policyId&&get('activityPolicies/'+i.policyId);
  N(cycle?.recordMode==='NEW_V2'&&cycle.revision===i.cycleRevision&&type?.recordMode==='NEW_V2'&&type.revision===i.typeRevision&&type.status==='ACTIVE'&&!type.portalModel?.memberPresentation,'ACCESS_REVOKED');
  N(type.attendancePolicyRef?.id===i.policyId||!type.attendancePolicyRef&&i.policyId===null,'ACCESS_REVOKED');if(i.policyId)N(policy?.revision===i.policyRevision,'ACCESS_REVOKED');
  if(v.operation==='CREATE')N(!a&&r.operationId===i.operationId&&hash(r.payload)===i.payloadHash&&r.payload.cycleRef?.id===i.cycleId&&r.payload.typeRef?.id===i.typeId,'COMMAND_PAYLOAD_INVALID');
  else{
   const owner=get('commandExecutorOwners/'+i.operationId),ledger=get('commandLedger/'+i.operationId);
   N(a?.recordMode==='NEW_V2'&&a.cycleRef?.id===i.cycleId&&a.typeRef?.id===i.typeId&&owner?.routeId===routeId(s.id,'CREATE')&&owner.activityId===i.activityId&&ledger?.state==='COMPLETED'&&ledger.digest===owner.digest,'ACCESS_REVOKED');
  }
  if(v.operation==='PUBLISH')N(['BOARD','MEMBERS'].includes(r.payload.visibility||a?.publishedVisibility||'BOARD'),'ACCESS_REVOKED');
  return s;
 }
 function admission(get,r,v,now){const s=fields(get,r,v);N(s.state==='ACTIVE'&&f4ExecutorWindow_().live(s,now)&&s.actorUids.includes(r.ownerUid)&&s.operationIds.length<s.maxOperations&&!s.operationIds.includes(r.operationId),'ACCESS_REVOKED');
  return [{path:PREFIX+s.id,data:{...s,revision:s.revision+1,operationIds:s.operationIds.concat(r.operationId)}}];
 }
 function portalHash(p){return hash({idConfig:p.idConfig,atividade:p.atividade});}
 // This returns only a pre-approved operation ID. The prepared command's full
 // canonical payload is checked again at admission and before any effect.
 function preview(get,config,p,uid,now){const scopes=all(get,config),mark=portalHash(p),matches=scopes.flatMap(s=>s.intents.filter(i=>i.portalPayloadHash===mark).map(i=>({s,i})));
  N(matches.length<=1,'ACCESS_REVOKED');if(!matches.length){N(!scopes.some(s=>s.intents.some(i=>i.typeId===p.idConfig&&(!p.atividade?.cycleId||i.cycleId===p.atividade.cycleId))),'ACCESS_REVOKED');return null;}
  const {s,i}=matches[0],v=route(get,config,routeId(s.id,'CREATE'));bound(get,v);
  N(s.state==='ACTIVE'&&v.state==='ACTIVE'&&s.actorUids.includes(uid)&&f4ExecutorWindow_().live(s,now)&&f4ExecutorWindow_().live(v,now),'ACCESS_REVOKED');return {executor:v.executor,routeId:v.routeId,creationOperationId:i.operationId};
 }
 return Object.freeze({routeId,route,scope,select,fields,admission,preview,portalHash,hash});
}
