// Host I/O exchanges encoded values; the shared domain keeps its own prototypes.
const nodeProfile=Object.freeze(JSON.parse(__profileJson));
const codec=f4Firestore_(),nodeConfig=Object.freeze({environment:nodeProfile.environment,projectId:nodeProfile.projectId,database:'(default)',host:'firestore.googleapis.com',port:443});
f4ExecutionConfig_=()=>nodeConfig;
const nodeWire=x=>JSON.parse(__io(JSON.stringify(x)));
const nodeRead=p=>{const v=nodeWire({kind:'get',path:p});return v===null?undefined:codec.decode(v);};
const nodeTx={get:p=>({data:()=>nodeRead(p)}),getAll:paths=>{nodeWire({kind:'getAll',paths});return paths.map(p=>({data:()=>nodeRead(p)}));},set:(p,d)=>nodeWire({kind:'write',path:p,data:codec.encode(d),mode:'set'}),update:(p,d)=>nodeWire({kind:'write',path:p,data:codec.encode(d),mode:'update'}),delete:p=>nodeWire({kind:'write',path:p,mode:'delete'})};
const nodeDb={get:nodeRead,getMany:paths=>{nodeWire({kind:'getAll',paths});return paths.map(nodeRead);},transaction:fn=>{nodeInspect(nodeTx);return fn(nodeTx);},query:plan=>nodeWire({kind:'query',plan:{...plan,filters:plan.filters.map(f=>({...f,value:codec.encode(f.value)})),...(plan.cursor?{cursor:plan.cursor.map(codec.encode)}:{})}}).map(r=>({...r,data:codec.decode(r.data)}))};
const nodeInspect=tx=>{f4Core_().need(tx===nodeTx,'M13_TRANSACTION_REQUIRED');return nodeWire({kind:'inspect'});};
f4Firestore_.value_=Object.freeze({...codec,connect:c=>{f4Core_().destination(c);nodeWire({kind:'guard'});return nodeDb;},trusted:(db,c)=>{f4Core_().destination(c);f4Core_().need(db===nodeDb,'DESTINATION_INVALID');nodeWire({kind:'guard'});return db;},isSynthetic:db=>{nodeWire({kind:'guard'});return db===nodeDb;},inspectTransaction:nodeInspect,previewTransaction:(tx,effects)=>{nodeInspect(tx);return nodeWire({kind:'preview',effects:effects.map(w=>({...w,data:codec.encode(w.data)}))});}});
// The host enforces an exclusive DEV pilot namespace and protected control on every attempt.
f4DevBoundary_=()=>({check:()=>nodeWire({kind:'guard'}),permit:c=>{nodeWire({kind:'guard'});return nodeConfig.environment==='DEV'&&c?.projectId===nodeConfig.projectId&&c?.environment==='DEV'&&c?.database==='(default)';},get root(){return nodeWire({kind:'pilotConfig'}).root;},project:nodeConfig.projectId,path:p=>p});
f4DevReadTiming_=()=>({stage:(name,fn)=>fn(),measure:(name,fn)=>fn(),command:()=>{},rpc:()=>{}});
f4CommandExecutorId_=()=>nodeWire({kind:'pilotConfig'}).executor;
const nodeSettings=()=>nodeWire({kind:'pilotConfig'}).settings;
f4ActivitiesAdmissionLimits_=uid=>({counterId:'ACT_'+f4Core_().hash(uid).slice(0,32),maximum:120,limit:nodeSettings().admissionsPerHour});
// No implicit DEV_ADMIN: the original named profile contract remains active.
function nodeIdentity(v){f4Core_().exact(v,['uid','expiresAt']);globalThis.nodeActor=v;}
function nodePipeline(){const m=f4MemberProtectedInputs_();return f4CorePipeline_({config:nodeConfig,clock:Date.now,verifyIdentity:()=>nodeActor,membersAuthorization:m.authorization,membersRevisions:m.revisions});}
function nodePrepare(input){const ports=f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:nodeActor,settings:nodeSettings()});return f4ActivitiesActivityCompat_({config:nodeConfig,ports,clock:Date.now}).prepare(input);}
function nodePreview(input){const ports=f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:nodeActor,settings:nodeSettings()});return f4ActivitiesActivityCompat_({config:nodeConfig,ports,clock:Date.now}).preview(input);}
// Pin only the technical snapshot's start within one SDK transaction attempt.
// Authorization, token validity and the final proof expiry keep the real clock.
const nodeCapacitySource=f4ActivitiesCapacitySource_;
f4ActivitiesCapacitySource_=options=>nodeCapacitySource({...options,clock:()=>nodeWire({kind:'preparationClock'})});
function nodeFlowPrepare({input,portal}){
 const ports=f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:nodeActor,settings:nodeSettings()}),options={config:nodeConfig,ports,clock:Date.now,pipeline:nodePipeline(),jobs:[]};
 if(!portal){f4Core_().need(f4ActFlowExecutor_().supports(input),'COMMAND_TYPE_INVALID');return ports.prepare(input);}
 const route=input.route;if(route==='/admin/atividades/conferencia')return f4ActivityReview_().prepare(options,input);if(f4ActRemainingConsumers_().routes.includes(route))return f4ActRemainingConsumers_().prepare(options,input);
 if(route==='/atividades/chamada/preparar'){f4Core_().exact(input,['operationId','route','payload']);f4Core_().exact(input.payload,['activityRef']);const rosterRef=f4ActivitiesAttendanceRoster_(options).prepare(input.payload.activityRef);return ports.prepare({operationId:input.operationId,commandType:'ACT.ATTENDANCE_ROSTER_APPLY',payload:{activityRef:input.payload.activityRef,rosterRef}});}
 if(route==='/atividades/chamada/salvar'){const {route,...request}=input;return f4ActivitiesPortalCompat_(options).prepareChamada(request);}
 if(['/v2/justificativas/enviar','/v2/justificativas/analisar'].includes(route))return f4ActivitiesFrequencyCompat_(options).prepare(input);
 if(['/admin/justificativas/janelas/abrir','/admin/justificativas/janelas/encerrar'].includes(route))return f4ActivitiesGovernanceCompat_(options).prepare(input);
 if(route==='/v2/apresentacoes/finalizacao')return f4ActivitiesFinalizationCompat_(options).prepare(input);
 if(route==='/v2/apresentacoes/realizacao')return f4ActivitiesRealizationCompat_(options).prepare(input);
 if(route==='/v2/apresentacoes/compatibilidade')return f4ActivitiesCompatibilityCompat_(options).prepare(input);
 if(route==='/v2/apresentacoes/agenda')return f4ActivitiesCalendarCompat_(options).prepare(input);
 if(route==='/v2/apresentacoes/reorganizacao')return f4ActivitiesReorganizationCompat_(options).prepare(input);
 if(route==='/admin/apresentacoes/planejamento/salvar')return f4ActivitiesPlanningCompat_(options).prepare(input);
 if(['/v2/apresentacoes/titulo-eixo/enviar','/v2/apresentacoes/titulo-eixo/revisar','/v2/apresentacoes/titulo-eixo/reprovar','/v2/apresentacoes/titulo-eixo/retirar','/v2/apresentacoes/titulo-eixo/natureza'].includes(route))return f4ActivitiesPresentationCompat_(options).prepare(input);
 if(['/admin/atividades/salvar','/admin/atividades/publicar','/admin/atividades/ocultar','/admin/atividades/cancelar','/admin/atividades/reabrir','/admin/atividades/realizacao'].includes(route))return f4ActivitiesActivityCompat_(options).prepare(input);
 if(route==='/atividades/modelo/criar')return f4ActivitiesActivityCompat_(options).prepare(input);
 f4Core_().fail('COMMAND_TYPE_INVALID');
}
function nodeFlowPortalType(input,optional=false){
 const p=input.payload||{},r=input.route,table={
 '/admin/atividades/conferencia':'ACTIVITY_ADMIN_REVIEW',
 '/admin/atividades/salvar':'ACTIVITY_SAVE','/admin/atividades/publicar':'ACTIVITY_PUBLICATION_SET','/admin/atividades/ocultar':'ACTIVITY_PUBLICATION_SET','/admin/atividades/cancelar':'ACTIVITY_PUBLICATION_SET','/admin/atividades/reabrir':'ACTIVITY_PUBLICATION_SET','/admin/atividades/realizacao':'ACTIVITY_REALIZE',
 '/atividades/chamada/preparar':'ATTENDANCE_ROSTER_APPLY',
 '/atividades/chamada/salvar':{SALVAR:'ATTENDANCE_SAVE_DRAFT',FINALIZAR:'ATTENDANCE_FINALIZE',REABRIR:'ATTENDANCE_REOPEN',REGISTRAR:'ATTENDANCE_RECORD',CORRIGIR:'ATTENDANCE_CORRECT'}[p.operacao],
 '/v2/justificativas/enviar':'ATTENDANCE_JUSTIFICATION_REQUEST','/v2/justificativas/analisar':'ATTENDANCE_JUSTIFICATION_DECIDE',
 '/admin/justificativas/janelas/abrir':'ATTENDANCE_EXCEPTIONAL_WINDOW_OPEN','/admin/justificativas/janelas/encerrar':'ATTENDANCE_EXCEPTIONAL_WINDOW_CLOSE',
 '/admin/apresentacoes/planejamento/salvar':p.kind==='OPEN'?'CYCLE_OPEN':p.kind==='CAPACITY'?'CAPACITY_ASSESSMENT_SEAL':p.kind==='CYCLE'?'CYCLE_CREATE':p.kind==='PREAGENDA'?'HISTORICAL_PREAGENDA_RECOGNIZE':p.kind==='CONTINUITY'?'PRESENTATION_CONTINUITY_CONNECT':p.kind==='SESSION'?'SESSION_CONFIGURE':null,
 '/v2/apresentacoes/compatibilidade':{REQUEST:'COMPATIBILITY_REQUEST',DECIDE:'COMPATIBILITY_DECIDE'}[p.action],
 '/v2/apresentacoes/agenda':{AGENDAR:'CALENDAR_RESERVE',CANCELAR:'CALENDAR_RELEASE'}[p.acao],
 '/v2/apresentacoes/reorganizacao':{PROPOSE:'CALENDAR_CHANGE_PROPOSE',CONFIRM:'CALENDAR_CHANGE_CONFIRM',CONSENT:'JOINT_CONSENT',JOINT:'JOINT_PROPOSE',COMPOSE:'COMPOSITION_CONFIRM'}[p.action],
 '/v2/apresentacoes/titulo-eixo/enviar':'THEME_SUBMIT','/v2/apresentacoes/titulo-eixo/revisar':'THEME_VALIDATE','/v2/apresentacoes/titulo-eixo/reprovar':'THEME_VALIDATE',
 '/v2/apresentacoes/titulo-eixo/retirar':'THEME_WITHDRAW','/v2/apresentacoes/titulo-eixo/natureza':'THEME_MATERIALITY_DECIDE',
 '/v2/apresentacoes/finalizacao':{TITULO:'FINAL_TITLE_CONFIRM',CONFERIR:'OPERATIONAL_CHECK'}[p.acao],'/v2/apresentacoes/realizacao':'REALIZATION_RECOGNIZE','/atividades/modelo/validar':'MEMBER_PRESENTATION_ACTIVITY_CREATE','/atividades/modelo/criar':'MEMBER_PRESENTATION_ACTIVITY_CREATE'};
 let type=f4ActRemainingConsumers_().type(input,nodeRead)||table[r];if(r==='/v2/apresentacoes/titulo-eixo/enviar'&&typeof p.idApresentacao==='string'&&p.idApresentacao.startsWith('APU_')){const u=nodeRead('presentationUnitHeads/'+p.idApresentacao.slice(4)),t=u?.thematicRequestRef&&nodeRead(f4ActivitiesModel_().path(u.thematicRequestRef));if(t?.workflowState==='PAUSED')type='THEME_COMPLEMENT';}if(optional&&!type)return null;f4Core_().need(type&&f4ActFlowExecutor_().commands.includes(type),'COMMAND_TYPE_INVALID');return 'ACT.'+type;
}
function nodeActivityRoute(input){
 const ports=f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:nodeActor,settings:nodeSettings()}),auth=ports.readAuthority(['activityAttendanceRegistry/current'],false),p=input.payload;
 const operations={'/atividades/modelo/criar':'CREATE','/atividades/modelo/validar':'CREATE','/admin/atividades/salvar':'EDIT','/admin/atividades/publicar':'PUBLISH','/admin/atividades/ocultar':'HIDE','/admin/atividades/cancelar':'CANCEL','/admin/atividades/reabrir':'REOPEN','/admin/atividades/realizacao':'REALIZE'},op=operations[input.route];
 f4Core_().need(op&&p,'COMMAND_TYPE_INVALID');
 let payload={activityRef:{kind:'activity',id:p.idAtividade||'route_preview'},action:op,decision:{reasonCode:op==='CREATE'?'PORTAL_MODEL_CREATION':'PORTAL_ADMIN_EDIT'}};
 if(op==='CREATE'){
  const type=nodeRead('activityTypes/'+p.idConfig);f4Core_().need(type,'CONFIG_CAPABILITY_DISABLED');
  if(type.portalModel?.memberPresentation){ports.recheckAuthority(auth);f4Core_().need(nodeConfig.environment!=='PROD','CONFIG_CAPABILITY_DISABLED');return {executor:'APPS_SCRIPT_DEV',routeId:null};}
  const real=f4ActivityRealExecutor_().preview(nodeRead,nodeConfig,p,nodeActor.uid,Date.now());
  if(real){ports.recheckAuthority(auth);return real;}
  const policy=type.attendancePolicyRef&&nodeRead(f4ActivitiesModel_().path(type.attendancePolicyRef));payload={...payload,typeRef:{kind:'activityType',id:type.id},cycleRef:p.atividade?.cycleId?{kind:'cycle',id:p.atividade.cycleId}:policy?.cycleRef};
 }
 const r={commandType:op==='CREATE'||op==='EDIT'?'ACT.ACTIVITY_SAVE':op==='REALIZE'?'ACT.ACTIVITY_REALIZE':'ACT.ACTIVITY_PUBLICATION_SET',target:{expected:op==='CREATE'?[{kind:'activity',id:payload.activityRef.id,exists:false}]:[]},payload},routes=f4ActivityExecutorRoutes_(),id=routes.select(nodeRead,r,nodeConfig);
 if(!id){ports.recheckAuthority(auth);f4Core_().need(nodeConfig.environment!=='PROD','CONFIG_CAPABILITY_DISABLED');return {executor:'APPS_SCRIPT_DEV',routeId:null};}
 const v=routes.route(nodeRead,nodeConfig,id),now=Date.now();f4Core_().need(v.state==='ACTIVE'&&v.actorUids.includes(nodeActor.uid)&&v.startsAt<=now&&now<v.expiresAt,'ACCESS_REVOKED');ports.recheckAuthority(auth);return {executor:v.executor,routeId:id};
}
function nodeProject(ids){return f4PortalProjectionRebuild_({config:nodeConfig,identity:nodeActor,settings:nodeSettings(),clock:Date.now},ids);}
function nodeProjectionInspect(id){return f4PortalProjectionRecoveryInspect_({config:nodeConfig,identity:nodeActor,settings:nodeSettings(),clock:Date.now},id);}
function nodePilotOperation(key,create){
 const need=(v,code)=>{if(!v)throw Object.assign(new Error(code),{code});};
 const spec=nodeWire({kind:'pilotConfig'});if(spec.shared){f4ExecutorOwnership_().recover(nodeRead,key,nodeActor.uid,nodeConfig,Date.now());return;}const path='pilotOperationOwners/'+key.operationId,old=nodeRead(path),budget=nodeRead('pilotBudget/current');
 if(old){need(old.executor===spec.executor&&old.epoch===spec.epoch,'EXECUTOR_FENCE_CONFLICT');need(old.ownerUid===nodeActor.uid&&old.digest===key.digest,'OPERATION_CONFLICT');return;}
 need(create,'EXECUTOR_FENCE_CONFLICT');
 const control=nodeRead('pilotControl/runtime');need((budget?.count||0)<control.maxOperations,'PILOT_LIMIT_REACHED');
 nodeTx.set(path,{schemaVersion:'geapa-pilot-operation-owner-v1',operationId:key.operationId,executor:spec.executor,epoch:spec.epoch,ownerUid:nodeActor.uid,digest:key.digest,namespace:spec.root,createdAt:codec.Timestamp.fromMillis(Date.now())});
 nodeTx.set('pilotBudget/current',{count:(budget?.count||0)+1});
}

function nodeActRead(request){
 if(request.route==='/members/own'){
  f4Core_().exact(request,['route','params']);
  return f4MemberOwnReader_().wire(f4MemberOwnReader_().projectTransaction(nodeTx,nodeActor,request.params,Date.now()));
 }
 if(f4MemberOperationalReader_({config:nodeConfig,identity:nodeActor,clock:Date.now}).routes.includes(request.route)||f4MemberProfileCorrectionReader_({config:nodeConfig,identity:nodeActor,clock:Date.now}).routes.includes(request.route))return f4MemberOperationalReader_({config:nodeConfig,identity:nodeActor,clock:Date.now}).read(request);
 const ports=f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:nodeActor,settings:nodeSettings()});return f4ActivitiesPortalReader_({config:nodeConfig,ports,clock:Date.now},request);
}

// Both passes execute the original protected download resolver. No file ID,
// provider settings or authorization is accepted from the browser.
function nodeWorkspaceRead(input){
 let plan=null;const ports=f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:nodeActor,settings:nodeSettings()}),options={config:nodeConfig,ports,clock:Date.now,providerFor:settings=>({download:(request,version)=>{
  const value={settings,request,version};f4Core_().need(!plan,'SCOPE_TOO_LARGE');plan=value;
  if(input.expected)f4ActivitiesModel_().need(f4Core_().hash(codec.encode(value))===f4Core_().hash(codec.encode(input.expected)),'STALE_REVISION');
  return input.bytes||[];
 }})};
 const q=input.request;let result;
 if(/^APR-/.test(q.params?.idApresentacao||'')){
  const value=f4ActivityHistoricalFilePlan_(options,q);
  if(input.expected)f4ActivitiesModel_().need(f4Core_().hash(codec.encode(value))===f4Core_().hash(codec.encode(input.expected)),'STALE_REVISION');
  return {plan:value,result:null};
 }
 if(q.route==='/v2/apresentacoes/material/arquivo')result=f4ActivitiesMaterialDownload_(options).read(q.params);
 else if(q.route==='/v2/apresentacoes/foto/arquivo')result=f4ActivitiesPhotoCompat_(options).download(q.params);
 else if(q.route==='/v2/justificativas/documento/arquivo')result=f4ActivitiesAttendanceProofCompat_(options).download(q.params);
 else f4Core_().fail('COMMAND_TYPE_INVALID');
 f4Core_().need(plan,'UNRESOLVED_DEPENDENCY');return {plan,result};
}

function nodeStoragePhase(input){
 const ports=f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:nodeActor,settings:nodeSettings()}),options={config:nodeConfig,pipeline:nodePipeline(),authorize:ports.authorizeUpload,providerFor:()=>f4Core_().fail('EXTERNAL_SERVICE_DISABLED'),clock:Date.now};
 const phases=f4ActivitiesStorageTransfer_(options),upload=f4ActivitiesWorkspaceUpload_(options).phases;
 f4Core_().need(['source','reservation','transportPlan','acquire','initiating','saveSession','release','digest','commandPlan','attempt','evidence'].includes(input.name),'COMMAND_TYPE_INVALID');
 const fn=phases[input.name]||upload[input.name];return codec.encode(fn(...codec.decode(input.args)));
}
function nodeStorageSystem(plan){
 const p=plan.principal,q=plan.request;
 f4Core_().need(['ACT.OUTBOX_CLAIM','ACT.OUTBOX_OBSERVE','ACT.OUTBOX_RECONCILE','ACT.CORRECTION_NOTICE_RECORD'].includes(q.commandType),'COMMAND_TYPE_INVALID');
 const input={schemaVersion:'geapa-act-command-v13',operationId:q.operationId,commandType:q.commandType,target:{scope:q.scope,expected:q.expected,authorization:{kind:p.kind,servicePrincipal:p.servicePrincipal,scope:p.scope,ruleRef:p.ruleRef,jobKind:p.jobKind,runId:p.runId}},payload:q.payload,evidenceRefs:[],decisionRef:null};
 globalThis.storageSystem=nodePipeline().systemPhases(input);return storageSystem.key;
}

function nodeStorageNotice(input){
 const ports=f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:nodeActor,settings:nodeSettings()});
 return codec.encode(f4ActivitiesPortalCorrectionNotice_({config:nodeConfig,pipeline:nodePipeline(),ports,clock:Date.now}).prepare(input.request,input.readOnly===true));
}

// Stabilize only audit identity within a read transaction; authorization keeps the real clock.
const nodeCorrectionReader=f4MemberProfileCorrectionReader_;
f4MemberProfileCorrectionReader_=options=>nodeCorrectionReader({...options,auditClock:()=>nodeWire({kind:"auditClock"})});
