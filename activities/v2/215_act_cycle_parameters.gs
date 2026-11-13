// Cycle configuration is still presentationCycles/*/configRevisions. Evidence
// annotates the existing parameter values; it is not another configuration store.
function f4CycleParameterDocumentHash_(v){return f4Core_().hash(f4Firestore_().encode(v));}
// A catalog-only approval cannot certify deadlines copied into its technical
// payload. Interpret the documented approval scope without rewriting the record.
// Recognition or a new, governed approval remains the only way to establish them.
function f4CycleEffectiveConfiguration_(cfg){
 if(!cfg||cfg.approvalScope!=='OPERATIONAL_AXIS_CATALOG_AND_CLOSED_FEATURES')return cfg;
 const parameters={...cfg.parameters},items={...parameters.parameterEvidence?.items};
 for(const key of ['newDateMinimumDays','thematicMinimumDays','titleMinimumDays','dynamicRequestLeadDays']){
  items[key]={status:'UNESTABLISHED',basis:'A aprovação deste registro abrange o catálogo de eixos e funcionalidades fechadas, não comprova este prazo.'+(Object.hasOwn(parameters,key)?' Valor técnico anteriormente registrado: '+parameters[key]+'.':'')};delete parameters[key];
 }
 parameters.parameterEvidence={schemaVersion:'limited-legacy-approval-view',items};return {...cfg,parameters};
}
function f4CycleMeetingDefaults_(get,ref){const M=f4ActivitiesModel_(),c=get(M.path(ref)),cfg=c?.currentConfigRef?.kind==='config'&&get(M.path(c.currentConfigRef)),P=f4CycleParameters_(),p=cfg?.approved?cfg.parameters:{};return {id:ref.id,label:c?.cycleLabel||ref.id,meetingStartTime:P.entry(p,'meetingStartTime').value||null,meetingEndTime:P.entry(p,'meetingEndTime').value||null,configRevision:cfg?.id||null};}
function f4CycleParameterCatalog_(){
 const field=(key,label,section,type,extra={})=>({key,label,section,type,...extra});
 return [
  field('calendarMode','Formação da agenda',3,'select',{choices:['BOARD_INITIAL','MEMBER_SELF_SCHEDULE']}),
  field('planningMode','Base do planejamento',1,'select',{choices:['INCIDENCE','RECOGNIZED_PREAGENDA']}),
  field('preAgendaPolicyRef','Reconhecimento da pré-agenda',1,'policy',{usage:'HISTORICAL_PREAGENDA'}),
  field('meetingStartTime','Horário-padrão de início',3,'time'),field('meetingEndTime','Horário-padrão de término',3,'time'),
  field('selfSchedulingWindow','Período de autoagendamento',3,'window'),field('reorganizationWindow','Período de reorganização',3,'window'),
  field('administrativeAllocationCriteria','Critérios de alocação administrativa',3,'text'),
  field('reserveDates','Datas de reserva',3,'dates'),field('shared','Sessões compartilhadas',3,'boolean'),
  field('maxMemberPresentationsPerSession','Máximo de apresentações de membros por sessão',3,'integer',{min:1,max:2}),
  field('exclusiveSessionCriteria','Critérios de exclusividade',3,'text'),
  field('annualPresentationRequired','Obrigação anual de apresentação',4,'boolean'),
  field('enabledModalities','Modalidades habilitadas',4,'choices',{choices:['ORDINARY_INDIVIDUAL','ORDINARY_SHARED','MEMBER_DYNAMIC']}),
  field('dynamic','Apresentação dinâmica de membro',4,'boolean'),field('dynamicLimit','Limite de dinâmicas de membros no ciclo',4,'integer',{min:0,max:180}),
  field('early','Reserva temática antecipada',4,'boolean'),field('jointRequestsEnabled','Solicitações conjuntas',4,'boolean'),
  field('reschedulingCriteria','Critérios de remarcação',4,'text'),field('cancellationCriteria','Critérios de cancelamento',4,'text'),
  field('axisGuideRef','Versão do Guia de Eixos Temáticos',5,'policy',{usage:'AXIS_GUIDE'}),
  field('axisRefs','Eixos oficiais aplicáveis',5,'policies',{usage:'THEMATIC_AXIS'}),
  field('principalAxisRequired','Eixo principal obrigatório',5,'boolean'),field('secondaryAxisEnabled','Eixo secundário permitido',5,'boolean'),
  field('quota','Cotas temáticas',5,'boolean'),field('clockPolicyRef','Prazo e procedimento de análise temática',5,'policy',{usage:'THEMATIC_CLOCK'}),
  field('compatibilityGuideRef','Critérios de compatibilidade temática',5,'policy',{usage:'COMPATIBILITY_GUIDE'}),
  field('newDateMinimumDays','Antecedência de nova data (dias)',6,'integer',{min:0,max:3660}),
  field('thematicMinimumDays','Regularização temática e eixo principal (dias)',6,'integer',{min:0,max:3660}),
  field('titleMinimumDays','Título definitivo (dias)',6,'integer',{min:0,max:3660}),
  field('dynamicRequestLeadDays','Solicitação de dinâmica de membro (dias)',6,'integer',{min:0,max:3660}),
  field('jointAcceptanceHours','Aceite de solicitação conjunta (horas)',6,'integer',{min:1,max:8760}),
  field('reconsiderationHours','Solicitação de reconsideração (horas)',6,'integer',{min:1,max:8760}),
  field('ordinaryMaterialRequirementRef','Materiais da apresentação ordinária',7,'policy',{usage:'MATERIAL_REQUIREMENT'}),
  field('dynamicMaterialRequirementRef','Materiais da dinâmica de membro',7,'policy',{usage:'MATERIAL_REQUIREMENT'}),
  field('publicationPolicyRef','Política de divulgação',7,'policy',{usage:'PUBLICATION'}),
  field('operationalChecklist','Conferência para divulgação',7,'checklist')
 ];
}
function f4CycleParameters_(){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),catalog=f4CycleParameterCatalog_(),byKey=Object.fromEntries(catalog.map(f=>[f.key,f]));
 const schema='cycle-parameters-v1',statuses=['DEFINED','NOT_APPLICABLE','UNESTABLISHED'];
 const isManaged=p=>p?.parameterEvidence?.schemaVersion===schema;
 function entry(p,key){
  const e=p?.parameterEvidence?.items?.[key];
  if(e)return {...e,...(e.status==='DEFINED'&&Object.hasOwn(p,key)?{value:p[key]}:{})};
  // An explicit legacy value is visible with its legacy provenance. Absence is
  // unknown, including zero/false, and never inherits a later cycle's minimum.
  return p&&Object.hasOwn(p,key)?{status:'DEFINED',value:p[key],basis:'Configuração anterior registrada',legacy:true}:{status:'UNESTABLISHED',basis:''};
 }
 function value(p,key){const e=entry(p,key);N(e.status==='DEFINED'&&Object.hasOwn(e,'value'),'CONFIG_CAPABILITY_DISABLED','PARAMETER_NOT_ESTABLISHED_'+key);return e.value;}
 function days(p,key){const v=value(p,key);N(Number.isSafeInteger(v)&&v>=0,'CONFIG_CAPABILITY_DISABLED','INVALID_CYCLE_DEADLINE');return v;}
 function deadline(p,key){const e=entry(p,key);if(isManaged(p)&&e.status==='NOT_APPLICABLE')return {applicable:false};return {applicable:true,days:days(p,key)};}
 // Read models may display an unknown rule. Only the affected action is closed;
 // neither zero nor another cycle's rule is a substitute for missing evidence.
 function dateAvailable(p,key,from,to){const e=entry(p,key),gap=M.gap(from,to);return gap>=0&&(isManaged(p)&&e.status==='NOT_APPLICABLE'||e.status==='DEFINED'&&Number.isSafeInteger(e.value)&&e.value>=0&&gap>=e.value);}
 function modality(p,mode){if(!isManaged(p))return true;const e=entry(p,'enabledModalities');return e.status==='DEFINED'&&Array.isArray(e.value)&&e.value.includes(mode);}
 function axes(p,primary,secondary,get){
  const managed=isManaged(p),required=managed?value(p,'principalAxisRequired'):true;
  N(primary&&(primary.kind==='policy'||managed&&required===false&&primary.kind==='NONE'),'CONFIG_CAPABILITY_DISABLED','PRINCIPAL_AXIS_REQUIRED');
  if(secondary&&secondary.kind!=='NONE'&&managed)N(value(p,'secondaryAxisEnabled')===true,'CONFIG_CAPABILITY_DISABLED','SECONDARY_AXIS_DISABLED');
  for(const r of [primary,secondary].filter(r=>r&&r.kind!=='NONE')){N(r.kind==='policy'&&p.axisRefs?.some(a=>M.path(a)===M.path(r)),'CONFIG_CAPABILITY_DISABLED','AXIS_NOT_CONFIGURED');const d=get(M.path(r));N(d?.approved&&d.active,'CONFIG_CAPABILITY_DISABLED','AXIS_INACTIVE');}
  N(!secondary||secondary.kind==='NONE'||primary.id!==secondary.id,'VALIDATION_ERROR','AXIS_DUPLICATE');
 }
 function reconsideration(p,now){if(!isManaged(p))return {};const e=entry(p,'reconsiderationHours');return e.status==='DEFINED'?{reconsiderationDueAt:M.stamp(now+days(p,'reconsiderationHours')*3600000)}:{};}
 function validateValue(f,v,cycle,get){
  if(f.type==='boolean')N(typeof v==='boolean','VALIDATION_ERROR',f.key);
  else if(f.type==='integer')N(Number.isSafeInteger(v)&&v>=f.min&&v<=f.max,'VALIDATION_ERROR',f.key);
  else if(f.type==='select')N(f.choices.includes(v),'VALIDATION_ERROR',f.key);
  else if(f.type==='choices')N(Array.isArray(v)&&v.length>0&&v.length<=f.choices.length&&new Set(v).size===v.length&&v.every(a=>f.choices.includes(a)),'VALIDATION_ERROR',f.key);
  else if(f.type==='time')N(typeof v==='string'&&/^([01]\d|2[0-3]):[0-5]\d$/.test(v),'VALIDATION_ERROR',f.key);
  else if(f.type==='text')N(typeof v==='string'&&v.trim()&&v.length<=2000,'VALIDATION_ERROR',f.key);
  else if(f.type==='window'){C.exact(v,['opensAt','endsAtExclusive']);N(Number.isSafeInteger(v.opensAt)&&Number.isSafeInteger(v.endsAtExclusive)&&v.opensAt<v.endsAtExclusive,'VALIDATION_ERROR',f.key);}
  else if(f.type==='dates'){N(Array.isArray(v)&&v.length<=32&&new Set(v).size===v.length,'VALIDATION_ERROR',f.key);for(const d of v){M.civil(d);N(M.gap(cycle.civilStart,d)>=0&&M.gap(d,cycle.civilEnd)>=0,'VALIDATION_ERROR','RESERVE_OUTSIDE_CYCLE');}}
  else if(f.type==='policy'||f.type==='policies'){
   const refs=f.type==='policy'?[v]:v;N(Array.isArray(refs)&&refs.length>0&&refs.length<=12,'VALIDATION_ERROR',f.key);
   const seen=new Set();for(const r of refs){N(r?.kind==='policy'&&!seen.has(M.path(r)),'VALIDATION_ERROR',f.key);seen.add(M.path(r));const d=get(M.path(r));N(d?.approved===true&&d.active===true&&d.usage===f.usage&&(!d.cycleRef||d.cycleRef.id===cycle.id),'CONFIG_CAPABILITY_DISABLED','PARAMETER_POLICY_'+f.key);}
  }else if(f.type==='checklist')f4ActivitiesOperationalChecklist_({operationalChecklist:v});
 }
 function normative(policy,cycle,get){
  const g=policy?.parameterGovernance;
  N(policy?.approved===true&&policy.active===true&&g?.schemaVersion===schema&&g.status==='APPROVED'&&g.normativeReference&&g.rules&&g.sources?.length>0&&g.sources.length<=8,'HUMAN_DECISION_REQUIRED','CYCLE_NORMATIVE_BASIS_REQUIRED');
  N(!g.cycleIds||g.cycleIds.includes(cycle.id),'CONFIG_CAPABILITY_DISABLED','NORM_CYCLE_SCOPE');
  for(const s of g.sources){const source=get(M.path(s.ref));N(source&&f4CycleParameterDocumentHash_(source)===s.hash,'STALE_REVISION','NORMATIVE_SOURCE_CHANGED');N(!['DRAFT','MINUTA','PROPOSED'].includes(source.status)&&source.isDraft!==true,'HUMAN_DECISION_REQUIRED','DRAFT_IS_NOT_NORMATIVE');}
  return g;
 }
 function validate(parameters,cycle,policy,get,approved=false){
  N(isManaged(parameters),'VALIDATION_ERROR','PARAMETER_SCHEMA');
  const evidence=parameters.parameterEvidence;C.exact(evidence,['schemaVersion','items']);C.exact(evidence.items,catalog.map(f=>f.key));
  const governance=approved?normative(policy,cycle,get):policy?.parameterGovernance;
  for(const f of catalog){
   const e=evidence.items[f.key];N(e&&typeof e==='object','VALIDATION_ERROR','PARAMETER_EVIDENCE');C.exact(e,['status','basis',...('documentedFrom'in e?['documentedFrom']:[]),...('documentedUntil'in e?['documentedUntil']:[])]);
   N(statuses.includes(e.status)&&typeof e.basis==='string'&&e.basis.length<=2000,'VALIDATION_ERROR','PARAMETER_EVIDENCE');
   if(e.documentedFrom)M.civil(e.documentedFrom);if(e.documentedUntil)M.civil(e.documentedUntil);
   if(e.documentedFrom&&e.documentedUntil)N(M.gap(e.documentedFrom,e.documentedUntil)>=0,'VALIDATION_ERROR','EVIDENCE_INTERVAL');
   const present=Object.hasOwn(parameters,f.key);N(present===(e.status==='DEFINED'),'VALIDATION_ERROR','PARAMETER_VALUE_STATUS');
   if(e.status==='DEFINED')validateValue(f,parameters[f.key],cycle,get);
   if(!approved)continue;
   if(e.status==='UNESTABLISHED')continue;
   N(e.basis.trim(),'HUMAN_DECISION_REQUIRED','PARAMETER_DOCUMENTARY_BASIS');
   const rule=governance.rules[f.key];N(rule&&rule.basis,'HUMAN_DECISION_REQUIRED','PARAMETER_NOT_DELEGATED_'+f.key);
   if(e.status==='NOT_APPLICABLE'){N(rule.allowNotApplicable===true,'HUMAN_DECISION_REQUIRED','NORM_REQUIRES_PARAMETER');continue;}
   const v=parameters[f.key];if(Object.hasOwn(rule,'fixed'))N(M.hash(v)===M.hash(rule.fixed),'HUMAN_DECISION_REQUIRED','NORM_FIXED_'+f.key);
   if(rule.min!==undefined)N(typeof v==='number'&&v>=rule.min,'HUMAN_DECISION_REQUIRED','NORM_MINIMUM_'+f.key);
   if(rule.max!==undefined)N(typeof v==='number'&&v<=rule.max,'HUMAN_DECISION_REQUIRED','NORM_MAXIMUM_'+f.key);
   if(rule.allowed)N((Array.isArray(v)?v:[v]).every(a=>rule.allowed.some(b=>M.hash(a)===M.hash(b))),'HUMAN_DECISION_REQUIRED','NORM_ALLOWED_'+f.key);
   N(Object.hasOwn(rule,'fixed')||rule.editable===true,'HUMAN_DECISION_REQUIRED','PARAMETER_NOT_DELEGATED_'+f.key);
  }
  if(parameters.calendarMode==='MEMBER_SELF_SCHEDULE')N(parameters.selfSchedulingWindow,'HUMAN_DECISION_REQUIRED','SELF_SCHEDULING_WINDOW_REQUIRED');
  if(parameters.meetingStartTime&&parameters.meetingEndTime)N(parameters.meetingStartTime<parameters.meetingEndTime,'VALIDATION_ERROR','MEETING_TIME_INTERVAL');
  if(parameters.shared===false&&parameters.maxMemberPresentationsPerSession!==undefined)N(parameters.maxMemberPresentationsPerSession===1,'VALIDATION_ERROR','SHARED_OFF');
  if(parameters.dynamic===true)N(Number.isSafeInteger(parameters.dynamicLimit)&&parameters.dynamicLimit>0,'VALIDATION_ERROR','DYNAMIC_LIMIT_REQUIRED');
  if(parameters.jointRequestsEnabled===true)N(parameters.shared===true&&Number.isSafeInteger(parameters.jointAcceptanceHours),'VALIDATION_ERROR','JOINT_PARAMETERS_REQUIRED');
  if(parameters.enabledModalities){N(!parameters.enabledModalities.includes('ORDINARY_SHARED')||parameters.shared===true,'VALIDATION_ERROR','MODALITY_SHARED');N(!parameters.enabledModalities.includes('MEMBER_DYNAMIC')||parameters.dynamic===true,'VALIDATION_ERROR','MODALITY_DYNAMIC');}
  if(parameters.quota===true)N(parameters.axisRefs?.length&&parameters.principalAxisRequired===true,'VALIDATION_ERROR','QUOTA_AXES_REQUIRED');
  N(parameters.timezone==='America/Cuiaba','VALIDATION_ERROR','CYCLE_TIMEZONE');
  return parameters;
 }
 function opening(p){if(!isManaged(p))return;for(const key of ['calendarMode','shared','maxMemberPresentationsPerSession','annualPresentationRequired','enabledModalities','dynamic','early','quota'])value(p,key);}
 return {catalog,schema,statuses,isManaged,entry,value,days,deadline,dateAvailable,modality,axes,reconsideration,validate,normative,opening};
}
function f4CycleParameterCommands_(x){
 const {M,N,p,load,save,revision,now,actor}=x,P=f4CycleParameters_();
 function cycle(){const c=load(p.cycleRef,true);N(c.recordMode==='NEW_V2'&&['DRAFT','OPEN'].includes(c.status),'INVALID_STATE_TRANSITION','CYCLE_PARAMETER_STATE');return c;}
 function reason(){N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'HUMAN_DECISION_REQUIRED','PARAMETER_REASON');}
 return {
  CONFIG_DRAFT_SAVE(){
   const c=cycle();reason();N(!c.scheduledConfigRef||c.scheduledConfigRef.kind==='NONE','INVALID_STATE_TRANSITION','ACTIVATE_OR_REVIEW_PENDING_CONFIGURATION');
   const policy=load(c.policyVersionRef,true);P.validate(p.parameters,c,policy,x.get,false);
   N(Number.isSafeInteger(p.effectiveFrom)&&p.effectiveFrom>=now,'INVALID_STATE_TRANSITION','RETROACTIVE');
   const current=c.currentConfigRef&&load(c.currentConfigRef,true),before=current?.parameters||{};
   const changes=f4ActivitiesConfigurationChange_(before,p.parameters,'STRUCTURAL',c,policy,now);
   if(c.status==='OPEN')N(changes.every(k=>['parameterEvidence','selfSchedulingWindow','reorganizationWindow'].includes(k)),'HUMAN_DECISION_REQUIRED','STRUCTURAL_FREEZE');
   const r=M.ref('config',M.hash({operationId:x.operationId,kind:'configDraft'}).slice(0,32),c.id);
   if(c.draftConfigRef?.kind==='config')load(c.draftConfigRef,true);
   save(r,{cycleRef:p.cycleRef,approved:false,state:'DRAFT',parameters:p.parameters,definitionHash:M.hash(p.parameters),effectiveFrom:M.stamp(p.effectiveFrom),policyVersionRef:c.policyVersionRef,policyRevision:policy.revision,policyHash:f4CycleParameterDocumentHash_(policy),predecessorDraftRef:c.draftConfigRef||{kind:'NONE'},predecessorRef:c.currentConfigRef||{kind:'NONE'},reason:p.reason,authoredBy:actor,authoredAt:M.stamp(now)});
   revision(p.cycleRef,{...c,draftConfigRef:r});return {draftConfigRef:r,definitionHash:M.hash(p.parameters),state:'DRAFT'};
  },
  CONFIG_HISTORICAL_RECOGNIZE(){
   const c=cycle();reason();N(c.status==='OPEN'&&(!c.scheduledConfigRef||c.scheduledConfigRef.kind==='NONE'),'INVALID_STATE_TRANSITION','RECOGNITION_OPEN_CYCLE_REQUIRED');
   const approval=load(p.recognitionRef,true),prior=c.currentConfigRef?load(c.currentConfigRef,true):null;
   N(approval?.usage==='CYCLE_PARAMETER_RECOGNITION'&&approval.approved===true&&approval.active===true&&approval.cycleRef?.id===c.id&&approval.expectedCycleRevision===c.revision&&M.ms(approval.issuedAt)<=now&&now<M.ms(approval.expiresAt),'HUMAN_DECISION_REQUIRED','PARAMETER_RECOGNITION_REQUIRED');
   N(approval.parametersHash===M.hash(p.parameters)&&approval.currentConfigurationHash===f4CycleParameterDocumentHash_(prior)&&approval.preservesAcquiredRights===true&&approval.documentaryEvidence,'STALE_REVISION','RECOGNITION_CONTENT_CHANGED');
   N(Array.isArray(approval.sources)&&approval.sources.length>0&&approval.sources.length<=8,'HUMAN_DECISION_REQUIRED','RECOGNITION_SOURCE_REQUIRED');
   for(const source of approval.sources){const d=load(source.ref,true);N(f4CycleParameterDocumentHash_(d)===source.hash,'STALE_REVISION','RECOGNITION_SOURCE_CHANGED');N(!['DRAFT','MINUTA','PROPOSED'].includes(d.status)&&d.isDraft!==true,'HUMAN_DECISION_REQUIRED','DRAFT_IS_NOT_HISTORICAL_EVIDENCE');}
   P.validate(p.parameters,c,load(c.policyVersionRef),x.get,false);
   for(const e of Object.values(p.parameters.parameterEvidence.items))N(e.status==='UNESTABLISHED'||e.basis.trim(),'HUMAN_DECISION_REQUIRED','PARAMETER_DOCUMENTARY_BASIS');
   const r=M.ref('config',M.hash({operationId:x.operationId,kind:'historicalConfig'}).slice(0,32),c.id);
   save(r,{cycleRef:p.cycleRef,approved:true,state:'HISTORICAL_RECOGNITION',parameters:p.parameters,parameterClasses:'DOCUMENTARY',effectiveFrom:M.stamp(now),predecessorRef:c.currentConfigRef||{kind:'NONE'},policyVersionRef:c.policyVersionRef,recognitionRef:p.recognitionRef,recognizedAt:M.stamp(now),recognition:{reason:p.reason,documentaryEvidence:approval.documentaryEvidence,sources:approval.sources,originalApproval:approval.originalApproval||{precision:'UNKNOWN'}},approvalDecision:{actor,authorityBasis:x.authorityBasis,decidedAt:M.stamp(now),basis:{reason:p.reason,reasonCode:'DOCUMENTARY_RECOGNITION'}}});
   // Only configuration and its fence change. No activity, reconciliation,
   // attendance, obligation, assignment, or previously recognized plan is rewritten.
   revision(p.cycleRef,{...c,currentConfigRef:r,activationState:'ACTIVE'});
   const headRef=M.ref('capacity',c.id),head=load(headRef,true);save(headRef,{...head,state:'INVALIDATED',configRef:r,capacityEpoch:head.capacityEpoch+1});
   return {effectiveConfigRef:r,state:'HISTORICAL_RECOGNITION',preservedFrequencyPlanning:true};
  }
 };
}
function f4CycleParameterApproval_(x,c){
 const P=f4CycleParameters_(),{M,N,p,load}=x;
 if(!P.isManaged(p.parameters))return {};
 N(p.draftRef?.kind==='config'&&p.draftRef.parent===c.id&&c.draftConfigRef?.id===p.draftRef.id,'STALE_REVISION','PARAMETER_DRAFT_REQUIRED');
 const draft=load(p.draftRef,true),policy=load(c.policyVersionRef,true);
 N(draft.state==='DRAFT'&&!draft.approved&&draft.definitionHash===M.hash(p.parameters)&&M.hash(draft.parameters)===draft.definitionHash&&M.ms(draft.effectiveFrom)===p.effectiveFrom,'STALE_REVISION','PARAMETER_DRAFT_CHANGED');
 N(draft.policyRevision===policy.revision&&draft.policyHash===f4CycleParameterDocumentHash_(policy),'STALE_REVISION','NORMATIVE_POLICY_CHANGED');
 P.validate(p.parameters,c,policy,x.get,true);
 return {draftSourceRef:p.draftRef,state:'APPROVED',definitionHash:draft.definitionHash,normativeReference:policy.parameterGovernance.normativeReference};
}
function f4CycleParametersCompat_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),P=f4CycleParameters_(),db=F.connect(options.config),ports=options.ports,clock=options.clock||Date.now;
 function context(id){
  N(M.id(id),'VALIDATION_ERROR');const cycleRef=M.ref('cycle',id),auth=ports.readAuthority([M.path(cycleRef)],false,'PRESENTATION_BOARD_READ','PRESC_'+id),cycle=db.get(M.path(cycleRef));N(cycle,'NOT_FOUND');
  const get=r=>r?.kind&&r.kind!=='NONE'?db.get(M.path(r)):null;
  return {cycleRef,cycle,auth,policy:get(cycle.policyVersionRef),current:get(cycle.currentConfigRef),pending:get(cycle.scheduledConfigRef),draft:get(cycle.draftConfigRef)};
 }
 function recheck(z){ports.recheckAuthority(z.auth);N(db.get(M.path(z.cycleRef))?.revision===z.cycle.revision,'STALE_REVISION','DISPLAYED_CYCLE_CHANGED');}
 function refOptions(){let cursor,rows=[];for(let page=0;page<12;page++){
  const part=db.query({collection:'presentationPolicyReferences',filters:[],order:[{field:'__name__',direction:'ASCENDING'}],limit:20,...(cursor?{cursor:[cursor]}:{})});rows.push(...part.map(r=>r.data));if(part.length<20)return rows;
  cursor=new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+part[part.length-1].path);
 }N(false,'SCOPE_TOO_LARGE','PARAMETER_POLICY_CATALOG_BOUND');}
 function summary(d){return d?{id:d.id,revision:d.revision,state:d.state||(d.approved?'APPROVED':'DRAFT'),effectiveFrom:M.ms(d.effectiveFrom),recordedAt:M.ms(d.recognizedAt||d.createdAt),reason:d.reason||d.approvalDecision?.basis?.reason||'',normativeReference:d.normativeReference||'',predecessorId:d.predecessorRef?.id||'',definitionHash:d.definitionHash||'',parameters:f4CycleEffectiveConfiguration_(d).parameters}:null;}
 function read(input){
  C.exact(input,['cycleId',...('cursor'in input?['cursor']:[])]);if(input.cursor)N(M.id(input.cursor),'VALIDATION_ERROR');const z=context(input.cycleId),config=z.draft||z.pending||z.current;
  const parameters=f4CycleEffectiveConfiguration_(config)?.parameters||{},references=refOptions(),scoped=references.filter(r=>r.approved&&r.active&&(!r.cycleRef||r.cycleRef.id===z.cycle.id));
  const prefix='presentationCycles/'+z.cycle.id+'/configRevisions',cursor=input.cursor?new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+prefix+'/'+input.cursor):null;
  const page=db.query({collection:prefix,filters:[],order:[{field:'__name__',direction:'ASCENDING'}],limit:21,...(cursor?{cursor:[cursor]}:{})});
  const fields=P.catalog.map(f=>({...f,...P.entry(parameters,f.key),revision:config?.id||'',effectiveFrom:config?M.ms(config.effectiveFrom):null,origin:config?.state==='HISTORICAL_RECOGNITION'?'DOCUMENTARY_RECOGNITION':P.isManaged(parameters)?'CYCLE_CONFIGURATION':'LEGACY_CONFIGURATION',rule:z.policy?.parameterGovernance?.rules?.[f.key]||null,options:f.usage?scoped.filter(r=>r.usage===f.usage).map(r=>({id:r.id,label:r.label||r.name||r.code||r.id,revision:r.revision})):[]}));
  const recognition=scoped.filter(r=>r.usage==='CYCLE_PARAMETER_RECOGNITION'&&r.cycleRef?.id===z.cycle.id&&r.expectedCycleRevision===z.cycle.revision&&M.ms(r.issuedAt)<=clock()&&clock()<M.ms(r.expiresAt)).map(r=>({id:r.id,revision:r.revision,documentaryEvidence:r.documentaryEvidence,parameters:r.parameters}));
  const scope='PRESC_'+z.cycle.id,can=cmd=>ports.commandAvailable(cmd,scope);
  recheck(z);return {ok:true,data:{cycle:{id:z.cycle.id,label:z.cycle.cycleLabel,revision:z.cycle.revision,status:z.cycle.status,civilStart:z.cycle.civilStart,civilEnd:z.cycle.civilEnd},fields,parameters,current:summary(z.current),pending:summary(z.pending),draft:summary(z.draft),normativeReference:z.policy?.parameterGovernance?.normativeReference||'',normativeStatus:z.policy?.parameterGovernance?.status||'UNESTABLISHED',frequencyPlanning:z.cycle.frequencyPlanning||null,recognitions:recognition,rights:{canDraft:can('CONFIG_DRAFT_SAVE')&&['DRAFT','OPEN'].includes(z.cycle.status)&&!z.pending,canApprove:can('CONFIG_RECORD')&&!!z.draft,canActivate:can('CONFIG_ACTIVATE')&&!!z.pending&&M.ms(z.pending.effectiveFrom)<=clock(),canRecognize:can('CONFIG_HISTORICAL_RECOGNIZE')&&z.cycle.status==='OPEN'},history:page.slice(0,20).map(r=>summary(r.data)),nextCursor:page.length>20?page[19].data.id:null}};
 }
 function prepare(input){
  C.exact(input,['operationId','route','payload']);N(input.route==='/admin/ciclos-semestres/parametros/salvar','VALIDATION_ERROR');const p=input.payload;
  N(['DRAFT','APPROVE','ACTIVATE','RECOGNIZE'].includes(p.action),'VALIDATION_ERROR');
  const extra={DRAFT:['parameters','effectiveFrom','reason'],APPROVE:['draftId','definitionHash','reason'],ACTIVATE:['configId','configRevision'],RECOGNIZE:['recognitionId','reason']}[p.action];C.exact(p,['action','cycleId','cycleRevision',...extra]);
  const z=context(p.cycleId);N(z.cycle.revision===p.cycleRevision,'STALE_REVISION','DISPLAYED_CYCLE_CHANGED');recheck(z);
  if(p.action==='ACTIVATE')return f4ActivitiesConfigCompat_(options).prepare({operationId:input.operationId,route:'/admin/apresentacoes/configuracao/salvar',payload:{kind:'ACTIVATE',cycleId:p.cycleId,cycleRevision:p.cycleRevision,configId:p.configId,configRevision:p.configRevision}});
  let commandType,payload;
  if(p.action==='DRAFT'){commandType='CONFIG_DRAFT_SAVE';payload={cycleRef:z.cycleRef,parameters:p.parameters,effectiveFrom:p.effectiveFrom,reason:p.reason};}
  if(p.action==='APPROVE'){
   N(z.draft?.id===p.draftId&&z.draft.definitionHash===p.definitionHash,'STALE_REVISION','PARAMETER_DRAFT_CHANGED');
   // CONFIG_RECORD remains the only approval command. Its manifest is sealed
   // by the domain transaction, so no side write is needed during preparation.
   commandType='CONFIG_RECORD';payload={cycleRef:z.cycleRef,parameters:z.draft.parameters,parameterClasses:z.cycle.status==='OPEN'?'OPERATIONAL':'STRUCTURAL',effectiveFrom:M.ms(z.draft.effectiveFrom),predecessorRef:z.cycle.currentConfigRef||{kind:'NONE',id:'INITIAL'},approvalBasis:{reasonCode:'CYCLE_PARAMETERS_APPROVAL',reason:p.reason},draftRef:z.cycle.draftConfigRef,itemManifestRef:z.cycle.draftConfigRef};
  }
  if(p.action==='RECOGNIZE'){const r=M.ref('policy',p.recognitionId),recognition=db.get(M.path(r));N(recognition?.parameters,'UNRESOLVED_DEPENDENCY','RECOGNITION_PARAMETERS_REQUIRED');commandType='CONFIG_HISTORICAL_RECOGNIZE';payload={cycleRef:z.cycleRef,parameters:recognition.parameters,recognitionRef:r,reason:p.reason};}
  const prepared=ports.prepare({operationId:input.operationId,commandType:'ACT.'+commandType,payload});
  N(prepared.target.expected.some(e=>e.kind==='cycle'&&e.id===p.cycleId&&e.revision===p.cycleRevision),'STALE_REVISION');return prepared;
 }
 return {read,prepare};
}
// Preparing a revision needs its selected values and documentary sources, not
// the cycle's entire activity, frequency, or configuration predecessor graph.
function f4CycleParameterDependencies_(command,ref,d){
 if(!['CONFIG_DRAFT_SAVE','CONFIG_HISTORICAL_RECOGNIZE','CONFIG_RECORD','CONFIG_ACTIVATE','CYCLE_THEME_SETTINGS_RECORD','CYCLE_PUBLICATION_SETTINGS_RECORD','CYCLE_MATERIAL_SETTINGS_RECORD'].includes(command))return null;
 const M=f4ActivitiesModel_(),out=[],add=r=>{if(r?.kind&&r.kind!=='NONE')out.push(r);};
 if(ref.kind==='cycle'){for(const r of [d.policyVersionRef,d.currentConfigRef,d.scheduledConfigRef,d.draftConfigRef])add(r);add(M.ref('capacity',ref.id));}
 if(ref.kind==='config'){add(d.policyVersionRef);for(const f of f4CycleParameterCatalog_()){const v=d.parameters?.[f.key];if(f.type==='policy')add(v);if(f.type==='policies')for(const r of v||[])add(r);}}
 if(ref.kind==='policy'){for(const s of d.parameterGovernance?.sources||[])add(s.ref);if(d.usage==='CYCLE_PARAMETER_RECOGNITION')for(const s of d.sources||[])add(s.ref);if(command==='CYCLE_MATERIAL_SETTINGS_RECORD')add(d.executorRuleRef);}
 if(ref.kind==='assessment')for(const b of d.basisRevisionVector||[])add(b.ref);
 return out;
}
