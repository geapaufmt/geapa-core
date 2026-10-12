// Administrative planning on the existing presentations page. The browser
// supplies choices; shared Core authorization and commands own every write.
function f4ActivitiesPlanningCompat_(options) {
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_();
 const db=F.connect(options.config),ports=options.ports,clock=options.clock||Date.now;
 function authorize(cycleId) {
  return ports.readAuthority([cycleId?'presentationCycles/'+cycleId:'presentationCapacityHeads/registry'],false,'PRESENTATION_BOARD_READ',cycleId?'PRESC_'+cycleId:'PRESENTATIONS_PLAN');
 }
 function recheck(proof) {ports.recheckAuthority(proof);}
 function rows(collection,filters=[]) {
  let cursor;const result=[];
  for(let page=0;page<20;page++) {
   const values=db.query({collection,filters,order:[{field:'__name__',direction:'ASCENDING'}],limit:20,...(cursor?{cursor:[cursor]}:{})});
   result.push(...values.map(v=>v.data));if(values.length<20)return result;
   cursor=new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+collection+'/'+values[values.length-1].data.id);
  }
  N(false,'SCOPE_TOO_LARGE','PLANNING_LIST_BOUND');
 }
 function read(input) {
  C.exact(input,'cycleId'in input?['cycleId']:[]);if(input.cycleId)N(M.id(input.cycleId),'VALIDATION_ERROR');
  const proof=authorize(input.cycleId),proofs=[proof];let data;
  if(!input.cycleId) {
   const cycles=[];
   for(const c of rows('presentationCycles'))try {
    proofs.push(authorize(c.id));cycles.push({id:c.id,revision:c.revision,label:c.cycleLabel,start:c.civilStart,end:c.civilEnd,status:c.status});
   }catch(e){if(!['AUTHORIZATION_DENIED','AUTH_BINDING_INVALID','ACCESS_REVOKED','M13_AUTHZ_DENIED'].includes(e.code))throw e;}
   const policies=rows('presentationPolicyReferences').filter(p=>p.approved&&p.active&&p.usage==='CYCLE').map(p=>({id:p.id,label:p.label||p.id}));
   const policy=db.get('memberAuthzPolicies/'+proof.q.sourceId);data={cycles,policies,canCreateCycle:proof.q.kind==='DEV_ADMIN'||policy?.targets.resourceScope?.cycleAccess==='ALL'&&policy.capabilities.includes('PRESENTATION_PLAN_MANAGE')};
  }else{
   const c=db.get('presentationCycles/'+input.cycleId);N(c,'NOT_FOUND');
   const cfg=c.currentConfigRef&&db.get(M.path(c.currentConfigRef)),sessions=rows('presentationSessions',[{field:'cycleRef',op:'EQUAL',value:M.ref('cycle',c.id)}]);
   proofs.push(ports.readAuthority(['activityAttendanceRegistry/current'],false));
   const activities=rows('canonicalActivities').filter(a=>a.status==='SCHEDULED'&&M.gap(c.civilStart,a.civilDate)>=0&&M.gap(a.civilDate,c.civilEnd)>=0).map(a=>({id:a.id,revision:a.revision,title:a.title,date:a.civilDate,sessionId:a.presentationSessionRef?.id||''}));
   data={incidenceConfigured:db.get(M.path(c.policyVersionRef))?.incidenceRule==='ORDINARY_FIRST_EFFECTIVATION',configuration:f4ActivitiesConfigCompat_(options).read(c.id),cycle:{id:c.id,revision:c.revision,label:c.cycleLabel,status:c.status,start:c.civilStart,end:c.civilEnd},config:cfg?{revision:cfg.revision,parameters:cfg.parameters}:null,sessions:sessions.map(s=>({id:s.id,revision:s.revision,date:s.civilDate,capacity:s.config.capacity,destination:s.config.destination,activityId:s.activityRef?.id||'',occupied:s.slots.length})),activities};
   const capacity=db.get('presentationCapacityHeads/'+c.id),coverage=capacity?.incidenceCoverageRef&&db.get(M.path(capacity.incidenceCoverageRef));
   data.canOpen=!!(c.status==='DRAFT'&&cfg?.approved&&data.incidenceConfigured&&coverage&&ports.commandAvailable('CYCLE_OPEN','PRESC_'+c.id));
   data.openingUnavailableReason=data.canOpen?'':!data.incidenceConfigured?'A abertura aguarda a configuração da regra institucional de incidência.':!coverage?'Confira a população e as obrigações do ciclo antes da abertura.':'A abertura não está disponível neste estado ou para sua função.';
   N(db.get('presentationCycles/'+c.id)?.revision===c.revision,'STALE_REVISION');
  }
  proofs.forEach(recheck);return {ok:true,data};
 }
 function prepare(input) {
  C.exact(input,['operationId','route','payload']);N(input.route==='/admin/apresentacoes/planejamento/salvar','VALIDATION_ERROR');
  const p=input.payload;N(['CYCLE','SESSION','OPEN'].includes(p.kind),'VALIDATION_ERROR');
  N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'INCOMPLETE_REQUEST');
  const decision={reasonCode:'PORTAL_PLANNING_DECISION',reason:p.reason.trim()};let command,shown=[];
  if(p.kind==='CYCLE') {
   const proof=authorize(),source=db.get('memberAuthzPolicies/'+proof.q.sourceId);N(proof.q.kind==='DEV_ADMIN'||source?.targets.resourceScope?.cycleAccess==='ALL','AUTHORIZATION_DENIED','FUTURE_CYCLE_SCOPE_REQUIRED');recheck(proof);
   C.exact(p,['kind','label','start','end','policyId','reason']);N(typeof p.label==='string'&&p.label.trim()&&p.label.length<=120&&M.id(p.policyId),'VALIDATION_ERROR');M.civil(p.start);M.civil(p.end);
   const policy=db.get('presentationPolicyReferences/'+p.policyId);N(policy?.approved&&policy.active&&policy.usage==='CYCLE','CONFIG_CAPABILITY_DISABLED','CYCLE_POLICY_REQUIRED');
   command={commandType:'ACT.CYCLE_CREATE',payload:{cycleLabel:p.label.trim(),civilStart:p.start,civilEnd:p.end,timezone:'America/Cuiaba',policyRef:M.ref('policy',p.policyId),decisionBasis:decision}};
  }else if(p.kind==='OPEN'){
   C.exact(p,['kind','cycleId','cycleRevision','reason']);N(M.id(p.cycleId),'VALIDATION_ERROR');const cycleRef=M.ref('cycle',p.cycleId),cycle=db.get(M.path(cycleRef));N(cycle?.revision===p.cycleRevision&&cycle.status==='DRAFT'&&cycle.currentConfigRef,'STALE_REVISION','DISPLAYED_CYCLE_CHANGED');
   const head=db.get('presentationCapacityHeads/'+cycle.id),placeholder=M.ref('assessment','opening_'+input.operationId);N(head?.incidenceCoverageRef,'UNRESOLVED_DEPENDENCY','INCIDENCE_COVERAGE_REQUIRED');
   const intent={cycleRef,approvedConfigRef:cycle.currentConfigRef,capacityAssessmentRef:placeholder,openingDecision:decision},authorized=ports.prepare({operationId:input.operationId,commandType:'ACT.CYCLE_OPEN',payload:intent});ports.authorizePrepared(path=>db.get(path),authorized);
   const jobs=options.jobs||[],index=jobs.findIndex(j=>j.kind==='CAPACITY_SYNC'&&j.cycleId===cycle.id);N(index>=0&&options.pipeline,'CONFIG_CAPABILITY_DISABLED','CAPACITY_REFRESH_NOT_CONFIGURED');
   f4ActivitiesConfiguredJobs_({config:options.config,db,pipeline:options.pipeline,jobs,clock}).run(index);
   const fresh=db.get('presentationCapacityHeads/'+cycle.id);N(fresh?.state==='SEALED'&&db.get(M.path(cycleRef))?.revision===p.cycleRevision,'STALE_REVISION');
   command={commandType:'ACT.CYCLE_OPEN',payload:{...intent,capacityAssessmentRef:fresh.currentAssessmentRef}};shown=[[cycleRef,p.cycleRevision]];
  }else{
   C.exact(p,['kind','cycleId','cycleRevision','sessionId','sessionRevision','activityId','activityRevision','capacity','destination','reason']);
   N(M.id(p.cycleId)&&M.id(p.activityId)&&(!p.sessionId||M.id(p.sessionId))&&[1,2].includes(p.capacity)&&['ORDINARY','EXCLUSIVE','RESERVE','BLOCKED','CANCELLED'].includes(p.destination),'VALIDATION_ERROR');
   const cycleRef=M.ref('cycle',p.cycleId),activityRef=M.ref('activity',p.activityId),sessionRef=M.ref('session',p.sessionId||'ses_'+input.operationId);
   shown=[[cycleRef,p.cycleRevision],[activityRef,p.activityRevision],[sessionRef,p.sessionRevision]];
   for(const [ref,revision]of shown){const value=db.get(M.path(ref));N(Number.isSafeInteger(revision)&&(value?.revision||0)===revision,'STALE_REVISION','DISPLAYED_PLANNING_CHANGED');}
   const activity=db.get(M.path(activityRef));N(activity,'NOT_FOUND');
   command={commandType:'ACT.SESSION_CONFIGURE',payload:{cycleRef,sessionRef,activityRef,civilDate:activity.civilDate,capacity1or2:p.capacity,destination:p.destination,decision,...(['EXCLUSIVE','RESERVE'].includes(p.destination)?{exclusiveBasis:decision}:{})}};
  }
  if(p.kind==='SESSION'){
   const cycle=db.get(M.path(command.payload.cycleRef)),session=db.get(M.path(command.payload.sessionRef));
   if(cycle.status==='OPEN'&&session&&(p.capacity<session.config.capacity||p.destination!==session.config.destination)){
    const placeholder=M.ref('assessment',M.hash({operationId:input.operationId,kind:'SESSION_CAPACITY_AUTH'}).slice(0,32));
    const authorized=ports.prepare({operationId:input.operationId,...command,payload:{...command.payload,assessmentRef:placeholder}});
    const authorize=get=>{ports.authorizePrepared(get,authorized);for(const [ref,rev]of shown)N((get(M.path(ref))?.revision||0)===rev,'STALE_REVISION','DISPLAYED_PLANNING_CHANGED');return true;};
    const preparation=f4ActivitiesPreparation_({db,clock,capacitySource:f4ActivitiesCapacitySource_({db,clock,config:options.config}),authorize,executionGate:()=>authorize(path=>db.get(path))});
    command.payload.assessmentRef=preparation.prepareCapacity(null,command.payload.cycleRef,cycle.currentConfigRef,{sessionConfig:{sessionRef:command.payload.sessionRef,capacity:p.capacity,destination:p.destination}});
   }
  }
  const prepared=ports.prepare({operationId:input.operationId,...command});
  for(const [ref,revision]of shown)N(prepared.target.expected.some(v=>v.kind===ref.kind&&v.id===ref.id&&(revision?v.revision===revision:v.exists===false)),'STALE_REVISION');
  return prepared;
 }
 return {read,prepare};
}
