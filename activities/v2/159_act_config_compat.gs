// Explicit cycle settings. Prepared manifests are Core-owned proofs; they do
// not grant authority and are revalidated by the shared command transaction.
function f4ActivitiesConfigCompat_(options) {
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),db=f4Firestore_().connect(options.config);
 const ports=options.ports,clock=options.clock||Date.now;
 const fields=['calendarMode','shared','quota','early','dynamic','dynamicLimit','newDateMinimumDays','thematicMinimumDays','titleMinimumDays','selfSchedulingOpensAt','selfSchedulingEndsAt'];
 function context(id) {
  N(M.id(id),'VALIDATION_ERROR');const cycleRef=M.ref('cycle',id),cycle=db.get(M.path(cycleRef));N(cycle,'NOT_FOUND');
  const proof=ports.readAuthority([M.path(cycleRef)],false,'PRESENTATION_BOARD_READ','PRESC_'+id);
  const policy=db.get(M.path(cycle.policyVersionRef));N(policy?.approved,'CONFIG_CAPABILITY_DISABLED','CYCLE_POLICY_REQUIRED');
  const current=cycle.currentConfigRef&&db.get(M.path(cycle.currentConfigRef));
  const pending=cycle.scheduledConfigRef?.kind==='config'&&db.get(M.path(cycle.scheduledConfigRef));
  return {cycleRef,cycle,policy,current,pending,proof};
 }
 function read(id) {
  const z=context(id),parameters=z.pending?.parameters||z.current?.parameters||z.policy.initialParameters;
  f4SharedAuthz_().recheck(p=>db.get(p),z.proof.q,z.proof.proof,clock());
  return {parameters:parameters||null,pending:z.pending?{id:z.pending.id,revision:z.pending.revision,effectiveFrom:M.ms(z.pending.effectiveFrom),due:M.ms(z.pending.effectiveFrom)<=clock()}:null};
 }
 function seal(request,ref,document,basis) {
  ports.authorizePrepared(p=>db.get(p),request);
  db.transaction(tx=>{
   const get=p=>tx.get(p).data();ports.authorizePrepared(get,request);
   for(const v of basis)N(get(M.path(v.ref))?.revision===v.revision,'STALE_REVISION','CONFIG_PROOF_SOURCE_CHANGED');
   const old=get(M.path(ref));if(old)N(old.digest===document.digest&&old.proofDigest===document.proofDigest,'CONFLICT','CONFIG_PROOF_REUSE');else tx.set(M.path(ref),document);
  });
 }
 function manifest(operationId,z,kind,items,basis,extra={}) {
  N(items.length<=32,'SCOPE_TOO_LARGE');
  const proofDigest=M.hash({operationId,kind,items,basis,extra}),ref=M.ref('assessment',M.hash({operationId,kind}).slice(0,32));
  return {ref,document:{id:ref.id,revision:1,schemaVersion:'act-config-proof-v1',cycleRef:z.cycleRef,state:'SEALED',complete:true,items,digest:M.hash(items),proofDigest,basisRevisionVector:basis,recordedAt:M.stamp(clock()),validUntil:M.stamp(clock()+60000),...extra}};
 }
 function prepare(input) {
  C.exact(input,['operationId','route','payload']);N(input.route==='/admin/apresentacoes/configuracao/salvar','VALIDATION_ERROR');
  const p=input.payload;N(['RECORD','ACTIVATE'].includes(p.kind),'VALIDATION_ERROR');
  C.exact(p,['kind','cycleId','cycleRevision',...(p.kind==='RECORD'?['values','effectiveFrom','reason']:['configId','configRevision'])]);
  const z=context(p.cycleId);N(z.cycle.revision===p.cycleRevision,'STALE_REVISION','DISPLAYED_CYCLE_CHANGED');
  let command,proof;
  if(p.kind==='RECORD') {
   C.exact(p.values,[...fields,...['reorganizationOpensAt','reorganizationEndsAt'].filter(k=>k in p.values)]);N(('reorganizationOpensAt'in p.values)===('reorganizationEndsAt'in p.values),'VALIDATION_ERROR');N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000&&Number.isSafeInteger(p.effectiveFrom)&&p.effectiveFrom>=clock(),'VALIDATION_ERROR');
   const before=z.pending?.parameters||z.current?.parameters||z.policy.initialParameters;N(before,'CONFIG_CAPABILITY_DISABLED','INITIAL_PARAMETERS_REQUIRED');
   const parameters={...before};for(const key of fields)if(!key.startsWith('selfScheduling'))parameters[key]=p.values[key];
   N(['BOARD_INITIAL','MEMBER_SELF_SCHEDULE'].includes(parameters.calendarMode),'VALIDATION_ERROR');
   for(const key of ['shared','quota','early','dynamic'])N(typeof parameters[key]==='boolean','VALIDATION_ERROR');
   for(const [key,min]of [['dynamicLimit',0],['newDateMinimumDays',30],['thematicMinimumDays',20],['titleMinimumDays',10]])N(Number.isSafeInteger(parameters[key])&&parameters[key]>=min,'VALIDATION_ERROR');
   if(p.values.selfSchedulingOpensAt||p.values.selfSchedulingEndsAt)parameters.selfSchedulingWindow={opensAt:p.values.selfSchedulingOpensAt,endsAtExclusive:p.values.selfSchedulingEndsAt};
   else delete parameters.selfSchedulingWindow;
   if('reorganizationOpensAt'in p.values){if(p.values.reorganizationOpensAt||p.values.reorganizationEndsAt){N(Number.isSafeInteger(p.values.reorganizationOpensAt)&&Number.isSafeInteger(p.values.reorganizationEndsAt)&&p.values.reorganizationOpensAt<p.values.reorganizationEndsAt,'VALIDATION_ERROR');parameters.reorganizationWindow={opensAt:p.values.reorganizationOpensAt,endsAtExclusive:p.values.reorganizationEndsAt};}else delete parameters.reorganizationWindow;}
   N(parameters.calendarMode!=='MEMBER_SELF_SCHEDULE'||f4ActivitiesSelfSchedulingWindow_(parameters,clock()).configured,'CONFIG_CAPABILITY_DISABLED','SELF_SCHEDULING_WINDOW_REQUIRED');
   const changed=f4ActivitiesConfigChangedKeys_(before,parameters);
   const operational=changed.every(key=>['selfSchedulingWindow','reorganizationWindow'].includes(key)||z.policy.parameterClasses?.[key]==='OPERATIONAL'),parameterClasses=operational?'OPERATIONAL':'STRUCTURAL';
   f4ActivitiesConfigurationChange_(before,parameters,parameterClasses,z.cycle,z.policy,clock());
   N(z.cycle.status!=='OPEN'||parameterClasses==='OPERATIONAL','HUMAN_DECISION_REQUIRED','STRUCTURAL_FREEZE');
   const predecessor=z.pending?z.cycle.scheduledConfigRef:z.cycle.currentConfigRef,basis=[{ref:z.cycleRef,revision:z.cycle.revision},{ref:z.cycle.policyVersionRef,revision:z.policy.revision},...(predecessor?[{ref:predecessor,revision:(z.pending||z.current).revision}]:[])];
   proof=manifest(input.operationId,z,'CONFIG_RECORD',Object.entries(parameters).map(([key,value])=>({key,value})),basis);
   command={commandType:'ACT.CONFIG_RECORD',payload:{cycleRef:z.cycleRef,parameters,parameterClasses,effectiveFrom:p.effectiveFrom,predecessorRef:predecessor||{kind:'NONE',id:'INITIAL'},approvalBasis:{reasonCode:'PORTAL_CYCLE_CONFIGURATION',reason:p.reason.trim()},itemManifestRef:proof.ref}};
  }else{
   N(z.pending?.id===p.configId&&z.pending.revision===p.configRevision&&M.ms(z.pending.effectiveFrom)<=clock(),'STALE_REVISION','PENDING_CONFIG_CHANGED');
   const before=z.current?.parameters||{},after=z.pending.parameters,changed=f4ActivitiesConfigurationChange_(before,after,z.pending.parameterClasses,z.cycle,z.policy,clock());
   // Extending a granted window preserves its existing rights. An initial
   // configuration is safe only while there are no assignments or obligations.
   const extension=changed.every(k=>['selfSchedulingWindow','reorganizationWindow'].includes(k));
   if(!extension){N(z.cycle.status==='DRAFT','HUMAN_DECISION_REQUIRED','CONFIG_IMPACT_ASSESSMENT_REQUIRED');for(const collection of ['calendarAssignments','presentationObligations'])N(db.query({collection,filters:[{field:'cycleRef',op:'EQUAL',value:z.cycleRef}],order:[{field:'__name__',direction:'ASCENDING'}],limit:1}).length===0,'HUMAN_DECISION_REQUIRED','CONFIG_EXISTING_RIGHTS');}
   const registry=db.get('presentationCapacityHeads/registry'),head=db.get('presentationCapacityHeads/'+z.cycle.id);N(registry&&head,'UNRESOLVED_DEPENDENCY');
   const basis=[{ref:z.cycleRef,revision:z.cycle.revision},{ref:z.cycle.scheduledConfigRef,revision:z.pending.revision},{ref:M.ref('capacity','registry'),revision:registry.revision},{ref:M.ref('capacity',z.cycle.id),revision:head.revision},...(z.current?[{ref:z.cycle.currentConfigRef,revision:z.current.revision}]:[])];
   proof=manifest(input.operationId,z,'CONFIG_ACTIVATE',basis,basis,{membersEpoch:registry.membersEpoch,cycleRegistryEpoch:registry.cycleRegistryEpoch,preservesAcquiredRights:true});
   command={commandType:'ACT.CONFIG_ACTIVATE',payload:{cycleRef:z.cycleRef,approvedConfigRef:z.cycle.scheduledConfigRef,impactAssessmentRef:proof.ref}};
  }
  const request=ports.prepare({operationId:input.operationId,...command});
  seal(request,proof.ref,proof.document,proof.document.basisRevisionVector);
  return ports.prepare({operationId:input.operationId,...command});
 }
 return {read,prepare};
}
