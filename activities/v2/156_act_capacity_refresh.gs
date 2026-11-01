// One configured Core job refreshes the technical capacity proof. It cannot
// decide incidence, open a cycle, allocate a date or create institutional rules.
function f4ActivitiesCapacityRefresh_(options){
 const M=f4ActivitiesModel_(),N=M.need,db=options.db,clock=options.clock||Date.now;
 function run(entry){
  const cycleRef=M.ref('cycle',entry.cycleId),headRef=M.ref('capacity',entry.cycleId),scope='PRESC_'+entry.cycleId;
  N(entry.kind==='CAPACITY_SYNC'&&entry.ruleRef?.kind==='policy','VALIDATION_ERROR');
  function gate(get=p=>db.get(p)){
   const rule=get(M.path(entry.ruleRef));N(rule?.approved===true&&rule.active===true&&rule.cycleRef?.id===entry.cycleId&&rule.commands?.includes('ACT.CAPACITY_ASSESSMENT_SEAL'),'CONFIG_CAPABILITY_DISABLED','CAPACITY_JOB_NOT_APPROVED');return rule;
  }
  const rule=gate(),cycle=db.get(M.path(cycleRef)),head=db.get(M.path(headRef)),registry=db.get('presentationCapacityHeads/registry');
  N(cycle?.currentConfigRef&&head&&registry,'UNRESOLVED_DEPENDENCY','CAPACITY_CONFIGURATION');
  if(head.state==='SEALED'&&head.synchronizedMembersEpoch===registry.membersEpoch&&head.synchronizedCycleRegistryEpoch===registry.cycleRegistryEpoch&&M.ms(head.validUntil)>clock()+5000)return {state:'FRESH',results:[],nextCursor:null};
  const handle={},principal={kind:'SYSTEM_JOB',servicePrincipal:'CORE_ACT_DEV_JOB',scope,ruleRef:entry.ruleRef,jobKind:'CAPACITY_SYNC',runId:M.hash({cycle:entry.cycleId,head:head.revision,at:clock()}).slice(0,32)},principals={resolve:h=>h===handle?principal:undefined};
  const preparation=f4ActivitiesPreparation_({db,clock,principals,capacitySource:f4ActivitiesCapacitySource_({db,clock,config:options.config}),executionGate:()=>gate()});
  const sourceRef=preparation.prepareCapacity(handle,cycleRef,cycle.currentConfigRef,{}),source=db.get(M.path(sourceRef));
  const refs=new Map([[M.path(sourceRef),sourceRef],[M.path(entry.ruleRef),entry.ruleRef],['presentationCapacityHeads/registry',M.ref('capacity','registry')]]);
  for(const v of source.basisRevisionVector)refs.set(M.path(v.ref),v.ref);
  const expected=[...refs].map(([path,ref])=>({...ref,revision:db.get(path).revision}));
  const q={schemaVersion:'geapa-command-v2-candidate',operationId:M.hash({kind:'CAPACITY_SYNC',source:sourceRef.id,ruleRevision:rule.revision}).slice(0,32),commandType:'ACT.CAPACITY_ASSESSMENT_SEAL',payload:{cycleRef,sourceManifestRef:sourceRef,proposedChangeDigest:source.proposedChangeDigest,evaluationRuleRef:entry.ruleRef},expected,scope:{targetRefs:[cycleRef,headRef]}};
  const result=f4ActivitiesSharedSystemCommands_(options.pipeline,principals).execute(handle,q);
  return {state:'RAN',kind:'CAPACITY_SYNC',cycleId:entry.cycleId,results:[result],nextCursor:null};
 }
 return {run};
}
