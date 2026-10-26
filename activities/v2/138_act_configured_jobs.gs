// Explicit DEV jobs reuse the shared Core scheduler and command pipeline.
// No trigger is installed here; an absent configuration performs no discovery.
function f4ActivitiesConfiguredJobs_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),clock=options.clock||Date.now,db=options.db;
 const kinds={CAPACITY_SYNC:{},THEMATIC_TACIT:{query:'Q22',collection:'presentationThematicRequests',ref:'thematic',state:'clockState',value:'RUNNING',due:'nextEvaluationAt'},CORRECTION_EXPIRE:{query:'Q13S',collection:'presentationCorrectionOpportunities',ref:'opportunity',state:'status',value:'OPEN',due:'dueAt',exclusive:true}};
 function validate(entries){N(Array.isArray(entries)&&entries.length<=8,'VALIDATION_ERROR');const seen=new Set();for(const e of entries){C.exact(e,['kind','cycleId','ruleRef']);N(kinds[e.kind]&&M.id(e.cycleId)&&e.ruleRef?.kind==='policy','VALIDATION_ERROR');M.path(e.ruleRef);const key=e.kind+':'+e.cycleId;N(!seen.has(key),'VALIDATION_ERROR');seen.add(key);}return entries;}
 function run(entry,cursor){
  const configured=validate(options.jobs||[]);if(!configured.length)return {state:'DISABLED',results:[],nextCursor:null};
  N(Number.isInteger(entry)&&entry>=0&&entry<configured.length,'VALIDATION_ERROR');const selected=configured[entry],kind=selected.kind,spec=kinds[kind],scope='PRESC_'+selected.cycleId,cycleRef=M.ref('cycle',selected.cycleId),rule=db.get(M.path(selected.ruleRef));
  N(rule?.approved&&rule.active&&rule.cycleRef?.id===selected.cycleId,'CONFIG_CAPABILITY_DISABLED','JOB_NOT_APPROVED_FOR_CYCLE');
  if(kind==='CAPACITY_SYNC'){N(!cursor,'VALIDATION_ERROR');return f4ActivitiesCapacityRefresh_(options).run(selected);}
  const handles=new WeakMap(),principals={resolve:h=>handles.get(h)},core=f4ActivitiesSharedSystemCommands_(options.pipeline,principals);
  const services={db,core,refreshQuota:(ref,ruleRef,revision)=>f4ActivitiesQuotaQueue_({db,config:options.config,clock}).refresh(ref,get=>{const live=get(M.path(ruleRef)),bucket=get(M.path(ref));N(kind==='THEMATIC_TACIT'&&M.path(ruleRef)===M.path(selected.ruleRef)&&live?.revision===revision&&live.approved&&live.active&&live.commands?.includes('ACT.THEME_TACIT_APPLY')&&live.cycleRef?.id===selected.cycleId&&bucket?.cycleRef.id===selected.cycleId,'AUTHORIZATION_DENIED');return true;}),ruleFor:(requested,cycle)=>{N(requested===kind&&cycle.id===selected.cycleId,'AUTHORIZATION_DENIED');return selected.ruleRef;},principal:(requested,b)=>{N(requested===kind&&M.path(b.ruleRef)===M.path(selected.ruleRef),'AUTHORIZATION_DENIED');const h={};handles.set(h,{kind:'SYSTEM_JOB',servicePrincipal:'CORE_ACT_DEV_JOB',scope,ruleRef:selected.ruleRef,jobKind:kind,runId:C.hash({kind,cycle:selected.cycleId,at:clock()}).slice(0,32)});return h;},discover:(query,target,after,limit)=>{
   N(query===spec.query&&target===scope&&limit===20,'AUTHORIZATION_DENIED');let start;
   if(after){C.exact(after,['kind','cycleId','id','dueAt']);N(after.kind===kind&&after.cycleId===selected.cycleId&&M.id(after.id)&&Number.isSafeInteger(after.dueAt),'VALIDATION_ERROR');start=[M.stamp(after.dueAt),new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+spec.collection+'/'+after.id)];}
   const rows=db.query({collection:spec.collection,filters:[{field:'cycleRef',op:'EQUAL',value:cycleRef},{field:spec.state,op:'EQUAL',value:spec.value},{field:spec.due,op:spec.exclusive?'LESS_THAN':'LESS_THAN_OR_EQUAL',value:M.stamp(clock())}],order:[{field:spec.due,direction:'ASCENDING'},{field:'__name__',direction:'ASCENDING'}],limit,...(start?{cursor:start}:{})});
   const items=rows.map(({data:d})=>{N(d.cycleRef?.id===selected.cycleId&&d[spec.state]===spec.value&&(spec.exclusive?M.ms(d[spec.due])<clock():M.ms(d[spec.due])<=clock()),'UNRESOLVED_DEPENDENCY','JOB_QUERY_SCOPE');return {...d,ref:M.ref(spec.ref,d.id),dueAt:d[spec.due]}}),last=items[items.length-1];return {items,nextCursor:items.length===limit?{kind,cycleId:selected.cycleId,id:last.id,dueAt:M.ms(last.dueAt)}:null};
  }};
  return {state:'RAN',kind,cycleId:selected.cycleId,...f4ActivitiesScheduler_(services).run(kind,scope,cursor)};
 }
 return {validate,run};
}
