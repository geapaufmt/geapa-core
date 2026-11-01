// A human assessment applies only to its exact pair, content and context.
// Content includes canonical timestamps after a real submission. Encode those
// values before hashing, using the same wire representation on both executors.
function f4ActivitiesCompatibilityDigest_(basis){
 const F=f4Firestore_();
 function value(v){
  if(v instanceof F.Timestamp)return {firestoreTimestamp:F.encode(v).timestampValue};
  if(Array.isArray(v))return v.map(value);
  if(v&&typeof v==='object'){f4Core_().need(Object.getPrototypeOf(v)===Object.prototype,'VALIDATION_ERROR');return Object.fromEntries(Object.entries(v).map(([k,x])=>[k,value(x)]));}
  return v;
 }
 return f4ActivitiesModel_().hash(value(basis));
}
function f4ActivitiesCompatibilityBasis_(session,get){
 const M=f4ActivitiesModel_(),N=M.need;
 N(session&&session.slots.length===2&&new Set(session.slots.map(v=>v.unitRef.id)).size===2,'INVALID_STATE_TRANSITION','CONCRETE_PAIR_REQUIRED');
 const cycle=get(M.path(session.cycleRef)),config=cycle?.currentConfigRef&&get(M.path(cycle.currentConfigRef));
 N(cycle&&config?.approved&&!config.parameters.reviewRequired&&config.parameters.shared===true,'CONFIG_CAPABILITY_DISABLED','SHARED_OFF');
 const contextRef=config.parameters.compatibilityGuideRef||cycle.policyVersionRef,context=contextRef&&get(M.path(contextRef));
 N(context?.approved&&Number.isSafeInteger(context.revision),'CONFIG_CAPABILITY_DISABLED','COMPATIBILITY_CONTEXT_REQUIRED');
 const units=session.slots.map(slot=>{const u=get(M.path(slot.unitRef)),a=get(M.path(slot.assignmentRef)),t=u?.thematicRequestRef&&get(M.path(u.thematicRequestRef));
  N(u?.cycleRef.id===cycle.id&&a?.unitRef.id===u.id&&u.assignmentRef?.id===a.id&&a.sessionRef.id===session.id&&['ACTIVE','CONFIRMED'].includes(a.status),'UNRESOLVED_DEPENDENCY','ASSIGNMENT_SLOT_DRIFT');
  return {unitRef:slot.unitRef,assignmentRef:slot.assignmentRef,personRef:u.personRef,thematicRef:u.thematicRequestRef||null,thematicRevision:t?.submittedRevisionRef||null,sourceRevision:t?.revision||0,content:t?.candidate||null,state:t?.workflowState||'MISSING'};
 }).sort((a,b)=>a.unitRef.id.localeCompare(b.unitRef.id));
 N(units[0].personRef.id!==units[1].personRef.id,'UNRESOLVED_DEPENDENCY','DISTINCT_PARTICIPANTS_REQUIRED');
 return {sessionRef:M.ref('session',session.id),cycleRef:session.cycleRef,compositionRevision:session.compositionRevision,sessionConfigRevision:session.configRevision,configRef:cycle.currentConfigRef,configRevision:config.revision,contextRef,contextRevision:context.revision,units};
}
function f4ActivitiesCompatibilityCurrent_(assessment,session,get){
 if(!assessment?.exactBasis||session?.compatibilityRef?.id!==assessment.id)return false;
 try{return f4ActivitiesCompatibilityDigest_(assessment.exactBasis)===f4ActivitiesCompatibilityDigest_(f4ActivitiesCompatibilityBasis_(session,get));}
 catch(e){if(['INVALID_STATE_TRANSITION','CONFIG_CAPABILITY_DISABLED','UNRESOLVED_DEPENDENCY'].includes(e.code))return false;throw e;}
}
function f4ActivitiesRecordCompatibility_(x){
 const {M,N,p,load,event,revision}=x,s=load(p.sessionRef,true);x.config();const basis=f4ActivitiesCompatibilityBasis_(s,x.get);
 N(p.basisDigest===f4ActivitiesCompatibilityDigest_(basis),'STALE_REVISION','DISPLAYED_COMPATIBILITY_CHANGED');
 const conclusion=['COMPATIBLE','INCOMPATIBLE','REORGANIZATION_REQUIRED'].includes(p.outcome);
 N(conclusion||['IN_ANALYSIS','WAITING_INFORMATION','ADJUSTMENT_REQUESTED'].includes(p.outcome),'VALIDATION_ERROR');
 N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'HUMAN_DECISION_REQUIRED','COMPATIBILITY_REASON_REQUIRED');
 N(Array.isArray(p.twoUnitRefs)&&p.twoUnitRefs.length===2&&new Set(p.twoUnitRefs.map(r=>r.id)).size===2&&Array.isArray(p.thematicRevisionRefs)&&p.thematicRevisionRefs.length===2,'VALIDATION_ERROR','EXACT_PAIR_REQUIRED');
 for(let i=0;i<2;i++){const b=basis.units.find(b=>b.unitRef.id===p.twoUnitRefs[i].id),r=p.thematicRevisionRefs[i];N(b&&M.hash(b.thematicRevision||{kind:'NONE'})===M.hash(r),'STALE_REVISION','THEMATIC_BASIS_CHANGED');if(conclusion)N(b.content?.titleOrObject?.trim()&&b.content.summary?.trim()&&!['MISSING','WITHDRAWN','REJECTED','WAITING'].includes(b.state),'INCOMPLETE_REQUEST','THEMATIC_INFORMATION_REQUIRED');load(b.unitRef,true);if(b.thematicRef)load(b.thematicRef,true);}
 if(p.equivalencePriorRef)N(s.compatibilityRef&&M.path(p.equivalencePriorRef)===M.path(s.compatibilityRef),'STALE_REVISION','PREVIOUS_ASSESSMENT_CHANGED');
 const basisRevisionVector=basis.units.filter(b=>b.thematicRef).map(b=>({ref:b.thematicRef,revision:b.sourceRevision})),r=event('compatibility',{cycleRef:s.cycleRef,sessionRef:p.sessionRef,compositionRevision:s.compositionRevision,sessionConfigRevision:s.configRevision,units:basis.units.map(b=>b.unitRef),thematicRevisionRefs:basis.units.map(b=>b.thematicRevision),basisRevisionVector,exactBasis:basis,result:p.outcome,reason:p.reason.trim(),...(s.compatibilityRef?{predecessorRef:s.compatibilityRef}:{}),...(p.equivalencePriorRef?{equivalencePriorRef:p.equivalencePriorRef}:{})});
 for(const b of basis.units){const u=load(b.unitRef),ref=u.currentCompatibilityRequestRef,f=ref&&load(ref);if(f?.state==='OPEN'&&f.sessionRef.id===s.id&&f.compositionRevision===s.compositionRevision)revision(ref,{...f,state:'ANSWERED',assessmentRef:r});}
 revision(p.sessionRef,{...s,compatibilityRef:r});return {assessmentRef:r};
}
function f4ActivitiesCompatibilityRequest_(x){
 const {M,N,p,load,event,revision}=x,u=x.own(p.unitRef),a=u.assignmentRef&&load(u.assignmentRef,true),s=a&&load(a.sessionRef,true);
 N(x.config().parameters.shared&&s?.slots.length===2&&s.slots.some(v=>v.unitRef.id===u.id),'CONFIG_CAPABILITY_DISABLED','CONCRETE_PAIR_REQUIRED');
 N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'INCOMPLETE_REQUEST');
 const prior=u.currentCompatibilityRequestRef&&load(u.currentCompatibilityRequestRef);if(prior?.state==='OPEN'&&prior.sessionRef.id===s.id&&prior.compositionRevision===s.compositionRevision&&prior.reason===p.reason.trim())return {requestRef:u.currentCompatibilityRequestRef,replay:true};
 const requestRef=event('flow',{kind:'COMPATIBILITY_REQUEST',state:'OPEN',unitRef:p.unitRef,personRef:u.personRef,cycleRef:u.cycleRef,sessionRef:a.sessionRef,compositionRevision:s.compositionRevision,reason:p.reason.trim(),...(u.currentCompatibilityRequestRef?{predecessorRef:u.currentCompatibilityRequestRef}:{}),submittedAt:M.stamp(x.now)});
 revision(p.unitRef,{...u,currentCompatibilityRequestRef:requestRef});return {requestRef};
}
