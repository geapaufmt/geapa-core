// Recognition of an existing imported schedule. No incidence, obligation,
// presentation, session, theme approval or institutional permission is created.
function f4ActivitiesPreagenda_(){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),eq=(a,b)=>!!a&&!!b&&M.path(a)===M.path(b);
 const mode=p=>(p?.planningMode||'INCIDENCE')==='RECOGNIZED_PREAGENDA';
 const refs=(activityRef,presentationRef)=>{const id=M.hash({kind:'HISTORICAL_PREAGENDA',activityRef,presentationRef}).slice(0,32);return {unitRef:M.ref('unit',id),basisRef:M.ref('flow',id)};};
 function policy(parameters,cycleRef,get,now){
  N(mode(parameters)&&parameters.preAgendaPolicyRef?.kind==='policy','CONFIG_CAPABILITY_DISABLED','PREAGENDA_CONFIGURATION_REQUIRED');
  const r=parameters.preAgendaPolicyRef,p=get(M.path(r));
  N(p?.approved===true&&p.active===true&&p.usage==='HISTORICAL_PREAGENDA'&&eq(p.cycleRef,cycleRef)&&p.decisionBasis?.reason&&p.decisionBasis?.evidenceRef&&M.ms(p.validFrom)<=now&&M.ms(p.validUntil)>now,'HUMAN_DECISION_REQUIRED','PREAGENDA_POLICY_REQUIRED');
  N(Array.isArray(p.items)&&p.items.length>0&&p.items.length<=12&&new Set(p.items.map(i=>i.activityRef?.id)).size===p.items.length,'VALIDATION_ERROR','PREAGENDA_SCOPE');
  for(const i of p.items){N(i.activityRef?.kind==='activity'&&i.presentationRef?.kind==='presentation'&&i.personRef?.kind==='person','VALIDATION_ERROR','PREAGENDA_REFERENCE');M.civil(i.civilDate);N(Array.isArray(i.deadlineDecisions)&&i.deadlineDecisions.length<=3&&new Set(i.deadlineDecisions.map(d=>d.stage)).size===i.deadlineDecisions.length,'VALIDATION_ERROR','PREAGENDA_DEADLINES');for(const d of i.deadlineDecisions)N(['CALENDAR','THEME','TITLE'].includes(d.stage)&&['PREEXISTING_SCHEDULE','REQUIREMENT_NOT_APPLICABLE_AT_ESTABLISHMENT','SPECIFIC_REGULARIZATION'].includes(d.basis)&&typeof d.reason==='string'&&d.reason.trim()&&d.evidenceRef&&Number.isSafeInteger(d.until)&&d.until<=C.civilStart(i.civilDate),'HUMAN_DECISION_REQUIRED','INDIVIDUAL_DEADLINE_DECISION_REQUIRED');}
  return {ref:r,document:p};
 }
 function parameters(v,cycleRef,get,now){N(['INCIDENCE','RECOGNIZED_PREAGENDA'].includes(v.planningMode||'INCIDENCE'),'VALIDATION_ERROR','PLANNING_MODE');if(mode(v)){policy(v,cycleRef,get,now);N(!v.quota&&!v.early,'CONFIG_CAPABILITY_DISABLED','INCIDENCE_DEPENDENT_FEATURE');}return v;}
 function basis(unit,get,now){
  const b=unit.preAgendaBasisRef&&get(M.path(unit.preAgendaBasisRef));
  N(b?.kind==='HISTORICAL_PREAGENDA'&&b.approved===true&&b.active===true&&eq(b.unitRef,M.ref('unit',unit.id))&&eq(b.personRef,unit.personRef)&&eq(b.cycleRef,unit.cycleRef),'HUMAN_DECISION_REQUIRED','PREAGENDA_BASIS_REQUIRED');
  const cycle=get(M.path(unit.cycleRef)),cfg=cycle?.currentConfigRef&&get(M.path(cycle.currentConfigRef));N(cfg?.approved&&!cfg.parameters.reviewRequired,'CONFIG_CAPABILITY_DISABLED','PREAGENDA_CONFIGURATION_REQUIRED');
  const p=policy(cfg.parameters,unit.cycleRef,get,now),item=p.document.items.find(i=>eq(i.activityRef,b.activityRef)&&eq(i.presentationRef,b.presentationRef));
  N(eq(p.ref,b.policyRef)&&item&&M.hash(item)===b.itemDigest&&eq(item.personRef,unit.personRef),'STALE_REVISION','PREAGENDA_DECISION_CHANGED');
  const activity=get(M.path(b.activityRef));N(activity?.status==='SCHEDULED'&&!activity.realizationRef&&eq(activity.cycleRef,unit.cycleRef)&&eq(activity.personRef,unit.personRef)&&activity.civilDate===item.civilDate,'INVALID_STATE_TRANSITION','PREAGENDA_DATE_OR_OWNER_CHANGED');
  return {basis:b,item,activity,policy:p};
 }
 function deadline(x,s,stage,unit){
  let u=unit;
  if(!u&&stage==='SESSION'&&s.activityRef){const a=x.get(M.path(s.activityRef));if(!a)return false;const candidates=(a.historicalPresentationRefs||[]).map(pr=>refs(s.activityRef,pr).unitRef).map(r=>x.get(M.path(r))).filter(Boolean);if(candidates.length!==1)return false;u=candidates[0];}
  if(!u?.preAgendaBasisRef)return false;
  const b=x.get(M.path(u.preAgendaBasisRef));if(b?.kind!=='HISTORICAL_PREAGENDA')return false;
  const z=basis(u,x.get,x.now),d=z.item.deadlineDecisions.find(d=>d.stage===(stage==='SESSION'?'CALENDAR':stage));
  return !!(d&&d.until>x.now&&s.civilDate===z.item.civilDate&&eq(s.activityRef,z.basis.activityRef)&&s.civilDate>M.day(x.now));
 }
 function recognize(x){
  const {p,load,get,config,actor,now,revision,event}=x,a=load(p.activityRef,true),pr=load(p.presentationRef,true),cfg=config();
  N(actor.kind==='BOARD_ACTOR'&&typeof p.reason==='string'&&p.reason.trim(),'HUMAN_DECISION_REQUIRED');
  N(a.recordMode==='NEW_V2'&&a.promotion?.fromMode==='HISTORICAL_PARTIAL'&&a.status==='SCHEDULED'&&!a.realizationRef&&a.civilDate>M.day(now),'INVALID_STATE_TRANSITION','FUTURE_PROMOTED_ACTIVITY_REQUIRED');
  const cycle=load(a.cycleRef,true);N(['DRAFT','OPEN'].includes(cycle.status)&&M.gap(cycle.civilStart,a.civilDate)>=0&&M.gap(a.civilDate,cycle.civilEnd)>=0,'INVALID_STATE_TRANSITION','PREAGENDA_CYCLE_REQUIRED');
  N(pr.recordMode==='HISTORICAL_PARTIAL'&&!pr.unitRef&&!pr.continuity&&!pr.realizationRef&&f4PresentationContinuity_().sourceForActivity(pr,a),'INVALID_STATE_TRANSITION','IMPORTED_PRESENTATION_REQUIRED');
  const z=policy(cfg.parameters,a.cycleRef,get,now),item=z.document.items.find(i=>eq(i.activityRef,p.activityRef)&&eq(i.presentationRef,p.presentationRef));
  N(item&&eq(item.personRef,a.personRef)&&eq(item.personRef,pr.personRef?.ref)&&item.civilDate===a.civilDate,'AUTHORIZATION_DENIED','PREAGENDA_EXACT_SCOPE');
  const e=pr.provenanceHead&&get(M.path(pr.provenanceHead));N(e?.literal?.ID_APRESENTACAO===pr.id&&e.literal.ID_ATIVIDADE===a.id&&p.sourceHash===C.hash(F.encode(e)),'STALE_REVISION','PRESENTATION_SOURCE_CHANGED');
  const r=refs(p.activityRef,p.presentationRef);N(!load(r.unitRef,false,true)&&!load(r.basisRef,false,true),'CONFLICT','PREAGENDA_ALREADY_RECOGNIZED');
  revision(r.basisRef,{kind:'HISTORICAL_PREAGENDA',approved:true,active:true,cycleRef:a.cycleRef,personRef:item.personRef,...r,activityRef:p.activityRef,presentationRef:p.presentationRef,policyRef:z.ref,policyRevision:z.document.revision,itemDigest:M.hash(item),sourceHash:p.sourceHash,recognizedAt:M.stamp(now),reason:p.reason.trim()});
  revision(r.unitRef,{cycleRef:a.cycleRef,personRef:item.personRef,preAgendaBasisRef:r.basisRef,kind:'RECOGNIZED_PREAGENDA',state:'ELIGIBLE'});
  x.bumpCapacity();event('history',{action:'HISTORICAL_PREAGENDA_RECOGNIZED',subjectRef:p.presentationRef,activityRef:p.activityRef,ownerPersonRef:item.personRef.id,visibility:'BOARD',summary:p.reason.trim(),occurredAt:M.stamp(now)});
  return {...r,activityRef:p.activityRef,presentationRef:p.presentationRef,obligationCreated:false,presentationCreated:false,sessionCreated:false};
 }
 function availableDeadline(x,s,stage,u){try{return !!s&&deadline(x,s,stage,u);}catch(e){if(['HUMAN_DECISION_REQUIRED','STALE_REVISION','INVALID_STATE_TRANSITION','CONFIG_CAPABILITY_DISABLED','UNRESOLVED_DEPENDENCY'].includes(e.code))return false;throw e;}}
 return {mode,refs,policy,parameters,basis,deadline,availableDeadline,recognize};
}
function f4ActivitiesPreagendaCommands_(x){return {HISTORICAL_PREAGENDA_RECOGNIZE:()=>f4ActivitiesPreagenda_().recognize(x)};}
