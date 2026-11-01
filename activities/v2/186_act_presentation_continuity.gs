// A bridge to existing approved planning, never a producer of planning decisions.
// The source and the operational identity remain the same presentation document.
function f4PresentationContinuity_(){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_();
 const eq=(a,b)=>!!a&&!!b&&a.kind===b.kind&&a.id===b.id&&a.parent===b.parent;
 function source(pr){return pr?.recordMode==='HISTORICAL_PARTIAL'?pr:pr?.continuity?.schemaVersion==='presentation-continuity-v1'?pr.continuity.imported:null;}
 function sourceForActivity(pr,a){const old=source(pr);return !!old&&old.activityRef?.resolution==='RESOLVED'&&old.activityRef.ref?.id===a?.id&&(a.historicalPresentationRefs||[]).some(r=>r.kind==='presentation'&&r.id===pr.id);}
 // Existing calendar/theme commands must not materialize a second presentation.
 // Reading the imported identity only defers a write; it grants no authority.
 function deferMaterialization(unit,assignment,get){
  const session=assignment.sessionRef&&get(M.path(assignment.sessionRef)),activity=session?.activityRef&&get(M.path(session.activityRef));
  if(activity?.promotion?.fromMode!=='HISTORICAL_PARTIAL')return false;
  const refs=activity.historicalPresentationRefs||[];N(refs.length<=32,'SCOPE_TOO_LARGE','IMPORTED_PRESENTATION_REFERENCES');
  const matches=refs.map(r=>get(M.path(r))).filter(pr=>sourceForActivity(pr,activity)&&eq(source(pr)?.personRef?.ref,unit.personRef));
  N(matches.length<=1,'UNRESOLVED_DEPENDENCY','AMBIGUOUS_IMPORTED_PRESENTATION');
  if(!matches.length)return false;
  N(matches[0].recordMode==='HISTORICAL_PARTIAL'&&!matches[0].unitRef&&!matches[0].continuity,'CONFLICT','IMPORTED_PRESENTATION_ALREADY_CONNECTED');
  return true;
 }
 function inspect(activity,pr,get,now){
  const reasons=[],add=(ok,code)=>{if(!ok)reasons.push(code);};
  add(activity?.recordMode==='NEW_V2'&&activity.promotion?.fromMode==='HISTORICAL_PARTIAL','ACTIVITY_PROMOTION_REQUIRED');
  add(activity?.status==='SCHEDULED'&&!activity.realizationRef&&activity.civilDate>M.day(now),'FUTURE_SCHEDULED_ACTIVITY_REQUIRED');
  add(sourceForActivity(pr,activity),'IMPORTED_PRESENTATION_LINK_REQUIRED');
  const cycle=activity?.cycleRef?.kind==='cycle'&&get(M.path(activity.cycleRef));
  add(cycle?.status==='OPEN','CYCLE_OPEN_DECISION_REQUIRED');
  const cfg=cycle?.currentConfigRef?.kind==='config'&&get(M.path(cycle.currentConfigRef));
  add(cfg?.approved===true&&!cfg.parameters?.reviewRequired,'APPROVED_CURRENT_CONFIG_REQUIRED');
  add(cfg?.approvalScope!=='OPERATIONAL_AXIS_CATALOG_AND_CLOSED_FEATURES','PLANNING_CONFIG_APPROVAL_REQUIRED');
  add(!!cfg?.parameters?.ordinaryMaterialRequirementRef,'APPROVED_MATERIAL_TEMPLATE_REQUIRED');
  const head=cycle&&get('presentationCapacityHeads/'+cycle.id),registry=cycle&&get('presentationCapacityHeads/registry');
  add(head?.state==='SEALED'&&head.synchronizedMembersEpoch===registry?.membersEpoch&&head.synchronizedCycleRegistryEpoch===registry?.cycleRegistryEpoch&&M.ms(head.validUntil)>now,'CURRENT_CAPACITY_ASSESSMENT_REQUIRED');
  const s=activity?.presentationSessionRef&&get(M.path(activity.presentationSessionRef));
  add(s?.activityRef?.id===activity?.id&&s?.civilDate===activity?.civilDate&&eq(s?.cycleRef,activity?.cycleRef),'CANONICAL_SESSION_REQUIRED');
  return {schemaVersion:'presentation-continuity-readiness-v1',activityId:activity?.id,presentationId:pr?.id,cycleId:cycle?.id||null,cycleState:cycle?.status||null,reviewState:activity?.administrativeReview?.state||null,state:pr?.continuity?'CONNECTED':reasons.length?'BLOCKED':'PLANNING_REFERENCE_REQUIRED',blockers:reasons,grantsAuthority:false};
 }
 function connect(x){
  const {p,load,get,config,capacity,revision,event,now,actor,operationId}=x;
  N(actor.kind==='BOARD_ACTOR'&&typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'HUMAN_DECISION_REQUIRED');
  const a=load(p.activityRef,true),pr=load(p.presentationRef,true),u=load(p.unitRef,true);
  const ready=inspect(a,pr,get,now);N(!ready.blockers.length,'UNRESOLVED_DEPENDENCY',ready.blockers[0]);
  N(pr.recordMode==='HISTORICAL_PARTIAL'&&!pr.realizationRef&&!pr.unitRef&&!pr.continuity&&!u.presentationRef,'INVALID_STATE_TRANSITION','PRESENTATION_ALREADY_CONNECTED');
  const original=source(pr),e=original.provenanceHead&&get(M.path(original.provenanceHead));
  N(e?.literal?.ID_APRESENTACAO===pr.id&&e.literal.ID_ATIVIDADE===a.id&&p.sourceHash===C.hash(F.encode(e)),'STALE_REVISION','PRESENTATION_SOURCE_CHANGED');
  N(original.personRef?.resolution==='RESOLVED'&&eq(original.personRef.ref,u.personRef)&&eq(a.personRef,u.personRef)&&eq(a.cycleRef,u.cycleRef),'UNRESOLVED_DEPENDENCY','CONTINUITY_OWNER_OR_CYCLE');
  const cfg=config();capacity();
  const D=f4ActivitiesDomain_(x);D.eligible(u);
  N(u.assignmentRef&&u.thematicRequestRef,'UNRESOLVED_DEPENDENCY','APPROVED_PLANNING_REQUIRED');
  const assignment=load(u.assignmentRef,true),theme=load(u.thematicRequestRef,true),s=load(a.presentationSessionRef,true);
  N(assignment.status==='CONFIRMED'&&eq(assignment.unitRef,p.unitRef)&&eq(assignment.sessionRef,a.presentationSessionRef)&&eq(assignment.cycleRef,a.cycleRef)&&s.slots.some(v=>eq(v.unitRef,p.unitRef)&&eq(v.assignmentRef,u.assignmentRef)),'UNRESOLVED_DEPENDENCY','CONFIRMED_ALLOCATION_REQUIRED');
  N(theme.workflowState==='APPROVED'&&eq(theme.unitRef,p.unitRef)&&eq(theme.personRef,u.personRef)&&eq(theme.cycleRef,u.cycleRef),'UNRESOLVED_DEPENDENCY','APPROVED_THEME_REQUIRED');
  N(cfg.parameters.ordinaryMaterialRequirementRef,'UNRESOLVED_DEPENDENCY','ORDINARY_MATERIAL_POLICY');
  const continuity={schemaVersion:'presentation-continuity-v1',imported:M.clone(original),activityRef:p.activityRef,sourceHash:p.sourceHash,unitRef:p.unitRef,operationId,connectedAt:M.stamp(now),reason:p.reason.trim(),actorPersonRef:actor.personRef};
  D.materialize(p.unitRef,{presentationRef:p.presentationRef,prior:pr,continuity});
  event('history',{action:'PRESENTATION_CONTINUITY_CONNECTED',subjectRef:p.presentationRef,activityRef:p.activityRef,ownerPersonRef:u.personRef.id,visibility:'BOARD',summary:p.reason.trim(),occurredAt:M.stamp(now)});
  return {activityRef:p.activityRef,presentationRef:p.presentationRef,unitRef:p.unitRef,reviewState:a.administrativeReview?.state,planningCreated:false,realizationCreated:false};
 }
 return {source,sourceForActivity,deferMaterialization,inspect,connect};
}
function f4PresentationContinuityCommands_(x){return {PRESENTATION_CONTINUITY_CONNECT:()=>f4PresentationContinuity_().connect(x)};}
