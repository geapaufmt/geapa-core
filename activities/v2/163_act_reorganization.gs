// Bounded calendar proposals. All affected assignments change in one Core commit.
function f4ActivitiesReorganization_(x){
 const {M,N,p,load,revision,event,config,capacity,actor,now}=x,D=f4ActivitiesDomain_(x),safety=f4ActivitiesCalendarSafety_(x);
 function plan(moves){
  const sessions=new Map(),units=new Map(),affected=new Map(),basis=new Map();
  function remember(r,d){basis.set(M.path(r),{ref:r,revision:d.revision});return d;}
  function session(r){if(!sessions.has(r.id)){const s=remember(r,load(r,true));N(s.cycleRef.id===x.cycleRef.id,'VALIDATION_ERROR','CROSS_CYCLE');sessions.set(r.id,{ref:r,document:s});}return sessions.get(r.id).document;}
  for(const move of moves){const u=remember(move.unitRef,load(move.unitRef,true));N(u.cycleRef.id===x.cycleRef.id,'VALIDATION_ERROR','CROSS_CYCLE');const a=u.assignmentRef&&remember(u.assignmentRef,load(u.assignmentRef,true));if(a){safety.movable(u,a);session(a.sessionRef);N(a.sessionRef.id!==move.sessionRef.id,'INVALID_STATE_TRANSITION','DESTINATION_UNCHANGED');}else{D.eligible(u);N(!u.presentationRef||!load(u.presentationRef,true).realizationRef,'INVALID_STATE_TRANSITION','REALIZED_CALENDAR_HISTORY');N(!u.dynamicOccupancyRef||load(u.dynamicOccupancyRef,true).state==='RELEASED','HUMAN_DECISION_REQUIRED','DYNAMIC_REORGANIZATION_REQUIRED');}
   const dest=session(move.sessionRef);D.date(dest);const activity=dest.activityRef&&load(dest.activityRef,true);N(activity?.status==='SCHEDULED'&&activity.civilDate===dest.civilDate,'UNRESOLVED_DEPENDENCY','SESSION_ACTIVITY_REQUIRED');units.set(u.id,{ref:move.unitRef,document:u,assignment:a});affected.set(u.id,move.unitRef);
  }
  for(const {document:s}of sessions.values())for(const slot of s.slots){const u=remember(slot.unitRef,load(slot.unitRef,true));N(u.cycleRef.id===x.cycleRef.id,'VALIDATION_ERROR','CROSS_CYCLE');affected.set(u.id,slot.unitRef);}
  N(affected.size<=16,'SCOPE_TOO_LARGE','AFFECTED_CALENDAR_SCOPE');
  const virtual=new Map([...sessions].map(([id,v])=>[id,{...v.document,slots:v.document.slots.filter(slot=>!units.has(slot.unitRef.id))}]));
  for(const move of moves){const s=virtual.get(move.sessionRef.id);N(config().parameters.shared||s.slots.length===0,'CONFIG_CAPABILITY_DISABLED','SHARED_OFF');D.slot(s,move.unitRef);s.slots.push({unitRef:move.unitRef});}
  return {sessions,units,affected:[...affected.values()].sort((a,b)=>a.id.localeCompare(b.id)),basis:[...basis.values()]};
 }
 function assessment(moves){const a=x.manifest(p.assessmentRef),g=load(M.ref('capacity','registry'),true),h=capacity(),c=load(x.cycleRef,true),change={moves};N(a.cycleRef?.id===c.id&&a.configRef&&M.path(a.configRef)===M.path(c.currentConfigRef)&&a.membersEpoch===g.membersEpoch&&a.cycleRegistryEpoch===g.cycleRegistryEpoch&&a.basisEpoch===h.capacityEpoch,'STALE_REVISION','ASSESSMENT_CAPACITY_EPOCH');N(a.sourceCoverageComplete&&a.preservesAcquiredRights&&a.feasible&&a.proposedChangeDigest===M.hash(change)&&M.hash(a.proposedChange)===M.hash(change),'CAPACITY_EXHAUSTED','CALENDAR_RESIDUAL_FEASIBILITY');}
 return {
 CALENDAR_CHANGE_PROPOSE(){
  const cycle=load(p.cycleRef,true);N(cycle.status==='OPEN','INVALID_STATE_TRANSITION');config();capacity();N(['VOLUNTARY','ADMINISTRATIVE'].includes(p.mode),'VALIDATION_ERROR');N(typeof p.reason==='string'&&p.reason.length<=2000&&(p.mode==='VOLUNTARY'||p.reason.trim()),'INCOMPLETE_REQUEST','DECISION_REASON_REQUIRED');const deadline=M.ms(p.deadline);N(deadline>now&&deadline<=f4Core_().civilStart(cycle.civilEnd)+86400000,'DEADLINE_VIOLATION');const moves=f4ActivitiesCalendarMoves_(p.moves),scope=plan(moves);N(scope.basis.length<=48,'SCOPE_TOO_LARGE');
  if(p.mode==='VOLUNTARY'){const w=config().parameters.reorganizationWindow;N(w&&Number.isSafeInteger(w.opensAt)&&Number.isSafeInteger(w.endsAtExclusive)&&now>=w.opensAt&&now<w.endsAtExclusive&&deadline<=w.endsAtExclusive,'CONFIG_CAPABILITY_DISABLED','REORGANIZATION_WINDOW_CLOSED');}
  const changeSummary=moves.map(move=>{const a=scope.units.get(move.unitRef.id).assignment;return {unitRef:move.unitRef,...(a?{fromSessionRef:a.sessionRef,fromCivilDate:scope.sessions.get(a.sessionRef.id).document.civilDate}:{}),toSessionRef:move.sessionRef,toCivilDate:scope.sessions.get(move.sessionRef.id).document.civilDate};});
  const flowCaseRef=event('flow',{kind:'CALENDAR_CHANGE',state:'OPEN',cycleRef:p.cycleRef,mode:p.mode,moves,changeSummary,unitRefs:scope.affected,movingUnitRefs:moves.map(v=>v.unitRef),basisRevisionVector:scope.basis,configRef:cycle.currentConfigRef,proposalRevision:1,deadline:M.stamp(deadline),reason:p.reason.trim(),proposedBy:actor.personRef});return {flowCaseRef,state:'OPEN'};
 },
 CALENDAR_CHANGE_CONFIRM(){
  const f=load(p.flowCaseRef,true),cycle=load(x.cycleRef,true);N(f.kind==='CALENDAR_CHANGE'&&f.state==='OPEN'&&M.ms(f.deadline)>now&&cycle.status==='OPEN','INVALID_STATE_TRANSITION','FLOW');N(M.path(f.configRef)===M.path(cycle.currentConfigRef),'STALE_REVISION','PROPOSAL_CONFIG_CHANGED');N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'INCOMPLETE_REQUEST','DECISION_REASON_REQUIRED');
  for(const b of f.basisRevisionVector)N(load(b.ref,true).revision===b.revision,'STALE_REVISION','PROPOSAL_CONTEXT_CHANGED');
  const moves=f4ActivitiesCalendarMoves_(f.moves),scope=plan(moves);N(M.hash(scope.affected)===M.hash(f.unitRefs),'STALE_REVISION','AFFECTED_SCOPE_CHANGED');
  if(f.mode==='VOLUNTARY'){const w=config().parameters.reorganizationWindow;N(w&&now>=w.opensAt&&now<w.endsAtExclusive,'DEADLINE_VIOLATION','REORGANIZATION_WINDOW_CLOSED');for(const r of scope.affected){const u=load(r),c=x.get(M.path(p.flowCaseRef)+'/consents/'+u.personRef.id);N(c?.choice==='ACCEPT'&&c.proposalRevision===f.proposalRevision,'HUMAN_DECISION_REQUIRED','CONSENT');}}
  else N(f.mode==='ADMINISTRATIVE'&&actor.kind==='BOARD_ACTOR','AUTHORIZATION_DENIED');assessment(moves);
  for(const {document:s}of scope.sessions.values())s.slots=s.slots.filter(slot=>!scope.units.has(slot.unitRef.id));
  for(const move of moves){const {document:u,assignment:a}=scope.units.get(move.unitRef.id),dest=scope.sessions.get(move.sessionRef.id).document;let ar=u.assignmentRef;
   if(!ar){ar=M.ref('assignment',M.hash({operationId:x.operationId,unit:u.id}).slice(0,32));const cr=M.ref('claim',M.hash({kind:'ACTIVE_ASSIGNMENT',unit:u.id}).slice(0,32)),claim=load(cr,false,true);N(!claim||claim.status==='RELEASED','CONFLICT','ASSIGNMENT_CLAIM');x.save(cr,{claimKind:'ACTIVE_ASSIGNMENT',canonicalKeyComponents:{unitId:u.id},targetRef:ar,status:'ACTIVE'});x.save(move.unitRef,{...u,assignmentRef:ar});}
   revision(ar,{...a,unitRef:move.unitRef,personRef:u.personRef,cycleRef:x.cycleRef,status:'CONFIRMED',sessionRef:move.sessionRef,assignedCivilDate:dest.civilDate,slot:dest.slots.length,mode:f.mode,lastCalendarChangeRef:p.flowCaseRef});dest.slots.push({unitRef:move.unitRef,assignmentRef:ar});safety.synchronize(u,ar,move.sessionRef);
  }
  for(const {ref,document:s}of scope.sessions.values()){s.compositionRevision++;revision(ref,s);}
  for(const move of moves)D.materialize(move.unitRef);
  revision(p.flowCaseRef,{...f,state:'CONFIRMED',confirmedAt:M.stamp(now),decision:{reason:p.reason.trim(),actorAuthorityBasis:x.authorityBasis}});x.bumpCapacity();return {state:'CONFIRMED',flowCaseRef:p.flowCaseRef};
 }
 };
}
