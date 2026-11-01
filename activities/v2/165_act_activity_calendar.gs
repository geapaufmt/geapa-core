// The existing activity edit coordinates its linked session behind the same UI.
function f4ActivitiesActivityCalendarChange_(activity,payload){
 const M=f4ActivitiesModel_();return {sessionDate:{sessionRef:activity.presentationSessionRef,activityRef:payload.activityRef,startsAt:payload.startsAt,endsAt:payload.endsAt,civilDate:M.day(payload.startsAt)}};
}
function f4ActivitiesApplyActivityCalendar_(x,activity){
 const {M,N,p,load,revision}=x,change=f4ActivitiesActivityCalendarChange_(activity,p);
 N(p.calendarAssessmentRef||p.sessionRef,'UNRESOLVED_DEPENDENCY','SESSION_REORGANIZATION_REQUIRED');const s=load(activity.presentationSessionRef,true),c=load(s.cycleRef,true);N(x.cycleRef?.id===c.id&&(!p.sessionRef||p.sessionRef.id===s.id),'VALIDATION_ERROR','SESSION_ACTIVITY_REQUIRED');x.config();
 N(s.activityRef?.id===activity.id&&['DRAFT','OPEN'].includes(c.status),'INVALID_STATE_TRANSITION','SESSION_ACTIVITY_REQUIRED');
 if(c.status==='DRAFT')N(s.slots.length===0,'INVALID_STATE_TRANSITION','DRAFT_SESSION_OCCUPIED');
 else {
 N(p.calendarAssessmentRef,'UNRESOLVED_DEPENDENCY','SESSION_REORGANIZATION_REQUIRED');const a=x.manifest(p.calendarAssessmentRef),h=x.capacity(),g=load(M.ref('capacity','registry'),true);
 N(a.cycleRef?.id===c.id&&a.configRef&&M.path(a.configRef)===M.path(c.currentConfigRef)&&a.membersEpoch===g.membersEpoch&&a.cycleRegistryEpoch===g.cycleRegistryEpoch&&a.basisEpoch===h.capacityEpoch,'STALE_REVISION','ASSESSMENT_CAPACITY_EPOCH');
 N(a.sourceCoverageComplete===true&&a.feasible===true&&a.preservesAcquiredRights===true&&a.proposedChangeDigest===M.hash(change)&&M.hash(a.proposedChange)===M.hash(change),'CAPACITY_EXHAUSTED','SESSION_RESIDUAL_FEASIBILITY');
 }
 N(p.status==='SCHEDULED'&&activity.status==='SCHEDULED'&&M.gap(c.civilStart,change.sessionDate.civilDate)>=0&&M.gap(change.sessionDate.civilDate,c.civilEnd)>=0,'VALIDATION_ERROR','DATE_OUTSIDE_CYCLE');
 if(activity.civilDate!==change.sessionDate.civilDate)f4ActivitiesDomain_(x).date({civilDate:change.sessionDate.civilDate});
 for(const slot of s.slots){const u=load(slot.unitRef,true),assignment=load(slot.assignmentRef,true);f4ActivitiesCalendarSafety_(x).movable(u,assignment);N(assignment.sessionRef.id===s.id,'UNRESOLVED_DEPENDENCY','ASSIGNMENT_SLOT_DRIFT');revision(slot.assignmentRef,{...assignment,assignedCivilDate:change.sessionDate.civilDate});}
 revision(activity.presentationSessionRef,{...s,civilDate:change.sessionDate.civilDate,configRevision:s.configRevision+1,compositionRevision:s.compositionRevision+1});x.bumpCapacity();
}
function f4ActivitiesPrepareActivityCalendar_(options,request,activity){
 const M=f4ActivitiesModel_(),N=M.need,db=f4Firestore_().connect(options.config),clock=options.clock||Date.now,ports=options.ports,p=request.payload;
 const publication=request.commandType==='ACT.ACTIVITY_PUBLICATION_SET';
 if(!activity?.presentationSessionRef||(publication?!['CANCEL','REOPEN'].includes(p.action):M.ms(activity.startsAt)===p.startsAt&&M.ms(activity.endsAt)===p.endsAt))return request;
 const session=db.get(M.path(activity.presentationSessionRef));N(session?.activityRef?.id===activity.id,'UNRESOLVED_DEPENDENCY','SESSION_ACTIVITY_REQUIRED');const cycle=db.get(M.path(session.cycleRef));N(cycle?.currentConfigRef,'UNRESOLVED_DEPENDENCY');
 request={...request,payload:{...p,sessionRef:activity.presentationSessionRef}};
 if(cycle.status==='DRAFT'||publication&&p.action==='REOPEN')return request;
 const placeholder=M.ref('assessment',M.hash({operationId:request.operationId,kind:'ACTIVITY_CALENDAR_SCOPE'}).slice(0,32)),intent={...request,payload:{...request.payload,calendarAssessmentRef:placeholder}};
 let authorized=ports.prepare(intent);ports.authorizePrepared(path=>db.get(path),authorized);let h=db.get('presentationCapacityHeads/'+cycle.id),g=db.get('presentationCapacityHeads/registry');
 if(!(h?.state==='SEALED'&&h.synchronizedMembersEpoch===g?.membersEpoch&&h.synchronizedCycleRegistryEpoch===g?.cycleRegistryEpoch&&M.ms(h.validUntil)>clock()+5000)){
  const jobs=options.jobs||[],index=jobs.findIndex(j=>j.kind==='CAPACITY_SYNC'&&j.cycleId===cycle.id);N(index>=0&&options.pipeline,'UNRESOLVED_DEPENDENCY','CAPACITY_REFRESH_NOT_CONFIGURED');f4ActivitiesConfiguredJobs_({config:options.config,db,pipeline:options.pipeline,jobs,clock}).run(index);
 }
 authorized=ports.prepare(intent);
 const authorize=get=>{N(get(M.path(p.activityRef))?.revision===activity.revision&&get(M.path(activity.presentationSessionRef))?.revision===session.revision&&get(M.path(session.cycleRef))?.revision===cycle.revision,'STALE_REVISION','ACTIVITY_CALENDAR_CHANGED');ports.authorizePrepared(get,authorized);return true;};
 const change=publication?{sessionConfig:{sessionRef:activity.presentationSessionRef,capacity:session.config.capacity,destination:'CANCELLED'}}:f4ActivitiesActivityCalendarChange_(activity,p);
 const preparation=f4ActivitiesPreparation_({db,clock,capacitySource:f4ActivitiesCapacitySource_({db,clock,config:options.config}),authorize,executionGate:()=>authorize(path=>db.get(path))}),calendarAssessmentRef=preparation.prepareCapacity(null,session.cycleRef,cycle.currentConfigRef,change);
 authorize(path=>db.get(path));return {...request,payload:{...request.payload,calendarAssessmentRef}};
}
// Cancelling an occupied activity requires the existing reorganization flow
// first. Removing an empty place still proves capacity for everyone remaining.
function f4ActivitiesApplyActivityCancellation_(x,activity){
 const {M,N,p,load,revision}=x;if(!activity.presentationSessionRef||!['CANCEL','REOPEN'].includes(p.action))return;
 N(p.sessionRef?.id===activity.presentationSessionRef.id,'UNRESOLVED_DEPENDENCY','SESSION_REORGANIZATION_REQUIRED');
 const s=load(p.sessionRef,true),c=load(s.cycleRef,true);N(s.activityRef?.id===activity.id&&x.cycleRef?.id===c.id&&['DRAFT','OPEN'].includes(c.status),'INVALID_STATE_TRANSITION','SESSION_ACTIVITY_REQUIRED');x.config();
 N(s.slots.length===0,'INVALID_STATE_TRANSITION','SESSION_HAS_COMMITMENTS');
 if(p.action==='CANCEL'){
  N(activity.status==='SCHEDULED'&&s.config.destination!=='CANCELLED','INVALID_STATE_TRANSITION');
  if(c.status==='OPEN'){
   N(p.calendarAssessmentRef,'UNRESOLVED_DEPENDENCY','SESSION_RESIDUAL_ASSESSMENT_REQUIRED');const a=x.manifest(p.calendarAssessmentRef),h=x.capacity(),g=load(M.ref('capacity','registry'),true),change={sessionConfig:{sessionRef:p.sessionRef,capacity:s.config.capacity,destination:'CANCELLED'}};
   N(a.cycleRef?.id===c.id&&a.configRef&&M.path(a.configRef)===M.path(c.currentConfigRef)&&a.membersEpoch===g.membersEpoch&&a.cycleRegistryEpoch===g.cycleRegistryEpoch&&a.basisEpoch===h.capacityEpoch,'STALE_REVISION','ASSESSMENT_CAPACITY_EPOCH');
   N(a.sourceCoverageComplete===true&&a.feasible===true&&a.preservesAcquiredRights===true&&a.proposedChangeDigest===M.hash(change)&&M.hash(a.proposedChange)===M.hash(change),'CAPACITY_EXHAUSTED','SESSION_RESIDUAL_FEASIBILITY');
  }
  revision(p.sessionRef,{...s,config:{...s.config,destination:'CANCELLED'},activityCancellation:{previousDestination:s.config.destination,activityRevision:activity.revision},configRevision:s.configRevision+1});
 }else{
  N(activity.status==='CANCELLED'&&s.config.destination==='CANCELLED'&&['ORDINARY','EXCLUSIVE','RESERVE','BLOCKED','DYNAMIC_EXCLUSIVE'].includes(s.activityCancellation?.previousDestination),'UNRESOLVED_DEPENDENCY','SESSION_REOPEN_BASIS_REQUIRED');
  revision(p.sessionRef,{...s,config:{...s.config,destination:s.activityCancellation.previousDestination},configRevision:s.configRevision+1});
 }
 x.bumpCapacity();
}
