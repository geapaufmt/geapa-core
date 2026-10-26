// One confirmation in the existing model form creates one activity and one
// assignment atomically. Its cycle-scoped capability is explicitly granted;
// ordinary activity editing does not grant presentation allocation authority.
function f4ActivitiesMemberActivityCommands_(x){
 const {M,N,p,load,revision}=x;
 return {MEMBER_PRESENTATION_ACTIVITY_CREATE(){
  const cycle=load(p.cycleRef,true),cfg=x.config(),unit=load(p.unitRef,true),type=load(p.typeRef,true),D=f4ActivitiesDomain_(x);
  N(cycle.status==='OPEN'&&unit.cycleRef.id===cycle.id,'INVALID_STATE_TRANSITION');
  N(type.status==='ACTIVE'&&type.portalModel?.memberPresentation,'CONFIG_CAPABILITY_DISABLED');
  N(!load(p.activityRef,true,true)&&!load(p.sessionRef,true,true)&&!unit.assignmentRef&&!unit.presentationRef,'CONFLICT','ACTIVE_ASSIGNMENT_EXISTS');
  N(unit.obligationRef,'UNRESOLVED_DEPENDENCY','APPROVED_INCIDENCE_REQUIRED');D.eligible(unit);
  const date=M.day(p.startsAt);D.date({civilDate:date});
  N(M.gap(cycle.civilStart,date)>=0&&M.gap(date,cycle.civilEnd)>=0,'VALIDATION_ERROR','DATE_OUTSIDE_CYCLE');
  f4ActivitiesMemberActivityPerson_(x.get,unit.personRef,p.startsAt);
  // Adding a new, fully occupied single slot leaves all existing places intact.
  // A current complete proof for the existing calendar is therefore sufficient.
  const proof=x.manifest(p.assessmentRef),head=x.capacity(),registry=load(M.ref('capacity','registry'),true);
  N(proof.cycleRef?.id===cycle.id&&M.path(proof.configRef)===M.path(cycle.currentConfigRef)&&proof.basisEpoch===head.capacityEpoch&&proof.membersEpoch===registry.membersEpoch&&proof.cycleRegistryEpoch===registry.cycleRegistryEpoch,'STALE_REVISION','ASSESSMENT_CAPACITY_EPOCH');
  N(proof.sourceCoverageComplete&&proof.feasible&&proof.preservesAcquiredRights&&M.hash(proof.proposedChange)===M.hash({}),'CAPACITY_EXHAUSTED','CALENDAR_RESIDUAL_FEASIBILITY');
  const modelContent={principal:{kind:'CORE_PERSON',personRef:unit.personRef,participantType:'MEMBRO',participationRole:type.portalModel.principalRole||'APRESENTADOR',institutionLabel:''}};
  const activityPayload={cycleRef:p.cycleRef,...(p.academicSemesterRef?{academicSemesterRef:p.academicSemesterRef}:{}),activityRef:p.activityRef,typeRef:p.typeRef,title:type.name,startsAt:p.startsAt,endsAt:p.endsAt,location:p.location,status:'SCHEDULED',details:p.details,modelContent,decision:p.decision};
  f4AttendanceCommands_({...x,p:activityPayload,memberPresentationCreation:true}).ACTIVITY_SAVE();
  f4ActivitiesPlanning_({...x,p:{cycleRef:p.cycleRef,sessionRef:p.sessionRef,activityRef:p.activityRef,civilDate:date,capacity1or2:1,destination:'ORDINARY',decision:p.decision}}).SESSION_CONFIGURE();
  const assignmentRef=f4ActivitiesCalendar_(x).confirmNewAssignment(p.sessionRef,p.unitRef,cfg.parameters.calendarMode);
  D.materialize(p.unitRef);x.bumpCapacity();
  return {activityRef:p.activityRef,sessionRef:p.sessionRef,assignmentRef,unitRef:p.unitRef};
 }};
}
function f4ActivitiesMemberActivityPerson_(get,ref,at){
 const M=f4ActivitiesModel_(),N=M.need,head=get('memberAggregateHeads/'+ref.id);
 N(ref.kind==='person'&&head?.paths?.includes(M.path(ref)),'UNRESOLVED_DEPENDENCY','MEMBERS_HEAD');
 const paths=head.paths.filter(p=>/^member(Identities|Links|Events)\//.test(p));N(paths.length<=60,'SCOPE_TOO_LARGE');
 const docs=Object.fromEntries(paths.map(p=>[p,get(p)]));N(Object.values(docs).every(Boolean)&&!docs[M.path(ref)].technicalAccount,'UNRESOLVED_DEPENDENCY');
 const state=f4MemberModel_(true).derive(docs,ref.id,head.revision,at).state;
 N(state.state==='MEMBRO_EFETIVO'&&!state.blocked,'AUTHORIZATION_DENIED','PRESENTER_NOT_EFFECTIVE_AT_DATE');
 return docs[M.path(ref)];
}
