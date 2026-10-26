// Calendar changes preserve the individual presentation and its material facts.
function f4ActivitiesCalendarSafety_(x){
 const {M,N,load,revision}=x;
 function movable(unit,assignment){
  N(unit.assignmentRef&&assignment.unitRef?.id===unit.id&&unit.assignmentRef.id===assignment.id&&['ACTIVE','CONFIRMED'].includes(assignment.status),'INVALID_STATE_TRANSITION','CURRENT_ASSIGNMENT_REQUIRED');
  N(assignment.cycleRef?.id===x.cycleRef.id&&unit.cycleRef?.id===x.cycleRef.id,'VALIDATION_ERROR','CROSS_CYCLE');
  const presentation=unit.presentationRef&&load(unit.presentationRef,true);
  N(!presentation?.realizationRef,'INVALID_STATE_TRANSITION','REALIZED_CALENDAR_HISTORY');
  const dynamic=unit.dynamicOccupancyRef&&load(unit.dynamicOccupancyRef,true);
  N(!dynamic||dynamic.state==='RELEASED','HUMAN_DECISION_REQUIRED','DYNAMIC_REORGANIZATION_REQUIRED');
 }
 function synchronize(unit,assignmentRef,sessionRef){
  if(!unit.presentationRef)return;
  const presentation=load(unit.presentationRef,true);N(!presentation.realizationRef,'INVALID_STATE_TRANSITION','REALIZED_CALENDAR_HISTORY');
  N(presentation.personRef.id===unit.personRef.id&&presentation.unitRef.id===unit.id,'UNRESOLVED_DEPENDENCY','PRESENTATION_UNIT_DRIFT');
  if(presentation.assignmentRef?.id!==assignmentRef.id||presentation.sessionRef?.id!==sessionRef.id)revision(unit.presentationRef,{...presentation,assignmentRef,sessionRef});
 }
 return {movable,synchronize};
}

// An explicitly approved operational window; absent configuration stays closed.
function f4ActivitiesSelfSchedulingWindow_(parameters,now){
 const w=parameters.selfSchedulingWindow,valid=!!(w&&Object.keys(w).every(k=>['opensAt','endsAtExclusive'].includes(k))&&Number.isSafeInteger(w.opensAt)&&Number.isSafeInteger(w.endsAtExclusive)&&w.opensAt<w.endsAtExclusive);
 return {configured:valid,open:valid&&now>=w.opensAt&&now<w.endsAtExclusive};
}
