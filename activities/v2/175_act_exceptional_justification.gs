// An explicit administrative reopening changes eligibility to submit only.
// Neither an absence nor the ordinary 48-hour rule changes when this is opened.
function f4AttendanceExceptionalWindow_(get,collection,fact,activity,now){
 const M=f4ActivitiesModel_();
 if(fact.status!=='ABSENT'||fact.applicableAtDate===false||fact.countsAbsence===false||activity.typeSnapshot?.countsAbsence===false||['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(activity.status))return null;
 const candidates=(collection.exceptionalWindowRefs||[]).map(ref=>({ref,value:get(M.path(ref))})).filter(({value:w})=>w&&w.status==='OPEN'&&M.ms(w.openedAt)<=now&&now<M.ms(w.newDeadline)&&w.activityRefs.some(r=>r.id===activity.id)&&(!w.personRefs.length||w.personRefs.some(r=>r.id===fact.personRef.id)));
 candidates.sort((a,b)=>M.ms(b.value.openedAt)-M.ms(a.value.openedAt)||a.ref.id.localeCompare(b.ref.id));
 return candidates[0]||null;
}

function f4ActivitiesExceptionalJustificationCommands_(x){
 const {M,N,p,load,revision,event,save,now,actor}=x;
 const reason=()=>{N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'INCOMPLETE_REQUEST');return p.reason.trim();};
 function invalidate(){const ref=M.ref('attendanceRegistry','current'),old=load(ref,false,true);save(ref,{epoch:(old?.epoch||0)+1});}
 function log(ref,action){event('history',{action,subjectRef:ref,visibility:'BOARD',ownerPersonRef:'',summary:reason(),occurredAt:M.stamp(now)});}
 return {
  ATTENDANCE_EXCEPTIONAL_WINDOW_OPEN(){
   N(actor.kind==='BOARD_ACTOR','AUTHORIZATION_DENIED');
   N(Array.isArray(p.activityRefs)&&p.activityRefs.length>0&&p.activityRefs.length<=20&&new Set(p.activityRefs.map(M.path)).size===p.activityRefs.length,'VALIDATION_ERROR');
   N(Array.isArray(p.personRefs)&&p.personRefs.length<=30&&new Set(p.personRefs.map(M.path)).size===p.personRefs.length,'VALIDATION_ERROR');
   N(Number.isSafeInteger(p.newDeadline)&&p.newDeadline>now,'DEADLINE_VIOLATION');
   N(['PORTAL_UNAVAILABLE','SYSTEM_FAILURE','ADMINISTRATIVE_REGULARIZATION','OTHER_JUSTIFIED'].includes(p.reasonCode),'VALIDATION_ERROR');
   p.personRefs.forEach(r=>{N(r.kind==='person','VALIDATION_ERROR');load(r);});
   const collections=p.activityRefs.map(r=>{
    N(r.kind==='activity','VALIDATION_ERROR');const a=load(r,true),ref=M.ref('attendanceCollection',r.id),c=load(ref,true);
    N(!['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(a.status)&&a.typeSnapshot?.countsAbsence!==false&&a.typeSnapshot?.allowsJustification&&c.lastFinalizedAt,'INVALID_STATE_TRANSITION','EXCEPTION_REQUIRES_COMPLETED_ATTENDANCE');
    N((c.exceptionalWindowRefs||[]).length<32,'SCOPE_TOO_LARGE');return {ref,c};
   });
   const ref=event('exceptionalWindow',{status:'OPEN',activityRefs:p.activityRefs,personRefs:p.personRefs,openedAt:M.stamp(now),openedBy:actor.personRef,newDeadline:M.stamp(p.newDeadline),reasonCode:p.reasonCode,reason:reason(),originalRuleReference:collections.map(({ref,c})=>({activityId:ref.id,ordinaryDeadline:c.justificationEndsAt,policyRef:c.policyRef||null,policyRevision:c.policyRevision||null}))});
   for(const {ref:cr,c}of collections)revision(cr,{...c,exceptionalWindowRefs:[...(c.exceptionalWindowRefs||[]),ref]});
   invalidate();log(ref,'EXCEPTIONAL_JUSTIFICATION_WINDOW_OPENED');return {windowRef:ref,status:'OPEN'};
  },
  ATTENDANCE_EXCEPTIONAL_WINDOW_CLOSE(){
   N(actor.kind==='BOARD_ACTOR'&&p.windowRef.kind==='exceptionalWindow','AUTHORIZATION_DENIED');const w=load(p.windowRef,true);
   N(w.status==='OPEN'&&['CLOSE','CANCEL'].includes(p.action),'INVALID_STATE_TRANSITION');
   const status=p.action==='CANCEL'?'CANCELLED':'CLOSED';revision(p.windowRef,{...w,status,closedAt:M.stamp(now),closedBy:actor.personRef,closureReason:reason()});
   invalidate();log(p.windowRef,'EXCEPTIONAL_JUSTIFICATION_WINDOW_'+status);return {windowRef:p.windowRef,status};
  }
 };
}
