// Academic semesters are independent from both activity cycles and board terms.
function f4ActivitiesAssignSemesters_(x,cycleRef,cycle,refs,confirmed){
 const {M,N,load,revision}=x;
 N(Array.isArray(refs)&&refs.length<=24&&new Set(refs.map(M.path)).size===refs.length,'VALIDATION_ERROR');
 const previous=cycle.academicSemesterRefs||[],changed=M.hash(previous)!==M.hash(refs);
 N(!changed||confirmed===true,'HUMAN_DECISION_REQUIRED','CONFIRM_SEMESTER_ASSOCIATION');
 const all=[...new Map([...previous,...refs].map(r=>[r.id,r])).values()];
 for(const ref of all){
  N(ref.kind==='academicSemester','VALIDATION_ERROR');const s=load(ref,true),selected=refs.some(r=>r.id===ref.id);
  N(!s.cycleId||s.cycleId===cycleRef.id,'CONFLICT','SEMESTER_ALREADY_ASSIGNED');
  if(selected)N(M.gap(cycle.civilStart,s.civilStart)>=0&&M.gap(s.civilEnd,cycle.civilEnd)>=0,'VALIDATION_ERROR','SEMESTER_OUTSIDE_CYCLE');
  if((s.cycleId||null)!==(selected?cycleRef.id:null))revision(ref,{...s,cycleId:selected?cycleRef.id:null,associationDecision:{reason:x.p.reason||x.p.decisionBasis?.reason||'',decidedAt:M.stamp(x.now),actor:x.actor}});
 }
 return refs;
}
function f4ActivitiesAcademicCommands_(x){
 const {M,N,p,load,revision,event,now,actor}=x;
 function valid(){N(typeof p.label==='string'&&p.label.trim()&&p.label.length<=120&&M.gap(p.civilStart,p.civilEnd)>0,'VALIDATION_ERROR');N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'HUMAN_DECISION_REQUIRED');}
 function history(ref,action){event('history',{action,subjectRef:ref,ownerPersonRef:'',visibility:'BOARD',summary:p.reason,occurredAt:M.stamp(now)});}
 return {
  ACTIVITY_SEMESTER_SAVE(){
   valid();N(actor.kind==='BOARD_ACTOR'&&p.semesterRef.kind==='academicSemester','AUTHORIZATION_DENIED');const old=load(p.semesterRef,true,true);
   N(['PLANNED','ACTIVE','CLOSED'].includes(p.status),'VALIDATION_ERROR');
   if(old?.cycleId){const cycle=load(M.ref('cycle',old.cycleId));N(M.gap(cycle.civilStart,p.civilStart)>=0&&M.gap(p.civilEnd,cycle.civilEnd)>=0,'VALIDATION_ERROR','SEMESTER_OUTSIDE_CYCLE');}
   revision(p.semesterRef,{...old,label:p.label.trim(),civilStart:p.civilStart,civilEnd:p.civilEnd,status:p.status,cycleId:old?.cycleId||null,decision:{reason:p.reason,actor,decidedAt:M.stamp(now)}});
   history(p.semesterRef,old?'ACADEMIC_SEMESTER_EDITED':'ACADEMIC_SEMESTER_CREATED');return {semesterRef:p.semesterRef};
  },
  CYCLE_GOVERNANCE_EDIT(){
   valid();N(actor.kind==='BOARD_ACTOR','AUTHORIZATION_DENIED');const old=load(p.cycleRef,true);N(old.status!=='CLOSED'&&old.recordMode!=='HISTORICAL_PARTIAL','INVALID_STATE_TRANSITION','CLOSED_CYCLE_READ_ONLY');
   // A changed date must not silently invalidate existing presentation planning.
   N(old.status==='DRAFT'||old.civilStart===p.civilStart&&old.civilEnd===p.civilEnd,'INVALID_STATE_TRANSITION','ACTIVE_CYCLE_DATES_LOCKED');
   const next={...old,cycleLabel:p.label.trim(),civilStart:p.civilStart,civilEnd:p.civilEnd};
   next.academicSemesterRefs=f4ActivitiesAssignSemesters_(x,p.cycleRef,next,p.academicSemesterRefs,p.confirmAssociations);
   revision(p.cycleRef,next);history(p.cycleRef,'ACTIVITY_CYCLE_GOVERNANCE_EDITED');return {cycleRef:p.cycleRef};
  }
 };
}

// Frequency criteria belong to a versioned operational policy. Aggregation
// belongs to the activity cycle, never to a semester or legacy period.
function f4ActivityCycleContext_(get,cycleRef,semesterRef,civilDate){
 const M=f4ActivitiesModel_(),N=M.need;N(cycleRef?.kind==='cycle','CONFIG_CAPABILITY_DISABLED','ACTIVITY_CYCLE_REQUIRED');
 const c=get(M.path(cycleRef));N(c&&c.recordMode!=='HISTORICAL_PARTIAL','UNRESOLVED_DEPENDENCY','ACTIVITY_CYCLE_REQUIRED');
 N(M.gap(c.civilStart,civilDate)>=0&&M.gap(civilDate,c.civilEnd)>=0,'VALIDATION_ERROR','DATE_OUTSIDE_CYCLE');
 let academicSemesterId=null;
 if(semesterRef){N(semesterRef.kind==='academicSemester','VALIDATION_ERROR');const s=get(M.path(semesterRef));N(s?.cycleId===cycleRef.id&&(c.academicSemesterRefs||[]).some(r=>r.id===s.id),'CONFLICT','SEMESTER_CYCLE_MISMATCH');N(M.gap(s.civilStart,civilDate)>=0&&M.gap(civilDate,s.civilEnd)>=0,'VALIDATION_ERROR','DATE_OUTSIDE_SEMESTER');academicSemesterId=s.id;}
 return {cycleRef,cycleId:cycleRef.id,academicSemesterId};
}
function f4ActivityFrequencyPolicy_(policy){
 const N=f4ActivitiesModel_().need,b=policy?.frequencyBasis;
 N(policy?.approved&&b&&['ACTIVITY','MINUTE'].includes(b.unit)&&['EXCLUDE','COUNT_PRESENT','COUNT_ABSENT'].includes(b.justifiedMode)&&['EXCLUDE','COUNT_PRESENT','COUNT_ABSENT'].includes(b.excusedMode),'CONFIG_CAPABILITY_DISABLED','CANONICAL_FREQUENCY_POLICY_REQUIRED');
 return {unit:b.unit,justifiedMode:b.justifiedMode,excusedMode:b.excusedMode,policyRevision:policy.revision};
}
// New activity form: resolve context from canonical dates, never from a month
// split or identifier. Explicit references remain compatible with older clients.
function f4ActivityCreationContext_(get,type,value,now){
 const M=f4ActivitiesModel_(),N=M.need,day=value.dataAtividade,today=M.day(now);
 let cycleRef=value.cycleId?M.ref('cycle',value.cycleId):type.attendancePolicyRef?get(M.path(type.attendancePolicyRef))?.cycleRef:null;
 if(!cycleRef){const candidates=(type.governance?.usageAuthorizations||[]).filter(g=>{
  const c=get(M.path(g.cycleRef));return c&&c.status==='OPEN'&&c.civilStart<=today&&today<=c.civilEnd&&c.civilStart<=day&&day<=c.civilEnd&&f4OriginalModelUsage_(get,type,c.id);
 });N(candidates.length===1,'UNRESOLVED_DEPENDENCY',candidates.length?'AMBIGUOUS_CURRENT_CYCLE':'CURRENT_CYCLE_REQUIRED');cycleRef=candidates[0].cycleRef;}
 const cycle=get(M.path(cycleRef));N(cycle,'UNRESOLVED_DEPENDENCY');
 if(!value.cycleId)N(cycle.status==='OPEN'&&cycle.civilStart<=today&&today<=cycle.civilEnd,'UNRESOLVED_DEPENDENCY','CURRENT_CYCLE_REQUIRED');
 let semesterRef=value.academicSemesterId?M.ref('academicSemester',value.academicSemesterId):null;
 if(!semesterRef){const matches=(cycle.academicSemesterRefs||[]).map(r=>get(M.path(r))).filter(s=>s?.cycleId===cycle.id&&s.civilStart&&s.civilEnd&&s.civilStart<=day&&day<=s.civilEnd);N(matches.length<=1,'CONFLICT','AMBIGUOUS_ACADEMIC_SEMESTER');if(matches.length)semesterRef=M.ref('academicSemester',matches[0].id);}
 const context=f4ActivityCycleContext_(get,cycleRef,semesterRef,day),semester=semesterRef&&get(M.path(semesterRef));
 return {...context,academicSemesterRef:semesterRef,cycleLabel:cycle.cycleLabel||cycle.id,cycleRevision:cycle.revision,academicSemesterLabel:semester?.label||semester?.semesterLabel||semester?.id||null,academicSemesterRevision:semester?.revision||null};
}
function f4AttendanceCycleFields_(activity,collection){
 const M=f4ActivitiesModel_(),cycleId=activity.cycleId||activity.cycleRef?.id||collection.cycleId;
 M.need(M.id(cycleId),'UNRESOLVED_DEPENDENCY','ATTENDANCE_CYCLE_REQUIRED');
 return {cycleId,academicSemesterId:activity.academicSemesterId||null};
}
