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
   valid();N(actor.kind==='BOARD_ACTOR','AUTHORIZATION_DENIED');const old=load(p.cycleRef,true);
   // A changed date must not silently invalidate existing presentation planning.
   N(old.status==='DRAFT'||old.civilStart===p.civilStart&&old.civilEnd===p.civilEnd,'INVALID_STATE_TRANSITION','ACTIVE_CYCLE_DATES_LOCKED');
   const next={...old,cycleLabel:p.label.trim(),civilStart:p.civilStart,civilEnd:p.civilEnd};
   next.academicSemesterRefs=f4ActivitiesAssignSemesters_(x,p.cycleRef,next,p.academicSemesterRefs,p.confirmAssociations);
   revision(p.cycleRef,next);history(p.cycleRef,'ACTIVITY_CYCLE_GOVERNANCE_EDITED');return {cycleRef:p.cycleRef};
  }
 };
}
