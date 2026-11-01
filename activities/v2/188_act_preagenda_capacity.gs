// Same capacity solver, bounded nominal demand from recognized schedules.
// This source deliberately never scans the member population or emits incidence.
function f4ActivitiesPreagendaCapacity_(options){
 const M=f4ActivitiesModel_(),N=M.need,db=options.db,clock=options.clock||Date.now,P=f4ActivitiesPreagenda_();
 function snapshot(cycleRef,configRef,change){
  const now=clock(),seen=new Map(),get=p=>{if(!seen.has(p))seen.set(p,db.get(p));return seen.get(p);};
  N(Object.keys(change).every(k=>['unitRef','sessionRef','sessionConfig','moves'].includes(k)),'HUMAN_DECISION_REQUIRED','PREAGENDA_DATE_CHANGE_REQUIRES_NEW_DECISION');
  const cycle=get(M.path(cycleRef)),cfg=get(M.path(configRef)),registry=get('presentationCapacityHeads/registry'),head=get('presentationCapacityHeads/'+cycleRef.id),ar=get('activityAttendanceRegistry/current');
  N(cycle&&cfg?.approved&&registry&&head&&ar&&M.path(cycle.currentConfigRef)===M.path(configRef),'UNRESOLVED_DEPENDENCY','CAPACITY_CONFIGURATION');
  P.parameters(cfg.parameters,cycleRef,get,now);const policy=P.policy(cfg.parameters,cycleRef,get,now);
  const until=Math.min(now+180000,M.ms(policy.document.validUntil),f4Core_().civilStart(M.day(now+86400000))),members=[],sessions=[],vector=[{ref:cycleRef,revision:cycle.revision},{ref:configRef,revision:cfg.revision},{ref:M.ref('capacity',cycleRef.id),revision:head.revision},{ref:M.ref('attendanceRegistry','current'),revision:ar.revision},{ref:policy.ref,revision:policy.document.revision}],model=f4MemberModel_(true),heads=[];
  for(const item of policy.document.items){
   const r=P.refs(item.activityRef,item.presentationRef),u=get(M.path(r.unitRef));N(u,'HUMAN_DECISION_REQUIRED','PREAGENDA_RECOGNITION_REQUIRED');
   const b=P.basis(u,get,now),a=b.activity,s=a.presentationSessionRef&&get(M.path(a.presentationSessionRef));
   N(s&&s.activityRef?.id===a.id&&s.cycleRef?.id===cycleRef.id&&s.civilDate===a.civilDate&&[1,2].includes(s.config?.capacity)&&Array.isArray(s.slots)&&s.slots.length<=s.config.capacity,'UNRESOLVED_DEPENDENCY','PREAGENDA_SESSION_REQUIRED');
   N(M.gap(cycle.civilStart,a.civilDate)>=0&&M.gap(a.civilDate,cycle.civilEnd)>=0,'VALIDATION_ERROR','DATE_OUTSIDE_CYCLE');
   for(const slot of s.slots){const su=get(M.path(slot.unitRef)),assignment=get(M.path(slot.assignmentRef));N(su&&assignment&&['ACTIVE','CONFIRMED'].includes(assignment.status)&&assignment.unitRef?.id===su.id&&assignment.sessionRef?.id===s.id&&su.assignmentRef?.id===assignment.id&&policy.document.items.some(i=>P.refs(i.activityRef,i.presentationRef).unitRef.id===su.id),'UNRESOLVED_DEPENDENCY','PREAGENDA_SLOT_DRIFT');}
   const h=get('memberAggregateHeads/'+u.personRef.id);N(h?.kind==='PERSON'&&Array.isArray(h.paths)&&h.paths.length<=180,'UNRESOLVED_DEPENDENCY','MEMBERS_HEAD');
   const paths=h.paths.filter(p=>/^member(Identities|Links|Events)\//.test(p));N(paths.includes('memberIdentities/'+u.personRef.id),'UNRESOLVED_DEPENDENCY','MEMBERS_SOURCES');
   const docs=Object.fromEntries(paths.map(p=>[p,get(p)]));N(Object.values(docs).every(d=>d&&d.idPessoa===u.personRef.id),'UNRESOLVED_DEPENDENCY','MEMBERS_SOURCE_MISSING');
   const current=model.derive(docs,u.personRef.id,h.revision,now).state,atDate=model.derive(docs,u.personRef.id,h.revision,M.ms(a.startsAt)).state;
   N(current.state==='MEMBRO_EFETIVO'&&!current.blocked&&atDate.state==='MEMBRO_EFETIVO'&&!atDate.blocked,'HUMAN_DECISION_REQUIRED','PREAGENDA_MEMBERSHIP_INELIGIBLE');
   const assignment=u.assignmentRef&&get(M.path(u.assignmentRef)),assigned=!!(assignment&&['ACTIVE','CONFIRMED'].includes(assignment.status)&&assignment.sessionRef?.id===s.id&&s.slots.some(v=>v.unitRef.id===u.id&&v.assignmentRef.id===assignment.id));
   const available=M.gap(M.day(now),s.civilDate)>0&&(M.gap(M.day(now),s.civilDate)>=(cfg.parameters.newDateMinimumDays??30)||P.deadline({get,now},s,'CALENDAR',u));
   members.push({personId:u.personRef.id,demandKind:'RECOGNIZED_PREAGENDA',assigned,eligibleSessionIds:available||assigned?[s.id]:[],validUntil:until});
   N(!sessions.some(v=>v.ref.id===s.id),'UNRESOLVED_DEPENDENCY','PREAGENDA_SESSION_AMBIGUOUS');
   sessions.push({ref:M.ref('session',s.id),revision:s.revision,availableForNew:available,admittedCapacity:cfg.parameters.shared?s.config.capacity:1,allowedPersonId:u.personRef.id,validUntil:until});
   vector.push({ref:r.unitRef,revision:u.revision},{ref:r.basisRef,revision:b.basis.revision});heads.push({path:'memberAggregateHeads/'+u.personRef.id,revision:h.revision});
  }
  N(new Set(members.map(m=>m.personId)).size===members.length,'UNRESOLVED_DEPENDENCY','MULTIPLE_PREAGENDA_DEMANDS_REQUIRE_SEPARATE_ASSESSMENT');
  const entries=[...seen],fresh=db.getMany(entries.map(([p])=>p));entries.forEach(([,before],i)=>N((fresh[i]?.revision??null)===(before?.revision??null),'STALE_REVISION','PREAGENDA_SOURCE_CHANGED'));
  N(clock()<until,'STALE_REVISION','CAPACITY_SOURCE_EXPIRED');
  return {complete:true,membersEpoch:registry.membersEpoch,cycleRegistryEpoch:registry.cycleRegistryEpoch,sourceRevisionVector:vector,sourceRegistryDigest:M.hash({heads,policy:policy.ref,revision:policy.document.revision}),members,sessions,validUntil:until,preservesAcquiredRights:true};
 }
 return {snapshot};
}
