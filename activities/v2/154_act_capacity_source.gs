// Canonical capacity source. The approved incidence coverage is an ACT planning
// fact containing references, not a second member registry or an identity copy.
function f4ActivitiesCapacityDigest_(rows){
 const M=f4ActivitiesModel_();M.need(Array.isArray(rows)&&rows.length<=180,'SCOPE_TOO_LARGE','CAPACITY_SOURCE_BOUND');
 if(rows.length<=64)return M.hash(rows);
 const chunks=[];for(let i=0;i<rows.length;i+=64)chunks.push(M.hash(rows.slice(i,i+64)));
 return M.hash({count:rows.length,chunks});
}
function f4ActivitiesCapacitySource_(options){
 const M=f4ActivitiesModel_(),N=M.need,F=f4Firestore_(),db=options.db,clock=options.clock||Date.now;
 function snapshot(cycleRef,configRef,change){
  N(change&&Object.keys(change).every(k=>['unitRef','sessionRef','sessionConfig','sessionDate','moves'].includes(k)),'VALIDATION_ERROR','CAPACITY_CHANGE_SCOPE');
  const start=clock(),registry=db.get('presentationCapacityHeads/registry'),head=db.get('presentationCapacityHeads/'+cycleRef.id),cycle=db.get(M.path(cycleRef)),cfg=db.get(M.path(configRef)),activityRegistry=db.get('activityAttendanceRegistry/current');
  N(registry&&head&&cycle&&activityRegistry&&cfg?.approved&&M.path(cycle.currentConfigRef)===M.path(configRef),'UNRESOLVED_DEPENDENCY','CAPACITY_CONFIGURATION');
  const coverageRef=head.incidenceCoverageRef,coverage=coverageRef&&db.get(M.path(coverageRef));
  N(coverageRef?.kind==='assessment'&&(coverage?.schemaVersion==='act-incidence-coverage-v1'||coverage?.coverageKind==='act-incidence-coverage-v1')&&coverage.state==='SEALED'&&coverage.complete===true&&coverage.cycleRef?.id===cycleRef.id&&coverage.configRef&&M.path(coverage.configRef)===M.path(configRef)&&coverage.membersEpoch===registry.membersEpoch&&coverage.cycleRegistryEpoch===registry.cycleRegistryEpoch&&Array.isArray(coverage.items)&&coverage.items.length<=180&&coverage.digest===f4ActivitiesCapacityDigest_(coverage.items),'UNRESOLVED_DEPENDENCY','INCIDENCE_COVERAGE_REQUIRED');
  N(coverage.approved===true&&coverage.decisionBasis?.reason&&coverage.validUntil&&M.ms(coverage.validUntil)>start,'HUMAN_DECISION_REQUIRED','INCIDENCE_COVERAGE_UNAPPROVED');
  const lead=cfg.parameters.newDateMinimumDays??30;N(Number.isSafeInteger(lead)&&lead>=30,'CONFIG_CAPABILITY_DISABLED','CALENDAR_ADVANCE');
  const seen=new Map();function get(path){if(!seen.has(path))seen.set(path,db.get(path));return seen.get(path)}
  function all(collection,filters=[]){let cursor,rows=[];for(let page=0;page<10;page++){const part=db.query({collection,filters,limit:20,order:[{field:'__name__',direction:'ASCENDING'}],...(cursor?{cursor:[cursor]}:{})});rows.push(...part);if(part.length<20){N(rows.length<=180,'SCOPE_TOO_LARGE','CAPACITY_SOURCE_BOUND');return rows;}cursor=new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+part[part.length-1].path);}N(false,'SCOPE_TOO_LARGE','CAPACITY_SOURCE_BOUND');}
  const dateChange=change.sessionDate;if(dateChange){f4Core_().exact(change,['sessionDate']);f4Core_().exact(dateChange,['sessionRef','activityRef','startsAt','endsAt','civilDate']);N(dateChange.sessionRef?.kind==='session'&&dateChange.activityRef?.kind==='activity'&&Number.isSafeInteger(dateChange.startsAt)&&Number.isSafeInteger(dateChange.endsAt)&&dateChange.endsAt>dateChange.startsAt&&M.day(dateChange.startsAt)===dateChange.civilDate,'VALIDATION_ERROR','SESSION_DATE_CHANGE');}
  const sessions=all('presentationSessions',[{field:'cycleRef',op:'EQUAL',value:cycleRef}]).map(r=>r.data).filter(s=>s.cycleRef?.id===cycleRef.id).map(s=>dateChange?.sessionRef.id===s.id?{...s,civilDate:dateChange.civilDate}:s);N(sessions.length<=32,'SCOPE_TOO_LARGE','CAPACITY_SESSION_BOUND');if(dateChange)N(sessions.some(s=>s.id===dateChange.sessionRef.id&&s.activityRef?.id===dateChange.activityRef.id),'UNRESOLVED_DEPENDENCY','SESSION_ACTIVITY_REQUIRED');
  const covered=new Map();for(const item of coverage.items){N(item.personRef?.kind==='person'&&M.id(item.personRef.id)&&!covered.has(item.personRef.id)&&['INCIDENT','DEFERRED','NOT_INCIDENT'].includes(item.disposition),'UNRESOLVED_DEPENDENCY','INCIDENCE_COVERAGE_INVALID');covered.set(item.personRef.id,item);}
  for(const {data:o} of all('presentationObligations',[{field:'cycleRef',op:'EQUAL',value:cycleRef}]))if(o.cycleRef?.id===cycleRef.id&&o.incidenceState==='INCIDENT')N(covered.get(o.personRef.id)?.obligationRef?.id===o.id,'UNRESOLVED_DEPENDENCY','INCIDENCE_OBLIGATION_COVERAGE');
  const model=f4MemberModel_(true),members=[],sourceHeads=[];let validUntil=Math.min(M.ms(coverage.validUntil),start+60000,f4Core_().civilStart(M.day(start+86400000)));
  const eligibleSessions=new Map();let preservesAcquiredRights=true;
  for(const s of sessions){
   N(Number.isSafeInteger(s.revision)&&Array.isArray(s.slots)&&[1,2].includes(s.config?.capacity)&&s.slots.length<=s.config.capacity&&new Set(s.slots.map(v=>v.unitRef?.id)).size===s.slots.length&&new Set(s.slots.map(v=>v.assignmentRef?.id)).size===s.slots.length,'UNRESOLVED_DEPENDENCY','SESSION_CAPACITY_DRIFT');
   let activity=s.activityRef&&get(M.path(s.activityRef));if(dateChange?.sessionRef.id===s.id){N(activity?.presentationSessionRef?.id===s.id,'UNRESOLVED_DEPENDENCY','SESSION_ACTIVITY_REQUIRED');activity={...activity,civilDate:dateChange.civilDate,startsAt:M.stamp(dateChange.startsAt),endsAt:M.stamp(dateChange.endsAt)};}
   N(s.activityRef?.kind==='activity'&&activity&&activity.civilDate===s.civilDate,'UNRESOLVED_DEPENDENCY','SESSION_ACTIVITY_REQUIRED');
   // A cancelled/occurred activity cannot supply a future place. Explicit
   // session destinations still protect reservations and exclusivity.
   const usable=activity.status==='SCHEDULED'&&M.gap(M.day(start),s.civilDate)>=lead&&M.gap(cycle.civilStart,s.civilDate)>=0&&M.gap(s.civilDate,cycle.civilEnd)>=0;
   eligibleSessions.set(s.id,{s,activity,usable});
   for(const slot of s.slots){const assignment=get(M.path(slot.assignmentRef)),unit=get(M.path(slot.unitRef));N(assignment&&unit&&['ACTIVE','CONFIRMED'].includes(assignment.status)&&assignment.sessionRef.id===s.id&&assignment.unitRef.id===unit.id&&unit.assignmentRef?.id===assignment.id&&unit.cycleRef.id===cycleRef.id,'UNRESOLVED_DEPENDENCY','ASSIGNMENT_SLOT_DRIFT');}
  }
  for(const row of all('memberAggregateHeads')){
   const h=row.data;N(h.kind==='PERSON'&&h.id===row.path.split('/')[1]&&Array.isArray(h.paths),'UNRESOLVED_DEPENDENCY','MEMBERS_HEAD');sourceHeads.push({path:row.path,revision:h.revision});
   const paths=h.paths.filter(p=>/^member(Identities|Links|Events)\//.test(p));N(paths.includes('memberIdentities/'+h.id)&&paths.length<=180,'UNRESOLVED_DEPENDENCY','MEMBERS_SOURCES');
   const values=db.getMany(paths),docs=Object.fromEntries(paths.map((p,i)=>[p,values[i]]));N(values.every(v=>v&&v.idPessoa===h.id),'UNRESOLVED_DEPENDENCY','MEMBERS_SOURCE_MISSING');if(docs['memberIdentities/'+h.id].technicalAccount)continue;
   const current=model.derive(docs,h.id,h.revision,start).state,item=covered.get(h.id);
   // Observed membership never silently replaces approved ACT incidence.
   N(current.state!=='MEMBRO_EFETIVO'||item,'UNRESOLVED_DEPENDENCY','INCIDENCE_SYNC_REQUIRED');if(!item)continue;
   N(item.personRevision===h.revision,'STALE_REVISION','INCIDENCE_MEMBERS_CHANGED');covered.delete(h.id);
   if(item.disposition!=='INCIDENT'){N(!(coverage.coverageKind==='act-incidence-coverage-v1'&&item.reasonCode==='EFFECTIVATION_NOT_DUE'&&current.state==='MEMBRO_EFETIVO'),'UNRESOLVED_DEPENDENCY','INCIDENCE_SYNC_REQUIRED');continue;}
   const obligation=item.obligationRef&&get(M.path(item.obligationRef));N(item.obligationRef?.kind==='obligation'&&obligation?.personRef.id===h.id&&obligation.cycleRef.id===cycleRef.id&&obligation.incidenceState==='INCIDENT','UNRESOLVED_DEPENDENCY','INCIDENCE_OBLIGATION');
   const unit=get(M.path(obligation.unitRef)),presentation=unit?.presentationRef&&get(M.path(unit.presentationRef)),realization=presentation?.realizationRef&&get(M.path(presentation.realizationRef));N(unit?.personRef.id===h.id&&unit.obligationRef?.id===obligation.id,'UNRESOLVED_DEPENDENCY','INCIDENCE_UNIT');
   if(realization?.recognized===true)continue;
   const eligibleSessionIds=[];
   for(const {s,activity,usable} of eligibleSessions.values())if(usable){const state=model.derive(docs,h.id,h.revision,M.ms(activity.startsAt)).state;if(state.state==='MEMBRO_EFETIVO'&&!state.blocked)eligibleSessionIds.push(s.id);}
   const assignment=unit.assignmentRef&&get(M.path(unit.assignmentRef)),assignedSession=assignment&&eligibleSessions.get(assignment.sessionRef.id);
   // An existing assignment with less than D-30 remains valid; D-30 applies
   // when confirming a new date, not to an already acquired place.
   const assignedState=assignedSession&&model.derive(docs,h.id,h.revision,M.ms(assignedSession.activity.startsAt)).state;
   if(dateChange&&assignment?.sessionRef.id===dateChange.sessionRef.id)preservesAcquiredRights=preservesAcquiredRights&&assignedState?.state==='MEMBRO_EFETIVO'&&!assignedState.blocked;
   const assigned=!!(assignment&&['ACTIVE','CONFIRMED'].includes(assignment.status)&&assignedSession&&assignedSession.activity.status==='SCHEDULED'&&M.gap(M.day(start),assignedSession.s.civilDate)>=0&&assignedSession.s.slots.some(s=>s.unitRef.id===unit.id&&s.assignmentRef.id===assignment.id)&&assignedState.state==='MEMBRO_EFETIVO'&&!assignedState.blocked);
   members.push({personId:h.id,incident:true,assigned,eligibleSessionIds,validUntil});
  }
  N(covered.size===0,'UNRESOLVED_DEPENDENCY','INCIDENCE_PERSON_MISSING');if(dateChange)for(const slot of sessions.find(s=>s.id===dateChange.sessionRef.id).slots){const unit=get(M.path(slot.unitRef));N(members.some(m=>m.personId===unit.personRef.id),'UNRESOLVED_DEPENDENCY','SESSION_PARTICIPANT_COVERAGE');}
  // Recheck global fences and all observed heads: a partial or mixed scan is
  // never stamped complete. Every canonical Members mutation advances registry.
  N(db.get('presentationCapacityHeads/registry')?.revision===registry.revision&&db.get('presentationCapacityHeads/'+cycleRef.id)?.revision===head.revision&&(db.get('activityAttendanceRegistry/current')?.revision||0)===(activityRegistry?.revision||0),'STALE_REVISION','CAPACITY_SOURCE_CHANGED');
  const fresh=db.getMany(sourceHeads.map(h=>h.path));sourceHeads.forEach((h,i)=>N(fresh[i]?.revision===h.revision,'STALE_REVISION','MEMBERS_CHANGED'));
  N(clock()<validUntil,'STALE_REVISION','CAPACITY_SOURCE_EXPIRED');
  const vector=[{ref:cycleRef,revision:cycle.revision},{ref:configRef,revision:cfg.revision},{ref:M.ref('capacity',cycleRef.id),revision:head.revision},{ref:coverageRef,revision:coverage.revision},...(activityRegistry?[{ref:M.ref('attendanceRegistry','current'),revision:activityRegistry.revision}]:[])];
  return {complete:true,membersEpoch:registry.membersEpoch,cycleRegistryEpoch:registry.cycleRegistryEpoch,sourceRegistryDigest:M.hash({sourceHeadsDigest:f4ActivitiesCapacityDigest_(sourceHeads),coverageRef,coverageRevision:coverage.revision,capacityRevision:head.revision,activityRevision:activityRegistry?.revision||0}),sourceRevisionVector:vector,members,sessions:sessions.map(s=>({ref:M.ref('session',s.id),revision:s.revision,availableForNew:eligibleSessions.get(s.id).usable,admittedCapacity:cfg.parameters.shared?s.config.capacity:1,allowedPersonId:s.config.destination==='EXCLUSIVE'&&s.exclusiveUnitRef?get(M.path(s.exclusiveUnitRef))?.personRef?.id||'UNRESOLVED_EXCLUSIVE':'',validUntil})),validUntil,preservesAcquiredRights};
 }
 return {snapshot};
}
