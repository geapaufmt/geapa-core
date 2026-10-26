// An approved justification remains a canonical human decision. J/A pairing is
// a frequency derivation over complete, proven facts of one person and cycle.
// Imported explicit J/A is immutable; semesters never partition this sequence.
function f4AttendancePairJustifications_(rows,complete){
 const M=f4ActivitiesModel_(),N=M.need;N(Array.isArray(rows)&&complete===true,'UNRESOLVED_DEPENDENCY','PAIRING_SOURCE_INCOMPLETE');
 const result=rows.map(r=>({...r})),groups=new Map(),seen=new Set();
 for(const r of result){
  N(M.id(r.id)&&!seen.has(r.id),'VALIDATION_ERROR','DUPLICATE_ATTENDANCE');seen.add(r.id);
  const b=r.officialCountBasis||r;
  if(!r.countable||b.countsAbsence!==true||b.applicableAtDate===false||!r.justificationApproved||!['JUSTIFIED','EXCUSED'].includes(r.status))continue;
  // Unknown historical cycle is retained, never assigned by date or semester.
  if(!M.id(r.cycleId)){N(!r.justificationPairingRule,'UNRESOLVED_DEPENDENCY','PAIRING_CYCLE_UNRESOLVED');continue;}
  N(M.id(r.personRef?.id),'UNRESOLVED_DEPENDENCY','PAIRING_PERSON_UNRESOLVED');
  N(typeof r.activityCivilDate==='string','UNRESOLVED_DEPENDENCY','PAIRING_DATE_UNRESOLVED');M.civil(r.activityCivilDate);
  const key=r.personRef.id+'|'+r.cycleId;if(!groups.has(key))groups.set(key,[]);groups.get(key).push(r);
 }
 for(const group of groups.values()){
  group.sort((a,b)=>a.activityCivilDate.localeCompare(b.activityCivilDate)||(a.activityTime||'').localeCompare(b.activityTime||'')||a.activityRef.id.localeCompare(b.activityRef.id)||a.id.localeCompare(b.id));
  group.forEach((r,index)=>{
   if(r.recordMode==='HISTORICAL_PARTIAL'||!r.justificationPairingRule)return;
   N(r.justificationPairingRule==='TWO_J_ONE_A_PERSON_CYCLE_V1','CONFIG_CAPABILITY_DISABLED','PAIRING_RULE_UNKNOWN');
   r.status=index%2===0?'JUSTIFIED':'EXCUSED';
   r.pairingBasis={rule:r.justificationPairingRule,cycleId:r.cycleId,ordinal:index+1,sourceStatus:rows.find(v=>v.id===r.id).status};
  });
 }
 return result;
}
