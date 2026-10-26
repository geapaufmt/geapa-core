// Incidence consumes Members history. Projection timestamps never stand in for
// the date of effectivation. Residual cases remain explicit for human review.
function f4ActivitiesIncidenceSource_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=options.db,clock=options.clock||Date.now;
 function all(collection,filters=[]){let cursor;const rows=[];for(let page=0;page<10;page++){const part=db.query({collection,filters,order:[{field:'__name__',direction:'ASCENDING'}],limit:20,...(cursor?{cursor:[cursor]}:{})});rows.push(...part);N(rows.length<=180,'SCOPE_TOO_LARGE','INCIDENCE_POPULATION_BOUND');if(part.length<20)return rows;cursor=new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+part[part.length-1].path);}N(false,'SCOPE_TOO_LARGE');}
 function scan(cycleRef){
  const now=clock(),cycle=db.get(M.path(cycleRef)),registry=db.get('presentationCapacityHeads/registry'),head=db.get('presentationCapacityHeads/'+cycleRef.id),activityRegistry=db.get('activityAttendanceRegistry/current');
  N(cycle&&registry&&head&&activityRegistry,'UNRESOLVED_DEPENDENCY');
  const config=cycle.currentConfigRef&&db.get(M.path(cycle.currentConfigRef)),policy=db.get(M.path(cycle.policyVersionRef));
  N(config?.approved&&policy?.approved&&policy.incidenceRule==='ORDINARY_FIRST_EFFECTIVATION','CONFIG_CAPABILITY_DISABLED','INCIDENCE_RULE_REQUIRED');
  const lead=config.parameters.newDateMinimumDays??30;N(Number.isSafeInteger(lead)&&lead>=30,'CONFIG_CAPABILITY_DISABLED');
  const sessions=all('presentationSessions',[{field:'cycleRef',op:'EQUAL',value:cycleRef}]).map(r=>r.data);N(sessions.length<=32,'SCOPE_TOO_LARGE');
  const activities=new Map();for(const s of sessions){const a=s.activityRef&&db.get(M.path(s.activityRef));N(a&&a.civilDate===s.civilDate,'UNRESOLVED_DEPENDENCY','SESSION_ACTIVITY_REQUIRED');activities.set(s.id,a);}
  const obligations=all('presentationObligations',[{field:'cycleRef',op:'EQUAL',value:cycleRef}]).map(r=>r.data),byPerson=new Map();
  for(const o of obligations){N(!byPerson.has(o.personRef.id),'CONFLICT','DUPLICATE_ORDINARY_OBLIGATION');byPerson.set(o.personRef.id,o);}
  const model=f4MemberModel_(true),items=[],heads=[];
  for(const row of all('memberAggregateHeads')){
   const h=row.data;N(h.kind==='PERSON'&&h.id===row.path.split('/')[1]&&Array.isArray(h.paths),'UNRESOLVED_DEPENDENCY','MEMBERS_HEAD');heads.push({path:row.path,revision:h.revision});
   const paths=h.paths.filter(p=>/^member(Identities|Links|Events)\//.test(p));N(paths.length<=180&&paths.includes('memberIdentities/'+h.id),'UNRESOLVED_DEPENDENCY');
   const values=db.getMany(paths);N(values.every(v=>v&&v.idPessoa===h.id),'UNRESOLVED_DEPENDENCY','MEMBERS_SOURCE_MISSING');const docs=Object.fromEntries(paths.map((p,i)=>[p,values[i]])),person=docs['memberIdentities/'+h.id];
   if(person.technicalAccount)continue;
   const state=model.derive(docs,h.id,h.revision,now).state,existing=byPerson.get(h.id),links=model.rows(docs,'memberLinks').filter(v=>v.type==='MEMBRO_EFETIVO'),events=model.rows(docs,'memberEvents');
   const item={personRef:M.ref('person',h.id),personRevision:h.revision,disposition:'NOT_INCIDENT',reasonCode:'NO_EFFECTIVE_MEMBERSHIP'};
   if(existing){N(existing.incidenceState==='INCIDENT','HUMAN_DECISION_REQUIRED','INCIDENCE_EXISTING_REVIEW');item.disposition='INCIDENT';item.reasonCode='EXISTING_OBLIGATION_PRESERVED';item.obligationRef=M.ref('obligation',existing.id);if(existing.effectiveCivilDate)item.effectiveCivilDate=existing.effectiveCivilDate;if(existing.membershipFactRef?.kind==='membership')item.membershipRef=existing.membershipFactRef;byPerson.delete(h.id);}
   else if(links.length>1||events.some(e=>['RETORNO','RETIFICACAO'].includes(e.type))){item.disposition='REVIEW_REQUIRED';item.reasonCode='N03_RESIDUAL';}
   else if(links.length===1){
    const link=links[0],effective=C.civilStart(link.startDate);item.membershipRef=M.ref('membership',link.id);item.effectiveCivilDate=link.startDate;
    if(effective>now){item.reasonCode='EFFECTIVATION_NOT_DUE';}
    else if(state.state!=='MEMBRO_EFETIVO'||state.blocked){item.disposition='REVIEW_REQUIRED';item.reasonCode='MEMBERS_TRANSITION_REVIEW';}
    else{
     const atEffect=model.derive(docs,h.id,h.revision,effective).state;N(atEffect.state==='MEMBRO_EFETIVO'&&!atEffect.blocked,'UNRESOLVED_DEPENDENCY','EFFECTIVATION_PROOF');
     const eligible=sessions.filter(s=>{const a=activities.get(s.id),at=model.derive(docs,h.id,h.revision,M.ms(a.startsAt)).state;return a.status==='SCHEDULED'&&s.config.destination==='ORDINARY'&&M.gap(link.startDate,s.civilDate)>=lead&&M.gap(cycle.civilStart,s.civilDate)>=0&&M.gap(s.civilDate,cycle.civilEnd)>=0&&at.state==='MEMBRO_EFETIVO'&&!at.blocked;});
     // A later scan of an open cycle cannot reconstruct its historical planning
     // merely from today's sessions. Existing decisions remain authoritative.
     const knownAtEffect=s=>{const a=activities.get(s.id);return Number.isFinite(M.ms(s.createdAt))&&M.ms(s.createdAt)<=effective&&Number.isFinite(M.ms(s.updatedAt))&&M.ms(s.updatedAt)<=effective&&Number.isFinite(M.ms(a.updatedAt))&&M.ms(a.updatedAt)<=effective;};
     const historical=cycle.status==='DRAFT'?eligible:eligible.filter(knownAtEffect),canProveAbsence=cycle.status==='DRAFT'||cycle.openedAt&&M.ms(cycle.openedAt)<=effective&&sessions.every(knownAtEffect);
     if(cycle.status!=='DRAFT'&&(!cycle.openedAt||M.ms(cycle.openedAt)>effective||!historical.length&&!canProveAbsence)){item.disposition='REVIEW_REQUIRED';item.reasonCode='HISTORICAL_PLANNING_PROOF_REQUIRED';}
     else{item.disposition=historical.length?'INCIDENT':'DEFERRED';item.reasonCode=historical.length?'ELIGIBLE_DATE_AT_EFFECTIVATION':'NO_ELIGIBLE_DATE_AT_EFFECTIVATION';item.eligibleSessionRefs=historical.map(s=>M.ref('session',s.id));}
    }
   }
   items.push(item);
  }
  N(byPerson.size===0,'UNRESOLVED_DEPENDENCY','OBLIGATION_PERSON_MISSING');
  const basis=[{ref:cycleRef,revision:cycle.revision},{ref:cycle.currentConfigRef,revision:config.revision},{ref:cycle.policyVersionRef,revision:policy.revision},{ref:M.ref('capacity','registry'),revision:registry.revision},{ref:M.ref('capacity',cycle.id),revision:head.revision},{ref:M.ref('attendanceRegistry','current'),revision:activityRegistry.revision}];
  for(const b of basis)N(db.get(M.path(b.ref))?.revision===b.revision,'STALE_REVISION','INCIDENCE_SOURCE_CHANGED');const fresh=db.getMany(heads.map(h=>h.path));heads.forEach((h,i)=>N(fresh[i]?.revision===h.revision,'STALE_REVISION','MEMBERS_CHANGED'));
  N(clock()<Math.min(now+60000,C.civilStart(C.nextDay(M.day(now)))),'STALE_REVISION','INCIDENCE_SOURCE_EXPIRED');
  return {cycle,config,registry,head,basis,items,headsDigest:f4ActivitiesCapacityDigest_(heads),validUntil:Math.min(now+60000,C.civilStart(C.nextDay(M.day(now))))};
 }
 return {scan};
}
