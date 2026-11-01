function f4ActivitiesIncidenceProof_(x,ref){
 const a=x.load(ref);x.N(a.producer==='CORE_INCIDENCE_SOURCE'&&a.state==='SEALED'&&a.complete===true&&a.cycleRef.id===x.cycleRef.id&&a.validUntil&&x.M.ms(a.validUntil)>x.now,'UNRESOLVED_DEPENDENCY','INCIDENCE_PROOF_REQUIRED');
 x.N(a.digest===f4ActivitiesCapacityDigest_(a.items),'UNRESOLVED_DEPENDENCY','INCIDENCE_PROOF_DIGEST');
 for(const b of a.basisRevisionVector)x.N(x.load(b.ref).revision===b.revision,'STALE_REVISION','INCIDENCE_SOURCE_CHANGED');
 return a;
}
function f4ActivitiesIncidenceEvaluate_(x,a){
 const {M,N,p,load,revision,save,actor}=x;N(actor.kind==='BOARD_ACTOR'&&typeof p.decisionReason==='string'&&p.decisionReason.trim(),'HUMAN_DECISION_REQUIRED');
 x.config();a=f4ActivitiesIncidenceProof_(x,p.capacityAssessmentRef);N(a.items.length===1,'VALIDATION_ERROR');const item=a.items[0];
 N(item.personRef.id===p.personRef.id&&p.membersEvidenceRef.id===p.personRef.id&&x.get('memberAggregateHeads/'+item.personRef.id)?.revision===item.personRevision,'STALE_REVISION','INCIDENCE_PERSON_CHANGED');
 N(['INCIDENT','DEFERRED','NOT_INCIDENT'].includes(item.disposition),'HUMAN_DECISION_REQUIRED','N03_RESIDUAL');
 const key=M.hash({person:p.personRef.id,cycle:p.cycleRef.id,kind:'ORDINARY',version:1}).slice(0,32),claim=M.ref('claim',key),old=load(claim,false,true);
 if(old)return {incidence:'INCIDENT',obligationRef:old.targetRef,replay:true};
 if(item.disposition!=='INCIDENT')return {incidence:item.disposition,assessmentRef:p.capacityAssessmentRef};
 N(item.membershipRef?.kind==='membership'&&item.effectiveCivilDate&&item.eligibleSessionRefs?.length,'UNRESOLVED_DEPENDENCY','INCIDENCE_HISTORICAL_BASIS');
 const ob=M.ref('obligation',key),unit=M.ref('unit',key);N(!load(ob,true,true)&&!load(unit,false,true),'CONFLICT');
 save(claim,{claimKind:'ORDINARY_OBLIGATION',keyVersion:1,canonicalKeyComponents:{person:p.personRef.id,cycle:p.cycleRef.id},targetRef:ob});
 revision(ob,{personRef:p.personRef,cycleRef:p.cycleRef,unitRef:unit,membershipFactRef:item.membershipRef,membershipRevision:item.personRevision,effectiveCivilDate:item.effectiveCivilDate,incidenceAssessmentRef:p.capacityAssessmentRef,incidenceDecision:{reason:p.decisionReason,actor,authorityBasis:x.authorityBasis},incidenceState:'INCIDENT',result:'PENDING'});
 save(unit,{personRef:p.personRef,cycleRef:p.cycleRef,obligationRef:ob});x.bumpCapacity();return {incidence:'INCIDENT',obligationRef:ob,unitRef:unit};
}
function f4ActivitiesIncidenceCommands_(x){return {
 INCIDENCE_COVERAGE_APPROVE(){
  const {M,N,p,load,save,actor}=x;N(actor.kind==='BOARD_ACTOR'&&typeof p.reason==='string'&&p.reason.trim(),'HUMAN_DECISION_REQUIRED');x.config();
  const proof=f4ActivitiesIncidenceProof_(x,p.sourceManifestRef);N(proof.coverageReady===true&&proof.items.every(i=>['INCIDENT','DEFERRED','NOT_INCIDENT'].includes(i.disposition)&&(i.disposition!=='INCIDENT'||i.obligationRef)),'UNRESOLVED_DEPENDENCY','INCIDENCE_EVALUATION_REQUIRED');
  const c=load(p.cycleRef,true),headRef=M.ref('capacity',c.id),head=load(headRef,true),g=load(M.ref('capacity','registry'),true);
  const ref=x.event('assessment',{coverageKind:'act-incidence-coverage-v1',cycleRef:p.cycleRef,configRef:c.currentConfigRef,state:'SEALED',complete:true,approved:true,items:proof.items,digest:proof.digest,membersEpoch:g.membersEpoch,cycleRegistryEpoch:g.cycleRegistryEpoch,sourceManifestRef:p.sourceManifestRef,sourceHeadsDigest:proof.sourceHeadsDigest,decisionBasis:{reason:p.reason,actor,authorityBasis:x.authorityBasis},validUntil:M.stamp(f4Core_().civilStart(f4Core_().nextDay(c.civilEnd)))});
  save(headRef,{...head,incidenceCoverageRef:ref,capacityEpoch:head.capacityEpoch+1,state:'INVALIDATED'});return {coverageRef:ref,coveredCount:proof.items.length};
 }
};}
