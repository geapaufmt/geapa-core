// Exceptional recognition of a finite, reviewed import. No generic importer and
// no change to operational finalization. Route and protected approval expire.
function f4ImportedRecognitionHash_(value){
 const normalize=v=>v&&typeof v.toMillis==='function'?{_seconds:v.seconds,_nanoseconds:v.nanoseconds}:Array.isArray(v)?v.map(normalize):v&&typeof v==='object'?Object.fromEntries(Object.entries(v).map(([k,x])=>[k,normalize(x)])):v;
 return f4Core_().hash(JSON.parse(JSON.stringify(normalize(value))));
}
function f4ActivitiesImportRecognitionCommands_(x){
 return {ACTIVITY_IMPORT_RECOGNIZE(){
  const {M,N,p,actor,now,operationId,get,load,revision,event,prefetch,set}=x,H=f4ImportedRecognitionHash_;
  N(actor.kind==='BOARD_ACTOR'&&p.activityRef?.kind==='activity'&&p.recognitionRef?.kind==='policy','AUTHORIZATION_DENIED');
  N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'HUMAN_DECISION_REQUIRED');
  const approval=load(p.recognitionRef,true),intent=approval.intent;
  N(approval.usage==='ACTIVITY_IMPORT_RECOGNITION'&&approval.approved===true&&approval.active===true&&approval.actorPersonIds?.includes(actor.personId),'AUTHORIZATION_DENIED','RECOGNITION_NOT_APPROVED');
  N(M.ms(approval.issuedAt)<=now&&now<M.ms(approval.expiresAt),'AUTHORIZATION_DENIED','RECOGNITION_APPROVAL_EXPIRED');
  N(intent?.operationId===operationId&&intent.activityId===p.activityRef.id&&intent.cycleId==='GEAPA_2026'&&approval.intentHash===H(intent),'STALE_REVISION','RECOGNITION_INTENT_CHANGED');
  N(Array.isArray(intent.readSet)&&intent.readSet.length>0&&intent.readSet.length<=180&&intent.readSet.every(r=>(M.validPath(r.path)||/^member(Identities|Links)\/[A-Za-z0-9_-]{1,48}$/.test(r.path))&&r.path.split('/').length===2),'VALIDATION_ERROR');
  prefetch(intent.readSet.map(r=>r.path));
  const a=load(p.activityRef,true),criteria=approval.historicalCriteria;
  N(criteria?.schemaVersion==='import-historical-criteria-v1'&&criteria.historicalApprovalClaimed===false&&criteria.civilStart<=a.civilDate&&a.civilDate<=criteria.civilEnd&&criteria.basis&&criteria.sources?.length>0&&criteria.sources.length<=8,'UNRESOLVED_DEPENDENCY','HISTORICAL_CRITERIA_NOT_PROVEN');
  const policy=get('activityPolicies/'+intent.policyId);
  N(criteria.policyId===policy?.id&&criteria.policyRevision===policy.revision&&criteria.policyCriteriaHash===H(f4ActivityConfigNormativeFields_('attendancePolicy',policy)),'STALE_REVISION','HISTORICAL_POLICY_CHANGED');
  for(const source of criteria.sources){const path=M.path(source.ref);N(source.ref.kind==='evidence'&&intent.readSet.some(r=>r.path===path&&r.hash===source.hash)&&H(get(path))===source.hash,'UNRESOLVED_DEPENDENCY','HISTORICAL_SOURCE_NOT_SEALED');}
  // Planning 25/5 cannot replace this separate, field-level documentary basis.
  const plan=f4ImportedActivityReconstructionPlan_({intent,get,now,actorRef:M.ref('person',actor.personId)});
  N(approval.factsPreservedHash===plan.factsPreservedHash,'STALE_REVISION','FACTS_CHANGED');
  const refFor=path=>{const [col,id]=path.split('/'),kind=Object.keys(M.collections).find(k=>M.collections[k]===col);N(kind,'VALIDATION_ERROR');return M.ref(kind,id);};
  for(const w of plan.writes){
   const ref=refFor(w.path),old=get(w.path);
   // Keep the exact imported image, including its original provenance. Existing
   // immutable revision documents are checked, never replaced by a rollback.
   if(old&&(M.nested[w.path.split('/')[0]]||[]).includes('revisions')){
    const archive=w.path+'/revisions/'+old.revision,stored=get(archive);
    N(!stored||H(stored.snapshot)===H(old),'CONFLICT','IMPORT_ARCHIVE_CONFLICT');
    if(!stored)set(archive,{number:old.revision,snapshot:old,operationId,recordedAt:M.stamp(now),archiveBasis:'IMPORTED_BEFORE_IMAGE',notOriginalRevisionTimestamp:true});
   }
   revision(ref,w.data);
  }
  const fence=M.ref('attendanceRegistry','current'),old=get(M.path(fence));revision(fence,{epoch:(old?.epoch||0)+1});
  const capacity=M.ref('capacity','registry'),head=get(M.path(capacity));N(head&&Number.isSafeInteger(head.activityEpoch),'UNRESOLVED_DEPENDENCY','ACTIVITY_FENCE_REQUIRED');revision(capacity,{...head,activityEpoch:head.activityEpoch+1});
  event('history',{subjectRef:p.activityRef,action:plan.historyEvent,ownerPersonRef:'',visibility:'BOARD',summary:p.reason,occurredAt:M.stamp(now),recognitionRef:p.recognitionRef,intentHash:approval.intentHash,factsPreservedHash:plan.factsPreservedHash});
  return {activityRef:p.activityRef,recognitionRef:p.recognitionRef,manifestHash:plan.manifestHash,factsPreservedHash:plan.factsPreservedHash,attendanceCount:intent.factIds.length,pendingBasesReconciled:plan.writes.filter(w=>w.path.startsWith('activityJustifications/')).length};
 }};
}
