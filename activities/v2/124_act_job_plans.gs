// Resolves bounded job payloads from canonical heads and approved proof references.
// Missing proof references leave the job pending; no job creates a human decision.
function f4ActivitiesJobPlans_(services){
 const M=f4ActivitiesModel_(),N=M.need,get=r=>{const d=services.db.get(M.path(r));N(d,'UNRESOLVED_DEPENDENCY','JOB_SOURCE');return d;};
 function prepare(kind,item,scope){const r=item.ref,d=get(r),basis=d.automationBasis||{},cycleRef=d.cycleRef||basis.cycleRef||(r.kind==='cycle'?r:null),ruleRef=(N(cycleRef,'UNRESOLVED_DEPENDENCY','JOB_CYCLE'),services.ruleFor(kind,cycleRef)),rule=get(ruleRef);N(rule.approved&&rule.active,'UNRESOLVED_DEPENDENCY','JOB_RULE');let p;
  if(kind==='THEMATIC_TACIT')p={requestRef:r,basisRevision:d.revision,dueAt:M.ms(d.analysisClock.dueAt),clockPolicyRef:d.analysisClock.policyRef,jobKey:M.hash({kind,id:r.id,revision:d.revision})};
  else if(kind==='CONFIG_ACTIVATE')p={cycleRef:r,approvedConfigRef:d.scheduledConfigRef,impactAssessmentRef:basis.impactAssessmentRef};
  else if(kind==='OBLIGATION_EVALUATE'){const u=get(d.unitRef),pr=get(u.presentationRef),m=get(pr.materialRequirementRef);p={obligationRef:r,cutoffRef:basis.cutoffRef,materialProofRef:m.currentRegularizationRef,realizationRef:pr.realizationRef};}
  else if(kind==='MATERIAL_EVALUATE')p={presentationRef:d.presentationRef,requirementRevisionRef:d.currentRevisionRef,receiptCoverageManifestRef:basis.receiptCoverageManifestRef,cutoffRef:basis.cutoffRef};
  else if(kind==='CORRECTION_EXPIRE')p={opportunityRef:r,basisRevision:d.revision,dueAt:M.ms(d.dueAt),jobKey:M.hash({kind,id:r.id,revision:d.revision})};
  else if(kind==='PROJECTION_REBUILD')p={viewFamily:basis.viewFamily,phase:basis.phase,generationRef:basis.generationRef,basisManifestRef:basis.basisManifestRef,...(basis.ownerPersonRef?{ownerPersonRef:basis.ownerPersonRef}:{}),...(basis.cycleRef?{cycleRef:basis.cycleRef}:{})};
  else if(kind==='OUTBOX_DISPATCH')p={effectRef:r,expectedAttemptState:d.status};
  else if(kind==='OUTBOX_RECONCILE')p={effectRef:d.effectRef||r,attemptId:d.attemptId,reconciliationEvidenceRef:basis.reconciliationEvidenceRef,provenOutcome:basis.provenOutcome};
  else if(kind==='INCIDENCE_SYNC')p={personRef:basis.personRef,cycleRef,membersEvidenceRef:basis.membersEvidenceRef,capacityAssessmentRef:d.currentAssessmentRef};
  else if(kind==='MEMBER_EXIT_EFFECT')p={membersEventRef:basis.membersEventRef,personRevision:basis.personRevision,affectedScopeManifestRef:basis.affectedScopeManifestRef};
  else if(kind==='CAPACITY_SYNC')p={cycleRef,sourceManifestRef:basis.sourceManifestRef,proposedChangeDigest:basis.proposedChangeDigest,evaluationRuleRef:ruleRef};
  else if(kind==='ARCHIVE')p={artifactRef:basis.artifactRef,receiptRef:basis.receiptRef,targetArchivePolicyRef:basis.targetArchivePolicyRef,operationKind:basis.operationKind};
  N(p&&Object.values(p).every(v=>v!==undefined),'UNRESOLVED_DEPENDENCY','JOB_PROOF_PENDING');N(scope==='PRESC_'+cycleRef.id,'AUTHORIZATION_DENIED');
  const quotaRefs=[];if(kind==='THEMATIC_TACIT'){const c=get(cycleRef),cfg=get(c.currentConfigRef);if(cfg.parameters.quota){const u=get(d.unitRef),old=u.quotaOccupancyRef&&get(u.quotaOccupancyRef),axes=[...new Map([d.principalAxisRef,d.candidate?.principalAxisRef,old?.axisRef].filter(Boolean).map(r=>[r.id,r])).values()];for(const axis of axes){const ref=M.ref('bucket',M.hash({cycle:cycleRef.id,axis:axis.id}).slice(0,32));if(services.refreshQuota)services.refreshQuota(ref,ruleRef,rule.revision);quotaRefs.push(ref);}}}
  const refs=new Map(),seen=new Set(),expected=[];
  function visit(v){if(!v||typeof v!=='object')return;if(v.kind&&v.id&&(M.collections[v.kind]||v.kind==='config'||['person','memberState','memberEvent'].includes(v.kind))){const path=M.path(v);if(seen.has(path))return;seen.add(path);N(seen.size<=64,'SCOPE_TOO_LARGE','JOB_READ_SET');const doc=services.db.get(path);N(doc,'UNRESOLVED_DEPENDENCY','JOB_REF');if(M.rev(doc.revision)){refs.set(path,v);expected.push({...v,revision:doc.revision});}for(const [k,value]of Object.entries(doc))if(!['sourceManifestRef','predecessorRef','predecessorCoverageRef','initialCoverageRef','lastCalendarChangeRef'].includes(k)&&(k.endsWith('Ref')||k.endsWith('Refs')))visit(value);return;}if(Array.isArray(v))v.forEach(visit);else Object.values(v).forEach(visit);}
  visit(p);quotaRefs.forEach(visit);visit(M.ref('capacity','registry'));visit(M.ref('capacity',cycleRef.id));N(refs.size<=32,'SCOPE_TOO_LARGE','JOB_SCOPE');return {payload:p,expected,scope:{targetRefs:[...refs.values()]},ruleVersion:rule.revision,ruleRef};
 }
 return {prepare};
}
