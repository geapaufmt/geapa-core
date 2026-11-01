function f4ActivitiesQuotaCommands_(x){
 const {M,N,p,load,save,revision,event}=x,D=f4ActivitiesDomain_(x);
 function reason(){N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'HUMAN_DECISION_REQUIRED');}
 return {
  QUOTA_CONFIGURE(){
   reason();const c=load(p.cycleRef,true),cfg=x.config(),axes=cfg.parameters.axisRefs||[];N(cfg.parameters.quota&&axes.length>0&&axes.length<=12,'CONFIG_CAPABILITY_DISABLED','QUOTA_AXES_REQUIRED');
   N(['DISABLED','MANUAL_NEW_ANALYSIS'].includes(p.waitAdmissionMode)&&Array.isArray(p.axisLimits)&&p.axisLimits.length===axes.length,'VALIDATION_ERROR');
   const ids=new Set();for(const item of p.axisLimits){f4Core_().exact(item,['axisRef','limit']);N(item.axisRef.kind==='policy'&&axes.some(r=>M.path(r)===M.path(item.axisRef))&&!ids.has(item.axisRef.id)&&Number.isSafeInteger(item.limit)&&item.limit>=0&&item.limit<=180,'VALIDATION_ERROR');ids.add(item.axisRef.id);}
   const h=load(M.ref('capacity',c.id),true),prior=c.thematicQuotaConfigRef&&load(c.thematicQuotaConfigRef);let initialCount=prior?.initialObligationCount;
   if(!prior||c.status==='DRAFT'){N(c.status==='DRAFT','HUMAN_DECISION_REQUIRED','INITIAL_QUOTA_BEFORE_OPEN');const coverage=h.incidenceCoverageRef&&load(h.incidenceCoverageRef),g=load(M.ref('capacity','registry'),true);N(coverage?.approved&&coverage.complete&&coverage.coverageKind==='act-incidence-coverage-v1'&&coverage.membersEpoch===g.membersEpoch&&coverage.cycleRegistryEpoch===g.cycleRegistryEpoch,'UNRESOLVED_DEPENDENCY','INCIDENCE_COVERAGE_REQUIRED');initialCount=coverage.items.filter(i=>i.disposition==='INCIDENT'&&i.obligationRef).length;N(p.axisLimits.reduce((n,i)=>n+i.limit,0)===initialCount,'VALIDATION_ERROR','INITIAL_QUOTA_SUM');}
   for(const item of p.axisLimits){const br=D.bucketRef(item.axisRef),b=load(br,true,true);N(!b||b.reconciliationState==='CLEAN','UNRESOLVED_DEPENDENCY','QUOTA_DIRTY');N(!b||item.limit>=b.heldCount+b.confirmedCount+b.realizedCount,'HUMAN_DECISION_REQUIRED','QUOTA_EXISTING_RIGHTS');}
   const decisionRef=event('flow',{kind:'QUOTA_CONFIGURATION',state:'APPROVED',cycleRef:p.cycleRef,axisLimits:p.axisLimits,waitAdmissionMode:p.waitAdmissionMode,initialObligationCount:initialCount,reason:p.reason.trim(),...(c.thematicQuotaConfigRef?{predecessorRef:c.thematicQuotaConfigRef}:{})});
   for(const item of p.axisLimits){const br=D.bucketRef(item.axisRef),b=load(br,false,true)||{cycleRef:p.cycleRef,axisRef:item.axisRef,heldCount:0,confirmedCount:0,realizedCount:0,epoch:0,pendingQueueEpoch:0,reconciliationState:'CLEAN'};save(br,{...b,limit:item.limit,epoch:b.epoch+1,pendingQueueEpoch:b.pendingQueueEpoch+1,limitDecisionRef:decisionRef});}
   save(p.cycleRef,{...c,thematicQuotaConfigRef:decisionRef});x.bumpCapacity();return {decisionRef,initialObligationCount:initialCount};
  },
  THEME_WAIT_ADMIT(){
   reason();const t=load(p.requestRef,true),u=load(t.unitRef,true),c=load(u.cycleRef,true),cfg=x.config(),policy=c.thematicQuotaConfigRef&&load(c.thematicQuotaConfigRef);
   N(cfg.parameters.quota&&policy?.state==='APPROVED'&&policy.waitAdmissionMode==='MANUAL_NEW_ANALYSIS','CONFIG_CAPABILITY_DISABLED','WAIT_ADMISSION_POLICY_REQUIRED');
   N(t.workflowState==='WAITING'&&t.candidate&&t.candidate.materiality!=='UNDECIDED','HUMAN_DECISION_REQUIRED','WAITING_MATERIALITY');D.eligible(u);if(u.assignmentRef)D.date(load(load(u.assignmentRef).sessionRef),cfg.parameters.thematicMinimumDays??20,20);
   N(cfg.parameters.clockPolicyRef,'CONFIG_CAPABILITY_DISABLED','ANALYSIS_CLOCK_REQUIRED');const clockPolicy=load(cfg.parameters.clockPolicyRef);
   const target={...t,principalAxisRef:t.candidate.principalAxisRef};
   // A pending change keeps its original occupancy until the final approval.
   D.quotaCheck(target,t.unitRef);if(!t.acceptedRevisionRef)D.allocate(target,t.unitRef,'H');
   const decisionRef=event('flow',{kind:'THEMATIC_WAIT_ADMISSION',state:'ADMITTED',cycleRef:u.cycleRef,requestRef:p.requestRef,requestRevisionRef:t.submittedRevisionRef,policyRef:c.thematicQuotaConfigRef,reason:p.reason.trim()});
   const analysisClock={...f4ActivitiesClock_().start(clockPolicy,x.now),policyRef:cfg.parameters.clockPolicyRef,manualDecisionEpoch:(t.manualDecisionEpoch||0)+1};
   revision(p.requestRef,{...t,workflowState:'ANALYSIS',clockState:'RUNNING',analysisClock,nextEvaluationAt:analysisClock.dueAt,manualDecisionEpoch:analysisClock.manualDecisionEpoch,waitAdmissionRef:decisionRef});
   const br=D.bucketRef(target.principalAxisRef),b=load(br,true);save(br,{...b,pendingQueueEpoch:b.pendingQueueEpoch+1});x.bumpCapacity();return {state:'ANALYSIS',decisionRef};
  },
  QUOTA_PRIORITY_RESOLVE(){
   reason();const b=load(p.bucketRef,true),proof=x.manifest(b.queueManifestRef);N(proof.queueEpoch===b.pendingQueueEpoch&&proof.items.length>1&&!proof.unresolvedPriority,'STALE_REVISION','PRIORITY_TIE_CHANGED');
   N(proof.items.some(i=>i.ref.id===p.requestRef.id)&&proof.items.every(i=>i.priorityCompleteAt===proof.minimumPriorityAt),'VALIDATION_ERROR','PRIORITY_TIE_REQUIRED');
   const t=load(p.requestRef,true);N(t.cycleRef.id===b.cycleRef.id,'VALIDATION_ERROR');
   const resolutionRef=event('flow',{kind:'QUOTA_PRIORITY_RESOLUTION',state:'APPROVED',cycleRef:b.cycleRef,bucketRef:p.bucketRef,queueEpoch:b.pendingQueueEpoch,queueDigest:proof.digest,selectedRequestRef:p.requestRef,reason:p.reason.trim()});save(p.bucketRef,{...b,priorityResolutionRef:resolutionRef});return {resolutionRef};
  }
 };
}
function f4ActivitiesQuotaOpening_(x){
 const M=x.M,cfg=x.config();if(!cfg.parameters.quota)return;
 const c=x.load(x.cycleRef),decision=c.thematicQuotaConfigRef&&x.load(c.thematicQuotaConfigRef);x.N(decision?.kind==='QUOTA_CONFIGURATION'&&decision.state==='APPROVED','CONFIG_CAPABILITY_DISABLED','QUOTA_CONFIGURATION_REQUIRED');
 const axes=cfg.parameters.axisRefs||[];x.N(decision.axisLimits.length===axes.length&&axes.every(r=>decision.axisLimits.some(i=>i.axisRef.id===r.id)),'CONFIG_CAPABILITY_DISABLED','QUOTA_AXES_CHANGED');
 const h=x.load(M.ref('capacity',c.id)),coverage=h.incidenceCoverageRef&&x.load(h.incidenceCoverageRef);x.N(coverage?.approved&&coverage.complete&&coverage.items.filter(i=>i.disposition==='INCIDENT'&&i.obligationRef).length===decision.initialObligationCount&&decision.axisLimits.reduce((n,i)=>n+i.limit,0)===decision.initialObligationCount,'UNRESOLVED_DEPENDENCY','INITIAL_QUOTA_SUM');
 for(const i of decision.axisLimits){const b=x.load(f4ActivitiesDomain_(x).bucketRef(i.axisRef));x.N(b.limit===i.limit&&b.reconciliationState==='CLEAN'&&b.heldCount+b.confirmedCount+b.realizedCount<=b.limit,'UNRESOLVED_DEPENDENCY','QUOTA_NOT_READY');}
}
