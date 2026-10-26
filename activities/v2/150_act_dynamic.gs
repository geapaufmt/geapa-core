// Dynamic requests keep their submission date. Approval does not reset D-30.
function f4ActivitiesDynamic_(x){
 const {M,N,p,load,save,revision,event,config,capacity,now}=x,D=f4ActivitiesDomain_(x);
 const lead=cfg=>{const days=cfg.parameters.dynamicRequestLeadDays??30;N(Number.isSafeInteger(days)&&days>=30,'CONFIG_CAPABILITY_DISABLED','DYNAMIC_REQUEST_DEADLINE');return days;};
 function reason(value){N(typeof value==='string'&&value.trim()&&value.length<=2000,'HUMAN_DECISION_REQUIRED','DYNAMIC_REASON');}
 function currentAssignment(u,sr){N(u.assignmentRef,'UNRESOLVED_DEPENDENCY');const a=load(u.assignmentRef,true);N(['ACTIVE','CONFIRMED'].includes(a.status)&&a.unitRef.id===u.id&&a.sessionRef.id===sr.id,'STALE_REVISION','DYNAMIC_ASSIGNMENT_CHANGED');return a;}
 return {
  DYNAMIC_SUBMIT(){
   const cfg=config();N(cfg.parameters.dynamic,'CONFIG_CAPABILITY_DISABLED');const u=x.own(p.unitRef);N(u.assignmentRef,'UNRESOLVED_DEPENDENCY');const a=load(u.assignmentRef,true),s=load(a.sessionRef,true);currentAssignment(u,a.sessionRef);
   N(!u.presentationRef||!load(u.presentationRef,true).realizationRef,'INVALID_STATE_TRANSITION','DYNAMIC_AFTER_REALIZATION');
   N(Array.isArray(p.elements)&&p.elements.length===9&&p.elements.every(e=>typeof e==='string'&&e.trim()&&e.length<=2000),'INCOMPLETE_REQUEST');
   let prior;if(u.currentDynamicProposalRef){prior=load(u.currentDynamicProposalRef,true);N(prior.state!=='OPEN'&&p.predecessorRef&&D.exactRef(p.predecessorRef,u.currentDynamicProposalRef),'INVALID_STATE_TRANSITION','DYNAMIC_PROPOSAL_PENDING');}else N(!p.predecessorRef,'STALE_REVISION','DYNAMIC_PREDECESSOR');
   const continued=prior?.requestedCivilDate===s.civilDate&&prior.sessionRef.id===s.id;
   if(!continued)D.date(s,lead(cfg));else N(prior.submittedAt&&Number.isSafeInteger(prior.requestLeadDays)&&prior.requestLeadDays>=30&&M.gap(M.day(M.ms(prior.requestSubmittedAt||prior.submittedAt)),s.civilDate)>=prior.requestLeadDays&&M.gap(M.day(now),s.civilDate)>=0,'UNRESOLVED_DEPENDENCY','DYNAMIC_SUBMISSION_PROOF');
   const r=event('flow',{kind:'DYNAMIC_PROPOSAL',state:'OPEN',cycleRef:x.cycleRef,personRef:u.personRef,unitRef:p.unitRef,sessionRef:a.sessionRef,elements:p.elements.map(e=>e.trim()),predecessorRef:p.predecessorRef||{kind:'NONE'},submittedAt:M.stamp(now),requestSubmittedAt:continued?prior.requestSubmittedAt||prior.submittedAt:M.stamp(now),requestedCivilDate:s.civilDate,requestLeadDays:continued?prior.requestLeadDays:lead(cfg),requestConfigRef:continued?prior.requestConfigRef||load(x.cycleRef).currentConfigRef:load(x.cycleRef).currentConfigRef});
   save(p.unitRef,{...u,currentDynamicProposalRef:r});return {proposalRef:r};
  },
  DYNAMIC_DECIDE(){
   const cfg=config();N(cfg.parameters.dynamic&&Number.isSafeInteger(cfg.parameters.dynamicLimit)&&cfg.parameters.dynamicLimit>=0,'CONFIG_CAPABILITY_DISABLED');
   const f=load(p.proposalRef,true);N(f.kind==='DYNAMIC_PROPOSAL'&&f.state==='OPEN','INVALID_STATE_TRANSITION');N(['APPROVE','REJECT','REQUEST_CORRECTION'].includes(p.decision),'VALIDATION_ERROR');reason(p.reason);
   let u=load(f.unitRef,true);N(D.exactRef(u.currentDynamicProposalRef,p.proposalRef),'STALE_REVISION','DYNAMIC_PROPOSAL_CHANGED');
   let requirementRef,occupancyRef,occupancy,capacityHead;
   if(p.decision==='APPROVE'){
    capacityHead=capacity();const s=load(f.sessionRef,true),assignment=currentAssignment(u,f.sessionRef);
    N(Number.isSafeInteger(f.requestLeadDays)&&f.requestLeadDays>=30&&f.submittedAt&&f.requestedCivilDate&&M.gap(M.day(M.ms(f.requestSubmittedAt||f.submittedAt)),f.requestedCivilDate)>=f.requestLeadDays,'UNRESOLVED_DEPENDENCY','DYNAMIC_SUBMISSION_PROOF');
    // A new date is a new calendar decision. The original date remains timely
    // even if the director reviews it with fewer than thirty days remaining.
    if(s.civilDate!==f.requestedCivilDate)D.date(s,Math.max(30,lead(cfg)));
    N(M.gap(M.day(now),s.civilDate)>=0,'INVALID_STATE_TRANSITION','DYNAMIC_DATE_PASSED');
    const assessment=x.manifest(p.assessmentRef),registry=load(M.ref('capacity','registry'),true);
    N(assessment.feasible===true&&assessment.cycleRef?.id===x.cycleRef.id&&assessment.membersEpoch===registry.membersEpoch&&assessment.cycleRegistryEpoch===registry.cycleRegistryEpoch&&assessment.basisEpoch===capacityHead.capacityEpoch&&assessment.validUntil&&M.ms(assessment.validUntil)>now,'STALE_REVISION','DYNAMIC_CAPACITY_ASSESSMENT');
    N(assignment.status==='CONFIRMED'&&s.slots.some(slot=>slot.unitRef.id===u.id&&slot.assignmentRef.id===u.assignmentRef.id)&&s.slots.every(slot=>slot.unitRef.id===u.id),'CAPACITY_EXHAUSTED','THIRD_PARTY');
    N(['ORDINARY','EXCLUSIVE','DYNAMIC_EXCLUSIVE'].includes(s.config.destination),'CAPACITY_EXHAUSTED','PROTECTED_SESSION');
    if(s.config.destination==='EXCLUSIVE')N(s.exclusiveUnitRef?.id===u.id,'CAPACITY_EXHAUSTED','EXCLUSIVE');
    const theme=u.thematicRequestRef&&load(u.thematicRequestRef,true);N(theme?.workflowState==='APPROVED','UNRESOLVED_DEPENDENCY','DYNAMIC_THEME_REQUIRED');
    if(cfg.parameters.quota){const quota=u.quotaOccupancyRef&&load(u.quotaOccupancyRef,true);N(quota?.phase==='C'&&quota.axisRef.id===theme.principalAxisRef.id,'UNRESOLVED_DEPENDENCY','DYNAMIC_THEMATIC_OCCUPANCY');}
    D.materialize(f.unitRef);u=load(f.unitRef);N(u.presentationRef,'UNRESOLVED_DEPENDENCY');const pr=load(u.presentationRef,true);N(!pr.realizationRef,'INVALID_STATE_TRANSITION','DYNAMIC_AFTER_REALIZATION');
    occupancyRef=u.dynamicOccupancyRef||M.ref('dynamic',M.hash({unit:u.id,cycle:x.cycleRef.id}).slice(0,32));const old=load(occupancyRef,true,true);N(!old||['ACTIVE','RELEASED'].includes(old.state),'INVALID_STATE_TRANSITION','HISTORICAL_DYNAMIC');
    N(!old||old.unitRef.id===u.id&&old.sessionRef.id===s.id,'STALE_REVISION','DYNAMIC_OCCUPANCY_CHANGED');
    const increment=old?.state==='ACTIVE'?0:1;N(capacityHead.dynamicActiveCount+capacityHead.dynamicHistoricalCount+increment<=cfg.parameters.dynamicLimit,'CAPACITY_EXHAUSTED','DYNAMIC_LIMIT');
    const material=x.manifest(p.requirementManifestRef);N(material.presentationRef?.id===pr.id&&Array.isArray(material.items)&&material.items.length>0&&material.items.length<=32&&new Set(material.items.map(i=>i.itemId)).size===material.items.length&&material.items.every(i=>typeof i.itemId==='string'&&i.itemId)&&material.finalEndExclusive&&Number.isFinite(M.ms(material.finalEndExclusive))&&material.uploadPolicyRef,'UNRESOLVED_DEPENDENCY','DYNAMIC_MATERIAL_SCOPE');
    const uploadPolicy=load(material.uploadPolicyRef);N(uploadPolicy.approved&&uploadPolicy.active,'CONFIG_CAPABILITY_DISABLED','DYNAMIC_MATERIAL_POLICY');
    requirementRef=pr.materialRequirementRef||M.ref('requirement',pr.id);const req=load(requirementRef,true,true),materialKey=v=>M.hash({items:v.items,deadline:M.ms(v.finalEndExclusive),policy:v.uploadPolicyRef}),sameRequirement=req&&materialKey(req)===materialKey(material);
    N(sameRequirement||!req||!req.currentSubmissionRef&&!req.currentRegularizationRef&&!req.openCorrectionCount,'INVALID_STATE_TRANSITION','MATERIAL_RIGHTS');
    if(!sameRequirement)revision(requirementRef,{...req,items:material.items,finalEndExclusive:material.finalEndExclusive,uploadPolicyRef:material.uploadPolicyRef,...(material.operatingHours?{operatingHours:material.operatingHours}:{}),...(material.finalCivilDate?{finalCivilDate:material.finalCivilDate}:{}),approvalManifestRef:p.requirementManifestRef,presentationRef:u.presentationRef,personRef:u.personRef,cycleRef:x.cycleRef,currentRevisionRef:{id:requirementRef.id,number:(req?.currentRevisionRef?.number||0)+1},materialStateRevision:(req?.materialStateRevision||0)+1,reviewEpoch:req?.reviewEpoch||0,opportunitiesEpoch:req?.opportunitiesEpoch||0,openCorrectionCount:req?.openCorrectionCount||0});
    occupancy={...old,unitRef:f.unitRef,cycleRef:x.cycleRef,sessionRef:f.sessionRef,presentationRef:u.presentationRef,state:'ACTIVE',approvedProposalRef:p.proposalRef,...(old?.state==='ACTIVE'&&old.priorSessionConfig?{priorSessionConfig:old.priorSessionConfig}:s.config.destination!=='DYNAMIC_EXCLUSIVE'?{priorSessionConfig:s.config}:{})};
    save(f.unitRef,{...u,dynamicOccupancyRef:occupancyRef});revision(u.presentationRef,{...pr,materialRequirementRef:requirementRef,dynamicOccupancyRef:occupancyRef});
    revision(f.sessionRef,{...s,config:{...s.config,destination:'DYNAMIC_EXCLUSIVE'},configRevision:s.configRevision+1});save(M.ref('capacity',x.cycleRef.id),{...capacityHead,dynamicActiveCount:capacityHead.dynamicActiveCount+increment,capacityEpoch:capacityHead.capacityEpoch+1});
   }
   const d=event('dynamicDecision',{cycleRef:x.cycleRef,personRef:u.personRef,unitRef:f.unitRef,proposalRef:p.proposalRef,elements:f.elements,decision:p.decision,reason:p.reason,capacityAssessmentRef:p.assessmentRef,...(requirementRef?{requirementRef,requirementManifestRef:p.requirementManifestRef}:{})});
   if(occupancyRef)revision(occupancyRef,{...occupancy,approvalDecisionRef:d});revision(p.proposalRef,{...f,state:p.decision==='REQUEST_CORRECTION'?'CORRECTION_REQUESTED':'DECIDED',decisionRef:d});return {decisionRef:d};
  },
  DYNAMIC_CANCEL(){
   capacity();const d=load(p.occupancyRef,true),u=x.own(d.unitRef);N(['WITHDRAWAL','CANCELLATION','REVOCATION'].includes(p.cause),'VALIDATION_ERROR');N(x.actor.kind!=='PORTAL_AUTHENTICATED_MEMBER'||p.cause==='WITHDRAWAL','AUTHORIZATION_DENIED');reason(p.decisionBasis?.reason);
   N(d.state==='ACTIVE'&&(!u.presentationRef||!load(u.presentationRef,true).realizationRef),'INVALID_STATE_TRANSITION','HISTORY_NOT_RELEASED');
   D.releaseDynamic(p.occupancyRef,{releaseCause:p.cause,releaseDecision:p.decisionBasis,releasedAt:M.stamp(now)});return {state:'RELEASED'};
  }
 };
}
