function f4ActivitiesThematic_(x){
 const {M,N,p,load,save,revision,event,config,capacity,now}=x,D=f4ActivitiesDomain_(x);
 function request(){const t=x.own(p.requestRef);N(t.workflowState!=='WITHDRAWN','INVALID_STATE_TRANSITION');return t;}
 function queue(t,priorAxis){if(!config().parameters.quota)return;const axes=[...new Map([t.principalAxisRef,t.candidate?.principalAxisRef,priorAxis].filter(Boolean).map(r=>[r.id,r])).values()];for(const axis of axes){const r=D.bucketRef(axis),b=load(r,true);save(r,{...b,pendingQueueEpoch:b.pendingQueueEpoch+1});}x.bumpCapacity();}
 function reviewedContent(t,v){
  N(t.workflowState==='ANALYSIS'&&p.decision==='APPROVE'&&t.candidate,'INVALID_STATE_TRANSITION');
  f4Core_().exact(v,['titleOrObject','principalAxisRef','materiality','reason',...('secondaryAxisRef'in v?['secondaryAxisRef']:[])]);
  N(typeof v.titleOrObject==='string'&&v.titleOrObject.trim()&&v.titleOrObject.length<=4000&&['EDITORIAL','NON_MATERIAL_COMPLEMENT','SUBSTANTIAL'].includes(v.materiality)&&typeof v.reason==='string'&&v.reason.trim()&&v.reason.length<=4000,'HUMAN_DECISION_REQUIRED');
  const cfg=config().parameters;f4CycleParameters_().axes(cfg,v.principalAxisRef,v.secondaryAxisRef,x.get);
  N(!v.secondaryAxisRef||v.secondaryAxisRef.id!==v.principalAxisRef.id,'VALIDATION_ERROR','AXIS_DUPLICATE');
  const before=t.candidate,changed=before.titleOrObject!==v.titleOrObject.trim()||before.principalAxisRef.id!==v.principalAxisRef.id||(before.secondaryAxisRef?.id||'')!==(v.secondaryAxisRef?.id||'');
  t.candidate={...before,titleOrObject:v.titleOrObject.trim(),principalAxisRef:v.principalAxisRef,materiality:v.materiality,...(changed?{completeAt:M.stamp(now)}:{})};
  if(v.secondaryAxisRef)t.candidate.secondaryAxisRef=v.secondaryAxisRef;else delete t.candidate.secondaryAxisRef;
  if(changed)t.submittedRevisionRef={...t.submittedRevisionRef,number:t.revision+1};
  t.materialityDecision={basis:{kind:'BOARD_EDIT_AND_APPROVE'},reason:v.reason,actor:x.actor,decidedAt:M.stamp(now),materiality:v.materiality};
  if(v.materiality==='SUBSTANTIAL')t.priorityBasis={completeAt:t.candidate.completeAt,operationId:x.operationId,predecessor:t.priorityBasis};
  t.priorityCompleteAt=t.priorityBasis.completeAt;
 }
 function decide(t,choice,tacit){N(['APPROVE','REJECT','REQUEST_COMPLEMENT'].includes(choice));N(t.workflowState==='ANALYSIS','INVALID_STATE_TRANSITION');const cfg=config(),priorAxis=t.principalAxisRef;if(choice==='APPROVE'){N(!t.candidate||t.candidate.materiality!=='UNDECIDED','HUMAN_DECISION_REQUIRED','MATERIALITY');if(t.candidate){t.principalAxisRef=t.candidate.principalAxisRef;t.secondaryAxisRef=t.candidate.secondaryAxisRef||{kind:"NONE"};t.acceptedRevisionRef=t.submittedRevisionRef;}D.allocate(t,t.unitRef,'C');t.workflowState='APPROVED';}
  if(choice==='REJECT'){t.workflowState='REJECTED';const u=load(t.unitRef,true);if(!t.acceptedRevisionRef&&u.quotaOccupancyRef)D.releaseQuota(u.quotaOccupancyRef);}if(choice==='REQUEST_COMPLEMENT'){N(p.correctionTerms?.nonMaterial===true,'HUMAN_DECISION_REQUIRED','MATERIALITY');const r=event('flow',{kind:'THEMATIC_ADJUSTMENT',state:'OPEN',cycleRef:x.cycleRef,personRef:load(t.unitRef).personRef,requestRef:p.requestRef,nonMaterial:true,terms:p.correctionTerms});t.correctionCaseRef=r;t.workflowState='PAUSED';t.analysisClock=t.extraordinarySubmission?{...t.analysisClock,state:'MANUAL_PAUSED'}:f4ActivitiesClock_().pause(load(t.analysisClock.policyRef),t.analysisClock,now);}
  if(!tacit)t.manualDecisionEpoch=(t.manualDecisionEpoch||0)+1;
  if(!tacit)N(typeof p.reason==='string'&&p.reason.length<=4000&&(p.internalReason===undefined||typeof p.internalReason==='string'&&p.internalReason.length<=4000),'VALIDATION_ERROR');
  const vr=event('validation',{...(!tacit?{reason:p.reason,internalReason:p.internalReason||''}:{}),personRef:load(t.unitRef).personRef,cycleRef:x.cycleRef,...f4CycleParameters_().reconsideration(cfg.parameters,now),requestRevisionRef:t.submittedRevisionRef,result:choice,decidedAt:M.stamp(now),effectiveAt:tacit?t.analysisClock.dueAt:M.stamp(now),clockBasis:t.analysisClock,configRef:load(x.cycleRef).currentConfigRef},p.requestRef.id);
  t.lastValidationRef=vr;t.clockState=t.workflowState==='PAUSED'?'PAUSED':'STOPPED';revision(p.requestRef,t);if(choice==='APPROVE')D.materialize(t.unitRef);queue(t,priorAxis);return {state:t.workflowState,validationRef:vr};
 }
 return {
 THEME_SUBMIT(){const cfg=config().parameters,u=x.own(p.unitRef);D.eligible(u);f4CycleParameters_().axes(cfg,p.principalAxisRef,p.secondaryAxisRef,x.get);N(!p.secondaryAxisRef||p.secondaryAxisRef.id!==p.principalAxisRef.id,'VALIDATION_ERROR','AXIS_DUPLICATE');N(typeof p.titleOrObject==='string'&&p.titleOrObject.trim()&&p.summary.trim(),'INCOMPLETE_REQUEST');N(['INITIAL','CHANGE_PROPOSAL'].includes(p.submissionKind));N(u.assignmentRef||cfg.early,'CONFIG_CAPABILITY_DISABLED','EARLY_OFF');if(u.assignmentRef)D.date(load(load(u.assignmentRef,true).sessionRef,true),undefined,undefined,'THEME',u);const rr=u.thematicRequestRef||M.ref('thematic',M.hash({unit:p.unitRef.id,cycle:x.cycleRef.id}).slice(0,32)),old=load(rr,true,true);N(p.submissionKind==='INITIAL'?!old:!!old,'INVALID_STATE_TRANSITION');const nextRev=(old?.revision||0)+1,submitted={kind:'thematicRevision',id:rr.id,number:nextRev};const priority={completeAt:M.stamp(now),operationId:x.operationId};const t={...old,unitRef:p.unitRef,personRef:u.personRef,cycleRef:x.cycleRef,principalAxisRef:old?.principalAxisRef||p.principalAxisRef,submittedRevisionRef:submitted,workflowState:'ANALYSIS',completeness:'COMPLETE',priorityBasis:old?.priorityBasis||priority,candidate:{titleOrObject:p.titleOrObject,summary:p.summary,principalAxisRef:p.principalAxisRef,...(p.secondaryAxisRef?{secondaryAxisRef:p.secondaryAxisRef}:{}),completeAt:M.stamp(now),materiality:old?'UNDECIDED':'INITIAL'},manualDecisionEpoch:old?.manualDecisionEpoch||0};
  if(cfg.clockPolicyRef){const clock=load(cfg.clockPolicyRef);t.analysisClock={...f4ActivitiesClock_().start(clock,now),policyRef:cfg.clockPolicyRef,manualDecisionEpoch:t.manualDecisionEpoch};}else t.analysisClock={state:'UNRESOLVED'};t.clockState=t.analysisClock.state;if(t.analysisClock.dueAt)t.nextEvaluationAt=t.analysisClock.dueAt;t.priorityCompleteAt=t.priorityBasis.completeAt;
  t.id=rr.id;
  if(cfg.quota){const b=load(D.bucketRef(p.principalAxisRef),true),oc=u.quotaOccupancyRef&&load(u.quotaOccupancyRef),same=oc&&oc.phase!=='RELEASED'&&oc.axisRef.id===p.principalAxisRef.id;if(!same&&b.heldCount+b.confirmedCount+b.realizedCount>=b.limit)t.workflowState='WAITING';else if(!old)try{D.allocate(t,p.unitRef,'H');}catch(e){if(e.code==='THEMATIC_QUOTA_EXHAUSTED'||['PRIORITY_TIE','PRIORITY_UNRESOLVED'].includes(e.reasonCode)){t.workflowState='WAITING';t.waitReason=e.reasonCode||e.code;}else throw e;}}
  if(t.workflowState==='WAITING'){t.analysisClock={state:'WAITING_ADMISSION'};t.clockState='WAITING_ADMISSION';delete t.nextEvaluationAt;}
  revision(rr,t);save(p.unitRef,{...load(p.unitRef),thematicRequestRef:rr});queue(t);return {requestRef:rr,state:t.workflowState,priorityBasis:t.priorityBasis};},
 THEME_COMPLEMENT(){const t=request(),f=load(p.correctionCaseRef,true);N(t.workflowState==='PAUSED'&&f.state==='OPEN'&&f.nonMaterial===true,'HUMAN_DECISION_REQUIRED');N(p.requestRevisionRef.number===t.submittedRevisionRef.number,'STALE_REVISION');N(Object.keys(p.contentVersion).every(k=>['titleOrObject','summary'].includes(k)),'VALIDATION_ERROR','COMPLEMENT_SCOPE');t.candidate={...t.candidate,...p.contentVersion,materiality:'NON_MATERIAL_COMPLEMENT'};t.analysisClock=t.extraordinarySubmission?{...t.analysisClock,state:'MANUAL_ONLY'}:f4ActivitiesClock_().resume(load(t.analysisClock.policyRef),t.analysisClock,now);t.workflowState='ANALYSIS';t.clockState=t.analysisClock.state;if(t.analysisClock.dueAt)t.nextEvaluationAt=t.analysisClock.dueAt;else delete t.nextEvaluationAt;t.manualDecisionEpoch++;t.analysisClock.manualDecisionEpoch=t.manualDecisionEpoch;revision(p.requestRef,t);revision(p.correctionCaseRef,{...f,state:'ANSWERED'});queue(t);return {priorityBasis:t.priorityBasis};},
 THEME_MATERIALITY_DECIDE(){const t=request();N(t.candidate&&p.submittedRevisionRef.number===t.submittedRevisionRef.number,'STALE_REVISION');N(['EDITORIAL','NON_MATERIAL_COMPLEMENT','SUBSTANTIAL'].includes(p.materiality));t.candidate.materiality=p.materiality;t.materialityDecision={basis:p.decisionBasis,reason:p.reason,actor:x.actor,decidedAt:M.stamp(now)};if(p.materiality==='SUBSTANTIAL')t.priorityBasis={completeAt:t.candidate.completeAt,operationId:x.operationId,predecessor:t.priorityBasis};t.priorityCompleteAt=t.priorityBasis.completeAt;revision(p.requestRef,t);queue(t);return {priorityBasis:t.priorityBasis};},
 THEME_VALIDATE(){const t=request();N(p.submittedRevisionRef.number===t.submittedRevisionRef.number,'STALE_REVISION');if(t.analysisClock?.state==='RUNNING'&&M.ms(t.analysisClock.dueAt)<now)N(p.decision==='APPROVE'&&!p.reviewedContent,'HUMAN_DECISION_REQUIRED','LATE_MANUAL_ACT');if(p.reviewedContent)reviewedContent(t,p.reviewedContent);return decide(t,p.decision,false);},
 THEME_TACIT_APPLY(){const t=request();N(t.revision===p.basisRevision&&t.analysisClock?.state==='RUNNING','STALE_REVISION');N(t.analysisClock.policyRef?.id===p.clockPolicyRef.id&&t.analysisClock.manualDecisionEpoch===t.manualDecisionEpoch,'STALE_REVISION','CLOCK_BASE');N(p.dueAt===M.ms(t.analysisClock.dueAt)&&p.dueAt<=now,'DEADLINE_VIOLATION');N(load(p.clockPolicyRef).approved,'HUMAN_DECISION_REQUIRED','X03');return decide(t,'APPROVE',true);},
 THEME_WITHDRAW(){const t=request(),u=load(t.unitRef,true);if(u.quotaOccupancyRef)D.releaseQuota(u.quotaOccupancyRef);revision(p.requestRef,{...t,workflowState:'WITHDRAWN'});queue(t);return {state:'WITHDRAWN'};},
 COMPATIBILITY_REQUEST(){return f4ActivitiesCompatibilityRequest_(x);},
 COMPATIBILITY_DECIDE(){return f4ActivitiesRecordCompatibility_(x);},
 ...f4ActivitiesThemeExtraordinary_({...x,themeQueue:queue}),
 ...f4ActivitiesDynamic_(x)
 };
}


// Individual, revocable authorization. Uses the enclosing command transaction,
// canonical ROLE/SELF proof, expected vector, receipt and audit effect vector.
function f4ActivitiesThemeExtraordinary_(x){
 const {M,N,p,load,save,revision,event,now}=x;
 const text=(v,max=2000)=>{N(typeof v==='string'&&v.trim()&&v.length<=max,'INCOMPLETE_REQUEST');return v.trim();};
 function administrative(){const b=x.authorityBasis.functionBasis;N(x.actor.kind==='BOARD_ACTOR'&&b?.kind==='ROLE'&&b.capability==='PRESENTATION_THEME_REVIEW'&&x.get('memberRoleCatalog/'+b.roleId)?.roleKey==='DIRETOR_ADMINISTRATIVO','AUTHORIZATION_DENIED','THEME_EXTRAORDINARY_AUTHORITY');}
 function state(direct=false){const u=load(p.unitRef,true);
  if(direct&&u.presentationRef){
   const pr=load(u.presentationRef,true),r=pr.realizationRef&&load(pr.realizationRef);
   N(pr.unitRef?.id===u.id&&pr.personRef?.id===u.personRef.id&&pr.cycleRef?.id===u.cycleRef.id,'UNRESOLVED_DEPENDENCY','PRESENTATION_UNIT_MISMATCH');
   if(r){
    N(r.recognized===true&&r.occurred===true&&r.presentationRef?.id===pr.id,'INVALID_STATE_TRANSITION','DOCUMENTED_REALIZATION_REQUIRED');
    return {u,pr,r,mode:'REALIZED'};
   }
  }
  N(!u.membershipExit&&u.assignmentRef,'INVALID_STATE_TRANSITION','ASSIGNED_PRESENTATION_REQUIRED');const a=load(u.assignmentRef,true),s=load(a.sessionRef,true);N(a.status==='CONFIRMED'&&a.unitRef?.id===u.id,'INVALID_STATE_TRANSITION');const pr=u.presentationRef&&load(u.presentationRef,true),r=pr?.realizationRef&&load(pr.realizationRef);if(r){N(r.recognized===true&&r.occurred===true&&r.presentationRef?.id===pr.id,'INVALID_STATE_TRANSITION','DOCUMENTED_REALIZATION_REQUIRED');return {u,a,s,pr,r,mode:'REALIZED'};}N(M.day(now)<=s.civilDate,'UNRESOLVED_DEPENDENCY','PAST_PRESENTATION_REQUIRES_REALIZATION');f4ActivitiesDomain_(x).eligible(u);return {u,a,s,pr,mode:'SCHEDULED'};}
 function audit(kind,u,details){return event('flow',{kind,personRef:u.personRef,unitRef:p.unitRef,cycleRef:u.cycleRef,actor:x.actor,recordedAt:M.stamp(now),operationId:x.operationId,...details});}
 function authorize(){administrative();N(p.confirmed===true,'HUMAN_DECISION_REQUIRED');const v=state(),u=v.u;N(u.extraordinaryTheme?.status!=='ACTIVE','CONFLICT','THEME_AUTHORIZATION_ACTIVE');const reason=text(p.reason),record=audit('THEME_EXTRAORDINARY_AUTHORIZED',u,{reason,mode:v.mode,authorityBasis:x.authorityBasis});save(p.unitRef,{...u,extraordinaryTheme:{status:'ACTIVE',mode:v.mode,reason,authorizedAt:M.stamp(now),authorizedBy:x.actor.personRef,authorizationRef:record,assignmentRef:u.assignmentRef,sessionRevision:v.s.revision,...(v.r?{realizationRef:v.pr.realizationRef}:{})}});return {state:'ACTIVE',authorizationRef:record};}
 function revoke(){administrative();N(p.confirmed===true,'HUMAN_DECISION_REQUIRED');const u=load(p.unitRef,true),g=u.extraordinaryTheme;N(g?.status==='ACTIVE','INVALID_STATE_TRANSITION');audit('THEME_EXTRAORDINARY_REVOKED',u,{authorizationRef:g.authorizationRef});save(p.unitRef,{...u,extraordinaryTheme:{...g,status:'REVOKED',endedAt:M.stamp(now),endOperationId:x.operationId}});return {state:'REVOKED'};}
 function documentation(content,required){
  const e=p.institutionalEvidence;
  if(!required&&e===undefined)return {};
  N(e&&typeof e==='object','INCOMPLETE_REQUEST','INSTITUTIONAL_CONTENT_EVIDENCE_REQUIRED');
  f4Core_().exact(e,['reference','receivedAt']);
  const reference=text(e.reference);
  N(Number.isSafeInteger(e.receivedAt)&&e.receivedAt>0&&e.receivedAt<=now,'VALIDATION_ERROR','INSTITUTIONAL_RECEIPT_DATE');
  return {institutionalEvidence:{reference,receivedAt:M.stamp(e.receivedAt),attestedAt:M.stamp(now),attestedBy:x.actor.personRef,contentHash:M.hash(content)}};
 }
 function directContext(v){
  N(p.confirmed===true,'HUMAN_DECISION_REQUIRED');
  const reason=text(p.reason);
  if(v.mode==='SCHEDULED'){
   const cfg=x.config().parameters;
   N(!f4CycleParameters_().dateAvailable(cfg,'thematicMinimumDays',M.day(now),v.s.civilDate)&&!f4ActivitiesPreagenda_().availableDeadline(x,v.s,'THEME',v.u),'INVALID_STATE_TRANSITION','ORDINARY_PERIOD_OPEN');
  }
  return {reason,submissionChannel:'ADMIN_INSTITUTIONAL',presenterRef:v.u.personRef,registeredBy:x.actor.personRef,recordedAt:M.stamp(now)};
 }
 function submit(admin){
  if(admin)administrative();else N(x.actor.kind==='PORTAL_AUTHENTICATED_MEMBER','AUTHORIZATION_DENIED');
  const direct=admin&&p.reason!==undefined,v=state(direct),u=v.u,g=u.extraordinaryTheme;
  if(admin&&p.section!==undefined){
   N(direct,'INCOMPLETE_REQUEST');
   return f4ActivitiesAdministrativeFile_(x,v,directContext(v),documentation);
  }
  N(p.fileDescriptor===undefined&&p.effectRef===undefined&&p.fileRevision===undefined,'VALIDATION_ERROR');
  const details=direct?directContext(v):{};
  if(!direct){
   N(admin||u.personRef.id===x.actor.personId,'AUTHORIZATION_DENIED','OWNER_MISMATCH');
   N(g?.status==='ACTIVE'&&g.mode===v.mode&&g.assignmentRef.id===u.assignmentRef.id&&g.sessionRevision===v.s.revision,'STALE_REVISION','THEME_AUTHORIZATION_CHANGED');
   if(v.mode==='REALIZED')N(g.realizationRef?.id===v.pr.realizationRef.id,'STALE_REVISION');
  }
  const consumes=g?.status==='ACTIVE',grantRef=consumes?{authorizationRef:g.authorizationRef}:{};
  N(p.principalAxisRef&&typeof p.titleOrObject==='string'&&typeof p.summary==='string','INCOMPLETE_REQUEST');
  const cfg=x.config().parameters;f4CycleParameters_().axes(cfg,p.principalAxisRef,p.secondaryAxisRef,x.get);N(!p.secondaryAxisRef||p.secondaryAxisRef.id!==p.principalAxisRef.id,'VALIDATION_ERROR','AXIS_DUPLICATE');const content={titleOrObject:text(p.titleOrObject,4000),summary:text(p.summary,4000),principalAxisRef:p.principalAxisRef,...(p.secondaryAxisRef?{secondaryAxisRef:p.secondaryAxisRef}:{})};
  if(admin)N(p.confirmed===true,'HUMAN_DECISION_REQUIRED');
  const evidence=admin?documentation(content,!direct):{};
  const rr=u.thematicRequestRef||M.ref('thematic',M.hash({unit:u.id,cycle:u.cycleRef.id}).slice(0,32)),old=load(rr,true,true);N(direct&&v.mode==='REALIZED'||!old||!['ANALYSIS','PAUSED','WAITING'].includes(old.workflowState),'CONFLICT','THEME_REVIEW_PENDING');if(direct&&v.mode==='SCHEDULED')N(!old?.acceptedRevisionRef&&old?.workflowState!=='APPROVED'&&!v.pr?.finalConfirmationRef,'HUMAN_DECISION_REQUIRED','APPROVED_THEME_REQUIRES_CHANGE_PROCEDURE');
  const recorded=audit('THEME_EXTRAORDINARY_SUBMITTED',u,{...grantRef,mode:v.mode,submissionChannel:admin?'ADMIN_INSTITUTIONAL':'SELF',...details,...evidence,content});
  const closed=consumes?{extraordinaryTheme:{...g,status:'CONSUMED',endedAt:M.stamp(now),endOperationId:x.operationId,submissionRef:recorded}}:{};
  const grantResult=consumes?{authorizationState:'CONSUMED'}:{};
  // Post-realization is a documentary addendum, never a replacement proposal.
  // Keep the thematic head, accepted revision, quota and realization byte-identical.
  if(v.mode==='REALIZED'){save(p.unitRef,{...u,lastThemeRegularizationRef:recorded,...closed});return {recordRef:recorded,state:'RECORDED_NOT_APPROVED',...grantResult};}
  const t={...old,id:rr.id,unitRef:p.unitRef,personRef:u.personRef,cycleRef:u.cycleRef,principalAxisRef:old?.principalAxisRef||p.principalAxisRef,submittedRevisionRef:{kind:'thematicRevision',id:rr.id,number:(old?.revision||0)+1},workflowState:'ANALYSIS',completeness:'COMPLETE',candidate:{...content,completeAt:M.stamp(now),materiality:old?'UNDECIDED':'INITIAL'},priorityBasis:old?.priorityBasis||{completeAt:M.stamp(now),operationId:x.operationId},priorityCompleteAt:old?.priorityBasis?.completeAt||M.stamp(now),manualDecisionEpoch:(old?.manualDecisionEpoch||0)+1,analysisClock:{state:'MANUAL_ONLY'},clockState:'MANUAL_ONLY',extraordinarySubmission:{mode:v.mode,recordedAt:M.stamp(now),recordRef:recorded,...grantRef}};delete t.nextEvaluationAt;revision(rr,t);save(p.unitRef,{...u,thematicRequestRef:rr,...closed});x.themeQueue?.(t);return {requestRef:rr,state:'ANALYSIS',...grantResult};
 }
 return {THEME_EXTRAORDINARY_AUTHORIZE:authorize,THEME_EXTRAORDINARY_REVOKE:revoke,THEME_REGULARIZE:()=>submit(false),THEME_REGULARIZE_ADMIN:()=>submit(true)};
}

function f4ActivitiesAdministrativeFile_(x,v,record,documentation){
 const {M,N,p}=x,{u,pr}=v,photo=p.section==='PHOTO';
 N(photo||p.section==='MATERIAL','VALIDATION_ERROR');
 N(pr&&pr.unitRef?.id===u.id&&pr.personRef?.id===u.personRef.id&&pr.cycleRef?.id===u.cycleRef.id,'UNRESOLVED_DEPENDENCY','CANONICAL_PRESENTATION_REQUIRED');
 N(!['titleOrObject','summary','principalAxisRef','secondaryAxisRef'].some(k=>k in p),'VALIDATION_ERROR');
 N(!!p.fileDescriptor!==!!p.effectRef&&Number.isSafeInteger(p.fileRevision),'VALIDATION_ERROR');
 const rq=!photo&&pr.materialRequirementRef&&x.load(pr.materialRequirementRef,true);
 const policyRef=photo?f4ActivitiesPhotoPolicy_(x.get,pr).ref:rq?.uploadPolicyRef;
 const policy=policyRef&&x.load(policyRef);
 N(policy?.approved&&policy.active&&policy.workspace?.enabled,'CONFIG_CAPABILITY_DISABLED');
 if(photo)N(policy.allowBoardUpload===true,'AUTHORIZATION_DENIED');
 else N(rq?.presentationRef?.id===pr.id&&rq.personRef?.id===u.personRef.id&&!rq.currentRegularizationRef&&!rq.currentSubmissionRef&&!rq.openCorrectionCount&&!(rq.receiptRefs||[]).length,'INVALID_STATE_TRANSITION','MATERIAL_REQUIRES_SEPARATE_PROCEDURE');
 N(photo?!pr.photo&&p.fileRevision===0:rq.revision===p.fileRevision,'STALE_REVISION','DOCUMENT_CHANGED');
 const proof=p.effectRef&&x.load(p.effectRef,true);
 const descriptor=p.fileDescriptor||proof?.details?.fileDescriptor;
 N(descriptor,'UNRESOLVED_DEPENDENCY');
 const evidence=documentation(descriptor,true);
 const intentHash=M.hash({section:p.section,unitRef:p.unitRef,fileRevision:p.fileRevision,reason:record.reason,evidence:p.institutionalEvidence,descriptor});
 const registration={...record,...evidence,section:p.section,mode:v.mode,unitRef:p.unitRef};
 const context={...x,administrativeRecord:registration,p:{presentationRef:u.presentationRef,...(!photo?{requirementRevisionRef:rq.currentRevisionRef}:{})}};
 if(!proof){
  const f=p.fileDescriptor;
  const extension=String(f.name).split('.').pop().toLowerCase();
  const mime={pdf:'application/pdf',ppt:'application/vnd.ms-powerpoint',pptx:'application/vnd.openxmlformats-officedocument.presentationml.presentation',odp:'application/vnd.oasis.opendocument.presentation'}[extension];
  N((photo||mime&&mime===f.mime)&&policy.workspace.mimeTypes.includes(f.mime)&&f.size<=policy.workspace.maximumBytes&&/^[a-f0-9]{64}$/.test(f.sha256),'VALIDATION_ERROR','UPLOAD_POLICY');
  context.p.fileDescriptor=f;
  context.emit=(kind,target,policy,details)=>x.emit(kind,target,policy,{...details,administrativeRegistration:{unitRef:p.unitRef,fileRevision:p.fileRevision,section:p.section,intentHash}});
  const reservation=photo?f4ActivitiesPhotoCommands_(context).PHOTO_UPLOAD_REQUEST():f4ActivitiesJobs_(context).ARTIFACT_UPLOAD_REQUEST();
  return {...reservation,provider:f4ActivitiesFileStorage_().configured(policy)};
 }
 N(p.effectRef.kind==='outbox'&&proof.actor?.uid===x.actor.uid&&proof.authorityBasis?.functionBasis?.functionId===x.authorityBasis.functionBasis.functionId&&proof.details?.administrativeRegistration?.intentHash===intentHash&&proof.policyRef?.id===policyRef.id&&proof.policyRevision===policy.revision&&proof.status==='CONFIRMED'&&proof.targetRef?.id===pr.id,'AUTHORIZATION_DENIED','ADMIN_UPLOAD_CONTEXT');
 const artifact=x.load(proof.artifactRef);
 N(f4ActivitiesFileStorage_().isArtifact(artifact)&&artifact.accessClass==='PRIVATE'&&artifact.objectId===proof.workspaceReservation?.fileId&&artifact.sha256===descriptor.sha256,'UNRESOLVED_DEPENDENCY','UPLOAD_PROOF');
 context.p.notes=record.reason;
 if(photo){context.p.effectRef=p.effectRef;return f4ActivitiesPhotoCommands_(context).PHOTO_SUBMIT();}
 const items=[{ref:proof.artifactRef}],manifestRef=M.ref('policy','upl_'+p.effectRef.id);
 N(!x.load(manifestRef,false,true),'CONFLICT');
 x.save(manifestRef,f4ActivitiesUploadManifest_(M,manifestRef,items,proof,p.effectRef));
 context.p.artifactLinkManifestRef=manifestRef;context.p.trustedUploadReceiptRef=p.effectRef;
 return f4ActivitiesMaterials_(context).MATERIAL_SUBMIT();
}
