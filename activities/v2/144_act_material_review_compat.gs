// The existing "Aprovar slide" action can cover the complete requirement only
// when its approved policy explicitly maps that action to all required items.
function f4ActivitiesMaterialReviewCompat_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),ports=options.ports,clock=options.clock||Date.now;
 function context(unitId){
  N(typeof unitId==='string'&&unitId.startsWith('APU_')&&M.id(unitId.slice(4)),'VALIDATION_ERROR');const u=db.get('presentationUnitHeads/'+unitId.slice(4));N(u?.presentationRef,'UNRESOLVED_DEPENDENCY');
  const proof=ports.readAuthority([M.path(M.ref('unit',u.id))],false,'PRESENTATION_BOARD_READ','PRESC_'+u.cycleRef.id),pr=db.get(M.path(u.presentationRef)),rq=pr?.materialRequirementRef&&db.get(M.path(pr.materialRequirementRef)),policy=rq?.uploadPolicyRef&&db.get(M.path(rq.uploadPolicyRef)),rule=policy?.portalReview;
  N(rq&&policy?.approved&&policy.active&&rule?.approved===true&&rule.mode==='COMPLETE_FILE_REVIEW'&&Array.isArray(rule.coveredItemIds)&&rule.coveredItemIds.length>0&&rule.coveredItemIds.length<=32&&new Set(rule.coveredItemIds).size===rule.coveredItemIds.length&&rq.items.length===rule.coveredItemIds.length&&rq.items.every(i=>rule.coveredItemIds.includes(i.itemId)),'CONFIG_CAPABILITY_DISABLED','REVIEW_COVERAGE_NOT_CONFIGURED');
  return {u,pr,rq,policy,rule,proof};
 }
 function seal(z,kind,items,extra,commandFor){
  const ref=M.ref('policy',M.hash({kind,presentation:z.pr.id,requirement:z.rq.currentRevisionRef,items,extra}).slice(0,32)),value={id:ref.id,revision:1,kind,state:'SEALED',complete:true,items,digest:M.hash(items),cycleRef:z.u.cycleRef,presentationRef:z.u.presentationRef,requirementRevisionRef:z.rq.currentRevisionRef,...extra};
  const prepared=ports.prepare(commandFor(ref));
  db.transaction(tx=>{const get=p=>tx.get(p).data();ports.authorizePrepared(get,prepared);f4SharedAuthz_().recheck(get,z.proof.q,z.proof.proof,clock());const rq=get(M.path(z.pr.materialRequirementRef)),policy=get(M.path(z.rq.uploadPolicyRef)),prior=get(M.path(ref));N(rq.revision===z.rq.revision&&C.hash(F.encode(policy))===C.hash(F.encode(z.policy)),'STALE_REVISION');if(prior)N(C.hash(F.encode(prior))===C.hash(F.encode(value)),'CONFLICT');else tx.set(M.path(ref),value);});return ref;
 }
 function prepare(input){
  C.exact(input,['operationId','idApresentacao','basis','decision','publicReason','internalReason','stage',...('priorSelector'in input?['priorSelector']:[]),...('dueCivilDate'in input?['dueCivilDate']:[]),...('opportunityRevision'in input?['opportunityRevision']:[])]);
  C.exact(input.basis,['unitRevision','requirementRevision']);N(['APROVAR','DISPENSAR','SOLICITAR_AJUSTE'].includes(input.decision)&&['REVIEW','RECEIVE','EVALUATE','REQUEST','OPPORTUNITY','TERMS'].includes(input.stage),'VALIDATION_ERROR');
  N(typeof input.publicReason==='string'&&input.publicReason.length<=2000&&typeof input.internalReason==='string'&&input.internalReason.length<=2000,'VALIDATION_ERROR');const z=context(input.idApresentacao);let commandType,payload;const command=(type,payload)=>({operationId:input.operationId,commandType:'ACT.'+type,payload});
  let dueAt;if(input.decision==='SOLICITAR_AJUSTE'){M.civil(input.dueCivilDate);dueAt=C.civilStart(C.nextDay(input.dueCivilDate))-1;N(dueAt>clock()&&input.publicReason.trim(),'INCOMPLETE_REQUEST','CORRECTION_TERMS_REQUIRED');}
  if(input.stage==='TERMS'){
   N(input.decision==='SOLICITAR_AJUSTE'&&z.u.revision===input.basis.unitRevision&&z.rq.revision===input.basis.requirementRevision,'STALE_REVISION');const f=z.rq.currentCorrectionRequestRef&&db.get(M.path(z.rq.currentCorrectionRequestRef)),o=f?.opportunityRef&&db.get(M.path(f.opportunityRef));N(o?.status==='AWAITING_NOTICE'&&o.revision===input.opportunityRevision&&o.presentationRef.id===z.pr.id,'STALE_REVISION');commandType='CORRECTION_TERMS_REVISE';payload={opportunityRef:f.opportunityRef,basisRevision:o.revision,dueAt,reason:input.publicReason,authorityDecision:{kind:'REVISE_UNNOTIFIED_TERMS',reason:input.publicReason,dueCivilDate:input.dueCivilDate}};
  }else if(input.stage==='REVIEW'){
   N(z.u.revision===input.basis.unitRevision&&z.rq.revision===input.basis.requirementRevision,'STALE_REVISION','DISPLAYED_MATERIAL_CHANGED');
   if(input.decision==='DISPENSAR'){
    N(input.publicReason.trim()||input.internalReason.trim(),'INCOMPLETE_REQUEST');N(!z.rq.currentRegularizationRef,'INVALID_STATE_TRANSITION','MATERIAL_ALREADY_REGULARIZED');N(!z.rq.openCorrectionCount,'HUMAN_DECISION_REQUIRED','PENDING_OPPORTUNITY');
    const exceptionalDecision={scope:'FULL_REQUIREMENT',reason:input.publicReason.trim()||input.internalReason.trim(),coveredItemIds:z.rule.coveredItemIds.slice()},evidence=seal(z,'PORTAL_MATERIAL_EXCEPTION_BASIS',exceptionalDecision.coveredItemIds.map(itemId=>({itemId})),{decisionReason:input.publicReason,internalReason:input.internalReason},ref=>command('MATERIAL_EXCEPTION',{presentationRef:z.u.presentationRef,requirementRevisionRef:z.rq.currentRevisionRef,exceptionalDecision,reason:input.publicReason,evidenceManifestRef:ref}));
    commandType='MATERIAL_EXCEPTION';payload={presentationRef:z.u.presentationRef,requirementRevisionRef:z.rq.currentRevisionRef,exceptionalDecision,reason:input.publicReason,evidenceManifestRef:evidence};
   }else{
    N(z.rq.currentSubmissionRef&&!z.rq.currentRegularizationRef,'INVALID_STATE_TRANSITION');const cutoff=z.rule.evaluationCutoffRef&&db.get(M.path(z.rule.evaluationCutoffRef));N(cutoff?.approved,'CONFIG_CAPABILITY_DISABLED','MATERIAL_CUTOFF_REQUIRED');
    const outcome=input.decision==='SOLICITAR_AJUSTE'?'CORRECTABLE_DEFECT':'VALID',coverage=seal(z,'PORTAL_MATERIAL_COVERAGE',outcome==='VALID'?z.rule.coveredItemIds.map(itemId=>({itemId})):[],{completeCoverage:outcome==='VALID'},ref=>command('MATERIAL_REVIEW',{submissionRef:z.rq.currentSubmissionRef,requirementRevisionRef:z.rq.currentRevisionRef,outcome,coveredItemsManifestRef:ref,reason:input.publicReason,internalReason:input.internalReason}));commandType='MATERIAL_REVIEW';payload={submissionRef:z.rq.currentSubmissionRef,requirementRevisionRef:z.rq.currentRevisionRef,outcome,coveredItemsManifestRef:coverage,reason:input.publicReason,internalReason:input.internalReason};
   }
  }else if(input.decision==='SOLICITAR_AJUSTE'){
   N(['REQUEST','OPPORTUNITY'].includes(input.stage)&&input.priorSelector,'VALIDATION_ERROR');const original=ports.resume(input.priorSelector),receipt=db.get('commandReceipts/'+input.priorSelector.operationId),isRequest=input.stage==='REQUEST',expected=isRequest?'ACT.MATERIAL_REVIEW':'ACT.MATERIAL_CORRECTION_REQUEST';N(original.commandType===expected&&receipt?.state==='COMPLETED','UNRESOLVED_DEPENDENCY');
   const eventRef=M.ref(isRequest?'review':'flow',M.hash({operationId:input.priorSelector.operationId,kind:isRequest?'review':'flow',parent:''}).slice(0,32)),event=db.get(M.path(eventRef));N(event?.operationId===input.priorSelector.operationId&&event.presentationRef?.id===z.pr.id,'UNRESOLVED_DEPENDENCY');
   const reviewRef=isRequest?eventRef:event.reviewRef,review=isRequest?event:db.get(M.path(reviewRef));N(review?.outcome==='CORRECTABLE_DEFECT'&&review.submissionRef?.id===z.rq.currentSubmissionRef?.id&&z.rq.currentReviewRef?.id===reviewRef.id,'STALE_REVISION');
   if(isRequest){commandType='MATERIAL_CORRECTION_REQUEST';payload={reviewRef,correctableIssues:[input.publicReason],reason:input.publicReason,suggestedCaseTerms:{dueCivilDate:input.dueCivilDate,dueAt,reason:input.publicReason}};}
   else{N(event.suggestedCaseTerms?.dueAt===dueAt&&event.reason===input.publicReason,'STALE_REVISION');commandType='CORRECTION_OPPORTUNITY_OPEN';payload={correctionRequestRef:eventRef,sourceReviewRef:reviewRef,originalSubmissionRef:review.submissionRef,dueAt,reason:input.publicReason,authorityDecision:{reason:input.publicReason,decisionKind:'INDIVIDUAL_CORRECTION_OPPORTUNITY',dueCivilDate:input.dueCivilDate}};}
  }else{
   N(input.decision==='APROVAR'&&input.priorSelector,'VALIDATION_ERROR');const original=ports.resume(input.priorSelector),receipt=db.get('commandReceipts/'+input.priorSelector.operationId);
   // The protected original request pins the reviewed submission. Client output
   // is never accepted as a reference to an authority-bearing review/receipt.
   const expected=input.stage==='RECEIVE'?'ACT.MATERIAL_REVIEW':'ACT.MATERIAL_RECEIVE';N(original.commandType===expected&&original.payload.submissionRef?.id===z.rq.currentSubmissionRef?.id,'STALE_REVISION');
   const eventKind=input.stage==='RECEIVE'?'review':'receipt',eventRef=M.ref(eventKind,M.hash({operationId:input.priorSelector.operationId,kind:eventKind,parent:''}).slice(0,32)),event=db.get(M.path(eventRef));N(receipt?.state==='COMPLETED'&&event?.operationId===input.priorSelector.operationId&&event.presentationRef.id===z.pr.id,'UNRESOLVED_DEPENDENCY');
   if(input.stage==='RECEIVE'){N(event.outcome==='VALID'&&z.rq.currentReviewRef?.id===eventRef.id,'STALE_REVISION');commandType='MATERIAL_RECEIVE';payload={submissionRef:z.rq.currentSubmissionRef,validReviewRef:eventRef,requirementRevisionRef:z.rq.currentRevisionRef,coverageManifestRef:event.coveredItemsManifestRef};}
   else{N(z.rq.receiptRefs?.some(r=>r.id===eventRef.id),'STALE_REVISION');const receipts=seal(z,'PORTAL_MATERIAL_RECEIPTS',z.rq.receiptRefs.map(ref=>({ref})),{completeCoverage:true},ref=>command('MATERIAL_FINAL_EVALUATE',{presentationRef:z.u.presentationRef,requirementRevisionRef:z.rq.currentRevisionRef,receiptCoverageManifestRef:ref,cutoffRef:z.rule.evaluationCutoffRef}));commandType='MATERIAL_FINAL_EVALUATE';payload={presentationRef:z.u.presentationRef,requirementRevisionRef:z.rq.currentRevisionRef,receiptCoverageManifestRef:receipts,cutoffRef:z.rule.evaluationCutoffRef};}
  }
  return ports.prepare({operationId:input.operationId,commandType:'ACT.'+commandType,payload});
 }
 return {prepare};
}
