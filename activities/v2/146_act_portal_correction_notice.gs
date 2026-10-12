// An explicit acknowledgment is a fact supplied by the authenticated presenter.
// Core records it only under an approved Portal notice policy. Fetching a card,
// queuing a message or a scheduler run never counts as acknowledgment.
function f4ActivitiesPortalCorrectionNotice_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),clock=options.clock||Date.now;
 function acknowledge(input){
  C.exact(input,['idApresentacao','opportunityId','opportunityRevision','dueAt','acknowledged']);N(input.acknowledged===true&&typeof input.idApresentacao==='string'&&input.idApresentacao.startsWith('APU_')&&M.id(input.idApresentacao.slice(4))&&M.id(input.opportunityId)&&Number.isSafeInteger(input.opportunityRevision)&&input.opportunityRevision>0&&Number.isSafeInteger(input.dueAt),'VALIDATION_ERROR');
  const unitRef=M.ref('unit',input.idApresentacao.slice(4)),opportunityRef=M.ref('opportunity',input.opportunityId),auth=options.ports.readAuthority([M.path(unitRef)],true,'PRESENTATION_MATERIAL_NOTICE_SELF');
  function source(get){
   const u=get(M.path(unitRef)),pr=u?.presentationRef&&get(M.path(u.presentationRef)),rq=pr?.materialRequirementRef&&get(M.path(pr.materialRequirementRef)),o=get(M.path(opportunityRef)),policy=rq?.uploadPolicyRef&&get(M.path(rq.uploadPolicyRef)),notice=policy?.portalCorrectionNotice,rule=notice?.executorRuleRef&&get(M.path(notice.executorRuleRef));
   N(u?.personRef?.id===auth.personId&&pr?.personRef?.id===auth.personId&&o?.personRef?.id===auth.personId&&o.presentationRef.id===pr.id&&['AWAITING_NOTICE','OPEN'].includes(o.status)&&!rq.currentRegularizationRef,'AUTHORIZATION_DENIED');
   N(M.ms(o.dueAt)===input.dueAt&&clock()<input.dueAt,'DEADLINE_VIOLATION');N(policy?.approved&&policy.active&&notice?.approved&&notice.channel==='PORTAL_EXPLICIT_ACK'&&rule?.approved&&rule.active&&rule.cycleRef?.id===u.cycleRef.id&&rule.commands?.includes('ACT.CORRECTION_NOTICE_RECORD'),'CONFIG_CAPABILITY_DISABLED');return {u,pr,rq,o,policy,notice,rule};
  }
  const operationId=M.hash({kind:'PORTAL_CORRECTION_ACK',unitRef,opportunityRef,personId:auth.personId,revision:input.opportunityRevision,dueAt:input.dueAt}).slice(0,32),evidenceRef=M.ref('reconciliation',operationId);
  const accepted=db.transaction(tx=>{const get=p=>tx.get(p).data();f4SharedAuthz_().recheck(get,auth.q,auth.proof,clock());const z=source(get),prior=get(M.path(evidenceRef));
   if(prior){N(prior.kind==='PORTAL_EXPLICIT_ACK'&&prior.targetRef?.id===opportunityRef.id&&prior.personRef?.id===auth.personId&&prior.opportunityRevision===input.opportunityRevision&&prior.dueAt===input.dueAt,'CONFLICT');return {proof:prior,ruleRef:z.notice.executorRuleRef,cycleRef:z.u.cycleRef};}
   N(z.o.status==='AWAITING_NOTICE'&&z.o.revision===input.opportunityRevision,'STALE_REVISION','CORRECTION_NOTICE_CHANGED');
   const proof={id:evidenceRef.id,revision:1,kind:'PORTAL_EXPLICIT_ACK',status:'CONFIRMED',targetRef:opportunityRef,personRef:z.u.personRef,cycleRef:z.u.cycleRef,opportunityRevision:input.opportunityRevision,sourceReviewRef:z.o.sourceReviewRef,dueAt:input.dueAt,observedAt:clock(),actorAuthorityBasis:auth.proof,expected:[{...opportunityRef,revision:z.o.revision},{...z.pr.materialRequirementRef,revision:z.rq.revision},{...z.u.presentationRef,revision:z.pr.revision}],policyRef:z.rq.uploadPolicyRef,policyRevision:z.policy.revision};
   tx.set(M.path(evidenceRef),proof);return {proof,ruleRef:z.notice.executorRuleRef,cycleRef:z.u.cycleRef};
  });
  const handle={},core=f4ActivitiesSharedSystemCommands_(options.pipeline,{resolve:h=>h===handle?{kind:'WORKSPACE_EXECUTOR',servicePrincipal:'CORE_ACT_PORTAL_NOTICE',scope:'PRESC_'+accepted.cycleRef.id,ruleRef:accepted.ruleRef,runId:operationId}:null});
  core.execute(handle,{schemaVersion:'geapa-command-v2-candidate',operationId,commandType:'ACT.CORRECTION_NOTICE_RECORD',payload:{opportunityRef,notificationEvidenceRef:evidenceRef,observedAt:accepted.proof.observedAt,outcome:'CONFIRMED'},scope:{targetRefs:[opportunityRef,evidenceRef]},expected:accepted.proof.expected});
  f4SharedAuthz_().recheck(p=>db.get(p),auth.q,auth.proof,clock());const z=source(p=>db.get(p));return {ok:true,basis:{unitRevision:z.u.revision,requirementRevision:z.rq.revision},opportunityId:z.o.id};
 }
 return {acknowledge};
}
