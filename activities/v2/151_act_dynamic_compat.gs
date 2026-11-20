// Optional dynamic format on the existing presentation card, using common commands.
function f4ActivitiesDynamicCompat_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),ports=options.ports,clock=options.clock||Date.now;
 function context(id){
  N(typeof id==='string'&&id.startsWith('APU_')&&M.id(id.slice(4)),'VALIDATION_ERROR');const docs=new Map(),get=r=>{if(!r||r.kind==='NONE')return null;const p=typeof r==='string'?r:M.path(r);if(!docs.has(p))docs.set(p,db.get(p));return docs.get(p);},unitRef=M.ref('unit',id.slice(4)),u=get(unitRef);N(u?.cycleRef,'UNRESOLVED_DEPENDENCY');
  let self,board;for(const own of [true,false])try{const proof=ports.readAuthority([M.path(unitRef)],own,own?'PRESENTATION_SELF_READ':'PRESENTATION_BOARD_READ',own?undefined:'PRESC_'+u.cycleRef.id);if(own)self=proof;else board=proof;}catch(e){if(!['AUTHORIZATION_DENIED','AUTH_BINDING_INVALID','ACCESS_REVOKED','M13_AUTHZ_DENIED'].includes(e.code))throw e;}
  N(self||board,'AUTHORIZATION_DENIED');const cycle=get(u.cycleRef),cfg=get(cycle.currentConfigRef),assignment=get(u.assignmentRef),session=get(assignment?.sessionRef),presentation=get(u.presentationRef),proposal=get(u.currentDynamicProposalRef),occupancy=get(u.dynamicOccupancyRef),template=get(cfg.parameters.dynamicMaterialRequirementRef),capacityHead=get(M.ref('capacity',u.cycleRef.id)),assessment=get(capacityHead?.currentAssessmentRef);
  return {get,docs,unitRef,u,cycle,cfg,assignment,session,presentation,proposal,occupancy,template,capacityHead,assessment,self,board};
 }
 const same=(a,b)=>C.hash(F.encode(a||null))===C.hash(F.encode(b||null));
 function recheck(z,get=p=>db.get(p)){for(const [path,value]of z.docs)N(same(value,get(path)),'STALE_REVISION','DYNAMIC_CONTEXT_CHANGED');for(const proof of [z.self,z.board].filter(Boolean))f4SharedAuthz_().recheck(get,proof.q,proof.proof,clock());}
 function read(input){
  C.exact(input,['idApresentacao']);const z=context(input.idApresentacao),history=[],seen=new Set();let ref=z.u.currentDynamicProposalRef;
  while(ref&&ref.kind!=='NONE'){N(ref.kind==='flow'&&!seen.has(ref.id)&&history.length<50,'SCOPE_TOO_LARGE');seen.add(ref.id);const f=z.get(ref),d=z.get(f?.decisionRef);N(f?.kind==='DYNAMIC_PROPOSAL'&&f.unitRef?.id===z.u.id&&f.personRef?.id===z.u.personRef.id,'UNRESOLVED_DEPENDENCY');history.push({revision:f.revision,estado:f.state,elementos:f.elements,dataSolicitada:f.requestedCivilDate||'',enviadaEm:f.submittedAt?new Date(M.ms(f.submittedAt)).toISOString():'',decisao:d?.decision||'',motivo:d?.reason||''});ref=f.predecessorRef;}
  const enabled=z.cycle.status==='OPEN'&&z.cfg.approved&&z.cfg.parameters.dynamic===true,lead=z.cfg.parameters.dynamicRequestLeadDays??30,onTime=z.session&&(z.proposal?.requestedCivilDate===z.session.civilDate&&M.gap(M.day(clock()),z.session.civilDate)>=0||Number.isSafeInteger(lead)&&lead>=30&&M.gap(M.day(clock()),z.session.civilDate)>=lead),canSubmit=!!(z.self&&enabled&&onTime&&!z.presentation?.realizationRef&&z.proposal?.state!=='OPEN');
  recheck(z);return {ok:true,data:{unitRevision:z.u.revision,proposalRevision:z.proposal?.revision||0,habilitada:enabled,podeSolicitar:canSubmit,podeDecidir:!!(z.board&&enabled&&z.proposal?.state==='OPEN'),podeCancelar:!!((z.self||z.board)&&z.occupancy?.state==='ACTIVE'&&!z.presentation?.realizationRef),podeRevogar:!!z.board,ocupacao:z.occupancy?.state||'',motivoCancelamento:z.occupancy?.releaseDecision?.reason||'',materiaisObrigatorios:(z.template?.items||[]).map(i=>({id:i.itemId,rotulo:i.label||i.itemId})),materialConfigurado:!!(z.template?.approved&&z.template?.state==='SEALED'&&z.template?.complete),historico:history}};
 }
 function prepare(input){
  C.exact(input,['operationId','route','payload','basis']);N(input.route==='/v2/apresentacoes/dinamica','VALIDATION_ERROR');C.exact(input.basis,['unitRevision','thematicRevision']);const p=input.payload,z=context(p.idApresentacao);N(z.u.revision===p.unitRevision&&z.u.revision===input.basis.unitRevision,'STALE_REVISION','DISPLAYED_DYNAMIC_CHANGED');
  let type,payload,manifestToSeal;
  if(['SOLICITAR','DECIDIR'].includes(p.acao))N(z.cycle.status==='OPEN'&&z.cfg.approved&&z.cfg.parameters.dynamic===true,'CONFIG_CAPABILITY_DISABLED');
  if(p.acao==='SOLICITAR'){
   C.exact(p,['idApresentacao','acao','unitRevision','elementos']);N(z.self,'AUTHORIZATION_DENIED');type='DYNAMIC_SUBMIT';payload={unitRef:z.unitRef,elements:p.elementos,...(z.u.currentDynamicProposalRef?{predecessorRef:z.u.currentDynamicProposalRef}:{})};
  }else if(p.acao==='DECIDIR'){
   C.exact(p,['idApresentacao','acao','unitRevision','proposalRevision','decisao','motivo','confirmouMateriais']);N(z.board,'AUTHORIZATION_DENIED');N(z.proposal?.revision===p.proposalRevision,'STALE_REVISION');type='DYNAMIC_DECIDE';payload={proposalRef:z.u.currentDynamicProposalRef,decision:p.decisao,reason:p.motivo,assessmentRef:z.capacityHead?.currentAssessmentRef||z.cycle.policyVersionRef||z.cfg.parameters.clockPolicyRef};
   N(payload.assessmentRef,'CONFIG_CAPABILITY_DISABLED','DYNAMIC_DECISION_POLICY');
   if(p.decisao==='APPROVE'){
    N(p.confirmouMateriais===true,'INCOMPLETE_REQUEST','CONFIRM_DYNAMIC_MATERIALS');N(z.assessment&&z.capacityHead.currentAssessmentRef&&z.template?.approved&&z.template.state==='SEALED'&&z.template.complete&&z.template.digest===M.hash(z.template.items)&&z.template.items.length>0&&z.presentation,'CONFIG_CAPABILITY_DISABLED','DYNAMIC_APPROVAL_CONFIGURATION');
    const policy=z.get(z.template.uploadPolicyRef),review=policy?.portalReview;N(policy?.approved&&policy.active&&review?.approved&&review.mode==='COMPLETE_FILE_REVIEW'&&Array.isArray(review.coveredItemIds)&&review.coveredItemIds.length===z.template.items.length&&z.template.items.every(i=>review.coveredItemIds.includes(i.itemId)),'CONFIG_CAPABILITY_DISABLED','DYNAMIC_REVIEW_CONFIGURATION');
    const ref=M.ref('policy',M.hash({kind:'PORTAL_DYNAMIC_MATERIALS',operationId:input.operationId}).slice(0,32)),basisRevisionVector=[{ref:z.unitRef,revision:z.u.revision},{ref:z.cfg.parameters.dynamicMaterialRequirementRef,revision:z.template.revision},{ref:z.template.uploadPolicyRef,revision:policy.revision}];
    const value={id:ref.id,revision:1,kind:'PORTAL_DYNAMIC_MATERIALS',state:'SEALED',complete:true,items:z.template.items,digest:z.template.digest,presentationRef:z.u.presentationRef,cycleRef:z.u.cycleRef,finalEndExclusive:z.template.finalEndExclusive,uploadPolicyRef:z.template.uploadPolicyRef,...(z.template.operatingHours?{operatingHours:z.template.operatingHours}:{}),...(z.template.finalCivilDate?{finalCivilDate:z.template.finalCivilDate}:{}),basisRevisionVector};payload.requirementManifestRef=ref;manifestToSeal={ref,value};
   }
  }else{
   C.exact(p,['idApresentacao','acao','unitRevision','motivo']);N(['DESISTIR','REVOGAR'].includes(p.acao)&&z.occupancy,'VALIDATION_ERROR');N(p.acao==='REVOGAR'?z.board:z.self,'AUTHORIZATION_DENIED');type='DYNAMIC_CANCEL';payload={occupancyRef:z.u.dynamicOccupancyRef,cause:p.acao==='DESISTIR'?'WITHDRAWAL':'REVOCATION',decisionBasis:{reasonCode:'PORTAL_DYNAMIC_CANCELLATION',reason:p.motivo}};
  }
  recheck(z);const command={operationId:input.operationId,commandType:'ACT.'+type,payload},prepared=ports.prepare(command);
  if(manifestToSeal)db.transaction(tx=>{const get=p=>tx.get(p).data();ports.authorizePrepared(get,prepared);recheck(z,get);const path=M.path(manifestToSeal.ref),prior=get(path);if(prior)N(same(prior,manifestToSeal.value),'CONFLICT');else tx.set(path,manifestToSeal.value);});
  return manifestToSeal?ports.prepare(command):prepared;
 }
 return {read,prepare};
}
