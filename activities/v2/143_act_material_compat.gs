// Adapts the existing material form to upload + submission. Neither step issues
// the definitive material receipt, which remains a separate review decision.
function f4ActivitiesMaterialCompat_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),ports=options.ports,clock=options.clock||Date.now;
 function begin(input){
  C.exact(input,['operationId','idApresentacao','basis','fileDescriptor']);C.exact(input.basis,['unitRevision','requirementRevision']);
  N(typeof input.idApresentacao==='string'&&input.idApresentacao.startsWith('APU_')&&M.id(input.idApresentacao.slice(4)),'VALIDATION_ERROR');
  const u=db.get('presentationUnitHeads/'+input.idApresentacao.slice(4)),p=u?.presentationRef&&db.get(M.path(u.presentationRef)),rq=p?.materialRequirementRef&&db.get(M.path(p.materialRequirementRef));
  N(u&&p&&rq&&u.revision===input.basis.unitRevision&&rq.revision===input.basis.requirementRevision,'STALE_REVISION','DISPLAYED_MATERIAL_CHANGED');
  const payload={presentationRef:u.presentationRef,requirementRevisionRef:rq.currentRevisionRef,fileDescriptor:input.fileDescriptor};
  if(rq.openCorrectionCount){const rows=db.query({collection:'presentationCorrectionOpportunities',filters:[{field:'presentationRef',op:'EQUAL',value:u.presentationRef}],order:[{field:'__name__',direction:'ASCENDING'}],limit:20}),open=rows.map(r=>r.data).filter(o=>o.status==='OPEN'&&o.notifiedAt&&clock()<=M.ms(o.dueAt));N(rows.length<20&&open.length===1,'HUMAN_DECISION_REQUIRED','CORRECTION_OPPORTUNITY_UNRESOLVED');payload.opportunityRef=M.ref('opportunity',open[0].id);}
  return ports.prepare({operationId:input.operationId,commandType:'ACT.ARTIFACT_UPLOAD_REQUEST',payload});
 }
 function submit(input){
  C.exact(input,['operationId','effectRef','notes']);N(typeof input.notes==='string'&&input.notes.length<=4000&&input.effectRef?.kind==='outbox','VALIDATION_ERROR');
  const sealed=db.transaction(tx=>{const get=p=>tx.get(p).data(),out=get(M.path(input.effectRef));N(out&&!out.details?.purpose,'UNRESOLVED_DEPENDENCY');ports.authorizeUpload(get,out);N(out.status==='CONFIRMED'&&out.artifactRef,'UNRESOLVED_DEPENDENCY','UPLOAD_NOT_CONFIRMED');
   const artifact=get(M.path(out.artifactRef));N(artifact?.presentationRef?.id===out.targetRef.id&&artifact.versionId,'UNRESOLVED_DEPENDENCY');
   const provenance=get('commandProvenance/'+out.createdByOperation),ticket=get('memberAdmissionTickets/'+provenance.ticketId),original=ticket.requestEnvelope.payload;
   const ref=M.ref('policy','upl_'+input.effectRef.id),items=[{ref:out.artifactRef}],manifest={id:ref.id,revision:1,kind:'UPLOAD_ARTIFACT_MANIFEST',state:'SEALED',complete:true,items,digest:M.hash(items),presentationRef:out.targetRef,cycleRef:out.cycleRef,requirementRevisionRef:out.details.requirementRevisionRef,effectRef:input.effectRef};
   const prior=get(M.path(ref));if(prior)N(C.hash(F.encode(prior))===C.hash(F.encode(manifest)),'CONFLICT');else tx.set(M.path(ref),manifest);
   return {ref,presentationRef:out.targetRef,requirementRevisionRef:out.details.requirementRevisionRef,opportunityRef:original.opportunityRef||null};
  });
  const correction=!!sealed.opportunityRef,payload={presentationRef:sealed.presentationRef,requirementRevisionRef:sealed.requirementRevisionRef,trustedUploadReceiptRef:input.effectRef,notes:input.notes,...(correction?{opportunityRef:sealed.opportunityRef,artifactManifestRef:sealed.ref}:{artifactLinkManifestRef:sealed.ref})};
  return ports.prepare({operationId:input.operationId,commandType:correction?'ACT.MATERIAL_CORRECTION_SUBMIT':'ACT.MATERIAL_SUBMIT',payload});
 }
 return {begin,submit};
}
