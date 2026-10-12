// Resolve only the current submitted artifact. Drive remains private; callers
// receive bytes after current Membros authorization, never owner credentials.
function f4ActivitiesMaterialDownload_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),clock=options.clock||Date.now;
 function read(input){
  C.exact(input,['idApresentacao']);N(typeof input.idApresentacao==='string'&&input.idApresentacao.startsWith('APU_')&&M.id(input.idApresentacao.slice(4)),'VALIDATION_ERROR');
  const seen=new Map(),get=path=>{if(!seen.has(path))seen.set(path,db.get(path));return seen.get(path)},unitPath='presentationUnitHeads/'+input.idApresentacao.slice(4),u=get(unitPath);N(u?.presentationRef,'UNRESOLVED_DEPENDENCY');
  let auth;for(const [self,cap]of [[true,'PRESENTATION_SELF_READ'],[true,'PRESENTATION_MATERIAL_READ_SELF'],[false,'PRESENTATION_BOARD_READ']]){try{auth=options.ports.readAuthority([unitPath],self,cap,self?undefined:'PRESC_'+u.cycleRef.id);break;}catch(e){if(!['AUTHORIZATION_DENIED','M13_AUTHZ_DENIED','AUTH_BINDING_INVALID','ACCESS_REVOKED'].includes(e.code))throw e;}}N(auth,'AUTHORIZATION_DENIED');
  const pr=get(M.path(u.presentationRef)),rq=pr?.materialRequirementRef&&get(M.path(pr.materialRequirementRef)),submission=rq?.currentSubmissionRef&&get(M.path(rq.currentSubmissionRef));N(submission?.presentationRef?.id===pr.id&&submission.artifactRefs?.length===1,'UNRESOLVED_DEPENDENCY');
  const artifact=get(M.path(submission.artifactRefs[0])),evidence=artifact?.evidenceRef&&get(M.path(artifact.evidenceRef)),out=evidence?.effectId&&get('jobOutbox/'+evidence.effectId),policy=out?.policyRef&&get(M.path(out.policyRef));
  N(artifact?.provider==='GOOGLE_DRIVE_V3'&&artifact.targetKind==='FILE'&&artifact.accessClass==='PRIVATE'&&artifact.presentationRef?.id===pr.id&&out?.status==='CONFIRMED'&&out.artifactRef?.id===artifact.id&&out.submissionRef?.id===submission.id&&out.targetRef?.id===pr.id&&out.workspaceReservation?.fileId===artifact.objectId,'UNRESOLVED_DEPENDENCY','MATERIAL_FILE_PROVENANCE');
  N(policy?.approved&&policy.active&&policy.workspace?.enabled===true,'CONFIG_CAPABILITY_DISABLED');
  const plan={effectId:out.id,fileId:artifact.objectId,descriptor:out.workspaceReservation.descriptor};
  N(artifact.sha256===plan.descriptor.sha256&&artifact.md5===plan.descriptor.md5&&artifact.size===plan.descriptor.size&&artifact.mime===plan.descriptor.mime,'UNRESOLVED_DEPENDENCY');
  const provider=options.providerFor?options.providerFor(policy.workspace):f4ActivitiesWorkspaceDrive_({config:options.config,root:f4DevBoundary_().root,settings:policy.workspace}),bytes=provider.download(plan,artifact.versionId);
  // A role revocation or a replaced submission during the download discards it.
  f4SharedAuthz_().recheck(p=>db.get(p),auth.q,auth.proof,clock());for(const [path,before]of seen)N(C.hash(F.encode(before||null))===C.hash(F.encode(db.get(path)||null)),'STALE_REVISION');
  return {ok:true,data:{name:artifact.name,mime:artifact.mime,size:artifact.size,sha256:artifact.sha256,contentBase64:Utilities.base64Encode(bytes)}};
 }
 return {read};
}
