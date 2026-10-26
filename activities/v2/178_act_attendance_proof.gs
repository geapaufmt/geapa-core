function f4AttendanceProofMime_(name){return {pdf:'application/pdf',jpg:'image/jpeg',jpeg:'image/jpeg',png:'image/png',doc:'application/msword',docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}[String(name).split('.').pop().toLowerCase()];}
// A private supporting file is evidence for a justification, never an approval.
// The existing workspace coordinator owns reservation, upload and reconciliation.
function f4AttendanceProof_(x){
 const {M,N,p,load,save,now,actor}=x;
 function begin(context){
  const policyRef=context.c.proofUploadPolicyRef;N(policyRef?.kind==='attachmentPolicy','CONFIG_CAPABILITY_DISABLED','JUSTIFICATION_UPLOAD_NOT_CONFIGURED');
  const policy=load(policyRef,true),file=p.fileDescriptor;N(policy.approved&&policy.active&&policy.usage==='ATTENDANCE_PROOF'&&policy.workspace?.enabled===true,'CONFIG_CAPABILITY_DISABLED');
  f4Core_().exact(file,['name','mime','size','sha256']);N(typeof file.name==='string'&&file.name.length<=240&&f4AttendanceProofMime_(file.name)&&file.mime===f4AttendanceProofMime_(file.name)&&Number.isSafeInteger(file.size)&&file.size>0&&file.size<=policy.workspace.maximumBytes&&/^[a-f0-9]{64}$/.test(file.sha256),'VALIDATION_ERROR','PROOF_FILE');
  N(policy.workspace.mimeTypes.includes(file.mime),'CONFIG_CAPABILITY_DISABLED');
  const ref=M.ref('outbox',M.hash({operationId:x.operationId,kind:'ATTENDANCE_PROOF'}).slice(0,32));N(!load(ref,false,true),'CONFLICT');
  save(ref,{effectKind:f4ActivitiesFileStorage_().effect(policy),targetRef:p.attendanceRef,activityRef:context.fact.activityRef,ownerPersonRef:actor.personId,authorityScope:'ACTIVITIES',policyRef,policyRevision:policy.revision,details:{fileDescriptor:file,purpose:'ATTENDANCE_PROOF'},actor,authorityBasis:x.authorityBasis,status:'PENDING',state:'PENDING',nextAttemptAt:M.stamp(now),dueAt:M.stamp(now),leaseFence:0,attemptCount:0,createdByOperation:x.operationId});
  return {effectRef:ref};
 }
 function attach(justificationRef,fact){
  const ref=p.proofUploadRef;if(!ref)return {};
  N(ref.kind==='outbox'&&!p.proofLink,'VALIDATION_ERROR');const out=load(ref,true);
  N(f4ActivitiesFileStorage_().isCreate(out)&&out.details?.purpose==='ATTENDANCE_PROOF'&&out.authorityScope==='ACTIVITIES'&&out.targetRef?.id===p.attendanceRef.id&&out.activityRef?.id===fact.activityRef.id&&out.ownerPersonRef===actor.personId&&out.status==='CONFIRMED'&&out.artifactRef?.kind==='activityArtifact'&&!out.justificationRef,'AUTHORIZATION_DENIED','PROOF_CONTEXT');
  const artifact=load(out.artifactRef);N(artifact.attendanceRef?.id===fact.id||artifact.attendanceRef?.id===p.attendanceRef.id,'AUTHORIZATION_DENIED');N(artifact.ownerPersonRef===actor.personId&&artifact.objectId===out.workspaceReservation?.fileId,'UNRESOLVED_DEPENDENCY');
  save(ref,{...out,justificationRef});return {proofArtifactRef:out.artifactRef,proofUploadRef:ref};
 }
 return {begin,attach};
}
function f4AttendanceProofAuthorizeUpload_(reader,out,options){
 const M=f4ActivitiesModel_(),N=M.need,{identity,settings,clock}=options;
 N(settings.commands.includes('ATTENDANCE_PROOF_UPLOAD_REQUEST')&&out.details?.purpose==='ATTENDANCE_PROOF'&&out.authorityScope==='ACTIVITIES'&&out.actor?.uid===identity.uid&&out.ownerPersonRef===out.actor.personId,'AUTHORIZATION_DENIED');
 const provenance=reader('commandProvenance/'+out.createdByOperation),ticket=provenance&&reader('memberAdmissionTickets/'+provenance.ticketId),request=ticket?.requestEnvelope;
 N(provenance?.ownerUid===identity.uid&&request?.commandType==='ACT.ATTENDANCE_PROOF_UPLOAD_REQUEST'&&request.payload.attendanceRef.id===out.targetRef.id&&request.payload.activityRef.id===out.activityRef.id,'AUTHORIZATION_DENIED');
 f4ActivitiesPipeline_().statusAccess(reader,identity,{operationId:out.createdByOperation,commandType:request.commandType,digest:provenance.digest},clock());
 const activity=reader(M.path(out.activityRef));N(activity?.typeSnapshot?.allowsJustification&&!['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(activity.status),'AUTHORIZATION_DENIED','PROOF_CONTEXT_CHANGED');
}
