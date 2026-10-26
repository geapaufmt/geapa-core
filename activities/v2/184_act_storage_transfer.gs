// Durable phases for the Storage transport. No network calls and no credentials
// enter the domain. All phases revalidate the original authenticated upload.
function f4ActivitiesStorageTransfer_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),clock=options.clock||Date.now;
 const base=f4ActivitiesWorkspaceUpload_(options).phases,hash=v=>C.hash(F.encode(v));
 function source(ref,get=p=>db.get(p)){
  const z=base.source(ref,get);N(z.out.effectKind==='FILE_CREATE'&&f4ActivitiesFileStorage_().configured(z.policy)==='GOOGLE_CLOUD_STORAGE','CONFIG_CAPABILITY_DISABLED');
  N(!z.out.dispatchDisabled&&['PENDING','CLAIMED','UNKNOWN','CONFIRMED'].includes(z.out.status),'INVALID_STATE_TRANSITION');
  const d=z.out.details.fileDescriptor,w=z.policy.workspace;
  C.exact(d,['name','mime','size','sha256']);N(Number.isSafeInteger(d.size)&&d.size>0&&d.size<=w.maximumBytes&&d.size<=50*1024*1024&&w.mimeTypes.includes(d.mime)&&/^[a-f0-9]{64}$/.test(d.sha256),'VALIDATION_ERROR');
  if(z.out.workspaceReservation)N(z.out.workspaceReservation.provider==='GOOGLE_CLOUD_STORAGE'&&z.out.workspaceReservation.bucket===f4ActivitiesFileStorage_().bucket&&z.out.workspaceReservation.fileId===f4ExecutionEnvironment_().prefix+ref.id&&z.out.workspaceReservation.executor===f4CommandExecutorId_(),'EXECUTOR_FENCE_CONFLICT');return z;
 }
 function reservation(ref){return db.transaction(tx=>{const get=p=>tx.get(p).data(),z=source(ref,get),out=z.out;if(out.workspaceReservation)return out;
  N(out.status==='PENDING'&&out.attemptCount===0&&out.leaseFence===0,'INVALID_STATE_TRANSITION');
  const next={...out,revision:out.revision+1,workspaceReservation:{provider:'GOOGLE_CLOUD_STORAGE',bucket:f4ActivitiesFileStorage_().bucket,fileId:f4ExecutionEnvironment_().prefix+ref.id,executor:f4CommandExecutorId_(),descriptor:out.details.fileDescriptor,reservedAt:M.stamp(clock())}};
  tx.set(M.path(ref),next);return next;
 });}
 function transportPlan(ref){const {out}=source(ref);N(out.workspaceReservation,'UNRESOLVED_DEPENDENCY');return {effectId:ref.id,objectName:out.workspaceReservation.fileId,descriptor:out.details.fileDescriptor};}
 function acquire(ref,requestId){N(/^[a-f0-9]{32}$/.test(requestId),'VALIDATION_ERROR');return db.transaction(tx=>{const get=p=>tx.get(p).data(),{out}=source(ref,get);if(out.status==='CONFIRMED')return {status:out.status,artifactRef:out.artifactRef};
  let r=out.workspaceReservation;if(!r){N(out.status==='PENDING'&&out.attemptCount===0&&out.leaseFence===0,'INVALID_STATE_TRANSITION');r={provider:'GOOGLE_CLOUD_STORAGE',bucket:f4ActivitiesFileStorage_().bucket,fileId:f4ExecutionEnvironment_().prefix+ref.id,executor:f4CommandExecutorId_(),descriptor:out.details.fileDescriptor,reservedAt:M.stamp(clock())};}
  const old=r.transferLease;N(!old||old.until<=clock(),'CONFLICT','TRANSFER_IN_PROGRESS');N((r.transferCalls||0)<100,'SCOPE_TOO_LARGE','TRANSFER_BUDGET');const lease={requestId,fence:(old?.fence||0)+1,until:clock()+55000};
  // Transport metadata uses the Firestore CAS, not the business revision. A
  // retry must reproduce the same original system command key after a lost reply.
  const next={...out,revision:out.revision+(out.workspaceReservation?0:1),workspaceReservation:{...r,transferCalls:(r.transferCalls||0)+1,transferLease:lease}};tx.set(M.path(ref),next);return {lease,reservation:next.workspaceReservation,status:next.status,plan:{effectId:ref.id,objectName:r.fileId,descriptor:out.details.fileDescriptor}};
 });}
 function saveSession(ref,lease,sessionUri){
  N(typeof sessionUri==='string'&&sessionUri.startsWith('https://storage.googleapis.com/upload/storage/v1/b/'+f4ActivitiesFileStorage_().bucket+'/o?')&&sessionUri.length<=4096,'DESTINATION_INVALID');
  return mutate(ref,lease,(out,r)=>{N(!r.sessionUri,'CONFLICT');return {...r,sessionUri,sessionCreatedAt:clock()};});
 }
 function initiating(ref,lease){return mutate(ref,lease,(out,r)=>{N(!r.sessionUri&&!r.initiationStartedAt,'CONFLICT','SESSION_RECONCILIATION_REQUIRED');return {...r,initiationStartedAt:clock()};});}
 function mutate(ref,lease,change){return db.transaction(tx=>{const get=p=>tx.get(p).data(),{out}=source(ref,get),r=out.workspaceReservation;N(r?.transferLease?.requestId===lease.requestId&&r.transferLease.fence===lease.fence&&clock()<r.transferLease.until,'CONFLICT');const next={...out,workspaceReservation:change(out,r)};tx.set(M.path(ref),next);return next;});}
 function release(ref,lease){return mutate(ref,lease,(out,r)=>({...r,transferLease:{...r.transferLease,until:0}}));}
 function digest(ref,result){return db.transaction(tx=>{const get=p=>tx.get(p).data(),{out}=source(ref,get);N(result.outcome==='CONFIRMED','UNRESOLVED_DEPENDENCY');f4ActivitiesFileStorage_().reserved(out,result,result.externalId);const r=out.workspaceReservation;
  if(r.descriptor.md5){N(r.descriptor.md5===result.md5&&r.generation===result.versionId,'CONFLICT');return out;}
  const next={...out,workspaceReservation:{...r,descriptor:{...r.descriptor,md5:result.md5},generation:result.versionId}};tx.set(M.path(ref),next);return next;
 });}
 // A cancelled untouched synthetic effect is retained, never marked uploaded.
 // This helper is operator-only; it is not exported by any HTTP action.
 function cancelUnattempted(ref,expectedRevision,evidenceId){return db.transaction(tx=>{const get=p=>tx.get(p).data(),out=get(M.path(ref)),owner=out&&get('commandExecutorOwners/'+out.createdByOperation),receipt=out&&get('commandReceipts/'+out.createdByOperation),route=owner?.routeId&&get('commandExecutorRoutes/'+owner.routeId),scope=route?.scopeId&&get('actFlowExecutorScopes/'+route.scopeId);
  N(out?.effectKind==='DRIVE_CREATE'&&out.revision===expectedRevision&&out.status==='PENDING'&&out.attemptCount===0&&out.leaseFence===0&&!out.workspaceReservation&&!out.artifactRef&&receipt?.state==='COMPLETED'&&scope?.synthetic===true&&scope.projectId==='geapa-convergencia-dev'&&scope.resourcePaths.includes(M.path(ref)),'INVALID_STATE_TRANSITION');
  N(M.id(evidenceId),'VALIDATION_ERROR');const path=M.path(M.ref('reconciliation',evidenceId));N(!get(path),'CONFLICT');
  tx.set(path,{id:evidenceId,revision:1,kind:'SYNTHETIC_PROVIDER_SWITCH',effectRef:ref,createdByOperation:out.createdByOperation,outcome:'KNOWN_NO_EFFECT',basis:{attemptCount:0,leaseFence:0,reservationAbsent:true},recordedAt:M.stamp(clock())});
  tx.set(M.path(ref),{...out,revision:out.revision+1,status:'CANCELLED',state:'CANCELLED',dispatchDisabled:true,reconciliationRef:{kind:'reconciliation',id:evidenceId},cancelledAt:M.stamp(clock()),cancelReason:'SYNTHETIC_PROVIDER_SWITCH'});return {state:'CANCELLED',effectRef:ref};
 });}
 return {source,reservation,transportPlan,acquire,initiating,saveSession,release,digest,cancelUnattempted};
}
