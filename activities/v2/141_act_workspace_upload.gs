// Protected provider preparation and observations. Business state changes still
// run through shared Core commands; bytes never enter the command ledger.
function f4ActivitiesWorkspaceUpload_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),clock=options.clock||Date.now;
 N(f4DevBoundary_().permit(options.config)&&typeof options.authorize==='function'&&typeof options.providerFor==='function','DESTINATION_INVALID');
 const handles=new WeakMap(),core=f4ActivitiesSharedSystemCommands_(options.pipeline,{resolve:h=>handles.get(h)});
 const get=p=>db.get(p),hash=v=>C.hash(F.encode(v));
 function source(ref,reader=get){
  // Deduplicate repeated reads within this one source check. Each subsequent
  // provider/transaction boundary creates a fresh reader and revalidates access.
  const sourceReader=reader,seen=new Map();reader=p=>{if(!seen.has(p))seen.set(p,sourceReader(p));return seen.get(p);};
  N(ref?.kind==='outbox','VALIDATION_ERROR');const out=reader(M.path(ref));N(f4ActivitiesFileStorage_().isCreate(out)&&(out.targetRef?.kind==='presentation'||out.authorityScope==='ACTIVITIES'&&out.details?.purpose==='ATTENDANCE_PROOF'&&out.targetRef?.kind==='attendance'),'VALIDATION_ERROR');options.authorize(reader,out);
  const policy=reader(M.path(out.policyRef)),rule=policy?.executorRuleRef&&reader(M.path(policy.executorRuleRef));
  N(policy?.approved&&policy.active&&policy.revision===out.policyRevision&&rule?.approved&&rule.active&&(out.authorityScope==='ACTIVITIES'?rule.authorityScope==='ACTIVITIES':rule.cycleRef?.id===out.cycleRef.id)&&policy.workspace?.enabled===true,'CONFIG_CAPABILITY_DISABLED','WORKSPACE_NOT_CONFIGURED');
  N(['ACT.OUTBOX_CLAIM','ACT.OUTBOX_OBSERVE','ACT.OUTBOX_RECONCILE'].every(cmd=>rule.commands?.includes(cmd)),'CONFIG_CAPABILITY_DISABLED');
  return {out,policy,rule};
 }
 function commandPlan(ref,command,payload){
  const {out,policy}=source(ref),observe=command==='OUTBOX_OBSERVE',operationId=M.hash({effect:ref.id,revision:out.revision,command,payload}).slice(0,32);
  const principal={kind:observe?'WORKSPACE_EXECUTOR':'SYSTEM_JOB',servicePrincipal:'CORE_ACT_DEV_WORKSPACE',scope:out.authorityScope==='ACTIVITIES'?'ACTIVITIES':'PRESC_'+out.cycleRef.id,ruleRef:policy.executorRuleRef,jobKind:observe?'WORKSPACE':command==='OUTBOX_CLAIM'?'OUTBOX_DISPATCH':'OUTBOX_RECONCILE',runId:operationId};
  return {principal,request:{schemaVersion:'geapa-command-v2-candidate',operationId,commandType:'ACT.'+command,payload,scope:{targetRefs:[ref]},expected:[{...ref,revision:out.revision}]}};
 }
 function command(ref,type,payload){const plan=commandPlan(ref,type,payload),handle={};handles.set(handle,plan.principal);return core.execute(handle,plan.request);}
 function reservedPlan(ref,out){N(out.workspaceReservation?.provider==='GOOGLE_DRIVE_V3','UNRESOLVED_DEPENDENCY','FILE_RESERVATION_REQUIRED');return {effectId:ref.id,fileId:out.workspaceReservation.fileId,descriptor:out.workspaceReservation.descriptor};}
 function reservationInput(ref,bytes){
  const first=source(ref),descriptor=first.out.details.fileDescriptor;N(first.out.effectKind==='DRIVE_CREATE'&&f4ActivitiesFileStorage_().configured(first.policy)==='GOOGLE_DRIVE_V3','EXECUTOR_FENCE_CONFLICT');N(descriptor&&/^[a-f0-9]{64}$/.test(descriptor.sha256)&&Array.isArray(bytes)&&bytes.length===descriptor.size,'VALIDATION_ERROR','UPLOAD_DESCRIPTOR_REQUIRED');
  N(first.policy.workspace.mimeTypes.includes(descriptor.mime)&&descriptor.size<=first.policy.workspace.maximumBytes&&typeof descriptor.name==='string'&&descriptor.name.trim()&&!/[\x00-\x1f]/.test(descriptor.name),'VALIDATION_ERROR','UPLOAD_POLICY');
  const digest=alg=>Utilities.computeDigest(alg,bytes).map(b=>(b&255).toString(16).padStart(2,'0')).join('');N(digest(Utilities.DigestAlgorithm.SHA_256)===descriptor.sha256,'VALIDATION_ERROR','UPLOAD_CONTENT_CHANGED');
  const complete={...descriptor,md5:digest(Utilities.DigestAlgorithm.MD5)};return {first,descriptor,complete};
 }
 function reservationCommit(ref,input,fileId){const {first,descriptor,complete}=input;N(typeof fileId==='string'&&/^[A-Za-z0-9_-]{10,160}$/.test(fileId),'VALIDATION_ERROR');
  return db.transaction(tx=>{const reader=p=>tx.get(p).data(),fresh=source(ref,reader),out=fresh.out;N(hash(fresh.policy.workspace)===hash(first.policy.workspace)&&hash(out.details.fileDescriptor)===hash(descriptor),'STALE_REVISION');
   if(out.workspaceReservation){N(hash(out.workspaceReservation.descriptor)===hash(complete)&&out.workspaceReservation.folderId===fresh.policy.workspace.folderId,'CONFLICT');return out;}
   N(['PENDING','KNOWN_NO_EFFECT'].includes(out.status),'INVALID_STATE_TRANSITION');const next={...out,revision:out.revision+1,workspaceReservation:{provider:'GOOGLE_DRIVE_V3',fileId,folderId:fresh.policy.workspace.folderId,descriptor:complete,reservedAt:M.stamp(clock())}};tx.set(M.path(ref),next);return next;
  });
 }
 function reserve(ref,bytes){const input=reservationInput(ref,bytes),fileId=input.first.out.workspaceReservation?.fileId||options.providerFor(input.first.policy.workspace).reserveId();return reservationCommit(ref,input,fileId);}
 function attempt(ref){return db.transaction(tx=>{const reader=p=>tx.get(p).data(),{out}=source(ref,reader);N(out.status==='CLAIMED','INVALID_STATE_TRANSITION');const path=M.path(ref)+'/attempts/'+out.attemptId,old=reader(path);if(old){N(old.effectId===ref.id&&old.leaseFence===out.leaseFence,'CONFLICT');return {out,execute:false};}
  const execute=M.ms(out.leaseUntil)>clock();tx.set(path,{effectId:ref.id,attemptId:out.attemptId,leaseFence:out.leaseFence,fileId:out.workspaceReservation.fileId,startedAt:M.stamp(clock()),state:execute?'EXECUTING':'EXPIRED_BEFORE_EXECUTION'});return {out,execute};
 });}
 function evidence(ref,out,result){
  N(['CONFIRMED','UNKNOWN'].includes(result.outcome),'UNRESOLVED_DEPENDENCY');const data={...result,effectId:ref.id,attemptId:out.attemptId,leaseFence:out.leaseFence,...(out.cycleRef?{cycleRef:out.cycleRef}:{authorityScope:'ACTIVITIES'}),conclusive:result.outcome==='CONFIRMED'},id=M.hash({kind:'DRIVE_OBSERVATION',data}).slice(0,32),proof=M.ref('reconciliation',id);
  db.transaction(tx=>{const reader=p=>tx.get(p).data(),fresh=source(ref,reader),prior=reader(M.path(proof));N(fresh.out.attemptId===out.attemptId&&fresh.out.leaseFence===out.leaseFence,'CONFLICT');if(prior){N(hash(prior.observation)===hash(data),'CONFLICT');return;}tx.set(M.path(proof),{...data,id,revision:1,observation:data,recordedAt:M.stamp(clock())});});return proof;
 }
 function observePlan(ref,out,result){
  const current=source(ref).out;if(current.status==='CONFIRMED')return {result:{state:'CONFIRMED',artifactRef:current.artifactRef,effectRef:ref}};
  if(result.outcome==='UNKNOWN'&&current.status==='UNKNOWN')return {result:{state:'UNKNOWN',effectRef:ref}};
  const proof=evidence(ref,out,result);
  if(current.status==='UNKNOWN')return {command:commandPlan(ref,'OUTBOX_RECONCILE',{effectRef:ref,attemptId:out.attemptId,reconciliationEvidenceRef:proof,provenOutcome:result.outcome})};
  return {command:commandPlan(ref,'OUTBOX_OBSERVE',{effectRef:ref,attemptId:out.attemptId,leaseFence:out.leaseFence,outcome:result.outcome,evidenceRef:proof,...(result.externalId?{externalId:result.externalId}:{})})};
 }
 function observe(ref,out,result){const plan=observePlan(ref,out,result);if(plan.result)return plan.result;const handle={};handles.set(handle,plan.command.principal);core.execute(handle,plan.command.request);
  const final=source(ref).out;return {state:final.status,effectRef:ref,...(final.artifactRef?{artifactRef:final.artifactRef}:{})};
 }
 function reconcile(ref){const {out,policy}=source(ref);if(out.status==='CONFIRMED')return {state:'CONFIRMED',effectRef:ref,artifactRef:out.artifactRef};N(['CLAIMED','UNKNOWN'].includes(out.status),'INVALID_STATE_TRANSITION');const result=options.providerFor(policy.workspace).inspect(reservedPlan(ref,out));if(result.outcome==='UNKNOWN'&&out.status==='CLAIMED'&&M.ms(out.leaseUntil)>clock())return {state:'PENDING',effectRef:ref};if(out.status==='CLAIMED')attempt(ref);return observe(ref,out,result);}
 function upload(ref,bytes){
  let out=reserve(ref,bytes);if(out.status==='CONFIRMED')return {state:'CONFIRMED',effectRef:ref,artifactRef:out.artifactRef};if(out.status==='UNKNOWN')return reconcile(ref);
  if(['PENDING','KNOWN_NO_EFFECT'].includes(out.status))command(ref,'OUTBOX_CLAIM',{effectRef:ref,expectedAttemptState:out.status});
  const a=attempt(ref);if(!a.execute)return reconcile(ref);const {policy}=source(ref),provider=options.providerFor(policy.workspace);let result;try{result=provider.upload(reservedPlan(ref,a.out),bytes);}catch(e){if(['DESTINATION_INVALID','VALIDATION_ERROR','CONFIG_CAPABILITY_DISABLED'].includes(e.code))throw e;result={outcome:'UNKNOWN',reason:'PROVIDER_RESULT_UNRESOLVED'};}return observe(ref,a.out,result);
 }
 return {upload,reconcile,phases:Object.freeze({source,reservationInput,reservationCommit,reservedPlan,commandPlan,attempt,evidence,observePlan})};
}
