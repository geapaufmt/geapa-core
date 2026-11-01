// Internal Core factory. Transport verifies identity and supplies opaque principal handles.
// No web endpoint, trigger, deployment or remote initialization is performed by this file.
function f4ActivitiesCreate_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),contracts=f4ActivitiesContracts_();
 const db=options.db,clock=options.clock||Date.now,authority=options.authority||f4ActivitiesAuthority_();
 N(db&&typeof db.transaction==='function'&&options.principals&&typeof options.principals.resolve==='function','UNRESOLVED_DEPENDENCY','CORE_SERVICES');
 function envelope(q){
  N(q&&Object.getPrototypeOf(q)===Object.prototype);const fields=['schemaVersion','operationId','commandType','payload','expected','scope'];
  N(fields.every(k=>k in q)&&Object.keys(q).every(k=>fields.includes(k)||k==='clientContext'),'VALIDATION_ERROR','ENVELOPE_FIELDS');
  N(q.schemaVersion==='geapa-command-v2-candidate'&&/^[a-f0-9]{32}$/.test(q.operationId));
  const cmd=typeof q.commandType==='string'&&q.commandType.startsWith('ACT.')?q.commandType.slice(4):'',ct=contracts[cmd];N(ct,'VALIDATION_ERROR','COMMAND');
  N(q.payload&&Object.getPrototypeOf(q.payload)===Object.prototype);const allowed=ct.fields.map(k=>k.replace(/^\?/,''));
  N(Object.keys(q.payload).every(k=>allowed.includes(k)),'VALIDATION_ERROR','PAYLOAD_UNKNOWN');N(ct.fields.filter(k=>k[0]!=='?').every(k=>k in q.payload),'INCOMPLETE_REQUEST');
  const refTypes={attendancePolicyRef:'attendancePolicy',rosterRef:'attendanceRoster',activityRef:'activity',typeRef:'activityType',periodRef:'frequencyPeriod',attendanceRef:'attendance',membershipRef:'membership',justificationRef:'justification',personRef:'person',cycleRef:'cycle',policyRef:'policy',approvedConfigRef:'config',unitRef:'unit',unitRefs:'unit',twoUnitRefs:'unit',sessionRef:'session',destinationSessionRef:'session',assignmentRef:'assignment',requestRef:'thematic',flowCaseRef:'flow',caseRef:'flow',correctionRequestRef:'flow',correctionCaseRef:'flow',proposalRef:'flow',presentationRef:'presentation',presentationRefs:'presentation',submissionRef:'submission',originalSubmissionRef:'submission',reviewRef:'review',validReviewRef:'review',sourceReviewRef:'review',opportunityRef:'opportunity',effectRef:'outbox',artifactRef:'artifact',publicationRef:'publication',snapshotRef:'snapshot',obligationRef:'obligation',membersEvidenceRef:'memberState',membersEventRef:'memberEvent'};
  for(const [key,kind]of Object.entries(refTypes))if(key in q.payload)for(const r of Array.isArray(q.payload[key])?q.payload[key]:[q.payload[key]]){N(r?.kind===kind,'VALIDATION_ERROR','REF_KIND');M.path(r);}
  N(Array.isArray(q.expected)&&q.expected.length<=64,'SCOPE_TOO_LARGE');const keys=new Set();
  for(const e of q.expected){N(e&&Object.keys(e).every(k=>['kind','id','parent','revision','exists'].includes(k)));const r={kind:e.kind,id:e.id,...(e.parent?{parent:e.parent}:{})},p=M.path(r);N(!keys.has(e.kind+':'+p));keys.add(e.kind+':'+p);N(('revision'in e)!==('exists'in e));N('revision'in e?M.rev(e.revision):e.exists===false);}
  const s=q.scope;N(s&&typeof s==='object'&&Object.keys(s).every(k=>['targetRefs','sealedManifestRef','manifestDigest'].includes(k)));N(Array.isArray(s.targetRefs)!==!!s.sealedManifestRef,'VALIDATION_ERROR','SCOPE');
  if(s.targetRefs){N(s.targetRefs.length<=32,'SCOPE_TOO_LARGE');s.targetRefs.forEach(M.path);}else{M.path(s.sealedManifestRef);N(/^[a-f0-9]{64}$/.test(s.manifestDigest));}
  if(q.clientContext)N(Object.keys(q.clientContext).every(k=>['uiContractVersion','locale'].includes(k))&&Object.values(q.clientContext).every(v=>typeof v==='string'&&v.length<=32));
  const walk=(v,k='',depth=0)=>{N(v!==null||cmd==='ACTIVITY_TYPE_SAVE'&&depth===1&&k==='durationMinutes'&&v===q.payload.durationMinutes,'VALIDATION_ERROR','NULL');if(typeof v==='string')N(v.length<=(k==='title'?240:k==='reason'?2000:8000));if(Array.isArray(v)){N(v.length<=32,'SCOPE_TOO_LARGE');v.forEach(x=>walk(x,k,depth+1));}else if(v&&typeof v==='object'){N(!Object.keys(v).some(k=>['actor','actorType','uid','role','grant','personId','authorityBasis'].includes(k)),'VALIDATION_ERROR','INJECTED_AUTHORITY');for(const [key,x]of Object.entries(v))walk(x,key,depth+1);}};
  walk(q.payload);const normalized=M.json(q);N(Utilities.newBlob(JSON.stringify(normalized)).getBytes().length<=65536&&Utilities.newBlob(JSON.stringify(normalized.payload)).getBytes().length<=32768,'SCOPE_TOO_LARGE');
  normalized.expected.sort((a,b)=>(a.kind+a.id).localeCompare(b.kind+b.id));if(s.targetRefs)normalized.scope.targetRefs=[...new Map(s.targetRefs.map(r=>[M.path(r),r])).values()].sort((a,b)=>M.path(a).localeCompare(M.path(b)));
  const institutional={...normalized};delete institutional.clientContext;return {cmd,ct,digest:M.hash(institutional),q:normalized};
 }
 function plan(raw,handle,input){
  f4ActivitiesExecutionGate_(options)();
  const parsed=envelope(input),{cmd,ct,q}=parsed,p=q.payload,now=clock(),principal=options.principals.resolve(handle);
  N(principal&&ct.actors.includes(principal.kind),'AUTHORIZATION_DENIED','ACTOR_TYPE');

   const attendanceBatch=['ATTENDANCE_SAVE_DRAFT','ATTENDANCE_FINALIZE','ATTENDANCE_ROSTER_APPLY'].includes(cmd);const cache=new Map(),writes=new Map(),before=[],after=[],required=new Set(),effects=[];let actor,authorityBasis,cycleRef;
   function get(path){N(typeof path==='string');if(!cache.has(path)){N(cache.size<(attendanceBatch?320:100),'SCOPE_TOO_LARGE','READ_BUDGET');cache.set(path,raw.get(path).data());}return M.clone(writes.has(path)?writes.get(path):cache.get(path));}
   function prefetch(paths){const pending=[...new Set(paths)].filter(p=>!cache.has(p));N(cache.size+pending.length<=(attendanceBatch?320:100),'SCOPE_TOO_LARGE','READ_BUDGET');if(!pending.length)return;N(typeof raw.getAll==='function','UNRESOLVED_DEPENDENCY','BATCH_READER');const rows=raw.getAll(pending);pending.forEach((p,i)=>cache.set(p,rows[i].data()));}
   function original(path){get(path);return cache.get(path);}
   // All optimistic inputs are already known. Fetch them in one transaction
   // batch, then let the same validators check existence and revisions below.
   prefetch(q.expected.map(e=>M.path({kind:e.kind,id:e.id,...(e.parent?{parent:e.parent}:{})})));
   // Revision preconditions are deterministic. Keep speculative revision reads
   // in the transaction cache, outside the domain's actual effect vector.
   const revisionPaths=q.expected.filter(e=>!e.parent&&(M.nested[M.collections[e.kind]]||[]).includes('revisions')).map(e=>M.path({kind:e.kind,id:e.id})+'/revisions/'+((e.revision||0)+1));
   if(revisionPaths.length)raw.getAll(revisionPaths);
   function set(path,data){get(path);writes.set(path,M.clone(data));N(writes.size<=(attendanceBatch?320:200),'SCOPE_TOO_LARGE','WRITE_BUDGET');}
   function load(r,expected=false,optional=false){const path=M.path(r),v=get(path);N(v?.recordMode!=='HISTORICAL_PARTIAL'||['PRESENTATION_CONTINUITY_CONNECT','HISTORICAL_PREAGENDA_RECOGNIZE'].includes(cmd)&&r.kind==='presentation'&&r.id===p.presentationRef.id||['ACTIVITY_ADMIN_REVIEW','ACTIVITY_PROMOTE_OPERATIONAL'].includes(cmd)&&r.kind==='activity'&&r.id===p.activityRef.id&&v.administrativeReview,'UNRESOLVED_DEPENDENCY','HISTORICAL_RECORD_REQUIRES_RECONCILIATION');if(expected&&actor?.kind==='PORTAL_AUTHENTICATED_MEMBER'&&v){
    let owner=v.personRef?.id||v.ownerPersonRef;
    if(!owner&&v.presentationRef)owner=get(M.path(v.presentationRef))?.personRef?.id;
    if(!owner&&v.unitRef)owner=get(M.path(v.unitRef))?.personRef?.id;
    if(!owner&&v.submissionRef){const sub=get(M.path(v.submissionRef));if(sub?.presentationRef)owner=get(M.path(sub.presentationRef))?.personRef?.id;}
    const partners=cmd==='JOINT_PROPOSE'?p.unitRefs:cmd==='JOINT_CONSENT'?get(M.path(p.flowCaseRef))?.unitRefs:[];if(owner)N(owner===actor.personId||r.kind==='unit'&&partners?.some(u=>u.id===r.id)&&partners.some(u=>get(M.path(u))?.personRef?.id===actor.personId),'AUTHORIZATION_DENIED','EXPECTED_OWNER');
    if(['person','memberState','memberEvent'].includes(r.kind))N((v.idPessoa||r.id)===actor.personId,'AUTHORIZATION_DENIED','EXPECTED_MEMBERS_SCOPE');
   }if(expected){const e=q.expected.find(x=>M.path({kind:x.kind,id:x.id,...(x.parent?{parent:x.parent}:{})})===path);N(e,'VALIDATION_ERROR','MISSING_PRECONDITION');required.add(path);const old=original(path);N('revision'in e?old&&old.revision===e.revision:old===undefined,'STALE_REVISION','EXPECTED_VECTOR_CHANGED');}N(optional||v,'UNRESOLVED_DEPENDENCY','REFERENCE_MISSING');return v;}
   function save(r,d){const path=M.path(r),old=original(path);N(!M.immutable.has(r.kind)||!old,'INVALID_STATE_TRANSITION','IMMUTABLE');if(r.kind==='activity')f4ActivityReview_().guard(old,d,cmd);const n=(old?.revision||0)+1;const value={...d,id:r.id,schemaVersion:'act-v2-candidate',recordMode:d.recordMode||'NEW_V2',revision:n,createdAt:old?.createdAt||M.stamp(now),updatedAt:M.stamp(now),operationId:q.operationId,originRef:{kind:'CORE_COMMAND',operationId:q.operationId}};if(cmd==='ACTIVITY_ADMIN_REVIEW'&&r.kind==='activity'){value.recordMode=old.recordMode;value.createdAt=old.createdAt??null;if(Object.hasOwn(old,'originRef'))value.originRef=old.originRef;else delete value.originRef;}if(r.kind==='presentation'&&d.continuity?.schemaVersion==='presentation-continuity-v1')value.createdAt=old?.createdAt??null;if(r.kind==='activity'&&(cmd==='ACTIVITY_PROMOTE_OPERATIONAL'||old?.promotion?.fromMode==='HISTORICAL_PARTIAL'))value.createdAt=old.createdAt??null;M.validate(r,value);set(path,value);return value;}
   // Multiple changes to one entity in the same atomic command share one final
   // revision snapshot. A revision already stored before the command is immutable.
   function revision(r,d){const path=M.path(r),old=original(path);const v=save(r,d);if((M.nested[path.split('/')[0]]||[]).includes('revisions')){const rp=path+'/revisions/'+String(v.revision);N(!get(rp)||writes.has(rp),'CONFLICT','REVISION_EXISTS');set(rp,{number:v.revision,predecessorRevision:old?.revision||0,snapshot:v,recordedAt:M.stamp(now),actorAuthorityBasis:authorityBasis,actor,operationId:q.operationId});}return v;}
   function event(kind,d,parent){const r=M.ref(kind,M.hash({operationId:q.operationId,kind,parent:parent||''}).slice(0,32),parent);N(!load(r,false,true),'CONFLICT','EVENT_EXISTS');save(r,{...d,actorAuthorityBasis:authorityBasis,actor,recordedAt:M.stamp(now)});return r;}
   function manifest(r){const m=load(r);N(m.state==='SEALED'&&m.complete===true&&Array.isArray(m.items)&&m.items.length<=32,'UNRESOLVED_DEPENDENCY','MANIFEST_UNSEALED');N(m.digest===M.hash(m.items),'UNRESOLVED_DEPENDENCY','MANIFEST_DIGEST');if(m.validUntil)N(M.ms(m.validUntil)>now,'STALE_REVISION','MANIFEST_EXPIRED');for(const b of m.basisRevisionVector||[])N(load(b.ref).revision===b.revision,'STALE_REVISION','MANIFEST_BASIS');return m;}
   let human;
   const terminal=null;
   if(principal.kind==='SYSTEM_JOB'||principal.kind==='WORKSPACE_EXECUTOR'){
    N(!principal.uid&&principal.servicePrincipal&&principal.scope&&principal.ruleRef,'AUTHORIZATION_DENIED','INTERNAL_PRINCIPAL');
    const rule=load(principal.ruleRef);N(rule.approved===true&&rule.active===true&&rule.commands?.includes(q.commandType),'AUTHORIZATION_DENIED','JOB_RULE');
    N(principal.kind!=='SYSTEM_JOB'||ct.jobs.includes(principal.jobKind),'AUTHORIZATION_DENIED','JOB_KIND');
    actor={kind:principal.kind,servicePrincipal:principal.servicePrincipal,jobKind:principal.jobKind||'WORKSPACE',runId:principal.runId,ruleVersion:rule.revision};authorityBasis={ruleRef:principal.ruleRef,ruleRevision:rule.revision,scope:principal.scope};
   }else{
    human=authority.identity(get,principal.uid,now);actor={...human.actor,kind:principal.kind};authorityBasis=human.authorityBasis;
   }
   // Resolve targets from typed selectors and server links. Client scope never grants access.
   const targetRefs=[];function collect(v){if(v&&typeof v==='object'){if(v.kind&&v.id&&M.collections[M.aliases[v.kind]||v.kind]||v&&v.kind==='config'){targetRefs.push(v);return;}for(const x of Object.values(v))collect(x);}}collect(p);
   let targetPersonIds=new Set();for(const r of targetRefs){const d=load(r,false,true);if(d?.cycleRef?.kind&&d?.cycleRef?.id){if(cycleRef)N(M.path(cycleRef)===M.path(d.cycleRef),'VALIDATION_ERROR','CROSS_CYCLE');cycleRef=d.cycleRef;}if(d?.personRef)targetPersonIds.add(d.personRef.id||d.personRef);}
   if(p.cycleRef)cycleRef=p.cycleRef;
   if(human){
    const own=principal.kind==='PORTAL_AUTHENTICATED_MEMBER';
    if(own){
     N(cmd==='JOINT_PROPOSE'?targetPersonIds.has(actor.personId):[...targetPersonIds].every(x=>x===actor.personId),'AUTHORIZATION_DENIED','OWNER_MISMATCH');
     if(human.s.state==='SUSPENSO'){
      N(['MATERIAL_SUBMIT','MATERIAL_CORRECTION_SUBMIT','ARTIFACT_UPLOAD_REQUEST'].includes(cmd)&&p.presentationRef,'AUTHORIZATION_DENIED','SUSPENDED_OPERATION_FORBIDDEN');
      const pr=load(p.presentationRef,true),rq=load(pr.materialRequirementRef,true),rz=pr.realizationRef&&load(pr.realizationRef);
      N(pr.personRef.id===actor.personId&&rz?.recognized===true&&!rq.currentRegularizationRef,'AUTHORIZATION_DENIED','SUSPENDED_SCOPE');
      authorityBasis={...authorityBasis,kind:'SUSPENDED_MATERIAL_SCOPE',restrictionProof:authority.suspension(get,human,now),presentationRef:p.presentationRef,realizationRef:pr.realizationRef,requirementRevision:rq.revision,operationType:cmd};
     }else{N((human.s.state==='MEMBRO_EFETIVO'||human.s.state==='MEMBRO_INGRESSANTE'&&['ATTENDANCE_JUSTIFICATION_REQUEST','ATTENDANCE_PROOF_UPLOAD_REQUEST'].includes(cmd))&&!human.s.blocked,'AUTHORIZATION_DENIED');const cap=ct.capabilities.find(c=>c.endsWith('_SELF'))||'PRESENTATION_SELF_READ';authority.ordinary(get,principal.uid,'SELF',cap,now);}
    }else{
     const scope=f4AttendanceScope_(cmd)?'ACTIVITIES':cmd==='CYCLE_CREATE'?'PRESENTATIONS_PLAN':cycleRef?'PRESC_'+cycleRef.id:terminal?.authorityBasis?.scope;N(scope,'AUTHORIZATION_DENIED','SCOPE_UNRESOLVED');
     const cap=ct.capabilities.find(c=>!c.endsWith('_SELF'));N(cap,'AUTHORIZATION_DENIED','CAPABILITY');const proof=authority.ordinary(get,principal.uid,scope,cap,now);authorityBasis={...authorityBasis,scope,capability:cap,functionBasis:proof.g.entries?.[cap]||{}};
    }
   }else if(cycleRef)N(principal.scope==='PRESC_'+cycleRef.id,'AUTHORIZATION_DENIED','JOB_SCOPE');
   const leaseUntil=now+15000;
   for(const e of q.expected)load({kind:e.kind,id:e.id,...(e.parent?{parent:e.parent}:{})},true,e.exists===false);
   let scopeRefs=q.scope.targetRefs;if(!scopeRefs){const m=manifest(q.scope.sealedManifestRef);N(m.digest===q.scope.manifestDigest);scopeRefs=m.items.map(i=>i.ref);}
   for(const r of targetRefs)N(scopeRefs.some(x=>M.path(x)===M.path(r))||['policy','artifact','config','assessment','receipt','review','realization','reconciliation'].includes(r.kind),'VALIDATION_ERROR','TARGET_OUTSIDE_SCOPE');
   function config(){N(cycleRef,'UNRESOLVED_DEPENDENCY','CYCLE');const c=load(cycleRef,true);if(c.scheduledConfigRef&&c.scheduledConfigRef.kind!=='NONE'){const scheduled=load(c.scheduledConfigRef,true);N(scheduled.approved&&M.ms(scheduled.effectiveFrom)>now,'UNRESOLVED_DEPENDENCY','CONFIG_ACTIVATION_REQUIRED');}const cfg=load(c.currentConfigRef,true);N(cfg.approved===true&&!cfg.parameters?.reviewRequired,'HUMAN_DECISION_REQUIRED','PARAMETER_REVIEW_REQUIRED');return cfg;}
   function capacity(){N(cycleRef,'UNRESOLVED_DEPENDENCY');const r=load(M.ref('capacity','registry'),true),h=load(M.ref('capacity',cycleRef.id),true);N(h.state==='SEALED'&&h.synchronizedMembersEpoch===r.membersEpoch&&h.synchronizedCycleRegistryEpoch===r.cycleRegistryEpoch&&M.ms(h.validUntil)>now,'UNRESOLVED_DEPENDENCY','CAPACITY_STALE');return h;}
   function bumpCapacity(){const r=M.ref('capacity',cycleRef.id),h=load(r);save(r,{...h,capacityEpoch:h.capacityEpoch+1});}
   function own(r){const d=load(r,true);if(human&&principal.kind==='PORTAL_AUTHENTICATED_MEMBER')N(d.personRef?.id===actor.personId,'AUTHORIZATION_DENIED','OWNER_MISMATCH');return d;}
   function emit(kind,target,policyRef,details){const policy=load(policyRef);N(policy.approved&&policy.active,'UNRESOLVED_DEPENDENCY','EFFECT_POLICY');const r=M.ref('outbox',M.hash({operationId:q.operationId,kind,target}).slice(0,32));N(!load(r,false,true),'CONFLICT');save(r,{effectKind:kind,targetRef:target,cycleRef,policyRef,policyRevision:policy.revision,details,actor,authorityBasis,status:'PENDING',state:'PENDING',nextAttemptAt:M.stamp(now),leaseFence:0,attemptCount:0,dueAt:M.stamp(now),createdByOperation:q.operationId});effects.push(r);return r;}
   const x={M,N,get,prefetch,set,load,save,revision,event,manifest,own,config,capacity,bumpCapacity,emit,now,actor,principal,authorityBasis,cycleRef,p,cmd,operationId:q.operationId,options};
   const modules=[f4ActivitiesPreagendaCommands_,f4PresentationContinuityCommands_,f4ActivitiesAcademicCommands_,f4ActivitiesExceptionalJustificationCommands_,f4ActivitiesPhotoCommands_,f4ActivitiesMemberActivityCommands_,f4ActivitiesCyclePolicyCommands_,f4ActivitiesQuotaCommands_,f4ActivitiesReorganization_,f4ActivitiesIncidenceCommands_,f4ActivitiesPlanning_,f4ActivitiesCalendar_,f4ActivitiesThematic_,f4ActivitiesMaterials_,f4ActivitiesPublication_,f4ActivitiesJobs_,f4AttendanceCommands_];let result;
   let failure;
   try{for(const module of modules){const handlers=module(x);if(handlers[cmd]){result=f4DevReadTiming_().stage('validation',()=>handlers[cmd]());break;}}N(result!==undefined,'UNRESOLVED_DEPENDENCY','HANDLER');}
   catch(e){const domainCodes=['VALIDATION_ERROR','INCOMPLETE_REQUEST','AUTHORIZATION_DENIED','NOT_FOUND','STALE_REVISION','CONFLICT','INVALID_STATE_TRANSITION','UNRESOLVED_DEPENDENCY','HUMAN_DECISION_REQUIRED','CAPACITY_EXHAUSTED','THEMATIC_QUOTA_EXHAUSTED','CONFIG_CAPABILITY_DISABLED','DEADLINE_VIOLATION','SCOPE_TOO_LARGE'];if(!domainCodes.includes(e.code))throw e;failure={code:e.code,reasonCode:e.reasonCode||e.code,retryDirective:'STOP_OR_NEW_OPERATION'};writes.clear();effects.length=0;result={code:e.code};}
   const businessState=failure?'REJECTED':'COMMIT_APPLIED';
   // Any source change invalidates every affected public projection and private generation.
   const affected=new Map();
   function affectUnit(ur){const u=load(ur);if(u.presentationRef)affected.set(u.presentationRef.id,u.presentationRef);if(u.personRef){const k='presentationMemberViews/'+u.personRef.id,h=get(k);if(h)set(k,{...h,state:'INVALIDATED',viewEpoch:h.viewEpoch+1,rebuildDueAt:M.stamp(now)});}}
   for(const [path,d]of [...writes]){if(!d||path.split('/').length!==2)continue;const col=path.split('/')[0];if(col==='canonicalActivities'&&d.presentationSessionRef){const session=load(d.presentationSessionRef);for(const slot of session.slots||[])affectUnit(slot.unitRef);}if(['presentationSessions','calendarAssignments','presentationThematicRequests','presentationDynamicOccupancies','presentationUnitHeads'].includes(col)){if(d.unitRef)affectUnit(d.unitRef);if(col==='presentationUnitHeads')affectUnit(M.ref('unit',d.id));if(col==='presentationSessions')for(const sl of [...(d.slots||[]),...(cache.get(path)?.slots||[])])affectUnit(sl.unitRef);}}
   for(const pr of affected.values()){const k='publicPresentations/'+pr.id,pub=get(k);if(pub?.published&&cmd!=='PUBLICATION_APPROVE')set(k,{...pub,published:false,revision:(pub.revision||0)+1,updatedAt:M.stamp(now)});}
   // Invalidate private generations/public bases atomically with every changed source.
   const sourceChanges=[...writes.entries()];for(const [path,d]of sourceChanges){if(!d)continue;if(['presentationUnitHeads','canonicalPresentations','presentationObligations','presentationMaterialRequirements','presentationCorrectionOpportunities'].includes(path.split('/')[0])){
     const person=d.personRef?.id||(d.presentationRef?load(d.presentationRef).personRef?.id:undefined);if(person){const k='presentationMemberViews/'+person,h=get(k);if(h)set(k,{...h,state:'INVALIDATED',viewEpoch:h.viewEpoch+1,rebuildDueAt:M.stamp(now)});}
    }
    if(path.startsWith('canonicalPresentations/')){const k='publicPresentations/'+d.id,pub=get(k);if(pub?.published&&cmd!=='PUBLICATION_APPROVE')set(k,{...pub,published:false,revision:(pub.revision||0)+1,updatedAt:M.stamp(now)});}
   }
   if(!failure)f4ActivitiesProjectAttendanceReads_(x,[...writes.entries()]);
   for(const [path,value]of cache){if(M.validPath(path))before.push({path,revision:value?.revision??null});}
   for(const [path,value]of writes){if(M.validPath(path))after.push({path,revision:value?.revision??null});}
   N(before.length<=(attendanceBatch?320:64)&&after.length<=(attendanceBatch?320:64),'SCOPE_TOO_LARGE','VECTOR_BUDGET');
   const response={schemaVersion:'geapa-response-v2-candidate',operationId:q.operationId,commandType:q.commandType,businessState,...(failure?{error:failure}:{}),transportStage:'RESPONSE_READY',externalState:effects.length?'EXTERNAL_EFFECT_PENDING':'NONE',result,replay:false,receiptRef:q.operationId};
   N(clock()<leaseUntil,'TEMPORARY_UNAVAILABLE','LEASE_EXPIRED');
   for(const [path,d]of cache)if(M.validPath(path)&&d?.validUntil&&M.ms(d.validUntil)>now)N(M.ms(d.validUntil)>clock(),'STALE_REVISION','SOURCE_TIME_BOUNDARY');
   // Revalidate temporal authority at the final cut; snapshot changes are also protected by CAS.
   if(human){const finalNow=clock();const current=authority.identity(get,principal.uid,finalNow);N(current.h.revision===human.h.revision&&current.b.bindingVersion===human.b.bindingVersion,'AUTHORIZATION_DENIED','AUTHORITY_CHANGED');if(authorityBasis.kind==='SUSPENDED_MATERIAL_SCOPE')authority.suspension(get,current,finalNow);else authority.ordinary(get,principal.uid,authorityBasis.scope||'SELF',authorityBasis.capability||ct.capabilities.find(c=>c.endsWith('_SELF'))||'PRESENTATION_SELF_READ',finalNow);}
   if(!failure&&cmd==='CALENDAR_RESERVE'&&actor.kind==='PORTAL_AUTHENTICATED_MEMBER')N(f4ActivitiesSelfSchedulingWindow_(config().parameters,clock()).open,'DEADLINE_VIOLATION','SELF_SCHEDULING_WINDOW_CLOSED');
   // Return domain effects only. The shared coordinator owns CAS, leases and terminal records.
   return {response,writes:[...writes].map(([path,data])=>({path,data:data===undefined?null:data,kind:data===undefined?'delete':'set'}))};
 }
 function execute(){N(false,'UNRESOLVED_DEPENDENCY','SHARED_COMMAND_PIPELINE_REQUIRED');}
 function receipt(){N(false,'UNRESOLVED_DEPENDENCY','SHARED_COMMAND_STATUS_REQUIRED');}
 return Object.freeze({plan,execute,receipt,envelope});
}


