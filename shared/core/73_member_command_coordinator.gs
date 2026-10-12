// C20 transaction composition only. No command dispatcher, admission, token verifier or worker.
function f4MemberCommandCoordinator_(){
 function stage(tx,plan,technicalEffects,validateTechnical,now){
  const F=f4Firestore_(),P=f4MemberMutationPlan_();f4MemberModel_(true).need((F.inspectTransaction(tx).environment==='SYNTHETIC_LOCAL' || f4DevBoundary_().permit(F.inspectTransaction(tx))),'M13_OPERATIONAL_WRITER_OFF');
  f4MemberModel_(true).need(plan?.auditFacts?.mode==='M13','M13_AUTHZ_DENIED');
  const full=P.appendTechnical(tx,plan,technicalEffects,validateTechnical);return P.applyMutationPlan(tx,full,now);
 }
 return {stage,createLocal:options=>f4MemberCommandRecovery_().createLocal(options)};
}

// D finalization/recovery only. No installed handler, endpoint, worker or remote session.
function f4MemberCommandRecovery_(){
 const C=f4Core_(),F=f4Firestore_(),K=f4SharedCommandCatalog_(),L=f4MemberLedgerM13_(),A=f4SharedAuthz_(),P=f4MemberMutationPlan_(),N=C.need,T=F.Timestamp.fromMillis;
 const clone=v=>F.decode(F.encode(v));
 function createLocal(options){
  C.exact(options,['config','clock','verifyIdentity','producers']);C.destination(options.config);N((options.config.environment==='SYNTHETIC_LOCAL' || f4DevBoundary_().permit(options.config)),'DESTINATION_INVALID');N(typeof options.clock==='function'&&typeof options.verifyIdentity==='function','AUTH_BINDING_INVALID');
  const config=Object.freeze({...options.config}),clock=options.clock,verify=options.verifyIdentity,seals=new WeakMap(),producers={};
  for(const [type,v]of Object.entries(options.producers)){N(Object.hasOwn(K.types,type),'COMMAND_TYPE_INVALID');C.exact(v,['validate','resolve']);N(typeof v.validate==='function'&&typeof v.resolve==='function','COMMAND_TYPE_INVALID');producers[type]=Object.freeze({...v});}
  function now(){const n=clock();N(Number.isSafeInteger(n),'DOMAIN_STATE_INVALID');return n}
  function authenticate(token){const v=verify(token,now());C.exact(v,['uid','expiresAt']);N(C.id(v.uid)&&Number.isSafeInteger(v.expiresAt)&&now()<v.expiresAt,'AUTH_BINDING_INVALID');const principal=Object.freeze({});seals.set(principal,Object.freeze({...v}));return principal}
  function principal(tx,s,key){const v=seals.get(s),info=F.inspectTransaction(tx);N(v&&now()<v.expiresAt,'AUTH_BINDING_INVALID');N(info.environment===config.environment&&info.projectId===config.projectId&&info.database===config.database,'DESTINATION_INVALID');if(K.act(key?.commandType))f4ActivitiesPrefetchCommand_(tx,v.uid,key);return v}
  function ownAccess(get,a,n,key){
   if(K.act(key?.commandType))return f4ActivitiesPipeline_().statusAccess(get,a,key,n);
   const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),b=get('memberAuthBindings/'+a.uid);N(b,'AUTH_BINDING_INVALID');const subject=get('memberAuthzSubjects/'+b.idPessoa),head=get('memberAggregateHeads/'+b.idPessoa),sp=control.selfReadPolicy;N(subject&&head&&sp,'ACCESS_REVOKED');
   return A.authorize(get,{schemaVersion:'member-authz-request-v13',uid:a.uid,kind:'SELF',capability:'MEMBERS_SELF_READ',sourceId:sp.policyId,functionId:null,actorScope:null,expected:{authorityEpoch:control.epoch,policyEpoch:control.policyEpoch,readAuthzEpoch:control.readAuthzEpoch,subjectEpoch:subject.authzEpoch,bindingVersion:b.bindingVersion,personRevision:head.revision,authorizationRevision:sp.policyRevision},tokenExpiresAt:T(a.expiresAt),admissionExpiresAt:T(a.expiresAt),target:{personId:b.idPessoa,scopeKey:null,boardId:null,roleKey:null,resources:['memberCurrentStates/'+b.idPessoa],caseId:null,purpose:null,fieldMask:['state'],linkId:null,effect:null,effectiveDate:null},evidenceRefs:[]},n);
  }
  function readStatus(tx,s,key){const a=principal(tx,s,key);K.selector(key);const get=p=>tx.get(p).data();ownAccess(get,a,now(),key);const result=L.recover(get,key,a.uid,now());if(result)return clone(result);if(key.commandType==='members.onboarding.complete'&&typeof f4MemberOnboarding_==='function'){const r=get('commandRequests/'+key.operationId),run=get('memberOnboardingRuns/'+key.operationId);N(r&&r.ownerUid===a.uid&&r.digest===key.digest&&r.commandType===key.commandType,'OPERATION_CONFLICT');if(run)return {state:f4MemberOnboarding_().run(run,r).stage};}N(false,'RECONCILIATION_REQUIRED');}
  function ticket(v,r){C.exact(v,['schemaVersion','id','operationId','ownerUid','idPessoa','requestEnvelope','digest','authProofs','issuedAt','expiresAt','tokenExpiresAt','authorityEpoch','status']);N(v.schemaVersion==='member-admission-ticket-v13'&&v.id===r.ticketId&&v.operationId===r.operationId&&v.ownerUid===r.ownerUid&&v.digest===r.digest&&C.id(v.idPessoa)&&v.status==='ISSUED'&&C.rev(v.authorityEpoch)&&v.authorityEpoch>0&&K.equal(v.requestEnvelope,K.template(r))&&Array.isArray(v.authProofs)&&v.authProofs.length===K.capabilities(r).length,'AUTH_BINDING_INVALID');N(C.time(v.issuedAt,v.expiresAt,now())&&C.ms(v.expiresAt)<=C.ms(v.issuedAt)+120000&&C.ms(v.expiresAt)<=C.ms(v.tokenExpiresAt),'ACCESS_REVOKED');return v}
  function lease(v,id,handle,n){C.exact(handle,['workerId','fencingToken']);C.exact(v,['schemaVersion','operationId','domain','workerId','fencingToken','deadline']);N(v.schemaVersion==='member-command-lease-v13'&&v.operationId===id&&v.domain==='MEMBERS'&&C.id(v.workerId)&&v.workerId===handle.workerId&&C.rev(v.fencingToken)&&v.fencingToken>0&&v.fencingToken===handle.fencingToken&&Number.isSafeInteger(C.ms(v.deadline))&&n<C.ms(v.deadline),'LEASE_LOST');return v}
  function admitted(tx,a,key,handle,claiming=false){
   const get=p=>tx.get(p).data(),r=get('commandRequests/'+key.operationId),producer=producers[key.commandType];N(producer,'COMMAND_TYPE_INVALID');K.envelope(config.projectId,r,producer.validate);N(r.operationId===key.operationId&&r.ownerUid===a.uid&&r.digest===key.digest&&r.commandType===key.commandType&&C.ms(r.receivedAt)<=now(),'OPERATION_CONFLICT');
   const t=ticket(get('memberAdmissionTickets/'+r.ticketId),r),u=get('memberAdmissionUses/'+r.ticketId),p=get('commandProvenance/'+r.operationId),stored=get('commandLeases/'+r.operationId),control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),h=claiming?stored:(stored?.schemaVersion==='member-command-lease-v13-g1'?f4MemberClaim_().active(stored,key,handle,control,now()):lease(stored,r.operationId,handle,now()));
   C.exact(u,['schemaVersion','ticketId','operationId','ownerUid','reservedAt']);N(u.schemaVersion==='member-admission-use-v13'&&u.ticketId===r.ticketId&&u.operationId===r.operationId&&u.ownerUid===a.uid&&K.equal(u.reservedAt,r.receivedAt)&&C.time(t.issuedAt,t.expiresAt,C.ms(u.reservedAt)),'AUTH_BINDING_INVALID');
   C.exact(p,['schemaVersion','operationId','ownerUid','ticketId','digest','authRequests']);N(p.schemaVersion==='member-command-provenance-v13'&&p.operationId===r.operationId&&p.ownerUid===a.uid&&p.ticketId===r.ticketId&&p.digest===r.digest&&Array.isArray(p.authRequests)&&p.authRequests.length===t.authProofs.length,'AUTH_BINDING_INVALID');
   for(let i=0;i<p.authRequests.length;i++){const q=p.authRequests[i],proof=t.authProofs[i];N(q.uid===a.uid&&q.capability===K.capabilities(r)[i]&&proof.capability===q.capability&&proof.idPessoa===t.idPessoa&&proof.authorityEpoch===t.authorityEpoch&&C.ms(q.tokenExpiresAt)<=a.expiresAt&&C.ms(q.admissionExpiresAt)===C.ms(t.expiresAt),'AUTH_BINDING_INVALID');}
   if(claiming||h?.schemaVersion==='member-command-lease-v13-g1')N(control.mode==='CANONICAL'&&control.activeWriterVersion==='M13'&&control.epoch===t.authorityEpoch,'ACCESS_REVOKED');
   return {get,r,producer,t,p,h,control};
  }
  function recheck(x){
   const ingress=x.r.commandType==='members.ingress.execute'&&typeof f4MemberIngressCommands_==='function'&&x.r.evidenceRefs.some(ref=>{const att=x.get('memberEvidenceAttestations/'+ref);return att&&x.get('memberEvidenceContents/'+att.contentHash)?.schemaVersion==='member-act-content-v13'});
   if(x.r.commandType==='members.onboarding.complete')return f4MemberOnboarding_().evaluate(x.get,x.r,x.p.authRequests,x.t.authProofs,now(),config).proofs;
   return x.p.authRequests.map((q,i)=>q.schemaVersion==='member-authz-governance-request-v13'?A.recheckGovernance(x.get,q,x.t.authProofs[i],now(),x.r):ingress?A.recheckCommand(x.get,q,x.t.authProofs[i],now(),x.r):A.recheck(x.get,q,x.t.authProofs[i],now()));
  }
  function claim(tx,s,key,workerId){
   const a=principal(tx,s,key);K.selector(key);const get=p=>tx.get(p).data(),prior=L.recover(get,key,a.uid,now());
   if(prior){ownAccess(get,a,now(),key);return {state:'RECOVERED',result:clone(prior)};}
   // An expired ACT ticket cannot execute again. Close only its technical
   // outcome under current authority; never renew a ticket or replay effects.
   if(K.act(key.commandType)){
    const r=get('commandRequests/'+key.operationId),t=r&&get('memberAdmissionTickets/'+r.ticketId),n=now();
    if(t&&Number.isSafeInteger(C.ms(t.expiresAt))&&n>=C.ms(t.expiresAt)){
     const currentAccess=ownAccess(get,a,n,key);K.envelope(config.projectId,r,producers[key.commandType].validate);
     N(r.ownerUid===a.uid&&r.operationId===key.operationId&&r.commandType===key.commandType&&r.digest===key.digest,'OPERATION_CONFLICT');
     N(t.schemaVersion==='member-admission-ticket-v13'&&t.id===r.ticketId&&t.ownerUid===a.uid&&t.operationId===r.operationId&&t.digest===r.digest&&t.status==='ISSUED'&&K.equal(t.requestEnvelope,K.template(r))&&C.time(t.issuedAt,t.expiresAt,C.ms(r.receivedAt))&&C.ms(t.expiresAt)<=C.ms(t.issuedAt)+120000,'AUTH_BINDING_INVALID');
     const use=get('memberAdmissionUses/'+r.ticketId),lease=get('commandLeases/'+key.operationId);
     N(use?.schemaVersion==='member-admission-use-v13'&&use.ticketId===r.ticketId&&use.operationId===r.operationId&&use.ownerUid===a.uid&&K.equal(use.reservedAt,r.receivedAt),'AUTH_BINDING_INVALID');
     if(lease){f4MemberClaim_().validate(lease);N(lease.operationId===key.operationId&&lease.digest===key.digest&&lease.domain==='ACT'&&lease.state!=='TERMINAL'&&C.ms(lease.deadline)<=C.ms(t.expiresAt),'LEASE_LOST');}
     const decision=L.decision({state:'FAILED',code:'DENIED',revisionBefore:[],revisionAfter:[],domainResult:{schemaVersion:'geapa-response-v2-candidate',operationId:r.operationId,commandType:r.commandType,businessState:'REJECTED',error:{code:'DEADLINE_VIOLATION',reasonCode:'ADMISSION_EXPIRED'},transportStage:'RESPONSE_READY',externalState:'NONE',result:{},replay:false,receiptRef:r.operationId}}),ledger=L.make(r,decision,n);
     const proof=currentAccess.proof,audit={schemaVersion:'member-audit-v13',operationId:r.operationId,ownerUid:a.uid,actorPersonId:t.idPessoa,capabilities:K.capabilities(r),proofDigest:C.hash(F.encode([proof])),authorizationRefs:[{kind:proof.kind,capability:proof.capability,sourceId:proof.kind==='DEV_ADMIN'?proof.profilePath:proof.grantId||proof.policyId,sourceRevision:proof.kind==='DEV_ADMIN'?(proof.devProfileRevision||proof.profileRevision):proof.grantRevision||proof.policyRevision,authorityEpoch:proof.authorityEpoch,policyEpoch:proof.policyEpoch,subjectEpoch:proof.subjectEpoch,bindingVersion:proof.bindingVersion,personRevision:proof.personRevision}],targetDigest:C.hash(r.target),evidenceRefs:r.evidenceRefs,decisionRef:r.decisionRef,digest:r.digest,revisionBefore:[],revisionAfter:[],resultCode:'DENIED',createdAt:T(n)};
     const writes=[['commandLedger',ledger],['commandReceipts',L.toReceipt(ledger)],['commandTombstones',L.toTombstone(ledger)],['memberAudit',audit]].map(([col,data])=>({path:col+'/'+r.operationId,data,kind:'set'}));
     N(!get('memberAudit/'+r.operationId),'OPERATION_CONFLICT');
     const plan=P.planTerminal(tx,t.idPessoa,n,n+15000),allowed=new Map(writes.map(w=>[w.path,C.hash(F.encode(w))]));
     P.applyMutationPlan(tx,P.appendTechnical(tx,plan,writes,w=>allowed.get(w.path)===C.hash(F.encode(w))),n);
     return {state:'RECOVERED',result:clone(decision)};
    }
   }
   // Exhaustion is a technical denial state, never a renewed execution permission.
   // Persist it even when the third lease and its ticket expired together.
   const exhausted=get('commandLeases/'+key.operationId);
   if(exhausted?.schemaVersion==='member-command-lease-v13-g1'&&exhausted.attempts===3&&(exhausted.state==='RECONCILIATION_REQUIRED'||exhausted.state==='CLAIMED'&&now()>=C.ms(exhausted.deadline))){
    f4MemberClaim_().validate(exhausted);const r=get('commandRequests/'+key.operationId),producer=producers[key.commandType];N(producer,'COMMAND_TYPE_INVALID');K.envelope(config.projectId,r,producer.validate);N(r.ownerUid===a.uid&&r.operationId===key.operationId&&r.commandType===key.commandType&&r.digest===key.digest&&exhausted.digest===key.digest,'OPERATION_CONFLICT');ownAccess(get,a,now(),key);
    if(exhausted.state!=='RECONCILIATION_REQUIRED')tx.set('commandLeases/'+key.operationId,f4MemberClaim_().exhaust(exhausted,key,now()));return {state:'RECONCILIATION_REQUIRED'};
   }
   const x=admitted(tx,a,key,null,true),proofs=recheck(x),H=f4MemberClaim_(),deadline=Math.min(a.expiresAt,C.ms(x.t.expiresAt),...proofs.map(p=>C.ms(p.expiresAt)));
   const result=H.claim(x.h,key,workerId,x.control,now(),deadline);
   if(result.changed)tx.set('commandLeases/'+key.operationId,result.lease);
   return result.acquired?{state:'CLAIMED',handle:H.handle(result.lease),attempts:result.lease.attempts}:{state:'RECONCILIATION_REQUIRED'};
  }
  function release(tx,s,key,handle){
   const a=principal(tx,s,key);K.selector(key);const get=p=>tx.get(p).data(),prior=L.recover(get,key,a.uid,now());
   if(prior){ownAccess(get,a,now(),key);return {state:'RECOVERED',result:clone(prior)};}
   const x=admitted(tx,a,key,handle),H=f4MemberClaim_();recheck(x);const v=H.release(x.h,key,handle,x.control,now());tx.set('commandLeases/'+key.operationId,v);return {state:v.state};
  }
  function rows(values){return values.map(v=>{const [resourceType,...rest]=v.path.split('/'),id=rest.join('/');return {resourceType,id,revision:v.revision}}).sort((a,b)=>{const x=a.resourceType+'/'+a.id,y=b.resourceType+'/'+b.id;return x<y?-1:x>y?1:0})}
  function finalize(tx,s,key,handle){
   const a=principal(tx,s,key);K.selector(key);const get=p=>tx.get(p).data();
   // Historical outcome takes precedence over expired admission and old source revisions.
   const prior=L.recover(get,key,a.uid,now());if(prior){ownAccess(get,a,now(),key);return clone(prior)}
   const x=admitted(tx,a,key,handle);let plan,d,outbox,saga,proofs=[];const deadline=Math.min(a.expiresAt,C.ms(x.t.expiresAt),C.ms(x.h.deadline));
   try{
    if(K.act(x.r.commandType)){proofs=recheck(x);const domain=f4ActivitiesPipeline_().plan(tx,x.r,x.p.authRequests[0],x.t.authProofs[0],now,deadline,config);plan=domain.plan;d=L.decision(domain.decision);}
    else {
    const onboarding=x.r.commandType==='members.onboarding.complete'&&typeof f4MemberOnboarding_==='function'&&x.producer.validate===f4MemberOnboarding_().validate;
    if(onboarding){saga=f4MemberOnboarding_().evaluate(get,x.r,x.p.authRequests,x.t.authProofs,now(),config);proofs=saga.proofs;plan=P.planTerminal(tx,x.t.idPessoa,now(),deadline);d=saga.terminal?L.decision({state:saga.run.stage,code:saga.run.stage==='COMPLETED'?'OK':'RECONCILIATION_REQUIRED',revisionBefore:[],revisionAfter:[]}):null;}
    else{
    const protectedIngress=x.r.commandType==='members.ingress.execute'&&typeof f4MemberIngressCommands_==='function'&&x.r.evidenceRefs.some(ref=>{const att=get('memberEvidenceAttestations/'+ref);return att&&get('memberEvidenceContents/'+att.contentHash)?.schemaVersion==='member-act-content-v13'});
    const governance=x.r.commandType==='members.governance.boardTerm'&&x.p.authRequests[0]?.schemaVersion==='member-authz-governance-request-v13';
    for(let i=0;i<x.p.authRequests.length;i++)proofs.push(governance?A.recheckGovernance(get,x.p.authRequests[i],x.t.authProofs[i],now(),x.r):protectedIngress?A.recheckCommand(get,x.p.authRequests[i],x.t.authProofs[i],now(),x.r):A.recheck(get,x.p.authRequests[i],x.t.authProofs[i],now()));
    // K2 is E's saga, not a single business plan; no implicit dispatch/handler.
    N(proofs.length===1,'COMMAND_TYPE_INVALID');const attemptTime=now(),ctx=x.producer.resolve(get,clone(x.r),attemptTime);N(ctx.mode==='M13'&&ctx.now===attemptTime&&K.equal(ctx.authorization.request,x.p.authRequests[0])&&K.equal(ctx.authorization.proof,x.t.authProofs[0]),'AUTH_BINDING_INVALID');
    if(Object.hasOwn(ctx,'outbox')){N(governance&&f4MemberBoardCommands_().outbox(ctx.outbox,x.r),'COMMAND_PAYLOAD_INVALID');outbox=ctx.outbox;delete ctx.outbox;}
    if(ctx.authorization.command)N(K.equal(ctx.authorization.command,x.r),'AUTH_BINDING_INVALID');
    const headType=governance?'memberBoardHeads':'memberAggregateHeads';N(ctx.expected===x.r.expectedRevisions.find(v=>v.resourceType===headType&&v.id===ctx.id)?.revision,'REVISION_CONFLICT');
    for(const c of ctx.changes){const [col,id]=c.path.split('/');N(c.expectedRevision===x.r.expectedRevisions.find(v=>v.resourceType===col&&v.id===id)?.revision,'REVISION_CONFLICT');}
    for(const [id,revision]of Object.entries(ctx.boardExpected))N(revision===x.r.expectedRevisions.find(v=>v.resourceType==='memberBoardHeads'&&v.id===id)?.revision,'REVISION_CONFLICT');
    const used=new Set([headType+'/'+ctx.id,...ctx.changes.map(c=>c.path),...Object.keys(ctx.boardExpected).map(id=>'memberBoardHeads/'+id)]);N(x.r.expectedRevisions.every(v=>used.has(v.resourceType+'/'+v.id)),'COMMAND_PAYLOAD_INVALID');
    ctx.deadline=Math.min(ctx.deadline,deadline);plan=P.planMutation(tx,ctx);d=L.decision({state:'COMPLETED',code:'OK',revisionBefore:rows(plan.revisionBefore),revisionAfter:rows(plan.revisionAfter)});
    }
    }
   }catch(e){
    const known=['STALE_REVISION','AUTHORIZATION_DENIED','VALIDATION_ERROR','UNRESOLVED_DEPENDENCY','CONFLICT','M13_AUTHZ_DENIED','M13_AUTHORITY_DENIED','AUTH_BINDING_INVALID','ACCESS_REVOKED','REVISION_CONFLICT','M1_REVISION_INVALID','SCHEMA_INVALID','COMMAND_PAYLOAD_INVALID','M13_ACT_INVALID','M1_IDENTIFIER_COLLISION','AGGREGATE_LIMIT_EXCEEDED'];if(!known.includes(e.code))throw e;
    saga=undefined;d=L.decision({state:'FAILED',code:['STALE_REVISION','REVISION_CONFLICT','M1_REVISION_INVALID'].includes(e.code)?'REVISION_CONFLICT':'DENIED',revisionBefore:[],revisionAfter:[]});plan=P.planTerminal(tx,x.t.idPessoa,now(),deadline);
   }
   const n=now();N(n<deadline,'LEASE_LOST');
   if(saga&&!d){if(saga.changed){const writes=[{path:'memberOnboardingRuns/'+key.operationId,data:saga.run,kind:'set'},saga.stageAudit],allowed=new Map(writes.map(w=>[w.path,C.hash(F.encode(w))]));const full=P.appendTechnical(tx,plan,writes,w=>allowed.get(w.path)===C.hash(F.encode(w)));N(now()<deadline,'LEASE_LOST');P.applyMutationPlan(tx,full,now());}return {state:saga.run.stage};}
   const l=L.make(x.r,d,n),r=L.toReceipt(l),t=d.state==='RECONCILIATION_REQUIRED'?null:L.toTombstone(l),audit={schemaVersion:'member-audit-v13',operationId:key.operationId,ownerUid:a.uid,actorPersonId:x.t.idPessoa,capabilities:K.capabilities(x.r),proofDigest:C.hash(F.encode(proofs)),authorizationRefs:proofs.map(p=>({kind:p.kind,capability:p.capability,sourceId:p.kind==='DEV_ADMIN'?p.profilePath:p.grantId||p.policyId,sourceRevision:p.kind==='DEV_ADMIN'?(p.devProfileRevision||p.profileRevision):p.grantRevision||p.policyRevision,authorityEpoch:p.authorityEpoch,policyEpoch:p.policyEpoch,subjectEpoch:p.subjectEpoch,bindingVersion:p.bindingVersion,personRevision:p.personRevision})),targetDigest:C.hash(x.r.target),evidenceRefs:x.r.evidenceRefs,decisionRef:x.r.decisionRef,digest:key.digest,revisionBefore:d.revisionBefore,revisionAfter:d.revisionAfter,resultCode:d.code,createdAt:T(n)};
   const data={'commandLedger':l,'commandReceipts':r,'memberAudit':audit,'commandLeases':x.h.schemaVersion==='member-command-lease-v13-g1'?f4MemberClaim_().finish(x.h,n):{...x.h,deadline:T(n)}};if(t)data.commandTombstones=t;if(outbox&&d.state==='COMPLETED')data.jobOutbox=outbox;if(saga?.terminal)data.memberOnboardingRuns=saga.run;
   const writes=Object.entries(data).map(([col,value])=>({path:col+'/'+key.operationId,data:value,kind:'set'}));
   if(saga?.changed)writes.push(saga.stageAudit);if(!K.act(x.r.commandType)&&!saga&&d.state==='COMPLETED'&&typeof f4MemberFamilyAudit_==='function'){const familyAudit=f4MemberFamilyAudit_(x.r,x.t.idPessoa,n);N(!get(familyAudit.path),'OPERATION_CONFLICT');writes.push(familyAudit);}const allowed=new Map(writes.map(w=>[w.path,C.hash(F.encode(w))]));
   const full=P.appendTechnical(tx,plan,writes,w=>allowed.get(w.path)===C.hash(F.encode(w)));N(now()<deadline,'LEASE_LOST');P.applyMutationPlan(tx,full,now());return clone(d);
  }
  return Object.freeze({authenticate,finalize,readStatus,claim,release});
 }
 return {createLocal};
}

