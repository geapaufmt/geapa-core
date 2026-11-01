// Protected local admission. No evidence issuer, HTTP endpoint or browser-supplied authority.
function f4MemberAdmission_(){
 const C=f4Core_(),F=f4Firestore_(),K=f4SharedCommandCatalog_(),A=f4SharedAuthz_(),L=f4MemberLedgerM13_(),N=C.need,T=F.Timestamp.fromMillis;
 function createLocal(options){
  C.exact(options,['config','clock','verifyIdentity','resolveAuthorization','resolveRevisions']);C.destination(options.config);
  N((options.config.environment==='SYNTHETIC_LOCAL' || f4OperationalDestination_(options.config))&&[options.clock,options.verifyIdentity,options.resolveAuthorization,options.resolveRevisions].every(x=>typeof x==='function'),'DESTINATION_INVALID');
  const config=Object.freeze({...options.config}),families=f4SharedFamilyCatalog_(config),seals=new WeakMap();
  function authenticate(token){const v=options.verifyIdentity(token,options.clock());C.exact(v,['uid','expiresAt']);N(C.id(v.uid)&&Number.isSafeInteger(v.expiresAt)&&options.clock()<v.expiresAt,'AUTH_BINDING_INVALID');const s=Object.freeze({});seals.set(s,v);return s;}
  function issue(tx,principal,input){
   const identity=seals.get(principal),n=options.clock(),info=F.inspectTransaction(tx);
   N(identity&&Number.isSafeInteger(n)&&n<identity.expiresAt,'AUTH_BINDING_INVALID');N(info.environment===config.environment&&info.projectId===config.projectId&&info.database===config.database,'DESTINATION_INVALID');
   C.exact(input,['schemaVersion','operationId','ownerUid','commandType','target','expectedRevisions','payload','evidenceRefs','decisionRef']);
   N(input.ownerUid===identity.uid&&Object.hasOwn(K.types,input.commandType),'AUTH_BINDING_INVALID');
   const get=p=>tx.get(p).data(),digest=K.digest(config.projectId,input),key={operationId:input.operationId,commandType:input.commandType,digest};K.selector(key);
   if(K.act(input.commandType))f4ActivitiesPrefetchCommand_(tx,identity.uid,key);
   // Existing operations go to the current-access status path; never issue a new execution ticket.
   const prior=L.recover(get,key,identity.uid,n);N(!prior&&!get('commandRequests/'+key.operationId),'OPERATION_CONFLICT');
   const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS'));N(control.mode==='CANONICAL'&&control.activeWriterVersion==='M13','ACCESS_REVOKED');
   const enrollment=f4MemberEnrollmentCommands_().is(input),candidate=enrollment?f4MemberEnrollmentAccess_().source(get,identity.uid,n):null;const binding=candidate?{active:true,idPessoa:candidate.enrollment.personId}:K.act(input.commandType)&&f4ActivitiesSystem_().is(input)&&identity.uid==='CORE_SYSTEM'?{active:true,idPessoa:'CORE_SYSTEM'}:get('memberAuthBindings/'+identity.uid);N(binding&&binding.active===true&&C.id(binding.idPessoa),'AUTH_BINDING_INVALID');
   const actLimits=K.act(input.commandType)&&typeof f4ActivitiesAdmissionLimits_==='function'?f4ActivitiesAdmissionLimits_(identity.uid):null; const cp='memberAdmissionCounters/'+(actLimits?.counterId||identity.uid),old=get(cp);
   if(old){C.exact(old,['schemaVersion','uid','revision','windowStart','issuedCount','pendingTicketIds']);N(old.schemaVersion==='member-admission-counter-v13'&&old.uid===identity.uid&&C.rev(old.revision)&&Number.isInteger(old.issuedCount)&&old.issuedCount>=0&&old.issuedCount<=(actLimits?.maximum||10)&&Array.isArray(old.pendingTicketIds)&&old.pendingTicketIds.length<=2&&new Set(old.pendingTicketIds).size===old.pendingTicketIds.length&&old.pendingTicketIds.every(C.id)&&Number.isSafeInteger(C.ms(old.windowStart))&&C.ms(old.windowStart)<=n,'SCHEMA_INVALID');}
   const pending=[];
   for(const id of old?.pendingTicketIds||[]){const t=get('memberAdmissionTickets/'+id);N(t&&t.ownerUid===identity.uid&&t.id===id&&t.schemaVersion==='member-admission-ticket-v13'&&t.status==='ISSUED'&&Number.isSafeInteger(C.ms(t.expiresAt)),'RECONCILIATION_REQUIRED');const existing=L.recover(get,{operationId:t.operationId,commandType:t.requestEnvelope.commandType,digest:t.digest},identity.uid,n);if(n<C.ms(t.expiresAt)&&!existing)pending.push(id);}
   const previous=get('commandProvenance/'+input.operationId);
   if(previous){N(previous.ownerUid===identity.uid&&previous.digest===digest&&previous.commandType===undefined,'OPERATION_CONFLICT');const t=get('memberAdmissionTickets/'+previous.ticketId);N(t&&t.digest===digest,'RECONCILIATION_REQUIRED');}
   const windowStart=!old||n-C.ms(old.windowStart)>=3600000?n:C.ms(old.windowStart),count=old&&windowStart===C.ms(old.windowStart)?old.issuedCount:0;
   const existing=previous?get('memberAdmissionTickets/'+previous.ticketId):null;
   const reuse=existing&&n<C.ms(existing.expiresAt);
   if(!reuse)N(pending.length<2&&count<(actLimits?.limit||10),'ACCESS_REVOKED');
   // Emission revision is persistent; wall-clock time is validity, not identity.
   // Retain historical IDs when recovering an existing immutable ticket.
   const revision=(old?.revision||0)+1,id=reuse?existing.id:C.hash({projectId:config.projectId,operationId:input.operationId,digest,revision}).slice(0,32);
   if(K.act(input.commandType))tx.getAll(['memberAdmissionTickets/'+id,'memberAdmissionUses/'+id]);
   N((reuse||!get('memberAdmissionTickets/'+id))&&!get('memberAdmissionUses/'+id),'OPERATION_CONFLICT');
   let deadline=reuse?C.ms(existing.expiresAt):Math.min(n+120000,identity.expiresAt);
   const r={...input,ticketId:id,digest,receivedAt:T(n),state:'RECEIVED'};
   K.envelope(config.projectId,r,families[r.commandType].validate);
   const requests=options.resolveAuthorization(get,F.decode(F.encode(r)),Object.freeze({...identity}),n,deadline);
   N(Array.isArray(requests)&&requests.length===K.capabilities(r).length,'AUTH_BINDING_INVALID');
   let proofs=requests.map((q,i)=>{
    N(q.uid===identity.uid&&q.capability===K.capabilities(r)[i]&&C.ms(q.tokenExpiresAt)===identity.expiresAt&&C.ms(q.admissionExpiresAt)===deadline,'AUTH_BINDING_INVALID');
    const result=q.schemaVersion==='member-authz-governance-request-v13'?A.authorizeGovernance(get,q,r,n):r.commandType==='members.ingress.execute'?A.authorizeCommand(get,q,r,n):A.authorize(get,q,n);
    N(result.proof.idPessoa===binding.idPessoa&&result.proof.authorityEpoch===control.epoch,'AUTH_BINDING_INVALID');return result.proof;
   });
   // If an authorization expires earlier, reissue the proof against that exact shortened ticket.
   deadline=Math.min(deadline,...proofs.map(p=>C.ms(p.expiresAt)));N(n<deadline,'ACCESS_REVOKED');
   let finalRequests=requests.map(q=>({...q,admissionExpiresAt:T(deadline)}));
   proofs=finalRequests.map(q=>(q.schemaVersion==='member-authz-governance-request-v13'?A.authorizeGovernance(get,q,r,n):r.commandType==='members.ingress.execute'?A.authorizeCommand(get,q,r,n):A.authorize(get,q,n)).proof);
   if(reuse){
    N(old&&old.issuedCount>0&&old.pendingTicketIds.includes(id),'RECONCILIATION_REQUIRED');
    C.exact(existing,['schemaVersion','id','operationId','ownerUid','idPessoa','requestEnvelope','digest','authProofs','issuedAt','expiresAt','tokenExpiresAt','authorityEpoch','status']);
    C.exact(previous,['schemaVersion','operationId','ownerUid','ticketId','digest','authRequests']);
    N(existing.schemaVersion==='member-admission-ticket-v13'&&existing.id===previous.ticketId&&existing.operationId===r.operationId&&existing.ownerUid===identity.uid&&existing.idPessoa===binding.idPessoa&&existing.digest===digest&&existing.authorityEpoch===control.epoch&&existing.status==='ISSUED'&&K.equal(existing.requestEnvelope,K.template(r))&&previous.schemaVersion==='member-command-provenance-v13'&&previous.operationId===r.operationId,'AUTH_BINDING_INVALID');
    N(C.time(existing.issuedAt,existing.expiresAt,n)&&C.ms(existing.expiresAt)<=C.ms(existing.issuedAt)+120000&&C.ms(existing.expiresAt)<=C.ms(existing.tokenExpiresAt)&&C.ms(existing.tokenExpiresAt)<=identity.expiresAt&&deadline===C.ms(existing.expiresAt),'ACCESS_REVOKED');
    N(Array.isArray(existing.authProofs)&&existing.authProofs.length===proofs.length&&Array.isArray(previous.authRequests)&&previous.authRequests.length===finalRequests.length,'AUTH_BINDING_INVALID');
    const stable=p=>Object.fromEntries(Object.entries(p).filter(([k])=>!['issuedAt','expiresAt'].includes(k)));
    for(let i=0;i<proofs.length;i++){
     const priorProof=existing.authProofs[i];
     N(K.equal({...finalRequests[i],tokenExpiresAt:existing.tokenExpiresAt},previous.authRequests[i])&&K.equal(stable(priorProof),stable(proofs[i]))&&C.time(priorProof.issuedAt,priorProof.expiresAt,n)&&C.ms(priorProof.issuedAt)===C.ms(existing.issuedAt)&&C.ms(priorProof.expiresAt)>=deadline,'AUTH_BINDING_INVALID');
    }
    finalRequests=previous.authRequests;proofs=existing.authProofs;
   }
   const ticket=reuse?existing:{schemaVersion:'member-admission-ticket-v13',id,operationId:r.operationId,ownerUid:identity.uid,idPessoa:binding.idPessoa,requestEnvelope:K.template(r),digest,authProofs:proofs,issuedAt:T(n),expiresAt:T(deadline),tokenExpiresAt:T(identity.expiresAt),authorityEpoch:control.epoch,status:'ISSUED'};
   const provenance={schemaVersion:'member-command-provenance-v13',operationId:r.operationId,ownerUid:identity.uid,ticketId:id,digest,authRequests:finalRequests};
   const plannedGet=path=>path==='commandRequests/'+r.operationId?r:path==='commandProvenance/'+r.operationId?provenance:path==='memberAdmissionTickets/'+id?ticket:get(path);
   if(K.act(r.commandType))f4ActivitiesPipeline_().preflight(plannedGet,r,finalRequests,proofs,n);
   else if(r.commandType==='members.onboarding.complete'&&!enrollment)f4MemberOnboarding_().evaluate(plannedGet,r,finalRequests,proofs,n,config);
   else{
    const ctx=families[r.commandType].resolve(plannedGet,r,n),headType=r.commandType==='members.governance.boardTerm'?'memberBoardHeads':'memberAggregateHeads';
    N(K.equal(ctx.authorization.request,finalRequests[0])&&K.equal(ctx.authorization.proof,proofs[0]),'AUTH_BINDING_INVALID');
    N(ctx.expected===r.expectedRevisions.find(v=>v.resourceType===headType&&v.id===ctx.id)?.revision,'REVISION_CONFLICT');
    const used=new Set([headType+'/'+ctx.id,...ctx.changes.map(c=>c.path),...Object.keys(ctx.boardExpected).map(id=>'memberBoardHeads/'+id)]);
    N(r.expectedRevisions.every(v=>used.has(v.resourceType+'/'+v.id)),'COMMAND_PAYLOAD_INVALID');
    for(const c of ctx.changes){const [resourceType,id]=c.path.split('/');N(c.expectedRevision===r.expectedRevisions.find(v=>v.resourceType===resourceType&&v.id===id)?.revision,'REVISION_CONFLICT');}
    for(const [id,revision]of Object.entries(ctx.boardExpected))N(revision===r.expectedRevisions.find(v=>v.resourceType==='memberBoardHeads'&&v.id===id)?.revision,'REVISION_CONFLICT');
    if(Object.hasOwn(ctx,'outbox'))delete ctx.outbox;
    // The final planner requires the persisted request/use pair (not yet created).
    // Family resolution checks act/target/preconditions now; C/D repeats all checks at commit.
   }
   const ownership=f4ExecutorOwnership_().admission(get,r,config,n);
   if(!reuse){
    ownership.forEach(w=>tx.set(w.path,w.data));
    tx.set('memberAdmissionTickets/'+id,ticket);
    tx.set('commandProvenance/'+r.operationId,provenance);
    tx.set(cp,{schemaVersion:'member-admission-counter-v13',uid:identity.uid,revision,windowStart:T(windowStart),issuedCount:count+1,pendingTicketIds:pending.concat(id)});
   }
   // Never return proofs, binding, idPessoa or internal provenance to a browser.
   return {ticketId:id,digest,expiresAt:ticket.expiresAt,requestEnvelope:ticket.requestEnvelope};
  }
  function issueBrowser(tx,principal,input){
   C.exact(input,['schemaVersion','operationId','commandType','target','payload','evidenceRefs','decisionRef']);
   const identity=seals.get(principal);N(identity&&options.clock()<identity.expiresAt,'AUTH_BINDING_INVALID');
   N(!f4ActivitiesSystem_().is(input),'ACCESS_REVOKED');
   N(!['members.binding.associate','members.binding.revoke'].includes(input.commandType),'ACCESS_REVOKED');
   if(K.act(input.commandType)){const M=f4ActivitiesModel_(),refs=input.target?.expected||[];tx.getAll(refs.map(e=>M.path({kind:e.kind,id:e.id,...(e.parent?{parent:e.parent}:{})})));}
   const revisions=options.resolveRevisions(p=>tx.get(p).data(),F.decode(F.encode(input)),Object.freeze({...identity}),options.clock());
   const value=issue(tx,principal,{...input,ownerUid:identity.uid,expectedRevisions:revisions});
   return {requestEnvelope:value.requestEnvelope,ticketId:value.ticketId,expiresAt:C.ms(value.expiresAt)};
  }
  return Object.freeze({authenticate,issue,issueBrowser});
 }
 return Object.freeze({createLocal});
}

