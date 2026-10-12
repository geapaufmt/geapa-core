// G local switch/quiescence. No exported endpoint, property, trigger or DEV session.
function f4MemberOperations_(){
 const C=f4Core_(),F=f4Firestore_(),N=C.need;
 const clone=v=>F.decode(F.encode(v));
 function quiescence(control,observations){
  f4MemberAuthority_().validate(control);
  C.exact(observations,['schemaVersion','authorityEpoch','packages','executions','requests','leases','outbox','canonicalChanges','legacyReconciled']);
  N(observations.schemaVersion==='member-quiescence-local-v13-g1'&&observations.authorityEpoch===control.epoch,'SCHEMA_INVALID');
  for(const k of ['packages','executions','requests','leases','outbox','canonicalChanges'])N(Array.isArray(observations[k]),'SCHEMA_INVALID');
  // No remote reconciliation certificate is implemented/accepted in this local slice.
  N(observations.legacyReconciled===false,'SCHEMA_INVALID');
  const failures=[];if(control.mode!=='PAUSED')failures.push('NOT_PAUSED');
  const repos=new Set();for(const p of observations.packages){C.exact(p,['repo','hash','epoch','fenced','triggersDisabled','acknowledged']);N(['core','members','portal'].includes(p.repo)&&!repos.has(p.repo)&&/^[a-f0-9]{64}$/.test(p.hash)&&C.rev(p.epoch)&&[p.fenced,p.triggersDisabled,p.acknowledged].every(v=>typeof v==='boolean'),'SCHEMA_INVALID');repos.add(p.repo);if(p.epoch!==control.epoch||!p.fenced||!p.triggersDisabled||!p.acknowledged)failures.push('EXECUTOR_NOT_DRAINED');}
  if(repos.size!==3)failures.push('PACKAGE_MISSING');
  for(const e of observations.executions){C.exact(e,['id','state']);N(C.id(e.id)&&['COMPLETED','ABORTED','RECONCILIATION_REQUIRED','IN_FLIGHT'].includes(e.state),'SCHEMA_INVALID');if(!['COMPLETED','ABORTED'].includes(e.state))failures.push('EXECUTION_UNCLASSIFIED');}
  for(const r of observations.requests){C.exact(r,['operationId','state']);N(/^[a-f0-9]{32}$/.test(r.operationId)&&['COMPLETED','FAILED','FENCED','RECONCILIATION_REQUIRED','RECEIVED'].includes(r.state),'SCHEMA_INVALID');if(!['COMPLETED','FAILED','FENCED'].includes(r.state))failures.push('REQUEST_UNCLASSIFIED');}
  for(const l of observations.leases){f4MemberClaim_().validate(l);if(l.state==='CLAIMED'&&l.authorityEpoch>=control.epoch)failures.push('LIVE_LEASE');}
  for(const o of observations.outbox){C.exact(o,['operationId','state']);N(/^[a-f0-9]{32}$/.test(o.operationId)&&['COMPLETED','FENCED','RECONCILIATION_REQUIRED'].includes(o.state),'SCHEMA_INVALID');if(o.state==='RECONCILIATION_REQUIRED')failures.push('RECONCILIATION_REQUIRED');}
  N(observations.canonicalChanges.every(x=>typeof x==='string'&&/^[a-f0-9]{32}$/.test(x)),'SCHEMA_INVALID');
  return Object.freeze({quiescent:failures.length===0,reasons:[...new Set(failures)],canRestoreLegacy:failures.length===0&&observations.canonicalChanges.length===0});
 }
 function createLocal(options){
  C.exact(options,['config','clock','observe','readiness']);C.destination(options.config);N((options.config.environment==='SYNTHETIC_LOCAL' || f4DevBoundary_().permit(options.config))&&typeof options.observe==='function'&&typeof options.readiness==='function'&&typeof options.clock==='function','DESTINATION_INVALID');
  // Both providers are trusted internal readers. No serialized client certificate is accepted.
  const config=Object.freeze({...options.config});
  function transition(tx,expected,next){
   const i=F.inspectTransaction(tx);N(i.environment===config.environment&&i.projectId===config.projectId&&i.database===config.database,'DESTINATION_INVALID');
   C.exact(expected,['revision','epoch']);N(C.rev(expected.revision)&&C.rev(expected.epoch),'SCHEMA_INVALID');
   const path='memberAuthorityControl/MEMBERS',before=f4MemberAuthority_().validate(tx.get(path).data()),n=options.clock();N(Number.isSafeInteger(n),'SCHEMA_INVALID');
   N(before.revision===expected.revision&&before.epoch===expected.epoch,'REVISION_CONFLICT');
   N((before.mode==='LEGACY'&&next==='PAUSED')||(before.mode==='CANONICAL'&&next==='PAUSED')||(before.mode==='PAUSED'&&['LEGACY','CANONICAL'].includes(next)),'ACCESS_REVOKED');
   let evidence=null;
   if(next!=='PAUSED'){
    evidence=options.observe(tx,clone(before));const q=quiescence(before,evidence);N(q.quiescent,'RECONCILIATION_REQUIRED');
    if(next==='LEGACY')N(q.canRestoreLegacy,'RECONCILIATION_REQUIRED');
    else N(options.readiness(tx,clone(before))==='READY_FOR_M64H','ACCESS_REVOKED');
   }
   const after={...before,revision:before.revision+1,epoch:before.epoch+1,policyEpoch:before.policyEpoch+1,readAuthzEpoch:before.readAuthzEpoch+1,mode:next,activeWriterVersion:next==='CANONICAL'?'M13':'NONE',updatedAt:F.Timestamp.fromMillis(n)};
   f4MemberAuthority_().validate(after);
   const id=C.hash({kind:'MEMBERS_SWITCH',revision:after.revision,epoch:after.epoch}).slice(0,32),auditPath='memberAudit/'+id;N(!tx.get(auditPath).data(),'OPERATION_CONFLICT');
   tx.set(path,after);tx.set(auditPath,{schemaVersion:'member-switch-audit-local-v13-g1',id,before:clone(before),after:clone(after),observationHash:evidence?C.hash(F.encode(evidence)):null,createdAt:F.Timestamp.fromMillis(n)});
   return clone(after);
  }
  return Object.freeze({transition});
 }
 return Object.freeze({quiescence,createLocal});
}

// The enumerator is a protected complete snapshot of the paused local namespace,
// not an operation list supplied by a client. Physical executor journals remain R01.
function f4MemberQuiescence_(){
 const C=f4Core_(),F=f4Firestore_(),N=C.need;
 function collect(tx,control,inventory,now){
  N((F.inspectTransaction(tx).environment==='SYNTHETIC_LOCAL' || f4DevBoundary_().permit(F.inspectTransaction(tx))),'DESTINATION_INVALID');
  f4MemberAuthority_().validate(control);N(control.mode==='PAUSED','ACCESS_REVOKED');
  C.exact(inventory,['schemaVersion','authorityEpoch','complete','operationIds','packages','executions']);
  N(inventory.schemaVersion==='member-paused-inventory-local-v13-g1'&&inventory.complete===true&&inventory.authorityEpoch===control.epoch&&Array.isArray(inventory.operationIds)&&new Set(inventory.operationIds).size===inventory.operationIds.length&&inventory.operationIds.every(C.op),'SCHEMA_INVALID');
  const get=p=>tx.get(p).data(),requests=[],leases=[],outbox=[],canonicalChanges=[];
  for(const id of inventory.operationIds){
   const r=get('commandRequests/'+id);let state='RECONCILIATION_REQUIRED';
   if(r&&r.schemaVersion==='geapa-members-command-v13'&&r.operationId===id){
    const key={operationId:id,commandType:r.commandType,digest:r.digest};
    try{
     const terminal=f4MemberLedgerM13_().recover(get,key,r.ownerUid,now);
     if(terminal){state=terminal.state;if(state==='COMPLETED')canonicalChanges.push(id);}
     else{
      const ticket=get('memberAdmissionTickets/'+r.ticketId),lease=get('commandLeases/'+id);
      if(lease){f4MemberClaim_().validate(lease);leases.push(lease);}
      if(ticket?.schemaVersion==='member-admission-ticket-v13'&&ticket.operationId===id&&ticket.digest===r.digest&&ticket.ownerUid===r.ownerUid&&Number.isSafeInteger(ticket.authorityEpoch)&&ticket.authorityEpoch<control.epoch&&(!lease||lease.authorityEpoch<control.epoch)&&lease?.state!=='RECONCILIATION_REQUIRED')state='FENCED';
     }
    }catch{state='RECONCILIATION_REQUIRED';}
    const o=get('jobOutbox/'+id);if(o){let s='RECONCILIATION_REQUIRED';try{if(f4MemberBoardCommands_().outbox(o,r)&&o.state==='PENDING')s='FENCED';}catch{}outbox.push({operationId:id,state:s});}
   }
   requests.push({operationId:id,state});
  }
  const observation={schemaVersion:'member-quiescence-local-v13-g1',authorityEpoch:control.epoch,packages:inventory.packages,executions:inventory.executions,requests,leases,outbox,canonicalChanges,legacyReconciled:false};
  f4MemberOperations_().quiescence(control,observation);return observation;
 }
 return Object.freeze({collect});
}
