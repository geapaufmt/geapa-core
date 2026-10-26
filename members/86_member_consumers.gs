// F local status egress. D remains the sole owner/replay authorization boundary.
// No worker, admission, grant, endpoint or external consumer is installed here.
function f4MemberConsumers_(){
 const C=f4Core_(),F=f4Firestore_(),K=f4MemberCommandCatalog_(),L=f4MemberLedgerM13_(),N=C.need;
 function createLocal(options){
  C.exact(options,['config','recovery']);C.destination(options.config);N((options.config.environment==='SYNTHETIC_LOCAL' || f4OperationalDestination_(options.config))&&typeof options.recovery?.readStatus==='function','DESTINATION_INVALID');
  const config={...options.config},recovery=options.recovery;
  function status(tx,principal,key){
   const info=F.inspectTransaction(tx);N((info.environment==='SYNTHETIC_LOCAL' || f4OperationalDestination_(info))&&info.projectId===config.projectId&&info.database===config.database,'DESTINATION_INVALID');
   K.selector(key);const d=recovery.readStatus(tx,principal,key); // Revalidates owner binding and SELF access, not original capability.
   let code=d.code;
   if(Object.hasOwn(d,'revisionBefore'))L.decision(d);
   else if(code==='RESULT_RETAINED_AS_TOMBSTONE'){C.exact(d,['state','code']);N(['COMPLETED','FAILED'].includes(d.state),'SCHEMA_INVALID');}
   else{C.exact(d,['state']);N(['INGRESS_PENDING','BINDING_PENDING'].includes(d.state),'SCHEMA_INVALID');code='PENDING';}
   let projection='NOT_CONFIRMED',projectionKind='UNSPECIFIED';
   if(d.state==='COMPLETED'&&code==='OK'){
    if(key.commandType==='members.governance.boardTerm'){
     const r=tx.get('commandRequests/'+key.operationId).data(),outbox=tx.get('jobOutbox/'+key.operationId).data();
     N(r&&r.digest===key.digest&&f4MemberBoardCommands_().outbox(outbox,r),'RECONCILIATION_REQUIRED');projection=f4MemberReprojection_().confirmed(p=>tx.get(p).data(),r,Date.now())?'CONFIRMED':'PENDING';projectionKind='GOVERNANCE_REPROJECTION';
    }else if(key.commandType==='members.self.update'){
     const r=tx.get('commandRequests/'+key.operationId).data();N(r&&r.digest===key.digest,'RECONCILIATION_REQUIRED');projectionKind='OWN_DISPLAY';
     const b=tx.get('memberAuthBindings/'+r.ownerUid).data();// A completed old receipt remains readable by its authenticated owner.
     // A later binding must not confirm the former person's projection.
     if(!b||b.idPessoa!==r.target.personId)return {operationId:key.operationId,receipt:{state:d.state,code},canonical:{state:'COMMITTED'},projection:{kind:projectionKind,state:'NOT_CONFIRMED'}};
     const expected=d.revisionAfter.find(x=>x.resourceType==='memberAggregateHeads'&&x.id===b.idPessoa),p=tx.get('memberPortalProfiles/'+b.idPessoa).data(),h=tx.get('memberAggregateHeads/'+b.idPessoa).data();
     // Only the synchronous own display projection is confirmed, never public/contact/outbox.
     if(expected&&p&&h&&h.revision===expected.revision&&p.sourceRevision===expected.revision){C.exact(p,['schemaVersion','idPessoa','displayLabel','sourceRevision']);N(p.schemaVersion==='safe-profile-v11'&&p.idPessoa===b.idPessoa&&typeof p.displayLabel==='string','SCHEMA_INVALID');projection='CONFIRMED';}
    }
   }
   return {operationId:key.operationId,receipt:{state:d.state,code},canonical:{state:d.state==='COMPLETED'?'COMMITTED':d.state==='FAILED'?'NOT_COMMITTED':'UNCONFIRMED'},projection:{kind:projectionKind,state:projection}};
  }
  return Object.freeze({status});
 }
 return {createLocal};
}
