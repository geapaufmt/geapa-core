// Protected, directed consumer. Never emits a business decision or external effect.
// A committed governance command identifies work; current canonical sources determine output.
function f4MemberReprojection_(){
 const C=f4Core_(),F=f4Firestore_(),N=C.need,hash=v=>C.hash(F.encode(v));
 function basis(get,control,loaded){const paths=[...new Set([...loaded.reg.paths,...Object.keys(loaded.docs),...Object.keys(loaded.boards).map(id=>'memberBoardHeads/'+id)])].sort();return hash({personRevision:loaded.reg.revision,authorityEpoch:control.epoch,readAuthzEpoch:control.readAuthzEpoch,sources:paths.map(path=>({path,data:get(path)}))});}
 const receiptPath=(operationId,personId,basisHash)=>'memberAudit/'+C.hash({kind:'MEMBERS_REPROJECTION',operationId,personId,basisHash}).slice(0,32);
 function confirmed(get,r,now){
  const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),outbox=get('jobOutbox/'+r.operationId);f4MemberBoardCommands_().outbox(outbox,r);
  return outbox.personIds.every(personId=>{const loaded=f4MemberRepository_(control).load(get,personId),basisHash=basis(get,control,loaded),v=get(receiptPath(r.operationId,personId,basisHash));return !!v&&v.schemaVersion==='member-projection-receipt-v1'&&v.operationId===r.operationId&&v.personId===personId&&v.sourceDigest===r.digest&&v.basisHash===basisHash&&now<C.ms(v.validUntil);});
 }
 function process(tx,operationId,personId,now){
  N(C.op(operationId)&&C.id(personId)&&Number.isSafeInteger(now),'COMMAND_PAYLOAD_INVALID');
  const info=F.inspectTransaction(tx);N(info.environment==='SYNTHETIC_LOCAL'||f4OperationalDestination_(info),'DESTINATION_INVALID');
  const get=p=>tx.get(p).data(),control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS'));
  N(control.mode==='CANONICAL'&&control.activeWriterVersion==='M13','ACCESS_REVOKED');
  const r=get('commandRequests/'+operationId);N(r&&r.operationId===operationId,'RECONCILIATION_REQUIRED');
  f4SharedCommandCatalog_().envelope(info.projectId,r,f4MemberBoardCommands_().validate);
  const outbox=get('jobOutbox/'+operationId);f4MemberBoardCommands_().outbox(outbox,r);N(outbox.personIds.includes(personId),'ACCESS_REVOKED');
  const result=f4MemberLedgerM13_().recover(get,{operationId,commandType:r.commandType,digest:r.digest},r.ownerUid,now);
  N(result?.state==='COMPLETED'&&result.code==='OK','RECONCILIATION_REQUIRED');
  const repo=f4MemberRepository_(control),loaded=repo.load(get,personId),basisHash=basis(get,control,loaded),auditPath=receiptPath(operationId,personId,basisHash),prior=get(auditPath);
  if(prior){C.exact(prior,['schemaVersion','operationId','personId','sourceDigest','basisHash','validUntil','projectionHash']);N(prior.schemaVersion==='member-projection-receipt-v1'&&prior.operationId===operationId&&prior.personId===personId&&prior.sourceDigest===r.digest&&prior.basisHash===basisHash,'RECONCILIATION_REQUIRED');return {state:'REPLAY',writes:0};}
  const output=repo.projections(loaded.docs,personId,loaded.reg.revision,now,loaded),writes=[];
  // Only the repository's derived surface is accepted. Sources, bindings and grants are excluded.
  for(const [path,data]of Object.entries(output)){
   N(/^(memberCurrentStates|memberPublicProfiles|memberCurrentFunctions|memberCurrentDirectors|institutionalAccounts|memberPortalProfiles|memberStatusViews|projectionProvenance|memberTransitionQueue|portalAccessStatus|memberFunctionProofs)\//.test(path),'M13_PRODUCER_COLLISION');
   const old=get(path);if(old===undefined||hash(old)!==hash(data))writes.push({path,kind:'set',data});
  }
  for(const collection of ['memberCurrentFunctions','memberCurrentDirectors'])for(const id of get(collection+'/'+personId+'/index/current')?.ids||[]){
   N(C.id(id),'SCHEMA_INVALID');for(const path of [collection+'/'+personId+'/items/'+id,...(collection==='memberCurrentFunctions'?['memberFunctionProofs/'+personId+'/items/'+id]:[])])if(!Object.hasOwn(output,path)&&get(path))writes.push({path,kind:'delete',data:null});
  }
  const due=output['memberTransitionQueue/'+personId].dueAt;
  writes.push({path:auditPath,kind:'set',data:{schemaVersion:'member-projection-receipt-v1',operationId,personId,sourceDigest:r.digest,basisHash,validUntil:due,projectionHash:hash(Object.entries(output).map(([path,data])=>({path,data})))}});
  // Preserve the same shared plan limits and single atomic write phase.
  const encoded=F.previewTransaction(tx,writes),bytes=v=>unescape(encodeURIComponent(JSON.stringify(v))).length;
  f4TransactionPlans_().assertBudget({writes:writes.length,reads:F.inspectTransaction(tx).readSet.length,requestBytes:bytes(encoded),documentBytes:Math.max(...encoded.writes.map(w=>w.update?bytes(w.update):0))});
  for(const w of writes)if(w.kind==='delete')tx.delete(w.path);else tx.set(w.path,w.data);
  return {state:'CONFIRMED',writes:writes.length};
 }
 return Object.freeze({process,confirmed});
}
