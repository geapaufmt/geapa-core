// One atomic edit, no external effect or unfinished admission. The shared SDK
// transaction planner seals the entire read set and applies it once. Recovery
// only returns an existing receipt; it never reconstructs or replays an edit.
function f4InstitutionalContactCommands_(){
 const C=f4Core_(),F=f4Firestore_(),N=C.need,M=f4InstitutionalDirectoryModel_(),P=f4TransactionPlans_();
 const CONTROL='institutionalEditControl/current',PATH='institutionalDirectory/identity',TYPE='institutional.contact.update';
 const hash=v=>C.hash(F.encode(v)),hex=x=>typeof x==='string'&&/^[a-f0-9]{64}$/.test(x);
 function control(get,config,identity,now,admit){
  C.destination(config);const c=get(CONTROL);
  N(c?.schemaVersion==='institutional-edit-control-v1'&&c.projectId===config.projectId&&c.resource===PATH&&c.publicationEnabled===false&&c.commandType===TYPE&&c.executor===f4CommandExecutorId_()&&C.rev(c.epoch)&&c.epoch>0&&Array.isArray(c.actorUids)&&c.actorUids.length>0&&c.actorUids.length<=16&&c.actorUids.includes(identity.uid),'ACCESS_REVOKED');
  N(Array.isArray(c.fields)&&c.fields.length>0&&c.fields.length<=6&&new Set(c.fields).size===c.fields.length&&c.fields.every(f=>Object.hasOwn(M.editable,f)),'ACCESS_REVOKED');
  N(c.rateLimit?.windowMs===3600000&&Number.isInteger(c.rateLimit.maxAdmissions)&&c.rateLimit.maxAdmissions>0&&c.rateLimit.maxAdmissions<=20,'ACCESS_REVOKED');
  if(admit)N(c.enabled===true&&f4ServiceAvailability_().live(c,now),'PILOT_DISABLED');
  const auth=f4InstitutionalPolicies_(get,identity,now,c.policyIds,'INSTITUTIONAL_CONTACT_EDIT',PATH);
  return {c,auth};
 }
 function fields(v,input,c,auth){
  // Field authorization precedes revision checks, including with a current revision.
  N(c.fields.includes(input.fieldId)&&auth.mask.includes(input.fieldId),'AUTHORIZATION_DENIED');
  return M.prepareContact(v,input);
 }
 function preview(tx,options,input){
  const get=p=>tx.get(p).data(),now=options.clock(),{c,auth}=control(get,options.config,options.identity,now,true),v=get(PATH);N(v,'NOT_FOUND');
  const p=fields(v,input,c,auth),prior=v.records.find(r=>r.id===p.fieldId).fields.find(f=>f.label==='Valor').value;
  N(prior!==p.value,'COMMAND_PAYLOAD_INVALID');
  return {...p,admissionEnabled:true,beforeValue:prior,documentHash:hash(v),controlRevision:c.revision,epoch:c.epoch,validUntil:Math.min(auth.validUntil,f4ServiceAvailability_().bound(c,now+60000)),publicUpdate:false};
 }
 function status(tx,options,key){
  C.exact(key,['operationId','digest']);N(C.op(key.operationId)&&hex(key.digest),'COMMAND_PAYLOAD_INVALID');
  const get=p=>tx.get(p).data(),now=options.clock(),{c,auth}=control(get,options.config,options.identity,now,false),r=get('institutionalContactReceipts/'+key.operationId),o=get('commandExecutorOwners/'+key.operationId);
  if(!r){N(!o,'RECONCILIATION_REQUIRED');return {state:'NOT_ADMITTED',operationId:key.operationId};}
  N(r.ownerUid===options.identity.uid&&r.digest===key.digest&&r.projectId===options.config.projectId&&r.executor===c.executor&&o?.digest===r.digest&&o.ownerUid===r.ownerUid&&o.executor===r.executor&&o.commandType===TYPE&&o.epoch===r.epoch&&r.state==='COMPLETED'&&auth.mask.includes(r.fieldId),'ACCESS_REVOKED');
  return {state:r.state,operationId:r.operationId,revision:r.revisionAfter,fieldId:r.fieldId,publicUpdate:false};
 }
 function digest(config,uid,r){return hash({projectId:config.projectId,ownerUid:uid,commandType:TYPE,...r});}
 function execute(tx,options,r){
  C.exact(r,['operationId','input','documentHash','epoch','controlRevision','validUntil','digest']);
  N(C.op(r.operationId)&&hex(r.documentHash)&&hex(r.digest),'COMMAND_PAYLOAD_INVALID');
  const {digest:given,...semantic}=r;N(given===digest(options.config,options.identity.uid,semantic),'OPERATION_CONFLICT');
  const get=p=>tx.get(p).data(),now=options.clock(),prior=get('institutionalContactReceipts/'+r.operationId);
  if(prior)return status(tx,options,{operationId:r.operationId,digest:r.digest});
  const q=preview(tx,options,r.input),{c,auth}=control(get,options.config,options.identity,now,true),v=get(PATH);
  N(r.epoch===c.epoch&&r.controlRevision===c.revision&&q.documentHash===r.documentHash,'REVISION_CONFLICT');
  N(Number.isSafeInteger(r.validUntil)&&now<r.validUntil&&r.validUntil<=now+60000,'ACCESS_REVOKED');
  N(!get('commandExecutorOwners/'+r.operationId)&&!get('commandRequests/'+r.operationId)&&!get('commandReceipts/'+r.operationId)&&!get('commandLedger/'+r.operationId),'OPERATION_CONFLICT');
  const bucket='institutionalContactRate/'+C.hash(options.identity.uid).slice(0,32),old=get(bucket),start=Math.floor(now/3600000)*3600000,count=old?.windowStart===start?old.count:0;
  N(Number.isInteger(count)&&count>=0&&count<c.rateLimit.maxAdmissions,'PILOT_LIMIT_REACHED');
  const next={...v,revision:v.revision+1,records:v.records.map(row=>row.id===r.input.fieldId?{...row,fields:row.fields.map(f=>f.label==='Valor'?{...f,value:r.input.value}:f)}:row)};M.validate(next);
  const stamp=F.Timestamp.fromMillis(now),receipt={schemaVersion:'institutional-contact-receipt-v1',projectId:options.config.projectId,operationId:r.operationId,ownerUid:options.identity.uid,actorPersonId:auth.actor.actor.h.id,commandType:TYPE,executor:c.executor,epoch:c.epoch,digest:r.digest,fieldId:r.input.fieldId,revisionBefore:v.revision,revisionAfter:next.revision,beforeHash:hash(v),afterHash:hash(next),state:'COMPLETED',publicUpdate:false,at:stamp};
  const writes=[{path:PATH,data:next},{path:bucket,data:{windowStart:start,count:count+1}},{path:'commandExecutorOwners/'+r.operationId,data:{schemaVersion:'institutional-command-owner-v1',routeId:'INSTITUTIONAL_CONTACT_EDIT',operationId:r.operationId,commandType:TYPE,activityId:PATH,ownerUid:options.identity.uid,digest:r.digest,executor:c.executor,epoch:c.epoch,createdAt:stamp}},{path:'institutionalContactReceipts/'+r.operationId,data:receipt},{path:'institutionalContactAudit/'+r.operationId,data:{...receipt,reason:r.input.reason,beforeValue:q.beforeValue,afterValue:r.input.value,policyRefs:auth.policies.map(p=>({id:p.id,revision:p.revision})),sourceHash:v.sourceHash}}];
  N(!get('institutionalContactAudit/'+r.operationId),'OPERATION_CONFLICT');
  const result={state:'COMPLETED',operationId:r.operationId,revision:next.revision,fieldId:r.input.fieldId,publicUpdate:false};
  const plan={kind:'INSTITUTIONAL_CONTACT',aggregateId:'identity',sourceWrites:writes.slice(0,1).map(w=>({...w,kind:'set'})),headWrites:[],claimWrites:[],projectionWrites:[],technicalWrites:writes.slice(1).map(w=>({...w,kind:'set'})),deletes:[]};
  return P.applyMutationPlan(tx,P.finalize(tx,plan,result,now,Math.min(auth.validUntil,r.validUntil)),options.clock());
 }
 return {control,preview,execute,status,digest,CONTROL,PATH,TYPE};
}
