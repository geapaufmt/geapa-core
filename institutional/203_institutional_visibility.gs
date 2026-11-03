// Presentation only. Atomic canonical transaction; no normative or public write.
function f4InstitutionalVisibilityCommands_(){
 const C=f4Core_(),F=f4Firestore_(),N=C.need,I=f4InstitutionalNorms_(),P=f4TransactionPlans_();
 const CONTROL='institutionalVisibilityControl/current',TYPE='institutional.document.visibility';
 const hash=v=>C.hash(F.encode(v)),hex=v=>typeof v==='string'&&/^[a-f0-9]{64}$/.test(v),path=s=>'institutionalPresentation/'+s;
 function control(get,config,identity,now,admit,section,inspection=false){
  C.destination(config);const c=get(CONTROL),actual=f4CommandExecutorId_(),expected=inspection&&actual==='CANONICAL_READ_ONLY'?(f4ExecutionEnvironment_().production?'CF_GEN2_PROD_V1':'CF_GEN2_DEV_EDIT_V1'):actual;
  N(c?.schemaVersion==='institutional-visibility-control-v1'&&c.projectId===config.projectId&&c.executor===expected&&c.publicationEnabled===false&&c.commandType===TYPE&&C.rev(c.revision)&&c.revision>0&&C.rev(c.epoch)&&c.epoch>0&&Array.isArray(c.actorUids)&&c.actorUids.length<=16&&c.actorUids.includes(identity.uid),'ACCESS_REVOKED');
  N(['norms','governance'].includes(section)&&Array.isArray(c.recordIds?.[section])&&c.recordIds[section].length<=150&&c.recordIds[section].every(C.id),'AUTHORIZATION_DENIED');
  N(c.rateLimit?.windowMs===3600000&&Number.isInteger(c.rateLimit.maxAdmissions)&&c.rateLimit.maxAdmissions>0&&c.rateLimit.maxAdmissions<=40,'ACCESS_REVOKED');
  if(admit)N(c.enabled===true&&f4ServiceAvailability_().live(c,now),'PILOT_DISABLED');
  const auth=f4InstitutionalPolicies_(get,identity,now,c.policyIds,'INSTITUTIONAL_DOCUMENT_VISIBILITY','institutionalDirectory/'+section);N(auth.mask.includes('presentation'),'AUTHORIZATION_DENIED');
  // Reading the record remains a separate competence.
  const read=get('institutionalDirectoryControl/current');N(read?.enabled===true,'FEATURE_DISABLED');const r=f4InstitutionalPolicies_(get,identity,now,read.policyIds,'INSTITUTIONAL_DIRECTORY_READ','institutionalDirectory/'+section);N(r.mask.includes('summary')&&r.mask.includes('fields'),'AUTHORIZATION_DENIED');auth.validUntil=Math.min(auth.validUntil,r.validUntil);
  return {c,auth};
 }
 function preview(tx,o,input){
  C.exact(input,['section','recordId','hidden','expectedRevision','expectedSourceRevision','reason']);N(['norms','governance'].includes(input.section)&&C.id(input.recordId)&&typeof input.hidden==='boolean'&&typeof input.reason==='string'&&input.reason===input.reason.trim()&&input.reason.length>=10&&input.reason.length<=1000,'COMMAND_PAYLOAD_INVALID');
  const get=p=>tx.get(p).data(),now=o.clock(),{c,auth}=control(get,o.config,o.identity,now,true,input.section);N(c.recordIds[input.section].includes(input.recordId),'AUTHORIZATION_DENIED');
  const source=get('institutionalDirectory/'+input.section);N(source,'RESOURCE_NOT_FOUND');f4InstitutionalDirectoryModel_().validate(source);const record=source.records.find(r=>r.id===input.recordId);N(record,'RESOURCE_NOT_FOUND');N(I.eligible(record),'DOCUMENT_PRESENTATION_PROTECTED');
  const v=I.overlay(get(path(input.section)),input.section),prior=v.entries.find(e=>e.recordId===input.recordId);N(v.revision===input.expectedRevision&&source.revision===input.expectedSourceRevision,'REVISION_CONFLICT');N((prior?.hidden===true)!==input.hidden,'COMMAND_PAYLOAD_INVALID');
  const history=get('institutionalPresentationHistory/'+input.section);N(!history||(history.schemaVersion==='institutional-presentation-history-v1'&&Array.isArray(history.events)&&history.events.length<500),'SCOPE_TOO_LARGE');
  return {section:input.section,recordId:input.recordId,beforeHidden:prior?.hidden===true,hidden:input.hidden,reason:input.reason,documentHash:hash({source,presentation:v}),epoch:c.epoch,controlRevision:c.revision,validUntil:Math.min(auth.validUntil,f4ServiceAvailability_().bound(c,now+60000)),publicUpdate:false};
 }
 function digest(config,uid,r){return hash({projectId:config.projectId,ownerUid:uid,commandType:TYPE,...r});}
 function status(tx,o,key){C.exact(key,['operationId','digest']);N(C.op(key.operationId)&&hex(key.digest),'COMMAND_PAYLOAD_INVALID');const get=p=>tx.get(p).data(),r=get('institutionalVisibilityReceipts/'+key.operationId),owner=get('commandExecutorOwners/'+key.operationId);
  if(!r){N(!owner,'RECONCILIATION_REQUIRED');return {state:'NOT_ADMITTED',operationId:key.operationId};}
  const {c}=control(get,o.config,o.identity,o.clock(),false,r.section);N(r.ownerUid===o.identity.uid&&r.digest===key.digest&&r.projectId===o.config.projectId&&r.executor===c.executor&&r.state==='COMPLETED'&&owner?.digest===r.digest&&owner.ownerUid===r.ownerUid&&owner.executor===r.executor&&owner.commandType===TYPE,'ACCESS_REVOKED');return {state:'COMPLETED',operationId:r.operationId,revision:r.revisionAfter,recordId:r.recordId,hidden:r.hidden,publicUpdate:false};
 }
 function execute(tx,o,r){C.exact(r,['operationId','input','documentHash','epoch','controlRevision','validUntil','digest']);N(C.op(r.operationId)&&hex(r.documentHash)&&hex(r.digest),'COMMAND_PAYLOAD_INVALID');const {digest:given,...semantic}=r;N(given===digest(o.config,o.identity.uid,semantic),'OPERATION_CONFLICT');
  const get=p=>tx.get(p).data(),now=o.clock();if(get('institutionalVisibilityReceipts/'+r.operationId))return status(tx,o,{operationId:r.operationId,digest:r.digest});
  const q=preview(tx,o,r.input),{c,auth}=control(get,o.config,o.identity,now,true,r.input.section);N(q.documentHash===r.documentHash&&c.epoch===r.epoch&&c.revision===r.controlRevision,'REVISION_CONFLICT');N(Number.isSafeInteger(r.validUntil)&&now<r.validUntil&&r.validUntil<=now+60000,'ACCESS_REVOKED');
  for(const prefix of ['commandExecutorOwners','commandRequests','commandReceipts','commandLedger','institutionalVisibilityAudit'])N(!get(prefix+'/'+r.operationId),'OPERATION_CONFLICT');
  const bucket='institutionalVisibilityRate/'+C.hash(o.identity.uid).slice(0,32),rate=get(bucket),start=Math.floor(now/3600000)*3600000,count=rate?.windowStart===start?rate.count:0;N(Number.isInteger(count)&&count>=0&&count<c.rateLimit.maxAdmissions,'PILOT_LIMIT_REACHED');
  const v=I.overlay(get(path(q.section)),q.section),event={recordId:q.recordId,hidden:q.hidden,reason:q.reason,actorPersonId:auth.actor.actor.h.id,at:now,operationId:r.operationId},next={...v,revision:v.revision+1,entries:[...v.entries.filter(e=>e.recordId!==q.recordId),event]};I.overlay(next,q.section);
  const historyPath='institutionalPresentationHistory/'+q.section,history=get(historyPath)||{schemaVersion:'institutional-presentation-history-v1',events:[]};
  const receipt={schemaVersion:'institutional-visibility-receipt-v1',projectId:o.config.projectId,operationId:r.operationId,ownerUid:o.identity.uid,actorPersonId:auth.actor.actor.h.id,executor:c.executor,epoch:c.epoch,commandType:TYPE,digest:r.digest,section:q.section,recordId:q.recordId,hidden:q.hidden,revisionBefore:v.revision,revisionAfter:next.revision,beforeHash:hash(v),afterHash:hash(next),state:'COMPLETED',publicUpdate:false,at:now};
  const writes=[{path:path(q.section),data:next},{path:historyPath,data:{...history,events:[...history.events,event]}},{path:bucket,data:{windowStart:start,count:count+1}},{path:'commandExecutorOwners/'+r.operationId,data:{schemaVersion:'institutional-command-owner-v1',routeId:'INSTITUTIONAL_DOCUMENT_VISIBILITY',operationId:r.operationId,commandType:TYPE,activityId:path(q.section),ownerUid:o.identity.uid,digest:r.digest,executor:c.executor,epoch:c.epoch,createdAt:F.Timestamp.fromMillis(now)}},{path:'institutionalVisibilityReceipts/'+r.operationId,data:receipt},{path:'institutionalVisibilityAudit/'+r.operationId,data:{...receipt,reason:q.reason,beforeHidden:q.beforeHidden,sourceHash:get('institutionalDirectory/'+q.section).sourceHash,policyRefs:auth.policies.map(p=>({id:p.id,revision:p.revision}))}}];
  const result={state:'COMPLETED',operationId:r.operationId,revision:next.revision,recordId:q.recordId,hidden:q.hidden,publicUpdate:false},plan={kind:'INSTITUTIONAL_VISIBILITY',aggregateId:q.section,sourceWrites:writes.slice(0,2).map(w=>({...w,kind:'set'})),headWrites:[],claimWrites:[],projectionWrites:[],technicalWrites:writes.slice(2).map(w=>({...w,kind:'set'})),deletes:[]};return P.applyMutationPlan(tx,P.finalize(tx,plan,result,now,Math.min(auth.validUntil,r.validUntil)),o.clock());
 }
 return {control,preview,execute,status,digest,CONTROL,TYPE};
}
