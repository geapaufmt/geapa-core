// Administrative drafts only. Original documentary rows, approvals, validity and
// publication are never written by this contract. A reference is not an approval.
function f4InstitutionalGovernance_(){
 const C=f4Core_(),F=f4Firestore_(),N=C.need,PATH='institutionalGovernanceDrafts/current',CONTROL='institutionalGovernanceControl/current',TYPE='institutional.governance.draft.save',CAP='INSTITUTIONAL_GOVERNANCE_DRAFT';
 const hash=v=>C.hash(F.encode(v)),hex=v=>typeof v==='string'&&/^[a-f0-9]{64}$/.test(v);
 const text=(v,max,min=0)=>typeof v==='string'&&v===v.trim()&&v.length>=min&&v.length<=max&&!/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(v);
 function state(get){const v=get(PATH)||{schemaVersion:'institutional-governance-drafts-v1',revision:0,records:[]};N(v.schemaVersion==='institutional-governance-drafts-v1'&&C.rev(v.revision)&&Array.isArray(v.records)&&v.records.length<=100&&new Set(v.records.map(r=>r.id)).size===v.records.length,'SCHEMA_INVALID');return v;}
 function control(get,config,identity,now,admit,inspection=false){
  C.destination(config);const c=get(CONTROL),actual=f4CommandExecutorId_(),executor=inspection&&actual==='CANONICAL_READ_ONLY'?(f4ExecutionEnvironment_().production?'CF_GEN2_PROD_V1':'CF_GEN2_DEV_EDIT_V1'):actual;
  N(c?.schemaVersion==='institutional-governance-control-v1'&&c.projectId===config.projectId&&c.executor===executor&&c.resource===PATH&&c.commandType===TYPE&&c.publicationEnabled===false&&c.formalApprovalEnabled===false&&c.appCheck?.mode==='REQUIRED'&&C.rev(c.revision)&&c.revision>0&&C.rev(c.epoch)&&c.epoch>0,'ACCESS_REVOKED');
  N(Array.isArray(c.actorUids)&&c.actorUids.length<=16&&c.actorUids.includes(identity.uid),'AUTHORIZATION_DENIED');
  if(admit)N(c.enabled===true&&f4ServiceAvailability_().live(c,now),'PILOT_DISABLED');
  N(c.rateLimit?.windowMs===3600000&&Number.isInteger(c.rateLimit.maxAdmissions)&&c.rateLimit.maxAdmissions>0&&c.rateLimit.maxAdmissions<=20,'ACCESS_REVOKED');
  const auth=f4InstitutionalPolicies_(get,identity,now,c.policyIds,CAP,PATH);N(auth.mask.includes('draft'),'AUTHORIZATION_DENIED');return {c,auth};
 }
 function prepare(get,input,v){
  C.exact(input,['id','expectedRevision','expectedRecordRevision','title','meetingDate','body','documentaryState','sourceRecordId','references','reason']);
  N(v.revision===input.expectedRevision&&C.rev(input.expectedRecordRevision),'REVISION_CONFLICT');
  N(input.id===''||/^GOV_[a-f0-9]{32}$/.test(input.id),'COMMAND_PAYLOAD_INVALID');
  N(text(input.title,240,5)&&text(input.body,8000,10)&&text(input.reason,1000,10)&&['DRAFT','UNDER_REVIEW','EVIDENCE_PENDING'].includes(input.documentaryState),'COMMAND_PAYLOAD_INVALID');
  N(input.meetingDate===null||typeof input.meetingDate==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(input.meetingDate)&&new Date(input.meetingDate+'T12:00:00Z').toISOString().slice(0,10)===input.meetingDate,'COMMAND_PAYLOAD_INVALID');
  const old=v.records.find(r=>r.id===input.id);N(input.id?old&&old.revision===input.expectedRecordRevision&&old.formalApproval===null&&old.publication==='CLOSED'&&['DRAFT','UNDER_REVIEW','EVIDENCE_PENDING'].includes(old.documentaryState):input.expectedRecordRevision===0,'REVISION_CONFLICT');
  const source=get('institutionalDirectory/governance'),sourceRow=input.sourceRecordId?source?.records.find(r=>r.id===input.sourceRecordId&&r.visibility==='BOARD'):null;
  N(input.sourceRecordId===null||C.id(input.sourceRecordId)&&sourceRow,'RESOURCE_NOT_FOUND');
  if(old)N(old.sourceRecordId===input.sourceRecordId,'COMMAND_PAYLOAD_INVALID');
  N(Array.isArray(input.references)&&input.references.length<=10,'COMMAND_PAYLOAD_INVALID');
  const references=input.references.map(r=>{
   C.exact(r,['label','url','sha256','targetId','evidence']);N(text(r.label,160,3)&&hex(r.sha256)&&text(r.evidence,1500,10)&&/^https:\/\/drive\.google\.com\/file\/d\/[A-Za-z0-9_-]+\/(?:view)?$/.test(r.url),'COMMAND_PAYLOAD_INVALID');
   N(r.targetId===null||C.id(r.targetId),'COMMAND_PAYLOAD_INVALID');
   const target=r.targetId?get('institutionalDirectory/norms')?.records.find(t=>t.id===r.targetId&&t.visibility==='BOARD'):null;N(!r.targetId||target,'RESOURCE_NOT_FOUND');
   // Only an explicit original documentary relation establishes a proven link.
   const proven=!!target&&sourceRow?.fields.some(f=>['DOC_ID','DOC_ALVO','DOCUMENTO_RELACIONADO_ID'].includes(f.label)&&f.value.split(/[;,]/).map(x=>x.trim()).includes(r.targetId));
   return {...r,association:proven?'DOCUMENTED_SOURCE_LINK':r.targetId?'PENDING_VERIFICATION':'DOCUMENT_REFERENCE',targetHash:target?hash(target):null};
  });
  N(new Set(references.map(r=>r.url)).size===references.length,'COMMAND_PAYLOAD_INVALID');
  return {old,references,sourceHash:sourceRow?hash(sourceRow):null};
 }
 function preview(tx,o,input){const get=p=>tx.get(p).data(),now=o.clock(),{c,auth}=control(get,o.config,o.identity,now,true),v=state(get),p=prepare(get,input,v);return {beforeValue:p.old?p.old.title+' · revisão '+p.old.revision:'Novo rascunho administrativo',value:input.title+'\nData: '+(input.meetingDate||'Não informada')+'\n'+({DRAFT:'Rascunho',UNDER_REVIEW:'Em conferência documental',EVIDENCE_PENDING:'Aguardando evidências'})[input.documentaryState]+'\n'+input.body+'\nDocumentos de apoio: '+p.references.length,reason:input.reason,documentHash:hash({v,references:p.references,sourceHash:p.sourceHash}),epoch:c.epoch,controlRevision:c.revision,validUntil:Math.min(auth.validUntil,f4ServiceAvailability_().bound(c,now+60000)),publicUpdate:false,formalApproval:false};}
 function digest(config,uid,r){return hash({projectId:config.projectId,ownerUid:uid,commandType:TYPE,...r});}
 function status(tx,o,key){C.exact(key,['operationId','digest']);N(C.op(key.operationId)&&hex(key.digest),'COMMAND_PAYLOAD_INVALID');const get=p=>tx.get(p).data(),{c}=control(get,o.config,o.identity,o.clock(),false),r=get('institutionalGovernanceReceipts/'+key.operationId),owner=get('commandExecutorOwners/'+key.operationId);if(!r){N(!owner,'RECONCILIATION_REQUIRED');return {state:'NOT_ADMITTED',operationId:key.operationId};}N(r.ownerUid===o.identity.uid&&r.digest===key.digest&&r.projectId===o.config.projectId&&r.executor===c.executor&&r.state==='COMPLETED'&&owner?.ownerUid===r.ownerUid&&owner.digest===r.digest&&owner.executor===r.executor&&owner.commandType===TYPE,'ACCESS_REVOKED');return r.result;}
 function execute(tx,o,r){
  C.exact(r,['operationId','input','documentHash','epoch','controlRevision','validUntil','digest']);N(C.op(r.operationId)&&hex(r.digest)&&hex(r.documentHash),'COMMAND_PAYLOAD_INVALID');const {digest:given,...semantic}=r;N(given===digest(o.config,o.identity.uid,semantic),'OPERATION_CONFLICT');const get=p=>tx.get(p).data(),now=o.clock();if(get('institutionalGovernanceReceipts/'+r.operationId))return status(tx,o,{operationId:r.operationId,digest:r.digest});
  const q=preview(tx,o,r.input),{c,auth}=control(get,o.config,o.identity,now,true),v=state(get),p=prepare(get,r.input,v);N(r.documentHash===q.documentHash&&r.epoch===c.epoch&&r.controlRevision===c.revision,'REVISION_CONFLICT');N(Number.isSafeInteger(r.validUntil)&&now<r.validUntil&&r.validUntil<=now+60000,'ACCESS_REVOKED');
  for(const prefix of ['commandExecutorOwners','commandRequests','commandReceipts','commandLedger','institutionalGovernanceAudit'])N(!get(prefix+'/'+r.operationId),'OPERATION_CONFLICT');
  const bucket='institutionalGovernanceRate/'+C.hash(o.identity.uid).slice(0,32),oldRate=get(bucket),start=Math.floor(now/3600000)*3600000,count=oldRate?.windowStart===start?oldRate.count:0;N(Number.isInteger(count)&&count>=0&&count<c.rateLimit.maxAdmissions,'PILOT_LIMIT_REACHED');
  const hp='institutionalGovernanceHistory/current',history=get(hp)||{schemaVersion:'institutional-governance-history-v1',events:[]};N(Array.isArray(history.events)&&history.events.length<500,'SCOPE_TOO_LARGE');
  const id=r.input.id||'GOV_'+r.operationId,record={id,revision:(p.old?.revision||0)+1,title:r.input.title,meetingDate:r.input.meetingDate,body:r.input.body,documentaryState:r.input.documentaryState,sourceRecordId:r.input.sourceRecordId,sourceHash:p.sourceHash,references:p.references,formalApproval:null,publication:'CLOSED',effectiveFrom:null,createdAt:p.old?.createdAt||now,updatedAt:now,actorPersonId:auth.actor.actor.h.id,operationId:r.operationId};
  N(p.old||v.records.length<100,'SCOPE_TOO_LARGE');const next={...v,revision:v.revision+1,records:p.old?v.records.map(x=>x.id===id?record:x):[...v.records,record]};N(JSON.stringify(next).length<500000,'SCOPE_TOO_LARGE');
  const result={state:'COMPLETED',operationId:r.operationId,recordId:id,revision:next.revision,recordRevision:record.revision,publicUpdate:false,formalApproval:false},receipt={schemaVersion:'institutional-governance-receipt-v1',projectId:o.config.projectId,executor:c.executor,ownerUid:o.identity.uid,operationId:r.operationId,digest:r.digest,epoch:c.epoch,commandType:TYPE,state:'COMPLETED',result,at:now},event={operationId:r.operationId,recordId:id,actorPersonId:record.actorPersonId,revision:record.revision,reason:r.input.reason,at:now};
  const writes=[{path:PATH,data:next},{path:hp,data:{...history,events:[...history.events,event]}},{path:bucket,data:{windowStart:start,count:count+1}},{path:'commandExecutorOwners/'+r.operationId,data:{schemaVersion:'institutional-command-owner-v1',routeId:'INSTITUTIONAL_GOVERNANCE_DRAFT',operationId:r.operationId,commandType:TYPE,activityId:PATH,ownerUid:o.identity.uid,digest:r.digest,executor:c.executor,epoch:c.epoch,createdAt:F.Timestamp.fromMillis(now)}},{path:'institutionalGovernanceReceipts/'+r.operationId,data:receipt},{path:'institutionalGovernanceAudit/'+r.operationId,data:{...receipt,...event,before:p.old||null,after:record,policyRefs:auth.policies.map(x=>({id:x.id,revision:x.revision}))}}];
  const P=f4TransactionPlans_(),plan={kind:'INSTITUTIONAL_GOVERNANCE',aggregateId:'current',sourceWrites:writes.slice(0,2).map(w=>({...w,kind:'set'})),headWrites:[],claimWrites:[],projectionWrites:[],technicalWrites:writes.slice(2).map(w=>({...w,kind:'set'})),deletes:[]};return P.applyMutationPlan(tx,P.finalize(tx,plan,result,now,Math.min(auth.validUntil,r.validUntil)),o.clock());
 }
 function read(get,o){const auth=f4InstitutionalPolicies_(get,o.identity,o.clock(),get('institutionalDirectoryControl/current').policyIds,'INSTITUTIONAL_DIRECTORY_READ','institutionalDirectory/governance');N(auth.mask.includes('summary')&&auth.mask.includes('fields'),'AUTHORIZATION_DENIED');const v=state(get),h=get('institutionalGovernanceHistory/current');let canEdit=false;try{control(get,o.config,o.identity,o.clock(),true,true);canEdit=true;}catch(e){if(!['ACCESS_REVOKED','AUTHORIZATION_DENIED','PILOT_DISABLED','M13_AUTHZ_DENIED'].includes(e.code))throw e;}return {revision:v.revision,canEdit,records:v.records,history:h?.events||[],publicationEnabled:false,formalApprovalEnabled:false};}
 return {PATH,CONTROL,TYPE,CAP,hash,state,control,prepare,preview,digest,status,execute,read};
}
