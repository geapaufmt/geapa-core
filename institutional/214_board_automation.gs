// Internal consequence executor. It never impersonates a director, receives a
// vote, edits a proposal, anticipates a deadline or confers generic privileges.
function f4BoardAutomation_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),G=f4InstitutionalGovernance_(),B=f4BoardProcedures_(),X=f4BoardExtraordinary_();
 const CONTROL='institutionalBoardAutomationControl/current',ACTOR='SYSTEM_BOARD_COLLECTIVE_DECISION',clone=v=>F.decode(F.encode(v)),hash=v=>C.hash(F.encode(v));
 function control(get,o){
  const c=get(CONTROL);N(c?.schemaVersion==='board-automation-control-v1'&&c.enabled===true&&c.projectId===o.config.projectId&&c.executor===f4CommandExecutorId_()&&C.rev(c.revision)&&c.revision>0&&C.rev(c.epoch)&&c.epoch>0&&Number.isSafeInteger(c.startsAt)&&Number.isSafeInteger(c.expiresAt)&&c.startsAt<=o.clock()&&o.clock()<c.expiresAt,'ACCESS_REVOKED');
  N(Array.isArray(c.boardIds)&&c.boardIds.length>0&&c.boardIds.length<=10&&c.boardIds.every(C.id)&&Array.isArray(c.policyContinuity)&&c.policyContinuity.length<=50,'SCHEMA_INVALID');
  return c;
 }
 function authority(get,o,r){
  const c=control(get,o),g=get(G.CONTROL),p=r.procedure,now=o.clock();
  N(r.kind==='BOARD_PROCEDURE'&&p.type==='EXTRAORDINARY_EXTENSION'&&c.boardIds.includes(p.boardId)&&g?.enabled===true&&f4ServiceAvailability_().live(g,now)&&g.projectId===o.config.projectId&&g.executor===c.executor&&g.boardProcedures?.enabled===true&&g.boardProcedures.types.includes(p.type)&&g.boardProcedures.boardIds.includes(p.boardId)&&g.boardProcedures.targetPersonIds.includes(p.target.personId),'SCOPE_INVALID');
  N(g.publicationEnabled===false&&g.formalApprovalEnabled===false&&g.appCheck?.mode==='REQUIRED','ACCESS_REVOKED');
  N(hash(B.content(p))===p.contentHash,'CONTENT_VERSION_CONFLICT');
  const current=B.composition(get,o,p.boardId);N(hash(current)===hash(p.electorate),'COMPOSITION_CHANGED');
  N(current.people.some(x=>x.personId===p.proposer&&x.functions.some(f=>f.roleKey==='PRESIDENTE')),'PRESIDENT_REQUIRED');
  // A typed procedure is insufficient without its original authenticated journal.
  const history=get('institutionalGovernanceHistory/current'),events=(history?.events||[]).filter(e=>e.recordId===r.id);N(events.length<=50,'SCOPE_TOO_LARGE');
  // Legacy records lack openingOperationId. Resolve the original immutable
  // audit without editing the proposal or fabricating a new opening event.
  const openings=p.openingOperationId?events.filter(e=>e.operationId===p.openingOperationId):events.filter(e=>{const a=get('institutionalGovernanceAudit/'+e.operationId);return a?.action==='BP_OPEN'&&a.after?.procedure?.openedAt===p.openedAt;});N(openings.length===1,'RECONCILIATION_REQUIRED');const opening=openings[0];
  const verified=(id,a,q)=>{const owner=get('commandExecutorOwners/'+id);N(q?.state==='COMPLETED'&&q.projectId===o.config.projectId&&q.executor===c.executor&&q.operationId===id&&q.ownerUid===a?.ownerUid&&q.digest===a?.digest&&owner?.ownerUid===q.ownerUid&&owner.executor===q.executor&&owner.digest===q.digest,'RECONCILIATION_REQUIRED');};
  const op=get('institutionalGovernanceAudit/'+opening.operationId),receipt=get('institutionalGovernanceReceipts/'+opening.operationId);
  verified(opening.operationId,op,receipt);
  N(op?.action==='BP_OPEN'&&op.actorPersonId===p.proposer&&op.after?.procedure?.contentHash===p.contentHash&&hash(op.after.procedure.electorate)===hash(p.electorate)&&op.after.procedure.deadline===p.deadline&&op.after.procedure.allowEarlyClose===p.allowEarlyClose&&receipt?.state==='COMPLETED'&&receipt.ownerUid===op.ownerUid&&receipt.digest===op.digest,'RECONCILIATION_REQUIRED');
  for(const v of p.votes){const a=get('institutionalGovernanceAudit/'+v.operationId),q=get('institutionalGovernanceReceipts/'+v.operationId);verified(v.operationId,a,q);N(a?.action==='BP_VOTE'&&a.actorPersonId===v.personId&&a.after?.procedure?.votes?.some(x=>hash(x)===hash(v)),'RECONCILIATION_REQUIRED');}
  const a=p.consent&&get('institutionalGovernanceAudit/'+p.consent.operationId),q=p.consent&&get('institutionalGovernanceReceipts/'+p.consent.operationId);N(p.consent,'DEFENSE_OR_CONSENT_PENDING');verified(p.consent.operationId,a,q);N(a?.action==='BP_CONSENT'&&a.actorPersonId===p.target.personId&&hash(a.after.procedure.consent)===hash(p.consent),'RECONCILIATION_REQUIRED');
  return {control:c,governance:g,validUntil:Math.min(c.expiresAt,f4ServiceAvailability_().bound(g,now+60000)),actor:ACTOR};
 }
 function policyChanges(get,o,r,effect,c){
  const out=[],p=r.procedure,end=C.civilStart(C.nextDay(p.target.endDate));
  for(const pin of c.policyContinuity.filter(x=>x.functionId===p.target.functionId)){
   C.exact(pin,['functionId','policyId','revision','sha256']);N(C.id(pin.policyId)&&C.rev(pin.revision)&&/^[a-f0-9]{64}$/.test(pin.sha256),'SCHEMA_INVALID');
   const path='memberAuthzPolicies/'+pin.policyId,old=get(path),fn=get('memberFunctions/'+pin.functionId),role=fn&&get('memberRoleCatalog/'+fn.idCargo);f4MemberAuthzContracts_().policy(old);
   N(old.active===true&&old.revision===pin.revision&&hash(old)===pin.sha256&&old.proofKind==='ROLE'&&old.actorSelector.kind==='ROLE'&&old.actorSelector.roleKey===role?.roleKey&&old.targets.allowedPersons.length===1&&old.targets.allowedPersons[0]===p.target.personId,'AUTHORIZATION_DENIED');
   N(C.ms(old.expiresAt)>o.clock(),'ACCESS_REVOKED');
   // No new capabilities, fields, resources, subjects or proof kinds. Only the
   // explicitly pinned nominal validity can follow a confirmed occupation.
   if(C.ms(old.expiresAt)<end){const next={...clone(old),revision:old.revision+1,expiresAt:F.Timestamp.fromMillis(end)};f4MemberAuthzContracts_().policy(next);out.push({path,kind:'set',data:next});}
  }
  N(new Set(out.map(x=>x.path)).size===out.length,'SCHEMA_INVALID');return out;
 }
 function key(r,phase){return C.hash({kind:'BOARD_AUTOMATIC_'+phase,id:r.id,contentHash:r.procedure.contentHash}).slice(0,32);}
 function execute(tx,o,id,phase){
  N(['SETTLE','ENACT'].includes(phase)&&C.id(id),'COMMAND_PAYLOAD_INVALID');const get=p=>tx.get(p).data(),c=control(get,o),v=G.state(get),r=v.records.find(x=>x.id===id);N(r?.kind==='BOARD_PROCEDURE'&&r.procedure.type==='EXTRAORDINARY_EXTENSION','RESOURCE_NOT_FOUND');
  const operationId=key(r,phase),rp='boardAutomaticReceipts/'+operationId,prior=get(rp);if(prior){N(prior.recordId===id&&prior.contentHash===r.procedure.contentHash&&prior.state==='COMPLETED','OPERATION_CONFLICT');return prior.result;}
  const p=r.procedure,now=o.clock();if(phase==='SETTLE'&&p.state!=='OPEN'||phase==='ENACT'&&p.enactment?.state==='ENACTED')return {state:'UNCHANGED',recordId:id};
  const auth=authority(get,o,r),decision=X.decision(p,now);if(phase==='SETTLE'&&decision.state==='OPEN')return {state:'WAITING',recordId:id,dueAt:p.deadline};
  const next=clone(r);let member=null,policies=[];
  if(phase==='SETTLE'){next.procedure.state=decision.state;next.procedure.closedAt=now;next.procedure.closedBy=ACTOR;next.documentaryState=decision.state;}
  else{
   N(p.state==='APPROVED'&&decision.state==='APPROVED','PROCEDURE_NOT_APPROVED');
   member=f4MemberMutationPlan_().planBoardAutomatic(tx,{...o,operationId,validUntil:auth.validUntil},r);
   policies=policyChanges(get,o,r,member,c);
   next.procedure.enactment={state:'ENACTED',mode:'AUTOMATIC_VALIDATED_DECISION',at:now,actorPersonId:ACTOR,operationId,functionId:member.functionId,evidence:member.evidence,tenure:member.summary,documentaryAuthorization:member.documentaryAuthorization,responsibleFunctionId:null};
  }
  next.revision=r.revision+1;next.updatedAt=now;next.actorPersonId=ACTOR;next.operationId=operationId;
  if(auth.governance.boardProcedures.pdfStorage){N(typeof o.renderProcedureDocument==='function','RECONCILIATION_REQUIRED');next.procedure.documentHistory=[...(p.documentHistory||[]),...(p.document?[{recordRevision:r.revision,descriptor:p.document}]:[])];N(next.procedure.documentHistory.length<=20,'SCOPE_TOO_LARGE');next.procedure.document=o.renderProcedureDocument(next);}
  const hp='institutionalGovernanceHistory/current',history=get(hp);N(history?.events?.length<500,'SCOPE_TOO_LARGE');const event={operationId,recordId:id,actorPersonId:ACTOR,revision:next.revision,reason:'Efeito automático de decisão definitiva validada pelo Core',action:'BP_AUTOMATIC_'+phase,at:now};
  const result={state:'COMPLETED',operationId,recordId:id,recordRevision:next.revision,procedureState:next.procedure.state,enactment:next.procedure.enactment?.state||'NOT_ENACTED'};
  const receipt={schemaVersion:'board-automatic-receipt-v1',state:'COMPLETED',recordId:id,contentHash:p.contentHash,operationId,phase,actorKind:'CANONICAL_SYSTEM_CONSEQUENCE',executor:c.executor,projectId:c.projectId,epoch:c.epoch,at:now,result};
  for(const path of ['commandExecutorOwners/'+operationId,'boardAutomaticAudit/'+operationId,'boardAutomaticJournal/'+operationId])N(!get(path),'OPERATION_CONFLICT');
  const plan={kind:'INSTITUTIONAL_GOVERNANCE',aggregateId:'current',sourceWrites:[{path:G.PATH,kind:'set',data:{...v,revision:v.revision+1,records:v.records.map(x=>x.id===id?next:x)}},{path:hp,kind:'set',data:{...history,events:[...history.events,event]}},...policies],headWrites:[],claimWrites:[],projectionWrites:[],technicalWrites:[{path:rp,kind:'set',data:receipt},{path:'commandExecutorOwners/'+operationId,kind:'set',data:{schemaVersion:'institutional-command-owner-v1',routeId:'BOARD_AUTOMATIC',commandType:'board.extraordinary.consequence',activityId:G.PATH,operationId,ownerUid:ACTOR,digest:hash(receipt),executor:c.executor,epoch:c.epoch,createdAt:F.Timestamp.fromMillis(now)}},{path:'boardAutomaticAudit/'+operationId,kind:'set',data:{...receipt,before:r,after:next,policyChanges:policies.map(w=>({path:w.path,before:get(w.path),after:w.data})),authorizationRevision:c.revision}},{path:'boardAutomaticJournal/'+operationId,kind:'set',data:{...receipt,expectedRecordRevision:r.revision,expectedAggregateRevision:v.revision}}],deletes:[]};
  if(next.procedure.document&&auth.governance.boardProcedures.pdfStorage){const path='boardAutomaticDocumentOutbox/'+operationId;N(!get(path),'OPERATION_CONFLICT');plan.technicalWrites.push({path,kind:'set',data:{schemaVersion:'board-automatic-document-v1',revision:1,state:'PENDING',operationId,recordId:id,recordRevision:next.revision,descriptor:next.procedure.document,createdAt:now}});}
  if(member)for(const k of ['sourceWrites','headWrites','claimWrites','projectionWrites','technicalWrites','deletes'])plan[k].push(...member.plan[k]);
  const P=f4TransactionPlans_(),finished=P.finalize(tx,plan,result,now,auth.validUntil);N(typeof o.sealAutomaticPlan==='function','RECONCILIATION_REQUIRED');o.sealAutomaticPlan(finished);return P.applyMutationPlan(tx,finished,o.clock());
 }
 // Storage is an external effect. Only the committed audit snapshot can be
 // delivered; aborted/concurrent transaction snapshots are never uploaded.
 function document(tx,o,operationId,proof){
  N(C.op(operationId),'COMMAND_PAYLOAD_INVALID');const get=p=>tx.get(p).data(),c=control(get,o),path='boardAutomaticDocumentOutbox/'+operationId,j=get(path);
  if(!j)return {state:'NONE'};
  const a=get('boardAutomaticAudit/'+operationId),r=get('boardAutomaticReceipts/'+operationId);
  N(j.schemaVersion==='board-automatic-document-v1'&&j.operationId===operationId&&['PENDING','COMPLETED'].includes(j.state)&&r?.state==='COMPLETED'&&r.executor===c.executor&&r.projectId===c.projectId&&a?.after?.id===j.recordId&&a.after.revision===j.recordRevision&&hash(a.after.procedure.document)===hash(j.descriptor)&&a.operationId===operationId&&a.contentHash===r.contentHash,'RECONCILIATION_REQUIRED');
  if(!proof)return {state:j.state,record:a.after,descriptor:j.descriptor};
  C.exact(proof,['sha256','versionId']);N(proof.sha256===j.descriptor.sha256&&typeof proof.versionId==='string'&&/^[0-9]+$/.test(proof.versionId),'RECONCILIATION_REQUIRED');
  if(j.state==='COMPLETED'){N(j.versionId===proof.versionId,'RECONCILIATION_REQUIRED');return {state:'COMPLETED'};}
  const now=o.clock(),result={state:'COMPLETED'},plan={kind:'INSTITUTIONAL_GOVERNANCE',aggregateId:'current',sourceWrites:[],headWrites:[],claimWrites:[],projectionWrites:[],technicalWrites:[{path,kind:'set',data:{...j,revision:j.revision+1,state:'COMPLETED',versionId:proof.versionId,deliveredAt:now}}],deletes:[]},P=f4TransactionPlans_(),finished=P.finalize(tx,plan,result,now,c.expiresAt);
  o.sealAutomaticPlan(finished);return P.applyMutationPlan(tx,finished,o.clock());
 }
 return {CONTROL,ACTOR,control,authority,policyChanges,key,execute,document};
}
