// Explicit individual exceptions. Documentary validation is not a vote, a
// historical approval or authority to extend an expired occupation.
function f4BoardExtraordinary_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),hash=v=>C.hash(F.encode(v));
 const articles=['ART_33_PAR_5','ART_47'];
 function declaration(value){
  N(value&&typeof value==='object','ART33_SPECIFIC_EXCEPTION_REQUIRED');
  C.exact(value,['kind','articles','individualGrounds']);
  N(value.kind==='INDIVIDUAL_TENURE_EXCEPTION'&&Array.isArray(value.articles)&&value.articles.length===2&&articles.every(a=>value.articles.includes(a)),'ART33_SPECIFIC_EXCEPTION_REQUIRED');
  N(typeof value.individualGrounds==='string'&&value.individualGrounds===value.individualGrounds.trim()&&value.individualGrounds.length>=30&&value.individualGrounds.length<=4000&&!/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(value.individualGrounds),'ART33_SPECIFIC_EXCEPTION_REQUIRED');
  return value;
 }
 // Documentary recognition is independent of the immutable voted content.
 // A control pin authorizes exactly one existing procedure, not all old votes.
 function candidate(c,r){return !!r.procedure?.exception||!!c?.legacyRecognitions?.some(x=>x.recordId===r.id&&x.contentHash===r.procedure?.contentHash);}
 function recognition(get,p,now){
  N(typeof get==='function'&&Number.isSafeInteger(now),'ART33_SPECIFIC_EXCEPTION_REQUIRED');
  const c=get('institutionalBoardAutomationControl/current'),E=f4ExecutionEnvironment_();
  N(c?.enabled===true&&c.projectId===E.project&&E.executors.includes(c.executor)&&c.startsAt<=now&&now<c.expiresAt,'ACCESS_REVOKED');
  const pins=c.legacyRecognitions;N(Array.isArray(pins)&&pins.length>0&&pins.length<=10,'ART33_SPECIFIC_EXCEPTION_REQUIRED');
  const matches=pins.filter(x=>x.contentHash===p.contentHash);N(matches.length===1,'ART33_SPECIFIC_EXCEPTION_REQUIRED');const pin=matches[0];
  C.exact(pin,['recordId','contentHash','recognitionId','sha256']);N(C.id(pin.recordId)&&C.id(pin.recognitionId)&&/^[a-f0-9]{64}$/.test(pin.sha256),'SCHEMA_INVALID');
  const r=get('institutionalGovernanceDrafts/current')?.records.find(x=>x.id===pin.recordId),d=get('institutionalBoardRecognitions/'+pin.recognitionId);
  N(r?.kind==='BOARD_PROCEDURE'&&r.procedure.contentHash===p.contentHash&&d?.schemaVersion==='board-legacy-exception-recognition-v1'&&hash(d)===pin.sha256&&d.projectId===c.projectId&&d.executor===c.executor&&d.recordId===r.id&&d.contentHash===p.contentHash,'RECONCILIATION_REQUIRED');
  N(d.meaning==='EXISTING_INDIVIDUAL_EXTRAORDINARY_PURPOSE'&&d.authority==='EXPLICIT_INSTITUTIONAL_RECOGNITION'&&Number.isSafeInteger(d.recognizedAt)&&d.recognizedAt<=now&&now<d.expiresAt,'ART33_SPECIFIC_EXCEPTION_REQUIRED');
  N(p.type==='EXTRAORDINARY_EXTENSION'&&p.rule==='ABSOLUTE_MAJORITY'&&d.targetHash===hash(p.target)&&d.consentHash===hash(p.consent)&&d.electorateHash===hash(p.electorate)&&d.originalGroundsHash===hash({basis:p.basis,reason:p.reason})&&typeof p.basis==='string'&&p.basis.trim().length>=10&&typeof p.reason==='string'&&p.reason.trim().length>=10,'CONTENT_VERSION_CONFLICT');
  const ref=d.sourceMinute,m=ref&&get('institutionalGovernanceDrafts/current')?.records.find(x=>x.id===ref.id);
  N(m?.revision===ref.revision&&m.minutes?.state==='ARCHIVED'&&m.minutes.archive?.contentHash===m.minutes.submitted?.contentHash&&m.minutes.archive.contentHash===ref.contentHash&&hash(m)===ref.sha256&&Number.isSafeInteger(m.minutes.archive.at)&&m.minutes.archive.at<=p.openedAt,'ELECTRONIC_FORMALIZATION_REQUIRED');
  return {id:pin.recognitionId,sha256:pin.sha256,recordId:r.id,document:d,sourceHash:ref.sha256};
 }
 function deadlineCorrection(get,p,opened,now){
  if(opened.deadline===p.deadline)return null;
  const recognized=recognition(get,p,now),q=recognized.document.deadlineCorrection;
  N(q&&C.op(q.operationId)&&q.beforeDeadline===opened.deadline&&q.afterDeadline===p.deadline,'RECONCILIATION_REQUIRED');
  const receipt=get('institutionalGovernanceReceipts/'+q.operationId),audit=get('institutionalGovernanceAudit/'+q.operationId),journal=get('institutionalGovernanceJournal/'+q.operationId),owner=get('commandExecutorOwners/'+q.operationId);
  N(receipt&&audit&&journal&&owner&&hash(receipt)===q.receiptHash&&hash(audit)===q.auditHash&&hash(journal)===q.journalHash&&hash(owner)===q.ownerHash,'RECONCILIATION_REQUIRED');
  N(receipt.schemaVersion==='institutional-governance-administrative-receipt-v1'&&receipt.executor==='AUTHORIZED_BOUNDED_MAINTENANCE'&&receipt.state==='COMPLETED'&&receipt.projectId===f4ExecutionEnvironment_().project&&receipt.recordId===recognized.recordId&&receipt.operationId===q.operationId&&receipt.manifestSha256===q.manifestSha256&&receipt.authorization==='USER_EXPLICIT_DEADLINE_ONLY_NO_ADDITIONAL_VOTES'&&receipt.beforeDeadline===opened.deadline&&receipt.afterDeadline===p.deadline&&receipt.at<=now&&receipt.at<p.deadline&&receipt.formalApproval===false&&receipt.enactment===false,'RECONCILIATION_REQUIRED');
  N(audit.operationId===receipt.operationId&&audit.manifestSha256===receipt.manifestSha256&&owner.manifestSha256===receipt.manifestSha256&&owner.owner===receipt.owner&&owner.projectId===receipt.projectId&&journal.phase==='COMMITTED'&&journal.operationId===receipt.operationId&&journal.manifestSha256===receipt.manifestSha256,'RECONCILIATION_REQUIRED');
  const before=audit.before,after=audit.after;N(before?.id===recognized.recordId&&after?.id===before.id&&after.revision===before.revision+1&&before.procedure.deadline===opened.deadline&&after.procedure.deadline===p.deadline&&before.procedure.state==='OPEN'&&after.procedure.state==='OPEN','RECONCILIATION_REQUIRED');
  const old=JSON.parse(JSON.stringify(before.procedure)),next=JSON.parse(JSON.stringify(after.procedure));delete old.deadline;delete next.deadline;
  N(hash(old)===hash(next)&&after.procedure.contentHash===p.contentHash&&hash(after.procedure.consent)===hash(p.consent)&&after.procedure.votes.every(v=>p.votes.some(w=>hash(v)===hash(w)))&&p.allowEarlyClose===opened.allowEarlyClose,'CONTENT_VERSION_CONFLICT');
  return {operationId:q.operationId,manifestSha256:q.manifestSha256,beforeDeadline:q.beforeDeadline,afterDeadline:q.afterDeadline};
 }
 function sealed(p,get,now){
  N(p?.type==='EXTRAORDINARY_EXTENSION'&&p.rule==='ABSOLUTE_MAJORITY','ART33_SPECIFIC_EXCEPTION_REQUIRED');
  const legacy=p.exception?null:recognition(get,p,now);if(p.exception)declaration(p.exception);
  N(f4BoardProcedures_().hash(f4BoardProcedures_().content(p))===p.contentHash,'CONTENT_VERSION_CONFLICT');
  N(p.consent?.accepted===true&&p.consent.personId===p.target.personId&&p.consent.contentHash===p.contentHash,'DEFENSE_OR_CONSENT_PENDING');
  return {contentHash:p.contentHash,articles:[...articles],individualGroundsHash:legacy?legacy.document.originalGroundsHash:hash(p.exception.individualGrounds),...(legacy?{recognitionId:legacy.id,recognitionHash:legacy.sha256,interpretationSourceHash:legacy.sourceHash}:{})};
 }
 function decision(p,now,get){
  sealed(p,get,now);
  N(Number.isSafeInteger(p.openedAt)&&Number.isSafeInteger(p.deadline)&&p.openedAt<p.deadline&&typeof p.allowEarlyClose==='boolean','PROCEDURE_NOT_APPROVED');
  const people=p.electorate?.people;
  N(Array.isArray(people)&&people.length>0&&new Set(people.map(x=>x.personId)).size===people.length&&people.every(x=>C.id(x.personId)),'COMPOSITION_CHANGED');
  N(Array.isArray(p.votes)&&p.votes.length<=people.length&&new Set(p.votes.map(x=>x.personId)).size===p.votes.length,'CONTENT_VERSION_CONFLICT');
  for(const v of p.votes)N(people.some(x=>x.personId===v.personId)&&['YES','NO','ABSTAIN'].includes(v.choice)&&v.contentHash===p.contentHash&&v.contentVersion===p.contentVersion&&Number.isSafeInteger(v.at)&&v.at>=p.openedAt&&v.at<p.deadline&&v.at<=now&&C.op(v.operationId),'CONTENT_VERSION_CONFLICT');
  const result=f4BoardProcedures_().outcome({...p,state:'OPEN'},now);
  N(['OPEN','APPROVED','REJECTED','INSUFFICIENT_MANIFESTATIONS'].includes(result),'PROCEDURE_NOT_APPROVED');
  return {state:result,required:Math.floor(people.length/2)+1,eligible:people.length,yes:p.votes.filter(v=>v.choice==='YES').length};
 }
 function evidence(get,b,p,now){
  const q=b.individualException;
  C.exact(q,['mode','articles','approvalRule','automaticEffect','interpretationSourceHash']);
  N(q.mode==='EXPRESS_SEALED_DELIBERATION'&&q.approvalRule==='ABSOLUTE_MAJORITY'&&q.automaticEffect===true&&Array.isArray(q.articles)&&q.articles.length===2&&articles.every(a=>q.articles.includes(a))&&/^[a-f0-9]{64}$/.test(q.interpretationSourceHash),'M13_ACT_INVALID');
  N(b.procedureRequirements?.EXTRAORDINARY_EXTENSION?.conventionalMinuteHash===q.interpretationSourceHash,'M13_ACT_INVALID');
  const proof=sealed(p,get,now),d=decision(p,now,get);N(!proof.recognitionId||proof.interpretationSourceHash===q.interpretationSourceHash,'M13_ACT_INVALID');N(p.state==='APPROVED'&&d.state==='APPROVED','PROCEDURE_NOT_APPROVED');
  return {...proof,interpretationSourceHash:q.interpretationSourceHash,required:d.required,eligible:d.eligible,yes:d.yes};
 }
 // Safe inspection preserves the old vote, deadline, composition and proposal.
 // It never adds an exception declaration to an already sealed procedure.
 function inspect(p,fn,now,get){
  const blocks=[];let tally=null;
  try{tally=decision(p,now,get);}catch(e){blocks.push(e.code||'ART33_SPECIFIC_EXCEPTION_REQUIRED');}
  if(fn){const end=fn.endDateReal||fn.endDatePlanned;if(Number.isSafeInteger(p.deadline)&&p.deadline>=C.civilStart(C.nextDay(end))&&!p.allowEarlyClose)blocks.push('DECISION_AFTER_FUNCTION_END');if(now>=C.civilStart(C.nextDay(end)))blocks.push('TARGET_FUNCTION_EXPIRED');}
  return {state:p.state,contentHash:p.contentHash,deadline:p.deadline,allowEarlyClose:p.allowEarlyClose,tally,blocks:[...new Set(blocks)],mayEnact:blocks.length===0&&p.state==='APPROVED'&&tally?.state==='APPROVED'};
 }
 return {articles,declaration,candidate,recognition,deadlineCorrection,sealed,decision,evidence,inspect};
}
