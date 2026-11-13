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
 function sealed(p){
  N(p?.type==='EXTRAORDINARY_EXTENSION'&&p.rule==='ABSOLUTE_MAJORITY','ART33_SPECIFIC_EXCEPTION_REQUIRED');
  declaration(p.exception);
  N(f4BoardProcedures_().hash(f4BoardProcedures_().content(p))===p.contentHash,'CONTENT_VERSION_CONFLICT');
  N(p.consent?.accepted===true&&p.consent.personId===p.target.personId&&p.consent.contentHash===p.contentHash,'DEFENSE_OR_CONSENT_PENDING');
  return {contentHash:p.contentHash,articles:[...articles],individualGroundsHash:hash(p.exception.individualGrounds)};
 }
 function decision(p,now){
  sealed(p);
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
  const proof=sealed(p),d=decision(p,now);N(p.state==='APPROVED'&&d.state==='APPROVED','PROCEDURE_NOT_APPROVED');
  return {...proof,interpretationSourceHash:q.interpretationSourceHash,required:d.required,eligible:d.eligible,yes:d.yes};
 }
 // Safe inspection preserves the old vote, deadline, composition and proposal.
 // It never adds an exception declaration to an already sealed procedure.
 function inspect(p,fn,now){
  const blocks=[];let tally=null;
  try{tally=decision(p,now);}catch(e){blocks.push(e.code||'ART33_SPECIFIC_EXCEPTION_REQUIRED');}
  if(fn){const end=fn.endDateReal||fn.endDatePlanned;if(Number.isSafeInteger(p.deadline)&&p.deadline>=C.civilStart(C.nextDay(end))&&!p.allowEarlyClose)blocks.push('DECISION_AFTER_FUNCTION_END');if(now>=C.civilStart(C.nextDay(end)))blocks.push('TARGET_FUNCTION_EXPIRED');}
  return {state:p.state,contentHash:p.contentHash,deadline:p.deadline,allowEarlyClose:p.allowEarlyClose,tally,blocks:[...new Set(blocks)],mayEnact:blocks.length===0&&p.state==='APPROVED'&&tally?.state==='APPROVED'};
 }
 return {articles,declaration,sealed,decision,evidence,inspect};
}
