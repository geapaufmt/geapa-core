// Voluntary suspension shares the request/index and Members transaction family.
// NC01 art. 15 + explicit decision 2026-10-13: effective members only.
// A suspension adds an event to the existing link; it never recreates a link.
function f4MemberVoluntarySuspension_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),T=F.Timestamp.fromMillis,R=f4MemberVoluntaryRules_();
 const modality='SUSPENSAO_TEMPORARIA',terminal=['CONCLUIDO','INDEFERIDO','CANCELADO_PELO_MEMBRO','CANCELADO_PELA_DIRETORIA'];
 function text(v,max,required=false){N(typeof v==='string'&&v.length<=max&&(!required||v.trim()),'COMMAND_PAYLOAD_INVALID');return v.trim();}
 function policy(v){
  C.exact(v,['schemaVersion','revision','enabled','minimumDays','maximumPerMembership','approvalRef','source']);
  N(v.schemaVersion==='member-voluntary-suspension-policy-v1'&&v.enabled===true&&C.rev(v.revision)&&v.revision>0&&v.minimumDays===30&&v.maximumPerMembership===1&&C.id(v.approvalRef)&&v.source?.normId==='NC01_2025_ART15','ACCESS_REVOKED');return v;
 }
 function validate(v){
  C.exact(v,['schemaVersion','id','revision','personId','linkId','modality','status','reason','semesterId','semesterEndSnapshot','policyRevision','createdAt','updatedAt','createdByUid','history','decision','submission','startDate','endDate','suspensionEventId','returnEventId']);
  N(v.schemaVersion==='member-voluntary-suspension-v1'&&C.op(v.id)&&C.rev(v.revision)&&v.revision>0&&C.id(v.personId)&&C.id(v.linkId)&&C.id(v.createdByUid)&&v.modality===modality,'SCHEMA_INVALID');
  N(['RECEBIDO','EM_ANALISE','AGUARDANDO_COMPLEMENTO','HOMOLOGADO_AGUARDANDO_INICIO','SUSPENSAO_ATIVA',...terminal].includes(v.status)&&C.id(v.semesterId)&&C.rev(v.policyRevision)&&v.policyRevision>0,'SCHEMA_INVALID');
  C.civilStart(v.startDate);C.civilStart(v.endDate);C.civilStart(v.semesterEndSnapshot);N(v.startDate<=v.endDate&&v.endDate<=v.semesterEndSnapshot,'SCHEMA_INVALID');
  text(v.reason,4000,true);N(Array.isArray(v.history)&&v.history.length>0&&v.history.length<=100&&Number.isSafeInteger(C.ms(v.createdAt))&&C.ms(v.updatedAt)>=C.ms(v.createdAt),'SCHEMA_INVALID');
  C.exact(v.submission,['reasonCategory','documentReference','acknowledged','version']);text(v.submission.reasonCategory,120,true);text(v.submission.documentReference,500);N(v.submission.acknowledged===true&&v.submission.version==='PORTAL_VINCULO_V2','SCHEMA_INVALID');
  for(const k of ['suspensionEventId','returnEventId'])N(v[k]===null||C.id(v[k]),'SCHEMA_INVALID');
  if(v.decision!==null){C.exact(v.decision,['action','actorUid','at','decisionReference','reason','exception']);N(['APPROVE','DENY'].includes(v.decision.action)&&C.id(v.decision.actorUid)&&Number.isSafeInteger(C.ms(v.decision.at))&&typeof v.decision.exception==='boolean','SCHEMA_INVALID');text(v.decision.decisionReference,500,true);text(v.decision.reason,4000,true);}
  return v;
 }
 function interval(p,semester,today,link){
  C.civilStart(p.startDate);C.civilStart(p.endDate);N(p.startDate>=today&&p.startDate<=p.endDate,'DOMAIN_STATE_INVALID');
  const days=Math.round((C.civilStart(p.endDate)-C.civilStart(p.startDate))/86400000)+1;N(days>=30,'DOMAIN_STATE_INVALID');
  N(semester&&semester.id===p.semesterId&&semester.civilStart<=today&&today<=semester.civilEnd&&p.endDate<=semester.civilEnd,'DOMAIN_STATE_INVALID');
  N(link.startDate<=p.startDate&&(link.endDate===null||p.endDate<=link.endDate),'DOMAIN_STATE_INVALID');
 }
 function entry(action,actorUid,now,before,after,note){return {action,actorUid,at:T(now),before,after,note};}
 function create(x){
  const {operationId,actorUid,personId,link,state,today,now,semester,priorIndex,payload}=x,p=policy(x.policy),W=f4MemberVoluntaryWorkflow_();
  C.exact(payload,['modality','reason','semesterId','startDate','endDate','submission']);N(payload.modality===modality&&link.type==='MEMBRO_EFETIVO','ACCESS_REVOKED');R.eligible(link,personId,state,today);interval(payload,semester,today,link);
  const i=W.index(priorIndex,personId);N(i.activeRequestId===null&&i.requestIds.length<100&&!i.requestIds.includes(operationId),'OPERATION_CONFLICT');
  const request=validate({schemaVersion:'member-voluntary-suspension-v1',id:operationId,revision:1,personId,linkId:link.id,modality,status:'RECEBIDO',reason:text(payload.reason,4000,true),semesterId:semester.id,semesterEndSnapshot:semester.civilEnd,policyRevision:p.revision,createdAt:T(now),updatedAt:T(now),createdByUid:actorUid,history:[entry('SUBMIT',actorUid,now,null,'RECEBIDO','')],decision:null,submission:payload.submission,startDate:payload.startDate,endDate:payload.endDate,suspensionEventId:null,returnEventId:null});
  return {request,index:{...i,revision:i.revision+1,activeRequestId:operationId,requestIds:i.requestIds.concat(operationId)},effect:null};
 }
 function review(x){
  const v=F.decode(F.encode(validate(x.request))),{payload,actorUid,actorPersonId,administrative,today,now,expectedRevision,semester,link,priorSuspensions}=x,p=policy(x.policy);
  N(v.revision===expectedRevision,'REVISION_CONFLICT');N(!terminal.includes(v.status),'DOMAIN_STATE_INVALID');
  N(p.revision===v.policyRevision,'RECONCILIATION_REQUIRED');
  C.exact(payload,['action','reason','minutesReference','confirmed',...(Object.hasOwn(payload,'graveException')?['graveException']:[]),...(Object.hasOwn(payload,'requestKind')?['requestKind']:[])]);text(payload.reason,4000,true);text(payload.minutesReference,500);N(typeof payload.confirmed==='boolean'&&(!payload.requestKind||payload.requestKind===modality),'COMMAND_PAYLOAD_INVALID');
  const self=['CANCEL_SELF','COMPLEMENT'].includes(payload.action);N(self?actorPersonId===v.personId:administrative===true,'ACCESS_REVOKED');
  let status,effect=null;const note=payload.reason.trim();
  if(['START','REQUEST_COMPLEMENT','COMPLEMENT','CANCEL_SELF','CANCEL_ADMIN','DENY'].includes(payload.action)){
   N(['RECEBIDO','EM_ANALISE','AGUARDANDO_COMPLEMENTO'].includes(v.status),'DOMAIN_STATE_INVALID');
   if(payload.action==='COMPLEMENT')N(v.status==='AGUARDANDO_COMPLEMENTO','DOMAIN_STATE_INVALID');
   status={START:'EM_ANALISE',REQUEST_COMPLEMENT:'AGUARDANDO_COMPLEMENTO',COMPLEMENT:'EM_ANALISE',CANCEL_SELF:'CANCELADO_PELO_MEMBRO',CANCEL_ADMIN:'CANCELADO_PELA_DIRETORIA',DENY:'INDEFERIDO'}[payload.action];
  }else if(payload.action==='APPROVE'){
   N(['RECEBIDO','EM_ANALISE'].includes(v.status)&&link?.id===v.linkId&&link.idPessoa===v.personId&&link.type==='MEMBRO_EFETIVO','DOMAIN_STATE_INVALID');
   interval(v,semester,today,link);N(Array.isArray(priorSuspensions),'RECONCILIATION_REQUIRED');
   const exception=payload.graveException===true;N(priorSuspensions.length<p.maximumPerMembership||exception,'DOMAIN_STATE_INVALID');
   N(payload.confirmed===true,'COMMAND_PAYLOAD_INVALID');
   // A formal human reference is always retained. Repetition additionally
   // records the exceptional decision; it is never automatically waived.
   v.decision={action:'APPROVE',actorUid,at:T(now),decisionReference:payload.minutesReference.trim()||'COMMAND_'+x.operationId,reason:note,exception};
   status=v.startDate===today?'SUSPENSAO_ATIVA':'HOMOLOGADO_AGUARDANDO_INICIO';
   v.suspensionEventId='EVT_'+x.operationId;effect={type:'SUSPENSAO',effectiveDate:v.startDate,endDate:v.endDate,id:v.suspensionEventId};
  }else if(payload.action==='START_SUSPENSION'){
   N(v.status==='HOMOLOGADO_AGUARDANDO_INICIO'&&today>=v.startDate&&today<=v.endDate&&v.suspensionEventId,'DOMAIN_STATE_INVALID');status='SUSPENSAO_ATIVA';
  }else if(payload.action==='RETURN'){
   N(['HOMOLOGADO_AGUARDANDO_INICIO','SUSPENSAO_ATIVA'].includes(v.status)&&today>v.endDate&&v.suspensionEventId,'DOMAIN_STATE_INVALID');
   const date=new Date(Date.parse(v.endDate+'T12:00:00Z')+86400000).toISOString().slice(0,10);
   status='CONCLUIDO';v.returnEventId='EVT_'+x.operationId;effect={type:'RETORNO',effectiveDate:date,endDate:null,id:v.returnEventId};
  }else N(false,'COMMAND_PAYLOAD_INVALID');
  if(payload.action==='DENY'){N(payload.confirmed===true,'COMMAND_PAYLOAD_INVALID');v.decision={action:'DENY',actorUid,at:T(now),decisionReference:payload.minutesReference.trim()||'COMMAND_'+x.operationId,reason:note,exception:false};}
  v.history.push(entry(payload.action,actorUid,now,v.status,status,note));v.status=status;v.revision++;v.updatedAt=T(now);
  return {request:validate(v),effect,terminal:terminal.includes(status)};
 }
 return {modality,terminal,policy,validate,interval,create,review};
}
