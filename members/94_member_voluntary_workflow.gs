// Canonical voluntary workflow. Pure decisions only: authentication and the
// shared transaction/receipt pipeline remain mandatory at the caller boundary.
function f4MemberVoluntaryWorkflow_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),R=f4MemberVoluntaryRules_();
 const terminal=['CONCLUIDO','EXECUTADO','INDEFERIDO','CANCELADO_PELO_MEMBRO','CANCELADO_PELA_DIRETORIA'];
 const modes=['DESLIGAMENTO_APOS_HOMOLOGACAO','DESLIGAMENTO_FIM_SEMESTRE'];
 const copy=v=>F.decode(F.encode(v)),stamp=n=>F.Timestamp.fromMillis(n);
 function text(v,max,required=false){N(typeof v==='string'&&v.length<=max&&(!required||v.trim().length>0),'COMMAND_PAYLOAD_INVALID');return v.trim();}
 function parameter(v){
  N(v&&v.schemaVersion==='member-voluntary-policy-v1'&&C.rev(v.revision)&&v.revision>0&&v.enabled===true,'ACCESS_REVOKED');
  C.exact(v,['schemaVersion','revision','enabled','minutes','presentationBlockDays','approvalRef','source']);
  N(C.id(v.approvalRef)&&v.source&&typeof v.source==='object'&&Number.isInteger(v.presentationBlockDays)&&v.presentationBlockDays>=0&&v.presentationBlockDays<=365,'COMMAND_PAYLOAD_INVALID');
  // Reuse the V2 validator, including the prohibition on silently waiving ata.
  return R.minutes(v.minutes);
 }
 function submission(v){if(v===null)return;C.exact(v,['reasonCategory','documentReference','acknowledged','version']);text(v.reasonCategory,120,true);text(v.documentReference,500);N(v.acknowledged===true&&v.version==='PORTAL_VINCULO_V2','COMMAND_PAYLOAD_INVALID');}
 function legacy(v){if(v.modality==='SUSPENSAO_TEMPORARIA')return {TIPO_SOLICITACAO:'SUSPENSAO_VOLUNTARIA',MODALIDADE_SOLICITADA:v.modality,STATUS_SOLICITACAO:v.status,ID_SEMESTRE_REFERENCIA:v.semesterId,DATA_INICIO_SUSPENSAO:v.startDate,DATA_FIM_SUSPENSAO:v.endDate};return {TIPO_SOLICITACAO:'DESLIGAMENTO_VOLUNTARIO',MODALIDADE_SOLICITADA:v.modality,STATUS_SOLICITACAO:v.status,ID_SEMESTRE_REFERENCIA:v.semesterId,DATA_FIM_SEMESTRE_SNAPSHOT:v.semesterEndSnapshot,PARAMETRO_ATA_DESLIGAMENTO_ID:v.minutesSnapshot.parametroId,ATA_DESLIGAMENTO_TIPO_VALOR_SNAPSHOT:v.minutesSnapshot.tipoValor,ATA_DESLIGAMENTO_BASE_LEGAL_SNAPSHOT:v.minutesSnapshot.baseLegal,ATA_DESLIGAMENTO_OBRIGATORIA_SNAPSHOT:v.minutesSnapshot.valor?'SIM':'NAO'};}
 function validate(v){
  if(v?.schemaVersion==='member-voluntary-suspension-v1')return f4MemberVoluntarySuspension_().validate(v);
  C.exact(v,['schemaVersion','id','revision','personId','linkId','modality','status','reason','semesterId','semesterEndSnapshot','minutesSnapshot','policyRevision','createdAt','updatedAt','createdByUid','history','decision','submission']);
  N(v.schemaVersion==='member-voluntary-request-v1'&&C.op(v.id)&&C.rev(v.revision)&&v.revision>0&&C.id(v.personId)&&C.id(v.linkId)&&C.id(v.createdByUid)&&modes.includes(v.modality),'SCHEMA_INVALID');
  N(['RECEBIDO','EM_ANALISE','AGUARDANDO_COMPLEMENTO','AGENDADO_PARA_ANALISE_FINAL','PRONTO_PARA_ANALISE_FINAL',...terminal].includes(v.status),'SCHEMA_INVALID');
  text(v.reason,4000,true);N(C.rev(v.policyRevision)&&v.policyRevision>0&&Array.isArray(v.history)&&v.history.length>0&&v.history.length<=100,'SCHEMA_INVALID');
  N(Number.isSafeInteger(C.ms(v.createdAt))&&C.ms(v.createdAt)<=C.ms(v.updatedAt),'SCHEMA_INVALID');
  if(v.modality===modes[1]){N(C.id(v.semesterId),'SCHEMA_INVALID');C.civilStart(v.semesterEndSnapshot);}else N(v.semesterId===null&&v.semesterEndSnapshot===null,'SCHEMA_INVALID');
  R.minutes(v.minutesSnapshot);submission(v.submission);
  if(v.decision!==null){C.exact(v.decision,['action','actorUid','at','minutesReference','reason','effectiveDate','calendarChanged']);N(['APPROVE','DENY'].includes(v.decision.action)&&C.id(v.decision.actorUid),'SCHEMA_INVALID');text(v.decision.minutesReference,500);text(v.decision.reason,4000,true);N(Number.isSafeInteger(C.ms(v.decision.at))&&typeof v.decision.calendarChanged==='boolean','SCHEMA_INVALID');if(v.decision.effectiveDate!==null)C.civilStart(v.decision.effectiveDate);}
  return v;
 }
 function index(v,personId){if(!v)return {schemaVersion:'member-voluntary-index-v1',personId,revision:0,activeRequestId:null,requestIds:[]};C.exact(v,['schemaVersion','personId','revision','activeRequestId','requestIds']);N(v.schemaVersion==='member-voluntary-index-v1'&&v.personId===personId&&C.rev(v.revision)&&Array.isArray(v.requestIds)&&v.requestIds.length<=100&&v.requestIds.every(C.op)&&new Set(v.requestIds).size===v.requestIds.length&&(v.activeRequestId===null||v.requestIds.includes(v.activeRequestId)),'SCHEMA_INVALID');return v;}
 function entry(action,actor,now,before,after,note){N(C.id(actor),'AUTH_BINDING_INVALID');return {action,actorUid:actor,at:stamp(now),before,after,note};}
 function create(input){
  if(input.payload?.modality==='SUSPENSAO_TEMPORARIA')return f4MemberVoluntarySuspension_().create(input);
  const {operationId,actorUid,personId,link,state,today,now,policy,semester,priorIndex,payload}=input;
  N(C.op(operationId)&&C.id(actorUid),'COMMAND_PAYLOAD_INVALID');C.exact(payload,['modality','reason','semesterId',...(Object.hasOwn(payload,'submission')?['submission']:[])]);submission(payload.submission??null);N(modes.includes(payload.modality),'COMMAND_PAYLOAD_INVALID');
  R.eligible(link,personId,state,today);const minutes=parameter(policy),i=index(priorIndex,personId);
  N(i.activeRequestId===null&&!i.requestIds.includes(operationId)&&i.requestIds.length<100,'OPERATION_CONFLICT');
  if(payload.modality===modes[1]){N(semester&&semester.id===payload.semesterId&&semester.civilStart<=today&&today<=semester.civilEnd,'DOMAIN_STATE_INVALID');C.civilStart(semester.civilStart);C.civilStart(semester.civilEnd);}else N(payload.semesterId===null,'COMMAND_PAYLOAD_INVALID');
  const request=validate({schemaVersion:'member-voluntary-request-v1',id:operationId,revision:1,personId,linkId:link.id,modality:payload.modality,status:'RECEBIDO',reason:text(payload.reason,4000,true),semesterId:semester?.id||null,semesterEndSnapshot:semester?.civilEnd||null,minutesSnapshot:copy(minutes),policyRevision:policy.revision,createdAt:stamp(now),updatedAt:stamp(now),createdByUid:actorUid,history:[entry('SUBMIT',actorUid,now,null,'RECEBIDO','')],decision:null,submission:payload.submission||null});
  return {request,index:{...i,revision:i.revision+1,activeRequestId:operationId,requestIds:i.requestIds.concat(operationId)}};
 }
 function review(input){
  if(input.request?.modality==='SUSPENSAO_TEMPORARIA')return f4MemberVoluntarySuspension_().review(input);
  const {request:raw,expectedRevision,actorUid,actorPersonId,administrative,today,now,policy,semester,payload}=input,v=copy(validate(raw));
  N(v.revision===expectedRevision,'REVISION_CONFLICT');N(!terminal.includes(v.status),'DOMAIN_STATE_INVALID');C.exact(payload,['action','reason','minutesReference','confirmed']);text(payload.reason,4000,true);text(payload.minutesReference,500);N(typeof payload.confirmed==='boolean','COMMAND_PAYLOAD_INVALID');
  const own=payload.action==='CANCEL_SELF'||payload.action==='COMPLEMENT';N(own?actorPersonId===v.personId:administrative===true,'ACCESS_REVOKED');
  const transitions={START:['EM_ANALISE','INICIAR_ANALISE'],REQUEST_COMPLEMENT:['AGUARDANDO_COMPLEMENTO','SOLICITAR_COMPLEMENTO'],COMPLEMENT:['EM_ANALISE','COMPLEMENTAR'],SCHEDULE:['AGENDADO_PARA_ANALISE_FINAL','ANALISE_PRELIMINAR_REGISTRADA'],READY:['PRONTO_PARA_ANALISE_FINAL','ANALISE_FINAL_DISPONIVEL'],CANCEL_SELF:['CANCELADO_PELO_MEMBRO','CANCELAR'],CANCEL_ADMIN:['CANCELADO_PELA_DIRETORIA','CANCELAR'],DENY:['INDEFERIDO','INDEFERIR'],APPROVE:['EXECUTADO','HOMOLOGAR_E_EFETIVAR_DESLIGAMENTO']};
  const change=transitions[payload.action];N(change,'COMMAND_PAYLOAD_INVALID');
  if(payload.action==='COMPLEMENT')N(v.status==='AGUARDANDO_COMPLEMENTO','DOMAIN_STATE_INVALID');
  if(['SCHEDULE','READY'].includes(payload.action))N(v.modality===modes[1],'DOMAIN_STATE_INVALID');
  if(payload.action==='READY'){N(semester&&semester.id===v.semesterId,'RECONCILIATION_REQUIRED');C.civilStart(semester.civilEnd);N(today>=semester.civilEnd,'DOMAIN_STATE_INVALID');}
  let effect=null;
  if(['APPROVE','DENY'].includes(payload.action)){
   parameter(policy);
   // Rejection also requires human minutes but need not wait for semester end.
   const checked=R.finalDecision(legacy(v),{confirmacaoReforcada:payload.confirmed,ataReferencia:payload.minutesReference},policy.minutes,today,semester,payload.action==='DENY');
   v.decision={action:payload.action,actorUid,at:stamp(now),minutesReference:payload.minutesReference.trim(),reason:payload.reason.trim(),effectiveDate:payload.action==='APPROVE'?checked.effectiveDate:null,calendarChanged:checked.calendarChanged};
   if(payload.action==='APPROVE')effect={personId:v.personId,linkId:v.linkId,effectiveDate:checked.effectiveDate,requestId:v.id,minutesReference:v.decision.minutesReference};
  }
  const status=R.transition(legacy(v),change[0],change[1]);v.history.push(entry(payload.action,actorUid,now,v.status,status,payload.reason.trim()));v.status=status;v.revision++;v.updatedAt=stamp(now);
  return {request:validate(v),effect,terminal:terminal.includes(status)};
 }
 return {validate,index,create,review,parameter,legacy,terminal};
}
