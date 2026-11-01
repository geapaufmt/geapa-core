// Existing jobOutbox and Mail Hub renderer; DEV delivery is an explicit dry-run.
// This adapter neither reads Sheets nor sends Gmail. Canonical changes commit
// independently of delivery, and the outbox is replayed by its command key.
function f4MemberNotifications_(){
 const C=f4Core_(),N=C.need,K=f4MemberCommandCatalog_();
 const messages={
  ENROLLMENT_INVITATION:['Convite de ingresso no GEAPA','Seu ingresso foi previamente aprovado. Acesse o Portal para aceitar ou recusar o convite e completar seu cadastro.','/ingresso'],
  ENROLLMENT_ACCEPTED:['Convite aceito — cadastro pendente','Seu aceite foi registrado. Complete os dados obrigatórios e registre a ciência das normas para confirmar o ingresso.','/ingresso'],
  ENROLLMENT_DECLINED:['Recusa de convite registrada','Sua decisão foi registrada. Nenhum vínculo foi ativado.','/ingresso'],
  ENROLLMENT_COMPLETED:['Ingresso confirmado','Seu cadastro foi concluído e o vínculo foi ativado. Consulte sua situação no Portal.','/app/minha-situacao'],
  VOLUNTARY_RECEIVED:['Solicitação recebida','Sua solicitação foi registrada e aguarda análise humana. O protocolo não altera seu vínculo.','/app/meu-vinculo'],
  VOLUNTARY_DECIDED:['Solicitação atualizada','A análise de sua solicitação foi registrada. Consulte a decisão e as datas aplicáveis no Portal.','/app/meu-vinculo'],
  VOLUNTARY_EXIT:['Desligamento voluntário registrado','A decisão humana sobre o desligamento foi registrada. Consulte a data de efeito e o histórico da solicitação no Portal.','/app/meu-vinculo'],
  VOLUNTARY_SUSPENSION:['Suspensão voluntária registrada','A decisão de suspensão temporária foi registrada. Consulte o período e o histórico no Portal.','/app/meu-vinculo'],
  VOLUNTARY_RETURN:['Retorno de suspensão registrado','O retorno da suspensão foi registrado no seu vínculo.','/app/meu-vinculo']
 };
 function policy(get){const p=get('memberOperationalConfig/notifications');N(p?.schemaVersion==='member-notification-policy-v1'&&p.enabled===true&&p.mode==='DRY_RUN'&&p.projectId===f4ExecutionEnvironment_().project&&p.environment===f4ExecutionEnvironment_().environment&&C.rev(p.revision)&&p.revision>0&&C.id(p.approvalRef),'ACCESS_REVOKED');return p;}
 function event(r,changes){
  if(f4MemberEnrollmentCommands_().is(r))return {ACCEPT:'ENROLLMENT_ACCEPTED',DECLINE:'ENROLLMENT_DECLINED',COMPLETE:'ENROLLMENT_COMPLETED'}[r.payload.action]||null;
  if(r.commandType==='members.voluntary.submit')return 'VOLUNTARY_RECEIVED';
  if(r.commandType==='members.membership.voluntaryExit')return 'VOLUNTARY_EXIT';
  if(r.commandType==='members.voluntary.review'){
   if(!['APPROVE','RETURN','DENY','REQUEST_COMPLEMENT','CANCEL_ADMIN'].includes(r.payload.action))return null;
   return r.payload.requestKind==='SUSPENSAO_TEMPORARIA'&&['APPROVE','RETURN'].includes(r.payload.action)?r.payload.action==='RETURN'?'VOLUNTARY_RETURN':'VOLUNTARY_SUSPENSION':'VOLUNTARY_DECIDED';
  }
  return null;
 }
 function plan(get,r,changes){
  const kind=event(r,changes);if(!kind)return null;
  const configured=get('memberOperationalConfig/notifications');if(!configured||configured.enabled===false)return null;
  const p=policy(get),request=changes.find(c=>c.path.startsWith('memberVoluntaryRequests/'))?.data;
  return {schemaVersion:'member-notification-outbox-v1',operationId:r.operationId,commandType:r.commandType,kind:'MEMBERS_NOTIFICATION',event:kind,state:'PENDING',idempotencyKey:r.operationId,attempts:0,personId:r.target.personId,enrollmentId:r.target.enrollmentId||null,requestId:request?.id||null,policyRevision:p.revision,projectId:p.projectId,environment:p.environment,mode:p.mode,sourceDigest:r.digest};
 }
 function source(get,id){
  const o=get('jobOutbox/'+id),r=get('commandRequests/'+id),l=get('commandLedger/'+id),audit=get('memberAudit/'+id);
  N(o?.kind==='MEMBERS_NOTIFICATION'&&o.schemaVersion==='member-notification-outbox-v1'&&o.operationId===id&&o.idempotencyKey===id&&o.environment===f4ExecutionEnvironment_().environment&&o.projectId===f4ExecutionEnvironment_().project&&o.mode==='DRY_RUN'&&messages[o.event],'COMMAND_PAYLOAD_INVALID');
  N(r?.operationId===id&&r.digest===o.sourceDigest&&r.target.personId===o.personId&&l?.state==='COMPLETED'&&l.digest===r.digest&&audit?.digest===r.digest,'RECONCILIATION_REQUIRED');
  if(o.event==='ENROLLMENT_INVITATION')N(r.commandType==='members.ingress.execute'&&r.payload.operation==='MANAGE_INVITATION'&&['SEND','RESEND'].includes(r.payload.action)&&r.payload.enrollmentId===o.enrollmentId,'RECONCILIATION_REQUIRED');
  else N(event(r,[])===o.event,'RECONCILIATION_REQUIRED');
  if(o.enrollmentId){const e=f4MemberEnrollmentContract_().validate(get('memberEnrollments/'+o.enrollmentId));N(e.personId===o.personId,'AUTH_BINDING_INVALID');const superseded=o.event==='ENROLLMENT_INVITATION'&&e.status!=='INVITED'||o.event==='ENROLLMENT_ACCEPTED'&&e.status!=='ACCEPTED';return {outbox:o,recipient:e.registration.email,superseded};}
  const h=get('memberAggregateHeads/'+o.personId);N(h?.id===o.personId&&Array.isArray(h.paths),'AUTH_BINDING_INVALID');
  const addresses=[...new Set(h.paths.filter(p=>p.startsWith('memberIdentifiers/')).map(get).filter(v=>v?.idPessoa===o.personId&&v.active&&v.kind==='EMAIL').map(v=>v.value))];
  N(addresses.length===1,'RECONCILIATION_REQUIRED');return {outbox:o,recipient:addresses[0]};
 }
 function render(get,id){
  const {outbox:o,recipient,superseded}=source(get,id),p=policy(get);N(p.revision===o.policyRevision,'REVISION_CONFLICT');
  if(superseded)return {recipient,policyRevision:p.revision,suppressed:true,reason:'ENROLLMENT_STATE_CHANGED'};
  const brand=get('memberOperationalConfig/mailBrand');
  N(brand?.schemaVersion==='member-mail-brand-v1'&&C.rev(brand.revision)&&brand.revision>0&&C.id(brand.approvalRef),'SCHEMA_INVALID');
  const b=brand.values;N(b&&['orgName','shortName','quote'].every(k=>typeof b[k]==='string'&&b[k].trim().length>0&&b[k].length<=500),'SCHEMA_INVALID');
  N(['officialEmail','universityName','universityShortName','instituteName','instituteShortName','courseName'].every(k=>typeof b[k]==='string'&&b[k].length<=200)&&['greenColor','brownColor','blackColor'].every(k=>/^#[0-9a-f]{6}$/i.test(b[k])),'SCHEMA_INVALID');
  const [subject,text,route]=messages[o.event],url='https://geapa-convergencia-dev.web.app/#'+route;
  const rendered=coreMailRenderEmailTemplate_(o.event==='ENROLLMENT_INVITATION'?'GEAPA_CONVITE':'GEAPA_OPERACIONAL','[DEV — SIMULAÇÃO] '+subject,{introText:text,cta:{label:'Acessar o Portal DEV',url},footerNote:'Mensagem de teste. Nenhum e-mail foi enviado.'},b);
  N(typeof recipient==='string'&&/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient),'SCHEMA_INVALID');
  return {recipient,rendered,policyRevision:p.revision,brandRevision:brand.revision};
 }
 function consumer({config,identity,clock=Date.now}){
  N(config.projectId==='geapa-convergencia-dev'&&f4DevBoundary_().permit(config),'ACCESS_REVOKED');const db=f4Firestore_().connect(config);
  return {run(id){N(C.op(id),'COMMAND_PAYLOAD_INVALID');return db.transaction(tx=>{
   const get=p=>tx.get(p).data(),now=clock();N(identity&&now<identity.expiresAt&&f4DevAdmin_().profile(get,identity.uid,now),'ACCESS_REVOKED');
   const {outbox:o}=source(get,id);policy(get);
   if(o.state==='COMPLETED'){N(o.notificationReceipt?.mode==='DRY_RUN'&&o.notificationReceipt.sent===false,'RECONCILIATION_REQUIRED');return {...o.notificationReceipt,replay:true};}
   N(o.state==='PENDING','RECONCILIATION_REQUIRED');const v=render(get,id),receipt={mode:'DRY_RUN',sent:false,operationId:id,event:o.event,policyRevision:v.policyRevision,brandRevision:v.brandRevision||null,contentDigest:v.suppressed?null:C.hash(v.rendered),recipientDigest:C.hash({email:v.recipient}),actorUid:identity.uid,completedAt:now,...(v.suppressed?{suppressed:true,reason:v.reason}:{})};
   tx.set('jobOutbox/'+id,{...o,state:'COMPLETED',attempts:o.attempts+1,notificationReceipt:receipt});return {...receipt,replay:false};
  });}};
 }
 return {plan,source,render,consumer};
}
