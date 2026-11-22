// Existing jobOutbox and Mail Hub renderer; DEV delivery is an explicit dry-run.
// This adapter neither reads Sheets nor sends Gmail. Canonical changes commit
// independently of delivery, and the outbox is replayed by its command key.
function f4MemberNotifications_(){
 const C=f4Core_(),N=C.need,K=f4MemberCommandCatalog_();
 const messages={
  ENROLLMENT_REGISTERED:['Novo pré-cadastro de membro','Um pré-cadastro está disponível para conferência da Presidência.','/admin/membros'],
  ENROLLMENT_INVITATION:['Convite de ingresso no GEAPA','Seu ingresso foi previamente aprovado. Acesse o Portal para aceitar ou recusar o convite e completar seu cadastro.','/ingresso'],
  ENROLLMENT_ACCEPTED:['Convite aceito — cadastro pendente','Seu aceite foi registrado. Complete os dados obrigatórios e registre a ciência das normas para confirmar o ingresso.','/ingresso'],
  ENROLLMENT_DECLINED:['Recusa de convite registrada','Sua decisão foi registrada. Nenhum vínculo foi ativado.','/ingresso'],
  ENROLLMENT_COMPLETED:['Ingresso confirmado','Seu cadastro foi concluído e o vínculo foi ativado. Consulte sua situação no Portal.','/app/minha-situacao'],
  VOLUNTARY_RECEIVED:['Solicitação recebida','Sua solicitação foi registrada e aguarda análise humana. O protocolo não altera seu vínculo.','/app/meu-vinculo'],
  VOLUNTARY_DECIDED:['Solicitação atualizada','A análise de sua solicitação foi registrada. Consulte a decisão e as datas aplicáveis no Portal.','/app/meu-vinculo'],
  VOLUNTARY_EXIT:['Desligamento voluntário registrado','A decisão humana sobre o desligamento foi registrada. Consulte a data de efeito e o histórico da solicitação no Portal.','/app/meu-vinculo'],
  VOLUNTARY_SUSPENSION:['Suspensão voluntária registrada','A decisão de suspensão temporária foi registrada. Consulte o período e o histórico no Portal.','/app/meu-vinculo'],
  VOLUNTARY_RETURN:['Retorno de suspensão registrado','O retorno da suspensão foi registrado no seu vínculo.','/app/meu-vinculo'],
  SYNTHETIC_TEST:['Notificação de ensaio sintético','Evento sintético gerado para validação da Central de Notificações.','/app/minha-situacao'],
  SYSTEM_ANNOUNCEMENT:['Comunicado institucional','Aviso institucional do GEAPA aos membros.','/app/minha-situacao']
 };
 function policy(get){const p=get('memberOperationalConfig/notifications');N(p?.schemaVersion==='member-notification-policy-v1'&&p.enabled===true&&p.mode==='DRY_RUN'&&p.projectId===f4ExecutionEnvironment_().project&&p.environment===f4ExecutionEnvironment_().environment&&C.rev(p.revision)&&p.revision>0&&C.id(p.approvalRef),'ACCESS_REVOKED');return p;}
 function event(r,changes){
  if(r.commandType==='members.ingress.execute'&&r.payload.operation==='PRE_REGISTER')return 'ENROLLMENT_REGISTERED';
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
  if(o.event==='ENROLLMENT_REGISTERED'){
   const recipient=policy(get).presidencyEmail;
   N(typeof recipient==='string'&&/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient),'UNRESOLVED_DEPENDENCY');
   return {outbox:o,recipient};
  }
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
  return {
   preview(id){
    N(C.op(id),'COMMAND_PAYLOAD_INVALID');
    return db.transaction(tx=>{
     const get=p=>tx.get(p).data(),now=clock();
     N(identity&&now<identity.expiresAt&&f4DevAdmin_().profile(get,identity.uid,now),'ACCESS_REVOKED');
     const v=render(get,id);
     return {operationId:id,recipient:v.recipient,rendered:v.rendered,policyRevision:v.policyRevision,brandRevision:v.brandRevision||null,suppressed:v.suppressed||false,mode:'DRY_RUN',isRealMail:false};
    });
   },
   run(id,options={}){
    N(C.op(id),'COMMAND_PAYLOAD_INVALID');
    return db.transaction(tx=>{
     const get=p=>tx.get(p).data(),now=clock();N(identity&&now<identity.expiresAt&&f4DevAdmin_().profile(get,identity.uid,now),'ACCESS_REVOKED');
     const {outbox:o}=source(get,id);policy(get);
     if(o.state==='COMPLETED'){N(o.notificationReceipt?.mode==='DRY_RUN'&&o.notificationReceipt.sent===false,'RECONCILIATION_REQUIRED');return {...o.notificationReceipt,replay:true};}
     N(o.state==='PENDING','RECONCILIATION_REQUIRED');
     if(options&&options.simulateFailure){
      const failureReceipt={mode:'DRY_RUN',sent:false,operationId:id,event:o.event,actorUid:identity.uid,failedAt:now,channel:'SIMULATED_EMAIL',error:options.failureCode||'EXTERNAL_CHANNEL_FAILED',errorMessage:options.failureMessage||'Falha simulada no canal de entrega externa.'};
      tx.set('jobOutbox/'+id,{...o,state:'RECONCILIATION_REQUIRED',attempts:o.attempts+1,lastError:failureReceipt});
      tx.set('commandReconciliation/'+id,{active:true,reason:'EXTERNAL_CHANNEL_FAILED',idempotencyKey:id,at:now});
      return {...failureReceipt,state:'RECONCILIATION_REQUIRED',replay:false};
     }
     const v=render(get,id),receipt={mode:'DRY_RUN',sent:false,operationId:id,event:o.event,policyRevision:v.policyRevision,brandRevision:v.brandRevision||null,contentDigest:v.suppressed?null:C.hash(v.rendered),recipientDigest:C.hash({email:v.recipient}),actorUid:identity.uid,completedAt:now,channel:'SIMULATED_EMAIL',...(v.suppressed?{suppressed:true,reason:v.reason}:{})};
     tx.set('jobOutbox/'+id,{...o,state:'COMPLETED',attempts:o.attempts+1,notificationReceipt:receipt});return {...receipt,replay:false};
    });
   }
  };
 }
 function record(v){
  C.exact(v,['schemaVersion','id','personId','revision','event','title','body','targetRoute','category','read','readAt','createdAt','operationId','metadata']);
  N(v.schemaVersion==='member-notification-v1'&&C.op(v.id)&&C.id(v.personId)&&C.rev(v.revision)&&v.revision>0,'SCHEMA_INVALID');
  N(typeof v.title==='string'&&v.title.length<=200&&typeof v.body==='string'&&v.body.length<=2000,'SCHEMA_INVALID');
  N(typeof v.read==='boolean'&&(v.readAt===null||Number.isSafeInteger(v.readAt)),'SCHEMA_INVALID');
  N(Number.isSafeInteger(v.createdAt)&&(v.operationId===null||C.op(v.operationId)),'SCHEMA_INVALID');
  return v;
 }
 function index(v,personId){
  if(!v)return {schemaVersion:'member-notification-index-v1',id:personId,personId,revision:0,unreadCount:0,notificationIds:[],updatedAt:0};
  C.exact(v,['schemaVersion','id','personId','revision','unreadCount','notificationIds','updatedAt']);
  N(v.schemaVersion==='member-notification-index-v1'&&v.id===personId&&v.personId===personId&&C.rev(v.revision)&&Number.isSafeInteger(v.unreadCount)&&v.unreadCount>=0&&Array.isArray(v.notificationIds)&&v.notificationIds.every(C.op),'SCHEMA_INVALID');
  return v;
 }
 function emit(tx,payload,now){
  const {personId,event:eventKind,title,body,targetRoute,category,operationId,metadata}=payload;
  N(C.id(personId),'COMMAND_PAYLOAD_INVALID');
  const get=p=>tx.get(p).data();
  const ix=index(get('memberNotificationIndexes/'+personId),personId);
  const notifId=operationId||C.hash({personId,event:eventKind,now,seq:ix.notificationIds.length}).slice(0,32);
  const old=get('memberNotifications/'+notifId);
  if(old)return {notification:record(old),index:ix,replay:true};
  const [defaultTitle,defaultBody,defaultRoute]=messages[eventKind]||['Notificação institucional','Nova mensagem do GEAPA.','/app/minha-situacao'];
  const notif={
   schemaVersion:'member-notification-v1',
   id:notifId,
   personId,
   revision:1,
   event:eventKind||'SYSTEM_ANNOUNCEMENT',
   title:typeof title==='string'&&title.trim().length?title.trim():defaultTitle,
   body:typeof body==='string'&&body.trim().length?body.trim():defaultBody,
   targetRoute:typeof targetRoute==='string'?targetRoute:defaultRoute,
   category:category||'OPERACIONAL',
   read:false,
   readAt:null,
   createdAt:now,
   operationId:operationId||null,
   metadata:metadata||{}
  };
  record(notif);
  const nextIds=[notifId,...ix.notificationIds.filter(x=>x!==notifId)].slice(0,100);
  const nextIx={
   schemaVersion:'member-notification-index-v1',
   id:personId,
   personId,
   revision:ix.revision+1,
   unreadCount:ix.unreadCount+1,
   notificationIds:nextIds,
   updatedAt:now
  };
  index(nextIx,personId);
  tx.set('memberNotifications/'+notifId,notif);
  tx.set('memberNotificationIndexes/'+personId,nextIx);
  return {notification:notif,index:nextIx,replay:false};
 }
 function resolvePersonId(get,identity,now){
  const binding=get('memberAuthBindings/'+identity.uid);
  if(binding?.idPessoa)return binding.idPessoa;
  const technical=f4DevAdmin_().profile(get,identity.uid,now);
  if(technical)return identity.uid;
  const actor=f4MemberAuthz_().identify(get,identity.uid,now);
  return actor.binding.idPessoa;
 }
 function readAuthorized(get,identity,now,params={}){
  C.exact(identity,['uid','expiresAt']);
  N(C.id(identity.uid)&&Number.isSafeInteger(identity.expiresAt)&&now<identity.expiresAt,'AUTH_BINDING_INVALID');
  const personId=resolvePersonId(get,identity,now);
  N(C.id(personId),'AUTH_BINDING_INVALID');
  if(params.personId){
   const technical=f4DevAdmin_().profile(get,identity.uid,now);
   N(params.personId===personId||technical,'ACCESS_REVOKED');
  }
  const targetPerson=params.personId||personId;
  const ix=index(get('memberNotificationIndexes/'+targetPerson),targetPerson);
  const items=ix.notificationIds.map(id=>get('memberNotifications/'+id)).filter(Boolean).map(record);
  const unreadCount=items.filter(i=>!i.read).length;
  return {
   schemaVersion:'member-notification-view-v1',
   personId:targetPerson,
   unreadCount,
   totalCount:items.length,
   items,
   asOf:now
  };
 }
 function markRead(tx,identity,notificationId,now){
  C.exact(identity,['uid','expiresAt']);
  N(C.id(identity.uid)&&Number.isSafeInteger(identity.expiresAt)&&now<identity.expiresAt,'AUTH_BINDING_INVALID');
  N(C.op(notificationId),'COMMAND_PAYLOAD_INVALID');
  const get=p=>tx.get(p).data(),personId=resolvePersonId(get,identity,now);
  N(C.id(personId),'AUTH_BINDING_INVALID');
  const raw=get('memberNotifications/'+notificationId);
  N(raw,'NOTIFICATION_NOT_FOUND');
  const notif=record(raw);
  const technical=f4DevAdmin_().profile(get,identity.uid,now);
  N(notif.personId===personId||technical,'ACCESS_REVOKED');
  const targetPerson=notif.personId,ix=index(get('memberNotificationIndexes/'+targetPerson),targetPerson);
  if(notif.read)return {ok:true,notification:notif,unreadCount:ix.unreadCount,replay:true};
  const updatedNotif={...notif,revision:notif.revision+1,read:true,readAt:now};
  record(updatedNotif);
  const newUnread=Math.max(0,ix.unreadCount-1);
  const updatedIx={...ix,revision:ix.revision+1,unreadCount:newUnread,updatedAt:now};
  index(updatedIx,targetPerson);
  tx.set('memberNotifications/'+notificationId,updatedNotif);
  tx.set('memberNotificationIndexes/'+targetPerson,updatedIx);
  return {ok:true,notification:updatedNotif,unreadCount:newUnread,replay:false};
 }
 function markAllRead(tx,identity,now){
  C.exact(identity,['uid','expiresAt']);
  N(C.id(identity.uid)&&Number.isSafeInteger(identity.expiresAt)&&now<identity.expiresAt,'AUTH_BINDING_INVALID');
  const get=p=>tx.get(p).data(),personId=resolvePersonId(get,identity,now);
  N(C.id(personId),'AUTH_BINDING_INVALID');
  const ix=index(get('memberNotificationIndexes/'+personId),personId);
  if(ix.unreadCount===0)return {ok:true,personId,unreadCount:0,updated:0,replay:true};
  let count=0;
  for(const id of ix.notificationIds){
   const raw=get('memberNotifications/'+id);
   if(raw&&!raw.read){
    const updated={...raw,revision:raw.revision+1,read:true,readAt:now};
    record(updated);
    tx.set('memberNotifications/'+id,updated);
    count++;
   }
  }
  const updatedIx={...ix,revision:ix.revision+1,unreadCount:0,updatedAt:now};
  index(updatedIx,personId);
  tx.set('memberNotificationIndexes/'+personId,updatedIx);
  return {ok:true,personId,unreadCount:0,updated:count,replay:false};
 }
 return {plan,source,render,consumer,record,index,emit,readAuthorized,markRead,markAllRead,messages};
}
