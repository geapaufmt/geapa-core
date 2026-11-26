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
     if(o.state==='COMPLETED'&&o.resolution==='LOCAL_SIMULATION_CONFIRMED_NO_RESEND'&&o.lastError?.mode==='DRY_RUN'&&o.lastError.sent===false){return {...o.lastError,state:'COMPLETED',resolution:o.resolution,replay:true};}
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
  N(['member-notification-index-v1','member-notification-index-v2'].includes(v.schemaVersion)&&v.id===personId&&v.personId===personId&&C.rev(v.revision)&&Number.isSafeInteger(v.unreadCount)&&v.unreadCount>=0&&Array.isArray(v.notificationIds)&&v.notificationIds.every(C.op),'SCHEMA_INVALID');
  return v;
 }
 const capacity=100;
 const routes={'/app/minha-situacao':'minha-situacao','/app/meu-vinculo':'minha-situacao','/ingresso':'ingresso','/admin/membros':'admin-membros'};
 function destination(value){const target=Object.hasOwn(routes,value)?routes[value]:value;return Object.values(routes).includes(target)?target:null;}
 function inbox(get,personId){
  const ix=index(get('memberNotificationIndexes/'+personId),personId);
  N(ix.notificationIds.length<=capacity&&new Set(ix.notificationIds).size===ix.notificationIds.length,'SCHEMA_INVALID');
  const records=ix.notificationIds.map(id=>get('memberNotifications/'+id));
  const items=records.map((item,i)=>{
   const id=ix.notificationIds[i];
   N(item?.personId===personId&&item.id===id,'RECONCILIATION_REQUIRED');
   return record(item);
  });
  return {ix,items,unreadCount:items.filter(item=>!item.read).length};
 }
 function delivery(get,payload,now){
  const {personId,operationId,event}=payload;
  N(C.id(personId)&&C.op(operationId)&&Object.hasOwn(messages,event),'COMMAND_PAYLOAD_INVALID');
  const [title,body,route]=messages[event],targetRoute=destination(payload.targetRoute??route);
  N(targetRoute,'COMMAND_PAYLOAD_INVALID');
  const id=C.hash({personId,operationId}).slice(0,32);
  const content={event,title:payload.title?.trim()||title,body:payload.body?.trim()||body,targetRoute,category:payload.category||'OPERACIONAL',metadata:payload.metadata||{}};
  const notification=record({schemaVersion:'member-notification-v1',id,personId,operationId,revision:1,...content,createdAt:now,read:false,readAt:null});
  const ix=index(get('memberNotificationIndexes/'+personId),personId);
  const legacy=get('memberNotifications/'+operationId);
  const previous=get('memberNotifications/'+id) || (legacy?.personId===personId?legacy:null);
  if(previous){
   record(previous);
   N(previous.personId===personId&&previous.operationId===operationId&&
    Object.keys(content).every(key=>C.hash(key==='targetRoute'?destination(previous[key]):previous[key])===C.hash(content[key])),'OPERATION_CONFLICT');
   return {notification:previous,index:ix,replay:true,writes:[]};
  }
  N(ix.notificationIds.length<=capacity&&new Set(ix.notificationIds).size===ix.notificationIds.length&&!ix.notificationIds.includes(id),'RECONCILIATION_REQUIRED');
  // V1 counts cover the visible window. Recount full legacy indexes before eviction.
  const retained=ix.notificationIds.slice(0,capacity-1);
  let unread=ix.unreadCount;
  if(ix.schemaVersion==='member-notification-index-v2'){
   N(unread<=ix.notificationIds.length,'SCHEMA_INVALID');
   if(ix.notificationIds.length===capacity){
    const last=ix.notificationIds[capacity-1],evicted=get('memberNotifications/'+last);
    N(evicted?.personId===personId&&evicted.id===last,'RECONCILIATION_REQUIRED');
    if(!record(evicted).read)unread--;
    N(unread>=0,'SCHEMA_INVALID');
   }
  }else{
   const box=inbox(get,personId);
   unread=box.items.slice(0,capacity-1).filter(item=>!item.read).length;
  }
  const next={...ix,schemaVersion:'member-notification-index-v2',notificationIds:[id,...retained],revision:ix.revision+1,updatedAt:now,unreadCount:unread+1};
  const writes=[{kind:'set',path:'memberNotifications/'+id,data:notification},{kind:'set',path:'memberNotificationIndexes/'+personId,data:next}];
  return {notification,index:next,replay:false,writes};
 }
 function emit(tx,payload,now){
  const {writes,...result}=delivery(p=>tx.get(p).data(),payload,now);
  for(const w of writes)tx.set(w.path,w.data);
  return result;
 }
 function ownPerson(get,identity,now){
  return f4MemberOwnReader_().read(get,identity,{view:'status'},now).data.idPessoa;
 }
 function view(items,personId,now){
  return {schemaVersion:'member-notification-view-v1',personId,asOf:now,totalCount:items.length,
   unreadCount:items.filter(item=>!item.read).length,items:items.map(({id,title,body,event,read,readAt,createdAt,targetRoute})=>
    ({id,title,body,event,read,readAt,createdAt,targetRoute:destination(targetRoute)}))};
 }
 function readAuthorized(get,identity,now){
  const person=ownPerson(get,identity,now);
  return view(inbox(get,person).items,person,now);
 }
 function mark(tx,identity,params,now){
  const get=p=>tx.get(p).data(),person=ownPerson(get,identity,now);
  const all=params?.action==='MARK_ALL_READ';
  C.exact(params,all?['action']:['action','notificationId']);
  N(all||params.action==='MARK_READ'&&C.op(params.notificationId),'COMMAND_PAYLOAD_INVALID');
  const {ix,items,unreadCount}=inbox(get,person);
  N(all||ix.notificationIds.includes(params.notificationId),'RESOURCE_NOT_FOUND');
  let count=0;
  const updated=items.map(item=>{
   if(item.read||!all&&item.id!==params.notificationId)return item;
   const next={...item,read:true,readAt:now,revision:item.revision+1};
   tx.set('memberNotifications/'+item.id,next);count++;
   return next;
  });
  if(count||ix.unreadCount!==unreadCount)tx.set('memberNotificationIndexes/'+person,{...ix,revision:ix.revision+1,updatedAt:now,unreadCount:unreadCount-count});
  return view(updated,person,now);
 }
 return {plan,source,render,consumer,delivery,emit,readAuthorized,mark};
}
