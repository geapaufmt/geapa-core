// Administrative verification is independent of operational/attendance facts.
// Initialization is an explicitly approved migration, never a browser command.
function f4ActivityReview_(){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_();
 const command='ACTIVITY_ADMIN_REVIEW',fields=['tituloPublico','local','descricaoPublica','descricaoInterna','responsavelInterno','responsavelEmail','publicoAlvo','observacoes'];
 function validate(v){C.exact(v,['schemaVersion','state','cycleRef','revision','initializedAt','initializationRef','confirmedAt','confirmedBy','confirmationOperationId']);N(v.schemaVersion==='activity-administrative-review-v1'&&['PENDING','CONFIRMED'].includes(v.state)&&v.cycleRef?.kind==='cycle'&&M.id(v.cycleRef.id)&&Number.isSafeInteger(v.revision)&&v.revision>0&&M.id(v.initializationRef)&&Number.isFinite(M.ms(v.initializedAt)),'VALIDATION_ERROR');if(v.state==='PENDING')N(v.confirmedAt===null&&v.confirmedBy===null&&v.confirmationOperationId===null,'VALIDATION_ERROR');else N(Number.isFinite(M.ms(v.confirmedAt))&&v.confirmedBy?.kind==='person'&&M.id(v.confirmedBy.id)&&C.op(v.confirmationOperationId),'VALIDATION_ERROR');return v;}
 function initial(cycleId,ref,now){N(M.id(cycleId)&&M.id(ref),'VALIDATION_ERROR');return {schemaVersion:'activity-administrative-review-v1',state:'PENDING',cycleRef:M.ref('cycle',cycleId),revision:1,initializedAt:M.stamp(now),initializationRef:ref,confirmedAt:null,confirmedBy:null,confirmationOperationId:null};}
 function guard(old,next,cmd){
  if(!old?.administrativeReview)return;
  const v=validate(old.administrativeReview);
  if(cmd!==command)N(M.hash(f4Firestore_().encode(next.administrativeReview||null))===M.hash(f4Firestore_().encode(v)),'AUTHORIZATION_DENIED','REVIEW_STATE_PROTECTED');
  if(v.state!=='CONFIRMED')return;
  const protectedFields=['title','location','format','startsAt','endsAt','civilDate','details','typeRef','typeSnapshot','typeRevision','modelContent','cycleRef','cycleId','academicSemesterId','personRef','guestParticipants','historicalPresentationRefs'];
  N(protectedFields.every(k=>C.hash(f4Firestore_().encode(old[k]??null))===C.hash(f4Firestore_().encode(next[k]??null))),'INVALID_STATE_TRANSITION','ACTIVITY_REVIEW_LOCKED');
 }
 function dto(a,admin){const v=a.administrativeReview;if(!v)return {};validate(v);return {estadoConferencia:v.state==='PENDING'?'PENDENTE DE CONFERÊNCIA':'CONFERIDA',conferenciaRevision:v.revision,...(admin?{conferenciaResponsavel:v.confirmedBy?.id||null,conferenciaEm:v.confirmedAt?new Date(M.ms(v.confirmedAt)).toISOString():null,podeConferir:v.state==='PENDING',camposConferencia:v.state==='PENDING'?fields:[]}:{} )};}
 function handle(x){const {p,load,revision,event,actor,now,operationId}=x;
  C.exact(p,['activityRef','action','patch','reason','confirmed']);N(['EDIT','CONFIRM'].includes(p.action)&&typeof p.reason==='string'&&p.reason.trim().length>0&&p.reason.length<=2000&&typeof p.confirmed==='boolean'&&p.patch&&Object.getPrototypeOf(p.patch)===Object.prototype,'VALIDATION_ERROR');
  N(actor.kind==='BOARD_ACTOR','AUTHORIZATION_DENIED');const a=load(p.activityRef,true),v=validate(a.administrativeReview);N(v.state==='PENDING','INVALID_STATE_TRANSITION','ACTIVITY_REVIEW_LOCKED');const cycle=load(v.cycleRef);N(cycle?.id===v.cycleRef.id,'UNRESOLVED_DEPENDENCY');
  N(!a.cycleId||a.cycleId===cycle.id,'VALIDATION_ERROR','REVIEW_CYCLE_MISMATCH');N(!a.cycleRef?.id||a.cycleRef.id===cycle.id,'VALIDATION_ERROR','REVIEW_CYCLE_MISMATCH');
  const keys=Object.keys(p.patch);N(keys.every(k=>fields.includes(k)),'AUTHORIZATION_DENIED','PROTECTED_REVIEW_FIELD');
  const next={...a,details:{...(a.details||{})}};
  // Optional corrections and explicit confirmation share one authorized CAS,
  // revision and audit event. No preliminary EDIT receipt is required.
  if(p.action==='CONFIRM')N(p.confirmed===true,'HUMAN_DECISION_REQUIRED');
  else N(keys.length>0&&p.confirmed===false,'VALIDATION_ERROR');
  if(keys.length)N(!['ARCHIVED','ARQUIVADA'].includes(a.status||a.legacyStatus),'INVALID_STATE_TRANSITION');
  for(const k of keys){const text=p.patch[k];N(typeof text==='string'&&text.length<=(k==='tituloPublico'||k==='local'?240:2000)&&!/[<>\x00-\x08]/.test(text),'VALIDATION_ERROR');if(k==='tituloPublico'){N(text.trim(),'INCOMPLETE_REQUEST');next.title=text.trim();}else if(k==='local')next.location=text.trim();else next.details[k]=text.trim();}
  next.administrativeReview=p.action==='CONFIRM'?{...v,state:'CONFIRMED',revision:v.revision+1,confirmedAt:M.stamp(now),confirmedBy:actor.personRef,confirmationOperationId:operationId}:{...v,revision:v.revision+1};
  revision(p.activityRef,next);event('history',{action:p.action==='CONFIRM'?'ACTIVITY_ADMIN_CONFIRMED':'ACTIVITY_ADMIN_CORRECTED',subjectRef:p.activityRef,ownerPersonRef:'',visibility:'BOARD',summary:p.reason.trim(),changedFields:keys,previousActivityRevision:a.revision,occurredAt:M.stamp(now)});
  return {activityRef:p.activityRef,revision:a.revision+1,reviewState:next.administrativeReview.state};
 }
 function prepare(options,input){const p=input.payload;C.exact(p,['idAtividade','revision','action','campos','motivo','confirmado']);N(M.id(p.idAtividade)&&Number.isSafeInteger(p.revision)&&p.revision>0,'VALIDATION_ERROR');const a=f4Firestore_().connect(options.config).get('canonicalActivities/'+p.idAtividade);N(a,'NOT_FOUND');N(a.revision===p.revision,'STALE_REVISION');return options.ports.prepare({operationId:input.operationId,commandType:'ACT.'+command,payload:{activityRef:M.ref('activity',a.id),action:p.action,patch:p.campos,reason:p.motivo,confirmed:p.confirmado}});}
 return {command,fields,validate,initial,guard,dto,handle,prepare};
}
