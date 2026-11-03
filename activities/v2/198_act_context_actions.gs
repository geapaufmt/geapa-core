// Request-local affordances, never an admission or a portable authorization proof.
function f4ActivityContextActions_(options,id){
 const M=f4ActivitiesModel_(),N=M.need,ports=options.ports,clock=options.clock||Date.now;
 N(M.id(id),'VALIDATION_ERROR');
 return ports.readWithAuthoritySnapshot((check,get)=>{
  const ref=M.ref('activity',id),auth=ports.readAuthority([M.path(ref)],false),a=get(M.path(ref));
  N(a&&a.sourceSnapshotState!=='ABSENT_FROM_CURRENT_SNAPSHOT','RESOURCE_NOT_FOUND');check([auth]);
  const now=clock(),c=get('activityAttendanceCollections/'+id),historical=a.recordMode==='HISTORICAL_PARTIAL',locked=a.administrativeReview?.state==='CONFIRMED';
  const actions=[],memo=new Map();
  function allowed(command,refs=[ref]){const key=command+refs.map(M.path).join();if(!memo.has(key))memo.set(key,ports.commandAccess(command,refs));return memo.get(key);}
  function add(id,label,command,enabled,reason,refs){if(!command||allowed(command,refs))actions.push({id,label,enabled:!!enabled,reason:enabled?'':reason});}
  add('detail','Ver detalhes',null,true,'');
  add('edit','Editar','ACTIVITY_SAVE',!historical&&!locked&&!a.realizationRef&&!['CANCELLED','ARCHIVED'].includes(a.status),locked?'Atividade conferida: edição bloqueada.':historical?'Registro histórico: utilize a conferência autorizada.':'O estado da atividade não permite edição comum.');
  if(a.administrativeReview)add('review','Conferência administrativa','ACTIVITY_ADMIN_REVIEW',!locked, 'Atividade já conferida.');
  const normal=!historical,notCancelled=!['CANCELLED','ARCHIVED'].includes(a.status);
  if(a.publicationStatus==='PUBLICADO')add('hide','Ocultar','ACTIVITY_PUBLICATION_SET',normal,'Registro histórico somente leitura.');
  else add('publish','Publicar','ACTIVITY_PUBLICATION_SET',normal&&notCancelled&&!a.realizationRef&&!!(a.title&&a.location&&a.format&&a.typeSnapshot?.portalModel),'Estado ou dados obrigatórios impedem a publicação.');
  if(a.status==='CANCELLED')add('reopen','Reabrir','ACTIVITY_PUBLICATION_SET',normal,'Registro histórico somente leitura.');
  else add('cancel','Cancelar atividade','ACTIVITY_PUBLICATION_SET',normal&&notCancelled&&!a.realizationRef,'Atividade histórica, encerrada ou já realizada.');
  add('realize','Registrar realização','ACTIVITY_REALIZE',normal&&notCancelled&&!a.realizationRef&&a.civilDate<=M.day(now),'A atividade precisa ser operacional, ainda não realizada e ocorrer até hoje.');
  add('attendance','Gerenciar chamada','ATTENDANCE_FINALIZE',normal&&notCancelled&&!!c,'Não existe chamada operacional elegível.');
  if(c?.state==='CLOSED')add('attendance-view','Visualizar chamada',null,true,'');
  N((c?.exceptionalWindowRefs||[]).length<=32&&Object.keys(c?.roster||{}).length<=200,'SCOPE_TOO_LARGE');
  const windows=(c?.exceptionalWindowRefs||[]).map(r=>({ref:r,value:get(M.path(r))})).filter(v=>v.value?.activityRefs?.some(r=>r.id===id));
  const active=windows.some(v=>v.value.status==='OPEN'&&M.ms(v.value.newDeadline)>now);
  const eligibility=f4ExceptionalWindowEligibility_(get,a,c,now);
  add('window',active?'Gerenciar prazo extraordinário':'Reabrir prazo de justificativas','ATTENDANCE_EXCEPTIONAL_WINDOW_OPEN',eligibility.eligible,eligibility.reason);
  add('history','Histórico e vínculos',null,true,'');
  return {ok:true,data:{schemaVersion:'activity-actions-v1',activityId:id,revision:a.revision,title:a.title||id,validUntil:now+30000,actions,
   extraordinary:allowed('ATTENDANCE_EXCEPTIONAL_WINDOW_OPEN')||allowed('ATTENDANCE_EXCEPTIONAL_WINDOW_CLOSE')?{
    eligible:eligibility.eligible,reason:eligibility.reason,canOpen:allowed('ATTENDANCE_EXCEPTIONAL_WINDOW_OPEN')&&eligibility.eligible&&!active,
    windows:windows.map(({ref,value:w})=>({id:ref.id,revision:w.revision,status:w.status==='OPEN'&&M.ms(w.newDeadline)<=now?'EXPIRED':w.status,deadline:M.ms(w.newDeadline),reasonCode:w.reasonCode,reason:w.reason,personIds:(w.personRefs||[]).map(r=>r.id),canClose:w.status==='OPEN'&&allowed('ATTENDANCE_EXCEPTIONAL_WINDOW_CLOSE',[ref])})),
    recipients:eligibility.eligible?Object.values(c?.roster||{}).map(r=>get(M.path(r))).filter(f=>f?.activityRef?.id===id&&f.status==='ABSENT'&&f.applicableAtDate!==false&&f.countsAbsence!==false).map(f=>({id:f.personRef.id,label:get('memberIdentities/'+f.personRef.id)?.displayName||f.personRef.id})):[]}:null}};
 });
}

// Shared by the reader and the transactional command. No global cycle defaults.
function f4ExceptionalWindowEligibility_(get,a,c,now){
 // Historical provenance may contain a NOT_RECONSTRUCTED marker, not a typed
 // cycle reference. It is never operational eligibility or authority.
 if(a.recordMode==='HISTORICAL_PARTIAL')return {eligible:false,reason:'Registro histórico somente leitura.'};
 const M=f4ActivitiesModel_(),cycle=a.cycleRef&&get(M.path(a.cycleRef)),day=M.day(now);
 const reason=a.recordMode==='HISTORICAL_PARTIAL'?'Registro histórico somente leitura.':!cycle||cycle.status!=='OPEN'||!cycle.civilStart||!cycle.civilEnd||day<cycle.civilStart||day>cycle.civilEnd?'O ciclo não está em andamento.':['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(a.status)?'A atividade não admite chamada disciplinar.':!a.typeSnapshot?.countsAbsence||!a.typeSnapshot?.allowsJustification?'A atividade não admite justificativas de falta.':!c?.lastFinalizedAt||c.state!=='CLOSED'?'A chamada precisa estar finalizada.':'';
 return {eligible:!reason,reason};
}
