// Calendar controls on the existing presentation card. The browser selects a
// destination; Core proves capacity and authorizes the resulting command.
function f4ActivitiesCalendarCompat_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),ports=options.ports,clock=options.clock||Date.now;
 function context(id){
  N(typeof id==='string'&&id.startsWith('APU_')&&M.id(id.slice(4)),'VALIDATION_ERROR');const unitRef=M.ref('unit',id.slice(4)),u=db.get(M.path(unitRef));N(u?.cycleRef,'NOT_FOUND');let auth,board=false;
  try{auth=ports.readAuthority([M.path(unitRef)],false,'PRESENTATION_BOARD_READ','PRESC_'+u.cycleRef.id);board=true}catch(e){if(!['AUTHORIZATION_DENIED','ACCESS_REVOKED','M13_AUTHZ_DENIED','AUTH_BINDING_INVALID'].includes(e.code))throw e;auth=ports.readAuthority([M.path(unitRef)],true,'PRESENTATION_SELF_READ');N(u.personRef.id===auth.personId,'AUTHORIZATION_DENIED');}
  const cycle=db.get(M.path(u.cycleRef)),cfg=cycle?.currentConfigRef&&db.get(M.path(cycle.currentConfigRef)),assignment=u.assignmentRef&&db.get(M.path(u.assignmentRef)),presentation=u.presentationRef&&db.get(M.path(u.presentationRef));N(cycle&&cfg?.approved,'CONFIG_CAPABILITY_DISABLED');
  const capacity=db.get('presentationCapacityHeads/'+u.cycleRef.id);N(capacity,'UNRESOLVED_DEPENDENCY');return {unitRef,u,cycle,cfg,assignment,presentation,auth,board,capacity};
 }
 function recheck(z,get=p=>db.get(p)){f4SharedAuthz_().recheck(get,z.auth.q,z.auth.proof,clock());for(const [ref,rev] of [[z.unitRef,z.u.revision],[z.u.cycleRef,z.cycle.revision],[z.cycle.currentConfigRef,z.cfg.revision],[M.ref('capacity',z.u.cycleRef.id),z.capacity.revision]])N(get(M.path(ref))?.revision===rev,'STALE_REVISION','CALENDAR_CONTEXT_CHANGED');}
 function freshContext(z,id,authorized){
  ports.authorizePrepared(p=>db.get(p),authorized);
  const g=db.get('presentationCapacityHeads/registry'),h=z.capacity;
  if(h.state==='SEALED'&&h.synchronizedMembersEpoch===g?.membersEpoch&&h.synchronizedCycleRegistryEpoch===g?.cycleRegistryEpoch&&M.ms(h.validUntil)>clock()+5000)return z;
  const jobs=options.jobs||[],entry=jobs.findIndex(j=>j.kind==='CAPACITY_SYNC'&&j.cycleId===z.u.cycleRef.id);
  N(entry>=0&&options.pipeline,'UNRESOLVED_DEPENDENCY','CAPACITY_REFRESH_NOT_CONFIGURED');
  f4ActivitiesConfiguredJobs_({config:options.config,db,pipeline:options.pipeline,jobs,clock}).run(entry);
  const fresh=context(id);N(fresh.u.revision===z.u.revision&&fresh.cycle.revision===z.cycle.revision&&fresh.cfg.revision===z.cfg.revision,'STALE_REVISION','CALENDAR_CONTEXT_CHANGED');return fresh;
 }
 function read(input){
  C.exact(input,['idApresentacao',...('historyCursor'in input?['historyCursor']:[])]);if(input.historyCursor){C.exact(input.historyCursor,['id','createdAt']);N(M.id(input.historyCursor.id)&&Number.isSafeInteger(input.historyCursor.createdAt),'VALIDATION_ERROR');}const z=context(input.idApresentacao),lead=z.cfg.parameters.newDateMinimumDays??30;N(Number.isSafeInteger(lead)&&lead>=30,'CONFIG_CAPABILITY_DISABLED','CALENDAR_ADVANCE');
  const window=f4ActivitiesSelfSchedulingWindow_(z.cfg.parameters,clock()),canSchedule=z.board||z.cfg.parameters.calendarMode==='MEMBER_SELF_SCHEDULE'&&window.open;const rows=[];let cursor;for(let page=0;page<2;page++){const next=db.query({collection:'presentationSessions',filters:[{field:'cycleRef',op:'EQUAL',value:z.u.cycleRef}],order:[{field:'__name__',direction:'ASCENDING'}],limit:20,...(cursor?{cursor:[cursor]}:{})});rows.push(...next);N(rows.length<=32,'SCOPE_TOO_LARGE','CALENDAR_LIST_REQUIRES_PAGING');if(next.length<20)break;cursor=new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+next[next.length-1].path);}
  const sessions=rows.map(({data:s})=>{const activity=s.activityRef&&db.get(M.path(s.activityRef)),remaining=s.config.capacity-s.slots.length,onTime=M.gap(M.day(clock()),s.civilDate)>=lead,ordinary=s.config.destination==='ORDINARY',exclusive=s.config.destination==='EXCLUSIVE'&&s.exclusiveUnitRef?.id===z.u.id;
   return {id:s.id,revision:s.revision,data:s.civilDate,titulo:activity?.title||'Sessão de apresentações',ocupadas:s.slots.length,capacidade:s.config.capacity,disponivel:!!(z.cycle.status==='OPEN'&&!z.u.assignmentRef&&!z.presentation?.realizationRef&&activity?.status==='SCHEDULED'&&activity.civilDate===s.civilDate&&remaining>0&&onTime&&(ordinary||exclusive)&&(canSchedule&&(z.board||s.slots.length===0)))};});
  const assignments=db.query({collection:'calendarAssignments',filters:[{field:'unitRef',op:'EQUAL',value:z.unitRef}],order:[{field:'createdAt',direction:'DESCENDING'},{field:'__name__',direction:'ASCENDING'}],limit:20,...(input.historyCursor?{cursor:[M.stamp(input.historyCursor.createdAt),new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+'calendarAssignments/'+input.historyCursor.id)]}:{})});
  const currentSession=z.assignment&&rows.find(r=>r.data.id===z.assignment.sessionRef.id)?.data,last=assignments[assignments.length-1]?.data;recheck(z);return {ok:true,data:{unitRevision:z.u.revision,modo:z.cfg.parameters.calendarMode,janelaAgendamento:window,antecedenciaMinima:lead,sessoes:sessions,atribuicaoAtual:z.assignment?{id:z.assignment.id,revision:z.assignment.revision,sessionId:z.assignment.sessionRef.id,status:z.assignment.status}:null,podeAgendar:!!(z.cycle.status==='OPEN'&&!z.u.assignmentRef&&!z.presentation?.realizationRef&&canSchedule),podeCancelar:!!(currentSession&&M.gap(M.day(clock()),currentSession.civilDate)>0&&z.assignment&&['ACTIVE','CONFIRMED'].includes(z.assignment.status)&&!z.presentation?.realizationRef),historico:assignments.map(({data:d})=>({id:d.id,sessionId:d.sessionRef.id,data:d.assignedCivilDate||null,status:d.status,revision:d.revision})),historicoCompleto:assignments.length<20,proximaPagina:assignments.length===20?{id:last.id,createdAt:M.ms(last.createdAt)}:null}};
 }
 function prepare(input){
  C.exact(input,['operationId','route','payload',...('basis'in input?['basis']:[])]);N(input.route==='/v2/apresentacoes/agenda','VALIDATION_ERROR');const p=input.payload;N(['AGENDAR','CANCELAR'].includes(p.acao),'VALIDATION_ERROR');C.exact(p,['idApresentacao','acao','unitRevision',...(p.acao==='AGENDAR'?['sessionId','sessionRevision']:['motivo'])]);
  let z=context(p.idApresentacao);N(z.u.revision===p.unitRevision,'STALE_REVISION','DISPLAYED_UNIT_CHANGED');N(!z.presentation?.realizationRef,'INVALID_STATE_TRANSITION','REALIZED_CALENDAR_HISTORY');let command;
  if(p.acao==='CANCELAR'){N(z.assignment&&typeof p.motivo==='string'&&p.motivo.trim()&&p.motivo.length<=2000,'INCOMPLETE_REQUEST');command={commandType:'ACT.CALENDAR_RELEASE',payload:{assignmentRef:z.u.assignmentRef,reasonCode:'PORTAL_CANCELLATION',reason:p.motivo.trim()}};z=freshContext(z,p.idApresentacao,ports.prepare({operationId:input.operationId,...command}));}
  else{
   N(z.cycle.status==='OPEN'&&!z.u.assignmentRef,'INVALID_STATE_TRANSITION');N(z.board||z.cfg.parameters.calendarMode==='MEMBER_SELF_SCHEDULE','CONFIG_CAPABILITY_DISABLED');if(!z.board)N(f4ActivitiesSelfSchedulingWindow_(z.cfg.parameters,clock()).open,'DEADLINE_VIOLATION','SELF_SCHEDULING_WINDOW_CLOSED');const sessionRef=M.ref('session',p.sessionId),session=db.get(M.path(sessionRef));N(session?.revision===p.sessionRevision&&session.cycleRef.id===z.u.cycleRef.id,'STALE_REVISION','DISPLAYED_SESSION_CHANGED');
   const intent={unitRef:z.unitRef,sessionRef,mode:z.cfg.parameters.calendarMode},placeholder=M.ref('assessment',M.hash({operationId:input.operationId,kind:'CAPACITY_AUTH_SCOPE'}).slice(0,32));
   // Authorize the exact write intent before producing its technical proof.
   let authorized=ports.prepare({operationId:input.operationId,commandType:'ACT.CALENDAR_RESERVE',payload:{...intent,assessmentRef:placeholder}});
   z=freshContext(z,p.idApresentacao,authorized);
   authorized=ports.prepare({operationId:input.operationId,commandType:'ACT.CALENDAR_RESERVE',payload:{...intent,assessmentRef:placeholder}});
   const authorize=get=>{recheck(z,get);N(get(M.path(sessionRef))?.revision===p.sessionRevision,'STALE_REVISION','DISPLAYED_SESSION_CHANGED');ports.authorizePrepared(get,authorized);return true;};
   const preparation=f4ActivitiesPreparation_({db,clock,capacitySource:f4ActivitiesCapacitySource_({db,clock,config:options.config}),authorize,executionGate:()=>authorize(path=>db.get(path))});
   const assessmentRef=preparation.prepareCapacity(null,z.u.cycleRef,z.cycle.currentConfigRef,{unitRef:z.unitRef,sessionRef});command={commandType:'ACT.CALENDAR_RESERVE',payload:{...intent,assessmentRef}};
  }
  recheck(z);const prepared=ports.prepare({operationId:input.operationId,...command});N(prepared.target.expected.some(v=>v.kind==='unit'&&v.id===z.u.id&&v.revision===p.unitRevision),'STALE_REVISION','DISPLAYED_UNIT_CHANGED');return prepared;
 }
 return {read,prepare,context,recheck,freshContext};
}
