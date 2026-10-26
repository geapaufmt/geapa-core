function f4ActivitiesReorganizationCompat_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),ports=options.ports,clock=options.clock||Date.now;
 const calendar=f4ActivitiesCalendarCompat_(options);
 function cursor(id,collection){N(M.id(id),'VALIDATION_ERROR');return [new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+collection+'/'+id)];}
 function read(input){
  C.exact(input,['idApresentacao',...['unitCursor','flowCursor'].filter(k=>k in input)]);const z=calendar.context(input.idApresentacao);
  const unitRows=z.board?db.query({collection:'presentationUnitHeads',filters:[{field:'cycleRef',op:'EQUAL',value:z.u.cycleRef}],order:[{field:'__name__',direction:'ASCENDING'}],limit:20,...(input.unitCursor?{cursor:cursor(input.unitCursor,'presentationUnitHeads')}:{})}):[];
  function unit(u){const a=u.assignmentRef&&db.get(M.path(u.assignmentRef)),s=a&&db.get(M.path(a.sessionRef)),person=db.get('memberIdentities/'+u.personRef.id);return {id:u.id,revision:u.revision,nome:person?.displayName||'Participante',sessionId:s?.id||null,data:s?.civilDate||null};}
  const units=unitRows.map(r=>unit(r.data));if(z.board&&!units.some(u=>u.id===z.u.id))units.unshift(unit(z.u));
  const rows=db.query({collection:'presentationFlowCases',filters:[{field:'cycleRef',op:'EQUAL',value:z.u.cycleRef}],order:[{field:'__name__',direction:'ASCENDING'}],limit:20,...(input.flowCursor?{cursor:cursor(input.flowCursor,'presentationFlowCases')}:{})});
  const sessions=calendar.sessionOptions(z).sessions,sessionDates=new Map(sessions.map(s=>[s.id,s.data]));
  const flows=rows.filter(r=>['CALENDAR_CHANGE','JOINT_PROPOSAL'].includes(r.data.kind)&&r.data.unitRefs.some(u=>u.id===z.u.id)).map(({data:f})=>{const affected=f.unitRefs.map(r=>{const u=db.get(M.path(r)),c=db.get('presentationFlowCases/'+f.id+'/consents/'+u.personRef.id);return {...unit(u),consentimento:c?.choice||'PENDING'};}),own=f.unitRefs.some(r=>db.get(M.path(r))?.personRef.id===z.auth.personId),consent=db.get('presentationFlowCases/'+f.id+'/consents/'+z.auth.personId),open=f.state==='OPEN'&&M.ms(f.deadline)>clock();return {id:f.id,revision:f.revision,proposalRevision:f.proposalRevision,tipo:f.kind,modo:f.kind==='JOINT_PROPOSAL'?'VOLUNTARY':f.mode,estado:!open&&f.state==='OPEN'?'EXPIRED':f.state,motivo:f.reason||f.terms,prazo:M.ms(f.deadline),afetados:affected,movimentos:(f.changeSummary||(f.moves||f.unitRefs.map(unitRef=>({unitRef,sessionRef:f.sessionRef}))).map(m=>({unitRef:m.unitRef,toSessionRef:m.sessionRef}))).map(m=>({unitId:m.unitRef.id,sessionId:m.toSessionRef?.id||null,data:m.toCivilDate||sessionDates.get(m.toSessionRef?.id)||null,dataAnterior:m.fromCivilDate||null})),podeConsentir:!!(open&&(f.mode==='VOLUNTARY'||f.kind==='JOINT_PROPOSAL')&&own&&!consent),podeConfirmar:!!(open&&z.board&&(f.kind!=='JOINT_PROPOSAL'||f.sessionRef?.kind==='session')&&(f.mode==='ADMINISTRATIVE'||affected.every(u=>u.consentimento==='ACCEPT')))};});
  calendar.recheck(z);
  return {ok:true,data:{unitRevision:z.u.revision,podePropor:z.board&&z.cycle.status==='OPEN',podeCompor:z.board&&z.cycle.status==='OPEN'&&z.cfg.parameters.shared===true,sessoes:sessions,unidades:units,propostas:flows,proximaPaginaUnidades:z.board&&unitRows.length===20?unitRows[unitRows.length-1].data.id:null,proximaPaginaPropostas:rows.length===20?rows[rows.length-1].data.id:null,janelaVoluntaria:z.cfg.parameters.reorganizationWindow||null}};
 }
 function seal(z,command,ref,evidence){
  const authorized=ports.prepare(command),same=(a,b)=>C.hash(F.encode(a||null))===C.hash(F.encode(b||null));
  db.transaction(tx=>{const get=p=>tx.get(p).data();ports.authorizePrepared(get,authorized);calendar.recheck(z,get);const prior=get(M.path(ref));if(prior)N(same(prior,evidence),'CONFLICT');else tx.set(M.path(ref),evidence);});
 }
 function prepare(input){
  C.exact(input,['operationId','route','payload',...('basis'in input?['basis']:[])]);N(input.route==='/v2/apresentacoes/reorganizacao','VALIDATION_ERROR');const p=input.payload;N(['PROPOSE','CONFIRM','CONSENT','JOINT','COMPOSE'].includes(p.action),'VALIDATION_ERROR');
  C.exact(p,['idApresentacao','unitRevision','action',...(['PROPOSE','JOINT'].includes(p.action)?['moves','mode','deadline','reason']:['CONFIRM','COMPOSE'].includes(p.action)?['flowId','flowRevision','reason']:['flowId','flowRevision','proposalRevision','choice'])]);let z=calendar.context(p.idApresentacao);N(z.u.revision===p.unitRevision,'STALE_REVISION','DISPLAYED_UNIT_CHANGED');let command;
  if(['PROPOSE','JOINT'].includes(p.action)){
   N(z.board&&Array.isArray(p.moves)&&p.moves.length>=1&&p.moves.length<=8,'AUTHORIZATION_DENIED');N(p.moves.some(m=>m.unitId===z.u.id),'VALIDATION_ERROR','FLOW_SCOPE');
   const moves=p.moves.map(v=>{C.exact(v,['unitId','unitRevision','sessionId','sessionRevision']);const unitRef=M.ref('unit',v.unitId),sessionRef=M.ref('session',v.sessionId),u=db.get(M.path(unitRef)),s=db.get(M.path(sessionRef));N(u?.revision===v.unitRevision&&s?.revision===v.sessionRevision&&u.cycleRef.id===z.u.cycleRef.id&&s.cycleRef.id===z.u.cycleRef.id,'STALE_REVISION','DISPLAYED_CALENDAR_CHANGED');return {unitRef,sessionRef};});
   if(p.action==='JOINT'){
    N(z.cfg.parameters.shared===true&&p.mode==='VOLUNTARY'&&moves.length===2&&moves[0].sessionRef.id===moves[1].sessionRef.id,'VALIDATION_ERROR','JOINT_SCOPE');
    N(Number.isSafeInteger(p.deadline)&&p.deadline>clock()&&typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'VALIDATION_ERROR');
    const ref=M.ref('assessment',M.hash({kind:'PORTAL_JOINT_DEADLINE',operationId:input.operationId}).slice(0,32)),items=moves.map(v=>({ref:v.unitRef})),evidence={id:ref.id,revision:1,kind:'PORTAL_JOINT_DEADLINE',state:'SEALED',complete:true,items,digest:M.hash(items),cycleRef:z.u.cycleRef,dueAt:M.stamp(p.deadline)};
    command={commandType:'ACT.JOINT_PROPOSE',payload:{unitRefs:moves.map(v=>v.unitRef),sessionRef:moves[0].sessionRef,terms:p.reason.trim(),deadlineRef:ref}};
    seal(z,{operationId:input.operationId,...command},ref,evidence);
   }else command={commandType:'ACT.CALENDAR_CHANGE_PROPOSE',payload:{cycleRef:z.u.cycleRef,moves:f4ActivitiesCalendarMoves_(moves),mode:p.mode,deadline:p.deadline,reason:p.reason}};
   z=calendar.freshContext(z,p.idApresentacao,ports.prepare({operationId:input.operationId,...command}));
  }else{
   const flowCaseRef=M.ref('flow',p.flowId),f=db.get(M.path(flowCaseRef));N(f?.revision===p.flowRevision&&f.cycleRef.id===z.u.cycleRef.id&&f.unitRefs.some(r=>r.id===z.u.id),'STALE_REVISION','DISPLAYED_FLOW_CHANGED');
   if(p.action==='COMPOSE'){
    N(z.board&&f.kind==='JOINT_PROPOSAL','AUTHORIZATION_DENIED');N(typeof p.reason==='string'&&p.reason.trim()&&p.reason.length<=2000,'HUMAN_DECISION_REQUIRED');
    const sealedChangeManifestRef=M.ref('assessment',M.hash({operationId:input.operationId,kind:'PORTAL_JOINT_MOVES'}).slice(0,32)),placeholder=M.ref('assessment',M.hash({operationId:input.operationId,kind:'CAPACITY_AUTH_SCOPE'}).slice(0,32)),intent={flowCaseRef,sealedChangeManifestRef,assessmentRef:placeholder};
    let authorized=ports.prepare({operationId:input.operationId,commandType:'ACT.COMPOSITION_CONFIRM',payload:intent});z=calendar.freshContext(z,p.idApresentacao,authorized);
    const units=f.unitRefs.map(ref=>({ref,u:db.get(M.path(ref))})),destination=db.get(M.path(f.sessionRef));N(destination&&units.every(v=>v.u?.assignmentRef),'UNRESOLVED_DEPENDENCY');
    const items=units.map(v=>({unitRef:v.ref,destinationSessionRef:f.sessionRef})),evidence={id:sealedChangeManifestRef.id,revision:1,kind:'PORTAL_JOINT_MOVES',state:'SEALED',complete:true,items,digest:M.hash(items),cycleRef:z.u.cycleRef,flowCaseRef,reason:p.reason.trim(),validUntil:f.deadline,basisRevisionVector:[{ref:flowCaseRef,revision:f.revision},{ref:f.sessionRef,revision:destination.revision},...units.map(v=>({ref:v.ref,revision:v.u.revision}))]};
    seal(z,{operationId:input.operationId,commandType:'ACT.COMPOSITION_CONFIRM',payload:intent},sealedChangeManifestRef,evidence);
    authorized=ports.prepare({operationId:input.operationId,commandType:'ACT.COMPOSITION_CONFIRM',payload:intent});
    const authorize=get=>{calendar.recheck(z,get);N(get(M.path(flowCaseRef))?.revision===p.flowRevision,'STALE_REVISION');ports.authorizePrepared(get,authorized);return true;},preparation=f4ActivitiesPreparation_({db,clock,capacitySource:f4ActivitiesCapacitySource_({db,clock,config:options.config}),authorize,executionGate:()=>authorize(path=>db.get(path))});
    const assessmentRef=preparation.prepareCapacity(null,z.u.cycleRef,z.cycle.currentConfigRef,{moves:f4ActivitiesCalendarMoves_(items.map(v=>({unitRef:v.unitRef,sessionRef:v.destinationSessionRef})))});command={commandType:'ACT.COMPOSITION_CONFIRM',payload:{...intent,assessmentRef}};
   }else if(p.action==='CONSENT')command={commandType:'ACT.JOINT_CONSENT',payload:{flowCaseRef,proposalRevisionRef:{number:p.proposalRevision},choice:p.choice}};
   else{N(z.board,'AUTHORIZATION_DENIED');const placeholder=M.ref('assessment',M.hash({operationId:input.operationId,kind:'CAPACITY_AUTH_SCOPE'}).slice(0,32)),intent={flowCaseRef,assessmentRef:placeholder,reason:p.reason};let authorized=ports.prepare({operationId:input.operationId,commandType:'ACT.CALENDAR_CHANGE_CONFIRM',payload:intent});z=calendar.freshContext(z,p.idApresentacao,authorized);authorized=ports.prepare({operationId:input.operationId,commandType:'ACT.CALENDAR_CHANGE_CONFIRM',payload:intent});const authorize=get=>{calendar.recheck(z,get);N(get(M.path(flowCaseRef))?.revision===p.flowRevision,'STALE_REVISION','DISPLAYED_FLOW_CHANGED');ports.authorizePrepared(get,authorized);return true;},preparation=f4ActivitiesPreparation_({db,clock,capacitySource:f4ActivitiesCapacitySource_({db,clock,config:options.config}),authorize,executionGate:()=>authorize(path=>db.get(path))}),assessmentRef=preparation.prepareCapacity(null,z.u.cycleRef,z.cycle.currentConfigRef,{moves:f4ActivitiesCalendarMoves_(f.moves)});command={commandType:'ACT.CALENDAR_CHANGE_CONFIRM',payload:{...intent,assessmentRef}};}
  }
  calendar.recheck(z);const request=ports.prepare({operationId:input.operationId,...command});
  if(['PROPOSE','JOINT'].includes(p.action))for(const v of p.moves)for(const [kind,id,rev]of [['unit',v.unitId,v.unitRevision],['session',v.sessionId,v.sessionRevision]])N(request.target.expected.some(e=>e.kind===kind&&e.id===id&&e.revision===rev),'STALE_REVISION','DISPLAYED_CALENDAR_CHANGED');
  else N(request.target.expected.some(e=>e.kind==='flow'&&e.id===p.flowId&&e.revision===p.flowRevision),'STALE_REVISION','DISPLAYED_FLOW_CHANGED');return request;
 }
 return {read,prepare};
}
