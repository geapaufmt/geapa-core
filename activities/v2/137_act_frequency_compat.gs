// Existing "Minha frequência" and justification forms, backed by canonical facts.
function f4ActivitiesFrequencyCompat_(options){
 const M=f4ActivitiesModel_(),N=M.need,C=f4Core_(),F=f4Firestore_(),db=F.connect(options.config),ports=options.ports,clock=options.clock||Date.now,cache=new Map();let temporalValidUntil=Infinity,semesters=[];
 function boundary(t){if(t){const ms=M.ms(t);if(ms>clock())temporalValidUntil=Math.min(temporalValidUntil,ms);}}
 const get=p=>{if(!cache.has(p))cache.set(p,db.get(p));return cache.get(p)},iso=t=>t?new Date(M.ms(t)).toISOString():'';
 function prefetch(paths){const missing=[...new Set(paths)].filter(p=>!cache.has(p));for(let i=0;i<missing.length;i+=180){const batch=missing.slice(i,i+180),values=db.getMany(batch);batch.forEach((p,j)=>cache.set(p,values[j]));}}
 function hydrate(facts,admin){semesters=facts.length?all('academicSemesters'):[];semesters.forEach(s=>cache.set('academicSemesters/'+s.id,s));prefetch(facts.flatMap(f=>[M.path(f.activityRef),'activityAttendanceCollections/'+f.activityRef.id,...(f.periodId?['activityFrequencyPeriods/'+f.periodId]:[]),...(f.cycleId?['presentationCycles/'+f.cycleId]:[]),...((f.pendingJustificationRef||f.lastJustificationRef||f.justificationDecisionRef)?[M.path(f.pendingJustificationRef||f.lastJustificationRef||f.justificationDecisionRef)]:[]),...(admin?['memberAggregateHeads/'+f.personRef.id,'memberIdentities/'+f.personRef.id,'memberPortalProfiles/'+f.personRef.id]:[])]));prefetch(facts.flatMap(f=>get('activityAttendanceCollections/'+f.activityRef.id)?.exceptionalWindowRefs||[]).map(M.path));prefetch(facts.flatMap(f=>{const a=get(M.path(f.activityRef)),c=get('activityAttendanceCollections/'+f.activityRef.id);return [...(a?.provenanceHead?.id?['migrationEvidence/'+a.provenanceHead.id]:[]),...(c?.periodRef?[M.path(c.periodRef)]:[])];}));}
 function all(collection,filters=[]){let cursor,rows=[];for(let page=0;page<50;page++){const next=db.query({collection,filters,limit:20,order:[{field:'__name__',direction:'ASCENDING'}],...(cursor?{cursor:[cursor]}:{})});rows.push(...next.map(r=>r.data));if(next.length<20)return rows;cursor=new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+collection+'/'+next[next.length-1].data.id);}N(false,'SCOPE_TOO_LARGE','FREQUENCY_REQUIRES_PAGING');}
 const justStatus=s=>({PREVIOUS:'PREVIA',PENDING:'ENVIADA',CORRECTION_REQUESTED:'AJUSTE_SOLICITADO',APPROVED:'DEFERIDA',EXCUSED:'ABONADA',REJECTED:'INDEFERIDA',CLOSED_WITHOUT_ABSENCE:'ENCERRADA_SEM_FALTA',HISTORICAL_PARTIAL:'HISTORICO_INCOMPLETO'}[s]||'');
 function row(fact,admin=false,selectedJustification){
  const activity=get(M.path(fact.activityRef)),collection=get('activityAttendanceCollections/'+fact.activityRef.id);N(activity&&collection,'UNRESOLVED_DEPENDENCY');
  const j=selectedJustification||((fact.pendingJustificationRef||fact.lastJustificationRef||fact.justificationDecisionRef)&&get(M.path(fact.pendingJustificationRef||fact.lastJustificationRef||fact.justificationDecisionRef))),now=clock(),late=!!collection.justificationEndsAt&&now>M.ms(collection.justificationEndsAt),before=!!activity.startsAt&&now<M.ms(activity.startsAt);
  // Submission eligibility follows the pending reference used by the command.
  // A terminal historical decision remains visible without blocking a new request.
  const pending=fact.pendingJustificationRef&&get(M.path(fact.pendingJustificationRef)),open=!fact.pendingJustificationRef||!!pending&&(pending.state==='CORRECTION_REQUESTED'||pending.state==='PREVIOUS'&&before);
  const exceptional=late?f4AttendanceExceptionalWindow_(get,collection,fact,activity,now):null;
  const outsideDeadline=late&&!exceptional&&!j?.submittedUnderExceptionalWindow;
  [activity.startsAt,collection.justificationOpensAt,collection.justificationEndsAt,exceptional?.value.newDeadline].forEach(boundary);
  const canRequest=!admin&&ports.commandAccess('ATTENDANCE_JUSTIFICATION_REQUEST',[M.ref('attendance',fact.id),M.ref('activity',activity.id)],'SELF')&&activity.recordMode!=='HISTORICAL_PARTIAL'&&activity.typeSnapshot?.allowsJustification&&fact.applicableAtDate!==false&&!['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(activity.status)&&open&&(before?collection.justificationOpensAt&&now>=M.ms(collection.justificationOpensAt):['ABSENT','UNRECORDED'].includes(fact.status))&&(!late||exceptional||collection.acceptLateJustification);
  const status={PRESENT:fact.modality==='REMOTA'?'PRESENTE_REMOTO':'PRESENTE_PRESENCIAL',ABSENT:'FALTA',JUSTIFIED:'JUSTIFICADA',EXCUSED:'ABONADA',NOT_APPLICABLE:'NAO_SE_APLICA',UNRECORDED:'SEM_REGISTRO'}[fact.status];
  const calendar=f4ActivityAcademicMetadata_(activity,get),cycle=f4ActivityFrequencyCycle_(activity,collection,fact,get,semesters);
  const dto={...calendar,...cycle,idRegistroPresenca:fact.id,idAtividade:activity.id,tituloAtividade:activity.title,tituloPublico:activity.title,dataAtividade:activity.civilDate,statusPresenca:status,statusPresencaRotulo:{PRESENTE_REMOTO:'Presente remoto',PRESENTE_PRESENCIAL:'Presente presencial',FALTA:'Falta',JUSTIFICADA:'Justificada',ABONADA:'Abonada',NAO_SE_APLICA:'Não se aplica',SEM_REGISTRO:'Sem registro'}[status],ciclo:cycle.cycleId||'',...(exceptional?{exceptionalWindowId:exceptional.ref.id,prazoExtraordinario:iso(exceptional.value.newDeadline),mensagemPrazoExtraordinario:'Prazo extraordinário para justificativa aberto até '+Utilities.formatDate(new Date(M.ms(exceptional.value.newDeadline)),'America/Cuiaba','dd/MM/yyyy HH:mm')}:{}) ,dataLimiteJustificativa:iso(collection.justificationEndsAt),prazoJustificativa:iso(collection.justificationEndsAt),idJustificativa:j?.id||'',statusJustificativa:justStatus(j?.state),justificativaPreviaPendente:j?.state==='PREVIOUS',foraPrazo:outsideDeadline,statusPrazo:outsideDeadline?'FORA_DO_PRAZO':'NO_PRAZO',envioForaDoPrazo:outsideDeadline?'SIM':'NAO',exigeCienciaForaPrazo:outsideDeadline,podeEnviarJustificativa:!!canRequest,podeComplementarJustificativa:!!canRequest&&j?.state==='CORRECTION_REQUESTED',podeVerJustificativa:!!j,acaoJustificativa:!canRequest?'':j?.state==='CORRECTION_REQUESTED'?'COMPLEMENTAR_JUSTIFICATIVA':outsideDeadline?'ENVIAR_JUSTIFICATIVA_FORA_PRAZO':'ENVIAR_JUSTIFICATIVA',motivoDeclarado:j?.reasonCode||'',descricaoJustificativa:j?.reason||'',observacaoPublica:j?.decisionReason||'',linkDocumentoComprobatorio:j?.proofLink||'',possuiDocumentoComprobatorio:j?.proofLink||j?.proofArtifactRef?'SIM':'NAO',canonicalProofAvailable:j?.proofArtifactRef?.kind==='activityArtifact',enviadaEm:iso(j?.submittedAt)};
  if(admin){const h=get('memberAggregateHeads/'+fact.personRef.id),person=get('memberIdentities/'+fact.personRef.id),profile=get('memberPortalProfiles/'+fact.personRef.id);dto.nomePessoa=profile&&h&&profile.sourceRevision===h.revision?profile.displayLabel:person?.displayName||'';dto.idPessoa=fact.personRef.id;dto.observacoesInternas=j?.internalReason||'';const readiness=f4JustificationDecisionReadiness_(j,fact,activity),authorized=ports.commandAccess('ATTENDANCE_JUSTIFICATION_DECIDE',[M.ref('justification',j.id)]),pending=readiness.eligible&&authorized;dto.impedimentoAnalise=readiness.eligible&&!authorized?'Sua autorização não permite decidir esta justificativa.':readiness.reason;dto.fila=queueGroup(j,fact,activity);dto.acoesGestao={podeDeferir:pending,podeAbonar:pending&&['EXCLUDE','COUNT_PRESENT','COUNT_ABSENT'].includes(fact.frequencyBasis?.excusedMode),podeIndeferir:pending,podeSolicitarAjuste:pending};if(!authorized){for(const field of ['observacoesInternas','motivoDeclarado','descricaoJustificativa','observacaoPublica','linkDocumentoComprobatorio'])delete dto[field];dto.possuiDocumentoComprobatorio='NAO';dto.canonicalProofAvailable=false;}}
  const countable=(fact.recordMode==='HISTORICAL_PARTIAL'&&fact.historicalAttendanceFact===true||!!collection.lastFinalizedAt&&!!fact.officialRecordedAt)&&activity.legacyStatus!=='CANCELADA'&&!['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(activity.status);
  return {dto,derived:{...fact,countable,cycleId:cycle.cycleId,activityCivilDate:activity.civilDate,activityTime:activity.startsAt?Utilities.formatDate(new Date(M.ms(activity.startsAt)),'America/Cuiaba','HH:mm'):'',justificationApproved:fact.justificationApproved===true||j?.state==='APPROVED'||fact.recordMode==='HISTORICAL_PARTIAL'&&['JUSTIFIED','EXCUSED'].includes(fact.status)}};
 }
 function pair(items){const rows=f4AttendancePairJustifications_(items.map(r=>r.derived),true);return items.map((r,i)=>{const derived=rows[i],status=derived.status;return {...r,derived,dto:{...r.dto,...(derived.justificationPairingRule&&['JUSTIFIED','EXCUSED'].includes(status)?{statusPresenca:status==='EXCUSED'?'ABONADA':'JUSTIFICADA',statusPresencaRotulo:status==='EXCUSED'?'Abonada':'Justificada',pairingBasis:derived.pairingBasis}: {})}};});}
 function summary(rows){
  const units=new Set(rows.filter(r=>r.countable&&r.status!=='NOT_APPLICABLE').map(r=>r.frequencyBasis?.unit).filter(Boolean));
  const value=f4AttendanceFrequency_(rows,units.size<=1),complete=value.basisComplete!==false&&units.size<=1;
  const net=value.counts.ABSENT+rows.filter(r=>r.countable&&(r.officialCountBasis||r).countsAbsence!==false&&(r.officialCountBasis||r).applicableAtDate!==false&&(r.status==='JUSTIFIED'&&r.frequencyBasis?.justifiedMode==='COUNT_ABSENT'||r.status==='EXCUSED'&&r.frequencyBasis?.excusedMode==='COUNT_ABSENT')).length;
  return {totalRegistros:rows.length,totalAtividades:['PRESENT','ABSENT','JUSTIFIED','EXCUSED'].reduce((n,k)=>n+value.counts[k],0),totalPresencas:value.counts.PRESENT,totalFaltas:value.counts.ABSENT,totalJustificadas:value.counts.JUSTIFIED,totalAbonadas:value.counts.EXCUSED,totalNaoAplicaveis:value.counts.NOT_APPLICABLE,totalPendentes:value.pending,faltasLiquidas:complete?net:null,percentualFrequencia:value.percentage===null?null:Math.round(value.percentage*100)/100,baseCompleta:complete,mensagemPortal:complete?'':'Totais dos fatos disponíveis. A base histórica ou os pesos não permitem apurar faltas líquidas e percentual completo com segurança.'};
 }
 function population(cycle,heads){
  const result=new Map();
  heads.forEach(h=>{N(h.kind==='PERSON'&&Array.isArray(h.paths),'UNRESOLVED_DEPENDENCY','MEMBERS_HEAD');cache.set('memberAggregateHeads/'+h.id,h);});
  prefetch(heads.flatMap(h=>h.paths.filter(p=>/^member(Identities|Links|Events)\//.test(p))));
  const model=f4MemberModel_(true,get('memberAuthorityControl/MEMBERS')||undefined);
  const start=C.civilStart(cycle.civilStart),end=C.civilStart(cycle.civilEnd)+86400000;
  for(const h of heads){
   N(h.kind==='PERSON'&&Array.isArray(h.paths),'UNRESOLVED_DEPENDENCY','MEMBERS_HEAD');
   const paths=h.paths.filter(p=>/^member(Identities|Links|Events)\//.test(p));N(paths.length<=180,'SCOPE_TOO_LARGE','MEMBERS_SOURCES');prefetch(paths);
   const docs=Object.fromEntries(paths.map(p=>[p,get(p)])),identity=docs['memberIdentities/'+h.id];
   if(identity?.technicalAccount)continue;
   if(!identity||Object.values(docs).some(d=>!d||d.idPessoa!==h.id)){result.set(h.id,{complete:false,periods:[]});continue;}
   for(const d of Object.values(docs))if(d.type==='RETIFICACAO'&&d.correctionPath?.startsWith('memberFunctions/')&&['m1-memberFunctions-v2','m1-memberFunctions-v3'].includes(d.prior?.schemaVersion)){const p='memberRoleCatalog/'+d.prior.idCargo;docs[p]=get(p);}
   // Canonical derivation at every change boundary preserves ingress, suspension,
   // dismissal and corrections without using today's membership projection.
   const points=new Set([start,end]);for(const d of Object.values(docs)){for(const k of ['startDate','effectiveDate','endDate','exitDate'])if(d[k]){const at=C.civilStart(d[k])+(k==='endDate'?86400000:0);if(start<at&&at<end)points.add(at);}}
   const sorted=[...points].sort((a,b)=>a-b),periods=[];
   for(let i=0;i<sorted.length-1;i++){const state=model.derive(docs,h.id,h.revision,sorted[i]).state.state;if(['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE','SUSPENSO'].includes(state))periods.push({start:sorted[i],end:sorted[i+1],state});}
   if(periods.length)result.set(h.id,{complete:true,personRevision:h.revision,periods});
  }
  return result;
 }
 // Administrative cycle view reuses the same factual rows, pairing and count
 // basis as SELF. It never derives membership sanctions from attendance.
 function central(params,check){
  C.exact(params,['cycleId',...('personId'in params?['personId']:[]),...('cursor'in params?['cursor']:[])]);
  N(M.id(params.cycleId)&&(!params.personId||M.id(params.personId)),'VALIDATION_ERROR','SELECT_CANONICAL_CYCLE');
  const cycle=get('presentationCycles/'+params.cycleId);N(cycle,'NOT_FOUND');
  const proof=ports.readAuthority(['presentationCycles/'+cycle.id],false);check([proof]);
  if(params.cursor){C.exact(params.cursor,['afterId','cycleId']);N(!params.personId&&M.id(params.cursor.afterId)&&params.cursor.cycleId===cycle.id,'VALIDATION_ERROR','RH_CURSOR_CYCLE_MISMATCH');}
  const after=params.cursor?.afterId,heads=params.personId?[get('memberAggregateHeads/'+params.personId)].filter(Boolean):db.query({collection:'memberAggregateHeads',filters:[],limit:20,order:[{field:'__name__',direction:'ASCENDING'}],...(after?{cursor:[new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+'memberAggregateHeads/'+after)]}:{})}).map(r=>r.data);
  const pagination={pageSize:20,hasMore:!params.personId&&heads.length===20,nextCursor:!params.personId&&heads.length===20?{cycleId:cycle.id,afterId:heads[heads.length-1].id}:null};
  const people=population(cycle,heads);
  const cycles=all('presentationCycles');cycles.forEach(c=>cache.set('presentationCycles/'+c.id,c));semesters=all('academicSemesters');
  semesters.forEach(s=>cache.set('academicSemesters/'+s.id,s));
  const activities=all('canonicalActivities').filter(a=>{const id=a.cycleId||a.cycleRef?.id||a.cycleRef?.ref?.id;return !id||id===cycle.id;}),authorized=[],missing=[];
  prefetch(activities.map(a=>'activityAttendanceCollections/'+a.id));
  for(const a of activities){
   cache.set('canonicalActivities/'+a.id,a);
   const collection=get('activityAttendanceCollections/'+a.id);
   const resolved=f4ActivityFrequencyCycle_(a,collection||{},{},get,semesters);
   if(resolved.cycleId!==cycle.id)continue;
   const access=ports.readAuthority(['canonicalActivities/'+a.id],false);check([access]);
   if(!['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(a.status)&&a.typeSnapshot?.countsAttendance&&a.civilDate<=Utilities.formatDate(new Date(clock()),'America/Cuiaba','yyyy-MM-dd')&&!collection?.lastFinalizedAt)missing.push(a.id);
   authorized.push(a);
  }
  const activityIds=new Set(authorized.map(a=>a.id));
  const facts=heads.flatMap(h=>all('activityAttendance',[{field:'ownerPersonRef',op:'EQUAL',value:h.id}])).filter(f=>activityIds.has(f.activityRef?.id));
  N(facts.every(f=>f.personRef?.id&&f.ownerPersonRef===f.personRef.id),'UNRESOLVED_DEPENDENCY','ATTENDANCE_OWNER_CONFLICT');
  hydrate(facts,true);
  // Resolve and filter BEFORE pairing: another cycle cannot consume a pair.
  const rows=pair(facts.map(f=>row(f)).filter(r=>r.derived.cycleId===cycle.id));
  const periods=(cycle.frequencyPlanning?.periods||[]).filter(p=>p.state==='APPROVED'),discipline=f4MemberAttendanceCases_().rules(get,cycle.id);
  prefetch(heads.map(h=>'memberPortalProfiles/'+h.id));
  const members=[...new Set([...people.keys(),...rows.map(r=>r.derived.personRef.id)])].sort().filter(id=>!params.personId||id===params.personId).map(id=>{
   const own=rows.filter(r=>r.derived.personRef.id===id),s=summary(own.map(r=>r.derived)),head=get('memberAggregateHeads/'+id),profile=get('memberPortalProfiles/'+id),identity=get('memberIdentities/'+id);
   const person=people.get(id),missingFacts=authorized.filter(a=>{if(['CANCELLED','NOT_OCCURRED','ARCHIVED'].includes(a.status)||!a.typeSnapshot?.countsAttendance||!a.civilDate||a.civilDate>Utilities.formatDate(new Date(clock()),'America/Cuiaba','yyyy-MM-dd'))return false;const at=a.startsAt?M.ms(a.startsAt):C.civilStart(a.civilDate);return person?.periods.some(p=>(p.state==='MEMBRO_EFETIVO'||p.state==='MEMBRO_INGRESSANTE'&&a.typeSnapshot.participantTypes?.includes('INGRESSANTE'))&&p.start<=at&&at<p.end)&&!own.some(r=>r.derived.activityRef.id===a.id);}).map(a=>a.id);
   const integrity=!person?.complete||!s.baseCompleta||missingFacts.length||missing.length||!periods.length||own.some(r=>r.derived.recordMode==='HISTORICAL_PARTIAL')?'INCOMPLETE':cycle.status!=='CLOSED'||s.totalPendentes||own.some(r=>r.derived.pendingJustificationRef)?'PROVISIONAL':'DEFINITIVE';
   const effectivePeriods=person?.periods.filter(v=>v.state==='MEMBRO_EFETIVO')||[];
   const plans=periods.filter(p=>effectivePeriods.some(v=>v.start<(C.civilStart(p.civilEnd)+86400000)&&C.civilStart(p.civilStart)<v.end)).map(p=>{const scoped=own.filter(r=>r.dto.dataAtividade>=p.civilStart&&r.dto.dataAtividade<=p.civilEnd),count=summary(scoped.map(r=>r.derived));return {periodId:p.periodId,civilStart:p.civilStart,civilEnd:p.civilEnd,plannedActivities:p.attendanceActivitiesCount,absenceLimit:p.absenceLimit,minimumAttendanceBasisPoints:p.rules?.minimumAttendanceBasisPoints??null,absenceNorm:p.rules?.absenceNorm||'',attendanceNorm:p.rules?.attendanceNorm||'',discipline:discipline.established?{normativeReference:discipline.rules.normativeReference,informalWarningAt:typeof p.absenceLimit==='number'?Math.ceil(p.absenceLimit*discipline.rules.informalBasisPoints/10000):null,formalWarningAt:typeof p.absenceLimit==='number'?Math.ceil(p.absenceLimit*discipline.rules.formalBasisPoints/10000):null,manifestationDays:discipline.rules.manifestationDays}:null,netAbsences:count.faltasLiquidas,atLimit:typeof p.absenceLimit==='number'&&p.absenceLimit>=0&&count.faltasLiquidas!==null?count.faltasLiquidas>=p.absenceLimit:null,decisionRequired:true};});
   return {personId:id,personRevision:head?.revision||null,label:profile&&head&&profile.sourceRevision===head.revision?profile.displayLabel:identity?.displayName||id,summary:s,integrity,participation:person?.periods.map(p=>({start:new Date(p.start).toISOString(),endExclusive:new Date(p.end).toISOString(),state:p.state}))||[],missingActivityIds:missingFacts,periods:plans,records:own.map(r=>({activityId:r.dto.idAtividade,title:r.dto.tituloAtividade,civilDate:r.dto.dataAtividade,status:r.dto.statusPresenca,label:r.dto.statusPresencaRotulo,cycleId:cycle.id,pending:!!r.derived.pendingJustificationRef}))};
  });
  ports.recheckAuthority(proof);
  const integrity=missing.length||!periods.length||members.some(m=>m.integrity==='INCOMPLETE')?'INCOMPLETE':cycle.status==='CLOSED'&&members.every(m=>m.integrity==='DEFINITIVE')?'DEFINITIVE':'PROVISIONAL';
  return {schemaVersion:'attendance-rh-cycle-v1',pagination,scope:'SELECTED_CYCLE_MEMBER_PAGE',cycle:{id:cycle.id,label:cycle.cycleLabel||cycle.id,status:cycle.status,revision:cycle.revision},cycles:cycles.map(c=>({id:c.id,label:c.cycleLabel||c.id,status:c.status})),activities:authorized.map(a=>({id:a.id,revision:a.revision,title:a.title,status:a.status,civilDate:a.civilDate,recordedAt:iso(a.recordedAt),replacementOfId:a.replacementOfRef?.id||null,replacedById:a.replacedByRef?.id||null})),canReplace:ports.commandAvailable?.('ACTIVITY_REPLACE')===true,members,missingAttendanceActivityIds:missing,coverage:'CANONICAL_CYCLE_PARTICIPANTS',populationComplete:[...people.values()].every(p=>p.complete),integrity,notice:integrity==='INCOMPLETE'?'Há fontes ou chamadas pendentes de conferência. Ausência de registro não equivale a falta.':integrity==='PROVISIONAL'?'O ciclo está em andamento ou há decisões pendentes. Os totais podem mudar.':'Apuração encerrada com fontes completas e parâmetros do próprio ciclo.',automaticDismissal:false};
 }
 // Filter by canonical cycle before pagination, never by submission date.
 // Cursor binds the selected cycle, queue and attendance epoch.
 function queueGroup(j,f,a){return ['APPROVED','EXCUSED','REJECTED','CLOSED_WITHOUT_ABSENCE'].includes(j.state)?'HISTORY':f4JustificationDecisionReadiness_(j,f,a).eligible?'ANALYSIS':'ACTION_REQUIRED';}
 function pendingPage(params,epoch,check){
  C.exact(params,['cursor','cycleId','queue'].filter(k=>Object.hasOwn(params,k)));
  const cycles=all('presentationCycles');cycles.forEach(c=>cache.set('presentationCycles/'+c.id,c));
  const current=cycles.filter(c=>c.status==='OPEN'),cycleId=params.cycleId||(current.length===1?current[0].id:null),queue=params.queue||'ALL';
  N(['ALL','ANALYSIS','ACTION_REQUIRED','HISTORY'].includes(queue),'VALIDATION_ERROR');
  // Do not choose an institutional cycle arbitrarily. Return the selector before
  // scanning private requests; the enclosing reader still authorizes/rechecks.
  if(!cycleId&&!Object.hasOwn(params,'cycleId')&&!params.cursor)return {items:[],cycleId:null,queue,requiresCycleSelection:true,cycles:cycles.map(c=>({id:c.id,label:c.cycleLabel||c.id,status:c.status})),pagination:{pageSize:10,hasMore:false,nextCursor:null}};
  N(M.id(cycleId)&&cycles.some(c=>c.id===cycleId),'VALIDATION_ERROR','SELECT_CANONICAL_CYCLE');
  const cursor=params.cursor;if(cursor){C.exact(cursor,['afterId','epoch','cycleId','queue']);N(M.id(cursor.afterId)&&cursor.epoch===epoch&&cursor.cycleId===cycleId&&cursor.queue===queue,'STALE_REVISION','JUSTIFICATION_PAGE_CHANGED');}
  semesters=all('academicSemesters');semesters.forEach(s=>cache.set('academicSemesters/'+s.id,s));
  const states=queue==='HISTORY'?['APPROVED','EXCUSED','REJECTED','CLOSED_WITHOUT_ABSENCE']:queue==='ALL'?['PREVIOUS','PENDING','CORRECTION_REQUESTED','HISTORICAL_PARTIAL','APPROVED','EXCUSED','REJECTED','CLOSED_WITHOUT_ABSENCE']:['PREVIOUS','PENDING','CORRECTION_REQUESTED','HISTORICAL_PARTIAL'];
  let afterId=cursor?.afterId,items=[],done=false;const authorizations=new Map();
  for(let scan=0;scan<50&&!done&&items.length<11;scan++){
   const after=afterId?new F.DocumentReference('projects/'+options.config.projectId+'/databases/(default)/documents/'+(f4DevBoundary_().root?f4DevBoundary_().root+'/':'')+'activityJustifications/'+afterId):null;
   const merged=states.flatMap(state=>db.query({collection:'activityJustifications',filters:[{field:'state',op:'EQUAL',value:state}],limit:20,order:[{field:'__name__',direction:'ASCENDING'}],...(after?{cursor:[after]}:{})}).map(r=>r.data)).sort((a,b)=>a.id<b.id?-1:a.id>b.id?1:0),batch=merged.slice(0,20);
   done=merged.length<20;
   prefetch(batch.flatMap(j=>[M.path(j.activityRef),...(j.attendanceRef?[M.path(j.attendanceRef)]:[]),'activityAttendanceCollections/'+j.activityRef.id]));
   for(const j of batch){
    afterId=j.id;const a=get(M.path(j.activityRef)),f=(j.attendanceRef?get(M.path(j.attendanceRef)):null),c=get('activityAttendanceCollections/'+j.activityRef.id);
    if(!a)continue;
    const cycle=f4ActivityFrequencyCycle_(a,c||{},f||{},get,semesters);if(cycle.cycleId!==cycleId)continue;
    if(!authorizations.has(a.id)){try{const proof=ports.readAuthority([M.path(j.activityRef)],false);check([proof]);authorizations.set(a.id,true);}catch(e){if(!['AUTHORIZATION_DENIED','ACCESS_REVOKED','M13_AUTHZ_DENIED'].includes(e.code))throw e;authorizations.set(a.id,false);}}
    if(!authorizations.get(a.id))continue;
    if(queue!=='ALL'&&queueGroup(j,f,a)!==queue)continue;
    cache.set('activityJustifications/'+j.id,j);items.push(j);if(items.length===11)break;
   }
  }
  N(done||items.length===11,'SCOPE_TOO_LARGE','JUSTIFICATION_SCAN_LIMIT');
  const page=items.slice(0,10);return {items:page,cycleId,queue,cycles:cycles.map(c=>({id:c.id,label:c.cycleLabel||c.id,status:c.status})),pagination:{pageSize:10,hasMore:items.length>10,nextCursor:items.length>10?{afterId:page[page.length-1].id,epoch,cycleId,queue}:null}};
 }
 function read(route,params={}){return ports.readWithAuthoritySnapshot(check=>readSnapshot(route,params,check));}
 function readSnapshot(route,params={},check){
  if(!['/v2/justificativas/pendencias','/admin/frequencia-rh'].includes(route))C.exact(params,[]);
  cache.clear();temporalValidUntil=Infinity;
  const admin=['/v2/justificativas/pendencias','/admin/frequencia-rh'].includes(route),auth=ports.readAuthority(admin?['activityAttendanceRegistry/current']:[],!admin),epoch=get('activityAttendanceRegistry/current')?.epoch||0;let data;check([auth]);
  if(route==='/admin/frequencia-rh'){data=central(params,check);}
  else if(route==='/v2/minha-frequencia'){
   const values=all('activityAttendance',[{field:'ownerPersonRef',op:'EQUAL',value:auth.personId}]);N(values.every(f=>f.ownerPersonRef===auth.personId&&f.personRef.id===auth.personId),'AUTHORIZATION_DENIED');hydrate(values,false);const items=pair(values.map(f=>row(f))),periods=[...new Set(items.map(r=>r.dto.ciclo).filter(Boolean))].sort().reverse();
   data={contrato:'MINHA_FREQUENCIA_DETALHADA_V2',cycleGrouping:'GEAPA_CYCLE_V1',registros:items.filter(r=>r.dto.ciclo).map(r=>r.dto),registrosPendentes:items.filter(r=>!r.dto.ciclo).map(r=>r.dto),ciclos:periods.map(id=>({ciclo:id,cycleId:id,rotuloCiclo:get('presentationCycles/'+id)?.cycleLabel||id,resumo:summary(items.filter(r=>r.dto.ciclo===id).map(r=>r.derived)),registros:items.filter(r=>r.dto.ciclo===id).map(r=>r.dto)})),ultimaAtualizacao:new Date(clock()).toISOString()};
  }else if(admin){
   const page=pendingPage(params,epoch,check),entries=page.items,facts=entries.map(j=>(j.attendanceRef?get(M.path(j.attendanceRef)):null)).filter(Boolean);hydrate(facts,true);
   const rows=entries.map(j=>{const f=(j.attendanceRef?get(M.path(j.attendanceRef)):null),a=get(M.path(j.activityRef)),c=get('activityAttendanceCollections/'+j.activityRef.id);
    if(f&&c)return row(f,true,j).dto;
    // No synthetic attendance is manufactured for an orphaned imported request.
    return {idJustificativa:j.id,idAtividade:j.activityRef.id,tituloAtividade:a?.title||j.activityRef.id,idPessoa:j.personRef?.id||'',ciclo:page.cycleId,statusJustificativa:justStatus(j.state)||'HISTORICO_INCOMPLETO',enviadaEm:iso(j.submittedAt),fila:queueGroup(j,f,a),impedimentoAnalise:'Referência de frequência ou chamada ausente; requer reconciliação documental.',acoesGestao:{podeDeferir:false,podeAbonar:false,podeIndeferir:false,podeSolicitarAjuste:false}};
   });data={justificativas:rows,pagination:page.pagination,cycleId:page.cycleId,queue:page.queue,cycles:page.cycles,requiresCycleSelection:page.requiresCycleSelection===true,ultimaAtualizacao:new Date(clock()).toISOString()};}

  else{N(route==='/v2/justificativas/config','VALIDATION_ERROR');data={motivos:[{valor:'SAUDE',rotulo:'Saúde'},{valor:'COMPROMISSO_ACADEMICO',rotulo:'Compromisso acadêmico'},{valor:'COMPROMISSO_PROFISSIONAL',rotulo:'Compromisso profissional'},{valor:'MOTIVO_PESSOAL_RELEVANTE',rotulo:'Motivo pessoal relevante'},{valor:'FORCA_MAIOR',rotulo:'Força maior'},{valor:'OUTRO',rotulo:'Outro'}]};}
  ports.recheckAuthority(auth);N((db.get('activityAttendanceRegistry/current')?.epoch||0)===epoch,'STALE_REVISION');return {ok:true,data};
 }
 function prepare(input){C.exact(input,['operationId','route','payload']);const p=input.payload;let command;
  if(input.route==='/v2/justificativas/enviar'){
   C.exact(p,['idRegistroPresenca','idAtividade','motivoDeclarado','descricaoJustificativa','possuiDocumentoComprobatorio','linkDocumentoComprobatorio','confirmouCienciaForaPrazo','observacoes',...('documentoComprobatorio'in p?['documentoComprobatorio']:[]),...('proofUploadRef'in p?['proofUploadRef']:[])]);
   N(!p.documentoComprobatorio,'UNRESOLVED_DEPENDENCY','JUSTIFICATION_UPLOAD_PROVIDER_REQUIRED');N(M.id(p.idAtividade)&&['SIM','NAO'].includes(p.possuiDocumentoComprobatorio)&&typeof p.confirmouCienciaForaPrazo==='boolean','VALIDATION_ERROR');
   const auth=ports.readAuthority([],true),derivedId='a_'+M.hash({activity:p.idAtividade,person:auth.personId}).slice(0,32),id=p.idRegistroPresenca||derivedId;N(M.id(id),'VALIDATION_ERROR');
   const fact=get('activityAttendance/'+id);
   if(id!==derivedId){
    // Imported facts retain their IDs. Only an existing, canonically recognized
    // SELF fact may use that path; absent facts still use the deterministic ID.
    const activity=get('canonicalActivities/'+p.idAtividade),recognition=fact?.importRecognition;
    N(fact?.id===id&&fact.recordMode==='NEW_V2'&&fact.personRef?.id===auth.personId&&fact.ownerPersonRef===auth.personId&&fact.activityRef?.kind==='activity'&&fact.activityRef.id===p.idAtividade&&recognition?.schemaVersion==='act-import-recognition-v1'&&/^[a-f0-9]{32}$/.test(recognition.operationId||'')&&activity?.recordMode==='NEW_V2'&&activity.importRecognition?.schemaVersion==='act-import-recognition-v1'&&recognition.operationId===activity.importRecognition.operationId,'AUTHORIZATION_DENIED');
   }
   N(fact?.recordMode!=='HISTORICAL_PARTIAL','UNRESOLVED_DEPENDENCY','HISTORICAL_RECORD_REQUIRES_RECONCILIATION');const prior=fact?.pendingJustificationRef&&get(M.path(fact.pendingJustificationRef));N(p.possuiDocumentoComprobatorio!=='SIM'||p.linkDocumentoComprobatorio||p.proofUploadRef||prior?.personRef?.id===auth.personId&&prior.proofArtifactRef?.kind==='activityArtifact','INCOMPLETE_REQUEST');
   N(typeof p.descricaoJustificativa==='string'&&typeof p.observacoes==='string','VALIDATION_ERROR');
   command={commandType:'ACT.ATTENDANCE_JUSTIFICATION_REQUEST',payload:{attendanceRef:M.ref('attendance',id),activityRef:M.ref('activity',p.idAtividade),reason:[p.descricaoJustificativa,p.observacoes].filter(Boolean).join('\n'),reasonCode:p.motivoDeclarado,proofLink:p.linkDocumentoComprobatorio||'',clearProof:p.possuiDocumentoComprobatorio==='NAO'&&!p.proofUploadRef&&!p.linkDocumentoComprobatorio,...(p.proofUploadRef?{proofUploadRef:p.proofUploadRef}:{}),lateAcknowledged:p.confirmouCienciaForaPrazo}};
  }else{
   N(input.route==='/v2/justificativas/analisar','VALIDATION_ERROR');N(get('activityJustifications/'+p.idJustificativa)?.recordMode!=='HISTORICAL_PARTIAL','UNRESOLVED_DEPENDENCY','HISTORICAL_RECORD_REQUIRES_RECONCILIATION');C.exact(p,['idJustificativa','decisao','observacaoPublica','observacoesInternas']);const decision={DEFERIR:'APPROVE',ABONAR:'EXCUSE',INDEFERIR:'REJECT',SOLICITAR_AJUSTE:'REQUEST_CORRECTION'}[p.decisao];N(decision&&typeof p.observacaoPublica==='string'&&typeof p.observacoesInternas==='string','VALIDATION_ERROR');N(!['REJECT','REQUEST_CORRECTION'].includes(decision)||p.observacaoPublica.trim()||p.observacoesInternas.trim(),'INCOMPLETE_REQUEST');
   command={commandType:'ACT.ATTENDANCE_JUSTIFICATION_DECIDE',payload:{justificationRef:M.ref('justification',p.idJustificativa),decision,reason:p.observacaoPublica||{APPROVE:'Justificativa deferida',EXCUSE:'Falta abonada',REJECT:'Justificativa indeferida',REQUEST_CORRECTION:'Solicitado ajuste da justificativa'}[decision],internalReason:p.observacoesInternas}};
  }
  return ports.prepare({operationId:input.operationId,...command});
 }
 return {read,prepare,row,pair,summary,validUntil:()=>temporalValidUntil};
}
