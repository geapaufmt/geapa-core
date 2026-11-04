// Calendar labels are evidence, not inferred associations. No date/ID heuristic.
function f4ActivityAcademicMetadata_(a,get){
 const text=v=>typeof v==='string'?v.trim():'';
 const academicSemesterId=text(a.academicSemesterId)||null,cycleId=text(a.cycleId)||text(a.cycleRef?.id)||(a.cycleRef?.resolution==='RESOLVED'?text(a.cycleRef.ref?.id):'')||null;
 const evidence=a.provenanceHead?.id&&get?get('migrationEvidence/'+a.provenanceHead.id):null;
 const literal=evidence?.schemaVersion==='act-migration-assertion-v1'&&evidence.literal?.ID_ATIVIDADE===a.id?evidence.literal:null;
 const year=text(literal?.ANO),half=text(literal?.SEMESTRE),historicalSemesterLabel=/^\d{4}$/.test(year)&&/^[12]$/.test(half)?year+'/'+half:'';
 const semester=academicSemesterId&&get?get('academicSemesters/'+academicSemesterId):null;
 const canonical=semester?.id===academicSemesterId&&text(semester?.label)?text(semester.label):'';
 return {academicSemesterId,cycleId,academicSemesterLabel:canonical,historicalSemesterLabel,historicalCycleId:text(literal?.CICLO)||null,semesterProvenanceRef:literal?{kind:'evidence',id:a.provenanceHead.id}:null,semesterResolution:academicSemesterId?(canonical?'CANONICAL':'UNRESOLVED_REFERENCE'):historicalSemesterLabel?'HISTORICAL_DECLARED':'UNDEFINED',rotuloSemestre:academicSemesterId?canonical:historicalSemesterLabel};
}

// Resolve membership in a real activity cycle from stored references only.
// Legacy period labels are documentary aliases, never cycle identifiers.
function f4ActivityFrequencyCycle_(activity,collection,fact,get,semesters){
 const text=v=>typeof v==='string'?v.trim():'',ref=v=>text(v?.id)||(v?.resolution==='RESOLVED'?text(v.ref?.id):'');
 const ids=new Set(),sources=[],unresolved=[];
 function cycle(id,source){if(!id)return;const c=get('presentationCycles/'+id);if(c?.id===id){ids.add(id);sources.push(source);}else unresolved.push(source);}
 function semester(id,source){if(!id)return;const s=get('academicSemesters/'+id);if(s?.id===id&&s.cycleId)cycle(s.cycleId,source);else unresolved.push(source);}
 for(const [row,label] of [[activity,'activity'],[collection,'collection'],[fact,'attendance']]){cycle(text(row?.cycleId)||ref(row?.cycleRef),label+'.cycle');semester(text(row?.academicSemesterId)||ref(row?.academicSemesterRef),label+'.semester');}
 const sessionId=ref(activity?.presentationSessionRef)||ref(activity?.sessionRef);
 if(sessionId){const s=get('presentationSessions/'+sessionId);if(s)cycle(text(s.cycleId)||ref(s.cycleRef),'session.cycle');else unresolved.push('session');}
 function result(kind){const id=ids.size===1&&!unresolved.length?[...ids][0]:null;return {cycleId:id,cycleLabel:id?(get('presentationCycles/'+id).cycleLabel||id):'',cycleResolution:id?kind:ids.size>1?'CONFLICT':'UNRESOLVED',cycleResolutionSources:sources,cycleResolutionPending:unresolved};}
 if(ids.size||unresolved.length)return result('CANONICAL');
 const meta=f4ActivityAcademicMetadata_(activity,get);
 if(meta.historicalCycleId)cycle(meta.historicalCycleId,'activity.documentaryCycle');
 function label(value,source){if(!value)return;const candidates=semesters.filter(s=>s.label===value);if(candidates.length===1)semester(candidates[0].id,source);else if(candidates.length>1)unresolved.push(source+'.ambiguous');}
 label(meta.historicalSemesterLabel,'activity.documentarySemester');
 const periodId=text(fact?.periodId)||ref(collection?.periodRef);
 if(periodId){const p=get('activityFrequencyPeriods/'+periodId);if(p){cycle(text(p.cycleId)||ref(p.cycleRef),'period.cycle');semester(text(p.academicSemesterId)||ref(p.academicSemesterRef),'period.semester');label(text(p.label),'period.documentedLabel');}}
 return result('DOCUMENTED');
}
function f4ActivityPortalDto_(a,c,admin,clock,presentations=[],get){
 const M=f4ActivitiesModel_(),iso=t=>t?new Date(M.ms(t)).toISOString():'',hour=t=>t?Utilities.formatDate(new Date(M.ms(t)),'America/Cuiaba','HH:mm'):'';
  const partial=a.recordMode==='HISTORICAL_PARTIAL',historical=partial||a.recordMode==='HISTORICAL_RECOGNIZED',type=a.typeSnapshot||{},model=type.portalModel||{},details=a.details||{},now=clock(),within=c&&now>=M.ms(c.opensAt)&&now<=M.ms(c.closesAt),finalized=c?.state==='CLOSED',state=finalized?'FINALIZADA':c?.draftSavedAt&&(!c.reopenedAt||M.ms(c.draftSavedAt)>=M.ms(c.reopenedAt))?'SALVA':c?.state==='REOPENED'?'REABERTA':'RASCUNHO';
  const hidden=a.publicationStatus==='OCULTA',publication=a.publicationStatus==='PUBLICADO'?'PUBLICADA':a.publicationStatus;
  return {...f4ActivityAcademicMetadata_(a,get),...f4ActivityReview_().dto(a,admin),recordMode:a.recordMode||'NEW_V2',reconstrucaoConcluida:a.recordMode==='HISTORICAL_RECOGNIZED'||!!a.importRecognition,somenteConsulta:historical,...f4ActivityContentDto_(a,admin,get),idAtividade:a.id,canonicalActivityRevision:a.revision,tituloPublico:a.title||'',tituloConteudoPublico:a.title||'',descricaoPublica:details.descricaoPublica||'',...(admin?{descricaoInterna:details.descricaoInterna||'',observacoes:details.observacoes||'',responsavelInterno:details.responsavelInterno||'',responsavelEmail:details.responsavelEmail||''}:{}),dataAtividade:a.civilDate,diaSemana:['domingo','segunda-feira','terça-feira','quarta-feira','quinta-feira','sexta-feira','sábado'][new Date(a.civilDate+'T12:00:00Z').getUTCDay()]||'',horarioInicio:hour(a.startsAt),horarioFim:hour(a.endsAt),horarioCompleto:a.startsAt&&a.endsAt?hour(a.startsAt)+' às '+hour(a.endsAt):'',dataHoraInicio:iso(a.startsAt),dataHoraFim:iso(a.endsAt),local:a.location||'',formato:a.format||'',publicoAlvo:details.publicoAlvo||'',tipoPublico:model.type||type.name||'',tipoAtividade:model.type||'',subtipoAtividade:model.subtype||'',idConfigModelo:a.typeRef?.id||null,nomeModeloPortal:type.name||null,classificacaoAcesso:{BOARD:'INTERNA_DIRETORIA',MEMBERS:'INTERNA_MEMBROS',PUBLIC:'PUBLICA'}[type.visibility]||null,contaPresenca:type.countsAttendance??null,contaFalta:historical?(type.countsAbsence??null):type.countsAbsence===true,geraCertificado:historical?(model.generatesCertificate??null):model.generatesCertificate===true,cargaHoraria:type.durationMinutes==null?null:type.durationMinutes/60,statusOperacional:partial?(a.legacyStatus||'NAO_INFORMADO'):{SCHEDULED:'PLANEJADA',OCCURRED:'REALIZADA',NOT_OCCURRED:'NAO_REALIZADA',CANCELLED:'CANCELADA',ARCHIVED:'ARQUIVADA'}[a.status],statusPublico:publication||null,statusPublicacaoPortal:publication||null,visibilidadePortal:hidden?'OCULTA':{BOARD:'DIRETORIA',MEMBERS:'MEMBROS',PUBLIC:'PUBLICA'}[a.publishedVisibility||'BOARD'],chamadaDisponivelEm:iso(c?.opensAt),chamadaEncerraEm:iso(c?.closesAt),podeRegistrarChamadaAgora:!!(!historical&&admin&&within&&!finalized&&a.status!=='CANCELLED'),podeRegistrarChamada:!!(!historical&&admin&&within&&!finalized&&a.status!=='CANCELLED'),podeVisualizarChamada:!!(!historical&&admin&&c&&(within||state!=='RASCUNHO')),chamadaFinalizada:historical?null:finalized,statusChamada:historical?'NAO_RECONSTRUIDA':state,statusChamadaRotulo:historical?(c?.historicalReadOnly&&c.recordedFactCount>0?'Registros históricos de presença: '+c.recordedFactCount:'Chamada histórica não reconstruída'):{FINALIZADA:'Chamada finalizada',SALVA:'Rascunho salvo',REABERTA:'Chamada reaberta',RASCUNHO:'Chamada pendente'}[state],statusChamadaAtualizadoEm:iso(c?.updatedAt),podeVerDetalhes:true,podeEditar:admin&&!historical&&a.administrativeReview?.state!=='CONFIRMED',possuiApresentacoes:presentations.length>0,qtdApresentacoes:presentations.length,apresentacoes:presentations,apresentacoesPublicas:presentations,envolvidos:[],pendencias:[],ultimaAtualizacao:iso(a.updatedAt),...(a.status==='CANCELLED'&&!c?{statusChamada:'NAO_APLICAVEL',statusChamadaRotulo:'Sem chamada — atividade cancelada',chamadaFinalizada:false}:{}),...(a.attendanceRecognition?.state==='NO_CALL'?{statusChamada:'SEM_CHAMADA',statusChamadaRotulo:'Sem chamada — condição histórica reconhecida',chamadaFinalizada:false,podeVisualizarChamada:false,podeRegistrarChamada:false,podeRegistrarChamadaAgora:false}:{})};
 }

// Updated in the same transaction as the facts. Private justification text,
// roster identities and board-only notes are never included in these views.
function f4ActivitiesProjectAttendanceReads_(x,changed){
 const activities=new Set(),{M,N}=x;
 for(const [path,d] of changed){if(!d||path.split('/').length!==2)continue;const col=path.split('/')[0];if(col==='canonicalActivities')activities.add(d.id);if(col==='activityAttendanceCollections')activities.add(d.activityRef.id);if(col==='presentationSessions'&&d.activityRef)activities.add(d.activityRef.id);if(col==='publicPresentations'){const pr=x.get('canonicalPresentations/'+d.publicId),s=pr?.sessionRef&&x.get(M.path(pr.sessionRef));if(s?.activityRef)activities.add(s.activityRef.id);}}
 // Presentation publication also changes the agenda read fence. A paginated
 // reader rejects a mixed generation, even when the activity revision is unchanged.
 if(changed.some(([path])=>path.startsWith('publicPresentations/'))&&activities.size){const ref=M.ref('attendanceRegistry','current'),h=x.get(M.path(ref));x.save(ref,{...h,epoch:(h?.epoch||0)+1});}
 for(const id of activities){
  const a=x.get('canonicalActivities/'+id),c=x.get('activityAttendanceCollections/'+id);N(a,'UNRESOLVED_DEPENDENCY');
  x.set('activityAgendaViews/'+id,{schemaVersion:'act-activity-agenda-v1',id,sourceActivityRevision:a.revision,sourceCollectionRevision:c?.revision||0,publicationStatus:a.publicationStatus||'RASCUNHO',visibility:a.publishedVisibility||'BOARD',data:f4ActivityPortalDto_(a,c,false,()=>x.now,f4ActivityPublishedPresentations_(a,x.get),x.get),allowsJustification:a.recordMode!=='HISTORICAL_PARTIAL'&&a.typeSnapshot?.allowsJustification===true,justificationOpensAt:c?.justificationOpensAt||null,projectedAt:M.stamp(x.now)});
 }
}

// Only the approved public payload can enter the existing activity detail UI.
function f4ActivityPublishedPresentations_(activity,get){
 const M=f4ActivitiesModel_(),ref=activity.presentationSessionRef;if(!ref)return [];
 const s=get(M.path(ref));M.need(s?.activityRef?.id===activity.id&&s.civilDate===activity.civilDate,'UNRESOLVED_DEPENDENCY','ACTIVITY_SESSION_DRIFT');
 const rows=[];for(const slot of s.slots){const unit=get(M.path(slot.unitRef)),pr=unit?.presentationRef&&get(M.path(unit.presentationRef)),pub=pr&&get('publicPresentations/'+pr.id);if(!pub?.published||pub.sanitized!==true||pub.schemaVersion!==1)continue;
  M.need(pr.sessionRef?.id===s.id&&pub.publicId===pr.id&&pub.sourceRevision===pr.revision,'STALE_REVISION','PUBLICATION_SOURCE_DRIFT');
  const row={};if(typeof pub.title==='string')row.titulo=pub.title;if(typeof pub.civilDate==='string')row.dataApresentacao=pub.civilDate;if(typeof pub.authorizedAxis==='string')row.eixoTematicoPrincipal=pub.authorizedAxisLabel||pub.authorizedAxis;if(typeof pub.permittedDisplay==='string')row.nomeApresentador=pub.permittedDisplay;if(pub.approvedFormat)row.formatoAprovado=pub.approvedFormat;rows.push(row);
 }return rows;
}
