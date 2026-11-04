// A canonical promotion changes the parent activity mode, never its imported children.
function f4ActivityHasHistoricalRelations_(a){return !!a&&(a.recordMode==='HISTORICAL_PARTIAL'||a.recordMode==='NEW_V2'&&Array.isArray(a.historicalPresentationRefs)&&(a.promotion?.schemaVersion==='activity-operational-promotion-v1'&&a.promotion.fromMode==='HISTORICAL_PARTIAL'&&a.promotion.sourceRef?.id===a.provenanceHead?.id||a.importRecognition?.schemaVersion==='act-import-recognition-v1'&&a.importRecognition.previousMode==='HISTORICAL_PARTIAL'&&!!a.provenanceHead));}
// Literal historical details only. Callers must authorize the parent activity.
// These DTOs never supply current policy, recognition, receipt or write authority.
function f4ActivityHistoricalDetails_(activity,get){
 if(activity.recordMode!=='HISTORICAL_PARTIAL')return {};
 const e=get('migrationEvidence/'+activity.provenanceHead.id);
 f4ActivitiesModel_().need(e?.schemaVersion==='act-migration-assertion-v1'&&e.literal?.ID_ATIVIDADE===activity.id,'UNRESOLVED_DEPENDENCY');
 const s={...e.literal},M=f4ActivitiesModel_();if(activity.scheduleRecognition?.kind==='HISTORICAL_RESCHEDULE'){s.HORARIO_INICIO=Utilities.formatDate(new Date(M.ms(activity.startsAt)),'America/Cuiaba','HH:mm');s.HORARIO_FIM=Utilities.formatDate(new Date(M.ms(activity.endsAt)),'America/Cuiaba','HH:mm');}const text=k=>typeof s[k]==='string'?s[k]:'';
 const flag=k=>s[k]==='SIM'?true:s[k]==='NAO'?false:null,hours=String(s.CARGA_HORARIA||'').replace(',','.');
 return {contaPresenca:flag('CONTA_PRESENCA'),contaFalta:flag('CONTA_FALTA'),geraCertificado:flag('GERA_CERTIFICADO'),cargaHoraria:/^\d+(?:\.\d+)?$/.test(hours)?Number(hours):null,classificacaoAcesso:text('CLASSIFICACAO_ACESSO')||null,tipoAtividade:text('TIPO_ATIVIDADE'),subtipoAtividade:text('SUBTIPO_ATIVIDADE'),tipoPublico:text('TIPO_ATIVIDADE'),horarioInicio:text('HORARIO_INICIO'),horarioFim:text('HORARIO_FIM'),horarioCompleto:[text('HORARIO_INICIO'),text('HORARIO_FIM')].filter(Boolean).join(' às '),responsavelInterno:activity.details?.responsavelInterno??text('RESPONSAVEL_INTERNO'),responsavelEmail:activity.details?.responsavelEmail??text('RESPONSAVEL_EMAIL'),eixoTematicoPrincipal:text('EIXO_TEMATICO_PRINCIPAL'),eixoTematicoSecundario:text('EIXO_TEMATICO_SECUNDARIO')};
}

// Historical links are exposed by activityRelations under ACTIVITY_BOARD_READ.
// Reuse that same institutional authority; a copied object creates no new right.
function f4ActivityHistoricalFilePlan_(options,request){
 const M=f4ActivitiesModel_(),C=f4Core_(),N=M.need,db=f4Firestore_().connect(options.config),get=p=>db.get(p),clock=options.clock||Date.now;
 C.exact(request.params,['idApresentacao',...('idArquivo'in request.params?['idArquivo']:[])]);const id=request.params.idApresentacao;
 if('idArquivo'in request.params)N(M.id(request.params.idArquivo),'VALIDATION_ERROR');
 N(typeof id==='string'&&/^APR-[A-Za-z0-9_-]+$/.test(id),'VALIDATION_ERROR');
 const kind=request.route==='/v2/apresentacoes/material/arquivo'?'SLIDE_APRESENTACAO':request.route==='/v2/apresentacoes/foto/arquivo'?'FOTO_REUNIAO':null;N(kind,'COMMAND_TYPE_INVALID');
 return options.ports.readWithAuthoritySnapshot(check=>{
  const pr=get('canonicalPresentations/'+id),original=f4PresentationContinuity_().source(pr),activity=original?.activityRef?.resolution==='RESOLVED'&&get('canonicalActivities/'+original.activityRef.ref.id);
  N(original&&f4ActivityHasHistoricalRelations_(activity)&&activity.sourceSnapshotState!=='ABSENT_FROM_CURRENT_SNAPSHOT','UNRESOLVED_DEPENDENCY');
  const auth=options.ports.readAuthority(['canonicalActivities/'+activity.id],false,'ACTIVITY_BOARD_READ');
  const control=get('actHistoricalFileControl/current');N(f4ExecutionEnvironment_().destination(options.config)&&control?.projectId===options.config.projectId&&control.enabled===true&&control.bucket===f4ExecutionEnvironment_().project+'.firebasestorage.app','CONFIG_CAPABILITY_DISABLED');
  N(Array.isArray(control.files)&&control.files.length<=67,'SCOPE_TOO_LARGE');const matches=control.files.filter(f=>f.presentationId===id&&f.kind===kind&&(!request.params.idArquivo||f.artifactId===request.params.idArquivo)).sort((a,b)=>a.artifactId<b.artifactId?-1:a.artifactId>b.artifactId?1:0);N(matches.length>0&&(!request.params.idArquivo||matches.length===1),'UNRESOLVED_DEPENDENCY');const artifactId=matches[0].artifactId;N(typeof artifactId==='string'&&M.id(artifactId),'UNRESOLVED_DEPENDENCY');
  const art=get('presentationArtifacts/'+artifactId),binding=get('actHistoricalFileBindings/'+artifactId),copy=binding&&get('actHistoricalFileCopies/'+binding.copyId),r=copy?.receipt;
  N(art?.recordMode==='HISTORICAL_PARTIAL'&&art.activityRef?.ref?.id===activity.id&&art.presentationRef?.ref?.id===id&&binding?.state==='ACTIVE'&&binding.manifestDigest===control.manifestDigest&&binding.artifactRevision===art.revision&&binding.activityId===activity.id&&binding.presentationId===id&&binding.kind===kind&&binding.originalDriveId===art.externalId&&binding.originalLink===art.link,'STALE_REVISION');
  const provenance=get(M.path(art.provenanceHead));N(provenance?.literal?.TIPO_ARQUIVO_ATIVIDADE===kind&&provenance.literal.ID_ATIVIDADE===activity.id,'UNRESOLVED_DEPENDENCY');
  N(copy.state==='VERIFIED'&&r?.institutionalReceipt===false&&r.source.id===art.externalId&&r.migrationId===binding.copyId&&r.provider==='GOOGLE_CLOUD_STORAGE'&&r.bucket===control.bucket&&r.objectName===(f4ExecutionEnvironment_().production?'act-history/':'act-dev/')+binding.copyId&&/^[0-9]+$/.test(r.generation)&&/^[a-f0-9]{64}$/.test(r.descriptor?.sha256)&&r.descriptor.name===art.name&&r.descriptor.mime===art.mimeType,'UNRESOLVED_DEPENDENCY');
  check([auth]);return {historical:true,controlRevision:control.revision,bindingRevision:binding.revision,copyRevision:copy.revision,settings:{provider:'GOOGLE_CLOUD_STORAGE',bucket:control.bucket},request:{effectId:binding.copyId,fileId:r.objectName,descriptor:r.descriptor},version:r.generation};
 });
}
function f4ActivityHistoricalRelations_(activity,get,artifacts){
 const M=f4ActivitiesModel_(),N=M.need;N(f4ActivityHasHistoricalRelations_(activity),'VALIDATION_ERROR');
 const safeLink=s=>typeof s==='string'&&/^https:\/\/(drive\.google\.com|docs\.google\.com)\//.test(s)?s:'';
 const rows=(activity.historicalPresentationRefs||[]).map(ref=>{
  const current=get(M.path(ref)),pr=f4PresentationContinuity_().source(current);N(pr?.recordMode==='HISTORICAL_PARTIAL'&&pr.activityRef?.resolution==='RESOLVED'&&pr.activityRef.ref.id===activity.id,'UNRESOLVED_DEPENDENCY');
  const evidence=get(M.path(pr.provenanceHead));N(evidence?.schemaVersion==='act-migration-assertion-v1'&&evidence.literal?.ID_APRESENTACAO===pr.id,'UNRESOLVED_DEPENDENCY');
  const literal=evidence.literal,personId=pr.personRef?.resolution==='RESOLVED'?pr.personRef.ref.id:null,person=personId&&get('memberIdentities/'+personId);
  const files=artifacts.filter(v=>v.recordMode==='HISTORICAL_PARTIAL'&&v.activityRef?.resolution==='RESOLVED'&&v.activityRef.ref.id===activity.id&&v.presentationRef?.resolution==='RESOLVED'&&v.presentationRef.ref.id===pr.id);
  const links={};for(const file of files){const e=get(M.path(file.provenanceHead));N(e?.schemaVersion==='act-migration-assertion-v1','UNRESOLVED_DEPENDENCY');const kind=e.literal.TIPO_ARQUIVO_ATIVIDADE,link=safeLink(file.link);if(link&&!links[kind])links[kind]=link;}
  const material=links.SLIDE_APRESENTACAO||'',photo=links.FOTO_REUNIAO||'';
  // Documentary states describe the historical record, not a new canonical
  // approval. A file reference must never replace an explicit receipt state.
  const text=k=>typeof literal[k]==='string'?literal[k].trim():'';
  const title=pr.title||text('TITULO_APRESENTACAO');
  return {...f4ActivityAcademicMetadata_(activity,get),idApresentacao:pr.id,idAtividade:activity.id,recordMode:'HISTORICAL_PARTIAL',somenteConsulta:true,dataAtividade:activity.civilDate||'',tema:title,tituloApresentacao:title,idPessoaPrincipal:personId||'',nomeApresentador:person?.displayName||'',eixoTematicoPrincipal:text('EIXO_TEMATICO_PRINCIPAL'),eixoTematicoSecundario:text('EIXO_TEMATICO_SECUNDARIO'),statusApresentacao:pr.legacyStatus||'NAO_INFORMADO',statusTituloEixo:text('STATUS_TITULO_EIXO')||'NAO_RECONSTRUIDO',statusMaterial:text('STATUS_ENVIO_MATERIAL')||(material?'REFERENCIA_HISTORICA':'NAO_RECONSTRUIDO'),statusFotoReuniao:text('STATUS_ENVIO_FOTO_REUNIAO')||(photo?'REFERENCIA_HISTORICA':'NAO_RECONSTRUIDO'),linkMaterial:material,linkFotoReuniao:photo,acoesMembro:{podeEditarTituloEixo:false,podeEnviarMaterial:false,podeReenviarMaterial:false,podeAbrirMaterial:!!material,podeEnviarFotoReuniao:false,podeReenviarFotoReuniao:false,podeAbrirFotoReuniao:!!photo,podeAbrirPastaAtividade:false},acoesGestao:{},canonicalRealizationAvailable:false,canonicalMaterialFile:false,canonicalFinalizationAvailable:false};
 });
 const involved=(activity.historicalInvolvedRefs||[]).map(r=>{const e=get(M.path(r));N(e?.literal?.ID_ATIVIDADE===activity.id,'UNRESOLVED_DEPENDENCY');const x=e.literal,id=e.subjectRefs?.[0]?.id||'',person=id&&get('memberIdentities/'+id);return {idPessoa:id,nome:x.NOME_PUBLICO||person?.displayName||'',papel:x.PAPEL_NA_ATIVIDADE||''};});
 return {apresentacoes:rows,envolvidos:involved.length?involved:rows.filter(v=>v.idPessoaPrincipal).map(v=>({idPessoa:v.idPessoaPrincipal,nome:v.nomeApresentador,papel:'APRESENTADOR'})),pendencias:[]};
}
