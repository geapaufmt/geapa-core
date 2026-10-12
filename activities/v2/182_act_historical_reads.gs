// Literal historical details only. Callers must authorize the parent activity.
// These DTOs never supply current policy, recognition, receipt or write authority.
function f4ActivityHistoricalDetails_(activity,get){
 if(activity.recordMode!=='HISTORICAL_PARTIAL')return {};
 const e=get('migrationEvidence/'+activity.provenanceHead.id);
 f4ActivitiesModel_().need(e?.schemaVersion==='act-migration-assertion-v1'&&e.literal?.ID_ATIVIDADE===activity.id,'UNRESOLVED_DEPENDENCY');
 const s=e.literal,text=k=>typeof s[k]==='string'?s[k]:'';
 const flag=k=>s[k]==='SIM'?true:s[k]==='NAO'?false:null,hours=String(s.CARGA_HORARIA||'').replace(',','.');
 return {contaPresenca:flag('CONTA_PRESENCA'),contaFalta:flag('CONTA_FALTA'),geraCertificado:flag('GERA_CERTIFICADO'),cargaHoraria:/^\d+(?:\.\d+)?$/.test(hours)?Number(hours):null,classificacaoAcesso:text('CLASSIFICACAO_ACESSO')||null,tipoAtividade:text('TIPO_ATIVIDADE'),subtipoAtividade:text('SUBTIPO_ATIVIDADE'),tipoPublico:text('TIPO_ATIVIDADE'),rotuloSemestre:text('CICLO'),horarioInicio:text('HORARIO_INICIO'),horarioFim:text('HORARIO_FIM'),horarioCompleto:[text('HORARIO_INICIO'),text('HORARIO_FIM')].filter(Boolean).join(' às '),responsavelInterno:text('RESPONSAVEL_INTERNO'),responsavelEmail:text('RESPONSAVEL_EMAIL'),eixoTematicoPrincipal:text('EIXO_TEMATICO_PRINCIPAL'),eixoTematicoSecundario:text('EIXO_TEMATICO_SECUNDARIO')};
}
function f4ActivityHistoricalRelations_(activity,get,artifacts){
 const M=f4ActivitiesModel_(),N=M.need;N(activity.recordMode==='HISTORICAL_PARTIAL','VALIDATION_ERROR');
 const safeLink=s=>typeof s==='string'&&/^https:\/\/(drive\.google\.com|docs\.google\.com)\//.test(s)?s:'';
 const rows=(activity.historicalPresentationRefs||[]).map(ref=>{
  const pr=get(M.path(ref));N(pr?.recordMode==='HISTORICAL_PARTIAL'&&pr.activityRef?.resolution==='RESOLVED'&&pr.activityRef.ref.id===activity.id,'UNRESOLVED_DEPENDENCY');
  const evidence=get(M.path(pr.provenanceHead));N(evidence?.schemaVersion==='act-migration-assertion-v1'&&evidence.literal?.ID_APRESENTACAO===pr.id,'UNRESOLVED_DEPENDENCY');
  const literal=evidence.literal,personId=pr.personRef?.resolution==='RESOLVED'?pr.personRef.ref.id:null,person=personId&&get('memberIdentities/'+personId);
  const files=artifacts.filter(v=>v.recordMode==='HISTORICAL_PARTIAL'&&v.activityRef?.resolution==='RESOLVED'&&v.activityRef.ref.id===activity.id&&v.presentationRef?.resolution==='RESOLVED'&&v.presentationRef.ref.id===pr.id);
  const links={};for(const file of files){const e=get(M.path(file.provenanceHead));N(e?.schemaVersion==='act-migration-assertion-v1','UNRESOLVED_DEPENDENCY');const kind=e.literal.TIPO_ARQUIVO_ATIVIDADE,link=safeLink(file.link);if(link&&!links[kind])links[kind]=link;}
  const material=links.SLIDE_APRESENTACAO||'',photo=links.FOTO_REUNIAO||'';
  return {idApresentacao:pr.id,idAtividade:activity.id,recordMode:'HISTORICAL_PARTIAL',somenteConsulta:true,dataAtividade:activity.civilDate||'',tema:pr.title||'',tituloApresentacao:pr.title||'',idPessoaPrincipal:personId||'',nomeApresentador:person?.displayName||'',eixoTematicoPrincipal:literal.EIXO_TEMATICO_PRINCIPAL||'',eixoTematicoSecundario:literal.EIXO_TEMATICO_SECUNDARIO||'',statusApresentacao:pr.legacyStatus||'NAO_INFORMADO',statusTituloEixo:'NAO_RECONSTRUIDO',statusMaterial:material?'REFERENCIA_HISTORICA':literal.STATUS_ENVIO_MATERIAL||'NAO_RECONSTRUIDO',statusFotoReuniao:photo?'REFERENCIA_HISTORICA':'NAO_RECONSTRUIDO',linkMaterial:material,linkFotoReuniao:photo,acoesMembro:{podeEditarTituloEixo:false,podeEnviarMaterial:false,podeReenviarMaterial:false,podeAbrirMaterial:!!material,podeEnviarFotoReuniao:false,podeReenviarFotoReuniao:false,podeAbrirFotoReuniao:!!photo,podeAbrirPastaAtividade:false},acoesGestao:{},canonicalRealizationAvailable:false,canonicalMaterialFile:false,canonicalFinalizationAvailable:false};
 });
 const involved=(activity.historicalInvolvedRefs||[]).map(r=>{const e=get(M.path(r));N(e?.literal?.ID_ATIVIDADE===activity.id,'UNRESOLVED_DEPENDENCY');const x=e.literal,id=e.subjectRefs?.[0]?.id||'',person=id&&get('memberIdentities/'+id);return {idPessoa:id,nome:x.NOME_PUBLICO||person?.displayName||'',papel:x.PAPEL_NA_ATIVIDADE||''};});
 return {apresentacoes:rows,envolvidos:involved.length?involved:rows.filter(v=>v.idPessoaPrincipal).map(v=>({idPessoa:v.idPessoaPrincipal,nome:v.nomeApresentador,papel:'APRESENTADOR'})),pendencias:[]};
}
