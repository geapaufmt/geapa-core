// Pure rules reused from geapa-membros V2. No endpoint, writer, authorization
// issuer, Sheets or jobs. Caller must authenticate, authorize and recheck all
// canonical revisions inside the shared command transaction.
// Incoming-member SELF eligibility: explicit user decision, 2026-10-13.
function f4MemberVoluntaryRules_(){
 const C=f4Core_(),N=C.need;
 const MEMBERS_VINCULO_CFG={"normativeIds":{"suspensionMinimum":"SUSPENSAO_MINIMA","dismissalPresentationBlock":"BLOQUEIO_DESLIGAMENTO_ANTES_APRESENTACAO","dismissalFinalMinutes":"DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA"},"types":{"suspension":"SUSPENSAO_VOLUNTARIA","dismissal":"DESLIGAMENTO_VOLUNTARIO"},"modalities":{"suspension":"SUSPENSAO_TEMPORARIA","dismissalImmediate":"DESLIGAMENTO_APOS_HOMOLOGACAO","dismissalSemesterEnd":"DESLIGAMENTO_FIM_SEMESTRE"},"statuses":{"received":"RECEBIDO","analyzing":"EM_ANALISE","complement":"AGUARDANDO_COMPLEMENTO","scheduledFinal":"AGENDADO_PARA_ANALISE_FINAL","readyFinal":"PRONTO_PARA_ANALISE_FINAL","approvedFuture":"HOMOLOGADO_AGUARDANDO_INICIO","suspensionActive":"SUSPENSAO_ATIVA","completed":"CONCLUIDO","denied":"INDEFERIDO","executed":"EXECUTADO","memberCancelled":"CANCELADO_PELO_MEMBRO","adminCancelled":"CANCELADO_PELA_DIRETORIA","executionError":"ERRO_EXECUCAO"},"transitions":{"RECEBIDO":["EM_ANALISE","AGUARDANDO_COMPLEMENTO","INDEFERIDO","CANCELADO_PELO_MEMBRO","CANCELADO_PELA_DIRETORIA"],"EM_ANALISE":["AGUARDANDO_COMPLEMENTO","AGENDADO_PARA_ANALISE_FINAL","HOMOLOGADO_AGUARDANDO_INICIO","SUSPENSAO_ATIVA","INDEFERIDO","EXECUTADO","CANCELADO_PELO_MEMBRO","CANCELADO_PELA_DIRETORIA"],"AGUARDANDO_COMPLEMENTO":["EM_ANALISE","INDEFERIDO","CANCELADO_PELO_MEMBRO","CANCELADO_PELA_DIRETORIA"],"AGENDADO_PARA_ANALISE_FINAL":["PRONTO_PARA_ANALISE_FINAL","AGUARDANDO_COMPLEMENTO","INDEFERIDO","CANCELADO_PELO_MEMBRO","CANCELADO_PELA_DIRETORIA"],"PRONTO_PARA_ANALISE_FINAL":["AGUARDANDO_COMPLEMENTO","INDEFERIDO","EXECUTADO","CANCELADO_PELO_MEMBRO","CANCELADO_PELA_DIRETORIA"],"HOMOLOGADO_AGUARDANDO_INICIO":["SUSPENSAO_ATIVA","CANCELADO_PELA_DIRETORIA","ERRO_EXECUCAO"],"SUSPENSAO_ATIVA":["CONCLUIDO","ERRO_EXECUCAO"],"ERRO_EXECUCAO":["HOMOLOGADO_AGUARDANDO_INICIO","SUSPENSAO_ATIVA","CONCLUIDO","EXECUTADO","CANCELADO_PELA_DIRETORIA"]}};
function membersVinculoError_(code, message, details, fieldErrors) {
  var error = new Error(message);
  error.legacyCode = code;
  error.code = code.indexOf('PARAMETRO_') === 0 ? 'RECONCILIATION_REQUIRED' : code === 'TRANSICAO_STATUS_INVALIDA' ? 'DOMAIN_STATE_INVALID' : 'COMMAND_PAYLOAD_INVALID';
  error.errorCode = code;
  error.details = details || {};
  error.fieldErrors = fieldErrors || {};
  return error;
}

function membersVinculoToken_(value) {
  return String(value == null ? '' : value).trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[\s-]+/g, '_');
}

function membersVinculoNormativeBoolean_(value, parameterId) {
  if (value === true || value === false) return value;
  var token = membersVinculoToken_(value);
  if (['SIM', 'TRUE', '1'].indexOf(token) >= 0) return true;
  if (['NAO', 'FALSE', '0'].indexOf(token) >= 0) return false;
  throw membersVinculoError_('PARAMETRO_NORMATIVO_VALOR_BOOLEANO_INVALIDO', 'O parametro normativo booleano esta indisponivel.', { parametroId: parameterId });
}

function membersVinculoValidateFinalMinutesParameter_(parameter) {
  var expectedId = MEMBERS_VINCULO_CFG.normativeIds.dismissalFinalMinutes;
  var id = membersVinculoToken_(parameter && (parameter.parametroId || parameter.PARAMETRO_ID));
  if (id !== expectedId) throw membersVinculoError_('PARAMETRO_NORMATIVO_ID_INCOMPATIVEL', 'Parametro normativo inesperado.', { esperado: expectedId, recebido: id });
  var vigente = parameter && (parameter.vigente != null ? parameter.vigente : parameter.VIGENTE);
  if (!(vigente === true || membersVinculoToken_(vigente) === 'SIM')) throw membersVinculoError_('PARAMETRO_NORMATIVO_NAO_VIGENTE', 'O parametro normativo esta indisponivel.', { parametroId: id });
  var tipo = membersVinculoToken_(parameter.tipoValor || parameter.TIPO_VALOR);
  if (tipo !== 'BOOLEANO') throw membersVinculoError_('PARAMETRO_NORMATIVO_TIPO_INVALIDO', 'O parametro normativo deve possuir TIPO_VALOR BOOLEANO.', { parametroId: id, tipoValor: tipo });
  var unidade = membersVinculoToken_(parameter.unidade || parameter.UNIDADE || 'NAO_APLICAVEL');
  if (unidade !== 'NAO_APLICAVEL') throw membersVinculoError_('PARAMETRO_NORMATIVO_UNIDADE_INVALIDA', 'O parametro normativo booleano deve usar NAO_APLICAVEL.', { parametroId: id });
  var baseLegal = String(parameter.baseLegal || parameter.BASE_LEGAL || '').trim();
  if (!baseLegal) throw membersVinculoError_('PARAMETRO_NORMATIVO_BASE_LEGAL_AUSENTE', 'O parametro normativo nao possui BASE_LEGAL.', { parametroId: id });
  var moduleSystem = String(parameter.moduloSistema || parameter.MODULO_SISTEMA || '').split(/[;,|]/).map(membersVinculoToken_).filter(Boolean);
  if (!moduleSystem.some(function(token) { return ['GEAPA_MEMBROS', 'MEMBROS', 'VINCULOS_GEAPA', 'TODOS', 'SISTEMA_GEAPA'].indexOf(token) >= 0; })) {
    throw membersVinculoError_('PARAMETRO_NORMATIVO_MODULO_INCOMPATIVEL', 'O parametro normativo esta indisponivel para o GEAPA Membros.', { parametroId: id });
  }
  var value = membersVinculoNormativeBoolean_(parameter.valor != null ? parameter.valor : parameter.VALOR, id);
  if (value === false && baseLegal === 'NC01-2025-ART16-IV') {
    throw membersVinculoError_('PARAMETRO_NORMATIVO_BASE_LEGAL_INCOMPATIVEL', 'A dispensa de ata exige nova BASE_LEGAL compativel com a alteracao normativa.', { parametroId: id, baseLegal: baseLegal });
  }
  return { parametroId: id, valor: value, tipoValor: 'BOOLEANO', unidade: 'NAO_APLICAVEL', baseLegal: baseLegal, vigente: true, moduloSistema: parameter.moduloSistema || parameter.MODULO_SISTEMA };
}

function membersVinculoAssertTransition_(from, to) {
  var source = membersVinculoToken_(from);
  var target = membersVinculoToken_(to);
  var allowed = MEMBERS_VINCULO_CFG.transitions[source] || [];
  if (allowed.indexOf(target) < 0) throw membersVinculoError_('TRANSICAO_STATUS_INVALIDA', 'Transicao de status nao permitida.', { statusAtual: source, statusPretendido: target });
  return target;
}

function membersVinculoAssertRequestTransition_(request, to, action) {
  var source = membersVinculoToken_(request.STATUS_SOLICITACAO);
  var target = membersVinculoToken_(to);
  var type = membersVinculoToken_(request.TIPO_SOLICITACAO);
  var modality = membersVinculoToken_(request.MODALIDADE_SOLICITADA);
  var operation = membersVinculoToken_(action);
  if (source === MEMBERS_VINCULO_CFG.statuses.received && type === MEMBERS_VINCULO_CFG.types.dismissal) {
    if (target === MEMBERS_VINCULO_CFG.statuses.executed && modality === MEMBERS_VINCULO_CFG.modalities.dismissalImmediate && operation === 'HOMOLOGAR_E_EFETIVAR_DESLIGAMENTO') return target;
    if (target === MEMBERS_VINCULO_CFG.statuses.scheduledFinal && modality === MEMBERS_VINCULO_CFG.modalities.dismissalSemesterEnd && operation === 'ANALISE_PRELIMINAR_REGISTRADA') return target;
  }
  return membersVinculoAssertTransition_(source, target);
}

function membersVinculoRequiresFinalMinutes_(request) {
  return membersVinculoToken_(request.TIPO_SOLICITACAO) === MEMBERS_VINCULO_CFG.types.dismissal;
}

function membersVinculoAssertFinalDecisionDocument_(request, payload, requirement) {
  if (!membersVinculoRequiresFinalMinutes_(request) || !requirement || requirement.exigeAta !== true) return true;
  var ata = String((payload && payload.ataReferencia) || request.ATA_REFERENCIA || '').trim();
  var idAta = String((payload && payload.idAtaDeliberacao) || request.ID_ATA_DELIBERACAO || '').trim();
  if (!ata && !idAta) throw membersVinculoError_('ATA_DECISAO_FINAL_OBRIGATORIA', 'A decisao final do desligamento voluntario exige referencia oficial de ata.', {});
  return true;
}

 function eligible(link,personId,state,today){
  C.civilStart(today);
  N(link&&link.idPessoa===personId&&['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE'].includes(link.type),'ACCESS_REVOKED');
  N(state&&state.idPessoa===personId&&state.state===link.type&&state.blocked===false&&state.stale===false,'ACCESS_REVOKED');
  N(link.startDate<=today&&(link.endDate===null||today<=link.endDate),'ACCESS_REVOKED');
  C.civilStart(link.startDate);if(link.endDate!==null)C.civilStart(link.endDate);
  return true;
 }
 function voluntary(request){N(request?.TIPO_SOLICITACAO==='DESLIGAMENTO_VOLUNTARIO'&&['DESLIGAMENTO_APOS_HOMOLOGACAO','DESLIGAMENTO_FIM_SEMESTRE'].includes(request.MODALIDADE_SOLICITADA),'COMMAND_PAYLOAD_INVALID');}
 function transition(request,to,action){voluntary(request);return membersVinculoAssertRequestTransition_(request,to,action);}
 function finalDecision(request,payload,current,today,semester,reject=false){
  voluntary(request);C.civilStart(today);
  const parameter=membersVinculoValidateFinalMinutesParameter_(current);
  // Missing or changed snapshots must go through explicit normative review;
  // absence never means that minutes can be waived.
  N(request.PARAMETRO_ATA_DESLIGAMENTO_ID===parameter.parametroId&&request.ATA_DESLIGAMENTO_TIPO_VALOR_SNAPSHOT==='BOOLEANO'&&request.ATA_DESLIGAMENTO_BASE_LEGAL_SNAPSHOT===parameter.baseLegal,'RECONCILIATION_REQUIRED');
  N(membersVinculoNormativeBoolean_(request.ATA_DESLIGAMENTO_OBRIGATORIA_SNAPSHOT,parameter.parametroId)===parameter.valor,'RECONCILIATION_REQUIRED');
  N(payload.confirmacaoReforcada===true,'COMMAND_PAYLOAD_INVALID');
  membersVinculoAssertFinalDecisionDocument_(request,payload,{exigeAta:parameter.valor});
  let effectiveDate=today;
  if(request.MODALIDADE_SOLICITADA==='DESLIGAMENTO_FIM_SEMESTRE'&&!reject){
   N(semester&&semester.id===request.ID_SEMESTRE_REFERENCIA,'RECONCILIATION_REQUIRED');
   C.civilStart(semester.civilStart);C.civilStart(semester.civilEnd);
   N(semester.civilStart<=semester.civilEnd&&today>=semester.civilEnd,'DOMAIN_STATE_INVALID');
   effectiveDate=semester.civilEnd;
  }
  return {effectiveDate,requiresMinutes:parameter.valor,calendarChanged:!!semester&&semester.civilEnd!==request.DATA_FIM_SEMESTRE_SNAPSHOT};
 }
 return {eligible,transition,finalDecision,minutes:membersVinculoValidateFinalMinutesParameter_};
}
