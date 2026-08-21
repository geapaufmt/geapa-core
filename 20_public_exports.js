/**
 * ============================================================
 * 20_public_exports.gs
 * ============================================================
 *
 * API PÚBLICA do GEAPA-CORE como Library.
 *
 * IMPORTANTE:
 * - Apps Script Libraries exportam FUNÇÕES GLOBAIS.
 * - Objetos/constantes (ex.: const GEAPA_CORE = {...}) NÃO são exportados.
 *
 * Portanto, aqui existem apenas wrappers:
 *   function coreXxx(...) { return core_xxx_(...) }
 *
 * Exemplo no módulo:
 *   GEAPA_CORE.coreGetRegistry()
 *   GEAPA_CORE.coreGetSheetByKey("MEMBERS_ATUAIS")
 *   GEAPA_CORE.coreSendEmailHtml({...})
 */

/* ============================================================
 * REGISTRY
 * ============================================================ */

/** Retorna o registry inteiro (lido da planilha-mãe e cacheado). */
function coreGetRegistry() {
  return core_getRegistry_();
}

/**
 * Retorna a referência {id, sheet} de uma KEY do registry.
 * Ex.: coreGetRegistryRefByKey("MEMBERS_ATUAIS") -> {id, sheet}
 */
function coreGetRegistryRefByKey(key) {
  // ajuste o nome interno conforme você corrigir no 10_core_registry.gs:
  // recomendado: core_getRegistryRefByKey_(key)
  return core_getRegistryRefByKey_(key);
}

/**
 * Abre e retorna o Sheet diretamente via KEY do registry.
 * Ex.: coreGetSheetByKey("MEMBERS_ATUAIS") -> Sheet
 */
function coreGetSheetByKey(key) {
  return core_getSheetByKey_(key);
}

function coreGetCurrentEnv() {
  return core_getCurrentEnv_();
}

function coreGetRegistryMetaByKey(key) {
  return core_getRegistryMetaByKey_(key);
}

/* ============================================================
 * CONFIG_GEAPA
 * ============================================================ */

function coreGetGeapaConfigValue(key, opts) {
  return core_getGeapaConfigValue_(key, arguments.length >= 2 ? opts : {});
}

function coreGetGeapaConfigMap(opts) {
  return core_getGeapaConfigMap_(opts || {});
}

function coreGetGeapaConfigObject(opts) {
  return core_getGeapaConfigObject_(opts || {});
}

function coreDebugGeapaConfig(opts) {
  return core_debugGeapaConfig_(opts || {});
}

/* ============================================================
 * DOMINIOS CENTRAIS V2
 * ============================================================ */

function coreGetDomainsV2Schemas() {
  return core_getDomainsV2Schemas_();
}

function coreGetDomainsV2ContractKeys() {
  return core_getDomainsV2ContractKeys_();
}

function coreGetDomainSpreadsheet(domain, options) {
  options = options || {};
  return core_openDomainSpreadsheet_(domain, {
    ambiente: options.ambiente || options.environment,
    // Entregar a planilha inteira e um contrato potencialmente mutavel. Por
    // seguranca, a validacao de escrita e o default; leitura tolerante deve
    // usar coreGetDomainSheet/coreReadDomainRecords.
    forWrite: options.forWrite !== false
  });
}

function coreGetDomainSheet(domain, logicalSheet, options) {
  options = options || {};
  return core_getDomainSheet_(domain, logicalSheet, {
    ambiente: options.ambiente || options.environment,
    forWrite: options.forWrite === true || String(options.access || '').toUpperCase() === 'WRITE'
  });
}

function coreReadDomainRecords(domain, logicalSheet, options) {
  options = options || {};
  var sheet = core_getDomainSheet_(domain, logicalSheet, {
    ambiente: options.ambiente || options.environment
  });
  return core_readSheetRecords_(sheet, {
    headerRow: Number(options.headerRow || 1),
    skipBlankRows: options.skipBlankRows !== false
  });
}

/** Resolve um parametro normativo operacional sem fallback entre ambientes. */
function coreResolverParametroNormativoOperacional(parametroId, options) {
  return core_resolverParametroNormativoOperacional_(parametroId, options || {});
}

/** Resolve varios parametros normativos com uma unica leitura da fonte. */
function coreResolverParametrosNormativosOperacionais(parametroIds, options) {
  return core_resolverParametrosNormativosOperacionais_(parametroIds || [], options || {});
}

/** Diagnostico somente leitura da configuracao normativa por ambiente. */
function coreDiagnosticarParametrosNormativosOperacionais(options) {
  return core_diagnosticarParametrosNormativosOperacionais_(options || {});
}

/** Invalida somente o cache de leitura dos parametros normativos. */
function coreInvalidarCacheParametrosNormativosOperacionais(options) {
  return core_invalidarCacheParametrosNormativosOperacionais_(options || {});
}

function corePrepararParametrosNormativosTipados(options) {
  return core_prepararParametrosNormativosTipados_(options || {});
}

function coreAuditarPessoasV2(options) {
  return coreAuditarPessoasV2_(options || {});
}

function coreAuditarVigenciasV2(options) {
  return coreAuditarVigenciasV2_(options || {});
}

function coreAuditarDominiosCentraisV2(options) {
  return coreAuditarDominiosCentraisV2_(options || {});
}

function coreValidateDomainRegistry(domain, options) {
  options = options || {};
  return core_validateDomainRegistry_(domain, { ambiente: options.ambiente || options.environment });
}

function coreValidateAllDomainRegistries(options) {
  options = options || {};
  return core_validateAllDomainRegistries_({ ambiente: options.ambiente || options.environment });
}

function corePessoasV2Diagnostico(options) {
  return core_pessoasV2Diagnostico_(options || {});
}

function corePessoasV2ConferirConsistencia(options) {
  return core_pessoasV2ConferirConsistencia_(options || {});
}

function coreVigenciasV2Diagnostico(options) {
  return core_vigenciasV2Diagnostico_(options || {});
}

function coreVigenciasV2ConferirConsistencia(options) {
  return core_vigenciasV2ConferirConsistencia_(options || {});
}

function coreV2DiagnosticoGeral(options) {
  return core_v2DiagnosticoGeral_(options || {});
}

function corePessoasV2AtualizarResumoOperacional(options) {
  return core_pessoasV2AtualizarResumoOperacional_(options || {});
}

function coreVigenciasV2AtualizarResumoAtual(options) {
  return core_vigenciasV2AtualizarResumoAtual_(options || {});
}

function coreV2RunTesteDiagnosticoGeral() {
  return coreV2RunTesteDiagnosticoGeral_();
}

function coreV2RunTestePessoasResumo() {
  return coreV2RunTestePessoasResumo_();
}

function coreV2RunTesteVigenciasResumo() {
  return coreV2RunTesteVigenciasResumo_();
}

function coreV2_runTesteJobDiarioDryRun() {
  return coreV2_runTesteJobDiarioDryRun_();
}

function coreV2_jobDiarioManutencao(options) {
  return coreV2_jobDiarioManutencao_(options || {});
}

function coreV2InstalarTriggerJobDiario(options) {
  return coreV2_instalarTriggerJobDiario_(options || {});
}

function coreV2RemoverTriggerJobDiario() {
  return coreV2_removerTriggerJobDiario_();
}

function coreV2ListarTriggerJobDiario() {
  return coreV2_listarTriggerJobDiario_();
}

function coreV2_conferirConfiguracao(options) {
  return coreV2_conferirConfiguracao_(options || {});
}

function coreV2_bootstrapConfiguracao(options) {
  return coreV2_bootstrapConfiguracao_(options || {});
}

function coreV2_runTesteBootstrapDryRun() {
  return coreV2_runTesteBootstrapDryRun_();
}

function coreV2_runTesteResolverRegistryV2() {
  return coreV2_runTesteResolverRegistryV2_();
}

function coreCompararLegadoComV2(opts) {
  return coreCompararLegadoComV2_(opts || {});
}

function coreRecalcularVigenciasResumoAtualV2(options) {
  return coreRecalcularVigenciasResumoAtualV2_(options || {});
}

function coreRecalcularPessoasResumoOperacionalV2(options) {
  return coreRecalcularPessoasResumoOperacionalV2_(options || {});
}

function coreRecalcularMembrosDetalhesSemestreAtualV2(options) {
  return coreRecalcularMembrosDetalhesSemestreAtualV2_(options || {});
}

function coreDiagnosticarPessoasResumoOperacionalV2(options) {
  return coreDiagnosticarPessoasResumoOperacionalV2_(options || {});
}

function corePessoasGetById(idPessoa) {
  return corePessoasGetById_(idPessoa);
}

function corePessoasFindByEmail(email) {
  return corePessoasFindByEmail_(email);
}

function corePessoasFindByRga(rga) {
  return corePessoasFindByRga_(rga);
}

function corePessoasGetOperationalSummary(idPessoa) {
  return corePessoasGetOperationalSummary_(idPessoa);
}

function corePessoasListCurrentMembers(opts) {
  return corePessoasListCurrentMembers_(opts || {});
}

function corePessoasListEffectiveMembers(opts) {
  return corePessoasListEffectiveMembers_(opts || {});
}

function corePessoasListExMembers(opts) {
  return corePessoasListExMembers_(opts || {});
}

function corePessoasListWaitingMembers(opts) {
  return corePessoasListWaitingMembers_(opts || {});
}

function corePessoasListAcademicCollaborators(opts) {
  return corePessoasListAcademicCollaborators_(opts || {});
}

function corePessoasListExternalParticipants(opts) {
  return corePessoasListExternalParticipants_(opts || {});
}

function corePessoasListLinksPerfis(idPessoa, opts) {
  return corePessoasListLinksPerfis_(idPessoa, opts || {});
}

function corePessoasListLinksPerfisPublicos(idPessoa) {
  return corePessoasListLinksPerfisPublicos_(idPessoa);
}

function corePessoasV2PrepareLinksPerfis(options) {
  return corePessoasV2PrepareLinksPerfis_(options || {});
}

function corePessoasV2PrepareLinksPerfisReal() {
  return corePessoasV2PrepareLinksPerfisReal_();
}

function corePessoasV2MigrarLinksPerfisLegados(options) {
  return corePessoasV2MigrarLinksPerfisLegados_(options || {});
}

function corePessoasV2MigrarLinksPerfisLegadosReal() {
  return corePessoasV2MigrarLinksPerfisLegadosReal_();
}

function corePessoasV2VerificarRemocaoLinkLattesLegado(options) {
  return corePessoasV2VerificarRemocaoLinkLattesLegado_(options || {});
}

function corePessoasV2RemoverColunaLinkLattesLegado(options) {
  return corePessoasV2RemoverColunaLinkLattesLegado_(options || {});
}

function corePessoasV2RemoverColunaLinkLattesLegadoReal() {
  return corePessoasV2RemoverColunaLinkLattesLegadoReal_();
}

function coreVigenciasGetCurrentFunctionByPessoa(idPessoa) {
  return coreVigenciasGetCurrentFunctionByPessoa_(idPessoa);
}

function coreVigenciasListCurrentFunctions(opts) {
  return coreVigenciasListCurrentFunctions_(opts || {});
}

function coreVigenciasGetPortalPermissionsByPessoa(idPessoa) {
  return coreVigenciasGetPortalPermissionsByPessoa_(idPessoa);
}

/* ============================================================
 * MODULOS_CONFIG
 * ============================================================ */

function coreGetModuleConfig(moduleName, flowName, opts) {
  return core_getModuleConfig_(moduleName, flowName, opts || {});
}

function coreIsModuleEnabled(moduleName, flowName, opts) {
  return core_isModuleEnabled_(moduleName, flowName, opts || {});
}

function coreGetModuleMode(moduleName, flowName, opts) {
  return core_getModuleMode_(moduleName, flowName, opts || {});
}

function coreCanModuleUseCapability(moduleName, flowName, capability, opts) {
  return core_canModuleUseCapability_(moduleName, flowName, capability, opts || {});
}

function coreAssertModuleExecutionAllowed(moduleName, flowName, capability, opts) {
  return core_assertModuleExecutionAllowed_(moduleName, flowName, capability, opts || {});
}

function coreGetModulesConfigDebug() {
  return core_debugModulesConfig_();
}

function coreClearModulesConfigCache() {
  return core_modulesConfigCacheClear_();
}

function coreApplyModulesConfigSheetUx(opts) {
  return core_applyModulesConfigSheetUx_(opts || {});
}

/* ============================================================
 * MODULOS_STATUS
 * ============================================================ */

function coreModuleStatusGet(moduleName, flowName, opts) {
  return core_moduleStatusGet_(moduleName, flowName, opts || {});
}

function coreModuleStatusEnsureRow(moduleName, flowName, opts) {
  return core_moduleStatusEnsureRow_(moduleName, flowName, opts || {});
}

function coreModuleStatusMarkExecution(moduleName, flowName, capability, opts) {
  return core_moduleStatusMarkExecution_(moduleName, flowName, capability, opts || {});
}

function coreModuleStatusMarkSuccess(moduleName, flowName, capability, opts) {
  return core_moduleStatusMarkSuccess_(moduleName, flowName, capability, opts || {});
}

function coreModuleStatusMarkError(moduleName, flowName, errorOrMessage, capability, opts) {
  return core_moduleStatusMarkError_(moduleName, flowName, errorOrMessage, capability, opts || {});
}

function coreModuleStatusMarkBlocked(moduleName, flowName, reasonCode, reasonMessage, capability, modeRead, opts) {
  return core_moduleStatusMarkBlocked_(moduleName, flowName, reasonCode, reasonMessage, capability, modeRead, opts || {});
}

function coreGetModulesStatusDebug() {
  return core_debugModulesStatus_();
}

/* ============================================================
 * REGISTRY / DEBUG
 * ============================================================ */

function coreClearRegistryCache() {
  return core_registryCacheClear_();
}

/* ============================================================
 * SHEETS
 * ============================================================ */

function coreOpenSpreadsheetById(id) {
  return core_openSpreadsheetById_(id);
}

function coreGetSheetById(spreadsheetId, sheetName) {
  return core_getSheetById_(spreadsheetId, sheetName);
}

function coreHeaderMap(sheet, headerRow) {
  return core_headerMap_(sheet, headerRow || 1);
}

function coreGetCol(headerMap, headerName) {
  return core_getCol_(headerMap, headerName);
}

function coreNormalizeHeader(s) {
  return core_normalizeHeader_(s);
}

function coreNormalizeText(value, opts) {
  return core_normalizeText_(value, opts || {});
}

function coreOnlyDigits(value) {
  return core_onlyDigits_(value);
}

function coreBuildHeaderIndexMap(headers, opts) {
  return core_buildHeaderIndexMap_(headers, opts || {});
}

function coreFindHeaderIndex(headerMap, headerName, opts) {
  return core_findHeaderIndex_(headerMap, headerName, opts || {});
}

function coreSetRowValueByHeader(rowArr, headerMap, headerName, value, opts) {
  return core_setRowValueByHeader_(rowArr, headerMap, headerName, value, opts || {});
}

function coreGetCellByHeader(rowArr, headerMap, headerName, opts) {
  return core_getCellByHeader_(rowArr, headerMap, headerName, opts || {});
}

function coreFindFirstExistingHeader(headerMap, headerNames, opts) {
  return core_findFirstExistingHeader_(headerMap, headerNames, opts || {});
}

function coreWriteCellByHeader(sheet, rowNumber, headerMap, headerName, value, opts) {
  return core_writeCellByHeader_(sheet, rowNumber, headerMap, headerName, value, opts || {});
}

function coreFreezeHeaderRow(sheet, headerRow) {
  return core_freezeHeaderRow_(sheet, headerRow || 1);
}

function coreEnsureFilter(sheet, headerRow, opts) {
  return core_ensureFilter_(sheet, headerRow || 1, opts || {});
}

function coreApplyHeaderNotes(sheet, notesByHeader, headerRow) {
  return core_applyHeaderNotes_(sheet, notesByHeader || {}, headerRow || 1);
}

function coreApplyHeaderColors(sheet, groups, headerRow, opts) {
  return core_applyHeaderColors_(sheet, groups || [], headerRow || 1, opts || {});
}

function coreApplyDropdownValidationByHeader(sheet, rulesByHeader, headerRow, opts) {
  return core_applyDropdownValidationByHeader_(sheet, rulesByHeader || {}, headerRow || 1, opts || {});
}

/* ============================================================
 * IDENTITY
 * ============================================================ */

function coreFillMissingProfessorIds() {
  return core_fillMissingProfessorIds_();
}

function coreFillMissingExternalIds() {
  return core_fillMissingExternalIds_();
}

function coreEnsureProfessorIdForRow(rowNumber) {
  return core_ensureProfessorIdForRow_(rowNumber);
}

function coreEnsureExternalIdForRow(rowNumber) {
  return core_ensureExternalIdForRow_(rowNumber);
}

function coreFindExternalByEmail(email) {
  return core_identityFindExternalByEmail_(email);
}

function coreValidateExternalEmailDuplicates() {
  return core_identityValidateExternalEmailDuplicates_();
}

/* ============================================================
 * EX-MEMBROS / COMUNICACOES ABERTAS
 * ============================================================ */

function coreGetExMembersCommunicationRecipients(options) {
  return core_getExMembersCommunicationRecipients_(options || {});
}

function coreDebugExMembersCommunicationRecipients(options) {
  return core_debugExMembersCommunicationRecipients_(options || {});
}

/* ============================================================
 * PORTAL GEAPA
 * ============================================================ */

function geapaCoreBuscarMembroParaPortal(emailOuRga) {
  try {
    return core_buscarMembroParaPortal_(emailOuRga);
  } catch (err) {
    Logger.log('[WARN] geapaCoreBuscarMembroParaPortal: falha interna ao consultar membro para portal.');
    return null;
  }
}

function geapaCoreBuscarUsuarioPortal(emailOuRga) {
  try {
    return core_buscarUsuarioPortal_(emailOuRga);
  } catch (err) {
    Logger.log('[WARN] geapaCoreBuscarUsuarioPortal: falha interna ao consultar usuario para portal.');
    return core_buildPortalError_(
      'ERRO_BUSCAR_USUARIO_PORTAL',
      'Nao foi possivel buscar o usuario do portal.'
    );
  }
}

function geapaCoreBuscarMinhaSituacaoParaPortal(emailOuRga, options) {
  try {
    return core_buscarMinhaSituacaoParaPortal_(emailOuRga, options || {});
  } catch (err) {
    Logger.log('[WARN] geapaCoreBuscarMinhaSituacaoParaPortal ' + JSON.stringify({
      errorCode: String(err && err.code || 'ERRO_BUSCAR_MINHA_SITUACAO'),
      failedStage: 'core_buscarMinhaSituacaoParaPortal',
      traceId: String(options && (options.traceId || options.requestId) || '').slice(0, 80)
    }));
    return core_buildPortalError_(
      String(err && err.code || 'ERRO_BUSCAR_MINHA_SITUACAO'),
      'Nao foi possivel buscar a situacao do membro.'
    );
  }
}

function geapaCoreBuscarMeuPerfilParaPortal(emailOuRga, options) {
  try {
    return core_buscarMeuPerfilParaPortal_(emailOuRga, options || {});
  } catch (err) {
    Logger.log('[WARN] geapaCoreBuscarMeuPerfilParaPortal: falha interna ao consultar perfil do usuario para portal.');
    return core_buildPortalError_(
      'ERRO_BUSCAR_MEU_PERFIL',
      'Nao foi possivel buscar o perfil do usuario.'
    );
  }
}

/** Atualiza somente campos cadastrais de baixo risco do usuario autenticado. */
function geapaCoreAtualizarMeuPerfilParaPortal(payload, contexto) {
  return core_atualizarMeuPerfilParaPortal_(payload || {}, contexto || {});
}

/** Registra uma solicitacao, sem alterar diretamente campos sensiveis. */
function geapaCoreSolicitarCorrecaoMeuPerfilParaPortal(payload, contexto) {
  return core_solicitarCorrecaoMeuPerfilParaPortal_(payload || {}, contexto || {});
}

/** Lista exclusivamente as solicitacoes da pessoa autenticada. */
function geapaCoreListarMinhasSolicitacoesCadastraisPortal(contexto) {
  return core_listarMinhasSolicitacoesCadastraisPortal_(contexto || {});
}

/** Confirma uma solicitacao propria por ID ou chave idempotente, sem reenviar a escrita. */
function geapaCoreConsultarMinhaSolicitacaoCadastralPortal(consulta, contexto) {
  return core_consultarMinhaSolicitacaoCadastralPortal_(consulta || {}, contexto || {});
}

/** Lista solicitacoes para Secretaria/Diretoria autorizada no backend. */
function geapaCoreListarSolicitacoesCadastraisAdministracaoPortal(filtros, contexto) {
  return core_listarSolicitacoesCadastraisAdministracaoPortal_(filtros || {}, contexto || {});
}

/** Retorna um unico detalhe administrativo, com revelacao protegida e auditada. */
function geapaCoreDetalharSolicitacaoCadastralAdministracaoPortal(payload, contexto) {
  return core_detalharSolicitacaoCadastralAdministracaoPortal_(payload || {}, contexto || {});
}

/** Registra somente estados administrativos que nao aplicam alteracoes. */
function geapaCoreAnalisarSolicitacaoCadastralPortal(payload, contexto) {
  return core_analisarSolicitacaoCadastralPortal_(payload || {}, contexto || {});
}

/** Aprova e aplica a correcao sensivel em uma unica operacao coordenada. */
function geapaCoreAprovarEAplicarSolicitacaoCadastralPortal(payload, contexto) {
  return core_aprovarEAplicarSolicitacaoCadastralPortal_(payload || {}, contexto || {});
}

/** Alias protegido para solicitacoes legadas previamente aprovadas. */
function geapaCoreAplicarSolicitacaoCadastralAprovadaPortal(payload, contexto) {
  return core_aplicarSolicitacaoCadastralAprovadaPortal_(payload || {}, contexto || {});
}

/** Setup idempotente e protegido da base DEV usada pelo projeto HOMOLOG. */
function geapaCoreSetupSolicitacoesAtualizacaoCadastralDev(options) {
  return core_setupSolicitacoesAtualizacaoCadastralDev_(options || {});
}

/** Entrada manual do editor para o setup real, ainda recusada quando GEAPA_ENV=PROD. */
function geapaCoreSetupSolicitacoesAtualizacaoCadastralDevReal() {
  return core_setupSolicitacoesAtualizacaoCadastralDevReal_();
}

/** Diagnostica ou prepara explicitamente a base PROD, sem depender de GEAPA_ENV. */
function geapaCoreSetupSolicitacoesAtualizacaoCadastralProd(options) {
  return core_setupSolicitacoesAtualizacaoCadastralProd_(options || {});
}

/** Entrada manual de escrita PROD, protegida por confirmacao dedicada. */
function geapaCoreSetupSolicitacoesAtualizacaoCadastralProdReal() {
  return core_setupSolicitacoesAtualizacaoCadastralProdReal_();
}

/** Catalogos oficiais versionados de paises, UFs e municipios. */
function geapaCoreListarCatalogoLocalidadesPortal(options) {
  return core_getLocalityCatalogV2_(options || {});
}

/** Validacao backend do conjunto de origem/naturalidade. */
function geapaCoreValidarOrigemV2(payload) {
  return core_validateOriginV2_(payload || {});
}

/** Validacao de curso contra registros previamente lidos de CURSOS_CATALOGO. */
function geapaCoreValidarCursoV2(payload, registrosCursos) {
  return core_validateCourseV2_(payload || {}, registrosCursos || []);
}

/** Calculo cronologico baseado exclusivamente nos semestres institucionais informados. */
function geapaCoreCalcularSemestreCursoV2(rga, registrosSemestres, dataReferencia) {
  return core_calculateAcademicSemesterV2_(rga, registrosSemestres || [], dataReferencia);
}

/** Planeja por padrao; escrita real aceita somente DEV e exige token explicito. */
function geapaCoreSetupEvolucaoMembrosV2Dev(options) {
  return core_setupMemberEvolutionV2_(options || {});
}

/** Lista membros administrativos para o Portal com autorizacao e sanitizacao no Core. */
function geapaCoreListarMembrosAdministracaoPortal(filtros, contexto) {
  try {
    return core_listarMembrosAdministracaoPortal_(filtros || {}, contexto || {});
  } catch (err) {
    Logger.log('[WARN] geapaCoreListarMembrosAdministracaoPortal: falha interna sem payload sensivel.');
    return {
      ok: false,
      errorCode: 'MEMBROS_ADMIN_INDISPONIVEIS',
      message: 'Nao foi possivel consultar os membros neste momento.'
    };
  }
}

function geapaCoreListarMembrosParaChamada(dataAtividade, contexto) {
  try {
    return core_listarMembrosParaChamada_(dataAtividade, contexto || {});
  } catch (err) {
    Logger.log('[WARN] geapaCoreListarMembrosParaChamada: falha interna ao listar membros para chamada.');
    return core_buildChamadaError_(
      'ERRO_LISTAR_MEMBROS_CHAMADA',
      'Nao foi possivel listar membros para chamada.'
    );
  }
}

function coreListarMembrosParaChamada(dataAtividade, contexto) {
  return geapaCoreListarMembrosParaChamada(dataAtividade, contexto || {});
}

function listarMembrosParaChamada(dataAtividade, contexto) {
  return geapaCoreListarMembrosParaChamada(dataAtividade, contexto || {});
}

function geapaCoreInvalidarCacheMembrosChamada(dataAtividade) {
  return core_invalidarCacheMembrosChamada_(dataAtividade);
}

function coreInvalidarCacheMembrosChamada(dataAtividade) {
  return geapaCoreInvalidarCacheMembrosChamada(dataAtividade);
}

function geapaCoreRunTesteUsuarioPortal() {
  try {
    var identificador = PropertiesService
      .getScriptProperties()
      .getProperty('GEAPA_CORE_PORTAL_TESTE_IDENTIFICADOR');

    if (!String(identificador || '').trim()) {
      return core_buildPortalError_(
        'CONFIG_TESTE_AUSENTE',
        'Configure a Script Property GEAPA_CORE_PORTAL_TESTE_IDENTIFICADOR para executar este teste.'
      );
    }

    return core_buildUsuarioPortalTesteResumo_(
      geapaCoreBuscarUsuarioPortal(identificador)
    );
  } catch (err) {
    Logger.log('[WARN] geapaCoreRunTesteUsuarioPortal: falha interna no teste manual de usuario do portal.');
    return core_buildPortalError_(
      'ERRO_TESTE_USUARIO_PORTAL',
      'Nao foi possivel executar o teste de usuario do portal.'
    );
  }
}

function geapaCoreRunTesteListarMembrosParaChamada() {
  try {
    return core_runTesteListarMembrosParaChamada_();
  } catch (err) {
    Logger.log('[WARN] geapaCoreRunTesteListarMembrosParaChamada: falha interna no teste manual de chamada.');
    return core_buildChamadaError_(
      'ERRO_TESTE_MEMBROS_CHAMADA',
      'Nao foi possivel executar o teste de membros para chamada.'
    );
  }
}

function geapaCoreRunTesteMinhaSituacaoParaPortal() {
  try {
    var identificador = PropertiesService
      .getScriptProperties()
      .getProperty('GEAPA_CORE_PORTAL_TESTE_IDENTIFICADOR');

    if (!String(identificador || '').trim()) {
      return core_buildPortalError_(
        'CONFIG_TESTE_AUSENTE',
        'Configure a Script Property GEAPA_CORE_PORTAL_TESTE_IDENTIFICADOR para executar este teste.'
      );
    }

    return geapaCoreBuscarMinhaSituacaoParaPortal(identificador);
  } catch (err) {
    Logger.log('[WARN] geapaCoreRunTesteMinhaSituacaoParaPortal: falha interna no teste manual do portal.');
    return core_buildPortalError_(
      'ERRO_TESTE_MINHA_SITUACAO',
      'Nao foi possivel executar o teste de Minha situacao do portal.'
    );
  }
}

/* ============================================================
 * PORTAL GEAPA / ACESSO, PERFIS E PERMISSOES
 * ============================================================ */

function corePortalGetConfig() {
  return corePortalReadConfig_();
}

function corePortalGetOperationalConfig(opts) {
  return corePortalReadConfig_(opts || {});
}

function corePortalClearConfigCache() {
  return corePortalConfigCacheClear_();
}

function corePortalGetProfiles() {
  return corePortalReadProfiles_();
}

function corePortalGetPermissionsByProfile(perfilPortal) {
  return corePortalBuildPermissionsForProfile_(perfilPortal);
}

function corePortalAuthorizeEmail(email, opts) {
  return corePortalAuthorizeEmail_(email, opts || {});
}

function corePortalHasPermission(sessionOrEmail, permission, opts) {
  return corePortalHasPermission_(sessionOrEmail, permission, opts || {});
}

function corePortalLogAccess(payload) {
  return corePortalAppendAccessLog_(payload || {});
}

function corePortalDiagnostics() {
  return corePortalDiagnostics_();
}

function corePortalResolverUsuarioAtual(entrada, opts) {
  return corePortalResolverUsuarioAtual_(entrada, opts || {});
}

function corePortalBuildFirestoreUserSnapshot(entrada, opts) {
  return corePortalBuildFirestoreUserSnapshot_(entrada, opts || {});
}

function corePortalGerarSnapshotFirestoreUsuario(entrada, opts) {
  return corePortalBuildFirestoreUserSnapshot_(entrada, opts || {});
}

function corePortalSincronizarUsuarioFirestore(entrada, opts) {
  return corePortalSincronizarUsuarioFirestore_(entrada, opts || {});
}

function corePortalProvisionarFirestoreUserAutenticado(firebaseIdentity, opts) {
  return corePortalProvisionarFirestoreUserAutenticado_(firebaseIdentity || {}, opts || {});
}

function corePortalMarcarFirestoreUserInativoPorUid(uid, opts) {
  return corePortalMarcarFirestoreUserInativoPorUid_(uid, opts || {});
}

function corePortalInvalidarCacheFirestoreUsuario(idPessoaOuEmail, opts) {
  return corePortalInvalidarCacheFirestoreUsuario_(idPessoaOuEmail, opts || {});
}

function corePortalSyncFirestoreUserByEmail(email, opts) {
  return corePortalSyncFirestoreUserByEmail_(email, opts || {});
}

function corePortalSyncFirestoreUserByIdPessoa(idPessoa, opts) {
  return corePortalSyncFirestoreUserByIdPessoa_(idPessoa, opts || {});
}

function corePortalSyncFirestoreUsersFromPessoasV2(opts) {
  return corePortalSyncFirestoreUsersFromPessoasV2_(opts || {});
}

function corePortalDiagnosticarFirestoreUsersDev(opts) {
  return corePortalDiagnosticarFirestoreUsersDev_(opts || {});
}

function coreFirestoreSetDocument(path, data, options) {
  return coreFirestoreSetDocument_(path, data || {}, options || {});
}

function coreFirestoreGetDocument(path, options) {
  return coreFirestoreGetDocument_(path, options || {});
}

function coreFirestoreListDocuments(collectionPath, options) {
  return coreFirestoreListDocuments_(collectionPath, options || {});
}

function coreFirestoreDeleteDocument(path, options) {
  return coreFirestoreDeleteDocument_(path, options || {});
}

function coreFirestoreBatchSetDocuments(items, options) {
  return coreFirestoreBatchSetDocuments_(items || [], options || {});
}

function coreFirestoreDiagnosticar(options) {
  return coreFirestoreDiagnosticar_(options || {});
}

function corePortalCalcularPerfilEfetivo(idPessoa, opts) {
  return corePortalCalcularPerfilEfetivo_(idPessoa, opts || {});
}

function corePortalListarPermissoesEfetivas(idPessoa, opts) {
  return corePortalListarPermissoesEfetivas_(idPessoa, opts || {});
}

function corePortalValidarAcesso(idPessoa, permissaoOuPerfil, opts) {
  return corePortalValidarAcesso_(idPessoa, permissaoOuPerfil, opts || {});
}

function corePortalGetMeuResumo(email, opts) {
  return corePortalGetMeuResumo_(email, opts || {});
}

function corePortalListarApresentacoesPermitidas(email, options) {
  return corePortalListarApresentacoesPermitidas_(email, options || {});
}

function corePortalListarApresentacoesParaEgresso(idPessoa, opts) {
  return corePortalListarApresentacoesParaEgresso_(idPessoa, opts || {});
}

function corePortalDiagnosticarPerfisEPermissoes(opts) {
  return corePortalDiagnosticarPerfisEPermissoes_(opts || {});
}

function corePortalDiagnosticarAcessoPortalDev(opts) {
  return corePortalDiagnosticarAcessoPortalDev_(opts || {});
}

function geapaCore_diagnosticarAcessoPortalDev_(opts) {
  return corePortalDiagnosticarAcessoPortalDev_(opts || {});
}

function corePrepararPortalParaV2(opts) {
  return corePrepararPortalParaV2_(opts || {});
}

/* ============================================================
 * PORTAL GEAPA / CONTEUDO PUBLICO EDITORIAL
 * ============================================================ */

function corePortalPublicContentGetDefinitions() {
  return corePortalPublicContentGetDefinitions_();
}

function corePortalPublicContentEnsureStructure(options) {
  return corePortalPublicContentEnsureStructure_(options || {});
}

function corePortalPublicContentCreateSpreadsheet(options) {
  return corePortalPublicContentCreateSpreadsheet_(options || {});
}

function corePortalPublicContentEnsureSheets(options) {
  return corePortalPublicContentEnsureSheets_(options || {});
}

function corePortalPublicContentEnsureHeaders(options) {
  return corePortalPublicContentEnsureHeaders_(options || {});
}

function corePortalPublicContentDiagnostics(options) {
  return corePortalPublicContentDiagnostics_(options || {});
}

function corePortalPublicContentReadRows(key, options) {
  return corePortalPublicContentReadRows_(key, options || {});
}

function corePortalPublicContentGetPage(slug, options) {
  return corePortalPublicContentGetPage_(slug, options || {});
}

function corePortalPublicContentGetHome(options) {
  return corePortalPublicContentGetHome_(options || {});
}

function corePortalPublicContentGetSobre(options) {
  return corePortalPublicContentGetSobre_(options || {});
}

function corePortalPublicContentGetHistoria(options) {
  return corePortalPublicContentGetHistoria_(options || {});
}

function corePortalPublicContentGetParceiros(options) {
  return corePortalPublicContentGetParceiros_(options || {});
}

function corePortalPublicContentGetDocumentos(options) {
  return corePortalPublicContentGetDocumentos_(options || {});
}

function corePortalPublicContentGetConfig(options) {
  return corePortalPublicContentGetConfig_(options || {});
}

function corePortalPublicContentGetMidias(options) {
  return corePortalPublicContentGetMidias_(options || {});
}

function corePortalPublicContentGetDiretoriaComplementos(options) {
  return corePortalPublicContentGetDiretoriaComplementos_(options || {});
}

function corePortalPublicContentGetPessoasComplementos(options) {
  return corePortalPublicContentGetPessoasComplementos_(options || {});
}

function corePortalPublicContentGetGestoesComplementos(options) {
  return corePortalPublicContentGetGestoesComplementos_(options || {});
}

function corePortalPublicContentGetPessoasConfig(options) {
  return corePortalPublicContentGetPessoasConfig_(options || {});
}

function corePortalPublicContentBuildPublicSnapshot(options) {
  return corePortalPublicContentBuildPublicSnapshot_(options || {});
}

function corePortalPublicContentBuildPublicBoard(options) {
  return corePortalPublicContentBuildPublicBoard_(options || {});
}

/* ============================================================
 * DATES
 * ============================================================ */

function coreNow() {
  return core_now_();
}

function coreStartOfDay(date) {
  return core_startOfDay_(date);
}

function coreAddDays(date, days) {
  return core_addDays_(date, days);
}

function coreIsSameDay(d1, d2) {
  return core_isSameDay_(d1, d2);
}

function coreInWindowDay(date, startInclusive, endExclusive) {
  return core_inWindowDay_(date, startInclusive, endExclusive);
}

function coreFormatDate(date, tz, pattern) {
  return core_formatDate_(date, tz, pattern);
}

/* ============================================================
 * EMAIL
 * ============================================================ */

function coreIsValidEmail(email) {
  return core_isValidEmail_(email);
}

function coreNormalizeEmail(value) {
  return core_normalizeEmail_(value);
}

function coreSendEmailText(opts) {
  return core_sendEmailText_(opts);
}

function coreSendEmailHtml(opts) {
  return core_sendEmailHtml_(opts);
}

function coreSendTrackedEmail(params) {
  return core_sendTrackedEmail_(params);
}

function coreExtractEmailAddress(value) {
  return core_extractEmailAddress_(value);
}

function coreExtractDisplayName(value) {
  return core_extractDisplayName_(value);
}

function coreUniqueEmails(values) {
  return core_uniqueEmails_(values);
}

/* ============================================================
 * GMAIL (labels)
 * ============================================================ */

function coreEnsureLabel(name) {
  return core_ensureLabel_(name);
}

function coreGetLabel(name) {
  return core_getLabel_(name);
}

function coreGetOrCreateLabel(name) {
  return core_getOrCreateLabel_(name);
}

function coreThreadHasLabel(thread, labelName) {
  return core_threadHasLabel_(thread, labelName);
}

function coreSearchThreads(query, start, max) {
  return core_searchThreads_(query, start, max);
}

function coreMarkThread(thread, labelIn, labelOut) {
  return core_markThread_(thread, labelIn, labelOut);
}

function coreReplyThreadHtml(thread, subject, htmlBody, opts) {
  return core_replyThreadHtml_(thread, subject, htmlBody, opts || {});
}

/* ============================================================
 * MAIL HUB
 * ============================================================ */

function coreMailIngestInbox(opts) {
  return core_mailIngestInbox_(opts || {});
}

function coreMailGetConfig(key, defaultValue) {
  return coreMailHubGetConfig_(key, defaultValue);
}

function coreMailGetConfigBoolean(key, defaultValue) {
  return coreMailHubGetConfigBooleanByKey_(key, defaultValue === true);
}

function coreMailGetConfigList(key) {
  return coreMailHubGetConfigListByKey_(key);
}

function coreMailListPendingByModule(moduleName) {
  return core_mailListPendingByModule_(moduleName);
}

function coreMailRegisterModuleAdapter(adapter) {
  return coreMailRegisterModuleAdapter_(adapter);
}

function coreMailGetModuleAdapter(moduleCodeOrName) {
  return coreMailAdapterToSnapshot_(coreMailGetModuleAdapter_(moduleCodeOrName));
}

function coreMailListModuleAdapters() {
  return coreMailListModuleAdapters_().map(function(adapter) {
    return coreMailAdapterToSnapshot_(adapter);
  });
}

function coreMailBuildCorrelationKey(moduleCodeOrName, ctx) {
  return coreMailBuildCorrelationKey_(moduleCodeOrName, ctx || {});
}

function coreMailParseCorrelationKey(key) {
  return coreMailParseCorrelationKey_(key);
}

function coreMailResolveRouting(msgCtx) {
  return coreMailResolveRouting_(msgCtx || {});
}

function coreMailNormalizeOutgoingSubject(moduleCodeOrName, subject, ctx) {
  return coreMailNormalizeOutgoingSubject_(moduleCodeOrName, subject, ctx || {});
}

function coreMailRenderEmailTemplate(templateKey, subjectHuman, payload) {
  return coreMailRenderEmailTemplate_(templateKey, subjectHuman, payload || {});
}

function coreMailBuildFinalSubject(subjectHuman, correlationKey) {
  return coreMailBuildFinalSubject_(subjectHuman, correlationKey);
}

function coreMailBuildOutgoingDraft(contract) {
  return coreMailBuildOutgoingDraft_(contract || {});
}

function coreMailQueueOutgoing(contract) {
  return coreMailQueueOutgoing_(contract || {});
}

function coreMailProcessOutbox() {
  return coreMailProcessOutbox_();
}

function coreMailGetLatestEvent(opts) {
  return core_mailGetLatestEvent_(opts || {});
}

function coreMailListAttachments(opts) {
  return core_mailListAttachments_(opts || {});
}

function coreMailListPendingAttachments(opts) {
  return core_mailListPendingAttachments_(opts || {});
}

function coreMailGetLatestPendingEventWithAttachment(opts) {
  return core_mailGetLatestPendingEventWithAttachment_(opts || {});
}

function coreMailListAttachmentsByEvent(eventId, opts) {
  return core_mailListAttachmentsByEvent_(eventId, opts || {});
}

function coreMailGetAttachmentById(attachmentId, opts) {
  return core_mailGetAttachmentById_(attachmentId, opts || {});
}

function coreMailGetAttachmentsByEvent(eventId, opts) {
  return core_mailGetAttachmentsByEvent_(eventId, opts || {});
}

function coreMailMarkLatestPendingByModule(moduleName, processorName) {
  return core_mailMarkLatestPendingByModule_(moduleName, processorName);
}

function coreMailMarkEventProcessed(eventId, processorName) {
  return core_mailMarkEventProcessed_(eventId, processorName);
}

function coreMailMarkAttachmentProcessed(attachmentId, processorName, observations) {
  return core_mailMarkAttachmentProcessed_(attachmentId, processorName, observations || '');
}

function coreMailMarkAttachmentSavedToDrive(attachmentId, processorName, driveInfo) {
  return core_mailMarkAttachmentSavedToDrive_(attachmentId, processorName, driveInfo || {});
}

function coreMailMarkAttachmentIgnored(attachmentId, processorName, observations) {
  return core_mailMarkAttachmentIgnored_(attachmentId, processorName, observations || '');
}

function coreMailMarkAttachmentError(attachmentId, processorName, observations) {
  return core_mailMarkAttachmentError_(attachmentId, processorName, observations || '');
}

function coreMailCleanupNoiseEvents() {
  return coreMailCleanupNoiseEvents_();
}

function coreMailApplyOperationalSheetUx(opts) {
  return coreMailApplyOperationalSheetUx_(opts || {});
}

/* ============================================================
 * LOGS
 * ============================================================ */

function coreRunId() {
  return core_runId_();
}

function coreLogInfo(runId, msg, obj) {
  return core_logInfo_(runId, msg, obj);
}

function coreLogWarn(runId, msg, obj) {
  return core_logWarn_(runId, msg, obj);
}

function coreLogError(runId, msg, obj) {
  return core_logError_(runId, msg, obj);
}

function coreLogSummarize(runId, title, startedAt, counters) {
  return core_logSummarize_(runId, title, startedAt, counters || {});
}

/* ============================================================
 * ASSERT
 * ============================================================ */

function coreAssertRequired(value, msg) {
  return core_assertRequired_(value, msg);
}

/* ============================================================
 * DRIVE
 * ============================================================ */

function coreDriveGetFolderById(folderId) {
  return core_driveGetFolderById_(folderId);
}

function coreDriveGetFileById(fileId) {
  return core_driveGetFileById_(fileId);
}

function coreDriveEnsureFolder(parentFolderId, childName) {
  return core_driveEnsureFolder_(parentFolderId, childName);
}

function coreDriveMoveFileToFolder(fileId, folderId) {
  return core_driveMoveFileToFolder_(fileId, folderId);
}

function coreDriveListFiles(folderId, max) {
  return core_driveListFiles_(folderId, max || 200);
}

function coreAppendMemberLifecycleEvent(payload) {
  return core_appendMemberLifecycleEvent_(payload || {});
}

function coreListMemberLifecycleEvents(filters, opts) {
  return core_memberLifecycleListEvents_(filters || {}, opts || {});
}

function coreGetLatestMemberLifecycleEventByRga(rga, opts) {
  return core_memberLifecycleGetLatestEventByRga_(rga, opts || {});
}

function coreUpdateMemberLifecycleEvent(eventId, patch) {
  return core_updateMemberLifecycleEvent_(eventId, patch || {});
}

function coreUpdateMemberLifecycleEventStatus(eventId, nextStatus, opts) {
  return core_updateMemberLifecycleEventStatus_(eventId, nextStatus, opts || {});
}

/* ============================================================
 * HTTP
 * ============================================================ */

function coreHttpPostJson(opts) {
  return core_httpPostJson_(opts);
}

/* ============================================================
 * ASSETS / EMAIL COM IMAGENS
 * ============================================================ */

function coreGetAssetBlob(assetIdOrUrl) {
  return coreGetAssetBlob_(assetIdOrUrl);
}

function coreInlineImagesDefault() {
  return core_inlineImagesDefault_();
}

function coreSendHtmlEmail(opts) {
  // envia HTML + logo padrão + inlineImages adicionais do módulo
  return core_sendEmailHtmlWithDefaultInline_(opts);
}

/* ============================================================
 * LOCK (Library limitation)
 * ============================================================ *
 * Callback NÃO atravessa library, então não exportamos execução com fn().
 * Módulos devem usar LockService localmente.
 */
function coreWithLock() {
  throw new Error("coreWithLock não suportado via Library. Use LockService no módulo.");
}

/* ============================================================
 * GOVERNANCE / VIGÊNCIAS
 * ============================================================ */
function coreGetCurrentBoard(refDate) {
  return core_getCurrentBoard_(refDate);
}

function coreGetCurrentBoardSlogan(refDate) {
  return core_getCurrentBoardSlogan_(refDate);
}

function coreGetCurrentBoardMembers(refDate) {
  return core_getCurrentBoardMembers_(refDate);
}

function coreGetCurrentBoardMembersByOccupation(occupation, refDate) {
  return core_getCurrentBoardMembersByOccupation_(occupation, refDate);
}

function coreGetCurrentBoardMembersByRole(role, refDate) {
  return core_getCurrentBoardMembersByRole_(role, refDate);
}

function coreGetCurrentBoardMemberByOccupation(occupation, refDate) {
  return core_getCurrentBoardMemberByOccupation_(occupation, refDate);
}

function coreGetCurrentBoardMemberByRole(role, refDate) {
  return core_getCurrentBoardMemberByRole_(role, refDate);
}

function coreGetCurrentLeadership(refDate) {
  return core_getCurrentLeadership_(refDate);
}

/* ============================================================
 * SEMESTRES / RGA
 * ============================================================ */
function coreGetCurrentSemester(refDate) {
  return core_getCurrentSemester_(refDate);
}

function coreParseEntrySemesterFromRga(rga) {
  return core_parseEntrySemesterFromRga_(rga);
}

function coreGetStudentCurrentSemesterFromRga(rga, refDate) {
  return core_getStudentCurrentSemesterFromRga_(rga, refDate);
}

/* ============================================================
 * SEMESTRES / MEMBERS_ATUAIS
 * ============================================================ */
function coreGetSemesterForDate(refDate) {
  return core_getSemesterForDate_(refDate);
}

function coreGetSemesterIdForDate(refDate) {
  return core_getSemesterIdForDate_(refDate);
}

function coreGetLastCompletedSemester(refDate) {
  return core_getLastCompletedSemester_(refDate);
}

function coreGetCompletedGroupSemesterCountFromEntrySemester(entrySemesterShort, refDate) {
  return core_getCompletedGroupSemesterCountFromEntrySemester_(entrySemesterShort, refDate);
}

function coreSyncMembersCurrentDerivedFields() {
  return core_syncMembersCurrentDerivedFields_();
}


/* ============================================================
 * CARGOS INSTITUCIONAIS / CONFIG
 * ============================================================ */

function coreGetInstitutionalRolesActive() {
  return core_getInstitutionalRolesActive_();
}

function coreFindInstitutionalRoleByKey(cargoKey) {
  return core_findInstitutionalRoleByKey_(cargoKey);
}

function coreFindInstitutionalRoleByPublicName(publicName) {
  return core_findInstitutionalRoleByPublicName_(publicName);
}

function coreFindInstitutionalRoleByAnyName(text) {
  return core_findInstitutionalRoleByAnyName_(text);
}

function coreGetInstitutionalRolesByEmailGroup(groupName) {
  return core_getInstitutionalRolesByEmailGroup_(groupName);
}

function coreDebugInstitutionalRolesConfig() {
  return core_debugInstitutionalRolesConfig_();
}

function coreClearInstitutionalRolesConfigCache() {
  return core_rolesConfigCacheClear_();
}

/* ============================================================
 * CARGOS INSTITUCIONAIS / PROJEÇÃO ATUAL
 * ============================================================ */

function coreGetCurrentInstitutionalAssignments(refDate) {
  return core_getCurrentInstitutionalAssignments_(refDate);
}

function coreGetCurrentOccupantsByEmailGroup(groupName, refDate) {
  return core_getCurrentOccupantsByEmailGroup_(groupName, refDate);
}

function coreGetCurrentContactsHtmlByEmailGroup(groupName, refDate) {
  return core_getCurrentContactsHtmlByEmailGroup_(groupName, refDate);
}

function coreSyncMembersCurrentInstitutionalRoles(refDate) {
  return core_syncMembersCurrentInstitutionalRoles_(refDate);
}

function coreSyncMembersCurrentInstitutionalOccupations(refDate) {
  return core_syncMembersCurrentInstitutionalOccupations_(refDate);
}

function coreDebugCurrentInstitutionalProjection(refDate) {
  return core_debugCurrentInstitutionalProjection_(refDate);
}

function coreGetCurrentEmailsByEmailGroup(groupName, refDate) {
  return core_getCurrentEmailsByEmailGroup_(groupName, refDate);
}

function coreGetCurrentEmailsByRole(roleName, refDate) {
  return core_getCurrentEmailsByRole_(roleName, refDate);
}

function coreGetCurrentEmailsByOccupation(occupationName, refDate) {
  return core_getCurrentEmailsByOccupation_(occupationName, refDate);
}

/* ============================================================
 * IDENTIDADE DE MEMBRO / AUTOFILL
 * ============================================================ */

function coreFindMemberIdentityByAny(identity) {
  return core_memberIdentityFindByAny_(identity);
}

function coreNormalizeIdentityKey(value) {
  return core_normalizeIdentityKey_(value);
}

function coreFindMemberCurrentRowByAny(identity) {
  return core_findMemberCurrentRowByAny_(identity);
}

function coreAutofillIdentityRowInSheet(sheet, rowNumber, opts) {
  return core_autofillIdentityRowInSheet_(sheet, rowNumber, opts || {});
}
