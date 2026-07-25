/**
 * ============================================================
 * 31_core_domains_v2_operational_api.js
 * ============================================================
 *
 * APIs operacionais permanentes para consumo de Pessoas v2 e Vigencias v2.
 *
 * Regras:
 * - nao executa migracao legado -> v2;
 * - nao altera planilhas legadas;
 * - leituras retornam objetos derivados das bases v2;
 * - recalculos escrevem apenas caches v2 e exigem confirmacao explicita.
 */

function core_domainsV2NewReadReport_(name) {
  return core_domainsV2AuditNewReport_(name || 'DOMINIOS_V2');
}

function core_domainsV2OpenPessoas_(report, options) {
  return core_domainsV2AuditOpenDomain_(
    'PESSOAS',
    report || core_domainsV2NewReadReport_('PESSOAS_V2'),
    options || {}
  );
}

function core_domainsV2OpenVigencias_(report, options) {
  return core_domainsV2AuditOpenDomain_(
    'VIGENCIAS',
    report || core_domainsV2NewReadReport_('VIGENCIAS_V2'),
    options || {}
  );
}

/*
 * Cache estritamente por execucao. Diferentemente do CacheService, este mapa
 * nunca sobrevive a uma execucao Apps Script e, portanto, nao compartilha
 * registros pessoais entre requisicoes.
 */
var __core_domains_v2_records_execution_cache = {};

function core_domainsV2ResetRecordsExecutionCache_() {
  __core_domains_v2_records_execution_cache = {};
}

/**
 * Le somente as abas logicas necessarias ao contrato atual.
 *
 * O leitor historico de auditoria abre todas as abas do dominio. Isso e
 * apropriado para diagnosticos, mas muito caro para login e telas pontuais.
 */
function core_domainsV2OpenSubset_(domain, logicalSheets, report, options) {
  var opts = options || {};
  var environment = core_normalizeDomainEnv_(opts);
  var definition = core_getDomainDefinition_(domain);
  var output = {};

  (logicalSheets || []).forEach(function(logicalSheet) {
    var logical = String(logicalSheet || '').trim().toUpperCase();
    var canonicalName = definition.definition.sheets[logical];
    if (!canonicalName || output[canonicalName]) return;

    try {
      var resolved = core_getDomainSheet_(definition.key, logical, Object.assign({}, opts, {
        ambiente: environment,
        includeResolution: true
      }));
      var resolution = resolved.resolution || {};
      var cacheKey = [
        environment,
        definition.key,
        logical,
        String(resolution.spreadsheetId || ''),
        String(resolution.sheetName || canonicalName)
      ].join('|');
      var cached = __core_domains_v2_records_execution_cache[cacheKey];

      if (!cached) {
        var startedAt = Date.now();
        cached = {
          sheet: resolved.sheet,
          records: core_readSheetRecords_(resolved.sheet, { skipBlankRows: true }),
          headers: resolved.sheet.getLastColumn() > 0
            ? resolved.sheet.getRange(1, 1, 1, resolved.sheet.getLastColumn()).getDisplayValues()[0]
            : [],
          resolution: resolution,
          durationMs: Math.max(0, Date.now() - startedAt)
        };
        __core_domains_v2_records_execution_cache[cacheKey] = cached;
        if (typeof corePortalTraceStage_ === 'function') {
          corePortalTraceStage_(
            definition.key + '.' + logical + '.leitura',
            startedAt,
            'READ',
            String(resolution.origin || '')
          );
        }
      } else if (typeof corePortalTraceStage_ === 'function') {
        corePortalTraceStage_(
          definition.key + '.' + logical + '.leitura',
          Date.now(),
          'CACHE_EXECUCAO',
          String(resolution.origin || '')
        );
      }

      output[canonicalName] = cached;
    } catch (error) {
      core_domainsV2AuditIssue_(
        report,
        'ERRO',
        error && error.code ? error.code : 'DOMINIO_ABA_INDISPONIVEL',
        'Aba necessaria ao contrato operacional esta indisponivel.',
        {
          domain: definition.key,
          logicalSheet: logical,
          ambiente: environment,
          error: String(error && error.message || error || '').slice(0, 300)
        }
      );
      output[canonicalName] = { sheet: null, records: [], headers: [], resolution: null, durationMs: 0 };
    }
  });

  report.ambiente = environment;
  return output;
}

function core_domainsV2OpenPessoasSubset_(logicalSheets, report, options) {
  return core_domainsV2OpenSubset_(
    'PESSOAS',
    logicalSheets,
    report || core_domainsV2NewReadReport_('PESSOAS_V2_SUBSET'),
    options || {}
  );
}

function core_domainsV2CloneRecord_(record) {
  var out = {};
  Object.keys(record || {}).forEach(function(key) {
    if (key !== '__rowNumber') out[key] = record[key];
  });
  return out;
}

function core_domainsV2Email_(value) {
  return String(value || '').trim().toLowerCase();
}

function core_domainsV2Rga_(value) {
  return String(value || '').trim();
}

var CORE_PESSOAS_V2_LINKS_PERFIS_TYPES = Object.freeze([
  'LATTES',
  'LINKEDIN',
  'ORCID',
  'INSTAGRAM',
  'SITE_PESSOAL',
  'GOOGLE_SCHOLAR',
  'RESEARCHGATE',
  'OUTRO'
]);

var CORE_PESSOAS_V2_LINKS_PERFIS_LABELS = Object.freeze({
  LATTES: 'Curriculo Lattes',
  LINKEDIN: 'LinkedIn',
  ORCID: 'ORCID',
  INSTAGRAM: 'Instagram',
  SITE_PESSOAL: 'Site pessoal',
  GOOGLE_SCHOLAR: 'Google Scholar',
  RESEARCHGATE: 'ResearchGate',
  OUTRO: 'Link externo'
});

/** Normaliza tipos de links para o catalogo canonico de Pessoas V2. */
function core_domainsV2NormalizeLinkPerfilType_(value) {
  var normalized = core_domainsV2AuditStatus_(value).replace(/[\s-]+/g, '_');
  if (normalized === 'CURRICULO_LATTES' || normalized === 'CURRICULO') normalized = 'LATTES';
  if (normalized === 'GOOGLE_SCHOLAR_PROFILE') normalized = 'GOOGLE_SCHOLAR';
  return CORE_PESSOAS_V2_LINKS_PERFIS_TYPES.indexOf(normalized) >= 0 ? normalized : 'OUTRO';
}

/** Retorna somente URLs HTTP(S) seguras para contratos consumidos pelo Portal. */
function core_domainsV2NormalizeProfileUrl_(value) {
  var raw = String(value || '').trim();
  if (!raw || /[\s<>]/.test(raw)) return '';
  if (/^https?:\/\//i.test(raw)) return raw;
  if (/^(www\.)?[a-z0-9.-]+\.[a-z]{2,}(?:\/[^\s<>]*)?$/i.test(raw)) return 'https://' + raw;
  return '';
}

/** Indica se o registro de link esta ativo para leitura. */
function core_domainsV2IsLinkPerfilActive_(record) {
  return core_domainsV2AuditIsSim_((record || {}).ATIVO);
}

/** Converte um registro bruto de link no formato seguro para consumidores. */
function core_domainsV2MapLinkPerfil_(record) {
  var source = record || {};
  var url = core_domainsV2NormalizeProfileUrl_(source.URL);
  if (!url) return null;
  var tipo = core_domainsV2NormalizeLinkPerfilType_(source.TIPO_LINK);
  return {
    idLink: String(source.ID_LINK || '').trim(),
    tipo: tipo,
    url: url,
    rotulo: String(source.ROTULO || CORE_PESSOAS_V2_LINKS_PERFIS_LABELS[tipo] || 'Link externo').trim(),
    publicavel: core_domainsV2AuditIsSim_(source.PUBLICAVEL),
    visivelPortal: core_domainsV2AuditIsSim_(source.VISIVEL_PORTAL)
  };
}

/** Lista links de uma pessoa, aplicando o filtro publico quando solicitado. */
function core_domainsV2LinksPerfisByPessoa_(pessoasData, idPessoa, options) {
  var opts = options || {};
  var id = String(idPessoa || '').trim();
  if (!id) return [];
  var records = (pessoasData.PESSOAS_LINKS_PERFIS && pessoasData.PESSOAS_LINKS_PERFIS.records) || [];
  return records.filter(function(record) {
    if (String(record.ID_PESSOA || '').trim() !== id) return false;
    if (opts.includeInactive !== true && !core_domainsV2IsLinkPerfilActive_(record)) return false;
    if (opts.publicOnly === true && !(core_domainsV2AuditIsSim_(record.PUBLICAVEL) && core_domainsV2AuditIsSim_(record.VISIVEL_PORTAL))) return false;
    return true;
  }).map(core_domainsV2MapLinkPerfil_).filter(function(link) {
    return !!link;
  }).sort(function(a, b) {
    var aPriority = a.tipo === 'LATTES' ? 0 : 1;
    var bPriority = b.tipo === 'LATTES' ? 0 : 1;
    if (aPriority !== bPriority) return aPriority - bPriority;
    return String(a.rotulo || a.tipo).localeCompare(String(b.rotulo || b.tipo));
  });
}

function core_domainsV2Active_(record) {
  return core_domainsV2AuditStatus_(record.STATUS_VINCULO || record.STATUS || record.STATUS_VIGENCIA) === 'ATIVO' ||
    core_domainsV2AuditStatus_(record.STATUS_VINCULO || record.STATUS || record.STATUS_VIGENCIA) === 'ATIVA' ||
    core_domainsV2AuditIsSim_(record.ATIVO);
}

function core_domainsV2CurrentVigencia_(record) {
  return core_domainsV2ResumoVigenciaAtual_(record);
}

function core_domainsV2IndexFirstBy_(records, field) {
  var out = {};
  (records || []).forEach(function(record) {
    var key = String(record[field] || '').trim();
    if (key && !out[key]) out[key] = record;
  });
  return out;
}

function core_domainsV2IndexManyBy_(records, field) {
  return core_domainsV2AuditIndexBy_(records || [], field);
}

function core_domainsV2FindPessoaById_(pessoasData, idPessoa) {
  var id = String(idPessoa || '').trim();
  if (!id) return null;
  var records = (pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [];
  for (var i = 0; i < records.length; i++) {
    if (String(records[i].ID_PESSOA || '').trim() === id) return records[i];
  }
  return null;
}

function core_domainsV2FindPessoaIdByIdentifier_(pessoasData, type, value) {
  var normalizedType = core_domainsV2AuditStatus_(type);
  var normalizedValue = normalizedType === 'EMAIL'
    ? core_domainsV2Email_(value)
    : core_domainsV2Rga_(value);
  if (!normalizedValue) return '';

  var identifiers = (pessoasData.PESSOAS_IDENTIFICADORES && pessoasData.PESSOAS_IDENTIFICADORES.records) || [];
  for (var i = identifiers.length - 1; i >= 0; i--) {
    var record = identifiers[i];
    var currentType = core_domainsV2AuditStatus_(record.TIPO_IDENTIFICADOR);
    var currentValue = currentType === 'EMAIL'
      ? core_domainsV2Email_(record.VALOR_IDENTIFICADOR)
      : core_domainsV2Rga_(record.VALOR_IDENTIFICADOR);
    if (currentType === normalizedType && currentValue === normalizedValue && core_domainsV2AuditIsSim_(record.ATIVO)) {
      return String(record.ID_PESSOA || '').trim();
    }
  }
  return '';
}

function core_domainsV2FindPessoaIdByEmail_(pessoasData, email) {
  var normalized = core_domainsV2Email_(email);
  if (!normalized) return '';
  var base = (pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [];
  for (var i = 0; i < base.length; i++) {
    if (core_domainsV2Email_(base[i].EMAIL_PRINCIPAL) === normalized) return String(base[i].ID_PESSOA || '').trim();
  }
  return core_domainsV2FindPessoaIdByIdentifier_(pessoasData, 'EMAIL', normalized);
}

function core_domainsV2FindPessoaIdByRga_(pessoasData, rga) {
  var normalized = core_domainsV2Rga_(rga);
  if (!normalized) return '';
  var detalhes = (pessoasData.MEMBROS_DETALHES && pessoasData.MEMBROS_DETALHES.records) || [];
  for (var i = detalhes.length - 1; i >= 0; i--) {
    if (core_domainsV2Rga_(detalhes[i].RGA) === normalized) return String(detalhes[i].ID_PESSOA || '').trim();
  }
  return core_domainsV2FindPessoaIdByIdentifier_(pessoasData, 'RGA', normalized);
}

function core_domainsV2PessoaBundle_(pessoasData, idPessoa) {
  var pessoa = core_domainsV2FindPessoaById_(pessoasData, idPessoa);
  if (!pessoa) return null;

  var id = String(idPessoa || '').trim();
  var byPessoa = function(sheetName) {
    return ((pessoasData[sheetName] && pessoasData[sheetName].records) || []).filter(function(record) {
      return String(record.ID_PESSOA || '').trim() === id;
    }).map(core_domainsV2CloneRecord_);
  };

  return {
    pessoa: core_domainsV2CloneRecord_(pessoa),
    identificadores: byPessoa('PESSOAS_IDENTIFICADORES'),
    membrosDetalhes: byPessoa('MEMBROS_DETALHES')[0] || null,
    linksPerfis: core_domainsV2LinksPerfisByPessoa_(pessoasData, id, { includeInactive: false }),
    colaboradoresAcademicos: byPessoa('COLABORADORES_ACADEMICOS'),
    participantesExternosDetalhes: byPessoa('PARTICIPANTES_EXTERNOS_DETALHES'),
    vinculos: byPessoa('VINCULOS_GEAPA'),
    resumoOperacional: byPessoa('PESSOAS_RESUMO_OPERACIONAL')[0] || null,
    comunicacaoConsentimentos: byPessoa('PESSOAS_COMUNICACAO_CONSENTIMENTOS'),
    portalExcecoes: byPessoa('PORTAL_ACESSOS_EXCECOES')
  };
}

function corePessoasGetById_(idPessoa, options) {
  var report = core_domainsV2NewReadReport_('PESSOAS_GET_BY_ID');
  var pessoasData = core_domainsV2OpenPessoasSubset_([
    'BASE',
    'IDENTIFICADORES',
    'MEMBROS_DETALHES',
    'LINKS_PERFIS',
    'COLABORADORES',
    'EXTERNOS',
    'VINCULOS',
    'RESUMO',
    'CONSENTIMENTOS',
    'ACESSOS_EXCECOES'
  ], report, options || {});
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));
  return core_domainsV2PessoaBundle_(pessoasData, idPessoa);
}

function corePessoasFindByEmail_(email, options) {
  var report = core_domainsV2NewReadReport_('PESSOAS_FIND_BY_EMAIL');
  var pessoasData = core_domainsV2OpenPessoasSubset_([
    'BASE',
    'IDENTIFICADORES',
    'MEMBROS_DETALHES',
    'LINKS_PERFIS',
    'COLABORADORES',
    'EXTERNOS',
    'VINCULOS',
    'RESUMO',
    'CONSENTIMENTOS',
    'ACESSOS_EXCECOES'
  ], report, options || {});
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));
  var idPessoa = core_domainsV2FindPessoaIdByEmail_(pessoasData, email);
  return idPessoa ? core_domainsV2PessoaBundle_(pessoasData, idPessoa) : null;
}

function corePessoasFindByRga_(rga, options) {
  var report = core_domainsV2NewReadReport_('PESSOAS_FIND_BY_RGA');
  var pessoasData = core_domainsV2OpenPessoasSubset_([
    'BASE',
    'IDENTIFICADORES',
    'MEMBROS_DETALHES',
    'LINKS_PERFIS',
    'COLABORADORES',
    'EXTERNOS',
    'VINCULOS',
    'RESUMO',
    'CONSENTIMENTOS',
    'ACESSOS_EXCECOES'
  ], report, options || {});
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));
  var idPessoa = core_domainsV2FindPessoaIdByRga_(pessoasData, rga);
  return idPessoa ? core_domainsV2PessoaBundle_(pessoasData, idPessoa) : null;
}

function corePessoasGetOperationalSummary_(idPessoa, options) {
  var report = core_domainsV2NewReadReport_('PESSOAS_GET_OPERATIONAL_SUMMARY');
  var pessoasData = core_domainsV2OpenPessoasSubset_(['RESUMO'], report, options || {});
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));
  var id = String(idPessoa || '').trim();
  var resumo = (pessoasData.PESSOAS_RESUMO_OPERACIONAL && pessoasData.PESSOAS_RESUMO_OPERACIONAL.records) || [];
  for (var i = 0; i < resumo.length; i++) {
    if (String(resumo[i].ID_PESSOA || '').trim() === id) return core_domainsV2CloneRecord_(resumo[i]);
  }
  return null;
}

/** Normaliza texto para busca e comparacao dos filtros administrativos. */
function core_pessoasAdminPortalNormalize_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
}

/** Converte valores de planilha/filtro em booleano quando a intencao estiver explicita. */
function core_pessoasAdminPortalBoolean_(value) {
  if (value === true || value === false) return value;
  var normalized = core_pessoasAdminPortalNormalize_(value);
  if (['SIM', 'TRUE', '1', 'ATIVO'].indexOf(normalized) >= 0) return true;
  if (['NAO', 'FALSE', '0', 'INATIVO'].indexOf(normalized) >= 0) return false;
  return null;
}

/** Converte indicador numerico e devolve nulo quando a origem estiver vazia ou invalida. */
function core_pessoasAdminPortalNumber_(value) {
  if (value === '' || value == null) return null;
  var number = Number(value);
  return isFinite(number) ? number : null;
}

/** Indica se a linha representa um vinculo de membro administravel nesta primeira versao. */
function core_pessoasAdminPortalIsMember_(record) {
  var tipo = core_pessoasAdminPortalNormalize_((record || {}).TIPO_VINCULO_ATUAL);
  return tipo === 'MEMBRO' || tipo.indexOf('MEMBRO_') === 0;
}

/** Indica se o resumo registra alguma pendencia operacional aberta. */
function core_pessoasAdminPortalHasPending_(value) {
  var normalized = core_pessoasAdminPortalNormalize_(value);
  return !!normalized && ['SEM_PENDENCIAS', 'SEM PENDENCIAS', 'NAO', 'NENHUMA'].indexOf(normalized) < 0;
}

/** Mapeia uma linha do cache operacional para o contrato estritamente sanitizado do Portal. */
function core_pessoasAdminPortalMapRow_(record) {
  var source = record || {};
  return Object.freeze({
    idPessoa: String(source.ID_PESSOA || '').trim(),
    rga: String(source.RGA || '').trim(),
    nomeExibicao: String(source.NOME_EXIBICAO || '').trim(),
    email: core_domainsV2Email_(source.EMAIL),
    tipoVinculoAtual: String(source.TIPO_VINCULO_ATUAL || '').trim(),
    statusVinculoAtual: String(source.STATUS_VINCULO_ATUAL || '').trim(),
    cargoFuncaoAtual: String(source.CARGO_FUNCAO_ATUAL || '').trim(),
    perfilPortalCalculado: String(source.PERFIL_PORTAL_CALCULADO || '').trim(),
    portalAtivo: core_pessoasAdminPortalBoolean_(source.PORTAL_ATIVO),
    tempoEfetivoNoGrupo: String(source.TEMPO_EFETIVO_NO_GRUPO || '').trim(),
    qtdSemestresNoGrupo: core_pessoasAdminPortalNumber_(source.QTD_SEMESTRES_NO_GRUPO),
    qtdApresentacoesRealizadas: core_pessoasAdminPortalNumber_(source.QTD_APRESENTACOES_REALIZADAS),
    cicloUltimaApresentacao: String(source.CICLO_ULTIMA_APRESENTACAO || '').trim(),
    frequenciaResumida: String(source.FREQUENCIA_RESUMIDA || '').trim(),
    pendenciasAbertas: String(source.PENDENCIAS_ABERTAS || '').trim(),
    flagJaFoiSuspenso: String(source.FLAG_JA_FOI_SUSPENSO || '').trim(),
    statusElegibilidadeDiretoria: String(source.STATUS_ELEGIBILIDADE_DIRETORIA || '').trim(),
    ultimaAtualizacao: source.ULTIMA_ATUALIZACAO || ''
  });
}

/** Normaliza e limita os filtros aceitos pelo contrato administrativo. */
function core_pessoasAdminPortalFilters_(filters) {
  var source = filters && typeof filters === 'object' ? filters : {};
  var pageSize = Math.min(Math.max(Number(source.pageSize || source.limite || 50) || 50, 1), 100);
  return Object.freeze({
    texto: core_pessoasAdminPortalNormalize_(source.texto || source.busca),
    tipoVinculo: core_pessoasAdminPortalNormalize_(source.tipoVinculo || source.tipoVinculoAtual),
    statusVinculo: core_pessoasAdminPortalNormalize_(source.statusVinculo || source.statusVinculoAtual),
    perfilPortal: core_pessoasAdminPortalNormalize_(source.perfilPortal || source.perfilPortalCalculado),
    portalAtivo: core_pessoasAdminPortalBoolean_(source.portalAtivo),
    comPendencias: core_pessoasAdminPortalBoolean_(source.comPendencias),
    situacaoFrequencia: core_pessoasAdminPortalNormalize_(source.situacaoFrequencia || source.frequencia),
    pagina: Math.max(Number(source.pagina || source.page || 1) || 1, 1),
    pageSize: pageSize
  });
}

/** Aplica os filtros homologados sem recalcular qualquer indicador operacional. */
function core_pessoasAdminPortalMatches_(item, filters) {
  if (filters.texto) {
    var haystack = core_pessoasAdminPortalNormalize_([item.nomeExibicao, item.rga, item.email].join(' '));
    if (haystack.indexOf(filters.texto) < 0) return false;
  }
  if (filters.tipoVinculo && core_pessoasAdminPortalNormalize_(item.tipoVinculoAtual) !== filters.tipoVinculo) return false;
  if (filters.statusVinculo && core_pessoasAdminPortalNormalize_(item.statusVinculoAtual) !== filters.statusVinculo) return false;
  if (filters.perfilPortal && core_pessoasAdminPortalNormalize_(item.perfilPortalCalculado) !== filters.perfilPortal) return false;
  if (filters.portalAtivo !== null && item.portalAtivo !== filters.portalAtivo) return false;
  if (filters.comPendencias !== null && core_pessoasAdminPortalHasPending_(item.pendenciasAbertas) !== filters.comPendencias) return false;
  if (filters.situacaoFrequencia) {
    var frequency = core_pessoasAdminPortalNormalize_(item.frequenciaResumida);
    if (filters.situacaoFrequencia === 'SEM_DADOS' && frequency) return false;
    if (filters.situacaoFrequencia !== 'SEM_DADOS' && frequency.indexOf(filters.situacaoFrequencia) < 0) return false;
  }
  return true;
}

/** Extrai opcoes de filtro nao sensiveis a partir da colecao completa de membros. */
function core_pessoasAdminPortalFilterOptions_(items) {
  function distinct(field) {
    var seen = {};
    return (items || []).map(function(item) { return String(item[field] || '').trim(); }).filter(function(value) {
      var key = core_pessoasAdminPortalNormalize_(value);
      if (!value || seen[key]) return false;
      seen[key] = true;
      return true;
    }).sort(function(a, b) { return a.localeCompare(b, 'pt-BR'); });
  }
  return Object.freeze({
    tiposVinculo: Object.freeze(distinct('tipoVinculoAtual')),
    statusVinculo: Object.freeze(distinct('statusVinculoAtual')),
    perfisPortal: Object.freeze(distinct('perfilPortalCalculado')),
    frequencias: Object.freeze(distinct('frequenciaResumida'))
  });
}

/** Valida no Core se a sessao oficial possui acesso ativo e a permissao membros:ler. */
function core_pessoasAdminPortalAuthorize_(session) {
  var permissions = session && Array.isArray(session.permissoes) ? session.permissoes : [];
  return !!(
    session &&
    session.ok !== false &&
    session.autenticado === true &&
    session.portalAtivo === true &&
    permissions.map(corePortalNormalizePermission_).indexOf('membros:ler') >= 0
  );
}

/** Registra diagnostico tecnico da listagem sem nomes, e-mails, RGA ou filtros textuais. */
function core_pessoasAdminPortalLog_(payload) {
  var source = payload || {};
  Logger.log('[geapa-core][portal][admin-members] ' + JSON.stringify({
    ok: source.ok === true,
    code: String(source.code || '').slice(0, 80),
    totalBase: Number(source.totalBase || 0),
    totalFiltrado: Number(source.totalFiltrado || 0),
    pagina: Number(source.pagina || 0),
    pageSize: Number(source.pageSize || 0),
    duracaoMs: Number(source.duracaoMs || 0)
  }));
}

/**
 * Lista membros para a area administrativa do Portal usando apenas o cache operacional V2.
 * A identidade do solicitante e resolvida novamente pelo Core e nenhum dado sensivel e retornado.
 */
function core_listarMembrosAdministracaoPortal_(filtros, contexto) {
  var startedAt = Date.now();
  var ctx = contexto && typeof contexto === 'object' ? contexto : {};
  var environment = core_normalizeDomainEnv_(ctx);
  var requester = ctx.idPessoa || ctx.identificadorSolicitante || ctx.identificador || '';
  var suppliedSession = ctx.sessao && typeof ctx.sessao === 'object' ? ctx.sessao : null;
  var session = suppliedSession &&
    String(suppliedSession.idPessoa || '').trim() === String(requester || '').trim()
    ? suppliedSession
    : corePortalResolverUsuarioAtual_(requester, {
        origem: 'adminMembrosListar',
        ambiente: environment,
        traceId: ctx.traceId || ctx.requestId || ''
      });
  if (!core_pessoasAdminPortalAuthorize_(session)) {
    core_pessoasAdminPortalLog_({ ok: false, code: 'ACESSO_NEGADO', duracaoMs: Date.now() - startedAt });
    return Object.freeze({
      ok: false,
      errorCode: 'ACESSO_NEGADO',
      message: 'Usuario sem permissao para consultar membros.'
    });
  }

  var normalizedFilters = core_pessoasAdminPortalFilters_(filtros);
  // Usa o resolvedor oficial do dominio para respeitar o ambiente efetivo de Pessoas V2.
  var report = core_domainsV2NewReadReport_('PORTAL_ADMIN_MEMBROS_LISTAR');
  var pessoasData = core_domainsV2OpenPessoasSubset_(['RESUMO'], report, { ambiente: environment });
  if (report.totalErros) {
    throw new Error('PESSOAS_V2_RESUMO_OPERACIONAL_INDISPONIVEL');
  }
  var records = (pessoasData.PESSOAS_RESUMO_OPERACIONAL && pessoasData.PESSOAS_RESUMO_OPERACIONAL.records) || [];
  var members = records.filter(core_pessoasAdminPortalIsMember_).map(core_pessoasAdminPortalMapRow_);
  members.sort(function(a, b) {
    return String(a.nomeExibicao || a.idPessoa).localeCompare(String(b.nomeExibicao || b.idPessoa), 'pt-BR');
  });
  var filtered = members.filter(function(item) { return core_pessoasAdminPortalMatches_(item, normalizedFilters); });
  var totalPages = Math.max(Math.ceil(filtered.length / normalizedFilters.pageSize), 1);
  var page = Math.min(normalizedFilters.pagina, totalPages);
  var start = (page - 1) * normalizedFilters.pageSize;
  var pageItems = filtered.slice(start, start + normalizedFilters.pageSize);
  var lastUpdate = members.reduce(function(latest, item) {
    var current = item.ultimaAtualizacao ? new Date(item.ultimaAtualizacao).getTime() : 0;
    return current > latest ? current : latest;
  }, 0);
  var data = Object.freeze({
    itens: Object.freeze(pageItems),
    paginacao: Object.freeze({
      pagina: page,
      pageSize: normalizedFilters.pageSize,
      totalItens: filtered.length,
      totalPaginas: totalPages,
      temAnterior: page > 1,
      temProxima: page < totalPages
    }),
    totalMembrosBase: members.length,
    opcoesFiltros: core_pessoasAdminPortalFilterOptions_(members),
    ultimaAtualizacaoResumo: lastUpdate ? new Date(lastUpdate) : '',
    somenteLeitura: true,
    fonte: 'PESSOAS_V2_RESUMO_OPERACIONAL'
  });
  core_pessoasAdminPortalLog_({
    ok: true,
    code: 'MEMBROS_ADMIN_LISTADOS',
    totalBase: members.length,
    totalFiltrado: filtered.length,
    pagina: page,
    pageSize: normalizedFilters.pageSize,
    ambiente: environment,
    duracaoMs: Date.now() - startedAt
  });
  return Object.freeze({ ok: true, data: data });
}

function core_domainsV2ListByActiveLinkType_(tipoVinculo, opts) {
  opts = opts || {};
  var report = core_domainsV2NewReadReport_('PESSOAS_LIST_BY_LINK');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));

  var tipo = core_domainsV2AuditTipoVinculo_(tipoVinculo);
  var baseById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [], 'ID_PESSOA');
  var detalhesById = core_domainsV2IndexFirstBy_((pessoasData.MEMBROS_DETALHES && pessoasData.MEMBROS_DETALHES.records) || [], 'ID_PESSOA');
  var resumoById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_RESUMO_OPERACIONAL && pessoasData.PESSOAS_RESUMO_OPERACIONAL.records) || [], 'ID_PESSOA');
  var out = [];
  var seen = {};

  ((pessoasData.VINCULOS_GEAPA && pessoasData.VINCULOS_GEAPA.records) || []).forEach(function(vinculo) {
    var idPessoa = String(vinculo.ID_PESSOA || '').trim();
    if (!idPessoa || seen[idPessoa]) return;
    if (core_domainsV2AuditTipoVinculo_(vinculo.TIPO_VINCULO) !== tipo) return;
    if (opts.includeInactive !== true && !core_domainsV2Active_(vinculo)) return;
    var pessoa = baseById[idPessoa] || {};
    var detalhes = detalhesById[idPessoa] || {};
    var resumo = resumoById[idPessoa] || {};
    out.push({
      idPessoa: idPessoa,
      nome: pessoa.NOME_EXIBICAO || pessoa.NOME_COMPLETO || resumo.NOME_EXIBICAO || '',
      email: pessoa.EMAIL_PRINCIPAL || resumo.EMAIL || '',
      rga: detalhes.RGA || resumo.RGA || '',
      tipoVinculo: vinculo.TIPO_VINCULO || '',
      statusVinculo: vinculo.STATUS_VINCULO || '',
      vinculo: core_domainsV2CloneRecord_(vinculo),
      resumoOperacional: core_domainsV2CloneRecord_(resumo)
    });
    seen[idPessoa] = true;
  });

  return out.sort(function(a, b) {
    return String(a.nome || a.idPessoa).localeCompare(String(b.nome || b.idPessoa));
  });
}

function corePessoasListCurrentMembers_(opts) {
  return core_domainsV2ListByActiveLinkType_('MEMBRO_EFETIVO', opts || {});
}

function corePessoasListExMembers_(opts) {
  return core_domainsV2ListByActiveLinkType_('EGRESSO', opts || {});
}

function corePessoasListWaitingMembers_(opts) {
  return core_domainsV2ListByActiveLinkType_('MEMBRO_EM_ESPERA', opts || {});
}

function core_domainsV2ListDetailWithPessoa_(detailSheetName, opts) {
  opts = opts || {};
  var report = core_domainsV2NewReadReport_('PESSOAS_LIST_DETAIL');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));
  var baseById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [], 'ID_PESSOA');
  return ((pessoasData[detailSheetName] && pessoasData[detailSheetName].records) || []).filter(function(record) {
    return opts.includeInactive === true || core_domainsV2AuditIsSim_(record.ATIVO);
  }).map(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    var pessoa = baseById[idPessoa] || {};
    return {
      idPessoa: idPessoa,
      nome: pessoa.NOME_EXIBICAO || pessoa.NOME_COMPLETO || '',
      email: pessoa.EMAIL_PRINCIPAL || '',
      pessoa: core_domainsV2CloneRecord_(pessoa),
      detalhe: core_domainsV2CloneRecord_(record)
    };
  });
}

function corePessoasListAcademicCollaborators_(opts) {
  return core_domainsV2ListDetailWithPessoa_('COLABORADORES_ACADEMICOS', opts || {});
}

function corePessoasListExternalParticipants_(opts) {
  return core_domainsV2ListDetailWithPessoa_('PARTICIPANTES_EXTERNOS_DETALHES', opts || {});
}

function coreVigenciasListCurrentFunctions_(opts) {
  opts = opts || {};
  var report = core_domainsV2NewReadReport_('VIGENCIAS_LIST_CURRENT_FUNCTIONS');
  var vigenciasData = core_domainsV2OpenVigencias_(report);
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) throw new Error('Vigencias v2 indisponivel: ' + JSON.stringify(report.erros));
  var pessoaById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [], 'ID_PESSOA');
  var cargoByKey = core_domainsV2IndexFirstBy_((vigenciasData.CARGOS_CONFIG && vigenciasData.CARGOS_CONFIG.records) || [], 'CARGO_KEY');
  var idFilter = String(opts.idPessoa || '').trim();
  return ((vigenciasData.VIGENCIAS_FUNCOES && vigenciasData.VIGENCIAS_FUNCOES.records) || []).filter(function(record) {
    if (!core_domainsV2CurrentVigencia_(record)) return false;
    if (idFilter && String(record.ID_PESSOA || '').trim() !== idFilter) return false;
    return true;
  }).map(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    var cargo = cargoByKey[String(record.CARGO_KEY || '').trim()] || {};
    var pessoa = pessoaById[idPessoa] || {};
    return {
      idPessoa: idPessoa,
      nome: pessoa.NOME_EXIBICAO || pessoa.NOME_COMPLETO || '',
      vigencia: core_domainsV2CloneRecord_(record),
      cargoConfig: core_domainsV2CloneRecord_(cargo)
    };
  });
}

function coreVigenciasGetCurrentFunctionByPessoa_(idPessoa) {
  var list = coreVigenciasListCurrentFunctions_({ idPessoa: idPessoa });
  list.sort(function(a, b) {
    return core_domainsV2ResumoTimestamp_(b.vigencia.DATA_INICIO) - core_domainsV2ResumoTimestamp_(a.vigencia.DATA_INICIO);
  });
  return list[0] || null;
}

function coreVigenciasGetCurrentSummaryByPessoa_(idPessoa, options) {
  var report = core_domainsV2NewReadReport_('VIGENCIAS_GET_CURRENT_SUMMARY_BY_PESSOA');
  var vigenciasData = core_domainsV2OpenSubset_('VIGENCIAS', ['RESUMO'], report, options || {});
  if (report.totalErros) throw new Error('Vigencias v2 indisponivel: ' + JSON.stringify(report.erros));
  var id = String(idPessoa || '').trim();
  var records = (vigenciasData.VIGENCIAS_RESUMO_ATUAL && vigenciasData.VIGENCIAS_RESUMO_ATUAL.records) || [];
  for (var i = 0; i < records.length; i++) {
    if (String(records[i].ID_PESSOA || '').trim() === id) return core_domainsV2CloneRecord_(records[i]);
  }
  return null;
}

function coreVigenciasGetPortalPermissionsByPessoa_(idPessoa) {
  var functions = coreVigenciasListCurrentFunctions_({ idPessoa: idPessoa });
  var perfis = [];
  var permissoes = [];
  functions.forEach(function(item) {
    var cargo = item.cargoConfig || {};
    if (core_domainsV2AuditIsSim_(cargo.GERA_PERFIL_PORTAL)) {
      core_domainsV2ResumoPushDistinct_(perfis, cargo.PERFIL_PORTAL_PADRAO || cargo.CARGO_KEY);
    }
    core_domainsV2ResumoPermissionsFromCargo_(cargo).forEach(function(permission) {
      core_domainsV2ResumoPushDistinct_(permissoes, permission);
    });
  });
  return {
    idPessoa: String(idPessoa || '').trim(),
    perfisPortalCalculados: perfis,
    permissoesCalculadas: permissoes,
    funcoesConsideradas: functions.length
  };
}

function core_domainsV2PickActiveVinculo_(vinculos) {
  return core_domainsV2PickCurrentVinculo_(vinculos);
}

function core_domainsV2NormalizeTipoVinculo_(value) {
  var tipo = core_domainsV2AuditTipoVinculo_(value);
  if (tipo === 'EX_MEMBRO' || tipo === 'EX-MEMBRO') return 'EGRESSO';
  return tipo;
}

function core_domainsV2PickCurrentVinculo_(vinculos) {
  var priority = {
    MEMBRO_EFETIVO_ATIVO: 1,
    MEMBRO_EM_ESPERA_ATIVO: 2,
    EGRESSO: 3,
    OUTRO_ATIVO: 4,
    MEMBRO_EFETIVO_ENCERRADO: 5,
    MEMBRO_EM_ESPERA_ENCERRADO: 6,
    OUTRO: 9
  };
  return (vinculos || []).slice().sort(function(a, b) {
    var pa = core_domainsV2CurrentVinculoRank_(a, priority);
    var pb = core_domainsV2CurrentVinculoRank_(b, priority);
    if (pa !== pb) return pa - pb;
    return core_domainsV2ResumoTimestamp_(b.DATA_INICIO) - core_domainsV2ResumoTimestamp_(a.DATA_INICIO);
  })[0] || null;
}

function core_domainsV2CurrentVinculoRank_(vinculo, priority) {
  var tipo = core_domainsV2NormalizeTipoVinculo_(vinculo.TIPO_VINCULO);
  var active = core_domainsV2Active_(vinculo);
  if (tipo === 'MEMBRO_EFETIVO' && active) return priority.MEMBRO_EFETIVO_ATIVO;
  if (tipo === 'MEMBRO_EM_ESPERA' && active) return priority.MEMBRO_EM_ESPERA_ATIVO;
  if (tipo === 'EGRESSO') return priority.EGRESSO;
  if (active) return priority.OUTRO_ATIVO;
  if (tipo === 'MEMBRO_EFETIVO') return priority.MEMBRO_EFETIVO_ENCERRADO;
  if (tipo === 'MEMBRO_EM_ESPERA') return priority.MEMBRO_EM_ESPERA_ENCERRADO;
  return priority.OUTRO;
}

function core_domainsV2Date_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  var text = String(value || '').trim();
  var m = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    var year = Number(m[3]);
    if (year < 100) year += 2000;
    return new Date(
      year,
      Number(m[2]) - 1,
      Number(m[1]),
      Number(m[4] || 0),
      Number(m[5] || 0),
      Number(m[6] || 0)
    );
  }
  var parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function core_domainsV2DaysBetween_(start, end) {
  if (!start || !end || end < start) return 0;
  var a = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  var b = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  return Math.floor((b.getTime() - a.getTime()) / 86400000) + 1;
}

function core_domainsV2IntervalsOverlap_(aStart, aEnd, bStart, bEnd) {
  if (!aStart || !bStart) return false;
  var aLast = aEnd || aStart;
  var bLast = bEnd || bStart;
  return aStart <= bLast && bStart <= aLast;
}

function core_domainsV2FormatDuration_(days) {
  if (!days) return '';
  var years = Math.floor(days / 365.25);
  var months = Math.floor((days - (years * 365.25)) / 30.44);
  if (years && months) return years + ' anos, ' + months + ' meses';
  if (years) return years + ' anos';
  if (months) return months + ' meses';
  return days + ' dias';
}

/**
 * Mescla intervalos sobrepostos ou contiguos para evitar dupla contagem.
 *
 * @param {Array<Object>} intervals
 * @return {Array<Object>}
 */
function core_domainsV2MergeIntervals_(intervals) {
  var sorted = (intervals || []).slice().sort(function(a, b) {
    return a.start.getTime() - b.start.getTime();
  });
  var merged = [];
  sorted.forEach(function(interval) {
    if (!merged.length) {
      merged.push({ start: interval.start, end: interval.end });
      return;
    }
    var current = merged[merged.length - 1];
    var nextDay = new Date(current.end.getTime());
    nextDay.setDate(nextDay.getDate() + 1);
    if (interval.start.getTime() <= nextDay.getTime()) {
      if (interval.end.getTime() > current.end.getTime()) current.end = interval.end;
      return;
    }
    merged.push({ start: interval.start, end: interval.end });
  });
  return merged;
}

function core_domainsV2EffectiveMemberIntervals_(vinculos, today) {
  var intervals = (vinculos || []).filter(function(vinculo) {
    return core_domainsV2NormalizeTipoVinculo_(vinculo.TIPO_VINCULO) === 'MEMBRO_EFETIVO';
  }).map(function(vinculo) {
    var start = core_domainsV2Date_(vinculo.DATA_INICIO);
    if (!start || start > today) return null;
    var end = core_domainsV2Date_(vinculo.DATA_FIM);
    if (!end) end = core_domainsV2Active_(vinculo) ? today : start;
    if (end > today) end = today;
    if (end < start) return null;
    return { start: start, end: end };
  }).filter(function(interval) {
    return !!interval;
  });
  return core_domainsV2MergeIntervals_(intervals);
}

function core_domainsV2CountIntervalDays_(intervals) {
  return (intervals || []).reduce(function(total, interval) {
    return total + core_domainsV2DaysBetween_(interval.start, interval.end);
  }, 0);
}

function core_domainsV2CountSemestersForIntervals_(vigenciasData, intervals) {
  if (!intervals.length) return '';
  var records = ((vigenciasData.SEMESTRES && vigenciasData.SEMESTRES.records) || []);
  // QTD_SEMESTRES_NO_GRUPO deve seguir exclusivamente os semestres letivos oficiais de Vigencias v2.
  if (!records.length) return '';
  var seen = {};
  records.forEach(function(record) {
    var start = core_domainsV2Date_(record.DATA_INICIO);
    var end = core_domainsV2Date_(record.DATA_FIM);
    if (!start || !end) return;
    var crosses = intervals.some(function(interval) {
      return core_domainsV2IntervalsOverlap_(interval.start, interval.end, start, end);
    });
    if (crosses) {
      var key = record.ID_SEMESTRE || record.ID_PERIODO || record.NOME_PERIODO || [record.ANO, record.SEMESTRE].join('/');
      seen[String(key || record.__rowNumber || '').trim()] = true;
    }
  });
  return Object.keys(seen).length || '';
}

function core_domainsV2GetRga_(pessoasData, idPessoa, detalhes) {
  var fallback = detalhes && detalhes.RGA ? detalhes.RGA : '';
  var records = (pessoasData.PESSOAS_IDENTIFICADORES && pessoasData.PESSOAS_IDENTIFICADORES.records) || [];
  for (var i = records.length - 1; i >= 0; i--) {
    var record = records[i];
    if (String(record.ID_PESSOA || '').trim() !== idPessoa) continue;
    if (core_domainsV2AuditStatus_(record.TIPO_IDENTIFICADOR) !== 'RGA') continue;
    if (!core_domainsV2AuditIsSim_(record.ATIVO)) continue;
    if (core_domainsV2AuditIsSim_(record.PRINCIPAL)) return record.VALOR_IDENTIFICADOR || fallback;
    fallback = fallback || record.VALOR_IDENTIFICADOR;
  }
  return fallback || '';
}

function core_domainsV2CurrentSemesterFromVigencias_(vigenciasData, refDate) {
  var now = refDate || new Date();
  var semestres = (vigenciasData.SEMESTRES && vigenciasData.SEMESTRES.records) || [];
  var next = null;
  for (var i = 0; i < semestres.length; i++) {
    var record = semestres[i];
    var start = core_domainsV2Date_(record.DATA_INICIO);
    var end = core_domainsV2Date_(record.DATA_FIM);
    var id = String(record.ID_SEMESTRE || '').trim();
    if (!id || !start) continue;
    if (core_domainsV2IntervalsOverlap_(now, now, start, end || start)) {
      return { id: id, startDate: start, endDate: end };
    }
    if (start > now && (!next || start < next.startDate)) {
      next = { id: id, startDate: start, endDate: end };
    }
  }
  return next;
}

function core_domainsV2StudentCurrentSemesterFromRga_(rga, vigenciasData, refDate) {
  var entry = core_parseEntrySemesterFromRga_(rga);
  if (!entry) return null;
  var currentSemester = core_domainsV2CurrentSemesterFromVigencias_(vigenciasData, refDate);
  if (!currentSemester || !currentSemester.id) return null;
  var current = core_parseSemesterId_(currentSemester.id);
  if (!current) return null;
  var diff = (current.year - entry.year) * 2 + (current.semester - entry.semester);
  var semesterNumber = diff + 1;
  return semesterNumber < 1 ? null : semesterNumber;
}

function core_domainsV2BuildMembrosDetalhesSemesterRows_(pessoasData, vigenciasData, opts, report) {
  var refDate = opts.refDate || new Date();
  var detalhes = (pessoasData.MEMBROS_DETALHES && pessoasData.MEMBROS_DETALHES.records) || [];
  var out = [];
  detalhes.forEach(function(record) {
    var idPessoa = String(record.ID_PESSOA || '').trim();
    if (opts.idPessoa && idPessoa !== opts.idPessoa) return;
    var rga = core_domainsV2GetRga_(pessoasData, idPessoa, record);
    var semester = core_domainsV2StudentCurrentSemesterFromRga_(rga, vigenciasData, refDate);
    if (!rga) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'MEMBRO_DETALHE_SEM_RGA', 'MEMBROS_DETALHES sem RGA para calcular SEMESTRE_ATUAL.', { idPessoa: idPessoa });
    } else if (semester === null) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'RGA_SEM_SEMESTRE_CALCULAVEL', 'RGA nao permitiu calcular SEMESTRE_ATUAL.', { idPessoa: idPessoa, rga: rga });
    }
    out.push({
      rowNumber: record.__rowNumber,
      idPessoa: idPessoa,
      rga: rga,
      semestreAtualAnterior: record.SEMESTRE_ATUAL || '',
      semestreAtualCalculado: semester === null ? '' : semester
    });
  });
  return out;
}

function coreRecalcularMembrosDetalhesSemestreAtualV2_(options) {
  options = options || {};
  var opts = core_domainsV2ResumoOptions_(options);
  opts.idPessoa = options.idPessoa ? String(options.idPessoa).trim() : '';
  opts.refDate = options.refDate ? core_domainsV2Date_(options.refDate) : new Date();
  if (!opts.dryRun && opts.confirmacao !== 'RECALCULAR_MEMBROS_DETALHES_SEMESTRE_ATUAL_V2') {
    throw new Error('Para escrever MEMBROS_DETALHES.SEMESTRE_ATUAL, informe confirmacao: "RECALCULAR_MEMBROS_DETALHES_SEMESTRE_ATUAL_V2".');
  }

  var report = core_domainsV2AuditNewReport_('RECALCULAR_MEMBROS_DETALHES_SEMESTRE_ATUAL_V2');
  report.dryRun = opts.dryRun;
  report.options = {
    idPessoa: opts.idPessoa || null,
    refDate: opts.refDate
  };
  var pessoasData = core_domainsV2OpenPessoas_(report);
  var vigenciasData = core_domainsV2OpenVigencias_(report);
  if (report.totalErros) return report;

  var rows = core_domainsV2BuildMembrosDetalhesSemesterRows_(pessoasData, vigenciasData, opts, report);
  var calculaveis = rows.filter(function(row) {
    return row.semestreAtualCalculado !== '';
  });
  var alteradas = calculaveis.filter(function(row) {
    return String(row.semestreAtualAnterior || '').trim() !== String(row.semestreAtualCalculado);
  });
  report.resumoQuantitativo = {
    linhasAnalisadas: rows.length,
    linhasCalculaveis: calculaveis.length,
    linhasComMudanca: alteradas.length,
    linhasAtualizadas: opts.dryRun ? 0 : alteradas.length
  };
  report.amostra = rows.slice(0, 10);

  if (opts.dryRun) {
    core_domainsV2AuditRecommendation_(report, 'Conferir a amostra; se estiver correta, rodar coreRecalcularMembrosDetalhesSemestreAtualV2({ dryRun: false, confirmacao: "RECALCULAR_MEMBROS_DETALHES_SEMESTRE_ATUAL_V2" }).');
    return report;
  }

  var headers = pessoasData.MEMBROS_DETALHES.headers || [];
  var headerMap = core_buildHeaderIndexMap_(headers);
  var semesterCol = core_findHeaderIndex_(headerMap, 'SEMESTRE_ATUAL');
  if (semesterCol < 0) {
    throw new Error('Cabecalho SEMESTRE_ATUAL nao encontrado em MEMBROS_DETALHES.');
  }
  var sheet = pessoasData.MEMBROS_DETALHES.sheet;
  alteradas.forEach(function(row) {
    if (!row.rowNumber) return;
    sheet.getRange(row.rowNumber, semesterCol + 1).setValue(row.semestreAtualCalculado);
  });
  return report;
}

function core_domainsV2ActivePortalExceptions_(pessoasData, idPessoa, today) {
  return ((pessoasData.PORTAL_ACESSOS_EXCECOES && pessoasData.PORTAL_ACESSOS_EXCECOES.records) || []).filter(function(record) {
    if (String(record.ID_PESSOA || '').trim() !== idPessoa) return false;
    var status = core_domainsV2AuditStatus_(record.STATUS);
    if (['REVOGADO', 'ENCERRADO', 'INATIVO', 'EXPIRADO'].indexOf(status) >= 0) return false;
    var start = core_domainsV2Date_(record.DATA_INICIO);
    var end = core_domainsV2Date_(record.DATA_FIM);
    if (start && start > today) return false;
    if (end && end < today) return false;
    return true;
  });
}

function core_domainsV2BuildPortalState_(vinculo, vigResumo, exceptions) {
  var tipo = core_domainsV2NormalizeTipoVinculo_(vinculo.TIPO_VINCULO);
  var active = core_domainsV2Active_(vinculo);
  var profile = String(vigResumo.PERFIS_PORTAL_CALCULADOS || '').trim();
  var allow = false;
  var block = false;
  (exceptions || []).forEach(function(record) {
    var status = core_domainsV2AuditStatus_(record.STATUS);
    if (['BLOQUEADO', 'NEGADO', 'SUSPENSO'].indexOf(status) >= 0) block = true;
    if (String(record.PERFIL_EXTRA || '').trim()) {
      profile = profile ? profile + '; ' + record.PERFIL_EXTRA : record.PERFIL_EXTRA;
      allow = true;
    }
    if (String(record.PERMISSAO_EXTRA || '').trim()) allow = true;
  });
  if (!profile && tipo === 'MEMBRO_EFETIVO' && active) profile = 'MEMBRO';
  if (!profile && allow) profile = 'EXCECAO_PORTAL';
  if (block) return { perfil: profile, portalAtivo: 'NAO' };
  if (allow && profile) return { perfil: profile, portalAtivo: 'SIM' };
  if (tipo === 'MEMBRO_EFETIVO' && active && profile) return { perfil: profile, portalAtivo: 'SIM' };
  return { perfil: profile, portalAtivo: 'NAO' };
}

function core_domainsV2ReadActivitySheetSoft_(logicalSheet, unavailable, options) {
  try {
    var sheet = core_getDomainSheet_('ATIVIDADES', logicalSheet, options || {});
    return core_readSheetRecords_(sheet, { skipBlankRows: true }) || [];
  } catch (err) {
    unavailable[logicalSheet] = err && err.message ? err.message : String(err);
    return [];
  }
}

/**
 * Le as fontes operacionais de Atividades v2 usadas pelo resumo de Pessoas.
 *
 * @return {Object}
 */
function core_domainsV2ActivityData_(options) {
  var unavailable = {};
  return {
    atividades: core_domainsV2ReadActivitySheetSoft_('ATIVIDADES', unavailable, options),
    apresentacoes: core_domainsV2ReadActivitySheetSoft_('APRESENTACOES', unavailable, options),
    envolvidos: core_domainsV2ReadActivitySheetSoft_('ENVOLVIDOS', unavailable, options),
    portalAtividadesDetalhes: core_domainsV2ReadActivitySheetSoft_('PORTAL_ATIVIDADES_DETALHES', unavailable, options),
    presencas: core_domainsV2ReadActivitySheetSoft_('PRESENCAS_REGISTROS', unavailable, options),
    portalFrequencia: core_domainsV2ReadActivitySheetSoft_('PORTAL_FREQUENCIA_MEMBROS', unavailable, options),
    justificativas: core_domainsV2ReadActivitySheetSoft_('JUSTIFICATIVAS', unavailable, options),
    portalJustificativas: core_domainsV2ReadActivitySheetSoft_('PORTAL_JUSTIFICATIVAS', unavailable, options),
    unavailable: unavailable
  };
}

function core_domainsV2ParseJsonArray_(value) {
  if (Array.isArray(value)) return value;
  var text = String(value || '').trim();
  if (!text) return [];
  try {
    var parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function core_domainsV2PresentationRecordsFromPortalDetails_(details) {
  var records = [];
  (details || []).forEach(function(detail) {
    var presentations = core_domainsV2ParseJsonArray_(core_domainsV2LegacyValue_(detail, [
      'APRESENTACOES_PUBLICAS_JSON',
      'apresentacoesPublicas'
    ]));
    if (!presentations.length) return;

    var base = {
      ID_ATIVIDADE: core_domainsV2LegacyValue_(detail, ['ID_ATIVIDADE', 'idAtividade']),
      DATA_ATIVIDADE: core_domainsV2LegacyValue_(detail, ['DATA_ATIVIDADE', 'dataAtividade', 'DATA']),
      CICLO: core_domainsV2LegacyValue_(detail, ['CICLO', 'ID_CICLO']),
      SEMESTRE: core_domainsV2LegacyValue_(detail, ['SEMESTRE']),
      ROTULO_SEMESTRE: core_domainsV2LegacyValue_(detail, ['ROTULO_SEMESTRE', 'PERIODO', 'ID_PERIODO', 'periodo'])
    };
    presentations.forEach(function(item) {
      if (!item || typeof item !== 'object' || Array.isArray(item)) return;
      var out = {};
      Object.keys(item).forEach(function(key) {
        out[key] = item[key];
      });
      Object.keys(base).forEach(function(key) {
        if (out[key] === undefined || out[key] === null || out[key] === '') out[key] = base[key];
      });
      records.push(out);
    });
  });
  return records;
}

/**
 * Enriquece apresentacoes com identidade e periodo vindos da atividade.
 *
 * A extensao Atividades_Apresentacoes guarda o estado operacional, enquanto
 * a identidade do apresentador permanece em Atividades. A view publica e
 * usada como complemento, inclusive para bases historicas migradas.
 *
 * @param {Object} activityData
 * @return {Array<Object>}
 */
function core_domainsV2EnrichedPresentationRecords_(activityData) {
  var activitiesById = core_domainsV2IndexFirstBy_(activityData.atividades || [], 'ID_ATIVIDADE');
  var portalRecords = core_domainsV2PresentationRecordsFromPortalDetails_(activityData.portalAtividadesDetalhes || []);
  var byId = {};
  var withoutId = [];

  portalRecords.forEach(function(record) {
    var id = String(core_domainsV2LegacyValue_(record, ['ID_APRESENTACAO', 'idApresentacao']) || '').trim();
    if (id) byId[id] = record;
    else withoutId.push(record);
  });

  (activityData.apresentacoes || []).forEach(function(record) {
    var id = String(core_domainsV2LegacyValue_(record, ['ID_APRESENTACAO']) || '').trim();
    var idAtividade = String(core_domainsV2LegacyValue_(record, ['ID_ATIVIDADE']) || '').trim();
    var atividade = activitiesById[idAtividade] || {};
    var enriched = Object.assign({}, id && byId[id] ? byId[id] : {}, record, {
      ID_ATIVIDADE: idAtividade,
      ID_PESSOA: core_domainsV2LegacyValue_(record, ['ID_PESSOA', 'ID_PESSOA_APRESENTADOR']) || atividade.ID_PESSOA_PRINCIPAL || '',
      RGA: core_domainsV2LegacyValue_(record, ['RGA', 'RGA_APRESENTADOR']) || atividade.RGA_PESSOA_PRINCIPAL || '',
      EMAIL: core_domainsV2LegacyValue_(record, ['EMAIL', 'EMAIL_APRESENTADOR']) || atividade.EMAIL_PESSOA_PRINCIPAL || '',
      DATA_ATIVIDADE: core_domainsV2LegacyValue_(record, ['DATA_ATIVIDADE', 'DATA_APRESENTACAO']) || atividade.DATA_ATIVIDADE || atividade.DATA_REALIZACAO || '',
      CICLO: core_domainsV2LegacyValue_(record, ['CICLO']) || atividade.CICLO || '',
      SEMESTRE: core_domainsV2LegacyValue_(record, ['SEMESTRE']) || atividade.SEMESTRE || '',
      ROTULO_SEMESTRE: core_domainsV2LegacyValue_(record, ['ROTULO_SEMESTRE', 'PERIODO_REFERENCIA']) ||
        atividade.ROTULO_SEMESTRE ||
        atividade.PERIODO_REFERENCIA ||
        ([atividade.ANO, atividade.SEMESTRE].filter(String).join('/') || '')
    });
    if (id) byId[id] = enriched;
    else withoutId.push(enriched);
  });

  return Object.keys(byId).map(function(id) { return byId[id]; }).concat(withoutId);
}

function core_domainsV2RecordMatchesPessoa_(record, ctx) {
  var id = String(core_domainsV2LegacyValue_(record, [
    'ID_PESSOA',
    'ID_PESSOA_APRESENTADOR',
    'ID_PESSOA_PRINCIPAL',
    'idPessoa',
    'idPessoaApresentador',
    'idPessoaPrincipal',
    'Id Pessoa'
  ]) || '').trim();
  var rga = core_domainsV2Rga_(core_domainsV2LegacyValue_(record, ['RGA', 'rga', 'rgaApresentador']));
  var email = core_domainsV2Email_(core_domainsV2LegacyValue_(record, [
    'EMAIL',
    'EMAIL_APRESENTADOR',
    'Email',
    'E-mail',
    'email',
    'emailApresentador'
  ]));
  return (!!ctx.idPessoa && id === ctx.idPessoa) || (!!ctx.rga && rga === ctx.rga) || (!!ctx.email && email === ctx.email);
}

function core_domainsV2StatusConcluido_(value) {
  var status = core_domainsV2AuditStatus_(value);
  return ['REALIZADA', 'CONCLUIDA', 'CONCLUÍDA', 'FINALIZADA', 'APROVADA', 'CONFIRMADA'].indexOf(status) >= 0;
}

function core_domainsV2StatusPendente_(value) {
  var status = core_domainsV2AuditStatus_(value);
  return ['PENDENTE', 'AGENDADA', 'EM_ABERTO', 'ABERTA', 'SOLICITADA', 'PLANEJADA'].indexOf(status) >= 0;
}

function core_domainsV2PresentationSummary_(activityData, ctx) {
  var records = core_domainsV2EnrichedPresentationRecords_(activityData);
  var concluded = [];
  var pending = false;
  var filePending = false;
  records.forEach(function(record) {
    if (!core_domainsV2RecordMatchesPessoa_(record, ctx)) return;
    var status = core_domainsV2LegacyValue_(record, ['STATUS_APRESENTACAO', 'STATUS', 'SITUACAO']);
    var fileStatus = core_domainsV2LegacyValue_(record, ['STATUS_ENVIO_MATERIAL', 'STATUS_ARQUIVO', 'ARQUIVO_STATUS', 'STATUS_DRIVE']);
    if (core_domainsV2StatusConcluido_(status)) concluded.push(record);
    if (core_domainsV2StatusPendente_(status)) pending = true;
    if (core_domainsV2StatusPendente_(fileStatus)) filePending = true;
  });
  concluded.sort(function(a, b) {
    var da = core_domainsV2Date_(core_domainsV2LegacyValue_(a, ['DATA_ATIVIDADE', 'DATA_APRESENTACAO', 'DATA_REALIZADA', 'DATA', 'DATA_INICIO', 'dataAtividade', 'dataApresentacao']));
    var db = core_domainsV2Date_(core_domainsV2LegacyValue_(b, ['DATA_ATIVIDADE', 'DATA_APRESENTACAO', 'DATA_REALIZADA', 'DATA', 'DATA_INICIO', 'dataAtividade', 'dataApresentacao']));
    return core_domainsV2ResumoTimestamp_(db) - core_domainsV2ResumoTimestamp_(da);
  });
  var last = concluded[0] || {};
  return {
    count: concluded.length,
    lastCycle: core_domainsV2LegacyValue_(last, ['CICLO', 'ID_CICLO', 'NOME_CICLO']),
    lastPeriod: core_domainsV2LegacyValue_(last, ['ROTULO_SEMESTRE', 'PERIODO', 'PERIODO_REFERENCIA', 'ID_PERIODO', 'SEMESTRE']),
    hasPending: pending,
    hasFilePending: filePending
  };
}

/**
 * Retorna links da pessoa resolvida por ID_PESSOA.
 * A leitura exposta filtra links publicos por padrao; uso privado exige opt-in
 * explicito por um backend ja autorizado.
 */
function corePessoasListLinksPerfis_(idPessoa, options) {
  var opts = options || {};
  var report = core_domainsV2NewReadReport_('PESSOAS_LIST_LINKS_PERFIS');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));
  return core_domainsV2LinksPerfisByPessoa_(pessoasData, idPessoa, {
    includeInactive: opts.includeInactive === true,
    publicOnly: opts.publicOnly !== false
  });
}

/** Retorna apenas links autorizados para superficies publicas do Portal. */
function corePessoasListLinksPerfisPublicos_(idPessoa) {
  return corePessoasListLinksPerfis_(idPessoa, { publicOnly: true });
}

function core_domainsV2FrequencySummary_(activityData, ctx, periodoReferencia) {
  var portal = activityData.portalFrequencia || [];
  for (var i = 0; i < portal.length; i++) {
    var record = portal[i];
    if (!core_domainsV2RecordMatchesPessoa_(record, ctx)) continue;
    var periodo = core_domainsV2LegacyValue_(record, ['CICLO', 'ROTULO_SEMESTRE', 'PERIODO', 'PERIODO_REFERENCIA', 'ID_PERIODO', 'SEMESTRE']);
    if (periodoReferencia && String(periodo || '').trim() !== String(periodoReferencia || '').trim()) continue;
    var ready = core_domainsV2LegacyValue_(record, ['FREQUENCIA_RESUMIDA', 'RESUMO_FREQUENCIA', 'FREQUENCIA', 'STATUS_FREQUENCIA']);
    if (ready) return String(ready);
    var percentual = core_domainsV2LegacyValue_(record, ['PERCENTUAL_FREQUENCIA', 'FREQUENCIA_PERCENTUAL']);
    var presencas = core_domainsV2LegacyValue_(record, ['TOTAL_PRESENCAS', 'PRESENCAS', 'QTD_PRESENCAS']);
    var faltas = core_domainsV2LegacyValue_(record, ['TOTAL_FALTAS', 'FALTAS', 'QTD_FALTAS']);
    var justificadas = core_domainsV2LegacyValue_(record, ['TOTAL_JUSTIFICADAS', 'JUSTIFICADAS', 'FALTAS_JUSTIFICADAS']);
    var abonadas = core_domainsV2LegacyValue_(record, ['TOTAL_ABONADAS', 'ABONADAS']);
    var faltasLiquidas = core_domainsV2LegacyValue_(record, ['FALTAS_LIQUIDAS']);
    var situacao = core_domainsV2LegacyValue_(record, ['SITUACAO_DISCIPLINAR', 'STATUS_DISCIPLINAR']);
    var percentualText = String(percentual === '' ? '' : percentual).trim();
    return [
      percentualText ? 'Frequencia ' + percentualText + (percentualText.indexOf('%') >= 0 ? '' : '%') : '',
      presencas !== '' ? 'Presencas ' + presencas : '',
      faltas !== '' ? 'Faltas ' + faltas : '',
      justificadas !== '' ? 'Justificadas ' + justificadas : '',
      abonadas !== '' ? 'Abonadas ' + abonadas : '',
      faltasLiquidas !== '' ? 'Faltas liquidas ' + faltasLiquidas : '',
      situacao ? 'Situacao ' + situacao : ''
    ].filter(String).join('; ');
  }
  var counts = { presenca: 0, falta: 0, justificada: 0 };
  (activityData.presencas || []).forEach(function(record) {
    if (!core_domainsV2RecordMatchesPessoa_(record, ctx)) return;
    var periodo = core_domainsV2LegacyValue_(record, ['PERIODO', 'PERIODO_REFERENCIA', 'ID_PERIODO', 'SEMESTRE']);
    if (periodoReferencia && String(periodo || '').trim() !== String(periodoReferencia || '').trim()) return;
    var status = core_domainsV2AuditStatus_(core_domainsV2LegacyValue_(record, ['STATUS_PRESENCA', 'STATUS', 'SITUACAO']));
    if (status.indexOf('JUST') >= 0) counts.justificada++;
    else if (status.indexOf('FALTA') >= 0 || status === 'AUSENTE') counts.falta++;
    else if (status.indexOf('PRES') >= 0 || status === 'PRESENTE') counts.presenca++;
  });
  if (!counts.presenca && !counts.falta && !counts.justificada) return '';
  return 'Presencas ' + counts.presenca + '; Faltas ' + counts.falta + '; Justificadas ' + counts.justificada;
}

/**
 * Conta justificativas ainda abertas para a pessoa.
 *
 * @param {Object} activityData
 * @param {Object} ctx
 * @return {number}
 */
function core_domainsV2PendingJustifications_(activityData, ctx) {
  var records = activityData.justificativas && activityData.justificativas.length
    ? activityData.justificativas
    : (activityData.portalJustificativas || []);
  return records.filter(function(record) {
    if (!core_domainsV2RecordMatchesPessoa_(record, ctx)) return false;
    var ativo = core_domainsV2AuditStatus_(core_domainsV2LegacyValue_(record, ['ATIVO']));
    if (ativo === 'NAO' || ativo === 'INATIVO') return false;
    var status = core_domainsV2AuditStatus_(core_domainsV2LegacyValue_(record, [
      'STATUS_ANALISE',
      'STATUS_ANALISE_JUSTIFICATIVA',
      'STATUS'
    ]));
    if (['DEFERIDA', 'DEFERIDO', 'INDEFERIDA', 'INDEFERIDO', 'APROVADA', 'APROVADO', 'RECUSADA', 'RECUSADO'].indexOf(status) >= 0) return false;
    return !status || ['PENDENTE', 'EM_ANALISE', 'AGUARDANDO_ANALISE', 'AJUSTE_SOLICITADO'].indexOf(status) >= 0;
  }).length;
}

function core_domainsV2HasCriticalFrequency_(frequencyText) {
  var text = core_domainsV2AuditStatus_(frequencyText);
  return text.indexOf('CRITICA') >= 0 || text.indexOf('CRÍTICA') >= 0 || text.indexOf('BAIXA') >= 0;
}

function core_domainsV2SuspensionFlag_(eventos, idPessoa) {
  return (eventos || []).some(function(record) {
    if (String(record.ID_PESSOA || '').trim() !== idPessoa) return false;
    var tipo = core_domainsV2AuditStatus_(record.TIPO_EVENTO);
    var status = core_domainsV2AuditStatus_(record.STATUS_EVENTO);
    return tipo.indexOf('SUSPEN') >= 0 && ['HOMOLOGADO', 'PROCESSADO', 'CONCLUIDO', 'CONCLUÍDO', 'APROVADO'].indexOf(status) >= 0;
  }) ? 'SIM' : 'NAO';
}

function core_domainsV2Eligibility_(vinculo) {
  var tipo = core_domainsV2NormalizeTipoVinculo_(vinculo.TIPO_VINCULO);
  var active = core_domainsV2Active_(vinculo);
  if (tipo === 'MEMBRO_EFETIVO' && active) return 'ELEGIVEL';
  if (tipo === 'MEMBRO_EM_ESPERA') return 'INELEGIVEL_MEMBRO_EM_ESPERA';
  if (tipo === 'EGRESSO') return 'INELEGIVEL_EX_MEMBRO';
  if (!active) return 'INELEGIVEL_SEM_VINCULO_ATIVO';
  return 'PENDENTE_ANALISE';
}

function core_domainsV2BuildPessoasResumoRows_(pessoasData, vigenciasData, options, report) {
  options = options || {};
  var today = new Date();
  var detalhesById = core_domainsV2IndexFirstBy_((pessoasData.MEMBROS_DETALHES && pessoasData.MEMBROS_DETALHES.records) || [], 'ID_PESSOA');
  var vinculosById = core_domainsV2IndexManyBy_((pessoasData.VINCULOS_GEAPA && pessoasData.VINCULOS_GEAPA.records) || [], 'ID_PESSOA');
  var existingResumoById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_RESUMO_OPERACIONAL && pessoasData.PESSOAS_RESUMO_OPERACIONAL.records) || [], 'ID_PESSOA');
  var vigResumoById = core_domainsV2IndexFirstBy_((vigenciasData.VIGENCIAS_RESUMO_ATUAL && vigenciasData.VIGENCIAS_RESUMO_ATUAL.records) || [], 'ID_PESSOA');
  var eventos = (pessoasData.MEMBROS_EVENTOS_VINCULO && pessoasData.MEMBROS_EVENTOS_VINCULO.records) || [];
  var activityData = core_domainsV2ActivityData_();
  var unavailable = activityData.unavailable || {};
  unavailable.DATA_LIMITE_ESTIMADA_DIRETORIA = 'Regra de limite de diretoria ainda nao esta segura para calculo automatico.';

  var rows = ((pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || []).filter(function(pessoa) {
    return !options.idPessoa || String(pessoa.ID_PESSOA || '').trim() === String(options.idPessoa || '').trim();
  }).map(function(pessoa) {
    var idPessoa = String(pessoa.ID_PESSOA || '').trim();
    var detalhes = detalhesById[idPessoa] || {};
    var vinculos = vinculosById[idPessoa] || [];
    var vinculo = core_domainsV2PickCurrentVinculo_(vinculos) || {};
    var existingResumo = existingResumoById[idPessoa] || {};
    var vigResumo = vigResumoById[idPessoa] || {};
    var rga = core_domainsV2GetRga_(pessoasData, idPessoa, detalhes);
    var email = pessoa.EMAIL_PRINCIPAL || '';
    var ctx = { idPessoa: idPessoa, rga: core_domainsV2Rga_(rga), email: core_domainsV2Email_(email) };
    var intervals = core_domainsV2EffectiveMemberIntervals_(vinculos, today);
    var presentation = core_domainsV2PresentationSummary_(activityData, ctx);
    var frequency = core_domainsV2FrequencySummary_(activityData, ctx, options.periodoReferencia);
    var pendingJustifications = core_domainsV2PendingJustifications_(activityData, ctx);
    var portal = core_domainsV2BuildPortalState_(vinculo, vigResumo, core_domainsV2ActivePortalExceptions_(pessoasData, idPessoa, today));
    var tipoAtual = vinculo.TIPO_VINCULO || '';
    var normalizedTipoAtual = core_domainsV2NormalizeTipoVinculo_(tipoAtual);
    var activeCurrent = core_domainsV2Active_(vinculo);
    var cargoAtual = vigResumo.CARGO_FUNCAO_ATUAL || '';
    var pendencias = [];
    if (!email) pendencias.push('SEM_EMAIL');
    if (normalizedTipoAtual === 'MEMBRO_EFETIVO' && !rga) pendencias.push('SEM_RGA');
    if (!pessoa.NOME_EXIBICAO && !pessoa.NOME_COMPLETO) pendencias.push('CADASTRO_INCOMPLETO');
    if (presentation.hasPending) pendencias.push('APRESENTACAO_PENDENTE');
    if (presentation.hasFilePending) pendencias.push('ARQUIVO_APRESENTACAO_PENDENTE');
    if (core_domainsV2HasCriticalFrequency_(frequency)) pendencias.push('FREQUENCIA_CRITICA');
    if (frequency && core_domainsV2AuditStatus_(frequency).indexOf('JUSTIFICATIVA_PENDENTE') >= 0) pendencias.push('JUSTIFICATIVA_PENDENTE');
    if (pendingJustifications > 0 && pendencias.indexOf('JUSTIFICATIVA_PENDENTE') < 0) pendencias.push('JUSTIFICATIVA_PENDENTE');
    if (!cargoAtual && normalizedTipoAtual === 'MEMBRO_EFETIVO' && activeCurrent) cargoAtual = 'Membro';

    return {
      ID_PESSOA: idPessoa,
      RGA: rga,
      NOME_EXIBICAO: pessoa.NOME_EXIBICAO || pessoa.NOME_COMPLETO || '',
      EMAIL: email,
      TIPO_VINCULO_ATUAL: tipoAtual,
      STATUS_VINCULO_ATUAL: vinculo.STATUS_VINCULO || '',
      CARGO_FUNCAO_ATUAL: cargoAtual,
      PERFIL_PORTAL_CALCULADO: portal.perfil,
      PORTAL_ATIVO: portal.portalAtivo,
      TEMPO_EFETIVO_NO_GRUPO: core_domainsV2FormatDuration_(core_domainsV2CountIntervalDays_(intervals)),
      QTD_SEMESTRES_NO_GRUPO: core_domainsV2CountSemestersForIntervals_(vigenciasData, intervals),
      QTD_APRESENTACOES_REALIZADAS: presentation.count,
      CICLO_ULTIMA_APRESENTACAO: presentation.lastCycle || '',
      PERIODO_ULTIMA_APRESENTACAO: presentation.lastPeriod || '',
      FREQUENCIA_RESUMIDA: frequency || '',
      PENDENCIAS_ABERTAS: pendencias.length ? pendencias.join('; ') : 'SEM_PENDENCIAS',
      FLAG_JA_FOI_SUSPENSO: core_domainsV2SuspensionFlag_(eventos, idPessoa),
      STATUS_ELEGIBILIDADE_DIRETORIA: existingResumo.STATUS_ELEGIBILIDADE_DIRETORIA || core_domainsV2Eligibility_(vinculo),
      DATA_LIMITE_ESTIMADA_DIRETORIA: existingResumo.DATA_LIMITE_ESTIMADA_DIRETORIA || '',
      ULTIMA_ATUALIZACAO: new Date()
    };
  });
  report.fontesIndisponiveis = unavailable;
  return rows;
}

function coreRecalcularPessoasResumoOperacionalV2_(options) {
  var opts = core_domainsV2ResumoOptions_(options);
  opts.idPessoa = options && options.idPessoa ? String(options.idPessoa).trim() : '';
  opts.periodoReferencia = options && options.periodoReferencia ? String(options.periodoReferencia).trim() : '';
  if (!opts.dryRun && opts.confirmacao !== 'RECALCULAR_PESSOAS_RESUMO_V2') {
    throw new Error('Para escrever PESSOAS_RESUMO_OPERACIONAL, informe confirmacao: "RECALCULAR_PESSOAS_RESUMO_V2".');
  }

  var report = core_domainsV2AuditNewReport_('RECALCULAR_PESSOAS_RESUMO_OPERACIONAL_V2');
  report.dryRun = opts.dryRun;
  var pessoasData = core_domainsV2OpenPessoas_(report);
  var vigenciasData = core_domainsV2OpenVigencias_(report);
  if (report.totalErros) return report;

  var rows = core_domainsV2BuildPessoasResumoRows_(pessoasData, vigenciasData, opts, report);
  var headers = pessoasData.PESSOAS_RESUMO_OPERACIONAL.headers || [];
  var stats = core_domainsV2PessoasResumoStats_(rows);
  var fieldStats = core_domainsV2PessoasResumoFieldStats_(rows, report.fontesIndisponiveis || {});
  report.camposNaoCalculaveis = fieldStats.naoCalculaveis;
  report.resumoQuantitativo = {
    totalPessoasAnalisadas: rows.length,
    totalMembrosAtivos: stats.totalMembrosAtivos,
    totalExMembros: stats.totalExMembros,
    totalMembrosEmEspera: stats.totalMembrosEmEspera,
    pessoasBase: (pessoasData.PESSOAS_BASE.records || []).length,
    linhasCalculadas: rows.length,
    linhasExistentesAntes: (pessoasData.PESSOAS_RESUMO_OPERACIONAL.records || []).length,
    resumosAtualizados: opts.dryRun ? 0 : rows.length,
    camposPreenchidos: fieldStats.preenchidos,
    camposSemValor: fieldStats.semValor
  };
  report.options = {
    idPessoa: opts.idPessoa || null,
    periodoReferencia: opts.periodoReferencia || null
  };
  core_domainsV2AttachLegacyComparisonSummary_(report);

  if (opts.dryRun) {
    report.amostra = rows.slice(0, 10);
    core_domainsV2AuditRecommendation_(report, 'Conferir a amostra; se estiver correta, rodar coreRecalcularPessoasResumoOperacionalV2({ dryRun: false, confirmacao: "RECALCULAR_PESSOAS_RESUMO_V2" }).');
    return report;
  }

  if (rows.length) {
    core_domainsV2WritePessoasResumoRows_(pessoasData.PESSOAS_RESUMO_OPERACIONAL, headers, rows, opts);
  }
  report.resumoQuantitativo.linhasEscritas = rows.length;
  return report;
}

function core_domainsV2AttachLegacyComparisonSummary_(report) {
  try {
    var comparison = coreCompararLegadoComV2_({ limit: 10 });
    report.divergenciasLegado = {
      totalErros: comparison.totalErros,
      totalAvisos: comparison.totalAvisos,
      resumoQuantitativo: comparison.resumoQuantitativo || {},
      avisos: (comparison.avisos || []).slice(0, 10)
    };
  } catch (err) {
    report.divergenciasLegado = {
      comparavel: false,
      motivo: err && err.message ? err.message : String(err)
    };
  }
}

function core_domainsV2PessoasResumoStats_(rows) {
  var stats = { totalMembrosAtivos: 0, totalExMembros: 0, totalMembrosEmEspera: 0 };
  (rows || []).forEach(function(row) {
    var tipo = core_domainsV2NormalizeTipoVinculo_(row.TIPO_VINCULO_ATUAL);
    var status = core_domainsV2AuditStatus_(row.STATUS_VINCULO_ATUAL);
    if (tipo === 'MEMBRO_EFETIVO' && (status === 'ATIVO' || row.PORTAL_ATIVO === 'SIM')) stats.totalMembrosAtivos++;
    if (tipo === 'EGRESSO') stats.totalExMembros++;
    if (tipo === 'MEMBRO_EM_ESPERA') stats.totalMembrosEmEspera++;
  });
  return stats;
}

/**
 * Resume o preenchimento dos campos derivados do cache operacional.
 *
 * @param {Array<Object>} rows
 * @param {Object} unavailable
 * @return {Object}
 */
function core_domainsV2PessoasResumoFieldStats_(rows, unavailable) {
  var fields = [
    'ID_PESSOA',
    'RGA',
    'NOME_EXIBICAO',
    'EMAIL',
    'TIPO_VINCULO_ATUAL',
    'STATUS_VINCULO_ATUAL',
    'CARGO_FUNCAO_ATUAL',
    'PERFIL_PORTAL_CALCULADO',
    'PORTAL_ATIVO',
    'TEMPO_EFETIVO_NO_GRUPO',
    'QTD_SEMESTRES_NO_GRUPO',
    'QTD_APRESENTACOES_REALIZADAS',
    'CICLO_ULTIMA_APRESENTACAO',
    'FREQUENCIA_RESUMIDA',
    'PENDENCIAS_ABERTAS',
    'STATUS_ELEGIBILIDADE_DIRETORIA'
  ];
  var preenchidos = {};
  var semValor = {};
  fields.forEach(function(field) {
    preenchidos[field] = 0;
    semValor[field] = 0;
  });
  (rows || []).forEach(function(row) {
    fields.forEach(function(field) {
      if (row[field] === 0 || String(row[field] || '').trim() !== '') preenchidos[field]++;
      else semValor[field]++;
    });
  });
  var naoCalculaveis = fields.filter(function(field) {
    return semValor[field] > 0;
  }).map(function(field) {
    return {
      campo: field,
      linhasSemValor: semValor[field],
      motivo: unavailable[field] || 'Fonte ausente, identidade nao conciliada ou campo nao aplicavel para parte das pessoas.'
    };
  });
  Object.keys(unavailable || {}).forEach(function(source) {
    naoCalculaveis.push({ campo: source, linhasSemValor: null, motivo: unavailable[source] });
  });
  return { preenchidos: preenchidos, semValor: semValor, naoCalculaveis: naoCalculaveis };
}

function core_domainsV2WritePessoasResumoRows_(targetData, headers, rows, opts) {
  var sheet = targetData.sheet;
  if (opts.idPessoa) {
    var existingById = core_domainsV2IndexFirstBy_(targetData.records || [], 'ID_PESSOA');
    rows.forEach(function(row) {
      var existingRow = existingById[String(row.ID_PESSOA || '').trim()];
      if (existingRow && existingRow.__rowNumber) {
        var mergedRow = Object.assign({}, core_domainsV2CloneRecord_(existingRow), row);
        sheet.getRange(existingRow.__rowNumber, 1, 1, headers.length).setValues([core_buildRowFromObjectByHeaders_(headers, mergedRow)]);
      } else {
        sheet.appendRow(core_buildRowFromObjectByHeaders_(headers, row));
      }
    });
    return rows.length;
  }

  var existing = (targetData.records || []).map(function(record) {
    return core_domainsV2CloneRecord_(record);
  });
  var existingIndexById = {};
  existing.forEach(function(record, index) {
    var id = String(record.ID_PESSOA || '').trim();
    if (id) existingIndexById[id] = index;
  });
  var appendRows = [];
  rows.forEach(function(row) {
    var id = String(row.ID_PESSOA || '').trim();
    if (Object.prototype.hasOwnProperty.call(existingIndexById, id)) {
      existing[existingIndexById[id]] = Object.assign({}, existing[existingIndexById[id]], row);
    } else {
      appendRows.push(row);
    }
  });
  if (existing.length) {
    sheet.getRange(2, 1, existing.length, headers.length).setValues(existing.map(function(row) {
      return core_buildRowFromObjectByHeaders_(headers, row);
    }));
  }
  appendRows.forEach(function(row) {
    sheet.appendRow(core_buildRowFromObjectByHeaders_(headers, row));
  });
  return rows.length;
}

function coreDiagnosticarPessoasResumoOperacionalV2_(options) {
  options = options || {};
  var report = core_domainsV2AuditNewReport_('DIAGNOSTICAR_PESSOAS_RESUMO_OPERACIONAL_V2');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  var vigenciasData = core_domainsV2OpenVigencias_(report);
  if (report.totalErros) return report;

  var baseById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [], 'ID_PESSOA');
  var detalhesById = core_domainsV2IndexFirstBy_((pessoasData.MEMBROS_DETALHES && pessoasData.MEMBROS_DETALHES.records) || [], 'ID_PESSOA');
  var vinculosById = core_domainsV2IndexManyBy_((pessoasData.VINCULOS_GEAPA && pessoasData.VINCULOS_GEAPA.records) || [], 'ID_PESSOA');
  var resumoById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_RESUMO_OPERACIONAL && pessoasData.PESSOAS_RESUMO_OPERACIONAL.records) || [], 'ID_PESSOA');
  var vigResumoById = core_domainsV2IndexFirstBy_((vigenciasData.VIGENCIAS_RESUMO_ATUAL && vigenciasData.VIGENCIAS_RESUMO_ATUAL.records) || [], 'ID_PESSOA');

  Object.keys(baseById).forEach(function(idPessoa) {
    if (!resumoById[idPessoa]) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'PESSOA_SEM_RESUMO', 'Pessoa sem linha em PESSOAS_RESUMO_OPERACIONAL.', { idPessoa: idPessoa });
    }
  });

  Object.keys(resumoById).forEach(function(idPessoa) {
    var resumo = resumoById[idPessoa];
    var pessoa = baseById[idPessoa];
    var detalhes = detalhesById[idPessoa] || {};
    var vinculo = core_domainsV2PickCurrentVinculo_(vinculosById[idPessoa] || []) || {};
    var tipo = core_domainsV2NormalizeTipoVinculo_(vinculo.TIPO_VINCULO || resumo.TIPO_VINCULO_ATUAL);
    var active = core_domainsV2Active_(vinculo);
    var vigResumo = vigResumoById[idPessoa] || {};
    if (!pessoa) {
      core_domainsV2AuditIssue_(report, 'ERRO', 'RESUMO_SEM_PESSOA', 'Resumo operacional aponta para pessoa inexistente.', { idPessoa: idPessoa });
      return;
    }
    if (tipo === 'MEMBRO_EFETIVO' && active && resumo.PORTAL_ATIVO !== 'SIM') {
      core_domainsV2AuditIssue_(report, 'AVISO', 'MEMBRO_ATIVO_SEM_PORTAL_ATIVO', 'Membro efetivo ativo sem PORTAL_ATIVO=SIM.', { idPessoa: idPessoa });
    }
    if (tipo === 'EGRESSO' && resumo.PORTAL_ATIVO === 'SIM') {
      core_domainsV2AuditIssue_(report, 'AVISO', 'EGRESSO_COM_PORTAL_ATIVO', 'Egresso com PORTAL_ATIVO=SIM; verificar excecao formal.', { idPessoa: idPessoa });
    }
    if (tipo === 'MEMBRO_EFETIVO' && active && !core_domainsV2GetRga_(pessoasData, idPessoa, detalhes)) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'MEMBRO_ATIVO_SEM_RGA', 'Membro efetivo ativo sem RGA.', { idPessoa: idPessoa });
    }
    if (tipo === 'MEMBRO_EFETIVO' && active && !pessoa.EMAIL_PRINCIPAL) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'MEMBRO_ATIVO_SEM_EMAIL', 'Membro efetivo ativo sem e-mail principal.', { idPessoa: idPessoa });
    }
    if (vigResumo.CARGO_FUNCAO_ATUAL && resumo.CARGO_FUNCAO_ATUAL !== vigResumo.CARGO_FUNCAO_ATUAL) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'CARGO_ATUAL_DIVERGENTE_VIGENCIAS', 'Cargo atual no resumo de Pessoas diverge de VIGENCIAS_RESUMO_ATUAL.', {
        idPessoa: idPessoa,
        pessoas: resumo.CARGO_FUNCAO_ATUAL,
        vigencias: vigResumo.CARGO_FUNCAO_ATUAL
      });
    }
    if (vigResumo.PERFIS_PORTAL_CALCULADOS && String(resumo.PERFIL_PORTAL_CALCULADO || '').indexOf(String(vigResumo.PERFIS_PORTAL_CALCULADOS || '')) < 0) {
      core_domainsV2AuditIssue_(report, 'AVISO', 'PERFIL_PORTAL_DIVERGENTE_VIGENCIAS', 'Perfil portal no resumo de Pessoas diverge de VIGENCIAS_RESUMO_ATUAL.', {
        idPessoa: idPessoa,
        pessoas: resumo.PERFIL_PORTAL_CALCULADO,
        vigencias: vigResumo.PERFIS_PORTAL_CALCULADOS
      });
    }
    ['QTD_SEMESTRES_NO_GRUPO', 'QTD_APRESENTACOES_REALIZADAS', 'FREQUENCIA_RESUMIDA', 'PENDENCIAS_ABERTAS'].forEach(function(field) {
      if (String(resumo[field] || '').trim() === '') {
        core_domainsV2AuditIssue_(report, 'AVISO', field + '_VAZIO', field + ' vazio em PESSOAS_RESUMO_OPERACIONAL.', { idPessoa: idPessoa });
      }
    });
  });

  report.resumoQuantitativo = {
    pessoasBase: Object.keys(baseById).length,
    resumos: Object.keys(resumoById).length,
    totalErros: report.totalErros,
    totalAvisos: report.totalAvisos
  };
  return report;
}

function core_domainsV2ReadLegacyRecordsSafe_(key) {
  try {
    return core_readRecordsByKey_(key, { skipBlankRows: true }) || [];
  } catch (err) {
    return [];
  }
}

function core_domainsV2LegacyValue_(record, headers) {
  var keys = Object.keys(record || {});
  for (var i = 0; i < headers.length; i++) {
    var target = core_normalizeHeader_(headers[i]);
    for (var j = 0; j < keys.length; j++) {
      if (core_normalizeHeader_(keys[j]) === target) return record[keys[j]];
    }
  }
  return '';
}

function core_domainsV2CompareLegacyList_(report, label, legacyRecords, pessoasData, expectedType, limit) {
  var missing = [];
  legacyRecords.forEach(function(record) {
    var email = core_domainsV2Email_(core_domainsV2LegacyValue_(record, ['EMAIL', 'Email', 'E-mail']));
    var rga = core_domainsV2Rga_(core_domainsV2LegacyValue_(record, ['RGA']));
    var idPessoa = rga
      ? core_domainsV2FindPessoaIdByRga_(pessoasData, rga)
      : core_domainsV2FindPessoaIdByEmail_(pessoasData, email);
    if (!idPessoa) {
      missing.push({ rga: rga, email: email, row: record.__rowNumber });
      return;
    }
    var vinculos = ((pessoasData.VINCULOS_GEAPA && pessoasData.VINCULOS_GEAPA.records) || []).filter(function(vinculo) {
      return String(vinculo.ID_PESSOA || '').trim() === idPessoa &&
        core_domainsV2AuditTipoVinculo_(vinculo.TIPO_VINCULO) === expectedType &&
        core_domainsV2Active_(vinculo);
    });
    if (!vinculos.length) {
      missing.push({ idPessoa: idPessoa, rga: rga, email: email, row: record.__rowNumber, reason: 'SEM_VINCULO_ATIVO_' + expectedType });
    }
  });
  report.resumoQuantitativo[label] = {
    legado: legacyRecords.length,
    divergencias: missing.length
  };
  if (missing.length) {
    core_domainsV2AuditIssue_(report, 'AVISO', 'LEGADO_V2_DIVERGENCIA_' + label, 'Divergencias diagnosticas entre legado e v2.', {
      total: missing.length,
      amostra: missing.slice(0, limit || 20)
    });
  }
}

function coreCompararLegadoComV2_(opts) {
  opts = opts || {};
  var report = core_domainsV2AuditNewReport_('COMPARAR_LEGADO_COM_V2');
  var pessoasData = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) return report;
  var limit = Number(opts.limit || 20);

  core_domainsV2CompareLegacyList_(report, 'MEMBROS_ATUAIS', core_domainsV2ReadLegacyRecordsSafe_('MEMBERS_ATUAIS'), pessoasData, 'MEMBRO_EFETIVO', limit);
  core_domainsV2CompareLegacyList_(report, 'MEMBROS_HIST', core_domainsV2ReadLegacyRecordsSafe_('MEMBERS_HIST'), pessoasData, 'EGRESSO', limit);
  core_domainsV2CompareLegacyList_(report, 'MEMBROS_FUTURO', core_domainsV2ReadLegacyRecordsSafe_('MEMBERS_FUTURO'), pessoasData, 'MEMBRO_EM_ESPERA', limit);
  return report;
}
