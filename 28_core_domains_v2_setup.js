/**
 * ============================================================
 * 28_core_domains_v2_setup.js
 * ============================================================
 *
 * Contratos permanentes das planilhas v2 dos dominios centrais
 * PESSOAS e VIGENCIAS.
 *
 * As funcoes historicas de preparacao/ensure inicial permanecem internas
 * por compatibilidade do codigo, mas nao sao API publica depois da migracao
 * inicial legado -> v2.
 */

var CORE_DOMAINS_V2_INITIAL_SHEET_NAMES = Object.freeze([
  'Página1',
  'Pagina1',
  'Page1',
  'Sheet1'
]);

var CORE_DOMAINS_V2_HEADER_NOTES = Object.freeze({
  ID_PESSOA: 'Identificador tecnico unificado da pessoa no ecossistema GEAPA.',
  ID_LINK: 'Identificador tecnico unico do link/perfil associado a uma pessoa.',
  ID_VINCULO: 'Identificador tecnico do vinculo da pessoa com o GEAPA.',
  ID_EVENTO_MEMBRO: 'Identificador tecnico do evento de ciclo de vida do membro.',
  ID_VIGENCIA: 'Identificador tecnico da vigencia de funcao/cargo.',
  RGA: 'Identificador oficial de membros discentes; nao substituir por ID_PESSOA nesta fase.',
  EMAIL: 'E-mail de contato normalizado quando aplicavel.',
  EMAIL_PRINCIPAL: 'E-mail principal da pessoa.',
  TIPO_LINK: 'Categoria canonica do link: LATTES, LINKEDIN, ORCID, INSTAGRAM, SITE_PESSOAL, GOOGLE_SCHOLAR, RESEARCHGATE ou OUTRO.',
  URL: 'Endereco do link. Leituras para o Portal aceitam somente URLs HTTP(S) validas.',
  ROTULO: 'Rotulo legivel do link para exibicao no Portal.',
  PUBLICAVEL: 'Use SIM somente quando o link puder integrar leituras publicas.',
  VISIVEL_PORTAL: 'Use SIM somente quando o link puder ser exibido em superficies do Portal.',
  FONTE: 'Origem cadastral ou processo que forneceu o link.',
  VALIDADO_EM: 'Data em que a URL foi validada manualmente ou por processo confiavel.',
  ATIVO: 'Use SIM/NAO para indicar se o registro esta operacionalmente ativo.',
  CRIADO_EM: 'Timestamp de criacao do registro.',
  ATUALIZADO_EM: 'Timestamp da ultima atualizacao do registro.',
  STATUS_CADASTRAL: 'Estado cadastral da pessoa; nao representa frequencia ou cargo.',
  STATUS_VINCULO: 'Estado consolidado do vinculo com o GEAPA.',
  STATUS_EVENTO: 'Estado do evento; apenas eventos homologados/processaveis devem gerar efeitos.',
  STATUS_VIGENCIA: 'Estado da vigencia temporal de cargo/funcao.',
  TIPO_VINCULO: 'Tipo relacional da pessoa com o GEAPA.',
  TIPO_EVENTO: 'Tipo do evento de ciclo de vida do membro.',
  TIPO_FUNCAO: 'Tipo de funcao temporal; nao cria categoria nova de membro.',
  CARGO_KEY: 'Chave canonica do cargo conforme CARGOS_CONFIG.',
  PERFIL_PORTAL_CALCULADO: 'Perfil derivado para portal; nao cria cargo institucional.',
  PERMISSOES_CALCULADAS: 'Permissoes derivadas; nao devem ser editadas como regra normativa.',
  FREQUENCIA_RESUMIDA: 'Resumo vindo de Atividades; Pessoas nao e fonte de presenca/falta.',
  QTD_APRESENTACOES_REALIZADAS: 'Resumo vindo de Atividades; Pessoas nao e fonte de apresentacoes.',
  CARGO_FUNCAO_ATUAL: 'Resumo vindo de Vigencias; Pessoas nao e fonte de cargos/funcoes.',
  OBS_INTERNA: 'Observacao interna operacional.',
  OBS: 'Observacoes gerais.'
});

var CORE_DOMAINS_V2_SCHEMAS = Object.freeze({
  PESSOAS: Object.freeze([
    Object.freeze({
      sheetName: 'PESSOAS_BASE',
      classification: 'FONTE',
      headers: Object.freeze([
        'ID_PESSOA',
        'NOME_COMPLETO',
        'NOME_CIVIL',
        'NOME_EXIBICAO',
        'EMAIL_PRINCIPAL',
        'TELEFONE_PRINCIPAL',
        'CPF',
        'DATA_NASCIMENTO',
        'INSTAGRAM',
        'CIDADE_NATAL',
        'UF_ORIGEM',
        'SEXO',
        'STATUS_CADASTRAL',
        'OBS_INTERNA',
        'CRIADO_EM',
        'ATUALIZADO_EM',
        'ATIVO'
      ]),
      optionalHeaders: Object.freeze([
        'NOME_CIVIL'
      ])
    }),
    Object.freeze({
      sheetName: 'PESSOAS_IDENTIFICADORES',
      classification: 'FONTE',
      headers: Object.freeze([
        'ID_IDENTIFICADOR',
        'ID_PESSOA',
        'TIPO_IDENTIFICADOR',
        'VALOR_IDENTIFICADOR',
        'PRINCIPAL',
        'ATIVO',
        'OBS'
      ])
    }),
    Object.freeze({
      sheetName: 'PESSOAS_LINKS_PERFIS',
      classification: 'FONTE',
      // Optional during rollout so existing Portal reads keep the legacy Lattes fallback.
      optional: true,
      headers: Object.freeze([
        'ID_LINK',
        'ID_PESSOA',
        'TIPO_LINK',
        'URL',
        'ROTULO',
        'PUBLICAVEL',
        'VISIVEL_PORTAL',
        'FONTE',
        'VALIDADO_EM',
        'ATIVO',
        'CRIADO_EM',
        'ATUALIZADO_EM',
        'OBS'
      ])
    }),
    Object.freeze({
      sheetName: 'MEMBROS_DETALHES',
      classification: 'FONTE',
      headers: Object.freeze([
        'ID_PESSOA',
        'RGA',
        'SEMESTRE_ENTRADA',
        'SEMESTRE_ATUAL',
        'DATA_INTEGRACAO_ORIGINAL',
        'HISTORICO_ATIVIDADES_ACADEMICAS',
        'OBS_MEMBRO',
        'ATUALIZADO_EM'
      ]),
      optionalHeaders: Object.freeze([
        'ATUALIZADO_EM'
      ])
    }),
    Object.freeze({
      sheetName: 'SOLICITACOES_ATUALIZACAO_CADASTRAL',
      classification: 'FONTE_EVENTOS',
      optional: true,
      headers: Object.freeze([
        'ID_SOLICITACAO',
        'ID_PESSOA',
        'TIPO_SOLICITACAO',
        'CAMPO',
        'VALOR_ATUAL_MASCARADO',
        'VALOR_ATUAL_HASH',
        'VALOR_SOLICITADO',
        'JUSTIFICATIVA',
        'STATUS',
        'SOLICITADO_EM',
        'SOLICITADO_POR',
        'ANALISADO_EM',
        'ANALISADO_POR',
        'DECISAO',
        'MOTIVO_DECISAO',
        'APLICADO_EM',
        'ID_LOG',
        'CHAVE_IDEMPOTENCIA',
        'ATIVO',
        'CRIADO_EM',
        'ATUALIZADO_EM'
      ])
    }),
    Object.freeze({
      sheetName: 'COLABORADORES_ACADEMICOS',
      classification: 'FONTE',
      headers: Object.freeze([
        'ID_PESSOA',
        'ID_PROFESSOR',
        'TIPO_COLABORADOR',
        'INSTITUICAO',
        'SETOR',
        'TITULACAO',
        'AREA_ATUACAO',
        'EIXO_ASSOCIADO',
        'EMAIL_INSTITUCIONAL',
        'OBS_ACADEMICA',
        'ATIVO'
      ]),
      deprecatedHeaders: Object.freeze([
        'EIXO_ASSOCIADO_1',
        'EIXO_ASSOCIADO_2'
      ])
    }),
    Object.freeze({
      sheetName: 'PARTICIPANTES_EXTERNOS_DETALHES',
      classification: 'FONTE',
      headers: Object.freeze([
        'ID_PESSOA',
        'ID_PARTICIPANTE_EXTERNO',
        'TIPO_PUBLICO',
        'INSTITUICAO_ORIGEM',
        'EMPRESA_ORIGEM',
        'CARGO_PROFISSAO',
        'CURSO',
        'CIDADE_ATUAL',
        'EVENTO_ORIGEM',
        'AUTORIZA_CONTATO',
        'OBS_EXTERNA',
        'ATIVO'
      ])
    }),
    Object.freeze({
      sheetName: 'VINCULOS_GEAPA',
      classification: 'FONTE',
      headers: Object.freeze([
        'ID_VINCULO',
        'ID_PESSOA',
        'TIPO_VINCULO',
        'STATUS_VINCULO',
        'DATA_INICIO',
        'DATA_FIM',
        'MOTIVO_INICIO',
        'MOTIVO_FIM',
        'FONTE',
        'LINK_ATA_OU_PROCESSO',
        'OBS_PUBLICA',
        'OBS_INTERNA',
        'ATIVO'
      ])
    }),
    Object.freeze({
      sheetName: 'MEMBROS_EVENTOS_VINCULO',
      classification: 'FONTE_EVENTOS',
      headers: Object.freeze([
        'ID_EVENTO_MEMBRO',
        'RGA',
        'ID_PESSOA',
        'TIPO_EVENTO',
        'DATA_EVENTO',
        'STATUS_EVENTO',
        'MODULO_ORIGEM',
        'CHAVE_ORIGEM',
        'OBSERVACOES',
        'ATUALIZADO_EM',
        'PROCESSADO_POR_MODULO',
        'DATA_PROCESSAMENTO',
        'ERRO_PROCESSAMENTO'
      ])
    }),
    Object.freeze({
      sheetName: 'PESSOAS_COMUNICACAO_CONSENTIMENTOS',
      classification: 'FONTE',
      headers: Object.freeze([
        'ID_PESSOA',
        'EMAIL',
        'RECEBE_COMUNICACOES_GEAPA',
        'STATUS_COMUNICACAO',
        'EIXOS_INTERESSE',
        'INTERESSE_EIXO_I',
        'INTERESSE_EIXO_II',
        'INTERESSE_EIXO_III',
        'INTERESSE_EIXO_IV',
        'INTERESSE_EIXO_V',
        'INTERESSE_EIXO_VI',
        'INTERESSE_EIXO_VII',
        'INTERESSE_EIXO_VIII',
        'ORIGEM_CONSENTIMENTO',
        'DATA_CONSENTIMENTO',
        'DATA_REVOGACAO',
        'OBS'
      ])
    }),
    Object.freeze({
      sheetName: 'PORTAL_ACESSOS_EXCECOES',
      classification: 'FONTE',
      headers: Object.freeze([
        'ID_EXCECAO',
        'ID_PESSOA',
        'PERFIL_EXTRA',
        'PERMISSAO_EXTRA',
        'DATA_INICIO',
        'DATA_FIM',
        'STATUS',
        'JUSTIFICATIVA',
        'APROVADO_POR',
        'LINK_ATA',
        'OBS'
      ])
    }),
    Object.freeze({
      sheetName: 'PESSOAS_RESUMO_OPERACIONAL',
      classification: 'CACHE',
      headers: Object.freeze([
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
        'FLAG_JA_FOI_SUSPENSO',
        'STATUS_ELEGIBILIDADE_DIRETORIA',
        'DATA_LIMITE_ESTIMADA_DIRETORIA',
        'ULTIMA_ATUALIZACAO'
      ])
    })
  ]),
  VIGENCIAS: Object.freeze([
    Object.freeze({
      sheetName: 'SEMESTRES',
      classification: 'FONTE',
      readAliases: Object.freeze({
        ID_CICLO: Object.freeze(['ID_PERIODO'])
      }),
      headers: Object.freeze([
        'ID_SEMESTRE',
        'ANO',
        'SEMESTRE',
        'DATA_INICIO',
        'DATA_FIM',
        'ID_CICLO',
        'NUMERO_REUNIOES_PREVISTAS',
        'INICIO_MATRICULAS_ONLINE',
        'FIM_MATRICULAS_ONLINE',
        'INICIO_AJUSTE_ALUNO',
        'FIM_AJUSTE_ALUNO',
        'INICIO_AJUSTE_COORDENADOR',
        'FIM_AJUSTE_COORDENADOR',
        'STATUS',
        'OBS'
      ])
    }),
    Object.freeze({
      sheetName: 'CICLOS',
      classification: 'FONTE',
      readAliases: Object.freeze({
        ID_CICLO: Object.freeze(['ID_PERIODO']),
        NOME_CICLO: Object.freeze(['NOME_PERIODO']),
        TIPO_CICLO: Object.freeze(['TIPO_PERIODO'])
      }),
      headers: Object.freeze([
        'ID_CICLO',
        'NOME_CICLO',
        'TIPO_CICLO',
        'DATA_INICIO',
        'DATA_FIM',
        'NUMERO_MEMBROS_PREVISTOS',
        'TOTAL_ATIVIDADES_QUE_CONTAM_FALTA_PLANEJADAS',
        'LIMITE_FALTAS_PERIODO_CONGELADO',
        'DATA_FECHAMENTO_PLANEJAMENTO',
        'NUMERO_SEI',
        'ANO_SEI',
        'COORDENADOR',
        'STATUS',
        'OBS'
      ])
    }),
    Object.freeze({
      sheetName: 'DIRETORIAS',
      classification: 'FONTE',
      headers: Object.freeze([
        'ID_DIRETORIA',
        'NOME_GESTAO',
        'DATA_INICIO',
        'DATA_FIM_PREVISTA',
        'DATA_FIM_REAL',
        'STATUS_DIRETORIA',
        'LINK_ATA_POSSE',
        'LEMA',
        'OBS'
      ])
    }),
    Object.freeze({
      sheetName: 'SEMESTRES_DIRETORIA',
      classification: 'FONTE',
      headers: Object.freeze([
        'ID_SEMESTRE_DIRETORIA',
        'ID_DIRETORIA',
        'ORDEM',
        'DATA_INICIO',
        'DATA_FIM',
        'TOTAL_DIAS',
        'STATUS',
        'PESO_LIMITE_DIRETORIA',
        'OBS'
      ])
    }),
    Object.freeze({
      sheetName: 'CARGOS_CONFIG',
      classification: 'FONTE',
      headers: Object.freeze([
        'CARGO_KEY',
        'CARGO_NOME',
        'NOME_PUBLICO',
        'TIPO_FUNCAO',
        'GRUPO_FUNCAO',
        'GRUPO_CARGO',
        'NIVEL_HIERARQUICO',
        'ESCRITA_VARIACAO',
        'EMAILS_GRUPO',
        'OBRIGATORIO_COMPOSICAO_INICIAL',
        'RECEBE_EMAILS',
        'E_CARGO_UNICO',
        'PERMITIR_NOMEACAO_VIA_FORM',
        'EXIGE_MEMBRO_EFETIVO',
        'EXIGE_VINCULO_ATIVO',
        'CONTA_LIMITE_DIRETORIA',
        'CONTA_PARA_LIMITE_DIRETORIA',
        'DIREITO_A_VOTO_DIRETORIA',
        'EXIGIR_HOMOLOGACAO_PREVIA',
        'GERA_PERFIL_PORTAL',
        'PERFIL_PORTAL_PADRAO',
        'PODE_VER_AREA_DIRETORIA',
        'PODE_GERENCIAR_ATIVIDADES',
        'PODE_REGISTRAR_CHAMADA',
        'PODE_EDITAR_ATIVIDADE',
        'PODE_ANALISAR_JUSTIFICATIVAS',
        'PODE_GERENCIAR_CERTIFICADOS',
        'PODE_GERENCIAR_COMUNICACAO',
        'ATIVO',
        'BASE_NORMATIVA',
        'OBS'
      ]),
      optionalHeaders: Object.freeze([
        'PODE_VER_AREA_DIRETORIA',
        'PODE_GERENCIAR_ATIVIDADES',
        'PODE_REGISTRAR_CHAMADA',
        'PODE_EDITAR_ATIVIDADE',
        'PODE_ANALISAR_JUSTIFICATIVAS',
        'PODE_GERENCIAR_CERTIFICADOS',
        'PODE_GERENCIAR_COMUNICACAO'
      ])
    }),
    Object.freeze({
      sheetName: 'VIGENCIAS_FUNCOES',
      classification: 'FONTE',
      headers: Object.freeze([
        'ID_VIGENCIA',
        'ID_PESSOA',
        'ID_VINCULO',
        'TIPO_FUNCAO',
        'CARGO_KEY',
        'CARGO_NOME_SNAPSHOT',
        'ID_DIRETORIA',
        'ID_SEMESTRE_DIRETORIA',
        'DATA_INICIO',
        'DATA_FIM_PREVISTA',
        'DATA_FIM_REAL',
        'STATUS_VIGENCIA',
        'FONTE_NOMEACAO',
        'LINK_ATA',
        'CRIADO_POR',
        'CRIADO_EM',
        'ATUALIZADO_POR',
        'ATUALIZADO_EM',
        'OBS',
        'ATIVO'
      ])
    }),
    Object.freeze({
      sheetName: 'VIGENCIAS_RESUMO_ATUAL',
      classification: 'CACHE',
      headers: Object.freeze([
        'ID_PESSOA',
        'NOME_EXIBICAO',
        'RGA',
        'CARGO_FUNCAO_ATUAL',
        'TIPO_FUNCAO_ATUAL',
        'GRUPO_FUNCAO_ATUAL',
        'ID_DIRETORIA_ATUAL',
        'PERFIS_PORTAL_CALCULADOS',
        'PERMISSOES_CALCULADAS',
        'DATA_INICIO_FUNCAO_ATUAL',
        'DATA_FIM_PREVISTA',
        'ULTIMA_ATUALIZACAO'
      ])
    })
  ])
});

function core_getDomainsV2Schemas_() {
  return CORE_DOMAINS_V2_SCHEMAS;
}

function core_isDomainsV2InitialBlankSheet_(sheet) {
  if (!sheet) return false;
  if (CORE_DOMAINS_V2_INITIAL_SHEET_NAMES.indexOf(sheet.getName()) < 0) return false;
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) return true;
  if (sheet.getLastRow() > 1 || sheet.getLastColumn() > 1) return false;
  return String(sheet.getRange(1, 1).getDisplayValue() || '').trim() === '';
}

function core_getOrCreateDomainsV2Sheet_(spreadsheet, expectedSheetName, sheetIndex) {
  var sheet = spreadsheet.getSheetByName(expectedSheetName);
  if (sheet) {
    return {
      sheet: sheet,
      created: false,
      renamedInitialSheet: false
    };
  }

  var sheets = spreadsheet.getSheets();
  if (sheetIndex === 0 && sheets.length === 1 && core_isDomainsV2InitialBlankSheet_(sheets[0])) {
    sheets[0].setName(expectedSheetName);
    return {
      sheet: sheets[0],
      created: false,
      renamedInitialSheet: true
    };
  }

  return {
    sheet: spreadsheet.insertSheet(expectedSheetName),
    created: true,
    renamedInitialSheet: false
  };
}

function core_dropDeprecatedDomainsV2Headers_(sheet, deprecatedHeaders) {
  deprecatedHeaders = deprecatedHeaders || [];
  if (!deprecatedHeaders.length || sheet.getLastColumn() < 1) {
    return {
      removedDeprecatedHeaders: [],
      skippedDeprecatedHeaders: []
    };
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
    .map(function(value) { return String(value || '').trim(); });
  var deprecatedMap = {};
  deprecatedHeaders.forEach(function(header) {
    deprecatedMap[core_normalizeHeader_(header)] = header;
  });

  var removed = [];
  var skipped = [];
  for (var c = headers.length - 1; c >= 0; c--) {
    var normalized = core_normalizeHeader_(headers[c]);
    if (!Object.prototype.hasOwnProperty.call(deprecatedMap, normalized)) continue;

    var hasData = false;
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var values = sheet.getRange(2, c + 1, lastRow - 1, 1).getDisplayValues();
      hasData = values.some(function(row) {
        return String(row[0] || '').trim() !== '';
      });
    }
    if (hasData) {
      skipped.push(headers[c]);
      continue;
    }
    sheet.deleteColumn(c + 1);
    removed.push(headers[c]);
  }

  return {
    removedDeprecatedHeaders: removed,
    skippedDeprecatedHeaders: skipped
  };
}

function core_ensureDomainsV2Headers_(sheet, expectedHeaders, deprecatedHeaders) {
  var deprecatedResult = core_dropDeprecatedDomainsV2Headers_(sheet, deprecatedHeaders);
  var lastColumn = sheet.getLastColumn();
  var existingHeaders = lastColumn > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(value) {
        return String(value || '').trim();
      })
    : [];
  var existingMap = core_buildHeaderIndexMap_(existingHeaders, {
    normalize: true,
    oneBased: false,
    keepFirst: true
  });
  var missing = [];

  expectedHeaders.forEach(function(header) {
    if (!Object.prototype.hasOwnProperty.call(existingMap, core_normalizeHeader_(header))) {
      missing.push(header);
    }
  });

  if (!existingHeaders.filter(function(header) { return !!header; }).length) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders.slice()]);
    return {
      wroteInitialHeaders: true,
      addedHeaders: expectedHeaders.slice(),
      missingHeaders: [],
      removedDeprecatedHeaders: deprecatedResult.removedDeprecatedHeaders,
      skippedDeprecatedHeaders: deprecatedResult.skippedDeprecatedHeaders
    };
  }

  if (missing.length) {
    sheet.getRange(1, existingHeaders.length + 1, 1, missing.length).setValues([missing]);
  }

  return {
    wroteInitialHeaders: false,
    addedHeaders: missing.slice(),
    missingHeaders: missing.slice(),
    removedDeprecatedHeaders: deprecatedResult.removedDeprecatedHeaders,
    skippedDeprecatedHeaders: deprecatedResult.skippedDeprecatedHeaders
  };
}

function core_applyDomainsV2SheetUx_(sheet) {
  var operations = [];

  function run(name, fn) {
    try {
      operations.push({
        name: name,
        ok: true,
        result: fn()
      });
    } catch (err) {
      operations.push({
        name: name,
        ok: false,
        error: err.message
      });
    }
  }

  run('freezeHeaderRow', function() {
    return core_freezeHeaderRow_(sheet, 1);
  });
  run('ensureFilter', function() {
    return core_ensureFilter_(sheet, 1, { recreate: false });
  });
  run('applyHeaderNotes', function() {
    return core_applyHeaderNotes_(sheet, CORE_DOMAINS_V2_HEADER_NOTES, 1);
  });

  return operations;
}

function core_ensureDomainsV2Spreadsheet_(domainKey, opts) {
  opts = opts || {};
  var schema = CORE_DOMAINS_V2_SCHEMAS[domainKey];
  if (!schema) {
    throw new Error('Dominio v2 nao reconhecido: ' + domainKey);
  }

  var environment = core_normalizeDomainEnv_({ ambiente: opts.ambiente || 'DEV' });
  var registryReport = core_validateDomainRegistry_(domainKey, {
    ambiente: environment,
    registryRaw: opts.registryRaw,
    openSpreadsheetById: opts.openSpreadsheetById
  });
  if (registryReport.duplicates.length || registryReport.divergences.length) {
    throw new Error('Setup V2 bloqueado por duplicidade ou divergencia no Registry de ' + domainKey + '/' + environment + '.');
  }
  var ref = core_getDomainSpreadsheetRef_(domainKey, {
    ambiente: environment,
    registryRaw: opts.registryRaw
  });
  var spreadsheet = core_openDomainSpreadsheet_(domainKey, {
    ambiente: environment,
    registryRaw: opts.registryRaw,
    openSpreadsheetById: opts.openSpreadsheetById
  });
  var out = {
    domain: domainKey,
    environment: environment,
    spreadsheetId: ref.spreadsheetIdMasked,
    spreadsheetName: spreadsheet.getName(),
    sheets: []
  };

  for (var i = 0; i < schema.length; i++) {
    var definition = schema[i];
    var resolved = core_getOrCreateDomainsV2Sheet_(spreadsheet, definition.sheetName, i);
    var headersResult = core_ensureDomainsV2Headers_(resolved.sheet, definition.headers, definition.deprecatedHeaders);
    var uxOperations = opts.applyUx === false ? [] : core_applyDomainsV2SheetUx_(resolved.sheet);

    out.sheets.push({
      sheetName: definition.sheetName,
      classification: definition.classification,
      created: resolved.created,
      renamedInitialSheet: resolved.renamedInitialSheet,
      wroteInitialHeaders: headersResult.wroteInitialHeaders,
      addedHeaders: headersResult.addedHeaders,
      removedDeprecatedHeaders: headersResult.removedDeprecatedHeaders,
      skippedDeprecatedHeaders: headersResult.skippedDeprecatedHeaders,
      uxOperations: uxOperations
    });
  }

  return out;
}

function coreEnsurePessoasV2DevSheets_(opts) {
  return core_ensureDomainsV2Spreadsheet_('PESSOAS', Object.assign({}, opts || {}, { ambiente: 'DEV' }));
}

/**
 * Prepara exclusivamente a aba geral de links e perfis de Pessoas V2.
 * A operacao e manual, exige confirmacao para escrita e nao modifica abas legadas.
 *
 * @param {Object=} options Opcoes de dry-run e confirmacao explicita.
 * @return {Object} Diagnostico ou resultado do provisionamento da aba.
 */
function corePessoasV2PrepareLinksPerfis_(options) {
  var opts = options || {};
  var dryRun = opts.dryRun !== false;
  var confirmacao = String(opts.confirmacao || '').trim();
  var definition = CORE_DOMAINS_V2_SCHEMAS.PESSOAS.filter(function(item) {
    return item.sheetName === 'PESSOAS_LINKS_PERFIS';
  })[0];
  var environment = core_normalizeDomainEnv_({ ambiente: opts.ambiente || 'DEV' });
  var validation = core_validateDomainRegistry_('PESSOAS', { ambiente: environment });
  if (validation.duplicates.length || validation.divergences.length) {
    throw new Error('Preparacao bloqueada por duplicidade ou divergencia no Registry de PESSOAS/' + environment + '.');
  }
  var spreadsheet = core_openDomainSpreadsheet_('PESSOAS', { ambiente: environment });
  var existing = spreadsheet.getSheetByName(definition.sheetName);
  var report = {
    ok: true,
    dryRun: dryRun,
    sheetName: definition.sheetName,
    spreadsheetName: spreadsheet.getName(),
    existedBefore: !!existing,
    requiredHeaders: definition.headers.slice(),
    created: false,
    addedHeaders: [],
    uxOperations: [],
    environment: environment,
    nextStep: 'Manter a key especifica apenas como fallback temporario de leitura; escritas usam PESSOAS_V2_DB.'
  };

  if (dryRun) {
    Logger.log('[geapa-core][PESSOAS_LINKS_PERFIS][PREPARE] ' + JSON.stringify(report));
    return report;
  }
  if (confirmacao !== 'PREPARAR_PESSOAS_V2_LINKS_PERFIS') {
    report.ok = false;
    report.blocked = true;
    report.reason = 'CONFIRMACAO_OBRIGATORIA';
    Logger.log('[geapa-core][PESSOAS_LINKS_PERFIS][PREPARE] ' + JSON.stringify(report));
    return report;
  }

  var sheetIndex = CORE_DOMAINS_V2_SCHEMAS.PESSOAS.indexOf(definition);
  var resolved = core_getOrCreateDomainsV2Sheet_(spreadsheet, definition.sheetName, sheetIndex);
  var headersResult = core_ensureDomainsV2Headers_(resolved.sheet, definition.headers, definition.deprecatedHeaders);
  report.created = !!resolved.created;
  report.addedHeaders = headersResult.addedHeaders;
  report.uxOperations = core_applyDomainsV2SheetUx_(resolved.sheet);
  Logger.log('[geapa-core][PESSOAS_LINKS_PERFIS][PREPARE] ' + JSON.stringify(report));
  return report;
}

/**
 * Executa o provisionamento real da aba de links sem exigir parametros no editor.
 *
 * @return {Object} Relatorio da criacao da aba e dos cabecalhos.
 */
function corePessoasV2PrepareLinksPerfisReal_() {
  return corePessoasV2PrepareLinksPerfis_({
    dryRun: false,
    confirmacao: 'PREPARAR_PESSOAS_V2_LINKS_PERFIS'
  });
}

function coreEnsureVigenciasV2DevSheets_(opts) {
  return core_ensureDomainsV2Spreadsheet_('VIGENCIAS', Object.assign({}, opts || {}, { ambiente: 'DEV' }));
}

function coreEnsureCentralDomainsV2DevSheets_(opts) {
  opts = opts || {};
  return {
    pessoas: coreEnsurePessoasV2DevSheets_(opts),
    vigencias: coreEnsureVigenciasV2DevSheets_(opts)
  };
}

function coreDiagnosticarCentralDomainsV2DevSheets_(opts) {
  opts = opts || {};
  var diagnostics = [];
  var domains = ['PESSOAS', 'VIGENCIAS'];

  domains.forEach(function(domainKey) {
    var spreadsheet = core_openDomainSpreadsheet_(domainKey, { ambiente: opts.ambiente || 'DEV' });
    CORE_DOMAINS_V2_SCHEMAS[domainKey].forEach(function(definition) {
      var sheet = spreadsheet.getSheetByName(definition.sheetName);
      var missingHeaders = [];
      if (sheet) {
        var headers = sheet.getLastColumn() > 0
          ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
          : [];
        var headerMap = core_buildHeaderIndexMap_(headers, {
          normalize: true,
          oneBased: false,
          keepFirst: true
        });
        definition.headers.forEach(function(header) {
          if (!Object.prototype.hasOwnProperty.call(headerMap, core_normalizeHeader_(header))) {
            missingHeaders.push(header);
          }
        });
      }

      diagnostics.push({
        domain: domainKey,
        environment: core_normalizeDomainEnv_({ ambiente: opts.ambiente || 'DEV' }),
        sheetName: definition.sheetName,
        classification: definition.classification,
        exists: !!sheet,
        missingHeaders: missingHeaders,
        ok: !!sheet ? missingHeaders.length === 0 : definition.optional === true,
        recommendation: !sheet
          ? (definition.optional === true
            ? 'Aba opcional durante rollout; preparar antes de usar links/perfis.'
            : 'Criar aba com cabecalhos oficiais.')
          : (missingHeaders.length ? 'Adicionar cabecalhos faltantes ao final.' : 'Estrutura aderente.')
      });
    });
  });

  return {
    ok: diagnostics.every(function(item) { return item.ok; }),
    diagnostics: diagnostics
  };
}
