/**
 * Resolucao canonica e estrita das planilhas centrais V2 por dominio.
 *
 * Regras principais:
 * - ambiente explicito tem precedencia sobre GEAPA_ENV;
 * - nunca existe fallback entre DEV e PROD;
 * - leitura pode usar key especifica apenas como compatibilidade controlada;
 * - escrita usa exclusivamente a key *_V2_DB e a aba canonica;
 * - IDs nunca sao registrados sem mascara.
 */

var CORE_DOMAIN_V2_MAP = Object.freeze({
  PESSOAS: Object.freeze({
    registryKey: 'PESSOAS_V2_DB',
    anchor: 'BASE',
    sheets: Object.freeze({
      BASE: 'PESSOAS_BASE',
      IDENTIFICADORES: 'PESSOAS_IDENTIFICADORES',
      MEMBROS_DETALHES: 'MEMBROS_DETALHES',
      SOLICITACOES_ATUALIZACAO: 'SOLICITACOES_ATUALIZACAO_CADASTRAL',
      LINKS_PERFIS: 'PESSOAS_LINKS_PERFIS',
      COLABORADORES: 'COLABORADORES_ACADEMICOS',
      EXTERNOS: 'PARTICIPANTES_EXTERNOS_DETALHES',
      VINCULOS: 'VINCULOS_GEAPA',
      EVENTOS: 'MEMBROS_EVENTOS_VINCULO',
      CONSENTIMENTOS: 'PESSOAS_COMUNICACAO_CONSENTIMENTOS',
      ACESSOS_EXCECOES: 'PORTAL_ACESSOS_EXCECOES',
      RESUMO: 'PESSOAS_RESUMO_OPERACIONAL'
    }),
    specificRegistryKeys: Object.freeze({
      BASE: 'PESSOAS_V2_BASE',
      IDENTIFICADORES: 'PESSOAS_V2_IDENTIFICADORES',
      MEMBROS_DETALHES: 'PESSOAS_V2_MEMBROS_DETALHES',
      SOLICITACOES_ATUALIZACAO: 'PESSOAS_V2_SOLICITACOES_ATUALIZACAO_CADASTRAL',
      LINKS_PERFIS: 'PESSOAS_V2_LINKS_PERFIS',
      COLABORADORES: 'PESSOAS_V2_COLABORADORES_ACADEMICOS',
      EXTERNOS: 'PESSOAS_V2_PARTICIPANTES_EXTERNOS_DETALHES',
      VINCULOS: 'PESSOAS_V2_VINCULOS_GEAPA',
      EVENTOS: 'PESSOAS_V2_MEMBROS_EVENTOS_VINCULO',
      CONSENTIMENTOS: 'PESSOAS_V2_COMUNICACAO_CONSENTIMENTOS',
      ACESSOS_EXCECOES: 'PESSOAS_V2_PORTAL_ACESSOS_EXCECOES',
      RESUMO: 'PESSOAS_V2_RESUMO_OPERACIONAL'
    })
  }),
  VIGENCIAS: Object.freeze({
    registryKey: 'VIGENCIAS_V2_DB',
    anchor: 'SEMESTRES',
    sheets: Object.freeze({
      SEMESTRES: 'SEMESTRES',
      CICLOS: 'CICLOS',
      DIRETORIAS: 'DIRETORIAS',
      SEMESTRES_DIRETORIA: 'SEMESTRES_DIRETORIA',
      CARGOS_CONFIG: 'CARGOS_CONFIG',
      FUNCOES: 'VIGENCIAS_FUNCOES',
      RESUMO: 'VIGENCIAS_RESUMO_ATUAL'
    }),
    specificRegistryKeys: Object.freeze({
      SEMESTRES: 'VIGENCIAS_V2_SEMESTRES',
      CICLOS: 'VIGENCIAS_V2_CICLOS',
      DIRETORIAS: 'VIGENCIAS_V2_DIRETORIAS',
      SEMESTRES_DIRETORIA: 'VIGENCIAS_V2_SEMESTRES_DIRETORIA',
      CARGOS_CONFIG: 'VIGENCIAS_V2_CARGOS_CONFIG',
      FUNCOES: 'VIGENCIAS_V2_FUNCOES',
      RESUMO: 'VIGENCIAS_V2_RESUMO_ATUAL'
    })
  }),
  ATIVIDADES: Object.freeze({
    registryKey: 'ATIVIDADES_V2_DB',
    anchor: 'ATIVIDADES',
    sheets: Object.freeze({
      ATIVIDADES: 'Atividades',
      APRESENTACOES: 'Atividades_Apresentacoes',
      ARQUIVOS: 'Atividades_Arquivos',
      ENVOLVIDOS: 'Atividades_Envolvidos',
      PRESENCAS_REGISTROS: 'Atividades_Presencas_Registros',
      CONVITES: 'Atividades_Convites',
      JUSTIFICATIVAS: 'Justificativas_Faltas',
      CONFIG: 'Atividades_Config',
      LOG: 'Atividades_Log',
      PORTAL_ACOES: 'Portal_Acoes',
      PORTAL_ATIVIDADES_CALENDARIO: 'PORTAL_ATIVIDADES_CALENDARIO',
      PORTAL_ATIVIDADES_DETALHES: 'PORTAL_ATIVIDADES_DETALHES',
      PORTAL_FREQUENCIA_MEMBROS: 'PORTAL_FREQUENCIA_MEMBROS',
      PORTAL_JUSTIFICATIVAS: 'PORTAL_JUSTIFICATIVAS',
      PORTAL_PENDENCIAS_DIRETORIA: 'PORTAL_PENDENCIAS_DIRETORIA',
      PORTAL_STATUS_ATIVIDADES: 'PORTAL_STATUS_ATIVIDADES'
    }),
    specificRegistryKeys: Object.freeze({
      ATIVIDADES: 'ATIVIDADES_V2_ATIVIDADES',
      APRESENTACOES: 'ATIVIDADES_V2_APRESENTACOES',
      ARQUIVOS: 'ATIVIDADES_V2_ARQUIVOS',
      ENVOLVIDOS: 'ATIVIDADES_V2_ENVOLVIDOS',
      PRESENCAS_REGISTROS: 'ATIVIDADES_V2_PRESENCAS_REGISTROS',
      CONVITES: 'ATIVIDADES_V2_CONVITES',
      JUSTIFICATIVAS: 'ATIVIDADES_V2_JUSTIFICATIVAS',
      CONFIG: 'ATIVIDADES_V2_CONFIG',
      LOG: 'ATIVIDADES_V2_LOG',
      PORTAL_ACOES: 'ATIVIDADES_V2_PORTAL_ACOES',
      PORTAL_ATIVIDADES_CALENDARIO: 'ATIVIDADES_V2_PORTAL_ATIVIDADES_CALENDARIO',
      PORTAL_ATIVIDADES_DETALHES: 'ATIVIDADES_V2_PORTAL_ATIVIDADES_DETALHES',
      PORTAL_FREQUENCIA_MEMBROS: 'ATIVIDADES_V2_PORTAL_FREQUENCIA_MEMBROS',
      PORTAL_JUSTIFICATIVAS: 'ATIVIDADES_V2_PORTAL_JUSTIFICATIVAS',
      PORTAL_PENDENCIAS_DIRETORIA: 'ATIVIDADES_V2_PORTAL_PENDENCIAS_DIRETORIA',
      PORTAL_STATUS_ATIVIDADES: 'ATIVIDADES_V2_PORTAL_STATUS_ATIVIDADES'
    })
  })
});

var __core_domain_ss_cache = {};

function core_domainResolverError_(code, message, details) {
  var err = new Error(message);
  err.code = code;
  err.details = details || {};
  return err;
}

function core_normalizeDomainEnv_(options) {
  options = options || {};
  var raw = options.ambiente || options.environment || options.env;
  if (!raw) raw = core_getCurrentEnv_();
  var env = String(raw || '').trim().toUpperCase();
  if (env !== 'DEV' && env !== 'PROD') {
    throw core_domainResolverError_('DOMAIN_ENV_INVALIDO', 'Ambiente invalido para dominio V2: "' + env + '". Use DEV ou PROD.', { ambiente: env });
  }
  return env;
}

function core_getDomainDefinition_(domain) {
  var key = String(domain || '').trim().toUpperCase();
  var definition = CORE_DOMAIN_V2_MAP[key];
  if (!definition) {
    throw core_domainResolverError_('DOMAIN_INVALIDO', 'Dominio V2 nao reconhecido: "' + key + '".', { dominio: key });
  }
  return { key: key, definition: definition };
}

function core_maskSpreadsheetId_(id) {
  var value = String(id || '').trim();
  if (!value) return '';
  if (value.length <= 10) return value.substring(0, 2) + '...' + value.substring(value.length - 2);
  return value.substring(0, 5) + '...' + value.substring(value.length - 4);
}

function core_domainRegistryRaw_(options) {
  return options && options.registryRaw ? options.registryRaw : core_getRegistryRaw_();
}

function core_domainRegistryEntry_(registryKey, environment, options) {
  var key = String(registryKey || '').trim().toUpperCase();
  var raw = core_domainRegistryRaw_(options || {});
  var envMap = raw[key];
  if (!envMap || !envMap[environment]) {
    return { key: key, environment: environment, entry: null, available: false, reason: 'AUSENTE' };
  }
  var entry = envMap[environment];
  var activeLines = (entry.duplicateActiveLines || (entry.ativo ? [entry.lineNo] : [])).filter(function(line) { return line != null; });
  if (activeLines.length > 1) {
    throw core_domainResolverError_('DOMAIN_REGISTRY_DUPLICADO', 'Registry possui mais de uma linha ativa para "' + key + '" em ' + environment + '.', {
      registryKey: key,
      ambiente: environment,
      linhas: activeLines.slice()
    });
  }
  if (entry.ativo !== true) {
    return { key: key, environment: environment, entry: entry, available: false, reason: 'INATIVA' };
  }
  if (!String(entry.id || '').trim()) {
    throw core_domainResolverError_('DOMAIN_REGISTRY_ID_AUSENTE', 'Registry key "' + key + '" sem SPREADSHEET_ID em ' + environment + '.', {
      registryKey: key,
      ambiente: environment
    });
  }
  return { key: key, environment: environment, entry: entry, available: true, reason: '' };
}

function core_domainWarn_(code, message, details, options) {
  var item = { level: 'WARN', code: code, message: message, details: details || {} };
  if (options && Array.isArray(options.warnings)) options.warnings.push(item);
  if (typeof core_logWarn_ === 'function') {
    core_logWarn_(typeof core_runId_ === 'function' ? core_runId_() : 'DOMAIN_V2', message, Object.assign({ code: code }, details || {}));
  } else if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
    Logger.log('[GEAPA_CORE][DOMAIN_V2][WARN] ' + JSON.stringify(item));
  }
  return item;
}

function core_getDomainSpreadsheetRef_(domain, options) {
  options = options || {};
  var resolvedDomain = core_getDomainDefinition_(domain);
  var environment = core_normalizeDomainEnv_(options);
  var state = core_domainRegistryEntry_(resolvedDomain.definition.registryKey, environment, options);
  if (!state.available) {
    throw core_domainResolverError_('DOMAIN_DB_INDISPONIVEL', 'Registry key "' + resolvedDomain.definition.registryKey + '" indisponivel em ' + environment + '.', {
      dominio: resolvedDomain.key,
      registryKey: resolvedDomain.definition.registryKey,
      ambiente: environment,
      motivo: state.reason
    });
  }
  if (options.forWrite === true || String(options.access || '').toUpperCase() === 'WRITE') {
    Object.keys(resolvedDomain.definition.specificRegistryKeys).forEach(function(logical) {
      var specificKey = resolvedDomain.definition.specificRegistryKeys[logical];
      var specificState = core_domainRegistryEntry_(specificKey, environment, options);
      if (!specificState.available || String(specificState.entry.id).trim() === String(state.entry.id).trim()) return;
      var divergence = {
        dominio: resolvedDomain.key,
        abaLogica: logical,
        ambiente: environment,
        dbKey: resolvedDomain.definition.registryKey,
        specificKey: specificKey,
        dbSpreadsheetId: core_maskSpreadsheetId_(state.entry.id),
        specificSpreadsheetId: core_maskSpreadsheetId_(specificState.entry.id)
      };
      core_domainWarn_('DOMAIN_REGISTRY_DIVERGENCIA', 'DB key e key especifica apontam para planilhas diferentes.', divergence, options);
      throw core_domainResolverError_('DOMAIN_WRITE_REGISTRY_DIVERGENTE', 'Escrita bloqueada: DB key e key especifica divergem em ' + environment + '.', divergence);
    });
  }
  return Object.freeze({
    domain: resolvedDomain.key,
    environment: environment,
    origin: 'DOMAIN_DB',
    registryKey: resolvedDomain.definition.registryKey,
    spreadsheetId: String(state.entry.id).trim(),
    spreadsheetIdMasked: core_maskSpreadsheetId_(state.entry.id),
    anchorSheetName: resolvedDomain.definition.sheets[resolvedDomain.definition.anchor],
    lineNo: state.entry.lineNo
  });
}

function core_openDomainSpreadsheet_(domain, options) {
  options = options || {};
  var ref = core_getDomainSpreadsheetRef_(domain, options);
  var cacheKey = ref.environment + '|' + ref.domain + '|' + ref.spreadsheetId;
  if (!__core_domain_ss_cache[cacheKey]) {
    var opener = options.openSpreadsheetById || core_openSpreadsheetById_;
    __core_domain_ss_cache[cacheKey] = opener(ref.spreadsheetId);
  }
  return __core_domain_ss_cache[cacheKey];
}

function core_resetDomainResolverExecutionCache_() {
  __core_domain_ss_cache = {};
}

function core_bindDomainWriteContext_(context, resolution) {
  if (!context) return;
  if (!context.spreadsheetId) {
    context.domain = resolution.domain;
    context.environment = resolution.environment;
    context.spreadsheetId = resolution.spreadsheetId;
    return;
  }
  if (context.domain !== resolution.domain || context.environment !== resolution.environment || context.spreadsheetId !== resolution.spreadsheetId) {
    throw core_domainResolverError_('DOMAIN_WRITE_CONTEXT_DIVERGENTE', 'O fluxo de escrita tentou alternar dominio, ambiente ou planilha.', {
      esperado: { dominio: context.domain, ambiente: context.environment, spreadsheetId: core_maskSpreadsheetId_(context.spreadsheetId) },
      recebido: { dominio: resolution.domain, ambiente: resolution.environment, spreadsheetId: resolution.spreadsheetIdMasked }
    });
  }
}

function core_resolveDomainSheet_(domain, logicalSheet, options) {
  options = options || {};
  var resolvedDomain = core_getDomainDefinition_(domain);
  var logical = String(logicalSheet || '').trim().toUpperCase();
  var canonicalName = resolvedDomain.definition.sheets[logical];
  if (!canonicalName) {
    throw core_domainResolverError_('DOMAIN_SHEET_INVALIDA', 'Aba logica "' + logical + '" nao pertence ao dominio ' + resolvedDomain.key + '.', {
      dominio: resolvedDomain.key,
      abaLogica: logical
    });
  }
  var environment = core_normalizeDomainEnv_(options);
  var writeMode = options.forWrite === true || String(options.access || '').toUpperCase() === 'WRITE';
  var dbState = core_domainRegistryEntry_(resolvedDomain.definition.registryKey, environment, options);
  var specificKey = resolvedDomain.definition.specificRegistryKeys[logical] || '';
  var specificState = specificKey ? core_domainRegistryEntry_(specificKey, environment, options) : null;

  if (dbState.available && specificState && specificState.available && String(dbState.entry.id).trim() !== String(specificState.entry.id).trim()) {
    var divergence = {
      dominio: resolvedDomain.key,
      abaLogica: logical,
      ambiente: environment,
      dbKey: resolvedDomain.definition.registryKey,
      specificKey: specificKey,
      dbSpreadsheetId: core_maskSpreadsheetId_(dbState.entry.id),
      specificSpreadsheetId: core_maskSpreadsheetId_(specificState.entry.id)
    };
    core_domainWarn_('DOMAIN_REGISTRY_DIVERGENCIA', 'DB key e key especifica apontam para planilhas diferentes.', divergence, options);
    if (writeMode) {
      throw core_domainResolverError_('DOMAIN_WRITE_REGISTRY_DIVERGENTE', 'Escrita bloqueada: DB key e key especifica divergem em ' + environment + '.', divergence);
    }
  }

  if (dbState.available) {
    var dbRef = {
      domain: resolvedDomain.key,
      environment: environment,
      origin: 'DOMAIN_DB',
      registryKey: resolvedDomain.definition.registryKey,
      spreadsheetId: String(dbState.entry.id).trim(),
      spreadsheetIdMasked: core_maskSpreadsheetId_(dbState.entry.id),
      logicalSheet: logical,
      sheetName: canonicalName
    };
    var spreadsheet = core_openDomainSpreadsheet_(resolvedDomain.key, Object.assign({}, options, { ambiente: environment }));
    var canonicalSheet = spreadsheet.getSheetByName(canonicalName);
    if (canonicalSheet) {
      if (writeMode) core_bindDomainWriteContext_(options.writeContext, dbRef);
      return { sheet: canonicalSheet, resolution: dbRef };
    }
    if (writeMode) {
      throw core_domainResolverError_('DOMAIN_WRITE_ABA_CANONICA_AUSENTE', 'Escrita bloqueada: aba canonica "' + canonicalName + '" ausente na DB key.', dbRef);
    }
  } else if (writeMode) {
    throw core_domainResolverError_('DOMAIN_WRITE_DB_INDISPONIVEL', 'Escrita bloqueada: DB key "' + resolvedDomain.definition.registryKey + '" indisponivel em ' + environment + '.', {
      dominio: resolvedDomain.key,
      abaLogica: logical,
      ambiente: environment,
      motivo: dbState.reason
    });
  }

  if (!writeMode && specificState && specificState.available) {
    var fallbackId = String(specificState.entry.id).trim();
    var fallbackCacheKey = environment + '|LEGACY|' + specificKey + '|' + fallbackId;
    if (!__core_domain_ss_cache[fallbackCacheKey]) {
      var fallbackOpener = options.openSpreadsheetById || core_openSpreadsheetById_;
      __core_domain_ss_cache[fallbackCacheKey] = fallbackOpener(fallbackId);
    }
    var fallbackSheetName = String(specificState.entry.sheet || canonicalName).trim();
    var fallbackSheet = __core_domain_ss_cache[fallbackCacheKey].getSheetByName(fallbackSheetName);
    if (fallbackSheet) {
      var fallbackResolution = {
        domain: resolvedDomain.key,
        environment: environment,
        origin: 'SPECIFIC_KEY_FALLBACK',
        registryKey: specificKey,
        spreadsheetId: fallbackId,
        spreadsheetIdMasked: core_maskSpreadsheetId_(fallbackId),
        logicalSheet: logical,
        sheetName: fallbackSheetName
      };
      core_domainWarn_('DOMAIN_SPECIFIC_KEY_FALLBACK', 'Leitura usando fallback temporario por key especifica.', {
        dominio: resolvedDomain.key,
        abaLogica: logical,
        ambiente: environment,
        origem: fallbackResolution.origin,
        registryKey: specificKey,
        spreadsheetId: fallbackResolution.spreadsheetIdMasked
      }, options);
      return { sheet: fallbackSheet, resolution: fallbackResolution };
    }
  }

  throw core_domainResolverError_('DOMAIN_SHEET_INDISPONIVEL', 'Aba "' + canonicalName + '" indisponivel para leitura em ' + environment + '.', {
    dominio: resolvedDomain.key,
    abaLogica: logical,
    ambiente: environment,
    dbKey: resolvedDomain.definition.registryKey,
    specificKey: specificKey || null
  });
}

function core_getDomainSheet_(domain, logicalSheet, options) {
  var resolved = core_resolveDomainSheet_(domain, logicalSheet, options || {});
  return options && options.includeResolution === true ? resolved : resolved.sheet;
}

function core_validateDomainRegistry_(domain, options) {
  options = options || {};
  var resolvedDomain = core_getDomainDefinition_(domain);
  var environment = core_normalizeDomainEnv_(options);
  var report = {
    ok: true,
    domain: resolvedDomain.key,
    environment: environment,
    registryKey: resolvedDomain.definition.registryKey,
    spreadsheetIdMasked: '',
    missingSheets: [],
    divergences: [],
    duplicates: [],
    legacySpecificKeys: [],
    legacySpecificKeysReferenced: Object.keys(resolvedDomain.definition.specificRegistryKeys).map(function(logical) {
      return resolvedDomain.definition.specificRegistryKeys[logical];
    })
  };
  var dbState;
  try {
    dbState = core_domainRegistryEntry_(resolvedDomain.definition.registryKey, environment, options);
  } catch (err) {
    report.ok = false;
    report.duplicates.push({ registryKey: resolvedDomain.definition.registryKey, code: err.code, details: err.details || {} });
    return report;
  }
  var spreadsheet = null;
  if (dbState.available) {
    report.spreadsheetIdMasked = core_maskSpreadsheetId_(dbState.entry.id);
    try {
      spreadsheet = core_openDomainSpreadsheet_(resolvedDomain.key, Object.assign({}, options, { ambiente: environment }));
    } catch (err) {
      report.ok = false;
      report.dbOpenError = {
        code: err && err.code ? err.code : 'DOMAIN_DB_OPEN_ERROR',
        message: err && err.message ? String(err.message) : 'Falha ao abrir planilha do dominio.'
      };
    }
  } else {
    report.ok = false;
    report.dbUnavailableReason = dbState.reason;
  }

  Object.keys(resolvedDomain.definition.sheets).forEach(function(logical) {
    var canonicalName = resolvedDomain.definition.sheets[logical];
    if (spreadsheet && !spreadsheet.getSheetByName(canonicalName)) {
      report.missingSheets.push({ logicalSheet: logical, sheetName: canonicalName });
    }
    var specificKey = resolvedDomain.definition.specificRegistryKeys[logical];
    if (!specificKey) return;
    var specificState;
    try {
      specificState = core_domainRegistryEntry_(specificKey, environment, options);
    } catch (err) {
      report.ok = false;
      report.duplicates.push({ registryKey: specificKey, code: err.code, details: err.details || {} });
      return;
    }
    if (!specificState.available) return;
    report.legacySpecificKeys.push(specificKey);
    if (dbState.available && String(dbState.entry.id).trim() !== String(specificState.entry.id).trim()) {
      report.divergences.push({
        logicalSheet: logical,
        dbKey: resolvedDomain.definition.registryKey,
        specificKey: specificKey,
        dbSpreadsheetId: core_maskSpreadsheetId_(dbState.entry.id),
        specificSpreadsheetId: core_maskSpreadsheetId_(specificState.entry.id)
      });
    }
  });
  if (report.missingSheets.length || report.divergences.length || report.duplicates.length) report.ok = false;
  return report;
}

function core_validateAllDomainRegistries_(options) {
  return Object.keys(CORE_DOMAIN_V2_MAP).map(function(domain) {
    return core_validateDomainRegistry_(domain, options || {});
  });
}
