/**
 * Parametros normativos operacionais.
 *
 * A planilha e fonte de configuracao, nunca fonte autonoma de norma. Cada
 * parametro precisa apontar para uma BASE_LEGAL vigente e compativel com o
 * modulo consumidor. Nao existe fallback de DEV para PROD ou de PROD para DEV.
 */

var CORE_NORMATIVE_REGISTRY_KEY = 'NORMAS_PARAMETROS_OPERACIONAIS';
var CORE_NORMATIVE_CACHE_TTL_SECONDS = 300;
var CORE_NORMATIVE_KNOWN_IDS = Object.freeze([
  'SUSPENSAO_MINIMA',
  'BLOQUEIO_DESLIGAMENTO_ANTES_APRESENTACAO',
  'DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA'
]);
var CORE_NORMATIVE_PARAMETER_DEFINITIONS = Object.freeze({
  SUSPENSAO_MINIMA: Object.freeze({ tipoValor: 'NUMERO', unidade: 'DIAS' }),
  BLOQUEIO_DESLIGAMENTO_ANTES_APRESENTACAO: Object.freeze({ tipoValor: 'NUMERO', unidade: 'DIAS' }),
  DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA: Object.freeze({
    tipoValor: 'BOOLEANO',
    unidade: 'NAO_APLICAVEL',
    baseLegalQueExigeValorTrue: 'NC01-2025-ART16-IV'
  })
});
var CORE_NORMATIVE_REQUIRED_HEADERS = Object.freeze([
  'PARAMETRO_ID', 'DESCRICAO', 'TIPO_VALOR', 'VALOR', 'UNIDADE', 'BASE_LEGAL',
  'MODULO_SISTEMA', 'VIGENTE'
]);
var __core_normative_execution_cache = {};

function core_normativeError_(code, message, details) {
  var error = new Error(message);
  error.code = code;
  error.errorCode = code;
  error.details = details || {};
  return error;
}

function core_normativeToken_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[\s-]+/g, '_');
}

function core_normativeEnvironment_(options) {
  options = options || {};
  var raw = options.ambiente || options.environment || options.env;
  if (!raw && options.requireExplicitEnvironment === false) raw = core_getCurrentEnv_();
  var environment = core_normativeToken_(raw);
  if (environment !== 'DEV' && environment !== 'PROD') {
    throw core_normativeError_(
      'PARAMETRO_NORMATIVO_AMBIENTE_INVALIDO',
      'Informe explicitamente DEV ou PROD para resolver parametros normativos.',
      { ambienteRecebido: environment || '(vazio)' }
    );
  }
  return environment;
}

function core_normativeRegistryRaw_(options) {
  return options && options.registryRaw ? options.registryRaw : core_getRegistryRaw_();
}

function core_normativeRegistryEntry_(environment, options) {
  var raw = core_normativeRegistryRaw_(options || {});
  var environmentMap = raw[CORE_NORMATIVE_REGISTRY_KEY] || null;
  var entry = environmentMap && environmentMap[environment];
  if (!entry) {
    throw core_normativeError_(
      'NORMAS_PARAMETROS_REGISTRY_' + environment + '_AUSENTE',
      'O parametro normativo esta indisponivel: a key ' + CORE_NORMATIVE_REGISTRY_KEY +
        ' nao possui entrada explicita em ' + environment + '.',
      {
        registryKey: CORE_NORMATIVE_REGISTRY_KEY,
        ambiente: environment,
        ambientesDisponiveis: environmentMap ? Object.keys(environmentMap).sort() : []
      }
    );
  }
  var activeLines = (entry.duplicateActiveLines || (entry.ativo ? [entry.lineNo] : []))
    .filter(function(line) { return line != null; });
  if (activeLines.length > 1) {
    throw core_normativeError_(
      'NORMAS_PARAMETROS_REGISTRY_DUPLICADO',
      'Existem duas linhas ativas da key normativa no mesmo ambiente.',
      { registryKey: CORE_NORMATIVE_REGISTRY_KEY, ambiente: environment, linhas: activeLines }
    );
  }
  if (entry.ativo !== true) {
    throw core_normativeError_(
      'NORMAS_PARAMETROS_REGISTRY_INATIVO',
      'O parametro normativo esta indisponivel: a key normativa esta inativa.',
      { registryKey: CORE_NORMATIVE_REGISTRY_KEY, ambiente: environment, linha: entry.lineNo }
    );
  }
  if (!String(entry.id || '').trim() || !String(entry.sheet || '').trim()) {
    throw core_normativeError_(
      'NORMAS_PARAMETROS_REGISTRY_REFERENCIA_INVALIDA',
      'O parametro normativo esta indisponivel: Registry sem planilha ou aba valida.',
      { registryKey: CORE_NORMATIVE_REGISTRY_KEY, ambiente: environment }
    );
  }
  return entry;
}

function core_normativeReadRecords_(environment, options) {
  options = options || {};
  if (Array.isArray(options.records)) return options.records.slice();
  if (options.recordsByEnvironment && Array.isArray(options.recordsByEnvironment[environment])) {
    return options.recordsByEnvironment[environment].slice();
  }
  var entry = core_normativeRegistryEntry_(environment, options);
  var sheet;
  if (typeof options.openSheet === 'function') {
    sheet = options.openSheet(entry.id, entry.sheet, environment);
  } else {
    sheet = core_getSheetById_(entry.id, entry.sheet);
  }
  if (!sheet) {
    throw core_normativeError_(
      'NORMAS_PARAMETROS_ABA_INDISPONIVEL',
      'O parametro normativo esta indisponivel: aba configurada nao encontrada.',
      { registryKey: CORE_NORMATIVE_REGISTRY_KEY, ambiente: environment, sheetName: entry.sheet }
    );
  }
  return core_readSheetRecords_(sheet, { skipBlankRows: true }) || [];
}

function core_normativeModuleAliases_(expectedModule) {
  var expected = core_normativeToken_(expectedModule || 'GEAPA_MEMBROS');
  if (expected === 'GEAPA_MEMBROS' || expected === 'MEMBROS') {
    return ['GEAPA_MEMBROS', 'MEMBROS', 'VINCULOS_GEAPA'];
  }
  return [expected];
}

function core_normativeModuleCompatible_(configured, expectedModule) {
  var configuredTokens = String(configured || '')
    .split(/[;,|]/)
    .map(core_normativeToken_)
    .filter(Boolean);
  var expected = core_normativeModuleAliases_(expectedModule);
  return configuredTokens.some(function(token) {
    return token === 'TODOS' || token === 'SISTEMA_GEAPA' || expected.indexOf(token) >= 0;
  });
}

function core_normativePositiveNumber_(value) {
  if (typeof value === 'string') value = value.trim().replace(',', '.');
  var number = Number(value);
  return isFinite(number) && number > 0 ? number : null;
}

function core_normativeValueType_(record, parametroId) {
  var configured = core_normativeToken_(record && record.TIPO_VALOR);
  var inferred = configured || 'NUMERO';
  if (inferred !== 'NUMERO' && inferred !== 'BOOLEANO') {
    throw core_normativeError_(
      'PARAMETRO_NORMATIVO_TIPO_VALOR_INVALIDO',
      'O parametro normativo esta indisponivel: TIPO_VALOR deve ser NUMERO ou BOOLEANO.',
      { parametroId: parametroId, tipoValor: inferred || '(vazio)' }
    );
  }
  var definition = CORE_NORMATIVE_PARAMETER_DEFINITIONS[parametroId] || null;
  if (definition && inferred !== definition.tipoValor) {
    throw core_normativeError_(
      'PARAMETRO_NORMATIVO_TIPO_VALOR_INCOMPATIVEL',
      'O tipo do parametro normativo e incompatível com a regra conhecida.',
      { parametroId: parametroId, esperado: definition.tipoValor, recebido: inferred }
    );
  }
  return inferred;
}

function core_normativeBoolean_(value, parametroId) {
  if (value === true || value === 1) return true;
  if (value === false || value === 0) return false;
  var token = core_normativeToken_(value);
  if (token === 'SIM' || token === 'TRUE' || token === '1') return true;
  if (token === 'NAO' || token === 'FALSE' || token === '0') return false;
  throw core_normativeError_(
    'PARAMETRO_NORMATIVO_VALOR_BOOLEANO_INVALIDO',
    'O parametro normativo booleano aceita somente SIM/NAO, TRUE/FALSE ou 1/0.',
    { parametroId: parametroId }
  );
}

function core_normativeTypedValue_(record, parametroId, tipoValor) {
  var definition = CORE_NORMATIVE_PARAMETER_DEFINITIONS[parametroId] || {};
  var unit = core_normativeToken_(record.UNIDADE);
  var value;
  if (tipoValor === 'BOOLEANO') {
    value = core_normativeBoolean_(record.VALOR, parametroId);
    if (unit && unit !== 'NAO_APLICAVEL') {
      throw core_normativeError_(
        'PARAMETRO_NORMATIVO_UNIDADE_INVALIDA',
        'O parametro normativo booleano deve usar UNIDADE vazia ou NAO_APLICAVEL.',
        { parametroId: parametroId, unidade: unit }
      );
    }
    unit = 'NAO_APLICAVEL';
  } else {
    value = core_normativePositiveNumber_(record.VALOR);
    if (value == null) {
      throw core_normativeError_(
        'PARAMETRO_NORMATIVO_VALOR_INVALIDO',
        'O parametro normativo esta indisponivel: VALOR deve ser numerico e positivo.',
        { parametroId: parametroId }
      );
    }
    if (definition.unidade && unit !== definition.unidade) {
      throw core_normativeError_(
        'PARAMETRO_NORMATIVO_UNIDADE_INVALIDA',
        'O parametro normativo deve usar UNIDADE = ' + definition.unidade + '.',
        { parametroId: parametroId, unidade: unit || '(vazio)', esperada: definition.unidade }
      );
    }
  }
  return { valor: value, unidade: unit };
}

function core_normativeValidateRecord_(record, parametroId, environment, options) {
  var tipoValor = core_normativeValueType_(record, parametroId);
  var typed = core_normativeTypedValue_(record, parametroId, tipoValor);
  var baseLegal = String(record.BASE_LEGAL || '').trim();
  if (!baseLegal) {
    throw core_normativeError_(
      'PARAMETRO_NORMATIVO_BASE_LEGAL_AUSENTE',
      'O parametro normativo esta indisponivel: BASE_LEGAL nao foi informada.',
      { parametroId: parametroId, ambiente: environment }
    );
  }
  var definition = CORE_NORMATIVE_PARAMETER_DEFINITIONS[parametroId] || {};
  if (definition.baseLegalQueExigeValorTrue && typed.valor === false && core_normativeToken_(baseLegal) === core_normativeToken_(definition.baseLegalQueExigeValorTrue)) {
    throw core_normativeError_(
      'PARAMETRO_NORMATIVO_BASE_LEGAL_INCOMPATIVEL',
      'A BASE_LEGAL informada ainda exige a regra ativa e nao autoriza VALOR = NAO.',
      { parametroId: parametroId, ambiente: environment, baseLegal: baseLegal }
    );
  }
  var moduleSystem = String(record.MODULO_SISTEMA || '').trim();
  if (!core_normativeModuleCompatible_(moduleSystem, options.moduloSistema || 'GEAPA_MEMBROS')) {
    throw core_normativeError_(
      'PARAMETRO_NORMATIVO_MODULO_INCOMPATIVEL',
      'O parametro normativo esta indisponivel para o modulo solicitante.',
      { parametroId: parametroId, ambiente: environment, moduloSistema: moduleSystem || '(vazio)' }
    );
  }
  return Object.freeze({
    parametroId: parametroId,
    tipoValor: tipoValor,
    valor: typed.valor,
    unidade: typed.unidade,
    baseLegal: baseLegal,
    moduloSistema: moduleSystem,
    vigente: true,
    ambiente: environment,
    registryKey: CORE_NORMATIVE_REGISTRY_KEY,
    origem: 'REGISTRY_EXPLICITO',
    atualizadoEm: record.ATUALIZADO_EM || record.VIGENTE_DESDE || ''
  });
}

function core_prepararParametrosNormativosTipados_(options) {
  options = options || {};
  var environment = core_normativeEnvironment_(options);
  var headers = (options.headers || []).map(core_normativeToken_).filter(Boolean);
  var records = options.records || [];
  var missingHeaders = CORE_NORMATIVE_REQUIRED_HEADERS.filter(function(header) {
    return headers.indexOf(header) < 0;
  });
  var parameterId = 'DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA';
  var existing = records.filter(function(record) {
    return core_normativeToken_(record.PARAMETRO_ID) === parameterId;
  });
  return Object.freeze({
    ok: missingHeaders.length === 0 && existing.length === 1,
    dryRun: true,
    somentePlanejamento: true,
    ambiente: environment,
    registryKey: CORE_NORMATIVE_REGISTRY_KEY,
    cabecalhosObrigatorios: CORE_NORMATIVE_REQUIRED_HEADERS.slice(),
    cabecalhosAusentes: missingHeaders,
    parametroExistente: existing.length === 1,
    duplicidadeParametro: existing.length > 1,
    linhaRecomendada: Object.freeze({
      PARAMETRO_ID: parameterId,
      DESCRICAO: 'Exige ata na decisao final de desligamento voluntario',
      TIPO_VALOR: 'BOOLEANO',
      VALOR: 'SIM',
      UNIDADE: 'NAO_APLICAVEL',
      BASE_LEGAL: 'NC01-2025-ART16-IV',
      MODULO_SISTEMA: 'GEAPA_MEMBROS',
      VIGENTE: 'SIM'
    }),
    escritaExecutada: false
  });
}

function core_normativeResolveFromRecords_(records, parametroId, environment, options) {
  var wanted = core_normativeToken_(parametroId);
  if (!wanted) {
    throw core_normativeError_('PARAMETRO_NORMATIVO_ID_OBRIGATORIO', 'PARAMETRO_ID e obrigatorio.', {});
  }
  var matches = (records || []).filter(function(record) {
    return core_normativeToken_(record.PARAMETRO_ID) === wanted;
  });
  if (!matches.length) {
    throw core_normativeError_(
      'PARAMETRO_NORMATIVO_NAO_ENCONTRADO',
      'O parametro normativo esta indisponivel: PARAMETRO_ID nao encontrado.',
      { parametroId: wanted, ambiente: environment }
    );
  }
  var current = matches.filter(function(record) {
    return core_normativeToken_(record.VIGENTE) === 'SIM';
  });
  if (!current.length) {
    throw core_normativeError_(
      'PARAMETRO_NORMATIVO_NAO_VIGENTE',
      'O parametro normativo esta indisponivel: nao existe linha VIGENTE = SIM.',
      { parametroId: wanted, ambiente: environment }
    );
  }
  if (current.length > 1) {
    throw core_normativeError_(
      'PARAMETRO_NORMATIVO_VIGENTE_DUPLICADO',
      'O parametro normativo esta indisponivel: mais de uma linha esta vigente.',
      { parametroId: wanted, ambiente: environment, quantidade: current.length }
    );
  }
  return core_normativeValidateRecord_(current[0], wanted, environment, options || {});
}

function core_normativeCacheKey_(environment, parametroId, moduleSystem) {
  return ['CORE_NORMATIVE', environment, core_normativeToken_(moduleSystem || 'GEAPA_MEMBROS'), core_normativeToken_(parametroId)].join(':');
}

function core_normativeCacheGet_(key, options) {
  if (options.records || options.recordsByEnvironment || options.registryRaw || options.disableCache === true) return null;
  if (__core_normative_execution_cache[key]) return __core_normative_execution_cache[key];
  if (typeof CacheService === 'undefined') return null;
  var serialized = CacheService.getScriptCache().get(key);
  if (!serialized) return null;
  try {
    var parsed = Object.freeze(JSON.parse(serialized));
    __core_normative_execution_cache[key] = parsed;
    return parsed;
  } catch (ignored) {
    return null;
  }
}

function core_normativeCachePut_(key, value, options) {
  if (options.records || options.recordsByEnvironment || options.registryRaw || options.disableCache === true) return;
  __core_normative_execution_cache[key] = value;
  if (typeof CacheService !== 'undefined') {
    CacheService.getScriptCache().put(key, JSON.stringify(value), CORE_NORMATIVE_CACHE_TTL_SECONDS);
  }
}

function core_resolverParametroNormativoOperacional_(parametroId, options) {
  options = options || {};
  var environment = core_normativeEnvironment_(options);
  var id = core_normativeToken_(parametroId);
  var cacheKey = core_normativeCacheKey_(environment, id, options.moduloSistema);
  var cached = core_normativeCacheGet_(cacheKey, options);
  if (cached) return cached;
  core_normativeRegistryEntry_(environment, options);
  var records = core_normativeReadRecords_(environment, options);
  var resolved = core_normativeResolveFromRecords_(records, id, environment, options);
  core_normativeCachePut_(cacheKey, resolved, options);
  return resolved;
}

function core_resolverParametrosNormativosOperacionais_(parametroIds, options) {
  options = options || {};
  var environment = core_normativeEnvironment_(options);
  var ids = (parametroIds || []).map(core_normativeToken_).filter(Boolean);
  if (!ids.length) {
    throw core_normativeError_('PARAMETROS_NORMATIVOS_IDS_OBRIGATORIOS', 'Informe ao menos um PARAMETRO_ID.', {});
  }
  core_normativeRegistryEntry_(environment, options);
  var records = core_normativeReadRecords_(environment, options);
  var output = {};
  ids.forEach(function(id) {
    output[id] = core_normativeResolveFromRecords_(records, id, environment, options);
    core_normativeCachePut_(core_normativeCacheKey_(environment, id, options.moduloSistema), output[id], options);
  });
  return Object.freeze(output);
}

function core_invalidarCacheParametrosNormativosOperacionais_(options) {
  options = options || {};
  var environments = options.ambiente || options.environment
    ? [core_normativeEnvironment_(options)]
    : ['DEV', 'PROD'];
  var ids = (options.parametroIds || CORE_NORMATIVE_KNOWN_IDS).map(core_normativeToken_);
  var moduleSystem = options.moduloSistema || 'GEAPA_MEMBROS';
  var removed = [];
  environments.forEach(function(environment) {
    ids.forEach(function(id) {
      var key = core_normativeCacheKey_(environment, id, moduleSystem);
      delete __core_normative_execution_cache[key];
      if (typeof CacheService !== 'undefined') CacheService.getScriptCache().remove(key);
      removed.push(key);
    });
  });
  return Object.freeze({ ok: true, somenteCache: true, chavesRemovidas: removed.length });
}

function core_diagnosticarParametrosNormativosOperacionais_(options) {
  options = options || {};
  var environment = core_normativeEnvironment_(options);
  var ids = options.parametroIds || CORE_NORMATIVE_KNOWN_IDS;
  var report = { ok: true, ambiente: environment, registryKey: CORE_NORMATIVE_REGISTRY_KEY, parametros: [], erros: [] };
  ids.forEach(function(id) {
    try {
      var value = core_resolverParametroNormativoOperacional_(id, Object.assign({}, options, { disableCache: true }));
      report.parametros.push(value);
    } catch (error) {
      report.ok = false;
      report.erros.push({ parametroId: core_normativeToken_(id), code: error.code || 'PARAMETRO_NORMATIVO_ERRO', message: error.message });
    }
  });
  return Object.freeze(report);
}
