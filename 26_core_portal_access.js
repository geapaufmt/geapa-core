/***************************************
 * 26_core_portal_access.js
 *
 * Camada inicial de autorizacao/permissoes do Portal GEAPA.
 *
 * Escopo desta etapa:
 * - ler perfis, permissoes, configuracoes e pessoas via Registry;
 * - autorizar temporariamente por e-mail enquanto Firebase Auth nao existe;
 * - registrar logs seguros de acesso/autorizacao;
 * - nao validar ID Token, nao usar Firebase e nao alterar o portal visual.
 ***************************************/

const CORE_PORTAL_ACCESS_CFG = Object.freeze({
  requiredRegistryKeys: Object.freeze([
    'PORTAL_PERFIS',
    'PORTAL_PERMISSOES',
    'PORTAL_CONFIG',
    'PORTAL_LOG_ACESSOS'
  ]),
  peopleSpreadsheetId: '1j3ea96-ySjz7qrn4PJ4Ds3wlQc19cfpJrIiAA72560Q',
  portalConfigCacheKey: 'GEAPA_CORE_PORTAL_CONFIG_V1',
  portalConfigCacheTtlSeconds: 10 * 60,
  sources: Object.freeze({
    current: Object.freeze({
      type: 'MEMBROS_ATUAIS',
      registryKeys: Object.freeze(['MEMBERS_ATUAIS', 'PESSOAS_MEMBROS_ATUAIS']),
      sheetNames: Object.freeze(['Membros Atuais', 'MEMBERS_ATUAIS'])
    }),
    waiting: Object.freeze({
      type: 'MEMBROS_EM_ESPERA',
      registryKeys: Object.freeze(['PESSOAS_MEMBROS_ESPERA', 'MEMBROS_EM_ESPERA', 'MEMBROS_ESPERA']),
      sheetNames: Object.freeze(['Membros em Espera', 'Membros Espera', 'MEMBROS_EM_ESPERA'])
    }),
    former: Object.freeze({
      type: 'EX_MEMBROS',
      registryKeys: Object.freeze(['PESSOAS_EX_MEMBROS', 'EX_MEMBROS', 'PESSOAS_EX_MEMBERS', 'PESSOAS_EX_MEMBROS_BASE']),
      sheetNames: Object.freeze(['Ex-Membros', 'Ex Membros', 'Ex_Membros'])
    })
  }),
  yesValues: Object.freeze(['SIM', 'S', 'TRUE', '1', 'YES']),
  noValues: Object.freeze(['NAO', 'N', 'FALSE', '0', 'NO']),
  headers: Object.freeze({
    profileKey: Object.freeze(['PERFIL_PORTAL', 'PERFIL', 'PROFILE', 'CHAVE_PERFIL']),
    profileName: Object.freeze(['NOME', 'NOME_PERFIL', 'DESCRICAO', 'DESCRICAO_PERFIL']),
    permission: Object.freeze(['PERMISSAO', 'PERMISSAO_KEY', 'CHAVE', 'PERMISSION']),
    configKey: Object.freeze(['CHAVE', 'KEY', 'CONFIG', 'CONFIG_KEY']),
    configValue: Object.freeze(['VALOR', 'VALUE']),
    configDescription: Object.freeze(['DESCRICAO', 'DESCRIÇÃO', 'DESCRIPTION']),
    active: Object.freeze(['ATIVO', 'ACTIVE']),
    email: Object.freeze(['EMAIL', 'E-MAIL', 'Email', 'E-mail', 'EMAIL_PRINCIPAL']),
    name: Object.freeze(['NOME', 'NOME_COMPLETO', 'NOME_MEMBRO', 'MEMBRO', 'Membro', 'Nome']),
    rga: Object.freeze(['RGA']),
    status: Object.freeze(['STATUS', 'STATUS_CADASTRAL', 'STATUS_REGISTRO', 'SITUACAO', 'SITUACAO_GERAL']),
    portalActive: Object.freeze(['PORTAL_ATIVO']),
    portalProfile: Object.freeze(['PERFIL_PORTAL']),
    portalObs: Object.freeze(['PORTAL_OBS'])
  }),
  accessLogHeaders: Object.freeze([
    'TIMESTAMP',
    'EMAIL',
    'UID_FIREBASE',
    'NOME',
    'PERFIL_PORTAL',
    'ACAO',
    'RESULTADO',
    'MOTIVO',
    'ORIGEM',
    'USER_AGENT',
    'OBS'
  ])
});

var __core_portal_registry_sheet_execution_cache = {};
var __core_portal_registry_spreadsheet_execution_cache = {};
var __core_portal_trace_stages = [];

function corePortalTraceReset_() {
  __core_portal_trace_stages = [];
}

function corePortalTraceStage_(stage, startedAt, code, origin) {
  __core_portal_trace_stages.push(Object.freeze({
    etapa: String(stage || '').slice(0, 80),
    duracaoMs: Math.max(0, Date.now() - Number(startedAt || Date.now())),
    code: String(code || '').slice(0, 80),
    origem: String(origin || '').slice(0, 80)
  }));
}

function corePortalResolveEnvironment_(opts) {
  return core_normalizeDomainEnv_(opts || {});
}

/**
 * Le uma key de configuracao no ambiente explicitamente pedido.
 * Nao usa o helper legado dependente do GEAPA_ENV global.
 */
function corePortalReadRegistryRecordsForEnv_(registryKey, opts) {
  var startedAt = Date.now();
  var options = opts || {};
  var environment = corePortalResolveEnvironment_(options);
  var state = core_domainRegistryEntry_(registryKey, environment, options);
  if (!state.available) {
    throw core_domainResolverError_(
      'PORTAL_REGISTRY_KEY_INDISPONIVEL',
      'Registry key "' + registryKey + '" indisponivel em ' + environment + '.',
      { registryKey: registryKey, ambiente: environment, motivo: state.reason }
    );
  }

  var entry = state.entry || {};
  var spreadsheetId = String(entry.id || '').trim();
  var sheetName = String(entry.sheet || '').trim();
  var cacheKey = [environment, registryKey, spreadsheetId, sheetName].join('|');
  var sheet = __core_portal_registry_sheet_execution_cache[cacheKey];
  if (!sheet) {
    var spreadsheetCacheKey = [environment, spreadsheetId].join('|');
    var spreadsheet = __core_portal_registry_spreadsheet_execution_cache[spreadsheetCacheKey];
    if (!spreadsheet) {
      spreadsheet = core_openSpreadsheetById_(spreadsheetId);
      __core_portal_registry_spreadsheet_execution_cache[spreadsheetCacheKey] = spreadsheet;
    }
    sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw core_domainResolverError_(
        'PORTAL_REGISTRY_ABA_INDISPONIVEL',
        'Aba da key "' + registryKey + '" indisponivel em ' + environment + '.',
        { registryKey: registryKey, ambiente: environment }
      );
    }
    __core_portal_registry_sheet_execution_cache[cacheKey] = sheet;
  }
  var records = core_readSheetRecords_(sheet, { skipBlankRows: true });
  corePortalTraceStage_('registry.' + registryKey, startedAt, 'OK', environment);
  return records;
}

function corePortalNormalizeToken_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  }).replace(/[^A-Z0-9:_-]+/g, '_').replace(/^_+|_+$/g, '');
}

function corePortalNormalizeEmail_(email) {
  var extracted = core_extractEmailAddress_(email);
  return extracted && core_isValidEmail_(extracted) ? extracted.toLowerCase() : '';
}

function corePortalIsYes_(value) {
  return CORE_PORTAL_ACCESS_CFG.yesValues.indexOf(corePortalNormalizeToken_(value)) >= 0;
}

function corePortalIsExplicitNo_(value) {
  return CORE_PORTAL_ACCESS_CFG.noValues.indexOf(corePortalNormalizeToken_(value)) >= 0;
}

function corePortalGetRecordValue_(record, aliases) {
  var names = Array.isArray(aliases) ? aliases : [aliases];
  var keys = Object.keys(record || {});

  for (var i = 0; i < names.length; i++) {
    var wanted = corePortalNormalizeToken_(names[i]);
    for (var j = 0; j < keys.length; j++) {
      if (corePortalNormalizeToken_(keys[j]) === wanted) {
        return record[keys[j]];
      }
    }
  }

  return '';
}

function corePortalRecordIsActive_(record) {
  var raw = corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.active);
  if (raw === '' || raw == null) return false;
  return corePortalIsYes_(raw);
}

function corePortalReadProfiles_(opts) {
  opts = opts || {};
  var records = opts.records || corePortalReadRegistryRecordsForEnv_('PORTAL_PERFIS', opts);
  var out = [];
  var seen = {};

  records.forEach(function(record) {
    if (!corePortalRecordIsActive_(record)) return;

    var profile = corePortalNormalizeToken_(
      corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.profileKey)
    );
    if (!profile || seen[profile]) return;

    seen[profile] = true;
    out.push(Object.freeze({
      perfilPortal: profile,
      nome: String(corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.profileName) || profile).trim(),
      ativo: true
    }));
  });

  out.sort(function(a, b) {
    return a.perfilPortal.localeCompare(b.perfilPortal, 'pt-BR');
  });

  return Object.freeze(out);
}

function corePortalReadPermissions_(opts) {
  opts = opts || {};
  var records = opts.records || corePortalReadRegistryRecordsForEnv_('PORTAL_PERMISSOES', opts);
  var out = [];

  records.forEach(function(record) {
    if (!corePortalRecordIsActive_(record)) return;

    var profile = corePortalNormalizeToken_(
      corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.profileKey)
    );
    var permission = String(
      corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.permission) || ''
    ).trim();

    if (!profile || !permission) return;

    out.push(Object.freeze({
      perfilPortal: profile,
      permissao: permission,
      ativo: true
    }));
  });

  return Object.freeze(out);
}

function corePortalConfigCacheKey_(opts) {
  return CORE_PORTAL_ACCESS_CFG.portalConfigCacheKey + ':' + corePortalResolveEnvironment_(opts || {});
}

function corePortalConfigCacheGet_(opts) {
  try {
    var raw = CacheService.getScriptCache().get(corePortalConfigCacheKey_(opts));
    return raw ? JSON.parse(raw) : null;
  } catch (err) {
    return null;
  }
}

function corePortalConfigCacheSet_(config, opts) {
  try {
    CacheService.getScriptCache().put(
      corePortalConfigCacheKey_(opts),
      JSON.stringify(config || {}),
      CORE_PORTAL_ACCESS_CFG.portalConfigCacheTtlSeconds
    );
  } catch (err) {}
}

function corePortalConfigCacheClear_(opts) {
  try {
    CacheService.getScriptCache().remove(corePortalConfigCacheKey_(opts));
  } catch (err) {}
  return Object.freeze({
    ok: true,
    cacheCleared: true,
    ambiente: corePortalResolveEnvironment_(opts || {})
  });
}

function corePortalConfigIsSensitiveKey_(key) {
  var normalized = corePortalNormalizeToken_(key);
  return /(^|_)(TOKEN|SECRET|SEGREDO|SENHA|PASSWORD|API_KEY|PRIVATE_KEY|CHAVE_PRIVADA|CREDENTIAL|CREDENCIAL)(_|$)/.test(normalized);
}

function corePortalParseConfigValue_(value) {
  var text = String(value == null ? '' : value).trim();
  var token = corePortalNormalizeToken_(text);
  if (CORE_PORTAL_ACCESS_CFG.yesValues.indexOf(token) >= 0) return true;
  if (CORE_PORTAL_ACCESS_CFG.noValues.indexOf(token) >= 0) return false;
  if (/^-?\d+([.,]\d+)?$/.test(text)) return Number(text.replace(',', '.'));
  return text;
}

function corePortalReadConfig_(opts) {
  opts = opts || {};
  if (!opts.records && opts.forceRefresh !== true) {
    var cached = corePortalConfigCacheGet_(opts);
    if (cached) return Object.freeze(cached);
  }

  var records = opts.records || corePortalReadRegistryRecordsForEnv_('PORTAL_CONFIG', opts);
  var out = {};

  records.forEach(function(record) {
    if (!corePortalRecordIsActive_(record)) return;

    var key = corePortalNormalizeToken_(
      corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.configKey)
    );
    if (!key) return;
    if (corePortalConfigIsSensitiveKey_(key)) return;

    out[key] = corePortalParseConfigValue_(
      corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.configValue)
    );
  });

  if (!opts.records) corePortalConfigCacheSet_(out, opts);
  return Object.freeze(out);
}

function corePortalAccessMode_(config) {
  var mode = corePortalNormalizeToken_((config || {}).PORTAL_MODO_ACESSO || 'MEMBROS_ATIVOS');
  if (['TESTE', 'MEMBROS_ATIVOS', 'PUBLICO_LIMITADO'].indexOf(mode) >= 0) return mode;
  return 'MEMBROS_ATIVOS';
}

function corePortalTestEmails_(config) {
  var seen = {};
  var out = [];
  String((config || {}).PORTAL_EMAILS_TESTE || '').split(/[;,\n\r\t ]+/).forEach(function(part) {
    var email = corePortalNormalizeEmail_(part);
    if (email && !seen[email]) {
      seen[email] = true;
      out.push(email);
    }
  });
  return Object.freeze(out);
}

function corePortalIsEmailInTestList_(email, config) {
  var normalized = corePortalNormalizeEmail_(email);
  return !!normalized && corePortalTestEmails_(config).indexOf(normalized) >= 0;
}

function corePortalAccessDeniedMessage_() {
  return 'Seu e-mail nao esta liberado para acessar o Portal GEAPA ou nao possui vinculo ativo no grupo. Entre com o mesmo e-mail cadastrado junto ao GEAPA.';
}

function corePortalLinkAllowsMemberAccess_(link) {
  var tipo = corePortalNormalizeToken_(link && link.tipoVinculo);
  var status = corePortalNormalizeToken_(link && link.statusVinculo);
  if (status !== 'ATIVO' && status !== 'ATIVA') return false;
  return ['MEMBRO_INGRESSANTE', 'MEMBRO_EFETIVO', 'MEMBRO'].indexOf(tipo) >= 0;
}

function corePortalEvaluateAccessMode_(config, email, profileResult, permissions, link) {
  var mode = corePortalAccessMode_(config);
  var profile = corePortalNormalizeToken_(profileResult && profileResult.perfilPortalEfetivo);
  var hasPortalAccess = (permissions || []).indexOf('portal:acessar') >= 0;
  var blockInactive = !Object.prototype.hasOwnProperty.call(config || {}, 'PORTAL_BLOCK_INACTIVE_MEMBERS') ||
    corePortalIsYes_((config || {}).PORTAL_BLOCK_INACTIVE_MEMBERS);

  if (mode === 'TESTE' && !corePortalIsEmailInTestList_(email, config)) {
    return Object.freeze({ allowed: false, mode: mode, reason: 'EMAIL_FORA_PORTAL_EMAILS_TESTE' });
  }

  if (profile === 'VISITANTE' && !corePortalIsYes_((config || {}).AUTH_ALLOW_VISITANTE)) {
    return Object.freeze({ allowed: false, mode: mode, reason: 'VISITANTE_NAO_PERMITIDO' });
  }

  if (mode === 'PUBLICO_LIMITADO' && profile === 'VISITANTE') {
    return Object.freeze({
      allowed: corePortalIsYes_((config || {}).AUTH_ALLOW_VISITANTE),
      mode: mode,
      reason: corePortalIsYes_((config || {}).AUTH_ALLOW_VISITANTE) ? 'PUBLICO_LIMITADO_VISITANTE' : 'VISITANTE_NAO_PERMITIDO'
    });
  }

  if (profile === 'ADMIN' && hasPortalAccess) {
    return Object.freeze({ allowed: true, mode: mode, reason: 'ADMIN_COM_PORTAL_ACESSAR' });
  }

  if (blockInactive && !corePortalLinkAllowsMemberAccess_(link)) {
    return Object.freeze({ allowed: false, mode: mode, reason: 'VINCULO_ATIVO_AUSENTE' });
  }

  return Object.freeze({ allowed: true, mode: mode, reason: 'MEMBRO_ATIVO_COM_PORTAL_ACESSAR' });
}

function corePortalBuildPermissionsForProfile_(perfilPortal, opts) {
  opts = opts || {};
  var profile = corePortalNormalizeToken_(perfilPortal);
  if (!profile) return Object.freeze([]);

  var permissions = opts.permissions || corePortalReadPermissions_(opts.permissionsOpts || {});
  var seen = {};
  var out = [];

  permissions.forEach(function(item) {
    if (corePortalNormalizeToken_(item.perfilPortal) !== profile) return;
    var permission = String(item.permissao || '').trim();
    if (!permission || seen[permission]) return;
    seen[permission] = true;
    out.push(permission);
  });

  out.sort();
  return Object.freeze(out);
}

const CORE_PORTAL_V2_REQUIRED_PROFILES = Object.freeze([
  'ADMIN',
  'DIRETORIA',
  'SECRETARIA',
  'COMUNICACAO',
  'CONSELHO',
  'MEMBRO',
  'MEMBRO_INGRESSANTE',
  'EGRESSO',
  'COLABORADOR',
  'EXTERNO',
  'VISITANTE'
]);

const CORE_PORTAL_V2_PROFILE_LEVELS = Object.freeze({
  ADMIN: 100,
  DIRETORIA: 80,
  SECRETARIA: 70,
  COMUNICACAO: 60,
  CONSELHO: 40,
  COLABORADOR: 30,
  MEMBRO: 10,
  MEMBRO_INGRESSANTE: 8,
  EGRESSO: 5,
  EXTERNO: 2,
  VISITANTE: 1
});

const CORE_PORTAL_V2_MIN_PERMISSIONS = Object.freeze({
  MEMBRO_INGRESSANTE: Object.freeze([
    'portal:acessar',
    'situacao:ver_propria'
  ]),
  EGRESSO: Object.freeze([
    'portal:acessar',
    'situacao:ver_propria',
    'apresentacoes:ver_ate_saida',
    'certificados:ver_proprios'
  ]),
  CONSELHO: Object.freeze([
    'portal:acessar',
    'apresentacoes:ver_publicas',
    'eixos:ver'
  ]),
  COLABORADOR: Object.freeze([
    'portal:acessar',
    'apresentacoes:ver_publicas',
    'eixos:ver'
  ]),
  EXTERNO: Object.freeze([
    'portal:acessar',
    'apresentacoes:ver_publicas',
    'inscricoes:ver_proprias'
  ]),
  VISITANTE: Object.freeze([
    'apresentacoes:ver_publicas'
  ])
});

function corePortalNormalizePermission_(permission) {
  return String(permission || '').trim().toLowerCase();
}

function corePortalSplitProfiles_(value) {
  var seen = {};
  var out = [];
  String(value || '').split(/[;,|]/).forEach(function(part) {
    var profile = corePortalNormalizeToken_(part);
    if (!profile || seen[profile]) return;
    seen[profile] = true;
    out.push(profile);
  });
  return out;
}

function corePortalActiveProfileMap_(profiles) {
  var out = {};
  (profiles || corePortalReadProfiles_()).forEach(function(profile) {
    out[corePortalNormalizeToken_(profile.perfilPortal)] = profile;
  });
  return out;
}

function corePortalPermissionsByProfile_(permissions) {
  var out = {};
  (permissions || corePortalReadPermissions_()).forEach(function(item) {
    var profile = corePortalNormalizeToken_(item.perfilPortal);
    var permission = corePortalNormalizePermission_(item.permissao);
    if (!profile || !permission) return;
    if (!out[profile]) out[profile] = [];
    if (out[profile].indexOf(permission) < 0) out[profile].push(permission);
  });
  Object.keys(out).forEach(function(profile) {
    out[profile].sort();
  });
  return out;
}

function corePortalIsExceptionActive_(record, refDate) {
  var status = corePortalNormalizeToken_(record && record.STATUS);
  if (['REVOGADO', 'ENCERRADO', 'INATIVO', 'EXPIRADO'].indexOf(status) >= 0) return false;
  var today = refDate || new Date();
  var start = core_domainsV2Date_(record && record.DATA_INICIO);
  var end = core_domainsV2Date_(record && record.DATA_FIM);
  if (start && start > today) return false;
  if (end && end < today) return false;
  return true;
}

function corePortalExceptionBlocks_(record) {
  var status = corePortalNormalizeToken_(record && record.STATUS);
  return ['BLOQUEADO', 'NEGADO', 'SUSPENSO'].indexOf(status) >= 0;
}

function corePortalGetBundleById_(idPessoa, opts) {
  var id = String(idPessoa || '').trim();
  return id ? corePessoasGetById_(id, opts || {}) : null;
}

function corePortalGetBundleByEmail_(email, opts) {
  var normalizedEmail = corePortalNormalizeEmail_(email);
  return normalizedEmail ? corePessoasFindByEmail_(normalizedEmail, opts || {}) : null;
}

function corePortalResolveInput_(entrada) {
  var raw = entrada;
  var opts = {};

  if (entrada && typeof entrada === 'object') {
    opts = entrada;
    raw = entrada.email || entrada.emailOuRga || entrada.identificador || entrada.idPessoa || entrada.rga || '';
  }

  var text = String(raw || '').trim();
  var email = corePortalNormalizeEmail_(opts.email || text);
  var idPessoa = String(opts.idPessoa || '').trim();
  var rga = String(opts.rga || '').trim();
  var identifier = String(opts.identificador || opts.emailOuRga || text).trim();

  if (!idPessoa && /^PES-\d+$/i.test(text)) idPessoa = text.toUpperCase();
  if (!email && !rga && !idPessoa && text && text.indexOf('@') < 0) rga = text;

  return Object.freeze({
    raw: text,
    email: email,
    idPessoa: idPessoa,
    rga: rga,
    identificador: identifier
  });
}

function corePortalGetBundleByInput_(entrada, opts) {
  var input = corePortalResolveInput_(entrada);
  var options = opts || {};
  if (input.idPessoa) return corePortalGetBundleById_(input.idPessoa, options);
  if (input.email) return corePortalGetBundleByEmail_(input.email, options);
  if (input.rga) return corePessoasFindByRga_(input.rga, options);
  if (input.identificador && input.identificador.indexOf('@') >= 0) {
    return corePortalGetBundleByEmail_(input.identificador, options);
  }
  if (input.identificador) return corePessoasFindByRga_(input.identificador, options);
  return null;
}

function corePortalGetCurrentLink_(bundle) {
  var resumo = (bundle && bundle.resumoOperacional) || {};
  var tipoResumo = corePortalNormalizeToken_(resumo.TIPO_VINCULO_ATUAL);
  if (tipoResumo) {
    return {
      tipoVinculo: tipoResumo,
      statusVinculo: corePortalNormalizeToken_(resumo.STATUS_VINCULO_ATUAL),
      dataFim: ''
    };
  }
  var vinculo = core_domainsV2PickCurrentVinculo_((bundle && bundle.vinculos) || []) || {};
  return {
    tipoVinculo: corePortalNormalizeToken_(vinculo.TIPO_VINCULO),
    statusVinculo: corePortalNormalizeToken_(vinculo.STATUS_VINCULO),
    dataFim: vinculo.DATA_FIM || ''
  };
}

function corePortalGetEgressoExitDate_(bundle) {
  var best = null;
  ((bundle && bundle.vinculos) || []).forEach(function(vinculo) {
    var tipo = corePortalNormalizeToken_(vinculo.TIPO_VINCULO);
    var status = corePortalNormalizeToken_(vinculo.STATUS_VINCULO);
    var date = core_domainsV2Date_(vinculo.DATA_FIM);
    if (tipo !== 'MEMBRO_EFETIVO' || !date) return;
    if (status === 'ATIVO' || corePortalIsYes_(vinculo.ATIVO)) return;
    if (!best || date > best) best = date;
  });
  return best;
}

function corePortalGetIdentityMarker_(bundle, tipoIdentificador) {
  var wanted = corePortalNormalizeToken_(tipoIdentificador);
  var found = '';
  ((bundle && bundle.identificadores) || []).forEach(function(record) {
    if (found) return;
    if (corePortalNormalizeToken_(record.TIPO_IDENTIFICADOR) === wanted && corePortalIsYes_(record.ATIVO)) {
      found = String(record.VALOR_IDENTIFICADOR || '').trim();
    }
  });
  return found;
}

function corePortalGetVigenciasResumoAtualSafe_(idPessoa, opts) {
  try {
    return coreVigenciasGetCurrentSummaryByPessoa_(idPessoa, opts || {}) || null;
  } catch (err) {
    return null;
  }
}

function corePortalSanitizeCurrentFunctionFromResumo_(record) {
  record = record || {};
  var cargoNome = String(record.CARGO_FUNCAO_ATUAL || '').trim();
  if (!cargoNome) return null;
  return Object.freeze({
    cargoKey: '',
    cargoNome: cargoNome,
    tipoFuncao: String(record.TIPO_FUNCAO_ATUAL || '').trim(),
    grupoCargo: String(record.GRUPO_FUNCAO_ATUAL || '').trim(),
    dataInicio: record.DATA_INICIO_FUNCAO_ATUAL || '',
    dataFimPrevista: record.DATA_FIM_PREVISTA || ''
  });
}

function corePortalListCurrentFunctionsSafe_(idPessoa, opts) {
  opts = opts || {};
  var resumo = opts.vigenciasResumo || corePortalGetVigenciasResumoAtualSafe_(idPessoa);
  var current = corePortalSanitizeCurrentFunctionFromResumo_(resumo);
  return Object.freeze(current ? [current] : []);
}

function corePortalListActiveExceptions_(bundle) {
  return ((bundle && bundle.portalExcecoes) || []).filter(function(record) {
    return corePortalIsExceptionActive_(record);
  });
}

function corePortalProfilesFromVigenciasResumo_(record) {
  return corePortalSplitProfiles_(record && record.PERFIS_PORTAL_CALCULADOS);
}

function corePortalCalcularPerfilEfetivo_(idPessoa, opts) {
  opts = opts || {};
  var bundle = opts.bundle || corePortalGetBundleById_(idPessoa, opts);
  if (!bundle || !bundle.pessoa) {
    return Object.freeze({
      ok: false,
      idPessoa: String(idPessoa || '').trim(),
      perfilPortalEfetivo: '',
      perfisPortal: Object.freeze([]),
      origemPerfil: '',
      portalAtivo: false,
      motivoBloqueio: 'PESSOA_NAO_ENCONTRADA'
    });
  }

  var id = String(bundle.pessoa.ID_PESSOA || idPessoa || '').trim();
  var profileMap = opts.profileMap || corePortalActiveProfileMap_(opts.profiles);
  var resumo = bundle.resumoOperacional || {};
  var vigenciasResumo = opts.vigenciasResumo || null;
  var link = corePortalGetCurrentLink_(bundle);
  var exceptions = corePortalListActiveExceptions_(bundle);
  var blocked = exceptions.some(corePortalExceptionBlocks_);
  var profiles = [];
  var origem = '';

  exceptions.forEach(function(record) {
    corePortalSplitProfiles_(record.PERFIL_EXTRA).forEach(function(profile) {
      if (profile === 'ADMIN' && profiles.indexOf('ADMIN') < 0) {
        profiles.unshift('ADMIN');
        origem = 'PORTAL_ACESSOS_EXCECOES_ADMIN';
      }
    });
  });

  if (!profiles.length) {
    profiles = corePortalSplitProfiles_(resumo.PERFIL_PORTAL_CALCULADO).filter(function(profile) {
      return profile !== 'ADMIN';
    });
    if (profiles.length) origem = 'PESSOAS_RESUMO_OPERACIONAL';
  }

  if (!profiles.length) {
    if (!vigenciasResumo) vigenciasResumo = corePortalGetVigenciasResumoAtualSafe_(id, opts);
    profiles = corePortalProfilesFromVigenciasResumo_(vigenciasResumo).filter(function(profile) {
      return profile !== 'ADMIN';
    });
    if (profiles.length) origem = 'VIGENCIAS_RESUMO_ATUAL';
  }

  if (!profiles.length) {
    if (link.tipoVinculo === 'MEMBRO_EFETIVO' && link.statusVinculo === 'ATIVO') {
      profiles = ['MEMBRO'];
      origem = 'VINCULO_MEMBRO_EFETIVO';
    } else if (link.tipoVinculo === 'MEMBRO_INGRESSANTE' && link.statusVinculo === 'ATIVO') {
      profiles = ['MEMBRO_INGRESSANTE'];
      origem = 'VINCULO_MEMBRO_INGRESSANTE';
    } else if (link.tipoVinculo === 'EGRESSO' || link.tipoVinculo === 'EX_MEMBRO') {
      profiles = ['EGRESSO'];
      origem = 'VINCULO_EGRESSO';
    } else if (corePortalGetIdentityMarker_(bundle, 'ID_PROFESSOR')) {
      profiles = ['COLABORADOR'];
      origem = 'IDENTIFICADOR_COLABORADOR';
    } else if (corePortalGetIdentityMarker_(bundle, 'ID_PARTICIPANTE_EXTERNO')) {
      profiles = ['EXTERNO'];
      origem = 'IDENTIFICADOR_EXTERNO';
    } else {
      profiles = ['VISITANTE'];
      origem = 'SEM_VINCULO_RECONHECIDO';
    }
  }

  exceptions.forEach(function(record) {
    corePortalSplitProfiles_(record.PERFIL_EXTRA).forEach(function(profile) {
      if (profiles.indexOf(profile) < 0) profiles.push(profile);
    });
  });

  if (link.tipoVinculo === 'MEMBRO_INGRESSANTE' && link.statusVinculo === 'ATIVO') {
    profiles = ['MEMBRO_INGRESSANTE'];
    origem = 'VINCULO_MEMBRO_INGRESSANTE';
  }

  var invalidProfiles = profiles.filter(function(profile) {
    return !profileMap[profile];
  });
  var profileEffective = profiles[0] || '';
  var portalAtivo = !blocked && !invalidProfiles.length && !!profileEffective;
  if (resumo.PORTAL_ATIVO && corePortalIsExplicitNo_(resumo.PORTAL_ATIVO) && profileEffective !== 'EGRESSO') {
    portalAtivo = false;
  }

  return Object.freeze({
    ok: invalidProfiles.length === 0,
    idPessoa: id,
    perfilPortalEfetivo: profileEffective,
    perfisPortal: Object.freeze(profiles),
    perfisInvalidos: Object.freeze(invalidProfiles),
    origemPerfil: origem,
    portalAtivo: portalAtivo,
    motivoBloqueio: blocked ? 'EXCECAO_PORTAL_BLOQUEIO' : (invalidProfiles.length ? 'PERFIL_PORTAL_INVALIDO' : ''),
    regraApresentacoes: profileEffective === 'EGRESSO' ? 'ATE_DATA_SAIDA' : ''
  });
}

function corePortalListarPermissoesEfetivas_(idPessoa, opts) {
  opts = opts || {};
  var bundle = opts.bundle || corePortalGetBundleById_(idPessoa, opts);
  var profileResult = opts.profileResult || corePortalCalcularPerfilEfetivo_(idPessoa, opts);
  if (!profileResult.ok && !profileResult.perfisPortal.length) {
    return Object.freeze({
      ok: false,
      idPessoa: String(idPessoa || '').trim(),
      permissoes: Object.freeze([]),
      origemPermissoes: '',
      motivo: profileResult.motivoBloqueio || 'PERFIL_NAO_RESOLVIDO'
    });
  }

  if (profileResult.perfilPortalEfetivo === 'MEMBRO_INGRESSANTE') {
    return Object.freeze({
      ok: true,
      idPessoa: profileResult.idPessoa,
      perfilPortalEfetivo: 'MEMBRO_INGRESSANTE',
      perfisPortal: Object.freeze(['MEMBRO_INGRESSANTE']),
      permissoes: Object.freeze(CORE_PORTAL_V2_MIN_PERMISSIONS.MEMBRO_INGRESSANTE.slice()),
      origemPermissoes: 'MINIMO_MEMBRO_INGRESSANTE',
      usaCargosConfigComoFonteFinal: false
    });
  }

  var permissionsByProfile = opts.permissionsByProfile || corePortalPermissionsByProfile_(opts.permissions);
  var seen = {};
  var out = [];
  profileResult.perfisPortal.forEach(function(profile) {
    (permissionsByProfile[profile] || []).forEach(function(permission) {
      if (!seen[permission]) {
        seen[permission] = true;
        out.push(permission);
      }
    });
  });

  ((bundle && bundle.portalExcecoes) || []).filter(corePortalIsExceptionActive_).forEach(function(record) {
    String(record.PERMISSAO_EXTRA || '').split(/[;,|]/).forEach(function(part) {
      var permission = corePortalNormalizePermission_(part);
      if (permission && !seen[permission]) {
        seen[permission] = true;
        out.push(permission);
      }
    });
  });

  out.sort();
  return Object.freeze({
    ok: true,
    idPessoa: profileResult.idPessoa,
    perfilPortalEfetivo: profileResult.perfilPortalEfetivo,
    perfisPortal: profileResult.perfisPortal,
    permissoes: Object.freeze(out),
    origemPermissoes: 'PORTAL_PERMISSOES',
    usaCargosConfigComoFonteFinal: false
  });
}

function corePortalResolverUsuarioAtual_(entrada, opts) {
  opts = opts || {};
  corePortalTraceReset_();
  var startedAt = Date.now();
  var environment = corePortalResolveEnvironment_(opts);
  var input = corePortalResolveInput_(entrada);
  var environmentOptions = Object.assign({}, opts, { ambiente: environment });
  var config = opts.config || corePortalReadConfig_(
    Object.assign({}, opts.configOpts || {}, { ambiente: environment })
  );
  var profiles = opts.profiles || corePortalReadProfiles_(
    Object.assign({}, opts.profilesOpts || {}, { ambiente: environment })
  );
  var permissions = opts.permissions || corePortalReadPermissions_(
    Object.assign({}, opts.permissionsOpts || {}, { ambiente: environment })
  );
  var bundle = corePortalGetBundleByInput_(entrada, environmentOptions);
  if (!bundle || !bundle.pessoa) {
    return Object.freeze({
      ok: false,
      autenticado: false,
      email: input.email,
      idPessoa: input.idPessoa,
      rga: input.rga,
      modoAcesso: corePortalAccessMode_(config),
      motivoBloqueio: 'PESSOA_NAO_ENCONTRADA',
      mensagemBloqueio: corePortalAccessDeniedMessage_(),
      metaTecnica: Object.freeze({
        traceId: String(opts.traceId || opts.requestId || '').trim().slice(0, 80),
        ambiente: environment,
        etapa: 'corePortalResolverUsuarioAtual',
        duracaoMs: Math.max(0, Date.now() - startedAt),
        etapas: Object.freeze(__core_portal_trace_stages.slice())
      })
    });
  }
  var idPessoa = String(bundle.pessoa.ID_PESSOA || '').trim();
  var vigenciasResumo = opts.vigenciasResumo || corePortalGetVigenciasResumoAtualSafe_(idPessoa, environmentOptions);
  var profileResult = corePortalCalcularPerfilEfetivo_(idPessoa, Object.assign({}, opts, {
    bundle: bundle,
    vigenciasResumo: vigenciasResumo,
    profiles: profiles
  }));
  var permissionsResult = corePortalListarPermissoesEfetivas_(idPessoa, Object.assign({}, opts, {
    bundle: bundle,
    profileResult: profileResult,
    permissions: permissions
  }));
  var resumo = bundle.resumoOperacional || {};
  var link = corePortalGetCurrentLink_(bundle);
  var permissoes = permissionsResult.permissoes || [];
  var cargosAtuais = corePortalListCurrentFunctionsSafe_(idPessoa, { vigenciasResumo: vigenciasResumo });
  var emailPrincipal = corePortalNormalizeEmail_(resumo.EMAIL || bundle.pessoa.EMAIL_PRINCIPAL || input.email || '');
  var modeDecision = corePortalEvaluateAccessMode_(config, emailPrincipal, profileResult, permissoes, link);
  var hasAccess = permissoes.indexOf('portal:acessar') >= 0 || (
    profileResult.perfilPortalEfetivo === 'VISITANTE' && corePortalIsYes_(config.AUTH_ALLOW_VISITANTE)
  );
  var portalAtivo = profileResult.portalAtivo && hasAccess && modeDecision.allowed;
  var motivoBloqueio = '';
  if (!portalAtivo) {
    if (!profileResult.portalAtivo) {
      motivoBloqueio = profileResult.motivoBloqueio || 'PERFIL_PORTAL_INATIVO';
    } else if (!hasAccess) {
      motivoBloqueio = 'PERMISSAO_PORTAL_ACESSAR_AUSENTE';
    } else {
      motivoBloqueio = modeDecision.reason || 'ACESSO_PORTAL_BLOQUEADO';
    }
  }
  var cargoFuncaoAtual = String(resumo.CARGO_FUNCAO_ATUAL || '').trim();
  if (!cargoFuncaoAtual && vigenciasResumo) cargoFuncaoAtual = String(vigenciasResumo.CARGO_FUNCAO_ATUAL || '').trim();
  if (!cargoFuncaoAtual && cargosAtuais.length) {
    cargoFuncaoAtual = cargosAtuais.map(function(item) {
      return item.cargoNome || item.cargoKey;
    }).filter(String).join('; ');
  }

  return Object.freeze({
    ok: true,
    autenticado: true,
    idPessoa: idPessoa,
    nomeExibicao: String(resumo.NOME_EXIBICAO || bundle.pessoa.NOME_EXIBICAO || bundle.pessoa.NOME_COMPLETO || '').trim(),
    email: emailPrincipal,
    rga: String(resumo.RGA || (bundle.membrosDetalhes && bundle.membrosDetalhes.RGA) || '').trim(),
    tipoVinculoAtual: resumo.TIPO_VINCULO_ATUAL || link.tipoVinculo || '',
    statusVinculoAtual: resumo.STATUS_VINCULO_ATUAL || link.statusVinculo || '',
    cargoFuncaoAtual: cargoFuncaoAtual,
    cargosAtuais: cargosAtuais,
    perfilPortalEfetivo: profileResult.perfilPortalEfetivo,
    perfisPortal: profileResult.perfisPortal,
    permissoes: Object.freeze(permissoes.slice()),
    portalAtivo: portalAtivo,
    modoAcesso: modeDecision.mode,
    motivoBloqueio: motivoBloqueio,
    mensagemBloqueio: portalAtivo ? '' : corePortalAccessDeniedMessage_(),
    origemPerfil: profileResult.origemPerfil,
    origemPermissoes: permissionsResult.origemPermissoes,
    regraApresentacoes: profileResult.regraApresentacoes || '',
    metaTecnica: Object.freeze({
      traceId: String(opts.traceId || opts.requestId || '').trim().slice(0, 80),
      ambiente: environment,
      etapa: 'corePortalResolverUsuarioAtual',
      duracaoMs: Math.max(0, Date.now() - startedAt),
      etapas: Object.freeze(__core_portal_trace_stages.slice())
    })
  });
}

const CORE_PORTAL_FIRESTORE_USER_SNAPSHOT_VERSION = 'portal-user-v2';

function corePortalMaskEmailForDiagnostic_(email) {
  var normalized = corePortalNormalizeEmail_(email || '');
  var parts = normalized.split('@');
  if (parts.length !== 2 || !parts[0] || !parts[1]) return '';
  return parts[0].slice(0, Math.min(2, parts[0].length)) + '***@' + parts[1];
}

function corePortalTruncateUidForLog_(uid) {
  var value = String(uid || '').trim();
  if (!value) return '';
  return value.length <= 10 ? value : value.slice(0, 6) + '...' + value.slice(-4);
}

function corePortalFirestoreOperationalProfile_(session) {
  var profiles = corePortalNormalizeFirestoreStringArray_((session && session.perfisPortal) || []);
  var effective = corePortalNormalizeToken_(session && session.perfilPortalEfetivo || '');
  var candidates = [effective].concat(profiles.map(corePortalNormalizeToken_));
  if (candidates.indexOf('ADMIN_TECNICO') >= 0 || candidates.indexOf('ADMIN') >= 0) return 'ADMIN_TECNICO';
  if (candidates.indexOf('SECRETARIA') >= 0) return 'SECRETARIA';
  if (candidates.indexOf('DIRETORIA') >= 0) return 'DIRETORIA';
  if (candidates.indexOf('ORIENTADOR') >= 0) return 'ORIENTADOR';
  return 'MEMBRO';
}

function corePortalFirestorePermissionMap_(permissions) {
  var out = {};
  corePortalNormalizeFirestoreStringArray_(permissions).forEach(function(permission) {
    out[permission] = true;
  });
  return Object.freeze(out);
}

function corePortalNormalizeFirestoreStringArray_(values) {
  var seen = {};
  var out = [];
  (Array.isArray(values) ? values : []).forEach(function(value) {
    var text = String(value || '').trim();
    if (!text || seen[text]) return;
    seen[text] = true;
    out.push(text);
  });
  return Object.freeze(out);
}

function corePortalBuildFirestoreUserSnapshot_(entrada, opts) {
  opts = opts || {};
  var sessao = opts.sessao || opts.session || opts.resolvedSession || null;
  if (!sessao) {
    sessao = corePortalResolverUsuarioAtual_(entrada, opts);
  }
  if (!sessao || sessao.ok === false || sessao.autenticado === false) {
    return Object.freeze({
      ok: false,
      code: sessao && sessao.motivoBloqueio ? sessao.motivoBloqueio : 'USUARIO_PORTAL_NAO_RESOLVIDO',
      message: 'Usuario do Portal nao resolvido pela PESSOAS v2.'
    });
  }

  var entradaObj = entrada && typeof entrada === 'object' ? entrada : {};
  var uid = String(opts.uid || entradaObj.uid || entradaObj.firebaseUid || sessao.uid || '').trim();
  var agora = new Date();
  var cacheTtlMs = Math.max(0, Number(opts.cacheTtlMs || opts.ttlMs || 6 * 60 * 60 * 1000));
  var cacheUpdatedAt = agora.toISOString();
  var cacheExpiresAtDate = opts.cacheExpiresAt
    ? new Date(opts.cacheExpiresAt)
    : new Date(agora.getTime() + cacheTtlMs);
  var cacheExpiresAt = isNaN(cacheExpiresAtDate.getTime())
    ? new Date(agora.getTime() + cacheTtlMs).toISOString()
    : cacheExpiresAtDate.toISOString();
  var sourceUpdatedAtValue = opts.sourceUpdatedAt || sessao.sourceUpdatedAt || sessao.atualizadoEm || cacheUpdatedAt;
  var sourceUpdatedAtDate = new Date(sourceUpdatedAtValue);
  var sourceUpdatedAt = isNaN(sourceUpdatedAtDate.getTime())
    ? cacheUpdatedAt
    : sourceUpdatedAtDate.toISOString();
  var existingSnapshot = opts.existingSnapshot || {};
  var permissions = corePortalNormalizeFirestoreStringArray_(sessao.permissoes);
  var roles = corePortalNormalizeFirestoreStringArray_(sessao.perfisPortal);
  var active = sessao.portalAtivo === true;
  // O documento privado deve acompanhar a identidade Firebase que possui o UID.
  // O e-mail canonico da PESSOAS v2 pode ser diferente quando o membro entrou
  // por um alias previamente resolvido e confirmado pelo Core.
  var normalizedEmail = corePortalNormalizeEmail_(
    opts.authenticatedEmail || opts.firebaseEmail || sessao.email || ''
  );
  var lastLoginAtValue = opts.lastLoginAt || cacheUpdatedAt;
  var lastLoginAtDate = new Date(lastLoginAtValue);
  var lastLoginAt = isNaN(lastLoginAtDate.getTime()) ? cacheUpdatedAt : lastLoginAtDate.toISOString();
  var provisionedAtValue = existingSnapshot.provisionedAt || opts.provisionedAt || cacheUpdatedAt;
  var provisionedAtDate = new Date(provisionedAtValue);
  var provisionedAt = isNaN(provisionedAtDate.getTime()) ? cacheUpdatedAt : provisionedAtDate.toISOString();
  var effectiveProfile = String(sessao.perfilPortalEfetivo || '').trim();

  return Object.freeze({
    uid: uid,
    idPessoa: String(sessao.idPessoa || '').trim(),
    nomePublico: String(sessao.nomeExibicao || '').trim(),
    nomeExibicao: String(sessao.nomeExibicao || '').trim(),
    email: normalizedEmail,
    emailNormalizado: normalizedEmail,
    perfilOperacional: corePortalFirestoreOperationalProfile_(sessao),
    ativo: active,
    podeAcessarPortal: active,
    podeLerDadosPrivados: active,
    roles: roles,
    permissions: corePortalFirestorePermissionMap_(permissions),
    portalAtivo: active,
    perfilPortalEfetivo: effectiveProfile,
    perfisPortal: roles,
    permissoes: permissions,
    source: 'PESSOAS_V2',
    sourceSystem: 'geapa-core',
    sourceUpdatedAt: sourceUpdatedAt,
    cacheUpdatedAt: cacheUpdatedAt,
    cacheExpiresAt: cacheExpiresAt,
    lastLoginAt: lastLoginAt,
    provisionedAt: provisionedAt,
    stale: !active,
    staleReason: active ? '' : String(opts.staleReason || sessao.motivoBloqueio || 'USUARIO_NAO_AUTORIZADO').trim(),
    schemaVersion: CORE_PORTAL_FIRESTORE_USER_SNAPSHOT_VERSION
  });
}

function corePortalReadFirestoreUserSnapshotByUid_(uid, opts) {
  opts = opts || {};
  var id = String(uid || '').trim();
  if (!id) {
    return Object.freeze({ ok: false, found: false, code: 'UID_FIREBASE_AUSENTE' });
  }

  var response = coreFirestoreEnvironmentGetDocument_('portalUsers/' + id, opts);
  var result = {
    ok: response.ok,
    found: response.found,
    reader: 'APPS_SCRIPT_FIRESTORE_REST',
    code: response.code === 'FIRESTORE_GET_OK' ? 'FIRESTORE_READ_OK' : response.code,
    httpStatus: response.httpStatus || 0
  };
  if (response.found) result.snapshot = response.data;
  if (response.firestoreError) result.firestoreError = response.firestoreError;
  return Object.freeze(result);
}
function corePortalFirestoreEnvironment_(opts) {
  return coreFirestoreNormalizeEnvironment_(opts && (opts.ambiente || opts.environment));
}
function corePortalPostFirestoreUserSnapshot_(snapshot, opts) {
  opts = opts || {};
  if (!snapshot || snapshot.ok === false) {
    return Object.freeze({ ok: false, synced: false, code: 'SNAPSHOT_INVALIDO' });
  }

  if (!String(snapshot.uid || '').trim()) {
    return Object.freeze({
      ok: false,
      synced: false,
      code: 'UID_FIREBASE_AUSENTE',
      message: 'Informe uid no opts ou entrada para gravar em portalUsers/{uid}.'
    });
  }

  var response = coreFirestoreEnvironmentSetDocument_('portalUsers/' + snapshot.uid, snapshot, Object.assign({}, opts, {
    dryRun: opts.dryRun === true
  }));
  var result = {
    ok: response.ok,
    synced: response.written === true,
    writer: 'APPS_SCRIPT_FIRESTORE_REST',
    code: response.code === 'FIRESTORE_SET_OK'
      ? 'FIRESTORE_SYNC_OK'
      : (response.code === 'FIRESTORE_SET_FALHOU' ? 'FIRESTORE_SYNC_FALHOU' : response.code),
    httpStatus: response.httpStatus || 0
  };
  if (opts.dryRun === true) result.snapshot = snapshot;
  if (response.firestoreError) result.firestoreError = response.firestoreError;

  return Object.freeze(result);
}
function corePortalSincronizarUsuarioFirestore_(entrada, opts) {
  opts = opts || {};
  var environment = corePortalFirestoreEnvironment_(opts);
  opts = Object.assign({}, opts, { ambiente: environment, environment: environment });
  var entradaObj = entrada && typeof entrada === 'object' ? entrada : {};
  var uid = String(opts.uid || entradaObj.uid || entradaObj.firebaseUid || '').trim();
  var existingSnapshot = opts.existingSnapshot || null;
  if (!existingSnapshot && uid) {
    var existingResult = corePortalReadFirestoreUserSnapshotByUid_(uid, opts);
    if (existingResult && existingResult.found) existingSnapshot = existingResult.snapshot;
  }
  var snapshot = corePortalBuildFirestoreUserSnapshot_(entradaObj, Object.assign({}, opts, {
    existingSnapshot: existingSnapshot || {}
  }));
  return corePortalPostFirestoreUserSnapshot_(snapshot, Object.assign({}, opts, {
    merge: opts.merge === true
  }));
}

function corePortalProvisionarFirestoreUserAutenticado_(firebaseIdentity, opts) {
  opts = opts || {};
  var identity = firebaseIdentity || {};
  var uid = String(identity.uid || '').trim();
  var email = corePortalNormalizeEmail_(identity.email || '');
  if (opts.identityVerified !== true) {
    return Object.freeze({ ok: false, synced: false, code: 'IDENTIDADE_FIREBASE_NAO_VERIFICADA' });
  }
  if (!uid || !email) {
    return Object.freeze({ ok: false, synced: false, code: 'IDENTIDADE_FIREBASE_INCOMPLETA' });
  }
  if (identity.emailVerified !== true) {
    return Object.freeze({ ok: false, synced: false, code: 'FIREBASE_EMAIL_NAO_VERIFICADO' });
  }
  corePortalFirestoreEnvironment_(opts);

  var session = opts.sessao || corePortalResolverUsuarioAtual_({ email: email }, opts);
  if (!session || session.ok === false || session.autenticado === false) {
    return Object.freeze({ ok: false, synced: false, code: 'IDENTIDADE_FIREBASE_DIVERGENTE' });
  }

  var sessionEmail = corePortalNormalizeEmail_(session.email || '');
  if (sessionEmail !== email) {
    // Nao confia apenas na anotacao enviada pelo Portal. Resolve novamente o
    // e-mail Firebase na fonte oficial e exige que ele aponte para a mesma
    // pessoa antes de aceitar um alias.
    var sessionByFirebaseEmail = corePortalResolverUsuarioAtual_({ email: email }, opts);
    var suppliedPersonId = String(session.idPessoa || '').trim();
    var resolvedPersonId = String(sessionByFirebaseEmail && sessionByFirebaseEmail.idPessoa || '').trim();
    if (
      !sessionByFirebaseEmail ||
      sessionByFirebaseEmail.ok === false ||
      sessionByFirebaseEmail.autenticado === false ||
      !suppliedPersonId ||
      !resolvedPersonId ||
      suppliedPersonId !== resolvedPersonId
    ) {
      return Object.freeze({ ok: false, synced: false, code: 'IDENTIDADE_FIREBASE_DIVERGENTE' });
    }
    session = sessionByFirebaseEmail;
  }
  if (session.portalAtivo !== true) {
    return Object.freeze({ ok: false, synced: false, code: 'USUARIO_NAO_AUTORIZADO' });
  }

  return corePortalSincronizarUsuarioFirestore_({ email: email, uid: uid }, Object.assign({}, opts, {
    uid: uid,
    sessao: session,
    authenticatedEmail: email,
    lastLoginAt: opts.lastLoginAt || new Date().toISOString()
  }));
}

function corePortalMarcarFirestoreUserInativoPorUid_(uid, opts) {
  opts = opts || {};
  corePortalFirestoreEnvironment_(opts);
  var id = String(uid || '').trim();
  if (opts.identityVerified !== true) {
    return Object.freeze({ ok: false, synced: false, code: 'IDENTIDADE_FIREBASE_NAO_VERIFICADA' });
  }
  if (!id) return Object.freeze({ ok: false, synced: false, code: 'UID_FIREBASE_AUSENTE' });

  var existing = corePortalReadFirestoreUserSnapshotByUid_(id, opts);
  if (!existing.ok) return Object.freeze({ ok: false, synced: false, code: existing.code || 'FIRESTORE_READ_FALHOU' });
  if (!existing.found) {
    return Object.freeze({ ok: true, synced: false, skipped: true, code: 'PORTAL_USER_INEXISTENTE_NAO_CRIADO' });
  }

  var now = new Date().toISOString();
  var inactive = Object.freeze({
    uid: id,
    ativo: false,
    podeAcessarPortal: false,
    podeLerDadosPrivados: false,
    portalAtivo: false,
    stale: true,
    staleReason: String(opts.staleReason || 'USUARIO_NAO_AUTORIZADO').trim().slice(0, 120),
    cacheUpdatedAt: now,
    cacheExpiresAt: new Date(Date.now() - 1000).toISOString(),
    sourceSystem: 'geapa-core',
    schemaVersion: CORE_PORTAL_FIRESTORE_USER_SNAPSHOT_VERSION
  });
  return corePortalPostFirestoreUserSnapshot_(inactive, Object.assign({}, opts, { merge: true }));
}

function corePortalSyncFirestoreUserByEmail_(email, opts) {
  opts = opts || {};
  return corePortalSincronizarUsuarioFirestore_({
    email: email,
    uid: opts.uid || ''
  }, opts);
}

function corePortalSyncFirestoreUserByIdPessoa_(idPessoa, opts) {
  opts = opts || {};
  return corePortalSincronizarUsuarioFirestore_({
    idPessoa: idPessoa,
    uid: opts.uid || ''
  }, opts);
}

function corePortalInvalidarCacheFirestoreUsuario_(idPessoaOuEmail, opts) {
  opts = opts || {};
  corePortalFirestoreEnvironment_(opts);
  var chave = String(idPessoaOuEmail || '').trim();
  var entrada = opts.entrada || {};
  if (chave) {
    entrada = Object.assign({}, entrada, chave.indexOf('@') >= 0 ? { email: chave } : { idPessoa: chave });
  }
  if (opts.uid) entrada.uid = opts.uid;

  var snapshot = corePortalBuildFirestoreUserSnapshot_(entrada, Object.assign({}, opts, {
    cacheExpiresAt: new Date(Date.now() - 1000).toISOString()
  }));

  if (!snapshot || snapshot.ok === false) {
    return Object.freeze({ ok: false, synced: false, code: 'SNAPSHOT_INVALIDO' });
  }

  var invalidado = Object.assign({}, snapshot, {
    ativo: false,
    podeAcessarPortal: false,
    podeLerDadosPrivados: false,
    portalAtivo: false,
    stale: true,
    staleReason: String(opts.staleReason || 'CACHE_INVALIDADO').trim().slice(0, 120),
    cacheUpdatedAt: new Date().toISOString(),
    cacheExpiresAt: new Date(Date.now() - 1000).toISOString()
  });

  return corePortalPostFirestoreUserSnapshot_(Object.freeze(invalidado), opts);
}

function corePortalSyncFirestoreUsersFromPessoasV2_(opts) {
  opts = opts || {};
  var environment = corePortalFirestoreEnvironment_(opts);
  opts = Object.assign({}, opts, { ambiente: environment, environment: environment });
  var report = core_domainsV2NewReadReport_('PORTAL_FIRESTORE_USERS_SYNC');
  var pessoasData = core_domainsV2OpenPessoas_(report, { ambiente: environment });
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));

  var resumo = (pessoasData.PESSOAS_RESUMO_OPERACIONAL && pessoasData.PESSOAS_RESUMO_OPERACIONAL.records) || [];
  var baseById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [], 'ID_PESSOA');
  var uidByEmail = opts.uidByEmail || {};
  var uidByIdPessoa = opts.uidByIdPessoa || {};
  var limite = Math.max(0, Number(opts.limit || opts.limite || 0));
  var contadores = { criados: 0, atualizados: 0, ignorados: 0, erros: 0 };
  var erros = [];
  var processados = 0;

  for (var i = 0; i < resumo.length; i++) {
    if (limite && processados >= limite) break;
    var row = resumo[i] || {};
    var idPessoa = String(row.ID_PESSOA || '').trim();
    var base = baseById[idPessoa] || {};
    var email = corePortalNormalizeEmail_(row.EMAIL || base.EMAIL_PRINCIPAL || '');
    var portalAtivo = corePortalIsYes_(row.PORTAL_ATIVO);

    if (!idPessoa || !email || (!portalAtivo && opts.includeInactive !== true)) {
      contadores.ignorados++;
      continue;
    }

    var uid = String(uidByIdPessoa[idPessoa] || uidByEmail[email] || '').trim();
    if (!uid) {
      contadores.ignorados++;
      continue;
    }

    processados++;
    try {
      var resultado = corePortalSyncFirestoreUserByIdPessoa_(idPessoa, Object.assign({}, opts, { uid: uid }));
      if (resultado.ok && resultado.synced) {
        contadores.atualizados++;
      } else if (resultado.ok && resultado.code === 'DRY_RUN') {
        contadores.ignorados++;
      } else if (resultado.ok && resultado.code === 'FIRESTORE_PROJECT_ID_NAO_CONFIGURADO') {
        contadores.ignorados++;
      } else {
        contadores.erros++;
        erros.push({ indice: i + 1, code: resultado.code || 'ERRO_SYNC' });
      }
    } catch (erro) {
      contadores.erros++;
      erros.push({ indice: i + 1, code: 'EXCEPTION_SYNC' });
    }
  }

  Logger.log('GEAPA_CORE_PORTAL_FIRESTORE_USERS_SYNC ' + JSON.stringify({
    totalResumo: resumo.length,
    processados: processados,
    criados: contadores.criados,
    atualizados: contadores.atualizados,
    ignorados: contadores.ignorados,
    erros: contadores.erros
  }));

  return Object.freeze({
    ok: contadores.erros === 0,
    writer: opts.dryRun === true ? 'DRY_RUN' : 'APPS_SCRIPT_FIRESTORE_REST',
    schemaVersion: CORE_PORTAL_FIRESTORE_USER_SNAPSHOT_VERSION,
    contadores: Object.freeze(contadores),
    erros: Object.freeze(erros.slice(0, 20))
  });
}

/** Diagnostico read-only da cobertura do cache portalUsers, sem expor PII. */
function corePortalDiagnosticarFirestoreUsersDev_(opts) {
  opts = opts || {};
  var report = core_domainsV2NewReadReport_('PORTAL_FIRESTORE_USERS_DIAGNOSTICO');
  var pessoasData;
  try {
    pessoasData = core_domainsV2OpenPessoas_(report, { ambiente: 'DEV' });
  } catch (pessoasErr) {
    return Object.freeze({
      ok: false,
      errorCode: 'PESSOAS_V2_INDISPONIVEL',
      message: String(pessoasErr && pessoasErr.message || pessoasErr || '').slice(0, 300)
    });
  }
  if (report.totalErros) {
    return Object.freeze({
      ok: false,
      errorCode: 'PESSOAS_V2_INDISPONIVEL',
      totalErros: report.totalErros
    });
  }

  var resumo = (pessoasData.PESSOAS_RESUMO_OPERACIONAL && pessoasData.PESSOAS_RESUMO_OPERACIONAL.records) || [];
  var baseById = core_domainsV2IndexFirstBy_((pessoasData.PESSOAS_BASE && pessoasData.PESSOAS_BASE.records) || [], 'ID_PESSOA');
  var ambiente = '';
  try { ambiente = corePortalNormalizeToken_(core_getCurrentEnv_()); } catch (envErr) {}
  var includeEmails = opts.includeEmails === true &&
    ['DEV', 'DESENVOLVIMENTO', 'HOMOLOG', 'HOMOLOGACAO', 'PREVIEW'].indexOf(ambiente) >= 0;
  var detailLimit = Math.max(1, Math.min(200, Number(opts.limit || 50)));
  var authorized = {};
  var officialByEmail = {};
  var officialById = {};
  var totalMembrosAtivos = 0;
  var totalMembrosEfetivosAtivos = 0;
  var totalMembrosIngressantesAtivos = 0;
  var totalDiretoriaSecretariaAdminTecnico = 0;
  var totalOrientadoresAutorizados = 0;

  function addIndex(index, key, value) {
    if (!key) return;
    if (!index[key]) index[key] = [];
    index[key].push(value);
  }

  function displayEmail(email) {
    var value = String(email || '').trim().toLowerCase();
    if (value.indexOf('***@') > 0) return value;
    return includeEmails ? value : corePortalMaskEmailForDiagnostic_(value);
  }

  resumo.forEach(function(row, index) {
    row = row || {};
    var idPessoa = String(row.ID_PESSOA || '').trim();
    var base = baseById[idPessoa] || {};
    var email = corePortalNormalizeEmail_(row.EMAIL || row.EMAIL_PRINCIPAL || base.EMAIL_PRINCIPAL || '');
    var profileRaw = row.PERFIL_PORTAL_CALCULADO || row.PERFIL_PORTAL_BASE || row.PERFIL_PORTAL || '';
    var profiles = corePortalSplitProfiles_(profileRaw);
    var profile = profiles[0] || corePortalNormalizeToken_(profileRaw);
    var linkType = corePortalNormalizeToken_(row.TIPO_VINCULO_ATUAL || row.TIPO_VINCULO || '');
    var linkStatus = corePortalNormalizeToken_(row.STATUS_VINCULO_ATUAL || row.STATUS_VINCULO || '');
    var item = { idPessoa: idPessoa, email: email, profile: profile, linkType: linkType, line: index + 2 };
    addIndex(officialById, idPessoa, item);
    addIndex(officialByEmail, email, item);
    if (!idPessoa || !corePortalIsYes_(row.PORTAL_ATIVO)) return;
    if (!authorized[idPessoa]) authorized[idPessoa] = item;
    if ((linkStatus === 'ATIVO' || linkStatus === 'ATIVA') && (linkType === 'MEMBRO_INGRESSANTE' || linkType === 'MEMBRO_EFETIVO' || linkType === 'MEMBRO')) {
      totalMembrosAtivos++;
      if (linkType === 'MEMBRO_INGRESSANTE') totalMembrosIngressantesAtivos++;
      else totalMembrosEfetivosAtivos++;
    }
    if (profiles.some(function(value) { return ['DIRETORIA', 'SECRETARIA', 'ADMIN', 'ADMIN_TECNICO'].indexOf(value) >= 0; }) ||
      ['DIRETORIA', 'SECRETARIA', 'ADMIN', 'ADMIN_TECNICO'].indexOf(profile) >= 0) {
      totalDiretoriaSecretariaAdminTecnico++;
    }
    if (profiles.indexOf('ORIENTADOR') >= 0 || profile === 'ORIENTADOR' || linkType === 'ORIENTADOR') totalOrientadoresAutorizados++;
  });

  var firestoreDocuments = [];
  var pageToken = '';
  try {
    do {
      var page = coreFirestoreEnvironmentListDocuments_('portalUsers', {
        ambiente: 'DEV',
        pageSize: 500,
        pageToken: pageToken
      });
      if (!page || page.ok !== true) throw new Error(String(page && page.code || 'FIRESTORE_LIST_FALHOU'));
      firestoreDocuments = firestoreDocuments.concat(page.documents || []);
      pageToken = String(page.nextPageToken || '');
      if (firestoreDocuments.length > 5000) throw new Error('LIMITE_DEFENSIVO_EXCEDIDO');
    } while (pageToken);
  } catch (firestoreErr) {
    return Object.freeze({
      ok: false,
      errorCode: 'PORTAL_USERS_FIRESTORE_INDISPONIVEL',
      message: String(firestoreErr && firestoreErr.message || firestoreErr || '').slice(0, 300),
      totalUsuariosAutorizados: Object.keys(authorized).length
    });
  }

  var cachedById = {};
  var cachedByEmail = {};
  var totalAtivos = 0;
  var totalInativos = 0;
  firestoreDocuments.forEach(function(item) {
    var data = item && item.data || {};
    var idPessoa = String(data.idPessoa || '').trim();
    var email = corePortalNormalizeEmail_(data.emailNormalizado || data.email || '');
    addIndex(cachedById, idPessoa, item);
    addIndex(cachedByEmail, email, item);
    var active = data.ativo === true && data.podeAcessarPortal === true;
    if (typeof data.ativo === 'undefined' && typeof data.podeAcessarPortal === 'undefined') active = data.portalAtivo === true;
    if (active) totalAtivos++;
    else totalInativos++;
  });

  var covered = 0;
  var missingAuthorizedEmails = 0;
  var waitingFirstLogin = [];
  Object.keys(authorized).forEach(function(idPessoa) {
    var person = authorized[idPessoa];
    if ((cachedById[idPessoa] && cachedById[idPessoa].length) || (person.email && cachedByEmail[person.email] && cachedByEmail[person.email].length)) {
      covered++;
      return;
    }
    if (person.email) missingAuthorizedEmails++;
    if (waitingFirstLogin.length < detailLimit) {
      waitingFirstLogin.push(Object.freeze({
        idPessoa: idPessoa,
        email: displayEmail(person.email),
        status: 'AGUARDANDO_PRIMEIRO_LOGIN_FIREBASE'
      }));
    }
  });

  var orphanUsers = [];
  var orphanUsersCount = 0;
  firestoreDocuments.forEach(function(item) {
    var data = item && item.data || {};
    var idPessoa = String(data.idPessoa || '').trim();
    var email = corePortalNormalizeEmail_(data.emailNormalizado || data.email || '');
    if ((idPessoa && officialById[idPessoa]) || (email && officialByEmail[email])) return;
    orphanUsersCount++;
    if (orphanUsers.length < detailLimit) {
      orphanUsers.push(Object.freeze({
        uid: corePortalTruncateUidForLog_(item && item.id || data.uid || ''),
        idPessoa: idPessoa,
        email: displayEmail(email),
        status: 'SEM_CORRESPONDENCIA_BASE_OFICIAL'
      }));
    }
  });

  function collectDuplicates(index, source) {
    return Object.keys(index).filter(function(key) {
      return key && index[key].length > 1;
    }).slice(0, detailLimit).map(function(key) {
      return Object.freeze({
        origem: source,
        valor: key.indexOf('@') >= 0 ? displayEmail(key) : key,
        quantidade: index[key].length
      });
    });
  }

  var duplicateEmails = collectDuplicates(officialByEmail, 'PESSOAS_V2').concat(
    collectDuplicates(cachedByEmail, 'FIRESTORE_PORTAL_USERS')
  ).slice(0, detailLimit);
  var duplicateIds = collectDuplicates(officialById, 'PESSOAS_V2').concat(
    collectDuplicates(cachedById, 'FIRESTORE_PORTAL_USERS')
  ).slice(0, detailLimit);

  var recentErrors = [];
  var accessLogAvailable = false;
  try {
    var accessLogs = corePortalReadRegistryRecordsForEnv_('PORTAL_LOG_ACESSOS', {
      ambiente: 'DEV',
      skipBlankRows: true
    }) || [];
    accessLogAvailable = true;
    recentErrors = accessLogs.filter(function(row) {
      var action = corePortalNormalizeToken_(row.ACAO || '');
      var result = corePortalNormalizeToken_(row.RESULTADO || '');
      return action.indexOf('PORTAL_FIRESTORE_USER_PROVISION_') === 0 &&
        (result === 'ERROR' || result === 'ERRO' || result === 'DENY' || result === 'RECUSADO');
    }).slice(-Math.max(1, Math.min(20, Number(opts.errorLimit || 10)))).reverse().map(function(row) {
      return Object.freeze({
        timestamp: String(row.TIMESTAMP || '').slice(0, 40),
        resultado: corePortalNormalizeToken_(row.RESULTADO || ''),
        motivo: String(row.MOTIVO || '').slice(0, 160),
        origem: String(row.ORIGEM || '').slice(0, 80),
        email: displayEmail(row.EMAIL || ''),
        uid: corePortalTruncateUidForLog_(row.UID_FIREBASE || '')
      });
    });
  } catch (logErr) {}

  var totalAuthorized = Object.keys(authorized).length;
  var totalMissing = Math.max(0, totalAuthorized - covered);
  var recommendations = [];
  if (totalMissing) recommendations.push('Orientar os usuarios pendentes a realizar o primeiro login Firebase; nao criar UID manualmente.');
  if (orphanUsers.length) recommendations.push('Revisar portalUsers sem correspondencia oficial e manter esses documentos inativos.');
  if (duplicateEmails.length) recommendations.push('Corrigir duplicidades de e-mail antes de liberar read models privados.');
  if (duplicateIds.length) recommendations.push('Corrigir duplicidades de ID_PESSOA antes de liberar read models privados.');
  if (!recommendations.length) recommendations.push('Cobertura consistente; manter provisionamento automatico no login.');
  var structurallyReady = orphanUsersCount === 0 && !duplicateEmails.length && !duplicateIds.length;
  return Object.freeze({
    ok: true,
    modo: 'DEV_READ_ONLY',
    ambiente: ambiente || 'NAO_INFORMADO',
    includeEmails: includeEmails,
    prontoParaHomologacao: structurallyReady,
    coberturaCompleta: totalMissing === 0 && totalInativos === 0,
    status: structurallyReady ? 'PRONTO_PARA_HOMOLOGACAO' : 'REQUER_CORRECAO',
    totalPessoasAutorizadas: totalAuthorized,
    totalUsuariosAutorizados: totalAuthorized,
    totalMembrosAtivos: totalMembrosAtivos,
    totalMembrosEfetivosAtivos: totalMembrosEfetivosAtivos,
    totalMembrosIngressantesAtivos: totalMembrosIngressantesAtivos,
    totalDiretoriaSecretariaAdminTecnico: totalDiretoriaSecretariaAdminTecnico,
    totalOrientadoresAutorizados: totalOrientadoresAutorizados,
    totalDocumentosPortalUsers: firestoreDocuments.length,
    totalPortalUsersAtivos: totalAtivos,
    totalPortalUsersInativos: totalInativos,
    totalCachesValidos: totalAtivos,
    totalUsuariosComCacheConhecido: covered,
    totalUsuariosSemCacheConhecido: totalMissing,
    totalEmailsAutorizadosSemPortalUsers: missingAuthorizedEmails,
    totalPortalUsersSemCorrespondenciaOficial: orphanUsersCount,
    aguardandoPrimeiroLoginFirebase: Object.freeze(waitingFirstLogin),
    portalUsersSemCorrespondenciaOficial: Object.freeze(orphanUsers),
    duplicidadesPorEmail: Object.freeze(duplicateEmails),
    duplicidadesPorIdPessoa: Object.freeze(duplicateIds),
    ultimosErrosProvisionamento: Object.freeze(recentErrors),
    ultimosErrosPortalLoginFirebase: Object.freeze(recentErrors),
    recomendacoes: Object.freeze(recommendations),
    logPersistidoDisponivel: accessLogAvailable,
    observacoes: Object.freeze([
      'portalUsers/{uid} depende do UID emitido pelo Firebase Auth; o Core nao inventa UID por e-mail.',
      'Usuarios sem documento sao classificados como AGUARDANDO_PRIMEIRO_LOGIN_FIREBASE, nao como erro.',
      accessLogAvailable && recentErrors.length
        ? 'Erros persistidos foram resumidos com e-mail mascarado e UID truncado.'
        : 'Erros de provisionamento podem existir apenas no Execution Log e nao estar disponiveis neste relatorio.'
    ])
  });
}
function corePortalValidarAcesso_(idPessoa, permissaoOuPerfil, opts) {
  opts = opts || {};
  var wanted = String(permissaoOuPerfil || '').trim();
  if (!wanted) return Object.freeze({ ok: false, allowed: false, motivo: 'ALVO_AUSENTE' });
  var profileResult = corePortalCalcularPerfilEfetivo_(idPessoa, opts);
  var permissionResult = corePortalListarPermissoesEfetivas_(idPessoa, Object.assign({}, opts, {
    profileResult: profileResult
  }));
  var isPermission = wanted.indexOf(':') >= 0;
  var allowed = profileResult.portalAtivo && (
    isPermission
      ? permissionResult.permissoes.indexOf(corePortalNormalizePermission_(wanted)) >= 0
      : profileResult.perfisPortal.indexOf(corePortalNormalizeToken_(wanted)) >= 0
  );
  return Object.freeze({
    ok: true,
    allowed: allowed,
    idPessoa: String(idPessoa || '').trim(),
    alvo: wanted,
    tipoAlvo: isPermission ? 'PERMISSAO' : 'PERFIL',
    motivo: allowed ? 'ACESSO_PERMITIDO' : 'ACESSO_NEGADO'
  });
}

function corePortalGetMeuResumo_(email, opts) {
  var usuario = corePortalResolverUsuarioAtual_(email, opts || {});
  if (!usuario.ok) return usuario;
  var resumo = corePessoasGetOperationalSummary_(usuario.idPessoa) || {};
  return Object.freeze({
    ok: true,
    usuario: usuario,
    resumoOperacional: Object.freeze({
      tipoVinculoAtual: resumo.TIPO_VINCULO_ATUAL || '',
      statusVinculoAtual: resumo.STATUS_VINCULO_ATUAL || '',
      cargoFuncaoAtual: resumo.CARGO_FUNCAO_ATUAL || '',
      perfilPortalCalculado: resumo.PERFIL_PORTAL_CALCULADO || '',
      portalAtivo: resumo.PORTAL_ATIVO || '',
      tempoEfetivoNoGrupo: resumo.TEMPO_EFETIVO_NO_GRUPO || '',
      qtdSemestresNoGrupo: resumo.QTD_SEMESTRES_NO_GRUPO || '',
      qtdApresentacoesRealizadas: resumo.QTD_APRESENTACOES_REALIZADAS || '',
      periodoUltimaApresentacao: resumo.CICLO_ULTIMA_APRESENTACAO || resumo.PERIODO_ULTIMA_APRESENTACAO || '',
      frequenciaResumida: resumo.FREQUENCIA_RESUMIDA || '',
      pendenciasAbertas: resumo.PENDENCIAS_ABERTAS || '',
      flagJaFoiSuspenso: resumo.FLAG_JA_FOI_SUSPENSO || '',
      statusElegibilidadeDiretoria: resumo.STATUS_ELEGIBILIDADE_DIRETORIA || ''
    })
  });
}

function corePortalReadPresentationRecords_() {
  var records = [];
  try {
    var details = core_readSheetRecords_(core_getDomainSheet_('ATIVIDADES', 'PORTAL_ATIVIDADES_DETALHES', {}), { skipBlankRows: true }) || [];
    details.forEach(function(detail) {
      records = records.concat(corePortalPresentationRecordsFromActivityDetail_(detail));
    });
  } catch (err) {}
  return records;
}

function corePortalParseJsonArray_(value) {
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

function corePortalPresentationRecordsFromActivityDetail_(detail) {
  var presentations = corePortalParseJsonArray_(core_domainsV2LegacyValue_(detail, [
    'APRESENTACOES_PUBLICAS_JSON',
    'apresentacoesPublicas'
  ]));
  if (!presentations.length) return [];

  var base = {
    ID_ATIVIDADE: core_domainsV2LegacyValue_(detail, ['ID_ATIVIDADE', 'idAtividade']),
    DATA_ATIVIDADE: core_domainsV2LegacyValue_(detail, ['DATA_ATIVIDADE', 'dataAtividade', 'DATA']),
    PERIODO: core_domainsV2LegacyValue_(detail, ['PERIODO', 'ID_PERIODO', 'periodo']),
    VISIBILIDADE: core_domainsV2LegacyValue_(detail, ['VISIBILIDADE_PORTAL', 'visibilidadePortal', 'STATUS_PUBLICO', 'statusPublico'])
  };

  return presentations.filter(function(item) {
    return item && typeof item === 'object' && !Array.isArray(item);
  }).map(function(item) {
    var out = {};
    Object.keys(item).forEach(function(key) {
      out[key] = item[key];
    });
    Object.keys(base).forEach(function(key) {
      if (out[key] === undefined || out[key] === null || out[key] === '') out[key] = base[key];
    });
    return out;
  });
}

function corePortalPresentationDate_(record) {
  return core_domainsV2Date_(core_domainsV2LegacyValue_(record, [
    'DATA_ATIVIDADE',
    'DATA_APRESENTACAO',
    'DATA_REALIZADA',
    'DATA',
    'DATA_INICIO'
  ]));
}

function corePortalPresentationIsPublic_(record) {
  var visibility = corePortalNormalizeToken_(core_domainsV2LegacyValue_(record, [
    'VISIBILIDADE',
    'PORTAL_VISIBILIDADE',
    'LIBERADA_PORTAL',
    'PUBLICA',
    'PUBLICO'
  ]));
  return ['PUBLICA', 'PUBLICO', 'SIM', 'LIBERADA', 'PORTAL'].indexOf(visibility) >= 0;
}

function corePortalPresentationBelongsToUser_(record, usuario) {
  var idPessoa = String(core_domainsV2LegacyValue_(record, [
    'ID_PESSOA',
    'ID_PESSOA_APRESENTADOR',
    'ID_PESSOA_PRINCIPAL',
    'idPessoa',
    'idPessoaApresentador',
    'idPessoaPrincipal',
    'Id Pessoa'
  ]) || '').trim();
  var rga = String(core_domainsV2LegacyValue_(record, ['RGA', 'rga', 'rgaApresentador']) || '').trim();
  var email = corePortalNormalizeEmail_(core_domainsV2LegacyValue_(record, [
    'EMAIL',
    'EMAIL_APRESENTADOR',
    'Email',
    'E-mail',
    'email',
    'emailApresentador'
  ]));
  return (!!idPessoa && idPessoa === usuario.idPessoa) ||
    (!!rga && rga === usuario.rga) ||
    (!!email && email === usuario.email);
}

function corePortalSanitizePresentation_(record) {
  return Object.freeze({
    idAtividade: String(core_domainsV2LegacyValue_(record, ['ID_ATIVIDADE', 'ID_APRESENTACAO', 'ID', 'idAtividade', 'idApresentacao']) || '').trim(),
    titulo: String(core_domainsV2LegacyValue_(record, ['TITULO', 'TITULO_APRESENTACAO', 'TEMA', 'titulo', 'tituloApresentacao']) || '').trim(),
    dataAtividade: core_domainsV2LegacyValue_(record, ['DATA_ATIVIDADE', 'DATA_APRESENTACAO', 'DATA_REALIZADA', 'DATA', 'dataAtividade', 'dataApresentacao']),
    periodo: String(core_domainsV2LegacyValue_(record, ['PERIODO', 'CICLO', 'SEMESTRE', 'ID_PERIODO', 'periodo']) || '').trim(),
    eixo: String(core_domainsV2LegacyValue_(record, ['EIXO', 'EIXO_TEMATICO', 'EIXO_TEMATICO_PRINCIPAL', 'eixo', 'eixoTematicoPrincipal']) || '').trim(),
    status: String(core_domainsV2LegacyValue_(record, ['STATUS', 'STATUS_APRESENTACAO', 'SITUACAO', 'status', 'statusApresentacao']) || '').trim()
  });
}

function corePortalListarApresentacoesParaEgresso_(idPessoa, opts) {
  opts = opts || {};
  var bundle = opts.bundle || corePortalGetBundleById_(idPessoa);
  if (!bundle || !bundle.pessoa) {
    return Object.freeze({ ok: false, motivo: 'PESSOA_NAO_ENCONTRADA', apresentacoes: Object.freeze([]) });
  }
  var exitDate = corePortalGetEgressoExitDate_(bundle);
  if (!exitDate) {
    return Object.freeze({
      ok: false,
      motivo: 'DATA_FIM_EGRESSO_AUSENTE',
      status: 'PENDENTE',
      apresentacoes: Object.freeze([])
    });
  }
  var user = {
    idPessoa: String(bundle.pessoa.ID_PESSOA || '').trim(),
    email: corePortalNormalizeEmail_(bundle.pessoa.EMAIL_PRINCIPAL || ''),
    rga: String((bundle.membrosDetalhes && bundle.membrosDetalhes.RGA) || '').trim()
  };
  var out = corePortalReadPresentationRecords_().filter(function(record) {
    var date = corePortalPresentationDate_(record);
    return date && date <= exitDate && (corePortalPresentationIsPublic_(record) || corePortalPresentationBelongsToUser_(record, user));
  }).map(corePortalSanitizePresentation_);
  return Object.freeze({
    ok: true,
    regra: 'ATE_DATA_SAIDA',
    dataSaida: exitDate,
    apresentacoes: Object.freeze(out)
  });
}

function corePortalListarApresentacoesPermitidas_(email, options) {
  options = options || {};
  var usuario = corePortalResolverUsuarioAtual_(email, options);
  if (!usuario.ok || !usuario.portalAtivo) {
    return Object.freeze({
      ok: false,
      motivo: usuario.motivoBloqueio || 'USUARIO_SEM_ACESSO',
      apresentacoes: Object.freeze([])
    });
  }
  if (usuario.perfilPortalEfetivo === 'EGRESSO') {
    var egresso = corePortalListarApresentacoesParaEgresso_(usuario.idPessoa);
    return Object.freeze(Object.assign({ usuario: usuario }, egresso));
  }
  var canManage = usuario.permissoes.indexOf('apresentacoes:gerir') >= 0 || usuario.permissoes.indexOf('apresentacoes:ler') >= 0;
  var canOwn = usuario.permissoes.indexOf('apresentacoes:ver_propria') >= 0 || usuario.perfilPortalEfetivo === 'MEMBRO';
  var canPublic = usuario.permissoes.indexOf('apresentacoes:ver_publicas') >= 0 || ['VISITANTE', 'EXTERNO'].indexOf(usuario.perfilPortalEfetivo) >= 0;
  var out = corePortalReadPresentationRecords_().filter(function(record) {
    if (canManage) return true;
    if (canOwn && corePortalPresentationBelongsToUser_(record, usuario)) return true;
    return canPublic && corePortalPresentationIsPublic_(record);
  }).map(corePortalSanitizePresentation_);
  return Object.freeze({
    ok: true,
    usuario: usuario,
    apresentacoes: Object.freeze(out)
  });
}

function corePortalDiagnosticarPerfisEPermissoes_(opts) {
  opts = opts || {};
  var report = {
    ok: true,
    bloqueios: [],
    perfisObrigatoriosAusentes: [],
    sugestoesPerfis: [],
    perfisSemPermissoes: [],
    permissoesComPerfilInexistente: [],
    permissoesInativas: [],
    permissoesNaoRecomendadas: [],
    usuariosPortalAtivoSemPerfil: [],
    membrosAtivosSemPortalAcessar: [],
    egressosSemPerfilEgresso: [],
    conselhoSemCadastroPerfil: [],
    excecoesAdminInvalidas: [],
    avisos: [],
    recomendacoes: []
  };
  var profiles = [];
  var permissions = [];
  var rawPermissions = [];
  try {
    profiles = corePortalReadProfiles_();
  } catch (errProfiles) {
    report.ok = false;
    report.bloqueios.push('PORTAL_PERFIS_INDISPONIVEL');
    report.avisos.push('Nao foi possivel ler PORTAL_PERFIS: ' + (errProfiles && errProfiles.message ? errProfiles.message : errProfiles));
  }
  try {
    permissions = corePortalReadPermissions_();
    rawPermissions = core_readRecordsByKey_('PORTAL_PERMISSOES', { skipBlankRows: true });
  } catch (errPermissions) {
    report.ok = false;
    report.bloqueios.push('PORTAL_PERMISSOES_INDISPONIVEL');
    report.avisos.push('Nao foi possivel ler PORTAL_PERMISSOES: ' + (errPermissions && errPermissions.message ? errPermissions.message : errPermissions));
  }
  var profileMap = corePortalActiveProfileMap_(profiles);
  var permissionsByProfile = corePortalPermissionsByProfile_(permissions);

  CORE_PORTAL_V2_REQUIRED_PROFILES.forEach(function(profile) {
    if (!profileMap[profile]) {
      report.perfisObrigatoriosAusentes.push(profile);
      report.sugestoesPerfis.push({
        PERFIL_PORTAL: profile,
        NOME: profile,
        NIVEL: CORE_PORTAL_V2_PROFILE_LEVELS[profile] || '',
        ATIVO: 'SIM'
      });
    }
  });
  Object.keys(profileMap).forEach(function(profile) {
    if (!(permissionsByProfile[profile] || []).length) report.perfisSemPermissoes.push(profile);
  });
  rawPermissions.forEach(function(record) {
    var profile = corePortalNormalizeToken_(corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.profileKey));
    var permission = corePortalNormalizePermission_(corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.permission));
    if (!corePortalRecordIsActive_(record)) {
      if (profile || permission) report.permissoesInativas.push({ perfilPortal: profile, permissao: permission });
      return;
    }
    if (profile && !profileMap[profile]) {
      report.permissoesComPerfilInexistente.push({ perfilPortal: profile, permissao: permission });
    }
  });
  Object.keys(CORE_PORTAL_V2_MIN_PERMISSIONS).forEach(function(profile) {
    CORE_PORTAL_V2_MIN_PERMISSIONS[profile].forEach(function(permission) {
      if ((permissionsByProfile[profile] || []).indexOf(permission) < 0) {
        report.recomendacoes.push('Adicionar permissao ' + permission + ' ao perfil ' + profile + '.');
      }
    });
  });
  ['membros:ler', 'presencas:ler', 'atividades:gerir', 'logs:ler'].forEach(function(permission) {
    if ((permissionsByProfile.EGRESSO || []).indexOf(permission) >= 0) {
      report.permissoesNaoRecomendadas.push({
        perfilPortal: 'EGRESSO',
        permissao: permission,
        motivo: 'Egresso deve ter acesso limitado.'
      });
    }
  });
  (permissionsByProfile.CONSELHO || []).forEach(function(permission) {
    if (permission.indexOf(':gerir') >= 0 || permission === 'sistema:admin' || permission === 'logs:ler') {
      report.permissoesNaoRecomendadas.push({
        perfilPortal: 'CONSELHO',
        permissao: permission,
        motivo: 'Conselho nao deve receber permissao administrativa por padrao.'
      });
    }
  });
  (permissionsByProfile.MEMBRO_INGRESSANTE || []).forEach(function(permission) {
    if (CORE_PORTAL_V2_MIN_PERMISSIONS.MEMBRO_INGRESSANTE.indexOf(permission) < 0) {
      report.permissoesNaoRecomendadas.push({
        perfilPortal: 'MEMBRO_INGRESSANTE',
        permissao: permission,
        motivo: 'Membro ingressante deve receber somente acesso basico ao proprio perfil e situacao.'
      });
    }
  });

  var pessoasReport = core_domainsV2AuditNewReport_('PORTAL_DIAGNOSTICO_PESSOAS_V2');
  var pessoasData = core_domainsV2OpenPessoas_(pessoasReport);
  if (!pessoasReport.totalErros) {
    var resumo = (pessoasData.PESSOAS_RESUMO_OPERACIONAL && pessoasData.PESSOAS_RESUMO_OPERACIONAL.records) || [];
    resumo.forEach(function(record) {
      var idPessoa = String(record.ID_PESSOA || '').trim();
      var profile = corePortalNormalizeToken_(record.PERFIL_PORTAL_CALCULADO);
      var tipo = corePortalNormalizeToken_(record.TIPO_VINCULO_ATUAL);
      var active = corePortalIsYes_(record.PORTAL_ATIVO);
      if (active && !profile) report.usuariosPortalAtivoSemPerfil.push(idPessoa);
      if (active && profile && (permissionsByProfile[profile] || []).indexOf('portal:acessar') < 0) {
        report.membrosAtivosSemPortalAcessar.push({ idPessoa: idPessoa, perfilPortal: profile });
      }
      if ((tipo === 'EGRESSO' || tipo === 'EX_MEMBRO') && profile !== 'EGRESSO') {
        report.egressosSemPerfilEgresso.push({ idPessoa: idPessoa, perfilPortal: profile });
      }
    });
    ((pessoasData.PORTAL_ACESSOS_EXCECOES && pessoasData.PORTAL_ACESSOS_EXCECOES.records) || []).forEach(function(record) {
      if (corePortalNormalizeToken_(record.PERFIL_EXTRA) !== 'ADMIN') return;
      var activeException = corePortalIsExceptionActive_(record);
      if (!activeException || !String(record.JUSTIFICATIVA || '').trim()) {
        report.excecoesAdminInvalidas.push({
          idPessoa: record.ID_PESSOA || '',
          status: record.STATUS || '',
          dataFim: record.DATA_FIM || '',
          justificativa: record.JUSTIFICATIVA || ''
        });
      }
    });
  } else {
    report.avisos.push('Nao foi possivel abrir Pessoas v2 para diagnostico de usuarios.');
  }

  if (!profileMap.CONSELHO) report.conselhoSemCadastroPerfil.push('CONSELHO');
  var blocking = report.perfisObrigatoriosAusentes.length ||
    report.permissoesComPerfilInexistente.length ||
    report.usuariosPortalAtivoSemPerfil.length;
  report.ok = !blocking;
  return Object.freeze(report);
}

function corePortalDiagnosticarAcessoPortalDev_(opts) {
  opts = opts || {};
  var config = opts.config || corePortalReadConfig_(opts.configOpts || {});
  var modoAcesso = corePortalAccessMode_(config);
  var permissionsByProfile = corePortalPermissionsByProfile_(opts.permissions || corePortalReadPermissions_(opts.permissionsOpts || {}));
  var limit = Math.max(1, Number(opts.limit || 100));
  var bloqueados = [];
  var totais = {
    totalMembrosAtivos: 0,
    totalComEmail: 0,
    totalSemEmail: 0,
    totalComPortalAcessar: 0,
    totalForaListaTeste: 0
  };
  var records = [];
  var fonte = 'PESSOAS_V2_RESUMO_OPERACIONAL';

  try {
    records = core_readRecordsByKey_('PESSOAS_V2_RESUMO_OPERACIONAL', { skipBlankRows: true }) || [];
  } catch (err) {
    return Object.freeze({
      ok: false,
      modoAcesso: modoAcesso,
      errorCode: 'PESSOAS_RESUMO_OPERACIONAL_INDISPONIVEL',
      message: 'Nao foi possivel ler PESSOAS_V2_RESUMO_OPERACIONAL para diagnosticar acesso ao portal.'
    });
  }

  records.forEach(function(record) {
    var idPessoa = String(corePortalGetRecordValue_(record, ['ID_PESSOA']) || '').trim();
    var nome = String(corePortalGetRecordValue_(record, ['NOME_EXIBICAO', 'NOME_COMPLETO', 'NOME']) || '').trim();
    var email = corePortalNormalizeEmail_(corePortalGetRecordValue_(record, ['EMAIL_PRINCIPAL', 'EMAIL', 'E-MAIL']));
    var tipo = corePortalNormalizeToken_(corePortalGetRecordValue_(record, ['TIPO_VINCULO_ATUAL', 'TIPO_VINCULO']));
    var status = corePortalNormalizeToken_(corePortalGetRecordValue_(record, ['STATUS_VINCULO_ATUAL', 'STATUS_VINCULO', 'STATUS']));
    var profile = corePortalNormalizeToken_(corePortalGetRecordValue_(record, ['PERFIL_PORTAL_CALCULADO', 'PERFIL_PORTAL_BASE', 'PERFIL_PORTAL']));
    var activeMember = ['MEMBRO_INGRESSANTE', 'MEMBRO_EFETIVO', 'MEMBRO'].indexOf(tipo) >= 0 && ['ATIVO', 'ATIVA'].indexOf(status) >= 0;
    if (!activeMember) return;

    totais.totalMembrosAtivos++;
    if (email) totais.totalComEmail++;
    else {
      totais.totalSemEmail++;
      if (bloqueados.length < limit) {
        bloqueados.push({
          idPessoa: idPessoa,
          nome: nome,
          email: '',
          motivo: 'Membro ativo sem e-mail cadastrado'
        });
      }
    }

    if (!profile) {
      if (bloqueados.length < limit) {
        bloqueados.push({
          idPessoa: idPessoa,
          nome: nome,
          email: email,
          motivo: 'Membro ativo sem perfil portal calculado'
        });
      }
      return;
    }

    if ((permissionsByProfile[profile] || []).indexOf('portal:acessar') >= 0) {
      totais.totalComPortalAcessar++;
    } else if (bloqueados.length < limit) {
      bloqueados.push({
        idPessoa: idPessoa,
        nome: nome,
        email: email,
        motivo: 'Perfil portal sem permissao portal:acessar'
      });
    }

    if (modoAcesso === 'TESTE' && email && !corePortalIsEmailInTestList_(email, config)) {
      totais.totalForaListaTeste++;
      if (bloqueados.length < limit) {
        bloqueados.push({
          idPessoa: idPessoa,
          nome: nome,
          email: email,
          motivo: 'Membro ativo fora de PORTAL_EMAILS_TESTE no modo TESTE'
        });
      }
    }
  });

  return Object.freeze({
    ok: true,
    modoAcesso: modoAcesso,
    authAllowVisitante: corePortalIsYes_(config.AUTH_ALLOW_VISITANTE),
    portalBlockInactiveMembers: !Object.prototype.hasOwnProperty.call(config, 'PORTAL_BLOCK_INACTIVE_MEMBERS') ||
      corePortalIsYes_(config.PORTAL_BLOCK_INACTIVE_MEMBERS),
    fonte: fonte,
    totalMembrosAtivos: totais.totalMembrosAtivos,
    totalComEmail: totais.totalComEmail,
    totalSemEmail: totais.totalSemEmail,
    totalComPortalAcessar: totais.totalComPortalAcessar,
    totalForaListaTeste: totais.totalForaListaTeste,
    bloqueados: Object.freeze(bloqueados)
  });
}

function corePrepararPortalParaV2_(opts) {
  opts = opts || {};
  var requiredFunctions = [
    'corePortalResolverUsuarioAtual',
    'corePortalValidarAcesso',
    'corePortalCalcularPerfilEfetivo',
    'corePortalListarPermissoesEfetivas',
    'corePortalGetMeuResumo'
  ];
  var functionMap = {
    corePortalResolverUsuarioAtual: typeof corePortalResolverUsuarioAtual === 'function',
    corePortalValidarAcesso: typeof corePortalValidarAcesso === 'function',
    corePortalCalcularPerfilEfetivo: typeof corePortalCalcularPerfilEfetivo === 'function',
    corePortalListarPermissoesEfetivas: typeof corePortalListarPermissoesEfetivas === 'function',
    corePortalGetMeuResumo: typeof corePortalGetMeuResumo === 'function'
  };
  var available = [];
  var missing = [];
  requiredFunctions.forEach(function(name) {
    if (functionMap[name]) available.push(name);
    else missing.push(name);
  });
  var diagnostic = corePortalDiagnosticarPerfisEPermissoes_(opts);
  var bloqueios = [];
  var avisos = [];
  if (missing.length) bloqueios.push('FUNCOES_PUBLICAS_FALTANTES');
  (diagnostic.bloqueios || []).forEach(function(code) {
    if (bloqueios.indexOf(code) < 0) bloqueios.push(code);
  });
  if (diagnostic.perfisObrigatoriosAusentes.length) bloqueios.push('PERFIS_OBRIGATORIOS_AUSENTES');
  if (diagnostic.permissoesComPerfilInexistente.length) bloqueios.push('PERMISSOES_COM_PERFIL_INEXISTENTE');
  if (diagnostic.usuariosPortalAtivoSemPerfil.length) bloqueios.push('USUARIOS_PORTAL_ATIVO_SEM_PERFIL');
  if (diagnostic.perfisSemPermissoes.length) avisos.push('PERFIS_SEM_PERMISSOES');
  if (diagnostic.excecoesAdminInvalidas.length) avisos.push('EXCECOES_ADMIN_REVISAR');
  if (diagnostic.permissoesNaoRecomendadas.length) avisos.push('PERMISSOES_NAO_RECOMENDADAS');
  if (!diagnostic.recomendacoes.some(function(text) { return text.indexOf('apresentacoes:ver_ate_saida') >= 0 && text.indexOf('EGRESSO') >= 0; })) {
    avisos.push('EGRESSO_APRESENTACOES_ATE_SAIDA_VERIFICADO_OU_JA_CONFIGURADO');
  }
  return Object.freeze({
    status: bloqueios.length ? 'BLOQUEADO' : (avisos.length ? 'PARCIAL' : 'PRONTO'),
    bloqueios: Object.freeze(bloqueios),
    avisos: Object.freeze(avisos),
    funcoesDisponiveis: Object.freeze(available),
    funcoesFaltantes: Object.freeze(missing),
    perfisInvalidos: Object.freeze(diagnostic.perfisObrigatoriosAusentes.concat(
      diagnostic.permissoesComPerfilInexistente.map(function(item) { return item.perfilPortal; })
    )),
    permissoesInvalidas: Object.freeze(diagnostic.permissoesComPerfilInexistente),
    recomendacoes: Object.freeze(diagnostic.recomendacoes),
    diagnostico: diagnostic
  });
}

function corePortalResolveSheetByRegistryOrName_(sourceCfg) {
  for (var i = 0; i < sourceCfg.registryKeys.length; i++) {
    try {
      return core_getSheetByKey_(sourceCfg.registryKeys[i]);
    } catch (err) {}
  }

  var raw = core_getRegistryRaw_();
  var currentEnv = core_getCurrentEnv_();
  var wantedNames = sourceCfg.sheetNames.map(corePortalNormalizeToken_);
  var found = null;

  Object.keys(raw).some(function(key) {
    var envMap = raw[key] || {};
    var entry = envMap[currentEnv] || envMap.PROD || null;
    if (!entry || !entry.ativo) return false;
    if (wantedNames.indexOf(corePortalNormalizeToken_(entry.sheet)) < 0) return false;
    found = entry;
    return true;
  });

  return found ? core_getSheetById_(found.id, found.sheet) : null;
}

function corePortalGetPersonSources_(opts) {
  opts = opts || {};
  if (Array.isArray(opts.sources)) return opts.sources;

  var sources = [];
  var currentSheet = corePortalResolveSheetByRegistryOrName_(CORE_PORTAL_ACCESS_CFG.sources.current);
  if (currentSheet) {
    sources.push(Object.freeze({
      type: CORE_PORTAL_ACCESS_CFG.sources.current.type,
      sourceSheet: typeof currentSheet.getName === 'function' ? currentSheet.getName() : 'Membros Atuais',
      sheet: currentSheet
    }));
  }

  if (opts.includeWaiting === true) {
    var waitingSheet = corePortalResolveSheetByRegistryOrName_(CORE_PORTAL_ACCESS_CFG.sources.waiting);
    if (waitingSheet) {
      sources.push(Object.freeze({
        type: CORE_PORTAL_ACCESS_CFG.sources.waiting.type,
        sourceSheet: typeof waitingSheet.getName === 'function' ? waitingSheet.getName() : 'Membros em Espera',
        sheet: waitingSheet
      }));
    }
  }

  if (opts.includeFormer === true) {
    var formerSheet = corePortalResolveSheetByRegistryOrName_(CORE_PORTAL_ACCESS_CFG.sources.former);
    if (formerSheet) {
      sources.push(Object.freeze({
        type: CORE_PORTAL_ACCESS_CFG.sources.former.type,
        sourceSheet: typeof formerSheet.getName === 'function' ? formerSheet.getName() : 'Ex-Membros',
        sheet: formerSheet
      }));
    }
  }

  return sources;
}

function corePortalNormalizePersonRecord_(record, source) {
  var email = corePortalNormalizeEmail_(
    corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.email)
  );

  return Object.freeze({
    found: !!email,
    sourceSheet: source.sourceSheet || '',
    sourceType: source.type || '',
    nome: String(corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.name) || '').trim(),
    rga: String(corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.rga) || '').trim(),
    email: email,
    status: String(corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.status) || '').trim(),
    portalAtivo: String(corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.portalActive) || '').trim(),
    perfilPortal: corePortalNormalizeToken_(
      corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.portalProfile)
    ),
    rawRecord: record
  });
}

function corePortalFindPersonByEmail_(email, opts) {
  opts = opts || {};
  var targetEmail = corePortalNormalizeEmail_(email);
  if (!targetEmail) {
    return Object.freeze({ found: false, reason: 'EMAIL_INVALIDO' });
  }

  var sources = corePortalGetPersonSources_(opts);

  for (var i = 0; i < sources.length; i++) {
    var source = sources[i];
    if (!source || !source.sheet) continue;

    var records = core_readSheetRecords_(source.sheet, {
      skipBlankRows: true
    });

    for (var j = 0; j < records.length; j++) {
      var recordEmail = corePortalNormalizeEmail_(
        corePortalGetRecordValue_(records[j], CORE_PORTAL_ACCESS_CFG.headers.email)
      );
      if (!recordEmail || recordEmail !== targetEmail) continue;
      return corePortalNormalizePersonRecord_(records[j], source);
    }
  }

  return Object.freeze({
    found: false,
    reason: 'PESSOA_NAO_ENCONTRADA'
  });
}

function corePortalResolveProfile_(person, config, activeProfiles) {
  var explicitProfile = corePortalNormalizeToken_(person && person.perfilPortal);
  var profile = explicitProfile;
  var reason = '';

  if (!profile) {
    var defaultProfile = corePortalNormalizeToken_(config.PORTAL_DEFAULT_PROFILE || '');
    if (defaultProfile && defaultProfile !== 'ADMIN') {
      profile = defaultProfile;
      reason = 'PERFIL_PADRAO_CONFIG';
    } else if (person && person.sourceType === CORE_PORTAL_ACCESS_CFG.sources.current.type) {
      profile = 'MEMBRO';
      reason = 'PERFIL_PADRAO_MEMBRO_ATUAL';
    }
  }

  if (!profile) {
    return Object.freeze({ ok: false, profile: '', reason: 'PERFIL_PORTAL_AUSENTE' });
  }

  if (profile === 'VISITANTE' && !corePortalIsYes_(config.AUTH_ALLOW_VISITANTE)) {
    return Object.freeze({ ok: false, profile: profile, reason: 'VISITANTE_NAO_PERMITIDO' });
  }

  if (profile === 'ADMIN' && !explicitProfile) {
    return Object.freeze({ ok: false, profile: profile, reason: 'ADMIN_NAO_PODE_SER_PADRAO' });
  }

  var exists = activeProfiles.some(function(item) {
    return corePortalNormalizeToken_(item.perfilPortal) === profile;
  });

  if (!exists) {
    return Object.freeze({ ok: false, profile: profile, reason: 'PERFIL_PORTAL_INVALIDO_OU_INATIVO' });
  }

  return Object.freeze({
    ok: true,
    profile: profile,
    reason: reason || 'PERFIL_RESOLVIDO'
  });
}

function corePortalBuildAuthorizationResult_(authorized, data) {
  data = data || {};
  return Object.freeze({
    authorized: authorized === true,
    authMode: 'EMAIL_LEGACY',
    modoAcesso: String(data.modoAcesso || '').trim(),
    email: String(data.email || '').trim(),
    nome: String(data.nome || '').trim(),
    rga: String(data.rga || '').trim(),
    sourceSheet: String(data.sourceSheet || '').trim(),
    status: String(data.status || '').trim(),
    portalAtivo: String(data.portalAtivo || '').trim(),
    perfilPortal: String(data.perfilPortal || '').trim(),
    permissions: Object.freeze((data.permissions || []).slice()),
    reason: String(data.reason || '').trim(),
    message: data.message || (authorized === true ? '' : corePortalAccessDeniedMessage_()),
    timestamp: data.timestamp || new Date()
  });
}

function corePortalAuthorizeEmail_(email, opts) {
  opts = opts || {};
  var normalizedEmail = corePortalNormalizeEmail_(email);
  if (!normalizedEmail) {
    return corePortalBuildAuthorizationResult_(false, {
      email: '',
      reason: 'EMAIL_INVALIDO'
    });
  }

  var config = opts.config || corePortalReadConfig_(opts.configOpts || {});
  var mode = corePortalAccessMode_(config);
  if (mode === 'TESTE' && !corePortalIsEmailInTestList_(normalizedEmail, config)) {
    return corePortalBuildAuthorizationResult_(false, {
      email: normalizedEmail,
      modoAcesso: mode,
      reason: 'EMAIL_FORA_PORTAL_EMAILS_TESTE'
    });
  }
  var includeWaiting = opts.includeWaiting === true || corePortalIsYes_(config.PORTAL_ALLOW_WAITING);
  var includeFormer = true; // busca ex-membro para bloquear com motivo claro por padrao
  var person = corePortalFindPersonByEmail_(normalizedEmail, {
    sources: opts.sources,
    includeWaiting: true,
    includeFormer: includeFormer
  });

  if (!person || !person.found) {
    return corePortalBuildAuthorizationResult_(false, {
      email: normalizedEmail,
      modoAcesso: mode,
      reason: 'PESSOA_NAO_ENCONTRADA'
    });
  }

  if (person.sourceType === CORE_PORTAL_ACCESS_CFG.sources.former.type && opts.includeFormer !== true) {
    return corePortalBuildAuthorizationResult_(false, {
      email: normalizedEmail,
      modoAcesso: mode,
      nome: person.nome,
      rga: person.rga,
      sourceSheet: person.sourceSheet,
      status: person.status,
      portalAtivo: person.portalAtivo,
      perfilPortal: person.perfilPortal,
      reason: 'EX_MEMBRO_NAO_AUTORIZADO'
    });
  }

  if (person.sourceType === CORE_PORTAL_ACCESS_CFG.sources.waiting.type && !includeWaiting) {
    return corePortalBuildAuthorizationResult_(false, {
      email: normalizedEmail,
      modoAcesso: mode,
      nome: person.nome,
      rga: person.rga,
      sourceSheet: person.sourceSheet,
      status: person.status,
      portalAtivo: person.portalAtivo,
      perfilPortal: person.perfilPortal,
      reason: 'MEMBRO_EM_ESPERA_NAO_AUTORIZADO'
    });
  }

  if (corePortalIsExplicitNo_(person.portalAtivo)) {
    return corePortalBuildAuthorizationResult_(false, {
      email: normalizedEmail,
      modoAcesso: mode,
      nome: person.nome,
      rga: person.rga,
      sourceSheet: person.sourceSheet,
      status: person.status,
      portalAtivo: person.portalAtivo,
      perfilPortal: person.perfilPortal,
      reason: 'PORTAL_ATIVO_NAO'
    });
  }

  if (!corePortalIsYes_(person.portalAtivo)) {
    var blockBlank = !Object.prototype.hasOwnProperty.call(config, 'PORTAL_BLOCK_INACTIVE_MEMBERS') ||
      corePortalIsYes_(config.PORTAL_BLOCK_INACTIVE_MEMBERS);
    if (blockBlank) {
      return corePortalBuildAuthorizationResult_(false, {
        email: normalizedEmail,
        modoAcesso: mode,
        nome: person.nome,
        rga: person.rga,
        sourceSheet: person.sourceSheet,
        status: person.status,
        portalAtivo: person.portalAtivo,
        perfilPortal: person.perfilPortal,
        reason: 'PORTAL_ATIVO_AUSENTE'
      });
    }
  }

  var profiles = opts.profiles || corePortalReadProfiles_(opts.profilesOpts || {});
  var resolvedProfile = corePortalResolveProfile_(person, config, profiles);
  if (!resolvedProfile.ok) {
    return corePortalBuildAuthorizationResult_(false, {
      email: normalizedEmail,
      modoAcesso: mode,
      nome: person.nome,
      rga: person.rga,
      sourceSheet: person.sourceSheet,
      status: person.status,
      portalAtivo: person.portalAtivo,
      perfilPortal: resolvedProfile.profile || person.perfilPortal,
      reason: resolvedProfile.reason
    });
  }

  var permissions = corePortalBuildPermissionsForProfile_(resolvedProfile.profile, {
    permissions: opts.permissions || null,
    permissionsOpts: opts.permissionsOpts || {}
  });
  var hasAccess = permissions.indexOf('portal:acessar') >= 0;

  return corePortalBuildAuthorizationResult_(hasAccess, {
    email: normalizedEmail,
    modoAcesso: mode,
    nome: person.nome,
    rga: person.rga,
    sourceSheet: person.sourceSheet,
    status: person.status,
    portalAtivo: person.portalAtivo,
    perfilPortal: resolvedProfile.profile,
    permissions: permissions,
    reason: hasAccess ? 'AUTORIZADO' : 'PERMISSAO_PORTAL_ACESSAR_AUSENTE'
  });
}

function corePortalHasPermission_(sessionOrEmail, permission, opts) {
  opts = opts || {};
  var wanted = String(permission || '').trim();
  var session = typeof sessionOrEmail === 'string'
    ? corePortalAuthorizeEmail_(sessionOrEmail, opts)
    : sessionOrEmail;
  var allowed = !!(session && session.authorized && wanted && session.permissions && session.permissions.indexOf(wanted) >= 0);

  if (opts.detailed === true) {
    return Object.freeze({
      allowed: allowed,
      permission: wanted,
      reason: allowed ? 'PERMISSAO_CONCEDIDA' : 'PERMISSAO_NEGADA'
    });
  }

  return allowed;
}

function corePortalSanitizeLogPayload_(payload) {
  payload = payload || {};
  var rawEmail = String(payload.email || payload.EMAIL || '').trim().toLowerCase();
  var safeEmail = rawEmail.indexOf('***@') > 0 ? rawEmail : corePortalNormalizeEmail_(rawEmail);
  return {
    TIMESTAMP: payload.timestamp || payload.TIMESTAMP || new Date(),
    EMAIL: safeEmail,
    UID_FIREBASE: String(payload.uidFirebase || payload.UID_FIREBASE || '').trim(),
    NOME: String(payload.nome || payload.NOME || '').trim(),
    PERFIL_PORTAL: corePortalNormalizeToken_(payload.perfilPortal || payload.PERFIL_PORTAL || ''),
    ACAO: corePortalNormalizeToken_(payload.acao || payload.ACAO || ''),
    RESULTADO: corePortalNormalizeToken_(payload.resultado || payload.RESULTADO || ''),
    MOTIVO: String(payload.motivo || payload.MOTIVO || '').trim().slice(0, 500),
    ORIGEM: String(payload.origem || payload.ORIGEM || '').trim().slice(0, 120),
    USER_AGENT: String(payload.userAgent || payload.USER_AGENT || '').trim().slice(0, 500),
    OBS: String(payload.obs || payload.OBS || '').trim().slice(0, 500)
  };
}

function corePortalAppendAccessLogToSheet_(sheet, payload) {
  return core_appendObjectByHeaders_(sheet, corePortalSanitizeLogPayload_(payload), {
    headerRow: 1
  });
}

function corePortalResolveRegistrySheetForEnv_(registryKey, opts) {
  var options = opts || {};
  var environment = corePortalResolveEnvironment_(options);
  var state = core_domainRegistryEntry_(registryKey, environment, options);
  if (!state.available) {
    throw core_domainResolverError_(
      'PORTAL_REGISTRY_KEY_INDISPONIVEL',
      'Registry key "' + registryKey + '" indisponivel em ' + environment + '.',
      { registryKey: registryKey, ambiente: environment, motivo: state.reason }
    );
  }

  var entry = state.entry || {};
  if (!entry.ativo || !String(entry.id || '').trim() || !String(entry.sheet || '').trim()) {
    throw core_domainResolverError_(
      'PORTAL_REGISTRY_ENTRY_INVALIDA',
      'Registry key "' + registryKey + '" invalida em ' + environment + '.',
      { registryKey: registryKey, ambiente: environment }
    );
  }
  return core_getSheetById_(String(entry.id).trim(), String(entry.sheet).trim());
}

function corePortalAppendAccessLog_(payload, opts) {
  var options = opts || {};
  var environment = corePortalResolveEnvironment_(options);
  var sheet = corePortalResolveRegistrySheetForEnv_('PORTAL_LOG_ACESSOS', {
    ambiente: environment
  });
  corePortalAppendAccessLogToSheet_(sheet, payload || {});
  return Object.freeze({
    ok: true,
    logged: true,
    ambiente: environment
  });
}

function corePortalDiagnosticsCheckSheet_(key, requiredHeaders) {
  var result = {
    key: key,
    ok: false,
    registry: null,
    sheetName: '',
    missingHeaders: [],
    error: ''
  };

  try {
    var meta = core_getRegistryMetaByKey_(key);
    result.registry = {
      key: meta.key,
      id: meta.id,
      sheet: meta.sheet,
      ativo: meta.ativo,
      ambiente: meta.ambiente,
      lineNo: meta.lineNo
    };

    var sheet = core_getSheetByKey_(key);
    result.sheetName = typeof sheet.getName === 'function' ? sheet.getName() : meta.sheet;
    var headerMap = core_headerMap_(sheet, 1);

    result.missingHeaders = (requiredHeaders || []).filter(function(headerName) {
      return !core_getCol_(headerMap, headerName);
    });
    result.ok = result.registry.ativo === true && result.missingHeaders.length === 0;
  } catch (err) {
    result.error = String(err && err.message || err || '');
  }

  return Object.freeze(result);
}

function corePortalDiagnosticsCheckMemberSheet_(sourceCfg) {
  var result = {
    type: sourceCfg.type,
    ok: false,
    sheetName: '',
    missingPortalColumns: [],
    error: ''
  };

  try {
    var sheet = corePortalResolveSheetByRegistryOrName_(sourceCfg);
    if (!sheet) throw new Error('Aba nao encontrada pelo Registry.');
    result.sheetName = typeof sheet.getName === 'function' ? sheet.getName() : sourceCfg.sheetNames[0];
    var headerMap = core_headerMap_(sheet, 1);
    result.missingPortalColumns = ['PORTAL_ATIVO', 'PERFIL_PORTAL', 'PORTAL_OBS'].filter(function(headerName) {
      return !core_getCol_(headerMap, headerName);
    });
    result.ok = true;
  } catch (err) {
    result.error = String(err && err.message || err || '');
  }

  return Object.freeze(result);
}

function corePortalDiagnostics_() {
  var requiredHeadersByKey = {
    PORTAL_PERFIS: ['PERFIL_PORTAL', 'ATIVO'],
    PORTAL_PERMISSOES: ['PERFIL_PORTAL', 'PERMISSAO', 'ATIVO'],
    PORTAL_CONFIG: ['CHAVE', 'VALOR', 'ATIVO'],
    PORTAL_LOG_ACESSOS: CORE_PORTAL_ACCESS_CFG.accessLogHeaders
  };
  var registry = CORE_PORTAL_ACCESS_CFG.requiredRegistryKeys.map(function(key) {
    return corePortalDiagnosticsCheckSheet_(key, requiredHeadersByKey[key] || ['ATIVO']);
  });
  var profiles = [];
  var permissions = [];
  var config = {};

  try {
    profiles = corePortalReadProfiles_();
  } catch (errProfiles) {}

  try {
    permissions = corePortalReadPermissions_();
  } catch (errPermissions) {}

  try {
    config = corePortalReadConfig_();
  } catch (errConfig) {}

  var memberSheets = [
    CORE_PORTAL_ACCESS_CFG.sources.current,
    CORE_PORTAL_ACCESS_CFG.sources.waiting,
    CORE_PORTAL_ACCESS_CFG.sources.former
  ].map(corePortalDiagnosticsCheckMemberSheet_);

  return Object.freeze({
    ok: registry.every(function(item) { return item.ok; }) && profiles.length > 0 && permissions.length > 0,
    registry: Object.freeze(registry),
    profiles: Object.freeze({
      activeCount: profiles.length,
      items: Object.freeze(profiles)
    }),
    permissions: Object.freeze({
      activeCount: permissions.length
    }),
    config: Object.freeze({
      activeCount: Object.keys(config).length,
      keys: Object.freeze(Object.keys(config).sort())
    }),
    memberSheets: Object.freeze(memberSheets)
  });
}
