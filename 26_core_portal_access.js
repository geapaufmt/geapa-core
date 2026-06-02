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
  var records = opts.records || core_readRecordsByKey_('PORTAL_PERFIS', {
    skipBlankRows: true
  });
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
  var records = opts.records || core_readRecordsByKey_('PORTAL_PERMISSOES', {
    skipBlankRows: true
  });
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

function corePortalReadConfig_(opts) {
  opts = opts || {};
  var records = opts.records || core_readRecordsByKey_('PORTAL_CONFIG', {
    skipBlankRows: true
  });
  var out = {};

  records.forEach(function(record) {
    if (!corePortalRecordIsActive_(record)) return;

    var key = corePortalNormalizeToken_(
      corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.configKey)
    );
    if (!key) return;

    out[key] = String(
      corePortalGetRecordValue_(record, CORE_PORTAL_ACCESS_CFG.headers.configValue) || ''
    ).trim();
  });

  return Object.freeze(out);
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
    email: String(data.email || '').trim(),
    nome: String(data.nome || '').trim(),
    rga: String(data.rga || '').trim(),
    sourceSheet: String(data.sourceSheet || '').trim(),
    status: String(data.status || '').trim(),
    portalAtivo: String(data.portalAtivo || '').trim(),
    perfilPortal: String(data.perfilPortal || '').trim(),
    permissions: Object.freeze((data.permissions || []).slice()),
    reason: String(data.reason || '').trim(),
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
      reason: 'PESSOA_NAO_ENCONTRADA'
    });
  }

  if (person.sourceType === CORE_PORTAL_ACCESS_CFG.sources.former.type && opts.includeFormer !== true) {
    return corePortalBuildAuthorizationResult_(false, {
      email: normalizedEmail,
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
  return {
    TIMESTAMP: payload.timestamp || payload.TIMESTAMP || new Date(),
    EMAIL: corePortalNormalizeEmail_(payload.email || payload.EMAIL || ''),
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

function corePortalAppendAccessLog_(payload) {
  var sheet = core_getSheetByKey_('PORTAL_LOG_ACESSOS');
  corePortalAppendAccessLogToSheet_(sheet, payload || {});
  return Object.freeze({
    ok: true,
    logged: true
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
