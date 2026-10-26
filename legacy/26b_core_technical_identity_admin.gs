/**
 * Administracao estreita da unica persona tecnica CODEX_E2E em DEV.
 *
 * Estas funcoes nao formam uma API administrativa generica: campos, vinculo,
 * finalidade, ambiente, perfil e permissoes sao constantes deste contrato.
 */

var CORE_TECH_ADMIN_CHANNEL = 'MYSELF_EXECUTION_API';
var CORE_TECH_ADMIN_PURPOSE = 'CODEX_E2E';
var CORE_TECH_ADMIN_PROFILE = 'DIRETORIA';
var CORE_TECH_ADMIN_MAX_LIFETIME_MS = 30 * 24 * 60 * 60 * 1000;
var CORE_TECH_ADMIN_CONFIRM_CREATE = 'CRIAR_IDENTIDADE_TECNICA_DIRETORIA_DEV';
var CORE_TECH_ADMIN_CONFIRM_END = 'ENCERRAR_IDENTIDADE_TECNICA_DIRETORIA_DEV';
var CORE_TECH_ADMIN_CONFIRM_PERMISSIONS = 'APLICAR_TETO_IDENTIDADE_TECNICA_DIRETORIA_DEV';

function coreTechAdminError_(code) {
  var error = new Error(code);
  error.code = code;
  return error;
}

function coreTechAdminAssertDev_(request, requireIdentifier) {
  var input = request || {};
  // A decisao de ambiente ocorre antes de qualquer acesso a Properties/Sheets.
  if (String(input.environment || '').trim().toUpperCase() !== 'DEV') {
    throw coreTechAdminError_('IDENTIDADE_TECNICA_ADMIN_PROD_BLOQUEADO');
  }
  if (String(input.channel || '').trim().toUpperCase() !== CORE_TECH_ADMIN_CHANNEL) {
    throw coreTechAdminError_('IDENTIDADE_TECNICA_ADMIN_CANAL_INVALIDO');
  }
  var active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  var effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (!active || !effective || active !== effective) {
    throw coreTechAdminError_('IDENTIDADE_TECNICA_ADMIN_MYSELF_OBRIGATORIO');
  }
  var identifier = String(input.identifier || '').trim().toLowerCase();
  if (requireIdentifier !== false && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(identifier)) {
    throw coreTechAdminError_('IDENTIDADE_TECNICA_IDENTIFICADOR_INVALIDO');
  }
  return { identifier: identifier };
}

function coreTechAdminHash_(value) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value || ''), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '').slice(0, 16);
}

function coreTechAdminResponse_(payload) {
  return Object.freeze(Object.assign({ ok: true, environment: 'DEV' }, payload || {}));
}

function coreTechAdminSafe_(fn) {
  try {
    return fn();
  } catch (error) {
    var code = String(error && (error.code || error.message) || 'IDENTIDADE_TECNICA_ADMIN_FALHOU')
      .toUpperCase().replace(/[^A-Z0-9_:-]/g, '_').slice(0, 120);
    return Object.freeze({ ok: false, environment: 'DEV', code: code });
  }
}

function coreTechAdminOpenData_() {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:62:coreTechAdminOpenData_");try {

  var logical = ['BASE', 'IDENTIFICADORES', 'VINCULOS', 'ACESSOS_EXCECOES'];
  var out = {};
  logical.forEach(function(name) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:65:<anonymous>");try {

    var resolved = core_getDomainSheet_('PESSOAS', name, { ambiente: 'DEV', includeResolution: true });
    var sheet = resolved.sheet;
    out[name] = {
      sheet: sheet,
      headers: sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0] : [],
      records: core_readSheetRecords_(sheet, { skipBlankRows: true })
    };
  
} finally { f4MemberLegacyFence_().end(__gFenceToken); }
});
  return out;

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechAdminRequiredHeaders_() {
  return {
    BASE: ['ID_PESSOA', 'NOME_COMPLETO', 'NOME_EXIBICAO', 'EMAIL_PRINCIPAL', 'STATUS_CADASTRAL', 'TECHNICAL_ACCOUNT', 'NON_HUMAN_IDENTITY', 'ENVIRONMENT', 'PURPOSE', 'EXPIRES_AT', 'CRIADO_EM', 'ATUALIZADO_EM', 'ATIVO'],
    IDENTIFICADORES: ['ID_IDENTIFICADOR', 'ID_PESSOA', 'TIPO_IDENTIFICADOR', 'VALOR_IDENTIFICADOR', 'PRINCIPAL', 'ATIVO'],
    VINCULOS: ['ID_VINCULO', 'ID_PESSOA', 'TIPO_VINCULO', 'STATUS_VINCULO', 'DATA_INICIO', 'DATA_FIM', 'MOTIVO_INICIO', 'MOTIVO_FIM', 'FONTE', 'OBS_INTERNA', 'ATIVO'],
    ACESSOS_EXCECOES: ['ID_EXCECAO', 'ID_PESSOA', 'PERFIL_EXTRA', 'PERMISSAO_EXTRA', 'DATA_INICIO', 'DATA_FIM', 'STATUS', 'JUSTIFICATIVA', 'APROVADO_POR', 'OBS']
  };
}

function coreTechAdminAssertSchema_(data) {
  var required = coreTechAdminRequiredHeaders_();
  Object.keys(required).forEach(function(key) {
    var actual = {};
    (data[key].headers || []).forEach(function(header) { actual[String(header || '').trim()] = true; });
    required[key].forEach(function(header) {
      if (!actual[header]) throw coreTechAdminError_('IDENTIDADE_TECNICA_SCHEMA_' + key + '_INCOMPLETO');
    });
  });
}

function coreTechAdminMissingHeaders_(data) {
  var required = coreTechAdminRequiredHeaders_();
  var missing = {};
  Object.keys(required).forEach(function(key) {
    var actual = {};
    (data[key].headers || []).forEach(function(header) {
      actual[String(header || '').trim()] = true;
    });
    missing[key] = required[key].filter(function(header) { return !actual[header]; });
  });
  return missing;
}

function coreTechAdminEnsureFixedSchemaForCreate_(data) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:110:coreTechAdminEnsureFixedSchemaForCreate_");try {

  var missing = coreTechAdminMissingHeaders_(data);
  Object.keys(missing).forEach(function(key) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:112:<anonymous>");try {

    if (!missing[key].length) return;
    var sheet = data[key].sheet;
    var startColumn = Math.max(1, sheet.getLastColumn() + 1);
    (f4MemberLegacyFence_().enter("core/26b_core_technical_identity_admin.js:116:5120"),sheet.getRange(1, startColumn, 1, missing[key].length).setValues([missing[key].slice()]));
  
} finally { f4MemberLegacyFence_().end(__gFenceToken); }
});
  return missing;

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechAdminSchemaReady_(data) {
  var missing = coreTechAdminMissingHeaders_(data);
  return Object.keys(missing).every(function(key) { return missing[key].length === 0; });
}

function coreTechAdminFind_(data, identifier) {
  var matches = (data.IDENTIFICADORES.records || []).filter(function(row) {
    return String(row.TIPO_IDENTIFICADOR || '').trim().toUpperCase() === 'EMAIL' &&
      String(row.VALOR_IDENTIFICADOR || '').trim().toLowerCase() === identifier;
  });
  var ids = {};
  matches.forEach(function(row) { if (row.ID_PESSOA) ids[String(row.ID_PESSOA).trim()] = true; });
  (data.BASE.records || []).forEach(function(row) {
    if (String(row.EMAIL_PRINCIPAL || '').trim().toLowerCase() === identifier && row.ID_PESSOA) {
      ids[String(row.ID_PESSOA).trim()] = true;
    }
  });
  var idList = Object.keys(ids);
  if (idList.length > 1) throw coreTechAdminError_('IDENTIDADE_TECNICA_IDENTIFICADOR_DUPLICADO');
  if (!idList.length) return null;
  var id = idList[0];
  var base = (data.BASE.records || []).filter(function(row) { return String(row.ID_PESSOA || '').trim() === id; });
  if (base.length !== 1) throw coreTechAdminError_('IDENTIDADE_TECNICA_BASE_INCOMPATIVEL');
  return {
    id: id,
    pessoa: base[0],
    identificadores: (data.IDENTIFICADORES.records || []).filter(function(row) { return String(row.ID_PESSOA || '').trim() === id; }),
    vinculos: (data.VINCULOS.records || []).filter(function(row) { return String(row.ID_PESSOA || '').trim() === id; }),
    excecoes: (data.ACESSOS_EXCECOES.records || []).filter(function(row) { return String(row.ID_PESSOA || '').trim() === id; })
  };
}

function coreTechAdminIsActive_(row) {
  var status = String((row || {}).STATUS_VINCULO || (row || {}).STATUS || '').trim().toUpperCase();
  var active = String((row || {}).ATIVO || '').trim().toUpperCase();
  return status === 'ATIVO' || status === 'ATIVA' || active === 'SIM' || active === 'TRUE';
}

function coreTechAdminInspectState_(found, now) {
  if (!found) return { exists: false, compatible: false, active: false, expired: false };
  var activeLinks = found.vinculos.filter(coreTechAdminIsActive_);
  var bundle = { pessoa: found.pessoa, vinculos: activeLinks };
  var evaluation = coreTechnicalIdentityEvaluate_(bundle, { ambiente: 'DEV', now: now || new Date() });
  var compatibleLinks = activeLinks.length === 1 &&
    String(activeLinks[0].TIPO_VINCULO || '').trim().toUpperCase() === CORE_TECHNICAL_IDENTITY_LINK_TYPE;
  var noForbiddenLink = found.vinculos.every(function(row) {
    return String(row.TIPO_VINCULO || '').trim().toUpperCase() === CORE_TECHNICAL_IDENTITY_LINK_TYPE;
  });
  var noRga = found.identificadores.every(function(row) {
    return String(row.TIPO_IDENTIFICADOR || '').trim().toUpperCase() !== 'RGA';
  });
  var withinMaximumLifetime = !!evaluation.expiresAt &&
    new Date(evaluation.expiresAt).getTime() - (now || new Date()).getTime() <= CORE_TECH_ADMIN_MAX_LIFETIME_MS;
  return {
    exists: true,
    compatible: evaluation.valid && compatibleLinks && noForbiddenLink && noRga && withinMaximumLifetime && evaluation.purpose === CORE_TECH_ADMIN_PURPOSE,
    active: evaluation.valid && compatibleLinks,
    expired: evaluation.expired === true,
    expiresAt: evaluation.expiresAt || ''
  };
}

function coreTechAdminPermissionList_(found, now) {
  if (!found || !coreTechAdminInspectState_(found, now).compatible) return [];
  var granted = {};
  found.excecoes.filter(coreTechAdminIsActive_).forEach(function(row) {
    String(row.PERMISSAO_EXTRA || '').split(/[;,|]/).forEach(function(permission) {
      granted[String(permission || '').trim().toLowerCase()] = true;
    });
  });
  return CORE_TECHNICAL_IDENTITY_ALLOWED_PERMISSIONS.filter(function(permission) { return granted[permission] === true; });
}

function coreTechAdminSetCell_(sheet, row, header, value) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:194:coreTechAdminSetCell_");try {

  var map = core_headerMap_(sheet, 1);
  var col = core_getCol_(map, header);
  if (!col) throw coreTechAdminError_('IDENTIDADE_TECNICA_SCHEMA_WRITE_INCOMPLETO');
  (f4MemberLegacyFence_().enter("core/26b_core_technical_identity_admin.js:198:9345"),sheet.getRange(row, col).setValue(value));

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechAdminEndRows_(data, found, nowIso) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:201:coreTechAdminEndRows_");try {

  if (!found) return false;
  coreTechAdminSetCell_(data.BASE.sheet, found.pessoa.__rowNumber, 'STATUS_CADASTRAL', 'ENCERRADA');
  coreTechAdminSetCell_(data.BASE.sheet, found.pessoa.__rowNumber, 'ATIVO', 'NAO');
  coreTechAdminSetCell_(data.BASE.sheet, found.pessoa.__rowNumber, 'ATUALIZADO_EM', nowIso);
  found.identificadores.forEach(function(row) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:206:<anonymous>");try {

    coreTechAdminSetCell_(data.IDENTIFICADORES.sheet, row.__rowNumber, 'ATIVO', 'NAO');
  
} finally { f4MemberLegacyFence_().end(__gFenceToken); }
});
  found.vinculos.forEach(function(row) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:209:<anonymous>");try {

    if (!coreTechAdminIsActive_(row)) return;
    coreTechAdminSetCell_(data.VINCULOS.sheet, row.__rowNumber, 'STATUS_VINCULO', 'ENCERRADO');
    coreTechAdminSetCell_(data.VINCULOS.sheet, row.__rowNumber, 'DATA_FIM', nowIso);
    coreTechAdminSetCell_(data.VINCULOS.sheet, row.__rowNumber, 'MOTIVO_FIM', 'ROLLBACK_TECNICO_DEV');
    coreTechAdminSetCell_(data.VINCULOS.sheet, row.__rowNumber, 'ATIVO', 'NAO');
  
} finally { f4MemberLegacyFence_().end(__gFenceToken); }
});
  found.excecoes.forEach(function(row) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:216:<anonymous>");try {

    if (!coreTechAdminIsActive_(row)) return;
    coreTechAdminSetCell_(data.ACESSOS_EXCECOES.sheet, row.__rowNumber, 'STATUS', 'ENCERRADO');
    coreTechAdminSetCell_(data.ACESSOS_EXCECOES.sheet, row.__rowNumber, 'DATA_FIM', nowIso);
  
} finally { f4MemberLegacyFence_().end(__gFenceToken); }
});
  return true;

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechnicalIdentityAdminCapabilities_(request) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:224:coreTechnicalIdentityAdminCapabilities_");try {

  coreTechAdminAssertDev_(request, false);
  var data = coreTechAdminOpenData_();
  return coreTechAdminResponse_({
    code: 'IDENTIDADE_TECNICA_ADMIN_CAPABILITIES_OK',
    read: true,
    create: true,
    end: true,
    permissions: true,
    portalUsersInspect: typeof coreTechnicalPortalUserAdminInspect_ === 'function',
    portalUsersConditionalDelete: typeof coreTechnicalPortalUserAdminDelete_ === 'function',
    rollback: true,
    prodBlocked: true,
    schemaReady: coreTechAdminSchemaReady_(data),
    schemaPreparedOnCreate: true,
    contractVersion: 'CODEX_E2E_DEV_V1'
  });

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechnicalIdentityAdminInspect_(request) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:243:coreTechnicalIdentityAdminInspect_");try {

  var gate = coreTechAdminAssertDev_(request, true);
  var found = coreTechAdminFind_(coreTechAdminOpenData_(), gate.identifier);
  var state = coreTechAdminInspectState_(found, new Date());
  return coreTechAdminResponse_({
    code: 'IDENTIDADE_TECNICA_ADMIN_INSPECT_OK',
    identifierHash: coreTechAdminHash_(gate.identifier),
    exists: state.exists,
    compatible: state.compatible,
    active: state.active,
    expired: state.expired,
    expiresAt: state.expiresAt || ''
  });

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechnicalIdentityAdminCreate_(request) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:258:coreTechnicalIdentityAdminCreate_");try {

  var gate = coreTechAdminAssertDev_(request, true);
  if (String(request.confirmation || '') !== CORE_TECH_ADMIN_CONFIRM_CREATE) throw coreTechAdminError_('IDENTIDADE_TECNICA_CONFIRMACAO_CRIACAO_INVALIDA');
  var now = new Date();
  var expiresAt = new Date(request.expiresAt || '');
  if (isNaN(expiresAt.getTime()) || expiresAt <= now || expiresAt.getTime() - now.getTime() > CORE_TECH_ADMIN_MAX_LIFETIME_MS) {
    throw coreTechAdminError_('IDENTIDADE_TECNICA_EXPIRACAO_FORA_DO_CONTRATO');
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var data = coreTechAdminOpenData_();
    coreTechAdminEnsureFixedSchemaForCreate_(data);
    data = coreTechAdminOpenData_();
    coreTechAdminAssertSchema_(data);
    var found = coreTechAdminFind_(data, gate.identifier);
    if (found) {
      var state = coreTechAdminInspectState_(found, now);
      if (!state.compatible) throw coreTechAdminError_('IDENTIDADE_TECNICA_EXISTENTE_INCOMPATIVEL');
      return coreTechAdminResponse_({ code: 'IDENTIDADE_TECNICA_REUTILIZADA', identifierHash: coreTechAdminHash_(gate.identifier), created: false, compatible: true, expiresAt: state.expiresAt });
    }
    var id = Utilities.getUuid();
    var iso = now.toISOString();
    try {
      core_appendObjectByHeaders_(data.BASE.sheet, {
        ID_PESSOA: id, NOME_COMPLETO: 'Identidade tecnica CODEX E2E', NOME_EXIBICAO: 'DIRETORIA DEV',
        EMAIL_PRINCIPAL: gate.identifier, STATUS_CADASTRAL: 'ATIVA', TECHNICAL_ACCOUNT: 'SIM',
        NON_HUMAN_IDENTITY: 'SIM', ENVIRONMENT: 'DEV', PURPOSE: CORE_TECH_ADMIN_PURPOSE,
        EXPIRES_AT: expiresAt.toISOString(), OBS_INTERNA: 'IDENTIDADE_TECNICA_NAO_HUMANA',
        CRIADO_EM: iso, ATUALIZADO_EM: iso, ATIVO: 'SIM'
      });
      core_appendObjectByHeaders_(data.IDENTIFICADORES.sheet, {
        ID_IDENTIFICADOR: Utilities.getUuid(), ID_PESSOA: id, TIPO_IDENTIFICADOR: 'EMAIL',
        VALOR_IDENTIFICADOR: gate.identifier, PRINCIPAL: 'SIM', ATIVO: 'SIM', OBS: 'CODEX_E2E_DEV'
      });
      core_appendObjectByHeaders_(data.VINCULOS.sheet, {
        ID_VINCULO: Utilities.getUuid(), ID_PESSOA: id, TIPO_VINCULO: CORE_TECHNICAL_IDENTITY_LINK_TYPE,
        STATUS_VINCULO: 'ATIVO', DATA_INICIO: iso, MOTIVO_INICIO: 'CODEX_E2E_DEV', FONTE: 'EXECUTION_API_MYSELF',
        OBS_INTERNA: 'SEM_RGA_SEM_INGRESSO_SEM_CONSENTIMENTO', ATIVO: 'SIM'
      });
      core_appendObjectByHeaders_(data.ACESSOS_EXCECOES.sheet, {
        ID_EXCECAO: Utilities.getUuid(), ID_PESSOA: id, PERFIL_EXTRA: CORE_TECH_ADMIN_PROFILE,
        PERMISSAO_EXTRA: CORE_TECHNICAL_IDENTITY_ALLOWED_PERMISSIONS.join(';'), DATA_INICIO: iso,
        DATA_FIM: expiresAt.toISOString(), STATUS: 'ATIVO', JUSTIFICATIVA: 'CODEX_E2E_DEV',
        APROVADO_POR: 'EXECUTION_API_MYSELF', OBS: 'TETO_FIXO_SEM_ADMIN'
      });
    } catch (writeError) {
      var partialData = coreTechAdminOpenData_();
      var partial = coreTechAdminFind_(partialData, gate.identifier);
      if (partial) coreTechAdminEndRows_(partialData, partial, new Date().toISOString());
      throw writeError;
    }
    return coreTechAdminResponse_({ code: 'IDENTIDADE_TECNICA_CRIADA', identifierHash: coreTechAdminHash_(gate.identifier), created: true, compatible: true, expiresAt: expiresAt.toISOString() });
  } finally {
    lock.releaseLock();
  }

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechnicalIdentityAdminApplyPermissions_(request) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:316:coreTechnicalIdentityAdminApplyPermissions_");try {

  var gate = coreTechAdminAssertDev_(request, true);
  if (String(request.confirmation || '') !== CORE_TECH_ADMIN_CONFIRM_PERMISSIONS) throw coreTechAdminError_('IDENTIDADE_TECNICA_CONFIRMACAO_PERMISSOES_INVALIDA');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var data = coreTechAdminOpenData_();
    coreTechAdminAssertSchema_(data);
    var found = coreTechAdminFind_(data, gate.identifier);
    var state = coreTechAdminInspectState_(found, new Date());
    if (!state.compatible) throw coreTechAdminError_('IDENTIDADE_TECNICA_INCOMPATIVEL');
    var nowIso = new Date().toISOString();
    found.excecoes.forEach(function(row) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:328:<anonymous>");try {

      if (!coreTechAdminIsActive_(row)) return;
      coreTechAdminSetCell_(data.ACESSOS_EXCECOES.sheet, row.__rowNumber, 'STATUS', 'ENCERRADO');
      coreTechAdminSetCell_(data.ACESSOS_EXCECOES.sheet, row.__rowNumber, 'DATA_FIM', nowIso);
    
} finally { f4MemberLegacyFence_().end(__gFenceToken); }
});
    core_appendObjectByHeaders_(data.ACESSOS_EXCECOES.sheet, {
      ID_EXCECAO: Utilities.getUuid(), ID_PESSOA: found.id, PERFIL_EXTRA: CORE_TECH_ADMIN_PROFILE,
      PERMISSAO_EXTRA: CORE_TECHNICAL_IDENTITY_ALLOWED_PERMISSIONS.join(';'), DATA_INICIO: nowIso,
      DATA_FIM: state.expiresAt, STATUS: 'ATIVO', JUSTIFICATIVA: 'CODEX_E2E_DEV',
      APROVADO_POR: 'EXECUTION_API_MYSELF', OBS: 'TETO_FIXO_SEM_ADMIN'
    });
    return coreTechAdminResponse_({ code: 'IDENTIDADE_TECNICA_PERMISSOES_EXATAS', identifierHash: coreTechAdminHash_(gate.identifier), permissions: CORE_TECHNICAL_IDENTITY_ALLOWED_PERMISSIONS.slice() });
  } finally {
    lock.releaseLock();
  }

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechnicalIdentityAdminGetPermissions_(request) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:345:coreTechnicalIdentityAdminGetPermissions_");try {

  var gate = coreTechAdminAssertDev_(request, true);
  var found = coreTechAdminFind_(coreTechAdminOpenData_(), gate.identifier);
  return coreTechAdminResponse_({ code: 'IDENTIDADE_TECNICA_ADMIN_PERMISSIONS_OK', identifierHash: coreTechAdminHash_(gate.identifier), permissions: coreTechAdminPermissionList_(found, new Date()) });

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechnicalIdentityAdminEnd_(request) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:351:coreTechnicalIdentityAdminEnd_");try {

  var gate = coreTechAdminAssertDev_(request, true);
  if (String(request.confirmation || '') !== CORE_TECH_ADMIN_CONFIRM_END) throw coreTechAdminError_('IDENTIDADE_TECNICA_CONFIRMACAO_ENCERRAMENTO_INVALIDA');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var data = coreTechAdminOpenData_();
    coreTechAdminAssertSchema_(data);
    var found = coreTechAdminFind_(data, gate.identifier);
    var changed = coreTechAdminEndRows_(data, found, new Date().toISOString());
    return coreTechAdminResponse_({ code: changed ? 'IDENTIDADE_TECNICA_ENCERRADA' : 'IDENTIDADE_TECNICA_AUSENTE', identifierHash: coreTechAdminHash_(gate.identifier), ended: changed, complete: true });
  } finally {
    lock.releaseLock();
  }

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechnicalIdentityAdminRollback_(request) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26b_core_technical_identity_admin.js:367:coreTechnicalIdentityAdminRollback_");try {

  return coreTechnicalIdentityAdminEnd_(request);

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}
