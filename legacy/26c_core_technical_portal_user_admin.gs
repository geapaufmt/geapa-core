/**
 * Administracao estreita de portalUsers criada exclusivamente para o rollback
 * da identidade tecnica CODEX_E2E DEV. Nao lista documentos nem expoe delete
 * generico.
 */

var CORE_TECH_PORTAL_USER_DELETE_CONFIRMATION =
  'EXCLUIR_PORTAL_USER_TECNICO_CODEX_E2E_DEV';

function coreTechPortalUserUid_(value) {
  var uid = String(value || '').trim();
  if (!/^[A-Za-z0-9_-]{1,128}$/.test(uid)) {
    throw coreTechAdminError_('PORTAL_USER_TECNICO_UID_INVALIDO');
  }
  return uid;
}

function coreTechPortalUserPermissions_(data) {
  var raw = Array.isArray(data && data.permissoes)
    ? data.permissoes
    : Object.keys(data && data.permissions || {}).filter(function(key) {
        return data.permissions[key] === true;
      });
  return raw.map(function(value) {
    return String(value || '').trim().toLowerCase();
  }).filter(Boolean).sort();
}

function coreTechPortalUserCompatible_(data) {
  var record = data || {};
  var expiresAt = new Date(record.expiresAt || '');
  var actual = coreTechPortalUserPermissions_(record);
  var expected = CORE_TECHNICAL_IDENTITY_ALLOWED_PERMISSIONS.slice().sort();
  return coreTechnicalIdentityTrue_(record.technicalAccount) &&
    coreTechnicalIdentityTrue_(record.nonHumanIdentity) &&
    String(record.environment || '').trim().toUpperCase() === 'DEV' &&
    String(record.purpose || '').trim().toUpperCase() === CORE_TECH_ADMIN_PURPOSE &&
    !isNaN(expiresAt.getTime()) && expiresAt.getTime() > Date.now() &&
    actual.length === expected.length && actual.every(function(permission, index) {
      return permission === expected[index];
    });
}

function coreTechPortalUserRead_(uid) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26c_core_technical_portal_user_admin.js:44:coreTechPortalUserRead_");try {

  return coreFirestoreEnvironmentGetDocumentWithVersion_(
    'portalUsers/' + uid,
    { ambiente: 'DEV', environment: 'DEV' }
  );

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechnicalPortalUserAdminInspect_(request) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26c_core_technical_portal_user_admin.js:51:coreTechnicalPortalUserAdminInspect_");try {

  coreTechAdminAssertDev_(request || {}, false);
  var uid = coreTechPortalUserUid_(request && request.uid);
  var result = coreTechPortalUserRead_(uid);
  if (!result || result.ok !== true) {
    throw coreTechAdminError_('PORTAL_USER_TECNICO_INSPECAO_FALHOU');
  }
  if (result.found !== true) {
    return coreTechAdminResponse_({
      code: 'PORTAL_USER_TECNICO_AUSENTE', uidHash: coreTechAdminHash_(uid),
      found: false, compatible: false, updateTime: ''
    });
  }
  var compatible = coreTechPortalUserCompatible_(result.data);
  return coreTechAdminResponse_({
    code: compatible
      ? 'PORTAL_USER_TECNICO_COMPATIVEL'
      : 'PORTAL_USER_TECNICO_INCOMPATIVEL',
    uidHash: coreTechAdminHash_(uid), found: true, compatible: compatible,
    updateTime: compatible ? String(result.updateTime || '') : ''
  });

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechPortalUserCommitDelete_(uid, expectedUpdateTime) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26c_core_technical_portal_user_admin.js:74:coreTechPortalUserCommitDelete_");try {

  var config = coreFirestoreGetConfig_({ ambiente: 'DEV', environment: 'DEV' });
  if (!config.projectId || String(config.projectId).trim() !== 'geapa-dev') {
    throw coreTechAdminError_('PORTAL_USER_TECNICO_PROJECT_DEV_INVALIDO');
  }
  var url = CORE_FIRESTORE_API_BASE + '/' +
    coreFirestoreDatabasePath_(config) + '/documents:commit';
  var response = coreFirestoreRequest_(url, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify({ writes: [{
      delete: coreFirestoreBuildDocumentName_('portalUsers/' + uid, config),
      currentDocument: { updateTime: expectedUpdateTime }
    }] })
  });
  if (!response.ok) {
    var status = String(response.firestoreError && response.firestoreError.status || '').toUpperCase();
    throw coreTechAdminError_(status === 'FAILED_PRECONDITION' || status === 'ABORTED'
      ? 'PORTAL_USER_TECNICO_UPDATE_TIME_DIVERGENTE'
      : 'PORTAL_USER_TECNICO_DELETE_FALHOU');
  }
  var payload = JSON.parse(response.body || '{}');
  if (!Array.isArray(payload.writeResults) || payload.writeResults.length !== 1) {
    throw coreTechAdminError_('PORTAL_USER_TECNICO_DELETE_RESPOSTA_INCOMPLETA');
  }
  return true;

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}

function coreTechnicalPortalUserAdminDelete_(request) {
const __gFenceToken=f4MemberLegacyFence_().begin("core/26c_core_technical_portal_user_admin.js:102:coreTechnicalPortalUserAdminDelete_");try {

  coreTechAdminAssertDev_(request || {}, false);
  var uid = coreTechPortalUserUid_(request && request.uid);
  if (!request || request.baselineAbsentBeforeApply !== true) {
    throw coreTechAdminError_('PORTAL_USER_TECNICO_BASELINE_AUSENTE_OBRIGATORIO');
  }
  if (String(request.confirmation || '') !==
      CORE_TECH_PORTAL_USER_DELETE_CONFIRMATION) {
    throw coreTechAdminError_('PORTAL_USER_TECNICO_CONFIRMACAO_DELETE_INVALIDA');
  }
  var expectedUpdateTime = String(request.expectedUpdateTime || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,9})?Z$/.test(expectedUpdateTime) ||
      isNaN(new Date(expectedUpdateTime).getTime())) {
    throw coreTechAdminError_('PORTAL_USER_TECNICO_UPDATE_TIME_INVALIDO');
  }
  var current = coreTechPortalUserRead_(uid);
  if (!current || current.ok !== true) {
    throw coreTechAdminError_('PORTAL_USER_TECNICO_INSPECAO_FALHOU');
  }
  if (current.found !== true) {
    return coreTechAdminResponse_({
      code: 'PORTAL_USER_TECNICO_DELETE_IDEMPOTENTE', uidHash: coreTechAdminHash_(uid),
      deleted: false, complete: true
    });
  }
  if (!coreTechPortalUserCompatible_(current.data)) {
    throw coreTechAdminError_('PORTAL_USER_TECNICO_DOCUMENTO_INCOMPATIVEL');
  }
  if (String(current.updateTime || '') !== expectedUpdateTime) {
    throw coreTechAdminError_('PORTAL_USER_TECNICO_UPDATE_TIME_DIVERGENTE');
  }
  coreTechPortalUserCommitDelete_(uid, expectedUpdateTime);
  return coreTechAdminResponse_({
    code: 'PORTAL_USER_TECNICO_EXCLUIDO', uidHash: coreTechAdminHash_(uid),
    deleted: true, complete: true, preconditionApplied: true
  });

} finally { f4MemberLegacyFence_().end(__gFenceToken); }
}
