/**
 * Contrato canonico de identidades tecnicas nao humanas.
 *
 * Regra de seguranca: qualquer marcador tecnico, mesmo parcial ou invalido,
 * exclui a identidade de processos institucionais humanos. Autenticacao e
 * autorizacao exigem o contrato completo, vigente e executado em DEV.
 */

var CORE_TECHNICAL_IDENTITY_LINK_TYPE = 'IDENTIDADE_TECNICA';
var CORE_TECHNICAL_IDENTITY_ALLOWED_PERMISSIONS = Object.freeze([
  'portal:acessar',
  'atividades:gerir',
  'apresentacoes:gerir'
]);

function coreTechnicalIdentityText_(value) {
  return String(value == null ? '' : value).trim();
}

function coreTechnicalIdentityToken_(value) {
  return coreTechnicalIdentityText_(value).toUpperCase();
}

function coreTechnicalIdentityTrue_(value) {
  if (value === true) return true;
  return ['TRUE', 'SIM', '1'].indexOf(coreTechnicalIdentityToken_(value)) >= 0;
}

function coreTechnicalIdentityValue_(record, aliases) {
  var source = record && typeof record === 'object' ? record : {};
  for (var i = 0; i < aliases.length; i++) {
    if (Object.prototype.hasOwnProperty.call(source, aliases[i]) && source[aliases[i]] !== '') {
      return source[aliases[i]];
    }
  }
  return '';
}

function coreTechnicalIdentityMergeBundle_(bundle) {
  var source = bundle && typeof bundle === 'object' ? bundle : {};
  var pessoa = source.pessoa || source.PESSOA || {};
  var resumo = source.resumoOperacional || source.RESUMO || {};
  var vinculos = source.vinculos || source.VINCULOS || [];
  var currentLink = typeof corePortalGetCurrentLink_ === 'function'
    ? corePortalGetCurrentLink_(source)
    : {};
  if (!currentLink || !currentLink.tipoVinculo) {
    for (var i = 0; i < vinculos.length; i++) {
      if (coreTechnicalIdentityToken_(vinculos[i].TIPO_VINCULO || vinculos[i].tipoVinculo) === CORE_TECHNICAL_IDENTITY_LINK_TYPE) {
        currentLink = {
          tipoVinculo: vinculos[i].TIPO_VINCULO || vinculos[i].tipoVinculo,
          statusVinculo: vinculos[i].STATUS_VINCULO || vinculos[i].statusVinculo
        };
        break;
      }
    }
  }
  var merged = Object.assign({}, pessoa);
  Object.keys(resumo).forEach(function(key) {
    if (resumo[key] !== '' && resumo[key] != null) merged[key] = resumo[key];
  });
  return Object.assign(merged, {
    TIPO_VINCULO: (currentLink && currentLink.tipoVinculo) || resumo.TIPO_VINCULO_ATUAL || pessoa.TIPO_VINCULO || '',
    STATUS_VINCULO: (currentLink && currentLink.statusVinculo) || resumo.STATUS_VINCULO_ATUAL || pessoa.STATUS_VINCULO || ''
  });
}

function coreTechnicalIdentityEvaluate_(identity, options) {
  var opts = options || {};
  var record = identity && (identity.pessoa || identity.resumoOperacional || identity.vinculos)
    ? coreTechnicalIdentityMergeBundle_(identity)
    : (identity || {});
  var technicalRaw = coreTechnicalIdentityValue_(record, ['technicalAccount', 'TECHNICAL_ACCOUNT']);
  var nonHumanRaw = coreTechnicalIdentityValue_(record, ['nonHumanIdentity', 'NON_HUMAN_IDENTITY']);
  var environment = coreTechnicalIdentityToken_(coreTechnicalIdentityValue_(record, ['environment', 'ENVIRONMENT', 'AMBIENTE']));
  var purpose = coreTechnicalIdentityText_(coreTechnicalIdentityValue_(record, ['purpose', 'PURPOSE', 'FINALIDADE']));
  var expiresAtRaw = coreTechnicalIdentityValue_(record, ['expiresAt', 'EXPIRES_AT', 'EXPIRA_EM']);
  var linkType = coreTechnicalIdentityToken_(coreTechnicalIdentityValue_(record, ['tipoVinculo', 'tipoVinculoAtual', 'TIPO_VINCULO', 'TIPO_VINCULO_ATUAL']));
  var technicalAccount = coreTechnicalIdentityTrue_(technicalRaw);
  var nonHumanIdentity = coreTechnicalIdentityTrue_(nonHumanRaw);
  var isTechnicalIdentity = linkType === CORE_TECHNICAL_IDENTITY_LINK_TYPE ||
    technicalAccount || nonHumanIdentity ||
    coreTechnicalIdentityText_(technicalRaw) !== '' || coreTechnicalIdentityText_(nonHumanRaw) !== '' ||
    !!purpose || !!coreTechnicalIdentityText_(expiresAtRaw);
  var expiresAtDate = expiresAtRaw instanceof Date ? new Date(expiresAtRaw.getTime()) : new Date(expiresAtRaw);
  var expiresAtValid = !!coreTechnicalIdentityText_(expiresAtRaw) && !isNaN(expiresAtDate.getTime());
  var now = opts.now instanceof Date ? opts.now : new Date(opts.now || Date.now());
  var expired = expiresAtValid && expiresAtDate.getTime() <= now.getTime();
  var runtimeEnvironment = coreTechnicalIdentityToken_(opts.ambiente || opts.environment || opts.env || '');
  if (!runtimeEnvironment && typeof core_getCurrentEnv_ === 'function') {
    runtimeEnvironment = coreTechnicalIdentityToken_(core_getCurrentEnv_());
  }
  var reason = '';
  if (isTechnicalIdentity) {
    if (runtimeEnvironment !== 'DEV') reason = 'IDENTIDADE_TECNICA_FORA_DE_DEV';
    else if (linkType !== CORE_TECHNICAL_IDENTITY_LINK_TYPE) reason = 'IDENTIDADE_TECNICA_VINCULO_INVALIDO';
    else if (!technicalAccount || !nonHumanIdentity) reason = 'IDENTIDADE_TECNICA_MARCADORES_INVALIDOS';
    else if (environment !== 'DEV') reason = 'IDENTIDADE_TECNICA_AMBIENTE_INVALIDO';
    else if (!purpose) reason = 'IDENTIDADE_TECNICA_FINALIDADE_AUSENTE';
    else if (!expiresAtValid) reason = 'IDENTIDADE_TECNICA_EXPIRACAO_INVALIDA';
    else if (expired) reason = 'IDENTIDADE_TECNICA_EXPIRADA';
  }
  var valid = isTechnicalIdentity && !reason;
  return Object.freeze({
    isTechnicalIdentity: isTechnicalIdentity,
    technicalAccount: technicalAccount,
    nonHumanIdentity: nonHumanIdentity,
    environment: environment,
    purpose: purpose,
    expiresAt: expiresAtValid ? expiresAtDate.toISOString() : '',
    expired: expired,
    valid: valid,
    authorizedForDev: valid && runtimeEnvironment === 'DEV',
    blockedReason: reason,
    excludeFromInstitutionalProcesses: isTechnicalIdentity
  });
}

function coreTechnicalIdentityApplyPermissionCeiling_(permissions, evaluation) {
  var state = evaluation || {};
  if (!state.isTechnicalIdentity) return (permissions || []).slice();
  if (!state.valid || !state.authorizedForDev) return [];
  var normalized = {};
  (permissions || []).forEach(function(permission) {
    normalized[String(permission || '').trim().toLowerCase()] = true;
  });
  return CORE_TECHNICAL_IDENTITY_ALLOWED_PERMISSIONS.filter(function(permission) {
    return normalized[permission] === true;
  });
}

function coreTechnicalIdentityShouldExclude_(identity, options) {
  return coreTechnicalIdentityEvaluate_(identity, options || {}).isTechnicalIdentity;
}

function coreTechnicalIdentityListExcludedIds_(options) {
  var opts = options || {};
  var environment = coreTechnicalIdentityToken_(opts.ambiente || opts.environment || opts.env || '');
  if (!environment && typeof core_getCurrentEnv_ === 'function') environment = coreTechnicalIdentityToken_(core_getCurrentEnv_());
  if (environment !== 'DEV') {
    throw new Error('IDENTIDADE_TECNICA_FORA_DE_DEV');
  }
  var report = core_domainsV2NewReadReport_('IDENTIDADES_TECNICAS_LISTAR_IDS');
  var data = core_domainsV2OpenPessoasSubset_(['BASE', 'VINCULOS', 'RESUMO'], report, { ambiente: environment });
  if (report.totalErros) throw new Error('Pessoas v2 indisponivel: ' + JSON.stringify(report.erros));
  var base = (data.PESSOAS_BASE && data.PESSOAS_BASE.records) || [];
  var resumo = (data.PESSOAS_RESUMO_OPERACIONAL && data.PESSOAS_RESUMO_OPERACIONAL.records) || [];
  var vinculos = (data.VINCULOS_GEAPA && data.VINCULOS_GEAPA.records) || [];
  var byId = {};
  base.forEach(function(row) {
    var id = coreTechnicalIdentityText_(row.ID_PESSOA);
    if (id) byId[id] = Object.assign({}, byId[id] || {}, row);
  });
  resumo.forEach(function(row) {
    var id = coreTechnicalIdentityText_(row.ID_PESSOA);
    if (id) byId[id] = Object.assign({}, byId[id] || {}, row);
  });
  vinculos.forEach(function(row) {
    var id = coreTechnicalIdentityText_(row.ID_PESSOA);
    if (!id) return;
    var current = byId[id] || {};
    if (coreTechnicalIdentityToken_(row.TIPO_VINCULO) === CORE_TECHNICAL_IDENTITY_LINK_TYPE) {
      current.TIPO_VINCULO = CORE_TECHNICAL_IDENTITY_LINK_TYPE;
    }
    byId[id] = current;
  });
  return Object.freeze(Object.keys(byId).filter(function(id) {
    return coreTechnicalIdentityShouldExclude_(byId[id], { ambiente: environment });
  }).sort());
}
