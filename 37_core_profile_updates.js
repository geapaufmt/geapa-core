/**
 * Contratos seguros de atualizacao cadastral para o Portal HOMOLOG.
 *
 * Regras centrais:
 * - a identidade vem da sessao oficial; ID_PESSOA/RGA do payload sao rejeitados;
 * - campos de baixo risco escrevem somente nas fontes Pessoas V2;
 * - campos sensiveis sempre passam por solicitacao, analise e aplicacao separadas;
 * - nenhuma rotina deste arquivo escreve em PROD;
 * - PESSOAS_RESUMO_OPERACIONAL nunca e editada diretamente.
 */

var CORE_PERFIL_SOLICITACOES_KEY = 'PESSOAS_V2_SOLICITACOES_ATUALIZACAO_CADASTRAL';
var CORE_PERFIL_SOLICITACOES_SHEET = 'SOLICITACOES_ATUALIZACAO_CADASTRAL';
var CORE_PERFIL_ADMIN_PERMISSION = 'membros:analisar_correcoes';
var CORE_PERFIL_SETUP_CONFIRMATION = 'PREPARAR_SOLICITACOES_CADASTRAIS_DEV';
var CORE_PERFIL_MAX_RESUMO = 3000;

var CORE_PERFIL_STATUS = Object.freeze([
  'PENDENTE',
  'EM_ANALISE',
  'COMPLEMENTO_SOLICITADO',
  'APROVADA',
  'INDEFERIDA',
  'CANCELADA',
  'APLICADA',
  'ERRO_APLICACAO'
]);

var CORE_PERFIL_SENSITIVE_FIELDS = Object.freeze({
  NOME_COMPLETO: Object.freeze({ source: 'PESSOAS_BASE', header: 'NOME_COMPLETO', type: 'NAME' }),
  NOME_CIVIL: Object.freeze({ source: 'PESSOAS_BASE', header: 'NOME_CIVIL', type: 'NAME' }),
  CPF: Object.freeze({ source: 'PESSOAS_BASE', header: 'CPF', type: 'CPF' }),
  RGA: Object.freeze({ source: 'MEMBROS_DETALHES', header: 'RGA', type: 'RGA' }),
  DATA_NASCIMENTO: Object.freeze({ source: 'PESSOAS_BASE', header: 'DATA_NASCIMENTO', type: 'DATE' }),
  EMAIL_PRINCIPAL: Object.freeze({ source: 'PESSOAS_BASE', header: 'EMAIL_PRINCIPAL', type: 'EMAIL' })
});

var CORE_PERFIL_DIRECT_FIELDS = Object.freeze({
  TELEFONE: Object.freeze({ source: 'PESSOAS_BASE', header: 'TELEFONE_PRINCIPAL', type: 'PHONE' }),
  INSTAGRAM: Object.freeze({ source: 'PESSOAS_BASE', header: 'INSTAGRAM', type: 'INSTAGRAM' }),
  CIDADE_ORIGEM: Object.freeze({ source: 'PESSOAS_BASE', header: 'CIDADE_NATAL', type: 'CITY' }),
  UF_ORIGEM: Object.freeze({ source: 'PESSOAS_BASE', header: 'UF_ORIGEM', type: 'UF' }),
  HISTORICO_ACADEMICO: Object.freeze({ source: 'MEMBROS_DETALHES', header: 'HISTORICO_ATIVIDADES_ACADEMICAS', type: 'SUMMARY' })
});

var CORE_PERFIL_DIRECT_ALIASES = Object.freeze({
  telefone: 'TELEFONE',
  telefonePrincipal: 'TELEFONE',
  instagram: 'INSTAGRAM',
  cidadeOrigem: 'CIDADE_ORIGEM',
  cidadeNatal: 'CIDADE_ORIGEM',
  ufOrigem: 'UF_ORIGEM',
  historicoAcademico: 'HISTORICO_ACADEMICO',
  resumoAcademico: 'HISTORICO_ACADEMICO'
});

var CORE_PERFIL_LINK_TYPES = Object.freeze([
  'LATTES',
  'ORCID',
  'LINKEDIN',
  'GOOGLE_SCHOLAR',
  'RESEARCHGATE',
  'SITE_PESSOAL',
  'OUTRO'
]);

function corePerfilEnvelopeOk_(data, meta) {
  return Object.freeze({ ok: true, data: Object.freeze(data || {}), meta: Object.freeze(meta || {}) });
}

function corePerfilEnvelopeError_(code, message, details) {
  var out = {
    ok: false,
    errorCode: String(code || 'ERRO_ATUALIZACAO_CADASTRAL'),
    message: String(message || 'Nao foi possivel concluir a operacao cadastral.')
  };
  if (details) out.details = Object.freeze(details);
  return Object.freeze(out);
}

function corePerfilNormalizeToken_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  }).replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function corePerfilHash_(value, deps) {
  if (deps && typeof deps.hash === 'function') return deps.hash(String(value == null ? '' : value));
  var text = String(value == null ? '' : value);
  if (typeof Utilities !== 'undefined' && Utilities.computeDigest) {
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
    return Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '');
  }
  var hash = 2166136261;
  for (var i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
  }
  return ('00000000' + (hash >>> 0).toString(16)).slice(-8);
}

function corePerfilUuid_(prefix, deps) {
  var uuid = deps && typeof deps.uuid === 'function'
    ? deps.uuid()
    : Utilities.getUuid();
  return String(prefix || '') + String(uuid).toUpperCase();
}

function corePerfilNow_(deps) {
  return deps && typeof deps.now === 'function' ? deps.now() : new Date();
}

function corePerfilAssertHomologContext_(contexto, deps) {
  var ctx = contexto && typeof contexto === 'object' ? contexto : {};
  var env = String(deps && deps.environment || ctx.ambientePortal || ctx.ambiente || '').trim().toUpperCase();
  if (env !== 'DEV' && env !== 'HOMOLOG') throw new Error('CONTEXTO_HOMOLOG_OBRIGATORIO');
  return 'DEV';
}

function corePerfilRegistryMetaDev_(key) {
  var normalized = String(key || '').trim().toUpperCase();
  var raw = core_getRegistryRaw_();
  var entry = raw[normalized] && raw[normalized].DEV;
  if (!entry || entry.ativo !== true) throw new Error('REGISTRY_DEV_INDISPONIVEL_' + normalized);
  return entry;
}

function corePerfilGetSheetByKeyDev_(key) {
  var entry = corePerfilRegistryMetaDev_(key);
  return core_getSheetById_(entry.id, entry.sheet);
}

function corePerfilSafeLogPayload_(event, details) {
  var source = details || {};
  return {
    event: String(event || '').slice(0, 80),
    ok: source.ok === true,
    field: String(source.field || '').slice(0, 60),
    status: String(source.status || '').slice(0, 40),
    requestId: String(source.requestId || '').slice(0, 80),
    changedCount: Number(source.changedCount || 0),
    code: String(source.code || '').slice(0, 80)
  };
}

function corePerfilSafeLog_(event, details) {
  Logger.log('[geapa-core][perfil-cadastral] ' + JSON.stringify(corePerfilSafeLogPayload_(event, details)));
}

function corePerfilContextEmail_(contexto) {
  if (typeof contexto === 'string') {
    return contexto.indexOf('@') >= 0 ? corePortalNormalizeEmail_(contexto) : '';
  }
  var ctx = contexto && typeof contexto === 'object' ? contexto : {};
  var official = ctx.sessaoOficial && typeof ctx.sessaoOficial === 'object' ? ctx.sessaoOficial : {};
  var candidates = [
    official.email,
    ctx.emailAutenticado,
    ctx.emailSolicitante,
    ctx.email,
    ctx.identificadorSolicitante,
    ctx.identificador
  ];
  for (var i = 0; i < candidates.length; i++) {
    if (String(candidates[i] || '').indexOf('@') < 0) continue;
    var email = corePortalNormalizeEmail_(candidates[i]);
    if (email) return email;
  }
  return '';
}

function corePerfilResolveSession_(contexto, deps) {
  deps = deps || {};
  if (deps.session) return deps.session;
  var email = corePerfilContextEmail_(contexto);
  if (!email) return null;
  var resolver = deps.resolveSession || corePortalResolverUsuarioAtual_;
  var session = resolver(email, { origem: 'perfilCadastralPortal' });
  if (!session || session.ok === false || session.autenticado !== true || !String(session.idPessoa || '').trim()) return null;
  return session;
}

function corePerfilAuthorizeOwn_(contexto, deps) {
  var session = corePerfilResolveSession_(contexto, deps);
  if (!session || session.portalAtivo !== true) {
    return { ok: false, response: corePerfilEnvelopeError_('SESSAO_INVALIDA', 'Sessao autenticada e ativa obrigatoria.') };
  }
  return { ok: true, session: session };
}

function corePerfilAuthorizeAdmin_(contexto, deps) {
  try {
    corePerfilAssertHomologContext_(contexto, deps);
  } catch (envError) {
    return { ok: false, response: corePerfilEnvelopeError_('CONTEXTO_HOMOLOG_OBRIGATORIO', 'Operacao administrativa disponivel somente no Portal HOMOLOG.') };
  }
  var auth = corePerfilAuthorizeOwn_(contexto, deps);
  if (!auth.ok) return auth;
  var permissions = Array.isArray(auth.session.permissoes) ? auth.session.permissoes : [];
  var profiles = Array.isArray(auth.session.perfisPortal) ? auth.session.perfisPortal.slice() : [];
  if (auth.session.perfilPortalEfetivo) profiles.push(auth.session.perfilPortalEfetivo);
  var homologProfileAllowed = profiles.map(corePerfilNormalizeToken_).some(function(profile) {
    return profile === 'SECRETARIA' || profile === 'DIRETORIA';
  });
  var allowed = permissions.map(corePortalNormalizePermission_).indexOf(CORE_PERFIL_ADMIN_PERMISSION) >= 0 ||
    corePerfilSessionHasPermissionDev_(auth.session, CORE_PERFIL_ADMIN_PERMISSION, deps) ||
    homologProfileAllowed;
  if (!allowed) {
    return { ok: false, response: corePerfilEnvelopeError_('PERMISSAO_NEGADA', 'Usuario sem permissao para analisar correcoes cadastrais.') };
  }
  return auth;
}

function corePerfilSessionHasPermissionDev_(session, permission, deps) {
  if (deps && typeof deps.hasDevPermission === 'function') return deps.hasDevPermission(session, permission);
  var profiles = Array.isArray(session && session.perfisPortal) ? session.perfisPortal.slice() : [];
  if (session && session.perfilPortalEfetivo) profiles.push(session.perfilPortalEfetivo);
  profiles = profiles.map(corePerfilNormalizeToken_);
  if (!profiles.length) return false;
  var records = [];
  try {
    records = core_readSheetRecords_(corePerfilGetSheetByKeyDev_('PORTAL_PERMISSOES'), { skipBlankRows: true });
  } catch (missingDevPermissionSource) {
    return false;
  }
  return records.some(function(record) {
    return profiles.indexOf(corePerfilNormalizeToken_(record.PERFIL_PORTAL)) >= 0 &&
      corePortalNormalizePermission_(record.PERMISSAO) === corePortalNormalizePermission_(permission) &&
      core_domainsV2AuditIsSim_(record.ATIVO);
  });
}

function corePerfilAssertNoTargetIdentity_(payload) {
  var source = payload && typeof payload === 'object' ? payload : {};
  var forbidden = ['idPessoa', 'ID_PESSOA', 'rga', 'RGA', 'emailPessoa', 'pessoaId'];
  for (var i = 0; i < forbidden.length; i++) {
    if (Object.prototype.hasOwnProperty.call(source, forbidden[i]) && String(source[forbidden[i]] || '').trim()) {
      throw new Error('IDENTIDADE_ALVO_NAO_PERMITIDA');
    }
  }
}

function corePerfilNormalizePhone_(value) {
  var raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  var digits = raw.replace(/\D+/g, '');
  if ((digits.length === 12 || digits.length === 13) && digits.indexOf('55') === 0) digits = digits.slice(2);
  if (digits.length !== 10 && digits.length !== 11) throw new Error('TELEFONE_INVALIDO');
  if (digits.length === 11 && digits.charAt(2) !== '9') throw new Error('TELEFONE_INVALIDO');
  return '+55' + digits;
}

function corePerfilNormalizeInstagram_(value) {
  var text = String(value == null ? '' : value).trim().replace(/^@+/, '').replace(/\s+/g, '');
  if (!text) return '';
  if (!/^[A-Za-z0-9._]{1,30}$/.test(text) || /\.\./.test(text)) throw new Error('INSTAGRAM_INVALIDO');
  return text;
}

function corePerfilNormalizeUf_(value) {
  var uf = String(value == null ? '' : value).trim().toUpperCase();
  if (!uf) return '';
  var allowed = ['AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA','PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO'];
  if (allowed.indexOf(uf) < 0) throw new Error('UF_INVALIDA');
  return uf;
}

function corePerfilNormalizeTextField_(value, maxLength, errorCode) {
  var text = String(value == null ? '' : value).trim().replace(/[ \t]+/g, ' ').replace(/\r\n?/g, '\n');
  if (text.length > maxLength) throw new Error(errorCode);
  return text;
}

function corePerfilNormalizeEmailValue_(value) {
  var email = corePortalNormalizeEmail_(value);
  if (!email) throw new Error('EMAIL_INVALIDO');
  return email;
}

function corePerfilNormalizeCpf_(value) {
  var cpf = String(value || '').replace(/\D+/g, '');
  if (cpf.length !== 11 || /^(\d)\1{10}$/.test(cpf)) throw new Error('CPF_INVALIDO');
  function digit(base, factor) {
    var sum = 0;
    for (var i = 0; i < base.length; i++) sum += Number(base.charAt(i)) * (factor - i);
    var result = (sum * 10) % 11;
    return result === 10 ? 0 : result;
  }
  if (digit(cpf.slice(0, 9), 10) !== Number(cpf.charAt(9)) || digit(cpf.slice(0, 10), 11) !== Number(cpf.charAt(10))) {
    throw new Error('CPF_INVALIDO');
  }
  return cpf;
}

function corePerfilNormalizeDate_(value) {
  var raw = String(value || '').trim();
  var match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/) || raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) throw new Error('DATA_NASCIMENTO_INVALIDA');
  var year = raw.indexOf('-') >= 0 ? Number(match[1]) : Number(match[3]);
  var month = Number(match[2]);
  var day = raw.indexOf('-') >= 0 ? Number(match[3]) : Number(match[1]);
  var date = new Date(year, month - 1, day);
  var today = new Date();
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day || year < 1900 || date > today) {
    throw new Error('DATA_NASCIMENTO_INVALIDA');
  }
  return Utilities.formatString('%04d-%02d-%02d', year, month, day);
}

function corePerfilNormalizeSensitiveValue_(field, value) {
  var cfg = CORE_PERFIL_SENSITIVE_FIELDS[field];
  if (!cfg) throw new Error('CAMPO_SENSIVEL_NAO_PERMITIDO');
  if (cfg.type === 'CPF') return corePerfilNormalizeCpf_(value);
  if (cfg.type === 'EMAIL') return corePerfilNormalizeEmailValue_(value);
  if (cfg.type === 'DATE') return corePerfilNormalizeDate_(value);
  if (cfg.type === 'RGA') {
    var rga = String(value || '').trim().replace(/\s+/g, '');
    if (!/^[A-Za-z0-9.-]{4,30}$/.test(rga)) throw new Error('RGA_INVALIDO');
    return rga;
  }
  return corePerfilNormalizeTextField_(value, 200, 'NOME_INVALIDO');
}

function corePerfilNormalizeUrl_(type, value) {
  var raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  if (/\s|[<>]/.test(raw)) throw new Error('URL_INVALIDA');
  var normalized = /^https?:\/\//i.test(raw) ? raw : 'https://' + raw;
  if (!/^https:\/\//i.test(normalized)) throw new Error('URL_PROTOCOLO_INVALIDO');
  var hostMatch = normalized.match(/^https:\/\/([^\/?#:]+)(?::\d+)?(?:[\/?#]|$)/i);
  var host = hostMatch ? hostMatch[1].toLowerCase().replace(/^www\./, '') : '';
  if (!host || host.indexOf('.') < 0 || /(^|\.)localhost$/.test(host)) throw new Error('URL_INVALIDA');
  var allowedHost = {
    LATTES: /(^|\.)lattes\.cnpq\.br$/,
    ORCID: /(^|\.)orcid\.org$/,
    LINKEDIN: /(^|\.)linkedin\.com$/,
    GOOGLE_SCHOLAR: /(^|\.)scholar\.google\.[a-z.]+$/,
    RESEARCHGATE: /(^|\.)researchgate\.net$/
  };
  if (allowedHost[type] && !allowedHost[type].test(host)) throw new Error('URL_DOMINIO_INVALIDO');
  return normalized.replace(/\/+$/, '');
}

function corePerfilNormalizeLinkType_(value) {
  var type = corePerfilNormalizeToken_(value);
  if (type === 'SITE' || type === 'SITE_PESSOAL') type = 'SITE_PESSOAL';
  if (type === 'GOOGLE_SCHOLAR_PROFILE') type = 'GOOGLE_SCHOLAR';
  if (CORE_PERFIL_LINK_TYPES.indexOf(type) < 0) throw new Error('TIPO_LINK_NAO_PERMITIDO');
  return type;
}

function corePerfilMaskValue_(field, value) {
  var text = String(value == null ? '' : value).trim();
  if (!text) return '';
  if (field === 'CPF') {
    var cpf = text.replace(/\D+/g, '');
    return cpf.length === 11 ? '***.***.***-' + cpf.slice(-2) : '***';
  }
  if (field === 'EMAIL_PRINCIPAL') {
    var parts = text.split('@');
    return parts.length === 2 ? parts[0].slice(0, 2) + '***@' + parts[1] : '***';
  }
  if (field === 'DATA_NASCIMENTO') return '**/**/' + text.slice(0, 4);
  if (field === 'RGA') return '***' + text.slice(-3);
  if (field === 'NOME_COMPLETO' || field === 'NOME_CIVIL') {
    var names = text.split(/\s+/);
    return names[0].slice(0, 1) + '***' + (names.length > 1 ? ' ' + names[names.length - 1].slice(0, 1) + '***' : '');
  }
  if (field === 'TELEFONE') return '***' + text.replace(/\D+/g, '').slice(-4);
  return text.length <= 4 ? '***' : text.slice(0, 2) + '***' + text.slice(-2);
}

function corePerfilRedactSensitiveText_(value) {
  return String(value == null ? '' : value)
    .replace(/\b\d{3}[.\s-]?\d{3}[.\s-]?\d{3}[-.\s]?\d{2}\b/g, '***.***.***-**')
    .replace(/\b[A-Z0-9._%+-]{2,}@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi, function(email) {
      return corePerfilMaskValue_('EMAIL_PRINCIPAL', email);
    });
}

function corePerfilOpenPessoas_(deps) {
  if (deps && typeof deps.openPessoas === 'function') return deps.openPessoas();
  var report = core_domainsV2NewReadReport_('PERFIL_CADASTRAL_PORTAL');
  var data = core_domainsV2OpenPessoas_(report);
  if (report.totalErros) throw new Error('PESSOAS_V2_INDISPONIVEL');
  var requestSheet = corePerfilGetSheetByKeyDev_(CORE_PERFIL_SOLICITACOES_KEY);
  var requestHeaders = requestSheet.getLastColumn() > 0
    ? requestSheet.getRange(1, 1, 1, requestSheet.getLastColumn()).getDisplayValues()[0].map(function(value) { return String(value || '').trim(); })
    : [];
  data[CORE_PERFIL_SOLICITACOES_SHEET] = {
    sheet: requestSheet,
    headers: requestHeaders,
    records: core_readSheetRecords_(requestSheet, { skipBlankRows: true })
  };
  return data;
}

function corePerfilSource_(data, name) {
  var source = data && data[name];
  if (!source || !source.sheet || !Array.isArray(source.headers)) throw new Error('FONTE_INDISPONIVEL_' + name);
  return source;
}

function corePerfilFindById_(source, idPessoa) {
  var records = (source && source.records) || [];
  for (var i = 0; i < records.length; i++) {
    if (String(records[i].ID_PESSOA || '').trim() === String(idPessoa || '').trim()) return records[i];
  }
  return null;
}

function corePerfilWriteRecord_(source, record, deps) {
  if (deps && typeof deps.writeRecord === 'function') return deps.writeRecord(source, record);
  source.sheet.getRange(record.__rowNumber, 1, 1, source.headers.length)
    .setValues([core_buildRowFromObjectByHeaders_(source.headers, record)]);
}

function corePerfilAppendRecord_(source, record, deps) {
  if (deps && typeof deps.appendRecord === 'function') return deps.appendRecord(source, record);
  source.sheet.appendRow(core_buildRowFromObjectByHeaders_(source.headers, record));
}

function corePerfilWithLock_(key, fn, deps) {
  if (deps && typeof deps.withLock === 'function') return deps.withLock(key, fn);
  return core_withLock_(key, fn, 20000);
}

function corePerfilHasHeader_(source, header) {
  return (source.headers || []).map(core_normalizeHeader_).indexOf(core_normalizeHeader_(header)) >= 0;
}

function corePerfilDirectPayload_(payload) {
  var source = payload && typeof payload === 'object' ? payload : {};
  var fieldInput = source.campos && typeof source.campos === 'object' ? source.campos : source;
  var changes = [];
  Object.keys(CORE_PERFIL_DIRECT_ALIASES).forEach(function(alias) {
    if (!Object.prototype.hasOwnProperty.call(fieldInput, alias)) return;
    var field = CORE_PERFIL_DIRECT_ALIASES[alias];
    if (changes.some(function(item) { return item.field === field; })) return;
    var cfg = CORE_PERFIL_DIRECT_FIELDS[field];
    var value = fieldInput[alias];
    if (cfg.type === 'PHONE') value = corePerfilNormalizePhone_(value);
    else if (cfg.type === 'INSTAGRAM') value = corePerfilNormalizeInstagram_(value);
    else if (cfg.type === 'UF') value = corePerfilNormalizeUf_(value);
    else if (cfg.type === 'CITY') value = corePerfilNormalizeTextField_(value, 120, 'CIDADE_ORIGEM_MUITO_LONGA');
    else if (cfg.type === 'SUMMARY') value = corePerfilNormalizeTextField_(value, CORE_PERFIL_MAX_RESUMO, 'RESUMO_ACADEMICO_MUITO_LONGO');
    changes.push({ field: field, source: cfg.source, header: cfg.header, value: value });
  });

  var links = Array.isArray(source.linksPerfis) ? source.linksPerfis : (Array.isArray(source.links) ? source.links : []);
  links.forEach(function(link) {
    var type = corePerfilNormalizeLinkType_(link && (link.tipo || link.type));
    if (changes.some(function(item) { return item.field === 'LINK_' + type; })) throw new Error('LINK_DUPLICADO_NO_PAYLOAD');
    changes.push({
      field: 'LINK_' + type,
      source: 'PESSOAS_V2_LINKS_PERFIS',
      linkType: type,
      value: corePerfilNormalizeUrl_(type, link && link.url),
      label: corePerfilNormalizeTextField_(link && link.rotulo, 100, 'ROTULO_LINK_MUITO_LONGO') || CORE_PESSOAS_V2_LINKS_PERFIS_LABELS[type]
    });
  });
  if (!changes.length) throw new Error('NENHUM_CAMPO_EDITAVEL_INFORMADO');
  return changes;
}

function corePerfilIdempotencyKey_(payload) {
  var key = String(payload && (payload.chaveIdempotencia || payload.idempotencyKey) || '').trim();
  if (!/^[A-Za-z0-9._:-]{8,120}$/.test(key)) throw new Error('CHAVE_IDEMPOTENCIA_INVALIDA');
  return key;
}

function corePerfilFindIdempotent_(requestSource, idPessoa, type, key) {
  return ((requestSource && requestSource.records) || []).filter(function(record) {
    return String(record.ID_PESSOA || '').trim() === idPessoa &&
      corePerfilNormalizeToken_(record.TIPO_SOLICITACAO) === type &&
      String(record.CHAVE_IDEMPOTENCIA || '').trim() === key;
  });
}

function corePerfilBuildAuditRow_(params, deps) {
  var now = params.now || corePerfilNow_(deps);
  return {
    ID_SOLICITACAO: params.requestId,
    ID_PESSOA: params.idPessoa,
    TIPO_SOLICITACAO: params.type,
    CAMPO: params.field,
    VALOR_ATUAL_MASCARADO: params.currentDisplay,
    VALOR_ATUAL_HASH: corePerfilHash_(String(params.currentValue == null ? '' : params.currentValue), deps),
    VALOR_SOLICITADO: params.requestedValue,
    JUSTIFICATIVA: params.justification || '',
    STATUS: params.status,
    SOLICITADO_EM: now,
    SOLICITADO_POR: params.actor,
    ANALISADO_EM: params.analyzedAt || '',
    ANALISADO_POR: params.analyzedBy || '',
    DECISAO: params.decision || '',
    MOTIVO_DECISAO: params.decisionReason || '',
    APLICADO_EM: params.appliedAt || '',
    ID_LOG: params.logId || '',
    CHAVE_IDEMPOTENCIA: params.idempotencyKey,
    ATIVO: 'SIM',
    CRIADO_EM: now,
    ATUALIZADO_EM: now
  };
}

function corePerfilApplyDirectChanges_(data, session, changes, key, deps) {
  var idPessoa = String(session.idPessoa).trim();
  var requests = corePerfilSource_(data, CORE_PERFIL_SOLICITACOES_SHEET);
  var replay = corePerfilFindIdempotent_(requests, idPessoa, 'ALTERACAO_DIRETA', key);
  if (replay.length) {
    return corePerfilEnvelopeOk_({
      idempotente: true,
      camposAtualizados: Object.freeze(replay.map(function(row) { return String(row.CAMPO || ''); }))
    });
  }

  var now = corePerfilNow_(deps);
  var bySource = {};
  var changed = [];
  changes.forEach(function(change) {
    if (change.source === 'PESSOAS_V2_LINKS_PERFIS') return;
    var source = corePerfilSource_(data, change.source);
    var record = bySource[change.source] || corePerfilFindById_(source, idPessoa);
    if (!record) throw new Error('REGISTRO_PESSOA_NAO_ENCONTRADO_' + change.source);
    if (!corePerfilHasHeader_(source, change.header) || !corePerfilHasHeader_(source, 'ATUALIZADO_EM')) {
      throw new Error('SCHEMA_INCOMPATIVEL_' + change.source);
    }
    if (!bySource[change.source]) bySource[change.source] = Object.assign({}, record);
    var current = String(record[change.header] == null ? '' : record[change.header]).trim();
    if (current === String(change.value)) return;
    bySource[change.source][change.header] = change.value;
    bySource[change.source].ATUALIZADO_EM = now;
    changed.push({ field: change.field, current: current, requested: change.value });
  });

  changes.filter(function(change) { return change.source === 'PESSOAS_V2_LINKS_PERFIS'; }).forEach(function(change) {
    var source = corePerfilSource_(data, 'PESSOAS_V2_LINKS_PERFIS');
    var records = source.records || [];
    var existing = null;
    for (var i = records.length - 1; i >= 0; i--) {
      if (String(records[i].ID_PESSOA || '').trim() !== idPessoa) continue;
      if (core_domainsV2NormalizeLinkPerfilType_(records[i].TIPO_LINK) !== change.linkType) continue;
      if (core_domainsV2AuditIsSim_(records[i].ATIVO)) { existing = records[i]; break; }
      if (!existing) existing = records[i];
    }
    var currentUrl = existing ? String(existing.URL || '').trim() : '';
    var currentlyActive = !!existing && core_domainsV2AuditIsSim_(existing.ATIVO);
    if (currentUrl === change.value && currentlyActive === !!change.value) return;
    if (existing) {
      var updated = Object.assign({}, existing, {
        URL: change.value || currentUrl,
        ROTULO: change.label,
        ATIVO: change.value ? 'SIM' : 'NAO',
        ATUALIZADO_EM: now
      });
      corePerfilWriteRecord_(source, updated, deps);
    } else if (change.value) {
      corePerfilAppendRecord_(source, {
        ID_LINK: corePerfilUuid_('LNK-', deps),
        ID_PESSOA: idPessoa,
        TIPO_LINK: change.linkType,
        URL: change.value,
        ROTULO: change.label,
        PUBLICAVEL: 'NAO',
        VISIVEL_PORTAL: 'SIM',
        FONTE: 'PORTAL_HOMOLOG',
        VALIDADO_EM: '',
        ATIVO: 'SIM',
        CRIADO_EM: now,
        ATUALIZADO_EM: now,
        OBS: ''
      }, deps);
    }
    changed.push({ field: change.field, current: currentUrl, requested: change.value });
  });

  Object.keys(bySource).forEach(function(name) {
    corePerfilWriteRecord_(corePerfilSource_(data, name), bySource[name], deps);
  });

  changed.forEach(function(change) {
    var requestId = 'ALT-' + corePerfilHash_(idPessoa + '|' + key + '|' + change.field, deps).slice(0, 24).toUpperCase();
    corePerfilAppendRecord_(requests, corePerfilBuildAuditRow_({
      requestId: requestId,
      idPessoa: idPessoa,
      type: 'ALTERACAO_DIRETA',
      field: change.field,
      currentDisplay: change.current,
      currentValue: change.current,
      requestedValue: change.requested,
      status: 'APLICADA',
      actor: session.email || idPessoa,
      appliedAt: now,
      logId: 'LOG-' + corePerfilHash_(requestId, deps).slice(0, 24).toUpperCase(),
      idempotencyKey: key,
      now: now
    }, deps), deps);
  });
  corePerfilSafeLog_('DIRECT_UPDATE', { ok: true, changedCount: changed.length });
  return corePerfilEnvelopeOk_({
    idempotente: false,
    semAlteracoes: changed.length === 0,
    camposAtualizados: Object.freeze(changed.map(function(item) { return item.field; }))
  });
}

function core_atualizarMeuPerfilParaPortal_(payload, contexto, options) {
  var deps = options && options.deps ? options.deps : {};
  try {
    corePerfilAssertNoTargetIdentity_(payload);
    var auth = corePerfilAuthorizeOwn_(contexto, deps);
    if (!auth.ok) return auth.response;
    var key = corePerfilIdempotencyKey_(payload);
    var changes = corePerfilDirectPayload_(payload);
    if (payload && payload.dryRun === true) {
      return corePerfilEnvelopeOk_({ dryRun: true, camposValidados: Object.freeze(changes.map(function(item) { return item.field; })) });
    }
    corePerfilAssertHomologContext_(contexto, deps);
    return corePerfilWithLock_('PERFIL_DIRETO_' + auth.session.idPessoa, function() {
      return corePerfilApplyDirectChanges_(corePerfilOpenPessoas_(deps), auth.session, changes, key, deps);
    }, deps);
  } catch (err) {
    corePerfilSafeLog_('DIRECT_UPDATE_ERROR', { ok: false, code: err && err.message });
    return corePerfilEnvelopeError_(err && err.message, 'Dados invalidos ou operacao cadastral indisponivel.');
  }
}

function corePerfilSensitiveCurrent_(data, idPessoa, field) {
  var cfg = CORE_PERFIL_SENSITIVE_FIELDS[field];
  var source = corePerfilSource_(data, cfg.source);
  var record = corePerfilFindById_(source, idPessoa);
  if (!record) throw new Error('PESSOA_NAO_ENCONTRADA');
  if (!corePerfilHasHeader_(source, cfg.header)) throw new Error('CAMPO_NAO_DISPONIVEL_NA_FONTE');
  return { source: source, record: record, cfg: cfg, value: String(record[cfg.header] == null ? '' : record[cfg.header]).trim() };
}

function core_solicitarCorrecaoMeuPerfilParaPortal_(payload, contexto, options) {
  var deps = options && options.deps ? options.deps : {};
  try {
    corePerfilAssertNoTargetIdentity_(payload);
    var auth = corePerfilAuthorizeOwn_(contexto, deps);
    if (!auth.ok) return auth.response;
    var field = corePerfilNormalizeToken_(payload && (payload.campo || payload.field));
    if (!CORE_PERFIL_SENSITIVE_FIELDS[field]) throw new Error('CAMPO_SENSIVEL_NAO_PERMITIDO');
    var requested = corePerfilNormalizeSensitiveValue_(field, payload && (payload.valorSolicitado != null ? payload.valorSolicitado : payload.valor));
    var justification = corePerfilNormalizeTextField_(payload && payload.justificativa, 1000, 'JUSTIFICATIVA_MUITO_LONGA');
    if (justification.length < 20) throw new Error('JUSTIFICATIVA_OBRIGATORIA');
    var key = corePerfilIdempotencyKey_(payload);
    if (payload && payload.dryRun === true) {
      return corePerfilEnvelopeOk_({ dryRun: true, campo: field, valorSolicitadoMascarado: corePerfilMaskValue_(field, requested) });
    }
    corePerfilAssertHomologContext_(contexto, deps);
    return corePerfilWithLock_('PERFIL_SOLICITAR_' + auth.session.idPessoa, function() {
      var data = corePerfilOpenPessoas_(deps);
      var requestSource = corePerfilSource_(data, CORE_PERFIL_SOLICITACOES_SHEET);
      var idPessoa = String(auth.session.idPessoa).trim();
      var replay = corePerfilFindIdempotent_(requestSource, idPessoa, 'CORRECAO_SENSIVEL', key);
      if (replay.length) {
        return corePerfilEnvelopeOk_({ idSolicitacao: replay[0].ID_SOLICITACAO, status: replay[0].STATUS, idempotente: true });
      }
      var duplicate = (requestSource.records || []).some(function(record) {
        return String(record.ID_PESSOA || '').trim() === idPessoa &&
          corePerfilNormalizeToken_(record.CAMPO) === field &&
          ['PENDENTE', 'EM_ANALISE'].indexOf(corePerfilNormalizeToken_(record.STATUS)) >= 0 &&
          core_domainsV2AuditIsSim_(record.ATIVO);
      });
      if (duplicate) return corePerfilEnvelopeError_('SOLICITACAO_DUPLICADA', 'Ja existe solicitacao pendente ou em analise para este campo.');
      var current = corePerfilSensitiveCurrent_(data, idPessoa, field).value;
      if (String(current) === String(requested)) return corePerfilEnvelopeError_('VALOR_SEM_ALTERACAO', 'O valor solicitado ja consta na fonte oficial.');
      var requestId = corePerfilUuid_('SAC-', deps);
      corePerfilAppendRecord_(requestSource, corePerfilBuildAuditRow_({
        requestId: requestId,
        idPessoa: idPessoa,
        type: 'CORRECAO_SENSIVEL',
        field: field,
        currentDisplay: corePerfilMaskValue_(field, current),
        currentValue: current,
        requestedValue: requested,
        justification: justification,
        status: 'PENDENTE',
        actor: auth.session.email || idPessoa,
        idempotencyKey: key
      }, deps), deps);
      corePerfilSafeLog_('SENSITIVE_REQUEST', { ok: true, field: field, status: 'PENDENTE', requestId: requestId });
      return corePerfilEnvelopeOk_({ idSolicitacao: requestId, campo: field, status: 'PENDENTE', idempotente: false });
    }, deps);
  } catch (err) {
    corePerfilSafeLog_('SENSITIVE_REQUEST_ERROR', { ok: false, code: err && err.message });
    return corePerfilEnvelopeError_(err && err.message, 'Solicitacao invalida ou indisponivel.');
  }
}

function corePerfilMapOwnRequest_(record) {
  var field = corePerfilNormalizeToken_(record.CAMPO);
  return Object.freeze({
    id: String(record.ID_SOLICITACAO || ''),
    campo: field,
    data: record.SOLICITADO_EM || record.CRIADO_EM || '',
    status: corePerfilNormalizeToken_(record.STATUS),
    valorSolicitadoMascarado: corePerfilMaskValue_(field, record.VALOR_SOLICITADO),
    decisao: String(record.DECISAO || ''),
    motivoDecisao: corePerfilRedactSensitiveText_(record.MOTIVO_DECISAO),
    analisadoEm: record.ANALISADO_EM || '',
    aplicadoEm: record.APLICADO_EM || ''
  });
}

function core_listarMinhasSolicitacoesCadastraisPortal_(contexto, options) {
  var deps = options && options.deps ? options.deps : {};
  try {
    var auth = corePerfilAuthorizeOwn_(contexto, deps);
    if (!auth.ok) return auth.response;
    var data = corePerfilOpenPessoas_(deps);
    var source = corePerfilSource_(data, CORE_PERFIL_SOLICITACOES_SHEET);
    var id = String(auth.session.idPessoa).trim();
    var items = (source.records || []).filter(function(record) {
      return String(record.ID_PESSOA || '').trim() === id &&
        corePerfilNormalizeToken_(record.TIPO_SOLICITACAO) === 'CORRECAO_SENSIVEL' &&
        core_domainsV2AuditIsSim_(record.ATIVO);
    }).map(corePerfilMapOwnRequest_).sort(function(a, b) {
      return new Date(b.data).getTime() - new Date(a.data).getTime();
    });
    return corePerfilEnvelopeOk_({ solicitacoes: Object.freeze(items), total: items.length });
  } catch (err) {
    return corePerfilEnvelopeError_('SOLICITACOES_INDISPONIVEIS', 'Nao foi possivel consultar suas solicitacoes.');
  }
}

function corePerfilFindRequestById_(source, requestId) {
  var id = String(requestId || '').trim();
  var records = source.records || [];
  for (var i = 0; i < records.length; i++) {
    if (String(records[i].ID_SOLICITACAO || '').trim() === id) return records[i];
  }
  return null;
}

function core_listarSolicitacoesCadastraisAdministracaoPortal_(filtros, contexto, options) {
  var deps = options && options.deps ? options.deps : {};
  try {
    var auth = corePerfilAuthorizeAdmin_(contexto, deps);
    if (!auth.ok) return auth.response;
    var opts = filtros && typeof filtros === 'object' ? filtros : {};
    var status = corePerfilNormalizeToken_(opts.status);
    var field = corePerfilNormalizeToken_(opts.campo);
    var pageSize = Math.min(Math.max(Number(opts.pageSize || 50), 1), 100);
    var page = Math.max(Number(opts.pagina || 1), 1);
    var source = corePerfilSource_(corePerfilOpenPessoas_(deps), CORE_PERFIL_SOLICITACOES_SHEET);
    var items = (source.records || []).filter(function(record) {
      if (corePerfilNormalizeToken_(record.TIPO_SOLICITACAO) !== 'CORRECAO_SENSIVEL') return false;
      if (!core_domainsV2AuditIsSim_(record.ATIVO)) return false;
      if (status && corePerfilNormalizeToken_(record.STATUS) !== status) return false;
      if (field && corePerfilNormalizeToken_(record.CAMPO) !== field) return false;
      return true;
    }).map(function(record) {
      var requestField = corePerfilNormalizeToken_(record.CAMPO);
      return Object.freeze({
        id: String(record.ID_SOLICITACAO || ''),
        campo: requestField,
        solicitadoEm: record.SOLICITADO_EM || '',
        status: corePerfilNormalizeToken_(record.STATUS),
        valorAtualMascarado: String(record.VALOR_ATUAL_MASCARADO || ''),
        valorSolicitadoMascarado: corePerfilMaskValue_(requestField, record.VALOR_SOLICITADO),
        justificativa: corePerfilRedactSensitiveText_(record.JUSTIFICATIVA).slice(0, 1000),
        decisao: String(record.DECISAO || ''),
        motivoDecisao: corePerfilRedactSensitiveText_(record.MOTIVO_DECISAO),
        analisadoEm: record.ANALISADO_EM || '',
        aplicadoEm: record.APLICADO_EM || ''
      });
    });
    var total = items.length;
    var start = (page - 1) * pageSize;
    return corePerfilEnvelopeOk_({
      solicitacoes: Object.freeze(items.slice(start, start + pageSize)),
      paginacao: Object.freeze({ pagina: page, pageSize: pageSize, totalItens: total, totalPaginas: Math.max(Math.ceil(total / pageSize), 1) })
    });
  } catch (err) {
    return corePerfilEnvelopeError_('SOLICITACOES_ADMIN_INDISPONIVEIS', 'Nao foi possivel consultar solicitacoes cadastrais.');
  }
}

function core_analisarSolicitacaoCadastralPortal_(payload, contexto, options) {
  var deps = options && options.deps ? options.deps : {};
  try {
    var auth = corePerfilAuthorizeAdmin_(contexto, deps);
    if (!auth.ok) return auth.response;
    var action = corePerfilNormalizeToken_(payload && (payload.acao || payload.status));
    if (['EM_ANALISE', 'COMPLEMENTO_SOLICITADO', 'APROVADA', 'INDEFERIDA'].indexOf(action) < 0) throw new Error('ACAO_ANALISE_INVALIDA');
    var reason = corePerfilRedactSensitiveText_(corePerfilNormalizeTextField_(payload && (payload.motivo || payload.motivoDecisao), 1000, 'MOTIVO_DECISAO_MUITO_LONGO'));
    if (['COMPLEMENTO_SOLICITADO', 'INDEFERIDA'].indexOf(action) >= 0 && reason.length < 10) throw new Error('MOTIVO_DECISAO_OBRIGATORIO');
    if (payload && payload.dryRun === true) return corePerfilEnvelopeOk_({ dryRun: true, statusDestino: action });
    corePerfilAssertHomologContext_(contexto, deps);
    return corePerfilWithLock_('PERFIL_ANALISAR_SOLICITACAO', function() {
      var source = corePerfilSource_(corePerfilOpenPessoas_(deps), CORE_PERFIL_SOLICITACOES_SHEET);
      var record = corePerfilFindRequestById_(source, payload && payload.idSolicitacao);
      if (!record || corePerfilNormalizeToken_(record.TIPO_SOLICITACAO) !== 'CORRECAO_SENSIVEL') throw new Error('SOLICITACAO_NAO_ENCONTRADA');
      var currentStatus = corePerfilNormalizeToken_(record.STATUS);
      if (currentStatus === action) return corePerfilEnvelopeOk_({ idSolicitacao: record.ID_SOLICITACAO, status: action, idempotente: true });
      var allowedFrom = {
        PENDENTE: ['EM_ANALISE', 'COMPLEMENTO_SOLICITADO', 'APROVADA', 'INDEFERIDA'],
        EM_ANALISE: ['COMPLEMENTO_SOLICITADO', 'APROVADA', 'INDEFERIDA'],
        COMPLEMENTO_SOLICITADO: ['EM_ANALISE', 'APROVADA', 'INDEFERIDA'],
        ERRO_APLICACAO: ['APROVADA', 'INDEFERIDA']
      };
      if (!allowedFrom[currentStatus] || allowedFrom[currentStatus].indexOf(action) < 0) throw new Error('TRANSICAO_STATUS_INVALIDA');
      var now = corePerfilNow_(deps);
      var updated = Object.assign({}, record, {
        STATUS: action,
        ANALISADO_EM: now,
        ANALISADO_POR: auth.session.email || auth.session.idPessoa,
        DECISAO: action,
        MOTIVO_DECISAO: reason,
        ATUALIZADO_EM: now
      });
      corePerfilWriteRecord_(source, updated, deps);
      corePerfilSafeLog_('ADMIN_ANALYSIS', { ok: true, status: action, requestId: record.ID_SOLICITACAO });
      return corePerfilEnvelopeOk_({ idSolicitacao: record.ID_SOLICITACAO, status: action, idempotente: false });
    }, deps);
  } catch (err) {
    return corePerfilEnvelopeError_(err && err.message, 'Nao foi possivel analisar a solicitacao.');
  }
}

function corePerfilRecalculateViews_(idPessoa, deps) {
  if (deps && typeof deps.recalculateViews === 'function') return deps.recalculateViews(idPessoa);
  return coreRecalcularPessoasResumoOperacionalV2_({
    dryRun: false,
    confirmacao: 'RECALCULAR_PESSOAS_RESUMO_V2',
    idPessoa: idPessoa
  });
}

function core_aplicarSolicitacaoCadastralAprovadaPortal_(payload, contexto, options) {
  var deps = options && options.deps ? options.deps : {};
  try {
    var auth = corePerfilAuthorizeAdmin_(contexto, deps);
    if (!auth.ok) return auth.response;
    if (payload && payload.dryRun === true) return corePerfilEnvelopeOk_({ dryRun: true, idSolicitacao: String(payload.idSolicitacao || '') });
    corePerfilAssertHomologContext_(contexto, deps);
    return corePerfilWithLock_('PERFIL_APLICAR_SOLICITACAO', function() {
      var data = corePerfilOpenPessoas_(deps);
      var requestSource = corePerfilSource_(data, CORE_PERFIL_SOLICITACOES_SHEET);
      var record = corePerfilFindRequestById_(requestSource, payload && payload.idSolicitacao);
      if (!record || corePerfilNormalizeToken_(record.TIPO_SOLICITACAO) !== 'CORRECAO_SENSIVEL') throw new Error('SOLICITACAO_NAO_ENCONTRADA');
      var status = corePerfilNormalizeToken_(record.STATUS);
      if (status === 'APLICADA') return corePerfilEnvelopeOk_({ idSolicitacao: record.ID_SOLICITACAO, status: 'APLICADA', idempotente: true });
      if (status !== 'APROVADA') throw new Error('SOLICITACAO_NAO_APROVADA');
      var field = corePerfilNormalizeToken_(record.CAMPO);
      try {
        if (!CORE_PERFIL_SENSITIVE_FIELDS[field]) throw new Error('CAMPO_SENSIVEL_NAO_PERMITIDO');
        var requested = corePerfilNormalizeSensitiveValue_(field, record.VALOR_SOLICITADO);
        var current = corePerfilSensitiveCurrent_(data, String(record.ID_PESSOA || '').trim(), field);
        var currentHash = corePerfilHash_(current.value, deps);
        if (String(record.VALOR_ATUAL_HASH || '') !== currentHash && String(current.value) !== String(requested)) {
          throw new Error('VALOR_ATUAL_ALTERADO_INCOMPATIVEL');
        }
        if (!corePerfilHasHeader_(current.source, 'ATUALIZADO_EM')) throw new Error('SCHEMA_SEM_ATUALIZADO_EM');
        var now = corePerfilNow_(deps);
        var updatedSource = Object.assign({}, current.record);
        updatedSource[current.cfg.header] = requested;
        updatedSource.ATUALIZADO_EM = now;
        corePerfilWriteRecord_(current.source, updatedSource, deps);
        var viewResult = corePerfilRecalculateViews_(String(record.ID_PESSOA || '').trim(), deps);
        if (viewResult && viewResult.ok === false) throw new Error('ERRO_RECALCULO_VIEW');
        var updatedRequest = Object.assign({}, record, {
          STATUS: 'APLICADA',
          APLICADO_EM: now,
          ID_LOG: record.ID_LOG || ('LOG-' + corePerfilHash_(record.ID_SOLICITACAO + '|APLICADA', deps).slice(0, 24).toUpperCase()),
          ATUALIZADO_EM: now
        });
        corePerfilWriteRecord_(requestSource, updatedRequest, deps);
        corePerfilSafeLog_('ADMIN_APPLY', { ok: true, field: field, status: 'APLICADA', requestId: record.ID_SOLICITACAO });
        return corePerfilEnvelopeOk_({ idSolicitacao: record.ID_SOLICITACAO, campo: field, status: 'APLICADA', idempotente: false });
      } catch (applyError) {
        var failedAt = corePerfilNow_(deps);
        corePerfilWriteRecord_(requestSource, Object.assign({}, record, {
          STATUS: 'ERRO_APLICACAO',
          DECISAO: 'ERRO_APLICACAO',
          MOTIVO_DECISAO: 'Nao foi possivel aplicar automaticamente. Revise o cadastro e tente novamente.',
          ATUALIZADO_EM: failedAt
        }), deps);
        throw applyError;
      }
    }, deps);
  } catch (err) {
    corePerfilSafeLog_('ADMIN_APPLY_ERROR', { ok: false, code: err && err.message });
    return corePerfilEnvelopeError_(err && err.message, 'Nao foi possivel aplicar a solicitacao aprovada.');
  }
}

function corePerfilRegistryRowValues_(headers, values) {
  return (headers || []).map(function(header) {
    var key = core_normalizeHeader_(header);
    return Object.prototype.hasOwnProperty.call(values, key) ? values[key] : '';
  });
}

function corePerfilEnsureRegistryDev_(spreadsheetId, dryRun) {
  var registry = core_openSpreadsheetById_(CORE_REGISTRY_SPREADSHEET_ID).getSheetByName(CORE_REGISTRY_SHEET_NAME);
  if (!registry) throw new Error('REGISTRY_SHEET_INDISPONIVEL');
  var data = core_readSheetData_(registry, { headerRow: 1 });
  var headers = data.headers;
  var headerMap = core_buildHeaderIndexMap_(headers, { normalize: true, oneBased: false, keepFirst: true });
  ['KEY', 'SPREADSHEET_ID', 'SHEET_NAME', 'ATIVO', 'AMBIENTE'].forEach(function(header) {
    if (core_findHeaderIndex_(headerMap, header, { notFoundValue: -1 }) < 0) throw new Error('REGISTRY_HEADER_AUSENTE_' + header);
  });
  var existing = null;
  for (var i = 0; i < data.rows.length; i++) {
    var row = core_rowToObject_(headers, data.rows[i]);
    if (corePerfilNormalizeToken_(row.KEY) === CORE_PERFIL_SOLICITACOES_KEY && corePerfilNormalizeToken_(row.AMBIENTE) === 'DEV') {
      existing = row;
      break;
    }
  }
  if (existing) {
    if (String(existing.SPREADSHEET_ID || '').trim() !== spreadsheetId || String(existing.SHEET_NAME || '').trim() !== CORE_PERFIL_SOLICITACOES_SHEET) {
      throw new Error('REGISTRY_DEV_CONFLITANTE');
    }
    if (!dryRun) core_registryCacheClear_();
    return { existed: true, created: false, lineCompatible: true };
  }
  if (!dryRun) {
    var values = {};
    values[core_normalizeHeader_('KEY')] = CORE_PERFIL_SOLICITACOES_KEY;
    values[core_normalizeHeader_('SPREADSHEET_ID')] = spreadsheetId;
    values[core_normalizeHeader_('SHEET_NAME')] = CORE_PERFIL_SOLICITACOES_SHEET;
    values[core_normalizeHeader_('ATIVO')] = 'SIM';
    values[core_normalizeHeader_('AMBIENTE')] = 'DEV';
    values[core_normalizeHeader_('DISPLAY_NAME')] = 'Solicitacoes de atualizacao cadastral Pessoas V2';
    values[core_normalizeHeader_('TYPE')] = 'FONTE_EVENTOS';
    values[core_normalizeHeader_('NOTAS')] = 'Uso exclusivo DEV/HOMOLOG; nao cadastrar como PROD nesta etapa.';
    registry.appendRow(corePerfilRegistryRowValues_(headers, values));
    core_registryCacheClear_();
  }
  return { existed: false, created: !dryRun, planned: dryRun };
}

function corePerfilEnsureAdminPermissionDev_(dryRun) {
  var meta = null;
  try {
    meta = corePerfilRegistryMetaDev_('PORTAL_PERMISSOES');
  } catch (missingDevPermissionSource) {
    return {
      available: false,
      skipped: true,
      reason: 'PORTAL_PERMISSOES_DEV_INDISPONIVEL',
      fallback: 'PERFIS_SECRETARIA_DIRETORIA_SOMENTE_HOMOLOG'
    };
  }
  var sheet = core_getSheetById_(meta.id, meta.sheet);
  var data = core_readSheetData_(sheet, { headerRow: 1 });
  var required = ['SECRETARIA', 'DIRETORIA'];
  var missing = required.filter(function(profile) {
    return !data.rows.some(function(row) {
      var record = core_rowToObject_(data.headers, row);
      return corePerfilNormalizeToken_(record.PERFIL_PORTAL) === profile &&
        corePortalNormalizePermission_(record.PERMISSAO) === CORE_PERFIL_ADMIN_PERMISSION &&
        core_domainsV2AuditIsSim_(record.ATIVO);
    });
  });
  if (!dryRun) {
    missing.forEach(function(profile) {
      sheet.appendRow(core_buildRowFromObjectByHeaders_(data.headers, {
        PERFIL_PORTAL: profile,
        PERMISSAO: CORE_PERFIL_ADMIN_PERMISSION,
        ATIVO: 'SIM'
      }));
    });
  }
  return { requiredProfiles: required, missingBefore: missing, created: dryRun ? [] : missing.slice(), planned: dryRun ? missing.slice() : [] };
}

function core_setupSolicitacoesAtualizacaoCadastralDev_(options) {
  var opts = options || {};
  var dryRun = opts.dryRun !== false;
  var env = String(opts.environment || 'DEV').trim().toUpperCase();
  var observedEnv = '';
  try {
    observedEnv = core_getCurrentEnv_();
  } catch (ignoredEnvError) {
    observedEnv = 'NAO_IDENTIFICADO';
  }
  var report = {
    ok: false,
    dryRun: dryRun,
    environment: env,
    scriptEnvironmentObserved: observedEnv,
    sheetName: CORE_PERFIL_SOLICITACOES_SHEET,
    registryKey: CORE_PERFIL_SOLICITACOES_KEY,
    productionRefused: env === 'PROD',
    createdSheet: false,
    addedHeaders: [],
    registry: null,
    adminPermission: null,
    diagnostics: []
  };
  if (env !== 'DEV') {
    report.errorCode = 'SETUP_RECUSADO_FORA_DEV';
    report.message = 'O setup cadastral aceita somente environment=DEV e nunca altera a Script Property GEAPA_ENV.';
    Logger.log('[geapa-core][perfil-cadastral][setup] ' + JSON.stringify(report));
    return report;
  }
  if (!dryRun && String(opts.confirmacao || '').trim() !== CORE_PERFIL_SETUP_CONFIRMATION) {
    report.errorCode = 'CONFIRMACAO_OBRIGATORIA';
    report.message = 'Informe confirmacao: ' + CORE_PERFIL_SETUP_CONFIRMATION;
    Logger.log('[geapa-core][perfil-cadastral][setup] ' + JSON.stringify(report));
    return report;
  }
  return corePerfilWithLock_('SETUP_SOLICITACOES_CADASTRAIS_DEV', function() {
    var baseMeta = corePerfilRegistryMetaDev_('PESSOAS_V2_BASE');
    var spreadsheet = core_openSpreadsheetById_(baseMeta.id);
    var definition = CORE_DOMAINS_V2_SCHEMAS.PESSOAS.filter(function(item) {
      return item.sheetName === CORE_PERFIL_SOLICITACOES_SHEET;
    })[0];
    var detailsDefinition = CORE_DOMAINS_V2_SCHEMAS.PESSOAS.filter(function(item) {
      return item.sheetName === 'MEMBROS_DETALHES';
    })[0];
    var baseDefinition = CORE_DOMAINS_V2_SCHEMAS.PESSOAS.filter(function(item) {
      return item.sheetName === 'PESSOAS_BASE';
    })[0];
    var existing = spreadsheet.getSheetByName(CORE_PERFIL_SOLICITACOES_SHEET);
    report.spreadsheetName = spreadsheet.getName();
    report.existedBefore = !!existing;
    if (dryRun) {
      var existingHeaders = existing && existing.getLastColumn() > 0
        ? existing.getRange(1, 1, 1, existing.getLastColumn()).getDisplayValues()[0]
        : [];
      var existingMap = core_buildHeaderIndexMap_(existingHeaders, { normalize: true, oneBased: false, keepFirst: true });
      report.addedHeaders = definition.headers.filter(function(header) {
        return core_findHeaderIndex_(existingMap, header, { notFoundValue: -1 }) < 0;
      });
    } else {
      var resolved = core_getOrCreateDomainsV2Sheet_(spreadsheet, definition.sheetName, CORE_DOMAINS_V2_SCHEMAS.PESSOAS.indexOf(definition));
      report.createdSheet = resolved.created;
      var headerResult = core_ensureDomainsV2Headers_(resolved.sheet, definition.headers, []);
      report.addedHeaders = headerResult.addedHeaders;
      report.diagnostics = core_applyDomainsV2SheetUx_(resolved.sheet);
      core_applyDropdownValidationByHeader_(resolved.sheet, {
        STATUS: { values: CORE_PERFIL_STATUS, allowInvalid: false, helpText: 'Use somente status cadastrais permitidos.' },
        ATIVO: { values: ['SIM', 'NAO'], allowInvalid: false }
      }, 1, {});
      var detailsSheet = spreadsheet.getSheetByName(detailsDefinition.sheetName);
      if (!detailsSheet) throw new Error('MEMBROS_DETALHES_INDISPONIVEL');
      core_ensureDomainsV2Headers_(detailsSheet, detailsDefinition.headers, []);
      var baseSheet = spreadsheet.getSheetByName(baseDefinition.sheetName);
      if (!baseSheet) throw new Error('PESSOAS_BASE_INDISPONIVEL');
      core_ensureDomainsV2Headers_(baseSheet, baseDefinition.headers, []);
    }
    report.registry = corePerfilEnsureRegistryDev_(baseMeta.id, dryRun);
    report.adminPermission = corePerfilEnsureAdminPermissionDev_(dryRun);
    report.ok = true;
    report.message = dryRun ? 'Dry-run concluido sem escrita.' : 'Setup DEV/HOMOLOG concluido.';
    Logger.log('[geapa-core][perfil-cadastral][setup] ' + JSON.stringify(report));
    return report;
  }, opts.deps || {});
}

function core_setupSolicitacoesAtualizacaoCadastralDevReal_() {
  return core_setupSolicitacoesAtualizacaoCadastralDev_({
    dryRun: false,
    confirmacao: CORE_PERFIL_SETUP_CONFIRMATION
  });
}
