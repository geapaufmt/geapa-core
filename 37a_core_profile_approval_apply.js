/**
 * Aprovacao e aplicacao coordenadas de correcoes cadastrais sensiveis.
 *
 * Regra canonica de identificadores:
 * - EMAIL e RGA possuem exatamente um registro ATIVO=SIM e PRINCIPAL=SIM por pessoa;
 * - o identificador substituido e preservado com ATIVO=NAO e PRINCIPAL=NAO;
 * - CPF permanece somente em PESSOAS_BASE enquanto nao houver contrato oficial
 *   que o reconheca como identificador em PESSOAS_IDENTIFICADORES.
 */

var CORE_PERFIL_APPLY_SOURCES = Object.freeze([
  CORE_PERFIL_SOLICITACOES_SHEET,
  'PESSOAS_BASE',
  'PESSOAS_IDENTIFICADORES',
  'MEMBROS_DETALHES',
  'PESSOAS_RESUMO_OPERACIONAL'
]);

function corePerfilCloneRecord_(record) {
  return Object.keys(record || {}).reduce(function(out, key) {
    out[key] = record[key];
    return out;
  }, {});
}

function corePerfilRequireHeaders_(source, headers) {
  (headers || []).forEach(function(header) {
    if (!corePerfilHasHeader_(source, header)) {
      throw new Error('SCHEMA_HEADER_AUSENTE_' + source.name + '_' + header);
    }
  });
}

function corePerfilIdentifierTypeForField_(field) {
  if (field === 'EMAIL_PRINCIPAL') return 'EMAIL';
  if (field === 'RGA') return 'RGA';
  return '';
}

function corePerfilNormalizeIdentifierValue_(type, value) {
  if (type === 'EMAIL') return corePerfilNormalizeSensitiveValue_('EMAIL_PRINCIPAL', value);
  if (type === 'RGA') return corePerfilNormalizeSensitiveValue_('RGA', value);
  return String(value == null ? '' : value).trim();
}

function corePerfilIdentifierRecords_(source, idPessoa, type) {
  return (source.records || []).filter(function(record) {
    return String(record.ID_PESSOA || '').trim() === idPessoa &&
      corePerfilNormalizeToken_(record.TIPO_IDENTIFICADOR) === type;
  });
}

function corePerfilFindIdentifierRecord_(records, type, value) {
  var normalized = corePerfilNormalizeIdentifierValue_(type, value);
  for (var i = records.length - 1; i >= 0; i--) {
    if (corePerfilNormalizeIdentifierValue_(type, records[i].VALOR_IDENTIFICADOR) === normalized) {
      return records[i];
    }
  }
  return null;
}

function corePerfilAssertIdentifierAvailable_(data, idPessoa, type, value) {
  var normalized = corePerfilNormalizeIdentifierValue_(type, value);
  var base = corePerfilSource_(data, 'PESSOAS_BASE');
  var details = corePerfilSource_(data, 'MEMBROS_DETALHES');
  var identifiers = corePerfilSource_(data, 'PESSOAS_IDENTIFICADORES');

  if (type === 'EMAIL') {
    (base.records || []).forEach(function(record) {
      if (String(record.ID_PESSOA || '').trim() === idPessoa) return;
      if (corePerfilNormalizeIdentifierValue_('EMAIL', record.EMAIL_PRINCIPAL) === normalized) {
        throw new Error('EMAIL_JA_VINCULADO_A_OUTRA_PESSOA');
      }
    });
  }
  if (type === 'RGA') {
    (details.records || []).forEach(function(record) {
      if (String(record.ID_PESSOA || '').trim() === idPessoa) return;
      if (corePerfilNormalizeIdentifierValue_('RGA', record.RGA) === normalized) {
        throw new Error('RGA_JA_VINCULADO_A_OUTRA_PESSOA');
      }
    });
  }
  (identifiers.records || []).forEach(function(record) {
    if (String(record.ID_PESSOA || '').trim() === idPessoa) return;
    if (corePerfilNormalizeToken_(record.TIPO_IDENTIFICADOR) !== type) return;
    if (!core_domainsV2AuditIsSim_(record.ATIVO)) return;
    if (corePerfilNormalizeIdentifierValue_(type, record.VALOR_IDENTIFICADOR) !== normalized) return;
    throw new Error(type === 'EMAIL'
      ? 'EMAIL_JA_VINCULADO_A_OUTRA_PESSOA'
      : 'RGA_JA_VINCULADO_A_OUTRA_PESSOA');
  });
}

function corePerfilMutationWrite_(source, record, mutations, deps) {
  var existing = (source.records || []).filter(function(item) {
    return Number(item.__rowNumber) === Number(record.__rowNumber);
  })[0];
  var before = corePerfilCloneRecord_(existing || record);
  mutations.push({ type: 'WRITE', source: source, before: before });
  corePerfilWriteRecord_(source, record, deps);
}

function corePerfilMutationAppend_(source, record, mutations, deps) {
  var created = corePerfilCloneRecord_(record);
  if (!created.__rowNumber) {
    created.__rowNumber = deps && deps.appendRecord
      ? (source.records || []).length + 2
      : Number(source.sheet.getLastRow()) + 1;
  }
  corePerfilAppendRecord_(source, created, deps);
  mutations.push({ type: 'APPEND', source: source, created: created });
  return created;
}

function corePerfilMutationDeleteAppended_(mutation, deps) {
  if (deps && typeof deps.deleteRecord === 'function') {
    deps.deleteRecord(mutation.source, mutation.created);
    return;
  }
  if (mutation.source.sheet && typeof mutation.source.sheet.deleteRow === 'function') {
    mutation.source.sheet.deleteRow(Number(mutation.created.__rowNumber));
    return;
  }
  var inactive = Object.assign({}, mutation.created, {
    ATIVO: 'NAO',
    PRINCIPAL: 'NAO',
    OBS: corePerfilAppendObservation_(mutation.created.OBS, 'Registro desativado por compensacao de aplicacao cadastral.')
  });
  corePerfilWriteRecord_(mutation.source, inactive, deps);
}

function corePerfilCompensate_(mutations, deps) {
  var compensated = [];
  var errors = [];
  for (var i = mutations.length - 1; i >= 0; i--) {
    var mutation = mutations[i];
    try {
      if (mutation.type === 'WRITE') {
        corePerfilWriteRecord_(mutation.source, mutation.before, deps);
      } else if (mutation.type === 'APPEND') {
        corePerfilMutationDeleteAppended_(mutation, deps);
      }
      compensated.push(mutation.type + ':' + mutation.source.name);
    } catch (err) {
      errors.push(String(err && err.message || 'COMPENSACAO_FALHOU').slice(0, 100));
    }
  }
  return { compensated: compensated, errors: errors, ok: errors.length === 0 };
}

function corePerfilAppendObservation_(current, message) {
  return [String(current || '').trim(), String(message || '').trim()].filter(String).join(' | ').slice(0, 1000);
}

function corePerfilPrepareIdentifier_(source, idPessoa, type, currentValue, requestedValue, now, mutations, deps) {
  corePerfilRequireHeaders_(source, [
    'ID_IDENTIFICADOR', 'ID_PESSOA', 'TIPO_IDENTIFICADOR',
    'VALOR_IDENTIFICADOR', 'PRINCIPAL', 'ATIVO', 'OBS'
  ]);
  var records = corePerfilIdentifierRecords_(source, idPessoa, type);
  var requested = corePerfilNormalizeIdentifierValue_(type, requestedValue);
  var current = corePerfilNormalizeIdentifierValue_(type, currentValue);
  var target = corePerfilFindIdentifierRecord_(records, type, requested);

  if (target) {
    corePerfilMutationWrite_(source, Object.assign({}, target, {
      VALOR_IDENTIFICADOR: requested,
      PRINCIPAL: 'SIM',
      ATIVO: 'SIM',
      OBS: corePerfilAppendObservation_(target.OBS, 'Reativado como identificador principal por correcao cadastral em ' + now.toISOString() + '.')
    }), mutations, deps);
  } else {
    target = corePerfilMutationAppend_(source, {
      ID_IDENTIFICADOR: corePerfilUuid_('IDN-', deps),
      ID_PESSOA: idPessoa,
      TIPO_IDENTIFICADOR: type,
      VALOR_IDENTIFICADOR: requested,
      PRINCIPAL: 'SIM',
      ATIVO: 'SIM',
      OBS: 'Criado como identificador principal por correcao cadastral em ' + now.toISOString() + '.'
    }, mutations, deps);
  }

  records.forEach(function(record) {
    if (String(record.ID_IDENTIFICADOR || '') === String(target.ID_IDENTIFICADOR || '')) return;
    var value = corePerfilNormalizeIdentifierValue_(type, record.VALOR_IDENTIFICADOR);
    var isOld = current && value === current;
    var needsWrite = core_domainsV2AuditIsSim_(record.ATIVO) ||
      core_domainsV2AuditIsSim_(record.PRINCIPAL);
    if (!needsWrite) return;
    corePerfilMutationWrite_(source, Object.assign({}, record, {
      PRINCIPAL: 'NAO',
      ATIVO: 'NAO',
      OBS: corePerfilAppendObservation_(record.OBS, (isOld ? 'Substituido' : 'Desativado') + ' por correcao cadastral em ' + now.toISOString() + '.')
    }), mutations, deps);
  });
  return target;
}

function corePerfilCacheDigestHex_(value) {
  var text = String(value == null ? '' : value);
  if (typeof Utilities !== 'undefined' && Utilities.computeDigest) {
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
    return digest.map(function(byte) {
      var normalized = byte < 0 ? byte + 256 : byte;
      return ('0' + normalized.toString(16)).slice(-2);
    }).join('');
  }
  return corePerfilHash_(text, {}).replace(/[^a-z0-9]/gi, '').toLowerCase();
}

function corePerfilInvalidateLocalIdentityCaches_(identifiers, environment, deps) {
  var values = [];
  var seen = {};
  (identifiers || []).forEach(function(value) {
    var normalized = String(value || '').trim().toLowerCase();
    if (!normalized || seen[normalized]) return;
    seen[normalized] = true;
    values.push(normalized);
  });
  if (deps && typeof deps.invalidateCaches === 'function') {
    return deps.invalidateCaches(values.slice(), environment);
  }
  var removed = [];
  try {
    var cache = CacheService.getScriptCache();
    values.forEach(function(value) {
      ['sessaoCoreV2', 'minhaSituacaoV2'].forEach(function(type) {
        var key = ['portal', environment, type, corePerfilCacheDigestHex_(value).slice(0, 32)].join(':');
        cache.remove(key);
        removed.push(type);
      });
    });
  } catch (ignored) {}
  return { identifiers: values.length, removed: removed.length };
}

function corePerfilSafeTechnicalCode_(error) {
  return String(error && error.message || 'ERRO_APLICACAO')
    .toUpperCase().replace(/[^A-Z0-9_:-]+/g, '_').slice(0, 80);
}

function corePerfilApplicationLogId_(record, deps) {
  return 'LOG-' + corePerfilHash_(String(record.ID_SOLICITACAO || '') + '|APLICADA', deps)
    .slice(0, 24).toUpperCase();
}

function corePerfilApplicationErrorLogId_(record, deps) {
  return 'ERR-' + corePerfilHash_(String(record.ID_SOLICITACAO || '') + '|ERRO_APLICACAO', deps)
    .slice(0, 24).toUpperCase();
}

function corePerfilExecuteApprovalApply_(payload, contexto, auth, deps, mode) {
  var environment = corePerfilAssertPortalContext_(contexto, deps);
  var data = corePerfilOpenPessoas_(contexto, deps, {
    forWrite: true,
    sources: CORE_PERFIL_APPLY_SOURCES
  });
  var requestSource = corePerfilSource_(data, CORE_PERFIL_SOLICITACOES_SHEET);
  var baseSource = corePerfilSource_(data, 'PESSOAS_BASE');
  var identifiersSource = corePerfilSource_(data, 'PESSOAS_IDENTIFICADORES');
  var detailsSource = corePerfilSource_(data, 'MEMBROS_DETALHES');
  var summarySource = corePerfilSource_(data, 'PESSOAS_RESUMO_OPERACIONAL');
  var record = corePerfilFindRequestById_(requestSource, payload && payload.idSolicitacao);
  if (!record || corePerfilNormalizeToken_(record.TIPO_SOLICITACAO) !== 'CORRECAO_SENSIVEL') {
    throw new Error('SOLICITACAO_NAO_ENCONTRADA');
  }
  var status = corePerfilNormalizeToken_(record.STATUS);
  if (status === 'APLICADA') {
    return corePerfilEnvelopeOk_({
      idSolicitacao: record.ID_SOLICITACAO,
      status: 'APLICADA',
      aplicadoEm: record.APLICADO_EM || '',
      idLog: record.ID_LOG || '',
      idempotente: true
    });
  }
  if (['INDEFERIDA', 'CANCELADA'].indexOf(status) >= 0) throw new Error('SOLICITACAO_TERMINAL');
  if (mode && mode.legacyOnly === true && status !== 'APROVADA') {
    throw new Error('SOLICITACAO_LEGADA_NAO_APROVADA');
  }
  if (['PENDENTE', 'EM_ANALISE', 'COMPLEMENTO_SOLICITADO', 'APROVADA', 'ERRO_APLICACAO'].indexOf(status) < 0) {
    throw new Error('TRANSICAO_STATUS_INVALIDA');
  }

  var field = corePerfilNormalizeToken_(record.CAMPO);
  if (!CORE_PERFIL_SENSITIVE_FIELDS[field]) throw new Error('CAMPO_SENSIVEL_NAO_PERMITIDO');
  corePerfilRequireHeaders_(requestSource, [
    'STATUS', 'ANALISADO_EM', 'ANALISADO_POR', 'DECISAO',
    'MOTIVO_DECISAO', 'APLICADO_EM', 'ID_LOG', 'ATUALIZADO_EM'
  ]);
  corePerfilRequireHeaders_(baseSource, ['ID_PESSOA', 'EMAIL_PRINCIPAL']);
  corePerfilRequireHeaders_(detailsSource, ['ID_PESSOA', 'RGA']);
  corePerfilRequireHeaders_(summarySource, ['ID_PESSOA']);

  var idPessoa = String(record.ID_PESSOA || '').trim();
  var requested = corePerfilNormalizeSensitiveValue_(field, record.VALOR_SOLICITADO);
  var current = corePerfilSensitiveCurrent_(data, idPessoa, field);
  var currentHash = corePerfilHash_(current.value, deps);
  if (String(record.VALOR_ATUAL_HASH || '') !== currentHash && String(current.value) !== String(requested)) {
    throw new Error('VALOR_ATUAL_ALTERADO_INCOMPATIVEL');
  }
  if (!corePerfilHasHeader_(current.source, 'ATUALIZADO_EM')) throw new Error('SCHEMA_SEM_ATUALIZADO_EM');

  var identifierType = corePerfilIdentifierTypeForField_(field);
  if (identifierType) corePerfilAssertIdentifierAvailable_(data, idPessoa, identifierType, requested);

  var now = corePerfilNow_(deps);
  var reason = corePerfilRedactSensitiveText_(corePerfilNormalizeTextField_(
    payload && (payload.motivo || payload.motivoDecisao), 1000, 'MOTIVO_DECISAO_MUITO_LONGO'
  ));
  var actor = String(auth && auth.session && (auth.session.email || auth.session.idPessoa) || 'REPARACAO_OPERACIONAL').trim();
  var mutations = [];
  var summaryRecord = corePerfilFindById_(summarySource, idPessoa);
  var summarySnapshot = summaryRecord ? corePerfilCloneRecord_(summaryRecord) : null;
  var wrote = false;

  try {
    if (identifierType) {
      corePerfilPrepareIdentifier_(identifiersSource, idPessoa, identifierType, current.value, requested, now, mutations, deps);
      wrote = true;
    }
    var sourceUpdated = Object.assign({}, current.record);
    sourceUpdated[current.cfg.header] = requested;
    sourceUpdated.ATUALIZADO_EM = now;
    corePerfilMutationWrite_(current.source, sourceUpdated, mutations, deps);
    wrote = true;

    var viewResult = corePerfilRecalculateViews_(idPessoa, deps);
    if (viewResult && viewResult.ok === false) throw new Error('ERRO_RECALCULO_VIEW');

    var cacheIdentifiers = [idPessoa];
    if (field === 'EMAIL_PRINCIPAL') cacheIdentifiers.push(current.value, requested);
    if (field === 'RGA') cacheIdentifiers.push(current.value, requested);
    var cacheResult = corePerfilInvalidateLocalIdentityCaches_(cacheIdentifiers, environment, deps);
    var logId = corePerfilApplicationLogId_(record, deps);
    var applied = Object.assign({}, record, {
      STATUS: 'APLICADA',
      ANALISADO_EM: record.ANALISADO_EM || now,
      ANALISADO_POR: record.ANALISADO_POR || actor,
      DECISAO: 'APROVADA',
      MOTIVO_DECISAO: reason || record.MOTIVO_DECISAO || '',
      APLICADO_EM: now,
      ID_LOG: logId,
      ATUALIZADO_EM: now
    });
    corePerfilWriteRecord_(requestSource, applied, deps);

    corePerfilSafeLog_('ADMIN_APPROVE_APPLY', {
      ok: true, field: field, status: 'APLICADA', requestId: record.ID_SOLICITACAO,
      actorHash: corePerfilHash_(actor, deps)
    });
    return corePerfilEnvelopeOk_({
      idSolicitacao: record.ID_SOLICITACAO,
      campo: field,
      status: 'APLICADA',
      aplicadoEm: now,
      analisadoPorMascarado: corePerfilMaskValue_('EMAIL_PRINCIPAL', actor),
      idLog: logId,
      idempotente: false,
      sessaoDeveSerRenovada: identifierType === 'EMAIL',
      _cacheInvalidation: {
        ambiente: environment,
        identificadores: cacheIdentifiers,
        resultadoCore: cacheResult || {}
      }
    });
  } catch (applyError) {
    var compensation = corePerfilCompensate_(mutations, deps);
    if (summarySnapshot) {
      try { corePerfilWriteRecord_(summarySource, summarySnapshot, deps); } catch (summaryRestoreError) {
        compensation.ok = false;
        compensation.errors.push('RESUMO_RESTORE_FALHOU');
      }
    } else if (wrote) {
      try { corePerfilRecalculateViews_(idPessoa, deps); } catch (recalcRestoreError) {
        compensation.ok = false;
        compensation.errors.push('RESUMO_RECALCULO_COMPENSACAO_FALHOU');
      }
    }
    if (wrote) {
      var failedAt = corePerfilNow_(deps);
      try {
        corePerfilWriteRecord_(requestSource, Object.assign({}, record, {
          STATUS: 'ERRO_APLICACAO',
          DECISAO: 'ERRO_APLICACAO',
          MOTIVO_DECISAO: 'Nao foi possivel aplicar a alteracao. Nenhum resultado conclusivo foi registrado. A operacao pode ser reprocessada.',
          ID_LOG: corePerfilApplicationErrorLogId_(record, deps),
          ATUALIZADO_EM: failedAt
        }), deps);
      } catch (requestFailure) {
        compensation.ok = false;
        compensation.errors.push('SOLICITACAO_ERRO_NAO_REGISTRADO');
      }
    }
    corePerfilSafeLog_('ADMIN_APPROVE_APPLY_ERROR', {
      ok: false,
      field: field,
      status: wrote ? 'ERRO_APLICACAO' : status,
      requestId: record.ID_SOLICITACAO,
      code: corePerfilSafeTechnicalCode_(applyError)
    });
    var wrapped = new Error(corePerfilSafeTechnicalCode_(applyError));
    wrapped.compensation = compensation;
    throw wrapped;
  }
}

function core_aprovarEAplicarSolicitacaoCadastralPortal_(payload, contexto, options) {
  var deps = options && options.deps ? options.deps : {};
  var mode = options && options.mode ? options.mode : {};
  try {
    var auth = corePerfilAuthorizeAdmin_(contexto, deps);
    if (!auth.ok) return auth.response;
    if (!mode.legacyOnly && (!payload || payload.confirmacao !== true)) {
      throw new Error('CONFIRMACAO_APROVAR_APLICAR_OBRIGATORIA');
    }
    if (payload && payload.dryRun === true) {
      return corePerfilEnvelopeOk_({
        dryRun: true,
        idSolicitacao: String(payload.idSolicitacao || ''),
        statusDestino: 'APLICADA'
      });
    }
    corePerfilAssertPortalContext_(contexto, deps);
    return corePerfilWithLock_('PERFIL_APROVAR_APLICAR_' + String(payload && payload.idSolicitacao || ''), function() {
      return corePerfilExecuteApprovalApply_(payload || {}, contexto, auth, deps, mode);
    }, deps);
  } catch (err) {
    var details = err && err.compensation ? {
      compensacaoOk: err.compensation.ok === true,
      etapasCompensadas: (err.compensation.compensated || []).slice(),
      errosCompensacao: (err.compensation.errors || []).slice()
    } : null;
    return corePerfilEnvelopeError_(
      err && err.message || 'ERRO_APLICACAO',
      'Nao foi possivel aplicar a alteracao. Nenhum resultado conclusivo foi registrado. A operacao pode ser reprocessada.',
      details
    );
  }
}

function corePerfilRepairReport_(idSolicitacao, data, deps) {
  var requestSource = corePerfilSource_(data, CORE_PERFIL_SOLICITACOES_SHEET);
  var record = corePerfilFindRequestById_(requestSource, idSolicitacao);
  if (!record) throw new Error('SOLICITACAO_NAO_ENCONTRADA');
  var field = corePerfilNormalizeToken_(record.CAMPO);
  var idPessoa = String(record.ID_PESSOA || '').trim();
  var current = corePerfilSensitiveCurrent_(data, idPessoa, field);
  var requested = corePerfilNormalizeSensitiveValue_(field, record.VALOR_SOLICITADO);
  var type = corePerfilIdentifierTypeForField_(field);
  var identifiers = type
    ? corePerfilIdentifierRecords_(corePerfilSource_(data, 'PESSOAS_IDENTIFICADORES'), idPessoa, type)
    : [];
  var duplicates = [];
  if (type) {
    try { corePerfilAssertIdentifierAvailable_(data, idPessoa, type, requested); }
    catch (err) { duplicates.push(String(err && err.message || 'IDENTIFICADOR_DUPLICADO')); }
  }
  var hashCompatible = String(record.VALOR_ATUAL_HASH || '') === corePerfilHash_(current.value, deps) ||
    String(current.value) === String(requested);
  return {
    ok: true,
    environment: 'PROD',
    dryRun: true,
    solicitacaoEncontrada: true,
    idSolicitacao: record.ID_SOLICITACAO,
    statusAtual: corePerfilNormalizeToken_(record.STATUS),
    idPessoa: idPessoa,
    campo: field,
    valorAtualMascarado: corePerfilMaskValue_(field, current.value),
    valorSolicitadoMascarado: corePerfilMaskValue_(field, requested),
    estadoPessoasBase: {
      emailPrincipalMascarado: corePerfilMaskValue_('EMAIL_PRINCIPAL', corePerfilFindById_(corePerfilSource_(data, 'PESSOAS_BASE'), idPessoa).EMAIL_PRINCIPAL)
    },
    identificadoresEncontrados: identifiers.map(function(item) {
      return {
        idIdentificador: String(item.ID_IDENTIFICADOR || ''),
        valorMascarado: corePerfilMaskValue_(field, item.VALOR_IDENTIFICADOR),
        ativo: corePerfilNormalizeToken_(item.ATIVO),
        principal: corePerfilNormalizeToken_(item.PRINCIPAL)
      };
    }),
    duplicidades: duplicates,
    alteracoesPlanejadas: [
      type ? 'CRIAR_OU_REATIVAR_IDENTIFICADOR_' + type : '',
      'ATUALIZAR_' + current.cfg.source + '_' + current.cfg.header,
      type ? 'DESATIVAR_IDENTIFICADOR_ANTERIOR' : '',
      'RECALCULAR_PESSOAS_RESUMO_OPERACIONAL',
      'INVALIDAR_CACHES_IDENTIDADE',
      'MARCAR_SOLICITACAO_APLICADA'
    ].filter(String),
    cachesAInvalidar: ['EMAIL_ANTIGO', 'EMAIL_NOVO', 'ID_PESSOA', 'SESSAO', 'MINHA_SITUACAO'],
    riscoConflito: duplicates.length ? 'BLOQUEANTE' : (hashCompatible ? 'BAIXO' : 'VALOR_ATUAL_ALTERADO_INCOMPATIVEL'),
    mutacaoDisponivel: false,
    escritaExecutada: false,
    idempotente: true
  };
}

function core_diagnosticarReparacaoSolicitacaoCadastralProd_(options) {
  var opts = options || {};
  var id = String(opts.idSolicitacao || '').trim();
  if (!id) throw new Error('ID_SOLICITACAO_OBRIGATORIO');
  if (opts.dryRun === false) throw new Error('REPARACAO_PROD_MUTACAO_NAO_EXPOSTA');
  var contexto = { ambientePortal: 'PROD' };
  var deps = opts.deps || {};
  deps.environment = 'PROD';
  var data = corePerfilOpenPessoas_(contexto, deps, {
    forWrite: false,
    sources: CORE_PERFIL_APPLY_SOURCES
  });
  var report = corePerfilRepairReport_(id, data, deps);
  return Object.freeze(report);
}
