/**
 * Extensao do cliente Firestore REST para commits atomicos condicionais.
 *
 * Reutiliza configuracao, autenticacao, codec e request do cliente 24. Nao faz
 * fallback para batch simples e toda escrita permanece dry-run por padrao.
 */

function coreFirestorePreconditionResult_(code, details) {
  var info = details || {};
  return Object.freeze({
    ok: false,
    written: 0,
    requested: Number(info.requested || 0),
    dryRun: info.dryRun === true,
    atomic: false,
    preconditionsApplied: false,
    code: String(code || 'FIRESTORE_PRECONDITION_INVALIDA'),
    concurrent: info.concurrent === true,
    httpStatus: Number(info.httpStatus || 0),
    firestoreError: info.firestoreError || undefined
  });
}

function coreFirestorePreconditionAttachLatency_(result, trace) {
  try {
    var snapshot = coreFirestoreLatencyMaybeSnapshotDev_(trace);
    if (!snapshot) return result;
    return Object.freeze(Object.assign({}, result || {}, {
      latencyTelemetry: snapshot
    }));
  } catch (ignored) {
    return result;
  }
}

function coreFirestoreGetDocumentWithVersion_(path, options) {
  options = options || {};
  try {
    var config = coreFirestoreGetConfig_(options);
    if (!config.projectId) {
      return Object.freeze({
        ok: false,
        found: false,
        code: 'FIRESTORE_PROJECT_ID_NAO_CONFIGURADO'
      });
    }
    var response = coreFirestoreRequest_(
      coreFirestoreBuildDocumentUrl_(path, config),
      {
        method: 'get',
        __latencyTrace: options.__latencyTrace,
        __latencyOperation: options.__latencyOperation || 'GET_DOCUMENT_VERSION'
      }
    );
    if (response.httpStatus === 404) {
      return Object.freeze({
        ok: true,
        found: false,
        code: 'FIRESTORE_DOCUMENTO_NAO_ENCONTRADO',
        createTime: '',
        updateTime: ''
      });
    }
    if (!response.ok) {
      return Object.freeze({
        ok: false,
        found: false,
        code: 'FIRESTORE_GET_VERSIONADO_FALHOU',
        httpStatus: response.httpStatus,
        firestoreError: response.firestoreError || undefined
      });
    }
    var document = JSON.parse(response.body || '{}');
    var updateTime = String(document.updateTime || '').trim();
    if (!updateTime) {
      return Object.freeze({
        ok: false,
        found: false,
        code: 'FIRESTORE_GET_VERSIONADO_SEM_UPDATE_TIME'
      });
    }
    return Object.freeze({
      ok: true,
      found: true,
      code: 'FIRESTORE_GET_VERSIONADO_OK',
      data: coreFirestoreDecodeDocument_(document),
      createTime: String(document.createTime || ''),
      updateTime: updateTime
    });
  } catch (err) {
    return Object.freeze({
      ok: false,
      found: false,
      code: 'FIRESTORE_GET_VERSIONADO_INVALIDO',
      message: String(err && err.message || err || '').slice(0, 500)
    });
  }
}

function coreFirestorePreconditionNormalizePayload_(payload) {
  var input = payload || {};
  var writes = Array.isArray(input.writes) ? input.writes : [];
  var preconditions = Array.isArray(input.preconditions) ? input.preconditions : [];
  if (!writes.length || writes.length > CORE_FIRESTORE_MAX_BATCH_WRITES) {
    throw new Error('FIRESTORE_PRECONDITION_WRITES_INVALIDOS');
  }
  if (preconditions.length !== writes.length) {
    throw new Error('FIRESTORE_PRECONDITION_OBRIGATORIA');
  }
  var writesByPath = {};
  var conditionsByPath = {};
  var normalizedWrites = writes.map(function(item) {
    var path = String(item && item.path || '').trim().replace(/^\/+|\/+$/g, '');
    coreFirestoreNormalizeDocumentPath_(path);
    coreFirestoreEncodeDocument_(item && item.data || {});
    if (writesByPath[path]) throw new Error('FIRESTORE_PRECONDITION_PATH_DUPLICADO');
    writesByPath[path] = true;
    return Object.freeze({
      path: path,
      data: item && item.data || {},
      merge: item && item.merge !== false
    });
  });
  var normalizedConditions = preconditions.map(function(condition) {
    var path = String(condition && condition.path || '').trim()
      .replace(/^\/+|\/+$/g, '');
    coreFirestoreNormalizeDocumentPath_(path);
    if (!writesByPath[path] || conditionsByPath[path]) {
      throw new Error('FIRESTORE_PRECONDITION_ESCOPO_INVALIDO');
    }
    var hasSourceHash = !!String(condition && condition.sourceHash || '').trim();
    var hasUpdateTime = !!String(condition && condition.updateTime || '').trim();
    var hasExists = condition && typeof condition.exists === 'boolean';
    if (Number(hasSourceHash) + Number(hasUpdateTime) + Number(hasExists) !== 1) {
      throw new Error('FIRESTORE_PRECONDITION_TIPO_INVALIDO');
    }
    if (hasExists && condition.exists !== false) {
      throw new Error('FIRESTORE_PRECONDITION_EXISTS_TRUE_INSUFICIENTE');
    }
    if (hasSourceHash &&
        !/^sha256:[a-f0-9]{64}$/.test(String(condition.sourceHash).trim())) {
      throw new Error('FIRESTORE_PRECONDITION_SOURCE_HASH_INVALIDO');
    }
    conditionsByPath[path] = true;
    return Object.freeze({
      path: path,
      sourceHash: hasSourceHash ? String(condition.sourceHash).trim() : '',
      updateTime: hasUpdateTime ? String(condition.updateTime).trim() : '',
      exists: hasExists ? condition.exists : undefined
    });
  });
  normalizedWrites.forEach(function(item) {
    if (!conditionsByPath[item.path]) {
      throw new Error('FIRESTORE_PRECONDITION_OBRIGATORIA');
    }
  });
  return Object.freeze({
    writes: Object.freeze(normalizedWrites),
    preconditions: Object.freeze(normalizedConditions)
  });
}

function coreFirestorePreconditionRelativePath_(documentName, config) {
  var prefix = coreFirestoreDatabasePath_(config) + '/documents/';
  var name = String(documentName || '');
  return name.indexOf(prefix) === 0 ? name.slice(prefix.length) : '';
}

function coreFirestoreBatchGetVersions_(paths, config, latencyTrace) {
  var names = paths.map(function(path) {
    return coreFirestoreBuildDocumentName_(path, config);
  });
  var url = CORE_FIRESTORE_API_BASE + '/' + coreFirestoreDatabasePath_(config) +
    '/documents:batchGet';
  coreFirestoreLatencyMaybeMarkDev_(
    latencyTrace,
    'CORE_BATCH_GET_START',
    'BATCH_GET_INICIADO',
    { requested: paths.length }
  );
  var response = coreFirestoreRequest_(url, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify({ documents: names }),
    __latencyTrace: latencyTrace,
    __latencyOperation: 'BATCH_GET_PRECONDITIONS'
  });
  coreFirestoreLatencyMaybeMarkDev_(
    latencyTrace,
    'CORE_BATCH_GET_END',
    'BATCH_GET_CONCLUIDO',
    { requested: paths.length, httpStatus: response.httpStatus }
  );
  if (!response.ok) {
    return Object.freeze({
      ok: false,
      code: 'FIRESTORE_PRECONDITION_BATCH_GET_FALHOU',
      httpStatus: response.httpStatus,
      firestoreError: response.firestoreError || undefined,
      documents: Object.freeze({})
    });
  }
  var entries;
  try {
    entries = JSON.parse(response.body || '[]');
  } catch (err) {
    entries = null;
  }
  if (!Array.isArray(entries)) {
    return Object.freeze({
      ok: false,
      code: 'FIRESTORE_PRECONDITION_BATCH_GET_RESPOSTA_INVALIDA',
      documents: Object.freeze({})
    });
  }
  var resultByPath = {};
  entries.forEach(function(entry) {
    if (entry && entry.found) {
      var foundPath = coreFirestorePreconditionRelativePath_(
        entry.found.name,
        config
      );
      if (foundPath) {
        resultByPath[foundPath] = Object.freeze({
          found: true,
          data: coreFirestoreDecodeDocument_(entry.found),
          createTime: String(entry.found.createTime || ''),
          updateTime: String(entry.found.updateTime || '')
        });
      }
    } else if (entry && entry.missing) {
      var missingPath = coreFirestorePreconditionRelativePath_(
        entry.missing,
        config
      );
      if (missingPath) {
        resultByPath[missingPath] = Object.freeze({
          found: false,
          data: null,
          createTime: '',
          updateTime: ''
        });
      }
    }
  });
  var complete = paths.every(function(path) {
    return Object.prototype.hasOwnProperty.call(resultByPath, path);
  });
  return Object.freeze({
    ok: complete,
    code: complete
      ? 'FIRESTORE_PRECONDITION_BATCH_GET_OK'
      : 'FIRESTORE_PRECONDITION_BATCH_GET_INCOMPLETO',
    documents: Object.freeze(resultByPath)
  });
}

function coreFirestoreResolveNativePreconditions_(normalized, versions) {
  var nativeByPath = {};
  var failure = '';
  normalized.preconditions.some(function(condition) {
    var current = versions.documents[condition.path];
    if (!current) {
      failure = 'FIRESTORE_PRECONDITION_DOCUMENTO_NAO_RESOLVIDO';
      return true;
    }
    if (condition.sourceHash) {
      if (!current.found) {
        failure = 'FIRESTORE_PRECONDITION_DOCUMENTO_AUSENTE';
        return true;
      }
      if (!current.updateTime) {
        failure = 'FIRESTORE_PRECONDITION_UPDATE_TIME_AUSENTE';
        return true;
      }
      if (String(current.data && current.data.sourceHash || '').trim() !==
          condition.sourceHash) {
        failure = 'FIRESTORE_PRECONDITION_SOURCE_HASH_DIVERGENTE';
        return true;
      }
      nativeByPath[condition.path] = Object.freeze({
        updateTime: current.updateTime
      });
      return false;
    }
    if (condition.updateTime) {
      if (!current.found) {
        failure = 'FIRESTORE_PRECONDITION_DOCUMENTO_AUSENTE';
        return true;
      }
      if (!current.updateTime || current.updateTime !== condition.updateTime) {
        failure = 'FIRESTORE_PRECONDITION_UPDATE_TIME_DIVERGENTE';
        return true;
      }
      nativeByPath[condition.path] = Object.freeze({
        updateTime: current.updateTime
      });
      return false;
    }
    if (current.found !== condition.exists) {
      failure = 'FIRESTORE_PRECONDITION_EXISTS_DIVERGENTE';
      return true;
    }
    nativeByPath[condition.path] = Object.freeze({ exists: condition.exists });
    return false;
  });
  return Object.freeze({
    ok: !failure,
    code: failure || 'FIRESTORE_PRECONDITIONS_RESOLVIDAS',
    nativeByPath: Object.freeze(nativeByPath)
  });
}

function coreFirestoreCommitDocumentsWithPreconditions_(payload, options) {
  options = options || {};
  var latencyTrace = coreFirestoreLatencyMaybeStartDev_(options.__latencyTrace);
  coreFirestoreLatencyMaybeMarkDev_(
    latencyTrace,
    'P16',
    'CORE_INICIOU',
    { operation: 'COMMIT_WITH_PRECONDITIONS' }
  );
  var dryRun = options.dryRun !== false;
  var normalized;
  try {
    normalized = coreFirestorePreconditionNormalizePayload_(payload);
  } catch (err) {
    return coreFirestorePreconditionAttachLatency_(coreFirestorePreconditionResult_(
      String(err && err.message || 'FIRESTORE_PRECONDITION_PAYLOAD_INVALIDO'),
      { dryRun: dryRun }
    ), latencyTrace);
  }
  if (dryRun) {
    return coreFirestorePreconditionAttachLatency_(Object.freeze({
      ok: true,
      written: 0,
      requested: normalized.writes.length,
      dryRun: true,
      atomic: false,
      preconditionsApplied: false,
      code: 'DRY_RUN',
      paths: Object.freeze(normalized.writes.map(function(item) {
        return item.path;
      }))
    }), latencyTrace);
  }
  var config = coreFirestoreGetConfig_(options);
  if (!config.projectId) {
    return coreFirestorePreconditionAttachLatency_(coreFirestorePreconditionResult_(
      'FIRESTORE_PROJECT_ID_NAO_CONFIGURADO',
      { requested: normalized.writes.length }
    ), latencyTrace);
  }
  var paths = normalized.writes.map(function(item) { return item.path; });
  var versions = coreFirestoreBatchGetVersions_(
    paths,
    config,
    latencyTrace
  );
  if (!versions.ok) {
    return coreFirestorePreconditionAttachLatency_(coreFirestorePreconditionResult_(versions.code, {
      requested: normalized.writes.length,
      httpStatus: versions.httpStatus,
      firestoreError: versions.firestoreError
    }), latencyTrace);
  }
  var resolved = coreFirestoreResolveNativePreconditions_(normalized, versions);
  if (!resolved.ok) {
    return coreFirestorePreconditionAttachLatency_(coreFirestorePreconditionResult_(resolved.code, {
      requested: normalized.writes.length,
      concurrent: /DIVERGENTE/.test(resolved.code)
    }), latencyTrace);
  }
  coreFirestoreLatencyMaybeMarkDev_(
    latencyTrace,
    'P17',
    'CORE_VERSIONAMENTO_CONCLUIDO',
    {
      operation: 'BATCH_GET_PRECONDITIONS',
      requested: normalized.writes.length
    }
  );
  coreFirestoreLatencyMaybeMarkDev_(
    latencyTrace,
    'CORE_COMMIT_BUILD_START',
    'DOCUMENTS_COMMIT_MONTAGEM_INICIADA',
    { requested: normalized.writes.length }
  );
  var firestoreWrites = normalized.writes.map(function(item) {
    var write = {
      update: Object.assign(
        { name: coreFirestoreBuildDocumentName_(item.path, config) },
        coreFirestoreEncodeDocument_(item.data)
      ),
      currentDocument: resolved.nativeByPath[item.path]
    };
    if (item.merge !== false) {
      write.updateMask = { fieldPaths: Object.keys(item.data) };
    }
    return write;
  });
  coreFirestoreLatencyMaybeMarkDev_(
    latencyTrace,
    'CORE_COMMIT_BUILD_END',
    'DOCUMENTS_COMMIT_MONTAGEM_CONCLUIDA',
    { requested: normalized.writes.length }
  );
  var url = CORE_FIRESTORE_API_BASE + '/' + coreFirestoreDatabasePath_(config) +
    '/documents:commit';
  coreFirestoreLatencyMaybeMarkDev_(
    latencyTrace,
    'P18',
    'DOCUMENTS_COMMIT_INICIADO',
    { requested: normalized.writes.length }
  );
  var response = coreFirestoreRequest_(url, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify({ writes: firestoreWrites }),
    __latencyTrace: latencyTrace,
    __latencyOperation: 'DOCUMENTS_COMMIT'
  });
  coreFirestoreLatencyMaybeMarkDev_(
    latencyTrace,
    'P19',
    'DOCUMENTS_COMMIT_RETORNOU',
    { httpStatus: response.httpStatus }
  );
  if (!response.ok) {
    var status = String(response.firestoreError &&
      response.firestoreError.status || '').trim().toUpperCase();
    var concurrent = status === 'FAILED_PRECONDITION' || status === 'ABORTED';
    return coreFirestorePreconditionAttachLatency_(coreFirestorePreconditionResult_(
      concurrent
        ? 'FIRESTORE_PRECONDITION_CONCORRENCIA_DETECTADA'
        : 'FIRESTORE_PRECONDITION_COMMIT_FALHOU',
      {
        requested: normalized.writes.length,
        concurrent: concurrent,
        httpStatus: response.httpStatus,
        firestoreError: response.firestoreError
      }
    ), latencyTrace);
  }
  coreFirestoreLatencyMaybeMarkDev_(
    latencyTrace,
    'CORE_COMMIT_PARSE_START',
    'DOCUMENTS_COMMIT_PARSE_INICIADO'
  );
  var commitPayload;
  try {
    commitPayload = JSON.parse(response.body || '{}');
  } catch (err) {
    commitPayload = {};
  }
  coreFirestoreLatencyMaybeMarkDev_(
    latencyTrace,
    'CORE_COMMIT_PARSE_END',
    'DOCUMENTS_COMMIT_PARSE_CONCLUIDO'
  );
  var writeResults = Array.isArray(commitPayload.writeResults)
    ? commitPayload.writeResults : [];
  if (writeResults.length !== normalized.writes.length) {
    return coreFirestorePreconditionAttachLatency_(coreFirestorePreconditionResult_(
      'FIRESTORE_PRECONDITION_COMMIT_RESPOSTA_INCOMPLETA',
      { requested: normalized.writes.length }
    ), latencyTrace);
  }
  coreFirestoreLatencyMaybeMarkDev_(
    latencyTrace,
    'P20',
    'CORE_POS_COMMIT_VALIDADO',
    {
      requested: normalized.writes.length,
      result: 'FIRESTORE_PRECONDITION_COMMIT_OK'
    }
  );
  return coreFirestorePreconditionAttachLatency_(Object.freeze({
    ok: true,
    written: normalized.writes.length,
    requested: normalized.writes.length,
    dryRun: false,
    atomic: true,
    preconditionsApplied: true,
    code: 'FIRESTORE_PRECONDITION_COMMIT_OK',
    commitTime: String(commitPayload.commitTime || ''),
    concurrent: false
  }), latencyTrace);
}

function coreFirestoreEnvironmentGetDocumentWithVersion_(path, options) {
  try {
    return coreFirestoreGetDocumentWithVersion_(
      path,
      coreFirestoreEnvironmentOptions_(options || {})
    );
  } catch (err) {
    return Object.freeze({
      ok: false,
      found: false,
      code: 'FIRESTORE_ENVIRONMENT_INVALIDO',
      message: String(err && err.message || err || '').slice(0, 500)
    });
  }
}

function coreFirestoreEnvironmentCommitDocumentsWithPreconditions_(payload, options) {
  try {
    coreFirestoreAssertEnvironmentWriteAllowed_(options || {});
    return coreFirestoreCommitDocumentsWithPreconditions_(
      payload || {},
      coreFirestoreEnvironmentOptions_(options || {})
    );
  } catch (err) {
    return coreFirestorePreconditionResult_('FIRESTORE_ENVIRONMENT_INVALIDO', {
      requested: payload && Array.isArray(payload.writes)
        ? payload.writes.length : 0,
      dryRun: !options || options.dryRun !== false
    });
  }
}
