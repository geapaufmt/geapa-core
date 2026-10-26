/**
 * Delete atomico com preconditions para artefatos tecnicos Firestore DEV.
 *
 * Reutiliza o batchGet e o documents:commit do cliente REST existente. Nao ha
 * fallback para delete generico, batch delete, update vazio ou overwrite.
 */

var CORE_FIRESTORE_CONDITIONAL_DELETE_ALLOWED_COLLECTIONS = Object.freeze({
  presentations: true,
  presentationPrivate: true,
  presentationDecisionRequests: true
});

function coreFirestoreConditionalDeleteResult_(code, details) {
  var info = details || {};
  return Object.freeze({
    ok: info.ok === true,
    deleted: Number(info.deleted || 0),
    requested: Number(info.requested || 0),
    dryRun: info.dryRun === true,
    atomic: info.atomic === true,
    preconditionsApplied: info.preconditionsApplied === true,
    code: String(code || 'FIRESTORE_CONDITIONAL_DELETE_INVALIDO'),
    concurrent: info.concurrent === true,
    httpStatus: Number(info.httpStatus || 0),
    paths: Object.freeze((info.paths || []).slice()),
    commitTime: String(info.commitTime || ''),
    firestoreError: info.firestoreError || undefined
  });
}

function coreFirestoreConditionalDeleteUpdateTime_(value) {
  var updateTime = String(value || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,9})?Z$/.test(
        updateTime
      ) || isNaN(new Date(updateTime).getTime())) {
    throw new Error('FIRESTORE_CONDITIONAL_DELETE_UPDATE_TIME_INVALIDO');
  }
  return updateTime;
}

function coreFirestoreConditionalDeletePath_(value) {
  var path = String(value || '').trim();
  if (!path || path !== path.replace(/^\/+|\/+$/g, '') ||
      /^projects\//i.test(path) || /\/documents\//i.test(path)) {
    throw new Error('FIRESTORE_CONDITIONAL_DELETE_PATH_INVALIDO');
  }
  var segments = coreFirestoreNormalizeDocumentPath_(path);
  if (segments.length !== 2 ||
      !CORE_FIRESTORE_CONDITIONAL_DELETE_ALLOWED_COLLECTIONS[segments[0]] ||
      !/^[A-Za-z0-9_-]{1,1500}$/.test(segments[1]) ||
      segments[1] === '.' || segments[1] === '..' ||
      /^__.*__$/.test(segments[1])) {
    throw new Error('FIRESTORE_CONDITIONAL_DELETE_PATH_NAO_AUTORIZADO');
  }
  return segments.join('/');
}

function coreFirestoreConditionalDeleteNormalize_(documents) {
  if (!Array.isArray(documents) || !documents.length ||
      documents.length > CORE_FIRESTORE_MAX_BATCH_WRITES) {
    throw new Error('FIRESTORE_CONDITIONAL_DELETE_DOCUMENTOS_INVALIDOS');
  }
  var seen = {};
  return Object.freeze(documents.map(function(item) {
    if (!item || Object.prototype.toString.call(item) !== '[object Object]') {
      throw new Error('FIRESTORE_CONDITIONAL_DELETE_DOCUMENTO_INVALIDO');
    }
    var keys = Object.keys(item).sort();
    if (keys.length !== 2 || keys[0] !== 'expectedUpdateTime' ||
        keys[1] !== 'path') {
      throw new Error('FIRESTORE_CONDITIONAL_DELETE_CONTRATO_INVALIDO');
    }
    var path = coreFirestoreConditionalDeletePath_(item.path);
    if (seen[path]) {
      throw new Error('FIRESTORE_CONDITIONAL_DELETE_PATH_DUPLICADO');
    }
    seen[path] = true;
    return Object.freeze({
      path: path,
      expectedUpdateTime: coreFirestoreConditionalDeleteUpdateTime_(
        item.expectedUpdateTime
      )
    });
  }));
}

function coreFirestoreDeleteDocumentsWithPreconditions_(documents, options) {
  var opts = options || {};
  var dryRun = opts.dryRun !== false;
  var normalized;
  try {
    normalized = coreFirestoreConditionalDeleteNormalize_(documents);
  } catch (err) {
    return coreFirestoreConditionalDeleteResult_(
      String(err && err.message || 'FIRESTORE_CONDITIONAL_DELETE_INVALIDO'),
      {
        requested: Array.isArray(documents) ? documents.length : 0,
        dryRun: dryRun
      }
    );
  }
  var paths = normalized.map(function(item) { return item.path; });
  if (dryRun) {
    return coreFirestoreConditionalDeleteResult_('DRY_RUN', {
      ok: true,
      requested: normalized.length,
      dryRun: true,
      paths: paths
    });
  }

  var config = coreFirestoreGetConfig_(opts);
  if (!config.projectId) {
    return coreFirestoreConditionalDeleteResult_(
      'FIRESTORE_PROJECT_ID_NAO_CONFIGURADO',
      { requested: normalized.length, paths: paths }
    );
  }
  var versions = coreFirestoreBatchGetVersions_(paths, config);
  if (!versions.ok) {
    return coreFirestoreConditionalDeleteResult_(versions.code, {
      requested: normalized.length,
      paths: paths,
      httpStatus: versions.httpStatus,
      firestoreError: versions.firestoreError
    });
  }

  var failedCode = '';
  normalized.some(function(item) {
    var current = versions.documents[item.path];
    if (!current || current.found !== true) {
      failedCode = 'FIRESTORE_CONDITIONAL_DELETE_DOCUMENTO_AUSENTE';
      return true;
    }
    if (!current.updateTime ||
        current.updateTime !== item.expectedUpdateTime) {
      failedCode = 'FIRESTORE_CONDITIONAL_DELETE_UPDATE_TIME_DIVERGENTE';
      return true;
    }
    return false;
  });
  if (failedCode) {
    return coreFirestoreConditionalDeleteResult_(failedCode, {
      requested: normalized.length,
      paths: paths,
      concurrent: failedCode.indexOf('DIVERGENTE') >= 0
    });
  }

  var writes = normalized.map(function(item) {
    return Object.freeze({
      delete: coreFirestoreBuildDocumentName_(item.path, config),
      currentDocument: Object.freeze({
        updateTime: item.expectedUpdateTime
      })
    });
  });
  var url = CORE_FIRESTORE_API_BASE + '/' + coreFirestoreDatabasePath_(config) +
    '/documents:commit';
  var response = coreFirestoreRequest_(url, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify({ writes: writes })
  });
  if (!response.ok) {
    var status = String(response.firestoreError &&
      response.firestoreError.status || '').trim().toUpperCase();
    var concurrent = status === 'FAILED_PRECONDITION' || status === 'ABORTED';
    return coreFirestoreConditionalDeleteResult_(
      concurrent
        ? 'FIRESTORE_CONDITIONAL_DELETE_CONCORRENCIA_DETECTADA'
        : 'FIRESTORE_CONDITIONAL_DELETE_COMMIT_FALHOU',
      {
        requested: normalized.length,
        paths: paths,
        concurrent: concurrent,
        httpStatus: response.httpStatus,
        firestoreError: response.firestoreError
      }
    );
  }

  var commitPayload;
  try {
    commitPayload = JSON.parse(response.body || '{}');
  } catch (err) {
    commitPayload = {};
  }
  var writeResults = Array.isArray(commitPayload.writeResults)
    ? commitPayload.writeResults : [];
  if (writeResults.length !== normalized.length) {
    return coreFirestoreConditionalDeleteResult_(
      'FIRESTORE_CONDITIONAL_DELETE_COMMIT_RESPOSTA_INCOMPLETA',
      { requested: normalized.length, paths: paths }
    );
  }
  return coreFirestoreConditionalDeleteResult_(
    'FIRESTORE_CONDITIONAL_DELETE_OK',
    {
      ok: true,
      deleted: normalized.length,
      requested: normalized.length,
      atomic: true,
      preconditionsApplied: true,
      paths: paths,
      commitTime: String(commitPayload.commitTime || '')
    }
  );
}

function coreFirestoreEnvironmentDeleteDocumentsWithPreconditions_(
  documents,
  options
) {
  var requested = Array.isArray(documents) ? documents.length : 0;
  var dryRun = !options || options.dryRun !== false;
  try {
    coreFirestoreAssertEnvironmentWriteAllowed_(options || {});
    return coreFirestoreDeleteDocumentsWithPreconditions_(
      documents,
      coreFirestoreEnvironmentOptions_(options || {})
    );
  } catch (err) {
    return coreFirestoreConditionalDeleteResult_(
      'FIRESTORE_CONDITIONAL_DELETE_ENVIRONMENT_INVALIDO',
      { requested: requested, dryRun: dryRun }
    );
  }
}
