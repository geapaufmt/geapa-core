/**
 * Cliente Firestore REST compartilhado pelos modulos GEAPA.
 *
 * Usa somente OAuth do Apps Script. Escritas sao dry-run por padrao e nunca
 * registram token, payload ou identificadores de projeto em logs.
 */

var CORE_FIRESTORE_PROJECT_ID_PROPERTY = 'GEAPA_CORE_FIRESTORE_PROJECT_ID';
var CORE_FIRESTORE_DATABASE_ID_PROPERTY = 'GEAPA_CORE_FIRESTORE_DATABASE_ID';
var CORE_FIRESTORE_API_BASE = 'https://firestore.googleapis.com/v1';
var CORE_FIRESTORE_MAX_BATCH_WRITES = 500;

function coreFirestoreGetConfig_(options) {
  options = options || {};
  var props = PropertiesService.getScriptProperties();
  return Object.freeze({
    projectId: String(options.projectId || props.getProperty(CORE_FIRESTORE_PROJECT_ID_PROPERTY) || '').trim(),
    databaseId: String(options.databaseId || props.getProperty(CORE_FIRESTORE_DATABASE_ID_PROPERTY) || '(default)').trim() || '(default)'
  });
}

function coreFirestoreNormalizeDocumentPath_(path) {
  var text = String(path || '').trim().replace(/^\/+|\/+$/g, '');
  var segments = text ? text.split('/') : [];
  if (!segments.length || segments.length % 2 !== 0) {
    throw new Error('O caminho Firestore deve apontar para um documento.');
  }
  segments.forEach(function(segment) {
    if (!String(segment || '').trim()) throw new Error('Caminho Firestore invalido.');
  });
  return segments;
}

function coreFirestoreNormalizeCollectionPath_(path) {
  var text = String(path || '').trim().replace(/^\/+|\/+$/g, '');
  var segments = text ? text.split('/') : [];
  if (!segments.length || segments.length % 2 === 0) {
    throw new Error('O caminho Firestore deve apontar para uma colecao.');
  }
  segments.forEach(function(segment) {
    if (!String(segment || '').trim()) throw new Error('Caminho Firestore invalido.');
  });
  return segments;
}

function coreFirestoreDatabasePath_(config) {
  var projectId = encodeURIComponent(String(config.projectId || '').trim());
  var databaseId = String(config.databaseId || '(default)').trim() || '(default)';
  var encodedDatabase = databaseId === '(default)' ? '(default)' : encodeURIComponent(databaseId);
  return 'projects/' + projectId + '/databases/' + encodedDatabase;
}

function coreFirestoreBuildDocumentUrl_(path, options) {
  var config = coreFirestoreGetConfig_(options || {});
  var segments = coreFirestoreNormalizeDocumentPath_(path).map(function(segment) {
    return encodeURIComponent(String(segment).trim());
  });
  if (!config.projectId) throw new Error('GEAPA_CORE_FIRESTORE_PROJECT_ID nao configurado.');
  return CORE_FIRESTORE_API_BASE + '/' + coreFirestoreDatabasePath_(config) + '/documents/' + segments.join('/');
}

function coreFirestoreBuildDocumentName_(path, config) {
  var segments = coreFirestoreNormalizeDocumentPath_(path);
  if (!config.projectId) throw new Error('GEAPA_CORE_FIRESTORE_PROJECT_ID nao configurado.');
  return coreFirestoreDatabasePath_(config) + '/documents/' + segments.join('/');
}

function coreFirestoreEncodeValue_(value) {
  if (value === null) return { nullValue: 'NULL_VALUE' };
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return { nullValue: 'NULL_VALUE' };
    return { timestampValue: value.toISOString() };
  }
  if (typeof value === 'boolean') return { booleanValue: value };
  if (typeof value === 'number') {
    if (!isFinite(value)) return { nullValue: 'NULL_VALUE' };
    return Number.isInteger(value)
      ? { integerValue: String(value) }
      : { doubleValue: value };
  }
  if (Array.isArray(value)) {
    return {
      arrayValue: {
        values: value.filter(function(item) {
          return item !== undefined;
        }).map(coreFirestoreEncodeValue_)
      }
    };
  }
  if (value && typeof value === 'object') {
    return {
      mapValue: {
        fields: coreFirestoreEncodeDocument_(value).fields
      }
    };
  }
  return { stringValue: String(value == null ? '' : value) };
}

function coreFirestoreEncodeDocument_(object) {
  var fields = {};
  Object.keys(object || {}).forEach(function(key) {
    if (object[key] === undefined) return;
    fields[key] = coreFirestoreEncodeValue_(object[key]);
  });
  return Object.freeze({ fields: fields });
}

function coreFirestoreDecodeValue_(value) {
  value = value || {};
  if (Object.prototype.hasOwnProperty.call(value, 'nullValue')) return null;
  if (Object.prototype.hasOwnProperty.call(value, 'stringValue')) return String(value.stringValue || '');
  if (Object.prototype.hasOwnProperty.call(value, 'booleanValue')) return value.booleanValue === true;
  if (Object.prototype.hasOwnProperty.call(value, 'integerValue')) return Number(value.integerValue || 0);
  if (Object.prototype.hasOwnProperty.call(value, 'doubleValue')) return Number(value.doubleValue || 0);
  if (Object.prototype.hasOwnProperty.call(value, 'timestampValue')) return String(value.timestampValue || '');
  if (value.arrayValue) return (value.arrayValue.values || []).map(coreFirestoreDecodeValue_);
  if (value.mapValue) {
    var out = {};
    Object.keys(value.mapValue.fields || {}).forEach(function(key) {
      out[key] = coreFirestoreDecodeValue_(value.mapValue.fields[key]);
    });
    return out;
  }
  return null;
}

function coreFirestoreDecodeDocument_(document) {
  var out = {};
  Object.keys(document && document.fields || {}).forEach(function(key) {
    out[key] = coreFirestoreDecodeValue_(document.fields[key]);
  });
  return Object.freeze(out);
}

function coreFirestoreSafeError_(responseText) {
  try {
    var parsed = JSON.parse(String(responseText || '{}'));
    var error = parsed.error || {};
    return Object.freeze({
      status: String(error.status || '').trim(),
      message: String(error.message || '').trim().slice(0, 500)
    });
  } catch (err) {
    return Object.freeze({ status: '', message: 'Resposta Firestore invalida.' });
  }
}

function coreFirestoreRequest_(url, requestOptions) {
  try {
    var fetchOptions = Object.assign({ muteHttpExceptions: true }, requestOptions || {});
    fetchOptions.headers = Object.assign({
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }, requestOptions && requestOptions.headers || {});
    var response = UrlFetchApp.fetch(url, fetchOptions);
    var httpStatus = response.getResponseCode();
    var ok = httpStatus >= 200 && httpStatus < 300;
    return {
      ok: ok,
      httpStatus: httpStatus,
      body: ok ? String(response.getContentText() || '') : '',
      firestoreError: ok ? null : coreFirestoreSafeError_(response.getContentText())
    };
  } catch (err) {
    return {
      ok: false,
      httpStatus: 0,
      body: '',
      firestoreError: Object.freeze({ status: 'REQUEST_EXCEPTION', message: String(err && err.message || err || '').slice(0, 500) })
    };
  }
}

function coreFirestoreUpdateMaskQuery_(data, merge) {
  if (merge === false) return '';
  return Object.keys(data || {}).map(function(field) {
    return 'updateMask.fieldPaths=' + encodeURIComponent(field);
  }).join('&');
}

function coreFirestoreSetDocument_(path, data, options) {
  options = options || {};
  var dryRun = options.dryRun !== false;
  try {
    coreFirestoreNormalizeDocumentPath_(path);
    coreFirestoreEncodeDocument_(data || {});
    if (dryRun) {
      return Object.freeze({ ok: true, written: false, dryRun: true, code: 'DRY_RUN', path: String(path || '') });
    }
    var config = coreFirestoreGetConfig_(options);
    if (!config.projectId) {
      return Object.freeze({ ok: false, written: false, dryRun: dryRun, code: 'FIRESTORE_PROJECT_ID_NAO_CONFIGURADO' });
    }
    var url = coreFirestoreBuildDocumentUrl_(path, config);
    var mask = coreFirestoreUpdateMaskQuery_(data, options.merge);
    if (mask) url += '?' + mask;
    var response = coreFirestoreRequest_(url, {
      method: 'patch',
      contentType: 'application/json; charset=utf-8',
      payload: JSON.stringify(coreFirestoreEncodeDocument_(data || {}))
    });
    return Object.freeze({
      ok: response.ok,
      written: response.ok,
      dryRun: false,
      code: response.ok ? 'FIRESTORE_SET_OK' : 'FIRESTORE_SET_FALHOU',
      path: String(path || ''),
      httpStatus: response.httpStatus,
      firestoreError: response.firestoreError || undefined
    });
  } catch (err) {
    return Object.freeze({ ok: false, written: false, dryRun: dryRun, code: 'FIRESTORE_SET_INVALIDO', message: String(err && err.message || err || '').slice(0, 500) });
  }
}

function coreFirestoreGetDocument_(path, options) {
  options = options || {};
  try {
    var config = coreFirestoreGetConfig_(options);
    if (!config.projectId) return Object.freeze({ ok: true, found: false, code: 'FIRESTORE_PROJECT_ID_NAO_CONFIGURADO' });
    var response = coreFirestoreRequest_(coreFirestoreBuildDocumentUrl_(path, config), { method: 'get' });
    if (response.httpStatus === 404) return Object.freeze({ ok: true, found: false, code: 'FIRESTORE_DOCUMENTO_NAO_ENCONTRADO' });
    var result = {
      ok: response.ok,
      found: response.ok,
      code: response.ok ? 'FIRESTORE_GET_OK' : 'FIRESTORE_GET_FALHOU',
      httpStatus: response.httpStatus,
      firestoreError: response.firestoreError || undefined
    };
    if (response.ok) result.data = coreFirestoreDecodeDocument_(JSON.parse(response.body || '{}'));
    return Object.freeze(result);
  } catch (err) {
    return Object.freeze({ ok: false, found: false, code: 'FIRESTORE_GET_INVALIDO', message: String(err && err.message || err || '').slice(0, 500) });
  }
}

/** Lista uma pagina de documentos sem expor token ou identificadores do projeto. */
function coreFirestoreListDocuments_(collectionPath, options) {
  options = options || {};
  try {
    var config = coreFirestoreGetConfig_(options);
    if (!config.projectId) {
      return Object.freeze({ ok: false, documents: Object.freeze([]), code: 'FIRESTORE_PROJECT_ID_NAO_CONFIGURADO' });
    }
    var segments = coreFirestoreNormalizeCollectionPath_(collectionPath).map(function(segment) {
      return encodeURIComponent(String(segment).trim());
    });
    var pageSize = Math.floor(Number(options.pageSize || 500));
    if (!isFinite(pageSize) || pageSize < 1) pageSize = 500;
    pageSize = Math.min(pageSize, 500);
    var query = ['pageSize=' + pageSize];
    if (String(options.pageToken || '').trim()) query.push('pageToken=' + encodeURIComponent(String(options.pageToken).trim()));
    var url = CORE_FIRESTORE_API_BASE + '/' + coreFirestoreDatabasePath_(config) + '/documents/' + segments.join('/') + '?' + query.join('&');
    var response = coreFirestoreRequest_(url, { method: 'get' });
    if (!response.ok) {
      return Object.freeze({
        ok: false,
        documents: Object.freeze([]),
        code: 'FIRESTORE_LIST_FALHOU',
        httpStatus: response.httpStatus,
        firestoreError: response.firestoreError || undefined
      });
    }
    var payload = JSON.parse(response.body || '{}');
    var documents = (payload.documents || []).map(function(document) {
      var fullName = String(document && document.name || '');
      var marker = '/documents/';
      var markerIndex = fullName.indexOf(marker);
      var relativePath = markerIndex >= 0 ? fullName.slice(markerIndex + marker.length) : '';
      return Object.freeze({
        path: relativePath,
        id: relativePath ? relativePath.split('/').pop() : '',
        data: coreFirestoreDecodeDocument_(document || {}),
        createTime: String(document && document.createTime || ''),
        updateTime: String(document && document.updateTime || '')
      });
    });
    return Object.freeze({
      ok: true,
      documents: Object.freeze(documents),
      nextPageToken: String(payload.nextPageToken || ''),
      code: 'FIRESTORE_LIST_OK',
      httpStatus: response.httpStatus
    });
  } catch (err) {
    return Object.freeze({
      ok: false,
      documents: Object.freeze([]),
      code: 'FIRESTORE_LIST_INVALIDO',
      message: String(err && err.message || err || '').slice(0, 500)
    });
  }
}

function coreFirestoreDeleteDocument_(path, options) {
  options = options || {};
  var dryRun = options.dryRun !== false;
  try {
    coreFirestoreNormalizeDocumentPath_(path);
    if (dryRun) return Object.freeze({ ok: true, deleted: false, dryRun: true, code: 'DRY_RUN', path: String(path || '') });
    var config = coreFirestoreGetConfig_(options);
    if (!config.projectId) return Object.freeze({ ok: false, deleted: false, dryRun: dryRun, code: 'FIRESTORE_PROJECT_ID_NAO_CONFIGURADO' });
    var url = coreFirestoreBuildDocumentUrl_(path, config);
    var response = coreFirestoreRequest_(url, { method: 'delete' });
    return Object.freeze({
      ok: response.ok,
      deleted: response.ok,
      dryRun: false,
      code: response.ok ? 'FIRESTORE_DELETE_OK' : 'FIRESTORE_DELETE_FALHOU',
      path: String(path || ''),
      httpStatus: response.httpStatus,
      firestoreError: response.firestoreError || undefined
    });
  } catch (err) {
    return Object.freeze({ ok: false, deleted: false, dryRun: dryRun, code: 'FIRESTORE_DELETE_INVALIDO', message: String(err && err.message || err || '').slice(0, 500) });
  }
}

function coreFirestoreBatchSetDocuments_(items, options) {
  options = options || {};
  var list = Array.isArray(items) ? items : [];
  var dryRun = options.dryRun !== false;
  if (!list.length) return Object.freeze({ ok: true, written: 0, requested: 0, dryRun: dryRun, code: 'SEM_DOCUMENTOS' });
  if (list.length > CORE_FIRESTORE_MAX_BATCH_WRITES) {
    return Object.freeze({ ok: false, written: 0, requested: list.length, dryRun: dryRun, code: 'LIMITE_BATCH_EXCEDIDO' });
  }
  try {
    list.forEach(function(item) {
      coreFirestoreNormalizeDocumentPath_(item && item.path);
      coreFirestoreEncodeDocument_(item && item.data || {});
    });
    if (dryRun) {
      return Object.freeze({ ok: true, written: 0, requested: list.length, dryRun: true, code: 'DRY_RUN', paths: Object.freeze(list.map(function(item) { return String(item && item.path || ''); })) });
    }
    var config = coreFirestoreGetConfig_(options);
    if (!config.projectId) return Object.freeze({ ok: false, written: 0, requested: list.length, dryRun: dryRun, code: 'FIRESTORE_PROJECT_ID_NAO_CONFIGURADO' });
    var writes = list.map(function(item) {
      var data = item && item.data || {};
      var write = {
        update: Object.assign({ name: coreFirestoreBuildDocumentName_(item && item.path, config) }, coreFirestoreEncodeDocument_(data))
      };
      if (options.merge !== false) write.updateMask = { fieldPaths: Object.keys(data) };
      return write;
    });
    var url = CORE_FIRESTORE_API_BASE + '/' + coreFirestoreDatabasePath_(config) + '/documents:commit';
    var response = coreFirestoreRequest_(url, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      payload: JSON.stringify({ writes: writes })
    });
    return Object.freeze({
      ok: response.ok,
      written: response.ok ? list.length : 0,
      requested: list.length,
      dryRun: false,
      code: response.ok ? 'FIRESTORE_BATCH_SET_OK' : 'FIRESTORE_BATCH_SET_FALHOU',
      httpStatus: response.httpStatus,
      firestoreError: response.firestoreError || undefined
    });
  } catch (err) {
    return Object.freeze({ ok: false, written: 0, requested: list.length, dryRun: dryRun, code: 'FIRESTORE_BATCH_INVALIDO', message: String(err && err.message || err || '').slice(0, 500) });
  }
}

function coreFirestoreDiagnosticar_(options) {
  options = options || {};
  var config = coreFirestoreGetConfig_(options);
  var tokenAvailable = false;
  try {
    tokenAvailable = !!String(ScriptApp.getOAuthToken() || '').trim();
  } catch (err) {}
  return Object.freeze({
    ok: !!config.projectId && tokenAvailable,
    readOnly: true,
    projectIdConfigured: !!config.projectId,
    databaseId: config.databaseId,
    oauthTokenAvailable: tokenAvailable,
    datastoreScopeDeclared: true,
    writer: 'APPS_SCRIPT_FIRESTORE_REST',
    defaultDryRun: true
  });
}
