/**
 * Telemetria tecnica DEV-only do cliente Firestore REST.
 *
 * O contrato aceita somente correlation id e metadados de duracao. URLs,
 * projetos, paths, tokens, payloads e dados de documentos nunca sao logados.
 */

function coreFirestoreLatencyStartDev_(value) {
  var input = value || {};
  if (input.__coreLatencyTrace === true && input.environment === 'DEV') {
    return input;
  }
  if (String(input.environment || '').trim().toUpperCase() !== 'DEV') return null;
  var requestId = String(input.requestId || '').trim().toUpperCase();
  if (!/^[A-Z0-9][A-Z0-9._:-]{11,119}$/.test(requestId)) return null;
  var now = new Date().getTime();
  return {
    __coreLatencyTrace: true,
    environment: 'DEV',
    requestId: requestId,
    startedAtMs: now,
    lastAtMs: now,
    parentStartedAtMs: Number(input.parentStartedAtMs || 0),
    stages: [],
    counters: {
      oauthTokenRequests: 0,
      firestoreDocumentGets: 0,
      firestoreBatchGets: 0,
      firestoreCommits: 0
    }
  };
}

function coreFirestoreLatencyIncrementDev_(trace, counter, amount) {
  if (!trace || trace.environment !== 'DEV' || !trace.counters) return trace;
  if (!Object.prototype.hasOwnProperty.call(trace.counters, counter)) return trace;
  trace.counters[counter] = Math.max(
    0,
    Number(trace.counters[counter] || 0) + Number(
      amount === undefined ? 1 : amount
    )
  );
  return trace;
}

function coreFirestoreLatencySafeMetadata_(metadata) {
  var source = metadata || {};
  var safe = {};
  [
    'operation', 'durationMs', 'httpStatus', 'requested', 'result',
    'parentElapsedMs'
  ].forEach(function(key) {
    if (!Object.prototype.hasOwnProperty.call(source, key)) return;
    var value = source[key];
    if (typeof value === 'number') {
      safe[key] = Math.max(0, Math.round(value));
      return;
    }
    if (typeof value === 'boolean') {
      safe[key] = value;
      return;
    }
    safe[key] = String(value || '').slice(0, 80);
  });
  return safe;
}

function coreFirestoreLatencyMarkDev_(trace, code, label, metadata) {
  if (!trace || trace.environment !== 'DEV' || !trace.requestId) return trace;
  var now = new Date().getTime();
  var entry = Object.assign({
    event: 'GEAPA_PRESENTATIONS_DEV_LATENCY_V1',
    layer: 'CORE',
    requestId: trace.requestId,
    code: String(code || '').slice(0, 32),
    label: String(label || '').slice(0, 80),
    at: new Date(now).toISOString(),
    elapsedMs: Math.max(0, now - Number(trace.startedAtMs || now)),
    deltaMs: Math.max(0, now - Number(trace.lastAtMs || now))
  }, coreFirestoreLatencySafeMetadata_(metadata));
  if (trace.parentStartedAtMs) {
    entry.parentElapsedMs = Math.max(0, now - trace.parentStartedAtMs);
  }
  trace.lastAtMs = now;
  if (Array.isArray(trace.stages)) trace.stages.push(Object.freeze(entry));
  try {
    console.log(JSON.stringify(entry));
  } catch (ignored) {}
  return trace;
}

function coreFirestoreLatencySnapshotDev_(trace) {
  if (!trace || trace.environment !== 'DEV' || !trace.requestId) return null;
  return Object.freeze({
    environment: 'DEV',
    requestId: trace.requestId,
    startedAtMs: Number(trace.startedAtMs || 0),
    parentStartedAtMs: Number(trace.parentStartedAtMs || 0),
    stages: Object.freeze((trace.stages || []).map(function(entry) {
      return Object.freeze(Object.assign({}, entry));
    })),
    counters: Object.freeze(Object.assign({}, trace.counters || {}))
  });
}
