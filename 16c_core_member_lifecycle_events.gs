/**
 * Registro compartilhado de eventos de vinculo de membros.
 *
 * Objetivo:
 * - centralizar ingressos, desligamentos, suspensoes e retornos por RGA;
 * - permitir que modulos produtores e consumidores usem o mesmo contrato;
 * - manter trilha historica simples e reutilizavel no core.
 */

const CORE_MEMBER_LIFECYCLE_CFG = Object.freeze({
  registryKey: 'MEMBER_EVENTOS_VINCULO',
  headerRow: 1,
  idPrefix: 'MEV-',
  idPadLength: 6,
  eventTypes: Object.freeze([
    'INGRESSO',
    'DESLIGAMENTO_VOLUNTARIO',
    'DESLIGAMENTO_POR_FALTAS',
    'DESLIGAMENTO_ADMINISTRATIVO',
    'SUSPENSAO',
    'RETORNO'
  ]),
  eventStatuses: Object.freeze([
    'REGISTRADO',
    'HOMOLOGADO',
    'CANCELADO',
    'PROCESSADO_ATIVIDADES',
    'PROCESSADO_MEMBROS'
  ]),
  headers: Object.freeze({
    id: Object.freeze(['ID_EVENTO_MEMBRO']),
    rga: Object.freeze(['RGA']),
    type: Object.freeze(['TIPO_EVENTO', 'TIPO_EVENTO_VINCULO']),
    date: Object.freeze(['DATA_EVENTO', 'DATA_HOMOLOGACAO']),
    status: Object.freeze(['STATUS_EVENTO', 'STATUS']),
    reason: Object.freeze(['MOTIVO_EVENTO', 'MOTIVO_DESLIGAMENTO', 'Motivo']),
    sourceModule: Object.freeze(['ORIGEM_MODULO']),
    sourceKey: Object.freeze(['ORIGEM_CHAVE']),
    sourceRow: Object.freeze(['ORIGEM_ROW']),
    name: Object.freeze(['NOME_MEMBRO', 'MEMBRO', 'Membro', 'Nome']),
    email: Object.freeze(['EMAIL', 'E-mail', 'Email']),
    notes: Object.freeze(['OBSERVACOES', 'Observacoes']),
    createdAt: Object.freeze(['CRIADO_EM']),
    updatedAt: Object.freeze(['ATUALIZADO_EM'])
  })
});

function core_memberLifecycleNormalizeToken_(value) {
  return String(value == null ? '' : value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function core_memberLifecycleNormalizeType_(value) {
  return core_memberLifecycleNormalizeToken_(value);
}

function core_memberLifecycleNormalizeStatus_(value) {
  return core_memberLifecycleNormalizeToken_(value);
}

function core_memberLifecycleGetSheet_() {
  var sheet = core_getSheetByKey_(CORE_MEMBER_LIFECYCLE_CFG.registryKey);
  if (!sheet) {
    throw new Error('MEMBER_EVENTOS_VINCULO nao encontrada no Registry.');
  }
  return sheet;
}

function core_memberLifecycleGetByAliases_(record, aliases) {
  var list = Array.isArray(aliases) ? aliases : [aliases];
  var keys = Object.keys(record || {});

  for (var i = 0; i < list.length; i++) {
    var target = core_memberLifecycleNormalizeToken_(list[i]);
    for (var j = 0; j < keys.length; j++) {
      if (core_memberLifecycleNormalizeToken_(keys[j]) === target) {
        return record[keys[j]];
      }
    }
  }

  return '';
}

function core_memberLifecycleNormalizeEventRecord_(record) {
  return Object.freeze({
    eventId: String(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.id) || '').trim(),
    rga: String(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.rga) || '').trim(),
    eventType: core_memberLifecycleNormalizeType_(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.type)),
    eventDate: core_parseDateOrNull_(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.date)),
    eventStatus: core_memberLifecycleNormalizeStatus_(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.status)),
    reason: String(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.reason) || '').trim(),
    sourceModule: String(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.sourceModule) || '').trim(),
    sourceKey: String(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.sourceKey) || '').trim(),
    sourceRow: String(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.sourceRow) || '').trim(),
    memberName: String(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.name) || '').trim(),
    memberEmail: String(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.email) || '').trim(),
    notes: String(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.notes) || '').trim(),
    createdAt: core_parseDateOrNull_(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.createdAt)),
    updatedAt: core_parseDateOrNull_(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.updatedAt)),
    rowNumber: record.__rowNumber || null,
    raw: record
  });
}

function core_memberLifecycleListEvents_(filters, opts) {
  filters = filters || {};
  opts = opts || {};

  var records = core_readRecordsByKey_(CORE_MEMBER_LIFECYCLE_CFG.registryKey, {
    headerRow: CORE_MEMBER_LIFECYCLE_CFG.headerRow,
    skipBlankRows: true
  }).map(core_memberLifecycleNormalizeEventRecord_);

  var filtered = records.filter(function(event) {
    if (filters.rga && String(event.rga || '').trim() !== String(filters.rga || '').trim()) return false;
    if (filters.eventType && event.eventType !== core_memberLifecycleNormalizeType_(filters.eventType)) return false;
    if (filters.eventStatus && event.eventStatus !== core_memberLifecycleNormalizeStatus_(filters.eventStatus)) return false;
    if (filters.sourceModule && core_memberLifecycleNormalizeToken_(event.sourceModule) !== core_memberLifecycleNormalizeToken_(filters.sourceModule)) return false;
    if (filters.sourceRow && String(event.sourceRow || '').trim() !== String(filters.sourceRow || '').trim()) return false;
    if (filters.excludeCanceled && event.eventStatus === 'CANCELADO') return false;
    return true;
  });

  filtered.sort(function(a, b) {
    var aTime = a.eventDate ? a.eventDate.getTime() : 0;
    var bTime = b.eventDate ? b.eventDate.getTime() : 0;
    if (aTime !== bTime) return aTime - bTime;
    return String(a.eventId || '').localeCompare(String(b.eventId || ''), 'pt-BR');
  });

  if (opts.limit && filtered.length > opts.limit) {
    return filtered.slice(Math.max(0, filtered.length - Number(opts.limit)));
  }

  return filtered;
}

function core_memberLifecycleGetLatestEventByRga_(rga, opts) {
  var events = core_memberLifecycleListEvents_({
    rga: rga,
    excludeCanceled: !(opts && opts.includeCanceled)
  }, opts || {});
  return events.length ? events[events.length - 1] : null;
}

function core_memberLifecycleExtractNumericSuffix_(idValue) {
  var match = String(idValue || '').trim().match(/^MEV-(\d+)$/i);
  return match ? Number(match[1]) : 0;
}

function core_memberLifecycleGenerateNextId_(records) {
  var maxValue = 0;
  (records || []).forEach(function(record) {
    var current = core_memberLifecycleExtractNumericSuffix_(record.eventId || record.ID_EVENTO_MEMBRO || '');
    if (current > maxValue) maxValue = current;
  });

  var next = String(maxValue + 1);
  while (next.length < CORE_MEMBER_LIFECYCLE_CFG.idPadLength) {
    next = '0' + next;
  }
  return CORE_MEMBER_LIFECYCLE_CFG.idPrefix + next;
}

function core_memberLifecycleFindDuplicate_(events, payload) {
  var targetType = core_memberLifecycleNormalizeType_(payload.eventType);
  var targetStatus = core_memberLifecycleNormalizeStatus_(payload.eventStatus || 'REGISTRADO');
  var targetRga = String(payload.rga || '').trim();
  var targetDate = core_parseDateOrNull_(payload.eventDate);
  var targetSourceModule = String(payload.sourceModule || '').trim();
  var targetSourceKey = String(payload.sourceKey || '').trim();
  var targetSourceRow = String(payload.sourceRow || '').trim();

  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    if (!event || event.rga !== targetRga || event.eventType !== targetType) continue;

    if (targetSourceModule && targetSourceRow) {
      if (
        core_memberLifecycleNormalizeToken_(event.sourceModule) === core_memberLifecycleNormalizeToken_(targetSourceModule) &&
        String(event.sourceRow || '').trim() === targetSourceRow &&
        core_memberLifecycleNormalizeToken_(event.sourceKey) === core_memberLifecycleNormalizeToken_(targetSourceKey)
      ) {
        return event;
      }
    }

    if (
      targetDate && event.eventDate &&
      event.eventDate.getTime() === targetDate.getTime() &&
      event.eventStatus === targetStatus &&
      core_memberLifecycleNormalizeToken_(event.sourceModule) === core_memberLifecycleNormalizeToken_(targetSourceModule)
    ) {
      return event;
    }
  }

  return null;
}

function core_appendMemberLifecycleEvent_(payload) {
  core_assertRequired_(payload, 'payload');

  var normalizedPayload = {
    rga: String(payload.rga || payload.memberRga || '').trim(),
    eventType: core_memberLifecycleNormalizeType_(payload.eventType || payload.tipoEvento),
    eventDate: core_parseDateOrNull_(payload.eventDate || payload.dataEvento || payload.approvedAt || payload.integratedAt || payload.createdAt),
    eventStatus: core_memberLifecycleNormalizeStatus_(payload.eventStatus || payload.statusEvento || 'REGISTRADO'),
    reason: String(payload.reason || payload.motivoEvento || '').trim(),
    sourceModule: String(payload.sourceModule || payload.origemModulo || '').trim(),
    sourceKey: String(payload.sourceKey || payload.origemChave || '').trim(),
    sourceRow: String(payload.sourceRow || payload.origemRow || '').trim(),
    memberName: String(payload.memberName || payload.nomeMembro || '').trim(),
    memberEmail: String(payload.memberEmail || payload.email || '').trim(),
    notes: String(payload.notes || payload.observacoes || '').trim()
  };

  if (!normalizedPayload.rga) {
    throw new Error('core_appendMemberLifecycleEvent_: RGA obrigatorio.');
  }
  if (!normalizedPayload.eventType) {
    throw new Error('core_appendMemberLifecycleEvent_: eventType obrigatorio.');
  }
  if (!normalizedPayload.eventDate) {
    throw new Error('core_appendMemberLifecycleEvent_: eventDate obrigatorio.');
  }

  var allEvents = core_memberLifecycleListEvents_({}, { includeCanceled: true });
  var currentEvents = allEvents.filter(function(event) {
    return String(event.rga || '').trim() === normalizedPayload.rga;
  });
  var duplicate = core_memberLifecycleFindDuplicate_(currentEvents, normalizedPayload);
  if (duplicate) {
    return {
      ok: true,
      duplicated: true,
      eventId: duplicate.eventId,
      rowNumber: duplicate.rowNumber,
      event: duplicate
    };
  }

  var nextId = core_memberLifecycleGenerateNextId_(allEvents);
  var sheet = core_memberLifecycleGetSheet_();
  var now = new Date();

  core_appendObjectByHeaders_(sheet, {
    ID_EVENTO_MEMBRO: nextId,
    RGA: normalizedPayload.rga,
    TIPO_EVENTO: normalizedPayload.eventType,
    DATA_EVENTO: normalizedPayload.eventDate,
    STATUS_EVENTO: normalizedPayload.eventStatus,
    MOTIVO_EVENTO: normalizedPayload.reason,
    ORIGEM_MODULO: normalizedPayload.sourceModule,
    ORIGEM_CHAVE: normalizedPayload.sourceKey,
    ORIGEM_ROW: normalizedPayload.sourceRow,
    NOME_MEMBRO: normalizedPayload.memberName,
    EMAIL: normalizedPayload.memberEmail,
    OBSERVACOES: normalizedPayload.notes,
    CRIADO_EM: now,
    ATUALIZADO_EM: now
  }, {
    headerRow: CORE_MEMBER_LIFECYCLE_CFG.headerRow
  });

  return {
    ok: true,
    duplicated: false,
    eventId: nextId,
    eventType: normalizedPayload.eventType,
    rga: normalizedPayload.rga
  };
}


