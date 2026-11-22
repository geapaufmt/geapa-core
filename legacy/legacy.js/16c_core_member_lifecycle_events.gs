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
  inputAliases: Object.freeze({
    eventStatus: Object.freeze(['eventStatus', 'statusEvento', 'status', 'STATUS_EVENTO', 'STATUS']),
    notes: Object.freeze(['notes', 'observacoes', 'OBSERVACOES']),
    updatedAt: Object.freeze(['updatedAt', 'atualizadoEm', 'ATUALIZADO_EM']),
    processedByModule: Object.freeze(['processedByModule', 'processadoPorModulo', 'PROCESSADO_POR_MODULO']),
    processingDate: Object.freeze(['processingDate', 'dataProcessamento', 'DATA_PROCESSAMENTO']),
    processingError: Object.freeze(['processingError', 'erroProcessamento', 'ERRO_PROCESSAMENTO'])
  }),
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
    updatedAt: Object.freeze(['ATUALIZADO_EM']),
    processedByModule: Object.freeze(['PROCESSADO_POR_MODULO']),
    processingDate: Object.freeze(['DATA_PROCESSAMENTO']),
    processingError: Object.freeze(['ERRO_PROCESSAMENTO'])
  }),
  optionalExtensionFields: Object.freeze([
    'processedByModule',
    'processingDate',
    'processingError'
  ])
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

function core_memberLifecycleGetSheetContext_(sheet, opts) {
  core_assertRequired_(sheet, 'sheet');
  opts = opts || {};

  var headerRow = Number(opts.headerRow || CORE_MEMBER_LIFECYCLE_CFG.headerRow);
  var startRow = headerRow + 1;
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var headers = lastCol > 0
    ? sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(function(value) {
        return String(value || '').trim();
      })
    : [];

  var headerMapZero = core_buildHeaderIndexMap_(headers, {
    normalize: true,
    oneBased: false,
    keepFirst: true
  });

  var headerMapOne = core_buildHeaderIndexMap_(headers, {
    normalize: true,
    oneBased: true,
    keepFirst: true
  });

  var rows = opts.includeRows === true && lastCol > 0 && lastRow >= startRow
    ? sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues()
    : [];

  return {
    sheet: sheet,
    headerRow: headerRow,
    startRow: startRow,
    lastRow: lastRow,
    lastCol: lastCol,
    headers: headers,
    headerMapZero: headerMapZero,
    headerMapOne: headerMapOne,
    rows: rows
  };
}

function core_memberLifecycleFindExistingHeader_(ctx, aliases) {
  var list = Array.isArray(aliases) ? aliases : [aliases];

  for (var i = 0; i < list.length; i++) {
    var colNumber = core_findHeaderIndex_(ctx.headerMapOne, list[i], {
      normalize: true,
      notFoundValue: 0
    });
    if (colNumber > 0) {
      return {
        headerName: ctx.headers[colNumber - 1],
        colNumber: colNumber
      };
    }
  }

  return null;
}

function core_memberLifecycleWriteCellByAliases_(sheet, rowNumber, ctx, aliases, value, required) {
  var target = core_memberLifecycleFindExistingHeader_(ctx, aliases);
  if (!target) {
    if (required === true) {
      throw new Error(
        'Schema de MEMBER_EVENTOS_VINCULO invalido: cabecalho nao encontrado para ' +
        (Array.isArray(aliases) ? aliases.join(', ') : aliases)
      );
    }
    return false;
  }

  sheet.getRange(Number(rowNumber), Number(target.colNumber)).setValue(value);
  return true;
}

function core_memberLifecycleEnsureSchemaExtensions_(sheet) {
  var ctx = core_memberLifecycleGetSheetContext_(sheet);
  var missing = [];

  CORE_MEMBER_LIFECYCLE_CFG.optionalExtensionFields.forEach(function(fieldName) {
    var aliases = CORE_MEMBER_LIFECYCLE_CFG.headers[fieldName];
    if (!core_memberLifecycleFindExistingHeader_(ctx, aliases)) {
      missing.push(aliases[0]);
    }
  });

  if (!missing.length) {
    return Object.freeze({
      ok: true,
      sheetName: typeof sheet.getName === 'function' ? sheet.getName() : '',
      addedHeaders: []
    });
  }

  var startColumn = ctx.lastCol + 1;
  sheet.getRange(ctx.headerRow, startColumn, 1, missing.length).setValues([missing]);

  return Object.freeze({
    ok: true,
    sheetName: typeof sheet.getName === 'function' ? sheet.getName() : '',
    addedHeaders: missing
  });
}

function core_memberLifecycleValidateStatus_(statusValue, label) {
  var normalized = core_memberLifecycleNormalizeStatus_(statusValue);
  if (!normalized) {
    throw new Error((label || 'STATUS_EVENTO') + ' obrigatorio.');
  }
  if (CORE_MEMBER_LIFECYCLE_CFG.eventStatuses.indexOf(normalized) < 0) {
    throw new Error(
      'Status de evento de vinculo nao suportado: ' + String(statusValue || '').trim() +
      '. Permitidos: ' + CORE_MEMBER_LIFECYCLE_CFG.eventStatuses.join(', ')
    );
  }
  return normalized;
}

function core_memberLifecyclePickOwnValue_(source, aliases) {
  var list = Array.isArray(aliases) ? aliases : [aliases];
  for (var i = 0; i < list.length; i++) {
    if (Object.prototype.hasOwnProperty.call(source, list[i])) {
      return {
        found: true,
        value: source[list[i]]
      };
    }
  }

  return {
    found: false,
    value: undefined
  };
}

function core_memberLifecycleNormalizePatchDate_(value, label, allowBlank) {
  if (value === '' || value === null || typeof value === 'undefined') {
    if (allowBlank) return '';
    throw new Error(label + ' invalido: informe uma data valida.');
  }

  var parsed = core_parseDateOrNull_(value);
  if (!parsed) {
    throw new Error(label + ' invalido: informe uma data valida.');
  }
  return parsed;
}

function core_memberLifecycleNormalizeUpdatePatch_(patch) {
  patch = patch || {};

  var allowedInputKeys = Object.create(null);
  Object.keys(CORE_MEMBER_LIFECYCLE_CFG.inputAliases).forEach(function(fieldName) {
    CORE_MEMBER_LIFECYCLE_CFG.inputAliases[fieldName].forEach(function(alias) {
      allowedInputKeys[alias] = true;
    });
  });

  var unsupportedKeys = Object.keys(patch).filter(function(key) {
    return !allowedInputKeys[key];
  });

  if (unsupportedKeys.length) {
    throw new Error(
      'Patch de lifecycle event contem campos nao permitidos: ' + unsupportedKeys.join(', ')
    );
  }

  var normalized = {};
  var touchedFields = [];

  var statusInput = core_memberLifecyclePickOwnValue_(patch, CORE_MEMBER_LIFECYCLE_CFG.inputAliases.eventStatus);
  if (statusInput.found) {
    normalized.eventStatus = core_memberLifecycleValidateStatus_(statusInput.value, 'STATUS_EVENTO');
    touchedFields.push('eventStatus');
  }

  var notesInput = core_memberLifecyclePickOwnValue_(patch, CORE_MEMBER_LIFECYCLE_CFG.inputAliases.notes);
  if (notesInput.found) {
    normalized.notes = String(notesInput.value || '').trim();
    touchedFields.push('notes');
  }

  var updatedAtInput = core_memberLifecyclePickOwnValue_(patch, CORE_MEMBER_LIFECYCLE_CFG.inputAliases.updatedAt);
  if (updatedAtInput.found) {
    normalized.updatedAt = core_memberLifecycleNormalizePatchDate_(updatedAtInput.value, 'ATUALIZADO_EM', false);
    touchedFields.push('updatedAt');
  }

  var processedByModuleInput = core_memberLifecyclePickOwnValue_(patch, CORE_MEMBER_LIFECYCLE_CFG.inputAliases.processedByModule);
  if (processedByModuleInput.found) {
    normalized.processedByModule = String(processedByModuleInput.value || '').trim();
    touchedFields.push('processedByModule');
  }

  var processingDateInput = core_memberLifecyclePickOwnValue_(patch, CORE_MEMBER_LIFECYCLE_CFG.inputAliases.processingDate);
  if (processingDateInput.found) {
    normalized.processingDate = core_memberLifecycleNormalizePatchDate_(processingDateInput.value, 'DATA_PROCESSAMENTO', true);
    touchedFields.push('processingDate');
  }

  var processingErrorInput = core_memberLifecyclePickOwnValue_(patch, CORE_MEMBER_LIFECYCLE_CFG.inputAliases.processingError);
  if (processingErrorInput.found) {
    normalized.processingError = String(processingErrorInput.value || '').trim();
    touchedFields.push('processingError');
  }

  return Object.freeze({
    values: normalized,
    touchedFields: Object.freeze(touchedFields.slice())
  });
}

function core_memberLifecycleValuesEqual_(left, right) {
  var leftIsDate = Object.prototype.toString.call(left) === '[object Date]' && !isNaN(left);
  var rightIsDate = Object.prototype.toString.call(right) === '[object Date]' && !isNaN(right);

  if (leftIsDate || rightIsDate) {
    if (!leftIsDate || !rightIsDate) return false;
    return left.getTime() === right.getTime();
  }

  return String(left == null ? '' : left).trim() === String(right == null ? '' : right).trim();
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
    processedByModule: String(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.processedByModule) || '').trim(),
    processingDate: core_parseDateOrNull_(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.processingDate)),
    processingError: String(core_memberLifecycleGetByAliases_(record, CORE_MEMBER_LIFECYCLE_CFG.headers.processingError) || '').trim(),
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
  var targetStatus = core_memberLifecycleValidateStatus_(payload.eventStatus || 'REGISTRADO', 'STATUS_EVENTO');
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
    eventStatus: core_memberLifecycleValidateStatus_(payload.eventStatus || payload.statusEvento || 'REGISTRADO', 'STATUS_EVENTO'),
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

function core_memberLifecycleFindEventByIdInContext_(ctx, eventId) {
  var normalizedEventId = String(eventId || '').trim();
  if (!normalizedEventId) {
    throw new Error('ID_EVENTO_MEMBRO obrigatorio.');
  }

  for (var i = 0; i < ctx.rows.length; i++) {
    var record = core_rowToObject_(ctx.headers, ctx.rows[i]);
    record.__rowNumber = ctx.startRow + i;
    var normalizedRecord = core_memberLifecycleNormalizeEventRecord_(record);
    if (normalizedRecord.eventId === normalizedEventId) {
      return normalizedRecord;
    }
  }

  return null;
}

function core_memberLifecycleApplyPatchToSheet_(sheet, eventId, patch) {
  core_assertRequired_(sheet, 'sheet');
  core_assertRequired_(eventId, 'eventId');

  core_memberLifecycleEnsureSchemaExtensions_(sheet);

  var normalizedPatch = core_memberLifecycleNormalizeUpdatePatch_(patch || {});
  var ctx = core_memberLifecycleGetSheetContext_(sheet, { includeRows: true });
  var currentEvent = core_memberLifecycleFindEventByIdInContext_(ctx, eventId);

  if (!currentEvent) {
    throw new Error('Evento nao encontrado em MEMBER_EVENTOS_VINCULO: ' + eventId);
  }

  if (!normalizedPatch.touchedFields.length) {
    return Object.freeze({
      ok: true,
      eventId: currentEvent.eventId,
      rowNumber: currentEvent.rowNumber,
      updated: false,
      event: currentEvent
    });
  }

  var writePlan = [];
  var values = normalizedPatch.values;

  if (Object.prototype.hasOwnProperty.call(values, 'eventStatus') &&
      !core_memberLifecycleValuesEqual_(currentEvent.eventStatus, values.eventStatus)) {
    writePlan.push({
      fieldName: 'eventStatus',
      aliases: CORE_MEMBER_LIFECYCLE_CFG.headers.status,
      value: values.eventStatus
    });
  }

  if (Object.prototype.hasOwnProperty.call(values, 'notes') &&
      !core_memberLifecycleValuesEqual_(currentEvent.notes, values.notes)) {
    writePlan.push({
      fieldName: 'notes',
      aliases: CORE_MEMBER_LIFECYCLE_CFG.headers.notes,
      value: values.notes
    });
  }

  if (Object.prototype.hasOwnProperty.call(values, 'processedByModule') &&
      !core_memberLifecycleValuesEqual_(currentEvent.processedByModule, values.processedByModule)) {
    writePlan.push({
      fieldName: 'processedByModule',
      aliases: CORE_MEMBER_LIFECYCLE_CFG.headers.processedByModule,
      value: values.processedByModule
    });
  }

  if (Object.prototype.hasOwnProperty.call(values, 'processingDate') &&
      !core_memberLifecycleValuesEqual_(currentEvent.processingDate, values.processingDate)) {
    writePlan.push({
      fieldName: 'processingDate',
      aliases: CORE_MEMBER_LIFECYCLE_CFG.headers.processingDate,
      value: values.processingDate
    });
  }

  if (Object.prototype.hasOwnProperty.call(values, 'processingError') &&
      !core_memberLifecycleValuesEqual_(currentEvent.processingError, values.processingError)) {
    writePlan.push({
      fieldName: 'processingError',
      aliases: CORE_MEMBER_LIFECYCLE_CFG.headers.processingError,
      value: values.processingError
    });
  }

  if (Object.prototype.hasOwnProperty.call(values, 'updatedAt')) {
    if (!core_memberLifecycleValuesEqual_(currentEvent.updatedAt, values.updatedAt)) {
      writePlan.push({
        fieldName: 'updatedAt',
        aliases: CORE_MEMBER_LIFECYCLE_CFG.headers.updatedAt,
        value: values.updatedAt
      });
    }
  } else if (writePlan.length) {
    writePlan.push({
      fieldName: 'updatedAt',
      aliases: CORE_MEMBER_LIFECYCLE_CFG.headers.updatedAt,
      value: new Date()
    });
  }

  if (!writePlan.length) {
    return Object.freeze({
      ok: true,
      eventId: currentEvent.eventId,
      rowNumber: currentEvent.rowNumber,
      updated: false,
      event: currentEvent
    });
  }

  writePlan.forEach(function(item) {
    core_memberLifecycleWriteCellByAliases_(sheet, currentEvent.rowNumber, ctx, item.aliases, item.value, true);
  });

  var refreshedCtx = core_memberLifecycleGetSheetContext_(sheet, { includeRows: true });
  var refreshedEvent = core_memberLifecycleFindEventByIdInContext_(refreshedCtx, eventId);
  var changedFields = writePlan.map(function(item) {
    return item.fieldName;
  });

  return Object.freeze({
    ok: true,
    eventId: currentEvent.eventId,
    rowNumber: currentEvent.rowNumber,
    updated: true,
    changedFields: Object.freeze(changedFields),
    event: refreshedEvent
  });
}

function core_updateMemberLifecycleEvent_(eventId, patch) {
  core_assertRequired_(eventId, 'eventId');

  return core_withLock_('CORE_MEMBER_LIFECYCLE_UPDATE_EVENT', function() {
    return core_memberLifecycleApplyPatchToSheet_(
      core_memberLifecycleGetSheet_(),
      String(eventId || '').trim(),
      patch || {}
    );
  });
}

function core_updateMemberLifecycleEventStatus_(eventId, nextStatus, opts) {
  opts = opts || {};

  var patch = {
    eventStatus: nextStatus
  };

  if (Object.prototype.hasOwnProperty.call(opts, 'notes')) patch.notes = opts.notes;
  if (Object.prototype.hasOwnProperty.call(opts, 'observacoes')) patch.notes = opts.observacoes;
  if (Object.prototype.hasOwnProperty.call(opts, 'updatedAt')) patch.updatedAt = opts.updatedAt;
  if (Object.prototype.hasOwnProperty.call(opts, 'atualizadoEm')) patch.updatedAt = opts.atualizadoEm;
  if (Object.prototype.hasOwnProperty.call(opts, 'processedByModule')) patch.processedByModule = opts.processedByModule;
  if (Object.prototype.hasOwnProperty.call(opts, 'processadoPorModulo')) patch.processedByModule = opts.processadoPorModulo;
  if (Object.prototype.hasOwnProperty.call(opts, 'processingDate')) patch.processingDate = opts.processingDate;
  if (Object.prototype.hasOwnProperty.call(opts, 'dataProcessamento')) patch.processingDate = opts.dataProcessamento;
  if (Object.prototype.hasOwnProperty.call(opts, 'processingError')) patch.processingError = opts.processingError;
  if (Object.prototype.hasOwnProperty.call(opts, 'erroProcessamento')) patch.processingError = opts.erroProcessamento;

  return core_updateMemberLifecycleEvent_(eventId, patch);
}


