/**
 * ============================================================
 * 17_core_mail_hub.js
 * ============================================================
 *
 * Mail Hub V1
 *
 * Objetivo desta camada:
 * - ler mensagens do Gmail;
 * - registrar eventos em planilhas centrais;
 * - deduplicar por Id Mensagem Gmail;
 * - manter indice por chave de correlacao;
 * - registrar anexos recebidos;
 * - expor consultas minimas para posterior consumo por modulos.
 *
 * Fora de escopo nesta V1:
 * - migracao dos modulos consumidores;
 * - fila de envio;
 * - salvamento de anexos no Drive;
 * - roteamento complexo por regras.
 */

var CORE_MAIL_HUB_KEYS = Object.freeze({
  EVENTOS: 'MAIL_EVENTOS',
  INDICE: 'MAIL_INDICE',
  SAIDA: 'MAIL_SAIDA',
  ANEXOS: 'MAIL_ANEXOS',
  REGRAS: 'MAIL_REGRAS',
  CONFIG: 'MAIL_CONFIG'
});

var CORE_MAIL_HUB_STATUS = Object.freeze({
  PENDENTE: 'PENDENTE',
  PROCESSADO: 'PROCESSADO',
  IGNORADO: 'IGNORADO'
});

var CORE_MAIL_HUB_DEFAULTS = Object.freeze({
  query: 'in:inbox',
  maxThreads: 25,
  maxMessagesPerThread: 100,
  start: 0
});

var CORE_MAIL_HUB_SCHEMA = Object.freeze({
  MAIL_EVENTOS: Object.freeze([
    'Id Evento',
    'Data Hora Evento',
    'Direcao',
    'Tipo Evento',
    'Modulo Dono',
    'Chave de Correlacao',
    'Id Thread Gmail',
    'Id Mensagem Gmail',
    'Assunto',
    'Email Remetente',
    'Emails Destinatarios',
    'Status Processamento',
    'Processado Por',
    'Data Hora Processamento',
    'Possui Anexos',
    'Quantidade Anexos',
    'Criado Em',
    'Atualizado Em'
  ]),
  MAIL_INDICE: Object.freeze([
    'Chave de Correlacao',
    'Modulo Dono',
    'Tipo Entidade',
    'Id Entidade',
    'Etapa Atual',
    'Id Thread Gmail',
    'Id Ultima Mensagem',
    'Ultima Direcao',
    'Ultimo Tipo Evento',
    'Ultimo Email Remetente',
    'Ultimo Assunto',
    'Data Hora Ultimo Evento',
    'Ha Entrada Pendente',
    'Ha Anexo Pendente',
    'Quantidade Eventos',
    'Quantidade Entradas',
    'Quantidade Saidas',
    'Quantidade Anexos',
    'Criado Em',
    'Atualizado Em'
  ]),
  MAIL_ANEXOS: Object.freeze([
    'Id Anexo',
    'Id Evento',
    'Modulo Dono',
    'Chave de Correlacao',
    'Etapa Fluxo',
    'Id Mensagem Gmail',
    'Id Thread Gmail',
    'Nome Arquivo',
    'Tipo Mime',
    'Tamanho Bytes',
    'Status Anexo',
    'Criado Em',
    'Atualizado Em'
  ]),
  MAIL_CONFIG: Object.freeze([
    'Chave',
    'Valor',
    'Ativo'
  ])
});

function coreMailHubGetEventosSheet_() {
  return core_getSheetByKey_(CORE_MAIL_HUB_KEYS.EVENTOS);
}

function coreMailHubGetIndiceSheet_() {
  return core_getSheetByKey_(CORE_MAIL_HUB_KEYS.INDICE);
}

function coreMailHubGetSaidaSheet_() {
  return core_getSheetByKey_(CORE_MAIL_HUB_KEYS.SAIDA);
}

function coreMailHubGetAnexosSheet_() {
  return core_getSheetByKey_(CORE_MAIL_HUB_KEYS.ANEXOS);
}

function coreMailHubGetRegrasSheet_() {
  return core_getSheetByKey_(CORE_MAIL_HUB_KEYS.REGRAS);
}

function coreMailHubGetConfigSheet_() {
  return core_getSheetByKey_(CORE_MAIL_HUB_KEYS.CONFIG);
}

function coreMailHubGetSheetContext_(sheet, opts) {
  core_assertRequired_(sheet, 'sheet');
  opts = opts || {};

  var headerRow = Number(opts.headerRow || 1);
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

function coreMailHubFindHeaderIndexZero_(ctx, headerName) {
  return core_findHeaderIndex_(ctx.headerMapZero, headerName, {
    normalize: true,
    notFoundValue: -1
  });
}

function coreMailHubGetRowValue_(row, ctx, headerName, defaultValue) {
  var idx = coreMailHubFindHeaderIndexZero_(ctx, headerName);
  if (idx < 0) {
    return typeof defaultValue === 'undefined' ? '' : defaultValue;
  }
  return row[idx];
}

function coreMailHubSetRowValue_(row, ctx, headerName, value) {
  var idx = coreMailHubFindHeaderIndexZero_(ctx, headerName);
  if (idx < 0) return false;
  row[idx] = value;
  return true;
}

function coreMailHubAppendRow_(sheet, ctx, payload) {
  var row = new Array(ctx.headers.length);
  for (var i = 0; i < row.length; i++) row[i] = '';

  Object.keys(payload || {}).forEach(function(headerName) {
    coreMailHubSetRowValue_(row, ctx, headerName, payload[headerName]);
  });

  sheet.appendRow(row);
  return row;
}

function coreMailHubWriteCell_(sheet, rowNumber, ctx, headerName, value) {
  return core_writeCellByHeader_(sheet, rowNumber, ctx.headerMapOne, headerName, value, {
    normalize: true,
    oneBased: true
  });
}

function coreMailHubAssertSheetSchema_(sheetKey, sheet, requiredHeaders) {
  var ctx = coreMailHubGetSheetContext_(sheet);
  if (!ctx.headers.length) {
    throw new Error(
      'Mail Hub schema invalido em "' + sheetKey + '": a aba precisa ter cabecalhos na linha 1.'
    );
  }

  var missing = requiredHeaders.filter(function(headerName) {
    return coreMailHubFindHeaderIndexZero_(ctx, headerName) < 0;
  });

  if (missing.length) {
    throw new Error(
      'Mail Hub schema invalido em "' + sheetKey + '" (' + sheet.getName() + '). ' +
      'Cabecalhos obrigatorios ausentes: ' + missing.join(', ')
    );
  }

  return Object.freeze({
    key: sheetKey,
    sheetName: sheet.getName(),
    requiredHeaders: requiredHeaders.slice(),
    headerCount: ctx.headers.length
  });
}

function coreMailHubAssertSchema_() {
  var results = [
    coreMailHubAssertSheetSchema_(
      CORE_MAIL_HUB_KEYS.EVENTOS,
      coreMailHubGetEventosSheet_(),
      CORE_MAIL_HUB_SCHEMA.MAIL_EVENTOS
    ),
    coreMailHubAssertSheetSchema_(
      CORE_MAIL_HUB_KEYS.INDICE,
      coreMailHubGetIndiceSheet_(),
      CORE_MAIL_HUB_SCHEMA.MAIL_INDICE
    ),
    coreMailHubAssertSheetSchema_(
      CORE_MAIL_HUB_KEYS.ANEXOS,
      coreMailHubGetAnexosSheet_(),
      CORE_MAIL_HUB_SCHEMA.MAIL_ANEXOS
    ),
    coreMailHubAssertSheetSchema_(
      CORE_MAIL_HUB_KEYS.CONFIG,
      coreMailHubGetConfigSheet_(),
      CORE_MAIL_HUB_SCHEMA.MAIL_CONFIG
    )
  ];

  return Object.freeze({
    ok: true,
    validatedAt: new Date(),
    sheets: results
  });
}

function coreMailHubNormalizeFlag_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  });
}

function coreMailHubGetConfigMap_() {
  var sheet = coreMailHubGetConfigSheet_();
  var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
  var out = Object.create(null);

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var key = core_normalizeText_(coreMailHubGetRowValue_(row, ctx, 'Chave', ''), {
      collapseWhitespace: true,
      caseMode: 'upper'
    });

    if (!key) continue;

    var ativo = coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Ativo', 'SIM'));
    if (ativo && ativo !== 'SIM' && ativo !== 'TRUE' && ativo !== '1') continue;

    out[key] = coreMailHubGetRowValue_(row, ctx, 'Valor', '');
  }

  return out;
}

function coreMailHubGetConfigValue_(configMap, key, defaultValue) {
  var normalizedKey = core_normalizeText_(key, {
    collapseWhitespace: true,
    caseMode: 'upper'
  });

  if (!Object.prototype.hasOwnProperty.call(configMap, normalizedKey)) {
    return defaultValue;
  }

  var value = configMap[normalizedKey];
  if (value === '' || value === null || typeof value === 'undefined') {
    return defaultValue;
  }

  return value;
}

function coreMailHubGetConfigNumber_(configMap, key, defaultValue) {
  var raw = coreMailHubGetConfigValue_(configMap, key, defaultValue);
  var num = Number(raw);
  return isNaN(num) ? Number(defaultValue) : num;
}

function coreMailHubGetConfigBoolean_(configMap, key, defaultValue) {
  var raw = coreMailHubNormalizeFlag_(coreMailHubGetConfigValue_(configMap, key, defaultValue ? 'SIM' : 'NAO'));
  if (raw === 'SIM' || raw === 'TRUE' || raw === '1') return true;
  if (raw === 'NAO' || raw === 'FALSE' || raw === '0') return false;
  return defaultValue === true;
}

function coreMailHubParseConfigList_(value) {
  return String(value || '')
    .split(/[\r\n,;]+/)
    .map(function(item) { return String(item || '').trim(); })
    .filter(function(item) { return !!item; });
}

function coreMailHubGetConfigList_(configMap, key) {
  return coreMailHubParseConfigList_(coreMailHubGetConfigValue_(configMap, key, ''));
}

function coreMailHubGetConfigRegexList_(configMap, key) {
  return coreMailHubGetConfigList_(configMap, key).map(function(pattern) {
    try {
      return new RegExp(pattern, 'i');
    } catch (err) {
      throw new Error('Regex invalida em MAIL_CONFIG para "' + key + '": ' + pattern);
    }
  });
}

function coreMailHubGetConfig_(key, defaultValue) {
  return coreMailHubGetConfigValue_(coreMailHubGetConfigMap_(), key, defaultValue);
}

function coreMailHubGetConfigBooleanByKey_(key, defaultValue) {
  return coreMailHubGetConfigBoolean_(coreMailHubGetConfigMap_(), key, defaultValue === true);
}

function coreMailHubGetConfigListByKey_(key) {
  return coreMailHubGetConfigList_(coreMailHubGetConfigMap_(), key);
}

function coreMailHubBuildIngestConfig_(configMap, opts) {
  opts = opts || {};

  var useOnlyTaggedSubjects = Object.prototype.hasOwnProperty.call(opts, 'useOnlyTaggedSubjects')
    ? opts.useOnlyTaggedSubjects === true
    : coreMailHubGetConfigBoolean_(configMap, 'USAR_SOMENTE_ASSUNTOS_GEAPA', false);
  var requiredSubjectPrefix = String(coreMailHubGetConfigValue_(configMap, 'ASSUNTO_PREFIXO_OBRIGATORIO', '[GEAPA]') || '').trim();

  return {
    requiredSubjectPrefix: requiredSubjectPrefix || '[GEAPA]',
    useOnlyTaggedSubjects: useOnlyTaggedSubjects,
    ignoredSenders: coreMailHubGetConfigList_(configMap, 'IGNORAR_REMETENTES').map(function(item) {
      return core_extractEmailAddress_(item);
    }).filter(function(item) {
      return !!item;
    }),
    ignoredDomains: coreMailHubGetConfigList_(configMap, 'IGNORAR_DOMINIOS').map(function(item) {
      return core_normalizeText_(item, {
        collapseWhitespace: true,
        caseMode: 'lower'
      });
    }).filter(function(item) {
      return !!item;
    }),
    ignoredSubjectRegexes: coreMailHubGetConfigRegexList_(configMap, 'IGNORAR_ASSUNTOS_REGEX'),
    maxEventsPerExecution: Math.max(
      0,
      Number(
        Object.prototype.hasOwnProperty.call(opts, 'maxEventsPerExecution')
          ? opts.maxEventsPerExecution
          : coreMailHubGetConfigNumber_(configMap, 'MAX_EVENTOS_POR_EXECUCAO', 0)
      ) || 0
    ),
    saveFullBody: Object.prototype.hasOwnProperty.call(opts, 'saveFullBody')
      ? opts.saveFullBody === true
      : coreMailHubGetConfigBoolean_(configMap, 'SALVAR_CORPO_COMPLETO', false),
    markNoiseAsIgnored: Object.prototype.hasOwnProperty.call(opts, 'markNoiseAsIgnored')
      ? opts.markNoiseAsIgnored === true
      : coreMailHubGetConfigBoolean_(configMap, 'MARCAR_RUIDO_COMO_IGNORADO', false)
  };
}

function coreMailHubBuildEventosState_(sheet) {
  var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
  var messageIds = Object.create(null);
  var rowByEventId = Object.create(null);

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var messageId = String(coreMailHubGetRowValue_(row, ctx, 'Id Mensagem Gmail', '') || '').trim();
    var eventId = String(coreMailHubGetRowValue_(row, ctx, 'Id Evento', '') || '').trim();
    var rowNumber = ctx.startRow + i;

    if (messageId) messageIds[messageId] = true;
    if (eventId) rowByEventId[eventId] = rowNumber;
  }

  return {
    ctx: ctx,
    messageIds: messageIds,
    rowByEventId: rowByEventId
  };
}

function coreMailHubBuildIndiceState_(sheet) {
  var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
  var rowByCorrelationKey = Object.create(null);
  var recordByCorrelationKey = Object.create(null);

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var correlationKey = core_normalizeText_(
      coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', ''),
      { collapseWhitespace: true, caseMode: 'upper' }
    );

    if (!correlationKey) continue;

    rowByCorrelationKey[correlationKey] = ctx.startRow + i;
    recordByCorrelationKey[correlationKey] = row;
  }

  return {
    ctx: ctx,
    rowByCorrelationKey: rowByCorrelationKey,
    recordByCorrelationKey: recordByCorrelationKey
  };
}

function coreMailEventExistsByMessageId_(messageId, eventosState) {
  var key = String(messageId || '').trim();
  if (!key) return false;
  return eventosState.messageIds[key] === true;
}

function coreMailExtractCorrelationKey_(subject) {
  var text = String(subject || '').trim();
  if (!text) return '';

  var match = text.match(/\[GEAPA\]\[([^\]]+)\]/i);
  if (!match) return '';

  return String(match[1] || '').trim().toUpperCase();
}

function coreMailResolveModule_(correlationKey) {
  var parsed = coreMailParseCorrelationKey_(correlationKey);
  return parsed.isValid ? String(parsed.moduleName || '').trim() : 'NAO_IDENTIFICADO';
}

function coreMailHubBuildSnippet_(message) {
  try {
    var plainBody = String(message.getPlainBody() || '').replace(/\s+/g, ' ').trim();
    return plainBody.slice(0, 500);
  } catch (err) {
    return '';
  }
}

function coreMailHubGetThreadLabels_(thread) {
  var labels = thread.getLabels();
  var out = [];

  for (var i = 0; i < labels.length; i++) {
    out.push(String(labels[i].getName() || '').trim());
  }

  return out.join(' | ');
}

function coreMailHubGetEmailDomain_(email) {
  var normalized = core_extractEmailAddress_(email);
  if (!normalized || normalized.indexOf('@') === -1) return '';
  return normalized.split('@')[1];
}

function coreMailHubSubjectHasRequiredPrefix_(subject, ingestConfig) {
  var prefix = String((ingestConfig && ingestConfig.requiredSubjectPrefix) || '[GEAPA]').trim();
  if (!prefix) return true;
  return String(subject || '').trim().toUpperCase().indexOf(prefix.toUpperCase()) === 0;
}

function coreMailHubDetectNoise_(msgCtx, ingestConfig) {
  var fromEmail = core_extractEmailAddress_(msgCtx.fromEmail || msgCtx.fromRaw || '');
  var fromDomain = coreMailHubGetEmailDomain_(fromEmail);
  var subject = String(msgCtx.subject || '').trim();
  var lowerSubject = subject.toLowerCase();
  var lowerFrom = String(fromEmail || '').toLowerCase();
  var technicalPatterns = [
    /github/i,
    /\bcodex\b/i,
    /\b(alerta|alert|incident|workflow|pull request|issue|build failed|monitor|sentry)\b/i
  ];

  if (ingestConfig.ignoredSenders.indexOf(fromEmail) !== -1) {
    return { isNoise: true, reason: 'IGNORAR_REMETENTES' };
  }

  if (fromDomain && ingestConfig.ignoredDomains.indexOf(String(fromDomain || '').toLowerCase()) !== -1) {
    return { isNoise: true, reason: 'IGNORAR_DOMINIOS' };
  }

  for (var i = 0; i < ingestConfig.ignoredSubjectRegexes.length; i++) {
    if (ingestConfig.ignoredSubjectRegexes[i].test(subject)) {
      return { isNoise: true, reason: 'IGNORAR_ASSUNTOS_REGEX' };
    }
  }

  if (ingestConfig.useOnlyTaggedSubjects && !coreMailHubSubjectHasRequiredPrefix_(subject, ingestConfig)) {
    return { isNoise: true, reason: 'ASSUNTO_SEM_PREFIXO_OBRIGATORIO' };
  }

  if (
    fromDomain === 'github.com' ||
    fromDomain === 'githubusercontent.com' ||
    lowerFrom.indexOf('github') !== -1 ||
    lowerFrom.indexOf('codex') !== -1
  ) {
    return { isNoise: true, reason: 'REMETENTE_TECNICO' };
  }

  for (var j = 0; j < technicalPatterns.length; j++) {
    if (technicalPatterns[j].test(lowerSubject)) {
      return { isNoise: true, reason: 'ASSUNTO_TECNICO' };
    }
  }

  return { isNoise: false, reason: '' };
}

function coreMailHubBuildMessageContext_(thread, message) {
  var subject = String(message.getSubject() || '').trim();
  var fromRaw = String(message.getFrom() || '').trim();
  var msgCtx = {
    subject: subject,
    fromRaw: fromRaw,
    fromEmail: core_extractEmailAddress_(fromRaw),
    fromName: core_extractDisplayName_(fromRaw),
    to: String(message.getTo() || '').trim(),
    cc: String(message.getCc() || '').trim(),
    bcc: String(message.getBcc() || '').trim(),
    replyTo: String(message.getReplyTo() || '').trim(),
    messageId: String(message.getId() || '').trim(),
    threadId: String(thread.getId() || '').trim(),
    messageDate: message.getDate() || '',
    snippet: coreMailHubBuildSnippet_(message),
    plainBody: '',
    labels: coreMailHubGetThreadLabels_(thread),
    attachments: message.getAttachments() || []
  };

  try {
    msgCtx.plainBody = String(message.getPlainBody() || '');
  } catch (err) {
    msgCtx.plainBody = '';
  }

  return msgCtx;
}

function coreMailHubBuildEventPayload_(thread, message, ingestedAt, ingestConfig) {
  var msgCtx = coreMailHubBuildMessageContext_(thread, message);
  var noise = coreMailHubDetectNoise_(msgCtx, ingestConfig || coreMailHubBuildIngestConfig_(coreMailHubGetConfigMap_(), {}));

  var routing = coreMailResolveRouting_(msgCtx);
  var correlationKey = String(routing.correlationKey || coreMailExtractCorrelationKey_(msgCtx.subject) || '').trim().toUpperCase();
  var parsedCorrelation = correlationKey ? coreMailParseCorrelationKey_(correlationKey) : { isValid: false };
  var attachments = msgCtx.attachments || [];
  var shouldIgnore = noise.isNoise && ingestConfig && ingestConfig.markNoiseAsIgnored === true;

  return {
    eventId: 'MEV-' + msgCtx.messageId,
    messageId: msgCtx.messageId,
    threadId: msgCtx.threadId,
    messageDate: msgCtx.messageDate,
    ingestedAt: ingestedAt || new Date(),
    direction: 'ENTRADA',
    eventType: 'EMAIL_RECEBIDO',
    flowStep: String(routing.stage || parsedCorrelation.stage || routing.flowCode || parsedCorrelation.flowCode || 'INGESTAO').trim(),
    routingStatus: shouldIgnore ? CORE_MAIL_HUB_STATUS.IGNORADO : (routing.matched ? 'ROTEADO' : 'NAO_IDENTIFICADO'),
    subject: msgCtx.subject,
    fromRaw: msgCtx.fromRaw,
    fromEmail: msgCtx.fromEmail,
    fromName: msgCtx.fromName,
    to: msgCtx.to,
    cc: msgCtx.cc,
    bcc: msgCtx.bcc,
    replyTo: msgCtx.replyTo,
    correlationKey: correlationKey,
    correlationPrefix: coreMailHubExtractCorrelationPrefix_(correlationKey),
    moduleName: routing.matched
      ? String(routing.moduleName || '').trim()
      : (parsedCorrelation.isValid ? String(parsedCorrelation.moduleName || '').trim() : coreMailResolveModule_(correlationKey)),
    moduleCode: String(routing.moduleCode || parsedCorrelation.moduleCode || '').trim(),
    entityType: String(routing.entityType || parsedCorrelation.entityType || '').trim(),
    entityId: String(routing.entityId || parsedCorrelation.entityId || parsedCorrelation.businessId || '').trim(),
    routingReason: String(routing.reason || '').trim(),
    routingConfidence: Number(routing.confidence || 0),
    processingStatus: shouldIgnore ? CORE_MAIL_HUB_STATUS.IGNORADO : CORE_MAIL_HUB_STATUS.PENDENTE,
    processorName: shouldIgnore ? 'coreMailIngestInbox' : '',
    processedAt: shouldIgnore ? (ingestedAt || new Date()) : '',
    hasAttachments: attachments.length ? 'SIM' : 'NAO',
    attachmentCount: attachments.length,
    snippet: msgCtx.snippet,
    plainBody: ingestConfig && ingestConfig.saveFullBody === true ? msgCtx.plainBody : '',
    labels: msgCtx.labels,
    attachments: attachments,
    parsedCorrelation: parsedCorrelation,
    isNoise: noise.isNoise,
    noiseReason: noise.reason,
    shouldIgnore: shouldIgnore
  };
}

function coreMailRegisterEvent_(eventosSheet, eventPayload, eventosState) {
  var ctx = eventosState.ctx;

  coreMailHubAppendRow_(eventosSheet, ctx, {
    'Id Evento': eventPayload.eventId,
    'Data Hora Evento': eventPayload.messageDate,
    'Direcao': eventPayload.direction,
    'Tipo Evento': eventPayload.eventType,
    'Modulo Dono': eventPayload.moduleName,
    'Tipo Entidade': eventPayload.entityType,
    'Id Entidade': eventPayload.entityId,
    'Chave de Correlacao': eventPayload.correlationKey,
    'Etapa Fluxo': eventPayload.flowStep,
    'Id Mensagem Gmail': eventPayload.messageId,
    'Id Thread Gmail': eventPayload.threadId,
    'Id Mensagem Pai': '',
    'Assunto': eventPayload.subject,
    'Email Remetente': eventPayload.fromEmail,
    'Nome Remetente': eventPayload.fromName,
    'Emails Destinatarios': eventPayload.to,
    'Emails Cc': eventPayload.cc,
    'Emails Cco': eventPayload.bcc,
    'Trecho Corpo': eventPayload.snippet,
    'Corpo Texto': eventPayload.plainBody || '',
    'Possui Anexos': eventPayload.hasAttachments,
    'Quantidade Anexos': eventPayload.attachmentCount,
    'Nomes Anexos': (eventPayload.attachments || []).map(function(item) {
      return String(item.getName() || '').trim();
    }).join(' | '),
    'Status Roteamento': eventPayload.routingStatus,
    'Status Processamento': eventPayload.processingStatus,
    'Processado Por': eventPayload.processorName,
    'Data Hora Processamento': eventPayload.processedAt,
    'Observacoes': eventPayload.noiseReason ? ('NOISE_REASON=' + eventPayload.noiseReason) : '',
    'Json Bruto': JSON.stringify({
      messageId: eventPayload.messageId,
      threadId: eventPayload.threadId,
      labels: eventPayload.labels,
      replyTo: eventPayload.replyTo,
      fromRaw: eventPayload.fromRaw,
      moduleCode: eventPayload.moduleCode,
      entityType: eventPayload.entityType,
      entityId: eventPayload.entityId,
      routingReason: eventPayload.routingReason,
      routingConfidence: eventPayload.routingConfidence
    }),
    'Criado Em': eventPayload.ingestedAt,
    'Atualizado Em': eventPayload.ingestedAt
  });

  var rowNumber = eventosSheet.getLastRow();
  eventosState.messageIds[eventPayload.messageId] = true;
  eventosState.rowByEventId[eventPayload.eventId] = rowNumber;
  return Object.freeze({
    eventId: eventPayload.eventId,
    rowNumber: rowNumber
  });
}

function coreMailRegisterAttachments_(anexosSheet, eventPayload, anexosCtx) {
  var attachments = eventPayload.attachments || [];
  var inserted = 0;

  for (var i = 0; i < attachments.length; i++) {
    var attachment = attachments[i];
    var bytesLength = 0;

    try {
      bytesLength = attachment.getBytes().length;
    } catch (err) {
      bytesLength = 0;
    }

    coreMailHubAppendRow_(anexosSheet, anexosCtx, {
      'Id Anexo': 'MAN-' + eventPayload.messageId + '-' + String(i + 1),
      'Id Evento': eventPayload.eventId,
      'Modulo Dono': eventPayload.moduleName,
      'Tipo Entidade': eventPayload.entityType,
      'Id Entidade': eventPayload.entityId,
      'Chave de Correlacao': eventPayload.correlationKey,
      'Etapa Fluxo': eventPayload.flowStep,
      'Id Mensagem Gmail': eventPayload.messageId,
      'Id Thread Gmail': eventPayload.threadId,
      'Nome Arquivo': String(attachment.getName() || '').trim(),
      'Tipo Mime': String(attachment.getContentType() || '').trim(),
      'Tamanho Bytes': bytesLength,
      'Foi Salvo No Drive': 'NAO',
      'Id Arquivo Drive': '',
      'Link Arquivo Drive': '',
      'Pasta Destino Drive': '',
      'Status Anexo': 'PENDENTE',
      'Processado Por': '',
      'Data Hora Processamento': '',
      'Observacoes': 'Attachment Index=' + String(i + 1),
      'Criado Em': eventPayload.ingestedAt,
      'Atualizado Em': eventPayload.ingestedAt
    });

    inserted++;
  }

  return inserted;
}

function coreMailHubCoerceDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (value === '' || value === null || typeof value === 'undefined') return null;

  var parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function coreMailHubExtractCorrelationPrefix_(correlationKey) {
  var normalized = String(correlationKey || '').trim().toUpperCase();
  if (!normalized) return '';
  var firstDash = normalized.indexOf('-');
  return firstDash > 0 ? normalized.substring(0, firstDash) : normalized;
}

function coreMailHubCollectAttachmentStatsByCorrelationKey_(correlationKey) {
  var normalizedKey = core_normalizeText_(correlationKey, {
    collapseWhitespace: true,
    caseMode: 'upper'
  });
  if (!normalizedKey) {
    return {
      totalAttachments: 0,
      hasPendingAttachment: false
    };
  }

  var anexosSheet = coreMailHubGetAnexosSheet_();
  var ctx = coreMailHubGetSheetContext_(anexosSheet, { includeRows: true });
  var stats = {
    totalAttachments: 0,
    hasPendingAttachment: false
  };

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var rowKey = core_normalizeText_(coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', ''), {
      collapseWhitespace: true,
      caseMode: 'upper'
    });
    if (rowKey !== normalizedKey) continue;

    stats.totalAttachments++;

    var statusAnexo = coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Status Anexo', 'PENDENTE'));
    if (statusAnexo !== CORE_MAIL_HUB_STATUS.PROCESSADO && statusAnexo !== CORE_MAIL_HUB_STATUS.IGNORADO) {
      stats.hasPendingAttachment = true;
    }
  }

  return stats;
}

function coreMailHubCollectCorrelationStats_(correlationKey) {
  var normalizedKey = core_normalizeText_(correlationKey, {
    collapseWhitespace: true,
    caseMode: 'upper'
  });
  if (!normalizedKey) return null;

  var eventosSheet = coreMailHubGetEventosSheet_();
  var ctx = coreMailHubGetSheetContext_(eventosSheet, { includeRows: true });
  var stats = {
    correlationKey: normalizedKey,
    correlationPrefix: coreMailHubExtractCorrelationPrefix_(normalizedKey),
    moduleName: '',
    entityType: '',
    entityId: '',
    currentStage: '',
    latestThreadId: '',
    latestMessageId: '',
    latestDirection: '',
    latestEventType: '',
    latestFromEmail: '',
    latestSubject: '',
    latestEventAt: '',
    latestReplyAt: '',
    totalEvents: 0,
    totalEntries: 0,
    totalOutputs: 0,
    totalAttachments: 0,
    hasPendingEntry: false,
    hasPendingAttachment: false,
    statusConversa: 'PENDENTE',
    createdAt: '',
    updatedAt: new Date(),
    allIgnored: true
  };
  var latestTime = 0;
  var parsedCorrelation = coreMailParseCorrelationKey_(normalizedKey);

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var rowKey = core_normalizeText_(coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', ''), {
      collapseWhitespace: true,
      caseMode: 'upper'
    });
    if (rowKey !== normalizedKey) continue;

    stats.totalEvents++;

    var direction = coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Direcao', ''));
    var processingStatus = coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Status Processamento', ''));
    var quantityAttachments = Number(coreMailHubGetRowValue_(row, ctx, 'Quantidade Anexos', 0) || 0);

    if (direction === 'SAIDA') {
      stats.totalOutputs++;
    } else {
      stats.totalEntries++;
    }

    stats.totalAttachments += quantityAttachments;

    if (direction !== 'SAIDA' && processingStatus === CORE_MAIL_HUB_STATUS.PENDENTE) {
      stats.hasPendingEntry = true;
    }
    if (processingStatus !== CORE_MAIL_HUB_STATUS.IGNORADO) {
      stats.allIgnored = false;
    }

    var createdAt = coreMailHubGetRowValue_(row, ctx, 'Criado Em', '');
    if (!stats.createdAt || (coreMailHubCoerceDate_(createdAt) && coreMailHubCoerceDate_(createdAt).getTime() < coreMailHubCoerceDate_(stats.createdAt).getTime())) {
      stats.createdAt = createdAt;
    }

    var eventAt = coreMailHubCoerceDate_(coreMailHubGetRowValue_(row, ctx, 'Data Hora Evento', '')) ||
      coreMailHubCoerceDate_(createdAt);
    var eventTime = eventAt ? eventAt.getTime() : 0;

    if (eventTime >= latestTime) {
      latestTime = eventTime;
      stats.moduleName = String(coreMailHubGetRowValue_(row, ctx, 'Modulo Dono', '') || '').trim();
      stats.entityType = String(coreMailHubGetRowValue_(row, ctx, 'Tipo Entidade', '') || '').trim();
      stats.entityId = String(coreMailHubGetRowValue_(row, ctx, 'Id Entidade', '') || '').trim();
      stats.currentStage = String(coreMailHubGetRowValue_(row, ctx, 'Etapa Fluxo', '') || '').trim();
      stats.latestThreadId = String(coreMailHubGetRowValue_(row, ctx, 'Id Thread Gmail', '') || '').trim();
      stats.latestMessageId = String(coreMailHubGetRowValue_(row, ctx, 'Id Mensagem Gmail', '') || '').trim();
      stats.latestDirection = direction;
      stats.latestEventType = String(coreMailHubGetRowValue_(row, ctx, 'Tipo Evento', '') || '').trim();
      stats.latestFromEmail = String(coreMailHubGetRowValue_(row, ctx, 'Email Remetente', '') || '').trim();
      stats.latestSubject = String(coreMailHubGetRowValue_(row, ctx, 'Assunto', '') || '').trim();
      stats.latestEventAt = coreMailHubGetRowValue_(row, ctx, 'Data Hora Evento', '');
      if (direction !== 'SAIDA') {
        stats.latestReplyAt = stats.latestEventAt;
      }
    }
  }

  if (!stats.totalEvents) return null;

  if (parsedCorrelation.isValid) {
    if (!stats.moduleName) stats.moduleName = String(parsedCorrelation.moduleName || '').trim();
    if (!stats.entityType) stats.entityType = String(parsedCorrelation.entityType || '').trim();
    if (!stats.entityId) stats.entityId = String(parsedCorrelation.entityId || parsedCorrelation.businessId || '').trim();
    if (!stats.currentStage) stats.currentStage = String(parsedCorrelation.stage || parsedCorrelation.flowCode || '').trim();
  }

  var attachmentStats = coreMailHubCollectAttachmentStatsByCorrelationKey_(normalizedKey);
  if (attachmentStats.totalAttachments > 0) {
    stats.totalAttachments = attachmentStats.totalAttachments;
  }
  stats.hasPendingAttachment = attachmentStats.hasPendingAttachment;

  if (stats.allIgnored) {
    stats.statusConversa = CORE_MAIL_HUB_STATUS.IGNORADO;
  } else if (stats.hasPendingEntry || stats.hasPendingAttachment) {
    stats.statusConversa = CORE_MAIL_HUB_STATUS.PENDENTE;
  } else {
    stats.statusConversa = CORE_MAIL_HUB_STATUS.PROCESSADO;
  }

  return stats;
}

function coreMailHubRefreshIndexSummaryByCorrelationKey_(correlationKey, indiceSheet, indiceState) {
  var stats = coreMailHubCollectCorrelationStats_(correlationKey);
  if (!stats) {
    return Object.freeze({
      updated: false,
      skipped: true,
      reason: 'SEM_EVENTOS'
    });
  }

  indiceSheet = indiceSheet || coreMailHubGetIndiceSheet_();
  indiceState = indiceState || coreMailHubBuildIndiceState_(indiceSheet);

  var rowNumber = indiceState.rowByCorrelationKey[stats.correlationKey];
  var ctx = indiceState.ctx;
  var inserted = false;

  if (!rowNumber) {
    var row = coreMailHubAppendRow_(indiceSheet, ctx, {
      'Chave de Correlacao': stats.correlationKey,
      'Criado Em': stats.createdAt || stats.updatedAt,
      'Atualizado Em': stats.updatedAt
    });
    rowNumber = indiceSheet.getLastRow();
    indiceState.rowByCorrelationKey[stats.correlationKey] = rowNumber;
    indiceState.recordByCorrelationKey[stats.correlationKey] = row;
    inserted = true;
  }

  var currentRow = indiceState.recordByCorrelationKey[stats.correlationKey];
  var payload = {
    'Chave de Correlacao': stats.correlationKey,
    'Prefixo Correlacao': stats.correlationPrefix,
    'Modulo Dono': stats.moduleName,
    'Tipo Entidade': stats.entityType,
    'Id Entidade': stats.entityId,
    'Etapa Atual': stats.currentStage,
    'Id Thread Gmail': stats.latestThreadId,
    'Id Ultima Mensagem': stats.latestMessageId,
    'Ultima Direcao': stats.latestDirection,
    'Ultimo Tipo Evento': stats.latestEventType,
    'Ultimo Email Remetente': stats.latestFromEmail,
    'Ultimo Assunto': stats.latestSubject,
    'Data Hora Ultimo Evento': stats.latestEventAt,
    'Ha Entrada Pendente': stats.hasPendingEntry ? 'SIM' : 'NAO',
    'Ha Anexo Pendente': stats.hasPendingAttachment ? 'SIM' : 'NAO',
    'Quantidade Eventos': stats.totalEvents,
    'Quantidade Entradas': stats.totalEntries,
    'Quantidade Saidas': stats.totalOutputs,
    'Quantidade Anexos': stats.totalAttachments,
    'Ultima Resposta Em': stats.latestReplyAt,
    'Status Conversa': stats.statusConversa,
    'Atualizado Em': stats.updatedAt
  };

  Object.keys(payload).forEach(function(headerName) {
    coreMailHubWriteCell_(indiceSheet, rowNumber, ctx, headerName, payload[headerName]);
    coreMailHubSetRowValue_(currentRow, ctx, headerName, payload[headerName]);
  });

  if (inserted) {
    coreMailHubWriteCell_(indiceSheet, rowNumber, ctx, 'Criado Em', stats.createdAt || stats.updatedAt);
    coreMailHubSetRowValue_(currentRow, ctx, 'Criado Em', stats.createdAt || stats.updatedAt);
  }

  return Object.freeze({
    updated: true,
    inserted: inserted,
    rowNumber: rowNumber
  });
}

function coreMailUpsertIndex_(indiceSheet, eventPayload, indiceState) {
  var correlationKey = String(eventPayload.correlationKey || '').trim().toUpperCase();
  if (!correlationKey) {
    return Object.freeze({
      updated: false,
      skipped: true,
      reason: 'SEM_CHAVE_CORRELACAO'
    });
  }

  return coreMailHubRefreshIndexSummaryByCorrelationKey_(correlationKey, indiceSheet, indiceState);
}

function core_mailIngestInbox_(opts) {
  opts = opts || {};

  return core_withLock_('CORE_MAIL_HUB_INGEST_INBOX', function() {
    var runId = core_runId_();
    var startedAt = new Date();
    var configMap = coreMailHubGetConfigMap_();
    var query = String(
      Object.prototype.hasOwnProperty.call(opts, 'query')
        ? opts.query
        : coreMailHubGetConfigValue_(configMap, 'GMAIL_QUERY_INGEST', CORE_MAIL_HUB_DEFAULTS.query)
    );
    var start = Number(
      Object.prototype.hasOwnProperty.call(opts, 'start')
        ? opts.start
        : coreMailHubGetConfigNumber_(configMap, 'GMAIL_START', CORE_MAIL_HUB_DEFAULTS.start)
    );
    var maxThreads = Number(
      Object.prototype.hasOwnProperty.call(opts, 'maxThreads')
        ? opts.maxThreads
        : coreMailHubGetConfigNumber_(configMap, 'GMAIL_MAX_THREADS', CORE_MAIL_HUB_DEFAULTS.maxThreads)
    );
    var maxMessagesPerThread = Number(
      Object.prototype.hasOwnProperty.call(opts, 'maxMessagesPerThread')
        ? opts.maxMessagesPerThread
        : coreMailHubGetConfigNumber_(configMap, 'GMAIL_MAX_MESSAGES_PER_THREAD', CORE_MAIL_HUB_DEFAULTS.maxMessagesPerThread)
    );
    var dryRun = opts.dryRun === true;
    var ingestConfig = coreMailHubBuildIngestConfig_(configMap, opts);

    query = String(query || '').trim() || CORE_MAIL_HUB_DEFAULTS.query;
    start = isNaN(start) || start < 0 ? CORE_MAIL_HUB_DEFAULTS.start : start;
    maxThreads = isNaN(maxThreads) || maxThreads < 1 ? CORE_MAIL_HUB_DEFAULTS.maxThreads : maxThreads;
    maxMessagesPerThread =
      isNaN(maxMessagesPerThread) || maxMessagesPerThread < 1
        ? CORE_MAIL_HUB_DEFAULTS.maxMessagesPerThread
        : maxMessagesPerThread;

    coreMailHubAssertSchema_();

    var eventosSheet = coreMailHubGetEventosSheet_();
    var indiceSheet = coreMailHubGetIndiceSheet_();
    var anexosSheet = coreMailHubGetAnexosSheet_();

    var eventosState = coreMailHubBuildEventosState_(eventosSheet);
    var indiceState = coreMailHubBuildIndiceState_(indiceSheet);
    var anexosCtx = coreMailHubGetSheetContext_(anexosSheet);

    var threads = GmailApp.search(query, start, maxThreads);
    var counters = {
      threadsScanned: threads.length,
      messagesScanned: 0,
      newEvents: 0,
      duplicatesSkipped: 0,
      noiseSkipped: 0,
      ignoredNoiseEvents: 0,
      attachmentsRegistered: 0,
      indexUpserts: 0,
      withoutCorrelationKey: 0,
      executionLimitReached: false,
      errors: 0
    };

    core_logInfo_(runId, 'Mail Hub ingest iniciado', {
      query: query,
      start: start,
      maxThreads: maxThreads,
      maxMessagesPerThread: maxMessagesPerThread,
      dryRun: dryRun,
      ingestConfig: {
        requiredSubjectPrefix: ingestConfig.requiredSubjectPrefix,
        useOnlyTaggedSubjects: ingestConfig.useOnlyTaggedSubjects,
        ignoredSendersCount: ingestConfig.ignoredSenders.length,
        ignoredDomainsCount: ingestConfig.ignoredDomains.length,
        ignoredSubjectRegexCount: ingestConfig.ignoredSubjectRegexes.length,
        maxEventsPerExecution: ingestConfig.maxEventsPerExecution,
        saveFullBody: ingestConfig.saveFullBody,
        markNoiseAsIgnored: ingestConfig.markNoiseAsIgnored
      }
    });

    outerLoop:
    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var messages = thread.getMessages();
      var startIndex = Math.max(0, messages.length - maxMessagesPerThread);

      for (var j = startIndex; j < messages.length; j++) {
        if (ingestConfig.maxEventsPerExecution > 0 && counters.newEvents >= ingestConfig.maxEventsPerExecution) {
          counters.executionLimitReached = true;
          break outerLoop;
        }

        var message = messages[j];
        counters.messagesScanned++;

        try {
          var eventPayload = coreMailHubBuildEventPayload_(thread, message, new Date(), ingestConfig);
          if (!eventPayload.messageId) {
            counters.errors++;
            core_logWarn_(runId, 'Mail Hub pulou mensagem sem messageId', {
              threadId: String(thread.getId() || '').trim(),
              subject: String(message.getSubject() || '').trim()
            });
            continue;
          }

          if (coreMailEventExistsByMessageId_(eventPayload.messageId, eventosState)) {
            counters.duplicatesSkipped++;
            continue;
          }

          if (eventPayload.isNoise && !eventPayload.shouldIgnore) {
            counters.noiseSkipped++;
            continue;
          }

          if (!eventPayload.correlationKey) {
            counters.withoutCorrelationKey++;
          }

          if (dryRun) {
            counters.newEvents++;
            if (eventPayload.shouldIgnore) {
              counters.ignoredNoiseEvents++;
              if (eventPayload.correlationKey) counters.indexUpserts++;
            } else {
              counters.attachmentsRegistered += eventPayload.attachmentCount;
              if (eventPayload.correlationKey) counters.indexUpserts++;
            }
            continue;
          }

          var registeredEvent = coreMailRegisterEvent_(eventosSheet, eventPayload, eventosState);
          eventPayload.eventId = registeredEvent.eventId;

          counters.newEvents++;
          if (eventPayload.shouldIgnore) {
            counters.ignoredNoiseEvents++;
            if (eventPayload.correlationKey) {
              var ignoredUpsertInfo = coreMailUpsertIndex_(indiceSheet, eventPayload, indiceState);
              if (ignoredUpsertInfo.updated) counters.indexUpserts++;
            }
          } else {
            counters.attachmentsRegistered += coreMailRegisterAttachments_(anexosSheet, eventPayload, anexosCtx);

            var upsertInfo = coreMailUpsertIndex_(indiceSheet, eventPayload, indiceState);
            if (upsertInfo.updated) counters.indexUpserts++;
          }
        } catch (err) {
          counters.errors++;
          core_logError_(runId, 'Mail Hub erro ao ingerir mensagem', {
            error: err && err.message ? err.message : String(err || ''),
            threadId: String(thread.getId() || '').trim(),
            messageIndex: j
          });
        }
      }
    }

    core_logSummarize_(runId, 'core_mailIngestInbox_', startedAt, counters);

    return Object.freeze({
      ok: true,
      dryRun: dryRun,
      query: query,
      start: start,
      maxThreads: maxThreads,
      maxMessagesPerThread: maxMessagesPerThread,
      ingestConfig: Object.freeze({
        requiredSubjectPrefix: ingestConfig.requiredSubjectPrefix,
        useOnlyTaggedSubjects: ingestConfig.useOnlyTaggedSubjects,
        maxEventsPerExecution: ingestConfig.maxEventsPerExecution,
        saveFullBody: ingestConfig.saveFullBody,
        markNoiseAsIgnored: ingestConfig.markNoiseAsIgnored
      }),
      counters: Object.freeze(counters)
    });
  });
}

function core_mailListPendingByModule_(moduleName) {
  core_assertRequired_(moduleName, 'moduleName');
  return coreMailHubListEvents_({
    moduleName: moduleName,
    processingStatus: CORE_MAIL_HUB_STATUS.PENDENTE
  });
}

function coreMailHubNormalizeOptionalFilter_(value) {
  if (value === '' || value === null || typeof value === 'undefined') return '';
  return core_normalizeText_(value, {
    collapseWhitespace: true,
    caseMode: 'upper'
  });
}

function coreMailHubBuildEventRecord_(row, ctx, rowNumber) {
  return {
    eventId: String(coreMailHubGetRowValue_(row, ctx, 'Id Evento', '') || '').trim(),
    messageId: String(coreMailHubGetRowValue_(row, ctx, 'Id Mensagem Gmail', '') || '').trim(),
    threadId: String(coreMailHubGetRowValue_(row, ctx, 'Id Thread Gmail', '') || '').trim(),
    correlationKey: String(coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', '') || '').trim(),
    direction: String(coreMailHubGetRowValue_(row, ctx, 'Direcao', '') || '').trim(),
    moduleName: coreMailHubNormalizeOptionalFilter_(coreMailHubGetRowValue_(row, ctx, 'Modulo Dono', '')),
    entityType: String(coreMailHubGetRowValue_(row, ctx, 'Tipo Entidade', '') || '').trim(),
    entityId: String(coreMailHubGetRowValue_(row, ctx, 'Id Entidade', '') || '').trim(),
    flowStep: String(coreMailHubGetRowValue_(row, ctx, 'Etapa Fluxo', '') || '').trim(),
    eventType: String(coreMailHubGetRowValue_(row, ctx, 'Tipo Evento', '') || '').trim(),
    subject: String(coreMailHubGetRowValue_(row, ctx, 'Assunto', '') || '').trim(),
    fromEmail: String(coreMailHubGetRowValue_(row, ctx, 'Email Remetente', '') || '').trim(),
    fromName: String(coreMailHubGetRowValue_(row, ctx, 'Nome Remetente', '') || '').trim(),
    processingStatus: coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Status Processamento', '')),
    routingStatus: coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Status Roteamento', '')),
    receivedAt: coreMailHubGetRowValue_(row, ctx, 'Data Hora Evento', ''),
    ingestedAt: coreMailHubGetRowValue_(row, ctx, 'Criado Em', ''),
    updatedAt: coreMailHubGetRowValue_(row, ctx, 'Atualizado Em', ''),
    hasAttachments: String(coreMailHubGetRowValue_(row, ctx, 'Possui Anexos', '') || '').trim(),
    attachmentCount: Number(coreMailHubGetRowValue_(row, ctx, 'Quantidade Anexos', 0) || 0),
    processedBy: String(coreMailHubGetRowValue_(row, ctx, 'Processado Por', '') || '').trim(),
    processedAt: coreMailHubGetRowValue_(row, ctx, 'Data Hora Processamento', ''),
    rowNumber: rowNumber
  };
}

function coreMailHubListEvents_(opts) {
  opts = opts || {};
  coreMailHubAssertSchema_();

  var sheet = coreMailHubGetEventosSheet_();
  var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
  var out = [];
  var targetModule = coreMailHubNormalizeOptionalFilter_(opts.moduleName);
  var targetStatus = coreMailHubNormalizeOptionalFilter_(opts.processingStatus);
  var targetCorrelationKey = coreMailHubNormalizeOptionalFilter_(opts.correlationKey);
  var targetMessageId = String(opts.messageId || '').trim();
  var targetThreadId = String(opts.threadId || '').trim();
  var limit = Number(opts.limit || 0);

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var record = coreMailHubBuildEventRecord_(row, ctx, ctx.startRow + i);

    if (targetModule && record.moduleName !== targetModule) continue;
    if (targetStatus && record.processingStatus !== targetStatus) continue;
    if (targetCorrelationKey && coreMailHubNormalizeOptionalFilter_(record.correlationKey) !== targetCorrelationKey) continue;
    if (targetMessageId && record.messageId !== targetMessageId) continue;
    if (targetThreadId && record.threadId !== targetThreadId) continue;

    out.push(record);
  }

  out.sort(function(a, b) {
    var aDate = coreMailHubCoerceDate_(a.receivedAt) || coreMailHubCoerceDate_(a.ingestedAt);
    var bDate = coreMailHubCoerceDate_(b.receivedAt) || coreMailHubCoerceDate_(b.ingestedAt);
    var aTime = aDate ? aDate.getTime() : 0;
    var bTime = bDate ? bDate.getTime() : 0;
    return bTime - aTime;
  });

  if (limit > 0) {
    return out.slice(0, limit);
  }

  return out;
}

function core_mailGetLatestEvent_(opts) {
  var events = coreMailHubListEvents_(Object.assign({}, opts || {}, { limit: 1 }));
  return events.length ? events[0] : null;
}

function core_mailMarkLatestPendingByModule_(moduleName, processorName) {
  core_assertRequired_(moduleName, 'moduleName');
  core_assertRequired_(processorName, 'processorName');

  var latestPending = core_mailGetLatestEvent_({
    moduleName: moduleName,
    processingStatus: CORE_MAIL_HUB_STATUS.PENDENTE
  });

  if (!latestPending) {
    return Object.freeze({
      ok: true,
      found: false,
      moduleName: coreMailHubNormalizeOptionalFilter_(moduleName),
      processorName: String(processorName || '').trim()
    });
  }

  var result = core_mailMarkEventProcessed_(latestPending.eventId, processorName);
  return Object.freeze({
    ok: true,
    found: true,
    latestEvent: latestPending,
    result: result
  });
}

function core_mailMarkEventProcessed_(eventId, processorName) {
  core_assertRequired_(eventId, 'eventId');
  core_assertRequired_(processorName, 'processorName');

  return core_withLock_('CORE_MAIL_HUB_MARK_EVENT_PROCESSED', function() {
    coreMailHubAssertSchema_();

    var sheet = coreMailHubGetEventosSheet_();
    var eventosState = coreMailHubBuildEventosState_(sheet);
    var normalizedEventId = String(eventId || '').trim();
    var rowNumber = eventosState.rowByEventId[normalizedEventId];

    if (!rowNumber) {
      throw new Error('Evento nao encontrado em MAIL_EVENTOS: ' + eventId);
    }

    var processedAt = new Date();
    coreMailHubWriteCell_(sheet, rowNumber, eventosState.ctx, 'Status Processamento', CORE_MAIL_HUB_STATUS.PROCESSADO);
    coreMailHubWriteCell_(sheet, rowNumber, eventosState.ctx, 'Processado Por', String(processorName || '').trim());
    coreMailHubWriteCell_(sheet, rowNumber, eventosState.ctx, 'Data Hora Processamento', processedAt);
    coreMailHubWriteCell_(sheet, rowNumber, eventosState.ctx, 'Atualizado Em', processedAt);

    var row = sheet.getRange(rowNumber, 1, 1, eventosState.ctx.lastCol).getValues()[0];
    var correlationKey = String(coreMailHubGetRowValue_(row, eventosState.ctx, 'Chave de Correlacao', '') || '').trim();
    if (correlationKey) {
      coreMailHubRefreshIndexSummaryByCorrelationKey_(correlationKey);
    }

    return Object.freeze({
      ok: true,
      eventId: normalizedEventId,
      processorName: String(processorName || '').trim(),
      status: CORE_MAIL_HUB_STATUS.PROCESSADO,
      processedAt: processedAt,
      rowNumber: rowNumber
    });
  });
}

function coreMailCleanupNoiseEvents_() {
  return core_withLock_('CORE_MAIL_HUB_CLEANUP_NOISE_EVENTS', function() {
    coreMailHubAssertSchema_();

    var configMap = coreMailHubGetConfigMap_();
    var ingestConfig = coreMailHubBuildIngestConfig_(configMap, {
      markNoiseAsIgnored: true
    });
    var sheet = coreMailHubGetEventosSheet_();
    var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
    var touchedCorrelationKeys = Object.create(null);
    var counters = {
      scanned: 0,
      updated: 0,
      alreadyIgnored: 0,
      skipped: 0
    };
    var processedAt = new Date();

    for (var i = 0; i < ctx.rows.length; i++) {
      var row = ctx.rows[i];
      var rowNumber = ctx.startRow + i;
      counters.scanned++;

      var msgCtx = {
        subject: String(coreMailHubGetRowValue_(row, ctx, 'Assunto', '') || '').trim(),
        fromEmail: String(coreMailHubGetRowValue_(row, ctx, 'Email Remetente', '') || '').trim(),
        fromRaw: String(coreMailHubGetRowValue_(row, ctx, 'Email Remetente', '') || '').trim(),
        snippet: String(coreMailHubGetRowValue_(row, ctx, 'Trecho Corpo', '') || '').trim(),
        plainBody: String(coreMailHubGetRowValue_(row, ctx, 'Corpo Texto', '') || '').trim()
      };
      var noise = coreMailHubDetectNoise_(msgCtx, ingestConfig);
      if (!noise.isNoise) {
        counters.skipped++;
        continue;
      }

      var currentStatus = coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Status Processamento', ''));
      if (currentStatus === CORE_MAIL_HUB_STATUS.IGNORADO) {
        counters.alreadyIgnored++;
        continue;
      }

      var currentObservacoes = String(coreMailHubGetRowValue_(row, ctx, 'Observacoes', '') || '').trim();
      var noiseNote = 'NOISE_REASON=' + noise.reason;
      var nextObservacoes = currentObservacoes
        ? (currentObservacoes.indexOf(noiseNote) >= 0 ? currentObservacoes : currentObservacoes + ' | ' + noiseNote)
        : noiseNote;

      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Status Roteamento', CORE_MAIL_HUB_STATUS.IGNORADO);
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Status Processamento', CORE_MAIL_HUB_STATUS.IGNORADO);
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Processado Por', 'coreMailCleanupNoiseEvents');
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Data Hora Processamento', processedAt);
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Observacoes', nextObservacoes);
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Atualizado Em', processedAt);
      counters.updated++;

      var correlationKey = String(coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', '') || '').trim();
      if (correlationKey) {
        touchedCorrelationKeys[core_normalizeText_(correlationKey, {
          collapseWhitespace: true,
          caseMode: 'upper'
        })] = true;
      }
    }

    Object.keys(touchedCorrelationKeys).forEach(function(correlationKey) {
      coreMailHubRefreshIndexSummaryByCorrelationKey_(correlationKey);
    });

    return Object.freeze({
      ok: true,
      processedAt: processedAt,
      counters: Object.freeze(counters),
      refreshedIndexKeys: Object.keys(touchedCorrelationKeys)
    });
  });
}
