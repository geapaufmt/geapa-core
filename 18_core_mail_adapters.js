/**
 * ============================================================
 * 18_core_mail_adapters.js
 * ============================================================
 *
 * Registry e contrato de adapters de mensageria por modulo.
 *
 * Objetivo:
 * - permitir que o core conheca apenas um contrato minimo comum;
 * - delegar correlacao, parsing e roteamento aos adapters;
 * - manter o Mail Hub desacoplado das regras de negocio dos modulos.
 *
 * Observacao importante:
 * - em Apps Script, callbacks/funcoes entre projeto consumidor e Library
 *   podem ter limitacoes. Esta camada e totalmente segura para adapters
 *   declarados no proprio core ou no mesmo projeto.
 */

var __core_mail_adapter_registry = null;
var __core_mail_default_adapters_bootstrapped = false;

function coreMailEnsureAdapterRegistry_() {
  if (__core_mail_adapter_registry) return __core_mail_adapter_registry;

  __core_mail_adapter_registry = {
    byCode: Object.create(null),
    byName: Object.create(null),
    list: []
  };

  return __core_mail_adapter_registry;
}

function coreMailNormalizeAdapterId_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  });
}

function coreMailNormalizeCorrelationToken_(value) {
  var text = coreMailNormalizeAdapterId_(value);
  return text.replace(/[^A-Z0-9_-]+/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '');
}

function coreMailStripWrappedCorrelationKey_(value) {
  var text = String(value || '').trim();
  if (!text) return '';

  var match = text.match(/\[GEAPA\]\[([^\]]+)\]/i);
  return match ? String(match[1] || '').trim() : text;
}

function coreMailAdapterToSnapshot_(adapter) {
  if (!adapter) return null;

  return Object.freeze({
    moduleName: adapter.moduleName,
    moduleCode: adapter.moduleCode,
    hasNormalizeOutgoingSubject: typeof adapter.normalizeOutgoingSubject === 'function',
    hasSummarizeForHistory: typeof adapter.summarizeForHistory === 'function'
  });
}

function coreMailAssertAdapterContract_(adapter) {
  core_assertRequired_(adapter, 'adapter');

  var moduleName = String(adapter.moduleName || '').trim();
  var moduleCode = coreMailNormalizeAdapterId_(adapter.moduleCode);

  if (!moduleName) {
    throw new Error('Mail adapter invalido: moduleName obrigatorio.');
  }

  if (!moduleCode) {
    throw new Error('Mail adapter invalido: moduleCode obrigatorio.');
  }

  var requiredFns = [
    'buildCorrelationKey',
    'parseCorrelationKey',
    'matchMessage',
    'resolveRouting'
  ];

  requiredFns.forEach(function(fnName) {
    if (typeof adapter[fnName] !== 'function') {
      throw new Error('Mail adapter invalido (' + moduleCode + '): funcao obrigatoria ausente: ' + fnName);
    }
  });

  if (adapter.normalizeOutgoingSubject != null && typeof adapter.normalizeOutgoingSubject !== 'function') {
    throw new Error('Mail adapter invalido (' + moduleCode + '): normalizeOutgoingSubject deve ser funcao.');
  }

  if (adapter.summarizeForHistory != null && typeof adapter.summarizeForHistory !== 'function') {
    throw new Error('Mail adapter invalido (' + moduleCode + '): summarizeForHistory deve ser funcao.');
  }

  return Object.freeze({
    moduleName: moduleName,
    moduleCode: moduleCode
  });
}

function coreMailEnsureDefaultAdaptersRegistered_() {
  if (__core_mail_default_adapters_bootstrapped) return;

  var registry = coreMailEnsureAdapterRegistry_();
  function registerIfMissing(adapter) {
    var codeKey = coreMailNormalizeAdapterId_(adapter.moduleCode);
    if (registry.byCode[codeKey]) return;
    coreMailRegisterModuleAdapter_(adapter);
  }

  registerIfMissing(coreMailCreateExampleModuleAdapter_({
    moduleName: 'ATIVIDADES',
    moduleCode: 'ATV',
    entityTypeDefault: 'ATIVIDADE',
    keywordHints: [
      'atividade',
      'atividades',
      'presenca',
      'presencas',
      'falta',
      'faltas',
      'justificativa',
      'justificativas',
      'ata',
      'material',
      'convocacao',
      'lembrete'
    ]
  }));

  registerIfMissing(coreMailCreateExampleModuleAdapter_({
    moduleName: 'APRESENTACOES',
    moduleCode: 'APR',
    entityTypeDefault: 'APRESENTACAO',
    keywordHints: ['apresentacao', 'apresentacoes', 'seminario', 'palestra']
  }));

  registerIfMissing(coreMailCreateExampleModuleAdapter_({
    moduleName: 'SELETIVO',
    moduleCode: 'SEL',
    entityTypeDefault: 'PROCESSO_SELETIVO',
    keywordHints: ['seletivo', 'processo seletivo', 'inscricao', 'candidato']
  }));

  registerIfMissing(coreMailCreateExampleModuleAdapter_({
    moduleName: 'MEMBROS',
    moduleCode: 'MEM',
    entityTypeDefault: 'MEMBRO',
    keywordHints: ['membro', 'membros', 'ingressar no geapa', 'aceito', 'recuso']
  }));

  __core_mail_default_adapters_bootstrapped = true;
}

function coreMailRegisterModuleAdapter_(adapter) {
  var contract = coreMailAssertAdapterContract_(adapter);
  var registry = coreMailEnsureAdapterRegistry_();
  var codeKey = contract.moduleCode;
  var nameKey = coreMailNormalizeAdapterId_(contract.moduleName);
  var existing = registry.byCode[codeKey];
  var existingByName = registry.byName[nameKey];

  if (existing) {
    var existingNameKey = coreMailNormalizeAdapterId_(existing.moduleName);
    delete registry.byName[existingNameKey];

    for (var i = registry.list.length - 1; i >= 0; i--) {
      if (registry.list[i].moduleCode === codeKey) {
        registry.list.splice(i, 1);
      }
    }
  }

  if (existingByName && existingByName.moduleCode !== codeKey) {
    delete registry.byCode[existingByName.moduleCode];

    for (var j = registry.list.length - 1; j >= 0; j--) {
      if (registry.list[j].moduleCode === existingByName.moduleCode) {
        registry.list.splice(j, 1);
      }
    }
  }

  adapter.moduleName = contract.moduleName;
  adapter.moduleCode = contract.moduleCode;

  registry.byCode[codeKey] = adapter;
  registry.byName[nameKey] = adapter;
  registry.list.push(adapter);

  return coreMailAdapterToSnapshot_(adapter);
}

function coreMailGetModuleAdapter_(moduleCodeOrName) {
  coreMailEnsureDefaultAdaptersRegistered_();
  if (moduleCodeOrName === '' || moduleCodeOrName === null || typeof moduleCodeOrName === 'undefined') {
    return null;
  }

  var key = coreMailNormalizeAdapterId_(moduleCodeOrName);
  var registry = coreMailEnsureAdapterRegistry_();
  return registry.byCode[key] || registry.byName[key] || null;
}

function coreMailListModuleAdapters_() {
  coreMailEnsureDefaultAdaptersRegistered_();
  return coreMailEnsureAdapterRegistry_().list.slice();
}

function coreMailBuildCorrelationKey_(moduleCodeOrName, ctx) {
  var adapter = coreMailGetModuleAdapter_(moduleCodeOrName);
  if (!adapter) {
    throw new Error('Mail adapter nao encontrado: ' + moduleCodeOrName);
  }

  var key = String(adapter.buildCorrelationKey(ctx || {}) || '').trim();
  if (!key) {
    throw new Error('Mail adapter "' + adapter.moduleCode + '" retornou correlationKey vazio.');
  }

  return key;
}

function coreMailParseCorrelationKey_(key) {
  coreMailEnsureDefaultAdaptersRegistered_();

  var rawKey = coreMailStripWrappedCorrelationKey_(key);
  if (!rawKey) {
    return Object.freeze({
      isValid: false
    });
  }

  var adapters = coreMailListModuleAdapters_();
  for (var i = 0; i < adapters.length; i++) {
    var adapter = adapters[i];
    var parsed = adapter.parseCorrelationKey(rawKey);
    if (!parsed || parsed.isValid !== true) continue;

    return Object.freeze({
      isValid: true,
      moduleCode: String(parsed.moduleCode || adapter.moduleCode || '').trim(),
      moduleName: String(parsed.moduleName || adapter.moduleName || '').trim(),
      businessId: parsed.businessId || '',
      flowCode: parsed.flowCode || '',
      entityType: parsed.entityType || '',
      entityId: parsed.entityId || '',
      stage: parsed.stage || '',
      counterpartType: parsed.counterpartType || '',
      targetKey: parsed.targetKey || rawKey
    });
  }

  return Object.freeze({
    isValid: false
  });
}

function coreMailResolveRouting_(msgCtx) {
  coreMailEnsureDefaultAdaptersRegistered_();
  msgCtx = msgCtx || {};

  var ruleRouting = coreMailResolveRoutingByRules_(msgCtx);
  if (ruleRouting && ruleRouting.matched) {
    return ruleRouting;
  }

  var subjectKey = coreMailExtractCorrelationKey_(msgCtx.subject || '');
  if (subjectKey) {
    var parsed = coreMailParseCorrelationKey_(subjectKey);
    if (parsed.isValid) {
      return Object.freeze({
        matched: true,
        moduleName: parsed.moduleName || '',
        moduleCode: parsed.moduleCode || '',
        correlationKey: subjectKey,
        entityType: parsed.entityType || '',
        entityId: parsed.entityId || parsed.businessId || '',
        stage: parsed.stage || '',
        flowCode: parsed.flowCode || '',
        confidence: 1,
        reason: 'CORRELATION_KEY'
      });
    }
  }

  var adapters = coreMailListModuleAdapters_();
  var best = null;

  for (var i = 0; i < adapters.length; i++) {
    var adapter = adapters[i];
    var route = adapter.resolveRouting(msgCtx) || { matched: false, confidence: 0 };
    var match = adapter.matchMessage(msgCtx) || { matched: false, confidence: 0 };
    var candidate = null;

    if (route.matched) {
      candidate = Object.assign({}, route, {
        moduleName: route.moduleName || adapter.moduleName,
        moduleCode: route.moduleCode || adapter.moduleCode,
        confidence: Number(route.confidence || 0)
      });
    } else if (match.matched) {
      candidate = {
        matched: true,
        moduleName: adapter.moduleName,
        moduleCode: adapter.moduleCode,
        correlationKey: '',
        entityType: '',
        entityId: '',
        stage: '',
        flowCode: '',
        confidence: Number(match.confidence || 0),
        reason: match.reason || 'MATCH_MESSAGE'
      };
    }

    if (!candidate || candidate.matched !== true) continue;
    if (!best || Number(candidate.confidence || 0) > Number(best.confidence || 0)) {
      best = candidate;
    }
  }

  if (!best) {
    return Object.freeze({
      matched: false,
      moduleName: '',
      moduleCode: '',
      correlationKey: '',
      entityType: '',
      entityId: '',
      stage: '',
      flowCode: '',
      confidence: 0,
      reason: 'NO_ADAPTER_MATCH'
    });
  }

  return Object.freeze({
    matched: true,
    moduleName: best.moduleName || '',
    moduleCode: best.moduleCode || '',
    correlationKey: String(best.correlationKey || '').trim(),
    entityType: best.entityType || '',
    entityId: best.entityId || '',
    stage: best.stage || '',
    flowCode: best.flowCode || '',
    confidence: Number(best.confidence || 0),
    reason: best.reason || ''
  });
}

function coreMailNormalizeOutgoingSubject_(moduleCodeOrName, subject, ctx) {
  var adapter = coreMailGetModuleAdapter_(moduleCodeOrName);
  if (!adapter) {
    throw new Error('Mail adapter nao encontrado: ' + moduleCodeOrName);
  }

  var baseSubject = String(subject || '').trim();
  if (typeof adapter.normalizeOutgoingSubject === 'function') {
    return String(adapter.normalizeOutgoingSubject(baseSubject, ctx || {}) || '').trim();
  }

  var localCtx = ctx || {};
  var correlationKey = String(localCtx.correlationKey || '').trim();
  if (!correlationKey) {
    correlationKey = coreMailBuildCorrelationKey_(moduleCodeOrName, localCtx);
  }

  var tag = '[GEAPA][' + correlationKey + ']';
  if (!baseSubject) return tag;

  if (/\[GEAPA\]\[[^\]]+\]/i.test(baseSubject)) {
    return baseSubject.replace(/\[GEAPA\]\[[^\]]+\]/i, tag).trim();
  }

  return (tag + ' ' + baseSubject).trim();
}

function coreMailBuildMessageTextForMatching_(msgCtx) {
  return [
    String(msgCtx.subject || ''),
    String(msgCtx.fromEmail || ''),
    String(msgCtx.to || ''),
    String(msgCtx.cc || ''),
    String(msgCtx.bcc || ''),
    String(msgCtx.snippet || ''),
    String(msgCtx.plainBody || '')
  ].join(' ');
}

function coreMailNormalizeRoutingRuleText_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'lower'
  });
}

function coreMailNormalizeRoutingRuleToken_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  }).replace(/\s+/g, '_');
}

function coreMailReadRoutingRules_() {
  var sheet = coreMailHubGetRegrasSheet_();
  var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
  var rules = [];

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var active = coreMailNormalizeRoutingRuleToken_(coreMailHubGetRowValue_(row, ctx, 'Ativa', ''));
    var id = String(coreMailHubGetRowValue_(row, ctx, 'Id Regra', '') || '').trim();
    var field = coreMailNormalizeRoutingRuleToken_(coreMailHubGetRowValue_(row, ctx, 'Campo Analise', ''));
    var comparison = coreMailNormalizeRoutingRuleToken_(coreMailHubGetRowValue_(row, ctx, 'Tipo Comparacao', ''));
    var value = String(coreMailHubGetRowValue_(row, ctx, 'Valor Comparacao', '') || '').trim();
    var moduleName = coreMailNormalizeRoutingRuleToken_(coreMailHubGetRowValue_(row, ctx, 'Modulo Dono', ''));
    var action = coreMailNormalizeRoutingRuleToken_(coreMailHubGetRowValue_(row, ctx, 'Acao Quando Bater', 'ROTEAR'));

    if (active !== 'SIM') continue;
    if (!field || !comparison || !value || !moduleName) continue;
    if (action && action !== 'ROTEAR') continue;

    rules.push(Object.freeze({
      id: id || ('REG_ROW_' + String(i + 2)),
      rowNumber: i + 2,
      order: Number(coreMailHubGetRowValue_(row, ctx, 'Ordem', 999999) || 999999),
      field: field,
      comparison: comparison,
      value: value,
      moduleName: moduleName,
      entityType: coreMailNormalizeRoutingRuleToken_(coreMailHubGetRowValue_(row, ctx, 'Tipo Entidade', '')),
      flowStep: coreMailNormalizeRoutingRuleToken_(coreMailHubGetRowValue_(row, ctx, 'Etapa Fluxo', '')),
      action: action || 'ROTEAR',
      observations: String(coreMailHubGetRowValue_(row, ctx, 'Observacoes', '') || '').trim()
    }));
  }

  return rules.sort(function(a, b) {
    if (a.order !== b.order) return a.order - b.order;
    return a.rowNumber - b.rowNumber;
  });
}

function coreMailGetRoutingRuleFieldValue_(msgCtx, field) {
  msgCtx = msgCtx || {};
  var normalizedField = coreMailNormalizeRoutingRuleToken_(field);

  if (normalizedField === 'ASSUNTO') return String(msgCtx.subject || '');
  if (normalizedField === 'REMETENTE') return [msgCtx.fromRaw, msgCtx.fromEmail, msgCtx.fromName].join(' ');
  if (normalizedField === 'DESTINATARIO') return [msgCtx.to, msgCtx.cc, msgCtx.bcc].join(' ');
  if (normalizedField === 'CORPO') return [msgCtx.snippet, msgCtx.plainBody].join(' ');
  if (normalizedField === 'TUDO') return coreMailBuildMessageTextForMatching_(msgCtx);

  return '';
}

function coreMailRoutingRuleMatches_(rule, msgCtx) {
  var rawHaystack = coreMailGetRoutingRuleFieldValue_(msgCtx, rule.field);
  var rawNeedle = String(rule.value || '');
  var comparison = coreMailNormalizeRoutingRuleToken_(rule.comparison || 'CONTEM');

  if (!rawHaystack || !rawNeedle) return false;

  if (comparison === 'REGEX') {
    try {
      return new RegExp(rawNeedle, 'i').test(rawHaystack);
    } catch (err) {
      return false;
    }
  }

  var haystack = coreMailNormalizeRoutingRuleText_(rawHaystack);
  var needle = coreMailNormalizeRoutingRuleText_(rawNeedle);

  if (comparison === 'IGUAL') return haystack === needle;
  if (comparison === 'COMECA_COM') return haystack.indexOf(needle) === 0;
  if (comparison === 'TERMINA_COM') return haystack.slice(-needle.length) === needle;
  return haystack.indexOf(needle) !== -1;
}

function coreMailResolveRoutingByRules_(msgCtx) {
  var rules = coreMailReadRoutingRules_();
  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    if (!coreMailRoutingRuleMatches_(rule, msgCtx)) continue;

    var adapter = coreMailGetModuleAdapter_(rule.moduleName);
    var correlationKey = String(coreMailExtractCorrelationKey_(msgCtx && msgCtx.subject ? msgCtx.subject : '') || '').trim();
    var parsed = correlationKey ? coreMailParseCorrelationKey_(correlationKey) : { isValid: false };

    return Object.freeze({
      matched: true,
      moduleName: rule.moduleName,
      moduleCode: adapter ? adapter.moduleCode : (parsed.moduleCode || ''),
      correlationKey: correlationKey,
      entityType: rule.entityType || parsed.entityType || '',
      entityId: parsed.entityId || parsed.businessId || '',
      stage: rule.flowStep || parsed.stage || '',
      flowCode: parsed.flowCode || '',
      confidence: 0.95,
      reason: 'MAIL_REGRAS:' + rule.id
    });
  }

  return Object.freeze({
    matched: false,
    confidence: 0,
    reason: 'NO_MAIL_RULE_MATCH'
  });
}

function coreMailCreateExampleModuleAdapter_(opts) {
  var moduleName = String(opts.moduleName || '').trim();
  var moduleCode = coreMailNormalizeAdapterId_(opts.moduleCode);
  var entityTypeDefault = String(opts.entityTypeDefault || '').trim();
  var keywordHints = Array.isArray(opts.keywordHints) ? opts.keywordHints.slice() : [];

  function parseByCode(key) {
    var rawKey = coreMailNormalizeCorrelationToken_(coreMailStripWrappedCorrelationKey_(key));
    if (!rawKey) return { isValid: false };

    var directPrefix = moduleCode + '-';
    if (rawKey === moduleCode) {
      return {
        isValid: true,
        moduleCode: moduleCode,
        moduleName: moduleName,
        targetKey: rawKey
      };
    }

    if (rawKey.indexOf(directPrefix) !== 0) {
      return { isValid: false };
    }

    var parts = rawKey.split('-');
    var tail = parts.slice(1);
    var businessId = '';
    var flowCode = '';
    var stage = '';

    if (tail.length >= 3 && /[A-Z]/.test(tail[tail.length - 1]) && /[A-Z]/.test(tail[tail.length - 2])) {
      businessId = tail.slice(0, -2).join('-');
      flowCode = tail[tail.length - 2];
      stage = tail[tail.length - 1];
    } else if (tail.length >= 2 && /[A-Z]/.test(tail[tail.length - 1])) {
      businessId = tail.slice(0, -1).join('-');
      flowCode = tail[tail.length - 1];
    } else {
      businessId = tail.join('-');
    }

    if (!businessId && tail.length) {
      businessId = tail[0];
    }

    return {
      isValid: true,
      moduleCode: moduleCode,
      moduleName: moduleName,
      businessId: businessId,
      flowCode: flowCode,
      stage: stage,
      entityType: entityTypeDefault,
      entityId: businessId,
      targetKey: rawKey
    };
  }

  return {
    moduleName: moduleName,
    moduleCode: moduleCode,

    buildCorrelationKey: function(ctx) {
      ctx = ctx || {};
      var businessId = coreMailNormalizeCorrelationToken_(
        ctx.businessId || ctx.entityId || ctx.targetKey || ctx.threadId || ctx.messageId
      );

      if (!businessId) {
        throw new Error(
          'Adapter ' + moduleCode + ': informe ctx.businessId, ctx.entityId, ctx.targetKey, ctx.threadId ou ctx.messageId.'
        );
      }

      var flowCode = coreMailNormalizeCorrelationToken_(ctx.flowCode || '');
      var stage = coreMailNormalizeCorrelationToken_(ctx.stage || '');
      var out = [moduleCode, businessId];

      if (flowCode) out.push(flowCode);
      if (stage) out.push(stage);

      return out.join('-');
    },

    parseCorrelationKey: function(key) {
      return parseByCode(key);
    },

    matchMessage: function(msgCtx) {
      var haystack = core_normalizeText_(coreMailBuildMessageTextForMatching_(msgCtx || {}), {
        removeAccents: true,
        collapseWhitespace: true,
        caseMode: 'lower'
      });

      for (var i = 0; i < keywordHints.length; i++) {
        var hint = core_normalizeText_(keywordHints[i], {
          removeAccents: true,
          collapseWhitespace: true,
          caseMode: 'lower'
        });

        if (hint && haystack.indexOf(hint) !== -1) {
          return {
            matched: true,
            confidence: 0.65,
            reason: 'KEYWORD:' + keywordHints[i]
          };
        }
      }

      return {
        matched: false,
        confidence: 0
      };
    },

    resolveRouting: function(msgCtx) {
      var keyInSubject = coreMailExtractCorrelationKey_(msgCtx && msgCtx.subject ? msgCtx.subject : '');
      var parsed = keyInSubject ? parseByCode(keyInSubject) : { isValid: false };

      if (parsed.isValid) {
        return {
          matched: true,
          moduleName: moduleName,
          moduleCode: moduleCode,
          correlationKey: keyInSubject,
          entityType: parsed.entityType || '',
          entityId: parsed.entityId || '',
          stage: parsed.stage || '',
          flowCode: parsed.flowCode || '',
          confidence: 1,
          reason: 'CORRELATION_KEY'
        };
      }

      var match = this.matchMessage(msgCtx || {});
      if (!match.matched) {
        return {
          matched: false,
          confidence: 0
        };
      }

      return {
        matched: true,
        moduleName: moduleName,
        moduleCode: moduleCode,
        correlationKey: '',
        entityType: entityTypeDefault,
        entityId: '',
        stage: '',
        flowCode: '',
        confidence: Number(match.confidence || 0),
        reason: match.reason || 'MATCH_MESSAGE'
      };
    },

    normalizeOutgoingSubject: function(subject, ctx) {
      var correlationKey = String((ctx && ctx.correlationKey) || '').trim();
      if (!correlationKey) {
        correlationKey = this.buildCorrelationKey(ctx || {});
      }

      var cleanSubject = String(subject || '').trim();
      var tag = '[GEAPA][' + correlationKey + ']';
      if (!cleanSubject) return tag;

      if (/\[GEAPA\]\[[^\]]+\]/i.test(cleanSubject)) {
        return cleanSubject.replace(/\[GEAPA\]\[[^\]]+\]/i, tag).trim();
      }

      return (tag + ' ' + cleanSubject).trim();
    },

    summarizeForHistory: function(eventRow) {
      eventRow = eventRow || {};
      return {
        moduleName: moduleName,
        moduleCode: moduleCode,
        eventId: eventRow['Id Evento'] || '',
        correlationKey: eventRow['Chave de Correlacao'] || '',
        subject: eventRow.Assunto || '',
        fromEmail: eventRow['Email Remetente'] || ''
      };
    }
  };
}
