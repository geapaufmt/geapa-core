/***************************************
 * 23_core_modules_config.js
 *
 * Camada central de controle operacional por MODULOS_CONFIG.
 *
 * Importante:
 * - Registry resolve recursos: KEY -> spreadsheet/sheet/folder/etc.
 * - MODULOS_CONFIG decide comportamento operacional: modulo, fluxo,
 *   modo, ambiente e capabilities.
 ***************************************/

const CORE_MODULES_CONFIG_SHEET_NAME = 'MODULOS_CONFIG';
const CORE_MODULES_CONFIG_CACHE_KEY = 'GEAPA_CORE_MODULES_CONFIG_V1';
const CORE_MODULES_CONFIG_CACHE_TTL_SECONDS = 15 * 60;
const CORE_MODULES_CONFIG_GENERAL_FLOW = 'GERAL';
const CORE_MODULES_CONFIG_PROD_ENV = 'PROD';

const CORE_MODULES_CONFIG_CAPABILITY_HEADERS = Object.freeze({
  TRIGGER: 'PERMITE_TRIGGER',
  EMAIL: 'PERMITE_EMAIL',
  INBOX: 'PERMITE_INBOX',
  SYNC: 'PERMITE_SYNC',
  DRIVE: 'PERMITE_DRIVE'
});

const CORE_MODULES_CONFIG_UX = Object.freeze({
  modules: Object.freeze([
    'CORE',
    'MEMBROS',
    'SELETIVO',
    'COMUNICACOES',
    'ATIVIDADES',
    'DESLIGAMENTOS',
    'APRESENTACOES'
  ]),
  flows: Object.freeze([
    'GERAL',
    'ATIVIDADES_INTERNAS',
    'PRESENCAS',
    'JUSTIFICATIVAS_FALTAS',
    'MOTOR_DISCIPLINAR',
    'APRESENTACOES',
    'INBOX',
    'EMAIL',
    'SYNC',
    'DRIVE',
    'DIAGNOSTICO'
  ]),
  modes: Object.freeze(['ON', 'OFF', 'MANUAL', 'DRY_RUN']),
  environments: Object.freeze(['PROD', 'DEV']),
  yesNo: Object.freeze(['SIM', 'NAO']),
  notes: Object.freeze({
    MODULO: 'Modulo operacional dono da regra. Ex.: ATIVIDADES, MEMBROS, APRESENTACOES. EVENTOS nao deve ser usado como modulo ativo nesta fase.',
    FLUXO: 'Fluxo dentro do modulo. Use GERAL como fallback padrao do modulo.',
    ATIVO: 'SIM permite avaliar a regra; NAO bloqueia o fluxo antes da execucao.',
    MODO: 'ON executa normalmente; OFF bloqueia; MANUAL bloqueia trigger; DRY_RUN permite leitura, log e diagnostico sem envio real ou escrita destrutiva.',
    AMBIENTE: 'Ambiente da regra. O core procura primeiro o ambiente atual e pode usar PROD como fallback seguro.',
    PERMITE_TRIGGER: 'SIM libera execucoes automaticas por trigger, respeitando ATIVO e MODO.',
    PERMITE_EMAIL: 'SIM libera envio real de e-mail pelo fluxo.',
    PERMITE_INBOX: 'SIM libera leitura/ingestao de inbox pelo fluxo.',
    PERMITE_SYNC: 'SIM libera rotinas de sincronizacao pelo fluxo.',
    PERMITE_DRIVE: 'SIM libera operacoes em Drive pelo fluxo.',
    JANELA_MINUTOS: 'Janela operacional opcional, em minutos. Deixe vazio quando nao se aplicar.',
    ULTIMA_ALTERACAO: 'Data/hora da ultima alteracao operacional registrada.',
    ALTERADO_POR: 'Pessoa ou funcao responsavel pela alteracao.',
    OBS: 'Observacoes operacionais, contexto de contingencia ou motivo da regra.'
  }),
  colorGroups: Object.freeze([
    Object.freeze({ color: '#d9ead3', headers: ['MODULO', 'FLUXO', 'AMBIENTE'] }),
    Object.freeze({ color: '#fff2cc', headers: ['ATIVO', 'MODO'] }),
    Object.freeze({ color: '#d0e0e3', headers: ['PERMITE_TRIGGER', 'PERMITE_EMAIL', 'PERMITE_INBOX', 'PERMITE_SYNC', 'PERMITE_DRIVE'] }),
    Object.freeze({ color: '#ead1dc', headers: ['JANELA_MINUTOS', 'ULTIMA_ALTERACAO', 'ALTERADO_POR'] }),
    Object.freeze({ color: '#fce5cd', headers: ['OBS'] })
  ])
});

function core_getModulesConfigSheet_() {
  const ss = core_openSpreadsheetById_(CORE_REGISTRY_SPREADSHEET_ID);
  const sh = ss.getSheetByName(CORE_MODULES_CONFIG_SHEET_NAME);

  if (!sh) {
    throw new Error(
      'Aba "' + CORE_MODULES_CONFIG_SHEET_NAME + '" nao encontrada na planilha geral do Registry (' +
      CORE_REGISTRY_SPREADSHEET_ID + ').'
    );
  }

  return sh;
}

function core_modulesConfigNormalizeKey_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  }).replace(/\s+/g, '_');
}

function core_modulesConfigParseYesNo_(value, label, lineNo) {
  const s = core_modulesConfigNormalizeKey_(value);
  if (s === 'SIM') return true;
  if (s === 'NAO') return false;

  throw new Error(
    'MODULOS_CONFIG invalida na linha ' + lineNo + ': coluna ' + label +
    ' deve ser "SIM" ou "NAO". Valor recebido: "' + (s || '(vazio)') + '".'
  );
}

function core_modulesConfigParseMode_(value, lineNo) {
  const mode = core_modulesConfigNormalizeKey_(value || 'ON');
  const allowed = {
    ON: true,
    OFF: true,
    MANUAL: true,
    DRY_RUN: true
  };

  if (allowed[mode]) return mode;

  throw new Error(
    'MODULOS_CONFIG invalida na linha ' + lineNo +
    ': coluna MODO deve ser ON, OFF, MANUAL ou DRY_RUN. Valor recebido: "' +
    (mode || '(vazio)') + '".'
  );
}

function core_modulesConfigParseEnv_(value, lineNo) {
  const env = core_modulesConfigNormalizeKey_(value || CORE_MODULES_CONFIG_PROD_ENV);
  if (env === 'DEV' || env === 'PROD') return env;

  throw new Error(
    'MODULOS_CONFIG invalida na linha ' + lineNo +
    ': coluna AMBIENTE deve ser DEV ou PROD. Valor recebido: "' +
    (env || '(vazio)') + '".'
  );
}

function core_modulesConfigParseWindowMinutes_(value) {
  if (value === '' || value == null) return null;
  const n = Number(value);
  if (isNaN(n) || n < 0) return null;
  return n;
}

function core_getModulesConfigRows_() {
  const cached = core_modulesConfigCacheGet_();
  if (cached) return cached;

  const sh = core_getModulesConfigSheet_();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return Object.freeze([]);

  const headerMap = core_headerMap_(sh, 1);
  const required = [
    'MODULO',
    'FLUXO',
    'ATIVO',
    'MODO',
    'AMBIENTE',
    'PERMITE_TRIGGER',
    'PERMITE_EMAIL',
    'PERMITE_INBOX',
    'PERMITE_SYNC',
    'PERMITE_DRIVE'
  ];

  const missing = required.filter(function(headerName) {
    return !core_getCol_(headerMap, headerName);
  });

  if (missing.length) {
    throw new Error('MODULOS_CONFIG invalida. Cabecalhos ausentes: ' + missing.join(', '));
  }

  const colModulo = core_getCol_(headerMap, 'MODULO');
  const colFluxo = core_getCol_(headerMap, 'FLUXO');
  const colAtivo = core_getCol_(headerMap, 'ATIVO');
  const colModo = core_getCol_(headerMap, 'MODO');
  const colAmbiente = core_getCol_(headerMap, 'AMBIENTE');
  const colJanela = core_getCol_(headerMap, 'JANELA_MINUTOS');
  const colUltimaAlteracao = core_getCol_(headerMap, 'ULTIMA_ALTERACAO');
  const colAlteradoPor = core_getCol_(headerMap, 'ALTERADO_POR');
  const colObs = core_getCol_(headerMap, 'OBS');

  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = [];

  values.forEach(function(row, idx) {
    const lineNo = idx + 2;
    const moduleName = core_modulesConfigNormalizeKey_(row[colModulo - 1]);
    const flowName = core_modulesConfigNormalizeKey_(row[colFluxo - 1] || CORE_MODULES_CONFIG_GENERAL_FLOW);

    if (!moduleName) return;

    const capabilities = {};
    Object.keys(CORE_MODULES_CONFIG_CAPABILITY_HEADERS).forEach(function(capability) {
      const headerName = CORE_MODULES_CONFIG_CAPABILITY_HEADERS[capability];
      const col = core_getCol_(headerMap, headerName);
      capabilities[capability] = core_modulesConfigParseYesNo_(row[col - 1], headerName, lineNo);
    });

    rows.push(Object.freeze({
      moduleName: moduleName,
      flowName: flowName || CORE_MODULES_CONFIG_GENERAL_FLOW,
      active: core_modulesConfigParseYesNo_(row[colAtivo - 1], 'ATIVO', lineNo),
      mode: core_modulesConfigParseMode_(row[colModo - 1], lineNo),
      ambiente: core_modulesConfigParseEnv_(row[colAmbiente - 1], lineNo),
      capabilities: Object.freeze(capabilities),
      windowMinutes: colJanela ? core_modulesConfigParseWindowMinutes_(row[colJanela - 1]) : null,
      updatedAt: colUltimaAlteracao ? row[colUltimaAlteracao - 1] || '' : '',
      updatedBy: colAlteradoPor ? String(row[colAlteradoPor - 1] || '').trim() : '',
      notes: colObs ? String(row[colObs - 1] || '').trim() : '',
      lineNo: lineNo
    }));
  });

  const frozen = Object.freeze(rows);
  core_modulesConfigCacheSet_(frozen);
  return frozen;
}

function core_getModulesConfigIndex_() {
  const rows = core_getModulesConfigRows_();
  const byKey = {};
  const duplicates = [];

  rows.forEach(function(row) {
    const key = core_modulesConfigBuildLookupKey_(row.moduleName, row.flowName, row.ambiente);
    if (byKey[key]) {
      duplicates.push({
        key: key,
        firstLineNo: byKey[key].lineNo,
        duplicateLineNo: row.lineNo
      });
      return;
    }

    byKey[key] = row;
  });

  return Object.freeze({
    rows: rows,
    byKey: Object.freeze(byKey),
    duplicates: Object.freeze(duplicates)
  });
}

function core_modulesConfigBuildLookupKey_(moduleName, flowName, ambiente) {
  return [
    core_modulesConfigNormalizeKey_(moduleName),
    core_modulesConfigNormalizeKey_(flowName || CORE_MODULES_CONFIG_GENERAL_FLOW),
    core_modulesConfigNormalizeKey_(ambiente || CORE_MODULES_CONFIG_PROD_ENV)
  ].join('|');
}

function core_modulesConfigResolveEnv_(opts) {
  opts = opts || {};
  return opts.ambiente
    ? core_modulesConfigParseEnv_(opts.ambiente, 0)
    : core_getCurrentEnv_();
}

function core_getModuleConfig_(moduleName, flowName, opts) {
  core_assertRequired_(moduleName, 'moduleName');

  opts = opts || {};
  const moduleKey = core_modulesConfigNormalizeKey_(moduleName);
  const flowKey = core_modulesConfigNormalizeKey_(flowName || CORE_MODULES_CONFIG_GENERAL_FLOW);
  const currentEnv = core_modulesConfigResolveEnv_(opts);
  const allowProdFallback = opts.allowProdFallback !== false;
  const index = core_getModulesConfigIndex_();

  const attempts = [
    { flowName: flowKey, ambiente: currentEnv, fallbackType: 'EXACT' },
    { flowName: CORE_MODULES_CONFIG_GENERAL_FLOW, ambiente: currentEnv, fallbackType: 'MODULE_GENERAL' }
  ];

  if (allowProdFallback && currentEnv !== CORE_MODULES_CONFIG_PROD_ENV) {
    attempts.push(
      { flowName: flowKey, ambiente: CORE_MODULES_CONFIG_PROD_ENV, fallbackType: 'EXACT_PROD_FALLBACK' },
      { flowName: CORE_MODULES_CONFIG_GENERAL_FLOW, ambiente: CORE_MODULES_CONFIG_PROD_ENV, fallbackType: 'MODULE_GENERAL_PROD_FALLBACK' }
    );
  }

  for (var i = 0; i < attempts.length; i++) {
    const attempt = attempts[i];
    const key = core_modulesConfigBuildLookupKey_(moduleKey, attempt.flowName, attempt.ambiente);
    const row = index.byKey[key];
    if (!row) continue;

    return Object.freeze({
      moduleName: row.moduleName,
      requestedModuleName: moduleKey,
      flowName: row.flowName,
      requestedFlowName: flowKey,
      active: row.active,
      mode: row.mode,
      ambiente: row.ambiente,
      requestedAmbiente: currentEnv,
      fallbackType: attempt.fallbackType,
      isFallback: attempt.fallbackType !== 'EXACT',
      capabilities: row.capabilities,
      windowMinutes: row.windowMinutes,
      updatedAt: row.updatedAt,
      updatedBy: row.updatedBy,
      notes: row.notes,
      lineNo: row.lineNo
    });
  }

  if (Object.prototype.hasOwnProperty.call(opts, 'defaultWhenMissing')) {
    return opts.defaultWhenMissing;
  }

  throw new Error(
    'MODULOS_CONFIG nao possui configuracao para MODULO="' + moduleKey +
    '", FLUXO="' + flowKey + '", AMBIENTE="' + currentEnv +
    '". Cadastre a linha exata ou a linha ' + moduleKey + ' + GERAL.'
  );
}

function core_isModuleEnabled_(moduleName, flowName, opts) {
  const config = core_getModuleConfig_(moduleName, flowName, opts || {});
  return config.active === true && config.mode !== 'OFF';
}

function core_getModuleMode_(moduleName, flowName, opts) {
  return core_getModuleConfig_(moduleName, flowName, opts || {}).mode;
}

function core_canModuleUseCapability_(moduleName, flowName, capability, opts) {
  const config = core_getModuleConfig_(moduleName, flowName, opts || {});
  return core_modulesConfigEvaluateExecution_(config, capability, opts || {}).allowed;
}

function core_assertModuleExecutionAllowed_(moduleName, flowName, capability, opts) {
  opts = opts || {};
  const config = core_getModuleConfig_(moduleName, flowName, opts);
  const decision = core_modulesConfigEvaluateExecution_(config, capability, opts);

  if (!decision.allowed) {
    throw new Error(core_modulesConfigBuildBlockMessage_(config, capability, decision));
  }

  return Object.freeze({
    allowed: true,
    dryRun: config.mode === 'DRY_RUN',
    reason: decision.reason,
    config: config
  });
}

function core_modulesConfigEvaluateExecution_(config, capability, opts) {
  opts = opts || {};
  const cap = core_modulesConfigNormalizeKey_(capability || '');
  const executionType = core_modulesConfigNormalizeKey_(opts.executionType || cap || 'MANUAL');

  if (!config.active) {
    return Object.freeze({ allowed: false, reason: 'ATIVO=NAO' });
  }

  if (config.mode === 'OFF') {
    return Object.freeze({ allowed: false, reason: 'MODO=OFF' });
  }

  if (config.mode === 'MANUAL' && executionType === 'TRIGGER') {
    return Object.freeze({ allowed: false, reason: 'MODO=MANUAL bloqueia execucao por trigger' });
  }

  if (cap) {
    if (!Object.prototype.hasOwnProperty.call(CORE_MODULES_CONFIG_CAPABILITY_HEADERS, cap)) {
      return Object.freeze({ allowed: false, reason: 'Capability invalida: ' + cap });
    }

    if (config.capabilities[cap] !== true) {
      return Object.freeze({ allowed: false, reason: 'Capability bloqueada: ' + cap });
    }
  }

  return Object.freeze({
    allowed: true,
    reason: config.mode === 'DRY_RUN' ? 'MODO=DRY_RUN' : 'PERMITIDO'
  });
}

function core_modulesConfigBuildBlockMessage_(config, capability, decision) {
  return [
    'GEAPA-CORE: fluxo bloqueado por MODULOS_CONFIG.',
    'Modulo: ' + config.moduleName,
    'Fluxo: ' + config.flowName,
    'Ambiente: ' + config.ambiente,
    'Capability: ' + (core_modulesConfigNormalizeKey_(capability || '') || '(nao informada)'),
    'Motivo: ' + decision.reason,
    'Linha: ' + config.lineNo
  ].join('\n');
}

function core_debugModulesConfig_() {
  const index = core_getModulesConfigIndex_();
  const modules = {};
  const environments = {};
  const modes = {};

  index.rows.forEach(function(row) {
    if (!modules[row.moduleName]) modules[row.moduleName] = [];
    modules[row.moduleName].push(row.flowName + '@' + row.ambiente);
    environments[row.ambiente] = true;
    modes[row.mode] = (modes[row.mode] || 0) + 1;
  });

  const examples = {};
  ['CORE', 'MEMBROS', 'SELETIVO', 'COMUNICACOES', 'ATIVIDADES', 'DESLIGAMENTOS', 'APRESENTACOES'].forEach(function(moduleName) {
    try {
      examples[moduleName] = core_getModuleConfig_(moduleName, CORE_MODULES_CONFIG_GENERAL_FLOW, {
        defaultWhenMissing: null
      });
    } catch (e) {
      examples[moduleName] = { error: e.message };
    }
  });

  return Object.freeze({
    sheetName: CORE_MODULES_CONFIG_SHEET_NAME,
    totalRows: index.rows.length,
    modules: Object.keys(modules).sort(),
    environments: Object.keys(environments).sort(),
    modes: Object.freeze(modes),
    duplicates: index.duplicates,
    examples: Object.freeze(examples)
  });
}

function core_applyModulesConfigSheetUx_(opts) {
  opts = opts || {};
  const sheet = core_getModulesConfigSheet_();
  const headerRow = Number(opts.headerRow || 1);
  const operations = [];

  operations.push(core_modulesConfigTryUxOperation_('freezeHeaderRow', function() {
    return core_freezeHeaderRow_(sheet, headerRow);
  }));

  operations.push(core_modulesConfigTryUxOperation_('headerNotes', function() {
    return core_applyHeaderNotes_(sheet, CORE_MODULES_CONFIG_UX.notes, headerRow);
  }));

  operations.push(core_modulesConfigTryUxOperation_('headerColors', function() {
    return core_applyHeaderColors_(sheet, CORE_MODULES_CONFIG_UX.colorGroups, headerRow, {
      defaultColor: '#f3f3f3',
      fontColor: '#202124',
      fontWeight: 'bold',
      wrap: true
    });
  }));

  operations.push(core_modulesConfigTryUxOperation_('filter', function() {
    return core_ensureFilter_(sheet, headerRow, { recreate: false });
  }));

  operations.push(core_modulesConfigTryUxOperation_('dropdownValidation', function() {
    return core_applyDropdownValidationByHeader_(sheet, core_modulesConfigBuildDropdownRules_(), headerRow, {});
  }));

  operations.push(core_modulesConfigTryUxOperation_('columnWidths', function() {
    return core_modulesConfigApplyColumnWidths_(sheet, headerRow);
  }));

  operations.push(core_modulesConfigTryUxOperation_('dataFormatting', function() {
    return core_modulesConfigApplyDataFormatting_(sheet, headerRow);
  }));

  return Object.freeze({
    ok: true,
    sheetName: sheet.getName(),
    operations: Object.freeze(operations)
  });
}

function core_modulesConfigBuildDropdownRules_() {
  const yesNoRule = Object.freeze({
    values: CORE_MODULES_CONFIG_UX.yesNo,
    allowInvalid: false,
    helpText: 'Use SIM ou NAO.'
  });

  return Object.freeze({
    MODULO: Object.freeze({
      values: CORE_MODULES_CONFIG_UX.modules,
      allowInvalid: true,
      helpText: 'Modulo operacional conhecido pelo ecossistema nesta fase.'
    }),
    FLUXO: Object.freeze({
      values: CORE_MODULES_CONFIG_UX.flows,
      allowInvalid: true,
      helpText: 'Use GERAL para fallback do modulo ou um fluxo especifico.'
    }),
    ATIVO: yesNoRule,
    MODO: Object.freeze({
      values: CORE_MODULES_CONFIG_UX.modes,
      allowInvalid: false,
      helpText: 'ON, OFF, MANUAL ou DRY_RUN.'
    }),
    AMBIENTE: Object.freeze({
      values: CORE_MODULES_CONFIG_UX.environments,
      allowInvalid: false,
      helpText: 'Ambiente da regra operacional.'
    }),
    PERMITE_TRIGGER: yesNoRule,
    PERMITE_EMAIL: yesNoRule,
    PERMITE_INBOX: yesNoRule,
    PERMITE_SYNC: yesNoRule,
    PERMITE_DRIVE: yesNoRule
  });
}

function core_modulesConfigTryUxOperation_(operation, fn) {
  try {
    return Object.freeze({
      operation: operation,
      status: 'APPLIED',
      result: fn()
    });
  } catch (err) {
    return Object.freeze({
      operation: operation,
      status: 'ERROR',
      reason: err && err.message ? err.message : String(err || 'UX_ERROR')
    });
  }
}

function core_modulesConfigApplyColumnWidths_(sheet, headerRow) {
  const headerMap = core_headerMap_(sheet, headerRow || 1);
  const widths = {
    MODULO: 150,
    FLUXO: 210,
    ATIVO: 90,
    MODO: 110,
    AMBIENTE: 105,
    PERMITE_TRIGGER: 135,
    PERMITE_EMAIL: 125,
    PERMITE_INBOX: 125,
    PERMITE_SYNC: 120,
    PERMITE_DRIVE: 125,
    JANELA_MINUTOS: 135,
    ULTIMA_ALTERACAO: 170,
    ALTERADO_POR: 180,
    OBS: 360
  };
  let applied = 0;

  Object.keys(widths).forEach(function(headerName) {
    const col = core_getCol_(headerMap, headerName);
    if (!col) return;
    sheet.setColumnWidth(col, widths[headerName]);
    applied++;
  });

  return applied;
}

function core_modulesConfigApplyDataFormatting_(sheet, headerRow) {
  const row = Math.max(1, Number(headerRow || 1));
  const headerMap = core_headerMap_(sheet, row);
  const maxRows = Math.max(sheet.getMaxRows(), row + 1);
  const dataRows = Math.max(maxRows - row, 1);
  let applied = 0;

  const colJanela = core_getCol_(headerMap, 'JANELA_MINUTOS');
  if (colJanela) {
    sheet.getRange(row + 1, colJanela, dataRows, 1).setNumberFormat('0');
    applied++;
  }

  const colUltimaAlteracao = core_getCol_(headerMap, 'ULTIMA_ALTERACAO');
  if (colUltimaAlteracao) {
    sheet.getRange(row + 1, colUltimaAlteracao, dataRows, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    applied++;
  }

  const colObs = core_getCol_(headerMap, 'OBS');
  if (colObs) {
    sheet.getRange(1, colObs, maxRows, 1)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setVerticalAlignment('middle');
    applied++;
  }

  sheet.getRange(1, 1, maxRows, Math.max(sheet.getLastColumn(), 1)).setVerticalAlignment('middle');
  return applied;
}

function core_modulesConfigCacheGet_() {
  const s = CacheService.getScriptCache().get(CORE_MODULES_CONFIG_CACHE_KEY);
  if (!s) return null;

  try {
    const arr = JSON.parse(s);
    return Object.freeze(arr.map(function(item) {
      return Object.freeze({
        moduleName: item.moduleName,
        flowName: item.flowName,
        active: item.active,
        mode: item.mode,
        ambiente: item.ambiente,
        capabilities: Object.freeze(item.capabilities || {}),
        windowMinutes: item.windowMinutes,
        updatedAt: item.updatedAt,
        updatedBy: item.updatedBy,
        notes: item.notes,
        lineNo: item.lineNo
      });
    }));
  } catch (e) {
    return null;
  }
}

function core_modulesConfigCacheSet_(rows) {
  CacheService.getScriptCache().put(
    CORE_MODULES_CONFIG_CACHE_KEY,
    JSON.stringify(rows),
    CORE_MODULES_CONFIG_CACHE_TTL_SECONDS
  );
}

function core_modulesConfigCacheClear_() {
  CacheService.getScriptCache().remove(CORE_MODULES_CONFIG_CACHE_KEY);
}
