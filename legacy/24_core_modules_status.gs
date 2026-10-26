/***************************************
 * 24_core_modules_status.js
 *
 * Log/status operacional leve por MODULO + FLUXO.
 *
 * Importante:
 * - MODULOS_CONFIG controla se pode executar.
 * - MODULOS_STATUS registra o que aconteceu.
 ***************************************/

const CORE_MODULES_STATUS_SHEET_NAME = 'MODULOS_STATUS';
const CORE_MODULES_STATUS_GENERAL_FLOW = 'GERAL';

const CORE_MODULES_STATUS_HEADERS = Object.freeze([
  'MODULO',
  'FLUXO',
  'ULTIMA_EXECUCAO',
  'ULTIMO_SUCESSO',
  'ULTIMO_ERRO',
  'MENSAGEM_ULTIMO_ERRO',
  'ULTIMO_BLOQUEIO_CONFIG',
  'MOTIVO_ULTIMO_BLOQUEIO',
  'ULTIMO_MODO_LIDO',
  'ULTIMA_CAPABILITY',
  'EXECUCOES_24H',
  'BLOQUEIOS_24H',
  'SUCESSOS_24H',
  'ERROS_24H',
  'OBS'
]);

function core_getModulesStatusSheet_() {
  const ss = core_openSpreadsheetById_(CORE_REGISTRY_SPREADSHEET_ID);
  const sh = ss.getSheetByName(CORE_MODULES_STATUS_SHEET_NAME);

  if (!sh) {
    throw new Error(
      'Aba "' + CORE_MODULES_STATUS_SHEET_NAME + '" nao encontrada na planilha geral do Registry (' +
      CORE_REGISTRY_SPREADSHEET_ID + ').'
    );
  }

  return sh;
}

function core_moduleStatusNormalizeKey_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  }).replace(/\s+/g, '_');
}

function core_moduleStatusNormalizeCapability_(value) {
  return core_moduleStatusNormalizeKey_(value || '');
}

function core_moduleStatusNormalizeMode_(value) {
  return core_moduleStatusNormalizeKey_(value || '');
}

function core_moduleStatusAssertSchema_(sheet) {
  const sh = sheet || core_getModulesStatusSheet_();
  const headerMap = core_headerMap_(sh, 1);
  const missing = CORE_MODULES_STATUS_HEADERS.filter(function(headerName) {
    return !core_getCol_(headerMap, headerName);
  });

  if (missing.length) {
    throw new Error('MODULOS_STATUS invalida. Cabecalhos ausentes: ' + missing.join(', '));
  }

  return headerMap;
}

function core_moduleStatusRowToObject_(headers, row, rowNumber) {
  const out = {};
  headers.forEach(function(headerName, index) {
    if (!headerName) return;
    out[headerName] = row[index];
  });
  out.rowNumber = rowNumber;
  return Object.freeze(out);
}

function core_findModuleStatusRow_(moduleName, flowName) {
  core_assertRequired_(moduleName, 'moduleName');

  const moduleKey = core_moduleStatusNormalizeKey_(moduleName);
  const flowKey = core_moduleStatusNormalizeKey_(flowName || CORE_MODULES_STATUS_GENERAL_FLOW);
  const sheet = core_getModulesStatusSheet_();
  const headerMap = core_moduleStatusAssertSchema_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  const colModulo = core_getCol_(headerMap, 'MODULO');
  const colFluxo = core_getCol_(headerMap, 'FLUXO');

  if (lastRow < 2) {
    return Object.freeze({
      found: false,
      sheet: sheet,
      headerMap: headerMap,
      headers: Object.freeze(headers),
      moduleName: moduleKey,
      flowName: flowKey
    });
  }

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (var i = 0; i < values.length; i++) {
    const row = values[i];
    const rowModule = core_moduleStatusNormalizeKey_(row[colModulo - 1]);
    const rowFlow = core_moduleStatusNormalizeKey_(row[colFluxo - 1]);

    if (rowModule === moduleKey && rowFlow === flowKey) {
      return Object.freeze({
        found: true,
        sheet: sheet,
        headerMap: headerMap,
        headers: Object.freeze(headers),
        moduleName: moduleKey,
        flowName: flowKey,
        rowNumber: i + 2,
        record: core_moduleStatusRowToObject_(headers, row, i + 2)
      });
    }
  }

  return Object.freeze({
    found: false,
    sheet: sheet,
    headerMap: headerMap,
    headers: Object.freeze(headers),
    moduleName: moduleKey,
    flowName: flowKey
  });
}

function core_moduleStatusGet_(moduleName, flowName, opts) {
  opts = opts || {};
  const found = core_findModuleStatusRow_(moduleName, flowName);

  if (found.found) return found.record;

  if (opts.createIfMissing === true) {
    return core_moduleStatusEnsureRow_(moduleName, flowName, opts);
  }

  if (Object.prototype.hasOwnProperty.call(opts, 'defaultWhenMissing')) {
    return opts.defaultWhenMissing;
  }

  return null;
}

function core_moduleStatusEnsureRow_(moduleName, flowName, opts) {
  opts = opts || {};
  const found = core_findModuleStatusRow_(moduleName, flowName);
  if (found.found) return found.record;

  const payload = {
    MODULO: found.moduleName,
    FLUXO: found.flowName,
    EXECUCOES_24H: 0,
    BLOQUEIOS_24H: 0,
    SUCESSOS_24H: 0,
    ERROS_24H: 0,
    OBS: opts.obs || ''
  };

  core_appendObjectByHeaders_(found.sheet, payload, { headerRow: 1 });
  return core_moduleStatusGet_(found.moduleName, found.flowName, { defaultWhenMissing: null });
}

function core_moduleStatusNumber_(value) {
  const n = Number(value);
  return isNaN(n) ? 0 : n;
}

function core_moduleStatusBuildReason_(reasonCode, reasonMessage) {
  const code = String(reasonCode || '').trim();
  const msg = String(reasonMessage || '').trim();
  if (code && msg) return code + ': ' + msg;
  return code || msg || '';
}

function core_moduleStatusMessageFromError_(errorOrMessage) {
  if (!errorOrMessage) return '';
  if (errorOrMessage && errorOrMessage.message) return String(errorOrMessage.message || '').trim();
  return String(errorOrMessage || '').trim();
}

function core_moduleStatusWriteUpdates_(ctx, updates) {
  const sheet = ctx.sheet;
  const headerMap = ctx.headerMap;
  const rowNumber = ctx.rowNumber;
  let written = 0;

  Object.keys(updates || {}).forEach(function(headerName) {
    if (!core_getCol_(headerMap, headerName)) return;
    core_writeCellByHeader_(sheet, rowNumber, headerMap, headerName, updates[headerName], {
      oneBased: true
    });
    written++;
  });

  return written;
}

function core_moduleStatusUpdate_(moduleName, flowName, updates, opts) {
  opts = opts || {};
  core_moduleStatusEnsureRow_(moduleName, flowName, opts);

  const ctx = core_findModuleStatusRow_(moduleName, flowName);
  if (!ctx.found) {
    throw new Error('Nao foi possivel garantir linha em MODULOS_STATUS para ' + moduleName + ' / ' + flowName + '.');
  }

  const written = core_moduleStatusWriteUpdates_(ctx, updates || {});
  const after = core_moduleStatusGet_(ctx.moduleName, ctx.flowName, { defaultWhenMissing: null });

  return Object.freeze({
    ok: true,
    moduleName: ctx.moduleName,
    flowName: ctx.flowName,
    rowNumber: ctx.rowNumber,
    written: written,
    status: after
  });
}

function core_moduleStatusCommonUpdates_(record, capability, opts) {
  opts = opts || {};
  const updates = {};
  const cap = core_moduleStatusNormalizeCapability_(capability);
  const mode = core_moduleStatusNormalizeMode_(opts.modeRead || opts.mode || '');

  if (cap) updates.ULTIMA_CAPABILITY = cap;
  if (mode) updates.ULTIMO_MODO_LIDO = mode;
  if (Object.prototype.hasOwnProperty.call(opts, 'obs')) updates.OBS = opts.obs || '';

  return updates;
}

function core_moduleStatusMarkExecution_(moduleName, flowName, capability, opts) {
  opts = opts || {};
  const current = core_moduleStatusEnsureRow_(moduleName, flowName, opts);
  const updates = core_moduleStatusCommonUpdates_(current, capability, opts);
  updates.ULTIMA_EXECUCAO = opts.timestamp || new Date();
  updates.EXECUCOES_24H = core_moduleStatusNumber_(current.EXECUCOES_24H) + 1;

  return core_moduleStatusUpdate_(moduleName, flowName, updates, opts);
}

function core_moduleStatusMarkSuccess_(moduleName, flowName, capability, opts) {
  opts = opts || {};
  const current = core_moduleStatusEnsureRow_(moduleName, flowName, opts);
  const updates = core_moduleStatusCommonUpdates_(current, capability, opts);
  updates.ULTIMO_SUCESSO = opts.timestamp || new Date();
  updates.SUCESSOS_24H = core_moduleStatusNumber_(current.SUCESSOS_24H) + 1;

  return core_moduleStatusUpdate_(moduleName, flowName, updates, opts);
}

function core_moduleStatusMarkError_(moduleName, flowName, errorOrMessage, capability, opts) {
  opts = opts || {};
  const current = core_moduleStatusEnsureRow_(moduleName, flowName, opts);
  const updates = core_moduleStatusCommonUpdates_(current, capability, opts);
  updates.ULTIMO_ERRO = opts.timestamp || new Date();
  updates.MENSAGEM_ULTIMO_ERRO = core_moduleStatusMessageFromError_(errorOrMessage);
  updates.ERROS_24H = core_moduleStatusNumber_(current.ERROS_24H) + 1;

  return core_moduleStatusUpdate_(moduleName, flowName, updates, opts);
}

function core_moduleStatusMarkBlocked_(moduleName, flowName, reasonCode, reasonMessage, capability, modeRead, opts) {
  opts = opts || {};
  opts.modeRead = modeRead || opts.modeRead || opts.mode || '';

  const current = core_moduleStatusEnsureRow_(moduleName, flowName, opts);
  const updates = core_moduleStatusCommonUpdates_(current, capability, opts);
  updates.ULTIMO_BLOQUEIO_CONFIG = opts.timestamp || new Date();
  updates.MOTIVO_ULTIMO_BLOQUEIO = core_moduleStatusBuildReason_(reasonCode, reasonMessage);
  updates.BLOQUEIOS_24H = core_moduleStatusNumber_(current.BLOQUEIOS_24H) + 1;

  return core_moduleStatusUpdate_(moduleName, flowName, updates, opts);
}

function core_debugModulesStatus_() {
  const sheet = core_getModulesStatusSheet_();
  const headerMap = core_moduleStatusAssertSchema_(sheet);
  const lastRow = sheet.getLastRow();
  const rows = [];

  if (lastRow >= 2) {
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(header) {
      return String(header || '').trim();
    });
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    values.forEach(function(row, idx) {
      const record = core_moduleStatusRowToObject_(headers, row, idx + 2);
      if (record.MODULO || record.FLUXO) rows.push(record);
    });
  }

  return Object.freeze({
    sheetName: CORE_MODULES_STATUS_SHEET_NAME,
    totalRows: rows.length,
    headersOk: !!headerMap,
    modules: Object.freeze(rows.reduce(function(acc, record) {
      const moduleName = core_moduleStatusNormalizeKey_(record.MODULO);
      if (moduleName && acc.indexOf(moduleName) === -1) acc.push(moduleName);
      return acc;
    }, []).sort()),
    sample: Object.freeze(rows.slice(0, 10))
  });
}
