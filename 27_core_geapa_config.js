/**
 * ============================================================
 * 27_core_geapa_config.js
 * ============================================================
 *
 * Leitura centralizada da aba Config_GEAPA.
 *
 * Formato atual (vertical):
 * KEY | VALOR | TIPO | GRUPO | ATIVO | DESCRICAO | OBSERVACOES
 *
 * Transicao:
 * O fallback horizontal abaixo existe apenas para compatibilidade temporaria
 * com a versao antiga da aba, em que a linha 1 continha chaves e a linha 2
 * continha valores. Novas integracoes devem sempre usar o formato vertical.
 */

var CORE_GEAPA_CONFIG_CFG = Object.freeze({
  registryKeys: Object.freeze([
    'CONFIG_GEAPA',
    'DADOS_OFICIAIS_GEAPA',
    'GEAPA_CONFIG'
  ]),
  sheetNames: Object.freeze([
    'Config_GEAPA',
    'CONFIG_GEAPA',
    'Config GEAPA'
  ]),
  headerRow: 1,
  expectedHeaders: Object.freeze([
    'KEY',
    'VALOR',
    'TIPO',
    'GRUPO',
    'ATIVO',
    'DESCRICAO',
    'OBSERVACOES'
  ]),
  aliases: Object.freeze({
    'CURSO_MAE': 'CURSO_MAE',
    'CURSO MAE': 'CURSO_MAE',
    'INSTITUTO_MAE': 'INSTITUTO_MAE',
    'INSTITUTO MAE': 'INSTITUTO_MAE',
    'SIGLA_INSTITUTO_MAE': 'SIGLA_INSTITUTO_MAE',
    'SIGLA INSTITUTO MAE': 'SIGLA_INSTITUTO_MAE',
    'LOCAL_PADRAO_REUNIOES': 'LOCAL_PADRAO_REUNIOES',
    'LOCAL_PADRAO_REUNIÕES': 'LOCAL_PADRAO_REUNIOES'
  })
});

function core_normalizeGeapaConfigKey_(key) {
  var normalized = core_normalizeText_(key, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  });
  normalized = normalized.replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
  return CORE_GEAPA_CONFIG_CFG.aliases[normalized] || normalized;
}

function core_isGeapaConfigActive_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  }) === 'SIM';
}

function core_getGeapaConfigSheet_() {
  var keys = CORE_GEAPA_CONFIG_CFG.registryKeys;
  var errors = [];

  for (var i = 0; i < keys.length; i++) {
    try {
      return {
        sheet: core_getSheetByKey_(keys[i]),
        registryKey: keys[i],
        source: 'REGISTRY'
      };
    } catch (err) {
      errors.push(keys[i] + ': ' + err.message);
    }
  }

  var wantedNames = CORE_GEAPA_CONFIG_CFG.sheetNames.map(function(name) {
    return core_normalizeGeapaConfigKey_(name);
  });

  try {
    var currentEnv = core_getCurrentEnv_();
    var raw = core_getRegistryRaw_();
    var registryKeys = Object.keys(raw);

    for (var r = 0; r < registryKeys.length; r++) {
      var envMap = raw[registryKeys[r]];
      var envs = Object.keys(envMap || {});

      for (var e = 0; e < envs.length; e++) {
        var entry = envMap[envs[e]];
        if (!entry || !entry.ativo) continue;
        if (!core_registryMatchesEnv_(entry.ambiente, currentEnv)) continue;
        if (wantedNames.indexOf(core_normalizeGeapaConfigKey_(entry.sheet)) < 0) continue;

        return {
          sheet: core_getSheetById_(entry.id, entry.sheet),
          registryKey: entry.key,
          source: 'REGISTRY_SHEET_NAME'
        };
      }
    }
  } catch (fallbackErr) {
    errors.push('fallback por nome de aba: ' + fallbackErr.message);
  }

  throw new Error(
    'Config_GEAPA nao encontrada via Registry. Cadastre uma KEY como CONFIG_GEAPA ' +
    'ou mantenha a chave legada DADOS_OFICIAIS_GEAPA apontando para a aba Config_GEAPA. ' +
    'Detalhes: ' + errors.join(' | ')
  );
}

function core_readGeapaConfigDisplayTable_(sheet, opts) {
  opts = opts || {};
  var headerRow = Number(opts.headerRow || CORE_GEAPA_CONFIG_CFG.headerRow);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < headerRow || lastCol < 1) {
    return {
      headerRow: headerRow,
      headers: [],
      rows: [],
      headerMap: {}
    };
  }

  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0]
    .map(function(header) { return String(header || '').trim(); });
  var rows = lastRow > headerRow
    ? sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getDisplayValues()
    : [];

  return {
    headerRow: headerRow,
    headers: headers,
    rows: rows,
    headerMap: core_buildHeaderIndexMap_(headers, {
      normalize: true,
      oneBased: false,
      keepFirst: true
    })
  };
}

function core_isGeapaConfigVerticalTable_(table) {
  return Object.prototype.hasOwnProperty.call(table.headerMap, core_normalizeHeader_('KEY')) &&
    Object.prototype.hasOwnProperty.call(table.headerMap, core_normalizeHeader_('VALOR'));
}

function core_getGeapaConfigCell_(row, headerMap, headerName) {
  var index = headerMap[core_normalizeHeader_(headerName)];
  return typeof index === 'number' ? String(row[index] || '').trim() : '';
}

function core_isValidGeapaConfigDateText_(value) {
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return true;
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(value)) return true;
  return !isNaN(Date.parse(value));
}

function core_validateGeapaConfigValue_(key, value, type, rowNumber) {
  var normalizedType = core_normalizeGeapaConfigKey_(type || 'TEXTO');
  if (!value) return;

  if (normalizedType === 'EMAIL' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)) {
    throw new Error('Config_GEAPA invalida na linha ' + rowNumber + ' (KEY=' + key + '): valor nao parece EMAIL.');
  }

  if (normalizedType === 'URL' && !/^https?:\/\/\S+$/i.test(value)) {
    throw new Error('Config_GEAPA invalida na linha ' + rowNumber + ' (KEY=' + key + '): valor nao parece URL http/https.');
  }

  if (normalizedType === 'DATA' && !core_isValidGeapaConfigDateText_(value)) {
    throw new Error('Config_GEAPA invalida na linha ' + rowNumber + ' (KEY=' + key + '): valor nao parece DATA reconhecivel.');
  }

  if (normalizedType === 'COR_HEX' && !/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(value)) {
    throw new Error('Config_GEAPA invalida na linha ' + rowNumber + ' (KEY=' + key + '): valor nao parece COR_HEX.');
  }
}

function core_readGeapaConfigVertical_(table, opts) {
  opts = opts || {};
  var includeInactive = opts.includeInactive === true;
  var map = {};
  var meta = {};
  var ignoredInactive = 0;
  var duplicateKeys = [];

  for (var i = 0; i < table.rows.length; i++) {
    var row = table.rows[i];
    var rawKey = core_getGeapaConfigCell_(row, table.headerMap, 'KEY');
    if (!rawKey) continue;

    var key = core_normalizeGeapaConfigKey_(rawKey);
    var active = core_isGeapaConfigActive_(core_getGeapaConfigCell_(row, table.headerMap, 'ATIVO'));
    if (!includeInactive && !active) {
      ignoredInactive++;
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(map, key)) {
      duplicateKeys.push(key);
      continue;
    }

    var value = core_getGeapaConfigCell_(row, table.headerMap, 'VALOR');
    var type = core_getGeapaConfigCell_(row, table.headerMap, 'TIPO') || 'TEXTO';
    var rowNumber = table.headerRow + 1 + i;
    core_validateGeapaConfigValue_(key, value, type, rowNumber);

    map[key] = value;
    meta[key] = {
      key: key,
      originalKey: rawKey,
      value: value,
      type: core_normalizeGeapaConfigKey_(type || 'TEXTO'),
      group: core_getGeapaConfigCell_(row, table.headerMap, 'GRUPO'),
      active: active,
      description: core_getGeapaConfigCell_(row, table.headerMap, 'DESCRICAO'),
      observations: core_getGeapaConfigCell_(row, table.headerMap, 'OBSERVACOES'),
      rowNumber: rowNumber
    };
  }

  return {
    format: 'VERTICAL',
    map: map,
    meta: meta,
    ignoredInactive: ignoredInactive,
    duplicateKeys: duplicateKeys
  };
}

function core_readGeapaConfigHorizontalLegacy_(table) {
  var map = {};
  var meta = {};
  var values = table.rows.length ? table.rows[0] : [];

  for (var i = 0; i < table.headers.length; i++) {
    var rawKey = String(table.headers[i] || '').trim();
    if (!rawKey) continue;

    var key = core_normalizeGeapaConfigKey_(rawKey);
    var value = String(values[i] || '').trim();
    map[key] = value;
    meta[key] = {
      key: key,
      originalKey: rawKey,
      value: value,
      type: 'TEXTO',
      group: '',
      active: true,
      description: '',
      observations: 'Fallback temporario do formato horizontal legado.',
      rowNumber: table.headerRow + 1
    };
  }

  return {
    format: 'HORIZONTAL_LEGACY',
    map: map,
    meta: meta,
    ignoredInactive: 0,
    duplicateKeys: []
  };
}

function core_readGeapaConfig_(opts) {
  opts = opts || {};
  var source = core_getGeapaConfigSheet_();
  var table = core_readGeapaConfigDisplayTable_(source.sheet, opts);
  var parsed = core_isGeapaConfigVerticalTable_(table)
    ? core_readGeapaConfigVertical_(table, opts)
    : core_readGeapaConfigHorizontalLegacy_(table);

  parsed.registryKey = source.registryKey;
  parsed.source = source.source;
  parsed.sheetName = source.sheet.getName();
  parsed.expectedHeaders = CORE_GEAPA_CONFIG_CFG.expectedHeaders.slice();
  return parsed;
}

function core_getGeapaConfigMap_(opts) {
  return core_readGeapaConfig_(opts || {}).map;
}

function core_getGeapaConfigObject_(opts) {
  return core_getGeapaConfigMap_(opts || {});
}

function core_getGeapaConfigValue_(key, opts) {
  core_assertRequired_(key, 'Config_GEAPA KEY');
  if (arguments.length < 2 || opts == null) {
    opts = {};
  } else if (typeof opts !== 'object') {
    opts = { defaultValue: opts };
  }

  var normalizedKey = core_normalizeGeapaConfigKey_(key);
  var map = core_getGeapaConfigMap_(opts);
  if (Object.prototype.hasOwnProperty.call(map, normalizedKey)) {
    return map[normalizedKey];
  }

  if (Object.prototype.hasOwnProperty.call(opts, 'defaultValue')) {
    return opts.defaultValue;
  }

  if (opts.required === true) {
    throw new Error('Config_GEAPA obrigatoria nao encontrada ou inativa: ' + normalizedKey);
  }

  return '';
}

function core_debugGeapaConfig_(opts) {
  var parsed = core_readGeapaConfig_(opts || {});
  return {
    ok: true,
    format: parsed.format,
    registryKey: parsed.registryKey,
    source: parsed.source,
    sheetName: parsed.sheetName,
    keys: Object.keys(parsed.map).sort(),
    totalKeys: Object.keys(parsed.map).length,
    ignoredInactive: parsed.ignoredInactive,
    duplicateKeys: parsed.duplicateKeys,
    expectedHeaders: parsed.expectedHeaders,
    meta: parsed.meta
  };
}
