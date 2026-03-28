/**************************************
 * 01_core_sheets.gs (refatorado)
 *
 * Objetivo:
 * - Centralizar operacoes comuns de Sheets (open / get sheet / header map)
 * - Garantir falhas explicitas (erros claros) quando algo essencial estiver faltando
 * - Evitar repeticao e reduzir custo de abrir planilhas repetidamente
 **************************************/

/**
 * Cache por execucao (in-memory).
 * - Evita abrir a mesma planilha N vezes numa unica execucao.
 * - Nao "persiste" entre execucoes (o que e bom: sempre atualiza no proximo run).
 */
const __core_ss_cache = Object.create(null);

function core_normalizeText_(value, opts) {
  opts = opts || {};

  var text = String(value == null ? '' : value).trim();
  if (!text) return '';

  if (opts.removeAccents === true) {
    text = text
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  }

  if (opts.collapseWhitespace !== false) {
    text = text.replace(/\s+/g, ' ');
  }

  if (opts.caseMode === 'upper') {
    text = text.toUpperCase();
  } else if (opts.caseMode !== 'none') {
    text = text.toLowerCase();
  }

  return text;
}

function core_onlyDigits_(value) {
  return String(value == null ? '' : value).replace(/\D+/g, '');
}

function core_buildHeaderIndexMap_(headers, opts) {
  opts = opts || {};

  var normalize = opts.normalize !== false;
  var oneBased = opts.oneBased === true;
  var keepFirst = opts.keepFirst !== false;
  var map = Object.create(null);

  headers.forEach(function(header, index) {
    var raw = String(header || '').trim();
    if (!raw) return;

    var key = normalize ? core_normalizeHeader_(raw) : raw;
    if (!key) return;
    if (keepFirst && Object.prototype.hasOwnProperty.call(map, key)) return;

    map[key] = oneBased ? index + 1 : index;
  });

  return Object.freeze(map);
}

function core_findHeaderIndex_(headerMap, headerName, opts) {
  core_assertRequired_(headerMap, 'headerMap');
  opts = opts || {};

  var key = opts.normalize === false
    ? String(headerName || '').trim()
    : core_normalizeHeader_(headerName);

  if (!key || !Object.prototype.hasOwnProperty.call(headerMap, key)) {
    return opts.notFoundValue != null ? opts.notFoundValue : -1;
  }

  return headerMap[key];
}

function core_setRowValueByHeader_(rowArr, headerMap, headerName, value, opts) {
  core_assertRequired_(rowArr, 'rowArr');
  core_assertRequired_(headerMap, 'headerMap');

  var idx = core_findHeaderIndex_(headerMap, headerName, {
    normalize: !(opts && opts.normalize === false),
    notFoundValue: -1
  });

  if (idx < 0) return false;
  rowArr[idx] = value;
  return true;
}

function core_getCellByHeader_(rowArr, headerMap, headerName, opts) {
  core_assertRequired_(rowArr, 'rowArr');
  core_assertRequired_(headerMap, 'headerMap');

  var idx = core_findHeaderIndex_(headerMap, headerName, {
    normalize: !(opts && opts.normalize === false),
    notFoundValue: -1
  });

  if (idx < 0) {
    return opts && Object.prototype.hasOwnProperty.call(opts, 'defaultValue')
      ? opts.defaultValue
      : '';
  }

  return rowArr[idx];
}

function core_findFirstExistingHeader_(headerMap, headerNames, opts) {
  core_assertRequired_(headerMap, 'headerMap');
  core_assertRequired_(headerNames, 'headerNames');

  var names = Array.isArray(headerNames) ? headerNames : [headerNames];

  for (var i = 0; i < names.length; i++) {
    var idx = core_findHeaderIndex_(headerMap, names[i], {
      normalize: !(opts && opts.normalize === false),
      notFoundValue: -1
    });

    if (idx >= 0) {
      return {
        found: true,
        headerName: names[i],
        index: idx
      };
    }
  }

  return {
    found: false,
    headerName: '',
    index: opts && Object.prototype.hasOwnProperty.call(opts, 'notFoundValue')
      ? opts.notFoundValue
      : -1
  };
}

function core_writeCellByHeader_(sheet, rowNumber, headerMap, headerName, value, opts) {
  core_assertRequired_(sheet, 'sheet');
  core_assertRequired_(rowNumber, 'rowNumber');
  core_assertRequired_(headerMap, 'headerMap');

  opts = opts || {};
  var idx = core_findHeaderIndex_(headerMap, headerName, {
    normalize: opts.normalize !== false,
    notFoundValue: -1
  });

  if (idx < 0) return false;

  var colNumber = opts.oneBased === true ? idx : idx + 1;
  sheet.getRange(Number(rowNumber), colNumber).setValue(value);
  return true;
}

/**
 * ------------------------------------------------------------
 * Abre uma planilha pelo ID (com validacao + cache).
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Sempre que precisar de SpreadsheetApp.openById(id).
 *
 * Por que existe:
 * - Valida ID obrigatorio.
 * - Cacheia o Spreadsheet para reduzir chamadas repetidas.
 */
function core_openSpreadsheetById_(id) {
  core_assertRequired_(id, 'Spreadsheet ID');
  var key = String(id).trim();

  if (!__core_ss_cache[key]) {
    __core_ss_cache[key] = SpreadsheetApp.openById(key);
  }
  return __core_ss_cache[key];
}

/**
 * ------------------------------------------------------------
 * Retorna uma aba (Sheet) pelo (spreadsheetId + sheetName).
 * ------------------------------------------------------------
 *
 * Comportamento:
 * - Valida parametros obrigatorios.
 * - Lanca erro explicito se a aba nao existir.
 */
function core_getSheetById_(spreadsheetId, sheetName) {
  core_assertRequired_(spreadsheetId, 'Spreadsheet ID');
  core_assertRequired_(sheetName, 'Sheet name');

  var ss = core_openSpreadsheetById_(spreadsheetId);
  var sh = ss.getSheetByName(String(sheetName).trim());

  if (!sh) {
    throw new Error('Aba nao encontrada: "' + sheetName + '" no arquivo ' + spreadsheetId);
  }
  return sh;
}

/**
 * ------------------------------------------------------------
 * Cria mapa { headerNormalizado -> coluna(1-based) }.
 * ------------------------------------------------------------
 *
 * Quando usar:
 * - Para buscar colunas pelo nome do cabecalho, sem depender de indice fixo.
 *
 * Notas:
 * - Mantem o primeiro cabecalho em caso de duplicidade.
 * - Colunas retornam 1-based (compativel com getRange()).
 */
function core_headerMap_(sheet, headerRow) {
  core_assertRequired_(sheet, 'Sheet');

  var row = Number(headerRow || 1);
  if (row < 1) throw new Error('headerRow invalido: ' + headerRow);

  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return Object.freeze({});

  var headers = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  return core_buildHeaderIndexMap_(headers, {
    normalize: true,
    oneBased: true,
    keepFirst: true
  });
}

/**
 * ------------------------------------------------------------
 * Retorna a coluna (1-based) associada ao headerName.
 * ------------------------------------------------------------
 *
 * Retorna:
 * - coluna 1-based, se existir
 * - 0, se nao existir (o modulo decide se e obrigatorio ou opcional)
 */
function core_getCol_(headerMap, headerName) {
  core_assertRequired_(headerMap, 'headerMap');
  var idx = core_findHeaderIndex_(headerMap, headerName, {
    normalize: true,
    notFoundValue: -1
  });
  return idx >= 0 ? idx : 0;
}

/**
 * ------------------------------------------------------------
 * Normaliza texto de cabecalho para comparacao robusta.
 * ------------------------------------------------------------
 */
function core_normalizeHeader_(h) {
  return core_normalizeText_(h, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'lower'
  });
}
