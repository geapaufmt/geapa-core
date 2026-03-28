/**
 * Helpers de mais alto nivel para leitura tabular orientada a cabecalhos.
 *
 * Objetivo:
 * - reduzir repeticao nos modulos;
 * - padronizar leitura por objeto em vez de arrays posicionais;
 * - facilitar buscas simples por header/valor;
 * - permitir escrita por objeto respeitando a ordem real da planilha.
 */

function core_rowToObject_(headers, row) {
  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    obj[headers[i]] = row[i];
  }
  return obj;
}

function core_buildRowFromObjectByHeaders_(headers, payload) {
  core_assertRequired_(headers, 'headers');
  payload = payload || {};

  return headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(payload, header) ? payload[header] : '';
  });
}

function core_appendObjectByHeaders_(sheet, payload, opts) {
  core_assertRequired_(sheet, 'Sheet');
  opts = opts || {};

  var headerRow = Number(opts.headerRow || 1);
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    throw new Error('core_appendObjectByHeaders_: sheet sem cabecalhos.');
  }

  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });

  var row = core_buildRowFromObjectByHeaders_(headers, payload || {});
  sheet.appendRow(row);
  return row;
}

function core_readSheetData_(sheet, opts) {
  core_assertRequired_(sheet, 'Sheet');
  opts = opts || {};

  var headerRow = Number(opts.headerRow || 1);
  var startRow = Number(opts.startRow || (headerRow + 1));
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastCol < 1) {
    return {
      headers: [],
      rows: [],
      headerMap: {}
    };
  }

  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });

  var rows = lastRow >= startRow
    ? sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues()
    : [];

  return {
    headers: headers,
    rows: rows,
    headerMap: core_buildHeaderIndexMap_(headers, {
      normalize: opts.normalizeHeaderMap === true,
      oneBased: false,
      keepFirst: true
    })
  };
}

/**
 * Le todas as linhas de uma aba como objetos {header: valor}.
 *
 * opts:
 * - headerRow (default 1)
 * - startRow (default headerRow + 1)
 * - skipBlankRows (default true)
 */
function core_readSheetRecords_(sheet, opts) {
  core_assertRequired_(sheet, 'Sheet');
  opts = opts || {};

  var skipBlankRows = opts.skipBlankRows !== false;
  var data = core_readSheetData_(sheet, opts);
  var headerRow = Number((opts && opts.headerRow) || 1);
  var startRow = Number((opts && opts.startRow) || (headerRow + 1));
  var headers = data.headers;
  var values = data.rows;
  var out = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var isBlank = row.every(function(cell) {
      return cell === '' || cell === null;
    });
    if (skipBlankRows && isBlank) continue;

    var obj = core_rowToObject_(headers, row);
    obj.__rowNumber = startRow + i;
    out.push(obj);
  }

  return out;
}

/**
 * Le registros diretamente por KEY do Registry.
 */
function core_readRecordsByKey_(key, opts) {
  var sh = core_getSheetByKey_(key);
  return core_readSheetRecords_(sh, opts);
}

/**
 * Busca o primeiro registro por igualdade simples em um header.
 * Comparacao textual, trim e case-insensitive por padrao.
 */
function core_findFirstRecordByField_(records, headerName, value, opts) {
  core_assertRequired_(records, 'records');
  core_assertRequired_(headerName, 'headerName');
  opts = opts || {};

  var normalize = opts.normalize !== false;
  var normalizer = typeof opts.normalizer === 'function'
    ? opts.normalizer
    : function(input) {
        return core_normalizeText_(input, {
          collapseWhitespace: true,
          caseMode: 'lower'
        });
      };

  var target = normalize ? normalizer(value) : value;

  for (var i = 0; i < records.length; i++) {
    var candidate = records[i][headerName];
    var current = normalize ? normalizer(candidate) : candidate;
    if (current === target) return records[i];
  }

  return null;
}

function core_findFirstRecordByAnyField_(records, headerNames, value, opts) {
  core_assertRequired_(records, 'records');
  core_assertRequired_(headerNames, 'headerNames');

  for (var i = 0; i < headerNames.length; i++) {
    var found = core_findFirstRecordByField_(records, headerNames[i], value, opts || {});
    if (found) return found;
  }

  return null;
}

/**
 * Busca o primeiro registro diretamente por KEY do Registry.
 */
function core_findFirstRecordByKeyField_(key, headerName, value, opts) {
  var records = core_readRecordsByKey_(key, opts);
  return core_findFirstRecordByField_(records, headerName, value, opts);
}

function core_getNearestFilledValueUp_(sheet, startRow, colNumber) {
  core_assertRequired_(sheet, 'sheet');
  core_assertRequired_(startRow, 'startRow');
  core_assertRequired_(colNumber, 'colNumber');

  var firstDataRow = 2;
  var lastRow = Math.max(Number(startRow), firstDataRow);
  if (lastRow < firstDataRow) return '';

  var range = sheet.getRange(firstDataRow, Number(colNumber), lastRow - firstDataRow + 1, 1);
  var values = range.getValues();
  var displays = range.getDisplayValues();

  for (var i = values.length - 1; i >= 0; i--) {
    var value = values[i][0];
    var display = String(displays[i][0] || '').trim();

    if (value instanceof Date && !isNaN(value.getTime())) {
      return value;
    }

    if (value !== '' && value !== null && value !== undefined) {
      return value;
    }

    if (display) {
      return display;
    }
  }

  return '';
}
