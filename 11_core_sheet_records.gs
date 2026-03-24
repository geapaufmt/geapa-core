/**
 * Helpers de mais alto nível para leitura tabular orientada a cabeçalhos.
 *
 * Objetivo:
 * - reduzir repetição nos módulos
 * - padronizar leitura por objeto em vez de arrays posicionais
 * - facilitar buscas simples por header/valor
 */

/**
 * Lê todas as linhas de uma aba como objetos {header: valor}.
 *
 * opts:
 * - headerRow (default 1)
 * - startRow (default headerRow + 1)
 * - skipBlankRows (default true)
 */
function core_readSheetRecords_(sheet, opts) {
  core_assertRequired_(sheet, 'Sheet');
  opts = opts || {};

  const headerRow = Number(opts.headerRow || 1);
  const startRow = Number(opts.startRow || (headerRow + 1));
  const skipBlankRows = opts.skipBlankRows !== false;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < startRow || lastCol < 1) return [];

  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });

  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();
  const out = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var isBlank = row.every(function(cell) {
      return cell === '' || cell === null;
    });
    if (skipBlankRows && isBlank) continue;

    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      obj[headers[c]] = row[c];
    }
    obj.__rowNumber = startRow + i;
    out.push(obj);
  }

  return out;
}

/**
 * Lê registros diretamente por KEY do Registry.
 */
function core_readRecordsByKey_(key, opts) {
  var sh = core_getSheetByKey_(key);
  return core_readSheetRecords_(sh, opts);
}

/**
 * Busca o primeiro registro por igualdade simples em um header.
 * Comparação textual, trim e case-insensitive por padrão.
 */
function core_findFirstRecordByField_(records, headerName, value, opts) {
  core_assertRequired_(records, 'records');
  core_assertRequired_(headerName, 'headerName');
  opts = opts || {};

  var normalize = opts.normalize !== false;
  var target = normalize
    ? String(value || '').trim().toLowerCase()
    : value;

  for (var i = 0; i < records.length; i++) {
    var candidate = records[i][headerName];
    var current = normalize
      ? String(candidate || '').trim().toLowerCase()
      : candidate;
    if (current === target) return records[i];
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
