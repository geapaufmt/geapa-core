/**
 * Exports publicos para helpers tabulares orientados a cabecalhos.
 *
 * Objetivo:
 * - dar ao Core uma camada mais reutilizavel para todos os modulos;
 * - reduzir repeticao de mapeamento linha-array -> objeto;
 * - centralizar escrita por payload respeitando a ordem da sheet.
 */

function coreRowToObject(headers, row) {
  return core_rowToObject_(headers, row);
}

function coreBuildRowFromObjectByHeaders(headers, payload) {
  return core_buildRowFromObjectByHeaders_(headers, payload || {});
}

function coreAppendObjectByHeaders(sheet, payload, opts) {
  return core_appendObjectByHeaders_(sheet, payload || {}, opts || {});
}

function coreReadSheetData(sheet, opts) {
  return core_readSheetData_(sheet, opts || {});
}

function coreReadSheetRecords(sheet, opts) {
  return core_readSheetRecords_(sheet, opts || {});
}

function coreReadRecordsByKey(key, opts) {
  return core_readRecordsByKey_(key, opts || {});
}

function coreFindFirstRecordByField(records, headerName, value, opts) {
  return core_findFirstRecordByField_(records, headerName, value, opts || {});
}

function coreFindFirstRecordByAnyField(records, headerNames, value, opts) {
  return core_findFirstRecordByAnyField_(records, headerNames, value, opts || {});
}

function coreFindFirstRecordByKeyField(key, headerName, value, opts) {
  return core_findFirstRecordByKeyField_(key, headerName, value, opts || {});
}

function coreGetNearestFilledValueUp(sheet, startRow, colNumber) {
  return core_getNearestFilledValueUp_(sheet, startRow, colNumber);
}
