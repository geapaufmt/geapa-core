/**
 * Exports públicos para helpers tabulares orientados a cabeçalhos.
 *
 * Objetivo:
 * - dar ao Core uma camada mais reutilizável para todos os módulos
 * - reduzir repetição de mapeamento linha-array -> objeto
 */

function coreReadSheetRecords(sheet, opts) {
  return core_readSheetRecords_(sheet, opts || {});
}

function coreReadRecordsByKey(key, opts) {
  return core_readRecordsByKey_(key, opts || {});
}

function coreFindFirstRecordByField(records, headerName, value, opts) {
  return core_findFirstRecordByField_(records, headerName, value, opts || {});
}

function coreFindFirstRecordByKeyField(key, headerName, value, opts) {
  return core_findFirstRecordByKeyField_(key, headerName, value, opts || {});
}
