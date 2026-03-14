/**
 * ============================================================
 * 20_public_exports.gs
 * ============================================================
 *
 * API PÚBLICA do GEAPA-CORE como Library.
 *
 * IMPORTANTE:
 * - Apps Script Libraries exportam FUNÇÕES GLOBAIS.
 * - Objetos/constantes (ex.: const GEAPA_CORE = {...}) NÃO são exportados.
 *
 * Portanto, aqui existem apenas wrappers:
 *   function coreXxx(...) { return core_xxx_(...) }
 *
 * Exemplo no módulo:
 *   GEAPA_CORE.coreGetRegistry()
 *   GEAPA_CORE.coreGetSheetByKey("MEMBERS_ATUAIS")
 *   GEAPA_CORE.coreSendEmailHtml({...})
 */

/* ============================================================
 * REGISTRY
 * ============================================================ */

/** Retorna o registry inteiro (lido da planilha-mãe e cacheado). */
function coreGetRegistry() {
  return core_getRegistry_();
}

/**
 * Retorna a referência {id, sheet} de uma KEY do registry.
 * Ex.: coreGetRegistryRefByKey("MEMBERS_ATUAIS") -> {id, sheet}
 */
function coreGetRegistryRefByKey(key) {
  // ajuste o nome interno conforme você corrigir no 10_core_registry.gs:
  // recomendado: core_getRegistryRefByKey_(key)
  return core_getRegistryRefByKey_(key);
}

/**
 * Abre e retorna o Sheet diretamente via KEY do registry.
 * Ex.: coreGetSheetByKey("MEMBERS_ATUAIS") -> Sheet
 */
function coreGetSheetByKey(key) {
  return core_getSheetByKey_(key);
}

/* ============================================================
 * SHEETS
 * ============================================================ */

function coreOpenSpreadsheetById(id) {
  return core_openSpreadsheetById_(id);
}

function coreGetSheetById(spreadsheetId, sheetName) {
  return core_getSheetById_(spreadsheetId, sheetName);
}

function coreHeaderMap(sheet, headerRow) {
  return core_headerMap_(sheet, headerRow || 1);
}

function coreGetCol(headerMap, headerName) {
  return core_getCol_(headerMap, headerName);
}

function coreNormalizeHeader(s) {
  return core_normalizeHeader_(s);
}

/* ============================================================
 * DATES
 * ============================================================ */

function coreNow() {
  return core_now_();
}

function coreStartOfDay(date) {
  return core_startOfDay_(date);
}

function coreAddDays(date, days) {
  return core_addDays_(date, days);
}

function coreIsSameDay(d1, d2) {
  return core_isSameDay_(d1, d2);
}

function coreInWindowDay(date, startInclusive, endExclusive) {
  return core_inWindowDay_(date, startInclusive, endExclusive);
}

function coreFormatDate(date, tz, pattern) {
  return core_formatDate_(date, tz, pattern);
}

/* ============================================================
 * EMAIL
 * ============================================================ */

function coreIsValidEmail(email) {
  return core_isValidEmail_(email);
}

function coreSendEmailText(opts) {
  return core_sendEmailText_(opts);
}

function coreSendEmailHtml(opts) {
  return core_sendEmailHtml_(opts);
}

/* ============================================================
 * GMAIL (labels)
 * ============================================================ */

function coreEnsureLabel(name) {
  return core_ensureLabel_(name);
}

function coreGetLabel(name) {
  return core_getLabel_(name);
}

/* ============================================================
 * LOGS
 * ============================================================ */

function coreRunId() {
  return core_runId_();
}

function coreLogInfo(runId, msg, obj) {
  return core_logInfo_(runId, msg, obj);
}

function coreLogWarn(runId, msg, obj) {
  return core_logWarn_(runId, msg, obj);
}

function coreLogError(runId, msg, obj) {
  return core_logError_(runId, msg, obj);
}

function coreLogSummarize(runId, title, startedAt, counters) {
  return core_logSummarize_(runId, title, startedAt, counters || {});
}

/* ============================================================
 * ASSERT
 * ============================================================ */

function coreAssertRequired(value, msg) {
  return core_assertRequired_(value, msg);
}

/* ============================================================
 * DRIVE
 * ============================================================ */

function coreDriveGetFolderById(folderId) {
  return core_driveGetFolderById_(folderId);
}

function coreDriveGetFileById(fileId) {
  return core_driveGetFileById_(fileId);
}

function coreDriveEnsureFolder(parentFolderId, childName) {
  return core_driveEnsureFolder_(parentFolderId, childName);
}

function coreDriveMoveFileToFolder(fileId, folderId) {
  return core_driveMoveFileToFolder_(fileId, folderId);
}

function coreDriveListFiles(folderId, max) {
  return core_driveListFiles_(folderId, max || 200);
}

/* ============================================================
 * HTTP
 * ============================================================ */

function coreHttpPostJson(opts) {
  return core_httpPostJson_(opts);
}

/* ============================================================
 * ASSETS / EMAIL COM IMAGENS
 * ============================================================ */

function coreGetAssetBlob(assetIdOrUrl) {
  return coreGetAssetBlob_(assetIdOrUrl);
}

function coreInlineImagesDefault() {
  return coreInlineImagesDefault_();
}

function coreSendHtmlEmail(opts) {
  // envia HTML + logo padrão + inlineImages adicionais do módulo
  return core_sendEmailHtmlWithDefaultInline_(opts);
}

/* ============================================================
 * LOCK (Library limitation)
 * ============================================================ *
 * Callback NÃO atravessa library, então não exportamos execução com fn().
 * Módulos devem usar LockService localmente.
 */
function coreWithLock() {
  throw new Error("coreWithLock não suportado via Library. Use LockService no módulo.");
}

/* ============================================================
 * GOVERNANCE / VIGÊNCIAS
 * ============================================================ */
function coreGetCurrentBoard(refDate) {
  return core_getCurrentBoard_(refDate);
}

function coreGetCurrentBoardMembers(refDate) {
  return core_getCurrentBoardMembers_(refDate);
}

function coreGetCurrentBoardMembersByRole(role, refDate) {
  return core_getCurrentBoardMembersByRole_(role, refDate);
}

function coreGetCurrentBoardMemberByRole(role, refDate) {
  return core_getCurrentBoardMemberByRole_(role, refDate);
}

function coreGetCurrentLeadership(refDate) {
  return core_getCurrentLeadership_(refDate);
}

/* ============================================================
 * SEMESTRES / RGA
 * ============================================================ */
function coreGetCurrentSemester(refDate) {
  return core_getCurrentSemester_(refDate);
}

function coreParseEntrySemesterFromRga(rga) {
  return core_parseEntrySemesterFromRga_(rga);
}

function coreGetStudentCurrentSemesterFromRga(rga, refDate) {
  return core_getStudentCurrentSemesterFromRga_(rga, refDate);
}

/* ============================================================
 * SEMESTRES / MEMBERS_ATUAIS
 * ============================================================ */
function coreGetSemesterForDate(refDate) {
  return core_getSemesterForDate_(refDate);
}

function coreGetLastCompletedSemester(refDate) {
  return core_getLastCompletedSemester_(refDate);
}

function coreGetCompletedGroupSemesterCountFromEntrySemester(entrySemesterShort, refDate) {
  return core_getCompletedGroupSemesterCountFromEntrySemester_(entrySemesterShort, refDate);
}

function coreSyncMembersCurrentDerivedFields() {
  return core_syncMembersCurrentDerivedFields_();
}