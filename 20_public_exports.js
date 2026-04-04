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

function coreGetCurrentEnv() {
  return core_getCurrentEnv_();
}

function coreGetRegistryMetaByKey(key) {
  return core_getRegistryMetaByKey_(key);
}

/* ============================================================
 * REGISTRY / DEBUG
 * ============================================================ */

function coreClearRegistryCache() {
  return core_registryCacheClear_();
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

function coreNormalizeText(value, opts) {
  return core_normalizeText_(value, opts || {});
}

function coreOnlyDigits(value) {
  return core_onlyDigits_(value);
}

function coreBuildHeaderIndexMap(headers, opts) {
  return core_buildHeaderIndexMap_(headers, opts || {});
}

function coreFindHeaderIndex(headerMap, headerName, opts) {
  return core_findHeaderIndex_(headerMap, headerName, opts || {});
}

function coreSetRowValueByHeader(rowArr, headerMap, headerName, value, opts) {
  return core_setRowValueByHeader_(rowArr, headerMap, headerName, value, opts || {});
}

function coreGetCellByHeader(rowArr, headerMap, headerName, opts) {
  return core_getCellByHeader_(rowArr, headerMap, headerName, opts || {});
}

function coreFindFirstExistingHeader(headerMap, headerNames, opts) {
  return core_findFirstExistingHeader_(headerMap, headerNames, opts || {});
}

function coreWriteCellByHeader(sheet, rowNumber, headerMap, headerName, value, opts) {
  return core_writeCellByHeader_(sheet, rowNumber, headerMap, headerName, value, opts || {});
}

function coreFreezeHeaderRow(sheet, headerRow) {
  return core_freezeHeaderRow_(sheet, headerRow || 1);
}

function coreEnsureFilter(sheet, headerRow, opts) {
  return core_ensureFilter_(sheet, headerRow || 1, opts || {});
}

function coreApplyHeaderNotes(sheet, notesByHeader, headerRow) {
  return core_applyHeaderNotes_(sheet, notesByHeader || {}, headerRow || 1);
}

function coreApplyHeaderColors(sheet, groups, headerRow, opts) {
  return core_applyHeaderColors_(sheet, groups || [], headerRow || 1, opts || {});
}

function coreApplyDropdownValidationByHeader(sheet, rulesByHeader, headerRow, opts) {
  return core_applyDropdownValidationByHeader_(sheet, rulesByHeader || {}, headerRow || 1, opts || {});
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

function coreNormalizeEmail(value) {
  return core_normalizeEmail_(value);
}

function coreSendEmailText(opts) {
  return core_sendEmailText_(opts);
}

function coreSendEmailHtml(opts) {
  return core_sendEmailHtml_(opts);
}

function coreSendTrackedEmail(params) {
  return core_sendTrackedEmail_(params);
}

function coreExtractEmailAddress(value) {
  return core_extractEmailAddress_(value);
}

function coreExtractDisplayName(value) {
  return core_extractDisplayName_(value);
}

function coreUniqueEmails(values) {
  return core_uniqueEmails_(values);
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

function coreGetOrCreateLabel(name) {
  return core_getOrCreateLabel_(name);
}

function coreThreadHasLabel(thread, labelName) {
  return core_threadHasLabel_(thread, labelName);
}

function coreSearchThreads(query, start, max) {
  return core_searchThreads_(query, start, max);
}

function coreMarkThread(thread, labelIn, labelOut) {
  return core_markThread_(thread, labelIn, labelOut);
}

function coreReplyThreadHtml(thread, subject, htmlBody, opts) {
  return core_replyThreadHtml_(thread, subject, htmlBody, opts || {});
}

/* ============================================================
 * MAIL HUB
 * ============================================================ */

function coreMailIngestInbox(opts) {
  return core_mailIngestInbox_(opts || {});
}

function coreMailGetConfig(key, defaultValue) {
  return coreMailHubGetConfig_(key, defaultValue);
}

function coreMailGetConfigBoolean(key, defaultValue) {
  return coreMailHubGetConfigBooleanByKey_(key, defaultValue === true);
}

function coreMailGetConfigList(key) {
  return coreMailHubGetConfigListByKey_(key);
}

function coreMailListPendingByModule(moduleName) {
  return core_mailListPendingByModule_(moduleName);
}

function coreMailRegisterModuleAdapter(adapter) {
  return coreMailRegisterModuleAdapter_(adapter);
}

function coreMailGetModuleAdapter(moduleCodeOrName) {
  return coreMailAdapterToSnapshot_(coreMailGetModuleAdapter_(moduleCodeOrName));
}

function coreMailListModuleAdapters() {
  return coreMailListModuleAdapters_().map(function(adapter) {
    return coreMailAdapterToSnapshot_(adapter);
  });
}

function coreMailBuildCorrelationKey(moduleCodeOrName, ctx) {
  return coreMailBuildCorrelationKey_(moduleCodeOrName, ctx || {});
}

function coreMailParseCorrelationKey(key) {
  return coreMailParseCorrelationKey_(key);
}

function coreMailResolveRouting(msgCtx) {
  return coreMailResolveRouting_(msgCtx || {});
}

function coreMailNormalizeOutgoingSubject(moduleCodeOrName, subject, ctx) {
  return coreMailNormalizeOutgoingSubject_(moduleCodeOrName, subject, ctx || {});
}

function coreMailRenderEmailTemplate(templateKey, subjectHuman, payload) {
  return coreMailRenderEmailTemplate_(templateKey, subjectHuman, payload || {});
}

function coreMailBuildFinalSubject(subjectHuman, correlationKey) {
  return coreMailBuildFinalSubject_(subjectHuman, correlationKey);
}

function coreMailBuildOutgoingDraft(contract) {
  return coreMailBuildOutgoingDraft_(contract || {});
}

function coreMailQueueOutgoing(contract) {
  return coreMailQueueOutgoing_(contract || {});
}

function coreMailProcessOutbox() {
  return coreMailProcessOutbox_();
}

function coreMailGetLatestEvent(opts) {
  return core_mailGetLatestEvent_(opts || {});
}

function coreMailMarkLatestPendingByModule(moduleName, processorName) {
  return core_mailMarkLatestPendingByModule_(moduleName, processorName);
}

function coreMailMarkEventProcessed(eventId, processorName) {
  return core_mailMarkEventProcessed_(eventId, processorName);
}

function coreMailCleanupNoiseEvents() {
  return coreMailCleanupNoiseEvents_();
}

function coreMailApplyOperationalSheetUx(opts) {
  return coreMailApplyOperationalSheetUx_(opts || {});
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
  return core_inlineImagesDefault_();
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

function coreGetCurrentBoardSlogan(refDate) {
  return core_getCurrentBoardSlogan_(refDate);
}

function coreGetCurrentBoardMembers(refDate) {
  return core_getCurrentBoardMembers_(refDate);
}

function coreGetCurrentBoardMembersByRole(role, refDate) {
  return core_getCurrentBoardMembersByRole_(role, refDate);
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

function coreGetSemesterIdForDate(refDate) {
  return core_getSemesterIdForDate_(refDate);
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


/* ============================================================
 * CARGOS INSTITUCIONAIS / CONFIG
 * ============================================================ */

function coreGetInstitutionalRolesActive() {
  return core_getInstitutionalRolesActive_();
}

function coreFindInstitutionalRoleByKey(cargoKey) {
  return core_findInstitutionalRoleByKey_(cargoKey);
}

function coreFindInstitutionalRoleByPublicName(publicName) {
  return core_findInstitutionalRoleByPublicName_(publicName);
}

function coreFindInstitutionalRoleByAnyName(text) {
  return core_findInstitutionalRoleByAnyName_(text);
}

function coreGetInstitutionalRolesByEmailGroup(groupName) {
  return core_getInstitutionalRolesByEmailGroup_(groupName);
}

function coreDebugInstitutionalRolesConfig() {
  return core_debugInstitutionalRolesConfig_();
}

function coreClearInstitutionalRolesConfigCache() {
  return core_rolesConfigCacheClear_();
}

/* ============================================================
 * CARGOS INSTITUCIONAIS / PROJEÇÃO ATUAL
 * ============================================================ */

function coreGetCurrentInstitutionalAssignments(refDate) {
  return core_getCurrentInstitutionalAssignments_(refDate);
}

function coreGetCurrentOccupantsByEmailGroup(groupName, refDate) {
  return core_getCurrentOccupantsByEmailGroup_(groupName, refDate);
}

function coreGetCurrentContactsHtmlByEmailGroup(groupName, refDate) {
  return core_getCurrentContactsHtmlByEmailGroup_(groupName, refDate);
}

function coreSyncMembersCurrentInstitutionalRoles(refDate) {
  return core_syncMembersCurrentInstitutionalRoles_(refDate);
}

function coreDebugCurrentInstitutionalProjection(refDate) {
  return core_debugCurrentInstitutionalProjection_(refDate);
}

function coreGetCurrentEmailsByEmailGroup(groupName, refDate) {
  return core_getCurrentEmailsByEmailGroup_(groupName, refDate);
}

function coreGetCurrentEmailsByRole(roleName, refDate) {
  return core_getCurrentEmailsByRole_(roleName, refDate);
}

/* ============================================================
 * IDENTIDADE DE MEMBRO / AUTOFILL
 * ============================================================ */

function coreFindMemberIdentityByAny(identity) {
  return core_memberIdentityFindByAny_(identity);
}

function coreNormalizeIdentityKey(value) {
  return core_normalizeIdentityKey_(value);
}

function coreFindMemberCurrentRowByAny(identity) {
  return core_findMemberCurrentRowByAny_(identity);
}

function coreAutofillIdentityRowInSheet(sheet, rowNumber, opts) {
  return core_autofillIdentityRowInSheet_(sheet, rowNumber, opts || {});
}
