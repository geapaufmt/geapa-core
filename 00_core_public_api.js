/**
 * ============================================================
 * 00_core_public_api.gs
 * ============================================================
 *
 * Este arquivo NÃO é o que a Library exporta para outros projetos.
 * Ele é um “mapa mental” interno do Core:
 * - organiza as funções do Core por áreas (Sheets, Dates, Email…)
 * - serve como documentação viva (o que existe e pra que serve)
 *
 * A API realmente exportada para módulos fica em:
 *   -> 20_public_exports.gs
 *
 * Regra geral:
 * - Funções internas terminam com "_" (ex.: core_now_).
 * - Módulos enxergam APENAS as funções wrapper exportadas no 20.
 *
 * Você pode usar este objeto aqui dentro do Core para navegar e
 * entender rapidamente o que existe, sem ficar caçando função em arquivo.
 */

const GEAPA_CORE = Object.freeze({
  /**
   * ----------------------------------------------------------
   * sheets (Planilhas)
   * ----------------------------------------------------------
   * Quando usar:
   * - abrir planilhas por ID
   * - acessar abas por nome de forma consistente
   * - mapear cabeçalhos -> colunas
   *
   * Por que existe:
   * - evitar "número mágico" de coluna
   * - evitar erro silencioso quando o nome do cabeçalho muda
   */
  sheets: Object.freeze({
    /** Abre Spreadsheet pelo ID (equivalente ao SpreadsheetApp.openById). */
    openSpreadsheetById: core_openSpreadsheetById_,

    /**
     * Retorna uma Sheet específica (Spreadsheet ID + nome da aba).
     * Deve falhar/lançar erro se não existir (melhor que retornar null).
     */
    getSheetById: core_getSheetById_,

    /**
     * Lê a linha de cabeçalho e devolve um mapa:
     * { "CABECALHO": coluna(1-based) }
     * Útil para localizar colunas por nome.
     */
    headerMap: core_headerMap_,

    /**
     * Busca uma coluna dentro do headerMap.
     * Normalmente você chama: getCol(map, "EMAIL")
     */
    getCol: core_getCol_,

    /**
     * Normaliza um cabeçalho (tirar espaços extras, padronizar caixa etc.)
     * Ajuda quando há pequenas variações no texto do cabeçalho.
     */
    normalizeHeader: core_normalizeHeader_,
    normalizeText: core_normalizeText_,
    onlyDigits: core_onlyDigits_,
    buildHeaderIndexMap: core_buildHeaderIndexMap_,
    findHeaderIndex: core_findHeaderIndex_,
    setRowValueByHeader: core_setRowValueByHeader_,
    writeCellByHeader: core_writeCellByHeader_,
    freezeHeaderRow: core_freezeHeaderRow_,
    ensureFilter: core_ensureFilter_,
    applyHeaderNotes: core_applyHeaderNotes_,
    applyHeaderColors: core_applyHeaderColors_,
    applyDropdownValidationByHeader: core_applyDropdownValidationByHeader_,
  }),

  records: Object.freeze({
    rowToObject: core_rowToObject_,
    buildRowFromObjectByHeaders: core_buildRowFromObjectByHeaders_,
    appendObjectByHeaders: core_appendObjectByHeaders_,
    readSheetRecords: core_readSheetRecords_,
    readRecordsByKey: core_readRecordsByKey_,
    findFirstRecordByField: core_findFirstRecordByField_,
    findFirstRecordByAnyField: core_findFirstRecordByAnyField_,
    findFirstRecordByKeyField: core_findFirstRecordByKeyField_,
  }),

  identity: Object.freeze({
    fillMissingProfessorIds: core_fillMissingProfessorIds_,
    fillMissingExternalIds: core_fillMissingExternalIds_,
    ensureProfessorIdForRow: core_ensureProfessorIdForRow_,
    ensureExternalIdForRow: core_ensureExternalIdForRow_,
    findExternalByEmail: core_identityFindExternalByEmail_,
    validateExternalEmailDuplicates: core_identityValidateExternalEmailDuplicates_,
  }),

  memberLifecycle: Object.freeze({
    appendEvent: core_appendMemberLifecycleEvent_,
    listEvents: core_memberLifecycleListEvents_,
    getLatestEventByRga: core_memberLifecycleGetLatestEventByRga_,
    updateEvent: core_updateMemberLifecycleEvent_,
    updateEventStatus: core_updateMemberLifecycleEventStatus_,
  }),

  /**
   * ----------------------------------------------------------
   * dates (Datas / Tempo)
   * ----------------------------------------------------------
   * Quando usar:
   * - comparar datas ignorando horário (mesmo dia)
   * - criar janelas: hoje, ontem, próximos dias etc.
   * - formatar datas para e-mail/planilha
   *
   * Por que existe:
   * - Apps Script mistura fuso/horário facilmente
   * - "Hoje" pode dar errado se você comparar Date com horário diferente
   */
  dates: Object.freeze({
    /** Retorna a data truncada para 00:00 (início do dia). */
    startOfDay: core_startOfDay_,

    /** Soma N dias em uma data. */
    addDays: core_addDays_,

    /** True se duas datas caem no mesmo dia (ignorando hora). */
    isSameDay: core_isSameDay_,

    /**
     * Verifica se uma data está dentro de uma janela [start, end)
     * (muito útil para "hoje" e "ontem", usando fuso correto).
     */
    inWindowDay: core_inWindowDay_,

    /** Retorna "agora" (centralizado para padronizar o sistema). */
    now: core_now_,

    /** Formata uma data (Utilities.formatDate). */
    format: core_formatDate_,
  }),

  /**
   * ----------------------------------------------------------
   * email (Envio de e-mails)
   * ----------------------------------------------------------
   * Quando usar:
   * - enviar mensagens simples texto/HTML
   * - validação básica de email
   *
   * Observação:
   * - e-mail com imagens inline (logo) geralmente usa o pacote "assets/email-assets"
   */
  email: Object.freeze({
    /** Validação simples de e-mail (evita to="" vazio / inválido). */
    isValid: core_isValidEmail_,

    /** Normaliza e-mail institucional para comparação/chaves. */
    normalizeEmail: core_normalizeEmail_,
    extractAddress: core_extractEmailAddress_,
    uniqueEmails: core_uniqueEmails_,

    /** Envia e-mail em texto puro. */
    sendText: core_sendEmailText_,

    /** Envia e-mail HTML (sem inlineImages). */
    sendHtml: core_sendHtmlEmail_,

    /** Envia e-mail e tenta retornar threadId/messageId do Gmail. */
    sendTracked: core_sendTrackedEmail_,
  }),

  /**
   * ----------------------------------------------------------
   * gmail (Gmail: labels e utilitários)
   * ----------------------------------------------------------
   * Quando usar:
   * - garantir/criar labels
   * - organizar threads processadas/erros
   *
   * Obs:
   * - leitura de e-mails e threads pode estar em outros arquivos do módulo,
   *   mas o Core deve fornecer helpers básicos para evitar duplicação.
   */
  gmail: Object.freeze({
    /** Cria ou retorna a label (GmailLabel). */
    ensureLabel: core_ensureLabel_,

    /** Retorna label existente ou null. */
    getLabel: core_getLabel_,
  }),

  mailHub: Object.freeze({
    ingestInbox: core_mailIngestInbox_,
    getConfig: coreMailHubGetConfig_,
    getConfigBoolean: coreMailHubGetConfigBooleanByKey_,
    getConfigList: coreMailHubGetConfigListByKey_,
    queueOutgoing: coreMailQueueOutgoing_,
    processOutbox: coreMailProcessOutbox_,
    listPendingByModule: core_mailListPendingByModule_,
    getLatestEvent: core_mailGetLatestEvent_,
    listAttachments: core_mailListAttachments_,
    listPendingAttachments: core_mailListPendingAttachments_,
    getLatestPendingEventWithAttachment: core_mailGetLatestPendingEventWithAttachment_,
    listAttachmentsByEvent: core_mailListAttachmentsByEvent_,
    getAttachmentById: core_mailGetAttachmentById_,
    getAttachmentsByEvent: core_mailGetAttachmentsByEvent_,
    markLatestPendingByModule: core_mailMarkLatestPendingByModule_,
    markEventProcessed: core_mailMarkEventProcessed_,
    markAttachmentProcessed: core_mailMarkAttachmentProcessed_,
    markAttachmentSavedToDrive: core_mailMarkAttachmentSavedToDrive_,
    markAttachmentIgnored: core_mailMarkAttachmentIgnored_,
    markAttachmentError: core_mailMarkAttachmentError_,
    cleanupNoiseEvents: coreMailCleanupNoiseEvents_,
  }),

  mailAdapters: Object.freeze({
    register: coreMailRegisterModuleAdapter_,
    get: coreMailGetModuleAdapter_,
    list: coreMailListModuleAdapters_,
    buildCorrelationKey: coreMailBuildCorrelationKey_,
    parseCorrelationKey: coreMailParseCorrelationKey_,
    resolveRouting: coreMailResolveRouting_,
    normalizeOutgoingSubject: coreMailNormalizeOutgoingSubject_,
  }),

  mailRenderer: Object.freeze({
    renderTemplate: coreMailRenderEmailTemplate_,
    buildFinalSubject: coreMailBuildFinalSubject_,
    buildOutgoingDraft: coreMailBuildOutgoingDraft_,
  }),

  /**
   * ----------------------------------------------------------
   * lock (Concorrência)
   * ----------------------------------------------------------
   * Quando usar:
   * - evitar rodar o mesmo job em paralelo via triggers
   *
   * Atenção:
   * - callbacks NÃO atravessam Library
   * - então isso é mais útil dentro do próprio Core ou dentro do módulo,
   *   mas não como API exportada para outro projeto.
   */
  lock: Object.freeze({
    withLock: core_withLock_,
  }),

  /**
   * ----------------------------------------------------------
   * log (Logs)
   * ----------------------------------------------------------
   * Quando usar:
   * - todo job com trigger deve ter logs consistentes
   * - usar runId para rastrear execução
   * - sumarizar contadores no fim ("enviados=..., pulos=..., erros=...")
   */
  log: Object.freeze({
    runId: core_runId_,
    info: core_logInfo_,
    warn: core_logWarn_,
    error: core_logError_,
    summarize: core_logSummarize_,
  }),

  /**
   * ----------------------------------------------------------
   * assert (Validações/erros bons)
   * ----------------------------------------------------------
   * Quando usar:
   * - validar entradas obrigatórias
   * - falhar cedo com mensagem clara
   */
  assert: Object.freeze({
    required: core_assertRequired_,
  }),

  /**
   * ----------------------------------------------------------
   * registry (Cadastro institucional)
   * ----------------------------------------------------------
   * Quando usar:
   * - obter IDs e nomes de abas institucionais compartilhadas
   *
   * Regra do registry:
   * - SOMENTE ponteiros (id, sheet)
   * - NÃO colocar templates, regex, regras de módulos aqui dentro
   */
  registry: Object.freeze({
    /** Retorna o registry atual (congelado). */
    get: core_getRegistry_,
  }),

  /**
   * ----------------------------------------------------------
   * modulesConfig (Controle operacional)
   * ----------------------------------------------------------
   * Quando usar:
   * - decidir se um modulo/fluxo pode executar
   * - respeitar modo operacional e capabilities
   *
   * Regra:
   * - MODULOS_CONFIG controla comportamento operacional.
   * - Registry continua sendo apenas resolucao de recursos.
   */
  modulesConfig: Object.freeze({
    get: core_getModuleConfig_,
    isEnabled: core_isModuleEnabled_,
    getMode: core_getModuleMode_,
    canUseCapability: core_canModuleUseCapability_,
    assertExecutionAllowed: core_assertModuleExecutionAllowed_,
    debug: core_debugModulesConfig_,
    applySheetUx: core_applyModulesConfigSheetUx_,
  }),

  /**
   * ----------------------------------------------------------
   * modulesStatus (Observabilidade operacional)
   * ----------------------------------------------------------
   * Quando usar:
   * - registrar inicio, sucesso, erro e bloqueio de fluxos
   * - consultar status operacional central por modulo/fluxo
   *
   * Regra:
   * - MODULOS_CONFIG decide permissao.
   * - MODULOS_STATUS registra o resultado operacional.
   */
  modulesStatus: Object.freeze({
    get: core_moduleStatusGet_,
    ensureRow: core_moduleStatusEnsureRow_,
    markExecution: core_moduleStatusMarkExecution_,
    markSuccess: core_moduleStatusMarkSuccess_,
    markError: core_moduleStatusMarkError_,
    markBlocked: core_moduleStatusMarkBlocked_,
    debug: core_debugModulesStatus_,
  }),

  /**
   * ----------------------------------------------------------
   * drive (Drive: arquivos/pastas)
   * ----------------------------------------------------------
   * Quando usar:
   * - salvar anexos em pasta institucional
   * - organizar por semestre/data
   */
  drive: Object.freeze({
    getFolderById: core_driveGetFolderById_,
    getFileById: core_driveGetFileById_,
    ensureFolder: core_driveEnsureFolder_,
    listFiles: core_driveListFiles_,
    moveFileToFolder: core_driveMoveFileToFolder_,
  }),

  /**
   * ----------------------------------------------------------
   * http (Webhooks/APIs)
   * ----------------------------------------------------------
   * Quando usar:
   * - enviar JSON para webhook (Google Chat, etc.)
   */
  http: Object.freeze({
    postJson: core_httpPostJson_,
  }),

  /**
   * ----------------------------------------------------------
   * assets (Imagens oficiais)
   * ----------------------------------------------------------
   * Quando usar:
   * - obter Blob de logo/imagens
   * - montar inlineImages para e-mail com <img src="cid:logo">
   *
   * Observação:
   * - isso depende dos arquivos do serviço de assets (ex.: 22_core_assets_service)
   */
  assets: Object.freeze({
    getAssetBlob: coreGetAssetBlob,              // <- sua função pública interna existente
    inlineImagesDefault: coreInlineImagesDefault // <- sua função pública interna existente
  }),

  /**
   * ----------------------------------------------------------
   * emailAssets (E-mail com imagens inline)
   * ----------------------------------------------------------
   * Quando usar:
   * - enviar e-mail HTML institucional com logo centralizada
   */
  emailAssets: Object.freeze({
    sendHtmlEmail: coreSendHtmlEmail, // <- sua função pública interna existente
  }),

/* ============================================================
 * GOVERNANCE / VIGÊNCIAS
 *
 * Camada institucional para consulta da diretoria vigente.
 *
 * Finalidades:
 * - identificar a diretoria atualmente em exercício
 * - retornar os membros da diretoria vigente
 * - localizar ocupantes atuais de cargos/funções institucionais
 * - servir de base para módulos que dependem da gestão atual,
 *   como processo seletivo, comunicações e automações internas
 *
 * Observação:
 * - a composição da diretoria é lida das planilhas de Vigências
 * - dados de contato (telefone/email) são cruzados com MEMBERS_ATUAIS
 * ============================================================ */
  governance: Object.freeze({
    getCurrentBoard: core_getCurrentBoard_,
    getCurrentBoardSlogan: core_getCurrentBoardSlogan_,
    getCurrentBoardMembers: core_getCurrentBoardMembers_,
    getCurrentBoardMembersByOccupation: core_getCurrentBoardMembersByOccupation_,
    getCurrentBoardMembersByRole: core_getCurrentBoardMembersByRole_,
    getCurrentBoardMemberByOccupation: core_getCurrentBoardMemberByOccupation_,
    getCurrentBoardMemberByRole: core_getCurrentBoardMemberByRole_,
    getCurrentLeadership: core_getCurrentLeadership_,
  }),

  /* ============================================================
 * SEMESTRES / RGA
 *
 * Camada para:
 * - descobrir o semestre institucional vigente
 * - interpretar ano/semestre de ingresso a partir do RGA
 * - calcular o semestre atual teórico do aluno
 *
 * Fonte de verdade:
 * - VIGENCIA_SEMESTRES
 * ============================================================ */
  semester: Object.freeze({
    getCurrentSemester: core_getCurrentSemester_,
    parseEntrySemesterFromRga: core_parseEntrySemesterFromRga_,
    getStudentCurrentSemesterFromRga: core_getStudentCurrentSemesterFromRga_,
  }),

  /* ============================================================
 * SEMESTRES / VIGÊNCIAS
 *
 * Camada para:
 * - descobrir o semestre institucional vigente
 * - descobrir o semestre institucional correspondente a uma data
 * - descobrir o último semestre concluído
 * - interpretar ano/semestre de ingresso a partir do RGA
 * - calcular semestre atual teórico do aluno
 * - calcular semestres concluídos no grupo pela data de entrada
 * ============================================================ */
  semester: Object.freeze({
    getCurrentSemester: core_getCurrentSemester_,
    getSemesterForDate: core_getSemesterForDate_,
    getSemesterIdForDate: core_getSemesterIdForDate_,
    getLastCompletedSemester: core_getLastCompletedSemester_,
    parseEntrySemesterFromRga: core_parseEntrySemesterFromRga_,
    getStudentCurrentSemesterFromRga: core_getStudentCurrentSemesterFromRga_,
    getCompletedGroupSemesterCountFromEntryDate: core_getCompletedGroupSemesterCountFromEntryDate_,
  }),

  /* ============================================================
  * MEMBERS_ATUAIS - CAMPOS DERIVADOS
  *
  * Rotinas de sincronização de colunas calculadas em MEMBERS_ATUAIS,
  * como semestre atual e número de semestres concluídos no grupo.
  * ============================================================ */
  membersCurrent: Object.freeze({
    syncDerivedFields: core_syncMembersCurrentDerivedFields_,
  }),

  identity: Object.freeze({
    normalizeKey: core_normalizeIdentityKey_,
    findByAny: core_memberIdentityFindByAny_,
    findCurrentRowByAny: core_findMemberCurrentRowByAny_,
    autofillRowInSheet: core_autofillIdentityRowInSheet_,
  }),
});
