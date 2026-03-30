function test_core_rolesConfig_debug() {
  Logger.log(JSON.stringify(core_debugInstitutionalRolesConfig_(), null, 2));
}

function test_core_rolesConfig_byKey() {
  Logger.log(JSON.stringify(
    core_findInstitutionalRoleByKey_('SECRETARIO_GERAL'),
    null,
    2
  ));
}

function test_core_rolesConfig_byPublicName() {
  Logger.log(JSON.stringify(
    core_findInstitutionalRoleByPublicName_('Secretário(a) Geral'),
    null,
    2
  ));
}

function test_core_rolesConfig_byAlias() {
  Logger.log(JSON.stringify(
    core_findInstitutionalRoleByAnyName_('Secretario Geral'),
    null,
    2
  ));
}

function test_core_rolesConfig_secretaria() {
  Logger.log(JSON.stringify(
    core_getInstitutionalRolesByEmailGroup_('SECRETARIA'),
    null,
    2
  ));
}

function test_core_rolesConfig_comunicacao() {
  Logger.log(JSON.stringify(
    core_getInstitutionalRolesByEmailGroup_('COMUNICACAO'),
    null,
    2
  ));
}

function test_core_rolesConfig_eventos() {
  Logger.log(JSON.stringify(
    core_getInstitutionalRolesByEmailGroup_('EVENTOS'),
    null,
    2
  ));
}

function test_core_rolesConfig_allActive() {
  Logger.log(JSON.stringify(
    core_getInstitutionalRolesActive_(),
    null,
    2
  ));
}

function test_core_rolesConfig_clearCacheAndDebug() {
  core_rolesConfigCacheClear_();
  Logger.log(JSON.stringify(core_debugInstitutionalRolesConfig_(), null, 2));
}

function test_projection_debug() {
  Logger.log(JSON.stringify(core_debugCurrentInstitutionalProjection_(), null, 2));
}

function test_projection_secretaria_html() {
  Logger.log(core_getCurrentContactsHtmlByEmailGroup_('SECRETARIA'));
}

function test_projection_sync_members() {
  Logger.log(JSON.stringify(core_syncMembersCurrentInstitutionalRoles_(), null, 2));
}

function test_projection_sync_members_and_debug() {
  Logger.log(JSON.stringify(core_syncMembersCurrentInstitutionalRoles_(), null, 2));
  Logger.log(JSON.stringify(core_debugCurrentInstitutionalProjection_(), null, 2));
}

function test_memberIdentity_byRga() {
  Logger.log(JSON.stringify(core_memberIdentityFindByAny_('SEU_RGA_AQUI'), null, 2));
}

function test_memberIdentity_byEmail() {
  Logger.log(JSON.stringify(core_memberIdentityFindByAny_('email@exemplo.com'), null, 2));
}

function test_sync_all_derived_fields() {
  Logger.log(JSON.stringify(core_syncMembersCurrentDerivedFields_(), null, 2));
}

function test_core_mailHub_assertSchema() {
  Logger.log(JSON.stringify(coreMailHubAssertSchema_(), null, 2));
}

function test_core_mailHub_ingestInbox_dryRun() {
  Logger.log(JSON.stringify(core_mailIngestInbox_({
    dryRun: true,
    maxThreads: 10,
    maxMessagesPerThread: 20
  }), null, 2));
}

function test_core_mailHub_ingestInbox_real() {
  Logger.log(JSON.stringify(core_mailIngestInbox_({
    maxThreads: 10,
    maxMessagesPerThread: 20
  }), null, 2));
}

function test_core_mailHub_config_read() {
  Logger.log(JSON.stringify({
    assuntoPrefixoObrigatorio: coreMailHubGetConfig_('ASSUNTO_PREFIXO_OBRIGATORIO', '[GEAPA]'),
    usarSomenteAssuntosGeapa: coreMailHubGetConfigBooleanByKey_('USAR_SOMENTE_ASSUNTOS_GEAPA', false),
    ignorarRemetentes: coreMailHubGetConfigListByKey_('IGNORAR_REMETENTES'),
    ignorarDominios: coreMailHubGetConfigListByKey_('IGNORAR_DOMINIOS'),
    ignorarAssuntosRegex: coreMailHubGetConfigListByKey_('IGNORAR_ASSUNTOS_REGEX'),
    maxEventosPorExecucao: coreMailHubGetConfig_('MAX_EVENTOS_POR_EXECUCAO', 0),
    salvarCorpoCompleto: coreMailHubGetConfigBooleanByKey_('SALVAR_CORPO_COMPLETO', false),
    marcarRuidoComoIgnorado: coreMailHubGetConfigBooleanByKey_('MARCAR_RUIDO_COMO_IGNORADO', false)
  }, null, 2));
}

function test_core_mailHub_ingestInbox_higiene_dryRun() {
  Logger.log(JSON.stringify(core_mailIngestInbox_({
    dryRun: true,
    maxThreads: 10,
    maxMessagesPerThread: 20
  }), null, 2));
}

function test_core_mailHub_listPending_naoIdentificado() {
  Logger.log(JSON.stringify(core_mailListPendingByModule_('NAO_IDENTIFICADO'), null, 2));
}

function test_core_mailHub_listPending_membros() {
  Logger.log(JSON.stringify(core_mailListPendingByModule_('MEMBROS'), null, 2));
}

function test_core_mailHub_getLatestEvent() {
  Logger.log(JSON.stringify(core_mailGetLatestEvent_(), null, 2));
}

function test_core_mailHub_getLatestPending_membros() {
  Logger.log(JSON.stringify(core_mailGetLatestEvent_({
    moduleName: 'MEMBROS',
    processingStatus: 'PENDENTE'
  }), null, 2));
}

function test_core_mailHub_markLatestPending_membros_processed() {
  Logger.log(JSON.stringify(core_mailMarkLatestPendingByModule_(
    'MEMBROS',
    'test_core_mailHub_markLatestPending_membros_processed'
  ), null, 2));
}

function test_core_mailHub_markLatestPending_naoIdentificado_processed() {
  Logger.log(JSON.stringify(core_mailMarkLatestPendingByModule_(
    'NAO_IDENTIFICADO',
    'test_core_mailHub_markLatestPending_naoIdentificado_processed'
  ), null, 2));
}

function test_core_mailHub_cleanupNoiseEvents() {
  Logger.log(JSON.stringify(coreMailCleanupNoiseEvents_(), null, 2));
}

function test_core_mailHub_extractCorrelationKey() {
  Logger.log(coreMailExtractCorrelationKey_('[GEAPA][MEM-2026-001] Assunto de teste'));
}

function test_core_mailAdapters_list() {
  Logger.log(JSON.stringify(coreMailListModuleAdapters_().map(function(adapter) {
    return coreMailAdapterToSnapshot_(adapter);
  }), null, 2));
}

function test_core_mailAdapters_get_mem() {
  Logger.log(JSON.stringify(coreMailAdapterToSnapshot_(coreMailGetModuleAdapter_('MEM')), null, 2));
}

function test_core_mailAdapters_build_mem() {
  Logger.log(coreMailBuildCorrelationKey_('MEM', {
    businessId: '2026-001',
    flowCode: 'ENTRADA',
    stage: 'CONFIRMACAO'
  }));
}

function test_core_mailAdapters_parse_mem() {
  Logger.log(JSON.stringify(coreMailParseCorrelationKey_('MEM-2026-001-ENTRADA-CONFIRMACAO'), null, 2));
}

function test_core_mailAdapters_resolveRouting_mem() {
  Logger.log(JSON.stringify(coreMailResolveRouting_({
    subject: '[GEAPA][MEM-2026-001-ENTRADA-CONFIRMACAO] Confirmacao de ingresso',
    fromEmail: 'teste@exemplo.com',
    snippet: 'aceito ingressar no geapa'
  }), null, 2));
}

function test_core_mailAdapters_normalizeOutgoingSubject_mem() {
  Logger.log(coreMailNormalizeOutgoingSubject_('MEM', 'Confirmacao de ingresso', {
    businessId: '2026-001',
    flowCode: 'ENTRADA',
    stage: 'CONFIRMACAO'
  }));
}
