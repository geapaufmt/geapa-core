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

function test_core_civil_group_time_from_integration_date() {
  Logger.log(JSON.stringify({
    sameMonth: core_getCivilGroupTimeFromIntegrationDate_(new Date(2026, 2, 15), new Date(2026, 3, 1)),
    oneYearTwoMonths: core_getCivilGroupTimeFromIntegrationDate_(new Date(2025, 0, 15), new Date(2026, 2, 20)),
    formatted: core_formatCivilGroupTime_(
      core_getCivilGroupTimeFromIntegrationDate_(new Date(2024, 0, 10), new Date(2026, 3, 1))
    )
  }, null, 2));
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

function test_core_mailRenderer_render_operacional() {
  Logger.log(JSON.stringify(coreMailRenderEmailTemplate_(
    'GEAPA_OPERACIONAL',
    'Confirmacao de dados cadastrais',
    {
      subtitle: 'Fluxo institucional do GEAPA',
      introText: 'Esta mensagem usa o layout institucional centralizado do GEAPA, enquanto o modulo continua dono do conteudo de negocio.',
      blocks: [
        {
          title: 'Proximos passos',
          text: 'Revise os dados abaixo e responda este e-mail caso encontre qualquer inconsistencia.'
        },
        {
          title: 'Resumo',
          items: [
            { label: 'Modulo', value: 'MEMBROS' },
            { label: 'Chave', value: 'MEM-2026-001' },
            { label: 'Prazo', value: '03/04/2026' }
          ]
        }
      ],
      cta: {
        label: 'Responder este e-mail',
        helper: 'Se preferir, voce tambem pode responder diretamente nesta mesma conversa.'
      }
    }
  ), null, 2));
}

function test_core_mailRenderer_render_convite() {
  Logger.log(JSON.stringify(coreMailRenderEmailTemplate_(
    'GEAPA_CONVITE',
    'Convite para atividade do GEAPA',
    {
      subtitle: 'Participacao institucional',
      introText: 'Convidamos voce para uma atividade organizada pelo GEAPA.',
      blocks: [
        {
          title: 'Detalhes',
          items: [
            { label: 'Data', value: '05/04/2026' },
            { label: 'Horario', value: '19:00' },
            { label: 'Local', value: 'Sala do GEAPA' }
          ]
        }
      ],
      cta: {
        label: 'Confirmar presenca',
        url: 'https://example.com/confirmacao'
      }
    }
  ), null, 2));
}

function test_core_mailRenderer_render_classico() {
  Logger.log(JSON.stringify(coreMailRenderEmailTemplate_(
    'GEAPA_CLASSICO',
    'Feliz aniversario, Luis!',
    {
      subtitle: '02/04/2026',
      introText: 'O GEAPA deseja um dia especial, com saude, alegria e muitas realizacoes.',
      blocks: [
        {
          items: [
            { line1: 'Luis Putton', line2: 'Membro do GEAPA' },
            { line1: '@luisputton', line2: 'Instagram' }
          ]
        }
      ],
      footerNote: 'Mensagem institucional automatica do GEAPA.'
    }
  ), null, 2));
}

function test_core_mailRenderer_buildFinalSubject() {
  Logger.log(coreMailBuildFinalSubject_('Convite para atividade do GEAPA', 'APR-2026-010'));
}

function test_core_mailRenderer_buildOutgoingDraft() {
  Logger.log(JSON.stringify(coreMailBuildOutgoingDraft_({
    moduleName: 'APRESENTACOES',
    templateKey: 'GEAPA_CONVITE',
    correlationKey: 'APR-2026-010',
    to: 'destinatario@exemplo.com',
    cc: 'apoio@exemplo.com',
    subjectHuman: 'Convite para atividade do GEAPA',
    payload: {
      subtitle: 'Template central institucional',
      introText: 'O modulo informa o conteudo, e o geapa-core monta o HTML final.',
      blocks: [
        {
          title: 'Agenda',
          items: [
            { label: 'Tema', value: 'Apresentacao institucional' },
            { label: 'Data', value: '05/04/2026' }
          ]
        }
      ]
    }
  }), null, 2));
}

function test_core_governance_currentBoardSlogan() {
  Logger.log(core_getCurrentBoardSlogan_());
}

function test_core_mailOutbox_queue_operacional() {
  Logger.log(JSON.stringify(coreMailQueueOutgoing_({
    moduleName: 'COMEMORACOES',
    templateKey: 'GEAPA_OPERACIONAL',
    correlationKey: 'COM-2026-1-MAT-ABERTURA',
    entityType: 'SEMESTRE',
    entityId: '2026/1',
    flowCode: 'MAT',
    stage: 'ABERTURA',
    to: coreMailHubEnvelopeOfficialEmail_() || Session.getActiveUser().getEmail(),
    subjectHuman: 'Aviso institucional de teste',
    payload: {
      eyebrow: 'Aviso academico',
      title: 'Teste da fila MAIL_SAIDA',
      introText: 'Este teste valida a V1 minima da fila central de saida.',
      blocks: [
        {
          title: 'Contexto',
          text: 'A mensagem foi enfileirada pelo geapa-core para processamento posterior.'
        }
      ]
    },
    metadata: {
      source: 'test_core_mailOutbox_queue_operacional'
    }
  }), null, 2));
}

function test_core_mailOutbox_process() {
  Logger.log(JSON.stringify(coreMailProcessOutbox_(), null, 2));
}
