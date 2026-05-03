function test_core_rolesConfig_debug() {
  Logger.log(JSON.stringify(core_debugInstitutionalRolesConfig_(), null, 2));
}

function test_core_modulesConfig_debug() {
  Logger.log(JSON.stringify(core_debugModulesConfig_(), null, 2));
}

function test_core_modulesConfig_clearCacheAndDebug() {
  core_modulesConfigCacheClear_();
  Logger.log(JSON.stringify(core_debugModulesConfig_(), null, 2));
}

function test_core_modulesConfig_applySheetUx() {
  Logger.log(JSON.stringify(core_applyModulesConfigSheetUx_(), null, 2));
}

function test_core_modulesConfig_atividades_geral() {
  Logger.log(JSON.stringify(core_getModuleConfig_('ATIVIDADES', 'GERAL'), null, 2));
}

function test_core_modulesConfig_apresentacoes_geral() {
  Logger.log(JSON.stringify(core_getModuleConfig_('APRESENTACOES', 'GERAL'), null, 2));
}

function test_core_modulesConfig_canTrigger_atividades() {
  Logger.log(JSON.stringify({
    canTrigger: core_canModuleUseCapability_('ATIVIDADES', 'GERAL', 'TRIGGER', {
      executionType: 'TRIGGER'
    })
  }, null, 2));
}

function test_core_modulesConfig_assertTrigger_atividades() {
  Logger.log(JSON.stringify(core_assertModuleExecutionAllowed_('ATIVIDADES', 'GERAL', 'TRIGGER', {
    executionType: 'TRIGGER'
  }), null, 2));
}

function test_core_modulesConfig_canEmail_apresentacoes() {
  Logger.log(JSON.stringify({
    canEmail: core_canModuleUseCapability_('APRESENTACOES', 'GERAL', 'EMAIL', {
      executionType: 'MANUAL'
    })
  }, null, 2));
}

function test_core_modulesStatus_debug() {
  Logger.log(JSON.stringify(core_debugModulesStatus_(), null, 2));
}

function test_core_modulesStatus_get_atividades_geral() {
  Logger.log(JSON.stringify(core_moduleStatusGet_('ATIVIDADES', 'GERAL'), null, 2));
}

function test_core_modulesStatus_ensure_atividades_geral() {
  Logger.log(JSON.stringify(core_moduleStatusEnsureRow_('ATIVIDADES', 'GERAL'), null, 2));
}

function test_core_modulesStatus_markExecution_atividades_geral() {
  Logger.log(JSON.stringify(core_moduleStatusMarkExecution_('ATIVIDADES', 'GERAL', 'SYNC', {
    modeRead: 'ON',
    obs: 'Teste manual de registro de execucao.'
  }), null, 2));
}

function test_core_modulesStatus_markSuccess_atividades_geral() {
  Logger.log(JSON.stringify(core_moduleStatusMarkSuccess_('ATIVIDADES', 'GERAL', 'SYNC', {
    modeRead: 'ON'
  }), null, 2));
}

function test_core_modulesStatus_markError_atividades_geral() {
  Logger.log(JSON.stringify(core_moduleStatusMarkError_('ATIVIDADES', 'GERAL', 'Erro manual de teste', 'SYNC', {
    modeRead: 'ON'
  }), null, 2));
}

function test_core_modulesStatus_markBlocked_apresentacoes_geral() {
  Logger.log(JSON.stringify(core_moduleStatusMarkBlocked_(
    'APRESENTACOES',
    'GERAL',
    'MODO_OFF',
    'Teste manual de bloqueio por MODULOS_CONFIG',
    'TRIGGER',
    'OFF'
  ), null, 2));
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

function test_core_portalBuscarMembroParaPortal_fakeSheet() {
  var sheet = test_createFakeSheet_([
    ['ID_MEMBRO', 'Membro', 'EMAIL', 'RGA', 'Status', 'Cargo/fun\u00E7\u00E3o atual', 'Telefone'],
    ['MEM-TESTE-001', 'Membro Teste', 'Membro@Exemplo.com', 'RGA-TESTE-001', 'Ativo', 'Membro', '(00) 0000-0000'],
    ['MEM-TESTE-002', 'Outro Membro', 'outro@exemplo.com', 'RGA-TESTE-002', 'Suspenso', 'Membro', '(00) 0000-0001']
  ], 'MEMBERS_ATUAIS_FAKE');

  var byEmail = core_buscarMembroParaPortalInSheet_(sheet, '  MEMBRO@EXEMPLO.COM  ');
  var byRga = core_buscarMembroParaPortalInSheet_(sheet, 'rga-teste-002');
  var missing = core_buscarMembroParaPortalInSheet_(sheet, 'nao-encontrado@example.com');
  var keys = Object.keys(byEmail || {}).sort();

  test_assert_(!!byEmail, 'Busca por email deveria encontrar membro.');
  test_assert_(byEmail.id === 'RGA-TESTE-001', 'id publico deve usar RGA, nao ID_MEMBRO.');
  test_assert_(byEmail.nomeExibicao === 'Membro Teste', 'nomeExibicao incorreto.');
  test_assert_(byEmail.emailCadastrado === 'membro@exemplo.com', 'emailCadastrado deveria ser normalizado em lowercase.');
  test_assert_(byEmail.rga === 'RGA-TESTE-001', 'rga incorreto.');
  test_assert_(byEmail.situacaoGeral === 'Ativo', 'situacaoGeral deveria vir de Status.');
  test_assert_(byEmail.vinculo === 'Membro', 'vinculo deveria vir da ocupacao atual.');
  test_assert_(keys.join(',') === 'emailCadastrado,id,nomeExibicao,rga,situacaoGeral,vinculo', 'Contrato retornou campos fora do esperado.');

  test_assert_(!!byRga, 'Busca por RGA deveria encontrar membro.');
  test_assert_(byRga.emailCadastrado === 'outro@exemplo.com', 'Busca por RGA deve devolver o email cadastrado oficial.');
  test_assert_(byRga.situacaoGeral === 'Suspenso', 'Status nao ativo tambem deve ser preservado para o portal.');
  test_assert_(missing === null, 'Membro inexistente deveria retornar null.');

  Logger.log(JSON.stringify({
    ok: true,
    byEmail: byEmail,
    byRga: byRga,
    missing: missing
  }, null, 2));
}

function test_core_portalBuscarMinhaSituacaoParaPortal_fakeSheet() {
  var sheet = test_createFakeSheet_([
    [
      'ID_MEMBRO',
      'Membro',
      'EMAIL',
      'RGA',
      'Status',
      'Cargo/fun\u00E7\u00E3o atual',
      'Telefone',
      'PERIODO_ULTIMA_APRESENTACAO',
      'QTD_APRESENTACOES_REALIZADAS'
    ],
    ['MEM-TESTE-001', 'Membro Sem Pendencia', 'sem-pendencia@exemplo.com', 'RGA-TESTE-001', 'Ativo', 'Membro', '(00) 0000-0000', 'GEAPA_2025', '2'],
    ['MEM-TESTE-002', 'Membro Sem RGA', 'sem-rga@exemplo.com', '', 'Ativo', 'Membro', '(00) 0000-0001', '', ''],
    ['MEM-TESTE-003', '', 'sem-nome@exemplo.com', 'RGA-TESTE-003', 'Ativo', 'Membro', '(00) 0000-0002', 'GEAPA_2024', '1'],
    ['MEM-TESTE-004', 'Membro Indefinido', 'indefinido@exemplo.com', 'RGA-TESTE-004', 'Indefinido', 'Indefinido', '(00) 0000-0003', 'GEAPA_2025', 'invalido'],
    ['MEM-TESTE-005', 'Membro Email Invalido', 'email-invalido', 'RGA-TESTE-005', 'Ativo', 'Membro', '(00) 0000-0004', '', '']
  ], 'MEMBERS_ATUAIS_FAKE');

  var result = core_buscarMinhaSituacaoParaPortalInSheet_(sheet, 'RGA-TESTE-001');
  var semRga = core_buscarMinhaSituacaoParaPortalInSheet_(sheet, 'sem-rga@exemplo.com');
  var semNome = core_buscarMinhaSituacaoParaPortalInSheet_(sheet, 'RGA-TESTE-003');
  var indefinido = core_buscarMinhaSituacaoParaPortalInSheet_(sheet, 'RGA-TESTE-004');
  var emailInvalido = core_buscarMinhaSituacaoParaPortalInSheet_(sheet, 'RGA-TESTE-005');
  var missing = core_buscarMinhaSituacaoParaPortalInSheet_(sheet, 'ausente@example.com');
  var memberKeys = Object.keys(result.membro || {}).sort();

  test_assert_(result.ok === true, 'Minha situacao deveria retornar ok=true.');
  test_assert_(memberKeys.join(',') === 'emailCadastrado,id,nomeExibicao,rga,situacaoGeral,vinculo', 'membro retornou campos fora do contrato.');
  test_assert_(result.membro.id === 'RGA-TESTE-001', 'id nao deve expor ID_MEMBRO.');
  test_assert_(result.membro.emailCadastrado === 'sem-pendencia@exemplo.com', 'email cadastrado incorreto.');
  test_assert_(result.minhaSituacao.resumo.frequencia === '', 'frequencia deve ficar vazia enquanto nao houver fonte oficial integrada.');
  test_assert_(result.minhaSituacao.resumo.pendenciasAbertas === 0, 'pendenciasAbertas deve iniciar zerado.');
  test_assert_(result.minhaSituacao.resumo.certificadosDisponiveis === 0, 'certificadosDisponiveis deve iniciar zerado.');
  test_assert_(Array.isArray(result.minhaSituacao.pendencias) && result.minhaSituacao.pendencias.length === 0, 'pendencias deve ser lista vazia do proprio membro.');
  test_assert_(Array.isArray(result.minhaSituacao.participacao.atividadesRecentes) && result.minhaSituacao.participacao.atividadesRecentes.length === 0, 'atividadesRecentes deve iniciar vazia.');
  test_assert_(result.minhaSituacao.participacao.frequenciaGeral === '', 'frequenciaGeral deve continuar vazia nesta etapa.');
  test_assert_(result.minhaSituacao.participacao.apresentacoes.periodoUltimaApresentacao === 'GEAPA_2025', 'Periodo atual de apresentacao incorreto.');
  test_assert_(result.minhaSituacao.participacao.apresentacoes.quantidadeRealizadas === 2, 'Quantidade atual de apresentacoes deveria ser numero.');
  test_assert_(!Object.prototype.hasOwnProperty.call(result.minhaSituacao.participacao.apresentacoes, 'quantidadeRealizadasBaseLegado'), 'Quantidade legado nao deve ser exposta no portal.');
  test_assert_(!Object.prototype.hasOwnProperty.call(result.minhaSituacao.participacao.apresentacoes, 'periodoUltimaApresentacaoBaseLegado'), 'Periodo legado nao deve ser exposto no portal.');
  test_assert_(Array.isArray(result.minhaSituacao.certificados) && result.minhaSituacao.certificados.length === 0, 'certificados deve iniciar vazio.');
  test_assert_(Array.isArray(result.minhaSituacao.avisos) && result.minhaSituacao.avisos.length === 0, 'avisos deve iniciar vazio.');
  test_assert_(semRga.minhaSituacao.pendencias.length === 1, 'Membro sem RGA deveria ter uma pendencia.');
  test_assert_(semRga.minhaSituacao.pendencias[0].titulo === 'RGA nao informado', 'Pendencia de RGA incorreta.');
  test_assert_(semRga.minhaSituacao.resumo.pendenciasAbertas === semRga.minhaSituacao.pendencias.length, 'Contagem de pendencias sem RGA incorreta.');
  test_assert_(semRga.minhaSituacao.participacao.apresentacoes.quantidadeRealizadas === 0, 'Membro sem apresentacoes atuais deveria retornar quantidade 0.');
  test_assert_(semRga.minhaSituacao.participacao.apresentacoes.periodoUltimaApresentacao === '', 'Membro sem apresentacoes atuais deveria retornar periodo vazio.');
  test_assert_(semNome.minhaSituacao.pendencias.length === 1, 'Membro sem nome deveria ter uma pendencia.');
  test_assert_(semNome.minhaSituacao.pendencias[0].titulo === 'Nome de exibicao nao informado', 'Pendencia de nome incorreta.');
  test_assert_(semNome.minhaSituacao.resumo.pendenciasAbertas === semNome.minhaSituacao.pendencias.length, 'Contagem de pendencias sem nome incorreta.');
  test_assert_(semNome.minhaSituacao.participacao.apresentacoes.periodoUltimaApresentacao === 'GEAPA_2024', 'Periodo consolidado incorreto.');
  test_assert_(semNome.minhaSituacao.participacao.apresentacoes.quantidadeRealizadas === 1, 'Quantidade consolidada deveria ser numero.');
  test_assert_(indefinido.minhaSituacao.pendencias.length === 2, 'Membro com situacao/vinculo indefinido deveria ter duas pendencias.');
  test_assert_(indefinido.minhaSituacao.pendencias[0].titulo === 'Vinculo cadastral indefinido', 'Pendencia de vinculo incorreta.');
  test_assert_(indefinido.minhaSituacao.pendencias[1].titulo === 'Situacao geral indefinida', 'Pendencia de situacao geral incorreta.');
  test_assert_(indefinido.minhaSituacao.resumo.pendenciasAbertas === indefinido.minhaSituacao.pendencias.length, 'Contagem de pendencias indefinidas incorreta.');
  test_assert_(indefinido.minhaSituacao.participacao.apresentacoes.quantidadeRealizadas === 0, 'Quantidade atual invalida deveria virar 0.');
  test_assert_(emailInvalido.minhaSituacao.pendencias.length === 1, 'Membro com email invalido deveria ter uma pendencia.');
  test_assert_(emailInvalido.membro.emailCadastrado === '', 'Email invalido nao deve ser retornado como emailCadastrado.');
  test_assert_(emailInvalido.minhaSituacao.pendencias[0].titulo === 'E-mail cadastrado ausente ou invalido', 'Pendencia de email invalido incorreta.');
  test_assert_(emailInvalido.minhaSituacao.resumo.pendenciasAbertas === emailInvalido.minhaSituacao.pendencias.length, 'Contagem de pendencias de email invalido incorreta.');
  test_assert_(missing.ok === false && missing.code === 'MEMBRO_NAO_ENCONTRADO', 'Membro ausente deveria retornar erro controlado.');

  Logger.log(JSON.stringify({
    ok: true,
    result: result,
    semRga: semRga,
    semNome: semNome,
    indefinido: indefinido,
    emailInvalido: emailInvalido,
    missing: missing
  }, null, 2));
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

function test_core_mailHub_applyOperationalSheetUx() {
  Logger.log(JSON.stringify(coreMailApplyOperationalSheetUx_(), null, 2));
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

function test_core_mailRoutingRules_resolve_apresentacoes() {
  Logger.log(JSON.stringify(coreMailResolveRouting_({
    subject: '[GEAPA][APR-2026-001] Resposta de apresentacao',
    fromEmail: 'teste@exemplo.com',
    snippet: 'Mensagem de teste para regra de apresentacoes.'
  }), null, 2));
}

function test_core_mailRoutingRules_resolve_seletivo() {
  Logger.log(JSON.stringify(coreMailResolveRouting_({
    subject: '[GEAPA][SEL-2026-001] Resposta de processo seletivo',
    fromEmail: 'teste@exemplo.com',
    snippet: 'Mensagem de teste para regra de seletivo.'
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
function test_core_identity_fillMissingProfessorIds() {
  Logger.log(JSON.stringify(core_fillMissingProfessorIds_(), null, 2));
}

function test_core_identity_fillMissingExternalIds() {
  Logger.log(JSON.stringify(core_fillMissingExternalIds_(), null, 2));
}

function test_core_identity_validateExternalEmailDuplicates() {
  Logger.log(JSON.stringify(core_identityValidateExternalEmailDuplicates_(), null, 2));
}

function test_core_identity_findExternalByEmail_example() {
  Logger.log(JSON.stringify(core_identityFindExternalByEmail_('email@exemplo.com'), null, 2));
}

function test_core_mailHub_listPendingAttachments() {
  Logger.log(JSON.stringify(core_mailListPendingAttachments_({ limit: 20 }), null, 2));
}

function test_core_mailHub_getLatestPendingEventWithAttachment() {
  Logger.log(JSON.stringify(core_mailGetLatestPendingEventWithAttachment_({}), null, 2));
}

function test_core_mailHub_getAttachmentById_example(attachmentId) {
  Logger.log(JSON.stringify(core_mailGetAttachmentById_(attachmentId, { includeBlob: false }), null, 2));
}

function test_core_mailHub_markAttachmentProcessed_example(attachmentId) {
  Logger.log(JSON.stringify(core_mailMarkAttachmentProcessed_(attachmentId, 'test_core_mailHub_markAttachmentProcessed_example', 'Marcado manualmente em teste.'), null, 2));
}

function test_core_mailHub_markAttachmentSavedToDrive_example(attachmentId) {
  Logger.log(JSON.stringify(core_mailMarkAttachmentSavedToDrive_(attachmentId, 'test_core_mailHub_markAttachmentSavedToDrive_example', {
    driveFileId: 'DRIVE_FILE_ID_EXEMPLO',
    driveFileUrl: 'https://drive.google.com/file/d/DRIVE_FILE_ID_EXEMPLO/view',
    driveFolder: 'PASTA_DESTINO_EXEMPLO',
    observations: 'Salvo no Drive em teste manual.'
  }), null, 2));
}

function test_core_mailHub_markAttachmentError_example(attachmentId) {
  Logger.log(JSON.stringify(core_mailMarkAttachmentError_(attachmentId, 'test_core_mailHub_markAttachmentError_example', 'Erro simulado para validacao operacional.'), null, 2));
}

function test_assert_(condition, message) {
  if (!condition) {
    throw new Error(message || 'Assercao falhou.');
  }
}

function test_createFakeSheet_(matrix, sheetName) {
  var data = (matrix || []).map(function(row) {
    return row.slice();
  });

  function ensureCell_(rowNumber, colNumber) {
    while (data.length < rowNumber) data.push([]);
    while (data[rowNumber - 1].length < colNumber) data[rowNumber - 1].push('');
  }

  return {
    getName: function() {
      return sheetName || 'FAKE_SHEET';
    },
    getLastColumn: function() {
      return data.length ? data[0].length : 0;
    },
    getLastRow: function() {
      return data.length;
    },
    appendRow: function(row) {
      data.push((row || []).slice());
    },
    getRange: function(row, col, numRows, numCols) {
      var startRow = Number(row);
      var startCol = Number(col);
      var rowCount = Number(numRows || 1);
      var colCount = Number(numCols || 1);

      return {
        getValues: function() {
          var values = [];
          for (var r = 0; r < rowCount; r++) {
            var line = [];
            for (var c = 0; c < colCount; c++) {
              var sourceRow = data[startRow + r - 1] || [];
              line.push(typeof sourceRow[startCol + c - 1] === 'undefined' ? '' : sourceRow[startCol + c - 1]);
            }
            values.push(line);
          }
          return values;
        },
        setValues: function(values) {
          for (var r = 0; r < rowCount; r++) {
            for (var c = 0; c < colCount; c++) {
              ensureCell_(startRow + r, startCol + c);
              data[startRow + r - 1][startCol + c - 1] = values[r][c];
            }
          }
        },
        setValue: function(value) {
          ensureCell_(startRow, startCol);
          data[startRow - 1][startCol - 1] = value;
        }
      };
    }
  };
}

function test_core_memberLifecycle_updateEvent_patch_fakeSheet() {
  var baseUpdatedAt = new Date(2026, 3, 12, 8, 0, 0);
  var processingDate = new Date(2026, 3, 18, 9, 30, 0);
  var sheet = test_createFakeSheet_([
    [
      'ID_EVENTO_MEMBRO',
      'RGA',
      'TIPO_EVENTO',
      'DATA_EVENTO',
      'STATUS_EVENTO',
      'MOTIVO_EVENTO',
      'ORIGEM_MODULO',
      'ORIGEM_CHAVE',
      'ORIGEM_ROW',
      'NOME_MEMBRO',
      'EMAIL',
      'OBSERVACOES',
      'CRIADO_EM',
      'ATUALIZADO_EM'
    ],
    [
      'MEV-000001',
      '2023001',
      'DESLIGAMENTO_POR_FALTAS',
      new Date(2026, 3, 10, 0, 0, 0),
      'HOMOLOGADO',
      'Faltas recorrentes',
      'geapa-atividades',
      'ATV-2026-001',
      '17',
      'Membro Teste',
      'membro@exemplo.com',
      'Aguardando processamento em membros.',
      new Date(2026, 3, 10, 8, 0, 0),
      baseUpdatedAt
    ]
  ], 'MEMBER_EVENTOS_VINCULO_FAKE');

  var first = core_memberLifecycleApplyPatchToSheet_(sheet, 'MEV-000001', {
    eventStatus: 'PROCESSADO_MEMBROS',
    observacoes: 'Processado efetivamente pelo modulo de membros.',
    processedByModule: 'geapa-membros',
    processingDate: processingDate,
    processingError: ''
  });

  test_assert_(first.updated === true, 'O primeiro patch deveria atualizar o evento.');
  test_assert_(first.event.eventStatus === 'PROCESSADO_MEMBROS', 'STATUS_EVENTO nao foi atualizado.');
  test_assert_(first.event.notes === 'Processado efetivamente pelo modulo de membros.', 'OBSERVACOES nao foi atualizada.');
  test_assert_(first.event.processedByModule === 'geapa-membros', 'PROCESSADO_POR_MODULO nao foi atualizado.');
  test_assert_(first.event.processingDate && first.event.processingDate.getTime() === processingDate.getTime(), 'DATA_PROCESSAMENTO nao foi atualizada.');
  test_assert_(first.event.processingError === '', 'ERRO_PROCESSAMENTO deveria estar vazio.');
  test_assert_(first.event.updatedAt && first.event.updatedAt.getTime() !== baseUpdatedAt.getTime(), 'ATUALIZADO_EM deveria refletir a atualizacao.');

  var headersAfterExtension = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  test_assert_(headersAfterExtension.indexOf('PROCESSADO_POR_MODULO') >= 0, 'Cabecalho PROCESSADO_POR_MODULO nao foi criado.');
  test_assert_(headersAfterExtension.indexOf('DATA_PROCESSAMENTO') >= 0, 'Cabecalho DATA_PROCESSAMENTO nao foi criado.');
  test_assert_(headersAfterExtension.indexOf('ERRO_PROCESSAMENTO') >= 0, 'Cabecalho ERRO_PROCESSAMENTO nao foi criado.');

  var firstUpdatedAt = first.event.updatedAt;
  var second = core_memberLifecycleApplyPatchToSheet_(sheet, 'MEV-000001', {
    eventStatus: 'PROCESSADO_MEMBROS',
    observacoes: 'Processado efetivamente pelo modulo de membros.',
    processedByModule: 'geapa-membros',
    processingDate: processingDate,
    processingError: ''
  });

  test_assert_(second.updated === false, 'Retry identico deveria ser idempotente.');
  test_assert_(second.event.updatedAt && second.event.updatedAt.getTime() === firstUpdatedAt.getTime(), 'ATUALIZADO_EM nao deveria mudar em retry idempotente.');

  Logger.log(JSON.stringify({
    ok: true,
    first: first,
    second: second
  }, null, 2));
}

function test_core_memberLifecycle_updateEvent_invalidStatus_fakeSheet() {
  var sheet = test_createFakeSheet_([
    [
      'ID_EVENTO_MEMBRO',
      'RGA',
      'TIPO_EVENTO',
      'DATA_EVENTO',
      'STATUS_EVENTO',
      'MOTIVO_EVENTO',
      'ORIGEM_MODULO',
      'ORIGEM_CHAVE',
      'ORIGEM_ROW',
      'NOME_MEMBRO',
      'EMAIL',
      'OBSERVACOES',
      'CRIADO_EM',
      'ATUALIZADO_EM'
    ],
    [
      'MEV-000002',
      '2023002',
      'DESLIGAMENTO_POR_FALTAS',
      new Date(2026, 3, 11, 0, 0, 0),
      'REGISTRADO',
      'Teste',
      'geapa-atividades',
      'ATV-2026-002',
      '18',
      'Outro Membro',
      'outro@exemplo.com',
      '',
      new Date(2026, 3, 11, 8, 0, 0),
      new Date(2026, 3, 11, 8, 0, 0)
    ]
  ], 'MEMBER_EVENTOS_VINCULO_FAKE');

  var failed = false;

  try {
    core_memberLifecycleApplyPatchToSheet_(sheet, 'MEV-000002', {
      eventStatus: 'STATUS_INVALIDO'
    });
  } catch (err) {
    failed = true;
    test_assert_(
      String(err && err.message || '').indexOf('nao suportado') >= 0,
      'A validacao deveria rejeitar status invalido com mensagem clara.'
    );
  }

  test_assert_(failed === true, 'Era esperado erro para status invalido.');
  Logger.log(JSON.stringify({ ok: true, failed: failed }, null, 2));
}

function test_core_occupationCompat_headerAliases() {
  test_assert_(
    core_findOccupationHeaderIndexInHeaders_(['Nome', 'Cargo/Fun\u00E7\u00E3o'], 'occupation') === 1,
    'Alias legado de ocupacao deveria ser reconhecido.'
  );
  test_assert_(
    core_findOccupationHeaderIndexInHeaders_(['Nome', 'Ocupa\u00E7\u00E3o'], 'occupation') === 1,
    'Alias novo de ocupacao deveria ser reconhecido.'
  );
  test_assert_(
    core_findOccupationHeaderIndexInHeaders_(['Cargo/fun\u00E7\u00E3o atual'], 'currentOccupation') === 0,
    'Alias legado de ocupacao atual deveria ser reconhecido.'
  );
  test_assert_(
    core_findOccupationHeaderIndexInHeaders_(['Ocupa\u00E7\u00E3o atual'], 'currentOccupation') === 0,
    'Alias novo de ocupacao atual deveria ser reconhecido.'
  );

  Logger.log(JSON.stringify({ ok: true }, null, 2));
}

function test_core_occupationCompat_writePrefersOccupation_fakeSheet() {
  var sheet = test_createFakeSheet_([
    ['Cargo/fun\u00E7\u00E3o atual', 'Ocupa\u00E7\u00E3o atual'],
    ['Legado', 'Atual']
  ], 'MEMBERS_ATUAIS_FAKE');
  var headers = sheet.getRange(1, 1, 1, 2).getValues()[0];
  var headerMap = core_buildHeaderIndexMap_(headers, {
    normalize: true,
    oneBased: true,
    keepFirst: true
  });
  var wrote = core_writeOccupationValueByHeaderMap_(sheet, 2, headerMap, 'currentOccupation', 'Presidente');
  var row = sheet.getRange(2, 1, 1, 2).getValues()[0];

  test_assert_(wrote === true, 'A escrita de ocupacao atual deveria encontrar uma coluna compativel.');
  test_assert_(row[0] === 'Legado', 'A coluna legada nao deveria ser sobrescrita quando a nova existe.');
  test_assert_(row[1] === 'Presidente', 'A coluna nova de ocupacao atual deveria receber o valor.');

  Logger.log(JSON.stringify({ ok: true, row: row }, null, 2));
}

function test_core_rolesConfig_diretorComunicacao_legacyAliases() {
  var role = core_findInstitutionalRoleByAnyName_('Coordenador(a) de Comunicação');
  test_assert_(!!role, 'Alias legado da comunicação deveria resolver um cargo institucional.');
  test_assert_(role.roleKey === 'DIRETOR_COMUNICACAO', 'Alias legado deveria apontar para DIRETOR_COMUNICACAO.');

  var byLegacyKey = core_findInstitutionalRoleByAnyName_('COORDENADOR_COMUNICACAO');
  test_assert_(!!byLegacyKey, 'Cargo key legado da comunicação deveria ser aceito.');
  test_assert_(byLegacyKey.roleKey === 'DIRETOR_COMUNICACAO', 'Cargo key legado deveria apontar para DIRETOR_COMUNICACAO.');

  Logger.log(JSON.stringify({
    ok: true,
    publicName: role.publicName,
    roleKey: role.roleKey
  }, null, 2));
}
