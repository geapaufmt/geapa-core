function test_core_rolesConfig_debug() {
  Logger.log(JSON.stringify(core_debugInstitutionalRolesConfig_(), null, 2));
}

function test_core_exMembersCommunicationRecipients_debug() {
  Logger.log(JSON.stringify(core_debugExMembersCommunicationRecipients_(), null, 2));
}

function test_core_exMembersCommunicationRecipients_eixoI() {
  Logger.log(JSON.stringify(core_getExMembersCommunicationRecipients_({
    eixos: ['EIXO_I']
  }), null, 2));
}

function teste_getGeapaConfigValue_EMAIL_OFICIAL() {
  Logger.log(core_getGeapaConfigValue_('EMAIL_OFICIAL', {
    required: true
  }));
}

function teste_getGeapaConfigMap() {
  Logger.log(JSON.stringify(core_getGeapaConfigMap_(), null, 2));
}

function diagnosticarConfigGeapa() {
  Logger.log(JSON.stringify(core_debugGeapaConfig_(), null, 2));
}

function test_core_domainsV2_auditarPessoas() {
  Logger.log(JSON.stringify(coreAuditarPessoasV2_(), null, 2));
}

function test_core_domainsV2_auditarVigencias() {
  Logger.log(JSON.stringify(coreAuditarVigenciasV2_(), null, 2));
}

function test_core_domainsV2_auditarDominiosCentrais() {
  Logger.log(JSON.stringify(coreAuditarDominiosCentraisV2_(), null, 2));
}

function test_core_domainsV2_compararLegadoComV2() {
  Logger.log(JSON.stringify(coreCompararLegadoComV2_(), null, 2));
}

function test_core_domainsV2_recalcularVigenciasResumoAtual_dryRun() {
  Logger.log(JSON.stringify(coreRecalcularVigenciasResumoAtualV2_({
    dryRun: true
  }), null, 2));
}

function test_core_domainsV2_recalcularPessoasResumoOperacional_dryRun() {
  Logger.log(JSON.stringify(coreRecalcularPessoasResumoOperacionalV2_({
    dryRun: true
  }), null, 2));
}

function test_core_domainsV2_recalcularMembrosDetalhesSemestreAtual_dryRun() {
  Logger.log(JSON.stringify(coreRecalcularMembrosDetalhesSemestreAtualV2_({
    dryRun: true
  }), null, 2));
}

/**
 * Testa os consolidadores puros usados por PESSOAS_RESUMO_OPERACIONAL.
 *
 * @return {Object}
 */
function test_core_domainsV2_pessoasResumoOperationalHelpers() {
  var ctx = { idPessoa: 'PES-TESTE', rga: '202611801001', email: 'teste@example.com' };
  var activityData = {
    atividades: [{
      ID_ATIVIDADE: 'ATV-TESTE',
      ID_PESSOA_PRINCIPAL: ctx.idPessoa,
      RGA_PESSOA_PRINCIPAL: ctx.rga,
      EMAIL_PESSOA_PRINCIPAL: ctx.email,
      DATA_ATIVIDADE: '23/04/2026 18:30:00',
      ANO: 2026,
      SEMESTRE: 1,
      CICLO: 'GEAPA_2026'
    }],
    apresentacoes: [{
      ID_APRESENTACAO: 'APR-TESTE',
      ID_ATIVIDADE: 'ATV-TESTE',
      STATUS_APRESENTACAO: 'REALIZADA',
      STATUS_ENVIO_MATERIAL: 'RECEBIDO'
    }],
    portalAtividadesDetalhes: [],
    portalFrequencia: [{
      ID_PESSOA: ctx.idPessoa,
      CICLO: 'GEAPA_2026',
      TOTAL_PRESENCAS: 8,
      TOTAL_FALTAS: 2,
      TOTAL_JUSTIFICADAS: 1,
      FALTAS_LIQUIDAS: 1,
      PERCENTUAL_FREQUENCIA: 80,
      SITUACAO_DISCIPLINAR: 'REGULAR'
    }],
    presencas: [],
    justificativas: [{ ID_PESSOA: ctx.idPessoa, STATUS_ANALISE: '', ATIVO: 'SIM' }],
    portalJustificativas: []
  };
  var presentation = core_domainsV2PresentationSummary_(activityData, ctx);
  var frequency = core_domainsV2FrequencySummary_(activityData, ctx, 'GEAPA_2026');
  var pending = core_domainsV2PendingJustifications_(activityData, ctx);
  if (presentation.count !== 1 || presentation.lastCycle !== 'GEAPA_2026' || presentation.lastPeriod !== '2026/1') throw new Error('Resumo de apresentacoes incorreto.');
  if (frequency.indexOf('Presencas 8') < 0 || frequency.indexOf('Situacao REGULAR') < 0) throw new Error('Resumo de frequencia incorreto.');
  if (pending !== 1) throw new Error('Contagem de justificativas pendentes incorreta.');
  if (!core_domainsV2Date_('23/04/2026 18:30:00')) throw new Error('Parser de data brasileira com hora falhou.');
  var refDate = new Date(2026, 6, 6);
  var intervals = core_domainsV2EffectiveMemberIntervals_([
    { TIPO_VINCULO: 'MEMBRO_EFETIVO', DATA_INICIO: new Date(2026, 6, 1), DATA_FIM: new Date(2026, 6, 31), STATUS_VINCULO: 'ATIVO' },
    { TIPO_VINCULO: 'MEMBRO_EFETIVO', DATA_INICIO: new Date(2026, 7, 1), STATUS_VINCULO: 'ATIVO' },
    { TIPO_VINCULO: 'MEMBRO_EFETIVO', DATA_INICIO: new Date(2025, 0, 1), STATUS_VINCULO: 'ENCERRADO' }
  ], refDate);
  if (core_domainsV2CountIntervalDays_(intervals) !== 7) throw new Error('Tempo efetivo incluiu dias futuros ou prolongou vinculo encerrado sem fim.');
  var semesters = core_domainsV2CountSemestersForIntervals_({
    SEMESTRES: { records: [{ ID_SEMESTRE: '2026/1', DATA_INICIO: new Date(2026, 0, 1), DATA_FIM: new Date(2026, 6, 31) }] },
    CICLOS: { records: [{ ID_CICLO: 'GEAPA_2026', DATA_INICIO: new Date(2026, 0, 1), DATA_FIM: new Date(2026, 11, 31) }] }
  }, intervals);
  var semestersWithoutOfficialRows = core_domainsV2CountSemestersForIntervals_({
    SEMESTRES: { records: [] },
    CICLOS: { records: [{ ID_CICLO: 'GEAPA_2026', DATA_INICIO: new Date(2026, 0, 1), DATA_FIM: new Date(2026, 11, 31) }] }
  }, intervals);
  if (semesters !== 1 || semestersWithoutOfficialRows !== '') throw new Error('QTD_SEMESTRES_NO_GRUPO deve usar apenas SEMESTRES de Vigencias v2.');
  var result = { ok: true, presentation: presentation, frequency: frequency, pendingJustifications: pending };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function test_core_domainsV2_diagnosticarPessoasResumoOperacional() {
  Logger.log(JSON.stringify(coreDiagnosticarPessoasResumoOperacionalV2_(), null, 2));
}

function test_core_domainsV2_recalcularPessoasResumoOperacional_REAL_CONFIRMADO() {
  Logger.log(JSON.stringify(coreRecalcularPessoasResumoOperacionalV2_({
    dryRun: false,
    confirmacao: 'RECALCULAR_PESSOAS_RESUMO_V2'
  }), null, 2));
}

function test_core_domainsV2_recalcularMembrosDetalhesSemestreAtual_REAL_CONFIRMADO() {
  Logger.log(JSON.stringify(coreRecalcularMembrosDetalhesSemestreAtualV2_({
    dryRun: false,
    confirmacao: 'RECALCULAR_MEMBROS_DETALHES_SEMESTRE_ATUAL_V2'
  }), null, 2));
}

function test_core_domainsV2_pessoasListCurrentMembers() {
  Logger.log(JSON.stringify(corePessoasListCurrentMembers_(), null, 2));
}

/** Testa autorizacao, sanitizacao e filtros puros da listagem administrativa de membros. */
function test_core_portalAdminMembrosContratoPuro() {
  var allowed = core_pessoasAdminPortalAuthorize_({
    ok: true,
    autenticado: true,
    portalAtivo: true,
    permissoes: ['portal:acessar', 'membros:ler']
  });
  var denied = core_pessoasAdminPortalAuthorize_({
    ok: true,
    autenticado: true,
    portalAtivo: true,
    permissoes: ['portal:acessar']
  });
  var mapped = core_pessoasAdminPortalMapRow_({
    ID_PESSOA: 'PES-1',
    RGA: '2026001',
    NOME_EXIBICAO: 'Membro Teste',
    EMAIL: 'membro@example.com',
    TIPO_VINCULO_ATUAL: 'MEMBRO_EFETIVO',
    STATUS_VINCULO_ATUAL: 'ATIVO',
    CPF: '00000000000',
    TELEFONE: '65999999999',
    QTD_SEMESTRES_NO_GRUPO: 2,
    PENDENCIAS_ABERTAS: 'SEM_PENDENCIAS'
  });
  var matches = core_pessoasAdminPortalMatches_(mapped, core_pessoasAdminPortalFilters_({ texto: '2026001' }));
  if (!allowed) throw new Error('Diretoria/Secretaria/Admin com membros:ler deveria ser autorizada.');
  if (denied) throw new Error('Membro comum sem membros:ler deveria ser negado.');
  if (!matches) throw new Error('Busca por RGA deveria localizar o membro.');
  if (Object.prototype.hasOwnProperty.call(mapped, 'CPF') || Object.prototype.hasOwnProperty.call(mapped, 'telefone')) {
    throw new Error('Contrato administrativo vazou campo sensivel.');
  }
  return { ok: true, mapped: mapped };
}

/** Testa filtros e normalizacao puros do contrato geral de links/perfis. */
function test_core_domainsV2_linksPerfisHelpers() {
  var pessoasData = {
    PESSOAS_LINKS_PERFIS: {
      records: [
        { ID_LINK: 'LNK-000001', ID_PESSOA: 'PES-TESTE', TIPO_LINK: 'LATTES', URL: 'lattes.cnpq.br/123', ROTULO: '', PUBLICAVEL: 'SIM', VISIVEL_PORTAL: 'SIM', ATIVO: 'SIM' },
        { ID_LINK: 'LNK-000002', ID_PESSOA: 'PES-TESTE', TIPO_LINK: 'LINKEDIN', URL: 'https://linkedin.com/in/teste', ROTULO: 'Perfil profissional', PUBLICAVEL: 'SIM', VISIVEL_PORTAL: 'NAO', ATIVO: 'SIM' },
        { ID_LINK: 'LNK-000003', ID_PESSOA: 'PES-TESTE', TIPO_LINK: 'ORCID', URL: 'javascript:alert(1)', ROTULO: '', PUBLICAVEL: 'SIM', VISIVEL_PORTAL: 'SIM', ATIVO: 'SIM' },
        { ID_LINK: 'LNK-000004', ID_PESSOA: 'PES-TESTE', TIPO_LINK: 'OUTRO', URL: 'https://example.org/inativo', ROTULO: '', PUBLICAVEL: 'SIM', VISIVEL_PORTAL: 'SIM', ATIVO: 'NAO' }
      ]
    }
  };
  var privados = core_domainsV2LinksPerfisByPessoa_(pessoasData, 'PES-TESTE');
  var publicos = core_domainsV2LinksPerfisByPessoa_(pessoasData, 'PES-TESTE', { publicOnly: true });
  if (privados.length !== 2 || privados[0].tipo !== 'LATTES') throw new Error('Links ativos privados foram normalizados incorretamente.');
  if (publicos.length !== 1 || publicos[0].tipo !== 'LATTES') throw new Error('Filtro PUBLICAVEL/VISIVEL_PORTAL falhou.');
  return { ok: true, privados: privados, publicos: publicos };
}

/** Executa a migracao de Lattes somente em dry-run apos cadastrar a key no Registry. */
function test_core_domainsV2_migrarLinksPerfisLegados_dryRun() {
  return corePessoasV2MigrarLinksPerfisLegados_({ dryRun: true, ambiente: 'DEV' });
}

/** Confirma que Meu Perfil le exclusivamente a fonte geral de links de Pessoas V2. */
function test_core_portalMeuPerfilLinksPerfisV2() {
  var vazio = core_buildMeuPerfilLinksPortal_({
    pessoa: {},
    colaboradoresAcademicos: [{ LINK_LATTES: 'lattes.cnpq.br/legado' }],
    linksPerfis: []
  });
  var v2 = core_buildMeuPerfilLinksPortal_({
    pessoa: {},
    colaboradoresAcademicos: [{ LINK_LATTES: 'lattes.cnpq.br/legado' }],
    linksPerfis: [{ tipo: 'LATTES', url: 'https://lattes.cnpq.br/v2', rotulo: 'Lattes atualizado' }]
  });
  if (vazio.length !== 0) throw new Error('Meu Perfil nao deve consultar LINK_LATTES legado.');
  if (v2.length !== 1 || v2[0].url !== 'https://lattes.cnpq.br/v2') throw new Error('Link V2 deveria ser exibido.');
  return { ok: true, vazio: vazio, v2: v2 };
}

function test_core_domainsV2_recalcularVigenciasResumoAtual_REAL_CONFIRMADO() {
  Logger.log(JSON.stringify(coreRecalcularVigenciasResumoAtualV2_({
    dryRun: false,
    confirmacao: 'RECALCULAR_RESUMO_ATUAL_V2'
  }), null, 2));
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

function test_core_portalBuscarUsuario_buildProfilesAndPermissions() {
  var membro = Object.freeze({
    id: 'RGA-TESTE-001',
    nomeExibicao: 'Membro Teste',
    emailCadastrado: 'membro@exemplo.com',
    rga: 'RGA-TESTE-001'
  });

  var comum = core_buildUsuarioPortal_(membro, []);

  test_assert_(comum.ok !== false, 'Usuario comum nao deve retornar erro.');
  test_assert_(comum.perfilPrincipal === 'MEMBRO', 'Perfil principal de membro comum incorreto.');
  test_assert_(comum.perfis.join(',') === 'MEMBRO', 'Membro comum deve ter apenas perfil MEMBRO.');
  test_assert_(comum.cargosAtuais.length === 0, 'Membro comum nao deveria ter cargos atuais.');
  test_assert_(comum.permissoes.podeVerAreaDiretoria === false, 'Membro comum nao deve ver area da diretoria.');
  test_assert_(comum.permissoes.podeGerenciarAtividades === false, 'Membro comum nao deve gerenciar atividades.');

  var diretorSecretario = core_buildUsuarioPortal_(membro, [
    Object.freeze({
      sourceType: 'DIRETORIA',
      roleKey: 'SECRETARIO_GERAL',
      publicName: 'Secret\u00E1rio(a) Geral',
      roleConfig: Object.freeze({
        group: 'SECRETARIA'
      }),
      boardId: '2026-2027'
    })
  ]);

  test_assert_(diretorSecretario.perfilPrincipal === 'DIRETORIA', 'Secretaria em diretoria deve ter DIRETORIA como perfil principal operacional.');
  test_assert_(diretorSecretario.perfis.join(',') === 'MEMBRO,DIRETORIA,SECRETARIA', 'Perfis de secretario geral incorretos.');
  test_assert_(diretorSecretario.cargosAtuais.length === 1, 'Secretario deveria preservar um cargo atual.');
  test_assert_(diretorSecretario.cargosAtuais[0].cargoKey === 'SECRETARIO_GERAL', 'cargoKey do secretario incorreto.');
  test_assert_(diretorSecretario.cargosAtuais[0].grupoCargo === 'SECRETARIA', 'grupoCargo do secretario incorreto.');
  test_assert_(diretorSecretario.cargosAtuais[0].fonte === 'VIGENCIAS_DIRETORES', 'fonte do secretario incorreta.');
  test_assert_(diretorSecretario.cargosAtuais[0].idDiretoria === '2026-2027', 'idDiretoria deveria ser preservado.');
  test_assert_(diretorSecretario.permissoes.podeVerAreaDiretoria === true, 'Diretoria deve ver area da diretoria.');
  test_assert_(diretorSecretario.permissoes.podeGerenciarAtividades === true, 'Diretoria deve gerenciar atividades.');
  test_assert_(diretorSecretario.permissoes.podeRegistrarChamada === true, 'Diretoria deve registrar chamada.');
  test_assert_(diretorSecretario.permissoes.podeEditarAtividade === true, 'Diretoria deve editar atividade.');
  test_assert_(diretorSecretario.permissoes.podeAnalisarJustificativas === true, 'Diretoria deve analisar justificativas.');
  test_assert_(diretorSecretario.permissoes.podeGerenciarConfiguracoes === false, 'ADMIN_TECNICO nao deve ser derivado automaticamente.');

  var assessorComunicacao = core_buildUsuarioPortal_(membro, [
    Object.freeze({
      sourceType: 'ASSESSORIA',
      roleKey: 'ASSESSOR_COMUNICACAO',
      publicName: 'Assessor(a) de Comunica\u00E7\u00E3o'
    })
  ]);

  test_assert_(assessorComunicacao.perfis.join(',') === 'MEMBRO,ASSESSORIA,COMUNICACAO', 'Assessor de comunicacao deveria receber ASSESSORIA e COMUNICACAO.');
  test_assert_(assessorComunicacao.permissoes.podeGerenciarComunicacao === true, 'Comunicacao deve poder gerenciar comunicacao.');
  test_assert_(assessorComunicacao.permissoes.podeVerAreaDiretoria === false, 'Assessoria de comunicacao nao deve herdar area da diretoria sem DIRETORIA.');

  var conselheiro = core_buildUsuarioPortal_(membro, [
    Object.freeze({
      sourceType: 'CONSELHO',
      roleKey: 'CONSELHEIRO_CONSULTIVO',
      publicName: 'Conselheiro(a) Consultivo(a)'
    })
  ]);

  test_assert_(conselheiro.perfis.join(',') === 'MEMBRO,CONSELHO', 'Conselheiro consultivo deveria receber CONSELHO.');
  test_assert_(conselheiro.permissoes.podeVerAreaDiretoria === false, 'CONSELHO fica sem area da diretoria nesta etapa.');
}

function test_core_listarMembrosParaChamada_fakeSheet() {
  var sheet = test_createFakeSheet_([
    ['Membro', 'RGA', 'Status', 'Cargo/fun\u00E7\u00E3o atual', 'DATA_INTEGRACAO', 'DATA_DESLIGAMENTO', 'EMAIL', 'Telefone'],
    ['Membro Ativo', 'RGA-001', 'Ativo', 'Membro', '2026-01-10', '', 'ativo@example.com', '(00) 0000-0000'],
    ['Membro Futuro', 'RGA-002', 'Ativo', 'Membro', '2026-05-01', '', 'futuro@example.com', '(00) 0000-0001'],
    ['Membro Desligado', 'RGA-003', 'Ativo', 'Membro', '2025-01-01', '2026-03-01', 'desligado@example.com', '(00) 0000-0002'],
    ['Membro Suspenso', 'RGA-004', 'Suspenso', 'Membro', '2025-01-01', '', 'suspenso@example.com', '(00) 0000-0003']
  ], 'MEMBERS_ATUAIS_FAKE');

  var result = core_listarMembrosParaChamadaInSheet_(sheet, '2026-04-16', {
    perfil: 'DIRETORIA'
  }, {
    identityMaps: {},
    lifecycleByRga: {
      'RGA-005': [
        {
          eventId: 'MEV-000001',
          rga: 'RGA-005',
          eventType: 'INGRESSO',
          eventDate: new Date(2025, 0, 1),
          eventStatus: 'HOMOLOGADO',
          memberName: 'Membro Historico'
        },
        {
          eventId: 'MEV-000002',
          rga: 'RGA-005',
          eventType: 'DESLIGAMENTO_VOLUNTARIO',
          eventDate: new Date(2026, 5, 1),
          eventStatus: 'HOMOLOGADO',
          memberName: 'Membro Historico'
        }
      ]
    }
  });
  var denied = core_listarMembrosParaChamadaInSheet_(sheet, '2026-04-16', {
    perfil: 'MEMBRO'
  }, {
    identityMaps: {},
    lifecycleByRga: {}
  });
  var missingDate = core_listarMembrosParaChamadaInSheet_(sheet, '', {
    perfil: 'DIRETORIA'
  }, {
    identityMaps: {},
    lifecycleByRga: {}
  });
  var safeKeys = [
    'aplicavelNaData',
    'contaFalta',
    'contaPresenca',
    'email',
    'idPessoa',
    'motivoNaoAplicavel',
    'nomeExibicao',
    'rga',
    'situacao',
    'tipoParticipante',
    'vinculo'
  ].sort().join(',');

  test_assert_(result.ok === true, 'Lista de membros para chamada deveria retornar ok=true.');
  test_assert_(result.meta.total === 3, 'A chamada deveria incluir ativo, historico e suspenso/N/A, excluindo futuro e desligado.');
  test_assert_(result.meta.totalAplicaveis === 2, 'A chamada deveria ter dois membros aplicaveis.');
  test_assert_(result.meta.totalNaoAplicaveis === 1, 'A chamada deveria ter um membro nao aplicavel.');
  test_assert_(result.meta.cacheHit === false, 'Resultado direto em sheet fake nao deveria vir do cache.');
  test_assert_(result.meta.dataReferencia === '2026-04-16', 'Data de referencia deveria ser ISO local.');
  test_assert_(result.data[0].nomeExibicao === 'Membro Ativo', 'Ordenacao por nome incorreta.');
  test_assert_(result.data[0].email === 'ativo@example.com', 'Email seguro da chamada deveria ser normalizado.');
  test_assert_(result.data[0].situacao === 'ATIVO', 'Situacao ativa deveria ser normalizada.');
  test_assert_(result.data[0].contaPresenca === true, 'Membro ativo deve contar presenca.');
  test_assert_(result.data[0].contaFalta === true, 'Membro ativo deve contar falta.');
  test_assert_(result.data[1].nomeExibicao === 'Membro Historico', 'Membro historico deveria ser reconstruido pelo lifecycle.');
  test_assert_(result.data[1].situacao === 'ATIVO', 'Membro historico deveria estar ativo na data.');
  test_assert_(result.data[2].nomeExibicao === 'Membro Suspenso', 'Membro suspenso deveria aparecer como N/A seguro.');
  test_assert_(result.data[2].situacao === 'SUSPENSO', 'Situacao suspensa deveria ser normalizada.');
  test_assert_(result.data[2].aplicavelNaData === false, 'Suspenso deve aparecer como nao aplicavel na data.');
  test_assert_(result.data[2].contaPresenca === false, 'Suspenso nao deve contar presenca.');
  test_assert_(result.data[2].contaFalta === false, 'Suspenso nao deve contar falta.');
  test_assert_(result.data[2].motivoNaoAplicavel === 'SITUACAO_NAO_CONTABILIZA_CHAMADA', 'Motivo seguro de suspensao incorreto.');
  test_assert_(Object.keys(result.data[0]).sort().join(',') === safeKeys, 'Contrato retornou campos nao seguros.');
  test_assert_(!Object.prototype.hasOwnProperty.call(result.data[0], 'EMAIL'), 'Nao deve expor cabecalho bruto EMAIL.');
  test_assert_(!Object.prototype.hasOwnProperty.call(result.data[0], 'Telefone'), 'Nao deve expor telefone.');
  test_assert_(denied.ok === false && denied.errorCode === 'PERMISSAO_NEGADA', 'Perfil MEMBRO nao deve listar chamada quando contexto e informado.');
  test_assert_(missingDate.ok === false && missingDate.errorCode === 'DATA_ATIVIDADE_OBRIGATORIA', 'Data ausente deveria retornar erro controlado.');
}

function test_core_portalAccess_authorizeEmail_fakeSheets() {
  var profiles = [
    { PERFIL_PORTAL: 'MEMBRO', NOME: 'Membro', ATIVO: 'SIM' },
    { PERFIL_PORTAL: 'ADMIN', NOME: 'Administrador', ATIVO: 'SIM' }
  ];
  var permissions = [
    { PERFIL_PORTAL: 'MEMBRO', PERMISSAO: 'portal:acessar', ATIVO: 'SIM' },
    { PERFIL_PORTAL: 'MEMBRO', PERMISSAO: 'situacao:ver_propria', ATIVO: 'SIM' },
    { PERFIL_PORTAL: 'ADMIN', PERMISSAO: 'portal:acessar', ATIVO: 'SIM' },
    { PERFIL_PORTAL: 'ADMIN', PERMISSAO: 'sistema:admin', ATIVO: 'SIM' }
  ];
  var config = {
    PORTAL_MODO_ACESSO: 'MEMBROS_ATIVOS',
    PORTAL_EMAILS_TESTE: 'teste@example.com',
    PORTAL_BLOCK_INACTIVE_MEMBERS: 'SIM',
    PORTAL_DEFAULT_PROFILE: 'MEMBRO'
  };
  var currentSheet = test_createFakeSheet_([
    ['NOME', 'RGA', 'EMAIL', 'STATUS', 'PORTAL_ATIVO', 'PERFIL_PORTAL', 'PORTAL_OBS', 'TELEFONE'],
    ['Membro Ativo', 'RGA-001', 'ativo@example.com', 'ATIVO', 'SIM', 'MEMBRO', 'obs interna', '(00) 0000-0000'],
    ['Membro Bloqueado', 'RGA-002', 'bloqueado@example.com', 'ATIVO', 'NAO', 'MEMBRO', '', '(00) 0000-0001'],
    ['Admin Explicito', 'RGA-003', 'admin@example.com', 'ATIVO', 'SIM', 'ADMIN', '', '(00) 0000-0002']
  ], 'Membros Atuais');
  var formerSheet = test_createFakeSheet_([
    ['NOME', 'RGA', 'EMAIL', 'STATUS', 'PORTAL_ATIVO', 'PERFIL_PORTAL'],
    ['Ex Membro', 'RGA-004', 'ex@example.com', 'DESLIGADO', 'SIM', 'MEMBRO']
  ], 'Ex-Membros');
  var waitingSheet = test_createFakeSheet_([
    ['NOME', 'RGA', 'EMAIL', 'STATUS', 'PORTAL_ATIVO', 'PERFIL_PORTAL'],
    ['Membro Espera', 'RGA-005', 'espera@example.com', 'ESPERA', 'SIM', 'MEMBRO']
  ], 'Membros em Espera');
  var sources = [
    { type: 'MEMBROS_ATUAIS', sourceSheet: 'Membros Atuais', sheet: currentSheet },
    { type: 'MEMBROS_EM_ESPERA', sourceSheet: 'Membros em Espera', sheet: waitingSheet },
    { type: 'EX_MEMBROS', sourceSheet: 'Ex-Membros', sheet: formerSheet }
  ];
  var opts = {
    sources: sources,
    profiles: corePortalReadProfiles_({ records: profiles }),
    permissions: corePortalReadPermissions_({ records: permissions }),
    config: config
  };

  var active = corePortalAuthorizeEmail_(' Ativo@Example.com ', opts);
  var missing = corePortalAuthorizeEmail_('ausente@example.com', opts);
  var former = corePortalAuthorizeEmail_('ex@example.com', opts);
  var blocked = corePortalAuthorizeEmail_('bloqueado@example.com', opts);
  var admin = corePortalAuthorizeEmail_('admin@example.com', opts);
  var waiting = corePortalAuthorizeEmail_('espera@example.com', opts);
  var waitingAllowed = corePortalAuthorizeEmail_('espera@example.com', Object.assign({}, opts, {
    includeWaiting: true
  }));
  var testeBloqueado = corePortalAuthorizeEmail_('ativo@example.com', Object.assign({}, opts, {
    config: Object.assign({}, config, { PORTAL_MODO_ACESSO: 'TESTE' })
  }));
  var testeLiberado = corePortalAuthorizeEmail_('ativo@example.com', Object.assign({}, opts, {
    config: Object.assign({}, config, {
      PORTAL_MODO_ACESSO: 'TESTE',
      PORTAL_EMAILS_TESTE: 'ativo@example.com'
    })
  }));

  test_assert_(active.authorized === true, 'Membro ativo deveria ser autorizado.');
  test_assert_(active.modoAcesso === 'MEMBROS_ATIVOS', 'Modo MEMBROS_ATIVOS deveria aparecer na autorizacao.');
  test_assert_(active.authMode === 'EMAIL_LEGACY', 'Modo transitorio de auth incorreto.');
  test_assert_(active.email === 'ativo@example.com', 'Email deveria ser normalizado.');
  test_assert_(active.perfilPortal === 'MEMBRO', 'Perfil do membro ativo incorreto.');
  test_assert_(active.permissions.indexOf('situacao:ver_propria') >= 0, 'Permissao de membro ausente.');
  test_assert_(!Object.prototype.hasOwnProperty.call(active, 'rawRecord'), 'Sessao publica nao deve expor rawRecord.');
  test_assert_(!Object.prototype.hasOwnProperty.call(active, 'telefone'), 'Sessao publica nao deve expor telefone.');
  test_assert_(missing.authorized === false && missing.reason === 'PESSOA_NAO_ENCONTRADA', 'Email inexistente deveria bloquear.');
  test_assert_(former.authorized === false && former.reason === 'EX_MEMBRO_NAO_AUTORIZADO', 'Ex-membro deve bloquear por padrao.');
  test_assert_(blocked.authorized === false && blocked.reason === 'PORTAL_ATIVO_NAO', 'PORTAL_ATIVO=NAO deveria bloquear.');
  test_assert_(admin.authorized === true, 'ADMIN explicito e ativo deveria autorizar quando perfil/permissoes existem.');
  test_assert_(corePortalHasPermission_(admin, 'sistema:admin') === true, 'ADMIN deveria ter sistema:admin.');
  test_assert_(waiting.authorized === false && waiting.reason === 'MEMBRO_EM_ESPERA_NAO_AUTORIZADO', 'Espera deveria bloquear por padrao.');
  test_assert_(waitingAllowed.authorized === true, 'Espera deveria autorizar quando includeWaiting=true.');
  test_assert_(testeBloqueado.authorized === false && testeBloqueado.reason === 'EMAIL_FORA_PORTAL_EMAILS_TESTE', 'Modo TESTE deveria aplicar PORTAL_EMAILS_TESTE.');
  test_assert_(testeLiberado.authorized === true, 'Modo TESTE deveria autorizar e-mail listado em PORTAL_EMAILS_TESTE.');
}

function test_core_portalAccess_logAccess_fakeSheet() {
  var sheet = test_createFakeSheet_([
    ['TIMESTAMP', 'EMAIL', 'UID_FIREBASE', 'NOME', 'PERFIL_PORTAL', 'ACAO', 'RESULTADO', 'MOTIVO', 'ORIGEM', 'USER_AGENT', 'OBS']
  ], 'PORTAL_LOG_ACESSOS');

  corePortalAppendAccessLogToSheet_(sheet, {
    email: 'Pessoa@Example.com',
    uidFirebase: 'uid-futuro',
    nome: 'Pessoa Teste',
    perfilPortal: 'MEMBRO',
    acao: 'login',
    resultado: 'permitido',
    motivo: 'teste',
    origem: 'teste_unitario',
    userAgent: 'UnitTest',
    obs: 'sem segredo',
    token: 'NAO_DEVE_SER_REGISTRADO'
  });

  var row = sheet.getRange(2, 1, 1, 11).getValues()[0];

  test_assert_(row[1] === 'pessoa@example.com', 'Log deveria normalizar email.');
  test_assert_(row[2] === 'uid-futuro', 'UID futuro deveria ser registrado quando informado.');
  test_assert_(row[5] === 'LOGIN', 'Acao deveria ser normalizada.');
  test_assert_(row.join('|').indexOf('NAO_DEVE_SER_REGISTRADO') < 0, 'Log nao deve registrar token/segredo fora do contrato.');
}

function test_core_portalPublicContent_ensureStructure_fakeSpreadsheet() {
  var sheets = {
    PUBLIC_HOME: test_createFakeSheet_([
      ['ID_BLOCO', 'COLUNA_EXTRA'],
      ['home-1', 'valor manual']
    ], 'PUBLIC_HOME')
  };
  var fakeSpreadsheet = {
    getName: function() {
      return 'PORTAL_CONTEUDO_PUBLICO_FAKE';
    },
    getSheetByName: function(name) {
      return sheets[name] || null;
    },
    insertSheet: function(name) {
      sheets[name] = test_createFakeSheet_([], name);
      return sheets[name];
    }
  };

  var definitions = corePortalPublicContentGetDefinitions_();
  var first = corePortalPublicContentEnsureHeadersNoLock_({
    spreadsheet: fakeSpreadsheet,
    applyUx: false
  });
  var second = corePortalPublicContentEnsureHeadersNoLock_({
    spreadsheet: fakeSpreadsheet,
    applyUx: false
  });
  var homeHeaders = sheets.PUBLIC_HOME
    .getRange(1, 1, 1, sheets.PUBLIC_HOME.getLastColumn())
    .getValues()[0]
    .map(function(value) { return String(value || '').trim(); });
  var homeData = sheets.PUBLIC_HOME
    .getRange(2, 1, 1, 2)
    .getValues()[0];
  var homeHeaderCount = homeHeaders.filter(function(header) {
    return header === 'ID_BLOCO';
  }).length;

  test_assert_(definitions.sheets.length === 11, 'Deveria haver 11 abas editoriais definidas.');
  test_assert_(!!sheets.PUBLIC_CONFIG, 'PUBLIC_CONFIG deveria ter sido criada.');
  test_assert_(!!sheets.PUBLIC_PESSOAS_COMPLEMENTOS, 'PUBLIC_PESSOAS_COMPLEMENTOS deveria ter sido criada.');
  test_assert_(!!sheets.PUBLIC_GESTOES_COMPLEMENTOS, 'PUBLIC_GESTOES_COMPLEMENTOS deveria ter sido criada.');
  test_assert_(!!sheets.PUBLIC_PESSOAS_CONFIG, 'PUBLIC_PESSOAS_CONFIG deveria ter sido criada.');
  test_assert_(homeHeaders.indexOf('COLUNA_EXTRA') >= 0, 'Coluna extra existente deve ser preservada.');
  test_assert_(homeHeaders.indexOf('TIPO_BLOCO') >= 0, 'Coluna faltante deveria ser adicionada ao final.');
  test_assert_(homeHeaderCount === 1, 'Rodar duas vezes nao deve duplicar cabecalhos.');
  test_assert_(homeData[0] === 'home-1' && homeData[1] === 'valor manual', 'Dados existentes nao devem ser apagados.');
  test_assert_(first.actions.length === 11 && second.actions.length === 11, 'Relatorio deve cobrir todas as abas.');
}

function test_core_portalPublicContent_buildPublicSnapshot_fakeRecords() {
  var recordsByKey = {
    PORTAL_PUBLIC_HOME: [
      {
        ID_BLOCO: 'hero',
        TIPO_BLOCO: 'destaque',
        TITULO: 'Portal GEAPA',
        SUBTITULO: 'Publico',
        TEXTO: 'Texto <script>alert(1)</script> seguro',
        IMAGEM_URL: 'https://example.com/img.png',
        BOTAO_TEXTO: 'Saiba mais',
        BOTAO_URL: 'javascript:alert(1)',
        ORDEM: '2',
        STATUS_PUBLICACAO: 'PUBLICADO',
        PUBLICAR: 'SIM',
        ATIVO: 'SIM',
        ATUALIZADO_EM: '2026-06-01',
        EMAIL_INTERNO: 'nao-vaza@example.com'
      },
      {
        ID_BLOCO: 'rascunho',
        TIPO_BLOCO: 'destaque',
        TITULO: 'Rascunho',
        ORDEM: '1',
        STATUS_PUBLICACAO: 'RASCUNHO',
        PUBLICAR: 'SIM',
        ATIVO: 'SIM'
      }
    ],
    PORTAL_PUBLIC_SOBRE: [
      {
        ID_BLOCO: 'sobre-1',
        TITULO: 'Sobre',
        TEXTO: 'Descricao',
        ORDEM: '1',
        STATUS_PUBLICACAO: 'PUBLICADA',
        PUBLICAR: 'SIM',
        ATIVO: 'SIM',
        ATUALIZADO_EM: '2026-05-30'
      }
    ],
    PORTAL_PUBLIC_HISTORIA: [],
    PORTAL_PUBLIC_PARCEIROS: [
      {
        ID_PARCEIRO: 'p1',
        NOME: 'Parceiro',
        TIPO_PARCEIRO: 'institucional',
        DESCRICAO: 'Desc',
        LOGO_URL: 'https://example.com/logo.png',
        SITE_URL: 'https://example.com',
        INSTAGRAM_URL: '/instagram-local',
        ORDEM: '1',
        STATUS_PUBLICACAO: 'PUBLICADO',
        PUBLICAR: 'SIM',
        ATIVO: 'SIM',
        ATUALIZADO_EM: '2026-06-02'
      }
    ],
    PORTAL_PUBLIC_DOCUMENTOS: [
      {
        ID_DOCUMENTO: 'doc-1',
        TITULO: 'Regimento',
        TIPO_DOCUMENTO: 'pdf',
        VERSAO: '1',
        DATA_PUBLICACAO: '2026-01-15',
        DESCRICAO: 'Documento publico',
        URL_DOCUMENTO: 'https://example.com/doc.pdf',
        ORDEM: '1',
        STATUS_PUBLICACAO: 'PUBLICADO',
        PUBLICAR: 'SIM',
        ATIVO: 'SIM',
        ATUALIZADO_EM: '2026-06-03',
        OBS_INTERNA: 'nao deve vazar'
      }
    ],
    PORTAL_PUBLIC_CONFIG: [
      { KEY: 'TEMA', VALOR: 'claro', TIPO: 'texto', ATIVO: 'SIM' },
      { KEY: 'INATIVO', VALOR: 'nao', TIPO: 'texto', ATIVO: 'NAO' }
    ],
    PORTAL_PUBLIC_MIDIAS: [
      {
        ID_MIDIA: 'm1',
        NOME: 'Foto',
        TIPO: 'imagem',
        URL: 'https://example.com/foto.jpg',
        DESCRICAO: 'Foto publica',
        CATEGORIA: 'home',
        STATUS_PUBLICACAO: 'PUBLICADO',
        PUBLICAR: 'SIM',
        ATIVO: 'SIM',
        ATUALIZADO_EM: '2026-06-04'
      }
    ],
    PORTAL_PUBLIC_PESSOAS_COMPLEMENTOS: [
      {
        ID_PESSOA: 'P-1',
        GRUPO_PUBLICO: 'DIRETORIA',
        ID_DIRETORIA: '2026-2027',
        CARGO_PUBLICO: 'Presidencia',
        NOME_PUBLICO: 'Pessoa Publica',
        FOTO_URL: 'https://example.com/foto-pessoa.jpg',
        DESCRICAO_PUBLICA: 'Descricao publica de pessoa',
        PERIODO_PUBLICO: '2026-2027',
        DESTAQUE_HISTORICO: 'SIM',
        EXIBIR_NO_HISTORICO: 'SIM',
        AUTORIZACAO_PUBLICACAO: 'SIM',
        LINK_LATTES: 'https://lattes.cnpq.br/123',
        LINK_INSTAGRAM_PUBLICO: 'https://instagram.com/geapa',
        LINK_LINKEDIN_PUBLICO: 'javascript:alert(1)',
        ORDEM_PUBLICA: '1',
        STATUS_PUBLICACAO: 'PUBLICADO',
        PUBLICAR: 'SIM',
        ATIVO: 'SIM',
        ATUALIZADO_EM: '2026-06-05',
        EMAIL: 'nao-vaza@example.com'
      }
    ],
    PORTAL_PUBLIC_GESTOES_COMPLEMENTOS: [
      {
        ID_DIRETORIA: '2026-2027',
        NOME_PUBLICO_GESTAO: 'Gestao Publica',
        LEMA_PUBLICO: 'Lema',
        DESCRICAO_GESTAO: 'Descricao de gestao',
        FOTO_GESTAO_URL: 'https://example.com/gestao.jpg',
        LINK_ATA_POSSE_PUBLICA: 'https://example.com/ata.pdf',
        LINK_RELATORIO_GESTAO: 'https://example.com/relatorio.pdf',
        URL_GALERIA: 'https://example.com/galeria',
        ORDEM_PUBLICA: '1',
        STATUS_PUBLICACAO: 'PUBLICADA',
        PUBLICAR: 'SIM',
        ATIVO: 'SIM',
        ATUALIZADO_EM: '2026-06-06'
      }
    ],
    PORTAL_PUBLIC_PESSOAS_CONFIG: [
      { KEY: 'EXIBIR_MEMBROS_PUBLICOS', VALOR: 'SIM', TIPO: 'texto', ATIVO: 'SIM' }
    ],
    PORTAL_PUBLIC_DIRETORIA_COMPLEMENTOS: [
      {
        ID_PESSOA: 'P-1',
        ID_DIRETORIA: '2026-2027',
        FOTO_URL: 'https://example.com/foto-diretoria.jpg',
        DESCRICAO_PUBLICA: 'Descricao publica',
        LINK_LATTES: 'https://lattes.cnpq.br/123',
        LINK_INSTAGRAM_PUBLICO: 'https://instagram.com/geapa',
        ORDEM_PUBLICA: '1',
        STATUS_PUBLICACAO: 'PUBLICADO',
        PUBLICAR: 'SIM',
        ATIVO: 'SIM',
        ATUALIZADO_EM: '2026-06-05',
        EMAIL: 'nao-vaza@example.com'
      }
    ],
    PORTAL_PUBLIC_LOG_PUBLICACAO: [
      { ID_LOG: 'L1', EXECUTADO_POR: 'Pessoa Interna', STATUS: 'OK' }
    ]
  };
  var snapshot = corePortalPublicContentBuildPublicSnapshot_({
    recordsByKey: recordsByKey,
    dryRun: true,
    disableCache: true
  });
  var logRead = corePortalPublicContentReadRows_('PORTAL_PUBLIC_LOG_PUBLICACAO', {
    recordsByKey: recordsByKey,
    dryRun: true,
    disableCache: true
  });

  test_assert_(snapshot.ok === true, 'Snapshot publico deveria montar ok=true.');
  test_assert_(snapshot.data.pages.home.blocos.length === 1, 'Rascunho nao deve aparecer no home.');
  test_assert_(snapshot.data.pages.home.blocos[0].botaoUrl === '', 'URL insegura deveria ser removida.');
  test_assert_(snapshot.data.pages.home.blocos[0].texto.indexOf('<script>') < 0, 'Script tag deveria ser removida.');
  test_assert_(!Object.prototype.hasOwnProperty.call(snapshot.data.pages.home.blocos[0], 'EMAIL_INTERNO'), 'Coluna extra nao deve vazar.');
  test_assert_(snapshot.data.documents.length === 1, 'Documento publicado deveria aparecer.');
  test_assert_(!Object.prototype.hasOwnProperty.call(snapshot.data.documents[0], 'OBS_INTERNA'), 'Observacao interna nao deve vazar.');
  test_assert_(snapshot.data.config.TEMA === 'claro', 'Config ativa deveria aparecer.');
  test_assert_(!Object.prototype.hasOwnProperty.call(snapshot.data.config, 'INATIVO'), 'Config inativa nao deve aparecer.');
  test_assert_(snapshot.data.peopleComplements[0].idPessoa === 'P-1', 'Complemento publico de pessoa deveria aparecer.');
  test_assert_(snapshot.data.peopleComplements[0].linkLinkedinPublico === '', 'URL insegura de pessoa deveria ser removida.');
  test_assert_(!Object.prototype.hasOwnProperty.call(snapshot.data.peopleComplements[0], 'EMAIL'), 'Email nao deve vazar em complemento de pessoa.');
  test_assert_(snapshot.data.managementComplements[0].idDiretoria === '2026-2027', 'Complemento publico de gestao deveria aparecer.');
  test_assert_(snapshot.data.peopleConfig.EXIBIR_MEMBROS_PUBLICOS === 'SIM', 'Config publica de pessoas deveria aparecer.');
  test_assert_(snapshot.data.boardComplements[0].ID_PESSOA === 'P-1', 'Complemento legado de diretoria deveria continuar disponivel.');
  test_assert_(!Object.prototype.hasOwnProperty.call(snapshot.data.boardComplements[0], 'EMAIL'), 'Email nao deve vazar em complemento legado de diretoria.');
  test_assert_(logRead.ok === false && logRead.errorCode === 'PORTAL_PUBLIC_KEY_NAO_EXPONIVEL', 'Log de publicacao nao deve ser exposto por leitura publica.');
}

function test_core_portalV2_diagnosticarPerfisEPermissoes() {
  Logger.log(JSON.stringify(corePortalDiagnosticarPerfisEPermissoes_(), null, 2));
}

function test_core_portalV2_prepararPortalParaV2() {
  Logger.log(JSON.stringify(corePrepararPortalParaV2_(), null, 2));
}

function test_core_portalConfig_normalizacao_fakeRecords() {
  var config = corePortalReadConfig_({
    records: [
      { CHAVE: 'ATIVIDADES_CHAMADA_ANTECEDENCIA_MINUTOS', VALOR: '60', ATIVO: 'SIM', DESCRICAO: 'Numero' },
      { CHAVE: 'ATIVIDADES_PRELOAD_DETALHES', VALOR: 'SIM', ATIVO: 'SIM', DESCRICAO: 'Boolean' },
      { CHAVE: 'PORTAL_TEMA', VALOR: 'claro', ATIVO: 'SIM', DESCRICAO: 'Texto' },
      { CHAVE: 'CONFIG_INATIVA', VALOR: 'SIM', ATIVO: 'NAO', DESCRICAO: 'Ignorar' },
      { CHAVE: 'API_SECRET', VALOR: 'NAO_DEVE_SAIR', ATIVO: 'SIM', DESCRICAO: 'Sensivel' }
    ]
  });
  test_assert_(config.ATIVIDADES_CHAMADA_ANTECEDENCIA_MINUTOS === 60, 'Numero deveria ser normalizado.');
  test_assert_(config.ATIVIDADES_PRELOAD_DETALHES === true, 'SIM deveria virar boolean true.');
  test_assert_(config.PORTAL_TEMA === 'claro', 'Texto deveria ser preservado.');
  test_assert_(!Object.prototype.hasOwnProperty.call(config, 'CONFIG_INATIVA'), 'Linha inativa nao deveria sair.');
  test_assert_(!Object.prototype.hasOwnProperty.call(config, 'API_SECRET'), 'Chave sensivel nao deveria sair.');
  Logger.log(JSON.stringify(config, null, 2));
}

function test_core_portalConfig_lerOperacional() {
  Logger.log(JSON.stringify(corePortalReadConfig_({ forceRefresh: true }), null, 2));
}

function test_core_portalV2_resolverUsuarioAtual_emailExemplo() {
  Logger.log(JSON.stringify(corePortalResolverUsuarioAtual_('email@exemplo.com'), null, 2));
}

function test_core_portalV2_resolverUsuarioAtual_entradaObjeto() {
  Logger.log(JSON.stringify(corePortalResolverUsuarioAtual_({
    email: 'email@exemplo.com'
  }), null, 2));
}

function test_core_portalV2_resolverUsuarioAtual_rgaConfigurado() {
  var rga = PropertiesService
    .getScriptProperties()
    .getProperty('GEAPA_CORE_PORTAL_TESTE_RGA');
  if (!String(rga || '').trim()) {
    throw new Error('Configure GEAPA_CORE_PORTAL_TESTE_RGA com um RGA existente em Pessoas v2.');
  }
  Logger.log(JSON.stringify(corePortalResolverUsuarioAtual_({ rga: rga }), null, 2));
}

function test_core_portalV2_resolverUsuarioAtual_idPessoaConfigurado() {
  var idPessoa = PropertiesService
    .getScriptProperties()
    .getProperty('GEAPA_CORE_PORTAL_TESTE_ID_PESSOA');
  if (!String(idPessoa || '').trim()) {
    throw new Error('Configure GEAPA_CORE_PORTAL_TESTE_ID_PESSOA com um ID_PESSOA existente em Pessoas v2.');
  }
  Logger.log(JSON.stringify(corePortalResolverUsuarioAtual_({ idPessoa: idPessoa }), null, 2));
}

function test_core_portalV2_buscarMembroLegado_identificadorConfigurado() {
  var identificador = PropertiesService
    .getScriptProperties()
    .getProperty('GEAPA_CORE_PORTAL_TESTE_IDENTIFICADOR');
  if (!String(identificador || '').trim()) {
    throw new Error('Configure GEAPA_CORE_PORTAL_TESTE_IDENTIFICADOR com e-mail, RGA ou ID_PESSOA.');
  }
  Logger.log(JSON.stringify(geapaCoreBuscarMembroParaPortal(identificador), null, 2));
}

function test_core_portalV2_buscarUsuarioLegado_identificadorConfigurado() {
  var identificador = PropertiesService
    .getScriptProperties()
    .getProperty('GEAPA_CORE_PORTAL_TESTE_IDENTIFICADOR');
  if (!String(identificador || '').trim()) {
    throw new Error('Configure GEAPA_CORE_PORTAL_TESTE_IDENTIFICADOR com e-mail, RGA ou ID_PESSOA.');
  }
  Logger.log(JSON.stringify(geapaCoreBuscarUsuarioPortal(identificador), null, 2));
}

function test_core_portalFirestore_diagnosticarEscoposOAuth() {
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo', {
    method: 'post',
    payload: { access_token: token },
    muteHttpExceptions: true
  });
  var body = JSON.parse(response.getContentText() || '{}');
  var scopes = String(body.scope || '').split(/\s+/).filter(String).sort();
  Logger.log(JSON.stringify({
    ok: response.getResponseCode() >= 200 && response.getResponseCode() < 300,
    httpStatus: response.getResponseCode(),
    hasDatastore: scopes.indexOf('https://www.googleapis.com/auth/datastore') >= 0,
    hasCloudPlatform: scopes.indexOf('https://www.googleapis.com/auth/cloud-platform') >= 0,
    scopes: scopes
  }, null, 2));
}

function test_core_firestoreRest_encoderEDryRun() {
  var encoded = coreFirestoreEncodeDocument_({
    texto: 'ok',
    ativo: true,
    total: 2,
    badges: ['A', 'B'],
    flags: { publicado: true }
  });
  test_assert_(encoded.fields.texto.stringValue === 'ok', 'Firestore deveria codificar string.');
  test_assert_(encoded.fields.ativo.booleanValue === true, 'Firestore deveria codificar boolean.');
  test_assert_(encoded.fields.total.integerValue === '2', 'Firestore deveria codificar inteiro.');
  test_assert_(encoded.fields.badges.arrayValue.values.length === 2, 'Firestore deveria codificar array.');
  test_assert_(encoded.fields.flags.mapValue.fields.publicado.booleanValue === true, 'Firestore deveria codificar mapa.');

  var dryRun = coreFirestoreSetDocument_('portalActivities/ATV-2026-1-0001', {
    idAtividade: 'ATV-2026-1-0001'
  }, {
    dryRun: true,
    projectId: 'projeto-teste'
  });
  test_assert_(dryRun.ok === true && dryRun.written === false, 'Dry-run Firestore nao deveria escrever.');
  return { ok: true, encoded: encoded, dryRun: dryRun };
}

function test_core_firestoreRest_validarCaminhosDeColecao() {
  var collection = coreFirestoreNormalizeCollectionPath_('portalActivities');
  var nested = coreFirestoreNormalizeCollectionPath_('parents/P1/children');
  test_assert_(collection.length === 1, 'Colecao raiz deveria ser aceita.');
  test_assert_(nested.length === 3, 'Colecao aninhada deveria ser aceita.');
  var rejected = false;
  try { coreFirestoreNormalizeCollectionPath_('portalActivities/ATV-1'); } catch (err) { rejected = true; }
  test_assert_(rejected, 'Caminho de documento nao deveria ser aceito como colecao.');
  return { ok: true };
}
function test_core_portalFirestore_syncUsuarioEmailConfigurado() {
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty('GEAPA_CORE_PORTAL_TESTE_FIRESTORE_EMAIL') || '';
  var uid = props.getProperty('GEAPA_CORE_PORTAL_TESTE_FIRESTORE_UID') || '';
  if (!String(email || '').trim()) {
    throw new Error('Configure GEAPA_CORE_PORTAL_TESTE_FIRESTORE_EMAIL com um e-mail existente em Pessoas v2.');
  }
  if (!String(uid || '').trim()) {
    throw new Error('Configure GEAPA_CORE_PORTAL_TESTE_FIRESTORE_UID com o uid Firebase desse usuario.');
  }
  if (props.getProperty('GEAPA_CORE_PORTAL_TESTE_FIRESTORE_DEV_WRITE_AUTHORIZED') !== 'SIM') {
    throw new Error('Write remoto DEV nao autorizado. Configure GEAPA_CORE_PORTAL_TESTE_FIRESTORE_DEV_WRITE_AUTHORIZED=SIM somente apos aprovacao explicita.');
  }

  var resultado = corePortalProvisionarFirestoreUserAutenticado({
    uid: uid,
    email: email,
    emailVerified: true
  }, {
    ambiente: 'DEV',
    uid: uid,
    identityVerified: true,
    dryRun: false
  });
  Logger.log(JSON.stringify(resultado, null, 2));
  var detalheErro = resultado.firestoreError
    ? ' Detalhe Firestore: ' + JSON.stringify(resultado.firestoreError)
    : '';
  test_assert_(resultado.ok === true, 'Sync Firestore deveria retornar ok=true.' + detalheErro);
  test_assert_(resultado.synced === true, 'Sync Firestore deveria gravar portalUsers/{uid}.' + detalheErro);
  test_assert_(resultado.code === 'FIRESTORE_SYNC_OK', 'Sync Firestore deveria retornar FIRESTORE_SYNC_OK.' + detalheErro);
}

function test_core_portalFirestore_snapshotSeguroMinimo() {
  var snapshot = corePortalBuildFirestoreUserSnapshot_({ uid: 'uid-teste-123' }, {
    sessao: {
      ok: true,
      autenticado: true,
      idPessoa: 'P-TESTE',
      nomeExibicao: 'Pessoa Teste',
      email: 'pessoa@example.org',
      rga: 'DADO-NAO-PERMITIDO',
      portalAtivo: true,
      perfilPortalEfetivo: 'MEMBRO',
      perfisPortal: ['MEMBRO'],
      permissoes: ['portal:acessar']
    },
    cacheUpdatedAt: '2026-07-04T12:00:00.000Z'
  });
  test_assert_(snapshot.uid === 'uid-teste-123', 'Snapshot deve preservar UID validado.');
  test_assert_(snapshot.ativo === true && snapshot.podeAcessarPortal === true, 'Snapshot deve liberar apenas sessao ativa.');
  test_assert_(snapshot.permissions['portal:acessar'] === true, 'Snapshot deve mapear permissoes no backend.');
  test_assert_(snapshot.source === 'PESSOAS_V2' && snapshot.sourceSystem === 'geapa-core', 'Snapshot deve identificar fonte oficial.');
  test_assert_(!Object.prototype.hasOwnProperty.call(snapshot, 'rga'), 'Snapshot nao deve gravar RGA.');
  test_assert_(!Object.prototype.hasOwnProperty.call(snapshot, 'cpf'), 'Snapshot nao deve gravar CPF.');
  return { ok: true, snapshot: snapshot };
}

function test_core_portalFirestore_negarIdentidadeNaoVerificada() {
  var result = corePortalProvisionarFirestoreUserAutenticado_({
    uid: 'uid-teste-123',
    email: 'pessoa@example.org',
    emailVerified: true
  }, {});
  test_assert_(result.ok === false, 'Provisionamento deve negar identidade nao marcada como verificada pelo backend.');
  test_assert_(result.code === 'IDENTIDADE_FIREBASE_NAO_VERIFICADA', 'Codigo de negacao inesperado.');
  return result;
}

function test_core_portalFirestore_lerSnapshotUidConfigurado() {
  var uid = PropertiesService
    .getScriptProperties()
    .getProperty('GEAPA_CORE_PORTAL_TESTE_FIRESTORE_UID') || '';
  if (!String(uid || '').trim()) {
    throw new Error('Configure GEAPA_CORE_PORTAL_TESTE_FIRESTORE_UID com o uid Firebase desse usuario.');
  }

  var resultado = corePortalReadFirestoreUserSnapshotByUid_(uid, { ambiente: 'DEV' });
  Logger.log(JSON.stringify(resultado, null, 2));
  var detalheErro = resultado.firestoreError
    ? ' Detalhe Firestore: ' + JSON.stringify(resultado.firestoreError)
    : '';
  test_assert_(resultado.ok === true, 'Leitura Firestore deveria retornar ok=true.' + detalheErro);
  test_assert_(resultado.found === true, 'Leitura Firestore deveria encontrar portalUsers/{uid}.' + detalheErro);
  test_assert_(resultado.snapshot && resultado.snapshot.uid === uid, 'Snapshot Firestore deveria ter uid esperado.');
  test_assert_(resultado.snapshot.source === 'PESSOAS_V2', 'Snapshot Firestore deveria preservar source oficial.');
  test_assert_(resultado.snapshot.schemaVersion === 'portal-user-v2', 'Snapshot Firestore deveria preservar schemaVersion.');
}
function test_core_portalV2_diagnosticarSessoesPortal() {
  var raw = PropertiesService
    .getScriptProperties()
    .getProperty('GEAPA_CORE_PORTAL_TESTE_SESSOES');
  var entradas = raw ? JSON.parse(raw) : {
    MEMBRO: '',
    DIRETORIA: '',
    SECRETARIA: '',
    COMUNICACAO: '',
    CONSELHO: '',
    EGRESSO: '',
    COLABORADOR: '',
    EXTERNO: '',
    VISITANTE: '',
    ADMIN: '',
    SEM_PORTAL_ATIVO: '',
    EXCECAO_ACESSO: '',
    NAO_ENCONTRADO: 'usuario-nao-encontrado@example.invalid'
  };
  var out = {};
  Object.keys(entradas).forEach(function(label) {
    var entrada = entradas[label];
    out[label] = entrada
      ? corePortalResolverUsuarioAtual_(entrada)
      : { ok: false, skipped: true, motivo: 'ENTRADA_NAO_CONFIGURADA' };
  });
  Logger.log(JSON.stringify(out, null, 2));
}

function test_core_portalV2_minhaSituacaoV2_identificadorConfigurado() {
  var identificador = PropertiesService
    .getScriptProperties()
    .getProperty('GEAPA_CORE_PORTAL_TESTE_IDENTIFICADOR');
  if (!String(identificador || '').trim()) {
    throw new Error('Configure GEAPA_CORE_PORTAL_TESTE_IDENTIFICADOR com e-mail, RGA ou ID_PESSOA.');
  }
  Logger.log(JSON.stringify(core_buscarMinhaSituacaoParaPortal_(identificador), null, 2));
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
      'CICLO_ULTIMA_APRESENTACAO',
      'QTD_APRESENTACOES_REALIZADAS',
      'QTD_DIAS_QUE_CONTAM_PARA_LIMITE_DIRETORIA',
      'LIMITE_DIAS_DIRETORIA',
      'SALDO_DIAS_DIRETORIA',
      'STATUS_ELEGIBILIDADE_DIRETORIA',
      'DATA_LIMITE_ESTIMADA_DIRETORIA'
    ],
    ['MEM-TESTE-001', 'Membro Sem Pendencia', 'sem-pendencia@exemplo.com', 'RGA-TESTE-001', 'Ativo', 'Membro', '(00) 0000-0000', 'GEAPA_2025', '2', '0', '549', '549', 'APTO', '18/05/2027'],
    ['MEM-TESTE-002', 'Membro Sem RGA', 'sem-rga@exemplo.com', '', 'Ativo', 'Membro', '(00) 0000-0001', '', '', '120', '549', '429', 'APTO_COM_LIMITE', '01/01/2027'],
    ['MEM-TESTE-003', '', 'sem-nome@exemplo.com', 'RGA-TESTE-003', 'Ativo', 'Membro', '(00) 0000-0002', 'GEAPA_2024', '1', '', '', '', '', ''],
    ['MEM-TESTE-004', 'Membro Indefinido', 'indefinido@exemplo.com', 'RGA-TESTE-004', 'Indefinido', 'Indefinido', '(00) 0000-0003', 'GEAPA_2025', 'invalido', 'abc', '-10', 'nao-numero', 'APTO_COM_LIMITE', ''],
    ['MEM-TESTE-005', 'Membro Email Invalido', 'email-invalido', 'RGA-TESTE-005', 'Ativo', 'Membro', '(00) 0000-0004', '', '', '', '', '', '', '']
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
  test_assert_(result.minhaSituacao.diretoria.statusElegibilidade === 'APTO', 'Status de elegibilidade APTO incorreto.');
  test_assert_(result.minhaSituacao.diretoria.diasComputados === 0, 'Dias computados APTO deveriam ser 0.');
  test_assert_(result.minhaSituacao.diretoria.limiteDias === 549, 'Limite de dias APTO incorreto.');
  test_assert_(result.minhaSituacao.diretoria.saldoDias === 549, 'Saldo de dias APTO incorreto.');
  test_assert_(result.minhaSituacao.diretoria.dataLimiteEstimada === '18/05/2027', 'Data limite estimada deveria preservar o texto da planilha.');
  test_assert_(Array.isArray(result.minhaSituacao.certificados) && result.minhaSituacao.certificados.length === 0, 'certificados deve iniciar vazio.');
  test_assert_(Array.isArray(result.minhaSituacao.avisos) && result.minhaSituacao.avisos.length === 0, 'avisos deve iniciar vazio.');
  test_assert_(semRga.minhaSituacao.pendencias.length === 1, 'Membro sem RGA deveria ter uma pendencia.');
  test_assert_(semRga.minhaSituacao.pendencias[0].titulo === 'RGA nao informado', 'Pendencia de RGA incorreta.');
  test_assert_(semRga.minhaSituacao.resumo.pendenciasAbertas === semRga.minhaSituacao.pendencias.length, 'Contagem de pendencias sem RGA incorreta.');
  test_assert_(semRga.minhaSituacao.participacao.apresentacoes.quantidadeRealizadas === 0, 'Membro sem apresentacoes atuais deveria retornar quantidade 0.');
  test_assert_(semRga.minhaSituacao.participacao.apresentacoes.periodoUltimaApresentacao === '', 'Membro sem apresentacoes atuais deveria retornar periodo vazio.');
  test_assert_(semRga.minhaSituacao.diretoria.statusElegibilidade === 'APTO_COM_LIMITE', 'Status APTO_COM_LIMITE incorreto.');
  test_assert_(semRga.minhaSituacao.diretoria.diasComputados === 120, 'Dias computados APTO_COM_LIMITE incorretos.');
  test_assert_(semRga.minhaSituacao.diretoria.limiteDias === 549, 'Limite APTO_COM_LIMITE incorreto.');
  test_assert_(semRga.minhaSituacao.diretoria.saldoDias === 429, 'Saldo APTO_COM_LIMITE incorreto.');
  test_assert_(semNome.minhaSituacao.pendencias.length === 1, 'Membro sem nome deveria ter uma pendencia.');
  test_assert_(semNome.minhaSituacao.pendencias[0].titulo === 'Nome de exibicao nao informado', 'Pendencia de nome incorreta.');
  test_assert_(semNome.minhaSituacao.resumo.pendenciasAbertas === semNome.minhaSituacao.pendencias.length, 'Contagem de pendencias sem nome incorreta.');
  test_assert_(semNome.minhaSituacao.participacao.apresentacoes.periodoUltimaApresentacao === 'GEAPA_2024', 'Periodo consolidado incorreto.');
  test_assert_(semNome.minhaSituacao.participacao.apresentacoes.quantidadeRealizadas === 1, 'Quantidade consolidada deveria ser numero.');
  test_assert_(semNome.minhaSituacao.diretoria.statusElegibilidade === '', 'Status vazio deveria retornar string vazia.');
  test_assert_(semNome.minhaSituacao.diretoria.diasComputados === 0, 'Dias vazios deveriam retornar 0.');
  test_assert_(semNome.minhaSituacao.diretoria.limiteDias === 0, 'Limite vazio deveria retornar 0.');
  test_assert_(semNome.minhaSituacao.diretoria.saldoDias === 0, 'Saldo vazio deveria retornar 0.');
  test_assert_(semNome.minhaSituacao.diretoria.dataLimiteEstimada === '', 'Data vazia deveria retornar string vazia.');
  test_assert_(indefinido.minhaSituacao.pendencias.length === 2, 'Membro com situacao/vinculo indefinido deveria ter duas pendencias.');
  test_assert_(indefinido.minhaSituacao.pendencias[0].titulo === 'Vinculo cadastral indefinido', 'Pendencia de vinculo incorreta.');
  test_assert_(indefinido.minhaSituacao.pendencias[1].titulo === 'Situacao geral indefinida', 'Pendencia de situacao geral incorreta.');
  test_assert_(indefinido.minhaSituacao.resumo.pendenciasAbertas === indefinido.minhaSituacao.pendencias.length, 'Contagem de pendencias indefinidas incorreta.');
  test_assert_(indefinido.minhaSituacao.participacao.apresentacoes.quantidadeRealizadas === 0, 'Quantidade atual invalida deveria virar 0.');
  test_assert_(indefinido.minhaSituacao.diretoria.diasComputados === 0, 'Dias invalidos deveriam retornar 0.');
  test_assert_(indefinido.minhaSituacao.diretoria.limiteDias === 0, 'Limite negativo deveria retornar 0.');
  test_assert_(indefinido.minhaSituacao.diretoria.saldoDias === 0, 'Saldo invalido deveria retornar 0.');
  test_assert_(indefinido.minhaSituacao.diretoria.dataLimiteEstimada === '', 'Data vazia em diretoria deveria retornar string vazia.');
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
