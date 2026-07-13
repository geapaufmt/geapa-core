/** Testes unitarios dos contratos de atualizacao cadastral. Nao escrevem em planilhas reais. */

function corePerfilTestAssert_(condition, message) {
  if (!condition) throw new Error('TESTE_FALHOU: ' + message);
}

function corePerfilTestFixture_(permissions) {
  var now = new Date(2026, 6, 12, 10, 0, 0);
  var sources = {};
  function source(name, headers, records) {
    sources[name] = {
      sheet: { name: name },
      headers: headers.slice(),
      records: (records || []).map(function(record, index) {
        return Object.assign({ __rowNumber: index + 2 }, record);
      })
    };
  }
  source('PESSOAS_BASE', [
    'ID_PESSOA','NOME_COMPLETO','NOME_CIVIL','EMAIL_PRINCIPAL','TELEFONE_PRINCIPAL','CPF',
    'DATA_NASCIMENTO','INSTAGRAM','CIDADE_NATAL','UF_ORIGEM','ATUALIZADO_EM','ATIVO'
  ], [{
    ID_PESSOA: 'PES-1', NOME_COMPLETO: 'Pessoa Teste', NOME_CIVIL: '', EMAIL_PRINCIPAL: 'pessoa@example.org',
    TELEFONE_PRINCIPAL: '+5565999990000', CPF: '52998224725', DATA_NASCIMENTO: '1990-01-10',
    INSTAGRAM: '', CIDADE_NATAL: 'Cuiaba', UF_ORIGEM: 'MT', ATIVO: 'SIM'
  }]);
  source('MEMBROS_DETALHES', ['ID_PESSOA','RGA','HISTORICO_ATIVIDADES_ACADEMICAS','ATUALIZADO_EM'], [{
    ID_PESSOA: 'PES-1', RGA: '20201234', HISTORICO_ATIVIDADES_ACADEMICAS: 'Resumo anterior'
  }]);
  source('PESSOAS_V2_LINKS_PERFIS', [
    'ID_LINK','ID_PESSOA','TIPO_LINK','URL','ROTULO','PUBLICAVEL','VISIVEL_PORTAL','FONTE',
    'VALIDADO_EM','ATIVO','CRIADO_EM','ATUALIZADO_EM','OBS'
  ], []);
  source(CORE_PERFIL_SOLICITACOES_SHEET, CORE_DOMAINS_V2_SCHEMAS.PESSOAS.filter(function(item) {
    return item.sheetName === CORE_PERFIL_SOLICITACOES_SHEET;
  })[0].headers, []);

  var counters = { writes: 0, appends: 0, recalculations: 0 };
  var uuidCounter = 0;
  var session = {
    ok: true,
    autenticado: true,
    portalAtivo: true,
    idPessoa: 'PES-1',
    email: 'pessoa@example.org',
    permissoes: permissions || ['portal:acessar'],
    perfisPortal: []
  };
  var deps = {
    environment: 'DEV',
    session: session,
    hasDevPermission: function(currentSession, permission) {
      return (currentSession.permissoes || []).indexOf(permission) >= 0;
    },
    now: function() { return new Date(now.getTime()); },
    uuid: function() { uuidCounter++; return '00000000-0000-4000-8000-' + ('000000000000' + uuidCounter).slice(-12); },
    hash: function(value) { return 'HASH_' + String(value).replace(/[^A-Za-z0-9]/g, '').slice(0, 30); },
    withLock: function(key, fn) { return fn(); },
    openPessoas: function() { return sources; },
    writeRecord: function(sourceData, record) {
      counters.writes++;
      var index = sourceData.records.findIndex(function(item) { return item.__rowNumber === record.__rowNumber; });
      if (index < 0) throw new Error('FAKE_ROW_NAO_ENCONTRADA');
      sourceData.records[index] = Object.assign({}, record);
    },
    appendRecord: function(sourceData, record) {
      counters.appends++;
      sourceData.records.push(Object.assign({ __rowNumber: sourceData.records.length + 2 }, record));
    },
    recalculateViews: function() { counters.recalculations++; return { ok: true }; }
  };
  return { sources: sources, counters: counters, deps: deps, session: session };
}

function corePerfilTestRunAll_() {
  var results = [];
  function test(name, fn) {
    fn();
    results.push({ name: name, ok: true });
  }

  test('1_membro_atualiza_proprio_telefone', function() {
    var fx = corePerfilTestFixture_();
    var response = core_atualizarMeuPerfilParaPortal_({ telefone: '(65) 99999-1111', chaveIdempotencia: 'phone-0001' }, {}, { deps: fx.deps });
    corePerfilTestAssert_(response.ok && fx.sources.PESSOAS_BASE.records[0].TELEFONE_PRINCIPAL === '+5565999991111', 'telefone nao atualizado');
  });

  test('2_membro_atualiza_resumo_academico', function() {
    var fx = corePerfilTestFixture_();
    var response = core_atualizarMeuPerfilParaPortal_({ resumoAcademico: 'Novo resumo academico', chaveIdempotencia: 'summary-001' }, {}, { deps: fx.deps });
    corePerfilTestAssert_(response.ok && fx.sources.MEMBROS_DETALHES.records[0].HISTORICO_ATIVIDADES_ACADEMICAS === 'Novo resumo academico', 'resumo nao atualizado');
  });

  test('3_membro_adiciona_e_remove_lattes', function() {
    var fx = corePerfilTestFixture_();
    var add = core_atualizarMeuPerfilParaPortal_({ links: [{ tipo: 'LATTES', url: 'https://lattes.cnpq.br/123' }], chaveIdempotencia: 'lattes-add-1' }, {}, { deps: fx.deps });
    var remove = core_atualizarMeuPerfilParaPortal_({ links: [{ tipo: 'LATTES', url: '' }], chaveIdempotencia: 'lattes-remove-1' }, {}, { deps: fx.deps });
    corePerfilTestAssert_(add.ok && remove.ok && fx.sources.PESSOAS_V2_LINKS_PERFIS.records[0].ATIVO === 'NAO', 'ciclo do Lattes falhou');
  });

  test('4_membro_nao_escolhe_outra_pessoa', function() {
    var fx = corePerfilTestFixture_();
    var response = core_atualizarMeuPerfilParaPortal_({ idPessoa: 'PES-2', telefone: '(65) 99999-2222', chaveIdempotencia: 'target-0001' }, {}, { deps: fx.deps });
    corePerfilTestAssert_(!response.ok && response.errorCode === 'IDENTIDADE_ALVO_NAO_PERMITIDA', 'ID alvo deveria ser rejeitado');
  });

  test('5_cpf_gera_solicitacao_sem_escrita_direta', function() {
    var fx = corePerfilTestFixture_();
    var before = fx.sources.PESSOAS_BASE.records[0].CPF;
    var response = core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '11144477735', justificativa: 'Solicito a correcao conforme documento oficial.', chaveIdempotencia: 'cpf-request-1' }, {}, { deps: fx.deps });
    corePerfilTestAssert_(response.ok && fx.sources.PESSOAS_BASE.records[0].CPF === before && fx.sources[CORE_PERFIL_SOLICITACOES_SHEET].records.length === 1, 'CPF foi escrito diretamente');
  });

  test('6_solicitacao_duplicada_rejeitada', function() {
    var fx = corePerfilTestFixture_();
    core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '11144477735', justificativa: 'Solicito a correcao conforme documento oficial.', chaveIdempotencia: 'cpf-request-1' }, {}, { deps: fx.deps });
    var duplicate = core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '39053344705', justificativa: 'Outra correcao para o mesmo campo cadastral.', chaveIdempotencia: 'cpf-request-2' }, {}, { deps: fx.deps });
    corePerfilTestAssert_(!duplicate.ok && duplicate.errorCode === 'SOLICITACAO_DUPLICADA', 'duplicidade nao rejeitada');
  });

  test('7_usuario_sem_permissao_nao_analisa', function() {
    var fx = corePerfilTestFixture_();
    var response = core_analisarSolicitacaoCadastralPortal_({ idSolicitacao: 'SAC-1', acao: 'EM_ANALISE' }, {}, { deps: fx.deps });
    corePerfilTestAssert_(!response.ok && response.errorCode === 'PERMISSAO_NEGADA', 'permissao administrativa nao exigida');
  });

  test('8_secretaria_diretoria_autorizada_analisa', function() {
    var fx = corePerfilTestFixture_(['portal:acessar']);
    fx.session.perfisPortal = ['SECRETARIA'];
    core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '11144477735', justificativa: 'Solicito a correcao conforme documento oficial.', chaveIdempotencia: 'cpf-request-1' }, {}, { deps: fx.deps });
    var id = fx.sources[CORE_PERFIL_SOLICITACOES_SHEET].records[0].ID_SOLICITACAO;
    var response = core_analisarSolicitacaoCadastralPortal_({ idSolicitacao: id, acao: 'EM_ANALISE' }, {}, { deps: fx.deps });
    corePerfilTestAssert_(response.ok && response.data.status === 'EM_ANALISE', 'analise autorizada falhou');
  });

  test('9_aprovacao_nao_aplica_automaticamente', function() {
    var fx = corePerfilTestFixture_(['portal:acessar', CORE_PERFIL_ADMIN_PERMISSION]);
    var before = fx.sources.PESSOAS_BASE.records[0].CPF;
    core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '11144477735', justificativa: 'Solicito a correcao conforme documento oficial.', chaveIdempotencia: 'cpf-request-1' }, {}, { deps: fx.deps });
    var id = fx.sources[CORE_PERFIL_SOLICITACOES_SHEET].records[0].ID_SOLICITACAO;
    core_analisarSolicitacaoCadastralPortal_({ idSolicitacao: id, acao: 'APROVADA' }, {}, { deps: fx.deps });
    corePerfilTestAssert_(fx.sources.PESSOAS_BASE.records[0].CPF === before, 'aprovacao aplicou automaticamente');
  });

  test('10_aplicacao_atualiza_fonte_oficial', function() {
    var fx = corePerfilTestFixture_(['portal:acessar', CORE_PERFIL_ADMIN_PERMISSION]);
    core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '11144477735', justificativa: 'Solicito a correcao conforme documento oficial.', chaveIdempotencia: 'cpf-request-1' }, {}, { deps: fx.deps });
    var id = fx.sources[CORE_PERFIL_SOLICITACOES_SHEET].records[0].ID_SOLICITACAO;
    core_analisarSolicitacaoCadastralPortal_({ idSolicitacao: id, acao: 'APROVADA' }, {}, { deps: fx.deps });
    var applied = core_aplicarSolicitacaoCadastralAprovadaPortal_({ idSolicitacao: id }, {}, { deps: fx.deps });
    corePerfilTestAssert_(applied.ok && fx.sources.PESSOAS_BASE.records[0].CPF === '11144477735' && fx.counters.recalculations === 1, 'aplicacao nao atualizou fonte/view');
  });

  test('11_aplicacao_duplicada_idempotente', function() {
    var fx = corePerfilTestFixture_(['portal:acessar', CORE_PERFIL_ADMIN_PERMISSION]);
    core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '11144477735', justificativa: 'Solicito a correcao conforme documento oficial.', chaveIdempotencia: 'cpf-request-1' }, {}, { deps: fx.deps });
    var id = fx.sources[CORE_PERFIL_SOLICITACOES_SHEET].records[0].ID_SOLICITACAO;
    core_analisarSolicitacaoCadastralPortal_({ idSolicitacao: id, acao: 'APROVADA' }, {}, { deps: fx.deps });
    core_aplicarSolicitacaoCadastralAprovadaPortal_({ idSolicitacao: id }, {}, { deps: fx.deps });
    var second = core_aplicarSolicitacaoCadastralAprovadaPortal_({ idSolicitacao: id }, {}, { deps: fx.deps });
    corePerfilTestAssert_(second.ok && second.data.idempotente === true && fx.counters.recalculations === 1, 'segunda aplicacao nao foi idempotente');
  });

  test('12_logs_nao_contem_cpf_completo', function() {
    var safe = JSON.stringify(corePerfilSafeLogPayload_('TEST', { ok: false, field: 'CPF', code: 'CPF_INVALIDO', cpf: '52998224725', value: '52998224725' }));
    corePerfilTestAssert_(safe.indexOf('52998224725') < 0, 'log seguro contem CPF completo');
  });

  test('13_setup_recusa_prod', function() {
    var response = core_setupSolicitacoesAtualizacaoCadastralDev_({ dryRun: false, environment: 'PROD', confirmacao: CORE_PERFIL_SETUP_CONFIRMATION });
    corePerfilTestAssert_(!response.ok && response.productionRefused === true && response.errorCode === 'SETUP_RECUSADO_FORA_DEV', 'setup PROD nao recusado');
  });

  test('14_dry_run_nao_escreve', function() {
    var fx = corePerfilTestFixture_();
    var response = core_atualizarMeuPerfilParaPortal_({ telefone: '(65) 99999-3333', chaveIdempotencia: 'dry-run-0001', dryRun: true }, {}, { deps: fx.deps });
    corePerfilTestAssert_(response.ok && response.data.dryRun === true && fx.counters.writes === 0 && fx.counters.appends === 0, 'dry-run escreveu');
  });

  test('15_fila_admin_identifica_pessoa_sem_expor_dados_completos', function() {
    var fx = corePerfilTestFixture_(['portal:acessar', CORE_PERFIL_ADMIN_PERMISSION]);
    core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '11144477735', justificativa: 'Solicito a correcao conforme documento oficial.', chaveIdempotencia: 'cpf-request-1' }, {}, { deps: fx.deps });
    var response = core_listarSolicitacoesCadastraisAdministracaoPortal_({}, {}, { deps: fx.deps });
    var item = response.ok && response.data.solicitacoes[0];
    var serialized = JSON.stringify(item || {});
    corePerfilTestAssert_(item && item.pessoa.nomeExibicao === 'Pessoa Teste', 'nome de exibicao ausente na fila');
    corePerfilTestAssert_(item.pessoa.rgaMascarado === '***234', 'RGA nao foi mascarado');
    corePerfilTestAssert_(item.pessoa.emailMascarado === 'pe***@example.org', 'email nao foi mascarado');
    corePerfilTestAssert_(serialized.indexOf('20201234') < 0 && serialized.indexOf('pessoa@example.org') < 0, 'fila expos identificador completo');
    corePerfilTestAssert_(!Object.prototype.hasOwnProperty.call(item, 'idPessoa'), 'fila expos ID_PESSOA desnecessario');
  });

  test('16_filtro_admin_por_pessoa_e_aplicado_no_core', function() {
    var fx = corePerfilTestFixture_(['portal:acessar', CORE_PERFIL_ADMIN_PERMISSION]);
    core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '11144477735', justificativa: 'Solicito a correcao conforme documento oficial.', chaveIdempotencia: 'cpf-request-1' }, {}, { deps: fx.deps });
    var byName = core_listarSolicitacoesCadastraisAdministracaoPortal_({ pessoa: 'Pessoa Teste' }, {}, { deps: fx.deps });
    var byRga = core_listarSolicitacoesCadastraisAdministracaoPortal_({ pessoa: '20201234' }, {}, { deps: fx.deps });
    var noMatch = core_listarSolicitacoesCadastraisAdministracaoPortal_({ pessoa: 'Outra Pessoa' }, {}, { deps: fx.deps });
    corePerfilTestAssert_(byName.data.paginacao.totalItens === 1 && byRga.data.paginacao.totalItens === 1, 'filtro por pessoa nao encontrou a solicitacao');
    corePerfilTestAssert_(noMatch.data.paginacao.totalItens === 0, 'filtro por pessoa retornou solicitacao incorreta');
  });

  return Object.freeze({ ok: true, total: results.length, resultados: Object.freeze(results) });
}

function geapaCoreRunTestesAtualizacaoCadastral() {
  return corePerfilTestRunAll_();
}
