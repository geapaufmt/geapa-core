/** Testes unitarios dos contratos de atualizacao cadastral. Nao escrevem em planilhas reais. */

function corePerfilTestAssert_(condition, message) {
  if (!condition) throw new Error('TESTE_FALHOU: ' + message);
}

function corePerfilTestFixture_(permissions, environment) {
  var now = new Date(2026, 6, 12, 10, 0, 0);
  var sources = {};
  function source(name, headers, records) {
    sources[name] = {
      name: name,
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
  source('PESSOAS_IDENTIFICADORES', [
    'ID_IDENTIFICADOR','ID_PESSOA','TIPO_IDENTIFICADOR','VALOR_IDENTIFICADOR','PRINCIPAL','ATIVO','OBS'
  ], [{
    ID_IDENTIFICADOR: 'IDN-EMAIL-1', ID_PESSOA: 'PES-1', TIPO_IDENTIFICADOR: 'EMAIL',
    VALOR_IDENTIFICADOR: 'pessoa@example.org', PRINCIPAL: 'SIM', ATIVO: 'SIM', OBS: ''
  }, {
    ID_IDENTIFICADOR: 'IDN-RGA-1', ID_PESSOA: 'PES-1', TIPO_IDENTIFICADOR: 'RGA',
    VALOR_IDENTIFICADOR: '20201234', PRINCIPAL: 'SIM', ATIVO: 'SIM', OBS: ''
  }]);
  source('PESSOAS_RESUMO_OPERACIONAL', ['ID_PESSOA','EMAIL','RGA','ATUALIZADO_EM'], [{
    ID_PESSOA: 'PES-1', EMAIL: 'pessoa@example.org', RGA: '20201234'
  }]);
  source('PESSOAS_LINKS_PERFIS', [
    'ID_LINK','ID_PESSOA','TIPO_LINK','URL','ROTULO','PUBLICAVEL','VISIVEL_PORTAL','FONTE',
    'VALIDADO_EM','ATIVO','CRIADO_EM','ATUALIZADO_EM','OBS'
  ], []);
  source(CORE_PERFIL_SOLICITACOES_SHEET, CORE_DOMAINS_V2_SCHEMAS.PESSOAS.filter(function(item) {
    return item.sheetName === CORE_PERFIL_SOLICITACOES_SHEET;
  })[0].headers, []);

  var counters = { writes: 0, appends: 0, deletes: 0, recalculations: 0, cacheInvalidations: [], openedSourceSets: [] };
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
    environment: environment || 'DEV',
    session: session,
    hasDevPermission: function(currentSession, permission) {
      return (currentSession.permissoes || []).indexOf(permission) >= 0;
    },
    now: function() { return new Date(now.getTime()); },
    uuid: function() { uuidCounter++; return '00000000-0000-4000-8000-' + ('000000000000' + uuidCounter).slice(-12); },
    hash: function(value) { return 'HASH_' + String(value).replace(/[^A-Za-z0-9]/g, '').slice(0, 30); },
    withLock: function(key, fn) { return fn(); },
    openPessoas: function(sourceNames) {
      counters.openedSourceSets.push((sourceNames || []).slice());
      return sources;
    },
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
    deleteRecord: function(sourceData, record) {
      counters.deletes++;
      sourceData.records = sourceData.records.filter(function(item) {
        return Number(item.__rowNumber) !== Number(record.__rowNumber);
      });
    },
    recalculateViews: function() { counters.recalculations++; return { ok: true }; },
    invalidateCaches: function(values, env) {
      counters.cacheInvalidations.push({ values: values.slice(), environment: env });
      return { identifiers: values.length };
    }
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
    corePerfilTestAssert_(add.ok && remove.ok && fx.sources.PESSOAS_LINKS_PERFIS.records[0].ATIVO === 'NAO', 'ciclo do Lattes falhou');
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

  test('9_aprovacao_e_aplicacao_sao_uma_operacao', function() {
    var fx = corePerfilTestFixture_(['portal:acessar', CORE_PERFIL_ADMIN_PERMISSION]);
    core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '11144477735', justificativa: 'Solicito a correcao conforme documento oficial.', chaveIdempotencia: 'cpf-request-1' }, {}, { deps: fx.deps });
    var id = fx.sources[CORE_PERFIL_SOLICITACOES_SHEET].records[0].ID_SOLICITACAO;
    var applied = core_aprovarEAplicarSolicitacaoCadastralPortal_({ idSolicitacao: id, confirmacao: true }, {}, { deps: fx.deps });
    corePerfilTestAssert_(applied.ok && applied.data.status === 'APLICADA' && fx.sources.PESSOAS_BASE.records[0].CPF === '11144477735', 'aprovacao/aplicacao coordenada falhou');
  });

  test('10_aplicacao_atualiza_fonte_oficial', function() {
    var fx = corePerfilTestFixture_(['portal:acessar', CORE_PERFIL_ADMIN_PERMISSION]);
    core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '11144477735', justificativa: 'Solicito a correcao conforme documento oficial.', chaveIdempotencia: 'cpf-request-1' }, {}, { deps: fx.deps });
    var id = fx.sources[CORE_PERFIL_SOLICITACOES_SHEET].records[0].ID_SOLICITACAO;
    var applied = core_aprovarEAplicarSolicitacaoCadastralPortal_({ idSolicitacao: id, confirmacao: true }, {}, { deps: fx.deps });
    corePerfilTestAssert_(applied.ok && fx.sources.PESSOAS_BASE.records[0].CPF === '11144477735' && fx.counters.recalculations === 1, 'aplicacao nao atualizou fonte/view');
  });

  test('11_aplicacao_duplicada_idempotente', function() {
    var fx = corePerfilTestFixture_(['portal:acessar', CORE_PERFIL_ADMIN_PERMISSION]);
    core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '11144477735', justificativa: 'Solicito a correcao conforme documento oficial.', chaveIdempotencia: 'cpf-request-1' }, {}, { deps: fx.deps });
    var id = fx.sources[CORE_PERFIL_SOLICITACOES_SHEET].records[0].ID_SOLICITACAO;
    core_aprovarEAplicarSolicitacaoCadastralPortal_({ idSolicitacao: id, confirmacao: true }, {}, { deps: fx.deps });
    var second = core_aprovarEAplicarSolicitacaoCadastralPortal_({ idSolicitacao: id, confirmacao: true }, {}, { deps: fx.deps });
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

  test('17_prod_escreve_somente_com_contexto_prod', function() {
    var fx = corePerfilTestFixture_(['portal:acessar'], 'PROD');
    var response = core_atualizarMeuPerfilParaPortal_({ links: [{ tipo: 'LATTES', url: 'https://lattes.cnpq.br/123' }], chaveIdempotencia: 'prod-link-001' }, {}, { deps: fx.deps });
    var link = fx.sources.PESSOAS_LINKS_PERFIS.records[0];
    corePerfilTestAssert_(response.ok && link.FONTE === 'PORTAL_PROD', 'escrita PROD nao preservou a origem oficial');
  });

  test('18_prod_nao_concede_gestao_apenas_por_perfil', function() {
    var fx = corePerfilTestFixture_(['portal:acessar'], 'PROD');
    fx.session.perfisPortal = ['SECRETARIA'];
    var response = core_listarSolicitacoesCadastraisAdministracaoPortal_({}, {}, { deps: fx.deps });
    corePerfilTestAssert_(!response.ok && response.errorCode === 'PERMISSAO_NEGADA', 'PROD concedeu gestao sem permissao canonica');
  });

  test('19_setup_prod_exige_confirmacao_dedicada', function() {
    var response = core_setupSolicitacoesAtualizacaoCadastralProd_({ dryRun: false, environment: 'PROD' });
    corePerfilTestAssert_(!response.ok && response.errorCode === 'CONFIRMACAO_OBRIGATORIA' && response.message.indexOf(CORE_PERFIL_SETUP_CONFIRMATION_PROD) >= 0, 'setup PROD nao exigiu confirmacao dedicada');
  });

  test('20_setup_prod_recusa_ambiente_dev', function() {
    var response = core_setupSolicitacoesAtualizacaoCadastralProd_({ dryRun: true, environment: 'DEV' });
    corePerfilTestAssert_(!response.ok && response.errorCode === 'SETUP_RECUSADO_FORA_PROD', 'entrada PROD aceitou ambiente DEV');
  });

  test('21_registry_cadastral_nao_aceita_fallback_all', function() {
    var original = core_getRegistryRaw_;
    core_getRegistryRaw_ = function() {
      return { PESSOAS_V2_BASE: { ALL: { ativo: true, id: 'ID-ALL', sheet: 'PESSOAS_BASE' } } };
    };
    var refused = false;
    try {
      corePerfilRegistryMeta_('PESSOAS_V2_BASE', 'PROD');
    } catch (error) {
      refused = error && error.message === 'REGISTRY_PROD_INDISPONIVEL_PESSOAS_V2_BASE';
    } finally {
      core_getRegistryRaw_ = original;
    }
    corePerfilTestAssert_(refused, 'Registry cadastral aceitou entrada ALL');
  });

  test('22_contextos_portal_mapeiam_ambientes_isolados', function() {
    corePerfilTestAssert_(corePerfilResolveEnvironment_({ ambientePortal: 'HOMOLOG' }, {}) === 'DEV', 'HOMOLOG nao mapeou para DEV');
    corePerfilTestAssert_(corePerfilResolveEnvironment_({ ambientePortal: 'PROD' }, {}) === 'PROD', 'PROD nao permaneceu isolado');
  });

  test('23_data_normaliza_sem_timezone_e_mascara_ano', function() {
    var date = new Date(2004, 11, 13, 0, 0, 0);
    corePerfilTestAssert_(corePerfilCanonicalDate_(date) === '2004-12-13', 'Date nao normalizou para formato canonico');
    corePerfilTestAssert_(corePerfilFormatDateForPortal_(date) === '13/12/2004', 'data administrativa nao ficou em DD/MM/AAAA');
    corePerfilTestAssert_(corePerfilMaskValue_('DATA_NASCIMENTO', date) === '**/**/2004', 'mascara de data incorreta');
    var futureRejected = false;
    try { corePerfilCanonicalDate_('2999-01-01'); } catch (error) { futureRejected = error && error.message === 'DATA_NASCIMENTO_INVALIDA'; }
    corePerfilTestAssert_(futureRejected, 'data futura nao foi rejeitada');
  });

  test('24_timeout_reconcilia_por_chave_sem_duplicar', function() {
    var fx = corePerfilTestFixture_();
    var payload = { campo: 'RGA', valorSolicitado: '20209999', justificativa: 'Solicito correcao conforme documento academico.', chaveIdempotencia: 'rga-reconcile-1' };
    var created = core_solicitarCorrecaoMeuPerfilParaPortal_(payload, {}, { deps: fx.deps });
    var confirmed = core_consultarMinhaSolicitacaoCadastralPortal_({ chaveIdempotencia: payload.chaveIdempotencia }, {}, { deps: fx.deps });
    var replay = core_solicitarCorrecaoMeuPerfilParaPortal_(payload, {}, { deps: fx.deps });
    corePerfilTestAssert_(created.ok && confirmed.ok && confirmed.data.encontrada === true, 'solicitacao persistida nao foi reconciliada');
    corePerfilTestAssert_(confirmed.data.solicitacao.status === 'PENDENTE' && replay.data.idempotente === true, 'reconciliacao nao preservou status/idempotencia');
    corePerfilTestAssert_(fx.sources[CORE_PERFIL_SOLICITACOES_SHEET].records.length === 1, 'reconciliacao duplicou a linha');
  });

  test('25_listagem_admin_permanece_mascarada', function() {
    var fx = corePerfilTestFixture_(['portal:acessar', CORE_PERFIL_ADMIN_PERMISSION]);
    var created = core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'RGA', valorSolicitado: '20209999', justificativa: 'Solicito correcao conforme documento academico.', chaveIdempotencia: 'rga-list-mask-1' }, {}, { deps: fx.deps });
    var listed = core_listarSolicitacoesCadastraisAdministracaoPortal_({}, {}, { deps: fx.deps });
    var serialized = JSON.stringify(listed.data.solicitacoes[0]);
    corePerfilTestAssert_(created.ok && serialized.indexOf('20209999') < 0 && serialized.indexOf('20201234') < 0, 'listagem administrativa revelou valores completos');
  });

  test('26_detalhe_admin_compara_rga_autorizado', function() {
    var fx = corePerfilTestFixture_(['portal:acessar', CORE_PERFIL_ADMIN_PERMISSION]);
    var created = core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'RGA', valorSolicitado: '20209999', justificativa: 'Solicito correcao conforme documento academico.', chaveIdempotencia: 'rga-detail-001' }, {}, { deps: fx.deps });
    var detail = core_detalharSolicitacaoCadastralAdministracaoPortal_({ idSolicitacao: created.data.idSolicitacao }, {}, { deps: fx.deps });
    corePerfilTestAssert_(detail.ok && detail.data.valorAtual === '20201234' && detail.data.valorSolicitado === '20209999', 'detalhe autorizado nao permitiu comparar RGA');
    corePerfilTestAssert_(detail.data.pessoa.nomeExibicao === 'Pessoa Teste' && detail.data.pessoa.rgaMascarado === '***234', 'cabecalho do solicitante incorreto');
  });

  test('27_cpf_exige_revelacao_explicita_e_log_seguro', function() {
    var fx = corePerfilTestFixture_(['portal:acessar', CORE_PERFIL_ADMIN_PERMISSION]);
    var created = core_solicitarCorrecaoMeuPerfilParaPortal_({ campo: 'CPF', valorSolicitado: '11144477735', justificativa: 'Solicito correcao conforme documento oficial.', chaveIdempotencia: 'cpf-detail-001' }, {}, { deps: fx.deps });
    var masked = core_detalharSolicitacaoCadastralAdministracaoPortal_({ idSolicitacao: created.data.idSolicitacao }, {}, { deps: fx.deps });
    var revealed = core_detalharSolicitacaoCadastralAdministracaoPortal_({ idSolicitacao: created.data.idSolicitacao, revelarDados: true }, {}, { deps: fx.deps });
    var safeLog = JSON.stringify(corePerfilSafeLogPayload_('ADMIN_SENSITIVE_DETAIL_VIEW', { ok: true, field: 'CPF', requestId: created.data.idSolicitacao, actorHash: 'HASH_ADMIN', revealed: true, value: '11144477735' }));
    corePerfilTestAssert_(masked.ok && masked.data.dadosRevelados === false && masked.data.valorSolicitado.indexOf('11144477735') < 0, 'CPF foi revelado sem acao explicita');
    corePerfilTestAssert_(revealed.ok && revealed.data.dadosRevelados === true && revealed.data.valorSolicitado === '111.444.777-35', 'revelacao autorizada do CPF falhou');
    corePerfilTestAssert_(safeLog.indexOf('11144477735') < 0 && safeLog.indexOf('HASH_ADMIN') >= 0, 'log de revelacao expos valor sensivel ou perdeu ator');
  });

  test('28_solicitacao_le_apenas_fonte_do_campo_e_fila', function() {
    var fx = corePerfilTestFixture_();
    var response = core_solicitarCorrecaoMeuPerfilParaPortal_({
      campo: 'DATA_NASCIMENTO',
      valorSolicitado: '11/01/1990',
      justificativa: 'Solicito correcao conforme documento oficial apresentado.',
      chaveIdempotencia: 'date-selective-001'
    }, {}, { deps: fx.deps });
    var opened = fx.counters.openedSourceSets[0] || [];
    corePerfilTestAssert_(response.ok, 'solicitacao seletiva falhou');
    corePerfilTestAssert_(opened.length === 2, 'solicitacao abriu fontes desnecessarias');
    corePerfilTestAssert_(opened.indexOf('PESSOAS_BASE') >= 0 && opened.indexOf(CORE_PERFIL_SOLICITACOES_SHEET) >= 0, 'fontes obrigatorias nao foram abertas');
  });

  test('29_listagem_propria_le_somente_fila', function() {
    var fx = corePerfilTestFixture_();
    var response = core_listarMinhasSolicitacoesCadastraisPortal_({}, { deps: fx.deps });
    var opened = fx.counters.openedSourceSets[0] || [];
    corePerfilTestAssert_(response.ok && opened.length === 1 && opened[0] === CORE_PERFIL_SOLICITACOES_SHEET, 'listagem propria abriu fontes desnecessarias');
  });

  test('30_atualizacao_telefone_le_somente_base_e_fila', function() {
    var fx = corePerfilTestFixture_();
    var response = core_atualizarMeuPerfilParaPortal_({
      telefone: '(65) 99999-4444',
      chaveIdempotencia: 'phone-selective-001'
    }, {}, { deps: fx.deps });
    var opened = fx.counters.openedSourceSets[0] || [];
    corePerfilTestAssert_(response.ok && opened.length === 2, 'telefone abriu fontes desnecessarias');
    corePerfilTestAssert_(opened.indexOf('PESSOAS_BASE') >= 0 && opened.indexOf(CORE_PERFIL_SOLICITACOES_SHEET) >= 0, 'telefone nao abriu base e fila');
  });

  test('31_reparacao_prod_interna_nao_aceita_mutacao', function() {
    var fx = corePerfilTestFixture_(['portal:acessar', CORE_PERFIL_ADMIN_PERMISSION], 'PROD');
    var refused = false;
    try {
      core_diagnosticarReparacaoSolicitacaoCadastralProd_({
        idSolicitacao: 'SAC-LEGADA-1',
        dryRun: false,
        confirmacao: 'REPARAR_SOLICITACAO_CADASTRAL_PROD_SAC-LEGADA-1',
        deps: fx.deps
      });
    } catch (error) {
      refused = error && error.message === 'REPARACAO_PROD_MUTACAO_NAO_EXPOSTA';
    }
    corePerfilTestAssert_(refused, 'diagnostico interno permitiu tentativa de mutacao PROD');
    corePerfilTestAssert_(fx.counters.writes === 0 && fx.counters.appends === 0, 'tentativa recusada escreveu em fonte de dados');
  });

  return Object.freeze({ ok: true, total: results.length, resultados: Object.freeze(results) });
}

function geapaCoreRunTestesAtualizacaoCadastral() {
  return corePerfilTestRunAll_();
}
