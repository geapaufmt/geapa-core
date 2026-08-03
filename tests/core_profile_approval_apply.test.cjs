const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const context = {
  console,
  Logger: { log() {} },
  CORE_DOMAINS_V2_SCHEMAS: {
    PESSOAS: [
      'SOLICITACOES_ATUALIZACAO_CADASTRAL',
      'PESSOAS_BASE',
      'PESSOAS_IDENTIFICADORES',
      'MEMBROS_DETALHES',
      'PESSOAS_RESUMO_OPERACIONAL',
      'PESSOAS_LINKS_PERFIS'
    ].map((sheetName) => ({ sheetName }))
  },
  core_normalizeText_(value, options) {
    let text = String(value == null ? '' : value).trim().replace(/\s+/g, ' ');
    if (options && options.removeAccents) text = text.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (options && options.caseMode === 'upper') text = text.toUpperCase();
    return text;
  },
  core_normalizeHeader_(value) {
    return String(value || '').trim().toUpperCase().replace(/[^A-Z0-9]+/g, '_');
  },
  corePortalNormalizeEmail_(value) {
    const email = String(value || '').trim().toLowerCase();
    return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) ? email : '';
  },
  corePortalNormalizePermission_(value) {
    return String(value || '').trim().toLowerCase();
  },
  core_domainsV2AuditIsSim_(value) {
    return ['SIM', 'S', 'TRUE', '1'].includes(String(value || '').trim().toUpperCase());
  },
  core_buildRowFromObjectByHeaders_(headers, record) {
    return headers.map((header) => record[header] == null ? '' : record[header]);
  }
};
vm.createContext(context);
vm.runInContext(fs.readFileSync(path.join(root, '37_core_profile_updates.js'), 'utf8'), context);
vm.runInContext(fs.readFileSync(path.join(root, '37a_core_profile_approval_apply.js'), 'utf8'), context);

const REQUEST_HEADERS = [
  'ID_SOLICITACAO', 'ID_PESSOA', 'TIPO_SOLICITACAO', 'CAMPO',
  'VALOR_ATUAL_MASCARADO', 'VALOR_ATUAL_HASH', 'VALOR_SOLICITADO',
  'JUSTIFICATIVA', 'STATUS', 'SOLICITADO_EM', 'SOLICITADO_POR',
  'ANALISADO_EM', 'ANALISADO_POR', 'DECISAO', 'MOTIVO_DECISAO',
  'APLICADO_EM', 'ID_LOG', 'ATIVO', 'CRIADO_EM', 'ATUALIZADO_EM'
];

function fixture(options = {}) {
  const state = {
    now: new Date('2026-07-29T12:00:00.000Z'),
    writes: [],
    appends: [],
    deletes: [],
    recalculations: 0,
    cacheInvalidations: [],
    failWrite: options.failWrite || null,
    failWriteUsed: false,
    failRecalculate: options.failRecalculate === true,
    opened: []
  };
  function source(name, headers, records) {
    return {
      name,
      sheet: { name },
      headers: headers.slice(),
      records: records.map((record, index) => ({ __rowNumber: index + 2, ...record }))
    };
  }
  const idPessoa = options.idPessoa || 'PES-000036';
  const oldEmail = options.oldEmail || 'viniciuspocket1515@gmail.com';
  const newEmail = options.newEmail || 'viniciusmalamim@gmail.com';
  const requestId = options.requestId || 'SAC-F1BD0911-69B4-42EE-99CE-7F5FD0081CB2';
  const field = options.field || 'EMAIL_PRINCIPAL';
  const currentValue = field === 'RGA' ? '20201234' : oldEmail;
  const requestedValue = field === 'RGA' ? '20209999' : newEmail;
  const hash = (value) => 'HASH_' + String(value).replace(/[^A-Za-z0-9]/g, '').slice(0, 30);
  const sources = {
    PESSOAS_BASE: source('PESSOAS_BASE',
      ['ID_PESSOA', 'NOME_COMPLETO', 'NOME_CIVIL', 'EMAIL_PRINCIPAL', 'CPF', 'DATA_NASCIMENTO', 'ATUALIZADO_EM', 'ATIVO'],
      [{ ID_PESSOA: idPessoa, NOME_COMPLETO: 'Pessoa Teste', EMAIL_PRINCIPAL: oldEmail, CPF: '52998224725', DATA_NASCIMENTO: '1990-01-10', ATIVO: 'SIM' }]
    ),
    MEMBROS_DETALHES: source('MEMBROS_DETALHES',
      ['ID_PESSOA', 'RGA', 'CURSO_ID', 'ATUALIZADO_EM'],
      [{ ID_PESSOA: idPessoa, RGA: '20201234', CURSO_ID: 'AGRONOMIA_UFMT_SINOP' }]
    ),
    PESSOAS_IDENTIFICADORES: source('PESSOAS_IDENTIFICADORES',
      ['ID_IDENTIFICADOR', 'ID_PESSOA', 'TIPO_IDENTIFICADOR', 'VALOR_IDENTIFICADOR', 'PRINCIPAL', 'ATIVO', 'OBS'],
      field === 'RGA'
        ? [{ ID_IDENTIFICADOR: 'IDN-RGA-OLD', ID_PESSOA: idPessoa, TIPO_IDENTIFICADOR: 'RGA', VALOR_IDENTIFICADOR: '20201234', PRINCIPAL: 'SIM', ATIVO: 'SIM', OBS: '' }]
        : [{ ID_IDENTIFICADOR: 'IDN-EMAIL-OLD', ID_PESSOA: idPessoa, TIPO_IDENTIFICADOR: 'EMAIL', VALOR_IDENTIFICADOR: oldEmail, PRINCIPAL: 'SIM', ATIVO: 'SIM', OBS: '' }]
    ),
    PESSOAS_RESUMO_OPERACIONAL: source('PESSOAS_RESUMO_OPERACIONAL',
      ['ID_PESSOA', 'EMAIL', 'RGA', 'ATUALIZADO_EM'],
      [{ ID_PESSOA: idPessoa, EMAIL: oldEmail, RGA: '20201234' }]
    ),
    SOLICITACOES_ATUALIZACAO_CADASTRAL: source('SOLICITACOES_ATUALIZACAO_CADASTRAL',
      REQUEST_HEADERS,
      [{
        ID_SOLICITACAO: requestId,
        ID_PESSOA: idPessoa,
        TIPO_SOLICITACAO: 'CORRECAO_SENSIVEL',
        CAMPO: field,
        VALOR_ATUAL_HASH: hash(currentValue),
        VALOR_SOLICITADO: requestedValue,
        JUSTIFICATIVA: 'Solicitacao controlada para teste automatizado.',
        STATUS: options.status || 'PENDENTE',
        ANALISADO_EM: options.status === 'APROVADA' ? state.now : '',
        ANALISADO_POR: options.status === 'APROVADA' ? 'secretaria@example.org' : '',
        DECISAO: options.status === 'APROVADA' ? 'APROVADA' : '',
        ATIVO: 'SIM'
      }]
    )
  };
  const session = {
    ok: true,
    autenticado: true,
    portalAtivo: true,
    idPessoa: 'PES-ADMIN',
    email: 'diretoria@example.org',
    permissoes: options.unauthorized ? ['portal:acessar'] : ['portal:acessar', 'membros:analisar_correcoes'],
    perfisPortal: ['DIRETORIA']
  };
  const deps = {
    environment: options.environment || 'DEV',
    session,
    now: () => new Date(state.now),
    uuid: (() => { let n = 0; return () => `00000000-0000-4000-8000-${String(++n).padStart(12, '0')}`; })(),
    hash,
    withLock(_key, fn) { return fn(); },
    hasDevPermission(currentSession, permission) {
      return currentSession.permissoes.includes(permission);
    },
    openPessoas(names, openOptions) {
      state.opened.push({ names: names.slice(), forWrite: openOptions.forWrite === true, environment: deps.environment });
      return sources;
    },
    writeRecord(sourceData, record) {
      const failureMatches = state.failWrite &&
        state.failWrite.source === sourceData.name &&
        (!state.failWrite.whenStatus || state.failWrite.whenStatus === record.STATUS);
      if (failureMatches && !state.failWriteUsed) {
        state.failWriteUsed = true;
        throw new Error(state.failWrite.code || 'FAKE_WRITE_FAILURE');
      }
      const index = sourceData.records.findIndex((item) => Number(item.__rowNumber) === Number(record.__rowNumber));
      if (index < 0) throw new Error('FAKE_ROW_NAO_ENCONTRADA');
      sourceData.records[index] = { ...record };
      state.writes.push({ source: sourceData.name, row: record.__rowNumber, status: record.STATUS || '' });
    },
    appendRecord(sourceData, record) {
      const appended = { ...record, __rowNumber: record.__rowNumber || sourceData.records.length + 2 };
      sourceData.records.push(appended);
      state.appends.push({ source: sourceData.name, row: appended.__rowNumber });
    },
    deleteRecord(sourceData, record) {
      sourceData.records = sourceData.records.filter((item) => Number(item.__rowNumber) !== Number(record.__rowNumber));
      state.deletes.push({ source: sourceData.name, row: record.__rowNumber });
    },
    recalculateViews() {
      state.recalculations++;
      if (state.failRecalculate) {
        state.failRecalculate = false;
        throw new Error('ERRO_RECALCULO_VIEW');
      }
      return { ok: true };
    },
    invalidateCaches(values, environment) {
      state.cacheInvalidations.push({ values: values.slice(), environment });
      return { identifiers: values.length };
    }
  };
  function apply(extra = {}) {
    return context.core_aprovarEAplicarSolicitacaoCadastralPortal_(
      { idSolicitacao: requestId, confirmacao: true, ...extra },
      { ambientePortal: deps.environment, sessaoOficial: { email: session.email } },
      { deps }
    );
  }
  return { state, sources, deps, session, idPessoa, oldEmail, newEmail, requestId, apply };
}

function request(fx) {
  return fx.sources.SOLICITACOES_ATUALIZACAO_CADASTRAL.records[0];
}
function identifiers(fx, type = 'EMAIL') {
  return fx.sources.PESSOAS_IDENTIFICADORES.records.filter((row) => row.TIPO_IDENTIFICADOR === type);
}
function resolvesByEmail(fx, email) {
  const normalized = email.trim().toLowerCase();
  const base = fx.sources.PESSOAS_BASE.records.find((row) => String(row.EMAIL_PRINCIPAL || '').toLowerCase() === normalized);
  if (base) return base.ID_PESSOA;
  const found = identifiers(fx).find((row) =>
    String(row.VALOR_IDENTIFICADOR || '').toLowerCase() === normalized && row.ATIVO === 'SIM');
  return found ? found.ID_PESSOA : '';
}

test('aprova e aplica em uma chamada sem estado intermediario APROVADA', () => {
  const fx = fixture();
  const result = fx.apply();
  assert.equal(result.ok, true);
  assert.equal(result.data.status, 'APLICADA');
  assert.equal(request(fx).STATUS, 'APLICADA');
  assert.equal(request(fx).DECISAO, 'APROVADA');
});

test('solicitacao legada APROVADA usa o alias protegido', () => {
  const fx = fixture({ status: 'APROVADA' });
  const result = context.core_aplicarSolicitacaoCadastralAprovadaPortal_(
    { idSolicitacao: fx.requestId },
    { ambientePortal: 'DEV' },
    { deps: fx.deps }
  );
  assert.equal(result.ok, true);
  assert.equal(request(fx).STATUS, 'APLICADA');
});

test('alias legado recusa solicitacao nova ainda pendente', () => {
  const fx = fixture();
  const result = context.core_aplicarSolicitacaoCadastralAprovadaPortal_(
    { idSolicitacao: fx.requestId },
    { ambientePortal: 'DEV' },
    { deps: fx.deps }
  );
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'SOLICITACAO_LEGADA_NAO_APROVADA');
  assert.equal(request(fx).STATUS, 'PENDENTE');
});

test('novo contrato exige confirmacao explicita', () => {
  const fx = fixture();
  const result = context.core_aprovarEAplicarSolicitacaoCadastralPortal_(
    { idSolicitacao: fx.requestId },
    { ambientePortal: 'DEV' },
    { deps: fx.deps }
  );
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'CONFIRMACAO_APROVAR_APLICAR_OBRIGATORIA');
  assert.equal(fx.state.opened.length, 0);
});

test('retry de APLICADA e clique duplo sao idempotentes', () => {
  const fx = fixture();
  const first = fx.apply();
  const appends = fx.state.appends.length;
  const second = fx.apply();
  assert.equal(first.ok, true);
  assert.equal(second.data.idempotente, true);
  assert.equal(fx.state.appends.length, appends);
  assert.equal(fx.state.recalculations, 1);
});

test('EMAIL_PRINCIPAL sincroniza base, cria novo identificador e desativa o antigo', () => {
  const fx = fixture();
  fx.apply();
  assert.equal(fx.sources.PESSOAS_BASE.records[0].EMAIL_PRINCIPAL, fx.newEmail);
  const rows = identifiers(fx);
  assert.equal(rows.find((row) => row.VALOR_IDENTIFICADOR === fx.newEmail).ATIVO, 'SIM');
  assert.equal(rows.find((row) => row.VALOR_IDENTIFICADOR === fx.newEmail).PRINCIPAL, 'SIM');
  assert.equal(rows.find((row) => row.VALOR_IDENTIFICADOR === fx.oldEmail).ATIVO, 'NAO');
  assert.equal(rows.filter((row) => row.ATIVO === 'SIM' && row.PRINCIPAL === 'SIM').length, 1);
});

test('novo email existente para a mesma pessoa e reativado sem duplicar', () => {
  const fx = fixture();
  fx.sources.PESSOAS_IDENTIFICADORES.records.push({
    __rowNumber: 3, ID_IDENTIFICADOR: 'IDN-EMAIL-NEW', ID_PESSOA: fx.idPessoa,
    TIPO_IDENTIFICADOR: 'EMAIL', VALOR_IDENTIFICADOR: fx.newEmail,
    PRINCIPAL: 'NAO', ATIVO: 'NAO', OBS: ''
  });
  fx.apply();
  assert.equal(identifiers(fx).length, 2);
  assert.equal(fx.state.appends.length, 0);
});

test('email de outra pessoa e bloqueado de forma case-insensitive', () => {
  const fx = fixture();
  fx.sources.PESSOAS_BASE.records.push({
    __rowNumber: 3, ID_PESSOA: 'PES-OTHER', EMAIL_PRINCIPAL: fx.newEmail.toUpperCase(), ATIVO: 'SIM'
  });
  const result = fx.apply();
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'EMAIL_JA_VINCULADO_A_OUTRA_PESSOA');
  assert.equal(fx.sources.PESSOAS_BASE.records[0].EMAIL_PRINCIPAL, fx.oldEmail);
  assert.equal(request(fx).STATUS, 'PENDENTE');
});

test('falha em identificadores nao deixa PESSOAS_BASE incoerente', () => {
  const fx = fixture({ failWrite: { source: 'PESSOAS_IDENTIFICADORES', code: 'IDENTIFIER_WRITE_FAILED' } });
  const result = fx.apply();
  assert.equal(result.ok, false);
  assert.equal(fx.sources.PESSOAS_BASE.records[0].EMAIL_PRINCIPAL, fx.oldEmail);
  assert.equal(request(fx).STATUS, 'PENDENTE');
});

test('falha no recalculo compensa fontes e marca ERRO_APLICACAO', () => {
  const fx = fixture({ failRecalculate: true });
  const result = fx.apply();
  assert.equal(result.ok, false);
  assert.equal(fx.sources.PESSOAS_BASE.records[0].EMAIL_PRINCIPAL, fx.oldEmail);
  assert.equal(request(fx).STATUS, 'ERRO_APLICACAO');
  assert.equal(identifiers(fx).filter((row) => row.ATIVO === 'SIM')[0].VALOR_IDENTIFICADOR, fx.oldEmail);
});

test('falha ao gravar resultado final compensa e pode ser reprocessada', () => {
  const fx = fixture({ failWrite: { source: 'SOLICITACOES_ATUALIZACAO_CADASTRAL', whenStatus: 'APLICADA', code: 'REQUEST_FINAL_WRITE_FAILED' } });
  const failed = fx.apply();
  assert.equal(failed.ok, false);
  assert.equal(request(fx).STATUS, 'ERRO_APLICACAO');
  assert.equal(fx.sources.PESSOAS_BASE.records[0].EMAIL_PRINCIPAL, fx.oldEmail);
  const retry = fx.apply();
  assert.equal(retry.ok, true);
  assert.equal(request(fx).STATUS, 'APLICADA');
});

test('preenche ID_LOG, APLICADO_EM e ator oficial, ignorando ator do frontend', () => {
  const fx = fixture();
  fx.apply({ analisadoPor: 'atacante@example.org', idPessoa: 'PES-OTHER' });
  assert.match(String(request(fx).ID_LOG), /^LOG-/);
  assert.ok(request(fx).APLICADO_EM);
  assert.equal(request(fx).ANALISADO_POR, fx.session.email);
});

test('invalida caches do email antigo, novo e ID_PESSOA no ambiente correto', () => {
  const fx = fixture();
  fx.apply();
  assert.deepEqual(JSON.parse(JSON.stringify(fx.state.cacheInvalidations[0])), {
    values: [fx.idPessoa.toLowerCase(), fx.oldEmail, fx.newEmail],
    environment: 'DEV'
  });
  assert.equal(resolvesByEmail(fx, fx.oldEmail), '');
  assert.equal(resolvesByEmail(fx, fx.newEmail), fx.idPessoa);
});

test('usuario sem permissao e bloqueado antes de abrir fontes', () => {
  const fx = fixture({ unauthorized: true, environment: 'PROD' });
  const result = fx.apply();
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'PERMISSAO_NEGADA');
  assert.equal(fx.state.opened.length, 0);
});

test('PROD abre somente fontes PROD e teste nao executa escrita real', () => {
  const fx = fixture({ environment: 'PROD' });
  const result = fx.apply();
  assert.equal(result.ok, true);
  assert.ok(fx.state.opened.every((row) => row.environment === 'PROD'));
});

test('RGA sincroniza MEMBROS_DETALHES e identificadores oficiais', () => {
  const fx = fixture({ field: 'RGA' });
  const result = fx.apply();
  assert.equal(result.ok, true);
  assert.equal(fx.sources.MEMBROS_DETALHES.records[0].RGA, '20209999');
  assert.equal(identifiers(fx, 'RGA').filter((row) => row.ATIVO === 'SIM')[0].VALOR_IDENTIFICADOR, '20209999');
});

test('dry-run de reparacao cobre o caso PROD informado sem escrever', () => {
  const fx = fixture({ status: 'APROVADA', environment: 'PROD' });
  const before = JSON.stringify(fx.sources);
  const report = context.core_diagnosticarReparacaoSolicitacaoCadastralProd_({
    idSolicitacao: fx.requestId,
    dryRun: true,
    deps: fx.deps
  });
  assert.equal(report.ok, true);
  assert.equal(report.idSolicitacao, 'SAC-F1BD0911-69B4-42EE-99CE-7F5FD0081CB2');
  assert.equal(report.idPessoa, 'PES-000036');
  assert.equal(report.escritaExecutada, false);
  assert.equal(report.tokenConfirmacao, `REPARAR_SOLICITACAO_CADASTRAL_PROD_${fx.requestId}`);
  assert.equal(JSON.stringify(fx.sources), before);
});

test('resolvedor cadastral repassa explicitamente o ambiente DEV para a sessao oficial', () => {
  let receivedOptions = null;
  const session = context.corePerfilResolveSession_({
    ambientePortal: 'DEV',
    sessaoOficial: { email: 'pessoa.teste@example.org' },
    traceId: 'req-session-dev'
  }, {
    resolveSession(_email, options) {
      receivedOptions = options;
      return {
        ok: true,
        autenticado: true,
        portalAtivo: true,
        idPessoa: 'PES-TESTE',
        email: 'pessoa.teste@example.org'
      };
    }
  });

  assert.equal(session.ok, true);
  assert.equal(receivedOptions.ambiente, 'DEV');
  assert.equal(receivedOptions.environment, 'DEV');
  assert.equal(receivedOptions.traceId, 'req-session-dev');
});

test('falha de resolucao preserva causa e retorna somente diagnostico seguro', () => {
  const authorization = context.corePerfilAuthorizeOwn_({
    ambientePortal: 'DEV',
    sessaoOficial: { email: 'pessoa.teste@example.org' },
    traceId: 'req-session-failure'
  }, {
    resolveSession() {
      return {
        ok: false,
        autenticado: false,
        portalAtivo: false,
        motivoBloqueio: 'PESSOAS_V2_DB_INDISPONIVEL',
        failedStage: 'domainRegistry'
      };
    }
  });
  const response = JSON.parse(JSON.stringify(authorization.response));

  assert.equal(authorization.ok, false);
  assert.equal(response.errorCode, 'PESSOAS_V2_DB_INDISPONIVEL');
  assert.equal(response.details.etapa, 'domainRegistry');
  assert.equal(response.details.traceId, 'req-session-failure');
  assert.equal(response.details.ambienteEfetivo, 'DEV');
  assert.equal(response.details.versaoBackend, 'CORE_PROFILE_SESSION_ENV_V1');
  assert.deepEqual(Object.keys(response.details).sort(), [
    'ambienteEfetivo',
    'errorCode',
    'etapa',
    'traceId',
    'versaoBackend'
  ]);
  assert.equal(JSON.stringify(response).includes('pessoa.teste@example.org'), false);
});
