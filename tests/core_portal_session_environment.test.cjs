const assert = require('node:assert/strict');
const {test} = require('node:test');
const {isolatedDomain, enableCompatibility} = require('./canonical-domain.cjs');
const counters = { open: 0, read: 0 };
const sheets = {};

function makeSheet(name) {
  return sheets[name] || (sheets[name] = {
    name,
    getLastColumn: () => 2,
    getRange: () => ({ getDisplayValues: () => [['ID_PESSOA', 'VALOR']] })
  });
}

function makeSpreadsheet(prefix) {
  return {
    getSheetByName(name) {
      return makeSheet(prefix + ':' + name);
    }
  };
}

const registryRaw = {
  PESSOAS_V2_DB: {
    DEV: { id: 'synthetic-directory', sheet: 'PESSOAS_BASE', ativo: true, lineNo: 2 },
    PROD: { id: 'synthetic-directory', sheet: 'PESSOAS_BASE', ativo: true, lineNo: 3 }
  },
  PESSOAS_V2_RESUMO_OPERACIONAL: {
    DEV: { id: 'synthetic-directory', sheet: 'PESSOAS_RESUMO_OPERACIONAL', ativo: true, lineNo: 4 },
    PROD: { id: 'synthetic-directory', sheet: 'PESSOAS_RESUMO_OPERACIONAL', ativo: true, lineNo: 5 }
  }
};

test('cache de compatibilidade isola leituras repetidas por ambiente', () => {
  const domain = isolatedDomain();
  const {context} = domain;
  assert.throws(() => context.core_domainsV2OpenPessoasSubset_(['RESUMO'], {}, {ambiente:'DEV'}), {code:'MEMBERS_LEGACY_FENCED'});
  enableCompatibility(domain);
  let environment = 'DEV';
  Object.assign(context, {
    core_getCurrentEnv_: () => 'PROD',
    core_getRegistryRaw_: () => registryRaw,
    core_openSpreadsheetById_(id) {
      counters.open += 1;
      return makeSpreadsheet(id + environment);
    },
    core_readSheetRecords_(sheet) {
      counters.read += 1;
      return [{ ID_PESSOA: 'PES-1', VALOR: sheet.name }];
    },
    core_logWarn_() {},
    core_runId_: () => 'TEST'
  });
  const devReport = context.core_domainsV2NewReadReport_('DEV');
  const dev = context.core_domainsV2OpenPessoasSubset_(['RESUMO'], devReport, { ambiente: 'DEV' });
  context.core_domainsV2OpenPessoasSubset_(['RESUMO'], devReport, { ambiente: 'DEV' });
  assert.equal(counters.open, 1, 'a mesma planilha DEV deve abrir uma vez por execucao');
  assert.equal(counters.read, 1, 'a mesma aba DEV deve ser lida uma vez por execucao');
  assert.equal(devReport.ambiente, 'DEV');

  environment = 'PROD';
  const prodReport = context.core_domainsV2NewReadReport_('PROD');
  const prod = context.core_domainsV2OpenPessoasSubset_(['RESUMO'], prodReport, { ambiente: 'PROD' });
  assert.equal(counters.open, 2, 'PROD deve possuir cache separado de DEV');
  assert.equal(counters.read, 2, 'os registros PROD nao podem reutilizar o cache DEV');
  assert.equal(prodReport.ambiente, 'PROD');
  assert.equal(dev.PESSOAS_RESUMO_OPERACIONAL.records[0].VALOR,'synthetic-directoryDEV:PESSOAS_RESUMO_OPERACIONAL');
  assert.equal(prod.PESSOAS_RESUMO_OPERACIONAL.records[0].VALOR,'synthetic-directoryPROD:PESSOAS_RESUMO_OPERACIONAL');
});

test('o Registry e a chave de cache do Portal exigem o ambiente pedido', () => {
  const {context} = isolatedDomain();
  assert.notEqual(context.corePortalConfigCacheKey_({ambiente:'DEV'}), context.corePortalConfigCacheKey_({ambiente:'PROD'}));
  for (const [requested, other] of [['DEV','PROD'], ['PROD','DEV']]) {
    const registryRaw = {SYN_KEY:{[other]:{id:'synthetic', sheet:'SYN', ativo:true, lineNo:2}}};
    assert.throws(() => context.corePortalReadRegistryRecordsForEnv_('SYN_KEY', {ambiente:requested, registryRaw}),
      {code:'PORTAL_REGISTRY_KEY_INDISPONIVEL'});
  }
});

test('Minha situação encaminha a sessão à leitura de resumo e respeita o bloqueio do fallback', () => {
  const domain = isolatedDomain();
  enableCompatibility(domain);
  const {context} = domain;
  const session = {idPessoa:'PES-1', nomeExibicao:'Pessoa sintética'};
  context.core_resolverSessaoPortalCompat_ = () => { throw Error('UNEXPECTED_SESSION_LOOKUP'); };
  context.core_domainsV2NewReadReport_ = () => ({totalErros:0});
  context.core_domainsV2OpenPessoasSubset_ = (keys, report, opts) => {
    assert.deepEqual(Array.from(keys), ['RESUMO']);
    assert.equal(opts.ambiente, 'DEV');
    return {PESSOAS_RESUMO_OPERACIONAL:{records:[]}};
  };
  for (const key of ['sessao','session']) {
    const result = context.geapaCoreBuscarMinhaSituacaoParaPortal('', {ambiente:'DEV', [key]:session});
    assert.equal(result.ok, true);
    assert.equal(result.sessao, session);
  }
  const denied = context.geapaCoreBuscarMinhaSituacaoParaPortal('', {sessao:{}, allowLegacyFallback:false});
  assert.equal(denied.ok, false);
  assert.equal(denied.code, 'MINHA_SITUACAO_V2_NAO_RESOLVIDA');
});
