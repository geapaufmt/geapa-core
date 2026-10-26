const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const counters = { open: 0, read: 0 };
const sheets = {};

function makeSheet(name) {
  return sheets[name] || (sheets[name] = {
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
    DEV: { id: 'pessoas-dev', sheet: 'PESSOAS_BASE', ativo: true, lineNo: 2 },
    PROD: { id: 'pessoas-prod', sheet: 'PESSOAS_BASE', ativo: true, lineNo: 3 }
  },
  PESSOAS_V2_RESUMO_OPERACIONAL: {
    DEV: { id: 'pessoas-dev', sheet: 'PESSOAS_RESUMO_OPERACIONAL', ativo: true, lineNo: 4 },
    PROD: { id: 'pessoas-prod', sheet: 'PESSOAS_RESUMO_OPERACIONAL', ativo: true, lineNo: 5 }
  }
};

const context = {
  console,
  Object,
  Array,
  String,
  Number,
  Date,
  Math,
  JSON,
  isFinite,
  isNaN,
  Logger: { log() {} },
  core_getCurrentEnv_: () => 'PROD',
  core_getRegistryRaw_: () => registryRaw,
  core_openSpreadsheetById_(id) {
    counters.open += 1;
    return makeSpreadsheet(id);
  },
  core_readSheetRecords_() {
    counters.read += 1;
    return [{ ID_PESSOA: 'PES-1', VALOR: 'ok' }];
  },
  core_domainsV2AuditNewReport_: (name) => ({
    nome: name,
    totalErros: 0,
    erros: [],
    avisos: []
  }),
  core_domainsV2AuditIssue_(report, level, code, message, details) {
    if (level === 'ERRO') {
      report.totalErros += 1;
      report.erros.push({ code, message, details });
    }
  },
  core_buildHeaderIndexMap_: () => ({}),
  core_normalizeHeader_: (value) => String(value || '').toUpperCase(),
  core_logWarn_() {},
  core_runId_: () => 'TEST'
};
vm.createContext(context);

for (const file of [
  '28a_core_domains_v2_resolver.js',
  '31_core_domains_v2_operational_api.js'
]) {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), context, { filename: file });
}

const devReport = context.core_domainsV2NewReadReport_('DEV');
context.core_domainsV2OpenPessoasSubset_(['RESUMO'], devReport, { ambiente: 'DEV' });
context.core_domainsV2OpenPessoasSubset_(['RESUMO'], devReport, { ambiente: 'DEV' });
assert.equal(counters.open, 1, 'a mesma planilha DEV deve abrir uma vez por execucao');
assert.equal(counters.read, 1, 'a mesma aba DEV deve ser lida uma vez por execucao');
assert.equal(devReport.ambiente, 'DEV');

const prodReport = context.core_domainsV2NewReadReport_('PROD');
context.core_domainsV2OpenPessoasSubset_(['RESUMO'], prodReport, { ambiente: 'PROD' });
assert.equal(counters.open, 2, 'PROD deve possuir cache separado de DEV');
assert.equal(counters.read, 2, 'os registros PROD nao podem reutilizar o cache DEV');
assert.equal(prodReport.ambiente, 'PROD');

const membersSource = fs.readFileSync(path.join(root, '11_core_members.gs'), 'utf8');
const exportsSource = fs.readFileSync(path.join(root, '20_public_exports.js'), 'utf8');
const portalSource = fs.readFileSync(path.join(root, '26_core_portal_access.js'), 'utf8');
const operationalSource = fs.readFileSync(path.join(root, '31_core_domains_v2_operational_api.js'), 'utf8');
assert.match(membersSource, /opts\.sessao\s*\|\|\s*opts\.session/);
assert.match(membersSource, /allowLegacyFallback === false/);
assert.match(exportsSource, /geapaCoreBuscarMinhaSituacaoParaPortal\(emailOuRga, options\)/);
assert.match(portalSource, /corePortalReadRegistryRecordsForEnv_/);
assert.match(portalSource, /portalConfigCacheKey \+ ':' \+ corePortalResolveEnvironment_/);
assert.match(operationalSource, /pageItems = filtered\.slice\(start, start \+ normalizedFilters\.pageSize\)/);
assert.match(operationalSource, /core_domainsV2OpenPessoasSubset_\(\['RESUMO'\]/);

console.log(JSON.stringify({
  ok: true,
  cacheExecucao: { opensDev: 1, readsDev: 1, opensAposProd: 2, readsAposProd: 2 },
  ambienteExplicito: true,
  sessaoReutilizadaEmMinhaSituacao: true,
  fallbackCaroDesabilitavel: true,
  paginacaoAdministrativaBackend: true
}, null, 2));
