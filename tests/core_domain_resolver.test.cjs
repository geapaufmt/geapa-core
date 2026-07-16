const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, '28a_core_domains_v2_resolver.js'), 'utf8');
const context = {
  Object,
  Array,
  String,
  Error,
  JSON,
  Logger: { log() {} },
  core_getCurrentEnv_: () => 'PROD',
  core_getRegistryRaw_: () => ({}),
  core_openSpreadsheetById_: () => { throw new Error('OPEN_REAL_FORBIDDEN'); }
};
vm.createContext(context);
vm.runInContext(source, context, { filename: '28a_core_domains_v2_resolver.js' });

function entry(key, id, sheet, ambiente, extra = {}) {
  return Object.assign({
    key,
    id,
    sheet,
    ambiente,
    ativo: true,
    lineNo: extra.lineNo || 2,
    duplicateActiveLines: [extra.lineNo || 2]
  }, extra);
}

function registry(rows) {
  const out = {};
  rows.forEach((row) => {
    out[row.key] ||= {};
    out[row.key][row.ambiente] = row;
  });
  return out;
}

function spreadsheet(names) {
  const sheets = Object.fromEntries(names.map((name) => [name, { name }]));
  return { getSheetByName: (name) => sheets[name] || null };
}

function options(raw, books, environment, counters = {}) {
  return {
    ambiente: environment,
    registryRaw: raw,
    warnings: [],
    openSpreadsheetById(id) {
      counters[id] = (counters[id] || 0) + 1;
      if (!books[id]) throw new Error(`UNKNOWN_FAKE_SPREADSHEET:${id}`);
      return books[id];
    }
  };
}

function reset() {
  context.core_resetDomainResolverExecutionCache_();
}

const tests = [];
function test(name, fn) { tests.push({ name, fn }); }

test('mesma key em DEV e PROD resolve somente o ambiente pedido', () => {
  reset();
  const raw = registry([
    entry('PESSOAS_V2_DB', 'fake-dev-db', 'PESSOAS_BASE', 'DEV'),
    entry('PESSOAS_V2_DB', 'fake-prod-db', 'PESSOAS_BASE', 'PROD')
  ]);
  const books = {
    'fake-dev-db': spreadsheet(['PESSOAS_BASE']),
    'fake-prod-db': spreadsheet(['PESSOAS_BASE'])
  };
  const dev = context.core_getDomainSheet_('PESSOAS', 'BASE', Object.assign(options(raw, books, 'DEV'), { includeResolution: true }));
  const prod = context.core_getDomainSheet_('PESSOAS', 'BASE', Object.assign(options(raw, books, 'PROD'), { includeResolution: true }));
  assert.equal(dev.resolution.spreadsheetId, 'fake-dev-db');
  assert.equal(prod.resolution.spreadsheetId, 'fake-prod-db');
});

test('duplicidade ativa da mesma key e ambiente e bloqueante', () => {
  reset();
  const duplicate = entry('PESSOAS_V2_DB', 'fake-dev-db', 'PESSOAS_BASE', 'DEV', { duplicateActiveLines: [2, 3] });
  const raw = registry([duplicate]);
  assert.throws(() => context.core_getDomainSpreadsheetRef_('PESSOAS', { ambiente: 'DEV', registryRaw: raw }), (err) => err.code === 'DOMAIN_REGISTRY_DUPLICADO');
});

test('DB ausente permite fallback especifico somente para leitura', () => {
  reset();
  const raw = registry([entry('PESSOAS_V2_BASE', 'fake-legacy-dev', 'PESSOAS_BASE', 'DEV')]);
  const opts = options(raw, { 'fake-legacy-dev': spreadsheet(['PESSOAS_BASE']) }, 'DEV');
  const result = context.core_getDomainSheet_('PESSOAS', 'BASE', Object.assign(opts, { includeResolution: true }));
  assert.equal(result.resolution.origin, 'SPECIFIC_KEY_FALLBACK');
  assert.equal(opts.warnings[0].code, 'DOMAIN_SPECIFIC_KEY_FALLBACK');
});

test('DB ausente bloqueia escrita mesmo com key especifica', () => {
  reset();
  const raw = registry([entry('PESSOAS_V2_BASE', 'fake-legacy-dev', 'PESSOAS_BASE', 'DEV')]);
  const opts = options(raw, { 'fake-legacy-dev': spreadsheet(['PESSOAS_BASE']) }, 'DEV');
  assert.throws(() => context.core_getDomainSheet_('PESSOAS', 'BASE', Object.assign(opts, { forWrite: true })), (err) => err.code === 'DOMAIN_WRITE_DB_INDISPONIVEL');
});

test('aba canonica ausente usa fallback na leitura e falha na escrita', () => {
  reset();
  const raw = registry([
    entry('PESSOAS_V2_DB', 'fake-dev-db', 'PESSOAS_BASE', 'DEV'),
    entry('PESSOAS_V2_LINKS_PERFIS', 'fake-dev-db', 'PESSOAS_V2_LINKS_PERFIS', 'DEV')
  ]);
  const books = { 'fake-dev-db': spreadsheet(['PESSOAS_V2_LINKS_PERFIS']) };
  const read = context.core_getDomainSheet_('PESSOAS', 'LINKS_PERFIS', Object.assign(options(raw, books, 'DEV'), { includeResolution: true }));
  assert.equal(read.resolution.origin, 'SPECIFIC_KEY_FALLBACK');
  reset();
  assert.throws(() => context.core_getDomainSheet_('PESSOAS', 'LINKS_PERFIS', Object.assign(options(raw, books, 'DEV'), { forWrite: true })), (err) => err.code === 'DOMAIN_WRITE_ABA_CANONICA_AUSENTE');
});

test('divergencia DB versus especifica avisa na leitura e bloqueia escrita', () => {
  reset();
  const raw = registry([
    entry('PESSOAS_V2_DB', 'fake-dev-db', 'PESSOAS_BASE', 'DEV'),
    entry('PESSOAS_V2_BASE', 'fake-other-dev', 'PESSOAS_BASE', 'DEV')
  ]);
  const books = {
    'fake-dev-db': spreadsheet(['PESSOAS_BASE']),
    'fake-other-dev': spreadsheet(['PESSOAS_BASE'])
  };
  const readOpts = options(raw, books, 'DEV');
  const read = context.core_getDomainSheet_('PESSOAS', 'BASE', Object.assign(readOpts, { includeResolution: true }));
  assert.equal(read.resolution.origin, 'DOMAIN_DB');
  assert.equal(readOpts.warnings[0].code, 'DOMAIN_REGISTRY_DIVERGENCIA');
  reset();
  assert.throws(() => context.core_getDomainSheet_('PESSOAS', 'BASE', Object.assign(options(raw, books, 'DEV'), { forWrite: true })), (err) => err.code === 'DOMAIN_WRITE_REGISTRY_DIVERGENTE');
});

test('abertura da planilha inteira para escrita bloqueia qualquer key especifica divergente', () => {
  reset();
  const raw = registry([
    entry('ATIVIDADES_V2_DB', 'fake-atividades-dev', 'Atividades', 'DEV'),
    entry('ATIVIDADES_V2_APRESENTACOES', 'fake-atividades-legada', 'Atividades_Apresentacoes', 'DEV')
  ]);
  const books = {
    'fake-atividades-dev': spreadsheet(['Atividades']),
    'fake-atividades-legada': spreadsheet(['Atividades_Apresentacoes'])
  };
  assert.throws(
    () => context.core_openDomainSpreadsheet_('ATIVIDADES', Object.assign(options(raw, books, 'DEV'), { forWrite: true })),
    (err) => err.code === 'DOMAIN_WRITE_REGISTRY_DIVERGENTE'
  );
});

test('planilha do dominio e aberta uma vez por execucao', () => {
  reset();
  const raw = registry([entry('PESSOAS_V2_DB', 'fake-dev-db', 'PESSOAS_BASE', 'DEV')]);
  const counters = {};
  const opts = options(raw, { 'fake-dev-db': spreadsheet(['PESSOAS_BASE', 'MEMBROS_DETALHES']) }, 'DEV', counters);
  context.core_getDomainSheet_('PESSOAS', 'BASE', opts);
  context.core_getDomainSheet_('PESSOAS', 'MEMBROS_DETALHES', opts);
  assert.equal(counters['fake-dev-db'], 1);
});

test('nao existe fallback cruzado DEV para PROD', () => {
  reset();
  const raw = registry([
    entry('PESSOAS_V2_DB', 'fake-prod-db', 'PESSOAS_BASE', 'PROD'),
    entry('PESSOAS_V2_BASE', 'fake-prod-db', 'PESSOAS_BASE', 'PROD')
  ]);
  const opts = options(raw, { 'fake-prod-db': spreadsheet(['PESSOAS_BASE']) }, 'DEV');
  assert.throws(() => context.core_getDomainSheet_('PESSOAS', 'BASE', opts), (err) => err.code === 'DOMAIN_SHEET_INDISPONIVEL');
});

test('ambiente invalido e rejeitado', () => {
  reset();
  assert.throws(() => context.core_getDomainSpreadsheetRef_('PESSOAS', { ambiente: 'HOMOLOG', registryRaw: {} }), (err) => err.code === 'DOMAIN_ENV_INVALIDO');
});

test('codigo operacional nao contem IDs reais nem resolvedor antigo', () => {
  const operational = fs.readdirSync(root)
    .filter((name) => /\.(?:js|gs)$/.test(name))
    .map((name) => fs.readFileSync(path.join(root, name), 'utf8'))
    .join('\n');
  assert.doesNotMatch(operational, /1sa1CZTsqdDEWKWLd5uDAiM-Y59ko9FLZfABL0wc0HVM|1M_KPFn7sRjZmQMtfoVOSDuSwlJqq9BLBQ-UYahcDQJw/);
  assert.doesNotMatch(operational, /CORE_DOMAINS_V2_DEV_SPREADSHEETS|CORE_V2_ROTINAS_KEYS|core_v2RotinasGetSheetByKey_/);
});

let failed = 0;
for (const { name, fn } of tests) {
  try {
    fn();
    console.log(`ok - ${name}`);
  } catch (err) {
    failed += 1;
    console.error(`not ok - ${name}`);
    console.error(err.stack || err);
  }
}
console.log(JSON.stringify({ ok: failed === 0, total: tests.length, failed }));
if (failed) process.exitCode = 1;
