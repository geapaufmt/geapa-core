const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', '38_core_normative_parameters.js'), 'utf8');
const cacheStore = new Map();
const context = {
  Object, Array, String, Number, Error, JSON, isFinite,
  core_getCurrentEnv_: () => 'PROD',
  core_getRegistryRaw_: () => { throw new Error('REGISTRY_REAL_PROIBIDO'); },
  core_getSheetById_: () => { throw new Error('PLANILHA_REAL_PROIBIDA'); },
  core_readSheetRecords_: (sheet) => sheet.records,
  CacheService: {
    getScriptCache: () => ({
      get: (key) => cacheStore.get(key) || null,
      put: (key, value) => cacheStore.set(key, value),
      remove: (key) => cacheStore.delete(key)
    })
  }
};
vm.createContext(context);
vm.runInContext(source, context, { filename: '38_core_normative_parameters.js' });

function registry(environment = 'DEV', duplicate = false) {
  return {
    NORMAS_PARAMETROS_OPERACIONAIS: {
      [environment]: {
        key: 'NORMAS_PARAMETROS_OPERACIONAIS', id: `fake-${environment}`, sheet: 'PARAMETROS',
        ambiente: environment, ativo: true, lineNo: 2, duplicateActiveLines: duplicate ? [2, 3] : [2]
      }
    }
  };
}

function numeric(id, value, extra = {}) {
  return Object.assign({
    PARAMETRO_ID: id, VALOR: value, UNIDADE: 'DIAS', VIGENTE: 'SIM',
    BASE_LEGAL: `NORMA-${id}`, MODULO_SISTEMA: 'GEAPA_MEMBROS'
  }, extra);
}

function ata(value = 'SIM', extra = {}) {
  return Object.assign({
    PARAMETRO_ID: 'DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA',
    TIPO_VALOR: 'BOOLEANO', VALOR: value, UNIDADE: 'NAO_APLICAVEL',
    VIGENTE: 'SIM', BASE_LEGAL: value === 'NAO' || value === false || value === 0 ? 'NC02-2027-ART4' : 'NC01-2025-ART16-IV',
    MODULO_SISTEMA: 'GEAPA_MEMBROS'
  }, extra);
}

function records() {
  return [
    numeric('SUSPENSAO_MINIMA', 37),
    numeric('BLOQUEIO_DESLIGAMENTO_ANTES_APRESENTACAO', 11, { TIPO_VALOR: 'NUMERO' }),
    ata('SIM')
  ];
}

function resolve(id, rows = records(), options = {}) {
  return context.core_resolverParametroNormativoOperacional_(id, Object.assign({
    ambiente: 'DEV', registryRaw: registry('DEV'), records: rows, moduloSistema: 'GEAPA_MEMBROS'
  }, options));
}

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test('NUMERO legado sem TIPO_VALOR permanece compativel', () => {
  const result = resolve('SUSPENSAO_MINIMA');
  assert.equal(result.tipoValor, 'NUMERO');
  assert.equal(result.valor, 37);
  assert.equal(result.unidade, 'DIAS');
});

test('NUMERO explicito permanece compativel', () => {
  const result = resolve('BLOQUEIO_DESLIGAMENTO_ANTES_APRESENTACAO');
  assert.equal(result.tipoValor, 'NUMERO');
  assert.equal(result.valor, 11);
});

test('BOOLEANO SIM e NAO normalizam para boolean', () => {
  assert.equal(resolve('DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA', [ata('SIM')]).valor, true);
  assert.equal(resolve('DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA', [ata('NAO')]).valor, false);
});

test('BOOLEANO TRUE e FALSE normalizam para boolean', () => {
  assert.equal(resolve('DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA', [ata('TRUE')]).valor, true);
  assert.equal(resolve('DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA', [ata('FALSE', { BASE_LEGAL: 'NC02-2027-ART4' })]).valor, false);
});

test('BOOLEANO 1 e 0 normalizam para boolean', () => {
  assert.equal(resolve('DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA', [ata(1)]).valor, true);
  assert.equal(resolve('DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA', [ata(0)]).valor, false);
});

test('rejeita BOOLEANO ambiguo', () => {
  assert.throws(() => resolve('DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA', [ata('TALVEZ')]),
    (error) => error.code === 'PARAMETRO_NORMATIVO_VALOR_BOOLEANO_INVALIDO');
});

test('rejeita tipo desconhecido', () => {
  assert.throws(() => resolve('SUSPENSAO_MINIMA', [numeric('SUSPENSAO_MINIMA', 30, { TIPO_VALOR: 'TEXTO' })]),
    (error) => error.code === 'PARAMETRO_NORMATIVO_TIPO_VALOR_INVALIDO');
});

test('rejeita parametro ausente', () => {
  assert.throws(() => resolve('SUSPENSAO_MINIMA', []),
    (error) => error.code === 'PARAMETRO_NORMATIVO_NAO_ENCONTRADO');
});

test('rejeita parametro vigente duplicado', () => {
  assert.throws(() => resolve('SUSPENSAO_MINIMA', [numeric('SUSPENSAO_MINIMA', 30), numeric('SUSPENSAO_MINIMA', 31)]),
    (error) => error.code === 'PARAMETRO_NORMATIVO_VIGENTE_DUPLICADO');
});

test('rejeita key duplicada no mesmo ambiente', () => {
  assert.throws(() => resolve('SUSPENSAO_MINIMA', records(), { registryRaw: registry('DEV', true) }),
    (error) => error.code === 'NORMAS_PARAMETROS_REGISTRY_DUPLICADO');
});

test('rejeita linha nao vigente', () => {
  assert.throws(() => resolve('SUSPENSAO_MINIMA', [numeric('SUSPENSAO_MINIMA', 30, { VIGENTE: 'NAO' })]),
    (error) => error.code === 'PARAMETRO_NORMATIVO_NAO_VIGENTE');
});

test('rejeita BASE_LEGAL ausente e modulo incompativel', () => {
  assert.throws(() => resolve('SUSPENSAO_MINIMA', [numeric('SUSPENSAO_MINIMA', 30, { BASE_LEGAL: '' })]),
    (error) => error.code === 'PARAMETRO_NORMATIVO_BASE_LEGAL_AUSENTE');
  assert.throws(() => resolve('SUSPENSAO_MINIMA', [numeric('SUSPENSAO_MINIMA', 30, { MODULO_SISTEMA: 'OUTRO' })]),
    (error) => error.code === 'PARAMETRO_NORMATIVO_MODULO_INCOMPATIVEL');
});

test('NAO nao pode reutilizar base legal que exige ata', () => {
  assert.throws(() => resolve('DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA', [ata('NAO', { BASE_LEGAL: 'NC01-2025-ART16-IV' })]),
    (error) => error.code === 'PARAMETRO_NORMATIVO_BASE_LEGAL_INCOMPATIVEL');
});

test('DEV nao usa entrada PROD', () => {
  assert.throws(() => resolve('SUSPENSAO_MINIMA', records(), { registryRaw: registry('PROD') }),
    (error) => error.code === 'NORMAS_PARAMETROS_REGISTRY_DEV_AUSENTE');
});

test('cache inclui e invalida o novo parametro', () => {
  cacheStore.clear();
  const sheet = { records: [ata('SIM')] };
  context.core_getRegistryRaw_ = () => registry('DEV');
  const options = { ambiente: 'DEV', moduloSistema: 'GEAPA_MEMBROS', openSheet: () => sheet };
  assert.equal(context.core_resolverParametroNormativoOperacional_('DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA', options).valor, true);
  sheet.records = [ata('NAO')];
  assert.equal(context.core_resolverParametroNormativoOperacional_('DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA', options).valor, true);
  const invalidated = context.core_invalidarCacheParametrosNormativosOperacionais_({ ambiente: 'DEV' });
  assert.equal(invalidated.chavesRemovidas, 3);
  assert.equal(context.core_resolverParametroNormativoOperacional_('DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA', options).valor, false);
});

test('setup dry-run recomenda TIPO_VALOR e nova linha sem escrever', () => {
  const result = context.core_prepararParametrosNormativosTipados_({
    ambiente: 'DEV', headers: ['PARAMETRO_ID', 'VALOR', 'UNIDADE', 'BASE_LEGAL', 'MODULO_SISTEMA', 'VIGENTE'], records: []
  });
  assert.equal(result.dryRun, true);
  assert.equal(result.escritaExecutada, false);
  assert.ok(Array.from(result.cabecalhosAusentes).includes('TIPO_VALOR'));
  assert.equal(result.linhaRecomendada.VALOR, 'SIM');
  assert.equal(result.linhaRecomendada.TIPO_VALOR, 'BOOLEANO');
});

let failed = 0;
for (const { name, fn } of tests) {
  try { fn(); console.log(`ok - ${name}`); }
  catch (error) { failed += 1; console.error(`not ok - ${name}`); console.error(error.stack || error); }
}
console.log(JSON.stringify({ ok: failed === 0, total: tests.length, failed }));
if (failed) process.exitCode = 1;
