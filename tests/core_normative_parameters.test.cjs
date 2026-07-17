const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', '38_core_normative_parameters.js'), 'utf8');
const context = {
  Object, Array, String, Number, Error, JSON, isFinite,
  core_getCurrentEnv_: () => 'PROD',
  core_getRegistryRaw_: () => { throw new Error('REGISTRY_REAL_PROIBIDO'); },
  core_getSheetById_: () => { throw new Error('PLANILHA_REAL_PROIBIDA'); },
  core_readSheetRecords_: () => { throw new Error('PLANILHA_REAL_PROIBIDA'); }
};
vm.createContext(context);
vm.runInContext(source, context, { filename: '38_core_normative_parameters.js' });

function registry(environment = 'DEV') {
  return {
    NORMAS_PARAMETROS_OPERACIONAIS: {
      [environment]: {
        key: 'NORMAS_PARAMETROS_OPERACIONAIS', id: `fake-${environment}`, sheet: 'PARAMETROS',
        ambiente: environment, ativo: true, lineNo: 2, duplicateActiveLines: [2]
      }
    }
  };
}

function records(suspensionValue = 37, blockValue = 11) {
  return [
    { PARAMETRO_ID: 'SUSPENSAO_MINIMA', VALOR: suspensionValue, UNIDADE: 'DIAS', VIGENTE: 'SIM', BASE_LEGAL: 'NORMA-X-ART-A', MODULO_SISTEMA: 'GEAPA_MEMBROS' },
    { PARAMETRO_ID: 'BLOQUEIO_DESLIGAMENTO_ANTES_APRESENTACAO', VALOR: blockValue, UNIDADE: 'DIAS', VIGENTE: 'SIM', BASE_LEGAL: 'NORMA-X-ART-B', MODULO_SISTEMA: 'MEMBROS' }
  ];
}

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test('resolve valores dinamicos sem assumir o valor institucional vigente', () => {
  const result = context.core_resolverParametrosNormativosOperacionais_(
    ['SUSPENSAO_MINIMA', 'BLOQUEIO_DESLIGAMENTO_ANTES_APRESENTACAO'],
    { ambiente: 'DEV', registryRaw: registry('DEV'), records: records(37, 11), moduloSistema: 'GEAPA_MEMBROS' }
  );
  assert.equal(result.SUSPENSAO_MINIMA.valor, 37);
  assert.equal(result.BLOQUEIO_DESLIGAMENTO_ANTES_APRESENTACAO.valor, 11);
});

test('DEV nao usa entrada PROD', () => {
  assert.throws(
    () => context.core_resolverParametroNormativoOperacional_('SUSPENSAO_MINIMA', { ambiente: 'DEV', registryRaw: registry('PROD'), records: records() }),
    (error) => error.code === 'NORMAS_PARAMETROS_REGISTRY_DEV_AUSENTE'
  );
});

test('ambiente precisa ser explicito', () => {
  assert.throws(
    () => context.core_resolverParametroNormativoOperacional_('SUSPENSAO_MINIMA', { registryRaw: registry('PROD'), records: records() }),
    (error) => error.code === 'PARAMETRO_NORMATIVO_AMBIENTE_INVALIDO'
  );
});

for (const [field, value, code] of [
  ['VALOR', 0, 'PARAMETRO_NORMATIVO_VALOR_INVALIDO'],
  ['UNIDADE', 'MESES', 'PARAMETRO_NORMATIVO_UNIDADE_INVALIDA'],
  ['BASE_LEGAL', '', 'PARAMETRO_NORMATIVO_BASE_LEGAL_AUSENTE'],
  ['MODULO_SISTEMA', 'OUTRO_MODULO', 'PARAMETRO_NORMATIVO_MODULO_INCOMPATIVEL']
]) {
  test(`rejeita ${field} invalido`, () => {
    const rows = records();
    rows[0][field] = value;
    assert.throws(
      () => context.core_resolverParametroNormativoOperacional_('SUSPENSAO_MINIMA', { ambiente: 'DEV', registryRaw: registry('DEV'), records: rows }),
      (error) => error.code === code
    );
  });
}

test('rejeita linha nao vigente', () => {
  const rows = records();
  rows[0].VIGENTE = 'NAO';
  assert.throws(
    () => context.core_resolverParametroNormativoOperacional_('SUSPENSAO_MINIMA', { ambiente: 'DEV', registryRaw: registry('DEV'), records: rows }),
    (error) => error.code === 'PARAMETRO_NORMATIVO_NAO_VIGENTE'
  );
});

test('invalida somente cache', () => {
  const result = context.core_invalidarCacheParametrosNormativosOperacionais_({ ambiente: 'DEV' });
  assert.equal(result.ok, true);
  assert.equal(result.somenteCache, true);
});

let failed = 0;
for (const { name, fn } of tests) {
  try { fn(); console.log(`ok - ${name}`); }
  catch (error) { failed += 1; console.error(`not ok - ${name}`); console.error(error.stack || error); }
}
console.log(JSON.stringify({ ok: failed === 0, total: tests.length, failed }));
if (failed) process.exitCode = 1;
