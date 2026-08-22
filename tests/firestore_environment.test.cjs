const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const properties = new Map();
const context = {
  Object,
  Array,
  String,
  Number,
  Boolean,
  Date,
  JSON,
  Math,
  isFinite,
  encodeURIComponent,
  PropertiesService: {
    getScriptProperties() {
      return { getProperty: (key) => properties.get(key) || null };
    }
  },
  ScriptApp: { getOAuthToken: () => 'test-token' },
  UrlFetchApp: { fetch: () => { throw new Error('rede nao deve ser usada neste teste'); } }
};

vm.createContext(context);
vm.runInContext(
  fs.readFileSync(path.join(__dirname, '..', '24_core_firestore_rest.js'), 'utf8'),
  context,
  { filename: '24_core_firestore_rest.js' }
);

assert.throws(
  () => context.coreFirestoreGetEnvironmentConfig_({}),
  /Ambiente Firestore obrigatorio/
);
assert.throws(
  () => context.coreFirestoreGetEnvironmentConfig_({ environment: 'HOMOLOG' }),
  /Informe DEV ou PROD/
);
assert.throws(
  () => context.coreFirestoreGetEnvironmentConfig_({ environment: 'DEV' }),
  /GEAPA_CORE_FIRESTORE_DEV_PROJECT_ID nao configurado/
);

properties.set('GEAPA_CORE_FIRESTORE_DEV_PROJECT_ID', 'geapa-dev-test');
properties.set('GEAPA_CORE_FIRESTORE_DEV_DATABASE_ID', '(default)');
let config = context.coreFirestoreGetEnvironmentConfig_({ ambiente: 'dev' });
assert.equal(config.environment, 'DEV');
assert.equal(config.projectId, 'geapa-dev-test');
assert.equal(config.namespaced, false);

let devDryRun = context.coreFirestoreEnvironmentSetDocument_(
  'activities/ATV-2026-1-0001',
  { idAtividade: 'ATV-2026-1-0001' },
  { environment: 'DEV' }
);
assert.equal(devDryRun.ok, true);
assert.equal(devDryRun.dryRun, true);
assert.equal(devDryRun.written, false);

properties.set('GEAPA_CORE_FIRESTORE_PROD_PROJECT_ID', 'geapa-prod-test');
const prodWrite = context.coreFirestoreEnvironmentSetDocument_(
  'activities/ATV-2026-1-0001',
  { idAtividade: 'ATV-2026-1-0001' },
  { environment: 'PROD', dryRun: false }
);
assert.equal(prodWrite.ok, false);
assert.match(prodWrite.message, /PROD estao bloqueadas/);

properties.set('GEAPA_CORE_FIRESTORE_PROD_PROJECT_ID', 'geapa-dev-test');
assert.throws(
  () => context.coreFirestoreGetEnvironmentConfig_({ environment: 'DEV' }),
  /projetos Firebase diferentes/
);

console.log('firestore_environment.test.cjs: OK');
