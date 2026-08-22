const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const context = {
  console,
  Object,
  Array,
  String,
  Number,
  Boolean,
  Date,
  JSON,
  Math,
  isFinite,
  isNaN,
  core_extractEmailAddress_: (value) => String(value || '').trim().toLowerCase(),
  core_isValidEmail_: (value) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '')),
  core_normalizeText_: (value, opts) => {
    const normalized = String(value || '').trim().replace(/\s+/g, ' ');
    return opts && opts.caseMode === 'upper' ? normalized.toUpperCase() : normalized;
  }
};
vm.createContext(context);
vm.runInContext(
  fs.readFileSync(path.join(__dirname, '..', '26_core_portal_access.js'), 'utf8'),
  context,
  { filename: '26_core_portal_access.js' }
);

const canonicalSession = Object.freeze({
  ok: true,
  autenticado: true,
  portalAtivo: true,
  idPessoa: 'PES-1',
  email: 'canonico@geapa.test',
  nomeExibicao: 'Pessoa Teste',
  perfilPortalEfetivo: 'MEMBRO',
  perfisPortal: ['MEMBRO'],
  permissoes: ['portal.acessar']
});
const aliasIdentity = Object.freeze({
  uid: 'firebase-uid-1',
  email: 'alias@geapa.test',
  emailVerified: true
});

context.corePortalFirestoreEnvironment_ = (opts) => {
  assert.equal(String(opts.ambiente || opts.environment).toUpperCase(), 'DEV');
  return 'DEV';
};

let resolverCalls = 0;
let syncCall = null;
context.corePortalResolverUsuarioAtual_ = ({ email }) => {
  resolverCalls += 1;
  assert.equal(email, 'alias@geapa.test');
  return canonicalSession;
};
context.corePortalSincronizarUsuarioFirestore_ = (entrada, opts) => {
  syncCall = { entrada, opts };
  return Object.freeze({ ok: true, synced: true, code: 'FIRESTORE_SYNC_OK' });
};

const aliasResult = context.corePortalProvisionarFirestoreUserAutenticado_(aliasIdentity, {
  ambiente: 'DEV',
  identityVerified: true,
  sessao: canonicalSession
});
assert.equal(aliasResult.ok, true, 'alias oficial deve ser aceito');
assert.equal(resolverCalls, 1, 'alias deve ser revalidado pelo proprio Core');
assert.equal(syncCall.opts.authenticatedEmail, 'alias@geapa.test');
assert.equal(syncCall.opts.sessao.idPessoa, 'PES-1');

const snapshot = context.corePortalBuildFirestoreUserSnapshot_(
  { uid: 'firebase-uid-1' },
  {
    uid: 'firebase-uid-1',
    sessao: canonicalSession,
    authenticatedEmail: 'alias@geapa.test',
    cacheTtlMs: 1000
  }
);
assert.equal(snapshot.email, 'alias@geapa.test');
assert.equal(snapshot.emailNormalizado, 'alias@geapa.test');
assert.equal(snapshot.idPessoa, 'PES-1');

context.corePortalResolverUsuarioAtual_ = () => Object.assign({}, canonicalSession, { idPessoa: 'PES-2' });
syncCall = null;
const divergentResult = context.corePortalProvisionarFirestoreUserAutenticado_(aliasIdentity, {
  ambiente: 'DEV',
  identityVerified: true,
  sessao: canonicalSession
});
assert.equal(divergentResult.ok, false);
assert.equal(divergentResult.code, 'IDENTIDADE_FIREBASE_DIVERGENTE');
assert.equal(syncCall, null, 'identidades divergentes nunca devem chegar a escrita');

let registryEnvironment = '';
let appended = false;
context.corePortalResolveEnvironment_ = (opts) => String(opts.ambiente || opts.environment || '').toUpperCase();
context.core_domainRegistryEntry_ = (key, environment) => {
  registryEnvironment = environment;
  return {
    available: true,
    entry: { ativo: true, id: 'log-dev-id', sheet: 'PORTAL_LOG_ACESSOS' }
  };
};
context.core_getSheetById_ = (id, sheet) => ({ id, sheet });
context.corePortalAppendAccessLogToSheet_ = (sheet) => {
  appended = true;
  assert.equal(sheet.id, 'log-dev-id');
};

const logResult = context.corePortalAppendAccessLog_({ acao: 'TESTE' }, { ambiente: 'DEV' });
assert.equal(logResult.ok, true);
assert.equal(logResult.ambiente, 'DEV');
assert.equal(registryEnvironment, 'DEV', 'log DEV nao pode consultar a entrada PROD');
assert.equal(appended, true);

console.log('core_portal_firestore_provision.test.cjs: OK');
