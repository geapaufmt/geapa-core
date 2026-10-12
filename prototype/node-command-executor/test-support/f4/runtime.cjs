'use strict';
// In-memory REST double only. No sockets, credentials, Firebase SDK or local emulator.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');
const mapping = {};
const root = path.resolve(__dirname, '../..');
const packageDir = process.env.F4_DEV_PACKAGE || path.resolve(__dirname,'../../base/core');
const legacyFiles = fs.readdirSync(root).filter(f => /\.(gs|js)$/.test(f) && (!packageDir || !['80_testes.js', '81_core_profile_updates_tests.js'].includes(f))).sort();
function createRuntime({ reference, legacy = true, devRunId, scriptId, environment = 'DEV', preconditionStatus = 409, preconditionCategory = 'FAILED_PRECONDITION' } = {}) {
  const documents = new Map(), calls = [], transactions = new Map();
  let sequence = 0, loseResponse = false;
  const prefix = devRunId ? 'https://firestore.googleapis.com/v1/projects/geapa-convergencia-dev/databases/(default)/documents/memberDevHomologation/' + devRunId : 'https://firestore.googleapis.com/v1/projects/demo-geapa-m04-local/databases/(default)/documents';
  const rpcPrefix = devRunId ? prefix.split('/memberDevHomologation/')[0] : prefix;
  const namePrefix = prefix.slice('https://firestore.googleapis.com/v1/'.length);
  const response = (status, value) => ({ getResponseCode: () => status, getContentText: () => JSON.stringify(value) });
  const deny = service => new Proxy({}, { get() { throw Error('FORBIDDEN_SERVICE_' + service); } });
  const context = vm.createContext({
    console,
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256', SHA_512: 'sha512' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (algorithm, value) => [...crypto.createHash(algorithm).update(typeof value === 'string' ? value : Buffer.from(value.map(x => x & 255))).digest()].map(x => x > 127 ? x - 256 : x),
      base64DecodeWebSafe: value => [...Buffer.from(value, 'base64url')].map(x => x > 127 ? x - 256 : x),
      sleep: () => {},
      formatDate(date, zone, format) {
        if (zone !== 'America/Cuiaba' || format !== "yyyy-MM-dd'T'HH:mm:ss") throw Error('UNEXPECTED_FORMAT');
        const p = Object.fromEntries(new Intl.DateTimeFormat('en-CA', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }).formatToParts(date).map(x => [x.type, x.value]));
        return `${p.year}-${p.month}-${p.day}T${p.hour}:${p.minute}:${p.second}`;
      }
    },
    ScriptApp: { getOAuthToken: () => 'SYNTHETIC_NO_CREDENTIAL', getScriptId: () => scriptId || (devRunId ? '1M3GpuWwyRkV24nbHnopzj1q4J-Ea7zHv5eCG-ngOw0qR5lWDMvvv8UfD' : 'LOCAL_UNPUBLISHED') },
    UrlFetchApp: { fetch(url, options) {
      if (!(url === prefix || url.startsWith(prefix + '/') || url.startsWith(rpcPrefix + ':')) || !['get', 'post'].includes(options.method)) throw Error('NETWORK_FORBIDDEN');
      if (options.followRedirects !== false || options.headers.Authorization !== 'Bearer SYNTHETIC_NO_CREDENTIAL') throw Error('INVALID_OPTIONS');
      const suffix = url.slice(url.startsWith(rpcPrefix + ':') ? rpcPrefix.length : prefix.length), body = options.payload ? JSON.parse(options.payload) : {};
      calls.push({ suffix, method: options.method, writes: body.writes?.length || 0 });
      if (suffix === ':beginTransaction') {
        const transaction = 'T' + ++sequence;
        transactions.set(transaction, new Map(documents));
        return response(200, { transaction });
      }
      if (suffix === ':rollback') { transactions.delete(body.transaction); return response(200, {}); }
      if (options.method === 'get') {
        const [encoded, query] = suffix.slice(1).split('?');
        const transaction = new URLSearchParams(query).get('transaction');
        const source = transaction ? transactions.get(transaction) : documents;
        if (!source) throw Error('UNKNOWN_TRANSACTION');
        const doc = source.get(decodeURIComponent(encoded));
        return doc ? response(200, doc) : response(404, { error: { status: 'NOT_FOUND' } });
      }
      if (suffix !== ':commit') throw Error('UNSUPPORTED_REST');
      for (const write of body.writes) {
        const name = write.update?.name || write.delete;
        if (!name.startsWith(namePrefix + '/')) throw Error('INVALID_NAME');
        const current = documents.get(decodeURIComponent(name.slice(namePrefix.length + 1))), pre = write.currentDocument;
        if (pre && (('exists' in pre && pre.exists !== !!current) || (pre.updateTime && pre.updateTime !== current?.updateTime))) return response(preconditionStatus, { error: { status: preconditionCategory } });
      }
      const stamp = new Date(Date.UTC(2026, 0, 1) + ++sequence).toISOString();
      for (const write of body.writes) {
        const key = decodeURIComponent((write.update?.name || write.delete).slice(namePrefix.length + 1));
        if (write.delete) documents.delete(key);
        else documents.set(key, { ...write.update, fields: write.updateMask ? { ...documents.get(key)?.fields, ...write.update.fields } : write.update.fields, updateTime: stamp });
      }
      transactions.delete(body.transaction);
      if (loseResponse) { loseResponse = false; throw Error('SYNTHETIC_LOST_RESPONSE'); }
      return response(200, { commitTime: stamp, writeResults: body.writes.map(() => ({ updateTime: stamp })) });
    } }
  });
  for (const service of ['DriveApp', 'GmailApp', 'MailApp', 'SpreadsheetApp', 'FormsApp', 'PropertiesService', 'LockService', 'Logger']) context[service] = deny(service);
  if (devRunId) context.PropertiesService = { getScriptProperties: () => ({ getProperty: key => key === 'GEAPA_ENV' ? environment : null }) };
  const load = filename => vm.runInContext(fs.readFileSync(filename, 'utf8'), context, { filename });
  const usePackage = packageDir && !reference;
  const sources = usePackage ? fs.readdirSync(packageDir, { recursive: true }).filter(f => f.endsWith('.gs')).sort().map(f => fs.readFileSync(path.join(packageDir, f), 'utf8')) : legacy ? legacyFiles.map(file => fs.readFileSync(path.join(process.env.F4_G_LEGACY_ROOT || root, file), 'utf8')) : [];
  for (const file of Object.keys(mapping).sort()) {
    if (usePackage) continue;
    const filename = reference ? path.join(reference, file) : path.join(root, mapping[file]);
    if (reference && !fs.existsSync(filename)) continue; // Core F4A.1 has no M1.2 modules.
    sources.push(fs.readFileSync(filename, 'utf8'));
  }
  if (!reference && !usePackage) for (const file of ['members/65_member_dev.gs','members/66_member_contracts.gs','members/69_member_authority.gs','members/70_member_mutation_plan.gs','members/73_member_command_coordinator.gs','members/81_member_command_catalog.gs','members/86_member_consumers.gs','members/71_member_admission.gs','members/72_member_claim.gs','members/74_member_identity.gs','members/75_member_endpoint.gs','members/82_member_legacy_fence.gs','members/87_member_operations.gs','members/88_member_legacy_targets.gs']) sources.push(fs.readFileSync(path.join(root, file), 'utf8'));
  // Apps Script hoists global declarations across files before global initializers.
  vm.runInContext(sources.join('\n;\n'), context, { filename: 'combined-gas.gs', timeout: 10000 });
  const run = (fn, argument) => {
    const code = typeof fn === 'function' ? '(' + fn.toString() + ')(' + JSON.stringify(argument ?? null) + ')' : fn;
    return vm.runInContext(code, context, { timeout: 10000 });
  };
  // Historical fixture helper clones only JSON source records, within the VM realm.
  vm.runInContext('const structuredClone = value => JSON.parse(JSON.stringify(value));', context);
  if (!reference || fs.existsSync(path.join(reference, '61_member_model.gs'))) load(path.join(__dirname, 'fixtures.cjs'));
  return { run, context, documents, calls, loseNextCommitResponse: () => { loseResponse = true; } };
}
module.exports = { createRuntime, root, legacyFiles };
