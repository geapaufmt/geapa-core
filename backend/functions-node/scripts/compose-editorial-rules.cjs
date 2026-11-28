'use strict';
// Offline composition only. The authenticated snapshot stays outside Git.
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const assert = require('node:assert/strict');
const hash = value => crypto.createHash('sha256').update(value).digest('hex');
const EXPECTED = 'f7333a2f341c4e5116ebe216cbb57b871eede89c936c7929eaa8188c037ac309';
function compose(snapshot) {
  assert.equal(snapshot.release, 'projects/geapa-convergencia-dev/releases/cloud.firestore');
  assert.equal(snapshot.files.length, 1);
  const source = snapshot.files[0].content;
  assert.equal(hash(source), EXPECTED, 'Rules source changed: review the new full ruleset before composing');
  const start = source.indexOf('    match /publicPages/{id} {');
  const end = source.indexOf('    // B02:', start);
  assert.ok(start > 0 && end > start && end < source.indexOf('    match /coreDevValidation/'));
  const before = source.slice(start, end);
  // Retain existing DTO key allowlists, ID equality and size/type checks.
  const after = before
    .replace("allow get: if resource.data.id == id && published(resource.data,'geapa-public-v2.1',", "allow get: if database == '(default)'\n        && (id in ['home','about','eixos-tematicos'] || id.matches('EIXOS_[0-9A-Za-z_-]{1,32}_EIXO_[A-Z0-9_]{1,27}'))\n        && (resource == null || (resource.data.id == id && published(resource.data,'geapa-public-v2.1',")
    .replace('resource.data.sections.size() <= 30;', 'resource.data.sections.size() <= 30));')
    .replaceAll('allow get: if resource.data.id == id', "allow get: if database == '(default)' && resource.data.id == id")
    .replaceAll('boundedQuery(100)', "database == '(default)' && boundedQuery(101)");
  assert.notEqual(after, before);
  const result = source.slice(0, start) + after + source.slice(end);
  assert.equal(result.replace(after, before), source, 'Unexpected changes outside the three root public matches');
  return {rules: result, manifest: {schemaVersion: 1, remoteActionsAuthorized: false,
    release: snapshot.release, ruleset: snapshot.ruleset, readAt: snapshot.readAt,
    sourceSha256: hash(source), composedSha256: hash(result),
    changedScope: ['publicPages/{id}', 'publicTimeline/{id}', 'publicDocuments/{id}'],
    unchangedPrefixSha256: hash(source.slice(0, start)), unchangedSuffixSha256: hash(source.slice(end))}};
}
if (require.main === module) {
  const [input, output] = process.argv.slice(2);
  assert.ok(input && output, 'Usage: node scripts/compose-editorial-rules.cjs <source.json> <output.rules>');
  const result = compose(JSON.parse(fs.readFileSync(input, 'utf8')));
  fs.mkdirSync(path.dirname(output), {recursive: true});
  fs.writeFileSync(output, result.rules);
  fs.writeFileSync(output + '.manifest.json', JSON.stringify(result.manifest, null, 2) + '\n');
  console.log(JSON.stringify(result.manifest));
}
module.exports = {compose};
