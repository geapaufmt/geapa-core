'use strict';

// Replay a recorded Gen2 source package from a clean release worktree.
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { Readable } = require('node:stream');
const ZipStream = require('zip-stream');

const digest = data => crypto.createHash('sha256').update(data).digest('hex');
const json = data => Buffer.from(JSON.stringify(data, null, 2) + '\n');

async function replay(revision, checkout, output) {
  assert(/^geapacommand[a-z]*-\d{5}-[a-z]{3}$/.test(revision), 'Invalid revision');
  const record = require('../releases/20261120/' + revision + '.json');
  const git = (...args) => execFileSync('git', ['-C', checkout, ...args], {encoding:'utf8'}).trim();
  assert.equal(git('rev-parse', 'HEAD'), record.treeCommit);
  assert.equal(git('status', '--porcelain', '--untracked-files=normal'), '');
  const runtime = path.join(checkout, 'backend', 'functions-node');
  execFileSync(process.execPath, [path.join(runtime, 'scripts/build-domain.cjs')], {stdio:'inherit'});
  const sources = JSON.parse(fs.readFileSync(path.join(runtime, 'generated/source-manifest.json')));
  assert.equal(sources.sourceDigest, record.release.sourceDigest);
  assert.equal(sources.bundleSha256, record.bundleSha256);
  sources.sourceCommit = record.release.sourceCommit;
  const entries = new Map();
  for (const row of record.release.files) {
    const name = row.path;
    assert(/^[\w./-]+$/.test(name) && !name.startsWith('/') && !name.includes('..'));
    let bytes;
    switch (name) {
      case 'generated/source-manifest.json': bytes = json(sources); break;
      case 'package.json': bytes = json(record.package); break;
      case 'package-lock.json': bytes = fs.readFileSync(path.join(runtime,'package-lock.runtime.json')); break;
      case '.gcloudignore': bytes = Buffer.from('node_modules/\n.git/\n'); break;
      default: bytes = fs.readFileSync(path.join(runtime, name));
    }
    assert.equal(bytes.length, row.bytes, name);
    assert.equal(digest(bytes), row.sha256, name);
    assert(!entries.has(name));
    entries.set(name, bytes);
  }
  assert.equal(entries.size, 53);
  assert.equal(digest(JSON.stringify(record.release.files)), record.release.inventoryDigest);
  entries.set('release-manifest.json', json(record.release));
  assert.deepEqual([...entries.keys()].sort(), [...record.archiveOrder].sort());
  const zip = new ZipStream({zlib:{level:9}});
  const chunks = [];
  zip.on('data', chunk => chunks.push(chunk));
  const finished = require('node:stream/promises').finished(zip);
  for (const name of record.archiveOrder) {
    await new Promise((resolve,reject) => zip.entry(Readable.from([entries.get(name)]),
      {name,date:new Date(record.archiveDate),mode:0o100666}, err => err ? reject(err) : resolve()));
  }
  zip.finalize();
  await finished;
  const bytes = Buffer.concat(chunks);
  assert.equal(digest(bytes), record.zipSha256, 'Historical archive digest');
  if (output) fs.writeFileSync(output, bytes, {flag:'wx'});
  console.log(JSON.stringify({revision,files:53,zipSha256:digest(bytes)}));
}
replay(...process.argv.slice(2)).catch(err => { console.error(err); process.exitCode=1; });
