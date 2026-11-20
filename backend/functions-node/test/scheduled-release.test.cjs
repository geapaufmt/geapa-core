'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

// Load each environment in a fresh process so the SDK's module cache cannot
// reuse a previous profile. Inspect declarations without invoking a job.
for (const [environment, name] of [['DEV', 'geapaCommandDevPilot'], ['PROD', 'geapaCommand']]) {
  test(environment + ' exports the isolated scheduler only when explicitly enabled', () => {
    for (const enabled of ['true', 'false']) {
      const output = execFileSync(process.execPath, ['-e', [
        "const functions = require('./src/index.cjs');",
        "const endpoints = Object.entries(functions).map(([name, fn]) => [name, fn.__endpoint]);",
        "process.stdout.write(JSON.stringify(Object.fromEntries(endpoints)));"
      ].join('\n')], {
        cwd: path.resolve(__dirname, '..'),
        env: { ...process.env, GEAPA_TARGET_ENVIRONMENT: environment, GEAPA_BOARD_AUTOMATION_ENABLED: enabled },
        encoding: 'utf8'
      });
      const endpoints = JSON.parse(output);
      const worker = name + 'BoardConsequences';
      assert.deepEqual(Object.keys(endpoints).sort(), enabled === 'true' ? [name, worker] : [name]);
      assert(endpoints[name].httpsTrigger);
      assert.equal(endpoints[name].concurrency, 2);
      if (enabled === 'true') {
        assert.equal(endpoints[worker].scheduleTrigger.schedule, 'every 1 minutes');
        assert.equal(endpoints[worker].httpsTrigger, undefined);
        assert.equal(endpoints[worker].concurrency, 1);
        assert.equal(endpoints[worker].maxInstances, 1);
      }
    }
  });
}
