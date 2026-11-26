const assert = require('node:assert/strict');
const {createDomain} = require('../backend/functions-node/src/domain.cjs');
const {docs} = require('../backend/functions-node/test/fixtures/core-edit.json');

function isolatedDomain() {
  return createDomain(raw => {
    const request = JSON.parse(raw);
    assert.equal(request.kind, 'get');
    assert.equal(request.path, 'memberAuthorityControl/MEMBERS');
    return JSON.stringify({mapValue: docs[request.path]});
  });
}

// Exercise compatibility reads only with the real fence and a synthetic local control.
function enableCompatibility(domain) {
  domain.run(`() => {
    const control = {...nodeTx.get('memberAuthorityControl/MEMBERS').data(), mode:'LEGACY', activeWriterVersion:'NONE'};
    f4MemberLegacyFence_().installLocal({
      config:{environment:'SYNTHETIC_LOCAL',projectId:'demo-geapa-m04-local',database:'(default)',host:'127.0.0.1',port:8184},
      readControl:() => control,
      classify:() => { throw Error('UNEXPECTED_LEGACY_EFFECT'); }
    });
  }`);
}
module.exports = {isolatedDomain, enableCompatibility};
