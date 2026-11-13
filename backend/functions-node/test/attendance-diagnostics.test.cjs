'use strict';
const {test}=require('node:test'),a=require('node:assert/strict'),{handler}=require('../src/http.cjs'),E=require('../src/environment.cjs');
test('attendance failures log only allowlisted reason and preparation stage',async()=>{
 Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GCLOUD_PROJECT:E.DEMO,FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
 for(const [reason,stage]of [['PREPARATION_DEPENDENCIES','COMMAND_PREPARATION'],['ROSTER_DISCOVERY_LIMIT','ROSTER_DISCOVERY'],['ATTENDANCE_ROSTER_LIMIT','ROSTER_DISCOVERY'],['READ_BUDGET','COMMAND_PREPARATION'],['private-person@example.invalid','private-token']]){
  const logs=[],res={set(){return this},status(){return this},json(v){this.body=v;return this}};
  await handler({runtime:async()=>({}),execute:async()=>{throw Object.assign(Error('private-message'),{code:'SCOPE_TOO_LARGE',reasonCode:reason,preparationStage:stage});},log:s=>logs.push(JSON.parse(s)),actions:['flowPortalPrepare']})({method:'POST',headers:{authorization:'Bearer private-credential','content-type':'application/json'},body:{action:'flowPortalPrepare',request:{}}},res);
  a.equal(logs[0].internalCode,reason.startsWith('private')?null:reason);a.equal(logs[0].preparationStage,stage.startsWith('private')?null:stage);a.deepEqual(res.body,{error:{code:'SCOPE_TOO_LARGE'}});a(!JSON.stringify(logs).includes('private'));
 }
});
