'use strict';
require('../src/local-network.cjs');
const {test}=require('node:test'),a=require('node:assert/strict');
const E=require('../src/environment.cjs');
for(const k of Object.keys(process.env))if(/EMULATOR|^GOOGLE_|^GCLOUD_|^CLOUDSDK_|^GCE_METADATA|^FIREBASE_/i.test(k))delete process.env[k];
Object.assign(process.env,{GEAPA_RUNTIME:'DEV',GCLOUD_PROJECT:E.PROJECT,K_SERVICE:'candidate',K_REVISION:'local-test'});
const {GoogleAuth,Compute}=require('google-auth-library');
const original=GoogleAuth.prototype.getClient;
globalThis.fetch=async url=>({ok:true,headers:new Map([['metadata-flavor','Google']]),text:async()=>url.endsWith('project-id')?E.PROJECT:E.SERVICE_ACCOUNT});
test('real runtime rejects user ADC, then constructs official SDK with pinned Compute ADC',async()=>{
 const {runtime}=require('../src/runtime.cjs');
 try{
  GoogleAuth.prototype.getClient=async()=>({getAccessToken:async()=>({token:'not-a-cloud-token'})});
  await a.rejects(runtime(),/DESTINATION_INVALID/);
  const client=new Compute();
  GoogleAuth.prototype.getClient=async function(){this.cachedCredential=client;return client;};
  const deps=await runtime();a.equal(deps.db.projectId,E.PROJECT);a.equal(deps.db._settings.auth.cachedCredential,client);
  a.equal(deps.auth.app.options.projectId,E.PROJECT);a.equal(typeof deps.auth.app.options.credential.getAccessToken,'function');
  await deps.db.terminate();
 }finally{GoogleAuth.prototype.getClient=original;}
});
