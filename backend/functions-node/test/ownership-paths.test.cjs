'use strict';
const {test}=require('node:test'),a=require('node:assert/strict');
const {createDomain}=require('../src/domain.cjs');
test('Apps Script Firestore adapter accepts only bounded pilot ownership paths',()=>{
 const d=createDomain(()=>{throw Error('NO_IO');});
 for(const p of ['commandExecutorRoutes/ACTIVITY_CREATE','commandExecutorRoutes/ACTIVITY_EDIT','commandExecutorRoutes/ACTIVITY_PUBLISH','commandExecutorRoutes/ACTIVITY_REALIZE','activityExecutorScopes/fn_cycle','commandExecutorRoutes/CF_DEV_EDIT_ACTIVITY_V2','pilotResources/CF_DEV_EDIT_ACTIVITY_V2','commandExecutorOwners/'+'a'.repeat(32)])a.equal(JSON.parse(d.run('(p)=>f4Firestore_().path(p)',p)),p);
 for(const p of ['commandExecutorRoutes/ACTIVITY_DELETE','activityExecutorScopes/a/extra/b','commandExecutorRoutes/real','pilotResources/real','commandExecutorOwners/invalid','commandExecutorOwners/'+'a'.repeat(32)+'/extra/child'])a.throws(()=>d.run('(p)=>f4Firestore_().path(p)',p),/DESTINATION_INVALID/);
});
test('Apps Script accepts exact real scope configuration paths, not new commands or arbitrary nesting',()=>{
 const d=createDomain(()=>{throw Error('NO_IO');});
 for(const p of ['activityRealExecutorIndex/current','activityRealExecutorScopes/REVIEW_1','commandExecutorRoutes/ACTIVITY_REAL_REVIEW_1_CREATE','commandExecutorRoutes/ACTIVITY_REAL_REVIEW_1_EDIT'])a.equal(JSON.parse(d.run('(p)=>f4Firestore_().path(p)',p)),p);
 for(const p of ['activityRealExecutorIndex/other','activityRealExecutorScopes/a/extra/b','commandExecutorRoutes/ACTIVITY_REAL_REVIEW_1_REALIZE','commandExecutorRoutes/ACTIVITY_REAL_REVIEW_1_DELETE'])a.throws(()=>d.run('(p)=>f4Firestore_().path(p)',p),/DESTINATION_INVALID/);
});
