'use strict';
const {test} = require('node:test');
const assert = require('node:assert/strict');
const {Firestore, Timestamp} = require('@google-cloud/firestore');

test('full rules composition preserves owner, member directory, projections and private isolation', {
  skip: process.env.GEAPA_EDITORIAL_COMPOSITION_TEST !== '1'
}, async () => {
  assert.equal(process.env.GCLOUD_PROJECT, 'demo-geapa-functions-local');
  assert.equal(process.env.FIRESTORE_EMULATOR_HOST, '127.0.0.1:8187');
  assert.equal(process.env.FIREBASE_AUTH_EMULATOR_HOST, '127.0.0.1:9197');
  const project = 'demo-geapa-functions-local';
  const db = new Firestore({projectId: project});
  const root = `http://127.0.0.1:8187/v1/projects/${project}/databases/(default)/documents/`;
  const authRoot = 'http://127.0.0.1:9197/identitytoolkit.googleapis.com/v1/accounts:';
  const users = [];
  const records = {};
  try {
    for (let i=0;i<2;i++) {
      const response = await fetch(authRoot+'signUp?key=demo', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({returnSecureToken:true})});
      assert.equal(response.status,200); users.push(await response.json());
    }
    const prefix = 'coreDevValidation/ACT_FUNCTIONAL_20261008/';
    const now = Date.now(), uid = users[0].localId;
    const status = {schemaVersion:'portal-access-status-v1',status:'ACTIVE',updatedAt:Timestamp.fromMillis(now-1000),validUntil:Timestamp.fromMillis(now+3600000)};
    const publicShape = {schemaVersion:'geapa-public-v2.1',published:true,sanitized:true};
    Object.assign(records, {
      ['portalAccessStatus/'+uid]:status,
      [prefix+'portalAccessStatus/'+uid]:status,
      [prefix+'publicMemberDirectoryIndex/current']:{schemaVersion:'public-member-directory-index-v1',keys:[]},
      [prefix+'publicPresentations/synthetic']:{schemaVersion:1,published:true,sanitized:true},
      [prefix+'publicPages/home']:{...publicShape,id:'home',title:'Synthetic preserved namespace',sections:[{type:'TEXT'}]}
    });
    const privateCollections = ['canonicalEditorial','canonicalEditorialHistory','editorialControl','editorialReceipts','editorialAudit','editorialRate','commandExecutorOwners','memberAuthBindings','memberFunctions','memberCurrentStates','memberAuthzPolicies','institutionalDocuments','memberAdmissionApprovals','portalReaderControl'];
    for(const collection of privateCollections) records[collection+'/synthetic']={ownerUid:uid,published:true,sanitized:true};
    const batch=db.batch(); for(const [p,v] of Object.entries(records)) batch.set(db.doc(p),v); await batch.commit();
    for(const user of [null,...users]) {
      const headers={'Content-Type':'application/json',...(user?{Authorization:'Bearer '+user.idToken}:{})};
      for(const p of ['portalAccessStatus/'+uid,prefix+'portalAccessStatus/'+uid]) {
        assert.equal((await fetch(root+p,{headers})).status,user?.localId===uid?200:403,p);
      }
      for(const p of [prefix+'publicMemberDirectoryIndex/current',prefix+'publicPresentations/synthetic',prefix+'publicPages/home']) {
        assert.equal((await fetch(root+p,{headers})).status,200,p);
      }
      for(const collection of privateCollections) {
        assert.equal((await fetch(root+collection+'/synthetic',{headers})).status,403,collection);
      }
      for(const p of Object.keys(records)) {
        for(const method of ['PATCH','DELETE']) {
          assert.equal((await fetch(root+p,{method,headers,...(method==='PATCH'?{body:'{"fields":{}}'}:{})})).status,403,method+' '+p);
        }
      }
      for(const collection of [...privateCollections,'publicPages','portalAccessStatus']) {
        const query={structuredQuery:{from:[{collectionId:collection}],limit:1}};
        assert.equal((await fetch(root.slice(0,-1)+':runQuery',{method:'POST',headers,body:JSON.stringify(query)})).status,403,'list '+collection);
      }
    }
  } finally {
    const batch=db.batch();for(const p of Object.keys(records))batch.delete(db.doc(p));await batch.commit();await db.terminate();
    for(const u of users) await fetch(authRoot+'delete?key=demo',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({idToken:u.idToken})});
  }
});
