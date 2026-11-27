'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { Firestore } = require('@google-cloud/firestore');
const { randomUUID } = require('node:crypto');

test('editorial public queries and client isolation under Firestore rules', {
  skip: process.env.GEAPA_EDITORIAL_RULES_TEST !== '1'
}, async () => {
  const project = 'demo-geapa-functions-local';
  assert.equal(process.env.FIRESTORE_EMULATOR_HOST, '127.0.0.1:8187');
  assert.equal(process.env.FIREBASE_AUTH_EMULATOR_HOST, '127.0.0.1:9197');
  const db = new Firestore({projectId: project});
  const id = 'rule_' + randomUUID().replaceAll('-', '');
  const root = 'http://127.0.0.1:8187/v1/projects/' + project + '/databases/(default)/documents';
  const timeline = {schemaVersion:'geapa-public-v2.1', published:true, sanitized:true};
  const norm = {schemaVersion:'geapa-public-document-v1', published:true, sanitized:true, status:'VIGENTE'};
  const records = {
    ['publicTimeline/' + id]: timeline,
    ['publicTimeline/' + id + '_private']: {...timeline, published:false},
    ['publicDocuments/' + id]: norm,
    ['publicDocuments/' + id + '_old']: {...norm, status:'REVOGADA'},
    ['coreDevValidation/CF_DEV_EDIT_PILOT_V1/publicTimeline/' + id]: timeline
  };
  const signup = await fetch('http://127.0.0.1:9197/identitytoolkit.googleapis.com/v1/accounts:signUp?key=demo', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({returnSecureToken:true})
  });
  assert.equal(signup.status,200);
  const {idToken} = await signup.json();
  assert.ok(idToken);
  try {
    const batch = db.batch();
    for (const [path,data] of Object.entries(records)) batch.set(db.doc(path), data);
    await batch.commit();
    for (const token of [null,idToken]) {
      const headers = {'Content-Type':'application/json'};
      if (token) headers.Authorization = 'Bearer ' + token;
      for (const [path,status] of [
        ['publicTimeline/'+id,200], ['publicDocuments/'+id,200], ['publicTimeline/'+id+'_private',403],
        ['publicDocuments/'+id+'_old',403], ['coreDevValidation/CF_DEV_EDIT_PILOT_V1/publicTimeline/'+id,403],
        ['publicPages/unknown',403], ...['canonicalEditorial','canonicalEditorialHistory','editorialReceipts',
          'editorialAudit','editorialRate','commandExecutorOwners','editorialControl'].map(c=>[c+'/'+id,403])
      ]) assert.equal((await fetch(root+'/'+path,{headers})).status,status,path);
      for (const path of ['publicPages/home','publicTimeline/'+id,'publicDocuments/'+id,'canonicalEditorial/home','editorialControl/current']) {
        const response = await fetch(root+'/'+path,{method:'PATCH',headers,body:JSON.stringify({fields:{}})});
        assert.equal(response.status,403,path);
      }
      async function query(collection, fields, limit) {
        const filters = Object.entries(fields).map(([fieldPath,value])=>({fieldFilter:{
          field:{fieldPath}, op:'EQUAL', value:typeof value==='boolean'?{booleanValue:value}:{stringValue:value}
        }}));
        const structuredQuery = {from:[{collectionId:collection}],limit};
        if (filters.length) structuredQuery.where = {compositeFilter:{op:'AND',filters}};
        return fetch(root+':runQuery',{method:'POST',headers,body:JSON.stringify({structuredQuery})});
      }
      for (const [collection,fields] of [['publicTimeline',timeline],['publicDocuments',norm]]) {
        const response = await query(collection,fields,101);
        assert.equal(response.status,200);
        const rows = await response.json();
        assert.ok(rows.some(row=>row.document?.name.endsWith('/'+id)));
        assert.ok(!rows.some(row=>row.document?.name.endsWith('/'+id+'_old')));
        assert.equal((await query(collection,{},101)).status,403);
        assert.equal((await query(collection,fields,102)).status,403);
      }
      const {status,...withoutStatus} = norm;
      assert.equal((await query('publicDocuments',withoutStatus,101)).status,403);
    }
  } finally {
    const cleanup = db.batch();
    for (const path of Object.keys(records)) cleanup.delete(db.doc(path));
    await cleanup.commit();
    await db.terminate();
    await fetch('http://127.0.0.1:9197/identitytoolkit.googleapis.com/v1/accounts:delete?key=demo',{
      method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({idToken})
    });
  }
});
