'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GEAPA_PILOT_MODE:'SHARED_DEV',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),a=require('node:assert/strict'),{setup}=require('./people-fixture.cjs'),{Timestamp}=require('firebase-admin/firestore'),E=require('../src/environment.cjs');
const {fixture}=require('./import-rectification-compat-fixture.cjs');

test('documentary import rectification changes only active and revision and retains replay receipt',()=>{
 const h=fixture(),out=h.resolve();a.equal(out.replay,false);a.equal(out.mutation.changes.length,1);a.deepEqual(out.mutation.changes[0].data,{...h.person,active:true,revision:2});a.equal(out.receipt.state,'COMPLETED');a.equal(out.receipt.after.state,'MEMBRO_EFETIVO');a.equal(out.receipt.after.blocked,false);h.docs[out.receiptPath]=out.receipt;const replay=h.resolve();a.equal(replay.replay,true);a.deepEqual(replay.receipt,out.receipt);
});
test('import rectification rejects missing/revoked approval, stale source, unrelated operation, missing evidence and forged replay',()=>{
 for(const mutate of [h=>delete h.docs['memberImportRectificationApprovals/'+h.input.approvalId],h=>h.batch.active=false,h=>h.person.revision++,h=>h.input.operationId='b'.repeat(32),h=>h.docs['memberEvidenceAttestations/SYN_IMPORT_EVIDENCE'].active=false,h=>h.docs['memberAudit/'+h.input.operationId]={schemaVersion:'member-import-rectification-receipt-v1',state:'COMPLETED',digest:'b'.repeat(64)}]){const h=fixture();mutate(h);a.throws(h.resolve);}
});
test('normal profile scope cannot use imported inactivity exception without protected approval',()=>{
 const h=fixture();a.equal(JSON.parse(h.domain.run('(x)=>f4MemberProfileTargetEligible_(nodeRead,x.id,x.now)',{id:h.input.personId,now:h.now})),false);delete h.docs['memberImportRectificationApprovals/'+h.input.approvalId];a.throws(h.resolve);
});
