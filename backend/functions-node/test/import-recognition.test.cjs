const {test}=require('node:test'),a=require('node:assert/strict'),H=require('./fixtures/import-reconstruction.cjs'),{harness}=require('./model-frequency.test.cjs');
const now=Date.parse('2026-11-03T18:00Z'),op='78'.repeat(16),ref=H.ref;
function fixture(n=3,actor='SYN_BOARD'){
 const d=H.fixture(n),v=d['canonicalActivities/SYN_IMPORT_ACTIVITY'];
 v.administrativeReview={schemaVersion:'activity-administrative-review-v1',state:'CONFIRMED',cycleRef:ref('cycle','GEAPA_2026'),revision:2,initializedAt:now-1000,initializationRef:'SYN_INIT',confirmedAt:now-1000,confirmedBy:ref('person','SYN_BOARD'),confirmationOperationId:'77'.repeat(16)};
 d['presentationCapacityHeads/registry']={id:'registry',revision:1,kind:'DOMAIN_FENCE',membersEpoch:1,activityEpoch:1,cycleRegistryEpoch:1};
 const intent=H.intentFor(d);intent.operationId=op;
 const planned=H.setup(d).plan(intent,now);
 d['presentationPolicyReferences/SYN_IMPORT_APPROVAL']={id:'SYN_IMPORT_APPROVAL',revision:1,usage:'ACTIVITY_IMPORT_RECOGNITION',approved:true,active:true,actorPersonIds:[actor],issuedAt:now-60000,expiresAt:now+3600000,intent,intentHash:H.hash(intent),factsPreservedHash:planned.factsPreservedHash,historicalCriteria:{schemaVersion:'import-historical-criteria-v1',historicalApprovalClaimed:false,civilStart:'2026-01-01',civilEnd:'2026-12-31',basis:'Synthetic documentary confirmation only',policyId:'SYN_POLICY',policyRevision:1,policyCriteriaHash:H.hash(H.setup(d).ctx.f4ActivityConfigNormativeFields_('attendancePolicy',d['activityPolicies/SYN_POLICY'])),sources:[{ref:ref('evidence','EA'),hash:H.hash(d['migrationEvidence/EA'])}]}};
 return d;
}
const payload={activityRef:ref('activity','SYN_IMPORT_ACTIVITY'),recognitionRef:ref('policy','SYN_IMPORT_APPROVAL'),reason:'Synthetic sealed recognition'};
const run=(h,opt={})=>h.run('ACTIVITY_IMPORT_RECOGNIZE',payload,{op,allowed:['ACTIVITY_MANAGE'],...opt});
test('canonical runtime recognizes imports, archives exact before images and advances fences',()=>{
 const docs=fixture(25),h=harness(docs),before=H.plain(docs),result=run(h);a.equal(result.response.businessState,'COMMIT_APPLIED',JSON.stringify(result.response));
 a.equal(docs['canonicalActivities/SYN_IMPORT_ACTIVITY'].recordMode,'NEW_V2');a.deepEqual(docs['canonicalActivities/SYN_IMPORT_ACTIVITY'].administrativeReview,before['canonicalActivities/SYN_IMPORT_ACTIVITY'].administrativeReview);
 for(let i=0;i<25;i++){a.equal(docs['activityAttendance/SYN_F'+i].status,before['activityAttendance/SYN_F'+i].status);a.deepEqual(docs['activityAttendance/SYN_F'+i+'/revisions/1'].snapshot,before['activityAttendance/SYN_F'+i]);}
 a.equal(docs['activityJustifications/SYN_J'].state,'PENDING');a.equal(docs['activityJustifications/SYN_J'].basisAttendanceRevision,2);a.equal(docs['activityAttendanceRegistry/current'].epoch,2);a.equal(docs['presentationCapacityHeads/registry'].activityEpoch,2);a(docs['activityAgendaViews/SYN_IMPORT_ACTIVITY']);a.equal(docs['presentationCycles/GEAPA_2026'].status,'OPEN');
});
for(const [name,mutate]of Object.entries({revoked:d=>d['presentationPolicyReferences/SYN_IMPORT_APPROVAL'].active=false,expired:d=>d['presentationPolicyReferences/SYN_IMPORT_APPROVAL'].expiresAt=now,otherActor:d=>d['presentationPolicyReferences/SYN_IMPORT_APPROVAL'].actorPersonIds=['OTHER'],sourceDrift:d=>d['migrationEvidence/E0'].literal.CODIGO_PRESENCA='P',policyDrift:d=>d['activityPolicies/SYN_POLICY'].revision++,manifestDrift:d=>d['presentationPolicyReferences/SYN_IMPORT_APPROVAL'].intent.typeId='OTHER',missingHistoricalBasis:d=>delete d['presentationPolicyReferences/SYN_IMPORT_APPROVAL'].historicalCriteria,archiveConflict:d=>d['activityAttendance/SYN_F0/revisions/1']={snapshot:{wrong:true}}}))test('recognition refuses '+name+' without domain effects',()=>{const d=fixture();mutate(d);const before=H.hash(d),r=run(harness(d));a.equal(r.response.businessState,'REJECTED');a.equal(r.writes.length,0);a.equal(H.hash(d),before);});
test('ordinary authorization remains required independently of sealed import approval',()=>{a.throws(()=>run(harness(fixture()),{allowed:[]}),/AUTHORIZATION_DENIED/);});
test('stale activity revision is rejected before reconstruction',()=>{a.throws(()=>run(harness(fixture()),{stale:true}),/EXPECTED_VECTOR_CHANGED/);});
module.exports={fixture,payload,op};
