const {test}=require('node:test');
const assert=require('node:assert/strict');
const {fixture,payload,op}=require('./import-recognition.test.cjs');
const {harness}=require('./model-frequency.test.cjs');
const ref=(kind,id)=>({kind,id});
function scenario(){
 const docs=fixture(),core=harness(docs);
 assert.equal(core.run('ACTIVITY_IMPORT_RECOGNIZE',payload,{op,allowed:['ACTIVITY_MANAGE']}).response.businessState,'COMMIT_APPLIED');
 const activity=docs['canonicalActivities/SYN_IMPORT_ACTIVITY'];
 activity.personRef=ref('person','SYN_OTHER_SPEAKER');
 activity.ownerPersonRef='SYN_OTHER_SPEAKER';
 activity.presentationRef=ref('presentation','SYN_PRESENTATION');
 docs['canonicalPresentations/SYN_PRESENTATION']={personRef:ref('person','SYN_OTHER_SPEAKER')};
 const fact=docs['activityAttendance/SYN_F0'];
 fact.personRef=ref('person','SYN_BOARD');fact.ownerPersonRef='SYN_BOARD';
 delete fact.pendingJustificationRef;
 docs['activityAttendanceCollections/'+activity.id].justificationEndsAt=Date.parse('2026-11-04T18:00Z');
 const body={activityRef:ref('activity',activity.id),attendanceRef:ref('attendance',fact.id),reason:'Synthetic request',lateAcknowledged:false};
 const options={kind:'PORTAL_AUTHENTICATED_MEMBER',allowed:['ACTIVITY_JUSTIFICATION_SELF'],dependencyRefs:[ref('attendanceCollection',activity.id)]};
 return {docs,core,body,options,run:()=>core.run('ATTENDANCE_JUSTIFICATION_REQUEST',body,options)};
}
test('SELF submits own absence for an activity led by someone else and retains its ownership metadata',()=>{
 const s=scenario(),activity=JSON.stringify(s.docs['canonicalActivities/SYN_IMPORT_ACTIVITY']),result=s.run();
 assert.equal(result.response.businessState,'COMMIT_APPLIED',JSON.stringify(result.response));
 const j=s.docs['activityJustifications/'+result.response.result.justificationRef.id];
 assert.equal(j.personRef.id,'SYN_BOARD');assert.equal(j.state,'PENDING');
 assert.equal(s.docs['activityAttendance/SYN_F0'].status,'ABSENT');
 assert.equal(JSON.stringify(s.docs['canonicalActivities/SYN_IMPORT_ACTIVITY']),activity);
 const count=Object.keys(s.docs).filter(p=>p.startsWith('activityJustifications/')).length;
 const repeat=s.run();assert.equal(repeat.response.businessState,'REJECTED');
 assert.equal(Object.keys(s.docs).filter(p=>p.startsWith('activityJustifications/')).length,count);
});
test('SELF still rejects another person attendance even when caller is the activity speaker',()=>{
 const s=scenario();s.docs['canonicalActivities/SYN_IMPORT_ACTIVITY'].personRef=ref('person','SYN_BOARD');
 s.docs['activityAttendance/SYN_F0'].personRef=ref('person','OTHER');
 assert.throws(s.run,e=>e.reasonCode==='OWNER_MISMATCH');
});
test('SELF still rejects revoked capability and stale attendance revision',()=>{
 const s=scenario();s.options.allowed=[];assert.throws(s.run,/AUTHORIZATION_DENIED/);
 s.options.allowed=['ACTIVITY_JUSTIFICATION_SELF'];s.options.stale=true;assert.throws(s.run,/STALE_REVISION/);
});
test('SELF cannot switch activity or submit outside the effective window',()=>{
 const s=scenario();s.docs['activityAttendance/SYN_F0'].activityRef=ref('activity','ANOTHER');
 assert.equal(s.run().response.businessState,'REJECTED');
 const t=scenario();t.docs['activityAttendanceCollections/SYN_IMPORT_ACTIVITY'].justificationEndsAt=Date.parse('2026-11-02T18:00Z');
 assert.equal(t.run().response.error.code,'DEADLINE_VIOLATION');
});
