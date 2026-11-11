const {test}=require('node:test'),assert=require('node:assert/strict');
const {setup}=require('./activity-profile-fixture.cjs');
function scenario(){
 const h=setup('DIRETOR_ADMINISTRATIVO',docs=>{
  const p=docs['memberAuthzPolicies/CF_DEV_ACT_POLICY'];
  Object.assign(p,{proofKind:'SELF',actorSelector:{kind:'SELF'},capabilities:['ACTIVITY_JUSTIFICATION_SELF'],fieldMasks:{ACTIVITY_JUSTIFICATION_SELF:[]},targets:{...p.targets,ownPerson:true,allowedPersons:[],allowedTargetScopes:['SELF'],allowedResources:[],resourceScope:{kind:'OWN_ATTENDANCE'}}});
  docs['canonicalActivities/SYN_A']={id:'SYN_A',revision:1,recordMode:'NEW_V2',personRef:{kind:'person',id:'OTHER_SPEAKER'}};
  docs['activityAttendance/SYN_F']={id:'SYN_F',revision:1,recordMode:'NEW_V2',activityRef:{kind:'activity',id:'SYN_A'},personRef:{kind:'person',id:'PES-000001'},ownerPersonRef:'PES-000001'};
 });
 const factId=JSON.parse(h.domain.run(()=> 'a_'+f4ActivitiesModel_().hash({activity:'SYN_A',person:'PES-000001'}).slice(0,32)));
 h.docs['activityAttendance/'+factId]=h.docs['activityAttendance/SYN_F'];
 h.docs['activityAttendance/'+factId].id=factId;
 const read=()=>JSON.parse(h.domain.run(now=>f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:{uid:'CF_ADMIN',expiresAt:now+60000},settings:{commands:['ATTENDANCE_JUSTIFICATION_REQUEST'],policyIds:['CF_DEV_ACT_POLICY']},clock:()=>now}).commandAccess('ATTENDANCE_JUSTIFICATION_REQUEST',[{kind:'attendance',id:'a_'+f4ActivitiesModel_().hash({activity:'SYN_A',person:'PES-000001'}).slice(0,32)},{kind:'activity',id:'SYN_A'}],'SELF'),Date.now()));
 return {...h,read};
}
test('canonical SELF action probe accepts own attendance with another activity speaker',()=>assert.equal(scenario().read(),true));
test('canonical action probe denies inactive, expired or insufficient SELF policy',()=>{
 for(const change of [p=>p.active=false,p=>p.expiresAt=require('firebase-admin/firestore').Timestamp.fromMillis(1),p=>{p.capabilities=['ACTIVITY_ATTENDANCE_READ_SELF'];p.fieldMasks={ACTIVITY_ATTENDANCE_READ_SELF:[]};}]){
  const h=scenario();change(h.docs['memberAuthzPolicies/CF_DEV_ACT_POLICY']);assert.equal(h.read(),false);
 }
});
test('canonical action probe denies foreign attendance',()=>{const h=scenario();h.docs['activityAttendance/SYN_F'].personRef.id='OTHER';h.docs['activityAttendance/SYN_F'].ownerPersonRef='OTHER';assert.equal(h.read(),false);});
