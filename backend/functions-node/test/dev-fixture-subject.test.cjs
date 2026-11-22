const {test}=require('node:test'),a=require('node:assert/strict'),{fixtures,repairFixtureSubject}=require('../scripts/administer-dev-access.cjs');
function setup(){const ownerUid='SYN_OWNER',linkedUid='SYN_LINKED',f=fixtures(ownerUid,1000),personId=f.persons[0],docs={...f.docs},writes=[];
 delete docs['memberAuthzSubjects/'+personId];
 docs['coreDevAdministrativeFixtures/'+ownerUid]={uid:ownerUid,origin:'DEV_ADMIN_SYNTHETIC_V1',personIds:f.persons};
 docs['memberAuthBindings/'+linkedUid]={uid:linkedUid,idPessoa:personId,active:true,bindingVersion:1};
 const db={projectId:'geapa-convergencia-dev',doc:p=>p,runTransaction:fn=>fn({get:async p=>({data:()=>docs[p]}),create:(p,d)=>{a(!docs[p]);docs[p]=d;writes.push(p);}})};
 return {docs,writes,personId,input:{db,ownerUid,personId,linkedUid,expectedPersonRevision:1,approvalRef:'SYN_APPROVED',operator:'SYN_OPERATOR',clock:()=>1000}};
}
test('new administrative fixtures include a canonical authorization subject',()=>{const f=fixtures('SYN_OWNER',1000);for(const id of f.persons)a.deepEqual(f.docs['memberAuthzSubjects/'+id],{schemaVersion:'member-authz-subject-v13',idPessoa:id,revision:1,authzEpoch:1,active:true,approvalRef:'DEV_ADMIN_SYNTHETIC_V1',updatedAt:new Date(1000)});});
test('registered linked fixture repair is dry-run by default and creates subject with audit exactly once',async()=>{const h=setup();await repairFixtureSubject(h.input);a.equal(h.writes.length,0);const result=await repairFixtureSubject({...h.input,apply:true});a.equal(result.capabilitiesChanged,false);a.equal(h.writes.length,2);a.equal(h.docs['memberAuthzSubjects/'+h.personId].active,true);await a.rejects(repairFixtureSubject({...h.input,apply:true}),/SUBJECT_EXISTS/);});
test('fixture repair cannot reactivate a subject, target a real person or cross project/binding/revision boundaries',async()=>{
 for(const mutate of [h=>h.input.db.projectId='portal-geapa',h=>h.docs['memberAuthzSubjects/'+h.personId]={active:false},h=>h.docs['memberIdentities/'+h.personId].origin='REAL',h=>h.docs['memberAuthBindings/SYN_LINKED'].idPessoa='OTHER',h=>h.input.expectedPersonRevision=2,h=>h.docs['coreDevAdministrativeFixtures/SYN_OWNER'].personIds=[]]){const h=setup();mutate(h);await a.rejects(repairFixtureSubject({...h.input,apply:true}));a.equal(h.writes.length,0);}
});
