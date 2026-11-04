Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GEAPA_PILOT_MODE:'SHARED_DEV',GEAPA_ACTIVITY_COMMANDS_ENABLED:'true',GEAPA_ACT_FLOWS_ENABLED:'true',GEAPA_ACT_EXTENSIONS_ENABLED:'true',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),a=require('node:assert/strict'),{Firestore}=require('@google-cloud/firestore'),{execute}=require('../src/executor.cjs'),{PROJECT,DEMO,EXECUTOR}=require('../src/environment.cjs'),{setup}=require('./activity-profile-fixture.cjs'),{fixture,payload,op}=require('./import-recognition.test.cjs'),H=require('./fixtures/import-reconstruction.cjs');
test('import recognition through shared canonical executor, ROLE, receipts and recovery',{skip:process.env.GEAPA_IMPORT_EXECUTOR_TEST!=='1'},async t=>{
 const db=new Firestore({projectId:DEMO}),now=Date.now(),base=setup('PRESIDENTE').docs,docs=fixture(25,'PES-000001');let invalidAuth=false,invalidApp=false;
 delete docs['presentationCapacityHeads/registry'].activityEpoch;
 const approval=docs['presentationPolicyReferences/SYN_IMPORT_APPROVAL'];approval.issuedAt=now-1000;approval.expiresAt=now+3600000;
 const resources=Object.keys(docs).filter(p=>!p.startsWith('member')&&!p.endsWith('/registry'));
 const policy=base['memberAuthzPolicies/CF_DEV_ACT_POLICY'];policy.capabilities=['ACTIVITY_BOARD_READ','ACTIVITY_MANAGE'];policy.fieldMasks={ACTIVITY_BOARD_READ:[],ACTIVITY_MANAGE:[]};policy.targets.allowedTargetScopes=['ACTIVITIES'];policy.targets.allowedResources=resources;policy.targets.allowedPersons=['PES-000001'];delete policy.targets.resourceScope;
 const route='ACT_FLOW_ACTIVITY_IMPORT_RECOGNIZE',scope='SYN_IMPORT_SCOPE',record={...base,...docs,
  'activityExecutorControl/current':{schemaVersion:'activity-executor-control-v1',projectId:PROJECT,enabled:true,revision:1,actorUids:['CF_ADMIN'],policyIds:['CF_DEV_ACT_POLICY'],startsAt:now-1000,expiresAt:now+3600000,browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['SYN_APP']}},
  ['actFlowExecutorScopes/'+scope]:{id:scope,projectId:PROJECT,synthetic:true,revision:1,resourcePaths:resources},
  ['commandExecutorRoutes/'+route]:{schemaVersion:'geapa-act-flow-route-v1',projectId:PROJECT,routeId:route,commandType:'ACT.ACTIVITY_IMPORT_RECOGNIZE',scopeId:scope,syntheticOnly:true,executor:EXECUTOR,state:'ACTIVE',epoch:1,revision:1,actorUids:['CF_ADMIN'],operationIds:[],maxOperations:5,availabilityMode:'TIMEBOXED',startsAt:now-1000,expiresAt:now+3600000}};
 const auth={verifyIdToken:async(_,check)=>{a(check);if(invalidAuth)throw Error('revoked');return {uid:'CF_ADMIN',sub:'CF_ADMIN',aud:DEMO,iss:'https://securetoken.google.com/'+DEMO,iat:Math.floor(now/1000),auth_time:Math.floor(now/1000),exp:Math.floor(now/1000)+3600}}},appCheck={verifyToken:async()=>{if(invalidApp)throw Error('invalid');return {appId:'SYN_APP'}}};
 const call=(action,request,hooks={})=>execute({db,auth,appCheck,token:'emulator-only',appCheckToken:'emulator-only',origin:'https://geapa-convergencia-dev.web.app',body:{action,request},hooks});
 try{
  a.equal(db.projectId,'demo-geapa-functions-local');
  for(const s of (await db.collection('memberAdmissionCounters').get()).docs)await s.ref.delete();
  for(const p of Object.keys(docs).filter(p=>p.startsWith('activityAttendance/')||p.startsWith('canonicalActivities/')||p.startsWith('activityAttendanceCollections/')||p.startsWith('activityJustifications/')))await db.recursiveDelete(db.doc(p));
  const b=db.batch();for(const[p,v]of Object.entries(record))b.set(db.doc(p),v);await b.commit();
  const input={operationId:op,commandType:'ACT.ACTIVITY_IMPORT_RECOGNIZE',payload};
  invalidAuth=true;await a.rejects(call('flowPrepare',input));invalidAuth=false;invalidApp=true;await a.rejects(call('flowPrepare',input));invalidApp=false;
  const q=(await call('flowPrepare',input)).request;a(q);const before=H.hash((await db.doc('activityAttendance/SYN_F0').get()).data());let lost=false;
  try{await call('execute',q,{afterCommit(){lost=true;throw Error('SIMULATED_LOST_RESPONSE');}})}catch(e){if(!lost)throw e}
  const receipt=(await db.doc('commandReceipts/'+op).get()).data();a(receipt, 'RECEIPT_REQUIRED');a.equal(receipt.state,'COMPLETED',JSON.stringify(receipt));
  const facts=await db.collection('activityAttendance').where('activityRef.id','==','SYN_IMPORT_ACTIVITY').get();a.equal(facts.size,25);for(const s of facts.docs)a.equal(s.data().status,docs[s.ref.path].status);
  a.notEqual(H.hash((await db.doc('activityAttendance/SYN_F0').get()).data()),before);a.equal((await db.doc('activityJustifications/SYN_J').get()).data().basisAttendanceRevision,2);
  await call('execute',q);a.equal((await db.doc('activityAttendance/SYN_F0').get()).data().revision,2);
  a.equal((await db.doc('commandLedger/'+op).get()).data().state,'COMPLETED');a.equal((await db.doc('commandExecutorOwners/'+op).get()).data().executor,EXECUTOR);
  a.equal((await db.doc('activityAttendance/SYN_F0/revisions/1').get()).data().snapshot.recordMode,'HISTORICAL_PARTIAL');
  // Imported evidence has no operational revision; downstream decisions must
  // fence recognized facts without inventing revisions on source documents.
  policy.capabilities.push('ACTIVITY_ATTENDANCE_MANAGE');policy.fieldMasks.ACTIVITY_ATTENDANCE_MANAGE=[];
  await db.doc('memberAuthzPolicies/CF_DEV_ACT_POLICY').set(policy);
  const decisionRoute='ACT_FLOW_ATTENDANCE_JUSTIFICATION_DECIDE';await db.doc('commandExecutorRoutes/'+decisionRoute).set({...record['commandExecutorRoutes/'+route],routeId:decisionRoute,commandType:'ACT.ATTENDANCE_JUSTIFICATION_DECIDE',operationIds:[]});
  const decision=(await call('flowPrepare',{operationId:require('crypto').randomBytes(16).toString('hex'),commandType:'ACT.ATTENDANCE_JUSTIFICATION_DECIDE',payload:{justificationRef:H.ref('justification','SYN_J'),decision:'APPROVE',reason:'Synthetic decision after recognition'}})).request;
  a(!decision.target.expected.some(e=>e.kind==='evidence'));
  const result=await call('execute',decision);a.equal(result.result.state,'COMPLETED',JSON.stringify(result));a.equal((await db.doc('activityJustifications/SYN_J').get()).data().state,'APPROVED');
 }finally{await db.terminate();}
});


