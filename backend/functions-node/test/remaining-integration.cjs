'use strict';
require('../src/local-network.cjs');
const a=require('node:assert/strict'),fs=require('fs'),crypto=require('crypto');
const {initializeApp}=require('firebase-admin/app'),{getFirestore}=require('firebase-admin/firestore'),{getAuth}=require('firebase-admin/auth');
const {DEMO,PROJECT,EXECUTOR,ACTIVITY,fence}=require('../src/environment.cjs'),{decode}=require('../src/firestore.cjs'),{execute}=require('../src/executor.cjs');
const recoveryModule=require('../src/projection-recovery.cjs'),recoverOriginal=recoveryModule.recoverProjection;recoveryModule.recoverProjection=async x=>{try{return await recoverOriginal(x);}catch(e){console.error('PROJECTION_DIAGNOSTIC '+x.body.request.viewId+' '+e.code+' '+e.stack);throw e;}};
fence();const db=getFirestore(initializeApp({projectId:DEMO})),auth=getAuth(),tests=[],bench=[];
const base=()=>Object.fromEntries(Object.entries(require('./fixtures.cjs').seed().docs).map(([p,v])=>[p,decode({mapValue:{fields:v.fields}})]));
require('./flow-fixture.cjs'); // Initialize current Core even when running one filtered group.
const get=async p=>(await db.doc(p).get()).data(),id=()=>crypto.randomBytes(16).toString('hex');let token;
const call=(action,request,hooks)=>execute({db,auth,token,appCheck:{verifyToken:async()=>({appId:'local-app'})},appCheckToken:'local',origin:'https://geapa-convergencia-dev.web.app',body:{action,request},hooks});
async function check(name,fn){if(process.env.GEAPA_FLOW_TEST_FILTER&&!(process.env.GEAPA_CYCLE_BOUNDARY_TEST==='true'?new RegExp(process.env.GEAPA_FLOW_TEST_FILTER).test(name):name.includes(process.env.GEAPA_FLOW_TEST_FILTER)))return;try{await fn();tests.push({name,passed:true});console.log('PASS '+name);}catch(e){tests.push({name,passed:false,error:e.message,reason:e.reasonCode,stack:e.stack});console.error('FAIL '+name+' '+e.reasonCode+' '+e.stack);}}
async function seed(extra={},command='ATTENDANCE_FINALIZE',self=false,cap='ACTIVITY_ATTENDANCE_MANAGE',resources=[],absent=[]){
 resources=[...new Set(resources)];
 await fetch("http://127.0.0.1:8187/emulator/v1/projects/"+DEMO+"/databases/(default)/documents",{method:"DELETE"});
 const docs={...base(),...extra},now=Date.now();
 // The shared registry is an existing canonical fence, with immutable creation time.
 if(docs['presentationCapacityHeads/registry']&&!docs['presentationCapacityHeads/registry'].createdAt)docs['presentationCapacityHeads/registry'].createdAt=require('firebase-admin/firestore').Timestamp.fromMillis(now-1000);
 for(const path of absent)delete docs[path];
 const old=docs['memberAuthzPolicies/CF_DEV_ACT_POLICY'];
 // Use the existing synthetic presidential role, not an extra real capability.
 const promote=v=>{if(v&&typeof v==='object'){for(const k of Object.keys(v)){if(k==='roleKey'&&v[k]==='DIRETOR_ADMINISTRATIVO')v[k]='PRESIDENTE';else if(['scopeKey','actorScope'].includes(k)&&v[k]==='ADMINISTRACAO')v[k]='PRESIDENCIA';else promote(v[k]);}}};promote(docs);

 docs['memberAuthzPolicies/FLOW_POLICY']={...old,id:'FLOW_POLICY',capabilities:[cap],proofKind:self?'SELF':'ROLE',actorSelector:self?{kind:'SELF'}:old.actorSelector,targets:{ownPerson:self,allowedTargetScopes:self?['SELF']:['ACTIVITIES','PRESENTATIONS_PLAN','PRESC_'+'c'.repeat(32)],allowedBoards:[],allowedRoles:[],allowedPersons:self?[]:['PES-000001','PES-000002'],allowedResources:resources},fieldMasks:{[cap]:[]}};
 if(command.startsWith('ATTENDANCE')&&!self)docs['coreDevAdmins/CF_ADMIN']={schemaVersion:'core-dev-admin-v1',projectId:PROJECT,environment:'DEV',uid:'CF_ADMIN',active:true,revision:1,issuedAt:require('firebase-admin/firestore').Timestamp.fromMillis(now-1000),expiresAt:require('firebase-admin/firestore').Timestamp.fromMillis(now+3600000),permissions:'FUNCTIONAL_ALL',exportable:false};
 const c={schemaVersion:'activity-executor-control-v1',projectId:PROJECT,enabled:true,revision:1,actorUids:['CF_ADMIN'],policyIds:['FLOW_POLICY','CF_DEV_ACT_POLICY'],startsAt:now-1000,expiresAt:now+3600000,browserEnabled:true,appCheck:{mode:'REQUIRED',appIds:['local-app']}};
 if(command==='SESSION_CONFIGURE'){docs['memberAuthzPolicies/FLOW_READ']={...old,id:'FLOW_READ',capabilities:['PRESENTATION_BOARD_READ'],targets:{ownPerson:false,allowedTargetScopes:['PRESENTATIONS_PLAN','PRESC_'+'c'.repeat(32)],allowedBoards:[],allowedRoles:[],allowedPersons:['PES-000001'],allowedResources:[],resourceScope:{kind:'PRESENTATIONS',cycleAccess:'ALL'}},fieldMasks:{PRESENTATION_BOARD_READ:[]}};c.policyIds.push('FLOW_READ');}
 docs['activityExecutorControl/current']=c;
 docs['projectionRecoveryControl/current']={...c,schemaVersion:'activity-projection-recovery-v1',namespace:'',readModelCommands:['ACTIVITY_SAVE','ATTENDANCE_FINALIZE']};
 const resourcePaths=[...new Set([...Object.keys(docs),...resources].filter(p=>f4ActivitiesModel_().validPath(p)))];
 docs['actFlowExecutorScopes/FLOW_FIXTURE']={id:'FLOW_FIXTURE',revision:1,projectId:PROJECT,synthetic:true,resourcePaths};
 if(process.env.GEAPA_CYCLE_BOUNDARY_TEST==='true'){
  const scope=docs['actFlowExecutorScopes/FLOW_FIXTURE'];
  scope.resourcePaths=['presentationCycles/'+ 'c'.repeat(32),'presentationCycles/fn_cycle'];
  scope.operational={schemaVersion:'act-cycle-boundary-v1',cycleIds:['c'.repeat(32),'fn_cycle'],readOnlyPaths:resourcePaths.filter(p=>p.split('/').length===2&&(/^(activityTypes|activityPolicies|activityAttachmentPolicies|academicSemesters|presentationPolicyReferences|migrationEvidence)\//.test(p)||['presentationCapacityHeads/registry','activityAttendanceRegistry/current'].includes(p))),commands:require('../src/act-flows.cjs').types.filter(c=>!['ACT.ACTIVITY_PROMOTE_OPERATIONAL','ACT.ACTIVITY_ADMIN_REVIEW','ACT.HISTORICAL_PREAGENDA_RECOGNIZE'].includes(c))};
  if(process.env.GEAPA_CYCLE_POLICY_TEST==='true'){const policy=docs['memberAuthzPolicies/FLOW_POLICY'],kind=self?(cap.startsWith('ACTIVITY_')?'OWN_ATTENDANCE':'OWN_PRESENTATIONS'):cap.startsWith('ACTIVITY_')?'ACTIVITIES':'PRESENTATIONS';policy.targets.resourceScope={kind,cycleIds:scope.operational.cycleIds};policy.targets.allowedResources=scope.operational.readOnlyPaths;policy.targets.allowedTargetScopes=self?['SELF']:kind==='ACTIVITIES'?['ACTIVITIES']:['PRESENTATIONS_PLAN','PRESC_'+ 'c'.repeat(32),'PRESC_fn_cycle'];}

 }
 for(const cmd of require('../src/act-flows.cjs').catalog){const routeId='ACT_FLOW_'+cmd;docs['commandExecutorRoutes/'+routeId]={schemaVersion:'geapa-act-flow-route-v1',projectId:PROJECT,routeId,commandType:'ACT.'+cmd,syntheticOnly:true,scopeId:'FLOW_FIXTURE',executor:EXECUTOR,state:'ACTIVE',epoch:1,revision:1,actorUids:['CF_ADMIN'],operationIds:[],maxOperations:100,startsAt:now-1000,expiresAt:now+3600000};}
 require('./production-fixture-profile.cjs')(docs);
 const entries=Object.entries(docs);for(let i=0;i<entries.length;i+=300){const b=db.batch();for(const [p,d]of entries.slice(i,i+300))b.set(db.doc(p),d);await b.commit();}
 return docs;
}
if(require.main===module)(async()=>{try{
 await auth.createUser({uid:'CF_ADMIN',email:'flow@example.invalid',password:'LocalSyntheticOnly123!'});
 token=(await (await fetch('http://127.0.0.1:9197/identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=local',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email:'flow@example.invalid',password:'LocalSyntheticOnly123!',returnSecureToken:true})})).json()).idToken;

 for(const command of [...require('./remaining-fixture.cjs').extra,'DYNAMIC_SUBMIT','DYNAMIC_DECIDE','DYNAMIC_CANCEL','PUBLICATION_APPROVE','PUBLICATION_WITHDRAW','ARTIFACT_UPLOAD_REQUEST','MATERIAL_SUBMIT','MATERIAL_REVIEW','MATERIAL_RECEIVE','MATERIAL_FINAL_EVALUATE','MATERIAL_EXCEPTION','MATERIAL_CORRECTION_REQUEST','CORRECTION_OPPORTUNITY_OPEN','MATERIAL_CORRECTION_SUBMIT','CYCLE_CREATE','CONFIG_RECORD','CONFIG_ACTIVATE']){
  await check(command+' canonical SDK effect and replay',async()=>{
   const sample=require('./remaining-fixture.cjs').remaining(command),p=sample.request.payload;if(sample.request.target.authorization.kind==='ROLE')sample.request.target.authorization.actorScope='PRESIDENCIA';
   if(command==='CORRECTION_OPPORTUNITY_OPEN')p.dueAt=Date.now()+86400000;
   if(command==='CONFIG_RECORD'){p.effectiveFrom=Date.now()+86400000;p.parameters.calendarMode='BOARD_INITIAL';p.parameterClasses='STRUCTURAL';sample.docs['presentationCycles/'+'c'.repeat(32)].status='DRAFT';}
   if(command==='CONFIG_ACTIVATE')for(const [k,d]of Object.entries(sample.docs))if(k.includes('/configRevisions/'))d.parameters.calendarMode='BOARD_INITIAL';
   if(command==='DYNAMIC_DECIDE'){const f=sample.docs['presentationFlowCases/'+p.proposalRef.id];sample.docs['presentationUnitHeads/'+f.unitRef.id].currentDynamicProposalRef=p.proposalRef;p.decision='REQUEST_CORRECTION';}

   if(command==='PUBLICATION_APPROVE'){
    sample.docs['presentationPolicyReferences/policy'].publicFieldSelection=p.publicFieldSelection;
    const pr=sample.docs['canonicalPresentations/presentation'];pr.finalConfirmationRef={kind:'flow',id:'final_confirmation'};
    sample.docs['presentationFlowCases/final_confirmation']={id:'final_confirmation',revision:1,cycleRef:pr.cycleRef,kind:'FINAL_CONFIRMATION',state:'CONFIRMED',presentationRef:{kind:'presentation',id:pr.id},title:pr.title,thematicRevisionRef:sample.docs['presentationThematicRequests/theme'].submittedRevisionRef};
   }
   if(command==='PUBLICATION_WITHDRAW'){p.presentationRef={kind:'presentation',id:'presentation'};sample.request.target.scope.targetRefs.push(p.presentationRef);sample.docs['publicPresentations/presentation']={publicId:'presentation',revision:1,schemaVersion:1,sanitized:true,published:true,cycleId:'c'.repeat(32),sourceSnapshotRef:'snapshot'};}
   if(command==='MATERIAL_EXCEPTION'){const req=sample.docs['presentationMaterialRequirements/requirement'];p.exceptionalDecision.coveredItemIds=req.items.map(i=>i.itemId);const m=sample.docs['presentationCapacityAssessments/manifest'];m.items=req.items.map(i=>({itemId:i.itemId}));m.digest=f4ActivitiesModel_().hash(m.items);}
   await seed(sample.docs,command,sample.request.target.authorization.kind==='SELF',sample.capability,sample.request.target.scope.targetRefs.map(r=>f4ActivitiesModel_().path(r)),sample.request.target.expected.filter(r=>r.exists===false).map(r=>f4ActivitiesModel_().path({kind:r.kind,id:r.id})));
   const prepared=await call('flowPrepare',{operationId:sample.request.operationId,commandType:sample.request.commandType,payload:p});const out=await call('execute',prepared.request);a.equal(out.result.state,'COMPLETED',JSON.stringify(out.result));
   process.env.GEAPA_ACT_EXTENSIONS_ENABLED='false';let replay;try{replay=await call('recover',out.key);}finally{process.env.GEAPA_ACT_EXTENSIONS_ENABLED='true';}a.equal(replay.result.state,'RECOVERED');a.equal(replay.metrics.writes,0);
   if(process.env.GEAPA_CYCLE_BOUNDARY_TEST==='true')a.deepEqual((await get('actFlowExecutorScopes/FLOW_FIXTURE')).resourcePaths,['presentationCycles/'+ 'c'.repeat(32),'presentationCycles/fn_cycle']);
   bench.push({command,metrics:out.metrics,projection:out.projection});
  });
 }
 async function review(){
  const s=require('./remaining-fixture.cjs').remaining('MATERIAL_REVIEW');
  await seed(s.docs,'MATERIAL_REVIEW',false,s.capability,s.request.target.scope.targetRefs.map(f4ActivitiesModel_().path));
  return (await call('flowPrepare',{operationId:id(),commandType:'ACT.MATERIAL_REVIEW',payload:s.request.payload})).request;
 }
 await check('extension disabled refuses fresh admission; no effect',async()=>{const q=await review();process.env.GEAPA_ACT_EXTENSIONS_ENABLED='false';try{await a.rejects(call('execute',q),/COMMAND_TYPE_INVALID/)}finally{process.env.GEAPA_ACT_EXTENSIONS_ENABLED='true'}a.equal((await db.collection('commandRequests').get()).size,0)});
 await check('current revision field mask rejects injected fields before admission',async()=>{const q=await review(),before=await get('presentationMaterialRequirements/requirement');for(const value of [{actor:{accessProfile:'DEV_ADMIN'}},{receiptRef:{kind:'receipt',id:'forged'}},{status:'REGULARIZED'}])await a.rejects(call('execute',{...q,payload:{...q.payload,...value}}),e=>['VALIDATION_ERROR','SCHEMA_INVALID','COMMAND_PAYLOAD_INVALID'].includes(e.code));a.equal((await get('presentationMaterialRequirements/requirement')).revision,before.revision);a.equal((await db.collection('commandRequests').get()).size,0)});
 await check('capability revoked after preparation refuses admission',async()=>{const q=await review();await db.doc('memberAuthzPolicies/FLOW_POLICY').update({active:false});await a.rejects(call('execute',q));a.equal((await db.collection('commandRequests').get()).size,0)});
 await check('ordinary member cannot inherit a board capability',async()=>{const q=await review();await db.doc('memberFunctions/VIG-001').update({active:false});await a.rejects(call('execute',q));a.equal((await db.collection('commandRequests').get()).size,0)});
 await check('two current revisions race: exactly one review and no lost update',async()=>{const q=await review(),other={...q,operationId:id()};const results=await Promise.allSettled([call('execute',q),call('execute',other)]);a.equal(results.filter(r=>r.status==='fulfilled'&&r.value.result.state==='COMPLETED').length,1,JSON.stringify(results));a.equal((await db.collection('presentationMaterialReviews').get()).docs.filter(d=>[q.operationId,other.operationId].includes(d.data().operationId)).length,1)});
 await check('lost commit response recovers same operation after route rollback',async()=>{const q=await review();let key;await a.rejects(call('execute',q,{afterCommit:k=>{key=k;throw Error('SIMULATED_RESPONSE_LOSS')}}),/SIMULATED_RESPONSE_LOSS/);a(key);await db.doc('commandExecutorRoutes/ACT_FLOW_MATERIAL_REVIEW').update({executor:'APPS_SCRIPT_DEV',state:'CLOSED',epoch:2,revision:2});const r=await call('recover',key);a.equal(r.result.state,'RECOVERED');a.equal(r.metrics.writes,0);a.equal((await db.collection('presentationMaterialReviews').get()).docs.filter(d=>d.data().operationId===q.operationId).length,1)});
 await check('same operation with altered payload cannot duplicate review',async()=>{const q=await review(),r=await call('execute',q);a.equal(r.result.state,'COMPLETED');await a.rejects(call('execute',{...q,payload:{...q.payload,reason:'altered'}}));a.equal((await db.collection('presentationMaterialReviews').get()).docs.filter(d=>d.data().operationId===q.operationId).length,1)});
 await check('historical partial input stays read only',async()=>{const q=await review();await db.doc('canonicalPresentations/presentation').update({recordMode:'HISTORICAL_PARTIAL'});await a.rejects(call('execute',q),e=>['PILOT_RESOURCE_DENIED','DOMAIN_STATE_INVALID',...(process.env.GEAPA_CYCLE_BOUNDARY_TEST==='true'?['SCOPE_INVALID']:[])].includes(e.code));a.equal((await db.collection('commandRequests').get()).size,0)});
 for(const decision of ['APROVAR','SOLICITAR_AJUSTE'])await check('Portal material '+decision+' resumes canonical chained receipts',async()=>{
  await review();const policy=await get('memberAuthzPolicies/FLOW_POLICY');policy.capabilities=['PRESENTATION_BOARD_READ','PRESENTATION_MATERIAL_REVIEW'];policy.fieldMasks={PRESENTATION_BOARD_READ:[],PRESENTATION_MATERIAL_REVIEW:[]};policy.targets={ownPerson:false,allowedTargetScopes:['PRESC_'+'c'.repeat(32)],allowedBoards:[],allowedRoles:[],allowedPersons:['PES-000001'],allowedResources:[],resourceScope:{kind:'PRESENTATIONS',cycleAccess:'ALL'}};await db.doc('memberAuthzPolicies/FLOW_POLICY').set(policy);
  const rq=await get('presentationMaterialRequirements/requirement');await db.doc('presentationPolicyReferences/policy').update({portalReview:{approved:true,mode:'COMPLETE_FILE_REVIEW',coveredItemIds:rq.items.map(x=>x.itemId),evaluationCutoffRef:{kind:'policy',id:'cutoff'}}});await db.doc('presentationPolicyReferences/cutoff').set({id:'cutoff',revision:1,approved:true,active:true,cycleRef:rq.cycleRef,at:require('firebase-admin/firestore').Timestamp.fromMillis(Date.now()+3600000)});const scope=await get('actFlowExecutorScopes/FLOW_FIXTURE');scope.resourcePaths.push('presentationPolicyReferences/cutoff');await db.doc('actFlowExecutorScopes/FLOW_FIXTURE').set(scope);
  let priorSelector;const stages=decision==='APROVAR'?['REVIEW','RECEIVE','EVALUATE']:['REVIEW','REQUEST','OPPORTUNITY'];
  for(const stage of stages){const input={operationId:id(),route:'/v2/apresentacoes/material/revisar',payload:{idApresentacao:'APU_unit',basis:{unitRevision:1,requirementRevision:rq.revision},decision,publicReason:'Synthetic review',internalReason:'',stage,...(priorSelector?{priorSelector}:{}),...(decision==='SOLICITAR_AJUSTE'?{dueCivilDate:new Date(Date.now()+86400000*2).toISOString().slice(0,10)}:{})}};
   const p=await call('flowPortalPrepare',input),r=await call('execute',p.request);a.equal(r.result.state,'COMPLETED',JSON.stringify(r.result));priorSelector=r.key;
  }
  a.equal((await call('recover',priorSelector)).result.state,'RECOVERED');
 });
 await check('linked proof associates one confirmed file without deciding justification',async()=>{
  const s=require('./remaining-fixture.cjs').remaining('ATTENDANCE_PROOF_UPLOAD_REQUEST'),payload=s.request.payload;await seed(s.docs,'ATTENDANCE_PROOF_UPLOAD_REQUEST',true,s.capability,s.request.target.scope.targetRefs.map(f4ActivitiesModel_().path));
  const q=(await call('flowPrepare',{operationId:id(),commandType:'ACT.ATTENDANCE_PROOF_UPLOAD_REQUEST',payload})).request,r=await call('execute',q);a.equal(r.result.state,'COMPLETED');
  const effect=(await db.collection('jobOutbox').get()).docs.find(d=>d.data().createdByOperation===q.operationId),ref={kind:'outbox',id:effect.id},artifactRef={kind:'activityArtifact',id:'proof_fixture'};
  await effect.ref.update({status:'CONFIRMED',artifactRef,workspaceReservation:{fileId:'synthetic_file_123'}});await db.doc('activityArtifacts/proof_fixture').set({id:'proof_fixture',revision:1,attendanceRef:payload.attendanceRef,ownerPersonRef:'PES-000001',objectId:'synthetic_file_123'});
  const policy=await get('memberAuthzPolicies/FLOW_POLICY');policy.targets.allowedResources.push('jobOutbox/'+effect.id);await db.doc('memberAuthzPolicies/FLOW_POLICY').set(policy);const scope=await get('actFlowExecutorScopes/FLOW_FIXTURE');scope.resourcePaths.push('activityArtifacts/proof_fixture');await db.doc('actFlowExecutorScopes/FLOW_FIXTURE').set(scope);
  const request=(await call('flowPrepare',{operationId:id(),commandType:'ACT.ATTENDANCE_JUSTIFICATION_REQUEST',payload:{activityRef:payload.activityRef,attendanceRef:payload.attendanceRef,proofUploadRef:ref,reason:'Synthetic justification',lateAcknowledged:false}})).request;
  const result=await call('execute',request);a.equal(result.result.state,'COMPLETED',JSON.stringify(result.result));const out=await get(effect.ref.path),j=await get(f4ActivitiesModel_().path(out.justificationRef));a.equal(j.state,'PENDING');a.equal(j.proofArtifactRef.id,artifactRef.id);a.equal((await call('recover',result.key)).metrics.writes,0);
 });
 await check('linked photo submits confirmed effect once without external upload',async()=>{
  const s=require('./remaining-fixture.cjs').remaining('PHOTO_SUBMIT'),p=s.request.payload;delete p.link;p.effectRef={kind:'outbox',id:'photo_effect'};
  s.docs['jobOutbox/photo_effect']={id:'photo_effect',revision:1,effectKind:'DRIVE_CREATE',status:'CONFIRMED',policyRef:{kind:'policy',id:'photo_policy'},targetRef:p.presentationRef,actor:{uid:'CF_ADMIN'},details:{purpose:'PRESENTATION_PHOTO',photoRevision:0},artifactRef:{kind:'artifact',id:'photo_file'},workspaceReservation:{fileId:'synthetic_file_123'}};s.docs['presentationArtifacts/photo_file']={id:'photo_file',revision:1,presentationRef:p.presentationRef,objectId:'synthetic_file_123'};s.request.target.scope.targetRefs.push(p.effectRef);
  await seed(s.docs,'PHOTO_SUBMIT',true,s.capability,s.request.target.scope.targetRefs.map(f4ActivitiesModel_().path));const request=(await call('flowPrepare',{operationId:id(),commandType:'ACT.PHOTO_SUBMIT',payload:p})).request,r=await call('execute',request);a.equal(r.result.state,'COMPLETED',JSON.stringify(r.result));a.equal((await get('jobOutbox/photo_effect')).photoSubmitted,true);a.equal((await call('recover',r.key)).metrics.writes,0);
 });
 a.equal(tests.filter(t=>!t.passed).length,0,'All test groups must pass');
 }finally{fs.writeFileSync((process.env.GEAPA_TEST_EVIDENCE_DIR||'evidence')+'/remaining.json',JSON.stringify({tests,bench,passed:tests.filter(t=>t.passed).length},null,2));await db.terminate();}})().catch(e=>{console.error(e.stack);process.exitCode=1;});

module.exports={seed,db,auth};
