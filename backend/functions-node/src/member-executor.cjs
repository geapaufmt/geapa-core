'use strict';
const {Store,decode,encode}=require('./firestore.cjs'),{createDomain}=require('./domain.cjs');
const {fence,fail,SHARED,EXECUTOR,PROJECT,PRODUCTION,profile}=require('./environment.cjs'),{authenticate,appCheckGate}=require('./auth.cjs');
const {windowLive}=require('./execution-window.cjs');
const CONTROL='memberExecutorControl/current';
const types=['members.self.update','members.profile.correct','members.ingress.execute','members.voluntary.submit','members.voluntary.review','members.membership.voluntaryExit','members.onboarding.complete'];
function control(c,uid){fence();if(!SHARED||c?.schemaVersion!=='member-executor-control-v1'||c.projectId!==PROJECT||!c.enabled||!windowLive(c,Date.now())||!Array.isArray(c.actorUids)||c.actorUids.length>(require('./environment.cjs').PRODUCTION?128:16)||!c.actorUids.includes(uid)||!Array.isArray(c.scopeIds)||!c.scopeIds.length||c.scopeIds.length>4||!c.scopeIds.every(x=>/^[A-Za-z0-9_-]{1,48}$/.test(x))||c.appCheck?.mode!=='REQUIRED'||c.browserEnabled!==true)fail('PILOT_DISABLED');return c;}
class MemberStore extends Store{
 constructor(...args){super(...args);this.specializedReaders=true;this.queryConcurrency=4;}
 guard(){const c=control(this.hostRead(CONTROL),this.identity?.uid);if(this.pinned&&JSON.stringify(c)!==this.pinned)fail('ACCESS_REVOKED');this.pinned=JSON.stringify(c);this.control=c;for(const id of c.scopeIds)this.hostRead('memberExecutorScopes/'+id);return c;}
 io(raw){const x=JSON.parse(raw);if(x.kind==='auditClock'){if(!this.active?.native)fail('TRANSACTION_REQUIRED');return JSON.stringify(this.active.preparationAt);}if(x.kind==='pilotConfig')return JSON.stringify({root:'',shared:true,executor:EXECUTOR,settings:{commands:[],policyIds:[],admissionsPerHour:20,jobs:[],activityAgendaDirect:false}});return super.io(raw);}
 validateWrites(s){
  const op=this.operationId,person=this.personId;
  if(s.writes.length&&!person)fail('PILOT_RESOURCE_DENIED');
  const scoped=this.control.scopeIds.map(id=>this.hostRead('memberExecutorScopes/'+id));if(s.missing.size)return;
  const fresh=person==='P_'+op&&this.prepared?.payload.operation==='PRE_REGISTER'&&scoped.some(v=>v?.projectId===PROJECT&&v.synthetic===!PRODUCTION&&v.approvalRefs?.includes(this.prepared.payload.approvalRef));
  if(!scoped.some(v=>v?.synthetic===!PRODUCTION&&v.projectId===PROJECT&&v.personIds?.includes(person))&&!fresh)fail('PILOT_RESOURCE_DENIED');
  const technical=['memberAdmissionTickets','memberAdmissionCounters','memberAdmissionUses','commandProvenance','commandRequests','commandReceipts','commandLedger','commandTombstones','commandLeases','commandReconciliation','memberAudit','commandProjectionInvalidations','commandExecutorOwners','commandExecutorRoutes'];
  const personCollections=['memberIdentities','memberIdentifiers','memberIdentifierClaims','memberPrivateProfiles','memberLinks','memberEvents','memberAggregateHeads','memberCurrentStates','memberStatusViews','memberScopeGrants','memberPortalProfiles','memberPublicProfiles','memberCurrentFunctions','memberCurrentDirectors','memberTransitionQueue','memberAuthzSubjects','memberVoluntaryRequests','memberVoluntaryIndexes','memberEnrollments','memberEnrollmentIndexes','memberEnrollmentFacts','memberProfileCorrections','memberProfileCorrectionIndexes'];
  for(const w of s.writes){const [col,id]=w.path.split('/'),d=w.mode==='delete'?null:decode(w.data),old=this.hostRead(w.path);if(s.missing.size)return;if(w.mode==='delete'){if(['memberCurrentFunctions','memberCurrentDirectors','memberFunctionProofs'].includes(col)&&id===person&&w.path.split('/')[2]==='items'&&w.path.split('/').length===4)continue;fail('PILOT_WRITE_DENIED');}
   if(col==='memberEnrollmentSessions'&&this.opening&&id===this.identity.uid&&d.uid===id&&d.personId===person&&d.environment===profile.environment&&d.projectId===PROJECT)continue;
   if(this.preparation&&!['memberPreparedIntents','devProtectedInputs','memberEvidenceContents','memberEvidenceAttestations'].includes(col))fail('PILOT_WRITE_DENIED');
   if(['memberPreparedIntents','devProtectedInputs'].includes(col)){if(id!==op||d.ownerUid!==this.identity.uid)fail('PILOT_WRITE_DENIED');continue;}
   if(col==='memberEvidenceContents'){if(d.target?.personId!==person)fail('PILOT_WRITE_DENIED');continue;}
   if(col==='memberEvidenceAttestations'){if(id!=='MVE_'+op||d.holderId!==person)fail('PILOT_WRITE_DENIED');continue;}
   if(technical.includes(col)){if(['commandRequests','commandReceipts','commandLedger','commandProvenance','commandLeases','commandExecutorOwners','commandProjectionInvalidations'].includes(col)&&id!==op)fail('PILOT_WRITE_DENIED');if(col==='commandExecutorRoutes'&&id!==this.routeId)fail('PILOT_WRITE_DENIED');continue;}
   if(col==='institutionalAccounts'&&id===person)continue;
   if(col==='projectionProvenance'&&['memberPortalProfiles/'+person,'memberStatusViews/'+person].includes(d.target)&&id==='M1_'+d.target.replace('/','_'))continue;
   if(w.path==='portalAccessStatus/'+person+'/state/current'||col==='memberFunctionProofs'&&id===person)continue;
   if(w.path==='presentationMemberViews/'+person){if(!old||d.state!=='INVALIDATED'||d.viewEpoch!==old.viewEpoch+1)fail('PILOT_WRITE_DENIED');continue;}
   if(personCollections.includes(col)){const owner=d.idPessoa||d.personId||(id===person?person:null),prior=old&&(old.idPessoa||old.personId||(id===person?person:null));if(owner!==person||old&&prior!==person)fail('PILOT_RESOURCE_DENIED');continue;}
   if(['memberAuthBindings','portalUsers','portalAccessStatus'].includes(col)){if(d.idPessoa!==person||old&&old.idPessoa!==person||!this.control.actorUids.includes(id))fail('PILOT_RESOURCE_DENIED');continue;}
   if(col==='jobOutbox'&&id===op&&d.kind==='MEMBERS_NOTIFICATION'&&d.mode==='DRY_RUN'&&d.projectId===PROJECT&&d.environment===profile.environment&&d.personId===person)continue;
   if(w.path==='portalReaderControl/current'){if(!old||d.generation!==old.generation+1)fail('PILOT_WRITE_DENIED');for(const k of Object.keys(old))if(k!=='generation'&&JSON.stringify(encode(old[k]))!==JSON.stringify(encode(d[k])))fail('PILOT_WRITE_DENIED');continue;}
   if(w.path==='portalReaderResourceFences/members_operations'){if(d.revision!==(old?.revision||0)+1)fail('PILOT_WRITE_DENIED');continue;}
   if(w.path==='presentationCapacityHeads/registry'){if(!old||d.membersEpoch!==old.membersEpoch+1)fail('PILOT_WRITE_DENIED');for(const k of Object.keys(old))if(!['membersEpoch','revision','updatedAt'].includes(k)&&JSON.stringify(encode(d[k]))!==JSON.stringify(encode(old[k])))fail('PILOT_WRITE_DENIED');continue;}
   fail('PILOT_WRITE_DENIED');
  }
 }
}
async function executeMember({db,auth,appCheck,token,appCheckToken,origin,body,hooks={}}){
 const started=performance.now(),env=fence();if(!env.emulator&&Object.keys(hooks).length)fail('DESTINATION_INVALID');
 if(!body||Object.keys(body).sort().join()!=='action,request'||!['memberPrepare','memberExecute','memberStatus','memberRecover','memberEnrollmentOpen','memberEnrollmentRead','memberEnrollmentPrepare'].includes(body.action))fail('COMMAND_TYPE_INVALID');
 const identity=await authenticate(auth,token),store=new MemberStore(db,identity),domain=createDomain(x=>store.io(x));
 const run=(fn,v)=>{const r=domain.run(fn,v);return r===undefined?undefined:JSON.parse(r);};run('(x)=>nodeIdentity(x)',identity);
 await store.read(()=>store.guard());await appCheckGate(appCheck,appCheckToken,store.control,origin);store.beforeCommit=()=>authenticate(auth,token);
 if(store.control.commandTypes!==undefined&&(!Array.isArray(store.control.commandTypes)||!store.control.commandTypes.length||!store.control.commandTypes.every(t=>types.includes(t))))fail('PILOT_DISABLED');
 const commandEnabled=t=>!store.control.commandTypes||store.control.commandTypes.includes(t);
 if(body.action.startsWith('memberEnrollment')&&!commandEnabled('members.onboarding.complete'))fail('PILOT_DISABLED');
 if(body.action==='memberPrepare'&&store.control.portalRoutes&&!store.control.portalRoutes.includes(body.request?.route))fail('PILOT_RESOURCE_DENIED');
 const checked=fn=>()=>{store.guard();return fn();};
 const select=r=>{if(!types.includes(r?.commandType)||!/^[a-f0-9]{32}$/.test(r.operationId))fail('COMMAND_TYPE_INVALID');store.operationId=r.operationId;store.personId=r.target?.personId;store.prepared=r;
  if(store.preparation&&!commandEnabled(r.commandType))fail('PILOT_RESOURCE_DENIED');
  store.routeId=run('(r)=>{const f=f4MemberExecutor_(),o=nodeRead("commandExecutorOwners/"+r.operationId),id=o?.routeId||f.select(nodeRead,r,nodeConfig);f4Core_().need(id,"PILOT_RESOURCE_DENIED");const v=f.route(nodeRead,nodeConfig,id);f.fields(nodeRead,r,v);return id;}',r);
  if(store.preparation)run('(id)=>{const v=f4MemberExecutor_().route(nodeRead,nodeConfig,id);f4Core_().need(v.state==="ACTIVE"&&f4ExecutorWindow_().live(v,Date.now())&&v.actorUids.includes(nodeActor.uid),"ACCESS_REVOKED");}',store.routeId);
 };
 let key,result;
 if(['memberEnrollmentOpen','memberEnrollmentRead'].includes(body.action)){
  if(Object.keys(body.request).length)fail('COMMAND_PAYLOAD_INVALID');
  let verified;if(body.action==='memberEnrollmentOpen'){const tokenData=await auth.verifyIdToken(token,true);if(tokenData.email_verified!==true||typeof tokenData.email!=='string')fail('AUTH_BINDING_INVALID');verified={...identity,projectId:PROJECT,email:tokenData.email,emailVerified:true};}
  store.opening=body.action==='memberEnrollmentOpen';store.preparation=store.opening;
  try{result=await store.transaction(checked(()=>{if(store.opening){const opened=run('(v)=>f4MemberEnrollmentAccess_().establish(nodeTx,v,Date.now(),nodeConfig)',verified);store.personId=opened.personId;}else store.personId=run('()=>f4MemberEnrollmentAccess_().source(nodeRead,nodeActor.uid,Date.now()).enrollment.personId');return run('()=>f4MemberEnrollmentPortal_({config:nodeConfig,identity:nodeActor,clock:Date.now}).context({})');}));}catch(error){if(!store.opening||error.code!=='ACCESS_REVOKED')throw error;store.preparation=false;result=await store.transaction(checked(()=>{store.personId=store.hostRead('memberEnrollmentSessions/'+identity.uid)?.personId;return run('(v)=>f4MemberEnrollmentPortal_({config:nodeConfig,identity:nodeActor,clock:Date.now}).status(v)',verified);}));}
  return {result,metrics:{...store.metrics,totalMs:performance.now()-started}};
 }
 if(body.action==='memberPrepare'||body.action==='memberEnrollmentPrepare'){
  store.operationId=body.request.operationId;store.preparation=true;
  try{result=await store.transaction(checked(()=>{const r=run(body.action==='memberEnrollmentPrepare'?'(q)=>f4MemberEnrollmentPortal_({config:nodeConfig,identity:nodeActor,clock:Date.now}).prepare(q)':'(q)=>f4MemberPortalPrepare_({config:nodeConfig,identity:nodeActor,clock:Date.now}).prepare(q)',body.request);select(r);return r;}));}finally{store.preparation=false;}
  const key=await store.read(checked(()=>run('(r)=>({operationId:r.operationId,commandType:r.commandType,digest:f4MemberCommandCatalog_().digest(nodeConfig.projectId,{...r,ownerUid:nodeActor.uid,expectedRevisions:nodeRead("devProtectedInputs/"+r.operationId).expectedRevisions})})',result)));return {request:result,key,metrics:{...store.metrics,totalMs:performance.now()-started}};
 }
 await store.read(checked(()=>run('()=>{globalThis.pipeline=nodePipeline();return null;}')));
 if(body.action==='memberExecute'){
  await store.read(checked(()=>select(body.request)));
  key=await store.read(checked(()=>run('(r)=>{const C=f4Core_(),K=f4MemberCommandCatalog_();C.exact(r,["schemaVersion","operationId","commandType","target","payload","evidenceRefs","decisionRef"]);const p=nodeRead("devProtectedInputs/"+r.operationId);C.need(p?.ownerUid===nodeActor.uid,"ACCESS_REVOKED");return {operationId:r.operationId,commandType:r.commandType,digest:K.digest(nodeConfig.projectId,{...r,ownerUid:nodeActor.uid,expectedRevisions:p.expectedRevisions})};}',body.request)));
  const exists=await store.read(()=>!!store.hostRead('commandRequests/'+key.operationId));
  if(!exists){
   const ticket=await store.transaction(checked(()=>run('(r)=>pipeline.admission.issueBrowser(nodeTx,pipeline.admission.authenticate("verified"),r)',body.request)));
   await store.transaction(checked(()=>run('(t)=>{const C=f4Core_(),T=f4Firestore_().Timestamp,r={...t.requestEnvelope,ticketId:t.ticketId,receivedAt:T.fromMillis(Date.now()),state:"RECEIVED"},old=nodeRead("commandRequests/"+r.operationId),used=nodeRead("memberAdmissionUses/"+r.ticketId);f4ExecutorOwnership_().recover(nodeRead,{operationId:r.operationId,commandType:r.commandType,digest:r.digest},nodeActor.uid,nodeConfig,Date.now());if(old){C.need(old.ownerUid===r.ownerUid&&old.digest===r.digest,"OPERATION_CONFLICT");return;}C.need(!used,"OPERATION_CONFLICT");nodeTx.set("commandRequests/"+r.operationId,r);nodeTx.set("memberAdmissionUses/"+r.ticketId,{schemaVersion:"member-admission-use-v13",ticketId:r.ticketId,operationId:r.operationId,ownerUid:r.ownerUid,reservedAt:r.receivedAt});}',ticket)));
  }
 }else key=body.request;
 if(!key||Object.keys(key).sort().join()!=='commandType,digest,operationId'||!types.includes(key.commandType)||!/^[a-f0-9]{32}$/.test(key.operationId)||!/^[a-f0-9]{64}$/.test(key.digest))fail('COMMAND_PAYLOAD_INVALID');
 const guarded=fn=>checked(()=>{const r=store.hostRead('commandRequests/'+key.operationId);if(!r)fail('NOT_FOUND');select(r);return fn();});
 await authenticate(auth,token);store.beforeCommit=()=>authenticate(auth,token);
 if(body.action==='memberStatus')result=await store.transaction(guarded(()=>run('(k)=>pipeline.recovery.readStatus(nodeTx,pipeline.recovery.authenticate("verified"),k)',key)));
 else{const claim=await store.transaction(guarded(()=>run('(k)=>pipeline.recovery.claim(nodeTx,pipeline.recovery.authenticate("verified"),k,f4CommandExecutorId_())',key)));
  if(claim.state!=='CLAIMED')result=claim;else{await hooks.afterClaim?.(key);try{result=await store.transaction(guarded(()=>run('(x)=>pipeline.recovery.finalize(nodeTx,pipeline.recovery.authenticate("verified"),x.key,x.handle)',{key,handle:claim.handle})));}catch(e){await store.transaction(guarded(()=>run('(x)=>pipeline.recovery.release(nodeTx,pipeline.recovery.authenticate("verified"),x.key,x.handle)',{key,handle:claim.handle}))).catch(()=>{});throw e;}await hooks.afterCommit?.(key);}}
 const dto=await store.transaction(guarded(()=>run('(k)=>f4MemberConsumers_().createLocal({config:nodeConfig,recovery:pipeline.recovery}).status(nodeTx,pipeline.recovery.authenticate("verified"),k)',key)));
 return {key,result,dto,projection:{state:'INVALIDATED',rebuildRequired:true},metrics:{...store.metrics,totalMs:performance.now()-started}};
}
module.exports={MemberStore,executeMember,types,CONTROL,control};
