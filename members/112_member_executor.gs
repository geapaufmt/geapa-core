// Shared protected-input port used by Apps Script and Node.
function f4MemberProtectedInputs_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_();
  const K=f4MemberCommandCatalog_(),material=r=>Object.fromEntries(['commandType','target','payload','evidenceRefs','decisionRef'].map(k=>[k,r[k]]));
  function read(get,r,a){const v=get('devProtectedInputs/'+r.operationId);C.exact(v,['schemaVersion','operationId','ownerUid','materialHash','expectedRevisions','authRequests']);N(v.schemaVersion==='core-dev-protected-input-v1'&&v.operationId===r.operationId&&v.ownerUid===a.uid&&v.materialHash===C.hash(material(r)),'AUTH_BINDING_INVALID');return v;}
  return {revisions:(get,r,a)=>read(get,r,a).expectedRevisions.map(v=>{N((get(v.resourceType+'/'+v.id)?.revision||0)===v.revision,'REVISION_CONFLICT');return {...v};}),authorization:(get,r,a,n,end)=>{const v=read(get,r,a);if(f4MemberEnrollmentCommands_().is(r)){N(v.authRequests.length===1&&v.authRequests[0].schemaVersion===f4MemberEnrollmentAccess_().schema,'AUTH_BINDING_INVALID');return [f4MemberEnrollmentAccess_().request(get,a,r,n,end)];}const control=get('memberAuthorityControl/MEMBERS'),b=get('memberAuthBindings/'+a.uid),head=b&&get('memberAggregateHeads/'+b.idPessoa),subject=b&&get('memberAuthzSubjects/'+b.idPessoa);N(b&&head&&subject,'AUTH_BINDING_INVALID');return v.authRequests.map(raw=>{const q=F.decode(F.encode(raw));N(q.uid===a.uid&&K.types[r.commandType].includes(q.capability),'AUTH_BINDING_INVALID');const source=q.kind==='DEV_TECHNICAL'?f4DevAdministrativeAccess_().command(get,a,n,q.capability,q.target,r).source:q.kind==='DEV_ADMIN'?f4DevAdmin_().profile(get,a.uid,n,true):get((q.kind==='EXPLICIT'?'memberExplicitGrants/':'memberAuthzPolicies/')+q.sourceId);N(source,'AUTH_BINDING_INVALID');return {...q,expected:{authorityEpoch:control.epoch,policyEpoch:control.policyEpoch,readAuthzEpoch:control.readAuthzEpoch,subjectEpoch:subject.authzEpoch,bindingVersion:b.bindingVersion,personRevision:head.revision,authorizationRevision:source.revision},tokenExpiresAt:F.Timestamp.fromMillis(a.expiresAt),admissionExpiresAt:F.Timestamp.fromMillis(end)};});}};
}

// Technical scope only. All institutional decisions still use the member catalog.
function f4MemberExecutor_(){
 const C=f4Core_(),N=C.need,types=['members.binding.associate','members.attendance.record','members.discipline.dismiss','members.self.update','members.profile.correct','members.ingress.execute','members.voluntary.submit','members.voluntary.review','members.membership.voluntaryExit','members.onboarding.complete'];
 const supports=r=>types.includes(r?.commandType),key=t=>'MEMBERS_'+t.slice(8).replace(/\./g,'_').toUpperCase();
 function scope(get,v){const s=get('memberExecutorScopes/'+v.scopeId);N(s?.schemaVersion==='member-executor-scope-v1'&&f4ExecutionEnvironment_().scope(s)&&s.id===v.scopeId&&C.rev(s.revision)&&s.revision>0,'ACCESS_REVOKED');for(const k of ['personIds','approvalRefs'])N(Array.isArray(s[k])&&s[k].length<=100&&s[k].every(C.id)&&new Set(s[k]).size===s[k].length,'ACCESS_REVOKED');N(s.personIds.length>0||s.profileCorrectionCurrentMembers===true||s.bindingCurrentMembers===true,'ACCESS_REVOKED');N(s.profileCorrectionCurrentMembers===undefined||s.profileCorrectionCurrentMembers===true,'ACCESS_REVOKED');N(s.bindingCurrentMembers===undefined||s.bindingCurrentMembers===true,'ACCESS_REVOKED');return s;}
 function route(get,config,id){N(f4ExecutionEnvironment_().destination(config),'DESTINATION_INVALID');const v=get('commandExecutorRoutes/'+id);N(v?.schemaVersion==='member-executor-route-v1'&&v.projectId===config.projectId&&types.includes(v.commandType)&&id===key(v.commandType)&&v.routeId===id&&f4ExecutionEnvironment_().route(v)&&C.id(v.scopeId)&&f4ExecutionEnvironment_().executors.includes(v.executor)&&['ACTIVE','DRAINING','CLOSED'].includes(v.state)&&C.rev(v.revision)&&v.revision>0&&C.rev(v.epoch)&&v.epoch>0&&Array.isArray(v.actorUids)&&v.actorUids.length>0&&v.actorUids.length<=(f4ExecutionEnvironment_().production?128:16)&&v.actorUids.every(C.id)&&new Set(v.actorUids).size===v.actorUids.length&&Array.isArray(v.operationIds)&&v.operationIds.every(C.op)&&new Set(v.operationIds).size===v.operationIds.length&&v.operationIds.length<=100&&f4ExecutorBudget_().valid(v)&&f4ExecutorWindow_().valid(v),'ACCESS_REVOKED');scope(get,v);return v;}
 function enrollmentScope(get,personId,refs){
  if(!/^P_[a-f0-9]{32}$/.test(personId))return false;
  const id=personId.slice(2),e=get('memberEnrollments/'+id);
  if(!e||e.id!==id||e.personId!==personId)return false;
  return refs.includes(e.approvalId||f4MemberPreRegistration_().key(e.registration));
 }
 function invitationActor(get,r,uid,v){
  if(v.commandType!=='members.onboarding.complete'||r.payload?.operation!=='ENROLLMENT')return false;
  const session=get('memberEnrollmentSessions/'+uid);
  return session?.uid===uid&&session.personId===r.target.personId&&session.enrollmentId===r.target.enrollmentId;
 }
 function fields(get,r,v){N(supports(r)&&r.commandType===v.commandType,'COMMAND_TYPE_INVALID');const s=scope(get,v),id=r.target?.personId;N(C.id(id),'COMMAND_PAYLOAD_INVALID');
  if(s.profileCorrectionRequestsOnly===true&&r.commandType==='members.profile.correct')N(r.payload?.operation==='CORRECTION_REQUEST','PILOT_RESOURCE_DENIED');
  // New identity only from a specifically approved synthetic registration.
  const registration=r.commandType==='members.ingress.execute'&&r.payload?.operation==='PRE_REGISTER'&&id==='P_'+r.operationId&&s.approvalRefs.includes(r.payload.approvalRef);
  N(s.personIds.includes(id)||registration||currentProfile(get,r,s)||currentBinding(get,r,s)||enrollmentScope(get,id,s.approvalRefs)&&['members.ingress.execute','members.onboarding.complete'].includes(r.commandType),'PILOT_RESOURCE_DENIED');
  if(r.payload?.action==='AMEND'&&['rga','email'].includes(r.payload.field)&&!s.personIds.includes(id)&&!get('commandExecutorOwners/'+r.operationId)){
   const old=get('memberEnrollments/'+r.payload.enrollmentId);
   const registration=f4MemberEnrollmentContract_().quick({...old.registration,[r.payload.field]:r.payload.value});
   const next=f4MemberPreRegistration_().locate(get,registration,old.approval.reference,Date.now());
   N(s.approvalRefs.includes(next.id),'PILOT_RESOURCE_DENIED');
  }
  if(registration&&!get('commandExecutorOwners/'+r.operationId))N(!get('memberAggregateHeads/'+id)&&!get('memberIdentities/'+id),'OPERATION_CONFLICT');
 }
 function select(get,r,config){if(!supports(r))return null;const id=key(r.commandType);if(!get('commandExecutorRoutes/'+id))return null;const v=route(get,config,id);const s=scope(get,v);
  // Unselected institutional resources retain their previous executor.
  if(!s.personIds.includes(r.target?.personId)&&!currentProfile(get,r,s)&&!currentBinding(get,r,s)&&!(['members.ingress.execute','members.onboarding.complete'].includes(r.commandType)&&enrollmentScope(get,r.target?.personId,s.approvalRefs))&&!(r.payload?.operation==='PRE_REGISTER'&&s.approvalRefs.includes(r.payload.approvalRef)))return null;fields(get,r,v);return id;}
 function currentProfile(get,r,s){return s.profileCorrectionCurrentMembers===true&&r.commandType==='members.profile.correct'&&f4MemberProfileTargetEligible_(get,r.target?.personId,Date.now());}
 function currentBinding(get,r,s){return s.bindingCurrentMembers===true&&r.commandType==='members.binding.associate'&&f4MemberBindingTargetEligible_(get,r.target?.personId,Date.now());}
 return {types,supports,key,scope,route,fields,select,currentProfile,currentBinding,enrollmentScope,invitationActor,resource:r=>r.target.personId};
}
