// M63-D envelope/fingerprint boundary. Family validators/planners are trusted E producers,
// never a callback or write-list obtained from a client. No family handler is installed here.
function f4MemberCommandCatalog_(){
 const C=f4Core_(),F=f4Firestore_(),N=C.need;
 const types=Object.freeze({
  'members.self.update':['MEMBERS_SELF_UPDATE'],
  'members.profile.correct':['MEMBERS_PROFILE_CORRECT'],
  'members.ingress.execute':['MEMBERS_INGRESS_EXECUTE'],
  'members.discipline.suspend':['MEMBERS_DISCIPLINE_EXECUTE'],
  'members.discipline.dismiss':['MEMBERS_DISCIPLINE_EXECUTE'],
  'members.membership.voluntaryExit':['MEMBERS_VOLUNTARY_EXIT_EXECUTE'],
  'members.voluntary.submit':['MEMBERS_VOLUNTARY_REQUEST'],
  'members.voluntary.review':['MEMBERS_VOLUNTARY_REQUEST','MEMBERS_VOLUNTARY_EXIT_EXECUTE','MEMBERS_VOLUNTARY_SUSPENSION_EXECUTE'],
  'members.governance.appoint':['MEMBERS_GOVERNANCE_EXECUTE'],
  'members.governance.endFunction':['MEMBERS_GOVERNANCE_EXECUTE'],
  'members.governance.boardTerm':['MEMBERS_GOVERNANCE_EXECUTE'],
  'members.binding.associate':['MEMBERS_BINDING_MANAGE'],
  'members.binding.revoke':['MEMBERS_BINDING_MANAGE'],
  'members.onboarding.complete':['MEMBERS_INGRESS_EXECUTE','MEMBERS_BINDING_MANAGE']
 });
 const semantic=['schemaVersion','operationId','ownerUid','commandType','target','expectedRevisions','payload','evidenceRefs','decisionRef'];
 const hex=v=>typeof v==='string'&&/^[a-f0-9]{64}$/.test(v);
 function selector(v){C.exact(v,['operationId','digest','commandType']);N(C.op(v.operationId)&&hex(v.digest)&&Object.hasOwn(types,v.commandType),'COMMAND_PAYLOAD_INVALID');return v}
 function revisionList(rows){N(Array.isArray(rows),'SCHEMA_INVALID');let previous='';for(const row of rows){C.exact(row,['resourceType','id','revision']);N(C.id(row.resourceType)&&C.id(row.id)&&C.rev(row.revision),'SCHEMA_INVALID');const key=row.resourceType+'/'+row.id;N(key>previous,'SCHEMA_INVALID');previous=key;}return rows}
 function base(r){C.exact(r,[...semantic,'ticketId','digest','receivedAt','state']);N(r.schemaVersion==='geapa-members-command-v13'&&C.op(r.operationId)&&C.id(r.ownerUid)&&Object.hasOwn(types,r.commandType)&&C.id(r.ticketId)&&hex(r.digest)&&Number.isSafeInteger(C.ms(r.receivedAt))&&r.state==='RECEIVED','COMMAND_PAYLOAD_INVALID');revisionList(r.expectedRevisions);N(Array.isArray(r.evidenceRefs)&&r.evidenceRefs.every(C.id)&&new Set(r.evidenceRefs).size===r.evidenceRefs.length&&(r.decisionRef===null||C.id(r.decisionRef)),'COMMAND_PAYLOAD_INVALID');C.canonical(r.target);C.canonical(r.payload);return r}
 function digest(projectId,r){N(typeof projectId==='string'&&(/^demo-geapa-m(04|13)-local$/.test(projectId)||(projectId===f4ExecutionEnvironment_().project&&f4ExecutionEnvironment_().destination(f4ExecutionConfig_()))),'DESTINATION_INVALID');return C.hash({projectId,...Object.fromEntries(semantic.map(k=>[k,r[k]]))})}
 function envelope(projectId,r,validateFamily){base(r);N(typeof validateFamily==='function'&&validateFamily(r)===true,'COMMAND_PAYLOAD_INVALID');N(r.digest===digest(projectId,r),'OPERATION_CONFLICT');return r}
 const template=r=>Object.fromEntries([...semantic,'digest'].map(k=>[k,r[k]]));
 const equal=(a,b)=>C.hash(F.encode(a))===C.hash(F.encode(b));
 return {types,selector,revisionList,base,digest,envelope,template,equal,hex};
}

// E's exact local producer registry. Registration grants no authority and enables no endpoint.
function f4MemberFamilyCatalog_(config){
 const C=f4Core_();C.destination(config);C.need((config.environment==='SYNTHETIC_LOCAL' || f4OperationalDestination_(config)),'DESTINATION_INVALID');
 const modules={
  'members.self.update':f4MemberProfileCommands_(),
  'members.profile.correct':f4MemberProfileCommands_(),
  'members.ingress.execute':f4MemberIngressCommands_(),
  'members.discipline.suspend':f4MemberMembershipCommands_(),
  'members.discipline.dismiss':f4MemberMembershipCommands_(),
  'members.membership.voluntaryExit':f4MemberMembershipCommands_(),
  'members.voluntary.submit':f4MemberVoluntaryCommands_(),
  'members.voluntary.review':f4MemberVoluntaryCommands_(),
  'members.governance.appoint':f4MemberGovernanceCommands_(),
  'members.governance.endFunction':f4MemberGovernanceCommands_(),
  'members.governance.boardTerm':f4MemberBoardCommands_(),
  'members.binding.associate':f4MemberBindingCommands_(config),
  'members.binding.revoke':f4MemberBindingCommands_(config),
  'members.onboarding.complete':f4MemberOnboarding_()
 };C.exact(modules,Object.keys(f4MemberCommandCatalog_().types));return Object.freeze(Object.fromEntries(Object.entries(modules).map(([type,m])=>[type,Object.freeze({validate:m.validate,resolve:m.resolve})])));
}

// M621 family audit events are technical records, never additions to memberEvents.
function f4MemberFamilyAudit_(request,actorPersonId,now){
 const C=f4Core_(),F=f4Firestore_(),events={
  'members.self.update':'SELF_UPDATED','members.profile.correct':'PROFILE_CORRECTED',
  'members.ingress.execute':'INGRESS_EXECUTED','members.discipline.suspend':'SUSPENSAO',
  'members.discipline.dismiss':'DESLIGAMENTO','members.membership.voluntaryExit':'DESLIGAMENTO',
  'members.voluntary.submit':'VOLUNTARY_REQUESTED','members.voluntary.review':'VOLUNTARY_REVIEWED',
  'members.governance.appoint':'FUNCTION_APPOINTED','members.governance.endFunction':'FUNCTION_ENDED',
  'members.governance.boardTerm':'GOVERNANCE_CHANGED','members.binding.associate':'BINDING_ASSOCIATED',
  'members.binding.revoke':'BINDING_REVOKED'};
 const event=f4MemberProfileCorrections_().is(request)?'PROFILE_CORRECTION_'+request.payload.action:f4MemberEnrollmentCommands_().is(request)?'ENROLLMENT_'+request.payload.action:request.commandType==='members.ingress.execute'&&request.payload.operation==='PRE_REGISTER'?'MEMBER_PRE_REGISTERED':request.commandType==='members.ingress.execute'&&request.payload.operation==='MANAGE_INVITATION'?'INVITATION_'+request.payload.action:events[request.commandType];C.need(event&&C.id(actorPersonId),'COMMAND_TYPE_INVALID');
 return {path:'memberAudit/'+C.hash({operationId:request.operationId,event}).slice(0,32),kind:'set',data:{schemaVersion:'member-family-audit-v13',event,commandType:request.commandType,operationId:request.operationId,ownerUid:request.ownerUid,actorPersonId,nature:request.commandType==='members.membership.voluntaryExit'?'VOLUNTARY':null,evidenceRefs:request.evidenceRefs,decisionRef:request.decisionRef,targetDigest:C.hash(request.target),createdAt:F.Timestamp.fromMillis(now)}};
}
