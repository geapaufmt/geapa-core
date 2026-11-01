// Optional, explicit command scopes on the existing institutional policy.
// These do not authorize a role: the normal live-function/capability proof is
// still required, and every resource is derived by a closed family producer.
function f4MemberOperationalScope_(get,q,limits,actorPerson,command,now){
 const K=f4MemberAuthzContracts_(),N=K.need;K.targets(limits);
 const scopes=limits.commandScopes||[],t=q.target;
 if(scopes.includes('OWN_MEMBER_RESOURCES')&&q.kind==='SELF'&&['MEMBERS_SELF_READ','MEMBERS_SELF_UPDATE'].includes(q.capability)){
  N(limits.ownPerson&&t.personId===actorPerson&&t.scopeKey===null&&t.boardId===null&&t.roleKey===null);
  if(q.capability==='MEMBERS_SELF_UPDATE')N(t.resources.length===1&&t.resources[0]==='memberPrivateProfiles/'+actorPerson||t.fieldMask.length===0&&t.resources.length===2&&t.resources.includes('memberProfileCorrectionIndexes/'+actorPerson)&&t.resources.some(p=>/^memberProfileCorrections\/[a-f0-9]{32}$/.test(p)));
  else{const head=get('memberAggregateHeads/'+actorPerson);N(head?.id===actorPerson&&Array.isArray(head.paths));N(t.resources.every(path=>path==='memberCurrentStates/'+actorPerson||head.paths.includes(path)&&/^member(Identities|Identifiers|PrivateProfiles|Functions)\//.test(path)&&get(path)?.idPessoa===actorPerson));}
  // Field masks, live membership, policy revisions and subject/binding fences
  // are still checked by the existing M13 authorization computation.
  return true;
 }
 if(q.capability==='MEMBERS_PROFILE_CORRECT'&&scopes.includes('PROFILE_CORRECTION')){N(limits.allowedPersons.includes(t.personId)||!f4ExecutionEnvironment_().production&&limits.allowedPersons.length===0);N(!limits.ownPerson&&t.scopeKey==='ADMINISTRACAO'&&limits.allowedTargetScopes.includes(t.scopeKey)&&t.purpose==='PROFILE_CORRECTION'&&t.resources.every(p=>p==='memberIdentities/'+t.personId||p==='memberPrivateProfiles/'+t.personId||p==='memberProfileCorrectionIndexes/'+t.personId||p==='memberProfileCorrections/'+t.caseId||p.startsWith('memberIdentifiers/')&&get(p)?.idPessoa===t.personId||/^memberIdentifiers\/I_[a-f0-9]{32}$/.test(p)&&!get(p)));return true;}
 if(command?.commandType==='members.ingress.execute'&&q.capability==='MEMBERS_INGRESS_EXECUTE'&&scopes.includes('ADMINISTRATIVE_INGRESS')){
  N(!limits.ownPerson&&t.scopeKey==='ADMINISTRACAO'&&limits.allowedTargetScopes.includes(t.scopeKey));
  const body=f4MemberActContent_().act(get,command,now);
  N(body.facts.schemaVersion==='member-administrative-ingress-v1'&&command.payload.operation==='RECORD_INGRESS'||body.facts.schemaVersion==='member-preregistration-v1'&&command.payload.operation==='PRE_REGISTER'||body.facts.schemaVersion==='member-invitation-action-v1'&&command.payload.operation==='MANAGE_INVITATION');
  // authorizeCommand already compared resources with the exact prospective
  // canonical changes; no arbitrary path prefix becomes a write capability.
  N(t.resources.every(p=>/^member(Identities|Identifiers|Links|Events|Enrollments|EnrollmentIndexes)\/[A-Za-z0-9_-]{1,48}$/.test(p)));
  return true;
 }
 if(['MEMBERS_VOLUNTARY_EXIT_EXECUTE','MEMBERS_VOLUNTARY_SUSPENSION_EXECUTE'].includes(q.capability)&&scopes.includes('VOLUNTARY_REVIEW')){
  N(!limits.ownPerson&&t.scopeKey==='RECURSOS_HUMANOS'&&limits.allowedTargetScopes.includes(t.scopeKey));
  const request=get('memberVoluntaryRequests/'+t.caseId);f4MemberVoluntaryWorkflow_().validate(request);
  N(request.personId===t.personId&&request.linkId===t.linkId&&t.purpose===(q.capability==='MEMBERS_VOLUNTARY_SUSPENSION_EXECUTE'?'VOLUNTARY_SUSPENSION':'VOLUNTARY_EXIT'));
  N((request.modality==='SUSPENSAO_TEMPORARIA')===(q.capability==='MEMBERS_VOLUNTARY_SUSPENSION_EXECUTE'));
  N(t.resources.every(p=>p==='memberVoluntaryRequests/'+request.id||p==='memberVoluntaryIndexes/'+t.personId||p==='memberLinks/'+t.linkId||/^memberEvents\/EVT_[a-f0-9]{32}$/.test(p)));
  return true;
 }
 K.match(t,limits,actorPerson);return true;
}
