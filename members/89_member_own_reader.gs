// Small read facade over the existing Members authorization and derivation.
// The caller supplies a verified identity, never a browser-selected person or proof.
function f4MemberOwnReader_(){
 const C=f4Core_(),K=f4MemberAuthzContracts_(),A=f4MemberAuthz_(),F=f4Firestore_(),N=C.need;
 function read(get,identity,input,now){
  C.exact(identity,['uid','expiresAt']);C.exact(input,['view']);
  N(C.id(identity.uid)&&Number.isSafeInteger(identity.expiresAt)&&now<identity.expiresAt,'AUTH_BINDING_INVALID');
  N(['profile','status','functions','directors','contact'].includes(input.view),'COMMAND_PAYLOAD_INVALID');
  const technical=f4DevAdmin_().profile(get,identity.uid,now),actor=technical?f4DevAdmin_().identify(get,identity.uid,now):A.identify(get,identity.uid,now),id=actor.binding.idPessoa,control=actor.control,sp=control.selfReadPolicy;
  const functions=['functions','directors'].includes(input.view),fields=functions?f4MemberContracts_().dtoFields:input.view==='profile'?['idPessoa','displayLabel']:input.view==='status'?(actor.derived.state.state==='MEMBRO_INGRESSANTE'?['idPessoa','state']:['idPessoa','state','egressEndsAtExclusive']):['contactEmail'];
  const records=input.view==='directors'?actor.derived.currentDirectors:actor.derived.currentFunctions;
  const resources=functions?['memberCurrentStates/'+id,...records.map(f=>'memberFunctions/'+f.idVigencia)]:[(input.view==='profile'?'memberIdentities/':input.view==='contact'?'memberPrivateProfiles/':'memberCurrentStates/')+id];
  const request={schemaVersion:'member-authz-request-v13',uid:identity.uid,kind:technical?'DEV_ADMIN':'SELF',capability:'MEMBERS_SELF_READ',sourceId:technical?identity.uid:sp.policyId,functionId:null,actorScope:null,expected:{authorityEpoch:control.epoch,policyEpoch:control.policyEpoch,readAuthzEpoch:control.readAuthzEpoch,subjectEpoch:actor.subject.authzEpoch,bindingVersion:actor.binding.bindingVersion,personRevision:actor.actor.h.revision,authorizationRevision:technical?technical.revision:sp.policyRevision},tokenExpiresAt:F.Timestamp.fromMillis(identity.expiresAt),admissionExpiresAt:F.Timestamp.fromMillis(identity.expiresAt),target:{personId:id,scopeKey:null,boardId:null,roleKey:null,resources,caseId:null,purpose:null,fieldMask:fields,linkId:null,effect:null,effectiveDate:null},evidenceRefs:[]};
  const authorized=A.authorize(get,request,now);let data;
  if(input.view==='contact')data=f4MemberReadViews_().project(get,request,'SELF',now)['memberSelfProfiles/'+id];
  else if(input.view==='profile')data={schemaVersion:'safe-profile-v11',idPessoa:id,displayLabel:authorized.actor.d['memberIdentities/'+id].displayName,sourceRevision:authorized.actor.h.revision};
  else if(input.view==='status')data={schemaVersion:'safe-status-v11',idPessoa:id,state:authorized.targetDerived.state.state,egressEndsAtExclusive:authorized.targetDerived.state.egressEndsAtExclusive,sourceRevision:authorized.actor.h.revision};
  else {const projected=f4MemberPortalViews_().projectM13(actor.actor.d,id,actor.actor.h.revision,now,{boards:actor.actor.boards},control),collection=input.view==='directors'?'memberCurrentDirectors':'memberCurrentFunctions';data=records.map(f=>projected[collection+'/'+id+'/items/'+f.idVigencia]);}
  // Recheck against the caller's protected snapshot; no browser proof is accepted.
  A.recheck(get,request,authorized.proof,now);
  const expiry=Math.min(identity.expiresAt,C.ms(authorized.proof.expiresAt),C.ms(actor.derived.state.validUntil),...Object.values(actor.actor.d).filter(v=>v.startDate&&/^m1-member(Functions|Governance)/.test(v.schemaVersion)).flatMap(v=>{const i=actor.actor.M.interval(v.startDate,v.endDateReal||v.endDatePlanned||v.endDate);return [i.a,i.z].filter(t=>t>now);}));
  return {view:input.view,data,validUntil:expiry,decisionEligible:false};
 }
 function wire(value){if(value&&typeof value.toMillis==='function')return new Date(value.toMillis()).toISOString();if(Array.isArray(value))return value.map(wire);if(value&&typeof value==='object')return Object.fromEntries(Object.entries(value).map(([k,v])=>[k,wire(v)]));return value;}
 // Prefetch source dependencies in the SAME protected transaction. The reader
 // and its final authorization check still consume that transaction's snapshot.
 function readTransaction(tx,identity,input,now){
  const get=p=>tx.get(p).data(),seen=new Set();
  function batch(paths){const keys=[...new Set(paths)].filter(p=>!seen.has(p));for(let i=0;i<keys.length;i+=180){const part=keys.slice(i,i+180);tx.getAll(part);part.forEach(p=>seen.add(p));}}
  N(C.id(identity.uid),'AUTH_BINDING_INVALID');
  batch(['coreDevAdmins/'+identity.uid,'memberAuthBindings/'+identity.uid,'portalUsers/'+identity.uid,'memberAuthorityControl/MEMBERS']);
  const b=get('memberAuthBindings/'+identity.uid),control=get('memberAuthorityControl/MEMBERS');
  if(!b?.active||!C.id(b.idPessoa)||!control?.selfReadPolicy)return read(get,identity,input,now);
  N(C.id(control.selfReadPolicy.policyId),'AUTH_BINDING_INVALID');
  batch(['memberAggregateHeads/'+b.idPessoa,'memberCurrentStates/'+b.idPessoa,'memberAuthzSubjects/'+b.idPessoa,'memberAuthzPolicies/'+control.selfReadPolicy.policyId]);
  const head=get('memberAggregateHeads/'+b.idPessoa);
  N(Array.isArray(head?.paths)&&head.paths.every(p=>/^member(Identities|Identifiers|PrivateProfiles|Links|Events|Functions)\/[A-Za-z0-9_-]{1,48}$/.test(p)),'AUTH_BINDING_INVALID');
  batch(head.paths);
  const functions=head.paths.filter(p=>p.startsWith('memberFunctions/')).map(get),dependencies=[];
  for(const f of functions){N(f&&[f.idDiretoria,f.idCargo,f.idTermo].every(C.id),'AUTH_BINDING_INVALID');dependencies.push('memberBoardHeads/'+f.idDiretoria,'memberRoleCatalog/'+f.idCargo,'memberGovernanceTerms/'+f.idTermo);}
  batch(dependencies);
  const boardPaths=[];for(const id of new Set(functions.map(f=>f.idDiretoria))){const h=get('memberBoardHeads/'+id);N(Array.isArray(h?.paths)&&h.paths.every(p=>p==='memberGovernanceBoards/'+id||/^memberGovernanceBoardTerms\/[A-Za-z0-9_-]{1,48}$/.test(p)),'AUTH_BINDING_INVALID');boardPaths.push(...h.paths);}
  batch(boardPaths);
  const terms=boardPaths.filter(p=>p.startsWith('memberGovernanceBoardTerms/')).map(p=>{const r=get(p);N(r&&C.id(r.idTermo),'AUTH_BINDING_INVALID');return 'memberGovernanceTerms/'+r.idTermo;});batch(terms);
  return read(get,identity,input,now);
 }
 return Object.freeze({read,readTransaction,wire});
}
