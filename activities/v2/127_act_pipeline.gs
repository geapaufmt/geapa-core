// Domain producer for the shared admission/claim/transaction/receipt pipeline.
function f4ActivitiesPipeline_(){
 const C=f4Core_(),F=f4Firestore_(),M=f4ActivitiesModel_(),N=C.need,T=F.Timestamp.fromMillis,K=f4MemberCommandCatalog_();
 function runtime(options={}){return f4ActivitiesCreate_({db:{transaction:()=>C.fail('SHARED_COMMAND_PIPELINE_REQUIRED')},principals:{resolve:()=>null},...options})}
 function command(r){return {schemaVersion:'geapa-command-v2-candidate',operationId:r.operationId,commandType:r.commandType,payload:r.payload,scope:r.target.scope,expected:r.target.expected}}
 function validateCommand(q){return runtime().envelope(q)}
 function validate(r){N(r.schemaVersion==='geapa-act-command-v13','COMMAND_PAYLOAD_INVALID');C.exact(r.target,['scope','expected','authorization']);if(f4ActivitiesSystem_().is(r)){f4ActivitiesSystem_().validate(r);validateCommand(command(r));return true;}const a=r.target.authorization;C.exact(a,['kind','sourceId','functionId','actorScope']);N(['SELF','ROLE','EXPLICIT','DEV_ADMIN'].includes(a.kind)&&C.id(a.sourceId)&&(a.kind==='ROLE'?C.id(a.functionId):a.functionId===null)&&(a.kind==='DEV_ADMIN'?[null,'SELF'].includes(a.actorScope):a.kind==='SELF'?a.actorScope===null:f4MemberContracts_().scopes.includes(a.actorScope)),'AUTH_BINDING_INVALID');validateCommand(command(r));capability(r);return true}
 function capability(r){if(f4ActivitiesSystem_().is(r))return 'ACT_INTERNAL';const ct=f4ActivitiesContracts_()[r.commandType?.slice(4)];N(ct,'COMMAND_TYPE_INVALID');const self=r.target?.authorization?.kind==='SELF'||r.target?.authorization?.kind==='DEV_ADMIN'&&r.target.authorization.actorScope==='SELF';N(ct.actors.includes(self?'PORTAL_AUTHENTICATED_MEMBER':'BOARD_ACTOR'),'AUTH_BINDING_INVALID');const cap=ct.capabilities.find(c=>self?c.endsWith('_SELF'):!c.endsWith('_SELF'));N(cap,'AUTH_BINDING_INVALID');return cap}
 function authorization(get,r,identity,n,end){
  validate(r);if(f4ActivitiesSystem_().is(r))return f4ActivitiesSystem_().authorization(get,r,identity,n,end);const b=get('memberAuthBindings/'+identity.uid),control=get('memberAuthorityControl/MEMBERS'),subject=b&&get('memberAuthzSubjects/'+b.idPessoa),head=b&&get('memberAggregateHeads/'+b.idPessoa),selector=r.target.authorization;
  N(b&&control&&subject&&head,'AUTH_BINDING_INVALID');const source=selector.kind==='DEV_ADMIN'?f4DevAdmin_().profile(get,identity.uid,n,true):get((selector.kind==='EXPLICIT'?'memberExplicitGrants/':'memberAuthzPolicies/')+selector.sourceId);N(source,'AUTH_BINDING_INVALID');
  const q=command(r),refs=[];function collect(v){if(v&&typeof v==='object'){if(v.kind&&v.id){try{M.path(v);refs.push(v);return}catch{}}Object.values(v).forEach(collect)}}collect(q.payload);
  let cycle=q.payload.cycleRef;for(const ref of refs){const d=get(M.path(ref));if(d?.cycleRef?.kind&&d?.cycleRef?.id){N(!cycle||M.path(cycle)===M.path(d.cycleRef),'COMMAND_PAYLOAD_INVALID');cycle=d.cycleRef;}}
  const scopeKey=selector.kind==='SELF'||selector.kind==='DEV_ADMIN'&&selector.actorScope==='SELF'?'SELF':f4AttendanceScope_(q.commandType.slice(4))?'ACTIVITIES':q.commandType==='ACT.CYCLE_CREATE'?'PRESENTATIONS_PLAN':cycle?'PRESC_'+cycle.id:null;N(scopeKey,'AUTH_BINDING_INVALID');
  // Sealed scopes are resolved here and again by the handler; client selectors grant nothing.
  let resources=q.scope.targetRefs?.map(M.path);if(!resources){const m=get(M.path(q.scope.sealedManifestRef));N(m?.state==='SEALED'&&m.complete&&m.digest===q.scope.manifestDigest&&m.digest===M.hash(m.items),'AUTH_BINDING_INVALID');resources=m.items.map(i=>M.path(i.ref));}
  N(resources?.length&&resources.length<=32,'COMMAND_PAYLOAD_INVALID');
  return [{schemaVersion:'act-authz-request-v13',uid:identity.uid,...selector,capability:capability(r),expected:{authorityEpoch:control.epoch,policyEpoch:control.policyEpoch,readAuthzEpoch:control.readAuthzEpoch,subjectEpoch:subject.authzEpoch,bindingVersion:b.bindingVersion,personRevision:head.revision,authorizationRevision:source.revision},tokenExpiresAt:T(identity.expiresAt),admissionExpiresAt:T(end),target:{personId:b.idPessoa,scopeKey,boardId:null,roleKey:null,resources,fieldMask:[],command:q},evidenceRefs:[]}];
 }
 function preflight(get,r,requests,proofs,n){N(requests.length===1&&proofs.length===1&&K.equal(requests[0].target?.command||command(requests[0].command),command(r))&&requests[0].capability===capability(r),'AUTH_BINDING_INVALID');}
 function resolve(get,r,n){const p=get('commandProvenance/'+r.operationId),ticket=p&&get('memberAdmissionTickets/'+p.ticketId);N(p&&ticket,'AUTH_BINDING_INVALID');return {request:p.authRequests[0],proof:ticket.authProofs[0]}}
 function plan(tx,r,q,proof,clock,deadline,config){
  const get=p=>tx.get(p).data(),A=f4SharedAuthz_();preflight(get,r,[q],[proof],clock());
  function current(g,now){A.recheck(g,q,proof,now);return A.authorize(g,q,now)}
  const authority={identity:(g,uid,now)=>{N(uid===r.ownerUid,'AUTH_BINDING_INVALID');const a=current(g,now),b=a.binding,h=a.actor.h,s=g('memberCurrentStates/'+b.idPessoa),i=a.actor.d['memberIdentities/'+b.idPessoa];return {b,h,s,i,actor:{uid,personId:b.idPessoa,personRef:M.ref('person',b.idPessoa),bindingVersion:b.bindingVersion,...(q.kind==='DEV_ADMIN'?{accessProfile:'DEV_ADMIN'}:{})},authorityBasis:{...a.proof}}},ordinary:(g,uid,scope,cap,now)=>{N(uid===r.ownerUid&&scope===q.target.scopeKey&&cap===q.capability,'AUTH_BINDING_INVALID');const a=current(g,now);return {g:{entries:{[cap]:a.proof}}}},suspension:(g,base,now)=>{current(g,now);N(f4ActivitiesAuthzDomain_().suspended(g,q,base.s,base.b,now),'AUTH_BINDING_INVALID');return {kind:'MEMBROS_SUSPENSION_ONLY',personRevision:base.h.revision}}};
  const h=Object.freeze({}),principal=f4ActivitiesSystem_().is(r)?{...r.target.authorization}:{uid:r.ownerUid,kind:q.kind==='SELF'||q.kind==='DEV_ADMIN'&&q.actorScope==='SELF'?'PORTAL_AUTHENTICATED_MEMBER':'BOARD_ACTOR'};
  const v=runtime({clock,authority,executionGate:()=>N((F.inspectTransaction(tx).environment==='SYNTHETIC_LOCAL' || f4OperationalDestination_(F.inspectTransaction(tx))),'ACT_OPERATIONAL_WRITER_OFF'),principals:{resolve:p=>p===h?principal:null}}).plan(tx,h,command(r));
  const p=f4ActivitiesMutationPlan_(tx,proof.idPessoa,v.writes,v.response,clock(),deadline),rows=values=>values.map(v=>{const [resourceType,...id]=v.path.split('/');return {resourceType,id:id.join('/'),revision:v.revision}}).sort((a,b)=>{const x=a.resourceType+'/'+a.id,y=b.resourceType+'/'+b.id;return x<y?-1:x>y?1:0});
  return {plan:p,decision:{state:v.response.businessState==='REJECTED'?'FAILED':'COMPLETED',code:v.response.businessState==='REJECTED'?(v.response.error.code==='STALE_REVISION'?'REVISION_CONFLICT':'DENIED'):'OK',revisionBefore:v.response.businessState==='REJECTED'?[]:rows(p.revisionBefore),revisionAfter:v.response.businessState==='REJECTED'?[]:rows(p.revisionAfter),domainResult:v.response}};
 }
 function statusAccess(get,identity,key,n){
  const provenance=get('commandProvenance/'+key.operationId);
  N(provenance?.ownerUid===identity.uid&&provenance.operationId===key.operationId&&provenance.digest===key.digest&&provenance.authRequests?.length===1,'OPERATION_CONFLICT');
  const prior=provenance.authRequests[0];let r=get('commandRequests/'+key.operationId);
  // Provenance outlives the raw request and is retained with the protected result.
  // Reconstruct only the original intent; current sources still authorize every read.
  if(!r){const q=prior.target?.command;r=prior.command||(q&&{schemaVersion:'geapa-act-command-v13',operationId:q.operationId,commandType:q.commandType,ownerUid:identity.uid,target:{scope:q.scope,expected:q.expected,authorization:{kind:prior.kind,sourceId:prior.sourceId,functionId:prior.functionId,actorScope:prior.actorScope}},payload:q.payload,evidenceRefs:[],decisionRef:null,expectedRevisions:[],digest:key.digest});}
  N(r?.ownerUid===identity.uid&&r.digest===key.digest&&r.commandType===key.commandType&&r.operationId===key.operationId,'OPERATION_CONFLICT');
  const q=authorization(get,r,identity,n,identity.expiresAt)[0];N(q.capability===prior.capability,'AUTH_BINDING_INVALID');return f4SharedAuthz_().authorize(get,q,n);
 }
 return {command,validateCommand,validate,capability,authorization,preflight,resolve,plan,statusAccess};
}

