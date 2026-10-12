// Composition is local and inert. The destination and verifier are trusted server ports.
function f4CorePipeline_(options){
 const C=f4Core_(),N=C.need;C.exact(options,['config','clock','verifyIdentity','membersAuthorization','membersRevisions']);C.destination(options.config);N((options.config.environment==='SYNTHETIC_LOCAL' || f4DevBoundary_().permit(options.config)),'DESTINATION_INVALID');
 const {config,clock}=options,internal=new WeakMap(),verifyIdentity=token=>{if(token&&typeof token==='object'){const i=internal.get(token);N(i&&clock()<i.expiresAt,'AUTH_BINDING_INVALID');return i;}const v=options.verifyIdentity(token);N(v?.uid!=='CORE_SYSTEM','AUTH_BINDING_INVALID');return v},A=f4ActivitiesPipeline_(),K=f4SharedCommandCatalog_();
 const admission=f4MemberAdmission_().createLocal({config,clock,verifyIdentity,resolveAuthorization:(get,r,a,n,end)=>K.act(r.commandType)?A.authorization(get,r,a,n,end):options.membersAuthorization(get,r,a,n,end),resolveRevisions:(get,r,a,n)=>K.act(r.commandType)?[]:options.membersRevisions(get,r,a,n)});
 const recovery=f4MemberCommandRecovery_().createLocal({config,clock,verifyIdentity,producers:f4SharedFamilyCatalog_(config)});
 const endpoint=f4MemberEndpoint_().createLocal({config,scriptId:config.environment==='SYNTHETIC_LOCAL'?'LOCAL_M13_G':f4DevBoundary_().script,admission,recovery,consumers:{status:(tx,p,key)=>K.act(key.commandType)?recovery.readStatus(tx,p,key):f4MemberConsumers_().createLocal({config,recovery}).status(tx,p,key)}});
 function system(input){
  N(f4ActivitiesSystem_().is(input),'COMMAND_TYPE_INVALID');const token=Object.freeze({});internal.set(token,{uid:'CORE_SYSTEM',expiresAt:clock()+120000});
  const principal=admission.authenticate(token),db=f4Firestore_().connect(config);
  const semantic={...input,ownerUid:'CORE_SYSTEM',expectedRevisions:[]};A.validate(semantic);
  const key={operationId:input.operationId,commandType:input.commandType,digest:K.digest(config.projectId,semantic)};
  const existing=db.transaction(tx=>{const get=p=>tx.get(p).data(),prior=f4MemberLedgerM13_().recover(get,key,'CORE_SYSTEM',clock()),r=get('commandRequests/'+key.operationId);if(r)N(r.ownerUid==='CORE_SYSTEM'&&r.digest===key.digest&&r.commandType===key.commandType,'OPERATION_CONFLICT');return !!(prior||r)});
  if(!existing){
  const issued=db.transaction(tx=>admission.issue(tx,principal,semantic));
  const r={...issued.requestEnvelope,ticketId:issued.ticketId,receivedAt:f4Firestore_().Timestamp.fromMillis(clock()),state:'RECEIVED'};
  db.transaction(tx=>{N(!tx.get('commandRequests/'+r.operationId).data()&&!tx.get('memberAdmissionUses/'+r.ticketId).data(),'OPERATION_CONFLICT');tx.set('commandRequests/'+r.operationId,r);tx.set('memberAdmissionUses/'+r.ticketId,{schemaVersion:'member-admission-use-v13',ticketId:r.ticketId,operationId:r.operationId,ownerUid:r.ownerUid,reservedAt:r.receivedAt});});
  }
  return Object.freeze({key,attempt:worker=>endpoint.attempt(token,key,worker),status:()=>db.transaction(tx=>recovery.readStatus(tx,recovery.authenticate(token),key))});
 }
 return Object.freeze({admission,recovery,endpoint,system});
}

