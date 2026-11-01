// Candidate entry uses the same authenticated Core transport and command
// pipeline, without granting access to institutional member screens.
function f4MemberEnrollmentPortal_({config,identity,clock=Date.now}){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),K=f4MemberCommandCatalog_(),A=f4MemberEnrollmentAccess_(),E=f4MemberEnrollmentContract_(),D=f4MemberEnrollmentCommands_(),db=F.connect(config),T=F.Timestamp.fromMillis;
 const material=r=>Object.fromEntries(['commandType','target','payload','evidenceRefs','decisionRef'].map(k=>[k,r[k]]));
 function status(verified){
  C.exact(verified,['uid','expiresAt','projectId','email','emailVerified']);
  N(verified.uid===identity.uid&&verified.emailVerified===true&&verified.projectId===config.projectId&&f4ExecutionEnvironment_().destination(config)&&clock()<verified.expiresAt,'AUTH_BINDING_INVALID');
  return db.transaction(tx=>{const get=p=>tx.get(p).data(),s=get('memberEnrollmentSessions/'+identity.uid);N(s?.schemaVersion==='member-enrollment-session-v1'&&s.uid===identity.uid&&s.environment===config.environment&&s.projectId===config.projectId,'AUTH_BINDING_INVALID');
   const e=E.validate(get('memberEnrollments/'+s.enrollmentId)),claim=get('memberIdentifierClaims/'+C.hash({kind:'EMAIL',value:verified.email})),identifier=claim&&get('memberIdentifiers/'+claim.identifierId),binding=get('memberAuthBindings/'+identity.uid);
   N(claim?.active===true&&claim.idPessoa===s.personId&&identifier?.active===true&&identifier.kind==='EMAIL'&&identifier.idPessoa===s.personId&&identifier.value===verified.email&&e.personId===s.personId&&e.registration.email===verified.email&&(e.uid===null||e.uid===identity.uid),'AUTH_BINDING_INVALID');
   N(!binding||binding.active===true&&binding.idPessoa===s.personId,'ACCESS_REVOKED');
   const terminal=['DECLINED','CANCELLED','COMPLETED'].includes(e.status),expired=!terminal&&clock()>=e.expiresAt;N(terminal||expired,'ACCESS_REVOKED');
   // Read-only recovery discloses no CPF, profile, normative receipts or grants.
   // It does not renew a candidate session or authorize another command.
   return {ok:true,data:{terminal:true,status:expired?'EXPIRED':e.status,enrollmentId:e.id,revision:e.revision,registration:e.registration,expiresAt:e.expiresAt}};
  });
 }
 function context(input){
  C.exact(input,[]);
  return db.transaction(tx=>{
   const get=p=>tx.get(p).data(),now=clock(),s=A.source(get,identity.uid,now),v=s.enrollment;
   N(['INVITED','ACCEPTED'].includes(v.status)&&now<v.expiresAt&&get('memberEnrollmentIndexes/'+v.personId)?.activeEnrollmentId===v.id,'ACCESS_REVOKED');
   const c=D.catalog(get,now),today=Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'),complete=E.completeness(get,v.profile,today),ids=s.policy.courseIds;
   N(Array.isArray(ids)&&ids.length<=100&&ids.every(C.id),'SCHEMA_INVALID');
   const courses=ids.map(id=>get('memberCourseCatalog/'+id)).filter(v=>v?.active===true).map(v=>({id:v.id,name:v.name}));
   const axisIds=s.policy.interestAxisIds||[];N(Array.isArray(axisIds)&&axisIds.length<=100&&axisIds.every(C.id),'SCHEMA_INVALID');
   const interestAxes=axisIds.map(id=>get('presentationPolicyReferences/'+id)).filter(v=>v?.active===true&&v.approved===true).map(v=>({id:v.id,name:v.label||v.name||v.id}));
   // Identity, CPF and draft data are returned only to the verified invitee.
   return {ok:true,data:{enrollmentId:v.id,revision:v.revision,status:v.status,registration:v.registration,profile:v.profile,expiresAt:v.expiresAt,missing:complete.missing,courses,interestAxes,norms:c.norms.map(n=>({id:n.id,version:n.version,label:n.label||n.title||n.id,pdfUrl:n.pdfUrl})),normTermVersion:c.configuration.normTermVersion,acknowledgement:E.acknowledgement,normAcknowledgement:v.normAcknowledgement,imageTerm:c.configuration.imageTerm||null,imageManifestation:v.imageManifestation}};
  });
 }
 function prepare(input){
  C.exact(input,['operationId','revision','action','payload']);N(C.op(input.operationId)&&C.rev(input.revision)&&input.revision>0,'COMMAND_PAYLOAD_INVALID');
  return db.transaction(tx=>{
   const get=p=>tx.get(p).data(),now=clock(),s=A.source(get,identity.uid,now),v=s.enrollment,path='memberPreparedIntents/'+input.operationId,hash=C.hash(input),old=get(path);
   if(old){N(old.ownerUid===identity.uid&&old.inputHash===hash,'OPERATION_CONFLICT');return old.input;}
   N(v.revision===input.revision,'REVISION_CONFLICT');
   N(input.payload&&Object.getPrototypeOf(input.payload)===Object.prototype&&!Object.hasOwn(input.payload,'operation')&&!Object.hasOwn(input.payload,'action'),'COMMAND_PAYLOAD_INVALID');
   const r={schemaVersion:'geapa-members-command-v13',operationId:input.operationId,ownerUid:identity.uid,commandType:'members.onboarding.complete',target:{personId:v.personId,enrollmentId:v.id},payload:{operation:'ENROLLMENT',action:input.action,...input.payload},evidenceRefs:[],decisionRef:null};D.validate(r);
   r.expectedRevisions=D.paths(get,r,now).map(path=>({resourceType:path.split('/')[0],id:path.split('/')[1],revision:get(path)?.revision||0}));
   const end=Math.min(now+3600000,identity.expiresAt,C.ms(s.session.expiresAt),v.expiresAt),q=A.request(get,identity,r,now,end);A.authorize(get,q,now);D.proposal(get,r,now);
   const browser=Object.fromEntries(['schemaVersion','operationId','commandType','target','payload','evidenceRefs','decisionRef'].map(k=>[k,r[k]])),protectedPath='devProtectedInputs/'+r.operationId;N(!get(protectedPath),'OPERATION_CONFLICT');
   tx.set(path,{schemaVersion:'member-prepared-intent-v1',operationId:r.operationId,ownerUid:identity.uid,inputHash:hash,input:browser,createdAt:T(now),expiresAt:T(end),operationalChecks:null});
   tx.set(protectedPath,{schemaVersion:'core-dev-protected-input-v1',operationId:r.operationId,ownerUid:identity.uid,materialHash:C.hash(material(r)),expectedRevisions:r.expectedRevisions,authRequests:[q]});return browser;
  });
 }
 function resume(key){
  K.selector(key);N(key.commandType==='members.onboarding.complete','COMMAND_TYPE_INVALID');
  return db.transaction(tx=>{const get=p=>tx.get(p).data(),v=get('memberPreparedIntents/'+key.operationId),p=get('devProtectedInputs/'+key.operationId);N(v?.ownerUid===identity.uid&&p?.ownerUid===identity.uid&&D.is(v.input),'ACCESS_REVOKED');N(K.digest(config.projectId,{...v.input,ownerUid:identity.uid,expectedRevisions:p.expectedRevisions})===key.digest,'OPERATION_CONFLICT');
   if(get('commandRequests/'+key.operationId))A.statusAccess(get,identity,key,clock());else{const s=A.source(get,identity.uid,clock());N(v.input.target.personId===s.enrollment.personId&&v.input.target.enrollmentId===s.enrollment.id,'ACCESS_REVOKED');}return v.input;});
 }
 return {context,prepare,resume,status};
}
