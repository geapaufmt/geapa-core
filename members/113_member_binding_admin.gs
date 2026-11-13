// Administrative account association. No grants, Auth users or contact edits.
// The host supplies a verified Google account; the browser supplies no UID.
function f4MemberBindingAdmin_({config,identity,clock=Date.now}){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),T=F.Timestamp.fromMillis,A=f4MemberAuthz_(),K=f4MemberCommandCatalog_(),db=F.connect(config),cap='MEMBERS_BINDING_MANAGE';
 const resources=uid=>['memberAuthBindings/'+uid,'portalUsers/'+uid,'portalAccessStatus/'+uid];
 function access(get,personId,operationId,uid){
  const now=clock(),c=get('memberBindingAdminControl/current');
  N(c?.schemaVersion==='member-binding-admin-control-v1'&&c.enabled===true&&c.projectId===config.projectId&&C.rev(c.revision)&&c.revision>0&&Array.isArray(c.grantIds)&&c.grantIds.length<=16&&c.grantIds.every(C.id),'ACCESS_REVOKED');
  N(C.id(personId)&&now<identity.expiresAt,'ACCESS_REVOKED');const a=A.identify(get,identity.uid,now);
  N(get('memberIdentities/'+personId)&&get('memberAggregateHeads/'+personId),'RESOURCE_NOT_FOUND');
  for(const id of c.grantIds){const g=get('memberExplicitGrants/'+id);if(!g||g.uid!==identity.uid||!g.capabilities?.includes(cap)||!g.targets?.allowedPersons?.includes(personId)||!g.actorScopes?.includes('ADMINISTRACAO'))continue;
   const q={schemaVersion:'member-authz-request-v13',uid:identity.uid,kind:'EXPLICIT',sourceId:id,functionId:null,actorScope:'ADMINISTRACAO',capability:cap,expected:{authorityEpoch:a.control.epoch,policyEpoch:a.control.policyEpoch,readAuthzEpoch:a.control.readAuthzEpoch,subjectEpoch:a.subject.authzEpoch,bindingVersion:a.binding.bindingVersion,personRevision:a.actor.h.revision,authorizationRevision:g.revision},tokenExpiresAt:T(identity.expiresAt),admissionExpiresAt:T(Math.min(now+300000,identity.expiresAt)),target:{personId,scopeKey:'ADMINISTRACAO',boardId:null,roleKey:null,resources:uid?resources(uid):['memberIdentities/'+personId],caseId:operationId||'ACCOUNT_LOOKUP',purpose:'ACCOUNT_BINDING',fieldMask:[],linkId:null,effect:null,effectiveDate:null},evidenceRefs:[]};
   const auth=A.authorize(get,q,now);return {q,auth,control:c};
  }N(false,'ACCESS_REVOKED');
 }
 function bindings(personId){return db.query({collection:'memberAuthBindings',filters:[{field:'idPessoa',op:'EQUAL',value:personId}],order:[{field:'__name__',direction:'ASCENDING'}],limit:2}).map(v=>v.data);}
 function eligible(get,personId){const control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS')),l=A.load(get,personId,control),p=l.d['memberIdentities/'+personId],state=l.M.derive(l.d,personId,l.h.revision,clock()).state;N(p.active&&!p.technicalAccount&&!state.blocked&&['MEMBRO_EFETIVO','MEMBRO_INGRESSANTE','EGRESSO'].includes(state.state),'MEMBER_NOT_ELIGIBLE');return l;}
 function context(personId){return db.transaction(tx=>{const get=p=>tx.get(p).data();access(get,personId);const rows=bindings(personId);N(rows.length<=1,'BINDING_CONFLICT');return {canManage:true,state:rows.length?(rows[0].active?'ASSOCIATED':'REVOKED'):'UNASSOCIATED',uid:rows[0]?.uid||null};});}
 function check(get,v){
  N(v?.ownerUid===identity.uid&&v.projectId===config.projectId,'ACCESS_REVOKED');access(get,v.personId,v.operationId,v.account.uid);
  N(clock()<C.ms(v.expiresAt),'BINDING_VERIFICATION_EXPIRED');const l=eligible(get,v.personId);N(l.h.revision===v.personRevision,'REVISION_CONFLICT');
  N(!get('memberEnrollmentIndexes/'+v.personId)?.activeEnrollmentId,'BINDING_CONFLICT');
  N(!get('memberAuthBindings/'+v.account.uid)&&bindings(v.personId).length===0,'BINDING_CONFLICT');
  return v;
 }
 function verify(tx,input,account){
  C.exact(input,['operationId','personId','evidence']);C.exact(input.evidence,['kind','value','reference','identityConfirmed']);
  N(C.op(input.operationId)&&['CPF','RGA'].includes(input.evidence.kind)&&typeof input.evidence.value==='string'&&input.evidence.value.length<=40&&C.id(input.evidence.reference)&&input.evidence.identityConfirmed===true,'BINDING_EVIDENCE_REQUIRED');
  const get=p=>tx.get(p).data(),now=clock(),op=input.operationId,path='memberBindingVerifications/'+op;
  access(get,input.personId);const l=eligible(get,input.personId);
  const ids=l.M.rows(l.d,'memberIdentifiers').filter(v=>v.active&&v.kind===input.evidence.kind&&v.value===input.evidence.value);N(ids.length===1,'BINDING_EVIDENCE_REQUIRED');
  const evidence={kind:input.evidence.kind,identifierId:ids[0].id,identifierRevision:ids[0].revision,reference:input.evidence.reference,confirmedBy:identity.uid},prior=get(path);
  if(prior){N(prior.ownerUid===identity.uid&&prior.personId===input.personId&&K.equal(prior.account,account)&&K.equal(prior.evidence,evidence),'OPERATION_CONFLICT');check(get,prior);return summary(prior);}
  const v={schemaVersion:'member-binding-verification-v1',operationId:op,ownerUid:identity.uid,projectId:config.projectId,personId:input.personId,personRevision:l.h.revision,account,evidence,createdAt:T(now),expiresAt:T(Math.min(now+300000,identity.expiresAt))};
  // Authorization of the exact server-resolved UID uses this protected record.
  check(p=>p===path?v:get(p),v);tx.set(path,v);return summary(v);
 }
 function summary(v){return {operationId:v.operationId,personId:v.personId,email:v.account.email,provider:'google.com',emailVerified:true,expiresAt:C.ms(v.expiresAt),evidenceReference:v.evidence.reference};}
 function prepare(tx,input){
  C.exact(input,['operationId','personId','confirmed']);N(C.op(input.operationId)&&input.confirmed===true,'COMMAND_PAYLOAD_INVALID');const get=p=>tx.get(p).data(),op=input.operationId,v=get('memberBindingVerifications/'+op);N(v?.personId===input.personId,'ACCESS_REVOKED');
  access(get,input.personId,op,v.account.uid);const prior=get('memberPreparedIntents/'+op);if(prior){N(prior.ownerUid===identity.uid,'ACCESS_REVOKED');return prior.input;}
  check(get,v);const now=clock(),end=C.ms(v.expiresAt),r={schemaVersion:'geapa-members-command-v13',operationId:op,ownerUid:identity.uid,commandType:'members.binding.associate',target:{personId:v.personId,uid:v.account.uid},payload:{},evidenceRefs:['MVE_'+op],decisionRef:v.evidence.reference,expectedRevisions:[{resourceType:'memberAggregateHeads',id:v.personId,revision:v.personRevision},{resourceType:'memberAuthBindings',id:v.account.uid,revision:0}]};
  const q=access(get,v.personId,op,v.account.uid).q;q.admissionExpiresAt=T(end);
  const body={schemaVersion:'member-act-content-v13',...f4MemberActContent_().material(r),facts:{schemaVersion:'member-association-content-v13',uid:v.account.uid,idPessoa:v.personId,projectId:config.projectId,verifiedAt:C.ms(v.createdAt),expiresAt:end,verificationRef:op}},hash=C.hash(body),browser=Object.fromEntries(['schemaVersion','operationId','commandType','target','payload','evidenceRefs','decisionRef'].map(k=>[k,r[k]]));
  const pending={['memberPreparedIntents/'+op]:{schemaVersion:'member-prepared-intent-v1',operationId:op,ownerUid:identity.uid,inputHash:C.hash(input),input:browser,createdAt:T(now),expiresAt:T(end)},['memberEvidenceContents/'+hash]:body,['memberEvidenceAttestations/MVE_'+op]:{schemaVersion:'member-evidence-attestation-v13',id:'MVE_'+op,revision:1,active:true,kind:'CADASTRAL_EVIDENCE',holderId:v.personId,linkId:null,targetResourceIds:resources(v.account.uid),effect:'ACCOUNT_BINDING',effectiveDate:Utilities.formatDate(new Date(now),'America/Cuiaba','yyyy-MM-dd'),issuerRef:identity.uid,approvalRef:r.decisionRef,contentHash:hash,verifiedAt:T(now),expiresAt:T(end)},['devProtectedInputs/'+op]:{schemaVersion:'core-dev-protected-input-v1',operationId:op,ownerUid:identity.uid,materialHash:C.hash(Object.fromEntries(['commandType','target','payload','evidenceRefs','decisionRef'].map(k=>[k,r[k]]))),expectedRevisions:r.expectedRevisions,authRequests:[q]}};
  for(const p of Object.keys(pending))N(!get(p),'OPERATION_CONFLICT');for(const [p,d]of Object.entries(pending))tx.set(p,d);return browser;
 }
 return {access,context,verify,prepare,check,bindings};
}
