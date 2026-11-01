// Technical DEV access is a separate, protected UID record. It is never a role,
// institutional grant, membership, or portable production authorization.
function f4DevAdmin_(){
 const C=f4Core_(),N=C.need,F=f4Firestore_(),K=f4MemberAuthzContracts_();
 const project='geapa-convergencia-dev',path=uid=>'coreDevAdmins/'+uid;
 function profile(get,uid,now,required=false){
  N(C.id(uid)&&Number.isSafeInteger(now),'AUTH_BINDING_INVALID');
  const p=get(path(uid));if(!p){N(!required,'AUTHORIZATION_DENIED');return null;}
  // Runtime identity and the immutable DEV destination must agree even when a
  // copied record claims DEV. No client field can select this environment.
  f4DevBoundary_().check();
  N(f4DevBoundary_().permit({environment:'DEV',projectId:project,database:'(default)'}),'DESTINATION_INVALID');
  C.exact(p,['schemaVersion','projectId','environment','uid','active','revision','issuedAt','expiresAt','permissions','exportable']);
  N(p.schemaVersion==='core-dev-admin-v1'&&p.projectId===project&&p.environment==='DEV'&&p.uid===uid&&p.exportable===false&&p.permissions==='FUNCTIONAL_ALL'&&C.rev(p.revision)&&typeof p.active==='boolean','AUTH_BINDING_INVALID');
  const live=p.active&&C.ms(p.issuedAt)<=now&&now<C.ms(p.expiresAt);
  N(C.ms(p.issuedAt)<C.ms(p.expiresAt),'AUTH_BINDING_INVALID');
  if(!live){N(!required,'ACCESS_REVOKED');return null;}return p;
 }
 function identify(get,uid,now){
  const p=profile(get,uid,now,true),control=f4MemberAuthority_().validate(get('memberAuthorityControl/MEMBERS'));
  N(control.mode==='CANONICAL'&&control.activeWriterVersion==='M13','ACCESS_REVOKED');
  const binding=get('memberAuthBindings/'+uid);f4MemberModel_(true).validate('memberAuthBindings',binding);
  N(binding.uid===uid&&binding.active&&binding.domainVersion==='m11'&&binding.sourceRevision===binding.revision,'AUTH_BINDING_INVALID');
  const actor=f4MemberAuthz_().load(get,binding.idPessoa,control),person=actor.d['memberIdentities/'+binding.idPessoa];
  N(K.digest(get('portalUsers/'+uid))===K.digest(f4MemberRepository_().bindingView(binding,person)),'AUTH_BINDING_INVALID');
  const subject=K.subject(get('memberAuthzSubjects/'+binding.idPessoa)),state=get('memberCurrentStates/'+binding.idPessoa);
  N(state?.idPessoa===binding.idPessoa&&state.sourceRevision===actor.h.revision,'AUTH_BINDING_INVALID');
  return {profile:p,control,binding,actor,person,subject,state,derived:actor.M.derive(actor.d,binding.idPessoa,actor.h.revision,now)};
 }
 function expected(a){return {authorityEpoch:a.control.epoch,policyEpoch:a.control.policyEpoch,readAuthzEpoch:a.control.readAuthzEpoch,subjectEpoch:a.subject.authzEpoch,bindingVersion:a.binding.bindingVersion,personRevision:a.actor.h.revision,authorizationRevision:a.profile.revision};}
 function authorize(get,q,now){
  N(q.kind==='DEV_ADMIN','AUTH_BINDING_INVALID');
  f4ActivitiesAuthzDomain_().contracts().request(q);
  const a=identify(get,q.uid,now),e=expected(a);
  N(q.sourceId===q.uid&&K.digest(q.expected)===K.digest(e)&&q.target.personId===a.binding.idPessoa,'AUTH_BINDING_INVALID');
  const end=Math.min(C.ms(q.tokenExpiresAt),C.ms(q.admissionExpiresAt),C.ms(a.profile.expiresAt));N(now<end,'ACCESS_REVOKED');
  const proof={schemaVersion:'act-authz-proof-v13',kind:'DEV_ADMIN',uid:q.uid,idPessoa:a.binding.idPessoa,capability:q.capability,...e,stateRevision:a.state.sourceRevision,targetsDigest:K.digest(q.target),evidenceRevisions:[],projectId:project,profilePath:path(q.uid),profileRevision:a.profile.revision,issuedAt:F.Timestamp.fromMillis(now),expiresAt:F.Timestamp.fromMillis(end)};
  return {...a,proof,request:q,target:a.actor};
 }
 function recheck(get,q,proof,now){
  N(proof?.kind==='DEV_ADMIN'&&C.ms(proof.issuedAt)<=now&&now<C.ms(proof.expiresAt),'ACCESS_REVOKED');
  const next=authorize(get,q,now).proof,stable=p=>Object.fromEntries(Object.entries(p).filter(([k])=>!['issuedAt','expiresAt'].includes(k)));
  N(K.digest(stable(next))===K.digest(stable(proof))&&C.ms(next.expiresAt)>=C.ms(proof.expiresAt),'ACCESS_REVOKED');return next;
 }
 return {profile,identify,expected,authorize,recheck,path};
}
