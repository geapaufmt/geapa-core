'use strict';
// Operator-only pure plan. The caller must re-read every precondition in one
// transaction, rerun this plan, record an audit and apply its exact writes.
// No SDK connection, role assignment, admission approval or notification send.
const a=require('node:assert/strict'),crypto=require('node:crypto');
const {createDomain}=require('../src/domain.cjs'),{encode}=require('../src/firestore.cjs'),{profile}=require('../src/runtime-profile.cjs');
const digest=v=>crypto.createHash('sha256').update(JSON.stringify(encode(v??null))).digest('hex');
const ms=v=>v?.toMillis?v.toMillis():v instanceof Date?v.getTime():NaN;
function plan({docs,uid,projectId,sourcePolicyId,authorizationRef,policyId='RH_PREREGISTRATION_DRAFTS',scopeId='RH_PREREGISTRATION_DRAFTS',now=Date.now()}){
 a.equal(projectId,profile.projectId,'PROJECT_MISMATCH');a(Number.isSafeInteger(now));
 for(const id of [uid,sourcePolicyId,authorizationRef,policyId,scopeId])a.match(id,/^[A-Za-z0-9_-]{1,48}$/);
 a.notEqual(sourcePolicyId,policyId,'SOURCE_POLICY_IMMUTABLE');
 const observed=new Map(),get=path=>{const value=docs[path]??null;observed.set(path,{path,exists:value!==null,digest:digest(value)});return value;};
 const domain=createDomain(raw=>{const q=JSON.parse(raw);a.equal(q.kind,'get','READ_ONLY_DOMAIN');return JSON.stringify(encode(get(q.path)));});
 const run=(fn,v)=>JSON.parse(domain.run(fn,v));
 const actor=run('(v)=>{const a=f4MemberAuthz_().identify(nodeRead,v.uid,v.now);return {personId:a.binding.idPessoa,bindingVersion:a.binding.bindingVersion,subjectEpoch:a.subject.authzEpoch,personRevision:a.actor.h.revision,functions:a.derived.currentFunctions};}',{uid,now});
 const rh=actor.functions.filter(f=>f.roleKey==='DIRETOR_RECURSOS_HUMANOS'&&f.scopeKey==='RECURSOS_HUMANOS'&&f.group==='DIRETORIA');a.equal(rh.length,1,'LIVE_RH_ROLE_REQUIRED');
 const source=get('memberAuthzPolicies/'+sourcePolicyId);
 run('(v)=>f4MemberAuthzContracts_().policy(nodeRead(v))', 'memberAuthzPolicies/'+sourcePolicyId);
 a(source.active&&source.proofKind==='ROLE'&&source.actorSelector.roleKey===rh[0].roleKey&&source.actorSelector.actorScope===rh[0].scopeKey,'RH_POLICY_REQUIRED');
 a(ms(source.issuedAt)<=now&&now<ms(source.expiresAt),'RH_POLICY_EXPIRED');
 const portal=get('memberOperationalConfig/portal'),control=get('memberExecutorControl/current'),registration=get('memberOperationalConfig/registration'),notification=get('memberOperationalConfig/notifications');
 a(portal?.schemaVersion==='member-operational-portal-v1'&&portal.enabled&&Array.isArray(portal.commands)&&Array.isArray(portal.policyIds));
 a(control?.schemaVersion==='member-executor-control-v1'&&control.enabled&&control.projectId===projectId&&control.browserEnabled&&control.appCheck?.mode==='REQUIRED');
 a(control.actorUids.includes(uid),'ACTOR_NOT_IN_EXISTING_CONTROL');
 a(Array.isArray(control.commandTypes)&&Array.isArray(control.portalRoutes)&&Array.isArray(control.scopeIds),'EXPLICIT_CONTROL_FILTERS_REQUIRED');
 a(registration?.schemaVersion==='member-registration-policy-v1'&&registration.enabled&&registration.stagedEnrollment,'REGISTRATION_DISABLED');
 a(notification?.enabled&&notification.projectId===projectId&&notification.environment===profile.environment&&notification.mode==='DRY_RUN','DRY_RUN_REQUIRED');
 const policyPath='memberAuthzPolicies/'+policyId,scopePath='memberExecutorScopes/'+scopeId;
 a.equal(get(policyPath),null,'POLICY_ALREADY_EXISTS');a.equal(get(scopePath),null,'SCOPE_ALREADY_EXISTS');
 const policy={schemaVersion:'member-authz-policy-v13',id:policyId,revision:1,active:true,origin:source.origin,approvalRef:authorizationRef,issuedAt:new Date(now),expiresAt:source.expiresAt,proofKind:'ROLE',capabilities:['MEMBERS_INGRESS_EXECUTE'],actorSelector:{...source.actorSelector},targets:{ownPerson:false,allowedTargetScopes:['ADMINISTRACAO'],allowedBoards:[],allowedRoles:[],allowedPersons:[],allowedResources:[],commandScopes:['ADMINISTRATIVE_INGRESS']},fieldMasks:{MEMBERS_INGRESS_EXECUTE:[]}};
 const scope={schemaVersion:'member-executor-scope-v1',id:scopeId,projectId,synthetic:profile.environment==='DEV',revision:1,personIds:[],approvalRefs:[],preregistrationDrafts:true};
 const writes={[policyPath]:policy,[scopePath]:scope};
 const types=['members.ingress.execute','members.onboarding.complete'];
 for(const commandType of types){const routeId='MEMBERS_'+commandType.slice(8).replaceAll('.','_').toUpperCase(),path='commandExecutorRoutes/'+routeId;a.equal(get(path),null,'ROUTE_ALREADY_EXISTS');
  writes[path]={schemaVersion:'member-executor-route-v1',projectId,routeId,commandType,scopeId,syntheticOnly:profile.environment==='DEV',executor:profile.executor,state:'ACTIVE',epoch:1,revision:1,actorUids:[uid],operationIds:[],mode:profile.environment+'_SCOPED_OPERATIONAL',availabilityMode:'PERMANENT',startsAt:now,expiresAt:null,maxOperations:null,rateLimit:{windowMs:3600000,maxAdmissions:20,maxPending:20},rateUsage:{windowStart:now,count:0}};
 }
 const union=(old,extra)=>[...new Set([...old,...extra])];
 writes['memberOperationalConfig/portal']={...portal,revision:portal.revision+1,commands:union(portal.commands,['members.ingress.execute']),policyIds:union(portal.policyIds,[policyId])};
 writes['memberExecutorControl/current']={...control,revision:control.revision+1,scopeIds:union(control.scopeIds,[scopeId]),commandTypes:union(control.commandTypes,types),portalRoutes:union(control.portalRoutes,['/admin/ingressos-membros/cadastrar','/admin/ingressos-membros/convite'])};
 a(writes['memberOperationalConfig/portal'].policyIds.length<=16);a(writes['memberExecutorControl/current'].scopeIds.length<=4);
 // Validate the planned artifacts with the same canonical contracts as runtime.
 const check=createDomain(raw=>{const q=JSON.parse(raw);a.equal(q.kind,'get');return JSON.stringify(encode(Object.hasOwn(writes,q.path)?writes[q.path]:get(q.path)));});
 check.run('(v)=>{f4MemberAuthzContracts_().policy(nodeRead(v.policy));const f=f4MemberExecutor_();for(const id of v.routes)f.route(nodeRead,f4ExecutionConfig_(),id);return true;}',{policy:policyPath,routes:['MEMBERS_INGRESS_EXECUTE','MEMBERS_ONBOARDING_COMPLETE']});
 return {projectId,environment:profile.environment,actor:{uid,personId:actor.personId,functionId:rh[0].idVigencia,bindingVersion:actor.bindingVersion,subjectEpoch:actor.subjectEpoch,personRevision:actor.personRevision},expiresAt:ms(source.expiresAt),preconditions:[...observed.values()],writes,notificationMode:'DRY_RUN',individualApprovalsCreated:false};
}
function planDevDraft({docs,uid,personId,registration,expectedGrantRevision,close=false,now=Date.now()}){
 a.equal(profile.environment,'DEV','DEV_ONLY');
 const observed=new Map(),get=path=>{const value=docs[path]??null;observed.set(path,{path,exists:value!==null,digest:digest(value)});return value;};
 const domain=createDomain(raw=>{const q=JSON.parse(raw);if(q.kind==='guard')return 'null';a.equal(q.kind,'get');return JSON.stringify(encode(get(q.path)));}),gp='coreDevAdministrativeGrants/'+uid,g=get(gp);
 a.equal(g?.revision,expectedGrantRevision,'GRANT_CHANGED');a.equal(g.uid,uid);a.equal(g.projectId,profile.projectId);a.equal(g.environment,'DEV');a.equal(g.exportable,false);
 a(g.syntheticPersonIds.includes(personId));
 if(close){a.equal(g.preregistrationWindow?.personId,personId,'WINDOW_CHANGED');}
 else{
  a(!g.preregistrationWindow,'WINDOW_ALREADY_EXISTS');
  domain.run('(v)=>{const D=f4DevAdministrativeAccess_(),a=D.source(nodeRead,{uid:v.uid,expiresAt:v.now+3600000},v.now,"MEMBERS_INGRESS_EXECUTE"),C=f4Core_(),E=f4MemberEnrollmentContract_();C.need(a&&a.mask.includes("preRegister"),"AUTHORIZATION_DENIED");const r=E.quick(v.registration),p=nodeRead("memberIdentities/"+v.personId),resolved=f4MemberRegistrationIdentity_().resolve(nodeRead,{email:r.email,rga:r.rga});C.need(r.email.endsWith("@example.invalid")&&/^999/.test(r.rga)&&p?.origin==="DEV_ADMIN_SYNTHETIC_V1"&&p.technicalAccount===false&&resolved.personId===v.personId,"AUTHORIZATION_DENIED");const loaded=f4MemberAuthz_().load(nodeRead,v.personId,a.actor.control),state=loaded.M.derive(loaded.d,v.personId,loaded.h.revision,v.now).state;C.need(state.state==="SEM_VINCULO"&&!nodeRead("memberEnrollmentIndexes/"+v.personId)?.activeEnrollmentId,"DOMAIN_STATE_INVALID");return true;}',{uid,personId,registration,now});
 }
 const {preregistrationWindow:old,...base}=g,end=Math.min(now+3600000,ms(g.expiresAt));if(!close)a(end>now);
 const next={...base,revision:g.revision+1,...(close?{}:{preregistrationWindow:{personId,registration,issuedAt:new Date(now),expiresAt:new Date(end)}})};
 return {projectId:profile.projectId,environment:'DEV',action:close?'CLOSE':'OPEN',preconditions:[...observed.values()],writes:{[gp]:next},expiresAt:close?now:end,individualApprovalsCreated:false};
}
function draftFixture(uid,now=Date.now()){
 a.equal(profile.environment,'DEV','DEV_ONLY');a.match(uid,/^[A-Za-z0-9_-]{1,48}$/);a(Number.isSafeInteger(now));
 const hash=v=>crypto.createHash('sha256').update(JSON.stringify(v)).digest('hex'),suffix=hash(uid).slice(0,12),personId='DEVADM_'+suffix+'_4',origin='DEV_ADMIN_SYNTHETIC_V1';
 const registration={fullName:'Pessoa sintética de pré-cadastro independente',email:'devadm.draft.'+suffix+'@example.invalid',rga:'9994'+BigInt('0x'+suffix).toString(),phone:'+5566999990000'};
 const identity='memberIdentities/'+personId,docs={[identity]:{schemaVersion:'m1-memberIdentities-v1',id:personId,idPessoa:personId,revision:1,origin,displayName:registration.fullName,technicalAccount:false,active:true}},paths=[identity];
 for(const kind of ['EMAIL','RGA']){const value=registration[kind.toLowerCase()],id=personId+'_'+kind,path='memberIdentifiers/'+id;
  docs[path]={schemaVersion:'m1-memberIdentifiers-v1',id,idPessoa:personId,revision:1,origin,kind,value,active:true};paths.push(path);
  docs['memberIdentifierClaims/'+hash({kind,value})]={idPessoa:personId,identifierId:id,active:true};
 }
 docs['memberAggregateHeads/'+personId]={schemaVersion:'m11-head-v1',id:personId,kind:'PERSON',revision:1,paths};
 docs['memberAuthzSubjects/'+personId]={schemaVersion:'member-authz-subject-v13',idPessoa:personId,revision:1,authzEpoch:1,active:true,approvalRef:origin,updatedAt:new Date(now)};
 return {personId,registration,docs};
}
function planDevAdmissionTest({docs,uid,enrollmentId,expectedGrantRevision,close=false,now=Date.now()}){
 a.equal(profile.environment,'DEV','DEV_ONLY');
 const observed=new Map(),get=path=>{const value=docs[path]??null;observed.set(path,{path,exists:value!==null,digest:digest(value)});return value;};
 const gp='coreDevAdministrativeGrants/'+uid,g=get(gp),e=get('memberEnrollments/'+enrollmentId);
 a.equal(g?.revision,expectedGrantRevision,'GRANT_CHANGED');a.equal(g.uid,uid);a.equal(g.projectId,profile.projectId);a.equal(g.environment,'DEV');a.equal(g.exportable,false);
 a(e?.createdBy===uid&&e.draft===true&&g.syntheticPersonIds.includes(e.personId),'OUTSIDE_SYNTHETIC_GRANT');
 const {admissionTestWindow:old,...base}=g,end=Math.min(now+3600000,ms(g.expiresAt));
 if(close)a.equal(old?.enrollmentId,enrollmentId,'WINDOW_CHANGED');
 else {a(!old,'WINDOW_ALREADY_EXISTS');a(e.approval===null&&['PRE_REGISTERED','INVITED','ACCEPTED'].includes(e.status));a(end>now);}
 const approvedOn=new Intl.DateTimeFormat('en-CA',{timeZone:'America/Cuiaba',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(now));
 const next={...base,revision:g.revision+1,...(close?{}:{admissionTestWindow:{enrollmentId,personId:e.personId,registration:e.registration,membershipType:'MEMBRO_INGRESSANTE',approvedOn,reason:'TESTE_DEV: confirmação sintética explícita de admissão',issuedAt:new Date(now),expiresAt:new Date(end)}})};
 if(!close){const domain=createDomain(raw=>{const q=JSON.parse(raw);if(q.kind==='guard')return 'null';a.equal(q.kind,'get');return JSON.stringify(encode(q.path===gp?next:get(q.path)));});domain.run('(v)=>{const D=f4DevAdministrativeAccess_(),a=D.source(nodeRead,{uid:v.uid,expiresAt:v.now+60000},v.now,"MEMBERS_INGRESS_EXECUTE"),e=f4MemberEnrollmentContract_().validate(nodeRead("memberEnrollments/"+v.enrollmentId));f4Core_().need(a&&D.decisionAllowed(nodeRead,a,e,v.now),"AUTHORIZATION_DENIED");return true;}',{uid,enrollmentId,now});}
 return {projectId:profile.projectId,environment:'DEV',action:close?'CLOSE':'OPEN',preconditions:[...observed.values()],writes:{[gp]:next},expiresAt:close?now:end,individualApprovalsCreated:false};
}
if(require.main===module)throw Error('PURE_PLAN_ONLY_USE_AUTHORIZED_TRANSACTION_WRAPPER');
module.exports={plan,planDevDraft,planDevAdmissionTest,draftFixture,digest};
