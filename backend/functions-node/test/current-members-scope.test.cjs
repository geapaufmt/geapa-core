'use strict';
Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GEAPA_PILOT_MODE:'SHARED_DEV',GCLOUD_PROJECT:'demo-geapa-functions-local',FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
const {test}=require('node:test'),a=require('node:assert/strict'),{setup}=require('./people-fixture.cjs'),E=require('../src/environment.cjs');
function fixture(role='DIRETOR_ADMINISTRATIVO'){
 const h=setup(role,5),p=h.policy;p.capabilities=['MEMBERS_PROFILE_CORRECT'];p.fieldMasks={MEMBERS_PROFILE_CORRECT:['displayName']};p.targets={...p.targets,allowedTargetScopes:['ADMINISTRACAO'],allowedPersons:[],allowedResources:[],commandScopes:['PROFILE_CORRECTION','PROFILE_CORRECTION_CURRENT_MEMBERS']};
 h.docs['memberOperationalConfig/portal'].commands=['members.profile.correct'];
 const actor=h.docs['memberAuthBindings/CF_ADMIN'],s=h.docs['memberAuthzSubjects/'+actor.idPessoa];
 h.docs['memberExplicitGrants/SYN_DYNAMIC_BINDING']={schemaVersion:'member-explicit-grant-v13',id:'SYN_DYNAMIC_BINDING',revision:1,active:true,origin:'SYN_TEST',approvalRef:'SYN_APPROVED',issuerRef:'SYN_AUTHORITY',uid:'CF_ADMIN',idPessoa:actor.idPessoa,bindingVersion:actor.bindingVersion,subjectEpoch:s.authzEpoch,issuedAt:p.issuedAt,expiresAt:p.expiresAt,capabilities:['MEMBERS_BINDING_MANAGE'],actorScopes:['ADMINISTRACAO'],targets:{ownPerson:false,allowedTargetScopes:['ADMINISTRACAO'],allowedBoards:[],allowedRoles:[],allowedPersons:[],allowedResources:[],commandScopes:['ACCOUNT_BINDING','ACCOUNT_BINDING_CURRENT_MEMBERS']},fieldMasks:{MEMBERS_BINDING_MANAGE:[]},environment:E.profile.environment,runId:null};
 h.docs['memberBindingAdminControl/current']={schemaVersion:'member-binding-admin-control-v1',projectId:E.PROJECT,revision:1,enabled:true,grantIds:['SYN_DYNAMIC_BINDING']};
 h.docs['memberExecutorScopes/SYN_DYNAMIC']={schemaVersion:'member-executor-scope-v1',id:'SYN_DYNAMIC',projectId:E.PROJECT,synthetic:!E.PRODUCTION,revision:1,personIds:[],approvalRefs:[],profileCorrectionCurrentMembers:true,bindingCurrentMembers:true};
 require('./production-fixture-profile.cjs')(h.docs);
 h.run=(code,input={})=>JSON.parse(h.domain.run('(x)=>{const get=p=>f4Firestore_().connect(nodeConfig).get(p);'+code+'}',{now:h.now,...input}));return h;
}
test('existing current-member policies validate without poisoning operational context',()=>{
 for(const role of ['DIRETOR_ADMINISTRATIVO','DIRETOR_RECURSOS_HUMANOS']){const h=fixture(role);a.equal(h.run('f4MemberAuthzContracts_().policy(get("memberAuthzPolicies/PEOPLE_POLICY"));return true;'),true);a.equal(h.run('return !!f4MemberOperationalReader_({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:x.now+60000},clock:()=>x.now}).access(get,x.now).own;'),true);}
});
test('dynamic correction scope cannot be supplied as SELF or generic explicit grant',()=>{
 const h=fixture();const g=h.docs['memberExplicitGrants/SYN_DYNAMIC_BINDING'];g.targets.commandScopes=['PROFILE_CORRECTION','PROFILE_CORRECTION_CURRENT_MEMBERS'];a.throws(()=>h.run('return f4MemberAuthzContracts_().grant(get("memberExplicitGrants/SYN_DYNAMIC_BINDING"));'),/M13_AUTHZ_DENIED/);
 const p=h.policy;p.proofKind='SELF';p.actorSelector={kind:'SELF'};p.capabilities=['MEMBERS_SELF_READ'];p.fieldMasks={MEMBERS_SELF_READ:['idPessoa']};p.targets.ownPerson=true;a.throws(()=>h.run('return f4MemberAuthzContracts_().policy(get("memberAuthzPolicies/PEOPLE_POLICY"));'),/M13_AUTHZ_DENIED/);
});
test('dynamic scope requires paired base scope and the specific role/capability',()=>{
 for(const change of [p=>p.targets.commandScopes=['PROFILE_CORRECTION_CURRENT_MEMBERS'],p=>{p.capabilities=['PEOPLE_DIRECTORY_READ'];p.fieldMasks={PEOPLE_DIRECTORY_READ:['idPessoa']};},p=>{p.actorSelector={kind:'ROLE',roleKey:'DIRETOR_EXTENSAO',group:'DIRETORIA',actorScope:'EXTENSAO'};}]){const h=fixture();change(h.policy);a.throws(()=>h.run('return f4MemberAuthzContracts_().policy(get("memberAuthzPolicies/PEOPLE_POLICY"));'),/M13_AUTHZ_DENIED/);}
 const h=fixture();h.policy.targets.commandScopes.push('ACCOUNT_BINDING_CURRENT_MEMBERS');a.throws(()=>h.run('return f4MemberAuthzContracts_().policy(get("memberAuthzPolicies/PEOPLE_POLICY"));'),/M13_AUTHZ_DENIED/);
});
test('current target eligibility derives source graph and excludes nonmembers, inactive and technical people',()=>{
 const h=fixture();const eligible=id=>h.run('return f4MemberProfileTargetEligible_(get,x.id,x.now);',{id});a.equal(eligible('PERSON_000'),true);a.equal(eligible('PERSON_001'),false);a.equal(eligible('PERSON_004'),true);
 h.docs['memberIdentities/PERSON_000'].technicalAccount=true;a.equal(eligible('PERSON_000'),false);h.docs['memberIdentities/PERSON_000'].technicalAccount=false;h.docs['memberIdentities/PERSON_000'].active=false;a.equal(eligible('PERSON_000'),false);
});
test('correction reader retains explicit dynamic mode, live role, field mask and revocation',()=>{
 const h=fixture(),read=()=>h.run('const r=f4MemberProfileCorrectionReader_({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:x.now+60000},clock:()=>x.now}).access(get,x.now,false);return {dynamic:r.dynamic,persons:r.persons,mask:r.mask};');a.deepEqual(read(),{dynamic:true,persons:null,mask:['displayName']});h.policy.active=false;a.throws(read,/ACCESS_REVOKED/);h.policy.active=true;h.docs['memberFunctions/VIG-001'].endDateReal='2020-01-01';a.throws(read);
});
test('dynamic binding requires explicit grant AND current administrative function before lookup',()=>{
 const h=fixture(),access=id=>h.run('f4MemberBindingAdmin_({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:x.now+60000},clock:()=>x.now}).access(get,x.id);return true;',{id});a.equal(access('PERSON_000'),true);a.throws(()=>access('PERSON_001'),/M13_AUTHZ_DENIED/);h.docs['memberExplicitGrants/SYN_DYNAMIC_BINDING'].active=false;a.throws(()=>access('PERSON_000'));
 const rh=fixture('DIRETOR_RECURSOS_HUMANOS');a.throws(()=>rh.run('f4MemberBindingAdmin_({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:x.now+60000},clock:()=>x.now}).access(get,"PERSON_000");return true;'),/M13_AUTHZ_DENIED/);
});
test('executor dynamic opt-in is command specific and cannot become an empty universal scope',()=>{
 const h=fixture(),fields=(id,type)=>h.run('f4MemberExecutor_().fields(get,{commandType:x.type,target:{personId:x.id},payload:{}},{scopeId:"SYN_DYNAMIC",commandType:x.type});return true;',{id,type});a.equal(fields('PERSON_000','members.profile.correct'),true);a.equal(fields('PERSON_000','members.binding.associate'),true);a.throws(()=>fields('PERSON_001','members.profile.correct'),/PILOT_RESOURCE_DENIED/);a.throws(()=>fields('PERSON_000','members.ingress.execute'),/PILOT_RESOURCE_DENIED/);const s=h.docs['memberExecutorScopes/SYN_DYNAMIC'];delete s.profileCorrectionCurrentMembers;delete s.bindingCurrentMembers;a.throws(()=>fields('PERSON_000','members.profile.correct'),/ACCESS_REVOKED/);
});
module.exports={fixture};
