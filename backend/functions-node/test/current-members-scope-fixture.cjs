'use strict';
const {setup}=require('./people-fixture.cjs'),E=require('../src/environment.cjs');
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
module.exports={fixture};
