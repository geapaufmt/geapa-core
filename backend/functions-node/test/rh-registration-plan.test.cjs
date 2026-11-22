const {test}=require('node:test'),a=require('node:assert/strict');
const {seed}=require('./fixtures.cjs'),{decode}=require('../src/firestore.cjs'),{profile}=require('../src/runtime-profile.cjs'),{plan,digest}=require('../scripts/administer-rh-registration.cjs');
function fixture(){const f=seed(),docs=Object.fromEntries(Object.entries(f.docs).map(([p,v])=>[p,decode({mapValue:v})]));
 docs['memberRoleCatalog/CARGO-001'].roleKey='DIRETOR_RECURSOS_HUMANOS';docs['memberFunctions/VIG-001'].scopeKey='RECURSOS_HUMANOS';
 docs['memberAuthzPolicies/FAST_MEMBER_ROLE'].actorSelector.roleKey='DIRETOR_RECURSOS_HUMANOS';docs['memberAuthzPolicies/FAST_MEMBER_ROLE'].actorSelector.actorScope='RECURSOS_HUMANOS';
 docs['memberExecutorControl/current']={schemaVersion:'member-executor-control-v1',enabled:true,projectId:profile.projectId,revision:3,actorUids:['CF_ADMIN','SYN_OTHER'],scopeIds:['ORIGINAL'],commandTypes:['members.self.update'],portalRoutes:['/meu-perfil/salvar'],browserEnabled:true,appCheck:{mode:'REQUIRED'}};
 for(const id of ['MEMBERS_INGRESS_EXECUTE','MEMBERS_ONBOARDING_COMPLETE'])delete docs['commandExecutorRoutes/'+id];
 docs['memberOperationalConfig/notifications']={enabled:true,mode:'DRY_RUN',projectId:profile.projectId,environment:profile.environment};
 return {docs,now:f.now,uid:'CF_ADMIN',projectId:profile.projectId,sourcePolicyId:'FAST_MEMBER_ROLE',authorizationRef:'SYN_RH_ENABLE'};
}
test('pure RH plan preserves actors and existing permissions, rechecks canonical role and inherits expiry',()=>{
 const f=fixture(),before=digest(f.docs),p=plan(f),g=p.writes['memberAuthzPolicies/RH_PREREGISTRATION_DRAFTS'];
 a.equal(digest(f.docs),before);a.equal(Object.keys(p.writes).length,6);a.equal(p.individualApprovalsCreated,false);
 a.deepEqual(g.capabilities,['MEMBERS_INGRESS_EXECUTE']);a.equal(g.origin,f.docs['memberAuthzPolicies/FAST_MEMBER_ROLE'].origin);a.equal(g.expiresAt,f.docs['memberAuthzPolicies/FAST_MEMBER_ROLE'].expiresAt);
 a.deepEqual(p.writes['memberExecutorControl/current'].actorUids,['CF_ADMIN','SYN_OTHER']);
 a.deepEqual(p.writes['memberExecutorScopes/RH_PREREGISTRATION_DRAFTS'].personIds,[]);a.equal(p.writes['memberExecutorScopes/RH_PREREGISTRATION_DRAFTS'].preregistrationDrafts,true);
 a(p.preconditions.some(v=>v.path==='memberFunctions/VIG-001'));a(p.preconditions.some(v=>v.path==='memberAuthBindings/CF_ADMIN'));
 a(Object.keys(p.writes).every(v=>!v.startsWith('memberAdmissionApproval')&&!v.startsWith('memberFunctions/')&&!v.startsWith('projectionRecoveryControl/')));
});
test('RH planner fails closed on wrong project, stale role/policy, absent actor, existing route or live mail',()=>{
 for(const mutate of [f=>f.projectId='other',f=>f.uid='CF_NONE',f=>f.docs['memberFunctions/VIG-001'].scopeKey='ADMINISTRACAO',f=>f.docs['memberAuthzPolicies/FAST_MEMBER_ROLE'].active=false,f=>f.docs['memberAuthzPolicies/FAST_MEMBER_ROLE'].expiresAt=new Date(f.now-1),f=>f.docs['memberExecutorControl/current'].actorUids=[],f=>f.docs['commandExecutorRoutes/MEMBERS_INGRESS_EXECUTE']={revision:1},f=>f.docs['memberOperationalConfig/notifications'].mode='LIVE']){const f=fixture();mutate(f);const before=digest(f.docs);a.throws(()=>plan(f));a.equal(digest(f.docs),before);}
});
