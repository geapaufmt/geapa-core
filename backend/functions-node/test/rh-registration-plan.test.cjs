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

test('DEV fixture 4 has no membership or admission approval; nominal draft window closes without touching base grant',()=>{
 const {draftFixture,planDevDraft}=require('../scripts/administer-rh-registration.cjs'),{setup}=require('./people-fixture.cjs'),{grant}=require('./dev-administrative-fixture.cjs');
 const h=setup('DIRETOR_RECURSOS_HUMANOS',0),f=draftFixture('CF_ADMIN',h.now);Object.assign(h.docs,f.docs);const g=grant(h);
 g.capabilities.push('MEMBERS_INGRESS_EXECUTE');g.fieldMasks.MEMBERS_INGRESS_EXECUTE=['preRegister'];g.syntheticPersonIds.push(f.personId);
 a(Object.keys(f.docs).every(p=>!/^member(Links|Events|AdmissionApproval|AuthBindings|Enrollment)/.test(p)));
 const args={docs:h.docs,uid:'CF_ADMIN',personId:f.personId,registration:f.registration,expectedGrantRevision:g.revision,now:h.now},before=digest(h.docs),p=planDevDraft(args);
 a.equal(digest(h.docs),before);a.equal(Object.keys(p.writes).length,1);a.equal(p.individualApprovalsCreated,false);
 const gp='coreDevAdministrativeGrants/CF_ADMIN';Object.assign(h.docs,p.writes);
 a(h.docs[gp].preregistrationWindow.expiresAt.getTime()-h.now<=3600000);
 const allowed=registration=>JSON.parse(h.domain.run('(x)=>{const d=f4DevAdministrativeAccess_(),get=p=>f4Firestore_().connect(nodeConfig).get(p),a=d.source(get,{uid:"CF_ADMIN",expiresAt:x.now+60000},x.now,"MEMBERS_INGRESS_EXECUTE");return d.draftAllowed(get,a,x.registration,x.id,x.now);}',{now:h.now,id:f.personId,registration}));
 a.equal(allowed(f.registration),true);a.equal(allowed({...f.registration,email:'other@example.invalid'}),false);
 const close=planDevDraft({...args,expectedGrantRevision:g.revision+1,close:true});a.equal(close.writes[gp].preregistrationWindow,undefined);
 const {revision,...base}=g,{revision:newRevision,...closed}=close.writes[gp];a.deepEqual(closed,base);
 for(const change of [x=>x.registration={...f.registration,email:'real@example.com'},x=>x.expectedGrantRevision++,x=>x.docs['memberIdentities/'+f.personId]={...f.docs['memberIdentities/'+f.personId],origin:'REAL'}]){const x={...args,docs:{...h.docs,[gp]:g}};change(x);a.throws(()=>planDevDraft(x));}
});
