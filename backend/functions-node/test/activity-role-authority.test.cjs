'use strict';
const {test}=require('node:test'),a=require('node:assert/strict'),{seed}=require('./fixtures.cjs'),{decode,encode}=require('../src/firestore.cjs'),{createDomain}=require('../src/domain.cjs');
function setup(role='DIRETOR_RECURSOS_HUMANOS',change=()=>{}){
 const f=seed(),docs=Object.fromEntries(Object.entries(f.docs).map(([p,v])=>[p,decode({mapValue:v})]));
 const scope=role==='DIRETOR_RECURSOS_HUMANOS'?'RECURSOS_HUMANOS':'PRESIDENCIA';
 docs['memberRoleCatalog/CARGO-001'].roleKey=role;docs['memberFunctions/VIG-001'].scopeKey=scope;
 const policy=docs['memberAuthzPolicies/CF_DEV_ACT_POLICY'];policy.capabilities=['ACTIVITY_BOARD_READ'];policy.fieldMasks={ACTIVITY_BOARD_READ:[]};policy.actorSelector={kind:'ROLE',roleKey:role,group:'DIRETORIA',actorScope:scope};
 // No technical DEV_ADMIN and no production identity or credentials.
 for(const p of Object.keys(docs))if(p.startsWith('coreDevAdmins/'))delete docs[p];
 docs['projectionRecoveryControl/current']={actorUids:['CF_ADMIN']};
 change(docs);
 const domain=createDomain(raw=>{const x=JSON.parse(raw);if(x.kind==='query'){const rows=Object.entries(docs).filter(([p,v])=>p.split('/').length===2&&p.startsWith(x.plan.collection+'/')&&x.plan.filters.every(f=>v[f.field]===decode(f.value)));return JSON.stringify(rows.slice(0,x.plan.limit).map(([path,data])=>({path,data:encode(data)})));}if(x.kind==='get')return JSON.stringify(docs[x.path]?encode(docs[x.path]):null);if(['getAll','guard'].includes(x.kind))return 'null';if(x.kind==='pilotConfig')return JSON.stringify({root:'',executor:'CANONICAL_READ_ONLY',settings:{commands:[],policyIds:['CF_DEV_ACT_POLICY']}});throw Error('UNEXPECTED_IO_'+x.kind);});
 const read=()=>JSON.parse(domain.run('(now)=>f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:now+60000},settings:nodeSettings(),clock:()=>now}).readAuthority(["activityAttendanceRegistry/current"],false)',f.now));
 const board=()=>JSON.parse(domain.run('(now)=>{const options={config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:now+60000},settings:nodeSettings(),clock:()=>now};options.ports=f4ActivitiesFunctionalPorts_(options);return f4BoardCompetencies_(options);}',f.now));return {docs,read,board};
}
for(const role of ['PRESIDENTE','VICE_PRESIDENTE','DIRETOR_RECURSOS_HUMANOS'])test('canonical ROLE read without technical admin: '+role,()=>a.equal(setup(role).read().q.kind,'ROLE'));
for(const [name,change]of Object.entries({expiredPolicy:d=>d['memberAuthzPolicies/CF_DEV_ACT_POLICY'].expiresAt=d['memberAuthzPolicies/CF_DEV_ACT_POLICY'].issuedAt,revokedPolicy:d=>d['memberAuthzPolicies/CF_DEV_ACT_POLICY'].active=false,emptyPersons:d=>d['memberAuthzPolicies/CF_DEV_ACT_POLICY'].targets.allowedPersons=[],missingBinding:d=>delete d['memberAuthBindings/CF_ADMIN'],endedFunction:d=>d['memberFunctions/VIG-001'].endDateReal='2026-02-01',missingCapability:d=>d['memberAuthzPolicies/CF_DEV_ACT_POLICY'].capabilities=[]}))test('canonical read fails closed: '+name,()=>a.throws(setup(undefined,change).read));

test('board reader derives safe names and permissions from canonical ROLE without private grants',()=>{const h=setup(),r=h.board();a.equal(r.data.items.length,1);a.equal(r.data.items[0].roleKey,'DIRETOR_RECURSOS_HUMANOS');a.equal(r.data.items[0].permissions[0].status,'CONDITIONAL');const wire=JSON.stringify(r);for(const forbidden of ['CF_ADMIN','CF_DEV_ACT_POLICY','actorSelector','token','cpf','contactEmail','allowedResources'])a(!wire.includes(forbidden),forbidden);});
test('board reader rejects an ended or revoked caller, no technical fallback',()=>{a.throws(setup(undefined,d=>d['memberFunctions/VIG-001'].endDateReal='2026-02-01').board);a.throws(setup(undefined,d=>d['memberAuthzPolicies/CF_DEV_ACT_POLICY'].active=false).board);});
