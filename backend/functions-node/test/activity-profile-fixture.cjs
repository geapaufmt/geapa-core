'use strict';
const a=require('node:assert/strict'),{seed}=require('./fixtures.cjs'),{decode,encode}=require('../src/firestore.cjs'),{createDomain}=require('../src/domain.cjs');
function setup(role='DIRETOR_RECURSOS_HUMANOS',change=()=>{}){
 const f=seed(),docs=Object.fromEntries(Object.entries(f.docs).map(([p,v])=>[p,decode({mapValue:v})]));
 const scope=role==='DIRETOR_RECURSOS_HUMANOS'?'RECURSOS_HUMANOS':role==='DIRETOR_ADMINISTRATIVO'?'ADMINISTRACAO':'PRESIDENCIA';
 docs['memberRoleCatalog/CARGO-001'].roleKey=role;docs['memberFunctions/VIG-001'].scopeKey=scope;
 const policy=docs['memberAuthzPolicies/CF_DEV_ACT_POLICY'];policy.capabilities=['ACTIVITY_BOARD_READ'];policy.fieldMasks={ACTIVITY_BOARD_READ:[]};policy.actorSelector={kind:'ROLE',roleKey:role,group:'DIRETORIA',actorScope:scope};
 // No technical DEV_ADMIN and no production identity or credentials.
 for(const p of Object.keys(docs))if(p.startsWith('coreDevAdmins/'))delete docs[p];
 docs['projectionRecoveryControl/current']={actorUids:['CF_ADMIN']};
 change(docs);
 const domain=createDomain(raw=>{const x=JSON.parse(raw);if(x.kind==='query'){const rows=Object.entries(docs).filter(([p,v])=>p.split('/').length===2&&p.startsWith(x.plan.collection+'/')&&x.plan.filters.every(f=>require('node:util').isDeepStrictEqual(v[f.field],decode(f.value))));return JSON.stringify(rows.slice(0,x.plan.limit).map(([path,data])=>({path,data:encode(data)})));}if(x.kind==='get')return JSON.stringify(docs[x.path]?encode(docs[x.path]):null);if(['getAll','guard'].includes(x.kind))return 'null';if(x.kind==='pilotConfig')return JSON.stringify({root:'',executor:'CANONICAL_READ_ONLY',settings:{commands:[],policyIds:['CF_DEV_ACT_POLICY']}});throw Error('UNEXPECTED_IO_'+x.kind);});
 const read=()=>JSON.parse(domain.run('(now)=>f4ActivitiesFunctionalPorts_({config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:now+60000},settings:nodeSettings(),clock:()=>now}).readAuthority(["activityAttendanceRegistry/current"],false)',f.now));
 const board=()=>JSON.parse(domain.run('(now)=>{const options={config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:now+60000},settings:nodeSettings(),clock:()=>now};options.ports=f4ActivitiesFunctionalPorts_(options);return f4BoardCompetencies_(options);}',f.now));const list=(route="/admin/atividades")=>JSON.parse(domain.run('(now)=>{const options={config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:now+60000},settings:nodeSettings(),clock:()=>now.time};options.identity.expiresAt=now.time+60000;options.ports=f4ActivitiesFunctionalPorts_(options);return f4ActivitiesActivityCompat_(options).read(now.route,{});}',{time:f.now,route}));return {docs,read,board,list,domain};
}

module.exports={setup};

