'use strict';
const {runtime,seed}=require('./runtime.cjs'),{scenario}=require('./activities/scenarios.cjs'),{M,NOW}=require('./activities/harness.cjs'),fs=require('node:fs'),p=require('node:path');
function prepare(cmd,options={}){
 const s=scenario(cmd),ct=f4ActivitiesContracts_()[cmd],principal=s.principal(s.handle),system=['SYSTEM_JOB','WORKSPACE_EXECUTOR'].includes(principal.kind),kind=options.kind|| (system?principal.kind:principal.kind==='PORTAL_AUTHENTICATED_MEMBER'?'SELF':'ROLE');
 const remap=v=>typeof v==='string'?v==='p1'?'PES-000001':v==='p2'?'PES-000002':v==='u1'?'SYN_UID':v:v&&typeof v==='object'?Array.isArray(v)?v.map(remap):Object.fromEntries(Object.entries(v).map(([k,x])=>[k,remap(x)])):v;
 const docs={};for(const [path,d]of s.db.docs)if(!path.startsWith('member')&&!path.startsWith('portalUsers'))docs[path.replace(/\/p1(?=\/|$)/g,'/PES-000001').replace(/\/p2(?=\/|$)/g,'/PES-000002')]=remap(FIRE(d));
 function FIRE(v){return f4Firestore_().encode(v).mapValue.fields}
 const spec={q:remap(s.q),kind,principal:remap(principal),docs,now:NOW};const r=runtime();
 seed(r,r.run(spec=>{
  const F=f4Firestore_(),T=F.Timestamp.fromMillis,C=f4Core_();globalThis.tick=spec.now;
  const x=m13BFixture_(spec.kind==='SELF'?'SELF':'ROLE',spec.kind==='SELF'?'MEMBERS_SELF_READ':'MEMBERS_PROFILE_CORRECT','DIRETOR_ADMINISTRATIVO',tick);x.d['memberAuthorityControl/MEMBERS'].activeWriterVersion='M13';
  for(const [k,v]of Object.entries(spec.docs))x.d[k]=F.decode({mapValue:{fields:v}});
  const q=spec.q;if(q.commandType==='ACT.INCIDENCE_EVALUATE'){const old=q.expected.find(e=>e.kind==='obligation'&&e.exists===false);old.id=f4ActivitiesModel_().hash({person:'PES-000001',cycle:'c'.repeat(32),kind:'ORDINARY',version:1}).slice(0,32);}
  if(q.commandType==='ACT.MEMBER_EXIT_EFFECT'){x.d['memberEvents/exit']={idPessoa:'PES-000001',cycleRef:{kind:'cycle',id:'c'.repeat(32)},revision:1};x.d['memberCurrentStates/PES-000001'].state='DESLIGADO';}
  const sys=['SYSTEM_JOB','WORKSPACE_EXECUTOR'].includes(spec.kind),a=sys?{kind:spec.kind,servicePrincipal:spec.principal.servicePrincipal,scope:spec.principal.scope,ruleRef:spec.principal.ruleRef,jobKind:spec.principal.jobKind||'WORKSPACE',runId:spec.principal.runId}:{kind:spec.kind,sourceId:'ACT_POLICY',functionId:spec.kind==='ROLE'?'VIG-001':null,actorScope:spec.kind==='ROLE'?'ADMINISTRACAO':null};
  globalThis.input={schemaVersion:'geapa-act-command-v13',operationId:q.operationId,commandType:q.commandType,target:{scope:q.scope,expected:q.expected,authorization:a},payload:q.payload,evidenceRefs:[],decisionRef:null};
  const cap=sys?'ACT_INTERNAL':f4ActivitiesPipeline_().capability(input);
  if(!sys){const resources=q.scope.targetRefs.map(f4ActivitiesModel_().path),scope=spec.kind==='SELF'?'SELF':q.commandType==='ACT.CYCLE_CREATE'?'PRESENTATIONS_PLAN':'PRESC_'+'c'.repeat(32);x.d['memberAuthzPolicies/ACT_POLICY']={schemaVersion:'act-authz-policy-v13',id:'ACT_POLICY',revision:1,active:true,origin:'ACT_CANONICAL_APPROVAL',approvalRef:'SYN_APPROVAL',issuedAt:T(tick-1000),expiresAt:T(tick+3600000),proofKind:spec.kind,capabilities:[cap],actorSelector:spec.kind==='SELF'?{kind:'SELF'}:{kind:'ROLE',roleKey:'DIRETOR_ADMINISTRATIVO',group:'DIRETORIA',actorScope:'ADMINISTRACAO'},targets:{ownPerson:spec.kind==='SELF',allowedTargetScopes:[scope],allowedBoards:[],allowedRoles:[],allowedPersons:spec.kind==='SELF'?[]:['PES-000001'],allowedResources:resources},fieldMasks:{[cap]:[]}};}
  if(spec.kind==='EXPLICIT'){const old=x.d['memberAuthzPolicies/ACT_POLICY'];delete x.d['memberAuthzPolicies/ACT_POLICY'];const b=x.d['memberAuthBindings/SYN_UID'];const {proofKind,actorSelector,...common}=old;x.d['memberExplicitGrants/ACT_POLICY']={...common,schemaVersion:'act-explicit-grant-v13',issuerRef:'SYN_ISSUER',uid:'SYN_UID',idPessoa:b.idPessoa,bindingVersion:b.bindingVersion,subjectEpoch:1,actorScopes:['ADMINISTRACAO'],environment:'DEV',runId:null};input.target.authorization.actorScope='ADMINISTRACAO';}
  globalThis.pipeline=f4CorePipeline_({config:fixture.config,clock:()=>tick,verifyIdentity:token=>{C.need(token==='SYNTHETIC','AUTH_BINDING_INVALID');return {uid:'SYN_UID',expiresAt:tick+3600000}},membersAuthorization:()=>[],membersRevisions:()=>[]});
  return Object.fromEntries(Object.entries(x.d).map(([k,v])=>[k,F.encode(v).mapValue.fields]));
 },spec));
return {r,system};
}
module.exports={prepare};
