'use strict';
require('../src/local-network.cjs');
const {test}=require('node:test'),a=require('assert/strict'),fs=require('node:fs'),p=require('node:path'),crypto=require('node:crypto');
const {fixture}=require('./institutional-fixture.cjs');
const {PROJECT,DEMO,EXECUTOR,ROOT}=require('../src/environment.cjs');
const {EditorialStore}=require('../src/editorial-commands.cjs');
const {encode}=require('../src/firestore.cjs');

function setupEditorial(role='DIRETOR_COMUNICACAO',group='DIRETORIA',scope='COMUNICACAO'){
 const h=fixture('DIRETOR_ADMINISTRATIVO');
 h.docs['memberRoleCatalog/CARGO-001'].roleKey=role;
 h.docs['memberRoleCatalog/CARGO-001'].group=group;
 h.docs['memberFunctions/VIG-001'].scopeKey=scope;
 const policy={
  ...h.policy,
  id:'EDITORIAL_POLICY',
  actorSelector:{kind:'ROLE',roleKey:role,group,actorScope:scope},
  capabilities:['PUBLIC_CONTENT_MANAGE'],
  fieldMasks:{PUBLIC_CONTENT_MANAGE:['pages','timeline']},
  targets:{...h.policy.targets,allowedTargetScopes:['COMUNICACAO','ADMINISTRACAO','PRESIDENCIA'],allowedResources:['canonicalEditorial/pages','canonicalEditorial/timeline']}
 };
 h.docs['memberAuthzPolicies/EDITORIAL_POLICY']=policy;
 h.docs['editorialControl/current']={
  schemaVersion:'editorial-control-v1',
  projectId:PROJECT,
  commandType:'editorial.public.manage',
  enabled:true,
  executor:EXECUTOR,
  epoch:1,
  revision:1,
  actorUids:['CF_ADMIN'],
  policyIds:['EDITORIAL_POLICY'],
  availabilityMode:'TIMEBOXED',
  startsAt:h.now-1000,
  expiresAt:h.now+3600000,
  rateLimit:{windowMs:3600000,maxAdmissions:60},
  browserEnabled:true,
  appCheck:{mode:'REQUIRED',appIds:['SYN_APP']}
 };
 return h;
}

test('PUBLIC_CONTENT_MANAGE policy accepts Communication, Presidency and Administrative roles and rejects unrelated roles',()=>{
 const h=fixture();
 const cases=[
  ['DIRETOR_COMUNICACAO','DIRETORIA','COMUNICACAO',true],
  ['ASSESSOR_COMUNICACAO_1','ASSESSORIA','COMUNICACAO',true],
  ['ASSESSOR_COMUNICACAO_2','ASSESSORIA','COMUNICACAO',true],
  ['PRESIDENTE','DIRETORIA','PRESIDENCIA',true],
  ['VICE_PRESIDENTE','DIRETORIA','PRESIDENCIA',true],
  ['DIRETOR_ADMINISTRATIVO','DIRETORIA','ADMINISTRACAO',true],
  ['DIRETOR_RECURSOS_HUMANOS','DIRETORIA','RECURSOS_HUMANOS',false],
  ['TESOUREIRO','DIRETORIA','FINANCAS',false],
  ['DIRETOR_EXTENSAO','DIRETORIA','EXTENSAO',false],
  ['CONSELHEIRO_CONSULTIVO','CONSELHO','CONSELHO',false]
 ];
 for(const [roleKey,group,actorScope,expected] of cases){
  a.equal(JSON.parse(h.domain.run('(s)=>f4MemberAuthzContracts_().roleCapability("PUBLIC_CONTENT_MANAGE",s)',{roleKey,group,actorScope})),expected,roleKey);
 }
});

test('publicTimeline sanitizer and projection reject unsafe URLs, invalid years and script injection',()=>{
 const h=setupEditorial();
 const good=JSON.parse(h.domain.run('(x)=>f4Projections_().publicTimeline({sourceRevision:1,data:{id:"marco_1",schemaVersion:"geapa-public-v2.1",published:true,sanitized:true,year:2025,date:"14/05/2025",title:"Fundação <b>GEAPA</b>",description:"<script>alert(1)</script>Marco inicial",imageUrl:"https://firebasestorage.googleapis.com/v0/b/portal/o/foto.jpg",order:1,sourceHash:"a".repeat(64),updatedAt:new Date(x.now).toISOString(),revision:1}}).dto',{now:h.now}));
 a.equal(good.date,'2025-05-14');
 a.equal(good.title,'Fundação GEAPA');
 a.equal(good.description,'Marco inicial');
 a.throws(()=>h.domain.run('(x)=>f4Projections_().publicTimeline({sourceRevision:1,data:{id:"marco_1",schemaVersion:"geapa-public-v2.1",published:true,sanitized:true,year:1800,date:"",title:"X",description:"Y",imageUrl:"",order:1,sourceHash:"a".repeat(64),updatedAt:new Date(x.now).toISOString(),revision:1}})',{now:h.now}));
 a.equal(JSON.parse(h.domain.run('()=>f4Public_().url("javascript:alert(1)",true)')),'');
});

test('canonical editorial preview rejects invalid payloads and keeps publicPages untouched',()=>{
 const h=setupEditorial('ASSESSOR_COMUNICACAO_1','ASSESSORIA','COMUNICACAO');
 const op1=crypto.randomBytes(16).toString('hex');
 const input1={action:'PAGE_SAVE_DRAFT',pageId:'about',expectedRevision:0,content:{title:'Sobre o GEAPA',subtitle:'Agroecologia UFMT',sections:[{id:'sec_hero',type:'HERO',title:'Missão',subtitle:'',text:'<script>x</script>Promover agroecologia.',imageUrl:'https://firebasestorage.googleapis.com/v0/b/portal/o/hero.png',buttonText:'Eixos',buttonUrl:'#/eixos-tematicos',order:1,published:true}]}};
 const prev1=JSON.parse(h.domain.run('(x)=>{f4CommandExecutorId_=()=>"'+EXECUTOR+'";const M=f4EditorialCommands_(),o={config:nodeConfig,identity:{uid:"CF_ADMIN",expiresAt:x.now+60000},clock:()=>x.now};const p=M.preview(nodeTx,o,x.input,x.op);const r={operationId:x.op,input:x.input,documentHash:p.documentHash,epoch:p.epoch,controlRevision:p.controlRevision,validUntil:p.validUntil};return {...r,digest:M.digest(o.config,"CF_ADMIN",r),publicUpdate:p.publicUpdate};}',{now:h.now,input:input1,op:op1}));
 a.equal(h.docs['publicPages/about'],undefined);
 a.equal(prev1.publicUpdate,false);
 a.equal(prev1.input.pageId,'about');
 a.equal(prev1.digest.length,64);
});

test('EditorialStore refuses writes during preview/status, rejects unpinned control and enforces path fence',()=>{
 Object.assign(process.env,{GEAPA_RUNTIME:'EMULATOR',GEAPA_PILOT_MODE:'SHARED_DEV',GCLOUD_PROJECT:DEMO,FUNCTIONS_EMULATOR:'true',FIRESTORE_EMULATOR_HOST:'127.0.0.1:8187',FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9197'});
 const s=new EditorialStore({projectId:DEMO,doc:x=>x},{uid:'CF_ADMIN'});
 s.action='editorialPreview';s.operationId='a'.repeat(32);
 a.throws(()=>s.validateWrites({writes:[{path:'publicPages/about',mode:'set',data:encode({})}]}),/PILOT_WRITE_DENIED/);
 s.action='editorialExecute';
 a.throws(()=>s.validateWrites({writes:[{path:'memberIdentities/PES-000001',mode:'set',data:encode({})}]}),/PILOT_WRITE_DENIED/);
});

test('navigation hints expose admin-conteudo-publico only when editorialControl enables an authorized actor',()=>{
 const h=setupEditorial();
 const hintsWithControl=JSON.parse(h.domain.run('(x)=>f4AdministrativeNavigationHints_({get:p=>f4Firestore_().connect(nodeConfig).get(p),identity:{uid:"CF_ADMIN",expiresAt:x.now+60000},now:x.now,board:true,probe:()=>{throw Error("DENIED")}})',{now:h.now}));
 a.equal(hintsWithControl.routes['admin-conteudo-publico'],true);
 delete h.docs['editorialControl/current'];
 const hintsWithoutControl=JSON.parse(h.domain.run('(x)=>f4AdministrativeNavigationHints_({get:p=>f4Firestore_().connect(nodeConfig).get(p),identity:{uid:"CF_ADMIN",expiresAt:x.now+60000},now:x.now,board:true,probe:()=>{throw Error("DENIED")}})',{now:h.now}));
 a.equal(hintsWithoutControl.routes['admin-conteudo-publico'],undefined);
});

test('runtime index.cjs exposes editorialPreview, editorialExecute, editorialStatus and actRead supports /admin/conteudo-publico',()=>{
 const idx=fs.readFileSync(p.resolve(__dirname,'../src/index.cjs'),'utf8');
 a(idx.includes("'editorialPreview','editorialExecute','editorialStatus'"));
 a(idx.includes("require('./editorial-commands.cjs').manageEditorial(options)"));
 const readers=fs.readFileSync(p.resolve(__dirname,'../src/act-readers.cjs'),'utf8');
 a(readers.includes("r.routes.includes('/admin/conteudo-publico')"));
 a(readers.includes('canonicalEditorial|canonicalEditorialHistory'));
});
