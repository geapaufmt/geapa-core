'use strict';
const H={C:f4Core_()};
const M=f4MemberModel_(),RealR=f4MemberRepository_(),R={...RealR,connect:(c,id='PES-000001')=>RealR.connect(c,id)},C=H.C;
const NOW=C.civilStart('2026-09-14')+12*3600000,config=f4Commands_().fixed;
const entity=(col,id,fields)=>({schemaVersion:col==='memberAuthBindings'?'binding-v1':'m1-'+col+'-v1',id,revision:1,origin:'SYNTHETIC_M1',...fields});
function base(type='MEMBRO_EFETIVO',technical=false){const d={};const add=(c,id,v)=>d[c+'/'+id]=entity(c,id,v);
 add('memberIdentities','PES-000001',{idPessoa:'PES-000001',displayName:'Pessoa sintética',technicalAccount:technical,active:true});
 if(!technical&&type!==null)add('memberLinks','VINC-001',{idPessoa:'PES-000001',type,startDate:type==='EGRESSO'?'2026-09-01':'2026-01-01',endDate:null,exitDate:type==='EGRESSO'?'2026-08-31':null});
 add('memberAuthBindings','SYN_UID',{uid:'SYN_UID',idPessoa:'PES-000001',active:true,bindingVersion:1,policyVersion:1,sourceRevision:1,createdAt:NOW,updatedAt:NOW,domainVersion:'m1'});return d;
}
function role(d,group='DIRETORIA',start='2026-01-01',end='2026-12-31'){
 const add=(c,id,v)=>d[c+'/'+id]=entity(c,id,v);
 add('memberGovernanceBoards','DIR-001',{name:'Diretoria sintética',startDate:'2026-01-01',endDate:'2026-12-31'});
 add('memberGovernanceTerms','TERM-001',{name:'Termo sintético',startDate:'2026-01-01',endDate:'2026-12-31'});
 add('memberGovernanceBoardTerms','REL-001',{idDiretoria:'DIR-001',idTermo:'TERM-001',startDate:'2026-01-01',endDate:'2026-12-31'});
 add('memberRoleCatalog','CARGO-001',{roleKey:'DIRETOR_SINTETICO',group,active:true});
 add('memberFunctions','VIG-001',{idPessoa:'PES-000001',idVinculo:'VINC-001',idCargo:'CARGO-001',scopeKey:'EDITORIAL',idDiretoria:'DIR-001',idTermo:'TERM-001',startDate:start,endDatePlanned:end,endDateReal:null});
 d['memberRoleCatalog/M1_POLICY']={schemaVersion:'m1-policy-v1',id:'M1_POLICY',revision:1,origin:'PORTAL_PERMISSOES',rules:[{idCargo:'CARGO-001',scopeKey:'EDITORIAL',capabilities:['PUBLIC_CONTENT_MANAGE','MEMBER_SCOPED_READ','ACTIVITY_SCOPED_READ']}]};return d;
}
function event(d,type,date='2026-09-12',end=null,id='EV-001'){d['memberEvents/'+id]=entity('memberEvents',id,{idPessoa:'PES-000001',idVinculo:'VINC-001',type,effect:type,effectiveDate:date,endDate:end,supersedes:null,correctionPath:null,prior:null});return d}
function physical(d){const out=structuredClone(d);for(const b of M.rows(out,'memberAuthBindings'))b.domainVersion='m11';const pol=out['memberRoleCatalog/M1_POLICY'];if(pol){delete out['memberRoleCatalog/M1_POLICY'];for(const r of pol.rules){const id=r.idCargo+'__'+r.scopeKey;out[R.policyPath(r.idCargo,r.scopeKey)]={schemaVersion:'m11-policy-v1',id,revision:pol.revision,origin:pol.origin,...r}}}return out}
const changes=d=>Object.entries(physical(d)).map(([path,data])=>({path,data,expectedRevision:0}));
function projectionMap(d,rev=1,now=NOW){const out=physical(d),docs=structuredClone(d);for(const b of M.rows(docs,'memberAuthBindings'))b.domainVersion='m11';const paths=Object.keys(out).filter(p=>/^member(Identities|Identifiers|PrivateProfiles|Links|Events|Functions)\//.test(p));out[R.personPath('PES-000001')]={schemaVersion:'m11-head-v1',id:'PES-000001',kind:'PERSON',revision:rev,paths};const boards={},policies={};for(const b of M.rows(d,'memberGovernanceBoards')){const bh={schemaVersion:'m11-head-v1',id:b.id,kind:'BOARD',revision:1,paths:Object.keys(d).filter(p=>p==='memberGovernanceBoards/'+b.id||p.startsWith('memberGovernanceBoardTerms/')&&d[p].idDiretoria===b.id)};boards[b.id]=bh;out[R.boardPath(b.id)]=bh}for(const [p,v]of Object.entries(out))if(p.startsWith('memberRolePolicies/'))policies[p]=v;Object.assign(out,R.projections(docs,'PES-000001',rev,now,{boards,policies}));for(const b of M.rows(docs,'memberAuthBindings'))out['portalUsers/'+b.uid]=R.bindingView(b,docs['memberIdentities/'+b.idPessoa]);return out}
globalThis.fixture={M,R,C,NOW,config,entity,base,role,event,changes,projectionMap,physical};
