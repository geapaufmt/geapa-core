'use strict';
const {test}=require('node:test'),a=require('assert/strict'),fs=require('fs'),vm=require('vm');
const root=__dirname+'/../../..';
test('larger internal pages are limited to activity and artifact catalogs',()=>{
 const {queryPlan}=require('../src/query.cjs');let size=0;const q={orderBy(){return this;},limit(n){size=n;return this;}};const db={collection:()=>q};
 const plan=(collection,limit)=>({collection,limit,filters:[],order:[{field:'__name__',direction:'ASCENDING'}]});
 for(const collection of ['canonicalActivities','presentationArtifacts']){queryPlan(db,plan(collection,100),x=>x);a.equal(size,100);a.throws(()=>queryPlan(db,plan(collection,101),x=>x));}
 a.throws(()=>queryPlan(db,plan('memberAggregateHeads',21),x=>x));a.throws(()=>queryPlan(db,plan('privateUnknown',1),x=>x));
});
test('pending labels preserve legacy strings, render DTO status safely and filter DTOs',()=>{
 const ui={escaparHtml:s=>s.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;')};
 const c={window:{PortalGeapaUi:ui},document:{readyState:'loading',addEventListener(){}}};vm.createContext(c);
 const file=fs.readFileSync(root+'/portal-integration/act-readers/assets/act-ui/admin-atividades.js','utf8');
 vm.runInContext(file.replace('})(window);','window.test={renderPending,pendingKind,getActivePendencies};})(window);'),c);
 const x=c.window.test,out=x.renderPending(['MATERIAL',{tipo:'APRESENTACAO',status:'PENDENTE'},{tipo:'<script>',status:'<img>'},null]);
 a(!out.includes('[object'));a(out.includes('APRESENTACAO: PENDENTE'));a(out.includes('&lt;script&gt;'));a(!out.includes('<img>'));
 a.equal(x.pendingKind({tipo:'MATERIAL'}),'MATERIAL');a.equal(x.pendingKind('MATERIAL'),'MATERIAL');a.equal(x.getActivePendencies({statusOperacional:'CANCELADA',pendencias:['MATERIAL']}).length,0);
});
test('cycle creation hint uses separate write capability and retains future-cycle scope',()=>{
 let scope='ALL',available=true;const source={targets:{resourceScope:{cycleAccess:scope}},capabilities:['PRESENTATION_BOARD_READ']};
 const ports={readAuthority:()=>({q:{sourceId:'READ_ONLY',kind:'ROLE'}}),recheckAuthority(){},commandAvailable:(command,key)=>{a.equal(command,'CYCLE_CREATE');a.equal(key,'PRESENTATIONS_PLAN');return available;}};
 const c={f4Core_:()=>({exact(){}}),f4ActivitiesModel_:()=>({need:ok=>a(ok),id:()=>true}),f4Firestore_:()=>({connect:()=>({query:()=>[],get:()=>source})}),f4ActivitiesFunctionalPorts_:()=>ports};vm.createContext(c);
 vm.runInContext(fs.readFileSync(root+'/activities/v2/158_act_planning_compat.gs','utf8'),c);
 const producer=c.f4ActivitiesPlanningCompat_({clock:Date.now,ports,config:{projectId:'synthetic'}});
 a.equal(producer.read({}).data.canCreateCycle,true);available=false;a.equal(producer.read({}).data.canCreateCycle,false);
 available=true;source.targets.resourceScope.cycleAccess='EXPLICIT';a.equal(producer.read({}).data.canCreateCycle,false);
});

