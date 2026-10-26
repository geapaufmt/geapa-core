'use strict';
// Reuse historical scenario DATA, executing only the current versioned Core.
const fs=require('fs'),p=require('path'),Module=require('module');
const dir=p.resolve(__dirname,'../../../prototype/node-command-executor/test-support/reconvergence/activities');
const file=p.join(dir,'harness.cjs'),h=new Module(file,module);h.filename=file;h.paths=Module._nodeModulePaths(dir);
let text=fs.readFileSync(file,'utf8');
const start=text.indexOf('vm.runInThisContext('),end=text.indexOf('\nconst M=',start);
text=text.slice(0,start)+`vm.runInThisContext(fs.readFileSync(${JSON.stringify(p.resolve(__dirname,'../generated/domain.gs'))},'utf8'),{filename:'versioned-core.gs'});`+text.slice(end);
h._compile(text,file);
const sf=p.join(dir,'scenarios.cjs'),s=new Module(sf,module);s.filename=sf;s.paths=h.paths;s.require=name=>name==='./harness.cjs'?h.exports:require(name);s._compile(fs.readFileSync(sf,'utf8'),sf);
const {decode}=require('../src/firestore.cjs');
function sample(command){
 const f=s.exports.scenario(command),delta=Date.now()-h.exports.NOW;
 const convert=v=>typeof v==='string'?v==='p1'?'PES-000001':v==='p2'?'PES-000002':v==='u1'?'CF_ADMIN':v:v&&typeof v==='object'?v.timestampValue?{timestampValue:new Date(Date.parse(v.timestampValue)+delta).toISOString()}:Array.isArray(v)?v.map(convert):Object.fromEntries(Object.entries(v).map(([k,x])=>[k,convert(x)])):v;
 const docs={};for(const [path,value]of f.db.docs)if(!path.startsWith('member')&&!path.startsWith('portalUsers'))docs[path.replace('/p1','/PES-000001').replace('/p2','/PES-000002')]=decode(convert(h.exports.F.encode(value)));
 for(const [path,d]of Object.entries(docs)){if(path.includes('/configRevisions/')){d.parameters.axisRefs=[{kind:'policy',id:'axis'}];d.parameters.quota=false;}if(path==='presentationPolicyReferences/axis'){d.approved=true;d.active=true;}}
 const q=convert(f.q),request={schemaVersion:'geapa-act-command-v13',operationId:require('crypto').randomBytes(16).toString('hex'),commandType:q.commandType,target:{scope:q.scope,expected:q.expected,authorization:{kind:f.kind==='PORTAL_AUTHENTICATED_MEMBER'?'SELF':'ROLE',sourceId:'FLOW_POLICY',functionId:f.kind==='PORTAL_AUTHENTICATED_MEMBER'?null:'VIG-001',actorScope:f.kind==='PORTAL_AUTHENTICATED_MEMBER'?null:'ADMINISTRACAO'}},payload:q.payload,evidenceRefs:[],decisionRef:null};
 if(command==='COMPOSITION_CONFIRM'){const m=docs['presentationCapacityAssessments/manifest'];m.proposedChange={moves:f4ActivitiesCalendarMoves_(m.items.map(v=>({unitRef:v.unitRef,sessionRef:v.destinationSessionRef})))};m.proposedChangeDigest=f4ActivitiesModel_().hash(m.proposedChange);m.sourceCoverageComplete=true;}
 if(command==='COMPATIBILITY_DECIDE'){
 docs['calendarAssignments/assignment2'].sessionRef={kind:'session',id:'session'};
 for(const id of ['theme','theme2']){docs['presentationThematicRequests/'+id].candidate={titleOrObject:'Synthetic '+id,summary:'Synthetic content'};docs['presentationThematicRequests/'+id].workflowState='ANALYSIS';}
 const basis=f4ActivitiesCompatibilityBasis_(docs['presentationSessions/session'],p=>docs[p]);request.payload.basisDigest=f4ActivitiesCompatibilityDigest_(basis);
 }
 if(command==='SESSION_CONFIGURE'){
 const ref={kind:'activity',id:'flow_session_activity'};docs['canonicalActivities/'+ref.id]={...decode({mapValue:{fields:require('./fixtures.cjs').seed().docs['canonicalActivities/'+require('../src/environment.cjs').ACTIVITY].fields}}),id:ref.id,revision:1,status:'SCHEDULED',civilDate:request.payload.civilDate,cycleId:'c'.repeat(32),cycleRef:{kind:'cycle',id:'c'.repeat(32)},title:'Synthetic session activity'};request.payload.activityRef=ref;request.target.scope.targetRefs.push(ref);request.target.expected.push({...ref,revision:1});
 }
 if(command==='OPERATIONAL_CHECK'){
 const pr=docs['canonicalPresentations/presentation'],theme=docs['presentationThematicRequests/theme'];pr.title='Synthetic final title';docs['presentationFlowCases/flow']={...docs['presentationFlowCases/flow'],kind:'FINAL_CONFIRMATION',state:'CONFIRMED',presentationRef:{kind:'presentation',id:pr.id},title:pr.title,thematicRevisionRef:theme.submittedRevisionRef};
 const cfg=Object.values(docs).find(d=>d.parameters);cfg.parameters.operationalChecklist={approved:true,items:[{key:'title',label:'Título'},{key:'format',label:'Formato'}]};
 request.payload.basisRefs=[{kind:'presentation',id:'presentation'},{kind:'session',id:'session'},{kind:'thematic',id:'theme'},{kind:'assignment',id:'assignment'},{kind:'cycle',id:'c'.repeat(32)},{kind:'config',parent:'c'.repeat(32),id:'config'}];for(const r of request.payload.basisRefs)if(!request.target.scope.targetRefs.some(v=>v.kind===r.kind&&v.id===r.id))request.target.scope.targetRefs.push(r);
 }
 return {docs,request,capability:f4ActivitiesPipeline_().capability(request)};
}
module.exports={sample};

function calendarData(){const f=sample('CALENDAR_RELEASE'),{decode}=require('../src/firestore.cjs'),Ts=require('firebase-admin/firestore').Timestamp,b=Object.fromEntries(Object.entries(require('./fixtures.cjs').seed().docs).map(([p,v])=>[p,decode({mapValue:{fields:v.fields}})])),a=b['canonicalActivities/'+require('../src/environment.cjs').ACTIVITY],M=f4ActivitiesModel_(),c='c'.repeat(32);
 for(const sid of ['session','session2']){const session=f.docs['presentationSessions/'+sid],aid='FLOW_'+sid,at=Date.parse(session.civilDate+'T16:00:00-04:00');session.activityRef={kind:'activity',id:aid};f.docs['canonicalActivities/'+aid]={...a,id:aid,cycleId:c,cycleRef:{kind:'cycle',id:c},civilDate:session.civilDate,startsAt:Ts.fromMillis(at),endsAt:Ts.fromMillis(at+3600000),presentationSessionRef:{kind:'session',id:sid}};f.docs['activityAttendanceCollections/'+aid]={...b['activityAttendanceCollections/'+a.id],id:aid,activityRef:session.activityRef,roster:{},rosterCount:0,unrecordedCount:0,cycleId:c};}
 const items=[{personRef:{kind:'person',id:'PES-000001'},personRevision:1,disposition:'INCIDENT',obligationRef:{kind:'obligation',id:'obligation'}},{personRef:{kind:'person',id:'PES-000002'},personRevision:1,disposition:'NOT_INCIDENT',reasonCode:'SYNTHETIC_APPROVED'}];
 f.docs['presentationCapacityHeads/'+c].incidenceCoverageRef={kind:'assessment',id:'coverage'};f.docs['presentationCapacityAssessments/coverage']={id:'coverage',revision:1,schemaVersion:'act-incidence-coverage-v1',state:'SEALED',complete:true,approved:true,cycleRef:{kind:'cycle',id:c},configRef:{kind:'config',parent:c,id:'config'},membersEpoch:1,cycleRegistryEpoch:1,items,digest:M.hash(items),decisionBasis:{reason:'Synthetic approved planning fixture'},validUntil:Ts.fromMillis(Date.now()+3600000)};
 return f.docs;
}
module.exports.calendarData=calendarData;
