'use strict';
const assert=require('node:assert/strict'),L=require('./local.cjs'),O=require('../ops/api.cjs');
function harness(){
 const sample=L.sample('CYCLE_CREATE'),r=L.runtime();
 for(const [path,d] of sample.documents)r.documents.set(path,{...d,name:O.BASE+'/'+O.ROOT+'/'+path});
 r.run(x=>{globalThis.tick=x;Date.now=()=>tick},sample.now);
 const raw=r.run;r.run=(fn,arg)=>raw('f4DevBoundary_().run(()=>('+fn.toString()+')('+JSON.stringify(arg??null)+'))');
 r.run(()=>{globalThis.cfg=f4aTarget_();globalThis.pipe=f4CorePipeline_({config:cfg,clock:()=>tick,verifyIdentity:()=>({uid:'SYN_UID',expiresAt:tick+3600000}),membersAuthorization:()=>[],membersRevisions:()=>[]});globalThis.serial=1000});
 function command(cmd,payload,{self=false,expect='COMPLETED',stale=false,admitOnly=false}={}){
  r.run(()=>{tick+=61000});
  const input=r.run(({cmd,payload,self,stale})=>{
   const M=f4ActivitiesModel_(),F=f4Firestore_(),db=F.connect(cfg),refs=[];
   function collect(v){if(v&&typeof v==='object'){if(v.kind&&v.id){try{const path=M.path(v);if(M.validPath(path))refs.push(v);return}catch{}}Object.values(v).forEach(collect)}}collect(payload);
   const targetRefs=[...new Map(refs.map(v=>[M.path(v),v])).values()];
   const implicit=[M.ref('attendanceRegistry','current')];
   if(payload.activityRef)implicit.push(M.ref('attendanceCollection',payload.activityRef.id));
   if(cmd==='ATTENDANCE_ENROLL')implicit.push(M.ref('attendance','a_'+M.hash({activity:payload.activityRef.id,person:payload.personRef.id}).slice(0,32)));
   for(const ref of targetRefs){const d=db.get(M.path(ref));if(d?.activityRef)implicit.push(d.activityRef,M.ref('attendanceCollection',d.activityRef.id));if(d?.attendanceRef){implicit.push(d.attendanceRef);const fact=db.get(M.path(d.attendanceRef));implicit.push(fact.activityRef,M.ref('attendanceCollection',fact.activityRef.id))}}
   const expected=[...new Map([...targetRefs,...implicit].map(v=>[M.path(v),v])).values()].map(ref=>{const d=db.get(M.path(ref));return {...ref,...(d?{revision:d.revision+(stale?1:0)}:{exists:false})}});
   const capability=f4ActivitiesContracts_()[cmd].capabilities[0];
   // Fixture policy and quota reset isolate domain behavior from the unchanged hourly throttle.
   db.transaction(tx=>{const policy=tx.get('memberAuthzPolicies/ACT_POLICY').data(),counter=tx.get('memberAdmissionCounters/SYN_UID').data();tx.set('memberAuthzPolicies/ACT_POLICY',{...policy,expiresAt:F.Timestamp.fromMillis(tick+3600000),proofKind:self?'SELF':'ROLE',capabilities:[capability],actorSelector:self?{kind:'SELF'}:{kind:'ROLE',roleKey:'DIRETOR_ADMINISTRATIVO',group:'DIRETORIA',actorScope:'ADMINISTRACAO'},targets:{ownPerson:self,allowedTargetScopes:[self?'SELF':'ACTIVITIES'],allowedBoards:[],allowedRoles:[],allowedPersons:self?[]:['PES-000001'],allowedResources:targetRefs.map(M.path)},fieldMasks:{[capability]:[]}});if(counter)tx.set('memberAdmissionCounters/SYN_UID',{...counter,issuedCount:0})});
   return {schemaVersion:'geapa-act-command-v13',operationId:(++serial).toString(16).padStart(32,'0'),commandType:'ACT.'+cmd,target:{scope:{targetRefs},expected,authorization:{kind:self?'SELF':'ROLE',sourceId:'ACT_POLICY',functionId:self?null:'VIG-001',actorScope:self?null:'ADMINISTRACAO'}},payload,evidenceRefs:[],decisionRef:null};
  },{cmd,payload,self,stale});
  const ticket=r.run(input=>pipe.endpoint.post({method:'POST',projectId:cfg.projectId,scriptId:f4DevBoundary_().script,body:JSON.stringify({action:'core.admit.v13',idToken:'LOCAL',request:input})}),input);
  assert(ticket.requestEnvelope,JSON.stringify(ticket));
  if(admitOnly)return {ticket,input};
  const key=r.run(t=>{const F=f4Firestore_(),q={...t.requestEnvelope,ticketId:t.ticketId,receivedAt:F.Timestamp.fromMillis(tick),state:'RECEIVED'};F.connect(cfg).transaction(tx=>{tx.get('commandRequests/'+q.operationId);tx.get('memberAdmissionUses/'+q.ticketId);tx.set('commandRequests/'+q.operationId,q);tx.set('memberAdmissionUses/'+q.ticketId,{schemaVersion:'member-admission-use-v13',ticketId:q.ticketId,operationId:q.operationId,ownerUid:q.ownerUid,reservedAt:q.receivedAt})});return {operationId:q.operationId,commandType:q.commandType,digest:q.digest}},ticket);
  const result=r.run(key=>pipe.endpoint.attempt('LOCAL',key,'DEV_WORKER_A'),key);
  assert.equal(result.state,expect,cmd+': '+JSON.stringify(result));
  const receipt=r.run(key=>f4Firestore_().connect(cfg).get('commandReceipts/'+key.operationId),key);
  if(expect==='COMPLETED')assert.equal(r.run(key=>pipe.endpoint.attempt('LOCAL',key,'DEV_WORKER_B'),key).state,'RECOVERED');
  return {result,receipt,key};
 }
 const get=ref=>r.run(ref=>f4Firestore_().connect(cfg).get(f4ActivitiesModel_().path(ref)),ref);
 return {r,command,get,now:sample.now};
}
const ref=(kind,id)=>({kind,id}),decision={reasonCode:'DEV_CONTROLLED',reason:'Validação funcional com dados sintéticos'},type=ref('activityType','meeting'),activity=ref('activity','dev_meeting'),period=ref('frequencyPeriod','dev_period');
function setup(options={}){const h=harness(),c=h.command,t=h.now;
 c('ACTIVITY_TYPE_SAVE',{typeRef:type,name:'Reunião DEV',countsAttendance:true,durationMinutes:60,visibility:'MEMBERS',allowsJustification:true,status:'ACTIVE',decision});
 c('ACTIVITY_SAVE',{activityRef:activity,typeRef:type,title:'Reunião sintética DEV',startsAt:t+(options.future?3600000:-3600000),endsAt:t+7200000,location:'Ambiente DEV',status:'SCHEDULED',decision});
 const civil=h.r.run(()=>f4Firestore_().connect(cfg).get('canonicalActivities/dev_meeting').civilDate);
 c('FREQUENCY_PERIOD_SAVE',{periodRef:period,label:'Período sintético DEV',civilStart:civil,civilEnd:civil,unit:'ACTIVITY',justifiedMode:'EXCLUDE',excusedMode:'COUNT_PRESENT',decision});
 c('ATTENDANCE_CONFIGURE',{activityRef:activity,periodRef:period,opensAt:t-7200000,closesAt:t+7200000,justificationEndsAt:t+86400000,justificationOpensAt:t-14400000,acceptLateJustification:true,decision});
 const membership=h.r.run(()=>{const s=f4Firestore_().connect(cfg).get('memberLinks/VINC-001');return s?'VINC-001':null});
 assert(membership,'fixture membership missing');
 const payload={activityRef:activity,personRef:ref('person','PES-000001'),membershipRef:ref('membership',membership),decision};
 if(!options.withoutRoster)c('ATTENDANCE_ENROLL',payload);
 const attendance=h.r.run(()=>f4ActivitiesModel_().ref('attendance','a_'+f4ActivitiesModel_().hash({activity:'dev_meeting',person:'PES-000001'}).slice(0,32)));
 return {...h,attendance,civil,enrollment:payload};
}
function main(){const results=[];
 const h=setup(),c=h.command;
 c('ATTENDANCE_ENROLL',h.enrollment,{expect:'FAILED'});results.push('unique attendance');
 c('ATTENDANCE_OPEN',{activityRef:activity});
 c('ATTENDANCE_CLOSE',{activityRef:activity,decision},{expect:'FAILED'});results.push('no implicit absence');
 c('ATTENDANCE_RECORD',{attendanceRef:h.attendance,status:'PRESENT',reason:'Registro pelo diretor responsável'});results.push('director records attendance');
 c('ATTENDANCE_CLOSE',{activityRef:activity,decision});
 c('ATTENDANCE_CORRECT',{attendanceRef:h.attendance,status:'ABSENT',reason:'Correção administrativa sintética'});results.push('correction history');
 c('ATTENDANCE_JUSTIFICATION_REQUEST',{attendanceRef:h.attendance,reason:'Motivo sintético'},{self:true});
 const justification=h.get(h.attendance).pendingJustificationRef;
 c('ATTENDANCE_JUSTIFICATION_REQUEST',{attendanceRef:h.attendance,reason:'Duplicação'},{self:true,expect:'FAILED'});
 c('ATTENDANCE_JUSTIFICATION_DECIDE',{justificationRef:justification,decision:'APPROVE',reason:'Decisão sintética'});assert.equal(h.get(h.attendance).status,'JUSTIFIED');results.push('justification lifecycle');
 c('ACTIVITY_REALIZE',{activityRef:activity,occurred:true,occurrenceTemporalObject:{precision:'DAY',civilDate:h.civil},basis:decision,predecessorRef:{kind:'NONE'}});results.push('explicit realization');
 c('ATTENDANCE_CORRECT',{attendanceRef:h.attendance,status:'PRESENT',reason:'Correção concorrente'},{stale:true,expect:'FAILED'});results.push('stale revision rejected');
 const frequency=h.r.run(()=>f4AttendanceFrequency_([{countable:true,status:'PRESENT',frequencyBasis:{unit:'MINUTE',justifiedMode:'EXCLUDE'},durationMinutes:60},{countable:true,status:'ABSENT',frequencyBasis:{unit:'MINUTE',justifiedMode:'EXCLUDE'},durationMinutes:30},{countable:true,status:'JUSTIFIED',frequencyBasis:{unit:'MINUTE',justifiedMode:'EXCLUDE'},durationMinutes:60}],true));assert.equal(frequency.denominator,90);assert.equal(frequency.percentage,100*60/90);results.push('frequency explicit policy');
 O.save('attendance-local.json',{pass:true,remote:false,results});console.log(JSON.stringify({pass:true,results}));
}
module.exports={harness,setup};if(require.main===module)main();
