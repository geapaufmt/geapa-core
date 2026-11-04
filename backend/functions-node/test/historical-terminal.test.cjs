'use strict';
const {test}=require('node:test'),a=require('node:assert/strict'),H=require('./fixtures/import-reconstruction.cjs'),{fixture,payload,op}=require('./import-recognition.test.cjs'),{harness}=require('./model-frequency.test.cjs');
const ref=H.ref,now=Date.parse('2026-11-03T18:00Z');
function terminal(variant='HISTORICAL_NO_CALL',unknownDate=false){
 const d=fixture(),approval=d['presentationPolicyReferences/SYN_IMPORT_APPROVAL'];delete d['presentationPolicyReferences/SYN_IMPORT_APPROVAL'];
 for(const p of Object.keys(d))if(/^(activityAttendance|activityAttendanceCollections|activityJustifications)\//.test(p))delete d[p];
 const act=d['canonicalActivities/SYN_IMPORT_ACTIVITY'],source=d['migrationEvidence/EA'];act.historicalPresentationRefs=[];
 if(variant==='HISTORICAL_CANCELLED'){act.legacyStatus=source.literal.STATUS_OPERACIONAL='CANCELADA';if(unknownDate)act.civilDate=null;}
 d['migrationEvidence/SYN_TERMINAL_DECISION']={id:'SYN_TERMINAL_DECISION',synthetic:true};
 const dec={id:'SYN_TERMINAL',revision:1,usage:variant,active:true,approved:true,activityId:act.id,sourceHash:H.hash(source),reason:'Synthetic explicit recognition of historical terminal state',decisionEvidenceRef:ref('evidence','SYN_TERMINAL_DECISION'),decisionEvidenceHash:H.hash(d['migrationEvidence/SYN_TERMINAL_DECISION'])};d['presentationPolicyReferences/'+dec.id]=dec;
 const i=H.intentFor(d);Object.assign(i,{operationId:op,variant,factIds:[],justificationIds:[],documentaryAdjustmentRef:ref('policy',dec.id)});Object.assign(approval,{intent:i,intentHash:H.hash(i),factsPreservedHash:H.hash([])});approval.historicalCriteria.sources=[{ref:ref('evidence','EA'),hash:H.hash(source)}];d['presentationPolicyReferences/SYN_IMPORT_APPROVAL']=approval;
 return d;
}
function run(d){return harness(d).run('ACTIVITY_IMPORT_RECOGNIZE',payload,{op,allowed:['ACTIVITY_MANAGE']});}
for(const [variant,unknownDate] of [['HISTORICAL_NO_CALL',false],['HISTORICAL_CANCELLED',false],['HISTORICAL_CANCELLED',true]])test('canonical terminal recognition, accurate projection, immutable history: '+variant+' '+unknownDate,()=>{
 const d=terminal(variant,unknownDate),before=H.plain(d),r=run(d);a.equal(r.response.businessState,'COMMIT_APPLIED',JSON.stringify(r.response));
 const v=d['canonicalActivities/SYN_IMPORT_ACTIVITY'],dto=d['activityAgendaViews/'+v.id].data;
 a.equal(v.recordMode,'HISTORICAL_RECOGNIZED');a.equal(v.historicalResolution.state,'FINAL');a.equal(v.historicalResolution.ordinaryOperationsAllowed,false);
 a.equal(v.status,variant==='HISTORICAL_NO_CALL'?'OCCURRED':'CANCELLED');a.equal(dto.statusOperacional,variant==='HISTORICAL_NO_CALL'?'REALIZADA':'CANCELADA');a.equal(dto.statusChamada,variant==='HISTORICAL_NO_CALL'?'SEM_CHAMADA':'NAO_APLICAVEL');a.equal(dto.reconstrucaoConcluida,true);a.equal(dto.somenteConsulta,true);a.equal(dto.podeRegistrarChamada,false);
 a.deepEqual(v.administrativeReview,before['canonicalActivities/'+v.id].administrativeReview);a.deepEqual(d['canonicalActivities/'+v.id+'/revisions/1'].snapshot,before['canonicalActivities/'+v.id]);a.deepEqual(d['migrationEvidence/EA'],before['migrationEvidence/EA']);
 a(!r.writes.some(w=>/^(activityAttendance|activityAttendanceCollections|activityJustifications)\//.test(w.path)));
 if(unknownDate)a.equal(v.civilDate,null);
 if(variant==='HISTORICAL_NO_CALL'){const real=d['activityRealizations/'+v.realizationRef.id];a.equal(real.occurrenceTemporalObject.precision,'DAY');a.equal(real.occurrenceTemporalObject.civilDate,v.civilDate);a.equal(real.basis.reasonCode,'IMPORTED_NO_CALL_RECOGNIZED');}else a.equal(v.realizationRef,null);
 const h=harness(d);for(const command of ['ACTIVITY_SAVE','ACTIVITY_REALIZE','ACTIVITY_PUBLICATION_SET']){const attempt=h.run(command,{activityRef:ref('activity',v.id),reason:'Synthetic prohibited ordinary mutation',action:'REOPEN'},{allowed:['ACTIVITY_MANAGE']});a.equal(attempt.response.businessState,'REJECTED');a.equal(attempt.writes.length,0);}
 a.throws(()=>JSON.parse(h.domain.run('x=>f4ActivitiesModel_().validate({kind:"activity",id:x.id},{...x,historicalResolution:null})',v)),/HISTORICAL_RESOLUTION_REQUIRED|UNRESOLVED_DEPENDENCY/);
 const repeat=run(d);a.equal(repeat.writes.length,0);a.notEqual(repeat.response.businessState,'COMMIT_APPLIED');
});
for(const [label,mutate] of Object.entries({existingCall:d=>{d['activityAttendanceCollections/SYN_IMPORT_ACTIVITY']={id:'SYN_IMPORT_ACTIVITY',state:'CLOSED'};},wrongSource:d=>{d['migrationEvidence/EA'].literal.STATUS_OPERACIONAL='PLANEJADA';},missingDecision:d=>{d['presentationPolicyReferences/SYN_TERMINAL'].approved=false;},realizationAlreadyPresent:d=>{d['canonicalActivities/SYN_IMPORT_ACTIVITY'].realizationRef=ref('activityRealization','EXISTING');}}))test('terminal recognition refuses '+label,()=>{const d=terminal();mutate(d);const before=H.hash(d),r=run(d);a.equal(r.response.businessState,'REJECTED');a.equal(r.writes.length,0);a.equal(H.hash(d),before);});
