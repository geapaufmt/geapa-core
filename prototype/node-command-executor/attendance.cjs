'use strict';
require('node:dns').setDefaultResultOrder('ipv4first');
const a=require('node:assert/strict'),crypto=require('node:crypto'),fs=require('node:fs'),{session,save}=require('./session.cjs'),{Host}=require('./sdk-host.cjs'),S=require('./sdk.cjs'),W=require('./wire.cjs');
const id=()=>crypto.randomBytes(16).toString('hex');
async function main(){const s=await session(),board=await s.login('BOARD'),host=new Host(s.db,s.verify),doc=p=>s.db.doc(S.ROOT+'/'+p);let e=fs.existsSync(__dirname+'/evidence/attendance.json')?JSON.parse(fs.readFileSync(__dirname+'/evidence/attendance.json')):{samples:[]};
 try{
  if(!e.fixture){e.fixture=await require('./fixture-tools.cjs').attendance(s.db,board.uid);save('attendance.json',e);}
  if(!e.parity&&e.samples.length&&fs.existsSync(__dirname+'/evidence/attendance-parity-input.json')){e.parity=await require('./parity.cjs').verify(s,host,board.token,JSON.parse(fs.readFileSync(__dirname+'/evidence/attendance-parity-input.json')),'attendance');save('attendance.json',e);}
  for(let i=e.samples.length;i<4;i++){
   const activityId=e.fixture.activities[i];
   host.metrics=[];s.db.b0RpcEvents=[];const roster=await host.execute({mode:'roster',token:board.token,activityId}),rosterIo=[...host.metrics];
   const collectionBefore=(await doc('activityAttendanceCollections/'+activityId).get()).data();
   if(!collectionBefore.rosterCount){const applied=await host.execute({token:board.token,operationId:id(),commandType:'ACT.ATTENDANCE_ROSTER_APPLY',payload:{activityRef:{kind:'activity',id:activityId},rosterRef:roster.rosterRef}});a.equal(applied.result.state,'COMPLETED',JSON.stringify(applied.result));}
   const collection=(await doc('activityAttendanceCollections/'+activityId).get()).data();a.equal(collection.rosterCount,24);const refs=Object.values(collection.roster),before=await s.db.getAll(...refs.map(r=>doc('activityAttendance/'+r.id)));a(before.every(d=>d.data().status==='UNRECORDED'),'NO_ABSENCE_BEFORE_FINALIZATION');
   const records=Object.fromEntries(refs.slice(12).map(r=>[r.id,{statusPresenca:'PRESENTE_PRESENCIAL',observacoes:'Fixture B0'}]));
   host.metrics=[];host.written.clear();host.observedBefore.clear();s.db.b0RpcEvents=[];const clock=Date.now(),started=performance.now(),operationId=id();
   const outcome=await host.execute({token:board.token,operationId,commandType:'ACT.ATTENDANCE_FINALIZE',payload:{activityRef:{kind:'activity',id:activityId},collectionRevision:collection.revision,records}});
   const wallMs=performance.now()-started,rpcs=s.db.b0RpcEvents.map(x=>x.method),io=[...host.metrics];a.equal(outcome.result.state,'COMPLETED',JSON.stringify(outcome.result));
   const rt=performance.now(),back=await s.db.getAll(...refs.map(r=>doc('activityAttendance/'+r.id)),doc('activityAttendanceCollections/'+activityId),doc('commandReceipts/'+operationId),doc('commandProjectionInvalidations/'+operationId));
   const facts=back.slice(0,24).map(d=>d.data());a.equal(facts.filter(d=>d.status==='PRESENT').length,12);a.equal(facts.filter(d=>d.status==='ABSENT').length,12);a(facts.every(d=>d.officialRecordedAt&&d.personRef&&d.membershipRef));a.equal(back[24].data().state,'CLOSED');a(back[25].exists&&back[26].exists);
   e.samples.push({index:i,coldLike:i===0,wallMs,readbackMs:performance.now()-rt,rosterMs:roster.rosterMs,rosterIo,...outcome,rpcs,io,counts:{people:24,present:12,absent:12},attendanceRefs:refs});save('attendance.json',e);if(!e.parity){e.parity=await require('./parity.cjs').parity(s,host,board.token,outcome,clock,'attendance');save('attendance.json',e);}console.log({sample:i,wallMs:Math.round(wallMs),rosterMs:roster.rosterMs,people:24,rpcs:rpcs.length});
  }
 }finally{await s.close([host]);}}
main().catch(e=>{console.error({code:e.code||e.message,stack:e.stack});process.exitCode=1;});
