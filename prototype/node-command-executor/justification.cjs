'use strict';
const fs=require('node:fs'),a=require('node:assert/strict'),crypto=require('node:crypto'),{session,save}=require('./session.cjs'),{Host}=require('./sdk-host.cjs'),S=require('./sdk.cjs'),W=require('./wire.cjs'),FT=require('./fixture-tools.cjs');
const id=()=>crypto.randomBytes(16).toString('hex');
async function main(){const s=await session(),board=await s.login('BOARD'),host=new Host(s.db,s.verify),doc=p=>s.db.doc(S.ROOT+'/'+p),e=fs.existsSync(__dirname+'/evidence/justification.json')?JSON.parse(fs.readFileSync(__dirname+'/evidence/justification.json')):{samples:[]};
 try{
  if(!(await doc('memberAuthzPolicies/B0_ATTENDANCE_SELF').get()).exists){const base=(await doc('memberAuthzPolicies/ACT_POLICY').get()).data(),caps=['ACTIVITY_JUSTIFICATION_SELF','ACTIVITY_ATTENDANCE_READ_SELF','ACTIVITY_AGENDA_READ_SELF'],policy={...base,id:'B0_ATTENDANCE_SELF',proofKind:'SELF',actorSelector:{kind:'SELF'},capabilities:caps,targets:{ownPerson:true,allowedTargetScopes:['SELF'],allowedBoards:[],allowedRoles:[],allowedPersons:[],allowedResources:[],resourceScope:{kind:'OWN_ATTENDANCE'}},fieldMasks:Object.fromEntries(caps.map(c=>[c,[]]))};await FT.createFixtureDocs(s.db,{'memberAuthzPolicies/B0_ATTENDANCE_SELF':W.encode(policy).mapValue.fields});}
  const attendance=JSON.parse(fs.readFileSync(__dirname+'/evidence/attendance.json'));a.equal(attendance.samples.length,4);
  for(let i=e.samples.length;i<4;i++){
   const sample=attendance.samples[i],refs=sample.attendanceRefs,facts=await s.db.getAll(...refs.map(r=>doc('activityAttendance/'+r.id))),own=facts.find(d=>d.data().personRef.id==='PES-000001');a(own&&own.data().status==='ABSENT');let pending=own.data().pendingJustificationRef;
   if(!pending){const request=await host.execute({token:board.token,operationId:id(),commandType:'ACT.ATTENDANCE_JUSTIFICATION_REQUEST',policyIds:['B0_ATTENDANCE_SELF'],payload:{attendanceRef:{kind:'attendance',id:own.id},reason:'Motivo sintético informado pelo próprio participante no benchmark B0'}});a.equal(request.result.state,'COMPLETED',JSON.stringify(request.result));pending=(await own.ref.get()).data().pendingJustificationRef;}
   a(pending);const beforeDecision=(await doc('activityJustifications/'+pending.id).get()).data();a.equal(beforeDecision.state,'PENDING');
   host.metrics=[];host.written.clear();host.observedBefore.clear();s.db.b0RpcEvents=[];const clock=Date.now(),started=performance.now(),operationId=id();
   const outcome=await host.execute({token:board.token,operationId,commandType:'ACT.ATTENDANCE_JUSTIFICATION_DECIDE',payload:{justificationRef:pending,decision:'APPROVE',reason:'Deferimento humano sintético autorizado para o benchmark B0'}});
   const wallMs=performance.now()-started,rpcs=s.db.b0RpcEvents.map(x=>x.method),io=[...host.metrics];a.equal(outcome.result.state,'COMPLETED',JSON.stringify(outcome.result));
   const rt=performance.now(),back=await s.db.getAll(own.ref,doc('activityJustifications/'+pending.id),doc('commandReceipts/'+operationId),doc('memberAudit/'+operationId),doc('commandProjectionInvalidations/'+operationId));a.equal(back[0].data().status,'JUSTIFIED');a.equal(back[1].data().state,'APPROVED');a(back.every(d=>d.exists));a.equal(back[0].data().id,own.id);
   e.samples.push({index:i,coldLike:i===0,wallMs,readbackMs:performance.now()-rt,...outcome,rpcs,io});save('justification.json',e);
   if(!e.parity){e.parity=await require('./parity.cjs').parity(s,host,board.token,outcome,clock,'justification');save('justification.json',e);}console.log({sample:i,wallMs:Math.round(wallMs),rpcs:rpcs.length});
  }
 }finally{await s.close([host]);}}
main().catch(e=>{console.error({code:e.code||e.message,stack:e.stack});process.exitCode=1;});
