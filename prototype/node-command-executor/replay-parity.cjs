'use strict';
// Directed adapter check after cleanup: original synthetic readback, no Firestore writes.
const fs=require('node:fs'),a=require('node:assert/strict'),{session,save}=require('./session.cjs'),{Host}=require('./sdk-host.cjs');
async function main(){const s=await session(),actor=await s.login('BOARD'),host=new Host(s.db,s.verify),memoryDocs=JSON.parse(fs.readFileSync(__dirname+'/evidence/cleanup-backup.json')),request=JSON.parse(fs.readFileSync(__dirname+'/evidence/edit-final.json')).samples.at(-1).request;
 try{
  const run=request=>host.execute({token:actor.token,request,memoryDocs}),normal=await run(request);a.equal(normal.result.state,'RECOVERED');
  await a.rejects(()=>run({...request,commandType:'ACT.ACTIVITY_REALIZE'}),/OPERATION_CONFLICT/);
  await a.rejects(()=>run({...request,payload:{...request.payload,title:'B0_CHANGED'}}),/OPERATION_CONFLICT/);
  await a.rejects(()=>run({...request,decisionRef:'DIFFERENT_DECISION'}),/OPERATION_CONFLICT/);
  a.equal(host.touchedAll.size,0);a(host.metrics.every(v=>v.kind==='verify'));
  save('replay-parity.json',{passed:4,failed:0,remoteFirestoreWrites:0,checks:['same immutable request recovers','changed command type denied','changed payload denied','changed decision reference denied'],method:'canonical full intent digest reused from Core v43'});console.log({passed:4,remoteFirestoreWrites:0});
 }finally{await s.close([host]);}}
main().catch(e=>{console.error({code:e.code||e.message,stack:e.stack});process.exitCode=1;});
