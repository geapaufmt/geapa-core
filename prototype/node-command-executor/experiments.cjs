'use strict';
require('node:dns').setDefaultResultOrder('ipv4first');
const a=require('node:assert/strict'),crypto=require('node:crypto'),fs=require('node:fs'),{session,save}=require('./session.cjs'),{Host}=require('./sdk-host.cjs'),S=require('./sdk.cjs'),W=require('./wire.cjs');
const id=()=>crypto.randomBytes(16).toString('hex');
const semantic=v=>Array.isArray(v)?v.map(semantic):v&&typeof v==='object'?'timestampValue'in v?'TIMESTAMP':Object.fromEntries(Object.entries(v).filter(([k])=>!['updateTime','createTime','proofDigest','ultimaAtualizacao'].includes(k)).sort(([a],[b])=>a.localeCompare(b)).map(([k,x])=>[k,semantic(x)])):v;
async function main(){
 const mode=process.argv[2]||'create',s=await session(),board=await s.login('BOARD'),host=new Host(s.db,s.verify),e=fs.existsSync(__dirname+'/evidence/'+mode+'.json')?JSON.parse(fs.readFileSync(__dirname+'/evidence/'+mode+'.json')):{command:mode,samples:[],parity:[]},doc=p=>s.db.doc(S.ROOT+'/'+p);
 try{
  a.equal(mode,'create');
  for(let i=e.samples.length;i<4;i++){
   const operationId=id(),activityId='act_'+operationId,base=(await doc('canonicalActivities/dev_meeting').get()).data(),payload={activityRef:{kind:'activity',id:activityId},typeRef:base.typeRef,title:'B0_NODE_POC_CREATE_'+i,startsAt:base.startsAt.toMillis(),endsAt:base.endsAt.toMillis(),location:base.location,status:'SCHEDULED',decision:{reasonCode:'B0_CONTROLLED_CREATION',reason:'Criação sintética autorizada para benchmark B0'}};
   host.metrics=[];host.written.clear();host.observedBefore.clear();s.db.b0RpcEvents=[];const clock=Date.now(),started=performance.now();
   const outcome=await host.execute({token:board.token,operationId,commandType:'ACT.ACTIVITY_SAVE',payload});
   const wallMs=performance.now()-started,rpcs=s.db.b0RpcEvents.map(v=>v.method),io=[...host.metrics];a.equal(outcome.result.state,'COMPLETED',JSON.stringify(outcome.result));
   const readStart=performance.now(),back=await s.db.getAll(doc('canonicalActivities/'+activityId),doc('commandReceipts/'+operationId),doc('memberAudit/'+operationId),doc('commandProjectionInvalidations/'+operationId));
   a.equal(back[0].data().revision,1);a.equal(back[0].data().title,payload.title);a(back.every(d=>d.exists));const readbackMs=performance.now()-readStart;
   const sample={index:i,coldLike:i===0,wallMs,readbackMs,...outcome,rpcs,io};e.samples.push(sample);save(mode+'.json',e);
   if(!e.parity.length){
    const before=Object.fromEntries(host.observedBefore),paths=[...host.written],actual=await s.db.getAll(...paths.map(doc));
    save(mode+'-parity-input.json',{before,paths,actual:actual.map(d=>d.exists?W.document(d):null),request:outcome.request,clock});const local=await host.execute({token:board.token,request:outcome.request,memoryDocs:before,clock});a.equal(local.result.state,'COMPLETED',JSON.stringify(local.result));
    for(let j=0;j<paths.length;j++){const path=paths[j],actualFields=actual[j].exists?W.document(actual[j]).fields:null,expectedFields=local.memoryAfter[path]?.fields||null;if(path.startsWith('memberAdmissionTickets/')||path.startsWith('commandProvenance/')||path.startsWith('memberAdmissionUses/'))continue;a.deepEqual(semantic(actualFields),semantic(expectedFields),'PARITY '+path);}
    e.parity.push({pass:true,method:'same immutable v43 logic against in-memory adapter vs SDK transaction',comparedPaths:paths.filter(v=>!/^memberAdmissionTickets|commandProvenance|memberAdmissionUses/.test(v))});save(mode+'.json',e);
   }
   console.log({command:mode,sample:i,wallMs:Math.round(wallMs),reads:io.reduce((s,m)=>s+m.reads,0),rpcs:rpcs.length,writes:io.reduce((s,m)=>s+m.writes,0)});
  }
 }finally{await s.close([host]);}
}
main().catch(e=>{console.error({code:e.code||e.message,stack:e.stack});process.exitCode=1;});
