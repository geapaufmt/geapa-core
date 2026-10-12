'use strict';
const fs=require('node:fs'),a=require('node:assert/strict'),{session,save}=require('./session.cjs'),{Host}=require('./sdk-host.cjs'),S=require('./sdk.cjs'),W=require('./wire.cjs'),FT=require('./fixture-tools.cjs');
async function main(){const s=await session(),board=await s.login('BOARD'),host=new Host(s.db,s.verify),doc=p=>s.db.doc(S.ROOT+'/'+p),e=fs.existsSync(__dirname+'/evidence/projections.json')?JSON.parse(fs.readFileSync(__dirname+'/evidence/projections.json')):{samples:[]};
 try{
  if(!e.volumePrepared){const existing=await s.db.collection(S.ROOT+'/canonicalActivities').get(),base=existing.docs.find(d=>d.id==='dev_meeting').data(),extra={};for(let i=existing.size;i<55;i++){const id='B0_LIST_'+String(i).padStart(3,'0');extra['canonicalActivities/'+id]=W.encode({...base,id,title:'B0_NODE_POC_AGENDA_'+i,revision:1}).mapValue.fields;}if(Object.keys(extra).length)await FT.createFixtureDocs(s.db,extra);e.volumePrepared=55;save('projections.json',e);}
  for(const view of ['activity_dev_meeting','activities','frequency'])for(let i=e.samples.filter(v=>v.view===view).length;i<4;i++){
   host.metrics=[];s.db.b0RpcEvents=[];const at=performance.now(),outcome=await host.execute({mode:'projection',token:board.token,viewIds:[view],policyIds:['ACT_POLICY','B0_ATTENDANCE_SELF']}),wallMs=performance.now()-at,rpcs=s.db.b0RpcEvents.map(v=>v.method);
   a.deepEqual(outcome.viewPaths,['portalReaderViews/'+board.uid+'/items/'+view]);const projected=(await doc(outcome.viewPaths[0]).get()).data();a.equal(projected.basis.ownerUid,board.uid);a.equal(projected.basis.generation,(await doc('portalReaderControl/current').get()).data().generation);
   if(view==='activity_dev_meeting')a.equal(projected.data.admin.data.canonicalActivityRevision,(await doc('canonicalActivities/dev_meeting').get()).data().revision);
   if(view==='activities')a.equal(projected.data.bundle.data.calendario.length,55);
   if(view==='frequency'){a.equal(projected.data.epoch,(await doc('activityAttendanceRegistry/current').get()).data().epoch);a.equal(projected.data.result.data.registros.length,4);a(projected.data.result.data.registros.every(v=>v.statusPresenca==='JUSTIFICADA'));}
   e.samples.push({view,index:i,coldLike:i===0,wallMs,...outcome,rpcs,io:[...host.metrics]});save('projections.json',e);console.log({view,sample:i,wallMs:Math.round(wallMs),rpcs:rpcs.length});
  }
 }finally{await s.close([host]);}}
main().catch(e=>{console.error({code:e.code||e.message,stack:e.stack});process.exitCode=1;});
