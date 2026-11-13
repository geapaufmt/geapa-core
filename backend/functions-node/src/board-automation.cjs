'use strict';
// Internal worker only: no HTTP action accepts a system identity or a write plan.
const {Store,decode}=require('./firestore.cjs'),{createDomain}=require('./domain.cjs');
const {fence,fail,PROJECT,EXECUTOR,ROOT}=require('./environment.cjs');
const CONTROL='institutionalBoardAutomationControl/current',ACTOR='SYSTEM_BOARD_COLLECTIVE_DECISION';
class BoardAutomationStore extends Store{
 guard(){
  fence();const c=this.hostRead(CONTROL);
  if(c?.schemaVersion!=='board-automation-control-v1'||c.projectId!==PROJECT||c.executor!==EXECUTOR||c.enabled!==true||!Number.isSafeInteger(c.startsAt)||!Number.isSafeInteger(c.expiresAt)||Date.now()<c.startsAt||Date.now()>=c.expiresAt)fail('ACCESS_REVOKED');
  if(this.pinned&&this.pinned!==JSON.stringify(c))fail('ACCESS_REVOKED');this.pinned=JSON.stringify(c);this.control=c;return c;
 }
 io(raw){const x=JSON.parse(raw);
  if(x.kind==='pilotConfig')return JSON.stringify({root:ROOT,shared:true,executor:EXECUTOR,settings:{commands:[],policyIds:[],jobs:[]}});
  if(x.kind==='boardAutomaticPlan'){if(!['SETTLE','ENACT','DOCUMENT'].includes(this.phase))fail('PILOT_WRITE_DENIED');this.plan=x.plan;return 'null';}
  if(x.kind==='boardProcedureDocument'){
   if(x.record?.id!==this.recordId||x.record.kind!=='BOARD_PROCEDURE')fail('PILOT_WRITE_DENIED');
   const pdf=require('./board-procedure-pdf.cjs').renderBoardProcedurePdf(x.record),G=require('./workspace-storage.cjs');
   const descriptor={renderer:pdf.renderer,sourceHash:pdf.sourceHash,sha256:pdf.sha256,size:pdf.size,storage:'PRIVATE_STORAGE',objectName:G.PREFIX+pdf.sha256.slice(0,32),bucket:G.BUCKET};
   this.pendingPdf={pdf,descriptor};return JSON.stringify(descriptor);
  }
  return super.io(raw);
 }
 validateWrites(s){
  if(!s.writes.length)return;
  if(!this.plan||!['SETTLE','ENACT','DOCUMENT'].includes(this.phase))fail('PILOT_WRITE_DENIED');
  const expected=new Map(this.plan.map(w=>[w.path,w]));if(expected.size!==this.plan.length)fail('PILOT_WRITE_DENIED');
  const gate=this.phase==='ENACT'?this.hostRead('portalReaderControl/current'):null;
  for(const w of s.writes){
   if(w.path==='portalReaderControl/current'){
    if(this.phase!=='ENACT'||!gate?.enabled||w.mode!=='update'||JSON.stringify(decode(w.data))!==JSON.stringify({...gate,generation:gate.generation+1}))fail('PILOT_WRITE_DENIED');continue;
   }
   const e=expected.get(w.path);if(!e||w.mode!==(e.kind==='delete'?'delete':'set')||e.kind!=='delete'&&JSON.stringify(w.data)!==JSON.stringify(e.data))fail('PILOT_WRITE_DENIED');expected.delete(w.path);
  }
  if(expected.size)fail('PILOT_WRITE_DENIED');
 }
}
const seal='p=>nodeWire({kind:"boardAutomaticPlan",plan:[...p.sourceWrites,...p.headWrites,...p.claimWrites,...p.projectionWrites,...p.technicalWrites,...p.deletes].map(w=>({...w,data:w.kind==="delete"?null:codec.encode(w.data)}))})';
async function deliver({db,operationId,hooks}){
 const store=new BoardAutomationStore(db,{uid:ACTOR}),domain=createDomain(raw=>store.io(raw));store.phase='DOCUMENT';store.specializedReaders=true;
 const run=proof=>JSON.parse(domain.run('(x)=>f4BoardAutomation_().document(nodeTx,{config:nodeConfig,clock:Date.now,sealAutomaticPlan:'+seal+'},x.operationId,x.proof)',{operationId,proof}));
 const job=await store.read(()=>run(null));if(job.state==='NONE'||job.state==='COMPLETED')return;
 const pdf=require('./board-procedure-pdf.cjs').renderBoardProcedurePdf(job.record,job.descriptor.renderer),d=job.descriptor;
 if(pdf.sha256!==d.sha256||pdf.sourceHash!==d.sourceHash||pdf.size!==d.size)fail('RECONCILIATION_REQUIRED');
 const p={effectId:d.sha256.slice(0,32),objectName:d.objectName,descriptor:{name:'procedimento.pdf',mime:'application/pdf',size:d.size,sha256:d.sha256}},g=hooks.pdfProvider||await require('./workspace-storage.cjs').runtimeProvider();
 let found=await g.inspect(p);if(found.outcome!=='CONFIRMED'){
  // Two retries may observe the same missing object. Storage creation is
  // conditional; reconcile a losing upload through the immutable object hash.
  try{const u=await g.initiate(p);await g.chunk(p,u.sessionUri,0,pdf.bytes);}catch(error){found=await g.inspect(p);if(found.outcome!=='CONFIRMED')throw error;}
  found=await g.inspect(p);
 }
 if(found.outcome!=='CONFIRMED'||found.sha256!==d.sha256)fail('RECONCILIATION_REQUIRED');
 await store.transaction(()=>{store.plan=null;return run({sha256:found.sha256,versionId:found.versionId});});
}
async function phase({db,id,phase,hooks={}}){
 const env=fence();if(!env.emulator&&Object.keys(hooks).length)fail('DESTINATION_INVALID');
 if(!/^GOV_[a-f0-9]{32}$/.test(id)||!['SETTLE','ENACT'].includes(phase))fail('COMMAND_PAYLOAD_INVALID');
 const store=new BoardAutomationStore(db,{uid:ACTOR}),domain=createDomain(raw=>store.io(raw));store.phase=phase;store.recordId=id;store.specializedReaders=true;
 store.beforeCommit=async()=>{
  await hooks.beforeCommit?.();
  if(Date.now()>=store.control.expiresAt)fail('ACCESS_REVOKED');
 };
 const result=await store.transaction(()=>{store.plan=null;store.pendingPdf=null;return JSON.parse(domain.run('(x)=>f4BoardAutomation_().execute(nodeTx,{config:nodeConfig,clock:Date.now,sealAutomaticPlan:p=>nodeWire({kind:"boardAutomaticPlan",plan:[...p.sourceWrites,...p.headWrites,...p.claimWrites,...p.projectionWrites,...p.technicalWrites,...p.deletes].map(w=>({...w,data:w.kind==="delete"?null:codec.encode(w.data)}))}),renderProcedureDocument:record=>nodeWire({kind:"boardProcedureDocument",record})},x.id,x.phase)',{id,phase}));});
 await hooks.afterCommit?.();
 if(result.operationId)await deliver({db,operationId:result.operationId,hooks});
 return {result,metrics:store.metrics};
}
async function runRecord({db,id,hooks={}}){
 // Two commits intentionally preserve the definitive decision if its effect
 // cannot be applied. Both phases reuse deterministic original receipt IDs.
 const settled=await phase({db,id,phase:'SETTLE',hooks});
 if(settled.result.state==='WAITING'||['REJECTED','INSUFFICIENT_MANIFESTATIONS'].includes(settled.result.procedureState))return {settled};
 const enacted=await phase({db,id,phase:'ENACT',hooks});return {settled,enacted};
}
async function runDue({db}){
 const store=new BoardAutomationStore(db,{uid:ACTOR}),domain=createDomain(raw=>store.io(raw));
 const ids=await store.read(()=>JSON.parse(domain.run('()=>{const A=f4BoardAutomation_(),get=p=>nodeTx.get(p).data(),o={config:nodeConfig,clock:Date.now},c=A.control(get,o),v=f4InstitutionalGovernance_().state(get);return v.records.filter(r=>r.kind==="BOARD_PROCEDURE"&&r.procedure.type==="EXTRAORDINARY_EXTENSION"&&c.boardIds.includes(r.procedure.boardId)&&r.procedure.exception&&(r.procedure.enactment?.state!=="ENACTED"&&["OPEN","APPROVED"].includes(r.procedure.state)||["SETTLE","ENACT"].some(phase=>get("boardAutomaticDocumentOutbox/"+A.key(r,phase))?.state==="PENDING"))).map(r=>r.id);}',{})));
 if(ids.length>20)fail('SCOPE_TOO_LARGE');const results=[];
 for(const id of ids){try{await runRecord({db,id});results.push({state:'CHECKED'});}catch(e){results.push({state:'BLOCKED',code:/^[A-Z0-9_]{3,80}$/.test(e.code||'')?e.code:'INTERNAL_ERROR'});}}
 return results;
}
module.exports={BoardAutomationStore,phase,runRecord,runDue};
