'use strict';
const {Store,decode}=require('./firestore.cjs'),{createDomain}=require('./domain.cjs');
const {authenticate,appCheckGate}=require('./auth.cjs'),{fence,fail,PROJECT,EXECUTOR,ROOT}=require('./environment.cjs');
const CONTROL='institutionalGovernanceControl/current';
class GovernanceStore extends Store{
 guard(){
  const c=this.hostRead(CONTROL);
  if(c?.projectId!==PROJECT||c.schemaVersion!=='institutional-governance-control-v1'||c.executor!==EXECUTOR||c.publicationEnabled!==false||!c.actorUids?.includes(this.identity.uid)||c.appCheck?.mode!=='REQUIRED')fail('PILOT_DISABLED');
  if(this.pinned&&JSON.stringify(c)!==this.pinned)fail('ACCESS_REVOKED');
  this.pinned=JSON.stringify(c);this.control=c;return c;
 }
 io(raw){const x=JSON.parse(raw);if(x.kind==='pilotConfig')return JSON.stringify({root:ROOT,shared:true,executor:EXECUTOR,settings:{commands:[],policyIds:[],jobs:[]}});
  if(x.kind==='minuteDocument'){if(this.action!=='institutionalGovernanceExecute'||x.record?.minutes?.state!=='ARCHIVED')fail('PILOT_WRITE_DENIED');const pdf=require('./minute-pdf.cjs').renderMinutePdf(x.record),G=require('./workspace-storage.cjs');const descriptor={renderer:pdf.renderer,sourceHash:pdf.sourceHash,sha256:pdf.sha256,size:pdf.size,storage:x.storage?'PRIVATE_STORAGE':'ON_DEMAND',objectName:x.storage?G.PREFIX+pdf.sha256.slice(0,32):null,bucket:x.storage?G.BUCKET:null};this.pendingPdf=x.storage?{pdf,descriptor}:null;return JSON.stringify(descriptor);}
  return super.io(raw);}
 validateWrites(s){
  if(!s.writes.length)return;
  if(this.action!=='institutionalGovernanceExecute'||s.writes.length!==6)fail('PILOT_WRITE_DENIED');
  const allowed=new Set(['institutionalGovernanceDrafts/current','institutionalGovernanceHistory/current','institutionalGovernanceRate/'+require('crypto').createHash('sha256').update(JSON.stringify(this.identity.uid)).digest('hex').slice(0,32),'commandExecutorOwners/'+this.operationId,'institutionalGovernanceReceipts/'+this.operationId,'institutionalGovernanceAudit/'+this.operationId]);
  for(const w of s.writes){if(w.mode!=='set'||!allowed.has(w.path))fail('PILOT_WRITE_DENIED');const d=decode(w.data);if(w.path==='institutionalGovernanceDrafts/current'&&d.revision!==(this.hostRead(w.path)?.revision||0)+1)fail('REVISION_CONFLICT');if(w.path.includes('/'+this.operationId)&&(d.ownerUid!==this.identity.uid||d.executor!==EXECUTOR||this.hostRead(w.path)))fail('OPERATION_CONFLICT');}
 }
}
async function editGovernance({db,auth,appCheck,token,appCheckToken,origin,body,hooks={}}){
 const started=performance.now();if(!fence().emulator&&Object.keys(hooks).length)fail('DESTINATION_INVALID');
 if(!body||Object.keys(body).sort().join()!=='action,request'||!['institutionalGovernancePreview','institutionalGovernanceExecute','institutionalGovernanceStatus','institutionalGovernanceDocument'].includes(body.action))fail('COMMAND_TYPE_INVALID');
 const identity=await authenticate(auth,token),store=new GovernanceStore(db,identity),domain=createDomain(x=>store.io(x));store.specializedReaders=true;store.action=body.action;store.operationId=body.request?.operationId;
 const run=(fn,x)=>JSON.parse(domain.run(fn,x));run('(x)=>{nodeIdentity(x);return null}',identity);
 await store.read(()=>store.guard());await appCheckGate(appCheck,appCheckToken,store.control,origin);
 const pdfPlan=d=>({effectId:d.sha256.slice(0,32),objectName:d.objectName,descriptor:{name:'ata.pdf',mime:'application/pdf',size:d.size,sha256:d.sha256}});
 const provider=async()=>fence().emulator&&hooks.pdfProvider?hooks.pdfProvider:require('./workspace-storage.cjs').runtimeProvider();
 store.beforeCommit=async()=>{await hooks.beforeCommit?.();await authenticate(auth,token);await appCheckGate(appCheck,appCheckToken,store.control,origin);
  if(store.pendingPdf){const {pdf,descriptor}=store.pendingPdf,p=pdfPlan(descriptor),g=await provider();let found=await g.inspect(p);if(found.outcome!=='CONFIRMED'){const started=await g.initiate(p);await g.chunk(p,started.sessionUri,0,pdf.bytes);found=await g.inspect(p);}if(found.outcome!=='CONFIRMED'||found.sha256!==descriptor.sha256)fail('RECONCILIATION_REQUIRED');await authenticate(auth,token);await appCheckGate(appCheck,appCheckToken,store.control,origin);if(Date.now()>=body.request.validUntil)fail('ACCESS_REVOKED');}
 };
 if(body.action==='institutionalGovernanceDocument'){
  const q=body.request;if(!q||Object.keys(q).sort().join()!=='id,revision'||!/^GOV_[a-f0-9]{32}$/.test(q.id)||!Number.isSafeInteger(q.revision))fail('COMMAND_PAYLOAD_INVALID');
  const read=()=>run('(q)=>{const get=p=>nodeTx.get(p).data(),C=f4Core_();C.need(get("institutionalDirectoryControl/current")?.enabled===true&&Date.now()<nodeActor.expiresAt,"ACCESS_REVOKED");const v=f4InstitutionalGovernance_().read(get,{config:nodeConfig,identity:nodeActor,clock:Date.now}),r=v.records.find(r=>r.id===q.id);C.need(r&&r.revision===q.revision&&r.minutes?.state==="ARCHIVED","AUTHORIZATION_DENIED");return r;}',q);
  const r=await store.read(read),pdf=require('./minute-pdf.cjs').renderMinutePdf(r),d=r.minutes.archive.document;
  if(d&&(pdf.sha256!==d.sha256||pdf.size!==d.size||pdf.sourceHash!==d.sourceHash))fail('RECONCILIATION_REQUIRED');
  let bytes=pdf.bytes;if(d?.storage==='PRIVATE_STORAGE'){const g=await provider(),p=pdfPlan(d),found=await g.inspect(p);if(found.outcome!=='CONFIRMED')fail('RECONCILIATION_REQUIRED');bytes=await g.downloadRange(p,found.versionId,0,d.size);if(require('crypto').createHash('sha256').update(bytes).digest('hex')!==d.sha256)fail('RECONCILIATION_REQUIRED');}
  await authenticate(auth,token);await appCheckGate(appCheck,appCheckToken,store.control,origin);const again=await store.read(read);if(JSON.stringify(again)!==JSON.stringify(r))fail('REVISION_CONFLICT');
  return {result:{name:'Ata-'+r.minutes.number.replace(/[^A-Za-z0-9_-]/g,'-')+'.pdf',mime:'application/pdf',size:bytes.length,sha256:pdf.sha256,sourceHash:pdf.sourceHash,contentRevision:r.minutes.contentRevision,legacy:!r.minutes.fullText,contentBase64:bytes.toString('base64')},metrics:store.metrics};
 }
 const result=await store.transaction(()=>{store.pendingPdf=null;return run('(x)=>{const M=f4InstitutionalGovernance_(),o={config:nodeConfig,identity:nodeActor,clock:Date.now,renderDocument:(record,storage)=>JSON.parse(__io(JSON.stringify({kind:"minuteDocument",record,storage})))};if(x.action==="institutionalGovernancePreview"){f4Core_().exact(x.request,["operationId","input"]);f4Core_().need(f4Core_().op(x.request.operationId),"COMMAND_PAYLOAD_INVALID");const preview=M.preview(nodeTx,o,x.request.input),r={operationId:x.request.operationId,input:x.request.input,documentHash:preview.documentHash,epoch:preview.epoch,controlRevision:preview.controlRevision,validUntil:preview.validUntil};return {preview,request:{...r,digest:M.digest(nodeConfig,nodeActor.uid,r)}};}return M[x.action==="institutionalGovernanceExecute"?"execute":"status"](nodeTx,o,x.request);}',body);});
 await hooks.afterCommit?.();await authenticate(auth,token);
 return {result,metrics:{...store.metrics,totalMs:performance.now()-started}};
}
module.exports={GovernanceStore,editGovernance};
