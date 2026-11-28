'use strict';
const crypto=require('node:crypto');
const {Store,decode}=require('./firestore.cjs'),{createDomain}=require('./domain.cjs');
const {authenticate,appCheckGate}=require('./auth.cjs'),{fence,fail,PROJECT,EXECUTOR,ROOT}=require('./environment.cjs');
const CONTROL='editorialControl/current';

class EditorialStore extends Store{
 guard(){
  const c=this.hostRead(CONTROL);
  if(c?.projectId!==PROJECT||c.schemaVersion!=='editorial-control-v1'||c.executor!==EXECUTOR||!c.actorUids?.includes(this.identity.uid)||c.appCheck?.mode!=='REQUIRED')fail('PILOT_DISABLED');
  if(this.pinned&&JSON.stringify(c)!==this.pinned)fail('ACCESS_REVOKED');
  this.pinned=JSON.stringify(c);this.control=c;return c;
 }
 io(raw){
  if(JSON.parse(raw).kind==='pilotConfig')return JSON.stringify({root:ROOT,shared:true,executor:EXECUTOR,settings:{commands:[],policyIds:[],jobs:[]}});
  return super.io(raw);
 }
 validateWrites(s){
  if(!s.writes.length)return;
  if(this.action!=='editorialExecute'||!/^[a-f0-9]{32}$/.test(this.operationId||''))fail('PILOT_WRITE_DENIED');
  const ratePath='editorialRate/'+crypto.createHash('sha256').update(JSON.stringify(this.identity.uid)).digest('hex').slice(0,32);
  const fixed=new Set([ratePath,'commandExecutorOwners/'+this.operationId,'editorialReceipts/'+this.operationId,'editorialAudit/'+this.operationId]);
  let hasReceipt=false,hasOwner=false,hasAudit=false;
  for(const w of s.writes){
   const isPublicDelete=w.mode==='delete'&&(/^(publicPages\/(home|about)|publicTimeline\/[A-Za-z0-9_-]{1,48})$/.test(w.path));
   if(!isPublicDelete&&w.mode!=='set')fail('PILOT_WRITE_DENIED');
   const isCanonical=/^canonicalEditorial\/(home|about|timeline|timeline_[A-Za-z0-9_-]{1,48})$/.test(w.path);
   const isVersion=/^canonicalEditorialHistory\/version_[a-f0-9]{32}$/.test(w.path);
   if(isVersion&&this.hostRead(w.path))fail('OPERATION_CONFLICT');
   const isHistory=isVersion||/^canonicalEditorialHistory\/(home|about|timeline)$/.test(w.path);
   const isPublic=/^(publicPages\/(home|about)|publicTimeline\/[A-Za-z0-9_-]{1,48})$/.test(w.path);
   if(!isCanonical&&!isHistory&&!isPublic&&!fixed.has(w.path))fail('PILOT_WRITE_DENIED');
   if(isPublicDelete)continue;
   const d=decode(w.data);
   if(isCanonical){
    const prevRev=this.hostRead(w.path)?.revision||0;
    if(d.revision!==prevRev+1)fail('REVISION_CONFLICT');
   }
   if(w.path.endsWith('/'+this.operationId)){
    if(d.ownerUid!==this.identity.uid||d.executor!==EXECUTOR||this.hostRead(w.path))fail('OPERATION_CONFLICT');
    if(w.path.startsWith('editorialReceipts/'))hasReceipt=true;
    if(w.path.startsWith('commandExecutorOwners/'))hasOwner=true;
    if(w.path.startsWith('editorialAudit/'))hasAudit=true;
   }
  }
  if(!hasReceipt||!hasOwner||!hasAudit)fail('PILOT_WRITE_DENIED');
 }
}

async function manageEditorial({db,auth,appCheck,token,appCheckToken,origin,body,hooks={}}){
 const started=performance.now();
 if(!fence().emulator&&Object.keys(hooks).length)fail('DESTINATION_INVALID');
 if(!body||Object.keys(body).sort().join()!=='action,request'||!['editorialPreview','editorialExecute','editorialStatus'].includes(body.action))fail('COMMAND_TYPE_INVALID');
 const identity=await authenticate(auth,token),store=new EditorialStore(db,identity),domain=createDomain(x=>store.io(x));
 store.action=body.action;store.operationId=body.request?.operationId;
 const run=(fn,x)=>JSON.parse(domain.run(fn,x));
 run('(x)=>{nodeIdentity(x);return null}',identity);
 await store.read(()=>store.guard());
 await appCheckGate(appCheck,appCheckToken,store.control,origin);
 store.beforeCommit=async()=>{
  await hooks.beforeCommit?.();
  await authenticate(auth,token);
  await appCheckGate(appCheck,appCheckToken,store.control,origin);
 };
 const result=await store.transaction(()=>run('(x)=>{const M=f4EditorialCommands_(),o={config:nodeConfig,identity:nodeActor,clock:Date.now};if(x.action==="editorialPreview"){f4Core_().exact(x.request,["operationId","input"]);f4Core_().need(f4Core_().op(x.request.operationId),"COMMAND_PAYLOAD_INVALID");const preview=M.preview(nodeTx,o,x.request.input,x.request.operationId),r={operationId:x.request.operationId,input:x.request.input,documentHash:preview.documentHash,epoch:preview.epoch,controlRevision:preview.controlRevision,validUntil:preview.validUntil};return {preview,request:{...r,digest:M.digest(nodeConfig,nodeActor.uid,r)}};}return M[x.action==="editorialExecute"?"execute":"status"](nodeTx,o,x.request);}',body));
 await hooks.afterCommit?.();
 await authenticate(auth,token);
 return {result,metrics:{...store.metrics,totalMs:performance.now()-started}};
}

module.exports={EditorialStore,manageEditorial,CONTROL};
