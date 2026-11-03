'use strict';
const {Store,decode}=require('./firestore.cjs'),{createDomain}=require('./domain.cjs');
const {authenticate,appCheckGate}=require('./auth.cjs'),{fence,fail,PROJECT,EXECUTOR,ROOT}=require('./environment.cjs');
const CONTROL='institutionalEditControl/current';
class InstitutionalStore extends Store{
 guard(){
  const c=this.hostRead(CONTROL);
  if(c?.projectId!==PROJECT||c.schemaVersion!=='institutional-edit-control-v1'||c.executor!==EXECUTOR||c.publicationEnabled!==false||!c.actorUids?.includes(this.identity.uid)||c.appCheck?.mode!=='REQUIRED')fail('PILOT_DISABLED');
  if(this.pinned&&JSON.stringify(c)!==this.pinned)fail('ACCESS_REVOKED');
  this.pinned=JSON.stringify(c);this.control=c;return c;
 }
 io(raw){if(JSON.parse(raw).kind==='pilotConfig')return JSON.stringify({root:ROOT,shared:true,executor:EXECUTOR,settings:{commands:[],policyIds:[],jobs:[]}});return super.io(raw);}
 validateWrites(s){
  if(!s.writes.length)return;
  if(this.action!=='institutionalExecute'||s.writes.length!==5)fail('PILOT_WRITE_DENIED');
  const allowed=new Set(['institutionalDirectory/identity','institutionalContactRate/'+require('crypto').createHash('sha256').update(JSON.stringify(this.identity.uid)).digest('hex').slice(0,32),'commandExecutorOwners/'+this.operationId,'institutionalContactReceipts/'+this.operationId,'institutionalContactAudit/'+this.operationId]);
  for(const w of s.writes){if(w.mode!=='set'||!allowed.has(w.path))fail('PILOT_WRITE_DENIED');const d=decode(w.data);if(w.path.startsWith('institutionalDirectory/')&&d.revision!==this.hostRead(w.path).revision+1)fail('REVISION_CONFLICT');if(w.path.includes('/'+this.operationId)&&(d.ownerUid!==this.identity.uid||d.executor!==EXECUTOR||this.hostRead(w.path)))fail('OPERATION_CONFLICT');}
 }
}
async function editInstitutional({db,auth,appCheck,token,appCheckToken,origin,body,hooks={}}){
 const started=performance.now();if(!fence().emulator&&Object.keys(hooks).length)fail('DESTINATION_INVALID');
 if(!body||Object.keys(body).sort().join()!=='action,request'||!['institutionalPreview','institutionalExecute','institutionalStatus'].includes(body.action))fail('COMMAND_TYPE_INVALID');
 const identity=await authenticate(auth,token),store=new InstitutionalStore(db,identity),domain=createDomain(x=>store.io(x));store.action=body.action;store.operationId=body.request?.operationId;
 const run=(fn,x)=>JSON.parse(domain.run(fn,x));run('(x)=>{nodeIdentity(x);return null}',identity);
 await store.read(()=>store.guard());await appCheckGate(appCheck,appCheckToken,store.control,origin);
 store.beforeCommit=async()=>{await hooks.beforeCommit?.();await authenticate(auth,token);await appCheckGate(appCheck,appCheckToken,store.control,origin);};
 const result=await store.transaction(()=>run('(x)=>{const M=f4InstitutionalContactCommands_(),o={config:nodeConfig,identity:nodeActor,clock:Date.now};if(x.action==="institutionalPreview"){f4Core_().exact(x.request,["operationId","input"]);f4Core_().need(f4Core_().op(x.request.operationId),"COMMAND_PAYLOAD_INVALID");const preview=M.preview(nodeTx,o,x.request.input),r={operationId:x.request.operationId,input:x.request.input,documentHash:preview.documentHash,epoch:preview.epoch,controlRevision:preview.controlRevision,validUntil:preview.validUntil};return {preview,request:{...r,digest:M.digest(nodeConfig,nodeActor.uid,r)}};}return M[x.action==="institutionalExecute"?"execute":"status"](nodeTx,o,x.request);}',body));
 await hooks.afterCommit?.();await authenticate(auth,token);
 return {result,metrics:{...store.metrics,totalMs:performance.now()-started}};
}
module.exports={InstitutionalStore,editInstitutional};
